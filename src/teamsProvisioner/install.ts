/**
 * Team-install step of `teamsProvisioner@1` (epic byte5ai/omadia#860,
 * capability issue byte5ai/omadia-m365-connector#3): install a catalog Teams
 * app into ONE team — the last link of the "1 agent = 1 Entra app + 1 Azure
 * bot + 1 Teams app package" chain.
 *
 * One Graph write: `POST /teams/{teamId}/installedApps` with the
 * `teamsApp@odata.bind` reference to the catalog app. When the caller passes
 * a {@link ConsentedPermissionSet} (resource-specific consent the tenant
 * admin pre-approved), it is included in the request body — and ONLY then:
 * the body shape changes with it, so an absent set must not serialise as
 * `consentedPermissionSet: undefined`/`null`.
 *
 * IDEMPOTENCY — 409 is success. Graph answers 409 when the app is already
 * installed in the team (idempotency key: (teamId, teamsAppId)). The shared
 * http layer maps that to its `{ kind: 'conflict' }` signal and this module
 * maps it onward to `Idempotent<T>` `'already-existed'` — callers branch on
 * `outcome`, nobody string-matches Graph error bodies (see `types.ts`).
 *
 * All HTTP goes through the shared {@link ProvisioningHttp} choke point (one
 * token cache, Retry-After-honouring 429 backoff → `ProvisioningThrottledError`
 * when exhausted, 403 → `ConsentMissingError` carrying
 * {@link TEAM_INSTALL_SCOPE}) — the same one-fallback-branch family as the
 * catalog-upload step. This module opens no second token cache and does no
 * fetch of its own.
 */

import type { ProvisioningHttp } from './http.js';
import type {
  Idempotent,
  InstallToTeamInput,
  TeamAppInstallation,
} from './types.js';

/**
 * Graph APPLICATION permission this step needs. Documented in the
 * scopes/consent unit and granted by the wiring unit's manifest bump —
 * surfaced on 403 via `ConsentMissingError.missingScopes`.
 */
export const TEAM_INSTALL_SCOPE = 'TeamsAppInstallation.ReadWriteForTeam.All';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/** Graph `teamsAppInstallation.permissionType` values. */
export type ResourceSpecificPermissionType = 'application' | 'delegated';

/** One resource-specific (RSC) permission entry of a consented set. */
export interface ResourceSpecificPermission {
  /** RSC permission name, e.g. `ChannelMessage.Read.Group`. */
  readonly permissionValue: string;
  readonly permissionType: ResourceSpecificPermissionType;
}

/**
 * The RSC permissions the installing admin consents to
 * (`consentedPermissionSet` on the Graph install body). Optional — most
 * installs omit it entirely.
 */
export interface ConsentedPermissionSet {
  readonly resourceSpecificPermissions: readonly ResourceSpecificPermission[];
}

/**
 * Install request: the shared `InstallToTeamInput` of `types.ts` plus the
 * optional consented-permission set. The extension lives HERE (not in
 * `types.ts`) because only the install step knows about RSC consent; the
 * `TeamsProvisioner` interface stays satisfied since the extra field is
 * optional.
 */
export interface InstallToTeamRequest extends InstallToTeamInput {
  /** Sent to Graph verbatim when present; omitted from the body otherwise. */
  readonly consentedPermissionSet?: ConsentedPermissionSet;
}

export interface TeamInstallClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  readonly log?: (msg: string) => void;
}

/**
 * Install catalog Teams apps into teams. One step client per provisioner;
 * ordering across chain steps stays middleware-side (agent factory,
 * byte5ai/omadia#863-865).
 */
export class TeamInstallClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly log: (msg: string) => void;

  constructor(opts: TeamInstallClientOptions) {
    this.http = opts.http;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * `POST /teams/{teamId}/installedApps` — install the catalog app
   * (`teamsAppId` = `CatalogTeamsApp.teamsAppId`) into the team.
   *
   * - 2xx → `'created'`, with the Graph installation id when the response
   *   body carried one.
   * - 409 → `'already-existed'` (app already installed) — success, never an
   *   exception.
   * - 403 → `ConsentMissingError([TEAM_INSTALL_SCOPE], 'graph')` from the
   *   http layer, the same typed family as `uploadToCatalog`, so callers get
   *   ONE fallback branch.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   */
  async installToTeam(
    input: InstallToTeamRequest,
  ): Promise<Idempotent<TeamAppInstallation>> {
    const teamId = requireNonEmpty(input.teamId, 'teamId');
    const teamsAppId = requireNonEmpty(input.teamsAppId, 'teamsAppId');

    const jsonBody: Record<string, unknown> = {
      'teamsApp@odata.bind': `${GRAPH_BASE}/appCatalogs/teamsApps/${encodeURIComponent(teamsAppId)}`,
      // Only present when the caller consented — the key must not appear
      // (not even as undefined/null) on the plain install body shape.
      ...(input.consentedPermissionSet !== undefined
        ? { consentedPermissionSet: input.consentedPermissionSet }
        : {}),
    };

    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/teams/${encodeURIComponent(teamId)}/installedApps`,
      step: 'teams.installedApps.add',
      jsonBody,
      missingScopesOn403: [TEAM_INSTALL_SCOPE],
    });

    if (response.kind === 'conflict') {
      this.log(
        `provisioner teams.installedApps.add: app ${teamsAppId} already installed in team ${teamId} (409 → already-existed)`,
      );
      return {
        outcome: 'already-existed',
        value: installation(teamId, teamsAppId, response.json),
      };
    }

    return {
      outcome: 'created',
      value: installation(teamId, teamsAppId, response.json),
    };
  }
}

/** Build the result value; `installationId` only when Graph returned an id. */
function installation(
  teamId: string,
  teamsAppId: string,
  json: unknown,
): TeamAppInstallation {
  const id = installationId(json);
  return {
    teamId,
    teamsAppId,
    ...(id !== undefined ? { installationId: id } : {}),
  };
}

/** `teamsAppInstallation.id` from the response body, when present. */
function installationId(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const id = (json as Record<string, unknown>)['id'];
  return typeof id === 'string' && id.length > 0 ? id : undefined;
}

function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

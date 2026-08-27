/**
 * Team-install step of `teamsProvisioner@1` (epic byte5ai/omadia#860,
 * capability issue byte5ai/omadia-m365-connector#3): install a catalog Teams
 * app into ONE team — the last link of the "1 agent = 1 Entra app + 1 Azure
 * bot + 1 Teams app package" chain — and, since 0.4.0, take it back out
 * again ({@link TeamInstallClient.uninstallFromTeam}, byte5ai/omadia#900).
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
 * The uninstall direction mirrors that: Graph deletes an installation by its
 * INSTALLATION id, not by the catalog app id, so the removal is a
 * lookup-then-DELETE pair — and both halves of the "not installed" case (the
 * lookup finds nothing, or the DELETE races another remover into a 404)
 * collapse into the single idempotent `'already-absent'` outcome. That
 * literal deviates from the `'already-deleted'` of the app-registration /
 * bot rollbacks on purpose: nothing is destroyed here, an app that was never
 * in the team is simply absent (byte5ai/omadia#900 names the signal).
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

/**
 * Input for the uninstall step — the SAME (teamId, teamsAppId) key the
 * install is idempotent on. Callers never have to have kept the Graph
 * installation id around; the step resolves it (see
 * {@link TeamInstallClient.uninstallFromTeam}).
 */
export interface UninstallFromTeamInput {
  readonly teamId: string;
  /** Catalog id (`CatalogTeamsApp.teamsAppId`) — NOT the installation id. */
  readonly teamsAppId: string;
}

/**
 * Idempotency signal of the uninstall direction: `'uninstalled'` when this
 * call removed the installation, `'already-absent'` when the app was not
 * installed in the team (lookup miss, or the DELETE answered 404 because
 * someone else got there first). Both are SUCCESS — callers branch on
 * `outcome` instead of string-matching Graph error bodies.
 */
export type UninstallFromTeamOutcome = 'uninstalled' | 'already-absent';

/**
 * Result of {@link TeamInstallClient.uninstallFromTeam}. `value` mirrors the
 * install result shape ({@link TeamAppInstallation}) so a caller can log the
 * removed installation the same way it logged the created one;
 * `installationId` is present whenever the lookup resolved one — it is absent
 * on a pure lookup miss, where there is nothing to name.
 */
export interface UninstallFromTeamResult {
  readonly outcome: UninstallFromTeamOutcome;
  readonly value: TeamAppInstallation;
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

  /**
   * Remove a catalog app from ONE team — the reverse of
   * {@link installToTeam} (byte5ai/omadia#900).
   *
   * Graph deletes an installation by its INSTALLATION id, which the caller
   * generally does not hold (the install result's `installationId` is
   * optional, and an app installed outside omadia has none on our side), so
   * this is a two-step:
   *
   *   1. `GET /teams/{teamId}/installedApps?$expand=teamsApp&$filter=teamsApp/id eq '…'`
   *      — resolve the installation for the catalog app. No hit →
   *      `'already-absent'`, and NO delete is attempted.
   *   2. `DELETE /teams/{teamId}/installedApps/{installationId}` — 2xx →
   *      `'uninstalled'`; 404 (someone removed it between the two calls) →
   *      `'already-absent'`, via the http layer's `extraOkStatuses`, so a
   *      race never turns into an exception.
   *
   * - 403 → `ConsentMissingError([TEAM_INSTALL_SCOPE], 'graph')` from the
   *   http layer — the same typed family (and the same scope) as the install
   *   direction, so callers keep ONE fallback branch.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   */
  async uninstallFromTeam(
    input: UninstallFromTeamInput,
  ): Promise<UninstallFromTeamResult> {
    const teamId = requireNonEmpty(input.teamId, 'teamId');
    const teamsAppId = requireNonEmpty(input.teamsAppId, 'teamsAppId');

    const installationId = await this.findInstallationId(teamId, teamsAppId);
    if (installationId === undefined) {
      this.log(
        `provisioner teams.installedApps.remove: app ${teamsAppId} not installed in team ${teamId} (lookup miss → already-absent)`,
      );
      return { outcome: 'already-absent', value: { teamId, teamsAppId } };
    }

    const response = await this.http.request({
      resource: 'graph',
      method: 'DELETE',
      url: `${GRAPH_BASE}/teams/${encodeURIComponent(teamId)}/installedApps/${encodeURIComponent(installationId)}`,
      step: 'teams.installedApps.remove',
      missingScopesOn403: [TEAM_INSTALL_SCOPE],
      // Already gone = already removed. Keeps the delete race off the
      // exception path (same treatment as the rollback DELETEs).
      extraOkStatuses: [404],
    });

    const gone = response.kind === 'ok' && response.status === 404;
    if (gone) {
      this.log(
        `provisioner teams.installedApps.remove: installation ${installationId} in team ${teamId} already gone (404 → already-absent)`,
      );
    }

    return {
      outcome: gone ? 'already-absent' : 'uninstalled',
      value: { teamId, teamsAppId, installationId },
    };
  }

  /**
   * `GET /teams/{teamId}/installedApps?$expand=teamsApp&$filter=teamsApp/id eq '…'`
   * — the installation id for a catalog app, or `undefined` when the app is
   * not installed in the team.
   *
   * OData string-literal escaping (quote doubling) + `encodeURIComponent`
   * keep the `$filter` injection-safe for arbitrary ids, exactly as the
   * catalog `externalId` lookup does. The returned entries are re-checked
   * against `teamsApp.id` client-side, so a tenant that ignores the
   * `$filter` cannot make us delete the wrong installation.
   */
  private async findInstallationId(
    teamId: string,
    teamsAppId: string,
  ): Promise<string | undefined> {
    const filter = `teamsApp/id eq '${escapeODataString(teamsAppId)}'`;
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/teams/${encodeURIComponent(teamId)}/installedApps?$expand=teamsApp&$filter=${encodeURIComponent(filter)}`,
      step: 'teams.installedApps.lookup',
      missingScopesOn403: [TEAM_INSTALL_SCOPE],
      // A team that no longer exists has no installation either — the same
      // "nothing to remove" outcome, not an error.
      extraOkStatuses: [404],
    });

    if (response.kind !== 'ok') {
      throw new Error(
        `graph teams.installedApps.lookup unexpectedly answered ${String(response.status)} for team=${teamId}`,
      );
    }
    if (response.status === 404) return undefined;

    for (const entry of listEntries(response.json)) {
      const id = matchingInstallationId(entry, teamsAppId);
      if (id !== undefined) return id;
    }
    return undefined;
  }
}

/** OData string literal escaping: single quotes double up. */
function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

/** `value` array of a Graph collection response. */
function listEntries(json: unknown): unknown[] {
  if (!json || typeof json !== 'object') return [];
  const value = (json as Record<string, unknown>)['value'];
  return Array.isArray(value) ? value : [];
}

/**
 * The `teamsAppInstallation.id` of one lookup entry — but only when its
 * expanded `teamsApp.id` really is the requested catalog app.
 */
function matchingInstallationId(
  json: unknown,
  teamsAppId: string,
): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const entry = json as Record<string, unknown>;
  const teamsApp = entry['teamsApp'];
  if (!teamsApp || typeof teamsApp !== 'object') return undefined;
  if ((teamsApp as Record<string, unknown>)['id'] !== teamsAppId) {
    return undefined;
  }
  const id = entry['id'];
  return typeof id === 'string' && id.length > 0 ? id : undefined;
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

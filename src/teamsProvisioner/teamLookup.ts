/**
 * Team-lookup step of `teamsProvisioner@1` (byte5ai/omadia#860): resolve ONE
 * team id to its display name.
 *
 * WHY IT EXISTS. Everything else in this capability addresses a team by its
 * GUID, and so did every operator screen built on it — the omadia agent detail
 * page listed a provisioned agent's teams as bare
 * `19:…@thread.tacv2` / group-object-id strings, which no operator can read.
 * The middleware cannot resolve them itself: it holds no Graph credentials and
 * reaches Microsoft only through this capability. So the name has to come from
 * here.
 *
 * READ-ONLY, AND NOT AN ENUMERATION. This resolves a name for an id the caller
 * ALREADY has. It deliberately does not list the teams an app is installed in
 * — Graph offers no such query for an app across teams, and inventing one from
 * `getCatalogApp` (which proves tenant-catalog presence, never a team install)
 * would let consumers claim knowledge they do not have.
 *
 * `found: false` IS AN ORDINARY ANSWER. A deleted team, or one the tenant app
 * cannot see, answers 404 — a fact about the tenant, not a failure of the
 * call. It is therefore an `extraOkStatuses` outcome rather than a throw, so
 * callers branch on `found` and never string-match an error. The consumer's
 * contract says a nameless team renders as its id, so a miss degrades to
 * exactly the display everyone had before this step existed.
 *
 * NOT `directory.ts`. That module owns the ENTRA directory-object realities of
 * the app-registration step (replication windows, the recycle bin). This one
 * reads a Teams resource for display purposes and shares nothing with it but
 * the Graph base URL, which it imports rather than restating.
 *
 * All HTTP goes through the shared {@link ProvisioningHttp} choke point (one
 * token cache, Retry-After-honouring 429 backoff, 403 → `ConsentMissingError`
 * carrying {@link TEAM_READ_SCOPE}). This module opens no second token cache
 * and does no fetch of its own.
 */

import { GRAPH_BASE } from './directory.js';
import type { ProvisioningHttp } from './http.js';

/**
 * Graph APPLICATION permission this step needs — the narrowest one that
 * answers the question. `Team.ReadBasic.All` grants exactly id + display name
 * for teams; `Group.Read.All` would also work and reads far more of the
 * directory, so it is not what is asked for. Surfaced on 403 via
 * `ConsentMissingError.missingScopes`.
 */
export const TEAM_READ_SCOPE = 'Team.ReadBasic.All';

/** Input for {@link TeamLookupClient.getTeam}. */
export interface GetTeamInput {
  /** Teams team id — the AAD group object id the install steps use. */
  readonly teamId: string;
}

/** Lookup miss: no such team, or it is not visible to this tenant app. */
export interface TeamNotFound {
  readonly found: false;
}

/** Lookup hit. */
export interface TeamFound {
  readonly found: true;
  readonly teamId: string;
  readonly displayName: string;
}

export type GetTeamResult = TeamNotFound | TeamFound;

export interface TeamLookupClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  readonly log?: (msg: string) => void;
}

/** Read basic directory facts about a team. One client per provisioner. */
export class TeamLookupClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly log: (msg: string) => void;

  constructor(opts: TeamLookupClientOptions) {
    this.http = opts.http;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * `GET /teams/{id}?$select=id,displayName` — the team's display name.
   *
   * `$select` is not an optimisation here, it is the scope boundary made
   * literal: the caller needs a label, so nothing else is requested.
   *
   * A 404 (team gone / not visible) and a 2xx without a usable `displayName`
   * both answer `{ found: false }`. The second case matters: an entry that
   * exists but carries no name cannot be rendered any better than its id, and
   * reporting `found: true` with an empty string would push an empty label
   * into the UI instead.
   */
  async getTeam(input: GetTeamInput): Promise<GetTeamResult> {
    const url = `${GRAPH_BASE}/teams/${encodeURIComponent(input.teamId)}?$select=id,displayName`;
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url,
      step: 'teams.get',
      missingScopesOn403: [TEAM_READ_SCOPE],
      // A team that is not there is an answer, not an error — see module doc.
      extraOkStatuses: [404],
    });
    if (response.kind !== 'ok' || response.status === 404) {
      this.log(
        `[teams-provisioner] teams.get: no team '${input.teamId}' visible to this tenant app`,
      );
      return { found: false };
    }
    const body = response.json;
    const record =
      body !== null && typeof body === 'object'
        ? (body as Record<string, unknown>)
        : null;
    if (record === null) return { found: false };
    const displayName = record['displayName'];
    if (typeof displayName !== 'string' || displayName.trim() === '') {
      return { found: false };
    }
    const id = record['id'];
    return {
      found: true,
      teamId: typeof id === 'string' && id !== '' ? id : input.teamId,
      displayName,
    };
  }
}

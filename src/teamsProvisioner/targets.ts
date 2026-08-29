/**
 * Install-target ENUMERATION for `teamsProvisioner@1` (0.8.0): the lists an
 * operator picks from, instead of typing an id.
 *
 * WHY IT EXISTS. Until now every install path took an id the operator had to
 * produce by hand, and the field test showed what that costs: a pasted
 * identifier that was neither a team nor a channel travelled through several
 * steps and surfaced as an unexplained 404 far from where it was entered.
 * `installTarget.ts` already classifies an id once you HAVE one; this module
 * is the step before that — offering the ids that exist so a wrong one is
 * never entered in the first place.
 *
 * NOT `teamLookup.ts`. That module resolves ONE id the caller already holds
 * and is careful to say it is not an enumeration. This one IS the enumeration,
 * and it deliberately lives next to it rather than inside it, because the two
 * answer different questions under different permissions.
 *
 * ── TEAMS: application permissions, straightforward ────────────────────────
 *
 * `GET /teams` lists every team in the organisation under
 * {@link TEAM_LIST_SCOPE} — the SAME app role `getTeam` already needs and the
 * setup guide already asks for. Paging is `@odata.nextLink` and is followed to
 * the end: a tenant with 30 teams fits in one page and a tenant with 300 does
 * not, and the difference must not be a silently truncated picker.
 *
 * Deliberately NOT `GET /groups?$filter=resourceProvisioningOptions/Any(x:x eq
 * 'Team')`. That query works and returns the same set, but it reads the
 * DIRECTORY, so it needs `Group.Read.All` (or `Directory.Read.All`) — neither
 * of which this connector's documented consent list contains. It would
 * therefore succeed in a tenant whose admin happened to grant broad directory
 * rights and answer 403 in every tenant that followed the setup guide, which
 * is the worst possible failure distribution: it passes the test you run and
 * fails at the customer.
 *
 * ── CHATS: delegated, and there is no application alternative ──────────────
 *
 * Microsoft Graph offers NO tenant-wide application-permission route for
 * listing chats. The `chat: list` reference maps bare `GET /chats` to
 * delegated permissions only; the application form it documents is
 * `GET /users/{id}/chats`, i.e. one user at a time. Enumerating a tenant that
 * way means listing every user and paging each one's chats — O(users) round
 * trips to build a list that would then show an operator thousands of
 * conversations they are not in. That is not a slower version of the right
 * answer, it is the wrong answer.
 *
 * So the picker reads `GET /me/chats` with a DELEGATED token: exactly the
 * chats of the administrator who signed in, which are exactly the ones they
 * can meaningfully drop an agent into. The connector already holds such a
 * credential since 0.6.0 (the catalog-publish sign-in), and 0.8.0 widens that
 * sign-in by one scope, {@link CHAT_READ_DELEGATED_SCOPE}.
 *
 * THE PRICE, stated plainly because a caller has to tell a human about it: a
 * credential stored before 0.8.0 does not carry `Chat.ReadBasic` and cannot
 * acquire it by refreshing. The administrator has to run the device-code
 * sign-in once more. Publishing is unaffected — old credentials keep working
 * for it — and {@link InstallTargetsClient.listChats} reports the gap as a
 * typed `DelegatedScopeRequiredError` before spending a Graph call, so a
 * consumer can render "sign in again to pick a chat" instead of an error.
 *
 * A NOTE ON THE ASTERISK in Graph's permission table
 * (`Chat.ReadBasic.All*`): it marked the retired "protected/metered Teams
 * APIs" regime. The page it pointed at (`/graph/teams-protected-apis`) is now
 * 404 and `/graph/teams-licenses` records that the Teams APIs stopped being
 * metered on 2025-08-25. It is a stale footnote, not an approval gate — and it
 * is moot either way, because the blocker above is the missing tenant-wide
 * route, not a licensing flag.
 *
 * All app-only HTTP goes through the shared {@link ProvisioningHttp} choke
 * point (one token cache, Retry-After backoff, 403 → `ConsentMissingError`);
 * the delegated call rides the same layer with an explicit `bearerToken`.
 */

import {
  CHAT_READ_DELEGATED_SCOPE,
  coversChatList,
  isAccessTokenStale,
  type DelegatedTokenSet,
} from './delegatedAuth.js';
import { GRAPH_BASE } from './directory.js';
import {
  ConsentMissingError,
  DelegatedConsentRequiredError,
  DelegatedScopeRequiredError,
  DelegatedTokenExpiredError,
  ProvisioningRequestError,
} from './errors.js';
import type { ProvisioningHttp } from './http.js';

/**
 * Graph APPLICATION permission the team listing needs — the same role
 * `teamLookup.ts` uses, so enabling the picker consents nothing new.
 */
export const TEAM_LIST_SCOPE = 'Team.ReadBasic.All';

/** Step labels — reused across log lines and typed errors. */
const TEAM_LIST_STEP = 'teams.list';
const CHAT_LIST_STEP = 'chats.list';

/**
 * Page-follow budget. `GET /teams` pages at 100 by default, so 20 pages is
 * ~2000 teams — far past any tenant a human picks from in a dropdown, and a
 * hard stop is what keeps a paging bug from becoming an unbounded loop.
 */
const MAX_LIST_PAGES = 20;

/** One team, reduced to what a picker renders. */
export interface TeamSummary {
  /** Team id — the AAD group object id every install step takes. */
  readonly id: string;
  readonly displayName: string;
}

/** The chat kinds Graph reports. `unknownFutureValue` entries are dropped. */
export type ChatSummaryType = 'group' | 'oneOnOne' | 'meeting';

/** One chat, reduced to what a picker renders. */
export interface ChatSummary {
  /** Chat thread id (`19:…@thread.v2` group, `19:…@unq.gbl.spaces` 1:1). */
  readonly id: string;
  /**
   * Subject of a group chat, or `null`. Explicitly nullable rather than
   * optional: a 1:1 chat HAS no topic, and Graph says so with `null`. A
   * consumer must fall back to {@link memberNames}, and an absent key would
   * let that case be forgotten silently.
   */
  readonly topic: string | null;
  readonly chatType: ChatSummaryType;
  /**
   * Display names of the chat's members, when Graph expanded them. Present
   * for the sake of the untitled chats — without it a 1:1 conversation is an
   * opaque `19:…` string in the picker, which is the problem this whole
   * module exists to remove.
   *
   * Optional because the expansion is best-effort: Graph caps
   * `$expand=members` at 25 entries per chat regardless of paging, and a chat
   * whose members it declines to expand still belongs in the list.
   */
  readonly memberNames?: readonly string[];
}

/** Input for {@link InstallTargetsClient.listChats}. */
export interface ListChatsInput {
  /**
   * SECRET. The stored delegated credential — the same one the catalog
   * publish uses, which since 0.8.0 also carries `Chat.ReadBasic`.
   *
   * Optional in the TYPE so a consumer can call the method with whatever it
   * has and branch on the typed error, rather than having to decide up front
   * whether it is allowed to ask. Omitting it never reaches Graph: it throws
   * `DelegatedScopeRequiredError('no-token')`.
   */
  readonly tokens?: DelegatedTokenSet;
}

export interface InstallTargetsClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  /**
   * Admin-consent URL for the publisher app a token set belongs to. Used only
   * on the delegated 403 path; the provisioner wires the same resolver the
   * catalog client gets.
   */
  readonly adminConsentUrlFor?: (tokens: DelegatedTokenSet) => string;
  readonly log?: (msg: string) => void;
}

/** Enumerate the places an agent can be installed. One client per provisioner. */
export class InstallTargetsClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly adminConsentUrlFor?: (tokens: DelegatedTokenSet) => string;
  private readonly log: (msg: string) => void;

  constructor(opts: InstallTargetsClientOptions) {
    this.http = opts.http;
    if (opts.adminConsentUrlFor !== undefined) {
      this.adminConsentUrlFor = opts.adminConsentUrlFor;
    }
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * `GET /teams?$select=id,displayName` — every team in the organisation,
   * app-only, all pages.
   *
   * `$select` is the scope boundary made literal, exactly as in `getTeam`: a
   * picker needs an id and a label, so nothing else is asked for. (Graph would
   * return a full `team` object with almost every field `null` otherwise —
   * this endpoint only ever populates id, displayName, description and
   * visibility.)
   *
   * Entries without a usable `displayName` are dropped rather than shown as a
   * blank row: an unlabelled option in a picker is worse than one fewer
   * option, and the operator can still paste the id if they genuinely mean it.
   *
   * @throws {ConsentMissingError} 403, carrying {@link TEAM_LIST_SCOPE}.
   */
  async listTeams(): Promise<readonly TeamSummary[]> {
    const teams: TeamSummary[] = [];
    let url = `${GRAPH_BASE}/teams?$select=id,displayName&$top=100`;

    for (let page = 0; page < MAX_LIST_PAGES; page += 1) {
      const response = await this.http.request({
        resource: 'graph',
        method: 'GET',
        url,
        step: TEAM_LIST_STEP,
        missingScopesOn403: [TEAM_LIST_SCOPE],
      });
      if (response.kind !== 'ok') {
        throw new Error(`graph ${TEAM_LIST_STEP} unexpected conflict`);
      }
      const { items, nextLink } = readPage(response.json);
      for (const item of items) {
        const team = parseTeam(item);
        if (team !== undefined) teams.push(team);
      }
      if (nextLink === undefined) return teams;
      url = nextLink;
    }

    this.log(
      `provisioner ${TEAM_LIST_STEP}: stopped after ${String(MAX_LIST_PAGES)} ` +
        `pages with ${String(teams.length)} teams — the list may be truncated`,
    );
    return teams;
  }

  /**
   * `GET /me/chats?$expand=members` — the signed-in administrator's chats,
   * delegated, all pages.
   *
   * See the module doc for why this is delegated and cannot be otherwise. The
   * three pre-flight checks all exist to fail with something an operator can
   * act on, BEFORE a Graph call is spent producing a 401/403 that would need
   * interpreting anyway:
   *
   * - no credential → `DelegatedScopeRequiredError('no-token')`, remedy: sign in;
   * - credential without the chat scope (every credential stored before 0.8.0)
   *   → `DelegatedScopeRequiredError('scope-missing')`, remedy: sign in AGAIN
   *   — refreshing cannot widen a grant;
   * - stale access token → `DelegatedTokenExpiredError('access-token-expired')`,
   *   remedy: `refreshDelegatedToken`, persist the rotated set, retry.
   *
   * The last one is why this method does not refresh on its own: it returns no
   * token set, so a refresh here would rotate the caller's stored refresh
   * token and then discard it, stranding them on a dead credential.
   *
   * `unknownFutureValue` chat types are dropped. Graph adds that member to
   * every evolvable enum, and an entry whose kind this version does not
   * understand cannot be routed to an install method — offering it would put a
   * dead option in the picker.
   *
   * @throws {DelegatedScopeRequiredError} no credential, or one too narrow.
   * @throws {DelegatedTokenExpiredError} stale token, or Graph answered 401.
   * @throws {DelegatedConsentRequiredError} 403 — consent missing/withdrawn.
   */
  async listChats(input?: ListChatsInput): Promise<readonly ChatSummary[]> {
    const tokens = input?.tokens;
    if (tokens === undefined) {
      throw new DelegatedScopeRequiredError(
        CHAT_LIST_STEP,
        [CHAT_READ_DELEGATED_SCOPE],
        'no-token',
      );
    }
    if (!coversChatList(tokens.scopes)) {
      throw new DelegatedScopeRequiredError(
        CHAT_LIST_STEP,
        [CHAT_READ_DELEGATED_SCOPE],
        'scope-missing',
        tokens.scopes,
      );
    }
    if (isAccessTokenStale(tokens)) {
      throw new DelegatedTokenExpiredError('access-token-expired');
    }

    const chats: ChatSummary[] = [];
    // $top is capped at 50 on this endpoint, and $select is NOT among the
    // parameters it supports — only $expand/$top/$filter/$orderby are. So the
    // shape is trimmed here rather than at the API.
    let url = `${GRAPH_BASE}/me/chats?$expand=members&$top=50`;

    for (let page = 0; page < MAX_LIST_PAGES; page += 1) {
      let response;
      try {
        response = await this.http.request({
          resource: 'graph',
          method: 'GET',
          url,
          step: CHAT_LIST_STEP,
          // SECRET — forwarded, never cached, never logged by the http layer.
          bearerToken: tokens.accessToken,
          missingScopesOn403: [CHAT_READ_DELEGATED_SCOPE],
        });
      } catch (err) {
        throw this.translateDelegatedFailure(err, tokens);
      }
      if (response.kind !== 'ok') {
        throw new Error(`graph ${CHAT_LIST_STEP} unexpected conflict`);
      }
      const { items, nextLink } = readPage(response.json);
      for (const item of items) {
        const chat = parseChat(item);
        if (chat !== undefined) chats.push(chat);
      }
      if (nextLink === undefined) return chats;
      url = nextLink;
    }

    this.log(
      `provisioner ${CHAT_LIST_STEP}: stopped after ${String(MAX_LIST_PAGES)} ` +
        `pages with ${String(chats.length)} chats — the list may be truncated`,
    );
    return chats;
  }

  /**
   * Turn an app-only-shaped failure into the delegated vocabulary.
   *
   * A 403 here is NOT a missing app role — no app identity was involved — so
   * reporting `ConsentMissingError` would send an operator to grant an
   * application permission that could never have helped. Same reasoning as the
   * catalog client's delegated publish path.
   */
  private translateDelegatedFailure(
    err: unknown,
    tokens: DelegatedTokenSet,
  ): unknown {
    if (err instanceof ConsentMissingError) {
      return new DelegatedConsentRequiredError(
        CHAT_LIST_STEP,
        [CHAT_READ_DELEGATED_SCOPE],
        this.adminConsentUrlFor?.(tokens) ??
          '(resolve the publisher app first — call startDelegatedSignIn to obtain the admin-consent URL)',
        err,
      );
    }
    if (err instanceof ProvisioningRequestError && err.status === 401) {
      return new DelegatedTokenExpiredError('access-token-expired', err);
    }
    return err;
  }
}

// ---------------------------------------------------------------------------
// Parsing. Every helper answers `undefined` for an entry it cannot render,
// because a malformed row in a picker is a support ticket and a dropped one
// is not.
// ---------------------------------------------------------------------------

/** An OData collection page: its items and the link to the next one. */
function readPage(json: unknown): {
  items: readonly unknown[];
  nextLink: string | undefined;
} {
  if (json === null || typeof json !== 'object') {
    return { items: [], nextLink: undefined };
  }
  const record = json as Record<string, unknown>;
  const value = record['value'];
  const nextLink = record['@odata.nextLink'];
  return {
    items: Array.isArray(value) ? value : [],
    nextLink:
      typeof nextLink === 'string' && nextLink.length > 0 ? nextLink : undefined,
  };
}

function parseTeam(item: unknown): TeamSummary | undefined {
  if (item === null || typeof item !== 'object') return undefined;
  const record = item as Record<string, unknown>;
  const id = record['id'];
  const displayName = record['displayName'];
  if (typeof id !== 'string' || id === '') return undefined;
  if (typeof displayName !== 'string' || displayName.trim() === '') {
    return undefined;
  }
  return { id, displayName };
}

const CHAT_TYPES: ReadonlySet<string> = new Set([
  'group',
  'oneOnOne',
  'meeting',
]);

function parseChat(item: unknown): ChatSummary | undefined {
  if (item === null || typeof item !== 'object') return undefined;
  const record = item as Record<string, unknown>;
  const id = record['id'];
  if (typeof id !== 'string' || id === '') return undefined;

  const chatType = record['chatType'];
  // `unknownFutureValue` (and anything else a newer Graph invents) is dropped:
  // a kind this version cannot route to an install method is a dead option.
  if (typeof chatType !== 'string' || !CHAT_TYPES.has(chatType)) {
    return undefined;
  }

  const rawTopic = record['topic'];
  const topic =
    typeof rawTopic === 'string' && rawTopic.trim() !== '' ? rawTopic : null;

  const memberNames = parseMemberNames(record['members']);
  return {
    id,
    topic,
    chatType: chatType as ChatSummaryType,
    ...(memberNames !== undefined ? { memberNames } : {}),
  };
}

/** `$expand=members` → the display names, or `undefined` when there are none. */
function parseMemberNames(value: unknown): readonly string[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const names: string[] = [];
  for (const member of value) {
    if (member === null || typeof member !== 'object') continue;
    const displayName = (member as Record<string, unknown>)['displayName'];
    if (typeof displayName === 'string' && displayName.trim() !== '') {
      names.push(displayName);
    }
  }
  return names.length > 0 ? names : undefined;
}

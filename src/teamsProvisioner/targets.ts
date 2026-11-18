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
import { redactUnknown } from './redact.js';

/**
 * Graph APPLICATION permission the team listing needs — the same role
 * `teamLookup.ts` uses, so enabling the picker consents nothing new.
 */
export const TEAM_LIST_SCOPE = 'Team.ReadBasic.All';

/** Step labels — reused across log lines and typed errors. */
const TEAM_LIST_STEP = 'teams.list';
const CHAT_LIST_STEP = 'chats.list';

const CHAT_MEMBERS_STEP = 'chats.members.list';

/**
 * Page bounds. A bound is what keeps a paging bug from becoming an unbounded
 * loop; hitting one is REPORTED on the result (`truncated`), never only logged.
 *
 * Teams page at 100, so 20 pages is 2,000 teams. Chats page at 50 at most and
 * Graph often returns fewer: at the 0.9.0 rollout 20 pages held 725 chats and
 * the list stopped there, short and unmarked (byte5ai/omadia#2043, rollout
 * defect D8). 200 pages is up to 10,000 chats.
 */
const MAX_TEAM_PAGES = 20;
const MAX_CHAT_PAGES = 200;
/** Wall-clock bound of one chat listing, paging and member lookups together. */
const CHAT_LIST_TIME_BUDGET_MS = 60_000;
/** Most chats without a topic whose members are looked up in one listing. */
const MAX_MEMBER_LOOKUPS = 100;
const MEMBER_LOOKUP_CONCURRENCY = 5;
/** Which untitled chats get their names first when the lookups run out. */
const MEMBER_LOOKUP_ORDER: readonly ChatSummaryType[] = ['group', 'oneOnOne', 'meeting'];

/**
 * A listing plus whether it is complete. The marks are NON-ENUMERABLE
 * properties on the array, so a caller written against 0.8.0 (iterate,
 * `JSON.stringify`, deep-compare) sees the same array as before.
 */
export type TeamList = readonly TeamSummary[] & {
  /** `true` when a bound stopped the listing before Graph's last page. */
  readonly truncated: boolean;
};

export type ChatList = readonly ChatSummary[] & {
  /** `true` when a bound stopped the listing before Graph's last page. */
  readonly truncated: boolean;
  /**
   * How many chats without a topic came back without `memberNames`: the
   * lookup budget ran out, or Graph refused or returned no names.
   */
  readonly withoutMemberNames: number;
};

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
   * Display names of the chat's members. Present for the sake of the
   * untitled chats — without it a 1:1 conversation is an opaque `19:…` string
   * in the picker, which is the problem this whole module exists to remove.
   *
   * Since 0.9.1 only a chat WITHOUT a topic carries it, and only when its
   * lookup was within budget and answered: the listing reports the rest as
   * `withoutMemberNames`. A chat with a topic comes without it.
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
  /** Clock for the time budget of the chat listing. Tests inject one. */
  readonly now?: () => number;
}

/** Enumerate the places an agent can be installed. One client per provisioner. */
export class InstallTargetsClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly adminConsentUrlFor?: (tokens: DelegatedTokenSet) => string;
  private readonly log: (msg: string) => void;
  private readonly now: () => number;

  constructor(opts: InstallTargetsClientOptions) {
    this.http = opts.http;
    this.now = opts.now ?? Date.now;
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
  async listTeams(): Promise<TeamList> {
    const { items, truncated } = await this.followPages({
      firstUrl: `${GRAPH_BASE}/teams?$select=id,displayName&$top=100`,
      step: TEAM_LIST_STEP,
      maxPages: MAX_TEAM_PAGES,
      fetchPage: async (url) => {
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
        return response.json;
      },
    });
    return markList(items.flatMap((item) => parseTeam(item) ?? []), { truncated });
  }

  /**
   * `GET /me/chats` — the signed-in administrator's chats, delegated, paged to
   * the end.
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
   * MEMBERS (changed in 0.9.1). Until 0.9.0 every page carried
   * `$expand=members`, for every chat. A picker needs member names only where
   * there is no topic to show, so the bulk listing no longer expands, and the
   * chats WITHOUT a topic get their names from `GET /chats/{id}/members`
   * afterwards — at most {@link MAX_MEMBER_LOOKUPS} of them, group chats
   * first. A chat with a topic carries no `memberNames` any more. A lookup
   * that fails costs that chat its names, never the listing.
   *
   * COMPLETENESS. The result says whether it is whole: `truncated` when the
   * page or time bound stopped it, `withoutMemberNames` for the untitled chats
   * left unnamed (see {@link ChatList}).
   *
   * @throws {DelegatedScopeRequiredError} no credential, or one too narrow.
   * @throws {DelegatedTokenExpiredError} stale token, or Graph answered 401.
   * @throws {DelegatedConsentRequiredError} 403 — consent missing/withdrawn.
   */
  async listChats(input?: ListChatsInput): Promise<ChatList> {
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

    const deadline = this.now() + CHAT_LIST_TIME_BUDGET_MS;
    const { items, truncated } = await this.followPages({
      // $top is capped at 50 on this endpoint, and $select is NOT among the
      // parameters it supports — only $expand/$top/$filter/$orderby are. So
      // the shape is trimmed here rather than at the API.
      firstUrl: `${GRAPH_BASE}/me/chats?$top=50`,
      step: CHAT_LIST_STEP,
      maxPages: MAX_CHAT_PAGES,
      deadline,
      fetchPage: async (url) => {
        try {
          const response = await this.http.request({
            resource: 'graph',
            method: 'GET',
            url,
            step: CHAT_LIST_STEP,
            // SECRET — forwarded, never cached, never logged by the http layer.
            bearerToken: tokens.accessToken,
            missingScopesOn403: [CHAT_READ_DELEGATED_SCOPE],
          });
          if (response.kind !== 'ok') {
            throw new Error(`graph ${CHAT_LIST_STEP} unexpected conflict`);
          }
          return response.json;
        } catch (err) {
          throw this.translateDelegatedFailure(err, tokens);
        }
      },
    });

    const listed = items.flatMap((item) => parseChat(item) ?? []);
    const names = await this.lookUpMemberNames(listed, tokens, deadline);
    const chats = listed.map((chat) => {
      const memberNames = names.get(chat.id);
      return memberNames === undefined ? chat : { ...chat, memberNames };
    });
    const withoutMemberNames = chats.filter(
      (chat) => chat.topic === null && chat.memberNames === undefined,
    ).length;
    if (withoutMemberNames > 0) {
      this.log(
        `provisioner ${CHAT_MEMBERS_STEP}: ${String(withoutMemberNames)} of ` +
          `${String(chats.length)} chats have neither a topic nor member names`,
      );
    }
    return markList(chats, { truncated, withoutMemberNames });
  }

  /**
   * Member display names for the chats that have no topic, by chat id.
   * Best-effort and bounded: {@link MAX_MEMBER_LOOKUPS} chats, the listing's
   * deadline, and no further lookups once Graph refuses one for a reason that
   * would repeat (throttling, an expired token, a missing consent). A single
   * chat that answers 404 is skipped.
   */
  private async lookUpMemberNames(
    chats: readonly ChatSummary[],
    tokens: DelegatedTokenSet,
    deadline: number,
  ): Promise<ReadonlyMap<string, readonly string[]>> {
    const untitled = chats.filter(
      (chat) => chat.topic === null && chat.memberNames === undefined,
    );
    const queue = MEMBER_LOOKUP_ORDER.flatMap((type) =>
      untitled.filter((chat) => chat.chatType === type),
    ).slice(0, MAX_MEMBER_LOOKUPS);

    const names = new Map<string, readonly string[]>();
    for (let start = 0; start < queue.length; start += MEMBER_LOOKUP_CONCURRENCY) {
      if (this.now() >= deadline) {
        this.log(
          `provisioner ${CHAT_MEMBERS_STEP}: time budget spent after ${String(start)} lookups`,
        );
        break;
      }
      const batch = queue.slice(start, start + MEMBER_LOOKUP_CONCURRENCY);
      const results = await Promise.all(
        batch.map((chat) => this.memberNamesOf(chat.id, tokens)),
      );
      results.forEach((result, index) => {
        const chat = batch[index];
        if (chat !== undefined && result.names !== undefined) {
          names.set(chat.id, result.names);
        }
      });
      const repeating = results.find((result) => result.stop !== undefined);
      if (repeating?.stop !== undefined) {
        this.log(
          `provisioner ${CHAT_MEMBERS_STEP}: stopped looking up members — ${repeating.stop}`,
        );
        break;
      }
    }
    return names;
  }

  /** `GET /chats/{id}/members` (delegated `Chat.ReadBasic`). Never throws. */
  private async memberNamesOf(
    chatId: string,
    tokens: DelegatedTokenSet,
  ): Promise<{ names?: readonly string[]; stop?: string }> {
    try {
      const response = await this.http.request({
        resource: 'graph',
        method: 'GET',
        url: `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}/members`,
        step: CHAT_MEMBERS_STEP,
        // SECRET — forwarded, never cached, never logged by the http layer.
        bearerToken: tokens.accessToken,
        missingScopesOn403: [CHAT_READ_DELEGATED_SCOPE],
      });
      if (response.kind !== 'ok') return {};
      const names = parseMemberNames(readPage(response.json).items);
      return names === undefined ? {} : { names };
    } catch (err) {
      if (err instanceof ProvisioningRequestError && err.status === 404) return {};
      return { stop: redactUnknown(err) };
    }
  }

  /**
   * Follow `@odata.nextLink` from `firstUrl` to the last page, or to a bound:
   * `maxPages`, the `deadline`, or a next link that points at a page already
   * read. Reports which through `truncated`.
   */
  private async followPages(input: {
    readonly firstUrl: string;
    readonly step: string;
    readonly maxPages: number;
    readonly deadline?: number;
    readonly fetchPage: (url: string) => Promise<unknown>;
  }): Promise<{ items: readonly unknown[]; truncated: boolean }> {
    const items: unknown[] = [];
    const seen = new Set<string>();
    let url: string | undefined = input.firstUrl;
    let stoppedBy: string | undefined;

    while (url !== undefined) {
      if (seen.size >= input.maxPages) stoppedBy = `${String(input.maxPages)} pages`;
      else if (seen.has(url)) stoppedBy = 'a next link that repeats';
      else if (seen.size > 0 && input.deadline !== undefined && this.now() >= input.deadline) {
        stoppedBy = 'its time budget';
      }
      if (stoppedBy !== undefined) break;

      seen.add(url);
      const page = readPage(await input.fetchPage(url));
      items.push(...page.items);
      url = page.nextLink;
    }

    if (stoppedBy !== undefined) {
      this.log(
        `provisioner ${input.step}: stopped by ${stoppedBy} after ${String(seen.size)} ` +
          `pages with ${String(items.length)} entries — the list is truncated`,
      );
    }
    return { items, truncated: stoppedBy !== undefined };
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

/** Attach `marks` to `items` as non-enumerable, read-only properties. */
function markList<T, M extends Record<string, unknown>>(
  items: readonly T[],
  marks: M,
): readonly T[] & Readonly<M> {
  const descriptors = Object.fromEntries(
    Object.entries(marks).map(([key, value]) => [key, { value, enumerable: false }]),
  );
  return Object.defineProperties([...items], descriptors) as unknown as readonly T[] &
    Readonly<M>;
}

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

/** A member collection → the display names, or `undefined` when there are none. */
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

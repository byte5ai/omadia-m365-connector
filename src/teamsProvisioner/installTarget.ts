/**
 * What KIND of Teams scope an operator-supplied identifier names — the
 * pre-flight classification the chat-install step (`install.ts`) needs before
 * it can pick an endpoint.
 *
 * WHY IT EXISTS. Until the chat direction landed, every install target was a
 * team, so a target id needed no interpretation: it went into
 * `POST /teams/{id}/installedApps` and Graph judged it. With two endpoints
 * (`/teams/{id}` and `/chats/{id}`) the identifier itself decides which one is
 * even meaningful, and Teams identifiers wear their kind in their suffix:
 *
 * | shape                     | what it is        | installable        |
 * |---------------------------|-------------------|--------------------|
 * | `xxxxxxxx-xxxx-…` (GUID)  | team (group id)   | yes — via a team   |
 * | `19:…@thread.v2`          | group chat        | yes — via a chat   |
 * | `19:…@unq.gbl.spaces`     | 1:1 chat          | yes — via a chat   |
 * | `19:…@thread.tacv2`       | CHANNEL           | NO                 |
 * | anything else             | unknown           | no                 |
 *
 * THE CHANNEL CASE IS THE POINT. A Teams app is never installed "into a
 * channel" — it is installed into the channel's TEAM. An operator who copies a
 * channel link out of Teams gets a `@thread.tacv2` id, and letting that reach
 * Graph buys a bare 404 that says nothing about the actual mistake. So
 * {@link classifyInstallTarget} names it and {@link ChannelTarget.hint} says
 * what to paste instead.
 *
 * THE AMBIGUITY IS ALSO THE POINT (byte5 field test). A bare 32-hex string
 * such as `abc8af8ec7fc471785d3b83c4d84b667` — no dashes, no `19:` prefix, no
 * `@…` suffix — is a dash-stripped team GUID AND the body of a
 * `19:<32 hex>@thread.v2` chat thread id. Nothing in the string tells the two
 * apart. Guessing would be wrong roughly half the time and silently, so this
 * module answers {@link AmbiguousTarget} — its own outcome, carrying the
 * candidates — and lets the CALLER decide with context the string does not
 * have.
 *
 * PURE AND NETWORK-FREE. This is string shape only: it says what an id LOOKS
 * like, never whether the object exists or is visible to the tenant app. A
 * well-formed id for a deleted chat classifies fine and fails later at Graph
 * (typed `ChatNotFoundError`) — the two questions stay separate on purpose.
 */

/** The scope kinds an install-target identifier can name. */
export type InstallTargetKind =
  | 'team'
  | 'group-chat'
  | 'one-on-one-chat'
  | 'channel'
  | 'ambiguous'
  | 'unknown';

/** A team, addressed by its AAD group object id (dashed GUID). */
export interface TeamTarget {
  readonly kind: 'team';
  readonly teamId: string;
}

/** A chat — group or 1:1 — addressed by its thread id. */
export interface ChatTarget {
  readonly kind: 'group-chat' | 'one-on-one-chat';
  readonly chatId: string;
}

/**
 * A CHANNEL id (`19:…@thread.tacv2`). Well-formed, and not an install target:
 * apps are installed into the channel's team.
 */
export interface ChannelTarget {
  readonly kind: 'channel';
  readonly value: string;
  /** Operator-facing remedy — names the team id as the thing to use. */
  readonly hint: string;
}

/**
 * A bare 32-hex string, readable as BOTH a dash-stripped team GUID and the
 * body of a `19:<32 hex>@thread.v2` chat id. Deliberately not guessed.
 */
export interface AmbiguousTarget {
  readonly kind: 'ambiguous';
  readonly value: string;
  /** What this value could be — the caller picks with outside context. */
  readonly candidates: readonly InstallTargetKind[];
  readonly hint: string;
}

/** No known Teams identifier shape. */
export interface UnknownTarget {
  readonly kind: 'unknown';
  readonly value: string;
  readonly hint: string;
}

export type InstallTarget =
  | TeamTarget
  | ChatTarget
  | ChannelTarget
  | AmbiguousTarget
  | UnknownTarget;

/** Canonical dashed GUID (8-4-4-4-12) — a team is addressed by its group id. */
const DASHED_GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * 32 hex digits, nothing else: a team GUID with the dashes stripped, or the
 * body of a chat thread id with the `19:` / `@thread.v2` wrapper stripped.
 * Both are common ways an id arrives from a copy-paste.
 */
const BARE_32_HEX = /^[0-9a-f]{32}$/i;

/**
 * Suffix table, ordered longest-discriminator-first. `@thread.tacv2` must be
 * tested before `@thread.v2`; the anchored patterns cannot overlap, but the
 * ordering keeps that independent of regex subtleties.
 */
const SUFFIX_PATTERNS: readonly {
  readonly pattern: RegExp;
  readonly kind: 'channel' | 'group-chat' | 'one-on-one-chat';
}[] = [
  { pattern: /^19:.+@thread\.tacv2$/i, kind: 'channel' },
  { pattern: /^19:.+@thread\.v2$/i, kind: 'group-chat' },
  { pattern: /^19:.+@unq\.gbl\.spaces$/i, kind: 'one-on-one-chat' },
];

/**
 * Legacy Skype-interop threads. Teams shows them like chats, but the Graph
 * `/chats` resource does not accept them — a `19:…@thread.skype` id answers
 * 400 there rather than resolving (byte5 field test, omadia teams 0.19.2).
 * Classified `'unknown'` like any other unrecognised shape; the hint is what
 * differs, because "not a chat id" is a much shorter search than "400".
 */
const SKYPE_THREAD = /^19:.+@thread\.skype$/i;

const CHANNEL_HINT =
  'this is a CHANNEL id (`@thread.tacv2`). A Teams app is not installed into ' +
  "a channel — it is installed into the channel's TEAM, and every channel of " +
  'that team then has it. Use the team id (the dashed GUID / group id) as the ' +
  'install target instead.';

const AMBIGUOUS_HINT =
  '32 hex digits with no dashes, no `19:` prefix and no `@…` suffix. That ' +
  'reads BOTH as a team GUID with the dashes stripped AND as the body of a ' +
  '`19:<32 hex>@thread.v2` group-chat id — the string alone cannot decide. ' +
  'Re-enter it in full form: dashed (8-4-4-4-12) for a team, or ' +
  '`19:<value>@thread.v2` for a group chat.';

const SKYPE_HINT =
  'this is a legacy Skype-interop thread (`@thread.skype`). Teams renders it ' +
  'like a chat, but the Graph /chats resource does not accept it — it is not ' +
  'an install target. Use a `19:…@thread.v2` group chat or a team id.';

const UNKNOWN_HINT =
  'not a recognised Teams identifier. Expected a team id (dashed GUID), a ' +
  'group chat (`19:…@thread.v2`) or a 1:1 chat (`19:…@unq.gbl.spaces`).';

/**
 * Classify an operator-supplied install target by its SHAPE.
 *
 * Total and side-effect free: every input gets an answer, including the empty
 * string (`'unknown'`). Surrounding whitespace is trimmed — ids arrive by
 * copy-paste — but nothing else is normalised: the value is handed back the
 * way it will be sent to Graph, so a caller never installs into an id it did
 * not classify.
 *
 * The `'ambiguous'` outcome is a first-class answer, not a failure. See the
 * module doc: a bare 32-hex string genuinely is two things at once, and the
 * caller holds the context needed to choose.
 */
export function classifyInstallTarget(value: string): InstallTarget {
  const trimmed = typeof value === 'string' ? value.trim() : '';

  for (const { pattern, kind } of SUFFIX_PATTERNS) {
    if (!pattern.test(trimmed)) continue;
    if (kind === 'channel') {
      return { kind: 'channel', value: trimmed, hint: CHANNEL_HINT };
    }
    return { kind, chatId: trimmed };
  }

  if (DASHED_GUID.test(trimmed)) {
    return { kind: 'team', teamId: trimmed };
  }

  if (BARE_32_HEX.test(trimmed)) {
    return {
      kind: 'ambiguous',
      value: trimmed,
      candidates: ['team', 'group-chat'],
      hint: AMBIGUOUS_HINT,
    };
  }

  return {
    kind: 'unknown',
    value: trimmed,
    hint: SKYPE_THREAD.test(trimmed) ? SKYPE_HINT : UNKNOWN_HINT,
  };
}

/** Does this classification name something an app can be installed INTO a chat? */
export function isChatTarget(target: InstallTarget): target is ChatTarget {
  return target.kind === 'group-chat' || target.kind === 'one-on-one-chat';
}

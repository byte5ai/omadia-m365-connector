/**
 * Secret redaction for every string this subsystem is about to log, put into
 * an `Error` message, or hand back to a caller (byte5ai/omadia#924).
 *
 * The provisioner already had ONE secret to keep out of its output: the Entra
 * client secret from `addPassword`, handled by never touching the value. The
 * delegated catalog-publish flow adds four more that DO travel through HTTP
 * bodies, which is a different problem — a token endpoint error body, a Graph
 * 4xx body or a stringified `cause` can carry them by accident:
 *
 *   - the delegated ACCESS token   (a JWT)
 *   - the REFRESH token            (a long opaque string, ~90 days of access)
 *   - the DEVICE CODE              (the flow secret; whoever holds it can
 *                                   collect the token the admin just signed for)
 *   - the client secret            (app-only credential, unchanged)
 *
 * The `user_code` is deliberately NOT redacted: it exists to be shown to the
 * operator, and hiding it would break the flow it belongs to.
 *
 * This is a LAST LINE OF DEFENCE, not the strategy. The strategy is that no
 * call site ever passes a token into a message in the first place; this
 * function is what keeps a body we did not write from leaking one anyway. It
 * runs on the error path only, where an extra regex pass costs nothing.
 */

/** Placeholder substituted for anything that looks like a credential. */
export const REDACTED = '[redacted]';

/**
 * Named credential fields, in both wire encodings the identity platform uses:
 * `application/x-www-form-urlencoded` (our own token requests) and JSON (its
 * responses). Matched case-insensitively.
 */
const SECRET_FIELD_NAMES = [
  'access_token',
  'refresh_token',
  'id_token',
  'device_code',
  'client_secret',
  'client_assertion',
  'assertion',
  'secretText',
] as const;

/**
 * Deliberately NOT in that list: a bare `code` field.
 *
 * The OAuth authorization code would qualify on paper, but this subsystem
 * never handles one — the device authorization grant uses `device_code`, which
 * IS listed. Meanwhile `error.code` is how every Graph and ARM error envelope
 * names its machine-readable reason, and redacting that would blind the very
 * failure classification this codebase depends on: the taken-bot-handle verdict
 * (byte5ai/omadia#921) and the taken-uniqueName conflict rules both branch on
 * reading `code` back out of an error message.
 */

const FIELD_ALTERNATION = SECRET_FIELD_NAMES.join('|');

/** `"access_token": "…"` / `'refresh_token':'…'` inside a JSON body. */
const JSON_FIELD_PATTERN = new RegExp(
  `("|')(${FIELD_ALTERNATION})\\1\\s*:\\s*("|')(?:\\\\.|(?!\\3).)*\\3`,
  'gi',
);

/** `refresh_token=…` inside a form-encoded body or a query string. */
const FORM_FIELD_PATTERN = new RegExp(`\\b(${FIELD_ALTERNATION})=([^&\\s"']+)`, 'gi');

/** `Authorization: Bearer <token>` however it got into the text. */
const BEARER_PATTERN = /\bBearer\s+[A-Za-z0-9._~+/=-]{8,}/gi;

/**
 * A bare JWT (`eyJ…header.payload.signature`) — every Entra access/id token
 * starts `eyJ` because that is base64url for `{"`.
 */
const JWT_PATTERN = /\beyJ[A-Za-z0-9_-]{6,}\.[A-Za-z0-9_-]{6,}\.[A-Za-z0-9_-]*/g;

/**
 * A bare long opaque credential — refresh tokens and device codes carry no
 * structure to key on, so length is the only signal left.
 *
 * 60 is chosen to sit ABOVE everything this subsystem legitimately prints and
 * BELOW the shortest credential: Graph object ids and `uniqueName`s are ≤ 64
 * characters but contain separators that break the run, a tenant/app GUID is
 * 36, an ARM resource id contains slashes, and a `user_code` is ~9. Real
 * refresh tokens and device codes are hundreds of characters of unbroken
 * base64url.
 */
const LONG_OPAQUE_PATTERN = /\b[A-Za-z0-9_-]{60,}\b/g;

/**
 * Strip anything credential-shaped from `text`.
 *
 * Deliberately over-redacts: turning an unrecognised long blob into
 * `[redacted]` costs a little diagnostic detail, while missing one writes a
 * live credential into a log file that outlives the token.
 */
export function redactSecrets(text: string): string {
  return text
    .replace(JSON_FIELD_PATTERN, (_match, quote: string, field: string) =>
      `${quote}${field}${quote}:"${REDACTED}"`,
    )
    .replace(FORM_FIELD_PATTERN, (_match, field: string) => `${field}=${REDACTED}`)
    .replace(BEARER_PATTERN, `Bearer ${REDACTED}`)
    .replace(JWT_PATTERN, REDACTED)
    .replace(LONG_OPAQUE_PATTERN, REDACTED);
}

/**
 * `redactSecrets` for a value of unknown type — the shape an `unknown` caught
 * error arrives in. Never throws: a value whose `toString` blows up still has
 * to produce a log line.
 */
export function redactUnknown(value: unknown): string {
  let text: string;
  try {
    text = value instanceof Error ? `${value.name}: ${value.message}` : String(value);
  } catch {
    return '[unprintable]';
  }
  return redactSecrets(text);
}

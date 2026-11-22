import { URLSearchParams } from 'node:url';

/**
 * Minimal Microsoft Graph client wired to the bot's app credentials
 * (MICROSOFT_APP_ID + MICROSOFT_APP_PASSWORD + MICROSOFT_APP_TENANT_ID).
 * Used to download files that arrive in Teams channel messages as
 * SharePoint / OneDrive links rather than pre-signed download URLs —
 * the bot has `Files.Read.All` so a `/shares/u!<b64>/driveItem/content`
 * call resolves them directly.
 *
 * Access tokens are cached in-process until 60 s before expiry.
 */

export interface GraphClientOptions {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  log?: (msg: string) => void;
  fetchImpl?: typeof fetch;
}

/** Default scope for app-only Graph tokens (client-credentials flow). */
export const GRAPH_TOKEN_SCOPE = 'https://graph.microsoft.com/.default';

/** An AAD client-credentials identity (client-secret flow). */
export interface AadClientCredential {
  readonly tenantId: string;
  readonly clientId: string;
  readonly clientSecret: string;
}

/** One token request against the v2 client-credentials endpoint. */
export interface AadTokenRequest extends AadClientCredential {
  /** Resource audience, e.g. {@link GRAPH_TOKEN_SCOPE} or `https://management.azure.com/.default`. */
  readonly scope: string;
  /**
   * Budget for the token POST. Defaults to 2000 ms — the Teams hot-path
   * precedent (see `getUserMail`); background flows (provisioning) pass a
   * longer budget.
   */
  readonly timeoutMs?: number;
}

const TOKEN_EXPIRY_SLACK_MS = 60_000;
const DEFAULT_TOKEN_TIMEOUT_MS = 2000;

/**
 * In-process client-credentials token cache, keyed per
 * (tenantId, clientId, scope) so one instance can serve multiple audiences
 * (Graph + ARM) and multiple identities (connector app vs. dedicated ARM
 * service principal). Single-flight per key; entries refresh 60 s before
 * expiry; a failed fetch clears its entry so the next call retries.
 *
 * Generalisation of the former `GraphClient`-private single-token cache —
 * `GraphClient` now delegates here with the Graph scope.
 */
export class AadTokenCache {
  private readonly fetchImpl: typeof fetch;
  private readonly cache = new Map<
    string,
    Promise<{ token: string; expiresAt: number }>
  >();

  constructor(opts?: { fetchImpl?: typeof fetch }) {
    this.fetchImpl = opts?.fetchImpl ?? fetch;
  }

  async accessToken(req: AadTokenRequest): Promise<string> {
    const key = `${req.tenantId}\u0000${req.clientId}\u0000${req.scope}`;
    const now = Date.now();
    const pending = this.cache.get(key);
    if (pending) {
      const cached = await pending.catch(() => undefined);
      if (cached && cached.expiresAt - TOKEN_EXPIRY_SLACK_MS > now) {
        return cached.token;
      }
    }
    const fresh = this.fetchToken(req);
    this.cache.set(key, fresh);
    try {
      return (await fresh).token;
    } catch (err) {
      this.cache.delete(key);
      throw err;
    }
  }

  private async fetchToken(req: AadTokenRequest): Promise<{
    token: string;
    expiresAt: number;
  }> {
    const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(req.tenantId)}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: req.clientId,
      client_secret: req.clientSecret,
      scope: req.scope,
    });
    const response = await this.fetchImpl(tokenUrl, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
      // Bounded so a hung Entra token endpoint can't stall the caller (see getUserMail).
      signal: AbortSignal.timeout(req.timeoutMs ?? DEFAULT_TOKEN_TIMEOUT_MS),
    });
    if (!response.ok) {
      const body = await safeBody(response);
      // NEVER include client_secret (or the params string) here.
      throw new Error(
        `aad token ${String(response.status)} scope=${req.scope} body=${truncate(body, 200)}`,
      );
    }
    const json = (await response.json()) as {
      access_token?: unknown;
      expires_in?: unknown;
    };
    if (typeof json.access_token !== 'string' || json.access_token.length === 0) {
      throw new Error(`aad token: no access_token in response scope=${req.scope}`);
    }
    const expiresInSec =
      typeof json.expires_in === 'number' && Number.isFinite(json.expires_in)
        ? json.expires_in
        : 3600;
    return {
      token: json.access_token,
      expiresAt: Date.now() + expiresInSec * 1000,
    };
  }
}

/** Code of a {@link GraphRequestError} for HTTP 403: the app token may not read this resource. */
export const GRAPH_NOT_PERMITTED = 'not_permitted';

/**
 * A non-2xx Graph answer. The message keeps the `graph <op> <status> …` format;
 * `status` and `code` let callers tell "not permitted" (403, `code:
 * 'not_permitted'`) from other failures without parsing the text. Example: a
 * facilitator with its own Teams bot app may lack RSC `ChannelMessage.Read.Group`
 * in a team the deployment app is not installed in.
 */
export class GraphRequestError extends Error {
  readonly status: number;
  readonly code: string | undefined;

  constructor(status: number, message: string) {
    super(message);
    this.name = 'GraphRequestError';
    this.status = status;
    this.code = status === 403 ? GRAPH_NOT_PERMITTED : undefined;
  }
}

export class GraphClient {
  private readonly tenantId: string;
  private readonly clientId: string;
  private readonly clientSecret: string;
  private readonly log: (msg: string) => void;
  private readonly fetchImpl: typeof fetch;

  private readonly tokens: AadTokenCache;

  constructor(opts: GraphClientOptions) {
    this.tenantId = opts.tenantId;
    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
    this.fetchImpl = opts.fetchImpl ?? fetch;
    this.tokens = new AadTokenCache({ fetchImpl: this.fetchImpl });
  }

  /**
   * Fetch a Teams group-chat message by chat + message id. Used as a
   * fallback when Teams' inbound Activity does NOT carry file attachments
   * directly — a known pattern for group-chat bots, even with RSC
   * `ChatMessage.Read.Chat` granted. Graph's representation of the same
   * message DOES include the attachments with SharePoint-linked
   * `contentUrl`s, which we can then feed into `downloadBySharingUrl`.
   *
   * `chatId` is the `conversation.id` (`19:…@thread.skype`) on inbound;
   * `messageId` is `activity.id`. Both are URL-escaped by the caller-safe
   * Graph SDK pattern we emulate with `encodeURIComponent`.
   */
  async fetchChatMessage(chatId: string, messageId: string): Promise<{
    id: string;
    attachments: GraphChatMessageAttachment[];
  }> {
    const token = await this.accessToken();
    const url = `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`;
    const response = await this.fetchImpl(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) {
      const body = await safeBody(response);
      throw new Error(
        `graph chatMessage ${String(response.status)} chat=${chatId.slice(0, 60)} msg=${messageId.slice(0, 60)} body=${truncate(body, 200)}`,
      );
    }
    return parseGraphMessage(await response.json(), messageId);
  }

  /**
   * Fetch a Teams CHANNEL message (a thread root or one of its replies). The
   * channel counterpart of {@link fetchChatMessage}: a channel conversation
   * (`19:…@thread.tacv2;messageid=<root>`) is not a chat, and
   * `/chats/{id}/messages/{id}` answers 404 "Invalid ThreadId" for it.
   * Needs `ChannelMessage.Read.Group` (RSC) or `ChannelMessage.Read.All`.
   *
   * `teamId` is the team's AAD group id (`channelData.team.aadGroupId`),
   * `channelId` the channel's thread id, `rootMessageId` the thread root and
   * `messageId` the message itself (equal to the root for the root post).
   */
  async fetchChannelMessage(input: {
    teamId: string;
    channelId: string;
    rootMessageId: string;
    messageId: string;
  }): Promise<{ id: string; attachments: GraphChatMessageAttachment[] }> {
    const token = await this.accessToken();
    const base = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(input.teamId)}/channels/${encodeURIComponent(input.channelId)}/messages/${encodeURIComponent(input.rootMessageId)}`;
    const url = input.messageId === input.rootMessageId ? base : `${base}/replies/${encodeURIComponent(input.messageId)}`;
    const response = await this.fetchImpl(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) {
      const body = await safeBody(response);
      throw new GraphRequestError(
        response.status,
        `graph channelMessage ${String(response.status)} channel=${input.channelId.slice(0, 60)} msg=${input.messageId.slice(0, 60)} body=${truncate(body, 200)}`,
      );
    }
    return parseGraphMessage(await response.json(), input.messageId);
  }

  /**
   * Fetch the raw content of a file identified by a SharePoint / OneDrive
   * sharing URL. Returns the bytes + content-type + filename-hint (from
   * Content-Disposition when present). Throws on any non-2xx.
   */
  async downloadBySharingUrl(url: string): Promise<{
    bytes: Buffer;
    contentType: string;
    fileName?: string;
  }> {
    const token = await this.accessToken();
    const encoded = encodeSharingUrl(url);
    const graphUrl = `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem/content`;
    const response = await this.fetchImpl(graphUrl, {
      headers: { Authorization: `Bearer ${token}` },
      // Graph returns 302 → redirect to the actual storage host; Node follows
      // automatically unless we tell it otherwise.
      redirect: 'follow',
    });
    if (!response.ok) {
      const body = await safeBody(response);
      throw new Error(
        `graph download ${String(response.status)} url=${truncate(url, 80)} body=${truncate(body, 200)}`,
      );
    }
    const contentType =
      response.headers.get('content-type') ?? 'application/octet-stream';
    const fileName = fileNameFromDisposition(
      response.headers.get('content-disposition'),
    );
    const arr = await response.arrayBuffer();
    return {
      bytes: Buffer.from(arr),
      contentType,
      ...(fileName ? { fileName } : {}),
    };
  }

  /**
   * Post a Teams Activity Notification (Bell-icon push) into a specific
   * scope — a Team, a Chat, or a User. Requires the matching
   * `TeamsActivity.Send.*` permission to be consented on the bot's App
   * Registration and the activity-type to be declared in the Teams
   * App-Manifest's `activities.activityTypes[]`.
   *
   * Docs:
   *   /teams/{id}  → TeamsActivity.Send.Group
   *   /chats/{id}  → TeamsActivity.Send.Chat
   *   /users/{id}/teamwork → TeamsActivity.Send (User)
   */
  async sendActivityNotification(opts: {
    /**
     * Graph scope prefix — `/teams/{id}`, `/chats/{id}`, or
     * `/users/{id}/teamwork`. The caller picks based on which
     * permission is consented.
     */
    scope: string;
    /** Must match an entry in the App-Manifest's activityTypes. */
    activityType: string;
    /** Short text shown under the activity-feed entry. */
    previewText: string;
    /** Template-parameter map filled into the activityType's
     *  templateText placeholders. */
    templateParameters: ReadonlyArray<{ name: string; value: string }>;
    /**
     * Notification topic — either an entity URL (Graph object the user
     * lands on) or a webUrl (free-form deep link). Teams renders this
     * as the activity-feed item's destination.
     */
    topic: { source: 'entityUrl'; value: string } | { source: 'text'; value: string; webUrl: string };
  }): Promise<void> {
    const token = await this.accessToken();
    const url = `https://graph.microsoft.com/v1.0${opts.scope}/sendActivityNotification`;
    const body = {
      topic: opts.topic,
      activityType: opts.activityType,
      previewText: { content: opts.previewText },
      templateParameters: opts.templateParameters,
    };
    const response = await this.fetchImpl(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      const errBody = await safeBody(response);
      throw new Error(
        `graph sendActivityNotification ${String(response.status)} scope=${opts.scope} type=${opts.activityType} body=${truncate(errBody, 300)}`,
      );
    }
  }

  /**
   * Resolve a user's primary SMTP email by their AAD object id (app permission `User.Read.All`).
   * Returns the `mail` field — the addressable SMTP address — and deliberately NOT the
   * userPrincipalName, which can differ from the real email. null when the user has no mailbox,
   * the id is unknown, or the read isn't permitted (logged, returns null rather than throwing).
   * Used by the Teams channel to key a Conductor binding by email even in 1:1 chats, where the
   * conversation roster does not expose member email.
   */
  async getUserMail(aadObjectId: string): Promise<string | null> {
    const token = await this.accessToken();
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(aadObjectId)}?$select=mail,proxyAddresses`;
    const response = await this.fetchImpl(url, {
      headers: { Authorization: `Bearer ${token}` },
      // Bounded: this runs in the Teams hot path (awaited before the chat turn) — a hung Graph/Entra
      // must never stall the user's first reply. On timeout the fetch rejects → the caller degrades.
      signal: AbortSignal.timeout(2000),
    });
    if (!response.ok) {
      this.log(`graph getUserMail ${aadObjectId}: ${String(response.status)}`);
      return null;
    }
    const json = (await response.json()) as { mail?: unknown; proxyAddresses?: unknown };
    if (typeof json.mail === 'string' && json.mail.length > 0) return json.mail;
    // `mail` is often unpopulated for on-prem-synced users; the primary SMTP is then the
    // `SMTP:`-prefixed (uppercase = primary) entry in proxyAddresses. `smtp:` are secondary aliases.
    if (Array.isArray(json.proxyAddresses)) {
      const primary = json.proxyAddresses.find(
        (a): a is string => typeof a === 'string' && a.startsWith('SMTP:'),
      );
      if (primary) return primary.slice(5);
    }
    this.log(`graph getUserMail ${aadObjectId}: 200 but no mail / primary SMTP proxyAddress`);
    return null;
  }

  private async accessToken(): Promise<string> {
    // Hot-path default budget (2000 ms) — see AadTokenRequest.timeoutMs.
    return this.tokens.accessToken({
      tenantId: this.tenantId,
      clientId: this.clientId,
      clientSecret: this.clientSecret,
      scope: GRAPH_TOKEN_SCOPE,
    });
  }
}

/**
 * Normalised shape of an attachment as surfaced by
 * `GET /chats/{id}/messages/{id}`. Graph returns a superset of fields;
 * we only care about content-url (SharePoint), name, content-type, and
 * the contentType discriminator (`reference` means a file link).
 */
export interface GraphChatMessageAttachment {
  id: string;
  contentType: string;           // 'reference' for file links, 'messageReference', 'codeSnippet', …
  contentUrl: string;
  name: string;
  teamsFileType?: string;        // usually mirrors file extension
}

function normaliseGraphAttachment(
  raw: unknown,
): GraphChatMessageAttachment | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  const contentUrl = typeof r['contentUrl'] === 'string' ? r['contentUrl'] : '';
  if (!contentUrl) return undefined;
  const id = typeof r['id'] === 'string' ? r['id'] : contentUrl;
  const contentType =
    typeof r['contentType'] === 'string' ? r['contentType'] : 'reference';
  const name =
    typeof r['name'] === 'string' && r['name'].trim().length > 0
      ? r['name']
      : 'unnamed';
  const teamsFileType =
    typeof r['teamsFileType'] === 'string' ? r['teamsFileType'] : undefined;
  const out: GraphChatMessageAttachment = {
    id,
    contentType,
    contentUrl,
    name,
    ...(teamsFileType ? { teamsFileType } : {}),
  };
  return out;
}

/**
 * Encode a SharePoint sharing URL into Graph's `u!<base64url>` shape.
 * See https://learn.microsoft.com/graph/api/shares-get#encoding-sharing-urls.
 */
export function encodeSharingUrl(url: string): string {
  const b64 = Buffer.from(url, 'utf8').toString('base64');
  const b64url = b64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');
  return `u!${b64url}`;
}

/** Read a response body for an error message, tolerating unreadable bodies. */
export async function safeBody(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

function fileNameFromDisposition(header: string | null): string | undefined {
  if (!header) return undefined;
  // Prefer RFC 5987 filename*=UTF-8''<urlenc> over plain filename=
  const star = /filename\*\s*=\s*([^']*)''([^;]+)/i.exec(header);
  if (star && star[2]) {
    try {
      return decodeURIComponent(star[2]).trim();
    } catch {
      // fall through to plain
    }
  }
  const plain = /filename\s*=\s*"?([^";]+)"?/i.exec(header);
  if (plain && plain[1]) return plain[1].trim();
  return undefined;
}

/** Bound a string for error messages (never dump full API bodies). */
export function truncate(value: string, max: number): string {
  if (value.length <= max) return value;
  return `${value.slice(0, max - 1)}…`;
}

/** The id and file attachments of a Graph chat or channel message. */
function parseGraphMessage(
  raw: unknown,
  fallbackId: string,
): { id: string; attachments: GraphChatMessageAttachment[] } {
  const json = (raw ?? {}) as { id?: unknown; attachments?: unknown };
  const attachments = Array.isArray(json.attachments)
    ? json.attachments
        .map((a) => normaliseGraphAttachment(a))
        .filter((a): a is GraphChatMessageAttachment => a !== undefined)
    : [];
  return { id: typeof json.id === 'string' ? json.id : fallbackId, attachments };
}

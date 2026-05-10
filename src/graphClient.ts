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

export class GraphClient {
  private readonly tenantId: string;
  private readonly clientId: string;
  private readonly clientSecret: string;
  private readonly log: (msg: string) => void;
  private readonly fetchImpl: typeof fetch;

  private tokenPromise: Promise<{ token: string; expiresAt: number }> | undefined;

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
    const json = (await response.json()) as {
      id?: unknown;
      attachments?: unknown;
    };
    const attachments = Array.isArray(json.attachments)
      ? json.attachments
          .map((a) => normaliseGraphAttachment(a))
          .filter((a): a is GraphChatMessageAttachment => a !== undefined)
      : [];
    return {
      id: typeof json.id === 'string' ? json.id : messageId,
      attachments,
    };
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

  private async accessToken(): Promise<string> {
    const now = Date.now();
    const cached = await this.tokenPromise?.catch(() => undefined);
    if (cached && cached.expiresAt - 60_000 > now) return cached.token;

    this.tokenPromise = this.fetchAccessToken();
    try {
      const fresh = await this.tokenPromise;
      return fresh.token;
    } catch (err) {
      this.tokenPromise = undefined;
      throw err;
    }
  }

  private async fetchAccessToken(): Promise<{
    token: string;
    expiresAt: number;
  }> {
    const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(this.tenantId)}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
    });
    const response = await this.fetchImpl(tokenUrl, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });
    if (!response.ok) {
      const body = await safeBody(response);
      throw new Error(
        `graph token ${String(response.status)} body=${truncate(body, 200)}`,
      );
    }
    const json = (await response.json()) as {
      access_token?: unknown;
      expires_in?: unknown;
    };
    if (typeof json.access_token !== 'string' || json.access_token.length === 0) {
      throw new Error('graph token: no access_token in response');
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

async function safeBody(response: Response): Promise<string> {
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

function truncate(value: string, max: number): string {
  if (value.length <= max) return value;
  return `${value.slice(0, max - 1)}…`;
}

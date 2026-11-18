/**
 * `probeBot` of `teamsProvisioner@1` (since 0.9.2): does the Bot Connector
 * know this bot?
 *
 * WHY IT EXISTS. Azure Resource Manager can report a bot resource as healthy
 * (endpoint right, Teams channel enabled, state Succeeded) while the Bot
 * Connector rejects every call the bot makes with `401 Authorization has been
 * denied for this request` and Teams shows "invalid bot". That happened to a
 * bot re-created under an app id that had a bot before, and it did not heal
 * (byte5ai/omadia#2043, rollout defect D10). Nothing in Graph or ARM shows
 * it. The only place that knows is the Bot Connector, asked AS the bot.
 *
 * WHAT IT DOES. Two requests, both with the bot's own identity:
 *
 * 1. a client-credentials token for `https://api.botframework.com/.default`
 *    from the tenant's token endpoint, with the bot's app id and the secret
 *    this connector stored under `teams_bot_password:<appId>`. Source:
 *    "Authenticate requests with the Bot Connector API", single-tenant bot
 *    (learn.microsoft.com/azure/bot-service/rest-api/bot-framework-rest-connector-authentication).
 * 2. `POST {serviceUrl}/v3/conversations` naming a member that does not exist.
 *    Body shape from "Send proactive messages"
 *    (learn.microsoft.com/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages).
 *
 * NO MEMBERLESS PROBE. The Connector reference lists one operation that needs
 * neither a member nor a conversation, `GET /v3/conversations`, and says of
 * it: "only Direct Line and Web Chat support the get conversations endpoint"
 * (bot-framework-rest-connector-api-reference). Nothing is known about what
 * Teams answers there for either kind of bot, so it is not used.
 *
 * THE MAPPING, and what it rests on:
 *
 * - `401` → `rejected`. OBSERVED at the rollout for the dead bot, with a real
 *   member id, repeatedly over 52 minutes.
 * - any `2xx`, or `400` / `403` / `404` / `405` → `known`. OBSERVED for
 *   healthy bots with a real member id: `403 BadArgument` "Bot is not
 *   installed in user's personal scope", and `201`.
 * - anything else, a timeout, no token for a reason other than the
 *   credential → `unverified`.
 *
 * NOT VERIFIED AGAINST A TENANT: what a healthy bot answers for the
 * FABRICATED member id this probe sends. The mapping does not depend on which
 * 4xx that is; it would be wrong only if a healthy bot answered 401 to it,
 * and 401 is the Connector's answer about the caller, not about the member.
 *
 * It can create nothing: the member does not exist, and no activity is sent.
 *
 * SECRET HANDLING. The bot's secret is read from the vault, sent to the token
 * endpoint and dropped. The token goes to the Connector and is dropped. Every
 * `detail` passes `redactSecrets` and is bounded.
 */

import type { SecretsAccessor } from '@omadia/plugin-api';

import { truncate } from '../graphClient.js';
import { requireAppId } from './appRegistrationSupport.js';
import { redactSecrets } from './redact.js';
import { secretRefForApp } from './secretStore.js';

/** Bot Connector service URL of Teams in the public cloud. */
export const TEAMS_BOT_CONNECTOR_SERVICE_URL = 'https://smba.trafficmanager.net/teams';
/** Token audience of the Bot Connector. */
export const BOT_CONNECTOR_TOKEN_SCOPE = 'https://api.botframework.com/.default';
/** Member id sent with the probe. `29:` is the Teams user prefix; this one exists nowhere. */
export const PROBE_MEMBER_ID = '29:omadia-probe';

const PROBE_TIMEOUT_MS = 10_000;
const DETAIL_MAX_CHARS = 500;
/** Connector statuses that are an answer about the REQUEST, from a bot it knows. */
const KNOWN_BOT_STATUSES: ReadonlySet<number> = new Set([400, 403, 404, 405]);

export interface ProbeBotInput {
  /** Entra application (client) id the bot authenticates as. */
  readonly appId: string;
}

/** What the Bot Connector said about the bot. */
export type ProbeBotResult =
  | { readonly status: 'known'; readonly detail?: string }
  | { readonly status: 'rejected'; readonly httpStatus?: number; readonly detail?: string }
  | { readonly status: 'unverified'; readonly detail?: string };

export interface BotProbeOptions {
  /** Tenant the bot's app registration lives in. */
  readonly tenantId: string;
  /** Plugin vault accessor — holds `teams_bot_password:<appId>`. */
  readonly secrets: SecretsAccessor;
  readonly fetchImpl?: typeof fetch;
  /** Bot Connector base URL. Default: {@link TEAMS_BOT_CONNECTOR_SERVICE_URL}. */
  readonly serviceUrl?: string;
  readonly log?: (msg: string) => void;
}

/** Ask the Bot Connector, as the bot, whether it knows the bot. Never throws. */
export async function probeBot(
  opts: BotProbeOptions,
  input: ProbeBotInput,
): Promise<ProbeBotResult> {
  const appId = requireAppId(input.appId);
  const fetchImpl = opts.fetchImpl ?? fetch;
  try {
    const secret = await opts.secrets.get(secretRefForApp(appId));
    if (secret === undefined || secret === '') {
      return { status: 'unverified', detail: `no stored secret for app id '${appId}'` };
    }
    const token = await botToken(fetchImpl, opts.tenantId, appId, secret);
    if (typeof token !== 'string') return token;
    return await askConnector(fetchImpl, opts, appId, token);
  } catch (err) {
    // Network failure, timeout, an unreadable vault: nothing was learned.
    const detail = safeDetail(err instanceof Error ? `${err.name}: ${err.message}` : String(err));
    opts.log?.(`provisioner bot.probe: '${appId}' not verified (${detail})`);
    return { status: 'unverified', detail };
  }
}

/** The bot's Connector token, or the probe result when there is none. */
async function botToken(
  fetchImpl: typeof fetch,
  tenantId: string,
  appId: string,
  secret: string,
): Promise<string | ProbeBotResult> {
  const response = await fetchImpl(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: appId,
        client_secret: secret,
        scope: BOT_CONNECTOR_TOKEN_SCOPE,
      }).toString(),
      signal: AbortSignal.timeout(PROBE_TIMEOUT_MS),
    },
  );
  const text = await response.text();
  if (response.ok) {
    const token = (parseJson(text) as Record<string, unknown> | undefined)?.['access_token'];
    return typeof token === 'string' && token !== ''
      ? token
      : { status: 'unverified', detail: 'the token endpoint answered without an access token' };
  }
  const body = parseJson(text) as Record<string, unknown> | undefined;
  const description = typeof body?.['error_description'] === 'string' ? body['error_description'] : text;
  const detail = safeDetail(`token ${String(response.status)}: ${description}`);
  // `invalid_client` is Entra's verdict on the bot's own credential: the app
  // is unknown in the tenant or the secret is not its secret (AADSTS code in
  // the detail). Everything else says nothing about the bot.
  return body?.['error'] === 'invalid_client'
    ? { status: 'rejected', httpStatus: response.status, detail }
    : { status: 'unverified', detail };
}

async function askConnector(
  fetchImpl: typeof fetch,
  opts: BotProbeOptions,
  appId: string,
  token: string,
): Promise<ProbeBotResult> {
  const serviceUrl = (opts.serviceUrl ?? TEAMS_BOT_CONNECTOR_SERVICE_URL).replace(/\/+$/, '');
  const response = await fetchImpl(`${serviceUrl}/v3/conversations`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'content-type': 'application/json' },
    body: JSON.stringify({
      bot: { id: `28:${appId}` },
      members: [{ id: PROBE_MEMBER_ID }],
      channelData: { tenant: { id: opts.tenantId } },
      isGroup: false,
    }),
    signal: AbortSignal.timeout(PROBE_TIMEOUT_MS),
  });
  const detail = safeDetail(`connector ${String(response.status)}: ${await response.text()}`);
  if (response.status === 401) {
    opts.log?.(`provisioner bot.probe: the Bot Connector rejects '${appId}' (${detail})`);
    return { status: 'rejected', httpStatus: 401, detail };
  }
  if (response.ok || KNOWN_BOT_STATUSES.has(response.status)) {
    return { status: 'known', detail };
  }
  return { status: 'unverified', detail };
}

function safeDetail(text: string): string {
  return truncate(redactSecrets(text), DETAIL_MAX_CHARS);
}

function parseJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

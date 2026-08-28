/**
 * Azure-bot step of `teamsProvisioner@1` (epic byte5ai/omadia#860, capability
 * issue byte5ai/omadia-m365-connector#3): create — and roll back — the
 * "1 agent = 1 Azure bot" resource via ARM REST.
 *
 * `createBot` chains TWO ARM writes:
 *
 *   1. `PUT .../Microsoft.BotService/botServices/{name}` — the bot resource
 *   2. `PUT .../botServices/{name}/channels/MsTeamsChannel` — Teams enablement
 *
 * Both go through the shared {@link ProvisioningHttp} choke point (ARM
 * audience `https://management.azure.com/.default`, Retry-After-honouring 429
 * backoff, 403 → `ConsentMissingError` with the RBAC action, and the
 * 201/202 + `Azure-AsyncOperation` long-running poll mode for the PUTs).
 * This module opens no second token cache and does no fetch of its own.
 *
 * PINNED BODY LITERALS (spec, non-negotiable): `kind: 'registration'`,
 * `sku: { name: 'F0' }`, `properties.msaAppType: 'SingleTenant'` — the ARM
 * mirror of the SingleTenant app-registration invariant (`types.ts` module
 * doc). MultiSaaS/MultiTenant bots are deliberately not expressible.
 *
 * GRACEFUL DEGRADATION — this whole client is unreachable in
 * registration-only mode: when the ARM setup fields are absent
 * (`ArmConfigResult` of the config unit is `kind: 'registration-only'`),
 * every method returns that typed {@link RegistrationOnlyOutcome} BEFORE any
 * token acquisition or fetch — never a crash. Callers pre-flight via
 * {@link BotServiceClient.canCreateBots}.
 *
 * IDEMPOTENCY — ARM `PUT botServices/{name}` is an upsert keyed on the bot
 * handle: `201` means created now, `200` means the resource already existed →
 * the `Idempotent<AzureBot>` `'already-existed'` outcome. A `409` is NOT the
 * re-run signal here (that is Graph semantics) — for bot handles it means the
 * globally-unique name is taken by a FOREIGN resource, which is a real error.
 * `deleteBot` is the rollback half and is idempotent: a 404 answers
 * `'already-deleted'`, never an error, so job retries and double rollbacks
 * are safe.
 *
 * Style precedent: `appRegistration.ts` — options-bag constructor, hand-rolled
 * REST, non-2xx → typed/plain errors with step + status.
 */

import { ARM_MANAGEMENT_HOST, type ArmConfigResult } from './config.js';
import { BotHandleUnavailableError, ProvisioningRequestError } from './errors.js';
import type { ProvisioningHttp, ProvisioningOkResponse } from './http.js';
import type {
  AzureBot,
  BotProvisioningOutcome,
  CreateBotInput,
  IdempotentOutcome,
  RegistrationOnlyOutcome,
} from './types.js';

/** Stable ARM api-version for `Microsoft.BotService` (bot + channels). */
export const BOT_SERVICE_API_VERSION = '2022-09-15';

/** Pinned ARM body literal: registration-style bot (no hosted web app). */
export const BOT_KIND_REGISTRATION = 'registration';
/** Pinned ARM body literal: free tier — bots are per-agent, F0 suffices. */
export const BOT_SKU_F0 = 'F0';
/** Pinned ARM body literal: the SingleTenant invariant, ARM-side. */
export const BOT_MSA_APP_TYPE_SINGLE_TENANT = 'SingleTenant';
/**
 * Bot Framework handle bounds (byte5ai/omadia#921). Stricter than the ARM
 * resource-name rule the same string also has to satisfy: 4-42 chars,
 * alphanumerics and hyphens only, no leading/trailing hyphen.
 */
export const BOT_HANDLE_MIN_LENGTH = 4;
export const BOT_HANDLE_MAX_LENGTH = 42;
/** The handle grammar both ARM and Bot Framework accept. */
export const BOT_HANDLE_RE = /^[A-Za-z0-9](?:[A-Za-z0-9-]{2,40})[A-Za-z0-9]$/;

/** ARM channel sub-resource name for the Teams enablement PUT. */
export const MS_TEAMS_CHANNEL_NAME = 'MsTeamsChannel';

/** RBAC action reported when the bot PUT/GET answers 403. */
export const BOT_SERVICES_WRITE_ACTION = 'Microsoft.BotService/botServices/write';
/** RBAC action reported when the channel PUT answers 403. */
export const BOT_SERVICES_CHANNELS_WRITE_ACTION =
  'Microsoft.BotService/botServices/channels/write';
/** RBAC action reported when the bot DELETE answers 403. */
export const BOT_SERVICES_DELETE_ACTION = 'Microsoft.BotService/botServices/delete';
/** RBAC action reported when the status GET answers 403. */
export const BOT_SERVICES_READ_ACTION = 'Microsoft.BotService/botServices/read';

/**
 * Idempotent delete signal — the delete-direction mirror of
 * `IdempotentOutcome` in `types.ts` (same shape as the app-registration
 * unit's `DeleteAppRegistrationOutcome`).
 */
export type DeleteBotOutcome = 'deleted' | 'already-deleted';

/** Successful (idempotent) bot deletion. */
export interface BotDeletedResult {
  readonly kind: 'deleted';
  readonly outcome: DeleteBotOutcome;
}

/** Result of `deleteBot`: idempotent delete or registration-only fallback. */
export type DeleteBotResult = BotDeletedResult | RegistrationOnlyOutcome;

/** Status probe hit — the bot as ARM sees it. */
export interface BotFoundResult {
  readonly kind: 'found';
  readonly bot: AzureBot;
}

/** Status probe miss — no such bot resource (or already deleted). */
export interface BotNotFoundResult {
  readonly kind: 'not-found';
}

/** Result of `getBot`: found/not-found or registration-only fallback. */
export type GetBotResult = BotFoundResult | BotNotFoundResult | RegistrationOnlyOutcome;

export interface BotServiceClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  /**
   * The ARM mode decision of the config unit. `'registration-only'` makes
   * every method answer that outcome without touching the network.
   */
  readonly armConfig: ArmConfigResult;
  readonly log?: (msg: string) => void;
}

/**
 * Create/enable, probe and roll back Azure Bot resources via ARM REST — the
 * Azure-bot step of the chain. One step client per provisioner; ordering
 * across chain steps stays middleware-side (agent factory,
 * byte5ai/omadia#863-865).
 */
export class BotServiceClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly armConfig: ArmConfigResult;
  private readonly log: (msg: string) => void;

  constructor(opts: BotServiceClientOptions) {
    this.http = opts.http;
    this.armConfig = opts.armConfig;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /** `true` when the ARM setup fields are configured — feeds `TeamsProvisioner.canCreateBots`. */
  get canCreateBots(): boolean {
    return this.armConfig.kind === 'configured';
  }

  /**
   * The full Azure-bot step: PUT the `Microsoft.BotService/botServices`
   * resource (pinned `kind: 'registration'`, `sku: 'F0'`,
   * `msaAppType: 'SingleTenant'`), then enable the `MsTeamsChannel`
   * sub-resource with a second ARM PUT.
   *
   * The `Idempotent` outcome reflects the BOT resource: ARM PUT is an upsert,
   * so `200` (as opposed to a created/polled `201`/`202`) means the handle
   * already existed — the channel enablement still runs (itself idempotent,
   * a 409 counts as already-enabled).
   *
   * On a channel-enablement failure a bot THIS call created is rolled back
   * (best-effort, logged), and the original error is rethrown; a
   * pre-existing bot is left untouched.
   *
   * In registration-only mode the typed `RegistrationOnlyOutcome` is returned
   * before any token acquisition.
   */
  async createBot(input: CreateBotInput): Promise<BotProvisioningOutcome> {
    if (this.armConfig.kind === 'registration-only') {
      return this.armConfig;
    }
    const botName = requireBotName(input.botName);
    requireNonEmpty(input.displayName, 'displayName');
    requireNonEmpty(input.msaAppId, 'msaAppId');
    requireNonEmpty(input.msaAppTenantId, 'msaAppTenantId');
    requireNonEmpty(input.messagingEndpoint, 'messagingEndpoint');

    let response;
    try {
      response = await this.http.request({
        resource: 'arm',
        method: 'PUT',
        url: this.botUrl(botName),
        step: 'botServices.put',
        jsonBody: {
          location: this.armConfig.region,
          kind: BOT_KIND_REGISTRATION,
          sku: { name: BOT_SKU_F0 },
          properties: {
            displayName: input.displayName,
            endpoint: input.messagingEndpoint,
            msaAppId: input.msaAppId,
            // The SingleTenant invariant, ARM-side — always with the app tenant.
            msaAppType: BOT_MSA_APP_TYPE_SINGLE_TENANT,
            msaAppTenantId: input.msaAppTenantId,
          },
        },
        missingScopesOn403: [BOT_SERVICES_WRITE_ACTION],
        pollLongRunning: true,
      });
    } catch (err) {
      // ARM reports a taken GLOBAL handle as a 400 `InvalidBotData`, not a
      // 409 — an untyped client error that reads like a payload bug and gets
      // retried as one (byte5ai/omadia#921). Promote it to the typed verdict
      // so the caller fails fast with an explanation.
      if (isBotNameTakenError(err)) {
        throw new BotHandleUnavailableError(botName, err.status, err);
      }
      throw err;
    }
    if (response.kind === 'conflict') {
      // NOT the idempotent signal: our own re-run upserts with 200. A 409
      // means the globally-unique handle belongs to a foreign resource —
      // same verdict as the 400 above, so it carries the same typed error.
      throw new BotHandleUnavailableError(botName, 409);
    }
    const outcome: IdempotentOutcome =
      response.status === 200 ? 'already-existed' : 'created';
    const bot = this.parseBot(response, 'botServices.put', input);

    try {
      await this.enableMsTeamsChannel(botName);
    } catch (err) {
      await this.rollbackPartialCreate(botName, outcome);
      throw err;
    }

    return { kind: 'provisioned', bot: { outcome, value: bot } };
  }

  /**
   * Rollback half: delete the bot resource (channels are sub-resources and go
   * with it). Idempotent — a 404 answers `'already-deleted'`, never an error,
   * so job retries and double rollbacks are safe. In registration-only mode
   * the typed `RegistrationOnlyOutcome` is returned (nothing was ever
   * provisioned, and there is no ARM credential to ask with).
   */
  async deleteBot(botName: string): Promise<DeleteBotResult> {
    if (this.armConfig.kind === 'registration-only') {
      return this.armConfig;
    }
    const name = requireBotName(botName);
    const response = await this.http.request({
      resource: 'arm',
      method: 'DELETE',
      url: this.botUrl(name),
      step: 'botServices.delete',
      missingScopesOn403: [BOT_SERVICES_DELETE_ACTION],
      // Already gone = already rolled back — the idempotent delete signal.
      extraOkStatuses: [404],
    });
    if (response.kind === 'conflict') {
      throw new Error('arm botServices.delete unexpected 409');
    }
    return {
      kind: 'deleted',
      outcome: response.status === 404 ? 'already-deleted' : 'deleted',
    };
  }

  /**
   * Status probe: the bot resource as ARM sees it, `'not-found'` when the
   * handle does not (or no longer does) exist, or the registration-only
   * outcome when ARM is unconfigured.
   */
  async getBot(botName: string): Promise<GetBotResult> {
    if (this.armConfig.kind === 'registration-only') {
      return this.armConfig;
    }
    const name = requireBotName(botName);
    const response = await this.http.request({
      resource: 'arm',
      method: 'GET',
      url: this.botUrl(name),
      step: 'botServices.get',
      missingScopesOn403: [BOT_SERVICES_READ_ACTION],
      extraOkStatuses: [404],
    });
    if (response.kind === 'conflict') {
      throw new Error('arm botServices.get unexpected 409');
    }
    if (response.status === 404) {
      return { kind: 'not-found' };
    }
    return { kind: 'found', bot: this.parseBot(response, 'botServices.get') };
  }

  /**
   * Second ARM PUT: enable the `MsTeamsChannel` sub-resource. Idempotent by
   * upsert; a 409 (concurrent enablement) counts as already-enabled — the
   * desired state is reached either way.
   */
  private async enableMsTeamsChannel(botName: string): Promise<void> {
    if (this.armConfig.kind !== 'configured') {
      // Unreachable from createBot; kept for the compiler's narrowing.
      throw new Error('arm channels.put requires a configured ARM mode');
    }
    await this.http.request({
      resource: 'arm',
      method: 'PUT',
      url: `${this.botUrl(botName, false)}/channels/${MS_TEAMS_CHANNEL_NAME}?api-version=${BOT_SERVICE_API_VERSION}`,
      step: 'botServices.channels.put',
      jsonBody: {
        location: this.armConfig.region,
        properties: {
          channelName: MS_TEAMS_CHANNEL_NAME,
          properties: { isEnabled: true },
        },
      },
      missingScopesOn403: [BOT_SERVICES_CHANNELS_WRITE_ACTION],
      pollLongRunning: true,
    });
  }

  /**
   * Best-effort rollback after a partial create: only a bot THIS call created
   * (`'created'`) is deleted — a pre-existing handle is never destroyed by a
   * failed channel enablement. Failures are logged; the caller rethrows the
   * original error.
   */
  private async rollbackPartialCreate(
    botName: string,
    outcome: IdempotentOutcome,
  ): Promise<void> {
    if (outcome !== 'created') return;
    try {
      await this.deleteBot(botName);
    } catch (err) {
      this.log(
        `provisioner rollback: botServices.delete of '${botName}' failed: ${String(err)}`,
      );
    }
  }

  /** Full ARM resource URL of one bot (validated handle, encoded path parts). */
  private botUrl(botName: string, withApiVersion = true): string {
    if (this.armConfig.kind !== 'configured') {
      throw new Error('arm request requires a configured ARM mode');
    }
    const base =
      `https://${ARM_MANAGEMENT_HOST}/subscriptions/` +
      `${encodeURIComponent(this.armConfig.subscriptionId)}/resourceGroups/` +
      `${encodeURIComponent(this.armConfig.resourceGroup)}/providers/` +
      `Microsoft.BotService/botServices/${botName}`;
    return withApiVersion ? `${base}?api-version=${BOT_SERVICE_API_VERSION}` : base;
  }

  /**
   * Parse + validate an ARM `botServices` resource into the shared
   * {@link AzureBot}. The polled bodiless-202 path re-GETs the finished
   * resource, so a real body is always present; `input` (createBot only)
   * backfills properties an ARM response may omit on the async path.
   */
  private parseBot(
    response: ProvisioningOkResponse,
    step: string,
    input?: CreateBotInput,
  ): AzureBot {
    const body = asRecord(response.json, step);
    const properties =
      typeof body['properties'] === 'object' && body['properties'] !== null
        ? (body['properties'] as Record<string, unknown>)
        : {};
    const msaAppId = stringField(properties, 'msaAppId') ?? input?.msaAppId;
    const messagingEndpoint =
      stringField(properties, 'endpoint') ?? input?.messagingEndpoint;
    const botName = stringField(body, 'name') ?? input?.botName;
    if (msaAppId === undefined || messagingEndpoint === undefined || botName === undefined) {
      throw new Error(`arm ${step}: response body missing name/msaAppId/endpoint`);
    }
    return {
      botName,
      resourceId: requireStringField(body, 'id', step),
      msaAppId,
      messagingEndpoint,
    };
  }
}

/**
 * Does this failure mean the global bot handle is taken?
 *
 * ARM answers `400` with `{"error":{"code":"InvalidBotData","message":"Bot is
 * not valid. Errors: The bot name is already registered to another bot
 * application."}}`. The status alone cannot carry the meaning — a 400 is also
 * how ARM reports a genuinely malformed body — so the error code plus the
 * registered-to-another phrase are both required before promoting.
 */
function isBotNameTakenError(err: unknown): err is ProvisioningRequestError {
  if (!(err instanceof ProvisioningRequestError)) return false;
  if (err.status !== 400) return false;
  const body = err.message.toLowerCase();
  return (
    body.includes('invalidbotdata') && body.includes('already registered to another')
  );
}

function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

/**
 * One string has to satisfy TWO different rule sets, and the stricter one
 * wins (byte5ai/omadia#921).
 *
 * The ARM resource name for `Microsoft.BotService/botServices` tolerates 2-64
 * chars of `[A-Za-z0-9._-]`. The Bot Framework HANDLE the same string becomes
 * is narrower: {@link BOT_HANDLE_MIN_LENGTH}-{@link BOT_HANDLE_MAX_LENGTH}
 * chars, alphanumerics and hyphens only. Validating against the ARM rule (as
 * this did until #921) lets a `.`/`_` handle — or a 60-char one — through the
 * client and into a 400 from the service, which is precisely the class of
 * late, opaque failure this module exists to prevent.
 *
 * Callers compose the handle; this is the boundary that proves the
 * composition is expressible. See `buildBotHandle` in the middleware job
 * runner for the naming convention itself.
 */
function requireBotName(botName: string): string {
  requireNonEmpty(botName, 'botName');
  if (!BOT_HANDLE_RE.test(botName)) {
    throw new Error(
      `invalid_argument: 'botName' must be ${String(BOT_HANDLE_MIN_LENGTH)}-` +
        `${String(BOT_HANDLE_MAX_LENGTH)} chars of [A-Za-z0-9-] starting and ` +
        'ending with a letter or digit — Azure bot handles are stricter than ' +
        'ARM resource names (no dots, no underscores)',
    );
  }
  return botName;
}

function asRecord(json: unknown, step: string): Record<string, unknown> {
  if (json === null || typeof json !== 'object') {
    throw new Error(`arm ${step}: unexpected empty/non-object response body`);
  }
  return json as Record<string, unknown>;
}

function stringField(
  body: Record<string, unknown>,
  field: string,
): string | undefined {
  const value = body[field];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function requireStringField(
  body: Record<string, unknown>,
  field: string,
  step: string,
): string {
  const value = stringField(body, field);
  if (value === undefined) {
    throw new Error(`arm ${step}: response body missing '${field}'`);
  }
  return value;
}

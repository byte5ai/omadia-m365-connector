/**
 * `teamsProvisioner@1` — capability assembly + public surface of the
 * provisioning subsystem (wiring unit of wave W0b, epic byte5ai/omadia#860,
 * capability issue byte5ai/omadia-m365-connector#3).
 *
 * {@link createTeamsProvisioner} wires the step clients around ONE shared
 * {@link ProvisioningHttp} choke point (single token cache, one 403/409/429
 * matrix) and returns the {@link TeamsProvisionerAccessor} that `activate()`
 * publishes under {@link TEAMS_PROVISIONER_SERVICE_NAME}. The provisioning
 * STATE MACHINE (ordering, persistence, retries across steps) lives
 * middleware-side in the agent factory (byte5ai/omadia#863-865) — this
 * surface exposes exactly the individual chain steps.
 *
 * SERVICE-BOUNDARY INVARIANT — no cleartext secrets. The app-registration
 * step persists the generated password via `ctx.secrets` and hands back the
 * opaque `secretRef` (`teams_bot_password:<appId>`); the raw secret value
 * never crosses the service boundary. This deliberately supersedes the
 * `addClientSecret → AppClientSecret.secretText` sketch in `types.ts` /
 * `docs/teams-provisioner.md` — see the secret-store unit's rationale in
 * `secretStore.ts`.
 *
 * IDEMPOTENCY — `createAppRegistration` is safe to re-run: a taken
 * `uniqueName` adopts the existing registration (including one restored from
 * the recycle bin), Entra's replication windows are polled through rather
 * than reported as failures, and a transient failure never rolls the
 * registration back (byte5ai/omadia#916). Callers that persist the app id
 * should pass `onRegistrationCreated` so an interrupted chain leaves a
 * resumable row.
 *
 * GRACEFUL DEGRADATION — registration-only mode is observable on the
 * published capability: `canCreateBots` is `false` and `createBot` answers
 * the typed `RegistrationOnlyOutcome` without touching the network.
 * Activation itself NEVER throws over missing ARM setup fields.
 */

import type { SecretsAccessor } from '@omadia/plugin-api';

import type { AadClientCredential } from '../graphClient.js';
import {
  AppRegistrationClient,
  type CreateAppRegistrationInput,
  type DeleteAppRegistrationInput,
  type DeleteAppRegistrationResult,
  type ProvisionedAppRegistration,
} from './appRegistration.js';
import {
  buildAppPackage,
  type BuildAppPackageInput,
} from './appPackage.js';
import {
  BotServiceClient,
  type DeleteBotResult,
  type GetBotResult,
} from './botService.js';
import {
  CatalogUploadClient,
  type DelegatedCatalogUploadResult,
  type GetCatalogAppInput,
  type GetCatalogAppResult,
  type UploadToCatalogDelegatedInput,
} from './catalog.js';
import {
  DelegatedAuthClient,
  adminConsentUrl,
  describeSignInStatus,
  revokeInstructions,
  type DelegatedRevokeResult,
  type DelegatedSignInStatus,
  type DelegatedTokenSet,
  type DeviceCodeFlowHandle,
  type DeviceCodePollResult,
  type DeviceCodeStart,
} from './delegatedAuth.js';
import { PublisherAppClient } from './publisherApp.js';
import { isArmConfigured, type ArmConfigResult } from './config.js';
import {
  TeamLookupClient,
  type GetTeamInput,
  type GetTeamResult,
} from './teamLookup.js';
import { ProvisioningHttp } from './http.js';
import {
  TeamInstallClient,
  type InstallToTeamRequest,
  type UninstallFromTeamInput,
  type UninstallFromTeamResult,
} from './install.js';
import type {
  AppRegistration,
  BotProvisioningOutcome,
  CatalogTeamsApp,
  CreateBotInput,
  Idempotent,
  TeamAppInstallation,
  TenantMode,
  UploadToCatalogInput,
} from './types.js';

/**
 * The service object published under the ServiceRegistry key
 * `teamsProvisioner` (manifest capability ref `teamsProvisioner@1`).
 *
 * One method per chain step — the caller owns ordering, persistence and
 * retries across steps. Every method may throw the typed errors from
 * `errors.ts` (`ConsentMissingError` on 403, `ProvisioningThrottledError`
 * on exhausted 429 backoff, `CapabilityUnavailableError` when the vault has
 * no write capability); anything else propagates verbatim.
 *
 * Deviation from the `TeamsProvisioner` interface sketch in `types.ts`:
 * register-app + add-secret are ONE step ({@link createAppRegistration},
 * with rollback), and only the vault `secretRef` — never the cleartext
 * secret — is returned. Delete/get counterparts are exposed so the agent
 * factory can roll back and probe idempotently.
 */
export interface TeamsProvisionerAccessor {
  /** Where this provisioner registers apps (`'customer'` in normal installs). */
  readonly tenantMode: TenantMode;
  /** `true` when the ARM setup fields are configured (bot creation possible). */
  readonly canCreateBots: boolean;

  /**
   * Chain step 1+2 — register (or find via `uniqueName`) the SingleTenant
   * Entra app, rotate a client secret into the vault (only the `secretRef`
   * is returned) and ensure the service principal exists.
   *
   * Idempotent and interruption-friendly: a taken `uniqueName` adopts the
   * existing registration, Entra's replication windows are polled through,
   * and rollback of a partial failure is narrow — never on a transient error,
   * never deleting a registration that carries a `uniqueName`
   * (byte5ai/omadia#916).
   */
  createAppRegistration(
    input: CreateAppRegistrationInput,
  ): Promise<Idempotent<ProvisionedAppRegistration>>;

  /** Rollback counterpart of {@link createAppRegistration} (idempotent). */
  deleteAppRegistration(
    input: DeleteAppRegistrationInput,
  ): Promise<DeleteAppRegistrationResult>;

  /** Probe an existing registration by app (client) id. */
  getAppRegistration(
    appId: string,
    tenantMode: TenantMode,
  ): Promise<AppRegistration | undefined>;

  /**
   * Render the per-agent Teams app package (manifest template + icons) into
   * an in-memory zip for {@link uploadToCatalog}. Pure, no network.
   */
  buildAppPackage(input: BuildAppPackageInput): Uint8Array;

  /**
   * Chain step 3 — create the Azure bot + enable the Teams channel via ARM.
   * Answers the typed `RegistrationOnlyOutcome` when ARM is unconfigured.
   */
  createBot(input: CreateBotInput): Promise<BotProvisioningOutcome>;

  /** Rollback counterpart of {@link createBot} (idempotent). */
  deleteBot(botName: string): Promise<DeleteBotResult>;

  /** Probe an existing bot resource by handle. */
  getBot(botName: string): Promise<GetBotResult>;

  /**
   * Chain step 4 — publish the app package into the tenant catalog.
   *
   * SINCE 0.6.0 this step needs a DELEGATED access token in
   * `input.delegatedAccessToken` (byte5ai/omadia#924). Graph supports
   * `POST /appCatalogs/teamsApps` for delegated permissions only, so the
   * app-only call this used to make is refused by the service regardless of
   * consent. The parameter shape is unchanged — the field is optional — but a
   * call without a token now throws `DelegatedSignInRequiredError` instead of
   * reaching Graph.
   *
   * Prefer {@link uploadToCatalogDelegated}, which takes the stored token set
   * and handles the refresh.
   */
  uploadToCatalog(
    input: UploadToCatalogInput,
  ): Promise<Idempotent<CatalogTeamsApp>>;

  /**
   * Chain step 4, the form a caller with STORED credentials wants (since
   * 0.6.0): refresh the delegated access token when stale, publish, and return
   * the possibly-rotated token set for persistence.
   *
   * FEATURE-DETECT it (`typeof provisioner.uploadToCatalogDelegated ===
   * 'function'`) — same reason as {@link uninstallFromTeam}: the middleware
   * mirrors this contract structurally rather than importing it.
   */
  uploadToCatalogDelegated(
    input: UploadToCatalogDelegatedInput,
  ): Promise<DelegatedCatalogUploadResult>;

  /**
   * Begin the one-time admin sign-in that makes catalog publishing possible
   * (since 0.6.0, byte5ai/omadia#924).
   *
   * Provisions the tenant's publisher app on first use (idempotent, no secret,
   * one delegated scope) and starts an RFC 8628 device-code flow against it.
   * Show `userCode` and `verificationUri` to the operator; keep `flowHandle`
   * like a password and feed it to {@link pollDelegatedSignIn}.
   *
   * Call it only when the operator is ready — the code expires ~15 minutes
   * from this call, not from when they read it.
   */
  startDelegatedSignIn(input?: {
    /** Portal/consent-screen name of the publisher app. */
    readonly displayName?: string;
  }): Promise<DeviceCodeStart>;

  /**
   * Continue a device-code sign-in (since 0.6.0). Costs no Graph call — every
   * piece of state lives in the handle, so any process/instance can poll a flow
   * another one started.
   *
   * `'pending'` carries the interval to wait; `'succeeded'` carries the tokens
   * to persist; `'expired'` and `'declined'` are terminal and need a new flow.
   * Read `reason` on `'declined'` before blaming the admin — a tenant that
   * blocks device code flow by Conditional Access lands there too.
   */
  pollDelegatedSignIn(input: {
    readonly flowHandle: DeviceCodeFlowHandle;
  }): Promise<DeviceCodePollResult>;

  /**
   * What the stored credential currently is (since 0.6.0). Synchronous and
   * side-effect free — it inspects what the caller passes in and makes no
   * network call, because a status widget must not cost a token round trip.
   */
  getDelegatedSignInStatus(input: {
    /** The persisted token set, or `undefined` when nobody has signed in. */
    readonly tokens?: DelegatedTokenSet;
  }): DelegatedSignInStatus;

  /**
   * Renew a stored credential explicitly (since 0.6.0). Returns the ROTATED
   * token set — persist it, or Entra's rotation eventually strands the caller
   * on a dead refresh token.
   */
  refreshDelegatedToken(input: {
    readonly tokens: DelegatedTokenSet;
  }): Promise<DelegatedTokenSet>;

  /**
   * End the delegated sign-in (since 0.6.0).
   *
   * Honest about its limits: the connector holds no tokens, so revoking is
   * primarily the caller DISCARDING what it stored — that is what the result
   * instructs. Server-side revocation is deliberately not attempted here.
   * Removing the tenant grant needs `DelegatedPermissionGrant.ReadWrite.All`,
   * which this connector does not have and should not get for one upload; and
   * deleting the publisher app would reserve its `uniqueName` for 30 days and
   * lock the tenant out of signing in again (byte5ai/omadia#916). The result
   * therefore carries the portal URL where an admin can withdraw consent.
   */
  revokeDelegatedSignIn(input: {
    readonly tokens?: DelegatedTokenSet;
  }): DelegatedRevokeResult;

  /**
   * Lookup probe for step 4 — resolve an EXISTING catalog app by manifest id
   * (`externalId`) without uploading a package. `{ found: false }` is a plain
   * outcome, never an exception.
   */
  getCatalogApp(input: GetCatalogAppInput): Promise<GetCatalogAppResult>;

  /** Chain step 5 — install the catalog app into one team. */
  installToTeam(
    input: InstallToTeamRequest,
  ): Promise<Idempotent<TeamAppInstallation>>;

  /**
   * Reverse of {@link installToTeam} (since 0.4.0, byte5ai/omadia#900) —
   * remove the catalog app from one team, keyed by the same
   * (teamId, teamsAppId) pair. Idempotent: an app that is not installed
   * answers `{ outcome: 'already-absent' }` instead of throwing.
   *
   * Consumers must FEATURE-DETECT this method (`typeof
   * provisioner.uninstallFromTeam === 'function'`): the middleware mirrors
   * this contract structurally rather than importing it, so a middleware
   * running against a connector < 0.4.0 has to keep its old
   * not-supported branch.
   */
  uninstallFromTeam(
    input: UninstallFromTeamInput,
  ): Promise<UninstallFromTeamResult>;

  /**
   * Resolve one team id to its display name (since 0.5.0).
   *
   * Every other method here addresses a team by GUID, which is also all the
   * consumer can show an operator. This is the step that turns
   * `19:…@thread.tacv2` into "Marketing". Read-only, and NOT an enumeration:
   * it answers for an id the caller already holds.
   *
   * A team that is gone or not visible answers `{ found: false }` — an
   * ordinary outcome, not a throw, because a consumer's fallback is simply to
   * keep showing the id.
   *
   * FEATURE-DETECT it, for the same reason as {@link uninstallFromTeam}: the
   * middleware mirrors this contract structurally rather than importing it,
   * so a middleware running against a connector < 0.5.0 must keep its
   * id-only branch.
   */
  getTeam(input: GetTeamInput): Promise<GetTeamResult>;
}

/** Everything {@link createTeamsProvisioner} needs — assembled by `activate()`. */
export interface CreateTeamsProvisionerOptions {
  /**
   * The connector's Bot-Framework app credential — always the Graph identity
   * (`microsoft_tenant_id` / `microsoft_app_id` / `microsoft_app_password`).
   */
  readonly graphCredential: AadClientCredential;
  /**
   * The ARM mode decision from `readArmConfig` (config unit). The degraded
   * `'registration-only'` variant flows through unchanged — `createBot`
   * answers it, `canCreateBots` reflects it.
   */
  readonly armConfig: ArmConfigResult;
  /** Plugin vault accessor — where generated app passwords land. */
  readonly secrets: SecretsAccessor;
  /** Which tenant apps are registered in. Default: `'customer'`. */
  readonly tenantMode?: TenantMode;
  readonly log?: (msg: string) => void;
  /** Test seam — identical to the `GraphClient` injection point. */
  readonly fetchImpl?: typeof fetch;
  /** Test seam for backoff/poll waits. */
  readonly sleep?: (ms: number) => Promise<void>;
}

/**
 * Assemble the `teamsProvisioner@1` capability: one shared
 * {@link ProvisioningHttp} (single token cache — the step clients must never
 * open a second one) plus the four step clients and the pure app-package
 * builder. Construction is side-effect free: no token is acquired and no
 * network request is made until a chain step runs.
 */
export function createTeamsProvisioner(
  options: CreateTeamsProvisionerOptions,
): TeamsProvisionerAccessor {
  const tenantMode: TenantMode = options.tenantMode ?? 'customer';
  const { armConfig } = options;

  // "Reuse app" mode maps to http's own fallback (armCredential omitted →
  // graph credential reused for the ARM audience); a dedicated SP becomes a
  // separate credential so the two token-cache entries never collide.
  const armCredential: AadClientCredential | undefined =
    isArmConfigured(armConfig) && armConfig.credential.source === 'dedicated-sp'
      ? {
          tenantId: options.graphCredential.tenantId,
          clientId: armConfig.credential.clientId,
          clientSecret: armConfig.credential.clientSecret,
        }
      : undefined;

  const http = new ProvisioningHttp({
    graphCredential: options.graphCredential,
    armCredential,
    log: options.log,
    fetchImpl: options.fetchImpl,
    sleep: options.sleep,
  });

  const appRegistrations = new AppRegistrationClient({
    http,
    secrets: options.secrets,
    tenantId: options.graphCredential.tenantId,
    log: options.log,
    // Same wait seam the http unit uses — the app-registration step sleeps
    // through Entra's replication windows and tests must not really wait.
    ...(options.sleep !== undefined
      ? { replication: { sleep: options.sleep } }
      : {}),
  });
  const bots = new BotServiceClient({ http, armConfig, log: options.log });

  // Delegated publish plumbing (byte5ai/omadia#924). The publisher app is
  // resolved LAZILY, on the first sign-in — activation stays side-effect free,
  // so an install that never publishes a Teams app never registers one.
  const publisherApps = new PublisherAppClient({
    http,
    tenantId: options.graphCredential.tenantId,
    log: options.log,
    ...(options.sleep !== undefined ? { replication: { sleep: options.sleep } } : {}),
  });
  /**
   * A delegated-auth client bound to whichever publisher app a token set names.
   *
   * Deriving it from the TOKENS rather than from a resolved publisher app is
   * what lets refresh and status work with no Graph call at all — and it keeps
   * a credential minted against an older publisher app renewable instead of
   * silently pointed at a new one.
   */
  const authFor = (tenantId: string, clientId: string): DelegatedAuthClient =>
    new DelegatedAuthClient({
      tenantId,
      clientId,
      log: options.log,
      ...(options.fetchImpl !== undefined ? { fetchImpl: options.fetchImpl } : {}),
      ...(options.sleep !== undefined ? { sleep: options.sleep } : {}),
    });

  const catalog = new CatalogUploadClient({
    http,
    log: options.log,
    delegatedAuth: {
      ensureFreshToken: (tokens) =>
        authFor(tokens.tenantId, tokens.clientId).ensureFreshToken(tokens),
      adminConsentUrlFor: (tokens) =>
        tokens !== undefined
          ? adminConsentUrl(tokens.tenantId, tokens.clientId)
          : `run startDelegatedSignIn to obtain the admin-consent URL for tenant ${options.graphCredential.tenantId}`,
    },
  });
  const installs = new TeamInstallClient({ http, log: options.log });
  const teamLookup = new TeamLookupClient({ http, log: options.log });

  return {
    tenantMode,
    canCreateBots: bots.canCreateBots,
    createAppRegistration: (input) => appRegistrations.createAppRegistration(input),
    deleteAppRegistration: (input) => appRegistrations.deleteAppRegistration(input),
    getAppRegistration: (appId, mode) =>
      appRegistrations.getAppRegistration(appId, mode),
    buildAppPackage: (input) => buildAppPackage(input),
    createBot: (input) => bots.createBot(input),
    deleteBot: (botName) => bots.deleteBot(botName),
    getBot: (botName) => bots.getBot(botName),
    uploadToCatalog: (input) => catalog.uploadToCatalog(input),
    uploadToCatalogDelegated: (input) => catalog.uploadToCatalogDelegated(input),
    getCatalogApp: (input) => catalog.getCatalogApp(input),
    startDelegatedSignIn: async (input) => {
      const publisher = await publisherApps.ensurePublisherApp(input ?? {});
      return authFor(
        publisher.value.tenantId,
        publisher.value.appId,
      ).startDeviceCode();
    },
    pollDelegatedSignIn: (input) =>
      // The handle carries tenant + client id, so the constructor arguments
      // here are placeholders the poll never reads. Passing the connector's
      // own tenant keeps the object well-formed without a Graph lookup.
      authFor(
        options.graphCredential.tenantId,
        options.graphCredential.clientId,
      ).pollDeviceCode(input),
    getDelegatedSignInStatus: (input) => describeSignInStatus(input.tokens),
    refreshDelegatedToken: (input) =>
      authFor(input.tokens.tenantId, input.tokens.clientId).refresh({
        refreshToken: input.tokens.refreshToken,
        tenantId: input.tokens.tenantId,
        clientId: input.tokens.clientId,
      }),
    revokeDelegatedSignIn: (input) => revokeInstructions(input.tokens),
    installToTeam: (input) => installs.installToTeam(input),
    uninstallFromTeam: (input) => installs.uninstallFromTeam(input),
    getTeam: (input) => teamLookup.getTeam(input),
  };
}

// ---------------------------------------------------------------------------
// Public surface of the provisioning subsystem — re-exported by name from
// the package barrel (`src/index.ts`). Curated: implementation classes
// (ProvisioningHttp, the step clients) stay internal.
// ---------------------------------------------------------------------------

export {
  TEAMS_PROVISIONER_SERVICE_NAME,
  TEAMS_PROVISIONER_CAPABILITY,
  SINGLE_TENANT_SIGN_IN_AUDIENCE,
} from './types.js';
export type {
  AppRegistration,
  AppClientSecret,
  AzureBot,
  BotProvisionedOutcome,
  BotProvisioningOutcome,
  CatalogTeamsApp,
  CreateBotInput,
  Idempotent,
  IdempotentOutcome,
  InstallToTeamInput,
  RegisterApplicationInput,
  RegistrationOnlyOutcome,
  SignInAudience,
  TeamAppInstallation,
  TeamsProvisioner,
  TenantMode,
  UploadToCatalogInput,
} from './types.js';

export {
  TeamsProvisionerError,
  ConsentMissingError,
  ProvisioningThrottledError,
  ArmNotConfiguredError,
  CapabilityUnavailableError,
  DirectoryReplicationError,
  ProvisioningRequestError,
  UniqueNameReservedError,
  BotHandleUnavailableError,
  // Delegated catalog-publish taxonomy (byte5ai/omadia#924).
  DelegatedSignInRequiredError,
  DelegatedConsentRequiredError,
  DelegatedTokenExpiredError,
  DeviceCodeFlowError,
  DELETED_ITEM_RETENTION_DAYS,
  isTransientProvisioningFailure,
} from './errors.js';

// Delegated catalog publishing (byte5ai/omadia#924). The DelegatedAuthClient
// and PublisherAppClient classes stay internal — the accessor is the surface.
export {
  APP_CATALOG_DELEGATED_SCOPE,
  APP_CATALOG_DELEGATED_PERMISSION_ID,
  DELEGATED_PUBLISH_SCOPES,
  GRAPH_RESOURCE_APP_ID,
  adminConsentUrl,
  coversCatalogPublish,
  describeSignInStatus,
  isAccessTokenStale,
  revokeInstructions,
} from './delegatedAuth.js';
export type {
  DelegatedAccount,
  DelegatedRevokeResult,
  DelegatedSignInStatus,
  DelegatedSignedInStatus,
  DelegatedSignedOutStatus,
  DelegatedTokenSet,
  DeviceCodeDeclined,
  DeviceCodeExpired,
  DeviceCodeFlowHandle,
  DeviceCodePending,
  DeviceCodePollResult,
  DeviceCodeStart,
  DeviceCodeSucceeded,
} from './delegatedAuth.js';

export {
  PUBLISHER_APP_DISPLAY_NAME,
  PUBLISHER_APP_UNIQUE_NAME_PREFIX,
  publisherAppUniqueName,
} from './publisherApp.js';
export type { PublisherApp } from './publisherApp.js';

export { redactSecrets, redactUnknown, REDACTED } from './redact.js';

export {
  ARM_MANAGEMENT_HOST,
  ARM_TOKEN_SCOPE,
  ARM_CORE_SETUP_FIELD_KEYS,
  ARM_SETUP_FIELD_KEYS,
  AZURE_SUBSCRIPTION_ID_FIELD,
  AZURE_RESOURCE_GROUP_FIELD,
  AZURE_REGION_FIELD,
  AZURE_SP_CLIENT_ID_FIELD,
  AZURE_SP_CLIENT_SECRET_FIELD,
  TEAMS_PROVISIONER_ARM_SETUP_FIELDS,
  isArmConfigured,
  readArmConfig,
} from './config.js';
export type {
  ArmClientCredential,
  ArmConfigResult,
  ArmConfigSource,
  ArmConfigured,
  ReuseAppCredential,
  TeamsProvisionerSetupField,
} from './config.js';

export {
  TEAMS_BOT_PASSWORD_SECRET_PREFIX,
  appIdFromSecretRef,
  secretRefForApp,
} from './secretStore.js';
export type { TeamsBotPasswordSecretRef } from './secretStore.js';

export type {
  CreateAppRegistrationInput,
  DeleteAppRegistrationInput,
  DeleteAppRegistrationOutcome,
  DeleteAppRegistrationResult,
  ProvisionedAppRegistration,
} from './appRegistration.js';

export { AppPackageError } from './appPackage.js';
export type {
  AppPackageIcons,
  AppPackageParamValue,
  AppPackageParams,
  BuildAppPackageInput,
} from './appPackage.js';

export type {
  CatalogAppFound,
  CatalogAppNotFound,
  DelegatedCatalogUploadResult,
  DelegatedUploadAuthority,
  GetCatalogAppInput,
  GetCatalogAppResult,
  UploadToCatalogDelegatedInput,
} from './catalog.js';

export { TEAM_READ_SCOPE } from './teamLookup.js';
export type {
  GetTeamInput,
  GetTeamResult,
  TeamFound,
  TeamNotFound,
} from './teamLookup.js';

export type {
  BotDeletedResult,
  BotFoundResult,
  BotNotFoundResult,
  DeleteBotOutcome,
  DeleteBotResult,
  GetBotResult,
} from './botService.js';

export type {
  ConsentedPermissionSet,
  InstallToTeamRequest,
  ResourceSpecificPermission,
  ResourceSpecificPermissionType,
  UninstallFromTeamInput,
  UninstallFromTeamOutcome,
  UninstallFromTeamResult,
} from './install.js';

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
  type GetCatalogAppInput,
  type GetCatalogAppResult,
} from './catalog.js';
import { isArmConfigured, type ArmConfigResult } from './config.js';
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

  /** Chain step 4 — publish the app package into the tenant catalog. */
  uploadToCatalog(
    input: UploadToCatalogInput,
  ): Promise<Idempotent<CatalogTeamsApp>>;

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
  const catalog = new CatalogUploadClient({ http, log: options.log });
  const installs = new TeamInstallClient({ http, log: options.log });

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
    getCatalogApp: (input) => catalog.getCatalogApp(input),
    installToTeam: (input) => installs.installToTeam(input),
    uninstallFromTeam: (input) => installs.uninstallFromTeam(input),
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
  DELETED_ITEM_RETENTION_DAYS,
  isTransientProvisioningFailure,
} from './errors.js';

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
  GetCatalogAppInput,
  GetCatalogAppResult,
} from './catalog.js';

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

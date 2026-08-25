/**
 * @omadia/integration-microsoft365 — public surface.
 *
 * Ships both the library exports (clients, types, errors) and the
 * plugin-form `activate()` entry point. The plugin registers TWO services:
 *
 *   - `Microsoft365Accessor` under the ServiceRegistry name
 *     `microsoft365.graph` — consumers read it via
 *     `ctx.services.get<Microsoft365Accessor>('microsoft365.graph')`.
 *   - `TeamsProvisionerAccessor` under `teamsProvisioner` (manifest
 *     capability `teamsProvisioner@1`) — per-agent Teams bot provisioning,
 *     consumed by the middleware agent factory (byte5ai/omadia#863-865).
 *
 * ⚠️ Error-name pair on this surface: `ConsentRequiredError` (graphObo.ts,
 * delegated calendar flow, thrown when a USER has not consented) vs
 * `ConsentMissingError` (teamsProvisioner, application-permission flow,
 * thrown on Graph/ARM 403 with the missing scope set). Both are exported —
 * pick by flow, not by name similarity.
 */

export { activate } from './plugin.js';
export type { Microsoft365PluginHandle } from './plugin.js';

export {
  MICROSOFT365_SERVICE_NAME,
  type Microsoft365Accessor,
  type SlotCacheAccessor,
} from './accessor.js';

export { GraphClient, encodeSharingUrl } from './graphClient.js';
export type {
  GraphClientOptions,
  GraphChatMessageAttachment,
} from './graphClient.js';

export {
  GraphOboClient,
  ConsentRequiredError,
  SsoUnavailableError,
  CALENDAR_GRAPH_SCOPES,
  createGraphOboClient,
} from './graphObo.js';
export type { GraphOboConfig } from './graphObo.js';

export { GraphCalendarClient } from './graphCalendarClient.js';
export type {
  AttendeeType,
  FindSlotsOptions,
  MeetingSlotSuggestion,
  GetScheduleOptions,
  ScheduleEntry,
  CreateEventOptions,
  CreatedEvent,
  MailboxSettings,
} from './graphCalendarClient.js';

export { SlotCache } from './slotCache.js';
export type { CachedSlot, SlotCacheOptions } from './slotCache.js';

// ---------------------------------------------------------------------------
// teamsProvisioner@1 — per-agent Teams bot provisioning (wave W0b, epic
// byte5ai/omadia#860). Curated re-export of the subsystem's public surface;
// the step-client classes stay internal to src/teamsProvisioner/.
// ---------------------------------------------------------------------------

export {
  createTeamsProvisioner,
  TEAMS_PROVISIONER_SERVICE_NAME,
  TEAMS_PROVISIONER_CAPABILITY,
  SINGLE_TENANT_SIGN_IN_AUDIENCE,
  // Error taxonomy — note the ConsentRequiredError (graphObo, delegated) /
  // ConsentMissingError (provisioning, application-permission) name pair.
  TeamsProvisionerError,
  ConsentMissingError,
  ProvisioningThrottledError,
  ArmNotConfiguredError,
  CapabilityUnavailableError,
  AppPackageError,
  // ARM config / registration-only degradation (config unit).
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
  // Secret-store surface — only the opaque secretRef crosses the boundary.
  TEAMS_BOT_PASSWORD_SECRET_PREFIX,
  appIdFromSecretRef,
  secretRefForApp,
} from './teamsProvisioner/index.js';
export type {
  TeamsProvisionerAccessor,
  CreateTeamsProvisionerOptions,
  // Shared capability types (types unit).
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
  // ARM config shapes.
  ArmClientCredential,
  ArmConfigResult,
  ArmConfigSource,
  ArmConfigured,
  ReuseAppCredential,
  TeamsProvisionerSetupField,
  // Step inputs/results on the accessor surface.
  CatalogAppFound,
  CatalogAppNotFound,
  GetCatalogAppInput,
  GetCatalogAppResult,
  CreateAppRegistrationInput,
  DeleteAppRegistrationInput,
  DeleteAppRegistrationOutcome,
  DeleteAppRegistrationResult,
  ProvisionedAppRegistration,
  AppPackageIcons,
  AppPackageParamValue,
  AppPackageParams,
  BuildAppPackageInput,
  BotDeletedResult,
  BotFoundResult,
  BotNotFoundResult,
  DeleteBotOutcome,
  DeleteBotResult,
  GetBotResult,
  ConsentedPermissionSet,
  InstallToTeamRequest,
  ResourceSpecificPermission,
  ResourceSpecificPermissionType,
  TeamsBotPasswordSecretRef,
} from './teamsProvisioner/index.js';

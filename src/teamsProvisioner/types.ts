/**
 * `teamsProvisioner@1` — shared types for the Teams provisioning capability.
 *
 * Backs the "1 agent = 1 Entra app + 1 Azure bot + 1 Teams app package" model
 * of the multi-agent Teams identities epic (byte5ai/omadia#860, capability
 * issue byte5ai/omadia-m365-connector#3). Provisioning is token-based REST
 * only (Microsoft Graph + ARM `management.azure.com`) — no `az` CLI.
 *
 * The two-constant split mirrors `transcription.ts` in @omadia/plugin-api:
 * the bare name is the SERVICE-REGISTRY key, the `@1`-suffixed form is the
 * MANIFEST capability ref (`provides` / `requires`). Never use the versioned
 * form as a registry key — `declaredServiceNames()` strips the version before
 * grant classification.
 *
 * ARCHITECTURE INVARIANT — SingleTenant only. New MultiTenant app
 * registrations are deprecated (07/2025), so {@link SignInAudience} models
 * exactly one value: `'AzureADMyOrg'`. The app is registered in the CUSTOMER
 * tenant ({@link TenantMode} `'customer'`); `'home'` exists only to label the
 * legacy single-bot deployment mode where the connector's own credentials
 * live in the operator's home tenant. There is deliberately no way to express
 * a MultiTenant registration in these types.
 *
 * The provisioner exposes the individual chain steps (register app → client
 * secret → Azure bot → catalog upload → team install); the provisioning STATE
 * MACHINE lives middleware-side in the agent factory (byte5ai/omadia#863-865),
 * which consumes this capability. Keep this surface minimal and typed.
 */

/** Service-registry key (bare, unversioned — see module doc). */
export const TEAMS_PROVISIONER_SERVICE_NAME = 'teamsProvisioner';
/** Manifest capability ref (`provides:` / `requires:` form). */
export const TEAMS_PROVISIONER_CAPABILITY = 'teamsProvisioner@1';

/**
 * Which tenant the provisioner operates against. `'customer'` is the normal
 * (and target) mode: the Entra app is registered in the customer's tenant.
 * `'home'` labels the legacy deployment where the connector's credentials
 * belong to the operator's own tenant. MultiTenant is NOT a mode.
 */
export type TenantMode = 'customer' | 'home';

/**
 * Entra `signInAudience` for provisioned app registrations. SingleTenant
 * only — MultiTenant creation is deprecated and deliberately not modeled.
 */
export type SignInAudience = 'AzureADMyOrg';

/** The only sign-in audience the provisioner will ever create. */
export const SINGLE_TENANT_SIGN_IN_AUDIENCE: SignInAudience = 'AzureADMyOrg';

/**
 * Idempotency signal for steps whose remote API answers 409 (or an
 * equivalent duplicate lookup hit) on re-runs — catalog upload keyed by the
 * manifest `externalId`, team install keyed by (teamId, teamsAppId), app
 * registration keyed by `uniqueName`. Callers branch on {@link outcome}
 * instead of string-matching error bodies.
 */
export type IdempotentOutcome = 'created' | 'already-existed';

/** Result wrapper carrying the {@link IdempotentOutcome} alongside the value. */
export interface Idempotent<T> {
  readonly outcome: IdempotentOutcome;
  readonly value: T;
}

/** A provisioned (or found) Entra app registration. */
export interface AppRegistration {
  /** Application (client) id — what Bot Framework calls the MSA app id. */
  readonly appId: string;
  /** Directory object id of the `application` resource (needed for PATCH/DELETE). */
  readonly objectId: string;
  /** Tenant the app is registered in. */
  readonly tenantId: string;
  readonly tenantMode: TenantMode;
  readonly signInAudience: SignInAudience;
  readonly displayName: string;
  /** Stable idempotency key the registration was created/found under, if any. */
  readonly uniqueName?: string;
}

/**
 * A client secret added via `addPassword`. `secretText` is shown exactly once.
 *
 * @deprecated Part of the pre-implementation {@link TeamsProvisioner} sketch.
 * The SHIPPED service (`TeamsProvisionerAccessor`) never returns cleartext
 * secrets across the service boundary — `createAppRegistration` persists the
 * password to the vault and returns only the opaque `secretRef`.
 */
export interface AppClientSecret {
  /** The secret value — persist immediately, Graph never returns it again. */
  readonly secretText: string;
  readonly keyId: string;
  /** ISO-8601 expiry. */
  readonly endDateTime: string;
}

/** An Azure Bot resource created via ARM REST. */
export interface AzureBot {
  /** ARM resource name (bot handle). */
  readonly botName: string;
  /** Full ARM resource id (`/subscriptions/.../botServices/...`). */
  readonly resourceId: string;
  /** The Entra app id the bot authenticates as (SingleTenant). */
  readonly msaAppId: string;
  readonly messagingEndpoint: string;
}

/**
 * Typed degraded outcome for the ARM step: the connector is configured with
 * Graph credentials only (no subscription id / resource group / region / ARM
 * service principal), so the chain can register the Entra app but cannot
 * create the Azure bot. The degradation path (and the agent factory) branch
 * on `kind` instead of catching stringly errors.
 */
export interface RegistrationOnlyOutcome {
  readonly kind: 'registration-only';
  readonly reason: 'arm-not-configured';
  /** Setup-field keys that are missing, e.g. `['azure_subscription_id']`. */
  readonly missingSetupFields: readonly string[];
}

/** Successful ARM bot creation. */
export interface BotProvisionedOutcome {
  readonly kind: 'provisioned';
  readonly bot: Idempotent<AzureBot>;
}

/** Result of the Azure-bot step: full success or registration-only fallback. */
export type BotProvisioningOutcome = BotProvisionedOutcome | RegistrationOnlyOutcome;

/** A Teams app in the tenant app catalog. */
export interface CatalogTeamsApp {
  /** Catalog id (`teamsApp.id`) — what installs reference. */
  readonly teamsAppId: string;
  /** Manifest id (`externalId`) — the idempotency key for uploads. */
  readonly externalId: string;
  readonly displayName: string;
  readonly version: string;
}

/** An app installation into one team (`POST /teams/{id}/installedApps`). */
export interface TeamAppInstallation {
  readonly teamId: string;
  readonly teamsAppId: string;
  /** Graph installation id when the API returned/located one. */
  readonly installationId?: string;
}

/**
 * An app installation into one CHAT (`POST /chats/{id}/installedApps`) —
 * the chat-scope twin of {@link TeamAppInstallation}, since 0.7.0.
 *
 * Kept as its OWN type rather than widening `TeamAppInstallation` with an
 * optional `chatId`: a target is a team or a chat, never both, and a shape
 * where either id may be missing pushes that check into every consumer.
 */
export interface ChatAppInstallation {
  /** Chat thread id (`19:…@thread.v2` group, `19:…@unq.gbl.spaces` 1:1). */
  readonly chatId: string;
  readonly teamsAppId: string;
  /** Graph installation id when the API returned/located one. */
  readonly installationId?: string;
}

/**
 * Input for the Entra app-registration step.
 *
 * @deprecated Part of the pre-implementation {@link TeamsProvisioner} sketch.
 * The shipped `TeamsProvisionerAccessor` merges register-app + add-secret
 * into one rolled-back step — use `CreateAppRegistrationInput`
 * (`src/teamsProvisioner/appRegistration.ts`) instead.
 */
export interface RegisterApplicationInput {
  readonly displayName: string;
  /**
   * Stable idempotency key (Graph `uniqueName`). Re-runs find the existing
   * registration instead of creating a duplicate.
   */
  readonly uniqueName: string;
}

/**
 * Input for the client-secret step.
 *
 * @deprecated Part of the pre-implementation {@link TeamsProvisioner} sketch —
 * see {@link RegisterApplicationInput}.
 */
export interface AddClientSecretInput {
  /** Directory object id of the app registration (NOT the client id). */
  readonly appObjectId: string;
  /** Label shown in the portal, e.g. the agent slug. */
  readonly displayName: string;
}

/** Input for the ARM Azure-bot step. */
export interface CreateBotInput {
  /** ARM resource name / bot handle (also the idempotency key). */
  readonly botName: string;
  readonly displayName: string;
  /** Entra app id the bot authenticates as. */
  readonly msaAppId: string;
  /** Tenant of the SingleTenant app registration. */
  readonly msaAppTenantId: string;
  readonly messagingEndpoint: string;
}

/** Input for the catalog-upload step. */
export interface UploadToCatalogInput {
  /** The zipped Teams app package. */
  readonly packageZip: Uint8Array;
  /** Manifest id — used for the pre-flight `externalId` lookup on 409. */
  readonly externalId: string;
  /**
   * SECRET. Delegated (user) access token for `POST /appCatalogs/teamsApps`.
   *
   * REQUIRED IN PRACTICE, optional in the type (byte5ai/omadia#924): Graph
   * supports this verb for delegated permissions only, so an app-only upload is
   * refused no matter which app roles are consented. The field stays optional
   * so the interface remains signature-compatible for consumers compiled
   * against an older connector — omitting it does not fail to compile, it
   * throws the typed `DelegatedSignInRequiredError` at call time, which is the
   * signal telling the caller to start the device-code sign-in.
   *
   * Prefer `TeamsProvisionerAccessor.uploadToCatalogDelegated`, which takes the
   * whole stored token set and hands back a refreshed one; this field is the
   * low-level seam for a caller that already holds a fresh access token.
   */
  readonly delegatedAccessToken?: string;
}

/** Input for the team-install step. */
export interface InstallToTeamInput {
  readonly teamId: string;
  /** Catalog id (`CatalogTeamsApp.teamsAppId`). */
  readonly teamsAppId: string;
}

/** Input for the chat-install step (since 0.7.0). */
export interface InstallToChatInput {
  /** Chat thread id (`19:…@thread.v2` group, `19:…@unq.gbl.spaces` 1:1). */
  readonly chatId: string;
  /** Catalog id (`CatalogTeamsApp.teamsAppId`). */
  readonly teamsAppId: string;
}

/**
 * The ORIGINAL spec sketch of the `teamsProvisioner@1` service surface.
 *
 * @deprecated NOT what the service registry publishes. The shipped surface is
 * `TeamsProvisionerAccessor` (`src/teamsProvisioner/index.ts`), which
 * deliberately supersedes this sketch: register-app + add-secret are ONE
 * rolled-back step (`createAppRegistration`) and only the vault `secretRef`
 * — never `AppClientSecret.secretText` — crosses the service boundary.
 * Resolve the service as
 * `ctx.services.get<TeamsProvisionerAccessor>('teamsProvisioner')`; coding
 * against THIS interface compiles but fails at runtime (`registerApplication`
 * does not exist on the published object). Kept only as the historical
 * contract the W0b wave was reviewed against.
 */
export interface TeamsProvisioner {
  /** Where this provisioner registers apps. */
  readonly tenantMode: TenantMode;
  /** `true` when the ARM setup fields are configured (bot creation possible). */
  readonly canCreateBots: boolean;

  registerApplication(
    input: RegisterApplicationInput,
  ): Promise<Idempotent<AppRegistration>>;

  addClientSecret(input: AddClientSecretInput): Promise<AppClientSecret>;

  createBot(input: CreateBotInput): Promise<BotProvisioningOutcome>;

  uploadToCatalog(input: UploadToCatalogInput): Promise<Idempotent<CatalogTeamsApp>>;

  installToTeam(input: InstallToTeamInput): Promise<Idempotent<TeamAppInstallation>>;
}

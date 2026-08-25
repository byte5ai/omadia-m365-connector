/**
 * ARM configuration reader for `teamsProvisioner@1` — owns the
 * registration-only-mode decision (graceful degradation, issue
 * byte5ai/omadia-m365-connector#2, epic byte5ai/omadia#860).
 *
 * The Azure-bot step (ARM `PUT .../Microsoft.BotService/botServices/{name}`)
 * needs four things beyond the Graph credentials the plugin already requires:
 * a subscription id, a resource group, a region, and a client credential for
 * `https://management.azure.com/.default`. ALL of them are OPTIONAL setup
 * fields — an install that configures none of them keeps working exactly as
 * before, and the capability degrades to registration-only mode (Entra app +
 * catalog upload possible, `createBot` answers the typed
 * {@link RegistrationOnlyOutcome} instead of provisioning).
 *
 * INVARIANTS
 * - {@link readArmConfig} uses `ctx.config.get` / `ctx.secrets.get` ONLY —
 *   never `require` — so it NEVER throws `MissingConfigError` /
 *   `MissingSecretError`. Missing config is a first-class result
 *   (`kind: 'registration-only'`), not an error. Activation must stay up and
 *   still publish the capability when ARM is unconfigured.
 * - 'reuse app': the dedicated ARM service principal
 *   ({@link AZURE_SP_CLIENT_ID_FIELD} / {@link AZURE_SP_CLIENT_SECRET_FIELD})
 *   may be omitted entirely; the connector's own Bot-Framework app
 *   (`microsoft_app_id` / `microsoft_app_password`) is then reused for ARM —
 *   the caller passes those already-required values in as
 *   {@link ReuseAppCredential}. A HALF-configured dedicated SP (id without
 *   secret, or secret without id) is ambiguous and degrades to
 *   registration-only naming the missing half — it never guesses and never
 *   throws.
 *
 * WIRING-UNIT CONTRACT (executed elsewhere, authored here):
 * - manifest `setup.fields` += {@link TEAMS_PROVISIONER_ARM_SETUP_FIELDS}
 *   verbatim (all `required: false` — `required: true` would break every
 *   existing install; the secret field mirrors `microsoft_app_password`'s
 *   `type: "secret"`).
 * - manifest `permissions.network.outbound` += {@link ARM_MANAGEMENT_HOST}.
 *   The current allowlist has only `login.microsoftonline.com` and
 *   `graph.microsoft.com`; without this entry ARM egress is denied.
 * - `activate()` calls {@link readArmConfig} (never `require`) and publishes
 *   the capability in BOTH result kinds; `canCreateBots` is
 *   {@link isArmConfigured} of the result.
 */

import type { ConfigAccessor, SecretsAccessor } from '@omadia/plugin-api';

import type { RegistrationOnlyOutcome } from './types.js';

/** Host the wiring unit must add to manifest `permissions.network.outbound`. */
export const ARM_MANAGEMENT_HOST = 'management.azure.com';

/** OAuth2 scope the ARM client credential is exchanged for. */
export const ARM_TOKEN_SCOPE = 'https://management.azure.com/.default';

/** Setup-field key: Azure subscription the bot resources live in. */
export const AZURE_SUBSCRIPTION_ID_FIELD = 'azure_subscription_id';
/** Setup-field key: resource group for `Microsoft.BotService/botServices`. */
export const AZURE_RESOURCE_GROUP_FIELD = 'azure_resource_group';
/** Setup-field key: ARM location of the bot resources (usually `global`). */
export const AZURE_REGION_FIELD = 'azure_region';
/** Setup-field key: client id of the dedicated ARM service principal. */
export const AZURE_SP_CLIENT_ID_FIELD = 'azure_sp_client_id';
/** Setup-field key: client secret of the dedicated ARM service principal. */
export const AZURE_SP_CLIENT_SECRET_FIELD = 'azure_sp_client_secret';

/**
 * The fields that MUST be present for ARM mode. The service-principal pair is
 * deliberately not in here — its absence means 'reuse app', not degradation.
 */
export const ARM_CORE_SETUP_FIELD_KEYS = [
  AZURE_SUBSCRIPTION_ID_FIELD,
  AZURE_RESOURCE_GROUP_FIELD,
  AZURE_REGION_FIELD,
] as const;

/** Every ARM setup-field key this unit introduces, in manifest order. */
export const ARM_SETUP_FIELD_KEYS = [
  ...ARM_CORE_SETUP_FIELD_KEYS,
  AZURE_SP_CLIENT_ID_FIELD,
  AZURE_SP_CLIENT_SECRET_FIELD,
] as const;

/** GUID pattern, identical to the existing `microsoft_tenant_id` field. */
const GUID_PATTERN = '^[0-9a-fA-F-]{36}$';

/**
 * One manifest `setup.fields` entry, mirroring the shape already used by
 * `microsoft_tenant_id` / `microsoft_app_password` in `manifest.yaml`
 * (key / type / label / help / required / pattern).
 */
export interface TeamsProvisionerSetupField {
  readonly key: string;
  readonly type: 'string' | 'secret';
  readonly label: string;
  readonly help: string;
  /** ALWAYS `false` — `required: true` would break existing installs. */
  readonly required: false;
  readonly pattern?: string;
}

/**
 * The manifest field spec the wiring unit copies into `setup.fields`.
 * German help texts follow the existing manifest's convention.
 */
export const TEAMS_PROVISIONER_ARM_SETUP_FIELDS: readonly TeamsProvisionerSetupField[] =
  [
    {
      key: AZURE_SUBSCRIPTION_ID_FIELD,
      type: 'string',
      label: 'Azure Subscription ID',
      help:
        'GUID der Azure-Subscription, in der die Azure-Bot-Ressourcen angelegt werden. ' +
        'Leer lassen für Registration-only-Modus (kein automatisches Bot-Provisioning).',
      required: false,
      pattern: GUID_PATTERN,
    },
    {
      key: AZURE_RESOURCE_GROUP_FIELD,
      type: 'string',
      label: 'Azure Resource Group',
      help:
        'Name der Resource Group für Microsoft.BotService/botServices. ' +
        'Leer lassen für Registration-only-Modus.',
      required: false,
      pattern: '^[-\\w._()]{1,90}$',
    },
    {
      key: AZURE_REGION_FIELD,
      type: 'string',
      label: 'Azure Region',
      help:
        "ARM-Location der Bot-Ressourcen, für Azure Bot Services üblicherweise 'global'. " +
        'Leer lassen für Registration-only-Modus.',
      required: false,
      pattern: '^[a-z0-9]{2,32}$',
    },
    {
      key: AZURE_SP_CLIENT_ID_FIELD,
      type: 'string',
      label: 'ARM Service Principal Client ID',
      help:
        'Client-ID eines dedizierten Service Principals für management.azure.com. ' +
        'Leer lassen, um die Bot-Framework-App (App (Client) ID + Secret) für ARM wiederzuverwenden.',
      required: false,
      pattern: GUID_PATTERN,
    },
    {
      key: AZURE_SP_CLIENT_SECRET_FIELD,
      type: 'secret',
      label: 'ARM Service Principal Client Secret',
      help:
        'Client-Secret des dedizierten ARM Service Principals. ' +
        'Leer lassen, um das App (Client) Secret der Bot-Framework-App wiederzuverwenden.',
      required: false,
    },
  ];

/**
 * The already-required Bot-Framework app credentials (`microsoft_app_id` /
 * `microsoft_app_password`), passed in by `activate()` which has read them via
 * `require` anyway. Used for ARM when no dedicated SP is configured.
 */
export interface ReuseAppCredential {
  readonly clientId: string;
  readonly clientSecret: string;
}

/** The client credential the ARM token request will use. */
export interface ArmClientCredential {
  readonly clientId: string;
  readonly clientSecret: string;
  /** Whether a dedicated SP is configured or the app credential is reused. */
  readonly source: 'dedicated-sp' | 'reused-app';
}

/** Full ARM configuration — bot creation is possible. */
export interface ArmConfigured {
  readonly kind: 'configured';
  readonly subscriptionId: string;
  readonly resourceGroup: string;
  readonly region: string;
  readonly credential: ArmClientCredential;
}

/**
 * Result of reading the ARM setup fields. The degraded variant IS the shared
 * {@link RegistrationOnlyOutcome} from `types.ts`, so `createBot` can return
 * it unchanged and observers see one shape everywhere.
 */
export type ArmConfigResult = ArmConfigured | RegistrationOnlyOutcome;

/** `true` when bot creation is possible — feeds `TeamsProvisioner.canCreateBots`. */
export function isArmConfigured(
  result: ArmConfigResult,
): result is ArmConfigured {
  return result.kind === 'configured';
}

/**
 * The slice of `PluginContext` the reader needs. Only the OPTIONAL accessors
 * — the `require` variants are structurally unreachable from here.
 */
export interface ArmConfigSource {
  readonly config: Pick<ConfigAccessor, 'get'>;
  readonly secrets: Pick<SecretsAccessor, 'get'>;
}

/**
 * Reads a config value defensively: only a non-blank string counts as
 * configured. Anything else (absent, empty, whitespace-only, wrong type from
 * a hand-edited registry) is treated as "not set" — boundary validation
 * instead of a downstream crash.
 */
function readOptionalString(
  config: Pick<ConfigAccessor, 'get'>,
  key: string,
): string | undefined {
  const raw = config.get<unknown>(key);
  return normalize(raw);
}

/** Blank-safe trim: returns `undefined` for non-strings and empty strings. */
function normalize(raw: unknown): string | undefined {
  if (typeof raw !== 'string') {
    return undefined;
  }
  const trimmed = raw.trim();
  return trimmed === '' ? undefined : trimmed;
}

/**
 * Read the ARM setup fields and decide the mode.
 *
 * - All three core fields present → `kind: 'configured'`, with either the
 *   dedicated SP (both halves set) or the reused app credential (both absent).
 * - Any core field missing, or a half-configured SP pair → the typed
 *   `kind: 'registration-only'` outcome listing exactly the missing keys.
 *
 * Never throws for missing configuration (uses `get`, never `require`);
 * transport failures from the secrets vault propagate verbatim.
 */
export async function readArmConfig(
  ctx: ArmConfigSource,
  reuseApp: ReuseAppCredential,
): Promise<ArmConfigResult> {
  const subscriptionId = readOptionalString(ctx.config, AZURE_SUBSCRIPTION_ID_FIELD);
  const resourceGroup = readOptionalString(ctx.config, AZURE_RESOURCE_GROUP_FIELD);
  const region = readOptionalString(ctx.config, AZURE_REGION_FIELD);
  const spClientId = readOptionalString(ctx.config, AZURE_SP_CLIENT_ID_FIELD);
  const spClientSecret = normalize(
    await ctx.secrets.get(AZURE_SP_CLIENT_SECRET_FIELD),
  );

  const missingSetupFields: string[] = [];
  if (subscriptionId === undefined) {
    missingSetupFields.push(AZURE_SUBSCRIPTION_ID_FIELD);
  }
  if (resourceGroup === undefined) {
    missingSetupFields.push(AZURE_RESOURCE_GROUP_FIELD);
  }
  if (region === undefined) {
    missingSetupFields.push(AZURE_REGION_FIELD);
  }
  // A half-configured dedicated SP is ambiguous: never guess which credential
  // was meant — degrade and name the absent half.
  if (spClientId !== undefined && spClientSecret === undefined) {
    missingSetupFields.push(AZURE_SP_CLIENT_SECRET_FIELD);
  }
  if (spClientId === undefined && spClientSecret !== undefined) {
    missingSetupFields.push(AZURE_SP_CLIENT_ID_FIELD);
  }

  // The explicit `undefined` re-checks are redundant with the pushes above but
  // let the compiler narrow the core fields to `string` below — no casts.
  if (
    missingSetupFields.length > 0 ||
    subscriptionId === undefined ||
    resourceGroup === undefined ||
    region === undefined
  ) {
    return {
      kind: 'registration-only',
      reason: 'arm-not-configured',
      missingSetupFields,
    };
  }

  const credential: ArmClientCredential =
    spClientId !== undefined && spClientSecret !== undefined
      ? { clientId: spClientId, clientSecret: spClientSecret, source: 'dedicated-sp' }
      : {
          clientId: reuseApp.clientId,
          clientSecret: reuseApp.clientSecret,
          source: 'reused-app',
        };

  return { kind: 'configured', subscriptionId, resourceGroup, region, credential };
}

/**
 * Entra app-registration step of `teamsProvisioner@1` (epic byte5ai/omadia#860,
 * capability issue byte5ai/omadia-m365-connector#3): create — and roll back —
 * the "1 agent = 1 Entra app" identity.
 *
 * `createAppRegistration` chains THREE Graph writes:
 *
 *   1. `POST /applications`                    — the app registration itself
 *   2. `POST /applications/{id}/addPassword`   — the bot client secret
 *   3. `POST /servicePrincipals`               — the tenant service principal
 *
 * A partial failure rolls back what this call created (best-effort, logged,
 * original error rethrown); `deleteAppRegistration` is the rollback half and
 * is idempotent — an already-deleted app answers with the idempotent
 * `'already-deleted'` signal (mirroring `Idempotent<T>` in `types.ts` for the
 * delete direction), never an error.
 *
 * ARCHITECTURE INVARIANT — SingleTenant only. Every registration is created
 * with `signInAudience: 'AzureADMyOrg'` in the CUSTOMER tenant; MultiTenant is
 * deliberately not expressible (see `types.ts` module doc). An existing app
 * found via the `uniqueName` idempotency lookup is rejected if its audience
 * is anything else.
 *
 * SECRET HANDLING INVARIANT — the `addPassword` secret value never crosses
 * this module's API boundary. It is written straight into the plugin vault via
 * the secret-store unit (`storeAppPassword`, keyed `teams_bot_password:<appId>`)
 * and only the opaque `secretRef` is returned. No log line and no error
 * message ever contains the secret.
 *
 * All HTTP goes through the shared {@link ProvisioningHttp} choke point (one
 * token cache, Retry-After-honouring 429 backoff, 403 → `ConsentMissingError`
 * with {@link APP_REGISTRATION_SCOPE}, 409 → conflict signal). This module
 * opens no second token cache and does no fetch of its own.
 *
 * Style precedent: `graphClient.ts` — options-bag constructor, hand-rolled
 * REST against `https://graph.microsoft.com/v1.0`, non-2xx → typed/plain
 * errors with step + status.
 */

import type { SecretsAccessor } from '@omadia/plugin-api';

import type { ProvisioningHttp, ProvisioningOkResponse } from './http.js';
import {
  storeAppPassword,
  deleteAppPassword,
  secretRefForApp,
  type TeamsBotPasswordSecretRef,
} from './secretStore.js';
import {
  SINGLE_TENANT_SIGN_IN_AUDIENCE,
  type AppRegistration,
  type Idempotent,
  type IdempotentOutcome,
  type TenantMode,
} from './types.js';

/**
 * Graph APPLICATION permission every call in this module needs. Documented in
 * the scopes/consent unit and granted by the wiring unit's manifest bump —
 * surfaced on 403 via `ConsentMissingError.missingScopes`.
 */
export const APP_REGISTRATION_SCOPE = 'Application.ReadWrite.OwnedBy';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export interface AppRegistrationClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  /** Plugin vault accessor — where the generated app password lands. */
  readonly secrets: SecretsAccessor;
  /** Tenant the registrations are created in (stamped into results). */
  readonly tenantId: string;
  readonly log?: (msg: string) => void;
}

export interface CreateAppRegistrationInput {
  readonly displayName: string;
  /**
   * Which tenant this registration targets — `'customer'` is the normal mode.
   * Label only: the audience is ALWAYS SingleTenant (`'AzureADMyOrg'`).
   */
  readonly tenantMode: TenantMode;
  /**
   * Stable idempotency key (Graph `uniqueName`). Strongly recommended: on a
   * re-run Graph answers 409 and the existing registration is found and
   * reused (with a freshly rotated secret) instead of creating a duplicate.
   */
  readonly uniqueName?: string;
  /** Portal label for the generated secret. Default: `<displayName> bot password`. */
  readonly secretDisplayName?: string;
}

/** What `createAppRegistration` hands back — note: NO secret value, only the ref. */
export interface ProvisionedAppRegistration {
  /** Application (client) id — what Bot Framework calls the MSA app id. */
  readonly appId: string;
  /** Opaque vault reference to the generated password — never the cleartext. */
  readonly secretRef: TeamsBotPasswordSecretRef;
  /** The full registration resource (objectId, audience, tenant, …). */
  readonly registration: AppRegistration;
  /** `keyId` of the added password credential (needed for `removePassword`). */
  readonly secretKeyId: string;
  /** ISO-8601 expiry of the added password credential. */
  readonly secretEndDateTime: string;
  /** Whether the service principal was created now or already existed. */
  readonly servicePrincipalOutcome: IdempotentOutcome;
}

/**
 * Idempotent delete signal — the delete-direction mirror of
 * `IdempotentOutcome` in `types.ts` (whose members are creation-specific).
 */
export type DeleteAppRegistrationOutcome = 'deleted' | 'already-deleted';

export interface DeleteAppRegistrationInput {
  /** Entra application (client) id of the registration to delete. */
  readonly appId: string;
  /**
   * When provided, the stored bot password (`teams_bot_password:<appId>`) is
   * removed from the vault after the Graph delete succeeded.
   */
  readonly secretRef?: string;
}

export interface DeleteAppRegistrationResult {
  readonly outcome: DeleteAppRegistrationOutcome;
}

/**
 * Create/find, secret-rotate and roll back Entra app registrations. One step
 * client per provisioner; ordering across chain steps stays middleware-side
 * (agent factory, byte5ai/omadia#863-865).
 */
export class AppRegistrationClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly secrets: SecretsAccessor;
  private readonly tenantId: string;
  private readonly log: (msg: string) => void;

  constructor(opts: AppRegistrationClientOptions) {
    this.http = opts.http;
    this.secrets = opts.secrets;
    this.tenantId = requireNonEmpty(opts.tenantId, 'tenantId');
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * The full app-registration step: register (or find via `uniqueName`) the
   * SingleTenant app, add a client secret (persisted to the vault, only the
   * `secretRef` is returned), and ensure the service principal exists.
   *
   * The `Idempotent` outcome reflects the REGISTRATION: `'already-existed'`
   * means the app was found via `uniqueName` — the secret is still rotated
   * (deterministic vault key, intended overwrite) and the service principal
   * still ensured.
   *
   * On a partial failure everything this call created is rolled back
   * (app registration and/or vault entry and/or password credential),
   * best-effort with logging, and the ORIGINAL error is rethrown.
   */
  async createAppRegistration(
    input: CreateAppRegistrationInput,
  ): Promise<Idempotent<ProvisionedAppRegistration>> {
    requireNonEmpty(input.displayName, 'displayName');
    if (input.uniqueName !== undefined) {
      requireUniqueName(input.uniqueName);
    }
    const { registration, outcome } = await this.registerApplication(input);

    // Captured BEFORE the vault write: on the uniqueName-reuse path
    // storeAppPassword OVERWRITES a prior run's entry, and a rollback must
    // restore that entry instead of destroying a credential this call did
    // not create. The value stays local — never logged, never returned.
    const priorSecretValue = await this.secrets.get(
      secretRefForApp(registration.appId),
    );

    let secretRef: TeamsBotPasswordSecretRef | undefined;
    let secretKeyId: string | undefined;
    try {
      const password = await this.addPassword(registration, input);
      secretKeyId = password.keyId;
      secretRef = await storeAppPassword(
        this.secrets,
        registration.appId,
        password.secretText,
      );
      const servicePrincipalOutcome = await this.ensureServicePrincipal(
        registration.appId,
      );
      return {
        outcome,
        value: {
          appId: registration.appId,
          secretRef,
          registration,
          secretKeyId: password.keyId,
          secretEndDateTime: password.endDateTime,
          servicePrincipalOutcome,
        },
      };
    } catch (err) {
      await this.rollbackPartialCreate(
        registration,
        outcome,
        secretRef,
        secretKeyId,
        priorSecretValue,
      );
      throw err;
    }
  }

  /**
   * Rollback half: delete the app registration (Entra cascades the delete to
   * its service principal) and, when a `secretRef` is passed, the vault entry.
   * Idempotent — a 404 answers `'already-deleted'`, never an error, so job
   * retries and double rollbacks are safe.
   */
  async deleteAppRegistration(
    input: DeleteAppRegistrationInput,
  ): Promise<DeleteAppRegistrationResult> {
    const appId = requireAppId(input.appId);
    const response = await this.http.request({
      resource: 'graph',
      method: 'DELETE',
      url: `${GRAPH_BASE}/applications(appId='${appId}')`,
      step: 'applications.delete',
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
      // Already gone = already rolled back — the idempotent delete signal.
      extraOkStatuses: [404],
    });
    if (response.kind === 'conflict') {
      throw new Error('graph applications.delete unexpected 409');
    }
    if (input.secretRef !== undefined) {
      await deleteAppPassword(this.secrets, input.secretRef);
    }
    return {
      outcome: response.status === 404 ? 'already-deleted' : 'deleted',
    };
  }

  /**
   * Status probe: the registration as Graph sees it, or `undefined` when the
   * app does not (or no longer does) exist.
   */
  async getAppRegistration(
    appId: string,
    tenantMode: TenantMode,
  ): Promise<AppRegistration | undefined> {
    const validAppId = requireAppId(appId);
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/applications(appId='${validAppId}')`,
      step: 'applications.get',
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
      extraOkStatuses: [404],
    });
    if (response.kind === 'conflict') {
      throw new Error('graph applications.get unexpected 409');
    }
    if (response.status === 404) return undefined;
    return this.parseApplication(response, 'applications.get', tenantMode);
  }

  /** `POST /applications`; 409 (uniqueName taken) → find + reuse the existing app. */
  private async registerApplication(
    input: CreateAppRegistrationInput,
  ): Promise<{ registration: AppRegistration; outcome: IdempotentOutcome }> {
    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/applications`,
      step: 'applications.create',
      jsonBody: {
        displayName: input.displayName,
        // THE invariant: SingleTenant, always. Never a MultiTenant audience.
        signInAudience: SINGLE_TENANT_SIGN_IN_AUDIENCE,
        ...(input.uniqueName !== undefined
          ? { uniqueName: input.uniqueName }
          : {}),
      },
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });

    if (response.kind === 'ok') {
      return {
        registration: this.parseApplication(
          response,
          'applications.create',
          input.tenantMode,
        ),
        outcome: 'created',
      };
    }

    // 409: the uniqueName idempotency key is already taken — find and reuse.
    if (input.uniqueName === undefined) {
      throw new Error(
        'graph applications.create 409 without a uniqueName — cannot resolve ' +
          'the conflicting registration; pass a stable uniqueName for idempotent re-runs',
      );
    }
    const existing = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/applications(uniqueName='${encodeURIComponent(escapeODataQuotes(input.uniqueName))}')`,
      step: 'applications.findByUniqueName',
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });
    if (existing.kind !== 'ok') {
      throw new Error('graph applications.findByUniqueName unexpected 409');
    }
    return {
      registration: this.parseApplication(
        existing,
        'applications.findByUniqueName',
        input.tenantMode,
      ),
      outcome: 'already-existed',
    };
  }

  /** `POST /applications/{id}/addPassword` — the secret value stays local to the caller. */
  private async addPassword(
    registration: AppRegistration,
    input: CreateAppRegistrationInput,
  ): Promise<{ secretText: string; keyId: string; endDateTime: string }> {
    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/applications/${encodeURIComponent(registration.objectId)}/addPassword`,
      step: 'applications.addPassword',
      jsonBody: {
        passwordCredential: {
          displayName:
            input.secretDisplayName ?? `${input.displayName} bot password`,
        },
      },
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });
    if (response.kind === 'conflict') {
      throw new Error('graph applications.addPassword unexpected 409');
    }
    const body = asRecord(response.json, 'applications.addPassword');
    return {
      secretText: requireStringField(body, 'secretText', 'applications.addPassword'),
      keyId: requireStringField(body, 'keyId', 'applications.addPassword'),
      endDateTime: requireStringField(body, 'endDateTime', 'applications.addPassword'),
    };
  }

  /** `POST /servicePrincipals`; 409 = it already exists — the idempotent signal. */
  private async ensureServicePrincipal(
    appId: string,
  ): Promise<IdempotentOutcome> {
    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/servicePrincipals`,
      step: 'servicePrincipals.create',
      jsonBody: { appId },
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });
    return response.kind === 'conflict' ? 'already-existed' : 'created';
  }

  /**
   * Best-effort rollback after a partial create. Only artifacts of THIS call
   * are touched: a registration found via `uniqueName` (`'already-existed'`)
   * is never deleted — its freshly added password credential is removed
   * instead, and a vault entry the call OVERWROTE (rather than created) is
   * restored to its previous value instead of deleted. Rollback failures are
   * logged (never a secret value), the caller rethrows the original error.
   */
  private async rollbackPartialCreate(
    registration: AppRegistration,
    outcome: IdempotentOutcome,
    secretRef: TeamsBotPasswordSecretRef | undefined,
    secretKeyId: string | undefined,
    priorSecretValue: string | undefined,
  ): Promise<void> {
    if (secretRef !== undefined) {
      try {
        if (priorSecretValue !== undefined) {
          // The vault key existed before this call — storeAppPassword
          // overwrote it. Deleting would destroy a credential this call did
          // not create; restore the pre-call value instead.
          await storeAppPassword(
            this.secrets,
            registration.appId,
            priorSecretValue,
          );
        } else {
          await deleteAppPassword(this.secrets, secretRef);
        }
      } catch (err) {
        this.log(
          `provisioner rollback: vault ${priorSecretValue !== undefined ? 'restore' : 'delete'} of '${secretRef}' failed: ${String(err)}`,
        );
      }
    }
    if (outcome === 'created') {
      try {
        await this.deleteAppRegistration({ appId: registration.appId });
      } catch (err) {
        this.log(
          `provisioner rollback: applications.delete of '${registration.appId}' failed: ${String(err)}`,
        );
      }
      return;
    }
    // Pre-existing app: remove only the password credential this call added.
    if (secretKeyId !== undefined) {
      try {
        await this.http.request({
          resource: 'graph',
          method: 'POST',
          url: `${GRAPH_BASE}/applications/${encodeURIComponent(registration.objectId)}/removePassword`,
          step: 'applications.removePassword',
          jsonBody: { keyId: secretKeyId },
          missingScopesOn403: [APP_REGISTRATION_SCOPE],
          extraOkStatuses: [404],
        });
      } catch (err) {
        this.log(
          `provisioner rollback: applications.removePassword on '${registration.appId}' failed: ${String(err)}`,
        );
      }
    }
  }

  /** Parse + validate a Graph `application` resource; enforces the audience invariant. */
  private parseApplication(
    response: ProvisioningOkResponse,
    step: string,
    tenantMode: TenantMode,
  ): AppRegistration {
    const body = asRecord(response.json, step);
    const signInAudience = requireStringField(body, 'signInAudience', step);
    if (signInAudience !== SINGLE_TENANT_SIGN_IN_AUDIENCE) {
      // Invariant guard: a uniqueName-found app with any other audience is a
      // foreign/legacy registration this provisioner must not adopt.
      throw new Error(
        `graph ${step}: refusing non-SingleTenant registration ` +
          `(signInAudience='${signInAudience}', expected '${SINGLE_TENANT_SIGN_IN_AUDIENCE}')`,
      );
    }
    const uniqueName = body['uniqueName'];
    return {
      appId: requireStringField(body, 'appId', step),
      objectId: requireStringField(body, 'id', step),
      tenantId: this.tenantId,
      tenantMode,
      signInAudience: SINGLE_TENANT_SIGN_IN_AUDIENCE,
      displayName: requireStringField(body, 'displayName', step),
      ...(typeof uniqueName === 'string' ? { uniqueName } : {}),
    };
  }
}

function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

/** Entra client ids are GUID-shaped — reject anything that could break the OData URL. */
function requireAppId(appId: string): string {
  requireNonEmpty(appId, 'appId');
  if (/[\s'()/]/.test(appId)) {
    throw new Error(
      "invalid_argument: 'appId' must be a plain Entra application (client) id",
    );
  }
  return appId;
}

/**
 * `uniqueName` is an idempotency key that ends up in an OData alternate-key
 * URL path. Mirroring {@link requireAppId}, anything that could break the
 * URL or the OData literal — whitespace, quotes, parens, path separators,
 * `#`/`?`/`%`/`&`/`+` — is rejected up front (belt) even though the lookup
 * additionally percent-encodes the value (braces).
 */
function requireUniqueName(uniqueName: string): string {
  requireNonEmpty(uniqueName, 'uniqueName');
  if (/[\s'"()/\\#?%&+]/.test(uniqueName)) {
    throw new Error(
      "invalid_argument: 'uniqueName' must not contain whitespace, quotes, " +
        'parentheses, slashes or URL metacharacters (#, ?, %, &, +)',
    );
  }
  return uniqueName;
}

/** OData string-literal escaping for alternate-key lookups (`'` → `''`). */
function escapeODataQuotes(value: string): string {
  return value.replace(/'/g, "''");
}

function asRecord(json: unknown, step: string): Record<string, unknown> {
  if (json === null || typeof json !== 'object') {
    throw new Error(`graph ${step}: unexpected empty/non-object response body`);
  }
  return json as Record<string, unknown>;
}

function requireStringField(
  body: Record<string, unknown>,
  field: string,
  step: string,
): string {
  const value = body[field];
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`graph ${step}: response body missing '${field}'`);
  }
  return value;
}

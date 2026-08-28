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
 * Entra is EVENTUALLY CONSISTENT between (1) and (2): the create answers 201
 * and `addPassword` on that brand-new object can still answer 404
 * `Request_ResourceNotFound` for a few seconds. That is a replication window,
 * not a failure — the step polls through it (`directory.ts`) instead of
 * giving up (byte5ai/omadia#916).
 *
 * ROLLBACK IS NARROW, on purpose. A partial failure used to delete the app it
 * had just created; because a deleted Entra app is only SOFT-deleted, that
 * reserved its `uniqueName` for 30 days and every retry then collided with an
 * object nobody could see — one transient 404 burned a provisioning slug for a
 * month. Now: a transient failure rolls back NOTHING (the registration is the
 * durable, adoptable artifact), and a registration carrying a `uniqueName` is
 * never deleted by the rollback path at all — it is addressable by its natural
 * key, so the next run adopts it. `deleteAppRegistration` remains the explicit,
 * idempotent delete for deprovisioning — an already-deleted app answers with
 * the `'already-deleted'` signal (mirroring `Idempotent<T>` in `types.ts` for
 * the delete direction), never an error.
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
 * opens no second token cache and does no fetch of its own. Note that the
 * `uniqueName` duplicate does NOT arrive as a 409: Entra reports it as a 400
 * `Request_BadRequest`, mapped onto the conflict signal by
 * `UNIQUE_NAME_CONFLICT_RULES` in `directory.ts`.
 *
 * Style precedent: `graphClient.ts` — options-bag constructor, hand-rolled
 * REST against `https://graph.microsoft.com/v1.0`, non-2xx → typed/plain
 * errors with step + status.
 */

import type { SecretsAccessor } from '@omadia/plugin-api';

import {
  asRecord,
  escapeODataQuotes,
  requireAppId,
  requireNonEmpty,
  requireStringField,
  requireUniqueName,
  secretLabel,
  staleCredentialKeyIds,
} from './appRegistrationSupport.js';

import {
  GRAPH_BASE,
  UNIQUE_NAME_CONFLICT_RULES,
  findDeletedApplicationByUniqueName,
  restoreDeletedApplication,
  retryWhileReplicating,
  waitForApplicationAddressable,
  type ReplicationOptions,
} from './directory.js';
import {
  DELETED_ITEM_RETENTION_DAYS,
  UniqueNameReservedError,
  isTransientProvisioningFailure,
} from './errors.js';
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
  /**
   * Budget for the Entra replication windows (probe count, base interval)
   * and the test seam for the waits. Defaults cover roughly 40 s across 8
   * probes — long enough for the observed windows, far short of the job
   * runner's own retry cadence.
   */
  readonly replication?: ReplicationOptions;
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
  /**
   * Called the MOMENT the registration exists — after create/adopt, BEFORE
   * the secret and the service principal. The caller persists `appId` here,
   * so an interruption anywhere in the rest of the chain leaves a RESUMABLE
   * row instead of an orphaned app nobody knows about (byte5ai/omadia#916).
   *
   * Best effort by design: a throwing callback is logged and the step
   * continues. The registration carries a `uniqueName`, so it stays
   * adoptable even when the persistence attempt failed — turning a store
   * hiccup into a Graph-side abort would trade a small problem for a bigger
   * one.
   */
  readonly onRegistrationCreated?: (
    registration: AppRegistration,
    outcome: IdempotentOutcome,
  ) => void | Promise<void>;
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
  private readonly replication: ReplicationOptions;

  constructor(opts: AppRegistrationClientOptions) {
    this.http = opts.http;
    this.secrets = opts.secrets;
    this.tenantId = requireNonEmpty(opts.tenantId, 'tenantId');
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
    this.replication = { ...opts.replication, log: (msg) => { this.log(msg); } };
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

    // The app_id is now the one durable fact of this step — hand it to the
    // caller BEFORE anything else can fail, so an interruption leaves a
    // resumable row rather than an orphan (byte5ai/omadia#916).
    await this.notifyRegistrationCreated(input, registration, outcome);

    if (outcome === 'created') {
      // Eventual consistency: the object exists but may not be addressable
      // for the follow-up writes yet. Poll, do not fail.
      await waitForApplicationAddressable(
        this.http,
        registration.objectId,
        [APP_REGISTRATION_SCOPE],
        this.replication,
      );
    }

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
      if (outcome === 'already-existed') {
        await this.pruneStaleBotPasswords(registration, input, password.keyId);
      }
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
      if (isTransientProvisioningFailure(err)) {
        // NOT a reason to undo anything. The registration is real and
        // adoptable by its uniqueName; deleting it here is precisely what
        // reserved a slug for 30 days in byte5ai/omadia#916.
        this.log(
          `provisioner app-registration: transient failure after ${outcome === 'created' ? 'creating' : 'adopting'} ` +
            `'${registration.appId}' — leaving it in place for an idempotent re-run (${String(err)})`,
        );
        throw err;
      }
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

  /** Best-effort early persistence hook — see `onRegistrationCreated`. */
  private async notifyRegistrationCreated(
    input: CreateAppRegistrationInput,
    registration: AppRegistration,
    outcome: IdempotentOutcome,
  ): Promise<void> {
    if (input.onRegistrationCreated === undefined) return;
    try {
      await input.onRegistrationCreated(registration, outcome);
    } catch (err) {
      this.log(
        `provisioner app-registration: onRegistrationCreated for '${registration.appId}' ` +
          `failed: ${String(err)}`,
      );
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

  /** `POST /applications`; a taken `uniqueName` → find + adopt the existing app. */
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
      // Entra reports a taken uniqueName as 400, not 409 — without these
      // rules the adopt branch below can never be reached.
      conflictOn: UNIQUE_NAME_CONFLICT_RULES,
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

    // The uniqueName idempotency key is already taken — find and adopt.
    if (input.uniqueName === undefined) {
      throw new Error(
        'graph applications.create conflict without a uniqueName — cannot resolve ' +
          'the conflicting registration; pass a stable uniqueName for idempotent re-runs',
      );
    }
    return {
      registration: await this.adoptByUniqueName(input, input.uniqueName),
      outcome: 'already-existed',
    };
  }

  /**
   * Resolve the registration behind a taken `uniqueName`.
   *
   * Two possible owners of the name:
   *
   * - a LIVE application — the normal idempotent re-run: look it up by the
   *   alternate key and adopt it (fresh secret, service principal ensured).
   * - a SOFT-DELETED application — invisible in `GET /applications` yet still
   *   holding the name for {@link DELETED_ITEM_RETENTION_DAYS} days. Restoring
   *   it is the only way to free the name early, and it returns the SAME
   *   object (same appId, same uniqueName) — exactly what re-provisioning the
   *   agent slug is supposed to yield. Restricted to an object whose
   *   `uniqueName` matches ours exactly, i.e. one this provisioner owns.
   *
   * If neither resolves, the operator finally learns WHY the name is taken
   * instead of reading "already exists" about an object they cannot see.
   */
  private async adoptByUniqueName(
    input: CreateAppRegistrationInput,
    uniqueName: string,
  ): Promise<AppRegistration> {
    const live = await this.findByUniqueName(uniqueName, input.tenantMode);
    if (live !== undefined) return live;

    const deleted = await findDeletedApplicationByUniqueName(
      this.http,
      uniqueName,
      [APP_REGISTRATION_SCOPE],
      this.log,
    );
    if (deleted === undefined) {
      throw new UniqueNameReservedError(uniqueName, {
        hint:
          `'${uniqueName}' is already taken but no matching application is visible. ` +
          `A deleted Entra app keeps its uniqueName reserved for ${String(DELETED_ITEM_RETENTION_DAYS)} days ` +
          `while sitting in the recycle bin, and the recycle bin could not be read ` +
          `(it needs Application.ReadWrite.All, not just ${APP_REGISTRATION_SCOPE}). ` +
          `Check 'directory/deletedItems/microsoft.graph.application' in the tenant, ` +
          `then restore that object or provision under a different name.`,
      });
    }

    this.log(
      `provisioner applications.create: '${uniqueName}' is held by soft-deleted app ` +
        `'${deleted.objectId}'${deleted.deletedDateTime !== undefined ? ` (deleted ${deleted.deletedDateTime})` : ''} — restoring it`,
    );
    try {
      await restoreDeletedApplication(this.http, deleted.objectId, [
        APP_REGISTRATION_SCOPE,
      ]);
    } catch (err) {
      throw new UniqueNameReservedError(
        uniqueName,
        {
          deletedObjectId: deleted.objectId,
          ...(deleted.deletedDateTime !== undefined
            ? { deletedDateTime: deleted.deletedDateTime }
            : {}),
          hint:
            `'${uniqueName}' is reserved by soft-deleted application '${deleted.objectId}'` +
            `${deleted.deletedDateTime !== undefined ? ` (deleted ${deleted.deletedDateTime})` : ''}. ` +
            `Entra holds the name for ${String(DELETED_ITEM_RETENTION_DAYS)} days after a delete and restoring ` +
            `it failed (${String(err)}). Restore or purge it manually — note a purge does NOT ` +
            `release the name — or provision under a different name.`,
        },
        err,
      );
    }

    // The restore is itself replicated — the app is not readable instantly.
    const restored = await retryWhileReplicating(
      'applications.findByUniqueName',
      deleted.objectId,
      () => this.findByUniqueName(uniqueName, input.tenantMode),
      this.replication,
    );
    return restored;
  }

  /** `GET /applications(uniqueName='…')` — `undefined` when no live app holds it. */
  private async findByUniqueName(
    uniqueName: string,
    tenantMode: TenantMode,
  ): Promise<AppRegistration | undefined> {
    const existing = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/applications(uniqueName='${encodeURIComponent(escapeODataQuotes(uniqueName))}')`,
      step: 'applications.findByUniqueName',
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
      extraOkStatuses: [404],
    });
    if (existing.kind !== 'ok') {
      throw new Error('graph applications.findByUniqueName unexpected conflict');
    }
    if (existing.status === 404) return undefined;
    return this.parseApplication(
      existing,
      'applications.findByUniqueName',
      tenantMode,
    );
  }

  /**
   * `POST /applications/{id}/addPassword` — the secret value stays local to
   * the caller.
   *
   * A 404 here does NOT mean the app is gone: Graph replicates read and write
   * paths independently, so a brand-new registration can still answer
   * `Request_ResourceNotFound` on this write after it already reads back
   * fine. Retried through that window; an exhausted budget raises the
   * transient `DirectoryReplicationError` (byte5ai/omadia#916).
   */
  private addPassword(
    registration: AppRegistration,
    input: CreateAppRegistrationInput,
  ): Promise<{ secretText: string; keyId: string; endDateTime: string }> {
    return retryWhileReplicating(
      'applications.addPassword',
      registration.objectId,
      () => this.tryAddPassword(registration, input),
      this.replication,
    );
  }

  /** One `addPassword` attempt; `undefined` = "not replicated yet, retry". */
  private async tryAddPassword(
    registration: AppRegistration,
    input: CreateAppRegistrationInput,
  ): Promise<{ secretText: string; keyId: string; endDateTime: string } | undefined> {
    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/applications/${encodeURIComponent(registration.objectId)}/addPassword`,
      step: 'applications.addPassword',
      jsonBody: {
        passwordCredential: {
          displayName: secretLabel(input),
        },
      },
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
      extraOkStatuses: [404],
    });
    if (response.kind === 'conflict') {
      throw new Error('graph applications.addPassword unexpected 409');
    }
    if (response.status === 404) return undefined;
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
   * Best-effort rollback after a NON-transient partial create (transient
   * failures never get here — see `createAppRegistration`).
   *
   * Only artifacts of THIS call are touched, and the registration itself is
   * deleted in exactly one case: it was created here AND carries no
   * `uniqueName`, i.e. nothing could ever find it again. A registration WITH
   * a uniqueName is kept: deleting it soft-deletes the object and reserves
   * that name for ${DELETED_ITEM_RETENTION_DAYS} days, which is how a single
   * failed run used to make a slug unusable for a month (byte5ai/omadia#916).
   * Keeping it costs one adoptable app; deleting it costs the name.
   *
   * A registration found via `uniqueName` (`'already-existed'`) additionally
   * gets its freshly added password credential removed, and a vault entry the
   * call OVERWROTE (rather than created) is restored to its previous value
   * instead of deleted. Rollback failures are logged (never a secret value),
   * the caller rethrows the original error.
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
      if (registration.uniqueName !== undefined) {
        this.log(
          `provisioner rollback: keeping registration '${registration.appId}' ` +
            `(uniqueName='${registration.uniqueName}') — deleting it would reserve that name for ` +
            `${String(DELETED_ITEM_RETENTION_DAYS)} days; the next run adopts it instead`,
        );
        return;
      }
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

  /**
   * Adopting an existing registration always mints a FRESH secret — the
   * original was returned exactly once and never persisted, so it cannot be
   * recovered. That makes the credentials this provisioner created pile up on
   * every re-run, and Entra caps how many an app may hold.
   *
   * So: remove the provisioner's OWN older credentials (matched by the
   * deterministic secret label, minus the one just added) and leave everything
   * else alone — a credential an operator added by hand is not ours to revoke.
   * Best effort: the registration is already usable, a failed cleanup must not
   * fail the step.
   */
  private async pruneStaleBotPasswords(
    registration: AppRegistration,
    input: CreateAppRegistrationInput,
    keepKeyId: string,
  ): Promise<void> {
    const label = secretLabel(input);
    try {
      const response = await this.http.request({
        resource: 'graph',
        method: 'GET',
        url: `${GRAPH_BASE}/applications/${encodeURIComponent(registration.objectId)}?$select=passwordCredentials`,
        step: 'applications.readCredentials',
        missingScopesOn403: [APP_REGISTRATION_SCOPE],
      });
      if (response.kind !== 'ok') return;
      const stale = staleCredentialKeyIds(response.json, label, keepKeyId);
      for (const keyId of stale) {
        await this.http.request({
          resource: 'graph',
          method: 'POST',
          url: `${GRAPH_BASE}/applications/${encodeURIComponent(registration.objectId)}/removePassword`,
          step: 'applications.removePassword',
          jsonBody: { keyId },
          missingScopesOn403: [APP_REGISTRATION_SCOPE],
          extraOkStatuses: [404],
        });
      }
      if (stale.length > 0) {
        this.log(
          `provisioner app-registration: removed ${String(stale.length)} superseded ` +
            `'${label}' credential(s) from '${registration.appId}'`,
        );
      }
    } catch (err) {
      this.log(
        `provisioner app-registration: pruning superseded credentials on ` +
          `'${registration.appId}' failed: ${String(err)}`,
      );
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

// Pure value helpers live in `appRegistrationSupport.ts` — see its module doc.

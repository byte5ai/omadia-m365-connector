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
  lookUpDeletedApplication,
  purgeDeletedApplication,
  resolveReplicationOptions,
  retryWhileReplicating,
  waitForApplicationAddressable,
  type DeletedApplication,
  type DeletedApplicationLookup,
  type ReplicationOptions,
} from './directory.js';
import {
  DELETED_ITEM_RETENTION_DAYS,
  DeletedObjectIdMismatchError,
  RecycleBinUnreadableError,
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
  /** Clock for "deleted moments ago". Tests inject one. */
  readonly now?: () => number;
}

/**
 * How long a delete made through this client counts as "moments ago": within
 * it, an application that is neither live nor in the recycle bin is treated as
 * not replicated yet, never as gone (byte5ai/omadia#2043, rollout defect D10).
 */
const RECENT_DELETE_WINDOW_MS = 10 * 60_000;
/** Looks into the recycle bin for a tombstone this client expects there. */
const PURGE_CONFIRM_ATTEMPTS = 6;
/** Looks for a tombstone nobody announced, before the bin's answer is believed. */
const PURGE_RECHECK_ATTEMPTS = 2;

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
 * What a purge established.
 *
 * - `'purged'`: Entra removed the recycle-bin entry now.
 * - `'already-absent'`: the recycle bin was READ and does not hold the entry,
 *   and this client did not delete the application moments ago.
 * - `'unconfirmed'` (since 0.9.2): neither of the above could be established.
 *   The application was deleted moments ago and its entry has not appeared in
 *   the recycle bin yet, or the bin could not be read. The `uniqueName` and
 *   the client id may still be held: ask again, do not provision on top.
 */
export type PurgeDeletedAppRegistrationOutcome =
  | 'purged'
  | 'already-absent'
  | 'unconfirmed';

export interface PurgeDeletedAppRegistrationInput {
  /**
   * DIRECTORY OBJECT id of the soft-deleted application — `application.id`,
   * NOT the `appId` that {@link DeleteAppRegistrationInput} takes.
   *
   * The two are both GUIDs and cannot be told apart by shape, which is why
   * this field is named for the one that works and why the implementation
   * re-probes the recycle bin before it reports `'already-absent'`.
   * `ProvisionedAppRegistration.registration.objectId` is where a caller
   * normally has it from.
   */
  readonly objectId: string;
}

export interface PurgeDeletedAppRegistrationResult {
  readonly outcome: PurgeDeletedAppRegistrationOutcome;
  /** Why the outcome is `'unconfirmed'`. */
  readonly detail?: string;
}

export interface FindDeletedAppRegistrationInput {
  /** Entra application (client) id to look for in the recycle bin. */
  readonly appId: string;
}

/** A recycle bin that was READ, and what it holds for the application. */
export type FindDeletedAppRegistrationResult =
  | {
      readonly found: true;
      /** Directory object id — what {@link PurgeDeletedAppRegistrationInput} takes. */
      readonly objectId: string;
      readonly deletedDateTime?: string;
    }
  | { readonly found: false };

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
  private readonly now: () => number;
  /**
   * Applications deleted through this client, by lower-cased app id AND
   * object id → time of the delete. In memory on purpose: it only has to
   * outlive Entra's replication of that delete, which is seconds.
   */
  private readonly recentDeletes = new Map<string, number>();

  constructor(opts: AppRegistrationClientOptions) {
    this.now = opts.now ?? Date.now;
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
   * means a LIVE app was found via `uniqueName` — the secret is still rotated
   * (deterministic vault key, intended overwrite) and the service principal
   * still ensured. A soft-deleted app holding the name is purged and replaced
   * by a new registration (`'created'`, new client id), never restored.
   *
   * Safe to re-run, and interruption-friendly: `onRegistrationCreated` fires
   * the moment the registration exists so the caller can persist `appId`
   * before anything else can fail.
   *
   * On a partial failure the ORIGINAL error is always rethrown, and what gets
   * undone is deliberately narrow (byte5ai/omadia#916): a TRANSIENT failure
   * undoes nothing, and a registration carrying a `uniqueName` is never
   * deleted — see {@link rollbackPartialCreate}.
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
    // The purge that follows a delete addresses the OBJECT id, so learn it
    // while the application can still be read: it is what lets that purge
    // know "this was deleted moments ago" (see `recentDeletes`).
    const objectId = await this.objectIdOfLiveApplication(appId);
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
    // Live a moment ago and deleted now, by this call or by someone racing it.
    if (response.status !== 404 || objectId !== undefined) {
      this.rememberDelete(appId, objectId);
    }
    if (input.secretRef !== undefined) {
      await deleteAppPassword(this.secrets, input.secretRef);
    }
    return {
      outcome: response.status === 404 ? 'already-deleted' : 'deleted',
    };
  }

  /**
   * The other half of a real reset: PERMANENTLY delete the soft-deleted app
   * so its `uniqueName` stops being reserved.
   *
   * {@link deleteAppRegistration} only soft-deletes — Entra parks the object
   * in the recycle bin and holds its name for
   * {@link DELETED_ITEM_RETENTION_DAYS} days. That is why a "clean up and
   * retry" that stops at the delete makes the next attempt WORSE: it collides
   * with an object nobody can see (byte5ai/omadia#916). Call this when the
   * intent is to free the slug, not to keep the app recoverable — a purge is
   * irreversible, and the appId dies with the object.
   *
   * Idempotent: an entry the bin provably does not hold answers
   * `'already-absent'`. Since 0.9.2 that needs a bin that was read; right
   * after a delete made through this client the bin is asked up to
   * {@link PURGE_CONFIRM_ATTEMPTS} times for the tombstone, which is purged
   * when it appears, and `'unconfirmed'` is the answer when it does not.
   *
   * WITH ONE GUARD. Passing an `appId` here would 404 exactly like an entry
   * that was genuinely purged, so a bare 404 is not evidence of anything. This
   * re-scans the bin for that id and throws {@link DeletedObjectIdMismatchError}
   * — carrying the object id that WOULD have worked — rather than reporting a
   * freed name that is still reserved. The extra Graph call is spent only on
   * the 404 path, and only to avoid lying.
   */
  async purgeDeletedAppRegistration(
    input: PurgeDeletedAppRegistrationInput,
  ): Promise<PurgeDeletedAppRegistrationResult> {
    const objectId = requireNonEmpty(input.objectId, 'objectId');
    const purge = (): Promise<'purged' | 'already-absent'> =>
      purgeDeletedApplication(this.http, objectId, [APP_REGISTRATION_SCOPE]);
    if ((await purge()) === 'purged') return this.purged(objectId);

    // 404. One of: purged before, the wrong GUID, or a delete whose tombstone
    // has not replicated into the recycle bin yet. Only the bin can tell, and
    // only when it can be read; a bare 404 is evidence of nothing. That last
    // case is the one 0.9.1 reported as 'already-absent' milliseconds after a
    // delete, and the next run revived the application (D10).
    const deletedMomentsAgo = this.wasDeletedRecently(objectId);
    const { intervalMs, sleep } = resolveReplicationOptions(this.replication);
    const attempts = deletedMomentsAgo ? PURGE_CONFIRM_ATTEMPTS : PURGE_RECHECK_ATTEMPTS;
    let lookup: DeletedApplicationLookup = { readable: false };

    for (let attempt = 0; attempt < attempts; attempt += 1) {
      if (attempt > 0) await sleep(Math.min(intervalMs * 2 ** (attempt - 1), 8000));
      lookup = await this.lookUpDeleted(
        (candidate) =>
          candidate.objectId.toLowerCase() === objectId.toLowerCase() ||
          candidate.appId?.toLowerCase() === objectId.toLowerCase(),
        objectId,
      );
      const tombstone = lookup.found;
      if (tombstone === undefined) continue;
      if (tombstone.objectId.toLowerCase() !== objectId.toLowerCase()) {
        throw new DeletedObjectIdMismatchError(objectId, tombstone.objectId);
      }
      if ((await purge()) === 'purged') return this.purged(objectId);
      // Listed but not purgeable yet: the two paths replicate separately.
    }

    if (lookup.found !== undefined) {
      return this.unconfirmed(
        objectId,
        'the recycle bin lists the application, but the permanent delete answered 404',
      );
    }
    if (!lookup.readable) {
      return this.unconfirmed(
        objectId,
        'the recycle bin could not be read (it needs Application.ReadWrite.All), ' +
          'so a 404 on the permanent delete proves nothing',
      );
    }
    if (deletedMomentsAgo) {
      return this.unconfirmed(
        objectId,
        `the application was deleted moments ago and its entry did not appear in the ` +
          `recycle bin within ${String(attempts)} looks`,
      );
    }
    return { outcome: 'already-absent' };
  }

  /**
   * Is the application in the recycle bin (since 0.9.2)? Answers only for a
   * bin that was READ: when it cannot be, this throws
   * {@link RecycleBinUnreadableError} instead of answering `found: false`,
   * because "not found" would read as "the name and the client id are free".
   */
  async findDeletedAppRegistration(
    input: FindDeletedAppRegistrationInput,
  ): Promise<FindDeletedAppRegistrationResult> {
    const appId = requireAppId(input.appId).toLowerCase();
    const lookup = await this.lookUpDeleted(
      (candidate) => candidate.appId?.toLowerCase() === appId,
      appId,
    );
    if (lookup.found !== undefined) {
      return {
        found: true,
        objectId: lookup.found.objectId,
        ...(lookup.found.deletedDateTime !== undefined
          ? { deletedDateTime: lookup.found.deletedDateTime }
          : {}),
      };
    }
    if (!lookup.readable) throw new RecycleBinUnreadableError(appId);
    return { found: false };
  }

  private lookUpDeleted(
    predicate: (candidate: DeletedApplication) => boolean,
    label: string,
  ): Promise<DeletedApplicationLookup> {
    return lookUpDeletedApplication(
      this.http,
      predicate,
      label,
      [APP_REGISTRATION_SCOPE],
      this.log,
    );
  }

  private purged(objectId: string): PurgeDeletedAppRegistrationResult {
    this.recentDeletes.delete(objectId.toLowerCase());
    return { outcome: 'purged' };
  }

  private unconfirmed(
    objectId: string,
    detail: string,
  ): PurgeDeletedAppRegistrationResult {
    this.log(
      `provisioner directory.deletedItems.delete: purge of '${objectId}' unconfirmed — ${detail}`,
    );
    return { outcome: 'unconfirmed', detail };
  }

  /** `id` of the live application behind `appId`, when it can be read. */
  private async objectIdOfLiveApplication(
    appId: string,
  ): Promise<string | undefined> {
    try {
      const response = await this.http.request({
        resource: 'graph',
        method: 'GET',
        url: `${GRAPH_BASE}/applications(appId='${appId}')?$select=id`,
        step: 'applications.get',
        missingScopesOn403: [APP_REGISTRATION_SCOPE],
        extraOkStatuses: [404],
      });
      if (response.kind !== 'ok' || response.status === 404) return undefined;
      const id = (response.json as Record<string, unknown> | null)?.['id'];
      return typeof id === 'string' && id !== '' ? id : undefined;
    } catch (err) {
      // The delete must not depend on this read.
      this.log(
        `provisioner applications.delete: could not read the object id of '${appId}' ` +
          `before deleting it (${String(err)})`,
      );
      return undefined;
    }
  }

  private rememberDelete(appId: string, objectId: string | undefined): void {
    const at = this.now();
    this.recentDeletes.set(appId.toLowerCase(), at);
    if (objectId !== undefined) this.recentDeletes.set(objectId.toLowerCase(), at);
  }

  /** Was this app id or object id deleted through this client moments ago? */
  private wasDeletedRecently(id: string): boolean {
    const at = this.recentDeletes.get(id.toLowerCase());
    return at !== undefined && this.now() - at < RECENT_DELETE_WINDOW_MS;
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

  /** `POST /applications`; a taken `uniqueName` → {@link resolveTakenUniqueName}. */
  private async registerApplication(
    input: CreateAppRegistrationInput,
  ): Promise<{ registration: AppRegistration; outcome: IdempotentOutcome }> {
    const created = await this.postApplication(input);
    if (created !== undefined) return { registration: created, outcome: 'created' };

    // The uniqueName idempotency key is already taken.
    if (input.uniqueName === undefined) {
      throw new Error(
        'graph applications.create conflict without a uniqueName — cannot resolve ' +
          'the conflicting registration; pass a stable uniqueName for idempotent re-runs',
      );
    }
    return this.resolveTakenUniqueName(input, input.uniqueName);
  }

  /** One `POST /applications`; `undefined` = the `uniqueName` is taken. */
  private async postApplication(
    input: CreateAppRegistrationInput,
  ): Promise<AppRegistration | undefined> {
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
      // rules the taken-name branch can never be reached.
      conflictOn: UNIQUE_NAME_CONFLICT_RULES,
    });
    if (response.kind !== 'ok') return undefined;
    return this.parseApplication(response, 'applications.create', input.tenantMode);
  }

  /**
   * Resolve a taken `uniqueName`. Two possible owners of the name:
   *
   * - a LIVE application — the normal idempotent re-run: adopt it (fresh
   *   secret, service principal ensured). Not when this client deleted that
   *   application moments ago and the delete has only not replicated yet:
   *   that one is on its way into the recycle bin and is waited for.
   * - a SOFT-DELETED application — invisible in `GET /applications`, holding
   *   the name for {@link DELETED_ITEM_RETENTION_DAYS} days. It is PURGED and
   *   a new application is created, with a new client id.
   *
   * NEVER RESTORED (since 0.9.2). Until 0.9.1 the tombstone was restored,
   * which brought back the same client id. A bot resource re-created under an
   * app id that had one before is rejected by the Bot Connector for good
   * (401), while Azure reports it healthy (byte5ai/omadia#2043, rollout
   * defect D10). A deleted identity is over; its name goes to a new one.
   * Restricted to a tombstone whose `uniqueName` matches ours exactly, i.e.
   * one this provisioner owns.
   *
   * The purge and the freed name both replicate, so this polls on the
   * replication budget. If the name is still not free at the end, the
   * operator learns WHY instead of reading "already exists" about an object
   * they cannot see.
   */
  private async resolveTakenUniqueName(
    input: CreateAppRegistrationInput,
    uniqueName: string,
  ): Promise<{ registration: AppRegistration; outcome: IdempotentOutcome }> {
    const { maxAttempts, intervalMs, sleep } = resolveReplicationOptions(this.replication);
    const retiredAppIds = new Set<string>();
    let lastTombstone: DeletedApplication | undefined;

    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
      if (attempt > 0) await sleep(Math.min(intervalMs * 2 ** (attempt - 1), 8000));

      const live = await this.findByUniqueName(uniqueName, input.tenantMode);
      const liveIsBeingDeleted =
        live !== undefined &&
        (this.wasDeletedRecently(live.appId) || retiredAppIds.has(live.appId.toLowerCase()));
      if (live !== undefined && !liveIsBeingDeleted) {
        return { registration: live, outcome: 'already-existed' };
      }

      const lookup = await this.lookUpDeleted(
        (candidate) => candidate.uniqueName === uniqueName,
        uniqueName,
      );
      if (lookup.found !== undefined) {
        lastTombstone = lookup.found;
        if (lookup.found.appId !== undefined) {
          retiredAppIds.add(lookup.found.appId.toLowerCase());
        }
        await this.purgeTombstone(lookup.found, uniqueName);
      } else if (!lookup.readable && lastTombstone === undefined && live === undefined) {
        throw this.nameReservedUnseen(uniqueName);
      }
      if (live !== undefined) continue;

      const created = await this.postApplication(input);
      if (created === undefined) continue;
      if (retiredAppIds.has(created.appId.toLowerCase())) {
        throw new Error(
          `graph applications.create: the new registration for '${uniqueName}' came back with ` +
            `client id '${created.appId}' of the application that was just purged — refusing to use it`,
        );
      }
      return { registration: created, outcome: 'created' };
    }

    throw new UniqueNameReservedError(uniqueName, {
      ...(lastTombstone !== undefined ? { deletedObjectId: lastTombstone.objectId } : {}),
      ...(lastTombstone?.deletedDateTime !== undefined
        ? { deletedDateTime: lastTombstone.deletedDateTime }
        : {}),
      hint:
        `'${uniqueName}' is still taken after ${String(maxAttempts)} attempts` +
        (lastTombstone !== undefined
          ? `: soft-deleted application '${lastTombstone.objectId}' held it and was purged, but ` +
            `Entra has not released the name yet`
          : `, and no live application and no recycle-bin entry with that name is visible`) +
        `. Run provisioning again in a minute, or provision under a different name.`,
    });
  }

  /** Permanently delete the tombstone that holds `uniqueName`. */
  private async purgeTombstone(
    tombstone: DeletedApplication,
    uniqueName: string,
  ): Promise<void> {
    const deletedAt =
      tombstone.deletedDateTime !== undefined ? ` (deleted ${tombstone.deletedDateTime})` : '';
    this.log(
      `provisioner applications.create: '${uniqueName}' is held by soft-deleted app ` +
        `'${tombstone.objectId}'${deletedAt} — purging it; the new registration gets a new client id`,
    );
    try {
      await purgeDeletedApplication(this.http, tombstone.objectId, [APP_REGISTRATION_SCOPE]);
    } catch (err) {
      throw new UniqueNameReservedError(
        uniqueName,
        {
          deletedObjectId: tombstone.objectId,
          ...(tombstone.deletedDateTime !== undefined
            ? { deletedDateTime: tombstone.deletedDateTime }
            : {}),
          hint:
            `'${uniqueName}' is reserved by soft-deleted application '${tombstone.objectId}'${deletedAt}. ` +
            `Entra holds the name for ${String(DELETED_ITEM_RETENTION_DAYS)} days after a delete, and ` +
            `permanently deleting it failed (${String(err)}). Delete it permanently by hand ` +
            `(DELETE /directory/deletedItems/${tombstone.objectId}), which frees the name ` +
            `immediately, or provision under a different name. Do not restore it: a bot ` +
            `re-created under its old client id is rejected by Teams.`,
        },
        err,
      );
    }
  }

  private nameReservedUnseen(uniqueName: string): UniqueNameReservedError {
    return new UniqueNameReservedError(uniqueName, {
      hint:
        `'${uniqueName}' is already taken but no matching application is visible. ` +
        `A deleted Entra app keeps its uniqueName reserved for ${String(DELETED_ITEM_RETENTION_DAYS)} days ` +
        `while sitting in the recycle bin, and the recycle bin could not be read ` +
        `(it needs Application.ReadWrite.All, not just ${APP_REGISTRATION_SCOPE}). ` +
        `Check 'directory/deletedItems/microsoft.graph.application' in the tenant and ` +
        `permanently delete that object to free the name, or provision under a different ` +
        `name. Do not restore it: a bot re-created under its old client id is rejected by Teams.`,
    });
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

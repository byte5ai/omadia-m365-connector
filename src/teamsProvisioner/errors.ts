/**
 * Typed error taxonomy for `teamsProvisioner@1`.
 *
 * Follows the `graphObo.ts` precedent verbatim: named `Error` subclasses with
 * explicit `this.name` and structured `public readonly` fields, snake_case
 * message codes. Only failures a CALLER must branch on get a class here —
 * everything else (transport errors, unexpected AAD/ARM responses) propagates
 * verbatim. The 409 already-exists paths are deliberately NOT errors; they
 * surface as the `Idempotent<T>` outcome in `types.ts`.
 */

/**
 * Base class for all provisioner errors so consumers (the agent factory,
 * byte5ai/omadia#863-865) can catch the whole taxonomy with one
 * `instanceof` check.
 */
export abstract class TeamsProvisionerError extends Error {
  protected constructor(message: string, cause?: unknown) {
    super(message);
    if (cause !== undefined) {
      (this as { cause?: unknown }).cause = cause;
    }
  }
}

/**
 * Thrown when Graph or ARM answers 403 because an application permission /
 * role assignment has not been granted (or admin consent is missing). Carries
 * the missing scope set so the middleware agent factory can fall back (e.g.
 * to the deep-link consent card) and render an actionable message.
 *
 * NAME COLLISION, on purpose distinct: `graphObo.ts` exports
 * `ConsentRequiredError` for the DELEGATED calendar flow (user consent via
 * OAuthCard). This class covers the APPLICATION-permission provisioning flow
 * (tenant-admin consent). Both end up on the public surface — keep the names
 * apart when importing.
 */
export class ConsentMissingError extends TeamsProvisionerError {
  /** The scopes/app roles the caller must have granted, e.g. `Application.ReadWrite.OwnedBy`. */
  public readonly missingScopes: readonly string[];
  /** Which API rejected the call. */
  public readonly resource: 'graph' | 'arm';

  constructor(
    missingScopes: readonly string[],
    resource: 'graph' | 'arm' = 'graph',
    cause?: unknown,
  ) {
    super('consent_missing', cause);
    this.name = 'ConsentMissingError';
    this.missingScopes = missingScopes;
    this.resource = resource;
  }
}

/**
 * Thrown after the 429 retry/backoff budget is exhausted (Graph and ARM both
 * throttle). Carries the last `Retry-After` hint, when the API sent one, so
 * the job runner (byte5ai/omadia#864) can schedule a later attempt instead
 * of parsing headers out of a stringified response.
 */
export class ProvisioningThrottledError extends TeamsProvisionerError {
  /** Seconds from the final `Retry-After` header, if the API provided it. */
  public readonly retryAfterSeconds?: number;
  /** Which API throttled the call. */
  public readonly resource: 'graph' | 'arm';

  constructor(
    resource: 'graph' | 'arm',
    retryAfterSeconds?: number,
    cause?: unknown,
  ) {
    super('provisioning_throttled', cause);
    this.name = 'ProvisioningThrottledError';
    this.resource = resource;
    if (retryAfterSeconds !== undefined) {
      this.retryAfterSeconds = retryAfterSeconds;
    }
  }
}

/**
 * Thrown when an ARM-dependent step is invoked although the ARM setup fields
 * (subscription id, resource group, region, service-principal credential)
 * are not configured. The graceful-degradation path should never see this —
 * it branches on the `RegistrationOnlyOutcome` result in `types.ts` first;
 * this error exists for callers that DEMAND full provisioning.
 */
export class ArmNotConfiguredError extends TeamsProvisionerError {
  /** Setup-field keys that are missing, e.g. `['azure_subscription_id']`. */
  public readonly missingSetupFields: readonly string[];

  constructor(missingSetupFields: readonly string[]) {
    super('arm_not_configured');
    this.name = 'ArmNotConfiguredError';
    this.missingSetupFields = missingSetupFields;
  }
}

/**
 * Thrown when a runtime secret write/delete is attempted but the kernel did
 * not hand out `secrets.set` / `secrets.delete` — i.e. the manifest does not
 * declare `permissions.secrets.runtime_write` (or the kernel predates
 * Spec 004). Carries the manifest permission to declare so the operator/dev
 * message is actionable. Fourth member of the taxonomy (added by the
 * secret-store unit; the spec's original three classes cover the network
 * paths, this one covers the vault capability gate).
 */
export class CapabilityUnavailableError extends TeamsProvisionerError {
  /** The manifest permission whose absence caused this. */
  public readonly missingPermission = 'permissions.secrets.runtime_write';
  /** Which vault operation was attempted. */
  public readonly operation: 'set' | 'delete';

  constructor(operation: 'set' | 'delete') {
    super('capability_unavailable');
    this.name = 'CapabilityUnavailableError';
    this.operation = operation;
  }
}

/**
 * A Graph/ARM call answered with a status the choke point treats as an error.
 *
 * Carries the STRUCTURE (`resource`, `step`, `status`) that used to be
 * recoverable only by parsing the message, so callers can classify a failure
 * — notably {@link isTransientProvisioningFailure} — instead of string
 * matching. The message text is unchanged from the plain `Error` this
 * replaced (byte5ai/omadia#916), so existing log lines and assertions keep
 * reading the same.
 */
export class ProvisioningRequestError extends TeamsProvisionerError {
  /** Which API answered. */
  public readonly resource: 'graph' | 'arm';
  /** Step label, e.g. `applications.addPassword`. */
  public readonly step: string;
  /** The HTTP status that was not accepted. */
  public readonly status: number;

  constructor(
    resource: 'graph' | 'arm',
    step: string,
    status: number,
    message: string,
    cause?: unknown,
  ) {
    super(message, cause);
    this.name = 'ProvisioningRequestError';
    this.resource = resource;
    this.step = step;
    this.status = status;
  }
}

/**
 * A directory object that Graph confirmed as created was still not
 * addressable for follow-up writes when the replication budget ran out.
 *
 * Entra is eventually consistent: `POST /applications` answers 201 and the
 * very next `POST /applications/{id}/addPassword` can answer 404
 * `Request_ResourceNotFound` for a few seconds (byte5ai/omadia#916). The
 * app-registration step polls through that window; this error means the
 * window was longer than the budget. It is TRANSIENT — the object exists,
 * a later run finds it under its `uniqueName`. Never a reason to roll back.
 */
export class DirectoryReplicationError extends TeamsProvisionerError {
  /** Step label that kept answering 404, e.g. `applications.addPassword`. */
  public readonly step: string;
  /** Directory object id that was not addressable yet. */
  public readonly objectId: string;
  /** How many probes were spent. */
  public readonly attempts: number;
  /** Total time waited across those probes (ms). */
  public readonly waitedMs: number;

  constructor(
    step: string,
    objectId: string,
    attempts: number,
    waitedMs: number,
  ) {
    super('directory_replication_pending');
    this.name = 'DirectoryReplicationError';
    this.step = step;
    this.objectId = objectId;
    this.attempts = attempts;
    this.waitedMs = waitedMs;
  }
}

/** How long Entra keeps a deleted application in the recycle bin. */
export const DELETED_ITEM_RETENTION_DAYS = 30;

/**
 * The `uniqueName` idempotency key is taken by an object the provisioner
 * cannot adopt — in practice a SOFT-DELETED application: Entra reserves the
 * name for {@link DELETED_ITEM_RETENTION_DAYS} days after a delete, while the
 * object is invisible in every normal listing.
 *
 * Exists so the operator no longer sees a bare "already exists" for an object
 * they cannot find anywhere: the message says WHY the name is unavailable,
 * for how long, and which object to purge.
 */
export class UniqueNameReservedError extends TeamsProvisionerError {
  /** The idempotency key that is unavailable. */
  public readonly uniqueName: string;
  /** Recycle-bin object id, when the deleted-items probe could resolve one. */
  public readonly deletedObjectId?: string;
  /** ISO-8601 deletion timestamp, when Graph reported one. */
  public readonly deletedDateTime?: string;
  /** Days the name stays reserved after the delete. */
  public readonly retentionDays = DELETED_ITEM_RETENTION_DAYS;

  constructor(
    uniqueName: string,
    detail: {
      readonly deletedObjectId?: string;
      readonly deletedDateTime?: string;
      readonly hint: string;
    },
    cause?: unknown,
  ) {
    super(`unique_name_reserved: ${detail.hint}`, cause);
    this.name = 'UniqueNameReservedError';
    this.uniqueName = uniqueName;
    if (detail.deletedObjectId !== undefined) {
      this.deletedObjectId = detail.deletedObjectId;
    }
    if (detail.deletedDateTime !== undefined) {
      this.deletedDateTime = detail.deletedDateTime;
    }
  }
}

/**
 * The Azure bot handle is not available: it is registered to a bot
 * application that is not ours (byte5ai/omadia#921).
 *
 * Bot Service handles live in a SINGLE GLOBAL namespace shared by every Azure
 * customer — they behave like DNS labels, not like tenant- or
 * subscription-scoped resource names. ARM reports the collision in two
 * different shapes depending on how far the request got:
 *
 *   - `400 InvalidBotData` — "The bot name is already registered to another
 *     bot application" (the common case; the name is taken globally),
 *   - `409 Conflict` — the ARM resource name itself belongs to a foreign
 *     resource.
 *
 * Both are DETERMINISTIC: the identical request will fail identically
 * forever, so this error is deliberately excluded from
 * {@link isTransientProvisioningFailure}. The job runner
 * (byte5ai/omadia#864) branches on it to fail fast instead of burning a
 * retry budget on a verdict that cannot change.
 */
export class BotHandleUnavailableError extends TeamsProvisionerError {
  /** The handle ARM refused. */
  public readonly botName: string;
  /** How ARM reported the collision. */
  public readonly status: number;

  constructor(botName: string, status: number, cause?: unknown) {
    super(
      `bot_handle_unavailable: Azure bot handle '${botName}' is already ` +
        'registered to another bot application. Bot handles share ONE global ' +
        'namespace across all Azure customers (like a DNS label) — they are ' +
        'not scoped to your tenant, subscription or resource group. omadia ' +
        'qualifies the handle automatically from the agent slug plus the ' +
        "app registration's id, so a collision here means that qualified " +
        'name is taken too — rename the agent (bot slug) and re-run ' +
        'provisioning',
      cause,
    );
    this.name = 'BotHandleUnavailableError';
    this.botName = botName;
    this.status = status;
  }
}

/**
 * A chat install/uninstall named a chat Graph will not resolve: it does not
 * exist, or it is invisible to this tenant app.
 *
 * A SEPARATE CLASS FROM THE TEAM DIRECTION ON PURPOSE. Both are "404 on the
 * install target", and that is where the similarity ends — the operator's next
 * move is different enough that collapsing them costs a support round trip:
 *
 * - a team 404 means a group id that is wrong or a team that was deleted;
 * - a chat 404 means a thread id that was copied from the wrong place —
 *   overwhelmingly a CHANNEL id (`@thread.tacv2`), which is a well-formed
 *   Teams id for something an app is never installed into (see
 *   {@link InstallTargetMismatchError}, which catches that shape before the
 *   network) — or a chat the app-only identity genuinely cannot see.
 *
 * Deterministic, never transient: the identical call answers identically until
 * somebody changes the id or the chat. It is therefore excluded from
 * {@link isTransientProvisioningFailure} explicitly rather than by falling
 * through to the message-pattern default.
 */
export class ChatNotFoundError extends TeamsProvisionerError {
  /** The chat thread id Graph did not resolve. */
  public readonly chatId: string;
  /** Which step asked, e.g. `chats.installedApps.add`. */
  public readonly step: string;

  constructor(chatId: string, step: string, cause?: unknown) {
    super(
      `chat_not_found: '${step}' found no chat '${chatId}' — it does not ` +
        'exist, or it is not visible to this tenant app. Check that the id is ' +
        'a CHAT thread id (`19:…@thread.v2` for a group chat, ' +
        '`19:…@unq.gbl.spaces` for a 1:1 chat) and not a channel id ' +
        '(`19:…@thread.tacv2`), and that the app role ' +
        'TeamsAppInstallation.ReadWriteForChat.All is consented for the tenant',
      cause,
    );
    this.name = 'ChatNotFoundError';
    this.chatId = chatId;
    this.step = step;
  }
}

/**
 * Graph answered `400 ResourceSpecificPermissionsMismatch` to an install.
 *
 * NOT a generic bad request, and NOT a wrong target id: the install body is
 * well-formed and the team/chat exists. What Graph refuses is the RSC consent
 * itself — the app package declares resource-specific permissions, and either
 * the installing identity's app role may not consent to them, or the
 * `consentedPermissionSet` in the body does not match what the app's
 * `teamsAppDefinition` declares.
 *
 * THE TWO CASES READ DIFFERENTLY ON PURPOSE, keyed by
 * {@link sentPermissionCount}. A request that carried the set and was still
 * refused points at the ROLE; a request that carried none points at the
 * resolution that produced none. Reporting both as "grant the role" sent the
 * operator to a portal blade that was already correct
 * (byte5ai/omadia-m365-connector 0.8.2).
 *
 * Deterministic: no retry makes a missing grant appear, so it is excluded from
 * {@link isTransientProvisioningFailure}.
 */
export class RscPermissionsMismatchError extends TeamsProvisionerError {
  /** Which step was refused, e.g. `chats.installedApps.add`. */
  public readonly step: string;
  /** App role that lets THIS scope's install carry RSC consent. */
  public readonly consentRole: string;
  /**
   * How many resource-specific permissions the request actually carried.
   * `0` means the body had no `consentedPermissionSet` at all — a different
   * situation with a different remedy, which is why the count is a field and
   * not just prose.
   */
  public readonly sentPermissionCount: number;

  constructor(
    step: string,
    consentRole: string,
    sentPermissionCount: number,
    graphDetail: string,
    cause?: unknown,
  ) {
    super(
      // Graph's own code stays in the text verbatim: consumers (the omadia
      // middleware among them) duck-type on it rather than importing this
      // class, so removing it would silently reclassify the failure.
      `rsc_permissions_mismatch: '${step}' was refused with ` +
        'ResourceSpecificPermissionsMismatch — ' +
        (sentPermissionCount > 0
          ? `the request DID carry a consentedPermissionSet of ${String(sentPermissionCount)} ` +
            'resource-specific permission(s) read from the app package itself, so the ' +
            `set is not what is missing: the installing identity may not consent to it. Grant the app role ${consentRole} ` +
            'and admin-consent it in the customer tenant'
          : 'the app package declares resource-specific permissions but NO ' +
            'consentedPermissionSet could be resolved for it — the catalog ' +
            'lookup of the app definition returned none (see the log line ' +
            `above). Grant ${consentRole} AND AppCatalog.ReadWrite.All, then ` +
            're-run provisioning') +
        `. Graph said: ${graphDetail}`,
      cause,
    );
    this.name = 'RscPermissionsMismatchError';
    this.step = step;
    this.consentRole = consentRole;
    this.sentPermissionCount = sentPermissionCount;
  }
}

/**
 * The identifier handed to an install step names a Teams scope the step cannot
 * act on — a channel id passed to the chat install, a team GUID passed to the
 * chat install, or a shape that is no Teams identifier at all.
 *
 * PRE-FLIGHT, BY DESIGN. This is thrown from the SHAPE of the id
 * (`classifyInstallTarget`), before a token is fetched or a request is sent,
 * because the alternative is a Graph 404 whose message names none of the
 * actual problem. {@link hint} carries the remedy in operator language — for
 * the channel case, that an app is installed into the channel's TEAM.
 *
 * Deterministic: retrying the identical id cannot help, so it is excluded from
 * {@link isTransientProvisioningFailure}.
 */
export class InstallTargetMismatchError extends TeamsProvisionerError {
  /** The identifier as classified (trimmed). */
  public readonly value: string;
  /** What the identifier turned out to be — `InstallTargetKind`. */
  public readonly targetKind: string;
  /** Which step refused it, e.g. `chats.installedApps.add`. */
  public readonly step: string;
  /** Operator-facing remedy, from the classification. */
  public readonly hint: string;

  constructor(
    step: string,
    value: string,
    targetKind: string,
    hint: string,
  ) {
    super(
      `install_target_mismatch: '${step}' cannot use '${value}' — ${hint}`,
    );
    this.name = 'InstallTargetMismatchError';
    this.step = step;
    this.value = value;
    this.targetKind = targetKind;
    this.hint = hint;
  }
}

// ---------------------------------------------------------------------------
// Delegated catalog-publish taxonomy (byte5ai/omadia#924).
//
// `POST /appCatalogs/teamsApps` is DELEGATED-ONLY — Graph documents
// application permissions as "Not supported." for that verb, and the field
// test confirmed it: the same app-only token that RESOLVES a catalog app is
// rejected for the UPLOAD even with `AppCatalog.ReadWrite.All` assigned as an
// app role. No amount of admin consent changes that.
//
// So exactly one step of the chain runs on a user token, and the middleware
// has to tell three failure modes apart because each carries a DIFFERENT
// operator instruction:
//
//   - {@link DelegatedSignInRequiredError}  → "no admin has signed in yet"
//     → start the device-code flow.
//   - {@link DelegatedConsentRequiredError} → "signed in, but the tenant never
//     consented to the delegated scope" → send an admin to the consent URL.
//   - {@link DelegatedTokenExpiredError}    → "the stored token stopped
//     working" → refresh, or sign in again when the refresh token is dead too.
//
// Collapsing these into one error is what makes an operator retry the wrong
// remedy three times — hence three classes.
// ---------------------------------------------------------------------------

/**
 * The delegated catalog upload was attempted without a delegated token.
 *
 * Neither a misconfiguration nor transient: it means no tenant admin has
 * completed the one-time device-code sign-in yet. The caller's next move is to
 * start that flow — never to retry the upload.
 */
export class DelegatedSignInRequiredError extends TeamsProvisionerError {
  /** Which step needs the user token, e.g. `appCatalogs.teamsApps.publish`. */
  public readonly step: string;
  /** Delegated scopes the sign-in has to acquire. */
  public readonly requiredScopes: readonly string[];

  constructor(step: string, requiredScopes: readonly string[]) {
    super(
      `delegated_sign_in_required: '${step}' needs a delegated (user) access ` +
        'token. Microsoft Graph does not support application permissions for ' +
        'publishing to the tenant app catalog, so one tenant admin has to ' +
        'complete the device-code sign-in once. Required delegated scope(s): ' +
        requiredScopes.join(', '),
    );
    this.name = 'DelegatedSignInRequiredError';
    this.step = step;
    this.requiredScopes = requiredScopes;
  }
}

/**
 * A delegated call was rejected because the TENANT has not consented to the
 * delegated scope — distinct from "nobody signed in" and from "token expired".
 *
 * `AppCatalog.ReadWrite.All` is an admin-consent-required delegated scope, so a
 * non-admin sign-in — or a tenant whose user-consent policy is "do not allow
 * user consent" — lands here. {@link adminConsentUrl} is the exact link to hand
 * an admin; it carries only the public client id, never a secret.
 */
export class DelegatedConsentRequiredError extends TeamsProvisionerError {
  /** Delegated scopes that lack tenant consent. */
  public readonly requiredScopes: readonly string[];
  /** Tenant-wide admin-consent URL for the publisher app (secret-free). */
  public readonly adminConsentUrl: string;
  /** Which step was rejected. */
  public readonly step: string;

  constructor(
    step: string,
    requiredScopes: readonly string[],
    adminConsentUrl: string,
    cause?: unknown,
  ) {
    super(
      `delegated_consent_required: '${step}' was rejected because the tenant ` +
        'has not granted admin consent for the delegated scope(s) ' +
        `${requiredScopes.join(', ')}. Send a Global Administrator (or Cloud ` +
        'Application Administrator) to the consent URL carried on this error, ' +
        'then retry — signing in again is not enough on its own.',
      cause,
    );
    this.name = 'DelegatedConsentRequiredError';
    this.step = step;
    this.requiredScopes = requiredScopes;
    this.adminConsentUrl = adminConsentUrl;
  }
}

/**
 * The stored delegated credential no longer works. Two flavours, and the
 * caller's response differs:
 *
 * - `'access-token-expired'` — the access token aged out (~1 h). Recoverable
 *   WITHOUT a human: exchange the refresh token.
 * - `'refresh-token-invalid'` — Entra rejected the refresh token
 *   (`invalid_grant`): revoked, password changed, re-evaluated by Conditional
 *   Access, or simply unused past its inactivity window. A human has to sign
 *   in again.
 */
export class DelegatedTokenExpiredError extends TeamsProvisionerError {
  public readonly reason: 'access-token-expired' | 'refresh-token-invalid';
  /** `true` when a refresh exchange can fix this without a human. */
  public readonly recoverableByRefresh: boolean;

  constructor(
    reason: 'access-token-expired' | 'refresh-token-invalid',
    cause?: unknown,
  ) {
    super(
      reason === 'access-token-expired'
        ? 'delegated_token_expired: the delegated access token has expired — ' +
            'exchange the stored refresh token for a fresh one and retry'
        : 'delegated_token_expired: Entra rejected the stored refresh token ' +
            '(revoked, expired, or invalidated by a credential/policy change) ' +
            '— a tenant admin has to complete the device-code sign-in again',
      cause,
    );
    this.name = 'DelegatedTokenExpiredError';
    this.reason = reason;
    this.recoverableByRefresh = reason === 'access-token-expired';
  }
}

/**
 * The device-code protocol itself failed in a way that is neither "still
 * waiting" nor an ordinary terminal outcome.
 *
 * The three EXPECTED terminal states (`authorization_pending`,
 * `expired_token`, `authorization_declined`) are deliberately NOT errors —
 * they surface as the poll result's discriminant so the caller renders a
 * status instead of catching. This class covers the rest: a malformed handle,
 * an unregistered/mis-typed client, a tenant that blocks device code flow via
 * Conditional Access, or a token endpoint answering something unparseable.
 */
export class DeviceCodeFlowError extends TeamsProvisionerError {
  /** OAuth `error` field when Entra sent one, e.g. `invalid_client`. */
  public readonly oauthError?: string;
  /** HTTP status of the devicecode/token response, when there was one. */
  public readonly status?: number;

  constructor(
    message: string,
    detail?: { readonly oauthError?: string; readonly status?: number },
    cause?: unknown,
  ) {
    super(`device_code_flow_failed: ${message}`, cause);
    this.name = 'DeviceCodeFlowError';
    if (detail?.oauthError !== undefined) this.oauthError = detail.oauthError;
    if (detail?.status !== undefined) this.status = detail.status;
  }
}

/**
 * A delegated read was asked for with a credential that cannot perform it —
 * either none was passed, or the stored one carries the wrong scopes.
 *
 * DISTINCT FROM {@link DelegatedSignInRequiredError} on purpose. That one
 * explains the catalog-publish situation ("Graph supports no application
 * permission for this verb"), which is a different fact about a different
 * endpoint; reusing it for a chat listing would hand the operator a sentence
 * about the app catalog while they are looking at a target picker.
 *
 * DISTINCT FROM {@link DelegatedConsentRequiredError} too: nothing has been
 * REJECTED here. This is a pre-flight verdict about the credential in hand,
 * raised before a Graph call is spent producing a misleading answer.
 *
 * Both reasons share one remedy — run the device-code sign-in again — but
 * they read very differently to an operator, so {@link reason} carries which
 * one it is instead of forcing a caller to parse the message.
 */
export class DelegatedScopeRequiredError extends TeamsProvisionerError {
  /** Which step needs the user token, e.g. `chats.list`. */
  public readonly step: string;
  /** Delegated scopes the sign-in has to acquire. */
  public readonly requiredScopes: readonly string[];
  /** `'no-token'`: nobody signed in. `'scope-missing'`: signed in, too narrow. */
  public readonly reason: 'no-token' | 'scope-missing';
  /** Scopes the stored credential DOES carry (empty for `'no-token'`). */
  public readonly grantedScopes: readonly string[];

  constructor(
    step: string,
    requiredScopes: readonly string[],
    reason: 'no-token' | 'scope-missing',
    grantedScopes: readonly string[] = [],
  ) {
    super(
      reason === 'no-token'
        ? `delegated_scope_required: '${step}' reads on behalf of the signed-in ` +
            'administrator and no delegated credential was passed. Microsoft ' +
            'Graph offers no tenant-wide application-permission route for this ' +
            'data, so an admin has to complete the device-code sign-in once. ' +
            `Required delegated scope(s): ${requiredScopes.join(', ')}`
        : `delegated_scope_required: '${step}' needs the delegated scope(s) ` +
            `${requiredScopes.join(', ')}, which the stored credential does not ` +
            `carry (it has: ${grantedScopes.length > 0 ? grantedScopes.join(', ') : 'none'}). ` +
            'A credential minted before this scope was requested cannot grow ' +
            'one by refreshing — the administrator has to sign in again',
    );
    this.name = 'DelegatedScopeRequiredError';
    this.step = step;
    this.requiredScopes = requiredScopes;
    this.reason = reason;
    this.grantedScopes = grantedScopes;
  }
}

/**
 * A purge was addressed with an APPLICATION (client) id instead of the
 * directory OBJECT id, and the recycle bin proved it.
 *
 * WHY THIS DESERVES ITS OWN CLASS. `DELETE /directory/deletedItems/{id}` takes
 * the object id; both ids are GUIDs, so the wrong one does not fail loudly —
 * Graph answers a perfectly ordinary 404, indistinguishable from "already
 * purged". Reporting that as `'already-absent'` would tell an operator the
 * `uniqueName` is free while the entry, and its 30-day reservation, is still
 * sitting in the bin — the exact shape of the confusion that cost a slug a
 * month in byte5ai/omadia#916. The error carries the id that WOULD have
 * worked, so the fix is a retry with a value the caller already has.
 */
export class DeletedObjectIdMismatchError extends TeamsProvisionerError {
  /** What the caller passed — an `appId`, as it turns out. */
  public readonly passedId: string;
  /** The directory object id the recycle bin actually holds it under. */
  public readonly objectId: string;

  constructor(passedId: string, objectId: string) {
    super(
      `deleted_object_id_mismatch: '${passedId}' is the application (client) ` +
        'id, but DELETE /directory/deletedItems addresses the DIRECTORY OBJECT ' +
        `id. The recycle bin holds this app under objectId '${objectId}' — ` +
        'retry the purge with that. Reporting the 404 as "already absent" ' +
        'would have claimed the uniqueName was freed while the reservation ' +
        `still stands for up to ${String(DELETED_ITEM_RETENTION_DAYS)} days`,
    );
    this.name = 'DeletedObjectIdMismatchError';
    this.passedId = passedId;
    this.objectId = objectId;
  }
}

/**
 * HTTP statuses that mean "the same call can succeed later": throttling,
 * request timeouts and the 5xx family. Deliberately NOT 404 — a bare
 * not-found is a legitimate terminal answer for most steps; the ONE 404 that
 * is transient (replication after create) surfaces as
 * {@link DirectoryReplicationError} instead.
 */
const TRANSIENT_HTTP_STATUSES: ReadonlySet<number> = new Set([
  408, 425, 429, 500, 502, 503, 504,
]);

/** Transport-level failures that are retryable by nature (no HTTP status). */
const TRANSIENT_TRANSPORT_PATTERN =
  /fetch failed|network|socket hang up|ECONNRESET|ECONNREFUSED|ETIMEDOUT|EAI_AGAIN|ENOTFOUND/i;

/**
 * Would re-running this step later plausibly succeed?
 *
 * The app-registration step branches on this to decide whether a partial
 * failure may be ROLLED BACK. Rolling back a transient failure is what burned
 * a provisioning slug for 30 days in byte5ai/omadia#916: the delete
 * soft-deleted a perfectly good app and reserved its `uniqueName`, so every
 * retry then collided with an object nobody could see.
 */
export function isTransientProvisioningFailure(err: unknown): boolean {
  if (err instanceof ProvisioningThrottledError) return true;
  if (err instanceof DirectoryReplicationError) return true;
  // A taken global bot handle is a verdict, not a hiccup — never retry it.
  if (err instanceof BotHandleUnavailableError) return false;
  // Same shape of verdict for the chat direction: a missing/invisible chat and
  // a wrong-kind target id both answer identically until a human changes the
  // id. Listed explicitly so neither leans on the message-pattern fallback.
  if (err instanceof ChatNotFoundError) return false;
  if (err instanceof InstallTargetMismatchError) return false;
  // A refused RSC consent is a tenant-side role grant, not a hiccup: replaying
  // the identical install cannot make the grant appear.
  if (err instanceof RscPermissionsMismatchError) return false;
  // Every delegated-auth failure needs a DIFFERENT action (sign in, consent,
  // refresh) — none of them is fixed by replaying the identical call, so they
  // are listed explicitly rather than left to the message-pattern fallback
  // below, which could match one of their explanatory sentences by accident.
  if (err instanceof DelegatedSignInRequiredError) return false;
  if (err instanceof DelegatedScopeRequiredError) return false;
  if (err instanceof DelegatedConsentRequiredError) return false;
  if (err instanceof DelegatedTokenExpiredError) return false;
  if (err instanceof DeviceCodeFlowError) return false;
  // Retrying a purge with the same wrong id produces the same wrong answer.
  if (err instanceof DeletedObjectIdMismatchError) return false;
  if (err instanceof ProvisioningRequestError) {
    return TRANSIENT_HTTP_STATUSES.has(err.status);
  }
  if (err instanceof Error) {
    if (err.name === 'AbortError' || err.name === 'TimeoutError') return true;
    return TRANSIENT_TRANSPORT_PATTERN.test(err.message);
  }
  return false;
}

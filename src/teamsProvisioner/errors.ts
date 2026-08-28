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
  if (err instanceof ProvisioningRequestError) {
    return TRANSIENT_HTTP_STATUSES.has(err.status);
  }
  if (err instanceof Error) {
    if (err.name === 'AbortError' || err.name === 'TimeoutError') return true;
    return TRANSIENT_TRANSPORT_PATTERN.test(err.message);
  }
  return false;
}

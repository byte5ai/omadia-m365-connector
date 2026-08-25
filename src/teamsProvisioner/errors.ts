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

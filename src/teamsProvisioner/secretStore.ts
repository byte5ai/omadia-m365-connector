/**
 * secretRef store for `teamsProvisioner@1` — WHERE a generated bot app
 * password lands, and how callers reference it afterwards.
 *
 * The Graph `addPassword` call (app-registration unit) returns the client
 * secret exactly once. This module persists that value into the plugin's own
 * secret vault via `ctx.secrets` and hands back an OPAQUE reference — the
 * deterministic vault key — so the middleware agent factory
 * (byte5ai/omadia#863-865) can store, pass around and later revoke the
 * credential without ever holding the password itself.
 *
 * Key scheme: `teams_bot_password:<appId>`. One agent = one Entra app, so
 * keying by the Entra application (client) id is collision-free across
 * agents, deterministic on re-runs (idempotent overwrite on secret rotation)
 * and reversible for the deleteAppRegistration rollback path
 * ({@link appIdFromSecretRef}).
 *
 * CAPABILITY GATE — `permissions.secrets.runtime_write`. Per the plugin-api
 * contract (`SecretsAccessor`, Spec 004), `secrets.set` / `secrets.delete`
 * are OPTIONAL: the kernel only hands them out when the manifest declares
 * `permissions.secrets.runtime_write`. Every write/delete here is therefore
 * guarded with `if (ctx.secrets.set)` and raises the typed
 * {@link CapabilityUnavailableError} when write access was not granted.
 * The manifest edit itself is owned by the WIRING unit, so all manifest
 * changes land in one version bump.
 *
 * INVARIANT — the secret VALUE never appears in a secretRef, a log line or
 * an error. Only the key/handle leaves this module.
 */

import type { SecretsAccessor } from '@omadia/plugin-api';

import { TeamsProvisionerError } from './errors.js';

/** Vault-key prefix for provisioned bot app passwords. */
export const TEAMS_BOT_PASSWORD_SECRET_PREFIX = 'teams_bot_password';

/**
 * Opaque reference to a stored bot app password: the deterministic vault key
 * `teams_bot_password:<appId>` — NEVER the secret value. Callers persist and
 * pass this; only code holding a write-capable `ctx.secrets` can resolve it.
 */
export type TeamsBotPasswordSecretRef =
  `${typeof TEAMS_BOT_PASSWORD_SECRET_PREFIX}:${string}`;

/**
 * Thrown when a runtime secret write/delete is attempted but the kernel did
 * not hand out `secrets.set` / `secrets.delete` — i.e. the manifest does not
 * declare `permissions.secrets.runtime_write` (or the kernel predates
 * Spec 004). Carries the manifest permission to declare so the operator/dev
 * message is actionable. Part of the `TeamsProvisionerError` taxonomy for
 * the one-`instanceof` catch-all.
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
 * Deterministic vault key for an Entra app's bot password. Same `appId` in →
 * same key out; distinct apps (= distinct agents) can never collide.
 *
 * The `appId` must be a plain Entra application (client) id — non-empty,
 * no whitespace, no `:` — so the ref stays unambiguous and reversible.
 */
export function secretRefForApp(appId: string): TeamsBotPasswordSecretRef {
  assertValidAppId(appId);
  return `${TEAMS_BOT_PASSWORD_SECRET_PREFIX}:${appId}`;
}

/**
 * Reverse of {@link secretRefForApp} — recovers the Entra app id from a
 * secretRef so the deleteAppRegistration rollback can correlate the vault
 * entry with the Graph `DELETE /applications/...` call. Throws on anything
 * that is not a well-formed `teams_bot_password:<appId>` ref.
 */
export function appIdFromSecretRef(secretRef: string): string {
  const prefix = `${TEAMS_BOT_PASSWORD_SECRET_PREFIX}:`;
  if (!secretRef.startsWith(prefix)) {
    throw new Error(`invalid_secret_ref: expected '${prefix}<appId>'`);
  }
  const appId = secretRef.slice(prefix.length);
  assertValidAppId(appId);
  return appId;
}

/**
 * Persist a freshly generated app password (Graph `addPassword` returns it
 * exactly once) into the plugin's secret vault and return the opaque
 * secretRef. Re-running with the same `appId` overwrites — that is the
 * intended idempotent/rotation behaviour of the deterministic key.
 *
 * @throws CapabilityUnavailableError when the kernel did not hand out
 *   `secrets.set` (manifest missing `permissions.secrets.runtime_write`).
 */
export async function storeAppPassword(
  secrets: SecretsAccessor,
  appId: string,
  secretText: string,
): Promise<TeamsBotPasswordSecretRef> {
  const secretRef = secretRefForApp(appId);
  if (secretText.length === 0) {
    // Guards against persisting an empty credential; message carries the
    // ref only — the value never reaches any error or log.
    throw new Error(`empty_secret_value: refusing to store '${secretRef}'`);
  }
  if (!secrets.set) {
    throw new CapabilityUnavailableError('set');
  }
  await secrets.set(secretRef, secretText);
  return secretRef;
}

/**
 * Rollback counterpart of {@link storeAppPassword}: remove the stored
 * password when deleting/rolling back an app registration. Validates the
 * ref shape first (never issues a vault delete for a foreign key), then
 * delegates to `secrets.delete` — which is a no-op if the key is absent.
 *
 * @throws CapabilityUnavailableError when the kernel did not hand out
 *   `secrets.delete` (manifest missing `permissions.secrets.runtime_write`).
 */
export async function deleteAppPassword(
  secrets: SecretsAccessor,
  secretRef: string,
): Promise<void> {
  appIdFromSecretRef(secretRef); // shape check — throws on malformed refs
  if (!secrets.delete) {
    throw new CapabilityUnavailableError('delete');
  }
  await secrets.delete(secretRef);
}

function assertValidAppId(appId: string): void {
  if (appId.length === 0 || /[\s:]/.test(appId)) {
    throw new Error(
      'invalid_app_id: expected a plain Entra application (client) id ' +
        '(non-empty, no whitespace, no colon)',
    );
  }
}

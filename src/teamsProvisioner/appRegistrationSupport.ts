/**
 * Pure helpers of the Entra app-registration step: argument validation, OData
 * literal handling, Graph response field access and the credential-label
 * bookkeeping the adopt path needs.
 *
 * Split out of `appRegistration.ts` so the step client stays about the CHAIN
 * (create → secret → service principal, replication windows, adoption,
 * rollback) and this module stays about VALUES. Everything here is
 * side-effect free and independently testable.
 */

import type { CreateAppRegistrationInput } from './appRegistration.js';

/** Portal label for the credential this provisioner manages on an app. */
export function secretLabel(input: CreateAppRegistrationInput): string {
  return input.secretDisplayName ?? `${input.displayName} bot password`;
}

/**
 * `keyId`s of password credentials that carry OUR label and are not the one
 * just added. Anything without that exact label belongs to someone else and
 * is never returned.
 */
export function staleCredentialKeyIds(
  json: unknown,
  label: string,
  keepKeyId: string,
): readonly string[] {
  if (!json || typeof json !== 'object') return [];
  const credentials = (json as Record<string, unknown>)['passwordCredentials'];
  if (!Array.isArray(credentials)) return [];
  const keyIds: string[] = [];
  for (const credential of credentials) {
    if (!credential || typeof credential !== 'object') continue;
    const record = credential as Record<string, unknown>;
    const keyId = record['keyId'];
    if (typeof keyId !== 'string' || keyId === keepKeyId) continue;
    if (record['displayName'] !== label) continue;
    keyIds.push(keyId);
  }
  return keyIds;
}

export function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

/** Entra client ids are GUID-shaped — reject anything that could break the OData URL. */
export function requireAppId(appId: string): string {
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
export function requireUniqueName(uniqueName: string): string {
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
export function escapeODataQuotes(value: string): string {
  return value.replace(/'/g, "''");
}

export function asRecord(json: unknown, step: string): Record<string, unknown> {
  if (json === null || typeof json !== 'object') {
    throw new Error(`graph ${step}: unexpected empty/non-object response body`);
  }
  return json as Record<string, unknown>;
}

export function requireStringField(
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

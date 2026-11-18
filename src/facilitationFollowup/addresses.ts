/**
 * Address handling for the facilitation follow-up service.
 *
 * The grant is the whole security boundary of this service: a facilitator may
 * reach exactly the people the kernel listed, and nobody else. Everything here
 * is pure, so the boundary can be tested without any transport.
 */

import type { FollowupGrant } from '@omadia/plugin-api';

/**
 * A bare mailbox address: one `@`, no whitespace and none of the characters
 * that would let one string carry a second recipient or a display name
 * (`a@x.de; b@y.de`, `Name <a@x.de>`). Deliberately stricter than RFC 5322 —
 * a string that fails here is refused, never repaired.
 */
const BARE_ADDRESS_PATTERN = /^[^\s@,;<>()"]+@[^\s@,;<>()"]+\.[^\s@,;<>()"]+$/;

/** Comparison form of an address: trimmed and lower-cased. */
export function normaliseAddress(address: string): string {
  return address.trim().toLowerCase();
}

export function isBareAddress(address: string): boolean {
  return BARE_ADDRESS_PATTERN.test(address);
}

export type GrantCheck =
  | { readonly allowed: true; readonly addresses: readonly string[] }
  | { readonly allowed: false; readonly rejected: readonly string[] };

/**
 * Check every requested address against the grant. All or nothing: a single
 * address outside the grant refuses the whole call, and the caller learns
 * which ones. On success the addresses come back normalised and de-duplicated,
 * and only those are ever put into a Graph request.
 *
 * An address that is not a bare mailbox address is rejected even when the
 * same string appears in the grant.
 */
export function checkAgainstGrant(
  grant: FollowupGrant,
  requested: readonly string[],
): GrantCheck {
  // Defensive against a malformed call: anything that is not a list of
  // strings grants nothing, it never throws.
  const granted: readonly unknown[] = Array.isArray(grant?.allowed) ? grant.allowed : [];
  const allowed = new Set(
    granted.filter((a): a is string => typeof a === 'string').map(normaliseAddress),
  );
  const accepted = new Set<string>();
  const rejected: string[] = [];
  for (const raw of Array.isArray(requested) ? (requested as readonly unknown[]) : []) {
    const address = typeof raw === 'string' ? normaliseAddress(raw) : '';
    if (isBareAddress(address) && allowed.has(address)) {
      accepted.add(address);
    } else {
      rejected.push(String(raw));
    }
  }
  if (rejected.length > 0) return { allowed: false, rejected };
  return { allowed: true, addresses: [...accepted] };
}

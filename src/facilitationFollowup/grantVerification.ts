/**
 * The kernel's word on a grant, asked before every follow-up operation
 * (byte5ai/omadia#2078, spec 2043 §5.11, since 0.10.0).
 *
 * A caller that states `grant.grantId` runs under a kernel-issued grant. The
 * kernel then decides whether that grant is live (not revoked, not expired,
 * its execution scope open, no operator stop) and covers the operation and
 * every address it reaches. This module asks and turns the answer into one of
 * three results; `grant.allowed` stays the provider's own second check.
 *
 *   allowed     the kernel said yes
 *   refused     the kernel said no, gave no usable answer, or offers no
 *               `grantVerification` service at all (fail closed: a caller
 *               that states a grant expects it to be checked)
 *   not_asked   there was nothing to ask: the caller stated no grant id
 *
 * Nothing here throws.
 */

import type {
  GRANT_VERIFICATION_SERVICE_NAME as CONTRACT_SERVICE_NAME,
  FollowupGrant,
  FollowupRefusal,
  GrantOperation,
  GrantVerificationService,
} from '@omadia/plugin-api';

import { truncate } from '../graphClient.js';
import { redactSecrets } from '../teamsProvisioner/redact.js';
import { normaliseAddress } from './addresses.js';

/**
 * ServiceRegistry key, restated instead of imported as a value so the package
 * still loads on a kernel whose plugin-api has no such export.
 */
export const GRANT_VERIFICATION_SERVICE_NAME: typeof CONTRACT_SERVICE_NAME = 'grantVerification';

/** How long the kernel may take to answer before the operation is refused. */
export const GRANT_VERIFICATION_TIMEOUT_MS = 10_000;
const MAX_KERNEL_MESSAGE_LENGTH = 200;
const MAX_KERNEL_CODE_LENGTH = 40;

/**
 * `audit` is the value of the `grant=` field of the audit line:
 * `allowed`, `refused:<kernel code>`, `refused:unavailable`,
 * `refused:no_kernel_service` or `not_asked:no_grant_id`. (The service adds
 * `not_asked:refused_locally` for a call its own checks ended first.)
 */
export type GrantVerdict =
  | { readonly verdict: 'allowed'; readonly audit: string }
  | { readonly verdict: 'not_asked'; readonly audit: string }
  | { readonly verdict: 'refused'; readonly audit: string; readonly refusal: FollowupRefusal };

export interface GrantGateOptions {
  /**
   * Looks the kernel service up at every call (it is published after this
   * plugin activates). `undefined` means this kernel has none.
   */
  readonly lookup?: () => GrantVerificationService | undefined;
  readonly timeoutMs?: number;
}

export interface GrantGate {
  check(
    grant: FollowupGrant,
    operation: GrantOperation,
    people: readonly string[],
  ): Promise<GrantVerdict>;
}

export function createGrantGate(opts: GrantGateOptions): GrantGate {
  const timeoutMs = opts.timeoutMs ?? GRANT_VERIFICATION_TIMEOUT_MS;
  const refusedWithoutAnswer = (audit: string, message: string): GrantVerdict => ({
    verdict: 'refused',
    audit,
    refusal: { outcome: 'refused', code: 'grant_refused', message },
  });
  const unavailable = (reason: string): GrantVerdict =>
    refusedWithoutAnswer(
      'refused:unavailable',
      `the grant could not be verified with the kernel (${reason}); nothing was sent`,
    );

  return {
    async check(grant, operation, people) {
      const grantId: unknown = grant?.grantId;
      if (grantId === undefined) return { verdict: 'not_asked', audit: 'not_asked:no_grant_id' };
      if (typeof grantId !== 'string' || grantId.trim().length === 0) {
        return unavailable('the grant id is not a string');
      }

      let service: GrantVerificationService | undefined;
      try {
        service = opts.lookup?.();
      } catch (err) {
        return unavailable(`service lookup failed: ${safe(messageOf(err))}`);
      }
      if (service === undefined) {
        return refusedWithoutAnswer(
          'refused:no_kernel_service',
          `the kernel offers no grant verification (no '${GRANT_VERIFICATION_SERVICE_NAME}' service), so the stated grant cannot be checked; nothing was sent`,
        );
      }
      if (typeof service.verify !== 'function') return unavailable('the service has no verify()');

      let answer: unknown;
      try {
        answer = await withTimeout(
          service.verify({ grantId, operation, recipients: recipientsOf(people) }),
          timeoutMs,
        );
      } catch (err) {
        return unavailable(safe(messageOf(err)));
      }
      const outcome = asRecord(answer);
      if (outcome?.['allowed'] === true) return { verdict: 'allowed', audit: 'allowed' };
      if (outcome?.['allowed'] !== false || typeof outcome['code'] !== 'string') {
        return unavailable('the kernel gave no usable answer');
      }
      const code = outcome['code'].replace(/[^\w]/g, '_').slice(0, MAX_KERNEL_CODE_LENGTH);
      const message = typeof outcome['message'] === 'string' ? safe(outcome['message']) : '';
      const notCovered = Array.isArray(outcome['notCovered'])
        ? outcome['notCovered'].filter((a): a is string => typeof a === 'string')
        : [];
      return {
        verdict: 'refused',
        audit: `refused:${code}`,
        refusal: {
          outcome: 'refused',
          code: 'grant_refused',
          message: `the kernel refused the grant (${code})${message ? `: ${message}` : ''}`,
          ...(notCovered.length > 0 ? { rejected: notCovered } : {}),
        },
      };
    },
  };
}

/**
 * Everybody the operation would reach, as the kernel compares them. Nothing is
 * dropped: an entry that is not an address goes to the kernel as it is and the
 * kernel refuses the request.
 */
function recipientsOf(people: unknown): string[] {
  const list: readonly unknown[] = Array.isArray(people) ? people : [];
  return [...new Set(list.map((raw) => (typeof raw === 'string' ? normaliseAddress(raw) : String(raw))))];
}

function withTimeout<T>(pending: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`no answer within ${String(ms)} ms`));
    }, ms);
    Promise.resolve(pending).then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err: unknown) => {
        clearTimeout(timer);
        reject(err instanceof Error ? err : new Error(String(err)));
      },
    );
  });
}

function messageOf(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function safe(text: string): string {
  return truncate(redactSecrets(text), MAX_KERNEL_MESSAGE_LENGTH);
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

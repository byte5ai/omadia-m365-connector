/**
 * Setup field of the facilitation follow-up service: the mailbox that
 * organises the follow-up meetings and sends the protocol mail.
 *
 * Optional, with no default. A facilitator sends in the name of the instance,
 * so which mailbox that is must be an operator decision; without it the
 * service is still published and answers `not_configured`.
 */

import type { ConfigAccessor } from '@omadia/plugin-api';

import { isBareAddress, normaliseAddress } from './addresses.js';

/** Setup-field key: organiser and sender mailbox for facilitation follow-ups. */
export const FACILITATION_ORGANIZER_MAILBOX_FIELD = 'facilitation_organizer_mailbox';

export type OrganizerMailboxConfig =
  | { readonly kind: 'configured'; readonly mailbox: string }
  | { readonly kind: 'not-configured'; readonly reason: string };

/** Never throws: a missing or malformed value is a result, not an error. */
export function readOrganizerMailbox(
  config: Pick<ConfigAccessor, 'get'>,
): OrganizerMailboxConfig {
  const raw = config.get<unknown>(FACILITATION_ORGANIZER_MAILBOX_FIELD);
  if (typeof raw !== 'string' || raw.trim().length === 0) {
    return {
      kind: 'not-configured',
      reason: `setup field '${FACILITATION_ORGANIZER_MAILBOX_FIELD}' is empty`,
    };
  }
  const mailbox = normaliseAddress(raw);
  if (!isBareAddress(mailbox)) {
    return {
      kind: 'not-configured',
      reason: `setup field '${FACILITATION_ORGANIZER_MAILBOX_FIELD}' is not a mailbox address`,
    };
  }
  return { kind: 'configured', mailbox };
}

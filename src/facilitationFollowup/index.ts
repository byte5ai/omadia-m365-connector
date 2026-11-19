/**
 * Facilitation follow-up service (`facilitation.followup`, byte5ai/omadia#2043).
 * Contract types live in `@omadia/plugin-api`; this is the Microsoft 365
 * provider.
 */

export {
  createFacilitationFollowupService,
  FACILITATION_FOLLOWUP_CAPABILITY,
  FACILITATION_FOLLOWUP_SERVICE_NAME,
} from './service.js';
export type { FacilitationFollowupOptions } from './service.js';
export {
  createGrantGate,
  GRANT_VERIFICATION_SERVICE_NAME,
  GRANT_VERIFICATION_TIMEOUT_MS,
} from './grantVerification.js';
export type { GrantGate, GrantGateOptions, GrantVerdict } from './grantVerification.js';
export {
  FACILITATION_ORGANIZER_MAILBOX_FIELD,
  readOrganizerMailbox,
} from './config.js';
export type { OrganizerMailboxConfig } from './config.js';
export { createAppOnlyGraphTransport } from './transport.js';
export type {
  AppOnlyGraphTransportOptions,
  FollowupGraphOperation,
  FollowupGraphRequest,
  FollowupGraphResponse,
  FollowupGraphTransport,
} from './transport.js';

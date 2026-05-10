/**
 * @omadia/integration-microsoft365 — public surface.
 *
 * Ships both the library exports (clients, types, errors) and the
 * plugin-form `activate()` entry point. The plugin registers a
 * `Microsoft365Accessor` under the ServiceRegistry name
 * `microsoft365.graph`; consumers read it via
 * `ctx.services.get<Microsoft365Accessor>('microsoft365.graph')`.
 */

export { activate } from './plugin.js';
export type { Microsoft365PluginHandle } from './plugin.js';

export {
  MICROSOFT365_SERVICE_NAME,
  type Microsoft365Accessor,
  type SlotCacheAccessor,
} from './accessor.js';

export { GraphClient, encodeSharingUrl } from './graphClient.js';
export type {
  GraphClientOptions,
  GraphChatMessageAttachment,
} from './graphClient.js';

export {
  GraphOboClient,
  ConsentRequiredError,
  SsoUnavailableError,
  CALENDAR_GRAPH_SCOPES,
  createGraphOboClient,
} from './graphObo.js';
export type { GraphOboConfig } from './graphObo.js';

export { GraphCalendarClient } from './graphCalendarClient.js';
export type {
  AttendeeType,
  FindSlotsOptions,
  MeetingSlotSuggestion,
  GetScheduleOptions,
  ScheduleEntry,
  CreateEventOptions,
  CreatedEvent,
  MailboxSettings,
} from './graphCalendarClient.js';

export { SlotCache } from './slotCache.js';
export type { CachedSlot, SlotCacheOptions } from './slotCache.js';

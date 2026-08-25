import type { GraphClient } from './graphClient.js';
import type { GraphOboClient } from './graphObo.js';
import type { GraphCalendarClient } from './graphCalendarClient.js';
import type { CachedSlot } from './slotCache.js';

/**
 * Cache-facing subset of SlotCache exposed to consumers via the
 * Microsoft365Accessor. The full SlotCache class stays an internal
 * implementation detail — tools only need put/get/consume to round-trip
 * opaque slot ids through Adaptive-Card Action.Submit.
 */
export interface SlotCacheAccessor {
  put(entry: Omit<CachedSlot, 'slotId' | 'expiresAt'>): CachedSlot;
  get(slotId: string): CachedSlot | undefined;
  consume(slotId: string): CachedSlot | undefined;
}

/**
 * Service surface published by `@omadia/integration-microsoft365`
 * under the ServiceRegistry key `microsoft365.graph`.
 *
 * NOTE: this is one of TWO services the plugin publishes — the per-agent
 * Teams provisioning surface lives separately as `TeamsProvisionerAccessor`
 * under the key `teamsProvisioner` (capability `teamsProvisioner@1`, see
 * `src/teamsProvisioner/index.ts`). Deliberately not folded in here: its
 * consumers (the middleware agent factory) and permission profile
 * (`secrets.runtime_write`, ARM egress) differ from the Graph accessor's.
 *
 * Consumers (Teams channel, calendar tools, future Mail / OneDrive plugins)
 * read from here instead of holding their own Graph-client references:
 *
 *   const ms = ctx.services.get<Microsoft365Accessor>('microsoft365.graph');
 *   if (!ms) throw new Error('microsoft365 integration not installed');
 *
 * Four tiers of capability:
 *   - `app`: Microsoft-Graph client authenticated with the Bot-Framework
 *     App Registration (client-credentials flow). Tenant-level reads that
 *     do not hang off a user session — Teams attachment downloads.
 *   - `obo`: Delegated / OBO-style access-token acquisition for Graph on
 *     behalf of a single Teams user. Hand the token to `calendar.*`.
 *   - `calendar`: Stateless Graph Calendar wrapper. Each call takes a
 *     per-user access-token from `obo`.
 *   - `slots`: SlotCache surface the calendar tools use to round-trip
 *     opaque slot ids across Adaptive-Card submissions.
 */
export interface Microsoft365Accessor {
  readonly app: GraphClient;
  readonly obo: GraphOboClient;
  readonly calendar: GraphCalendarClient;
  readonly slots: SlotCacheAccessor;
}

/** Well-known ServiceRegistry key used by providers and consumers. */
export const MICROSOFT365_SERVICE_NAME = 'microsoft365.graph';

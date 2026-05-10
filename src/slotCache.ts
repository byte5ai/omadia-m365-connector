import { randomUUID } from 'node:crypto';

/**
 * In-memory TTL cache for calendar slots returned by `find_free_slots`.
 *
 * Why this exists: `book_meeting` needs the exact ISO start/end of a slot
 * the LLM *chose*, but the LLM has been known to drift a few minutes when
 * re-serialising times. The cache gives us an opaque `slotId` that the
 * Adaptive Card's Action.Submit round-trips back — no timestamp ever
 * leaves the cache boundary.
 *
 * Single-tenant, single-machine deployment on Fly → in-memory is fine;
 * a restart invalidates pending slots, but the UX degrades gracefully
 * (user re-runs the slot query). Swap for Redis if we ever go multi-instance.
 */
export interface CachedSlot {
  /** Opaque identifier the Adaptive Card submits back on click. */
  slotId: string;
  /** ISO-8601, UTC or with offset — passed straight to Graph `/me/events`. */
  start: string;
  /** ISO-8601. */
  end: string;
  /** Attendee UPNs/emails from the original `find_free_slots` call. */
  attendees: string[];
  /** Optional IANA TZ from the user's mailboxSettings; informational only. */
  timeZone?: string;
  /** Graph's own confidence score (0–100). Surfaced to the LLM for ranking. */
  confidence?: number;
  /** Epoch ms at which this entry expires and must be dropped. */
  expiresAt: number;
}

export interface SlotCacheOptions {
  /** Max age in ms. Default: 15 minutes. */
  ttlMs?: number;
  /** Hard cap on entries. Default: 2000. Oldest evicted first. */
  maxEntries?: number;
  /** Clock injection point for tests. */
  now?: () => number;
  /** RNG injection point for tests. */
  newId?: () => string;
}

const DEFAULT_TTL_MS = 60 * 60 * 1000;
const DEFAULT_MAX = 2000;

export class SlotCache {
  private readonly store = new Map<string, CachedSlot>();
  private readonly ttlMs: number;
  private readonly maxEntries: number;
  private readonly now: () => number;
  private readonly newId: () => string;

  constructor(opts: SlotCacheOptions = {}) {
    this.ttlMs = opts.ttlMs ?? DEFAULT_TTL_MS;
    this.maxEntries = opts.maxEntries ?? DEFAULT_MAX;
    this.now = opts.now ?? (() => Date.now());
    this.newId = opts.newId ?? (() => randomUUID());
  }

  /**
   * Register a slot and return its opaque id. Caller constructs the entry;
   * cache only fills in `slotId` + `expiresAt`.
   */
  put(entry: Omit<CachedSlot, 'slotId' | 'expiresAt'>): CachedSlot {
    this.evictExpired();
    if (this.store.size >= this.maxEntries) {
      const oldest = this.store.keys().next().value;
      if (oldest !== undefined) this.store.delete(oldest);
    }
    const slot: CachedSlot = {
      ...entry,
      slotId: this.newId(),
      expiresAt: this.now() + this.ttlMs,
    };
    this.store.set(slot.slotId, slot);
    return slot;
  }

  get(slotId: string): CachedSlot | undefined {
    const hit = this.store.get(slotId);
    if (!hit) return undefined;
    if (hit.expiresAt <= this.now()) {
      this.store.delete(slotId);
      return undefined;
    }
    return hit;
  }

  /** Remove after a successful book to prevent double-booking. */
  consume(slotId: string): CachedSlot | undefined {
    const hit = this.get(slotId);
    if (hit) this.store.delete(slotId);
    return hit;
  }

  size(): number {
    this.evictExpired();
    return this.store.size;
  }

  private evictExpired(): void {
    const cutoff = this.now();
    for (const [id, slot] of this.store) {
      if (slot.expiresAt <= cutoff) this.store.delete(id);
    }
  }
}

/**
 * `facilitation.followup` — the only reach a native facilitator has into mail
 * and calendars (epic byte5ai/omadia#2043, issue #18, kernel spec §5.11).
 *
 * Three operations, each on the configured organiser mailbox and each with
 * application permissions:
 *
 *   freeBusy          POST /users/{organiser}/calendar/getSchedule
 *   createInvitation  POST /users/{organiser}/events
 *   sendMail          POST /users/{organiser}/sendMail
 *
 * Order of checks in every operation, all before any Graph request:
 *   1. everyone the call names is inside `grant.allowed`, else
 *      `recipient_not_allowed` for the whole call;
 *   2. an organiser mailbox is configured, else `not_configured`;
 *   3. the request itself is well-formed, else `provider_error`.
 *
 * Nothing here throws: transport and token failures come back as
 * `provider_error`, because the consumer branches on `outcome`.
 */

import type {
  FACILITATION_FOLLOWUP_SERVICE_NAME as CONTRACT_SERVICE_NAME,
  FacilitationFollowupService,
  FollowupGrant,
  FollowupRefusal,
} from '@omadia/plugin-api';

import { truncate } from '../graphClient.js';
import { redactSecrets } from '../teamsProvisioner/redact.js';
import { checkAgainstGrant } from './addresses.js';
import type { OrganizerMailboxConfig } from './config.js';
import type {
  FollowupGraphOperation,
  FollowupGraphResponse,
  FollowupGraphTransport,
} from './transport.js';

/**
 * ServiceRegistry key. Restated here instead of imported as a value: this
 * package imports `@omadia/plugin-api` for types only, so it still loads on a
 * kernel older than 1.55.0 that has no such export. The annotation makes the
 * compiler fail if the contract's name ever differs.
 */
export const FACILITATION_FOLLOWUP_SERVICE_NAME: typeof CONTRACT_SERVICE_NAME =
  'facilitation.followup';
/** Versioned manifest ref of the service above (`provides:`). */
export const FACILITATION_FOLLOWUP_CAPABILITY = 'facilitation.followup@1';

/** Graph is asked for and answers in UTC, so no time zone is ever guessed. */
const UTC = 'UTC';
const PREFER_UTC_HEADER = { Prefer: 'outlook.timezone="UTC"' } as const;
const ONLINE_MEETING_PROVIDER = 'teamsForBusiness';
/**
 * `freeBusyStatus` values that do not block a slot. Everything else, including
 * `unknown` and values added later, counts as busy.
 */
const NOT_BUSY_STATUSES: ReadonlySet<string> = new Set(['free', 'workingElsewhere']);
/** An instant with an explicit offset; a local time would be ambiguous. */
const INSTANT_PATTERN =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})$/;
const MAX_ERROR_DETAIL_LENGTH = 200;
const MAX_LOGGED_ID_LENGTH = 80;

type Interval = { readonly start: string; readonly end: string };
type Prepared<T> =
  | { readonly ready: true; readonly mailbox: string; readonly addresses: readonly string[]; readonly value: T }
  | { readonly ready: false; readonly refusal: FollowupRefusal };

export interface FacilitationFollowupOptions {
  readonly organizer: OrganizerMailboxConfig;
  readonly transport: FollowupGraphTransport;
  /** Audit line per call: operation, facilitation, counts, outcome. No addresses, no content. */
  readonly log?: (msg: string) => void;
}

export function createFacilitationFollowupService(
  opts: FacilitationFollowupOptions,
): FacilitationFollowupService {
  const audit = (
    operation: string,
    grant: FollowupGrant,
    people: number,
    outcome: string,
  ): void => {
    opts.log?.(
      `[microsoft365] facilitation.followup op=${operation} facilitation=${loggableId(grant.facilitationId)} people=${String(people)} outcome=${outcome}`,
    );
  };

  const prepare = <T>(
    people: readonly string[],
    grant: FollowupGrant,
    validate: () => T | string,
  ): Prepared<T> => {
    const check = checkAgainstGrant(grant, people);
    if (!check.allowed) {
      return {
        ready: false,
        refusal: refuse(
          'recipient_not_allowed',
          `${String(check.rejected.length)} address(es) are outside the grant of this facilitation`,
          check.rejected,
        ),
      };
    }
    if (opts.organizer.kind !== 'configured') {
      return {
        ready: false,
        refusal: refuse(
          'not_configured',
          `no organiser mailbox for facilitation follow-ups: ${opts.organizer.reason}`,
        ),
      };
    }
    if (check.addresses.length === 0) {
      return { ready: false, refusal: refuse('provider_error', 'invalid request: no addresses given') };
    }
    const value = validate();
    if (typeof value === 'string') {
      return { ready: false, refusal: refuse('provider_error', `invalid request: ${value}`) };
    }
    return { ready: true, mailbox: opts.organizer.mailbox, addresses: check.addresses, value };
  };

  const post = async (
    operation: FollowupGraphOperation,
    mailbox: string,
    body: unknown,
    headers?: Readonly<Record<string, string>>,
  ): Promise<FollowupGraphResponse | FollowupRefusal> => {
    try {
      const response = await opts.transport.post({
        operation,
        mailbox,
        body,
        ...(headers ? { headers } : {}),
      });
      if (response.status >= 200 && response.status < 300) return response;
      return refuse('provider_error', describeGraphError(operation, response));
    } catch (err) {
      return refuse(
        'provider_error',
        `graph ${operation} failed: ${safeDetail(err instanceof Error ? err.message : String(err))}`,
      );
    }
  };

  return {
    async freeBusy(grant, request) {
      const prepared = prepare(request.attendees, grant, () =>
        toUtcWindow(request.from, request.to, 'from', 'to'),
      );
      if (!prepared.ready) {
        audit('freeBusy', grant, countOf(request.attendees), prepared.refusal.code);
        return prepared.refusal;
      }
      const response = await post(
        'getSchedule',
        prepared.mailbox,
        {
          schedules: prepared.addresses,
          startTime: { dateTime: prepared.value.start, timeZone: UTC },
          endTime: { dateTime: prepared.value.end, timeZone: UTC },
        },
        PREFER_UTC_HEADER,
      );
      if ('outcome' in response) {
        audit('freeBusy', grant, prepared.addresses.length, response.code);
        return response;
      }
      audit('freeBusy', grant, prepared.addresses.length, 'ok');
      return { outcome: 'ok', busy: readBusy(response.json, prepared.addresses) };
    },

    async createInvitation(grant, request) {
      const prepared = prepare(request.attendees, grant, () => {
        if (!isNonEmptyString(request.subject)) return 'subject is empty';
        if (typeof request.body !== 'string') return 'body is not a string';
        return toUtcWindow(request.start, request.end, 'start', 'end');
      });
      if (!prepared.ready) {
        audit('createInvitation', grant, countOf(request.attendees), prepared.refusal.code);
        return prepared.refusal;
      }
      const response = await post('createEvent', prepared.mailbox, {
        subject: request.subject,
        body: { contentType: 'text', content: request.body },
        start: { dateTime: prepared.value.start, timeZone: UTC },
        end: { dateTime: prepared.value.end, timeZone: UTC },
        attendees: prepared.addresses.map((address) => ({
          emailAddress: { address },
          type: 'required',
        })),
        ...(request.online === true
          ? { isOnlineMeeting: true, onlineMeetingProvider: ONLINE_MEETING_PROVIDER }
          : {}),
      });
      if ('outcome' in response) {
        audit('createInvitation', grant, prepared.addresses.length, response.code);
        return response;
      }
      // Only the id and the join link of the event just created are read.
      const created = asRecord(response.json);
      const eventId = created?.['id'];
      if (!isNonEmptyString(eventId)) {
        audit('createInvitation', grant, prepared.addresses.length, 'provider_error');
        return refuse('provider_error', 'graph createEvent answered without an event id');
      }
      const joinUrl = asRecord(created?.['onlineMeeting'])?.['joinUrl'];
      audit('createInvitation', grant, prepared.addresses.length, 'created');
      return {
        outcome: 'created',
        eventId,
        ...(isNonEmptyString(joinUrl) ? { joinUrl } : {}),
      };
    },

    async sendMail(grant, request) {
      const prepared = prepare(request.recipients, grant, () => {
        if (!isNonEmptyString(request.subject)) return 'subject is empty';
        if (!isNonEmptyString(request.body)) return 'body is empty';
        return true as const;
      });
      if (!prepared.ready) {
        audit('sendMail', grant, countOf(request.recipients), prepared.refusal.code);
        return prepared.refusal;
      }
      const response = await post('sendMail', prepared.mailbox, {
        message: {
          subject: request.subject,
          body: { contentType: 'text', content: request.body },
          toRecipients: prepared.addresses.map((address) => ({
            emailAddress: { address },
          })),
        },
        saveToSentItems: true,
      });
      if ('outcome' in response) {
        audit('sendMail', grant, prepared.addresses.length, response.code);
        return response;
      }
      audit('sendMail', grant, prepared.addresses.length, 'sent');
      return { outcome: 'sent' };
    },
  };
}

function refuse(
  code: FollowupRefusal['code'],
  message: string,
  rejected?: readonly string[],
): FollowupRefusal {
  return { outcome: 'refused', code, message, ...(rejected ? { rejected } : {}) };
}

/** Both ends as Graph `dateTime` in UTC, or the reason the window is unusable. */
function toUtcWindow(
  startInput: string,
  endInput: string,
  startName: string,
  endName: string,
): { start: string; end: string } | string {
  const start = parseInstant(startInput);
  if (start === undefined) return `${startName} is not an ISO 8601 instant with an offset`;
  const end = parseInstant(endInput);
  if (end === undefined) return `${endName} is not an ISO 8601 instant with an offset`;
  if (end <= start) return `${endName} is not after ${startName}`;
  return { start: toGraphDateTime(start), end: toGraphDateTime(end) };
}

function parseInstant(value: unknown): number | undefined {
  if (typeof value !== 'string' || !INSTANT_PATTERN.test(value)) return undefined;
  const ms = Date.parse(value);
  return Number.isNaN(ms) ? undefined : ms;
}

/** `2026-11-20T09:00:00` — Graph's `dateTimeTimeZone.dateTime`, zone given separately. */
function toGraphDateTime(ms: number): string {
  return new Date(ms).toISOString().slice(0, 19);
}

/**
 * Reduce a getSchedule answer to busy intervals per requested attendee.
 *
 * Only `status`, `start` and `end` of a schedule item are read; `subject`,
 * `location` and `isPrivate` are never touched. Overlapping and adjacent
 * intervals are merged, so the result does not show how many appointments
 * someone has either.
 *
 * An attendee whose schedule Graph could not resolve (an `error` entry, or no
 * entry at all) is left out: missing from the result means "availability
 * unknown", which is not the same as an empty `intervals` list.
 */
function readBusy(
  json: unknown,
  attendees: readonly string[],
): { attendee: string; intervals: Interval[] }[] {
  const entries = asRecord(json)?.['value'];
  if (!Array.isArray(entries)) return [];
  const requested = new Set(attendees);
  const result: { attendee: string; intervals: Interval[] }[] = [];
  for (const entry of entries) {
    const schedule = asRecord(entry);
    const scheduleId = schedule?.['scheduleId'];
    if (typeof scheduleId !== 'string' || schedule?.['error'] != null) continue;
    const attendee = scheduleId.trim().toLowerCase();
    if (!requested.delete(attendee)) continue;
    const items: unknown = schedule?.['scheduleItems'];
    result.push({ attendee, intervals: mergeIntervals(Array.isArray(items) ? items : []) });
  }
  return result;
}

function mergeIntervals(items: readonly unknown[]): Interval[] {
  const spans = items
    .map((raw) => {
      const item = asRecord(raw);
      const status = item?.['status'];
      if (typeof status === 'string' && NOT_BUSY_STATUSES.has(status)) return undefined;
      const start = parseGraphUtc(asRecord(item?.['start'])?.['dateTime']);
      const end = parseGraphUtc(asRecord(item?.['end'])?.['dateTime']);
      return start !== undefined && end !== undefined && end > start
        ? { start, end }
        : undefined;
    })
    .filter((span): span is { start: number; end: number } => span !== undefined)
    .sort((a, b) => a.start - b.start);

  const merged: { start: number; end: number }[] = [];
  for (const span of spans) {
    const last = merged[merged.length - 1];
    if (last && span.start <= last.end) {
      merged[merged.length - 1] = { start: last.start, end: Math.max(last.end, span.end) };
    } else {
      merged.push(span);
    }
  }
  return merged.map((span) => ({
    start: new Date(span.start).toISOString(),
    end: new Date(span.end).toISOString(),
  }));
}

/** Graph answers `2026-11-20T09:00:00.0000000` in the zone asked for (UTC). */
function parseGraphUtc(value: unknown): number | undefined {
  if (typeof value !== 'string') return undefined;
  const match = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.exec(value);
  if (!match) return undefined;
  const ms = Date.parse(`${match[0]}Z`);
  return Number.isNaN(ms) ? undefined : ms;
}

function describeGraphError(
  operation: FollowupGraphOperation,
  response: FollowupGraphResponse,
): string {
  const error = asRecord(asRecord(response.json)?.['error']);
  const code = typeof error?.['code'] === 'string' ? error['code'] : 'unknown';
  const message = typeof error?.['message'] === 'string' ? error['message'] : '';
  return `graph ${operation} ${String(response.status)} code=${code}${
    message ? ` message=${safeDetail(message)}` : ''
  }`;
}

function safeDetail(text: string): string {
  return truncate(redactSecrets(text), MAX_ERROR_DETAIL_LENGTH);
}

function loggableId(id: unknown): string {
  return truncate(String(id).replace(/[^\w.:-]/g, '_'), MAX_LOGGED_ID_LENGTH);
}

function countOf(list: unknown): number {
  return Array.isArray(list) ? list.length : 0;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

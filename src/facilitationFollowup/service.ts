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
 * `freeBusy` sends one getSchedule request per 20 attendees and refuses a
 * window of 62 days or more (both are documented Graph limits).
 *
 * Order of checks in every operation, all before any Graph request:
 *   1. everyone the call names is inside `grant.allowed`, else
 *      `recipient_not_allowed` for the whole call;
 *   2. an organiser mailbox is configured, else `not_configured`;
 *   3. the request itself is well-formed, else `provider_error`;
 *   4. when the call states `grant.grantId`: the kernel says the grant is live
 *      and covers the operation and everyone it reaches, else `grant_refused`
 *      (./grantVerification.ts; no answer and no such service refuse too).
 *      This is the last step before the token and Graph requests, so a use
 *      the kernel records as allowed is one that is about to be sent.
 *
 * Nothing here throws: transport and token failures come back as
 * `provider_error`, because the consumer branches on `outcome`.
 */

import type {
  FACILITATION_FOLLOWUP_SERVICE_NAME as CONTRACT_SERVICE_NAME,
  FacilitationFollowupService,
  FollowupGrant,
  FollowupRefusal,
  GrantVerificationService,
} from '@omadia/plugin-api';

import { truncate } from '../graphClient.js';
import { redactSecrets } from '../teamsProvisioner/redact.js';
import { checkAgainstGrant } from './addresses.js';
import type { OrganizerMailboxConfig } from './config.js';
import { createGrantGate } from './grantVerification.js';
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
 * Slot length asked of getSchedule (`availabilityViewInterval`, 5 to 1440).
 * Busy intervals are exact to this grid, counted from the start of the window.
 */
const AVAILABILITY_SLOT_MINUTES = 15;
const AVAILABILITY_SLOT_MS = AVAILABILITY_SLOT_MINUTES * 60 * 1000;
/**
 * `availabilityView` digits that do not block a slot: 0 free, 4 working
 * elsewhere. Everything else (1 tentative, 2 busy, 3 out of office, and any
 * value added later) counts as busy.
 */
const NOT_BUSY_SLOTS: ReadonlySet<string> = new Set(['0', '4']);
/** An instant with an explicit offset; a local time would be ambiguous. */
const INSTANT_PATTERN =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})$/;
/** Documented getSchedule limit: schedules per request. */
const MAX_SCHEDULES_PER_REQUEST = 20;
/** Documented getSchedule limit: the window has to stay below 62 days. */
const MAX_FREE_BUSY_WINDOW_DAYS = 62;
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const MAX_ERROR_DETAIL_LENGTH = 200;
const MAX_LOGGED_ID_LENGTH = 80;

type Interval = { readonly start: string; readonly end: string };
type Prepared<T> =
  | { readonly ready: true; readonly mailbox: string; readonly addresses: readonly string[]; readonly value: T }
  | { readonly ready: false; readonly refusal: FollowupRefusal };

export interface FacilitationFollowupOptions {
  readonly organizer: OrganizerMailboxConfig;
  readonly transport: FollowupGraphTransport;
  /**
   * The kernel's `grantVerification` service, looked up at every call.
   * Absent, or answering `undefined`: a kernel without kernel-issued grants.
   * A call that states a grant id is then refused; one without is not asked
   * about and runs on its allow-list alone.
   */
  readonly grantVerification?: () => GrantVerificationService | undefined;
  /** For tests: how long the kernel may take to answer. */
  readonly grantVerificationTimeoutMs?: number;
  /**
   * Audit line per call: operation, facilitation, counts, outcome and what
   * the kernel said about the grant. No addresses, no content.
   */
  readonly log?: (msg: string) => void;
}

export function createFacilitationFollowupService(
  opts: FacilitationFollowupOptions,
): FacilitationFollowupService {
  const grantGate = createGrantGate({
    ...(opts.grantVerification ? { lookup: opts.grantVerification } : {}),
    ...(opts.grantVerificationTimeoutMs !== undefined
      ? { timeoutMs: opts.grantVerificationTimeoutMs }
      : {}),
  });

  const audit = (
    operation: string,
    grant: FollowupGrant,
    verification: string,
    people: number,
    outcome: string,
  ): void => {
    opts.log?.(
      `[microsoft365] facilitation.followup op=${operation} facilitation=${loggableId(grant?.facilitationId)} people=${String(people)} outcome=${outcome} grant=${verification}`,
    );
  };

  /** `grant=` of a call that this service's own checks ended before the kernel was asked. */
  const beforeKernel = (grant: FollowupGrant): string =>
    (grant as { grantId?: unknown } | undefined)?.grantId === undefined
      ? 'not_asked:no_grant_id'
      : 'not_asked:refused_locally';

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
        toUtcWindow(request.from, request.to, 'from', 'to', MAX_FREE_BUSY_WINDOW_DAYS),
      );
      if (!prepared.ready) {
        audit('freeBusy', grant, beforeKernel(grant), countOf(request.attendees), prepared.refusal.code);
        return prepared.refusal;
      }
      const verified = await grantGate.check(grant, 'free_busy', prepared.addresses);
      if (verified.verdict === 'refused') {
        audit('freeBusy', grant, verified.audit, prepared.addresses.length, verified.refusal.code);
        return verified.refusal;
      }
      // One request per 20 schedules, one after the other. A failed request
      // fails the lookup (it has no side effect, so the caller can repeat it);
      // a single unresolvable schedule inside an answer does not.
      const busy: { attendee: string; intervals: Interval[] }[] = [];
      for (const schedules of chunk(prepared.addresses, MAX_SCHEDULES_PER_REQUEST)) {
        const response = await post(
          'getSchedule',
          prepared.mailbox,
          {
            schedules,
            startTime: { dateTime: prepared.value.start, timeZone: UTC },
            endTime: { dateTime: prepared.value.end, timeZone: UTC },
            availabilityViewInterval: AVAILABILITY_SLOT_MINUTES,
          },
          PREFER_UTC_HEADER,
        );
        if ('outcome' in response) {
          audit('freeBusy', grant, verified.audit, prepared.addresses.length, response.code);
          return response;
        }
        busy.push(
          ...readBusy(
            response.json,
            schedules,
            utcMs(prepared.value.start),
            utcMs(prepared.value.end),
          ),
        );
      }
      audit(
        'freeBusy',
        grant,
        verified.audit,
        prepared.addresses.length,
        busy.length === prepared.addresses.length
          ? 'ok'
          : `ok_unknown=${String(prepared.addresses.length - busy.length)}`,
      );
      return { outcome: 'ok', busy };
    },

    async createInvitation(grant, request) {
      const prepared = prepare(request.attendees, grant, () => {
        if (!isNonEmptyString(request.subject)) return 'subject is empty';
        if (typeof request.body !== 'string') return 'body is not a string';
        return toUtcWindow(request.start, request.end, 'start', 'end');
      });
      if (!prepared.ready) {
        audit('createInvitation', grant, beforeKernel(grant), countOf(request.attendees), prepared.refusal.code);
        return prepared.refusal;
      }
      const verified = await grantGate.check(grant, 'create_invitation', prepared.addresses);
      if (verified.verdict === 'refused') {
        audit('createInvitation', grant, verified.audit, prepared.addresses.length, verified.refusal.code);
        return verified.refusal;
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
        audit('createInvitation', grant, verified.audit, prepared.addresses.length, response.code);
        return response;
      }
      // Only the id and the join link of the event just created are read.
      const created = asRecord(response.json);
      const eventId = created?.['id'];
      if (!isNonEmptyString(eventId)) {
        audit('createInvitation', grant, verified.audit, prepared.addresses.length, 'provider_error');
        return refuse('provider_error', 'graph createEvent answered without an event id');
      }
      const joinUrl = asRecord(created?.['onlineMeeting'])?.['joinUrl'];
      // An online meeting was asked for and Graph created the event without a
      // link. The contract has no code for that: the event exists, so this is
      // `created` without `joinUrl`, and the audit line names it.
      const isLinkMissing = request.online === true && !isNonEmptyString(joinUrl);
      audit(
        'createInvitation',
        grant,
        verified.audit,
        prepared.addresses.length,
        isLinkMissing ? 'created_without_join_url' : 'created',
      );
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
        audit('sendMail', grant, beforeKernel(grant), countOf(request.recipients), prepared.refusal.code);
        return prepared.refusal;
      }
      const verified = await grantGate.check(grant, 'send_mail', prepared.addresses);
      if (verified.verdict === 'refused') {
        audit('sendMail', grant, verified.audit, prepared.addresses.length, verified.refusal.code);
        return verified.refusal;
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
        audit('sendMail', grant, verified.audit, prepared.addresses.length, response.code);
        return response;
      }
      audit('sendMail', grant, verified.audit, prepared.addresses.length, 'sent');
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
  maxDays?: number,
): { start: string; end: string } | string {
  const start = parseInstant(startInput);
  if (start === undefined) return `${startName} is not an ISO 8601 instant with an offset`;
  const end = parseInstant(endInput);
  if (end === undefined) return `${endName} is not an ISO 8601 instant with an offset`;
  if (end <= start) return `${endName} is not after ${startName}`;
  if (maxDays !== undefined && end - start >= maxDays * MS_PER_DAY) {
    return `the window from ${startName} to ${endName} has to be shorter than ${String(maxDays)} days`;
  }
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
 * PRIVACY. A schedule's `scheduleItems` can carry `subject` and `location`,
 * depending on how far that person shares their calendar with the organiser
 * (seen live in the byte5 tenant, 2026-11-18). This function therefore does
 * not look at `scheduleItems` at all. Of each schedule it reads three things:
 * `scheduleId`, whether there is an `error`, and `availabilityView` — a string
 * with one digit per slot that cannot hold any detail of an appointment.
 *
 * An attendee whose schedule Graph could not resolve (an `error` entry, no
 * entry, or no availability view) is left out: missing from the result means
 * "availability unknown", which is not the same as an empty `intervals` list.
 */
function readBusy(
  json: unknown,
  attendees: readonly string[],
  windowStartMs: number,
  windowEndMs: number,
): { attendee: string; intervals: Interval[] }[] {
  const entries = asRecord(json)?.['value'];
  if (!Array.isArray(entries)) return [];
  const requested = new Set(attendees);
  const result: { attendee: string; intervals: Interval[] }[] = [];
  for (const entry of entries) {
    const schedule = asRecord(entry);
    const scheduleId = schedule?.['scheduleId'];
    const view = schedule?.['availabilityView'];
    if (typeof scheduleId !== 'string' || schedule?.['error'] != null) continue;
    if (typeof view !== 'string' || view.length === 0) continue;
    const attendee = scheduleId.trim().toLowerCase();
    if (!requested.delete(attendee)) continue;
    result.push({ attendee, intervals: busyIntervals(view, windowStartMs, windowEndMs) });
  }
  return result;
}

/**
 * Turn an availability view into busy intervals. Slot `i` covers
 * `[start + i * slot, start + (i + 1) * slot)`, cut off at the end of the
 * window; neighbouring busy slots become one interval, so the result shows
 * neither how many appointments there are nor their exact bounds.
 */
function busyIntervals(view: string, windowStartMs: number, windowEndMs: number): Interval[] {
  const intervals: Interval[] = [];
  let runStart: number | undefined;
  const close = (endMs: number): void => {
    if (runStart === undefined) return;
    intervals.push({
      start: new Date(runStart).toISOString(),
      end: new Date(Math.min(endMs, windowEndMs)).toISOString(),
    });
    runStart = undefined;
  };
  for (let i = 0; i < view.length; i += 1) {
    const slotStart = windowStartMs + i * AVAILABILITY_SLOT_MS;
    if (slotStart >= windowEndMs) break;
    if (NOT_BUSY_SLOTS.has(view.charAt(i))) close(slotStart);
    else runStart ??= slotStart;
  }
  close(windowStartMs + view.length * AVAILABILITY_SLOT_MS);
  return intervals;
}

/** A Graph `dateTime` this service produced itself (`toGraphDateTime`), in UTC. */
function utcMs(graphDateTime: string): number {
  return Date.parse(`${graphDateTime}Z`);
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

function chunk<T>(items: readonly T[], size: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < items.length; i += size) chunks.push(items.slice(i, i + size));
  return chunks;
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

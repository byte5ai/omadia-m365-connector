import { Client } from '@microsoft/microsoft-graph-client';

/**
 * Graph Calendar wrapper. Every call takes the per-user access token
 * explicitly — the client is stateless and safe to share across turns
 * and users. Token acquisition lives in `graphObo.ts`.
 */

export type AttendeeType = 'required' | 'optional' | 'resource';

export interface FindSlotsOptions {
  /** Access token from `GraphOboClient.acquireTokenForUser`. */
  accessToken: string;
  /** Email/UPN list. Caller controls required vs optional. */
  attendees: Array<{ email: string; type?: AttendeeType }>;
  /** Meeting length in minutes. Graph accepts 15–1440. */
  durationMinutes: number;
  /** ISO-8601 with offset. Default: now. */
  windowStart?: string;
  /** ISO-8601. Default: windowStart + 5 days. */
  windowEnd?: string;
  /** Max candidates Graph should return (1..12). Default 5. */
  maxCandidates?: number;
  /** IANA TZ. Default: UTC. */
  timeZone?: string;
  /**
   * 0–100. Below this, Graph drops the suggestion. Default 100 (everyone
   * required is free); drop to 50 if we want "most of the room" slots.
   */
  minimumAttendeePercentage?: number;
}

export interface MeetingSlotSuggestion {
  start: string;
  end: string;
  timeZone: string;
  /** Graph's own 0–100 ranking. Surfaced to the LLM for picking top-N. */
  confidence: number;
  /** Per-attendee status: `free | tentative | busy | oof | workingElsewhere | unknown`. */
  attendees: Array<{ email: string; availability: string }>;
  /** Graph's optional text rationale (e.g. "suggested because …"). */
  reason?: string;
}

export interface GetScheduleOptions {
  accessToken: string;
  /** Email/UPN list — includes self. */
  users: string[];
  /** ISO-8601 with offset. */
  windowStart: string;
  windowEnd: string;
  /** In minutes. Default 30. */
  availabilityViewInterval?: number;
  timeZone?: string;
}

export interface ScheduleEntry {
  user: string;
  /** Dense string: `0=free, 1=tentative, 2=busy, 3=oof, 4=workingElsewhere`. */
  availabilityView: string;
  /** Optional detailed busy windows Graph echoes back. */
  busy: Array<{ start: string; end: string; status: string }>;
}

export interface CreateEventOptions {
  accessToken: string;
  subject: string;
  /** ISO-8601 — must match a slot from `findMeetingTimes` when booking. */
  start: string;
  end: string;
  timeZone?: string;
  attendees: Array<{ email: string; type?: AttendeeType }>;
  bodyHtml?: string;
  location?: string;
  /** When true, Graph issues a Teams meeting link and embeds it. */
  createTeamsMeeting?: boolean;
}

export interface CreatedEvent {
  id: string;
  webLink: string;
  onlineMeetingJoinUrl?: string;
  subject: string;
  start: string;
  end: string;
}

export interface MailboxSettings {
  timeZone: string;
  workingHoursStart: string;
  workingHoursEnd: string;
  workingDays: string[];
}

/**
 * Build a short-lived Graph client bound to a single user access token.
 * The SDK's default middleware stack handles throttling + 5xx retry with
 * sane backoff; we re-create one per call to avoid token-leak across
 * concurrent users.
 */
function clientForToken(token: string): Client {
  return Client.init({
    authProvider: (done) => done(null, token),
    defaultVersion: 'v1.0',
  });
}

export class GraphCalendarClient {
  async findMeetingTimes(opts: FindSlotsOptions): Promise<MeetingSlotSuggestion[]> {
    const client = clientForToken(opts.accessToken);
    const tz = opts.timeZone ?? 'UTC';
    const windowStart = opts.windowStart ?? new Date().toISOString();
    const windowEnd =
      opts.windowEnd ?? new Date(Date.parse(windowStart) + 5 * 24 * 3600 * 1000).toISOString();

    const body = {
      attendees: opts.attendees.map((a) => ({
        type: a.type ?? 'required',
        emailAddress: { address: a.email },
      })),
      timeConstraint: {
        timeSlots: [
          {
            start: { dateTime: windowStart, timeZone: tz },
            end: { dateTime: windowEnd, timeZone: tz },
          },
        ],
      },
      meetingDuration: minutesToIsoDuration(opts.durationMinutes),
      maxCandidates: opts.maxCandidates ?? 5,
      minimumAttendeePercentage: opts.minimumAttendeePercentage ?? 100,
      isOrganizerOptional: false,
      returnSuggestionReasons: true,
    };

    const result = (await client.api('/me/findMeetingTimes').post(body)) as {
      meetingTimeSuggestions?: Array<{
        meetingTimeSlot: {
          start: { dateTime: string; timeZone: string };
          end: { dateTime: string; timeZone: string };
        };
        confidence: number;
        attendeeAvailability?: Array<{
          attendee: { emailAddress: { address: string } };
          availability: string;
        }>;
        suggestionReason?: string;
      }>;
    };

    const suggestions = result.meetingTimeSuggestions ?? [];
    return suggestions.map((s) => ({
      start: s.meetingTimeSlot.start.dateTime,
      end: s.meetingTimeSlot.end.dateTime,
      timeZone: s.meetingTimeSlot.start.timeZone,
      confidence: s.confidence,
      attendees: (s.attendeeAvailability ?? []).map((a) => ({
        email: a.attendee.emailAddress.address,
        availability: a.availability,
      })),
      ...(s.suggestionReason ? { reason: s.suggestionReason } : {}),
    }));
  }

  async getSchedule(opts: GetScheduleOptions): Promise<ScheduleEntry[]> {
    const client = clientForToken(opts.accessToken);
    const tz = opts.timeZone ?? 'UTC';
    const body = {
      schedules: opts.users,
      startTime: { dateTime: opts.windowStart, timeZone: tz },
      endTime: { dateTime: opts.windowEnd, timeZone: tz },
      availabilityViewInterval: opts.availabilityViewInterval ?? 30,
    };

    const result = (await client.api('/me/calendar/getSchedule').post(body)) as {
      value?: Array<{
        scheduleId: string;
        availabilityView: string;
        scheduleItems?: Array<{
          start: { dateTime: string };
          end: { dateTime: string };
          status: string;
        }>;
      }>;
    };

    return (result.value ?? []).map((v) => ({
      user: v.scheduleId,
      availabilityView: v.availabilityView,
      busy: (v.scheduleItems ?? []).map((i) => ({
        start: i.start.dateTime,
        end: i.end.dateTime,
        status: i.status,
      })),
    }));
  }

  async createEvent(opts: CreateEventOptions): Promise<CreatedEvent> {
    const client = clientForToken(opts.accessToken);
    const tz = opts.timeZone ?? 'UTC';

    const body: Record<string, unknown> = {
      subject: opts.subject,
      start: { dateTime: opts.start, timeZone: tz },
      end: { dateTime: opts.end, timeZone: tz },
      attendees: opts.attendees.map((a) => ({
        type: a.type ?? 'required',
        emailAddress: { address: a.email },
      })),
    };
    if (opts.bodyHtml) {
      body.body = { contentType: 'HTML', content: opts.bodyHtml };
    }
    if (opts.location) {
      body.location = { displayName: opts.location };
    }
    if (opts.createTeamsMeeting) {
      body.isOnlineMeeting = true;
      body.onlineMeetingProvider = 'teamsForBusiness';
    }

    const created = (await client.api('/me/events').post(body)) as {
      id: string;
      webLink: string;
      subject: string;
      start: { dateTime: string };
      end: { dateTime: string };
      onlineMeeting?: { joinUrl?: string };
    };

    return {
      id: created.id,
      webLink: created.webLink,
      subject: created.subject,
      start: created.start.dateTime,
      end: created.end.dateTime,
      ...(created.onlineMeeting?.joinUrl ? { onlineMeetingJoinUrl: created.onlineMeeting.joinUrl } : {}),
    };
  }

  /**
   * Resolve the authenticated user's primary address from the access-token
   * JWT itself — no Graph call needed, so this works even when the token
   * doesn't carry `User.Read` (Calendar-only scopes don't grant `/me`
   * profile access and would 403 on any Graph `/me` query).
   *
   * Fallback order: `upn` → `preferred_username` → `unique_name` →
   * `/me` Graph call as last resort (may 403).
   */
  async getSelfAddress(accessToken: string): Promise<string> {
    const jwt = extractUpnFromToken(accessToken);
    if (jwt) return jwt;
    const client = clientForToken(accessToken);
    const result = (await client.api('/me').select('mail,userPrincipalName').get()) as {
      mail?: string;
      userPrincipalName?: string;
    };
    const addr = result.mail ?? result.userPrincipalName;
    if (!addr) throw new Error('GET /me returned no mail or userPrincipalName');
    return addr;
  }

  async getMailboxSettings(accessToken: string): Promise<MailboxSettings> {
    const client = clientForToken(accessToken);
    const result = (await client.api('/me/mailboxSettings').get()) as {
      timeZone?: string;
      workingHours?: {
        startTime: string;
        endTime: string;
        daysOfWeek: string[];
      };
    };
    return {
      timeZone: result.timeZone ?? 'UTC',
      workingHoursStart: result.workingHours?.startTime ?? '09:00:00.0000000',
      workingHoursEnd: result.workingHours?.endTime ?? '17:00:00.0000000',
      workingDays: result.workingHours?.daysOfWeek ?? [
        'monday',
        'tuesday',
        'wednesday',
        'thursday',
        'friday',
      ],
    };
  }
}

/**
 * Extract the user's UPN from an AAD access-token JWT without a Graph call.
 * Works for any AAD-issued token regardless of granted scopes — the claims
 * are in the token itself, independent of API permissions.
 * Returns `undefined` on any decode failure so callers can fall back to `/me`.
 */
function extractUpnFromToken(accessToken: string): string | undefined {
  try {
    const parts = accessToken.split('.');
    if (parts.length < 2) return undefined;
    const payloadPart = parts[1];
    if (!payloadPart) return undefined;
    // Base64url → base64 (padding for Node's Buffer strictness).
    const b64 = payloadPart.replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
    const json = Buffer.from(padded, 'base64').toString('utf-8');
    const payload = JSON.parse(json) as {
      upn?: string;
      preferred_username?: string;
      unique_name?: string;
      email?: string;
    };
    const addr =
      payload.upn ??
      payload.preferred_username ??
      payload.unique_name ??
      payload.email;
    if (!addr || typeof addr !== 'string') return undefined;
    // preferred_username sometimes contains a display name rather than an
    // address — reject anything without an `@`.
    if (!addr.includes('@')) return undefined;
    return addr;
  } catch {
    return undefined;
  }
}

/** `45` → `PT45M`, `90` → `PT1H30M`. Graph accepts ISO-8601 durations only. */
function minutesToIsoDuration(minutes: number): string {
  if (!Number.isFinite(minutes) || minutes <= 0) {
    throw new Error(`invalid duration: ${minutes}`);
  }
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (h === 0) return `PT${m}M`;
  if (m === 0) return `PT${h}H`;
  return `PT${h}H${m}M`;
}

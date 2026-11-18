import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  FACILITATION_FOLLOWUP_CAPABILITY,
  FACILITATION_FOLLOWUP_SERVICE_NAME,
  FACILITATION_ORGANIZER_MAILBOX_FIELD,
  createAppOnlyGraphTransport,
  createFacilitationFollowupService,
  readOrganizerMailbox,
  type OrganizerMailboxConfig,
} from '../src/facilitationFollowup/index.js';
import { checkAgainstGrant } from '../src/facilitationFollowup/addresses.js';

// The facilitation follow-up service against a FAKE Graph. There is no test
// tenant for this feature: the fake sits at the `fetch` seam, so the real
// transport, its URLs and its headers are what gets exercised, and the request
// and response shapes follow the Graph v1.0 reference for
//
//   POST /users/{id}/calendar/getSchedule
//   POST /users/{id}/events
//   POST /users/{id}/sendMail
//
// What this suite must prove:
//   * each operation sends exactly the documented request and maps the answer;
//   * a single address outside the grant refuses the WHOLE call, names the
//     rejected addresses and reaches neither Graph nor the token endpoint;
//   * without an organiser mailbox the answer is `not_configured`;
//   * no request this service can make reads message or event contents.

const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const ORGANIZER = 'facilitator@contoso.example';
const ANNA = 'anna@contoso.example';
const BEN = 'ben@contoso.example';
const OUTSIDER = 'mallory@elsewhere.example';
const GRANT = { facilitationId: 'fac-42', allowed: [ANNA, BEN] };
const GRAPH = 'https://graph.microsoft.com/v1.0';
const ORGANIZER_PATH = `${GRAPH}/users/${encodeURIComponent(ORGANIZER)}`;

interface GraphCall {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: unknown;
}

interface FakeGraph {
  fetchImpl: typeof fetch;
  /** Calls to graph.microsoft.com, in order. */
  graphCalls: GraphCall[];
  /** Calls to the token endpoint. */
  tokenCalls: string[];
}

type Responder = (call: GraphCall) => { status: number; body?: unknown };

function fakeGraph(respond: Responder): FakeGraph {
  const graphCalls: GraphCall[] = [];
  const tokenCalls: string[] = [];
  const fetchImpl = (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    const url = String(input);
    if (url.startsWith('https://login.microsoftonline.com/')) {
      tokenCalls.push(url);
      return new Response(JSON.stringify({ access_token: 'app-token', expires_in: 3600 }), {
        status: 200,
      });
    }
    const call: GraphCall = {
      url,
      method: init?.method ?? 'GET',
      headers: { ...(init?.headers as Record<string, string> | undefined) },
      body: typeof init?.body === 'string' ? (JSON.parse(init.body) as unknown) : undefined,
    };
    graphCalls.push(call);
    const { status, body } = respond(call);
    return new Response(body === undefined ? null : JSON.stringify(body), { status });
  }) as typeof fetch;
  return { fetchImpl, graphCalls, tokenCalls };
}

function serviceOn(
  graph: FakeGraph,
  organizer: OrganizerMailboxConfig = { kind: 'configured', mailbox: ORGANIZER },
  logs: string[] = [],
) {
  return createFacilitationFollowupService({
    organizer,
    transport: createAppOnlyGraphTransport({
      credential: { tenantId: 'tenant', clientId: 'client', clientSecret: 'secret' },
      fetchImpl: graph.fetchImpl,
    }),
    log: (msg) => logs.push(msg),
  });
}

const failIfCalled: Responder = (call) => {
  throw new Error(`unexpected Graph request ${call.method} ${call.url}`);
};

/** A getSchedule answer as documented, including the fields we must ignore. */
const scheduleAnswer = {
  value: [
    {
      scheduleId: ANNA,
      availabilityView: '002200',
      scheduleItems: [
        {
          isPrivate: false,
          status: 'busy',
          subject: 'Salary review',
          location: 'Room 4',
          start: { dateTime: '2026-11-20T09:00:00.0000000', timeZone: 'UTC' },
          end: { dateTime: '2026-11-20T10:00:00.0000000', timeZone: 'UTC' },
        },
        {
          status: 'tentative',
          subject: 'Overlapping call',
          start: { dateTime: '2026-11-20T09:30:00.0000000', timeZone: 'UTC' },
          end: { dateTime: '2026-11-20T10:30:00.0000000', timeZone: 'UTC' },
        },
        {
          status: 'free',
          subject: 'Lunch reminder',
          start: { dateTime: '2026-11-20T12:00:00.0000000', timeZone: 'UTC' },
          end: { dateTime: '2026-11-20T13:00:00.0000000', timeZone: 'UTC' },
        },
        {
          status: 'oof',
          start: { dateTime: '2026-11-20T15:00:00.0000000', timeZone: 'UTC' },
          end: { dateTime: '2026-11-20T16:00:00.0000000', timeZone: 'UTC' },
        },
      ],
      workingHours: { daysOfWeek: ['monday'] },
    },
    { scheduleId: BEN, availabilityView: '000000', scheduleItems: [] },
  ],
};

describe('service identity', () => {
  it('uses the contract service name and a versioned capability', () => {
    assert.equal(FACILITATION_FOLLOWUP_SERVICE_NAME, 'facilitation.followup');
    assert.equal(FACILITATION_FOLLOWUP_CAPABILITY, 'facilitation.followup@1');
  });

  it('the manifest announces the capability and the setup field', () => {
    const manifest = readFileSync(join(pkgRoot, 'manifest.yaml'), 'utf8');
    assert.ok(manifest.includes(`- "${FACILITATION_FOLLOWUP_CAPABILITY}"`));
    assert.ok(
      !manifest.includes(`service: "${FACILITATION_FOLLOWUP_SERVICE_NAME}"`),
      'not offered to builder-generated agents through service_types',
    );
    assert.match(
      manifest,
      new RegExp(`key: "${FACILITATION_ORGANIZER_MAILBOX_FIELD}"[\\s\\S]*?required: false`),
    );
  });
});

describe('freeBusy', () => {
  it('asks getSchedule on the organiser mailbox, in UTC, for the granted attendees', async () => {
    const graph = fakeGraph(() => ({ status: 200, body: scheduleAnswer }));
    await serviceOn(graph).freeBusy(GRANT, {
      attendees: [` ${ANNA.toUpperCase()} `, BEN, ANNA],
      from: '2026-11-20T09:00:00+01:00',
      to: '2026-11-20T18:00:00+01:00',
    });

    assert.equal(graph.graphCalls.length, 1);
    const call = graph.graphCalls[0]!;
    assert.equal(call.method, 'POST');
    assert.equal(call.url, `${ORGANIZER_PATH}/calendar/getSchedule`);
    assert.equal(call.headers['Authorization'], 'Bearer app-token');
    assert.equal(call.headers['Prefer'], 'outlook.timezone="UTC"');
    assert.deepEqual(call.body, {
      schedules: [ANNA, BEN],
      startTime: { dateTime: '2026-11-20T08:00:00', timeZone: 'UTC' },
      endTime: { dateTime: '2026-11-20T17:00:00', timeZone: 'UTC' },
    });
  });

  it('returns merged busy intervals only, with no event details', async () => {
    const graph = fakeGraph(() => ({ status: 200, body: scheduleAnswer }));
    const result = await serviceOn(graph).freeBusy(GRANT, {
      attendees: [ANNA, BEN],
      from: '2026-11-20T08:00:00Z',
      to: '2026-11-20T17:00:00Z',
    });

    assert.deepEqual(result, {
      outcome: 'ok',
      busy: [
        {
          attendee: ANNA,
          intervals: [
            { start: '2026-11-20T09:00:00.000Z', end: '2026-11-20T10:30:00.000Z' },
            { start: '2026-11-20T15:00:00.000Z', end: '2026-11-20T16:00:00.000Z' },
          ],
        },
        { attendee: BEN, intervals: [] },
      ],
    });
    const serialised = JSON.stringify(result);
    for (const leaked of ['Salary', 'Room 4', 'Overlapping', 'Lunch', 'subject', 'location']) {
      assert.ok(!serialised.includes(leaked), `result must not contain "${leaked}"`);
    }
  });

  it('leaves out an attendee whose schedule Graph could not resolve', async () => {
    const graph = fakeGraph(() => ({
      status: 200,
      body: {
        value: [
          { scheduleId: ANNA, scheduleItems: [] },
          {
            scheduleId: BEN,
            error: { message: 'Unable to resolve the recipient', responseCode: 'ErrorMailRecipientNotFound' },
          },
          { scheduleId: OUTSIDER, scheduleItems: [] },
        ],
      },
    }));
    const result = await serviceOn(graph).freeBusy(GRANT, {
      attendees: [ANNA, BEN],
      from: '2026-11-20T08:00:00Z',
      to: '2026-11-20T17:00:00Z',
    });
    assert.deepEqual(result, { outcome: 'ok', busy: [{ attendee: ANNA, intervals: [] }] });
  });

  it('refuses a window without an offset or with the end before the start, without a request', async () => {
    const graph = fakeGraph(failIfCalled);
    const service = serviceOn(graph);
    const local = await service.freeBusy(GRANT, {
      attendees: [ANNA],
      from: '2026-11-20T08:00:00',
      to: '2026-11-20T17:00:00Z',
    });
    const reversed = await service.freeBusy(GRANT, {
      attendees: [ANNA],
      from: '2026-11-20T17:00:00Z',
      to: '2026-11-20T08:00:00Z',
    });
    for (const result of [local, reversed]) {
      assert.equal(result.outcome, 'refused');
      assert.equal(result.outcome === 'refused' && result.code, 'provider_error');
    }
    assert.equal(graph.graphCalls.length + graph.tokenCalls.length, 0);
  });
});

describe('createInvitation', () => {
  const request = {
    attendees: [ANNA, BEN],
    subject: 'Follow-up: pricing decision',
    start: '2026-11-27T10:00:00+01:00',
    end: '2026-11-27T11:00:00+01:00',
    body: 'Agenda:\n1. Open points\nProtocol: https://omadia.example/p/42',
  };

  it('creates the event on the organiser calendar with the granted attendees', async () => {
    const graph = fakeGraph(() => ({ status: 201, body: { id: 'AAMkAGI=', onlineMeeting: null } }));
    const result = await serviceOn(graph).createInvitation(GRANT, request);

    assert.deepEqual(result, { outcome: 'created', eventId: 'AAMkAGI=' });
    assert.equal(graph.graphCalls.length, 1);
    const call = graph.graphCalls[0]!;
    assert.equal(call.method, 'POST');
    assert.equal(call.url, `${ORGANIZER_PATH}/events`);
    assert.deepEqual(call.body, {
      subject: request.subject,
      body: { contentType: 'text', content: request.body },
      start: { dateTime: '2026-11-27T09:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-11-27T10:00:00', timeZone: 'UTC' },
      attendees: [
        { emailAddress: { address: ANNA }, type: 'required' },
        { emailAddress: { address: BEN }, type: 'required' },
      ],
    });
  });

  it('asks for a Teams meeting when online and returns the join link', async () => {
    const graph = fakeGraph(() => ({
      status: 201,
      body: {
        id: 'AAMkAGI=',
        isOnlineMeeting: true,
        onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc' },
      },
    }));
    const result = await serviceOn(graph).createInvitation(GRANT, { ...request, online: true });

    assert.deepEqual(result, {
      outcome: 'created',
      eventId: 'AAMkAGI=',
      joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc',
    });
    const body = graph.graphCalls[0]!.body as Record<string, unknown>;
    assert.equal(body['isOnlineMeeting'], true);
    assert.equal(body['onlineMeetingProvider'], 'teamsForBusiness');
  });

  it('reports a Graph refusal as provider_error with status and code', async () => {
    const graph = fakeGraph(() => ({
      status: 403,
      body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } },
    }));
    const result = await serviceOn(graph).createInvitation(GRANT, request);
    assert.equal(result.outcome, 'refused');
    if (result.outcome !== 'refused') return;
    assert.equal(result.code, 'provider_error');
    assert.match(result.message, /createEvent 403 code=ErrorAccessDenied/);
  });

  it('refuses an empty subject without a request', async () => {
    const graph = fakeGraph(failIfCalled);
    const result = await serviceOn(graph).createInvitation(GRANT, { ...request, subject: '  ' });
    assert.equal(result.outcome === 'refused' && result.code, 'provider_error');
    assert.equal(graph.graphCalls.length, 0);
  });
});

describe('sendMail', () => {
  const request = {
    recipients: [ANNA, BEN],
    subject: 'Protocol: pricing decision',
    body: 'Decisions:\n- Tier B stays.\nNext steps:\n- Ben drafts the note by Friday.',
  };

  it('sends the given content from the organiser mailbox to the granted recipients', async () => {
    const graph = fakeGraph(() => ({ status: 202 }));
    const result = await serviceOn(graph).sendMail(GRANT, request);

    assert.deepEqual(result, { outcome: 'sent' });
    assert.equal(graph.graphCalls.length, 1);
    const call = graph.graphCalls[0]!;
    assert.equal(call.method, 'POST');
    assert.equal(call.url, `${ORGANIZER_PATH}/sendMail`);
    assert.deepEqual(call.body, {
      message: {
        subject: request.subject,
        body: { contentType: 'text', content: request.body },
        toRecipients: [{ emailAddress: { address: ANNA } }, { emailAddress: { address: BEN } }],
      },
      saveToSentItems: true,
    });
  });

  it('reports a transport failure as provider_error instead of throwing', async () => {
    const graph = fakeGraph(() => {
      throw new Error('socket hang up');
    });
    const result = await serviceOn(graph).sendMail(GRANT, request);
    assert.equal(result.outcome, 'refused');
    if (result.outcome !== 'refused') return;
    assert.equal(result.code, 'provider_error');
    assert.match(result.message, /socket hang up/);
  });

  it('reports a failed token request as provider_error without leaking the secret', async () => {
    const fetchImpl = (async () =>
      new Response(JSON.stringify({ error: 'invalid_client', client_secret: 'secret' }), {
        status: 401,
      })) as typeof fetch;
    const service = createFacilitationFollowupService({
      organizer: { kind: 'configured', mailbox: ORGANIZER },
      transport: createAppOnlyGraphTransport({
        credential: { tenantId: 'tenant', clientId: 'client', clientSecret: 'secret' },
        fetchImpl,
      }),
    });
    const result = await service.sendMail(GRANT, request);
    assert.equal(result.outcome, 'refused');
    if (result.outcome !== 'refused') return;
    assert.equal(result.code, 'provider_error');
    assert.ok(!result.message.includes('"secret"'));
  });
});

describe('refusal outside the grant', () => {
  const window = { from: '2026-11-20T08:00:00Z', to: '2026-11-20T17:00:00Z' };
  const invitation = {
    subject: 'Follow-up',
    start: '2026-11-27T09:00:00Z',
    end: '2026-11-27T10:00:00Z',
    body: 'Agenda',
  };
  const mail = { subject: 'Protocol', body: 'Decisions' };

  it('every operation refuses the whole call, names the outsider and makes no request', async () => {
    const graph = fakeGraph(failIfCalled);
    const service = serviceOn(graph);
    const results = [
      await service.freeBusy(GRANT, { attendees: [ANNA, OUTSIDER], ...window }),
      await service.createInvitation(GRANT, { attendees: [ANNA, OUTSIDER], ...invitation }),
      await service.sendMail(GRANT, { recipients: [ANNA, OUTSIDER], ...mail }),
    ];
    for (const result of results) {
      assert.deepEqual(
        { outcome: result.outcome, code: 'code' in result && result.code, rejected: 'rejected' in result && result.rejected },
        { outcome: 'refused', code: 'recipient_not_allowed', rejected: [OUTSIDER] },
      );
    }
    assert.equal(graph.graphCalls.length, 0, 'no Graph request');
    assert.equal(graph.tokenCalls.length, 0, 'not even a token request');
  });

  it('an empty grant allows nobody', async () => {
    const graph = fakeGraph(failIfCalled);
    const result = await serviceOn(graph).sendMail(
      { facilitationId: 'fac-42', allowed: [] },
      { recipients: [ANNA], ...mail },
    );
    assert.equal(result.outcome === 'refused' && result.code, 'recipient_not_allowed');
  });

  it('a string that smuggles a second recipient is rejected even if the grant lists it', () => {
    const smuggled = `${ANNA}; ${OUTSIDER}`;
    const named = `Anna <${ANNA}>`;
    const check = checkAgainstGrant(
      { facilitationId: 'fac-42', allowed: [ANNA, smuggled, named] },
      [smuggled, named, ANNA],
    );
    assert.deepEqual(check, { allowed: false, rejected: [smuggled, named] });
  });

  it('the grant is checked before the configuration', async () => {
    const graph = fakeGraph(failIfCalled);
    const service = serviceOn(graph, { kind: 'not-configured', reason: 'empty' });
    const result = await service.sendMail(GRANT, { recipients: [OUTSIDER], ...mail });
    assert.equal(result.outcome === 'refused' && result.code, 'recipient_not_allowed');
  });

  it('writes an audit line without addresses or content', async () => {
    const logs: string[] = [];
    const graph = fakeGraph(() => ({ status: 202 }));
    const service = serviceOn(graph, undefined, logs);
    await service.sendMail(GRANT, { recipients: [ANNA], ...mail });
    await service.sendMail(GRANT, { recipients: [OUTSIDER], ...mail });
    assert.deepEqual(logs, [
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=1 outcome=sent',
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=1 outcome=recipient_not_allowed',
    ]);
  });
});

describe('not_configured', () => {
  it('every operation answers not_configured without a request when no organiser mailbox is set', async () => {
    const graph = fakeGraph(failIfCalled);
    const organizer = readOrganizerMailbox({ get: () => undefined });
    assert.equal(organizer.kind, 'not-configured');
    const service = serviceOn(graph, organizer);
    const results = [
      await service.freeBusy(GRANT, {
        attendees: [ANNA],
        from: '2026-11-20T08:00:00Z',
        to: '2026-11-20T17:00:00Z',
      }),
      await service.createInvitation(GRANT, {
        attendees: [ANNA],
        subject: 'Follow-up',
        start: '2026-11-27T09:00:00Z',
        end: '2026-11-27T10:00:00Z',
        body: 'Agenda',
      }),
      await service.sendMail(GRANT, { recipients: [ANNA], subject: 'Protocol', body: 'Decisions' }),
    ];
    for (const result of results) {
      assert.equal(result.outcome, 'refused');
      assert.equal(result.outcome === 'refused' && result.code, 'not_configured');
    }
    assert.equal(graph.graphCalls.length + graph.tokenCalls.length, 0);
  });

  it('reads the mailbox from the setup field and invents no default', () => {
    const read = (value: unknown) =>
      readOrganizerMailbox({
        get: <T>(key: string) =>
          (key === FACILITATION_ORGANIZER_MAILBOX_FIELD ? value : undefined) as T | undefined,
      });
    assert.deepEqual(read(' Facilitator@Contoso.example '), { kind: 'configured', mailbox: ORGANIZER });
    assert.equal(read('').kind, 'not-configured');
    assert.equal(read('   ').kind, 'not-configured');
    assert.equal(read('not-an-address').kind, 'not-configured');
    assert.equal(read(42).kind, 'not-configured');
  });
});

describe('no request reads message or event contents', () => {
  it('a full run makes only the three POSTs on the organiser mailbox', async () => {
    const graph = fakeGraph((call) => {
      if (call.url.endsWith('/getSchedule')) return { status: 200, body: scheduleAnswer };
      if (call.url.endsWith('/events')) return { status: 201, body: { id: 'evt-1' } };
      return { status: 202 };
    });
    const service = serviceOn(graph);
    await service.freeBusy(GRANT, {
      attendees: [ANNA, BEN],
      from: '2026-11-20T08:00:00Z',
      to: '2026-11-20T17:00:00Z',
    });
    await service.createInvitation(GRANT, {
      attendees: [ANNA, BEN],
      subject: 'Follow-up',
      start: '2026-11-27T09:00:00Z',
      end: '2026-11-27T10:00:00Z',
      body: 'Agenda',
      online: true,
    });
    await service.sendMail(GRANT, { recipients: [ANNA, BEN], subject: 'Protocol', body: 'Decisions' });

    assert.deepEqual(
      graph.graphCalls.map((c) => `${c.method} ${c.url.slice(ORGANIZER_PATH.length)}`),
      ['POST /calendar/getSchedule', 'POST /events', 'POST /sendMail'],
    );
    for (const call of graph.graphCalls) {
      assert.ok(call.url.startsWith(`${ORGANIZER_PATH}/`), 'only the organiser mailbox is addressed');
      assert.ok(!call.url.includes('?'), 'no query, so no $expand or $select into contents');
    }
    assert.equal(graph.tokenCalls.length, 1, 'one app-only token, cached');
  });

  it('the service sources contain no route into messages, folders or stored events', () => {
    const dir = join(pkgRoot, 'src', 'facilitationFollowup');
    const sources = ['service.ts', 'transport.ts', 'addresses.ts', 'config.ts']
      .map((file) => readFileSync(join(dir, file), 'utf8'))
      .join('\n')
      // Comments name the forbidden routes in order to rule them out.
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .replace(/^\s*\/\/.*$/gm, '');
    for (const forbidden of [
      '/messages',
      '/mailFolders',
      '/calendarView',
      '/calendars',
      '/me/',
      "'GET'",
      "'PATCH'",
      "'DELETE'",
      'graphObo',
      'graphCalendarClient',
    ]) {
      assert.ok(!sources.includes(forbidden), `follow-up sources must not contain ${forbidden}`);
    }
  });
});

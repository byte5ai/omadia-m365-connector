import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  ProvisioningHttp,
  type ProvisioningHttpOptions,
  type ProvisioningRequest,
} from '../src/teamsProvisioner/http.js';
import {
  ConsentMissingError,
  ProvisioningRequestError,
  ProvisioningThrottledError,
} from '../src/teamsProvisioner/errors.js';
import { redactSecrets } from '../src/teamsProvisioner/redact.js';

// The single token+fetch choke point: dual-audience token plumbing
// (graph.microsoft.com vs management.azure.com, per-(tenant,clientId,scope)
// cache), the 403 → ConsentMissingError and 409 → conflict-signal mapping,
// Retry-After-honouring 429 backoff, and the ARM long-running poll mode.
// Mocked-fetch style follows omadia-channel-teams tests/teamsGraphResolver.

interface FetchCall {
  url: string;
  init?: RequestInit;
}

interface ResponseSpec {
  status: number;
  body?: unknown;
  headers?: Record<string, string>;
}

interface Route {
  match: string;
  responses: ResponseSpec[];
}

const route = (match: string, ...responses: ResponseSpec[]): Route => ({
  match,
  responses,
});

function makeResponse(spec: ResponseSpec): Response {
  const headerMap = new Map(
    Object.entries(spec.headers ?? {}).map(([k, v]) => [k.toLowerCase(), v]),
  );
  return {
    ok: spec.status >= 200 && spec.status < 300,
    status: spec.status,
    headers: { get: (name: string) => headerMap.get(name.toLowerCase()) ?? null },
    json: async () => {
      if (spec.body === undefined) throw new Error('no body');
      return spec.body;
    },
    text: async () => (spec.body === undefined ? '' : JSON.stringify(spec.body)),
  } as unknown as Response;
}

/**
 * Route table with per-route response QUEUES (the last entry repeats), so
 * retry/poll sequences (429 → 200, InProgress → Succeeded) are expressible.
 * The token route answers dynamically from the urlencoded request body so
 * dual-audience tests can assert which identity+scope produced which bearer.
 */
function mockFetch(routes: Route[], calls: FetchCall[]): typeof fetch {
  const queues = routes.map((r) => ({ ...r, next: 0 }));
  return (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, ...(init ? { init } : {}) });
    if (url.includes('login.microsoftonline.com')) {
      const params = new URLSearchParams(String(init?.body ?? ''));
      const scopeHost = new URL(params.get('scope') ?? 'https://none/').host;
      const access_token = `tok:${params.get('client_id') ?? '?'}:${scopeHost}`;
      return makeResponse({ status: 200, body: { access_token, expires_in: 3600 } });
    }
    for (const q of queues) {
      if (url.includes(q.match)) {
        const spec = q.responses[Math.min(q.next, q.responses.length - 1)];
        q.next += 1;
        if (!spec) break;
        return makeResponse(spec);
      }
    }
    return makeResponse({ status: 404, body: {} });
  }) as typeof fetch;
}

const GRAPH_CRED = {
  tenantId: 'tenant-1',
  clientId: 'app-graph',
  clientSecret: 'graph-secret-value',
};
const ARM_CRED = {
  tenantId: 'tenant-1',
  clientId: 'app-arm',
  clientSecret: 'arm-secret-value',
};

/** One test rig: recorded fetch calls + recorded backoff sleeps + the client. */
function harness(
  routes: Route[],
  overrides?: Partial<ProvisioningHttpOptions>,
): { http: ProvisioningHttp; calls: FetchCall[]; sleeps: number[] } {
  const calls: FetchCall[] = [];
  const sleeps: number[] = [];
  const http = new ProvisioningHttp({
    graphCredential: GRAPH_CRED,
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async (ms) => {
      sleeps.push(ms);
    },
    ...overrides,
  });
  return { http, calls, sleeps };
}

const tokenCalls = (calls: FetchCall[]): FetchCall[] =>
  calls.filter((c) => c.url.includes('login.microsoftonline.com'));

const bearerOf = (call: FetchCall): string | undefined =>
  (call.init?.headers as Record<string, string> | undefined)?.['Authorization'];

const rejection = async (p: Promise<unknown>): Promise<unknown> =>
  p.then(() => assert.fail('expected the request to reject'), (e: unknown) => e);

const CREATE_APP: ProvisioningRequest = {
  resource: 'graph',
  method: 'POST',
  url: 'https://graph.microsoft.com/v1.0/applications',
  step: 'applications.create',
  jsonBody: { displayName: 'HR Agent', uniqueName: 'omadia-agent-hr' },
  missingScopesOn403: ['Application.ReadWrite.OwnedBy'],
};

const BOT_URL =
  'https://management.azure.com/subscriptions/s/resourceGroups/rg/providers/Microsoft.BotService/botServices/omadia-agent-hr?api-version=2022-09-15';

const PUT_BOT: ProvisioningRequest = {
  resource: 'arm',
  method: 'PUT',
  url: BOT_URL,
  step: 'botServices.put',
  jsonBody: { location: 'global', properties: { msaAppId: 'app-graph' } },
  missingScopesOn403: ['Microsoft.BotService/botServices/write'],
  pollLongRunning: true,
};

const PUT_BOT_NO_POLL: ProvisioningRequest = { ...PUT_BOT, pollLongRunning: false };

describe('ProvisioningHttp token plumbing', () => {
  it('sends a Graph bearer, JSON body, and caches the token across calls', async () => {
    const { http, calls } = harness([
      route('/applications', { status: 201, body: { id: 'obj-1' } }),
    ]);
    const first = await http.request(CREATE_APP);
    const second = await http.request(CREATE_APP);

    assert.equal(first.kind, 'ok');
    assert.equal(first.status, 201);
    assert.deepEqual(first.json, { id: 'obj-1' });
    assert.equal(second.kind, 'ok');

    const tokens = tokenCalls(calls);
    assert.equal(tokens.length, 1, 'token fetched once, then cached');
    const tokenBody = String(tokens[0]?.init?.body ?? '');
    assert.ok(tokenBody.includes('grant_type=client_credentials'));
    assert.ok(tokenBody.includes('client_id=app-graph'));
    assert.ok(
      tokenBody.includes(
        `scope=${encodeURIComponent('https://graph.microsoft.com/.default')}`,
      ),
    );

    const apiCalls = calls.filter((c) => c.url.includes('/applications'));
    assert.equal(apiCalls.length, 2);
    assert.equal(bearerOf(apiCalls[0]!), 'Bearer tok:app-graph:graph.microsoft.com');
    const headers = apiCalls[0]?.init?.headers as Record<string, string>;
    assert.equal(headers['content-type'], 'application/json');
    assert.equal(apiCalls[0]?.init?.method, 'POST');
    assert.deepEqual(JSON.parse(String(apiCalls[0]?.init?.body)), {
      displayName: 'HR Agent',
      uniqueName: 'omadia-agent-hr',
    });
  });

  it('uses the ARM audience + dedicated service principal, cached separately from Graph', async () => {
    const { http, calls } = harness(
      [
        route('/applications', { status: 201, body: {} }),
        route('management.azure.com', { status: 200, body: {} }),
      ],
      { armCredential: ARM_CRED },
    );
    await http.request(CREATE_APP);
    await http.request(PUT_BOT_NO_POLL);
    await http.request(PUT_BOT_NO_POLL);

    const tokens = tokenCalls(calls);
    assert.equal(tokens.length, 2, 'one token per audience, ARM token reused');
    const armTokenBody = String(tokens[1]?.init?.body ?? '');
    assert.ok(armTokenBody.includes('client_id=app-arm'));
    assert.ok(
      armTokenBody.includes(
        `scope=${encodeURIComponent('https://management.azure.com/.default')}`,
      ),
    );
    const armCall = calls.find((c) => c.url.includes('management.azure.com'));
    assert.equal(bearerOf(armCall!), 'Bearer tok:app-arm:management.azure.com');
  });

  it('falls back to the Graph credential for ARM when no armCredential is set (reuse app)', async () => {
    const { http, calls } = harness([
      route('management.azure.com', { status: 200, body: {} }),
    ]);
    await http.request(PUT_BOT_NO_POLL);

    const tokens = tokenCalls(calls);
    assert.equal(tokens.length, 1);
    const body = String(tokens[0]?.init?.body ?? '');
    assert.ok(body.includes('client_id=app-graph'));
    assert.ok(
      body.includes(
        `scope=${encodeURIComponent('https://management.azure.com/.default')}`,
      ),
    );
  });
});

describe('ProvisioningHttp status mapping', () => {
  it('403 from Graph → ConsentMissingError with scope set, no secret/bearer leak', async () => {
    const { http } = harness([
      route('/applications', {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);
    const err = await rejection(http.request(CREATE_APP));
    assert.ok(err instanceof ConsentMissingError);
    assert.deepEqual(err.missingScopes, ['Application.ReadWrite.OwnedBy']);
    assert.equal(err.resource, 'graph');
    assert.equal(err.message, 'consent_missing');
    const cause = (err as { cause?: unknown }).cause;
    assert.ok(cause instanceof Error);
    assert.ok(cause.message.includes('applications.create'));
    assert.ok(cause.message.includes('Authorization_RequestDenied'));
    for (const text of [err.message, cause.message]) {
      assert.ok(!text.includes('graph-secret-value'), 'client secret must never leak');
      assert.ok(!text.includes('tok:'), 'bearer token must never leak');
    }
  });

  it('403 from ARM → ConsentMissingError with resource arm', async () => {
    const { http } = harness(
      [route('management.azure.com', { status: 403, body: {} })],
      { armCredential: ARM_CRED },
    );
    const err = await rejection(http.request(PUT_BOT_NO_POLL));
    assert.ok(err instanceof ConsentMissingError);
    assert.equal(err.resource, 'arm');
    assert.deepEqual(err.missingScopes, ['Microsoft.BotService/botServices/write']);
  });

  it('409 → conflict signal (idempotent already-exists), never an exception', async () => {
    const { http, calls } = harness([
      route('/appCatalogs/teamsApps', {
        status: 409,
        body: { error: { code: 'Conflict' } },
      }),
    ]);
    const result = await http.request({
      resource: 'graph',
      method: 'POST',
      url: 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps',
      step: 'appCatalogs.upload',
      rawBody: { bytes: new Uint8Array([80, 75, 3, 4]), contentType: 'application/zip' },
      missingScopesOn403: ['AppCatalog.ReadWrite.All'],
    });
    assert.equal(result.kind, 'conflict');
    assert.equal(result.status, 409);
    assert.deepEqual(result.json, { error: { code: 'Conflict' } });
    const upload = calls.find((c) => c.url.includes('/appCatalogs/teamsApps'));
    const headers = upload?.init?.headers as Record<string, string>;
    assert.equal(headers['content-type'], 'application/zip');
  });

  it('other non-2xx propagates verbatim as a bounded generic error', async () => {
    const { http } = harness([
      route('/applications', {
        status: 400,
        body: { error: { message: 'bad uniqueName' } },
      }),
    ]);
    const err = await rejection(http.request(CREATE_APP));
    assert.ok(err instanceof Error);
    assert.ok(!(err instanceof ConsentMissingError));
    assert.ok(err.message.includes('graph applications.create 400'));
    assert.ok(err.message.includes('bad uniqueName'));
    assert.ok(!err.message.includes('graph-secret-value'));
    assert.ok(!err.message.includes('tok:'));
  });

  it('extraOkStatuses lets DELETE rollbacks treat 404 as already-gone', async () => {
    const { http } = harness([route('/applications/obj-1', { status: 404 })]);
    const result = await http.request({
      resource: 'graph',
      method: 'DELETE',
      url: 'https://graph.microsoft.com/v1.0/applications/obj-1',
      step: 'applications.delete',
      missingScopesOn403: ['Application.ReadWrite.OwnedBy'],
      extraOkStatuses: [404],
    });
    assert.equal(result.kind, 'ok');
    assert.equal(result.status, 404);
    assert.equal(result.json, undefined);
  });

  it('token endpoint failure propagates without leaking the client secret', async () => {
    const failingFetch: typeof fetch = (async () =>
      makeResponse({ status: 400, body: { error: 'invalid_client' } })) as typeof fetch;
    const http = new ProvisioningHttp({
      graphCredential: GRAPH_CRED,
      fetchImpl: failingFetch,
      log: () => {},
    });
    const err = await rejection(http.request(CREATE_APP));
    assert.ok(err instanceof Error);
    assert.ok(err.message.includes('aad token 400'));
    assert.ok(err.message.includes('https://graph.microsoft.com/.default'));
    assert.ok(!err.message.includes('graph-secret-value'));
  });
});

describe('ProvisioningHttp 429 backoff', () => {
  it('honours Retry-After seconds and retries with the cached token', async () => {
    const { http, calls, sleeps } = harness([
      route(
        '/applications',
        { status: 429, headers: { 'Retry-After': '7' } },
        { status: 201, body: { id: 'obj-1' } },
      ),
    ]);
    const result = await http.request(CREATE_APP);
    assert.equal(result.kind, 'ok');
    assert.deepEqual(sleeps, [7000], 'sleep comes from the header, not a fixed value');
    assert.equal(tokenCalls(calls).length, 1, 'retry reuses the cached token');
    assert.equal(calls.filter((c) => c.url.includes('/applications')).length, 2);
  });

  it('falls back to exponential backoff when no Retry-After is sent', async () => {
    const { http, sleeps } = harness([
      route('/applications', { status: 429 }, { status: 429 }, { status: 201, body: {} }),
    ]);
    const result = await http.request(CREATE_APP);
    assert.equal(result.kind, 'ok');
    assert.deepEqual(sleeps, [1000, 2000]);
  });

  it('a Retry-After beyond the 60s cap aborts typed instead of sleeping for hours', async () => {
    // Graph/ARM can hint 'come back tomorrow' (Retry-After: 86400). An
    // unbounded await inside the provisioning job would hang it for a day —
    // the cap turns that into the typed reschedule signal, carrying the full
    // hint for the job runner.
    const { http, sleeps } = harness([
      route('/applications', {
        status: 429,
        body: {},
        headers: { 'Retry-After': '86400' },
      }),
    ]);
    const err = await rejection(http.request(CREATE_APP));
    assert.ok(err instanceof ProvisioningThrottledError);
    assert.equal(err.retryAfterSeconds, 86400, 'full hint preserved for rescheduling');
    assert.deepEqual(sleeps, [], 'no multi-hour sleep may be awaited');
  });

  it('clamps long-running poll pacing hints to the backoff cap', async () => {
    const OP_URL = 'https://management.azure.com/operations/op-clamp';
    const { http, sleeps } = harness(
      [
        route(
          '/operations/op-clamp',
          { status: 202, headers: { 'Retry-After': '86400' } },
          { status: 200, body: { status: 'Succeeded' } },
        ),
        route(
          'botServices/omadia-agent-hr',
          { status: 202, headers: { 'Azure-AsyncOperation': OP_URL } },
          { status: 200, body: { name: 'omadia-agent-hr' } },
        ),
      ],
      { pollIntervalMs: 50 },
    );
    const result = await http.request(PUT_BOT);
    assert.equal(result.kind, 'ok');
    assert.deepEqual(
      sleeps,
      [50, 60_000],
      'the 86400s pacing hint is clamped to the 60s cap',
    );
  });

  it('exhausted budget → ProvisioningThrottledError with resource + Retry-After hint', async () => {
    const { http, calls, sleeps } = harness(
      [route('management.azure.com', { status: 429, headers: { 'Retry-After': '3' } })],
      { armCredential: ARM_CRED, max429Retries: 2 },
    );
    const err = await rejection(http.request(PUT_BOT_NO_POLL));
    assert.ok(err instanceof ProvisioningThrottledError);
    assert.equal(err.resource, 'arm');
    assert.equal(err.retryAfterSeconds, 3);
    assert.deepEqual(sleeps, [3000, 3000], 'two retries before the budget ends');
    assert.equal(calls.filter((c) => c.url.includes('management.azure.com')).length, 3);
  });
});

describe('ProvisioningHttp ARM long-running poll mode', () => {
  const OP_URL = 'https://management.azure.com/operations/op-1';

  it('201 + Azure-AsyncOperation → polls until Succeeded, re-reads the finished resource', async () => {
    // The PUT answers with the MID-PROVISIONING representation; returning it
    // after the operation succeeded would hand callers a stale body (empty
    // endpoint, provisioningState 'Creating'). The finished resource must be
    // re-read once the operation reports Succeeded.
    const putBody = {
      name: 'omadia-agent-hr',
      properties: { provisioningState: 'Creating', endpoint: '' },
    };
    const finishedBody = {
      name: 'omadia-agent-hr',
      properties: {
        provisioningState: 'Succeeded',
        endpoint: 'https://example.invalid/api/teams/messages/hr',
      },
    };
    const { http, calls, sleeps } = harness(
      [
        route(
          '/operations/op-1',
          { status: 200, body: { status: 'InProgress' } },
          { status: 200, body: { status: 'Succeeded' } },
        ),
        route(
          'botServices/omadia-agent-hr',
          {
            status: 201,
            body: putBody,
            headers: { 'Azure-AsyncOperation': OP_URL },
          },
          { status: 200, body: finishedBody },
        ),
      ],
      { pollIntervalMs: 50 },
    );
    const result = await http.request(PUT_BOT);
    assert.equal(result.kind, 'ok');
    assert.equal(result.status, 201);
    assert.deepEqual(
      result.json,
      finishedBody,
      'the stale pre-poll PUT body must not be returned after the poll',
    );
    assert.equal(calls.filter((c) => c.url.includes('/operations/op-1')).length, 2);
    const botCalls = calls.filter((c) => c.url.includes('botServices'));
    assert.equal(botCalls.length, 2, 'PUT + final re-read GET');
    assert.equal(botCalls[1]?.init?.method, 'GET');
    assert.deepEqual(sleeps, [50, 50]);
  });

  it('202 with only a Location header polls the resource shape until terminal', async () => {
    // ARM's operation-results (Location) endpoint returns the RESOURCE
    // representation (properties.provisioningState), never a top-level
    // `status` — selecting the operation-status parser here would poll
    // forever (major review finding of the http unit).
    const done = {
      name: 'omadia-agent-hr',
      properties: { provisioningState: 'Succeeded' },
    };
    const { http, calls } = harness(
      [
        route(
          '/operationresults/loc-1',
          { status: 200, body: { properties: { provisioningState: 'Creating' } } },
          { status: 200, body: done },
        ),
        route('botServices/omadia-agent-hr', {
          status: 202,
          headers: { Location: 'https://management.azure.com/operationresults/loc-1' },
        }),
      ],
      { pollIntervalMs: 5 },
    );
    const result = await http.request(PUT_BOT);
    assert.equal(result.kind, 'ok');
    assert.deepEqual(result.json, done, 'the Location poll body is the fresh resource');
    assert.equal(
      calls.filter((c) => c.url.includes('/operationresults/loc-1')).length,
      2,
    );
  });

  it('honours Retry-After pacing hints from the 202 and from interim polls', async () => {
    const { http, sleeps } = harness(
      [
        route(
          '/operations/op-1',
          { status: 202, headers: { 'Retry-After': '4' } },
          { status: 200, body: { status: 'Succeeded' } },
        ),
        route(
          'botServices/omadia-agent-hr',
          { status: 202, headers: { 'Azure-AsyncOperation': OP_URL, 'Retry-After': '2' } },
          { status: 200, body: { name: 'omadia-agent-hr' } },
        ),
      ],
      { pollIntervalMs: 50 },
    );
    const result = await http.request(PUT_BOT);
    assert.equal(result.kind, 'ok');
    // Bodiless 202 → after the operation succeeds, the resource is re-read.
    assert.deepEqual(result.json, { name: 'omadia-agent-hr' });
    assert.deepEqual(sleeps, [2000, 4000]);
  });

  it('operation Failed → bounded error, no retry loop', async () => {
    const { http } = harness(
      [
        route('/operations/op-1', {
          status: 200,
          body: { status: 'Failed', error: { code: 'InvalidBotData' } },
        }),
        route('botServices/omadia-agent-hr', {
          status: 201,
          body: {},
          headers: { 'Azure-AsyncOperation': OP_URL },
        }),
      ],
      { pollIntervalMs: 1 },
    );
    const err = await rejection(http.request(PUT_BOT));
    assert.ok(err instanceof Error);
    assert.ok(err.message.includes('arm botServices.put'));
    assert.ok(err.message.includes('failed'));
  });

  it('201 with terminal provisioningState and no async header returns without polling', async () => {
    const { http, sleeps } = harness([
      route('botServices/omadia-agent-hr', {
        status: 201,
        body: { name: 'omadia-agent-hr', properties: { provisioningState: 'Succeeded' } },
      }),
    ]);
    const result = await http.request(PUT_BOT);
    assert.equal(result.kind, 'ok');
    assert.deepEqual(sleeps, []);
  });

  it('no async header + pending provisioningState → polls the resource itself', async () => {
    const done = {
      name: 'omadia-agent-hr',
      properties: { provisioningState: 'Succeeded' },
    };
    const { http, calls, sleeps } = harness(
      [
        route(
          'botServices/omadia-agent-hr',
          { status: 201, body: { properties: { provisioningState: 'Creating' } } },
          { status: 200, body: { properties: { provisioningState: 'Creating' } } },
          { status: 200, body: done },
        ),
      ],
      { pollIntervalMs: 25 },
    );
    const result = await http.request(PUT_BOT);
    assert.equal(result.kind, 'ok');
    assert.deepEqual(result.json, done);
    assert.equal(calls.filter((c) => c.url.includes('botServices')).length, 3);
    assert.deepEqual(sleeps, [25, 25]);
  });
});

// ---------------------------------------------------------------------------
// byte5ai/omadia#916 — some APIs report a duplicate without a 409. Entra
// answers a taken `uniqueName` on POST /applications with 400
// Request_BadRequest, which the choke point must map onto the SAME idempotent
// conflict signal — otherwise the app-registration adopt branch is
// unreachable and the retry budget burns on an impossible create.
// ---------------------------------------------------------------------------

describe('ProvisioningHttp non-409 conflict rules (#916)', () => {
  const CONFLICT_RULES = [
    { status: 400, codes: ['ObjectConflict'] },
    { status: 400, messageIncludes: ['same value for property uniquename'] },
  ] as const;

  const withRules: ProvisioningRequest = {
    ...CREATE_APP,
    conflictOn: CONFLICT_RULES,
  };

  const badRequest = (message: string, code = 'Request_BadRequest'): ResponseSpec => ({
    status: 400,
    body: { error: { code, message } },
  });

  it('matches on the Graph error code', async () => {
    const { http } = harness([
      route('/applications', badRequest('anything', 'ObjectConflict')),
    ]);
    const response = await http.request(withRules);
    assert.equal(response.kind, 'conflict');
    assert.equal(response.status, 400);
  });

  it('matches on the error message, whatever the code', async () => {
    const { http } = harness([
      route(
        '/applications',
        badRequest(
          'Another object with the same value for property uniqueName already exists.',
        ),
      ),
    ]);
    const response = await http.request(withRules);
    assert.equal(response.kind, 'conflict');
    assert.deepEqual(
      (response.json as { error: { code: string } }).error.code,
      'Request_BadRequest',
      'the parsed body travels with the conflict signal',
    );
  });

  it('leaves an unrelated 400 on the error path', async () => {
    const { http } = harness([
      route('/applications', badRequest('Property displayName is invalid.')),
    ]);
    const err = await rejection(http.request(withRules));
    assert.match((err as Error).message, /applications\.create 400/);
  });

  it('never fires for a status no rule declares', async () => {
    const { http } = harness([
      route('/applications', { status: 500, body: { error: { code: 'ObjectConflict' } } }),
    ]);
    const err = await rejection(http.request(withRules));
    assert.match((err as Error).message, /applications\.create 500/);
  });

  it('a request without rules keeps the plain 409-only mapping', async () => {
    const { http } = harness([
      route('/applications', badRequest('anything', 'ObjectConflict')),
    ]);
    const err = await rejection(http.request(CREATE_APP));
    assert.match((err as Error).message, /applications\.create 400/);
  });
});

// byte5ai/omadia#2043, rollout defect D7. A team install was refused by the
// team's resource-specific consent check. What reached the kernel from 0.9.0
// was `…backend request [redacted]. Workload Unknown. Request Url:
// https://teamsgraph.teams.microsoft.com/v1.0/groups/c6031502-025…`: the
// request name was redacted as a "long opaque credential" (it is 62 letters)
// and the URL that names the check a second time was cut at 200 characters.
// The kernel could not tell the refusal from a replication delay.
describe('ProvisioningHttp keeps what a Graph error says (0.9.1)', () => {
  const TEAM_ID = 'c6031502-025c-4297-92c6-b734e17e46c5';
  const BACKEND_REQUEST = 'IsUserAuthorizedToGrantGroupResourceSpecificPermissionsRequest';
  // The body Graph answered at the rollout (request ids replaced).
  const RSC_REFUSAL = {
    error: {
      code: 'BadRequest',
      message:
        `Failed to execute TeamsGraphService backend request ${BACKEND_REQUEST}. Workload Unknown. ` +
        `Request Url: https://teamsgraph.teams.microsoft.com/v1.0/groups/${TEAM_ID}/isUserAuthorizedToGrantResourceSpecificPermissions, ` +
        'Request Method: POST, Response Status Code: BadRequest, Response Headers: Strict-Transport-Security: max-age=2592000\r\n' +
        'request-id: 5f0c1d2e-0000-4000-8000-000000000001\r\n, Reason Phrase: Bad Request, ' +
        'Request ID 5f0c1d2e-0000-4000-8000-000000000001.',
      innerError: {
        code: 'InvalidRequest',
        date: '2026-11-18T17:01:42',
        'request-id': '9a1b2c3d-0000-4000-8000-000000000002',
        'client-request-id': '9a1b2c3d-0000-4000-8000-000000000003',
      },
    },
  };
  const INSTALL: ProvisioningRequest = {
    resource: 'graph',
    method: 'POST',
    url: `https://graph.microsoft.com/v1.0/teams/${TEAM_ID}/installedApps`,
    step: 'teams.installedApps.add',
    jsonBody: {
      'teamsApp@odata.bind': 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/app-1',
      consentedPermissionSet: { resourceSpecificPermissions: [] },
    },
    missingScopesOn403: ['TeamsAppInstallation.ReadWriteAndConsentForTeam.All'],
  };

  it('the message carries the backend request name and the whole Teams URL', async () => {
    const { http } = harness([route('/installedApps', { status: 400, body: RSC_REFUSAL })]);
    const err = await rejection(http.request(INSTALL));

    assert.ok(err instanceof ProvisioningRequestError);
    assert.equal(err.status, 400);
    assert.ok(err.message.includes(BACKEND_REQUEST), err.message);
    assert.ok(err.message.includes(`/groups/${TEAM_ID}/isUserAuthorizedToGrantResourceSpecificPermissions`));
    assert.ok(err.message.includes('"InvalidRequest"'));
    assert.ok(err.message.includes(INSTALL.url));
    assert.ok(!err.message.includes('[redacted]'));
    // The pattern the kernel classifies with (teamsProvisioningJob.ts).
    assert.match(err.message, /isUserAuthorizedToGrant\w*ResourceSpecificPermissions/i);
    assert.match(err.message, /backend request (\w+Request)\b/);
  });

  it('graphError carries code, inner code, message, request ids and body as fields', async () => {
    const { http } = harness([route('/installedApps', { status: 400, body: RSC_REFUSAL })]);
    const err = await rejection(http.request(INSTALL));

    assert.ok(err instanceof ProvisioningRequestError);
    assert.deepEqual(err.graphError, {
      code: 'BadRequest',
      innerCode: 'InvalidRequest',
      message: RSC_REFUSAL.error.message,
      requestId: '9a1b2c3d-0000-4000-8000-000000000002',
      clientRequestId: '9a1b2c3d-0000-4000-8000-000000000003',
      body: JSON.stringify(RSC_REFUSAL),
    });
  });

  it('holds no token and nothing of the request body', async () => {
    const { http } = harness([route('/installedApps', { status: 400, body: RSC_REFUSAL })]);
    const err = await rejection(http.request({ ...INSTALL, bearerToken: 'delegated-secret-token-value' }));

    assert.ok(err instanceof ProvisioningRequestError);
    const everything = `${err.message} ${JSON.stringify(err.graphError)}`;
    for (const secret of ['delegated-secret-token-value', 'graph-secret-value', 'tok:', 'consentedPermissionSet', 'odata.bind']) {
      assert.ok(!everything.includes(secret), `${secret} reached the error`);
    }
  });

  it('redacts a credential the API echoes, in the message and in every field', async () => {
    const jwt = 'eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJhZG1pbiJ9.c2lnbmF0dXJl';
    const opaque = `0.AAAA${'x9_'.repeat(30)}`;
    const { http } = harness([
      route('/installedApps', {
        status: 400,
        body: {
          error: { code: 'BadRequest', message: `token ${jwt} rejected, refresh ${opaque}`, access_token: jwt },
        },
      }),
    ]);
    const err = await rejection(http.request(INSTALL));

    assert.ok(err instanceof ProvisioningRequestError);
    const everything = `${err.message} ${JSON.stringify(err.graphError)}`;
    assert.ok(!everything.includes(jwt));
    assert.ok(!everything.includes('x9_x9_x9_'));
    assert.match(err.graphError?.message ?? '', /token \[redacted\] rejected/);
  });

  it('bounds the message detail at 2,000 and the body field at 4,000 characters', async () => {
    const { http } = harness([
      route('/installedApps', {
        status: 400,
        body: { error: { code: 'BadRequest', message: 'word '.repeat(2000) } },
      }),
    ]);
    const err = await rejection(http.request(INSTALL));

    assert.ok(err instanceof ProvisioningRequestError);
    assert.equal(err.graphError?.message?.length, 2000);
    assert.equal(err.graphError?.body?.length, 4000);
    assert.ok(err.message.length < 2000 + 300, String(err.message.length));
  });

  it('takes the request ids from the response headers when the body has none', async () => {
    const { http } = harness([
      route('/installedApps', {
        status: 400,
        body: { error: { code: 'BadRequest', message: 'no inner error', innererror: { code: 'Lowercase' } } },
        headers: { 'request-id': 'header-request-id', 'client-request-id': 'header-client-id' },
      }),
    ]);
    const err = await rejection(http.request(INSTALL));

    assert.ok(err instanceof ProvisioningRequestError);
    assert.equal(err.graphError?.innerCode, 'Lowercase');
    assert.equal(err.graphError?.requestId, 'header-request-id');
    assert.equal(err.graphError?.clientRequestId, 'header-client-id');
  });

  it('a body that is no error envelope yields no graphError, and the text stays in the message', async () => {
    const { http } = harness([route('/installedApps', { status: 502, body: 'Bad Gateway' })]);
    const err = await rejection(http.request(INSTALL));

    assert.ok(err instanceof ProvisioningRequestError);
    assert.equal(err.graphError, undefined);
    assert.ok(err.message.includes('Bad Gateway'));
  });

  it('redaction keeps a long name made of letters and still removes long credentials', () => {
    assert.equal(redactSecrets(`request ${BACKEND_REQUEST}.`), `request ${BACKEND_REQUEST}.`);
    const refreshToken = `M.R3_BAY.${'aB3dE6gH9_kL-'.repeat(8)}`;
    assert.ok(!redactSecrets(`refresh ${refreshToken}`).includes('aB3dE6gH9'));
    assert.equal(redactSecrets('a'.repeat(59) + '1'), '[redacted]');
  });
});


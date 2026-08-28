import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  TEAM_READ_SCOPE,
  TeamLookupClient,
} from '../src/teamsProvisioner/teamLookup.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import { ConsentMissingError } from '../src/teamsProvisioner/errors.js';

// The team-lookup step: GET /teams/{id}?$select=id,displayName behind the
// shared ProvisioningHttp choke point. What it must prove:
//
//   * a hit yields the NAME — the whole reason the step exists, since every
//     other method addresses a team by GUID and so does every consumer UI;
//   * a 404 is an ANSWER (`found: false`), not a throw, because a consumer's
//     fallback is simply to keep showing the id;
//   * a 403 still maps to the typed ConsentMissingError carrying the one
//     scope this step needs, so an unconsented tenant is actionable rather
//     than mysterious;
//   * `$select` is actually sent — it is the scope boundary made literal.
//
// Mocked-fetch style follows tests/provisionerInstall.test.ts.

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

function mockFetch(routes: Route[], calls: FetchCall[]): typeof fetch {
  const queues = routes.map((r) => ({ ...r, next: 0 }));
  return (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, ...(init ? { init } : {}) });
    if (url.includes('login.microsoftonline.com')) {
      return makeResponse({
        status: 200,
        body: { access_token: 'tok', expires_in: 3600 },
      });
    }
    for (const q of queues) {
      if (url.includes(q.match)) {
        const spec = q.responses[Math.min(q.next, q.responses.length - 1)];
        q.next += 1;
        if (!spec) break;
        return makeResponse(spec);
      }
    }
    return makeResponse({ status: 500, body: { error: 'unrouted ' + url } });
  }) as typeof fetch;
}

const TEAM_ID = 'team-0001';
const TEAM_URL_MATCH = `/teams/${TEAM_ID}`;

function harness(routes: Route[]): {
  client: TeamLookupClient;
  calls: FetchCall[];
} {
  const calls: FetchCall[] = [];
  const http = new ProvisioningHttp({
    graphCredential: {
      tenantId: 'tenant-1',
      clientId: 'app-graph',
      clientSecret: 'graph-secret-value',
    },
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async () => {},
    max429Retries: 2,
  });
  return { client: new TeamLookupClient({ http, log: () => {} }), calls };
}

function graphCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter((c) => c.url.includes(TEAM_URL_MATCH));
}

describe('TeamLookupClient.getTeam', () => {
  it('resolves the display name and selects only id + displayName', async () => {
    const { client, calls } = harness([
      route(TEAM_URL_MATCH, {
        status: 200,
        body: { id: TEAM_ID, displayName: 'Marketing' },
      }),
    ]);
    const result = await client.getTeam({ teamId: TEAM_ID });
    assert.deepEqual(result, {
      found: true,
      teamId: TEAM_ID,
      displayName: 'Marketing',
    });
    const call = graphCalls(calls)[0];
    assert.ok(call, 'the lookup must hit Graph');
    // The scope boundary made literal: a label is all that is requested.
    assert.match(call.url, /\$select=id,displayName/);
    assert.equal(call.init?.method, 'GET');
  });

  it('answers found:false on 404 instead of throwing', async () => {
    const { client } = harness([
      route(TEAM_URL_MATCH, {
        status: 404,
        body: { error: { code: 'ItemNotFound' } },
      }),
    ]);
    // A team that is gone (or invisible to this tenant app) is a fact about
    // the tenant, not a failed call — the consumer keeps showing the id.
    assert.deepEqual(await client.getTeam({ teamId: TEAM_ID }), { found: false });
  });

  it('answers found:false when Graph returns no usable displayName', async () => {
    const { client } = harness([
      route(TEAM_URL_MATCH, { status: 200, body: { id: TEAM_ID, displayName: '  ' } }),
    ]);
    // Reporting found:true here would push an empty label into the UI, which
    // is strictly worse than the id it replaced.
    assert.deepEqual(await client.getTeam({ teamId: TEAM_ID }), { found: false });
  });

  it('maps 403 to ConsentMissingError carrying Team.ReadBasic.All', async () => {
    const { client } = harness([
      route(TEAM_URL_MATCH, {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);
    await assert.rejects(
      () => client.getTeam({ teamId: TEAM_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.deepEqual(err.missingScopes, [TEAM_READ_SCOPE]);
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });

  it('falls back to the requested id when Graph omits one', async () => {
    const { client } = harness([
      route(TEAM_URL_MATCH, { status: 200, body: { displayName: 'Marketing' } }),
    ]);
    const result = await client.getTeam({ teamId: TEAM_ID });
    assert.deepEqual(result, {
      found: true,
      teamId: TEAM_ID,
      displayName: 'Marketing',
    });
  });
});

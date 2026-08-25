import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  TEAM_INSTALL_SCOPE,
  TeamInstallClient,
  type ConsentedPermissionSet,
} from '../src/teamsProvisioner/install.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  ConsentMissingError,
  ProvisioningThrottledError,
  TeamsProvisionerError,
} from '../src/teamsProvisioner/errors.js';

// The team-install step: POST /teams/{id}/installedApps behind the shared
// ProvisioningHttp choke point — 409 resolves to the idempotent
// 'already-existed' outcome (never an exception), 403 maps to the same typed
// ConsentMissingError family as the catalog upload, 429 rides the shared
// Retry-After backoff. Mocked-fetch style follows
// tests/teamsProvisionerHttp.test.ts.

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

/** Route table with per-route response queues (last entry repeats). */
function mockFetch(routes: Route[], calls: FetchCall[]): typeof fetch {
  const queues = routes.map((r) => ({ ...r, next: 0 }));
  return (async (input: RequestInfo | URL, init?: RequestInit) => {
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
const TEAMS_APP_ID = 'catalog-app-0001';
const INSTALLATION_ID = 'NmFiOTZlZm-installation-id';
const INSTALL_URL_MATCH = `/teams/${TEAM_ID}/installedApps`;

const CONSENTED: ConsentedPermissionSet = {
  resourceSpecificPermissions: [
    { permissionValue: 'ChannelMessage.Read.Group', permissionType: 'application' },
  ],
};

function harness(routes: Route[]): {
  client: TeamInstallClient;
  calls: FetchCall[];
  sleeps: number[];
} {
  const calls: FetchCall[] = [];
  const sleeps: number[] = [];
  const http = new ProvisioningHttp({
    graphCredential: {
      tenantId: 'tenant-1',
      clientId: 'app-graph',
      clientSecret: 'graph-secret-value',
    },
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async (ms) => {
      sleeps.push(ms);
    },
    max429Retries: 2,
  });
  const client = new TeamInstallClient({ http, log: () => {} });
  return { client, calls, sleeps };
}

/** The recorded Graph install calls (token calls filtered out). */
function installCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter((c) => c.url.includes(INSTALL_URL_MATCH));
}

function parsedBody(call: FetchCall): Record<string, unknown> {
  return JSON.parse(String(call.init?.body ?? '{}')) as Record<string, unknown>;
}

describe('TeamInstallClient.installToTeam', () => {
  it('POSTs the odata.bind body and returns created with the installation id (201)', async () => {
    const { client, calls } = harness([
      route(INSTALL_URL_MATCH, {
        status: 201,
        body: { id: INSTALLATION_ID },
      }),
    ]);

    const result = await client.installToTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, {
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
      installationId: INSTALLATION_ID,
    });

    const [call] = installCalls(calls);
    assert.ok(call, 'expected one install POST');
    assert.equal(call.init?.method, 'POST');
    assert.equal(
      call.url,
      `https://graph.microsoft.com/v1.0/teams/${TEAM_ID}/installedApps`,
    );
    const body = parsedBody(call);
    assert.equal(
      body['teamsApp@odata.bind'],
      `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${TEAMS_APP_ID}`,
    );
    // Plain install: the consent key must be ABSENT, not undefined/null.
    assert.ok(!('consentedPermissionSet' in body));
    const headers = call.init?.headers as Record<string, string>;
    assert.equal(headers['Authorization'], 'Bearer tok');
    assert.equal(headers['content-type'], 'application/json');
  });

  it('omits installationId when the 2xx response has no body', async () => {
    const { client } = harness([route(INSTALL_URL_MATCH, { status: 201 })]);

    const result = await client.installToTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, { teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID });
    assert.ok(!('installationId' in result.value));
  });

  it('sends consentedPermissionSet verbatim when (and only when) provided', async () => {
    const { client, calls } = harness([
      route(INSTALL_URL_MATCH, { status: 201, body: { id: INSTALLATION_ID } }),
    ]);

    const result = await client.installToTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
      consentedPermissionSet: CONSENTED,
    });

    assert.equal(result.outcome, 'created');
    const [call] = installCalls(calls);
    assert.ok(call, 'expected one install POST');
    const body = parsedBody(call);
    assert.deepEqual(body['consentedPermissionSet'], {
      resourceSpecificPermissions: [
        {
          permissionValue: 'ChannelMessage.Read.Group',
          permissionType: 'application',
        },
      ],
    });
  });

  it('resolves 409 (already installed) to already-existed — success, no throw', async () => {
    const { client, calls } = harness([
      route(INSTALL_URL_MATCH, {
        status: 409,
        body: { error: { code: 'Conflict', message: 'Duplicated app id' } },
      }),
    ]);

    const result = await client.installToTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-existed');
    assert.deepEqual(result.value, { teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID });
    assert.equal(installCalls(calls).length, 1);
  });

  it('maps 403 to ConsentMissingError carrying the install scope (graph)', async () => {
    const { client } = harness([
      route(INSTALL_URL_MATCH, {
        status: 403,
        body: { error: { code: 'Forbidden' } },
      }),
    ]);

    await assert.rejects(
      client.installToTeam({ teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        // Same typed family as uploadToCatalog → one fallback branch.
        assert.ok(err instanceof TeamsProvisionerError);
        assert.deepEqual(err.missingScopes, [TEAM_INSTALL_SCOPE]);
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });

  it('retries 429 honouring Retry-After, then succeeds', async () => {
    const { client, calls, sleeps } = harness([
      route(
        INSTALL_URL_MATCH,
        { status: 429, headers: { 'Retry-After': '7' } },
        { status: 201, body: { id: INSTALLATION_ID } },
      ),
    ]);

    const result = await client.installToTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.equal(result.value.installationId, INSTALLATION_ID);
    assert.equal(installCalls(calls).length, 2);
    assert.deepEqual(sleeps, [7000]);
  });

  it('throws ProvisioningThrottledError when the 429 budget is exhausted', async () => {
    const { client, calls } = harness([
      route(INSTALL_URL_MATCH, {
        status: 429,
        headers: { 'Retry-After': '3' },
      }),
    ]);

    await assert.rejects(
      client.installToTeam({ teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ProvisioningThrottledError);
        assert.equal(err.resource, 'graph');
        assert.equal(err.retryAfterSeconds, 3);
        return true;
      },
    );
    // max429Retries: 2 → initial attempt + 2 retries.
    assert.equal(installCalls(calls).length, 3);
  });

  it('rejects empty teamId / teamsAppId before any fetch', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.installToTeam({ teamId: '  ', teamsAppId: TEAMS_APP_ID }),
      /invalid_argument: 'teamId'/,
    );
    await assert.rejects(
      client.installToTeam({ teamId: TEAM_ID, teamsAppId: '' }),
      /invalid_argument: 'teamsAppId'/,
    );
    assert.equal(calls.length, 0);
  });
});

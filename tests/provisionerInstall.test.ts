import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  TEAM_INSTALL_CONSENT_SCOPE,
  TEAM_INSTALL_SCOPE,
  TeamInstallClient,
  type ConsentedPermissionSet,
} from '../src/teamsProvisioner/install.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  ConsentMissingError,
  ProvisioningThrottledError,
  RscPermissionsMismatchError,
  TeamsProvisionerError,
  isTransientProvisioningFailure,
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
const TEAMS_APP_ID = 'catalog-app-0001';
const INSTALLATION_ID = 'NmFiOTZlZm-installation-id';
const INSTALL_URL_MATCH = `/teams/${TEAM_ID}/installedApps`;

const CONSENTED: ConsentedPermissionSet = {
  resourceSpecificPermissions: [
    { permissionValue: 'ChannelMessage.Read.Group', permissionType: 'application' },
  ],
};

/** See the twin in tests/provisionerChatInstall.test.ts — the 0.8.2 lookup of
 *  the app's own declared RSC permissions. Default: the app declares none, so
 *  the body keeps its pre-0.8.2 shape. */
const RSC_LOOKUP_MATCH = '/appCatalogs/teamsApps?';

const rscLookupRoute = (
  permissions: readonly { permissionValue: string; permissionType: string }[],
): Route =>
  route(RSC_LOOKUP_MATCH, {
    status: 200,
    body: {
      value: [
        {
          id: TEAMS_APP_ID,
          appDefinitions: [
            {
              id: 'definition-1',
              authorization: {
                requiredPermissionSet: { resourceSpecificPermissions: permissions },
              },
            },
          ],
        },
      ],
    },
  });

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
    fetchImpl: mockFetch([...routes, rscLookupRoute([])], calls),
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

// The uninstall direction (byte5ai/omadia#900): Graph deletes an installation
// by INSTALLATION id, so it is a lookup-then-DELETE pair. "Not installed" —
// whether the lookup misses or the DELETE races into a 404 — is the
// idempotent 'already-absent' success, never an exception. 403/429 ride the
// same shared paths as the install direction.

/** Only the lookup GETs (they carry a query string). */
function lookupCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter(
    (c) => c.url.includes(INSTALL_URL_MATCH) && c.url.includes('$filter='),
  );
}

/** Only the DELETEs of a concrete installation. */
function deleteCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter((c) => c.init?.method === 'DELETE');
}

/** A lookup response body carrying one matching installation. */
const foundBody = (
  installationId = INSTALLATION_ID,
  teamsAppId = TEAMS_APP_ID,
): unknown => ({
  value: [{ id: installationId, teamsApp: { id: teamsAppId } }],
});

describe('TeamInstallClient.uninstallFromTeam', () => {
  it('resolves the installation id, DELETEs it, and reports uninstalled', async () => {
    const { client, calls } = harness([
      route(`${INSTALL_URL_MATCH}?`, { status: 200, body: foundBody() }),
      route(`${INSTALL_URL_MATCH}/${INSTALLATION_ID}`, { status: 204 }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'uninstalled');
    assert.deepEqual(result.value, {
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
      installationId: INSTALLATION_ID,
    });

    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected one lookup GET');
    assert.equal(lookup.init?.method, 'GET');
    assert.equal(
      lookup.url,
      `https://graph.microsoft.com/v1.0/teams/${TEAM_ID}/installedApps` +
        `?$expand=teamsApp&$filter=${encodeURIComponent(`teamsApp/id eq '${TEAMS_APP_ID}'`)}`,
    );

    const [del] = deleteCalls(calls);
    assert.ok(del, 'expected one DELETE');
    assert.equal(
      del.url,
      `https://graph.microsoft.com/v1.0/teams/${TEAM_ID}/installedApps/${INSTALLATION_ID}`,
    );
  });

  it('reports already-absent WITHOUT deleting when the lookup finds nothing', async () => {
    const { client, calls } = harness([
      route(`${INSTALL_URL_MATCH}?`, { status: 200, body: { value: [] } }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.deepEqual(result.value, { teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID });
    assert.ok(!('installationId' in result.value));
    assert.equal(lookupCalls(calls).length, 1);
    assert.equal(deleteCalls(calls).length, 0, 'must not DELETE on a lookup miss');
  });

  it('ignores lookup entries whose expanded teamsApp is a different app', async () => {
    const { client, calls } = harness([
      route(`${INSTALL_URL_MATCH}?`, {
        status: 200,
        // A tenant that ignores $filter must not make us delete the wrong one.
        body: { value: [{ id: 'other-installation', teamsApp: { id: 'other-app' } }] },
      }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.equal(deleteCalls(calls).length, 0);
  });

  it('treats a 404 on the DELETE as already-absent (removal race)', async () => {
    const { client, calls } = harness([
      route(`${INSTALL_URL_MATCH}?`, { status: 200, body: foundBody() }),
      route(`${INSTALL_URL_MATCH}/${INSTALLATION_ID}`, {
        status: 404,
        body: { error: { code: 'NotFound' } },
      }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.equal(result.value.installationId, INSTALLATION_ID);
    assert.equal(deleteCalls(calls).length, 1);
  });

  it('treats a 404 on the lookup (team gone) as already-absent', async () => {
    const { client, calls } = harness([
      route(`${INSTALL_URL_MATCH}?`, {
        status: 404,
        body: { error: { code: 'NotFound' } },
      }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.equal(deleteCalls(calls).length, 0);
  });

  it("escapes OData quotes in the teamsAppId filter (quote doubling)", async () => {
    const nastyAppId = "app' or id ne '";
    const { client, calls } = harness([
      route(`${INSTALL_URL_MATCH}?`, { status: 200, body: { value: [] } }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: nastyAppId,
    });

    assert.equal(result.outcome, 'already-absent');
    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected one lookup GET');
    const filter = new URL(lookup.url).searchParams.get('$filter');
    assert.equal(filter, "teamsApp/id eq 'app'' or id ne '''");
  });

  it('maps 403 to ConsentMissingError carrying the install scope (graph)', async () => {
    const { client } = harness([
      route(`${INSTALL_URL_MATCH}?`, {
        status: 403,
        body: { error: { code: 'Forbidden' } },
      }),
    ]);

    await assert.rejects(
      client.uninstallFromTeam({ teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.ok(err instanceof TeamsProvisionerError);
        assert.deepEqual(err.missingScopes, [TEAM_INSTALL_SCOPE]);
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });

  it('maps a 403 on the DELETE itself to ConsentMissingError too', async () => {
    const { client } = harness([
      route(`${INSTALL_URL_MATCH}?`, { status: 200, body: foundBody() }),
      route(`${INSTALL_URL_MATCH}/${INSTALLATION_ID}`, { status: 403 }),
    ]);

    await assert.rejects(
      client.uninstallFromTeam({ teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.deepEqual(err.missingScopes, [TEAM_INSTALL_SCOPE]);
        return true;
      },
    );
  });

  it('retries 429 honouring Retry-After, then uninstalls', async () => {
    const { client, calls, sleeps } = harness([
      route(
        `${INSTALL_URL_MATCH}?`,
        { status: 429, headers: { 'Retry-After': '5' } },
        { status: 200, body: foundBody() },
      ),
      route(`${INSTALL_URL_MATCH}/${INSTALLATION_ID}`, { status: 204 }),
    ]);

    const result = await client.uninstallFromTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'uninstalled');
    assert.equal(lookupCalls(calls).length, 2);
    assert.deepEqual(sleeps, [5000]);
  });

  it('throws ProvisioningThrottledError when the 429 budget is exhausted', async () => {
    const { client } = harness([
      route(`${INSTALL_URL_MATCH}?`, {
        status: 429,
        headers: { 'Retry-After': '3' },
      }),
    ]);

    await assert.rejects(
      client.uninstallFromTeam({ teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ProvisioningThrottledError);
        assert.equal(err.resource, 'graph');
        assert.equal(err.retryAfterSeconds, 3);
        return true;
      },
    );
  });

  it('rejects empty teamId / teamsAppId before any fetch', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.uninstallFromTeam({ teamId: '  ', teamsAppId: TEAMS_APP_ID }),
      /invalid_argument: 'teamId'/,
    );
    await assert.rejects(
      client.uninstallFromTeam({ teamId: TEAM_ID, teamsAppId: '' }),
      /invalid_argument: 'teamsAppId'/,
    );
    assert.equal(calls.length, 0);
  });
});

// ---------------------------------------------------------------------------
// Resource-specific consent (0.8.2) — the team direction carried the SAME bug
// as the chat direction, latently: it offered `consentedPermissionSet` as a
// caller field and no caller ever filled it, so an app declaring RSC was
// refused with 400 ResourceSpecificPermissionsMismatch the moment a tenant
// held the consent-capable role. It had simply never been run at a tenant
// that got that far. The resolution is now shared with the chat direction.
// ---------------------------------------------------------------------------

const DECLARED_RSC = [
  { permissionValue: 'ChannelMessage.Read.Group', permissionType: 'application' },
  { permissionValue: 'TeamsActivity.Send.Group', permissionType: 'application' },
];

describe('TeamInstallClient.installToTeam — resource-specific consent', () => {
  it('consents to the declared permissions without the caller passing any', async () => {
    const { client, calls } = harness([
      route(INSTALL_URL_MATCH, { status: 201, body: { id: INSTALLATION_ID } }),
      rscLookupRoute(DECLARED_RSC),
    ]);

    const result = await client.installToTeam({
      teamId: TEAM_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    const [call] = installCalls(calls);
    assert.ok(call);
    assert.deepEqual(parsedBody(call)['consentedPermissionSet'], {
      resourceSpecificPermissions: DECLARED_RSC,
    });
  });

  it('maps 400 ResourceSpecificPermissionsMismatch to the team consent role', async () => {
    const { client } = harness([
      route(INSTALL_URL_MATCH, {
        status: 400,
        body: {
          error: {
            code: 'ResourceSpecificPermissionsMismatch',
            message: 'The app requires resource-specific permissions.',
          },
        },
      }),
      rscLookupRoute(DECLARED_RSC),
    ]);

    await assert.rejects(
      client.installToTeam({ teamId: TEAM_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof RscPermissionsMismatchError);
        assert.equal(err.step, 'teams.installedApps.add');
        assert.equal(err.consentRole, TEAM_INSTALL_CONSENT_SCOPE);
        assert.equal(err.sentPermissionCount, DECLARED_RSC.length);
        assert.equal(isTransientProvisioningFailure(err), false);
        return true;
      },
    );
  });
});

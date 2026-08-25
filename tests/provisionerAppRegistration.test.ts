import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { SecretsAccessor } from '@omadia/plugin-api';

import {
  APP_REGISTRATION_SCOPE,
  AppRegistrationClient,
  type CreateAppRegistrationInput,
} from '../src/teamsProvisioner/appRegistration.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  ConsentMissingError,
  ProvisioningThrottledError,
} from '../src/teamsProvisioner/errors.js';

// createAppRegistration: three chained Graph writes (POST /applications →
// addPassword → POST /servicePrincipals) behind the shared ProvisioningHttp
// choke point, with rollback of partial failures; deleteAppRegistration as
// the idempotent rollback half. Mocked-fetch style follows
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

/** In-memory fake of the plugin-api SecretsAccessor (write-capable). */
function fakeSecrets(): { accessor: SecretsAccessor; vault: Map<string, string> } {
  const vault = new Map<string, string>();
  const accessor: SecretsAccessor = {
    get: async (key) => vault.get(key),
    require: async (key) => {
      const value = vault.get(key);
      if (value === undefined) throw new Error(`missing_secret: ${key}`);
      return value;
    },
    keys: async () => [...vault.keys()],
    set: async (key: string, value: string) => {
      vault.set(key, value);
    },
    delete: async (key: string) => {
      vault.delete(key);
    },
  };
  return { accessor, vault };
}

const TENANT_ID = 'tenant-1';
const APP_ID = 'aaaa1111-0000-0000-0000-000000000001';
const OBJECT_ID = 'obj-0001';
const SECRET_TEXT = 's3cr3t-value-never-returned';
const SECRET_REF = `teams_bot_password:${APP_ID}`;

const APP_BODY = {
  id: OBJECT_ID,
  appId: APP_ID,
  displayName: 'HR Agent',
  signInAudience: 'AzureADMyOrg',
  uniqueName: 'omadia-agent-hr',
};

const PASSWORD_BODY = {
  secretText: SECRET_TEXT,
  keyId: 'key-1',
  endDateTime: '2027-01-01T00:00:00Z',
};

const CREATE_INPUT: CreateAppRegistrationInput = {
  displayName: 'HR Agent',
  tenantMode: 'customer',
  uniqueName: 'omadia-agent-hr',
};

function harness(routes: Route[]): {
  client: AppRegistrationClient;
  calls: FetchCall[];
  vault: Map<string, string>;
} {
  const calls: FetchCall[] = [];
  const { accessor, vault } = fakeSecrets();
  const http = new ProvisioningHttp({
    graphCredential: {
      tenantId: TENANT_ID,
      clientId: 'connector-app',
      clientSecret: 'connector-secret',
    },
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async () => {},
    max429Retries: 1,
  });
  const client = new AppRegistrationClient({
    http,
    secrets: accessor,
    tenantId: TENANT_ID,
    log: () => {},
  });
  return { client, calls, vault };
}

const graphCalls = (calls: FetchCall[]): FetchCall[] =>
  calls.filter((c) => c.url.includes('graph.microsoft.com'));

const bodyOf = (call: FetchCall): Record<string, unknown> =>
  JSON.parse(String(call.init?.body ?? '{}')) as Record<string, unknown>;

const rejection = async (p: Promise<unknown>): Promise<unknown> =>
  p.then(() => assert.fail('expected the call to reject'), (e: unknown) => e);

// Route order matters (first match wins): specific paths before the plain
// `/applications` create route.
const HAPPY_ROUTES = (): Route[] => [
  route('/addPassword', { status: 200, body: PASSWORD_BODY }),
  route('/servicePrincipals', { status: 201, body: { id: 'sp-1', appId: APP_ID } }),
  route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
  route("/applications(appId='", { status: 204 }),
  route('/applications', { status: 201, body: APP_BODY }),
];

describe('createAppRegistration — happy path', () => {
  it('chains POST /applications → addPassword → POST /servicePrincipals', async () => {
    const { client, calls } = harness(HAPPY_ROUTES());
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'created');
    assert.equal(result.value.appId, APP_ID);
    assert.equal(result.value.secretRef, SECRET_REF);
    assert.equal(result.value.registration.objectId, OBJECT_ID);
    assert.equal(result.value.registration.tenantId, TENANT_ID);
    assert.equal(result.value.registration.tenantMode, 'customer');
    assert.equal(result.value.secretKeyId, 'key-1');
    assert.equal(result.value.secretEndDateTime, '2027-01-01T00:00:00Z');
    assert.equal(result.value.servicePrincipalOutcome, 'created');

    const graph = graphCalls(calls);
    assert.equal(graph.length, 3);
    assert.ok(graph[0]?.url.endsWith('/v1.0/applications'));
    assert.equal(graph[0]?.init?.method, 'POST');
    assert.ok(graph[1]?.url.endsWith(`/applications/${OBJECT_ID}/addPassword`));
    assert.ok(graph[2]?.url.endsWith('/servicePrincipals'));
    assert.deepEqual(bodyOf(graph[2]!), { appId: APP_ID });
  });

  it('persists the secret in the vault and never returns it in cleartext', async () => {
    const { client, vault } = harness(HAPPY_ROUTES());
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(vault.get(SECRET_REF), SECRET_TEXT);
    // The secret value must not cross the API boundary anywhere in the result.
    assert.ok(!JSON.stringify(result).includes(SECRET_TEXT));
  });

  it('pins signInAudience to AzureADMyOrg for BOTH tenant modes (invariant)', async () => {
    for (const tenantMode of ['customer', 'home'] as const) {
      const { client, calls } = harness(HAPPY_ROUTES());
      await client.createAppRegistration({ ...CREATE_INPUT, tenantMode });

      const create = graphCalls(calls)[0]!;
      const body = bodyOf(create);
      assert.equal(body['signInAudience'], 'AzureADMyOrg');
      assert.ok(!JSON.stringify(body).includes('AzureADMultipleOrgs'));
    }
  });

  it('sends the uniqueName idempotency key on create', async () => {
    const { client, calls } = harness(HAPPY_ROUTES());
    await client.createAppRegistration(CREATE_INPUT);
    assert.equal(bodyOf(graphCalls(calls)[0]!)['uniqueName'], 'omadia-agent-hr');
  });
});

describe('createAppRegistration — 409 idempotency', () => {
  it('409 on create → finds the existing app by uniqueName → already-existed', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 409, body: { error: 'exists' } }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', { status: 409, body: { error: 'uniqueName taken' } }),
    ];
    const { client, calls, vault } = harness(routes);
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'already-existed');
    assert.equal(result.value.appId, APP_ID);
    // Secret is still rotated: deterministic key, intended overwrite.
    assert.equal(vault.get(SECRET_REF), SECRET_TEXT);
    // 409 on the service principal is the idempotent signal, not an error.
    assert.equal(result.value.servicePrincipalOutcome, 'already-existed');

    const lookup = graphCalls(calls).find((c) =>
      c.url.includes("applications(uniqueName='omadia-agent-hr')"),
    );
    assert.ok(lookup, 'expected the uniqueName alternate-key lookup');
  });

  it('409 without a uniqueName cannot be resolved and rejects', async () => {
    const routes: Route[] = [
      route('/applications', { status: 409, body: { error: 'conflict' } }),
    ];
    const { client } = harness(routes);
    const err = await rejection(
      client.createAppRegistration({ displayName: 'HR Agent', tenantMode: 'customer' }),
    );
    assert.match((err as Error).message, /409 without a uniqueName/);
  });

  it('refuses to adopt a found registration that is not SingleTenant', async () => {
    const routes: Route[] = [
      route("/applications(uniqueName='", {
        status: 200,
        body: { ...APP_BODY, signInAudience: 'AzureADMultipleOrgs' },
      }),
      route('/applications', { status: 409, body: {} }),
    ];
    const { client, vault } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));
    assert.match((err as Error).message, /refusing non-SingleTenant/);
    assert.equal(vault.size, 0);
  });
});

describe('createAppRegistration — 403/429 typed errors', () => {
  it('403 on create → ConsentMissingError carrying Application.ReadWrite.OwnedBy', async () => {
    const routes: Route[] = [
      route('/applications', { status: 403, body: { error: 'forbidden' } }),
    ];
    const { client } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.ok(err instanceof ConsentMissingError);
    assert.deepEqual(err.missingScopes, [APP_REGISTRATION_SCOPE]);
    assert.equal(err.resource, 'graph');
  });

  it('exhausted 429 backoff → ProvisioningThrottledError with Retry-After', async () => {
    const routes: Route[] = [
      route('/applications', {
        status: 429,
        body: {},
        headers: { 'retry-after': '7' },
      }),
    ];
    const { client } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.ok(err instanceof ProvisioningThrottledError);
    assert.equal(err.retryAfterSeconds, 7);
    assert.equal(err.resource, 'graph');
  });
});

describe('createAppRegistration — rollback of partial failures', () => {
  it('addPassword fails → deletes the just-created app, vault untouched', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 500, body: { error: 'boom' } }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls, vault } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.match((err as Error).message, /applications\.addPassword 500/);
    assert.equal(vault.size, 0);
    const del = graphCalls(calls).find((c) => c.init?.method === 'DELETE');
    assert.ok(del, 'expected the rollback DELETE');
    assert.ok(del!.url.includes(`applications(appId='${APP_ID}')`));
  });

  it('service-principal failure → deletes app AND vault entry, rethrows original', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 500, body: { error: 'boom' } }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls, vault } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.match((err as Error).message, /servicePrincipals\.create 500/);
    assert.equal(vault.size, 0, 'rollback must remove the stored secret');
    assert.ok(graphCalls(calls).some((c) => c.init?.method === 'DELETE'));
  });

  it('pre-existing app is NOT deleted on rollback — only its new credential', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/removePassword', { status: 204 }),
      route('/servicePrincipals', { status: 500, body: { error: 'boom' } }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', { status: 409, body: {} }),
    ];
    const { client, calls, vault } = harness(routes);
    await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.equal(vault.size, 0);
    const graph = graphCalls(calls);
    assert.ok(
      !graph.some((c) => c.init?.method === 'DELETE'),
      'a found (pre-existing) registration must never be deleted by rollback',
    );
    const removed = graph.find((c) => c.url.includes('/removePassword'));
    assert.ok(removed, 'expected removePassword for the added credential');
    assert.deepEqual(bodyOf(removed!), { keyId: 'key-1' });
  });
});

describe('deleteAppRegistration — the idempotent rollback half', () => {
  it('204 → deleted, and removes the vault entry when a secretRef is passed', async () => {
    const routes: Route[] = [route("/applications(appId='", { status: 204 })];
    const { client, calls, vault } = harness(routes);
    vault.set(SECRET_REF, SECRET_TEXT);

    const result = await client.deleteAppRegistration({
      appId: APP_ID,
      secretRef: SECRET_REF,
    });

    assert.equal(result.outcome, 'deleted');
    assert.equal(vault.size, 0);
    const del = graphCalls(calls)[0]!;
    assert.equal(del.init?.method, 'DELETE');
    assert.ok(del.url.endsWith(`/applications(appId='${APP_ID}')`));
  });

  it('404 → already-deleted, not an error (safe to retry / double-rollback)', async () => {
    const routes: Route[] = [
      route("/applications(appId='", { status: 404, body: { error: 'gone' } }),
    ];
    const { client } = harness(routes);
    const result = await client.deleteAppRegistration({ appId: APP_ID });
    assert.equal(result.outcome, 'already-deleted');
  });

  it('403 → ConsentMissingError, vault entry is kept', async () => {
    const routes: Route[] = [
      route("/applications(appId='", { status: 403, body: { error: 'forbidden' } }),
    ];
    const { client, vault } = harness(routes);
    vault.set(SECRET_REF, SECRET_TEXT);

    const err = await rejection(
      client.deleteAppRegistration({ appId: APP_ID, secretRef: SECRET_REF }),
    );
    assert.ok(err instanceof ConsentMissingError);
    assert.equal(vault.get(SECRET_REF), SECRET_TEXT);
  });

  it('rejects an appId that could break the OData alternate-key URL', async () => {
    const { client } = harness([]);
    const err = await rejection(
      client.deleteAppRegistration({ appId: "x'/../" }),
    );
    assert.match((err as Error).message, /invalid_argument/);
  });
});

describe('getAppRegistration — status probe', () => {
  it('returns the parsed registration when the app exists', async () => {
    const routes: Route[] = [
      route("/applications(appId='", { status: 200, body: APP_BODY }),
    ];
    const { client } = harness(routes);
    const registration = await client.getAppRegistration(APP_ID, 'customer');
    assert.equal(registration?.appId, APP_ID);
    assert.equal(registration?.signInAudience, 'AzureADMyOrg');
    assert.equal(registration?.uniqueName, 'omadia-agent-hr');
  });

  it('returns undefined on 404', async () => {
    const routes: Route[] = [
      route("/applications(appId='", { status: 404, body: {} }),
    ];
    const { client } = harness(routes);
    assert.equal(await client.getAppRegistration(APP_ID, 'customer'), undefined);
  });
});

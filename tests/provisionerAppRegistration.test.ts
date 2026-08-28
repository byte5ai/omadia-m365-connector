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
  DirectoryReplicationError,
  ProvisioningThrottledError,
  UniqueNameReservedError,
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

function harness(
  routes: Route[],
  replicationOverrides: { maxAttempts?: number } = {},
): {
  client: AppRegistrationClient;
  calls: FetchCall[];
  vault: Map<string, string>;
  waits: number[];
  logs: string[];
} {
  const calls: FetchCall[] = [];
  const waits: number[] = [];
  const logs: string[] = [];
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
    log: (msg) => {
      logs.push(msg);
    },
    replication: {
      // Replication waits are recorded, never really slept.
      sleep: async (ms) => {
        waits.push(ms);
      },
      ...replicationOverrides,
    },
  });
  return { client, calls, vault, waits, logs };
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
    assert.equal(graph.length, 4);
    assert.ok(graph[0]?.url.endsWith('/v1.0/applications'));
    assert.equal(graph[0]?.init?.method, 'POST');
    // Between create and the first follow-up WRITE: the replication probe.
    assert.ok(graph[1]?.url.endsWith(`/applications/${OBJECT_ID}`));
    assert.equal(graph[1]?.init?.method, 'GET');
    assert.ok(graph[2]?.url.endsWith(`/applications/${OBJECT_ID}/addPassword`));
    assert.ok(graph[3]?.url.endsWith('/servicePrincipals'));
    assert.deepEqual(bodyOf(graph[3]!), { appId: APP_ID });
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

  it('rejects a uniqueName with URL/OData metacharacters before any network call', async () => {
    // The uniqueName lands in the OData alternate-key URL path; '#', '?',
    // '/', quotes and whitespace would silently truncate or reroute the
    // lookup (major review finding). Rejected up front, and the lookup URL
    // additionally percent-encodes the value.
    const { client, calls } = harness(HAPPY_ROUTES());
    for (const uniqueName of [
      'omadia agent#1',
      'agents/hr?x=1',
      "od'ata",
      'a+b&c',
    ]) {
      const err = await rejection(
        client.createAppRegistration({ ...CREATE_INPUT, uniqueName }),
      );
      assert.match((err as Error).message, /invalid_argument: 'uniqueName'/);
    }
    assert.equal(calls.length, 0, 'validation happens before any fetch');
  });

  it('409 without a uniqueName cannot be resolved and rejects', async () => {
    const routes: Route[] = [
      route('/applications', { status: 409, body: { error: 'conflict' } }),
    ];
    const { client } = harness(routes);
    const err = await rejection(
      client.createAppRegistration({ displayName: 'HR Agent', tenantMode: 'customer' }),
    );
    assert.match((err as Error).message, /conflict without a uniqueName/);
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
  // byte5ai/omadia#916: a 5xx is TRANSIENT. Deleting the app over one
  // soft-deletes it and reserves its uniqueName for 30 days, so the retry
  // collides with an object nobody can see. Nothing may be undone here.
  it('transient failure after create → NOTHING is rolled back', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 503, body: { error: 'boom' } }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls, vault, logs } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.match((err as Error).message, /servicePrincipals\.create 503/);
    assert.ok(
      !graphCalls(calls).some((c) => c.init?.method === 'DELETE'),
      'a transient failure must never delete the registration',
    );
    assert.ok(
      !graphCalls(calls).some((c) => c.url.includes('/removePassword')),
      'a transient failure must not revoke the credential either',
    );
    assert.equal(
      vault.get(SECRET_REF),
      SECRET_TEXT,
      'the stored secret matches a live credential — keep it for the re-run',
    );
    assert.ok(logs.some((l) => l.includes('leaving it in place')));
  });

  it('replication budget exhausted → transient, still no rollback', async () => {
    const routes: Route[] = [
      route('/addPassword', {
        status: 404,
        body: { error: { code: 'Request_ResourceNotFound' } },
      }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls } = harness(routes, { maxAttempts: 3 });
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.ok(err instanceof DirectoryReplicationError);
    assert.equal(err.step, 'applications.addPassword');
    assert.equal(err.objectId, OBJECT_ID);
    assert.ok(
      !graphCalls(calls).some((c) => c.init?.method === 'DELETE'),
      'the app exists and is adoptable — deleting it burns the uniqueName',
    );
  });

  it('non-transient failure keeps a registration that carries a uniqueName', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 400, body: { error: 'nope' } }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls, vault, logs } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.match((err as Error).message, /servicePrincipals\.create 400/);
    assert.equal(vault.size, 0, 'the vault entry this call wrote is rolled back');
    assert.ok(
      !graphCalls(calls).some((c) => c.init?.method === 'DELETE'),
      'deleting would reserve the uniqueName for 30 days',
    );
    assert.ok(logs.some((l) => l.includes('would reserve that name')));
  });

  it('non-transient failure DOES delete an app that carries no uniqueName', async () => {
    // Nothing could ever find this object again — an orphan, not an
    // adoptable identity. That is the one case rollback still deletes.
    const anonymous = { ...APP_BODY, uniqueName: undefined };
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 400, body: { error: 'nope' } }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: anonymous }),
    ];
    const { client, calls, vault } = harness(routes);
    await rejection(
      client.createAppRegistration({
        displayName: 'HR Agent',
        tenantMode: 'customer',
      }),
    );

    assert.equal(vault.size, 0);
    const del = graphCalls(calls).find((c) => c.init?.method === 'DELETE');
    assert.ok(del, 'expected the rollback DELETE');
    assert.ok(del!.url.includes(`applications(appId='${APP_ID}')`));
  });

  it('restores an OVERWRITTEN vault entry on rollback instead of deleting it', async () => {
    // Re-run over an existing registration: run 1 stored password P1 at the
    // deterministic key; this call overwrites it with P2, then fails at the
    // service principal. The rollback must put P1 back — deleting the entry
    // would destroy a credential this call did not create (blocker finding
    // of the app-registration unit review).
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/removePassword', { status: 204 }),
      route('/servicePrincipals', { status: 400, body: { error: 'boom' } }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', { status: 409, body: {} }),
    ];
    const { client, vault } = harness(routes);
    const PRIOR_SECRET = 'prior-run-password-p1';
    vault.set(SECRET_REF, PRIOR_SECRET);

    await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.equal(
      vault.get(SECRET_REF),
      PRIOR_SECRET,
      'the pre-call vault value must be restored, not deleted',
    );
  });

  it('pre-existing app is NOT deleted on rollback — only its new credential', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/removePassword', { status: 204 }),
      route('/servicePrincipals', { status: 400, body: { error: 'boom' } }),
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

// ---------------------------------------------------------------------------
// byte5ai/omadia#916 — the first real provisioning run against the byte5
// tenant. Every case below reproduces one link of the chain that burned a
// slug for 30 days: create 201 → addPassword 404 (replication) → rollback →
// soft-delete → uniqueName reserved → every retry collides.
// ---------------------------------------------------------------------------

/** Entra's answer to a taken uniqueName: 400, NOT 409. */
const UNIQUE_NAME_TAKEN = {
  status: 400,
  body: {
    error: {
      code: 'Request_BadRequest',
      message:
        'Another object with the same value for property uniqueName already exists.',
    },
  },
};

/** The same collision as Graph sometimes codes it. */
const OBJECT_CONFLICT = {
  status: 400,
  body: {
    error: { code: 'ObjectConflict', message: 'conflicting object exists' },
  },
};

const NOT_REPLICATED_YET = {
  status: 404,
  body: {
    error: {
      code: 'Request_ResourceNotFound',
      message: `Resource '${OBJECT_ID}' does not exist or one of its queried reference-property objects are not present.`,
    },
  },
};

describe('createAppRegistration — Entra replication window (#916)', () => {
  it('create 201 → addPassword 404 → polls → 200, no rollback', async () => {
    const routes: Route[] = [
      route('/addPassword', NOT_REPLICATED_YET, { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 201, body: { id: 'sp-1' } }),
      route("/applications(appId='", { status: 204 }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls, vault, waits } = harness(routes);
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'created');
    assert.equal(result.value.appId, APP_ID);
    assert.equal(vault.get(SECRET_REF), SECRET_TEXT);
    assert.equal(
      graphCalls(calls).filter((c) => c.url.includes('/addPassword')).length,
      2,
      'the 404 is retried in place, not reported as a step failure',
    );
    assert.deepEqual(waits, [1000], 'one replication wait before the retry');
    assert.ok(!graphCalls(calls).some((c) => c.init?.method === 'DELETE'));
  });

  it('waits for the created app to become addressable before writing to it', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 201, body: { id: 'sp-1' } }),
      route(
        `/applications/${OBJECT_ID}`,
        { status: 404, body: {} },
        { status: 200, body: APP_BODY },
      ),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls, waits } = harness(routes);
    await client.createAppRegistration(CREATE_INPUT);

    const probes = graphCalls(calls).filter(
      (c) => c.init?.method === 'GET' && c.url.endsWith(`/applications/${OBJECT_ID}`),
    );
    assert.equal(probes.length, 2, 'polled until the object was readable');
    assert.deepEqual(waits, [1000]);
  });

  it('adopting an existing app skips the replication probe', async () => {
    const routes: Route[] = [
      route('passwordCredentials', { status: 200, body: { passwordCredentials: [] } }),
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 409, body: {} }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', UNIQUE_NAME_TAKEN),
    ];
    const { client, calls, waits } = harness(routes);
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'already-existed');
    assert.deepEqual(waits, [], 'a found app is addressable by definition');
  });
});

describe('createAppRegistration — uniqueName conflict is a 400 (#916)', () => {
  for (const [label, spec] of [
    ['Request_BadRequest + uniqueName message', UNIQUE_NAME_TAKEN],
    ['ObjectConflict code', OBJECT_CONFLICT],
  ] as const) {
    it(`400 ${label} → adopts the live registration`, async () => {
      const routes: Route[] = [
        route('passwordCredentials', {
          status: 200,
          body: { passwordCredentials: [] },
        }),
        route('/addPassword', { status: 200, body: PASSWORD_BODY }),
        route('/servicePrincipals', { status: 409, body: {} }),
        route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
        route('/applications', spec),
      ];
      const { client, calls, vault } = harness(routes);
      const result = await client.createAppRegistration(CREATE_INPUT);

      assert.equal(result.outcome, 'already-existed');
      assert.equal(result.value.appId, APP_ID);
      // A fresh secret is minted — the original was never persisted.
      assert.equal(vault.get(SECRET_REF), SECRET_TEXT);
      assert.ok(
        graphCalls(calls).some((c) =>
          c.url.includes("applications(uniqueName='omadia-agent-hr')"),
        ),
        'expected the alternate-key lookup',
      );
    });
  }

  it('an unrelated 400 stays an error — the rules are not a catch-all', async () => {
    const routes: Route[] = [
      route('/applications', {
        status: 400,
        body: {
          error: {
            code: 'Request_BadRequest',
            message: 'Property displayName is invalid.',
          },
        },
      }),
    ];
    const { client, calls } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.match((err as Error).message, /applications\.create 400/);
    assert.ok(
      !graphCalls(calls).some((c) => c.url.includes("(uniqueName='")),
      'no adoption attempt for a genuine bad request',
    );
  });
});

describe('createAppRegistration — the recycle bin (#916)', () => {
  const DELETED_ITEM = {
    id: 'del-0001',
    appId: APP_ID,
    displayName: 'HR Agent',
    uniqueName: 'omadia-agent-hr',
    deletedDateTime: '2026-08-28T10:19:32Z',
  };

  it('soft-deleted holder → restores it and adopts the registration', async () => {
    const routes: Route[] = [
      route('passwordCredentials', { status: 200, body: { passwordCredentials: [] } }),
      route('/restore', { status: 200, body: APP_BODY }),
      route('/deletedItems/microsoft.graph.application', {
        status: 200,
        body: { value: [DELETED_ITEM] },
      }),
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 409, body: {} }),
      route(
        "/applications(uniqueName='",
        { status: 404, body: {} },
        { status: 200, body: APP_BODY },
      ),
      route('/applications', UNIQUE_NAME_TAKEN),
    ];
    const { client, calls } = harness(routes);
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'already-existed');
    assert.equal(result.value.appId, APP_ID);
    const restore = graphCalls(calls).find((c) => c.url.includes('/restore'));
    assert.ok(restore, 'expected the recycle-bin restore');
    assert.equal(restore!.init?.method, 'POST');
    assert.ok(restore!.url.includes('deletedItems/del-0001/restore'));
  });

  it('restore fails → an error that names the object and the 30-day window', async () => {
    const routes: Route[] = [
      route('/restore', { status: 400, body: { error: 'nope' } }),
      route('/deletedItems/microsoft.graph.application', {
        status: 200,
        body: { value: [DELETED_ITEM] },
      }),
      route("/applications(uniqueName='", { status: 404, body: {} }),
      route('/applications', UNIQUE_NAME_TAKEN),
    ];
    const { client } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.ok(err instanceof UniqueNameReservedError);
    assert.equal(err.uniqueName, 'omadia-agent-hr');
    assert.equal(err.deletedObjectId, 'del-0001');
    assert.equal(err.deletedDateTime, '2026-08-28T10:19:32Z');
    assert.equal(err.retentionDays, 30);
    assert.match(err.message, /30 days/);
    assert.match(err.message, /soft-deleted application 'del-0001'/);
  });

  it('recycle bin unreadable → still explains WHY the name is taken', async () => {
    // Application.ReadWrite.OwnedBy does not cover directory/deletedItems.
    // The probe degrades to a diagnostic; it must never escalate into a
    // consent error for the whole step.
    const routes: Route[] = [
      route('/deletedItems/microsoft.graph.application', {
        status: 403,
        body: { error: 'forbidden' },
      }),
      route("/applications(uniqueName='", { status: 404, body: {} }),
      route('/applications', UNIQUE_NAME_TAKEN),
    ];
    const { client } = harness(routes);
    const err = await rejection(client.createAppRegistration(CREATE_INPUT));

    assert.ok(err instanceof UniqueNameReservedError);
    assert.ok(!(err instanceof ConsentMissingError));
    assert.match(err.message, /30 days/);
    assert.match(err.message, /recycle bin/);
    assert.match(err.message, /Application\.ReadWrite\.All/);
  });
});

describe('createAppRegistration — app_id is persistable before anything else (#916)', () => {
  it('onRegistrationCreated fires after create and BEFORE addPassword', async () => {
    const routes: Route[] = [
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 201, body: { id: 'sp-1' } }),
      route('/applications', { status: 201, body: APP_BODY }),
    ];
    const { client, calls } = harness(routes);
    let seenWhenNotified: string[] = [];
    const result = await client.createAppRegistration({
      ...CREATE_INPUT,
      onRegistrationCreated: (registration, outcome) => {
        assert.equal(registration.appId, APP_ID);
        assert.equal(outcome, 'created');
        seenWhenNotified = graphCalls(calls).map((c) => c.url);
      },
    });

    assert.equal(result.value.appId, APP_ID);
    assert.equal(seenWhenNotified.length, 1, 'notified right after the create');
    assert.ok(!seenWhenNotified.some((u) => u.includes('/addPassword')));
  });

  it('an interrupted run leaves a resumable app_id, and the re-run adopts it', async () => {
    // Run 1: the app is created, then the chain dies in the replication
    // window. The caller has already persisted app_id.
    let persistedAppId: string | undefined;
    const onRegistrationCreated = (registration: { appId: string }): void => {
      persistedAppId = registration.appId;
    };

    const run1 = harness(
      [
        route('/addPassword', NOT_REPLICATED_YET),
        route("/applications(appId='", { status: 204 }),
        route('/applications', { status: 201, body: APP_BODY }),
      ],
      { maxAttempts: 2 },
    );
    await rejection(
      run1.client.createAppRegistration({ ...CREATE_INPUT, onRegistrationCreated }),
    );
    assert.equal(persistedAppId, APP_ID, 'app_id survives the interruption');
    assert.ok(
      !graphCalls(run1.calls).some((c) => c.init?.method === 'DELETE'),
      'the app stays in the tenant — that is what makes the re-run possible',
    );

    // Run 2: create collides with the LIVE app from run 1 and adopts it.
    const run2 = harness([
      route('passwordCredentials', { status: 200, body: { passwordCredentials: [] } }),
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 409, body: {} }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', UNIQUE_NAME_TAKEN),
    ]);
    const result = await run2.client.createAppRegistration({
      ...CREATE_INPUT,
      onRegistrationCreated,
    });

    assert.equal(result.outcome, 'already-existed');
    assert.equal(result.value.appId, persistedAppId);
    assert.equal(run2.vault.get(SECRET_REF), SECRET_TEXT);
  });

  it('a throwing persistence hook is logged, not fatal', async () => {
    const { client, logs } = harness(HAPPY_ROUTES());
    const result = await client.createAppRegistration({
      ...CREATE_INPUT,
      onRegistrationCreated: () => {
        throw new Error('store unavailable');
      },
    });

    assert.equal(result.value.appId, APP_ID);
    assert.ok(logs.some((l) => l.includes('onRegistrationCreated')));
  });
});

describe('createAppRegistration — superseded credentials on adoption (#916)', () => {
  it('removes only the provisioner-labelled predecessors', async () => {
    const routes: Route[] = [
      route('passwordCredentials', {
        status: 200,
        body: {
          passwordCredentials: [
            { keyId: 'key-1', displayName: 'HR Agent bot password' },
            { keyId: 'key-0', displayName: 'HR Agent bot password' },
            { keyId: 'key-operator', displayName: 'ops laptop' },
          ],
        },
      }),
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/removePassword', { status: 204 }),
      route('/servicePrincipals', { status: 409, body: {} }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', UNIQUE_NAME_TAKEN),
    ];
    const { client, calls } = harness(routes);
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'already-existed');
    const removed = graphCalls(calls).filter((c) =>
      c.url.includes('/removePassword'),
    );
    assert.equal(removed.length, 1, 'exactly the one superseded credential');
    assert.deepEqual(bodyOf(removed[0]!), { keyId: 'key-0' });
  });

  it('a failing cleanup never fails the step', async () => {
    const routes: Route[] = [
      route('passwordCredentials', { status: 500, body: { error: 'boom' } }),
      route('/addPassword', { status: 200, body: PASSWORD_BODY }),
      route('/servicePrincipals', { status: 409, body: {} }),
      route("/applications(uniqueName='", { status: 200, body: APP_BODY }),
      route('/applications', UNIQUE_NAME_TAKEN),
    ];
    const { client, vault } = harness(routes);
    const result = await client.createAppRegistration(CREATE_INPUT);

    assert.equal(result.outcome, 'already-existed');
    assert.equal(vault.get(SECRET_REF), SECRET_TEXT);
  });
});

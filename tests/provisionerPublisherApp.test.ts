import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  APP_CATALOG_DELEGATED_PERMISSION_ID,
  CHAT_READ_DELEGATED_PERMISSION_ID,
  GRAPH_RESOURCE_APP_ID,
} from '../src/teamsProvisioner/delegatedAuth.js';
import { UniqueNameReservedError } from '../src/teamsProvisioner/errors.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  PUBLISHER_APP_UNIQUE_NAME_PREFIX,
  PublisherAppClient,
  declaresCatalogScope,
  declaresChatScope,
  mergeCatalogScope,
  publisherAppUniqueName,
} from '../src/teamsProvisioner/publisherApp.js';

// The public-client publisher app: the ONE registration the delegated catalog
// upload signs in as.
//
// Its whole reason for existing is that the device code grant needs a public
// client, and the connector's own app — which holds Application.ReadWrite.OwnedBy
// — must never become one. So this app carries no secret, exactly one delegated
// scope, and can do nothing on its own.
//
// The idempotency machinery is deliberately the app-registration step's
// (byte5ai/omadia#916), not a second implementation: a taken uniqueName is
// adopted, replication is polled through, a soft-deleted holder is restored, and
// NOTHING is ever rolled back — a delete would reserve the name for 30 days and
// lock the tenant out of signing in.

const TENANT_ID = '99999999-8888-7777-6666-555555555555';
const OBJECT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const APP_ID = '11111111-2222-3333-4444-555555555555';
const UNIQUE_NAME = `${PUBLISHER_APP_UNIQUE_NAME_PREFIX}-${TENANT_ID}`;

interface FetchCall {
  url: string;
  method: string;
  body: unknown;
}

interface ResponseSpec {
  status: number;
  body?: unknown;
}

interface Route {
  match: string;
  method?: string;
  responses: ResponseSpec[];
}

const route = (
  match: string,
  method: string | undefined,
  ...responses: ResponseSpec[]
): Route => ({ match, ...(method !== undefined ? { method } : {}), responses });

function makeResponse(spec: ResponseSpec): Response {
  const text = spec.body === undefined ? '' : JSON.stringify(spec.body);
  return {
    ok: spec.status >= 200 && spec.status < 300,
    status: spec.status,
    headers: { get: () => null },
    json: async () => {
      if (text === '') throw new Error('no body');
      return JSON.parse(text);
    },
    text: async () => text,
  } as unknown as Response;
}

function mockFetch(routes: Route[], calls: FetchCall[]): typeof fetch {
  const queues = routes.map((r) => ({ ...r, next: 0 }));
  return (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    const url = String(input);
    const method = init?.method ?? 'GET';
    if (url.includes('login.microsoftonline.com')) {
      return makeResponse({
        status: 200,
        body: { access_token: 'tok', expires_in: 3600 },
      });
    }
    let body: unknown;
    try {
      body = init?.body === undefined ? undefined : JSON.parse(String(init.body));
    } catch {
      body = undefined;
    }
    calls.push({ url, method, body });
    for (const q of queues) {
      if (!url.includes(q.match)) continue;
      if (q.method !== undefined && q.method !== method) continue;
      const spec = q.responses[Math.min(q.next, q.responses.length - 1)];
      q.next += 1;
      if (spec) return makeResponse(spec);
    }
    return makeResponse({ status: 500, body: { error: `unrouted ${method} ${url}` } });
  }) as typeof fetch;
}

/** A Graph `application` as the publisher app should look once correct. */
function applicationBody(
  overrides: Record<string, unknown> = {},
): Record<string, unknown> {
  return {
    id: OBJECT_ID,
    appId: APP_ID,
    displayName: 'omadia Teams App Publisher',
    signInAudience: 'AzureADMyOrg',
    uniqueName: UNIQUE_NAME,
    isFallbackPublicClient: true,
    requiredResourceAccess: [
      {
        resourceAppId: GRAPH_RESOURCE_APP_ID,
        resourceAccess: [
          { id: APP_CATALOG_DELEGATED_PERMISSION_ID, type: 'Scope' },
          // 0.8.0 — the publisher app also declares Chat.ReadBasic so the
          // admin-consent URL can grant it. An app WITHOUT this is the
          // pre-0.8.0 shape; the repair path is covered separately below.
          { id: CHAT_READ_DELEGATED_PERMISSION_ID, type: 'Scope' },
        ],
      },
    ],
    ...overrides,
  };
}

/** Entra reports a taken uniqueName as 400 Request_BadRequest, not 409. */
const UNIQUE_NAME_TAKEN: ResponseSpec = {
  status: 400,
  body: {
    error: {
      code: 'Request_BadRequest',
      message:
        'Another object with the same value for property uniqueName already exists.',
    },
  },
};

function harness(routes: Route[]): {
  client: PublisherAppClient;
  calls: FetchCall[];
  sleeps: number[];
} {
  const calls: FetchCall[] = [];
  const sleeps: number[] = [];
  const http = new ProvisioningHttp({
    graphCredential: {
      tenantId: TENANT_ID,
      clientId: 'connector-app',
      clientSecret: 'connector-secret',
    },
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async (ms) => {
      sleeps.push(ms);
    },
  });
  const client = new PublisherAppClient({
    http,
    tenantId: TENANT_ID,
    log: () => {},
    replication: {
      sleep: async (ms) => {
        sleeps.push(ms);
      },
    },
  });
  return { client, calls, sleeps };
}

const FIND = "/applications(uniqueName='";

describe('publisherAppUniqueName', () => {
  it('qualifies the key by tenant and stays inside Graph limits', () => {
    const name = publisherAppUniqueName(TENANT_ID);
    assert.equal(name, UNIQUE_NAME);
    assert.ok(name.length <= 64, 'application.uniqueName is capped at 64 chars');
  });

  it('rejects a tenant id that would break the OData alternate-key URL', () => {
    assert.throws(() => publisherAppUniqueName("tenant'; DROP"), /invalid_argument/);
  });
});

describe('PublisherAppClient.ensurePublisherApp — first run', () => {
  it('registers a secret-less public client with exactly the delegated catalog scope', async () => {
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', {
        status: 201,
        body: applicationBody(),
      }),
      route(`/applications/${OBJECT_ID}`, 'GET', {
        status: 200,
        body: applicationBody(),
      }),
      route('/servicePrincipals', 'POST', { status: 201, body: { id: 'sp-1' } }),
    ]);

    const result = await client.ensurePublisherApp();

    assert.equal(result.outcome, 'created');
    assert.equal(result.value.appId, APP_ID);
    assert.equal(result.value.isPublicClient, true);
    assert.equal(result.value.declaresCatalogScope, true);
    assert.equal(result.value.declaresChatScope, true);

    const create = calls.find((c) => c.method === 'POST' && c.url.endsWith('/applications'));
    assert.ok(create);
    const body = create.body as Record<string, unknown>;
    // The public-client flag is what makes the device code grant possible at
    // all; without it Entra answers unauthorized_client forever.
    assert.equal(body['isFallbackPublicClient'], true);
    assert.equal(body['signInAudience'], 'AzureADMyOrg');
    assert.equal(body['uniqueName'], UNIQUE_NAME);
    // No redirect URI is registered — that is the point of choosing the device
    // code grant for a self-hosted product with a different URL per install.
    assert.deepEqual(body['publicClient'], { redirectUris: [] });
    // Two permissions since 0.8.0 — catalog publish and the chat read behind
    // the target picker — and BOTH must be DELEGATED ('Scope'). 'Role' would
    // be an application permission, which Graph refuses for the publish verb
    // and does not offer tenant-wide for chats at all.
    assert.deepEqual(body['requiredResourceAccess'], [
      {
        resourceAppId: GRAPH_RESOURCE_APP_ID,
        resourceAccess: [
          { id: APP_CATALOG_DELEGATED_PERMISSION_ID, type: 'Scope' },
          { id: CHAT_READ_DELEGATED_PERMISSION_ID, type: 'Scope' },
        ],
      },
    ]);

    // No client secret is ever minted for this app — that is what makes a
    // public-client flag acceptable on it.
    assert.equal(
      calls.filter((c) => c.url.includes('addPassword')).length,
      0,
      'the publisher app must never get a client secret',
    );
    assert.ok(calls.some((c) => c.url.endsWith('/servicePrincipals')));
  });

  it('polls through the Entra replication window instead of failing', async () => {
    // POST /applications answers 201 and the very next read can still 404 for
    // a few seconds. That is a replication window, not a failure
    // (byte5ai/omadia#916).
    const { client, sleeps } = harness([
      route('/v1.0/applications', 'POST', { status: 201, body: applicationBody() }),
      route(`/applications/${OBJECT_ID}`, 'GET',
        { status: 404 },
        { status: 404 },
        { status: 200, body: applicationBody() },
      ),
      route('/servicePrincipals', 'POST', { status: 201, body: {} }),
    ]);

    const result = await client.ensurePublisherApp();
    assert.equal(result.outcome, 'created');
    assert.deepEqual(sleeps, [1000, 2000]);
  });

  it('treats an existing service principal (409) as success', async () => {
    const { client } = harness([
      route('/v1.0/applications', 'POST', { status: 201, body: applicationBody() }),
      route(`/applications/${OBJECT_ID}`, 'GET', { status: 200, body: applicationBody() }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);
    const result = await client.ensurePublisherApp();
    assert.equal(result.outcome, 'created');
  });
});

describe('PublisherAppClient.ensurePublisherApp — adopt path', () => {
  it('adopts the live app behind a taken uniqueName without creating a duplicate', async () => {
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET', { status: 200, body: applicationBody() }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);

    const result = await client.ensurePublisherApp();

    assert.equal(result.outcome, 'already-existed');
    assert.equal(result.value.appId, APP_ID);
    assert.equal(
      calls.filter((c) => c.method === 'POST' && c.url.endsWith('/applications')).length,
      1,
      'exactly one create attempt, then adopt',
    );
    // Nothing is deleted on this path. A rollback here would soft-delete the
    // app and reserve its name for 30 days, locking the tenant out of the
    // sign-in it needs to publish anything at all.
    assert.equal(calls.filter((c) => c.method === 'DELETE').length, 0);
  });

  it('repairs an adopted app that is missing the public-client flag', async () => {
    // An app registered by an older connector version — or edited in the
    // portal — fails at sign-in time with a protocol error that names nothing
    // useful. Repairing here turns that into a non-event.
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET', {
        status: 200,
        body: applicationBody({ isFallbackPublicClient: false }),
      }),
      route(`/applications/${OBJECT_ID}?$select=requiredResourceAccess`, 'GET', {
        status: 200,
        body: { requiredResourceAccess: applicationBody()['requiredResourceAccess'] },
      }),
      route(`/applications/${OBJECT_ID}`, 'PATCH', { status: 204 }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);

    const result = await client.ensurePublisherApp();

    assert.equal(result.value.isPublicClient, true);
    const patch = calls.find((c) => c.method === 'PATCH');
    assert.ok(patch, 'expected a repair PATCH');
    assert.equal((patch.body as Record<string, unknown>)['isFallbackPublicClient'], true);
  });

  it('adds the delegated scope to an adopted app WITHOUT dropping other permissions', async () => {
    // PATCHing requiredResourceAccess is a full overwrite. Silently deleting a
    // permission somebody added on purpose would be a worse failure than the
    // one being fixed, so the write merges.
    const foreign = {
      resourceAppId: '00000002-0000-0000-c000-000000000000',
      resourceAccess: [{ id: 'some-other-permission', type: 'Scope' }],
    };
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET', {
        status: 200,
        body: applicationBody({ requiredResourceAccess: [foreign] }),
      }),
      route(`/applications/${OBJECT_ID}?$select=requiredResourceAccess`, 'GET', {
        status: 200,
        body: { requiredResourceAccess: [foreign] },
      }),
      route(`/applications/${OBJECT_ID}`, 'PATCH', { status: 204 }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);

    await client.ensurePublisherApp();

    const patch = calls.find((c) => c.method === 'PATCH');
    assert.ok(patch);
    const written = (patch.body as Record<string, unknown>)[
      'requiredResourceAccess'
    ] as unknown[];
    assert.equal(written.length, 2, 'the foreign permission must survive');
    assert.ok(declaresCatalogScope(written));
    assert.ok(written.some((e) => JSON.stringify(e) === JSON.stringify(foreign)));
  });

  it('repairs a pre-0.8.0 publisher app by declaring Chat.ReadBasic', async () => {
    // THE UPGRADE PATH. Every tenant that signed in before 0.8.0 has a
    // publisher app declaring only the catalog scope. Declaring is not
    // consenting — but the admin-consent URL grants what the app DECLARES, so
    // a tenant with user consent switched off could never consent to a scope
    // that is missing here, no matter how often the admin signs in.
    const catalogOnly = {
      resourceAppId: GRAPH_RESOURCE_APP_ID,
      resourceAccess: [
        { id: APP_CATALOG_DELEGATED_PERMISSION_ID, type: 'Scope' },
      ],
    };
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET', {
        status: 200,
        body: applicationBody({ requiredResourceAccess: [catalogOnly] }),
      }),
      route(`/applications/${OBJECT_ID}?$select=requiredResourceAccess`, 'GET', {
        status: 200,
        body: { requiredResourceAccess: [catalogOnly] },
      }),
      route(`/applications/${OBJECT_ID}`, 'PATCH', { status: 204 }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);

    const result = await client.ensurePublisherApp();

    const patch = calls.find((c) => c.method === 'PATCH');
    assert.ok(patch, 'a pre-0.8.0 app must be repaired, not left as it is');
    const written = (patch.body as Record<string, unknown>)[
      'requiredResourceAccess'
    ] as unknown[];
    // Both scopes present, and the catalog one is not disturbed on the way in.
    assert.ok(declaresCatalogScope(written));
    assert.ok(declaresChatScope(written));
    assert.equal(result.value.declaresChatScope, true);
  });

  it('does not PATCH an adopted app that is already correct', async () => {
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET', { status: 200, body: applicationBody() }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);

    await client.ensurePublisherApp();
    assert.equal(calls.filter((c) => c.method === 'PATCH').length, 0);
  });

  it('restores a soft-deleted app that still holds the uniqueName', async () => {
    // A deleted Entra app keeps its uniqueName reserved for 30 days while
    // being invisible in GET /applications. Restoring returns the SAME object,
    // which is exactly what re-provisioning should yield.
    const { client, calls } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET',
        { status: 404 },
        { status: 200, body: applicationBody() },
      ),
      route('/directory/deletedItems/microsoft.graph.application', 'GET', {
        status: 200,
        body: {
          value: [
            {
              id: OBJECT_ID,
              appId: APP_ID,
              uniqueName: UNIQUE_NAME,
              deletedDateTime: '2026-08-01T00:00:00Z',
            },
          ],
        },
      }),
      route(`/directory/deletedItems/${OBJECT_ID}/restore`, 'POST', {
        status: 200,
        body: applicationBody(),
      }),
      route('/servicePrincipals', 'POST', { status: 409, body: {} }),
    ]);

    const result = await client.ensurePublisherApp();

    assert.equal(result.outcome, 'already-existed');
    assert.equal(result.value.appId, APP_ID);
    assert.ok(calls.some((c) => c.url.includes('/restore')));
  });

  it('explains a reserved uniqueName it cannot recover, instead of a bare conflict', async () => {
    const { client } = harness([
      route('/v1.0/applications', 'POST', UNIQUE_NAME_TAKEN),
      route(FIND, 'GET', { status: 404 }),
      route('/directory/deletedItems/microsoft.graph.application', 'GET', {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);

    await assert.rejects(
      () => client.ensurePublisherApp(),
      (err: unknown) => {
        assert.ok(err instanceof UniqueNameReservedError);
        assert.equal(err.uniqueName, UNIQUE_NAME);
        assert.match(err.message, /30 days/);
        return true;
      },
    );
  });
});

describe('requiredResourceAccess merge helpers', () => {
  it('detects the delegated scope and ignores the Role spelling', () => {
    assert.ok(declaresCatalogScope(applicationBody()['requiredResourceAccess']));
    // 'Role' is an APPLICATION permission — present, but useless for a verb
    // Graph only supports delegated, so it must not count as declared.
    assert.ok(
      !declaresCatalogScope([
        {
          resourceAppId: GRAPH_RESOURCE_APP_ID,
          resourceAccess: [{ id: APP_CATALOG_DELEGATED_PERMISSION_ID, type: 'Role' }],
        },
      ]),
    );
    assert.ok(!declaresCatalogScope(undefined));
    assert.ok(!declaresCatalogScope([]));
  });

  it('is a no-op when the scope is already declared', () => {
    const existing = applicationBody()['requiredResourceAccess'] as unknown[];
    assert.deepEqual(mergeCatalogScope(existing), existing);
  });

  it('appends to an existing Graph entry rather than creating a second one', () => {
    const merged = mergeCatalogScope([
      {
        resourceAppId: GRAPH_RESOURCE_APP_ID,
        resourceAccess: [{ id: 'other-scope', type: 'Scope' }],
      },
    ]);
    assert.equal(merged.length, 1, 'one entry per resourceAppId');
    assert.ok(declaresCatalogScope(merged));
  });
});

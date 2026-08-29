import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { SecretsAccessor } from '@omadia/plugin-api';

import {
  APP_REGISTRATION_SCOPE,
  AppRegistrationClient,
} from '../src/teamsProvisioner/appRegistration.js';
import { CatalogUploadClient } from '../src/teamsProvisioner/catalog.js';
import {
  APP_CATALOG_DELEGATED_SCOPE,
  type DelegatedTokenSet,
} from '../src/teamsProvisioner/delegatedAuth.js';
import {
  ConsentMissingError,
  DelegatedScopeRequiredError,
  DelegatedTokenExpiredError,
  DeletedObjectIdMismatchError,
} from '../src/teamsProvisioner/errors.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';

// RESETTING A FAILED PROVISIONING RUN (0.8.0). A half-finished chain leaves
// leftovers that block a clean retry; these are the primitives that clear
// them. What the tests must prove:
//
//   * EVERY operation is idempotent in BOTH directions. A reset whose second
//     pass fails on the work its first pass already did is worthless — that is
//     the entire reason `'already-absent'` is a success and not an error;
//
//   * THE PURGE IS ADDRESSED BY OBJECT ID. `DELETE /directory/deletedItems/{id}`
//     takes the directory object id, not the appId every other rollback step
//     is keyed by. Both are GUIDs, so the wrong one does not fail loudly: it
//     404s exactly like an entry that was genuinely purged. Reporting that as
//     "already absent" would tell an operator the uniqueName is free while its
//     30-day reservation still stands — the confusion that cost a slug a month
//     in byte5ai/omadia#916. The guard is tested here explicitly;
//
//   * the catalog removal is DELEGATED, because Graph documents application
//     permissions for that verb as "Not supported." — same as the upload.
//
// deleteAppRegistration and deleteBot already existed as rollback halves and
// are reused unchanged; their idempotency lives in their own suites.

interface FetchCall {
  url: string;
  init?: RequestInit;
}

interface ResponseSpec {
  status: number;
  body?: unknown;
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
  return {
    ok: spec.status >= 200 && spec.status < 300,
    status: spec.status,
    headers: { get: () => null },
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
    if (url.includes('login.microsoftonline.com')) {
      return makeResponse({
        status: 200,
        body: { access_token: 'app-only-tok', expires_in: 3600 },
      });
    }
    calls.push({ url, ...(init ? { init } : {}) });
    for (const q of queues) {
      if (url.includes(q.match)) {
        const spec = q.responses[Math.min(q.next, q.responses.length - 1)];
        q.next += 1;
        if (spec) return makeResponse(spec);
      }
    }
    return makeResponse({ status: 500, body: { error: `unrouted ${url}` } });
  }) as typeof fetch;
}

const TENANT_ID = '99999999-8888-7777-6666-555555555555';
const OBJECT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const APP_ID = '11111111-2222-3333-4444-555555555555';
const DELETED_ITEMS = '/directory/deletedItems';
const RECYCLE_BIN_LIST = '/directory/deletedItems/microsoft.graph.application';

function fakeSecrets(): SecretsAccessor {
  const vault = new Map<string, string>();
  return {
    get: async (key: string) => vault.get(key),
    set: async (key: string, value: string) => {
      vault.set(key, value);
    },
    delete: async (key: string) => {
      vault.delete(key);
    },
  } as unknown as SecretsAccessor;
}

function http(routes: Route[], calls: FetchCall[]): ProvisioningHttp {
  return new ProvisioningHttp({
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
}

function appHarness(routes: Route[]): {
  client: AppRegistrationClient;
  calls: FetchCall[];
} {
  const calls: FetchCall[] = [];
  return {
    client: new AppRegistrationClient({
      http: http(routes, calls),
      secrets: fakeSecrets(),
      tenantId: TENANT_ID,
      log: () => {},
      replication: { sleep: async () => {} },
    }),
    calls,
  };
}

function catalogHarness(routes: Route[]): {
  client: CatalogUploadClient;
  calls: FetchCall[];
} {
  const calls: FetchCall[] = [];
  return {
    client: new CatalogUploadClient({
      http: http(routes, calls),
      log: () => {},
      delegatedAuth: {
        ensureFreshToken: async (t) => ({ tokens: t, refreshed: false }),
        adminConsentUrlFor: () => 'https://consent.example/admin',
      },
    }),
    calls,
  };
}

function tokens(overrides: Partial<DelegatedTokenSet> = {}): DelegatedTokenSet {
  return {
    accessToken: 'delegated-access-token',
    refreshToken: 'delegated-refresh-token',
    expiresAt: new Date(Date.now() + 30 * 60_000).toISOString(),
    scopes: [APP_CATALOG_DELEGATED_SCOPE, 'offline_access'],
    clientId: 'publisher-app',
    tenantId: TENANT_ID,
    ...overrides,
  };
}

describe('AppRegistrationClient.deleteAppRegistration (reused rollback half)', () => {
  it('deletes by appId and reports the idempotent outcome', async () => {
    const { client, calls } = appHarness([
      route(`/applications(appId='${APP_ID}')`, { status: 204 }),
    ]);
    assert.deepEqual(await client.deleteAppRegistration({ appId: APP_ID }), {
      outcome: 'deleted',
    });
    // The DELETE direction is keyed by appId — which is exactly why the purge
    // below must not be, and why mixing them up is so easy.
    assert.ok(calls[0]?.url.includes(`(appId='${APP_ID}')`));
    assert.equal(calls[0]?.init?.method, 'DELETE');
  });

  it('answers already-deleted on 404 so a second reset pass succeeds', async () => {
    const { client } = appHarness([
      route(`/applications(appId='${APP_ID}')`, {
        status: 404,
        body: { error: { code: 'Request_ResourceNotFound' } },
      }),
    ]);
    assert.deepEqual(await client.deleteAppRegistration({ appId: APP_ID }), {
      outcome: 'already-deleted',
    });
  });
});

describe('AppRegistrationClient.purgeDeletedAppRegistration', () => {
  it('purges via DELETE /directory/deletedItems/{objectId}', async () => {
    const { client, calls } = appHarness([
      route(`${DELETED_ITEMS}/${OBJECT_ID}`, { status: 204 }),
    ]);

    assert.deepEqual(
      await client.purgeDeletedAppRegistration({ objectId: OBJECT_ID }),
      { outcome: 'purged' },
    );

    const call = calls[0];
    assert.ok(call);
    // THE OBJECT ID, in the path. Not `(appId='…')`, not a filter — the
    // recycle bin has no alternate-key form, so this is the only address that
    // works and the assertion pins it.
    assert.ok(call.url.endsWith(`${DELETED_ITEMS}/${OBJECT_ID}`));
    assert.ok(!call.url.includes(APP_ID), 'the appId must not appear in the URL');
    assert.equal(call.init?.method, 'DELETE');
  });

  it('answers already-absent when the bin entry is genuinely gone', async () => {
    const { client } = appHarness([
      route(`${DELETED_ITEMS}/${OBJECT_ID}`, {
        status: 404,
        body: { error: { code: 'Request_ResourceNotFound' } },
      }),
      // The 404 triggers a confirming scan; an empty bin is the proof that
      // 'already-absent' is honest here.
      route(RECYCLE_BIN_LIST, { status: 200, body: { value: [] } }),
    ]);
    assert.deepEqual(
      await client.purgeDeletedAppRegistration({ objectId: OBJECT_ID }),
      { outcome: 'already-absent' },
    );
  });

  it('is idempotent across two passes of the same reset', async () => {
    const { client } = appHarness([
      route(`${DELETED_ITEMS}/${OBJECT_ID}`, { status: 204 }, { status: 404 }),
      route(RECYCLE_BIN_LIST, { status: 200, body: { value: [] } }),
    ]);
    const first = await client.purgeDeletedAppRegistration({ objectId: OBJECT_ID });
    const second = await client.purgeDeletedAppRegistration({ objectId: OBJECT_ID });
    assert.equal(first.outcome, 'purged');
    // A reset that fails on its second run because its first run worked is
    // useless — this is the property that makes it re-runnable.
    assert.equal(second.outcome, 'already-absent');
  });

  it('refuses to call an appId "already absent" and names the object id instead', async () => {
    // THE TRAP. Passing the appId 404s exactly like a purged entry. Reporting
    // 'already-absent' would claim the uniqueName is free while its 30-day
    // reservation still stands, and the next provisioning run would collide
    // with an object nobody can see — byte5ai/omadia#916, again.
    const { client } = appHarness([
      route(`${DELETED_ITEMS}/${APP_ID}`, {
        status: 404,
        body: { error: { code: 'Request_ResourceNotFound' } },
      }),
      route(RECYCLE_BIN_LIST, {
        status: 200,
        body: {
          value: [
            {
              id: OBJECT_ID,
              appId: APP_ID,
              uniqueName: 'omadia-agent-falcon',
              displayName: 'falcon',
            },
          ],
        },
      }),
    ]);

    await assert.rejects(
      () => client.purgeDeletedAppRegistration({ objectId: APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof DeletedObjectIdMismatchError);
        assert.equal(err.passedId, APP_ID);
        // The id that WOULD have worked, so the fix is a retry with a value
        // the caller already holds.
        assert.equal(err.objectId, OBJECT_ID);
        return true;
      },
    );
  });

  it('still answers already-absent when the recycle-bin probe is not permitted', async () => {
    // The bin listing is not covered by Application.ReadWrite.OwnedBy, so a
    // tenant that granted only the provisioner's scope gets 403 there. A
    // diagnostic that cannot run must not become a failure of its own: with no
    // evidence AGAINST 'already-absent', that is the honest answer.
    const { client } = appHarness([
      route(`${DELETED_ITEMS}/${OBJECT_ID}`, { status: 404 }),
      route(RECYCLE_BIN_LIST, {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);
    assert.deepEqual(
      await client.purgeDeletedAppRegistration({ objectId: OBJECT_ID }),
      { outcome: 'already-absent' },
    );
  });

  it('maps 403 on the purge itself to ConsentMissingError', async () => {
    const { client } = appHarness([
      route(`${DELETED_ITEMS}/${OBJECT_ID}`, {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);
    await assert.rejects(
      () => client.purgeDeletedAppRegistration({ objectId: OBJECT_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.deepEqual(err.missingScopes, [APP_REGISTRATION_SCOPE]);
        return true;
      },
    );
  });

  it('rejects an empty object id before touching Graph', async () => {
    const { client, calls } = appHarness([]);
    await assert.rejects(
      () => client.purgeDeletedAppRegistration({ objectId: '  ' }),
      /invalid_argument/,
    );
    assert.equal(calls.length, 0);
  });
});

describe('CatalogUploadClient.removeFromCatalog', () => {
  const TEAMS_APP_ID = 'catalog-app-0001';

  it('removes the published app with the delegated token', async () => {
    const { client, calls } = catalogHarness([
      route(`/appCatalogs/teamsApps/${TEAMS_APP_ID}`, { status: 204 }),
    ]);

    assert.deepEqual(
      await client.removeFromCatalog({
        teamsAppId: TEAMS_APP_ID,
        tokens: tokens(),
      }),
      { outcome: 'removed' },
    );

    const call = calls[0];
    assert.ok(call);
    assert.equal(call.init?.method, 'DELETE');
    const headers = call.init?.headers as Record<string, string> | undefined;
    // DELEGATED, like the upload: Graph documents application permissions for
    // this verb as "Not supported." in both privilege columns.
    assert.equal(headers?.['Authorization'], 'Bearer delegated-access-token');
  });

  it('answers already-absent on 404 so a reset can be re-run', async () => {
    const { client } = catalogHarness([
      route(`/appCatalogs/teamsApps/${TEAMS_APP_ID}`, {
        status: 404,
        body: { error: { code: 'NotFound' } },
      }),
    ]);
    // The realistic reset case: a run that failed before it ever published.
    assert.deepEqual(
      await client.removeFromCatalog({
        teamsAppId: TEAMS_APP_ID,
        tokens: tokens(),
      }),
      { outcome: 'already-absent' },
    );
  });

  it('needs no NEW scope — the publish credential already covers it', async () => {
    const { client } = catalogHarness([
      route(`/appCatalogs/teamsApps/${TEAMS_APP_ID}`, { status: 204 }),
    ]);
    // A credential minted for publishing (pre-0.8.0, no chat scope) removes
    // just fine. Unlike listChats, this costs the operator no second sign-in.
    const publishOnly = tokens({ scopes: ['AppCatalog.ReadWrite.All'] });
    assert.deepEqual(
      await client.removeFromCatalog({
        teamsAppId: TEAMS_APP_ID,
        tokens: publishOnly,
      }),
      { outcome: 'removed' },
    );
  });

  it('reports a credential without the catalog scope as scope-missing', async () => {
    const { client, calls } = catalogHarness([]);
    await assert.rejects(
      () =>
        client.removeFromCatalog({
          teamsAppId: TEAMS_APP_ID,
          tokens: tokens({ scopes: ['openid'] }),
        }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedScopeRequiredError);
        assert.equal(err.reason, 'scope-missing');
        assert.deepEqual(err.requiredScopes, [APP_CATALOG_DELEGATED_SCOPE]);
        return true;
      },
    );
    assert.equal(calls.length, 0);
  });

  it('reports a stale token rather than rotating one it cannot hand back', async () => {
    const { client, calls } = catalogHarness([]);
    await assert.rejects(
      () =>
        client.removeFromCatalog({
          teamsAppId: TEAMS_APP_ID,
          tokens: tokens({ expiresAt: new Date(Date.now() - 1000).toISOString() }),
        }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedTokenExpiredError);
        assert.equal(err.recoverableByRefresh, true);
        return true;
      },
    );
    assert.equal(calls.length, 0);
  });

  it('rejects an empty teamsAppId before touching Graph', async () => {
    const { client, calls } = catalogHarness([]);
    await assert.rejects(
      () => client.removeFromCatalog({ teamsAppId: '', tokens: tokens() }),
      /invalid_argument/,
    );
    assert.equal(calls.length, 0);
  });
});

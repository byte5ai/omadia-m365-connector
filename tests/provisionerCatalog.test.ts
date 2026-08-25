import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  APP_CATALOG_SCOPE,
  CatalogUploadClient,
} from '../src/teamsProvisioner/catalog.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  ConsentMissingError,
  ProvisioningThrottledError,
  TeamsProvisionerError,
} from '../src/teamsProvisioner/errors.js';

// The catalog-upload step: POST /appCatalogs/teamsApps behind the shared
// ProvisioningHttp choke point — 409 re-resolves the existing entry by
// externalId (a true idempotent no-op, never an exception, never a bare
// swallow), 403 maps to the typed ConsentMissingError the agent factory
// branches on, 429 rides the shared Retry-After backoff. Mocked-fetch style
// follows tests/provisionerInstall.test.ts / tests/teamsProvisionerHttp.test.ts.

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

const EXTERNAL_ID = 'b5561ec9-8cab-4aa3-8aa2-d8d7172e4311';
const CATALOG_ID = 'e3e29acb-8c79-412b-b746-e6c39ff4cd22';
const DISPLAY_NAME = 'Agent Bot';
const VERSION = '1.4.2';
// Any buffer works — deliberately NOT a real zip (no buildAppPackage dependency).
const PACKAGE_ZIP = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x2a]);

const UPLOAD_URL = 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps';
// The lookup is the same collection URL plus the $filter — match it FIRST
// (routes are checked in order) so the plain upload match doesn't shadow it.
const LOOKUP_URL_MATCH = '/appCatalogs/teamsApps?$filter=externalId';
const UPLOAD_URL_MATCH = '/appCatalogs/teamsApps';

/** Full teamsApp shape as the fresh-upload result should carry it. */
const EXPECTED_APP = {
  teamsAppId: CATALOG_ID,
  externalId: EXTERNAL_ID,
  displayName: DISPLAY_NAME,
  version: VERSION,
};

const LOOKUP_HIT: ResponseSpec = {
  status: 200,
  body: {
    value: [
      {
        id: CATALOG_ID,
        externalId: EXTERNAL_ID,
        displayName: DISPLAY_NAME,
        distributionMethod: 'organization',
        appDefinitions: [{ version: VERSION }],
      },
    ],
  },
};

function harness(routes: Route[]): {
  client: CatalogUploadClient;
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
  const client = new CatalogUploadClient({ http, log: () => {} });
  return { client, calls, sleeps };
}

/** The recorded catalog calls (token calls filtered out). */
function catalogCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter((c) => c.url.includes(UPLOAD_URL_MATCH));
}

function uploadCalls(calls: FetchCall[]): FetchCall[] {
  return catalogCalls(calls).filter((c) => c.init?.method === 'POST');
}

function lookupCalls(calls: FetchCall[]): FetchCall[] {
  return catalogCalls(calls).filter((c) => c.init?.method === 'GET');
}

describe('CatalogUploadClient.uploadToCatalog', () => {
  it('POSTs the raw zip and returns created from a full 201 body (no lookup)', async () => {
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, { status: 500, body: { error: 'must not be called' } }),
      route(UPLOAD_URL_MATCH, {
        status: 201,
        body: {
          id: CATALOG_ID,
          externalId: EXTERNAL_ID,
          displayName: DISPLAY_NAME,
          distributionMethod: 'organization',
          appDefinitions: [{ version: VERSION }],
        },
      }),
    ]);

    const result = await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, EXPECTED_APP);

    const [call] = uploadCalls(calls);
    assert.ok(call, 'expected one upload POST');
    assert.equal(call.url, UPLOAD_URL);
    const headers = call.init?.headers as Record<string, string>;
    assert.equal(headers['Authorization'], 'Bearer tok');
    assert.equal(headers['content-type'], 'application/zip');
    assert.deepEqual(new Uint8Array(call.init?.body as Uint8Array), PACKAGE_ZIP);
    assert.equal(lookupCalls(calls).length, 0);
  });

  it('re-resolves by externalId when the 2xx body lacks the manifest version', async () => {
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, LOOKUP_HIT),
      route(UPLOAD_URL_MATCH, {
        status: 201,
        // Graph's real POST response: no appDefinitions, no version.
        body: {
          id: CATALOG_ID,
          externalId: EXTERNAL_ID,
          displayName: DISPLAY_NAME,
          distributionMethod: 'organization',
        },
      }),
    ]);

    const result = await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, EXPECTED_APP);
    assert.equal(uploadCalls(calls).length, 1);
    assert.equal(lookupCalls(calls).length, 1);
  });

  it('resolves 409 to already-existed via the externalId lookup — same shape as a fresh upload', async () => {
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, LOOKUP_HIT),
      route(UPLOAD_URL_MATCH, {
        status: 409,
        body: { error: { code: 'Conflict', message: 'App already exists' } },
      }),
    ]);

    const result = await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
    });

    assert.equal(result.outcome, 'already-existed');
    // The load-bearing bit: the 409 path returns the SAME value shape a
    // fresh upload returns — a second provisioning run is a true no-op.
    assert.deepEqual(result.value, EXPECTED_APP);

    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected the 409 path to re-resolve by externalId');
    assert.ok(lookup.url.includes(`$filter=externalId%20eq%20'${EXTERNAL_ID}'`));
    assert.ok(lookup.url.includes('$expand=appDefinitions'));
    assert.equal(uploadCalls(calls).length, 1);
  });

  it('fails loudly when the 409 lookup finds no app under the externalId', async () => {
    const { client } = harness([
      route(LOOKUP_URL_MATCH, { status: 200, body: { value: [] } }),
      route(UPLOAD_URL_MATCH, { status: 409 }),
    ]);

    await assert.rejects(
      client.uploadToCatalog({ packageZip: PACKAGE_ZIP, externalId: EXTERNAL_ID }),
      new RegExp(`found no catalog app with externalId=${EXTERNAL_ID}`),
    );
  });

  it('maps 403 to ConsentMissingError carrying the catalog scope (graph)', async () => {
    const { client } = harness([
      route(UPLOAD_URL_MATCH, {
        status: 403,
        body: { error: { code: 'Forbidden' } },
      }),
    ]);

    await assert.rejects(
      client.uploadToCatalog({ packageZip: PACKAGE_ZIP, externalId: EXTERNAL_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        // Same typed family as installToTeam → the agent factory gets ONE
        // fallback branch (byte5ai/omadia#863-865).
        assert.ok(err instanceof TeamsProvisionerError);
        assert.deepEqual(err.missingScopes, [APP_CATALOG_SCOPE]);
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });

  it('retries 429 honouring Retry-After, then succeeds', async () => {
    const { client, calls, sleeps } = harness([
      route(LOOKUP_URL_MATCH, LOOKUP_HIT),
      route(
        UPLOAD_URL_MATCH,
        { status: 429, headers: { 'Retry-After': '7' } },
        {
          status: 201,
          body: {
            id: CATALOG_ID,
            externalId: EXTERNAL_ID,
            displayName: DISPLAY_NAME,
            appDefinitions: [{ version: VERSION }],
          },
        },
      ),
    ]);

    const result = await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, EXPECTED_APP);
    assert.equal(uploadCalls(calls).length, 2);
    assert.deepEqual(sleeps, [7000]);
  });

  it('throws ProvisioningThrottledError when the 429 budget is exhausted', async () => {
    const { client, calls } = harness([
      route(UPLOAD_URL_MATCH, {
        status: 429,
        headers: { 'Retry-After': '3' },
      }),
    ]);

    await assert.rejects(
      client.uploadToCatalog({ packageZip: PACKAGE_ZIP, externalId: EXTERNAL_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ProvisioningThrottledError);
        assert.equal(err.resource, 'graph');
        assert.equal(err.retryAfterSeconds, 3);
        return true;
      },
    );
    // max429Retries: 2 → initial attempt + 2 retries.
    assert.equal(uploadCalls(calls).length, 3);
  });

  it('rejects empty externalId / empty packageZip before any fetch', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.uploadToCatalog({ packageZip: PACKAGE_ZIP, externalId: '  ' }),
      /invalid_argument: 'externalId'/,
    );
    await assert.rejects(
      client.uploadToCatalog({ packageZip: new Uint8Array(0), externalId: EXTERNAL_ID }),
      /invalid_argument: 'packageZip'/,
    );
    assert.equal(calls.length, 0);
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  APP_CATALOG_SCOPE,
  CatalogUploadClient,
} from '../src/teamsProvisioner/catalog.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  ConsentMissingError,
  DelegatedConsentRequiredError,
  DelegatedSignInRequiredError,
  DelegatedTokenExpiredError,
  ProvisioningThrottledError,
  TeamsProvisionerError,
  isTransientProvisioningFailure,
} from '../src/teamsProvisioner/errors.js';
import { APP_CATALOG_DELEGATED_SCOPE } from '../src/teamsProvisioner/delegatedAuth.js';

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

const EXTERNAL_ID = 'b5561ec9-8cab-4aa3-8aa2-d8d7172e4311';
const CATALOG_ID = 'e3e29acb-8c79-412b-b746-e6c39ff4cd22';
const DISPLAY_NAME = 'Agent Bot';
const VERSION = '1.4.2';
// Any buffer works — deliberately NOT a real zip (no buildAppPackage dependency).
const PACKAGE_ZIP = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x2a]);
/**
 * Since 0.6.0 the upload is DELEGATED-ONLY (byte5ai/omadia#924): Graph does not
 * support application permissions for POST /appCatalogs/teamsApps, so every
 * upload here carries a user token. The catalog LOOKUP stays app-only, which is
 * why the 409 re-resolution below still works without one.
 */
const DELEGATED_TOKEN = 'delegated-user-access-token';

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
  it('picks the PUBLISHED appDefinition when several versions exist (409 path)', async () => {
    // Graph documents appDefinitions as one entry PER VERSION, in no
    // guaranteed order. Taking the first entry would silently report a stale
    // version on exactly the idempotent path this step exists for (major
    // review finding). The published definition must win; the $expand also
    // requests publishingState so the choice is possible.
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, {
        status: 200,
        body: {
          value: [
            {
              id: CATALOG_ID,
              externalId: EXTERNAL_ID,
              displayName: DISPLAY_NAME,
              appDefinitions: [
                { version: '1.0.0', publishingState: 'rejected' },
                { version: VERSION, publishingState: 'published' },
                { version: '1.0.1', publishingState: 'submitted' },
              ],
            },
          ],
        },
      }),
      route(UPLOAD_URL_MATCH, { status: 409, body: { error: { code: 'Conflict' } } }),
    ]);

    const result = await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
      delegatedAccessToken: DELEGATED_TOKEN,
    });

    assert.equal(result.outcome, 'already-existed');
    assert.equal(result.value.version, VERSION, 'published definition wins');
    const [lookup] = lookupCalls(calls);
    assert.ok(lookup?.url.includes('publishingState'), '$expand selects publishingState');
  });

  it('falls back to the highest version when no definition is marked published', async () => {
    const { client } = harness([
      route(LOOKUP_URL_MATCH, {
        status: 200,
        body: {
          value: [
            {
              id: CATALOG_ID,
              externalId: EXTERNAL_ID,
              displayName: DISPLAY_NAME,
              appDefinitions: [
                { version: '1.9.9' },
                { version: '1.10.0' },
                { version: '1.2.3' },
              ],
            },
          ],
        },
      }),
      route(UPLOAD_URL_MATCH, { status: 409, body: {} }),
    ]);

    const result = await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
      delegatedAccessToken: DELEGATED_TOKEN,
    });
    assert.equal(
      result.value.version,
      '1.10.0',
      'numeric-aware compare, not first-entry or lexicographic',
    );
  });

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
      delegatedAccessToken: DELEGATED_TOKEN,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, EXPECTED_APP);

    const [call] = uploadCalls(calls);
    assert.ok(call, 'expected one upload POST');
    assert.equal(call.url, UPLOAD_URL);
    const headers = call.init?.headers as Record<string, string>;
    // The DELEGATED token, not the app-only one from the token cache: Graph
    // refuses application permissions for this verb (byte5ai/omadia#924).
    assert.equal(headers['Authorization'], `Bearer ${DELEGATED_TOKEN}`);
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
      delegatedAccessToken: DELEGATED_TOKEN,
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
      delegatedAccessToken: DELEGATED_TOKEN,
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
      client.uploadToCatalog({
        packageZip: PACKAGE_ZIP,
        externalId: EXTERNAL_ID,
        delegatedAccessToken: DELEGATED_TOKEN,
      }),
      new RegExp(`found no catalog app with externalId=${EXTERNAL_ID}`),
    );
  });

  it('maps 403 to DelegatedConsentRequiredError, not the app-only ConsentMissingError', async () => {
    // Since 0.6.0 this call carries a USER token, so a 403 can no longer mean
    // "the connector app is missing an app role" — it means the tenant never
    // consented to the delegated scope on the PUBLISHER app. Reporting the
    // app-only error here would send an operator to fix the wrong registration
    // (byte5ai/omadia#924).
    const { client } = harness([
      route(UPLOAD_URL_MATCH, {
        status: 403,
        body: { error: { code: 'Forbidden' } },
      }),
    ]);

    await assert.rejects(
      client.uploadToCatalog({
        packageZip: PACKAGE_ZIP,
        externalId: EXTERNAL_ID,
        delegatedAccessToken: DELEGATED_TOKEN,
      }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedConsentRequiredError);
        assert.ok(!(err instanceof ConsentMissingError));
        // Still one catchable family for the agent factory.
        assert.ok(err instanceof TeamsProvisionerError);
        assert.deepEqual(err.requiredScopes, [APP_CATALOG_DELEGATED_SCOPE]);
        assert.equal(err.step, 'appCatalogs.teamsApps.publish');
        assert.ok(err.adminConsentUrl.length > 0);
        // The app-only 403 the http layer raised is preserved as the cause,
        // so nothing diagnostic is lost in the translation.
        assert.ok((err as { cause?: unknown }).cause instanceof ConsentMissingError);
        return true;
      },
    );
  });

  it('maps 401 to DelegatedTokenExpiredError, flagged as refresh-recoverable', async () => {
    // A user token that aged out mid-flight is fixed by a refresh, with no
    // human involved — a different remedy from every other failure here, so it
    // must be a different error.
    const { client } = harness([
      route(UPLOAD_URL_MATCH, {
        status: 401,
        body: { error: { code: 'InvalidAuthenticationToken' } },
      }),
    ]);

    await assert.rejects(
      client.uploadToCatalog({
        packageZip: PACKAGE_ZIP,
        externalId: EXTERNAL_ID,
        delegatedAccessToken: DELEGATED_TOKEN,
      }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedTokenExpiredError);
        assert.equal(err.reason, 'access-token-expired');
        assert.equal(err.recoverableByRefresh, true);
        return true;
      },
    );
  });

  it('refuses an upload with NO delegated token before it reaches Graph', async () => {
    // The app-only upload is known-unsupported, so spending a Graph call to
    // learn that again would only produce a misleading 403. Refuse locally, and
    // do so with the one error whose remedy is "start the device-code sign-in".
    const { client, calls } = harness([
      route(UPLOAD_URL_MATCH, { status: 201, body: {} }),
    ]);

    await assert.rejects(
      client.uploadToCatalog({ packageZip: PACKAGE_ZIP, externalId: EXTERNAL_ID }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedSignInRequiredError);
        assert.ok(err instanceof TeamsProvisionerError);
        assert.equal(err.step, 'appCatalogs.teamsApps.publish');
        assert.deepEqual(err.requiredScopes, [APP_CATALOG_DELEGATED_SCOPE]);
        assert.ok(!isTransientProvisioningFailure(err), 'never retryable');
        return true;
      },
    );
    assert.equal(catalogCalls(calls).length, 0, 'no Graph call may be made');
  });

  it('sends the DELEGATED token on the upload and the APP-ONLY token on the lookup', async () => {
    // The asymmetry is the design: only the upload is delegated-only, so a
    // missing/stale user token degrades exactly one operation instead of
    // blinding the provisioner to what is already published.
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, LOOKUP_HIT),
      route(UPLOAD_URL_MATCH, { status: 409 }),
    ]);

    await client.uploadToCatalog({
      packageZip: PACKAGE_ZIP,
      externalId: EXTERNAL_ID,
      delegatedAccessToken: DELEGATED_TOKEN,
    });

    const [upload] = uploadCalls(calls);
    const [lookup] = lookupCalls(calls);
    assert.ok(upload && lookup);
    assert.equal(
      (upload.init?.headers as Record<string, string>)['Authorization'],
      `Bearer ${DELEGATED_TOKEN}`,
    );
    assert.equal(
      (lookup.init?.headers as Record<string, string>)['Authorization'],
      'Bearer tok',
      'the catalog lookup must keep using the app-only token',
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
      delegatedAccessToken: DELEGATED_TOKEN,
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
      client.uploadToCatalog({
        packageZip: PACKAGE_ZIP,
        externalId: EXTERNAL_ID,
        delegatedAccessToken: DELEGATED_TOKEN,
      }),
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
      client.uploadToCatalog({
        packageZip: PACKAGE_ZIP,
        externalId: '  ',
        delegatedAccessToken: DELEGATED_TOKEN,
      }),
      /invalid_argument: 'externalId'/,
    );
    await assert.rejects(
      client.uploadToCatalog({
        packageZip: new Uint8Array(0),
        externalId: EXTERNAL_ID,
        delegatedAccessToken: DELEGATED_TOKEN,
      }),
      /invalid_argument: 'packageZip'/,
    );
    assert.equal(calls.length, 0);
  });
});

// getCatalogApp: the upload-free lookup probe — same Graph query (and same
// published-wins version selection) as the 409 idempotent path, but a miss is
// a plain { found: false } outcome, never an exception, and no POST happens.
describe('CatalogUploadClient.getCatalogApp', () => {
  it('resolves an existing app — published appDefinition wins over other versions', async () => {
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, {
        status: 200,
        body: {
          value: [
            {
              id: CATALOG_ID,
              externalId: EXTERNAL_ID,
              displayName: DISPLAY_NAME,
              appDefinitions: [
                { version: '1.0.0', publishingState: 'rejected' },
                { version: VERSION, publishingState: 'published' },
                { version: '9.9.9', publishingState: 'submitted' },
              ],
            },
          ],
        },
      }),
    ]);

    const result = await client.getCatalogApp({
      teamsAppExternalId: EXTERNAL_ID,
    });

    assert.deepEqual(result, {
      found: true,
      teamsAppId: CATALOG_ID,
      displayName: DISPLAY_NAME,
      publishedVersion: VERSION,
    });

    // Lookup only — this method must never upload anything.
    assert.equal(uploadCalls(calls).length, 0);
    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected exactly one GET lookup');
    assert.ok(lookup.url.includes(`$filter=externalId%20eq%20'${EXTERNAL_ID}'`));
    assert.ok(lookup.url.includes('$expand=appDefinitions'));
    assert.ok(lookup.url.includes('publishingState'), '$expand selects publishingState');
  });

  it('stays a hit when Graph omits displayName and version (lenient parse)', async () => {
    // The strict CatalogTeamsApp parse would turn this thin-but-real entry
    // into a miss — the lookup only requires id + matching externalId.
    const { client } = harness([
      route(LOOKUP_URL_MATCH, {
        status: 200,
        body: { value: [{ id: CATALOG_ID, externalId: EXTERNAL_ID }] },
      }),
    ]);

    const result = await client.getCatalogApp({
      teamsAppExternalId: EXTERNAL_ID,
    });

    assert.deepEqual(result, { found: true, teamsAppId: CATALOG_ID });
  });

  it('answers { found: false } on an empty result — an outcome, not an exception', async () => {
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, { status: 200, body: { value: [] } }),
    ]);

    const result = await client.getCatalogApp({
      teamsAppExternalId: EXTERNAL_ID,
    });

    assert.deepEqual(result, { found: false });
    assert.equal(lookupCalls(calls).length, 1);
    assert.equal(uploadCalls(calls).length, 0);
  });

  it('ignores entries whose externalId differs (server-side filter not trusted)', async () => {
    const { client } = harness([
      route(LOOKUP_URL_MATCH, {
        status: 200,
        body: {
          value: [
            { id: 'other-id', externalId: 'someone-else', displayName: 'Other' },
          ],
        },
      }),
    ]);

    const result = await client.getCatalogApp({
      teamsAppExternalId: EXTERNAL_ID,
    });
    assert.deepEqual(result, { found: false });
  });

  it("escapes quotes in the externalId (quote doubling + encodeURIComponent) — filter stays injection-safe", async () => {
    const trickyId = "agent's-app";
    const { client, calls } = harness([
      route(LOOKUP_URL_MATCH, { status: 200, body: { value: [] } }),
    ]);

    const result = await client.getCatalogApp({ teamsAppExternalId: trickyId });
    assert.deepEqual(result, { found: false });

    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected one GET lookup');
    // encodeURIComponent leaves single quotes as-is, so the doubled quote is
    // directly visible in the URL: externalId eq 'agent''s-app'.
    assert.ok(
      lookup.url.includes("$filter=externalId%20eq%20'agent''s-app'"),
      `filter must carry the doubled quote, got: ${lookup.url}`,
    );
  });

  it('maps 403 to ConsentMissingError carrying the catalog scope (graph)', async () => {
    const { client } = harness([
      route(LOOKUP_URL_MATCH, { status: 403, body: { error: { code: 'Forbidden' } } }),
    ]);

    await assert.rejects(
      client.getCatalogApp({ teamsAppExternalId: EXTERNAL_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.ok(err instanceof TeamsProvisionerError);
        assert.deepEqual(err.missingScopes, [APP_CATALOG_SCOPE]);
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });

  it('rides the shared 429 backoff and throws ProvisioningThrottledError when exhausted', async () => {
    const { client, calls, sleeps } = harness([
      route(LOOKUP_URL_MATCH, {
        status: 429,
        headers: { 'Retry-After': '5' },
      }),
    ]);

    await assert.rejects(
      client.getCatalogApp({ teamsAppExternalId: EXTERNAL_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ProvisioningThrottledError);
        assert.equal(err.resource, 'graph');
        assert.equal(err.retryAfterSeconds, 5);
        return true;
      },
    );
    // max429Retries: 2 → initial attempt + 2 retries, honouring Retry-After.
    assert.equal(lookupCalls(calls).length, 3);
    assert.deepEqual(sleeps, [5000, 5000]);
  });

  it('rejects an empty teamsAppExternalId before any fetch', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.getCatalogApp({ teamsAppExternalId: '  ' }),
      /invalid_argument: 'teamsAppExternalId'/,
    );
    assert.equal(calls.length, 0);
  });
});

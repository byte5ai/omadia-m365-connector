import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  APP_CATALOG_DELEGATED_SCOPE,
  DELEGATED_PUBLISH_SCOPES,
  DelegatedAuthClient,
  adminConsentUrl,
  coversCatalogPublish,
  describeSignInStatus,
  isAccessTokenStale,
  revokeInstructions,
  type DelegatedTokenSet,
} from '../src/teamsProvisioner/delegatedAuth.js';
import {
  DelegatedConsentRequiredError,
  DelegatedTokenExpiredError,
  DeviceCodeFlowError,
  TeamsProvisionerError,
  isTransientProvisioningFailure,
} from '../src/teamsProvisioner/errors.js';

// The device authorization grant (RFC 8628) against a mocked identity platform.
//
// The flow is spoken directly rather than through MSAL because it has to be
// POLLABLE FROM ANY PROCESS — `acquireTokenByDeviceCode` blocks until the user
// finishes, which would force flow state into a process-local map that dies on
// the next deploy. These tests pin that property: every poll below runs on a
// client instance that never saw the start call.
//
// Mocked-fetch style follows tests/provisionerCatalog.test.ts.

const TENANT_ID = '99999999-8888-7777-6666-555555555555';
const CLIENT_ID = '11111111-2222-3333-4444-555555555555';
const DEVICE_CODE = 'device-code-value';
const USER_CODE = 'BXTM4NCD';

interface FetchCall {
  url: string;
  body: URLSearchParams;
}

interface ResponseSpec {
  status: number;
  body?: unknown;
}

function makeResponse(spec: ResponseSpec): Response {
  const text = spec.body === undefined ? '' : JSON.stringify(spec.body);
  return {
    ok: spec.status >= 200 && spec.status < 300,
    status: spec.status,
    headers: { get: () => null },
    json: async () => JSON.parse(text),
    text: async () => text,
  } as unknown as Response;
}

/** Queue of responses (last entry repeats), plus a record of what was sent. */
function mockLogin(
  responses: ResponseSpec[],
  calls: FetchCall[],
): typeof fetch {
  let next = 0;
  return (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    calls.push({
      url: String(input),
      body: new URLSearchParams(String(init?.body ?? '')),
    });
    const spec = responses[Math.min(next, responses.length - 1)];
    next += 1;
    return makeResponse(spec ?? { status: 500 });
  }) as typeof fetch;
}

const START_OK: ResponseSpec = {
  status: 200,
  body: {
    device_code: DEVICE_CODE,
    user_code: USER_CODE,
    verification_uri: 'https://microsoft.com/devicelogin',
    expires_in: 900,
    interval: 5,
    message: `Open https://microsoft.com/devicelogin and enter ${USER_CODE}`,
  },
};

/** A minimal, decodable id_token payload — signature intentionally junk. */
function idToken(claims: Record<string, unknown>): string {
  const payload = Buffer.from(JSON.stringify(claims), 'utf8').toString('base64url');
  return `eyJhbGciOiJSUzI1NiJ9.${payload}.signature`;
}

function client(
  responses: ResponseSpec[],
  calls: FetchCall[] = [],
  sleeps: number[] = [],
): DelegatedAuthClient {
  return new DelegatedAuthClient({
    tenantId: TENANT_ID,
    clientId: CLIENT_ID,
    fetchImpl: mockLogin(responses, calls),
    sleep: async (ms) => {
      sleeps.push(ms);
    },
    log: () => {},
  });
}

describe('DelegatedAuthClient.startDeviceCode', () => {
  it('asks for the catalog scope plus offline_access and returns display material', async () => {
    const calls: FetchCall[] = [];
    const start = await client([START_OK], calls).startDeviceCode();

    const [call] = calls;
    assert.ok(call);
    assert.ok(call.url.endsWith(`/${TENANT_ID}/oauth2/v2.0/devicecode`));
    assert.equal(call.body.get('client_id'), CLIENT_ID);
    const scopes = (call.body.get('scope') ?? '').split(' ');
    assert.ok(scopes.includes(APP_CATALOG_DELEGATED_SCOPE));
    // Without offline_access there is no refresh token, and the whole
    // "one admin sign-in per tenant" promise collapses into an hourly one.
    assert.ok(scopes.includes('offline_access'));

    assert.equal(start.userCode, USER_CODE);
    assert.equal(start.verificationUri, 'https://microsoft.com/devicelogin');
    assert.equal(start.intervalSeconds, 5);
    assert.deepEqual(start.scopes, DELEGATED_PUBLISH_SCOPES);
    assert.equal(start.adminConsentUrl, adminConsentUrl(TENANT_ID, CLIENT_ID));
    assert.ok(Date.parse(start.expiresAt) > Date.now());
  });

  it('polls through the login stack replication window after a fresh app create', async () => {
    // Entra replicates the directory and the sign-in stack separately: Graph
    // can already serve a just-created application while login.microsoftonline
    // still answers `unauthorized_client`. Failing here would make the very
    // first sign-in after install fail for no visible reason
    // (byte5ai/omadia#916, one layer further out).
    const sleeps: number[] = [];
    const start = await client(
      [
        { status: 400, body: { error: 'unauthorized_client' } },
        { status: 400, body: { error: 'unauthorized_client' } },
        START_OK,
      ],
      [],
      sleeps,
    ).startDeviceCode();

    assert.equal(start.userCode, USER_CODE);
    assert.deepEqual(sleeps, [2000, 4000], 'exponential, capped backoff');
  });

  it('gives up with an actionable error when the client stays unusable', async () => {
    await assert.rejects(
      () =>
        client([{ status: 400, body: { error: 'unauthorized_client' } }]).startDeviceCode(),
      (err: unknown) => {
        assert.ok(err instanceof DeviceCodeFlowError);
        assert.equal(err.oauthError, 'unauthorized_client');
        // The message must name BOTH plausible causes — at the protocol level
        // "not replicated yet" and "not a public client" are indistinguishable.
        assert.match(err.message, /public client/i);
        assert.ok(!isTransientProvisioningFailure(err));
        return true;
      },
    );
  });

  it('rejects a non-replication failure immediately, without burning the budget', async () => {
    const calls: FetchCall[] = [];
    await assert.rejects(
      () =>
        client(
          [{ status: 400, body: { error: 'invalid_scope', error_description: 'bad scope' } }],
          calls,
        ).startDeviceCode(),
      (err: unknown) => {
        assert.ok(err instanceof DeviceCodeFlowError);
        assert.equal(err.oauthError, 'invalid_scope');
        return true;
      },
    );
    assert.equal(calls.length, 1, 'a deterministic rejection must not be retried');
  });
});

describe('DelegatedAuthClient.pollDeviceCode', () => {
  /** Start a flow on one client, then poll it on a DIFFERENT one. */
  async function handle(): Promise<string> {
    const start = await client([START_OK]).startDeviceCode();
    return start.flowHandle;
  }

  it('reports authorization_pending as a status, never as an exception', async () => {
    const flowHandle = await handle();
    const calls: FetchCall[] = [];
    const result = await client(
      [{ status: 400, body: { error: 'authorization_pending' } }],
      calls,
    ).pollDeviceCode({ flowHandle });

    assert.equal(result.status, 'pending');
    assert.ok(result.status === 'pending' && result.retryAfterSeconds === 5);

    const [call] = calls;
    assert.ok(call);
    assert.equal(
      call.body.get('grant_type'),
      'urn:ietf:params:oauth:grant-type:device_code',
    );
    assert.equal(call.body.get('device_code'), DEVICE_CODE);
    assert.equal(call.body.get('client_id'), CLIENT_ID);
  });

  it('widens the interval on slow_down instead of retrying at the same rate', async () => {
    // RFC 8628 §3.5: slow_down is a permanent instruction, not a one-off. The
    // caller polls at whatever interval we hand back, so widening it here is
    // the only thing that actually slows the loop.
    const flowHandle = await handle();
    const result = await client([
      { status: 400, body: { error: 'slow_down' } },
    ]).pollDeviceCode({ flowHandle });

    assert.equal(result.status, 'pending');
    assert.ok(result.status === 'pending' && result.retryAfterSeconds > 5);
  });

  it('reports expired_token as a terminal status', async () => {
    const flowHandle = await handle();
    const result = await client([
      {
        status: 400,
        body: { error: 'expired_token', error_description: 'AADSTS70019: code expired' },
      },
    ]).pollDeviceCode({ flowHandle });

    assert.equal(result.status, 'expired');
    assert.ok(result.status === 'expired' && result.reason?.includes('AADSTS70019'));
  });

  it('reports authorization_declined and access_denied as declined, keeping the AADSTS code', async () => {
    // access_denied is also where a Conditional Access "block device code flow"
    // policy lands — Microsoft recommends that policy, so it is a realistic
    // production answer. The AADSTS code in `reason` is the ONLY thing that
    // tells it apart from an admin who simply clicked cancel.
    for (const error of ['authorization_declined', 'access_denied']) {
      const flowHandle = await handle();
      const result = await client([
        {
          status: 400,
          body: { error, error_description: 'AADSTS50199: blocked by policy' },
        },
      ]).pollDeviceCode({ flowHandle });

      assert.equal(result.status, 'declined', error);
      assert.ok(result.status === 'declined' && result.reason?.includes('AADSTS50199'));
    }
  });

  it('promotes a missing tenant consent to DelegatedConsentRequiredError', async () => {
    const flowHandle = await handle();
    await assert.rejects(
      () =>
        client([
          {
            status: 400,
            body: {
              error: 'invalid_grant',
              error_description:
                'AADSTS65001: The user or administrator has not consented to use the application.',
            },
          },
        ]).pollDeviceCode({ flowHandle }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedConsentRequiredError);
        assert.equal(err.adminConsentUrl, adminConsentUrl(TENANT_ID, CLIENT_ID));
        assert.deepEqual(err.requiredScopes, DELEGATED_PUBLISH_SCOPES);
        return true;
      },
    );
  });

  it('returns a token set on success, including the refresh token and the account', async () => {
    const flowHandle = await handle();
    const result = await client([
      {
        status: 200,
        body: {
          token_type: 'Bearer',
          scope: 'AppCatalog.ReadWrite.All profile openid',
          expires_in: 3599,
          access_token: 'access-1',
          refresh_token: 'refresh-1',
          id_token: idToken({
            preferred_username: 'admin@contoso.com',
            name: 'Contoso Admin',
            oid: 'user-oid',
            tid: TENANT_ID,
          }),
        },
      },
    ]).pollDeviceCode({ flowHandle });

    assert.equal(result.status, 'succeeded');
    assert.ok(result.status === 'succeeded');
    const { tokens } = result;
    assert.equal(tokens.accessToken, 'access-1');
    assert.equal(tokens.refreshToken, 'refresh-1');
    // Carried on the token set so a later refresh needs no Graph round trip
    // to re-resolve the publisher app.
    assert.equal(tokens.clientId, CLIENT_ID);
    assert.equal(tokens.tenantId, TENANT_ID);
    assert.equal(tokens.account?.username, 'admin@contoso.com');
    assert.equal(tokens.account?.displayName, 'Contoso Admin');
    // Entra answers with SHORT scope names even when the request was
    // fully qualified — the coverage check has to accept both spellings.
    assert.ok(coversCatalogPublish(tokens.scopes));
  });

  it('fails loudly when the tenant issues no refresh token', async () => {
    // Silently accepting this would look fine for an hour and then start
    // demanding an admin sign-in per upload — exactly the outcome this design
    // exists to avoid, so it has to be a hard failure at sign-in time.
    const flowHandle = await handle();
    await assert.rejects(
      () =>
        client([
          { status: 200, body: { access_token: 'access-1', expires_in: 3599 } },
        ]).pollDeviceCode({ flowHandle }),
      (err: unknown) => {
        assert.ok(err instanceof DeviceCodeFlowError);
        assert.match(err.message, /offline_access/);
        return true;
      },
    );
  });

  it('survives the process boundary: any client can poll any handle', async () => {
    // The whole reason for talking to /devicecode directly instead of using
    // MSAL's blocking helper. A restart between phase 1 and phase 2 must not
    // strand the operator.
    const flowHandle = await handle();
    const rehydrated = new DelegatedAuthClient({
      tenantId: 'a-completely-different-tenant',
      clientId: 'a-completely-different-client',
      fetchImpl: mockLogin(
        [{ status: 200, body: { access_token: 'a', refresh_token: 'r', expires_in: 60 } }],
        [],
      ),
      log: () => {},
    });
    const result = await rehydrated.pollDeviceCode({ flowHandle });
    assert.equal(result.status, 'succeeded');
    assert.ok(result.status === 'succeeded' && result.tokens.tenantId === TENANT_ID);
  });
});

describe('DelegatedAuthClient.refresh', () => {
  const stored: DelegatedTokenSet = {
    accessToken: 'old-access',
    refreshToken: 'old-refresh',
    expiresAt: new Date(Date.now() - 1000).toISOString(),
    scopes: [APP_CATALOG_DELEGATED_SCOPE],
    clientId: CLIENT_ID,
    tenantId: TENANT_ID,
  };

  it('returns the ROTATED refresh token so the caller can persist it', async () => {
    // Entra hands out a new refresh token on every exchange and the old one
    // ages out. A caller that keeps writing back the original eventually needs
    // a pointless second admin sign-in.
    const calls: FetchCall[] = [];
    const tokens = await client(
      [
        {
          status: 200,
          body: {
            access_token: 'new-access',
            refresh_token: 'rotated-refresh',
            expires_in: 3599,
          },
        },
      ],
      calls,
    ).refresh({ refreshToken: 'old-refresh' });

    assert.equal(tokens.accessToken, 'new-access');
    assert.equal(tokens.refreshToken, 'rotated-refresh');
    const [call] = calls;
    assert.ok(call);
    assert.equal(call.body.get('grant_type'), 'refresh_token');
    assert.equal(call.body.get('refresh_token'), 'old-refresh');
    // A public client must never send a secret — there is none to send.
    assert.equal(call.body.get('client_secret'), null);
  });

  it('carries the previous refresh token forward when Entra rotates nothing', async () => {
    // Dropping it would strand a still-valid credential over an optional field.
    const tokens = await client([
      { status: 200, body: { access_token: 'new-access', expires_in: 3599 } },
    ]).refresh({ refreshToken: 'old-refresh' });

    assert.equal(tokens.refreshToken, 'old-refresh');
  });

  it('maps invalid_grant to the non-recoverable DelegatedTokenExpiredError', async () => {
    await assert.rejects(
      () =>
        client([
          {
            status: 400,
            body: { error: 'invalid_grant', error_description: 'AADSTS70008: expired' },
          },
        ]).refresh({ refreshToken: 'old-refresh' }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedTokenExpiredError);
        assert.equal(err.reason, 'refresh-token-invalid');
        assert.equal(err.recoverableByRefresh, false);
        assert.ok(err instanceof TeamsProvisionerError);
        // A dead credential is a verdict, not a hiccup — retrying it forever
        // would hide the fact that a human has to act.
        assert.ok(!isTransientProvisioningFailure(err));
        return true;
      },
    );
  });

  it('ensureFreshToken refreshes only when the access token is stale', async () => {
    const fresh: DelegatedTokenSet = {
      ...stored,
      expiresAt: new Date(Date.now() + 3600_000).toISOString(),
    };
    const calls: FetchCall[] = [];
    const untouched = await client([START_OK], calls).ensureFreshToken(fresh);
    assert.equal(untouched.refreshed, false);
    assert.equal(calls.length, 0, 'a valid token must not cost a token request');

    const renewed = await client([
      {
        status: 200,
        body: { access_token: 'a2', refresh_token: 'r2', expires_in: 3599 },
      },
    ]).ensureFreshToken(stored);
    assert.equal(renewed.refreshed, true);
    assert.equal(renewed.tokens.accessToken, 'a2');
  });
});

describe('delegated sign-in status and revoke are pure', () => {
  const tokens: DelegatedTokenSet = {
    accessToken: 'a',
    refreshToken: 'r',
    expiresAt: new Date(Date.now() + 3600_000).toISOString(),
    scopes: ['AppCatalog.ReadWrite.All'],
    clientId: CLIENT_ID,
    tenantId: TENANT_ID,
    account: { username: 'admin@contoso.com' },
  };

  it('reports signed-out with the scopes a sign-in would need', () => {
    const status = describeSignInStatus(undefined);
    assert.equal(status.signedIn, false);
    assert.ok(status.signedIn === false && status.requiredScopes.length > 0);
  });

  it('reports signed-in without ever exposing a token', () => {
    const status = describeSignInStatus(tokens);
    assert.equal(status.signedIn, true);
    assert.ok(status.signedIn === true);
    assert.equal(status.accessTokenStale, false);
    assert.equal(status.coversCatalogPublish, true);
    assert.equal(status.account?.username, 'admin@contoso.com');
    const serialised = JSON.stringify(status);
    assert.ok(!serialised.includes('"a"') && !serialised.includes('"r"'));
  });

  it('treats a nearly-expired access token as stale, but still signed in', () => {
    // A stale access token is not a signed-out state: the refresh token is
    // what keeps the sign-in alive and it long outlives the hourly one.
    const stale: DelegatedTokenSet = {
      ...tokens,
      expiresAt: new Date(Date.now() + 30_000).toISOString(),
    };
    assert.ok(isAccessTokenStale(stale));
    const status = describeSignInStatus(stale);
    assert.ok(status.signedIn === true && status.accessTokenStale);
  });

  it('revoke tells the caller to discard, and where an admin withdraws consent', () => {
    assert.equal(revokeInstructions(undefined).outcome, 'not-signed-in');
    const result = revokeInstructions(tokens);
    assert.equal(result.outcome, 'discard-stored-tokens');
    assert.equal(result.adminConsentUrl, adminConsentUrl(TENANT_ID, CLIENT_ID));
    // Deleting the publisher app would reserve its uniqueName for 30 days and
    // lock the tenant out of signing in (byte5ai/omadia#916) — the advice must
    // say so rather than tempting an operator into it.
    assert.match(result.note, /30 days/);
  });
});

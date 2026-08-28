import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { CatalogUploadClient } from '../src/teamsProvisioner/catalog.js';
import {
  DelegatedAuthClient,
  type DelegatedTokenSet,
} from '../src/teamsProvisioner/delegatedAuth.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import { redactSecrets, redactUnknown } from '../src/teamsProvisioner/redact.js';

// THE security suite for the delegated publish path (byte5ai/omadia#924).
//
// Four values must never leave this subsystem in a log line, an error message,
// an error's `cause` chain, or any field of a returned object: the delegated
// ACCESS token, the REFRESH token, the DEVICE CODE, and the app CLIENT SECRET.
// Each of them buys real access if it lands in a log file that outlives it.
//
// The tests below drive the real code paths with recognisable sentinel values
// and then assert on EVERYTHING the code produced — every captured log line,
// every message in the cause chain, every enumerable field of the result. That
// is deliberately broader than "the message does not contain the token":
// `cause` is exactly where a leaked credential hides from a message assertion.

/** Values that must never appear anywhere in output. Each is unique. */
const ACCESS_TOKEN =
  'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.QUNDRVNTX1RPS0VOX1NFTlRJTkVMX0RPX05PVF9MT0c.c2lnbmF0dXJlX3NlbnRpbmVsX3ZhbHVl';
const REFRESH_TOKEN =
  'REFRESHtokenSENTINELdoNOTlogTHISvalueEVERanywhereITisAcredentialXYZ0123456789abcdefghijklmnop';
const DEVICE_CODE =
  'DEVICEcodeSENTINELdoNOTlogTHISvalueEVERitISTHEflowSECRETabcdefghijklmnopqrstuvwxyz0123456789';
const CLIENT_SECRET = 'CLIENTsecretSENTINELdoNOTlogTHISvalueEVERabcdefghijklmnopqrstuvwxyz012345';

const FORBIDDEN = [
  ACCESS_TOKEN,
  REFRESH_TOKEN,
  DEVICE_CODE,
  CLIENT_SECRET,
] as const;

/** Short distinctive fragments — catches a TRUNCATED leak, which is still a leak. */
const FORBIDDEN_FRAGMENTS = [
  'ACCESS_TOKEN_SENTINEL',
  'REFRESHtokenSENTINEL',
  'DEVICEcodeSENTINEL',
  'CLIENTsecretSENTINEL',
] as const;

function assertNoSecret(haystack: string, where: string): void {
  for (const secret of FORBIDDEN) {
    assert.ok(
      !haystack.includes(secret),
      `${where} leaked a credential verbatim:\n${haystack}`,
    );
  }
  for (const fragment of FORBIDDEN_FRAGMENTS) {
    assert.ok(
      !haystack.includes(fragment),
      `${where} leaked a credential fragment '${fragment}':\n${haystack}`,
    );
  }
}

/** Flatten an error and its whole `cause` chain into inspectable text. */
function errorSurface(err: unknown): string {
  const parts: string[] = [];
  let current: unknown = err;
  for (let depth = 0; depth < 10 && current !== undefined && current !== null; depth += 1) {
    if (current instanceof Error) {
      parts.push(current.name, current.message, current.stack ?? '');
      // Own enumerable fields, e.g. a structured `response` someone attached.
      for (const [key, value] of Object.entries(current)) {
        parts.push(key, safeStringify(value));
      }
      current = (current as { cause?: unknown }).cause;
      continue;
    }
    parts.push(safeStringify(current));
    break;
  }
  return parts.join('\n');
}

function safeStringify(value: unknown): string {
  try {
    return typeof value === 'string' ? value : JSON.stringify(value) ?? String(value);
  } catch {
    return String(value);
  }
}

interface ResponseSpec {
  status: number;
  body?: unknown;
  text?: string;
  headers?: Record<string, string>;
}

function makeResponse(spec: ResponseSpec): Response {
  const headerMap = new Map(
    Object.entries(spec.headers ?? {}).map(([k, v]) => [k.toLowerCase(), v]),
  );
  const text =
    spec.text ?? (spec.body === undefined ? '' : JSON.stringify(spec.body));
  return {
    ok: spec.status >= 200 && spec.status < 300,
    status: spec.status,
    headers: { get: (name: string) => headerMap.get(name.toLowerCase()) ?? null },
    json: async () => {
      if (text === '') throw new Error('no body');
      return JSON.parse(text);
    },
    text: async () => text,
  } as unknown as Response;
}

const TOKENS: DelegatedTokenSet = {
  accessToken: ACCESS_TOKEN,
  refreshToken: REFRESH_TOKEN,
  expiresAt: new Date(Date.now() + 3600_000).toISOString(),
  scopes: ['https://graph.microsoft.com/AppCatalog.ReadWrite.All'],
  clientId: '11111111-2222-3333-4444-555555555555',
  tenantId: '99999999-8888-7777-6666-555555555555',
};

describe('redactSecrets scrubs every credential shape', () => {
  it('removes JWTs, long opaque tokens, bearer headers and named fields', () => {
    const raw =
      `Authorization: Bearer ${ACCESS_TOKEN}\n` +
      `{"access_token":"${ACCESS_TOKEN}","refresh_token":"${REFRESH_TOKEN}"}\n` +
      `grant_type=refresh_token&refresh_token=${REFRESH_TOKEN}&client_id=abc\n` +
      `device_code=${DEVICE_CODE}\n` +
      `client_secret=${CLIENT_SECRET}\n` +
      `bare token: ${REFRESH_TOKEN}`;
    assertNoSecret(redactSecrets(raw), 'redactSecrets');
  });

  it('keeps the user_code readable — it exists to be displayed', () => {
    // Redacting the one value the operator has to type would break the flow
    // this whole module exists to run.
    const redacted = redactSecrets('{"user_code":"BXTM4NCD","interval":5}');
    assert.ok(redacted.includes('BXTM4NCD'), redacted);
  });

  it('does not eat ordinary provisioning identifiers', () => {
    // Over-redaction is acceptable in principle, but not to the point where a
    // GUID or a uniqueName stops being readable in a log.
    const text =
      'appId=11111111-2222-3333-4444-555555555555 ' +
      'uniqueName=omadia-teams-publisher-99999999-8888-7777-6666-555555555555';
    const redacted = redactSecrets(text);
    assert.ok(redacted.includes('11111111-2222-3333-4444-555555555555'), redacted);
    assert.ok(redacted.includes('omadia-teams-publisher'), redacted);
  });

  it('never throws on a hostile value', () => {
    const hostile = {
      toString() {
        throw new Error('nope');
      },
    };
    assert.equal(redactUnknown(hostile), '[unprintable]');
  });
});

describe('the delegated upload never leaks a token on the failure path', () => {
  /**
   * Graph 403 on the upload. The body deliberately echoes the bearer token —
   * some gateways really do that — so this asserts the http layer's redaction,
   * not just our own message text.
   */
  it('403 → DelegatedConsentRequiredError with no credential in the cause chain', async () => {
    const logs: string[] = [];
    const fetchImpl = (async () =>
      makeResponse({
        status: 403,
        text: JSON.stringify({
          error: {
            code: 'Authorization_RequestDenied',
            message: `Insufficient privileges. token=${ACCESS_TOKEN}`,
          },
        }),
      })) as unknown as typeof fetch;

    const client = new CatalogUploadClient({
      http: new ProvisioningHttp({
        graphCredential: {
          tenantId: TOKENS.tenantId,
          clientId: TOKENS.clientId,
          clientSecret: CLIENT_SECRET,
        },
        fetchImpl,
        log: (msg) => logs.push(msg),
      }),
      delegatedAuth: {
        ensureFreshToken: async (tokens) => ({ tokens, refreshed: false }),
        adminConsentUrlFor: () => 'https://login.microsoftonline.com/t/adminconsent?client_id=c',
      },
      log: (msg) => logs.push(msg),
    });

    await assert.rejects(
      () => client.uploadToCatalogDelegated({
        packageZip: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
        externalId: 'b5561ec9-8cab-4aa3-8aa2-d8d7172e4311',
        tokens: TOKENS,
      }),
      (err: unknown) => {
        assert.equal((err as Error).name, 'DelegatedConsentRequiredError');
        assertNoSecret(errorSurface(err), 'DelegatedConsentRequiredError');
        return true;
      },
    );
    assertNoSecret(logs.join('\n'), 'upload 403 logs');
  });

  it('an unexpected 500 keeps the token out of the generic request error', async () => {
    const logs: string[] = [];
    const fetchImpl = (async () =>
      makeResponse({
        status: 500,
        text: `upstream failure; echoed authorization: Bearer ${ACCESS_TOKEN}`,
      })) as unknown as typeof fetch;

    const client = new CatalogUploadClient({
      http: new ProvisioningHttp({
        graphCredential: {
          tenantId: TOKENS.tenantId,
          clientId: TOKENS.clientId,
          clientSecret: CLIENT_SECRET,
        },
        fetchImpl,
        log: (msg) => logs.push(msg),
      }),
      log: (msg) => logs.push(msg),
    });

    await assert.rejects(
      () => client.uploadToCatalog({
        packageZip: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
        externalId: 'b5561ec9-8cab-4aa3-8aa2-d8d7172e4311',
        delegatedAccessToken: ACCESS_TOKEN,
      }),
      (err: unknown) => {
        assertNoSecret(errorSurface(err), 'ProvisioningRequestError');
        return true;
      },
    );
    assertNoSecret(logs.join('\n'), 'upload 500 logs');
  });
});

describe('the device-code flow never leaks the device code', () => {
  it('the start result logs and returns no device_code', async () => {
    const logs: string[] = [];
    const fetchImpl = (async () =>
      makeResponse({
        status: 200,
        body: {
          device_code: DEVICE_CODE,
          user_code: 'BXTM4NCD',
          verification_uri: 'https://microsoft.com/devicelogin',
          expires_in: 900,
          interval: 5,
          message: 'To sign in, use a web browser to open the page…',
        },
      })) as unknown as typeof fetch;

    const auth = new DelegatedAuthClient({
      tenantId: TOKENS.tenantId,
      clientId: TOKENS.clientId,
      fetchImpl,
      log: (msg) => logs.push(msg),
    });
    const start = await auth.startDeviceCode();

    assert.equal(start.userCode, 'BXTM4NCD');
    assertNoSecret(logs.join('\n'), 'startDeviceCode logs');

    // Every field EXCEPT the handle is display material and must be clean. The
    // handle is documented as secret-grade and does carry the device code — the
    // test pins that boundary so nobody widens it by accident.
    const { flowHandle, ...displayable } = start;
    assertNoSecret(JSON.stringify(displayable), 'DeviceCodeStart display fields');
    assert.ok(flowHandle.length > 0);
  });

  it('a malformed handle is rejected without echoing it back', async () => {
    const auth = new DelegatedAuthClient({
      tenantId: TOKENS.tenantId,
      clientId: TOKENS.clientId,
      fetchImpl: (async () => makeResponse({ status: 200, body: {} })) as unknown as typeof fetch,
    });
    await assert.rejects(
      () => auth.pollDeviceCode({ flowHandle: `not-base64-${DEVICE_CODE}` }),
      (err: unknown) => {
        assert.equal((err as Error).name, 'DeviceCodeFlowError');
        assertNoSecret(errorSurface(err), 'malformed handle error');
        return true;
      },
    );
  });

  it('a failing refresh keeps the refresh token out of the error', async () => {
    const logs: string[] = [];
    const fetchImpl = (async () =>
      makeResponse({
        status: 400,
        body: {
          error: 'invalid_grant',
          error_description: `AADSTS70008: the refresh token ${REFRESH_TOKEN} has expired.`,
        },
      })) as unknown as typeof fetch;

    const auth = new DelegatedAuthClient({
      tenantId: TOKENS.tenantId,
      clientId: TOKENS.clientId,
      fetchImpl,
      log: (msg) => logs.push(msg),
    });

    await assert.rejects(
      () => auth.refresh({ refreshToken: REFRESH_TOKEN }),
      (err: unknown) => {
        assert.equal((err as Error).name, 'DelegatedTokenExpiredError');
        assertNoSecret(errorSurface(err), 'DelegatedTokenExpiredError');
        return true;
      },
    );
    assertNoSecret(logs.join('\n'), 'refresh failure logs');
  });

  it('a declined poll reports its reason with the credential stripped out', async () => {
    const fetchImpl = (async (_url: string, init?: RequestInit) => {
      const body = String(init?.body ?? '');
      if (body.includes('devicecode') || !body.includes('grant_type')) {
        return makeResponse({ status: 200, body: {} });
      }
      return makeResponse({
        status: 400,
        body: {
          error: 'access_denied',
          error_description: `AADSTS50199: blocked by policy. device_code=${DEVICE_CODE}`,
        },
      });
    }) as unknown as typeof fetch;

    const auth = new DelegatedAuthClient({
      tenantId: TOKENS.tenantId,
      clientId: TOKENS.clientId,
      fetchImpl,
    });
    const handle = Buffer.from(
      JSON.stringify({
        v: 1,
        deviceCode: DEVICE_CODE,
        tenantId: TOKENS.tenantId,
        clientId: TOKENS.clientId,
        intervalSeconds: 5,
        expiresAt: new Date(Date.now() + 900_000).toISOString(),
      }),
      'utf8',
    ).toString('base64url');

    const result = await auth.pollDeviceCode({ flowHandle: handle });
    assert.equal(result.status, 'declined');
    assert.ok(result.status === 'declined' && result.reason !== undefined);
    // The AADSTS code survives — that is what tells a cancelled sign-in apart
    // from a Conditional-Access block — while the credential does not.
    assert.ok(result.status === 'declined' && result.reason?.includes('AADSTS50199'));
    assertNoSecret(JSON.stringify(result), 'declined poll result');
  });
});

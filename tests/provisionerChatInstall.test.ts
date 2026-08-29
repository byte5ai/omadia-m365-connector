import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  CHAT_INSTALL_SCOPE,
  ChatInstallClient,
} from '../src/teamsProvisioner/install.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  ChatNotFoundError,
  ConsentMissingError,
  InstallTargetMismatchError,
  ProvisioningRequestError,
  ProvisioningThrottledError,
  TeamsProvisionerError,
  isTransientProvisioningFailure,
} from '../src/teamsProvisioner/errors.js';
import type {
  ChatAppInstallation,
  Idempotent,
  InstallToChatInput,
} from '../src/teamsProvisioner/types.js';

// The chat-install step: POST /chats/{id}/installedApps behind the shared
// ProvisioningHttp choke point. Same mocked-fetch style as
// tests/provisionerInstall.test.ts, because the two directions are twins —
// what is tested HERE is where they are not: the app role, the pre-flight
// target check, the typed chat-specific 404, and the fact that "already
// installed" may arrive Bad-Request-shaped rather than as a 409.

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

const CHAT_ID = '19:0d8ba2ab5a5b4f2f9b0f2f7d0f5b6a11@thread.v2';
const ONE_ON_ONE_ID = '19:aaaa_bbbb@unq.gbl.spaces';
const CHANNEL_ID = '19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2';
const TEAM_GUID = 'abc8af8e-c7fc-4717-85d3-b83c4d84b667';
/** 32 hex, no dashes, no prefix/suffix — a team GUID AND a chat body. */
const AMBIGUOUS_ID = 'abc8af8ec7fc471785d3b83c4d84b667';

const TEAMS_APP_ID = 'catalog-app-0001';
const INSTALLATION_ID = 'NmFiOTZlZm-installation-id';

const CHAT_PATH = `/chats/${encodeURIComponent(CHAT_ID)}/installedApps`;
const CHAT_URL = `https://graph.microsoft.com/v1.0${CHAT_PATH}`;

function harness(routes: Route[]): {
  client: ChatInstallClient;
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
  const client = new ChatInstallClient({ http, log: () => {} });
  return { client, calls, sleeps };
}

/** Graph calls only (token POSTs filtered out). */
function graphCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter((c) => c.url.includes('graph.microsoft.com'));
}

function parsedBody(call: FetchCall): Record<string, unknown> {
  return JSON.parse(String(call.init?.body ?? '{}')) as Record<string, unknown>;
}

// ---------------------------------------------------------------------------
// Contract shape. The middleware compiles against the structural signature
// `(input: { chatId, teamsAppId }) => Promise<Idempotent<{ chatId, teamsAppId }>>`
// rather than importing our types, so assert that our richer result (which
// adds the optional installationId) still satisfies it.
// ---------------------------------------------------------------------------
type MiddlewareInstallToChat = (input: {
  readonly chatId: string;
  readonly teamsAppId: string;
}) => Promise<Idempotent<{ readonly chatId: string; readonly teamsAppId: string }>>;

const _contractCheck: MiddlewareInstallToChat = (
  input: InstallToChatInput,
): Promise<Idempotent<ChatAppInstallation>> =>
  new ChatInstallClient({
    http: { request: async () => ({ kind: 'ok', status: 201, json: undefined, header: () => null }) },
  }).installToChat(input);
void _contractCheck;

describe('ChatInstallClient.installToChat', () => {
  it('POSTs the odata.bind body to /chats and returns created with the installation id', async () => {
    const { client, calls } = harness([
      route(CHAT_PATH, { status: 201, body: { id: INSTALLATION_ID } }),
    ]);

    const result = await client.installToChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, {
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
      installationId: INSTALLATION_ID,
    });

    const [call] = graphCalls(calls);
    assert.ok(call, 'expected one install POST');
    assert.equal(call.init?.method, 'POST');
    assert.equal(call.url, CHAT_URL);
    const body = parsedBody(call);
    assert.equal(
      body['teamsApp@odata.bind'],
      `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${TEAMS_APP_ID}`,
    );
    // The chat scope cannot consent RSC — the key must never appear.
    assert.ok(!('consentedPermissionSet' in body));
    const headers = call.init?.headers as Record<string, string>;
    assert.equal(headers['Authorization'], 'Bearer tok');
    assert.equal(headers['content-type'], 'application/json');
  });

  it('omits installationId when the 201 has no body', async () => {
    const { client } = harness([route(CHAT_PATH, { status: 201 })]);

    const result = await client.installToChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.deepEqual(result.value, {
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });
    assert.ok(!('installationId' in result.value));
  });

  it('installs into a 1:1 chat (@unq.gbl.spaces) too', async () => {
    const path = `/chats/${encodeURIComponent(ONE_ON_ONE_ID)}/installedApps`;
    const { client, calls } = harness([route(path, { status: 201 })]);

    const result = await client.installToChat({
      chatId: ONE_ON_ONE_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.equal(graphCalls(calls).length, 1);
  });
});

describe('ChatInstallClient.installToChat — already installed is idempotent', () => {
  it('resolves a 409 Conflict to already-existed — success, no throw', async () => {
    // The observed duplicate shape of the sibling scopes: code 'Conflict',
    // message 'AppEntitlement id: … already exists in <scope>'. The scope tail
    // differs per endpoint ("TeamId: '19:…'" for teams, "thread" for the
    // per-user endpoint), which is exactly why nothing here reads the message.
    const { client, calls } = harness([
      route(CHAT_PATH, {
        status: 409,
        body: {
          error: {
            code: 'Conflict',
            message: `AppEntitlement id: '${TEAMS_APP_ID}' already exists in ChatId: '${CHAT_ID}'`,
          },
        },
      }),
    ]);

    const result = await client.installToChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-existed');
    assert.deepEqual(result.value, {
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });
    assert.equal(graphCalls(calls).length, 1);
  });

  it('also resolves a Bad-Request-shaped duplicate (400 + "already exists")', async () => {
    // Graph documents NO failure response for this verb at all, and Entra has
    // form for reporting a duplicate as a 400 (byte5ai/omadia#916). Both
    // shapes must land on the same idempotent path. 'already exists' is the
    // wording the documented sibling scopes actually use.
    const { client } = harness([
      route(CHAT_PATH, {
        status: 400,
        body: {
          error: {
            code: 'BadRequest',
            message: `AppEntitlement id: '${TEAMS_APP_ID}' already exists in ChatId: '${CHAT_ID}'`,
          },
        },
      }),
    ]);

    const result = await client.installToChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-existed');
  });

  it('resolves a 400 that still carries the Conflict code', async () => {
    // The CODE, not the prose, is the load-bearing signal — a 400 wrapping
    // the same 'Conflict' code is a duplicate however it is worded.
    const { client } = harness([
      route(CHAT_PATH, {
        status: 400,
        body: { error: { code: 'Conflict', message: 'unhelpfully worded' } },
      }),
    ]);

    const result = await client.installToChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-existed');
  });

  it('does NOT swallow an ordinary 400 — the leniency is duplicate-only', async () => {
    // The real 400 this endpoint is known to answer: the app declares RSC
    // permissions the caller's app role cannot consent to. It must stay a
    // failure — reporting it as 'already-existed' would claim an install that
    // never happened.
    const { client } = harness([
      route(CHAT_PATH, {
        status: 400,
        body: {
          error: {
            code: 'ResourceSpecificPermissionsMismatch',
            message:
              'The app requires resource-specific permissions that were not consented.',
          },
        },
      }),
    ]);

    await assert.rejects(
      client.installToChat({ chatId: CHAT_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ProvisioningRequestError);
        assert.equal(err.status, 400);
        assert.equal(err.step, 'chats.installedApps.add');
        return true;
      },
    );
  });
});

describe('ChatInstallClient.installToChat — a missing chat is distinguishable', () => {
  it('throws ChatNotFoundError (not a generic request error) on 404', async () => {
    const { client } = harness([
      route(CHAT_PATH, { status: 404, body: { error: { code: 'NotFound' } } }),
    ]);

    await assert.rejects(
      client.installToChat({ chatId: CHAT_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ChatNotFoundError);
        // Distinguishable from the team direction's not-found handling, which
        // never produces this class.
        assert.ok(!(err instanceof ProvisioningRequestError));
        assert.ok(err instanceof TeamsProvisionerError);
        assert.equal(err.chatId, CHAT_ID);
        assert.equal(err.step, 'chats.installedApps.add');
        return true;
      },
    );
  });

  it('names the channel-id confusion and the app role in the message', async () => {
    const { client } = harness([route(CHAT_PATH, { status: 404 })]);

    await assert.rejects(
      client.installToChat({ chatId: CHAT_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ChatNotFoundError);
        assert.match(err.message, /thread\.tacv2/);
        assert.match(err.message, /ReadWriteForChat\.All/);
        return true;
      },
    );
  });

  it('is not transient — retrying the identical call cannot help', async () => {
    assert.equal(
      isTransientProvisioningFailure(
        new ChatNotFoundError(CHAT_ID, 'chats.installedApps.add'),
      ),
      false,
    );
  });
});

describe('ChatInstallClient.installToChat — pre-flight target check', () => {
  it('refuses a CHANNEL id before any network call and points at the team', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.installToChat({ chatId: CHANNEL_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof InstallTargetMismatchError);
        assert.equal(err.targetKind, 'channel');
        assert.equal(err.value, CHANNEL_ID);
        assert.match(err.hint.toLowerCase(), /team id/);
        assert.equal(isTransientProvisioningFailure(err), false);
        return true;
      },
    );
    assert.equal(calls.length, 0, 'must not reach Graph');
  });

  it('refuses a team GUID and names installToTeam', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.installToChat({ chatId: TEAM_GUID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof InstallTargetMismatchError);
        assert.equal(err.targetKind, 'team');
        assert.match(err.hint, /installToTeam/);
        return true;
      },
    );
    assert.equal(calls.length, 0);
  });

  it('refuses the ambiguous 32-hex form and asks for the full chat form', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.installToChat({ chatId: AMBIGUOUS_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof InstallTargetMismatchError);
        assert.equal(err.targetKind, 'ambiguous');
        assert.match(err.hint, /thread\.v2/);
        return true;
      },
    );
    assert.equal(calls.length, 0);
  });

  it('rejects empty chatId / teamsAppId before any fetch', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.installToChat({ chatId: '  ', teamsAppId: TEAMS_APP_ID }),
      /invalid_argument: 'chatId'/,
    );
    await assert.rejects(
      client.installToChat({ chatId: CHAT_ID, teamsAppId: '' }),
      /invalid_argument: 'teamsAppId'/,
    );
    assert.equal(calls.length, 0);
  });
});

describe('ChatInstallClient.installToChat — shared 403/429 paths', () => {
  it('maps 403 to ConsentMissingError carrying the CHAT app role', async () => {
    const { client } = harness([
      route(CHAT_PATH, { status: 403, body: { error: { code: 'Forbidden' } } }),
    ]);

    await assert.rejects(
      client.installToChat({ chatId: CHAT_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.ok(err instanceof TeamsProvisionerError);
        assert.deepEqual(err.missingScopes, [CHAT_INSTALL_SCOPE]);
        assert.equal(
          CHAT_INSTALL_SCOPE,
          'TeamsAppInstallation.ReadWriteForChat.All',
        );
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });

  it('retries 429 honouring Retry-After, then succeeds', async () => {
    const { client, calls, sleeps } = harness([
      route(
        CHAT_PATH,
        { status: 429, headers: { 'Retry-After': '7' } },
        { status: 201, body: { id: INSTALLATION_ID } },
      ),
    ]);

    const result = await client.installToChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'created');
    assert.equal(graphCalls(calls).length, 2);
    assert.deepEqual(sleeps, [7000]);
  });

  it('throws ProvisioningThrottledError when the 429 budget is exhausted', async () => {
    const { client, calls } = harness([
      route(CHAT_PATH, { status: 429, headers: { 'Retry-After': '3' } }),
    ]);

    await assert.rejects(
      client.installToChat({ chatId: CHAT_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ProvisioningThrottledError);
        assert.equal(err.retryAfterSeconds, 3);
        return true;
      },
    );
    assert.equal(graphCalls(calls).length, 3);
  });
});

// ---------------------------------------------------------------------------
// The uninstall direction — the chat mirror of uninstallFromTeam
// (byte5ai/omadia#900): lookup-then-DELETE, and every flavour of "not there"
// collapses into the idempotent 'already-absent'.
// ---------------------------------------------------------------------------

/** Only the lookup GETs (they carry a query string). */
function lookupCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter(
    (c) => c.url.includes(CHAT_PATH) && c.url.includes('$filter='),
  );
}

function deleteCalls(calls: FetchCall[]): FetchCall[] {
  return calls.filter((c) => c.init?.method === 'DELETE');
}

const foundBody = (
  installationId = INSTALLATION_ID,
  teamsAppId = TEAMS_APP_ID,
): unknown => ({
  value: [{ id: installationId, teamsApp: { id: teamsAppId } }],
});

describe('ChatInstallClient.uninstallFromChat', () => {
  it('resolves the installation id, DELETEs it, and reports uninstalled', async () => {
    const { client, calls } = harness([
      route(`${CHAT_PATH}?`, { status: 200, body: foundBody() }),
      route(`${CHAT_PATH}/${INSTALLATION_ID}`, { status: 204 }),
    ]);

    const result = await client.uninstallFromChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'uninstalled');
    assert.deepEqual(result.value, {
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
      installationId: INSTALLATION_ID,
    });

    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected one lookup GET');
    assert.equal(lookup.init?.method, 'GET');
    assert.equal(
      lookup.url,
      `${CHAT_URL}?$expand=teamsApp&$filter=${encodeURIComponent(`teamsApp/id eq '${TEAMS_APP_ID}'`)}`,
    );

    const [del] = deleteCalls(calls);
    assert.ok(del, 'expected one DELETE');
    assert.equal(del.url, `${CHAT_URL}/${INSTALLATION_ID}`);
  });

  it('reports already-absent WITHOUT deleting when the lookup finds nothing', async () => {
    const { client, calls } = harness([
      route(`${CHAT_PATH}?`, { status: 200, body: { value: [] } }),
    ]);

    const result = await client.uninstallFromChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.deepEqual(result.value, {
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });
    assert.equal(deleteCalls(calls).length, 0, 'must not DELETE on a lookup miss');
  });

  it('ignores lookup entries whose expanded teamsApp is a different app', async () => {
    const { client, calls } = harness([
      route(`${CHAT_PATH}?`, {
        status: 200,
        body: { value: [{ id: 'other-installation', teamsApp: { id: 'other-app' } }] },
      }),
    ]);

    const result = await client.uninstallFromChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.equal(deleteCalls(calls).length, 0);
  });

  it('treats a 404 on the DELETE as already-absent (removal race)', async () => {
    const { client, calls } = harness([
      route(`${CHAT_PATH}?`, { status: 200, body: foundBody() }),
      route(`${CHAT_PATH}/${INSTALLATION_ID}`, { status: 404 }),
    ]);

    const result = await client.uninstallFromChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.equal(result.value.installationId, INSTALLATION_ID);
    assert.equal(deleteCalls(calls).length, 1);
  });

  it('treats a 404 on the lookup (chat gone) as already-absent, NOT ChatNotFoundError', async () => {
    const { client, calls } = harness([
      route(`${CHAT_PATH}?`, { status: 404, body: { error: { code: 'NotFound' } } }),
    ]);

    // Being asked to make sure an app is absent from a chat that no longer
    // exists is a request already satisfied — unlike the install direction.
    const result = await client.uninstallFromChat({
      chatId: CHAT_ID,
      teamsAppId: TEAMS_APP_ID,
    });

    assert.equal(result.outcome, 'already-absent');
    assert.equal(deleteCalls(calls).length, 0);
  });

  it('escapes OData quotes in the teamsAppId filter (quote doubling)', async () => {
    const { client, calls } = harness([
      route(`${CHAT_PATH}?`, { status: 200, body: { value: [] } }),
    ]);

    await client.uninstallFromChat({
      chatId: CHAT_ID,
      teamsAppId: "app' or id ne '",
    });

    const [lookup] = lookupCalls(calls);
    assert.ok(lookup, 'expected one lookup GET');
    assert.equal(
      new URL(lookup.url).searchParams.get('$filter'),
      "teamsApp/id eq 'app'' or id ne '''",
    );
  });

  it('maps 403 to ConsentMissingError carrying the CHAT app role', async () => {
    const { client } = harness([
      route(`${CHAT_PATH}?`, { status: 403, body: { error: { code: 'Forbidden' } } }),
    ]);

    await assert.rejects(
      client.uninstallFromChat({ chatId: CHAT_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        assert.deepEqual(err.missingScopes, [CHAT_INSTALL_SCOPE]);
        return true;
      },
    );
  });

  it('refuses a channel id before any network call', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.uninstallFromChat({ chatId: CHANNEL_ID, teamsAppId: TEAMS_APP_ID }),
      (err: unknown) => {
        assert.ok(err instanceof InstallTargetMismatchError);
        assert.equal(err.step, 'chats.installedApps.remove');
        return true;
      },
    );
    assert.equal(calls.length, 0);
  });

  it('rejects empty chatId / teamsAppId before any fetch', async () => {
    const { client, calls } = harness([]);

    await assert.rejects(
      client.uninstallFromChat({ chatId: '  ', teamsAppId: TEAMS_APP_ID }),
      /invalid_argument: 'chatId'/,
    );
    await assert.rejects(
      client.uninstallFromChat({ chatId: CHAT_ID, teamsAppId: '' }),
      /invalid_argument: 'teamsAppId'/,
    );
    assert.equal(calls.length, 0);
  });
});

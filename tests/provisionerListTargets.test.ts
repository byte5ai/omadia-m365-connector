import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  CHAT_READ_DELEGATED_SCOPE,
  type DelegatedTokenSet,
} from '../src/teamsProvisioner/delegatedAuth.js';
import {
  ConsentMissingError,
  DelegatedConsentRequiredError,
  DelegatedScopeRequiredError,
  DelegatedTokenExpiredError,
} from '../src/teamsProvisioner/errors.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  InstallTargetsClient,
  TEAM_LIST_SCOPE,
} from '../src/teamsProvisioner/targets.js';

// Install-target ENUMERATION (0.8.0): the lists a picker offers so an operator
// never types a team or chat id again. What these tests must prove:
//
//   * TEAMS come from `GET /teams` under Team.ReadBasic.All — the app role the
//     connector already has. Using /groups?$filter=… instead would need
//     Group.Read.All, which the setup guide never asks for, so it would pass
//     in a permissive tenant and 403 at the customer;
//   * PAGING is followed to the end. A tenant with 30 teams fits in one page
//     and one with 300 does not, and the difference must never be a silently
//     short picker;
//   * CHATS are delegated, because Graph has no tenant-wide app-only route for
//     them at all, and the three ways a credential can be unusable (absent,
//     too narrow, stale) each get their own typed answer BEFORE a Graph call
//     is spent producing a 401/403 someone would have to interpret;
//   * a 1:1 chat has NO topic, so member names are what make it identifiable —
//     that is the whole reason `memberNames` exists.
//
// Mocked-fetch style follows tests/provisionerTeamLookup.test.ts.

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

function harness(routes: Route[]): {
  client: InstallTargetsClient;
  calls: FetchCall[];
} {
  const calls: FetchCall[] = [];
  const http = new ProvisioningHttp({
    graphCredential: {
      tenantId: 'tenant-1',
      clientId: 'app-graph',
      clientSecret: 'graph-secret-value',
    },
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async () => {},
    max429Retries: 2,
  });
  return {
    client: new InstallTargetsClient({
      http,
      log: () => {},
      adminConsentUrlFor: () => 'https://consent.example/admin',
    }),
    calls,
  };
}

const NEXT_PAGE = 'https://graph.microsoft.com/v1.0/teams?$skiptoken=PAGE2';
const CHAT_NEXT_PAGE = 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=P2';

/** A credential that can do everything listChats needs. */
function tokens(overrides: Partial<DelegatedTokenSet> = {}): DelegatedTokenSet {
  return {
    accessToken: 'delegated-access-token',
    refreshToken: 'delegated-refresh-token',
    expiresAt: new Date(Date.now() + 30 * 60_000).toISOString(),
    scopes: [CHAT_READ_DELEGATED_SCOPE, 'offline_access'],
    clientId: 'publisher-app',
    tenantId: 'tenant-1',
    ...overrides,
  };
}

describe('InstallTargetsClient.listTeams', () => {
  it('lists teams from /teams and selects only id + displayName', async () => {
    const { client, calls } = harness([
      route('/teams', {
        status: 200,
        body: {
          value: [
            { id: 'team-a', displayName: 'Marketing' },
            { id: 'team-b', displayName: 'Engineering' },
          ],
        },
      }),
    ]);

    assert.deepEqual(await client.listTeams(), [
      { id: 'team-a', displayName: 'Marketing' },
      { id: 'team-b', displayName: 'Engineering' },
    ]);

    const call = calls[0];
    assert.ok(call);
    // /teams, NOT /groups?$filter=resourceProvisioningOptions/… — the group
    // query returns the same set but reads the directory, so it needs
    // Group.Read.All, which this connector's setup guide never asks for.
    assert.ok(call.url.includes('/v1.0/teams'));
    assert.ok(!call.url.includes('/groups'));
    // The scope boundary made literal: a picker needs an id and a label.
    assert.match(call.url, /\$select=id,displayName/);
    assert.equal(call.init?.method, 'GET');
  });

  it('follows @odata.nextLink until the last page', async () => {
    // 30 teams fit in one page; 300 do not. A picker that silently shows the
    // first hundred is the same class of bug as the one this feature fixes.
    const { client, calls } = harness([
      route('$skiptoken=PAGE2', {
        status: 200,
        body: { value: [{ id: 'team-c', displayName: 'Sales' }] },
      }),
      route('/teams', {
        status: 200,
        body: {
          value: [{ id: 'team-a', displayName: 'Marketing' }],
          '@odata.nextLink': NEXT_PAGE,
        },
      }),
    ]);

    assert.deepEqual(await client.listTeams(), [
      { id: 'team-a', displayName: 'Marketing' },
      { id: 'team-c', displayName: 'Sales' },
    ]);
    assert.equal(calls.length, 2, 'both pages must be fetched');
    assert.ok(calls[1]?.url.includes('$skiptoken=PAGE2'));
  });

  it('drops entries without a usable displayName instead of showing a blank row', async () => {
    const { client } = harness([
      route('/teams', {
        status: 200,
        body: {
          value: [
            { id: 'team-a', displayName: 'Marketing' },
            { id: 'team-b', displayName: '   ' },
            { id: '', displayName: 'Nameless id' },
            null,
          ],
        },
      }),
    ]);
    // An unlabelled option in a picker is worse than one fewer option.
    assert.deepEqual(await client.listTeams(), [
      { id: 'team-a', displayName: 'Marketing' },
    ]);
  });

  it('maps 403 to ConsentMissingError carrying Team.ReadBasic.All', async () => {
    const { client } = harness([
      route('/teams', {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);
    await assert.rejects(
      () => client.listTeams(),
      (err: unknown) => {
        assert.ok(err instanceof ConsentMissingError);
        // The same app role getTeam already needs — enabling the picker
        // consents nothing new.
        assert.deepEqual(err.missingScopes, [TEAM_LIST_SCOPE]);
        assert.equal(err.resource, 'graph');
        return true;
      },
    );
  });
});

describe('InstallTargetsClient.listChats', () => {
  it('lists the signed-in admin chats from /me/chats with the delegated token', async () => {
    const { client, calls } = harness([
      route('/me/chats', {
        status: 200,
        body: {
          value: [
            {
              id: '19:group@thread.v2',
              topic: 'Project Falcon',
              chatType: 'group',
              members: [
                { displayName: 'Ada Lovelace' },
                { displayName: 'Alan Turing' },
              ],
            },
          ],
        },
      }),
    ]);

    assert.deepEqual(await client.listChats({ tokens: tokens() }), [
      {
        id: '19:group@thread.v2',
        topic: 'Project Falcon',
        chatType: 'group',
        memberNames: ['Ada Lovelace', 'Alan Turing'],
      },
    ]);

    const call = calls[0];
    assert.ok(call);
    // /me/chats, not /users/{id}/chats: the delegated route answers for the
    // administrator who signed in, which is exactly whose chats they can drop
    // an agent into. The application-permission route is per-user and would
    // mean enumerating the whole tenant.
    assert.ok(call.url.includes('/v1.0/me/chats'));
    assert.match(call.url, /\$expand=members/);
    const headers = call.init?.headers as Record<string, string> | undefined;
    // The DELEGATED token, not the app-only one the token cache holds.
    assert.equal(headers?.['Authorization'], 'Bearer delegated-access-token');
  });

  it('keeps a 1:1 chat with topic null and identifies it by member names', async () => {
    const { client } = harness([
      route('/me/chats', {
        status: 200,
        body: {
          value: [
            {
              id: '19:solo@unq.gbl.spaces',
              topic: null,
              chatType: 'oneOnOne',
              members: [{ displayName: 'Grace Hopper' }],
            },
          ],
        },
      }),
    ]);
    const [chat] = await client.listChats({ tokens: tokens() });
    // A 1:1 chat HAS no topic — without the member names the picker would show
    // a bare `19:…` string, which is the problem this feature removes.
    assert.equal(chat?.topic, null);
    assert.deepEqual(chat?.memberNames, ['Grace Hopper']);
    assert.equal(chat?.chatType, 'oneOnOne');
  });

  it('drops chat kinds this version cannot route to an install method', async () => {
    const { client } = harness([
      route('/me/chats', {
        status: 200,
        body: {
          value: [
            { id: '19:m@thread.v2', topic: 'Standup', chatType: 'meeting' },
            { id: '19:x@thread.v2', topic: 'X', chatType: 'unknownFutureValue' },
            { id: '19:y@thread.v2', topic: 'Y' },
          ],
        },
      }),
    ]);
    // `unknownFutureValue` is on every evolvable Graph enum. Offering one is
    // putting a dead option in the picker.
    assert.deepEqual(await client.listChats({ tokens: tokens() }), [
      { id: '19:m@thread.v2', topic: 'Standup', chatType: 'meeting' },
    ]);
  });

  it('follows @odata.nextLink until the last page', async () => {
    const { client, calls } = harness([
      route('$skiptoken=P2', {
        status: 200,
        body: {
          value: [{ id: '19:b@thread.v2', topic: 'Second', chatType: 'group' }],
        },
      }),
      route('/me/chats', {
        status: 200,
        body: {
          value: [{ id: '19:a@thread.v2', topic: 'First', chatType: 'group' }],
          '@odata.nextLink': CHAT_NEXT_PAGE,
        },
      }),
    ]);
    const chats = await client.listChats({ tokens: tokens() });
    assert.deepEqual(
      chats.map((c) => c.id),
      ['19:a@thread.v2', '19:b@thread.v2'],
    );
    assert.equal(calls.length, 2);
  });

  it('throws DelegatedScopeRequiredError("no-token") without calling Graph', async () => {
    const { client, calls } = harness([]);
    await assert.rejects(
      () => client.listChats(),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedScopeRequiredError);
        assert.equal(err.reason, 'no-token');
        assert.deepEqual(err.requiredScopes, [CHAT_READ_DELEGATED_SCOPE]);
        assert.equal(err.step, 'chats.list');
        return true;
      },
    );
    // Not one Graph call spent to learn what we already knew.
    assert.equal(calls.length, 0);
  });

  it('throws DelegatedScopeRequiredError("scope-missing") for a pre-0.8.0 credential', async () => {
    // THE MIGRATION CASE, and the one an operator will actually hit: a
    // credential minted for publishing only. It cannot grow Chat.ReadBasic by
    // refreshing, so the remedy is a fresh sign-in — and saying that here is
    // better than letting Graph answer 403 and guessing at why.
    const { client, calls } = harness([]);
    const publishOnly = tokens({
      scopes: ['https://graph.microsoft.com/AppCatalog.ReadWrite.All'],
    });
    await assert.rejects(
      () => client.listChats({ tokens: publishOnly }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedScopeRequiredError);
        assert.equal(err.reason, 'scope-missing');
        assert.deepEqual(err.grantedScopes, publishOnly.scopes);
        assert.match(err.message, /sign in again/);
        return true;
      },
    );
    assert.equal(calls.length, 0);
  });

  it('accepts a wider chat scope than the least-privileged one', async () => {
    // An admin who consented to Chat.Read can list chats. Refusing them
    // because the string is not the narrowest permission would be a bug in
    // the check, not a safety property.
    const { client } = harness([
      route('/me/chats', { status: 200, body: { value: [] } }),
    ]);
    const wider = tokens({ scopes: ['https://graph.microsoft.com/Chat.Read'] });
    assert.deepEqual(await client.listChats({ tokens: wider }), []);
  });

  it('accepts the bare (unqualified) scope spelling Entra sometimes echoes', async () => {
    const { client } = harness([
      route('/me/chats', { status: 200, body: { value: [] } }),
    ]);
    assert.deepEqual(
      await client.listChats({ tokens: tokens({ scopes: ['Chat.ReadBasic'] }) }),
      [],
    );
  });

  it('reports a stale access token instead of silently refreshing it', async () => {
    const { client, calls } = harness([]);
    const stale = tokens({ expiresAt: new Date(Date.now() - 1000).toISOString() });
    await assert.rejects(
      () => client.listChats({ tokens: stale }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedTokenExpiredError);
        assert.equal(err.reason, 'access-token-expired');
        assert.equal(err.recoverableByRefresh, true);
        return true;
      },
    );
    // Refreshing here would rotate the caller's refresh token and then drop it
    // — the result carries no token set back. So it does not refresh.
    assert.equal(calls.length, 0);
  });

  it('translates a delegated 403 into DelegatedConsentRequiredError', async () => {
    const { client } = harness([
      route('/me/chats', {
        status: 403,
        body: { error: { code: 'Authorization_RequestDenied' } },
      }),
    ]);
    await assert.rejects(
      () => client.listChats({ tokens: tokens() }),
      (err: unknown) => {
        // NOT ConsentMissingError: no app identity was involved, so telling an
        // operator to grant an application permission would send them to fix
        // something that could never have helped.
        assert.ok(err instanceof DelegatedConsentRequiredError);
        assert.equal(err.adminConsentUrl, 'https://consent.example/admin');
        assert.deepEqual(err.requiredScopes, [CHAT_READ_DELEGATED_SCOPE]);
        return true;
      },
    );
  });

  it('translates a 401 into an expired access token', async () => {
    const { client } = harness([
      route('/me/chats', { status: 401, body: { error: { code: 'InvalidAuthenticationToken' } } }),
    ]);
    await assert.rejects(
      () => client.listChats({ tokens: tokens() }),
      (err: unknown) => {
        assert.ok(err instanceof DelegatedTokenExpiredError);
        assert.equal(err.reason, 'access-token-expired');
        return true;
      },
    );
  });
});

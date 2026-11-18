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
  return { ...harnessWith(mockFetch(routes, calls)), calls };
}

/** A client over any fake Graph, with its log lines and an optional clock. */
function harnessWith(
  fetchImpl: typeof fetch,
  now?: () => number,
): { client: InstallTargetsClient; logs: string[] } {
  const logs: string[] = [];
  const http = new ProvisioningHttp({
    graphCredential: {
      tenantId: 'tenant-1',
      clientId: 'app-graph',
      clientSecret: 'graph-secret-value',
    },
    fetchImpl,
    log: () => {},
    sleep: async () => {},
    max429Retries: 0,
  });
  return {
    client: new InstallTargetsClient({
      http,
      log: (msg) => logs.push(msg),
      adminConsentUrlFor: () => 'https://consent.example/admin',
      ...(now !== undefined ? { now } : {}),
    }),
    logs,
  };
}

/**
 * A fake Graph that pages a collection: `pageOf(n)` is page `n` (0-based) and
 * says whether another follows. Member lookups go to `members`. Records every
 * URL it was asked for.
 */
function pagedGraph(opts: {
  collection: string;
  pageOf: (page: number) => { value: unknown[]; more: boolean };
  members?: (chatId: string) => ResponseSpec;
}): { fetchImpl: typeof fetch; urls: string[] } {
  const urls: string[] = [];
  const fetchImpl = (async (input: Parameters<typeof fetch>[0]) => {
    const url = String(input);
    if (url.includes('login.microsoftonline.com')) {
      return makeResponse({ status: 200, body: { access_token: 'app-only-tok', expires_in: 3600 } });
    }
    urls.push(url);
    const memberMatch = /\/chats\/([^/]+)\/members$/.exec(url);
    if (memberMatch?.[1] !== undefined) {
      return makeResponse(
        opts.members?.(decodeURIComponent(memberMatch[1])) ?? { status: 200, body: { value: [] } },
      );
    }
    const page = Number(/skiptoken=(\d+)/.exec(url)?.[1] ?? '0');
    const { value, more } = opts.pageOf(page);
    return makeResponse({
      status: 200,
      body: {
        value,
        ...(more
          ? { '@odata.nextLink': `https://graph.microsoft.com/v1.0/${opts.collection}?$skiptoken=${String(page + 1)}` }
          : {}),
      },
    });
  }) as typeof fetch;
  return { fetchImpl, urls };
}

const titledChat = (n: number): unknown => ({
  id: `19:chat-${String(n)}@thread.v2`,
  topic: `Chat ${String(n)}`,
  chatType: 'group',
});

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
            },
          ],
        },
      }),
    ]);

    const chats = await client.listChats({ tokens: tokens() });
    assert.deepEqual(chats, [
      { id: '19:group@thread.v2', topic: 'Project Falcon', chatType: 'group' },
    ]);
    assert.equal(chats.truncated, false);
    assert.equal(chats.withoutMemberNames, 0);
    // A chat with a topic needs no member names: one request, no lookup.
    assert.equal(calls.length, 1);

    const call = calls[0];
    assert.ok(call);
    // /me/chats, not /users/{id}/chats: the delegated route answers for the
    // administrator who signed in, which is exactly whose chats they can drop
    // an agent into. The application-permission route is per-user and would
    // mean enumerating the whole tenant.
    assert.ok(call.url.includes('/v1.0/me/chats'));
    // 0.9.1: the bulk listing no longer expands members for every chat.
    assert.doesNotMatch(call.url, /\$expand/);
    const headers = call.init?.headers as Record<string, string> | undefined;
    // The DELEGATED token, not the app-only one the token cache holds.
    assert.equal(headers?.['Authorization'], 'Bearer delegated-access-token');
  });

  it('keeps a 1:1 chat with topic null and identifies it by member names', async () => {
    const { client, calls } = harness([
      route('/me/chats', {
        status: 200,
        body: {
          value: [
            {
              id: '19:solo@unq.gbl.spaces',
              topic: null,
              chatType: 'oneOnOne',
            },
          ],
        },
      }),
      route('/members', {
        status: 200,
        body: { value: [{ displayName: 'Grace Hopper' }, { displayName: ' ' }] },
      }),
    ]);
    const chats = await client.listChats({ tokens: tokens() });
    const [chat] = chats;
    assert.equal(chats.withoutMemberNames, 0);
    const lookup = calls[1];
    assert.equal(
      lookup?.url,
      'https://graph.microsoft.com/v1.0/chats/19%3Asolo%40unq.gbl.spaces/members',
    );
    assert.equal(
      (lookup?.init?.headers as Record<string, string> | undefined)?.['Authorization'],
      'Bearer delegated-access-token',
    );
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

// byte5ai/omadia#2043, rollout defect D8. On a real tenant the chat listing
// stopped after 20 pages with 725 chats and returned that short list as if it
// were all of them, with `$expand=members` on every page.
describe('InstallTargetsClient.listChats: paging to the end, honestly (0.9.1)', () => {
  it('returns all 725 chats of 30 short pages, unexpanded and not truncated', async () => {
    // Graph often returns fewer than $top per page; 0.9.0 gave up at page 20.
    const perPage = 25;
    const total = 725;
    const { fetchImpl, urls } = pagedGraph({
      collection: 'me/chats',
      pageOf: (page) => ({
        value: Array.from(
          { length: Math.min(perPage, total - page * perPage) },
          (_, i) => titledChat(page * perPage + i),
        ),
        more: (page + 1) * perPage < total,
      }),
    });
    const { client, logs } = harnessWith(fetchImpl);

    const chats = await client.listChats({ tokens: tokens() });
    assert.equal(chats.length, 725);
    assert.equal(new Set(chats.map((c) => c.id)).size, 725);
    assert.equal(chats.truncated, false);
    assert.equal(urls.length, 29);
    assert.ok(urls.every((url) => !url.includes('$expand')), 'no page may expand members');
    assert.deepEqual(logs, []);
  });

  it('stops at the page bound and says so on the result', async () => {
    const { fetchImpl, urls } = pagedGraph({
      collection: 'me/chats',
      pageOf: (page) => ({ value: [titledChat(page)], more: true }),
    });
    const { client, logs } = harnessWith(fetchImpl);

    const chats = await client.listChats({ tokens: tokens() });
    assert.equal(urls.length, 200);
    assert.equal(chats.length, 200);
    assert.equal(chats.truncated, true);
    assert.equal(logs.length, 1);
    assert.match(logs[0] ?? '', /chats\.list: stopped by 200 pages .* the list is truncated/);
  });

  it('stops at the time budget and says so on the result', async () => {
    let clock = 0;
    const { fetchImpl, urls } = pagedGraph({
      collection: 'me/chats',
      pageOf: (page) => {
        clock += 25_000;
        return { value: [titledChat(page)], more: true };
      },
    });
    const { client, logs } = harnessWith(fetchImpl, () => clock);

    const chats = await client.listChats({ tokens: tokens() });
    // 25 s per page against a 60 s budget: pages at 0, 25 and 50 s, not at 75 s.
    assert.equal(urls.length, 3);
    assert.equal(chats.truncated, true);
    assert.match(logs[0] ?? '', /stopped by its time budget/);
  });

  it('stops when the next link points at a page it already read', async () => {
    const urls: string[] = [];
    const fetchImpl = (async (input: Parameters<typeof fetch>[0]) => {
      urls.push(String(input));
      return makeResponse({
        status: 200,
        body: { value: [titledChat(urls.length)], '@odata.nextLink': CHAT_NEXT_PAGE },
      });
    }) as typeof fetch;
    const { client } = harnessWith(fetchImpl);

    const chats = await client.listChats({ tokens: tokens() });
    assert.equal(urls.length, 2);
    assert.equal(chats.truncated, true);
  });

  it('the marks do not change what an older caller sees of the array', async () => {
    const { fetchImpl } = pagedGraph({
      collection: 'me/chats',
      pageOf: () => ({ value: [titledChat(1)], more: false }),
    });
    const chats = await harnessWith(fetchImpl).client.listChats({ tokens: tokens() });

    assert.ok(Array.isArray(chats));
    assert.deepEqual(Object.keys(chats), ['0']);
    assert.equal(
      JSON.stringify(chats),
      '[{"id":"19:chat-1@thread.v2","topic":"Chat 1","chatType":"group"}]',
    );
  });
});

describe('InstallTargetsClient.listChats: member names only where there is no topic (0.9.1)', () => {
  const untitled = (n: number, chatType: string): unknown => ({
    id: `19:${chatType}-${String(n)}@thread.v2`,
    topic: null,
    chatType,
  });
  const named = (chatId: string): ResponseSpec => ({
    status: 200,
    body: { value: [{ displayName: `Member of ${chatId}` }] },
  });

  it('looks up at most 100 untitled chats, group chats first, and counts the rest', async () => {
    const value = [
      ...Array.from({ length: 90 }, (_, i) => untitled(i, 'oneOnOne')),
      titledChat(1),
      ...Array.from({ length: 40 }, (_, i) => untitled(i, 'group')),
    ];
    const { fetchImpl, urls } = pagedGraph({
      collection: 'me/chats',
      pageOf: () => ({ value, more: false }),
      members: named,
    });
    const { client, logs } = harnessWith(fetchImpl);

    const chats = await client.listChats({ tokens: tokens() });
    const lookups = urls.filter((url) => url.endsWith('/members'));
    assert.equal(lookups.length, 100);
    assert.ok(!lookups.some((url) => url.includes('chat-1')), 'a titled chat is never looked up');

    const withNames = chats.filter((chat) => chat.memberNames !== undefined);
    assert.equal(withNames.filter((chat) => chat.chatType === 'group').length, 40);
    assert.equal(withNames.filter((chat) => chat.chatType === 'oneOnOne').length, 60);
    assert.deepEqual(chats.find((chat) => chat.id === '19:group-0@thread.v2')?.memberNames, [
      'Member of 19:group-0@thread.v2',
    ]);
    // Order of the listing is Graph's, whatever the lookup order was.
    assert.equal(chats[0]?.id, '19:oneOnOne-0@thread.v2');
    assert.equal(chats.length, 131);
    assert.equal(chats.withoutMemberNames, 30);
    assert.equal(chats.truncated, false);
    assert.match(logs.join('\n'), /30 of 131 chats have neither a topic nor member names/);
  });

  it('a chat whose members answer 404 loses its names, not the listing', async () => {
    const { fetchImpl } = pagedGraph({
      collection: 'me/chats',
      pageOf: () => ({ value: [untitled(0, 'group'), untitled(1, 'group')], more: false }),
      members: (chatId) => (chatId.includes('group-0') ? { status: 404, body: {} } : named(chatId)),
    });
    const chats = await harnessWith(fetchImpl).client.listChats({ tokens: tokens() });

    assert.equal(chats[0]?.memberNames, undefined);
    assert.deepEqual(chats[1]?.memberNames, ['Member of 19:group-1@thread.v2']);
    assert.equal(chats.withoutMemberNames, 1);
  });

  for (const status of [401, 403, 429, 500]) {
    it(`stops looking up after a ${String(status)} and still returns every chat`, async () => {
      const { fetchImpl, urls } = pagedGraph({
        collection: 'me/chats',
        pageOf: () => ({
          value: Array.from({ length: 12 }, (_, i) => untitled(i, 'group')),
          more: false,
        }),
        members: () => ({
          status,
          body: { error: { code: 'Refused', message: 'Bearer abcdefgh12345678 was refused' } },
        }),
      });
      const { client, logs } = harnessWith(fetchImpl);

      const chats = await client.listChats({ tokens: tokens() });
      assert.equal(chats.length, 12);
      assert.equal(chats.withoutMemberNames, 12);
      // One batch of five went out; the refusal would repeat, so no second one.
      assert.equal(urls.filter((url) => url.endsWith('/members')).length, 5);
      assert.match(logs.join('\n'), /stopped looking up members/);
      assert.ok(!logs.join('\n').includes('abcdefgh12345678'), 'a token reached a log line');
      assert.ok(!logs.join('\n').includes('delegated-access-token'));
    });
  }

  it('spends no lookups once the time budget is gone', async () => {
    let clock = 0;
    const { fetchImpl, urls } = pagedGraph({
      collection: 'me/chats',
      pageOf: () => {
        clock += 61_000;
        return { value: [untitled(0, 'group')], more: false };
      },
      members: named,
    });
    const chats = await harnessWith(fetchImpl, () => clock).client.listChats({ tokens: tokens() });

    assert.equal(urls.length, 1);
    assert.equal(chats.withoutMemberNames, 1);
    // The listing itself reached Graph's last page.
    assert.equal(chats.truncated, false);
  });
});

describe('InstallTargetsClient.listTeams: truncation is reported (0.9.1)', () => {
  it('a complete listing says so', async () => {
    const { fetchImpl } = pagedGraph({
      collection: 'teams',
      pageOf: (page) => ({ value: [{ id: `t-${String(page)}`, displayName: 'T' }], more: page < 2 }),
    });
    const teams = await harnessWith(fetchImpl).client.listTeams();
    assert.equal(teams.length, 3);
    assert.equal(teams.truncated, false);
  });

  it('a listing cut at the page bound says so', async () => {
    const { fetchImpl, urls } = pagedGraph({
      collection: 'teams',
      pageOf: (page) => ({ value: [{ id: `t-${String(page)}`, displayName: 'T' }], more: true }),
    });
    const { client, logs } = harnessWith(fetchImpl);
    const teams = await client.listTeams();
    assert.equal(urls.length, 20);
    assert.equal(teams.truncated, true);
    assert.match(logs[0] ?? '', /teams\.list: stopped by 20 pages/);
  });
});


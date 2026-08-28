import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  BOT_KIND_REGISTRATION,
  BOT_MSA_APP_TYPE_SINGLE_TENANT,
  BOT_SERVICE_API_VERSION,
  BOT_SERVICES_CHANNELS_WRITE_ACTION,
  BOT_SERVICES_DELETE_ACTION,
  BOT_SERVICES_WRITE_ACTION,
  BOT_SKU_F0,
  BotServiceClient,
  MS_TEAMS_CHANNEL_NAME,
} from '../src/teamsProvisioner/botService.js';
import type { ArmConfigResult } from '../src/teamsProvisioner/config.js';
import { ProvisioningHttp } from '../src/teamsProvisioner/http.js';
import {
  BotHandleUnavailableError,
  ConsentMissingError,
  ProvisioningThrottledError,
  isTransientProvisioningFailure,
} from '../src/teamsProvisioner/errors.js';
import type { CreateBotInput, RegistrationOnlyOutcome } from '../src/teamsProvisioner/types.js';

// createBot: two chained ARM PUTs (Microsoft.BotService/botServices upsert
// with the PINNED body literals kind 'registration' / sku 'F0' / msaAppType
// 'SingleTenant', then the MsTeamsChannel enablement) behind the shared
// ProvisioningHttp choke point; deleteBot as the idempotent rollback half,
// getBot as status probe, and the registration-only degradation that answers
// BEFORE any token acquisition. Mocked-fetch style follows
// tests/teamsProvisionerHttp.test.ts (scope-aware token route).

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

/**
 * Route table with per-route response queues (last entry repeats; FIRST match
 * wins, so the more specific `/channels/` route must precede the bot route).
 * The token route answers scope-aware so ARM-audience bearers are assertable.
 */
function mockFetch(routes: Route[], calls: FetchCall[]): typeof fetch {
  const queues = routes.map((r) => ({ ...r, next: 0 }));
  return (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, ...(init ? { init } : {}) });
    if (url.includes('login.microsoftonline.com')) {
      const params = new URLSearchParams(String(init?.body ?? ''));
      const scopeHost = new URL(params.get('scope') ?? 'https://none/').host;
      const access_token = `tok:${params.get('client_id') ?? '?'}:${scopeHost}`;
      return makeResponse({ status: 200, body: { access_token, expires_in: 3600 } });
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

const ARM_CONFIGURED: ArmConfigResult = {
  kind: 'configured',
  subscriptionId: 'sub-1',
  resourceGroup: 'rg-1',
  region: 'global',
  credential: { clientId: 'app-arm', clientSecret: 'arm-secret', source: 'dedicated-sp' },
};

const REGISTRATION_ONLY: RegistrationOnlyOutcome = {
  kind: 'registration-only',
  reason: 'arm-not-configured',
  missingSetupFields: ['azure_subscription_id', 'azure_region'],
};

const BOT_NAME = 'omadia-agent-hr';
const APP_ID = 'aaaa1111-0000-0000-0000-000000000001';
const TENANT_ID = 'tenant-1';
const ENDPOINT = 'https://mw.example.com/api/teams/omadia-agent-hr/messages';
const RESOURCE_ID = `/subscriptions/sub-1/resourceGroups/rg-1/providers/Microsoft.BotService/botServices/${BOT_NAME}`;
const BOT_URL = `https://management.azure.com${RESOURCE_ID}?api-version=${BOT_SERVICE_API_VERSION}`;
const CHANNEL_URL = `https://management.azure.com${RESOURCE_ID}/channels/${MS_TEAMS_CHANNEL_NAME}?api-version=${BOT_SERVICE_API_VERSION}`;

const CREATE_INPUT: CreateBotInput = {
  botName: BOT_NAME,
  displayName: 'HR Agent',
  msaAppId: APP_ID,
  msaAppTenantId: TENANT_ID,
  messagingEndpoint: ENDPOINT,
};

const BOT_BODY = {
  id: RESOURCE_ID,
  name: BOT_NAME,
  properties: {
    msaAppId: APP_ID,
    endpoint: ENDPOINT,
    provisioningState: 'Succeeded',
  },
};

const CHANNEL_OK = { status: 200, body: { name: MS_TEAMS_CHANNEL_NAME } };

function harness(
  routes: Route[],
  armConfig: ArmConfigResult = ARM_CONFIGURED,
): { client: BotServiceClient; calls: FetchCall[]; sleeps: number[] } {
  const calls: FetchCall[] = [];
  const sleeps: number[] = [];
  const http = new ProvisioningHttp({
    graphCredential: { tenantId: TENANT_ID, clientId: 'app-graph', clientSecret: 'gs' },
    armCredential: { tenantId: TENANT_ID, clientId: 'app-arm', clientSecret: 'arm-secret' },
    fetchImpl: mockFetch(routes, calls),
    log: () => {},
    sleep: async (ms) => {
      sleeps.push(ms);
    },
    max429Retries: 1,
  });
  const client = new BotServiceClient({ http, armConfig, log: () => {} });
  return { client, calls, sleeps };
}

const armCalls = (calls: FetchCall[]): FetchCall[] =>
  calls.filter((c) => c.url.includes('management.azure.com'));

const methodOf = (call: FetchCall): string | undefined => call.init?.method;

const jsonBodyOf = (call: FetchCall): Record<string, unknown> =>
  JSON.parse(String(call.init?.body)) as Record<string, unknown>;

const rejection = async (p: Promise<unknown>): Promise<unknown> =>
  p.then(() => assert.fail('expected the call to reject'), (e: unknown) => e);

describe('BotServiceClient.createBot', () => {
  it('PUTs the bot with the pinned body literals, then enables MsTeamsChannel', async () => {
    const { client, calls } = harness([
      route('/channels/MsTeamsChannel', CHANNEL_OK),
      route('/botServices/omadia-agent-hr', { status: 201, body: BOT_BODY }),
    ]);
    const outcome = await client.createBot(CREATE_INPUT);

    assert.equal(outcome.kind, 'provisioned');
    if (outcome.kind !== 'provisioned') return;
    assert.equal(outcome.bot.outcome, 'created');
    assert.deepEqual(outcome.bot.value, {
      botName: BOT_NAME,
      resourceId: RESOURCE_ID,
      msaAppId: APP_ID,
      messagingEndpoint: ENDPOINT,
    });

    const arm = armCalls(calls);
    assert.equal(arm.length, 2);
    const [botPut, channelPut] = arm as [FetchCall, FetchCall];

    // Bot PUT: exact ARM URL + method + the spec-pinned literals.
    assert.equal(botPut.url, BOT_URL);
    assert.equal(methodOf(botPut), 'PUT');
    const botBody = jsonBodyOf(botPut);
    assert.equal(botBody['location'], 'global');
    assert.equal(botBody['kind'], BOT_KIND_REGISTRATION);
    assert.equal(botBody['kind'], 'registration');
    assert.deepEqual(botBody['sku'], { name: BOT_SKU_F0 });
    assert.deepEqual(botBody['sku'], { name: 'F0' });
    assert.deepEqual(botBody['properties'], {
      displayName: 'HR Agent',
      endpoint: ENDPOINT,
      msaAppId: APP_ID,
      msaAppType: BOT_MSA_APP_TYPE_SINGLE_TENANT,
      msaAppTenantId: TENANT_ID,
    });
    assert.equal(
      (botBody['properties'] as Record<string, unknown>)['msaAppType'],
      'SingleTenant',
    );
    // Both PUTs run against the ARM audience with the dedicated SP identity.
    const bearer = (botPut.init?.headers as Record<string, string>)['Authorization'];
    assert.equal(bearer, 'Bearer tok:app-arm:management.azure.com');

    // Channel PUT: sub-resource URL + enablement body.
    assert.equal(channelPut.url, CHANNEL_URL);
    assert.equal(methodOf(channelPut), 'PUT');
    assert.deepEqual(jsonBodyOf(channelPut), {
      location: 'global',
      properties: {
        channelName: MS_TEAMS_CHANNEL_NAME,
        properties: { isEnabled: true },
      },
    });
  });

  it('maps an upsert 200 to the already-existed idempotent outcome', async () => {
    const { client } = harness([
      route('/channels/MsTeamsChannel', CHANNEL_OK),
      route('/botServices/omadia-agent-hr', { status: 200, body: BOT_BODY }),
    ]);
    const outcome = await client.createBot(CREATE_INPUT);
    assert.equal(outcome.kind, 'provisioned');
    if (outcome.kind !== 'provisioned') return;
    assert.equal(outcome.bot.outcome, 'already-existed');
  });

  it('polls the Azure-AsyncOperation of a long-running PUT to Succeeded', async () => {
    const { client, calls, sleeps } = harness([
      route(
        '/operations/op-1',
        { status: 200, body: { status: 'InProgress' } },
        { status: 200, body: { status: 'Succeeded' } },
      ),
      route('/channels/MsTeamsChannel', CHANNEL_OK),
      route('/botServices/omadia-agent-hr', {
        status: 201,
        body: BOT_BODY,
        headers: {
          'Azure-AsyncOperation': 'https://management.azure.com/operations/op-1',
          'Retry-After': '3',
        },
      }),
    ]);
    const outcome = await client.createBot(CREATE_INPUT);
    assert.equal(outcome.kind, 'provisioned');
    if (outcome.kind !== 'provisioned') return;
    assert.equal(outcome.bot.outcome, 'created');
    assert.equal(outcome.bot.value.resourceId, RESOURCE_ID);
    const polls = calls.filter((c) => c.url.includes('/operations/op-1'));
    assert.equal(polls.length, 2);
    assert.deepEqual(sleeps, [3000, 2000]);
  });

  it('throws ConsentMissingError(arm) with the write action on a bot-PUT 403', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', { status: 403, body: { error: 'RBAC' } }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.ok(err instanceof ConsentMissingError);
    assert.equal(err.resource, 'arm');
    assert.deepEqual(err.missingScopes, [BOT_SERVICES_WRITE_ACTION]);
  });

  it('honours Retry-After on 429 for BOTH ARM PUTs, then succeeds', async () => {
    const { client, sleeps } = harness([
      route(
        '/channels/MsTeamsChannel',
        { status: 429, headers: { 'Retry-After': '2' } },
        CHANNEL_OK,
      ),
      route(
        '/botServices/omadia-agent-hr',
        { status: 429, headers: { 'Retry-After': '5' } },
        { status: 201, body: BOT_BODY },
      ),
    ]);
    const outcome = await client.createBot(CREATE_INPUT);
    assert.equal(outcome.kind, 'provisioned');
    assert.deepEqual(sleeps, [5000, 2000]);
  });

  it('throws ProvisioningThrottledError(arm) when the 429 budget is exhausted', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', {
        status: 429,
        headers: { 'Retry-After': '7' },
      }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.ok(err instanceof ProvisioningThrottledError);
    assert.equal(err.resource, 'arm');
    assert.equal(err.retryAfterSeconds, 7);
  });

  it('rolls back a created bot when the channel enablement fails, rethrowing', async () => {
    const { client, calls } = harness([
      route('/channels/MsTeamsChannel', { status: 403, body: { error: 'RBAC' } }),
      route(
        '/botServices/omadia-agent-hr',
        { status: 201, body: BOT_BODY },
        { status: 200, body: {} },
      ),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.ok(err instanceof ConsentMissingError);
    assert.deepEqual(err.missingScopes, [BOT_SERVICES_CHANNELS_WRITE_ACTION]);
    const deletes = armCalls(calls).filter((c) => methodOf(c) === 'DELETE');
    assert.equal(deletes.length, 1);
    assert.equal(deletes[0]?.url, BOT_URL);
  });

  it('leaves a pre-existing bot untouched when the channel enablement fails', async () => {
    const { client, calls } = harness([
      route('/channels/MsTeamsChannel', { status: 500, body: { error: 'boom' } }),
      route('/botServices/omadia-agent-hr', { status: 200, body: BOT_BODY }),
    ]);
    await rejection(client.createBot(CREATE_INPUT));
    assert.equal(armCalls(calls).filter((c) => methodOf(c) === 'DELETE').length, 0);
  });

  it('treats a channel-PUT 409 as already enabled', async () => {
    const { client, calls } = harness([
      route('/channels/MsTeamsChannel', { status: 409, body: {} }),
      route('/botServices/omadia-agent-hr', { status: 201, body: BOT_BODY }),
    ]);
    const outcome = await client.createBot(CREATE_INPUT);
    assert.equal(outcome.kind, 'provisioned');
    assert.equal(armCalls(calls).filter((c) => methodOf(c) === 'DELETE').length, 0);
  });

  it('rejects a bot-PUT 409 — a foreign handle, not the idempotent signal', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', { status: 409, body: {} }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.ok(err instanceof BotHandleUnavailableError);
    assert.equal(err.status, 409);
    assert.equal(err.botName, BOT_NAME);
  });

  // ---------------------------------------------------------------------
  // byte5ai/omadia#921 — the handle namespace is GLOBAL
  // ---------------------------------------------------------------------

  it('promotes a 400 InvalidBotData to the typed BotHandleUnavailableError', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', {
        status: 400,
        body: {
          error: {
            code: 'InvalidBotData',
            message:
              'Bot is not valid. Errors: The bot name is already registered to another bot application.',
          },
        },
      }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.ok(err instanceof BotHandleUnavailableError);
    assert.equal(err.name, 'BotHandleUnavailableError');
    assert.equal(err.status, 400);
    assert.equal(err.botName, BOT_NAME);
  });

  it('explains the global namespace and the automatic qualification', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', {
        status: 400,
        body: {
          error: {
            code: 'InvalidBotData',
            message:
              'Bot is not valid. Errors: The bot name is already registered to another bot application.',
          },
        },
      }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    const message = String((err as Error).message);
    assert.match(message, /^bot_handle_unavailable: /);
    assert.match(message, /global namespace across all Azure customers/);
    assert.match(message, /not scoped to your tenant, subscription or resource group/);
    assert.match(message, /qualifies the handle automatically/);
  });

  it('never classifies a taken handle as transient — no retry storm', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', {
        status: 400,
        body: {
          error: {
            code: 'InvalidBotData',
            message:
              'Bot is not valid. Errors: The bot name is already registered to another bot application.',
          },
        },
      }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.equal(isTransientProvisioningFailure(err), false);
  });

  it('issues exactly ONE bot PUT for a taken handle and never rolls back', async () => {
    const { client, calls, sleeps } = harness([
      route('/botServices/omadia-agent-hr', {
        status: 400,
        body: {
          error: {
            code: 'InvalidBotData',
            message:
              'Bot is not valid. Errors: The bot name is already registered to another bot application.',
          },
        },
      }),
    ]);
    await rejection(client.createBot(CREATE_INPUT));
    const arm = armCalls(calls);
    assert.equal(arm.filter((c) => methodOf(c) === 'PUT').length, 1);
    // Nothing was created, so nothing may be deleted.
    assert.equal(arm.filter((c) => methodOf(c) === 'DELETE').length, 0);
    assert.deepEqual(sleeps, []);
  });

  it('leaves an ordinary 400 as the untyped request error', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', {
        status: 400,
        body: { error: { code: 'InvalidRequest', message: 'malformed body' } },
      }),
    ]);
    const err = await rejection(client.createBot(CREATE_INPUT));
    assert.equal(err instanceof BotHandleUnavailableError, false);
    assert.match(String(err), /botServices\.put 400/);
  });

  it('rejects handles that ARM would accept but Bot Framework would not', async () => {
    // The ARM resource-name rule allows dots, underscores and 64 chars; the
    // Bot Framework handle rule does not. The stricter rule wins (#921).
    const tooLong = `omadia-${'a'.repeat(40)}`;
    for (const bad of ['abc', tooLong, 'omadia_hr_bot', 'omadia.hr.bot', '-omadia-hr', 'omadia-hr-']) {
      const { client, calls } = harness([]);
      const err = await rejection(client.createBot({ ...CREATE_INPUT, botName: bad }));
      assert.match(String(err), /invalid_argument: 'botName'/, `expected ${bad} to be rejected`);
      assert.equal(calls.length, 0, `expected no network call for ${bad}`);
    }
  });

  it('accepts a qualified handle at exactly the 42-char boundary', async () => {
    const boundary = `omadia-${'a'.repeat(26)}-7034c271`;
    assert.equal(boundary.length, 42);
    const resourceId = `/subscriptions/sub-1/resourceGroups/rg-1/providers/Microsoft.BotService/botServices/${boundary}`;
    const { client } = harness([
      route('/channels/MsTeamsChannel', CHANNEL_OK),
      route(`/botServices/${boundary}`, {
        status: 201,
        body: { ...BOT_BODY, id: resourceId, name: boundary },
      }),
    ]);
    const outcome = await client.createBot({ ...CREATE_INPUT, botName: boundary });
    assert.equal(outcome.kind, 'provisioned');
  });

  it('rejects an ARM-URL-breaking botName before any network call', async () => {
    const { client, calls } = harness([]);
    const err = await rejection(
      client.createBot({ ...CREATE_INPUT, botName: 'bad name/../oops' }),
    );
    assert.match(String(err), /invalid_argument: 'botName'/);
    assert.equal(calls.length, 0);
  });
});

describe('BotServiceClient registration-only degradation', () => {
  it('createBot answers the typed outcome without any token acquisition', async () => {
    const { client, calls } = harness([], REGISTRATION_ONLY);
    assert.equal(client.canCreateBots, false);
    const outcome = await client.createBot(CREATE_INPUT);
    assert.deepEqual(outcome, REGISTRATION_ONLY);
    assert.equal(calls.length, 0);
  });

  it('deleteBot and getBot degrade the same way', async () => {
    const { client, calls } = harness([], REGISTRATION_ONLY);
    assert.deepEqual(await client.deleteBot(BOT_NAME), REGISTRATION_ONLY);
    assert.deepEqual(await client.getBot(BOT_NAME), REGISTRATION_ONLY);
    assert.equal(calls.length, 0);
  });

  it('canCreateBots is true in configured mode', () => {
    const { client } = harness([]);
    assert.equal(client.canCreateBots, true);
  });
});

describe('BotServiceClient.deleteBot', () => {
  it('DELETEs the bot resource and reports deleted', async () => {
    const { client, calls } = harness([
      route('/botServices/omadia-agent-hr', { status: 200, body: {} }),
    ]);
    assert.deepEqual(await client.deleteBot(BOT_NAME), {
      kind: 'deleted',
      outcome: 'deleted',
    });
    const del = armCalls(calls)[0];
    assert.equal(del?.url, BOT_URL);
    assert.equal(methodOf(del as FetchCall), 'DELETE');
  });

  it('is idempotent: a 404 answers already-deleted, never an error', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', { status: 404, body: {} }),
    ]);
    assert.deepEqual(await client.deleteBot(BOT_NAME), {
      kind: 'deleted',
      outcome: 'already-deleted',
    });
  });

  it('throws ConsentMissingError(arm) with the delete action on 403', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', { status: 403, body: {} }),
    ]);
    const err = await rejection(client.deleteBot(BOT_NAME));
    assert.ok(err instanceof ConsentMissingError);
    assert.equal(err.resource, 'arm');
    assert.deepEqual(err.missingScopes, [BOT_SERVICES_DELETE_ACTION]);
  });
});

describe('BotServiceClient.getBot', () => {
  it('parses a found bot resource', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', { status: 200, body: BOT_BODY }),
    ]);
    assert.deepEqual(await client.getBot(BOT_NAME), {
      kind: 'found',
      bot: {
        botName: BOT_NAME,
        resourceId: RESOURCE_ID,
        msaAppId: APP_ID,
        messagingEndpoint: ENDPOINT,
      },
    });
  });

  it('answers not-found on 404', async () => {
    const { client } = harness([
      route('/botServices/omadia-agent-hr', { status: 404, body: {} }),
    ]);
    assert.deepEqual(await client.getBot(BOT_NAME), { kind: 'not-found' });
  });
});

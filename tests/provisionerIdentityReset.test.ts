/**
 * provisionerIdentityReset.test.ts — an identity that was deleted is over
 * (0.9.2, byte5ai/omadia#2043 rollout defect D10).
 *
 * What happened on a real tenant: a reset soft-deleted a facilitator's app
 * registration and asked for the purge milliseconds later. The tombstone had
 * not replicated into the recycle bin, the connector answered
 * `already-absent`, the next run RESTORED the application (same client id)
 * and re-created the bot resource under that app id. The Bot Connector has
 * answered 401 for that bot ever since, while Azure reports it healthy.
 *
 * These tests run the capability against ONE stateful fake of Entra, Azure
 * Resource Manager and the Bot Connector, in which a delete reaches the
 * recycle bin only after a number of looks — the replication window the old
 * code did not survive.
 */

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { SecretsAccessor } from '@omadia/plugin-api';

import {
  BotAppIdRetiredError,
  RecycleBinUnreadableError,
  UniqueNameReservedError,
} from '../src/teamsProvisioner/errors.js';
import { PROBE_MEMBER_ID } from '../src/teamsProvisioner/botProbe.js';
import { TEAMS_RETIRED_BOT_APP_PREFIX } from '../src/teamsProvisioner/retiredBotApps.js';
import {
  createTeamsProvisioner,
  type TeamsProvisionerAccessor,
} from '../src/teamsProvisioner/index.js';

const TENANT_ID = '99999999-8888-7777-6666-555555555555';
const UNIQUE_NAME = 'omadia-teams-bot-mara-test';
const BOT_NAME = 'omadia-mara-test-e30baf34';
const ENDPOINT = 'https://omadia.example/api/teams/mara-test/messages';

interface FakeApp {
  id: string;
  appId: string;
  displayName: string;
  uniqueName?: string;
  /** Where the application is. `replicating` = deleted, in neither listing. */
  state: 'live' | 'replicating' | 'deleted';
  /** Recycle-bin looks left until a `replicating` application shows up there. */
  looksUntilVisible: number;
}

interface Call {
  method: string;
  url: string;
  body: string;
}

interface FakeTenantOptions {
  /** Recycle-bin looks a delete needs before its tombstone is listed. */
  replicationLooks?: number;
  /** `false` → listing the recycle bin answers 403. */
  binReadable?: boolean;
  /** What the Bot Connector answers a bot, by app id. Default: 403 BadArgument. */
  connector?: (appId: string) => { status: number; body: unknown };
  /** What the token endpoint answers a bot's own sign-in. Default: a token. */
  botToken?: (appId: string, secret: string) => { status: number; body: unknown };
}

const json = (status: number, body?: unknown): Response =>
  ({
    ok: status >= 200 && status < 300,
    status,
    headers: { get: () => null },
    json: async () => {
      if (body === undefined) throw new Error('no body');
      return body;
    },
    text: async () => (body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body)),
  }) as unknown as Response;

/** Entra + ARM + Bot Connector, with the state a provisioning run changes. */
function fakeTenant(opts: FakeTenantOptions = {}) {
  const apps: FakeApp[] = [];
  const bots = new Map<string, { msaAppId: string }>();
  const calls: Call[] = [];
  let created = 0;

  const appView = (app: FakeApp): Record<string, unknown> => ({
    id: app.id,
    appId: app.appId,
    displayName: app.displayName,
    signInAudience: 'AzureADMyOrg',
    ...(app.uniqueName !== undefined ? { uniqueName: app.uniqueName } : {}),
  });
  const notFound = (): Response => json(404, { error: { code: 'Request_ResourceNotFound' } });

  const graph = (method: string, path: string, body: string): Response => {
    if (path.startsWith('/directory/deletedItems/microsoft.graph.application')) {
      if (opts.binReadable === false) return json(403, { error: { code: 'Authorization_RequestDenied' } });
      for (const app of apps) {
        if (app.state !== 'replicating') continue;
        app.looksUntilVisible -= 1;
        if (app.looksUntilVisible <= 0) app.state = 'deleted';
      }
      return json(200, {
        value: apps
          .filter((app) => app.state === 'deleted')
          .map((app) => ({ ...appView(app), deletedDateTime: '2026-11-18T17:24:49Z' })),
      });
    }
    const item = /^\/directory\/deletedItems\/([^/?]+)(\/restore)?$/.exec(path);
    if (item !== null) {
      const index = apps.findIndex((app) => app.id === item[1] && app.state === 'deleted');
      if (index < 0) return notFound();
      if (item[2] !== undefined) {
        (apps[index] as FakeApp).state = 'live';
        return json(200, appView(apps[index] as FakeApp));
      }
      apps.splice(index, 1);
      return json(204);
    }
    if (path === '/applications' && method === 'POST') {
      const wanted = JSON.parse(body) as { displayName: string; uniqueName?: string };
      // A soft-deleted (or not yet replicated) application still holds its name.
      if (wanted.uniqueName !== undefined && apps.some((app) => app.uniqueName === wanted.uniqueName)) {
        return json(400, {
          error: {
            code: 'Request_BadRequest',
            message: 'Another object with the same value for property uniqueName already exists.',
          },
        });
      }
      created += 1;
      const app: FakeApp = {
        id: `object-${String(created)}`,
        appId: `app-${String(created)}`,
        displayName: wanted.displayName,
        ...(wanted.uniqueName !== undefined ? { uniqueName: wanted.uniqueName } : {}),
        state: 'live',
        looksUntilVisible: 0,
      };
      apps.push(app);
      return json(201, appView(app));
    }
    const byKey = /^\/applications\((appId|uniqueName)='([^']+)'\)/.exec(path);
    const byId = /^\/applications\/([^/?]+)(\/addPassword|\/removePassword)?/.exec(path);
    const app = apps.find(
      (candidate) =>
        candidate.state === 'live' &&
        (byKey !== null
          ? (byKey[1] === 'appId' ? candidate.appId : candidate.uniqueName) === decodeURIComponent(byKey[2] ?? '')
          : candidate.id === byId?.[1]),
    );
    if (byKey === null && byId === null) {
      if (path === '/servicePrincipals') return json(201, {});
      return json(500, { error: `fake Graph: unrouted ${method} ${path}` });
    }
    if (app === undefined) return notFound();
    if (method === 'DELETE') {
      app.state = 'replicating';
      app.looksUntilVisible = opts.replicationLooks ?? 0;
      if (app.looksUntilVisible <= 0) app.state = 'deleted';
      return json(204);
    }
    if (byId?.[2] === '/addPassword') {
      return json(200, { secretText: `secret-of-${app.appId}`, keyId: 'key-1', endDateTime: '2027-01-01T00:00:00Z' });
    }
    return json(200, { ...appView(app), passwordCredentials: [] });
  };

  const arm = (method: string, path: string, body: string): Response => {
    const name = /\/botServices\/([^/?]+)/.exec(path)?.[1] ?? '';
    if (path.includes('/channels/')) return json(200, { name: 'MsTeamsChannel' });
    const bot = bots.get(name);
    const view = (msaAppId: string): Record<string, unknown> => ({
      id: `/subscriptions/s/resourceGroups/rg/providers/Microsoft.BotService/botServices/${name}`,
      name,
      properties: { msaAppId, endpoint: ENDPOINT, provisioningState: 'Succeeded' },
    });
    if (method === 'PUT') {
      const msaAppId = (JSON.parse(body) as { properties: { msaAppId: string } }).properties.msaAppId;
      bots.set(name, { msaAppId });
      return json(bot === undefined ? 201 : 200, view(msaAppId));
    }
    if (bot === undefined) return json(404, {});
    if (method === 'DELETE') {
      bots.delete(name);
      return json(200, {});
    }
    return json(200, view(bot.msaAppId));
  };

  const fetchImpl = (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
    const url = String(input);
    const method = init?.method ?? 'GET';
    const body = typeof init?.body === 'string' ? init.body : '';
    calls.push({ method, url, body });
    const parsed = new URL(url);
    if (parsed.host === 'login.microsoftonline.com') {
      const form = new URLSearchParams(body);
      if (form.get('scope') !== 'https://api.botframework.com/.default') {
        return json(200, { access_token: 'connector-app-token', expires_in: 3600 });
      }
      const answer = opts.botToken?.(form.get('client_id') ?? '', form.get('client_secret') ?? '');
      return json(answer?.status ?? 200, answer?.body ?? { access_token: `bot-token-of-${form.get('client_id') ?? ''}`, expires_in: 3600 });
    }
    if (parsed.host === 'smba.trafficmanager.net') {
      const appId = (JSON.parse(body) as { bot: { id: string } }).bot.id.replace(/^28:/, '');
      const answer = opts.connector?.(appId) ?? {
        status: 403,
        body: { error: { code: 'BadArgument', message: "Bot is not installed in user's personal scope" } },
      };
      return json(answer.status, answer.body);
    }
    const path = decodeURIComponent(parsed.pathname.replace(/^\/v1\.0/, ''));
    return parsed.host === 'management.azure.com' ? arm(method, path, body) : graph(method, path, body);
  }) as typeof fetch;

  return { fetchImpl, apps, bots, calls };
}

function rig(opts: FakeTenantOptions = {}) {
  const tenant = fakeTenant(opts);
  const vault = new Map<string, string>();
  const logs: string[] = [];
  const sleeps: number[] = [];
  const secrets = {
    get: async (key: string) => vault.get(key),
    set: async (key: string, value: string) => {
      vault.set(key, value);
    },
    delete: async (key: string) => {
      vault.delete(key);
    },
  } as unknown as SecretsAccessor;
  const build = (): TeamsProvisionerAccessor =>
    createTeamsProvisioner({
      graphCredential: { tenantId: TENANT_ID, clientId: 'connector-app', clientSecret: 'connector-secret' },
      armConfig: {
        kind: 'configured',
        subscriptionId: 's',
        resourceGroup: 'rg',
        region: 'global',
        credential: { clientId: 'app-arm', clientSecret: 'arm-secret', source: 'dedicated-sp' },
      },
      secrets,
      log: (msg) => logs.push(msg),
      fetchImpl: tenant.fetchImpl,
      sleep: async (ms) => {
        sleeps.push(ms);
      },
    });
  return { ...tenant, vault, logs, sleeps, provisioner: build(), restart: build };
}

const register = (provisioner: TeamsProvisionerAccessor) =>
  provisioner.createAppRegistration({
    displayName: 'Mara (Test)',
    tenantMode: 'customer',
    uniqueName: UNIQUE_NAME,
  });

const createBot = (provisioner: TeamsProvisionerAccessor, msaAppId: string) =>
  provisioner.createBot({
    botName: BOT_NAME,
    displayName: 'Mara (Test)',
    msaAppId,
    msaAppTenantId: TENANT_ID,
    messagingEndpoint: ENDPOINT,
  });

const rejection = async (p: Promise<unknown>): Promise<unknown> =>
  p.then(() => assert.fail('expected a rejection'), (e: unknown) => e);

describe('purge right after a delete (D10, cause 1)', () => {
  it('waits for the tombstone, purges it, and never says already-absent', async () => {
    const { provisioner, apps, sleeps } = rig({ replicationLooks: 3 });
    const first = await register(provisioner);
    const { objectId } = first.value.registration;

    await provisioner.deleteAppRegistration({ appId: first.value.appId });
    // Milliseconds later, as the kernel's reset does.
    const purge = await provisioner.purgeDeletedAppRegistration({ objectId });

    assert.deepEqual(purge, { outcome: 'purged' });
    assert.deepEqual(apps, [], 'the tombstone is still in the tenant');
    assert.ok(sleeps.length >= 2, 'the purge did not wait for replication');
  });

  it('reports unconfirmed when the tombstone never appears within the bound', async () => {
    const { provisioner, apps, logs } = rig({ replicationLooks: 1000 });
    const first = await register(provisioner);

    await provisioner.deleteAppRegistration({ appId: first.value.appId });
    const purge = await provisioner.purgeDeletedAppRegistration({
      objectId: first.value.registration.objectId,
    });

    assert.equal(purge.outcome, 'unconfirmed');
    assert.match(purge.detail ?? '', /deleted moments ago .* did not appear in the recycle bin within 6 looks/);
    assert.equal(apps.length, 1, 'nothing may be reported gone that is still there');
    assert.match(logs.join('\n'), /purge of 'object-1' unconfirmed/);
  });

  it('a later call finishes the purge once the tombstone has replicated', async () => {
    const { provisioner, apps } = rig({ replicationLooks: 8 });
    const first = await register(provisioner);
    const { objectId } = first.value.registration;
    await provisioner.deleteAppRegistration({ appId: first.value.appId });

    assert.equal((await provisioner.purgeDeletedAppRegistration({ objectId })).outcome, 'unconfirmed');
    assert.equal((await provisioner.purgeDeletedAppRegistration({ objectId })).outcome, 'purged');
    assert.deepEqual(apps, []);
  });

  it('still answers already-absent for an entry the readable bin does not hold and nobody just deleted', async () => {
    const { provisioner } = rig();
    assert.deepEqual(
      await provisioner.purgeDeletedAppRegistration({ objectId: 'object-of-last-month' }),
      { outcome: 'already-absent' },
    );
  });

  it('answers unconfirmed when the recycle bin cannot be read', async () => {
    const { provisioner } = rig({ binReadable: false });
    const purge = await provisioner.purgeDeletedAppRegistration({ objectId: 'object-x' });
    assert.equal(purge.outcome, 'unconfirmed');
  });
});

describe('findDeletedAppRegistration (0.9.2, what the kernel reset looks with)', () => {
  it('finds the tombstone by app id once it has replicated', async () => {
    const { provisioner } = rig({ replicationLooks: 2 });
    const first = await register(provisioner);
    await provisioner.deleteAppRegistration({ appId: first.value.appId });

    assert.deepEqual(await provisioner.findDeletedAppRegistration({ appId: first.value.appId }), { found: false });
    assert.deepEqual(await provisioner.findDeletedAppRegistration({ appId: first.value.appId }), {
      found: true,
      objectId: first.value.registration.objectId,
      deletedDateTime: '2026-11-18T17:24:49Z',
    });
  });

  it('throws instead of answering "not found" for a bin it could not read', async () => {
    const { provisioner } = rig({ binReadable: false });
    const err = await rejection(provisioner.findDeletedAppRegistration({ appId: 'app-1' }));
    assert.ok(err instanceof RecycleBinUnreadableError);
    assert.match(err.message, /recycle_bin_unreadable/);
  });
});

describe('registering under a name a deleted application holds (D10, cause 2)', () => {
  it('purges the tombstone and creates a new application with a new client id', async () => {
    const { provisioner, apps, calls, logs } = rig();
    const first = await register(provisioner);
    await provisioner.deleteAppRegistration({ appId: first.value.appId });

    const second = await register(provisioner);

    assert.equal(second.outcome, 'created');
    assert.notEqual(second.value.appId, first.value.appId, 'the old client id was reused');
    assert.notEqual(second.value.registration.objectId, first.value.registration.objectId);
    assert.deepEqual(apps.map((app) => `${app.appId}:${app.state}`), [`${second.value.appId}:live`]);
    assert.ok(!calls.some((call) => call.url.includes('/restore')), 'the tombstone was restored');
    assert.match(logs.join('\n'), /held by soft-deleted app 'object-1' .* purging it/);
  });

  it('does the same when the delete has not replicated yet: waits, purges, creates anew', async () => {
    // The rollout sequence: reset deletes, the purge is not confirmed, and the
    // next run starts while the application is in neither listing.
    const { provisioner, apps, calls } = rig({ replicationLooks: 4 });
    const first = await register(provisioner);
    await provisioner.deleteAppRegistration({ appId: first.value.appId });

    const second = await register(provisioner);

    assert.equal(second.outcome, 'created');
    assert.notEqual(second.value.appId, first.value.appId);
    assert.equal(apps.length, 1);
    assert.ok(!calls.some((call) => call.url.includes('/restore')));
  });

  it('a live application under the name is still adopted (the idempotent re-run)', async () => {
    const { provisioner } = rig();
    const first = await register(provisioner);
    const again = await register(provisioner);

    assert.equal(again.outcome, 'already-existed');
    assert.equal(again.value.appId, first.value.appId);
  });

  it('says why when the name is taken and the recycle bin cannot be read', async () => {
    const { provisioner, calls } = rig({ binReadable: false });
    const first = await register(provisioner);
    await provisioner.deleteAppRegistration({ appId: first.value.appId });

    const err = await rejection(register(provisioner));
    assert.ok(err instanceof UniqueNameReservedError);
    assert.match(err.message, /Do not restore it/);
    assert.ok(!calls.some((call) => call.url.includes('/restore')));
  });
});

describe('no second bot resource under an app id that had one (D10, cause 3)', () => {
  it('refuses to create a bot under the app id of a deleted bot, before any PUT', async () => {
    const { provisioner, bots, calls, vault } = rig();
    const app = await register(provisioner);
    await createBot(provisioner, app.value.appId);
    await provisioner.deleteBot(BOT_NAME);
    const putsBefore = calls.filter((call) => call.method === 'PUT').length;

    const err = await rejection(createBot(provisioner, app.value.appId));

    assert.ok(err instanceof BotAppIdRetiredError);
    assert.equal(err.msaAppId, app.value.appId);
    assert.match(err.message, /^bot_app_id_retired: /);
    assert.match(err.message, /provision it again: it gets a new app id/);
    assert.equal(bots.size, 0);
    assert.equal(calls.filter((call) => call.method === 'PUT').length, putsBefore);
    assert.ok(vault.has(`${TEAMS_RETIRED_BOT_APP_PREFIX}:${app.value.appId}`));
  });

  it('remembers it across a restart of the connector', async () => {
    const { provisioner, restart } = rig();
    const app = await register(provisioner);
    await createBot(provisioner, app.value.appId);
    await provisioner.deleteBot(BOT_NAME);

    const err = await rejection(createBot(restart(), app.value.appId));
    assert.ok(err instanceof BotAppIdRetiredError);
    assert.match(err.retiredAt ?? '', /^\d{4}-\d{2}-\d{2}T/);
  });

  it('a new app id gets its bot under the same handle', async () => {
    const { provisioner, bots } = rig();
    const first = await register(provisioner);
    await createBot(provisioner, first.value.appId);
    await provisioner.deleteBot(BOT_NAME);
    await provisioner.deleteAppRegistration({ appId: first.value.appId });
    const second = await register(provisioner);

    const outcome = await createBot(provisioner, second.value.appId);
    assert.equal(outcome.kind, 'provisioned');
    assert.equal(bots.get(BOT_NAME)?.msaAppId, second.value.appId);
  });

  it('re-running createBot for a bot that was never deleted still upserts', async () => {
    const { provisioner } = rig();
    const app = await register(provisioner);
    await createBot(provisioner, app.value.appId);
    const again = await createBot(provisioner, app.value.appId);
    assert.equal(again.kind, 'provisioned');
    assert.equal(again.kind === 'provisioned' ? again.bot.outcome : '', 'already-existed');
  });

  it('deleting a bot that does not exist retires nothing', async () => {
    const { provisioner, vault } = rig();
    assert.deepEqual(await provisioner.deleteBot(BOT_NAME), { kind: 'deleted', outcome: 'already-deleted' });
    assert.equal([...vault.keys()].filter((key) => key.startsWith(TEAMS_RETIRED_BOT_APP_PREFIX)).length, 0);
  });
});

describe('the whole rollout sequence: reset, then provision again', () => {
  it('ends with a new client id and a bot the old one never shared', async () => {
    const { provisioner, bots, calls } = rig({ replicationLooks: 3 });
    const mara = await register(provisioner);
    await createBot(provisioner, mara.value.appId);

    // "Lauf zurücksetzen": bot, then application, then the purge at once.
    await provisioner.deleteBot(BOT_NAME);
    await provisioner.deleteAppRegistration({ appId: mara.value.appId });
    const purge = await provisioner.purgeDeletedAppRegistration({
      objectId: mara.value.registration.objectId,
    });
    assert.equal(purge.outcome, 'purged');

    // "Provisioning starten".
    const again = await register(provisioner);
    assert.equal(again.outcome, 'created');
    assert.notEqual(again.value.appId, mara.value.appId);
    await createBot(provisioner, again.value.appId);
    assert.equal(bots.get(BOT_NAME)?.msaAppId, again.value.appId);
    assert.ok(!calls.some((call) => call.url.includes('/restore')));
    // And the old app id can never get a bot again.
    assert.ok((await rejection(createBot(provisioner, mara.value.appId))) instanceof BotAppIdRetiredError);
  });
});

describe('probeBot: does the Bot Connector know the bot (0.9.2)', () => {
  const DEAD = { status: 401, body: { message: 'Authorization has been denied for this request.' } };

  it('a healthy bot is known: 403 BadArgument for a member it is not installed for', async () => {
    const { provisioner, calls } = rig();
    const app = await register(provisioner);

    const result = await provisioner.probeBot({ appId: app.value.appId });

    assert.equal(result.status, 'known');
    assert.match(result.detail ?? '', /connector 403: .*not installed in user's personal scope/);

    const token = calls.find((call) => call.body.includes('api.botframework.com'));
    assert.equal(token?.url, `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`);
    const form = new URLSearchParams(token?.body ?? '');
    assert.equal(form.get('grant_type'), 'client_credentials');
    assert.equal(form.get('client_id'), app.value.appId);
    assert.equal(form.get('client_secret'), `secret-of-${app.value.appId}`);
    assert.equal(form.get('scope'), 'https://api.botframework.com/.default');

    const probe = calls.find((call) => call.url.includes('smba.trafficmanager.net'));
    assert.equal(probe?.method, 'POST');
    assert.equal(probe?.url, 'https://smba.trafficmanager.net/teams/v3/conversations');
    assert.deepEqual(JSON.parse(probe?.body ?? '{}'), {
      bot: { id: `28:${app.value.appId}` },
      members: [{ id: PROBE_MEMBER_ID }],
      channelData: { tenant: { id: TENANT_ID } },
      isGroup: false,
    });
  });

  it('a dead bot is rejected: 401 from the Bot Connector', async () => {
    const { provisioner } = rig({ connector: () => DEAD });
    const app = await register(provisioner);

    assert.deepEqual(await provisioner.probeBot({ appId: app.value.appId }), {
      status: 'rejected',
      httpStatus: 401,
      detail: 'connector 401: {"message":"Authorization has been denied for this request."}',
    });
  });

  for (const status of [200, 201, 400, 404, 405]) {
    it(`a ${String(status)} from the Bot Connector means the bot is known`, async () => {
      const { provisioner } = rig({ connector: () => ({ status, body: { id: 'a:1' } }) });
      const app = await register(provisioner);
      assert.equal((await provisioner.probeBot({ appId: app.value.appId })).status, 'known');
    });
  }

  for (const status of [429, 500, 503]) {
    it(`a ${String(status)} from the Bot Connector verifies nothing`, async () => {
      const { provisioner } = rig({ connector: () => ({ status, body: 'later' }) });
      const app = await register(provisioner);
      assert.equal((await provisioner.probeBot({ appId: app.value.appId })).status, 'unverified');
    });
  }

  it('invalid_client at the token endpoint is a rejection, with the AADSTS code', async () => {
    const { provisioner, calls } = rig({
      botToken: () => ({
        status: 401,
        body: {
          error: 'invalid_client',
          error_description: "AADSTS7000215: Invalid client secret provided. Trace ID: 5f0c1d2e-0000-4000-8000-000000000001",
        },
      }),
    });
    const app = await register(provisioner);

    const result = await provisioner.probeBot({ appId: app.value.appId });
    assert.equal(result.status, 'rejected');
    assert.equal(result.status === 'rejected' ? result.httpStatus : 0, 401);
    assert.match(result.detail ?? '', /AADSTS7000215/);
    assert.ok(!calls.some((call) => call.url.includes('smba.trafficmanager.net')));
  });

  it('a token endpoint that is down verifies nothing', async () => {
    const { provisioner } = rig({ botToken: () => ({ status: 503, body: 'unavailable' }) });
    const app = await register(provisioner);
    assert.equal((await provisioner.probeBot({ appId: app.value.appId })).status, 'unverified');
  });

  it('no stored secret verifies nothing and asks nobody', async () => {
    const { provisioner, calls } = rig();
    const result = await provisioner.probeBot({ appId: 'app-without-secret' });
    assert.deepEqual(result, { status: 'unverified', detail: "no stored secret for app id 'app-without-secret'" });
    assert.equal(calls.length, 0);
  });

  it('a network failure verifies nothing and does not throw', async () => {
    const { provisioner, vault } = rig({
      connector: () => {
        throw new TypeError('fetch failed');
      },
    });
    const app = await register(provisioner);
    assert.ok(vault.size > 0);
    const result = await provisioner.probeBot({ appId: app.value.appId });
    assert.equal(result.status, 'unverified');
    assert.match(result.detail ?? '', /fetch failed/);
  });

  it('neither the secret nor the token reaches a result or a log line', async () => {
    const { provisioner, logs } = rig({
      connector: (appId) => ({
        status: 401,
        body: `denied for Bearer bot-token-of-${appId} with client_secret=secret-of-${appId}`,
      }),
    });
    const app = await register(provisioner);

    const result = await provisioner.probeBot({ appId: app.value.appId });
    const everything = `${JSON.stringify(result)} ${logs.join('\n')}`;
    assert.equal(result.status, 'rejected');
    assert.ok(!everything.includes(`secret-of-${app.value.appId}`), 'the secret leaked');
    assert.ok(!everything.includes(`bot-token-of-${app.value.appId}`), 'the token leaked');
  });
});

/**
 * teamsProvisionerService.test.ts — proves the WIRING unit of wave W0b
 * (epic byte5ai/omadia#860, capability issue #3): capability assembly
 * (`createTeamsProvisioner`), service registration in `activate()`, the
 * barrel exports, and the manifest/package hub edits (version bump, ARM
 * setup fields, network allowlist, runtime_write, service_types, provides,
 * bilingual setup.guide scope + consent additions).
 */

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { PluginContext } from '@omadia/plugin-api';

import { activate } from '../src/plugin.js';
import * as barrel from '../src/index.js';
import {
  createTeamsProvisioner,
  TEAMS_PROVISIONER_SERVICE_NAME,
  type TeamsProvisionerAccessor,
} from '../src/teamsProvisioner/index.js';
import { MICROSOFT365_SERVICE_NAME } from '../src/accessor.js';
import {
  AZURE_REGION_FIELD,
  AZURE_RESOURCE_GROUP_FIELD,
  AZURE_SP_CLIENT_ID_FIELD,
  AZURE_SP_CLIENT_SECRET_FIELD,
  AZURE_SUBSCRIPTION_ID_FIELD,
  readArmConfig,
} from '../src/teamsProvisioner/config.js';

// Works both from tests/ (direct ts run) and .test-build/ (esbuild output):
// both directories sit directly under the package root.
const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const TENANT = '11111111-2222-3333-4444-555555555555';
const APP_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
const APP_SECRET = 'bot-framework-app-secret';

const ARM_CORE_CONFIG = {
  [AZURE_SUBSCRIPTION_ID_FIELD]: '22222222-3333-4444-5555-666666666666',
  [AZURE_RESOURCE_GROUP_FIELD]: 'rg-omadia-bots',
  [AZURE_REGION_FIELD]: 'global',
};

interface FakeContext {
  readonly ctx: PluginContext;
  readonly provided: Map<string, unknown>;
  readonly disposed: string[];
  readonly logs: string[];
}

/**
 * Fake `PluginContext` covering exactly what `activate()` touches: the
 * required base config/secret via `require`, the OPTIONAL ARM fields via the
 * non-throwing `get` accessors, `services.provide` and `log`.
 */
function fakeContext(
  config: Record<string, unknown> = {},
  secrets: Record<string, string> = {},
): FakeContext {
  const provided = new Map<string, unknown>();
  const disposed: string[] = [];
  const logs: string[] = [];

  const baseConfig: Record<string, unknown> = {
    microsoft_tenant_id: TENANT,
    microsoft_app_id: APP_ID,
    ...config,
  };
  const baseSecrets: Record<string, string> = {
    microsoft_app_password: APP_SECRET,
    ...secrets,
  };

  const ctx = {
    log: (msg: string): void => {
      logs.push(msg);
    },
    config: {
      get: <T = unknown>(key: string): T | undefined =>
        baseConfig[key] as T | undefined,
      require: <T = unknown>(key: string): T => {
        const value = baseConfig[key];
        if (value === undefined) {
          throw new Error(`missing config: ${key}`);
        }
        return value as T;
      },
    },
    secrets: {
      get: async (key: string): Promise<string | undefined> =>
        baseSecrets[key],
      require: async (key: string): Promise<string> => {
        const value = baseSecrets[key];
        if (value === undefined) {
          throw new Error(`missing secret: ${key}`);
        }
        return value;
      },
      keys: async (): Promise<string[]> => Object.keys(baseSecrets),
    },
    services: {
      provide: <T>(name: string, service: T): (() => void) => {
        provided.set(name, service);
        return () => {
          disposed.push(name);
        };
      },
      get: (): undefined => undefined,
    },
  } as unknown as PluginContext;

  return { ctx, provided, disposed, logs };
}

describe('activate() — service registration', () => {
  it('publishes teamsProvisioner alongside microsoft365.graph', async () => {
    const { ctx, provided } = fakeContext();

    const handle = await activate(ctx);

    assert.ok(provided.has(MICROSOFT365_SERVICE_NAME));
    assert.ok(provided.has(TEAMS_PROVISIONER_SERVICE_NAME));
    assert.equal(TEAMS_PROVISIONER_SERVICE_NAME, 'teamsProvisioner');
    await handle.close();
  });

  it('never throws when the ARM setup fields are absent (registration-only)', async () => {
    const { ctx, provided, logs } = fakeContext();

    const handle = await activate(ctx);
    const provisioner = provided.get(
      TEAMS_PROVISIONER_SERVICE_NAME,
    ) as TeamsProvisionerAccessor;

    // Degradation is observable on the published capability itself …
    assert.equal(provisioner.canCreateBots, false);
    assert.equal(provisioner.tenantMode, 'customer');
    const outcome = await provisioner.createBot({
      botName: 'omadia-agent-bot',
      displayName: 'Omadia Agent',
      msaAppId: APP_ID,
      msaAppTenantId: TENANT,
      messagingEndpoint: 'https://example.org/api/messages',
    });
    assert.equal(outcome.kind, 'registration-only');
    if (outcome.kind === 'registration-only') {
      assert.deepEqual(outcome.missingSetupFields, [
        AZURE_SUBSCRIPTION_ID_FIELD,
        AZURE_RESOURCE_GROUP_FIELD,
        AZURE_REGION_FIELD,
      ]);
    }

    // … and announced in the activation log.
    assert.ok(
      logs.some((line) => line.includes('registration-only')),
      'activation must log the registration-only degradation',
    );
    await handle.close();
  });

  it('enables bot creation when the ARM core fields are configured', async () => {
    const { ctx, provided, logs } = fakeContext(ARM_CORE_CONFIG);

    const handle = await activate(ctx);
    const provisioner = provided.get(
      TEAMS_PROVISIONER_SERVICE_NAME,
    ) as TeamsProvisionerAccessor;

    assert.equal(provisioner.canCreateBots, true);
    assert.ok(
      logs.some(
        (line) =>
          line.includes('teamsProvisioner@1') && line.includes('ARM configured'),
      ),
    );
    await handle.close();
  });

  it('degrades (never throws) on a half-configured dedicated SP', async () => {
    const { ctx, provided } = fakeContext({
      ...ARM_CORE_CONFIG,
      [AZURE_SP_CLIENT_ID_FIELD]: APP_ID,
      // azure_sp_client_secret deliberately absent.
    });

    const handle = await activate(ctx);
    const provisioner = provided.get(
      TEAMS_PROVISIONER_SERVICE_NAME,
    ) as TeamsProvisionerAccessor;

    assert.equal(provisioner.canCreateBots, false);
    await handle.close();
  });

  it('close() disposes both service registrations', async () => {
    const { ctx, disposed } = fakeContext();

    const handle = await activate(ctx);
    assert.deepEqual(disposed, []);
    await handle.close();
    assert.deepEqual(
      [...disposed].sort(),
      [MICROSOFT365_SERVICE_NAME, TEAMS_PROVISIONER_SERVICE_NAME].sort(),
    );
  });
});

describe('createTeamsProvisioner — capability assembly', () => {
  const trapFetch: typeof fetch = () => {
    throw new Error('no network call expected in this test');
  };

  it('is side-effect free at construction and defaults to the customer tenant', async () => {
    const { ctx } = fakeContext();
    const armConfig = await readArmConfig(ctx, {
      clientId: APP_ID,
      clientSecret: APP_SECRET,
    });
    const provisioner = createTeamsProvisioner({
      graphCredential: {
        tenantId: TENANT,
        clientId: APP_ID,
        clientSecret: APP_SECRET,
      },
      armConfig,
      secrets: ctx.secrets,
      fetchImpl: trapFetch,
      log: () => {},
    });

    assert.equal(provisioner.tenantMode, 'customer');
    assert.equal(provisioner.canCreateBots, false);
    // createBot answers registration-only WITHOUT touching the trap fetch.
    const outcome = await provisioner.createBot({
      botName: 'omadia-agent-bot',
      displayName: 'Omadia Agent',
      msaAppId: APP_ID,
      msaAppTenantId: TENANT,
      messagingEndpoint: 'https://example.org/api/messages',
    });
    assert.equal(outcome.kind, 'registration-only');
  });

  it('exposes every chain step plus the pure app-package builder', async () => {
    const { ctx } = fakeContext(ARM_CORE_CONFIG);
    const armConfig = await readArmConfig(ctx, {
      clientId: APP_ID,
      clientSecret: APP_SECRET,
    });
    const provisioner = createTeamsProvisioner({
      graphCredential: {
        tenantId: TENANT,
        clientId: APP_ID,
        clientSecret: APP_SECRET,
      },
      armConfig,
      secrets: ctx.secrets,
      fetchImpl: trapFetch,
      log: () => {},
    });

    assert.equal(provisioner.canCreateBots, true);
    for (const step of [
      'createAppRegistration',
      'deleteAppRegistration',
      'getAppRegistration',
      'buildAppPackage',
      'createBot',
      'deleteBot',
      'getBot',
      'uploadToCatalog',
      'getCatalogApp',
      'installToTeam',
      'uninstallFromTeam',
      // Chat install (0.7.0) — the group-chat target the team-only surface
      // could not express. Consumers feature-detect these two.
      'installToChat',
      'uninstallFromChat',
      // Delegated catalog publishing (0.6.0, byte5ai/omadia#924).
      'uploadToCatalogDelegated',
      'startDelegatedSignIn',
      'pollDelegatedSignIn',
      'getDelegatedSignInStatus',
      'refreshDelegatedToken',
      'revokeDelegatedSignIn',
      // Install-target ENUMERATION (0.8.0) — the lists a picker offers so an
      // operator never types a team or chat id by hand.
      'listTeams',
      'listChats',
      // Reset primitives (0.8.0). deleteAppRegistration and deleteBot above
      // already were the rollback halves and are reused unchanged; these two
      // are what a delete alone cannot do — free the reserved uniqueName, and
      // take the app back out of the tenant catalog.
      'purgeDeletedAppRegistration',
      'removeFromCatalog',
    ] as const) {
      assert.equal(typeof provisioner[step], 'function', `missing step ${step}`);
    }
  });

  it('answers listChats with a typed sign-in verdict, not a network call', async () => {
    // The whole delegated story in one assertion: chat listing cannot run on
    // the connector's application credentials, because Graph offers no
    // tenant-wide app-only route for chats at all. A consumer must learn that
    // from a typed error, not from a 403 it has to interpret — and learning it
    // must not cost a Graph round trip, hence the trap fetch.
    const { ctx } = fakeContext();
    const armConfig = await readArmConfig(ctx, {
      clientId: APP_ID,
      clientSecret: APP_SECRET,
    });
    const provisioner = createTeamsProvisioner({
      graphCredential: {
        tenantId: TENANT,
        clientId: APP_ID,
        clientSecret: APP_SECRET,
      },
      armConfig,
      secrets: ctx.secrets,
      fetchImpl: trapFetch,
      log: () => {},
    });

    await assert.rejects(
      () => provisioner.listChats(),
      (err: unknown) => {
        assert.ok(err instanceof barrel.DelegatedScopeRequiredError);
        assert.equal(err.reason, 'no-token');
        assert.deepEqual(err.requiredScopes, [barrel.CHAT_READ_DELEGATED_SCOPE]);
        return true;
      },
    );
  });

  it('stays side-effect free: no Graph call until a step runs', async () => {
    // The publisher app is provisioned LAZILY, on the first sign-in. Resolving
    // it during activation would make every install register an app it may
    // never use — and would make activation fail on a Graph hiccup.
    const { ctx } = fakeContext();
    const armConfig = await readArmConfig(ctx, {
      clientId: APP_ID,
      clientSecret: APP_SECRET,
    });
    const provisioner = createTeamsProvisioner({
      graphCredential: {
        tenantId: TENANT,
        clientId: APP_ID,
        clientSecret: APP_SECRET,
      },
      armConfig,
      secrets: ctx.secrets,
      fetchImpl: trapFetch,
      log: () => {},
    });

    // The pure members answer without touching the network at all.
    const signedOut = provisioner.getDelegatedSignInStatus({});
    assert.equal(signedOut.signedIn, false);
    assert.equal(
      provisioner.revokeDelegatedSignIn({}).outcome,
      'not-signed-in',
    );
  });
});

describe('barrel exports (src/index.ts)', () => {
  it('re-exports the capability constants and the assembly factory', () => {
    assert.equal(barrel.TEAMS_PROVISIONER_SERVICE_NAME, 'teamsProvisioner');
    assert.equal(barrel.TEAMS_PROVISIONER_CAPABILITY, 'teamsProvisioner@1');
    assert.equal(typeof barrel.createTeamsProvisioner, 'function');
    assert.equal(typeof barrel.readArmConfig, 'function');
    assert.equal(typeof barrel.secretRefForApp, 'function');
  });

  it('ships the ConsentRequiredError/ConsentMissingError name pair as distinct classes', () => {
    assert.equal(typeof barrel.ConsentRequiredError, 'function');
    assert.equal(typeof barrel.ConsentMissingError, 'function');
    assert.notEqual(
      barrel.ConsentRequiredError as unknown,
      barrel.ConsentMissingError as unknown,
    );
    const missing = new barrel.ConsentMissingError(['AppCatalog.ReadWrite.All'], 'graph');
    assert.equal(missing.name, 'ConsentMissingError');
    assert.ok(missing instanceof barrel.TeamsProvisionerError);
  });
});

describe('manifest.yaml / package.json hub edits', () => {
  const manifest = readFileSync(join(pkgRoot, 'manifest.yaml'), 'utf8');
  const pkg = JSON.parse(
    readFileSync(join(pkgRoot, 'package.json'), 'utf8'),
  ) as { version: string };

  it('bumps the version in BOTH files without drift', () => {
    // 0.5.3 — `teamsProvisioner@1.getTeam`, the team-name lookup.
    // 0.5.4 — the global bot-handle verdict + the stricter handle grammar (#921).
    // 0.6.0 — delegated catalog publishing via device code (#924).
    // 0.7.0 — chat install/uninstall + install-target classification.
    // 0.8.0 — install-target ENUMERATION (listTeams/listChats) + the reset
    //         primitives (purge, catalog removal).
    // 0.8.1 — the legacy `19:…@thread.skype` group chat is an install target
    //         again; listChats had been offering chats classify refused.
    assert.equal(pkg.version, '0.8.1');
    assert.ok(
      manifest.includes(`version: "${pkg.version}"`),
      'manifest.yaml must carry the same version as package.json',
    );
  });

  it('declares the teamsProvisioner service type and capability', () => {
    assert.ok(manifest.includes('service: "teamsProvisioner"'));
    assert.ok(manifest.includes('name: "TeamsProvisionerAccessor"'));
    assert.ok(manifest.includes('- "teamsProvisioner@1"'));
  });

  it('adds the OPTIONAL ARM setup fields from the arm-config unit', () => {
    for (const key of [
      AZURE_SUBSCRIPTION_ID_FIELD,
      AZURE_RESOURCE_GROUP_FIELD,
      AZURE_REGION_FIELD,
      AZURE_SP_CLIENT_ID_FIELD,
      AZURE_SP_CLIENT_SECRET_FIELD,
    ]) {
      assert.ok(
        manifest.includes(`key: "${key}"`),
        `manifest.yaml must declare setup field ${key}`,
      );
    }
    // ALL ARM fields are optional — a `required: true` on any of them would
    // break every existing install.
    const armFieldsBlock = manifest.slice(manifest.indexOf('azure_subscription_id'));
    assert.ok(!armFieldsBlock.includes('required: true'));
    // The SP secret mirrors microsoft_app_password's secret typing.
    const spSecretBlock = manifest.slice(
      manifest.indexOf(`key: "${AZURE_SP_CLIENT_SECRET_FIELD}"`),
    );
    assert.ok(spSecretBlock.includes('type: "secret"'));
  });

  it('extends permissions: ARM egress + runtime secret writes', () => {
    assert.ok(manifest.includes('- "management.azure.com"'));
    assert.match(manifest, /secrets:\s*\n\s*runtime_write: true/);
  });

  it('documents the provisioning scopes in the bilingual setup.guide', () => {
    for (const scope of [
      'Application.ReadWrite.OwnedBy',
      'AppCatalog.ReadWrite.All',
      'TeamsAppInstallation.ReadWriteForTeam.All',
      // 0.7.0 — the chat install target.
      'TeamsAppInstallation.ReadWriteForChat.All',
    ]) {
      const first = manifest.indexOf(`\`${scope}\``);
      const last = manifest.lastIndexOf(`\`${scope}\``);
      assert.ok(first !== -1, `setup.guide must list ${scope}`);
      assert.ok(last !== first, `setup.guide must list ${scope} in BOTH languages`);
    }
    assert.ok(manifest.includes('teamsProvisioner@1'));
  });

  it('documents renewed admin consent + the appRoleAssignments/restart fallback bilingually', () => {
    assert.match(manifest, /renewed admin consent/i);
    assert.ok(manifest.includes('erneuten Admin-Consent'));
    const first = manifest.indexOf('appRoleAssignments');
    const last = manifest.lastIndexOf('appRoleAssignments');
    assert.ok(first !== -1 && last !== first, 'appRoleAssignments note must be bilingual');
    assert.match(manifest, /restart the middleware/);
    assert.ok(manifest.includes('Middleware neu starten'));
  });
});

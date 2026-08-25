import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  ARM_CORE_SETUP_FIELD_KEYS,
  ARM_MANAGEMENT_HOST,
  ARM_SETUP_FIELD_KEYS,
  ARM_TOKEN_SCOPE,
  AZURE_REGION_FIELD,
  AZURE_RESOURCE_GROUP_FIELD,
  AZURE_SP_CLIENT_ID_FIELD,
  AZURE_SP_CLIENT_SECRET_FIELD,
  AZURE_SUBSCRIPTION_ID_FIELD,
  TEAMS_PROVISIONER_ARM_SETUP_FIELDS,
  isArmConfigured,
  readArmConfig,
  type ArmConfigResult,
  type ArmConfigSource,
} from '../src/teamsProvisioner/config.js';
import type { RegistrationOnlyOutcome } from '../src/teamsProvisioner/types.js';

const SUBSCRIPTION = '11111111-2222-3333-4444-555555555555';
const SP_CLIENT_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';

const REUSE_APP = {
  clientId: 'ffffffff-0000-1111-2222-333333333333',
  clientSecret: 'bot-framework-app-secret',
};

/**
 * Fake `PluginContext` slice. The `require` members THROW on purpose: the
 * reader must use `get` only — registration-only mode is a result, never a
 * `MissingConfigError` crash. Any `require` call fails the test loudly.
 */
function fakeCtx(
  config: Record<string, unknown>,
  secrets: Record<string, string> = {},
): ArmConfigSource {
  // Intermediate consts on purpose: they carry the booby-trapped `require`
  // past the excess-property check that a direct object literal would hit.
  const configAccessor = {
    get: <T = unknown>(key: string): T | undefined =>
      config[key] as T | undefined,
    require: (): never => {
      throw new Error('config.require must never be used for ARM fields');
    },
  };
  const secretsAccessor = {
    get: async (key: string): Promise<string | undefined> => secrets[key],
    require: async (): Promise<never> => {
      throw new Error('secrets.require must never be used for ARM fields');
    },
  };
  return { config: configAccessor, secrets: secretsAccessor };
}

const FULL_CORE_CONFIG = {
  [AZURE_SUBSCRIPTION_ID_FIELD]: SUBSCRIPTION,
  [AZURE_RESOURCE_GROUP_FIELD]: 'rg-omadia-bots',
  [AZURE_REGION_FIELD]: 'global',
};

function patternOf(key: string): RegExp {
  const field = TEAMS_PROVISIONER_ARM_SETUP_FIELDS.find((f) => f.key === key);
  assert.ok(field, `field spec missing for ${key}`);
  assert.ok(field.pattern, `field ${key} has no pattern`);
  return new RegExp(field.pattern);
}

describe('ARM setup-field spec (wiring-unit contract)', () => {
  it('declares all five keys, in manifest order, every one optional', () => {
    assert.deepEqual(
      TEAMS_PROVISIONER_ARM_SETUP_FIELDS.map((f) => f.key),
      [...ARM_SETUP_FIELD_KEYS],
    );
    assert.deepEqual(
      [...ARM_SETUP_FIELD_KEYS],
      [
        AZURE_SUBSCRIPTION_ID_FIELD,
        AZURE_RESOURCE_GROUP_FIELD,
        AZURE_REGION_FIELD,
        AZURE_SP_CLIENT_ID_FIELD,
        AZURE_SP_CLIENT_SECRET_FIELD,
      ],
    );
    for (const field of TEAMS_PROVISIONER_ARM_SETUP_FIELDS) {
      // required: true would break every existing install of this plugin.
      assert.equal(field.required, false, `${field.key} must stay optional`);
      assert.ok(field.label.length > 0);
      assert.ok(field.help.length > 0);
    }
  });

  it('stores the SP secret as type "secret", like microsoft_app_password', () => {
    const secretField = TEAMS_PROVISIONER_ARM_SETUP_FIELDS.find(
      (f) => f.key === AZURE_SP_CLIENT_SECRET_FIELD,
    );
    assert.ok(secretField);
    assert.equal(secretField.type, 'secret');
    // Secret values are opaque — a pattern would leak shape expectations.
    assert.equal(secretField.pattern, undefined);
    const stringFields = TEAMS_PROVISIONER_ARM_SETUP_FIELDS.filter(
      (f) => f.key !== AZURE_SP_CLIENT_SECRET_FIELD,
    );
    for (const field of stringFields) {
      assert.equal(field.type, 'string');
    }
  });

  it('validates GUIDs, resource-group names and regions via patterns', () => {
    for (const key of [AZURE_SUBSCRIPTION_ID_FIELD, AZURE_SP_CLIENT_ID_FIELD]) {
      const guid = patternOf(key);
      assert.ok(guid.test(SUBSCRIPTION));
      assert.ok(!guid.test('not-a-guid'));
    }
    const rg = patternOf(AZURE_RESOURCE_GROUP_FIELD);
    assert.ok(rg.test('rg-omadia_bots.v2(prod)'));
    assert.ok(!rg.test(''));
    assert.ok(!rg.test('bad/name'));
    const region = patternOf(AZURE_REGION_FIELD);
    assert.ok(region.test('global'));
    assert.ok(region.test('westeurope'));
    assert.ok(!region.test('West Europe'));
  });

  it('pins the ARM egress host and token scope for the wiring unit', () => {
    assert.equal(ARM_MANAGEMENT_HOST, 'management.azure.com');
    assert.equal(ARM_TOKEN_SCOPE, 'https://management.azure.com/.default');
    assert.deepEqual(
      [...ARM_CORE_SETUP_FIELD_KEYS],
      [AZURE_SUBSCRIPTION_ID_FIELD, AZURE_RESOURCE_GROUP_FIELD, AZURE_REGION_FIELD],
    );
  });
});

describe('readArmConfig — configured mode', () => {
  it('returns the dedicated SP when both credential halves are set', async () => {
    const result = await readArmConfig(
      fakeCtx(
        { ...FULL_CORE_CONFIG, [AZURE_SP_CLIENT_ID_FIELD]: SP_CLIENT_ID },
        { [AZURE_SP_CLIENT_SECRET_FIELD]: 'dedicated-sp-secret' },
      ),
      REUSE_APP,
    );
    assert.equal(result.kind, 'configured');
    assert.ok(isArmConfigured(result));
    assert.equal(result.subscriptionId, SUBSCRIPTION);
    assert.equal(result.resourceGroup, 'rg-omadia-bots');
    assert.equal(result.region, 'global');
    assert.deepEqual(result.credential, {
      clientId: SP_CLIENT_ID,
      clientSecret: 'dedicated-sp-secret',
      source: 'dedicated-sp',
    });
  });

  it("reuses the app credential when the SP pair is omitted ('reuse app')", async () => {
    const result = await readArmConfig(fakeCtx(FULL_CORE_CONFIG), REUSE_APP);
    assert.ok(isArmConfigured(result));
    assert.deepEqual(result.credential, {
      clientId: REUSE_APP.clientId,
      clientSecret: REUSE_APP.clientSecret,
      source: 'reused-app',
    });
  });

  it('trims whitespace around configured values', async () => {
    const result = await readArmConfig(
      fakeCtx({
        [AZURE_SUBSCRIPTION_ID_FIELD]: `  ${SUBSCRIPTION}  `,
        [AZURE_RESOURCE_GROUP_FIELD]: 'rg-omadia-bots ',
        [AZURE_REGION_FIELD]: ' global',
      }),
      REUSE_APP,
    );
    assert.ok(isArmConfigured(result));
    assert.equal(result.subscriptionId, SUBSCRIPTION);
    assert.equal(result.resourceGroup, 'rg-omadia-bots');
    assert.equal(result.region, 'global');
  });
});

describe('readArmConfig — registration-only degradation', () => {
  it('degrades with all three core keys when nothing is configured', async () => {
    // require() throws in the fake: reaching the assertion proves the reader
    // survives a completely ARM-less install without touching require.
    const result = await readArmConfig(fakeCtx({}), REUSE_APP);
    assert.equal(result.kind, 'registration-only');
    assert.ok(!isArmConfigured(result));
    assert.equal(result.reason, 'arm-not-configured');
    assert.deepEqual(result.missingSetupFields, [...ARM_CORE_SETUP_FIELD_KEYS]);
  });

  it('names exactly the absent core fields', async () => {
    const result = await readArmConfig(
      fakeCtx({ [AZURE_SUBSCRIPTION_ID_FIELD]: SUBSCRIPTION }),
      REUSE_APP,
    );
    assert.equal(result.kind, 'registration-only');
    assert.deepEqual(result.missingSetupFields, [
      AZURE_RESOURCE_GROUP_FIELD,
      AZURE_REGION_FIELD,
    ]);
  });

  it('treats blank and non-string config values as absent', async () => {
    const result = await readArmConfig(
      fakeCtx({
        [AZURE_SUBSCRIPTION_ID_FIELD]: '   ',
        [AZURE_RESOURCE_GROUP_FIELD]: 42,
        [AZURE_REGION_FIELD]: 'global',
      }),
      REUSE_APP,
    );
    assert.equal(result.kind, 'registration-only');
    assert.deepEqual(result.missingSetupFields, [
      AZURE_SUBSCRIPTION_ID_FIELD,
      AZURE_RESOURCE_GROUP_FIELD,
    ]);
  });

  it('degrades on a half-configured SP: id without secret', async () => {
    const result = await readArmConfig(
      fakeCtx({ ...FULL_CORE_CONFIG, [AZURE_SP_CLIENT_ID_FIELD]: SP_CLIENT_ID }),
      REUSE_APP,
    );
    assert.equal(result.kind, 'registration-only');
    assert.deepEqual(result.missingSetupFields, [AZURE_SP_CLIENT_SECRET_FIELD]);
  });

  it('degrades on a half-configured SP: secret without id', async () => {
    const result = await readArmConfig(
      fakeCtx(FULL_CORE_CONFIG, {
        [AZURE_SP_CLIENT_SECRET_FIELD]: 'dedicated-sp-secret',
      }),
      REUSE_APP,
    );
    assert.equal(result.kind, 'registration-only');
    assert.deepEqual(result.missingSetupFields, [AZURE_SP_CLIENT_ID_FIELD]);
  });

  it('produces the shared RegistrationOnlyOutcome shape from types.ts', async () => {
    const result: ArmConfigResult = await readArmConfig(fakeCtx({}), REUSE_APP);
    assert.equal(result.kind, 'registration-only');
    if (result.kind !== 'registration-only') {
      assert.fail('expected registration-only');
    }
    // Compile-time proof: the degraded config result IS the typed outcome
    // createBot returns — one observable degradation shape everywhere.
    const outcome: RegistrationOnlyOutcome = result;
    assert.equal(outcome.reason, 'arm-not-configured');
    assert.ok(outcome.missingSetupFields.length > 0);
  });
});

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import type { SecretsAccessor } from '@omadia/plugin-api';

import {
  CapabilityUnavailableError,
  TEAMS_BOT_PASSWORD_SECRET_PREFIX,
  appIdFromSecretRef,
  deleteAppPassword,
  secretRefForApp,
  storeAppPassword,
} from '../src/teamsProvisioner/secretStore.js';
import { TeamsProvisionerError } from '../src/teamsProvisioner/errors.js';

const APP_ID = '11111111-2222-3333-4444-555555555555';
const OTHER_APP_ID = '99999999-8888-7777-6666-555555555555';
const PASSWORD = 'super~secret~app~password~value';

/**
 * In-memory fake of the plugin-api `SecretsAccessor`. `writable: false`
 * models a kernel that did NOT hand out runtime_write (set/delete are
 * `undefined`, exactly like the real optional members).
 */
function fakeSecrets(opts: { writable: boolean }): {
  accessor: SecretsAccessor;
  vault: Map<string, string>;
  deleted: string[];
} {
  const vault = new Map<string, string>();
  const deleted: string[] = [];
  const accessor: SecretsAccessor = {
    get: async (key) => vault.get(key),
    require: async (key) => {
      const value = vault.get(key);
      if (value === undefined) throw new Error(`missing_secret: ${key}`);
      return value;
    },
    keys: async () => [...vault.keys()],
    ...(opts.writable
      ? {
          set: async (key: string, value: string) => {
            vault.set(key, value);
          },
          delete: async (key: string) => {
            deleted.push(key);
            vault.delete(key);
          },
        }
      : {}),
  };
  return { accessor, vault, deleted };
}

describe('deterministic secretRef key scheme', () => {
  it('builds teams_bot_password:<appId> and is deterministic', () => {
    assert.equal(TEAMS_BOT_PASSWORD_SECRET_PREFIX, 'teams_bot_password');
    const ref = secretRefForApp(APP_ID);
    assert.equal(ref, `teams_bot_password:${APP_ID}`);
    // Deterministic: same app id, same key — re-runs overwrite, never fork.
    assert.equal(secretRefForApp(APP_ID), ref);
  });

  it('is collision-free across agents (distinct app ids, distinct keys)', () => {
    assert.notEqual(secretRefForApp(APP_ID), secretRefForApp(OTHER_APP_ID));
  });

  it('is reversible for the deleteAppRegistration rollback', () => {
    assert.equal(appIdFromSecretRef(secretRefForApp(APP_ID)), APP_ID);
  });

  it('rejects app ids that would break the ref scheme', () => {
    for (const bad of ['', ' ', 'has space', 'has:colon', 'tab\there']) {
      assert.throws(() => secretRefForApp(bad), /invalid_app_id/);
    }
  });

  it('rejects malformed refs instead of guessing', () => {
    for (const bad of [
      'microsoft_app_password',
      'teams_bot_password',
      'teams_bot_password:',
      `other_prefix:${APP_ID}`,
    ]) {
      assert.throws(() => appIdFromSecretRef(bad), /invalid_/);
    }
  });
});

describe('storeAppPassword', () => {
  it('persists via ctx.secrets.set and returns the opaque ref, never the value', async () => {
    const { accessor, vault } = fakeSecrets({ writable: true });

    const ref = await storeAppPassword(accessor, APP_ID, PASSWORD);

    assert.equal(ref, `teams_bot_password:${APP_ID}`);
    // The ref is a key/handle — the password must not leak into it.
    assert.ok(!ref.includes(PASSWORD));
    // The vault holds the value under exactly that key.
    assert.equal(vault.get(ref), PASSWORD);
    assert.equal(vault.size, 1);
  });

  it('overwrites on re-run with the same appId (rotation, not duplication)', async () => {
    const { accessor, vault } = fakeSecrets({ writable: true });

    const first = await storeAppPassword(accessor, APP_ID, PASSWORD);
    const second = await storeAppPassword(accessor, APP_ID, 'rotated-value');

    assert.equal(first, second);
    assert.equal(vault.size, 1);
    assert.equal(vault.get(second), 'rotated-value');
  });

  it('throws CapabilityUnavailableError when the kernel gave no write access', async () => {
    const { accessor, vault } = fakeSecrets({ writable: false });

    await assert.rejects(
      storeAppPassword(accessor, APP_ID, PASSWORD),
      (err: unknown) => {
        assert.ok(err instanceof CapabilityUnavailableError);
        // Part of the taxonomy: one instanceof catches everything.
        assert.ok(err instanceof TeamsProvisionerError);
        assert.equal(err.name, 'CapabilityUnavailableError');
        assert.equal(err.message, 'capability_unavailable');
        assert.equal(err.operation, 'set');
        assert.equal(err.missingPermission, 'permissions.secrets.runtime_write');
        // The secret value must never leak into the error.
        const serialized =
          JSON.stringify(err, Object.getOwnPropertyNames(err)) +
          String(err.stack);
        assert.ok(!serialized.includes(PASSWORD));
        return true;
      },
    );
    assert.equal(vault.size, 0);
  });

  it('refuses to persist an empty secret value', async () => {
    const { accessor, vault } = fakeSecrets({ writable: true });
    await assert.rejects(
      storeAppPassword(accessor, APP_ID, ''),
      /empty_secret_value/,
    );
    assert.equal(vault.size, 0);
  });

  it('validates the appId before touching the vault', async () => {
    const { accessor, vault } = fakeSecrets({ writable: true });
    await assert.rejects(
      storeAppPassword(accessor, 'bad:app:id', PASSWORD),
      /invalid_app_id/,
    );
    assert.equal(vault.size, 0);
  });
});

describe('deleteAppPassword (rollback)', () => {
  it('deletes exactly the stored key via ctx.secrets.delete', async () => {
    const { accessor, vault, deleted } = fakeSecrets({ writable: true });
    const ref = await storeAppPassword(accessor, APP_ID, PASSWORD);

    await deleteAppPassword(accessor, ref);

    assert.deepEqual(deleted, [ref]);
    assert.equal(vault.size, 0);
  });

  it('is a no-op for an absent key (idempotent rollback)', async () => {
    const { accessor, deleted } = fakeSecrets({ writable: true });
    await deleteAppPassword(accessor, secretRefForApp(APP_ID));
    assert.deepEqual(deleted, [secretRefForApp(APP_ID)]);
  });

  it('rejects refs outside the teams_bot_password namespace before deleting', async () => {
    const { accessor, deleted } = fakeSecrets({ writable: true });
    // Never delete a foreign vault entry, e.g. the connector's own
    // 'microsoft_app_password'.
    await assert.rejects(
      deleteAppPassword(accessor, 'microsoft_app_password'),
      /invalid_secret_ref/,
    );
    assert.deepEqual(deleted, []);
  });

  it('throws CapabilityUnavailableError when delete was not handed out', async () => {
    const { accessor } = fakeSecrets({ writable: false });
    await assert.rejects(
      deleteAppPassword(accessor, secretRefForApp(APP_ID)),
      (err: unknown) => {
        assert.ok(err instanceof CapabilityUnavailableError);
        assert.equal(err.operation, 'delete');
        return true;
      },
    );
  });
});

/**
 * manifestSetup.test.ts — proves the provisioning credentials documentation
 * (wave W0b, epic byte5ai/omadia#860, credentials issue #2).
 *
 * Asserts that README.md and INTEGRATION.md document the three additional
 * Graph application scopes required by teamsProvisioner@1 and the
 * renewed-admin-consent semantics (including the REST appRoleAssignments +
 * restart fallback).
 *
 * NOTE for the wiring unit: once manifest.yaml's bilingual setup.guide
 * (en + de) gains the scope list and consent text authored in
 * INTEGRATION.md ("Manifest `setup.guide` wording"), extend this suite with
 * the equivalent assertions against manifest.yaml — the docs assertions
 * below stay untouched.
 */

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

// Works both from tests/ (direct ts run) and .test-build/ (esbuild output):
// both directories sit directly under the package root.
const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const PROVISIONING_SCOPES = [
  'Application.ReadWrite.OwnedBy',
  'AppCatalog.ReadWrite.All',
  'TeamsAppInstallation.ReadWriteForTeam.All',
] as const;

const docs: ReadonlyArray<readonly [name: string, content: string]> = [
  ['README.md', readFileSync(join(pkgRoot, 'README.md'), 'utf8')],
  ['INTEGRATION.md', readFileSync(join(pkgRoot, 'INTEGRATION.md'), 'utf8')],
];

describe('provisioning Graph application scopes are documented', () => {
  for (const [name, content] of docs) {
    it(`${name} lists all three teamsProvisioner@1 scopes`, () => {
      for (const scope of PROVISIONING_SCOPES) {
        assert.ok(
          content.includes(`\`${scope}\``),
          `${name} must list the Graph application scope \`${scope}\``,
        );
      }
    });

    it(`${name} keeps the scopes tied to the provisioning capability`, () => {
      assert.ok(
        content.includes('teamsProvisioner@1'),
        `${name} must attribute the extra scopes to teamsProvisioner@1`,
      );
    });
  }
});

describe('renewed admin consent is documented', () => {
  for (const [name, content] of docs) {
    it(`${name} states that scope changes require renewed admin consent`, () => {
      assert.match(
        content,
        /renewed admin consent/i,
        `${name} must document that extending scopes requires renewed admin consent`,
      );
      // \s+ tolerates markdown hard-wrapping inside the phrase.
      assert.match(
        content,
        /Grant\s+admin\s+consent/,
        `${name} must name the portal action (Grant admin consent)`,
      );
    });

    it(`${name} documents the REST appRoleAssignments + restart fallback`, () => {
      assert.ok(
        content.includes('appRoleAssignments'),
        `${name} must document the REST appRoleAssignments fallback`,
      );
      assert.match(
        content,
        /restart/i,
        `${name} must document that a restart is needed for a fresh token`,
      );
    });
  }
});

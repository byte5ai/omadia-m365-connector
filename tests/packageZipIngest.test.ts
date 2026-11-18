/**
 * packageZipIngest.test.ts — the release zip, as the kernel will judge it.
 *
 * Runs the real `npm run package` script and pushes every entry of the zip it
 * writes through a mirror of the kernel's zip ingest rules. A package the
 * kernel refuses is published, offered as the latest version by the Hub, and
 * answers 422 on every install (byte5ai/omadia#2043, rollout defect D1, in
 * the Teams channel). Needs `npm run build` first: the script packages `dist/`.
 */

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  APP_PACKAGE_EXTENSIONS,
  BASENAME_ALLOWLIST,
  DECL_EXTENSIONS,
  EXTENSION_ALLOWLIST,
  UI_BUNDLE_EXTENSIONS,
  kernelIngestRefusal,
  kernelIngestRules,
  readZipEntries,
} from './support/kernelZipIngest.js';

// Works both from tests/ and .test-build/: both sit directly under the root.
const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const pkgJson = JSON.parse(readFileSync(join(pkgRoot, 'package.json'), 'utf8')) as {
  name: string;
  version: string;
};
const safeName = pkgJson.name.replace(/^@/, '').replace(/\//g, '-');
const zipPath = join(pkgRoot, 'out', `${safeName}-${pkgJson.version}.zip`);

describe('release artifact against the kernel zip ingest', () => {
  it('npm run package builds the zip', () => {
    assert.ok(
      existsSync(join(pkgRoot, 'dist', 'plugin.js')),
      'dist/plugin.js is missing: run `npm run build` before `npm test`',
    );
    const res = spawnSync(process.execPath, [join(pkgRoot, 'scripts', 'build-zip.mjs')], {
      cwd: pkgRoot,
      encoding: 'utf8',
    });
    assert.equal(res.status, 0, `build-zip.mjs failed:\n${res.stdout}\n${res.stderr}`);
    assert.ok(existsSync(zipPath), `missing release zip: ${zipPath}`);
  });

  // Runs after the test above, in the same suite: the zip read is the one just built.
  it('the built zip contains no entry the kernel ingest refuses', () => {
    const entries = readZipEntries(zipPath);

    assert.ok(entries.length > 50, `only ${String(entries.length)} entries read from the zip`);
    const refused = entries.flatMap((entry) => {
      const reason = kernelIngestRefusal(entry);
      return reason ? [`${entry.name}: ${reason}`] : [];
    });
    assert.deepEqual(refused, [], 'the kernel would answer 422 for this package');

    const names = entries.map((entry) => entry.name);
    for (const required of ['manifest.yaml', 'package.json', 'dist/plugin.js', 'dist/index.js']) {
      assert.ok(names.includes(required), `the zip is missing ${required}`);
    }
    assert.ok(!names.some((name) => name.includes('node_modules/')), 'node_modules is in the zip');
  });

  it('the ingest mirror refuses what the kernel refuses', () => {
    const file = (name: string) => ({ name, unixMode: 0o100644 });
    assert.equal(kernelIngestRefusal(file('appPackage/TEMPLATE_REVISION')), 'zip.forbidden_extension (<none>)');
    assert.equal(kernelIngestRefusal(file('appPackage/manifest.json.template')), undefined);
    assert.equal(kernelIngestRefusal(file('dist/manifest.json.template')), 'zip.forbidden_extension (.template)');
    assert.equal(kernelIngestRefusal(file('assets/ui/font.woff2')), undefined);
    assert.equal(kernelIngestRefusal(file('assets/font.woff2')), 'zip.forbidden_extension (.woff2)');
    assert.equal(kernelIngestRefusal(file('dist/style.css')), 'zip.forbidden_extension (.css)');
    assert.equal(kernelIngestRefusal(file('dist/Makefile')), 'zip.forbidden_extension (<none>)');
    assert.equal(kernelIngestRefusal(file('dist/index.d.ts')), undefined);
    assert.equal(kernelIngestRefusal(file('dist/index.js.map')), undefined);
    assert.equal(kernelIngestRefusal(file('LICENSE')), undefined);
    assert.equal(kernelIngestRefusal(file('dist/')), undefined);
    assert.equal(kernelIngestRefusal(file('dist/../x.js')), 'zip.path_escape');
    assert.equal(kernelIngestRefusal({ name: 'dist/link.js', unixMode: 0o120777 }), 'zip.symlink');
  });

  it('the ingest mirror states the same sets as the kernel source, where a kernel checkout is present', () => {
    // Without a kernel checkout next to this repo there is nothing to compare
    // and the mirror stands as written. Locally the sibling `../odoo-bot` is
    // the kernel this connector is built against.
    const rules = kernelIngestRules(join(pkgRoot, '..', 'odoo-bot', 'middleware', 'src', 'plugins', 'zipExtractor.ts'));
    if (!rules) return;
    const sorted = (list: readonly string[]) => [...list].sort();
    assert.deepEqual(sorted(rules.extensionAllowlist), sorted(EXTENSION_ALLOWLIST));
    assert.deepEqual(sorted(rules.declExtensions), sorted(DECL_EXTENSIONS));
    assert.deepEqual(sorted(rules.uiBundleExtensions), sorted(UI_BUNDLE_EXTENSIONS));
    assert.deepEqual(sorted(rules.appPackageExtensions), sorted(APP_PACKAGE_EXTENSIONS));
    assert.deepEqual(sorted(rules.basenames), sorted(BASENAME_ALLOWLIST));
  });
});

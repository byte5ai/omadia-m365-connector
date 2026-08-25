import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { strFromU8, unzipSync } from 'fflate';

import {
  APP_PACKAGE_COLOR_ICON_ENTRY,
  APP_PACKAGE_MANIFEST_ENTRY,
  APP_PACKAGE_OUTLINE_ICON_ENTRY,
  AppPackageError,
  buildAppPackage,
  type AppPackageParams,
} from '../src/teamsProvisioner/appPackage.js';
import { TeamsProvisionerError } from '../src/teamsProvisioner/errors.js';

/**
 * Fixture mirroring the placeholder contract of the canonical template in
 * omadia-channel-teams (appPackage/manifest.json.template): every placeholder
 * sits inside JSON-string quotes EXCEPT {{VALID_DOMAINS}}, which is a raw
 * JSON-array slot — the fixture, like the real template, is not valid JSON
 * until rendered. The real template is caller input and deliberately not
 * vendored into this repo.
 */
const TEMPLATE = `{
  "manifestVersion": "1.17",
  "version": "{{VERSION}}",
  "id": "{{APP_ID}}",
  "name": { "short": "{{NAME_SHORT}}", "full": "{{NAME_FULL}}" },
  "icons": { "color": "color.png", "outline": "outline.png" },
  "bots": [{ "botId": "{{BOT_ID}}" }],
  "validDomains": {{VALID_DOMAINS}},
  "webApplicationInfo": { "id": "{{BOT_ID}}", "resource": "api://{{MIDDLEWARE_HOST}}/{{BOT_ID}}" }
}`;

const PARAMS: AppPackageParams = {
  VERSION: '1.3.0',
  APP_ID: 'bc0bd6cf-7037-4c7c-9e29-de86c4b77177',
  BOT_ID: '737c6ddd-6d4e-4599-8dc3-260281ea906e',
  NAME_SHORT: 'omadia-agent',
  NAME_FULL: 'omadia-agent — byte5 "HR" Assistent \\ Umläute',
  VALID_DOMAINS: ['odoo-bot-middleware.fly.dev', 'odoo-bot-harness.fly.dev'],
  MIDDLEWARE_HOST: 'odoo-bot-middleware.fly.dev',
};

const COLOR_PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1]);
const OUTLINE_PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 2]);
const ICONS = { color: COLOR_PNG, outline: OUTLINE_PNG };

function buildDefault(overrides: Partial<Parameters<typeof buildAppPackage>[0]> = {}) {
  return buildAppPackage({
    manifestTemplate: TEMPLATE,
    params: PARAMS,
    icons: ICONS,
    ...overrides,
  });
}

function assertAppPackageError(problemFragment: string) {
  return (err: unknown): boolean => {
    assert.ok(err instanceof Error);
    assert.ok(err instanceof TeamsProvisionerError);
    assert.ok(err instanceof AppPackageError);
    assert.equal(err.name, 'AppPackageError');
    assert.equal(err.message, 'app_package_invalid');
    assert.ok(
      err.problems.some((p) => p.includes(problemFragment)),
      `problems ${JSON.stringify(err.problems)} should mention ${JSON.stringify(problemFragment)}`,
    );
    return true;
  };
}

describe('buildAppPackage', () => {
  it('produces a zip with exactly manifest.json, color.png and outline.png at the root', () => {
    const zip = buildDefault();
    assert.ok(zip instanceof Uint8Array);
    assert.ok(zip.length > 0);

    const entries = unzipSync(zip);
    assert.deepEqual(Object.keys(entries).sort(), [
      APP_PACKAGE_COLOR_ICON_ENTRY,
      APP_PACKAGE_MANIFEST_ENTRY,
      APP_PACKAGE_OUTLINE_ICON_ENTRY,
    ]);
    assert.deepEqual(entries[APP_PACKAGE_COLOR_ICON_ENTRY], COLOR_PNG);
    assert.deepEqual(entries[APP_PACKAGE_OUTLINE_ICON_ENTRY], OUTLINE_PNG);
  });

  it('substitutes string params JSON-escaped and raw-array slots verbatim', () => {
    const entries = unzipSync(buildDefault());
    const manifestBytes = entries[APP_PACKAGE_MANIFEST_ENTRY];
    assert.ok(manifestBytes !== undefined);
    const manifest = JSON.parse(strFromU8(manifestBytes)) as {
      version: string;
      id: string;
      name: { short: string; full: string };
      bots: readonly { botId: string }[];
      validDomains: readonly string[];
      webApplicationInfo: { id: string; resource: string };
    };

    assert.equal(manifest.version, '1.3.0');
    assert.equal(manifest.id, PARAMS['APP_ID']);
    // JSON-escaping round-trips quotes, backslashes and non-ASCII intact.
    assert.equal(manifest.name.full, 'omadia-agent — byte5 "HR" Assistent \\ Umläute');
    // {{VALID_DOMAINS}} is the raw JSON-array slot.
    assert.deepEqual(manifest.validDomains, [
      'odoo-bot-middleware.fly.dev',
      'odoo-bot-harness.fly.dev',
    ]);
    // A repeated placeholder ({{BOT_ID}}) substitutes at EVERY occurrence.
    assert.equal(manifest.bots[0]?.botId, PARAMS['BOT_ID']);
    assert.equal(manifest.webApplicationInfo.id, PARAMS['BOT_ID']);
    assert.equal(
      manifest.webApplicationInfo.resource,
      `api://${String(PARAMS['MIDDLEWARE_HOST'])}/${String(PARAMS['BOT_ID'])}`,
    );
  });

  it('is deterministic — same inputs yield byte-identical zips', () => {
    assert.deepEqual(buildDefault(), buildDefault());
  });

  it('rejects a template placeholder without a param', () => {
    const { MIDDLEWARE_HOST: _omitted, ...incomplete } = PARAMS;
    assert.throws(
      () => buildDefault({ params: incomplete }),
      assertAppPackageError('missing params for placeholders: MIDDLEWARE_HOST'),
    );
  });

  it('rejects a param without a matching placeholder (typo guard)', () => {
    assert.throws(
      () => buildDefault({ params: { ...PARAMS, MIDLEWARE_HOST: 'oops' } }),
      assertAppPackageError('params without a matching placeholder: MIDLEWARE_HOST'),
    );
  });

  it('collects every problem in one error instead of failing one at a time', () => {
    const { VERSION: _omitted, ...rest } = PARAMS;
    assert.throws(
      () => buildDefault({ params: { ...rest, TYPO: 'x' }, icons: { color: new Uint8Array(), outline: OUTLINE_PNG } }),
      (err: unknown) => {
        assert.ok(err instanceof AppPackageError);
        assert.equal(err.problems.length, 3);
        return true;
      },
    );
  });

  it('rejects empty icon buffers', () => {
    assert.throws(
      () => buildDefault({ icons: { color: COLOR_PNG, outline: new Uint8Array() } }),
      assertAppPackageError('icons.outline is empty'),
    );
  });

  it('rejects a render that is not valid JSON (string param in a raw-array slot)', () => {
    assert.throws(
      () => buildDefault({ params: { ...PARAMS, VALID_DOMAINS: 'not-an-array' } }),
      assertAppPackageError('rendered manifest is not a valid JSON object'),
    );
  });

  it('rejects a manifest that references other icon file names', () => {
    const template = TEMPLATE.replace('"outline.png"', '"sketch.png"');
    assert.throws(
      () => buildDefault({ manifestTemplate: template }),
      assertAppPackageError('icons.outline must be "outline.png"'),
    );
  });
});

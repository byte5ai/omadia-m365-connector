/**
 * `buildAppPackage` — render a Teams app-package ZIP (manifest + icons) fully
 * in memory.
 *
 * Part of `teamsProvisioner@1` (epic byte5ai/omadia#860, capability issue
 * byte5ai/omadia-m365-connector#3): the agent factory renders one app package
 * per agent identity and feeds the resulting buffer to `uploadToCatalog`.
 *
 * The manifest TEMPLATE is caller input, deliberately not vendored here. The
 * canonical template lives in `omadia-channel-teams` (`appPackage/
 * manifest.json.template`, wave W0a) and its README is the rendering
 * contract this function implements:
 *
 * - Placeholders use `{{NAME}}` syntax.
 * - Every placeholder substitutes a JSON STRING value — the template carries
 *   the surrounding quotes, so the generator inserts the JSON-escaped string
 *   content and must not add quotes.
 * - A placeholder given a string ARRAY (the template's raw-JSON slots, e.g.
 *   `{{VALID_DOMAINS}}`) substitutes the serialized array verbatim; the
 *   template deliberately writes such slots unquoted, which is why the raw
 *   template is not valid JSON until rendered.
 *
 * The rendered package ZIP contains exactly three files at the archive root
 * (`manifest.json`, `color.png`, `outline.png`); the manifest references the
 * icons by those fixed names. No filesystem is touched — this connector's
 * manifest declares `permissions.filesystem.scratch: false` — so the ZIP is
 * produced fully in memory by the dependency-free writer in `zip.ts`
 * (`node:zlib` deflate; see that module's doc for why no runtime dependency
 * is possible here) and returned as a buffer.
 */

import { TeamsProvisionerError } from './errors.js';
import { createZip } from './zip.js';

/** Archive entry name of the rendered manifest. */
export const APP_PACKAGE_MANIFEST_ENTRY = 'manifest.json';
/** Archive entry name of the 192×192 color icon (fixed by the manifest). */
export const APP_PACKAGE_COLOR_ICON_ENTRY = 'color.png';
/** Archive entry name of the 32×32 transparent outline icon (fixed). */
export const APP_PACKAGE_OUTLINE_ICON_ENTRY = 'outline.png';

/**
 * Fixed timestamp stamped on every archive entry so that rendering the same
 * inputs twice yields byte-identical ZIPs (re-runs of the provisioning chain
 * stay comparable; nothing downstream reads the entry mtime).
 */
const APP_PACKAGE_ENTRY_MTIME = new Date('2020-01-01T00:00:00Z');

/**
 * Placeholder syntax of the channel-teams template: `{{UPPER_SNAKE_CASE}}`.
 */
const PLACEHOLDER_RE = /\{\{([A-Z][A-Z0-9_]*)\}\}/g;

/** Per-agent icon assets, replacing the template's default omadia icons. */
export interface AppPackageIcons {
  /** PNG bytes for `color.png` (192×192). */
  readonly color: Uint8Array;
  /** PNG bytes for `outline.png` (32×32, transparent). */
  readonly outline: Uint8Array;
}

/**
 * Value for one `{{NAME}}` placeholder. A `string` is substituted as the
 * JSON-escaped string CONTENT (the template supplies the quotes); a
 * `readonly string[]` is substituted as a raw JSON array (the template's
 * unquoted slots, e.g. `{{VALID_DOMAINS}}`).
 */
export type AppPackageParamValue = string | readonly string[];

/** Placeholder name → substitution value, keyed WITHOUT the `{{ }}` braces. */
export type AppPackageParams = Readonly<Record<string, AppPackageParamValue>>;

/** Input for {@link buildAppPackage}. */
export interface BuildAppPackageInput {
  /**
   * The `manifest.json.template` text from `omadia-channel-teams`
   * (`appPackage/` directory) — caller-supplied, never vendored here.
   */
  readonly manifestTemplate: string;
  /** One value per `{{NAME}}` placeholder occurring in the template. */
  readonly params: AppPackageParams;
  /** The per-agent icon PNGs. */
  readonly icons: AppPackageIcons;
}

/**
 * Thrown when the template, the params or the icons do not line up (missing
 * or unused placeholders, a render that is not valid JSON, empty icon
 * buffers, a manifest that references other icon file names). Carries every
 * detected problem so callers fix one round-trip, not one field at a time.
 */
export class AppPackageError extends TeamsProvisionerError {
  /** Human-readable description of each detected problem. */
  public readonly problems: readonly string[];

  constructor(problems: readonly string[], cause?: unknown) {
    super('app_package_invalid', cause);
    this.name = 'AppPackageError';
    this.problems = problems;
  }
}

/**
 * Render the manifest template with `params`, bundle it with the icons and
 * return the Teams app-package ZIP as an in-memory buffer (the shape
 * `UploadToCatalogInput.packageZip` expects). Throws {@link AppPackageError}
 * on any contract violation; never touches the filesystem.
 */
export function buildAppPackage(input: BuildAppPackageInput): Uint8Array {
  const { manifestTemplate, params, icons } = input;

  const problems: string[] = [
    ...validatePlaceholderCoverage(manifestTemplate, params),
    ...validateIcons(icons),
  ];
  if (problems.length > 0) {
    throw new AppPackageError(problems);
  }

  const rendered = renderTemplate(manifestTemplate, params);
  const manifest = parseRenderedManifest(rendered);
  const iconProblems = validateManifestIconRefs(manifest);
  if (iconProblems.length > 0) {
    throw new AppPackageError(iconProblems);
  }

  return createZip(
    [
      { name: APP_PACKAGE_MANIFEST_ENTRY, data: new TextEncoder().encode(rendered) },
      { name: APP_PACKAGE_COLOR_ICON_ENTRY, data: icons.color },
      { name: APP_PACKAGE_OUTLINE_ICON_ENTRY, data: icons.outline },
    ],
    APP_PACKAGE_ENTRY_MTIME,
  );
}

/** Every placeholder needs a param; every param needs a placeholder. */
function validatePlaceholderCoverage(
  template: string,
  params: AppPackageParams,
): readonly string[] {
  const placeholders = new Set<string>();
  for (const match of template.matchAll(PLACEHOLDER_RE)) {
    const name = match[1];
    if (name !== undefined) {
      placeholders.add(name);
    }
  }

  const provided = new Set(Object.keys(params));
  const missing = [...placeholders].filter((name) => !provided.has(name));
  const unused = [...provided].filter((name) => !placeholders.has(name));

  const problems: string[] = [];
  if (missing.length > 0) {
    problems.push(`missing params for placeholders: ${missing.join(', ')}`);
  }
  if (unused.length > 0) {
    // An unused param is almost always a typo'd placeholder name — failing
    // fast here beats shipping a manifest with a silently absent value.
    problems.push(`params without a matching placeholder: ${unused.join(', ')}`);
  }
  return problems;
}

function validateIcons(icons: AppPackageIcons): readonly string[] {
  const problems: string[] = [];
  if (icons.color.length === 0) {
    problems.push('icons.color is empty');
  }
  if (icons.outline.length === 0) {
    problems.push('icons.outline is empty');
  }
  return problems;
}

function renderTemplate(template: string, params: AppPackageParams): string {
  return template.replace(PLACEHOLDER_RE, (token, name: string) => {
    const value = params[name];
    if (value === undefined) {
      // Unreachable after validatePlaceholderCoverage; kept for safety under
      // noUncheckedIndexedAccess.
      throw new AppPackageError([`missing params for placeholders: ${name}`]);
    }
    if (typeof value === 'string') {
      // JSON-escape the CONTENT only — the template supplies the quotes.
      return JSON.stringify(value).slice(1, -1);
    }
    // Raw-JSON slot ({{VALID_DOMAINS}} style): serialized array, verbatim.
    return JSON.stringify(value);
  });
}

function parseRenderedManifest(rendered: string): Record<string, unknown> {
  try {
    const parsed: unknown = JSON.parse(rendered);
    if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
      throw new Error('manifest root is not a JSON object');
    }
    return parsed as Record<string, unknown>;
  } catch (cause) {
    throw new AppPackageError(
      ['rendered manifest is not a valid JSON object'],
      cause,
    );
  }
}

/**
 * The archive always carries `color.png` / `outline.png`; a manifest that
 * references anything else would ship a package pointing at files that are
 * not in the ZIP, which Teams rejects at upload time. Catch it here instead.
 */
function validateManifestIconRefs(
  manifest: Record<string, unknown>,
): readonly string[] {
  const icons = manifest['icons'];
  if (typeof icons !== 'object' || icons === null) {
    return ['rendered manifest has no "icons" object'];
  }
  const { color, outline } = icons as { color?: unknown; outline?: unknown };

  const problems: string[] = [];
  if (color !== APP_PACKAGE_COLOR_ICON_ENTRY) {
    problems.push(
      `manifest icons.color must be "${APP_PACKAGE_COLOR_ICON_ENTRY}", got ${JSON.stringify(color)}`,
    );
  }
  if (outline !== APP_PACKAGE_OUTLINE_ICON_ENTRY) {
    problems.push(
      `manifest icons.outline must be "${APP_PACKAGE_OUTLINE_ICON_ENTRY}", got ${JSON.stringify(outline)}`,
    );
  }
  return problems;
}

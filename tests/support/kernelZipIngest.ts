// The kernel's plugin zip ingest, mirrored (byte5ai/omadia
// `middleware/src/plugins/zipExtractor.ts`, default plugin-package limits).
//
// Same module as in omadia-channel-teams, whose 0.36.0 shipped a file without
// an extension and was refused by every kernel with `422
// zip.forbidden_extension`: nothing there had ever pushed the real zip through
// those rules. This connector had no such check either. This module holds the
// rules so a test can.
//
// A MIRROR CAN DRIFT. `kernelIngestRules` below reads the same sets from a
// kernel checkout when one is next to this repo, and a test compares them.

import { existsSync, readFileSync } from 'node:fs';
import { posix } from 'node:path';

/** Accepted anywhere in a package. */
export const EXTENSION_ALLOWLIST: readonly string[] = [
  '.yaml', '.yml', '.md', '.json', '.js', '.mjs', '.cjs', '.map', '.png', '.svg',
  '.jpg', '.jpeg', '.txt', '.license', '.html', '.sql',
];
/** Declaration files (`.d.ts` and friends). */
export const DECL_EXTENSIONS: readonly string[] = ['.ts', '.mts', '.cts'];
/** Accepted only below a `ui/` directory. */
export const UI_BUNDLE_EXTENSIONS: readonly string[] = ['.woff2'];
/** Accepted only below an `appPackage/` directory. */
export const APP_PACKAGE_EXTENSIONS: readonly string[] = ['.template'];
/** Extension-less (or dot-only) names accepted by their basename. */
export const BASENAME_ALLOWLIST: readonly string[] = ['LICENSE', 'NOTICE', 'README', '.npmignore'];

const UI_BUNDLE_PATH = /(?:^|\/)ui\//;
const APP_PACKAGE_PATH = /(?:^|\/)appPackage\//;
const SYMLINK_MODE = 0o120000;
const FILE_TYPE_MASK = 0o170000;

export interface ZipEntry {
  name: string;
  /** Upper 16 bits of the entry's external file attributes (UNIX mode). */
  unixMode: number;
}

/** Why the kernel ingest refuses this entry, or `undefined` when it accepts it. */
export function kernelIngestRefusal(entry: ZipEntry): string | undefined {
  if ((entry.unixMode & FILE_TYPE_MASK) === SYMLINK_MODE) return 'zip.symlink';
  const name = entry.name.replace(/\\/g, '/');
  if (name.startsWith('/') || name.includes('..')) return 'zip.path_escape';
  if (name.endsWith('/')) return undefined;
  const ext = posix.extname(name).toLowerCase();
  const isAllowed =
    EXTENSION_ALLOWLIST.includes(ext) ||
    DECL_EXTENSIONS.includes(ext) ||
    (UI_BUNDLE_EXTENSIONS.includes(ext) && UI_BUNDLE_PATH.test(name)) ||
    (APP_PACKAGE_EXTENSIONS.includes(ext) && APP_PACKAGE_PATH.test(name)) ||
    BASENAME_ALLOWLIST.includes(posix.basename(name));
  return isAllowed ? undefined : `zip.forbidden_extension (${ext || '<none>'})`;
}

export interface KernelIngestRules {
  extensionAllowlist: string[];
  declExtensions: string[];
  uiBundleExtensions: string[];
  appPackageExtensions: string[];
  basenames: string[];
}

/** The rule sets as the kernel source at `zipExtractorPath` states them, or
 *  `undefined` when there is no kernel checkout at that path. */
export function kernelIngestRules(zipExtractorPath: string): KernelIngestRules | undefined {
  if (!existsSync(zipExtractorPath)) return undefined;
  const source = readFileSync(zipExtractorPath, 'utf8');
  const setOf = (constant: string): string[] => {
    const body = new RegExp(`const ${constant}\\b[^=]*=\\s*new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(source)?.[1];
    if (body === undefined) throw new Error(`${constant} not found in ${zipExtractorPath}`);
    const withoutComments = body.replace(/\/\/.*$/gm, '');
    return [...withoutComments.matchAll(/'([^']+)'/g)].flatMap((match) => (match[1] ? [match[1]] : []));
  };
  return {
    extensionAllowlist: setOf('EXTENSION_ALLOWLIST'),
    declExtensions: setOf('DECL_EXTENSIONS'),
    uiBundleExtensions: setOf('UI_BUNDLE_EXTENSIONS'),
    appPackageExtensions: setOf('APP_PACKAGE_EXTENSIONS'),
    basenames: [...source.matchAll(/baseName === '([^']+)'/g)].flatMap((match) => (match[1] ? [match[1]] : [])),
  };
}

const END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const CENTRAL_FILE_HEADER = 0x02014b50;
const END_RECORD_MIN_BYTES = 22;
const CENTRAL_HEADER_BYTES = 46;

/** Every entry of a zip archive, read from its central directory. Enough for
 *  the archives `npm run package` builds (no zip64, no multi-disk). */
export function readZipEntries(zipPath: string): ZipEntry[] {
  const zip = readFileSync(zipPath);
  let end = zip.length - END_RECORD_MIN_BYTES;
  while (end >= 0 && zip.readUInt32LE(end) !== END_OF_CENTRAL_DIRECTORY) end -= 1;
  if (end < 0) throw new Error(`${zipPath}: no end-of-central-directory record`);
  const count = zip.readUInt16LE(end + 10);
  let offset = zip.readUInt32LE(end + 16);
  const entries: ZipEntry[] = [];
  for (let i = 0; i < count; i += 1) {
    if (zip.readUInt32LE(offset) !== CENTRAL_FILE_HEADER) {
      throw new Error(`${zipPath}: broken central directory at entry ${String(i)}`);
    }
    const nameLength = zip.readUInt16LE(offset + 28);
    const extraLength = zip.readUInt16LE(offset + 30);
    const commentLength = zip.readUInt16LE(offset + 32);
    const unixMode = (zip.readUInt32LE(offset + 38) >>> 16) & 0xffff;
    const name = zip.toString('utf8', offset + CENTRAL_HEADER_BYTES, offset + CENTRAL_HEADER_BYTES + nameLength);
    entries.push({ name, unixMode });
    offset += CENTRAL_HEADER_BYTES + nameLength + extraLength + commentLength;
  }
  return entries;
}

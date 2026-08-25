/**
 * Minimal deterministic ZIP writer for the app-package step — dependency-free
 * on purpose.
 *
 * WHY THIS EXISTS instead of a zip library: the Omadia plugin host installs
 * nothing from a plugin ZIP (`build-zip.mjs` ships plain `tsc` output, no
 * `node_modules`), so a runtime `dependencies` entry can never be resolved at
 * load time — `import 'fflate'` would throw `ERR_MODULE_NOT_FOUND` in
 * production. Peers work only for packages the host itself ships. Rather than
 * bundling (which `build-zip.mjs` deliberately avoids for this connector) or
 * coupling to a host-shipped zipper, the ~120 lines of ZIP container format
 * live here: local file headers + central directory + EOCD, entries
 * deflate-compressed via `node:zlib` (always available, no dependency).
 *
 * Determinism: a fixed DOS timestamp is stamped on every entry and
 * `deflateRawSync` with a pinned level is deterministic for identical input,
 * so the same entries always produce a byte-identical archive (the
 * re-run-comparability property the app-package step relies on).
 *
 * Scope: write-only, no ZIP64 (a Teams app package is a few KB), no
 * encryption, no streaming. Tests unzip the output with `fflate` (a DEV
 * dependency — never shipped) as an independent format cross-check.
 */

import { deflateRawSync } from 'node:zlib';

/** One archive entry: name at the archive root + raw content bytes. */
export interface ZipEntry {
  readonly name: string;
  readonly data: Uint8Array;
}

const LOCAL_FILE_HEADER_SIG = 0x04034b50;
const CENTRAL_DIR_HEADER_SIG = 0x02014b50;
const END_OF_CENTRAL_DIR_SIG = 0x06054b50;
/** "version needed to extract" 2.0 — deflate support, no ZIP64 features. */
const ZIP_VERSION = 20;
const METHOD_DEFLATE = 8;
/** Pinned deflate level — part of the byte-identical-output guarantee. */
const DEFLATE_LEVEL = 9;

const CRC32_TABLE = buildCrc32Table();

function buildCrc32Table(): Uint32Array {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
}

/** Standard CRC-32 (IEEE 802.3), as required by the ZIP format. */
export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = (CRC32_TABLE[(crc ^ byte) & 0xff] ?? 0) ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** MS-DOS date+time pair (2-second resolution), the ZIP timestamp format. */
function dosDateTime(date: Date): { dosTime: number; dosDate: number } {
  const dosTime =
    (date.getUTCHours() << 11) |
    (date.getUTCMinutes() << 5) |
    Math.floor(date.getUTCSeconds() / 2);
  const dosDate =
    ((date.getUTCFullYear() - 1980) << 9) |
    ((date.getUTCMonth() + 1) << 5) |
    date.getUTCDate();
  return { dosTime, dosDate };
}

/**
 * Build a ZIP archive from `entries`, every entry stamped with `mtime`.
 * Entry order is preserved — identical input yields a byte-identical archive.
 */
export function createZip(
  entries: readonly ZipEntry[],
  mtime: Date,
): Uint8Array {
  const encoder = new TextEncoder();
  const { dosTime, dosDate } = dosDateTime(mtime);

  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    if (nameBytes.length === 0 || nameBytes.length > 0xffff) {
      throw new Error(`zip: invalid entry name '${entry.name}'`);
    }
    const checksum = crc32(entry.data);
    const compressed = deflateRawSync(entry.data, { level: DEFLATE_LEVEL });

    const local = Buffer.alloc(30);
    local.writeUInt32LE(LOCAL_FILE_HEADER_SIG, 0);
    local.writeUInt16LE(ZIP_VERSION, 4);
    local.writeUInt16LE(0, 6); // general-purpose flags
    local.writeUInt16LE(METHOD_DEFLATE, 8);
    local.writeUInt16LE(dosTime, 10);
    local.writeUInt16LE(dosDate, 12);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(entry.data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28); // extra-field length
    localParts.push(local, Buffer.from(nameBytes), compressed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(CENTRAL_DIR_HEADER_SIG, 0);
    central.writeUInt16LE(ZIP_VERSION, 4); // version made by
    central.writeUInt16LE(ZIP_VERSION, 6); // version needed
    central.writeUInt16LE(0, 8); // flags
    central.writeUInt16LE(METHOD_DEFLATE, 10);
    central.writeUInt16LE(dosTime, 12);
    central.writeUInt16LE(dosDate, 14);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(entry.data.length, 24);
    central.writeUInt16LE(nameBytes.length, 28);
    central.writeUInt16LE(0, 30); // extra-field length
    central.writeUInt16LE(0, 32); // comment length
    central.writeUInt16LE(0, 34); // disk number
    central.writeUInt16LE(0, 36); // internal attributes
    central.writeUInt32LE(0, 38); // external attributes
    central.writeUInt32LE(offset, 42); // local-header offset
    centralParts.push(central, Buffer.from(nameBytes));

    offset += 30 + nameBytes.length + compressed.length;
  }

  const centralStart = offset;
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(END_OF_CENTRAL_DIR_SIG, 0);
  eocd.writeUInt16LE(0, 4); // this disk
  eocd.writeUInt16LE(0, 6); // central-dir start disk
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(centralStart, 16);
  eocd.writeUInt16LE(0, 20); // comment length

  return Uint8Array.from(
    Buffer.concat([...localParts, ...centralParts, eocd]),
  );
}

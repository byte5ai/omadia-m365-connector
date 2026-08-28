/**
 * Directory-object lifecycle helpers for `teamsProvisioner@1` — the two Entra
 * realities the app-registration step has to live with (byte5ai/omadia#916):
 *
 * 1. **Eventual consistency.** `POST /applications` answers 201 and the very
 *    next write against that object can answer 404 `Request_ResourceNotFound`
 *    for a few seconds — the object exists, it is just not replicated to the
 *    node serving the follow-up call. {@link waitForApplicationAddressable}
 *    and {@link retryWhileReplicating} poll through that window on a bounded
 *    budget instead of reporting a step failure.
 *
 * 2. **The recycle bin.** A deleted application is SOFT-deleted and its
 *    `uniqueName` stays reserved for {@link DELETED_ITEM_RETENTION_DAYS} days
 *    while the object is invisible in `GET /applications`. That is why an
 *    operator saw "already exists" for an app they could not find anywhere.
 *    {@link findDeletedApplicationByUniqueName} and
 *    {@link restoreDeletedApplication} make that state observable and, where
 *    permissions allow, recoverable.
 *
 * Everything here goes through the shared {@link ProvisioningHttp} choke
 * point — no second token cache, no fetch of its own.
 */

import { DirectoryReplicationError } from './errors.js';
import type { ProvisioningConflictRule, ProvisioningHttp } from './http.js';

/** Microsoft Graph v1.0 base — every provisioning URL is built from it. */
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Graph signatures for "that `uniqueName` is taken". Entra does NOT answer
 * 409 here: `POST /applications` reports the duplicate as **400** with
 * `Request_BadRequest` / "Another object with the same value for property
 * uniqueName already exists." The message constraint keeps ordinary 400s
 * (a malformed displayName, a rejected audience) on the error path.
 */
export const UNIQUE_NAME_CONFLICT_RULES: readonly ProvisioningConflictRule[] = [
  { status: 400, codes: ['ObjectConflict'] },
  {
    status: 400,
    messageIncludes: [
      'same value for property uniquename',
      'uniquename already exists',
    ],
  },
];

/** Probes spent waiting for a freshly created object to become addressable. */
export const DEFAULT_REPLICATION_MAX_ATTEMPTS = 8;
/** First wait between replication probes; doubles up to {@link MAX_REPLICATION_DELAY_MS}. */
export const DEFAULT_REPLICATION_INTERVAL_MS = 1000;
/** Cap for a single replication wait — the whole budget stays well under a minute. */
const MAX_REPLICATION_DELAY_MS = 8000;

/** Injection points so the wait budget is configurable and tests never sleep. */
export interface ReplicationOptions {
  readonly maxAttempts?: number;
  readonly intervalMs?: number;
  readonly sleep?: (ms: number) => Promise<void>;
  readonly log?: (msg: string) => void;
}

interface ResolvedReplicationOptions {
  readonly maxAttempts: number;
  readonly intervalMs: number;
  readonly sleep: (ms: number) => Promise<void>;
  readonly log: (msg: string) => void;
}

const defaultSleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms));

export function resolveReplicationOptions(
  opts: ReplicationOptions | undefined,
): ResolvedReplicationOptions {
  return {
    maxAttempts: opts?.maxAttempts ?? DEFAULT_REPLICATION_MAX_ATTEMPTS,
    intervalMs: opts?.intervalMs ?? DEFAULT_REPLICATION_INTERVAL_MS,
    sleep: opts?.sleep ?? defaultSleep,
    log:
      opts?.log ??
      ((msg: string): void => {
        console.error(msg);
      }),
  };
}

/** Exponential, capped — attempt 0 waits `intervalMs`, then 2x per probe. */
function replicationDelayMs(intervalMs: number, attempt: number): number {
  return Math.min(intervalMs * 2 ** attempt, MAX_REPLICATION_DELAY_MS);
}

/**
 * Poll `GET /applications/{objectId}` until Graph admits the object exists.
 *
 * Returns as soon as the read succeeds (usually on the first probe). Throws
 * the TRANSIENT {@link DirectoryReplicationError} when the budget runs out —
 * transient because the object DOES exist; a later run finds it under its
 * `uniqueName`. Never a reason to roll anything back.
 */
export async function waitForApplicationAddressable(
  http: Pick<ProvisioningHttp, 'request'>,
  objectId: string,
  scopes: readonly string[],
  opts: ReplicationOptions | undefined,
): Promise<void> {
  const { maxAttempts, intervalMs, sleep, log } = resolveReplicationOptions(opts);
  let waitedMs = 0;

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const response = await http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/applications/${encodeURIComponent(objectId)}`,
      step: 'applications.awaitReplication',
      missingScopesOn403: scopes,
      extraOkStatuses: [404],
    });
    if (response.kind === 'ok' && response.status !== 404) return;

    const delayMs = replicationDelayMs(intervalMs, attempt);
    log(
      `provisioner applications.awaitReplication: '${objectId}' not addressable yet ` +
        `(probe ${String(attempt + 1)}/${String(maxAttempts)}), waiting ${String(delayMs)}ms`,
    );
    await sleep(delayMs);
    waitedMs += delayMs;
  }

  throw new DirectoryReplicationError(
    'applications.awaitReplication',
    objectId,
    maxAttempts,
    waitedMs,
  );
}

/**
 * Run a follow-up write against a freshly created object, retrying while it
 * still answers "not found".
 *
 * {@link waitForApplicationAddressable} closes the common window, but Graph
 * replicates READ and WRITE paths independently — a successful read is no
 * guarantee the next write lands. `op` therefore reports "still replicating"
 * by returning `undefined`, and the budget is spent here rather than being
 * paid twice.
 */
export async function retryWhileReplicating<T>(
  step: string,
  objectId: string,
  op: () => Promise<T | undefined>,
  opts: ReplicationOptions | undefined,
): Promise<T> {
  const { maxAttempts, intervalMs, sleep, log } = resolveReplicationOptions(opts);
  let waitedMs = 0;

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const result = await op();
    if (result !== undefined) return result;

    const delayMs = replicationDelayMs(intervalMs, attempt);
    log(
      `provisioner ${step}: object '${objectId}' still replicating ` +
        `(attempt ${String(attempt + 1)}/${String(maxAttempts)}), waiting ${String(delayMs)}ms`,
    );
    await sleep(delayMs);
    waitedMs += delayMs;
  }

  throw new DirectoryReplicationError(step, objectId, maxAttempts, waitedMs);
}

/** A soft-deleted application as the recycle bin reports it. */
export interface DeletedApplication {
  /** Directory object id — what `restore` and `purge` address. */
  readonly objectId: string;
  readonly uniqueName: string;
  readonly appId?: string;
  readonly displayName?: string;
  /** ISO-8601 timestamp of the delete, when Graph reported one. */
  readonly deletedDateTime?: string;
}

/** Recycle-bin pages to scan before giving up on the lookup. */
const MAX_DELETED_ITEM_PAGES = 5;

/**
 * Find a soft-deleted application by its `uniqueName`.
 *
 * Deliberately a client-side scan of `GET
 * /directory/deletedItems/microsoft.graph.application` rather than an
 * `$filter`: `uniqueName` is not reliably filterable there, and a rejected
 * filter would turn a diagnostic into a second failure.
 *
 * BEST EFFORT — this endpoint is not covered by
 * `Application.ReadWrite.OwnedBy`, so a tenant that granted only the
 * provisioner's own scope gets a 403 here. That must never escalate: every
 * failure answers `undefined`, and the caller degrades to an explanatory
 * error instead of a consent error.
 */
export async function findDeletedApplicationByUniqueName(
  http: Pick<ProvisioningHttp, 'request'>,
  uniqueName: string,
  scopes: readonly string[],
  log: (msg: string) => void,
): Promise<DeletedApplication | undefined> {
  let url =
    `${GRAPH_BASE}/directory/deletedItems/microsoft.graph.application` +
    `?$select=id,appId,displayName,uniqueName,deletedDateTime&$top=100`;

  try {
    for (let page = 0; page < MAX_DELETED_ITEM_PAGES; page += 1) {
      const response = await http.request({
        resource: 'graph',
        method: 'GET',
        url,
        step: 'directory.deletedItems.list',
        missingScopesOn403: scopes,
      });
      if (response.kind !== 'ok') return undefined;
      const body = response.json;
      if (!body || typeof body !== 'object') return undefined;
      const record = body as Record<string, unknown>;

      const items = Array.isArray(record['value']) ? record['value'] : [];
      for (const item of items) {
        const parsed = parseDeletedApplication(item);
        if (parsed?.uniqueName === uniqueName) return parsed;
      }

      const nextLink = record['@odata.nextLink'];
      if (typeof nextLink !== 'string' || nextLink.length === 0) return undefined;
      url = nextLink;
    }
    return undefined;
  } catch (err) {
    // A probe is a diagnostic, never a failure mode of its own.
    log(
      `provisioner directory.deletedItems.list: recycle-bin probe for ` +
        `'${uniqueName}' unavailable (${String(err)})`,
    );
    return undefined;
  }
}

function parseDeletedApplication(item: unknown): DeletedApplication | undefined {
  if (!item || typeof item !== 'object') return undefined;
  const record = item as Record<string, unknown>;
  const objectId = record['id'];
  const uniqueName = record['uniqueName'];
  if (typeof objectId !== 'string' || typeof uniqueName !== 'string') {
    return undefined;
  }
  const appId = record['appId'];
  const displayName = record['displayName'];
  const deletedDateTime = record['deletedDateTime'];
  return {
    objectId,
    uniqueName,
    ...(typeof appId === 'string' ? { appId } : {}),
    ...(typeof displayName === 'string' ? { displayName } : {}),
    ...(typeof deletedDateTime === 'string' ? { deletedDateTime } : {}),
  };
}

/**
 * Bring a soft-deleted application back out of the recycle bin.
 *
 * `POST /directory/deletedItems/{id}/restore` is the ONLY way to free a
 * reserved `uniqueName` before the retention window ends — a purge keeps the
 * name reserved too. Restoring returns the ORIGINAL object (same appId, same
 * uniqueName), which is exactly what re-provisioning the same agent slug is
 * supposed to yield.
 *
 * Throws on failure; the caller turns that into the explanatory
 * `UniqueNameReservedError`.
 */
export async function restoreDeletedApplication(
  http: Pick<ProvisioningHttp, 'request'>,
  objectId: string,
  scopes: readonly string[],
): Promise<void> {
  const response = await http.request({
    resource: 'graph',
    method: 'POST',
    url: `${GRAPH_BASE}/directory/deletedItems/${encodeURIComponent(objectId)}/restore`,
    step: 'directory.deletedItems.restore',
    missingScopesOn403: scopes,
  });
  if (response.kind !== 'ok') {
    throw new Error(
      `graph directory.deletedItems.restore unexpected conflict for '${objectId}'`,
    );
  }
}

/**
 * Catalog-upload step of `teamsProvisioner@1` (epic byte5ai/omadia#860,
 * capability issue byte5ai/omadia-m365-connector#3): publish a zipped Teams
 * app package into the TENANT app catalog — the link between "1 Entra app +
 * 1 Azure bot" and the team install that follows.
 *
 * One Graph write: `POST /appCatalogs/teamsApps` with the raw zip body. The
 * zip comes from the buildAppPackage step, but this module takes it as a
 * plain `Uint8Array` parameter — no compile/test dependency on the packaging
 * unit; any buffer works.
 *
 * IDEMPOTENCY — 409 is success, but NOT a bare swallow. Graph answers 409
 * when an app with the same manifest id (`externalId`) is already published.
 * The shared http layer maps that to its `{ kind: 'conflict' }` signal and
 * this module RE-RESOLVES the existing catalog entry by `externalId`
 * (`GET /appCatalogs/teamsApps?$filter=externalId eq '…'`) so the
 * `'already-existed'` outcome carries the SAME `CatalogTeamsApp` shape as a
 * fresh upload — a second provisioning run is a true no-op for the caller.
 *
 * All HTTP goes through the shared {@link ProvisioningHttp} choke point (one
 * token cache, Retry-After-honouring 429 backoff → `ProvisioningThrottledError`
 * when exhausted, 403 → `ConsentMissingError` carrying
 * {@link APP_CATALOG_SCOPE}) — the typed 403 is load-bearing: the middleware
 * agent factory (byte5ai/omadia#863-865) branches on it to fall back. This
 * module opens no second token cache and does no fetch of its own.
 */

import type { ProvisioningHttp } from './http.js';
import type {
  CatalogTeamsApp,
  Idempotent,
  UploadToCatalogInput,
} from './types.js';

/**
 * Graph APPLICATION permission this step needs. Documented in the
 * scopes/consent unit (`INTEGRATION.md`, `docs/teams-provisioner.md`) —
 * surfaced on 403 via `ConsentMissingError.missingScopes`.
 */
export const APP_CATALOG_SCOPE = 'AppCatalog.ReadWrite.All';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const ZIP_CONTENT_TYPE = 'application/zip';

export interface CatalogUploadClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  readonly log?: (msg: string) => void;
}

/**
 * Publish Teams app packages into the tenant app catalog. One step client
 * per provisioner; ordering across chain steps stays middleware-side (agent
 * factory, byte5ai/omadia#863-865).
 */
export class CatalogUploadClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly log: (msg: string) => void;

  constructor(opts: CatalogUploadClientOptions) {
    this.http = opts.http;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * `POST /appCatalogs/teamsApps` — publish the zipped app package into the
   * tenant catalog.
   *
   * - 2xx → `'created'`. Graph's POST response omits the appDefinitions
   *   (and with them the manifest version), so the entry is re-read by
   *   `externalId` whenever the response body alone cannot fill the
   *   {@link CatalogTeamsApp} shape.
   * - 409 → `'already-existed'` (same `externalId` already published) —
   *   success, never an exception. The existing entry is re-resolved by
   *   `externalId` so both outcomes carry an identical value shape.
   * - 403 → `ConsentMissingError([APP_CATALOG_SCOPE], 'graph')` from the
   *   http layer, so the agent factory gets ONE typed fallback branch.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   */
  async uploadToCatalog(
    input: UploadToCatalogInput,
  ): Promise<Idempotent<CatalogTeamsApp>> {
    const externalId = requireNonEmpty(input.externalId, 'externalId');
    if (
      !(input.packageZip instanceof Uint8Array) ||
      input.packageZip.byteLength === 0
    ) {
      throw new Error(
        "invalid_argument: 'packageZip' must be a non-empty Uint8Array",
      );
    }

    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/appCatalogs/teamsApps`,
      step: 'appCatalogs.teamsApps.publish',
      rawBody: { bytes: input.packageZip, contentType: ZIP_CONTENT_TYPE },
      missingScopesOn403: [APP_CATALOG_SCOPE],
    });

    if (response.kind === 'conflict') {
      this.log(
        `provisioner appCatalogs.teamsApps.publish: externalId=${externalId} already in catalog (409 → already-existed)`,
      );
      return {
        outcome: 'already-existed',
        value: await this.resolveByExternalId(externalId),
      };
    }

    const uploaded = catalogApp(response.json);
    if (uploaded !== undefined) {
      if (uploaded.externalId !== externalId) {
        throw new Error(
          `graph appCatalogs.teamsApps.publish returned externalId=${uploaded.externalId} but the input declared externalId=${externalId} — package/manifest mismatch`,
        );
      }
      return { outcome: 'created', value: uploaded };
    }

    return {
      outcome: 'created',
      value: await this.resolveByExternalId(externalId),
    };
  }

  /**
   * `GET /appCatalogs/teamsApps?$filter=externalId eq '…'` (expanding
   * `appDefinitions` for the manifest version) — the shared lookup for the
   * 409 idempotent path and for POST responses too thin to build the result
   * from. A conflict entry that cannot be found afterwards is an
   * inconsistency worth failing loudly on, not an outcome.
   */
  private async resolveByExternalId(
    externalId: string,
  ): Promise<CatalogTeamsApp> {
    const filter = `externalId eq '${escapeODataString(externalId)}'`;
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/appCatalogs/teamsApps?$filter=${encodeURIComponent(filter)}&$expand=appDefinitions($select=version)`,
      step: 'appCatalogs.teamsApps.lookup',
      missingScopesOn403: [APP_CATALOG_SCOPE],
    });

    if (response.kind !== 'ok') {
      throw new Error(
        `graph appCatalogs.teamsApps.lookup unexpectedly answered ${String(response.status)} for externalId=${externalId}`,
      );
    }

    const match = listEntries(response.json)
      .map(catalogApp)
      .find((app) => app !== undefined && app.externalId === externalId);
    if (match === undefined) {
      throw new Error(
        `graph appCatalogs.teamsApps.lookup found no catalog app with externalId=${externalId}`,
      );
    }
    return match;
  }
}

/** OData string literal escaping: single quotes double up. */
function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

/** `value` array of a Graph collection response. */
function listEntries(json: unknown): unknown[] {
  if (!json || typeof json !== 'object') return [];
  const value = (json as Record<string, unknown>)['value'];
  return Array.isArray(value) ? value : [];
}

/**
 * Parse one Graph `teamsApp` into {@link CatalogTeamsApp} — `undefined` when
 * any required field (incl. the appDefinitions version) is missing, which
 * routes the caller to the externalId lookup instead.
 */
function catalogApp(json: unknown): CatalogTeamsApp | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const rec = json as Record<string, unknown>;
  const teamsAppId = nonEmptyString(rec['id']);
  const externalId = nonEmptyString(rec['externalId']);
  const displayName = nonEmptyString(rec['displayName']);
  const version = appVersion(rec);
  if (!teamsAppId || !externalId || !displayName || !version) return undefined;
  return { teamsAppId, externalId, displayName, version };
}

/** Manifest version — from the expanded `appDefinitions`, else top-level. */
function appVersion(rec: Record<string, unknown>): string | undefined {
  const definitions = rec['appDefinitions'];
  if (Array.isArray(definitions)) {
    for (const definition of definitions) {
      if (definition && typeof definition === 'object') {
        const version = nonEmptyString(
          (definition as Record<string, unknown>)['version'],
        );
        if (version !== undefined) return version;
      }
    }
  }
  return nonEmptyString(rec['version']);
}

function nonEmptyString(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

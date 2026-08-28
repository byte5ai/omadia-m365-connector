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
 * THE UPLOAD IS THE ONE DELEGATED STEP OF THE WHOLE CHAIN (byte5ai/omadia#924).
 * Graph documents application permissions for `POST /appCatalogs/teamsApps` as
 * "Not supported.", and the field test matches: with a single app-only token
 * the catalog LOOKUP below succeeds while the UPLOAD is rejected, although
 * `AppCatalog.ReadWrite.All` is assigned as an app role and consented. No
 * consent fixes it, because it is not a consent problem. So the upload carries
 * a delegated (user) token — acquired once per tenant through the device-code
 * flow in `delegatedAuth.ts` — and EVERYTHING ELSE in this module, including
 * the `externalId` lookup on the idempotent 409 path, stays app-only.
 *
 * The asymmetry is deliberate and worth keeping: it means a stale or missing
 * user token degrades exactly one operation instead of blinding the provisioner
 * to what is already in the catalog.
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
 * when exhausted, 403 → `ConsentMissingError`). The typed 403 is load-bearing:
 * the middleware agent factory (byte5ai/omadia#863-865) branches on it to fall
 * back. This module opens no second token cache and does no fetch of its own.
 *
 * The DELEGATED publish translates that 403 one step further, into
 * `DelegatedConsentRequiredError` — on that path an app role was never
 * involved, so reporting a missing APPLICATION permission would point the
 * operator at the wrong registration entirely.
 */

import {
  APP_CATALOG_DELEGATED_SCOPE,
  type DelegatedTokenSet,
} from './delegatedAuth.js';
import {
  DelegatedConsentRequiredError,
  DelegatedSignInRequiredError,
  DelegatedTokenExpiredError,
  ProvisioningRequestError,
  ConsentMissingError,
} from './errors.js';
import type { ProvisioningHttp, ProvisioningResponse } from './http.js';
import type {
  CatalogTeamsApp,
  Idempotent,
  UploadToCatalogInput,
} from './types.js';

/** Input for the catalog-lookup step ({@link CatalogUploadClient.getCatalogApp}). */
export interface GetCatalogAppInput {
  /** Manifest id (`externalId`) of the catalog app to resolve. */
  readonly teamsAppExternalId: string;
}

/** Lookup miss — no catalog app carries the requested `externalId`. */
export interface CatalogAppNotFound {
  readonly found: false;
}

/**
 * Lookup hit. `displayName` / `publishedVersion` are optional on purpose:
 * a catalog entry can exist (and be installable by `teamsAppId`) while Graph
 * omits either field, so the lookup never turns a thin-but-real entry into a
 * miss the way the strict {@link CatalogTeamsApp} parse would.
 */
export interface CatalogAppFound {
  readonly found: true;
  /** Catalog id (`teamsApp.id`) — what installs reference. */
  readonly teamsAppId: string;
  readonly displayName?: string;
  /**
   * Manifest version, selected like the upload path: the CURRENT
   * (`publishingState === 'published'`) appDefinition wins, else the highest
   * version (numeric-aware dotted compare).
   */
  readonly publishedVersion?: string;
}

/** Result of {@link CatalogUploadClient.getCatalogApp}. */
export type GetCatalogAppResult = CatalogAppNotFound | CatalogAppFound;

/**
 * Graph APPLICATION permission this step needs. Documented in the
 * scopes/consent unit (`INTEGRATION.md`, `docs/teams-provisioner.md`) —
 * surfaced on 403 via `ConsentMissingError.missingScopes`.
 */
export const APP_CATALOG_SCOPE = 'AppCatalog.ReadWrite.All';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const ZIP_CONTENT_TYPE = 'application/zip';
/** Step label of the delegated publish — reused across its typed errors. */
const UPLOAD_STEP = 'appCatalogs.teamsApps.publish';
/**
 * Fallback for DelegatedConsentRequiredError.adminConsentUrl when this client
 * was built without a delegated-auth client (only reachable via the low-level
 * uploadToCatalog seam). A sentence beats an empty string: the caller still has
 * to show an operator SOMETHING actionable.
 */
const UNKNOWN_CONSENT_URL_HINT =
  '(resolve the publisher app first — call startDelegatedSignIn to obtain the admin-consent URL)';

/** Result of the token-carrying upload — the credential may have rotated. */
export interface DelegatedCatalogUploadResult {
  /** Same idempotent outcome as the app-only surface used to produce. */
  readonly app: Idempotent<CatalogTeamsApp>;
  /**
   * The token set to persist. `refreshed` says whether it CHANGED — a caller
   * that ignores it and never writes back will eventually force a needless
   * second admin sign-in, because Entra rotates the refresh token.
   */
  readonly tokens: DelegatedTokenSet;
  readonly refreshed: boolean;
}

/**
 * What the catalog client needs from the delegated-auth layer.
 *
 * A narrow port rather than the `DelegatedAuthClient` class: the publisher app's
 * client id is only known once a token set exists, so both members take the
 * tokens as an argument instead of being bound at construction time. That keeps
 * the provisioner constructible without a Graph round trip — activation must
 * stay side-effect free.
 */
export interface DelegatedUploadAuthority {
  /** Refresh a stale access token; `refreshed` says whether it changed. */
  ensureFreshToken(tokens: DelegatedTokenSet): Promise<{
    readonly tokens: DelegatedTokenSet;
    readonly refreshed: boolean;
  }>;
  /**
   * Admin-consent URL for the publisher app these tokens belong to. Called on
   * the failure path only; `undefined` tokens (the low-level seam) yield an
   * actionable sentence instead.
   */
  adminConsentUrlFor(tokens: DelegatedTokenSet | undefined): string;
}

/** Input for {@link CatalogUploadClient.uploadToCatalogDelegated}. */
export interface UploadToCatalogDelegatedInput {
  /** The zipped Teams app package. */
  readonly packageZip: Uint8Array;
  /** Manifest id — used for the pre-flight `externalId` lookup on 409. */
  readonly externalId: string;
  /** SECRET. The stored delegated credential, refreshed here when stale. */
  readonly tokens: DelegatedTokenSet;
}

export interface CatalogUploadClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  /**
   * Delegated-auth client for the upload's user token. Optional so the module
   * stays constructible (and testable) without one; `uploadToCatalogDelegated`
   * then reports the missing sign-in the same way a missing token does.
   */
  readonly delegatedAuth?: DelegatedUploadAuthority;
  readonly log?: (msg: string) => void;
}

/**
 * Publish Teams app packages into the tenant app catalog. One step client
 * per provisioner; ordering across chain steps stays middleware-side (agent
 * factory, byte5ai/omadia#863-865).
 */
export class CatalogUploadClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly delegatedAuth?: DelegatedUploadAuthority;
  private readonly log: (msg: string) => void;

  constructor(opts: CatalogUploadClientOptions) {
    this.http = opts.http;
    if (opts.delegatedAuth !== undefined) this.delegatedAuth = opts.delegatedAuth;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * The upload as a caller holding a STORED credential wants it: refresh the
   * access token when it is stale, publish, and hand the (possibly rotated)
   * token set back for persistence.
   *
   * This is the method the middleware should call. {@link uploadToCatalog}
   * remains the lower-level seam for a caller that already has a fresh access
   * token in hand.
   *
   * @throws {DelegatedTokenExpiredError} `'refresh-token-invalid'` — the stored
   *   credential is dead; an admin must sign in again.
   * @throws {DelegatedConsentRequiredError} tenant consent is missing or was
   *   withdrawn; the error carries the admin-consent URL.
   */
  async uploadToCatalogDelegated(
    input: UploadToCatalogDelegatedInput,
  ): Promise<DelegatedCatalogUploadResult> {
    if (this.delegatedAuth === undefined) {
      throw new DelegatedSignInRequiredError(UPLOAD_STEP, [
        APP_CATALOG_DELEGATED_SCOPE,
      ]);
    }
    const { tokens, refreshed } = await this.delegatedAuth.ensureFreshToken(
      input.tokens,
    );
    const app = await this.publishAndResolve(
      input.packageZip,
      tokens.accessToken,
      requireNonEmpty(input.externalId, 'externalId'),
      tokens,
    );
    return { app, tokens, refreshed };
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
   * - 401 → `DelegatedTokenExpiredError('access-token-expired')` — the user
   *   token aged out mid-flight; refresh and retry.
   * - 403 → `DelegatedConsentRequiredError` carrying the admin-consent URL.
   *   The app-only `ConsentMissingError` cannot apply here: the call did not
   *   use an app identity, so telling an operator to grant an app role would
   *   send them to fix something that was never the problem.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   *
   * @throws {DelegatedSignInRequiredError} when no delegated token was passed.
   *   Distinct from every other failure on purpose: it means "nobody has signed
   *   in yet", whose remedy is a device-code flow, not a retry and not a
   *   permission change.
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
    const delegatedAccessToken = input.delegatedAccessToken;
    if (
      typeof delegatedAccessToken !== 'string' ||
      delegatedAccessToken.length === 0
    ) {
      // Refusing BEFORE the request, not after Graph's 403: an app-only upload
      // is known to be unsupported, and letting it go out would spend a Graph
      // call to produce a misleading "consent missing" answer.
      throw new DelegatedSignInRequiredError(UPLOAD_STEP, [
        APP_CATALOG_DELEGATED_SCOPE,
      ]);
    }

    return this.publishAndResolve(
      input.packageZip,
      delegatedAccessToken,
      externalId,
      undefined,
    );
  }

  /**
   * Publish, then normalise both outcomes onto the identical
   * {@link CatalogTeamsApp} shape — shared by the low-level and the
   * token-carrying entry points so the 409 semantics can never drift apart.
   */
  private async publishAndResolve(
    packageZip: Uint8Array,
    delegatedAccessToken: string,
    externalId: string,
    tokens: DelegatedTokenSet | undefined,
  ): Promise<Idempotent<CatalogTeamsApp>> {
    if (!(packageZip instanceof Uint8Array) || packageZip.byteLength === 0) {
      throw new Error(
        "invalid_argument: 'packageZip' must be a non-empty Uint8Array",
      );
    }

    const response = await this.publish(
      packageZip,
      delegatedAccessToken,
      externalId,
      tokens,
    );

    if (response.kind === 'conflict') {
      this.log(
        `provisioner ${UPLOAD_STEP}: externalId=${externalId} already in catalog (409 → already-existed)`,
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
   * The one delegated Graph write, with its 401/403 answers translated into
   * the delegated taxonomy.
   *
   * The translation is the point. Left alone, the shared http layer maps a 403
   * to `ConsentMissingError` naming an APPLICATION permission — correct for
   * every other step, actively misleading for this one, where the fix is a
   * tenant grant on the publisher app and not an app role on the connector.
   * A 401 likewise arrives as a generic request error when it actually means
   * "your user token aged out; refresh it".
   */
  private async publish(
    packageZip: Uint8Array,
    delegatedAccessToken: string,
    externalId: string,
    tokens: DelegatedTokenSet | undefined,
  ): Promise<ProvisioningResponse> {
    try {
      return await this.http.request({
        resource: 'graph',
        method: 'POST',
        url: `${GRAPH_BASE}/appCatalogs/teamsApps`,
        step: UPLOAD_STEP,
        rawBody: { bytes: packageZip, contentType: ZIP_CONTENT_TYPE },
        // SECRET — forwarded, never cached, never logged by the http layer.
        bearerToken: delegatedAccessToken,
        missingScopesOn403: [APP_CATALOG_DELEGATED_SCOPE],
      });
    } catch (err) {
      if (err instanceof ConsentMissingError) {
        throw new DelegatedConsentRequiredError(
          UPLOAD_STEP,
          [APP_CATALOG_DELEGATED_SCOPE],
          this.delegatedAuth?.adminConsentUrlFor(tokens) ??
            UNKNOWN_CONSENT_URL_HINT,
          err,
        );
      }
      if (err instanceof ProvisioningRequestError && err.status === 401) {
        throw new DelegatedTokenExpiredError('access-token-expired', err);
      }
      this.log(
        `provisioner ${UPLOAD_STEP}: publishing externalId=${externalId} failed`,
      );
      throw err;
    }
  }

  /**
   * Resolve an EXISTING catalog app by its manifest id (`externalId`) —
   * `GET /appCatalogs/teamsApps?$filter=externalId eq '…'`, the same query
   * (and the same `$expand=appDefinitions` version selection) the 409
   * idempotent upload path uses, but WITHOUT uploading anything: consumers
   * that only need the `teamsAppId` of an already-published app (e.g. to
   * install it into a team) call this instead of round-tripping a package.
   *
   * - hit → `{ found: true, teamsAppId, displayName?, publishedVersion? }`
   * - miss → `{ found: false }` — a plain outcome, never an exception
   *   (unlike the 409 path, where a vanished entry is an inconsistency).
   * - 403 → `ConsentMissingError([APP_CATALOG_SCOPE], 'graph')`;
   *   429 → shared Retry-After backoff → `ProvisioningThrottledError`.
   */
  async getCatalogApp(input: GetCatalogAppInput): Promise<GetCatalogAppResult> {
    const externalId = requireNonEmpty(
      input.teamsAppExternalId,
      'teamsAppExternalId',
    );
    for (const entry of await this.queryByExternalId(externalId)) {
      const match = foundCatalogApp(entry, externalId);
      if (match !== undefined) return match;
    }
    this.log(
      `provisioner appCatalogs.teamsApps.lookup: externalId=${externalId} not in catalog (found=false)`,
    );
    return { found: false };
  }

  /**
   * The shared lookup for the 409 idempotent path and for POST responses too
   * thin to build the result from. A conflict entry that cannot be found
   * afterwards is an inconsistency worth failing loudly on, not an outcome.
   */
  private async resolveByExternalId(
    externalId: string,
  ): Promise<CatalogTeamsApp> {
    const match = (await this.queryByExternalId(externalId))
      .map(catalogApp)
      .find((app) => app !== undefined && app.externalId === externalId);
    if (match === undefined) {
      throw new Error(
        `graph appCatalogs.teamsApps.lookup found no catalog app with externalId=${externalId}`,
      );
    }
    return match;
  }

  /**
   * `GET /appCatalogs/teamsApps?$filter=externalId eq '…'` (expanding
   * `appDefinitions` for the manifest version) — the ONE catalog query behind
   * both {@link getCatalogApp} and the upload path's re-resolution. OData
   * string literal escaping (quote doubling) + `encodeURIComponent` keep the
   * `$filter` injection-safe for arbitrary externalIds.
   */
  private async queryByExternalId(externalId: string): Promise<unknown[]> {
    const filter = `externalId eq '${escapeODataString(externalId)}'`;
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/appCatalogs/teamsApps?$filter=${encodeURIComponent(filter)}&$expand=appDefinitions($select=version,publishingState)`,
      step: 'appCatalogs.teamsApps.lookup',
      missingScopesOn403: [APP_CATALOG_SCOPE],
    });

    if (response.kind !== 'ok') {
      throw new Error(
        `graph appCatalogs.teamsApps.lookup unexpectedly answered ${String(response.status)} for externalId=${externalId}`,
      );
    }
    return listEntries(response.json);
  }
}

/** OData string literal escaping: single quotes double up. */
function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Parse one Graph `teamsApp` into a {@link CatalogAppFound} — lenient on
 * purpose (only `id` + a matching `externalId` are required); `undefined`
 * when the entry is not the requested app.
 */
function foundCatalogApp(
  json: unknown,
  externalId: string,
): CatalogAppFound | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const rec = json as Record<string, unknown>;
  if (nonEmptyString(rec['externalId']) !== externalId) return undefined;
  const teamsAppId = nonEmptyString(rec['id']);
  if (teamsAppId === undefined) return undefined;
  const displayName = nonEmptyString(rec['displayName']);
  const publishedVersion = appVersion(rec);
  return {
    found: true,
    teamsAppId,
    ...(displayName !== undefined ? { displayName } : {}),
    ...(publishedVersion !== undefined ? { publishedVersion } : {}),
  };
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

/**
 * Manifest version — from the expanded `appDefinitions`, else top-level.
 *
 * `appDefinitions` holds one entry PER PUBLISHED VERSION of the app, in an
 * order Graph does not guarantee — taking "the first entry with a version"
 * would silently report a stale/arbitrary version on the 409 idempotent
 * path. Selection is therefore deterministic: the CURRENT (`publishingState
 * === 'published'`) definition wins; when none is marked published, the
 * highest version (numeric-aware dotted compare) is used.
 */
function appVersion(rec: Record<string, unknown>): string | undefined {
  const definitions = rec['appDefinitions'];
  if (Array.isArray(definitions)) {
    const candidates: { version: string; publishingState: string | undefined }[] =
      [];
    for (const definition of definitions) {
      if (definition === null || typeof definition !== 'object') continue;
      const entry = definition as Record<string, unknown>;
      const version = nonEmptyString(entry['version']);
      if (version === undefined) continue;
      candidates.push({
        version,
        publishingState: nonEmptyString(entry['publishingState']),
      });
    }
    const published = candidates.find(
      (candidate) => candidate.publishingState === 'published',
    );
    if (published !== undefined) return published.version;
    const highest = [...candidates]
      .sort((a, b) => compareDottedVersions(a.version, b.version))
      .pop();
    if (highest !== undefined) return highest.version;
  }
  return nonEmptyString(rec['version']);
}

/** Numeric-aware dotted-version compare (`1.10.0` > `1.9.9`); ascending. */
function compareDottedVersions(a: string, b: string): number {
  const partsA = a.split('.');
  const partsB = b.split('.');
  const length = Math.max(partsA.length, partsB.length);
  for (let i = 0; i < length; i += 1) {
    const numA = Number.parseInt(partsA[i] ?? '0', 10) || 0;
    const numB = Number.parseInt(partsB[i] ?? '0', 10) || 0;
    if (numA !== numB) return numA - numB;
  }
  return a.localeCompare(b);
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

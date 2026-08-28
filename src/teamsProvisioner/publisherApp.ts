/**
 * The PUBLISHER app — the minimal public-client Entra registration the
 * delegated catalog upload signs in as (byte5ai/omadia#924).
 *
 * WHY A SECOND APP AT ALL. The device authorization grant only works for a
 * PUBLIC client, i.e. an app Entra will issue tokens to without a client
 * secret. The connector's own app registration must never become one: it holds
 * `Application.ReadWrite.OwnedBy`, so it can mint app registrations and client
 * secrets across the tenant. An identity with that much reach and no secret
 * required to speak as it is not a trade worth making for one upload. So the
 * public-client flag goes on a SEPARATE app that can do exactly one thing.
 *
 * WHAT IT IS ALLOWED TO DO. One delegated scope —
 * `AppCatalog.ReadWrite.All` — and nothing else. No client secret is ever
 * created for it, no app role is ever assigned to it, and it therefore cannot
 * act on its own: it only ever borrows the rights of the admin who signed in.
 * That is what makes handing it a public-client flag acceptable.
 *
 * WHO CREATES IT. The connector, through Graph, with the permission it already
 * has (`Application.ReadWrite.OwnedBy` covers apps it owns). No operator has to
 * register anything by hand — the first sign-in provisions it.
 *
 * IDEMPOTENCY reuses the app-registration step's hard-won machinery rather than
 * re-deriving it (byte5ai/omadia#916): the qualified `uniqueName` is the
 * natural key, a taken name is ADOPTED rather than treated as a failure, Entra's
 * replication windows are polled through instead of reported, and a name held
 * by a soft-deleted app is recovered by restoring that app.
 *
 * AND THERE IS NO ROLLBACK. Deliberately, not by omission. A delete
 * soft-deletes the object and reserves its `uniqueName` for 30 days, so a
 * rollback on a half-created publisher app would make the tenant unable to sign
 * in for a month over a transient error. There is also nothing worth undoing:
 * the app holds no secret and grants nobody anything until an admin consents.
 * A half-created registration is simply adopted and completed on the next run.
 */

import { APP_REGISTRATION_SCOPE } from './appRegistration.js';
import {
  asRecord,
  escapeODataQuotes,
  requireNonEmpty,
  requireUniqueName,
  requireStringField,
} from './appRegistrationSupport.js';
import {
  APP_CATALOG_DELEGATED_PERMISSION_ID,
  GRAPH_RESOURCE_APP_ID,
} from './delegatedAuth.js';
import {
  GRAPH_BASE,
  UNIQUE_NAME_CONFLICT_RULES,
  findDeletedApplicationByUniqueName,
  restoreDeletedApplication,
  retryWhileReplicating,
  waitForApplicationAddressable,
  type ReplicationOptions,
} from './directory.js';
import {
  DELETED_ITEM_RETENTION_DAYS,
  UniqueNameReservedError,
} from './errors.js';
import type { ProvisioningHttp } from './http.js';
import { redactUnknown } from './redact.js';
import { SINGLE_TENANT_SIGN_IN_AUDIENCE, type Idempotent } from './types.js';

/** `uniqueName` prefix of the per-tenant publisher app. */
export const PUBLISHER_APP_UNIQUE_NAME_PREFIX = 'omadia-teams-publisher';
/** Default portal display name — what an admin sees on the consent screen. */
export const PUBLISHER_APP_DISPLAY_NAME = 'omadia Teams App Publisher';
/** Graph's limit for `application.uniqueName`. */
const MAX_UNIQUE_NAME_LENGTH = 64;

/**
 * The publisher app's idempotency key, qualified by tenant.
 *
 * The app lives inside the tenant, so the qualifier is redundant for
 * uniqueness — it is there to make the object self-describing in a directory
 * an operator is reading by hand, and to keep the key stable across
 * reinstalls (the same tenant always resolves to the same app instead of
 * accumulating one per install).
 */
export function publisherAppUniqueName(tenantId: string): string {
  requireNonEmpty(tenantId, 'tenantId');
  const candidate = `${PUBLISHER_APP_UNIQUE_NAME_PREFIX}-${tenantId}`;
  if (candidate.length > MAX_UNIQUE_NAME_LENGTH) {
    throw new Error(
      `invalid_argument: derived uniqueName '${candidate}' exceeds Graph's ` +
        `${String(MAX_UNIQUE_NAME_LENGTH)}-character limit for application.uniqueName`,
    );
  }
  return requireUniqueName(candidate);
}

/** The provisioned (or adopted) publisher app registration. */
export interface PublisherApp {
  /** Application (client) id — the `client_id` of the device-code flow. */
  readonly appId: string;
  /** Directory object id (what PATCH addresses). */
  readonly objectId: string;
  /** Tenant the app is registered in. */
  readonly tenantId: string;
  readonly displayName: string;
  /** The idempotency key it was created/found under. */
  readonly uniqueName: string;
  /** `true` once the app is flagged as a public client (device code works). */
  readonly isPublicClient: boolean;
  /**
   * `true` when the delegated `AppCatalog.ReadWrite.All` scope is declared on
   * the registration. Declaring is NOT consenting — the tenant grant happens
   * at sign-in (or through the admin-consent URL).
   */
  readonly declaresCatalogScope: boolean;
}

export interface PublisherAppClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — never open a second
   * token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  /** Tenant the publisher app is registered in. */
  readonly tenantId: string;
  readonly log?: (msg: string) => void;
  /** Entra replication budget + the test seam for its waits. */
  readonly replication?: ReplicationOptions;
}

/** The Entra payload that makes an app usable for the device code grant. */
function publisherAppBody(displayName: string, uniqueName: string): Record<string, unknown> {
  return {
    displayName,
    // Same invariant as every other app this provisioner creates.
    signInAudience: SINGLE_TENANT_SIGN_IN_AUDIENCE,
    uniqueName,
    // "Allow public client flows" in the portal. Without it the device code
    // grant answers `unauthorized_client` and no amount of consent helps.
    isFallbackPublicClient: true,
    // An explicitly EMPTY redirect list is the point of this whole design:
    // the device code grant needs none, so a self-hosted omadia never has to
    // register its deployment URL anywhere.
    publicClient: { redirectUris: [] },
    requiredResourceAccess: [catalogResourceAccess()],
  };
}

/** The single delegated permission the publisher app declares. */
function catalogResourceAccess(): Record<string, unknown> {
  return {
    resourceAppId: GRAPH_RESOURCE_APP_ID,
    resourceAccess: [
      // `type: 'Scope'` = DELEGATED. `'Role'` would be an app permission,
      // which is precisely what Graph refuses for the catalog upload.
      { id: APP_CATALOG_DELEGATED_PERMISSION_ID, type: 'Scope' },
    ],
  };
}

/**
 * Create, adopt and repair the per-tenant publisher app. One instance per
 * provisioner; every call is safe to repeat.
 */
export class PublisherAppClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly tenantId: string;
  private readonly log: (msg: string) => void;
  private readonly replication: ReplicationOptions;

  constructor(opts: PublisherAppClientOptions) {
    this.http = opts.http;
    this.tenantId = requireNonEmpty(opts.tenantId, 'tenantId');
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
    this.replication = {
      ...opts.replication,
      log: (msg) => {
        this.log(msg);
      },
    };
  }

  /**
   * Ensure the tenant has a usable publisher app, and return it.
   *
   * `'created'` on the first run, `'already-existed'` afterwards. An adopted
   * app is RECONCILED rather than trusted: an app registered by an older
   * connector version (or edited in the portal) may be missing the
   * public-client flag or the delegated scope, and both are silent failures at
   * sign-in time — so both are repaired here instead.
   */
  async ensurePublisherApp(input?: {
    readonly displayName?: string;
  }): Promise<Idempotent<PublisherApp>> {
    const uniqueName = publisherAppUniqueName(this.tenantId);
    const displayName = input?.displayName ?? PUBLISHER_APP_DISPLAY_NAME;

    const response = await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/applications`,
      step: 'publisherApp.create',
      jsonBody: publisherAppBody(displayName, uniqueName),
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
      // Entra reports a taken uniqueName as 400, not 409 — without these rules
      // the adopt branch below is unreachable.
      conflictOn: UNIQUE_NAME_CONFLICT_RULES,
    });

    if (response.kind === 'ok') {
      const created = this.parsePublisherApp(response.json, 'publisherApp.create');
      await waitForApplicationAddressable(
        this.http,
        created.objectId,
        [APP_REGISTRATION_SCOPE],
        this.replication,
      );
      await this.ensureServicePrincipal(created.appId);
      this.log(
        `provisioner publisherApp: registered public client '${created.appId}' ` +
          `(uniqueName='${uniqueName}') for delegated catalog publishing`,
      );
      return { outcome: 'created', value: created };
    }

    const adopted = await this.adoptByUniqueName(uniqueName);
    const reconciled = await this.reconcile(adopted);
    await this.ensureServicePrincipal(reconciled.appId);
    return { outcome: 'already-existed', value: reconciled };
  }

  /** Probe: the publisher app as Graph sees it, or `undefined` when absent. */
  async getPublisherApp(): Promise<PublisherApp | undefined> {
    return this.findByUniqueName(publisherAppUniqueName(this.tenantId));
  }

  /**
   * Resolve the registration behind a taken `uniqueName` — a live app on the
   * ordinary re-run, otherwise a soft-deleted one that still holds the name
   * and has to be restored. Mirrors the app-registration step's recovery
   * (byte5ai/omadia#916) so a publisher app deleted by hand does not lock the
   * tenant out of signing in for 30 days.
   */
  private async adoptByUniqueName(uniqueName: string): Promise<PublisherApp> {
    const live = await this.findByUniqueName(uniqueName);
    if (live !== undefined) return live;

    const deleted = await findDeletedApplicationByUniqueName(
      this.http,
      uniqueName,
      [APP_REGISTRATION_SCOPE],
      this.log,
    );
    if (deleted === undefined) {
      throw new UniqueNameReservedError(uniqueName, {
        hint:
          `'${uniqueName}' is already taken but no matching application is visible. ` +
          `A deleted Entra app keeps its uniqueName reserved for ${String(DELETED_ITEM_RETENTION_DAYS)} days ` +
          'while sitting in the recycle bin, and the recycle bin could not be read ' +
          `(it needs Application.ReadWrite.All, not just ${APP_REGISTRATION_SCOPE}). ` +
          "Check 'directory/deletedItems/microsoft.graph.application' in the tenant, then " +
          'restore that object, or permanently delete it to free the name immediately.',
      });
    }

    this.log(
      `provisioner publisherApp: '${uniqueName}' is held by soft-deleted app ` +
        `'${deleted.objectId}' — restoring it`,
    );
    try {
      await restoreDeletedApplication(this.http, deleted.objectId, [
        APP_REGISTRATION_SCOPE,
      ]);
    } catch (err) {
      throw new UniqueNameReservedError(
        uniqueName,
        {
          deletedObjectId: deleted.objectId,
          ...(deleted.deletedDateTime !== undefined
            ? { deletedDateTime: deleted.deletedDateTime }
            : {}),
          hint:
            `'${uniqueName}' is reserved by soft-deleted application '${deleted.objectId}'. ` +
            `Entra holds the name for ${String(DELETED_ITEM_RETENTION_DAYS)} days after a delete and ` +
            `restoring it failed (${redactUnknown(err)}). Restore it manually, or ` +
            `permanently delete it (DELETE /directory/deletedItems/${deleted.objectId}) ` +
            'which frees the name immediately.',
        },
        err,
      );
    }

    // The restore replicates too — the app is not readable instantly.
    return retryWhileReplicating(
      'publisherApp.findByUniqueName',
      deleted.objectId,
      () => this.findByUniqueName(uniqueName),
      this.replication,
    );
  }

  /**
   * Bring an adopted registration up to what the device-code flow needs.
   *
   * The `requiredResourceAccess` write MERGES rather than replaces: a PATCH of
   * that collection is a full overwrite, and silently dropping a permission
   * someone added on purpose is a worse failure than the one being fixed. Only
   * the Graph entry is touched, and within it only our scope is added.
   */
  private async reconcile(app: PublisherApp): Promise<PublisherApp> {
    if (app.isPublicClient && app.declaresCatalogScope) return app;

    const missing: string[] = [];
    if (!app.isPublicClient) missing.push('isFallbackPublicClient');
    if (!app.declaresCatalogScope) missing.push('AppCatalog.ReadWrite.All (delegated)');
    this.log(
      `provisioner publisherApp: repairing adopted app '${app.appId}' — ` +
        `missing ${missing.join(', ')}`,
    );

    const existing = await this.readResourceAccess(app.objectId);
    const patch: Record<string, unknown> = {
      isFallbackPublicClient: true,
      requiredResourceAccess: mergeCatalogScope(existing),
    };

    const response = await this.http.request({
      resource: 'graph',
      method: 'PATCH',
      url: `${GRAPH_BASE}/applications/${encodeURIComponent(app.objectId)}`,
      step: 'publisherApp.patch',
      jsonBody: patch,
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });
    if (response.kind !== 'ok') {
      throw new Error('graph publisherApp.patch unexpected conflict');
    }

    return { ...app, isPublicClient: true, declaresCatalogScope: true };
  }

  /** Current `requiredResourceAccess` of the app, for the merge above. */
  private async readResourceAccess(objectId: string): Promise<unknown[]> {
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/applications/${encodeURIComponent(objectId)}?$select=requiredResourceAccess`,
      step: 'publisherApp.readResourceAccess',
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });
    if (response.kind !== 'ok') return [];
    const value = asRecord(response.json, 'publisherApp.readResourceAccess')[
      'requiredResourceAccess'
    ];
    return Array.isArray(value) ? value : [];
  }

  /** `GET /applications(uniqueName='…')` — `undefined` when no live app holds it. */
  private async findByUniqueName(
    uniqueName: string,
  ): Promise<PublisherApp | undefined> {
    const response = await this.http.request({
      resource: 'graph',
      method: 'GET',
      url: `${GRAPH_BASE}/applications(uniqueName='${encodeURIComponent(escapeODataQuotes(uniqueName))}')`,
      step: 'publisherApp.findByUniqueName',
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
      extraOkStatuses: [404],
    });
    if (response.kind !== 'ok') {
      throw new Error('graph publisherApp.findByUniqueName unexpected conflict');
    }
    if (response.status === 404) return undefined;
    return this.parsePublisherApp(response.json, 'publisherApp.findByUniqueName');
  }

  /** `POST /servicePrincipals`; 409 = it already exists — the idempotent signal. */
  private async ensureServicePrincipal(appId: string): Promise<void> {
    // The consent an admin gives at sign-in is recorded against the service
    // principal, so the app is not consentable until one exists. Creating it
    // here means the very first sign-in can succeed.
    await this.http.request({
      resource: 'graph',
      method: 'POST',
      url: `${GRAPH_BASE}/servicePrincipals`,
      step: 'publisherApp.servicePrincipals.create',
      jsonBody: { appId },
      missingScopesOn403: [APP_REGISTRATION_SCOPE],
    });
  }

  /** Parse + validate a Graph `application`; enforces the audience invariant. */
  private parsePublisherApp(json: unknown, step: string): PublisherApp {
    const body = asRecord(json, step);
    const signInAudience = requireStringField(body, 'signInAudience', step);
    if (signInAudience !== SINGLE_TENANT_SIGN_IN_AUDIENCE) {
      throw new Error(
        `graph ${step}: refusing non-SingleTenant publisher app ` +
          `(signInAudience='${signInAudience}', expected '${SINGLE_TENANT_SIGN_IN_AUDIENCE}')`,
      );
    }
    return {
      appId: requireStringField(body, 'appId', step),
      objectId: requireStringField(body, 'id', step),
      tenantId: this.tenantId,
      displayName: requireStringField(body, 'displayName', step),
      uniqueName:
        typeof body['uniqueName'] === 'string'
          ? body['uniqueName']
          : publisherAppUniqueName(this.tenantId),
      isPublicClient: body['isFallbackPublicClient'] === true,
      declaresCatalogScope: declaresCatalogScope(body['requiredResourceAccess']),
    };
  }
}

/** Does this `requiredResourceAccess` collection already declare our scope? */
export function declaresCatalogScope(requiredResourceAccess: unknown): boolean {
  if (!Array.isArray(requiredResourceAccess)) return false;
  for (const entry of requiredResourceAccess) {
    if (entry === null || typeof entry !== 'object') continue;
    const record = entry as Record<string, unknown>;
    if (record['resourceAppId'] !== GRAPH_RESOURCE_APP_ID) continue;
    const access = record['resourceAccess'];
    if (!Array.isArray(access)) continue;
    for (const item of access) {
      if (item === null || typeof item !== 'object') continue;
      const grant = item as Record<string, unknown>;
      if (
        grant['id'] === APP_CATALOG_DELEGATED_PERMISSION_ID &&
        grant['type'] === 'Scope'
      ) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Add the delegated catalog scope to an existing `requiredResourceAccess`
 * collection without disturbing anything else in it.
 */
export function mergeCatalogScope(existing: readonly unknown[]): unknown[] {
  if (declaresCatalogScope(existing)) return [...existing];

  const merged: unknown[] = [];
  let addedToGraphEntry = false;
  for (const entry of existing) {
    if (
      entry !== null &&
      typeof entry === 'object' &&
      (entry as Record<string, unknown>)['resourceAppId'] === GRAPH_RESOURCE_APP_ID
    ) {
      const record = entry as Record<string, unknown>;
      const access = Array.isArray(record['resourceAccess'])
        ? [...(record['resourceAccess'] as unknown[])]
        : [];
      access.push({ id: APP_CATALOG_DELEGATED_PERMISSION_ID, type: 'Scope' });
      merged.push({ ...record, resourceAccess: access });
      addedToGraphEntry = true;
      continue;
    }
    merged.push(entry);
  }
  if (!addedToGraphEntry) merged.push(catalogResourceAccess());
  return merged;
}

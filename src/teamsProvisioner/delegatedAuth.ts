/**
 * Delegated (user) authentication for the ONE provisioning step that cannot
 * run app-only: the tenant app-catalog upload (byte5ai/omadia#924).
 *
 * WHY THIS EXISTS AT ALL. `POST /appCatalogs/teamsApps` is delegated-only —
 * Graph's own reference lists application permissions as "Not supported." for
 * that verb, and the field test confirms it end to end: with one app-only
 * token the catalog *lookup* succeeds and the *upload* is refused, although
 * `AppCatalog.ReadWrite.All` is assigned as an app role and consented. There
 * is no consent, no role and no API version that changes this. Everything else
 * in the chain (app registration, ARM bot, catalog lookup, team install) stays
 * app-only; only the upload rides a user token.
 *
 * WHY DEVICE CODE AND NOT AUTHORIZATION CODE. A redirect flow needs a redirect
 * URI registered in Entra per deployment URL. omadia is self-hosted, every
 * instance answers on a different host, and moving an instance would silently
 * break the flow until someone edited the app registration. The device
 * authorization grant (RFC 8628) has no redirect URI at all, so the same
 * publisher app works for every install forever.
 *
 * WHY RAW `fetch` AND NOT MSAL. `@azure/msal-node` is available (see
 * `graphObo.ts`), but `acquireTokenByDeviceCode` BLOCKS until the user
 * finishes and reports the user code through a callback. A device-code flow
 * inherently spans two HTTP requests of the operator — "show me the code" and,
 * minutes later, "am I done yet?" — so a blocking call would force the
 * connector to hold flow state in a process-local map. That map dies on every
 * deploy and does not exist on a second instance, which turns a restart during
 * sign-in into a silent dead end. Talking to `/devicecode` and `/token`
 * directly makes the flow POLLABLE FROM ANY PROCESS: all state lives in the
 * handle the caller holds. It is also the house style — `AadTokenCache` in
 * `graphClient.ts` already speaks the same endpoint with hand-rolled form
 * posts, and it keeps msal confined to the OBO path.
 *
 * WHO OWNS THE TOKENS. Not this module. A plugin has no database, so every
 * method takes credentials as parameters and hands renewed ones back; nothing
 * is cached across calls and nothing is written anywhere. The caller persists.
 *
 * SECRETS. Four values here must never reach a log, an error message or a
 * result field: the access token, the refresh token, the device code, and
 * anything Entra echoes back. Bodies on the error path go through
 * {@link redactSecrets} before they are allowed into a message. The `user_code`
 * is the exception — it exists to be displayed.
 */

import {
  DelegatedConsentRequiredError,
  DelegatedTokenExpiredError,
  DeviceCodeFlowError,
} from './errors.js';
import { redactSecrets } from './redact.js';

/** Microsoft Graph's well-known resource (client) id — the `resourceAppId`. */
export const GRAPH_RESOURCE_APP_ID = '00000003-0000-0000-c000-000000000000';

/**
 * The delegated scope `POST /appCatalogs/teamsApps` requires.
 *
 * NOT `AppCatalog.Submit`, even though Graph marks it as the least-privileged
 * delegated permission for the endpoint: the reference note is explicit that
 * `AppCatalog.Submit` only submits apps FOR REVIEW and cannot publish to the
 * catalog. Publishing needs `AppCatalog.ReadWrite.All` (delegated), which is
 * admin-consent-required — hence the "one admin, once per tenant" design.
 */
export const APP_CATALOG_DELEGATED_SCOPE =
  'https://graph.microsoft.com/AppCatalog.ReadWrite.All';

/** `resourceAccess.id` of the DELEGATED `AppCatalog.ReadWrite.All` scope. */
export const APP_CATALOG_DELEGATED_PERMISSION_ID =
  '1ca167d5-1655-44a1-8adf-1414072e1ef9';

/**
 * What the device-code sign-in asks for.
 *
 * `offline_access` is load-bearing: without it Entra issues no refresh token
 * and the admin would have to sign in again every hour, which defeats the
 * whole point. `openid`/`profile` cost no consent and are what makes the
 * result able to say WHICH admin is signed in.
 */
export const DELEGATED_PUBLISH_SCOPES: readonly string[] = [
  APP_CATALOG_DELEGATED_SCOPE,
  'offline_access',
  'openid',
  'profile',
];

const LOGIN_HOST = 'https://login.microsoftonline.com';
/** Device-code polling is a background operation, not a hot path. */
const DEFAULT_TIMEOUT_MS = 10_000;
/** RFC 8628 default when Entra sends no `interval`. */
const DEFAULT_POLL_INTERVAL_SECONDS = 5;
/** Bound on how far a `slow_down` may push the caller's poll cadence. */
const MAX_POLL_INTERVAL_SECONDS = 60;
/** Refresh this long before expiry — same slack as `AadTokenCache`. */
const TOKEN_EXPIRY_SLACK_MS = 60_000;
/** Bound for redacted diagnostic text copied out of a token-endpoint body. */
const MAX_REASON_LENGTH = 200;

/**
 * Probes spent waiting for a freshly created publisher app to be usable at the
 * LOGIN endpoint. Same Entra eventual consistency that bit the app-registration
 * step (byte5ai/omadia#916), one layer further out: Graph can already serve the
 * new application while `login.microsoftonline.com` still answers
 * `unauthorized_client`, because the sign-in stack replicates separately from
 * the directory. Without this the very first sign-in after install fails.
 */
const DEVICE_CODE_CLIENT_REPLICATION_ATTEMPTS = 6;
const DEVICE_CODE_CLIENT_REPLICATION_INTERVAL_MS = 2000;
const MAX_DEVICE_CODE_CLIENT_REPLICATION_DELAY_MS = 8000;

/**
 * OAuth `error` values that mean "the client id is not usable at the login
 * endpoint YET" — indistinguishable at the protocol level from "the client id
 * is wrong", which is why the retry budget is small and the final error still
 * names both possibilities.
 */
const CLIENT_REPLICATION_ERRORS: ReadonlySet<string> = new Set([
  'invalid_client',
  'unauthorized_client',
  'invalid_request',
]);

/**
 * Opaque continuation token for an in-flight device-code sign-in.
 *
 * SECRET-GRADE — treat it exactly like a password. It carries the RFC 8628
 * `device_code`, and whoever holds that can redeem the token the admin is
 * about to sign for. Store it the way a secret is stored, never log it, never
 * render it in a UI, and drop it as soon as the flow reaches a terminal state.
 *
 * It is a plain `string` on purpose: the middleware mirrors this contract
 * structurally rather than importing it, so a branded type would not survive
 * the boundary. Its contents are base64url-encoded JSON — FRAMING, not
 * encryption. It is short-lived (~15 minutes) by protocol design, which bounds
 * the damage, but it is not protected by the encoding.
 */
export type DeviceCodeFlowHandle = string;

/** What the operator has to be shown to complete the sign-in. */
export interface DeviceCodeStart {
  /** SHOW THIS. The code the admin types at {@link verificationUri}. */
  readonly userCode: string;
  /** SHOW THIS. Where the admin enters the code (usually microsoft.com/devicelogin). */
  readonly verificationUri: string;
  /** Entra's own full instruction sentence, when it sent one. */
  readonly message?: string;
  /** ISO-8601 instant after which the code stops working (~15 min out). */
  readonly expiresAt: string;
  /** Minimum seconds between polls — going faster earns a `slow_down`. */
  readonly intervalSeconds: number;
  /** SECRET. Hand back to {@link DelegatedAuthClient.pollDeviceCode}. */
  readonly flowHandle: DeviceCodeFlowHandle;
  /** Delegated scopes this sign-in acquires. */
  readonly scopes: readonly string[];
  /**
   * Where to send a Global Administrator if the sign-in page answers "approval
   * required". Contains only the public client id — never a secret.
   */
  readonly adminConsentUrl: string;
}

/** Who signed in — best-effort, read from the id_token we were issued. */
export interface DelegatedAccount {
  /** `preferred_username`, normally the admin's UPN. */
  readonly username?: string;
  /** `name` claim — the display name. */
  readonly displayName?: string;
  /** `oid` claim — the user's directory object id. */
  readonly objectId?: string;
  /** `tid` claim — the tenant the admin signed into. */
  readonly tenantId?: string;
}

/**
 * A usable delegated credential. BOTH token fields are secrets; only
 * {@link expiresAt}, {@link scopes}, {@link clientId} and {@link account} are
 * safe to show.
 *
 * {@link clientId} travels WITH the token set on purpose: refreshing needs the
 * publisher app's client id, and carrying it here means the caller never has to
 * re-resolve the publisher app (a Graph round trip) just to renew a token.
 */
export interface DelegatedTokenSet {
  /** SECRET. Bearer token for the delegated catalog upload. */
  readonly accessToken: string;
  /** SECRET. Persist this — it is what avoids a second admin sign-in. */
  readonly refreshToken: string;
  /** ISO-8601 expiry of {@link accessToken}. */
  readonly expiresAt: string;
  /** Scopes the access token actually carries (Entra's answer, not our ask). */
  readonly scopes: readonly string[];
  /** Publisher app the tokens belong to — needed for the refresh exchange. */
  readonly clientId: string;
  /** Tenant the tokens are valid in. */
  readonly tenantId: string;
  /** Best-effort identity of the admin who signed in. */
  readonly account?: DelegatedAccount;
}

/**
 * Poll outcome. The three EXPECTED terminal states of RFC 8628 are results,
 * not exceptions — a caller renders them, it does not catch them.
 */
export type DeviceCodePollResult =
  | DeviceCodePending
  | DeviceCodeSucceeded
  | DeviceCodeExpired
  | DeviceCodeDeclined;

/** The admin has not finished yet. Poll again after {@link retryAfterSeconds}. */
export interface DeviceCodePending {
  readonly status: 'pending';
  /**
   * Honour this — it already accounts for a `slow_down` from Entra. Polling
   * faster than the server asked is how a flow earns a hard rejection.
   */
  readonly retryAfterSeconds: number;
}

/** The admin signed in. {@link tokens} must be persisted by the caller. */
export interface DeviceCodeSucceeded {
  readonly status: 'succeeded';
  readonly tokens: DelegatedTokenSet;
}

/** The 15-minute window elapsed. Start a new flow; the handle is dead. */
export interface DeviceCodeExpired {
  readonly status: 'expired';
  /** Redacted diagnostic from Entra, when it sent one. */
  readonly reason?: string;
}

/**
 * The sign-in did not complete and will not: the admin cancelled, OR the
 * tenant refuses this flow.
 *
 * The second case is the one worth reading twice. Microsoft's own guidance is
 * to *block device code flow wherever possible*, and Conditional Access has a
 * dedicated "authentication flows" condition for exactly that. In a tenant
 * with that policy the browser leg is refused and the poll lands HERE, not on
 * a distinguishable error. {@link reason} carries Entra's redacted
 * `error_description` (with its AADSTS code) — the only thing that tells a
 * cancelled sign-in apart from a policy-blocked one.
 */
export interface DeviceCodeDeclined {
  readonly status: 'declined';
  /** Redacted `error_description` — read the AADSTS code before blaming the admin. */
  readonly reason?: string;
}

export interface DelegatedAuthClientOptions {
  /** Tenant the sign-in targets — the customer tenant, never `/common`. */
  readonly tenantId: string;
  /** Publisher app (public client) the flow authenticates as. */
  readonly clientId: string;
  /** Override the requested scopes (tests, future steps). */
  readonly scopes?: readonly string[];
  /** Test seam — identical to the `GraphClient` injection point. */
  readonly fetchImpl?: typeof fetch;
  /** Test seam for the client-replication waits. */
  readonly sleep?: (ms: number) => Promise<void>;
  readonly log?: (msg: string) => void;
  /** Budget for a single token/devicecode POST (default 10 s). */
  readonly timeoutMs?: number;
}

const defaultSleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms));

/**
 * The device-code half of the delegated publish path: start a sign-in, poll it,
 * and renew the credential it produced.
 *
 * Stateless between calls by construction — no token cache, no flow map, no
 * disk. Two instances in two processes serve the same flow interchangeably as
 * long as they are handed the same {@link DeviceCodeFlowHandle}.
 */
export class DelegatedAuthClient {
  private readonly tenantId: string;
  private readonly clientId: string;
  private readonly scopes: readonly string[];
  private readonly fetchImpl: typeof fetch;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly log: (msg: string) => void;
  private readonly timeoutMs: number;

  constructor(opts: DelegatedAuthClientOptions) {
    this.tenantId = requireNonEmpty(opts.tenantId, 'tenantId');
    this.clientId = requireNonEmpty(opts.clientId, 'clientId');
    this.scopes = opts.scopes ?? DELEGATED_PUBLISH_SCOPES;
    this.fetchImpl = opts.fetchImpl ?? fetch;
    this.sleep = opts.sleep ?? defaultSleep;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
    this.timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  }

  /** Tenant-wide admin-consent URL for the publisher app (carries no secret). */
  get adminConsentUrl(): string {
    return adminConsentUrl(this.tenantId, this.clientId);
  }

  /**
   * Phase 1 — `POST /oauth2/v2.0/devicecode`.
   *
   * Only call this when the operator is actually ready to sign in: the code
   * expires ~15 minutes from HERE, not from when they get around to it.
   *
   * Retries a small number of times while the login stack still answers
   * `unauthorized_client` for a just-created publisher app — see
   * {@link DEVICE_CODE_CLIENT_REPLICATION_ATTEMPTS}. That window is why a
   * first-install sign-in used to fail once for no visible reason.
   */
  async startDeviceCode(): Promise<DeviceCodeStart> {
    const body = new URLSearchParams({
      client_id: this.clientId,
      scope: this.scopes.join(' '),
    });

    let lastFailure: DeviceCodeFlowError | undefined;
    for (
      let attempt = 0;
      attempt < DEVICE_CODE_CLIENT_REPLICATION_ATTEMPTS;
      attempt += 1
    ) {
      const { status, json, text } = await this.postForm('devicecode', body);

      if (status >= 200 && status < 300) {
        return this.parseDeviceCodeStart(json, text);
      }

      const oauthError = oauthErrorCode(json);
      if (oauthError === undefined || !CLIENT_REPLICATION_ERRORS.has(oauthError)) {
        throw new DeviceCodeFlowError(
          `the device authorization request was rejected: ${describe(json, text)}`,
          { ...(oauthError !== undefined ? { oauthError } : {}), status },
        );
      }

      lastFailure = new DeviceCodeFlowError(
        `the publisher app '${this.clientId}' is not accepted at the sign-in ` +
          'endpoint. Either the app registration has not replicated to the ' +
          'login stack yet, or "Allow public client flows" is off on it — the ' +
          `device code grant requires a public client. Detail: ${describe(json, text)}`,
        { oauthError, status },
      );

      const delayMs = Math.min(
        DEVICE_CODE_CLIENT_REPLICATION_INTERVAL_MS * 2 ** attempt,
        MAX_DEVICE_CODE_CLIENT_REPLICATION_DELAY_MS,
      );
      this.log(
        `provisioner delegated.devicecode: publisher app '${this.clientId}' not ` +
          `usable at the login endpoint yet (probe ${String(attempt + 1)}/` +
          `${String(DEVICE_CODE_CLIENT_REPLICATION_ATTEMPTS)}), waiting ${String(delayMs)}ms`,
      );
      await this.sleep(delayMs);
    }

    throw (
      lastFailure ??
      new DeviceCodeFlowError('the device authorization request never succeeded')
    );
  }

  /**
   * Phase 2 — `POST /oauth2/v2.0/token` with
   * `grant_type=urn:ietf:params:oauth:grant-type:device_code`.
   *
   * Safe to call from a different process than {@link startDeviceCode}: every
   * piece of state the exchange needs travels in the handle. Idempotent while
   * pending; after `'succeeded'` the code is spent and a further poll answers
   * `'declined'`/`'expired'`.
   */
  async pollDeviceCode(input: {
    readonly flowHandle: DeviceCodeFlowHandle;
  }): Promise<DeviceCodePollResult> {
    const flow = decodeFlowHandle(input.flowHandle);

    const body = new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: flow.clientId,
      device_code: flow.deviceCode,
    });
    const { status, json, text } = await this.postForm('token', body, flow.tenantId);

    if (status >= 200 && status < 300) {
      return {
        status: 'succeeded',
        tokens: this.parseTokenSet(json, text, flow.tenantId, flow.clientId),
      };
    }

    const oauthError = oauthErrorCode(json);
    const reason = diagnosticReason(json, text);

    switch (oauthError) {
      case 'authorization_pending':
        return { status: 'pending', retryAfterSeconds: flow.intervalSeconds };
      case 'slow_down':
        // RFC 8628 §3.5: back off permanently, do not just retry once. The
        // caller polls at whatever we return, so widening it here is what
        // actually slows the loop down.
        return {
          status: 'pending',
          retryAfterSeconds: Math.min(
            flow.intervalSeconds + DEFAULT_POLL_INTERVAL_SECONDS,
            MAX_POLL_INTERVAL_SECONDS,
          ),
        };
      case 'expired_token':
        return { status: 'expired', ...(reason !== undefined ? { reason } : {}) };
      case 'authorization_declined':
      case 'access_denied':
        return { status: 'declined', ...(reason !== undefined ? { reason } : {}) };
      default:
        break;
    }

    if (indicatesMissingConsent(json, text)) {
      throw new DelegatedConsentRequiredError(
        'delegated.deviceCode.poll',
        this.scopes,
        adminConsentUrl(flow.tenantId, flow.clientId),
        new Error(describe(json, text)),
      );
    }

    throw new DeviceCodeFlowError(
      `polling the device code failed: ${describe(json, text)}`,
      { ...(oauthError !== undefined ? { oauthError } : {}), status },
    );
  }

  /**
   * `grant_type=refresh_token` — a fresh access token from a stored refresh
   * token, with no human involved.
   *
   * The returned set carries Entra's ROTATED refresh token. Persisting it is
   * not optional bookkeeping: Entra hands out a new refresh token on every
   * exchange and the old one ages out on its own inactivity window, so a
   * caller that keeps writing back the original one eventually forces the
   * admin to sign in again for no reason. When Entra returns no new refresh
   * token, the previous one is carried forward unchanged.
   *
   * @throws {DelegatedTokenExpiredError} `'refresh-token-invalid'` when Entra
   *   answers `invalid_grant` — the credential is dead, a human must sign in.
   * @throws {DelegatedConsentRequiredError} when tenant consent was withdrawn
   *   after the sign-in.
   */
  async refresh(input: {
    readonly refreshToken: string;
    /** Defaults to this client's tenant/app; pass the stored ones when they differ. */
    readonly tenantId?: string;
    readonly clientId?: string;
  }): Promise<DelegatedTokenSet> {
    const refreshToken = requireNonEmpty(input.refreshToken, 'refreshToken');
    const tenantId = input.tenantId ?? this.tenantId;
    const clientId = input.clientId ?? this.clientId;

    const body = new URLSearchParams({
      grant_type: 'refresh_token',
      client_id: clientId,
      refresh_token: refreshToken,
      scope: this.scopes.join(' '),
    });
    const { status, json, text } = await this.postForm('token', body, tenantId);

    if (status >= 200 && status < 300) {
      // Carry the previous refresh token forward when Entra rotates nothing —
      // dropping it would strand a still-valid credential.
      return this.parseTokenSet(json, text, tenantId, clientId, refreshToken);
    }

    if (indicatesMissingConsent(json, text)) {
      throw new DelegatedConsentRequiredError(
        'delegated.token.refresh',
        this.scopes,
        adminConsentUrl(tenantId, clientId),
        new Error(describe(json, text)),
      );
    }

    if (oauthErrorCode(json) === 'invalid_grant') {
      throw new DelegatedTokenExpiredError(
        'refresh-token-invalid',
        new Error(describe(json, text)),
      );
    }

    throw new DeviceCodeFlowError(
      `refreshing the delegated token failed: ${describe(json, text)}`,
      {
        ...(oauthErrorCode(json) !== undefined
          ? { oauthError: oauthErrorCode(json) as string }
          : {}),
        status,
      },
    );
  }

  /**
   * Return a token set whose access token is good for the next minute,
   * refreshing only when it is not.
   *
   * `refreshed` tells the caller whether the credential CHANGED and therefore
   * has to be written back — the difference between a persistence write per
   * upload and one per hour.
   */
  async ensureFreshToken(tokens: DelegatedTokenSet): Promise<{
    readonly tokens: DelegatedTokenSet;
    readonly refreshed: boolean;
  }> {
    if (!isAccessTokenStale(tokens)) return { tokens, refreshed: false };
    const refreshed = await this.refresh({
      refreshToken: tokens.refreshToken,
      tenantId: tokens.tenantId,
      clientId: tokens.clientId,
    });
    return { tokens: refreshed, refreshed: true };
  }

  /** One form POST to the identity platform. Never logs the request body. */
  private async postForm(
    endpoint: 'devicecode' | 'token',
    body: URLSearchParams,
    tenantId: string = this.tenantId,
  ): Promise<{ status: number; json: unknown; text: string }> {
    const url = `${LOGIN_HOST}/${encodeURIComponent(tenantId)}/oauth2/v2.0/${endpoint}`;
    const response = await this.fetchImpl(url, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
      signal: AbortSignal.timeout(this.timeoutMs),
    });
    // The body is read exactly once, then reused for both the success parse
    // and the diagnostic text — and the diagnostic copy is redacted at every
    // use site, never here, so the success path keeps the real values.
    let text = '';
    try {
      text = await response.text();
    } catch {
      text = '';
    }
    let json: unknown;
    try {
      json = JSON.parse(text);
    } catch {
      json = undefined;
    }
    return { status: response.status, json, text };
  }

  private parseDeviceCodeStart(json: unknown, text: string): DeviceCodeStart {
    const record = asRecord(json);
    const deviceCode = stringField(record, 'device_code');
    const userCode = stringField(record, 'user_code');
    const verificationUri =
      stringField(record, 'verification_uri') ??
      stringField(record, 'verification_url');
    if (
      deviceCode === undefined ||
      userCode === undefined ||
      verificationUri === undefined
    ) {
      throw new DeviceCodeFlowError(
        `the device authorization response was missing device_code/user_code/verification_uri: ${describe(json, text)}`,
      );
    }
    const expiresInSeconds = numberField(record, 'expires_in') ?? 900;
    const intervalSeconds =
      numberField(record, 'interval') ?? DEFAULT_POLL_INTERVAL_SECONDS;
    const message = stringField(record, 'message');
    const expiresAt = new Date(Date.now() + expiresInSeconds * 1000).toISOString();

    // Logged WITHOUT the device code — the user code is display material, the
    // device code is the flow's secret.
    this.log(
      `provisioner delegated.devicecode: sign-in started for publisher app ` +
        `'${this.clientId}', code valid until ${expiresAt}`,
    );

    return {
      userCode,
      verificationUri,
      ...(message !== undefined ? { message } : {}),
      expiresAt,
      intervalSeconds,
      flowHandle: encodeFlowHandle({
        v: 1,
        deviceCode,
        tenantId: this.tenantId,
        clientId: this.clientId,
        intervalSeconds,
        expiresAt,
      }),
      scopes: this.scopes,
      adminConsentUrl: this.adminConsentUrl,
    };
  }

  private parseTokenSet(
    json: unknown,
    text: string,
    tenantId: string,
    clientId: string,
    fallbackRefreshToken?: string,
  ): DelegatedTokenSet {
    const record = asRecord(json);
    const accessToken = stringField(record, 'access_token');
    if (accessToken === undefined) {
      throw new DeviceCodeFlowError(
        `the token response carried no access_token: ${describe(json, text)}`,
      );
    }
    const refreshToken =
      stringField(record, 'refresh_token') ?? fallbackRefreshToken;
    if (refreshToken === undefined) {
      // Without a refresh token the "one sign-in per tenant" promise is dead —
      // fail loudly here rather than a week later when the access token ages
      // out and the upload starts demanding a human.
      throw new DeviceCodeFlowError(
        'the token response carried no refresh_token — the sign-in did not ' +
          "request 'offline_access', or the tenant's token lifetime policy " +
          'suppresses refresh tokens for this flow. Without one, every catalog ' +
          'upload would need a fresh admin sign-in',
      );
    }
    const expiresInSeconds = numberField(record, 'expires_in') ?? 3600;
    const scopeField = stringField(record, 'scope');
    const account = accountFromIdToken(stringField(record, 'id_token'));

    return {
      accessToken,
      refreshToken,
      expiresAt: new Date(Date.now() + expiresInSeconds * 1000).toISOString(),
      scopes:
        scopeField !== undefined && scopeField.length > 0
          ? scopeField.split(' ')
          : this.scopes,
      clientId,
      tenantId,
      ...(account !== undefined ? { account } : {}),
    };
  }
}

/**
 * What the caller can safely SHOW about a stored delegated credential.
 *
 * Everything here is display-grade by construction — no token, not even a
 * prefix of one. `signedIn` is about whether a credential exists at all;
 * `accessTokenStale` is about whether the next upload will spend a refresh.
 * A stale access token is NOT "signed out": the refresh token is what keeps
 * the sign-in alive, and it long outlives the hour-long access token.
 */
export type DelegatedSignInStatus =
  | DelegatedSignedOutStatus
  | DelegatedSignedInStatus;

/** No credential stored — the operator has to run the device-code flow. */
export interface DelegatedSignedOutStatus {
  readonly signedIn: false;
  /** The delegated scopes a sign-in would acquire. */
  readonly requiredScopes: readonly string[];
}

/** A credential is stored. Whether it still WORKS is only knowable by using it. */
export interface DelegatedSignedInStatus {
  readonly signedIn: true;
  /** ISO-8601 expiry of the access token (not of the sign-in). */
  readonly accessTokenExpiresAt: string;
  /** `true` when the next call will refresh before it can publish. */
  readonly accessTokenStale: boolean;
  /** Scopes the stored access token carries. */
  readonly scopes: readonly string[];
  /** `true` when those scopes actually cover the catalog upload. */
  readonly coversCatalogPublish: boolean;
  /** Publisher app the credential belongs to. */
  readonly clientId: string;
  readonly tenantId: string;
  /** Which admin signed in, when the id_token said so. */
  readonly account?: DelegatedAccount;
  /** Where an admin can withdraw the tenant grant. */
  readonly adminConsentUrl: string;
}

/** Outcome of {@link revokeDelegatedSignIn} — see its doc for the honest limits. */
export interface DelegatedRevokeResult {
  /**
   * `'discard-stored-tokens'` — there was a credential and the caller must now
   * delete it from its own store; `'not-signed-in'` — nothing to do.
   */
  readonly outcome: 'discard-stored-tokens' | 'not-signed-in';
  /**
   * Portal location where an admin withdraws the tenant-wide consent. Present
   * only when a credential existed, since it names the publisher app.
   */
  readonly adminConsentUrl?: string;
  /** One sentence an operator can act on, safe to render verbatim. */
  readonly note: string;
}

/** Describe a stored credential without calling anything. Never returns a token. */
export function describeSignInStatus(
  tokens: DelegatedTokenSet | undefined,
): DelegatedSignInStatus {
  if (tokens === undefined) {
    return { signedIn: false, requiredScopes: DELEGATED_PUBLISH_SCOPES };
  }
  return {
    signedIn: true,
    accessTokenExpiresAt: tokens.expiresAt,
    accessTokenStale: isAccessTokenStale(tokens),
    scopes: tokens.scopes,
    coversCatalogPublish: coversCatalogPublish(tokens.scopes),
    clientId: tokens.clientId,
    tenantId: tokens.tenantId,
    ...(tokens.account !== undefined ? { account: tokens.account } : {}),
    adminConsentUrl: adminConsentUrl(tokens.tenantId, tokens.clientId),
  };
}

/**
 * Does this scope list actually permit the catalog upload?
 *
 * Entra reports granted scopes in SHORT form (`AppCatalog.ReadWrite.All`) even
 * when the request used the fully-qualified resource URI, so both spellings
 * count — comparing only the string we sent would report a working credential
 * as insufficient.
 */
export function coversCatalogPublish(scopes: readonly string[]): boolean {
  const shortName = APP_CATALOG_DELEGATED_SCOPE.slice(
    APP_CATALOG_DELEGATED_SCOPE.lastIndexOf('/') + 1,
  ).toLowerCase();
  return scopes.some((scope) => {
    const normalised = scope.toLowerCase();
    return (
      normalised === APP_CATALOG_DELEGATED_SCOPE.toLowerCase() ||
      normalised === shortName
    );
  });
}

/** Caller-side revoke instructions. Pure — see the accessor doc for why. */
export function revokeInstructions(
  tokens: DelegatedTokenSet | undefined,
): DelegatedRevokeResult {
  if (tokens === undefined) {
    return {
      outcome: 'not-signed-in',
      note: 'No delegated credential was stored — nothing to revoke.',
    };
  }
  return {
    outcome: 'discard-stored-tokens',
    adminConsentUrl: adminConsentUrl(tokens.tenantId, tokens.clientId),
    note:
      'Delete the stored access and refresh tokens now. The connector holds no ' +
      'copy, so discarding them ends this sign-in. To also withdraw the ' +
      "tenant's consent, remove the publisher app under Microsoft Entra ID → " +
      'Enterprise applications → Permissions; leave the app registration itself ' +
      'in place, since deleting it reserves its name for 30 days and blocks the ' +
      'next sign-in.',
  };
}

/** Tenant-wide admin-consent URL — public client id only, never a secret. */
export function adminConsentUrl(tenantId: string, clientId: string): string {
  return (
    `${LOGIN_HOST}/${encodeURIComponent(tenantId)}/adminconsent` +
    `?client_id=${encodeURIComponent(clientId)}`
  );
}

/** Is the access token within the refresh slack of expiry (or already past it)? */
export function isAccessTokenStale(tokens: DelegatedTokenSet): boolean {
  const expiresAt = Date.parse(tokens.expiresAt);
  if (!Number.isFinite(expiresAt)) return true;
  return expiresAt - TOKEN_EXPIRY_SLACK_MS <= Date.now();
}

// ---------------------------------------------------------------------------
// Flow handle — base64url JSON. FRAMING, not encryption; see the type doc.
// ---------------------------------------------------------------------------

interface FlowHandlePayload {
  readonly v: 1;
  readonly deviceCode: string;
  readonly tenantId: string;
  readonly clientId: string;
  readonly intervalSeconds: number;
  readonly expiresAt: string;
}

function encodeFlowHandle(payload: FlowHandlePayload): DeviceCodeFlowHandle {
  return Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
}

function decodeFlowHandle(handle: string): FlowHandlePayload {
  requireNonEmpty(handle, 'flowHandle');
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(handle, 'base64url').toString('utf8'));
  } catch {
    // The handle itself is the secret — the message says WHAT is wrong and
    // never echoes the value back.
    throw new DeviceCodeFlowError(
      'the device-code flow handle is not readable — it must be passed back ' +
        'verbatim from startDeviceCode',
    );
  }
  const record = asRecord(parsed);
  const deviceCode = stringField(record, 'deviceCode');
  const tenantId = stringField(record, 'tenantId');
  const clientId = stringField(record, 'clientId');
  if (
    record['v'] !== 1 ||
    deviceCode === undefined ||
    tenantId === undefined ||
    clientId === undefined
  ) {
    throw new DeviceCodeFlowError(
      'the device-code flow handle is malformed or from an incompatible version',
    );
  }
  return {
    v: 1,
    deviceCode,
    tenantId,
    clientId,
    intervalSeconds:
      numberField(record, 'intervalSeconds') ?? DEFAULT_POLL_INTERVAL_SECONDS,
    expiresAt: stringField(record, 'expiresAt') ?? new Date().toISOString(),
  };
}

// ---------------------------------------------------------------------------
// Response helpers. Everything that can end up in a message goes through
// redactSecrets — a token endpoint body is exactly where a credential hides.
// ---------------------------------------------------------------------------

function asRecord(json: unknown): Record<string, unknown> {
  return json !== null && typeof json === 'object'
    ? (json as Record<string, unknown>)
    : {};
}

function stringField(
  record: Record<string, unknown>,
  field: string,
): string | undefined {
  const value = record[field];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function numberField(
  record: Record<string, unknown>,
  field: string,
): number | undefined {
  const value = record[field];
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function oauthErrorCode(json: unknown): string | undefined {
  return stringField(asRecord(json), 'error');
}

/** Redacted, bounded `error_description` for a poll result's `reason`. */
function diagnosticReason(json: unknown, text: string): string | undefined {
  const description = stringField(asRecord(json), 'error_description');
  const raw = description ?? (text.length > 0 ? text : undefined);
  if (raw === undefined) return undefined;
  return truncateReason(redactSecrets(raw));
}

/** Redacted, bounded description for an error MESSAGE. */
function describe(json: unknown, text: string): string {
  const record = asRecord(json);
  const code = stringField(record, 'error');
  const description = stringField(record, 'error_description');
  if (code !== undefined) {
    return truncateReason(
      redactSecrets(
        description !== undefined ? `${code} — ${description}` : code,
      ),
    );
  }
  return text.length > 0 ? truncateReason(redactSecrets(text)) : '(empty body)';
}

function truncateReason(value: string): string {
  const collapsed = value.replace(/\s+/g, ' ').trim();
  return collapsed.length <= MAX_REASON_LENGTH
    ? collapsed
    : `${collapsed.slice(0, MAX_REASON_LENGTH - 1)}…`;
}

/**
 * Does this failure mean "the tenant has not consented", as opposed to any
 * other rejection? Entra signals it three different ways depending on which
 * leg answered, so all three are checked.
 */
function indicatesMissingConsent(json: unknown, text: string): boolean {
  const record = asRecord(json);
  const code = stringField(record, 'error');
  if (code === 'consent_required' || code === 'interaction_required') return true;
  const description = (
    stringField(record, 'error_description') ??
    stringField(record, 'suberror') ??
    text
  ).toUpperCase();
  // AADSTS65001 = no consent on record; AADSTS90094 = the scope needs an admin.
  return description.includes('AADSTS65001') || description.includes('AADSTS90094');
}

/**
 * Read the account claims out of the id_token we were issued.
 *
 * Reading OUR OWN id_token's payload for display is the intended use; this
 * deliberately does NOT validate a signature and the result is never used for
 * an authorization decision — it exists so an operator can see WHICH admin the
 * stored credential belongs to. Any parse failure yields `undefined`: a
 * cosmetic field must never fail a sign-in.
 */
function accountFromIdToken(idToken: string | undefined): DelegatedAccount | undefined {
  if (idToken === undefined) return undefined;
  const payloadSegment = idToken.split('.')[1];
  if (payloadSegment === undefined || payloadSegment.length === 0) return undefined;
  let claims: Record<string, unknown>;
  try {
    claims = asRecord(
      JSON.parse(Buffer.from(payloadSegment, 'base64url').toString('utf8')),
    );
  } catch {
    return undefined;
  }
  const username =
    stringField(claims, 'preferred_username') ?? stringField(claims, 'upn');
  const displayName = stringField(claims, 'name');
  const objectId = stringField(claims, 'oid');
  const tenantId = stringField(claims, 'tid');
  if (
    username === undefined &&
    displayName === undefined &&
    objectId === undefined &&
    tenantId === undefined
  ) {
    return undefined;
  }
  return {
    ...(username !== undefined ? { username } : {}),
    ...(displayName !== undefined ? { displayName } : {}),
    ...(objectId !== undefined ? { objectId } : {}),
    ...(tenantId !== undefined ? { tenantId } : {}),
  };
}

function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

/**
 * Single token+fetch choke point for every Graph and ARM call of the
 * `teamsProvisioner@1` capability (epic byte5ai/omadia#860, capability issue
 * byte5ai/omadia-m365-connector#3).
 *
 * Every provisioning verb — `POST /applications`, `addPassword`,
 * `servicePrincipals`, the ARM `PUT .../botServices/{name}` (incl. the
 * MsTeamsChannel enable), `POST /appCatalogs/teamsApps`,
 * `POST /teams/{id}/installedApps`, and the DELETE rollbacks — routes through
 * {@link ProvisioningHttp.request}. That is what makes the 403/409/429 matrix
 * provable ONCE instead of per call:
 *
 * - 403 → typed {@link ConsentMissingError} carrying the missing scope set
 *   and which API rejected, so the agent factory can fall back.
 * - 409 → NOT an error: the `{ kind: 'conflict' }` result is the idempotent
 *   signal the step units map to `Idempotent<T>` `'already-existed'`. Some
 *   APIs report a duplicate WITHOUT a 409 — Entra answers a taken
 *   `uniqueName` on `POST /applications` with **400 Request_BadRequest**
 *   (byte5ai/omadia#916) — so a request may declare extra
 *   {@link ProvisioningConflictRule}s that map onto the same signal.
 * - 429 → backoff honouring the `Retry-After` header (never a fixed sleep);
 *   exhausted budget → {@link ProvisioningThrottledError}.
 *
 * Tokens come from the shared {@link AadTokenCache} (the generalised
 * `GraphClient` plumbing) — cached per (tenant, clientId, scope), so the two
 * audiences (`graph.microsoft.com/.default` vs `management.azure.com/.default`)
 * and the optional dedicated ARM service principal never collide.
 *
 * ONE EXCEPTION, and only one: `ProvisioningRequest.bearerToken`
 * (byte5ai/omadia#924). `POST /appCatalogs/teamsApps` is delegated-only in
 * Graph — application permissions are documented as "Not supported." and the
 * field test confirms the upload is refused for an app-only token that the
 * catalog LOOKUP accepts. That one verb therefore carries a user token supplied
 * by the caller. It is never cached here and never falls back to the app
 * identity, so the two auth modes cannot be confused at runtime.
 *
 * Invariants: token-based REST only (no `az` CLI, no child_process, no msal —
 * msal stays confined to the delegated OBO flow in `graphObo.ts`); error
 * messages NEVER contain the client secret or a bearer token, bodies are
 * bounded via `truncate(body, 200)` following the `GraphClient` precedent.
 */

import {
  AadTokenCache,
  GRAPH_TOKEN_SCOPE,
  safeBody,
  truncate,
  type AadClientCredential,
} from '../graphClient.js';
import {
  ConsentMissingError,
  ProvisioningRequestError,
  ProvisioningThrottledError,
} from './errors.js';
import { redactSecrets } from './redact.js';

/**
 * Bound AND scrub a response body before it may appear in a message.
 *
 * `truncate` alone was enough while every credential this layer handled was a
 * client secret it never read back. Since the delegated catalog upload
 * (byte5ai/omadia#924) a bearer token travels on requests here, and a 4xx body
 * — or an identity-platform body reached through this path — can echo one. The
 * redaction runs on the ERROR path only, so the success path pays nothing.
 */
function safeDetail(bodyText: string): string {
  return truncate(redactSecrets(bodyText), 200);
}

/** Which API a request targets — also the audience key for token + errors. */
export type ProvisioningResource = 'graph' | 'arm';

/** Token audience for ARM (`management.azure.com`) requests. */
const ARM_TOKEN_SCOPE = 'https://management.azure.com/.default';

const DEFAULT_MAX_429_RETRIES = 3;
/** Fallback backoff base when a 429 arrives WITHOUT Retry-After (doubles per attempt). */
const DEFAULT_BACKOFF_BASE_MS = 1000;
/**
 * Upper bound for any single Retry-After-derived wait. Graph/ARM can emit
 * hints of hours (or an HTTP-date far in the future); sleeping that long
 * inside a provisioning job is an unbounded hang. A 429 hint above this cap
 * aborts with {@link ProvisioningThrottledError} carrying the full
 * `retryAfterSeconds` so the job runner (byte5ai/omadia#864) reschedules;
 * long-running-poll pacing hints are clamped to the cap instead.
 */
const MAX_BACKOFF_MS = 60_000;
/** Provisioning is a background job, not the Teams hot path — token POSTs get 10 s. */
const TOKEN_TIMEOUT_MS = 10_000;
/** Per-request budgets: Graph verbs are quick; ARM PUTs are long-running. */
const DEFAULT_REQUEST_TIMEOUT_MS: Record<ProvisioningResource, number> = {
  graph: 15_000,
  arm: 60_000,
};
const DEFAULT_POLL_INTERVAL_MS = 2000;
const DEFAULT_MAX_POLL_ATTEMPTS = 60;

export interface ProvisioningHttpOptions {
  /** Connector app credential — always the Graph identity. */
  readonly graphCredential: AadClientCredential;
  /**
   * Optional dedicated ARM service principal (`azure_sp_client_id` /
   * `azure_sp_client_secret`). Absent → the Graph credential is reused for
   * the ARM audience ("reuse app" mode of the config unit).
   */
  readonly armCredential?: AadClientCredential;
  readonly log?: (msg: string) => void;
  /** Test seam — identical to the `GraphClient` injection point. */
  readonly fetchImpl?: typeof fetch;
  /** Test seam for backoff/poll waits. */
  readonly sleep?: (ms: number) => Promise<void>;
  /** 429 retries per request before ProvisioningThrottledError (default 3). */
  readonly max429Retries?: number;
  /** Override the per-resource request budgets (ms). */
  readonly requestTimeoutMs?: Partial<Record<ProvisioningResource, number>>;
  /** ARM long-running poll cadence when the API sends no Retry-After (ms). */
  readonly pollIntervalMs?: number;
  /** ARM long-running poll budget (default 60 attempts). */
  readonly maxPollAttempts?: number;
}

export interface ProvisioningRequest {
  readonly resource: ProvisioningResource;
  readonly method: 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';
  /** Absolute URL (Graph v1.0 or ARM incl. api-version). */
  readonly url: string;
  /** Short step label for log/error messages, e.g. `applications.create`. */
  readonly step: string;
  /** JSON payload (serialised with content-type application/json). */
  readonly jsonBody?: unknown;
  /** Raw payload (e.g. the app-package zip for the catalog upload). */
  readonly rawBody?: {
    readonly bytes: Uint8Array;
    readonly contentType: string;
  };
  /** Scopes reported when this call answers 403 (→ ConsentMissingError). */
  readonly missingScopesOn403: readonly string[];
  /**
   * Use THIS bearer token instead of the app-only one from the token cache.
   *
   * Exists for exactly one verb: `POST /appCatalogs/teamsApps`, which Graph
   * supports for DELEGATED permissions only (byte5ai/omadia#924). Everything
   * else stays app-only, so this stays an opt-in override rather than a mode.
   *
   * SECRET. Never logged, never put into an error message, never cached — the
   * caller owns this token's lifetime and this layer only forwards it. When
   * set, no client-credentials token is requested at all, so a delegated call
   * cannot silently fall back to the app identity that Graph would refuse.
   */
  readonly bearerToken?: string;
  /** Override the resource-default request budget (ms). */
  readonly timeoutMs?: number;
  /**
   * ARM long-running mode: on 201/202 follow the `Azure-AsyncOperation`
   * header (or the resource's `provisioningState`) until terminal.
   */
  readonly pollLongRunning?: boolean;
  /**
   * Extra statuses treated as success, e.g. `[404]` on a DELETE rollback
   * (already gone = rolled back). Keeps callers from string-matching errors.
   */
  readonly extraOkStatuses?: readonly number[];
  /**
   * Non-409 answers that STILL mean "the object already exists" and must
   * surface as {@link ProvisioningConflictResponse} rather than an error.
   * The matching stays HERE — the one place that owns the status matrix —
   * so step units keep branching on `kind === 'conflict'` only.
   */
  readonly conflictOn?: readonly ProvisioningConflictRule[];
}

/**
 * One "this is really a duplicate" signature: a status plus optional Graph
 * `error.code` values and/or `error.message` substrings (matched
 * case-insensitively). A rule fires when the status matches AND every
 * constraint it declares matches — a rule with no constraint fires on the
 * status alone, so keep those narrow.
 */
export interface ProvisioningConflictRule {
  readonly status: number;
  /** Qualifying `error.code` values, e.g. `['ObjectConflict']`. */
  readonly codes?: readonly string[];
  /** Qualifying `error.message` substrings (lower-cased comparison). */
  readonly messageIncludes?: readonly string[];
}

/** 2xx (or extraOkStatuses) outcome. `json` is undefined for empty bodies. */
export interface ProvisioningOkResponse {
  readonly kind: 'ok';
  readonly status: number;
  readonly json: unknown;
  readonly header: (name: string) => string | null;
}

/**
 * Duplicate outcome — the idempotent "already exists" signal, NOT an error.
 * Usually a 409; `status` is the raw status because a {@link
 * ProvisioningConflictRule} can promote another one (Entra's 400 on a taken
 * `uniqueName`) onto this same branch.
 */
export interface ProvisioningConflictResponse {
  readonly kind: 'conflict';
  readonly status: number;
  readonly json: unknown;
}

export type ProvisioningResponse =
  | ProvisioningOkResponse
  | ProvisioningConflictResponse;

/**
 * What one HTTP exchange settled on: the raw `Response` for every status the
 * caller treats as non-error, or an already-parsed conflict for a status a
 * {@link ProvisioningConflictRule} promoted (its body had to be read to
 * inspect the error code, and a body can only be read once).
 */
type FetchOutcome =
  | { readonly kind: 'response'; readonly response: Response }
  | { readonly kind: 'conflict'; readonly conflict: ProvisioningConflictResponse };

const defaultSleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms));

export class ProvisioningHttp {
  private readonly graphCredential: AadClientCredential;
  private readonly armCredential: AadClientCredential;
  private readonly log: (msg: string) => void;
  private readonly fetchImpl: typeof fetch;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly max429Retries: number;
  private readonly requestTimeoutMs: Record<ProvisioningResource, number>;
  private readonly pollIntervalMs: number;
  private readonly maxPollAttempts: number;
  private readonly tokens: AadTokenCache;

  constructor(opts: ProvisioningHttpOptions) {
    this.graphCredential = opts.graphCredential;
    this.armCredential = opts.armCredential ?? opts.graphCredential;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
    this.fetchImpl = opts.fetchImpl ?? fetch;
    this.sleep = opts.sleep ?? defaultSleep;
    this.max429Retries = opts.max429Retries ?? DEFAULT_MAX_429_RETRIES;
    this.requestTimeoutMs = {
      ...DEFAULT_REQUEST_TIMEOUT_MS,
      ...opts.requestTimeoutMs,
    };
    this.pollIntervalMs = opts.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS;
    this.maxPollAttempts = opts.maxPollAttempts ?? DEFAULT_MAX_POLL_ATTEMPTS;
    this.tokens = new AadTokenCache({ fetchImpl: this.fetchImpl });
  }

  /** The single choke point — every Graph/ARM provisioning verb goes here. */
  async request(req: ProvisioningRequest): Promise<ProvisioningResponse> {
    const outcome = await this.fetchWithRetry(req);

    // A rule-matched conflict already consumed the body (it had to, to read
    // the error code) — hand the parsed payload straight through.
    if (outcome.kind === 'conflict') return outcome.conflict;

    const { response } = outcome;

    if (response.status === 409) {
      return {
        kind: 'conflict',
        status: 409,
        json: await safeJson(response),
      };
    }

    if (
      req.pollLongRunning &&
      (response.status === 201 || response.status === 202)
    ) {
      return this.pollUntilTerminal(req, response);
    }

    return {
      kind: 'ok',
      status: response.status,
      json: await safeJson(response),
      header: (name) => response.headers.get(name),
    };
  }

  /** Token per audience — cached per (tenant, clientId, scope) in AadTokenCache. */
  private accessToken(resource: ProvisioningResource): Promise<string> {
    const credential =
      resource === 'arm' ? this.armCredential : this.graphCredential;
    const scope = resource === 'arm' ? ARM_TOKEN_SCOPE : GRAPH_TOKEN_SCOPE;
    return this.tokens.accessToken({
      ...credential,
      scope,
      timeoutMs: TOKEN_TIMEOUT_MS,
    });
  }

  /**
   * One HTTP exchange incl. the 429 loop and 403 mapping. Returns the raw
   * response for every status the caller treats as non-error (2xx,
   * extraOkStatuses, 409); throws typed/generic errors for the rest.
   */
  private async fetchWithRetry(req: ProvisioningRequest): Promise<FetchOutcome> {
    let lastRetryAfterSeconds: number | undefined;

    for (let attempt = 0; ; attempt += 1) {
      // An explicit bearer token REPLACES the app-only one; it is never merged
      // with it and never cached (byte5ai/omadia#924).
      const token = req.bearerToken ?? (await this.accessToken(req.resource));
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
      };
      let body: RequestInit['body'];
      if (req.rawBody) {
        headers['content-type'] = req.rawBody.contentType;
        body = req.rawBody.bytes as unknown as RequestInit['body'];
      } else if (req.jsonBody !== undefined) {
        headers['content-type'] = 'application/json';
        body = JSON.stringify(req.jsonBody);
      }
      const response = await this.fetchImpl(req.url, {
        method: req.method,
        headers,
        ...(body !== undefined ? { body } : {}),
        signal: AbortSignal.timeout(
          req.timeoutMs ?? this.requestTimeoutMs[req.resource],
        ),
      });

      if (
        response.ok ||
        response.status === 409 ||
        (req.extraOkStatuses?.includes(response.status) ?? false)
      ) {
        return { kind: 'response', response };
      }

      // The body can be read exactly ONCE, and three branches below need it —
      // so read it here, on the error path only, and pass the text along.
      const bodyText = await safeBody(response);

      const conflict = matchConflictRule(req, response.status, bodyText);
      if (conflict !== undefined) {
        this.log(
          `provisioner ${req.step}: ${String(response.status)} reads as an already-exists conflict — continuing on the idempotent path`,
        );
        return { kind: 'conflict', conflict };
      }

      if (response.status === 403) {
        throw new ConsentMissingError(
          req.missingScopesOn403,
          req.resource,
          new Error(
            `${req.resource} ${req.step} 403 body=${safeDetail(bodyText)}`,
          ),
        );
      }

      if (response.status === 429) {
        lastRetryAfterSeconds = retryAfterSeconds(response) ?? lastRetryAfterSeconds;
        if (attempt >= this.max429Retries) {
          throw new ProvisioningThrottledError(
            req.resource,
            lastRetryAfterSeconds,
            new Error(
              `${req.resource} ${req.step} 429 after ${String(attempt + 1)} attempts body=${safeDetail(bodyText)}`,
            ),
          );
        }
        const delayMs = this.throttleDelayMs(response, attempt);
        if (delayMs > MAX_BACKOFF_MS) {
          // A wait beyond the cap is a "come back much later" signal, not a
          // backoff — surface it typed so the job runner can reschedule
          // instead of holding an awaited sleep for hours.
          throw new ProvisioningThrottledError(
            req.resource,
            lastRetryAfterSeconds,
            new Error(
              `${req.resource} ${req.step} 429 Retry-After ${String(delayMs)}ms exceeds the ${String(MAX_BACKOFF_MS)}ms backoff cap body=${safeDetail(bodyText)}`,
            ),
          );
        }
        this.log(
          `provisioner ${req.step}: 429, retrying in ${String(delayMs)}ms (attempt ${String(attempt + 1)}/${String(this.max429Retries + 1)})`,
        );
        await this.sleep(delayMs);
        continue;
      }

      throw new ProvisioningRequestError(
        req.resource,
        req.step,
        response.status,
        `${req.resource} ${req.step} ${String(response.status)} ${req.method} ${truncate(req.url, 120)} body=${safeDetail(bodyText)}`,
      );
    }
  }

  /** Retry-After when present (never a fixed sleep), else exponential fallback. */
  private throttleDelayMs(response: Response, attempt: number): number {
    const headerSeconds = retryAfterSeconds(response);
    if (headerSeconds !== undefined) return headerSeconds * 1000;
    return DEFAULT_BACKOFF_BASE_MS * 2 ** attempt;
  }

  /**
   * ARM long-running follow-up: prefer the `Azure-AsyncOperation` status
   * endpoint; fall back to `provisioningState` on the resource itself. The
   * spec text does not cover this, but ARM `PUT botServices` answers
   * 201/202 + async operation — a plain "ok" would report unprovisioned bots.
   */
  private async pollUntilTerminal(
    req: ProvisioningRequest,
    initial: Response,
  ): Promise<ProvisioningOkResponse> {
    const initialJson = await safeJson(initial);
    const asyncOperationUrl = initial.headers.get('azure-asyncoperation');
    const locationUrl = initial.headers.get('location');
    const pollUrl = asyncOperationUrl ?? locationUrl ?? req.url;
    /**
     * Which body SHAPE the poll target answers with: the
     * `Azure-AsyncOperation` endpoint returns an operation-status document
     * (top-level `status`), while a `Location` target (ARM operation-results)
     * and the resource URL itself return the RESOURCE representation
     * (`properties.provisioningState`). The parser is selected by this kind
     * — with the other shape accepted as fallback — so a Location-only 202
     * can terminate instead of polling forever on the wrong parser.
     */
    const pollKind: 'operation' | 'resource' =
      asyncOperationUrl !== null ? 'operation' : 'resource';

    if (asyncOperationUrl === null && locationUrl === null) {
      const state = provisioningState(initialJson);
      if (state === undefined || state.toLowerCase() === 'succeeded') {
        return {
          kind: 'ok',
          status: initial.status,
          json: initialJson,
          header: (name) => initial.headers.get(name),
        };
      }
    }

    const initialRetryAfter = retryAfterSeconds(initial);
    let nextWaitMs =
      initialRetryAfter !== undefined
        ? Math.min(initialRetryAfter * 1000, MAX_BACKOFF_MS)
        : this.pollIntervalMs;

    for (let attempt = 0; attempt < this.maxPollAttempts; attempt += 1) {
      await this.sleep(nextWaitMs);
      nextWaitMs = this.pollIntervalMs;

      const token = req.bearerToken ?? (await this.accessToken(req.resource));
      const poll = await this.fetchImpl(pollUrl, {
        method: 'GET',
        headers: { Authorization: `Bearer ${token}` },
        signal: AbortSignal.timeout(
          req.timeoutMs ?? this.requestTimeoutMs[req.resource],
        ),
      });

      if (poll.status === 429 || poll.status === 202) {
        // Not terminal yet — honour the API's own pacing hint when present,
        // clamped so an hours-long hint cannot hang the provisioning job.
        const hinted = retryAfterSeconds(poll);
        if (hinted !== undefined) {
          nextWaitMs = Math.min(hinted * 1000, MAX_BACKOFF_MS);
        }
        continue;
      }
      if (!poll.ok) {
        const bodyText = await safeBody(poll);
        throw new ProvisioningRequestError(
          req.resource,
          `${req.step} poll`,
          poll.status,
          `${req.resource} ${req.step} poll ${String(poll.status)} body=${safeDetail(bodyText)}`,
        );
      }

      const json = await safeJson(poll);
      const status =
        pollKind === 'operation'
          ? (operationStatus(json) ?? provisioningState(json))
          : (provisioningState(json) ?? operationStatus(json));
      const terminal = status?.toLowerCase();
      if (terminal === 'succeeded') {
        // Resource-shaped polls (Location target or the resource URL itself)
        // already carry the finished resource — the poll body IS the
        // freshest representation. An operation-status body is not the
        // resource, and the pre-poll PUT body is known-stale by the time the
        // operation succeeds, so that path always re-GETs the resource.
        const finalJson =
          pollKind === 'resource' ? json : await this.fetchResource(req);
        return {
          kind: 'ok',
          status: initial.status,
          json: finalJson,
          header: (name) => initial.headers.get(name),
        };
      }
      if (terminal === 'failed' || terminal === 'canceled') {
        throw new Error(
          `${req.resource} ${req.step} long-running operation ${terminal} body=${safeDetail(JSON.stringify(json ?? ''))}`,
        );
      }
    }

    throw new Error(
      `${req.resource} ${req.step} long-running operation still pending after ${String(this.maxPollAttempts)} polls`,
    );
  }

  /** Final GET of the just-provisioned resource (bodiless 202 async path). */
  private async fetchResource(req: ProvisioningRequest): Promise<unknown> {
    const token = req.bearerToken ?? (await this.accessToken(req.resource));
    const response = await this.fetchImpl(req.url, {
      method: 'GET',
      headers: { Authorization: `Bearer ${token}` },
      signal: AbortSignal.timeout(
        req.timeoutMs ?? this.requestTimeoutMs[req.resource],
      ),
    });
    if (!response.ok) {
      const bodyText = await safeBody(response);
      throw new ProvisioningRequestError(
        req.resource,
        `${req.step} final read`,
        response.status,
        `${req.resource} ${req.step} final read ${String(response.status)} body=${safeDetail(bodyText)}`,
      );
    }
    return safeJson(response);
  }
}

/** Parse `Retry-After` — delta-seconds or HTTP-date — into whole seconds. */
function retryAfterSeconds(response: Response): number | undefined {
  const raw = response.headers.get('retry-after');
  if (raw === null || raw.trim() === '') return undefined;
  const asInt = Number.parseInt(raw, 10);
  if (Number.isFinite(asInt) && String(asInt) === raw.trim()) {
    return Math.max(0, asInt);
  }
  const asDate = Date.parse(raw);
  if (Number.isFinite(asDate)) {
    return Math.max(0, Math.ceil((asDate - Date.now()) / 1000));
  }
  return undefined;
}

/**
 * Does this error body match one of the request's extra conflict rules?
 * Returns the ready conflict response, or `undefined` to keep the normal
 * error path.
 */
function matchConflictRule(
  req: ProvisioningRequest,
  status: number,
  bodyText: string,
): ProvisioningConflictResponse | undefined {
  const rules = req.conflictOn;
  if (rules === undefined || rules.length === 0) return undefined;
  const candidates = rules.filter((rule) => rule.status === status);
  if (candidates.length === 0) return undefined;

  const json = parseJson(bodyText);
  const code = graphErrorField(json, 'code')?.toLowerCase();
  const message = graphErrorField(json, 'message')?.toLowerCase();

  for (const rule of candidates) {
    if (
      rule.codes !== undefined &&
      !rule.codes.some((c) => c.toLowerCase() === code)
    ) {
      continue;
    }
    if (
      rule.messageIncludes !== undefined &&
      !rule.messageIncludes.some((m) => message?.includes(m.toLowerCase()) === true)
    ) {
      continue;
    }
    return { kind: 'conflict', status, json };
  }
  return undefined;
}

/** `error.code` / `error.message` of a Graph/ARM error envelope, if present. */
function graphErrorField(json: unknown, field: 'code' | 'message'): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const error = (json as Record<string, unknown>)['error'];
  if (!error || typeof error !== 'object') return undefined;
  const value = (error as Record<string, unknown>)[field];
  return typeof value === 'string' ? value : undefined;
}

function parseJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

async function safeJson(response: Response): Promise<unknown> {
  try {
    return await response.json();
  } catch {
    return undefined;
  }
}

function provisioningState(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const properties = (json as Record<string, unknown>)['properties'];
  if (!properties || typeof properties !== 'object') return undefined;
  const state = (properties as Record<string, unknown>)['provisioningState'];
  return typeof state === 'string' ? state : undefined;
}

function operationStatus(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const status = (json as Record<string, unknown>)['status'];
  return typeof status === 'string' ? status : undefined;
}

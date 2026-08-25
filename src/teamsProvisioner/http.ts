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
 *   signal the step units map to `Idempotent<T>` `'already-existed'`.
 * - 429 → backoff honouring the `Retry-After` header (never a fixed sleep);
 *   exhausted budget → {@link ProvisioningThrottledError}.
 *
 * Tokens come from the shared {@link AadTokenCache} (the generalised
 * `GraphClient` plumbing) — cached per (tenant, clientId, scope), so the two
 * audiences (`graph.microsoft.com/.default` vs `management.azure.com/.default`)
 * and the optional dedicated ARM service principal never collide.
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
import { ConsentMissingError, ProvisioningThrottledError } from './errors.js';

/** Which API a request targets — also the audience key for token + errors. */
export type ProvisioningResource = 'graph' | 'arm';

/** Token audience for ARM (`management.azure.com`) requests. */
const ARM_TOKEN_SCOPE = 'https://management.azure.com/.default';

const DEFAULT_MAX_429_RETRIES = 3;
/** Fallback backoff base when a 429 arrives WITHOUT Retry-After (doubles per attempt). */
const DEFAULT_BACKOFF_BASE_MS = 1000;
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
}

/** 2xx (or extraOkStatuses) outcome. `json` is undefined for empty bodies. */
export interface ProvisioningOkResponse {
  readonly kind: 'ok';
  readonly status: number;
  readonly json: unknown;
  readonly header: (name: string) => string | null;
}

/** 409 outcome — the idempotent "already exists" signal, NOT an error. */
export interface ProvisioningConflictResponse {
  readonly kind: 'conflict';
  readonly status: 409;
  readonly json: unknown;
}

export type ProvisioningResponse =
  | ProvisioningOkResponse
  | ProvisioningConflictResponse;

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
    const response = await this.fetchWithRetry(req);

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
  private async fetchWithRetry(req: ProvisioningRequest): Promise<Response> {
    let lastRetryAfterSeconds: number | undefined;

    for (let attempt = 0; ; attempt += 1) {
      const token = await this.accessToken(req.resource);
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
        return response;
      }

      if (response.status === 403) {
        const bodyText = await safeBody(response);
        throw new ConsentMissingError(
          req.missingScopesOn403,
          req.resource,
          new Error(
            `${req.resource} ${req.step} 403 body=${truncate(bodyText, 200)}`,
          ),
        );
      }

      if (response.status === 429) {
        lastRetryAfterSeconds = retryAfterSeconds(response) ?? lastRetryAfterSeconds;
        if (attempt >= this.max429Retries) {
          const bodyText = await safeBody(response);
          throw new ProvisioningThrottledError(
            req.resource,
            lastRetryAfterSeconds,
            new Error(
              `${req.resource} ${req.step} 429 after ${String(attempt + 1)} attempts body=${truncate(bodyText, 200)}`,
            ),
          );
        }
        const delayMs = this.throttleDelayMs(response, attempt);
        this.log(
          `provisioner ${req.step}: 429, retrying in ${String(delayMs)}ms (attempt ${String(attempt + 1)}/${String(this.max429Retries + 1)})`,
        );
        await this.sleep(delayMs);
        continue;
      }

      const bodyText = await safeBody(response);
      throw new Error(
        `${req.resource} ${req.step} ${String(response.status)} ${req.method} ${truncate(req.url, 120)} body=${truncate(bodyText, 200)}`,
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
    const asyncUrl =
      initial.headers.get('azure-asyncoperation') ??
      initial.headers.get('location');

    if (!asyncUrl) {
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
        ? initialRetryAfter * 1000
        : this.pollIntervalMs;

    for (let attempt = 0; attempt < this.maxPollAttempts; attempt += 1) {
      await this.sleep(nextWaitMs);
      nextWaitMs = this.pollIntervalMs;

      const pollUrl = asyncUrl ?? req.url;
      const token = await this.accessToken(req.resource);
      const poll = await this.fetchImpl(pollUrl, {
        method: 'GET',
        headers: { Authorization: `Bearer ${token}` },
        signal: AbortSignal.timeout(
          req.timeoutMs ?? this.requestTimeoutMs[req.resource],
        ),
      });

      if (poll.status === 429 || poll.status === 202) {
        // Not terminal yet — honour the API's own pacing hint when present.
        const hinted = retryAfterSeconds(poll);
        if (hinted !== undefined) nextWaitMs = hinted * 1000;
        continue;
      }
      if (!poll.ok) {
        const bodyText = await safeBody(poll);
        throw new Error(
          `${req.resource} ${req.step} poll ${String(poll.status)} body=${truncate(bodyText, 200)}`,
        );
      }

      const json = await safeJson(poll);
      const status = asyncUrl ? operationStatus(json) : provisioningState(json);
      const terminal = status?.toLowerCase();
      if (terminal === 'succeeded') {
        // The PUT response usually carries the resource representation; for
        // the resource-URL fallback the freshest body is the poll itself. A
        // bodiless 202 + async operation re-GETs the finished resource.
        const finalJson =
          asyncUrl === null
            ? json
            : (initialJson ?? (await this.fetchResource(req)));
        return {
          kind: 'ok',
          status: initial.status,
          json: finalJson,
          header: (name) => initial.headers.get(name),
        };
      }
      if (terminal === 'failed' || terminal === 'canceled') {
        throw new Error(
          `${req.resource} ${req.step} long-running operation ${terminal} body=${truncate(JSON.stringify(json ?? ''), 200)}`,
        );
      }
    }

    throw new Error(
      `${req.resource} ${req.step} long-running operation still pending after ${String(this.maxPollAttempts)} polls`,
    );
  }

  /** Final GET of the just-provisioned resource (bodiless 202 async path). */
  private async fetchResource(req: ProvisioningRequest): Promise<unknown> {
    const token = await this.accessToken(req.resource);
    const response = await this.fetchImpl(req.url, {
      method: 'GET',
      headers: { Authorization: `Bearer ${token}` },
      signal: AbortSignal.timeout(
        req.timeoutMs ?? this.requestTimeoutMs[req.resource],
      ),
    });
    if (!response.ok) {
      const bodyText = await safeBody(response);
      throw new Error(
        `${req.resource} ${req.step} final read ${String(response.status)} body=${truncate(bodyText, 200)}`,
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

/**
 * App-only Graph transport for the facilitation follow-up service.
 *
 * The transport cannot express anything but the three follow-up calls: it
 * takes an operation name, not a URL, and every operation is a POST against
 * the organiser mailbox. There is no GET here and no path into `/messages`,
 * `/events/{id}`, `/calendarView` or `/mailFolders`, so "never reads mailbox
 * or calendar contents" is a property of this file and not a promise the
 * service has to keep call by call.
 *
 * Tokens come from the client-credentials flow (application permissions). The
 * delegated clients (`graphObo.ts`, `graphCalendarClient.ts`, `/me/...`) are
 * not involved: a facilitator acts for no user.
 */

import {
  AadTokenCache,
  GRAPH_TOKEN_SCOPE,
  type AadClientCredential,
} from '../graphClient.js';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const DEFAULT_REQUEST_TIMEOUT_MS = 15_000;
/** Follow-ups run after a facilitation closed, not on the Teams hot path. */
const TOKEN_TIMEOUT_MS = 10_000;

/** The complete set of Graph calls this service can make. */
export type FollowupGraphOperation = 'getSchedule' | 'createEvent' | 'sendMail';

/** Path below `/users/{organiser}` for each operation. All are POSTs. */
const OPERATION_PATHS: Readonly<Record<FollowupGraphOperation, string>> = {
  getSchedule: '/calendar/getSchedule',
  createEvent: '/events',
  sendMail: '/sendMail',
};

export interface FollowupGraphRequest {
  readonly operation: FollowupGraphOperation;
  /** The organiser mailbox the call is made on. */
  readonly mailbox: string;
  readonly body: unknown;
  readonly headers?: Readonly<Record<string, string>>;
}

export interface FollowupGraphResponse {
  readonly status: number;
  /** Parsed JSON body, or undefined for an empty or non-JSON body. */
  readonly json: unknown;
}

export interface FollowupGraphTransport {
  post(request: FollowupGraphRequest): Promise<FollowupGraphResponse>;
}

export interface AppOnlyGraphTransportOptions {
  readonly credential: AadClientCredential;
  /** Test seam, same injection point as `GraphClient`. */
  readonly fetchImpl?: typeof fetch;
  readonly requestTimeoutMs?: number;
}

export function followupGraphUrl(
  operation: FollowupGraphOperation,
  mailbox: string,
): string {
  return `${GRAPH_BASE_URL}/users/${encodeURIComponent(mailbox)}${OPERATION_PATHS[operation]}`;
}

export function createAppOnlyGraphTransport(
  opts: AppOnlyGraphTransportOptions,
): FollowupGraphTransport {
  const fetchImpl = opts.fetchImpl ?? fetch;
  const tokens = new AadTokenCache({ fetchImpl });
  const requestTimeoutMs = opts.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS;

  return {
    async post(request: FollowupGraphRequest): Promise<FollowupGraphResponse> {
      const token = await tokens.accessToken({
        ...opts.credential,
        scope: GRAPH_TOKEN_SCOPE,
        timeoutMs: TOKEN_TIMEOUT_MS,
      });
      const response = await fetchImpl(
        followupGraphUrl(request.operation, request.mailbox),
        {
          method: 'POST',
          headers: {
            ...request.headers,
            Authorization: `Bearer ${token}`,
            'content-type': 'application/json',
          },
          body: JSON.stringify(request.body),
          signal: AbortSignal.timeout(requestTimeoutMs),
        },
      );
      return { status: response.status, json: await readJson(response) };
    },
  };
}

async function readJson(response: Response): Promise<unknown> {
  const text = await response.text();
  if (text.length === 0) return undefined;
  try {
    return JSON.parse(text) as unknown;
  } catch {
    // A non-JSON body (gateway error page) carries nothing the service reads.
    return undefined;
  }
}

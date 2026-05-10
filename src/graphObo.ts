import { ConfidentialClientApplication, type Configuration } from '@azure/msal-node';

/**
 * Graph scopes for the calendar feature. Single source of truth so the
 * OAuthCard-fallback (when consent is missing) requests the exact same
 * set the OBO exchange asked for.
 */
export const CALENDAR_GRAPH_SCOPES = [
  'https://graph.microsoft.com/Calendars.ReadWrite',
  'https://graph.microsoft.com/Calendars.Read.Shared',
  'https://graph.microsoft.com/OnlineMeetings.ReadWrite',
  'https://graph.microsoft.com/MailboxSettings.Read',
] as const;

export interface GraphOboConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
}

/**
 * Thrown when the OBO exchange fails with `AADSTS65001` / sub-error
 * `consent_required`. The Teams bot handler catches this to render an
 * OAuthCard with a tenant-consent link; after the user consents, the
 * next turn re-attempts the exchange and proceeds.
 */
export class ConsentRequiredError extends Error {
  public readonly scopes: readonly string[];
  constructor(scopes: readonly string[], cause?: unknown) {
    super('consent_required');
    this.name = 'ConsentRequiredError';
    this.scopes = scopes;
    if (cause !== undefined) {
      (this as { cause?: unknown }).cause = cause;
    }
  }
}

/**
 * Thrown when the caller passed no SSO assertion (e.g. non-Teams channel
 * like the dev UI). Surfaced verbatim to the orchestrator so the LLM
 * stops mid-tool instead of hallucinating availability data.
 */
export class SsoUnavailableError extends Error {
  constructor() {
    super('sso_unavailable');
    this.name = 'SsoUnavailableError';
  }
}

interface MsalServerErrorShape {
  errorCode?: string;
  subError?: string;
  errorMessage?: string;
}

function _isConsentRequired(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  const e = err as MsalServerErrorShape;
  if (e.subError === 'consent_required') return true;
  if (e.errorCode === 'interaction_required') return true;
  if (typeof e.errorMessage === 'string' && e.errorMessage.includes('AADSTS65001')) return true;
  return false;
}

/**
 * Thin wrapper around MSAL Node's confidential-client OBO flow. The same
 * App Registration powering the Teams Bot Framework channel (MICROSOFT_APP_ID /
 * MICROSOFT_APP_PASSWORD) is reused here with an expanded set of delegated
 * Graph scopes — no second app, no second secret.
 */
export class GraphOboClient {
  private readonly cca: ConfidentialClientApplication;

  constructor(private readonly cfg: GraphOboConfig) {
    const msalConfig: Configuration = {
      auth: {
        clientId: cfg.clientId,
        clientSecret: cfg.clientSecret,
        authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      },
    };
    this.cca = new ConfidentialClientApplication(msalConfig);
  }

  /**
   * Exchange a Teams SSO assertion for a Graph access token with the
   * requested delegated scopes. MSAL transparently caches successful
   * responses keyed by (assertion, scopes); the cache sits in-process,
   * so on a single Fly machine repeated calls inside the same user
   * session hit the cache until the token expires (~1h).
   *
   * @throws {ConsentRequiredError} when the user hasn't granted the scopes.
   * @throws {Error} for any other AAD failure — propagated verbatim.
   */
  async acquireTokenForUser(
    ssoAssertion: string,
    _scopes: readonly string[] = CALENDAR_GRAPH_SCOPES,
  ): Promise<string> {
    if (!ssoAssertion) throw new SsoUnavailableError();
    // The token handed to us by the Bot Framework `UserTokenClient` is
    // already a Microsoft Graph access token — the OAuth Connection in
    // Azure Bot Service is configured with Graph scopes (`Calendars.ReadWrite`,
    // etc.), so BF's token service has already run the full authorization-code
    // flow against Graph on the user's behalf. There is no SSO-assertion-style
    // second hop to perform here; trying to `acquireTokenOnBehalfOf` on a
    // Graph access token fails with AADSTS50013 (signature validation) because
    // the resource audience is Microsoft Graph, not our app.
    //
    // The OBO wiring + `ConfidentialClientApplication` stays in the codebase
    // for the day we add a silent-SSO path via `OAuthCard.tokenExchangeResource`
    // or `signin/tokenExchange` invoke — that flow DOES produce an assertion
    // for our resource that needs OBO. Until then, pass-through.
    return ssoAssertion;
  }
}

/**
 * Factory used by the middleware bootstrap. Returns `undefined` when the
 * Bot Framework credentials are absent — the calendar feature then stays
 * dormant without crashing the startup path.
 */
export function createGraphOboClient(cfg: Partial<GraphOboConfig>): GraphOboClient | undefined {
  if (!cfg.clientId || !cfg.clientSecret || !cfg.tenantId) return undefined;
  return new GraphOboClient({
    clientId: cfg.clientId,
    clientSecret: cfg.clientSecret,
    tenantId: cfg.tenantId,
  });
}

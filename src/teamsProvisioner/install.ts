/**
 * Install step of `teamsProvisioner@1` (epic byte5ai/omadia#860, capability
 * issue byte5ai/omadia-m365-connector#3): install a catalog Teams app into ONE
 * team — the last link of the "1 agent = 1 Entra app + 1 Azure bot + 1 Teams
 * app package" chain — and, since 0.4.0, take it back out again
 * ({@link TeamInstallClient.uninstallFromTeam}, byte5ai/omadia#900).
 *
 * SINCE 0.7.0, ALSO INTO A CHAT ({@link ChatInstallClient}). A team was never
 * the only place an agent belongs: the common case is a group chat, which had
 * no path here at all. `POST /chats/{id}/installedApps` is the chat-scope
 * counterpart and — unlike the catalog upload — it DOES support application
 * permissions, so the chat direction needs no delegated sign-in; it needs the
 * `TeamsAppInstallation.ReadWriteForChat.All` app role
 * ({@link CHAT_INSTALL_SCOPE}) and nothing else.
 *
 * The two directions are near-twins and deliberately do not share a class:
 * see {@link ChatInstallClient} for what actually differs (app role,
 * not-found meaning, the pre-flight target check). What they DO share — the
 * installation lookup, OData escaping, result assembly — lives in the
 * module-private helpers at the bottom.
 *
 * One Graph write: `POST /teams/{teamId}/installedApps` with the
 * `teamsApp@odata.bind` reference to the catalog app, plus — since 0.8.2, in
 * BOTH directions — the {@link ConsentedPermissionSet} the app package itself
 * declares. It is included only when there is one: the body shape changes with
 * it, so an absent set must not serialise as `consentedPermissionSet:
 * undefined`/`null`.
 *
 * WHERE THAT SET COMES FROM (0.8.2, the field failure this release fixes).
 * Graph requires the installer to state the resource-specific permissions it
 * consents to whenever the app declares any, and requires them to MATCH the
 * app's own `teamsAppDefinition` — "the permissions consented to during the
 * installation must match the resource-specific permissions defined in the
 * teamsAppDefinition of the app". So the set is neither a fixed list nor
 * something a caller should have to retype: it is read back from Graph's own
 * record of the app being installed
 * ({@link resolveDeclaredPermissionSet}, the query Graph's own docs point at).
 * A caller MAY still pass one explicitly — an explicit set wins and skips the
 * lookup — but nothing has to.
 *
 * Before 0.8.2 the team direction offered the field and no caller ever filled
 * it, and the chat direction did not offer it at all. Our generated packages
 * declare seven RSC permissions, so BOTH directions were refused with
 * `400 ResourceSpecificPermissionsMismatch` the moment the tenant had the
 * consent-capable app role — the team direction latently, because nobody had
 * run it yet.
 *
 * IDEMPOTENCY — 409 is success. Graph answers 409 when the app is already
 * installed in the team (idempotency key: (teamId, teamsAppId)). The shared
 * http layer maps that to its `{ kind: 'conflict' }` signal and this module
 * maps it onward to `Idempotent<T>` `'already-existed'` — callers branch on
 * `outcome`, nobody string-matches Graph error bodies (see `types.ts`).
 *
 * The uninstall direction mirrors that: Graph deletes an installation by its
 * INSTALLATION id, not by the catalog app id, so the removal is a
 * lookup-then-DELETE pair — and both halves of the "not installed" case (the
 * lookup finds nothing, or the DELETE races another remover into a 404)
 * collapse into the single idempotent `'already-absent'` outcome. That
 * literal deviates from the `'already-deleted'` of the app-registration /
 * bot rollbacks on purpose: nothing is destroyed here, an app that was never
 * in the team is simply absent (byte5ai/omadia#900 names the signal).
 *
 * All HTTP goes through the shared {@link ProvisioningHttp} choke point (one
 * token cache, Retry-After-honouring 429 backoff → `ProvisioningThrottledError`
 * when exhausted, 403 → `ConsentMissingError` carrying
 * {@link TEAM_INSTALL_SCOPE}) — the same one-fallback-branch family as the
 * catalog-upload step. This module opens no second token cache and does no
 * fetch of its own.
 */

import {
  ChatNotFoundError,
  InstallTargetMismatchError,
  ProvisioningRequestError,
  RscPermissionsMismatchError,
} from './errors.js';
import type {
  ProvisioningConflictRule,
  ProvisioningHttp,
  ProvisioningResponse,
} from './http.js';
import {
  classifyInstallTarget,
  isChatTarget,
  type AmbiguousTarget,
  type ChannelTarget,
  type TeamTarget,
  type UnknownTarget,
} from './installTarget.js';
import type {
  ChatAppInstallation,
  Idempotent,
  InstallToChatInput,
  InstallToTeamInput,
  TeamAppInstallation,
} from './types.js';

/**
 * Graph APPLICATION permission this step needs. Documented in the
 * scopes/consent unit and granted by the wiring unit's manifest bump —
 * surfaced on 403 via `ConsentMissingError.missingScopes`.
 */
export const TEAM_INSTALL_SCOPE = 'TeamsAppInstallation.ReadWriteForTeam.All';

/**
 * The app role a TEAM install needs ON TOP of {@link TEAM_INSTALL_SCOPE} once
 * the package declares resource-specific permissions.
 *
 * Graph is explicit that the plain `…ReadWriteForTeam.All` "cannot be used to
 * install apps that require consent to resource-specific consent permissions".
 * Reported on 403 (and named in {@link RscPermissionsMismatchError}) only when
 * a set is actually being sent — a plain install must keep pointing at the
 * plain role, or an operator grants a wider role than the install needs.
 */
export const TEAM_INSTALL_CONSENT_SCOPE =
  'TeamsAppInstallation.ReadWriteAndConsentForTeam.All';

/**
 * Graph app role the catalog lookup behind {@link resolveDeclaredPermissionSet}
 * needs. The same role the catalog step already requires — the chain has
 * always held it by the time an install runs, which is why the resolution can
 * live inside the install step instead of becoming a caller obligation.
 */
const APP_CATALOG_SCOPE = 'AppCatalog.ReadWrite.All';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/** Graph `teamsAppInstallation.permissionType` values. */
export type ResourceSpecificPermissionType = 'application' | 'delegated';

/** One resource-specific (RSC) permission entry of a consented set. */
export interface ResourceSpecificPermission {
  /** RSC permission name, e.g. `ChannelMessage.Read.Group`. */
  readonly permissionValue: string;
  readonly permissionType: ResourceSpecificPermissionType;
}

/**
 * The RSC permissions the installing admin consents to
 * (`consentedPermissionSet` on the Graph install body). Optional — most
 * installs omit it entirely.
 */
export interface ConsentedPermissionSet {
  readonly resourceSpecificPermissions: readonly ResourceSpecificPermission[];
}

/**
 * Install request: the shared `InstallToTeamInput` of `types.ts` plus the
 * optional consented-permission set. The extension lives HERE (not in
 * `types.ts`) because only the install step knows about RSC consent; the
 * `TeamsProvisioner` interface stays satisfied since the extra field is
 * optional.
 */
export interface InstallToTeamRequest extends InstallToTeamInput {
  /** Sent to Graph verbatim when present; omitted from the body otherwise. */
  readonly consentedPermissionSet?: ConsentedPermissionSet;
}

/**
 * Input for the uninstall step — the SAME (teamId, teamsAppId) key the
 * install is idempotent on. Callers never have to have kept the Graph
 * installation id around; the step resolves it (see
 * {@link TeamInstallClient.uninstallFromTeam}).
 */
export interface UninstallFromTeamInput {
  readonly teamId: string;
  /** Catalog id (`CatalogTeamsApp.teamsAppId`) — NOT the installation id. */
  readonly teamsAppId: string;
}

/**
 * Idempotency signal of the uninstall direction: `'uninstalled'` when this
 * call removed the installation, `'already-absent'` when the app was not
 * installed in the team (lookup miss, or the DELETE answered 404 because
 * someone else got there first). Both are SUCCESS — callers branch on
 * `outcome` instead of string-matching Graph error bodies.
 */
export type UninstallFromTeamOutcome = 'uninstalled' | 'already-absent';

/**
 * Result of {@link TeamInstallClient.uninstallFromTeam}. `value` mirrors the
 * install result shape ({@link TeamAppInstallation}) so a caller can log the
 * removed installation the same way it logged the created one;
 * `installationId` is present whenever the lookup resolved one — it is absent
 * on a pure lookup miss, where there is nothing to name.
 */
export interface UninstallFromTeamResult {
  readonly outcome: UninstallFromTeamOutcome;
  readonly value: TeamAppInstallation;
}

export interface TeamInstallClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  readonly log?: (msg: string) => void;
}

/**
 * Install catalog Teams apps into teams. One step client per provisioner;
 * ordering across chain steps stays middleware-side (agent factory,
 * byte5ai/omadia#863-865).
 */
export class TeamInstallClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly log: (msg: string) => void;

  constructor(opts: TeamInstallClientOptions) {
    this.http = opts.http;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * `POST /teams/{teamId}/installedApps` — install the catalog app
   * (`teamsAppId` = `CatalogTeamsApp.teamsAppId`) into the team.
   *
   * - 2xx → `'created'`, with the Graph installation id when the response
   *   body carried one.
   * - 409 → `'already-existed'` (app already installed) — success, never an
   *   exception.
   * - 403 → `ConsentMissingError([TEAM_INSTALL_SCOPE], 'graph')` from the
   *   http layer, the same typed family as `uploadToCatalog`, so callers get
   *   ONE fallback branch.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   */
  async installToTeam(
    input: InstallToTeamRequest,
  ): Promise<Idempotent<TeamAppInstallation>> {
    const teamId = requireNonEmpty(input.teamId, 'teamId');
    const teamsAppId = requireNonEmpty(input.teamsAppId, 'teamsAppId');
    const step = 'teams.installedApps.add';

    const permissions = await resolveConsentedPermissionSet(this.http, {
      teamsAppId,
      explicit: input.consentedPermissionSet,
      step,
      log: this.log,
    });

    const response = await installRequest(this.http, {
      url: `${GRAPH_BASE}/teams/${encodeURIComponent(teamId)}/installedApps`,
      step,
      teamsAppId,
      permissions,
      baseScope: TEAM_INSTALL_SCOPE,
      consentScope: TEAM_INSTALL_CONSENT_SCOPE,
    });

    if (response.kind === 'conflict') {
      this.log(
        `provisioner teams.installedApps.add: app ${teamsAppId} already installed in team ${teamId} (409 → already-existed)`,
      );
      return {
        outcome: 'already-existed',
        value: installation(teamId, teamsAppId, response.json),
      };
    }

    return {
      outcome: 'created',
      value: installation(teamId, teamsAppId, response.json),
    };
  }

  /**
   * Remove a catalog app from ONE team — the reverse of
   * {@link installToTeam} (byte5ai/omadia#900).
   *
   * Graph deletes an installation by its INSTALLATION id, which the caller
   * generally does not hold (the install result's `installationId` is
   * optional, and an app installed outside omadia has none on our side), so
   * this is a two-step:
   *
   *   1. `GET /teams/{teamId}/installedApps?$expand=teamsApp&$filter=teamsApp/id eq '…'`
   *      — resolve the installation for the catalog app. No hit →
   *      `'already-absent'`, and NO delete is attempted.
   *   2. `DELETE /teams/{teamId}/installedApps/{installationId}` — 2xx →
   *      `'uninstalled'`; 404 (someone removed it between the two calls) →
   *      `'already-absent'`, via the http layer's `extraOkStatuses`, so a
   *      race never turns into an exception.
   *
   * - 403 → `ConsentMissingError([TEAM_INSTALL_SCOPE], 'graph')` from the
   *   http layer — the same typed family (and the same scope) as the install
   *   direction, so callers keep ONE fallback branch.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   */
  async uninstallFromTeam(
    input: UninstallFromTeamInput,
  ): Promise<UninstallFromTeamResult> {
    const teamId = requireNonEmpty(input.teamId, 'teamId');
    const teamsAppId = requireNonEmpty(input.teamsAppId, 'teamsAppId');

    const installationId = await this.findInstallationId(teamId, teamsAppId);
    if (installationId === undefined) {
      this.log(
        `provisioner teams.installedApps.remove: app ${teamsAppId} not installed in team ${teamId} (lookup miss → already-absent)`,
      );
      return { outcome: 'already-absent', value: { teamId, teamsAppId } };
    }

    const response = await this.http.request({
      resource: 'graph',
      method: 'DELETE',
      url: `${GRAPH_BASE}/teams/${encodeURIComponent(teamId)}/installedApps/${encodeURIComponent(installationId)}`,
      step: 'teams.installedApps.remove',
      missingScopesOn403: [TEAM_INSTALL_SCOPE],
      // Already gone = already removed. Keeps the delete race off the
      // exception path (same treatment as the rollback DELETEs).
      extraOkStatuses: [404],
    });

    const gone = response.kind === 'ok' && response.status === 404;
    if (gone) {
      this.log(
        `provisioner teams.installedApps.remove: installation ${installationId} in team ${teamId} already gone (404 → already-absent)`,
      );
    }

    return {
      outcome: gone ? 'already-absent' : 'uninstalled',
      value: { teamId, teamsAppId, installationId },
    };
  }

  /** See {@link findInstallationId} — the team-scoped installation lookup. */
  private findInstallationId(
    teamId: string,
    teamsAppId: string,
  ): Promise<string | undefined> {
    return findInstallationId(this.http, {
      collectionUrl: `${GRAPH_BASE}/teams/${encodeURIComponent(teamId)}/installedApps`,
      teamsAppId,
      step: 'teams.installedApps.lookup',
      scope: TEAM_INSTALL_SCOPE,
      subject: `team=${teamId}`,
    });
  }
}

/**
 * Graph APPLICATION permission the chat direction needs (since 0.7.0).
 *
 * NOT the `…SelfForChat.All` variant, which only lets an app install ITSELF —
 * the provisioner installs the per-agent app it generated, a different app
 * from the connector's own identity, so "self" never applies.
 *
 * Surfaced on 403 via `ConsentMissingError.missingScopes`.
 */
export const CHAT_INSTALL_SCOPE = 'TeamsAppInstallation.ReadWriteForChat.All';

/**
 * Chat twin of {@link TEAM_INSTALL_CONSENT_SCOPE} — the app role a chat
 * install needs once the package declares resource-specific permissions.
 *
 * This is the role the field test had already granted while the request body
 * still omitted the set, which is precisely why the install kept failing:
 * having the role is half of it, the caller must also SAY what it consents to.
 */
export const CHAT_INSTALL_CONSENT_SCOPE =
  'TeamsAppInstallation.ReadWriteAndConsentForChat.All';

/**
 * Input for {@link ChatInstallClient.installToChat}: the shared
 * `InstallToChatInput` of `types.ts` plus the optional consented-permission
 * set — the exact shape {@link InstallToTeamRequest} has, deliberately, so the
 * two directions stay one mechanism rather than two.
 *
 * OPTIONAL BECAUSE IT IS RESOLVED, NOT BECAUSE IT IS RARE. Left out, the step
 * reads the app's declared permissions from Graph itself
 * ({@link resolveDeclaredPermissionSet}); passed in, the caller's set is sent
 * verbatim and no lookup happens.
 */
export interface InstallToChatRequest extends InstallToChatInput {
  /** Sent to Graph verbatim when present; resolved from the app otherwise. */
  readonly consentedPermissionSet?: ConsentedPermissionSet;
}

/**
 * How the chat endpoint reports "this app is already installed here" — and
 * what is actually KNOWN about that, which is less than one would like.
 *
 * WHAT IS ESTABLISHED. `409` with `error.code` **`Conflict`** is the duplicate
 * signal for the sibling scopes, and the http layer already treats 409 as its
 * NATIVE conflict — no rule needed for it. The message body reads
 * `AppEntitlement id: '<app>' already exists in <scope>`, and the SCOPE TAIL
 * VARIES: `… already exists in TeamId: '19:…'` for the team endpoint,
 * `… already exists in thread` for the per-user endpoint.
 *
 * WHAT IS NOT. Graph's reference for this verb documents no failure response
 * at all — only `201 Created` — and a search across Microsoft Q&A, Stack
 * Overflow and GitHub turns up no report of the CHAT variant's duplicate
 * answer specifically. So the 409 is an inference from two neighbouring
 * scopes, not an observed fact about this one.
 *
 * WHY A 400 RULE ANYWAY. `teams`, `chats` and `users/…/teamwork` are separate
 * endpoints that have already been shown to word the same condition
 * differently, and Entra set the precedent that a duplicate can arrive
 * Bad-Request-shaped: a taken `uniqueName` on `POST /applications` answers
 * **400** rather than 409 (byte5ai/omadia#916). A re-run of a provisioning
 * chain hard-failing on an app that is simply already there is the expensive
 * mistake here, so the 400 shapes fold onto the same idempotent path.
 *
 * THE SUBSTRINGS ARE THE OBSERVED ONES. `'already exists'` is the wording both
 * documented scopes use; `'already installed'` is a cheap second guess. What
 * is deliberately NOT here: `NameAlreadyExists` and `AppAlreadyInstalled` —
 * neither error code exists anywhere in this API surface, so matching them
 * would be decoration. Nor `AddAppBotToChatRosterFailed`, which is an HTTP 500
 * from the Teams web client's own internal API, never a Graph answer.
 */
const CHAT_ALREADY_INSTALLED_RULES: readonly ProvisioningConflictRule[] = [
  { status: 400, codes: ['Conflict'] },
  { status: 400, messageIncludes: ['already exists', 'already installed'] },
];

/** Input for {@link ChatInstallClient.uninstallFromChat}. */
export interface UninstallFromChatInput {
  readonly chatId: string;
  /** Catalog id (`CatalogTeamsApp.teamsAppId`) — NOT the installation id. */
  readonly teamsAppId: string;
}

/**
 * Idempotency signal of the chat uninstall — the same vocabulary as
 * {@link UninstallFromTeamOutcome}, deliberately: a consumer that already
 * branches on `'uninstalled' | 'already-absent'` for teams needs no second
 * shape for chats.
 */
export type UninstallFromChatOutcome = 'uninstalled' | 'already-absent';

/** Result of {@link ChatInstallClient.uninstallFromChat}. */
export interface UninstallFromChatResult {
  readonly outcome: UninstallFromChatOutcome;
  readonly value: ChatAppInstallation;
}

export interface ChatInstallClientOptions {
  /**
   * The SHARED token/429 choke point of the http unit. Pass the one
   * `ProvisioningHttp` instance of the provisioner — this module must never
   * open a second token cache.
   */
  readonly http: Pick<ProvisioningHttp, 'request'>;
  readonly log?: (msg: string) => void;
}

/**
 * Install catalog Teams apps into CHATS — group chats and 1:1 chats (since
 * 0.7.0). The chat-scope twin of {@link TeamInstallClient}; one step client
 * per provisioner, ordering stays middleware-side.
 *
 * WHY A SECOND CLIENT AND NOT A FLAG. The two directions share their body
 * shape and their idempotency story but nothing else that matters: a different
 * Graph collection, a different app role, a different not-found meaning, and a
 * pre-flight target check that only the chat side needs. A `scope: 'team' |
 * 'chat'` parameter would have to branch on all four inside every method.
 * The genuinely shared parts — the installation lookup, the OData escaping,
 * the result assembly — are module-private functions both clients call.
 *
 * `consentedPermissionSet` IS SENT HERE, exactly as {@link installToTeam}
 * sends it (since 0.8.2). Until then it was left out on the reading that
 * `TeamsAppInstallation.ReadWriteForChat.All` "cannot be used to install apps
 * that require consent to resource-specific permissions", so offering the
 * field would build a request refused by construction. That reading was right
 * about the WEAK role and wrong about the outcome: with the weak role the
 * install is refused whether or not the field is sent, and with the
 * consent-capable `…ReadWriteAndConsentForChat.All` role Graph REQUIRES the
 * field, because an installer has to state which resource-specific
 * permissions it consents to. Omitting it is what produced the field failure.
 *
 * The 400 `ResourceSpecificPermissionsMismatch` is still deliberately NOT
 * folded into the idempotent path — it surfaces as
 * {@link RscPermissionsMismatchError}, whose text differs by whether a set was
 * actually sent.
 */
export class ChatInstallClient {
  private readonly http: Pick<ProvisioningHttp, 'request'>;
  private readonly log: (msg: string) => void;

  constructor(opts: ChatInstallClientOptions) {
    this.http = opts.http;
    this.log =
      opts.log ??
      ((msg: string): void => {
        console.error(msg);
      });
  }

  /**
   * `POST /chats/{chatId}/installedApps` — install the catalog app
   * (`teamsAppId` = `CatalogTeamsApp.teamsAppId`) into the chat.
   *
   * APPLICATION PERMISSIONS WORK HERE. Unlike the catalog upload
   * (`POST /appCatalogs/teamsApps`, delegated-only since byte5ai/omadia#924),
   * this verb supports app-only tokens — {@link CHAT_INSTALL_SCOPE} is an
   * application app role. No device-code sign-in, no protected-API request
   * form to Microsoft.
   *
   * - PRE-FLIGHT → the `chatId` is classified by shape first
   *   ({@link classifyInstallTarget}). A channel id, a team GUID or an
   *   unrecognised string throws `InstallTargetMismatchError` BEFORE any
   *   network call, carrying the remedy — a Graph 404 would name none of it.
   * - 2xx → `'created'`, with the Graph installation id when the response
   *   body carried one.
   * - 409 (and a Bad-Request-shaped duplicate, see
   *   {@link CHAT_ALREADY_INSTALLED_RULES}) → `'already-existed'` — success,
   *   never an exception.
   * - 404 → `ChatNotFoundError`, DISTINCT from the team direction's 404: the
   *   operator's remedy differs (see the error's doc). Graph does not
   *   separate "no such chat" from "no such catalog app" on this verb, but
   *   the app id comes from our own upload step, so the chat is the honest
   *   thing to name.
   * - 400 `ResourceSpecificPermissionsMismatch` → `RscPermissionsMismatchError`,
   *   naming {@link CHAT_INSTALL_CONSENT_SCOPE} and whether a permission set
   *   was actually carried.
   * - 403 → `ConsentMissingError([CHAT_INSTALL_SCOPE], 'graph')` from the
   *   http layer — plus {@link CHAT_INSTALL_CONSENT_SCOPE} when this install
   *   carries a permission set, since that is the role it then needs. Same
   *   typed family as every other step, so callers keep ONE fallback branch.
   * - 429 → retried by the http layer honouring `Retry-After`; exhausted
   *   budget → `ProvisioningThrottledError`.
   */
  async installToChat(
    input: InstallToChatRequest,
  ): Promise<Idempotent<ChatAppInstallation>> {
    const step = 'chats.installedApps.add';
    const chatId = this.requireChatId(input.chatId, step);
    const teamsAppId = requireNonEmpty(input.teamsAppId, 'teamsAppId');

    const permissions = await resolveConsentedPermissionSet(this.http, {
      teamsAppId,
      explicit: input.consentedPermissionSet,
      step,
      log: this.log,
    });

    const response = await installRequest(this.http, {
      url: `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}/installedApps`,
      step,
      teamsAppId,
      permissions,
      baseScope: CHAT_INSTALL_SCOPE,
      consentScope: CHAT_INSTALL_CONSENT_SCOPE,
      conflictOn: CHAT_ALREADY_INSTALLED_RULES,
      // Taken off the generic error path so it can be re-thrown as the typed,
      // chat-specific not-found below.
      extraOkStatuses: [404],
    });

    if (response.kind === 'conflict') {
      this.log(
        `provisioner chats.installedApps.add: app ${teamsAppId} already installed in chat ${chatId} (${String(response.status)} → already-existed)`,
      );
      return {
        outcome: 'already-existed',
        value: chatInstallation(chatId, teamsAppId, response.json),
      };
    }

    if (response.status === 404) {
      throw new ChatNotFoundError(chatId, 'chats.installedApps.add');
    }

    return {
      outcome: 'created',
      value: chatInstallation(chatId, teamsAppId, response.json),
    };
  }

  /**
   * Remove a catalog app from ONE chat — the reverse of
   * {@link installToChat}, and the chat-scope mirror of
   * {@link TeamInstallClient.uninstallFromTeam}.
   *
   * Same two-step for the same reason: Graph deletes an installation by its
   * INSTALLATION id, which the caller generally does not hold.
   *
   *   1. `GET /chats/{chatId}/installedApps?$expand=teamsApp&$filter=teamsApp/id eq '…'`
   *      — no hit → `'already-absent'`, and NO delete is attempted.
   *   2. `DELETE /chats/{chatId}/installedApps/{installationId}` — 2xx →
   *      `'uninstalled'`; 404 → `'already-absent'`, so a removal race never
   *      turns into an exception.
   *
   * A CHAT THAT IS GONE IS `'already-absent'` HERE, not `ChatNotFoundError`.
   * The install direction has to be loud about a bad chat id because it failed
   * to do what was asked; the uninstall direction was asked to make sure an
   * app is not in a chat, and a chat that does not exist satisfies that. Same
   * reasoning the team direction already applies to its 404 lookup.
   */
  async uninstallFromChat(
    input: UninstallFromChatInput,
  ): Promise<UninstallFromChatResult> {
    const chatId = this.requireChatId(
      input.chatId,
      'chats.installedApps.remove',
    );
    const teamsAppId = requireNonEmpty(input.teamsAppId, 'teamsAppId');

    const installationId = await findInstallationId(this.http, {
      collectionUrl: `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}/installedApps`,
      teamsAppId,
      step: 'chats.installedApps.lookup',
      scope: CHAT_INSTALL_SCOPE,
      subject: `chat=${chatId}`,
    });

    if (installationId === undefined) {
      this.log(
        `provisioner chats.installedApps.remove: app ${teamsAppId} not installed in chat ${chatId} (lookup miss → already-absent)`,
      );
      return { outcome: 'already-absent', value: { chatId, teamsAppId } };
    }

    const response = await this.http.request({
      resource: 'graph',
      method: 'DELETE',
      url: `${GRAPH_BASE}/chats/${encodeURIComponent(chatId)}/installedApps/${encodeURIComponent(installationId)}`,
      step: 'chats.installedApps.remove',
      missingScopesOn403: [CHAT_INSTALL_SCOPE],
      // Already gone = already removed.
      extraOkStatuses: [404],
    });

    const gone = response.kind === 'ok' && response.status === 404;
    if (gone) {
      this.log(
        `provisioner chats.installedApps.remove: installation ${installationId} in chat ${chatId} already gone (404 → already-absent)`,
      );
    }

    return {
      outcome: gone ? 'already-absent' : 'uninstalled',
      value: { chatId, teamsAppId, installationId },
    };
  }

  /**
   * Non-empty AND shaped like a chat. Everything the classifier does not call
   * a chat is refused here, with the classification's own remedy text: a
   * channel id points at the team, a bare 32-hex says which full form to
   * re-enter, a team GUID says which method to call instead.
   */
  private requireChatId(value: string, step: string): string {
    const chatId = requireNonEmpty(value, 'chatId');
    const target = classifyInstallTarget(chatId);
    if (isChatTarget(target)) return target.chatId;
    throw new InstallTargetMismatchError(
      step,
      chatId,
      target.kind,
      targetHint(target),
    );
  }
}

/**
 * The ONE install POST, shared by both directions — including the
 * `consentedPermissionSet` handling and the typed 400 both of them need.
 *
 * The body key is present only when a set was resolved: the shape changes with
 * it, so an absent set must not serialise as `consentedPermissionSet:
 * undefined`/`null`.
 *
 * WHY THE 400 IS CAUGHT HERE AND NOT DECLARED AS A CONFLICT RULE. The http
 * layer's `extraOkStatuses` is consulted BEFORE its conflict rules, so listing
 * 400 there would disable the chat direction's Bad-Request-shaped
 * already-installed detection. Catching the thrown
 * {@link ProvisioningRequestError} keeps both behaviours, at the price of
 * matching Graph's error code in the message — the same duck-typing the
 * middleware does one level up, and for the same reason: the code is the only
 * thing carried across that seam.
 */
async function installRequest(
  http: Pick<ProvisioningHttp, 'request'>,
  opts: {
    readonly url: string;
    readonly step: string;
    readonly teamsAppId: string;
    readonly permissions: ConsentedPermissionSet | undefined;
    readonly baseScope: string;
    readonly consentScope: string;
    readonly conflictOn?: readonly ProvisioningConflictRule[];
    readonly extraOkStatuses?: readonly number[];
  },
): Promise<ProvisioningResponse> {
  const sentCount = opts.permissions?.resourceSpecificPermissions.length ?? 0;
  try {
    return await http.request({
      resource: 'graph',
      method: 'POST',
      url: opts.url,
      step: opts.step,
      jsonBody: {
        'teamsApp@odata.bind': `${GRAPH_BASE}/appCatalogs/teamsApps/${encodeURIComponent(opts.teamsAppId)}`,
        ...(opts.permissions !== undefined
          ? { consentedPermissionSet: opts.permissions }
          : {}),
      },
      // The consent-capable role is named only when this install actually
      // consents to something — otherwise a 403 would ask an operator to grant
      // a wider role than the call needs.
      missingScopesOn403:
        sentCount > 0
          ? [opts.baseScope, opts.consentScope]
          : [opts.baseScope],
      ...(opts.conflictOn !== undefined ? { conflictOn: opts.conflictOn } : {}),
      ...(opts.extraOkStatuses !== undefined
        ? { extraOkStatuses: opts.extraOkStatuses }
        : {}),
    });
  } catch (err) {
    if (isRscMismatchResponse(err)) {
      throw new RscPermissionsMismatchError(
        opts.step,
        opts.consentScope,
        sentCount,
        err.message,
        err,
      );
    }
    throw err;
  }
}

/** Graph's `ResourceSpecificPermissionsMismatch` on a 400, as it reaches us. */
function isRscMismatchResponse(err: unknown): err is ProvisioningRequestError {
  return (
    err instanceof ProvisioningRequestError &&
    err.status === 400 &&
    /ResourceSpecificPermissionsMismatch/i.test(err.message)
  );
}

/**
 * The permission set an install should carry: the caller's, when it passed
 * one, else the app's own declared set read back from Graph.
 *
 * An EXPLICIT set is authoritative and skips the lookup entirely — including
 * an explicitly EMPTY one, which is how a caller says "send nothing" without
 * the step second-guessing it.
 */
async function resolveConsentedPermissionSet(
  http: Pick<ProvisioningHttp, 'request'>,
  opts: {
    readonly teamsAppId: string;
    readonly explicit: ConsentedPermissionSet | undefined;
    readonly step: string;
    readonly log: (msg: string) => void;
  },
): Promise<ConsentedPermissionSet | undefined> {
  if (opts.explicit !== undefined) {
    return opts.explicit.resourceSpecificPermissions.length > 0
      ? opts.explicit
      : undefined;
  }
  return resolveDeclaredPermissionSet(http, opts);
}

/**
 * The resource-specific permissions the catalog app itself declares —
 * `GET /appCatalogs/teamsApps?$filter=id eq '…'&$expand=appDefinitions($select=id,authorization)`,
 * reading `appDefinitions[].authorization.requiredPermissionSet`.
 *
 * THIS QUERY IS NOT A GUESS. Graph's own install reference points at it as the
 * way to obtain the set an install must consent to, and requires the consented
 * set to match what the app declares. Reading it back from Graph — rather than
 * from the manifest template that produced the package — also means the set
 * matches the DEFINITION THAT IS ACTUALLY PUBLISHED, which is the thing Graph
 * compares against; a package rebuilt locally after the upload cannot drift
 * this set out from under the install.
 *
 * Entries are passed through VERBATIM, `delegated` ones included: the docs
 * allow omitting the field when the app declares delegated permissions ONLY,
 * they never ask for a filtered set, and filtering would be exactly the
 * mismatch this call exists to avoid.
 *
 * DEGRADES, LOUDLY. A 403/404 on the lookup, an app with no declared
 * permissions, or a definition Graph returns without an `authorization` block
 * all answer `undefined` — the install then proceeds with the pre-0.8.2 body
 * shape and, if the app really did need consent, fails with a
 * {@link RscPermissionsMismatchError} that says the set could not be resolved.
 * Every one of those paths logs why. What it does NOT do is swallow a
 * transport error or an unexpected status: those propagate.
 */
async function resolveDeclaredPermissionSet(
  http: Pick<ProvisioningHttp, 'request'>,
  opts: {
    readonly teamsAppId: string;
    readonly step: string;
    readonly log: (msg: string) => void;
  },
): Promise<ConsentedPermissionSet | undefined> {
  const filter = `id eq '${escapeODataString(opts.teamsAppId)}'`;
  const response = await http.request({
    resource: 'graph',
    method: 'GET',
    url:
      `${GRAPH_BASE}/appCatalogs/teamsApps?$filter=${encodeURIComponent(filter)}` +
      `&$expand=${encodeURIComponent('appDefinitions($select=id,authorization)')}`,
    step: 'appCatalogs.teamsApps.rscLookup',
    missingScopesOn403: [APP_CATALOG_SCOPE],
    // Both mean "nothing declared that we can read": a tenant that withheld
    // AppCatalog.ReadWrite.All, or an app id the catalog does not resolve.
    // Neither should turn an install that used to work into an exception.
    extraOkStatuses: [403, 404],
  });

  if (response.kind !== 'ok' || response.status === 403 || response.status === 404) {
    opts.log(
      `provisioner ${opts.step}: could not read the declared RSC permissions of app ` +
        `${opts.teamsAppId} (appCatalogs.teamsApps.rscLookup answered ${String(response.status)}) — ` +
        `installing without a consentedPermissionSet; grant ${APP_CATALOG_SCOPE} if the install is refused`,
    );
    return undefined;
  }

  const permissions = declaredPermissions(response.json, opts.teamsAppId);
  if (permissions.length === 0) {
    opts.log(
      `provisioner ${opts.step}: app ${opts.teamsAppId} declares no resource-specific permissions — installing without a consentedPermissionSet`,
    );
    return undefined;
  }
  opts.log(
    `provisioner ${opts.step}: consenting to ${String(permissions.length)} resource-specific permission(s) declared by app ${opts.teamsAppId}`,
  );
  return { resourceSpecificPermissions: permissions };
}

/**
 * `authorization.requiredPermissionSet.resourceSpecificPermissions` of the
 * matching catalog entry, flattened across its app definitions and
 * de-duplicated on (value, type).
 *
 * The entry is re-checked against the requested `id` client-side for the same
 * reason the installation lookup re-checks `teamsApp.id`: a tenant that
 * ignores the `$filter` must not be able to make us consent to a DIFFERENT
 * app's permissions. Malformed entries are skipped rather than defaulted —
 * a permission we cannot read is one we must not claim to have consented to.
 */
function declaredPermissions(
  json: unknown,
  teamsAppId: string,
): readonly ResourceSpecificPermission[] {
  const out: ResourceSpecificPermission[] = [];
  const seen = new Set<string>();
  for (const entry of listEntries(json)) {
    if (!entry || typeof entry !== 'object') continue;
    const app = entry as Record<string, unknown>;
    if (app['id'] !== teamsAppId) continue;
    const definitions = app['appDefinitions'];
    if (!Array.isArray(definitions)) continue;
    for (const definition of definitions) {
      for (const permission of permissionsOfDefinition(definition)) {
        const key = `${permission.permissionType}:${permission.permissionValue}`;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(permission);
      }
    }
  }
  return out;
}

/** The RSC permissions of ONE `teamsAppDefinition`, ignoring malformed entries. */
function permissionsOfDefinition(
  definition: unknown,
): readonly ResourceSpecificPermission[] {
  if (!definition || typeof definition !== 'object') return [];
  const authorization = (definition as Record<string, unknown>)['authorization'];
  if (!authorization || typeof authorization !== 'object') return [];
  const set = (authorization as Record<string, unknown>)['requiredPermissionSet'];
  if (!set || typeof set !== 'object') return [];
  const list = (set as Record<string, unknown>)['resourceSpecificPermissions'];
  if (!Array.isArray(list)) return [];

  const out: ResourceSpecificPermission[] = [];
  for (const raw of list) {
    if (!raw || typeof raw !== 'object') continue;
    const record = raw as Record<string, unknown>;
    const value = record['permissionValue'];
    const type = record['permissionType'];
    if (typeof value !== 'string' || value.length === 0) continue;
    if (typeof type !== 'string') continue;
    // Graph's enum members are lowercase; its own examples are inconsistent
    // about casing, so normalise rather than pass an unusable literal on.
    const normalized = type.toLowerCase();
    if (normalized !== 'application' && normalized !== 'delegated') continue;
    out.push({ permissionValue: value, permissionType: normalized });
  }
  return out;
}

/**
 * `GET {collection}?$expand=teamsApp&$filter=teamsApp/id eq '…'` — the
 * installation id for a catalog app in one team or chat, or `undefined` when
 * the app is not installed there.
 *
 * OData string-literal escaping (quote doubling) + `encodeURIComponent` keep
 * the `$filter` injection-safe for arbitrary ids, exactly as the catalog
 * `externalId` lookup does. The returned entries are re-checked against
 * `teamsApp.id` client-side, so a tenant that ignores the `$filter` cannot
 * make us delete the wrong installation.
 *
 * A 404 on the collection (the team/chat is gone) answers `undefined` — the
 * same "nothing to remove" outcome, not an error.
 */
async function findInstallationId(
  http: Pick<ProvisioningHttp, 'request'>,
  opts: {
    /** Absolute `…/installedApps` URL of the team or chat. */
    readonly collectionUrl: string;
    readonly teamsAppId: string;
    readonly step: string;
    /** App role reported on 403. */
    readonly scope: string;
    /** Label for the unexpected-status message, e.g. `chat=19:…`. */
    readonly subject: string;
  },
): Promise<string | undefined> {
  const filter = `teamsApp/id eq '${escapeODataString(opts.teamsAppId)}'`;
  const response = await http.request({
    resource: 'graph',
    method: 'GET',
    url: `${opts.collectionUrl}?$expand=teamsApp&$filter=${encodeURIComponent(filter)}`,
    step: opts.step,
    missingScopesOn403: [opts.scope],
    extraOkStatuses: [404],
  });

  if (response.kind !== 'ok') {
    throw new Error(
      `graph ${opts.step} unexpectedly answered ${String(response.status)} for ${opts.subject}`,
    );
  }
  if (response.status === 404) return undefined;

  for (const entry of listEntries(response.json)) {
    const id = matchingInstallationId(entry, opts.teamsAppId);
    if (id !== undefined) return id;
  }
  return undefined;
}

/**
 * Operator-facing remedy for a target that is not a chat. Most kinds carry
 * their own `hint`; a team GUID is well-formed and correct — just for the
 * other method — so it gets the one sentence the classifier has no business
 * knowing (it does not know which step asked).
 */
function targetHint(
  target: TeamTarget | ChannelTarget | AmbiguousTarget | UnknownTarget,
): string {
  if (target.kind === 'team') {
    return (
      'this is a TEAM id (dashed GUID / group id), not a chat thread id. ' +
      'Install it with installToTeam instead.'
    );
  }
  return target.hint;
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
 * The `teamsAppInstallation.id` of one lookup entry — but only when its
 * expanded `teamsApp.id` really is the requested catalog app.
 */
function matchingInstallationId(
  json: unknown,
  teamsAppId: string,
): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const entry = json as Record<string, unknown>;
  const teamsApp = entry['teamsApp'];
  if (!teamsApp || typeof teamsApp !== 'object') return undefined;
  if ((teamsApp as Record<string, unknown>)['id'] !== teamsAppId) {
    return undefined;
  }
  const id = entry['id'];
  return typeof id === 'string' && id.length > 0 ? id : undefined;
}

/** Build the result value; `installationId` only when Graph returned an id. */
function installation(
  teamId: string,
  teamsAppId: string,
  json: unknown,
): TeamAppInstallation {
  const id = installationId(json);
  return {
    teamId,
    teamsAppId,
    ...(id !== undefined ? { installationId: id } : {}),
  };
}

/** Chat twin of {@link installation}. */
function chatInstallation(
  chatId: string,
  teamsAppId: string,
  json: unknown,
): ChatAppInstallation {
  const id = installationId(json);
  return {
    chatId,
    teamsAppId,
    ...(id !== undefined ? { installationId: id } : {}),
  };
}

/** `teamsAppInstallation.id` from the response body, when present. */
function installationId(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined;
  const id = (json as Record<string, unknown>)['id'];
  return typeof id === 'string' && id.length > 0 ? id : undefined;
}

function requireNonEmpty(value: string, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`invalid_argument: '${field}' must be a non-empty string`);
  }
  return value;
}

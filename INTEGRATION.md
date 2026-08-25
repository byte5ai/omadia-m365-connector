# `@omadia/integration-microsoft365` — Integration Surface

**Source of truth for Builder-Agents that depend on this plugin.** Read this
before writing cross-integration code; do NOT trust your training-memory of
the Microsoft Graph SDK — this plugin exposes a curated 4-surface accessor,
not the raw Graph SDK.

## Service Registry

| Service Name           | TypeScript Type        | Purpose                                              |
|------------------------|------------------------|------------------------------------------------------|
| `microsoft365.graph`   | `Microsoft365Accessor` | Bundled accessor: app-auth + on-behalf-of + calendar + slot-cache |

## Consumption Pattern

Plugin must declare `de.byte5.integration.microsoft365` in `manifest.yaml`'s
`depends_on`. Then in `activate()`:

```typescript
import type { Microsoft365Accessor } from '@omadia/integration-microsoft365';

const ms = ctx.services.get<Microsoft365Accessor>('microsoft365.graph');
if (!ms) {
  throw new Error(
    'microsoft365.graph unavailable — ensure depends_on includes ' +
    '"de.byte5.integration.microsoft365" and the integration is installed/active',
  );
}
```

`peerDependencies` in `package.json` must include
`"@omadia/integration-microsoft365": "*"`.

## `Microsoft365Accessor` API — verbatim

```typescript
interface Microsoft365Accessor {
  /** Microsoft Graph client authenticated via the Bot-Framework App
   *  Registration (client-credentials flow). Use for tenant-level reads
   *  that do NOT belong to a specific user — Teams attachment downloads,
   *  service-principal-scoped operations. */
  readonly app: GraphClient;

  /** Delegated / on-behalf-of access-token acquisition for Graph on
   *  behalf of a single Teams user. Hand the resulting token to
   *  `calendar.*` calls. */
  readonly obo: GraphOboClient;

  /** Stateless Graph Calendar wrapper. Each call takes a per-user
   *  access-token from `obo`. */
  readonly calendar: GraphCalendarClient;

  /** SlotCache surface for round-tripping opaque slot ids across
   *  Adaptive-Card Action.Submit submissions. */
  readonly slots: SlotCacheAccessor;
}
```

### `app: GraphClient`

```typescript
class GraphClient {
  // Fetch a Teams chat message (e.g. for attachment access).
  async fetchChatMessage(chatId: string, messageId: string): Promise<{
    /* upstream Graph message */
  }>;

  // Download a file by Graph sharing-URL.
  async downloadBySharingUrl(url: string): Promise<{
    /* { mimeType, body: ArrayBuffer, …meta } */
  }>;
}
```

### `obo: GraphOboClient`

```typescript
class GraphOboClient {
  readonly scopes: readonly string[];

  // Acquire a per-user delegated access-token via OBO flow.
  // The Teams channel hands the user's bot-framework token in;
  // returns a Graph-scoped access token.
  async acquireTokenForUser(
    userBotToken: string,
    scopes?: readonly string[],
  ): Promise<{ accessToken: string; expiresOn: Date }>;
}
```

OBO failures throw `GraphOboError` with `scopes` set so the caller can
prompt the user to reconsent.

### `calendar: GraphCalendarClient`

```typescript
class GraphCalendarClient {
  // Query free/busy slots. accessToken from obo.acquireTokenForUser.
  async findMeetingTimes(opts: FindSlotsOptions): Promise<MeetingSlotSuggestion[]>;
  async getSchedule(opts: GetScheduleOptions): Promise<ScheduleEntry[]>;
  async createEvent(opts: CreateEventOptions): Promise<CreatedEvent>;

  // Per-user metadata.
  async getSelfAddress(accessToken: string): Promise<string>;
  async getMailboxSettings(accessToken: string): Promise<MailboxSettings>;
}
```

`FindSlotsOptions`, `CreateEventOptions` etc. are exported from the
package; check `./src/graphCalendarClient.ts` for the exact shapes
when needed.

### `slots: SlotCacheAccessor`

```typescript
interface SlotCacheAccessor {
  put(entry: Omit<CachedSlot, 'slotId' | 'expiresAt'>): CachedSlot;
  get(slotId: string): CachedSlot | undefined;
  consume(slotId: string): CachedSlot | undefined;  // get + delete
}
```

Use only inside Adaptive-Card response cycles where you need an opaque
slot id round-tripped without exposing schedule details to the user.

## Concrete Snippets

### Find free slots for the current user

```typescript
const userToken = await ms.obo.acquireTokenForUser(userBotToken);
const slots = await ms.calendar.findMeetingTimes({
  accessToken: userToken.accessToken,
  attendees: ['user@contoso.com', 'colleague@contoso.com'],
  durationMinutes: 30,
  windowStart: new Date(),
  windowEnd: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
});
```

### Create a calendar event

```typescript
const userToken = await ms.obo.acquireTokenForUser(userBotToken);
const event = await ms.calendar.createEvent({
  accessToken: userToken.accessToken,
  subject: '1:1 with team lead',
  start: new Date('2026-05-10T10:00:00Z'),
  end: new Date('2026-05-10T10:30:00Z'),
  attendees: ['lead@contoso.com'],
  body: { contentType: 'text', content: 'Weekly sync.' },
});
```

### Tenant-level: download a Teams attachment

```typescript
const file = await ms.app.downloadBySharingUrl(sharingUrl);
// file.body is ArrayBuffer; file.mimeType is the upstream Content-Type.
```

## Graph application scopes

Machine-readable list of the Graph **application** permissions this connector
requests. The baseline set is use-case-driven (see `manifest.yaml`
`setup.guide`, e.g. `Calendars.Read`, `Mail.Read`, `User.Read.All`,
`Files.Read.All`); the Teams provisioning capability (`teamsProvisioner@1`,
spec in [`docs/teams-provisioner.md`](docs/teams-provisioner.md), credentials
issue [#2](https://github.com/byte5ai/omadia-m365-connector/issues/2))
extends it by exactly three scopes:

| Scope | Chain step it unlocks |
|---|---|
| `Application.ReadWrite.OwnedBy` | `registerApplication` / `addClientSecret` — `POST /applications`, `POST /applications/{id}/addPassword`; per-agent apps stay owned by the connector app |
| `AppCatalog.ReadWrite.All` | `uploadToCatalog` — `POST /appCatalogs/teamsApps` |
| `TeamsAppInstallation.ReadWriteForTeam.All` | `installToTeam` — `POST /teams/{id}/installedApps` |

### Consent semantics (renewed admin consent)

- Adding any of these scopes to an existing app registration **requires
  renewed admin consent**: an admin must click **Grant admin consent for
  &lt;Tenant&gt;** again after the permission change. Consent granted before
  the change does not cover the new scopes.
- **Consent sometimes only lands via REST `appRoleAssignments` + restart.**
  Field-tested failure mode: portal/CLI consent (`az ad app permission
  admin-consent`) reports success but Graph keeps returning `403`. Fix: grant
  the app roles directly — `POST
  /servicePrincipals/{graph-sp-object-id}/appRoleAssignments`
  (`principalId` = connector SP object id, `appRoleId` = the role id of the
  missing permission) — then **restart the middleware** so a fresh token is
  acquired; cached tokens never gain roles retroactively.
- Missing or unconsented provisioning scopes surface as the typed
  `ConsentMissingError` (`403`, carries `missingScopes` + `resource`) from
  `src/teamsProvisioner/errors.ts` — callers get the scope list to render a
  consent prompt, nobody string-matches Graph error bodies.

### Manifest `setup.guide` wording (handoff to the wiring unit)

`manifest.yaml`'s bilingual `setup.guide` (en **and** de) section
`### 3. Graph permissions` must gain the following text verbatim — the edit
itself is executed by the wiring unit so parallel W0b units never collide in
`manifest.yaml`:

**en** — append after the existing example-permissions step:

> For the Teams provisioning capability (`teamsProvisioner@1`) additionally
> add: `Application.ReadWrite.OwnedBy`, `AppCatalog.ReadWrite.All`,
> `TeamsAppInstallation.ReadWriteForTeam.All`.
>
> Extending an existing app registration requires **renewed admin consent**:
> click **Grant admin consent for &lt;Tenant&gt;** again. If Graph still
> answers 403 afterwards, grant the app roles via REST `appRoleAssignments`
> and restart the middleware — portal/CLI consent sometimes silently fails
> to apply.

**de** — nach dem bestehenden Beispiel-Berechtigungen-Schritt anfügen:

> Für die Teams-Provisionierung (`teamsProvisioner@1`) zusätzlich hinzufügen:
> `Application.ReadWrite.OwnedBy`, `AppCatalog.ReadWrite.All`,
> `TeamsAppInstallation.ReadWriteForTeam.All`.
>
> Erweiterte Berechtigungen einer bestehenden App-Registration erfordern
> **erneuten Admin-Consent**: **Grant admin consent for &lt;Tenant&gt;**
> erneut klicken. Antwortet Graph danach weiterhin mit 403, die App-Rollen
> per REST `appRoleAssignments` zuweisen und die Middleware neu starten —
> Portal-/CLI-Consent greift manchmal stillschweigend nicht.

## Was NICHT geht

- ❌ **Importing the Microsoft Graph SDK directly** (`@microsoft/microsoft-graph-client`)
  — bypasses our auth/credential rotation. Always go through `Microsoft365Accessor`.
- ❌ **`ms.graph.api(...)` style chained calls** — we don't expose the raw SDK.
  The 4 sub-clients (`app` / `obo` / `calendar` / `slots`) are the surface.
- ❌ **OBO without a user-bot-token** — `obo.acquireTokenForUser` requires the
  Teams user's bot-framework token. Cron / background jobs without a user
  context use `app: GraphClient` instead (tenant-level only).
- ❌ **Mixing `app` and `obo` tokens** — the `app` token is service-principal-
  scoped, the `obo` token is user-scoped. Calendar reads/writes need OBO.
  Tenant-level reads (e.g. file metadata) use `app`.
- ❌ **Don't construct `Microsoft365Accessor` yourself** — credentials live
  in the vault under the integration's scope. Always
  `ctx.services.get<Microsoft365Accessor>('microsoft365.graph')`.

## Reference implementations

- `harness-channel-teams` — uses `app.fetchChatMessage` + `app.downloadBySharingUrl`
  for attachment handling.
- Calendar tools in the orchestrator — use `obo.acquireTokenForUser` +
  `calendar.findMeetingTimes` / `calendar.createEvent`.

## Versioning

Adding new methods to any sub-client is non-breaking. Renaming or
removing a sub-client (e.g. dropping `obo` if a future SSO flow replaces
OBO) is a major-version event — check git-blame on this file.

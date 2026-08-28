# teamsProvisioner@1 — capability spec

Status: **spec-first** (this document lands before any capability code — it is
the contract the W0b/W1b implementation units build against).
Epic: [byte5ai/omadia#860](https://github.com/byte5ai/omadia/issues/860) ·
capability issue [byte5ai/omadia-m365-connector#3](https://github.com/byte5ai/omadia-m365-connector/issues/3) ·
credentials issue [byte5ai/omadia-m365-connector#2](https://github.com/byte5ai/omadia-m365-connector/issues/2).

## Goal

Run multiple omadia agents as separate named bot identities in the same
Microsoft Teams channel: **1 agent = 1 Entra app + 1 Azure bot + 1 generated
Teams app package**. The M365 connector gains a `teamsProvisioner@1`
capability covering the full chain:

1. **Entra app registration** — Graph `POST /applications`
2. **Client secret** — Graph `POST /applications/{id}/addPassword`
3. **Azure bot** — ARM REST `PUT .../Microsoft.BotService/botServices/{name}`
4. **Catalog upload** — Graph `POST /appCatalogs/teamsApps`
5. **Team install** — Graph `POST /teams/{id}/installedApps`

The provisioning **state machine** (ordering, persistence, retries across
steps) lives middleware-side in the agent factory
(byte5ai/omadia#863–#865); this capability exposes exactly the individual
steps, keeping the public surface minimal and typed.

## Invariants

- **Token-based REST only** (Graph + ARM). No `az` CLI, no SDK lock-in.
- **SingleTenant, customer tenant.** New MultiTenant registrations are
  deprecated since 07/2025. The types model `signInAudience` as the single
  value `'AzureADMyOrg'` and `tenantMode: 'customer' | 'home'` — MultiTenant
  is not expressible. SingleTenant apps are messaging-only outside their home
  tenant (epic deployment breakpoint).
- **Spec before code.** This file is the root of the wave; capability code
  must conform to `src/teamsProvisioner/types.ts` + `errors.ts`.

## Service naming

Two-constant split, mirroring `transcription.ts` in `@omadia/plugin-api`:

| Constant | Value | Used as |
|---|---|---|
| `TEAMS_PROVISIONER_SERVICE_NAME` | `teamsProvisioner` | ServiceRegistry key |
| `TEAMS_PROVISIONER_CAPABILITY` | `teamsProvisioner@1` | manifest `provides:` ref |

Registration itself (plugin wiring, `src/index.ts` barrel export, manifest
`provides`) belongs to the **wiring unit**, not this spec's unit.

## Public surface (`src/teamsProvisioner/types.ts`)

`TeamsProvisioner` — one method per chain step:

- `registerApplication(input) → Idempotent<AppRegistration>`
- `addClientSecret(input) → AppClientSecret` (secret value returned exactly once)
- `createBot(input) → BotProvisioningOutcome`
- `uploadToCatalog(input) → Idempotent<CatalogTeamsApp>`
- `installToTeam(input) → Idempotent<TeamAppInstallation>`

plus two readonly probes: `tenantMode` and `canCreateBots`.

> **Superseded at implementation time.** The SHIPPED service surface is
> `TeamsProvisionerAccessor` (`src/teamsProvisioner/index.ts`): register-app +
> add-secret became ONE rolled-back step (`createAppRegistration`) and only
> the opaque vault `secretRef` (`teams_bot_password:<appId>`) — never
> `AppClientSecret.secretText` — crosses the service boundary. The sketch
> above is kept (and exported `@deprecated`) as the reviewed historical
> contract; resolve the service as
> `ctx.services.get<TeamsProvisionerAccessor>('teamsProvisioner')`.

### Catalog lookup — `getCatalogApp` (since 0.3.1)

`getCatalogApp({ teamsAppExternalId })` on the shipped
`TeamsProvisionerAccessor` resolves an EXISTING catalog app by its manifest id
(`externalId`) **without uploading a package** — for consumers that only need
the `teamsAppId` of an already-published app (e.g. to drive `installToTeam`).
It reuses the exact query of the 409 idempotent upload path
(`GET /appCatalogs/teamsApps?$filter=externalId eq '…'` with
`$expand=appDefinitions($select=version,publishingState)`; quote-doubling +
`encodeURIComponent` keep the filter injection-safe) and the same version
selection: the `published` appDefinition wins, else the highest version.

Result: `{ found: false }` (a plain outcome, never an exception) or
`{ found: true, teamsAppId, displayName?, publishedVersion? }`. Errors map
like every catalog call: 403 → `ConsentMissingError(['AppCatalog.ReadWrite.All'], 'graph')`,
exhausted 429 backoff → `ProvisioningThrottledError`.

### Team uninstall — `uninstallFromTeam` (since 0.4.0)

`uninstallFromTeam({ teamId, teamsAppId })` on the shipped
`TeamsProvisionerAccessor` is the reverse of `installToTeam`
(byte5ai/omadia#900): it removes an agent's Teams app from ONE team, keyed by
the same `(teamId, teamsAppId)` pair the install is idempotent on.

Graph deletes an installation by its **installation id**, not by the catalog
app id, so the step is a lookup-then-DELETE pair:

1. `GET /teams/{teamId}/installedApps?$expand=teamsApp&$filter=teamsApp/id eq '…'`
   — quote-doubling + `encodeURIComponent` keep the filter injection-safe, and
   the returned entries are re-checked against `teamsApp.id` client-side so a
   tenant that ignores `$filter` cannot make us delete the wrong installation.
2. `DELETE /teams/{teamId}/installedApps/{installationId}`.

Result: `{ outcome: 'uninstalled' | 'already-absent', value: TeamAppInstallation }`.
**"Not installed" is success, never an exception** — the lookup missing, the
lookup answering 404 (team gone) and the DELETE answering 404 (another
remover won the race) all collapse into `'already-absent'`; only the first of
those has no `installationId` to report. The literal deviates from the
`'already-deleted'` of the app-registration / bot rollbacks on purpose:
nothing is destroyed here, an app that was never in the team is simply absent.

Errors map like the install direction: 403 →
`ConsentMissingError(['TeamsAppInstallation.ReadWriteForTeam.All'], 'graph')`
(the same scope covers reading and removing an installation), exhausted 429
backoff → `ProvisioningThrottledError`.

> **Consumers must feature-detect.** The middleware mirrors this contract
> structurally rather than importing it, so a middleware talking to a
> connector `< 0.4.0` must keep its not-supported branch
> (`typeof provisioner.uninstallFromTeam === 'function'`) instead of crashing.

### Idempotency — "already exists" is not an error

Steps that can hit "already exists" on re-runs (catalog upload via
`externalId` lookup, team install, app registration via `uniqueName`) return
`Idempotent<T> = { outcome: 'created' | 'already-existed', value }`. Callers
branch on `outcome`; nobody string-matches Graph error bodies.

The signal is usually a 409 — but not always. Entra reports a taken
`uniqueName` on `POST /applications` as **400 `Request_BadRequest`**
("Another object with the same value for property uniqueName already
exists."). A request may therefore declare `conflictOn` rules, and the choke
point maps a matching status/code/message onto the same
`{ kind: 'conflict' }` result. See `UNIQUE_NAME_CONFLICT_RULES` in
`src/teamsProvisioner/directory.ts`.

### Entra replication and the recycle bin (byte5ai/omadia#916)

Two Entra behaviours the app-registration step has to survive, both found on
the first real run against a customer tenant:

**Eventual consistency.** `POST /applications` answers 201 and the immediate
`POST /applications/{id}/addPassword` can answer 404
`Request_ResourceNotFound` for a few seconds — the object exists, it is not
replicated to the node serving the follow-up write. The step polls the new
object until it is readable and retries `addPassword` through the same window
(8 probes, ~40 s total, all seams injectable). An exhausted budget raises
`DirectoryReplicationError`, which is **transient**: the app exists and is
adoptable.

**The recycle bin.** A deleted application is only SOFT-deleted, and its
`uniqueName` stays reserved for 30 days while the object is invisible in
`GET /applications`. On a conflict the step therefore looks in
`directory/deletedItems/microsoft.graph.application` and **restores** the
holder when it is one of ours (exact `uniqueName` match) — a restore returns
the original object, which is exactly what re-provisioning the same agent slug
should yield. If the recycle bin cannot be read (it needs
`Application.ReadWrite.All`, which this connector deliberately does not ask
for) or the restore fails, the step raises `UniqueNameReservedError`, whose
message says *why* the name is unavailable and for how long instead of a bare
"already exists" about an object the operator cannot find.

### Rollback is narrow, on purpose

The failure that motivated all of the above: a transient 404 was treated as a
step failure, the partially-created app was rolled back, the delete
soft-deleted it, its `uniqueName` became unavailable for 30 days, and all
four remaining retries collided with an object nobody could see. One
transient error burned a provisioning slug for a month.

So `createAppRegistration` now:

1. hands the caller the `appId` via `onRegistrationCreated` the moment the
   registration exists — **before** the secret and the service principal — so
   an interruption leaves a resumable row instead of an orphan;
2. rolls back **nothing** when the failure is transient
   (`isTransientProvisioningFailure`: throttling, pending replication, 5xx,
   408/425/429, transport errors);
3. never deletes a registration that carries a `uniqueName`, even on a
   non-transient failure — it is addressable by its natural key, so the next
   run adopts it. Deleting costs the name; keeping it costs one adoptable app.
   A registration with no `uniqueName` is an orphan nothing could find again
   and is still deleted.

`deleteAppRegistration` stays the explicit, idempotent delete for
deprovisioning.

Adoption always mints a **fresh** client secret (the original was returned
once and never persisted). Because that would otherwise accumulate a
credential per re-run against Entra's cap, the step removes its own superseded
credentials — matched by the deterministic secret label — and leaves every
other credential on the app untouched.

### Graceful degradation — registration-only mode

When the ARM setup fields are absent, `createBot` returns the typed
`RegistrationOnlyOutcome` (`kind: 'registration-only'`,
`reason: 'arm-not-configured'`, `missingSetupFields`). The chain can still
register the Entra app + upload the package; the operator creates the bot
manually. `canCreateBots` lets callers pre-flight this.

## Error taxonomy (`src/teamsProvisioner/errors.ts`)

Precedent: `ConsentRequiredError` / `SsoUnavailableError` in `src/graphObo.ts`
(named `Error` subclasses, explicit `this.name`, structured readonly fields,
snake_case messages). Base class `TeamsProvisionerError` for catch-all.

| Error | When | Carries |
|---|---|---|
| `ConsentMissingError` | Graph/ARM 403, permission or admin consent missing | `missingScopes`, `resource: 'graph' \| 'arm'` — so the middleware factory can fall back (deep-link consent card) |
| `ProvisioningThrottledError` | 429 backoff budget exhausted, or a Retry-After hint beyond the 60 s backoff cap | `retryAfterSeconds?`, `resource` |
| `ArmNotConfiguredError` | ARM step demanded while setup fields missing | `missingSetupFields` |
| `CapabilityUnavailableError` | Vault write/delete attempted but the kernel did not hand out `secrets.set`/`secrets.delete` (manifest lacks `permissions.secrets.runtime_write`) | `missingPermission`, `operation: 'set' \| 'delete'` |

Anything else propagates verbatim (matching `GraphOboClient`). ⚠️ Name
collision by design: `graphObo.ts` already exports `ConsentRequiredError`
(delegated calendar flow); `ConsentMissingError` is the provisioning-flow
(application-permission) counterpart. Both live on the public surface.

## Credentials & scopes (issue #2 — config unit)

Additional Graph **application** permissions: `Application.ReadWrite.OwnedBy`,
`AppCatalog.ReadWrite.All`, `TeamsAppInstallation.ReadWriteForTeam.All`.
New setup fields: Azure subscription id, resource group, region, and a
service-principal credential for `management.azure.com` (ARM). All ARM fields
are optional — absence triggers registration-only mode, never a crash.

## Consumers

The middleware agent factory (byte5ai/omadia#863–#865) resolves the service
via the registry and drives the chain; the channel-teams installer units
(W2) consume the install step's results.

## Testing

`npm test` → `scripts/test.mjs` (ported from `omadia-channel-teams`): esbuild
transpiles `tests/*.test.ts` into `.test-build/`, then `node --test` runs
them. `tests/` sits outside `tsconfig`'s `include` (`src/**/*.ts`), so tests
never leak into `dist/`.

> Note for reviewers: `docs/` is a new directory in this repo (previously only
> `README.md` / `INTEGRATION.md`) — introduced by wave W0b as the home for
> capability specs.

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

### Idempotency — 409 is not an error

Steps that can hit "already exists" on re-runs (catalog upload via
`externalId` lookup, team install, app registration via `uniqueName`) return
`Idempotent<T> = { outcome: 'created' | 'already-existed', value }`. Callers
branch on `outcome`; nobody string-matches Graph error bodies.

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

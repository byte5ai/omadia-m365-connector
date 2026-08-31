<div align="center">

# @omadia/integration-microsoft365

### Shared Microsoft Graph layer for omadia's Microsoft 365 consumers: app-auth client, on-behalf-of user-token client, Calendar client, and a slot cache.

A **Microsoft 365** integration for [omadia](https://github.com/byte5ai/omadia).
It publishes the `microsoft365.graph` service (`Microsoft365Accessor`) to the
service registry, consumed by the Teams channel and calendar-related tools.

[![License: MIT](https://img.shields.io/badge/License-MIT-black.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/built%20with-TypeScript-3178C6.svg?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)

</div>

---

## How it works

| Concern | Implementation |
|---|---|
| Auth | Microsoft Entra ID (Azure AD) app registration, client-credentials (app-only) flow via `@azure/msal-node` |
| Delegated access | On-behalf-of (OBO) user-token flow for calls that need a signed-in user's context |
| Graph calls | `@microsoft/microsoft-graph-client`, wrapped by `graphClient.ts` / `graphCalendarClient.ts` |
| Caching | `slotCache.ts` — short-lived cache for calendar free/busy slot lookups |
| Service surface | `accessor.ts` exposes `Microsoft365Accessor`, registered as the `microsoft365.graph` service |
| Entry point | `plugin.ts` → bundled to `dist/plugin.js` on `npm run build` |

This package has no user-facing tools of its own — it is implicitly installed
whenever a Microsoft 365 consumer (the Teams channel, calendar agents, and
eventually Mail/OneDrive integrations) is installed, and provides the shared
Graph client via the service registry.

## Setup (one time)

1. [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App
   registrations** → **New registration**.
2. Copy the **Directory (tenant) ID** and **Application (client) ID** from
   **Overview**.
3. **Certificates & secrets** → **New client secret** → copy the value
   immediately (shown once).
4. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Application permissions** → add what your use case needs (e.g.
   `Calendars.Read`, `Mail.Read`, `User.Read.All`, `Files.Read.All`) → **Grant
   admin consent**.
5. For the Teams provisioning capability (`teamsProvisioner@1` — one bot
   identity per omadia agent, see
   [`docs/teams-provisioner.md`](docs/teams-provisioner.md)) additionally add
   these **application** permissions:
   - `Application.ReadWrite.OwnedBy` — create the per-agent Entra app
     registrations and their client secrets (`POST /applications`,
     `POST /applications/{id}/addPassword`)
   - `AppCatalog.ReadWrite.All` — **resolve** apps in the tenant app catalog
     (`GET /appCatalogs/teamsApps`). Since 0.6.0 this app permission no longer
     covers the *upload*: Graph documents application permissions for
     `POST /appCatalogs/teamsApps` as **"Not supported."**, and the field test
     confirms it — the same token that resolves a catalog app is refused for
     the upload no matter what is consented. The upload runs on a **delegated**
     token instead, acquired once per tenant through a device-code sign-in. See
     [`docs/teams-provisioner-delegated-publish.md`](docs/teams-provisioner-delegated-publish.md).
   - `TeamsAppInstallation.ReadWriteForTeam.All` — install the catalog app
     into the target team (`POST /teams/{id}/installedApps`) and, since
     0.4.0, remove it again (`GET /teams/{id}/installedApps` +
     `DELETE /teams/{id}/installedApps/{installationId}`)
   - `TeamsAppInstallation.ReadWriteForChat.All` — since 0.7.0, install the
     catalog app into a group or 1:1 **chat**
     (`POST /chats/{id}/installedApps`) and remove it again. Unlike the
     catalog upload this verb DOES support application permissions — no
     device-code sign-in, and it is not a Teams protected API. Not the
     `…SelfForChat.All` variant, which only lets an app install *itself*.
     Needed only for chat installs; team-only deployments can skip it.
   - `TeamsAppInstallation.ReadWriteAndConsentForTeam.All` and
     `TeamsAppInstallation.ReadWriteAndConsentForChat.All` — required in
     practice since 0.8.2. The generated app packages declare seven
     resource-specific permissions, and the two plain roles above may not
     consent to them: without these, an install answers **400
     `ResourceSpecificPermissionsMismatch`**. The connector sends the matching
     `consentedPermissionSet` itself, read back from the published app
     definition; the role is the half a tenant admin grants.
   - `Team.ReadBasic.All` — since 0.5.0, resolve a team id to its display name
     (`GET /teams/{id}?$select=id,displayName`) so consumers can show operators
     a team NAME instead of a GUID. Read-only, and the narrowest scope that
     answers the question. Without it the lookup 403s and consumers fall back
     to the id — nothing else in the chain depends on it.

This is the app-only (client-credentials) flow: no interactive user sign-in;
access is tenant-wide within the granted permissions. The same app
registration is reused by the Teams channel plugin.

### Renewed admin consent

Extending the permissions of an **existing** app registration requires
**renewed admin consent** — previously granted consent does not stretch to
cover newly added scopes. After every permission change, click **Grant admin
consent for &lt;Tenant&gt;** again; app-only tokens carry only the app roles
consented at token-issue time.

Two field-tested gotchas:

- **Portal/CLI consent sometimes silently fails to apply** (observed with
  `az ad app permission admin-consent`: the command succeeds, Graph keeps
  answering `403`). In that case grant the app roles directly via REST
  `appRoleAssignments`, one call per missing permission — on the
  **connector's own** service principal:

  ```
  POST /servicePrincipals/{connector-sp-object-id}/appRoleAssignments
  {
    "principalId": "{connector-sp-object-id}",
    "resourceId":  "{graph-sp-object-id}",
    "appRoleId":   "{app-role-id-of-the-missing-permission}"
  }
  ```

  The Microsoft Graph service principal's object id (`resourceId`) is
  resolved via
  `GET /servicePrincipals(appId='00000003-0000-0000-c000-000000000000')`.
  Verify with `GET
  /servicePrincipals/{connector-sp-object-id}/appRoleAssignments`.
- **Restart after consent.** Acquired tokens are cached; newly consented
  roles only appear in a *fresh* token. Restart the middleware (or wait for
  token expiry) after granting consent, otherwise the `403`s persist even
  though consent is in place.

## Build & install

```bash
npm install
npm run typecheck   # tsc --noEmit (src) + tsc -p tsconfig.tests.json (tests incl. type-level assertions)
npm run build        # tsc
npm test             # esbuild-transpiled node:test suite (scripts/test.mjs)
```

`@omadia/plugin-api` is a **peer dependency**, provided by the omadia host at
runtime. For local typechecking, `tsconfig.json` maps it to
`../odoo-bot/middleware/packages/plugin-api/dist` — check out `odoo-bot`
alongside this repo, or adjust the `paths` entry to your local plugin-api
type sources.

## Manifest

See [`manifest.yaml`](manifest.yaml) for the full plugin manifest (service
types, setup fields, network/permission declarations) and
[`INTEGRATION.md`](INTEGRATION.md) for integration-specific implementation
notes.

## License

MIT © byte5 GmbH — see [LICENSE](LICENSE).

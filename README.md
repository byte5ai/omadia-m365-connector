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

This is the app-only (client-credentials) flow: no interactive user sign-in;
access is tenant-wide within the granted permissions. The same app
registration is reused by the Teams channel plugin.

## Build & install

```bash
npm install
npm run typecheck   # tsc --noEmit
npm run build        # tsc
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

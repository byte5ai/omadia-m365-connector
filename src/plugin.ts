import type { PluginContext } from '@omadia/plugin-api';

import {
  MICROSOFT365_SERVICE_NAME,
  type Microsoft365Accessor,
} from './accessor.js';
import { GraphClient } from './graphClient.js';
import { GraphOboClient } from './graphObo.js';
import { GraphCalendarClient } from './graphCalendarClient.js';
import { SlotCache } from './slotCache.js';

/**
 * @omadia/integration-microsoft365 — plugin entry point.
 *
 * `kind: integration`. Exposes one thing to the kernel on activate():
 * the `microsoft365.graph` service (`Microsoft365Accessor`) bundling the
 * four underlying clients:
 *
 *   - `GraphClient`           app-credentials client (attachment downloads)
 *   - `GraphOboClient`        OBO / delegated token client (calendar users)
 *   - `GraphCalendarClient`   stateless Calendar API wrapper
 *   - `SlotCache`             in-memory TTL cache for calendar slot ids
 *
 * Required config (via `ctx.config`):
 *   - `microsoft_tenant_id`   Azure AD tenant GUID
 *   - `microsoft_app_id`      Bot-Framework App (Client) ID
 *
 * Required secret (via `ctx.secrets`):
 *   - `microsoft_app_password`  Bot-Framework App (Client) Secret
 *
 * Consumers (Teams channel + calendar tools) reach the accessor via
 * `ctx.services.get<Microsoft365Accessor>('microsoft365.graph')`. They
 * should declare `de.byte5.integration.microsoft365` in their manifest's
 * `depends_on` so the kernel activates this plugin before them.
 */

export interface Microsoft365PluginHandle {
  close(): Promise<void>;
}

export async function activate(
  ctx: PluginContext,
): Promise<Microsoft365PluginHandle> {
  ctx.log('activating microsoft365 integration');

  const tenantId = ctx.config.require<string>('microsoft_tenant_id');
  const clientId = ctx.config.require<string>('microsoft_app_id');
  const clientSecret = await ctx.secrets.require('microsoft_app_password');

  const app = new GraphClient({ tenantId, clientId, clientSecret });
  const obo = new GraphOboClient({ tenantId, clientId, clientSecret });
  const calendar = new GraphCalendarClient();
  const slotCache = new SlotCache();

  const accessor: Microsoft365Accessor = {
    app,
    obo,
    calendar,
    slots: {
      put: (entry) => slotCache.put(entry),
      get: (id) => slotCache.get(id),
      consume: (id) => slotCache.consume(id),
    },
  };

  const disposeService = ctx.services.provide<Microsoft365Accessor>(
    MICROSOFT365_SERVICE_NAME,
    accessor,
  );

  ctx.log(
    `[microsoft365] ready (tenant=${tenantId}, app=${clientId}) — service '${MICROSOFT365_SERVICE_NAME}' published`,
  );

  return {
    async close(): Promise<void> {
      ctx.log('deactivating microsoft365 integration');
      disposeService();
    },
  };
}

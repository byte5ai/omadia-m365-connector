import type { PluginContext } from '@omadia/plugin-api';

import {
  MICROSOFT365_SERVICE_NAME,
  type Microsoft365Accessor,
} from './accessor.js';
import { GraphClient } from './graphClient.js';
import { GraphOboClient } from './graphObo.js';
import { GraphCalendarClient } from './graphCalendarClient.js';
import { SlotCache } from './slotCache.js';
import {
  TEAMS_PROVISIONER_CAPABILITY,
  TEAMS_PROVISIONER_SERVICE_NAME,
  createTeamsProvisioner,
  readArmConfig,
  type TeamsProvisionerAccessor,
} from './teamsProvisioner/index.js';

/**
 * @omadia/integration-microsoft365 — plugin entry point.
 *
 * `kind: integration`. Exposes two things to the kernel on activate():
 *
 * 1. The `microsoft365.graph` service (`Microsoft365Accessor`) bundling the
 *    four underlying clients:
 *
 *    - `GraphClient`           app-credentials client (attachment downloads)
 *    - `GraphOboClient`        OBO / delegated token client (calendar users)
 *    - `GraphCalendarClient`   stateless Calendar API wrapper
 *    - `SlotCache`             in-memory TTL cache for calendar slot ids
 *
 * 2. The `teamsProvisioner` service (`TeamsProvisionerAccessor`, capability
 *    `teamsProvisioner@1`) — per-agent Teams bot provisioning (Entra app +
 *    Azure bot + app package + catalog + team install) for the agent factory
 *    (byte5ai/omadia#863-865).
 *
 * Required config (via `ctx.config`):
 *   - `microsoft_tenant_id`   Azure AD tenant GUID
 *   - `microsoft_app_id`      Bot-Framework App (Client) ID
 *
 * Required secret (via `ctx.secrets`):
 *   - `microsoft_app_password`  Bot-Framework App (Client) Secret
 *
 * OPTIONAL ARM setup fields (`azure_subscription_id`, `azure_resource_group`,
 * `azure_region`, `azure_sp_client_id` + `azure_sp_client_secret`) are read
 * with the non-throwing accessors (`ctx.config.get` / `ctx.secrets.get`):
 * when absent, activation still succeeds and `teamsProvisioner` is published
 * in registration-only mode (`canCreateBots === false`, `createBot` answers
 * the typed `RegistrationOnlyOutcome`). Activation must never throw over
 * missing ARM configuration.
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

  // teamsProvisioner@1 — optional ARM fields via the NON-throwing accessors:
  // a missing/half configuration degrades to registration-only mode, it never
  // fails activation (readArmConfig uses config.get / secrets.get only).
  const armConfig = await readArmConfig(ctx, { clientId, clientSecret });
  const provisioner = createTeamsProvisioner({
    graphCredential: { tenantId, clientId, clientSecret },
    armConfig,
    secrets: ctx.secrets,
    tenantMode: 'customer',
    log: (msg) => ctx.log(msg),
  });
  const disposeProvisioner = ctx.services.provide<TeamsProvisionerAccessor>(
    TEAMS_PROVISIONER_SERVICE_NAME,
    provisioner,
  );

  ctx.log(
    provisioner.canCreateBots
      ? `[microsoft365] ${TEAMS_PROVISIONER_CAPABILITY} ready — service '${TEAMS_PROVISIONER_SERVICE_NAME}' published (ARM configured, bot creation enabled)`
      : `[microsoft365] ${TEAMS_PROVISIONER_CAPABILITY} in registration-only mode — service '${TEAMS_PROVISIONER_SERVICE_NAME}' published (missing setup fields: ${
          armConfig.kind === 'registration-only'
            ? armConfig.missingSetupFields.join(', ')
            : ''
        })`,
  );

  return {
    async close(): Promise<void> {
      ctx.log('deactivating microsoft365 integration');
      disposeProvisioner();
      disposeService();
    },
  };
}

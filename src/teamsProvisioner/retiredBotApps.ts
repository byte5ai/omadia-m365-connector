/**
 * Which app ids already had an Azure bot resource that was deleted
 * (since 0.9.2, byte5ai/omadia#2043 rollout defect D10).
 *
 * A bot resource re-created under such an app id is rejected by the Bot
 * Connector with 401 for good, while Azure Resource Manager reports it
 * healthy. `deleteBot` therefore records the app id here and `createBot`
 * refuses to create a NEW resource under a recorded one.
 *
 * WHERE IT IS KEPT. In the plugin vault, one entry per app id under
 * `teams_retired_bot_app:<appId>`, next to `teams_bot_password:<appId>`. The
 * value is not a secret (a timestamp and the handle); the vault is the only
 * store this connector has that survives a restart, and the record has to:
 * the delete and the next create are separate provisioning runs. Without
 * `permissions.secrets.runtime_write` the record lives in memory only, which
 * still covers a delete and a create within one process.
 */

import type { SecretsAccessor } from '@omadia/plugin-api';

/** Vault-key prefix of the retired-app-id records. */
export const TEAMS_RETIRED_BOT_APP_PREFIX = 'teams_retired_bot_app';

/** What is known about a retired app id. */
export interface RetiredBotApp {
  /** ISO-8601 time of the bot delete, when the record carries one. */
  readonly retiredAt?: string;
  /** Handle of the deleted bot resource, when the record carries one. */
  readonly botName?: string;
}

export interface RetiredBotAppsOptions {
  /** Plugin vault accessor. Absent → memory only. */
  readonly secrets?: SecretsAccessor;
  readonly log: (msg: string) => void;
  readonly now?: () => number;
}

/** Record and look up retired bot app ids. One instance per provisioner. */
export class RetiredBotApps {
  private readonly secrets: SecretsAccessor | undefined;
  private readonly log: (msg: string) => void;
  private readonly now: () => number;
  private readonly memory = new Map<string, RetiredBotApp>();

  constructor(opts: RetiredBotAppsOptions) {
    this.secrets = opts.secrets;
    this.log = opts.log;
    this.now = opts.now ?? Date.now;
  }

  /** Record that the bot resource of `msaAppId` was deleted. Never throws. */
  async retire(msaAppId: string, botName: string): Promise<void> {
    const record: RetiredBotApp = {
      retiredAt: new Date(this.now()).toISOString(),
      botName,
    };
    this.memory.set(keyFor(msaAppId), record);
    if (this.secrets?.set === undefined) {
      this.log(
        `provisioner botServices.delete: app id '${msaAppId}' recorded as retired in memory only ` +
          '(no vault write access) — the record is lost on restart',
      );
      return;
    }
    try {
      await this.secrets.set(keyFor(msaAppId), JSON.stringify(record));
    } catch (err) {
      this.log(
        `provisioner botServices.delete: could not persist that app id '${msaAppId}' is retired ` +
          `(${String(err)}) — recorded in memory only`,
      );
    }
  }

  /** The record for `msaAppId`, or `undefined` when it never had a bot deleted. */
  async find(msaAppId: string): Promise<RetiredBotApp | undefined> {
    const remembered = this.memory.get(keyFor(msaAppId));
    if (remembered !== undefined) return remembered;
    const stored = await this.secrets?.get(keyFor(msaAppId));
    if (stored === undefined || stored === '') return undefined;
    return parseRecord(stored);
  }
}

function keyFor(msaAppId: string): string {
  return `${TEAMS_RETIRED_BOT_APP_PREFIX}:${msaAppId.toLowerCase()}`;
}

/** A stored record; anything unreadable still means "retired". */
function parseRecord(stored: string): RetiredBotApp {
  try {
    const json: unknown = JSON.parse(stored);
    if (json === null || typeof json !== 'object') return {};
    const record = json as Record<string, unknown>;
    return {
      ...(typeof record['retiredAt'] === 'string' ? { retiredAt: record['retiredAt'] } : {}),
      ...(typeof record['botName'] === 'string' ? { botName: record['botName'] } : {}),
    };
  } catch {
    return {};
  }
}

import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  SINGLE_TENANT_SIGN_IN_AUDIENCE,
  TEAMS_PROVISIONER_CAPABILITY,
  TEAMS_PROVISIONER_SERVICE_NAME,
  type AppRegistration,
  type AzureBot,
  type BotProvisioningOutcome,
  type Idempotent,
  type TenantMode,
} from '../src/teamsProvisioner/types.js';
import {
  ArmNotConfiguredError,
  ConsentMissingError,
  ProvisioningThrottledError,
  TeamsProvisionerError,
} from '../src/teamsProvisioner/errors.js';
import { ConsentRequiredError } from '../src/graphObo.js';

describe('teamsProvisioner@1 constants', () => {
  it('splits service-registry key and manifest capability ref', () => {
    assert.equal(TEAMS_PROVISIONER_SERVICE_NAME, 'teamsProvisioner');
    assert.equal(TEAMS_PROVISIONER_CAPABILITY, 'teamsProvisioner@1');
    // The versioned ref must be the bare key + '@1' so
    // declaredServiceNames()' version-stripping maps them onto each other.
    assert.equal(
      TEAMS_PROVISIONER_CAPABILITY,
      `${TEAMS_PROVISIONER_SERVICE_NAME}@1`,
    );
  });

  it('models the SingleTenant architecture invariant', () => {
    // MultiTenant is deliberately not expressible: the SignInAudience type
    // has exactly one member and this constant is it.
    assert.equal(SINGLE_TENANT_SIGN_IN_AUDIENCE, 'AzureADMyOrg');
    const modes: readonly TenantMode[] = ['customer', 'home'];
    assert.deepEqual(modes, ['customer', 'home']);
  });
});

describe('idempotency and degradation result types', () => {
  const registration: AppRegistration = {
    appId: '11111111-2222-3333-4444-555555555555',
    objectId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    tenantId: '99999999-8888-7777-6666-555555555555',
    tenantMode: 'customer',
    signInAudience: SINGLE_TENANT_SIGN_IN_AUDIENCE,
    displayName: 'HR Agent',
    uniqueName: 'omadia-agent-hr',
  };

  it('Idempotent<T> lets callers branch on outcome without string-matching', () => {
    const rerun: Idempotent<AppRegistration> = {
      outcome: 'already-existed',
      value: registration,
    };
    assert.equal(rerun.outcome, 'already-existed');
    assert.equal(rerun.value.appId, registration.appId);

    const first: Idempotent<AppRegistration> = {
      outcome: 'created',
      value: registration,
    };
    assert.equal(first.outcome, 'created');
  });

  it('BotProvisioningOutcome discriminates provisioned vs registration-only', () => {
    const bot: AzureBot = {
      botName: 'omadia-agent-hr',
      resourceId:
        '/subscriptions/s/resourceGroups/rg/providers/Microsoft.BotService/botServices/omadia-agent-hr',
      msaAppId: registration.appId,
      messagingEndpoint: 'https://example.invalid/api/teams/messages/hr',
    };
    const outcomes: readonly BotProvisioningOutcome[] = [
      { kind: 'provisioned', bot: { outcome: 'created', value: bot } },
      {
        kind: 'registration-only',
        reason: 'arm-not-configured',
        missingSetupFields: ['azure_subscription_id', 'azure_resource_group'],
      },
    ];
    for (const outcome of outcomes) {
      if (outcome.kind === 'provisioned') {
        assert.equal(outcome.bot.value.msaAppId, registration.appId);
      } else {
        // The degradation unit branches here — typed, no error catching.
        assert.equal(outcome.reason, 'arm-not-configured');
        assert.ok(outcome.missingSetupFields.length > 0);
      }
    }
  });
});

describe('error taxonomy', () => {
  it('ConsentMissingError carries the missing scope set and resource', () => {
    const cause = new Error('403 from graph');
    const err = new ConsentMissingError(
      ['Application.ReadWrite.OwnedBy', 'AppCatalog.ReadWrite.All'],
      'graph',
      cause,
    );
    assert.ok(err instanceof Error);
    assert.ok(err instanceof TeamsProvisionerError);
    assert.ok(err instanceof ConsentMissingError);
    assert.equal(err.name, 'ConsentMissingError');
    assert.equal(err.message, 'consent_missing');
    assert.deepEqual(err.missingScopes, [
      'Application.ReadWrite.OwnedBy',
      'AppCatalog.ReadWrite.All',
    ]);
    assert.equal(err.resource, 'graph');
    assert.equal((err as { cause?: unknown }).cause, cause);
  });

  it('ConsentMissingError defaults to the graph resource', () => {
    const err = new ConsentMissingError(['AppCatalog.ReadWrite.All']);
    assert.equal(err.resource, 'graph');
  });

  it('is distinct from graphObo.ts ConsentRequiredError (name collision on purpose)', () => {
    const provisioning = new ConsentMissingError(['AppCatalog.ReadWrite.All']);
    const delegated = new ConsentRequiredError(['Calendars.ReadWrite']);
    assert.notEqual(provisioning.name, delegated.name);
    assert.ok(!(delegated instanceof TeamsProvisionerError));
    assert.ok(!(provisioning instanceof ConsentRequiredError));
  });

  it('ProvisioningThrottledError carries resource and optional Retry-After', () => {
    const err = new ProvisioningThrottledError('arm', 27);
    assert.ok(err instanceof TeamsProvisionerError);
    assert.equal(err.name, 'ProvisioningThrottledError');
    assert.equal(err.message, 'provisioning_throttled');
    assert.equal(err.resource, 'arm');
    assert.equal(err.retryAfterSeconds, 27);

    const noHint = new ProvisioningThrottledError('graph');
    assert.equal(noHint.retryAfterSeconds, undefined);
  });

  it('ArmNotConfiguredError names the missing setup fields', () => {
    const err = new ArmNotConfiguredError(['azure_subscription_id']);
    assert.ok(err instanceof TeamsProvisionerError);
    assert.equal(err.name, 'ArmNotConfiguredError');
    assert.equal(err.message, 'arm_not_configured');
    assert.deepEqual(err.missingSetupFields, ['azure_subscription_id']);
  });
});

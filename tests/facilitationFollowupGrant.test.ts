import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import type {
  FacilitationFollowupService,
  GrantVerificationOutcome,
  GrantVerificationRequest,
  GrantVerificationService,
} from '@omadia/plugin-api';

import {
  GRANT_VERIFICATION_SERVICE_NAME,
  createAppOnlyGraphTransport,
  createFacilitationFollowupService,
} from '../src/facilitationFollowup/index.js';

// The kernel-issued grant (byte5ai/omadia#2078) against a FAKE kernel service
// and a fake Graph. The fake kernel follows the order and the wording of
// `createGrantVerificationService` in the kernel (middleware/src/execution/
// grantService.ts): revoked, scope ended, expired, kill switch, operation not
// covered, recipient not covered.
//
// What this suite must prove:
//   * with a grant id, the kernel is asked before every operation, with the
//     operation and every address it reaches, as the last step before the
//     token endpoint and Graph, and only for a call the local checks passed;
//   * every refusal, and every missing answer, ends the call before the token
//     endpoint and Graph are reached, with code `grant_refused`;
//   * the allow-list of the call still refuses on its own;
//   * a stated grant id on a kernel without the service is refused;
//   * without a grant id nothing changed.

const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const ORGANIZER = 'facilitator@contoso.example';
const ANNA = 'anna@contoso.example';
const BEN = 'ben@contoso.example';
const CARLA = 'carla@contoso.example';
const GRANT_ID = '0b6f6c1e-5a43-4c6b-9a55-2f1c7a3d9e10';
const GRANT = { facilitationId: 'fac-42', allowed: [ANNA, BEN], grantId: GRANT_ID };
const WINDOW = { from: '2026-11-20T08:00:00Z', to: '2026-11-20T17:00:00Z' };

type Operation = GrantVerificationRequest['operation'];

interface KernelGrant {
  kind: 'protocol_dispatch' | 'followup_scheduling' | 'followup_reminders';
  recipients: string[];
  revoked?: boolean;
  scopeEnded?: boolean;
  expired?: boolean;
  killed?: boolean;
}

const KIND_OPERATIONS: Record<KernelGrant['kind'], Operation[]> = {
  protocol_dispatch: ['send_mail'],
  followup_scheduling: ['free_busy', 'create_invitation'],
  followup_reminders: ['send_mail'],
};

function fakeKernel(grants: Record<string, KernelGrant>) {
  const asked: GrantVerificationRequest[] = [];
  const service: GrantVerificationService = {
    verify(request) {
      asked.push(request);
      const refused = (
        code: Exclude<GrantVerificationOutcome, { allowed: true }>['code'],
        message: string,
        notCovered?: string[],
      ): GrantVerificationOutcome => ({
        allowed: false,
        code,
        message,
        ...(notCovered ? { notCovered } : {}),
      });
      const grant = grants[request.grantId];
      if (!grant) return Promise.resolve(refused('unknown_grant', 'no such grant'));
      if (grant.revoked) return Promise.resolve(refused('revoked', 'this grant was revoked'));
      if (grant.scopeEnded) {
        return Promise.resolve(refused('scope_ended', "this grant's execution scope has ended"));
      }
      if (grant.expired) return Promise.resolve(refused('expired', 'this grant has expired'));
      if (grant.killed) return Promise.resolve(refused('kill_switch', 'stopped by an operator'));
      if (!KIND_OPERATIONS[grant.kind].includes(request.operation)) {
        return Promise.resolve(
          refused(
            'operation_not_covered',
            `a '${grant.kind}' grant does not cover '${request.operation}'`,
          ),
        );
      }
      const notCovered = request.recipients.filter((a) => !grant.recipients.includes(a));
      if (notCovered.length > 0) {
        return Promise.resolve(
          refused('recipient_not_covered', 'the grant does not cover every recipient', notCovered),
        );
      }
      return Promise.resolve({ allowed: true, kind: grant.kind, processId: 'proc-1' });
    },
  };
  return { service, asked };
}

function fakeGraph() {
  const graphCalls: string[] = [];
  const tokenCalls: string[] = [];
  const fetchImpl = (async (input: Parameters<typeof fetch>[0]) => {
    const url = String(input);
    if (url.startsWith('https://login.microsoftonline.com/')) {
      tokenCalls.push(url);
      return new Response(JSON.stringify({ access_token: 'app-token', expires_in: 3600 }), {
        status: 200,
      });
    }
    graphCalls.push(url);
    if (url.endsWith('/calendar/getSchedule')) {
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    }
    if (url.endsWith('/events')) {
      return new Response(JSON.stringify({ id: 'event-1' }), { status: 201 });
    }
    return new Response(null, { status: 202 });
  }) as typeof fetch;
  return { fetchImpl, graphCalls, tokenCalls };
}

function setUp(
  lookup: (() => GrantVerificationService | undefined) | undefined,
  timeoutMs?: number,
  isConfigured = true,
) {
  const graph = fakeGraph();
  const logs: string[] = [];
  const service: FacilitationFollowupService = createFacilitationFollowupService({
    organizer: isConfigured
      ? { kind: 'configured', mailbox: ORGANIZER }
      : { kind: 'not-configured', reason: 'setup field is empty' },
    transport: createAppOnlyGraphTransport({
      credential: { tenantId: 'tenant', clientId: 'client', clientSecret: 'secret' },
      fetchImpl: graph.fetchImpl,
    }),
    ...(lookup ? { grantVerification: lookup } : {}),
    ...(timeoutMs !== undefined ? { grantVerificationTimeoutMs: timeoutMs } : {}),
    log: (msg) => logs.push(msg),
  });
  const auditLines = () => logs.filter((l) => l.includes(' op='));
  const untouched = () => {
    assert.deepEqual(graph.graphCalls, [], 'Graph must not be reached');
    assert.deepEqual(graph.tokenCalls, [], 'the token endpoint must not be reached');
  };
  return { service, graph, logs, auditLines, untouched };
}

const mail = (recipients: string[]) => ({ recipients, subject: 'Protocol', body: 'Attached.' });
const invitation = (attendees: string[]) => ({
  attendees,
  subject: 'Follow-up',
  start: '2026-11-25T09:00:00Z',
  end: '2026-11-25T10:00:00Z',
  body: 'As agreed.',
});

describe('facilitation.followup under a kernel-issued grant (0.10.0)', () => {
  it('live grant: asks the kernel before each of the three operations, then calls Graph', async () => {
    const scheduling = fakeKernel({
      [GRANT_ID]: { kind: 'followup_scheduling', recipients: [ANNA, BEN] },
    });
    const a = setUp(() => scheduling.service);
    const busy = await a.service.freeBusy(GRANT, { attendees: [' Anna@Contoso.example ', BEN], ...WINDOW });
    assert.equal(busy.outcome, 'ok');
    const created = await a.service.createInvitation(GRANT, invitation([ANNA, BEN]));
    assert.equal(created.outcome, 'created');
    assert.deepEqual(scheduling.asked, [
      { grantId: GRANT_ID, operation: 'free_busy', recipients: [ANNA, BEN] },
      { grantId: GRANT_ID, operation: 'create_invitation', recipients: [ANNA, BEN] },
    ]);
    assert.equal(a.graph.graphCalls.length, 2);

    const dispatch = fakeKernel({ [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA] } });
    const b = setUp(() => dispatch.service);
    const sent = await b.service.sendMail(GRANT, mail([ANNA]));
    assert.deepEqual(sent, { outcome: 'sent' });
    assert.deepEqual(dispatch.asked, [
      { grantId: GRANT_ID, operation: 'send_mail', recipients: [ANNA] },
    ]);
    assert.deepEqual(b.auditLines(), [
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=1 outcome=sent grant=allowed',
    ]);
  });

  it('asks again at every call: a grant revoked between two sends stops the second', async () => {
    const grants: Record<string, KernelGrant> = {
      [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA] },
    };
    const kernel = fakeKernel(grants);
    const t = setUp(() => kernel.service);
    assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), { outcome: 'sent' });
    grants[GRANT_ID] = { kind: 'protocol_dispatch', recipients: [ANNA], revoked: true };
    const second = await t.service.sendMail(GRANT, mail([ANNA]));
    assert.deepEqual(second, {
      outcome: 'refused',
      code: 'grant_refused',
      message: 'the kernel refused the grant (revoked): this grant was revoked',
    });
    assert.equal(kernel.asked.length, 2);
    assert.equal(t.graph.graphCalls.length, 1, 'only the first send reached Graph');
    assert.equal(
      t.auditLines()[1],
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=1 outcome=grant_refused grant=refused:revoked',
    );
  });

  for (const [name, state, code] of [
    ['revoked', { revoked: true }, 'revoked'],
    ['expired', { expired: true }, 'expired'],
    ['scope ended', { scopeEnded: true }, 'scope_ended'],
    ['operator stop', { killed: true }, 'kill_switch'],
  ] as const) {
    it(`${name}: all three operations are refused before Graph`, async () => {
      const kernel = fakeKernel({
        [GRANT_ID]: { kind: 'followup_scheduling', recipients: [ANNA, BEN], ...state },
      });
      const t = setUp(() => kernel.service);
      const answers = [
        await t.service.freeBusy(GRANT, { attendees: [ANNA], ...WINDOW }),
        await t.service.createInvitation(GRANT, invitation([ANNA])),
        await t.service.sendMail(GRANT, mail([ANNA])),
      ];
      for (const answer of answers) {
        assert.equal(answer.outcome, 'refused');
        assert.equal(answer.outcome === 'refused' && answer.code, 'grant_refused');
        assert.match(answer.outcome === 'refused' ? answer.message : '', new RegExp(`\\(${code}\\)`));
      }
      assert.equal(kernel.asked.length, 3);
      t.untouched();
      assert.ok(t.auditLines().every((l) => l.endsWith(`outcome=grant_refused grant=refused:${code}`)));
    });
  }

  it('unknown grant id: refused', async () => {
    const kernel = fakeKernel({});
    const t = setUp(() => kernel.service);
    const answer = await t.service.sendMail(GRANT, mail([ANNA]));
    assert.equal(answer.outcome === 'refused' && answer.code, 'grant_refused');
    assert.match(answer.outcome === 'refused' ? answer.message : '', /\(unknown_grant\)/);
    t.untouched();
  });

  it('operation not covered: a dispatch grant sends mail but creates no invitation and reads no calendar', async () => {
    const kernel = fakeKernel({ [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA, BEN] } });
    const t = setUp(() => kernel.service);
    const invited = await t.service.createInvitation(GRANT, invitation([ANNA]));
    assert.deepEqual(invited, {
      outcome: 'refused',
      code: 'grant_refused',
      message:
        "the kernel refused the grant (operation_not_covered): a 'protocol_dispatch' grant does not cover 'create_invitation'",
    });
    const busy = await t.service.freeBusy(GRANT, { attendees: [ANNA], ...WINDOW });
    assert.equal(busy.outcome === 'refused' && busy.code, 'grant_refused');
    t.untouched();
    assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), { outcome: 'sent' });
  });

  it('recipient not covered by the kernel grant: whole call refused, addresses named, although the call allow-list has them', async () => {
    const kernel = fakeKernel({ [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA] } });
    const t = setUp(() => kernel.service);
    const answer = await t.service.sendMail(GRANT, mail([ANNA, BEN]));
    assert.deepEqual(answer, {
      outcome: 'refused',
      code: 'grant_refused',
      message: 'the kernel refused the grant (recipient_not_covered): the grant does not cover every recipient',
      rejected: [BEN],
    });
    t.untouched();
    assert.deepEqual(t.auditLines(), [
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=2 outcome=grant_refused grant=refused:recipient_not_covered',
    ]);
  });

  it('local checks come first: a call they refuse never reaches the kernel, so no use is recorded', async () => {
    const kernel = fakeKernel({
      [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA, BEN, CARLA] },
    });
    const t = setUp(() => kernel.service);
    // The kernel grant covers Carla, the allow-list of the call does not.
    const outside = await t.service.sendMail(GRANT, mail([ANNA, CARLA]));
    assert.equal(outside.outcome === 'refused' && outside.code, 'recipient_not_allowed');
    assert.deepEqual(outside.outcome === 'refused' ? outside.rejected : [], [CARLA]);
    const empty = await t.service.sendMail(GRANT, { recipients: [ANNA], subject: '', body: 'x' });
    assert.equal(empty.outcome === 'refused' && empty.code, 'provider_error');
    const noList = await t.service.sendMail(GRANT, mail(undefined as unknown as string[]));
    assert.equal(noList.outcome === 'refused' && noList.code, 'provider_error');
    const window = await t.service.freeBusy(GRANT, { attendees: [ANNA], from: 'tomorrow', to: WINDOW.to });
    assert.equal(window.outcome === 'refused' && window.code, 'provider_error');

    const unconfigured = setUp(() => kernel.service, undefined, false);
    const noMailbox = await unconfigured.service.createInvitation(GRANT, invitation([ANNA]));
    assert.equal(noMailbox.outcome === 'refused' && noMailbox.code, 'not_configured');

    assert.deepEqual(kernel.asked, [], 'the kernel must not record a use for these calls');
    t.untouched();
    unconfigured.untouched();
    assert.ok(
      [...t.auditLines(), ...unconfigured.auditLines()].every((l) =>
        l.endsWith('grant=not_asked:refused_locally'),
      ),
    );
    assert.equal(
      t.auditLines()[0],
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=2 outcome=recipient_not_allowed grant=not_asked:refused_locally',
    );
  });

  it('the kernel is the last step before the token and Graph requests', async () => {
    const order: string[] = [];
    const kernel = fakeKernel({ [GRANT_ID]: { kind: 'followup_scheduling', recipients: [ANNA, BEN] } });
    const t = setUp(() => ({
      verify: (request) => {
        order.push(`kernel:${String(t.graph.tokenCalls.length + t.graph.graphCalls.length)}`);
        return kernel.service.verify(request);
      },
    }));
    await t.service.freeBusy(GRANT, { attendees: [ANNA, BEN], ...WINDOW });
    assert.deepEqual(order, ['kernel:0'], 'asked once, before any request left the process');
    assert.equal(t.graph.graphCalls.length, 1);
  });

  describe('kernel unreachable = refuse', () => {
    const cases: [string, () => GrantVerificationService | undefined, RegExp][] = [
      [
        'verify() answers unavailable',
        () => ({
          verify: () =>
            Promise.resolve({ allowed: false, code: 'unavailable', message: 'the grant could not be verified' }),
        }),
        /\(unavailable\)/,
      ],
      [
        'verify() rejects',
        () => ({ verify: () => Promise.reject(new Error('connection terminated')) }),
        /could not be verified with the kernel \(connection terminated\)/,
      ],
      [
        'verify() throws synchronously',
        () => ({
          verify: () => {
            throw new Error('boom');
          },
        }),
        /could not be verified with the kernel/,
      ],
      [
        'verify() never answers',
        () => ({ verify: () => new Promise<GrantVerificationOutcome>(() => undefined) }),
        /no answer within 20 ms/,
      ],
      [
        'verify() answers something else',
        () => ({ verify: () => Promise.resolve({ ok: true } as unknown as GrantVerificationOutcome) }),
        /no usable answer/,
      ],
      [
        'the service lookup throws',
        () => {
          throw new Error("service 'grantVerification' is not declared");
        },
        /service lookup failed/,
      ],
    ];
    for (const [name, lookup, message] of cases) {
      it(name, async () => {
        const t = setUp(lookup, 20);
        const answers = [
          await t.service.freeBusy(GRANT, { attendees: [ANNA], ...WINDOW }),
          await t.service.createInvitation(GRANT, invitation([ANNA])),
          await t.service.sendMail(GRANT, mail([ANNA])),
        ];
        for (const answer of answers) {
          assert.equal(answer.outcome === 'refused' && answer.code, 'grant_refused');
          assert.match(answer.outcome === 'refused' ? answer.message : '', message);
        }
        t.untouched();
        assert.ok(t.auditLines().every((l) => l.endsWith('grant=refused:unavailable')));
      });
    }

    it('a service that was there and is gone refuses', async () => {
      const kernel = fakeKernel({ [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA] } });
      let present = true;
      const t = setUp(() => (present ? kernel.service : undefined));
      assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), { outcome: 'sent' });
      present = false;
      const answer = await t.service.sendMail(GRANT, mail([ANNA]));
      assert.equal(answer.outcome === 'refused' && answer.code, 'grant_refused');
      assert.equal(t.graph.graphCalls.length, 1);
    });
  });

  it('a grant id that is not a string is refused without asking', async () => {
    const kernel = fakeKernel({});
    const t = setUp(() => kernel.service);
    const answer = await t.service.sendMail(
      { facilitationId: 'fac-42', allowed: [ANNA], grantId: 7 as unknown as string },
      mail([ANNA]),
    );
    assert.equal(answer.outcome === 'refused' && answer.code, 'grant_refused');
    assert.equal(kernel.asked.length, 0);
    t.untouched();
  });

  it('the kernel message is redacted and cut before it is passed on', async () => {
    const secret = `eyJ${'A1b2'.repeat(40)}`;
    const t = setUp(() => ({
      verify: () =>
        Promise.resolve({ allowed: false, code: 'unavailable', message: `db said ${secret} ${'x '.repeat(400)}` }),
    }));
    const answer = await t.service.sendMail(GRANT, mail([ANNA]));
    const message = answer.outcome === 'refused' ? answer.message : '';
    assert.ok(!message.includes(secret));
    assert.ok(message.length < 300);
  });
});

describe('facilitation.followup on a kernel without grantVerification', () => {
  const NO_SERVICE = {
    outcome: 'refused',
    code: 'grant_refused',
    message:
      "the kernel offers no grant verification (no 'grantVerification' service), so the stated grant cannot be checked; nothing was sent",
  };

  it('a stated grant id is refused for all three operations: it cannot be checked', async () => {
    let lookups = 0;
    const t = setUp(() => {
      lookups += 1;
      return undefined;
    });
    assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), NO_SERVICE);
    assert.deepEqual(await t.service.freeBusy(GRANT, { attendees: [ANNA], ...WINDOW }), NO_SERVICE);
    assert.deepEqual(await t.service.createInvitation(GRANT, invitation([ANNA])), NO_SERVICE);
    assert.equal(lookups, 3, 'the service is looked up at every call');
    t.untouched();
    assert.ok(
      t.auditLines().every((l) => l.endsWith('outcome=grant_refused grant=refused:no_kernel_service')),
    );
  });

  it('no lookup wired at all behaves the same', async () => {
    const t = setUp(undefined);
    assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), NO_SERVICE);
    t.untouched();
  });

  it('once the kernel publishes the service, the same call goes through', async () => {
    const kernel = fakeKernel({ [GRANT_ID]: { kind: 'protocol_dispatch', recipients: [ANNA] } });
    let published = false;
    const t = setUp(() => (published ? kernel.service : undefined));
    assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), NO_SERVICE);
    published = true;
    assert.deepEqual(await t.service.sendMail(GRANT, mail([ANNA])), { outcome: 'sent' });
  });

  it('caller without a grant id: allow-list only, as before 0.10.0', async () => {
    const t = setUp(() => undefined);
    const legacy = { facilitationId: 'fac-42', allowed: [ANNA] };
    assert.deepEqual(await t.service.sendMail(legacy, mail([ANNA])), { outcome: 'sent' });
    const outside = await t.service.sendMail(legacy, mail([CARLA]));
    assert.equal(outside.outcome === 'refused' && outside.code, 'recipient_not_allowed');
    assert.deepEqual(
      t.auditLines().map((l) => l.split(' grant=')[1]),
      ['not_asked:no_grant_id', 'not_asked:no_grant_id'],
    );
  });
});

describe('facilitation.followup without a grant id on a kernel with grants', () => {
  it('the kernel is not asked', async () => {
    const kernel = fakeKernel({});
    const t = setUp(() => kernel.service);
    const legacy = { facilitationId: 'fac-42', allowed: [ANNA] };
    assert.deepEqual(await t.service.sendMail(legacy, mail([ANNA])), { outcome: 'sent' });
    assert.equal(kernel.asked.length, 0);
    assert.deepEqual(t.auditLines(), [
      '[microsoft365] facilitation.followup op=sendMail facilitation=fac-42 people=1 outcome=sent grant=not_asked:no_grant_id',
    ]);
  });
});

describe('grantVerification wiring', () => {
  it('the manifest declares the service as optional and the plugin looks it up per call', () => {
    const manifest = readFileSync(join(pkgRoot, 'manifest.yaml'), 'utf8');
    assert.match(manifest, /^optional_requires:\n {2}- "grantVerification@\^1"$/m);
    assert.ok(!/^requires:/m.test(manifest), 'a hard requirement would keep the plugin off an older kernel');
    const plugin = readFileSync(join(pkgRoot, 'src', 'plugin.ts'), 'utf8');
    assert.match(
      plugin,
      /grantVerification: \(\) =>\s+ctx\.services\.get<GrantVerificationService>\(GRANT_VERIFICATION_SERVICE_NAME\)/,
    );
    assert.equal(GRANT_VERIFICATION_SERVICE_NAME, 'grantVerification');
  });
});

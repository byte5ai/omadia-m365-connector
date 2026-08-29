import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  classifyInstallTarget,
  isChatTarget,
} from '../src/teamsProvisioner/installTarget.js';

// Pure shape classification — no network, no Graph. The two cases that carry
// their weight here are the CHANNEL id (well-formed, and never an install
// target) and the bare 32-hex string that is a team GUID and a chat thread
// body at the same time. Both come from the byte5 field test.

describe('classifyInstallTarget — teams', () => {
  it('classifies a dashed GUID as a team and hands back the id', () => {
    const target = classifyInstallTarget(
      'abc8af8e-c7fc-4717-85d3-b83c4d84b667',
    );
    assert.equal(target.kind, 'team');
    assert.equal(
      target.kind === 'team' ? target.teamId : undefined,
      'abc8af8e-c7fc-4717-85d3-b83c4d84b667',
    );
    assert.equal(isChatTarget(target), false);
  });

  it('accepts an upper-case GUID', () => {
    assert.equal(
      classifyInstallTarget('ABC8AF8E-C7FC-4717-85D3-B83C4D84B667').kind,
      'team',
    );
  });

  it('trims surrounding whitespace (ids arrive by copy-paste)', () => {
    const target = classifyInstallTarget(
      '  abc8af8e-c7fc-4717-85d3-b83c4d84b667\n',
    );
    assert.equal(target.kind, 'team');
    assert.equal(
      target.kind === 'team' ? target.teamId : undefined,
      'abc8af8e-c7fc-4717-85d3-b83c4d84b667',
    );
  });
});

describe('classifyInstallTarget — chats', () => {
  it('classifies 19:…@thread.v2 as a group chat', () => {
    const target = classifyInstallTarget('19:abc123def456@thread.v2');
    assert.equal(target.kind, 'group-chat');
    assert.ok(isChatTarget(target));
    assert.equal(target.chatId, '19:abc123def456@thread.v2');
  });

  it('classifies 19:…@unq.gbl.spaces as a 1:1 chat', () => {
    const target = classifyInstallTarget('19:aaa_bbb@unq.gbl.spaces');
    assert.equal(target.kind, 'one-on-one-chat');
    assert.ok(isChatTarget(target));
    assert.equal(target.chatId, '19:aaa_bbb@unq.gbl.spaces');
  });

  it('matches the suffix case-insensitively', () => {
    assert.equal(
      classifyInstallTarget('19:ABC@THREAD.V2').kind,
      'group-chat',
    );
  });
});

describe('classifyInstallTarget — the channel trap', () => {
  it('classifies 19:…@thread.tacv2 as a CHANNEL, never a group chat', () => {
    const target = classifyInstallTarget('19:abc123@thread.tacv2');
    assert.equal(target.kind, 'channel');
    assert.equal(isChatTarget(target), false);
  });

  it('says an app goes into the TEAM, not the channel', () => {
    const target = classifyInstallTarget('19:abc123@thread.tacv2');
    assert.ok(target.kind === 'channel');
    const hint = target.hint.toLowerCase();
    // The whole point: not a generic "not found", but the actual remedy.
    assert.match(hint, /channel/);
    assert.match(hint, /team id/);
  });

  it('does not confuse the tacv2 suffix with the v2 suffix', () => {
    // '…@thread.tacv2' must NOT satisfy the '@thread.v2' pattern.
    assert.notEqual(
      classifyInstallTarget('19:x@thread.tacv2').kind,
      'group-chat',
    );
    assert.notEqual(classifyInstallTarget('19:x@thread.v2').kind, 'channel');
  });
});

describe('classifyInstallTarget — the ambiguous 32-hex form', () => {
  // The value an operator actually pasted in the byte5 field test.
  const FIELD_TEST_VALUE = 'abc8af8ec7fc471785d3b83c4d84b667';

  it('reports ambiguity instead of guessing', () => {
    const target = classifyInstallTarget(FIELD_TEST_VALUE);
    assert.equal(target.kind, 'ambiguous');
    assert.equal(
      target.kind === 'ambiguous' ? target.value : undefined,
      FIELD_TEST_VALUE,
    );
  });

  it('names both readings as candidates', () => {
    const target = classifyInstallTarget(FIELD_TEST_VALUE);
    assert.ok(target.kind === 'ambiguous');
    assert.deepEqual([...target.candidates].sort(), ['group-chat', 'team']);
  });

  it('tells the operator which full form to re-enter', () => {
    const target = classifyInstallTarget(FIELD_TEST_VALUE);
    assert.ok(target.kind === 'ambiguous');
    assert.match(target.hint, /thread\.v2/);
    assert.match(target.hint, /8-4-4-4-12/);
  });

  it('is not a chat target — the caller must decide first', () => {
    assert.equal(isChatTarget(classifyInstallTarget(FIELD_TEST_VALUE)), false);
  });

  it('needs exactly 32 hex digits — 31 or 33 are unknown', () => {
    assert.equal(classifyInstallTarget('a'.repeat(31)).kind, 'unknown');
    assert.equal(classifyInstallTarget('a'.repeat(33)).kind, 'unknown');
  });

  it('needs HEX — 32 non-hex characters are unknown', () => {
    assert.equal(classifyInstallTarget('z'.repeat(32)).kind, 'unknown');
  });
});

describe('classifyInstallTarget — unknown shapes', () => {
  it('answers unknown for the empty string instead of throwing', () => {
    assert.equal(classifyInstallTarget('').kind, 'unknown');
    assert.equal(classifyInstallTarget('   ').kind, 'unknown');
  });

  it('answers unknown for a team display name', () => {
    const target = classifyInstallTarget('Marketing');
    assert.equal(target.kind, 'unknown');
    assert.ok(target.kind === 'unknown');
    assert.match(target.hint, /thread\.v2/);
  });

  it('gives a legacy @thread.skype id its own remedy text', () => {
    const target = classifyInstallTarget('19:abc@thread.skype');
    // Still 'unknown' — the contract lists exactly four known shapes.
    assert.equal(target.kind, 'unknown');
    assert.ok(target.kind === 'unknown');
    assert.match(target.hint.toLowerCase(), /skype/);
  });

  it('rejects a 19: prefix with no recognised suffix', () => {
    assert.equal(classifyInstallTarget('19:abc123').kind, 'unknown');
  });
});

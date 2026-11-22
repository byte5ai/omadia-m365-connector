import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { GraphClient, GraphRequestError } from '../src/graphClient.js';

// Chat vs. channel message lookup (live finding 2026-11-21): the Teams plugin
// asked `/chats/{id}/messages/{id}` for a channel thread and got 404 "Invalid
// ThreadId". A channel message lives under /teams/{team}/channels/{channel}.

function client(calls: string[], status = 200): GraphClient {
  const fetchImpl = (async (url: string | URL) => {
    const u = String(url);
    if (u.includes('login.microsoftonline.com')) {
      return new Response(JSON.stringify({ access_token: 'tok', expires_in: 3600 }), { status: 200 });
    }
    calls.push(u);
    const body = status === 200
      ? { id: 'm', attachments: [{ contentType: 'reference', contentUrl: 'https://contoso.sharepoint.com/f.pdf', name: 'f.pdf' }] }
      : { error: { code: 'NotFound', message: 'Invalid ThreadId' } };
    return new Response(JSON.stringify(body), { status });
  }) as typeof fetch;
  return new GraphClient({ tenantId: 't', clientId: 'c', clientSecret: 's', fetchImpl, log: () => {} });
}

describe('GraphClient channel message lookup', () => {
  it('reads a thread root under the team and channel', async () => {
    const calls: string[] = [];
    const msg = await client(calls).fetchChannelMessage({
      teamId: 'team-guid',
      channelId: '19:abc@thread.tacv2',
      rootMessageId: '1700000000000',
      messageId: '1700000000000',
    });
    assert.deepEqual(calls, [
      'https://graph.microsoft.com/v1.0/teams/team-guid/channels/19%3Aabc%40thread.tacv2/messages/1700000000000',
    ]);
    assert.equal(msg.attachments.length, 1);
    assert.equal(msg.attachments[0]?.name, 'f.pdf');
  });

  it('reads a reply under its root', async () => {
    const calls: string[] = [];
    await client(calls).fetchChannelMessage({
      teamId: 'team-guid',
      channelId: '19:abc@thread.tacv2',
      rootMessageId: '1700000000000',
      messageId: '1700000000999',
    });
    assert.deepEqual(calls, [
      'https://graph.microsoft.com/v1.0/teams/team-guid/channels/19%3Aabc%40thread.tacv2/messages/1700000000000/replies/1700000000999',
    ]);
  });

  it('throws with the status on a non-2xx answer', async () => {
    await assert.rejects(
      client([], 404).fetchChannelMessage({ teamId: 't', channelId: 'c', rootMessageId: 'r', messageId: 'r' }),
      /graph channelMessage 404/,
    );
  });

  it('marks a 403 as not permitted, keeping the message format', async () => {
    await assert.rejects(
      client([], 403).fetchChannelMessage({ teamId: 't', channelId: 'c', rootMessageId: 'r', messageId: 'r' }),
      (err: unknown) =>
        err instanceof GraphRequestError &&
        err.status === 403 &&
        err.code === 'not_permitted' &&
        /^graph channelMessage 403 /.test(err.message),
    );
  });

  it('a 404 carries its status but no not_permitted code', async () => {
    await assert.rejects(
      client([], 404).fetchChannelMessage({ teamId: 't', channelId: 'c', rootMessageId: 'r', messageId: 'r' }),
      (err: unknown) => err instanceof GraphRequestError && err.status === 404 && err.code === undefined,
    );
  });

  it('chat lookup is unchanged', async () => {
    const calls: string[] = [];
    const msg = await client(calls).fetchChatMessage('19:x@thread.v2', '42');
    assert.deepEqual(calls, ['https://graph.microsoft.com/v1.0/chats/19%3Ax%40thread.v2/messages/42']);
    assert.equal(msg.attachments.length, 1);
  });
});

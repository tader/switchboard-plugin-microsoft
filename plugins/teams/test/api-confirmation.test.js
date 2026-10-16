import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { TeamsAPI, encodeConversation } from '../lib/teams-api.js';
import { SendLedger } from '../lib/ledger.js';

const thread = '19:verification@thread.v2', chatId = encodeConversation('chat', thread);
function fixture({ nativeQuoteSeparators = false } = {}) {
  let raw = { id: '100', content: 'Before', messagetype: 'Text', from: '8:orgid:self', clientmessageid: '101', version: '1', properties: {} };
  let outcome = 'observed', error = null, gets = 0;
  const writes = [], session = { credentials: { identity: { oid: 'self', tenant: 'tenant', name: 'Self' },
    messageOrigin: 'https://emea.ng.msg.teams.microsoft.com', chatToken: { value: 'test' }, tokens: { chatsvcagg: { value: 'test' } } } };
  const api = new TeamsAPI(session, { fetcher: async (url, options) => {
    const pathname = new URL(url).pathname;
    if (pathname.endsWith('/teams/users/me')) return Response.json({ chats: [{ id: thread }], teams: [] });
    if (options.method && options.method !== 'GET') {
      writes.push(options);
      if (options.method === 'POST') { const payload = JSON.parse(options.body); if (outcome === 'observed') {
        raw = { ...raw, ...payload };
        if (nativeQuoteSeparators) raw.content = raw.content
          .replace(/(<blockquote[^>]*>)\s*(<strong)/, '$1\r\n$2')
          .replace(/(<\/strong>)\s*(<p itemprop="preview">)/, '$1\r\n$2')
          .replace(/(<\/p>)\s*(<\/blockquote>)/, '$1\r\n$2')
          .replace(/(<\/blockquote>)\s*(<p>)/, '$1\r\n$2');
      } return Response.json({ OriginalArrivalTime: 100 }); }
      if (outcome === 'observed') {
        if (pathname.endsWith('/properties')) { const key = JSON.parse(options.body).emotions.key; raw.properties.emotions = options.method === 'DELETE' ? [] : [{ key, users: [{ mri: '8:orgid:self' }] }]; }
        else if (options.method === 'DELETE') raw = { ...raw, messagetype: 'MessageDelete' };
        else { raw = { ...raw, ...JSON.parse(options.body), version: '2' }; }
      }
      return new Response(null, { status: 204 });
    }
    gets++; if (error) return new Response(null, { status: error });
    if (pathname.endsWith('/messages')) return Response.json({ messages: [raw], _metadata: {} });
    return Response.json(raw);
  } });
  return { api, session, writes, get raw() { return raw; }, get gets() { return gets; }, setOutcome: value => outcome = value, setError: value => error = value };
}
async function ledger(t) { const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-confirm-')); t.after(() => fs.rm(dir, { recursive: true, force: true })); return new SendLedger(dir); }

test('send readback runs after durable acceptance, checks exact client/content identity, and never resends on replay', async t => {
  const f = fixture(), journal = await ledger(t), input = { text: 'After' };
  const verify = async (result, context) => {
    const saved = Object.values(await journal.read())[0];
    assert.equal(saved.result.status, 'accepted_by_api'); assert.ok(saved.confirmation.contentStamp);
    assert.ok(!JSON.stringify(saved.confirmation).includes('After'));
    return f.api.confirm(result, context);
  };
  const result = await journal.run('confirmed-send-key', input, () => f.api.prepareSend(chatId, input), p => f.api.send(p), 'send_uncertain', verify);
  assert.equal(result.verification.status, 'observed'); assert.equal(result.deliveryConfirmed, false);
  const count = f.gets;
  const replay = await new SendLedger(path.dirname(journal.file)).run('confirmed-send-key', input, () => assert.fail('No preparation'), () => assert.fail('No send'), 'send_uncertain', verify);
  assert.equal(replay.replayed, true); assert.equal(f.gets, count); assert.equal(f.writes.length, 1);
});

test('unobserved and failed readbacks retain accepted receipts and can retry reads after restart without a write', async t => {
  const f = fixture(), journal = await ledger(t), input = { text: 'After' }; f.setOutcome('lagging');
  const first = await journal.run('lagging-send-key', input, () => f.api.prepareSend(chatId, input), p => f.api.send(p), 'send_uncertain', (r, c) => f.api.confirm(r, c));
  assert.equal(first.status, 'accepted_by_api'); assert.equal(first.verification.status, 'not_observed'); assert.equal(f.gets, 2);
  const payload = JSON.parse(f.writes[0].body); Object.assign(f.raw, payload);
  const next = await new SendLedger(path.dirname(journal.file)).run('lagging-send-key', input, () => assert.fail('No preparation'), () => assert.fail('No resend'), 'send_uncertain', (r, c) => f.api.confirm(r, c));
  assert.equal(next.verification.status, 'observed'); assert.equal(f.writes.length, 1);
  const prepared = await f.api.prepareSend(chatId, { text: 'Another' }), accepted = await f.api.send(prepared);
  f.setError(401); const failed = await f.api.confirm(accepted, prepared.confirmation);
  assert.equal(failed.status, 'unavailable'); assert.equal(failed.reason, 'api_unauthorized');
});

test('same-text messages from another sender/client and mismatching rich content do not confirm a send', async () => {
  const f = fixture(), prepared = await f.api.prepareSend(chatId, { text: 'After' }), accepted = await f.api.send(prepared);
  f.raw.clientmessageid = 'unrelated'; assert.equal((await f.api.confirm(accepted, prepared.confirmation)).status, 'not_observed');
  f.raw.clientmessageid = accepted.message.clientMessageId; f.raw.from = '8:orgid:other';
  assert.equal((await f.api.confirm(accepted, prepared.confirmation)).status, 'not_observed');
  f.raw.from = '8:orgid:self'; f.raw.messagetype = 'RichText/Html';
  assert.equal((await f.api.confirm(accepted, prepared.confirmation)).status, 'not_observed');
  f.session.credentials.identity.oid = 'different-account';
  assert.equal((await f.api.confirm(accepted, prepared.confirmation)).reason, 'verification_context_unavailable');
});

test('edit/delete/reaction readbacks verify desired state; 404 is not proof of deletion', async () => {
  for (const [kind, input] of [['edit', { expectedText: 'Before', text: 'After' }], ['delete', { expectedText: 'Before' }], ['reaction', { reaction: 'like', selected: true }]]) {
    const f = fixture(), prepared = await f.api.prepareMutation(chatId, '100', kind, input), result = await f.api.mutate(prepared);
    assert.equal((await f.api.confirm(result, prepared.confirmation)).status, 'observed');
    if (kind === 'delete') { f.setError(404); assert.equal((await f.api.confirm(result, prepared.confirmation)).reason, 'api_not_found'); }
    if (kind === 'reaction') {
      const remove = await f.api.prepareMutation(chatId, '100', 'reaction', { reaction: 'like', selected: false });
      assert.equal((await f.api.confirm(await f.api.mutate(remove), remove.confirmation)).status, 'observed');
      f.raw.properties.emotions = 'malformed'; assert.equal((await f.api.confirm(result, remove.confirmation)).status, 'unavailable');
    }
  }
});

test('confirmation persistence failure cannot erase the durably accepted result or trigger resend', async t => {
  const f = fixture(), journal = await ledger(t), write = journal.write.bind(journal); let count = 0;
  journal.write = records => ++count === 3 ? Promise.reject(new Error('disk full')) : write(records);
  const input = { text: 'After' }, result = await journal.run('disk-full-confirmation', input, () => f.api.prepareSend(chatId, input), p => f.api.send(p), 'send_uncertain', (r, c) => f.api.confirm(r, c));
  assert.equal(result.status, 'accepted_by_api'); assert.equal(result.verification.reason, 'verification_persistence_failed');
  const replay = await new SendLedger(path.dirname(journal.file)).run('disk-full-confirmation', input, () => assert.fail('No prepare'), () => assert.fail('No write'), 'send_uncertain', (r, c) => f.api.confirm(r, c));
  assert.equal(replay.verification.status, 'observed'); assert.equal(f.writes.length, 1);
});

test('legacy accepted records replay without inventing a confirmation target', async t => {
  const journal = await ledger(t);
  await journal.run('legacy-accepted-key', ['legacy'], async () => ({}), async () => ({ status: 'accepted_by_api' }), 'send_uncertain');
  const result = await journal.run('legacy-accepted-key', ['legacy'], () => assert.fail(), () => assert.fail(), 'send_uncertain', async (_r, context) => {
    assert.equal(context, undefined); return { status: 'unavailable', reason: 'verification_context_unavailable' };
  });
  assert.equal(result.replayed, true); assert.equal(result.status, 'accepted_by_api');
});

test('reaction availability includes observed custom keys with truthful source and mutability metadata', async () => {
  const f = fixture(); f.raw.properties.emotions = [{ key: 'celebrate', users: [{ mri: '8:orgid:self' }] }, { key: '😀', users: [{ mri: '8:orgid:peer' }] }];
  const result = await f.api.reactions(chatId, '100');
  assert.equal(result.availabilityComplete, false);
  assert.deepEqual(result.available.find(x => x.reaction === 'celebrate'), { reaction: 'celebrate', selected: true, source: 'observed_on_message', canSet: true });
  assert.equal(result.available.find(x => x.reaction === '😀').canSet, false);
  assert.equal(result.available.find(x => x.reaction === 'heart').source, 'adapter_preset');
  f.raw.properties.emotions = 'not-json'; await assert.rejects(f.api.reactions(chatId, '100'), { code: 'api_schema_changed' });
});


test('native quote block separators preserve exact send confirmation', async t => {
  const f = fixture({ nativeQuoteSeparators: true }), journal = await ledger(t);
  const input = { text: 'Quoted reply', replyToMessageId: '100' };
  const result = await journal.run('native-quote-confirmation', input,
    () => f.api.prepareSend(chatId, input), p => f.api.send(p), 'send_uncertain',
    (r, c) => f.api.confirm(r, c));
  assert.equal(result.verification.status, 'observed');
  assert.match(f.raw.content, /<\/blockquote>\r\n<p>Quoted reply<\/p>$/);
  const context = Object.values(await journal.read())[0].confirmation;
  f.raw.content = f.raw.content.replace('itemid="100"', 'itemid="999"');
  assert.equal((await f.api.confirm(result, context)).status, 'not_observed');
  f.raw.content = f.raw.content.replace('itemid="999"', 'itemid="100"').replace('Quoted reply', 'Changed reply');
  assert.equal((await f.api.confirm(result, context)).status, 'not_observed');
  assert.equal(f.writes.length, 1);
});

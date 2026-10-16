import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { TeamsAPI, encodeConversation, messagePayload } from '../lib/teams-api.js';
import { TeamsSession } from '../lib/api-auth.js';
import { SendLedger } from '../lib/ledger.js';
import { openapi } from '../lib/api-openapi.js';
import { encodeTokenBundle, validateTokenBundle } from '../lib/token-bundle.js';
import { signedBundle, signedToken, renewal, oid } from './token-fixture.js';
const peer = '33333333-3333-3333-3333-333333333333';
const chat = '19:chat@thread.v2', channel = '19:channel@thread.tacv2';
const chatId = encodeConversation('chat', chat), channelId = encodeConversation('channel', channel);
const message = (id, text, parent) => ({ id, content: text, messagetype: 'Text', from: `8:orgid:${peer}`, imdisplayname: 'Peer', originalarrivaltime: Number(id), version: '1', properties: {}, ...(parent ? { parentMessageId: parent } : {}) });
function fixture() {
  const requests = [], credentials = { ...signedBundle(), tokens: { ...signedBundle().tokens, substrate: { value: signedToken('skype', { aud: 'https://substrate.office.com' }), expiresAt: Date.now() + 3600_000 } },
    messageOrigin: 'https://emea.ng.msg.teams.microsoft.com', region: 'emea', chatToken: { value: 'fixture-chat' } };
  let existing = false, changed = false, badPage = false, pageHost = null, searchShape = false;
  const response = value => new Response(JSON.stringify(value));
  const fetcher = async (raw, options) => {
    const url = new URL(raw), decoded = decodeURIComponent(url.pathname); requests.push({ url, ...options });
    if (url.pathname.endsWith('/teams/users/me')) return response({ chats: [{ id: chat, isOneOnOne: true, members: existing ? [{ mri: `8:orgid:${oid}` }, { mri: `8:orgid:${peer}` }] : [] }], teams: [{ channels: [{ id: channel, isMember: true }] }] });
    if (decoded.includes('/users/peer@example.com/')) return response({ value: { email: 'peer@example.com', mri: `8:orgid:${peer}`, displayName: 'Peer & Co' } });
    if (url.pathname.endsWith('/suggestions')) return response({ Groups: [{ Type: 'Chat', Suggestions: [{ Name: 'Not a person' }] }, { Type: 'People', Suggestions: [
      { DisplayName: 'Peer & Co', MRI: `8:orgid:${peer}`, EmailAddresses: ['peer@example.com'] },
      { DisplayName: 'Duplicate', MRI: `8:orgid:${peer}`, EmailAddresses: ['peer@example.com'] }, { DisplayName: 'No email' },
    ] }] });
    if (url.pathname.endsWith('/v2/query')) {
      if (searchShape) return response({ unexpected: [] });
      const from = JSON.parse(options.body).EntityRequests[0].from;
      const hits = from ? [{ Source: { WebUrl: `https://teams.cloud.microsoft/l/message/${encodeURIComponent(channel)}/100?parentMessageId=100`, Preview: 'root' } }] : [
        { Source: { WebUrl: `https://teams.microsoft.com/l/message/${encodeURIComponent(chat)}/101`, Preview: '<b>match</b>' } },
        { Source: { WebUrl: `https://teams.microsoft.com/l/message/${encodeURIComponent(chat)}/101`, Preview: 'duplicate' } },
      ];
      return response({ EntitySets: [{ ResultSets: [{ Results: hits, MoreResultsAvailable: !from }] }] });
    }
    if (decoded.endsWith('/messages/100')) return response(message('100', 'root'));
    if (decoded.endsWith('/messages/101')) return response(message('101', changed ? 'changed' : '<unsafe>'));
    if (decoded.endsWith('/messages/199')) return response(message('199', 'reply', '100'));
    if (decoded.endsWith('/messages') && options.method === 'POST') return response({ OriginalArrivalTime: 999 });
    if (decoded.endsWith('/messages') && (options.method ?? 'GET') === 'GET') {
      const thread = decoded.includes(';messageid=100');
      const older = url.searchParams.get('startTime') === '2';
      const items = thread ? older ? [message('102', 'overlap', '100'), message('101', 'older', '100')] :
        [message('103', 'recent', badPage ? '200' : '100'), message('102', 'overlap', '100')] : [message('101', '<unsafe>')];
      const link = new URL(url); link.searchParams.set('startTime', '2'); if (pageHost) link.host = pageHost;
      return response({ messages: items, _metadata: { backwardLink: thread && !older ? link.href : null } });
    }
    throw new Error('Unexpected fixture endpoint');
  };
  const session = { credentials }, api = new TeamsAPI(session, { fetcher });
  return { api, requests, session, setExisting: value => existing = value, setChanged: value => changed = value,
    setBadPage: value => badPage = value, setPageHost: value => pageHost = value, setSearchShape: value => searchShape = value };
}

test('server message search deduplicates pages and opens exact account-bound chat/channel context', async () => {
  const f = fixture(), results = await f.api.search('match', 10, 3);
  assert.equal(results.items.length, 2); assert.equal(results.pages, 2); assert.equal(results.completeSearch, true);
  assert.equal(results.items[0].text, 'match');
  assert.deepEqual(f.requests.filter(r => r.url.pathname.endsWith('/query')).map(r => JSON.parse(r.body).EntityRequests[0].from), [0, 10]);
  const opened = await f.api.openSearchResult(results.items[0].resultId);
  assert.equal(opened.chatId, chatId); assert.equal(opened.selectedMessage.id, '101'); assert.equal(opened.mayMarkRead, false);
  const thread = await f.api.openSearchResult(results.items[1].resultId);
  assert.equal(thread.channelId, channelId); assert.equal(thread.parentMessageId, '100');
  f.session.credentials.identity = { ...f.session.credentials.identity, oid: peer };
  await assert.rejects(f.api.openSearchResult(results.items[0].resultId), { code: 'invalid_search_result' });
  assert.equal(f.requests.filter(r => r.method === 'POST' && r.url.pathname.endsWith('/messages')).length, 0);
});

test('search locators reject foreign hosts and path injection; malformed server data never becomes an empty success', async () => {
  const f = fixture();
  for (const WebUrl of ['https://evil.example/l/message/chat/1', 'https://teams.microsoft.com@evil.example/l/message/chat/1', 'https://teams.microsoft.com/l/message/19%3Ax%3Bmessageid%3D2/1']) assert.equal(f.api.searchLocator({ WebUrl }), null);
  const result = await f.api.search('match', 10, 1); assert.equal(result.truncated, true);
  f.api.searchResults.get(result.items[0].resultId).expires = 1;
  await assert.rejects(f.api.openSearchResult(result.items[0].resultId), { code: 'invalid_search_result' });
  f.setSearchShape(true); await assert.rejects(f.api.search('match'), { code: 'api_schema_changed' });
});

test('people suggestions deduplicate candidates; virtual one-to-one chats bind account and exact recipient across restart', async () => {
  const f = fixture(), people = await f.api.people('Peer');
  assert.equal(people.items.length, 1); assert.equal(people.items[0].email, 'peer@example.com');
  const virtual = await f.api.existingConversation('peer@example.com'); assert.equal(virtual.persisted, false);
  assert.equal((await f.api.resolve(virtual.chatId)).id, `19:${oid}_${peer}@unq.gbl.spaces`);
  const forged = JSON.parse(Buffer.from(virtual.chatId, 'base64url')); forged.threadId = '19:foreign@thread.v2';
  await assert.rejects(f.api.resolve(Buffer.from(JSON.stringify(forged)).toString('base64url')), { code: 'conversation_unavailable' });
  const restarted = new TeamsAPI(f.session, { fetcher: f.api.fetcher });
  assert.equal((await restarted.resolve(virtual.chatId)).members[1].mri, `8:orgid:${peer}`);
  f.setExisting(true); assert.equal((await f.api.existingConversation('peer@example.com')).chatId, chatId);
});

test('person mentions use re-resolved server identity and escaped markup; exact quotes refuse changed targets before dispatch', async () => {
  const f = fixture();
  const prepared = await f.api.prepareSend(chatId, { content: [{ type: 'paragraph', runs: [{ mention: { email: 'peer@example.com', name: 'Untrusted label' } }, { text: ' <script>' }] }], replyToMessageId: '101' });
  assert.ok(prepared.payload.content.includes('Peer &amp; Co'));
  assert.ok(!prepared.payload.content.includes('Untrusted label'));
  assert.ok(prepared.payload.content.includes('&lt;script&gt;'));
  assert.ok(prepared.payload.content.includes('itemtype="http://schema.skype.com/Reply" itemid="101"'));
  assert.deepEqual(JSON.parse(prepared.payload.properties.mentions)[0], { '@type': 'http://schema.skype.com/Mention', itemid: 0, mri: `8:orgid:${peer}`, mentionType: 'person', displayName: 'Peer & Co' });
  f.setChanged(true); await assert.rejects(f.api.send(prepared), { code: 'message_changed' });
  assert.equal(f.requests.filter(r => r.method === 'POST' && r.url.pathname.endsWith('/messages')).length, 0);
  assert.throws(() => messagePayload({ content: [{ type: 'paragraph', runs: [{ mention: { email: 'peer@example.com', name: 'Peer' } }] }] }, 'Self'), { code: 'mention_unresolved' });
});

test('quoted/first-chat sends replay after restart and bind the selected quote or exact recipient', async t => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-parity-ledger-')); t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const f = fixture(), ledger = new SendLedger(dir), input = { text: 'Reply', replyToMessageId: '101' };
  const first = await ledger.run('quote-key-0001', ['quote', chatId, input], () => f.api.prepareSend(chatId, input), p => f.api.send(p), 'send_uncertain');
  assert.equal(first.status, 'accepted_by_api');
  const restarted = new SendLedger(dir);
  assert.equal((await restarted.run('quote-key-0001', ['quote', chatId, input], () => assert.fail('No preparation on replay'), () => assert.fail('No resend'), 'send_uncertain')).replayed, true);
  await assert.rejects(restarted.run('quote-key-0001', ['quote', chatId, { ...input, replyToMessageId: '100' }], () => {}, () => {}, 'send_uncertain'), { code: 'idempotency_conflict' });
  const virtual = await f.api.existingConversation('peer@example.com');
  assert.equal((await f.api.send(await f.api.prepareSend(virtual.chatId, { text: 'First message' }))).status, 'accepted_by_api');
});

test('per-thread history includes an older root once, deduplicates overlap, exports NDJSON and scopes cursor replay', async () => {
  const f = fixture(), args = { kind: 'channel', parentMessageId: '100', limit: 2, format: 'ndjson' };
  const first = await f.api.history(channelId, args);
  assert.deepEqual(first.items.map(m => m.id), ['100', '102', '103']); assert.equal(first.hasMore, true);
  const next = await f.api.history(channelId, { ...args, cursor: first.nextCursor });
  assert.deepEqual(next.items.map(m => m.id), ['101']); assert.equal(JSON.parse(next.ndjson).id, '101');
  assert.deepEqual(await f.api.history(channelId, { ...args, cursor: first.nextCursor }), next);
  await assert.rejects(f.api.history(channelId, { ...args, parentMessageId: '200', cursor: first.nextCursor }), { code: 'invalid_cursor' });
  await assert.rejects(f.api.history(channelId, { ...args, rootsOnly: true, cursor: first.nextCursor }), { code: 'invalid_cursor' });
  f.session.credentials.identity = { ...f.session.credentials.identity, oid: peer };
  await assert.rejects(f.api.history(channelId, { ...args, cursor: first.nextCursor }), { code: 'invalid_cursor' });
});

test('thread reads reject reply-as-root, sibling messages, composite injection and redirected history links', async () => {
  const f = fixture();
  await assert.rejects(f.api.threadMessages(channelId, '199'), { code: 'invalid_thread_root' });
  await assert.rejects(f.api.threadMessages(channelId, '100;messageid=200'), { code: 'invalid_message_id' });
  f.setBadPage(true); await assert.rejects(f.api.threadMessages(channelId, '100'), { code: 'api_schema_changed' });
  f.setBadPage(false); f.setPageHost('evil.example');
  const first = await f.api.history(channelId, { kind: 'channel', parentMessageId: '100' });
  await assert.rejects(f.api.history(channelId, { kind: 'channel', parentMessageId: '100', cursor: first.nextCursor }), { code: 'api_cursor_rejected' });
});

test('optional search token acquisition uses HTTP, preserves base credentials and returns rotations for encrypted persistence', async () => {
  let count = 0;
  const credentials = { ...signedBundle(), profile: 'test', authMode: 'tokens', renewal: renewal() };
  const session = new TeamsSession('test', '/unused', {}, { createAuthenticator: () => assert.fail('No browser'), fetcher: async (_url, options) => {
    count++; const form = new URLSearchParams(options.body);
    assert.equal(form.get('scope'), 'https://substrate.office.com/.default openid profile offline_access');
    assert.equal(form.get('refresh_token'), credentials.renewal.value);
    return new Response(JSON.stringify({ access_token: signedToken('skype', { aud: 'https://substrate.office.com' }), token_type: 'Bearer', expires_in: 3600, refresh_token: 'rotated-search' }));
  } });
  session.adopt(credentials); await session.ensureSearch(); await session.ensureSearch();
  assert.equal(count, 1); assert.equal(session.credentials.tokens.skype.value, credentials.tokens.skype.value);
  assert.equal(session.credentials.renewal.value, 'rotated-search');
  const encoded = encodeTokenBundle(session.credentials), restored = await validateTokenBundle(encoded);
  assert.equal(restored.tokens.substrate.value, session.credentials.tokens.substrate.value);
});

test('schema restores legacy search/thread operation IDs and structured mentions/exact quote input', () => {
  for (const route of ['/search/messages', '/search/messages/open', '/channels/{channelId}/threads', '/channels/{channelId}/threads/{parentMessageId}/messages', '/channels/{channelId}/threads/{parentMessageId}/history']) assert.ok(openapi.paths[route]);
  const schema = openapi.paths['/chats/{chatId}/messages'].post.requestBody.content['application/json'].schema;
  assert.ok(schema.properties.replyToMessageId);
  assert.ok(schema.properties.content.items.properties.runs.items.properties.mention);
  assert.ok(!openapi.paths['/channels/{channelId}/threads'].post);
});


test('captured structured search locators resolve exact IDs and reject inconsistent conversation/root fields', () => {
  const f = fixture();
  const source = { ClientConversationId: chat, ClientThreadId: chat, InternetMessageId: '101' };
  assert.deepEqual(f.api.searchLocator(source), { threadId: chat, kind: 'chat', messageId: '101', parentMessageId: null });
  assert.deepEqual(f.api.searchLocator({ ...source, ClientConversationId: channel, ClientThreadId: `${channel};messageid=100` }), { threadId: channel, kind: 'channel', messageId: '101', parentMessageId: '100' });
  for (const invalid of [{ ClientThreadId: channel }, { InternetMessageId: '101/other' }, { ClientConversationId: `${chat};messageid=100`, ClientThreadId: `${chat};messageid=100` }]) assert.equal(f.api.searchLocator({ ...source, ...invalid }), null);
});

test('structured edits and triage preparation share exact-email mention metadata', async () => {
  const f = fixture();
  f.session.credentials.identity = { ...f.session.credentials.identity, oid: peer };
  const input = { content: [{ type: 'paragraph', runs: [{ mention: { email: 'peer@example.com', name: 'Ignored label' } }] }] };
  const edit = await f.api.prepareMutation(chatId, '101', 'edit', { ...input, expectedText: '<unsafe>' });
  assert.equal(edit.payload.skypeeditedid, '101');
  assert.equal(JSON.parse(edit.payload.properties.mentions)[0].mri, `8:orgid:${peer}`);
  const triage = await f.api.prepareSendForTriage({ threadId: chat }, input);
  assert.equal(triage.payload.content, edit.payload.content);
  assert.equal(triage.payload.properties.mentions, edit.payload.properties.mentions);
});

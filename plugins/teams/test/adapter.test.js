import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { EventEmitter } from 'node:events';
import { Readable } from 'node:stream';
import { createAdapter } from '../legacy-index.js';

async function fixture(t) {
  const dataDir = await fs.mkdtemp('/tmp/teams-adapter-');
  const profiles = new Map();
  let handler;
  const instance = await createAdapter({ dataDir, settings: {} }, {
    createBrowser: dir => {
      const browser = {
        serial: fn => fn(), login: async () => ({ opened: true }), dom: async action => ({ ready: true, action }),
        prepareMessageAction: async (id, mid, kind, payload) => { browser.mutationPrepared = (browser.mutationPrepared ?? 0) + 1; return { args: { kind, id, mid, ...payload } }; },
        prepareUnread: async id => ({ args: { kind: 'unread', id }, cleanup: async () => { browser.cleaned = true; } }),
        executeMessageAction: async prepared => { browser.mutated = (browser.mutated ?? 0) + 1; return { status: prepared.args.kind + '_observed', ...prepared.args }; },
        reactions: async (_id, mid) => ({ messageId: mid, available: [{ reaction: 'like' }] }),
        scanChats: async (unreadOnly, maxWindows) => ({ unreadOnly, maxWindows }),
        prepare: async id => ({ chat: { title: id } }),
        prepareQuote: async (id, replyToMessageId) => ({ chat: { title: id }, replyToMessageId }),
        searchMessages: async (query, limit, maxPages) => ({ query, limit, maxPages }),
        listChats: async unreadOnly => ({ profile: dir, unreadOnly }),
        channels: async maxWindows => ({ maxWindows, items: [] }),
        channelThreads: async id => ({ channelId: id }),
        threadMessages: async (id, parent) => ({ channelId: id, parentMessageId: parent }),
        history: async (id, options) => ({ conversationId: id, ...options }),
        prepareContent: async (id, content, parent) => { browser.richPrepared = (browser.richPrepared ?? 0) + 1; return { args: { id, parent }, content }; },
        sendContent: async prepared => { browser.richSent = (browser.richSent ?? 0) + 1; return { status: 'observed_in_ui', ...prepared.args }; },
        unreadMessages: async (maxChats, limitPerChat) => ({ maxChats, limitPerChat, mayMarkRead: true }),
        searchPeople: async query => ({ query, items: [{ email: 'alice@example.com' }] }),
        conversation: async email => ({ email, messageSent: false }),
        prepareRecipient: async email => { browser.preparedCount = (browser.preparedCount ?? 0) + 1; return { chat: { title: 'Alice', key: { attribute: 'recipient', value: email } } }; },
        send: async () => { browser.sentCount = (browser.sentCount ?? 0) + 1; return { status: 'observed_in_chat', message: { id: 'sent1' } }; },
        chatId: () => 'person-chat-id', close: async () => {},
      };
      profiles.set(dir, browser); return browser;
    },
    createServer: callback => {
      handler = callback;
      const server = new EventEmitter();
      server.listen = (_port, host, ready) => { assert.equal(host, '127.0.0.1'); ready(); };
      server.address = () => ({ port: 12345 });
      server.close = done => done();
      server.closeIdleConnections = () => {};
      return server;
    },
  });
  t.after(async () => { await instance.dispose(); await fs.rm(dataDir, { recursive: true, force: true }); });
  const service = instance.services[0], auth = service.authMethods[0];
  async function connection() {
    const flow = await auth.connect({ config: { label: 'test' } });
    return auth.poll({ pending: flow.pending });
  }
  function request(conn, pathname, method = 'GET') {
    const outgoing = { url: new URL(pathname, service.baseUrl), method, headers: new Headers() };
    auth.authorize(outgoing, conn); return outgoing;
  }
  async function invoke(out, { pathname, method, body = '', destroyed = false } = {}) {
    const request = Readable.from(body ? [Buffer.from(body)] : []);
    request.headers = Object.fromEntries(out.headers);
    request.url = pathname ?? out.url.pathname + out.url.search;
    request.method = method ?? out.method;
    const result = {};
    const response = { destroyed, writeHead: status => result.status = status, end: text => result.body = JSON.parse(text) };
    await handler(request, response); return result;
  }
  return { service, auth, connection, request, invoke, profiles };
}

test('single-use invocation tokens bind account, method, path and query', async t => {
  const f = await fixture(t), conn = await f.connection();
  const out = f.request(conn, '/status');
  assert.equal((await f.invoke(out)).status, 200);
  assert.equal((await f.invoke(out)).status, 401);
  assert.equal((await f.invoke({ ...out, headers: new Headers() })).status, 401);
  assert.equal((await f.invoke(f.request(conn, '/status'), { pathname: '/chats' })).status, 403);
  assert.equal((await f.invoke(f.request(conn, '/status?a=1'), { pathname: '/status?a=2' })).status, 403);
  assert.equal((await f.invoke(f.request(conn, '/status'), { method: 'POST' })).status, 403);
});

test('channel threads, history and structured sends validate inputs and preserve content-bound retries', async t => {
  const f = await fixture(t), conn = await f.connection();
  assert.equal((await f.invoke(f.request(conn, '/channels?maxWindows=7'))).body.maxWindows, 7);
  assert.equal((await f.invoke(f.request(conn, '/channels/channel-id/threads/root/messages'))).body.parentMessageId, 'root');
  const page = await f.invoke(f.request(conn, '/chats/chat-id/history?limit=17&maxWindows=8&format=ndjson'));
  assert.equal(page.body.limit, 17); assert.equal(page.body.format, 'ndjson');
  assert.equal((await f.invoke(f.request(conn, '/chats/chat-id/history?format=html'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/chats/chat-id/history?cursor=forged'))).status, 400);
  const content = [{ type: 'paragraph', runs: [{ text: 'Hi ', marks: ['bold'] }, { mention: { email: 'alice@example.com', name: 'Alice' } }] }];
  const payload = { content, idempotencyKey: 'rich-channel-0001' };
  const send = data => f.invoke(f.request(conn, '/channels/channel-id/threads/root/messages', 'POST'), { body: JSON.stringify(data) });
  assert.equal((await send(payload)).status, 200); assert.equal((await send(payload)).body.replayed, true);
  assert.equal((await send({ ...payload, content: [{ type: 'paragraph', runs: [{ text: 'Changed' }] }] })).status, 409);
  assert.equal((await send({ ...payload, text: 'Also text' })).status, 400);
  const browser = [...f.profiles.values()][0]; assert.equal(browser.richPrepared, 1); assert.equal(browser.richSent, 1);
  const chatSend = data => f.invoke(f.request(conn, '/chats/chat-id/messages', 'POST'), { body: JSON.stringify(data) });
  assert.equal((await chatSend({ content, idempotencyKey: 'rich-chat-0001' })).status, 200);
  assert.equal((await chatSend({ content, replyToMessageId: 'root', idempotencyKey: 'rich-chat-0002' })).status, 400);
  const edit = data => f.invoke(f.request(conn, '/chats/chat-id/messages/own-message', 'PATCH'), { body: JSON.stringify(data) });
  assert.equal((await edit({ content, expectedText: 'Original', idempotencyKey: 'rich-edit-0001' })).status, 200);
  assert.equal((await edit({ content, text: 'Also text', expectedText: 'Original', idempotencyKey: 'rich-edit-0002' })).status, 400);
});

test('connections have separate profiles and malformed requests fail before browser', async t => {
  const f = await fixture(t), a = await f.connection(), b = await f.connection();
  assert.notEqual(a.credentials.profile, b.credentials.profile);
  assert.equal(f.profiles.size, 2);
  assert.equal(f.service.baseUrl, 'http://teams.localhost');
  const result = await f.invoke(f.request(a, '/chats/invalid/messages?limit=-1'));
  assert.equal(result.status, 400);
  assert.equal(result.body.error.code, 'invalid_query');
  const wrongBody = await f.invoke(f.request(a, '/chats/invalid/messages', 'POST'), { body: '{"text":"hello","extra":true}' });
  assert.equal(wrongBody.status, 400);
  assert.equal(wrongBody.body.error.code, 'invalid_body');
  assert.equal((await f.invoke(f.request(a, '/status'), { destroyed: true })).status, 499);
});

test('revocation invalidates pending invocations', async t => {
  const f = await fixture(t), conn = await f.connection();
  const out = f.request(conn, '/status');
  await f.auth.revoke(conn);
  assert.equal((await f.invoke(out)).status, 401);
});

test('unread and people operations validate bounds and route the requested filters', async t => {
  const f = await fixture(t), conn = await f.connection();
  assert.equal((await f.invoke(f.request(conn, '/unread/chats'))).body.unreadOnly, true);
  assert.equal((await f.invoke(f.request(conn, '/chats?unreadOnly=false'))).body.unreadOnly, false);
  assert.equal((await f.invoke(f.request(conn, '/chats?unreadOnly=maybe'))).status, 400);
  assert.deepEqual((await f.invoke(f.request(conn, '/unread/messages?maxChats=2&limitPerChat=10'))).body, { maxChats: 2, limitPerChat: 10, mayMarkRead: true });
  assert.equal((await f.invoke(f.request(conn, '/unread/messages?maxChats=6'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/people'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/people?query=Alice'))).body.query, 'Alice');
});

test('opening a conversation does not send; initial messages have durable recipient-bound retries', async t => {
  const f = await fixture(t), conn = await f.connection();
  const browser = [...f.profiles.values()][0];
  const post = input => f.invoke(f.request(conn, '/conversations', 'POST'), { body: JSON.stringify(input) });
  assert.deepEqual((await post({ email: 'Alice@Example.COM' })).body, { email: 'alice@example.com', messageSent: false });
  assert.equal(browser.sentCount, undefined);
  assert.equal((await post({ email: 'Alice Adams', text: 'Hi', idempotencyKey: 'person-send-1' })).status, 400);
  assert.equal((await post({ email: 'alice@example.com', text: 'Hi' })).status, 400);
  const first = await post({ email: 'alice@example.com', text: 'Hi', idempotencyKey: 'person-send-1' });
  const retry = await post({ email: 'ALICE@EXAMPLE.COM', text: 'Hi', idempotencyKey: 'person-send-1' });
  assert.equal(first.body.chatId, 'person-chat-id');
  assert.equal(retry.body.replayed, true);
  assert.equal(browser.sentCount, 1);
  assert.equal(browser.preparedCount, 1);
  assert.equal((await post({ email: 'bob@example.com', text: 'Hi', idempotencyKey: 'person-send-1' })).body.error.code, 'idempotency_conflict');
});

test('search and quote requests validate inputs; retries bind the exact quoted message', async t => {
  const f = await fixture(t), conn = await f.connection();
  assert.equal((await f.invoke(f.request(conn, '/unread/chats?maxWindows=1'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/search/messages'))).status, 400);
  assert.equal((await f.invoke(f.request(conn, '/search/messages?query=hi&maxPages=11'))).status, 400);
  assert.deepEqual((await f.invoke(f.request(conn, '/search/messages?query=hi&limit=20&maxPages=2'))).body, { query: 'hi', limit: 20, maxPages: 2 });
  const post = input => f.invoke(f.request(conn, '/chats/chat-id/messages', 'POST'), { body: JSON.stringify(input) });
  const input = { text: 'Reply', idempotencyKey: 'quoted-message-1', replyToMessageId: '100' };
  assert.equal((await post({ ...input, replyToMessageId: '' })).status, 400);
  assert.equal((await post(input)).body.status, 'observed_in_chat');
  assert.equal((await post(input)).body.replayed, true);
  assert.equal((await post({ ...input, replyToMessageId: '101' })).body.error.code, 'idempotency_conflict');
  assert.equal((await post({ text: 'Reply', idempotencyKey: 'quoted-message-1' })).body.error.code, 'idempotency_conflict');
  assert.equal([...f.profiles.values()][0].sentCount, 1);
});


test('message changes require explicit state and expected content; replay never reopens menus', async t => {
  const f = await fixture(t), conn = await f.connection();
  const route = (path, method, value) => f.invoke(f.request(conn, path, method), { body: JSON.stringify(value) });
  const base = '/chats/chat-id/messages/100';
  assert.equal((await f.invoke(f.request(conn, base + '/reactions'))).body.available[0].reaction, 'like');
  assert.equal((await route(base + '/reactions', 'POST', { reaction: 'like', idempotencyKey: 'reaction-001' })).status, 400);
  assert.equal((await route(base, 'PATCH', { text: 'New', idempotencyKey: 'editing-001' })).status, 400);
  assert.equal((await route(base, 'DELETE', { idempotencyKey: 'deleting-001' })).status, 400);
  const input = { reaction: 'yes', selected: true, idempotencyKey: 'reaction-001' };
  assert.equal((await route(base + '/reactions', 'POST', input)).body.reaction, 'like');
  assert.equal((await route(base + '/reactions', 'POST', { ...input, reaction: 'like' })).body.replayed, true);
  assert.equal((await route(base + '/reactions', 'POST', { ...input, selected: false })).body.error.code, 'idempotency_conflict');
  assert.equal((await route(base, 'PATCH', { text: 'New', expectedText: 'Old', idempotencyKey: 'editing-001' })).body.status, 'edit_observed');
  assert.equal((await route(base, 'DELETE', { expectedText: 'New', idempotencyKey: 'deleting-001' })).body.status, 'delete_observed');
  assert.equal((await route('/chats/chat-id/unread', 'POST', { idempotencyKey: 'unread-001' })).body.status, 'unread_observed');
  assert.equal((await route('/chats/chat-id/unread', 'POST', { idempotencyKey: 'unread-001' })).body.replayed, true);
  assert.equal((await route('/chats/chat-id/unread', 'POST', { selected: false, idempotencyKey: 'unread-002' })).status, 400);
  const browser = [...f.profiles.values()][0];
  assert.equal(browser.mutated, 4); assert.equal(browser.mutationPrepared, 3); assert.equal(browser.cleaned, true);
});

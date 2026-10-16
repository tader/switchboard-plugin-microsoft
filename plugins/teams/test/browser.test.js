import test from 'node:test';
import assert from 'node:assert/strict';
import { TeamsBrowser, profilePath } from '../lib/browser.js';
import { selectorsFrom } from '../lib/dom.js';

test('profile queue serializes operations after failures and isolates profiles', async () => {
  const a = new TeamsBrowser('/tmp/a'), b = new TeamsBrowser('/tmp/b');
  const events = [];
  let unlock;
  const barrier = new Promise(resolve => unlock = resolve);
  const first = a.serial(async () => { events.push('a1'); await barrier; throw new Error('failure'); });
  const second = a.serial(async () => events.push('a2'));
  await b.serial(async () => events.push('b1'));
  assert.deepEqual(events, ['a1', 'b1']);
  unlock(); await assert.rejects(first); await second;
  assert.deepEqual(events, ['a1', 'b1', 'a2']);
});

test('profile IDs cannot traverse filesystem and chat IDs are validated', () => {
  assert.throws(() => profilePath('/tmp', '../escape'), { code: 'invalid_profile' });
  const browser = new TeamsBrowser('/tmp/a');
  for (const id of ['bad', '', Buffer.from(JSON.stringify({key:{attribute:'onclick', value:'x'},title:'x'})).toString('base64url')]) assert.throws(() => browser.decodeChat(id), { code: 'invalid_chat' });
  const chat = { key: { attribute: 'data-chat-id', value: 'a"b' }, title: 'A chat' };
  assert.deepEqual(browser.decodeChat(Buffer.from(JSON.stringify(chat)).toString('base64url')), chat);
});

test('selector overrides reject unsupported fields and empty selectors', () => {
  assert.equal(selectorsFrom('{"send":"button.send"}').send, 'button.send');
  for (const value of ['[]', '{"cookie":"x"}', '{"send":""}', '{"send":true}']) assert.throws(() => selectorsFrom(value));
});

test('unread reads snapshot flags before opening chats, cap work and separate fallback messages', async () => {
  const browser = new TeamsBrowser('/tmp/unread');
  const events = [];
  browser.scanChats = async unreadOnly => {
    assert.equal(unreadOnly, true); events.push('snapshot');
    return { items: [{ id: 'a', title: 'Alice' }, { id: 'b', title: 'Bob' }, { id: 'c', title: 'Carol' }] };
  };
  browser.openChat = async id => { events.push('open-' + id); return { title: id }; };
  browser.dom = async (_action, chat) => chat.title === 'a' ? { items: [{ id: 'm1', text: 'Unread' }], boundaryFound: true } : { items: [], recentMessages: [{ id: 'old', text: 'Unconfirmed' }], boundaryFound: false };
  const result = await browser.unreadMessages(2, 10);
  assert.deepEqual(events, ['snapshot', 'open-a', 'open-b']);
  assert.equal(result.items.length, 1);
  assert.equal(result.chats[1].recentMessages[0].id, 'old');
  assert.equal(result.remainingChats, 1);
  assert.equal(result.mayMarkRead, true);
  assert.equal(result.completeAccount, false);
});

test('person conversations reselect by normalized email and never trust an encoded display name', async () => {
  const browser = new TeamsBrowser('/tmp/person');
  const person = { key: { attribute: 'recipient', value: 'alice@example.com' }, title: 'Verified Alice', threadId: null };
  browser.openRecipient = async email => { assert.equal(email, 'alice@example.com'); return person; };
  const forged = { ...person, title: 'Someone else', threadId: 'wrong-thread' };
  assert.deepEqual(await browser.openChat(browser.chatId(forged)), person);
  const opened = await browser.conversation('Alice@Example.COM');
  assert.equal(opened.messageSent, false);
  assert.equal(opened.email, 'alice@example.com');
  assert.equal(browser.decodeChat(opened.chatId).title, 'Verified Alice');
});

test('unread discovery walks virtualized windows, deduplicates threads, requires a stable end and restores UI', async () => {
  const browser = new TeamsBrowser('/tmp/discover'), events = [];
  const a = { key: { attribute: 'data-chat-id', value: 'a' }, threadId: 'a', title: 'A', unread: true };
  const b = { ...a, key: { attribute: 'data-chat-id', value: 'b' }, threadId: 'b', title: 'B' };
  let next = 0;
  browser.dom = async (action, args) => {
    events.push(action); assert.ok(args.token);
    if (action === 'scanWindow') return { items: [a], atEnd: false, position: 0, height: 300 };
    if (action === 'scanNext') { next++; return { items: [a, b], atEnd: true, position: 200, height: 300 }; }
    return {};
  };
  const result = await browser.scanChats(true, 10);
  assert.equal(result.items.length, 2); assert.equal(result.discoveryComplete, true); assert.equal(next, 2);
  assert.equal(result.completeAccount, false); assert.equal(events.at(-1), 'scanEnd');
  next = 0; const partial = await browser.scanChats(true, 2);
  assert.equal(partial.discoveryComplete, false); assert.equal(partial.truncated, true);
  browser.dom = async action => { events.push(action); if (action === 'scanWindow') throw new Error('layout changed'); return {}; };
  await assert.rejects(browser.scanChats(true), /layout changed/); assert.equal(events.at(-1), 'scanEnd');
});

test('native message search deduplicates paginated results and caps work', async () => {
  const browser = new TeamsBrowser('/tmp/search'), calls = []; let page = 0;
  browser.page = { call: async (method, params) => calls.push([method, params]) };
  browser.dom = async (action, args) => {
    assert.ok(action === 'searchFocus' || args.query === 'find this');
    if (action === 'searchFocus') return { focused: true };
    if (action === 'searchNext') { page++; return { advanced: true }; }
    if (action === 'searchResults') return { items: [{ key: 'serp-message-card-content-' + page, text: 'match' }, { key: 'serp-message-card-content-' + (page + 1), text: 'match' }], hasNext: page < 2 };
    return {};
  };
  const result = await browser.searchMessages('find this', 100, 3);
  assert.equal(result.items.length, 4); assert.equal(result.completeSearch, true); assert.equal(result.pagesScanned, 3);
  assert.equal(calls[1][1].text, '\r');
  const decoded = browser.decodeSearch(result.items[0].resultId); assert.equal(decoded.query, 'find this');
  assert.throws(() => browser.decodeSearch('invalid'), { code: 'invalid_search_result' });
});

test('quote preparation finds an exact older ID after reopening resets the message window', async () => {
  const browser = new TeamsBrowser('/tmp/older-quote');
  const chat = { title: 'Alice', threadId: 'thread-a' }; let history = 0;
  browser.openChat = async () => chat;
  browser.page = { call: async (_method, args) => assert.equal(args.button, 'right') };
  browser.dom = async (action, args) => {
    if (action === 'quoteTarget') { if (history < 2) { const error = new Error('not rendered'); error.code = 'message_unavailable'; throw error; } return { x: 1, y: 2 }; }
    if (action === 'older') { history++; return { changed: true }; }
    assert.equal(action, 'quoteSelect'); assert.equal(args.replyToMessageId, 'older-id');
    return { replyToMessageId: args.replyToMessageId, beforeIds: ['older-id'] };
  };
  assert.equal((await browser.prepareQuote('chat-id', 'older-id')).replyToMessageId, 'older-id');
  assert.equal(history, 2);
});


test('marking a rendered chat unread does not open it; edit input uses only the prepared inline editor', async () => {
  const browser = new TeamsBrowser('/tmp/actions');
  const chat = { key: { attribute: 'data-chat-id', value: 'thread-a' }, title: 'Alice', threadId: 'thread-a' };
  browser.openChat = () => assert.fail('Unread marking should not open the conversation');
  const calls = [];
  browser.dom = async (action, args) => { calls.push(action); if (action === 'action:unreadTarget') return { x: 1, y: 2, alreadyUnread: false }; return {}; };
  browser.page = { call: async (method, params) => { calls.push(method); if (method === 'Input.dispatchMouseEvent') assert.equal(params.button, 'right'); } };
  const prepared = await browser.prepareUnread(browser.chatId(chat));
  await browser.executeMessageAction(prepared); await prepared.cleanup();
  assert.deepEqual(calls, ['status', 'action:unreadTarget', 'Input.dispatchMouseEvent', 'Input.dispatchMouseEvent', 'action:unreadPrepare', 'action:unreadCommit', 'action:release']);
  calls.length = 0;
  await browser.executeMessageAction({ args: { kind: 'edit', text: 'Changed' } });
  assert.deepEqual(calls, ['action:editSelect', 'Input.insertText', 'action:editCommit']);
});

test('channel search IDs stay bound to their verified parent and cannot silently become a top-level post', async () => {
  const browser = new TeamsBrowser('/tmp/channel-search');
  const context = { kind: 'channel', title: 'Example', threadId: 'channel@thread.tacv2', parentMessageId: 'root', key: { attribute: 'search', value: JSON.stringify({ query: 'find', resultKey: 'serp-message-card-content-fixture' }) } };
  const id = browser.chatId(context); let opened = 0;
  browser.openSearch = async () => { opened++; return context; };
  assert.deepEqual(await browser.openChannel(id, 'root'), context);
  await assert.rejects(browser.openChannel(id, null), { code: 'invalid_channel' });
  await assert.rejects(browser.openChannel(id, 'different-root'), { code: 'invalid_channel' });
  assert.equal(opened, 1);
  browser.openSearch = async () => ({ ...context, parentMessageId: 'different-root' });
  await assert.rejects(browser.openChannel(id, 'root'), { code: 'channel_changed' });
});

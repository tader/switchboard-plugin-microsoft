import test from 'node:test';
import assert from 'node:assert/strict';
import { TeamsBrowser } from '../lib/browser.js';

test('history cursors traverse beyond five windows, deduplicate overlaps, replay pages and stop only at a stable UI start', async () => {
  const browser = new TeamsBrowser('/tmp/history-test');
  const chat = { title: 'Alice', threadId: 'thread-a', key: { attribute: 'data-chat-id', value: 'a' } };
  const id = browser.chatId(chat); let end = 30, opens = 0, scrolls = 0;
  const items = () => Array.from({ length: Math.min(10, end) }, (_, i) => ({ id: String(Math.max(1, end - 9) + i), text: 'Message' }));
  browser.openChat = async () => { opens++; end = 30; return chat; };
  browser.dom = async action => {
    if (action === 'extended:messages') return { items: items() };
    assert.equal(action, 'extended:older'); scrolls++; const before = end; end = Math.max(1, end - 3);
    return { items: items(), changed: before !== end, atStart: end === 1, loading: false };
  };
  const seen = new Set(); let cursor, previous;
  for (let i = 0; i < 30; i++) {
    const page = await browser.history(id, { cursor, limit: 4, maxWindows: 1, format: 'ndjson' });
    if (cursor) assert.deepEqual(await browser.history(id, { cursor, limit: 4, maxWindows: 1, format: 'ndjson' }), page);
    for (const message of page.items) { assert.ok(!seen.has(message.id)); seen.add(message.id); }
    assert.equal(page.completeHistory, false); assert.equal(page.ndjson.trim().split('\n').filter(Boolean).length, page.items.length);
    previous = page; cursor = page.nextCursor; if (!cursor) break;
  }
  assert.equal(seen.size, 30); assert.ok(scrolls > 5); assert.equal(opens, 1); assert.equal(previous.uiStartReached, true);
  assert.equal(previous.hasMore, false);
});

test('history cursors bind parameters and profile, and seek an anchor after another operation changes views', async () => {
  const browser = new TeamsBrowser('/tmp/history-resume');
  const chat = { title: 'Alice', threadId: 'thread-a', key: { attribute: 'data-chat-id', value: 'a' } }, id = browser.chatId(chat);
  let changed = false, step = 0;
  browser.openChat = async () => { changed = false; step = 0; return chat; };
  browser.dom = async action => {
    if (changed) { const e = new Error('Changed'); e.code = 'chat_changed'; throw e; }
    if (action === 'extended:messages') return { items: [{ id: '3' }, { id: '4' }, { id: '5' }] };
    step++; return { items: step === 1 ? [{ id: '2' }, { id: '3' }, { id: '4' }] : [{ id: '1' }, { id: '2' }], changed: step <= 2, atStart: step >= 2, loading: false };
  };
  const first = await browser.history(id, { limit: 1, maxWindows: 1 });
  await assert.rejects(browser.history('other', { cursor: first.nextCursor, limit: 1 }), { code: 'invalid_history_cursor' });
  await assert.rejects(browser.history(id, { cursor: first.nextCursor, limit: 2 }), { code: 'invalid_history_cursor' });
  const otherProfile = new TeamsBrowser('/tmp/history-other');
  await assert.rejects(otherProfile.history(id, { cursor: first.nextCursor, limit: 1 }), { code: 'invalid_history_cursor' });
  changed = true;
  const resumed = await browser.history(id, { cursor: first.nextCursor, limit: 1, maxWindows: 1 });
  assert.deepEqual(resumed.items, [{ id: '4' }]); assert.equal(resumed.hasMore, true);
});

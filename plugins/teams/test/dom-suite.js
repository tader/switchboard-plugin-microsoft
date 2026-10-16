// This fixture models the inspected 2026 combined-sidebar Teams DOM. It is a
// separate blank page, never the user's Teams session. No Microsoft traffic.
export const fixtureHTML = `<!doctype html><html><body>
<div data-tid="simple-collab-dnd-rail" role="tree">
 <div role="treeitem" data-item-type="chat" data-fui-tree-item-value="Folder|folder/OneGQL_ChatConversation|thread-a" onclick="document.querySelector('[data-tid=chat-title]').textContent='Alice';document.querySelector('button').setAttribute('data-track-thread-id','thread-a')"><span id="title-chat-list-item_a">Alice</span></div>
 <div role="treeitem" data-item-type="chat" data-fui-tree-item-value="Folder|folder/OneGQL_ChatConversation|thread-b" onclick="document.querySelector('[data-tid=chat-title]').textContent='Bob';document.querySelector('button').setAttribute('data-track-thread-id','thread-b')"><span id="title-chat-list-item_b">Bob</span></div>
 <div role="treeitem" data-item-type="channel"><span>A channel, excluded</span></div>
</div>
<h2 data-tid="chat-title">Alice</h2>
<div data-tid="message-pane-list-viewport" style="height:150px;overflow:auto">
 <div data-tid="chat-pane-item"><span id="author-100">Alice</span><time id="timestamp-100" datetime="2026-10-07T10:00:00Z">10:00</time><div data-tid="chat-pane-message" data-mid="100"><div data-message-content>Hello Thomas</div></div></div>
</div>
<div id="fixture-composer" data-tid="ckeditor" role="textbox" contenteditable="true" style="border:1px solid;padding:10px"></div>
<button data-tid="sendMessageCommands-send" data-track-thread-id="thread-a" onclick="const editor=document.querySelector('[contenteditable]');const item=document.createElement('div');item.dataset.tid='chat-pane-item';const message=document.createElement('div');message.dataset.tid='chat-pane-message';message.dataset.mid=String(++window.messageCount);const content=document.createElement('div');content.setAttribute('data-message-content','');content.textContent=editor.innerText;message.append(content);item.append(message);document.querySelector('[data-tid=message-pane-list-viewport]').append(item);editor.textContent='';window.sentCount++">Send</button>
<script>window.messageCount=100;window.sentCount=0;</script>
</body></html>`;

export async function runDOMSuite(page, dom, selectors, html) {
  const results = [];
  const check = (condition, label) => { if (!condition) throw new Error(label); results.push(label); };
  await page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  const run = (action, args = {}, custom = selectors) => page.evaluate(dom, action, args, custom);
  check((await run('status')).ready, 'detects authenticated layout');
  const chats = await run('chats');
  check(chats.items.length === 2, 'lists chats and excludes channels');
  const alice = chats.items[0], bob = chats.items[1];
  check((await run('open', alice)).title === 'Alice', 'opens and verifies matching thread');
  const messages = await run('messages', { ...alice, limit: 50 });
  check(messages.items[0].id === '100' && messages.items[0].author === 'Alice' && messages.items[0].timestamp === '2026-10-07T10:00:00Z', 'extracts message ID, author and timestamp');
  check((await run('messages', { ...bob, limit: 50 })).error.code === 'chat_changed', 'refuses wrong conversation');
  check((await run('open', { ...alice, threadId: null })).error.code === 'invalid_chat', 'refuses forged thread identity');
  await page.evaluate(() => document.querySelector('[contenteditable]').textContent = 'Manual draft');
  check((await run('prepare', alice)).error.code === 'draft_exists', 'preserves existing drafts');
  await page.evaluate(() => document.querySelector('[contenteditable]').textContent = '');
  const prepared = await run('prepare', alice);
  check(prepared.beforeIds[0] === '100', 'records existing IDs before sending');
  const payload = 'Plain text <script> is not HTML\nSecond line';
  await page.call('Input.insertText', { text: payload });
  check((await run('sendPoint', { ...alice, text: 'wrong' })).error.code === 'draft_mismatch', 'refuses unexpected editor contents');
  const point = await run('sendPoint', { ...alice, text: payload });
  await page.clickPoint(point);
  const confirmed = await run('confirm', { ...alice, text: payload, beforeIds: prepared.beforeIds });
  check(confirmed.status === 'observed_in_chat' && confirmed.message.id === '101', 'observes a new matching message after a trusted send click');
  check(await page.evaluate(() => window.sentCount === 1), 'clicks Send exactly once');
  await run('prepare', alice);
  await page.call('Input.insertText', { text: 'Guarded send' });
  const guardedPoint = await run('sendPoint', { ...alice, text: 'Guarded send' });
  await page.evaluate(() => document.querySelector('button').setAttribute('data-track-thread-id', 'thread-b'));
  await page.clickPoint(guardedPoint);
  check(await page.evaluate(() => window.sentCount === 1), 'blocks send if the thread changes between verification and click');
  const bad = await run('diagnostics', {}, { ...selectors, send: '???' });
  check(bad.error.code === 'selector_mismatch', 'fails closed on broken selectors');
  const diagnostic = await run('diagnostics');
  check(!JSON.stringify(diagnostic).includes(payload), 'diagnostics omit message content');
  return results;
}

function initializeExtendedFixture() {
  const main = document.querySelector('[data-tid="app-layout-area--main"]');
  window.extCounts = { sent: 0, edited: 0 };
  const message = (id, text) => `<div data-tid="chat-pane-item"><div data-tid="chat-pane-message" data-mid="${id}"><div data-message-content>${text}</div></div><button aria-label="Reply in thread">1 reply</button></div>`;
  const addPopup = e => {
    document.getElementById('autocomplete-picker-list')?.remove();
    if (!e.innerText.includes('@alice@example.com') && !e.innerText.includes('@Alice Example')) return;
    const ul = document.createElement('ul'); ul.id = 'autocomplete-picker-list'; ul.setAttribute('role', 'listbox');
    const li = document.createElement('li'); li.setAttribute('role', 'option'); li.setAttribute('itemtype', 'person'); li.setAttribute('aria-label', 'Alice Example alice@example.com'); li.innerHTML = '<span data-person-mri="8:orgid:fixture-alice">Alice Example</span><span>alice@example.com</span>';
    li.onclick = () => {
      const walker = document.createTreeWalker(e, NodeFilter.SHOW_TEXT); let n;
      while ((n = walker.nextNode())) {
        const query = n.data.includes('@alice@example.com') ? '@alice@example.com' : '@Alice Example';
        const index = n.data.indexOf(query); if (index < 0) continue;
        const r = document.createRange(); r.setStart(n, index); r.setEnd(n, index + query.length); r.deleteContents();
        const card = document.createElement('readonly'); card.setAttribute('itemtype', 'http://schema.skype.com/Mention'); card.innerHTML = '<mention mri="8:orgid:fixture-alice" type="person">Alice Example</mention>'; r.insertNode(card); break;
      }
      ul.remove(); e.focus();
    };
    ul.append(li); document.body.append(ul);
  };
  const wire = root => {
    const e = root.querySelector('[data-tid=ckeditor]'), send = root.querySelector('[data-tid$="-send"]');
    e.addEventListener('paste', event => { event.preventDefault(); document.execCommand('insertHTML', false, event.clipboardData.getData('text/html')); });
    e.addEventListener('input', () => addPopup(e));
    send.onclick = () => {
      const list = root.querySelector('[data-tid="message-pane-list-viewport"]');
      const node = document.createElement('div'); node.dataset.tid = 'chat-pane-message'; node.dataset.mid = 'new-' + ++window.extCounts.sent;
      const body = document.createElement('div'); body.setAttribute('data-message-content', ''); body.innerHTML = e.innerHTML; node.append(body); list.append(node); e.innerHTML = '<p><br></p>';
    };
  };
  wire(main);
  const openThread = () => {
    const area = document.querySelector('[data-tid="app-layout-area--end"]'); area.style.display = 'block';
    area.innerHTML = `<div data-tid="message-pane-layout"><div data-tid="message-pane-list-viewport" style="height:180px;overflow-y:auto">${message('100', 'Parent note')}${message('101', 'A reply')}</div><div data-tid="ckeditor" role="textbox" contenteditable="true" id="thread-editor"><p><br></p></div><button data-tid="newMessageCommands-send" data-track-thread-id="channel@thread.tacv2" data-track-child-thread-id="100">Send</button></div>`;
    wire(area);
  };
  main.querySelector('[aria-label="Reply in thread"]').onclick = openThread;
  main.querySelector('[data-tid="chat-pane-message"]').addEventListener('contextmenu', event => {
    event.preventDefault(); document.querySelector('[data-tid="message-actions-menu-renderer-getV9MessageActions"]')?.remove();
    const menu = document.createElement('div'); menu.dataset.tid = 'message-actions-menu-renderer-getV9MessageActions';
    const button = document.createElement('button'); button.dataset.tid = 'message-actions-threaded-reply'; button.textContent = 'Reply in thread'; button.onclick = () => { menu.remove(); openThread(); };
    menu.append(button); document.body.append(menu);
  });
}
export const extendedFixtureHTML = `<style>[contenteditable]{min-height:40px;padding:8px;border:1px solid} [data-tid="app-layout-area--main"]{width:400px;display:inline-block} [data-tid="app-layout-area--end"]{width:400px;display:inline-block;vertical-align:top} [data-tid="chat-pane-message"]{padding:12px} [role=option]{padding:8px;background:#eee}</style>
<div data-tid="simple-collab-dnd-rail"><div role="treeitem" data-item-type="channel" data-fui-tree-item-value="folder/Channel|channel@thread.tacv2"><span id="title-chat-list-item_channel">Example channel</span></div></div><h1 data-tid="channelTitle-text">Example channel</h1>
<div data-tid="app-layout-area--main"><div data-tid="message-pane-layout"><div data-tid="message-pane-list-viewport" style="height:180px;overflow-y:auto"><div data-tid="chat-pane-item"><div data-tid="chat-pane-message" data-mid="100"><div data-message-content>Parent note</div></div><button aria-label="Reply in thread">1 reply</button></div></div><div data-tid="ckeditor" role="textbox" contenteditable="true" id="main-editor"><p><br></p></div><button data-tid="sendMessageCommands-send" data-track-thread-id="channel@thread.tacv2">Send</button></div></div><div data-tid="app-layout-area--end" style="display:none"></div><script>(${initializeExtendedFixture.toString()})()</script>`;

export async function runExtendedDOMSuite(page, dom, selectors, html) {
  const results = [], check = (condition, message) => { if (!condition) throw Error(message); results.push(message); };
  const run = (action, args) => page.evaluate(dom, action, args, selectors);
  const base = { kind: 'channel', title: 'Example channel', threadId: 'channel@thread.tacv2', key: { attribute: 'data-fui-tree-item-value', value: 'folder/Channel|channel@thread.tacv2' } };
  const thread = { ...base, parentMessageId: '100', token: 'fixture-rich-token' };
  const content = [{ type: 'paragraph', runs: [{ text: 'Hello ', marks: ['bold'] }, { mention: { email: 'alice@example.com', name: 'Alice Example' } }, { text: '!', marks: ['italic'] }] }, { type: 'bulletedList', items: [[{ text: 'One', marks: [] }], [{ text: 'Two', marks: [] }]] }];
  const reset = () => page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  await reset();
  check((await run('channels', {})).items[0].threadId === base.threadId, 'lists channels with exact native identities');
  check((await run('openChannel', base)).opened, 'opens the exact channel timeline');
  check((await run('threads', base)).items[0].parentMessageId === '100', 'returns parent IDs for channel threads');
  check((await run('openThread', thread)).parentMessageId === '100', 'opens a native reply pane bound to its exact parent');
  check((await run('messages', thread)).items.length === 2, 'reads only the thread pane rather than mixing both histories');
  await run('prepare', thread);
  const pasted = await run('paste', { ...thread, content });
  check(pasted.mentions[0].email === 'alice@example.com', 'pastes structured formatting and reserves mention placeholders');
  const focused = await run('mentionFocus', { ...thread, index: 0 });
  await page.call('Input.dispatchKeyEvent', { type: 'keyDown', key: '@', code: 'Digit2', windowsVirtualKeyCode: 50, modifiers: 8, text: '@', unmodifiedText: '@' });
  await page.call('Input.dispatchKeyEvent', { type: 'keyUp', key: '@', code: 'Digit2', windowsVirtualKeyCode: 50, modifiers: 8 });
  await page.call('Input.insertText', { text: focused.query.slice(1) });
  check((await run('mentionSelect', { ...thread, index: 0 })).inserted, 'selects a native person by exact email and verifies its MRI');
  const text = 'Hello Alice Example!\nOne\nTwo';
  check((await run('checkContent', { ...thread, content, text })).verified, 'checks text, native mentions and retained rich formatting');
  check((await run('commit', { ...thread, content, text })).status === 'observed_in_channel_thread', 'sends only to the verified channel reply pane');
  check(await page.evaluate(() => !document.getElementById('main-editor').innerText.trim() && window.extCounts.sent === 1), 'leaves the main channel composer untouched');
  await run('release', thread);
  await reset(); await page.evaluate(() => document.querySelector('[data-tid="app-layout-area--main"] [aria-label="Reply in thread"]').remove());
  const menu = await run('openThread', thread);
  check(menu.requiresContextMenu, 'handles a channel message with no existing reply counter');
  for (const type of ['mousePressed', 'mouseReleased']) await page.call('Input.dispatchMouseEvent', { type, x: menu.x, y: menu.y, button: 'right', clickCount: 1 });
  check((await run('threadMenu', thread)).opened, 'opens the exact parent using its guarded native Reply in thread menu');
  await reset(); await page.evaluate(() => document.querySelector('[data-tid="app-layout-area--main"] [aria-label="Reply in thread"]').remove());
  const stale = await run('openThread', thread);
  for (const type of ['mousePressed', 'mouseReleased']) await page.call('Input.dispatchMouseEvent', { type, x: stale.x, y: stale.y, button: 'right', clickCount: 1 });
  await page.evaluate(() => document.querySelector('[data-tid="app-layout-area--main"] [data-tid$="-send"]').setAttribute('data-track-thread-id', 'other@thread.tacv2'));
  check((await run('threadMenu', thread)).error.code === 'thread_changed', 'refuses a thread menu after channel identity changes');
  await reset(); await run('openThread', thread); await run('prepare', thread);
  await run('paste', { ...thread, content: [{ type: 'paragraph', runs: [{ text: 'Guarded', marks: [] }] }] });
  await run('checkContent', { ...thread, text: 'Guarded' });
  await page.evaluate(() => document.querySelector('[data-track-child-thread-id]').setAttribute('data-track-child-thread-id', 'other-parent'));
  check((await run('commit', { ...thread, text: 'Guarded' })).error.code === 'channel_changed' && await page.evaluate(() => window.extCounts.sent === 0), 'blocks a send if the parent thread changes');
  await run('release', thread);
  await reset(); await run('openThread', thread); await run('prepare', thread); await run('paste', { ...thread, content });
  await run('mentionFocus', { ...thread, index: 0 }); await page.evaluate(() => document.getElementById('main-editor').focus());
  await page.call('Input.insertText', { text: '@alice@example.com' });
  check(await page.evaluate(() => !document.getElementById('main-editor').innerText.trim()), 'blocks mention typing if focus moves into another composer');
  await run('release', thread);
  await reset(); await run('openThread', thread); await run('prepare', thread); await run('paste', { ...thread, content: [{ type: 'paragraph', runs: [{ text: 'Bold', marks: ['bold'] }] }] });
  await page.evaluate(() => { const e = document.getElementById('thread-editor'); e.textContent = 'Bold'; });
  check((await run('checkContent', { ...thread, content: [{ type: 'paragraph', runs: [{ text: 'Bold', marks: ['bold'] }] }], text: 'Bold' })).error.code === 'format_not_preserved', 'refuses silently flattened rich content');
  await run('release', thread);
  await reset(); await run('openThread', thread); await run('prepare', thread); await run('paste', { ...thread, content: [{ type: 'paragraph', runs: [{ text: '@ordinary', marks: [] }] }] });
  check((await run('checkContent', { ...thread, text: 'ordinary' })).error.code === 'draft_mismatch', 'does not ignore literal @ characters when checking message text');
  await run('release', thread);
  await reset(); await page.evaluate(() => document.getElementById('main-editor').textContent = 'Manual draft');
  check((await run('openThread', thread)).error.code === 'draft_exists', 'preserves existing channel drafts');
  await reset();
  const recipient = { title: 'Alice Example', key: { attribute: 'recipient', value: 'alice@example.com' }, threadId: null, token: 'first-rich-message' };
  await page.evaluate(() => {
    const heading = document.querySelector('[data-tid="channelTitle-text"]'); heading.dataset.tid = 'chat-title'; heading.textContent = 'Alice Example';
    document.querySelector('[data-tid="message-pane-list-viewport"]').remove();
    const empty = document.createElement('div'); empty.dataset.tid = 'chat-pane-new-chat'; document.body.append(empty);
    const e = document.getElementById('main-editor'), send = document.querySelector('[data-tid="sendMessageCommands-send"]'); send.removeAttribute('data-track-thread-id');
    window[Symbol.for('switchboard.teams.recipient')] = { email: 'alice@example.com', composer: e, send, clicked: false };
    send.onclick = () => {
      const list = document.createElement('div'); list.dataset.tid = 'message-pane-list-viewport'; list.innerHTML = '<div data-tid="chat-pane-message" data-mid="first-rich"><div data-message-content>' + e.innerHTML + '</div></div>'; document.body.append(list);
      const replacement = e.cloneNode(false); replacement.innerHTML = '<p><br></p>'; e.replaceWith(replacement); empty.remove(); send.setAttribute('data-track-thread-id', 'new-thread');
    };
  });
  check(Array.isArray((await run('prepare', recipient)).beforeIds), 'prepares a verified new recipient with no existing history');
  const firstContent = [{ type: 'paragraph', runs: [{ text: 'First rich message', marks: ['bold'] }] }];
  await run('paste', { ...recipient, content: firstContent }); await run('checkContent', { ...recipient, content: firstContent, text: 'First rich message' });
  check((await run('commit', { ...recipient, text: 'First rich message' })).message?.id === 'first-rich', 'confirms a first rich message even when Teams replaces the new-chat composer');
  await run('release', recipient);
  return results;
}

export async function runChannelSearchDOMSuite(page, baseDOM, dom, selectors, html) {
  const results = [], check = (condition, label) => { if (!condition) throw Error(label); results.push(label); };
  await page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  await page.evaluate(() => {
    const input = document.createElement('input'); input.dataset.tid = 'AUTOSUGGEST_INPUT'; input.value = 'find'; document.body.append(input);
    const content = document.createElement('div'); content.dataset.tid = 'search-content'; content.innerHTML = '<div data-tid="search-card"><div role="row"><div data-tid="message-app-card-header">Example channel</div><span id="serp-message-card-content-fixture">Result</span></div></div>';
    document.body.append(content); window.searchOpenClicks = 0;
    content.querySelector('[role=row]').onclick = () => { window.searchOpenClicks++; document.querySelector('[aria-label="Reply in thread"]').click(); };
  });
  const args = { query: 'find', resultKey: 'serp-message-card-content-fixture' };
  const open = () => page.evaluate(baseDOM, 'searchOpen', args, selectors);
  check((await open()).channelContext, 'opens a channel result only after observing a context transition');
  const context = await page.evaluate(dom, 'channelSearchContext', args, selectors);
  check(context.kind === 'channel' && context.parentMessageId === '100', 'returns verified channel and parent identities from native search context');
  check((await open()).channelContext && await page.evaluate(() => window.searchOpenClicks === 1), 'reuses only a result-bound channel context with unchanged native controls');
  await page.evaluate(() => { document.querySelector('[data-track-child-thread-id]').setAttribute('data-track-child-thread-id', 'other-parent'); document.querySelector('[data-tid="search-card"] [role=row]').onclick = () => window.searchOpenClicks++; });
  check((await open()).error?.code === 'search_context_unavailable', 'does not reuse an unrelated previously open channel thread');
  return results;
}

export async function runRichEditDOMSuite(page, dom, actions, selectors, html) {
  const results = [], check = (condition, label) => { if (!condition) throw Error(label); results.push(label); };
  await page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  const args = { title: 'Alice', threadId: 'thread-a', kind: 'edit', messageId: '200', expectedText: 'Original note', token: 'fixture-rich-edit', content: [{ type: 'paragraph', runs: [{ text: 'Updated', marks: ['bold'] }] }] };
  const run = (action, extra = {}) => page.evaluate(dom, action, { ...args, ...extra }, selectors);
  const act = action => page.evaluate(actions, action, args, selectors);
  // Existing links cannot safely be flattened by the old plain-text edit path.
  await page.evaluate(() => document.querySelector('[data-mid="200"] [data-message-content]').innerHTML = '<a href="https://example.com">Original note</a>');
  const target = await act('messageTarget');
  check(!target.error, 'structured edits explicitly permit replacing existing rich text');
  for (const type of ['mousePressed', 'mouseReleased']) await page.call('Input.dispatchMouseEvent', { type, ...target, button: 'right', clickCount: 1 });
  await act('menuPrepare'); await act('editPrepare');
  check((await run('adoptEdit')).adopted, 'adopts only the verified own-message inline editor'); await act('release');
  await page.evaluate(() => { const e = document.querySelector('[data-tid="chat-pane-item"] [data-tid="ckeditor"]'); e.addEventListener('paste', event => { event.preventDefault(); document.execCommand('insertHTML', false, event.clipboardData.getData('text/html')); }); });
  await run('paste', { replace: true, content: args.content });
  check((await run('checkContent', { text: 'Updated', content: args.content })).verified, 'rich edit replacement retains requested native formatting');
  const edited = await run('editCommit', { text: 'Updated' });
  check(edited.status === 'edited_in_ui' && edited.message.richText.includes('<strong>Updated</strong>'), 'rich edit verifies the same message ID and rich body after saving');
  check(await page.evaluate(() => !document.querySelector('#fixture-composer').innerText.trim()), 'rich editing preserves the main composer');
  await run('release'); return results;
}

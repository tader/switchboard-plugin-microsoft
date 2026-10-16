import { fixtureHTML } from './dom-suite.js';

function initializeDiscoveryFixture() {
  const rail = document.querySelector('[data-tid="simple-collab-dnd-rail"]');
  rail.style.cssText = 'height:120px;overflow:auto';
  rail.replaceChildren();
  const spacer = document.createElement('div'); spacer.style.height = '600px'; rail.append(spacer);
  const folder = document.createElement('div'); folder.setAttribute('role', 'treeitem'); folder.dataset.itemType = 'custom-folder'; folder.dataset.fuiTreeItemValue = 'folder-a'; folder.setAttribute('aria-expanded', 'false');
  const toggle = document.createElement('div'); toggle.dataset.testid = 'conversation-folder-header'; toggle.textContent = 'Section'; folder.append(toggle);
  const rows = document.createElement('div'); folder.append(rows); spacer.append(folder);
  const unread = document.createElement('button'); unread.dataset.testid = 'simple-collab-left-rail-sticky-filter-toggle-button-UNREAD'; unread.setAttribute('aria-pressed', 'false'); unread.textContent = 'Unread';
  const channels = document.createElement('button'); channels.dataset.testid = 'simple-collab-left-rail-sticky-filter-toggle-button-TEAMS_AND_CHANNELS'; channels.setAttribute('aria-pressed', 'true'); channels.textContent = 'Channels';
  rail.before(unread, channels);
  for (const button of [unread, channels]) button.onclick = () => { button.setAttribute('aria-pressed', String(button.getAttribute('aria-pressed') !== 'true')); render(); };
  window.chatOpens = 0;
  const render = () => {
    rows.replaceChildren(); if (folder.getAttribute('aria-expanded') !== 'true') return;
    const index = Math.min(2, Math.floor(rail.scrollTop / 150));
    for (const value of new Set([index, Math.min(2, index + 1)])) {
      const row = document.createElement('div'); row.setAttribute('role', 'treeitem'); row.dataset.itemType = 'chat'; row.dataset.chatId = 'thread-' + value;
      row.setAttribute('aria-labelledby', 'chat_list_unread_text'); const name = document.createElement('span'); name.dataset.tid = 'chat-list-item-title'; name.textContent = 'Unread ' + value; row.append(name); row.onclick = () => window.chatOpens++; rows.append(row);
    }
  };
  toggle.onclick = () => { folder.setAttribute('aria-expanded', String(folder.getAttribute('aria-expanded') !== 'true')); render(); };
  rail.addEventListener('scroll', render);

  const search = document.createElement('input'); search.dataset.tid = 'AUTOSUGGEST_INPUT'; document.body.append(search);
  const tab = document.createElement('button'); tab.dataset.tid = 'messages-tab'; tab.setAttribute('aria-selected', 'false'); tab.textContent = 'Messages'; document.body.append(tab);
  const content = document.createElement('div'); content.dataset.tid = 'search-content'; document.body.append(content);
  window.searchPage = 0;
  const renderSearch = () => {
    content.replaceChildren();
    if (search.value.startsWith('absent')) { const empty = document.createElement('div'); empty.dataset.tid = 'search-no-results'; content.append(empty); return; }
    for (const value of [window.searchPage, window.searchPage + 1]) {
      const card = document.createElement('div'); card.dataset.tid = 'search-card';
      card.innerHTML = '<div role="row"><div data-tid="message-app-card-header">Project chat</div><div role="gridcell"><div id="serp-message-card-content-' + value + '">Alice<br>10:00</div></div><div role="gridcell">Matching text ' + value + '</div></div>';
      card.querySelector('[role="row"]').onclick = () => {
        document.querySelector('[data-tid="chat-title"]').textContent = 'Project chat';
        document.querySelector('[data-tid="sendMessageCommands-send"]').setAttribute('data-track-thread-id', 'search-thread-' + value);
        const list = document.querySelector('[data-tid="message-pane-list-viewport"]');
        list.innerHTML = '<div data-tid="chat-pane-message" data-mid="result-message-' + value + '"><div data-message-content>Matching text ' + value + '</div></div>';
      };
      content.append(card);
    }
    const pagination = document.createElement('div'); pagination.dataset.tid = 'search-pagination-previous-next';
    const previous = document.createElement('button'); previous.textContent = 'Previous'; pagination.append(previous);
    if (window.searchPage < 1) { const next = document.createElement('button'); next.textContent = 'Next'; next.onclick = () => { window.searchPage++; renderSearch(); }; pagination.append(next); }
    content.append(pagination);
  };
  search.addEventListener('keydown', event => { if (event.key === 'Enter') { window.searchPage = 0; tab.setAttribute('aria-selected', 'false'); renderSearch(); } });
  tab.onclick = () => { tab.setAttribute('aria-selected', 'true'); search.value = search.value.trim() + ' is:Messages '; renderSearch(); };
}

export const discoveryFixtureHTML = fixtureHTML.replace('</body>', '<script>(' + initializeDiscoveryFixture.toString() + ')()</script></body>');

export async function runDiscoveryDOMSuite(page, dom, selectors, html) {
  const results = [], check = (condition, label) => { if (!condition) throw new Error(label); results.push(label); };
  const run = (action, args = {}) => page.evaluate(dom, action, args, selectors);
  await page.evaluate(html => { document.open(); document.write(html); document.close(); }, html);
  await run('scanBegin', { token: 'scan', unreadOnly: true });
  const first = await run('scanWindow', { token: 'scan' });
  check(first.items.length === 2 && !first.atEnd, 'expands collapsed sections and reads the first virtualized unread window');
  check(await page.evaluate(() => document.querySelector('[data-testid$="button-UNREAD"]').getAttribute('aria-pressed') === 'true' && document.querySelector('[data-testid$="button-TEAMS_AND_CHANNELS"]').getAttribute('aria-pressed') === 'false'), 'clears competing sidebar filters during unread discovery');
  const ids = new Set(first.items.map(n => n.threadId)); let last;
  for (let i = 0; i < 8; i++) { last = await run('scanNext', { token: 'scan' }); last.items.forEach(n => ids.add(n.threadId)); if (last.atEnd) break; }
  check(ids.size === 3 && last.atEnd, 'discovers unread chats outside the initial DOM window');
  await run('scanEnd', { token: 'scan' });
  check(await page.evaluate(() => document.querySelector('[aria-expanded]').getAttribute('aria-expanded') === 'false' && document.querySelector('[data-testid$="button-TEAMS_AND_CHANNELS"]').getAttribute('aria-pressed') === 'true' && document.querySelector('[data-testid$="button-UNREAD"]').getAttribute('aria-pressed') === 'false'), 'restores collapsed sections and prior filter settings');
  check(await page.evaluate(() => window.chatOpens === 0), 'discovery never opens chats or deliberately clears unread flags');
  await page.evaluate(() => { const input = document.createElement('input'); input.placeholder = 'Filter by name or group name'; input.value = 'A name'; document.body.append(input); });
  check((await run('scanBegin', { token: 'filtered', unreadOnly: true })).error.code === 'sidebar_filtered', 'refuses to claim discovery while a sidebar name filter hides results');
  await page.evaluate(() => document.querySelector('input[placeholder]').remove());
  const search = async query => {
    const focus = await run('searchFocus'); if (focus.error) throw new Error(focus.error.message); await page.call('Input.insertText', { text: query });
    const prepared = await run('searchSubmit', { query }); if (prepared.error) throw new Error(prepared.error.message);
    await page.call('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r', unmodifiedText: '\r' });
    await page.call('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
    return run('searchResults', { query });
  };
  const found = await search('matching');
  check(found.items.length === 2 && found.items[0].author === 'Alice' && found.items[0].timestamp === '10:00' && found.items[0].text === 'Matching text 0' && found.hasNext, 'extracts native search metadata and tolerates Teams is:Messages suffix');
  await run('searchNext', { query: 'matching' });
  const second = await run('searchResults', { query: 'matching' });
  check(!second.hasNext && second.items[0].key === found.items[1].key, 'paginates native search and does not mistake Previous for Next');
  const context = await run('searchOpen', { query: 'matching', resultKey: second.items[0].key });
  check(context.threadId === 'search-thread-1' && context.title === 'Project chat', 'opens the exact search result into a verified conversation');
  check((await run('searchOpen', { query: 'matching', resultKey: second.items[0].key })).threadId === context.threadId, 'reuses only the verified result context with the same composer and thread');
  check((await run('searchResults', { query: 'other' })).error.code === 'search_changed', 'refuses search results for a changed query');
  check((await search('absent')).items.length === 0, 'returns empty search results only with a native empty state');
  await run('searchFocus');
  await page.evaluate(() => document.querySelector('[contenteditable]').focus());
  await page.call('Input.insertText', { text: 'Misdirected query' });
  check(await page.evaluate(() => !document.querySelector('[contenteditable]').innerText.trim()), 'blocks search text input if focus shifts to the composer');
  await run('searchFocus'); await page.call('Input.insertText', { text: 'matching' }); await run('searchSubmit', { query: 'matching' });
  await page.evaluate(() => { const editor = document.querySelector('[contenteditable]'); window.accidentalSends = 0; editor.addEventListener('keydown', e => { if (e.key === 'Enter') window.accidentalSends++; }); editor.focus(); });
  await page.call('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r', unmodifiedText: '\r' });
  await page.call('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
  check(await page.evaluate(() => window.accidentalSends === 0 && !document.querySelector('[contenteditable]').innerText.trim()), 'blocks search Enter and keypress after a composer focus race');
  await page.evaluate(() => document.querySelector('[contenteditable]').textContent = 'Manual draft');
  check((await run('searchFocus')).error.code === 'draft_exists', 'message search preserves a manual draft');
  return results;
}

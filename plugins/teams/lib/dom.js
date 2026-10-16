import { ACTION_SELECTORS } from './actions-dom.js';
import { EXTENDED_SELECTORS } from './extended-dom.js';
export const DEFAULT_SELECTORS = Object.freeze({
  ...ACTION_SELECTORS,
  ...EXTENDED_SELECTORS,
  chatApp: '[data-tid="3b64df9d-7e97-4d9c-ac5c-2e0a5d8e6f40"]',
  chatList: '[data-tid="simple-collab-dnd-rail"], [data-tid="chat-list"], [data-tid="chatList"]',
  chatRow: '[role="treeitem"][data-item-type="chat"], [data-tid="chat-list-item"]',
  chatTitle: '[id^="title-chat-list-item_"], [data-tid="chat-list-item-title"], [data-tid="chat-pane-item-title"]',
  header: '[data-tid="chat-header-title"], [data-tid="chat-title"]',
  messageList: '[data-tid="message-pane-list-viewport"], [data-tid="message-pane-list"], [data-tid="chat-message-list"]',
  message: '[data-tid="chat-pane-message"], [data-tid="message-pane-message"], [data-message-id]',
  messageBody: '[data-message-content], [data-tid="message-body"], [data-tid="chat-pane-message-content"]',
  author: '[data-tid="message-author-name"], [data-tid="message-author"]',
  timestamp: 'time, [data-tid="message-timestamp"]',
  composer: '[data-tid="ckeditor"][contenteditable="true"], [data-tid="ckeditor"] [contenteditable="true"], [role="textbox"][contenteditable="true"]',
  send: '[data-tid="sendMessageCommands-send"], [data-tid="newMessageCommands-send"], [data-tid="sendMessageButton"], [data-tid="send-message-button"]',
  pending: '[data-tid="message-status-sending"], [data-tid="message-status-failed"], [data-tid="message-error"]',
  chatPreview: '[id^="message-preview-chat-list-item_"], [data-tid="chat-list-item-preview"]',
  chatTime: '[id^="time-chat-list-item_"], [data-tid="chat-list-item-time"]',
  lastRead: '[data-tid="last-read-line"]',
  newChat: '[data-testid="simple-collab-left-rail-header-new-message-button"]',
  peopleInput: '[data-tid="people-picker-search"]',
  peopleResult: '[data-tid^="people-picker-entry-"][data-type="ADUser"], [data-tid^="people-picker-entry-"][data-type="Person"]',
  participant: '[data-tid^="participant-"]',
  sidebarFilters: '[data-testid^="simple-collab-left-rail-sticky-filter-toggle-button-"]',
  unreadFilter: '[data-testid="simple-collab-left-rail-sticky-filter-toggle-button-UNREAD"]',
  sidebarNameFilter: 'input[placeholder="Filter by name or group name"]',
  folder: '[role="treeitem"][aria-expanded][data-item-type="custom-folder"], [role="treeitem"][aria-expanded][data-item-type="chats"], [role="treeitem"][aria-expanded][data-item-type="team"]',
  folderToggle: '[data-testid="conversation-folder-header"]',
  sidebarMore: '[data-testid="chat-list-item-see-all"], [data-testid="list-item-seeall-chats"]',
  searchInput: '[data-tid="AUTOSUGGEST_INPUT"]',
  searchTab: '[data-tid="messages-tab"]',
  searchContent: '[data-tid="search-content"]',
  searchCard: '[data-tid="search-card"]',
  searchEmpty: '[data-tid="search-no-results"]',
  searchNext: '[data-tid="search-pagination-previous-next"] button',
  quoteMenu: '[data-tid="message-actions-quoted-reply"]',
  quoteRendered: '[data-tid="quoted-reply-card"]',
  quotePreview: '[data-tid="quoted-reply-preview-content"]',
  quoteCard: 'card[itemscope="QuotedReplyCard"]',
  emptyChat: '[data-tid="chat-pane-new-chat"]',
});

export function selectorsFrom(value) {
  if (!value) return { ...DEFAULT_SELECTORS };
  const overrides = typeof value === 'string' ? JSON.parse(value) : value;
  if (!overrides || Array.isArray(overrides) || typeof overrides !== 'object') throw new Error('selectors must be a JSON object');
  for (const [key, selector] of Object.entries(overrides)) {
    if (!(key in DEFAULT_SELECTORS) || typeof selector !== 'string' || !selector.trim() || selector.length > 2000) {
      throw new Error(`Invalid selector override: ${key}`);
    }
  }
  return { ...DEFAULT_SELECTORS, ...overrides };
}

// Executed as one small CDP evaluation; no screenshots or accessibility-tree dumps.
// Keep this function self-contained: Chrome receives its source, not module imports.
export async function teamsDOM(action, args, selectors) {
  const text = node => (node?.innerText ?? node?.textContent ?? '').replace(/\s+/g, ' ').trim();
  const visible = node => !!node && node.getClientRects().length > 0 && getComputedStyle(node).visibility !== 'hidden';
  const all = (key, root = document) => [...root.querySelectorAll(selectors[key])].filter(visible);
  const one = (key, root = document) => {
    const nodes = all(key, root);
    if (nodes.length !== 1) throw fail('selector_mismatch', `Expected one ${key} element; found ${nodes.length}. Run diagnostics and adjust selectors.`);
    return nodes[0];
  };
  const fail = (code, message, status = 409) => ({ error: { code, message, status } });
  const title = row => text(row.querySelector(selectors.chatTitle)) || row.getAttribute('title') || row.getAttribute('aria-label') || '';
  const threadId = row => row.getAttribute('data-fui-tree-item-value')?.split('|').at(-1) || row.getAttribute('data-chat-id') || null;
  const key = row => {
    for (const attribute of ['data-fui-tree-item-value', 'data-chat-id', 'data-itemid', 'id']) {
      const value = row.getAttribute(attribute);
      if (value) return { attribute, value };
    }
    return { attribute: 'title', value: title(row) };
  };
  const chatData = () => rows().map(row => ({
    key: key(row), title: title(row), threadId: threadId(row),
    unread: row.getAttribute('data-is-unread') === 'true' || (row.getAttribute('aria-labelledby') ?? '').split(/\s+/).includes('chat_list_unread_text') || !!row.querySelector('[data-tid="unread-badge"], [data-testid="chat-list-unread-text"]'),
    preview: text(row.querySelector(selectors.chatPreview)) || null,
    updatedAt: text(row.querySelector(selectors.chatTime)) || null,
  })).filter(item => item.title && item.key.value);
  const scrollContainer = root => {
    let node = root;
    while (node && node !== document.body && node !== document.documentElement) {
      if (node.clientHeight > 0 && /auto|scroll/.test(getComputedStyle(node).overflowY)) return node;
      node = node.parentElement;
    }
    return root;
  };
  const quoteId = root => {
    const cards = all('quoteCard', root);
    if (!cards.length) return null;
    if (cards.length !== 1) throw fail('quote_mismatch', 'Expected one quoted message.');
    try { return String(JSON.parse(cards[0].getAttribute('data-itemprops')).messageId); }
    catch { throw fail('quote_identity_unavailable', 'Teams did not expose the quoted message ID.'); }
  };
  const plainText = root => {
    // Text nodes retain native editor line breaks in innerText on the real node;
    // subtract the quote's rendered text rather than inserting HTML into the editor.
    const cards = all('quoteCard', root);
    if (!cards.length) cards.push(...all('quoteRendered', root).map(n => n.closest('[data-track-module-name="messageQuotedReply"]') || n));
    let value = root.innerText ?? root.textContent ?? '';
    for (const card of cards) value = value.replace(card.innerText ?? card.textContent ?? '', '');
    return value.replace(/\s+/g, ' ').trim();
  };
  const assertQuote = () => {
    if (quoteId(one('composer')) !== (args.replyToMessageId ?? null)) throw fail('quote_mismatch', 'The composer quote does not match the requested message. No message was sent.');
  };
  const rows = () => all('chatRow', one('chatList'));
  const heading = () => text(one('header'));
  const recipientKey = Symbol.for('switchboard.teams.recipient');
  const assertChat = () => {
    if (heading() !== args.title) throw fail('chat_changed', 'The selected chat changed. No message was sent.');
    if (args.threadId && one('send').getAttribute('data-track-thread-id') !== args.threadId) throw fail('chat_changed', 'The reply composer belongs to a different Teams thread.');
    if (args.key?.attribute === 'recipient') {
      const proof = window[recipientKey];
      if (!proof || proof.email !== args.key.value || (!proof.clicked && (
        proof.composer !== one('composer') || proof.send !== one('send') || proof.composerId !== one('composer').id ||
        (proof.participants.length && JSON.stringify(proof.participants) !== JSON.stringify(all('participant').map(n => n.getAttribute('data-tid')).sort()))
      ))) throw fail('recipient_changed', 'The verified recipient or composer changed. No message was sent.');
    }
    if (all('messageList').length === 0 && args.key?.attribute === 'recipient' && all('emptyChat').length === 1) return null;
    return one('messageList');
  };
  const messageData = (root, after) => !root ? [] : all('message', root).filter(node => !after || !!(after.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)).map(node => {
    const body = node.querySelector(selectors.messageBody);
    if (!body) return null;
    const mid = node.getAttribute('data-mid') || node.getAttribute('data-message-id');
    const author = (mid && document.getElementById(`author-${mid}`)) || node.querySelector(selectors.author);
    const timestamp = (mid && document.getElementById(`timestamp-${mid}`)) || node.querySelector(selectors.timestamp);
    return {
      id: mid || node.getAttribute('data-itemid') || node.id || null,
      text: (body.innerText ?? body.textContent ?? '').trim(), isOwn: !!node.closest(selectors.ownMessage), replyToMessageId: quoteId(body), hasQuote: all('quoteRendered', body).length > 0, quotePreview: text(body.querySelector(selectors.quotePreview)) || null, author: text(author) || null,
      timestamp: timestamp?.getAttribute('datetime') || text(timestamp) || null,
      pending: !!node.querySelector(selectors.pending),
    };
  }).filter(Boolean);
  const wait = (predicate, milliseconds = 10_000, stable = 0) => new Promise(resolve => {
    let stableTimer;
    let done = false;
    const finish = value => {
      if (done) return;
      done = true;
      observer.disconnect(); clearTimeout(deadline); clearTimeout(stableTimer); resolve(value);
    };
    const check = () => {
      clearTimeout(stableTimer);
      try {
        const value = predicate();
        if (value) {
          if (stable) stableTimer = setTimeout(() => { try { const next = predicate(); if (next) finish(next); } catch {} }, stable);
          else finish(value);
        }
      } catch {}
    };
    const observer = new MutationObserver(check);
    observer.observe(document.documentElement, { subtree: true, childList: true, attributes: true, characterData: true });
    const deadline = setTimeout(() => finish(null), milliseconds);
    check();
  });
  const searchMatches = () => { const value = one('searchInput').value.trim(); return value === args.query || value === `${args.query} is:Messages`; };
  const nextSearch = () => all('searchNext').filter(n => /^(Next|Next page)$/i.test(n.getAttribute('aria-label') || text(n)));
  const people = () => all('peopleResult').map(node => ({
    email: (node.getAttribute('data-tid') ?? '').replace(/^people-picker-entry-/, '').toLowerCase(),
    name: (node.innerText ?? '').split('\n')[0].trim(),
  })).filter(person => /^[^\s,]+@[^\s,]+$/.test(person.email));
  try {
    // Also validates custom selectors early, including selectors not used in this action.
    for (const selector of Object.values(selectors)) document.querySelector(selector);
    const needsSidebar = ['status', 'chats', 'scanBegin', 'scanWindow', 'scanNext', 'scanEnd', 'open', 'peopleFocus', 'peopleResults', 'selectPerson'].includes(action);
    if (needsSidebar && all('chatList').length === 0 && all('chatApp').length === 1) {
      one('chatApp').click();
      await wait(() => all('chatList').length === 1, 10_000, 200);
    }
    if (action === 'status') return {
      ready: all('chatList').length === 1,
      url: location.origin + location.pathname,
      selectorCounts: Object.fromEntries(['chatList', 'chatRow', 'header', 'messageList', 'composer', 'send'].map(k => [k, all(k).length])),
    };
    if (action === 'diagnostics') return {
      url: location.origin + location.pathname,
      counts: Object.fromEntries(Object.keys(selectors).map(k => [k, all(k).length])),
      // Attribute inventory only: never includes message text, tokens, HTML or cookies.
      attributes: [...new Set([...document.querySelectorAll('[data-tid]')].map(n => n.getAttribute('data-tid')))].slice(0, 200),
    };
    if (needsSidebar && all('chatList').length !== 1) return fail('signin_or_layout_required', 'Sign in in the dedicated browser. If already signed in, inspect diagnostics and update selectors.', 503);
    if (action === 'chats') return { items: chatData(), coverage: 'rendered_chat_list' };
    const scanKey = Symbol.for('switchboard.teams.sidebarScan');
    if (action === 'scanBegin') {
      if (window[scanKey]) throw fail('scan_in_progress', 'A previous sidebar scan requires restoration.');
      if (all('sidebarNameFilter').some(n => n.value.trim())) throw fail('sidebar_filtered', 'Clear the sidebar name filter before discovering chats.');
      const filters = all('sidebarFilters');
      const unread = all('unreadFilter');
      if (args.unreadOnly && unread.length !== 1) throw fail('unread_filter_unavailable', 'Could not identify the native Unread filter.');
      const scroll = scrollContainer(one('chatList'));
      const state = { token: args.token, scrollTop: scroll.scrollTop, filters: filters.filter(n => n.getAttribute('aria-pressed') === 'true').map(n => n.getAttribute('data-testid')), folders: [], more: new Set() };
      window[scanKey] = state;
      for (const filter of filters) {
        const desired = !!args.unreadOnly && filter === unread[0];
        if ((filter.getAttribute('aria-pressed') === 'true') !== desired) filter.click();
      }
      if (!await wait(() => all('sidebarFilters').every(n => (n.getAttribute('aria-pressed') === 'true') === (!!args.unreadOnly && n.matches(selectors.unreadFilter))), 3000, 250)) throw fail('sidebar_filter_not_ready', 'The requested sidebar filter was not applied; discovery is incomplete.');
      scrollContainer(one('chatList')).scrollTop = 0;
      await wait(() => true, 1000, 200);
      return { started: true };
    }
    if (['scanWindow', 'scanNext', 'scanEnd'].includes(action)) {
      const state = window[scanKey];
      if (!state || state.token !== args.token) throw fail('scan_changed', 'Sidebar scan state changed.');
      if (action === 'scanEnd') {
        try {
          // Restore only controls changed by this scan, identified by stable attributes.
          let unrestoredSections = 0;
          for (const folderKey of state.folders.reverse()) {
            const folder = all('folder').find(n => JSON.stringify(key(n)) === JSON.stringify(folderKey));
            if (!folder) unrestoredSections++;
            if (folder?.getAttribute('aria-expanded') === 'true') folder.querySelector(selectors.folderToggle)?.click();
          }
          for (const filter of all('sidebarFilters')) {
            if ((filter.getAttribute('aria-pressed') === 'true') !== state.filters.includes(filter.getAttribute('data-testid'))) filter.click();
          }
          if (!await wait(() => all('sidebarFilters').every(n => (n.getAttribute('aria-pressed') === 'true') === state.filters.includes(n.getAttribute('data-testid'))), 3000, 200)) throw fail('sidebar_restore_failed', 'The sidebar filters could not be restored. Inspect them manually.');
          scrollContainer(one('chatList')).scrollTop = state.scrollTop;
          return { restored: unrestoredSections === 0, unrestoredSections };
        } finally { delete window[scanKey]; }
      }
      for (let i = 0; i < 30; i++) {
        const folder = all('folder').find(n => n.getAttribute('aria-expanded') === 'false');
        if (!folder) break;
        const toggle = folder.querySelector(selectors.folderToggle);
        if (!toggle) throw fail('folder_toggle_unavailable', 'A collapsed chat section cannot be expanded. Discovery is incomplete.');
        state.folders.push(key(folder)); toggle.click();
        if (!await wait(() => folder.getAttribute('aria-expanded') === 'true', 2000, 100)) throw fail('folder_not_ready', 'Could not expand a chat section.');
      }
      const scroll = scrollContainer(one('chatList'));
      if (action === 'scanNext') {
        const before = JSON.stringify(chatData().map(n => n.key));
        scroll.scrollTop += Math.max(100, scroll.clientHeight * 0.8);
        await wait(() => JSON.stringify(chatData().map(n => n.key)) !== before, 600, 100);
      }
      const more = all('sidebarMore', one('chatList')).find(n => !state.more.has(JSON.stringify(key(n))));
      if (more) { state.more.add(JSON.stringify(key(more))); more.click(); await wait(() => true, 1200, 250); }
      return { items: chatData(), atEnd: scroll.scrollTop + scroll.clientHeight >= scroll.scrollHeight - 2, position: scroll.scrollTop, height: scroll.scrollHeight, collapsed: all('folder').some(n => n.getAttribute('aria-expanded') === 'false'), more: !!more };
    }
    if (action === 'searchFocus') {
      if (all('composer').some(n => text(n) || all('quoteCard', n).length)) throw fail('draft_exists', 'Clear or send the existing draft before searching messages.');
      const input = one('searchInput'); input.focus(); input.select();
      const guardKey = Symbol.for('switchboard.teams.searchInputGuard');
      if (window[guardKey]) document.removeEventListener('beforeinput', window[guardKey], true);
      const guard = event => {
        if (event.target !== input || document.activeElement !== input) { event.preventDefault(); event.stopImmediatePropagation(); }
        document.removeEventListener('beforeinput', guard, true);
        if (window[guardKey] === guard) delete window[guardKey];
      };
      window[guardKey] = guard; document.addEventListener('beforeinput', guard, true);
      // Retain the one-input guard until input occurs or a new search replaces it.
      // A delayed CDP command must not outlive its protection.
      return { focused: document.activeElement === input };
    }
    if (action === 'searchSubmit') {
      if (all('composer').some(n => text(n) || all('quoteCard', n).length)) throw fail('draft_exists', 'Clear or send the existing draft before searching messages.');
      const input = one('searchInput');
      if (input.value !== args.query || document.activeElement !== input) throw fail('search_focus_changed', 'Search input lost focus or changed. No Enter event was dispatched.');
      // Cancel Enter synchronously if Teams moves focus to the composer between
      // this evaluation and the trusted key event. Never submit an editor draft.
      const guardKey = Symbol.for('switchboard.teams.searchGuard');
      if (window[guardKey]) for (const type of ['keydown', 'keypress', 'keyup']) document.removeEventListener(type, window[guardKey], true);
      const deadline = Date.now() + 5000;
      const guard = event => {
        if (event.key !== 'Enter') return;
        if (Date.now() > deadline || event.target !== input || document.activeElement !== input || input.value !== args.query) {
          event.preventDefault(); event.stopImmediatePropagation();
        }
        if (event.type === 'keyup') { for (const type of ['keydown', 'keypress', 'keyup']) document.removeEventListener(type, guard, true); if (window[guardKey] === guard) delete window[guardKey]; }
      };
      window[guardKey] = guard; for (const type of ['keydown', 'keypress', 'keyup']) document.addEventListener(type, guard, true);
      // Keep the guard until Enter keyup (or the next search), even on timeout.
      return { enterPrepared: true };
    }
    if (action === 'searchResults') {
      if (!searchMatches()) throw fail('search_changed', 'The search query changed.');
      if (!await wait(() => all('searchTab').length === 1, 8000)) throw fail('search_unavailable', 'Message search did not open.');
      if (one('searchTab').getAttribute('aria-selected') !== 'true') one('searchTab').click();
      const result = await wait(() => {
        if (one('searchTab').getAttribute('aria-selected') !== 'true') return false;
        const content = one('searchContent');
        if (content.querySelector('[role="progressbar"], [aria-busy="true"]')) return false;
        const cards = all('searchCard', content).filter(n => n.querySelector('[id^="serp-message-card-content-"]'));
        return cards.length ? cards : all('searchEmpty', content).length ? [] : false;
      }, 10_000, 400);
      if (!result) throw fail('search_not_ready', 'Message results could not be verified; no empty-results claim was made.');
      if (!searchMatches()) throw fail('search_changed', 'The search query changed.');
      return { items: result.map(card => {
        const metadata = card.querySelector('[id^="serp-message-card-content-"]');
        const lines = (metadata.innerText ?? '').split('\n').map(n => n.trim()).filter(Boolean);
        const cells = [...card.querySelectorAll('[role="gridcell"]')];
        const header = card.querySelector('[data-tid="message-app-card-header"]');
        return { key: metadata.id, author: lines[0] ?? null, timestamp: lines.slice(1).join(' ') || null, conversation: (header?.innerText ?? '').split('\n')[0].trim() || null, text: (cells[1]?.innerText ?? '').trim() };
      }), hasNext: nextSearch().some(n => !n.disabled && n.getAttribute('aria-disabled') !== 'true'), coverage: 'native_message_search' };
    }
    if (action === 'searchOpen') {
      if (!searchMatches()) throw fail('search_changed', 'The search query changed.');
      if (all('composer').some(n => text(n) || all('quoteCard', n).length)) throw fail('draft_exists', 'Clear or send the existing draft before opening a search result.');
      const matches = all('searchCard', one('searchContent')).filter(n => n.querySelector('[id^="serp-message-card-content-"]')?.id === args.resultKey);
      if (matches.length !== 1) throw fail('search_result_unavailable', 'The selected search result is absent or ambiguous.', 404);
      const row = matches[0].querySelector('[role="row"]'); if (!row) throw fail('search_result_unavailable', 'The selected result cannot be opened.');
      const contextKey = Symbol.for('switchboard.teams.searchContext');
      const proof = window[contextKey];
      const channelProof = window[Symbol.for('switchboard.teams.channelSearchContext')];
      if (channelProof?.query === args.query && channelProof.resultKey === args.resultKey && channelProof.composer.isConnected && channelProof.send.isConnected && all('composer').includes(channelProof.composer) && all('send').includes(channelProof.send) && channelProof.send.getAttribute('data-track-thread-id') === channelProof.chat.threadId && channelProof.send.getAttribute('data-track-child-thread-id') === channelProof.chat.parentMessageId && all('channelHeader').length === 1 && text(one('channelHeader')) === channelProof.chat.title) return { channelContext: true };
      if (proof && proof.query === args.query && proof.resultKey === args.resultKey && proof.composer === all('composer')[0] && proof.send === all('send')[0] && heading() === proof.chat.title && one('send').getAttribute('data-track-thread-id') === proof.chat.threadId) return proof.chat;
      const expectedTitle = (matches[0].querySelector('[data-tid="message-app-card-header"]')?.innerText ?? '').split('\n')[0].trim().replace(/^Chat with /, '');
      const previousComposer = all('composer')[0], previousThread = all('send')[0]?.getAttribute('data-track-thread-id');
      const previousIds = all('message').map(n => n.getAttribute('data-mid') || n.getAttribute('data-message-id')).join('|');
      row.click();
      const contextChanged = () => !previousComposer || previousComposer !== all('composer')[0] || previousThread !== all('send')[0]?.getAttribute('data-track-thread-id') || previousIds !== all('message').map(n => n.getAttribute('data-mid') || n.getAttribute('data-message-id')).join('|');
      const channelContext = () => contextChanged() && all('channelHeader').length === 1 && all('send').some(n => n.getAttribute('data-track-child-thread-id'));
      if (!await wait(() => channelContext() || (all('header').length === 1 && heading() === expectedTitle && all('send').length === 1 && all('composer').length === 1 && all('messageList').length === 1 && (!previousComposer || previousComposer !== one('composer') || previousThread !== one('send').getAttribute('data-track-thread-id') || previousIds !== all('message').map(n => n.getAttribute('data-mid') || n.getAttribute('data-message-id')).join('|'))), 8000, 500)) throw fail('search_context_unavailable', 'The message context did not open into a verified conversation.');
      if (channelContext()) return { channelContext: true };
      const thread = one('send').getAttribute('data-track-thread-id');
      if (!thread || /@thread\.tacv2$/.test(thread)) throw fail('search_context_unsupported', 'Search can return channel messages; opening channel threads for replies is not yet supported.');
      const chat = { title: heading(), threadId: thread, key: { attribute: 'search', value: JSON.stringify({ query: args.query, resultKey: args.resultKey }) } };
      window[contextKey] = { query: args.query, resultKey: args.resultKey, chat, composer: one('composer'), send: one('send') };
      return chat;
    }
    if (action === 'searchNext') {
      if (!searchMatches()) throw fail('search_changed', 'The search query changed.');
      const matches = nextSearch(); if (matches.length !== 1) throw fail('search_pagination_unavailable', 'No unique Next page button.'); const next = matches[0];
      if (next.disabled || next.getAttribute('aria-disabled') === 'true') return { advanced: false };
      const before = one('searchContent').innerText; next.click();
      if (!await wait(() => one('searchContent').innerText !== before, 8000, 300)) throw fail('search_not_ready', 'Search pagination did not advance.');
      return { advanced: true };
    }
    if (action === 'quoteTarget') {
      const list = assertChat();
      if (text(one('composer')) || all('quoteCard', one('composer')).length) throw fail('draft_exists', 'Clear or send the existing draft before quoting a message.');
      const targets = all('message', list).filter(n => (n.getAttribute('data-mid') || n.getAttribute('data-message-id') || n.getAttribute('data-itemid') || n.id) === args.replyToMessageId);
      if (targets.length !== 1) throw fail('message_unavailable', 'Quoted message is absent or ambiguous in the rendered history. Load older messages first.', 404);
      targets[0].scrollIntoView({ block: 'center' });
      const rect = targets[0].getBoundingClientRect(); const x = rect.x + rect.width / 2, y = rect.y + rect.height / 2;
      if (!targets[0].contains(document.elementFromPoint(x, y))) throw fail('message_obscured', 'The quoted message is obscured.');
      return { x, y };
    }
    if (action === 'quoteSelect') {
      assertChat();
      if (text(one('composer')) || all('quoteCard', one('composer')).length) throw fail('draft_exists', 'The composer changed before quoting.');
      if (!await wait(() => all('quoteMenu').length === 1, 3000)) throw fail('quote_unavailable', 'The message has no unique Reply with quote action.');
      one('quoteMenu').click();
      if (!await wait(() => quoteId(one('composer')) === args.replyToMessageId, 3000, 100)) throw fail('quote_mismatch', 'Teams did not insert the requested message quote. Inspect the draft manually.');
      assertChat(); assertQuote();
      const composer = one('composer'); composer.focus();
      const range = document.createRange(); range.selectNodeContents(composer); range.collapse(false);
      const selection = getSelection(); selection.removeAllRanges(); selection.addRange(range);
      const before = messageData(assertChat()); if (before.some(m => !m.id)) throw fail('message_ids_unavailable', 'Message IDs are required to verify sends.');
      const quotePreview = JSON.parse(one('quoteCard', composer).getAttribute('data-itemprops')).preview;
      if (typeof quotePreview !== 'string' || !quotePreview.trim()) throw fail('quote_identity_unavailable', 'The quote has no verifiable preview.');
      return { beforeIds: before.map(m => m.id), replyToMessageId: args.replyToMessageId, quotePreview: quotePreview.replace(/\s+/g, ' ').trim() };
    }
    if (action === 'peopleFocus') {
      if (all('composer').some(node => text(node))) return fail('draft_exists', 'The composer contains a draft. Clear or send it manually before searching people.');
      one('newChat').click();
      if (!await wait(() => all('peopleInput').length === 1, 5000)) return fail('people_picker_unavailable', 'Could not open the Teams people picker.');
      const input = one('peopleInput');
      const previousValue = input.value;
      input.focus(); input.select();
      return { previousValue, before: JSON.stringify(people()) };
    }
    if (action === 'peopleResults') {
      const result = await wait(() => {
        if (one('peopleInput').value !== args.query) throw fail('search_changed', 'The people search changed.');
        const query = args.query.toLowerCase();
        // The picker can briefly render recent contacts before search finishes.
        // Never mistake those unrelated suggestions for the requested results.
        const items = people().filter(person => query.includes('@') ? person.email === query : query.split(/\s+/).every(part => `${person.name} ${person.email}`.toLowerCase().includes(part)));
        return items.length ? items : false;
      }, 10_000, 300);
      if (one('peopleInput').value !== args.query) return fail('search_changed', 'The people search changed.');
      return { items: result ?? [], coverage: 'rendered_people_results' };
    }
    if (action === 'selectPerson') {
      if (one('peopleInput').value.toLowerCase() !== args.email) return fail('search_changed', 'Search must contain the exact recipient email before selection.');
      const matches = all('peopleResult').filter(node => (node.getAttribute('data-tid') ?? '').toLowerCase() === `people-picker-entry-${args.email}`);
      if (matches.length !== 1) return fail('recipient_unavailable', 'No unique directory person matches that email. Search people and use a returned email.', 404);
      const name = (matches[0].innerText ?? '').split('\n')[0].trim();
      matches[0].click();
      const ready = await wait(() => all('peopleInput').length === 0 && heading() === name && all('composer').length === 1 && all('send').length === 1 && (all('messageList').length === 1 || all('emptyChat').length === 1), 10_000, 500);
      if (!ready) return fail('recipient_not_ready', 'The selected person did not produce a verified conversation. No message was sent.');
      const composer = one('composer'), send = one('send');
      const participants = all('participant').map(n => n.getAttribute('data-tid')).sort();
      if (!send.getAttribute('data-track-thread-id') && !composer.id && !participants.length) return fail('recipient_identity_unavailable', 'The new chat has no stable composer or participant identity. No message was sent.');
      window[recipientKey] = { email: args.email, composer, composerId: composer.id, participants, send, clicked: false };
      return { key: { attribute: 'recipient', value: args.email }, title: heading(), threadId: send.getAttribute('data-track-thread-id') || null };
    }
    if (action === 'open') {
      const matches = rows().filter(row => JSON.stringify(key(row)) === JSON.stringify(args.key));
      if (matches.length !== 1) return fail('chat_unavailable', 'Chat is absent or ambiguous in the rendered list. Scroll it into view manually and list chats again.');
      if ((args.threadId ?? null) !== threadId(matches[0])) return fail('invalid_chat', 'Chat identity does not match the sidebar. List chats again.', 400);
      if (!args.threadId && rows().filter(row => title(row) === args.title).length !== 1) return fail('ambiguous_title', 'Multiple rendered chats have this title. Rename the chat or adjust the sidebar before sending.');
      if (title(matches[0]) !== args.title) return fail('chat_changed', 'The chat title changed. List chats again.');
      matches[0].scrollIntoView({ block: 'nearest' }); matches[0].click();
      const ready = await wait(() => heading() === args.title && all('composer').length === 1 && all('messageList').length === 1 && (!args.threadId || one('send').getAttribute('data-track-thread-id') === args.threadId), 10_000, 300);
      return ready ? { title: heading() } : fail('chat_not_ready', 'Could not verify chat title and composer. No message was sent.');
    }
    if (action === 'messages') return { items: messageData(assertChat()).slice(-args.limit), coverage: 'rendered_message_window' };
    if (action === 'unreadMessages') {
      const list = assertChat();
      const markers = list ? all('lastRead', list) : [];
      if (markers.length !== 1) return { items: [], recentMessages: messageData(list).slice(-args.limit), boundaryFound: false, coverage: 'rendered_message_window' };
      const unread = messageData(list, markers[0]);
      return { items: unread.slice(0, args.limit), boundaryFound: true, truncated: unread.length > args.limit, coverage: 'rendered_messages_after_last_read' };
    }
    if (action === 'older') {
      const list = assertChat();
      let scroll = list;
      while (scroll && scroll.scrollHeight <= scroll.clientHeight) scroll = scroll.parentElement;
      if (!scroll || scroll === document.body || scroll === document.documentElement) return fail('history_unavailable', 'Could not identify the message history scroll container.');
      const before = JSON.stringify(messageData(list).map(m => [m.id, m.text]));
      scroll.scrollTop = Math.max(0, scroll.scrollTop - Math.max(scroll.clientHeight, 600));
      const changed = await wait(() => JSON.stringify(messageData(assertChat()).map(m => [m.id, m.text])) !== before, 1500, 100);
      return { changed: !!changed, items: messageData(assertChat()) };
    }
    if (action === 'prepare') {
      const list = assertChat();
      const composer = one('composer');
      if (text(composer) || all('quoteCard', composer).length || all('quoteRendered', composer).length) return fail('draft_exists', 'The composer contains a draft. Clear or send it manually first.');
      if (composer.getAttribute('aria-disabled') === 'true') return fail('composer_disabled', 'This chat does not allow replies.');
      const before = messageData(list);
      if (before.some(m => !m.id)) return fail('message_ids_unavailable', 'Message IDs are required to verify sends. Adjust the message selector.');
      composer.focus();
      return { beforeIds: before.map(m => m.id) };
    }
    if (action === 'sendPoint') {
      assertChat(); assertQuote();
      if (plainText(one('composer')) !== args.text.replace(/\s+/g, ' ').trim()) return fail('draft_mismatch', 'Composer text does not match the requested message. No send click was made.');
      const button = one('send');
      if (button.disabled || button.getAttribute('aria-disabled') === 'true') return fail('send_disabled', 'Send is disabled. No send click was made.');
      button.scrollIntoView({ block: 'nearest' });
      const rect = button.getBoundingClientRect();
      const x = rect.x + rect.width / 2, y = rect.y + rect.height / 2;
      if (!button.contains(document.elementFromPoint(x, y))) return fail('send_obscured', 'Send button is obscured. No send click was made.');
      // Recheck synchronously at click dispatch, closing the race between the
      // CDP evaluation and input events if the user navigates in this window.
      const guardKey = Symbol.for('switchboard.teams.sendGuard');
      if (window[guardKey]) document.removeEventListener('click', window[guardKey], true);
      const deadline = Date.now() + 5000;
      const guard = event => {
        if (!event.target.closest?.(selectors.send)) return;
        try {
          assertChat(); assertQuote();
          if (Date.now() > deadline || plainText(one('composer')) !== args.text.replace(/\s+/g, ' ').trim()) throw new Error();
          if (args.key?.attribute === 'recipient') window[recipientKey].clicked = true;
        } catch { event.preventDefault(); event.stopImmediatePropagation(); }
        document.removeEventListener('click', guard, true);
        delete window[guardKey];
      };
      window[guardKey] = guard;
      document.addEventListener('click', guard, true);
      return { x, y };
    }
    if (action === 'confirm') {
      const observed = await wait(() => {
        const list = assertChat();
        if (text(one('composer'))) return false;
        return messageData(list).find(m => m.id && !args.beforeIds.includes(m.id) && !m.pending && (args.replyToMessageId ? m.hasQuote && m.quotePreview === args.quotePreview : !m.hasQuote && !m.replyToMessageId) && (args.replyToMessageId ? plainText(all('message', list).find(n => (n.getAttribute('data-mid') || n.getAttribute('data-message-id') || n.getAttribute('data-itemid') || n.id) === m.id).querySelector(selectors.messageBody)) : m.text).replace(/\s+/g, ' ').trim() === args.text.replace(/\s+/g, ' ').trim());
      }, 10_000, 300);
      return observed ? { status: 'observed_in_chat', message: observed, ...(args.replyToMessageId ? { replyToMessageId: args.replyToMessageId, quoteObserved: true } : {}) } : fail('send_uncertain', 'Send was attempted but could not be confirmed. Check Teams manually; do not retry with a new idempotency key.', 504);
    }
    return fail('unknown_action', 'Unknown DOM action.', 400);
  } catch (error) {
    return error?.error ? error : fail('selector_mismatch', 'DOM selectors do not match this Teams layout. Run diagnostics.', 409);
  }
}

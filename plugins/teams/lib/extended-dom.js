export const EXTENDED_SELECTORS = {
  channelRow: '[role="treeitem"][data-item-type="channel"]',
  channelHeader: '[data-tid="channelTitle-text"]',
  mainPane: '[data-tid="app-layout-area--main"]',
  threadPane: '[data-tid="app-layout-area--end"]',
  paneLayout: '[data-tid="message-pane-layout"]',
  replyButton: 'button',
  historyLoading: '[role="progressbar"], [data-tid="message-list-loading"], [data-tid="message-pane-loading"]',
  mentionOption: '#autocomplete-picker-list [role="option"][itemtype="person"], [role="option"][data-email], [data-tid^="mention-suggestion"], [data-tid^="people-picker-entry-"]',
  mentionCard: '[itemtype="http://schema.skype.com/Mention"], [data-mention-id], [data-tid="mention"]',
  threadedReply: '[data-tid="message-actions-threaded-reply"]',
};

// Runs in the Teams page. Native UI and editor events only; no private APIs.
export async function teamsExtended(action, args, selectors) {
  const fail = (code, message, status = 409) => ({ error: { code, message, status } });
  const visible = n => n && n.getClientRects().length && getComputedStyle(n).visibility !== 'hidden';
  const nodes = (s, root = document) => [...root.querySelectorAll(s)].filter(visible);
  const all = (key, root = document) => nodes(selectors[key], root);
  const one = (key, root = document) => { const n = all(key, root); if (n.length !== 1) throw fail('selector_mismatch', `Expected one ${key}; found ${n.length}.`); return n[0]; };
  const raw = n => (n?.innerText ?? n?.textContent ?? '').trim();
  const norm = s => s.replace(/[\u200b-\u200d\u2060\ufeff]/g, '').replace(/\s+/g, ' ').trim();
  const mid = n => n.getAttribute('data-mid') || n.getAttribute('data-message-id') || n.getAttribute('data-itemid') || n.id;
  const wait = (fn, timeout = 8000, stable = 150) => new Promise(resolve => {
    let done, timer;
    const finish = value => { if (done) return; done = true; observer.disconnect(); clearTimeout(timer); clearTimeout(deadline); resolve(value); };
    const check = () => { clearTimeout(timer); try { if (fn()) timer = setTimeout(() => { try { const value = fn(); if (value) finish(value); } catch {} }, stable); } catch {} };
    const observer = new MutationObserver(check); observer.observe(document.documentElement, { subtree: true, childList: true, attributes: true, characterData: true });
    const deadline = setTimeout(() => finish(null), timeout); check();
  });
  const pane = () => {
    if (args.kind !== 'channel') return document;
    const area = one(args.parentMessageId ? 'threadPane' : 'mainPane');
    const root = one('paneLayout', area), send = all('send', root).filter(n => !n.closest(selectors.messageList));
    if (send.length !== 1 || send[0].getAttribute('data-track-thread-id') !== args.threadId || (send[0].getAttribute('data-track-child-thread-id') || null) !== (args.parentMessageId || null)) throw fail('channel_changed', 'The channel or parent thread changed.');
    if (norm(raw(one('channelHeader'))) !== args.title) throw fail('channel_changed', 'The channel title changed.');
    return root;
  };
  const editor = root => {
    const list = all('messageList', root);
    const candidates = all('composer', root).filter(n => !list.some(l => l.contains(n)));
    if (candidates.length !== 1) throw fail('draft_exists', 'The composer is absent or ambiguous, possibly due to an inline edit.');
    return candidates[0];
  };
  const sender = root => { const list = all('messageList', root); const n = all('send', root).filter(n => !list.some(l => l.contains(n))); if (n.length !== 1) throw fail('selector_mismatch', 'No unique Send control.'); return n[0]; };
  const assertContext = () => {
    const root = pane();
    if (args.kind !== 'channel' && (norm(raw(one('header'))) !== args.title || (args.threadId && sender(root).getAttribute('data-track-thread-id') !== args.threadId))) throw fail('chat_changed', 'The conversation changed.');
    if (args.key?.attribute === 'recipient') {
      const recipient = window[Symbol.for('switchboard.teams.recipient')];
      if (!recipient || recipient.email !== args.key.value || !recipient.clicked && (recipient.composer !== editor(root) || recipient.send !== sender(root))) throw fail('recipient_changed', 'The verified recipient changed.');
    }
    return root;
  };
  const data = root => {
    if (!all('messageList', root).length && args.key?.attribute === 'recipient' && all('emptyChat').length === 1) return [];
    return all('message', one('messageList', root)).map(n => {
    const body = n.querySelector(selectors.messageBody); if (!body) return null;
    const id = mid(n), author = document.getElementById(`author-${id}`) || n.querySelector(selectors.author), time = document.getElementById(`timestamp-${id}`) || n.querySelector(selectors.timestamp);
    return { id, text: raw(body), author: raw(author) || null, timestamp: time?.getAttribute('datetime') || raw(time) || null, isOwn: !!n.closest(selectors.ownMessage), pending: !!n.querySelector(selectors.pending), richText: body.innerHTML, mentions: all('mentionCard', body).map(n => ({ name: raw(n), identity: n.getAttribute('data-mention-id') || n.getAttribute('itemid') || null })) };
    }).filter(Boolean);
  };
  const nativeText = root => {
    let value = raw(root);
    for (const card of all('mentionCard', root)) value = value.replace(raw(card), raw(card).replace(/^@/, ''));
    return norm(value);
  };
  const matchingMessage = (root, predicate) => {
    const item = data(root).find(n => predicate(n) && !n.pending && nativeText(all('message', one('messageList', root)).find(node => mid(node) === n.id).querySelector(selectors.messageBody)) === norm(args.text));
    return item;
  };
  const stateKey = Symbol.for('switchboard.teams.richDraft');
  const threadKey = Symbol.for('switchboard.teams.openThread');
  const proof = () => {
    const p = window[stateKey]; if (!p || p.token !== args.token) throw fail('draft_changed', 'The prepared rich draft changed.');
    const root = assertContext();
    if (!p.editor.isConnected || p.invalidated || (p.edit ? !p.wrapper.isConnected || !p.wrapper.contains(p.editor) || !p.wrapper.contains(p.send) || !p.wrapper.querySelector(selectors.ownMessage) || one('composer', p.wrapper) !== p.editor || one('editSave', p.wrapper) !== p.send : editor(root) !== p.editor || sender(root) !== p.send)) throw fail('draft_changed', 'The prepared editor or target changed.');
    return p;
  };
  const selectedRange = p => { const selection = window.getSelection(); if (document.activeElement !== p.editor || !selection?.rangeCount || !p.editor.contains(selection.getRangeAt(0).commonAncestorContainer)) throw fail('composer_focus_changed', 'Focus or selection left the verified editor.'); };
  const guardInput = (p, persistent = false) => {
    p.editor.focus(); const guard = event => { try { proof(); selectedRange(p); if (event.target !== p.editor) throw Error(); } catch { p.invalidated = true; event.preventDefault(); event.stopImmediatePropagation(); } if (!persistent) document.removeEventListener('beforeinput', guard, true); };
    document.addEventListener('beforeinput', guard, true); p.inputGuards.push(guard);
  };
  const mentionIdentity = n => {
    for (const attr of ['data-email', 'data-upn', 'data-mention-id', 'itemid', 'data-tid']) {
      const value = n.getAttribute(attr); if (value) { const email = value.replace(/^(people-picker-entry-|mention-suggestion-)/, '').toLowerCase(); if (/^[^\s]+@[^\s]+$/.test(email)) return email; }
    }
    const rendered = ((n.getAttribute('aria-label') || '') + ' ' + raw(n)).match(/[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i); return rendered?.[0].toLowerCase();
  };
  try {
    if (action === 'channels') {
      return { items: all('channelRow', one('chatList')).map(row => ({ key: { attribute: 'data-fui-tree-item-value', value: row.getAttribute('data-fui-tree-item-value') }, title: norm(raw(row.querySelector(selectors.chatTitle)) || row.getAttribute('aria-label') || ''), threadId: row.getAttribute('data-fui-tree-item-value')?.split('|').at(-1), kind: 'channel' })).filter(n => n.title && n.threadId?.endsWith('@thread.tacv2')), coverage: 'rendered_channel_sidebar' };
    }
    if (action === 'openChannel') {
      if (all('composer').some(n => raw(n))) throw fail('draft_exists', 'Clear the existing draft before changing channels.');
      const rows = all('channelRow', one('chatList')).filter(n => n.getAttribute(args.key.attribute) === args.key.value);
      if (rows.length !== 1 || rows[0].getAttribute('data-fui-tree-item-value')?.split('|').at(-1) !== args.threadId) throw fail('channel_unavailable', 'Channel is absent or changed; refresh its ID.', 404);
      rows[0].click();
      if (!await wait(() => pane())) throw fail('channel_layout_unsupported', 'This channel did not expose a verified Threads layout and composer.');
      return { opened: true };
    }
    if (action === 'threads') {
      const root = assertContext(); return { items: data(root).map(message => ({ ...message, parentMessageId: message.id })), coverage: 'rendered_channel_threads' };
    }
    if (action === 'openThread') {
      // Start in the channel timeline; never use a display name or ordinal.
      const parent = args.parentMessageId;
      const area = one('mainPane'), root = one('paneLayout', area);
      if (sender(root).getAttribute('data-track-thread-id') !== args.threadId || norm(raw(one('channelHeader'))) !== args.title) throw fail('channel_changed', 'Wrong channel.');
      if (all('composer').some(n => raw(n))) throw fail('draft_exists', 'Clear the existing draft before opening a thread.');
      const matches = all('message', one('messageList', root)).filter(n => mid(n) === parent);
      if (matches.length !== 1) throw fail('message_unavailable', 'Exact parent message is outside the rendered timeline.', 404);
      const wrapper = matches[0].closest(selectors.messageWrapper) || matches[0].parentElement;
      const buttons = nodes(selectors.replyButton, wrapper).filter(n => /^\d+ repl(?:y|ies)\b/i.test(raw(n)) || /^(Reply in thread|Reply)$/i.test(n.getAttribute('aria-label') || raw(n)));
      if (!buttons.length) {
        const old = window[threadKey]; if (old) document.removeEventListener('contextmenu', old.guard, true);
        const p = { node: matches[0], token: args.token, parent, threadId: args.threadId, opened: false };
        p.guard = event => {
          if (window[threadKey] !== p || !p.node.contains(event.target) || !p.node.isConnected || sender(root).getAttribute('data-track-thread-id') !== p.threadId) { p.invalidated = true; event.preventDefault(); event.stopImmediatePropagation(); }
          else p.opened = true;
        };
        window[threadKey] = p; document.addEventListener('contextmenu', p.guard, true);
        p.node.scrollIntoView({ block: 'center' }); const r = p.node.getBoundingClientRect(), x = r.x + r.width / 2, y = r.y + r.height / 2;
        if (!p.node.contains(document.elementFromPoint(x, y))) { document.removeEventListener('contextmenu', p.guard, true); delete window[threadKey]; throw fail('thread_obscured', 'Parent message is obscured.'); }
        return { requiresContextMenu: true, x, y };
      }
      if (buttons.length !== 1) throw fail('thread_reply_unavailable', 'No unique native reply control for this parent.');
      buttons[0].click();
      if (!await wait(() => pane())) throw fail('thread_not_ready', 'The reply pane did not expose the exact channel and parent IDs.');
      return { opened: true, parentMessageId: parent, channelId: args.threadId };
    }
    if (action === 'threadMenu') {
      const p = window[threadKey];
      try {
        if (!p || p.token !== args.token || !p.opened || p.invalidated || !p.node.isConnected || mid(p.node) !== args.parentMessageId || p.threadId !== args.threadId || sender(one('paneLayout', one('mainPane'))).getAttribute('data-track-thread-id') !== args.threadId) throw fail('thread_changed', 'The exact parent menu could not be verified.');
        if (!await wait(() => all('messageMenu').length === 1, 3000)) throw fail('thread_reply_unavailable', 'Native message menu did not open.');
        const button = one('threadedReply', one('messageMenu'));
        if (button.disabled || button.getAttribute('aria-disabled') === 'true') throw fail('thread_reply_unavailable', 'Replies are disabled for this thread.');
        button.click();
        if (!await wait(() => pane())) throw fail('thread_not_ready', 'Reply composer did not expose the exact channel and parent IDs.');
        return { opened: true, parentMessageId: args.parentMessageId, channelId: args.threadId };
      } finally { if (p) document.removeEventListener('contextmenu', p.guard, true); delete window[threadKey]; }
    }
    if (action === 'channelSearchContext') {
      const query = one('searchInput').value.trim(); if (query !== args.query && query !== `${args.query} is:Messages`) throw fail('search_changed', 'The native search query changed.');
      const headers = all('channelHeader'); if (headers.length !== 1) throw fail('search_context_unavailable', 'No channel context opened.');
      const sends = all('send').filter(n => n.getAttribute('data-track-child-thread-id'));
      if (sends.length !== 1) throw fail('search_context_unavailable', 'Search did not expose a unique native channel thread.');
      const chat = { kind: 'channel', title: norm(raw(headers[0])), threadId: sends[0].getAttribute('data-track-thread-id'), parentMessageId: sends[0].getAttribute('data-track-child-thread-id'), key: { attribute: 'search', value: JSON.stringify({ query: args.query, resultKey: args.resultKey }) } };
      const root = sends[0].closest(selectors.paneLayout); if (!root) throw fail('search_context_unavailable', 'No native channel reply pane.');
      window[Symbol.for('switchboard.teams.channelSearchContext')] = { query: args.query, resultKey: args.resultKey, chat, send: sends[0], composer: editor(root) };
      return chat;
    }
    if (action === 'messages') return { items: data(assertContext()) };
    if (action === 'older') {
      const root = assertContext(), list = one('messageList', root); let scroll = list;
      while (scroll && scroll !== document.body && !(scroll.clientHeight > 0 && /auto|scroll/.test(getComputedStyle(scroll).overflowY))) scroll = scroll.parentElement;
      if (!scroll || scroll === document.body) throw fail('history_unavailable', 'No verified history scroller.');
      const before = JSON.stringify(data(root).map(n => [n.id, n.text]));
      scroll.scrollTop = Math.max(0, scroll.scrollTop - Math.max(100, scroll.clientHeight * .8));
      const changed = await wait(() => JSON.stringify(data(assertContext()).map(n => [n.id, n.text])) !== before, 1800, 100);
      return { items: data(assertContext()), changed: !!changed, atStart: scroll.scrollTop <= 1, loading: all('historyLoading', root).length > 0 };
    }
    if (action === 'prepare') {
      const root = assertContext(), e = editor(root), send = sender(root);
      if (raw(e) || all('composer').some(n => n !== e && raw(n))) throw fail('draft_exists', 'An existing message or inline draft must be cleared manually.');
      if (e.getAttribute('aria-disabled') === 'true') throw fail('composer_disabled', 'Replies are disabled.');
      const p = { token: args.token, editor: e, send, beforeIds: data(root).map(n => n.id), inputGuards: [], mentions: [] };
      if (p.beforeIds.some(id => !id)) throw fail('message_ids_unavailable', 'Stable message IDs are required.');
      window[stateKey] = p; e.focus(); return { beforeIds: p.beforeIds };
    }
    if (action === 'adoptEdit') {
      const existing = window[Symbol.for('switchboard.teams.messageAction')];
      if (!existing || existing.token !== args.token || existing.messageId !== args.messageId || existing.kind !== 'edit' || existing.invalidated || !existing.wrapper.contains(existing.editor) || norm(raw(existing.editor)) !== norm(args.expectedText)) throw fail('edit_changed', 'No verified original inline editor.');
      window[stateKey] = { token: args.token, editor: existing.editor, send: existing.save, wrapper: existing.wrapper, edit: true, inputGuards: [], mentions: [], beforeIds: [] };
      existing.editor.focus(); return { adopted: true };
    }
    if (action === 'paste') {
      const p = proof(); selectedRange(p);
      if (args.replace) { const r = document.createRange(); r.selectNodeContents(p.editor); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); }
      const escape = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      let ordinal = 0; p.placeholders = [];
      const runs = list => list.map(run => {
        if (run.mention) { const marker = `SBMENTION${args.token.replace(/-/g, '')}N${ordinal++}END`; p.placeholders.push({ marker, ...run.mention }); return marker; }
        let html = escape(run.text).replace(/\n/g, '<br>');
        const tags = { bold: 'strong', italic: 'em', underline: 'u', strike: 's', code: 'code' };
        for (const mark of run.marks) html = `<${tags[mark]}>${html}</${tags[mark]}>`;
        if (run.link) html = `<a href="${escape(run.link)}">${html}</a>`; return html;
      }).join('');
      const html = args.content.map(b => b.runs ? b.type === 'quote' ? `<blockquote><p>${runs(b.runs)}</p></blockquote>` : `<p>${runs(b.runs)}</p>` : `<${b.type === 'bulletedList' ? 'ul' : 'ol'}>${b.items.map(list => `<li>${runs(list)}</li>`).join('')}</${b.type === 'bulletedList' ? 'ul' : 'ol'}>`).join('');
      const dt = new DataTransfer(); dt.setData('text/html', html); dt.setData('text/plain', html.replace(/<[^>]*>/g, ''));
      p.editor.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
      if (!await wait(() => p.editor.innerHTML !== '<p><br></p>' && raw(p.editor), 2500)) throw fail('rich_paste_unsupported', 'The native editor did not accept rich content.');
      return { mentions: p.placeholders.map(({ email, name }, index) => ({ email, name, index })) };
    }
    if (action === 'mentionFocus') {
      const p = proof(), target = p.placeholders[args.index]; if (!target) throw fail('invalid_mention', 'Missing mention target.');
      const walker = document.createTreeWalker(p.editor, NodeFilter.SHOW_TEXT); let node, range;
      while ((node = walker.nextNode())) { const i = node.data.indexOf(target.marker); if (i >= 0) { if (range) throw fail('mention_changed', 'Ambiguous mention placeholder.'); range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + target.marker.length); } }
      if (!range) throw fail('mention_changed', 'The mention placeholder changed.');
      p.editor.focus(); const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range); guardInput(p, true);
      return { query: '@' + target.name };
    }
    if (action === 'mentionSelect') {
      const p = proof(), target = p.placeholders[args.index]; selectedRange(p);
      const options = await wait(() => { const n = all('mentionOption').filter(n => mentionIdentity(n) === target.email); return n.length === 1 ? n : false; }, 6000, 250);
      if (!options) throw fail('mention_identity_unavailable', 'Teams did not expose a unique exact-email mention candidate. No message was sent.');
      const mri = options[0].querySelector('[data-person-mri]')?.getAttribute('data-person-mri');
      const before = all('mentionCard', p.editor); options[0].click();
      const card = await wait(() => all('mentionCard', p.editor).find(n => !before.includes(n) && norm(raw(n)).replace(/^@/, '') === target.name), 2500);
      if (!card || mri && !card.querySelector(`mention[mri="${CSS.escape(mri)}"]`)) throw fail('mention_not_inserted', 'No native mention card with the requested name and person identity appeared.');
      p.mentions.push({ node: card, email: target.email, name: target.name, mri }); return { inserted: true };
    }
    if (action === 'inputFocus') { const p = proof(); p.editor.focus(); guardInput(p); return { focused: true }; }
    if (action === 'checkContent' || action === 'commit' || action === 'editCommit') {
      const p = proof();
      if (p.placeholders?.some(x => raw(p.editor).includes(x.marker)) || p.mentions.some(m => !m.node.isConnected || !p.editor.contains(m.node) || norm(raw(m.node)).replace(/^@/, '') !== m.name || m.mri && !m.node.querySelector(`mention[mri="${CSS.escape(m.mri)}"]`))) throw fail('mention_changed', 'The native mention draft changed.');
      if (nativeText(p.editor) !== norm(args.text)) throw fail('draft_mismatch', 'The editor text differs from the requested content.');
      // Snapshot native markup, then require identical markup at the final click.
      if (action === 'checkContent') {
        const tags = { bold: 'strong,b', italic: 'em,i', underline: 'u,span[style*="underline"]', strike: 's,strike,span[style*="line-through"]', code: 'code' };
        for (const b of args.content ?? []) {
          if (b.type === 'quote' && !p.editor.querySelector('blockquote')) throw fail('format_not_preserved', 'Quote formatting was lost.');
          if (b.type === 'bulletedList' && !p.editor.querySelector('ul') || b.type === 'numberedList' && !p.editor.querySelector('ol')) throw fail('format_not_preserved', 'List formatting was lost.');
          for (const run of b.runs ?? b.items.flat()) {
            for (const mark of run.marks ?? []) if (!nodes(tags[mark], p.editor).some(n => raw(n).includes(run.text.trim()))) throw fail('format_not_preserved', `Native editor did not preserve ${mark}.`);
            if (run.link && !nodes('a[href]', p.editor).some(n => n.href === run.link && raw(n).includes(run.text.trim()))) throw fail('format_not_preserved', 'Native editor did not preserve the link.');
          }
        }
        p.html = p.editor.innerHTML; return { verified: true };
      }
      if (p.html !== p.editor.innerHTML || p.send.disabled || p.send.getAttribute('aria-disabled') === 'true') throw fail('draft_changed', 'The draft or Send control changed.');
      // Programmatic click is synchronous with all identity and draft checks.
      if (args.key?.attribute === 'recipient') window[Symbol.for('switchboard.teams.recipient')].clicked = true;
      p.send.click();
      if (p.edit) {
        const edited = await wait(() => !p.editor.isConnected && matchingMessage(assertContext(), n => n.id === args.messageId), 10_000, 300);
        if (!edited) throw fail('mutation_uncertain', 'Edit was attempted without verification; inspect Teams manually.', 504);
        return { status: 'edited_in_ui', message: edited, mentionCount: p.mentions.length };
      }
      const observed = await wait(() => (!p.editor.isConnected || !norm(raw(p.editor))) && matchingMessage(assertContext(), n => n.id && !p.beforeIds.includes(n.id)), 10_000, 300);
      if (!observed) throw fail('send_uncertain', 'Send was attempted without a verified new message; inspect Teams manually.', 504);
      return { status: args.kind === 'channel' ? 'observed_in_channel_thread' : 'observed_in_chat', message: observed, ...(args.kind === 'channel' ? { channelId: args.threadId, parentMessageId: args.parentMessageId ?? null } : {}), mentionCount: p.mentions.length };
    }
    if (action === 'release') { const p = window[stateKey]; if (p?.token === args.token) { for (const guard of p.inputGuards) document.removeEventListener('beforeinput', guard, true); delete window[stateKey]; } return { released: true }; }
    return fail('unknown_action', 'Unknown extended DOM action.', 400);
  } catch (e) { return e?.error ? e : fail('selector_mismatch', 'Native channel or rich-editor layout could not be verified.'); }
}

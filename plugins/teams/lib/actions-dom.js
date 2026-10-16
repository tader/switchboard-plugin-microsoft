export const ACTION_SELECTORS = Object.freeze({
  messageWrapper: '[data-tid="chat-pane-item"]',
  ownMessage: '.fui-ChatMyMessage',
  messageMenu: '[data-tid="message-actions-menu-renderer-getV9MessageActions"]',
  editMenu: '[data-tid="message-actions-edit"]',
  deleteMenu: '[data-tid="message-actions-delete"]',
  editSave: '[data-tid="newMessageCommands-send"], [data-tid="edit-message-save"]',
  deletedMessage: '[data-tid="message-deleted"], [data-tid="deleted-message"], [data-tid="deleted-message-placeholder"], [data-tid="message-deleted-placeholder"], [data-tid="deleted-message-renderer"]',
  attachment: '[data-tid="message-attachments"], [data-tid="attachment-list"], [data-tid="attachment-card"]',
  reactionToolbar: '[data-tid="message-reactions-toolbar"], [data-tid="message-reactions-on-menu"]',
  reactionButton: 'button[data-tid^="message-actions-"]',
  reactionPill: '[data-tid="diverse-reaction-pill-button"]',
  reactionEmoji: '[itemtype="http://schema.skype.com/Emoji"][itemid]',
  unreadMenu: '[data-tid="change-unread-status-menu-item"]',
});

// A separate compact DOM evaluation for message mutations. No private API calls,
// React internals, screenshots, or keyboard shortcuts that might send a draft.
export async function teamsActions(action, args, selectors) {
  const fail = (code, message, status = 409) => ({ error: { code, message, status } });
  const visible = n => !!n && n.getClientRects().length > 0 && getComputedStyle(n).visibility !== 'hidden';
  const all = (key, root = document) => [...root.querySelectorAll(selectors[key])].filter(visible);
  const one = (key, root = document) => {
    const nodes = all(key, root);
    if (nodes.length !== 1) throw fail('selector_mismatch', `Expected one ${key}; found ${nodes.length}. Run diagnostics.`);
    return nodes[0];
  };
  const raw = n => (n?.innerText ?? n?.textContent ?? '').trim();
  const normalized = value => value.replace(/\s+/g, ' ').trim();
  const text = n => normalized(raw(n));
  const id = n => n.getAttribute('data-mid') || n.getAttribute('data-message-id') || n.getAttribute('data-itemid') || n.id || null;
  const own = n => !!n.closest(selectors.ownMessage);
  const stateKey = Symbol.for('switchboard.teams.messageAction');
  const canonical = value => value === 'yes' ? 'like' : value;
  const wait = (predicate, milliseconds = 8000, stable = 150) => new Promise(resolve => {
    let timer, done = false;
    const finish = value => { if (done) return; done = true; observer.disconnect(); clearTimeout(timer); clearTimeout(deadline); resolve(value); };
    const check = () => {
      clearTimeout(timer);
      try { if (predicate()) timer = setTimeout(() => { try { const value = predicate(); if (value) finish(value); } catch {} }, stable); } catch {}
    };
    const observer = new MutationObserver(check);
    observer.observe(document.documentElement, { subtree: true, childList: true, attributes: true, characterData: true });
    const deadline = setTimeout(() => finish(null), milliseconds); check();
  });
  const assertChat = () => {
    if (text(one('header')) !== args.title || !args.threadId) throw fail('chat_changed', 'A verified chat thread is required for message changes.');
    const list = one('messageList');
    const normalSends = all('send').filter(n => !list.contains(n));
    if (normalSends.length !== 1 || normalSends[0].getAttribute('data-track-thread-id') !== args.threadId) throw fail('chat_changed', 'The selected chat thread changed.');
    return list;
  };
  const message = () => {
    const matches = all('message', assertChat()).filter(n => id(n) === args.messageId);
    if (matches.length !== 1) throw fail('message_unavailable', 'The exact message ID is absent or ambiguous in the loaded history.', 404);
    return matches[0];
  };
  const noInlineEdit = () => {
    const list = assertChat();
    if (all('composer', list).length) throw fail('edit_exists', 'An inline message edit is already open. Finish or cancel it manually.');
  };
  const proof = () => {
    const value = window[stateKey];
    if (!value || value.token !== args.token || value.threadId !== args.threadId || value.messageId !== args.messageId || value.kind !== args.kind) throw fail('action_changed', 'The prepared message action changed.');
    return value;
  };
  const checkMessage = (p, editing = false) => {
    assertChat();
    if (p.invalidated) throw fail('action_changed', 'Another message menu was opened during this operation.');
    if (!p.wrapper.isConnected) throw fail('message_changed', 'The target message wrapper changed.');
    if (editing) {
      if (!p.editor?.isConnected || !p.wrapper.contains(p.editor) || one('composer', p.wrapper) !== p.editor || one('editSave', p.wrapper) !== p.save) throw fail('edit_changed', 'The inline editor changed.');
    } else {
      const node = message();
      if (node !== p.node || raw(node.querySelector(selectors.messageBody)) !== p.originalText) throw fail('message_changed', 'The target message content or identity changed.');
    }
    if (['edit', 'delete'].includes(p.kind) && !p.wrapper.querySelector(selectors.ownMessage) && !p.wrapper.matches(selectors.ownMessage)) throw fail('not_own_message', 'Only your own messages can be edited or deleted.', 403);
  };
  const point = node => {
    node.scrollIntoView({ block: 'center' });
    const rect = node.getBoundingClientRect(), x = rect.x + rect.width / 2, y = rect.y + rect.height / 2;
    if (!node.contains(document.elementFromPoint(x, y))) throw fail('action_obscured', 'The target control is obscured.');
    return { x, y };
  };
  const toolbar = () => {
    const p = window[stateKey];
    if (p?.token === args.token && p.menuOpened && !p.invalidated) {
      const menus = all('messageMenu');
      if (menus.length === 1) return all('reactionToolbar', menus[0])[0] || null;
    }
    const surface = document.getElementById(`${args.messageId}-popover-surface`);
    return visible(surface) ? all('reactionToolbar', surface)[0] : null;
  };
  const buttons = () => {
    const root = toolbar(); if (!root) return [];
    return all('reactionButton', root).map(node => {
      const name = node.querySelector(selectors.reactionEmoji)?.getAttribute('itemid') || node.getAttribute('data-tid').replace(/^message-actions-/, '');
      return { node, reaction: canonical(name), selected: node.getAttribute('aria-pressed') === 'true', stateKnown: ['true', 'false'].includes(node.getAttribute('aria-pressed')), label: node.getAttribute('aria-label') || null };
    }).filter(n => /^[a-z][a-z0-9_-]{0,63}$/.test(n.reaction));
  };
  const pills = wrapper => all('reactionPill', wrapper).map(node => {
    const emojiId = node.querySelector(selectors.reactionEmoji)?.getAttribute('itemid');
    const label = (node.getAttribute('aria-labelledby') ?? '').split(/\s+/).map(value => document.getElementById(value)?.textContent ?? '').join(' ');
    const count = Number.parseInt(label, 10);
    return { node, reaction: canonical(emojiId), emojiId, selected: node.getAttribute('aria-pressed') === 'true', stateKnown: ['true', 'false'].includes(node.getAttribute('aria-pressed')), count: Number.isFinite(count) ? count : null };
  }).filter(n => n.emojiId);
  const reactionState = p => {
    const current = buttons().filter(b => b.reaction === args.reaction && b.stateKnown);
    if (current.length === 1) return current[0].selected;
    const existing = pills(p.wrapper).filter(b => b.reaction === args.reaction && b.stateKnown);
    if (existing.length === 1) return existing[0].selected;
    return null;
  };
  const rowKey = row => {
    for (const attribute of ['data-fui-tree-item-value', 'data-chat-id', 'data-itemid', 'id']) {
      const value = row.getAttribute(attribute); if (value) return { attribute, value };
    }
    return { attribute: 'title', value: text(row.querySelector(selectors.chatTitle)) };
  };
  const rowThread = row => row.getAttribute('data-fui-tree-item-value')?.split('|').at(-1) || row.getAttribute('data-chat-id');
  const unread = row => row.getAttribute('data-is-unread') === 'true' || (row.getAttribute('aria-labelledby') ?? '').split(/\s+/).includes('chat_list_unread_text') || !!row.querySelector('[data-tid="unread-badge"], [data-testid="chat-list-unread-text"]');
  const chatRow = () => {
    const matches = all('chatRow', one('chatList')).filter(row => args.key?.attribute === 'recipient' || args.key?.attribute === 'search' ? rowThread(row) === args.threadId : JSON.stringify(rowKey(row)) === JSON.stringify(args.key));
    if (matches.length !== 1) throw fail('chat_unavailable', 'Chat is absent or ambiguous in the sidebar.', 404);
    if (!args.threadId || rowThread(matches[0]) !== args.threadId || text(matches[0].querySelector(selectors.chatTitle)) !== args.title) throw fail('chat_changed', 'The sidebar chat identity changed.');
    return matches[0];
  };
  const contextGuard = (p, target, validate) => {
    const guardKey = Symbol.for('switchboard.teams.actionContextGuard');
    if (window[guardKey]) document.removeEventListener('contextmenu', window[guardKey], true);
    const guard = event => {
      try { if (window[stateKey] !== p || !target.contains(event.target)) throw new Error(); validate(); p.menuOpened = true; }
      catch { p.invalidated = true; event.preventDefault(); event.stopImmediatePropagation(); }
    };
    window[guardKey] = guard; document.addEventListener('contextmenu', guard, true);
    const clickKey = Symbol.for('switchboard.teams.actionMenuClickGuard');
    if (window[clickKey]) for (const type of ['pointerdown', 'click']) document.removeEventListener(type, window[clickKey], true);
    const outside = event => {
      if (window[stateKey] !== p) return;
      const surface = event.target.closest?.('[data-tid="message-actions-container"]');
      if (![p.node, p.wrapper, p.row, p.menu, p.menuButton].some(n => n?.contains(event.target)) && surface?.id !== `${p.messageId}-popover-surface`) p.invalidated = true;
    };
    window[clickKey] = outside;
    for (const type of ['pointerdown', 'click']) document.addEventListener(type, outside, true);
  };
  try {
    for (const selector of Object.values(selectors)) document.querySelector(selector);
    if (action === 'messageTarget') {
      noInlineEdit(); const node = message(), wrapper = node.closest(selectors.messageWrapper);
      if (!wrapper) throw fail('message_identity_unavailable', 'Could not identify the native message wrapper.');
      const originalText = raw(node.querySelector(selectors.messageBody));
      if (['edit', 'delete'].includes(args.kind)) {
        if (!own(node)) throw fail('not_own_message', 'Only your own messages can be edited or deleted.', 403);
        if (originalText !== args.expectedText) throw fail('message_changed', 'expectedText does not match the current message. Read it again.');
      }
      if (args.kind === 'edit') {
        if (all('composer').some(n => raw(n) || n.querySelector(selectors.quoteCard))) throw fail('draft_exists', 'Clear or send the existing draft before editing.');
        const body = node.querySelector(selectors.messageBody);
        if (!body || (!args.content && body.querySelector('a,img,card,table,pre,blockquote,[contenteditable="false"]')) || body.querySelector(selectors.quoteRendered) || wrapper.querySelector(selectors.attachment)) throw fail('edit_content_unsupported', 'Use structured content for rich editing. Quotes and attachments cannot be replaced.');
      }
      const p = { token: args.token, kind: args.kind, threadId: args.threadId, messageId: args.messageId, node, wrapper, originalText };
      window[stateKey] = p; const location = point(node);
      if (['edit', 'delete', 'reaction', 'reactions'].includes(args.kind)) contextGuard(p, node, () => checkMessage(p));
      return location;
    }
    if (action === 'reactionReady') {
      const p = proof(); checkMessage(p);
      if (!await wait(() => toolbar() && buttons().length > 0, 3000)) throw fail('reaction_unavailable', 'Could not identify the target message’s reaction toolbar.');
      checkMessage(p);
      if (p.menuOpened) p.menu = one('messageMenu');
      const available = buttons(); const current = pills(p.wrapper);
      return { available: available.map(({ node, stateKnown, ...item }) => item), items: current.map(({ node, stateKnown, ...item }) => item), coverage: 'visible_native_reactions' };
    }
    if (action === 'reactionPrepare') {
      const p = proof(); checkMessage(p);
      const state = reactionState(p);
      let matches = buttons().filter(n => n.reaction === args.reaction && n.stateKnown);
      p.reactionSource = 'toolbar';
      if (!matches.length) { p.reactionSource = 'pill'; matches = pills(p.wrapper).filter(n => n.reaction === args.reaction && n.stateKnown); }
      if (matches.length !== 1 || state === null) throw fail('reaction_unavailable', 'Use a reaction ID from the target’s available reactions or existing visible pills.');
      p.reactionNode = matches[0].node; p.reaction = args.reaction; p.desired = args.selected; p.beforeSelected = state;
      return { alreadySelected: state === args.selected };
    }
    if (action === 'reactionCommit') {
      const p = proof(); checkMessage(p);
      if (p.menuOpened && one('messageMenu') !== p.menu) throw fail('reaction_changed', 'The verified reaction menu changed.');
      if (p.reaction !== args.reaction || p.desired !== args.selected) throw fail('action_changed', 'The requested reaction state changed.');
      const state = reactionState(p);
      if (state === args.selected) return { status: 'reaction_observed', messageId: args.messageId, reaction: args.reaction, selected: args.selected, changed: false };
      const controls = (p.reactionSource === 'toolbar' ? buttons() : pills(p.wrapper)).filter(n => n.reaction === args.reaction && n.stateKnown);
      if (controls.length !== 1 || controls[0].node !== p.reactionNode || state !== p.beforeSelected || !visible(p.reactionNode) || p.reactionNode.disabled) throw fail('reaction_changed', 'The reaction or its control changed before clicking.');
      p.reactionNode.click();
      const observed = await wait(() => { checkMessage(p); return reactionState(p) === args.selected; });
      if (!observed) throw fail('mutation_uncertain', 'Reaction was attempted but not confirmed. Inspect Teams; reuse the same key.', 504);
      return { status: 'reaction_observed', messageId: args.messageId, reaction: args.reaction, selected: args.selected, changed: true };
    }
    if (action === 'menuPrepare') {
      const p = proof(); checkMessage(p);
      if (!p.menuOpened || !await wait(() => all('messageMenu').length === 1, 3000)) throw fail('message_menu_unavailable', 'The verified message menu did not open.');
      const menu = one('messageMenu'), button = one(args.kind === 'edit' ? 'editMenu' : 'deleteMenu', menu);
      if (button.disabled || button.getAttribute('aria-disabled') === 'true') throw fail('action_disabled', 'Teams does not allow this message action.');
      p.menu = menu; p.menuButton = button; return { ready: true };
    }
    if (action === 'editPrepare') {
      const p = proof(); checkMessage(p);
      if (!p.menuOpened || one('messageMenu') !== p.menu || one('editMenu', p.menu) !== p.menuButton) throw fail('action_changed', 'The edit menu changed.');
      p.menuButton.click();
      if (!await wait(() => p.wrapper.isConnected && all('composer', p.wrapper).length === 1 && all('editSave', p.wrapper).length === 1, 5000)) throw fail('edit_unavailable', 'The inline edit did not open. Inspect Teams manually.');
      p.editor = one('composer', p.wrapper); p.save = one('editSave', p.wrapper);
      checkMessage(p, true);
      const clean = value => args.content ? normalized(value.replace(/[\u200b-\u200d\u2060\ufeff]/g, '')) : normalized(value);
      if (clean(raw(p.editor)) !== clean(args.expectedText) || p.editor.querySelector(selectors.quoteCard)) throw fail('edit_changed', 'The inline editor does not contain the expected original text.');
      return { ready: true };
    }
    if (action === 'editSelect') {
      const p = proof(); checkMessage(p, true);
      if (text(p.editor) !== normalized(args.expectedText)) throw fail('edit_changed', 'The edit draft changed. No text was replaced.');
      p.editor.focus(); const range = document.createRange(); range.selectNodeContents(p.editor);
      const selection = getSelection(); selection.removeAllRanges(); selection.addRange(range);
      const guardKey = Symbol.for('switchboard.teams.editInputGuard');
      if (window[guardKey]) document.removeEventListener('beforeinput', window[guardKey], true);
      const guard = event => {
        try { checkMessage(p, true); if (window[stateKey] !== p || event.target !== p.editor || document.activeElement !== p.editor || text(p.editor) !== normalized(args.expectedText)) throw new Error(); p.inputObserved = true; }
        catch { event.preventDefault(); event.stopImmediatePropagation(); }
        document.removeEventListener('beforeinput', guard, true); if (window[guardKey] === guard) delete window[guardKey];
      };
      window[guardKey] = guard; document.addEventListener('beforeinput', guard, true); return { selected: true };
    }
    if (action === 'editCommit') {
      const p = proof(); checkMessage(p, true);
      if (!p.inputObserved || text(p.editor) !== normalized(args.text) || p.save.disabled || p.save.getAttribute('aria-disabled') === 'true') throw fail('edit_changed', 'The replacement text or Done control changed. No save click was made.');
      p.save.click();
      if (!await wait(() => { const node = message(); return !all('composer', p.wrapper).length && normalized(raw(node.querySelector(selectors.messageBody))) === normalized(args.text) && !node.querySelector(selectors.pending); })) throw fail('mutation_uncertain', 'Edit was attempted but not confirmed. Inspect Teams; reuse the same key.', 504);
      return { status: 'edited_in_ui', messageId: args.messageId, text: args.text };
    }
    if (action === 'deleteCommit') {
      const p = proof(); checkMessage(p);
      if (!p.menuOpened || one('messageMenu') !== p.menu || one('deleteMenu', p.menu) !== p.menuButton) throw fail('action_changed', 'The delete menu changed.');
      p.menuButton.click();
      if (!await wait(() => {
        assertChat(); if (!p.wrapper.isConnected) return false;
        if (all('deletedMessage', p.wrapper).length) return true;
        const tombstone = /(?:You deleted this message|This message (?:has been|was) deleted)\.?/i.test(text(p.wrapper));
        const undo = [...p.wrapper.querySelectorAll('button,[role="button"]')].some(n => visible(n) && /^(Undo|Undo delete|Undo deletion)$/i.test(n.getAttribute('aria-label') || text(n)));
        return tombstone && undo;
      })) throw fail('mutation_uncertain', 'Deletion was attempted but no native tombstone was verified. Inspect Teams; reuse the same key.', 504);
      return { status: 'deleted_in_ui', messageId: args.messageId };
    }
    if (action === 'unreadTarget') {
      const row = chatRow(), p = { token: args.token, kind: 'unread', threadId: args.threadId, row, alreadyUnread: unread(row) };
      window[stateKey] = p;
      if (p.alreadyUnread) return { alreadyUnread: true };
      const location = point(row); contextGuard(p, row, () => { if (chatRow() !== row) throw new Error(); });
      return { ...location, alreadyUnread: false };
    }
    if (action === 'unreadPrepare') {
      const p = window[stateKey];
      if (!p || p.invalidated || p.token !== args.token || p.row !== chatRow()) throw fail('action_changed', 'The sidebar action changed.');
      if (unread(p.row)) return { alreadyUnread: true };
      if (!p.menuOpened || !await wait(() => all('unreadMenu').length === 1, 3000)) throw fail('unread_action_unavailable', 'The verified chat’s unread menu did not open.');
      const button = one('unreadMenu');
      if (!/\bMark as unread\b/i.test(text(button)) || button.disabled || button.getAttribute('aria-disabled') === 'true') throw fail('unread_action_unavailable', 'The menu does not offer Mark as unread.');
      p.menuButton = button; return { alreadyUnread: false };
    }
    if (action === 'unreadCommit') {
      const p = window[stateKey];
      if (!p || p.invalidated || p.token !== args.token || p.row !== chatRow()) throw fail('action_changed', 'The sidebar action changed.');
      if (unread(p.row)) return { status: 'unread_observed', changed: false };
      if (!p.menuOpened || one('unreadMenu') !== p.menuButton || !/\bMark as unread\b/i.test(text(p.menuButton))) throw fail('unread_action_unavailable', 'The unread menu changed. No toggle was clicked.');
      p.menuButton.click();
      if (!await wait(() => unread(chatRow()))) throw fail('mutation_uncertain', 'Mark unread was attempted but not confirmed in the sidebar. Inspect Teams; reuse the same key.', 504);
      return { status: 'unread_observed', changed: true };
    }
    if (action === 'release') {
      if (window[stateKey]?.token === args.token) {
        const contextKey = Symbol.for('switchboard.teams.actionContextGuard');
        if (window[contextKey]) { document.removeEventListener('contextmenu', window[contextKey], true); delete window[contextKey]; }
        const clickKey = Symbol.for('switchboard.teams.actionMenuClickGuard');
        if (window[clickKey]) { for (const type of ['pointerdown', 'click']) document.removeEventListener(type, window[clickKey], true); delete window[clickKey]; }
        const inputKey = Symbol.for('switchboard.teams.editInputGuard');
        if (window[inputKey]) { document.removeEventListener('beforeinput', window[inputKey], true); delete window[inputKey]; }
        delete window[stateKey];
      }
      return { released: true };
    }
    throw fail('unknown_action', 'Unknown message action.', 400);
  } catch (error) { return error?.error ? error : fail('selector_mismatch', 'Message-action selectors do not match this Teams layout. Run diagnostics.'); }
}

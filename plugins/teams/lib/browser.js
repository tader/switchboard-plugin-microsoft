import fs from 'node:fs/promises';
import path from 'node:path';
import { CDP, findBrowser } from './cdp.js';
import { teamsActions } from './actions-dom.js';
import { teamsExtended } from './extended-dom.js';
import { richContent, contentText } from './content.js';
import { teamsDOM, selectorsFrom } from './dom.js';
import { AdapterError } from './errors.js';
import { randomUUID } from 'node:crypto';
import { recipientEmail, peopleQuery, messageQuery, messageId, messageText, reactionInput } from './input.js';

const TEAMS_URL = 'https://teams.microsoft.com/v2/';
export class TeamsBrowser {
  constructor(profileDir, settings = {}, options = {}) {
    this.profileDir = profileDir;
    this.settings = settings;
    this.options = options;
    this.selectors = selectorsFrom(settings.selectors);
    this.queue = Promise.resolve();
    this.queued = 0;
    this.disposed = false;
    this.historyCursors = new Map();
  }
  serial(fn) {
    if (this.disposed) return Promise.reject(new AdapterError('disposed', 'Browser is being disposed.', 503));
    if (this.queued >= 20) return Promise.reject(new AdapterError('queue_full', 'Too many queued Teams operations. Try later.', 429));
    this.queued++;
    const result = this.queue.then(() => {
      if (this.disposed) throw new AdapterError('disposed', 'Browser has been disposed.', 503);
      return fn();
    }).finally(() => this.queued--);
    this.queue = result.catch(() => {});
    return result;
  }
  async start() {
    if (this.cdp && !this.cdp.closed) return;
    await fs.mkdir(this.profileDir, { recursive: true, mode: 0o700 });
    await fs.chmod(this.profileDir, 0o700);
    this.cdp = new CDP(await findBrowser(this.settings.browserPath), this.profileDir, this.options);
    try { this.page = await this.cdp.page(TEAMS_URL); }
    catch (error) { await this.cdp.close(); throw error; }
  }
  async dom(action, args = {}) {
    await this.start();
    const location = await this.page.evaluate(() => ({ host: location.hostname, protocol: location.protocol }));
    if (location.protocol !== 'https:' || !['teams.microsoft.com', 'teams.cloud.microsoft'].includes(location.host)) {
      if (action === 'status') return { ready: false, signInRequired: true };
      if (action === 'diagnostics') return { ready: false, signInRequired: true, host: location.host };
      throw new AdapterError('signin_required', 'Sign in in the dedicated Teams browser using the login operation.', 503);
    }
    return this.page.evaluate(action.startsWith('action:') ? teamsActions : action.startsWith('extended:') ? teamsExtended : teamsDOM, action.replace(/^(action|extended):/, ''), args, this.selectors);
  }
  async login() {
    await this.start();
    await this.page.call('Page.bringToFront');
    return { opened: true, instructions: 'Complete sign-in in this browser, then check status. Keep this dedicated Teams window open.' };
  }
  chatId(chat) { return Buffer.from(JSON.stringify(chat)).toString('base64url'); }
  async listChats(unreadOnly = false) {
    const result = await this.dom('chats');
    return { ...result, items: result.items.filter(item => !unreadOnly || item.unread).map(({ key, threadId, ...item }) => ({ ...item, id: this.chatId({ key, title: item.title, threadId }) })) };
  }
  encodeChats(result, unreadOnly = false) {
    return { ...result, items: result.items.filter(item => !unreadOnly || item.unread).map(({ key, threadId, ...item }) => ({ ...item, id: this.chatId({ key, title: item.title, threadId }) })) };
  }
  async scanChats(unreadOnly = true, maxWindows = 100, target) {
    const token = randomUUID(), found = new Map();
    let complete = false, windows = 0, endSignature, matched, restoration;
    try {
      await this.dom('scanBegin', { token, unreadOnly });
      for (; windows < maxWindows; windows++) {
        const result = await this.dom(windows ? 'scanNext' : 'scanWindow', { token });
        for (const item of result.items) found.set(item.threadId || JSON.stringify(item.key), item);
        if (target) {
          matched = result.items.find(item => JSON.stringify(item.key) === JSON.stringify(target.key));
          if (matched) { await this.dom('open', target); break; }
        }
        const signature = JSON.stringify([result.position, result.height, [...found.keys()]]);
        // Observe a stable terminal window twice: lazy loading may append rows.
        if (result.atEnd && !result.collapsed && !result.more && signature === endSignature) { complete = true; windows++; break; }
        endSignature = result.atEnd ? signature : null;
      }
    } finally { restoration = await this.dom('scanEnd', { token }).catch(error => { if (error.code !== 'scan_changed') throw error; }); }
    if (target && !matched) throw new AdapterError('chat_unavailable', 'Chat was not found during sidebar discovery. Refresh chat IDs.', 404);
    return this.encodeChats({ items: [...found.values()], discoveryComplete: complete, truncated: !complete, windowsScanned: windows, coverage: 'native_unread_chat_list', completeAccount: false, uiRestored: restoration?.restored ?? false, unrestoredSections: restoration?.unrestoredSections ?? 0, excludes: ['channels', 'other_tenants', 'hidden_chats'] }, unreadOnly);
  }
  async searchMessages(query, limit = 50, maxPages = 3) {
    query = messageQuery(query);
    const focus = await this.dom('searchFocus');
    if (!focus.focused) throw new AdapterError('search_focus_changed', 'Could not focus the Teams search input.');
    await this.page.call('Input.insertText', { text: query });
    await this.dom('searchSubmit', { query });
    await this.page.call('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r', unmodifiedText: '\r' });
    await this.page.call('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
    const items = new Map(); let hasNext = false, pages = 0;
    for (; pages < maxPages; pages++) {
      const result = await this.dom('searchResults', { query }); hasNext = result.hasNext;
      for (const { key, ...item } of result.items) items.set(key, { ...item, resultId: this.chatId({ query, key }), messageId: null });
      if (!hasNext || items.size >= limit) { pages++; break; }
      if (pages + 1 < maxPages) await this.dom('searchNext', { query });
    }
    return { query, items: [...items.values()].slice(0, limit), pagesScanned: pages, hasMore: hasNext || items.size > limit, coverage: 'native_message_search', completeSearch: !hasNext && items.size <= limit };
  }
  decodeSearch(resultId) {
    try {
      if (typeof resultId !== 'string' || resultId.length > 4096) throw new Error();
      const value = JSON.parse(Buffer.from(resultId, 'base64url').toString());
      value.query = messageQuery(value.query);
      if (typeof value.key !== 'string' || !value.key.startsWith('serp-message-card-content-') || value.key.length > 2000) throw new Error();
      return value;
    } catch { throw new AdapterError('invalid_search_result', 'Use resultId returned by searchTeamsMessages.', 400); }
  }
  async openSearch(resultId) {
    const { query, key } = this.decodeSearch(resultId);
    // Re-run the native query, then locate this exact result key. Never open by
    // display name, ordinal, or snippet, which may change between requests.
    await this.searchMessages(query, 1, 1);
    for (let page = 0; page < 10; page++) {
      const result = await this.dom('searchResults', { query });
      if (result.items.some(item => item.key === key)) {
        const opened = await this.dom('searchOpen', { query, resultKey: key });
        return opened.channelContext ? this.dom('extended:channelSearchContext', { query, resultKey: key }) : opened;
      }
      if (!result.hasNext) break;
      await this.dom('searchNext', { query });
    }
    throw new AdapterError('search_result_unavailable', 'The selected result was not found in the first ten search pages. Search again.', 404);
  }
  async searchContext(resultId) {
    const chat = await this.openSearch(resultId);
    return { chatId: this.chatId(chat), title: chat.title, kind: chat.kind ?? 'chat', ...(chat.kind === 'channel' ? { channelId: this.chatId(chat), parentMessageId: chat.parentMessageId } : {}), ...await this.dom(chat.kind === 'channel' ? 'extended:messages' : 'messages', { ...chat, limit: 100 }), completeHistory: false, mayMarkRead: true };
  }
  async unreadMessages(maxChats = 3, limitPerChat = 50) {
    // Snapshot unread flags BEFORE opening anything: Teams marks opened chats read.
    const snapshot = await this.scanChats(true);
    const chats = [];
    for (const summary of snapshot.items.slice(0, maxChats)) {
      try {
        const chat = await this.openChat(summary.id);
        chats.push({ chatId: summary.id, title: summary.title, unreadAtSnapshot: true, ...await this.dom('unreadMessages', { ...chat, limit: limitPerChat }) });
      } catch (error) {
        chats.push({ chatId: summary.id, title: summary.title, items: [], error: { code: error.code ?? 'read_failed', message: error.message } });
      }
    }
    return {
      items: chats.flatMap(chat => chat.items.map(message => ({ chatId: chat.chatId, chatTitle: chat.title, ...message }))),
      chats, unreadChatsAtSnapshot: snapshot.items.length, remainingChats: Math.max(0, snapshot.items.length - maxChats),
      mayMarkRead: true, discoveryComplete: snapshot.discoveryComplete, completeAccount: false, coverage: 'native_unread_chat_list',
    };
  }
  async searchPeople(query) {
    query = peopleQuery(query);
    const before = await this.dom('peopleFocus');
    await this.page.call('Input.insertText', { text: query });
    return this.dom('peopleResults', { ...before, query });
  }
  async openRecipient(email) {
    email = recipientEmail(email);
    const result = await this.searchPeople(email);
    if (result.items.filter(person => person.email === email).length !== 1) throw new AdapterError('recipient_unavailable', 'No unique directory person matches that email. Search people and use a returned email.', 404);
    return this.dom('selectPerson', { email });
  }
  async conversation(email) {
    email = recipientEmail(email);
    const chat = await this.openRecipient(email);
    return { chatId: this.chatId(chat), title: chat.title, email, status: 'conversation_open', messageSent: false };
  }
  decodeChat(id) {
    try {
      if (typeof id !== 'string' || id.length > 4096) throw new Error();
      const value = JSON.parse(Buffer.from(id, 'base64url').toString());
      if (!['data-fui-tree-item-value', 'data-chat-id', 'data-itemid', 'id', 'title', 'recipient', 'search'].includes(value.key?.attribute) || typeof value.key.value !== 'string' || !value.key.value || typeof value.title !== 'string' || !value.title || (value.threadId != null && typeof value.threadId !== 'string')) throw new Error();
      if (value.key.attribute === 'recipient') value.key.value = recipientEmail(value.key.value);
      return value;
    } catch { throw new AdapterError('invalid_chat', 'Use a chat ID returned by listTeamsChats.', 400); }
  }
  async openChat(id) {
    const chat = this.decodeChat(id);
    if (chat.key.attribute === 'search') {
      let target; try { target = JSON.parse(chat.key.value); } catch { throw new AdapterError('invalid_chat', 'Invalid search conversation.', 400); }
      const opened = await this.openSearch(this.chatId({ query: target.query, key: target.resultKey }));
      if (opened.threadId !== chat.threadId) throw new AdapterError('chat_changed', 'The search result opened a different thread.');
      return opened;
    }
    if (chat.key.attribute === 'recipient') return this.openRecipient(chat.key.value);
    try { await this.dom('open', chat); }
    catch (error) { if (error.code !== 'chat_unavailable') throw error; await this.scanChats(false, 100, chat); }
    return chat;
  }
  async messages(id, limit = 50, olderPages = 0) {
    const chat = await this.openChat(id);
    const result = await this.dom('messages', { ...chat, limit: 200 });
    const items = result.items;
    for (let page = 0; page < olderPages; page++) {
      const older = await this.dom('older', chat);
      items.unshift(...older.items);
      if (!older.changed) break;
    }
    const seen = new Set();
    const unique = items.filter(m => {
      const key = m.id ?? JSON.stringify([m.text, m.author, m.timestamp]);
      if (seen.has(key)) return false;
      seen.add(key); return true;
    });
    return { chatId: id, title: chat.title, items: olderPages ? unique.slice(0, limit) : unique.slice(-limit), coverage: 'rendered_message_window', completeHistory: false };
  }
  async channels(maxWindows = 100, target) {
    const token = randomUUID(), found = new Map(); let signature, complete = false, windows = 0, restoration;
    try {
      await this.dom('scanBegin', { token, unreadOnly: false });
      for (; windows < maxWindows; windows++) {
        const window = await this.dom(windows ? 'scanNext' : 'scanWindow', { token });
        const rendered = (await this.dom('extended:channels')).items;
        for (const item of rendered) found.set(item.threadId, item);
        if (target && rendered.some(item => item.threadId === target.threadId)) { await this.dom('extended:openChannel', target); complete = true; windows++; break; }
        const next = JSON.stringify([window.position, window.height, [...found.keys()]]);
        if (window.atEnd && !window.collapsed && !window.more && next === signature) { complete = true; windows++; break; }
        signature = window.atEnd ? next : null;
      }
    } finally { restoration = await this.dom('scanEnd', { token }); }
    return { items: [...found.values()].map(item => ({ ...item, channelId: this.chatId(item) })), discoveryComplete: complete, windowsScanned: windows, uiRestored: restoration?.restored, coverage: 'exposed_channel_sidebar', completeAccount: false };
  }
  decodeChannel(id) {
    const channel = this.decodeChat(id);
    if (channel.kind !== 'channel' || !channel.threadId?.endsWith('@thread.tacv2')) throw new AdapterError('invalid_channel', 'Use channelId returned by listTeamsChannels.', 400);
    if (channel.parentMessageId != null) messageId(channel.parentMessageId);
    return channel;
  }
  async openChannel(id, parentMessageId) {
    const channel = this.decodeChannel(id), base = { ...channel, parentMessageId: null };
    if (channel.key.attribute === 'search') {
      const target = JSON.parse(channel.key.value), requested = parentMessageId === undefined ? channel.parentMessageId : parentMessageId;
      if (requested !== channel.parentMessageId) throw new AdapterError('invalid_channel', 'A search channel ID is bound to one parent. Use a sidebar channel ID for other threads or top-level posts.', 400);
      const opened = await this.openSearch(this.chatId({ query: target.query, key: target.resultKey }));
      if (opened.threadId !== channel.threadId || opened.parentMessageId !== channel.parentMessageId) throw new AdapterError('channel_changed', 'Search context changed.'); return opened;
    }
    try { await this.dom('extended:openChannel', base); }
    catch (error) {
      if (error.code !== 'channel_unavailable') throw error;
      const found = await this.channels(100, base); if (!found.items.some(n => n.threadId === base.threadId)) throw error;
    }
    const parent = parentMessageId === undefined ? channel.parentMessageId : parentMessageId;
    if (!parent) return base;
    const context = { ...base, parentMessageId: messageId(parent) };
    for (let i = 0; i < 100; i++) {
      try {
        const args = { ...context, token: randomUUID() }, opened = await this.dom('extended:openThread', args);
        if (opened.requiresContextMenu) { await this.contextClick(opened); await this.dom('extended:threadMenu', args); }
        return context;
      }
      catch (error) { if (error.code !== 'message_unavailable' || i === 99) throw error; if (!(await this.dom('extended:older', base)).changed) throw error; }
    }
  }
  async channelThreads(id) {
    const channel = await this.openChannel(id, null);
    return { channelId: id, ...await this.dom('extended:threads', channel), completeHistory: false, mayMarkRead: true };
  }
  async threadMessages(id, parentMessageId) {
    const context = await this.openChannel(id, parentMessageId);
    return { channelId: id, parentMessageId, conversationId: this.chatId(context), ...await this.dom('extended:messages', context), completeHistory: false, mayMarkRead: true };
  }
  async prepareContent(id, content, parentMessageId) {
    const decoded = this.decodeChat(id);
    const context = decoded.kind === 'channel' ? await this.openChannel(id, parentMessageId) : await this.openChat(id);
    const args = { ...context, token: randomUUID() };
    const cleanup = () => this.dom('extended:release', args);
    try { await this.dom('extended:prepare', args); return { args, content, cleanup }; }
    catch (error) { await cleanup(); throw error; }
  }
  async sendContent(prepared, checkCancelled = () => {}) {
    const { args, content } = prepared;
    checkCancelled();
    const result = await this.dom('extended:paste', { ...args, content, replace: prepared.edit === true });
    for (const mention of result.mentions) {
      checkCancelled(); const focused = await this.dom('extended:mentionFocus', { ...args, index: mention.index });
      // Native @mention autocomplete starts on key events, not bulk insertText.
      // A persistent beforeinput guard pins every character to this editor.
      await this.page.call('Input.dispatchKeyEvent', { type: 'keyDown', key: '@', code: 'Digit2', windowsVirtualKeyCode: 50, modifiers: 8, text: '@', unmodifiedText: '@' });
      await this.page.call('Input.dispatchKeyEvent', { type: 'keyUp', key: '@', code: 'Digit2', windowsVirtualKeyCode: 50, modifiers: 8 });
      checkCancelled(); await this.page.call('Input.insertText', { text: focused.query.slice(1) });
      await this.dom('extended:mentionSelect', { ...args, index: mention.index });
    }
    const text = contentText(content);
    await this.dom('extended:checkContent', { ...args, content, text }); checkCancelled();
    return this.dom(prepared.edit ? 'extended:editCommit' : 'extended:commit', { ...args, content, text });
  }
  async history(id, { cursor, limit = 100, maxWindows = 10, format = 'json' } = {}) {
    const now = Date.now();
    for (const [key, value] of this.historyCursors) if (value.expires < now) this.historyCursors.delete(key);
    let state;
    if (cursor) {
      const entry = this.historyCursors.get(cursor);
      if (!entry || entry.id !== id || entry.limit !== limit || entry.format !== format) throw new AdapterError('invalid_history_cursor', 'Cursor expired or belongs to different history parameters. Restart without a cursor.', 400);
      if (entry.result) return entry.result;
      state = entry.state;
    } else {
      if (this.historyCursors.size >= 1000) throw new AdapterError('history_capacity', 'Too many history cursors; wait for expiry.', 429);
      const decoded = this.decodeChat(id);
      const context = decoded.kind === 'channel' ? await this.openChannel(id, decoded.parentMessageId) : await this.openChat(id);
      state = { context, pending: [], seen: new Set(), exhausted: false, stable: 0, anchor: null, windows: 0 };
      const first = await this.dom('extended:messages', context);
      state.pending = first.items; for (const item of first.items) { if (!item.id) throw new AdapterError('message_ids_unavailable', 'History paging requires stable message IDs.'); state.seen.add(item.id); }
    }
    const add = items => {
      for (const item of items) {
        if (!item.id) throw new AdapterError('message_ids_unavailable', 'History paging requires stable message IDs.');
        if (!state.seen.has(item.id)) { state.pending.unshift(item); state.seen.add(item.id); }
      }
      // Incoming windows are chronological; reverse the new prefix back below.
    };
    let scanned = 0;
    // The cursor retains the scan position. If another operation changed views,
    // reopen, then seek the retained anchor across subsequent bounded requests.
    try { await this.dom('extended:messages', state.context); }
    catch (error) {
      if (!['chat_changed', 'channel_changed', 'selector_mismatch', 'thread_identity_unavailable'].includes(error.code)) throw error;
      state.context = state.context.kind === 'channel' ? await this.openChannel(id, state.context.parentMessageId) : await this.openChat(id);
      state.seeking = !!state.anchor;
    }
    while (scanned < maxWindows && !state.exhausted && (state.seeking || state.pending.length < limit)) {
      const page = await this.dom('extended:older', state.context); scanned++; state.windows++;
      if (state.seeking) {
        if (page.items.some(n => n.id === state.anchor)) state.seeking = false;
        else { if (!page.changed && page.atStart && !page.loading) throw new AdapterError('history_anchor_unavailable', 'The cursor anchor is no longer exposed. Restart history.', 409); continue; }
      }
      const fresh = page.items.filter(n => !state.seen.has(n.id));
      add([...fresh].reverse());
      state.stable = !page.changed && page.atStart && !page.loading ? state.stable + 1 : 0;
      if (state.stable >= 2) state.exhausted = true;
    }
    if (state.seen.size > 100000) throw new AdapterError('history_capacity', 'This export reached 100,000 observed messages. Split work into another session.', 429);
    const items = state.seeking ? [] : state.pending.splice(Math.max(0, state.pending.length - limit));
    state.anchor = state.pending[0]?.id ?? items[0]?.id ?? state.anchor;
    const hasMore = state.seeking || !state.exhausted || state.pending.length > 0;
    const nextCursor = hasMore ? randomUUID() : null;
    const result = { conversationId: id, items, nextCursor, hasMore, windowsScanned: scanned, totalWindowsScanned: state.windows, completeHistory: false, uiStartReached: state.exhausted, coverage: 'native_history_pagination', mayMarkRead: true, ...(format === 'ndjson' ? { ndjson: items.map(item => JSON.stringify(item)).join('\n') + (items.length ? '\n' : '') } : {}) };
    if (nextCursor) this.historyCursors.set(nextCursor, { id, limit, format, state, expires: now + 30 * 60_000 });
    if (cursor) this.historyCursors.get(cursor).result = result;
    return result;
  }
  async prepare(id) {
    const chat = await this.openChat(id);
    return { chat, ...await this.dom('prepare', chat) };
  }
  async prepareRecipient(email) {
    const chat = await this.openRecipient(email);
    return { chat, ...await this.dom('prepare', chat) };
  }
  async prepareQuote(id, replyToMessageId) {
    replyToMessageId = messageId(replyToMessageId);
    const chat = await this.openChat(id);
    let point;
    for (let page = 0; page <= 5; page++) {
      try { point = await this.dom('quoteTarget', { ...chat, replyToMessageId }); break; }
      catch (error) {
        if (error.code !== 'message_unavailable' || page === 5) throw error;
        const older = await this.dom('older', chat);
        if (!older.changed) throw error;
      }
    }
    await this.page.call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'right', clickCount: 1 });
    await this.page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'right', clickCount: 1 });
    return { chat, ...await this.dom('quoteSelect', { ...chat, replyToMessageId }) };
  }
  async contextClick(point) {
    await this.page.call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'right', clickCount: 1 });
    await this.page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'right', clickCount: 1 });
  }
  async prepareMessageAction(id, mid, kind, payload = {}) {
    mid = messageId(mid);
    if (['edit', 'delete'].includes(kind)) messageText(payload.expectedText, 'expectedText', true);
    if (kind === 'edit') { if (payload.content) richContent(payload.content); else messageText(payload.text); }
    if (kind === 'reaction') payload = reactionInput(payload.reaction, payload.selected);
    const chat = await this.openChat(id), args = { ...chat, messageId: mid, token: randomUUID(), kind, ...payload };
    let point;
    const cleanup = () => this.dom('action:release', args);
    try {
      for (let page = 0; page <= 5; page++) {
        try { point = await this.dom('action:messageTarget', args); break; }
        catch (error) { if (error.code !== 'message_unavailable' || page === 5) throw error; if (!(await this.dom('older', chat)).changed) throw error; }
      }
      if (['reaction', 'reactions'].includes(kind)) {
        await this.contextClick(point);
        const result = await this.dom('action:reactionReady', args);
        if (kind === 'reactions') return { result, cleanup };
        await this.dom('action:reactionPrepare', args);
      } else {
        await this.contextClick(point); await this.dom('action:menuPrepare', args);
        if (kind === 'edit') await this.dom('action:editPrepare', args);
      }
      if (kind === 'edit' && payload.content) {
        await this.dom('extended:adoptEdit', args); await this.dom('action:release', args);
        return { args, content: payload.content, edit: true, cleanup: () => this.dom('extended:release', args) };
      }
      return { args, cleanup };
    } catch (error) { await cleanup(); throw error; }
  }
  async reactions(id, mid) {
    const prepared = await this.prepareMessageAction(id, mid, 'reactions');
    try { return prepared.result; } finally { await prepared.cleanup(); }
  }
  async executeMessageAction(prepared, checkCancelled = () => {}) {
    const { args } = prepared;
    if (prepared.content) return this.sendContent(prepared, checkCancelled);
    if (args.kind === 'edit') {
      await this.dom('action:editSelect', args); checkCancelled();
      await this.page.call('Input.insertText', { text: args.text });
    }
    checkCancelled();
    return this.dom(`action:${args.kind}Commit`, args);
  }
  async prepareUnread(id) {
    await this.dom('status');
    let chat = this.decodeChat(id);
    if (['recipient', 'search'].includes(chat.key.attribute)) chat = await this.openChat(id);
    const args = { ...chat, token: randomUUID(), kind: 'unread' };
    let scanToken, point;
    const cleanup = async () => {
      try { await this.dom('action:release', args); }
      finally { if (scanToken) { const token = scanToken; scanToken = null; await this.dom('scanEnd', { token }); } }
    };
    try {
      try { point = await this.dom('action:unreadTarget', args); }
      catch (error) {
        if (error.code !== 'chat_unavailable') throw error;
        scanToken = randomUUID(); await this.dom('scanBegin', { token: scanToken, unreadOnly: false });
        for (let page = 0; page < 100; page++) {
          const window = await this.dom(page ? 'scanNext' : 'scanWindow', { token: scanToken });
          if (window.items.some(item => item.threadId === chat.threadId)) { point = await this.dom('action:unreadTarget', args); break; }
          if (window.atEnd && !window.more && !window.collapsed) break;
        }
        if (!point) throw new AdapterError('chat_unavailable', 'Could not find this chat during sidebar discovery.', 404);
      }
      if (!point.alreadyUnread) { await this.contextClick(point); await this.dom('action:unreadPrepare', args); }
      return { args, cleanup };
    } catch (error) { await cleanup(); throw error; }
  }
  async send(prepared, text) {
    // Input.insertText dispatches normal editor input. Never inject HTML or press Enter
    // (Teams' Enter-to-send setting varies). Click only the verified Send button.
    await this.page.call('Input.insertText', { text });
    const point = await this.dom('sendPoint', { ...prepared.chat, text, replyToMessageId: prepared.replyToMessageId });
    await this.page.clickPoint(point);
    return this.dom('confirm', { ...prepared.chat, text, replyToMessageId: prepared.replyToMessageId, quotePreview: prepared.quotePreview, beforeIds: prepared.beforeIds });
  }
  async close() { this.disposed = true; await this.queue; await this.cdp?.close(); }
}

export const profilePath = (dataDir, profile) => {
  if (!/^[a-f0-9-]{36}$/.test(profile)) throw new AdapterError('invalid_profile', 'Invalid browser profile.', 400);
  return path.join(dataDir, 'profiles', profile);
};

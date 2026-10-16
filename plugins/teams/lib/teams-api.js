import { randomUUID, randomBytes, createHash } from 'node:crypto';
import { AdapterError } from './errors.js';
import { requestJSON } from './api-http.js';
import { messageId, recipientEmail, messageText, reactionInput, peopleQuery, messageQuery } from './input.js';
import { richContent } from './content.js';
import { linkedText } from './linkify.js';

export const unsupported = () => { throw new AdapterError('api_operation_unsupported', 'This operation has not been migrated to the Teams API. See getTeamsCapabilities.', 501); };
export const capabilities = {
  transport: 'api', browserUse: 'authentication_only', authModes: ['browser', 'tokens'],
  supported: ['status', 'chats', 'unread_chats', 'unread_messages', 'messages', 'chat_history', 'channels', 'channel_history', 'exact_email_lookup', 'people_search', 'message_search', 'open_search_result', 'new_conversation', 'mentions', 'quoted_replies', 'channel_thread_history', 'channel_root_listing', 'existing_one_to_one', 'send_chat', 'formatted_chat', 'edit_chat', 'delete_chat', 'chat_reactions', 'mutation_state_readback', 'chat_discovery_pagination', 'automatic_url_links', 'channel_thread_replies'],
  unsupported: ['mark_unread', 'send_channel_thread', 'attachments'],
  liveValidated: ['oauth_tokens', 'profile_read', 'conversation_list', 'recent_messages', 'people_search', 'message_search', 'open_search_result.chat', 'open_search_result.channel', 'exact_message_read', 'channel_root_listing', 'channel_thread_reads', 'channel_thread_history', 'channel_thread_history.deep_pagination', 'quoted_replies.send', 'mentions.server_metadata', 'mutation_state_readback.live_writes', 'chat_discovery_pagination', 'automatic_url_links.send_edit'],
  pendingValidation: ['new_conversation.first_send', 'mentions.notifications', 'channel_thread_replies.send'],
};
export const encodeConversation = (kind, threadId) => Buffer.from(JSON.stringify({ version: 1, kind, threadId })).toString('base64url');
const encode = encodeConversation;
const escape = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const entities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
export const arrivalMillis = value => {
  if (typeof value === 'number') return value;
  if (typeof value === 'string' && /^\d{13}$/.test(value)) return Number(value);
  return typeof value === 'string' ? Date.parse(value) : NaN;
};
export function consumptionBoundary(conversation) {
  let value = conversation.userConsumptionHorizon ?? conversation.consumptionHorizon;
  if (conversation.userConsumptionHorizon && Number(value.originalArrivalTime ?? value.OriginalArrivalTime) === 0) value = conversation.consumptionHorizon;
  const boundary = Number(value?.originalArrivalTime ?? value?.OriginalArrivalTime);
  return Number.isFinite(boundary) && boundary > 0 ? boundary : null;
}
export function plainText(content, type = '') {
  if (!/html/i.test(type)) return String(content ?? '');
  return String(content ?? '').replace(/<br\s*\/?\s*>/gi, '\n').replace(/<\/(p|div|li|blockquote)>/gi, '\n').replace(/<[^>]*>/g, '')
    .replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (full, entity) => {
      if (!entity.startsWith('#')) return entities[entity.toLowerCase()] ?? full;
      const number = entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
      return number > 0 && number <= 0x10ffff ? String.fromCodePoint(number) : full;
    }).trimEnd();
}
function arrayProperty(value) {
  if (typeof value === 'string') { try { value = JSON.parse(value); } catch { return []; } }
  return Array.isArray(value) ? value : [];
}
export function normalize(raw) {
  const parent = raw.parentMessageId ?? raw.properties?.parentMessageId ?? raw.rootMessageId;
  return { id: String(raw.id), text: plainText(raw.content, raw.messagetype ?? raw.messageType),
    author: raw.imdisplayname ?? raw.imDisplayName ?? null, authorId: typeof raw.from === 'string' ? raw.from.split('/contacts/').at(-1) : null,
    time: raw.originalarrivaltime ?? raw.originalArrivalTime ?? raw.composetime ?? raw.composeTime ?? null,
    richText: /html/i.test(raw.messagetype ?? raw.messageType ?? '') ? String(raw.content ?? '') : null,
    deleted: Boolean(raw.properties?.deletetime ?? raw.properties?.deleteTime) || (raw.messagetype ?? raw.messageType) === 'MessageDelete',
    reactions: arrayProperty(raw.properties?.emotions), mentions: arrayProperty(raw.properties?.mentions),
    parentMessageId: parent == null || String(parent) === String(raw.id) ? null : String(parent) };
}
function strictArray(value) {
  if (value === undefined) return [];
  if (typeof value === 'string') { try { value = JSON.parse(value); } catch { throw new AdapterError('api_schema_changed', 'Teams message metadata changed.', 502); } }
  if (!Array.isArray(value)) throw new AdapterError('api_schema_changed', 'Teams message metadata changed.', 502);
  return value;
}
function confirmationContent(raw) {
  if ((raw.messagetype ?? raw.messageType) !== 'RichText/Html' || typeof raw.content !== 'string') return raw.content;
  // Live quote readback inserts CRLFs at these four block boundaries, even
  // when separators were supplied. Ignore only those separators in our exact
  // native quote prefix; retain quote IDs, author, preview and reply content.
  return raw.content.replace(/^(<blockquote itemscope="" itemtype="http:\/\/schema\.skype\.com\/Reply" itemid="[^"]+">)(?:\r\n)*(<strong itemprop="mri" itemid="[^"]+">[^<]*<\/strong>)(?:\r\n)*(<p itemprop="preview">[^<]*<\/p>)(?:\r\n)*(<\/blockquote>)(?:\r\n)*/, (_, ...parts) => parts.slice(0, 4).join(''));
}
function contentStamp(raw) {
  const mentions = strictArray(raw.properties?.mentions).map(item => [item.itemid, item.mri, item.mentionType, item.displayName]);
  return createHash('sha256').update(JSON.stringify([confirmationContent(raw), raw.messagetype ?? raw.messageType, mentions])).digest('hex');
}
const presets = ['like', 'heart', 'laugh', 'surprised', 'sad', 'angry'];
const writableReaction = value => typeof value === 'string' && /^[a-z][a-z0-9_-]{0,63}$/.test(value);
function order(items) { return items.sort((a, b) => String(a.time ?? a.id).localeCompare(String(b.time ?? b.id))); }
export function messagePayload(input, name, people = new Map()) {
  const mentions = [];
  let content, messagetype;
  if (input.content !== undefined) {
    const blocks = richContent(input.content);
    const runs = values => values.map(run => {
      if (run.mention) {
        const person = people.get(run.mention.email);
        if (!person || !/^8:orgid:[a-f0-9-]{36}$/i.test(person.mri)) throw new AdapterError('mention_unresolved', 'Resolve every mention by exact email before sending.', 400);
        const itemid = mentions.length;
        mentions.push({ '@type': 'http://schema.skype.com/Mention', itemid, mri: person.mri, mentionType: 'person', displayName: person.name });
        return `<span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="${itemid}">${escape(person.name)}</span>`;
      }
      const pieces = run.link || run.marks.includes('code') ? [{ text: run.text }] : linkedText(run.text);
      let value = pieces.map(piece => {
        const text = escape(piece.text).replace(/\n/g, '<br>');
        return piece.link ? `<a href="${escape(piece.link)}">${text}</a>` : text;
      }).join('');
      for (const mark of run.marks) { const tag = { bold: 'strong', italic: 'em', underline: 'u', strike: 's', code: 'code' }[mark]; value = `<${tag}>${value}</${tag}>`; }
      return run.link ? `<a href="${escape(run.link)}">${value}</a>` : value;
    }).join('');
    content = blocks.map(block => {
      if (block.runs) { const tag = block.type === 'quote' ? 'blockquote' : 'p'; return `<${tag}>${runs(block.runs)}</${tag}>`; }
      const tag = block.type === 'numberedList' ? 'ol' : 'ul'; return `<${tag}>${block.items.map(item => `<li>${runs(item)}</li>`).join('')}</${tag}>`;
    }).join(''); messagetype = 'RichText/Html';
  } else {
    const text = messageText(input.text), pieces = linkedText(text);
    if (pieces.some(piece => piece.link)) {
      content = '<p>' + pieces.map(piece => {
        const value = escape(piece.text).replace(/\n/g, '<br>');
        return piece.link ? `<a href="${escape(piece.link)}">${value}</a>` : value;
      }).join('') + '</p>';
      messagetype = 'RichText/Html';
    } else { content = text; messagetype = 'Text'; }
  }
  return { content, messagetype, contenttype: 'text', imdisplayname: name, properties: { importance: '', subject: null, ...(mentions.length ? { mentions: JSON.stringify(mentions) } : {}) } };
}

export function addQuote(payload, selected) {
  const preview = selected.text.length > 200 ? `${selected.text.slice(0, 197)}...` : selected.text;
  const quote = `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${escape(selected.id)}"><strong itemprop="mri" itemid="${escape(selected.authorId)}">${escape(selected.author || selected.authorId)}</strong><p itemprop="preview">${escape(preview)}</p></blockquote>`;
  const content = payload.messagetype === 'Text' ? `<p>${escape(payload.content).replace(/\n/g, '<br>')}</p>` : payload.content;
  payload.content = quote + content; payload.messagetype = 'RichText/Html';
}

export class TeamsAPI {
  constructor(session, { fetcher = fetch } = {}) { this.session = session; this.fetcher = fetcher; this.cursors = new Map(); this.searchResults = new Map(); this.discoveryCursors = new Map(); }
  get credentials() {
    const c = this.session.credentials;
    if (!c) throw new AdapterError('signin_required', 'Reconnect this connection to enable API authentication.', 503);
    return c;
  }
  async request(service, suffix, options = {}) {
    const c = this.credentials;
    if (!['messages', 'csa', 'mt', 'search'].includes(service)) throw new AdapterError('api_url_rejected', 'Unknown Teams service.', 400);
    const base = service === 'search' ? 'https://substrate.office.com/search/api/' : service === 'messages' ? c.messageOrigin + '/v1/' : service === 'csa' ? 'https://teams.microsoft.com/api/csa/api/v1/' : `https://teams.microsoft.com/api/mt/${c.region}/beta/`;
    const url = new URL(suffix, base);
    if (!url.href.startsWith(base) || url.username || url.password) throw new AdapterError('api_url_rejected', 'Invalid Teams API target.', 400);
    if (service === 'search' && !c.tokens.substrate?.value) throw new AdapterError('search_token_required', 'Search requires a Substrate token acquired using renewal credentials.', 503);
    const auth = service === 'messages' ? { Authentication: `skypetoken=${c.chatToken.value}` } : { Authorization: `Bearer ${c.tokens[service === 'search' ? 'substrate' : service === 'csa' ? 'chatsvcagg' : 'skype']?.value}` };
    return requestJSON(url.href, { ...options, headers: { ...auth, ...(options.body ? { 'content-type': 'application/json' } : {}) } }, this.fetcher);
  }
  conversationURL(limit, link) {
    const first = new URL(`users/ME/conversations?view=mychats&pageSize=${limit}`, this.credentials.messageOrigin + '/v1/');
    let url, pathname;
    try { url = new URL(link ?? first.href, first); pathname = decodeURIComponent(url.pathname); }
    catch { throw new AdapterError('api_cursor_rejected', 'Invalid conversation continuation.', 502); }
    const paths = ['/v1/users/ME/conversations', `/v1/users/${this.selfMri()}/conversations`];
    if (url.origin !== this.credentials.messageOrigin || url.username || url.password || url.hash || !paths.includes(pathname)) {
      throw new AdapterError('api_cursor_rejected', 'Conversation continuation changed host, account or resource.', 502);
    }
    return url;
  }
  async conversationPage(limit = 100, link, signal) {
    const url = this.conversationURL(limit, link);
    const value = await this.request('messages', url.href, { signal });
    if (!Array.isArray(value?.conversations) || value.conversations.some(c => typeof c?.id !== 'string' || !c.id || /[\u0000-\u001f/]/.test(c.id)) ||
        value._metadata?.backwardLink != null && typeof value._metadata.backwardLink !== 'string') throw new AdapterError('api_schema_changed', 'Conversation discovery response changed.', 502);
    return { raw: value.conversations, next: value._metadata?.backwardLink ? this.conversationURL(limit, value._metadata.backwardLink).href : null };
  }
  discoveredChat(raw) {
    const type = raw.threadProperties?.threadType;
    if (!['chat', 'meeting'].includes(type) || !/^19:[^/]+@(?:thread\.v2|unq\.gbl\.spaces)$/.test(raw.id)) return null;
    const props = raw.properties ?? {};
    const marker = value => typeof value === 'string' && /^\d+;\d+;\d+$/.test(value) && Number.isSafeInteger(Number(value.split(';')[0])) ? Number(value.split(';')[0]) : null;
    const normal = marker(props.consumptionhorizon), bookmark = props.consumptionHorizonBookmark == null ? 0 : marker(props.consumptionHorizonBookmark);
    const boundary = normal > 0 && bookmark !== null ? bookmark || normal : null;
    const last = arrivalMillis(props.lastimreceivedtime);
    return { id: raw.id, title: raw.threadProperties?.topic || props.displayName || 'Teams chat',
      discoverySource: 'regional_conversations', isRead: props.isemptyconversation === 'True' ? true : boundary !== null && Number.isFinite(last) ? last <= boundary : null,
      hidden: raw.threadProperties?.hidden === 'true' || raw.threadProperties?.hidden === 'True',
      ...(normal > 0 ? { consumptionHorizon: { originalArrivalTime: normal } } : {}),
      ...(bookmark !== null ? { userConsumptionHorizon: { originalArrivalTime: bookmark } } : {}),
      lastMessage: { originalArrivalTime: props.lastimreceivedtime ?? null } };
  }
  async discoverChats({ limit = 100, cursor, signal } = {}) {
    if (!Number.isInteger(limit) || limit < 1 || limit > 100) throw new AdapterError('invalid_limit', 'Discovery limit must be 1–100.', 400);
    for (const [key, state] of this.discoveryCursors) if (state.expires < Date.now()) this.discoveryCursors.delete(key);
    let state = { account: this.account(), limit, seen: [], links: [] };
    if (cursor) {
      state = this.discoveryCursors.get(cursor);
      if (!state || state.account !== this.account() || state.limit !== limit) throw new AdapterError('invalid_cursor', 'Discovery cursor expired or parameters/account changed.', 400);
      if (state.result) return state.result;
    }
    const page = await this.conversationPage(limit, state.link, signal), seen = new Set(state.seen), items = [];
    let excluded = 0;
    for (const raw of page.raw) {
      if (seen.has(raw.id)) continue;
      seen.add(raw.id); const chat = this.discoveredChat(raw);
      if (chat) items.push({ id: encode('chat', chat.id), title: chat.title, unread: typeof chat.isRead === 'boolean' ? !chat.isRead : null, hidden: chat.hidden, time: chat.lastMessage.originalArrivalTime });
      else excluded++;
    }
    // Normalize links before comparing so relative/absolute aliases cannot loop.
    const key = page.next ? new URL(page.next, this.credentials.messageOrigin + '/v1/users/ME/conversations').href : null;
    const repeated = key && [...state.links, state.link].includes(key), capped = seen.size >= 100_000;
    let nextCursor;
    if (page.next && !repeated && !capped) {
      if (this.discoveryCursors.size >= 1000) throw new AdapterError('cursor_capacity', 'Too many discovery cursors; restart discovery after existing cursors expire.', 503);
      nextCursor = randomUUID(); this.discoveryCursors.set(nextCursor, { account: state.account, limit, seen: [...seen], links: [...state.links, ...(state.link ? [state.link] : [])], link: key, expires: Date.now() + 1800_000 });
    }
    const result = { items, hasMore: Boolean(nextCursor), ...(nextCursor ? { nextCursor } : {}), discoveryComplete: !page.next,
      truncated: Boolean(repeated || capped), excludedConversations: excluded, completeAccount: false, coverage: 'regional_chat_discovery_page', mayMarkRead: false };
    if (cursor) state.result = result;
    return result;
  }
  async directory(force = false, signal, discover = false) {
    if (!force && this.directoryCache?.account === this.account() && this.directoryCache.expires > Date.now()) return this.directoryCache.value;
    const value = await this.request('csa', 'teams/users/me?isPrefetch=false&enableMembershipSummary=true', { signal });
    if (!value || !Array.isArray(value.chats) || !Array.isArray(value.teams)) throw new AdapterError('api_schema_changed', 'Teams conversation response no longer matches the expected shape.', 502);
    const partial = Boolean(value.metadata?.isPartialData || value.metadata?.hasMoreChats || value.metadata?.forwardSyncToken);
    if (partial || discover) {
      const seen = new Set(value.chats.map(c => c.id)), links = new Set(); let next, pages = 0, added = 0;
      const budget = signal ? AbortSignal.any([signal, AbortSignal.timeout(30_000)]) : AbortSignal.timeout(30_000);
      try {
        do {
          const page = await this.conversationPage(100, next, budget); pages++;
          for (const raw of page.raw) {
            if (seen.has(raw.id)) continue;
            seen.add(raw.id); const chat = this.discoveredChat(raw);
            if (chat) { value.chats.push(chat); added++; }
          }
          const key = page.next ? new URL(page.next, this.credentials.messageOrigin + '/v1/users/ME/conversations').href : null;
          if (key && links.has(key)) { next = key; break; }
          if (key) links.add(key); next = key;
        } while (next && pages < 10 && seen.size < 100_000);
        value.discovery = { pages, added, complete: !next, truncated: Boolean(next), coverage: 'csa_and_bounded_regional_chats' };
      } catch (error) {
        if (['api_unauthorized', 'api_rate_limited'].includes(error.code) || signal?.aborted) throw error;
        value.discovery = { pages, added, complete: false, truncated: true, reason: error instanceof AdapterError ? error.code : 'discovery_failed', coverage: 'csa_and_bounded_regional_chats' };
      }
    }
    this.directoryCache = { value, account: this.account(), expires: Date.now() + 30_000 }; return value;
  }
  async ensureDiscoveredMembership(conversation) {
    if (conversation.discoverySource !== 'regional_conversations') return;
    const roster = await this.request('messages', `threads/${encodeURIComponent(conversation.id)}/members`);
    if (!Array.isArray(roster?.members) || !roster.members.some(member => member.id === this.selfMri())) throw new AdapterError('conversation_unavailable', 'Current chat membership was not verified.', 404);
  }
  selfMri() { return `8:orgid:${this.credentials.identity.oid}`; }
  async profile(email = this.credentials.identity.email) {
    email = recipientEmail(email);
    const result = await this.request('mt', `users/${encodeURIComponent(email)}/?throwIfNotFound=false&isMailAddress=true&enableGuest=true&includeIBBarredUsers=true&skypeTeamsInfo=true`);
    const person = result?.value;
    const emails = [person?.email, person?.userPrincipalName, person?.upn].filter(value => typeof value === 'string').map(value => value.toLowerCase());
    if (!person || !emails.includes(email) || typeof person.mri !== 'string') throw new AdapterError('person_unavailable', 'Teams did not return an exact matching directory person.', 404);
    return { email, name: person.displayName || email, mri: person.mri };
  }
  async people(query) {
    query = peopleQuery(query);
    if (query.includes('@')) return { query, items: [await this.profile(query)], coverage: 'exact_email_lookup' };
    const response = await this.request('search', 'v1/suggestions?scenario=peoplepicker.newChat', { method: 'POST', body: JSON.stringify({
      EntityRequests: [{ Query: { QueryString: query, DisplayQueryString: query }, EntityType: 'People', Size: 20, From: 0,
        Fields: ['Id', 'MRI', 'DisplayName', 'EmailAddresses', 'UserPrincipalName', 'PeopleType', 'PeopleSubtype'],
        Filter: { And: [{ Or: [{ Term: { PeopleType: 'Person' } }, { Term: { PeopleType: 'Other' } }] },
          { Or: [{ Term: { PeopleSubtype: 'OrganizationUser' } }, { Term: { PeopleSubtype: 'MTOUser' } }, { Term: { PeopleSubtype: 'Guest' } }] },
          { Or: [{ Term: { Flags: 'NonHidden' } }] }] }, Provenances: ['Mailbox', 'Directory'] }],
      Cvid: randomUUID(), Scenario: { Name: 'peoplepicker.newChat' }, AppName: 'Microsoft Teams',
    }) });
    if (!Array.isArray(response?.Groups)) throw new AdapterError('api_schema_changed', 'Teams people-search response changed.', 502);
    const suggestions = response.Groups.find(group => group.Type === 'People')?.Suggestions ?? [];
    if (!Array.isArray(suggestions)) throw new AdapterError('api_schema_changed', 'Teams people-search results changed.', 502);
    const items = [], seen = new Set();
    for (const person of suggestions.slice(0, 20)) {
      let email; try { email = recipientEmail(person.EmailAddresses?.[0] ?? person.UserPrincipalName); } catch { continue; }
      if (typeof person.MRI !== 'string' || !person.MRI.startsWith('8:orgid:') || seen.has(email)) continue;
      seen.add(email); items.push({ email, name: person.DisplayName || email, mri: person.MRI });
    }
    return { query, items, coverage: 'api_people_suggestions', completeDirectory: false, truncated: suggestions.length >= 20 };
  }
  account() { const { tenant, oid } = this.credentials.identity; return `${tenant}:${oid}`; }
  rememberSearch(locator) {
    for (const [key, value] of this.searchResults) if (value.expires <= Date.now() || value.account !== this.account()) this.searchResults.delete(key);
    if (this.searchResults.size >= 1000) this.searchResults.delete(this.searchResults.keys().next().value);
    const id = randomUUID(); this.searchResults.set(id, { ...locator, account: this.account(), expires: Date.now() + 1800_000 }); return id;
  }
  searchLocator(source) {
    // Deep links supply locators only. Never fetch a server-supplied URL or
    // accept its account/membership assertions.
    if (typeof source.ClientConversationId === 'string' && typeof source.InternetMessageId === 'string') {
      const parse = value => typeof value === 'string' ? value.match(/^(19:[^;\/\s\u0000-\u001f]{1,1000}@(?:thread\.v2|thread\.tacv2|unq\.gbl\.spaces))(?:;messageid=(\d{1,20}))?$/) : null;
      const conversation = parse(source.ClientConversationId), thread = parse(source.ClientThreadId ?? source.ClientConversationId);
      if (!conversation || !thread || conversation[1] !== thread[1] || !/^\d{1,20}$/.test(source.InternetMessageId)) return null;
      const kind = conversation[1].endsWith('@thread.tacv2') ? 'channel' : 'chat';
      const parentMessageId = conversation[2] ?? thread[2] ?? null;
      if (conversation[2] && thread[2] && conversation[2] !== thread[2] || kind === 'chat' && parentMessageId) return null;
      return { threadId: conversation[1], kind, messageId: source.InternetMessageId, parentMessageId };
    }
    const raw = source.WebUrl ?? source.webUrl ?? source.Path;
    if (typeof raw !== 'string') return null;
    let url; try { url = new URL(raw); } catch { return null; }
    if (!['https://teams.microsoft.com', 'https://teams.cloud.microsoft'].includes(url.origin) || url.username || url.password) return null;
    const parts = url.pathname.match(/^\/l\/message\/([^/]+)\/([^/]+)\/?$/);
    if (!parts) return null;
    let threadId, mid; try { threadId = decodeURIComponent(parts[1]); mid = messageId(decodeURIComponent(parts[2])); } catch { return null; }
    if (!threadId || threadId.length > 1024 || /[\u0000-\u001f/;]/.test(threadId)) return null;
    const kind = threadId.endsWith('@thread.tacv2') ? 'channel' : 'chat';
    let parentMessageId = url.searchParams.get('parentMessageId');
    if (kind === 'channel') {
      if (!parentMessageId) {
        try { parentMessageId = JSON.parse(url.searchParams.get('context') ?? '{}').parentMessageId ?? null; } catch { return null; }
      }
      if (parentMessageId !== null && (typeof parentMessageId !== 'string' || !/^\d{1,20}$/.test(parentMessageId))) return null;
    } else parentMessageId = null;
    return { threadId, kind, messageId: mid, parentMessageId };
  }
  async search(query, limit = 50, maxPages = 3) {
    query = messageQuery(query);
    const items = [], seen = new Set(); let hasMore = false, pages = 0;
    const size = Math.min(limit, 50);
    for (; pages < maxPages; pages++) {
      const response = await this.request('search', 'v2/query', { method: 'POST', body: JSON.stringify({
        EntityRequests: [{ entityType: 'Message', contentSources: ['Teams'], propertySet: 'Optimized',
          fields: ['Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String', 'Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String', 'Extension_SkypeSpaces_ConversationPost_Extension_SkypeGroupId_String'],
          query: { queryString: query, displayQueryString: query }, from: pages * size, size }],
        cvid: randomUUID(), scenario: { Name: 'powerbar', Dimensions: [{ DimensionName: 'QueryType', DimensionValue: 'All' }, { DimensionName: 'FormFactor', DimensionValue: 'general.desktop.reactSearch' }] },
      }) });
      const sets = response?.EntitySets ?? response?.entitySets;
      const results = sets?.[0]?.ResultSets ?? sets?.[0]?.resultSets;
      const hits = results?.[0]?.Results ?? results?.[0]?.results;
      if (!Array.isArray(sets) || !Array.isArray(results) || !Array.isArray(hits)) throw new AdapterError('api_schema_changed', 'Teams message-search response changed.', 502);
      const more = results[0].MoreResultsAvailable ?? results[0].moreResultsAvailable;
      hasMore = typeof more === 'boolean' ? more : hits.length >= size;
      for (const hit of hits) {
        const source = hit.Source ?? hit.source ?? {}, locator = this.searchLocator(source);
        const key = locator ? JSON.stringify(locator) : String(hit.Id ?? hit.id ?? source.WebUrl ?? source.Path ?? JSON.stringify(source));
        if (seen.has(key)) continue; seen.add(key);
        const from = source.From ?? source.from ?? source.Creator ?? source.creator;
        items.push({ resultId: locator ? this.rememberSearch(locator) : null, canOpen: Boolean(locator),
          text: plainText(source.Preview ?? source.preview ?? hit.HitHighlightedSummary ?? hit.Summary ?? '', 'html'),
          author: typeof from === 'string' ? from : from?.EmailAddress?.Name ?? from?.DisplayName ?? from?.displayName ?? null,
          time: source.DateTimeSent ?? source.ItemDate ?? source.LastModifiedTime ?? source.itemDate ?? null,
          ...(locator ? { kind: locator.kind } : { reason: 'exact_locator_unavailable' }) });
        if (items.length >= limit) { hasMore ||= hits.indexOf(hit) < hits.length - 1; break; }
      }
      if (!hasMore || items.length >= limit) { pages++; break; }
    }
    return { query, items, pages, hasMore, truncated: hasMore, completeSearch: !hasMore,
      completeAccount: false, coverage: 'api_server_search', mayMarkRead: false };
  }
  async openSearchResult(resultId) {
    const locator = this.searchResults.get(resultId);
    if (!locator || locator.expires <= Date.now() || locator.account !== this.account()) throw new AdapterError('invalid_search_result', 'Search result expired or belongs to another account. Search again.', 400);
    await this.directory(true);
    const conversationId = encode(locator.kind, locator.threadId);
    await this.resolve(conversationId, locator.kind);
    let parent = locator.parentMessageId, found;
    if (locator.kind === 'channel' && parent) found = await this.exactMessage(`${locator.threadId};messageid=${parent}`, locator.messageId);
    else found = await this.exactMessage(locator.threadId, locator.messageId);
    if (locator.kind === 'channel') {
      parent ??= found.message.parentMessageId ?? found.message.id;
      const context = await this.threadMessages(conversationId, parent, 50);
      if (found.message.parentMessageId && found.message.parentMessageId !== parent) throw new AdapterError('message_unavailable', 'Search hit no longer matches its channel thread.', 404);
      return { ...context, channelId: conversationId, parentMessageId: parent, selectedMessage: found.message };
    }
    const context = await this.messages(conversationId, 50);
    return { ...context, chatId: conversationId, selectedMessage: found.message };
  }
  async chats(unreadOnly = false) {
    const data = await this.directory(true, undefined, true);
    const items = data.chats.filter(chat => !chat.isConversationDeleted && (!unreadOnly || chat.isRead === false)).map(chat => ({
      id: encode('chat', chat.id), title: chat.title || chat.members?.filter(member => member.mri !== this.selfMri()).map(member => member.friendlyName || member.mri).join(', ') || 'Teams chat',
      unread: typeof chat.isRead === 'boolean' ? !chat.isRead : null, hidden: Boolean(chat.hidden),
      preview: chat.lastMessage ? plainText(chat.lastMessage.content, chat.lastMessage.messageType) : null,
      time: chat.lastMessage?.originalArrivalTime ?? null,
    }));
    const partial = Boolean(data.metadata?.isPartialData || data.metadata?.hasMoreChats || data.metadata?.forwardSyncToken);
    return { items, discoveryComplete: data.discovery ? data.discovery.complete : !partial, truncated: data.discovery ? data.discovery.truncated : partial, ...(data.discovery ? { discovery: data.discovery } : {}), completeAccount: false, coverage: data.discovery?.coverage ?? 'api_conversation_snapshot', mayMarkRead: false };
  }
  async channels() {
    const data = await this.directory();
    return { items: data.teams.flatMap(team => (team.channels || []).filter(channel => !channel.isDeleted).map(channel => ({ id: encode('channel', channel.id), title: channel.displayName, teamTitle: team.displayName }))), coverage: 'api_conversation_snapshot', completeAccount: false, mayMarkRead: false };
  }
  async resolve(id, kind = 'chat') {
    let value;
    try { if (typeof id !== 'string' || id.length > 4096) throw new Error(); value = JSON.parse(Buffer.from(id, 'base64url')); }
    catch { throw new AdapterError('invalid_conversation', 'Use a conversation ID returned by the Teams listing.', 400); }
    // Old DOM IDs with an exact threadId can migrate after server membership lookup.
    if (typeof value.threadId !== 'string' || !value.threadId || value.threadId.length > 1024 || /[\u0000-\u001f/]/.test(value.threadId) || value.kind && value.kind !== kind || value.parentMessageId) throw new AdapterError('invalid_conversation', 'Refresh the conversation ID with listTeamsChats or listTeamsChannels.', 400);
    const data = await this.directory();
    let match = kind === 'chat' ? data.chats.find(chat => chat.id === value.threadId) : data.teams.flatMap(team => team.channels || []).find(channel => channel.id === value.threadId);
    if (!match && kind === 'chat' && !value.recipientEmail) {
      // A selected ID may lie beyond the bounded directory scan. Verify the
      // current roster, then fetch its exact account-scoped metadata directly.
      try {
        await this.ensureDiscoveredMembership({ id: value.threadId, discoverySource: 'regional_conversations' });
        const raw = await this.request('messages', `users/ME/conversations/${encodeURIComponent(value.threadId)}?view=msnp24Equivalent`);
        if (raw?.id !== value.threadId) throw new AdapterError('api_schema_changed', 'Teams returned a different conversation.', 502);
        match = this.discoveredChat(raw);
      } catch (error) {
        if (!['api_not_found', 'api_forbidden', 'conversation_unavailable'].includes(error.code)) throw error;
      }
    }
    if (!match && kind === 'chat' && value.recipientEmail) {
      const person = await this.profile(value.recipientEmail);
      const peer = person.mri.match(/^8:orgid:([a-f0-9-]{36})$/i)?.[1];
      if (peer && peer !== this.credentials.identity.oid && value.threadId === `19:${this.credentials.identity.oid}_${peer}@unq.gbl.spaces`) {
        match = { id: value.threadId, isOneOnOne: true, members: [{ mri: this.selfMri() }, { mri: person.mri }] };
      }
    }
    if (!match || match.isDeleted || match.isConversationDeleted || kind === 'channel' && match.isMember === false) throw new AdapterError('conversation_unavailable', 'Conversation was not returned for this account.', 404);
    return match;
  }
  async existingConversation(email) {
    const person = await this.profile(email), data = await this.directory(true);
    const matches = data.chats.filter(chat => chat.isOneOnOne && !chat.isConversationDeleted && chat.members?.some(member => member.mri === person.mri) && chat.members?.some(member => member.mri === this.selfMri()));
    if (!matches.length) {
      const peer = person.mri.match(/^8:orgid:([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})$/i)?.[1];
      if (!peer || person.mri === this.selfMri()) throw new AdapterError('person_unavailable', 'A new chat requires another organization directory user.', 400);
      const threadId = `19:${this.credentials.identity.oid}_${peer}@unq.gbl.spaces`;
      const chatId = Buffer.from(JSON.stringify({ version: 1, kind: 'chat', threadId, recipientEmail: person.email })).toString('base64url');
      return { email: person.email, title: person.name, chatId, messageSent: false, persisted: false };
    }
    if (matches.length !== 1) throw new AdapterError('ambiguous_conversation', 'Multiple one-to-one chats match this person. Select a chat from listTeamsChats.', 409);
    return { email: person.email, title: person.name, chatId: encode('chat', matches[0].id), messageSent: false, persisted: true };
  }
  messagePath(threadId) { return `users/ME/conversations/${encodeURIComponent(threadId)}/messages`; }
  async page(threadId, limit = 50, backwardLink, signal) {
    const expected = this.messagePath(threadId);
    let suffix = `${expected}?view=msnp24Equivalent%7CsupportsMessageProperties&pageSize=${limit}&startTime=1`;
    if (backwardLink) {
      let url;
      try { url = new URL(backwardLink, this.credentials.messageOrigin); }
      catch { throw new AdapterError('api_cursor_rejected', 'Teams returned an invalid history link.', 502); }
      if (url.origin !== this.credentials.messageOrigin || url.username || url.password || url.hash || decodeURIComponent(url.pathname) !== decodeURIComponent('/v1/' + expected)) throw new AdapterError('api_cursor_rejected', 'Teams history link changed host or conversation.', 502);
      suffix = url.href;
    }
    const response = await this.request('messages', suffix, { signal });
    if (!Array.isArray(response?.messages) || response.messages.some(item => typeof item?.id !== 'string' && typeof item?.id !== 'number')) throw new AdapterError('api_schema_changed', 'Teams message response no longer matches the expected shape.', 502);
    return { raw: response.messages, items: order(response.messages.map(normalize)), next: response._metadata?.backwardLink || null };
  }
  async messages(id, limit = 50, olderPages = 0, kind = 'chat') {
    const conversation = await this.resolve(id, kind), seen = new Map(); let page, next;
    for (let i = 0; i <= olderPages; i++) {
      page = await this.page(conversation.id, limit, next);
      for (const message of page.items) seen.set(message.id, message);
      if (!page.next || page.next === next) break; next = page.next;
    }
    const items = order([...seen.values()]);
    return { conversationId: id, items: olderPages ? items.slice(0, limit) : items.slice(-limit), hasMore: Boolean(page.next), completeHistory: false, coverage: 'api_messages', mayMarkRead: false };
  }
  async history(id, { limit = 100, cursor, format = 'json', kind = 'chat', parentMessageId = null, rootsOnly = false } = {}) {
    for (const [key, state] of this.cursors) if (state.expires < Date.now()) this.cursors.delete(key);
    let state;
    if (cursor) {
      state = this.cursors.get(cursor);
      if (!state || state.id !== id || state.limit !== limit || state.format !== format || state.kind !== kind || state.parentMessageId !== parentMessageId || state.rootsOnly !== rootsOnly || state.account !== this.account()) throw new AdapterError('invalid_cursor', 'History cursor expired or belongs to different parameters.', 400);
      if (state.result) return state.result;
    } else {
      const conversation = await this.resolve(id, kind);
      const thread = parentMessageId ? await this.threadSource(id, parentMessageId) : null;
      state = { id, limit, format, kind, parentMessageId, rootsOnly, account: this.account(),
        threadId: thread?.threadId ?? conversation.id, root: thread?.root, seen: [], links: [], next: null };
    }
    if (this.cursors.size >= 1000 || state.seen.length >= 100_000) throw new AdapterError('history_capacity', 'History traversal capacity reached.', 429);
    const page = await this.page(state.threadId, limit, state.next);
    const sourceItems = state.root ? order([...page.items.filter(item => item.id !== state.root.id), state.root]) : page.items;
    if (parentMessageId && sourceItems.some(item => item.parentMessageId && item.parentMessageId !== parentMessageId)) throw new AdapterError('api_schema_changed', 'Teams returned messages from a different thread.', 502);
    const seen = new Set(state.seen), items = sourceItems.filter(item => !rootsOnly || !item.parentMessageId).filter(message => { if (seen.has(message.id)) return false; seen.add(message.id); return true; });
    const hasMore = Boolean(page.next && page.next !== state.next && !state.links.includes(page.next)), nextCursor = hasMore ? randomUUID() : null;
    const result = { conversationId: id, items, hasMore, nextCursor, completeHistory: false, coverage: parentMessageId ? 'api_channel_thread_history' : rootsOnly ? 'api_channel_roots' : 'api_history', mayMarkRead: false,
      ...(parentMessageId ? { parentMessageId } : {}),
      ...(format === 'ndjson' ? { ndjson: items.map(item => JSON.stringify(item)).join('\n') + (items.length ? '\n' : '') } : {}) };
    if (nextCursor) this.cursors.set(nextCursor, { ...state, seen: [...seen], links: [...state.links, page.next], next: page.next, result: undefined, expires: Date.now() + 1800_000 });
    if (cursor) state.result = result;
    return result;
  }
  async threadSource(id, parentMessageId) {
    if (typeof parentMessageId !== 'string' || !/^\d{1,20}$/.test(parentMessageId)) throw new AdapterError('invalid_message_id', 'Use the numeric parentMessageId returned by a channel read.', 400);
    const conversation = await this.resolve(id, 'channel');
    const { message: root, raw } = await this.exactMessage(conversation.id, parentMessageId);
    if (root.parentMessageId && root.parentMessageId !== root.id) throw new AdapterError('invalid_thread_root', 'The selected channel message is a reply, not a root.', 400);
    return { threadId: `${conversation.id};messageid=${parentMessageId}`, root, rootVersion: raw.version ?? null, conversation, channelId: id, parentMessageId };
  }
  async threadMessages(id, parentMessageId, limit = 50) {
    const { threadId, root } = await this.threadSource(id, parentMessageId);
    const page = await this.page(threadId, limit);
    if (page.items.some(item => item.parentMessageId && item.parentMessageId !== parentMessageId)) throw new AdapterError('api_schema_changed', 'Teams returned messages from a different thread.', 502);
    return { conversationId: id, parentMessageId, root, items: order([...page.items.filter(item => item.id !== root.id), root]),
      hasMore: Boolean(page.next), completeHistory: false, coverage: 'api_channel_thread', mayMarkRead: false };
  }
  async unread(maxChats = 3, limit = 50) {
    const snapshot = await this.chats(true), chats = [], items = [];
    for (const chat of snapshot.items.slice(0, maxChats)) {
      const conversation = await this.resolve(chat.id), page = await this.messages(chat.id, limit);
      const boundary = consumptionBoundary(conversation);
      const boundaryFound = boundary !== null;
      const unread = boundaryFound ? page.items.filter(item => arrivalMillis(item.time) > boundary && item.authorId !== this.selfMri()) : [];
      chats.push({ chatId: chat.id, boundaryFound, items: unread, ...(boundaryFound ? {} : { recentMessages: page.items }) });
      items.push(...unread.map(item => ({ ...item, chatId: chat.id, chatTitle: chat.title })));
    }
    return { items, chats, remainingChats: Math.max(0, snapshot.items.length - maxChats), completeAccount: false, coverage: 'api_consumption_horizon_window', mayMarkRead: false };
  }
  async findMessage(threadId, mid, signal) {
    mid = messageId(mid); let next;
    for (let i = 0; i < 6; i++) {
      const page = await this.page(threadId, 200, next, signal), raw = page.raw.find(item => String(item.id) === mid);
      if (raw) return { raw, message: normalize(raw) };
      if (!page.next || page.next === next) break; next = page.next;
    }
    throw new AdapterError('message_unavailable', 'Message was not found in six recent API pages.', 404);
  }
  async exactMessage(threadId, mid) {
    mid = messageId(mid);
    const raw = await this.request('messages', `${this.messagePath(threadId)}/${encodeURIComponent(mid)}`);
    if (!raw || String(raw.id) !== mid) throw new AdapterError('api_schema_changed', 'Teams did not return the exact requested message.', 502);
    const message = normalize(raw);
    if (message.deleted || !message.authorId || !/^(Text|RichText(?:\/.*)?)$/.test(raw.messagetype ?? raw.messageType ?? '')) throw new AdapterError('message_unavailable', 'The selected message is unavailable for this operation.', 404);
    return { raw, message };
  }
  async payload(input) {
    const people = new Map();
    if (input.content !== undefined) {
      const blocks = richContent(input.content);
      const emails = new Set(blocks.flatMap(block => block.runs ?? block.items.flat()).filter(run => run.mention).map(run => run.mention.email));
      if (emails.size > 20) throw new AdapterError('invalid_content', 'At most 20 distinct person mentions are allowed.', 400);
      for (const email of emails) people.set(email, await this.profile(email));
    }
    return messagePayload(input, this.credentials.identity.name || this.credentials.identity.email, people);
  }
  async prepareSend(id, input) {
    const conversation = await this.resolve(id);
    await this.ensureDiscoveredMembership(conversation);
    if (conversation.isMessagingDisabled || conversation.isDisabled) throw new AdapterError('api_forbidden', 'Messaging is disabled for this chat.', 403);
    const payload = await this.payload(input);
    let quote;
    if (input.replyToMessageId !== undefined) {
      quote = await this.exactMessage(conversation.id, input.replyToMessageId);
      addQuote(payload, quote.message);
    }
    return { threadId: conversation.id, confirmation: this.confirmationContext(conversation.id, payload), ...(quote ? { quote } : {}), payload: { ...payload, clientmessageid: String(BigInt('0x' + randomBytes(8).toString('hex')) >> 1n) } };
  }
  async prepareChannelReply(id, parentMessageId, input) {
    const source = await this.threadSource(id, parentMessageId);
    if (source.conversation.isMessagingDisabled || source.conversation.isDisabled || source.conversation.isArchived) throw new AdapterError('api_forbidden', 'Messaging is disabled for this channel.', 403);
    const payload = await this.payload(input);
    return { threadId: source.threadId,
      channelReplyGuard: { channelId: id, parentMessageId, root: source.root, rootVersion: source.rootVersion },
      confirmation: { ...this.confirmationContext(source.threadId, payload), parentMessageId },
      payload: { ...payload, clientmessageid: String(BigInt('0x' + randomBytes(8).toString('hex')) >> 1n) } };
  }
  async prepareSendForTriage(source, input) {
    if (source.kind === 'channel' || source.parentMessageId) return this.prepareChannelReply(source.conversationId, source.parentMessageId, input);
    if (source.kind === 'chat') await this.ensureDiscoveredMembership(await this.resolve(source.conversationId));
    const payload = await this.payload(input);
    const threadId = source.threadId;
    return { threadId, confirmation: this.confirmationContext(threadId, payload), payload: { ...payload, clientmessageid: String(BigInt('0x' + randomBytes(8).toString('hex')) >> 1n) } };
  }
  confirmationContext(threadId, payload, kind = 'send', mid = null) {
    return { account: this.account(), threadId, kind, mid,
      ...(kind === 'send' || kind === 'edit' ? { contentStamp: contentStamp(payload) } : {}),
      ...(kind === 'reaction' ? { reaction: payload.reaction, selected: payload.selected } : {}) };
  }
  async confirm(result, context, signal) {
    if (!context || context.account !== this.account()) return { status: 'unavailable', reason: 'verification_context_unavailable' };
    const mid = context.kind === 'send' ? result.message?.id : context.mid;
    if (!mid) return { status: 'unavailable', reason: 'message_id_unavailable' };
    const bounded = signal ? AbortSignal.any([signal, AbortSignal.timeout(5000)]) : AbortSignal.timeout(5000);
    // GET retries only: the accepted write was saved before this method runs.
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const raw = await this.request('messages', `${this.messagePath(context.threadId)}/${encodeURIComponent(messageId(mid))}`, { signal: bounded });
        if (!raw || String(raw.id) !== mid) throw new AdapterError('api_schema_changed', 'Teams returned a different message for confirmation.', 502);
        const message = normalize(raw);
        let observed = false;
        if (context.kind === 'delete') observed = message.deleted;
        else if (context.kind === 'reaction') {
          const emotions = strictArray(raw.properties?.emotions);
          if (emotions.some(item => typeof item?.key !== 'string' || !Array.isArray(item.users) || item.users.some(user => typeof user?.mri !== 'string'))) throw new AdapterError('api_schema_changed', 'Teams reaction metadata changed.', 502);
          const selected = emotions.some(item => item.key === context.reaction && item.users.some(user => user.mri === this.selfMri()));
          observed = !message.deleted && selected === context.selected;
        } else if (['send', 'edit'].includes(context.kind)) {
          observed = !message.deleted && message.authorId === this.selfMri() && contentStamp(raw) === context.contentStamp &&
            (context.kind !== 'send' || String(raw.clientmessageid ?? raw.clientMessageId) === result.message.clientMessageId) &&
            (!context.parentMessageId || message.parentMessageId === context.parentMessageId &&
              [raw.parentMessageId, raw.properties?.parentMessageId, raw.rootMessageId].filter(value => value != null).every(value => String(value) === context.parentMessageId) &&
              (raw.conversationid == null || raw.conversationid === context.threadId.split(';messageid=')[0]));
        }
        if (observed) return { status: 'observed', operation: context.kind, messageId: mid, ...(context.parentMessageId ? { parentMessageId: context.parentMessageId, threadConfirmed: true } : {}), recipientDeliveryConfirmed: false };
      } catch (error) {
        // 404 is not a confirmed tombstone: retention or visibility can also
        // remove a message. Authentication/transport failure does not undo acceptance.
        return { status: 'unavailable', reason: error instanceof AdapterError ? error.code : 'verification_failed' };
      }
    }
    return { status: 'not_observed', reason: 'state_not_observed_in_bounded_readback' };
  }
  async send(prepared, signal) {
    if (prepared.channelReplyGuard) {
      const guard = prepared.channelReplyGuard;
      await this.directory(true, signal);
      const fresh = await this.threadSource(guard.channelId, guard.parentMessageId);
      if (fresh.conversation.isMessagingDisabled || fresh.conversation.isDisabled || fresh.conversation.isArchived) throw new AdapterError('api_forbidden', 'Messaging is disabled for this channel.', 403);
      if (fresh.threadId !== prepared.threadId || fresh.rootVersion !== guard.rootVersion || JSON.stringify(fresh.root) !== JSON.stringify(guard.root)) throw new AdapterError('message_changed', 'Channel root changed before sending. Read the thread again.', 409);
    }
    if (prepared.quote) {
      const fresh = await this.exactMessage(prepared.threadId, prepared.quote.message.id);
      if (JSON.stringify(fresh.message) !== JSON.stringify(prepared.quote.message) || fresh.raw.version !== prepared.quote.raw.version) throw new AdapterError('message_changed', 'The quoted message changed before sending. Read it again.', 409);
    }
    try {
      const result = await this.request('messages', this.messagePath(prepared.threadId), { method: 'POST', body: JSON.stringify(prepared.payload), signal });
      const id = result?.OriginalArrivalTime ?? result?.originalarrivaltime ?? result?.id;
      if (id == null) throw new Error();
      return { status: 'accepted_by_api', message: { id: String(id), clientMessageId: prepared.payload.clientmessageid }, ...(prepared.channelReplyGuard ? { channelId: prepared.channelReplyGuard.channelId, parentMessageId: prepared.channelReplyGuard.parentMessageId } : {}), deliveryConfirmed: false };
    } catch { throw new AdapterError('send_uncertain', 'The send was attempted. Inspect Teams before further action; reuse this idempotency key.', 409); }
  }
  async reactions(id, mid) {
    const conversation = await this.resolve(id), { raw } = await this.exactMessage(conversation.id, mid);
    const entries = strictArray(raw.properties?.emotions);
    if (entries.some(entry => typeof entry?.key !== 'string' || !Array.isArray(entry.users) || entry.users.some(user => typeof user?.mri !== 'string'))) throw new AdapterError('api_schema_changed', 'Teams reaction metadata changed.', 502);
    const items = entries.map(entry => ({ reaction: entry.key, count: entry.users.length, selected: entry.users.some(user => user.mri === this.selfMri()) }));
    const keys = [...new Set([...presets, ...items.map(item => item.reaction)])];
    return { messageId: mid, items, available: keys.map(reaction => ({ reaction, selected: items.some(item => item.reaction === reaction && item.selected),
      source: items.some(item => item.reaction === reaction) ? 'observed_on_message' : 'adapter_preset', canSet: writableReaction(reaction) })),
      availabilityComplete: false, availabilityCoverage: 'adapter_presets_and_observed_message', coverage: 'api_message_reactions' };
  }
  async prepareMutation(id, mid, kind, input) {
    const conversation = await this.resolve(id); await this.ensureDiscoveredMembership(conversation);
    const found = await this.findMessage(conversation.id, mid);
    if (kind === 'reaction') return { threadId: conversation.id, mid, kind, payload: reactionInput(input.reaction, input.selected), confirmation: this.confirmationContext(conversation.id, reactionInput(input.reaction, input.selected), kind, mid) };
    if (found.message.authorId !== this.selfMri()) throw new AdapterError('not_message_owner', 'Only your own messages can be changed.', 403);
    if (found.message.text !== input.expectedText) throw new AdapterError('message_changed', 'Message text changed; read it again before changing it.', 409);
    if (found.message.deleted) throw new AdapterError('message_deleted', 'This message was already deleted.', 409);
    if (kind === 'edit' && (found.raw.properties?.files && found.raw.properties.files !== '[]' || /schema\.skype\.com\/Reply/.test(found.message.richText || ''))) unsupported();
    const payload = kind === 'edit' ? { ...await this.payload(input), skypeeditedid: mid } : undefined;
    return { threadId: conversation.id, mid, kind, payload, confirmation: this.confirmationContext(conversation.id, payload, kind, mid), expected: found.message.text, version: found.raw.version };
  }
  async mutate(prepared) {
    // Re-read immediately before a write so changed ownership/text is refused.
    // The private endpoint offers no conditional-write guarantee: a remote race remains possible.
    if (prepared.kind !== 'reaction') {
      const current = await this.findMessage(prepared.threadId, prepared.mid);
      if (current.message.authorId !== this.selfMri() || current.message.text !== prepared.expected || current.raw.version !== prepared.version) throw new AdapterError('message_changed', 'Message changed before the API write.', 409);
    }
    let method = prepared.kind === 'delete' ? 'DELETE' : 'PUT', suffix = this.messagePath(prepared.threadId) + '/' + encodeURIComponent(prepared.mid), payload = prepared.payload;
    if (prepared.kind === 'reaction') { suffix += '/properties?name=emotions'; method = payload.selected ? 'PUT' : 'DELETE'; payload = { emotions: { key: payload.reaction, value: prepared.mid } }; }
    try {
      await this.request('messages', suffix, { method, ...(payload ? { body: JSON.stringify(payload) } : {}) });
      return { status: 'accepted_by_api', operation: prepared.kind, messageId: prepared.mid };
    } catch { throw new AdapterError('mutation_uncertain', 'The change was attempted. Inspect Teams and reuse the same idempotency key.', 409); }
  }
}

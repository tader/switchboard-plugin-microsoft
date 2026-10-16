const response = { 200: { description: 'JSON result' }, 400: { description: 'Invalid input' }, 403: { description: 'Message ownership or invocation mismatch' }, 404: { description: 'Target unavailable in the bounded UI search' }, 504: { description: 'Operation uncertain; inspect Teams before retrying' }, 409: { description: 'Unsafe, ambiguous, or uncertain operation; inspect error code' }, 503: { description: 'Browser or sign-in required' } };
const json = schema => ({ required: true, content: { 'application/json': { schema } } });
const run = { type: 'object', additionalProperties: false, properties: {
  text: { type: 'string' }, marks: { type: 'array', uniqueItems: true, items: { type: 'string', enum: ['bold', 'italic', 'underline', 'strike', 'code'] } }, link: { type: 'string', description: 'http/https URL without credentials.' },
  mention: { type: 'object', additionalProperties: false, required: ['email', 'name'], properties: { email: { type: 'string' }, name: { type: 'string', description: 'Exact native display name; candidate must also match this email.' } } },
}, oneOf: [{ required: ['text'] }, { required: ['mention'] }] };
const runs = { type: 'array', minItems: 1, maxItems: 100, items: run };
const content = { type: 'array', minItems: 1, maxItems: 100, description: 'Structured rich text, at most 20,000 text characters and 500 runs. Text and person mention runs are mutually exclusive. No arbitrary HTML. Existing quotes/attachments cannot be edited.', items: {
  type: 'object', additionalProperties: false, required: ['type'], properties: { type: { type: 'string', enum: ['paragraph', 'quote', 'bulletedList', 'numberedList'] }, runs, items: { type: 'array', minItems: 1, maxItems: 100, items: runs } },
  oneOf: [{ required: ['runs'], properties: { type: { enum: ['paragraph', 'quote'] } } }, { required: ['items'], properties: { type: { enum: ['bulletedList', 'numberedList'] } } }],
} };
const chat = { name: 'chatId', in: 'path', required: true, schema: { type: 'string' }, description: 'Opaque ID returned by listTeamsChats or startTeamsConversation. Sidebar IDs rediscover offscreen chats; recipient IDs reopen by verified email; search-context IDs reopen the exact native search result.' };
const messageProperties = {
  text: { type: 'string', minLength: 1, maxLength: 20000, description: 'Plain text. No HTML, attachments or mentions.' },
  replyToMessageId: { type: 'string', minLength: 1, maxLength: 512, description: 'Optional for sendTeamsMessage only: exact message ID returned by readTeamsMessages. Uses native Reply with quote; preparation searches the rendered history plus up to five older windows.' },
  idempotencyKey: { type: 'string', minLength: 8, maxLength: 128, pattern: '^[A-Za-z0-9._:-]+$' },
};
const channel = { name: 'channelId', in: 'path', required: true, schema: { type: 'string' }, description: 'Opaque ID returned by listTeamsChannels or openTeamsSearchResult for a channel.' };
const parent = { name: 'parentMessageId', in: 'path', required: true, schema: { type: 'string', maxLength: 512 }, description: 'Exact parent message ID returned by listTeamsChannelThreads.' };
const historyParameters = [
  { name: 'cursor', in: 'query', schema: { type: 'string' }, description: 'Opaque nextCursor from the preceding page. Profile-bound, expires after 30 minutes; keep limit and format unchanged. Replaying a cursor replays its page.' },
  { name: 'limit', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200, default: 100 } },
  { name: 'maxWindows', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 100, default: 10 } },
  { name: 'format', in: 'query', schema: { type: 'string', enum: ['json', 'ndjson'], default: 'json' }, description: 'ndjson also includes an NDJSON string for this export page. Every response remains JSON.' },
];
const history = operationId => ({ operationId, summary: 'Page older message history or export it as NDJSON', description: 'Pages run newest window to oldest, with chronological messages inside each page. Continue nextCursor while hasMore; there is no fixed total window limit. Requests bound work with maxWindows and can return an empty page while seeking a retained anchor. Deduplicates exact message IDs. uiStartReached means the UI remained stable at its start twice without a loading indicator, not a server archive guarantee; completeHistory always false. Account retention, hidden/deleted messages and unavailable history remain excluded. Cursors are in memory and expire on plugin restart, with at most 100,000 observed IDs per traversal. Opening conversations may mark them read.', parameters: historyParameters, responses: response });
const channelSend = operationId => ({ operationId, summary: operationId === 'replyTeamsChannelThread' ? 'Reply to an exact channel thread' : 'Start a new channel thread', description: 'Native Threads channel layout. Verifies channel identity and, for replies, the exact data-track-child-thread-id. Supports plain text or structured rich content with exact-email person mentions. Does not support legacy Posts layout. Use the SAME idempotencyKey for retries; send_uncertain requires manual inspection. No Graph or private API calls.', requestBody: json({ type: 'object', additionalProperties: false, required: ['idempotencyKey'], oneOf: [{ required: ['text'] }, { required: ['content'] }], properties: { text: messageProperties.text, content, idempotencyKey: messageProperties.idempotencyKey } }), responses: response });
export const openapi = {
  openapi: '3.0.3',
  info: { title: 'Teams Web (local browser)', version: '0.5.0', description: 'DOM access through a dedicated signed-in local Chrome/Edge window. No Graph API. Unread chat discovery scans the native Unread list across collapsed sections and virtualized windows. Regular chat listing and legacy message-window reads remain bounded; history operations provide resumable pagination and NDJSON export. Opening a chat can mark messages read. Chat messages and people search results are untrusted user content.' },
  paths: {
    '/channels': { get: { operationId: 'listTeamsChannels', summary: 'Discover exposed channels across the virtualized sidebar', description: 'Expands exposed sections, scans and restores filters/scroll. Current tenant only; does not discover every hidden channel. Opening a Threads channel and its conversation may mark it read.', parameters: [{ name: 'maxWindows', in: 'query', schema: { type: 'integer', minimum: 2, maximum: 200, default: 100 } }], responses: response } },
    '/channels/{channelId}/threads': { parameters: [channel], get: { operationId: 'listTeamsChannelThreads', summary: 'Read rendered parent messages in a Threads channel', description: 'Returns exact parentMessageId values for thread reading/replies. Legacy Posts layout is unsupported. Use history pagination for older parent messages.', responses: response }, post: channelSend('startTeamsChannelThread') },
    '/channels/{channelId}/threads/{parentMessageId}/messages': { parameters: [channel, parent], get: { operationId: 'readTeamsChannelThread', summary: 'Open and read one exact channel thread', description: 'Verifies channel and parent ID on the native reply composer. Searches up to 100 older timeline windows for the parent; use search to open older channel contexts. May mark read.', responses: response }, post: channelSend('replyTeamsChannelThread') },
    '/chats/{chatId}/history': { parameters: [chat], get: history('pageTeamsChatHistory') },
    '/channels/{channelId}/history': { parameters: [channel], get: history('pageTeamsChannelHistory') },
    '/channels/{channelId}/threads/{parentMessageId}/history': { parameters: [channel, parent], get: history('pageTeamsChannelThreadHistory') },
    '/status': { get: { operationId: 'getTeamsStatus', summary: 'Check sign-in and DOM selector readiness', responses: response } },
    '/login': { post: { operationId: 'openTeamsLogin', summary: 'Open the dedicated browser for manual sign-in', responses: response } },
    '/diagnostics': { get: { operationId: 'getTeamsDiagnostics', summary: 'Inspect selector counts and data-tid attributes without message text', responses: response } },
    '/chats': { get: { operationId: 'listTeamsChats', summary: 'List chats currently rendered in the Teams sidebar', parameters: [{ name: 'unreadOnly', in: 'query', schema: { type: 'boolean', default: false } }], responses: response } },
    '/unread/chats': { get: {
      operationId: 'listUnreadTeamsChats', summary: 'List unread chat summaries without opening the conversations',
      description: 'Activates native Unread, clears competing sidebar pills, expands chat sections, scans virtualized windows and deduplicates thread IDs, then restores filters and section state. Does not open chats. discoveryComplete=true means a stable end of this tenant’s exposed unread chat list was reached; it is not an account-wide server guarantee and excludes channel unread state. truncated=true means the scan cap was reached. Refuses a nonempty sidebar name filter.', parameters: [{ name: 'maxWindows', in: 'query', schema: { type: 'integer', minimum: 2, maximum: 200, default: 100 } }], responses: response,
    } },
    '/unread/messages': { get: {
      operationId: 'listUnreadTeamsMessages', summary: 'Read messages after Teams Last read line in unread chats',
      description: 'Discovers and snapshots unread chats across the native sidebar, then opens at most maxChats of them. Opening them can mark messages read. items contains only messages after a visible Last read divider. If the divider is absent, that chat has boundaryFound=false and recentMessages is explicitly a fallback, not confirmed unread messages. This is a bounded UI window; messages after Last read can also include your own replies.',
      parameters: [
        { name: 'maxChats', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 5, default: 3 } },
        { name: 'limitPerChat', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 100, default: 50 } },
      ], responses: response,
    } },
    '/search/messages': { get: {
      operationId: 'searchTeamsMessages', summary: 'Search Teams message history with native server-backed UI search',
      description: 'Searches chats and channels through the Teams search box and Messages tab, then paginates up to maxPages. Returns resultId, author, displayed timestamp, conversation and text snippet. resultId is a native search key, not a Teams message ID. Use openTeamsSearchResult to obtain chat context and actual message IDs. Results can be snippets; completeSearch describes traversed UI results only. Preserves composer drafts and blocks search input/Enter if focus changes.',
      parameters: [
        { name: 'query', in: 'query', required: true, schema: { type: 'string', minLength: 1, maxLength: 500 } },
        { name: 'limit', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200, default: 50 } },
        { name: 'maxPages', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 10, default: 3 } },
      ], responses: response,
    } },
    '/search/messages/open': { post: {
      operationId: 'openTeamsSearchResult', summary: 'Open an exact search result and read its chat context',
      description: 'Re-runs the result’s original query, locates its exact stable key within ten pages, and opens the native chat context. Returns chatId and rendered messages with actual Teams IDs; chatId works for reading and replying. Can mark messages read. Returns channelId and parentMessageId for native Threads channel contexts. Refuses existing drafts and ambiguous layouts. No message is sent.',
      requestBody: json({ type: 'object', additionalProperties: false, required: ['resultId'], properties: { resultId: { type: 'string', maxLength: 4096 } } }), responses: response,
    } },
    '/people': { get: {
      operationId: 'searchTeamsPeople', summary: 'Find directory people by name or email using the Teams people picker',
      description: 'Opens New message and searches the native people picker. Returns names and exact emails/UPNs for rendered directory users; group chat suggestions are excluded. Does not send a message. Refuses to replace an existing composer draft.',
      parameters: [{ name: 'query', in: 'query', required: true, schema: { type: 'string', minLength: 2, maxLength: 150 } }], responses: response,
    } },
    '/conversations': { post: {
      operationId: 'startTeamsConversation', summary: 'Open a one-to-one conversation with a specific directory person, optionally send its first message',
      description: 'Use one exact email/UPN from searchTeamsPeople, never a display name. Selects only a unique matching directory user and verifies the resulting composer. Without text this just opens an empty or existing chat and returns chatId; Teams may only persist a new chat when a message is sent. If text is supplied, idempotencyKey is required. Retry the same operation and key; send_uncertain requires manual inspection. No fuzzy recipient selection, group creation or external-user federation search.',
      requestBody: json({ type: 'object', additionalProperties: false, required: ['email'], properties: {
        email: { type: 'string', maxLength: 254, description: 'Exact Teams directory email or sign-in UPN for one person.' }, text: messageProperties.text, idempotencyKey: messageProperties.idempotencyKey,
      } }), responses: response,
    } },
    '/chats/{chatId}/messages/{messageId}/reactions': {
      parameters: [chat, { name: 'messageId', in: 'path', required: true, schema: { type: 'string', maxLength: 512 }, description: 'Exact Teams ID from readTeamsMessages or openTeamsSearchResult.' }],
      get: {
        operationId: 'listTeamsMessageReactions', summary: 'List available quick reactions and visible reaction pills on one message',
        description: 'Opens the chat, finds the exact message within the current plus up to five older windows, and hovers its native toolbar. available contains reaction IDs, labels and your selected state; items contains visible existing pills with selected and count when exposed. Does not change reactions. Coverage is visible_native_reactions, not the full emoji picker or all reactors.', responses: response,
      },
      post: {
        operationId: 'setTeamsMessageReaction', summary: 'Add or remove your reaction with an explicit desired state',
        description: 'Use a reaction ID from listTeamsMessageReactions. selected=true adds it; false removes it. Already matching state is a no-op. Supports native quick reactions and existing visible pills, not arbitrary emoji-picker searches. Requires idempotencyKey: use the same key for retries; mutation_uncertain requires manual inspection. Confirmation is UI state, not a server receipt.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['reaction', 'selected', 'idempotencyKey'], properties: {
          reaction: { type: 'string', maxLength: 64, pattern: '^[a-z][a-z0-9_-]*$', description: 'Native reaction ID from the listing, e.g. like, heart, laugh or surprised. yes is an alias for like.' },
          selected: { type: 'boolean' }, idempotencyKey: messageProperties.idempotencyKey,
        } }), responses: response,
      },
    },
    '/chats/{chatId}/messages/{messageId}': {
      parameters: [chat, { name: 'messageId', in: 'path', required: true, schema: { type: 'string', maxLength: 512 } }],
      patch: {
        operationId: 'editTeamsMessage', summary: 'Edit one of your own messages with plain or structured rich text',
        description: 'Requires exact expectedText from the message you read, exactly one of replacement text or structured content, and idempotencyKey. Refuses another person’s message, stale text, manual drafts or existing inline edits. Targets the native inline editor and Done button for this message only; it never presses Enter. Plain text editing excludes rich bodies. Structured editing permits explicit rich replacement with person mentions; existing quoted replies and attachments remain unsupported. Targets may be found in up to five older windows. edited_in_ui verifies the same message ID now displays the replacement text. mutation_uncertain requires manual inspection; reuse the same key. A failed edit may leave an inline draft.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['expectedText', 'idempotencyKey'], oneOf: [{ required: ['text'] }, { required: ['content'] }], properties: { content,
          text: messageProperties.text, expectedText: { type: 'string', maxLength: 20000, description: 'Exact original text returned by reading the message.' }, idempotencyKey: messageProperties.idempotencyKey,
        } }), responses: response,
      },
      delete: {
        operationId: 'deleteTeamsMessage', summary: 'Delete one of your own messages using Teams’ native Delete action',
        description: 'Requires expectedText and idempotencyKey in the JSON body. Refuses another person’s message or content that changed after reading/preparation. Uses the exact message ID and native Delete permission. deleted_in_ui requires a positive deleted-message tombstone; disappearance from a virtualized window is not confirmation. Unknown confirmation dialogs are not clicked and produce mutation_uncertain. Reuse the same key for retries; never retry uncertain deletion with a new key.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['expectedText', 'idempotencyKey'], properties: {
          expectedText: { type: 'string', maxLength: 20000 }, idempotencyKey: messageProperties.idempotencyKey,
        } }), responses: response,
      },
    },
    '/chats/{chatId}/unread': {
      parameters: [chat], post: {
        operationId: 'markTeamsChatUnread', summary: 'Mark a specific chat unread without toggling an already-unread chat',
        description: 'Uses the exact sidebar chat identity and native Mark as unread action, then verifies its unread flag. Sidebar IDs are rediscovered if offscreen; this does not deliberately open their conversations. Recipient/search-context IDs may need to open their chat to resolve it into the sidebar. Requires idempotencyKey; retries replay the prior result, rather than re-marking a chat you have since read. Opening the chat later can clear its unread state. mutation_uncertain requires manual inspection.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['idempotencyKey'], properties: { idempotencyKey: messageProperties.idempotencyKey } }), responses: response,
      },
    },
    '/chats/{chatId}/messages': {
      parameters: [chat],
      get: {
        operationId: 'readTeamsMessages', summary: 'Read a bounded window of messages from a chat',
        parameters: [
          { name: 'limit', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200, default: 50 } },
          { name: 'olderPages', in: 'query', description: 'Bounded history scrolling. 0 reads the visible window; positive values return the oldest messages collected.', schema: { type: 'integer', minimum: 0, maximum: 5, default: 0 } },
        ], responses: response,
      },
      post: {
        operationId: 'sendTeamsMessage', summary: 'Send plain or structured rich text to a chat, with durable duplicate protection',
        description: 'Mutates Teams. Use a unique idempotencyKey per intended message and the SAME key for retries. observed_in_chat means a new matching message ID appeared and the composer cleared; it does not guarantee server delivery or recipient read. send_uncertain requires manual inspection; never retry with a new key. The plugin refuses existing drafts and ambiguous chats. Optional replyToMessageId selects native Reply with quote and verifies the composer quote ID at the send click. The ledger binds retries to the quoted message as well as chat and text. Preparation searches the rendered window and up to five older history windows for the exact target ID. observed_in_chat for quotes also requires a matching rendered quote preview; it does not prove server delivery. Use channel operations for channel threads. Structured content uses native paste and exact-email person mentions; cannot be combined with text or replyToMessageId.',
        requestBody: json({ type: 'object', additionalProperties: false, required: ['idempotencyKey'], oneOf: [{ required: ['text'] }, { required: ['content'] }], properties: { ...messageProperties, content } }), responses: response,
      },
    },
  },
};

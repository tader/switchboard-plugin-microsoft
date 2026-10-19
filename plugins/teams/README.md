# Teams in the Microsoft plugin family

Install `plugins/teams` from `tader/switchboard-plugin-microsoft`. The plugin and service ID remain `teams-web`, so existing connections, API routes, saved calls, profile directories and durable send/read-state ledgers continue to work. The complete current API adapter and convenience methods are retained.

For one Microsoft sign-in, first connect **Microsoft Authentication** (Teams web sign-in), then connect Teams using **Use Microsoft Authentication** and select that source. Select the same source for Outlook, Graph, OneDrive, To Do and Excel. For a server, explicitly share authentication from the laptop to the direct peer and select the received source on a local Teams connection. Tokens renew by HTTP with cached credentials while the laptop is offline. A fresh provider-required sign-in happens on the authentication source device.

Teams setup offers **Use Microsoft Authentication**, with the same authentication connection selector as the other Microsoft products. A Switchboard version supporting authentication sources and hidden legacy methods is required. Existing **Microsoft sign-in through local Chrome** and **Import Teams tokens** connections continue to work, but these legacy choices are hidden from setup. Reconnect an existing connection and select **Use Microsoft Authentication** to migrate it in place. Outlook-issued web credentials cannot authorize Teams chat even when token issuance succeeds; reconnect the source using its default Teams sign-in. All data operations use APIs. Browser automation is used for authentication only.

To diagnose authentication without reading Teams content, call `getTeamsStatus` with `renew=true`. This renews tokens and returns status only. If a consumer has an expired session, it first obtains the authentication source’s current cached session; another browser sign-in is requested only if that session is also rejected.

Legacy standalone documentation and scripts follow for compatibility:

# Teams API plugin for Switchboard

Version 0.11.0 supports direct Teams private API calls with local browser sign-in or imported tokens. Chrome or Edge captures renewal credentials during normal Teams sign-in, then closes. Access tokens renew through HTTP using the existing Teams public client ID, without a client secret or browser. A central Switchboard can import tokens from another machine and run without Chrome or Edge. No Go, Electron, Microsoft Graph registration, npm dependencies, or persistent Teams tab is required. Node 24+ is required.

For unread triage, prefer `getTeamsTriageInbox`, `getTeamsReplyContexts` and `replyTeamsBatch`. The inbox groups chats and channel threads; native quoted chat replies have durable duplicate protection. `setTeamsReadState` marks chat messages read or unread at an exact selected boundary. **Exact-root channel replies are supported; channel read state and automatic read-after-reply remain disabled.** See [triage protocol status](docs/triage-protocol.md).

The service ID remains `teams-web`, and the existing browser authentication method remains `browser`, so existing Switchboard connections can be reconnected in place. The previous DOM adapter is retained in `legacy-index.js` for regression testing; production does not invoke it.

## Install locally

From this checkout:

```sh
npm test
npm run install:local -- ~/switchboard/.data
```

The installer stages code outside the plugin watcher, replaces `.data/plugins/teams-web`, and saves the previous code under `.data/plugin-backups/`. It does not change connection records, encrypted credentials, or `.data/plugin-data/teams-web`. Reload or restart Switchboard after installing. This branch has not been published as a GitHub release.

In Switchboard, reconnect the existing Teams connection once (or add a new Teams connection). Choose Microsoft sign-in through local Chrome, or Import Teams tokens for a browser-free host (see below). For browser sign-in, enter your work email as the optional Microsoft work email hint. Complete sign-in in the dedicated browser window; it closes automatically when authentication finishes. `LOCAL-BROWSER` is only a placeholder in the connection dialog.

Old DOM-only connections report `signin_required` until reconnected. Refresh chat/channel IDs by listing them again; exact legacy thread IDs can migrate after a membership check, but UI search, person and history cursors cannot. Existing send ledger entries remain preserved; a reused key with changed request semantics is rejected.

## Central Switchboard without Chrome or Edge

On a machine with Node 24+ and Chrome or Edge, run from the plugin checkout:

```sh
npm run auth:export -- --out ~/teams-tokens.json --login-hint your.name@example.com
```

Complete Microsoft sign-in in the dedicated browser window. The script closes the browser and writes a version 2 JSON bundle containing access and refresh tokens with owner-only permissions (`0600`). It requires a new output filename, refuses existing files and symlinks, removes incomplete output on failure or interruption, and never prints tokens. It does not require a local Switchboard or send Teams messages. `TEAMS_BROWSER_PATH` overrides browser detection; `TEAMS_LOGIN_HINT` is an alternative to `--login-hint`.

Install this plugin version on the central Switchboard. Add Teams or reconnect an existing connection, select **Import Teams tokens**, enter a label and paste the complete file contents into the secret bundle field. The plugin checks bundle structure and resource/account metadata, proves the refresh token by obtaining both access tokens over HTTP (even when the bundled access tokens are fresh), then validates those access tokens through Teams: the auth service accepts the Skype token and confirms its account, and the conversation service accepts the chatsvcagg token. Microsoft validates their proprietary signatures; see [Microsoft access-token guidance](https://learn.microsoft.com/entra/identity-platform/access-tokens). No credentials are accepted before both checks succeed. Switchboard encrypts accepted credentials; a failed import leaves an existing session intact. Delete the plaintext export file after import. The JSON contains bearer credentials; browser cookies remain on the authentication machine.

The central host never opens or discovers a browser for imported-token connections. Before data calls, access tokens approaching expiry renew over HTTP; rejected access tokens can trigger forced HTTP renewal. The plugin follows refresh-token rotation and returns updated credentials to Switchboard for encrypted storage, including when a later renewal step fails. After a restart, it uses the saved refresh token. If Microsoft expires or revokes renewal credentials, export a replacement and reconnect the same central connection. Existing version 1 bundles still work while their access tokens remain fresh, but require a new version 2 export to enable HTTP renewal. Tenant/account IDs must match; reconnection preserves the profile and durable send ledger. A different account requires a separate connection.

Silent authentication can reuse the saved local browser profile when Microsoft permits it. Use a new output filename:

```sh
npm run auth:export -- --silent --out ~/teams-tokens-renewed.json
```

If Microsoft requires sign-in, repeat without `--silent`. The export profile is separate from diagnostic and Switchboard profiles. Existing browser connections can switch to token import through reconnect; selecting browser authentication again requires a browser on that Switchboard host.

`getTeamsStatus` reports `authMode="tokens"`, expiry and `tokenImportRequired`. `renewalAvailable` and `renewalMethod="http_refresh_token"` confirm refresh support. Terminal refresh-token rejection produces `token_import_required` with re-import instructions; transient renewal failures remain retryable. `openTeamsLogin` returns those instructions instead of opening a browser. Status and capability calls remain available after expiry. A real exported bundle has passed read-only authentication, account binding and conversation validation on the authentication machine. Verify import on your central host too; network access and host compatibility remain environment-specific.

## Sync tokens from your laptop

Set `SWITCHBOARD_TOKEN` to a full-access Switchboard API token, then run from this checkout on your laptop:

```sh
npm run auth:sync -- --url https://switchboard.example.com --connection your-teams-connection --login-hint your.name@example.com
```

The connection can be its Switchboard name or ID. The command checks auth expiry and readiness. If tokens have more than two minutes remaining, it exits without a browser or upload. Otherwise it first makes a read-only chat-list request so Switchboard can attempt its existing HTTP renewal; returned chat data is discarded. If fresh sign-in is required, it acquires tokens using a dedicated local browser profile, closes the browser, sends the version 2 bundle through Switchboard's reconnect API, and verifies readiness. Reconnect preserves the connection and ledger and validates the same Microsoft account. The account label is retained. Bundles stay in memory; no plaintext export file is created or tokens printed.

The command tries the saved browser session headlessly first, then opens a sign-in window if Microsoft requires interaction. Use `--silent` for scheduled runs: it exits with an error instead of opening an interactive window. Use `--force` to acquire and upload replacement credentials even when server tokens are healthy. An existing sign-in in progress is never replaced. Transient server renewal failures stop without browser acquisition; rerun later. Uploads are not automatically retried.

You can set `SWITCHBOARD_URL`, `TEAMS_CONNECTION` and `TEAMS_LOGIN_HINT` instead of supplying those flags. With these variables and `SWITCHBOARD_TOKEN` configured, the command is simply:

```sh
npm run auth:sync
```

`TEAMS_BROWSER_PATH` selects Chrome or Edge. Each server/connection has its own local profile under `.local-profile/token-sync/`. The server needs the Teams plugin version 0.8.0 or later. It can run without a browser. Remote URLs require HTTPS; HTTP is accepted for localhost. Switchboard's reconnect API requires full access, so connection-limited or MCP-only tokens cannot perform this task. The command reports access/chat-token expiry; it does not infer a refresh token's remaining lifetime from capture time. Silent browser recovery and server connectivity depend on your environment and tenant policy.

## Test without sending messages

```sh
npm test
npm run smoke
TEAMS_LOGIN_HINT=Thomas.De-Ruiter@dsm-firmenich.com npm run check:api
TEAMS_LOGIN_HINT=Thomas.De-Ruiter@dsm-firmenich.com npm run check:api -- --silent
```

`npm test` runs unit and adapter tests, including the retained DOM regression tests. `npm run smoke` runs the new API adapter and authentication tests with fixtures; neither command contacts Teams or sends real messages.

`check:api` reuses the shared diagnostic session, renews expired tokens over HTTP, then reads your profile, conversation list and five recent messages from one chat. It prints only PASS markers and token expiry, never tokens or message content. The second command reuses the same session without interactive sign-in. `check:api`, `check:triage` and `check:auth-refresh` share `.local-profile/token-export/diagnostic-session.json` and its persistent browser profile. The first successful run saves the session; subsequent checks reuse it and renew over HTTP. The session file contains credentials, is owner-readable/writable only (0600), and is excluded from Git. These checks do not create a Switchboard connection. Use `--capture` only when explicitly testing a fresh browser capture. Set `TEAMS_BROWSER_PATH` to override browser detection. A successful silent check demonstrates current session reuse, not a guarantee about future company sign-in policy.

After installing, run `getTeamsStatus`, `getTeamsCapabilities`, `listTeamsChats` and `readTeamsMessages` in Switchboard. Chrome should remain closed for data calls with fresh tokens. Verify an intended write manually in a designated test chat before depending on write operations. Authorized live chat writes and exact-message readback passed on 2026-10-08; temporary messages were deleted. See [protocol evidence](docs/parity-protocol.md).

## Authentication lifetime

### Read-only refresh-token compatibility check

The production acquisition and HTTP renewal helpers also have an isolated diagnostic:

```sh
npm run check:auth-refresh -- --login-hint your.name@example.com
```

Complete normal Teams sign-in in the dedicated browser. With `--capture`, the diagnostic captures a refresh token only from a successful Microsoft token exchange for the existing Teams web client, closes the browser and waits for its process to exit, then requests fresh Skype and chatsvcagg access tokens using HTTP, with no client secret or cookies. It follows refresh-token rotation between the two requests. It validates the renewed credentials through the Teams chat-token exchange, confirms the account through a profile read, and reads the conversation list. It prints only PASS markers, sanitized errors and access/chat-token expiry; it does not print tokens or message content, export a bundle, or send messages.

`--login-hint` identifies the account to validate through Teams after renewal; it does not prefill the Teams web login. `TEAMS_LOGIN_HINT` is an alternative. `TEAMS_BROWSER_PATH` overrides browser detection. A later `npm run check:auth-refresh -- --silent --login-hint your.name@example.com` attempts capture headlessly using the saved diagnostic session; success depends on Microsoft session and tenant policy.

The browser profile is separate at `.local-profile/refresh-check/auth-browser`. Before each check, only Teams app caches in that profile are cleared to force a network token exchange; Microsoft login cookies are retained. The diagnostic keeps captured credentials in memory, but Teams itself can cache credentials in this browser profile. Protect it like a logged-in browser session.

By default, this command reuses the shared diagnostic session and forces an HTTP renewal. Add `--capture` to repeat the browser-capture compatibility probe. Neither mode changes existing Switchboard connections. Fixture success is not live compatibility evidence. A successful live run establishes capture, browser-closed HTTP renewal, account binding and the two API reads for that account on that machine. It does not establish refresh-token lifetime, recovery after its expiry, cookie-only renewal, or central-host compatibility. A captured rotated token can inherit an earlier SPA expiry; capture time is not a new 24-hour lifetime.

On 2026-10-08, this diagnostic passed real browser capture after Microsoft Authenticator approval, browser-process exit, HTTP renewal of both Microsoft access tokens, chat-token exchange, account binding, profile and conversation reads on the user's Mac. Subsequent silent diagnostics reused the saved credential cache and forced HTTP renewal without opening a browser. Headless browser reacquisition after credential expiry has not passed. See [validation status](ROADMAP.md).

Tokens have individual expiry timestamps. Before a data call, the plugin renews when less than two minutes remain. An expired chat-service token can be exchanged over HTTP while its source token remains valid. New browser connections and version 2 imported bundles renew Microsoft access tokens over HTTP. Existing browser connections acquire refresh support on reconnect or their next successful browser authentication. Browser connections attempt saved-session browser reacquisition only after terminal refresh rejection; imported connections require re-export instead.

Microsoft documents a 24-hour limit for refresh tokens issued to SPA redirect URIs; rotated tokens inherit the original expiry. Rotation does not provide indefinite browser-free authentication. Tenant policy, MFA and revocation can shorten availability. Capture time is recorded as metadata, not a promised expiry. See [Microsoft refresh-token lifetime](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens). Cookie-only recovery is not implemented.

There is no fixed manual-login interval: Microsoft session expiry, MFA, revocation and tenant policy decide when silent renewal stops working. `getTeamsStatus` reports `tokenExpiresAt`, readiness and sign-in errors. Use `openTeamsLogin` and poll status, or reconnect in Switchboard, when interactive sign-in is required. Renewal occurs on demand, not on a background schedule. Control/status calls remain available when tokens expire.

## Operations and migration limits

| Operation | API version |
| --- | --- |
| Chat/channel listing, recent chat messages | Supported |
| Unread chats/messages | Supported, bounded CSA/regional chat discovery and consumption horizon |
| Chat/channel history, NDJSON | Supported, server backward links and bounded cursors |
| People search | Names through server suggestions; exact email/UPN verified separately |
| Start conversation operation | Resolves an existing or account-bound virtual one-to-one chat; first send persists a new chat |
| Chat sends and edits | Text with automatic URL links, structured formatting and exact-email person mentions |
| Delete own chat message; add/remove reaction | Supported |
| Grouped unread triage and batch context | Chats and channel threads; bounded, paginated coverage |
| Native quoted chat replies | `replyTeamsBatch` or ordinary sends with exact `replyToMessageId` |
| Channel roots, exact thread reads/history | Direct restored routes, cursor replay and NDJSON; also triage context |
| Channel replies | Exact-root replies with durable replay and root-bound verification; live send validation pending |
| Automatic read-after-reply | Implemented behind a disabled protocol-verification gate |
| Selected-message read/unread | `setTeamsReadState` supports chat read/unread; channel gates defer |
| Message search/open result | Server-backed queries and exact account-bound result context |
| Attachments and starting channel threads | Unsupported |

Unsupported operations are removed from the advertised OpenAPI schema; saved calls fail explicitly. `getTeamsCapabilities` reports the supported subset. Conversation data can be partial, and retained history is not a guaranteed complete archive. API reads do not send read markers.

## Search and direct message context

`searchTeamsMessages` searches the server across chats and channels, bounded by `limit` (1–200) and `maxPages` (1–10). It deduplicates hits and reports truncation. `resultId` is an account-bound handle valid for 30 minutes or until reload/capacity eviction. Use `openTeamsSearchResult` to get current `selectedMessage`, `chatId` or `channelId`/`parentMessageId` and recent context. Hits with no exact locator report `canOpen=false`. Deleted/unavailable results fail instead of selecting another message. Reads do not mark messages read.

Name-based people search and message search acquire an optional Substrate access token over HTTP from existing renewal credentials; no sign-in window is opened. Version 1 bundles without renewal credentials return `search_token_required`; reconnect with a version 2 bundle. Rotation is returned to Switchboard for encrypted persistence before the operation runs. Search-token failures do not silently fall back to scanning recent messages.

`startTeamsConversation` accepts one exact directory email. With no text, it returns an existing chat or a virtual one-to-one chat with `persisted=false` and sends nothing. The first message persists a new chat; use an idempotency key. The virtual ID binds the account and re-resolved recipient and works after reload. This does not create groups or search external federation.

Ordinary `sendTeamsMessage` accepts `replyToMessageId` from an ordinary read or opened search hit. It reads the exact message directly and rechecks content/version immediately before dispatch. Structured content accepts person mention runs; the exact email is resolved again, and server directory names/MRIs supply escaped markup and mention metadata. These sends, triage replies and edits share mention handling. Existing quoted messages and attachments cannot be edited.

`replyTeamsChannelThread` sends text or structured content beneath an exact
`channelId`/`parentMessageId`. It supports mentions and automatic URL links,
rechecks membership, archived/disabled status and root content/version before
dispatch, and preserves durable acceptance/replay. `verification.threadConfirmed`
requires matching root metadata in readback; acceptance alone is not proof of
thread placement or delivery. `replyTeamsBatch` also supports channel targets.
Channel read markers and automatic read-after-reply remain disabled. The send
contract is source-verified and fixture-tested; designated live send validation
remains pending.

`listTeamsChannelThreads` pages channel roots independently of unread state. `readTeamsChannelThread` reads the exact root and composite reply conversation. `pageTeamsChannelThreadHistory` supports backward-link cursors, replay, overlap deduplication and NDJSON. Cursors bind account, channel/root, limit and format. Each thread response includes its root; the initial page can contain that root in addition to `limit` replies. Continue while `hasMore`, including empty filtered root pages. `completeHistory=false` reflects retention/unavailable history.

Live checks passed message/people search, chat result opening, exact message reads, channel roots/continuation and exact thread/NDJSON reads. Mention notifications, first-send chat persistence, quoted sends, channel search-hit opening and deep reply pagination remain separate live checks. No real messages were sent.

## Unread triage and replies

1. Call `getTeamsTriageInbox`. Defaults: 10 groups, five context messages and a 50-message unread window. Mentions come first, then direct chats, then other conversations; oldest unread first within each bucket. Priority applies to scanned sources. Follow `nextCursor` with unchanged limits while `hasMore`, including after an empty page. Each scan bounds work to 20 sources and 20 roots per channel; missing flags/horizons and unavailable sources appear in `coverage.issues`.
2. Review each group's `unread`, `context`, factual `signals` and `canReply`. Messages are not duplicated across unread/context arrays. Own messages, deleted messages and system events are excluded from unread results. Read-only calls leave everything unread. The plugin does not decide urgency or generate wording.
3. If needed, call `getTeamsReplyContexts` with up to ten reply targets and optional `contextLimit` (default 20, maximum 50). It refreshes context and returns new reply targets on individual messages. The inbox group-level target chooses the latest incoming message; a context refresh preserves the selected message as its group-level target.
4. Call `replyTeamsBatch` with up to ten explicit replies. Use a unique key per item and retain the exact input for retries:

```json
{"replies":[{"target":"<replyTarget from inbox>","text":"Thanks, I'll review this today.","idempotencyKey":"review-reply-20261008-001"}]}
```

Each result includes independent `send` and `read` outcomes. `context_changed` means fetch fresh context and reconsider the reply. Sends are sequential; authentication/rate/transport failures or an uncertain send stop remaining items as `not_attempted`. Reuse the original key for accepted/uncertain retries, never invent a replacement key to bypass protection. Accepted sends replay after restart even when reply targets have expired. Content conflicts are refused.

Reply targets and cursors are account-bound, expire after 30 minutes or plugin reload, and may be evicted at the 1,000-record capacity. A refreshed target identifies a new operation; retain the old target/key when checking an already attempted operation.

Currently successful replies return `read.status="deferred"`, `reason="protocol_unverified"`. The durable read stage is implemented and tested: after protocol verification it can retry read failures without resending, preserves later arrivals/newer known horizons, and defers channel-wide advancement while sibling threads remain unhandled or coverage is incomplete. An old thread root never proves coverage of omitted unread replies. No automatic marking is enabled in this release.

`setTeamsReadState` (`POST /triage/read-state`) accepts 1–10 explicit updates. `read` means read through the selected message; `unread` leaves that message and later messages unread. Use an individual message's `replyTarget` as `target`, or an exact `conversationId` and `messageId`; channel locators also require `parentMessageId`.

```json
{"updates":[{"target":"<message replyTarget>","state":"unread","idempotencyKey":"review-unread-001"}]}
```

Keep the original input and key for retries. Results contain `status` (`updated`, `unchanged`, `deferred`, `uncertain`, `failed` or `not_attempted`) and the input locator/key. Successful results identify `throughMessageId` or `fromMessageId` and `observedReadBoundary`; completed retries add `replayed=true`. Uncertain writes stop the batch and are never resent. Explicit unread intents durably supersede older automatic read grants, including across restarts. Incomplete read prefixes and unhandled sibling channel threads defer.

Check `capabilities.triage.readState.chat.read`, `.chat.unread`, `.channel.read` and `.channel.unread`, or each group's `canMarkRead`/`canMarkUnread`. **Chat read and unread are enabled and live-verified; both channel gates remain false** and return `deferred/protocol_unverified`. The transport keeps the ordinary horizon monotonic and changes a separate bookmark for explicit unread intent. Reading a prefix advances that bookmark instead of clearing later unread messages. A cleared zero bookmark falls back to the ordinary horizon. Private endpoints provide no atomic guard against concurrent external changes. See [native capture and remaining verification](docs/triage-protocol.md).

Run `npm run check:triage -- --silent` for a read-only live check using the saved export browser session, or omit `--silent` for interactive authentication. It reuses the shared diagnostic session, prints counts and sanitized errors, sends no messages, and writes no read markers.

## Writes and history

Every write requires an `idempotencyKey`. Use one new key per intended action and reuse it for retries. The durable ledger records intent before the HTTP write; there are no automatic write retries. `accepted_by_api` means server acceptance. After saving that receipt, the adapter performs at most two exact-message reads within five seconds. `verification.status` is `observed`, `not_observed`, or `unavailable`; observation confirms matching server state, never recipient delivery. Reusing the same key can retry an unfinished readback without another write. Once observed, the saved observation replays; it does not promise the message is still unchanged. If `send_uncertain` or `mutation_uncertain` occurs, inspect Teams before further action; retrying with a new key can duplicate the action.

Edits and deletion require `expectedText`, exact message IDs and ownership. The adapter rechecks text and version immediately before writing, but the private endpoint does not provide an atomic conditional-write guarantee. Edits refuse existing mentions, quoted replies and attachments. Mutation target lookup searches at most six pages of 200 messages. Reaction listing uses an exact-message read. Its `available` entries identify adapter presets and IDs observed on the message; `availabilityComplete=false` makes the incomplete catalog explicit. `canSet` means the ID fits the adapter input contract, not that the server grants permission or supports it for this account.

For formatted sends/edits, use `content` instead of `text`, for example:

```json
{"content":[{"type":"paragraph","runs":[{"text":"Hello","marks":["bold"]}]}],"idempotencyKey":"example-unique-key"}
```

Bare `http://` and `https://` URLs automatically become HTML links in ordinary
text and structured text runs, consistently across sends, replies and edits.
Trailing sentence punctuation and unmatched closing brackets stay outside the
link; balanced parentheses in URLs are retained. Text without eligible URLs
keeps the plain-text payload. URLs with credentials, backslashes, malformed
addresses or more than 2,000 characters remain literal. Explicit links and runs
marked `code` are preserved. No Markdown or caller HTML is interpreted.

For custom labels or URLs with ambiguous trailing punctuation, use an explicit
link run such as `{"text":"View the PR","link":"https://github.com/tader/switchboard-plugin-teams/pull/4"}`.

Supported blocks: paragraph, quote, bulletedList and numberedList. Marks: bold, italic, underline, strike and code; links must use HTTP(S). Structured quotes are formatting, not replies to another message. Person mention runs require exact directory emails; arbitrary HTML is refused. Message HTML returned by reads is untrusted data.

History cursors are bound to the connection, conversation, page size and format. They expire after 30 minutes and reset on plugin reload. Reusing a cursor repeats its page. Keep `limit` and `format` unchanged while following `nextCursor`; pages proceed from newer to older history with chronological records inside each page. Capacity is 1,000 cursor records and 100,000 observed IDs per traversal. `format=ndjson` adds an export string to each JSON response.

Unread message reads select messages newer than the conversation consumption horizon and exclude your own messages. If the horizon is absent, `items` is empty and a separate `recentMessages` fallback is returned. These are bounded recent windows; `completeAccount=false` remains explicit.

## Storage and validation

Switchboard encrypts API access and refresh tokens as connection credentials for both authentication methods. The plugin holds decrypted tokens only in memory. For browser authentication, Chrome retains Microsoft session cookies in its separate per-connection `auth-browser` profile under the plugin data directory. Protect that directory like a logged-in browser profile. Disconnecting a connection deletes its profile and ledger; installation preserves them. Switchboard's normal call audit still applies. The ledger can contain request results and message text.

The normal Teams web app performs its authorization-code/PKCE flow. The plugin observes only successful HTTPS token responses for the fixed Teams client and allowed Teams origins, then checks resource, tenant, account and expiry consistency on renewed tokens. JWT decoding is a metadata check; the receiving Microsoft services validate the access tokens. The Teams chat-token exchange confirms the account, and token imports also require a successful conversation-service read. The retained legacy OAuth helper has state, nonce and ID-token signature regression coverage. HTTP calls use fixed Teams service hosts and the auth-service regional chat host; history links cannot change host or conversation. The loopback adapter exposes short-lived single-use invocation capabilities, never Microsoft tokens. Diagnostics omit tokens and message contents.

The earlier Go compatibility probe was live-validated for OAuth, profile, conversation and recent-message reads. This Node implementation has automated fixture coverage for authentication, renewal, portable token export/import, restart recovery, account isolation, request routing, pagination, token redaction, write payloads and duplicate protection. A real portable bundle passed read-only token exchange, account binding and conversation validation on the authentication machine. Central reconnect upload, renewal across the SPA expiry boundary and saved-session browser recovery still need validation. Authorized chat sends, quotes/mentions, edit/delete and reaction readback passed; recipient notification and new-chat persistence remain unverified. Teams private endpoints can change independently of this plugin. See [API sources](THIRD_PARTY.md).

This implementation and its tests were written by OpenAI Codex at Thomas de Ruiter's request. Independent code review remains outstanding.

### Conversation discovery

`listTeamsChats` combines the CSA snapshot with up to ten regional pages of 100
conversations within a 30-second scan budget. A partial CSA snapshot triggers
the same fallback for triage. Existing CSA metadata takes precedence; new chat
and meeting IDs are deduplicated and unread state remains unknown when native
horizons/bookmarks are unavailable. `discovery` reports pages, added chats and
incomplete/error coverage. Team/channel discovery still uses CSA.

For larger inventories, use `pageTeamsChatDiscovery` with `limit` (1–100) and
follow `nextCursor` while `hasMore=true`, including after an empty page. Keep
limit unchanged; cursors bind account and expire after 30 minutes or reload.
Replay returns the same page, overlapping IDs are deduplicated and cycles/caps
report truncation. Notes, spaces and system streams are excluded. An ended
regional traversal does not establish a complete archive. Exact selected chats
outside the automatic scan can resolve through current roster and metadata
reads; newly discovered chats require a fresh roster check before message writes.

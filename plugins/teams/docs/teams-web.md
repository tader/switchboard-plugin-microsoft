---
title: Connecting Teams
services: [teams-web]
---

# Connecting Teams

Teams 0.10.0 uses direct APIs with local browser sign-in or imported tokens. Browser sign-in captures a refresh token in Chrome or Edge and closes the browser. Access tokens then renew over HTTP without a client secret. Imported-token connections use no browser on Switchboard.

For local browser authentication, in Switchboard Connections, add Teams or reconnect your existing Teams connection once after upgrading from 0.5. Enter an account label and optionally your work email. Complete Microsoft sign-in in the dedicated browser window. `LOCAL-BROWSER` is a placeholder, not a code to enter at Microsoft. Each connection retains a separate browser session and encrypted API credentials.

Call `getTeamsStatus`, then `listTeamsChats` and `readTeamsMessages`. Chrome stays closed during HTTP renewal. The plugin renews automatically before data calls when tokens approach expiry. After terminal refresh rejection, browser connections attempt saved-session browser authentication. If Microsoft requires interactive authentication, use `openTeamsLogin` and poll `getTeamsStatus`, or reconnect through Switchboard. Your tenant policy determines how often this happens.

For unread triage, use `getTeamsTriageInbox` to obtain grouped unread messages and reply targets. Use `getTeamsReplyContexts` for extra context, then `replyTeamsBatch` for explicit native quoted chat replies. Keep each reply's input and idempotency key for retries. These calls report partial coverage and preserve read state. Channel replies and automatic read updates are disabled pending live protocol verification; successful chat replies report a deferred read update. Check `getTeamsCapabilities` and [protocol status](triage-protocol.md).

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

Microsoft limits SPA refresh tokens to 24 hours; rotations inherit the original expiry. Re-export may therefore be needed daily or sooner under tenant policy. Cookie-only recovery is not implemented. See [Microsoft refresh-token lifetime](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens).

## Update central tokens from your laptop

Set `SWITCHBOARD_TOKEN` to a full-access Switchboard API token. From the plugin checkout on your laptop, run:

```sh
npm run auth:sync -- --url https://switchboard.example.com --connection your-teams-connection --login-hint your.name@example.com
```

It checks expiry, tries server-side HTTP renewal first, and acquires and uploads a replacement bundle through reconnect only when fresh authentication is needed. A read-only chat-list request can trigger server renewal; its data is discarded. Tokens stay in memory, and the connection and ledger are preserved. Saved-session browser authentication is attempted first; a sign-in window opens if required. Add `--silent` for scheduled runs or `--force` to replace healthy credentials. An in-progress sign-in is never replaced. Set `SWITCHBOARD_URL`, `TEAMS_CONNECTION` and `TEAMS_LOGIN_HINT` to omit their flags. See the checkout README for details. The server needs plugin 0.8.0 or later.

## Operations

Use `getTeamsCapabilities` for the current operation subset. Chat and channel listing, unread windows, chat messages, paged chat/channel history, exact-email lookup, existing one-to-one chat resolution, plain/formatted chat sends, edits, deletion and reactions are implemented. Name search, message search, new chats, mentions, attachments and channel-thread writes are not yet migrated. Native quoted chat replies and bounded channel-thread reads are available through triage. Refresh old conversation IDs through the listing operations. API reads do not send read markers.

For every send or mutation, use a unique `idempotencyKey` and reuse it for retries. `accepted_by_api` is server acceptance, not a delivery receipt. On `send_uncertain` or `mutation_uncertain`, inspect Teams before further action; using a new key can duplicate the action. Edits/deletion require your own exact message ID and current `expectedText`. Concurrent external changes cannot be protected atomically by the private API. Live writes still require validation in a designated test chat.

Person lookup requires an exact email or UPN. `startTeamsConversation` only resolves an existing one-to-one chat; it does not create one. Optional text sends to that resolved chat. History follows `nextCursor`; keep limit/format unchanged. Cursors expire after 30 minutes and on plugin reload. Unread reads are bounded windows, not an account-wide archive. Rich HTML returned by reads is untrusted data.

For local browser authentication, the browser profile persists in plugin data. Both methods retain a send ledger in plugin data. Protect this directory: it contains a Microsoft browser session and may contain message results. Disconnecting deletes both. Token credentials are encrypted by Switchboard; diagnostics do not expose them. This version was written by OpenAI Codex at Thomas de Ruiter's request. The original compatibility probe passed real reads; the new Node adapter is covered by fixture tests and includes a read-only `npm run check:api` diagnostic.

`setTeamsReadState` accepts explicit selected-message read/unread updates using a triage target or exact conversation/message locator. Read includes the selected message; unread includes it in the unread suffix. Each update requires a durable idempotency key. Check `capabilities.triage.readState` for separate chat/channel read/unread gates; chat read/unread are enabled and live-verified; channel gates return `deferred/protocol_unverified`. The transport preserves a later ordinary horizon while advancing an explicit unread bookmark only through the selected message. Native requests and browser-closed readback passed in a user-designated regular chat. Private endpoints have no atomic guard against external changes. See [protocol status](triage-protocol.md).

Live diagnostic checks (`check:api`, `check:triage`, `check:auth-refresh`) share `.local-profile/token-export/diagnostic-session.json`. The first successful run saves credentials in this Git-ignored, owner-only file; subsequent checks reuse them and renew over HTTP. `--silent` prevents interactive sign-in. `--capture` explicitly repeats browser capture. Run diagnostics sequentially so token rotations remain consistent.

# Shared Messaging v1

The `teams-web` service implements the optional shared Messaging contract beside
its existing provider API. Enable shared contracts on a compatible Switchboard
core to discover it. No provider operation is removed.

Supported operations: conversations.list/get, messages.list/get/send and reactions.update for chat messages. Channel reactions remain unsupported and return a rejection.
Conversation IDs retain the existing account-bound chat/channel IDs. Direct/group
classification is verified from membership when not present in discovery. Channel
replies require an exact thread root; new channel roots are not advertised.
Chat replies quote an exact existing message. Messages contain plain-text projections;
restricted, deleted and unavailable content is returned without its body.

New connections and reconnects default to read-only; select read/write explicitly
to advertise common sends. Explicit read-only connections also block legacy provider
writes. Existing connections without an access setting retain legacy API behavior,
but cannot use common writes until reconnected with read/write access.

History is partial, not an archive. Discovery and message cursors are signed and
bound to the account, connection and query. Provider continuation state is bounded
and expires after thirty minutes or plugin reload; invalid continuations require a
fresh snapshot. Thread filters are explicit; channel message reads need threadId.

The existing durable send ledger handles common idempotency keys in a separate
namespace. Accepted and uncertain attempts are never automatically resent. Reusing
a key for changed input is rejected. Receipts and change events contain metadata,
not message bodies. Readback confirmation does not prove recipient delivery.

Change collection performs a bounded authenticated discovery request and emits a
broad metadata invalidation. This adapter does not provide a native push watcher.
The core controls activity expiry, fallback scheduling and terminal auth pauses.
Expired/revoked credentials require reconnecting the same connection; repeating a
refresh does not bypass the pause. A resource-specific forbidden response does not
pause the entire account. Provider retry-after values are preserved.

Validation, from the repository root:

```sh
npm ci
npm run check
npm test
SWITCHBOARD_ROOT=/path/to/compatible/core npm run integration:messaging
```

`lib/messaging-v1.json` must be byte-identical to the core's
`server/contracts/messaging-v1.json`; the integration script asserts this.
The shared Microsoft Authentication source path is covered in
`test/teams-source.test.js` against this repository's
`plugins/microsoft-authentication`, not a separately pinned checkout.

Tests use offline provider fixtures and real local HTTP/SSE paths. They cover
read-only scope, schema parity, restricted content, replay and uncertain sends,
expiry without repeated provider requests, credential recovery and resource 403
isolation. They do not establish live tenant compatibility or send real messages.

The common Messaging contract supports the current Microsoft Authentication source
method as well as existing browser/token connections. Shared-source credentials
retain their Microsoft session mode, profile and identity. Discovery and requests
use the same access setting: read-only connections never advertise or permit sends.

Conversation discovery shares each page between regional chats and joined channels,
never exceeding the requested limit. Continuations bind the account, connection and
page size; a changed channel membership snapshot requires restarting discovery.
Provider chat cursors retain a stable native page size across mixed pages.

The account-private Notes to self conversation is included when Teams returns its
exact `48:notes` / `streamofnotes` identity. Reads and send preparation revalidate
it through the authenticated `ME` endpoint; callers cannot supply another user's
notes stream. Other system streams are excluded. No conversation is created as a
side effect of discovery.

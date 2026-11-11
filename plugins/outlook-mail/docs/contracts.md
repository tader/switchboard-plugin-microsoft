# Shared Email v1

Outlook Mail implements all fifteen optional common Email operations beside its
existing API. The adapter uses Microsoft Graph for Graph connections and the
existing private Outlook web compatibility backend for Outlook-audience bindings.
Graph credentials and IDs are never sent to the Outlook audience or vice versa.
The public Outlook REST API is retired; fixture conformance is not a claim of
ongoing tenant compatibility for this private backend.

Effective Mail.Read/ReadWrite/Send grants and configured read/send/full access
jointly determine discovery. Basic-mail grants that cannot read bodies do not
advertise full common reads. Read-only bindings never advertise writes. Optional
conditional `ifVersion` is explicitly rejected: these adapters do not advertise
an unverified atomic conditional-write guarantee. Observed versions remain useful
for comparing snapshots, not for pretending a preflight read locks the resource.

Mailboxes are folders. Signed, persistent pagination walks child folders as well
as top-level pages. Graph includes hidden folders; the compatibility backend
explicitly reports visible-only coverage. One inaccessible child does not hide
siblings. Directory traversal is bounded at 256 queued/1,024 visited parents and a
40 KiB continuation state; large/incomplete trees report partial coverage. Page
cursors bind connection, account, backend, filters and page size. Provider next
links must retain the exact audience, API prefix and requested collection path.

Folder `role` comes only from well-known folder names, never from localized display
names: `inbox`, `sentitems` (sent), `drafts`, `deleteditems` (trash), `archive` and
`junkemail` (spam). Graph v1.0 folders carry no `wellKnownName`, so the folder IDs
are resolved with one JSON batch (`POST /$batch` of `/me/mailFolders/{name}?$select=id`);
the compatibility backend reads them with at most three requests in flight. Roles are
best effort and never slow down or fail a listing: resolution runs alongside the
directory read, a listing waits at most 5 s for it and then continues with the roles
found so far (others `null`), and lookup errors are only logged. The role map is
cached in memory per connection for 6 hours (failed or incomplete lookups for 2
minutes), so later pages and listings make no lookups; an expired map is served
while it refreshes. A `wellKnownName` returned by the provider is honoured as well.
A missing folder (for example no archive folder yet) has no role. Other
folders report `role: null`. Moving to the archive-role folder is an ordinary move.

Message projections preserve text or HTML bodies, recipients, folder membership,
read/flag/draft state and thread identity. Unknown state stays null. Missing bodies
are unavailable; body previews are not substituted as complete content. Consumers
must sanitize HTML. Graph and Outlook expose a single selected body representation;
HTML is used if both text and HTML are supplied. Reads never mark messages read.

Organization changes patch only requested state. Graph supports read and flag;
flag writes on the compatibility backend are explicitly unsupported. Move targets
one folder and returns the resulting provider ID. Trash moves to deleteditems,
never permanently deletes. Draft delete/update/send first require an actual draft;
patches omit unspecified fields and attachments, preserving them on the provider.
There is one provider write per common operation after any preflight reads.

Send, reply, reply-all and forward use native actions, preserving provider reply
threading and forwarded attachments without download/re-upload or intermediate
server-created drafts. Reply defaults to sender-only. Original Bcc is not supplied
as new recipients. Graph actions accept a typed body and forward recipient fields;
compatibility reply/forward accepts text Comment only, and compatibility forward
rejects extra Cc/Bcc rather than dropping them. These unsupported options fail
before any provider write. Per-binding input constraints in discovery expose these
limits; the core validates them before persisting a write intent. Native action acceptance is not proof of delivery and
may not provide a sent message ID.

The core requires durable idempotency keys for writes. Accepted and uncertain
attempts never automatically resend. Metadata receipts/events/audit contain no
bodies. Reconciliation makes one bounded authenticated folder request and emits
scope invalidation. It is not native Microsoft push. Core activity leases and
terminal auth pause govern polling; resource 403 is not global auth failure.

Run `npm run check`, `npm test` and, with a compatible core,
`SWITCHBOARD_ROOT=/path/to/core npm run integration:email`. Integration exercises
both audiences through real local HTTP/SSE and offline provider fixtures. It does
not send real email, expand scopes or register cloud subscriptions.

## References

- [Message patches preserve omitted properties](https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0)
- [Reply action](https://learn.microsoft.com/en-us/graph/api/message-reply?view=graph-rest-1.0)
- [Forward action](https://learn.microsoft.com/en-us/graph/api/message-forward?view=graph-rest-1.0)
- [JSON batching](https://learn.microsoft.com/en-us/graph/json-batching)
- [Well-known folder names](https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0)
- [Child folder traversal and hidden folders](https://learn.microsoft.com/en-us/graph/api/mailfolder-list-childfolders?view=graph-rest-1.0)

The private compatibility routes and PascalCase wire format follow the existing
`microsoft-authentication/web-spec.ts`, separately tested alongside Graph.

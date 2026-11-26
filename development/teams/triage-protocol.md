# Triage protocol verification

The triage helpers use the existing regional Teams API transport. Private
contracts are verified independently of fixture success. `getTeamsCapabilities`
reports the production gates; callers also receive `canReply` and separate send
and read outcomes.

Live read check, 2026-10-08: after Microsoft Authenticator approval, ten inbox
groups and three batched reply-context reads succeeded with the browser closed.
Coverage issues remained explicit. This validates bounded chat reads, not channel
completeness or writes.

| Contract | Evidence | Production status |
| --- | --- | --- |
| Chat quote | [Inline quote builder](https://github.com/Maxim-Mazurok/teams-api/blob/e109d58122160c59cc210ec3284cd5a3d0f5ac8c/src/html-utils.ts), [chat send](https://github.com/Maxim-Mazurok/teams-api/blob/e109d58122160c59cc210ec3284cd5a3d0f5ac8c/src/api/chat-service.ts) | Source-verified and fixture-tested, not live-validated |
| Channel thread reads | [Thread conversation IDs and reads](https://github.com/Maxim-Mazurok/teams-api/blob/e109d58122160c59cc210ec3284cd5a3d0f5ac8c/src/teams-client.ts) | Source-verified and fixture-tested, bounded reads |
| Channel unread metadata | [CSA channel and horizon models](https://github.com/fossteams/teams-api/blob/dbbdc3681f32/pkg/models/teams.go) | Require explicit unread flag and valid horizon; unknown data is a coverage issue |
| Channel reply | Pinned primary message action appends the root to its regional POST; existing live reply reads verified rootMessageId association | **Enabled**, source-verified and fixture-tested; designated live send pending |
| Explicit chat read/unread | Native horizon/bookmark requests and browser-closed regional API readback, 2026-10-08 | **Enabled** for selected-message chat boundaries |
| Channel read/unread and automatic read-after-reply | Scope/bookmark interaction not live-verified | **Disabled**; successful replies report `read.status=deferred`, `reason=protocol_unverified` |

Chat quotes independently construct escaped microdata identifying the exact
message and sender, include at most 200 characters of preview, and use the
existing rich-text send contract. No upstream implementation is bundled.

Channel replies now share the exact-root direct reply implementation. Automatic
read-after-reply remains behind a disabled immutable production gate in `lib/triage.js`. Explicit chat read state has
separate enabled gates. Tests inject enabled contracts only
to exercise queueing, durable read stages and conservative advancement. Those
tests establish local behavior, not server compatibility. Synthetic redacted
examples in `test/fixtures/triage.json` are labeled accordingly.

Before enabling automatic read or channel read-state gates, capture native marker
requests in an explicitly designated test chat/channel. Record the endpoint, non-secret body,
read-marker scope and successful response; replace identifiers, people and text
with synthetic values before adding fixtures. Verify the actual channel reply
appears beneath its intended root. Verify the marker advances only through the
observed cutoff, preserves later arrivals/newer markers, and cannot clear an
unhandled sibling thread. If scope cannot be established, leave the gate disabled.
Do not send to an arbitrary unread conversation to obtain evidence.

`npm run check:triage -- --silent` authenticates through the saved export profile,
closes the browser and exercises bounded inbox/context reads. Without `--silent`
it permits interactive sign-in. It prints counts and sanitized errors, never
tokens or message contents, and performs no sends or read-marker writes.

## Selected-message read state

`setTeamsReadState` has separate immutable gates: chat read/unread enabled, channel read/unread disabled. Automatic read-after-reply remains disabled. Fixtures verify validation, exact boundaries, conservative coverage, account membership, replay, uncertainty, cancellation and durable unread generations. An unread generation blocks older automatic read grants even after a later explicit operation clears the uncertainty barrier.

Native capture on 2026-10-08 first used the user-designated `48:notes` self-chat. Message-level **Mark as unread** issued a successful PUT to `/api/chatsvc/emea/v1/users/ME/conversations/<thread>/properties?name=consumptionHorizonBookmark`, with `consumptionHorizonBookmark` containing selected message ID minus one, request time and its numeric client-message ID. Sidebar **Mark as read** used the same PUT with `0;<request time>;0`. Normal reading also captured a successful PUT with `name=consumptionhorizon` and `consumptionhorizon` containing message ID, request time and client-message ID. Notes is absent from standard CSA listing and remains unsupported by the membership-bound production resolver.

A subsequently user-designated regular meeting chat confirmed the same unread bookmark request and bookmark clearing when opened. Browser-closed regional GETs and CSA reads showed the ordinary `consumptionHorizon` at the later message, the explicit `userConsumptionHorizon` at selected ID minus one, and `isRead=false`. Numeric identifiers in the [redacted capture](../test/fixtures/read-state-native.json) are synthetic; no credentials or message contents are retained.

The implemented transport then passed browser-closed live checks in that regular chat:

- Explicit unread set the bookmark immediately before the selected message.
- Explicit read advanced the bookmark through the selected message while the later message remained unread and the ordinary horizon stayed at that later message.
- Replaying the original unread key returned its original result without overwriting the newer read boundary.
- Reading through the latest normal boundary cleared the bookmark and restored `isRead=true`. A positive bookmark at the latest message still forces the chat unread, so clearing is necessary in this case.
- The captured normal-horizon PUT was accepted at the already-read latest boundary without rewinding it.

No message was sent. The regular test chat was restored to its original read state. These checks used existing messages, not a newly injected arrival, and establish this account's private API compatibility rather than all tenant configurations. New arrivals and newer marker preservation have fixture coverage; no atomic compare-and-set is provided by the private endpoint.

The effective boundary is the positive explicit bookmark (`userConsumptionHorizon`), otherwise the ordinary horizon. A zero bookmark is cleared and falls back to the ordinary horizon. Reading with a later ordinary horizon advances the bookmark only to the selected cutoff. If the ordinary horizon is no newer than that cutoff, the transport advances it when needed, rechecks both native properties, then clears the bookmark. Explicit unread changes only the bookmark; it never rewinds the ordinary horizon. Regional properties are checked against CSA during preparation and again before writing. A lost response or mismatching readback is quarantined as uncertain; no write is automatically resent.

No channel was designated or mutated. Enable channel gates only after proving marker scope preserves unrelated unread threads; otherwise continue returning explicit deferral. Automatic read-after-reply needs separate verification of bookmark handling and keeps its disabled gate.

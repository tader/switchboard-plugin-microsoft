---
title: Shared Calendar contract (staged)
services: [outlook-calendar]
---

# Shared Calendar contract

Adds the optional Calendar v1 contract beside the existing Outlook Calendar API.
It depends on the staged core API in [Switchboard #40](https://github.com/tader/switchboard/pull/40),
which is disabled by default. The canonical JSON schema and wrapper are vendored
from the core; Google Calendar implements the same schema. No existing connection,
authentication method, configured permissions or consent is changed.

The adapter handles both Microsoft Graph and the Outlook web compatibility backend
selected by Microsoft Authentication. Backend audience, URLs, property casing,
scopes, page links and timezone names remain internal to the plugin. Capability
discovery checks actual grants for the selected audience and the connection's read
binding; individual calendar edit rights are reported separately when available.
Mailbox timezone is unknown rather than requesting additional mailbox permissions.

Event lists use a bounded calendar view with expanded recurring occurrences.
All-day ends remain exclusive civil dates, including across daylight-saving changes.
Timed events preserve the instant and report the original timezone using the CLDR
Windows/IANA mapping. Ambiguous wall-clock writes use UTC to preserve the requested
instant. Unsupported legacy custom timezones fail explicitly instead of inventing
an all-day date. Cursors bind to the connection, backend, calendar and time window;
next links cannot change the request host, resource or window.

Update/delete require the previously read version and use `If-Match`. Updates also
reject a stale version before dispatch. Online-meeting description replacement is
rejected because replacing the provider meeting blob could break the join link.
The common schema excludes attendees, recurrence-rule editing and reminders.
Creates have no attendees. Editing/deleting existing meetings can trigger normal
provider notifications; this adapter does not claim to suppress them.

Reconciliation currently performs one calendar-list read and emits a broad
connection-wide calendar invalidation. It is not native Microsoft push or a delta
feed. Core activity leases, authentication pause, cooldown and SSE replay control
this fallback. Native subscriptions and full cross-provider conformance remain
follow-up work. Write and mapping tests use fixtures; live backend verification is
required before enabling the staged API in production.

References: [Graph calendar view](https://learn.microsoft.com/graph/api/calendar-list-calendarview?view=graph-rest-1.0),
[event update](https://learn.microsoft.com/graph/api/event-update?view=graph-rest-1.0),
[event resource](https://learn.microsoft.com/graph/api/resources/event?view=graph-rest-1.0),
and [Outlook REST reference](https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/calendar-rest-operations).

`windows-zones.json` is derived from Unicode CLDR release 48
[`windowsZones.xml`](https://github.com/unicode-org/cldr/blob/release-48/common/supplemental/windowsZones.xml):
territory `001` gives Windows-to-IANA defaults; every territory supplies IANA aliases
for conversion back to Windows names. Copyright Unicode, Inc.; see the included
`UNICODE-LICENSE.txt`. This data is bundled; runtime calls do not download mappings.

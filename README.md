# Microsoft

Microsoft APIs for Switchboard: browser-session authentication, Microsoft Graph, Outlook Mail and Calendar, OneDrive, To Do, Excel, Teams, Word, PowerPoint and Org Explorer. Use an existing Microsoft web session locally or through a trusted peer, or keep using an app registration with OAuth.

Maintained independently from [Switchboard](https://github.com/tader/switchboard).

## Installation

Open **Plugins → Community** to install individual plugins. Alternatively install `tader/switchboard-plugin-microsoft` without a path to install the whole family, or specify one of these paths:

| Plugin | Path |
| --- | --- |
| Microsoft Authentication | `plugins/microsoft-authentication` |
| Microsoft | `plugins/microsoft` |
| Outlook Mail | `plugins/outlook-mail` |
| Outlook Calendar | `plugins/outlook-calendar` |
| OneDrive | `plugins/onedrive` |
| SharePoint | `plugins/sharepoint` |
| Microsoft To Do | `plugins/microsoft-todo` |
| Excel | `plugins/excel` |
| Teams | `plugins/teams` |
| Word | `plugins/word` |
| PowerPoint | `plugins/powerpoint` |
| Org Explorer | `plugins/org-explorer` |

Each plugin retains its own ID and version. Shared helpers are resolved through plugin dependencies.

## Existing installations

Teams retains the `teams-web` plugin/service ID and all current adapter routes, convenience methods, saved-call paths, browser sign-in and token imports. Install `plugins/teams` from this repository to replace the separate Teams plugin. Existing encrypted credentials, profile data and send ledgers stay under the same plugin ID. To consolidate sign-in, reconnect a Teams connection with **Use Microsoft Authentication** and select the same source used by the other Microsoft services. See the [Teams guide](plugins/teams/README.md).

Install this repository before upgrading a Switchboard instance that bundles `microsoft`. The plugin ID, service IDs, authentication methods, settings and persistent data paths are preserved. Installed plugins replace the built-in copy. Existing connections and saved calls continue to work; keep the instance data directory and encryption key.

## Setup guides

- [Web authentication without a new app registration](plugins/microsoft-authentication/docs/web-authentication.md)
- [Word, PowerPoint and Org Explorer APIs](plugins/microsoft/content/README.md)
- [app-registration](plugins/microsoft/docs/app-registration.md)

## OpenAPI descriptions

The slim Graph descriptions ship with the plugin. Rebuild the generated descriptions with `npm run build:openapi -- /path/to/openapi.yaml`; the focused SharePoint description is maintained separately in `plugins/microsoft/openapi/sharepoint.json`. The build helper is self-contained and uses YAML only during development.

## Development

Use Node.js 24 or newer. Run `npm ci`, `npm run check` and `npm test`. Run `npm run check:access -- --silent` for live read-only checks with the persistent credential cache; see the web authentication guide for initial sign-in and optional write validation. The plugin folder includes local type declarations, so development does not require a Switchboard checkout. Type imports are erased at runtime; shared helpers are provided by `ctx.require()` and the manifest declares compatible versions.

Keep `plugin.json` and `package.json` versions in sync. Commit version changes before tagging a release; `gh release create` does not bump versions. Switchboard follows the branch, tag or commit selected during installation.

## Releases

Release-please opens version and changelog pull requests from Conventional Commits. Merge the release PR to publish its tag and GitHub release. Plugin manifests are updated with their package versions. Family repositories maintain an independent version for each plugin.

## SharePoint

SharePoint uses a work or school account. Select read access (`Sites.Read.All`) or read and edit access (`Sites.ReadWrite.All`); the app registration or existing Microsoft Authentication session must already have the selected permission, and site permissions still apply. Graph returns pages of results with `@odata.nextLink`; follow it using the same connection. Read modern page contents using `$expand=canvasLayout`, and list item fields using `$expand=fields`. Graph supports only its documented modern page web parts; new pages are drafts and are not published automatically. Classic pages are not supported by the page API. Use the SharePoint File Download connection to download files up to 25 MiB; it follows Microsoft transfer redirects without sending the Graph token to the transfer host. Word and PowerPoint plugins can read structured Office documents from SharePoint drives when configured for all accessible files.

## Activity response previews

Switchboard 0.17.0 can retain a small reviewed projection of an explicitly called HTTP response. These declarations follow the existing Spotify example; core owns capture, encryption, owner access and retention. Calls use their existing scopes and authorization. No extra request is made.

| Service | Exact query-free GET | Retained string | Authentication |
| --- | --- | --- | --- |
| Microsoft Graph, OneDrive | `https://graph.microsoft.com/v1.0/me/drive` | `driveType` | `oauth`, `device`, `authentication-source` on Graph |
| Outlook Calendar | `https://graph.microsoft.com/v1.0/me/calendar` | `defaultOnlineMeetingProvider` | `oauth`, `device`, `authentication-source` on Graph |
| Word, PowerPoint | Setup's allocated loopback origin + `/capabilities` | `kind` | `authentication-source` |
| Teams | Setup's allocated loopback origin + `/capabilities` | `transport` | `authentication-source` |

Each field is limited to 32 UTF-8 bytes. Graph fixtures follow the documented [drive](https://learn.microsoft.com/en-us/graph/api/resources/drive?view=graph-rest-1.0) and [calendar](https://learn.microsoft.com/en-us/graph/api/resources/calendar?view=graph-rest-1.0) shapes. The capability fixtures exercise the actual adapter response. Word and PowerPoint declarations live in Microsoft's shared content helper; install its updated release to receive them. No new helper option or dependency minimum is introduced.

Only an audited local final HTTP 200 with eligible JSON and clean body completion qualifies. Core drops all unselected fields and omits previews for invalid or missing selected strings, credential echoes, responses over 16 KiB, errors, empty/binary/streaming-event responses, interrupted/cancelled bodies, query strings and mismatched routes or auth methods. Finite JSON streams can qualify at completion. Capture requires provisioned isolated preview storage and nonzero output retention; older cores continue serving responses. Removing declarations stops future capture; existing previews expire under the instance policy. Zero retention purges previews across plugins.

The Outlook web backend, Excel workbook data, Outlook Mail, To Do, Org Explorer, SharePoint and `sharepoint-download` remain metadata-only. Microsoft Authentication supplies credentials rather than audited data responses. Teams browser/token profiles, document contents, mail, meetings, chats, sign-in, plugin maintenance, native contracts and subscriptions have no preview declaration.

Verification with synthetic data (no live Microsoft sign-in):

```sh
npm ci
npm run check
npm test
# In a separate Switchboard 0.17.0 checkout, run npm ci first:
SWITCHBOARD_ROOT=/path/to/switchboard npm run integration:output
# Core's retention, owner/scoped-token access and storage regression suites:
cd /path/to/switchboard
node --test test/activity-output.test.ts test/activity-preview-store.test.ts
```

The preview integration provisions disposable storage and drives real core `execute` calls through the returned service declarations: nine Graph auth profiles, private fields and credential echoes, exact matching, byte limits, HTTP/media failures, cancellation, interruption, retention disablement and storage unavailability. It also verifies the Outlook backend omission and six actual loopback adapters across two ports per service. Each call checks the upstream count and retained projection; synthetic Graph responses check identical caller bytes/status/headers. OAuth bearer hooks and loopback credential/session acquisition are fixtures; authentication-source Graph request authorization is real. OAuth/device 401 retries are exercised here; source rotation/retry is covered separately by `integration:authentication` and existing authentication tests. Core's own suites cover reveal permissions, encryption, retention changes and restore fencing. Live provider accounts and released-package installation are release verification steps.

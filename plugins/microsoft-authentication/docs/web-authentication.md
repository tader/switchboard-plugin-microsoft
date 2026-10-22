---
title: Microsoft web authentication
services: [microsoft-authentication, microsoft-graph, outlook-mail, outlook-calendar, onedrive, microsoft-todo, excel, teams-web, word, powerpoint, org-explorer]
order: 10
---

# Microsoft web authentication

Use your existing Microsoft web access without registering an application. This method requires Switchboard authentication-source support and an account that can sign in to Teams web. Microsoft still applies your account permissions and company sign-in policies.

## Connect on your laptop

1. Install **Microsoft Authentication** (`plugins/microsoft-authentication`) on the computer used for sign-in, with Edge or Chrome available. Install the Microsoft product plugins you want. Update the shared **Microsoft** helper before adding Word, PowerPoint or Org Explorer to an existing installation.
2. Add a **Microsoft Authentication** connection and choose **Sign in through local browser**. Complete sign-in in the dedicated Edge or Chrome window. The window closes when authentication finishes. Switchboard shows browser sign-in instructions; no device code is needed. Upgrade Switchboard if browser sign-in is reported as unavailable.
3. Add Outlook Mail, Calendar, OneDrive, To Do, Excel, Teams, Word, PowerPoint, Org Explorer or Microsoft Graph. Choose **Use Microsoft Authentication**, select your authentication connection and choose its access level.

The browser is used only for authentication. Credentials are encrypted by Switchboard. Data operations call Microsoft APIs directly. Graph and Outlook product connections share a cached credential bundle and serialize renewal. Teams caches its derived API and chat tokens from the same source.

## Use a server through a peer

Pair the laptop and server through Switchboard Peers. On the laptop, open the authentication connection's **Sharing** dialog. Share the connection with the server, then separately enable **Allow authentication** for that server. Normal connection sharing alone does not grant credential access.

On the server, create a local Microsoft product connection using the shared authentication source. The server receives encrypted-at-rest access and refresh credentials; browser profiles and cookies stay on the laptop. Data requests go directly from the server to Microsoft, including while the laptop is offline.

When HTTP renewal stops working, the server requests fresh authentication from the laptop. The laptop first tries its saved browser profile headlessly. If Microsoft requires interaction, reconnect the authentication connection on the laptop and complete sign-in there. The server retries source acquisition on its next request.

Unsharing blocks new transfers and disables linked use once that change reaches the server. Exported Microsoft tokens cannot be recalled from an offline machine. Microsoft session revocation controls the provider's tokens themselves.

## Available APIs

The plugin uses Graph for file and Excel operations. If the existing session does not grant Graph mail, calendar or task scopes, those services use Outlook's own API audience. The API reference and connection base URL change accordingly; native Outlook bodies use PascalCase, such as `Subject` and `IsRead`. Graph bodies and IDs are not interchangeable with native Outlook ones.

Microsoft retired the public Outlook REST API. This plugin's private compatibility backend was live-tested using both Outlook and Teams web credentials and may change. The task backend supports Outlook task folders and tasks; modern To Do features such as My Day and checklists are not exposed by that backend.

Excel uses workbook APIs for modern workbooks in OneDrive for Business and SharePoint. Select **all files you can access** for another drive or SharePoint. Existing file permissions still apply. Use persistent workbook sessions for related edits; changes made without a session also persist. The reference includes `workbook-session-id` headers, range values/formulas, formatting, tables and calculation.

The generic Graph connection accepts permissions already present in the web session. It never opens a new consent flow. Requests are limited to supported resource/permission mappings; use the existing OAuth methods for other Graph APIs. Missing permissions are reported rather than requesting admin consent.

## Development checks and cache

Run `npm run check:teams -- --silent` to validate the consolidated Teams adapter, including triage, without changing messages or read state. Run `npm run check:access -- --silent` to reuse the diagnostic session without interactive sign-in. Without `--silent`, the check opens an interactive browser only after cached credentials, HTTP renewal and silent authentication cannot finish. `--capture` explicitly requests fresh browser authentication.

The shared diagnostic cache is `.local-profile/microsoft/`, excluded from Git. Its credential file is owner-only (`0600`), its directory/profile is owner-only (`0700`), and refresh-token rotations are saved atomically even if a later check fails. `MICROSOFT_AUTH_CACHE` chooses another persistent location; `MICROSOFT_BROWSER_PATH` selects a browser executable. The cache contains bearer credentials, not returned mail or document content. Keep it between checks to avoid repeated sign-in.

Set `MICROSOFT_TEST_SITE`, `MICROSOFT_TEST_DRIVE` and `MICROSOFT_TEST_WORKBOOK` to selected IDs for additional read checks. `npm run check:writes -- --write --silent` explicitly creates, verifies and removes disposable drafts, an event without attendees, a task and workbooks. It never sends mail. Set `MICROSOFT_TEST_SHAREPOINT_DRIVE` to a writable SharePoint document-library drive for its workbook test, and `MICROSOFT_TEST_SHAREPOINT_FOLDER` to an optional folder path within that drive.

Web-app refresh tokens can reach Microsoft's 24-hour SPA lifetime limit; token rotation does not extend that limit. A fresh browser authentication can reuse the saved Microsoft session, but MFA, revocation and company policies may require you to sign in again. Capture time is not treated as proof of a refresh token's remaining lifetime.


## One sign-in including Teams

Install Teams from `tader/switchboard-plugin-microsoft`, folder `plugins/teams`, instead of the standalone Teams repository. The plugin ID remains `teams-web`, preserving existing connection and plugin-data identities during migration.

Microsoft Authentication now signs into Teams web by default. Live validation confirmed that this same Teams refresh credential can issue Graph, Outlook, Skype/Teams, chat aggregation and Substrate tokens. Select this source for every Microsoft service, including the consolidated `plugins/teams` plugin. Each service caches its required resource tokens and enforces its own request boundaries.

Previously cached Outlook sign-ins continue to work for Outlook and Graph, but Teams rejects their chat authorization with `410 ApiRestricted`. Reconnect the source once using Teams sign-in to enable the whole family. The source never changes a refresh token's issuing client ID or origin. Graph permissions differ between the Outlook and Teams web clients; only already-granted permissions are used. Some Outlook operations still require the native Outlook backend (for example Mail.Send), with the compatibility limits described above.

Word and PowerPoint use Graph file transfer and server-side Open XML editing with version checks, rather than a live Office session. Org Explorer uses the existing Graph directory grants for read-only people and hierarchy queries. These three services use the same authentication source; they do not request new permissions.

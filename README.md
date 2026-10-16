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

The slim Graph descriptions ship with the plugin. Rebuild them with `npm run build:openapi -- /path/to/openapi.yaml`. The build helper is self-contained and uses YAML only during development.

## Development

Use Node.js 24 or newer. Run `npm ci`, `npm run check` and `npm test`. Run `npm run check:access -- --silent` for live read-only checks with the persistent credential cache; see the web authentication guide for initial sign-in and optional write validation. The plugin folder includes local type declarations, so development does not require a Switchboard checkout. Type imports are erased at runtime; shared helpers are provided by `ctx.require()` and the manifest declares compatible versions.

Keep `plugin.json` and `package.json` versions in sync. Commit version changes before tagging a release; `gh release create` does not bump versions. Switchboard follows the branch, tag or commit selected during installation.

## Releases

Release-please opens version and changelog pull requests from Conventional Commits. Merge the release PR to publish its tag and GitHub release. Plugin manifests are updated with their package versions. Family repositories maintain an independent version for each plugin.

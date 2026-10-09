# Microsoft

Microsoft sign-in for Microsoft Graph (Microsoft 365, Outlook.com, OneDrive). Other plugins build services such as Outlook Mail on top of it.

Maintained independently from [Switchboard](https://github.com/tader/switchboard).

## Installation

Open **Plugins → Community** to install individual plugins. Alternatively install `tader/switchboard-plugin-microsoft` without a path to install the whole family, or specify one of these paths:

| Plugin | Path |
| --- | --- |
| Microsoft | `plugins/microsoft` |
| Outlook Mail | `plugins/outlook-mail` |
| Outlook Calendar | `plugins/outlook-calendar` |
| OneDrive | `plugins/onedrive` |
| Microsoft To Do | `plugins/microsoft-todo` |

Each plugin retains its own ID and version. Shared helpers are resolved through plugin dependencies.

## Existing installations

Install this repository before upgrading a Switchboard instance that bundles `microsoft`. The plugin ID, service IDs, authentication methods, settings and persistent data paths are preserved. Installed plugins replace the built-in copy. Existing connections and saved calls continue to work; keep the instance data directory and encryption key.

## Setup guides

- [app-registration](plugins/microsoft/docs/app-registration.md)

## OpenAPI descriptions

The slim Graph descriptions ship with the plugin. Rebuild them with `npm run build:openapi -- /path/to/openapi.yaml`. The build helper is self-contained and uses YAML only during development.

## Development

Use Node.js 24 or newer. Run `npm ci` and `npm run check`. The plugin folder includes local type declarations, so development does not require a Switchboard checkout. Type imports are erased at runtime; shared helpers are provided by `ctx.require()` and the manifest declares compatible versions.

Keep `plugin.json` and `package.json` versions in sync. Commit version changes before tagging a release; `gh release create` does not bump versions. Switchboard follows the branch, tag or commit selected during installation.

## Releases

Release-please opens version and changelog pull requests from Conventional Commits. Merge the release PR to publish its tag and GitHub release. Plugin manifests are updated with their package versions. Family repositories maintain an independent version for each plugin.

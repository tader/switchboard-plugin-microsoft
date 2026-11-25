# Immediate shared-profile sign-in

Connect and Reconnect now request a visible browser directly for both shared and legacy profiles. Shared-profile acquisition previously attempted a silent capture with a 60-second timeout before opening visibly. Background renewal remains silent; profile serialization, cancellation, cookie isolation, and account matching are unchanged.

The dialog and user documentation now describe the visible sign-in flow. A profile already in use can still queue sign-in until it becomes available.

## Verification

- The new shared-profile regression failed before the change: capture requested `silent: true` instead of `false`.
- Regression coverage checks immediate Connect and Reconnect, no second browser after 60 seconds, cancellation, silent background recovery, and account mismatch rejection for shared and legacy profiles.
- `npm run check`: passed.
- `npm test`: 313 passed, no failures.
- `npm run build:content` and vendor diff check: passed, no generated changes.
- Live account sign-in was not exercised. The pull request workflow runs the pinned-core email and messaging contract integrations.

No core or macOS app change is required. Versioning remains with Release Please.

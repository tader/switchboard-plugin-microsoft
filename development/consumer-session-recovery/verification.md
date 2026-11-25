# Microsoft consumer session recovery

Forced retries now check the selected source for its current session, then renew the requested API audience. Only explicit peer-offline failures allow cached HTTP renewal. Other source failures and account validation errors propagate.

Source adoption retains sibling resource scopes with expired cached tokens, requiring renewal before use. Native resource rotation persists the complete validated token bundle; failed native acquisition no longer replaces the durable grant with an incomplete source bundle. Graph rotations still persist before identity checks.

Verification: typecheck and all 323 tests passed, including Graph/Outlook recovery, transient failures, sibling resource preservation, offline fallback, and adapter option forwarding. Content generation produced no vendor changes. The new offline integration exercises real core shared credentials and proxy retries for Graph and Outlook, including recovery after a failed exchange. Existing email (Graph and native Outlook) and Teams HTTP/SSE integrations passed against the locally available core. CI also runs these against its pinned core revision. No live account sign-in was exercised.

Changes are confined to the Microsoft plugin repository. Versioning remains with Release Please.

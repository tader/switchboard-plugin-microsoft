---
title: Word
services: [word]
order: 20
---

# Word

Install Microsoft Authentication and connect it using the shared Teams web sign-in. Add a Word connection, choose **Use Microsoft Authentication**, and select the same source used by the other Microsoft services. You can also select a source explicitly shared by a direct peer. Existing credentials renew over HTTP while the authentication device is offline.

Choose read access, edit access to your own drive, or edit access to all files you can access (including SharePoint). Read the file first to obtain its current version and editable element IDs. Submit an edit batch with that version; stale versions and locked files fail without overwriting. Creation supports simple files and existing templates. Data operations use Graph and server-side Open XML processing, with no live Office or browser session.

See the API reference for request schemas and examples. Full limits and operation details are in the Microsoft helper documentation: [Office content adapter](https://github.com/tader/switchboard-plugin-microsoft/blob/main/plugins/microsoft/content/README.md).

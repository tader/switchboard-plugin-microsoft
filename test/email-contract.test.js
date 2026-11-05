import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { Ajv } from "ajv";
import { createContracts, project } from "../plugins/outlook-mail/contracts.ts";
const ajv = new Ajv({ strict: false });
const json = (body, status = 200, headers = {}) =>
  new Response([202, 204].includes(status) ? null : JSON.stringify(body), {
    status,
    headers,
  });
async function fixture(t, respond, native = false) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "outlook-email-"));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const requests = [],
    contracts = createContracts(root);
  const ctx = {
    connection: {
      id: "c",
      account: { id: "a" },
      methodId: native ? "authentication-source" : "oauth",
      config: { access: "full", backend: native ? "outlook" : "graph" },
      credentials: native
        ? {
            scope: "User.Read",
            resources: {
              outlook: {
                scope:
                  "https://outlook.office.com/Mail.ReadWrite https://outlook.office.com/Mail.Send",
              },
            },
          }
        : { scope: "Mail.ReadWrite Mail.Send" },
    },
    signal: new AbortController().signal,
    idempotencyKey: "key",
    async request(r) {
      requests.push(r);
      return respond(r);
    },
  };
  async function call(name, args = {}) {
    const op = contracts.operations["email." + name];
    assert.ok(ajv.validate(op.inputSchema, args), ajv.errorsText());
    const result = await op.invoke(args, ctx);
    assert.ok(ajv.validate(op.outputSchema, result), ajv.errorsText());
    return result;
  }
  return { root, ctx, requests, contracts, call };
}
const mail = (id = "m", draft = false) => ({
  id,
  "@odata.etag": 'W/"version"',
  subject: "Subject",
  body: { contentType: "HTML", content: "<p>Hello</p>" },
  from: { emailAddress: { address: "person@example.test", name: "Person" } },
  toRecipients: [{ emailAddress: { address: "self@example.test" } }],
  ccRecipients: [],
  bccRecipients: [],
  parentFolderId: "inbox",
  isRead: false,
  isDraft: draft,
  flag: { flagStatus: "flagged" },
  conversationId: "thread",
  hasAttachments: true,
});
// Role resolution reads each well-known folder by name; directory assertions ignore those lookups.
const WELL_KNOWN =
  /\/mailFolders\/(inbox|sentitems|drafts|deleteditems|archive|junkemail)\?/i;
const directory = (requests) => requests.filter((r) => !WELL_KNOWN.test(r.url));
const pascal = (value) =>
  Array.isArray(value)
    ? value.map(pascal)
    : value && typeof value === "object"
      ? Object.fromEntries(
          Object.entries(value).map(([k, v]) => [
            k.startsWith("@") ? k : k[0].toUpperCase() + k.slice(1),
            pascal(v),
          ]),
        )
      : value;
test("Outlook effective grants use the correct resource audience and binding access", async (t) => {
  const f = await fixture(t, () => json({}));
  assert.equal(f.contracts.available(f.ctx.connection).length, 15);
  f.ctx.connection.config.access = "read";
  assert.equal(f.contracts.available(f.ctx.connection).length, 5);
  await assert.rejects(
    () =>
      f.contracts.operations["email.messages.send"].invoke(
        { subject: "", to: [], text: "" },
        f.ctx,
      ),
    { code: "permission_denied" },
  );
  assert.equal(f.requests.length, 0);
  f.ctx.connection.config.access = "send";
  f.ctx.connection.credentials.scope = "Mail.Read Mail.Send";
  assert.ok(
    f.contracts.available(f.ctx.connection).includes("email.drafts.send"),
  );
  assert.ok(
    !f.contracts.available(f.ctx.connection).includes("email.drafts.update"),
  );
  f.ctx.connection.methodId = "authentication-source";
  f.ctx.connection.config.backend = "outlook";
  assert.deepEqual(f.contracts.available(f.ctx.connection), []);
  f.ctx.connection.credentials.resources = {
    outlook: {
      scope:
        "https://outlook.office.com/Mail.Read https://graph.microsoft.com/Mail.Send",
    },
  };
  assert.equal(f.contracts.available(f.ctx.connection).length, 5);
});
test("Graph and Outlook body/recipient projections are equivalent without inventing read state", () => {
  assert.deepEqual(project(mail()), project(pascal(mail())));
  const m = project(mail());
  assert.equal(m.text, null);
  assert.equal(m.html, "<p>Hello</p>");
  assert.equal(m.contentState, "available");
  assert.equal(m.isRead, false);
  assert.equal(m.isFlagged, true);
  const missing = mail();
  delete missing.body;
  delete missing.isRead;
  delete missing.flag;
  const p = project(missing);
  assert.equal(p.contentState, "unavailable");
  assert.equal(p.text, null);
  assert.equal(p.isRead, null);
  assert.equal(p.isFlagged, null);
});
test("signed message continuations survive restart but reject tampering, account/filter and foreign endpoints", async (t) => {
  let foreign = false;
  const f = await fixture(t, () =>
    json({
      value: [mail()],
      "@odata.nextLink": foreign
        ? "https://evil.example/v1.0/me/messages?$skip=2"
        : "https://graph.microsoft.com/v1.0/me/messages?$skip=2",
    }),
  );
  const first = await f.call("messages.list", { limit: 1, isRead: false });
  assert.equal(first.coverage.complete, false);
  assert.equal(first.items[0].html, "<p>Hello</p>");
  const second = createContracts(f.root);
  await second.operations["email.messages.list"].invoke(
    { limit: 1, isRead: false, cursor: first.nextCursor },
    f.ctx,
  );
  assert.equal(new URL(f.requests[1].url).searchParams.get("$skip"), "2");
  await assert.rejects(
    () =>
      f.call("messages.list", {
        limit: 2,
        isRead: false,
        cursor: first.nextCursor,
      }),
    { code: "invalid_cursor" },
  );
  f.ctx.connection.account.id = "different";
  await assert.rejects(
    () =>
      f.call("messages.list", {
        limit: 1,
        isRead: false,
        cursor: first.nextCursor,
      }),
    { code: "invalid_cursor" },
  );
  f.ctx.connection.account.id = "a";
  const parts = first.nextCursor.split(".");
  parts[0] = Buffer.from(JSON.stringify(["forged", "/me/sendMail"])).toString(
    "base64url",
  );
  await assert.rejects(
    () =>
      f.call("messages.list", {
        limit: 1,
        isRead: false,
        cursor: parts.join("."),
      }),
    { code: "invalid_cursor" },
  );
  foreign = true;
  await assert.rejects(() => f.call("messages.list", {}), {
    code: "invalid_cursor",
  });
});
test("folder discovery walks child pages and detects provider cycles", async (t) => {
  const f = await fixture(t, (r) =>
    new URL(r.url).pathname.endsWith("/childFolders")
      ? json({
          value: [
            {
              id: "parent",
              displayName: "Parent repeated",
              childFolderCount: 1,
            },
            { id: "nested", displayName: "Nested", childFolderCount: 0 },
          ],
        })
      : json({
          value: [{ id: "parent", displayName: "Parent", childFolderCount: 1 }],
        }),
  );
  const first = await f.call("mailboxes.list", { limit: 2 });
  assert.ok(first.nextCursor);
  const second = await f.call("mailboxes.list", {
    limit: 2,
    cursor: first.nextCursor,
  });
  assert.equal(second.items[1].id, "nested");
  assert.equal(second.nextCursor, undefined);
  assert.equal(second.coverage.complete, true);
  assert.ok(directory(f.requests)[1].url.includes("/parent/childFolders"));
});
test("native Outlook pagination and OData filters use its audience and PascalCase", async (t) => {
  const f = await fixture(
    t,
    () =>
      json({
        value: [pascal(mail("d", true))],
        "@odata.nextLink":
          "https://outlook.office.com/api/v2.0/me/messages?$skip=1",
      }),
    true,
  );
  const result = await f.call("drafts.list", { limit: 1 });
  assert.equal(result.items[0].id, "d");
  assert.equal(result.items[0].message.draft, true);
  assert.equal(
    new URL(f.requests[0].url).searchParams.get("$filter"),
    "IsDraft eq true",
  );
  assert.ok(
    f.requests[0].url.startsWith("https://outlook.office.com/api/v2.0/"),
  );
  await f.call("drafts.list", { limit: 1, cursor: result.nextCursor });
  assert.equal(new URL(f.requests[1].url).searchParams.get("$skip"), "1");
});
for (const native of [false, true])
  test(`draft patch/delete/send preflight and one-write behavior (${native ? "Outlook" : "Graph"})`, async (t) => {
    let draft = true;
    const f = await fixture(
      t,
      (r) =>
        r.method === "GET"
          ? json(native ? pascal(mail("d", draft)) : mail("d", draft))
          : r.method === "PATCH"
            ? json(native ? { Id: "d" } : { id: "d" })
            : json({}, 202),
      native,
    );
    const result = await f.call("drafts.update", {
      id: "d",
      changes: { subject: "Changed", to: [] },
    });
    assert.equal(result.status, "accepted");
    const patch = JSON.parse(f.requests.find((r) => r.method === "PATCH").body);
    assert.deepEqual(
      patch,
      native
        ? { Subject: "Changed", ToRecipients: [] }
        : { subject: "Changed", toRecipients: [] },
    );
    assert.equal(patch.Attachments ?? patch.attachments, undefined);
    draft = false;
    const before = f.requests.filter((r) => r.method !== "GET").length;
    const rejected = await f.call("drafts.delete", { id: "d" });
    assert.equal(rejected.status, "rejected");
    assert.equal(rejected.code, "not_found");
    assert.equal(f.requests.filter((r) => r.method !== "GET").length, before);
    draft = true;
    assert.equal((await f.call("drafts.send", { id: "d" })).status, "accepted");
    assert.equal(f.requests.at(-1).url.endsWith("/d/send"), true);
  });
test("send/reply/forward use native actions without multi-write draft creation", async (t) => {
  const f = await fixture(t, () => json({}, 202));
  await f.call("messages.send", {
    subject: "Hello",
    to: [{ address: "person@example.test" }],
    text: "Hello",
  });
  assert.equal(f.requests[0].url.endsWith("/me/sendMail"), true);
  assert.equal(JSON.parse(f.requests[0].body).saveToSentItems, true);
  await f.call("messages.reply", { id: "m", text: "Reply", replyAll: true });
  assert.equal(f.requests.at(-1).url.endsWith("/replyAll"), true);
  assert.deepEqual(JSON.parse(f.requests.at(-1).body), {
    message: { body: { contentType: "Text", content: "Reply" } },
  });
  await f.call("messages.forward", {
    id: "m",
    to: [{ address: "next@example.test" }],
    text: "Forward",
  });
  const body = JSON.parse(f.requests.at(-1).body);
  assert.equal(
    body.message.toRecipients[0].emailAddress.address,
    "next@example.test",
  );
  assert.equal(body.message.attachments, undefined);
  assert.equal(f.requests.length, 3);
});
test("Outlook compatibility actions reject unsupported options before dispatch", async (t) => {
  const f = await fixture(t, () => json({}, 202), true);
  await f.call("messages.forward", {
    id: "m",
    to: [{ address: "person@example.test" }],
    text: "Intro",
  });
  assert.deepEqual(JSON.parse(f.requests[0].body), {
    Comment: "Intro",
    ToRecipients: [{ EmailAddress: { Address: "person@example.test" } }],
  });
  const n = f.requests.length;
  assert.equal(
    (
      await f.call("messages.forward", {
        id: "m",
        to: [{ address: "person@example.test" }],
        html: "<p>Intro</p>",
      })
    ).code,
    "unsupported_option",
  );
  assert.equal(
    (await f.call("messages.update", { id: "m", changes: { isFlagged: true } }))
      .code,
    "unsupported_option",
  );
  assert.equal(f.requests.length, n);
});
test("trash moves to deleteditems, narrow state edits preserve other fields and conditional writes are explicit", async (t) => {
  const f = await fixture(t, () => json({ id: "moved" }));
  await f.call("messages.trash", { id: "m" });
  assert.deepEqual(JSON.parse(f.requests[0].body), {
    destinationId: "deleteditems",
  });
  assert.equal(f.requests[0].method, "POST");
  await f.call("messages.update", {
    id: "m",
    changes: { isRead: true, isFlagged: false },
  });
  assert.deepEqual(JSON.parse(f.requests[1].body), {
    isRead: true,
    flag: { flagStatus: "notFlagged" },
  });
  assert.equal(
    ajv.validate(f.contracts.operations["email.messages.update"].inputSchema, {
      id: "m",
      ifVersion: "v",
      changes: { isRead: true },
    }),
    false,
  );
  await assert.rejects(
    () =>
      f.contracts.operations["email.messages.update"].invoke(
        { id: "m", ifVersion: "v", changes: { isRead: true } },
        f.ctx,
      ),
    { code: "unsupported_precondition" },
  );
});
test("terminal auth, scoped permission errors and rate backoff remain distinct", async (t) => {
  let status = 401;
  const f = await fixture(t, () =>
    json({ error: { message: "private provider text" } }, status, {
      "retry-after": "120",
    }),
  );
  await assert.rejects(() => f.call("messages.get", { id: "m" }), {
    code: "authentication_required",
  });
  status = 403;
  await assert.rejects(() => f.call("messages.get", { id: "m" }), {
    code: "permission_denied",
  });
  status = 429;
  await assert.rejects(
    () => f.call("messages.get", { id: "m" }),
    (e) =>
      e.code === "rate_limited" &&
      e.retryAfterMs === 120000 &&
      !String(e).includes("private provider text"),
  );
});

test("one inaccessible child folder does not hide siblings and reports partial coverage", async (t) => {
  const f = await fixture(t, (r) =>
    r.url.includes("/blocked/childFolders")
      ? json({}, 403)
      : r.url.includes("/open/childFolders")
        ? json({
            value: [
              { id: "nested", displayName: "Nested", childFolderCount: 0 },
            ],
          })
        : json({
            value: [
              { id: "blocked", displayName: "Blocked", childFolderCount: 1 },
              { id: "open", displayName: "Open", childFolderCount: 1 },
            ],
          }),
  );
  const first = await f.call("mailboxes.list", { limit: 2 });
  assert.equal(
    new URL(directory(f.requests)[0].url).searchParams.get(
      "includeHiddenFolders",
    ),
    "true",
  );
  const second = await f.call("mailboxes.list", {
    limit: 2,
    cursor: first.nextCursor,
  });
  assert.equal(second.items.length, 0);
  assert.ok(second.nextCursor);
  assert.equal(second.coverage.complete, false);
  const third = await f.call("mailboxes.list", {
    limit: 2,
    cursor: second.nextCursor,
  });
  assert.equal(third.items[0].id, "nested");
  assert.equal(third.coverage.complete, false);
  assert.equal(third.coverage.reason, "some_folders_unavailable");
});

for (const native of [false, true])
  test(`${native ? "Outlook" : "Graph"} folder roles come from well-known folder names, never display names`, async (t) => {
    const ids = {
      inbox: "id-inbox",
      sentitems: "id-sent",
      drafts: "id-drafts",
      deleteditems: "id-deleted",
      archive: "id-archive",
      junkemail: "id-junk",
    };
    const f = await fixture(
      t,
      (r) => {
        const u = new URL(r.url);
        const name = /\/mailFolders\/([^/]+)$/i.exec(u.pathname)?.[1];
        if (name && !u.pathname.toLowerCase().endsWith("/childfolders")) {
          assert.equal(u.searchParams.get("$select"), native ? "Id" : "id");
          const id = ids[name.toLowerCase()];
          const body = { id };
          return id ? json(native ? pascal(body) : body) : json({}, 404);
        }
        const value = [
          { id: "id-inbox", displayName: "Postvak IN" },
          { id: "id-archive", displayName: "Archief" },
          // A user folder named "Archive" must not get the archive role.
          { id: "user-archive", displayName: "Archive" },
          { id: "id-sent", displayName: "Verzonden items" },
          { id: "id-drafts", displayName: "Concepten" },
          { id: "id-deleted", displayName: "Verwijderde items" },
          { id: "id-junk", displayName: "Ongewenste e-mail" },
          { id: "beta", displayName: "Beta", wellKnownName: "archive" },
        ];
        return json({ value: native ? pascal(value) : value });
      },
      native,
    );
    const result = await f.call("mailboxes.list", {});
    assert.deepEqual(
      Object.fromEntries(result.items.map((m) => [m.id, m.role])),
      {
        "id-inbox": "inbox",
        "id-archive": "archive",
        "user-archive": null,
        "id-sent": "sent",
        "id-drafts": "drafts",
        "id-deleted": "trash",
        "id-junk": "spam",
        // Beta-style responses that carry wellKnownName are honoured too.
        beta: "archive",
      },
    );
    assert.equal(result.items[1].kind, "folder");
    assert.equal(result.items[1].canWrite, true);
    assert.equal(f.requests.filter((r) => WELL_KNOWN.test(r.url)).length, 6);
  });

test("folder roles survive pagination without re-resolving and a missing archive folder is no role", async (t) => {
  const f = await fixture(t, (r) => {
    const u = new URL(r.url);
    if (WELL_KNOWN.test(r.url))
      return u.pathname.endsWith("/archive")
        ? json({}, 404)
        : json({ id: "id-" + u.pathname.split("/").pop() });
    return u.pathname.endsWith("/childFolders")
      ? json({ value: [{ id: "id-sentitems", displayName: "Sent" }] })
      : json({
          value: [
            { id: "id-inbox", displayName: "Inbox", childFolderCount: 1 },
          ],
        });
  });
  const first = await f.call("mailboxes.list", { limit: 1 });
  assert.equal(first.items[0].role, "inbox");
  const second = await f.call("mailboxes.list", {
    limit: 1,
    cursor: first.nextCursor,
  });
  assert.equal(second.items[0].role, "sent");
  assert.equal(f.requests.filter((r) => WELL_KNOWN.test(r.url)).length, 6);
});

test("a role lookup outage fails the listing instead of hiding roles", async (t) => {
  const f = await fixture(t, (r) =>
    WELL_KNOWN.test(r.url) ? json({}, 503) : json({ value: [] }),
  );
  await assert.rejects(() => f.call("mailboxes.list", {}), {
    code: "source_unavailable",
  });
});

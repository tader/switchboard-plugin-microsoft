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
async function fixture(t, respond, native = false, options = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "outlook-email-"));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const requests = [],
    contracts = createContracts(root, options);
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
// Role resolution batches (Graph) or reads (Outlook REST) each well-known folder; directory assertions ignore those lookups.
const WELL_KNOWN =
  /\/\$batch$|\/mailFolders\/(inbox|sentitems|drafts|deleteditems|archive|junkemail)\?/i;
const isBatch = (r) => new URL(r.url).pathname.endsWith("/$batch");
// Answers a Graph JSON batch of well-known folder reads; `lookup(name)` gives [status, body].
const batch = (r, lookup) =>
  json({
    responses: JSON.parse(r.body).requests.map((q) => {
      const name = /\/mailFolders\/([^?]+)/.exec(q.url)[1];
      const [status, body] = lookup(name, q);
      return { id: q.id, status, body };
    }),
  });
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

const ROLE_IDS = {
  inbox: "id-inbox",
  sentitems: "id-sent",
  drafts: "id-drafts",
  deleteditems: "id-deleted",
  archive: "id-archive",
  junkemail: "id-junk",
};
for (const native of [false, true])
  test(`${native ? "Outlook" : "Graph"} folder roles come from well-known folder names, never display names`, async (t) => {
    const f = await fixture(
      t,
      (r) => {
        const u = new URL(r.url);
        if (isBatch(r)) {
          assert.equal(native, false);
          assert.equal(r.method, "POST");
          assert.equal(u.href, "https://graph.microsoft.com/v1.0/$batch");
          return batch(r, (name, q) => {
            assert.equal(q.method, "GET");
            assert.equal(new URL(q.url, u).searchParams.get("$select"), "id");
            const id = ROLE_IDS[name];
            return id ? [200, { id }] : [404, { error: {} }];
          });
        }
        const name = /\/mailFolders\/([^/]+)$/i.exec(u.pathname)?.[1];
        if (name && !u.pathname.toLowerCase().endsWith("/childfolders")) {
          assert.equal(native, true);
          assert.equal(u.searchParams.get("$select"), "Id");
          const id = ROLE_IDS[name.toLowerCase()];
          return id ? json(pascal({ id })) : json({}, 404);
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
    // Graph: one batch request; Outlook REST: one read per well-known folder.
    assert.equal(
      f.requests.filter((r) => WELL_KNOWN.test(r.url)).length,
      native ? 6 : 1,
    );
  });

test("folder roles are cached per connection across pages and listings, and a missing archive folder is no role", async (t) => {
  const f = await fixture(t, (r) => {
    if (isBatch(r))
      return batch(r, (name) =>
        name === "archive" ? [404, {}] : [200, { id: "id-" + name }],
      );
    return new URL(r.url).pathname.endsWith("/childFolders")
      ? json({ value: [{ id: "id-sentitems", displayName: "Sent" }] })
      : json({
          value: [
            { id: "id-inbox", displayName: "Inbox", childFolderCount: 1 },
          ],
        });
  });
  const first = await f.call("mailboxes.list", { limit: 1 });
  assert.equal(first.items[0].role, "inbox");
  // The cursor no longer carries the role map.
  const second = await f.call("mailboxes.list", {
    limit: 1,
    cursor: first.nextCursor,
  });
  assert.equal(second.items[0].role, "sent");
  const again = await f.call("mailboxes.list", { limit: 1 });
  assert.equal(again.items[0].role, "inbox");
  assert.equal(f.requests.filter(isBatch).length, 1);
});

test("a complete role map expires after its TTL and is refreshed without blocking", async (t) => {
  let clock = 1000,
    archiveId = "old-archive";
  const f = await fixture(
    t,
    (r) =>
      isBatch(r)
        ? batch(r, (name) =>
            name === "archive" ? [200, { id: archiveId }] : [404, {}],
          )
        : json({ value: [{ id: "new-archive" }, { id: "old-archive" }] }),
    false,
    { now: () => clock, roleTtlMs: 10000 },
  );
  const roleOf = (page, id) => page.items.find((m) => m.id === id).role;
  assert.equal(
    roleOf(await f.call("mailboxes.list"), "old-archive"),
    "archive",
  );
  clock += 9999;
  await f.call("mailboxes.list");
  assert.equal(f.requests.filter(isBatch).length, 1);
  clock += 2;
  archiveId = "new-archive";
  // Expired: the previous map is served while the refresh runs.
  const stale = await f.call("mailboxes.list");
  assert.equal(roleOf(stale, "old-archive"), "archive");
  assert.equal(f.requests.filter(isBatch).length, 2);
  const fresh = await f.call("mailboxes.list");
  assert.equal(roleOf(fresh, "new-archive"), "archive");
  assert.equal(roleOf(fresh, "old-archive"), null);
});

for (const failure of [
  ["an outage", () => json({}, 503)],
  ["a network error", () => Promise.reject(new TypeError("fetch failed"))],
  ["a throttled sub-request", null],
])
  test(`role lookup failure (${failure[0]}) never fails the listing and is cached only briefly`, async (t) => {
    let clock = 0;
    const warnings = [];
    const f = await fixture(
      t,
      (r) =>
        isBatch(r)
          ? failure[1]
            ? failure[1]()
            : batch(r, (name) =>
                name === "inbox" ? [200, { id: "id-inbox" }] : [429, {}],
              )
          : json({ value: [{ id: "id-inbox", displayName: "Inbox" }] }),
      false,
      {
        now: () => clock,
        roleFailureTtlMs: 120000,
        log: { warn: (...a) => warnings.push(a) },
      },
    );
    const result = await f.call("mailboxes.list");
    assert.equal(result.items[0].role, failure[1] ? null : "inbox");
    assert.equal(result.coverage.complete, true);
    if (failure[1]) assert.equal(warnings.length, 1);
    clock += 119999;
    await f.call("mailboxes.list");
    assert.equal(f.requests.filter(isBatch).length, 1);
    clock += 2;
    await f.call("mailboxes.list");
    assert.equal(f.requests.filter(isBatch).length, 2);
  });

for (const native of [false, true])
  test(`${native ? "Outlook" : "Graph"} hanging role lookups stay within the time budget and keep roles found so far`, async (t) => {
    const hang = new Promise(() => {});
    const f = await fixture(
      t,
      (r) => {
        const u = new URL(r.url);
        if (isBatch(r)) return hang;
        const name = /\/mailFolders\/([^/]+)$/i.exec(u.pathname)?.[1];
        if (name && !u.pathname.toLowerCase().endsWith("/childfolders"))
          return name === "inbox" ? json(pascal({ id: "id-inbox" })) : hang;
        return json({
          value: pascal([
            { id: "id-inbox", displayName: "Inbox" },
            { id: "id-archive", displayName: "Archive" },
            { id: "beta", displayName: "Beta", wellKnownName: "archive" },
          ]),
        });
      },
      native,
      { roleBudgetMs: 100 },
    );
    const started = Date.now();
    const result = await f.call("mailboxes.list");
    const elapsed = Date.now() - started;
    assert.ok(elapsed >= 90 && elapsed < 1000, `took ${elapsed} ms`);
    assert.deepEqual(
      result.items.map((m) => m.role),
      // Graph's single batch hangs entirely; Outlook REST found the inbox before the budget ran out.
      [native ? "inbox" : null, null, "archive"],
    );
    // A lookup that is still running is shared, not repeated.
    await f.call("mailboxes.list");
    assert.equal(
      f.requests.filter((r) => WELL_KNOWN.test(r.url)).length,
      // Outlook REST: at most 3 reads in flight (inbox answered, then 3 hang).
      native ? 4 : 1,
    );
  });

test("the directory read and the role lookup run concurrently", async (t) => {
  let release;
  const gate = new Promise((resolve) => (release = resolve));
  const order = [];
  const f = await fixture(t, async (r) => {
    if (isBatch(r)) {
      order.push("batch");
      await gate;
      return batch(r, () => [404, {}]);
    }
    order.push("directory");
    release();
    return json({ value: [] });
  });
  await f.call("mailboxes.list");
  assert.deepEqual(order, ["batch", "directory"]);
});

test("a directory read failure still fails the listing", async (t) => {
  const f = await fixture(t, (r) =>
    isBatch(r) ? batch(r, () => [404, {}]) : json({}, 503),
  );
  await assert.rejects(() => f.call("mailboxes.list", {}), {
    code: "source_unavailable",
  });
});

for (const native of [false, true])
  test(`Outlook ${native ? "REST" : "Graph"} retries explicit throttling with backoff`, async (t) => {
    let attempts = 0;
    const f = await fixture(t, () => ++attempts <= 2 ? json({}, 429) : json({ id: "m" }), native);
    const started = Date.now();
    assert.deepEqual(await f.call("messages.move", { id: "m", destinationId: "archive" }), { status: "accepted", resourceId: "m" });
    assert.equal(attempts, 3);
    assert.ok(Date.now() - started >= 650, "retries must back off");
    assert.ok(f.requests.every((r) => r.method === "POST"));
    assert.equal(new Set(f.requests.map((r) => r.body)).size, 1, "repeat the exact same operation");
  });

test("Outlook never retries a refusal, ambiguous server error or dropped connection", async (t) => {
  const denied = await fixture(t, () => json({}, 403));
  assert.deepEqual(await denied.call("messages.update", { id: "m", changes: { isRead: true } }), { status: "rejected", code: "permission_denied" });
  assert.equal(denied.requests.length, 1);
  const unavailable = await fixture(t, () => json({}, 503));
  await assert.rejects(() => unavailable.call("messages.update", { id: "m", changes: { isRead: true } }), { code: "source_unavailable" });
  assert.equal(unavailable.requests.length, 1);
  const disconnected = await fixture(t, () => { throw new Error("fixture connection lost"); });
  await assert.rejects(() => disconnected.call("messages.update", { id: "m", changes: { isRead: true } }));
  assert.equal(disconnected.requests.length, 1);
});

test("Outlook retries are bounded and honor Retry-After seconds and dates", async (t) => {
  const limited = await fixture(t, () => json({}, 429));
  await assert.rejects(() => limited.call("messages.get", { id: "m" }), { code: "rate_limited", status: 429 });
  assert.equal(limited.requests.length, 4);
  for (const header of ["120", new Date(Date.now() + 120000).toUTCString()]) {
    const longWait = await fixture(t, () => json({}, 429, { "retry-after": header }));
    await assert.rejects(() => longWait.call("messages.get", { id: "m" }), (e) => e.code === "rate_limited" && e.retryAfterMs > 100000);
    assert.equal(longWait.requests.length, 1);
  }
});

test("Outlook cancellation during backoff prevents another request", async (t) => {
  const f = await fixture(t, () => json({}, 429));
  const controller = new AbortController(); f.ctx.signal = controller.signal;
  const call = f.call("messages.update", { id: "m", changes: { isRead: true } });
  setTimeout(() => controller.abort(), 20);
  await assert.rejects(call);
  assert.equal(f.requests.length, 1);
});

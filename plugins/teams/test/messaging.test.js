import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { Ajv } from "ajv";
import { fixture, chatId, channelId } from "./contracts-fixture.js";
import { messagingOperations } from "../lib/messaging-schema.js";
const ajv = new Ajv({ strict: false });
async function call(f, name, input, ctx = f.invocation()) {
  const op = f.service.contracts.operations[name];
  for (const schema of [
    messagingOperations[name].inputSchema,
    op.inputSchema,
  ]) {
    const valid = ajv.compile(schema);
    assert.ok(valid(input), JSON.stringify(valid.errors));
  }
  const result = await op.invoke(input, ctx),
    valid = ajv.compile(op.outputSchema);
  assert.ok(valid(result), JSON.stringify(valid.errors));
  return result;
}
test("common reads cover chats/channels and preserve partial/restricted/deleted content", async (t) => {
  const f = await fixture(t),
    contract = f.service.contracts;
  assert.equal(contract.watch, undefined);
  const list = await call(f, "messaging.conversations.list", {});
  assert.equal(list.coverage.complete, false);
  assert.equal(list.items[0].kind, "direct");
  assert.equal(list.items[1].kind, "channel");
  assert.equal(list.items[0].unread, null);
  const page = await call(f, "messaging.messages.list", {
    conversationId: chatId,
  });
  assert.equal(page.coverage.complete, false);
  assert.equal(page.items.find((m) => m.id === "100").isOwn, true);
  assert.equal(page.items.find((m) => m.id === "101").text, null);
  assert.equal(
    page.items.find((m) => m.id === "101").contentState,
    "restricted",
  );
  assert.equal(page.items.find((m) => m.id === "102").contentState, "deleted");
  assert.ok(!JSON.stringify(page).includes("PRIVATE-"));
  await assert.rejects(
    call(f, "messaging.messages.list", {
      conversationId: channelId,
      cursor: page.nextCursor,
    }),
    (e) => e.code === "invalid_cursor",
  );
  const second = await call(f, "messaging.messages.list", {
    conversationId: chatId,
    cursor: page.nextCursor,
  });
  assert.equal(second.nextCursor, undefined);
  const get = await call(f, "messaging.messages.get", {
    conversationId: chatId,
    id: "100",
  });
  assert.equal(get.text, "Root");
  assert.ok(
    !f.requests.some((r) => r.options.method && r.options.method !== "GET"),
  );
});
test("common sends replay the durable native ledger and block conflicts and uncertain redispatch", async (t) => {
  const f = await fixture(t),
    input = { conversationId: chatId, text: "PRIVATE-OUTGOING" };
  const first = await call(f, "messaging.messages.send", input);
  assert.equal(first.status, "accepted");
  assert.deepEqual(await call(f, "messaging.messages.send", input), first);
  assert.equal(f.posts.length, 1);
  await assert.rejects(
    call(f, "messaging.messages.send", { ...input, text: "changed" }),
    (e) => e.code === "idempotency_conflict",
  );
  f.failSend(true);
  for (let n = 0; n < 2; n++)
    assert.deepEqual(
      await call(
        f,
        "messaging.messages.send",
        input,
        f.invocation(f.connection, "uncertain-key"),
      ),
      { status: "uncertain" },
    );
  assert.equal(f.posts.length, 2);
  const files = await fs.readdir(path.join(f.directory, "profiles"), {
    recursive: true,
  });
  for (const file of files.filter((x) => x.endsWith("send-ledger.json"))) {
    const ledger = await fs.readFile(
      path.join(f.directory, "profiles", file),
      "utf8",
    );
    assert.ok(!ledger.includes("PRIVATE-OUTGOING"));
  }
});
test("thread history and replies retain exact roots through opaque message ids", async (t) => {
  const f = await fixture(t);
  const page = await call(f, "messaging.messages.list", {
    conversationId: channelId,
    threadId: "100",
  });
  assert.equal(page.coverage.complete, false);
  const sent = await call(f, "messaging.messages.send", {
    conversationId: channelId,
    threadId: "100",
    text: "Reply",
  });
  assert.equal(sent.status, "accepted");
  assert.notEqual(sent.resourceId, "200");
  const get = await call(f, "messaging.messages.get", {
    conversationId: channelId,
    id: sent.resourceId,
  });
  assert.equal(get.threadId, "100");
  assert.equal(get.text, "Reply");
  await assert.rejects(
    call(f, "messaging.messages.send", {
      conversationId: channelId,
      text: "No implicit new thread",
    }),
    (e) => e.code === "invalid_request",
  );
});
test("read-only guards include raw writes; auth expiry is terminal and resource denial is not", async (t) => {
  const f = await fixture(t),
    ro = { ...f.connection, config: { access: "read" } };
  assert.equal(f.service.contracts.available(ro).length, 4);
  await assert.rejects(
    call(
      f,
      "messaging.messages.send",
      { conversationId: chatId, text: "blocked" },
      f.invocation(ro),
    ),
    (e) => e.code === "permission_denied",
  );
  await assert.rejects(
    f
      .invocation(ro)
      .request({
        method: "POST",
        url: "/chats/" + encodeURIComponent(chatId) + "/messages",
        body: "{}",
      }),
    (e) => e.code === "permission_denied",
  );
  f.expire(true);
  const count = f.requests.length;
  await assert.rejects(
    f.service.contracts.collect({}, f.invocation()),
    (e) => e.code === "authentication_required" && e.reauthenticationRequired,
  );
  assert.equal(f.requests.length, count);
  f.expire(false);
  f.deny(true);
  await assert.rejects(
    call(f, "messaging.messages.list", { conversationId: chatId }),
    (e) => e.code === "permission_denied" && !e.reauthenticationRequired,
  );
  f.deny(false);
  f.rate(true);
  await assert.rejects(
    f.service.contracts.collect({}, f.invocation()),
    (e) => e.code === "rate_limited" && e.retryAfterMs === 123000,
  );
});

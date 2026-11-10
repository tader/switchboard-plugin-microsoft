import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { Ajv } from "ajv";
import { fixture, chatId, channelId } from "./contracts-fixture.js";
import { project } from "../lib/messaging.js";
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

for (const method of ["browser", "tokens", "authentication-source"]) {
 test(`Messaging discovery and requests preserve access with ${method}`, async t => {
  const f=await fixture(t,method), contract=f.service.contracts;
  assert.equal(contract.available(f.connection).length,6);
  const list=await call(f,"messaging.conversations.list",{});
  assert.equal(list.items[0].kind,"direct");
  const history=await call(f,"messaging.messages.list",{conversationId:chatId});
  assert.equal(history.items[0].text,"Root");
  const sent=await call(f,"messaging.messages.send",{conversationId:chatId,text:"source fixture"});
  assert.equal(sent.status,"accepted");assert.equal(f.posts.length,1);
  const read={...f.connection,config:{...f.connection.config,access:"read"}};
  assert.equal(contract.available(read).length,4);
  await assert.rejects(call(f,"messaging.messages.send",{conversationId:chatId,text:"blocked"},f.invocation(read)),{code:"permission_denied"});
  await assert.rejects(f.invocation(read).request({method:"POST",url:"/chats/"+encodeURIComponent(chatId)+"/messages",body:"{}"}),{code:"permission_denied"});
  assert.equal(f.posts.length,1);
  f.expire(true);const before=f.requests.length;
  await assert.rejects(contract.collect({},f.invocation()),e=>e.code==="authentication_required"&&e.reauthenticationRequired);
  assert.equal(f.requests.length,before);
 });
}
test("shared authentication requires its own credential mode and existing profile/identity", async t => {
 const f=await fixture(t,"authentication-source"), available=f.service.contracts.available;
 assert.deepEqual(available({...f.connection,credentials:{...f.connection.credentials,authMode:"tokens"}}),[]);
 assert.deepEqual(available({...f.connection,credentials:{...f.connection.credentials,profile:null}}),[]);
 assert.deepEqual(available({...f.connection,credentials:{...f.connection.credentials,identity:null}}),[]);
 assert.deepEqual(available({...f.connection,methodId:"unknown"}),[]);
});

test("mixed chat/channel discovery stays bounded and pages every channel exactly once",async t=>{
 const f=await fixture(t,"authentication-source");f.setChannelCount(130);
 let cursor;const ids=[];let pages=0;
 do{const page=await call(f,"messaging.conversations.list",{limit:20,...(cursor?{cursor}:{})});
 assert.ok(page.items.length<=20);ids.push(...page.items.map(c=>c.id));cursor=page.nextCursor;pages++;assert.ok(pages<10);
 }while(cursor);
 assert.equal(ids.length,131);assert.equal(new Set(ids).size,131);assert.ok(ids.includes(chatId));
});
test("discovery cursors remain bound to query/account and reject changed channel membership",async t=>{
 const f=await fixture(t,"authentication-source");f.setChannelCount(5);
 const page=await call(f,"messaging.conversations.list",{limit:2});
 await assert.rejects(call(f,"messaging.conversations.list",{limit:3,cursor:page.nextCursor}),{code:"invalid_cursor"});
 const other={...f.connection,id:"other-account"};
 await assert.rejects(call(f,"messaging.conversations.list",{limit:2,cursor:page.nextCursor},f.invocation(other)),{code:"invalid_cursor"});
 // The provider directory cache may last 30s; use another adapter with the same
 // signing key to exercise a refreshed snapshot without sleeping in this test.
 const {createMessaging}=await import("../lib/messaging.js");
 const contract=createMessaging(f.directory);
 const ctx=f.invocation();const original=ctx.request;
 ctx.request=async req=>req.url==="/channels"?Response.json({items:[]}):original(req);
 await assert.rejects(contract.operations["messaging.conversations.list"].invoke({limit:2,cursor:page.nextCursor},ctx),{code:"invalid_cursor"});
});

test("common discovery, history and durable sends support the account-private self conversation",async t=>{
 const f=await fixture(t,"authentication-source");f.includeSelfNotes();
 const {encodeConversation}=await import("../lib/teams-api.js");const id=encodeConversation("chat","48:notes");
 const list=await call(f,"messaging.conversations.list",{});
 assert.equal(list.items.find(c=>c.id===id).title,"Notes to self");
 assert.equal(list.items.find(c=>c.id===id).kind,"direct");
 const history=await call(f,"messaging.messages.list",{conversationId:id});assert.equal(history.items[0].isOwn,true);
 const input={conversationId:id,text:"Self fixture"};const result=await call(f,"messaging.messages.send",input);
 assert.equal(result.status,"accepted");assert.deepEqual(await call(f,"messaging.messages.send",input),result);assert.equal(f.posts.length,1);
});

test("common reaction writes use the chat's native durable action",async t=>{
 const f=await fixture(t);const input={conversationId:chatId,messageId:"100",emoji:"👍",remove:false};
 assert.deepEqual(await call(f,"messaging.reactions.update",input),{status:"accepted"});
 const writes=f.requests.filter(r=>r.options.method==="PUT"&&r.url.includes("name=emotions"));assert.equal(writes.length,1);
 assert.deepEqual(await call(f,"messaging.reactions.update",input),{status:"accepted"});assert.equal(f.requests.filter(r=>r.options.method==="PUT"&&r.url.includes("name=emotions")).length,1);
 assert.deepEqual(await call(f,"messaging.reactions.update",{...input,conversationId:channelId},f.invocation(f.connection,"channel-key")),{status:"rejected",code:"unsupported"});
});

test("reaction projection retains unknown users instead of inventing ownership",()=>{
 const row=project({id:"100",text:"Hi",time:"2026-11-03T09:00:00Z",reactions:[{key:"like",users:[{mri:"self"},{mri:"other"}]},{key:"unknown"}]},chatId,"self");
 assert.deepEqual(row.reactions,[{emoji:"👍",count:2,me:true,users:[{id:"self",displayName:null},{id:"other",displayName:null}],usersComplete:true},{emoji:":unknown:",count:null,me:null}]);
});

test("reaction names use observed identities without guessing incomplete people",()=>{
 const row=project({id:"100",text:"Hi",authorId:"other",author:"Alice",reactions:[{key:"like",users:[{mri:"other"},{mri:"self"},{noId:true}]}]},chatId,"self",new Map([["self","Me"]]));
 assert.deepEqual(row.reactions[0].users,[{id:"other",displayName:"Alice"},{id:"self",displayName:"Me"}]);
 assert.equal(row.reactions[0].usersComplete,false);
 assert.equal(row.reactions[0].me,true);
});
test("empty provider reaction entries do not appear as active reactions",()=>{
 const row=project({id:"100",text:"Hi",reactions:[{key:"like",users:[]}]},chatId,"self");
 assert.deepEqual(row.reactions,[]);
});

test("malformed reaction metadata does not erase a previously known state",()=>{
 const row=project({id:"100",content:"Hi",messagetype:"Text",properties:{emotions:"{invalid"}},chatId,"self");
 assert.equal(row.reactions,undefined);
});

// Real core HTTP/SSE and existing Teams HTTP adapter; provider traffic is fixture-only.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import assert from "node:assert/strict";
import { pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { fixture, chatId } from "../test/contracts-fixture.js";
const core = process.env.SWITCHBOARD_ROOT;
assert(core, "Set SWITCHBOARD_ROOT");
const root = path.resolve(core),
  temp = fs.mkdtempSync(path.join(os.tmpdir(), "teams-contract-http-"));
process.env.SWITCHBOARD_DATA_DIR = temp;
process.env.SWITCHBOARD_SHARED_CONTRACTS = "1";
const load = (file) => import(pathToFileURL(path.join(root, file)).href);
const database = await load("server/db.ts");
database.initDb();
const { initKey, encrypt } = await load("server/crypto.ts");
initKey();
const { createUser, createToken } = await load("server/users.ts"),
  { plugins } = await load("server/plugins/manager.ts"),
  { api } = await load("server/routes/connections.ts"),
  { collectors, pushSources, startContractCollectors } = await load(
    "server/contracts/service.ts",
  );
const cleanup = [],
  f = await fixture({ after: (fn) => cleanup.push(fn) }, "authentication-source");
(await load("server/contracts/model.ts")).validateDefinition(
  f.service.contracts,
);
plugins.services.set("teams-web", {
  pluginId: "teams-web",
  service: f.service,
});
assert.equal(
  fs.readFileSync(new URL("../lib/messaging-v1.json", import.meta.url), "utf8"),
  fs.readFileSync(
    path.join(root, "server/contracts/messaging-v1.json"),
    "utf8",
  ),
);
const { serve } = createRequire(path.join(root, "package.json"))(
  "@hono/node-server",
);
let ready;
const listening = new Promise((r) => (ready = r));
const server = serve(
  { fetch: api.fetch, hostname: "127.0.0.1", port: 0 },
  ready,
);
await listening;
const stop = startContractCollectors();
const base = "http://127.0.0.1:" + server.address().port,
  user = createUser("fixture", "admin"),
  id = "teams-fixture",
  ro = "teams-readonly";
// Exercise the real credential-consumer boundary too: cached shared sessions
// still require a selected, owned authentication provider in the core catalogue.
plugins.services.set("fixture-authentication", {
 pluginId:"fixture-authentication",
 service:{id:"fixture-authentication",name:"Fixture authentication",authMethods:[{
  id:"fixture",name:"Fixture",fields:[],authorize:async()=>{},
  credentialProvider:{format:"microsoft/web-session/v1",provide:async()=>{throw Error("Unexpected export: fixture uses cached session");}},
 }]},
});
database.run("INSERT INTO connections(id,user_id,service_id,method_id,name,config_enc,credentials_enc,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
 "source-fixture",user.id,"fixture-authentication","fixture","source-fixture",encrypt({}),encrypt({}),"ok",Date.now(),Date.now());
for (const [key, access] of [
  [id, "write"],
  [ro, "read"],
])
  database.run(
    "INSERT INTO connections(id,user_id,service_id,method_id,name,config_enc,credentials_enc,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
    key,
    user.id,
    "teams-web",
    f.connection.methodId,
    key,
    encrypt({ access, authenticationSource: "source-fixture" }),
    encrypt(f.connection.credentials),
    "ok",
    Date.now(),
    Date.now(),
  );
const token = createToken(user.id, "fixture", [id]).secret,
  roToken = createToken(user.id, "readonly", [ro]).secret;
const request = (
  url,
  body,
  auth = token,
  method = body === undefined ? "GET" : "POST",
) =>
  fetch(base + url, {
    method,
    headers: {
      authorization: "Bearer " + auth,
      "content-type": "application/json",
    },
    ...(body === undefined ? {} : { body: JSON.stringify(body) }),
  });
const call = (operation, args, key, conn = id, auth = token) =>
  request(
    `/connections/${conn}/contracts/v1/call`,
    { operation, arguments: args, ...(key ? { idempotencyKey: key } : {}) },
    auth,
  );
try {
  const descriptor = await (
    await request(`/connections/${id}/contracts/v1`)
  ).json();
  assert.equal(descriptor.operations.length, 5);
  const discovery = await (await request("/contracts/v1/connections?category=messaging")).json();
  assert.equal(discovery.connections[0].id, id);
  assert.deepEqual(descriptor.changes, { native: false, reconciliation: true });
  assert.equal((await request(`/connections/${ro}/contracts/v1`)).status, 403);
  assert.equal(
    (
      await call(
        "messaging.messages.send",
        { conversationId: chatId, text: "No" },
        "readonly",
        ro,
        roToken,
      )
    ).status,
    403,
  );
  const list = await (await call("messaging.conversations.list", {})).json();
  assert.equal(list.items?.length, 2, JSON.stringify(list));
  assert.equal(list.coverage.complete, false);
  const history = await (
    await call("messaging.messages.list", { conversationId: chatId })
  ).json();
  assert.equal(history.items.find((m) => m.id === "101").text, null);
  assert.ok(!JSON.stringify(history).includes("PRIVATE-"));
  for (let n = 0; n < 2; n++)
    assert.equal(
      (
        await (
          await call(
            "messaging.messages.send",
            { conversationId: chatId, text: "PRIVATE-SEND" },
            "same-key",
          )
        ).json()
      ).status,
      "accepted",
    );
  assert.equal(f.posts.length, 1);
  f.failSend(true);
  for (let n = 0; n < 2; n++)
    assert.equal(
      (
        await (
          await call(
            "messaging.messages.send",
            { conversationId: chatId, text: "Unknown" },
            "uncertain",
          )
        ).json()
      ).status,
      "uncertain",
    );
  assert.equal(f.posts.length, 2);
  f.failSend(false);
  const sub = await (
    await request("/contracts/v1/subscriptions", { connections: [id] })
  ).json();
  await request(`/contracts/v1/subscriptions/${sub.id}/activity`, {
    active: true,
  });
  const refresh = await (
    await request(`/contracts/v1/subscriptions/${sub.id}/refresh`, {
      connectionId: id,
    })
  ).json();
  assert.equal(refresh.status, "completed");
  const abort = new AbortController(),
    timer = setTimeout(() => abort.abort(), 3000);
  try {
    const res = await fetch(
      base +
        `/contracts/v1/subscriptions/${sub.id}/events?cursor=${sub.cursor}`,
      { headers: { authorization: "Bearer " + token }, signal: abort.signal },
    );
    const reader = res.body.getReader();
    let text = "";
    try {
      while (!text.includes('"type":"invalidated"')) {
        const chunk = await reader.read();
        assert.equal(chunk.done, false);
        text += new TextDecoder().decode(chunk.value);
      }
      assert.ok(!text.includes("PRIVATE-"));
    } finally {
      await reader.cancel();
    }
  } finally {
    clearTimeout(timer);
    abort.abort();
  }
  f.expire(true);
  const reads = f.requests.length;
  assert.equal(
    (await call("messaging.messages.list", { conversationId: chatId })).status,
    409,
  );
  assert.equal(collectors().state(id).status, "auth_required");
  for (let n = 0; n < 2; n++)
    assert.equal(
      (
        await (
          await request(`/contracts/v1/subscriptions/${sub.id}/refresh`, {
            connectionId: id,
          })
        ).json()
      ).status,
      "auth_required",
    );
  assert.equal(f.requests.length, reads);
  f.expire(false);
  const { saveCredentials } = await load("server/connections.ts");
  saveCredentials(id, {
    ...f.connection.credentials,
    sessionRevision: "verified-new-auth",
  });
  assert.notEqual(collectors().state(id).status, "auth_required");
  f.deny(true);
  assert.equal(
    (await call("messaging.messages.list", { conversationId: chatId })).status,
    403,
  );
  assert.notEqual(collectors().state(id).status, "auth_required");
  assert.ok(
    !JSON.stringify(database.all("SELECT * FROM contract_receipts")).includes(
      "PRIVATE-SEND",
    ),
  );
  assert.ok(
    !JSON.stringify(database.all("SELECT * FROM audit_log")).includes(
      "PRIVATE-SEND",
    ),
  );
  await request(
    `/contracts/v1/subscriptions/${sub.id}`,
    undefined,
    token,
    "DELETE",
  );
  console.log(
    "PASS: Teams canonical HTTP/SSE, read-only scope, restricted content, durable writes, terminal auth pause without provider retry, credential recovery and resource-specific 403 isolation; offline only",
  );
} finally {
  await stop();
  server.closeAllConnections();
  await new Promise((r) => server.close(r));
  for (const fn of cleanup) await fn();
  database.db.close();
  fs.rmSync(temp, { recursive: true, force: true });
}

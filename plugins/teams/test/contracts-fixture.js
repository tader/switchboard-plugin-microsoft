import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createAdapter } from "../index.js";
import { TeamsAPI, encodeConversation } from "../lib/teams-api.js";
import { AdapterError } from "../lib/errors.js";
export const oid = "11111111-1111-4111-8111-111111111111",
  tenant = "22222222-2222-4222-8222-222222222222",
  profile = "33333333-3333-4333-8333-333333333333",
  thread = "19:chat@thread.v2",
  channel = "19:channel@thread.tacv2",
  chatId = encodeConversation("chat", thread),
  channelId = encodeConversation("channel", channel);
export async function fixture(t, methodId = "tokens") {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "teams-contract-"));
  const requests = [],
    posts = [];
  let channelCount = 1;
  let includeSelfNotes = false;
  let expired = false,
    sendFailure = false,
    forbidden = false,
    rate = false;
  const credentials = {
    profile,
    authVersion: 1,
    identity: { oid, tenant, name: "Fixture", email: "fixture@example.com" },
    messageOrigin: "https://emea.ng.msg.teams.microsoft.com",
    region: "emea",
    chatToken: { value: "PRIVATE-CHAT-TOKEN", expiresAt: Date.now() + 3600000 },
    tokens: {
      skype: { value: "PRIVATE-SKYPE" },
      chatsvcagg: { value: "PRIVATE-CSA" },
    },
    authMode: methodId === "authentication-source" ? "microsoft" : methodId,
  };
  const root = {
    id: "100",
    content: "Root",
    messagetype: "Text",
    from: "8:orgid:" + oid,
    originalarrivaltime: "2026-11-03T09:00:00Z",
    version: "v1",
    properties: {},
  };
  const fetcher = async (url, options = {}) => {
    requests.push({ url: String(url), options });
    const u = new URL(url),
      p = decodeURIComponent(u.pathname);
    if (rate)
      return new Response("{}", {
        status: 429,
        headers: { "retry-after": "123" },
      });
    if (p.endsWith("/teams/users/me"))
      return Response.json({
        chats: [{ id: thread, title: "Chat", isOneOnOne: true, members: [] }],
        teams: [
          {
            channels: Array.from({length:channelCount},(_,n)=>({id:n?channel+"-"+n:channel,displayName:"Channel "+n,isMember:true})),
          },
        ],
      });
    if (p.endsWith("/users/ME/conversations/48:notes"))return Response.json({id:"48:notes",threadProperties:{threadType:"streamofnotes"}});
    if (p.endsWith("/users/ME/conversations"))
      return Response.json({
        conversations: [
          ...(includeSelfNotes?[{id:"48:notes",threadProperties:{threadType:"streamofnotes"}}]:[]),
          {
            id: thread,
            threadProperties: { threadType: "chat", topic: "Chat" },
            properties: {},
          },
        ],
        _metadata: {},
      });
    if (p.includes("/messages/") && p.endsWith("/properties") && ["PUT","DELETE"].includes(options.method)) {
      const key=JSON.parse(options.body).emotions.key;
      root.properties.emotions=options.method==="PUT"?[{key,users:[{mri:"8:orgid:"+oid}]}]:[];
      return Response.json({});
    }
    if (p.includes("/messages/") && options.method === undefined) {
      const id = p.split("/").at(-1);
      const sent = posts.find((x) => x.id === id);
      return Response.json(sent?.raw ?? { ...root, id });
    }
    if (p.endsWith("/messages") && options.method === "POST") {
      const payload = JSON.parse(options.body),
        id = String(200 + posts.length);
      posts.push({
        id,
        raw: {
          ...root,
          id,
          content: payload.content,
          messagetype: payload.messagetype,
          clientmessageid: payload.clientmessageid,
          ...(p.includes(";messageid=") ? { parentMessageId: "100" } : {}),
        },
      });
      if (sendFailure) throw Error("ambiguous provider disconnect");
      return Response.json({ id });
    }
    if (p.endsWith("/messages")) {
      if (forbidden) return new Response("{}", { status: 403 });
      return Response.json({
        messages: [
          root,
          {
            ...root,
            id: "101",
            content: "PRIVATE-RESTRICTED",
            properties: { isRestricted: true },
          },
          {
            ...root,
            id: "102",
            content: "PRIVATE-DELETED",
            messagetype: "MessageDelete",
          },
        ],
        _metadata: {
          backwardLink: u.searchParams.has("older")
            ? null
            : u.origin + u.pathname + "?older=1",
        },
      });
    }
    throw Error("Unexpected fixture route " + p);
  };
  const adapter = await createAdapter(
    { dataDir: directory, settings: {} },
    {
      createSession: (id) => ({
        profile: id,
        credentials: null,
        queue: Promise.resolve(),
        adopt(c) {
          this.credentials = c;
        },
        async ensure() {
          if (expired)
            throw new AdapterError(
              "token_import_required",
              "Reconnect fixture",
              503,
            );
          return this.credentials;
        },
        async serial(fn) {
          const task = this.queue.then(fn);
          this.queue = task.catch(() => {});
          return task;
        },
        status() {
          return { ready: !expired };
        },
        async close() {},
      }),
      createAPI: (s) => new TeamsAPI(s, { fetcher }),
    },
  );
  t.after(async () => {
    await adapter.dispose();
    await fs.rm(directory, { recursive: true, force: true });
  });
  const service = adapter.services[0],
    connection = {
      id: "fixture-connection",
      methodId,
      serviceId: "teams-web",
      credentials,
      config: { access: "write" },
    };
  const invocation = (conn = connection, key = "fixture-key") => ({
    connection: conn,
    idempotencyKey: key,
    signal: new AbortController().signal,
    async request(input) {
      const req = {
        method: input.method ?? "GET",
        url: new URL(input.url, "http://teams.localhost"),
        headers: new Headers(input.headers),
        body: input.body,
      };
      await service.authMethods
        .find((m) => m.id === conn.methodId)
        .authorize(req, conn);
      return fetch(req.url, {
        method: req.method,
        headers: req.headers,
        body: req.body,
      });
    },
  });
  return {
    directory,
    adapter,
    service,
    connection,
    invocation,
    requests,
    posts,
    expire: (v) => (expired = v),
    failSend: (v) => (sendFailure = v),
    deny: (v) => (forbidden = v),
    rate: (v) => (rate = v),
    setChannelCount: n => {channelCount=n;},
    includeSelfNotes: () => {includeSelfNotes=true;},
  };
}

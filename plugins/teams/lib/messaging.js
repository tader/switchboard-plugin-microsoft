import { createHash, randomUUID } from "node:crypto";
import { messagingOperations } from "./messaging-schema.js";
import { MessagingCursors } from "./messaging-cursors.js";
import { AdapterError } from "./errors.js";
import { normalize, arrivalMillis } from "./teams-api.js";
const REACTIONS = {"👍":"like","❤️":"heart","😂":"laugh","😮":"surprised","😢":"sad","😡":"angry"};
const REACTION_EMOJI = Object.fromEntries(Object.entries(REACTIONS).map(([emoji,name]) => [name,emoji]));
const fail = (code, status = 400) => new AdapterError(code, code, status);
const terminal = new Set([
  "token_import_required",
  "signin_required",
  "signin_failed",
  "api_unauthorized",
  "account_mismatch",
]);
function decode(id) {
  try {
    const value = JSON.parse(Buffer.from(id, "base64url"));
    if (
      !["chat", "channel"].includes(value.kind) ||
      typeof value.threadId !== "string" ||
      !value.threadId ||
      value.parentMessageId
    )
      throw Error();
    return value;
  } catch {
    throw fail("invalid_request");
  }
}
function date(value) {
  const n = arrivalMillis(value),
    d = new Date(n);
  return Number.isFinite(d.valueOf()) ? d.toISOString() : null;
}
function mid(id, parent) {
  return parent && parent !== id
    ? Buffer.from(JSON.stringify([id, parent])).toString("base64url")
    : id;
}
function target(id) {
  if (typeof id === "string" && /^\d{1,20}$/.test(id)) return { id };
  try {
    const [value, parent] = JSON.parse(Buffer.from(id, "base64url"));
    if (!/^\d{1,20}$/.test(value) || !/^\d{1,20}$/.test(parent)) throw Error();
    return { id: value, parent };
  } catch {
    throw fail("invalid_request");
  }
}
export function project(raw, conversationId, self) {
  const m = raw.messagetype || raw.messageType ? normalize(raw) : raw;
  const kind = raw.messagetype ?? raw.messageType;
  const restricted = [
    raw.isRestricted,
    raw.isHidden,
    raw.properties?.isRestricted,
  ].some((v) => v === true || v === "true" || v === "True");
  const state = m.deleted
    ? "deleted"
    : restricted
      ? "restricted"
      : (m.contentState ??
        (typeof m.text === "string" &&
        (!kind || /^(Text|RichText(?:\/.*)?)$/.test(kind))
          ? "available"
          : "unavailable"));
  const parent = m.parentMessageId;
  let reactionMetadataValid=true;
  if(raw.messagetype || raw.messageType){
    let value=raw.properties?.emotions;
    if(typeof value==="string"){try{value=JSON.parse(value)}catch{reactionMetadataValid=false}}
    if(value!==undefined&&!Array.isArray(value))reactionMetadataValid=false;
  }
  return {
    id: mid(m.id, parent),
    conversationId,
    text: state === "available" ? m.text : null,
    authorId: m.authorId ?? null,
    authorDisplayName: m.author?.trim().slice(0, 512) || null,
    createdAt: date(m.time),
    isOwn: m.authorId && self ? m.authorId === self : null,
    ...(reactionMetadataValid && Array.isArray(m.reactions) ? {reactions:m.reactions.filter(r=>typeof r?.key==="string").map(r=>({emoji:REACTION_EMOJI[r.key]??`:${r.key}:`,count:Array.isArray(r.users)?r.users.length:null,me:Array.isArray(r.users)&&self?r.users.some(u=>u?.mri===self):null}))}:{}),
    contentState: state,
    contentType:
      kind && !/^(Text|RichText(?:\/.*)?)$/.test(kind) ? "system" : "text",
    ...(parent ? { threadId: parent, replyToId: parent } : {}),
  };
}
export function createMessaging(root) {
  const cursors = new MessagingCursors(root),
    operations = {};
  async function request(ctx, url, args = {}, method = "GET") {
    ctx.signal.throwIfAborted();
    const query = new URLSearchParams(
      Object.entries(args)
        .filter(([, v]) => v !== undefined)
        .map(([k, v]) => [k, String(v)]),
    ).toString();
    let response;
    try {
      response = await ctx.request({
        url: url + (method === "GET" && query ? "?" + query : ""),
        method,
        ...(method === "GET"
          ? {}
          : {
              headers: [["content-type", "application/json"]],
              body: JSON.stringify(args),
            }),
      });
    } catch (e) {
      if (
        e.reauthenticationRequired ||
        terminal.has(e.code) ||
        e.status === 401
      )
        throw fail("authentication_required", 401);
      throw e;
    }
    const data = await response.json();
    ctx.signal.throwIfAborted();
    if (!response.ok || data.error) {
      const e = data.error ?? {};
      if (
        e.reauthenticationRequired ||
        terminal.has(e.code) ||
        response.status === 401
      )
        throw fail("authentication_required", 401);
      if (
        ["send_uncertain", "mutation_uncertain"].includes(e.code) ||
        response.status >= 500
      )
        throw fail(
          method === "GET" ? "source_unavailable" : "uncertain",
          method === "GET" ? 502 : 409,
        );
      if (e.code === "idempotency_conflict")
        throw fail("idempotency_conflict", 409);
      if (e.code === "invalid_cursor") throw fail("invalid_cursor", 409);
      if (response.status === 403) throw fail("permission_denied", 403);
      if (response.status === 404) throw fail("not_found", 404);
      if (response.status === 429)
        throw Object.assign(fail("rate_limited", 429), {
          retryAfterMs:
            e.retryAfterMs ??
            Number(e.retryAfter || response.headers.get("retry-after") || 60) *
              1000,
        });
      throw fail("invalid_request");
    }
    return data;
  }
  const available = (c) =>
    (["browser", "tokens"].includes(c.methodId) ||
      (c.methodId === "authentication-source" && c.credentials?.authMode === "microsoft")) &&
    c.credentials?.profile &&
    c.credentials?.identity
      ? Object.keys(operations).filter(
          (n) => operations[n].readOnly || c.config.access === "write",
        )
      : [];
  function op(name, invoke) {
    operations[name] = {
      category: "messaging",
      readOnly: /\.(get|list)$/.test(name),
      ...messagingOperations[name],
      async invoke(input, ctx) {
        if (!available(ctx.connection).includes(name))
          throw fail("permission_denied", 403);
        return invoke(input, ctx);
      },
    };
  }
  op("messaging.conversations.list", async (input, ctx) => {
    const limit = input.limit ?? 50;
    const bind = [ctx.connection.id, ctx.connection.credentials.identity.tenant,
      ctx.connection.credentials.identity.oid, "conversations-v2", limit];
    const state = input.cursor ? JSON.parse(await cursors.decode(bind, input.cursor)) : null;
    // Keep the provider's page size stable: its native cursor is bound to it.
    const raw = state?.chatsDone ? {items:[]} : await request(ctx, "/chats/discovery", {
      limit:Math.ceil(limit/2), cursor:state?.chat,
    });
    let channels = {items:[]}, unavailable = false;
    try { channels = await request(ctx, "/channels"); }
    catch(e){if(!["permission_denied","not_found"].includes(e.code))throw e;unavailable=true;}
    const directory = [...new Map(channels.items.map(c=>[c.id,c])).values()]
      .sort((a,b)=>a.id<b.id?-1:a.id>b.id?1:0);
    const hash = createHash("sha256").update(JSON.stringify(directory.map(c=>c.id))).digest("hex");
    // Restart instead of silently skipping/repeating channels if membership changes.
    if(state && state.channelHash!==hash)throw fail("invalid_cursor",409);
    const items=[];
    for(const item of raw.items){
      try{
        const c=await request(ctx,"/contract/conversations/"+encodeURIComponent(item.id));
        items.push({id:c.id,title:c.title,kind:c.kind,unread:c.unreadCount,updatedAt:date(c.time)});
      }catch(e){if(!["permission_denied","not_found"].includes(e.code))throw e;unavailable=true;}
    }
    const offset=state?.channelOffset??0;
    const selected=directory.slice(offset,offset+Math.max(0,limit-items.length));
    items.push(...selected.map(c=>({id:c.id,title:c.title||"Teams channel",kind:"channel",unread:null,updatedAt:null})));
    const next={chat:raw.nextCursor??null,chatsDone:state?.chatsDone||!raw.nextCursor,
      channelOffset:offset+selected.length,channelHash:hash};
    return {items,...(!next.chatsDone||next.channelOffset<directory.length?
      {nextCursor:await cursors.encode(bind,JSON.stringify(next))}:{}),
      coverage:{complete:false,reason:unavailable?"available_snapshot_some_resources_unavailable":"available_chat_and_channel_snapshot_only"}};
  });
  op("messaging.conversations.get", async (input, ctx) => {
    decode(input.id);
    const raw = await request(
      ctx,
      "/contract/conversations/" + encodeURIComponent(input.id),
    );
    return {
      id: raw.id,
      title: raw.title,
      kind: raw.kind,
      unread: raw.unreadCount,
      updatedAt: date(raw.time),
    };
  });
  op("messaging.messages.list", async (input, ctx) => {
    const kind = decode(input.conversationId).kind,
      limit = input.limit ?? 50;
    if (input.threadId && kind !== "channel") throw fail("invalid_request");
    const route = input.threadId
      ? `/channels/${encodeURIComponent(input.conversationId)}/threads/${encodeURIComponent(input.threadId)}/history`
      : `/${kind === "chat" ? "chats" : "channels"}/${encodeURIComponent(input.conversationId)}/history`;
    const bind = [
      ctx.connection.id,
      ctx.connection.credentials.identity.tenant,
      ctx.connection.credentials.identity.oid,
      "messages",
      input.conversationId,
      input.threadId ?? null,
      limit,
    ];
    const cursor = input.cursor
        ? await cursors.decode(bind, input.cursor)
        : undefined,
      raw = await request(ctx, route, { limit, cursor });
    const self = "8:orgid:" + ctx.connection.credentials.identity.oid;
    return {
      items: raw.items.map((m) => project(m, input.conversationId, self)),
      ...(raw.nextCursor
        ? { nextCursor: await cursors.encode(bind, raw.nextCursor) }
        : {}),
      coverage: {
        complete: false,
        reason: input.threadId
          ? "available_thread_history_only"
          : "available_history_only",
      },
    };
  });
  op("messaging.messages.get", async (input, ctx) => {
    decode(input.conversationId);
    const t = target(input.id);
    const data = await request(
      ctx,
      `/contract/conversations/${encodeURIComponent(input.conversationId)}/messages/${encodeURIComponent(t.id)}`,
      { threadId: t.parent },
    );
    return project(data.raw, input.conversationId, data.account);
  });
  op("messaging.messages.send", async (input, ctx) => {
    const kind = decode(input.conversationId).kind,
      reply = input.replyToId ? target(input.replyToId) : null;
    const thread =
      input.threadId ??
      (kind === "channel" ? (reply?.parent ?? reply?.id) : null);
    if (
      (kind === "chat" && input.threadId) ||
      (kind === "channel" && !thread) ||
      (reply && input.threadId && (reply.parent ?? reply.id) !== input.threadId)
    )
      throw fail("invalid_request");
    if (typeof ctx.idempotencyKey !== "string" || !ctx.idempotencyKey)
      throw fail("invalid_idempotency_key");
    const idempotencyKey =
      "contract:" +
      createHash("sha256")
        .update(JSON.stringify([ctx.connection.id, ctx.idempotencyKey]))
        .digest("hex");
    const route =
      kind === "chat"
        ? `/chats/${encodeURIComponent(input.conversationId)}/messages`
        : `/channels/${encodeURIComponent(input.conversationId)}/threads/${encodeURIComponent(thread)}/messages`;
    try {
      const result = await request(
        ctx,
        route,
        {
          text: input.text,
          idempotencyKey,
          ...(kind === "chat" && reply ? { replyToMessageId: reply.id } : {}),
        },
        "POST",
      );
      if (result.status !== "accepted_by_api" || !result.message?.id)
        throw fail("uncertain", 409);
      return { status: "accepted", resourceId: mid(result.message.id, thread) };
    } catch (e) {
      if (
        [
          "authentication_required",
          "idempotency_conflict",
          "permission_denied",
          "rate_limited",
        ].includes(e.code) ||
        ctx.signal.aborted
      )
        throw e;
      return ["invalid_request", "not_found"].includes(e.code)
        ? { status: "rejected", code: e.code }
        : { status: "uncertain" };
    }
  });
  op("messaging.reactions.update", async (input, ctx) => {
    const kind=decode(input.conversationId).kind;
    if(kind!=="chat") return {status:"rejected",code:"unsupported"};
    const message=target(input.messageId);
    if(message.parent) return {status:"rejected",code:"unsupported"};
    const reaction=REACTIONS[input.emoji] ?? (/^:[a-z][a-z0-9_-]{0,61}:$/.test(input.emoji)?input.emoji.slice(1,-1):null);
    if(!reaction) return {status:"rejected",code:"invalid_request"};
    if(typeof ctx.idempotencyKey!=="string"||!ctx.idempotencyKey)throw fail("invalid_idempotency_key");
    const idempotencyKey="contract:"+createHash("sha256").update(JSON.stringify([ctx.connection.id,ctx.idempotencyKey])).digest("hex");
    try {
      const result=await request(ctx,`/chats/${encodeURIComponent(input.conversationId)}/messages/${encodeURIComponent(message.id)}/reactions`,{reaction,selected:!input.remove,idempotencyKey},"POST");
      return result.status==="accepted_by_api" || result.status==="reaction_observed" ? {status:"accepted"} : {status:"uncertain"};
    } catch(e) {
      if(["authentication_required","idempotency_conflict","permission_denied","rate_limited"].includes(e.code)||ctx.signal.aborted)throw e;
      return ["invalid_request","not_found"].includes(e.code)?{status:"rejected",code:e.code}:{status:"uncertain"};
    }
  });
  operations["messaging.messages.send"].inputSchema = structuredClone(
    operations["messaging.messages.send"].inputSchema,
  );
  operations["messaging.messages.send"].inputSchema.properties.text.maxLength =
    20000;
  return {
    version: 1,
    operations,
    available,
    async collect(_input, ctx) {
      await request(ctx, "/chats/discovery", { limit: 1 });
      const checkpoint = randomUUID();
      return {
        checkpoint,
        events: [
          {
            sourceId: "teams-reconcile:" + checkpoint,
            event: {
              category: "messaging",
              type: "invalidated",
              resourceType: "scope",
            },
          },
        ],
      };
    },
  };
}

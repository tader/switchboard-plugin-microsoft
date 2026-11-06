import { randomUUID } from "node:crypto";
import { setTimeout as sleep } from "node:timers/promises";
import { emailOperations } from "./email-schema.ts";
import { Cursors } from "./cursors.ts";
import type {
  Connection,
  ContractContext,
  ContractDefinition,
  EmailMessage,
  Logger,
} from "./types/api.d.ts";
const GRAPH = "https://graph.microsoft.com/v1.0",
  OUTLOOK = "https://outlook.office.com/api/v2.0";
const web = (c: Connection) =>
  c.methodId === "authentication-source" && c.config.backend === "outlook";
const root = (c: Connection) => (web(c) ? OUTLOOK : GRAPH);
const fail = (code: string, status = 400): never => {
  throw Object.assign(Error(code), { code, status });
};
const segment = (value: unknown) => {
  if (
    typeof value !== "string" ||
    !value ||
    [".", ".."].includes(value) ||
    /[\\]/.test(value)
  )
    return fail("invalid_request");
  return encodeURIComponent(value);
};
function normalize(value: any): any {
  if (Array.isArray(value)) return value.map(normalize);
  if (value && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value).map(([k, v]) => [
        k.startsWith("@") ? k : k[0].toLowerCase() + k.slice(1),
        normalize(v),
      ]),
    );
  return value;
}
function wire(value: any): any {
  if (Array.isArray(value)) return value.map(wire);
  if (value && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value).map(([k, v]) => [
        k.startsWith("@") ? k : k[0].toUpperCase() + k.slice(1),
        wire(v),
      ]),
    );
  return value;
}
const field = (c: Connection, name: string) =>
  web(c) ? name[0].toUpperCase() + name.slice(1) : name;
const query = (values: Record<string, unknown>) =>
  "?" +
  new URLSearchParams(
    Object.entries(values)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => [k, String(v)]),
  );
async function request(
  ctx: ContractContext,
  path: string,
  method = "GET",
  body?: unknown,
) {
  const retryAfter = (response: Response): number | null => {
    const header = response.headers.get("retry-after");
    if (header === null || !header.trim()) return null;
    const seconds = Number(header);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const date = Date.parse(header);
    return Number.isFinite(date) ? Math.max(0, date - Date.now()) : null;
  };
  // A 429 explicitly rejects the request. Other errors, especially transport
  // failures and 5xx responses to writes, can have an ambiguous outcome.
  const delays = [250, 500, 1000];
  let r: Response;
  for (let attempt = 0; ; attempt++) {
    ctx.signal.throwIfAborted();
    r = await ctx.request({
      url: root(ctx.connection) + path,
      method,
      headers: [["content-type", "application/json"]],
      body: body === undefined ? undefined : JSON.stringify(web(ctx.connection) ? wire(body) : body),
    });
    ctx.signal.throwIfAborted();
    if (r.status !== 429 || attempt >= delays.length) break;
    const requested = retryAfter(r) ?? 0;
    // Respect a longer provider delay rather than retrying prematurely.
    if (requested > 3000) break;
    await sleep(Math.max(delays[attempt], requested), undefined, { signal: ctx.signal });
  }
  if (!r.ok) {
    const code = r.status === 401 ? "authentication_required"
      : r.status === 403 ? "permission_denied"
      : r.status === 404 ? "not_found"
      : r.status === 409 || r.status === 412 ? "conflict"
      : r.status === 429 ? "rate_limited"
      : r.status >= 500 ? "source_unavailable" : "invalid_request";
    const delay = retryAfter(r);
    throw Object.assign(Error(code), {
      code,
      status: r.status,
      ...(r.status === 429
        ? { retryAfterMs: delay === null ? 60000 : Math.max(1000, Math.min(86400000, delay)) }
        : {}),
    });
  }
  return [202, 204].includes(r.status) ? {} : normalize(await r.json());
}
function grants(c: Connection) {
  const audience = web(c)
    ? "https://outlook.office.com/"
    : "https://graph.microsoft.com/";
  const raw = String(
    (web(c) ? c.credentials.resources?.outlook : c.credentials)?.scope ?? "",
  ).split(/\s+/);
  const s = new Set(
    raw
      .filter((x) => !x.startsWith("https://") || x.startsWith(audience))
      .map((x) => x.replace(audience, "")),
  );
  const read = [
    "Mail.Read",
    "Mail.Read.Shared",
    "Mail.ReadWrite",
    "Mail.ReadWrite.Shared",
    "Mail.ReadWrite.All",
  ].some((x) => s.has(x));
  return {
    read,
    write:
      c.config.access !== "read" &&
      c.config.access !== "send" &&
      ["Mail.ReadWrite", "Mail.ReadWrite.Shared", "Mail.ReadWrite.All"].some(
        (x) => s.has(x),
      ),
    send:
      c.config.access !== "read" &&
      (s.has("Mail.Send") || s.has("Mail.Send.Shared")),
  };
}
const addresses = (list: any) =>
  Array.isArray(list)
    ? list
        .filter((x) => typeof x?.emailAddress?.address === "string")
        .map((x) => ({
          address: x.emailAddress.address,
          ...(x.emailAddress.name ? { name: x.emailAddress.name } : {}),
        }))
    : [];
const recipients = (list: any[]) =>
  list.map((a) => ({
    emailAddress: { address: a.address, ...(a.name ? { name: a.name } : {}) },
  }));
const date = (s: unknown) => {
  const d = new Date(typeof s === "string" ? s : NaN);
  return Number.isFinite(d.valueOf()) ? d.toISOString() : null;
};
export function project(raw: any): EmailMessage {
  const r = normalize(raw),
    type = String(r.body?.contentType ?? "").toLowerCase(),
    body = typeof r.body?.content === "string" ? r.body.content : null;
  const available = body !== null && ["text", "html"].includes(type);
  return {
    id: r.id,
    ...(r["@odata.etag"]
      ? { version: r["@odata.etag"] }
      : r.changeKey
        ? { version: r.changeKey }
        : {}),
    subject: r.subject ?? "",
    from: addresses(r.from ? [r.from] : []),
    to: addresses(r.toRecipients),
    cc: addresses(r.ccRecipients),
    bcc: addresses(r.bccRecipients),
    replyTo: addresses(r.replyTo),
    collectionIds: r.parentFolderId ? [r.parentFolderId] : [],
    text: available && type === "text" ? body : null,
    ...(available && type === "html" ? { html: body } : {}),
    contentState: available ? "available" : "unavailable",
    isRead: typeof r.isRead === "boolean" ? r.isRead : null,
    isFlagged: r.flag?.flagStatus
      ? ["flagged"].includes(String(r.flag.flagStatus).toLowerCase())
      : null,
    draft: r.isDraft === true,
    receivedAt: date(r.receivedDateTime),
    sentAt: date(r.sentDateTime),
    ...(r.conversationId ? { threadId: r.conversationId } : {}),
    hasAttachments:
      typeof r.hasAttachments === "boolean" ? r.hasAttachments : null,
  };
}
function message(input: Record<string, any>) {
  const result: Record<string, unknown> = {};
  if (input.subject !== undefined) result.subject = input.subject;
  for (const [key, native] of [
    ["to", "toRecipients"],
    ["cc", "ccRecipients"],
    ["bcc", "bccRecipients"],
  ])
    if (input[key] !== undefined) result[native] = recipients(input[key]);
  if (input.html !== undefined || input.text !== undefined)
    result.body = {
      contentType: input.html !== undefined ? "HTML" : "Text",
      content: input.html ?? input.text,
    };
  return result;
}
export interface ContractOptions {
  log?: Pick<Logger, "warn">;
  /** Total time a folder listing waits for role resolution. */
  roleBudgetMs?: number;
  /** How long a complete role map stays cached per connection. */
  roleTtlMs?: number;
  /** How long a failed or incomplete role map stays cached. */
  roleFailureTtlMs?: number;
  now?: () => number;
}
export function createContracts(
  dataDir: string,
  options: ContractOptions = {},
): ContractDefinition {
  const cursors = new Cursors(dataDir),
    operations: ContractDefinition["operations"] = {};
  const available = (c: Connection) =>
    Object.keys(operations).filter((n) =>
      /\.(list|get)$/.test(n)
        ? grants(c).read
        : /\.(send|reply|forward)$/.test(n)
          ? grants(c).send && (n !== "email.drafts.send" || grants(c).read)
          : grants(c).write,
    );
  const op = (
    short: string,
    invoke: (
      input: Record<string, any>,
      ctx: ContractContext,
    ) => Promise<unknown>,
  ) => {
    const name = "email." + short,
      schema = structuredClone(emailOperations[name]);
    if (
      schema.inputSchema.properties &&
      typeof schema.inputSchema.properties === "object"
    )
      delete (schema.inputSchema.properties as any).ifVersion;
    operations[name] = {
      ...schema,
      category: "email",
      inputConstraints(connection) {
        if (!web(connection)) return {};
        if (short === "messages.update")
          return {
            properties: { changes: { properties: { isFlagged: false } } },
          };
        if (short === "messages.reply") return { properties: { html: false } };
        if (short === "messages.forward")
          return {
            properties: {
              html: false,
              cc: { maxItems: 0 },
              bcc: { maxItems: 0 },
            },
          };
        return {};
      },
      readOnly: /\.(list|get)$/.test(name),
      async invoke(input, ctx) {
        if (!available(ctx.connection).includes(name))
          return fail("permission_denied", 403);
        if (input.ifVersion !== undefined)
          return fail("unsupported_precondition");
        if (!/\.(list|get)$/.test(name) && !ctx.idempotencyKey)
          return fail("invalid_idempotency_key");
        try {
          return await invoke(input, ctx);
        } catch (e: any) {
          if (
            !/\.(list|get)$/.test(name) &&
            [
              "not_found",
              "conflict",
              "permission_denied",
              "invalid_request",
              "unsupported_option",
            ].includes(e.code)
          )
            return { status: "rejected", code: e.code };
          throw e;
        }
      },
    };
  };
  const bind = (
    ctx: ContractContext,
    kind: string,
    input: Record<string, any>,
  ) => [
    ctx.connection.id,
    ctx.connection.account?.id ?? null,
    root(ctx.connection),
    kind,
    input.limit ?? 50,
    input.collectionId ?? null,
    input.isRead ?? null,
  ];
  const continuation = (data: any, ctx: ContractContext, route: string) => {
    const link =
      data["@odata.nextLink"] ??
      data["@odata.nextlink"] ??
      data["odata.nextLink"];
    if (!link) return undefined;
    const u = new URL(link, root(ctx.connection));
    const base = new URL(root(ctx.connection));
    if (
      u.origin !== base.origin ||
      u.username ||
      u.password ||
      u.hash ||
      u.pathname.toLowerCase() !== (base.pathname + route).toLowerCase()
    )
      return fail("invalid_cursor", 409);
    return u.pathname.slice(base.pathname.length) + u.search;
  };
  const get = (ctx: ContractContext, id: unknown) =>
    request(ctx, "/me/messages/" + segment(id));
  const draft = async (ctx: ContractContext, id: unknown) => {
    const r = await get(ctx, id);
    if (r.isDraft !== true) return fail("not_found", 404);
    return r;
  };
  const page = async (
    items: unknown[],
    next: string | undefined,
    binding: unknown,
    reason?: string,
  ) => ({
    items,
    ...(next ? { nextCursor: await cursors.encode(binding, next) } : {}),
    coverage: {
      complete: !next && !reason,
      ...(next || reason ? { reason: reason ?? "more_pages" } : {}),
    },
  });
  // Graph/Outlook well-known folder names are locale independent; never infer roles from display names.
  const wellKnown: Record<string, string> = {
    inbox: "inbox",
    sentitems: "sent",
    drafts: "drafts",
    deleteditems: "trash",
    archive: "archive",
    junkemail: "spam",
  };
  // v1.0 folders carry no wellKnownName, so resolve the well-known folder IDs once and cache them.
  // Roles are best effort: they must never slow down or fail the folder listing itself.
  const roleBudgetMs = options.roleBudgetMs ?? 5000,
    roleTtlMs = options.roleTtlMs ?? 6 * 3600000,
    roleFailureTtlMs = options.roleFailureTtlMs ?? 120000,
    now = options.now ?? Date.now;
  type RoleEntry = {
    /** Map served while no lookup is running (or the previous one during a refresh). */
    roles: Record<string, string>;
    /** Roles found so far by the running lookup. */
    found: Record<string, string>;
    until: number;
    started: number;
    pending?: Promise<Record<string, string>>;
  };
  const roleCache = new Map<string, RoleEntry>();
  // Provider answers meaning "no such folder" (no role); anything else is transient and cached briefly.
  const absent = ["not_found", "permission_denied", "invalid_request"];
  const statusAbsent = (status: number) => [400, 403, 404].includes(status);
  const lookupRoles = async (
    ctx: ContractContext,
    found: Record<string, string>,
  ) => {
    const names = Object.keys(wellKnown);
    let settled = 0;
    const select = query({ $select: field(ctx.connection, "id") });
    if (!web(ctx.connection)) {
      // One Graph JSON batch instead of one request per folder (also stays under mailbox concurrency limits).
      const r = await request(ctx, "/$batch", "POST", {
        requests: names.map((name, i) => ({
          id: String(i),
          method: "GET",
          url: "/me/mailFolders/" + name + select,
        })),
      });
      for (const item of Array.isArray(r.responses) ? r.responses : []) {
        const name = names[Number(item?.id)];
        if (!name) continue;
        const id = item.body?.id;
        if (item.status === 200 && typeof id === "string" && id) {
          found[id] = wellKnown[name];
          settled++;
        } else if (statusAbsent(item.status)) settled++;
      }
      return settled === names.length;
    }
    // The Outlook REST compatibility backend has no JSON batch: bounded parallel reads.
    const queue = [...names];
    const worker = async () => {
      for (let name; (name = queue.shift()); ) {
        try {
          const r = await request(ctx, "/me/mailFolders/" + name + select);
          if (typeof r.id === "string" && r.id) found[r.id] = wellKnown[name];
          settled++;
        } catch (error: any) {
          // A mailbox without e.g. an archive folder has no such role.
          if (absent.includes(error.code)) settled++;
        }
      }
    };
    await Promise.all([worker(), worker(), worker()]);
    return settled === names.length;
  };
  const resolveRoles = (ctx: ContractContext) => {
    const key = ctx.connection.id + "\0" + root(ctx.connection);
    const hit = roleCache.get(key);
    // A lookup that never settles (hung provider) is abandoned after a minute so it can be retried.
    if (hit && (hit.pending ? now() - hit.started < 60000 : hit.until > now()))
      return hit;
    const stale = hit?.roles ?? {},
      found: Record<string, string> = {};
    const entry: RoleEntry = {
      roles: stale,
      found,
      until: 0,
      started: now(),
    };
    const settle = (roles: Record<string, string>, ttl: number) => {
      // Only the newest lookup may replace the cache entry.
      if (roleCache.get(key) === entry)
        roleCache.set(key, {
          roles,
          found: roles,
          until: now() + ttl,
          started: 0,
        });
    };
    entry.pending = lookupRoles(ctx, found).then(
      (complete) => {
        settle(found, complete ? roleTtlMs : roleFailureTtlMs);
        return found;
      },
      (error: any) => {
        options.log?.warn(
          "outlook-mail: folder role lookup failed",
          error?.code ?? error?.message ?? String(error),
        );
        const roles = { ...stale, ...found };
        settle(roles, roleFailureTtlMs);
        return roles;
      },
    );
    roleCache.set(key, entry);
    return entry;
  };
  // Waits at most until `deadline` for roles; afterwards the listing continues with what is known.
  // A refresh of an earlier map does not wait at all: the previous map is served meanwhile.
  const rolesBy = async (ctx: ContractContext, deadline: number) => {
    const entry = resolveRoles(ctx);
    if (!entry.pending || Object.keys(entry.roles).length) return entry.roles;
    let timer: Parameters<typeof clearTimeout>[0];
    const late = new Promise<undefined>((resolve) => {
      timer = setTimeout(resolve, Math.max(0, deadline - now()));
    });
    try {
      return (await Promise.race([entry.pending, late])) ?? { ...entry.found };
    } finally {
      clearTimeout(timer);
    }
  };
  op("mailboxes.list", async (input, ctx) => {
    const binding = bind(ctx, "folders", input),
      cursor = await cursors.decode(binding, input.cursor);
    const state: {
      route: string;
      queue: string[];
      seen: string[];
      bounded: boolean;
      partial?: string;
      roles?: unknown;
    } = cursor
      ? JSON.parse(cursor)
      : {
          route:
            "/me/mailFolders" +
            query({
              $top: input.limit ?? 50,
              includeHiddenFolders: web(ctx.connection) ? undefined : true,
            }),
          queue: [],
          seen: [],
          bounded: false,
          ...(web(ctx.connection) ? { partial: "visible_folders_only" } : {}),
        };
    // The cursor is signed; still bound work/memory when the provider has a huge or cyclic hierarchy.
    if (
      !state.route.startsWith("/me/mailFolders") ||
      state.seen.length > 1024 ||
      state.queue.length > 256
    )
      return fail("invalid_cursor", 409);
    // Cursors from earlier versions carried the role map; roles now live in the per-connection cache.
    delete state.roles;
    const deadline = now() + roleBudgetMs;
    // Start role resolution alongside the directory read; it never throws.
    const pendingRoles = rolesBy(ctx, deadline);
    const route = state.route.split("?")[0];
    let data: any;
    try {
      data = await request(ctx, state.route);
    } catch (error: any) {
      if (
        !route.toLowerCase().endsWith("/childfolders") ||
        !["permission_denied", "not_found"].includes(error.code)
      )
        throw error;
      data = { value: [] };
      state.partial = "some_folders_unavailable";
    }
    const roles = await pendingRoles;
    const items = (data.value ?? []).map((r: any) => ({
      id: r.id,
      name: r.displayName ?? r.id,
      kind: "folder",
      role:
        (Object.hasOwn(roles, r.id) ? roles[r.id] : undefined) ??
        (typeof r.wellKnownName === "string" &&
        Object.hasOwn(wellKnown, r.wellKnownName.toLowerCase())
          ? wellKnown[r.wellKnownName.toLowerCase()]
          : null),
      canWrite: grants(ctx.connection).write,
    }));
    for (const r of data.value ?? [])
      if (r.childFolderCount > 0 && !state.seen.includes(r.id)) {
        if (state.queue.length < 256 && state.seen.length < 1024) {
          state.queue.push(r.id);
          state.seen.push(r.id);
        } else state.bounded = true;
      }
    let next = continuation(data, ctx, route);
    if (!next && state.queue.length)
      next =
        "/me/mailFolders/" +
        segment(state.queue.shift()) +
        "/childFolders" +
        query({
          $top: input.limit ?? 50,
          includeHiddenFolders: web(ctx.connection) ? undefined : true,
        });
    state.route = next ?? "";
    const token = next ? JSON.stringify(state) : undefined;
    if (token && Buffer.byteLength(token) > 40000)
      return page(items, undefined, binding, "folder_directory_bounded");
    return page(
      items,
      token,
      binding,
      state.bounded ? "folder_directory_bounded" : state.partial,
    );
  });
  op("messages.get", async (input, ctx) => project(await get(ctx, input.id)));
  const list = async (
    input: Record<string, any>,
    ctx: ContractContext,
    drafts = false,
  ) => {
    const route = input.collectionId
        ? "/me/mailFolders/" + segment(input.collectionId) + "/messages"
        : "/me/messages",
      binding = bind(ctx, drafts ? "drafts" : "messages", input);
    const token = await cursors.decode(binding, input.cursor);
    const filter = [
      ...(drafts ? [field(ctx.connection, "isDraft") + " eq true"] : []),
      ...(input.isRead !== undefined
        ? [field(ctx.connection, "isRead") + " eq " + input.isRead]
        : []),
    ].join(" and ");
    const data = await request(
      ctx,
      token ??
        route +
          query({ $top: input.limit ?? 50, $filter: filter || undefined }),
    );
    const next = continuation(data, ctx, route),
      items = (data.value ?? []).map((m: any) =>
        drafts ? { id: m.id, message: project(m) } : project(m),
      );
    if (drafts && (data.value ?? []).some((m: any) => m.isDraft !== true))
      return fail("invalid_contract_response", 502);
    return page(items, next, binding);
  };
  op("messages.list", (input, ctx) => list(input, ctx));
  op("messages.update", async (input, ctx) => {
    const changes: Record<string, unknown> = {};
    if (input.changes.isRead !== undefined)
      changes.isRead = input.changes.isRead;
    if (input.changes.isFlagged !== undefined) {
      if (web(ctx.connection)) return fail("unsupported_option");
      changes.flag = {
        flagStatus: input.changes.isFlagged ? "flagged" : "notFlagged",
      };
    }
    const r = await request(
      ctx,
      "/me/messages/" + segment(input.id),
      "PATCH",
      changes,
    );
    return { status: "accepted", resourceId: r.id ?? input.id };
  });
  op("messages.move", async (input, ctx) => {
    if (input.sourceCollectionId !== undefined) {
      const current = await get(ctx, input.id);
      if (current.parentFolderId !== input.sourceCollectionId)
        return fail("conflict", 409);
    }
    const r = await request(
      ctx,
      "/me/messages/" + segment(input.id) + "/move",
      "POST",
      { destinationId: input.destinationId },
    );
    return { status: "accepted", resourceId: r.id };
  });
  op("messages.trash", async (input, ctx) => {
    const r = await request(
      ctx,
      "/me/messages/" + segment(input.id) + "/move",
      "POST",
      { destinationId: "deleteditems" },
    );
    return { status: "accepted", resourceId: r.id };
  });
  op("messages.send", async (input, ctx) => {
    await request(ctx, "/me/sendMail", "POST", {
      message: message(input),
      saveToSentItems: true,
    });
    return { status: "accepted" };
  });
  op("messages.reply", async (input, ctx) => {
    // Compatibility web actions only have a text Comment parameter.
    if (web(ctx.connection) && input.html !== undefined)
      return fail("unsupported_option");
    await request(
      ctx,
      "/me/messages/" +
        segment(input.id) +
        (input.replyAll ? "/replyAll" : "/reply"),
      "POST",
      web(ctx.connection)
        ? { comment: input.text }
        : { message: { body: message(input).body } },
    );
    return { status: "accepted" };
  });
  op("messages.forward", async (input, ctx) => {
    if (
      web(ctx.connection) &&
      (input.html !== undefined || input.cc?.length || input.bcc?.length)
    )
      return fail("unsupported_option");
    const body = web(ctx.connection)
      ? { comment: input.text, toRecipients: recipients(input.to) }
      : { message: message(input) };
    await request(
      ctx,
      "/me/messages/" + segment(input.id) + "/forward",
      "POST",
      body,
    );
    return { status: "accepted" };
  });
  op("drafts.list", (input, ctx) => list(input, ctx, true));
  op("drafts.get", async (input, ctx) => {
    const r = await draft(ctx, input.id);
    return { id: r.id, message: project(r) };
  });
  op("drafts.create", async (input, ctx) => {
    const r = await request(ctx, "/me/messages", "POST", message(input));
    return { status: "accepted", resourceId: r.id };
  });
  op("drafts.update", async (input, ctx) => {
    await draft(ctx, input.id);
    const r = await request(
      ctx,
      "/me/messages/" + segment(input.id),
      "PATCH",
      message(input.changes),
    );
    return { status: "accepted", resourceId: r.id ?? input.id };
  });
  op("drafts.delete", async (input, ctx) => {
    await draft(ctx, input.id);
    await request(ctx, "/me/messages/" + segment(input.id), "DELETE");
    return { status: "accepted", resourceId: input.id };
  });
  op("drafts.send", async (input, ctx) => {
    const r = await draft(ctx, input.id);
    if (
      ![
        ...addresses(r.toRecipients),
        ...addresses(r.ccRecipients),
        ...addresses(r.bccRecipients),
      ].length
    )
      return fail("invalid_request");
    await request(
      ctx,
      "/me/messages/" + segment(input.id) + "/send",
      "POST",
      {},
    );
    return { status: "accepted" };
  });
  return {
    version: 1,
    operations,
    available,
    async collect(_input, ctx) {
      await request(ctx, "/me/mailFolders?$top=1");
      const checkpoint = randomUUID();
      return {
        checkpoint,
        events: [
          {
            sourceId: "outlook-email-reconcile:" + checkpoint,
            event: {
              category: "email",
              type: "invalidated",
              resourceType: "scope",
            },
          },
        ],
      };
    },
  };
}

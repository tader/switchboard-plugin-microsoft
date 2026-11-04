import { randomUUID } from "node:crypto";
import zones from "./windows-zones.json" with { type: "json" };
import { calendarOperations } from "./calendar-schema.ts";
import type {
  CalendarEvent,
  Connection,
  ContractContext,
  ContractDefinition,
  ContractOperation,
  WriteReceipt,
} from "./types/api.d.ts";
const GRAPH = "https://graph.microsoft.com/v1.0",
  OUTLOOK = "https://outlook.office.com/api/v2.0";
function native(c: Connection) {
  return (
    c.methodId === "authentication-source" && c.config.backend === "outlook"
  );
}
function root(c: Connection) {
  return native(c) ? OUTLOOK : GRAPH;
}
function fail(code: string): never {
  throw Object.assign(Error(code), { code });
}
function segment(value: unknown) {
  if (
    typeof value !== "string" ||
    !value ||
    [".", ".."].includes(value) ||
    /[\\/]/.test(value)
  )
    return fail("invalid_request");
  return encodeURIComponent(value);
}
function granted(c: Connection) {
  const audience = native(c)
    ? "https://outlook.office.com/"
    : "https://graph.microsoft.com/";
  const scopes = String(
    (native(c) ? c.credentials.resources?.outlook : c.credentials)?.scope ?? "",
  ).split(/\s+/);
  return new Set(
    scopes
      .filter((s) => !s.startsWith("https://") || s.startsWith(audience))
      .map((s) => s.replace(audience, "")),
  );
}
function canWrite(c: Connection) {
  return (
    c.config.access !== "read" &&
    [
      "Calendars.ReadWrite",
      "Calendars.ReadWrite.Shared",
      "Calendars.ReadWrite.All",
    ].some((s) => granted(c).has(s))
  );
}
function normalize(data: any): any {
  if (Array.isArray(data)) return data.map(normalize);
  if (data && typeof data === "object")
    return Object.fromEntries(
      Object.entries(data).map(([k, v]) => [
        k[0] === "@" ? k : k[0].toLowerCase() + k.slice(1),
        normalize(v),
      ]),
    );
  return data;
}
function wire(data: any): any {
  if (Array.isArray(data)) return data.map(wire);
  if (data && typeof data === "object")
    return Object.fromEntries(
      Object.entries(data).map(([k, v]) => [
        k[0].toUpperCase() + k.slice(1),
        wire(v),
      ]),
    );
  return data;
}
async function request(
  ctx: ContractContext,
  path: string,
  method = "GET",
  body?: unknown,
  headers: [string, string][] = [],
) {
  const response = await ctx.request({
    method,
    url: root(ctx.connection) + path,
    headers: [
      ["content-type", "application/json"],
      ["Prefer", 'outlook.timezone="UTC", outlook.body-content-type="text"'],
      ...headers,
    ],
    body:
      body === undefined
        ? undefined
        : JSON.stringify(native(ctx.connection) ? wire(body) : body),
  });
  if (response.status === 401) return fail("authentication_required");
  if (response.status === 429 || response.status >= 500) {
    const retry = response.headers.get("retry-after");
    const seconds = Number(retry);
    throw Object.assign(Error("source_unavailable"), {
      code: "source_unavailable",
      retryAfterMs: retry
        ? Number.isFinite(seconds)
          ? seconds * 1000
          : Math.max(0, Date.parse(retry) - Date.now())
        : undefined,
    });
  }
  if (!response.ok) {
    const code =
      response.status === 412 || response.status === 409
        ? "conflict"
        : response.status === 404
          ? "not_found"
          : response.status === 403
            ? "permission_denied"
            : "invalid_request";
    throw Object.assign(Error(code), { code, status: response.status });
  }
  return response.status === 204 ? {} : normalize(await response.json());
}
function query(values: Record<string, unknown>) {
  return (
    "?" +
    new URLSearchParams(
      Object.entries(values)
        .filter(([, v]) => v !== undefined)
        .map(([k, v]) => [k, String(v)]),
    )
  );
}
function binding(
  ctx: ContractContext,
  kind: string,
  input: Record<string, unknown>,
) {
  return {
    connection: ctx.connection.id,
    backend: root(ctx.connection),
    kind,
    ...(kind === "events"
      ? {
          calendarId: input.calendarId,
          timeMin: input.timeMin,
          timeMax: input.timeMax,
        }
      : {}),
  };
}
function pageArgs(cursor: unknown, bound: unknown): Record<string, string> {
  if (!cursor) return {};
  try {
    const x = JSON.parse(Buffer.from(String(cursor), "base64url").toString());
    if (
      JSON.stringify(x.binding) !== JSON.stringify(bound) ||
      !x.page ||
      typeof x.page !== "object" ||
      Object.keys(x.page).some((k) => !["$skip", "$skiptoken"].includes(k)) ||
      Object.values(x.page).some((v) => typeof v !== "string")
    )
      return fail("invalid_cursor");
    return x.page;
  } catch {
    return fail("invalid_cursor");
  }
}
function nextCursor(
  data: any,
  ctx: ContractContext,
  path: string,
  bound: unknown,
) {
  const link =
    data["@odata.nextLink"] ??
    data["@odata.nextlink"] ??
    data["odata.nextLink"];
  if (!link) return undefined;
  const url = new URL(link, root(ctx.connection));
  const expected = new URL(root(ctx.connection) + path);
  const actualParts = decodeURIComponent(url.pathname).split("/");
  const expectedParts = decodeURIComponent(expected.pathname).split("/");
  const idIndex = path.toLowerCase().endsWith("/calendarview")
    ? expectedParts.length - 2
    : -1;
  const samePath =
    actualParts.length === expectedParts.length &&
    actualParts.every((part, index) =>
      index === idIndex
        ? part === expectedParts[index]
        : part.toLowerCase() === expectedParts[index].toLowerCase(),
    );
  if (
    url.origin !== expected.origin ||
    url.username ||
    url.password ||
    !samePath
  )
    return fail("invalid_provider_response");
  const page = Object.fromEntries(
    [...url.searchParams].filter(([k]) => ["$skip", "$skiptoken"].includes(k)),
  );
  if (!Object.keys(page).length) return fail("invalid_provider_response");
  return Buffer.from(JSON.stringify({ binding: bound, page })).toString(
    "base64url",
  );
}
function iana(value: string | undefined) {
  const mapped =
    (zones.windowsToIana as Record<string, string>)[value ?? ""] ?? value;
  try {
    if (!mapped) throw Error();
    new Intl.DateTimeFormat("en", { timeZone: mapped });
    return mapped;
  } catch {
    return fail("unsupported_time_zone");
  }
}
function instant(value: any) {
  if (!value?.dateTime) return fail("invalid_provider_response");
  const utc = ["UTC", "Etc/UTC", "GMT"].includes(value.timeZone);
  const text = /Z$|[+-]\d{2}:\d{2}$/.test(value.dateTime)
    ? value.dateTime
    : utc
      ? value.dateTime + "Z"
      : fail("unsupported_time_zone");
  const date = new Date(text);
  if (!Number.isFinite(date.valueOf()))
    return fail("invalid_provider_response");
  return date;
}
function civil(date: Date, zone: string) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: zone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  return ["year", "month", "day"]
    .map((k) => parts.find((p) => p.type === k)!.value)
    .join("-");
}
function mappedEvent(raw: any, calendarId: string): CalendarEvent {
  const startZone = iana(raw.originalStartTimeZone ?? raw.start?.timeZone),
    endZone = iana(
      raw.originalEndTimeZone ?? raw.end?.timeZone ?? raw.originalStartTimeZone,
    );
  const start = instant(raw.start),
    end = instant(raw.end);
  return {
    id: raw.id,
    calendarId,
    title: raw.subject ?? "",
    ...(raw["@odata.etag"] ? { version: raw["@odata.etag"] } : {}),
    start: raw.isAllDay
      ? { date: civil(start, startZone) }
      : { dateTime: start.toISOString(), timeZone: startZone },
    end: raw.isAllDay
      ? { date: civil(end, endZone) }
      : { dateTime: end.toISOString(), timeZone: endZone },
    ...(raw.body?.content ? { description: raw.body.content } : {}),
    ...(raw.location?.displayName
      ? { location: raw.location.displayName }
      : {}),
    ...(raw.seriesMasterId
      ? { seriesId: raw.seriesMasterId, occurrenceId: raw.id }
      : {}),
  };
}
function validTimes(event: any) {
  const a = event.start,
    b = event.end;
  if (!a || !b || Boolean(a.date) !== Boolean(b.date))
    return fail("invalid_request");
  for (const t of [a, b]) {
    const date = new Date(t.date ?? t.dateTime);
    if (
      !Number.isFinite(date.valueOf()) ||
      (t.date && date.toISOString().slice(0, 10) !== t.date)
    )
      return fail("invalid_request");
    if (!t.date) {
      const civilDate = new Date(t.dateTime.slice(0, 10));
      if (
        !Number.isFinite(civilDate.valueOf()) ||
        civilDate.toISOString().slice(0, 10) !== t.dateTime.slice(0, 10)
      )
        return fail("invalid_request");
      iana(t.timeZone);
    }
  }
  if (Date.parse(a.date ?? a.dateTime) >= Date.parse(b.date ?? b.dateTime))
    return fail("invalid_request");
}
function providerTime(time: any, c: Connection) {
  if (time.date) return { dateTime: time.date + "T00:00:00", timeZone: "UTC" };
  const zone = iana(time.timeZone),
    date = new Date(time.dateTime);
  const formatter = new Intl.DateTimeFormat("en-GB", {
    timeZone: zone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  });
  const wall = formatter.format(date);
  const fold = [-7200000, -3600000, -1800000, 1800000, 3600000, 7200000].some(
    (delta) => formatter.format(new Date(date.valueOf() + delta)) === wall,
  );
  if (fold)
    return { dateTime: date.toISOString().slice(0, -1), timeZone: "UTC" };
  const parts = formatter.formatToParts(date);
  const p = (kind: string) => parts.find((v) => v.type === kind)!.value;
  const providerZone = native(c)
    ? ((zones.ianaToWindows as Record<string, string>)[zone] ??
      (zone === "UTC" ? "UTC" : fail("unsupported_time_zone")))
    : zone;
  return {
    dateTime: `${p("year")}-${p("month")}-${p("day")}T${p("hour")}:${p("minute")}:${p("second")}`,
    timeZone: providerZone,
  };
}
function payload(value: any, c: Connection) {
  return {
    ...(value.title !== undefined ? { subject: value.title } : {}),
    ...(value.description !== undefined
      ? {
          body: {
            contentType: native(c) ? "Text" : "text",
            content: value.description,
          },
        }
      : {}),
    ...(value.location !== undefined
      ? { location: { displayName: value.location } }
      : {}),
    ...(value.start
      ? {
          start: providerTime(value.start, c),
          isAllDay: Boolean(value.start.date),
        }
      : {}),
    ...(value.end ? { end: providerTime(value.end, c) } : {}),
  };
}
function rejected(error: any): WriteReceipt {
  if (
    [
      "conflict",
      "not_found",
      "invalid_request",
      "permission_denied",
      "unsupported_time_zone",
    ].includes(error.code)
  )
    return { status: "rejected", code: error.code };
  throw error;
}
function op(
  name: string,
  invoke: ContractOperation["invoke"],
): ContractOperation {
  return {
    category: "calendar",
    readOnly: /\.(list|get)$/.test(name),
    ...calendarOperations[name],
    invoke,
  };
}
const operations: ContractDefinition["operations"] = {
  "calendar.calendars.list": op(
    "calendar.calendars.list",
    async (input, ctx) => {
      const path = "/me/calendars",
        bound = binding(ctx, "calendars", input);
      const data = await request(
        ctx,
        path +
          query({ $top: input.limit ?? 100, ...pageArgs(input.cursor, bound) }),
      );
      const cursor = nextCursor(data, ctx, path, bound);
      return {
        items: (data.value ?? []).map((c: any) => ({
          id: c.id,
          name: c.name ?? c.id,
          timeZone: null,
          canWrite: canWrite(ctx.connection)
            ? typeof c.canEdit === "boolean"
              ? c.canEdit
              : null
            : false,
        })),
        ...(cursor ? { nextCursor: cursor } : {}),
        coverage: {
          complete: !cursor,
          ...(cursor ? { reason: "more_pages" } : {}),
        },
      };
    },
  ),
  "calendar.events.list": op("calendar.events.list", async (input, ctx) => {
    if (
      !Number.isFinite(Date.parse(String(input.timeMin))) ||
      !Number.isFinite(Date.parse(String(input.timeMax))) ||
      Date.parse(String(input.timeMin)) >= Date.parse(String(input.timeMax))
    )
      return fail("invalid_request");
    const path = "/me/calendars/" + segment(input.calendarId) + "/calendarView",
      bound = binding(ctx, "events", input);
    const data = await request(
      ctx,
      path +
        query({
          startDateTime: input.timeMin,
          endDateTime: input.timeMax,
          $top: input.limit ?? 100,
          ...pageArgs(input.cursor, bound),
        }),
    );
    const cursor = nextCursor(data, ctx, path, bound);
    return {
      items: (data.value ?? [])
        .filter((e: any) => !e.isCancelled)
        .map((e: any) => mappedEvent(e, String(input.calendarId))),
      ...(cursor ? { nextCursor: cursor } : {}),
      coverage: {
        complete: !cursor,
        ...(cursor ? { reason: "more_pages" } : {}),
      },
    };
  }),
  "calendar.events.get": op("calendar.events.get", async (input, ctx) => {
    const data = await request(
      ctx,
      "/me/calendars/" +
        segment(input.calendarId) +
        "/events/" +
        segment(input.id),
    );
    if (data.isCancelled) return fail("not_found");
    return mappedEvent(data, String(input.calendarId));
  }),
  "calendar.events.create": op("calendar.events.create", async (input, ctx) => {
    try {
      validTimes(input.event);
      const data = await request(
        ctx,
        "/me/calendars/" + segment(input.calendarId) + "/events",
        "POST",
        payload(input.event, ctx.connection),
      );
      return {
        status: "accepted",
        resourceId: data.id,
        ...(data["@odata.etag"] ? { version: data["@odata.etag"] } : {}),
      };
    } catch (e) {
      return rejected(e);
    }
  }),
  "calendar.events.update": op("calendar.events.update", async (input, ctx) => {
    try {
      const path =
        "/me/calendars/" +
        segment(input.calendarId) +
        "/events/" +
        segment(input.id);
      const data = await request(ctx, path);
      if (data["@odata.etag"] !== input.version)
        return { status: "rejected", code: "conflict" };
      const changes = input.changes as any;
      if (data.isOnlineMeeting && changes.description !== undefined)
        return {
          status: "rejected",
          code: "online_meeting_body_edit_unsupported",
        };
      const merged = {
        ...mappedEvent(data, String(input.calendarId)),
        ...changes,
      };
      validTimes(merged);
      // Sending both endpoints avoids leaving a mixed all-day/timed representation.
      const patch = {
        ...changes,
        ...(changes.start || changes.end
          ? { start: merged.start, end: merged.end }
          : {}),
      };
      const updated = await request(
        ctx,
        path,
        "PATCH",
        payload(patch, ctx.connection),
        [["if-match", String(input.version)]],
      );
      return {
        status: "accepted",
        resourceId: updated.id ?? String(input.id),
        ...(updated["@odata.etag"] ? { version: updated["@odata.etag"] } : {}),
      };
    } catch (e) {
      return rejected(e);
    }
  }),
  "calendar.events.delete": op("calendar.events.delete", async (input, ctx) => {
    try {
      await request(
        ctx,
        "/me/calendars/" +
          segment(input.calendarId) +
          "/events/" +
          segment(input.id),
        "DELETE",
        undefined,
        [["if-match", String(input.version)]],
      );
      return { status: "accepted", resourceId: String(input.id) };
    } catch (e) {
      return rejected(e);
    }
  }),
};
export const contracts: ContractDefinition = {
  version: 1,
  operations,
  available(c) {
    const scopes = granted(c);
    if (
      ![
        "Calendars.Read",
        "Calendars.Read.Shared",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
        "Calendars.ReadWrite.All",
      ].some((s) => scopes.has(s))
    )
      return [];
    return Object.keys(operations).filter(
      (name) => operations[name].readOnly || canWrite(c),
    );
  },
  async collect(_input, ctx) {
    await request(ctx, "/me/calendars?$top=1");
    const id = randomUUID();
    return {
      events: [
        {
          sourceId: id,
          event: {
            category: "calendar",
            type: "invalidated",
            resourceType: "calendars",
          },
        },
      ],
      checkpoint: id,
    };
  },
};

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { Ajv } from "ajv";
import { contracts } from "../plugins/outlook-calendar/contracts.ts";
const ajv = new Ajv({ strict: false });
const json = (value, status = 200, headers = {}) =>
  new Response(status === 204 ? null : JSON.stringify(value), {
    status,
    headers,
  });
function fixture(respond, web = false) {
  const requests = [];
  const ctx = {
    connection: {
      id: "c",
      methodId: web ? "authentication-source" : "oauth",
      config: { access: "full", backend: web ? "outlook" : "graph" },
      credentials: web
        ? {
            scope: "User.Read",
            resources: {
              outlook: {
                scope: "https://outlook.office.com/Calendars.ReadWrite",
              },
            },
          }
        : { scope: "Calendars.ReadWrite" },
    },
    signal: new AbortController().signal,
    async request(r) {
      requests.push(r);
      return respond(r);
    },
  };
  return { ctx, requests };
}
async function call(name, input, ctx) {
  const op = contracts.operations[name];
  assert.ok(ajv.validate(op.inputSchema, input), ajv.errorsText());
  const output = await op.invoke(input, ctx);
  assert.ok(ajv.validate(op.outputSchema, output), ajv.errorsText());
  return output;
}
const timed = {
  id: "event",
  "@odata.etag": '"v1"',
  subject: "Planning",
  start: { dateTime: "2026-11-03T08:00:00.0000000", timeZone: "UTC" },
  end: { dateTime: "2026-11-03T09:00:00.0000000", timeZone: "UTC" },
  originalStartTimeZone: "W. Europe Standard Time",
  originalEndTimeZone: "W. Europe Standard Time",
};
function pascal(x) {
  if (Array.isArray(x)) return x.map(pascal);
  if (x && typeof x === "object")
    return Object.fromEntries(
      Object.entries(x).map(([k, v]) => [
        k.startsWith("@") ? k : k[0].toUpperCase() + k.slice(1),
        pascal(v),
      ]),
    );
  return x;
}
test("effective capabilities use the selected audience, grants and binding limits", () => {
  const f = fixture(() => json({}));
  assert.equal(contracts.available(f.ctx.connection).length, 6);
  f.ctx.connection.config.access = "read";
  assert.equal(contracts.available(f.ctx.connection).length, 3);
  f.ctx.connection.config.access = "full";
  f.ctx.connection.credentials.scope = "Calendars.Read";
  assert.equal(contracts.available(f.ctx.connection).length, 3);
  f.ctx.connection.credentials.scope = "Calendars.ReadBasic";
  assert.deepEqual(contracts.available(f.ctx.connection), []);
  const g = fixture(() => json({}), true);
  assert.equal(contracts.available(g.ctx.connection).length, 6);
  g.ctx.connection.credentials.resources.outlook.scope =
    "https://graph.microsoft.com/Calendars.ReadWrite";
  assert.deepEqual(contracts.available(g.ctx.connection), []);
});
for (const web of [false, true]) {
  const label = web ? "Outlook web" : "Graph";
  test(
    label +
      ": lists calendars without assuming mailbox timezone or write rights",
    async () => {
      const f = fixture(
        () =>
          json(
            web
              ? { value: [{ Id: "cal", Name: "Work", CanEdit: false }] }
              : { value: [{ id: "cal", name: "Work", canEdit: false }] },
          ),
        web,
      );
      assert.deepEqual(await call("calendar.calendars.list", {}, f.ctx), {
        items: [{ id: "cal", name: "Work", timeZone: null, canWrite: false }],
        coverage: { complete: true },
      });
      assert.equal(
        new URL(f.requests[0].url).origin,
        web ? "https://outlook.office.com" : "https://graph.microsoft.com",
      );
    },
  );
  test(
    label +
      ": bounded snapshots normalize times, all-day dates and recurring instances",
    async () => {
      const allDay = {
        ...timed,
        id: "day",
        isAllDay: true,
        start: { dateTime: "2026-11-02T23:00:00", timeZone: "UTC" },
        end: { dateTime: "2026-11-03T23:00:00", timeZone: "UTC" },
      };
      const values = [
        timed,
        allDay,
        { ...timed, id: "occurrence", seriesMasterId: "master" },
        { id: "cancelled", isCancelled: true },
      ];
      const f = fixture(
        () => json({ value: web ? pascal(values) : values }),
        web,
      );
      const result = await call(
        "calendar.events.list",
        {
          calendarId: "cal",
          timeMin: "2026-11-01T00:00:00Z",
          timeMax: "2026-12-01T00:00:00Z",
        },
        f.ctx,
      );
      assert.equal(result.items.length, 3);
      assert.equal(result.items[0].start.dateTime, "2026-11-03T08:00:00.000Z");
      assert.equal(result.items[0].start.timeZone, "Europe/Berlin");
      assert.deepEqual(result.items[1].start, { date: "2026-11-03" });
      assert.deepEqual(result.items[1].end, { date: "2026-11-04" });
      assert.equal(result.items[2].seriesId, "master");
      assert.equal(
        new URL(f.requests[0].url).searchParams.get("startDateTime"),
        "2026-11-01T00:00:00Z",
      );
    },
  );
  test(
    label + ": creates map shared fields to provider names without attendees",
    async () => {
      const f = fixture(() => json(web ? pascal(timed) : timed), web);
      const input = {
        calendarId: "cal",
        event: {
          title: "Planning",
          start: {
            dateTime: "2026-11-03T09:00:00+01:00",
            timeZone: "Europe/Amsterdam",
          },
          end: {
            dateTime: "2026-11-03T10:00:00+01:00",
            timeZone: "Europe/Amsterdam",
          },
          description: "Notes",
        },
      };
      const receipt = await call("calendar.events.create", input, f.ctx);
      assert.equal(receipt.status, "accepted");
      assert.equal(receipt.version, '"v1"');
      const body = JSON.parse(f.requests[0].body);
      assert.equal(body[web ? "Subject" : "subject"], "Planning");
      assert.equal(
        body[web ? "Start" : "start"][web ? "DateTime" : "dateTime"],
        "2026-11-03T09:00:00",
      );
      assert.equal(
        body[web ? "Start" : "start"][web ? "TimeZone" : "timeZone"],
        web ? "W. Europe Standard Time" : "Europe/Amsterdam",
      );
      assert.ok(!("attendees" in body) && !("Attendees" in body));
    },
  );
  test(
    label + ": stale updates reject; actual mutations carry If-Match",
    async () => {
      const f = fixture((r) => json(web ? pascal(timed) : timed), web);
      const input = {
        calendarId: "cal",
        id: "event",
        version: "stale",
        changes: { title: "Updated" },
      };
      assert.deepEqual(await call("calendar.events.update", input, f.ctx), {
        status: "rejected",
        code: "conflict",
      });
      assert.equal(f.requests.length, 1);
      await call(
        "calendar.events.update",
        { ...input, version: '"v1"' },
        f.ctx,
      );
      assert.equal(f.requests.at(-1).method, "PATCH");
      assert.ok(
        f.requests
          .at(-1)
          .headers.some(([k, v]) => k === "if-match" && v === '"v1"'),
      );
      assert.deepEqual(
        JSON.parse(f.requests.at(-1).body),
        web ? { Subject: "Updated" } : { subject: "Updated" },
      );
    },
  );
  test(
    label + ": delete precondition failure is rejected, never silently retried",
    async () => {
      const f = fixture(() => json({}, 412), web);
      assert.deepEqual(
        await call(
          "calendar.events.delete",
          { calendarId: "cal", id: "event", version: '"v1"' },
          f.ctx,
        ),
        { status: "rejected", code: "conflict" },
      );
      assert.equal(f.requests.length, 1);
    },
  );
}
test("all-day ranges crossing DST keep civil dates instead of UTC dates", async () => {
  const f = fixture(() =>
    json({
      value: [
        {
          ...timed,
          isAllDay: true,
          start: { dateTime: "2026-03-28T23:00:00", timeZone: "UTC" },
          end: { dateTime: "2026-03-29T22:00:00", timeZone: "UTC" },
        },
      ],
    }),
  );
  const result = await call(
    "calendar.events.list",
    {
      calendarId: "cal",
      timeMin: "2026-03-01T00:00:00Z",
      timeMax: "2026-04-01T00:00:00Z",
    },
    f.ctx,
  );
  assert.deepEqual(result.items[0].start, { date: "2026-03-29" });
  assert.deepEqual(result.items[0].end, { date: "2026-03-30" });
});
test("pagination reconstructs the original query and rejects off-origin or wrong-resource links", async () => {
  const path = "https://graph.microsoft.com/v1.0/me/calendars/cal/calendarview";
  let link = path + "?$skiptoken=second&startDateTime=changed";
  const f = fixture(() => json({ value: [], "@odata.nextLink": link }));
  const args = {
    calendarId: "cal",
    timeMin: "2026-11-01T00:00:00Z",
    timeMax: "2026-12-01T00:00:00Z",
  };
  const first = await call("calendar.events.list", args, f.ctx);
  assert.equal(first.coverage.complete, false);
  link = undefined;
  await call(
    "calendar.events.list",
    { ...args, cursor: first.nextCursor },
    f.ctx,
  );
  const url = new URL(f.requests[1].url);
  assert.equal(url.searchParams.get("$skiptoken"), "second");
  assert.equal(url.searchParams.get("startDateTime"), args.timeMin);
  await assert.rejects(
    () =>
      call(
        "calendar.events.list",
        { ...args, calendarId: "other", cursor: first.nextCursor },
        f.ctx,
      ),
    { code: "invalid_cursor" },
  );
  for (const bad of [
    "https://evil.example/v1.0/me/calendars/cal/calendarview?$skip=2",
    "https://graph.microsoft.com/v1.0/me/calendars/other/calendarview?$skip=2",
  ]) {
    link = bad;
    await assert.rejects(() => call("calendar.events.list", args, f.ctx), {
      code: "invalid_provider_response",
    });
  }
});
test("writes reject invalid date ranges and protect online meeting body content", async () => {
  const f = fixture(() => json({ ...timed, isOnlineMeeting: true }));
  const input = {
    calendarId: "cal",
    event: {
      title: "Day",
      start: { date: "2026-02-30" },
      end: { date: "2026-03-03" },
    },
  };
  assert.deepEqual(await call("calendar.events.create", input, f.ctx), {
    status: "rejected",
    code: "invalid_request",
  });
  assert.equal(f.requests.length, 0);
  assert.deepEqual(
    await call(
      "calendar.events.update",
      {
        calendarId: "cal",
        id: "event",
        version: '"v1"',
        changes: { description: "replacement" },
      },
      f.ctx,
    ),
    { status: "rejected", code: "online_meeting_body_edit_unsupported" },
  );
  assert.equal(f.requests.length, 1);
});
test("collector checks provider auth, returns invalidation only and honors retry delay", async () => {
  const f = fixture(() => json({ value: [] }));
  assert.deepEqual((await contracts.collect({}, f.ctx)).events[0].event, {
    category: "calendar",
    type: "invalidated",
    resourceType: "calendars",
  });
  assert.equal(f.requests.length, 1);
  const g = fixture(() => json({}, 401));
  await assert.rejects(() => contracts.collect({}, g.ctx), {
    code: "authentication_required",
  });
  const h = fixture(() => json({}, 429, { "retry-after": "120" }));
  await assert.rejects(() => contracts.collect({}, h.ctx), {
    retryAfterMs: 120000,
  });
});
test("schema snapshot covers exactly the advertised operation set", () => {
  // Each provider executes the same normative schemas; the core independently enforces them.
  const local = JSON.parse(
    fs.readFileSync(
      new URL("../plugins/outlook-calendar/calendar-v1.json", import.meta.url),
    ),
  );
  assert.equal(local.version, 1);
  assert.deepEqual(
    Object.keys(local.operations),
    Object.keys(contracts.operations),
  );
});

test("a DST fold write preserves the exact instant instead of choosing the wrong offset", async () => {
  const f = fixture(() => json(timed));
  await call(
    "calendar.events.create",
    {
      calendarId: "cal",
      event: {
        title: "Fold",
        start: {
          dateTime: "2026-10-25T02:30:00+01:00",
          timeZone: "Europe/Amsterdam",
        },
        end: {
          dateTime: "2026-10-25T03:30:00+01:00",
          timeZone: "Europe/Amsterdam",
        },
      },
    },
    f.ctx,
  );
  const body = JSON.parse(f.requests[0].body);
  assert.deepEqual(body.start, {
    dateTime: "2026-10-25T01:30:00.000",
    timeZone: "UTC",
  });
});

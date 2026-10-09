// Builds slim OpenAPI descriptions of Microsoft Graph for the console's API reference.
//
// Microsoft's full description (github.com/microsoftgraph/msgraph-metadata) is 44 MB of YAML with
// 11,000+ operations; parsing it takes seconds and over a gigabyte of memory. This keeps only the
// operations each service needs, with parameters and example bodies inlined, so the plugin can ship
// small JSON files. Run it again to pick up changes to Graph:
//
//   node plugins/microsoft/openapi/build.ts [path-or-url-of-openapi.yaml]
import fs from 'node:fs';
import path from 'node:path';
import YAML from 'yaml';
import { exampleJson, makeResolver } from './schema.ts';

const SOURCE = 'https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/v1.0/openapi.yaml';

/** Path patterns per output file. */
export const SUBSETS: Record<string, { title: string; paths: RegExp[] }> = {
  mail: { title: 'Outlook Mail', paths: [/^\/me\/(messages|mailFolders|sendMail|inferenceClassification|outlook\/masterCategories|mailboxSettings)(\/|$)/] },
  calendar: { title: 'Outlook Calendar', paths: [/^\/me\/(calendar|calendars|calendarGroups|calendarView|events|findMeetingTimes)(\/|$)/] },
  onedrive: { title: 'OneDrive', paths: [/^\/me\/drive(\/|$)/, /^\/me\/drive$/, /^\/drives\/\{drive-id\}\/items\/\{driveItem-id\}(\/children|\/content)?$/, /^\/shares(\/|$)/] },
  todo: { title: 'Microsoft To Do', paths: [/^\/me\/todo(\/|$)/] },
};
// The general Microsoft Graph service gets all of the above, plus profile, contacts and people.
SUBSETS.graph = {
  title: 'Microsoft Graph',
  paths: [/^\/me$/, /^\/me\/(contacts|contactFolders|people|photo|memberOf|manager|directReports)(\/|$)/, ...Object.values(SUBSETS).flatMap((s) => s.paths)],
};

const METHODS = ['get', 'put', 'post', 'patch', 'delete'];

const to = [{ emailAddress: { address: 'someone@example.com' } }];
const at = (h: number) => ({ dateTime: `2026-01-15T${String(h).padStart(2, '0')}:00:00`, timeZone: 'UTC' });

/** Hand-written examples for common operations; generated ones list every property of the schema. */
const EXAMPLES: Record<string, unknown> = {
  'POST /me/sendMail': { message: { subject: 'Hello', body: { contentType: 'Text', content: 'Sent through Switchboard' }, toRecipients: to }, saveToSentItems: true },
  'POST /me/messages': { subject: 'Draft', body: { contentType: 'Text', content: 'Hello' }, toRecipients: to },
  'PATCH /me/messages/{message-id}': { isRead: true },
  'POST /me/messages/{message-id}/reply': { comment: 'Thanks!' },
  'POST /me/messages/{message-id}/forward': { comment: 'FYI', toRecipients: to },
  'POST /me/messages/{message-id}/move': { destinationId: 'archive' },
  'POST /me/events': { subject: 'Planning', start: at(10), end: at(11), attendees: [{ ...to[0], type: 'required' }], isOnlineMeeting: false },
  'PATCH /me/events/{event-id}': { subject: 'Planning (moved)', start: at(14), end: at(15) },
  'POST /me/calendar/getSchedule': { schedules: ['someone@example.com'], startTime: at(8), endTime: at(18), availabilityViewInterval: 30 },
  'POST /me/findMeetingTimes': { attendees: [{ ...to[0], type: 'required' }], meetingDuration: 'PT1H', maxCandidates: 5 },
  'POST /me/todo/lists': { displayName: 'Groceries' },
  'POST /me/todo/lists/{todoTaskList-id}/tasks': { title: 'Buy milk', dueDateTime: at(0) },
  'PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}': { status: 'completed' },
  'POST /me/drive/root/children': { name: 'New folder', folder: {}, '@microsoft.graph.conflictBehavior': 'rename' },
  'POST /me/drive/items/{driveItem-id}/children': { name: 'New folder', folder: {}, '@microsoft.graph.conflictBehavior': 'rename' },
};

/** Properties the service sets itself; Graph's description does not mark them read-only. */
const SERVER_SET = new Set(['id', '@odata.type', '@odata.etag', 'changeKey', 'createdDateTime', 'lastModifiedDateTime', 'webLink', 'iCalUId', 'parentFolderId', 'conversationId', 'conversationIndex', 'internetMessageId']);
function clean(v: any): any {
  if (Array.isArray(v)) return v.map(clean);
  if (v && typeof v === 'object') return Object.fromEntries(Object.entries(v).filter(([k]) => !SERVER_SET.has(k)).map(([k, x]) => [k, clean(x)]));
  return v;
}
const trim = (s: unknown, max: number) => (typeof s === 'string' && s.trim() ? (s.length > max ? `${s.slice(0, max)}…` : s.trim()) : undefined);

/**
 * Graph's description has drive items only under /drives/{drive-id}/…, but most use the
 * /me/drive/… aliases (as Microsoft's own docs do). Add those, and search, which is missing.
 */
function addDriveAliases(doc: any) {
  const ITEM = '/drives/{drive-id}/items/{driveItem-id}';
  const suffixes = ['', '/children', '/content', '/copy', '/createLink', '/createUploadSession', '/invite', '/permissions', '/preview', '/restore', '/thumbnails', '/versions', '/delta()', '/checkin', '/checkout'];
  const rootSuffixes = ['', '/children', '/content', '/delta()', '/createUploadSession'];
  for (const sfx of suffixes) {
    const item = doc.paths[ITEM + sfx];
    if (!item) continue;
    doc.paths[`/me/drive/items/{driveItem-id}${sfx}`] = item;
    if (rootSuffixes.includes(sfx)) doc.paths[`/me/drive/root${sfx}`] = item;
  }
  for (const fn of ['recent()', 'sharedWithMe()']) if (doc.paths[`/drives/{drive-id}/${fn}`]) doc.paths[`/me/drive/${fn}`] = doc.paths[`/drives/{drive-id}/${fn}`];
  doc.paths["/me/drive/root/search(q='{q}')"] = {
    get: {
      operationId: 'drive.root.search',
      summary: 'Search for files',
      description: 'Searches file names, metadata and content in your OneDrive.',
      tags: ['me.drive'],
      parameters: [{ name: 'q', in: 'path', required: true, description: 'The search text', schema: { type: 'string' } }, { $ref: '#/components/parameters/top' }, { $ref: '#/components/parameters/select' }],
    },
  };
}

async function load(src: string) {
  const text = /^https?:/.test(src) ? await (await fetch(src)).text() : fs.readFileSync(src, 'utf8');
  console.log(`parsing ${Math.round(text.length / 1e6)} MB…`);
  return YAML.parse(text, { maxAliasCount: -1 });
}

function slim(doc: any, patterns: RegExp[]) {
  const resolve = makeResolver(doc);
  const paths: Record<string, any> = {};
  for (const [p, rawItem] of Object.entries<any>(doc.paths)) {
    if (!patterns.some((re) => re.test(p))) continue;
    const item = resolve(rawItem) ?? {};
    const shared = (item.parameters ?? []).map(resolve);
    for (const m of METHODS) {
      const op = item[m];
      if (!op) continue;
      const params = [...shared, ...(op.parameters ?? []).map(resolve)]
        .filter((x: any) => x?.name && ['path', 'query', 'header'].includes(x.in))
        // Aliases drop some path parameters, e.g. {drive-id} for /me/drive/….
        .filter((x: any) => x.in !== 'path' || p.includes(`{${x.name}}`))
        .map((x: any) => {
          const schema = resolve(x.schema) ?? {};
          return {
            name: x.name,
            in: x.in,
            required: x.in === 'path' || !!x.required,
            description: trim(x.description, 300),
            schema: { type: schema.type ?? resolve(schema.items)?.type, enum: Array.isArray(schema.enum) ? schema.enum.slice(0, 30) : undefined, default: schema.default },
          };
        });
      let requestBody;
      const rb = resolve(op.requestBody);
      const json = rb?.content?.['application/json'];
      if (json) {
        const curated = EXAMPLES[`${m.toUpperCase()} ${p}`];
        const generated = curated ? undefined : exampleJson(json.schema, resolve);
        const example = curated ?? (generated ? clean(JSON.parse(generated)) : undefined);
        requestBody = { required: !!rb.required, content: { 'application/json': example !== undefined ? { example } : {} } };
      } else if (rb?.content) {
        requestBody = { required: !!rb.required, content: Object.fromEntries(Object.keys(rb.content).map((t) => [t, {}])) };
      }
      (paths[p] ??= {})[m] = {
        operationId: op.operationId,
        summary: trim(op.summary, 200),
        description: trim(op.description, 600),
        tags: op.tags?.slice(0, 1),
        deprecated: op.deprecated || undefined,
        parameters: params,
        requestBody,
      };
    }
  }
  return paths;
}

const doc = await load(process.argv[2] ?? SOURCE);
addDriveAliases(doc);
const out = path.dirname(new URL(import.meta.url).pathname);
for (const [name, subset] of Object.entries(SUBSETS)) {
  const paths = slim(doc, subset.paths);
  const file = path.join(out, `${name}.json`);
  const spec = { openapi: '3.0.3', info: { title: `Microsoft Graph: ${subset.title}`, version: 'v1.0', 'x-generated-from': SOURCE }, servers: [{ url: 'https://graph.microsoft.com/v1.0' }], paths };
  fs.writeFileSync(file, JSON.stringify(spec));
  const ops = Object.values(paths).reduce((n: number, i: any) => n + Object.keys(i).length, 0);
  console.log(`${name}.json: ${ops} operations, ${Math.round(fs.statSync(file).size / 1024)} KB`);
}

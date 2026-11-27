// Run against core 0.17.0; provider traffic is synthetic, loopback HTTP is real.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { graphServices, credentials, config as connectionConfig, drive, calendarBody, contentContext } from '../test/fixtures/output-preview.js';
import { createContentPlugin } from '../plugins/microsoft/content/service.js';
import { createAdapter } from '../plugins/teams/index.js';
const root = process.env.SWITCHBOARD_ROOT;
assert(root, 'Set SWITCHBOARD_ROOT to a core 0.17.0 checkout with npm ci completed');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'microsoft-output-'));
process.env.SWITCHBOARD_DATA_DIR = path.join(temp, 'data');
process.env.SWITCHBOARD_ACTIVITY_PREVIEW_ROOT = path.join(temp, 'outputs');
process.env.SWITCHBOARD_ACTIVITY_PREVIEW_RESTORE_SAFE = '1';
process.env.SWITCHBOARD_ACTIVITY_PREVIEW_INSTANCE = 'synthetic-microsoft';
process.env.SWITCHBOARD_ACTIVITY_PREVIEW_EPOCH = 'synthetic-epoch';
fs.mkdirSync(process.env.SWITCHBOARD_DATA_DIR);
const load = (file: string) => import(pathToFileURL(path.join(root, file)).href);
const database = await load('server/db.ts'); database.initDb();
const { encrypt, initKey } = await load('server/crypto.ts'); initKey();
const { openPreviewStore } = await load('shared/activity-preview-store.ts');
openPreviewStore(process.env.SWITCHBOARD_ACTIVITY_PREVIEW_ROOT, [process.env.SWITCHBOARD_DATA_DIR], 'synthetic-microsoft');
const { config } = await load('server/config.ts');
config.previewStoreId = JSON.parse(fs.readFileSync(path.join(process.env.SWITCHBOARD_ACTIVITY_PREVIEW_ROOT, 'store.json'), 'utf8')).storeId;
const output = await load('server/activity-output.ts'); output.initOutputStore(); output.setOutputRetention(4);
const { plugins } = await load('server/plugins/manager.ts');
const { createUser } = await load('server/users.ts');
const { execute } = await load('server/proxy.ts');
const owner = createUser('fixture', 'admin');
const caller = { source: 'proxy', tokenId: null, tokenName: null, ip: null, userAgent: null };
const originalFetch = globalThis.fetch;
let upstream = () => Response.json(drive), requests = 0, assertions = 0, sequence = 0;
const adapters: any[] = [];
function register(service: any, methodId: string, cfg = connectionConfig, creds = credentials()) {
  plugins.services.set(service.id, { pluginId: 'microsoft-fixture', service });
  const id = 'fixture-' + ++sequence;
  database.run('INSERT INTO connections(id,user_id,service_id,method_id,name,config_enc,credentials_enc,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)', id, owner.id, service.id, methodId, id, encrypt({ authenticationSource: 'fixture-source', ...cfg }), encrypt(creds), 'ok', Date.now(), Date.now());
  return id;
}
plugins.services.set('fixture-source', { pluginId: 'microsoft-fixture', service: { id: 'fixture-source', name: 'Synthetic authentication source', authMethods: [{ id: 'source', credentialProvider: { format: 'microsoft/web-session/v1' } }] } });
database.run('INSERT INTO connections(id,user_id,service_id,method_id,name,config_enc,credentials_enc,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)', 'fixture-source', owner.id, 'fixture-source', 'source', 'fixture-source', encrypt({}), encrypt({}), 'ok', Date.now(), Date.now());
async function call(id: string, url: string, preview: any, body?: string | null, status = 200, method = 'GET', requestCount = 1) {
  const before = requests;
  const result = await execute(owner, id, { method, url: url.startsWith('/v1.0/') ? 'https://graph.microsoft.com' + url : url }, undefined, caller);
  assert.equal(result.response.status, status);
  const text = await result.response.text();
  if (body !== undefined) assert.equal(text, body ?? '');
  assert.equal(requests - before, requestCount);
  assert.deepEqual(output.readOutput(owner.id, result.auditId, true).preview, preview, JSON.stringify({ availability: output.outputCaptureAvailability(), stored: output.readOutput(owner.id, result.auditId, true), url: result.url.href, service: database.one('SELECT service_id FROM connections WHERE id = ?', id) }));
  assert.equal(database.all('SELECT * FROM audit_outputs WHERE audit_id = ?', result.auditId).length, preview ? 1 : 0);
  assertions++;
  return result;
}
globalThis.fetch = async (input: any, opts: any) => {
  requests++;
  if (new URL(String(input)).hostname === '127.0.0.1') return originalFetch(input, opts);
  assert(['graph.microsoft.com', 'outlook.office.com'].includes(new URL(String(input)).hostname), 'Unexpected upstream');
  return upstream();
};
try {
  for (const service of graphServices()) {
    const rule = service.outputPreviews[0], field = Object.keys(rule.fields)[0];
    const fixture = service.id === 'outlook-calendar' ? calendarBody : drive;
    const expected = { [field]: fixture[field] }, body = JSON.stringify(fixture);
    for (const methodId of rule.authMethods) {
      const id = register(service, methodId);
      const response = (text = body, status = 200, type = 'application/json') => new Response(status === 204 ? null : text, { status, headers: { 'content-type': type, 'x-fixture': 'preserved' } });
      upstream = () => response();
      const result = await call(id, rule.path, expected, body);
      assert.equal(result.response.headers.get('x-fixture'), 'preserved');
      assert.equal(result.response.headers.get('content-type'), 'application/json');
      // Exact matching, with source policy still permitting these neighboring calls.
      await call(id, rule.path + '?$select=' + field, undefined, body);
      await call(id, rule.path + '/other', undefined, body);
      await call(id, rule.path, undefined, body, 200, 'POST');
      for (const invalid of [undefined, 1, '', 'bad value', 'a'.repeat(33), 'echo-fixture-access']) {
        const bad = JSON.stringify({ ...fixture, [field]: invalid });
        upstream = () => response(bad);
        await call(id, rule.path, undefined, bad);
      }
      for (const [status, type] of [[201, 'application/json'], [204, 'application/json'], [302, 'application/json'], [500, 'application/json'], [200, 'text/plain'], [200, 'application/octet-stream'], [200, 'text/event-stream']] as const) {
        upstream = () => response(body, status, type);
        await call(id, rule.path, undefined, status === 204 ? null : body, status);
      }
      for (const bad of ['', '{bad', '[]']) {
        upstream = () => response(bad); await call(id, rule.path, undefined, bad);
      }
      for (const size of [16384, 16385]) {
        const padded = body.padEnd(size);
        upstream = () => response(padded);
        await call(id, rule.path, size === 16384 ? expected : undefined, padded);
      }
      upstream = () => response(body, 200, 'application/graph+json');
      await call(id, rule.path, expected, body);
      output.setOutputRetention(0);
      await call(id, rule.path, undefined, body);
      output.setOutputRetention(4);
      config.previewRestoreSafe = false; output.initOutputStore();
      await call(id, rule.path, undefined, body);
      config.previewRestoreSafe = true; output.initOutputStore();
      upstream = () => response();
      // A cancelled response never commits a projection.
      const cancelled = await execute(owner, id, { method: 'GET', url: rule.origin + rule.path }, undefined, caller);
      await cancelled.response.body.cancel();
      assert.equal(output.readOutput(owner.id, cancelled.auditId, true).preview, undefined);
      upstream = () => new Response(new ReadableStream({ start(c) { c.enqueue(Buffer.from(body)); c.error(Error('fixture interrupted')); } }), { headers: { 'content-type': 'application/json' } });
      const interrupted = await execute(owner, id, { method: 'GET', url: rule.origin + rule.path }, undefined, caller);
      await assert.rejects(interrupted.response.text(), /fixture interrupted/);
      assert.equal(output.readOutput(owner.id, interrupted.auditId, true).preview, undefined);
      // Core-owned OAuth refresh is represented by the fixture's bearer hook.
      if (methodId !== 'authentication-source') {
        let attempt = 0; upstream = () => response(body, ++attempt === 1 ? 401 : 200);
        const retried = await call(id, rule.path, expected, body, 200, 'GET', 2);
        assert.equal(retried.sent.retried, true);
        upstream = () => response(body, 401);
        await call(id, rule.path, undefined, body, 401, 'GET', 2);
      }
    }
    // Unsupported auth IDs never qualify, even with an otherwise identical call.
    const excluded = { ...service, authMethods: [...service.authMethods, { id: 'excluded', authorize() {} }] };
    upstream = () => Response.json(fixture);
    await call(register(excluded, 'excluded'), rule.path, undefined, body);
    // Existing host policy rejects an untrusted destination before any fetch.
    const before = requests;
    await assert.rejects(execute(owner, register(service, 'oauth'), { method: 'GET', url: 'https://wrong.example' + rule.path }, undefined, caller), /not an allowed host/);
    assert.equal(requests, before);
    if (service.id === 'outlook-calendar') {
      // Real authentication-source's Outlook backend is explicitly omitted.
      await call(register(service, 'authentication-source', { ...connectionConfig, backend: 'outlook' }), '/me/calendar', undefined, body);
    }
  }
  for (const kind of ['word', 'powerpoint', 'teams']) {
    const ports = new Set();
    for (let start = 0; start < 2; start++) {
      const adapter = kind === 'teams'
        ? await createAdapter({ dataDir: temp, settings: {} }, { createSession: profile => ({ profile, adopt() {}, close: async () => {} }), createAPI: () => ({}) })
        : await createContentPlugin(contentContext, { id: kind, name: kind });
      adapters.push(adapter);
      const service = adapter.services[0], rule = service.outputPreviews[0]; ports.add(rule.origin);
      const field = kind === 'teams' ? 'transport' : 'kind';
      const selected = kind === 'teams' ? 'api' : kind;
      const privateFixture = { [field]: selected, privateText: 'Private document or chat', nested: { token: 'fixture-access' } };
      assert.deepEqual({ ...output.projectOutput(Buffer.from(JSON.stringify(privateFixture)), rule, ['fixture-access']) }, { [field]: selected });
      for (const value of ['echo-fixture-access', 'a'.repeat(33), 'bad value', '', 1, undefined]) {
        assert.equal(output.projectOutput(Buffer.from(JSON.stringify({ ...privateFixture, [field]: value })), rule, ['fixture-access']), undefined);
      }
      const id = register(service, 'authentication-source', { access: 'read' }, { profile: '44444444-4444-4444-8444-444444444444' });
      await call(id, '/capabilities', kind === 'teams' ? { transport: 'api' } : { kind });
      await call(id, '/capabilities?extra=1', undefined);
      // The previous start's port is never in this instance's declaration.
      if (start) assert.notEqual(rule.origin, adapters.at(-2).services[0].outputPreviews[0].origin);
    }
    assert.equal(ports.size, 2);
  }
  console.log(`PASS: ${assertions} audited calls, nine Graph auth profiles, Outlook backend omission, six loopback instances; exact projections and unchanged status/headers/bytes/counts.`);
} finally {
  globalThis.fetch = originalFetch;
  for (const adapter of adapters) await adapter.dispose();
  database.db.close(); fs.rmSync(temp, { recursive: true, force: true });
}

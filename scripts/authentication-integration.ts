// Real shared-credential and proxy paths; all provider/source traffic is synthetic.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { createConsumerAuth, baseUrl } from '../plugins/microsoft-authentication/consumer.ts';
import { FORMAT, TEAMS_CLIENT_ID } from '../plugins/microsoft-authentication/lib/session.js';
const root = process.env.SWITCHBOARD_ROOT;
assert(root, 'Set SWITCHBOARD_ROOT');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'authentication-recovery-'));
process.env.SWITCHBOARD_DATA_DIR = temp;
const load = (name: string) => import(pathToFileURL(path.join(root, name)).href);
const database = await load('server/db.ts'); database.initDb();
(await load('server/crypto.ts')).initKey();
const { plugins } = await load('server/plugins/manager.ts');
const { createUser } = await load('server/users.ts');
const { startConnect } = await load('server/connections.ts');
const { execute } = await load('server/proxy.ts');
const user = createUser('fixture', 'admin');
const graph = 'https://graph.microsoft.com', outlook = 'https://outlook.office.com';
const tenant = '11111111-1111-4111-8111-111111111111', oid = '22222222-2222-4222-8222-222222222222';
let renewed = false, failExchange = false, acquisitions = 0;
const original = globalThis.fetch;
const bundle = () => ({ format: FORMAT, clientId: TEAMS_CLIENT_ID, origin: 'https://teams.microsoft.com', identity: { tenant, oid }, refreshToken: renewed ? 'new-session' : 'old-session', accessToken: renewed ? 'new-graph' : 'old-graph', expiresAt: Date.now() + 3600000, scope: `${graph}/User.Read ${graph}/Calendars.Read`, generation: renewed ? '44444444-4444-4444-8444-444444444444' : '33333333-3333-4333-8333-333333333333' });
globalThis.fetch = async (input: any, opts: any) => {
  const url = String(input);
  if (url.includes('/token')) {
    if (failExchange) return Response.json({ error: 'temporarily_unavailable' }, { status: 503 });
    const session = opts.body.get('refresh_token') === 'new-session' ? 'new' : 'old';
    return Response.json({ access_token: session + '-resource', refresh_token: session + '-session', token_type: 'Bearer', expires_in: 3600, scope: `${graph}/User.Read ${graph}/Calendars.Read ${outlook}/Mail.ReadWrite` });
  }
  if (url.includes('/me?$select=')) return Response.json({ id: oid });
  assert.ok(url.startsWith(graph + '/') || url.startsWith(outlook + '/'), 'No unexpected network traffic');
  const rejected = renewed && !opts.headers.get('authorization')?.includes('new-');
  return Response.json({}, { status: rejected ? 401 : 200 });
};
try {
  for (const [id, scope, endpoint] of [['outlook-mail', 'Mail.ReadWrite', '/me/messages'], ['outlook-calendar', 'Calendars.Read', '/me/events']]) {
    renewed = false;
    const auth = createConsumerAuth({ credentials: { obtain: async (_ref: string, opts: any) => {
      acquisitions++;
      assert.equal(opts.forceAuthentication, false);
      return { format: FORMAT, bundle: bundle(), account: { id: `${tenant}:${oid}` } };
    } } } as any, { id, scopes: [scope] });
    // Core validates this source binding before every token call.
    plugins.services.set('fixture-source', { pluginId: 'fixture', service: { id: 'fixture-source', name: 'Fixture source', authMethods: [{ id: 'source', credentialProvider: { format: FORMAT } }] } });
    const { encrypt } = await load('server/crypto.ts');
    const source = 'source-' + id;
    database.run('INSERT INTO connections(id,user_id,service_id,method_id,name,config_enc,credentials_enc,status,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)', source,user.id,'fixture-source','source',source,encrypt({}),encrypt({}),'ok',Date.now(),Date.now());
    plugins.services.set(id, { pluginId: 'fixture', service: { id, name: id, baseUrl, authMethods: [auth] } });
    const connected = await startConnect(user, { service: id, method: auth.id, name: id, config: { authenticationSource: source } });
    assert.equal(connected.status, 'connected');
    const connection = connected.connection.id;
    assert.ok(database.one('SELECT credential_id FROM connections WHERE id=?', connection).credential_id);
    renewed = true;
    if (id === 'outlook-mail') {
      failExchange = true;
      await assert.rejects(execute(user, connection, { method: 'GET', url: endpoint }));
      failExchange = false;
    }
    const before = acquisitions;
    const result = await execute(user, connection, { method: 'GET', url: endpoint });
    assert.equal(result.response.status, 200);
    assert.equal(result.sent.retried, true);
    assert.equal(acquisitions, before + 1);
    await result.response.body?.cancel();
  }
  console.log('PASS: Graph and Outlook shared credentials recover after source renewal without reconnecting; transient failure remains recoverable.');
} finally { globalThis.fetch = original; database.db.close(); fs.rmSync(temp, { recursive: true, force: true }); }

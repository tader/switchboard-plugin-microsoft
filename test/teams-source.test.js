import { test } from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { createAdapter } from '../plugins/teams/index.js';
import { createConsumerAuth } from '../plugins/microsoft-authentication/consumer.ts';
import assert from 'node:assert/strict';
import { MicrosoftTeamsSession } from '../plugins/teams/lib/microsoft-session.js';
import * as web from '../plugins/microsoft-authentication/lib/session.js';
const tenant = '11111111-1111-4111-8111-111111111111', oid = '22222222-2222-4222-8222-222222222222';
const bundle = () => ({ format:web.FORMAT,clientId:web.TEAMS_CLIENT_ID,origin:'https://teams.microsoft.com',identity:{tenant,oid},refreshToken:'fixture-refresh',accessToken:'graph-access',expiresAt:Date.now()+3600_000,scope:'https://graph.microsoft.com/User.ReadBasic.All',generation:'33333333-3333-4333-8333-333333333333' });
const json = (value,status=200) => new Response(JSON.stringify(value),{status});
const chatToken = `e30.${Buffer.from(JSON.stringify({skypeid:`orgid:${oid}`,tid:tenant})).toString('base64url')}.fixture`;
function fixture() {
  const supplied=[],saved=[],requests=[];let rejectRefresh=false,rejectChat=false;
  const fetcher=async(url,opts)=>{
    requests.push(String(url));
    if(String(url).endsWith('/token')){
      assert.equal(opts.body.get('client_id'),web.TEAMS_CLIENT_ID);
      assert.equal(opts.headers.Origin,'https://teams.microsoft.com');
      if(rejectRefresh){rejectRefresh=false;return json({error:'invalid_grant'},400);}
      return json({access_token:'resource-'+requests.length,refresh_token:'rotated-'+requests.length,token_type:'Bearer',expires_in:3600,scope:opts.body.get('scope')});
    }
    if(String(url).includes('graph.microsoft.com'))return json({id:oid,mail:'fixture@example.test'});
    if(rejectChat)return json({},401);
    return json({tokens:{skypeToken:chatToken,expiresIn:3600},region:'emea',regionGtms:{}});
  };
  const webSession={...web,format:web.FORMAT,teamsClientId:web.TEAMS_CLIENT_ID,refresh:(b,o)=>web.refresh(b,{...o,fetcher}),identify:(b,o)=>web.identify(b,{...o,fetcher})};
  const ctx={require:()=>({webSession}),credentials:{obtain:async(ref,opts)=>{supplied.push({ref,...opts});return {format:web.FORMAT,bundle:bundle(),account:{id:`${tenant}:${oid}`,label:'fixture@example.test'}};}}};
  const make=(profile='profile',dir='/unused',settings={})=>new MicrosoftTeamsSession(ctx,profile,dir,settings, {fetcher,createAuthenticator:()=>{throw Error('Unexpected browser');}});
  const session=make();session.persistCredentials=async value=>saved.push(structuredClone(value));
  return {ctx,session,make,supplied,saved,requests,rejectRefresh:()=>{rejectRefresh=true;},rejectChat:()=>{rejectChat=true;}};
}
test('Teams source renews the issuing client, persists rotations, and works from a cached bundle offline',async()=>{
 const f=fixture();const credentials=await f.session.connectSource({authenticationSource:'laptop'});
 assert.equal(credentials.authMode,'microsoft');assert.equal(f.supplied.length,1);assert.ok(f.saved.some(c=>c.microsoft.refreshToken.startsWith('rotated-')));
 const restarted=f.make();restarted.adopt(credentials);restarted.sourceConfig={authenticationSource:'laptop'};
 await restarted.ensure();assert.equal(f.supplied.length,1);
 await restarted.ensureSearch();assert.ok(restarted.credentials.tokens.substrate);assert.equal(f.supplied.length,1);
 f.rejectChat();await assert.rejects(f.session.ensure(true),{code:'api_unauthorized'});
 assert.ok(f.saved.at(-1).microsoft.refreshToken.startsWith('rotated-'));
 await f.session.close();await restarted.close();
});
test('terminal Teams renewal rejection reacquires from the selected source, never an independent browser',async()=>{
 const f=fixture();await f.session.connectSource({authenticationSource:'laptop'});f.rejectRefresh();await f.session.ensure(true);
 assert.equal(f.supplied.length,2);assert.equal(f.supplied[1].ref,'laptop');assert.equal(f.supplied[1].forceAuthentication,false);
 await assert.rejects(f.session.login(),{code:'signin_required'});await f.session.close();
});
test('Teams and Outlook clients cannot exchange origins during import or renewal',()=>{
 const b=bundle();assert.equal(web.validateBundle(b).clientId,web.TEAMS_CLIENT_ID);
 assert.throws(()=>web.validateBundle({...b,origin:'https://outlook.cloud.microsoft'}),{code:'invalid_credentials'});
 assert.throws(()=>web.validateBundle({...b,clientId:web.CLIENT_ID}),{code:'invalid_credentials'});
});


test('Teams offers the same shared authentication picker as other Microsoft products and migrates in place', async t => {
  const f = fixture();
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-source-adapter-'));
  let directoryChecks = 0;
  const adapter = await createAdapter({ ...f.ctx, dataDir: directory, settings: {} }, {
    createSession: f.make,
    createAPI: () => ({ directory: async () => { directoryChecks++; } }),
  });
  t.after(async () => { await adapter.dispose(); await fs.rm(directory, { recursive: true, force: true }); });
  const methods = adapter.services[0].authMethods;
  assert.deepEqual(methods.filter(method => !method.hidden).map(method => method.id), ['authentication-source']);
  const method = methods[0];
  const outlook = createConsumerAuth(f.ctx, { id: 'outlook-mail', scopes: [] });
  assert.equal(method.name, outlook.name);
  assert.deepEqual(method.fields, outlook.fields);
  assert.deepEqual(method.credentialConsumer, outlook.credentialConsumer);
  assert.equal(method.unavailable, undefined);
  const profile = randomUUID();
  const config = { authenticationSource: 'peer/laptop/microsoft-source' };
  const connected = await method.connect({ config, connection: { credentials: { profile, authMode: 'browser' } } });
  assert.equal(connected.credentials.profile, profile);
  assert.equal(connected.credentials.authMode, 'microsoft');
  assert.deepEqual(connected.config, config);
  assert.equal(connected.account.id, `${tenant}:${oid}`);
  assert.equal(f.supplied[0].ref, config.authenticationSource);
  assert.equal(directoryChecks, 1);
  const request = { method: 'GET', url: new URL('/status', adapter.services[0].baseUrl), headers: new Headers() };
  await method.authorize(request, connected);
  const response = await fetch(request.url, { headers: request.headers });
  assert.equal(response.status, 200);
  assert.equal((await response.json()).authMode, 'microsoft');
  const before = f.requests.length;
  const renewal = { method: 'GET', url: new URL('/status?renew=true', adapter.services[0].baseUrl), headers: new Headers() };
  await method.authorize(renewal, connected);
  assert(f.requests.length > before, 'the explicit status probe exercises token renewal');
  const renewed = await fetch(renewal.url, { headers: renewal.headers });
  assert.equal(renewed.status, 200);
  assert.equal((await renewed.json()).ready, true);
  assert.equal(directoryChecks, 1, 'status renewal must not list chats or probe directory data');

});

test('Teams shows an upgrade message instead of silently omitting shared authentication on older cores', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-source-upgrade-'));
  const adapter = await createAdapter({ dataDir: directory, settings: {} });
  t.after(async () => { await adapter.dispose(); await fs.rm(directory, { recursive: true, force: true }); });
  const methods = adapter.services[0].authMethods;
  const visible = methods.filter(method => !method.hidden);
  assert.equal(visible.length, 1);
  assert.equal(visible[0].id, 'authentication-source');
  assert.match(visible[0].unavailable, /Upgrade Switchboard/);
  await assert.rejects(visible[0].connect({ config: {} }), { code: 'core_upgrade_required' });
  assert(methods.find(method => method.id === 'browser').authorize);
  assert(methods.find(method => method.id === 'tokens').authorize);
});


test('Teams retries a reconnected source before forcing another browser authentication', async () => {
  const f = fixture();
  await f.session.connectSource({ authenticationSource: 'laptop' });
  const previous = structuredClone(f.session.credentials);
  const latest = { ...bundle(), generation: randomUUID(), refreshToken: 'fresh-local-login' };
  const calls = [];
  f.ctx.credentials.obtain = async (ref, opts) => {
    calls.push(opts.forceAuthentication);
    assert.equal(opts.forceAuthentication, false, 'a fresh local login must not trigger another browser');
    return { format: web.FORMAT, bundle: latest, account: { id: `${tenant}:${oid}`, label: 'fixture@example.test' } };
  };
  f.rejectRefresh();
  await f.session.ensure(true);
  assert.deepEqual(calls, [false]);
  assert.notEqual(f.session.credentials.microsoft.generation, previous.microsoft.generation);
  assert.equal(f.session.credentials.microsoft.generation, latest.generation);
  assert.equal(f.session.status().ready, true);
  await f.session.close();
});

test('Teams forces authentication only after the current source session also fails renewal', async () => {
  const f = fixture();
  await f.session.connectSource({ authenticationSource: 'laptop' });
  const original = f.session.web.refresh;
  let forced = false, refreshes = 0;
  f.session.web.refresh = async (b, opts) => {
    refreshes++;
    if (!forced) { const e = new Error('Expired fixture session'); e.reauthenticationRequired = true; throw e; }
    return original(b, opts);
  };
  f.ctx.credentials.obtain = async (ref, opts) => {
    f.supplied.push({ ref, ...opts });
    forced = opts.forceAuthentication;
    return { format: web.FORMAT, bundle: bundle(), account: { id: `${tenant}:${oid}`, label: 'fixture@example.test' } };
  };
  await f.session.ensure(true);
  assert.deepEqual(f.supplied.slice(1).map(call => call.forceAuthentication), [false, true]);
  assert.equal(refreshes, 4);
  await f.session.close();
});

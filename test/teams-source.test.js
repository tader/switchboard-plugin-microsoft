import { test } from 'node:test';
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
  const make=()=>new MicrosoftTeamsSession(ctx,'profile','/unused',{}, {fetcher,createAuthenticator:()=>{throw Error('Unexpected browser');}});
  const session=make();session.persistCredentials=async value=>saved.push(structuredClone(value));
  return {session,make,supplied,saved,requests,rejectRefresh:()=>{rejectRefresh=true;},rejectChat:()=>{rejectChat=true;}};
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
 assert.equal(f.supplied.length,2);assert.equal(f.supplied[1].ref,'laptop');assert.equal(f.supplied[1].forceAuthentication,true);
 await assert.rejects(f.session.login(),{code:'signin_required'});await f.session.close();
});
test('Teams and Outlook clients cannot exchange origins during import or renewal',()=>{
 const b=bundle();assert.equal(web.validateBundle(b).clientId,web.TEAMS_CLIENT_ID);
 assert.throws(()=>web.validateBundle({...b,origin:'https://outlook.cloud.microsoft'}),{code:'invalid_credentials'});
 assert.throws(()=>web.validateBundle({...b,clientId:web.CLIENT_ID}),{code:'invalid_credentials'});
});

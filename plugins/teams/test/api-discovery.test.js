import test from 'node:test';
import assert from 'node:assert/strict';
import { TeamsAPI, encodeConversation } from '../lib/teams-api.js';
import { TeamsTriage } from '../lib/triage.js';
import { openapi } from '../lib/api-openapi.js';
const origin='https://emea.ng.msg.teams.microsoft.com', thread='19:older@thread.v2';
const chat=(id=thread, properties={})=>({id,threadProperties:{threadType:'chat',topic:'Older chat'},properties:{consumptionhorizon:'100;200;300',consumptionHorizonBookmark:'0;200;0',lastimreceivedtime:'1970-01-01T00:00:00.200Z',...properties}});
function fixture({pages=[{conversations:[chat()]}],partial=true,roster=true,error=null}={}){
 const calls=[],credentials={identity:{oid:'self',tenant:'tenant'},messageOrigin:origin,chatToken:{value:'test'},tokens:{chatsvcagg:{value:'test'}}};
 const api=new TeamsAPI({credentials},{fetcher:async(raw,options)=>{
  const u=new URL(raw);calls.push({u,...options});
  if(u.pathname.endsWith('/teams/users/me'))return Response.json({chats:[{id:'19:known@thread.v2',title:'CSA title',isRead:true}],teams:[],metadata:{isPartialData:partial}});
  if(u.pathname.endsWith('/members'))return Response.json({members:roster?[{id:'8:orgid:self'}]:[{id:'8:orgid:other'}]});
  if(decodeURIComponent(u.pathname).endsWith('/conversations/'+thread)) return Response.json(chat());
  if(u.pathname.endsWith('/messages'))return Response.json({messages:[{id:'200',content:'Incoming',messagetype:'Text',from:'8:orgid:peer',originalarrivaltime:200}],_metadata:{}});
  if(error)return new Response(null,{status:error});
  const index=Number(u.searchParams.get('page')??0);return Response.json(pages[index]);
 }});return{api,calls,credentials};
}
const link=page=>`${origin}/v1/users/8:orgid:self/conversations?page=${page}`;
test('regional paging follows empty filtered pages, deduplicates overlaps and replays scoped cursors',async()=>{
 const f=fixture({pages:[{conversations:[{id:'48:notes',threadProperties:{threadType:'streamofnotes'}}],_metadata:{backwardLink:link(1)}},{conversations:[chat()],_metadata:{backwardLink:link(2)}},{conversations:[chat(),chat('19:last@thread.v2')]}]});
 const first=await f.api.discoverChats({limit:2});assert.equal(first.items.length,0);assert.equal(first.hasMore,true);assert.equal(first.excludedConversations,1);
 const second=await f.api.discoverChats({limit:2,cursor:first.nextCursor});assert.equal(second.items.length,1);const count=f.calls.length;
 assert.deepEqual(await f.api.discoverChats({limit:2,cursor:first.nextCursor}),second);assert.equal(f.calls.length,count);
 const last=await f.api.discoverChats({limit:2,cursor:second.nextCursor});assert.equal(last.items.length,1);assert.equal(last.discoveryComplete,true);
 await assert.rejects(f.api.discoverChats({limit:3,cursor:first.nextCursor}),{code:'invalid_cursor'});
 f.credentials.identity.oid='other';await assert.rejects(f.api.discoverChats({limit:2,cursor:first.nextCursor}),{code:'invalid_cursor'});
});
test('continuations reject foreign origins, accounts, resource paths, credentials and fragments before fetching',async()=>{
 for(const url of ['https://example.com/v1/users/ME/conversations',`${origin}/v1/users/other/conversations`,`${origin}/v1/users/ME/conversations/messages`,`${origin}/v1/users/ME/conversations#x`,`https://user@emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations`,`${origin}/v1/users/%ZZ/conversations`]){
  const f=fixture({pages:[{conversations:[],_metadata:{backwardLink:url}}]});
  await assert.rejects(f.api.discoverChats(),{code:'api_cursor_rejected'});assert.equal(f.calls.length,1);
  await assert.rejects(f.api.conversationPage(100,url),{code:'api_cursor_rejected'});assert.equal(f.calls.length,1);
 }
});
test('relative account-scoped links work and cycles remain explicitly incomplete',async()=>{
 const f=fixture({pages:[{conversations:[chat()],_metadata:{backwardLink:'?page=1'}},{conversations:[chat()],_metadata:{backwardLink:'?page=1'}}]});
 const first=await f.api.discoverChats(),second=await f.api.discoverChats({cursor:first.nextCursor});
 assert.equal(second.items.length,0);assert.equal(second.hasMore,false);assert.equal(second.truncated,true);assert.equal(second.discoveryComplete,false);
});
test('projection preserves cleared/active bookmarks and unknown marker states; excludes spaces and system streams',()=>{
 const {api}=fixture();assert.equal(api.discoveredChat(chat()).isRead,false);
 assert.equal(api.discoveredChat(chat(thread,{consumptionHorizonBookmark:'250;300;400'})).isRead,true);
 assert.equal(api.discoveredChat(chat(thread,{consumptionhorizon:undefined})).isRead,null);
 assert.equal(api.discoveredChat(chat(thread,{consumptionHorizonBookmark:'malformed'})).isRead,null);
 assert.equal(api.discoveredChat(chat(thread,{consumptionhorizon:'999999999999999999999;1;1'})).isRead,null);
 for(const type of ['space','streamofnotes','streamofmentions','unknown'])assert.equal(api.discoveredChat({...chat(),threadProperties:{threadType:type}}),null);
});
test('partial snapshots merge missing chats into listing and triage without overwriting CSA metadata',async()=>{
 const f=fixture({pages:[{conversations:[chat('19:known@thread.v2'),chat()]}]});
 const listing=await f.api.chats();assert.equal(listing.items.length,2);assert.equal(listing.items[0].title,'CSA title');assert.equal(listing.discovery.added,1);assert.equal(listing.discoveryComplete,true);
 const inbox=await new TeamsTriage(f.api,null).inbox();assert.equal(inbox.items.length,1);assert.equal(inbox.items[0].conversationId,encodeConversation('chat',thread));assert.equal(inbox.coverage.discoveryPartial,true);assert.equal(inbox.mayMarkRead,false);
 assert.ok(f.calls.every(c=>!c.method||c.method==='GET'));
});
test('bounded automatic discovery retains continuation limits and failures without dropping known chats',async()=>{
 const pages=Array.from({length:11},(_,i)=>({conversations:[chat(`19:page${i}@thread.v2`)],_metadata:{backwardLink:link(i+1)}}));
 const f=fixture({pages});const directory=await f.api.directory(true);assert.equal(directory.discovery.pages,10);assert.equal(directory.discovery.truncated,true);assert.equal(directory.chats.length,11);
 const bad=fixture({error:503});const listing=await bad.api.chats();assert.equal(listing.items.length,1);assert.equal(listing.discoveryComplete,false);assert.equal(listing.discovery.reason,'api_response_failed');
 await assert.rejects(fixture({error:401}).api.chats(),{code:'api_unauthorized'});
});
test('newly discovered chat writes require a current roster containing the signed-in account',async()=>{
 const f=fixture({roster:false});const id=encodeConversation('chat',thread);
 await f.api.chats();await assert.rejects(f.api.prepareSend(id,{text:'Do not send'}),{code:'conversation_unavailable'});
 await assert.rejects(f.api.prepareMutation(id,'200','reaction',{reaction:'like',selected:true}),{code:'conversation_unavailable'});
 await assert.rejects(f.api.prepareSendForTriage({kind:'chat',conversationId:id,threadId:thread},{text:'Do not send'}),{code:'conversation_unavailable'});
 assert.ok(f.calls.every(c=>!c.method||c.method==='GET'));
});
test('malformed discovery never becomes an empty success and the schema exposes pagination',async()=>{
 for(const page of [{},{conversations:[null]},{conversations:[{id:'bad/path'}]},{conversations:[],_metadata:{backwardLink:{}}}])await assert.rejects(fixture({pages:[page]}).api.discoverChats(),{code:'api_schema_changed'});
 assert.equal(openapi.paths['/chats/discovery'].get.operationId,'pageTeamsChatDiscovery');
});


test('selected chat IDs outside the directory scan resolve through exact metadata and current roster', async () => {
 const f=fixture({partial:false,pages:[{conversations:[]}]});
 const resolved=await f.api.resolve(encodeConversation('chat',thread));
 assert.equal(resolved.id,thread);assert.equal(resolved.discoverySource,'regional_conversations');
 assert.ok(f.calls.some(c=>c.u.pathname.endsWith('/members')));
 const denied=fixture({partial:false,roster:false});
 await assert.rejects(denied.api.resolve(encodeConversation('chat',thread)),{code:'conversation_unavailable'});
});

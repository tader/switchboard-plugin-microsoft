import test from 'node:test';
import assert from 'node:assert/strict';
import { TeamsAPI, encodeConversation } from '../lib/teams-api.js';
const id=encodeConversation('chat','48:notes');
function fixture({wrongId=false,wrongType=false,missing=false}={}){
 const calls=[];const api=new TeamsAPI({credentials:{identity:{oid:'self',tenant:'tenant',name:'Me'},messageOrigin:'https://emea.ng.msg.teams.microsoft.com',chatToken:{value:'fixture'},tokens:{}}},{fetcher:async(url,options)=>{
  calls.push({url,options});
  const path=decodeURIComponent(new URL(url).pathname);
  if(path==='/v1/users/ME/conversations/48:notes')return missing?new Response('{}',{status:404}):Response.json({id:wrongId?'48:other':'48:notes',threadProperties:{threadType:wrongType?'chat':'streamofnotes'}});
  if(path.endsWith('/48:notes/messages')){
   if(options.method==='POST')return Response.json({id:'300'});
   return Response.json({messages:[{id:'200',content:'Self note',from:'8:orgid:self',messagetype:'Text',originalarrivaltime:1000}],_metadata:{}});
  }
  throw Error('Unexpected self-chat request '+path);
 }});return{api,calls};
}
test('self notes are discoverable while other system streams remain excluded',()=>{
 const {api}=fixture();assert.equal(api.discoveredChat({id:'48:notes',threadProperties:{threadType:'streamofnotes'}}).isOneOnOne,true);
 for(const raw of [{id:'48:other',threadProperties:{threadType:'streamofnotes'}},{id:'48:notes',threadProperties:{threadType:'chat'}},{id:'48:mentions',threadProperties:{threadType:'streamofmentions'}}])assert.equal(api.discoveredChat(raw),null);
});
test('self history and send preparation revalidate the exact authenticated private stream',async()=>{
 const {api,calls}=fixture();const chat=await api.resolve(id);assert.equal(chat.title,'Notes to self');
 const page=await api.history(id,{limit:5});assert.equal(page.items[0].text,'Self note');
 const prepared=await api.prepareSend(id,{text:'Self test'});assert.equal(prepared.threadId,'48:notes');
 assert.ok(calls.every(c=>new URL(c.url).pathname.startsWith('/v1/users/ME/conversations/48%3Anotes')));
 assert.ok(calls.every(c=>!c.options.method||c.options.method==='GET'),'preparation must not send');
});
test('missing or misidentified private streams cannot be used for a self-chat send',async()=>{
 for(const opts of [{wrongId:true},{wrongType:true},{missing:true}]){
  const {api,calls}=fixture(opts);await assert.rejects(api.prepareSend(id,{text:'No send'}));
  assert.ok(calls.every(c=>!c.options.method||c.options.method==='GET'));
 }
});

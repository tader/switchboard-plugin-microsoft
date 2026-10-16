import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { TeamsAPI, encodeConversation, normalize } from '../lib/teams-api.js';
import { SendLedger } from '../lib/ledger.js';
const channel='19:channel@thread.tacv2',id=encodeConversation('channel',channel);
function fixture(){
 const state={archived:false,member:true,changed:false,fail:false,wrongRoot:false,omitRoot:false,writes:0,reply:null},calls=[];
 const api=new TeamsAPI({credentials:{identity:{tenant:'tenant',oid:'self',name:'Self'},messageOrigin:'https://emea.ng.msg.teams.microsoft.com',chatToken:{value:'test'},tokens:{chatsvcagg:{value:'test'}}}}, {fetcher:async(raw,options)=>{
  const u=new URL(raw),p=decodeURIComponent(u.pathname);calls.push({u,...options});
  if(p.endsWith('/teams/users/me'))return Response.json({chats:[],teams:[{channels:[{id:channel,isMember:state.member,isArchived:state.archived}]}]});
  if(options.method==='POST'){
   state.writes++;if(state.fail)throw new Error('lost response');const payload=JSON.parse(options.body);
   state.reply={...payload,id:'200',from:'8:orgid:self',conversationid:channel,...(state.omitRoot?{}:{rootMessageId:state.wrongRoot?'999':'100'})};return Response.json({OriginalArrivalTime:200});
  }
  if(p.endsWith('/messages/100'))return Response.json({id:'100',rootMessageId:'100',content:state.changed?'Changed root':'Root',messagetype:'Text',from:'8:orgid:peer',version:state.changed?'2':'1'});
  if(p.endsWith('/messages/200'))return Response.json(state.reply);
  if(p.endsWith('/messages/101'))return Response.json({id:'101',rootMessageId:'100',content:'A reply',messagetype:'Text',from:'8:orgid:peer'});
  throw new Error('unexpected request');
 }});return{api,state,calls};
}
async function ledger(t){const dir=await fs.mkdtemp(path.join(os.tmpdir(),'teams-channel-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));return new SendLedger(dir);}
test('exact channel replies use only the composite root path, confirm root metadata and replay without writes',async t=>{
 const f=fixture(),journal=await ledger(t),input={text:'See https://example.com'},identity=['channel-reply',f.api.account(),id,'100',input];
 const verify=(r,c)=>f.api.confirm(r,c);
 const result=await journal.run('channel-reply-live-key',identity,()=>f.api.prepareChannelReply(id,'100',input),p=>f.api.send(p),'send_uncertain',verify);
 assert.equal(result.verification.status,'observed');assert.equal(result.verification.threadConfirmed,true);assert.equal(result.parentMessageId,'100');
 const post=f.calls.find(c=>c.method==='POST');assert.ok(decodeURIComponent(post.u.pathname).endsWith(`/conversations/${channel};messageid=100/messages`));assert.match(JSON.parse(post.body).content,/<a href=/);
 const replay=await new SendLedger(path.dirname(journal.file)).run('channel-reply-live-key',identity,()=>assert.fail('reprepare'),()=>assert.fail('resend'),'send_uncertain',verify);
 assert.equal(replay.replayed,true);assert.equal(f.state.writes,1);
 await assert.rejects(journal.run('channel-reply-live-key',['different-root'],()=>assert.fail(),()=>assert.fail(),'send_uncertain'),{code:'idempotency_conflict'});
});
test('reply-as-root, invalid root IDs, lost membership and archived channels prevent dispatch',async()=>{
 const f=fixture();await assert.rejects(f.api.prepareChannelReply(id,'101',{text:'No'}),{code:'invalid_thread_root'});
 await assert.rejects(f.api.prepareChannelReply(id,'100;messageid=101',{text:'No'}),{code:'invalid_message_id'});
 f.state.archived=true;await f.api.directory(true);await assert.rejects(f.api.prepareChannelReply(id,'100',{text:'No'}),{code:'api_forbidden'});
 f.state.archived=false;f.state.member=false;await f.api.directory(true);await assert.rejects(f.api.prepareChannelReply(id,'100',{text:'No'}),{code:'conversation_unavailable'});
 assert.equal(f.state.writes,0);
});
test('root and permission rechecks refuse changes between preparation and dispatch',async()=>{
 const f=fixture(),p=await f.api.prepareChannelReply(id,'100',{text:'No'});f.state.changed=true;
 await assert.rejects(f.api.send(p),{code:'message_changed'});assert.equal(f.state.writes,0);
 f.state.changed=false;f.state.archived=true;await assert.rejects(f.api.send(p),{code:'api_forbidden'});assert.equal(f.state.writes,0);
});
test('matching content with wrong or missing root metadata never confirms thread placement',async()=>{
 for(const field of ['wrongRoot','omitRoot']){
  const f=fixture();f.state[field]=true;const p=await f.api.prepareChannelReply(id,'100',{text:'Reply'}),r=await f.api.send(p);
  assert.equal((await f.api.confirm(r,p.confirmation)).status,'not_observed');
 }
 const f=fixture(),p=await f.api.prepareChannelReply(id,'100',{text:'Reply'}),r=await f.api.send(p);
 f.state.reply.parentMessageId='999';assert.equal((await f.api.confirm(r,p.confirmation)).status,'not_observed');
 f.state.reply.parentMessageId='100';f.state.reply.conversationid='19:sibling@thread.tacv2';assert.equal((await f.api.confirm(r,p.confirmation)).status,'not_observed');
});
test('uncertain channel sends remain quarantined after restart',async t=>{
 const f=fixture(),journal=await ledger(t);f.state.fail=true;const run=()=>journal.run('uncertain-channel-key',['channel',id,'100'],()=>f.api.prepareChannelReply(id,'100',{text:'Test'}),p=>f.api.send(p),'send_uncertain');
 await assert.rejects(run(),{code:'send_uncertain'});await assert.rejects(run(),{code:'send_uncertain'});assert.equal(f.state.writes,1);
});
test('native rootMessageId distinguishes replies from roots in reads',()=>{
 assert.equal(normalize({id:'100',rootMessageId:'100'}).parentMessageId,null);
 assert.equal(normalize({id:'200',rootMessageId:'100'}).parentMessageId,'100');
});

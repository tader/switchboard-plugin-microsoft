import test from 'node:test';
import assert from 'node:assert/strict';
import { TeamsAPI } from '../lib/teams-api.js';
import { project } from '../lib/messaging.js';
const peer='8:orgid:33333333-3333-4333-8333-333333333333';
function api(){return new TeamsAPI({credentials:{identity:{oid:'self',tenant:'tenant'}}});}
test('identity titles resolve exact names and exclude self; named chats and notes survive',async()=>{
 const a=api();let calls=0;
 a.request=async()=>{calls++;return {value:{mri:peer,displayName:'Alex (Al) Example'}}};
 assert.equal(await a.conversationTitle({title:peer}),'Alex (Al) Example');
 assert.equal(await a.conversationTitle({members:[{mri:a.selfMri()},{mri:peer}]}),'Alex (Al) Example');assert.equal(calls,1);
 assert.equal(await a.conversationTitle({title:'My group'}),'My group');assert.equal(await a.conversationTitle({title:'Notes to self'}),'Notes to self');
});
test('mismatched profiles fall back, negative cache and account separation',async()=>{
 const a=api();let calls=0;a.request=async()=>{calls++;return {value:{mri:'other',displayName:'Wrong'}}};
 assert.equal(await a.conversationTitle({title:peer}),peer);assert.equal(await a.participantName(peer),null);assert.equal(calls,1);
 a.credentials.identity.tenant='second';await a.participantName(peer);assert.equal(calls,2);
});
test('roster titles require current-account membership',async()=>{
 const a=api();a.request=async()=>({members:[{id:a.selfMri()},{id:peer,friendlyName:'Alex'}]});assert.equal(await a.conversationTitle({id:'chat',title:'Teams chat'}),'Alex');
 a.request=async()=>({members:[{id:peer,friendlyName:'Alex'}]});assert.equal(await a.conversationTitle({id:'chat',title:'Teams chat'}),'Teams chat');
});
test('lookups coalesce, bound concurrency, and propagate terminal errors',async()=>{
 const a=api();let calls=0,active=0,max=0;a.request=async(_s,path)=>{calls++;active++;max=Math.max(active,max);await new Promise(resolve=>setImmediate(resolve));active--;return {value:{mri:decodeURIComponent(path.split('/')[1]),displayName:'Name'}}};
 await Promise.all(Array.from({length:12},(_,i)=>a.participantName('8:orgid:'+String(i).padStart(8,'0')+'-3284-4293-9123-2975995045b0')));assert.equal(max,4);
 await Promise.all([a.participantName(peer),a.participantName(peer)]);assert.equal(calls,13);
 const b=api();b.request=async()=>{throw Object.assign(new Error('expired'),{code:'api_unauthorized'})};await assert.rejects(b.participantName(peer));assert.equal(b.participantNames.size,0);
});
test('projected names remain separate from routing identity',()=>{
 const m=project({id:'1',text:'hello',authorId:peer,author:'Alex'},'chat','self');assert.equal(m.authorId,peer);assert.equal(m.authorDisplayName,'Alex');
});

import test from 'node:test';
import assert from 'node:assert/strict';
import { linkedText } from '../lib/linkify.js';
import { messagePayload, addQuote, plainText, TeamsAPI } from '../lib/teams-api.js';
const links = text => linkedText(text).filter(part => part.link);
test('bare URLs preserve displayed text, query/fragment and balanced parentheses while leaving prose punctuation outside', () => {
  const text = 'See (https://example.com/wiki/Thing_(detail)). Then https://example.com/?a=1&b=2#here!';
  const parts = linkedText(text);
  assert.equal(parts.map(part => part.text).join(''), text);
  assert.deepEqual(links(text), [
    {text:'https://example.com/wiki/Thing_(detail)',link:'https://example.com/wiki/Thing_(detail)'},
    {text:'https://example.com/?a=1&b=2#here',link:'https://example.com/?a=1&b=2#here'},
  ]);
  assert.equal(links('[https://example.com/path]).')[0].text,'https://example.com/path');
  assert.equal(links('http://[::1]/')[0].link,'http://[::1]/');
  assert.equal(links('HTTPS://example.com')[0].text,'HTTPS://example.com');
});
test('malformed, credential-bearing, overlong and non-HTTP URLs stay literal', () => {
  for (const text of ['https://', 'https://user:password@example.com', 'https://example.com\\@evil.test', 'https://example.com/'+'a'.repeat(2000), 'javascript:alert(1)', 'ftp://example.com', 'www.example.com', 'prefixhttps://example.com']) {
    assert.equal(links(text).length,0);assert.equal(linkedText(text).map(part=>part.text).join(''),text);
  }
});
test('plain sends render escaped anchors/newlines only when an eligible URL exists', () => {
  const input='<script>\nhttps://example.com/?a=1&b=2.';
  const payload=messagePayload({text:input},'Self');
  assert.equal(payload.messagetype,'RichText/Html');
  assert.equal(payload.content,'<p>&lt;script&gt;<br><a href="https://example.com/?a=1&amp;b=2">https://example.com/?a=1&amp;b=2</a>.</p>');
  assert.equal(plainText(payload.content,payload.messagetype),input);
  assert.equal(messagePayload({text:'No link <here>'},'Self').messagetype,'Text');
  assert.equal(messagePayload({text:'https://user:secret@example.com'},'Self').messagetype,'Text');
});
test('structured links remain explicit, code stays literal and automatic links retain formatting in paragraphs/lists', () => {
  const p=messagePayload({content:[{type:'paragraph',runs:[
    {text:'https://display.example',link:'https://target.example'},
    {text:'https://code.example',marks:['code']},
    {text:'Go https://auto.example',marks:['bold']},
  ]},{type:'bulletedList',items:[[{text:'https://list.example'}]]}]},'Self');
  assert.match(p.content,/<a href="https:\/\/target.example\/">https:\/\/display.example<\/a>/);
  assert.match(p.content,/<code>https:\/\/code.example<\/code>/);
  assert.match(p.content,/<strong>Go <a href="https:\/\/auto.example\/">https:\/\/auto.example<\/a><\/strong>/);
  assert.match(p.content,/<li><a href="https:\/\/list.example\/">https:\/\/list.example<\/a><\/li>/);
  assert.equal((p.content.match(/<a /g)??[]).length,3);
});
test('automatic links share preparation for sends, edits and triage replies; quote previews stay escaped text', async () => {
  const api=new TeamsAPI({credentials:{identity:{oid:'self',tenant:'tenant',name:'Self'}}});
  api.resolve=async()=>({id:'19:chat@thread.v2'});
  api.findMessage=async()=>({message:{authorId:'8:orgid:self',text:'Before'},raw:{version:'1'}});
  const input={text:'See https://example.com/path'};
  const send=await api.prepareSend('chat',input),edit=await api.prepareMutation('chat','100','edit',{...input,expectedText:'Before'});
  const reply=await api.prepareSendForTriage({kind:'chat',conversationId:'chat',threadId:'19:chat@thread.v2'},input);
  assert.equal(send.payload.content,edit.payload.content);assert.equal(send.payload.content,reply.payload.content);
  addQuote(reply.payload,{id:'100',authorId:'8:orgid:peer',author:'Peer',text:'https://quoted.example'});
  assert.equal((reply.payload.content.match(/<a /g)??[]).length,1);
  assert.match(reply.payload.content,/<p itemprop="preview">https:\/\/quoted.example<\/p>/);
});

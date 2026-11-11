import { test } from 'node:test';
import assert from 'node:assert/strict';
import { captureRefreshResponse } from '../plugins/microsoft-authentication/lib/browser.js';
import { CLIENT_ID } from '../plugins/microsoft-authentication/lib/session.js';

function fixture({ origin = 'https://outlook.cloud.microsoft', host = 'login.microsoftonline.com', clientId = CLIENT_ID } = {}) {
  const listeners = new Set();
  const tenant = '11111111-1111-4111-8111-111111111111', oid = '22222222-2222-4222-8222-222222222222';
  const url = `https://${host}/${tenant}/oauth2/v2.0/token`;
  const emit = (method, params) => { for (const listener of listeners) listener({ sessionId: 'page', method, params }); };
  const page = { sessionId: 'page', async call(method) {
    if (method === 'Page.navigate') queueMicrotask(() => {
      emit('Network.requestWillBeSent', { requestId: 'token', request: { method: 'POST', url, headers: { Origin: origin }, postData: new URLSearchParams({ client_id: clientId, grant_type: 'authorization_code' }).toString() } });
      emit('Network.responseReceived', { requestId: 'token', response: { status: 200, url } });
      emit('Network.loadingFinished', { requestId: 'token' });
    });
    if (method === 'Network.getResponseBody') return { body: JSON.stringify({ refresh_token: 'fixture-refresh-secret', id_token: `e30.${Buffer.from(JSON.stringify({ tid: tenant, oid })).toString('base64url')}.fixture` }) };
    if (method === 'Page.getFrameTree') return { frameTree: { frame: { url: 'https://outlook.cloud.microsoft/' } } };
    return {};
  } };
  const cdp = { subscribe(listener) { listeners.add(listener); return () => listeners.delete(listener); }, exit: new Promise(() => {}) };
  return { page, cdp, listeners };
}

test('capture binds a Microsoft HTTPS token response to the observed Outlook client and origin', async () => {
  const f = fixture();
  const value = await captureRefreshResponse(f.page, f.cdp, { timeout: 100 });
  assert.equal(value.refreshToken, 'fixture-refresh-secret');
  assert.equal(value.clientId, CLIENT_ID); assert.equal(value.origin, 'https://outlook.cloud.microsoft');
  assert.equal(value.sourceGrant, 'authorization_code'); assert.equal(f.listeners.size, 0);
});

test('foreign hosts, clients and origins never yield captured credentials', async () => {
  for (const options of [{ host: 'login.microsoftonline.com.evil.example' }, { clientId: 'another-client' }, { origin: 'https://evil.example' }]) {
    const f = fixture(options);
    await assert.rejects(captureRefreshResponse(f.page, f.cdp, { timeout: 5 }), { code: 'refresh_capture_timeout' });
    assert.equal(f.listeners.size, 0);
  }
});

test('cancelling capture removes token listeners and never returns partial credentials', async () => {
  const f = fixture(), controller = new AbortController(); controller.abort();
  await assert.rejects(captureRefreshResponse(f.page, f.cdp, { timeout: 100, signal: controller.signal }), { code: 'refresh_check_cancelled' });
  assert.equal(f.listeners.size, 0);
});

test('Teams capture accepts only the Teams client paired with a Teams origin', async () => {
  const clientId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
  const f = fixture({ clientId, origin: 'https://teams.microsoft.com' });
  assert.equal((await captureRefreshResponse(f.page, f.cdp, { app: 'teams', timeout: 100 })).clientId, clientId);
  for (const options of [{ clientId, origin: 'https://outlook.cloud.microsoft' }, { clientId: CLIENT_ID, origin: 'https://teams.microsoft.com' }]) {
    const wrong = fixture(options);
    await assert.rejects(captureRefreshResponse(wrong.page, wrong.cdp, { app: 'teams', timeout: 5 }), { code: 'refresh_capture_timeout' });
  }
});

test('core profile uses an ephemeral app context, preserves shared storage, and disposes it', async () => {
  const { acquireBrowserRefresh } = await import('../plugins/microsoft-authentication/lib/browser.js');
  const f = fixture({ clientId: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346', origin: 'https://teams.microsoft.com' });
  const calls = [];
  f.cdp.call = async (method, args) => {
    calls.push([method, args]);
    if(method==='Target.createBrowserContext') return {browserContextId:'isolated'};
    if(method==='Storage.getCookies') return {cookies:[{name:'sso',value:'synthetic',domain:'login.microsoftonline.com',path:'/',size:10,session:true}]};
    return {};
  };
  f.cdp.page = async (url, context) => { assert.equal(context,'isolated'); return f.page; };
  let options;
  const sharedProfile={run:async(o,fn)=>{options=o;return fn(f.cdp,new AbortController().signal);}};
  const value=await acquireBrowserRefresh('/unused',{silent:true,sharedProfile});
  assert.equal(value.refreshToken,'fixture-refresh-secret');assert.equal(options.headless,true);
  assert.ok(!calls.some(([method])=>method==='Storage.clearDataForOrigin'));
  assert.equal(calls.filter(([method])=>method==='Storage.setCookies').length,2);
  assert.equal(calls.filter(([method])=>method==='Storage.setCookies')[0][1].cookies[0].size,undefined);
  assert.deepEqual(calls.at(-1),['Target.disposeBrowserContext',{browserContextId:'isolated'}]);
});

test('core temporary context is disposed when acquisition is cancelled', async () => {
  const { acquireBrowserRefresh }=await import('../plugins/microsoft-authentication/lib/browser.js');
  const f=fixture(),calls=[]; const controller=new AbortController();
  f.cdp.call=async method=>{calls.push(method);if(method==='Target.createBrowserContext')return {browserContextId:'isolated'};if(method==='Storage.getCookies')return {cookies:[]};return {};};
  f.cdp.page=async()=>{controller.abort();return f.page;};
  await assert.rejects(acquireBrowserRefresh('/unused',{silent:true,sharedProfile:{run:(_,fn)=>fn(f.cdp,controller.signal)}}),{code:'refresh_check_cancelled'});
  assert.equal(calls.at(-1),'Target.disposeBrowserContext');assert.equal(f.listeners.size,0);
});

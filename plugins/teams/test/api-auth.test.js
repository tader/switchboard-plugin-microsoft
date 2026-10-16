import test from 'node:test';
import assert from 'node:assert/strict';
import { generateKeyPairSync, sign } from 'node:crypto';
import { authorization, callback, validateIdentity, clientId, capture } from '../lib/oauth.js';

test('OAuth callbacks enforce exact origin, state, stage, and sanitized errors', () => {
  const request = authorization('skype');
  const url = `https://teams.microsoft.com/go#state=${request.state}&access_token=secret`;
  assert.equal(callback(url, request), 'secret');
  assert.equal(callback(url.replace('teams.microsoft.com', 'teams.microsoft.com.evil.test'), request), null);
  assert.throws(() => callback(url.replace(request.state, 'wrong'), request), /state_mismatch/);
  assert.throws(() => callback(url.replace('access_token', 'id_token'), request), /missing_token/);
  assert.throws(() => callback(url + '&access_token=other', request), /missing_token/);
  assert.throws(() => callback(`https://teams.microsoft.com/go#state=${request.state}&error=access_denied&error_description=AADSTS12345%20secret`, request), error => error.message === 'oauth_error:access_denied:AADSTS12345');
  assert.equal(new URL(request.url).searchParams.get('resource'), 'https://api.spaces.skype.com');
});

test('identity verification rejects tampering, expired tokens, nonce and tenant mismatch', async () => {
  const { publicKey, privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
  const tenant = '11111111-1111-1111-1111-111111111111';
  const request = authorization('teams', tenant);
  const key = { ...publicKey.export({ format: 'jwk' }), kid: 'test', use: 'sig' };
  const fetcher = async () => ({ ok: true, json: async () => ({ keys: [key] }) });
  const claims = { tid: tenant, aud: clientId, iss: `https://sts.windows.net/${tenant}/`, nonce: request.nonce, exp: Date.now() / 1000 + 3600 };
  const token = (changes = {}, signingKey = privateKey) => {
    const parts = [{ alg: 'RS256', kid: 'test' }, { ...claims, ...changes }].map(value => Buffer.from(JSON.stringify(value)).toString('base64url'));
    return [...parts, sign('RSA-SHA256', Buffer.from(parts.join('.')), signingKey).toString('base64url')].join('.');
  };
  assert.equal(await validateIdentity(token(), request, fetcher), tenant);
  await assert.rejects(validateIdentity(token({ nonce: 'other' }), request, fetcher), /claims_rejected/);
  await assert.rejects(validateIdentity(token({ exp: 1 }), request, fetcher), /claims_rejected/);
  await assert.rejects(validateIdentity(token({ iss: 'https://evil.test/' }), request, fetcher), /claims_rejected/);
  await assert.rejects(validateIdentity(token(), { ...request, tenant: '22222222-2222-2222-2222-222222222222' }, fetcher), /claims_rejected/);
  const other = generateKeyPairSync('rsa', { modulusLength: 2048 });
  await assert.rejects(validateIdentity(token({}, other.privateKey), request, fetcher), /signature_rejected/);
});

test('capture catches transient fragment callbacks, isolates sessions and removes listener', async () => {
  const request = authorization('skype');
  let listener, removed = false;
  const cdp = { subscribe(fn) { listener = fn; return () => { removed = true; }; } };
  const page = { sessionId: 'auth', async call(method) {
    if (method === 'Page.getFrameTree') return { frameTree: { frame: { id: 'main' } } };
    if (method !== 'Page.navigate') return {};
    listener({ sessionId: 'other', method: 'Page.navigatedWithinDocument', params: { url: `https://teams.microsoft.com/go#state=wrong&access_token=other` } });
    listener({ sessionId: 'auth', method: 'Page.navigatedWithinDocument', params: { frameId: 'main', url: `https://teams.microsoft.com/go#state=${request.state}&access_token=credential` } });
    return {};
  } };
  assert.equal(await capture(page, cdp, request, 1000), 'credential');
  assert.equal(removed, true);
});

function captureFixture(request, events) {
  let listener;
  const calls = [], state = { removed: false };
  const cdp = { subscribe(fn) { listener = fn; return () => { state.removed = true; }; } };
  const page = { sessionId: 'auth', async call(method) {
    calls.push(method);
    if (method === 'Page.getFrameTree') return { frameTree: { frame: { id: 'main' } } };
    if (method === 'Page.navigate') for (const event of events) listener({ sessionId: 'auth', ...event });
    return {};
  } };
  return { result: capture(page, cdp, request, 30), calls, state };
}

test('capture reconstructs a full navigation callback from CDP urlFragment', async () => {
  const request = authorization('teams');
  const fixture = captureFixture(request, [{ method: 'Page.frameNavigated', params: { frame: {
    id: 'main', url: 'https://teams.microsoft.com/go', urlFragment: `#state=${request.state}&id_token=credential`,
  } } }]);
  assert.equal(await fixture.result, 'credential');
  assert.equal(fixture.state.removed, true);
});

test('capture sees the callback request even when /go redirects without committing a document', async () => {
  const request = authorization('skype');
  const fixture = captureFixture(request, [
    { method: 'Network.requestWillBeSent', params: { frameId: 'main', type: 'Document', request: {
      url: 'https://teams.microsoft.com/go', urlFragment: `#state=${request.state}&access_token=credential`,
    } } },
    { method: 'Page.frameNavigated', params: { frame: { id: 'main', url: 'https://teams.cloud.microsoft/v2/' } } },
  ]);
  assert.equal(await fixture.result, 'credential');
  assert.ok(fixture.calls.indexOf('Network.enable') < fixture.calls.indexOf('Page.navigate'));
  assert.equal(fixture.state.removed, true);
});

test('capture ignores iframe and non-document credentials and emits secret-free timeout diagnostics', async () => {
  const request = authorization('skype'), url = `https://teams.microsoft.com/go#state=${request.state}&access_token=secret`;
  const fixture = captureFixture(request, [
    { method: 'Page.navigatedWithinDocument', params: { frameId: 'iframe', url } },
    { method: 'Page.frameNavigated', params: { frame: { id: 'iframe', parentId: 'main', url } } },
    { method: 'Network.requestWillBeSent', params: { frameId: 'main', type: 'Fetch', request: { url } } },
    { method: 'Network.requestWillBeSent', params: { frameId: 'iframe', type: 'Document', request: { url } } },
    { method: 'Page.frameNavigated', params: { frame: { id: 'main', url: 'https://teams.microsoft.com/go' } } },
    { method: 'Page.frameNavigated', params: { frame: { id: 'main', url: 'https://teams.cloud.microsoft/v2/?private=value' } } },
  ]);
  await assert.rejects(fixture.result, error => {
    assert.equal(error.message, 'oauth_timeout');
    assert.deepEqual(error.diagnostics, { callbackSeen: true, callbackHadFragment: false, teamsAppSeen: true });
    assert.ok(!JSON.stringify(error).includes('secret'));
    assert.ok(!JSON.stringify(error).includes('private'));
    return true;
  });
  assert.equal(fixture.state.removed, true);
});

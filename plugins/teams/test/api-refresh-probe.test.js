import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { captureRefreshResponse, acquireBrowserRefresh, renewProbeTokens, checkRefreshCompatibility } from '../lib/refresh-probe.js';
import { clientId, resources } from '../lib/oauth.js';
import { signedToken, signedBundle, tenant, oid } from './token-fixture.js';

const captured = () => ({ clientId, origin: 'https://teams.microsoft.com', refreshToken: 'captured-secret',
  identity: signedBundle().identity, sourceGrant: 'authorization_code' });
function captureFixture({ requests, responses, body, getPostData, exit } = {}) {
  let listener; let removed = false; const calls = [];
  const emit = (method, params, sessionId = 'auth') => listener({ method, params, sessionId });
  const request = (id, changes = {}) => emit('Network.requestWillBeSent', { requestId: id, request: {
    method: 'POST', url: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token?client-request-id=fixture`,
    headers: { Origin: 'https://teams.microsoft.com' },
    postData: new URLSearchParams({ client_id: clientId, grant_type: 'authorization_code', code: 'private-code' }).toString(), ...changes,
  } });
  const response = (id, changes = {}) => {
    emit('Network.responseReceived', { requestId: id, response: { status: 200, url: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token?client-request-id=fixture`, ...changes } });
    emit('Network.loadingFinished', { requestId: id });
  };
  const page = { sessionId: 'auth', async call(method, args) {
    calls.push({ method, args });
    if (method === 'Network.getResponseBody') return body || { body: JSON.stringify({ refresh_token: 'captured-secret', access_token: signedToken('skype') }) };
    if (method === 'Network.getRequestPostData') return getPostData;
    if (method === 'Page.navigate') { (requests || (() => request('token')))({ request, emit }); (responses || (() => response('token')))({ response, emit }); }
    return {};
  } };
  const cdp = { exit, subscribe(fn) { listener = fn; return () => { removed = true; }; } };
  return { page, cdp, calls, removed: () => removed };
}

test('captures only the matching Microsoft Teams exchange, including base64 responses and extra Origin headers', async () => {
  const f = captureFixture({
    requests({ request, emit }) {
      request('other-client', { postData: 'client_id=other&grant_type=authorization_code' });
      request('evil', { url: `https://login.microsoftonline.com.evil.test/${tenant}/oauth2/v2.0/token` });
      emit('Network.requestWillBeSentExtraInfo', { requestId: 'valid', headers: { origin: 'https://teams.cloud.microsoft' } });
      request('valid', { headers: {}, postData: undefined });
    },
    getPostData: { postData: new URLSearchParams({ client_id: clientId, grant_type: 'refresh_token', refresh_token: 'old-secret' }).toString() },
    responses({ response }) { response('other-client'); response('evil'); response('valid'); },
    body: { base64Encoded: true, body: Buffer.from(JSON.stringify({ refresh_token: 'captured-secret', access_token: signedToken('skype') })).toString('base64') },
  });
  const result = await captureRefreshResponse(f.page, f.cdp, { timeout: 1000 });
  assert.equal(result.clientId, clientId); assert.equal(result.refreshToken, 'captured-secret');
  assert.equal(result.identity.oid, oid); assert.equal(result.origin, 'https://teams.cloud.microsoft');
  assert.equal(result.sourceGrant, 'refresh_token'); assert.equal(f.removed(), true);
  assert.deepEqual(f.calls.filter(call => call.method === 'Network.getResponseBody').map(call => call.args.requestId), ['valid']);
});

test('untrusted, redirected and cached responses cannot produce renewal credentials', async () => {
  for (const changes of [{ fromServiceWorker: true }, { fromDiskCache: true }, { status: 400 }, { url: `https://login.microsoftonline.com/common/oauth2/v2.0/token` }]) {
    const f = captureFixture({ responses({ response }) { response('token', changes); } });
    await assert.rejects(captureRefreshResponse(f.page, f.cdp, { timeout: 10 }), { code: 'refresh_capture_timeout' });
    assert.equal(f.removed(), true);
  }
  const f = captureFixture({ requests({ request }) { request('token', { headers: { Origin: 'https://evil.test' } }); } });
  await assert.rejects(captureRefreshResponse(f.page, f.cdp, { timeout: 10 }), { code: 'refresh_capture_timeout' });
});

test('captures authentication worker traffic in its own CDP session without pausing sign-in', async () => {
  let listener; const calls = [], url = `https://login.microsoftonline.com/example.com/oauth2/v2.0/token?client-request-id=fixture`;
  const emit = (method, params, sessionId) => listener({ method, params, sessionId });
  const cdp = {
    subscribe(fn) { listener = fn; return () => {}; },
    async call(method, args, sessionId) {
      calls.push({ method, sessionId }); assert.equal(sessionId, 'worker');
      if (method === 'Network.enable') {
        emit('Network.requestWillBeSent', { requestId: 'request', request: { method: 'POST', url,
          headers: { Origin: 'https://teams.cloud.microsoft' }, postData: `client_id=${clientId}&grant_type=authorization_code` } }, 'worker');
        emit('Network.responseReceived', { requestId: 'request', response: { status: 200, url } }, 'worker');
        emit('Network.loadingFinished', { requestId: 'request' }, 'worker');
      }
      return method === 'Network.getResponseBody' ? { body: JSON.stringify({ refresh_token: 'worker-secret', access_token: signedToken('skype') }) } : {};
    },
  };
  const page = { sessionId: 'page', async call(method) {
    if (method === 'Page.navigate') emit('Target.attachedToTarget', { sessionId: 'worker' }, 'page');
    return {};
  } };
  const result = await captureRefreshResponse(page, cdp, { timeout: 1000 });
  assert.equal(result.refreshToken, 'worker-secret');
  assert.deepEqual(calls.map(call => call.method), ['Network.enable', 'Network.getResponseBody']);
});

test('capture cancellation and invalid bodies clean up without exposing provider data', async () => {
  const f = captureFixture({ body: { body: 'private malformed response' } });
  await assert.rejects(captureRefreshResponse(f.page, f.cdp, { timeout: 1000 }), error => error.code === 'refresh_capture_failed' && !error.message.includes('private'));
  assert.equal(f.removed(), true);
  const controller = new AbortController(); controller.abort();
  const cancelled = captureFixture();
  await assert.rejects(captureRefreshResponse(cancelled.page, cancelled.cdp, { signal: controller.signal }), { code: 'refresh_check_cancelled' });
  assert.equal(cancelled.removed(), true);
});

test('HTTP refresh rotates the token between resources, sends Origin and preserves exact account and expiry', async () => {
  const requests = [], c = captured();
  const result = await renewProbeTokens(c, { fetcher: async (url, options) => {
    requests.push({ url, options }); const service = requests.length === 1 ? 'skype' : 'chatsvcagg';
    return new Response(JSON.stringify({ access_token: signedToken(service), token_type: 'Bearer', expires_in: 3600, refresh_token: 'rotated-secret' }));
  } });
  assert.equal(requests.length, 2); assert.equal(result.authMode, 'tokens');
  assert.deepEqual(result.identity, c.identity);
  for (const [index, service] of ['skype', 'chatsvcagg'].entries()) {
    const { url, options } = requests[index], params = new URLSearchParams(options.body);
    assert.equal(url, `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`);
    assert.equal(options.redirect, 'error'); assert.equal(options.headers.Origin, c.origin);
    assert.equal(params.get('client_id'), clientId); assert.equal(params.get('grant_type'), 'refresh_token');
    assert.equal(params.get('refresh_token'), index ? 'rotated-secret' : c.refreshToken);
    assert.equal(params.get('scope'), `${resources[service]}/.default openid profile offline_access`);
    assert.equal(params.has('client_secret'), false); assert.equal('Cookie' in options.headers, false);
    assert.ok(result.tokens[service].expiresAt > Date.now() + 120_000);
  }
});

test('renewal rejects wrong account, audience, expiry and metadata and sanitizes OAuth errors', async () => {
  for (const changes of [{ oid: '33333333-3333-3333-3333-333333333333' }, { aud: 'https://evil.test' }, { exp: 1 }, { nbf: 'invalid' }, { nbf: Date.now() / 1000 + 600 }]) {
    await assert.rejects(renewProbeTokens(captured(), { fetcher: async () => new Response(JSON.stringify({ access_token: signedToken('skype', changes), token_type: 'Bearer', expires_in: 3600 })) }));
  }
  await assert.rejects(renewProbeTokens({ ...captured(), origin: 'https://evil.test' }, { fetcher: () => { assert.fail('must not send credentials'); } }), { code: 'refresh_metadata_invalid' });
  await assert.rejects(renewProbeTokens(captured(), { fetcher: async () => new Response(JSON.stringify({ error: 'invalid_grant', error_codes: [700084], error_description: 'captured-secret identity-private' }), { status: 400 }) }), error => {
    assert.equal(error.code, 'refresh_exchange_rejected'); assert.match(error.message, /AADSTS700084/);
    assert.ok(!error.message.includes('secret') && !error.message.includes('identity-private')); return true;
  });
});

test('malformed and oversized renewal responses fail without revealing their bodies', async () => {
  for (const body of ['private malformed token response', 'private-secret'.repeat(100_000)]) {
    await assert.rejects(renewProbeTokens(captured(), { fetcher: async () => new Response(body) }), error => {
      assert.equal(error.code, 'refresh_response_invalid'); assert.ok(!error.message.includes('private')); return true;
    });
  }
});

test('browser acquisition closes on capture failure and does not return credentials', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-refresh-failed-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const f = captureFixture({ body: { body: 'private invalid response' } }); let closed = false;
  await assert.rejects(acquireBrowserRefresh(directory, {}, {
    discoverBrowser: async () => 'fixture-browser', createCDP: () => ({ ...f.cdp, page: async () => f.page, close: async () => { closed = true; } }),
  }), { code: 'refresh_capture_failed' });
  assert.equal(closed, true); assert.equal(f.removed(), true);
});

test('browser acquisition uses an isolated profile, clears only Teams app caches and closes before returning', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-refresh-probe-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const f = captureFixture(); let closed = false;
  const result = await acquireBrowserRefresh(directory, { silent: true }, {
    discoverBrowser: async () => 'fixture-browser', createCDP(exe, profile, options) {
      assert.equal(exe, 'fixture-browser'); assert.equal(profile, path.join(directory, 'auth-browser')); assert.equal(options.headless, true);
      return { ...f.cdp, page: async () => f.page, close: async () => { closed = true; } };
    },
  });
  assert.equal(closed, true); assert.equal(result.refreshToken, 'captured-secret');
  assert.equal((await fs.stat(path.join(directory, 'auth-browser'))).mode & 0o777, 0o700);
  const cleared = f.calls.filter(call => call.method === 'Storage.clearDataForOrigin');
  assert.equal(cleared.length, 2); assert.ok(cleared.every(call => !call.args.storageTypes.includes('cookies')));
  assert.deepEqual(await fs.readdir(directory), ['auth-browser']);
});

test('compatibility check renews only after acquisition returns and validates both services with read-only APIs', async () => {
  const events = [], reports = [], bundle = signedBundle();
  const result = await checkRefreshCompatibility({ directory: 'fixture', report: message => reports.push(message) }, {
    acquire: async () => { events.push('browser-closed'); return captured(); },
    renew: async () => { assert.deepEqual(events, ['browser-closed']); events.push('http-renew'); return bundle; },
    exchange: async credentials => { events.push('chat-exchange'); return { ...credentials, chatToken: { expiresAt: Date.now() + 3600_000 } }; },
    createAPI: () => ({ async profile() { events.push('profile'); return { mri: `8:orgid:${oid}` }; }, async chats() { events.push('chats'); return { items: [] }; } }),
  });
  assert.deepEqual(events, ['browser-closed', 'http-renew', 'chat-exchange', 'profile', 'chats']);
  assert.equal(reports.length, 5); assert.ok(result.expiresAt > Date.now());
  assert.ok(!JSON.stringify(reports).includes('captured-secret'));
  await assert.rejects(checkRefreshCompatibility({ directory: 'fixture' }, {
    acquire: async () => captured(), renew: async () => bundle, exchange: async c => c,
    createAPI: () => ({ profile: async () => ({ mri: '8:orgid:other' }), chats: () => assert.fail('must stop before conversation reads') }),
  }), { code: 'account_mismatch' });
});

test('invalid CLI arguments fail without launching a browser or exposing arguments', () => {
  const result = spawnSync(process.execPath, ['scripts/auth-refresh-check.js', '--unknown=private-secret'], { encoding: 'utf8' });
  assert.equal(result.status, 1); assert.ok(!result.stderr.includes('private-secret')); assert.equal(result.stdout, '');
});

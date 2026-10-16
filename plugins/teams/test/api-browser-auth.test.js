import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { generateKeyPairSync, sign } from 'node:crypto';
import { LegacyBrowserAuthenticator as BrowserAuthenticator } from '../lib/api-auth.js';
import { clientId } from '../lib/oauth.js';

const tenant = '11111111-1111-1111-1111-111111111111', oid = '22222222-2222-2222-2222-222222222222';
const { publicKey, privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const key = { ...publicKey.export({ format: 'jwk' }), kid: 'test', use: 'sig' };
const jwt = claims => {
  const encoded = [{ alg: 'RS256', kid: 'test' }, claims].map(value => Buffer.from(JSON.stringify(value)).toString('base64url')).join('.');
  return encoded + '.' + sign('RSA-SHA256', Buffer.from(encoded), privateKey).toString('base64url');
};
async function fixture(t, { error, wrongAccessAccount = false } = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-auth-test-')), launches = [], requests = [];
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const auth = new BrowserAuthenticator(directory, { browserPath: process.execPath }, {
    fetcher: async url => { assert.equal(url, 'https://login.microsoftonline.com/common/discovery/keys'); return new Response(JSON.stringify({ keys: [key] })); },
    createCDP(exe, profile, options) {
      const launch = { exe, profile, options, closed: false }; launches.push(launch); let listener;
      const page = { sessionId: 'auth', async call(method, args) {
        if (method === 'Page.getFrameTree') return { frameTree: { frame: { id: 'main' } } };
        if (method !== 'Page.navigate' || args.url === 'about:blank') return {};
        const url = new URL(args.url), query = url.searchParams; requests.push(query);
        const fragment = new URLSearchParams({ state: query.get('state') });
        if (error) fragment.set('error', error);
        else {
          const isID = query.get('response_type') === 'id_token';
          fragment.set(isID ? 'id_token' : 'access_token', jwt({ tid: tenant, oid: !isID && wrongAccessAccount ? '33333333-3333-3333-3333-333333333333' : oid,
            aud: isID ? clientId : query.get('resource'), nonce: query.get('nonce'), iss: `https://sts.windows.net/${tenant}/`, exp: Math.floor(Date.now() / 1000) + 3600, upn: 'test@example.com', name: 'Test user' }));
          fragment.set('expires_in', '3600');
        }
        listener({ sessionId: 'auth', method: 'Network.requestWillBeSent', params: { frameId: 'main', type: 'Document', request: { url: 'https://teams.microsoft.com/go', urlFragment: '#' + fragment } } });
        return {};
      } };
      return { async page() { return page; }, subscribe(fn) { listener = fn; return () => {}; }, async close() { launch.closed = true; } };
    },
  });
  return { auth, directory, launches, requests };
}

test('interactive OAuth retains a private browser profile, closes Chrome and captures token expiry', async t => {
  const f = await fixture(t), result = await f.auth.acquire({ interactive: true, loginHint: 'test@example.com' });
  assert.equal(result.identity.oid, oid); assert.equal(result.identity.tenant, tenant);
  assert.ok(result.tokens.skype.expiresAt > Date.now());
  assert.equal(result.tokens.teams, undefined, 'ID token is validated then discarded');
  assert.equal(f.launches[0].options.headless, false); assert.equal(f.launches[0].closed, true);
  assert.ok(f.requests.every(query => query.get('prompt') === null && query.get('login_hint') === 'test@example.com'));
  assert.equal((await fs.stat(path.join(f.directory, 'auth-browser'))).mode & 0o777, 0o700);
  assert.deepEqual(await fs.readdir(f.directory), ['auth-browser'], 'no token file is written');
});

test('silent renewal uses the same browser profile and prompt=none and binds the identity', async t => {
  const f = await fixture(t);
  await f.auth.acquire({ previous: { identity: { tenant, oid, email: 'test@example.com' } } });
  assert.equal(f.launches[0].options.headless, true); assert.ok(f.requests.every(query => query.get('prompt') === 'none'));
  await assert.rejects(f.auth.acquire({ interactive: true, previous: { identity: { tenant, oid: '33333333-3333-3333-3333-333333333333' } } }), { code: 'account_mismatch' });
  assert.ok(f.launches.every(launch => launch.closed)); assert.equal(f.launches[0].profile, f.launches[1].profile);
});

test('silent MFA challenges and cross-account access tokens fail without exposing credentials', async t => {
  const challenge = await fixture(t, { error: 'interaction_required' });
  await assert.rejects(challenge.auth.acquire(), { code: 'signin_required' });
  assert.equal(challenge.launches[0].closed, true);
  const mismatch = await fixture(t, { wrongAccessAccount: true });
  await assert.rejects(mismatch.auth.acquire({ interactive: true }), { code: 'account_mismatch' });
  assert.equal(mismatch.launches[0].closed, true);
});

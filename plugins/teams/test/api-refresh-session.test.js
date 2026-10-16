import test from 'node:test';
import assert from 'node:assert/strict';
import { BrowserAuthenticator, TeamsSession } from '../lib/api-auth.js';
import { AdapterError } from '../lib/errors.js';
import { signedBundle, renewal } from './token-fixture.js';

const credentials = () => ({ ...signedBundle(), authMode: 'browser', profile: 'fixture', renewal: renewal(), chatToken: { value: 'chat', expiresAt: Date.now() + 3600_000 } });

test('production browser authenticator captures renewal before HTTP exchange and enforces the existing account', async () => {
  const events = [], c = credentials();
  const auth = new BrowserAuthenticator('fixture-directory', { browserPath: 'fixture-browser' }, {
    acquire: async (dir, options) => {
      assert.equal(dir, 'fixture-directory'); assert.equal(options.silent, false); assert.equal(options.browserPath, 'fixture-browser');
      events.push('browser-closed'); return { ...c.renewal, refreshToken: c.renewal.value, identity: c.identity };
    },
    refresh: async captured => { assert.deepEqual(events, ['browser-closed']); events.push('http'); return { ...c, renewal: { ...c.renewal, value: captured.refreshToken } }; },
  });
  assert.equal((await auth.acquire({ interactive: true, previous: c, loginHint: 'TEST@example.com' })).renewal.value, c.renewal.value);
  events.length = 0;
  await assert.rejects(auth.acquire({ interactive: true, previous: { identity: { ...c.identity, oid: 'other' } } }), { code: 'account_mismatch' });
  assert.deepEqual(events, ['browser-closed']); await auth.close();
});

test('production browser auth sanitizes capture errors and maps silent failure to sign-in guidance', async () => {
  const auth = new BrowserAuthenticator('fixture', {}, { acquire: async () => { throw new Error('private-cookie'); } });
  await assert.rejects(auth.acquire(), error => error.code === 'signin_required' && !error.message.includes('private-cookie'));
  await assert.rejects(auth.acquire({ interactive: true }), { code: 'signin_failed' }); await auth.close();
});

test('browser connections use HTTP first and only reacquire in a browser on terminal refresh rejection', async () => {
  const c = credentials(); c.tokens.skype.expiresAt = 1;
  let browserCalls = 0, refreshCalls = 0, terminal = false;
  const session = new TeamsSession(c.profile, 'fixture', {}, {
    refresh: async () => { refreshCalls++; const e = new AdapterError('refresh_exchange_rejected', 'Renewal failed.', 503); e.reauthenticationRequired = terminal; throw e; },
    exchange: async value => ({ ...value, chatToken: c.chatToken }),
    createAuthenticator: () => ({ acquire: async options => { browserCalls++; assert.equal(options.interactive, false); return { ...signedBundle(), renewal: renewal() }; }, close: async () => {} }),
  });
  session.adopt(c);
  await assert.rejects(session.ensure(), { code: 'refresh_exchange_rejected' }); assert.equal(browserCalls, 0);
  terminal = true; await session.ensure(); assert.equal(browserCalls, 1); assert.equal(refreshCalls, 2);
  assert.equal(session.credentials.authMode, 'browser'); assert.equal(session.status().ready, true); await session.close();
});

test('rotation remains available for persistence when chat-token exchange fails after HTTP renewal', async () => {
  const c = credentials(); c.tokens.skype.expiresAt = 1;
  const session = new TeamsSession(c.profile, 'fixture', {}, {
    refresh: async (_, { onRotate }) => { const r = { ...c.renewal, value: 'rotated' }; await onRotate(r); return { ...signedBundle(), renewal: r }; },
    exchange: async () => { throw new AdapterError('api_transport', 'Exchange failed.', 503); },
    createAuthenticator: () => assert.fail('must not open browser'),
  });
  session.adopt(c); await assert.rejects(session.ensure(), { code: 'api_transport' });
  assert.notEqual(session.credentials, c); assert.equal(session.credentials.renewal.value, 'rotated');
  assert.equal(session.credentials.authMode, 'browser'); assert.equal(session.credentials.profile, c.profile); await session.close();
});

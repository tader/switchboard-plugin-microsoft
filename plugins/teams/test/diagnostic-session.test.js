import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { DiagnosticSession } from '../lib/diagnostic-session.js';
import { signedBundle, renewal } from './token-fixture.js';

async function fixture(t) {
  const directory = await fs.mkdtemp('/tmp/teams-diagnostic-session-');
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  let logins = 0;
  const bundle = () => ({ ...signedBundle(), renewal: renewal() });
  const options = {
    createAuthenticator: () => ({ async acquire() { logins++; return bundle(); }, async close() {} }),
    exchange: async c => ({ ...c, chatToken: { value: 'fixture-chat', expiresAt: Date.now() + 3600_000 }, messageOrigin: 'https://emea.ng.msg.teams.microsoft.com' }),
  };
  return { directory, bundle, options, logins: () => logins };
}

test('live checks reuse one saved diagnostic session across restart without opening another sign-in', async t => {
  const f = await fixture(t), first = new DiagnosticSession(f.directory, {}, f.options);
  await first.acquire(); await first.close();
  assert.equal(f.logins(), 1);
  assert.equal((await fs.stat(first.storage.file)).mode & 0o777, 0o600);
  const second = new DiagnosticSession(f.directory, {}, f.options);
  const result = await second.acquire({ silent: true });
  assert.equal(result.chatToken.value, 'fixture-chat'); assert.equal(f.logins(), 1);
  await second.close();
});

test('cached sessions renew over HTTP and persist rotations before a later exchange fails', async t => {
  const f = await fixture(t), first = new DiagnosticSession(f.directory, {}, f.options);
  await first.acquire(); await first.close();
  const c = JSON.parse(await fs.readFile(first.storage.file, 'utf8')); c.tokens.skype.expiresAt = 1;
  await fs.writeFile(first.storage.file, JSON.stringify(c));
  let refreshes = 0;
  const session = new DiagnosticSession(f.directory, {}, { ...f.options,
    refresh: async (_captured, options) => {
      refreshes++; await options.onRotate({ ...renewal(), value: 'rotated-fixture' });
      assert.equal(JSON.parse(await fs.readFile(first.storage.file, 'utf8')).renewal.value, 'rotated-fixture');
      return { ...f.bundle(), renewal: { ...renewal(), value: 'rotated-fixture' } };
    }, exchange: async () => { throw new Error('network failed'); },
  });
  await assert.rejects(session.acquire({ silent: true }), /network failed/);
  await session.close();
  assert.equal(refreshes, 1); assert.equal(f.logins(), 1);
  assert.equal(JSON.parse(await fs.readFile(first.storage.file, 'utf8')).renewal.value, 'rotated-fixture');
});

test('concurrent checks and account mismatches never start a new sign-in', async t => {
  const f = await fixture(t), first = new DiagnosticSession(f.directory, {}, f.options);
  await first.acquire();
  const second = new DiagnosticSession(f.directory, {}, f.options);
  await assert.rejects(second.acquire(), { code: 'diagnostic_session_busy' }); await second.close();
  await first.close();
  const third = new DiagnosticSession(f.directory, {}, f.options);
  await assert.rejects(third.acquire({ loginHint: 'other@example.com' }), { code: 'account_mismatch' });
  await third.close(); assert.equal(f.logins(), 1);
});

test('fresh capture is explicit and corrupted session data does not silently trigger sign-in', async t => {
  const f = await fixture(t), first = new DiagnosticSession(f.directory, {}, f.options);
  await first.acquire(); await first.close();
  const second = new DiagnosticSession(f.directory, {}, f.options);
  await second.acquire({ capture: true }); await second.close(); assert.equal(f.logins(), 2);
  await fs.writeFile(first.storage.file, 'invalid');
  const third = new DiagnosticSession(f.directory, {}, f.options);
  await assert.rejects(third.acquire(), { code: 'diagnostic_session_invalid' });
  await third.close(); assert.equal(f.logins(), 2);
});

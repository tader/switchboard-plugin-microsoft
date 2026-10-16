import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createServer } from 'node:http';
import { syncTokens } from '../scripts/auth-sync.js';
import { AdapterError } from '../lib/errors.js';
import { signedBundle, renewal } from './token-fixture.js';
import { validateTokenBundle } from '../lib/token-bundle.js';

const options = { url: 'https://switchboard.example/prefix/', connection: 'My Teams', token: 'switchboard-secret', loginHint: 'TEST@example.com' };
const valid = () => ({ transport: 'api', ready: true, authenticating: false, tokenImportRequired: false, signInRequired: false, tokenExpiresAt: new Date(Date.now() + 3600_000).toISOString() });
const expired = () => ({ ...valid(), ready: false, tokenExpiresAt: new Date(1).toISOString() });
const required = () => ({ ...expired(), tokenImportRequired: true, signInRequired: true });
const response = (body, status = 200) => new Response(JSON.stringify(body), { status });
function fixture({ statuses = [valid()], acquireError, httpStatus = 200, reconnectStatus = 200, metadata, signal, cancelOnAcquire = false } = {}) {
  const requests = [], acquisitions = [], reports = []; let closed = 0, statusIndex = 0;
  const deps = {
    report: value => reports.push(value),
    createAuthenticator(directory) {
      assert.match(directory, /token-sync\/[a-f0-9]{24}$/);
      return { async acquire(args) {
        acquisitions.push(args);
        if (cancelOnAcquire) signal.abort();
        if (acquireError && !args.interactive) throw acquireError;
        return { ...signedBundle(), renewal: renewal() };
      }, async close() { closed++; } };
    },
    async fetcher(url, args) {
      requests.push({ url, args });
      assert.equal(args.redirect, 'error'); assert.equal(args.headers.Authorization, 'Bearer switchboard-secret');
      assert.equal(url.origin, 'https://switchboard.example'); assert.ok(url.pathname.startsWith('/prefix/'));
      if (url.pathname.endsWith('/api/connections/My%20Teams')) return response(metadata ?? { id: 'c_test', serviceId: 'teams-web', account: { label: 'Work Teams' } });
      if (url.pathname.endsWith('/proxy/c_test/status')) return response(statuses[Math.min(statusIndex++, statuses.length - 1)]);
      if (url.pathname.endsWith('/proxy/c_test/chats')) return response({ items: [{ text: 'private-chat-content' }] }, httpStatus);
      if (url.pathname.endsWith('/api/connections/c_test/reconnect')) return response({ status: 'connected', connection: { id: 'c_test' }, private: 'remote-private-data' }, reconnectStatus);
      assert.fail('Unexpected request');
    },
  };
  return { deps, requests, acquisitions, reports, closed: () => closed };
}

test('sync skips browser and upload for healthy credentials', async () => {
  const f = fixture(); const result = await syncTokens(options, f.deps);
  assert.equal(result.action, 'unchanged'); assert.equal(f.acquisitions.length, 0); assert.equal(f.requests.length, 2);
});

test('expired access tokens are renewed by the server before trying a browser; chat data is discarded', async () => {
  const f = fixture({ statuses: [expired(), valid()] }); const result = await syncTokens(options, f.deps);
  assert.equal(result.action, 'server_renewed'); assert.equal(f.acquisitions.length, 0);
  assert.deepEqual(f.requests.map(r => r.args.method), ['GET', 'GET', 'GET', 'GET']);
  assert.ok(!JSON.stringify(f.reports).includes('private-chat-content'));
});

test('terminal server renewal failure acquires locally, uploads v2 in memory and verifies status', async () => {
  const f = fixture({ statuses: [expired(), required(), valid()], httpStatus: 502 });
  const result = await syncTokens(options, f.deps);
  assert.equal(result.action, 'uploaded'); assert.equal(f.acquisitions.length, 1); assert.equal(f.acquisitions[0].interactive, false);
  assert.equal(f.acquisitions[0].loginHint, 'test@example.com'); assert.equal(f.closed(), 1);
  const upload = f.requests.find(r => r.url.pathname.endsWith('/reconnect'));
  const payload = JSON.parse(upload.args.body), bundle = JSON.parse(payload.config.tokenBundle);
  assert.equal(payload.method, 'tokens'); assert.equal(payload.config.label, 'Work Teams');
  assert.equal(bundle.version, 2); assert.equal(bundle.renewal.value, 'fixture-refresh-secret');
  assert.ok(!JSON.stringify(f.reports).includes('secret'));
});

test('sync tries silent acquisition first and opens interactive sign-in only when required', async () => {
  const f = fixture({ statuses: [required(), valid()], acquireError: new AdapterError('signin_required', 'Sign in.', 503) });
  assert.equal((await syncTokens(options, f.deps)).action, 'uploaded');
  assert.deepEqual(f.acquisitions.map(a => a.interactive), [false, true]); assert.equal(f.closed(), 2);
  const silent = fixture({ statuses: [required()], acquireError: new AdapterError('signin_required', 'Sign in.', 503) });
  await assert.rejects(syncTokens({ ...options, silent: true }, silent.deps), { code: 'signin_required' });
  assert.deepEqual(silent.acquisitions.map(a => a.interactive), [false]); assert.equal(silent.closed(), 1);
  assert.ok(!silent.requests.some(r => r.url.pathname.endsWith('/reconnect')));
});

test('force reacquires even healthy tokens, but pending server sign-in is never replaced', async () => {
  const f = fixture(); assert.equal((await syncTokens({ ...options, force: true }, f.deps)).action, 'uploaded'); assert.equal(f.acquisitions.length, 1);
  const busy = fixture({ statuses: [{ ...valid(), authenticating: true }] });
  await assert.rejects(syncTokens({ ...options, force: true }, busy.deps), { code: 'sync_signin_pending' }); assert.equal(busy.acquisitions.length, 0);
});

test('transient server errors do not launch a browser or upload replacements', async () => {
  const f = fixture({ statuses: [expired(), { ...expired(), error: { code: 'refresh_transport_failed' } }], httpStatus: 502 });
  await assert.rejects(syncTokens(options, f.deps), { code: 'sync_renewal_failed' });
  assert.equal(f.acquisitions.length, 0); assert.ok(!f.requests.some(r => r.url.pathname.endsWith('/reconnect')));
});

test('bad input and non-Teams connections fail before browser acquisition or credential upload', async () => {
  for (const changes of [{ url: 'http://remote.example' }, { url: 'https://user:password@example.com' }, { url: 'https://example.com/?secret=x' }, { connection: '' }, { token: '' }]) {
    const f = fixture(); await assert.rejects(syncTokens({ ...options, ...changes }, f.deps)); assert.equal(f.requests.length, 0);
  }
  const f = fixture({ metadata: { id: 'c_test', serviceId: 'other-service' } });
  await assert.rejects(syncTokens(options, f.deps), { code: 'sync_connection_invalid' }); assert.equal(f.acquisitions.length, 0);
  const malformed = fixture({ statuses: [{ ...valid(), tokenExpiresAt: 'bad-expiry' }] });
  await assert.rejects(syncTokens(options, malformed.deps), { code: 'sync_status_invalid' });
});

test('redirects, transport failures and denied requests are sanitized and not retried', async () => {
  for (const status of [401, 403, 500]) {
    let calls = 0;
    await assert.rejects(syncTokens(options, { fetcher: async () => { calls++; return response({ error: 'private-provider-secret' }, status); }, createAuthenticator: () => assert.fail('no browser') }), error => {
      assert.ok(!error.message.includes('private')); return true;
    }); assert.equal(calls, 1);
  }
  await assert.rejects(syncTokens(options, { fetcher: async () => { throw new Error('private network secret'); } }), error => error.code === 'sync_transport_failed' && !error.message.includes('private'));
});

test('upload rejection and post-upload failure stop without duplicate reconnect requests', async () => {
  const denied = fixture({ statuses: [required()], reconnectStatus: 403 });
  await assert.rejects(syncTokens(options, denied.deps), { code: 'sync_access_denied' });
  assert.equal(denied.requests.filter(r => r.url.pathname.endsWith('/reconnect')).length, 1);
  const failed = fixture({ statuses: [required(), required()] });
  await assert.rejects(syncTokens(options, failed.deps), { code: 'sync_verification_failed' });
  assert.equal(failed.requests.filter(r => r.url.pathname.endsWith('/reconnect')).length, 1);
});

test('cancellation during acquisition closes auth and prevents upload', async () => {
  const controller = new AbortController(), f = fixture({ statuses: [required()], signal: controller, cancelOnAcquire: true });
  await assert.rejects(syncTokens({ ...options, signal: controller.signal }, f.deps), { name: 'AbortError' });
  assert.ok(f.closed() >= 1); assert.ok(!f.requests.some(r => r.url.pathname.endsWith('/reconnect')));
});

test('CLI help needs no credentials and malformed arguments never echo secrets', () => {
  const help = spawnSync(process.execPath, ['scripts/auth-sync.js', '--help'], { encoding: 'utf8', cwd: new URL('..', import.meta.url) });
  assert.equal(help.status, 0); assert.match(help.stdout, /SWITCHBOARD_TOKEN/);
  const invalid = spawnSync(process.execPath, ['scripts/auth-sync.js', '--token=private-secret'], { encoding: 'utf8', cwd: new URL('..', import.meta.url) });
  assert.equal(invalid.status, 1); assert.ok(!invalid.stderr.includes('private-secret'));
});

test('sync uses real HTTP for status and reconnect without a plaintext export', async t => {
  let uploaded = false, calls = 0;
  const server = createServer(async (req, res) => {
    calls++; assert.equal(req.headers.authorization, 'Bearer local-test-token');
    const send = value => { res.setHeader('content-type', 'application/json'); res.end(JSON.stringify(value)); };
    if (req.url === '/api/connections/work') return send({ id: 'c_live_fixture', serviceId: 'teams-web' });
    if (req.url === '/proxy/c_live_fixture/status') return send(uploaded ? valid() : required());
    if (req.url === '/api/connections/c_live_fixture/reconnect') {
      assert.equal(req.method, 'POST'); const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const body = JSON.parse(Buffer.concat(chunks).toString());
      assert.equal(body.method, 'tokens');
      assert.equal((await validateTokenBundle(body.config.tokenBundle)).renewal.value, 'fixture-refresh-secret');
      uploaded = true; return send({ status: 'connected', connection: { id: 'c_live_fixture' } });
    }
    res.writeHead(404); res.end();
  });
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
  t.after(() => new Promise(resolve => server.close(resolve)));
  const result = await syncTokens({ url: `http://127.0.0.1:${server.address().port}`, connection: 'work', token: 'local-test-token' }, {
    createAuthenticator: () => ({ acquire: async () => ({ ...signedBundle(), renewal: renewal() }), close: async () => {} }),
  });
  assert.equal(result.action, 'uploaded'); assert.equal(calls, 4);
});

test('real HTTP redirects are refused before the token can reach a redirect target', async t => {
  let targetReached = false;
  const server = createServer((req, res) => {
    if (req.url.startsWith('/api/connections/')) { res.writeHead(302, { location: '/redirect-target' }); res.end(); }
    else { targetReached = true; res.end('{}'); }
  });
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
  t.after(() => new Promise(resolve => server.close(resolve)));
  await assert.rejects(syncTokens({ url: `http://127.0.0.1:${server.address().port}`, connection: 'work', token: 'local-test-token' }), { code: 'sync_transport_failed' });
  assert.equal(targetReached, false);
});

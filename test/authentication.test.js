import { test } from 'node:test';
import setupAuthentication from '../plugins/microsoft-authentication/index.ts';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { FORMAT, CLIENT_ID, DiagnosticSession, refresh, validateBundle } from '../plugins/microsoft-authentication/lib/session.js';
import { createConsumerAuth, coversScopes, allowRequest } from '../plugins/microsoft-authentication/consumer.ts';
import { outlookSpec } from '../plugins/microsoft-authentication/web-spec.ts';

const tenant = '11111111-1111-4111-8111-111111111111', oid = '22222222-2222-4222-8222-222222222222';
const graph = 'https://graph.microsoft.com', outlook = 'https://outlook.office.com';
const bundle = () => ({ format: FORMAT, clientId: CLIENT_ID, origin: 'https://outlook.cloud.microsoft', identity: { tenant, oid },
  refreshToken: 'fixture-refresh', accessToken: 'fixture-access', expiresAt: Date.now() + 3600_000,
  scope: `${graph}/User.Read ${graph}/Files.ReadWrite.All`, generation: '33333333-3333-4333-8333-333333333333' });
const response = (value, status = 200) => new Response(JSON.stringify(value), { status });

test('cached diagnostics reuse credentials across runs without opening a browser', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'microsoft-cache-'));
  try {
    const a = new DiagnosticSession(dir, { capture: () => { throw new Error('Unexpected browser'); } });
    await a.save(bundle());
    assert.equal((await a.acquire()).accessToken, 'fixture-access'); await a.close();
    const b = new DiagnosticSession(dir, { capture: () => { throw new Error('Unexpected browser'); } });
    assert.equal((await b.acquire({ silent: true })).refreshToken, 'fixture-refresh');
    assert.equal((await fs.stat(b.file)).mode & 0o777, 0o600); await b.close();
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
});

test('HTTP renewal persists a rotation before a subsequent identity-check failure', async () => {
  let saved;
  await assert.rejects(refresh(bundle(), { onRotate: async value => { saved = value; }, fetcher: async url => {
    if (url.includes('/token')) return response({ access_token: 'opaque-graph', refresh_token: 'rotated-refresh', token_type: 'Bearer', expires_in: 3600, scope: `${graph}/User.Read` });
    return response({}, 503);
  } }), { code: 'identity_unavailable' });
  assert.equal(saved.refreshToken, 'rotated-refresh');
  assert.equal(saved.accessToken, 'fixture-access');
});

test('Outlook resource renewal uses its own audience and preserves the Graph token', async () => {
  const updated = await refresh(bundle(), { resource: 'outlook', fetcher: async (url, options) => {
    assert.equal(new URL(url).host, 'login.microsoftonline.com');
    assert.equal(options.headers.Origin, 'https://outlook.cloud.microsoft');
    assert.equal(options.body.get('scope'), `${outlook}/.default openid profile offline_access`);
    assert.equal(options.redirect, 'error');
    return response({ access_token: 'opaque-outlook', refresh_token: 'rotated', token_type: 'Bearer', expires_in: 3600, scope: `${outlook}/Mail.ReadWrite ${outlook}/Mail.Send` });
  } });
  assert.equal(updated.accessToken, 'fixture-access');
  assert.equal(updated.resources.outlook.accessToken, 'opaque-outlook');
  assert.equal(updated.refreshToken, 'rotated');
});

test('terminal renewal rejection requests authentication; throttling preserves the cache', async () => {
  for (const [status, error, expected] of [[400, 'invalid_grant', true], [429, 'temporarily_unavailable', false], [400, 'consent_required', false]]) {
    await assert.rejects(refresh(bundle(), { fetcher: async () => response({ error, error_description: 'fixture-secret-must-not-appear', error_codes: [700084] }, status) }), value => {
      assert.equal(value.reauthenticationRequired, expected);
      assert.ok(!value.message.includes('fixture-secret')); return true;
    });
  }
});

test('imports cannot switch accounts, clients or origins, and profiles never cross the boundary', () => {
  const original = bundle();
  assert.throws(() => validateBundle({ ...original, clientId: 'other' }), { code: 'invalid_credentials' });
  assert.throws(() => validateBundle({ ...original, origin: 'https://evil.example' }), { code: 'invalid_credentials' });
  assert.throws(() => validateBundle({ ...original, identity: { ...original.identity, oid: tenant } }, original), { code: 'account_mismatch' });
  assert.equal(validateBundle({ ...original, profile: '/private/profile', cookies: ['secret'] }).profile, undefined);
});

test('permission implication never transfers a permission between Graph and Outlook audiences', () => {
  assert.ok(coversScopes([`${outlook}/Mail.ReadWrite`], [`${outlook}/Mail.Read`]));
  assert.ok(!coversScopes([`${outlook}/Mail.ReadWrite`], [`${graph}/Mail.Read`]));
  assert.ok(coversScopes([`${graph}/Files.ReadWrite.All`], [`${graph}/Files.Read`]));
  assert.ok(!coversScopes([`${graph}/Files.ReadWrite`], [`${graph}/Files.ReadWrite.All`]));
});

test('mail, file, read-only and generic Graph boundaries block use of broader grants', () => {
  const conn = { methodId: 'authentication-source', config: { backend: 'outlook', access: 'read' } };
  const req = (method, p, origin = outlook + '/api/v2.0') => ({ method, url: new URL(origin + p) });
  assert.ok(allowRequest('outlook-mail', req('GET', '/me/messages'), conn, []));
  assert.ok(!allowRequest('outlook-mail', req('POST', '/me/messages'), conn, []));
  assert.ok(!allowRequest('outlook-mail', req('GET', '/me/tasks'), conn, []));
  conn.config.access = 'send';
  assert.ok(allowRequest('outlook-mail', req('POST', '/me/sendmail'), conn, []));
  assert.ok(!allowRequest('outlook-mail', req('PATCH', '/me/messages/id'), conn, []));
  conn.config = { backend: 'graph', access: 'full' };
  assert.ok(!allowRequest('onedrive', req('GET', '/sites/root', graph + '/v1.0'), conn, []));
  assert.ok(allowRequest('onedrive', req('GET', '/me/drive/root', graph + '/v1.0'), conn, []));
  assert.ok(!allowRequest('microsoft-graph', req('GET', '/users', graph + '/v1.0'), conn, [`${graph}/Files.ReadWrite.All`]));
  assert.ok(!allowRequest('microsoft-graph', req('POST', '/$batch', graph + '/v1.0'), conn, [`${graph}/Directory.ReadWrite.All`]));
});

test('source-backed mail selects a native Outlook backend without requesting Graph consent', async () => {
  const original = globalThis.fetch;
  let exchanges = 0;
  globalThis.fetch = async (url, options) => {
    if (String(url).includes('/token')) {
      exchanges++;
      assert.ok(options.body.get('scope').startsWith(outlook));
      return response({ access_token: 'outlook-access', token_type: 'Bearer', expires_in: 3600, scope: `${outlook}/Mail.ReadWrite ${outlook}/Mail.Send` });
    }
    return response({ id: oid, mail: 'fixture@example.com' });
  };
  try {
    const method = createConsumerAuth({ credentials: { obtain: async () => ({ format: FORMAT, bundle: bundle(), account: { id: `${tenant}:${oid}`, label: 'Fixture' } }) } }, { id: 'outlook-mail', scopes: ['Mail.ReadWrite', 'Mail.Send'] });
    const result = await method.connect({ config: { authenticationSource: 'source', access: 'full' } });
    assert.equal(result.config.backend, 'outlook'); assert.equal(exchanges, 1);
    const request = { method: 'GET', url: new URL(`${outlook}/api/v2.0/me/messages`), headers: new Headers() };
    await method.authorize(request, { methodId: method.id, config: result.config, credentials: result.credentials }, {});
    assert.equal(request.headers.get('authorization'), 'Bearer outlook-access'); assert.equal(exchanges, 1);
  } finally { globalThis.fetch = original; }
});

test('native Outlook descriptions use their actual JSON dialect and never replace Graph descriptions', () => {
  const mail = outlookSpec('outlook-mail');
  assert.equal(mail.paths['/me/sendmail'].post.requestBody.content['application/json'].example.Message.Subject, 'Hello');
  assert.ok(outlookSpec('microsoft-todo').paths['/me/tasks/{task-id}/complete']);
  assert.ok(!mail.paths['/me/todo/lists']);
});


test('authentication connection tests probe identity without exposing product APIs', async () => {
  const plugin = setupAuthentication({ dataDir: '/unused', settings: {} });
  const service = plugin.services[0], method = service.authMethods[0];
  const credentials = { profile: 'fixture-profile', bundle: bundle() };
  const conn = { credentials, config: {} };
  try {
    for (const verb of ['GET', 'HEAD']) for (const suffix of ['', '/', '?$expand=messages']) {
      const req = { method: verb, url: new URL(service.baseUrl + suffix), headers: new Headers() };
      await method.authorize(req, conn, {});
      assert.equal(req.url.origin, graph);
      assert.equal(req.url.pathname, '/v1.0/me');
      assert.equal(req.url.searchParams.get('$select'), 'id,displayName,mail,userPrincipalName');
      assert.equal(req.url.searchParams.has('$expand'), false);
      assert.equal(req.headers.get('authorization'), 'Bearer fixture-access');
    }
    for (const [verb, url] of [['POST', graph + '/v1.0'], ['GET', graph + '/v1.0/me/messages'], ['GET', graph + '/v1.0/users'], ['GET', 'https://evil.example/v1.0']]) {
      const req = { method: verb, url: new URL(url), headers: new Headers() };
      await assert.rejects(method.authorize(req, conn, {}), { code: 'request_forbidden' });
      assert.equal(req.headers.has('authorization'), false);
    }
  } finally { plugin.dispose(); }
});


test('expired Microsoft consumers reuse a reconnected source before forcing browser authentication', async () => {
  const original = globalThis.fetch;
  try {
    for (const backend of ['graph', 'outlook']) {
      const stale = { ...bundle(), refreshToken: 'stale-refresh', expiresAt: 0 };
      const current = bundle(), obtained = [], saved = [];
      let exchanges = 0;
      globalThis.fetch = async (url, options) => {
        if (String(url).includes('/token')) {
          exchanges++;
          if (options.body.get('refresh_token') === 'stale-refresh') return response({ error: 'invalid_grant' }, 400);
          assert.equal(options.body.get('refresh_token'), current.refreshToken);
          assert.ok(options.body.get('scope').startsWith(outlook));
          return response({ access_token: 'new-outlook-access', token_type: 'Bearer', expires_in: 3600, scope: `${outlook}/Mail.ReadWrite` });
        }
        return response({ id: oid, mail: 'fixture@example.test' });
      };
      const method = createConsumerAuth({ credentials: { obtain: async (ref, opts) => {
        obtained.push({ ref, ...opts });
        assert.equal(opts.forceAuthentication, false, 'the source already has a valid reconnected session');
        return { format: FORMAT, bundle: current, account: { id: `${tenant}:${oid}`, label: 'Fixture' } };
      } } }, { id: backend === 'graph' ? 'onedrive' : 'outlook-mail', scopes: backend === 'graph' ? ['Files.ReadWrite.All'] : ['Mail.ReadWrite'] });
      const result = await method.token({ methodId: 'authentication-source', credentials: stale, config: { authenticationSource: 'laptop', backend } }, { persistCredentials: value => { saved.push(value); } });
      assert.equal(result.accessToken, backend === 'graph' ? current.accessToken : 'new-outlook-access');
      assert.equal(obtained.length, 1);
      assert.equal(obtained[0].ref, 'laptop');
      assert.equal(exchanges, backend === 'graph' ? 1 : 2);
      assert.equal(saved[0].refreshToken, current.refreshToken, 'retain the reconnected session before subsequent renewal');
    }
  } finally { globalThis.fetch = original; }
});

test('Outlook forces authentication only after current source renewal also requires sign-in', async () => {
  const original = globalThis.fetch;
  try {
    for (const retryError of ['invalid_grant', 'temporarily_unavailable']) {
      const obtained = [], saved = [];
      globalThis.fetch = async (url, options) => {
        if (!String(url).includes('/token')) return response({ id: oid });
        const rt = options.body.get('refresh_token');
        if (rt === 'stale-refresh') return response({ error: 'invalid_grant' }, 400);
        if (rt === 'fixture-refresh') return response({ error: retryError }, retryError === 'invalid_grant' ? 400 : 429);
        assert.equal(rt, 'reauthenticated-refresh');
        return response({ access_token: 'outlook-access', token_type: 'Bearer', expires_in: 3600, scope: `${outlook}/Mail.ReadWrite` });
      };
      const method = createConsumerAuth({ credentials: { obtain: async (_ref, opts) => {
        obtained.push(opts.forceAuthentication);
        return { format: FORMAT, bundle: { ...bundle(), refreshToken: opts.forceAuthentication ? 'reauthenticated-refresh' : 'fixture-refresh' }, account: { id: `${tenant}:${oid}` } };
      } } }, { id: 'outlook-mail', scopes: ['Mail.ReadWrite'] });
      const request = method.token({ methodId: 'authentication-source', credentials: { ...bundle(), refreshToken: 'stale-refresh' }, config: { authenticationSource: 'laptop', backend: 'outlook' } }, { persistCredentials: value => { saved.push(value); } });
      if (retryError === 'invalid_grant') {
        assert.equal((await request).accessToken, 'outlook-access');
        assert.deepEqual(obtained, [false, true]);
        assert.equal(saved.at(-1).refreshToken, 'reauthenticated-refresh');
      } else {
        await assert.rejects(request, { code: 'refresh_rate_limited' });
        assert.deepEqual(obtained, [false], 'transient errors never request browser authentication');
      }
    }
  } finally { globalThis.fetch = original; }
});

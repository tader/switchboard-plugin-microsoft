import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { encodeTokenBundle, validateTokenBundle } from '../lib/token-bundle.js';
import { exportTokens } from '../scripts/auth-export.js';
import { signedBundle, signedToken, oid, renewal } from './token-fixture.js';

test('portable bundles retain only supported secrets and verified account identity', async () => {
  const source = { ...signedBundle(), profile: 'local-only', messageOrigin: 'https://evil.example', chatToken: { value: 'derived-secret' } };
  const raw = encodeTokenBundle(source), decoded = JSON.parse(raw);
  assert.deepEqual(Object.keys(decoded), ['format', 'version', 'identity', 'tokens']);
  assert.ok(!raw.includes('local-only') && !raw.includes('evil.example') && !raw.includes('derived-secret'));
  const imported = await validateTokenBundle(raw, source);
  assert.equal(imported.authMode, 'tokens'); assert.equal(imported.identity.oid, oid);
  assert.ok(imported.tokens.skype.expiresAt <= source.tokens.skype.expiresAt);
});

test('imports reject malformed, oversize, expired and mismatched bundles without leaking secrets', async () => {
  const valid = JSON.parse(encodeTokenBundle(signedBundle()));
  const invalidCases = ['not json', 'x'.repeat(128 * 1024 + 1), JSON.stringify({ ...valid, version: 2 }), JSON.stringify({ ...valid, profile: 'injected' })];
  for (const raw of invalidCases) await assert.rejects(validateTokenBundle(raw, undefined), { code: 'invalid_token_bundle' });
  for (const changes of [{ aud: 'wrong-service' }, { nbf: Date.now() / 1000 + 600 }]) {
    const bundle = signedBundle(); bundle.tokens.skype.value = signedToken('skype', changes);
    await assert.rejects(validateTokenBundle(encodeTokenBundle(bundle), undefined), { code: 'invalid_token_bundle' });
  }
  const expired = signedBundle(); expired.tokens.skype.value = signedToken('skype', { exp: Date.now() / 1000 + 60 });
  await assert.rejects(validateTokenBundle(encodeTokenBundle(expired), undefined), { code: 'token_import_required' });
  const declaredExpired = structuredClone(valid); declaredExpired.tokens.skype.expiresAt = 1;
  await assert.rejects(validateTokenBundle(JSON.stringify(declaredExpired), undefined), { code: 'token_import_required' });
  const mixed = signedBundle(); mixed.tokens.chatsvcagg.value = signedToken('chatsvcagg', { oid: '33333333-3333-3333-3333-333333333333' });
  await assert.rejects(validateTokenBundle(encodeTokenBundle(mixed), undefined), { code: 'account_mismatch' });
  const tenantMixed = signedBundle(); tenantMixed.tokens.skype.value = signedToken('skype', { tid: '33333333-3333-3333-3333-333333333333' });
  await assert.rejects(validateTokenBundle(encodeTokenBundle(tenantMixed), undefined), { code: 'invalid_token_bundle' });
  await assert.rejects(validateTokenBundle(JSON.stringify(valid), { identity: { ...valid.identity, oid: '33333333-3333-3333-3333-333333333333' } }), { code: 'account_mismatch' });
});

test('export creates an owner-only file, supports silent reuse, closes auth and refuses overwrite', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-export-test-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const out = path.join(directory, 'tokens.json'); let closed = 0, acquired = 0;
  const createAuthenticator = () => ({ async acquire(options) { acquired++; assert.equal(options.interactive, false); assert.equal(options.loginHint, 'test@example.com'); return { ...signedBundle(), renewal: renewal() }; }, async close() { closed++; } });
  await exportTokens({ out, silent: true, loginHint: 'TEST@example.com' }, { createAuthenticator });
  assert.equal((await fs.stat(out)).mode & 0o777, 0o600); assert.equal(closed, 1);
  assert.equal((await validateTokenBundle(await fs.readFile(out, 'utf8'), undefined)).identity.oid, oid);
  const contents = await fs.readFile(out, 'utf8');
  assert.equal(JSON.parse(contents).version, 2);
  await assert.rejects(exportTokens({ out }, { createAuthenticator }), { code: 'export_file_unavailable' });
  assert.equal(await fs.readFile(out, 'utf8'), contents); assert.equal(acquired, 1);
  const link = path.join(directory, 'link.json'); await fs.symlink(out, link);
  await assert.rejects(exportTokens({ out: link }, { createAuthenticator }), { code: 'export_file_unavailable' });
});

test('failed and interrupted exports remove incomplete files and close the browser', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'teams-export-fail-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const out = path.join(directory, 'tokens.json'); let closed = false;
  await assert.rejects(exportTokens({ out }, { createAuthenticator: () => ({ async acquire() { throw new Error('private provider response'); }, async close() { closed = true; } }) }));
  assert.equal(closed, true); assert.deepEqual(await fs.readdir(directory), []);
  const controller = new AbortController(); closed = false;
  await assert.rejects(exportTokens({ out, signal: controller.signal }, { createAuthenticator: () => ({ async acquire() { controller.abort(); return signedBundle(); }, async close() { closed = true; } }) }), { name: 'AbortError' });
  assert.equal(closed, true); assert.deepEqual(await fs.readdir(directory), []);
  const result = spawnSync(process.execPath, ['scripts/auth-export.js'], { encoding: 'utf8', cwd: new URL('..', import.meta.url) });
  assert.equal(result.status, 1); assert.match(result.stderr, /export_path_required/);
  assert.ok(!result.stderr.includes('private provider response'));
});

test('v2 portable bundles retain renewal metadata and permit expired access tokens for online refresh validation', async () => {
  const c = { ...signedBundle({ exp: 1 }), renewal: renewal() };
  const raw = encodeTokenBundle(c), decoded = JSON.parse(raw);
  assert.equal(decoded.version, 2); assert.deepEqual(decoded.renewal, c.renewal);
  assert.equal((await validateTokenBundle(raw)).renewal.value, c.renewal.value);
  for (const changes of [{ clientId: 'wrong' }, { origin: 'https://evil.test' }, { capturedAt: Date.now() + 3600_000 }, { sourceGrant: 'password' }, { value: '' }, { cookie: 'injected' }]) {
    await assert.rejects(validateTokenBundle(JSON.stringify({ ...decoded, renewal: { ...decoded.renewal, ...changes } })), { code: 'invalid_token_bundle' });
  }
  await assert.rejects(validateTokenBundle(JSON.stringify({ ...decoded, version: 1 })), { code: 'invalid_token_bundle' });
});

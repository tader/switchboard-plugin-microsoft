import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { AdapterError } from './errors.js';

export const FORMAT = 'microsoft/web-session/v1';
export const CLIENT_ID = '9199bf20-a13f-4107-85dc-02114787ef48';
export const ORIGINS = ['https://outlook.cloud.microsoft', 'https://outlook.office.com'];
export const TEAMS_CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
export const TEAMS_ORIGINS = ['https://teams.microsoft.com', 'https://teams.cloud.microsoft'];
export const validClientOrigin = (client, origin) => client === CLIENT_ID && ORIGINS.includes(origin) || client === TEAMS_CLIENT_ID && TEAMS_ORIGINS.includes(origin);
export const RESOURCES = { graph: 'https://graph.microsoft.com', outlook: 'https://outlook.office.com', skype: 'https://api.spaces.skype.com', chatsvcagg: 'https://chatsvcagg.teams.microsoft.com', substrate: 'https://substrate.office.com' };
export const GRAPH = 'https://graph.microsoft.com/v1.0';
const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const secret = value => typeof value === 'string' && value.length > 0 && value.length <= 100_000;
export const fail = (code, message, status = 503) => new AdapterError(code, message, status);
export const fresh = bundle => secret(bundle?.accessToken) && Number.isFinite(bundle.expiresAt) && bundle.expiresAt > Date.now() + 120_000;
export function claims(token) {
  try { return JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString()); } catch { return {}; }
}

export function validateBundle(bundle, previous) {
  if (!bundle || bundle.format !== FORMAT || !uuid.test(bundle.identity?.tenant ?? '') || !uuid.test(bundle.identity?.oid ?? '') ||
      !validClientOrigin(bundle.clientId, bundle.origin) || !secret(bundle.refreshToken) ||
      !secret(bundle.accessToken) || !Number.isFinite(bundle.expiresAt) || typeof bundle.scope !== 'string' || bundle.scope.length > 32_000 ||
      typeof bundle.generation !== 'string' || !uuid.test(bundle.generation)) throw fail('invalid_credentials', 'Invalid Microsoft authentication bundle.', 400);
  if (previous && (previous.identity.tenant !== bundle.identity.tenant || previous.identity.oid !== bundle.identity.oid)) {
    throw fail('account_mismatch', 'The authentication source belongs to a different Microsoft account.', 403);
  }
  // Only explicitly supported fields cross the authentication boundary. Profiles and cookies never do.
  const resources = {};
  for (const key of Object.keys(RESOURCES).filter(key => key !== 'graph')) {
    const token = bundle.resources?.[key];
    if (!token) continue;
    if (!secret(token.accessToken) || !Number.isFinite(token.expiresAt) || typeof token.scope !== 'string' || token.scope.length > 32_000) throw fail('invalid_credentials', 'Invalid Microsoft resource credentials.', 400);
    resources[key] = { accessToken: token.accessToken, expiresAt: token.expiresAt, scope: token.scope };
  }
  return { resources, format: FORMAT, clientId: bundle.clientId, origin: bundle.origin, identity: { tenant: bundle.identity.tenant, oid: bundle.identity.oid },
    refreshToken: bundle.refreshToken, accessToken: bundle.accessToken, expiresAt: bundle.expiresAt, scope: bundle.scope, generation: bundle.generation };
}

export async function jsonResponse(response) {
  if (!response.body) throw fail('invalid_response', 'Microsoft returned an empty response.');
  const reader = response.body.getReader(); const chunks = []; let size = 0;
  try {
    for (;;) {
      const chunk = await reader.read(); if (chunk.done) break;
      size += chunk.value.length;
      if (size > 512 * 1024) throw fail('invalid_response', 'Microsoft response exceeded the size limit.');
      chunks.push(Buffer.from(chunk.value));
    }
    return JSON.parse(Buffer.concat(chunks).toString());
  } catch (error) {
    await reader.cancel().catch(() => {});
    throw error instanceof AdapterError ? error : fail('invalid_response', 'Microsoft returned an invalid response.');
  } finally { reader.releaseLock(); }
}

export async function identify(bundle, { fetcher = fetch, signal } = {}) {
  const response = await fetcher(`${GRAPH}/me?$select=id,displayName,mail,userPrincipalName`, {
    headers: { authorization: `Bearer ${bundle.accessToken}` }, redirect: 'error',
    signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(20_000)]) : AbortSignal.timeout(20_000),
  }).catch(() => { throw fail('api_transport_failed', 'Microsoft Graph identity check failed or timed out.'); });
  if (!response.ok) {
    await response.body?.cancel();
    throw fail(response.status === 401 ? 'api_unauthorized' : 'identity_unavailable', `Microsoft Graph rejected the identity check (HTTP ${response.status}).`, response.status === 401 ? 401 : 503);
  }
  const me = await jsonResponse(response);
  if (me.id !== bundle.identity.oid) throw fail('account_mismatch', 'Microsoft Graph authenticated a different account.', 403);
  return { id: `${bundle.identity.tenant}:${me.id}`, label: me.mail || me.userPrincipalName || me.displayName || me.id };
}

export async function refresh(bundle, { fetcher = fetch, signal, onRotate = async () => {}, resource = 'graph' } = {}) {
  if (!Object.hasOwn(RESOURCES, resource)) throw fail('invalid_resource', 'Unsupported Microsoft API resource.', 400);
  // Captures have no access token yet. Their metadata came from a filtered Microsoft HTTPS response.
  if (!validClientOrigin(bundle.clientId, bundle.origin) || !uuid.test(bundle.identity?.tenant ?? '') || !uuid.test(bundle.identity?.oid ?? '') || !secret(bundle.refreshToken)) {
    throw fail('invalid_credentials', 'Invalid Microsoft renewal metadata.', 400);
  }
  const response = await fetcher(`https://login.microsoftonline.com/${bundle.identity.tenant}/oauth2/v2.0/token`, {
    method: 'POST', redirect: 'error', headers: { 'content-type': 'application/x-www-form-urlencoded', Origin: bundle.origin },
    body: new URLSearchParams({ client_id: bundle.clientId, grant_type: 'refresh_token', refresh_token: bundle.refreshToken,
      scope: `${RESOURCES[resource]}/.default openid profile offline_access` }),
    signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(20_000)]) : AbortSignal.timeout(20_000),
  }).catch(() => { throw fail('refresh_transport_failed', 'Microsoft token renewal failed or timed out.'); });
  const value = await jsonResponse(response);
  if (!response.ok || value.error) {
    const aadsts = value.error_codes?.find?.(code => Number.isSafeInteger(code) && code > 0);
    const error = fail(response.status === 429 ? 'refresh_rate_limited' : 'refresh_rejected', `Microsoft rejected token renewal${aadsts ? ` (AADSTS${aadsts})` : ` (HTTP ${response.status})`}.`);
    error.reauthenticationRequired = ['invalid_grant', 'interaction_required', 'login_required'].includes(value.error);
    throw error;
  }
  const seconds = Number(value.expires_in);
  if (!secret(value.access_token) || value.token_type?.toLowerCase() !== 'bearer' || !Number.isFinite(seconds) || seconds <= 120 || seconds > 604800 ||
      typeof value.scope !== 'string' || value.scope.length > 32_000 || value.refresh_token !== undefined && !secret(value.refresh_token)) {
    throw fail('invalid_response', 'Microsoft returned invalid renewal credentials.');
  }
  const updated = { ...bundle, refreshToken: value.refresh_token ?? bundle.refreshToken };
  // Retain rotations even if a subsequent Graph check fails. Never persist an unvalidated initial import.
  if (bundle.format === FORMAT && value.refresh_token) await onRotate(updated);
  const token = { accessToken: value.access_token, expiresAt: Date.now() + seconds * 1000, scope: value.scope };
  const next = validateBundle(resource === 'graph' ? { ...updated, ...token, format: FORMAT, generation: bundle.generation ?? randomUUID() } : { ...updated, resources: { ...updated.resources, [resource]: token } });
  // Graph owns its tokens. They may be opaque, or use either Graph audience representation.
  if (resource === 'graph') await identify(next, { fetcher, signal });
  return next;
}

export class DiagnosticSession {
  constructor(directory, { browserPath, fetcher = fetch, capture } = {}) {
    Object.assign(this, { directory, browserPath, fetcher, capture });
    this.file = path.join(directory, 'credentials.json');
    this.controller = new AbortController();
  }
  async save(bundle) {
    const temp = path.join(this.directory, `.credentials-${randomUUID()}.tmp`);
    try { await fs.writeFile(temp, JSON.stringify(bundle), { mode: 0o600, flag: 'wx' }); await fs.rename(temp, this.file); }
    finally { await fs.unlink(temp).catch(() => {}); }
    this.credentials = bundle;
  }
  async acquire({ silent = false, capture = false } = {}) {
    await fs.mkdir(this.directory, { recursive: true, mode: 0o700 }); await fs.chmod(this.directory, 0o700);
    if (!this.lock) {
      try { this.lock = await fs.open(path.join(this.directory, 'session.lock'), 'wx', 0o600); await this.lock.writeFile(String(process.pid)); }
      catch (error) { if (error.code === 'EEXIST') throw fail('session_busy', 'Another diagnostic is using the saved Microsoft session.'); throw error; }
    }
    let previous;
    try { previous = validateBundle(JSON.parse(await fs.readFile(this.file, 'utf8'))); }
    catch (error) { if (error.code !== 'ENOENT') throw fail('cache_invalid', 'The saved Microsoft credential cache is invalid.'); }
    if (previous && !capture) {
      if (fresh(previous)) { this.credentials = previous; return previous; }
      try {
        const renewed = await refresh(previous, { fetcher: this.fetcher, signal: this.controller.signal, onRotate: value => this.save(value) });
        await this.save(renewed); return renewed;
      } catch (error) { if (!error.reauthenticationRequired) throw error; }
    }
    const acquire = this.capture ?? (await import('./browser.js')).acquireBrowserRefresh;
    let captured;
    try { captured = await acquire(this.directory, { silent: true, browserPath: this.browserPath, signal: this.controller.signal }); }
    catch (error) {
      if (silent) throw fail('signin_required', 'Microsoft requires interactive sign-in. Run the access check without --silent.');
      captured = await acquire(this.directory, { silent: false, browserPath: this.browserPath, signal: this.controller.signal });
    }
    if (previous && (previous.identity.oid !== captured.identity.oid || previous.identity.tenant !== captured.identity.tenant)) throw fail('account_mismatch', 'Sign in with the cached Microsoft account.', 403);
    const credentials = await refresh(captured, { fetcher: this.fetcher, signal: this.controller.signal });
    await this.save(credentials); return credentials;
  }
  async close() {
    this.controller.abort();
    if (this.lock) { await this.lock.close(); this.lock = null; await fs.unlink(path.join(this.directory, 'session.lock')); }
  }
}

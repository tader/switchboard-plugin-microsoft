// Browser capture and HTTP renewal shared by production and the read-only diagnostic.
import fs from 'node:fs/promises';
import path from 'node:path';
import { CDP, findBrowser } from './cdp.js';
import { clientId, resources } from './oauth.js';
import { tokenClaims } from './api-http.js';
import { AdapterError } from './errors.js';

export const fresh = (token, now = Date.now()) => typeof token?.value === 'string' && token.value.length > 0 && Number.isFinite(token.expiresAt) && token.expiresAt > now + 120_000;

const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const origins = ['https://teams.microsoft.com', 'https://teams.cloud.microsoft'];
const fail = (code, message) => new AdapterError(code, message, 503);
function endpoint(raw) {
  try {
    const url = new URL(raw);
    return url.origin === 'https://login.microsoftonline.com' && !url.username && !url.password && !url.hash &&
      /^\/[a-z0-9._-]+\/oauth2\/(?:v2\.0\/)?token$/i.test(url.pathname) ? url : null;
  } catch { return null; }
}
function origin(headers = {}) {
  return Object.entries(headers).find(([key]) => key.toLowerCase() === 'origin')?.[1];
}
function identity(value) {
  // Metadata from Microsoft's HTTPS token response. Teams validates the access
  // tokens and confirms the account below; decoding a JWT is not signature validation.
  const claims = tokenClaims(value.id_token || value.access_token || '');
  if (!uuid.test(claims.tid ?? '') || !uuid.test(claims.oid ?? '')) throw fail('refresh_capture_identity', 'Microsoft returned no usable account metadata.');
  return { tenant: claims.tid, oid: claims.oid, email: claims.preferred_username || claims.upn || claims.email || '', name: claims.name || '' };
}

async function tokenResponse(response) {
  if (!response.body) return null;
  const reader = response.body.getReader(), chunks = []; let size = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.length;
      if (size > 512 * 1024) throw fail('refresh_response_invalid', 'Microsoft token response exceeded the size limit.');
      chunks.push(value);
    }
    return JSON.parse(Buffer.concat(chunks).toString('utf8'));
  } catch (error) {
    await reader.cancel().catch(() => {});
    throw error instanceof AdapterError ? error : fail('refresh_response_invalid', 'Could not read a valid Microsoft token response.');
  } finally { reader.releaseLock(); }
}

export async function captureRefreshResponse(page, cdp, { signal, timeout = 300_000 } = {}) {
  await page.call('Network.enable', { maxTotalBufferSize: 2 * 1024 * 1024, maxResourceBufferSize: 512 * 1024 });
  return new Promise((resolve, reject) => {
    const requests = new Map(), extraHeaders = new Map(); let settled = false;
    const sessions = new Map([[page.sessionId, page]]);
    const observed = { microsoftRequests: 0, teamsRequests: 0, successfulResponses: 0, targetFailures: 0 };
    const finish = (error, value) => {
      if (settled) return;
      settled = true; clearTimeout(timer); unsubscribe(); signal?.removeEventListener('abort', abort);
      requests.clear(); extraHeaders.clear(); sessions.clear();
      if (error) reject(error); else resolve(value);
    };
    const abort = () => finish(fail('refresh_check_cancelled', 'Teams authentication cancelled.'));
    const unsubscribe = cdp.subscribe(event => {
      if (settled || !sessions.has(event.sessionId)) return;
      const p = event.params;
      if (event.method === 'Target.attachedToTarget') {
        const child = { call: (method, args) => cdp.call(method, args, p.sessionId) };
        sessions.set(p.sessionId, child);
        (async () => {
          await child.call('Network.enable', { maxTotalBufferSize: 2 * 1024 * 1024, maxResourceBufferSize: 512 * 1024 });
          // Observe without pausing authentication: some workers cannot process
          // debugger commands until the parent page's startup has finished.
          if (p.targetInfo?.type === 'iframe') await child.call('Target.setAutoAttach', {
            autoAttach: true, waitForDebuggerOnStart: false, flatten: true,
          }).catch(() => {});
        })().catch(() => { observed.targetFailures++; sessions.delete(p.sessionId); });
        return;
      }
      if (event.method === 'Target.detachedFromTarget') { sessions.delete(p.sessionId); return; }
      const target = sessions.get(event.sessionId), key = `${event.sessionId}:${p.requestId}`;
      if (event.method === 'Network.requestWillBeSentExtraInfo') {
        if (extraHeaders.size < 1000) extraHeaders.set(key, p.headers);
        return;
      }
      if (event.method === 'Network.requestWillBeSent') {
        // A redirect invalidates the earlier candidate, even if its request ID is reused.
        requests.delete(key);
        if (p.request.method !== 'POST' || !endpoint(p.request.url) || requests.size >= 100) return;
        observed.microsoftRequests++;
        const metadata = (async () => {
          const postData = p.request.postData ?? (await target.call('Network.getRequestPostData', { requestId: p.requestId })).postData;
          if (typeof postData !== 'string' || postData.length > 256 * 1024) return null;
          const form = new URLSearchParams(postData);
          if (form.getAll('client_id').length !== 1 || form.get('client_id') !== clientId || form.has('client_secret') || form.getAll('grant_type').length !== 1 ||
              !['authorization_code', 'refresh_token'].includes(form.get('grant_type'))) return null;
          observed.teamsRequests++;
          return { headers: p.request.headers, grantType: form.get('grant_type'), url: p.request.url };
        })().catch(() => null);
        requests.set(key, { metadata, accepted: false });
      } else if (event.method === 'Network.responseReceived') {
        const request = requests.get(key);
        if (request) {
          request.responseURL = p.response.url;
          request.accepted = p.response.status === 200 && Boolean(endpoint(p.response.url)) &&
            !p.response.fromDiskCache && !p.response.fromServiceWorker;
        }
      } else if (event.method === 'Network.loadingFailed') {
        requests.delete(key); extraHeaders.delete(key);
      } else if (event.method === 'Network.loadingFinished') {
        const request = requests.get(key);
        if (!request?.accepted) return;
        // CDP handlers must never throw or leak response bodies through browser errors.
        (async () => {
          const metadata = await request.metadata;
          if (settled || !metadata || metadata.url !== request.responseURL) return;
          const requestOrigin = origin(extraHeaders.get(key)) || origin(metadata.headers);
          if (!origins.includes(requestOrigin)) return;
          observed.successfulResponses++;
          const response = await target.call('Network.getResponseBody', { requestId: p.requestId });
          if (settled || typeof response.body !== 'string' || response.body.length > 700_000) return;
          const body = response.base64Encoded ? Buffer.from(response.body, 'base64').toString('utf8') : response.body;
          if (Buffer.byteLength(body) > 512 * 1024) return;
          const value = JSON.parse(body);
          if (typeof value.refresh_token !== 'string' || !value.refresh_token || value.refresh_token.length > 100_000) return;
          const account = identity(value);
          finish(null, { refreshToken: value.refresh_token, clientId, origin: requestOrigin, identity: account,
            capturedAt: Date.now(), sourceGrant: metadata.grantType });
        })().catch(() => finish(fail('refresh_capture_failed', 'Could not read valid Teams renewal credentials from Microsoft.')));
      }
    });
    const timer = setTimeout(async () => {
      let location = 'unknown';
      try {
        const tree = await page.call('Page.getFrameTree');
        const host = new URL(tree.frameTree.frame.url).hostname;
        location = origins.some(value => new URL(value).hostname === host) ? 'teams_app' : host === 'login.microsoftonline.com' ? 'microsoft_signin' : 'other_page';
      } catch {}
      finish(fail('refresh_capture_timeout', `No Teams refresh token was captured (page: ${location}; Microsoft requests: ${observed.microsoftRequests}; Teams requests: ${observed.teamsRequests}; successful responses: ${observed.successfulResponses}; target failures: ${observed.targetFailures}). Complete sign-in in the authentication browser and retry.`));
    }, timeout);
    cdp.exit?.then(() => finish(fail('refresh_browser_closed', 'The authentication browser closed before refresh-token capture.')));
    signal?.addEventListener('abort', abort, { once: true });
    if (signal?.aborted) { abort(); return; }
    page.call('Target.setAutoAttach', { autoAttach: true, waitForDebuggerOnStart: false, flatten: true })
      .then(() => page.call('Page.navigate', { url: 'https://teams.microsoft.com/v2/' })).then(result => {
      if (result.errorText) finish(fail('refresh_navigation_failed', 'Could not open Teams for authentication.'));
    }, () => finish(fail('refresh_navigation_failed', 'Could not open Teams for authentication.')));
  });
}

export async function acquireBrowserRefresh(directory, { silent = false, browserPath, signal } = {}, {
  createCDP = (exe, dir, options) => new CDP(exe, dir, options), discoverBrowser = findBrowser,
} = {}) {
  signal?.throwIfAborted();
  const profile = path.join(directory, 'auth-browser');
  await fs.mkdir(profile, { recursive: true, mode: 0o700 }); await fs.chmod(profile, 0o700);
  const cdp = createCDP(await discoverBrowser(browserPath), profile, { headless: silent });
  const abort = () => { cdp.close().catch(() => {}); };
  signal?.addEventListener('abort', abort, { once: true });
  try {
    signal?.throwIfAborted();
    const page = await cdp.page('about:blank');
    // Only this dedicated authentication profile is touched. Keep Microsoft login cookies,
    // but clear Teams app caches so an existing AT cannot conceal the token exchange.
    for (const appOrigin of origins) await page.call('Storage.clearDataForOrigin', {
      origin: appOrigin, storageTypes: 'local_storage,indexeddb,cache_storage,service_workers',
    });
    return await captureRefreshResponse(page, cdp, { signal, timeout: silent ? 60_000 : 300_000 });
  } finally {
    signal?.removeEventListener('abort', abort);
    await cdp.close();
    // Acquisition must establish that the browser process has exited.
    if (cdp.exit) {
      let timer;
      try { await Promise.race([cdp.exit, new Promise((_, reject) => { timer = setTimeout(() => reject(fail('refresh_browser_exit_failed', 'The authentication browser did not exit. Authentication could not finish.')), 5000); })]); }
      finally { clearTimeout(timer); }
    }
  }
}

export async function refreshAccessTokens(captured, { fetcher = fetch, signal, onRotate = async () => {}, tokenResources = resources } = {}) {
  if (captured.clientId !== clientId || !origins.includes(captured.origin) || !uuid.test(captured.identity?.tenant ?? '') || !uuid.test(captured.identity?.oid ?? '')) {
    throw fail('refresh_metadata_invalid', 'Invalid Teams refresh-token metadata.');
  }
  let refreshToken = captured.refreshToken; const tokens = {};
  if (typeof refreshToken !== 'string' || !refreshToken || refreshToken.length > 100_000) throw fail('refresh_metadata_invalid', 'Invalid Teams refresh-token metadata.');
  for (const [service, resource] of Object.entries(tokenResources)) {
    signal?.throwIfAborted();
    const response = await fetcher(`https://login.microsoftonline.com/${captured.identity.tenant}/oauth2/v2.0/token`, {
      method: 'POST', redirect: 'error', signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(20_000)]) : AbortSignal.timeout(20_000),
      headers: { 'content-type': 'application/x-www-form-urlencoded', Origin: captured.origin },
      body: new URLSearchParams({ client_id: clientId, grant_type: 'refresh_token', refresh_token: refreshToken,
        scope: `${resource}/.default openid profile offline_access` }).toString(),
    }).catch(() => { throw fail('refresh_transport_failed', 'Microsoft token renewal failed or timed out.'); });
    // Bounded response reader, including for OAuth errors. Never expose error_description.
    const value = await tokenResponse(response);
    if (!response.ok || value?.error) {
      const aadsts = (Array.isArray(value?.error_codes) ? value.error_codes : []).find(code => Number.isSafeInteger(code) && code > 0);
      const error = fail(response.status === 429 ? 'refresh_rate_limited' : 'refresh_exchange_rejected', `Microsoft rejected HTTP renewal for ${service}${aadsts ? ` (AADSTS${aadsts})` : ` (HTTP ${response.status})`}.`);
      error.reauthenticationRequired = ['invalid_grant', 'interaction_required', 'login_required', 'consent_required'].includes(value?.error);
      throw error;
    }
    const seconds = Number(value?.expires_in), claims = tokenClaims(value?.access_token ?? '');
    if (typeof value?.access_token !== 'string' || value.access_token.length > 100_000 || typeof value.token_type !== 'string' || value.token_type.toLowerCase() !== 'bearer' ||
        !Number.isFinite(seconds) || seconds <= 120 || seconds > 604800 || !Number.isFinite(claims.exp)) throw fail('refresh_response_invalid', 'Microsoft returned an invalid access token.');
    if (claims.aud !== resource || claims.tid !== captured.identity.tenant || claims.oid !== captured.identity.oid) {
      throw fail('account_mismatch', 'Microsoft renewed tokens for a different account or service.');
    }
    const token = { value: value.access_token, expiresAt: Math.min(Date.now() + seconds * 1000, claims.exp * 1000) };
    if (!fresh(token) || claims.nbf !== undefined && (!Number.isFinite(claims.nbf) || claims.nbf > Date.now() / 1000 + 60)) throw fail('refresh_response_invalid', 'Microsoft returned an unusable access token.');
    tokens[service] = token;
    if (value.refresh_token !== undefined) {
      if (typeof value.refresh_token !== 'string' || !value.refresh_token || value.refresh_token.length > 100_000) throw fail('refresh_response_invalid', 'Microsoft returned invalid renewal credentials.');
      refreshToken = value.refresh_token;
      await onRotate({ value: refreshToken, clientId, origin: captured.origin, capturedAt: captured.capturedAt, sourceGrant: captured.sourceGrant });
    }
  }
  return { authVersion: 1, authMode: 'tokens', identity: captured.identity, tokens,
    renewal: { value: refreshToken, clientId, origin: captured.origin, capturedAt: captured.capturedAt, sourceGrant: captured.sourceGrant } };
}

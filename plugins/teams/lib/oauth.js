import { randomUUID, createPublicKey, verify } from 'node:crypto';

export const clientId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
export const resources = { skype: 'https://api.spaces.skype.com', chatsvcagg: 'https://chatsvcagg.teams.microsoft.com' };

export function authorization(stage, tenant = 'common', loginHint, silent = false) {
  if (!['teams', 'skype', 'chatsvcagg'].includes(stage)) throw new Error('invalid_stage');
  if (tenant !== 'common' && !/^[a-f0-9-]{36}$/i.test(tenant)) throw new Error('invalid_tenant');
  const state = randomUUID(), nonce = randomUUID();
  const url = new URL(`https://login.microsoftonline.com/${tenant}/oauth2/authorize`);
  url.search = new URLSearchParams({ client_id: clientId, redirect_uri: 'https://teams.microsoft.com/go',
    response_type: stage === 'teams' ? 'id_token' : 'token', response_mode: 'fragment', state, nonce, 'client-request-id': randomUUID() });
  if (resources[stage]) url.searchParams.set('resource', resources[stage]);
  if (loginHint) url.searchParams.set('login_hint', loginHint);
  if (silent) url.searchParams.set('prompt', 'none');
  return { stage, state, nonce, tenant, url: url.href };
}

export function callback(raw, request, details = false) {
  let url;
  try { url = new URL(raw); } catch { return null; }
  if (url.origin !== 'https://teams.microsoft.com' || url.pathname !== '/go') return null;
  const params = new URLSearchParams(url.hash.slice(1));
  if (!params.has('state')) return null;
  if (params.getAll('state').length !== 1 || params.get('state') !== request.state) throw new Error('oauth_state_mismatch');
  if (params.has('error')) {
    // Never print error_description: providers can include identity or callback data.
    const aadsts = params.get('error_description')?.match(/AADSTS\d+/)?.[0];
    throw new Error(`oauth_error:${params.get('error')?.replace(/[^a-z_]/gi, '').slice(0, 80)}${aadsts ? ':' + aadsts : ''}`);
  }
  const field = request.stage === 'teams' ? 'id_token' : 'access_token';
  const token = params.get(field);
  if (params.getAll(field).length !== 1 || !token || token.length > 100_000) throw new Error('oauth_missing_token');
  if (!details) return token;
  const seconds = Number(params.get('expires_in'));
  if (request.stage !== 'teams' && (!Number.isFinite(seconds) || seconds <= 0 || seconds > 604800)) throw new Error('oauth_invalid_expiry');
  return { value: token, expiresAt: request.stage === 'teams' ? null : Date.now() + seconds * 1000 };
}

export async function validateIdentity(token, request, fetcher = fetch) {
  return (await validateIdentityClaims(token, request, fetcher)).tid;
}

export async function validateIdentityClaims(token, request, fetcher = fetch) {
  const parts = token.split('.');
  if (parts.length !== 3) throw new Error('invalid_identity_token');
  let header, claims;
  try { header = JSON.parse(Buffer.from(parts[0], 'base64url')); claims = JSON.parse(Buffer.from(parts[1], 'base64url')); }
  catch { throw new Error('invalid_identity_token'); }
  const now = Date.now() / 1000;
  if (header.alg !== 'RS256' || claims.aud !== clientId || claims.nonce !== request.nonce ||
      !/^[a-f0-9-]{36}$/i.test(claims.tid ?? '') || !Number.isFinite(claims.exp) || claims.exp <= now ||
      (claims.nbf !== undefined && (!Number.isFinite(claims.nbf) || claims.nbf > now + 60)) ||
      claims.iss !== `https://sts.windows.net/${claims.tid}/` ||
      (request.tenant !== 'common' && claims.tid !== request.tenant)) throw new Error('identity_claims_rejected');
  // Fixed Microsoft host; no issuer or key URL supplied by the token is followed.
  const response = await fetcher('https://login.microsoftonline.com/common/discovery/keys', { redirect: 'error', signal: AbortSignal.timeout(20_000) });
  if (!response.ok) throw new Error('identity_keys_unavailable');
  const keys = (await response.json()).keys;
  const key = keys?.find(key => key.kid === header.kid && key.kty === 'RSA' && key.use === 'sig');
  if (!key || !verify('RSA-SHA256', Buffer.from(parts.slice(0, 2).join('.')),
    createPublicKey({ key, format: 'jwk' }), Buffer.from(parts[2], 'base64url'))) throw new Error('identity_signature_rejected');
  return claims;
}

export async function capture(page, cdp, request, timeout = 300_000, { details = false, signal } = {}) {
  await page.call('Page.enable');
  await page.call('Network.enable');
  const { frameTree } = await page.call('Page.getFrameTree');
  let mainFrameId = frameTree.frame.id;
  const diagnostics = { callbackSeen: false, callbackHadFragment: false, teamsAppSeen: false };
  // Both Page.Frame and Network.Request split the fragment off the URL.
  // OAuth implicit-flow credentials live in that fragment, not in `url`.
  const fullURL = value => value.urlFragment === undefined ? value.url : value.url.split('#')[0] + value.urlFragment;
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (error, token) => {
      if (settled) return;
      settled = true; clearTimeout(timer); unsubscribe(); signal?.removeEventListener('abort', aborted);
      if (error) reject(error); else resolve(token);
    };
    const unsubscribe = cdp.subscribe(event => {
      if (settled || event.sessionId !== page.sessionId) return;
      const params = event.params;
      let raw;
      if (event.method === 'Page.frameNavigated' && !params.frame.parentId) {
        mainFrameId = params.frame.id;
        raw = fullURL(params.frame);
      } else if (event.method === 'Page.navigatedWithinDocument' && params.frameId === mainFrameId) {
        raw = params.url;
      } else if (event.method === 'Network.requestWillBeSent' && params.frameId === mainFrameId && params.type === 'Document') {
        // /go can issue a server redirect, so it may never appear in frameNavigated.
        raw = fullURL(params.request);
      }
      if (!raw) return;
      try {
        const url = new URL(raw);
        if (url.origin === 'https://teams.microsoft.com' && url.pathname === '/go') {
          diagnostics.callbackSeen = true;
          diagnostics.callbackHadFragment ||= Boolean(url.hash);
        }
        diagnostics.teamsAppSeen ||= url.origin === 'https://teams.cloud.microsoft' ||
          url.origin === 'https://teams.microsoft.com' && url.pathname.startsWith('/v2');
        const token = callback(raw, request, details);
        if (token) finish(null, token);
      }
      catch (error) { finish(error); }
    });
    const timer = setTimeout(() => finish(Object.assign(new Error('oauth_timeout'), { diagnostics })), timeout);
    const aborted = () => finish(new Error('oauth_cancelled'));
    cdp.exit?.then(() => finish(new Error('oauth_browser_closed')));
    signal?.addEventListener('abort', aborted, { once: true });
    if (signal?.aborted) { aborted(); return; }
    page.call('Page.navigate', { url: request.url }).then(result => {
      if (result.errorText) finish(new Error('oauth_navigation_failed'));
    }, () => finish(new Error('oauth_navigation_failed')));
  });
}

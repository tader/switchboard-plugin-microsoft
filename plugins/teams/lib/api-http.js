import { AdapterError } from './errors.js';

export async function requestJSON(url, options = {}, fetcher = fetch) {
  let response;
  const timeout = AbortSignal.timeout(20_000);
  const signal = options.signal ? AbortSignal.any([options.signal, timeout]) : timeout;
  try { response = await fetcher(url, { ...options, redirect: 'error', signal }); }
  catch { throw new AdapterError('api_transport_failed', 'Teams API request failed or timed out.', 502); }
  if (!response.ok) {
    await response.body?.cancel();
    if (response.status === 401) throw new AdapterError('api_unauthorized', 'Teams rejected the token. Renew authentication.', 401);
    if (response.status === 403) throw new AdapterError('api_forbidden', 'Teams denied this operation.', 403);
    if (response.status === 404) throw new AdapterError('api_not_found', 'Teams could not find this resource.', 404);
    if (response.status === 429) throw new AdapterError('api_rate_limited', 'Teams rate limit reached. Try again later.', 429);
    throw new AdapterError('api_response_failed', `Teams API returned HTTP ${response.status}.`, 502);
  }
  if (!response.body) return null;
  const reader = response.body.getReader(), chunks = []; let size = 0;
  try {
    while (true) {
      const { done, value } = await reader.read(); if (done) break;
      size += value.length;
      if (size > 8 * 1024 * 1024) throw new AdapterError('api_response_too_large', 'Teams response exceeded 8 MiB.', 502);
      chunks.push(value);
    }
  } catch (error) { await reader.cancel().catch(() => {}); throw error instanceof AdapterError ? error : new AdapterError('api_transport_failed', 'Teams response was interrupted.', 502); }
  const text = Buffer.concat(chunks).toString('utf8');
  if (!text.trim()) return null;
  try { return JSON.parse(text); } catch { throw new AdapterError('api_invalid_json', 'Teams returned an invalid JSON response.', 502); }
}

export function tokenClaims(value) {
  // Metadata only. Identity trust comes from validated ID tokens; APIs validate access tokens.
  try { return JSON.parse(Buffer.from(value.split('.')[1], 'base64url')); } catch { return {}; }
}

export function chatOrigin(authz) {
  const region = authz.region;
  if (typeof region !== 'string' || !/^[a-z]{2,10}$/.test(region)) throw new AdapterError('api_region_invalid', 'Teams returned an unsupported region.', 502);
  const url = new URL(authz.regionGtms?.chatService || `https://${region}.ng.msg.teams.microsoft.com`);
  if (url.protocol !== 'https:' || url.username || url.password || url.port || !/^[a-z0-9-]+\.ng\.msg\.teams\.microsoft\.com$/.test(url.hostname)) {
    throw new AdapterError('api_region_invalid', 'Teams returned an unsupported chat service host.', 502);
  }
  return url.origin;
}

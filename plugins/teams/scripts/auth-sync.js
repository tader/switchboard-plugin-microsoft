import path from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { BrowserAuthenticator } from '../lib/api-auth.js';
import { encodeTokenBundle } from '../lib/token-bundle.js';
import { recipientEmail } from '../lib/input.js';
import { AdapterError } from '../lib/errors.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../.local-profile/token-sync');
const fail = (code, message) => new AdapterError(code, message, 503);
const needsSignIn = status => status.tokenImportRequired === true || status.signInRequired === true;
const healthy = status => status.ready === true && Date.parse(status.tokenExpiresAt) > Date.now() + 120_000;

function settings({ url, connection, token, loginHint }) {
  let base;
  try { base = new URL(url); } catch { throw fail('sync_url_invalid', 'Provide --url or SWITCHBOARD_URL with your Switchboard URL.'); }
  if (base.username || base.password || base.search || base.hash ||
      !(base.protocol === 'https:' || base.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(base.hostname))) {
    throw fail('sync_url_invalid', 'Use an HTTPS Switchboard URL, or HTTP on localhost, without embedded credentials, query or fragment.');
  }
  if (typeof connection !== 'string' || !connection.trim() || connection.length > 256) throw fail('sync_connection_required', 'Provide --connection or TEAMS_CONNECTION with an existing Teams connection name or ID.');
  if (typeof token !== 'string' || !token.trim() || /[\r\n]/.test(token)) throw fail('sync_token_required', 'Set SWITCHBOARD_TOKEN to a full-access Switchboard API token.');
  if (loginHint) loginHint = recipientEmail(loginHint);
  base.pathname = base.pathname.replace(/\/+$/, '') + '/';
  return { base, connection: connection.trim(), token: token.trim(), loginHint };
}

async function json(response) {
  const reader = response.body?.getReader();
  if (!reader) throw fail('sync_response_invalid', 'Switchboard returned an empty response.');
  const chunks = []; let length = 0;
  try {
    while (true) {
      const { done, value } = await reader.read(); if (done) break;
      length += value.length;
      if (length > 256 * 1024) throw fail('sync_response_invalid', 'Switchboard returned an oversized response.');
      chunks.push(value);
    }
    return JSON.parse(Buffer.concat(chunks).toString('utf8'));
  } catch (error) {
    await reader.cancel().catch(() => {});
    throw error instanceof AdapterError ? error : fail('sync_response_invalid', 'Switchboard returned an invalid JSON response.');
  } finally { reader.releaseLock(); }
}

function validateStatus(status) {
  if (!status || status.transport !== 'api' || typeof status.ready !== 'boolean' || typeof status.authenticating !== 'boolean' ||
      !(status.tokenExpiresAt === null || typeof status.tokenExpiresAt === 'string' && Number.isFinite(Date.parse(status.tokenExpiresAt)))) {
    throw fail('sync_status_invalid', 'The connection did not return supported Teams auth status. Update the Teams plugin.');
  }
  if (status.authenticating) throw fail('sync_signin_pending', 'Teams sign-in is already in progress. Complete it before syncing.');
  return status;
}

export async function syncTokens(options, {
  fetcher = fetch,
  createAuthenticator = directory => new BrowserAuthenticator(directory, { browserPath: process.env.TEAMS_BROWSER_PATH }),
  report = () => {},
} = {}) {
  const { base, connection, token, loginHint } = settings(options);
  const { force = false, silent = false, signal } = options;
  const request = async (route, body, { discard = false, allowFailure = false } = {}) => {
    signal?.throwIfAborted();
    let response;
    try {
      response = await fetcher(new URL(route, base), {
        method: body === undefined ? 'GET' : 'POST', redirect: 'error',
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', ...(body === undefined ? {} : { 'content-type': 'application/json' }) },
        ...(body === undefined ? {} : { body: JSON.stringify(body) }),
        signal: signal ? AbortSignal.any([signal, AbortSignal.timeout(120_000)]) : AbortSignal.timeout(120_000),
      });
    } catch {
      signal?.throwIfAborted();
      throw fail('sync_transport_failed', 'Could not reach Switchboard, or the request timed out. Check the server before retrying.');
    }
    if (!response.ok && !allowFailure || response.status === 401) {
      await response.body?.cancel().catch(() => {});
      if (response.status === 401) throw fail('sync_access_denied', 'Switchboard rejected the API token. Use a valid full-access API token.');
      if (response.status === 403) throw fail('sync_access_denied', 'Switchboard refused this request. Reconnect requires a full-access API token.');
      throw fail('sync_request_failed', `Switchboard rejected the request (HTTP ${response.status}). Check the connection on the server.`);
    }
    if (discard) { await response.body?.cancel().catch(() => {}); return response.status; }
    return json(response);
  };
  const metadata = await request(`api/connections/${encodeURIComponent(connection)}`);
  if (!metadata || typeof metadata.id !== 'string' || !metadata.id ||
      !(metadata.serviceId === 'teams-web' || /^sat\/[^/]+\/teams-web$/.test(metadata.serviceId))) {
    throw fail('sync_connection_invalid', 'Choose an existing Teams connection. No credentials were uploaded.');
  }
  const id = encodeURIComponent(metadata.id);
  const status = () => request(`proxy/${id}/status`).then(validateStatus);
  let current = await status();
  report('Checked Teams authentication on Switchboard.');
  if (!force && healthy(current) && !needsSignIn(current)) return { action: 'unchanged', expiresAt: current.tokenExpiresAt };
  if (!force && !needsSignIn(current)) {
    // A read-only call lets the server's on-demand HTTP renewal run. Discard chat
    // data rather than sending it to stdout or storing it on the laptop.
    report('Checking server-side HTTP renewal.');
    const code = await request(`proxy/${id}/chats`, undefined, { discard: true, allowFailure: true });
    current = await status();
    if (code === 200 && healthy(current) && !needsSignIn(current)) return { action: 'server_renewed', expiresAt: current.tokenExpiresAt };
    if (!needsSignIn(current)) throw fail('sync_renewal_failed', 'Server renewal did not restore Teams authentication. Retry later or use --force to acquire replacement credentials.');
  }
  const key = createHash('sha256').update(`${base.href}\n${metadata.id}`).digest('hex').slice(0, 24);
  const directory = path.join(root, key);
  const acquire = async interactive => {
    signal?.throwIfAborted();
    const auth = createAuthenticator(directory);
    const abort = () => { auth.close().catch(() => {}); };
    signal?.addEventListener('abort', abort, { once: true });
    try { signal?.throwIfAborted(); return await auth.acquire({ interactive, loginHint }); }
    finally { signal?.removeEventListener('abort', abort); await auth.close(); }
  };
  report('Obtaining replacement tokens using the saved local browser session.');
  let credentials;
  try { credentials = await acquire(false); }
  catch (error) {
    signal?.throwIfAborted();
    if (silent || !['signin_required', 'signin_failed'].includes(error.code)) throw error;
    report('Complete Microsoft sign-in in the dedicated browser window.');
    credentials = await acquire(true);
  }
  signal?.throwIfAborted();
  const tokenBundle = encodeTokenBundle(credentials);
  report('Browser closed. Uploading the replacement token bundle.');
  // Keep bearer credentials in memory; no export file and no automatic POST retry.
  const result = await request(`api/connections/${id}/reconnect`, { method: 'tokens', config: { label: metadata.account?.label || 'Teams', tokenBundle } });
  if (result?.status !== 'connected' || result.connection?.id !== metadata.id) throw fail('sync_upload_unconfirmed', 'Switchboard did not confirm reconnect. Check server status before retrying.');
  report('Tokens uploaded. Verifying Switchboard authentication.');
  current = await status();
  if (!healthy(current) || needsSignIn(current)) throw fail('sync_verification_failed', 'Tokens were uploaded, but Switchboard is not ready. Check the connection on the server.');
  return { action: 'uploaded', expiresAt: current.tokenExpiresAt };
}

export async function main() {
  const controller = new AbortController();
  const abort = () => controller.abort();
  process.once('SIGINT', abort); process.once('SIGTERM', abort);
  try {
    let values;
    try { ({ values } = parseArgs({ options: { url: { type: 'string' }, connection: { type: 'string' }, 'login-hint': { type: 'string' }, force: { type: 'boolean' }, silent: { type: 'boolean' }, help: { type: 'boolean' } }, strict: true, allowPositionals: false })); }
    catch { throw fail('sync_arguments_invalid', 'Usage: npm run auth:sync -- --url <Switchboard URL> --connection <Teams name or ID> [--login-hint <email>] [--force] [--silent]'); }
    if (values.help) {
      console.log('Usage: npm run auth:sync -- --url <Switchboard URL> --connection <Teams name or ID> [--login-hint <email>] [--force] [--silent]');
      console.log('Environment: SWITCHBOARD_URL, SWITCHBOARD_TOKEN (full-access API token), TEAMS_CONNECTION, TEAMS_LOGIN_HINT, TEAMS_BROWSER_PATH.');
      return;
    }
    const result = await syncTokens({ url: values.url || process.env.SWITCHBOARD_URL, token: process.env.SWITCHBOARD_TOKEN,
      connection: values.connection || process.env.TEAMS_CONNECTION, loginHint: values['login-hint'] || process.env.TEAMS_LOGIN_HINT,
      force: values.force, silent: values.silent, signal: controller.signal }, { report: message => console.log(message) });
    console.log(result.action === 'unchanged' ? 'Teams tokens are still valid; no browser or upload needed.' : result.action === 'server_renewed' ? 'Switchboard renewed Teams tokens; no browser or upload needed.' : 'Teams tokens replaced successfully.');
    console.log('Earliest access/chat-token expiry:', result.expiresAt);
  } catch (error) {
    console.error(controller.signal.aborted ? 'Teams token sync cancelled. Check server status if an upload was in progress.' : error instanceof AdapterError ? `${error.code}: ${error.message}` : 'Teams token sync failed. Check server status before retrying.');
    process.exitCode = controller.signal.aborted ? 130 : 1;
  } finally { process.removeListener('SIGINT', abort); process.removeListener('SIGTERM', abort); }
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) await main();

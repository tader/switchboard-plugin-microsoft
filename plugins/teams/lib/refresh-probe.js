// Read-only compatibility diagnostic using the same renewal code as production.
import { acquireBrowserRefresh, refreshAccessTokens } from './oauth-refresh.js';
import { exchangeChatToken } from './api-auth.js';
import { TeamsAPI } from './teams-api.js';
import { recipientEmail } from './input.js';
import { AdapterError } from './errors.js';
const fail = (code, message) => new AdapterError(code, message, 503);
export { captureRefreshResponse, acquireBrowserRefresh } from './oauth-refresh.js';
export const renewProbeTokens = refreshAccessTokens;

export async function checkRefreshCompatibility({ directory, silent, browserPath, loginHint, signal, report = () => {} }, {
  acquire = acquireBrowserRefresh, renew = renewProbeTokens, exchange = exchangeChatToken,
  createAPI = credentials => new TeamsAPI({ credentials }),
} = {}) {
  if (loginHint) loginHint = recipientEmail(loginHint);
  const captured = await acquire(directory, { silent, browserPath, signal });
  report('PASS: Teams refresh token captured; diagnostic browser exited');
  signal?.throwIfAborted();
  const bundle = await renew(captured, { signal });
  report('PASS: Skype and chatsvcagg access tokens renewed over HTTP');
  signal?.throwIfAborted();
  const credentials = await exchange(bundle);
  report('PASS: Teams chat-token exchange accepted the renewed token and confirmed the account');
  const api = createAPI(credentials);
  const email = loginHint || credentials.identity.email;
  const profile = await api.profile(email);
  if (profile.mri !== `8:orgid:${credentials.identity.oid}`) throw fail('account_mismatch', 'Teams profile does not match the renewed account.');
  report('PASS: direct API profile read and account binding');
  signal?.throwIfAborted();
  await api.chats();
  report('PASS: direct API conversation list using the renewed chatsvcagg token');
  return { expiresAt: Math.min(credentials.tokens.skype.expiresAt, credentials.tokens.chatsvcagg.expiresAt, credentials.chatToken.expiresAt),
    sourceGrant: captured.sourceGrant };
}

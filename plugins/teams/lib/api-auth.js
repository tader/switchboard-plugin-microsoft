import fs from 'node:fs/promises';
import path from 'node:path';
import { CDP, findBrowser } from './cdp.js';
import { authorization, capture, validateIdentityClaims, resources } from './oauth.js';
import { AdapterError } from './errors.js';
import { requestJSON, tokenClaims, chatOrigin } from './api-http.js';
import { fresh, acquireBrowserRefresh, refreshAccessTokens } from './oauth-refresh.js';
import { recipientEmail } from './input.js';

export { fresh } from './oauth-refresh.js';
export const readyCredentials = credentials => credentials?.authVersion === 1 && fresh(credentials.tokens?.skype) && fresh(credentials.tokens?.chatsvcagg) && fresh(credentials.chatToken);
export const importRequired = () => new AdapterError('token_import_required', 'Export fresh tokens on a machine with Chrome or Edge, then reconnect using Import Teams tokens.', 503);

export async function exchangeChatToken(credentials, fetcher = fetch) {
  const value = await requestJSON('https://teams.microsoft.com/api/authsvc/v1.0/authz', {
    method: 'POST', headers: { Authorization: `Bearer ${credentials.tokens.skype.value}`, 'ms-teams-authz-type': 'TokenRefresh' },
  }, fetcher);
  const token = value?.tokens?.skypeToken, expiresIn = Number(value?.tokens?.expiresIn);
  if (typeof token !== 'string' || !token || !Number.isFinite(expiresIn) || expiresIn <= 120 || expiresIn > 604800) {
    throw new AdapterError('api_authz_invalid', 'Teams returned an invalid chat token.', 502);
  }
  if (['tokens', 'microsoft'].includes(credentials.authMode) || credentials.renewal) {
    // This token came directly from the fixed Teams auth service over HTTPS,
    // after that service authenticated the imported Skype access token.
    const claims = tokenClaims(token);
    if (claims.skypeid !== `orgid:${credentials.identity.oid}` || claims.tid !== credentials.identity.tenant) {
      throw new AdapterError('account_mismatch', 'Teams authenticated a different account than the token bundle specifies.', 403);
    }
  }
  return { ...credentials, chatToken: { value: token, expiresAt: Date.now() + expiresIn * 1000 }, region: value.region, messageOrigin: chatOrigin(value) };
}

export class LegacyBrowserAuthenticator {
  constructor(profileDir, settings = {}, { fetcher = fetch, createCDP = (exe, dir, options) => new CDP(exe, dir, options) } = {}) {
    Object.assign(this, { profileDir, settings, fetcher, createCDP });
    this.controller = new AbortController();
  }
  async acquire({ interactive = false, previous, loginHint } = {}) {
    const directory = path.join(this.profileDir, 'auth-browser');
    await fs.mkdir(directory, { recursive: true, mode: 0o700 }); await fs.chmod(directory, 0o700);
    const cdp = this.cdp = this.createCDP(await findBrowser(this.settings.browserPath), directory, { headless: !interactive });
    try {
      const page = await cdp.page('about:blank');
      let tenant = previous?.identity?.tenant ?? 'common', identity;
      const tokens = {};
      for (const stage of ['teams', 'skype', 'chatsvcagg']) {
        const req = authorization(stage, tenant, loginHint || previous?.identity?.email, !interactive);
        const token = await capture(page, cdp, req, interactive ? 300_000 : 45_000, { details: true, signal: this.controller.signal });
        await page.call('Page.stopLoading'); await page.call('Page.navigate', { url: 'about:blank' });
        if (stage === 'teams') {
          const claims = await validateIdentityClaims(token.value, req, this.fetcher);
          if (typeof claims.oid !== 'string' || !/^[a-f0-9-]{36}$/i.test(claims.oid)) throw new Error('identity_claims_rejected');
          identity = { tenant: claims.tid, oid: claims.oid, email: claims.preferred_username || claims.upn || claims.email || '', name: claims.name || '' };
          if (previous?.identity && (identity.tenant !== previous.identity.tenant || identity.oid !== previous.identity.oid)) throw new AdapterError('account_mismatch', 'Sign in with the same Teams account as this connection.', 403);
          tenant = identity.tenant;
        } else {
          const claims = tokenClaims(token.value);
          if (claims.aud && claims.aud !== resources[stage] || claims.tid && claims.tid !== tenant || claims.oid && claims.oid !== identity.oid) throw new AdapterError('account_mismatch', 'Microsoft returned a token for a different account or service.', 403);
          if (Number.isFinite(claims.exp)) token.expiresAt = Math.min(token.expiresAt, claims.exp * 1000);
          if (!fresh(token)) throw new Error('oauth_invalid_expiry');
          tokens[stage] = token;
        }
      }
      return { authVersion: 1, identity, tokens };
    } catch (error) {
      if (error instanceof AdapterError) throw error;
      if (!interactive) throw new AdapterError('signin_required', 'Microsoft requires interactive sign-in. Run openTeamsLogin or reconnect this account.', 503);
      const code = /^(oauth_[a-z_:0-9]+|identity_[a-z_]+|browser_[a-z_]+)$/i.test(error.code ?? error.message) ? error.code ?? error.message : 'authentication_failed';
      throw new AdapterError('signin_failed', `Microsoft sign-in failed (${code}). Reconnect to try again.`, 503);
    } finally { await cdp.close(); if (this.cdp === cdp) this.cdp = null; }
  }
  async close() { this.controller.abort(); await this.cdp?.close(); }
}

export class BrowserAuthenticator {
  constructor(profileDir, settings = {}, { fetcher = fetch, createCDP = (exe, dir, options) => new CDP(exe, dir, options), acquire = acquireBrowserRefresh, refresh = refreshAccessTokens } = {}) {
    Object.assign(this, { profileDir, settings, fetcher, createCDP, acquireRefresh: acquire, refresh });
    this.controller = new AbortController();
  }
  async acquire({ interactive = false, previous, loginHint } = {}) {
    try {
      const captured = await this.acquireRefresh(this.profileDir, { silent: !interactive, browserPath: this.settings.browserPath, signal: this.controller.signal }, {
        createCDP: (...args) => { this.cdp = this.createCDP(...args); return this.cdp; },
      });
      if (previous?.identity && (captured.identity.tenant !== previous.identity.tenant || captured.identity.oid !== previous.identity.oid)) {
        throw new AdapterError('account_mismatch', 'Sign in with the same Teams account as this connection.', 403);
      }
      if (loginHint && recipientEmail(loginHint) !== recipientEmail(captured.identity.email)) throw new AdapterError('account_mismatch', 'Sign in with the requested Teams account.', 403);
      return await this.refresh(captured, { fetcher: this.fetcher, signal: this.controller.signal });
    } catch (error) {
      if (error.code === 'account_mismatch' || /^refresh_(transport_failed|rate_limited|response_invalid|exchange_rejected)$/.test(error.code ?? '')) throw error;
      throw new AdapterError(interactive ? 'signin_failed' : 'signin_required', interactive ? 'Teams sign-in failed. Reconnect to try again.' : 'Microsoft requires interactive sign-in. Run openTeamsLogin or reconnect this account.', 503);
    } finally { this.cdp = null; }
  }
  async close() { this.controller.abort(); await this.cdp?.close(); }
}

export class TeamsSession {
  constructor(profile, directory, settings = {}, { fetcher = fetch, createAuthenticator = (...args) => new BrowserAuthenticator(...args), exchange = exchangeChatToken, refresh = refreshAccessTokens } = {}) {
    Object.assign(this, { profile, directory, settings, fetcher, createAuthenticator, exchange, refresh });
    this.queue = Promise.resolve(); this.queued = 0; this.disposed = false; this.credentials = null; this.loginTask = null; this.authError = null;
  }
  serial(fn) {
    if (this.disposed) return Promise.reject(new AdapterError('disposed', 'Teams connection is closed.', 503));
    if (this.queued >= 20) return Promise.reject(new AdapterError('queue_full', 'Too many Teams operations.', 429));
    this.queued++;
    const result = this.queue.then(() => { if (this.disposed) throw new AdapterError('disposed', 'Teams connection is closed.', 503); return fn(); }).finally(() => this.queued--);
    this.queue = result.catch(() => {}); return result;
  }
  adopt(credentials) {
    if (credentials?.profile !== this.profile) throw new AdapterError('invalid_profile', 'Connection profile mismatch.', 403);
    if (!this.credentials && credentials.authVersion === 1 && credentials.identity?.tenant && credentials.identity?.oid && credentials.tokens?.skype?.value && credentials.tokens?.chatsvcagg?.value) this.credentials = credentials;
  }
  async authenticate(interactive, loginHint) {
    if (this.credentials?.authMode === 'tokens') throw importRequired();
    const auth = this.authenticator = this.createAuthenticator(this.directory, this.settings, { fetcher: this.fetcher });
    try {
      const bundle = await auth.acquire({ interactive, previous: this.credentials, loginHint });
      const credentials = await this.exchange({ ...bundle, authMode: 'browser', profile: this.profile }, this.fetcher);
      if (this.disposed) throw new AdapterError('disposed', 'Teams connection is closed.', 503);
      this.credentials = credentials; this.authError = null; return credentials;
    } finally { await auth.close(); if (this.authenticator === auth) this.authenticator = null; }
  }
  async login(loginHint) {
    if (this.disposed) throw new AdapterError('disposed', 'Teams connection is closed.', 503);
    if (this.credentials?.authMode === 'tokens') throw importRequired();
    if (!this.loginTask) {
      this.authError = null;
      this.loginTask = this.serial(() => this.authenticate(true, loginHint)).catch(error => { this.authError = error; }).finally(() => { this.loginTask = null; });
    }
    return { opened: true, authenticating: true, instructions: 'Complete Microsoft sign-in in the dedicated Chrome window. It closes automatically. Poll getTeamsStatus for completion.' };
  }
  async ensure(force = false) {
    if (this.loginTask) throw new AdapterError('signin_pending', 'Complete sign-in and poll getTeamsStatus.', 503);
    return this.serial(async () => {
      if (this.credentials?.authMode === 'tokens' && this.authError?.code === 'token_import_required') throw this.authError;
      if (!force && readyCredentials(this.credentials) && !this.authError) return this.credentials;
      if (!this.credentials) throw new AdapterError('signin_required', 'Reconnect this connection once to enable API authentication.', 503);
      try {
        if ((force || !fresh(this.credentials.tokens.skype) || !fresh(this.credentials.tokens.chatsvcagg)) && this.credentials.renewal) {
          try { await this.renew(); }
          catch (error) {
            if (!error.reauthenticationRequired) throw error;
            if (this.credentials.authMode === 'tokens') throw importRequired();
            await this.authenticate(false);
          }
        } else if (this.credentials.authMode === 'tokens' && (force || !fresh(this.credentials.tokens.skype) || !fresh(this.credentials.tokens.chatsvcagg))) throw importRequired();
        else if (!force && fresh(this.credentials.tokens.skype) && fresh(this.credentials.tokens.chatsvcagg)) {
          this.credentials = await this.exchange(this.credentials, this.fetcher);
        } else await this.authenticate(false);
        this.authError = null; return this.credentials;
      } catch (error) {
        if (this.credentials.authMode === 'tokens' && error.code === 'api_unauthorized') error = importRequired();
        this.authError = error; throw error;
      }
    });
  }
  async renew() {
    const previous = this.credentials, r = previous.renewal;
    const updated = await this.refresh({ ...r, refreshToken: r.value, identity: previous.identity }, {
      fetcher: this.fetcher,
      onRotate: async renewal => { if (this.disposed) throw new AdapterError('disposed', 'Teams connection is closed.', 503); this.credentials = { ...this.credentials, renewal }; },
    });
    if (this.disposed) throw new AdapterError('disposed', 'Teams connection is closed.', 503);
    this.credentials = { ...this.credentials, ...updated, authMode: previous.authMode, profile: this.profile };
    this.credentials = await this.exchange(this.credentials, this.fetcher);
    return this.credentials;
  }
  async ensureSearch() {
    return this.serial(async () => {
      if (fresh(this.credentials?.tokens?.substrate)) return this.credentials;
      const previous = this.credentials, r = previous?.renewal;
      if (!r) throw new AdapterError('search_token_required', 'Search requires renewal credentials. Reconnect with a version 2 token bundle.', 503);
      const updated = await this.refresh({ ...r, refreshToken: r.value, identity: previous.identity }, {
        fetcher: this.fetcher, tokenResources: { substrate: 'https://substrate.office.com' },
        onRotate: async renewal => { this.credentials = { ...this.credentials, renewal }; },
      });
      if (this.disposed) throw new AdapterError('disposed', 'Teams connection is closed.', 503);
      this.credentials = { ...this.credentials, tokens: { ...this.credentials.tokens, ...updated.tokens }, renewal: updated.renewal };
      return this.credentials;
    });
  }
  status() {
    const c = this.credentials;
    const tokenImportRequired = c?.authMode === 'tokens' && ((!c.renewal && (!fresh(c.tokens?.skype) || !fresh(c.tokens?.chatsvcagg))) || this.authError?.code === 'token_import_required');
    return { ready: readyCredentials(c) && !this.loginTask && !this.authError, transport: 'api', authenticating: Boolean(this.loginTask),
      authMode: c?.authMode ?? 'browser', tokenImportRequired: Boolean(tokenImportRequired),
      renewalAvailable: Boolean(c?.renewal), renewalMethod: c?.renewal ? 'http_refresh_token' : 'access_tokens_only',
      signInRequired: !this.loginTask && (!c || tokenImportRequired || this.authError?.code === 'signin_required' || this.authError?.code === 'signin_failed'),
      browserOpen: Boolean(this.authenticator?.cdp), tokenExpiresAt: c?.chatToken?.expiresAt && c.tokens?.skype?.expiresAt && c.tokens?.chatsvcagg?.expiresAt ? new Date(Math.min(c.tokens.skype.expiresAt, c.tokens.chatsvcagg.expiresAt, c.chatToken.expiresAt)).toISOString() : null,
      ...(this.authError ? { error: { code: this.authError.code || 'authentication_failed', message: this.authError instanceof AdapterError ? this.authError.message : 'Microsoft authentication failed.' } } : {}) };
  }
  async close() { this.disposed = true; await this.authenticator?.close(); await this.queue; this.credentials = null; }
}

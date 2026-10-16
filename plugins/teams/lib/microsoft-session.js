import { TeamsSession, exchangeChatToken, fresh, readyCredentials } from './api-auth.js';
import { AdapterError } from './errors.js';

/** Preserve the existing Teams adapter while sourcing Microsoft credentials centrally. */
export class MicrosoftTeamsSession extends TeamsSession {
  constructor(ctx, ...args) { super(...args); this.ctx = ctx; }
  get web() { return this.ctx.require('microsoft-authentication').webSession; }
  async source(previous, forceAuthentication = false) {
    const supplied = await this.ctx.credentials.obtain(this.sourceConfig.authenticationSource, { format: this.web.format, forceAuthentication, signal: this.signal });
    if (supplied.format !== this.web.format) throw new AdapterError('invalid_credentials', 'Incompatible Microsoft authentication source.', 403);
    const bundle = this.web.validateBundle(supplied.bundle, previous);
    const account = await this.web.identify(bundle, { signal: this.signal });
    if (account.id !== supplied.account.id) throw new AdapterError('account_mismatch', 'The authentication source account does not match Microsoft Graph.', 403);
    if (bundle.clientId !== this.web.teamsClientId) throw new AdapterError('teams_signin_required', 'Reconnect Microsoft Authentication using Teams sign-in to enable this connection.', 403);
    return { bundle, account };
  }
  async connectSource(config, previous) {
    this.sourceConfig = config;
    const { bundle, account } = await this.source(previous?.microsoft);
    this.credentials = { profile: this.profile, authVersion: 1, authMode: 'microsoft', identity: { ...bundle.identity, email: account.label, name: account.label }, microsoft: bundle, tokens: {} };
    return this.renewSource();
  }
  adopt(credentials) {
    if (credentials?.authMode !== 'microsoft') return super.adopt(credentials);
    if (credentials.profile !== this.profile) throw new AdapterError('invalid_profile', 'Connection profile mismatch.', 403);
    this.web.validateBundle(credentials.microsoft);
    // A failed operation may already have persisted an intermediate token rotation.
    if (!this.credentials || this.credentials.authMode !== 'microsoft') this.credentials = credentials;
  }
  async persist(bundle) {
    this.credentials = { ...this.credentials, microsoft: bundle };
    await this.persistCredentials?.(this.credentials);
  }
  async renewSource(force = false, search = false) {
    let bundle = this.web.validateBundle(this.credentials.microsoft);
    const renew = async resource => {
      bundle = await this.web.refresh(bundle, { resource, signal: this.signal, onRotate: next => this.persist(next) });
      await this.persist(bundle);
    };
    try {
      if (!this.web.fresh(bundle)) await renew('graph');
      for (const resource of ['skype', 'chatsvcagg', ...(search ? ['substrate'] : [])]) {
        if (force || !this.web.fresh(bundle.resources?.[resource])) await renew(resource);
      }
    } catch (error) {
      if (!error.reauthenticationRequired) throw error;
      const supplied = await this.source(bundle, true); bundle = supplied.bundle;
      await this.persist(bundle);
      for (const resource of ['skype', 'chatsvcagg', ...(search ? ['substrate'] : [])]) await renew(resource);
    }
    const tokens = Object.fromEntries(['skype', 'chatsvcagg', 'substrate'].filter(key => bundle.resources?.[key]).map(key => [key, { value: bundle.resources[key].accessToken, expiresAt: bundle.resources[key].expiresAt }]));
    this.credentials = { ...this.credentials, microsoft: bundle, tokens };
    if (force || !fresh(this.credentials.chatToken)) this.credentials = await exchangeChatToken(this.credentials, this.fetcher);
    await this.persistCredentials?.(this.credentials);
    return this.credentials;
  }
  async ensure(force = false) {
    if (this.credentials?.authMode !== 'microsoft') return super.ensure(force);
    return this.serial(async () => {
      if (!force && readyCredentials(this.credentials) && !this.authError) return this.credentials;
      try { const value = await this.renewSource(force); this.authError = null; return value; }
      catch (error) { this.authError = error; throw error; }
    });
  }
  async ensureSearch() {
    if (this.credentials?.authMode !== 'microsoft') return super.ensureSearch();
    return this.serial(() => fresh(this.credentials.tokens?.substrate) ? this.credentials : this.renewSource(false, true));
  }
  async login(hint) {
    if (this.credentials?.authMode !== 'microsoft') return super.login(hint);
    throw new AdapterError('signin_required', 'Sign in to the Microsoft Authentication connection on its device, then retry.', 503);
  }
  status() {
    const status = super.status();
    return this.credentials?.authMode === 'microsoft' ? { ...status, renewalAvailable: true, renewalMethod: 'microsoft_authentication_source', signInRequired: !!this.authError } : status;
  }
}

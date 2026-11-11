import { createConsumerAuth, baseUrl, backend, outlookSpec } from './consumer.ts';
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import type { PluginContext, Connection, AuthMethod, AuthenticationProfile } from './types/api.d.ts';
import { browserAuthorization } from './lib/browser-presentation.js';
import { acquireBrowserRefresh } from './lib/browser.js';
import { FORMAT, GRAPH, TEAMS_CLIENT_ID, RESOURCES, fresh, refresh, identify, validateBundle, fail, type Bundle } from './lib/session.js';

interface Credentials { profile: string; coreProfile?: boolean; bundle: Bundle }
interface PendingFlow { profile: string; controller: AbortController; result?: Credentials; error?: Error; expires: number; newProfile: boolean; coreProfile: boolean; work?: Promise<void>; account?: { id: string; label: string } }
const profilePath = (root: string, id: string) => {
  if (!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(id)) throw fail('invalid_profile', 'Invalid Microsoft browser profile.');
  return path.join(root, 'profiles', id);
};

export default function setup(ctx: PluginContext, { captureBrowser = acquireBrowserRefresh } = {}) {
  const pending = new Map<string, PendingFlow>();
  const controllers = new Set<AbortController>();
  async function cancelPending(id: string) {
    const flow = pending.get(id); if (!flow) return;
    pending.delete(id); flow.controller.abort(); await flow.work;
    if (flow.newProfile && !flow.coreProfile) await fs.rm(profilePath(ctx.dataDir, flow.profile), {recursive:true,force:true});
  }
  async function capture(profile: string, silent: boolean, previous?: Bundle, signal?: AbortSignal, shared?: AuthenticationProfile) {
    const captured = await captureBrowser(profilePath(ctx.dataDir, profile), { silent, browserPath: ctx.settings.browserPath, signal, sharedProfile: shared });
    if (previous && (previous.identity.tenant !== captured.identity.tenant || previous.identity.oid !== captured.identity.oid)) throw fail('account_mismatch', 'Sign in with the existing Microsoft account.', 403);
    return refresh(captured, { signal });
  }
  async function usable(conn: Connection, opts: { force?: boolean; signal?: AbortSignal; persistCredentials?: (credentials: any) => void | Promise<void> } = {}, forceAuthentication = false) {
    const credentials = conn.credentials as Credentials;
    let bundle = validateBundle(credentials.bundle);
    if (!forceAuthentication && !opts.force && fresh(bundle)) return credentials;
    if (!forceAuthentication) try {
      bundle = await refresh(bundle, { signal: opts.signal, onRotate: async next => { await opts.persistCredentials?.({ ...credentials, bundle: next }); } });
      return { ...credentials, bundle };
    } catch (error: any) { if (!error.reauthenticationRequired) throw error; }
    try { bundle = await capture(credentials.profile, true, bundle, opts.signal, conn.browserProfileId ? ctx.browserProfiles?.forConnection(conn) : undefined); }
    catch (error: any) {
      if (opts.signal?.aborted) throw opts.signal.reason;
      if (error.code === 'account_mismatch') throw error;
      throw fail('signin_required', 'Sign in to the Microsoft Authentication connection on your laptop.');
    }
    return { ...credentials, bundle };
  }
  const method: AuthMethod = {
    id: 'browser', name: 'Sign in through local browser',
    description: 'Sign in to Teams web for a shared Microsoft session. The browser closes after authentication. API access uses cached credentials and HTTP renewal.',
    browserProfile: ctx.browserProfiles ? { browserPath: ctx.settings.browserPath, browserFamily: 'edge', legacyPath: conn => `profiles/${path.basename(profilePath(ctx.dataDir, conn.credentials.profile))}/auth-browser` } : undefined,
    fields: [],
    unavailable: ctx.capabilities?.deviceBrowser ? undefined : 'Upgrade Switchboard to use browser sign-in.',
    async cancel(state) { await cancelPending(state?.id); },
    async connect({ connection, browserProfile }) {
      const device = browserAuthorization(ctx.capabilities);
      const profile = connection?.credentials?.profile ?? randomUUID();
      const id = randomUUID(), controller = new AbortController(); controllers.add(controller);
      const flow: PendingFlow = { profile, controller, expires: Date.now() + 600_000, newProfile: !connection, coreProfile: !!browserProfile };
      pending.set(id, flow);
      flow.work = (async () => {
        let bundle: Bundle;
        try { bundle = await capture(profile, !!browserProfile, connection?.credentials?.bundle, controller.signal, browserProfile); }
        catch (error: any) {
          if (!browserProfile || controller.signal.aborted || error.code === 'account_mismatch') throw error;
          bundle = await capture(profile, false, connection?.credentials?.bundle, controller.signal, browserProfile);
        }
        flow.account = await identify(bundle, { signal: controller.signal });
        flow.result = { profile, bundle, ...(browserProfile ? { coreProfile: true } : {}) };
      })().catch(error => { flow.error = error; }).finally(() => controllers.delete(controller));
      return { device: browserProfile ? { ...device, instructions: 'Waiting for this browser profile. Sign-in first runs headlessly; complete any requested sign-in in the browser on this device. You can cancel while waiting.' } : device, pending: { id } };
    },
    async poll({ pending: state }) {
      const flow = pending.get(state.id);
      if (!flow || flow.expires < Date.now()) throw fail('signin_expired', 'Microsoft sign-in expired. Reconnect to try again.');
      if (flow.error) { await cancelPending(state.id); throw flow.error; }
      if (!flow.result) return { wait: true };
      pending.delete(state.id);
      return { credentials: flow.result, account: flow.account, config: {} };
    },
    async authorize(req, conn, opts) {
      // Authentication sources expose identity only; sharing one is an explicit credential-provider grant.
      // Switchboard's connection test requests the base URL. Probe identity rather than rejecting a valid sign-in.
      if (['GET', 'HEAD'].includes(req.method) && req.url.origin === 'https://graph.microsoft.com' && ['/v1.0', '/v1.0/'].includes(req.url.pathname)) {
        req.url.pathname = '/v1.0/me';
        req.url.search = '?$select=id,displayName,mail,userPrincipalName';
      }
      if (!['GET', 'HEAD'].includes(req.method) || req.url.origin !== 'https://graph.microsoft.com' || req.url.pathname !== '/v1.0/me') throw fail('request_forbidden', 'Use a Microsoft product connection for data operations.', 403);
      const credentials = await usable(conn, opts);
      req.headers.set('authorization', `Bearer ${credentials.bundle.accessToken}`);
      return credentials !== conn.credentials ? { credentials } : undefined;
    },
    credentialProvider: {
      format: FORMAT,
      async provide(conn, opts) {
        const credentials = await usable(conn, opts, opts.forceAuthentication);
        return { bundle: validateBundle(credentials.bundle), credentials, account: await identify(credentials.bundle, opts) };
      },
    },
    async revoke(conn) { for (const [id, flow] of pending) if (flow.profile === conn.credentials.profile) { await cancelPending(id); } if (!conn.browserProfileId && !conn.credentials.coreProfile) await fs.rm(profilePath(ctx.dataDir, conn.credentials.profile), { recursive: true, force: true }); },
  };
  const prune = setInterval(() => { for (const [id, flow] of pending) if (flow.expires < Date.now()) { void cancelPending(id).catch(() => {}); } }, 30_000); prune.unref();
  return {
    services: [{ id: 'microsoft-authentication', name: 'Microsoft Authentication', description: 'Browser authentication for Microsoft API connections on this device and trusted peers.', icon: 'icon.svg', baseUrl: GRAPH, allowedHosts: ['graph.microsoft.com'], authMethods: [method] }],
    exports: { createConsumerAuth, baseUrl, backend, outlookSpec, webSession: { browserAuthorization: (expiresIn?: number) => browserAuthorization(ctx.capabilities, expiresIn), format: FORMAT, teamsClientId: TEAMS_CLIENT_ID, resources: RESOURCES, validateBundle, identify, refresh, fresh, fail } },
    async dispose() { clearInterval(prune); for (const controller of controllers) controller.abort(); await Promise.allSettled([...pending.keys()].map(cancelPending)); pending.clear(); },
  };
}

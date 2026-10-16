import type { AuthMethod, Connection, Field, PluginContext } from './types/api.d.ts';
import { FORMAT, GRAPH, RESOURCES, fresh, identify, refresh, validateBundle, fail, type Bundle } from './lib/session.js';
import { outlookSpec } from './web-spec.ts';

type Config = Record<string, any>;
export interface ConsumerOptions { id: string; fields?: Field[]; scopes: string[] | ((config: Config) => string[]) }
export const scopeNames = (scope: string) => scope.split(/\s+/).filter(Boolean);
const qualified = (scope: string, resource: string) => scope.startsWith('https://') ? scope : `${resource}/${scope}`;
function split(scope: string) { const match = scope.match(/^(https:\/\/[^/]+)\/(.+)$/); return match ? [match[1], match[2]] : ['', scope]; }
export function coversScopes(granted: string[], required: string[]) {
  return required.every(want => {
    const [resource, name] = split(want);
    return granted.some(have => {
      const [audience, permission] = split(have);
      if (audience !== resource) return false;
      if (permission === name) return true;
      const implications: Record<string, string[]> = {
        'Files.ReadWrite.All': ['Files.Read', 'Files.ReadWrite', 'Files.Read.All'], 'Files.ReadWrite': ['Files.Read'], 'Files.Read.All': ['Files.Read'],
        'Mail.ReadWrite': ['Mail.Read'], 'Mail.ReadWrite.All': ['Mail.ReadWrite', 'Mail.Read'], 'Mail.ReadWrite.Shared': ['Mail.ReadWrite', 'Mail.Read'],
        'Calendars.ReadWrite': ['Calendars.Read'], 'Calendars.ReadWrite.Shared': ['Calendars.ReadWrite', 'Calendars.Read'], 'Calendars.ReadWrite.All': ['Calendars.ReadWrite', 'Calendars.Read'],
        'Tasks.ReadWrite': ['Tasks.Read'], 'Tasks.ReadWrite.Shared': ['Tasks.ReadWrite', 'Tasks.Read'], 'Contacts.ReadWrite': ['Contacts.Read'],
        'Sites.ReadWrite.All': ['Sites.Read.All'],
        'User.ReadBasic.All': ['User.Read'], 'User.Read.All': ['User.Read', 'User.ReadBasic.All'], 'User.ReadWrite.All': ['User.Read', 'User.ReadBasic.All', 'User.Read.All'],
      };
      return implications[permission]?.includes(name) === true;
    });
  });
}
export function grantedScopes(bundle: Bundle) {
  return [...scopeNames(bundle.scope).map(scope => qualified(scope, RESOURCES.graph)), ...scopeNames(bundle.resources?.outlook?.scope ?? '').map(scope => qualified(scope, RESOURCES.outlook))];
}
export function backend(conn: Connection) { return conn.methodId === 'authentication-source' && conn.config.backend === 'outlook' ? 'outlook' : 'graph'; }
export function baseUrl(conn: Connection) { return backend(conn) === 'outlook' ? 'https://outlook.office.com/api/v2.0' : GRAPH; }

/** Resource boundaries also apply to broad, independently imported web grants. */
export function allowRequest(service: string, req: { method: string; url: URL }, conn: Connection, required: string[]) {
  if (req.url.origin !== (backend(conn) === 'outlook' ? RESOURCES.outlook : RESOURCES.graph) || req.url.username || req.url.password) return false;
  const prefix = backend(conn) === 'outlook' ? '/api/v2.0' : '/v1.0';
  const p = req.url.pathname.slice(prefix.length);
  if (!req.url.pathname.startsWith(prefix + '/') || /(?:^|\/)\$(?:batch|ref)(?:\/|$)/i.test(p)) return false;
  try { if (decodeURIComponent(p).split('/').some(x => x === '.' || x === '..' || x.includes('\\'))) return false; } catch { return false; }
  const read = ['GET', 'HEAD'].includes(req.method.toUpperCase());
  if (conn.config.access === 'read' && !read) return false;
  if (service === 'outlook-mail') {
    if (!/^\/me\/(messages|mailfolders|sendmail)(\/|$)/i.test(p)) return false;
    if (conn.config.access === 'send' && !read && !(req.method === 'POST' && (/^\/me\/sendmail$/i.test(p) || /^\/me\/messages\/[^/]+\/(reply|replyall|forward|send)$/i.test(p)))) return false;
    return true;
  }
  if (service === 'outlook-calendar') return /^\/me\/(events|calendar|calendars|calendarview|calendargroups|findmeetingtimes)(\/|$)/i.test(p);
  if (service === 'microsoft-todo') return backend(conn) === 'outlook' ? /^\/me\/(tasks|taskfolders|taskgroups)(\/|$)/i.test(p) : /^\/me\/todo(?:\/|$)/i.test(p);
  if (service === 'onedrive' || service === 'excel') {
    if (/^\/me\/drive(?:\/|$|:)/i.test(p)) return true;
    return conn.config.access === 'all' && /^\/(drives|sites|shares)(\/|$)/i.test(p);
  }
  if (service === 'microsoft-graph') {
    if (p === '/me') return read;
    const permissions = required.filter(x => x.startsWith(RESOURCES.graph + '/')).map(x => split(x)[1]);
    const candidates: [RegExp, string[]][] = [
      [/^\/me\/(messages|mailfolders|sendmail)(\/|$)/i, ['Mail.']], [/^\/me\/(events|calendar|calendars|calendarview)(\/|$)/i, ['Calendars.']],
      [/^\/me\/todo(?:\/|$)/i, ['Tasks.']], [/^\/me\/drive(?:\/|$|:)/i, ['Files.']], [/^\/(drives|shares)(\/|$)/i, ['Files.ReadWrite.All', 'Files.Read.All']],
      [/^\/sites(?:\/|$)/i, ['Sites.']], [/^\/me\/(contacts|contactfolders)(\/|$)/i, ['Contacts.']], [/^\/me\/people(?:\/|$)/i, ['People.']],
      [/^\/users(?:\/|$)/i, ['User.Read.All', 'User.ReadWrite.All', 'Directory.']], [/^\/groups(?:\/|$)/i, ['Group.']],
    ];
    const candidatesForPath = candidates.find(([pattern]) => pattern.test(p))?.[1] ?? [];
    const selected = permissions.filter(permission => candidatesForPath.some(root => root.endsWith('.') ? permission.startsWith(root) : permission === root));
    if (/^\/me\/(messages|mailfolders|sendmail)(\/|$)/i.test(p)) {
      const sending = req.method === 'POST' && (/^\/me\/sendmail$/i.test(p) || /^\/me\/messages\/[^/]+\/(reply|replyall|forward|send)$/i.test(p));
      return selected.some(permission => sending ? permission === 'Mail.Send' : read ? /Read/.test(permission) : /ReadWrite/.test(permission));
    }
    return selected.some(permission => read ? /Read/.test(permission) : /ReadWrite/.test(permission));
  }
  return false;
}

export function createConsumerAuth(ctx: PluginContext, options: ConsumerOptions): AuthMethod {
  const scopes = (c: Config) => typeof options.scopes === 'function' ? options.scopes(c) : options.scopes;
  const required = (c: Config) => [qualified('User.Read', RESOURCES.graph), ...scopes(c).map(scope => qualified(scope, c.backend === 'outlook' ? RESOURCES.outlook : RESOURCES.graph))];
  const fields = (options.fields ?? []).map(field => options.id === 'microsoft-graph' && field.key === 'scopes' ? { ...field, default: 'User.Read Files.ReadWrite.All', description: 'Graph permissions already granted to the web session. This method never requests new consent.' } : field);
  const policy = (req: Parameters<AuthMethod['authorize']>[0], conn: Connection) => allowRequest(options.id, req, conn, required(conn.config));
  async function obtain(config: Config, forceAuthentication: boolean, previous?: Bundle, signal?: AbortSignal) {
    if (!ctx.credentials) throw fail('core_upgrade_required', 'Upgrade Switchboard to use authentication sources.');
    const provided = await ctx.credentials.obtain(config.authenticationSource, { format: FORMAT, forceAuthentication, signal });
    if (provided.format !== FORMAT) throw fail('invalid_credentials', 'Incompatible Microsoft authentication source.');
    const bundle = validateBundle(provided.bundle, previous);
    const account = await identify(bundle, { signal });
    if (account.id !== provided.account.id) throw fail('account_mismatch', 'The authentication source account does not match Microsoft Graph.', 403);
    return bundle;
  }
  async function usable(conn: Connection, opts: { force?: boolean; signal?: AbortSignal; persistCredentials?: (value: any) => void | Promise<void> }) {
    let bundle = validateBundle(conn.credentials);
    const resource = backend(conn);
    const token = resource === 'graph' ? bundle : bundle.resources?.outlook;
    if (!opts.force && token && fresh(token as Bundle)) return bundle;
    try { bundle = await refresh(bundle, { resource, signal: opts.signal, onRotate: async next => { await opts.persistCredentials?.(next); } }); }
    catch (error: any) {
      if (!error.reauthenticationRequired) throw error;
      bundle = await obtain(conn.config, true, bundle, opts.signal);
      if (resource === 'outlook') bundle = await refresh(bundle, { resource, signal: opts.signal, onRotate: async next => { await opts.persistCredentials?.(next); } });
    }
    if (!coversScopes(grantedScopes(bundle), required(conn.config))) throw fail('permission_unavailable', 'The web session no longer grants the permissions required by this connection.', 403);
    return bundle;
  }
  const changed = (a: Bundle, b: Bundle) => a.accessToken !== b.accessToken || a.refreshToken !== b.refreshToken || a.scope !== b.scope || a.expiresAt !== b.expiresAt || a.generation !== b.generation || a.resources?.outlook?.accessToken !== b.resources?.outlook?.accessToken || a.resources?.outlook?.scope !== b.resources?.outlook?.scope || a.resources?.outlook?.expiresAt !== b.resources?.outlook?.expiresAt;
  return {
    id: 'authentication-source', name: 'Use Microsoft Authentication',
    description: 'Use cached web credentials from this device or an explicitly shared authentication source on a peer.',
    unavailable: ctx.credentials ? undefined : 'Upgrade Switchboard to enable authentication sources',
    fields: [...fields, { key: 'authenticationSource', label: 'Authentication connection', type: 'credential-source', credentialFormat: FORMAT, required: true }],
    credentialConsumer: { format: FORMAT, field: 'authenticationSource' },
    reusable: { kind: 'oauth', format: FORMAT, context: c => JSON.stringify([FORMAT, c.authenticationSource]), configKeys: ['authenticationSource'],
      requiredScopes: required, grantedScopes, coversScopes, validate: conn => identify(conn.credentials), allowRequest: policy,
      applyToken: (req, _conn, token) => { req.headers.set('authorization', `Bearer ${token.accessToken}`); } },
    async connect({ config, connection }) {
      let bundle = await obtain(config, false, connection?.methodId === 'authentication-source' ? connection.credentials : undefined);
      const settings = { ...config, backend: 'graph' };
      if (['outlook-mail', 'outlook-calendar', 'microsoft-todo'].includes(options.id) && !coversScopes(grantedScopes(bundle), required(settings))) {
        settings.backend = 'outlook';
        if (!bundle.resources?.outlook || !fresh(bundle.resources.outlook as Bundle)) bundle = await refresh(bundle, { resource: 'outlook' });
      }
      if (!coversScopes(grantedScopes(bundle), required(settings))) throw fail('permission_unavailable', 'These permissions are not granted to the existing web session. Choose an available access level or use another authentication method.', 403);
      return { credentials: bundle, config: settings, account: await identify(bundle) };
    },
    async authorize(req, conn, opts) {
      if (!policy(req, conn)) throw fail('request_forbidden', 'Request exceeds this Microsoft connection’s access boundary.', 403);
      const credentials = await usable(conn, opts);
      req.headers.set('authorization', `Bearer ${backend(conn) === 'outlook' ? credentials.resources!.outlook!.accessToken : credentials.accessToken}`);
      return changed(credentials, conn.credentials) ? { credentials } : undefined;
    },
    async token(conn, opts) {
      const credentials = await usable(conn, opts);
      const token = backend(conn) === 'outlook' ? credentials.resources!.outlook! : credentials;
      return { accessToken: token.accessToken, tokenType: 'Bearer', expiresAt: token.expiresAt, credentials: changed(credentials, conn.credentials) ? credentials : undefined };
    },
  };
}
export { outlookSpec };

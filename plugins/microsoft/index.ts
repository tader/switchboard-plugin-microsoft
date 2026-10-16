import { createContentPlugin } from './content/service.js';
import fs from 'node:fs';
import path from 'node:path';
import type { Field, PluginContext, ServiceDefinition } from './types/api.d.ts';
import type { AuthenticationExports } from './types/authentication.d.ts';
import type * as OAuth2 from './types/oauth2.d.ts';

type Cfg = Record<string, any>;

export interface MicrosoftServiceOptions {
  id: string;
  name: string;
  description?: string;
  icon?: string;
  docsUrl?: string;
  /** Microsoft Graph permissions, e.g. Mail.ReadWrite. */
  scopes: string[] | ((config: Cfg) => string[]);
  fields?: Field[];
  /** Name of a slim API description in openapi/, built by openapi/build.ts. */
  spec?: string;
}

const GRAPH = 'https://graph.microsoft.com/v1.0';
// offline_access gives a refresh token; User.Read lets the hub see who signed in.
const IDENTITY = ['offline_access', 'openid', 'profile', 'email', 'User.Read'];

export default function setup(ctx: PluginContext) {
  const authentication = ctx.require<AuthenticationExports>('microsoft-authentication');
  const oauth = ctx.require<typeof OAuth2>('oauth2');
  const shared = !!ctx.settings.clientId;
  const tenantOf = (c: Cfg) => String(c.tenant || ctx.settings.tenant || 'common').trim();
  const login = (c: Cfg, endpoint: string) => `https://login.microsoftonline.com/${encodeURIComponent(tenantOf(c))}/oauth2/v2.0/${endpoint}`;
  const clientId = (c: Cfg) => c.clientId || ctx.settings.clientId;

  // Users can bring their own app registration; it is required when the administrator did not configure one.
  const clientIdField: Field = {
    key: 'clientId',
    label: 'Application (client) ID',
    required: !shared,
    advanced: shared,
    description: shared ? 'Leave empty to use Switchboard’s app registration' : 'From an app registration in the Microsoft Entra admin center; see the setup guide',
  };
  const tenantField: Field = { key: 'tenant', label: 'Tenant', advanced: true, placeholder: ctx.settings.tenant || 'common', description: 'A tenant ID or domain, to sign in to one organization only' };

  const specs = new Map<string, object>();
  const spec = (name: string) => {
    if (!specs.has(name)) specs.set(name, JSON.parse(fs.readFileSync(path.join(ctx.dir, 'openapi', `${name}.json`), 'utf8')));
    return specs.get(name)!;
  };

  async function identify(creds: { accessToken: string }) {
    const res = await fetch(`${GRAPH}/me?$select=id,displayName,mail,userPrincipalName`, { headers: { authorization: `Bearer ${creds.accessToken}` } });
    if (!res.ok) return undefined;
    const me = await res.json();
    return { id: me.id, label: me.mail || me.userPrincipalName || me.displayName };
  }

  function microsoftService(o: MicrosoftServiceOptions): ServiceDefinition {
    const scopes = (c: Cfg) => [...IDENTITY, ...(typeof o.scopes === 'function' ? o.scopes(c) : o.scopes)];
    const common = { tokenUrl: (c: Cfg) => login(c, 'token'), clientId, scopes, identify };
    return {
      id: o.id,
      name: o.name,
      description: o.description,
      icon: o.icon,
      docsUrl: o.docsUrl,
      baseUrl: (conn) => authentication.baseUrl(conn),
      allowedHosts: (conn) => [authentication.backend(conn) === 'outlook' ? 'outlook.office.com' : 'graph.microsoft.com'],
      openapi: o.spec ? (conn) => authentication.backend(conn) === 'outlook' ? authentication.outlookSpec(o.id) : spec(o.spec!) : undefined,
      authMethods: [
        authentication.createConsumerAuth(ctx, o),
        oauth.authorizationCode({
          ...common,
          id: 'oauth',
          name: 'Sign in with Microsoft',
          fields: [...(o.fields ?? []), clientIdField, { key: 'clientSecret', label: 'Client secret', type: 'secret', advanced: true, description: 'Only with your own app registration' }, tenantField],
          clientSecret: (c) => (c.clientId ? c.clientSecret : ctx.settings.clientSecret),
          authorizeUrl: (c) => login(c, 'authorize'),
          authorizeParams: { prompt: 'select_account', response_mode: 'query' },
        }),
        oauth.deviceCode({
          ...common,
          id: 'device',
          name: 'Sign in with a code',
          description: 'Enter a code at microsoft.com/devicelogin',
          fields: [...(o.fields ?? []), clientIdField, tenantField],
          // Microsoft refuses a client secret in the device flow: the app must allow public client flows.
          clientSecret: () => undefined,
          deviceAuthorizationUrl: (c) => login(c, 'devicecode'),
        }),
      ],
    };
  }

  const graph = microsoftService({
    id: 'microsoft-graph',
    name: 'Microsoft Graph',
    description: 'Microsoft 365 APIs, with the permissions you choose',
    icon: 'icon.svg',
    docsUrl: 'https://learn.microsoft.com/graph/api/overview',
    spec: 'graph',
    scopes: (c) => String(c.scopes ?? '').split(/[\s,]+/).filter(Boolean),
    fields: [
      {
        key: 'scopes',
        label: 'Permissions',
        type: 'textarea',
        required: true,
        default: 'Mail.ReadWrite Mail.Send Calendars.ReadWrite Contacts.ReadWrite Files.ReadWrite Tasks.ReadWrite',
        description: 'Microsoft Graph delegated permissions, separated by spaces. The app registration must include them.',
      },
    ],
  });

  return { services: [graph], exports: { microsoftService, spec, createContentPlugin } };
}

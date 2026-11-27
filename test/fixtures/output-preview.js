// Synthetic wire responses based on Microsoft Graph's drive/calendar resources.
import microsoft from '../../plugins/microsoft/index.ts';
import onedrive from '../../plugins/onedrive/index.ts';
import calendar from '../../plugins/outlook-calendar/index.ts';
import * as authentication from '../../plugins/microsoft-authentication/consumer.ts';
import { FORMAT, TEAMS_CLIENT_ID } from '../../plugins/microsoft-authentication/lib/session.js';
export const credentials = () => ({
  format: FORMAT, clientId: TEAMS_CLIENT_ID, origin: 'https://teams.microsoft.com',
  identity: { tenant: '11111111-1111-4111-8111-111111111111', oid: '22222222-2222-4222-8222-222222222222' },
  accessToken: 'fixture-access', refreshToken: 'fixture-refresh', expiresAt: Date.now() + 3600_000,
  scope: 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/Calendars.ReadWrite',
  generation: '33333333-3333-4333-8333-333333333333',
  resources: { outlook: { accessToken: 'fixture-outlook', expiresAt: Date.now() + 3600_000, scope: 'Calendars.ReadWrite' } },
});
export const drive = { driveType: 'business', id: 'private-drive', name: 'Private files', owner: { user: { email: 'private@example.test' } }, nested: { token: 'fixture-access' } };
export const calendarBody = { defaultOnlineMeetingProvider: 'teamsForBusiness', id: 'private-calendar', name: 'Private calendar', owner: { address: 'private@example.test' }, nested: { token: 'fixture-access' } };
export function graphServices() {
  // OAuth protocol is core-owned; the fixture only applies the synthetic bearer.
  const oauth = options => ({ ...options, fields: options.fields ?? [], connect: async () => ({ credentials: credentials() }), authorize: async req => req.headers.set('authorization', 'Bearer fixture-access') });
  const hub = microsoft({ settings: {}, require: id => id === 'microsoft-authentication' ? authentication : { authorizationCode: oauth, deviceCode: oauth } });
  const ctx = { require: () => hub.exports };
  return [hub.services[0], onedrive(ctx).services[0], calendar(ctx).services[0]];
}
export const config = { access: 'full', scopes: 'Files.ReadWrite', backend: 'graph' };
export const contentContext = { require: () => ({ createConsumerAuth: () => ({ id: 'authentication-source', fields: [], token: async () => ({ accessToken: 'fixture-access' }) }) }) };

import { generateKeyPairSync, sign } from 'node:crypto';
import { resources, clientId } from '../lib/oauth.js';

export const renewal = () => ({ value: 'fixture-refresh-secret', clientId, origin: 'https://teams.microsoft.com', capturedAt: Date.now(), sourceGrant: 'authorization_code' });

export const tenant = '11111111-1111-1111-1111-111111111111';
export const oid = '22222222-2222-2222-2222-222222222222';
const { publicKey, privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const issued = new Set();
export const serverAccepts = token => issued.has(token);
const key = { ...publicKey.export({ format: 'jwk' }), kid: 'import-test', use: 'sig' };
export function signedToken(service, changes = {}, headerChanges = {}) {
  const claims = { tid: tenant, oid, aud: resources[service], iss: `https://sts.windows.net/${tenant}/`, exp: Math.floor(Date.now() / 1000) + 3600, upn: 'test@example.com', ...changes };
  const encoded = [{ alg: 'RS256', kid: key.kid, ...headerChanges }, claims].map(value => Buffer.from(JSON.stringify(value)).toString('base64url')).join('.');
  const token = encoded + '.' + sign('RSA-SHA256', Buffer.from(encoded), privateKey).toString('base64url');
  issued.add(token); return token;
}
export function proprietaryToken(service, changes = {}) {
  const ordinary = signedToken(service, changes, { nonce: 'proprietary-header-nonce' });
  const token = ordinary.split('.').slice(0, 2).join('.') + '.proprietary-service-signature';
  issued.add(token); return token;
}
export function signedBundle(changes = {}) {
  const tokens = Object.fromEntries(Object.keys(resources).map(service => [service, { value: signedToken(service, changes), expiresAt: Date.now() + 3600_000 }]));
  return { authVersion: 1, identity: { tenant: changes.tid ?? tenant, oid: changes.oid ?? oid, email: 'test@example.com', name: 'Test user' }, tokens };
}

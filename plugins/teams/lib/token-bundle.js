import { AdapterError } from './errors.js';
import { fresh, importRequired } from './api-auth.js';
import { resources, clientId } from './oauth.js';
import { tokenClaims } from './api-http.js';
import { recipientEmail } from './input.js';

const format = 'switchboard-teams-tokens';
const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const invalid = (reason = 'bundle_schema') => new AdapterError('invalid_token_bundle', `Invalid Teams token bundle (${reason}). Paste the complete JSON file from auth:export.`, 400);
function object(value, keys) {
  return value && typeof value === 'object' && !Array.isArray(value) && Object.keys(value).every(key => keys.includes(key));
}
function shape(value) {
  if (!object(value, ['format', 'version', 'identity', 'tokens', 'renewal']) || value.format !== format || ![1, 2].includes(value.version) ||
      !object(value.identity, ['tenant', 'oid', 'email', 'name']) || !uuid.test(value.identity.tenant ?? '') || !uuid.test(value.identity.oid ?? '') ||
      typeof value.identity.email !== 'string' || typeof value.identity.name !== 'string' || value.identity.name.length > 200 ||
      !object(value.tokens, ['skype', 'chatsvcagg', 'substrate'])) throw invalid();
  if (value.version === 1 && value.renewal !== undefined) throw invalid();
  if (value.version === 2) {
    const r = value.renewal;
    if (!object(r, ['value', 'clientId', 'origin', 'capturedAt', 'sourceGrant']) || typeof r.value !== 'string' || !r.value || r.value.length > 100_000 ||
        r.clientId !== clientId || !['https://teams.microsoft.com', 'https://teams.cloud.microsoft'].includes(r.origin) ||
        !Number.isFinite(r.capturedAt) || r.capturedAt <= 0 || r.capturedAt > Date.now() + 60_000 || !['authorization_code', 'refresh_token'].includes(r.sourceGrant)) throw invalid('renewal_metadata');
  }
  try { recipientEmail(value.identity.email); } catch { throw invalid(); }
  for (const service of Object.keys({ ...resources, ...(value.tokens.substrate !== undefined ? { substrate: 'https://substrate.office.com' } : {}) })) {
    const token = value.tokens[service];
    if (!object(token, ['value', 'expiresAt']) || typeof token.value !== 'string' || token.value.length > 100_000 || !Number.isFinite(token.expiresAt)) throw invalid();
    if (!token.value) throw invalid();
    if (value.version === 1 && !fresh(token)) throw importRequired();
  }
  return value;
}

export function encodeTokenBundle(credentials) {
  const { tenant, oid, email, name } = credentials.identity;
  const tokens = Object.fromEntries(Object.keys({ ...resources, ...(credentials.tokens.substrate ? { substrate: 'https://substrate.office.com' } : {}) }).map(service => [service, { value: credentials.tokens[service].value, expiresAt: credentials.tokens[service].expiresAt }]));
  const encoded = JSON.stringify(shape({ format, version: credentials.renewal ? 2 : 1, identity: { tenant, oid, email, name }, tokens,
    ...(credentials.renewal ? { renewal: { ...credentials.renewal } } : {}) }));
  if (Buffer.byteLength(encoded) > 128 * 1024) throw invalid();
  return encoded;
}

export async function validateTokenBundle(raw, previous) {
  if (typeof raw !== 'string' || Buffer.byteLength(raw) > 128 * 1024) throw invalid('bundle_size');
  let value;
  try { value = JSON.parse(raw); } catch { throw invalid('invalid_json'); }
  shape(value);
  const identity = { ...value.identity, email: recipientEmail(value.identity.email) };
  const tokens = {};
  for (const service of Object.keys({ ...resources, ...(value.tokens.substrate !== undefined ? { substrate: 'https://substrate.office.com' } : {}) })) {
    // Consistency checks only: Microsoft access tokens can use proprietary
    // signatures. Their receiving Teams APIs validate authenticity at connect.
    const claims = tokenClaims(value.tokens[service].value);
    if (claims.aud !== (resources[service] ?? 'https://substrate.office.com')) throw invalid(`${service}: audience_mismatch`);
    if (claims.tid !== identity.tenant) throw invalid(`${service}: tenant_mismatch`);
    if (!Number.isFinite(claims.exp) || claims.nbf !== undefined && (!Number.isFinite(claims.nbf) || claims.nbf > Date.now() / 1000 + 60)) throw invalid(`${service}: invalid_expiry`);
    if (claims.oid !== identity.oid) throw new AdapterError('account_mismatch', 'All tokens must belong to the same Teams account.', 403);
    tokens[service] = { value: value.tokens[service].value, expiresAt: Math.min(value.tokens[service].expiresAt, claims.exp * 1000) };
    if (value.version === 1 && !fresh(tokens[service])) throw importRequired();
  }
  if (previous?.identity && (identity.tenant !== previous.identity.tenant || identity.oid !== previous.identity.oid)) {
    throw new AdapterError('account_mismatch', 'Import tokens for the same Teams account as this connection.', 403);
  }
  return { authVersion: 1, authMode: 'tokens', identity, tokens, ...(value.version === 2 ? { renewal: { ...value.renewal } } : {}) };
}

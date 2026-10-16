// Read-only validation through the consolidated Teams adapter and shared credential source.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DiagnosticSession, FORMAT, TEAMS_CLIENT_ID, RESOURCES, validateBundle, identify, refresh, fresh, fail } from '../plugins/microsoft-authentication/lib/session.js';
import setup from '../plugins/teams/index.js';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const session = new DiagnosticSession(process.env.MICROSOFT_AUTH_CACHE || path.join(root, '.local-profile/microsoft'));
const scratch = await fs.mkdtemp(path.join(os.tmpdir(), 'microsoft-teams-check-'));
let adapter;
try {
  await session.acquire({ silent: process.argv.includes('--silent') });
  const webSession = { format: FORMAT, teamsClientId: TEAMS_CLIENT_ID, resources: RESOURCES, validateBundle, identify, refresh, fresh, fail };
  adapter = await setup({ dataDir: scratch, settings: {}, require: () => ({ webSession }), credentials: { obtain: async () => ({ format: FORMAT, bundle: session.credentials, account: await identify(session.credentials) }) } });
  const service = adapter.services[0], method = service.authMethods.find(method => method.id === 'authentication-source');
  const result = await method.connect({ config: { authenticationSource: 'diagnostic' } });
  let conn = { id: 'diagnostic', methodId: method.id, config: result.config, credentials: result.credentials };
  await session.save(conn.credentials.microsoft);
  const paths = ['/status', '/chats', '/channels', '/triage/inbox?limit=1&contextLimit=1'];
  for (const relative of paths) {
    const request = { method: 'GET', url: new URL('http://teams.localhost' + relative), headers: new Headers() };
    const authorized = await method.authorize(request, conn, { persistCredentials: async credentials => { conn.credentials = credentials; await session.save(credentials.microsoft); } });
    if (authorized?.credentials) { conn.credentials = authorized.credentials; await session.save(authorized.credentials.microsoft); }
    const response = await fetch(request.url, { method: request.method, headers: request.headers, redirect: 'error', signal: AbortSignal.timeout(60_000) });
    if (!response.ok) { await response.body?.cancel(); throw fail('teams_check_failed', `Teams ${relative.split('?')[0]} returned HTTP ${response.status}.`); }
    await response.body?.cancel();
    console.log(`PASS: shared-source Teams ${relative.split('?')[0]}`);
  }
  console.log('No messages sent or read states changed; shared credentials cached.');
} catch (error) { console.error(`${error.code || 'validation_failed'}: Teams validation did not complete.`); process.exitCode = 1; }
finally { await adapter?.dispose(); await session.close(); await fs.rm(scratch, {recursive:true,force:true}); }

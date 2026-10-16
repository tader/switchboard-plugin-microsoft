import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DiagnosticSession } from '../lib/diagnostic-session.js';
import { TeamsAPI } from '../lib/teams-api.js';
import { AdapterError } from '../lib/errors.js';

const silent = process.argv.includes('--silent');
const directory = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../.local-profile/token-export');
const auth = new DiagnosticSession(directory, { browserPath: process.env.TEAMS_BROWSER_PATH });
process.once('SIGINT', () => { auth.close().finally(() => { process.exitCode = 130; }); });
try {
  console.log('Reusing the shared diagnostic session; expired tokens renew over HTTP.');
  const credentials = await auth.acquire({ silent, capture: process.argv.includes('--capture'), loginHint: process.env.TEAMS_LOGIN_HINT });
  console.log('PASS: diagnostic session ready; authentication browser closed');
  const api = new TeamsAPI({ credentials });
  await api.profile(process.env.TEAMS_LOGIN_HINT || credentials.identity.email);
  console.log('PASS: direct API profile read');
  const chats = await api.chats();
  console.log('PASS: direct API conversation list');
  if (chats.items.length) {
    await api.messages(chats.items[0].id, 5);
    console.log('PASS: direct API recent messages read');
  } else console.log('SKIP: recent messages (no chats returned)');
  console.log('Earliest token expiry:', new Date(Math.min(credentials.tokens.skype.expiresAt, credentials.tokens.chatsvcagg.expiresAt, credentials.chatToken.expiresAt)).toISOString());
  console.log('API check complete. No messages were sent. The shared diagnostic session is saved locally for reuse.');
} catch (error) {
  console.error('API check failed:', error instanceof AdapterError ? `${error.code}: ${error.message}` : 'runtime_or_network_failure');
  process.exitCode = 1;
} finally { await auth.close(); }

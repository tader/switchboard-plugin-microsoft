import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { BrowserAuthenticator } from '../lib/api-auth.js';
import { encodeTokenBundle } from '../lib/token-bundle.js';
import { recipientEmail } from '../lib/input.js';
import { AdapterError } from '../lib/errors.js';

const directory = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../.local-profile/token-export');

export async function exportTokens({ out, silent = false, loginHint, signal }, {
  createAuthenticator = () => new BrowserAuthenticator(directory, { browserPath: process.env.TEAMS_BROWSER_PATH }),
} = {}) {
  if (typeof out !== 'string' || !out.trim()) throw new AdapterError('export_path_required', 'Provide --out with a new output filename.', 400);
  if (loginHint) loginHint = recipientEmail(loginHint);
  signal?.throwIfAborted();
  const filename = path.resolve(out);
  let file;
  try { file = await fs.open(filename, 'wx', 0o600); }
  catch { throw new AdapterError('export_file_unavailable', 'Cannot create the output file. Use a new filename in an existing writable directory.', 400); }
  let auth, complete = false;
  const abort = () => { auth?.close().catch(() => {}); };
  signal?.addEventListener('abort', abort, { once: true });
  try {
    await file.chmod(0o600);
    signal?.throwIfAborted();
    auth = createAuthenticator();
    const credentials = await auth.acquire({ interactive: !silent, loginHint });
    signal?.throwIfAborted();
    const encoded = encodeTokenBundle(credentials);
    await file.writeFile(encoded + '\n');
    await file.sync();
    signal?.throwIfAborted();
    await auth.close(); auth = null;
    await file.close(); file = null;
    signal?.throwIfAborted();
    complete = true;
    return { expiresAt: Math.min(credentials.tokens.skype.expiresAt, credentials.tokens.chatsvcagg.expiresAt) };
  } finally {
    signal?.removeEventListener('abort', abort);
    try { await auth?.close(); }
    finally { try { await file?.close(); } finally { if (!complete) await fs.unlink(filename); } }
  }
}

async function main() {
  const controller = new AbortController();
  const interrupt = () => controller.abort();
  process.once('SIGINT', interrupt); process.once('SIGTERM', interrupt);
  try {
    let values;
    try { ({ values } = parseArgs({ options: { out: { type: 'string' }, silent: { type: 'boolean' }, 'login-hint': { type: 'string' } }, strict: true, allowPositionals: false })); }
    catch { throw new AdapterError('export_arguments_invalid', 'Usage: npm run auth:export -- --out <new-file> [--silent] [--login-hint <email>]', 400); }
    if (!values.out) throw new AdapterError('export_path_required', 'Provide --out with a new output filename.', 400);
    console.log(values.silent ? 'Obtaining tokens using the saved local sign-in session.' : 'Complete Microsoft sign-in in the dedicated browser window.');
    const result = await exportTokens({ out: values.out, silent: values.silent, loginHint: values['login-hint'] || process.env.TEAMS_LOGIN_HINT, signal: controller.signal });
    console.log('Bundle saved with owner-only permissions. Chrome closed.');
    console.log('Earliest Microsoft access-token expiry:', new Date(result.expiresAt).toISOString());
    console.log('Paste the file contents into Import Teams tokens on your central Switchboard. Delete the file after import.');
  } catch (error) {
    console.error(controller.signal.aborted ? 'Token export cancelled.' : error instanceof AdapterError ? `${error.code}: ${error.message}` : 'Token export failed. No bundle was saved.');
    process.exitCode = controller.signal.aborted ? 130 : 1;
  } finally { process.removeListener('SIGINT', interrupt); process.removeListener('SIGTERM', interrupt); }
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) await main();

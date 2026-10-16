import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';

// Only plugin code is replaced. Connection credentials and plugin-data stay in place.
const source = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dataDir = path.resolve(process.argv[2] || path.join(os.homedir(), 'switchboard/.data'));
const target = path.join(dataDir, 'plugins/teams-web');
const backup = path.join(dataDir, 'plugin-backups', `teams-web-${new Date().toISOString().replace(/[:.]/g, '-')}-${randomUUID()}`);
let stage, moved = false;
try {
  const manifest = JSON.parse(await fs.readFile(path.join(source, 'plugin.json'), 'utf8'));
  if (manifest.id !== 'teams-web' || manifest.main !== 'index.js') throw new Error('Unexpected plugin manifest.');
  const existing = await fs.lstat(target).catch(error => { if (error.code !== 'ENOENT') throw error; return null; });
  if (existing && (!existing.isDirectory() || existing.isSymbolicLink())) throw new Error('Installed plugin must be a regular directory.');
  await fs.mkdir(path.join(dataDir, 'plugin-backups'), { recursive: true });
  // Stage outside the watched plugins directory so partial copies are never loaded.
  stage = await fs.mkdtemp(path.join(dataDir, '.teams-install-'));
  for (const entry of ['plugin.json', 'package.json', 'index.js', 'icon.svg', 'README.md', 'THIRD_PARTY.md', 'lib', 'docs']) {
    await fs.cp(path.join(source, entry), path.join(stage, entry), { recursive: true });
  }
  await fs.mkdir(path.dirname(target), { recursive: true });
  if (existing) { await fs.rename(target, backup); moved = true; }
  try { await fs.rename(stage, target); stage = null; }
  catch (error) { if (moved) { await fs.rename(backup, target); moved = false; } throw error; }
  console.log(`Installed Teams ${manifest.version} in ${target}`);
  if (moved) console.log(`Previous plugin saved in ${backup}`);
  console.log('Reload or restart Switchboard, then reconnect the Teams connection once.');
} catch (error) {
  console.error(`Install failed: ${error.message}`);
  process.exitCode = 1;
} finally { if (stage) await fs.rm(stage, { recursive: true, force: true }); }

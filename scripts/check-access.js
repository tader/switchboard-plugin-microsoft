import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DiagnosticSession, FORMAT, identify } from '../plugins/microsoft-authentication/lib/session.js';
import { createConsumerAuth, baseUrl, backend } from '../plugins/microsoft-authentication/consumer.ts';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = process.env.MICROSOFT_AUTH_CACHE || path.join(root, '.local-profile', 'microsoft');
const session = new DiagnosticSession(directory, { browserPath: process.env.MICROSOFT_BROWSER_PATH });
process.once('SIGINT', () => session.controller.abort());
try {
  console.log('Reusing cached Microsoft credentials; HTTP renewal and silent authentication precede interactive sign-in.');
  await session.acquire({ silent: process.argv.includes('--silent'), capture: process.argv.includes('--capture') });
  const account = await identify(session.credentials);
  console.log('PASS: Microsoft Graph identity; credentials cached and browser closed');
  const context = { credentials: { obtain: async () => ({ format: FORMAT, bundle: session.credentials, account }) } };
  let selected = {};
  try { selected = JSON.parse(await fs.readFile(path.join(directory, 'selection.json'), 'utf8')); } catch {}
  const site = process.env.MICROSOFT_TEST_SITE || selected.site;
  const drive = process.env.MICROSOFT_TEST_DRIVE || selected.workbook?.drive, item = process.env.MICROSOFT_TEST_WORKBOOK || selected.workbook?.item;
  const probes = [
    ['outlook-mail', ['Mail.ReadWrite', 'Mail.Send'], '/me/messages?$top=1&$select=Id', 'full'],
    ['outlook-calendar', ['Calendars.ReadWrite'], '/me/calendars?$top=1&$select=Id', 'full'],
    ['microsoft-todo', ['Tasks.ReadWrite'], '/me/taskfolders?$top=1', 'full'],
    ['onedrive', ['Files.ReadWrite.All'], '/me/drive?$select=id', 'all'],
  ];
  if (site) probes.push(['onedrive', ['Files.ReadWrite.All'], `/sites/${encodeURIComponent(site)}/drives?$top=1&$select=id`, 'all']);
  if (drive && item) probes.push(['excel', ['Files.ReadWrite.All'], `/drives/${encodeURIComponent(drive)}/items/${encodeURIComponent(item)}/workbook/worksheets?$select=id`, 'all']);
  let failures = 0;
  for (const [id, scopes, endpoint, access] of probes) {
    try {
      const method = createConsumerAuth(context, { id, scopes });
      const connected = await method.connect({ config: { authenticationSource: 'diagnostic', access } });
      await session.save(connected.credentials);
      const conn = { id: 'diagnostic', methodId: method.id, config: connected.config, credentials: connected.credentials };
      const relative = backend(conn) === 'graph' && id === 'microsoft-todo' ? '/me/todo/lists?$top=1&$select=id' : backend(conn) === 'graph' ? endpoint.replace('$select=Id', '$select=id') : endpoint;
      const request = { method: 'GET', url: new URL(baseUrl(conn) + relative), headers: new Headers() };
      const updated = await method.authorize(request, conn, { persistCredentials: bundle => session.save(bundle) });
      if (updated?.credentials) await session.save(updated.credentials);
      const response = await fetch(request.url, { headers: request.headers, redirect: 'error', signal: AbortSignal.timeout(30000) });
      console.log(`${response.ok ? 'PASS' : 'UNAVAILABLE'}: ${id}${endpoint.startsWith('/sites') ? ' SharePoint' : ''} via ${backend(conn)} (HTTP ${response.status})`);
      if (!response.ok) failures++;
      await response.body?.cancel();
    } catch (error) { failures++; console.log(`UNAVAILABLE: ${id} (${error.code || 'request_failed'})`); }
  }
  if (!site) console.log('SKIP: SharePoint; set MICROSOFT_TEST_SITE to a selected site ID');
  if (!drive || !item) console.log('SKIP: workbook; set MICROSOFT_TEST_DRIVE and MICROSOFT_TEST_WORKBOOK');
  console.log('No content was written. The saved credential cache will be reused.');
  if (failures) process.exitCode = 2;
} catch (error) {
  console.error(`${error.code || 'access_check_failed'}: ${error.code ? error.message : 'The Microsoft access check failed.'}`); process.exitCode = 1;
} finally { await session.close(); }

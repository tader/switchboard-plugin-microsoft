// Explicit opt-in diagnostic. Only disposable artifacts created by this run are changed/deleted.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import assert from 'node:assert/strict';
import { DiagnosticSession, FORMAT, identify } from '../plugins/microsoft-authentication/lib/session.js';
import { createConsumerAuth, baseUrl, backend } from '../plugins/microsoft-authentication/consumer.ts';

if (!process.argv.includes('--write')) { console.error('Use --write to create, verify and delete disposable validation artifacts. No mail is sent.'); process.exit(1); }
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const directory = process.env.MICROSOFT_AUTH_CACHE || path.join(root, '.local-profile', 'microsoft');
const session = new DiagnosticSession(directory, { browserPath: process.env.MICROSOFT_BROWSER_PATH });
const cleanup = []; let cleanupFailed = false;
// Excel may retain a short-lived file lock after closeSession.
async function removeWithRetry(remove) {
  for (let attempt = 0; ; attempt++) {
    try { return await remove(); } catch (error) {
      if (error.code === 'http_404') return;
      if (error.code !== 'http_423' || attempt >= 11) throw error;
      await new Promise(resolve => setTimeout(resolve, 5000));
    }
  }
}
try {
  await session.acquire({ silent: process.argv.includes('--silent') });
  const account = await identify(session.credentials);
  const ctx = { credentials: { obtain: async () => ({ format: FORMAT, bundle: session.credentials, account }) } };
  const label = `Switchboard API validation ${randomUUID()}`;
  async function connection(id, scopes, access = 'full') {
    const method = createConsumerAuth(ctx, { id, scopes });
    const result = await method.connect({ config: { access, authenticationSource: 'diagnostic' } });
    await session.save(result.credentials);
    return { method, conn: { id: 'diagnostic', methodId: method.id, config: result.config, credentials: result.credentials } };
  }
  async function request(target, method, relative, body, extraHeaders = {}) {
    target.conn.credentials = session.credentials;
    const req = { method, url: new URL(baseUrl(target.conn) + relative), headers: new Headers(extraHeaders) };
    const authorized = await target.method.authorize(req, target.conn, { persistCredentials: value => session.save(value) });
    if (authorized?.credentials) await session.save(authorized.credentials);
    if (body !== undefined && !(body instanceof Uint8Array)) req.headers.set('content-type', 'application/json');
    const response = await fetch(req.url, { method, headers: req.headers, redirect: 'error', body: body === undefined ? undefined : body instanceof Uint8Array ? body : JSON.stringify(body), signal: AbortSignal.timeout(60_000) });
    if (!response.ok) { await response.body?.cancel(); throw Object.assign(new Error(`Validation operation failed (HTTP ${response.status}).`), { code: `http_${response.status}` }); }
    if (response.status === 204 || response.status === 202) { await response.body?.cancel(); return {}; }
    return response.json();
  }
  const mail = await connection('outlook-mail', ['Mail.ReadWrite', 'Mail.Send']);
  const nativeMail = backend(mail.conn) === 'outlook';
  const draft = await request(mail, 'POST', '/me/messages', nativeMail ? { Subject: label, Body: { ContentType: 'Text', Content: 'Disposable API validation draft; never sent.' } } : { subject: label, body: { contentType: 'Text', content: 'Disposable API validation draft; never sent.' } });
  const draftPath = `/me/messages/${encodeURIComponent(draft.Id ?? draft.id)}`;
  cleanup.push(() => request(mail, 'DELETE', draftPath));
  const readDraft = await request(mail, 'GET', draftPath);
  assert.equal(readDraft.Subject ?? readDraft.subject, label);
  await request(mail, 'PATCH', draftPath, nativeMail ? { Subject: label + ' updated' } : { subject: label + ' updated' });
  console.log('PASS: create, read and edit disposable mail draft; no mail sent');

  const calendar = await connection('outlook-calendar', ['Calendars.ReadWrite']);
  const nativeCalendar = backend(calendar.conn) === 'outlook';
  const start = new Date(Date.now() + 86400_000).toISOString().slice(0, 19), end = new Date(Date.now() + 86400_000 + 1800_000).toISOString().slice(0, 19);
  const event = await request(calendar, 'POST', '/me/events', nativeCalendar ? { Subject: label, Start: { DateTime: start, TimeZone: 'UTC' }, End: { DateTime: end, TimeZone: 'UTC' }, ShowAs: 'Free', IsReminderOn: false } : { subject: label, start: { dateTime: start, timeZone: 'UTC' }, end: { dateTime: end, timeZone: 'UTC' }, showAs: 'free', isReminderOn: false });
  const eventPath = `/me/events/${encodeURIComponent(event.Id ?? event.id)}`;
  cleanup.push(() => request(calendar, 'DELETE', eventPath));
  assert.equal((await request(calendar, 'GET', eventPath))[nativeCalendar ? 'Subject' : 'subject'], label);
  await request(calendar, 'PATCH', eventPath, nativeCalendar ? { Subject: label + ' updated' } : { subject: label + ' updated' });
  console.log('PASS: create, read and edit disposable calendar event without attendees or reminders');

  const todo = await connection('microsoft-todo', ['Tasks.ReadWrite']);
  if (backend(todo.conn) === 'outlook') {
    const task = await request(todo, 'POST', '/me/tasks', { Subject: label });
    const taskPath = `/me/tasks/${encodeURIComponent(task.Id)}`;
    cleanup.push(() => request(todo, 'DELETE', taskPath));
    assert.equal((await request(todo, 'GET', taskPath)).Subject, label);
    await request(todo, 'POST', taskPath + '/complete');
    assert.equal((await request(todo, 'GET', taskPath)).Status, 'Completed');
    console.log('PASS: create, read and complete disposable Outlook task');
  } else {
    const list = await request(todo, 'POST', '/me/todo/lists', { displayName: label });
    const listPath = `/me/todo/lists/${encodeURIComponent(list.id)}`;
    cleanup.push(() => request(todo, 'DELETE', listPath));
    const task = await request(todo, 'POST', listPath + '/tasks', { title: label });
    await request(todo, 'PATCH', listPath + `/tasks/${encodeURIComponent(task.id)}`, { status: 'completed' });
    console.log('PASS: create and complete disposable Graph task');
  }

  const excel = await connection('excel', ['Files.ReadWrite.All'], 'all');
  const bytes = await fs.readFile(path.join(root, 'test', 'fixtures', 'blank-workbook.xlsx'));
  async function workbook(drive, name, folder = '') {
    const file = await request(excel, 'PUT', `${drive}/root:/${folder ? folder.split('/').map(encodeURIComponent).join('/') + '/' : ''}${encodeURIComponent(label + '.xlsx')}:/content`, bytes, { 'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const itemPath = `${drive}/items/${encodeURIComponent(file.id)}`;
    cleanup.push(() => request(excel, 'DELETE', itemPath));
    const created = await request(excel, 'POST', itemPath + '/workbook/createSession', { persistChanges: true });
    const headers = { 'workbook-session-id': created.id };
    cleanup.push(() => request(excel, 'POST', itemPath + '/workbook/closeSession', undefined, headers));
    const rangePath = itemPath + "/workbook/worksheets/Sheet1/range(address='A1:B2')";
    await request(excel, 'PATCH', rangePath, { values: [[2, 3], [4, null]], formulas: [[null, null], [null, '=SUM(A1:B1)']], numberFormat: [['0', '0'], ['0', '0']] }, headers);
    const read = await request(excel, 'GET', rangePath, undefined, headers);
    assert.deepEqual(read.values, [[2, 3], [4, 5]]);
    console.log(`PASS: ${name} workbook session, cell values, formula calculation, formatting and readback`);
  }
  await workbook('/me/drive', 'OneDrive');
  let selection = {};
  try { selection = JSON.parse(await fs.readFile(path.join(directory, 'selection.json'), 'utf8')); } catch {}
  let sharepointDrive = process.env.MICROSOFT_TEST_SHAREPOINT_DRIVE || selection.sharepointDrive;
  if (!sharepointDrive && (process.env.MICROSOFT_TEST_SITE || selection.site)) {
    const drive = await request(excel, 'GET', `/sites/${encodeURIComponent(process.env.MICROSOFT_TEST_SITE || selection.site)}/drive`);
    sharepointDrive = drive.id;
  }
  if (sharepointDrive) await workbook(`/drives/${encodeURIComponent(sharepointDrive)}`, 'SharePoint', process.env.MICROSOFT_TEST_SHAREPOINT_FOLDER || selection.sharepointFolder || '');
  else console.log('SKIP: disposable SharePoint workbook; select a writable document-library drive');
} catch (error) {
  console.error(`${error.code || 'validation_failed'}: Live write validation did not complete.`); process.exitCode = 1;
} finally {
  for (const remove of cleanup.reverse()) try { await removeWithRetry(remove); } catch (error) { cleanupFailed = true; console.error(`Cleanup failed for a disposable validation artifact (${error.code || 'request_failed'}).`); }
  await session.close();
  if (cleanupFailed) process.exitCode = 1;
  else if (cleanup.length) console.log('PASS: all disposable validation artifacts removed; credential cache retained');
}

// Explicit live diagnostic: reads Org Explorer; optional writes affect only UUID-named files created here.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import assert from 'node:assert/strict';
import {
  DiagnosticSession,
  FORMAT,
  identify,
} from '../plugins/microsoft-authentication/lib/session.js';
import { createConsumerAuth } from '../plugins/microsoft-authentication/consumer.ts';
import { createContentPlugin } from '../plugins/microsoft/content/service.js';
import { GraphClient } from '../plugins/microsoft/content/graph.js';
import { readWord } from '../plugins/microsoft/content/word.js';
import { readPowerPoint } from '../plugins/microsoft/content/powerpoint.js';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..'),
  directory =
    process.env.MICROSOFT_AUTH_CACHE ||
    path.join(root, '.local-profile/microsoft');
const session = new DiagnosticSession(directory, {
    browserPath: process.env.MICROSOFT_BROWSER_PATH,
  }),
  plugins = [],
  cleanup = [];
const manifest = path.join(directory, 'content-cleanup.json');
let failed = false;
const saveCleanup = () =>
  fs.writeFile(manifest, JSON.stringify(cleanup), { mode: 0o600 });
try {
  await session.acquire({ silent: process.argv.includes('--silent') });
  const account = await identify(session.credentials),
    ctx = {
      require: () => ({ createConsumerAuth }),
      credentials: {
        obtain: async () => ({
          format: FORMAT,
          bundle: session.credentials,
          account,
        }),
      },
    };
  async function plugin(id) {
    const p = await createContentPlugin(ctx, { id, name: id });
    plugins.push(p);
    const method = p.services[0].authMethods[0],
      result = await method.connect({
        config: { authenticationSource: 'diagnostic', access: 'all' },
      });
    await session.save(result.credentials);
    return {
      method,
      conn: {
        id,
        methodId: method.id,
        config: result.config,
        credentials: result.credentials,
      },
    };
  }
  async function call(target, method, relative, body) {
    target.conn.credentials = session.credentials;
    const request = {
      method,
      url: new URL(`http://${target.conn.id}.localhost` + relative),
      headers: new Headers(
        body === undefined ? {} : { 'content-type': 'application/json' },
      ),
    };
    const updated = await target.method.authorize(request, target.conn, {
      persistCredentials: (value) => session.save(value),
    });
    if (updated?.credentials) await session.save(updated.credentials);
    const response = await fetch(request.url, {
      method,
      headers: request.headers,
      body: body === undefined ? undefined : JSON.stringify(body),
      redirect: 'error',
      signal: AbortSignal.timeout(120_000),
    });
    const value = await response.json();
    if (!response.ok)
      throw Object.assign(Error('Content validation failed.'), {
        code: value?.error?.code ?? 'http_' + response.status,
        status: response.status,
      });
    return value;
  }
  const org = await plugin('org-explorer');
  const profile = await call(org, 'GET', '/people/me');
  assert.ok(profile.id);
  await call(
    org,
    'GET',
    '/people?q=' +
      encodeURIComponent((profile.displayName || 'Test').slice(0, 4)) +
      '&limit=1',
  );
  await call(org, 'GET', '/people/me/manager-chain');
  await call(org, 'GET', '/people/me/reports?limit=1');
  await call(org, 'GET', '/people/me/peers?limit=1');
  await call(org, 'GET', '/people/me/subtree?depth=1&maxPeople=5');
  console.log(
    'PASS: Org Explorer profiles, manager chain, reports, peers and bounded subtree',
  );
  if (process.argv.includes('--write')) {
    try {
      const pending = JSON.parse(await fs.readFile(manifest, 'utf8'));
      assert.equal(
        pending.length,
        0,
        'Clean up artifacts from the earlier run first.',
      );
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    const graph = new GraphClient(session.credentials.accessToken, {
        access: 'all',
      }),
      own = await graph.request(
        'GET',
        '/me/drive/root?$select=id,parentReference',
      );
    const destinations = [
      { name: 'OneDrive', drive: own.parentReference.driveId, parent: own.id },
    ];
    let selection = {};
    try {
      selection = JSON.parse(
        await fs.readFile(path.join(directory, 'selection.json'), 'utf8'),
      );
    } catch {}
    const drive =
        process.env.MICROSOFT_TEST_SHAREPOINT_DRIVE ||
        selection.sharepointDrive,
      folder =
        process.env.MICROSOFT_TEST_SHAREPOINT_FOLDER ||
        selection.sharepointFolder;
    if (drive) {
      const parent = await graph.request(
        'GET',
        `/drives/${encodeURIComponent(drive)}/root${folder ? ':/' + folder.split('/').map(encodeURIComponent).join('/') : ''}?$select=id`,
      );
      destinations.push({ name: 'SharePoint', drive, parent: parent.id });
    }
    for (const kind of ['word', 'powerpoint']) {
      const target = await plugin(kind),
        plural = kind === 'word' ? 'documents' : 'presentations';
      for (const destination of destinations) {
        const name = `Switchboard content validation ${randomUUID()}.${kind === 'word' ? 'docx' : 'pptx'}`;
        const content =
          kind === 'word'
            ? {
                paragraphs: [
                  { text: 'Validation heading', style: 'Heading1' },
                  'Original content',
                ],
                tables: [
                  [
                    ['A', 'B'],
                    ['C', 'D'],
                  ],
                ],
              }
            : {
                slides: [
                  {
                    title: 'Original title',
                    body: 'First\nSecond',
                    notes: 'Original notes',
                  },
                ],
              };
        const created = await call(target, 'POST', '/' + plural, {
          drive: destination.drive,
          parent: destination.parent,
          name,
          content,
        });
        cleanup.push({ drive: destination.drive, item: created.file.id });
        await saveCleanup();
        const filePath = `/${plural}/${encodeURIComponent(destination.drive)}/${encodeURIComponent(created.file.id)}`;
        // A repeated create must never overwrite the file.
        await assert.rejects(
          call(target, 'POST', '/' + plural, {
            drive: destination.drive,
            parent: destination.parent,
            name,
            content,
          }),
          (error) => [409, 412].includes(error.status),
        );
        const read = await call(target, 'GET', filePath),
          operations =
            kind === 'word'
              ? [
                  {
                    op: 'replaceText',
                    id: read.content.paragraphs[1].id,
                    find: 'Original',
                    text: 'Updated',
                  },
                  {
                    op: 'insertParagraph',
                    after: read.content.paragraphs[1].id,
                    text: 'Added paragraph',
                  },
                  {
                    op: 'setCellText',
                    id: read.content.tables[0].rows[1][0].id,
                    text: 'Changed cell',
                  },
                ]
              : [
                  {
                    op: 'setShapeText',
                    id: read.content.slides[0].shapes[0].id,
                    text: 'Updated title',
                  },
                  {
                    op: 'setNotes',
                    id: read.content.slides[0].id,
                    text: 'Updated notes',
                  },
                  {
                    op: 'duplicateSlide',
                    id: read.content.slides[0].id,
                    notes: 'Duplicate notes',
                  },
                  { op: 'addSlide', title: 'Added slide', body: 'Added body' },
                ];
        const updated = await call(target, 'POST', filePath + '/edits', {
          version: read.file.version,
          operations,
        });
        assert.ok(updated.file.version);
        assert.notEqual(updated.file.version, read.file.version);
        await assert.rejects(
          call(target, 'POST', filePath + '/edits', {
            version: read.file.version,
            operations,
          }),
          { code: 'version_conflict' },
        );
        const download = await new GraphClient(
          session.credentials.accessToken,
          { access: 'all' },
        ).read(destination.drive, created.file.id);
        const parsed =
          kind === 'word'
            ? readWord(download.bytes)
            : readPowerPoint(download.bytes);
        if (kind === 'word') {
          assert.ok(
            parsed.paragraphs.some((p) => p.text === 'Updated content'),
          );
          assert.ok(
            parsed.paragraphs.some((p) => p.text === 'Added paragraph'),
          );
          assert.equal(parsed.tables[0].rows[1][0].text, 'Changed cell');
        } else {
          assert.equal(parsed.slides.length, 3);
          assert.equal(parsed.slides[0].shapes[0].text, 'Updated title');
          assert.equal(parsed.slides[0].notes.text, 'Updated notes');
          assert.equal(parsed.slides[1].notes.text, 'Duplicate notes');
          const moved = await call(target, 'POST', filePath + '/edits', {
            version: download.metadata.eTag,
            operations: [
              { op: 'moveSlide', id: parsed.slides[2].id, index: 0 },
              { op: 'deleteSlide', id: parsed.slides[1].id },
            ],
          });
          assert.ok(moved.file.version);
        }
        if (process.argv.includes('--pdf')) {
          const pdf = await new GraphClient(session.credentials.accessToken, {
            access: 'all',
          }).download(destination.drive, created.file.id, { format: 'pdf' });
          assert.equal(pdf.subarray(0, 5).toString(), '%PDF-');
          console.log(
            `PASS: Microsoft Office rendered the edited ${destination.name} ${kind} file as PDF`,
          );
        }
        // Explicitly test Microsoft's precondition handling, bypassing the local stale-version check.
        await assert.rejects(
          new GraphClient(session.credentials.accessToken, {
            access: 'all',
          }).replace(
            destination.drive,
            created.file.id,
            read.file.version,
            download.bytes,
          ),
          { code: 'version_conflict' },
        );
        const copy = await call(target, 'POST', '/' + plural, {
          drive: destination.drive,
          parent: destination.parent,
          name: name.replace(/\.(docx|pptx)$/, ' copy.$1'),
          template: { drive: destination.drive, item: created.file.id },
        });
        cleanup.push({ drive: destination.drive, item: copy.file.id });
        await saveCleanup();
        console.log(
          `PASS: ${destination.name} ${kind} creation, content edits, structural edits, template copy, collision and conditional-version checks`,
        );
      }
    }
  } else
    console.log(
      'SKIP: document writes; use --write to create and remove disposable files.',
    );
} catch (error) {
  failed = true;
  console.error(
    `${error.code || 'validation_failed'}: content validation did not complete.`,
  );
} finally {
  for (const p of plugins) await p.dispose();
  for (let i = cleanup.length - 1; i >= 0; i--) {
    let removed = false;
    for (let attempt = 0; attempt < 12; attempt++) {
      try {
        await new GraphClient(session.credentials.accessToken, {
          access: 'all',
        }).request(
          'DELETE',
          `/drives/${encodeURIComponent(cleanup[i].drive)}/items/${encodeURIComponent(cleanup[i].item)}`,
        );
        removed = true;
        break;
      } catch (error) {
        if (error.status === 404) {
          removed = true;
          break;
        }
        if (error.status !== 423) {
          console.error(`Cleanup failed (${error.code || 'request_failed'}).`);
          break;
        }
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }
    }
    if (removed) {
      cleanup.splice(i, 1);
      await saveCleanup();
    } else failed = true;
  }
  await session.close();
  if (cleanup.length === 0)
    console.log(
      'PASS: no disposable validation files remain; credentials retained.',
    );
  process.exitCode = failed ? 1 : 0;
}

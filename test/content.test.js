import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {
  createWord,
  readWord,
  editWord,
} from '../plugins/microsoft/content/word.js';
import {
  createPowerPoint,
  readPowerPoint,
  editPowerPoint,
} from '../plugins/microsoft/content/powerpoint.js';
import {
  parseZip,
  writeZip,
  packageParts,
  xml,
  NS,
  descendants,
} from '../plugins/microsoft/content/package.js';
import { GraphClient } from '../plugins/microsoft/content/graph.js';
import { OrgExplorer } from '../plugins/microsoft/content/org.js';
import { createContentPlugin } from '../plugins/microsoft/content/service.js';
const fixture = (name) =>
  fs.readFile(new URL('./fixtures/content/' + name, import.meta.url));
const json = (value, status = 200) =>
  new Response(JSON.stringify(value), { status });
const unchanged = (before, after, changed) => {
  for (const [name, bytes] of parseZip(before))
    if (!changed.includes(name))
      assert.deepEqual(parseZip(after).get(name), bytes, name);
};
test('Word edits split formatted runs and table cells, preserving every untouched package part', async () => {
  const bytes = await fixture('complex.docx'),
    read = readWord(bytes),
    paragraph = read.paragraphs.find((p) => p.text === 'Hello World');
  const edited = editWord(bytes, [
    {
      op: 'replaceText',
      id: paragraph.id,
      find: 'lo Wor',
      text: 'lo Wonderful Wor',
    },
    { op: 'setCellText', id: 'table0/r1/c0', text: 'Unicode 雪 & <text>' },
  ]).bytes;
  const result = readWord(edited);
  assert.equal(
    result.paragraphs.find((p) => p.id === paragraph.id).text,
    'Hello Wonderful World',
  );
  assert.equal(result.tables[0].rows[1][0].text, 'Unicode 雪 & <text>');
  assert.equal(
    result.paragraphs.find((p) => p.id === paragraph.id).runs[0].bold,
    true,
  );
  assert.equal(
    result.paragraphs.find((p) => p.id === paragraph.id).runs[1].italic,
    true,
  );
  unchanged(bytes, edited, ['word/document.xml']);
  assert.throws(
    () =>
      editWord(bytes, [
        { op: 'replaceText', id: paragraph.id, find: 'missing', text: 'x' },
      ]),
    { code: 'text_conflict' },
  );
});
test('Word creation and structural batches preserve version-bound targets and valid cell structure', () => {
  const bytes = createWord({
    paragraphs: [{ text: 'Heading', style: 'Heading1' }, 'Second'],
    tables: [[['A', 'B']]],
  });
  const edited = editWord(bytes, [
    { op: 'insertParagraph', before: 'p1', text: 'Inserted' },
    { op: 'deleteParagraph', id: 'p1' },
    { op: 'insertTable', rows: [['More']] },
  ]).bytes;
  assert.deepEqual(
    readWord(edited)
      .paragraphs.slice(0, 2)
      .map((p) => p.text),
    ['Heading', 'Inserted'],
  );
  assert.equal(readWord(edited).tables.length, 2);
  assert.throws(() => editWord(bytes, [{ op: 'deleteParagraph', id: 'p2' }]), {
    code: 'invalid_structure',
  });
});
test('PowerPoint text, tables, notes and structural edits retain layouts/themes and isolate duplicate notes', async () => {
  const bytes = await fixture('complex.pptx'),
    read = readPowerPoint(bytes),
    slide = read.slides[0],
    shape = slide.shapes.find((s) => s.text === 'Hello World'),
    table = slide.shapes.find((s) => s.type === 'table');
  const edited = editPowerPoint(bytes, [
    {
      op: 'replaceText',
      id: shape.paragraphs[0].id,
      find: 'World',
      text: '雪',
    },
    { op: 'setCellText', id: table.rows[0][0].id, text: 'Changed' },
    { op: 'setNotes', id: slide.id, text: 'New notes' },
  ]).bytes;
  const result = readPowerPoint(edited);
  assert.equal(
    result.slides[0].shapes.find((s) => s.id === shape.id).text,
    'Hello 雪',
  );
  assert.equal(result.slides[0].notes.text, 'New notes');
  assert.equal(
    result.slides[0].shapes.find((s) => s.type === 'table').rows[0][0].text,
    'Changed',
  );
  unchanged(bytes, edited, [
    'ppt/slides/slide1.xml',
    'ppt/notesSlides/notesSlide1.xml',
  ]);
  const duplicated = editPowerPoint(edited, [
    {
      op: 'duplicateSlide',
      id: slide.id,
      index: 0,
      notes: 'Independent notes',
    },
    { op: 'addSlide', title: 'Added', body: 'Body', notes: 'New slide notes' },
  ]).bytes;
  const slides = readPowerPoint(duplicated).slides;
  assert.equal(slides.length, 3);
  assert.equal(slides[0].notes.text, 'Independent notes');
  assert.equal(slides[1].notes.text, 'New notes');
  assert.equal(slides[2].notes.text, 'New slide notes');
  const moved = editPowerPoint(duplicated, [
    { op: 'moveSlide', id: slides[2].id, index: 0 },
    { op: 'deleteSlide', id: slides[0].id },
  ]).bytes;
  assert.equal(readPowerPoint(moved).slides[0].id, slides[2].id);
});
test('PowerPoint creates simple slides and can add notes to a package without an existing notes master', () => {
  const bytes = createPowerPoint({
    slides: [
      { title: 'Title', body: 'A\nB' },
      { title: 'Two', notes: 'Notes' },
    ],
  });
  assert.equal(readPowerPoint(bytes).slides.length, 2);
  const parts = parseZip(bytes),
    presentation = xml(parts.get('ppt/presentation.xml')),
    rels = xml(parts.get('ppt/_rels/presentation.xml.rels'));
  for (const n of descendants(presentation, NS.p, 'notesMasterIdLst'))
    n.parentNode.removeChild(n);
  for (const n of descendants(rels, NS.rel, 'Relationship'))
    if (n.getAttribute('Type').endsWith('/notesMaster'))
      n.parentNode.removeChild(n);
  parts.set('ppt/presentation.xml', Buffer.from(presentation.toString()));
  parts.set('ppt/_rels/presentation.xml.rels', Buffer.from(rels.toString()));
  for (const name of [...parts.keys()])
    if (
      name.startsWith('ppt/notesSlides/') ||
      name.startsWith('ppt/notesMasters/')
    )
      parts.delete(name);
  for (const name of [...parts.keys()].filter((n) =>
    /^ppt\/slides\/_rels\//.test(n),
  )) {
    const doc = xml(parts.get(name));
    for (const n of descendants(doc, NS.rel, 'Relationship'))
      if (n.getAttribute('Type').endsWith('/notesSlide'))
        n.parentNode.removeChild(n);
    parts.set(name, Buffer.from(doc.toString()));
  }
  const result = editPowerPoint(writeZip(parts), [
    {
      op: 'setNotes',
      id: readPowerPoint(writeZip(parts)).slides[0].id,
      text: 'Added notes',
    },
  ]).bytes;
  assert.equal(readPowerPoint(result).slides[0].notes.text, 'Added notes');
});
test('package validation rejects corruption, XML entities, macros, protection, oversized data and unsafe names', async () => {
  const bytes = await fixture('complex.docx');
  const corrupt = Buffer.from(bytes);
  corrupt[40] ^= 1;
  assert.throws(() => parseZip(corrupt), { code: 'invalid_package' });
  assert.throws(() => xml('<!DOCTYPE doc [<!ENTITY xx "x">]><doc/>'), {
    code: 'unsupported_xml',
  });
  const parts = parseZip(bytes);
  parts.set('../evil', Buffer.from('x'));
  assert.throws(() => packageParts(writeZip(parts), 'word'), {
    code: 'invalid_package',
  });
  parts.delete('../evil');
  parts.set('word/vbaProject.bin', Buffer.from('x'));
  assert.throws(() => readWord(writeZip(parts)), {
    code: 'unsupported_package',
  });
  parts.delete('word/vbaProject.bin');
  parts.set(
    'word/settings.xml',
    Buffer.from(
      `<w:settings xmlns:w="${NS.w}"><w:documentProtection w:enforcement="1"/></w:settings>`,
    ),
  );
  assert.throws(
    () =>
      editWord(writeZip(parts), [
        { op: 'setParagraphText', id: 'p0', text: 'x' },
      ]),
    { code: 'protected_document' },
  );
  assert.throws(() => parseZip(Buffer.alloc(25 * 1024 * 1024 + 1)), {
    code: 'file_too_large',
  });
});
test('Graph downloads never forward bearer credentials to signed URLs and enforce own-drive and version boundaries', async () => {
  const calls = [],
    client = new GraphClient('private-token', {
      access: 'full',
      fetcher: async (url, opts) => {
        calls.push({ url: String(url), opts });
        if (String(url).includes('/me/drive?')) return json({ id: 'own' });
        if (String(url).endsWith('/content'))
          return new Response(null, {
            status: 302,
            headers: {
              location: 'https://example.sharepoint.com/download?secret=opaque',
            },
          });
        assert.equal(opts.headers, undefined);
        return new Response('file');
      },
    });
  assert.equal((await client.download('own', 'item')).toString(), 'file');
  assert.equal(calls.at(-1).opts.headers, undefined);
  await assert.rejects(client.metadata('other', 'item'), {
    code: 'request_forbidden',
  });
  const conflict = new GraphClient('token', {
    access: 'all',
    fetcher: async (url, opts) => {
      assert.equal(opts.headers['If-Match'], 'old-version');
      return json({}, 412);
    },
  });
  await assert.rejects(
    conflict.replace('d', 'i', 'old-version', Buffer.from('x')),
    { code: 'version_conflict' },
  );
  const evil = new GraphClient('token', {
    access: 'all',
    fetcher: async () =>
      new Response(null, {
        status: 302,
        headers: { location: 'https://evil.example/secret' },
      }),
  });
  await assert.rejects(evil.download('d', 'i'), { code: 'invalid_transfer' });
});
test('Org Explorer handles pagination, absent managers, cycles, bounded subtrees and missing properties', async () => {
  const data = {
    me: { id: 'a', displayName: 'A' },
    a: { id: 'a', displayName: 'A' },
    b: { id: 'b', displayName: 'B' },
    c: { id: 'c' },
  };
  const org = new OrgExplorer({
    request: async (method, path) => {
      assert.equal(method, 'GET');
      const id = decodeURIComponent(
        path.match(/^\/(?:users\/([^/?]+)|me)/)?.[1] ?? 'me',
      );
      if (path.includes('/manager')) {
        if (id === 'c') throw Object.assign(Error(), { status: 404 });
        return data[id === 'b' ? 'a' : 'b'];
      }
      if (path.includes('/directReports'))
        return {
          value: id === 'a' ? [data.b, data.c] : id === 'b' ? [data.a] : [],
        };
      return data[id];
    },
  });
  assert.equal((await org.chain('me')).cycle, true);
  assert.equal(await org.manager('c'), null);
  const tree = await org.subtree('a', { depth: '5', maxPeople: '3' });
  assert.equal(tree.people.length, 3);
  assert.equal(tree.truncated, true);
  assert.equal(
    (await org.peers('b', {})).people.some((p) => p.id === 'b'),
    false,
  );
  await assert.rejects(org.subtree('me', { depth: '6' }), {
    code: 'invalid_input',
  });
  await assert.rejects(
    org.page('/users', {
      cursor: Buffer.from('https://evil.example/users').toString('base64url'),
    }),
    { code: 'invalid_input' },
  );
});
test('adapter exposes no raw-token hook and binds one-use invocations to method/path with read-only enforcement', async () => {
  const token = {
      format: 'fixture',
      accessToken: 'never-return-this',
      refreshToken: 'private',
    },
    auth = {
      id: 'authentication-source',
      fields: [],
      credentialConsumer: { format: 'fixture', field: 'authenticationSource' },
      connect: async () => ({ credentials: token }),
      authorize() {},
      token: async () => ({ accessToken: token.accessToken }),
    };
  const plugin = await createContentPlugin(
    { require: () => ({ createConsumerAuth: () => auth }) },
    { id: 'word', name: 'Word' },
    {
      fetcher: async () => {
        throw Error('Unexpected API');
      },
    },
  );
  try {
    const method = plugin.services[0].authMethods[0];
    assert.equal(method.token, undefined);
    assert.equal(method.reusable, undefined);
    const conn = { credentials: token, config: { access: 'read' } },
      request = {
        method: 'GET',
        url: new URL('http://word.localhost/capabilities'),
        headers: new Headers(),
      };
    await method.authorize(request, conn, {});
    const result = await fetch(request.url, { headers: request.headers });
    assert.equal(result.status, 200);
    assert.equal((await result.json()).kind, 'word');
    assert.equal(
      (await fetch(request.url, { headers: request.headers })).status,
      403,
    );
    await assert.rejects(
      method.authorize(
        {
          method: 'POST',
          url: new URL('http://word.localhost/documents'),
          headers: new Headers(),
        },
        conn,
        {},
      ),
      { code: 'request_forbidden' },
    );
  } finally {
    await plugin.dispose();
  }
});

test('deleting a slide removes its XML and notes rather than retaining hidden text', () => {
  const bytes = createPowerPoint({
    slides: [
      { title: 'Retained' },
      { title: 'Delete this', notes: 'Sensitive deleted notes' },
    ],
  });
  const slides = readPowerPoint(bytes).slides,
    parts = parseZip(bytes);
  const edited = editPowerPoint(bytes, [
    { op: 'deleteSlide', id: slides[1].id },
  ]).bytes;
  const after = parseZip(edited);
  assert.equal(after.has('ppt/slides/slide2.xml'), false);
  assert.equal(after.has('ppt/notesSlides/notesSlide2.xml'), false);
  assert.equal(
    [...after.values()].some((b) =>
      b.toString().includes('Sensitive deleted notes'),
    ),
    false,
  );
});
test('deletion refuses dangling cross-slide links; text edits reject non-text shapes and protected content', async () => {
  const bytes = await fixture('complex.pptx'),
    slide = readPowerPoint(bytes).slides[0],
    table = slide.shapes.find((s) => s.type === 'table');
  assert.throws(
    () =>
      editPowerPoint(bytes, [
        { op: 'setShapeText', id: table.id, text: 'Unsafe' },
      ]),
    { code: 'invalid_input' },
  );
  const two = editPowerPoint(bytes, [
      { op: 'duplicateSlide', id: slide.id },
    ]).bytes,
    parts = parseZip(two);
  const rels = xml(parts.get('ppt/slides/_rels/slide2.xml.rels')),
    node = rels.createElementNS(NS.rel, 'Relationship');
  node.setAttribute('Id', 'crossLink');
  node.setAttribute('Type', NS.r + '/slide');
  node.setAttribute('Target', 'slide1.xml');
  rels.documentElement.appendChild(node);
  parts.set('ppt/slides/_rels/slide2.xml.rels', Buffer.from(rels.toString()));
  assert.throws(
    () =>
      editPowerPoint(writeZip(parts), [{ op: 'deleteSlide', id: slide.id }]),
    { code: 'unsupported_structure' },
  );
  const doc = xml(parts.get('ppt/presentation.xml')),
    protection = doc.createElementNS(NS.p, 'p:modifyVerifier');
  doc.documentElement.appendChild(protection);
  parts.set('ppt/presentation.xml', Buffer.from(doc.toString()));
  assert.throws(
    () =>
      editPowerPoint(writeZip(parts), [
        { op: 'setShapeText', id: slide.shapes[0].id, text: 'x' },
      ]),
    { code: 'protected_document' },
  );
});
test('bounded decompression rejects a forged uncompressed size before expanding a ZIP bomb', () => {
  const parts = new Map([['data', Buffer.alloc(1024 * 1024, 65)]]),
    bytes = writeZip(parts),
    end = bytes.length - 22,
    central = bytes.readUInt32LE(end + 16);
  bytes.writeUInt32LE(1, central + 24);
  assert.throws(() => parseZip(bytes), { code: 'invalid_package' });
});
test('Org Explorer paginates only the same resource and terminates repeated continuation cursors', async () => {
  const cursor =
    'https://graph.microsoft.com/v1.0/users/a/directReports?$skiptoken=next';
  let calls = 0;
  const org = new OrgExplorer({
    request: async (_method, url) => {
      calls++;
      if (url.startsWith('/users/a?')) return { id: 'a' };
      return {
        value: [{ id: 'b', displayName: 'B' }],
        '@odata.nextLink': cursor,
      };
    },
  });
  const first = await org.reports('a', { limit: 1 });
  assert.ok(first.nextCursor);
  assert.equal(
    (await org.reports('a', { limit: 1, cursor: first.nextCursor })).people[0]
      .id,
    'b',
  );
  await assert.rejects(org.reports('different', { cursor: first.nextCursor }), {
    code: 'invalid_input',
  });
  const tree = await org.subtree('a', { depth: 1, maxPeople: 500 });
  assert.equal(tree.truncated, true);
  assert.ok(calls < 10);
});

test('text replacement does not join text across an existing line break', () => {
  const parts = parseZip(createWord({ paragraphs: ['HelloWorld'] }));
  const doc = xml(parts.get('word/document.xml')),
    run = descendants(doc, NS.w, 'r')[0],
    t = descendants(run, NS.w, 't')[0];
  t.textContent = 'Hello';
  run.appendChild(doc.createElementNS(NS.w, 'w:br'));
  const tail = doc.createElementNS(NS.w, 'w:t');
  tail.textContent = 'World';
  run.appendChild(tail);
  parts.set('word/document.xml', Buffer.from(doc.toString()));
  const bytes = writeZip(parts);
  assert.equal(readWord(bytes).paragraphs[0].text, 'Hello\nWorld');
  assert.throws(
    () =>
      editWord(bytes, [
        { op: 'replaceText', id: 'p0', find: 'HelloWorld', text: 'Joined' },
      ]),
    { code: 'text_conflict' },
  );
  const replaced = editWord(bytes, [
    { op: 'replaceText', id: 'p0', find: 'World', text: 'There' },
  ]).bytes;
  assert.equal(readWord(replaced).paragraphs[0].text, 'Hello\nThere');
  assert.equal(
    readWord(
      editWord(bytes, [{ op: 'setParagraphText', id: 'p0', text: 'One line' }])
        .bytes,
    ).paragraphs[0].text,
    'One line',
  );
});

test('organization contacts remain visible leaf nodes without invalid user traversal', async () => {
  const org = new OrgExplorer({
    request: async (_method, path) => {
      if (path.startsWith('/users/a?')) return { id: 'a' };
      if (path.startsWith('/users/a/directReports'))
        return {
          value: [
            {
              '@odata.type': '#microsoft.graph.orgContact',
              id: 'external-contact',
              displayName: 'Contact',
            },
          ],
        };
      throw Error('Contact must not be traversed as a user');
    },
  });
  const result = await org.subtree('a', { depth: 5 });
  assert.equal(result.people.length, 2);
  assert.equal(result.people[1].type, 'contact');
  assert.equal(result.depthLimited, false);
});

import fs from 'node:fs';
import path from 'node:path';
import {
  NS,
  xml,
  serialize,
  descendants,
  elements,
  text,
  esc,
  check,
  fail,
  writeZip,
  packageParts,
  replaceText,
  setText,
} from './package.js';
const P = NS.p,
  A = NS.a,
  R = NS.r,
  REL = NS.rel;
const child = (n, ns, name) =>
  elements(n).filter((c) => c.namespaceURI === ns && c.localName === name);
const relPath = (part) =>
  path.posix.join(
    path.posix.dirname(part),
    '_rels',
    path.posix.basename(part) + '.rels',
  );
const resolve = (part, target) => {
  check(
    target && !target.includes('\\') && !target.includes(':'),
    'Unsupported package relationship.',
  );
  const name = path.posix.normalize(
    target.startsWith('/')
      ? target.slice(1)
      : path.posix.join(path.posix.dirname(part), target),
  );
  check(
    !name.startsWith('../') && name !== '..',
    'Invalid package relationship.',
  );
  return name;
};
const plain = (node) =>
  descendants(node, A, 'p')
    .map((p) =>
      elements(p)
        .map((n) =>
          n.localName === 'br'
            ? '\n'
            : descendants(n, A, 't')
                .map((t) => t.textContent)
                .join(''),
        )
        .join(''),
    )
    .join('\n');
const relationship = (doc, type) =>
  descendants(doc, REL, 'Relationship').find(
    (r) =>
      r.getAttribute('Type') === R + '/' + type &&
      r.getAttribute('TargetMode') !== 'External',
  );
const blank = () =>
  packageParts(
    fs.readFileSync(new URL('./blank.pptx', import.meta.url)),
    'powerpoint',
  );
function model(parts) {
  const doc = xml(parts.get('ppt/presentation.xml') ?? ''),
    rels = xml(parts.get('ppt/_rels/presentation.xml.rels') ?? '');
  check(
    doc.documentElement.namespaceURI === P,
    'Unsupported presentation structure.',
  );
  const targets = new Map(),
    slides = [],
    list = descendants(doc, P, 'sldIdLst')[0];
  check(list, 'Presentation has no slide list.');
  const usedParts = new Set();
  for (const item of child(list, P, 'sldId')) {
    const rid = item.getAttributeNS(R, 'id'),
      r = descendants(rels, REL, 'Relationship').find(
        (r) =>
          r.getAttribute('Id') === rid &&
          r.getAttribute('TargetMode') !== 'External',
      );
    check(
      r && r.getAttribute('Type') === R + '/slide',
      'Invalid slide relationship.',
    );
    const part = resolve('ppt/presentation.xml', r.getAttribute('Target')),
      slide = xml(parts.get(part) ?? ''),
      id = 'slide' + item.getAttribute('id'),
      relationships = xml(
        parts.get(relPath(part)) ?? `<Relationships xmlns="${REL}"/>`,
      );
    check(
      !targets.has(id) && !usedParts.has(part),
      'Duplicate slide identifiers or parts.',
    );
    usedParts.add(part);
    const entry = {
      id,
      part,
      item,
      doc: slide,
      relationships,
      shapes: [],
      notes: null,
    };
    targets.set(id, { ...entry, entry, kind: 'slide' });
    for (const shape of [
      ...descendants(slide, P, 'sp'),
      ...descendants(slide, P, 'graphicFrame'),
    ]) {
      const nativeId = descendants(shape, P, 'cNvPr')[0]?.getAttribute('id');
      check(
        nativeId && !targets.has(`${id}/shape${nativeId}`),
        'Missing or duplicate shape identifier.',
      );
      const shapeId = `${id}/shape${nativeId}`,
        type = shape.localName === 'sp' ? 'text' : 'table',
        paragraphs = descendants(shape, A, 'p').map((p, i) => {
          const pid = `${shapeId}/p${i}`;
          targets.set(pid, { node: p, entry, kind: 'paragraph' });
          return {
            id: pid,
            text: descendants(p, A, 't')
              .map((t) => t.textContent)
              .join(''),
            runs: descendants(p, A, 'r').map((run) => ({
              text: descendants(run, A, 't')
                .map((t) => t.textContent)
                .join(''),
              bold: descendants(run, A, 'rPr')[0]?.getAttribute('b') === '1',
            })),
          };
        });
      targets.set(shapeId, { node: shape, entry, kind: 'shape' });
      const result = {
        id: shapeId,
        type,
        text: plain(shape),
        paragraphs,
        editable: type === 'text' && descendants(shape, P, 'txBody').length > 0,
      };
      const table = descendants(shape, A, 'tbl')[0];
      if (table) {
        result.type = 'table';
        result.rows = child(table, A, 'tr').map((row, r) =>
          child(row, A, 'tc').map((cell, c) => {
            const cid = `${shapeId}/r${r}/c${c}`;
            targets.set(cid, { node: cell, entry, kind: 'cell' });
            return { id: cid, text: plain(cell) };
          }),
        );
      } else if (type === 'table') result.type = 'unsupported';
      entry.shapes.push(result);
    }
    const notesRel = relationship(relationships, 'notesSlide');
    if (notesRel) {
      const notesPart = resolve(part, notesRel.getAttribute('Target')),
        notes = xml(parts.get(notesPart) ?? ''),
        body = descendants(notes, P, 'sp').find((n) =>
          descendants(n, P, 'ph').some(
            (ph) => ph.getAttribute('type') === 'body',
          ),
        );
      entry.notesPart = notesPart;
      entry.notesDoc = notes;
      entry.notes = {
        id: `${id}/notes`,
        text: body ? plain(body) : '',
        editable: !!body,
      };
      if (body)
        targets.set(entry.notes.id, { node: body, entry, kind: 'notes' });
    }
    slides.push(entry);
  }
  check(slides.length <= 1000, 'Presentations are limited to 1000 slides.');
  return { doc, rels, list, slides, targets };
}
function output(m) {
  return {
    slides: m.slides.map((s) => ({
      id: s.id,
      shapes: s.shapes,
      notes: s.notes,
    })),
    capabilities: [
      'replaceText',
      'setShapeText',
      'setCellText',
      'setNotes',
      'addSlide',
      'duplicateSlide',
      'moveSlide',
      'deleteSlide',
    ],
  };
}
export function readPowerPoint(input) {
  return output(model(packageParts(input, 'powerpoint')));
}
const free = (parts, folder, prefix, ext = 'xml') => {
  for (let n = 1; ; n++) {
    const name = `${folder}/${prefix}${n}.${ext}`;
    if (!parts.has(name)) return name;
  }
};
function newRel(rels, type, target) {
  let n = 1;
  const used = new Set(
    descendants(rels, REL, 'Relationship').map((r) => r.getAttribute('Id')),
  );
  while (used.has('rId' + n)) n++;
  const node = rels.createElementNS(REL, 'Relationship');
  node.setAttribute('Id', 'rId' + n);
  node.setAttribute('Type', R + '/' + type);
  node.setAttribute('Target', target);
  rels.documentElement.appendChild(node);
  return node;
}
function contentType(parts, part, type) {
  const doc = xml(parts.get('[Content_Types].xml'));
  if (
    !descendants(doc, NS.ct, 'Override').some(
      (n) => n.getAttribute('PartName') === '/' + part,
    )
  ) {
    const n = doc.createElementNS(NS.ct, 'Override');
    n.setAttribute('PartName', '/' + part);
    n.setAttribute(
      'ContentType',
      'application/vnd.openxmlformats-officedocument.presentationml.' +
        type +
        '+xml',
    );
    doc.documentElement.appendChild(n);
    parts.set('[Content_Types].xml', serialize(doc));
  }
}
function notesMaster(parts, m) {
  const current = relationship(m.rels, 'notesMaster');
  if (current)
    return resolve('ppt/presentation.xml', current.getAttribute('Target'));
  const preset = blank(),
    master = free(parts, 'ppt/notesMasters', 'notesMaster'),
    theme = free(parts, 'ppt/theme', 'theme'),
    source = 'ppt/notesMasters/notesMaster1.xml',
    rels = xml(preset.get(relPath(source))),
    themeRel = relationship(rels, 'theme');
  check(themeRel, 'Notes template has no theme.');
  const themeSource = resolve(source, themeRel.getAttribute('Target'));
  parts.set(theme, preset.get(themeSource));
  themeRel.setAttribute(
    'Target',
    path.posix.relative(path.posix.dirname(master), theme),
  );
  parts.set(master, preset.get(source));
  parts.set(relPath(master), serialize(rels));
  contentType(parts, master, 'notesMaster');
  const types = xml(parts.get('[Content_Types].xml')),
    t = types.createElementNS(NS.ct, 'Override');
  t.setAttribute('PartName', '/' + theme);
  t.setAttribute(
    'ContentType',
    'application/vnd.openxmlformats-officedocument.theme+xml',
  );
  types.documentElement.appendChild(t);
  parts.set('[Content_Types].xml', serialize(types));
  m.notesMasterDirty = true;
  const r = newRel(m.rels, 'notesMaster', path.posix.relative('ppt', master)),
    list = m.doc.createElementNS(P, 'p:notesMasterIdLst'),
    id = m.doc.createElementNS(P, 'p:notesMasterId');
  id.setAttributeNS(R, 'r:id', r.getAttribute('Id'));
  list.appendChild(id);
  const before = elements(m.doc.documentElement).find((n) =>
    [
      'handoutMasterIdLst',
      'sldIdLst',
      'sldSz',
      'notesSz',
      'defaultTextStyle',
      'extLst',
    ].includes(n.localName),
  );
  m.doc.documentElement.insertBefore(list, before ?? null);
  return master;
}
function paragraphs(value) {
  return text(value)
    .split(/\r?\n/)
    .map(
      (line) =>
        `<a:p><a:r><a:rPr lang="en-US"/><a:t xml:space="preserve">${esc(line)}</a:t></a:r></a:p>`,
    )
    .join('');
}
function bodyText(node, value) {
  text(value);
  check(
    !descendants(node, A, 'fld').length,
    'Fields cannot be edited as text.',
  );
  const container =
    descendants(node, P, 'txBody')[0] ?? descendants(node, A, 'txBody')[0];
  check(container, 'Element has no editable text body.');
  const ps = child(container, A, 'p');
  if (!value.includes('\n') && ps.length && descendants(ps[0], A, 't').length) {
    setText(ps[0], A, value);
    for (const p of ps.slice(1)) container.removeChild(p);
  } else {
    const doc = node.ownerDocument;
    for (const p of ps) container.removeChild(p);
    const source = xml(`<a:root xmlns:a="${A}">${paragraphs(value)}</a:root>`);
    for (const p of elements(source.documentElement))
      container.appendChild(doc.importNode(p, true));
  }
}
function setNotes(parts, m, entry, value) {
  text(value);
  if (entry.notesDoc) {
    const body = descendants(entry.notesDoc, P, 'sp').find((n) =>
      descendants(n, P, 'ph').some((p) => p.getAttribute('type') === 'body'),
    );
    check(body, 'Notes body is unsupported.');
    bodyText(body, value);
    parts.set(entry.notesPart, serialize(entry.notesDoc));
    return;
  }
  const master = notesMaster(parts, m),
    part = free(parts, 'ppt/notesSlides', 'notesSlide');
  const doc = xml(
    `<p:notes xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs(value)}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
  );
  const rels = xml(`<Relationships xmlns="${REL}"/>`);
  newRel(
    rels,
    'notesMaster',
    path.posix.relative(path.posix.dirname(part), master),
  );
  newRel(
    rels,
    'slide',
    path.posix.relative(path.posix.dirname(part), entry.part),
  );
  newRel(
    entry.relationships,
    'notesSlide',
    path.posix.relative(path.posix.dirname(entry.part), part),
  );
  parts.set(part, serialize(doc));
  parts.set(relPath(part), serialize(rels));
  contentType(parts, part, 'notesSlide');
  entry.notesPart = part;
  entry.notesDoc = doc;
  entry.relationshipsDirty = true;
}
function validateLinks(parts) {
  for (const [part, bytes] of parts) {
    if (!part.endsWith('.rels')) continue;
    const owner =
      part === '_rels/.rels'
        ? ''
        : path.posix.join(
            path.posix.dirname(path.posix.dirname(part)),
            path.posix.basename(part).slice(0, -5),
          );
    for (const r of descendants(xml(bytes), REL, 'Relationship')) {
      if (r.getAttribute('TargetMode') === 'External') continue;
      const target = resolve(owner, r.getAttribute('Target').split('#')[0]);
      check(
        parts.has(target),
        'Presentation contains a missing package relationship.',
      );
    }
  }
}
export function editPowerPoint(input, operations) {
  check(
    Array.isArray(operations) &&
      operations.length > 0 &&
      operations.length <= 500,
    'Provide 1–500 edit operations.',
  );
  const parts = packageParts(input, 'powerpoint', true),
    m = model(parts),
    changes = [],
    dirty = new Set();
  let structureDirty = false;
  const target = (id, kind) => {
    const t = m.targets.get(id);
    check(
      t && t.entry.item.parentNode === m.list && (!kind || t.kind === kind),
      'Unknown or incompatible presentation element.',
    );
    return t;
  };
  const structural = () => {
    structureDirty = true;
    if (
      Array.from(m.doc.getElementsByTagName('*')).some((n) =>
        /^(custShowLst|sectionLst|sectionList)$/.test(n.localName),
      )
    )
      throw fail(
        'unsupported_structure',
        'Structural edits to presentations with custom shows or sections are unsupported.',
        409,
      );
  };
  for (const op of operations) {
    check(op && typeof op === 'object', 'Invalid edit operation.');
    if (op.op === 'replaceText') {
      const t = target(op.id, 'paragraph');
      check(
        !descendants(t.node, A, 'fld').length,
        'Fields cannot be edited as text.',
      );
      replaceText(t.node, A, op.find, op.text, op.expectedMatches ?? 1);
      dirty.add(t.entry);
    } else if (op.op === 'setShapeText') {
      const t = target(op.id, 'shape');
      check(
        t.node.localName === 'sp' && child(t.node, P, 'txBody').length > 0,
        'Select a text shape.',
      );
      bodyText(t.node, op.text);
      dirty.add(t.entry);
    } else if (op.op === 'setCellText') {
      const t = target(op.id, 'cell');
      check(
        !['hMerge', 'vMerge'].some((a) =>
          ['1', 'true'].includes(t.node.getAttribute(a)),
        ),
        'Edit the leading cell of a merged table region.',
      );
      bodyText(t.node, op.text);
      dirty.add(t.entry);
    } else if (op.op === 'setNotes') {
      const t = target(op.id, 'slide');
      setNotes(parts, m, t.entry, op.text);
    } else if (op.op === 'moveSlide') {
      structural();
      const t = target(op.id, 'slide'),
        ids = child(m.list, P, 'sldId').filter((n) => n !== t.item);
      check(
        Number.isInteger(op.index) && op.index >= 0 && op.index <= ids.length,
        'Invalid destination slide index.',
      );
      m.list.removeChild(t.item);
      m.list.insertBefore(t.item, ids[op.index] ?? null);
    } else if (op.op === 'deleteSlide') {
      structural();
      const t = target(op.id, 'slide');
      check(
        child(m.list, P, 'sldId').length > 1,
        'A presentation must retain at least one slide.',
      );
      m.list.removeChild(t.item);
      const rid = t.item.getAttributeNS(R, 'id'),
        rel = descendants(m.rels, REL, 'Relationship').find(
          (r) => r.getAttribute('Id') === rid,
        );
      m.rels.documentElement.removeChild(rel);
      const removed = new Set([
        t.entry.part,
        ...(t.entry.notesPart ? [t.entry.notesPart] : []),
      ]);
      for (const [name, bytes] of parts) {
        if (
          !name.endsWith('.rels') ||
          name === 'ppt/_rels/presentation.xml.rels'
        )
          continue;
        const owner =
          name === '_rels/.rels'
            ? ''
            : path.posix.join(
                path.posix.dirname(path.posix.dirname(name)),
                path.posix.basename(name).slice(0, -5),
              );
        if (removed.has(owner)) continue;
        for (const link of descendants(xml(bytes), REL, 'Relationship'))
          if (
            link.getAttribute('TargetMode') !== 'External' &&
            removed.has(
              resolve(owner, link.getAttribute('Target').split('#')[0]),
            )
          )
            throw fail(
              'unsupported_structure',
              'Another package part links to this slide or its notes. Remove that link before deleting the slide.',
              409,
            );
      }
      for (const part of removed) {
        parts.delete(part);
        parts.delete(relPath(part));
      }
      const types = xml(parts.get('[Content_Types].xml'));
      for (const entry of descendants(types, NS.ct, 'Override'))
        if (removed.has(entry.getAttribute('PartName').slice(1)))
          entry.parentNode.removeChild(entry);
      parts.set('[Content_Types].xml', serialize(types));
    } else if (op.op === 'duplicateSlide' || op.op === 'addSlide') {
      structural();
      check(
        child(m.list, P, 'sldId').length < 1000,
        'Presentations are limited to 1000 slides.',
      );
      const source =
        op.op === 'duplicateSlide' ? target(op.id, 'slide').entry : m.slides[0];
      const part = free(parts, 'ppt/slides', 'slide'),
        rels = xml(serialize(source.relationships));
      let slide;
      if (op.op === 'duplicateSlide') slide = xml(serialize(source.doc));
      else {
        for (const r of [...descendants(rels, REL, 'Relationship')])
          if (r.getAttribute('Type') !== R + '/slideLayout')
            rels.documentElement.removeChild(r);
        const size = descendants(m.doc, P, 'sldSz')[0],
          width = Number(size?.getAttribute('cx')) || 12192000,
          height = Number(size?.getAttribute('cy')) || 6858000;
        const shape = (id, name, value, y, h, font) =>
          `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(width * 0.06)}" y="${y}"/><a:ext cx="${Math.round(width * 0.88)}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${text(
            value,
          )
            .split(/\r?\n/)
            .map(
              (line) =>
                `<a:p><a:r><a:rPr lang="en-US" sz="${font}"/><a:t xml:space="preserve">${esc(line)}</a:t></a:r></a:p>`,
            )
            .join('')}</p:txBody></p:sp>`;
        slide = xml(
          `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shape(2, 'Title', op.title ?? '', Math.round(height * 0.06), Math.round(height * 0.2), 3200)}${shape(3, 'Body', op.body ?? '', Math.round(height * 0.3), Math.round(height * 0.6), 2200)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
        );
      }
      const note = relationship(rels, 'notesSlide');
      if (note) {
        const original = resolve(source.part, note.getAttribute('Target')),
          newPart = free(parts, 'ppt/notesSlides', 'notesSlide'),
          nr = xml(parts.get(relPath(original)));
        const back = relationship(nr, 'slide');
        check(back, 'Notes have no slide relationship.');
        back.setAttribute(
          'Target',
          path.posix.relative(path.posix.dirname(newPart), part),
        );
        parts.set(
          newPart,
          source.notesDoc ? serialize(source.notesDoc) : parts.get(original),
        );
        parts.set(relPath(newPart), serialize(nr));
        note.setAttribute(
          'Target',
          path.posix.relative(path.posix.dirname(part), newPart),
        );
        contentType(parts, newPart, 'notesSlide');
      }
      parts.set(part, serialize(slide));
      parts.set(relPath(part), serialize(rels));
      contentType(parts, part, 'slide');
      const relationshipNode = newRel(
          m.rels,
          'slide',
          path.posix.relative('ppt', part),
        ),
        id = m.doc.createElementNS(P, 'p:sldId');
      id.setAttribute(
        'id',
        String(
          Math.max(
            255,
            ...child(m.list, P, 'sldId').map((n) =>
              Number(n.getAttribute('id')),
            ),
          ) + 1,
        ),
      );
      id.setAttributeNS(R, 'r:id', relationshipNode.getAttribute('Id'));
      const ids = child(m.list, P, 'sldId'),
        index = op.index ?? ids.length;
      check(
        Number.isInteger(index) && index >= 0 && index <= ids.length,
        'Invalid slide index.',
      );
      m.list.insertBefore(id, ids[index] ?? null);
      if (op.notes !== undefined) {
        const noteRel = relationship(rels, 'notesSlide'),
          notesPart = noteRel
            ? resolve(part, noteRel.getAttribute('Target'))
            : undefined;
        const entry = {
          part,
          doc: slide,
          relationships: rels,
          notesPart,
          notesDoc: notesPart ? xml(parts.get(notesPart)) : undefined,
        };
        setNotes(parts, m, entry, op.notes);
        parts.set(relPath(part), serialize(rels));
      }
    } else
      throw fail(
        'unsupported_operation',
        'Unsupported PowerPoint edit operation.',
      );
    changes.push({ op: op.op, id: op.id ?? null });
  }
  for (const entry of dirty)
    if (entry.item.parentNode === m.list)
      parts.set(entry.part, serialize(entry.doc));
  for (const entry of m.slides)
    if (entry.relationshipsDirty && entry.item.parentNode === m.list)
      parts.set(relPath(entry.part), serialize(entry.relationships));
  if (
    structureDirty ||
    m.notesMasterDirty ||
    m.slides.some((s) => s.relationshipsDirty)
  ) {
    parts.set('ppt/presentation.xml', serialize(m.doc));
    parts.set('ppt/_rels/presentation.xml.rels', serialize(m.rels));
  }
  validateLinks(parts);
  const bytes = writeZip(parts);
  readPowerPoint(bytes);
  return { bytes, changes };
}
export function createPowerPoint({ slides = [{ title: '', body: '' }] } = {}) {
  check(
    Array.isArray(slides) && slides.length > 0 && slides.length <= 499,
    'Provide 1–499 slides.',
  );
  let bytes = writeZip(blank());
  const first = readPowerPoint(bytes).slides[0];
  bytes = editPowerPoint(bytes, [
    { op: 'setShapeText', id: first.shapes[0].id, text: slides[0].title ?? '' },
    { op: 'setShapeText', id: first.shapes[1].id, text: slides[0].body ?? '' },
    { op: 'setNotes', id: first.id, text: slides[0].notes ?? '' },
  ]).bytes;
  if (slides.length > 1)
    bytes = editPowerPoint(
      bytes,
      slides.slice(1).map((s) => ({ op: 'addSlide', ...s })),
    ).bytes;
  return bytes;
}

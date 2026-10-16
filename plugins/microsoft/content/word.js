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
const W = NS.w;
const child = (node, name) =>
  elements(node).filter((n) => n.namespaceURI === W && n.localName === name);
function editable(p) {
  for (let n = p.parentNode; n; n = n.parentNode)
    if (
      n.namespaceURI === W &&
      ['sdt', 'ins', 'del', 'moveFrom', 'moveTo'].includes(n.localName)
    )
      return false;
  return ![
    'ins',
    'del',
    'moveFrom',
    'moveTo',
    'fldChar',
    'instrText',
    'fldSimple',
    'sdt',
    'altChunk',
    'drawing',
    'pict',
    'object',
  ].some((name) => descendants(p, W, name).length);
}
function paragraph(doc, value, style = 'Normal') {
  text(value);
  check(
    !/[\r\n\t]/.test(value),
    'Use separate Word paragraphs for line changes.',
  );
  check(
    /^(Normal|Title|Subtitle|Heading[1-9])$/.test(style),
    'Unsupported paragraph style.',
  );
  return doc.importNode(
    xml(
      `<w:p xmlns:w="${W}"><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t xml:space="preserve">${esc(value)}</w:t></w:r></w:p>`,
    ).documentElement,
    true,
  );
}
const plain = (node) => {
  let value = '';
  const visit = (n) => {
    if (n.nodeType !== 1) return;
    if (n.namespaceURI === W && n.localName === 't') value += n.textContent;
    else if (n.namespaceURI === W && ['br', 'cr', 'tab'].includes(n.localName))
      value += n.localName === 'tab' ? '\t' : '\n';
    else for (const c of elements(n)) visit(c);
  };
  visit(node);
  return value;
};
function model(parts) {
  const doc = xml(parts.get('word/document.xml') ?? ''),
    body = descendants(doc, W, 'body')[0];
  check(
    body && doc.documentElement.namespaceURI === W,
    'Unsupported Word document structure.',
  );
  const paragraphs = descendants(body, W, 'p'),
    tables = descendants(body, W, 'tbl'),
    targets = new Map();
  const ps = paragraphs.map((p, i) => {
    const id = `p${i}`;
    targets.set(id, p);
    return {
      id,
      text: plain(p),
      style:
        descendants(p, W, 'pStyle')[0]?.getAttributeNS(W, 'val') || 'Normal',
      runs: descendants(p, W, 'r').map((r) => ({
        text: plain(r),
        bold: descendants(r, W, 'b').length > 0,
        italic: descendants(r, W, 'i').length > 0,
      })),
      editable: editable(p),
    };
  });
  const ts = tables.map((table, i) => {
    const id = `table${i}`;
    targets.set(id, table);
    return {
      id,
      rows: child(table, 'tr').map((row, r) =>
        child(row, 'tc').map((cell, c) => {
          const cellId = `${id}/r${r}/c${c}`;
          targets.set(cellId, cell);
          return {
            id: cellId,
            text: child(cell, 'p').map(plain).join('\n'),
            paragraphIds: child(cell, 'p').map(
              (p) => `p${paragraphs.indexOf(p)}`,
            ),
            editable: editable(cell),
          };
        }),
      ),
    };
  });
  return { doc, body, targets, paragraphs: ps, tables: ts };
}
export function readWord(input) {
  const parts = packageParts(input, 'word'),
    m = model(parts);
  return {
    paragraphs: m.paragraphs,
    tables: m.tables,
    unavailableParts: [...parts.keys()].filter((k) =>
      /^word\/(header|footer|footnotes|endnotes|comments|charts|embeddings)/.test(
        k,
      ),
    ),
    capabilities: [
      'replaceText',
      'setParagraphText',
      'insertParagraph',
      'deleteParagraph',
      'insertTable',
      'setCellText',
    ],
  };
}
export function editWord(input, operations) {
  check(
    Array.isArray(operations) &&
      operations.length > 0 &&
      operations.length <= 500,
    'Provide 1–500 edit operations.',
  );
  const parts = packageParts(input, 'word', true),
    m = model(parts),
    changes = [];
  const target = (id, local) => {
    const n = m.targets.get(id);
    check(
      n && n.parentNode && (!local || n.localName === local),
      `Unknown or incompatible element ${String(id).slice(0, 100)}.`,
    );
    if (!editable(n))
      throw fail(
        'unsupported_element',
        'Tracked changes, fields and content controls cannot be edited.',
        409,
      );
    return n;
  };
  const append = (node) => {
    const section = child(m.body, 'sectPr')[0];
    m.body.insertBefore(node, section ?? null);
  };
  for (const op of operations) {
    check(op && typeof op === 'object', 'Invalid edit operation.');
    if (op.op === 'replaceText') {
      const n = target(op.id, 'p');
      replaceText(n, W, op.find, op.text, op.expectedMatches ?? 1);
    } else if (op.op === 'setParagraphText') {
      const n = target(op.id, 'p');
      text(op.text);
      check(
        !/[\r\n\t]/.test(op.text),
        'Use separate Word paragraphs for line changes.',
      );
      if (!descendants(n, W, 't').length)
        n.appendChild(paragraph(m.doc, op.text).lastChild);
      else setText(n, W, op.text);
    } else if (op.op === 'insertParagraph') {
      const p = paragraph(m.doc, op.text, op.style ?? 'Normal');
      if (op.after !== undefined) {
        const n = target(op.after, 'p');
        n.parentNode.insertBefore(p, n.nextSibling);
      } else if (op.before !== undefined) {
        const n = target(op.before, 'p');
        n.parentNode.insertBefore(p, n);
      } else append(p);
    } else if (op.op === 'deleteParagraph') {
      const n = target(op.id, 'p');
      if (
        n.parentNode.localName === 'tc' &&
        child(n.parentNode, 'p').length === 1
      )
        throw fail(
          'invalid_structure',
          'A table cell must retain a paragraph.',
        );
      n.parentNode.removeChild(n);
    } else if (op.op === 'insertTable') {
      check(
        Array.isArray(op.rows) &&
          op.rows.length > 0 &&
          op.rows.length <= 200 &&
          op.rows.every(
            (row) => Array.isArray(row) && row.length > 0 && row.length <= 50,
          ),
        'Provide table rows with 1–50 cells and at most 200 rows.',
      );
      check(
        op.rows.every((row) => row.length === op.rows[0].length),
        'Table rows must have the same number of cells.',
      );
      for (const row of op.rows)
        for (const value of row)
          check(
            typeof value === 'string' && !/[\r\n\t]/.test(value),
            'Word table cell text must be single-line.',
          );
      const markup = `<w:tbl xmlns:w="${W}"><w:tblPr><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((edge) => `<w:${edge} w:val="single" w:sz="4"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid>${op.rows[0].map(() => '<w:gridCol w:w="2400"/>').join('')}</w:tblGrid>${op.rows.map((row) => `<w:tr>${row.map((value) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${esc(text(value))}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`;
      const table = m.doc.importNode(xml(markup).documentElement, true);
      if (op.after !== undefined) {
        const n = target(op.after, 'p');
        check(n.parentNode === m.body, 'Insert tables after a body paragraph.');
        m.body.insertBefore(table, n.nextSibling);
      } else append(table);
    } else if (op.op === 'setCellText') {
      const n = target(op.id, 'tc');
      check(
        !child(n, 'tbl').length,
        'Cells containing nested tables cannot be replaced as plain text.',
      );
      text(op.text);
      check(
        !/[\r\n\t]/.test(op.text),
        'Use separate Word paragraphs for line changes.',
      );
      const ps = child(n, 'p');
      if (ps.length) {
        if (descendants(ps[0], W, 't').length) setText(ps[0], W, op.text);
        else ps[0].appendChild(paragraph(m.doc, op.text).lastChild);
        for (const p of ps.slice(1)) n.removeChild(p);
      } else n.appendChild(paragraph(m.doc, op.text));
    } else
      throw fail('unsupported_operation', 'Unsupported Word edit operation.');
    changes.push({ op: op.op, id: op.id ?? null });
  }
  if (!child(m.body, 'p').length) append(paragraph(m.doc, ''));
  parts.set('word/document.xml', serialize(m.doc));
  const output = writeZip(parts);
  readWord(output);
  return { bytes: output, changes };
}
export function createWord({ paragraphs = [{ text: '' }], tables = [] } = {}) {
  check(
    Array.isArray(paragraphs) &&
      paragraphs.length <= 499 &&
      Array.isArray(tables) &&
      tables.length <= 20 &&
      paragraphs.length + tables.length <= 499,
    'Provide at most 499 combined paragraphs/tables, including at most 20 tables.',
  );
  const parts = new Map(
    Object.entries({
      '[Content_Types].xml': `<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${NS.r}/styles" Target="styles.xml"/></Relationships>`,
      'word/document.xml': `<w:document xmlns:w="${W}" xmlns:r="${NS.r}"><w:body><w:p><w:r><w:t/></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
      'word/styles.xml': `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>${['Title', 'Subtitle', ...Array.from({ length: 9 }, (_, i) => `Heading${i + 1}`)].map((style, i) => `<w:style w:type="paragraph" w:styleId="${style}"><w:name w:val="${style}"/><w:basedOn w:val="Normal"/><w:pPr>${i >= 2 ? `<w:outlineLvl w:val="${i - 2}"/>` : ''}</w:pPr><w:rPr><w:b/><w:sz w:val="${Math.max(24, 40 - i * 2)}"/></w:rPr></w:style>`).join('')}</w:styles>`,
    }).map(([name, value]) => [name, Buffer.from(value)]),
  );
  const operations = paragraphs.map((p) => ({
    op: 'insertParagraph',
    text: typeof p === 'string' ? p : p.text,
    style: typeof p === 'string' ? 'Normal' : p.style,
  }));
  operations.push(
    { op: 'deleteParagraph', id: 'p0' },
    ...tables.map((rows) => ({ op: 'insertTable', rows })),
  );
  return editWord(writeZip(parts), operations).bytes;
}

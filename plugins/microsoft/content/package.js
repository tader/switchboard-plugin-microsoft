import { inflateRawSync, deflateRawSync, crc32 } from 'node:zlib';
import { DOMParser, XMLSerializer } from '../vendor/xml.js';
export const MAX_FILE = 25 * 1024 * 1024,
  MAX_EXPANDED = 100 * 1024 * 1024;
export const fail = (code, message, status = 400) =>
  Object.assign(new Error(message), { code, status });
export function check(value, message) {
  if (!value) throw fail('invalid_input', message);
}
export function text(value, name = 'text') {
  check(
    typeof value === 'string' &&
      value.length <= 100_000 &&
      !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value),
    `Provide ${name} as text up to 100000 characters.`,
  );
  return value;
}
export const esc = (value) =>
  String(value).replace(
    /[&<>"']/g,
    (c) =>
      ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&apos;',
      })[c],
  );
export function xml(value) {
  const source =
    typeof value === 'string'
      ? value
      : new TextDecoder('utf-8', { fatal: true }).decode(value);
  if (source.length > 4 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(source))
    throw fail(
      'unsupported_xml',
      'XML entities, document types and XML parts over 4 MiB are unsupported.',
    );
  try {
    return new DOMParser({
      onError: () => {
        throw Error('Invalid XML');
      },
    }).parseFromString(source, 'application/xml');
  } catch {
    throw fail('invalid_package', 'The Office package contains invalid XML.');
  }
}
export const serialize = (doc) =>
  Buffer.from(new XMLSerializer().serializeToString(doc));
export const elements = (node) =>
  Array.from(node.childNodes ?? []).filter((n) => n.nodeType === 1);
export const descendants = (node, ns, local) =>
  Array.from(node.getElementsByTagNameNS(ns, local));
export function parseZip(input) {
  const bytes = Buffer.from(input);
  if (bytes.length > MAX_FILE)
    throw fail('file_too_large', 'Office packages are limited to 25 MiB.', 413);
  try {
    let end = bytes.length - 22;
    for (; end >= Math.max(0, bytes.length - 65557); end--)
      if (
        bytes.readUInt32LE(end) === 0x06054b50 &&
        end + 22 + bytes.readUInt16LE(end + 20) === bytes.length
      )
        break;
    check(
      end >= 0 &&
        bytes.readUInt16LE(end + 4) === 0 &&
        bytes.readUInt16LE(end + 6) === 0,
      'Unsupported ZIP archive.',
    );
    const count = bytes.readUInt16LE(end + 10),
      start = bytes.readUInt32LE(end + 16),
      size = bytes.readUInt32LE(end + 12);
    check(
      count > 0 &&
        count < 10000 &&
        bytes.readUInt16LE(end + 8) === count &&
        start + size === end,
      'Unsupported ZIP directory.',
    );
    const parts = new Map();
    let at = start,
      expanded = 0;
    for (let i = 0; i < count; i++) {
      check(
        at + 46 <= end && bytes.readUInt32LE(at) === 0x02014b50,
        'Invalid ZIP entry.',
      );
      const flags = bytes.readUInt16LE(at + 8),
        method = bytes.readUInt16LE(at + 10),
        crc = bytes.readUInt32LE(at + 16),
        compressed = bytes.readUInt32LE(at + 20),
        original = bytes.readUInt32LE(at + 24),
        namesize = bytes.readUInt16LE(at + 28),
        extra = bytes.readUInt16LE(at + 30),
        comment = bytes.readUInt16LE(at + 32),
        local = bytes.readUInt32LE(at + 42);
      check(
        at + 46 + namesize + extra + comment <= end,
        'Invalid ZIP directory bounds.',
      );
      const name = new TextDecoder('utf-8', { fatal: true }).decode(
        bytes.subarray(at + 46, at + 46 + namesize),
      );
      check(
        !(flags & 1) &&
          [0, 8].includes(method) &&
          !name.startsWith('/') &&
          !name.includes('\\') &&
          !name.includes('\0') &&
          !name.split('/').includes('..') &&
          !parts.has(name),
        'Encrypted, duplicate or unsafe ZIP entries are unsupported.',
      );
      expanded += original;
      if (expanded > MAX_EXPANDED)
        throw fail(
          'file_too_large',
          'Expanded Office packages are limited to 100 MiB.',
          413,
        );
      check(
        local + 30 <= start &&
          bytes.readUInt32LE(local) === 0x04034b50 &&
          bytes.readUInt16LE(local + 8) === method,
        'Invalid ZIP local entry.',
      );
      const data =
        local +
        30 +
        bytes.readUInt16LE(local + 26) +
        bytes.readUInt16LE(local + 28);
      check(
        data + compressed <= start &&
          bytes
            .subarray(local + 30, local + 30 + bytes.readUInt16LE(local + 26))
            .equals(Buffer.from(name)),
        'Invalid ZIP content bounds.',
      );
      const raw = bytes.subarray(data, data + compressed),
        content =
          method === 0
            ? raw
            : inflateRawSync(raw, { maxOutputLength: Math.max(1, original) });
      check(
        content.length === original && crc32(content) === crc,
        'ZIP content failed validation.',
      );
      parts.set(name, Buffer.from(content));
      at += 46 + namesize + extra + comment;
    }
    check(at === end, 'Invalid ZIP directory size.');
    return parts;
  } catch (error) {
    if (error.code === 'file_too_large') throw error;
    throw fail(
      'invalid_package',
      'Unsupported or malformed Office ZIP package.',
    );
  }
}
export function writeZip(parts) {
  const local = [],
    directory = [];
  let offset = 0,
    expanded = 0;
  for (const [name, content] of parts) {
    const raw = Buffer.from(content);
    expanded += raw.length;
    if (expanded > MAX_EXPANDED)
      throw fail(
        'file_too_large',
        'Expanded Office packages are limited to 100 MiB.',
        413,
      );
    const filename = Buffer.from(name),
      compressed = deflateRawSync(raw),
      crc = crc32(raw);
    const header = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50);
    header.writeUInt16LE(20, 4);
    header.writeUInt16LE(0x800, 6);
    header.writeUInt16LE(8, 8);
    header.writeUInt16LE(0x21, 12);
    header.writeUInt32LE(crc, 14);
    header.writeUInt32LE(compressed.length, 18);
    header.writeUInt32LE(raw.length, 22);
    header.writeUInt16LE(filename.length, 26);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50);
    entry.writeUInt16LE(20, 4);
    entry.writeUInt16LE(20, 6);
    entry.writeUInt16LE(0x800, 8);
    entry.writeUInt16LE(8, 10);
    entry.writeUInt16LE(0x21, 14);
    entry.writeUInt32LE(crc, 16);
    entry.writeUInt32LE(compressed.length, 20);
    entry.writeUInt32LE(raw.length, 24);
    entry.writeUInt16LE(filename.length, 28);
    entry.writeUInt32LE(offset, 42);
    local.push(header, filename, compressed);
    directory.push(entry, filename);
    offset += header.length + filename.length + compressed.length;
  }
  const central = Buffer.concat(directory),
    end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(parts.size, 8);
  end.writeUInt16LE(parts.size, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(offset, 16);
  const output = Buffer.concat([...local, central, end]);
  if (output.length > MAX_FILE)
    throw fail('file_too_large', 'Office packages are limited to 25 MiB.', 413);
  return output;
}
export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};
export function packageParts(input, kind, editing = false) {
  const parts = parseZip(input),
    types = xml(parts.get('[Content_Types].xml') ?? '');
  const expected =
    kind === 'word'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
      : 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
  const overrides = descendants(types, NS.ct, 'Override');
  check(
    overrides.some((n) => n.getAttribute('ContentType') === expected),
    `Provide a ${kind === 'word' ? '.docx' : '.pptx'} package.`,
  );
  if (
    [...parts.keys()].some((k) => /vbaProject|_xmlsignatures/i.test(k)) ||
    /macroEnabled/i.test(serialize(types).toString())
  )
    throw fail(
      'unsupported_package',
      'Macro-enabled and signed packages are unsupported.',
    );
  for (const [name, content] of parts)
    if (/\.xml$|\.rels$/.test(name)) xml(content);
  if (editing && kind === 'word' && parts.has('word/settings.xml')) {
    const doc = xml(parts.get('word/settings.xml'));
    if (
      descendants(doc, NS.w, 'documentProtection').some(
        (n) =>
          !['false', '0', 'off'].includes(
            n.getAttributeNS(NS.w, 'enforcement'),
          ),
      )
    )
      throw fail(
        'protected_document',
        'Editing protected Word documents is unsupported.',
        409,
      );
  }
  if (
    editing &&
    kind === 'powerpoint' &&
    parts.has('ppt/presentation.xml') &&
    descendants(xml(parts.get('ppt/presentation.xml')), NS.p, 'modifyVerifier')
      .length
  )
    throw fail(
      'protected_document',
      'Editing password-protected presentations is unsupported.',
      409,
    );
  return parts;
}
export function replaceText(node, ns, find, replacement, expectedMatches = 1) {
  text(find, 'find');
  text(replacement);
  check(
    !/[\r\n\t]/.test(find + replacement),
    'Text replacements are single-line; use paragraph or shape operations for line changes.',
  );
  check(
    find.length > 0 &&
      Number.isInteger(expectedMatches) &&
      expectedMatches >= 1 &&
      expectedMatches <= 1000,
    'Provide a nonempty find and expectedMatches between 1 and 1000.',
  );
  const nodes = [];
  const walk = (current) => {
    if (current.nodeType !== 1) return;
    if (current.namespaceURI === ns && current.localName === 't')
      nodes.push({ node: current });
    else if (
      current.namespaceURI === ns &&
      ['br', 'cr', 'tab'].includes(current.localName)
    )
      nodes.push({ boundary: current.localName === 'tab' ? '\t' : '\n' });
    else for (const child of elements(current)) walk(child);
  };
  walk(node);
  const combined = nodes
      .map((entry) => entry.boundary ?? entry.node.textContent)
      .join(''),
    matches = [];
  for (let at = 0; (at = combined.indexOf(find, at)) >= 0; at += find.length)
    matches.push(at);
  if (matches.length !== expectedMatches)
    throw fail(
      'text_conflict',
      'Text occurrences do not match expectedMatches.',
      409,
    );
  for (const at of matches.reverse()) {
    let position = 0;
    for (const entry of nodes) {
      if (entry.boundary) {
        position += entry.boundary.length;
        continue;
      }
      const n = entry.node;
      const current = n.textContent,
        start = position,
        end = position + current.length;
      position = end;
      if (end <= at || start >= at + find.length) continue;
      const before = current.slice(0, Math.max(0, at - start)),
        after = current.slice(
          Math.min(current.length, at + find.length - start),
        );
      n.textContent = before + (at >= start ? replacement : '') + after;
      n.setAttributeNS(
        'http://www.w3.org/XML/1998/namespace',
        'xml:space',
        'preserve',
      );
    }
  }
  return matches.length;
}
export function setText(node, ns, value) {
  text(value);
  check(!/[\r\n\t]/.test(value), 'Use separate paragraphs for line changes.');
  for (const tag of ['br', 'cr', 'tab'])
    for (const n of descendants(node, ns, tag)) n.parentNode.removeChild(n);
  const nodes = descendants(node, ns, 't');
  check(nodes.length > 0, 'This element has no editable text runs.');
  nodes[0].textContent = value;
  nodes[0].setAttributeNS(
    'http://www.w3.org/XML/1998/namespace',
    'xml:space',
    'preserve',
  );
  for (const n of nodes.slice(1)) n.textContent = '';
}

import { AdapterError } from './errors.js';
import { recipientEmail } from './input.js';

// Structured content only: never accept caller-supplied HTML or mention markup.
export function richContent(value) {
  const bad = () => { throw new AdapterError('invalid_content', 'content must contain 1–100 blocks with text runs. Supported blocks: paragraph, quote, bulletedList, numberedList. Supported marks: bold, italic, underline, strike, code. Person mentions require exact directory emails.', 400); };
  const object = (v, keys) => v && typeof v === 'object' && !Array.isArray(v) && Object.keys(v).every(k => keys.includes(k));
  if (!Array.isArray(value) || !value.length || value.length > 100) bad();
  let size = 0, count = 0;
  const runs = list => {
    if (!Array.isArray(list) || !list.length || list.length > 100) bad();
    return list.map(run => {
      if (!object(run, ['text', 'marks', 'mention', 'link']) || ++count > 500) bad();
      if (run.mention !== undefined) {
        if (run.text !== undefined || run.marks !== undefined || run.link !== undefined || !object(run.mention, ['email', 'name']) || typeof run.mention.name !== 'string' || !run.mention.name.trim() || run.mention.name.length > 150 || /[\u0000-\u001f<>]/.test(run.mention.name)) bad();
        size += run.mention.name.length;
        return { mention: { email: recipientEmail(run.mention.email), name: run.mention.name.trim() } };
      }
      if (typeof run.text !== 'string' || !run.text || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(run.text)) bad();
      const marks = run.marks ?? [];
      if (!Array.isArray(marks) || marks.some(x => !['bold', 'italic', 'underline', 'strike', 'code'].includes(x)) || new Set(marks).size !== marks.length) bad();
      size += run.text.length;
      let link;
      if (run.link !== undefined) { try { if (typeof run.link !== 'string' || run.link.length > 2000) bad(); const url = new URL(run.link); if (!['https:', 'http:'].includes(url.protocol) || url.username || url.password) bad(); link = url.href; } catch { bad(); } }
      return { text: run.text, marks: [...marks].sort(), ...(link ? { link } : {}) };
    });
  };
  const content = value.map(block => {
    if (!object(block, ['type', 'runs', 'items']) || !['paragraph', 'quote', 'bulletedList', 'numberedList'].includes(block.type)) bad();
    if (['paragraph', 'quote'].includes(block.type)) { if (block.items !== undefined) bad(); return { type: block.type, runs: runs(block.runs) }; }
    if (block.runs !== undefined || !Array.isArray(block.items) || !block.items.length || block.items.length > 100) bad();
    return { type: block.type, items: block.items.map(runs) };
  });
  if (!size || size > 20_000) bad();
  return content;
}

export function contentText(content) {
  const runs = list => list.map(run => run.mention ? run.mention.name : run.text).join('');
  return content.map(block => block.runs ? runs(block.runs) : block.items.map(runs).join('\n')).join('\n');
}

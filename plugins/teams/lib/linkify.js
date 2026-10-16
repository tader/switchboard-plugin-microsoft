// Detect bare HTTP(S) URLs in text, never in HTML. Callers escape every segment
// before rendering and skip runs with explicit links or code formatting.
export function linkedText(text) {
  const parts = []; let position = 0;
  for (const match of text.matchAll(/\bhttps?:\/\/[^\s<>"`]+/gi)) {
    const start = match.index;
    let candidate = match[0], end = candidate.length;
    // Sentence punctuation and unmatched wrapping brackets belong to the text.
    const counts = { '(': 0, ')': 0, '[': 0, ']': 0, '{': 0, '}': 0 };
    const opening = { ')': '(', ']': '[', '}': '{' };
    for (const character of candidate) if (Object.hasOwn(counts, character)) counts[character]++;
    while (end) {
      const last = candidate[end - 1];
      if (/[.,!?;:'\u2019\u201d]/.test(last)) end--;
      else if (opening[last] && counts[last] > counts[opening[last]]) { counts[last]--; end--; }
      else break;
    }
    candidate = candidate.slice(0, end);
    let url;
    try {
      if (candidate.length > 2000 || candidate.includes('\\')) continue;
      url = new URL(candidate);
      if (!['https:', 'http:'].includes(url.protocol) || !url.hostname || url.username || url.password) continue;
    } catch { continue; }
    if (start > position) parts.push({ text: text.slice(position, start) });
    parts.push({ text: candidate, link: url.href });
    position = start + candidate.length;
  }
  if (position < text.length) parts.push({ text: text.slice(position) });
  return parts;
}

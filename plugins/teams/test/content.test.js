import test from 'node:test';
import assert from 'node:assert/strict';
import { richContent, contentText } from '../lib/content.js';

test('structured messages canonicalize formatting and person identity without accepting arbitrary markup', () => {
  const content = richContent([{ type: 'paragraph', runs: [{ text: '<hello>', marks: ['italic', 'bold'] }, { mention: { email: 'Alice@Example.COM', name: 'Alice Example' } }] }, { type: 'numberedList', items: [[{ text: 'One', link: 'https://example.com' }], [{ text: 'Two' }]] }]);
  assert.deepEqual(content[0].runs[0].marks, ['bold', 'italic']);
  assert.equal(content[0].runs[1].mention.email, 'alice@example.com');
  assert.equal(contentText(content), '<hello>Alice Example\nOne\nTwo');
  for (const value of [[], [{ type: 'html', runs: [{ text: '<script>' }] }], [{ type: 'paragraph', runs: [{ text: 'x', marks: ['script'] }] }], [{ type: 'paragraph', runs: [{ text: 'x', link: 'javascript:alert(1)' }] }], [{ type: 'paragraph', runs: [{ mention: { email: 'everyone', name: 'Everyone' } }] }], [{ type: 'paragraph', runs: [{ text: 'x', mention: { email: 'alice@example.com', name: 'Alice' } }] }], [{ type: 'paragraph', runs: [{ text: 'x'.repeat(20001) }] }]]) assert.throws(() => richContent(value));
});

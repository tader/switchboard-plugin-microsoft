import test from 'node:test';
import assert from 'node:assert/strict';
import { recipientEmail, peopleQuery, messageText, reactionInput } from '../lib/input.js';

test('recipient identity accepts one email/UPN and rejects names, groups and URLs', () => {
  assert.equal(recipientEmail(' Alice+Teams@Example.COM '), 'alice+teams@example.com');
  for (const value of ['Alice Adams', 'alice@example.com,bob@example.com', 'alice@example.com bob@example.com', 'https://example.com', null]) assert.throws(() => recipientEmail(value), { code: 'invalid_recipient' });
});

test('people search is bounded and requires meaningful text', () => {
  assert.equal(peopleQuery(' Alice Adams '), 'Alice Adams');
  for (const value of ['a', '', 'a\0b', 'a'.repeat(151), null]) assert.throws(() => peopleQuery(value), { code: 'invalid_query' });
});


test('mutation payloads require plain text and explicit reaction state', () => {
  assert.equal(messageText('', 'expectedText', true), '');
  assert.throws(() => messageText(''), { code: 'invalid_message' });
  assert.deepEqual(reactionInput('yes', true), { reaction: 'like', selected: true });
  for (const value of ['Like', '👍', 'button[evil]', '', 'a'.repeat(65)]) assert.throws(() => reactionInput(value, true), { code: 'invalid_reaction' });
  assert.throws(() => reactionInput('like', undefined), { code: 'invalid_reaction' });
});

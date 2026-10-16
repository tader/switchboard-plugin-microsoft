import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { SendLedger } from '../lib/ledger.js';

test('send intent survives restart, replays result, and binds key to content', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  let sends = 0;
  const result = { status: 'observed_in_chat', message: { id: 'm2' } };
  const ledger = new SendLedger(dir);
  await ledger.send('request-001', 'chat1', 'hello', async () => 'prepared', async prepared => {
    assert.equal(prepared, 'prepared');
    const records = await ledger.read();
    assert.equal(Object.values(records).length, 1);
    assert.equal(Object.values(records)[0].result, undefined);
    sends++; return result;
  });
  assert.deepEqual(await new SendLedger(dir).send('request-001', 'chat1', 'hello', () => assert.fail('prepare repeated'), () => assert.fail('send repeated')), { ...result, replayed: true });
  await assert.rejects(() => ledger.send('request-001', 'chat2', 'hello', () => {}, () => {}), { code: 'idempotency_conflict' });
  await assert.rejects(() => ledger.send('request-001', 'chat1', 'changed', () => {}, () => {}), { code: 'idempotency_conflict' });
  assert.equal(sends, 1);
});

test('uncertain attempts never send again, including after a crash', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir);
  await assert.rejects(() => ledger.send('request-002', 'chat1', 'hello', async () => ({}), async () => { throw new Error('browser crashed'); }), /browser crashed/);
  await assert.rejects(() => new SendLedger(dir).send('request-002', 'chat1', 'hello', () => assert.fail(), () => assert.fail()), { code: 'send_uncertain' });
});

test('preflight failures remain retryable and invalid payloads never reach browser', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir);
  await assert.rejects(() => ledger.send('request-003', 'chat1', 'hello', async () => { throw new Error('draft exists'); }, () => assert.fail()), /draft exists/);
  assert.deepEqual(await ledger.read(), {});
  for (const [key, text] of [['short', 'hello'], ['request-003', ' '], ['request-003', '\0'], ['request-003', 'a'.repeat(20001)]]) {
    await assert.rejects(() => ledger.send(key, 'chat1', text, () => assert.fail(), () => assert.fail()), error => error.status === 400);
  }
});

test('corrupted durable state stops sends rather than resetting duplicate protection', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-ledger-');
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir);
  await fs.writeFile(ledger.file, 'broken JSON');
  await assert.rejects(() => ledger.send('request-004', 'chat1', 'hello', () => assert.fail(), () => assert.fail()), SyntaxError);
});


test('mutation intent binds operation, target and payload, survives restart, and releases preparation', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-mutation-ledger-'); t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir); let executions = 0, cleanups = 0;
  const target = ['reaction', 'chat1', 'message1'], payload = { reaction: 'like', selected: true };
  const prepare = async () => ({ cleanup: async () => cleanups++ });
  await ledger.mutate('mutation-001', target, payload, prepare, async () => { executions++; return { selected: true }; });
  assert.equal((await new SendLedger(dir).mutate('mutation-001', target, payload, () => assert.fail(), () => assert.fail())).replayed, true);
  assert.equal(executions, 1); assert.equal(cleanups, 1);
  await assert.rejects(ledger.mutate('mutation-001', ['delete', 'chat1', 'message1'], {}, () => assert.fail(), () => assert.fail()), { code: 'idempotency_conflict' });
  await assert.rejects(ledger.mutate('mutation-002', ['delete', 'chat1', 'message1'], {}, prepare, async () => { throw new Error('connection lost after click'); }), /connection lost/);
  assert.equal(cleanups, 2);
  await assert.rejects(new SendLedger(dir).mutate('mutation-002', ['delete', 'chat1', 'message1'], {}, () => assert.fail(), () => assert.fail()), { code: 'mutation_uncertain' });
});

test('read-state generations are separate from operation keys and uncertainty persists before dispatch', async t => {
  const dir = await fs.mkdtemp('/tmp/teams-read-ledger-'); t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const ledger = new SendLedger(dir), guard = { account: 'a', threadId: 'chat', generation: 0, unread: true };
  await assert.rejects(ledger.run('unread-ledger-001', ['unread'], async () => ({ readStateGuard: guard }), async () => {
    const state = (await new SendLedger(dir).readStates('a')).get('chat');
    assert.deepEqual(state, { account: 'a', threadId: 'chat', generation: 1, blocked: true });
    throw new Error('crashed');
  }, 'mutation_uncertain'), /crashed/);
  assert.equal((await new SendLedger(dir).readStates('a')).get('chat').blocked, true);
  assert.equal((await ledger.readStates('b')).size, 0);
  const metadataKey = Object.keys(await ledger.read()).find(key => key.startsWith('read-state:'));
  await ledger.run(metadataKey, ['send'], async () => ({}), async () => ({ status: 'sent' }), 'send_uncertain');
  assert.equal((await ledger.readStates('a')).get('chat').generation, 1);
  await assert.rejects(ledger.run('stale-generation-001', ['read'], async () => ({ readStateGuard: { ...guard, unread: false } }), () => assert.fail(), 'mutation_uncertain'), { code: 'read_state_changed' });
  const result = await ledger.run('fresh-generation-001', ['read'], async () => ({ readStateGuard: { ...guard, generation: 1, unread: false } }), async () => ({ status: 'updated' }), 'mutation_uncertain');
  assert.equal((await ledger.readStates('a')).get('chat').blocked, false);
  assert.equal((await ledger.readStates('a')).get('chat').generation, 1);
  assert.deepEqual(await new SendLedger(dir).run('fresh-generation-001', ['read'], () => assert.fail(), () => assert.fail(), 'mutation_uncertain'), { ...result, replayed: true });
});

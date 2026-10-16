import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { PassThrough } from 'node:stream';
import { CDP, Page } from '../lib/cdp.js';

function transport(t) {
  const child = new EventEmitter();
  child.stdio = [null, null, null, new PassThrough(), new PassThrough()];
  child.exitCode = null;
  child.kill = () => { child.exitCode = 0; child.emit('exit', 0); };
  const requests = [];
  child.stdio[3].on('data', chunk => requests.push(JSON.parse(chunk.toString().replace(/\0$/, ''))));
  const cdp = new CDP('fake', '/tmp/fake', { spawnProcess: (_exe, args, opts) => {
    assert.ok(args.includes('--remote-debugging-pipe'));
    assert.ok(!args.some(arg => arg.startsWith('--remote-debugging-port')));
    assert.equal(opts.stdio.length, 5);
    return child;
  } });
  t.after(() => child.kill());
  return { child, cdp, requests, reply: message => child.stdio[4].write(JSON.stringify(message) + '\0') };
}

test('private pipe handles fragmented frames, interleaved replies and target sessions', async t => {
  const f = transport(t);
  const first = f.cdp.call('Runtime.evaluate', {}, 'session-a');
  const second = f.cdp.call('Runtime.evaluate', {}, 'session-b');
  assert.equal(f.requests[0].sessionId, 'session-a');
  f.reply({ id: f.requests[1].id, result: { value: 'second' } });
  const bytes = JSON.stringify({ id: f.requests[0].id, result: { value: 'first' } }) + '\0';
  f.child.stdio[4].write(bytes.slice(0, 9)); f.child.stdio[4].write(bytes.slice(9));
  assert.deepEqual(await first, { value: 'first' });
  assert.deepEqual(await second, { value: 'second' });
});

test('disconnects and protocol timeouts reject pending commands', async t => {
  const f = transport(t);
  await assert.rejects(f.cdp.call('hang', {}, undefined, 5), { code: 'browser_timeout' });
  const pending = f.cdp.call('pending');
  f.child.kill();
  await assert.rejects(pending, { code: 'browser_closed' });
  await assert.rejects(f.cdp.call('later'), { code: 'browser_closed' });
});

test('page evaluations surface DOM errors and escape input as data', async t => {
  const f = transport(t), page = new Page(f.cdp, 'page-session', 'target');
  const input = '"; window.sideEffect = true; //';
  const pending = page.evaluate(text => text, input);
  assert.ok(f.requests[0].params.expression.includes(JSON.stringify([input])));
  f.reply({ id: f.requests[0].id, result: { result: { value: { error: { code: 'draft_exists', message: 'Draft exists', status: 409 } } } } });
  await assert.rejects(pending, { code: 'draft_exists', status: 409 });
});

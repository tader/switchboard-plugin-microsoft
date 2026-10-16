import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { AdapterError } from './errors.js';

// Durable intent BEFORE any editor input or send click. A process crash must never
// turn a repeated request into a second message. Called inside the profile queue.
export class SendLedger {
  constructor(directory) { this.file = path.join(directory, 'send-ledger.json'); }
  async read() {
    try { return JSON.parse(await fs.readFile(this.file, 'utf8')); }
    catch (error) { if (error.code === 'ENOENT') return {}; throw error; }
  }
  async write(records) {
    await fs.mkdir(path.dirname(this.file), { recursive: true, mode: 0o700 });
    const temp = `${this.file}.${randomUUID()}.tmp`;
    const handle = await fs.open(temp, 'wx', 0o600);
    try { await handle.writeFile(JSON.stringify(records)); await handle.sync(); }
    finally { await handle.close(); }
    await fs.rename(temp, this.file);
    const directory = await fs.open(path.dirname(this.file), 'r');
    try { await directory.sync(); } finally { await directory.close(); }
  }
  async send(key, chatId, text, prepare, execute) {
    if (typeof key !== 'string' || !/^[A-Za-z0-9._:-]{8,128}$/.test(key)) throw new AdapterError('invalid_idempotency_key', 'idempotencyKey must be 8–128 letters, digits or . _ : - characters.', 400);
    if (typeof text !== 'string' || !text.trim() || text.length > 20_000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)) throw new AdapterError('invalid_message', 'text must be nonempty plain text, at most 20,000 characters, without control characters.', 400);
    return this.run(key, [chatId, text], prepare, execute, 'send_uncertain');
  }
  async mutate(key, target, payload, prepare, execute, verify) {
    return this.run(key, ['mutation', target, payload], prepare, execute, 'mutation_uncertain', verify);
  }
  async sendContent(key, target, content, prepare, execute) {
    return this.run(key, ['structured-send', target, content], prepare, execute, 'send_uncertain');
  }
  async readStates(account) {
    const states = new Map();
    for (const record of Object.values(await this.read())) {
      const state = record.readState;
      if (state?.account !== account) continue;
      if (typeof state.threadId !== 'string' || !state.threadId || !Number.isSafeInteger(state.generation) || state.generation < 0 || typeof state.blocked !== 'boolean') {
        throw new AdapterError('ledger_invalid', 'Read-state protection could not be loaded.', 503);
      }
      states.set(state.threadId, state);
    }
    return states;
  }
  // Separate durable read stages: never replace or discard an accepted send.
  // Invoked inside the same profile queue as run(). Keys remain hashed on disk.
  async updateTriageReads(account, update) {
    const records = await this.read();
    const entries = Object.entries(records).filter(([, record]) => record.result?._triageRead?.account === account)
      .map(([key, record]) => ({ key, intent: record.result._triageRead, read: record.triageRead }));
    const outcomes = await update(entries);
    let changed = false;
    for (const [key, outcome] of outcomes) {
      if (!records[key]) continue;
      if (JSON.stringify(records[key].triageRead) !== JSON.stringify(outcome)) { records[key].triageRead = outcome; changed = true; }
    }
    if (changed) await this.write(records);
    return outcomes;
  }
  async verifyRecord(records, index, verify) {
    const record = records[index];
    if (!verify || record.result.verification?.status === 'observed') return record.result;
    let verification;
    try { verification = await verify(record.result, record.confirmation); }
    catch { verification = { status: 'unavailable', reason: 'verification_failed' }; }
    const result = { ...record.result, verification };
    // Acceptance is already durable. A failed read or second persistence must
    // never hide that receipt or turn a replay into a second write.
    record.result = result;
    try { await this.write(records); }
    catch { return { ...result, verification: { status: 'unavailable', reason: 'verification_persistence_failed' } }; }
    return result;
  }
  async run(key, identity, prepare, execute, uncertainCode, verify) {
    if (typeof key !== 'string' || !/^[A-Za-z0-9._:-]{8,128}$/.test(key)) throw new AdapterError('invalid_idempotency_key', 'idempotencyKey must be 8–128 letters, digits or . _ : - characters.', 400);
    const records = await this.read();
    const index = createHash('sha256').update(key).digest('hex');
    const fingerprint = createHash('sha256').update(JSON.stringify(identity)).digest('hex');
    const previous = records[index];
    if (previous) {
      if (previous.fingerprint !== fingerprint) throw new AdapterError('idempotency_conflict', 'This key was already used for a different operation or content.');
      if (previous.result) return { ...await this.verifyRecord(records, index, verify), replayed: true };
      throw new AdapterError(uncertainCode, 'This operation was already attempted. Check Teams manually; do not retry with a new key.');
    }
    if (Object.keys(records).length >= 10_000) throw new AdapterError('ledger_full', 'Operation ledger is full. Archive it manually after checking outstanding sends.', 503);
    const prepared = await prepare();
    try {
      records[index] = { fingerprint, attemptedAt: new Date().toISOString() };
      // An unread intent supersedes older automatic read grants BEFORE the
      // request. A crash/timeout keeps the barrier blocked until a fresh explicit
      // operation verifies state. Generation metadata shares the atomic journal.
      let stateIndex;
      if (prepared?.readStateGuard) {
        const { account, threadId, generation, unread } = prepared.readStateGuard;
        // Namespace metadata separately from hashes of caller-supplied keys.
        stateIndex = `read-state:${createHash('sha256').update(JSON.stringify([account, threadId])).digest('hex')}`;
        if (!records[stateIndex] && Object.keys(records).length >= 10_000) throw new AdapterError('ledger_full', 'Operation ledger is full. Archive it manually after checking outstanding sends.', 503);
        const current = records[stateIndex]?.readState ?? { account, threadId, generation: 0, blocked: false };
        if (current.generation !== generation || unread && generation === Number.MAX_SAFE_INTEGER) throw new AdapterError('read_state_changed', 'Read state changed. Fetch fresh context.', 409);
        records[stateIndex] = { readState: { ...current, generation: generation + (unread ? 1 : 0), blocked: unread || current.blocked } };
      }
      await this.write(records);
      const result = await execute(prepared);
      records[index].result = result;
      if (prepared?.confirmation) records[index].confirmation = prepared.confirmation;
      if (stateIndex && ['updated', 'unchanged'].includes(result.status)) records[stateIndex].readState.blocked = false;
      await this.write(records);
      return await this.verifyRecord(records, index, verify);
    } finally { await prepared?.cleanup?.(); }
  }
}

import fs from 'node:fs/promises';
import path from 'node:path';
import { TeamsSession } from './api-auth.js';
import { SendLedger } from './ledger.js';
import { AdapterError } from './errors.js';
import { recipientEmail } from './input.js';

// Shared by live diagnostics only; production credentials remain in Switchboard.
// Keep refresh rotations on disk even when a later exchange fails.
export class DiagnosticSession extends TeamsSession {
  constructor(directory, settings = {}, options = {}) {
    super('diagnostic', directory, settings, options);
    this.storage = new SendLedger(directory);
    this.storage.file = path.join(directory, 'diagnostic-session.json');
    this.lock = path.join(directory, 'diagnostic-session.lock');
    const refresh = this.refresh;
    this.refresh = (captured, options) => refresh(captured, { ...options, onRotate: async renewal => {
      await options.onRotate?.(renewal);
      await this.save();
    } });
  }
  async save() { if (this.credentials) await this.storage.write(this.credentials); }
  async authenticate(interactive, loginHint) {
    const result = await super.authenticate(interactive, loginHint);
    await this.save(); return result;
  }
  async ensureSearch() { try { return await super.ensureSearch(); } finally { await this.save(); } }
  async renew() { try { return await super.renew(); } finally { await this.save(); } }
  async acquire({ silent = false, loginHint, renew = false, capture = false } = {}) {
    if (loginHint) loginHint = recipientEmail(loginHint);
    await fs.mkdir(this.directory, { recursive: true, mode: 0o700 });
    try { this.lockHandle = await fs.open(this.lock, 'wx', 0o600); }
    catch (error) {
      if (error.code === 'EEXIST') throw new AdapterError('diagnostic_session_busy', 'Another diagnostic holds the shared session. Run checks sequentially; remove a stale diagnostic-session.lock only after that process has exited.', 409);
      throw error;
    }
    await this.lockHandle.writeFile(String(process.pid));
    let cached;
    try { cached = JSON.parse(await fs.readFile(this.storage.file, 'utf8')); }
    catch (error) { if (error.code !== 'ENOENT') throw new AdapterError('diagnostic_session_invalid', 'The saved diagnostic session could not be loaded.', 503); }
    if (cached) {
      if (cached.profile !== this.profile || cached.authMode !== 'browser' || cached.authVersion !== 1 ||
          !cached.identity?.tenant || !cached.identity?.oid || !cached.tokens?.skype?.value || !cached.tokens?.chatsvcagg?.value) throw new AdapterError('diagnostic_session_invalid', 'The saved diagnostic session is invalid.', 503);
      if (loginHint && recipientEmail(cached.identity.email) !== loginHint) throw new AdapterError('account_mismatch', 'The saved diagnostic session belongs to a different account.', 403);
      this.adopt(cached);
    }
    if (!this.credentials || capture) return this.authenticate(!silent, loginHint);
    try { const result = await this.ensure(renew); await this.save(); return result; }
    catch (error) {
      if (!silent && error.code === 'signin_required') return this.authenticate(true, loginHint);
      throw error;
    }
  }
  async close() {
    try { await super.close(); }
    finally {
      const handle = this.lockHandle; this.lockHandle = null;
      if (handle) {
        await handle.close();
        await fs.unlink(this.lock);
      }
    }
  }
}

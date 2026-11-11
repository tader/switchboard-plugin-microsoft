import fs from "node:fs/promises";
import path from "node:path";
import {
  createHmac,
  randomBytes,
  randomUUID,
  timingSafeEqual,
} from "node:crypto";
const fail = (code: string, status = 400): never => {
  throw Object.assign(Error(code), { code, status });
};
export class Cursors {
  root: string;
  pending?: Promise<Buffer>;
  constructor(root: string) {
    this.root = root;
  }
  async key() {
    return (this.pending ??= (async () => {
      await fs.mkdir(this.root, { recursive: true, mode: 0o700 });
      const target = path.join(this.root, "email-cursors.key");
      try {
        const key = await fs.readFile(target);
        if (key.length !== 32) fail("cursor_key_invalid", 503);
        return key;
      } catch (e: any) {
        if (e.code !== "ENOENT") throw e;
      }
      const temp = target + "." + randomUUID(),
        key = randomBytes(32);
      const handle = await fs.open(temp, "wx", 0o600);
      try {
        await handle.writeFile(key);
        await handle.sync();
      } finally {
        await handle.close();
      }
      try {
        try {
          await fs.link(temp, target);
        } catch (e: any) {
          if (e.code !== "EEXIST") throw e;
        }
        const dir = await fs.open(this.root, "r");
        try {
          await dir.sync();
        } finally {
          await dir.close();
        }
      } finally {
        await fs.unlink(temp);
      }
      const saved = await fs.readFile(target);
      if (saved.length !== 32) fail("cursor_key_invalid", 503);
      return saved;
    })());
  }
  async encode(binding: unknown, token: string) {
    const key = await this.key();
    const scope = createHmac("sha256", key)
      .update(JSON.stringify(binding))
      .digest("base64url");
    const payload = Buffer.from(JSON.stringify([scope, token])).toString(
      "base64url",
    );
    return (
      payload +
      "." +
      createHmac("sha256", key).update(payload).digest("base64url")
    );
  }
  async decode(binding: unknown, cursor?: unknown) {
    if (cursor === undefined) return undefined;
    const key = await this.key();
    try {
      if (typeof cursor !== "string" || cursor.length > 65000) throw Error();
      const [payload, mac, extra] = cursor.split(".");
      const expected = createHmac("sha256", key).update(payload).digest(),
        actual = Buffer.from(mac, "base64url");
      if (
        extra !== undefined ||
        actual.length !== expected.length ||
        !timingSafeEqual(actual, expected)
      )
        throw Error();
      const [scope, token] = JSON.parse(
        Buffer.from(payload, "base64url").toString(),
      );
      if (
        scope !==
          createHmac("sha256", key)
            .update(JSON.stringify(binding))
            .digest("base64url") ||
        typeof token !== "string"
      )
        throw Error();
      return token;
    } catch {
      return fail("invalid_cursor", 409);
    }
  }
}

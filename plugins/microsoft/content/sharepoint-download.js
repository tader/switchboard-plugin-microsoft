import http from 'node:http';
import { randomBytes } from 'node:crypto';
import { GraphClient } from './graph.js';

// An invocation key binds a local request to the caller's Graph token and exact URL.
export async function createSharePointDownload(ctx, graphService, { fetcher = fetch } = {}) {
  const pending = new Map(), controllers = new Set();
  let disposed = false;
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, 'http://adapter.localhost');
    const key = /^Bearer ([A-Za-z0-9_-]{43})$/.exec(req.headers.authorization ?? '')?.[1];
    const invocation = key && pending.get(key);
    if (key) pending.delete(key);
    const match = /^\/files\/([^/]+)\/([^/]+)\/content$/.exec(url.pathname);
    if (!invocation || invocation.expires < Date.now() || invocation.path !== url.pathname + url.search || req.method !== 'GET' || !match || url.search) {
      res.writeHead(403).end(); return;
    }
    const controller = new AbortController(); controllers.add(controller);
    res.once('close', () => controller.abort());
    try {
      const ids = match.slice(1).map(decodeURIComponent);
      if (ids.some(id => !id || id.length > 512 || /[\/\\\x00-\x1f]/.test(id))) throw Object.assign(new Error('Invalid file identifier.'), { status: 400 });
      const graph = new GraphClient(invocation.token, { access: 'all', fetcher, signal: invocation.signal ? AbortSignal.any([controller.signal, invocation.signal]) : controller.signal });
      const bytes = await graph.download(ids[0], ids[1]);
      res.writeHead(200, { 'content-type': 'application/octet-stream', 'content-length': bytes.length, 'cache-control': 'no-store' }); res.end(bytes);
    } catch (error) {
      if (!res.destroyed) { res.writeHead(Number.isInteger(error.status) ? error.status : 500, { 'content-type': 'application/json', 'cache-control': 'no-store' }); res.end(JSON.stringify({ error: { code: error.code || 'adapter_error', message: error.status ? error.message : 'Download failed.' } })); }
    } finally { controllers.delete(controller); }
  });
  server.requestTimeout = 120000; server.headersTimeout = 15000;
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
  const origin = `http://127.0.0.1:${server.address().port}`, placeholder = 'http://sharepoint-download.localhost';
  const prune = setInterval(() => { for (const [key, value] of pending) if (value.expires < Date.now()) pending.delete(key); }, 10000); prune.unref();
  const methods = graphService.authMethods.map(auth => {
    const { reusable, token, ...base } = auth;
    if (!token) throw new Error('SharePoint download requires token support');
    return { ...base, async authorize(request, conn, opts) {
      if (disposed || (request.url.origin !== placeholder && request.url.origin !== origin) || request.method !== 'GET' || !/^\/files\/[^/]+\/[^/]+\/content$/.test(request.url.pathname) || request.url.search) throw Object.assign(new Error('Unsupported download request.'), { code: 'request_forbidden', status: 403 });
      if (pending.size >= 1000) throw Object.assign(new Error('Too many pending downloads.'), { code: 'adapter_busy', status: 429 });
      const result = await token(conn, opts);
      if (opts.signal?.aborted) throw Object.assign(new Error('Request cancelled.'), { code: 'cancelled', status: 499 });
      const key = randomBytes(32).toString('base64url');
      pending.set(key, { token: result.accessToken, path: request.url.pathname + request.url.search, expires: Date.now() + 10000, signal: opts.signal });
      request.url.host = new URL(origin).host; request.headers.set('authorization', `Bearer ${key}`);
      return result.credentials ? { credentials: result.credentials } : undefined;
    } };
  });
  const path = '/files/{drive}/{item}/content';
  const spec = { openapi: '3.0.3', info: { title: 'SharePoint file download', version: '1.0.0' }, servers: [{ url: placeholder }], paths: { [path]: { get: { operationId: 'downloadSharePointFile', summary: 'Download a file from a SharePoint library (maximum 25 MiB)', parameters: ['drive','item'].map(name => ({ name, in: 'path', required: true, schema: { type: 'string' } })), responses: { 200: { description: 'File bytes', content: { 'application/octet-stream': { schema: { type: 'string', format: 'binary' } } } } } } } } };
  return { service: { id: 'sharepoint-download', name: 'SharePoint File Download', description: 'Download SharePoint files securely', icon: 'icon.svg', baseUrl: placeholder, allowedHosts: [new URL(placeholder).host, new URL(origin).host], openapi: spec, authMethods: methods }, async dispose() { disposed = true; clearInterval(prune); pending.clear(); for (const controller of controllers) controller.abort(); await new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }); } };
}

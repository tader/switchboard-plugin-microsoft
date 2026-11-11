import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { allowRequest } from '../plugins/microsoft-authentication/consumer.ts';
import { createSharePointDownload } from '../plugins/microsoft/content/sharepoint-download.js';
import setup from '../plugins/sharepoint/index.ts';

test('SharePoint registers read-first scopes and both services', async () => {
  let options;
  const instance = await setup({ require: () => ({ microsoftService(o) { options = o; return { id: o.id }; }, createSharePointDownload: async () => ({ service: { id: 'sharepoint-download' }, dispose: async () => {} }) }) });
  assert.deepEqual(instance.services.map(x => x.id), ['sharepoint', 'sharepoint-download']);
  assert.deepEqual(options.scopes({}), ['Sites.Read.All']);
  assert.deepEqual(options.scopes({access:'full'}), ['Sites.ReadWrite.All']);
  assert.equal(options.spec, 'sharepoint');
});
test('browser credentials restrict SharePoint paths and read mode', () => {
  const req = (method, path) => ({ method, url: new URL('https://graph.microsoft.com/v1.0' + path) });
  const conn = { methodId: 'authentication-source', config: {} };
  assert.ok(allowRequest('sharepoint', req('GET', '/sites/a/pages'), conn, []));
  assert.ok(allowRequest('sharepoint', req('GET', '/drives/a/items/b'), conn, []));
  assert.ok(!allowRequest('sharepoint', req('PATCH', '/sites/a'), conn, []));
  assert.ok(!allowRequest('sharepoint', req('GET', '/me/messages'), conn, []));
  assert.ok(!allowRequest('sharepoint', req('GET', '/sites/a/$batch'), conn, []));
  conn.config.access = 'full';
  assert.ok(allowRequest('sharepoint', req('PATCH', '/sites/a/lists/b/items/c/fields'), conn, []));
});
test('reference exposes page content, library and list operations', async () => {
  const spec = JSON.parse(await fs.readFile(new URL('../plugins/microsoft/openapi/sharepoint.json', import.meta.url)));
  for (const [method,path] of [['get','/sites'],['get','/sites/{site-id}/pages/{page-id}/microsoft.graph.sitePage'],['get','/sites/{site-id}/drives'],['get','/drives/{drive-id}/items/{item-id}/children'],['get','/sites/{site-id}/lists/{list-id}/items'],['post','/sites/{site-id}/lists/{list-id}/items'],['patch','/sites/{site-id}/lists/{list-id}/items/{item-id}/fields']]) assert.ok(spec.paths[path]?.[method], `${method} ${path}`);
});
test('file download follows trusted redirects without forwarding the Graph token', async () => {
  const calls = [];
  const adapter = await createSharePointDownload({}, { authMethods: [{ id:'test', token: async () => ({ accessToken:'secret' }) }] }, { fetcher: async (url, opts) => {
    calls.push([String(url), opts]);
    return calls.length === 1 ? new Response(null, { status:302, headers: { location:'https://tenant.sharepoint.com/file?token=secret' } }) : new Response('file bytes');
  } });
  try {
    const req = { url: new URL('http://sharepoint-download.localhost/files/drive/item/content'), method:'GET', headers:new Headers() };
    await adapter.service.authMethods[0].authorize(req, { config:{} }, {});
    const result = await fetch(req.url, { headers:req.headers });
    assert.equal(await result.text(), 'file bytes');
    assert.equal(calls[0][1].headers.authorization, 'Bearer secret');
    assert.equal(calls[1][1].headers, undefined);
    assert.equal((await fetch(req.url, { headers:req.headers })).status, 403);
  } finally { await adapter.dispose(); }
});
test('file adapter rejects untrusted redirects and unsupported requests', async () => {
  const adapter = await createSharePointDownload({}, { authMethods: [{ id:'test', token: async () => ({ accessToken:'secret' }) }] }, { fetcher: async () => new Response(null, { status:302, headers:{ location:'https://attacker.example/steal' } }) });
  try {
    const method = adapter.service.authMethods[0];
    const bad = { url:new URL('http://sharepoint-download.localhost/files/d/i/content'), method:'POST', headers:new Headers() };
    await assert.rejects(method.authorize(bad, { config:{} }, {}), { code:'request_forbidden' });
    const req = { url:new URL('http://sharepoint-download.localhost/files/d/i/content'), method:'GET', headers:new Headers() };
    await method.authorize(req, { config:{} }, {});
    const result = await fetch(req.url, { headers:req.headers });
    assert.equal(result.status, 502);
    assert.equal((await result.json()).error.code, 'invalid_transfer');
  } finally { await adapter.dispose(); }
});

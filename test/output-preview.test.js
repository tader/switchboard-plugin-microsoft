import { test } from 'node:test';
import assert from 'node:assert/strict';
import { graphServices, config, credentials, contentContext } from './fixtures/output-preview.js';
import { createContentPlugin } from '../plugins/microsoft/content/service.js';
import { createAdapter } from '../plugins/teams/index.js';

for (const service of graphServices()) {
  test(`${service.id} declares only the reviewed Graph field for all three profiles`, async () => {
    const calendar = service.id === 'outlook-calendar';
    assert.deepEqual(service.outputPreviews, [{
      origin: 'https://graph.microsoft.com', method: 'GET', path: '/v1.0/me/' + (calendar ? 'calendar' : 'drive'),
      authMethods: ['oauth', 'device', 'authentication-source'],
      fields: { [calendar ? 'defaultOnlineMeetingProvider' : 'driveType']: { maxBytes: 32 } },
    }]);
    for (const methodId of service.outputPreviews[0].authMethods) {
      const conn = { methodId, config, credentials: credentials() };
      const request = { method: 'GET', url: new URL(service.outputPreviews[0].origin + service.outputPreviews[0].path), headers: new Headers() };
      await service.authMethods.find(m => m.id === methodId).authorize(request, conn, {});
      assert.equal(request.headers.get('authorization'), 'Bearer fixture-access');
      assert.equal(request.url.origin, service.outputPreviews[0].origin);
    }
    if (calendar) assert.equal(service.baseUrl({ methodId: 'authentication-source', config: { backend: 'outlook' } }), 'https://outlook.office.com/api/v2.0');
  });
}
for (const kind of ['word', 'powerpoint', 'teams']) {
  test(`${kind} binds capabilities to each allocated port and preserves authorization`, async t => {
    const plugins = [];
    t.after(async () => { for (const plugin of plugins) await plugin.dispose(); });
    for (let start = 0; start < 2; start++) {
      const plugin = kind === 'teams'
        ? await createAdapter({ dataDir: '/unused', settings: {} }, { createSession: profile => ({ profile, adopt() {}, status: () => ({}), close: async () => {} }), createAPI: () => ({}) })
        : await createContentPlugin(contentContext, { id: kind, name: kind });
      plugins.push(plugin);
      const service = plugin.services[0], rule = service.outputPreviews[0];
      assert.deepEqual(rule, { origin: rule.origin, method: 'GET', path: '/capabilities', authMethods: ['authentication-source'], fields: { [kind === 'teams' ? 'transport' : 'kind']: { maxBytes: 32 } } });
      assert.match(rule.origin, /^http:\/\/127\.0\.0\.1:\d+$/);
      const req = { method: 'GET', url: new URL('/capabilities', service.baseUrl), headers: new Headers() };
      await service.authMethods[0].authorize(req, { credentials: { profile: '44444444-4444-4444-8444-444444444444' }, config: { access: 'read' } }, {});
      assert.equal(req.url.origin, rule.origin);
      const response = await fetch(req.url, { headers: req.headers });
      assert.equal(response.status, 200);
      assert.equal((await response.json())[kind === 'teams' ? 'transport' : 'kind'], kind === 'teams' ? 'api' : kind);
      assert.equal((await fetch(req.url, { headers: req.headers })).status, kind === 'teams' ? 401 : 403);
    }
    assert.notEqual(plugins[0].services[0].outputPreviews[0].origin, plugins[1].services[0].outputPreviews[0].origin);
  });
}
test('Org Explorer inherits no content preview', async () => {
  const plugin = await createContentPlugin(contentContext, { id: 'org-explorer', name: 'Org Explorer' });
  try { assert.equal(plugin.services[0].outputPreviews, undefined); } finally { await plugin.dispose(); }
});

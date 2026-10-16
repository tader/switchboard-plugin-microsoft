import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { CDP, Page, findBrowser } from '../lib/cdp.js';
import { teamsActions } from '../lib/actions-dom.js';
import { actionsFixtureHTML, runActionsDOMSuite } from '../test/actions-dom-suite.js';
import { teamsDOM, DEFAULT_SELECTORS } from '../lib/dom.js';
import { fixtureHTML, runDOMSuite } from '../test/dom-suite.js';
import { featuresFixtureHTML, runFeaturesDOMSuite } from '../test/features-dom-suite.js';
import { discoveryFixtureHTML, runDiscoveryDOMSuite } from '../test/discovery-dom-suite.js';
import { quotesFixtureHTML, runQuotesDOMSuite } from '../test/quotes-dom-suite.js';
import { teamsExtended } from '../lib/extended-dom.js';
import { extendedFixtureHTML, runExtendedDOMSuite, runRichEditDOMSuite, runChannelSearchDOMSuite } from '../test/extended-dom-suite.js';
import { createAdapter } from '../legacy-index.js';

// Runs outside the Codex sandbox on the user's Mac. It sends only fixture
// messages to a blank local page, never to Teams or any external recipient.
const directory = await fs.mkdtemp('/tmp/teams-smoke-');
let cdp, adapter;
try {
  cdp = new CDP(await findBrowser(process.env.TEAMS_BROWSER_PATH), directory + '/browser', { headless: true });
  const page = await cdp.page('about:blank');
  assert.ok(page instanceof Page);
  console.log(await runDOMSuite(page, teamsDOM, DEFAULT_SELECTORS, fixtureHTML));
  console.log(await runFeaturesDOMSuite(page, teamsDOM, DEFAULT_SELECTORS, featuresFixtureHTML));
  console.log(await runDiscoveryDOMSuite(page, teamsDOM, DEFAULT_SELECTORS, discoveryFixtureHTML));
  console.log(await runQuotesDOMSuite(page, teamsDOM, DEFAULT_SELECTORS, quotesFixtureHTML));
  console.log(await runActionsDOMSuite(page, teamsActions, DEFAULT_SELECTORS, actionsFixtureHTML));
  console.log(await runExtendedDOMSuite(page, teamsExtended, DEFAULT_SELECTORS, extendedFixtureHTML));
  console.log(await runRichEditDOMSuite(page, teamsExtended, teamsActions, DEFAULT_SELECTORS, actionsFixtureHTML));
  console.log(await runChannelSearchDOMSuite(page, teamsDOM, teamsExtended, DEFAULT_SELECTORS, extendedFixtureHTML));
  adapter = await createAdapter({ dataDir: directory + '/adapter', settings: {} }, {
    createBrowser: () => ({ serial: fn => fn(), login: async () => ({}), dom: async () => ({ ready: true }), close: async () => {} }),
  });
  const service = adapter.services[0], auth = service.authMethods[0];
  const flow = await auth.connect({ config: { label: 'Smoke fixture' } });
  const conn = await auth.poll({ pending: flow.pending });
  const req = { method: 'GET', url: new URL('/status', service.baseUrl), headers: new Headers() };
  auth.authorize(req, conn);
  const response = await fetch(req.url, { headers: req.headers });
  assert.equal(response.status, 200);
  assert.equal((await response.json()).ready, true);
  assert.equal((await fetch(req.url, { headers: req.headers })).status, 401);
  console.log('Private Chrome pipe and real loopback HTTP smoke tests passed.');
} finally {
  await adapter?.dispose();
  await cdp?.close();
  await fs.rm(directory, { recursive: true, force: true });
}

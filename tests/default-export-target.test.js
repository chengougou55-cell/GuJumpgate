const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

test('default export target is SUB2API in shared capabilities', () => {
  const sandbox = {};
  vm.runInNewContext(readProjectFile('shared/flow-capabilities.js'), sandbox);

  const capabilities = sandbox.MultiPageFlowCapabilities;
  assert.equal(capabilities.DEFAULT_PANEL_MODE, 'sub2api');
  assert.equal(capabilities.normalizePanelMode(), 'sub2api');
  assert.equal(capabilities.normalizePanelMode('local-cpa-json'), 'local-cpa-json');
});

test('sidepanel initial export target selects SUB2API', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');

  assert.match(html, /<option value="sub2api" selected>SUB2API<\/option>/);
  assert.doesNotMatch(html, /<option value="local-cpa-json" selected>/);
});

test('default SUB2API url points to configured admin accounts page', () => {
  const backgroundScript = readProjectFile('background.js');

  assert.match(backgroundScript, /const DEFAULT_SUB2API_URL = 'http:\/\/156\.239\.40\.207:18080\/admin\/accounts';/);
  assert.match(backgroundScript, /case 'sub2apiUrl':\n\s+return normalizeSub2ApiUrl\(value\);/);
});

test('sidepanel initial account access strategy selects session JSON import', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');
  const backgroundScript = readProjectFile('background.js');
  const sidepanelScript = readProjectFile('sidepanel/sidepanel.js');

  assert.match(html, /<option value="session_json" selected>SESSION JSON导入<\/option>/);
  assert.doesNotMatch(html, /<option value="oauth" selected>Oauth<\/option>/);
  assert.match(backgroundScript, /const DEFAULT_PLUS_ACCOUNT_ACCESS_STRATEGY = PLUS_ACCOUNT_ACCESS_STRATEGY_SUB2API_CODEX_SESSION;/);
  assert.match(backgroundScript, /plusAccountAccessStrategy: DEFAULT_PLUS_ACCOUNT_ACCESS_STRATEGY,/);
  assert.match(sidepanelScript, /const DEFAULT_PLUS_ACCOUNT_ACCESS_STRATEGY = PLUS_ACCOUNT_ACCESS_STRATEGY_SUB2API_CODEX_SESSION;/);
  assert.match(sidepanelScript, /let currentPlusAccountAccessStrategy = DEFAULT_PLUS_ACCOUNT_ACCESS_STRATEGY;/);
});

test('sidepanel initial checkout conversion uses cloud mode', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');

  assert.match(html, /<input type="checkbox" id="input-plus-checkout-cloud-conversion-enabled" checked \/>/);
});

test('outlookEmail is the default mail service', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');
  const backgroundScript = readProjectFile('background.js');
  const sidepanelScript = readProjectFile('sidepanel/sidepanel.js');

  assert.match(html, /<option value="outlook-email" selected>outlookEmail<\/option>/);
  assert.doesNotMatch(html, /<option value="hotmail-api" selected>/);
  assert.match(backgroundScript, /mailProvider: OUTLOOK_EMAIL_PROVIDER,/);
  assert.match(backgroundScript, /emailGenerator: OUTLOOK_EMAIL_GENERATOR,/);
  assert.match(sidepanelScript, /return OUTLOOK_EMAIL_PROVIDER;\n}/);
});

test('sidepanel exposes one-click log copy and txt export controls', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');
  const script = readProjectFile('sidepanel/sidepanel.js');

  assert.match(html, /id="btn-copy-log"/);
  assert.match(html, /id="btn-export-log-txt"/);
  assert.match(script, /const btnCopyLog = document\.getElementById\('btn-copy-log'\);/);
  assert.match(script, /const btnExportLogTxt = document\.getElementById\('btn-export-log-txt'\);/);
  assert.match(script, /function buildCurrentLogText\(\)/);
  assert.ok(script.includes("downloadTextFile(`${logText}\\n`, buildLogExportFileName(), 'text/plain;charset=utf-8');"));
});

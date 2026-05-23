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

test('sidepanel initial checkout conversion uses cloud mode', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');

  assert.match(html, /<input type="checkbox" id="input-plus-checkout-cloud-conversion-enabled" checked \/>/);
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

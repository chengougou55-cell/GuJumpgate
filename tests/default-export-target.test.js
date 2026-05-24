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

test('skip checkout plus payment mode omits checkout creation step', () => {
  const sandbox = {};
  vm.runInNewContext(readProjectFile('data/step-definitions.js'), sandbox);

  const stepDefinitions = sandbox.MultiPageStepDefinitions;
  const steps = stepDefinitions.getSteps({
    activeFlowId: 'openai',
    plusModeEnabled: true,
    plusPaymentMethod: 'skip-checkout',
    plusAccountAccessStrategy: 'oauth',
  });

  assert.equal(stepDefinitions.normalizePlusPaymentMethod('skip-checkout'), 'skip-checkout');
  assert.deepEqual(
    Array.from(steps, (step) => [step.id, step.key]),
    [
      [1, 'open-chatgpt'],
      [2, 'submit-signup-email'],
      [3, 'fill-password'],
      [4, 'fetch-signup-code'],
      [5, 'fill-profile'],
      [7, 'oauth-login'],
      [8, 'fetch-login-code'],
      [9, 'confirm-oauth'],
      [10, 'platform-verify'],
    ]
  );
  assert.equal(steps.some((step) => step.key === 'plus-checkout-create'), false);
  assert.equal(steps.some((step) => step.id === 6), false);
});

test('skip checkout keeps session import strategy when selected', () => {
  const sandbox = {};
  vm.runInNewContext(readProjectFile('data/step-definitions.js'), sandbox);

  const stepDefinitions = sandbox.MultiPageStepDefinitions;
  const steps = stepDefinitions.getSteps({
    activeFlowId: 'openai',
    plusModeEnabled: true,
    plusPaymentMethod: 'skip-checkout',
    plusAccountAccessStrategy: 'sub2api_codex_session',
  });

  assert.deepEqual(
    Array.from(steps, (step) => [step.id, step.key]),
    [
      [1, 'open-chatgpt'],
      [2, 'submit-signup-email'],
      [3, 'fill-password'],
      [4, 'fetch-signup-code'],
      [5, 'fill-profile'],
      [7, 'sub2api-session-import'],
    ]
  );
  assert.equal(steps.some((step) => step.key === 'plus-checkout-create'), false);
  assert.equal(steps.some((step) => step.id === 6), false);
});

test('sidepanel always passes plus account access strategy to step definitions', () => {
  const script = readProjectFile('sidepanel/sidepanel.js');

  assert.match(script, /plusAccountAccessStrategy: normalizedAccountAccessStrategy,/);
  assert.doesNotMatch(script, /normalizedAccountAccessStrategy\s*&&\s*normalizedAccountAccessStrategy\s*!==\s*defaultAccountAccessStrategy/);
  assert.doesNotMatch(script, /getStepIdByKeyForCurrentMode\('plus-checkout-create'\)\s*\|\|\s*6/);
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

test('PayPal hosted checkout dialog auto retries after 15 seconds', () => {
  const script = readProjectFile('sidepanel/sidepanel.js');
  const routerScript = readProjectFile('background/message-router.js');
  const autoRunScript = readProjectFile('background/auto-run-controller.js');
  const checkoutScript = readProjectFile('background/steps/create-plus-checkout.js');
  const backgroundScript = readProjectFile('background.js');

  assert.match(script, /const PAYPAL_HOSTED_GENERIC_ERROR_AUTO_RETRY_DELAY_MS = 15000;/);
  assert.match(script, /autoSelect:\s*\{\s*actionId:\s*'retry',\s*delayMs:\s*PAYPAL_HOSTED_GENERIC_ERROR_AUTO_RETRY_DELAY_MS,\s*\}/);
  assert.match(script, /\$\{retryDelaySeconds\} 秒内未处理将自动重试/);
  assert.match(script, /function clearActionModalAutoSelectTimer\(\)/);
  assert.match(script, /button\.textContent = `\$\{label\} \(\$\{remainingSeconds\}\)`;/);
  assert.match(script, /resolveModalChoice\(actionId,\s*\{\s*autoSelected:\s*true\s*\}\);/);
  assert.match(script, /buildResult:\s*\(choice,\s*meta\)\s*=>\s*\(\{\s*action:\s*choice,\s*autoSelected:\s*Boolean\(meta\?\.autoSelected\),\s*\}\),/);
  assert.match(script, /autoSelected:\s*Boolean\(choice\?\.autoSelected\),/);
  assert.match(routerScript, /AUTO_RUN_MAX_RETRIES_PER_ROUND = 3,/);
  assert.match(routerScript, /function resumeAutoRunAfterPaypalHostedGenericError\(state = \{\}\)/);
  assert.match(routerScript, /return Boolean\(state\?\.plusManualConfirmationAutoRunContext\);/);
  assert.match(routerScript, /autoSelected && hasAutoRunContextForPaypalRetryResume\(currentState\)/);
  assert.match(routerScript, /autoRetryLimitReached:\s*true/);
  assert.match(routerScript, /autoRunRetryPaypalCallback:\s*true/);
  assert.match(routerScript, /resumeAttemptRun:\s*resumeOptions\.nextAttempt/);
  assert.match(autoRunScript, /function waitForManualPaypalHostedGenericErrorAutoRetry\(requestId = '', delayMs = 15000\)/);
  assert.match(autoRunScript, /waitForHostedCheckoutGenericErrorAutoRetry = blockedByHostedCheckoutGenericError/);
  assert.match(autoRunScript, /waitForManualPaypalHostedGenericErrorAutoRetry\(\s*hostedCheckoutGenericErrorRequestId,\s*15000\s*\)/);
  assert.match(autoRunScript, /autoRunRetryPaypalCallback:\s*true/);
  assert.match(autoRunScript, /plusManualConfirmationAutoRunContext:\s*false/);
  assert.match(autoRunScript, /attemptRun \+= 1;/);
  assert.match(checkoutScript, /plusManualConfirmationAutoRunContext:\s*hasAutoRunContext/);
  assert.match(checkoutScript, /plusManualConfirmationResolvedAutoSelected:\s*false/);
  assert.match(backgroundScript, /AUTO_RUN_MAX_RETRIES_PER_ROUND,\n\s+AUTO_RUN_TIMER_KIND_SCHEDULED_START,/);
});

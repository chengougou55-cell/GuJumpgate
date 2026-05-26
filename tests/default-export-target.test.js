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

test('xiaohongshu auto mode starts at checkout and keeps session import tail', () => {
  const sidepanelHtml = readProjectFile('sidepanel/sidepanel.html');
  const sidepanelCss = readProjectFile('sidepanel/sidepanel.css');
  const sidepanelScript = readProjectFile('sidepanel/sidepanel.js');
  const routerScript = readProjectFile('background/message-router.js');
  const checkoutScript = readProjectFile('background/steps/create-plus-checkout.js');
  const sub2ApiImportScript = readProjectFile('background/steps/sub2api-session-import.js');
  const cpaImportScript = readProjectFile('background/steps/cpa-session-import.js');
  const backgroundScript = readProjectFile('background.js');
  const sub2ApiScript = readProjectFile('background/sub2api-api.js');
  const runtimeStateScript = readProjectFile('background/runtime-state.js');
  const autoRunScript = readProjectFile('background/auto-run-controller.js');

  assert.match(sidepanelScript, /const XIAOHONGSHU_SUB2API_GROUP_NAME = 'xiaohongshu';/);
  assert.match(sidepanelScript, /xiaohongshu 是小红书模式专用分组，普通模式不能添加/);
  assert.match(sidepanelScript, /function openAutoRunModeDialog\(\)/);
  assert.match(sidepanelScript, /小红书模式会跳过前置注册\/登录/);
  assert.match(sidepanelHtml, /id="btn-auto-start-extra"/);
  assert.match(sidepanelCss, /\.modal-actions \{[\s\S]*flex-wrap: wrap;/);
  assert.match(sidepanelScript, /btnAutoStartExtra/);
  assert.match(sidepanelScript, /btnAutoStartCancel,\s*btnAutoStartRestart,\s*btnAutoStartExtra,\s*btnAutoStartContinue/);
  assert.match(sidepanelScript, /id: 'normal', label: '普通模式'/);
  assert.match(sidepanelScript, /type:\s*'AUTO_RUN_XIAOHONGSHU'/);
  assert.match(sidepanelScript, /function openXiaohongshuAccessTokenDialog\(\)/);
  assert.match(sidepanelScript, /extractAccessTokenFromInput\(result\.accessToken\)/);
  assert.match(routerScript, /case 'AUTO_RUN_XIAOHONGSHU':/);
  assert.match(routerScript, /const startNodeId = 'plus-checkout-create';/);
  assert.match(routerScript, /buildXiaohongshuNodeStatuses\(xiaohongshuState, startNodeId\)/);
  assert.match(routerScript, /sub2apiGroupName:\s*'xiaohongshu'/);
  assert.match(routerScript, /xiaohongshuModeEnabled:\s*true/);
  assert.match(routerScript, /function buildXiaohongshuRuntimeReset\(state = \{\}\)/);
  assert.match(routerScript, /function hasXiaohongshuRuntimeResidue\(state = \{\}\)/);
  assert.match(routerScript, /normalizePublicSub2ApiGroupNames\(state\?\.sub2apiGroupNames\)/);
  assert.match(routerScript, /case 'AUTO_RUN':[\s\S]*const preflightState = await getState\(\);[\s\S]*isAutoRunPausedSnapshot\(preflightState\)[\s\S]*getPendingAutoRunTimerPlan\(preflightState\)[\s\S]*clearStopRequest\(\);[\s\S]*\.\.\.buildXiaohongshuRuntimeReset\(state\),/);
  assert.match(routerScript, /\.\.\.buildXiaohongshuRuntimeReset\(state\),\n\s+autoRunSkipFailures,/);
  assert.match(routerScript, /case 'AUTO_RUN_XIAOHONGSHU':[\s\S]*const accessToken = resolveXiaohongshuAccessToken\(message\.payload \|\| \{\}\);[\s\S]*isAutoRunPausedSnapshot\(currentState\)[\s\S]*clearStopRequest\(\);[\s\S]*sub2apiGroupName:\s*'xiaohongshu'/);
  assert.match(routerScript, /case 'SCHEDULE_AUTO_RUN':[\s\S]*const preflightState = await getState\(\);[\s\S]*isAutoRunPausedSnapshot\(preflightState\)[\s\S]*const result = await scheduleAutoRun[\s\S]*clearStopRequest\(\);[\s\S]*\.\.\.buildXiaohongshuRuntimeReset\(state\),/);
  assert.match(routerScript, /case 'SCHEDULE_AUTO_RUN':[\s\S]*const result = await scheduleAutoRun[\s\S]*\.\.\.buildXiaohongshuRuntimeReset\(state\),[\s\S]*return result;/);
  assert.match(routerScript, /sub2apiSessionId:\s*null/);
  assert.match(routerScript, /sub2apiGroupIds:\s*\[\]/);
  assert.match(routerScript, /startAutoRunLoop\(1,\s*\{\s*autoRunSkipFailures,/);
  assert.match(routerScript, /xiaohongshuModeEnabled:\s*true,\n\s+mode:\s*'continue'/);
  assert.match(routerScript, /mode:\s*'continue'/);
  assert.match(checkoutScript, /function getDirectCheckoutAccessToken\(state = \{\}\)/);
  assert.match(checkoutScript, /if \(!state\?\.xiaohongshuModeEnabled\) \{\n\s+return '';/);
  assert.match(checkoutScript, /小红书模式正在请求云端服务生成订阅长链/);
  assert.match(sub2ApiImportScript, /function resolveDirectSessionAccessToken\(state = \{\}\)/);
  assert.match(sub2ApiImportScript, /if \(!state\?\.xiaohongshuModeEnabled\) \{\n\s+return '';/);
  assert.match(sub2ApiImportScript, /小红书模式已接收 accessToken，正在直接导入 SUB2API/);
  assert.match(cpaImportScript, /function resolveDirectSessionAccessToken\(state = \{\}\)/);
  assert.match(cpaImportScript, /if \(!state\?\.xiaohongshuModeEnabled\) \{\n\s+return '';/);
  assert.match(cpaImportScript, /小红书模式已接收 accessToken，正在直接导入 CPA/);
  assert.match(backgroundScript, /redactSensitiveStateForLog\(updates\)/);
  assert.match(backgroundScript, /const XIAOHONGSHU_SUB2API_GROUP_NAME = 'xiaohongshu';/);
  assert.match(backgroundScript, /function normalizePublicSub2ApiGroupNames\(value = ''\)/);
  assert.match(backgroundScript, /case 'sub2apiGroupName':\n\s+return isXiaohongshuSub2ApiGroupName\(value\) \? DEFAULT_SUB2API_GROUP_NAME : String\(value \|\| ''\)\.trim\(\);/);
  assert.match(sub2ApiScript, /function resolveSub2ApiGroupNamesForState\(state = \{\}\)/);
  assert.match(sub2ApiScript, /return Boolean\(state\?\.xiaohongshuModeEnabled\);/);
  assert.match(sub2ApiScript, /return XIAOHONGSHU_SUB2API_GROUP_NAME;/);
  assert.match(sub2ApiScript, /const storedGroupIds = isXiaohongshuModeState\(state\)\n\s+\? \[\]/);
  assert.match(runtimeStateScript, /'xiaohongshuModeEnabled'/);
  assert.match(runtimeStateScript, /'sub2apiGroupName'/);
  assert.match(runtimeStateScript, /'sub2apiGroupNames'/);
  assert.match(autoRunScript, /const preserveXiaohongshuRuntime = Boolean\(options\.xiaohongshuModeEnabled\);/);
  assert.match(autoRunScript, /const resetStaleXiaohongshuRuntime = !preserveXiaohongshuRuntime && Boolean\(preflightState\?\.xiaohongshuModeEnabled\);/);
  assert.match(autoRunScript, /const effectiveInitialMode = resetStaleXiaohongshuRuntime \? 'restart' : initialMode;/);
  assert.match(autoRunScript, /if \(xiaohongshuModeEnabled\) \{\n\s+return false;\n\s+\}/);
  assert.match(autoRunScript, /const shouldResetXiaohongshuRuntime = Boolean\(options\?\.xiaohongshuModeEnabled \|\| latestState\?\.xiaohongshuModeEnabled\);/);
  assert.match(autoRunScript, /\.\.\.\(preserveXiaohongshuRuntime \? \{/);
  assert.match(autoRunScript, /function buildXiaohongshuRuntimeReset\(state = \{\}\)/);
  assert.match(autoRunScript, /function hasXiaohongshuRuntimeResidue\(state = \{\}\)/);
  assert.match(autoRunScript, /\.\.\.\(shouldClearCredentialAliases \? \{\n\s+chatgptAccessToken: '',\n\s+accessToken: '',\n\s+\} : \{\}\),/);
  assert.match(autoRunScript, /\.\.\.\(preserveXiaohongshuRuntime \? buildXiaohongshuRuntimeReset\(await getState\(\)\) : \{\}\),/);
});

test('xiaohongshu runtime reset preserves ordinary access token aliases', () => {
  function isXiaohongshuSub2ApiGroupName(value = '') {
    return String(value || '').trim().toLowerCase() === 'xiaohongshu';
  }

  function normalizePublicSub2ApiGroupNames(value = '') {
    const source = Array.isArray(value)
      ? value
      : String(value || '').split(/[\r\n,，、;；]+/);
    const names = [];
    const seen = new Set();
    for (const item of source) {
      const name = String(item || '').trim();
      const key = name.toLowerCase();
      if (!key || key === 'xiaohongshu' || seen.has(key)) {
        continue;
      }
      seen.add(key);
      names.push(name);
    }
    return names;
  }

  const sandbox = { isXiaohongshuSub2ApiGroupName, normalizePublicSub2ApiGroupNames };
  const script = readProjectFile('background/auto-run-controller.js');
  const helperMatch = script.match(/function hasXiaohongshuRuntimeResidue[\s\S]*?\n    }\n\n    function buildXiaohongshuRuntimeReset[\s\S]*?\n    }\n\n    async function autoRunLoop/);
  assert.ok(helperMatch, 'expected xiaohongshu reset helpers');
  vm.runInNewContext(helperMatch[0].replace(/\n    async function autoRunLoop[\s\S]*$/, ''), sandbox);

  assert.deepEqual(JSON.parse(JSON.stringify(sandbox.buildXiaohongshuRuntimeReset({
    sub2apiGroupName: 'codex',
    sub2apiGroupNames: ['codex', 'openai-plus'],
    accessToken: 'ordinary-token',
    chatgptAccessToken: 'ordinary-chatgpt-token',
  }))), {
    sub2apiGroupName: 'codex',
    sub2apiGroupNames: ['codex', 'openai-plus'],
    xiaohongshuModeEnabled: false,
    xiaohongshuAccessToken: '',
    directCheckoutAccessToken: '',
    manualCheckoutAccessToken: '',
  });

  assert.deepEqual(JSON.parse(JSON.stringify(sandbox.buildXiaohongshuRuntimeReset({
    sub2apiGroupName: 'xiaohongshu',
    sub2apiGroupNames: ['codex', 'xiaohongshu'],
    xiaohongshuModeEnabled: true,
    accessToken: 'xiaohongshu-token',
    chatgptAccessToken: 'xiaohongshu-token',
    directCheckoutAccessToken: 'xiaohongshu-token',
  }))), {
    sub2apiGroupName: 'codex',
    sub2apiGroupNames: ['codex'],
    xiaohongshuModeEnabled: false,
    xiaohongshuAccessToken: '',
    directCheckoutAccessToken: '',
    manualCheckoutAccessToken: '',
    chatgptAccessToken: '',
    accessToken: '',
  });
});

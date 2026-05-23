const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

function extractAutoRunBranch(script, branchName, nextBranchPattern) {
  const pattern = new RegExp(
    `if \\(${branchName}\\) \\{[\\s\\S]*?\\n            \\}\\n\\n            ${nextBranchPattern}`
  );
  return script.match(pattern)?.[0];
}

function assertContinuesNextRound(block, label) {
  assert.ok(block, `${label} block should be present`);
  assert.match(block, /自动流程将继续下一轮。/);
  assert.match(block, /forceFreshTabsNextRun = true;\n\s+break;/);
  assert.doesNotMatch(block, /stoppedEarly = true/);
  assert.doesNotMatch(block, /broadcastAutoRunStatus\('stopped'/);
}

test('Plus non-free trial failures continue the next auto-run round', () => {
  const createCheckoutScript = readProjectFile('background/steps/create-plus-checkout.js');
  const fillCheckoutScript = readProjectFile('background/steps/fill-plus-checkout.js');
  const autoRunScript = readProjectFile('background/auto-run-controller.js');

  assert.match(createCheckoutScript, /function shouldLetAutoRunHandleNonFreeTrial\(state = \{\}\) \{/);
  assert.match(createCheckoutScript, /return Boolean\(state\?\.autoRunning\);/);
  assert.match(createCheckoutScript, /shouldRetryNonFreeTrial \|\| letAutoRunHandleNonFreeTrial/);
  assert.match(createCheckoutScript, /failNodeFromBackground\('plus-checkout-create', `PLUS_CHECKOUT_NON_FREE_TRIAL::\$\{stopReason\}`\);/);

  assert.match(fillCheckoutScript, /const letAutoRunHandleNonFreeTrial = Boolean\(state\?\.autoRunning\);/);
  assert.match(fillCheckoutScript, /if \(letAutoRunHandleNonFreeTrial\) \{\n\s+throw new Error\(`PLUS_CHECKOUT_NON_FREE_TRIAL::\$\{stopReason\}`\);/);

  const plusBlock = extractAutoRunBranch(
    autoRunScript,
    'blockedByPlusNonFreeTrial',
    'if \\(blockedByGpcTaskEnded\\)'
  );
  assertContinuesNextRound(plusBlock, 'Plus non-free-trial');
});

test('recoverable checkout terminal failures do not stop multi-round auto-run', () => {
  const autoRunScript = readProjectFile('background/auto-run-controller.js');

  const paypalResendBlock = extractAutoRunBranch(
    autoRunScript,
    'blockedByHostedCheckoutVerificationResendLimit',
    'if \\(blockedByCloudCheckoutAlreadyPaid\\)'
  );
  assertContinuesNextRound(paypalResendBlock, 'PayPal resend-limit');

  const cloudAlreadyPaidBlock = extractAutoRunBranch(
    autoRunScript,
    'blockedByCloudCheckoutAlreadyPaid',
    'if \\(blockedBySignupUserAlreadyExists\\)'
  );
  assertContinuesNextRound(cloudAlreadyPaidBlock, 'Cloud already-paid');

  const genericErrorBlock = extractAutoRunBranch(
    autoRunScript,
    'blockedByHostedCheckoutGenericError',
    'if \\(blockedByHostedCheckoutVerificationResendLimit\\)'
  );
  assert.ok(genericErrorBlock, 'Hosted checkout genericError block should be present');
  assert.match(genericErrorBlock, /if \(!autoRunRetryPaypalCallback\) \{/);
  assert.match(genericErrorBlock, /自动流程将继续下一轮。/);
  assert.match(genericErrorBlock, /forceFreshTabsNextRun = true;\n\s+break;/);
});

test('known account-level blockers continue the next auto-run round', () => {
  const autoRunScript = readProjectFile('background/auto-run-controller.js');
  const branchCases = [
    ['blockedByAddPhone', 'if \\(blockedByPhoneNoSupply\\)', 'add-phone'],
    ['blockedByPhoneNoSupply', 'if \\(blockedByPlusNonFreeTrial\\)', 'phone no supply'],
    ['blockedByGpcTaskEnded', 'if \\(blockedByHostedCheckoutGenericError\\)', 'GPC task ended'],
    ['blockedBySignupUserAlreadyExists', 'if \\(blockedByStep4Route405\\)', 'signup user already exists'],
    ['blockedByStep4Route405', 'if \\(canRetry\\)', 'step 4 route 405'],
  ];

  for (const [branchName, nextBranchPattern, label] of branchCases) {
    const block = extractAutoRunBranch(autoRunScript, branchName, nextBranchPattern);
    assertContinuesNextRound(block, label);
    assert.doesNotMatch(block, /自动重试未开启，当前自动运行将停止/);
  }
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

function loadPasswordPolicy(relativePath, endMarker, generatorName) {
  const source = readProjectFile(relativePath);
  const start = source.indexOf('const HOSTED_PAYPAL_PASSWORD_MIN_LENGTH');
  const end = source.indexOf(endMarker, start);
  assert.notEqual(start, -1, `${relativePath} should define PayPal password policy constants`);
  assert.notEqual(end, -1, `${relativePath} should expose password policy before ${endMarker}`);
  const context = {};
  vm.runInNewContext(`${source.slice(start, end)}
globalThis.__policy = {
  buildPassword: ${generatorName},
  isHostedPayPalPasswordCompliant,
  hasHostedPayPalPasswordConsecutiveKeys,
};`, context);
  return context.__policy;
}

function assertPasswordMeetsPayPalPolicy(password, policy) {
  assert.equal(typeof password, 'string');
  assert.ok(password.length >= 8, `password should be at least 8 chars: ${password}`);
  assert.ok(password.length <= 20, `password should be at most 20 chars: ${password}`);
  assert.match(password, /[0-9!@#$%^]/, `password should contain a digit or allowed symbol: ${password}`);
  assert.equal(policy.hasHostedPayPalPasswordConsecutiveKeys(password), false, `password should avoid 4 consecutive keys: ${password}`);
  assert.equal(policy.isHostedPayPalPasswordCompliant(password), true, `password should pass policy helper: ${password}`);
}

test('PayPal hosted checkout passwords satisfy PayPal account policy', () => {
  const contentSource = readProjectFile('content/paypal-flow.js');
  assert.match(contentSource, /const payloadPassword = String\(payload\.password \|\| ''\)/);
  assert.match(contentSource, /isHostedPayPalPasswordCompliant\(payloadPassword\)/);

  const policies = [
    loadPasswordPolicy('content/paypal-flow.js', 'function buildHostedVisaCard', 'buildHostedRandomPassword'),
    loadPasswordPolicy('background/steps/create-plus-checkout.js', 'function buildHostedCheckoutVisaCard', 'buildHostedCheckoutRandomPassword'),
  ];

  for (const policy of policies) {
    for (let index = 0; index < 200; index += 1) {
      assertPasswordMeetsPayPalPolicy(policy.buildPassword(), policy);
    }
    assert.equal(policy.isHostedPayPalPasswordCompliant('abcd9!Lm'), false);
    assert.equal(policy.isHostedPayPalPasswordCompliant('qwer9!Lm'), false);
    assert.equal(policy.isHostedPayPalPasswordCompliant('AAAA9!Lm'), false);
    assert.equal(policy.isHostedPayPalPasswordCompliant('NoDigitSymbol'), false);
    assert.equal(policy.isHostedPayPalPasswordCompliant('Pa9!Vx2@Lm7#Qr'), true);
  }
});

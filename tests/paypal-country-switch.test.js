const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

function readHar(relativePath = 'www.paypal.com.har') {
  return JSON.parse(readProjectFile(relativePath));
}

test('PayPal login flow exposes US country switch command', () => {
  const script = readProjectFile('content/paypal-flow.js');

  assert.match(script, /PAYPAL_HOSTED_COUNTRY_TARGET_CODE = 'US'/);
  assert.match(script, /message\.type === 'PAYPAL_ENSURE_US_COUNTRY'/);
  assert.match(script, /case 'PAYPAL_ENSURE_US_COUNTRY':\n\s+return ensureHostedCountryUnitedStates/);
  assert.match(script, /el\.matches\?\.\(':disabled'\)/);
  assert.match(script, /fieldset\[disabled\]/);
  assert.match(script, /function buildHostedCountryUrl/);
  assert.match(script, /url\.searchParams\.set\('country\.x', target\)/);
  assert.match(script, /method: 'country_url'/);
  assert.match(script, /function getHostedCountryServerDefaultCode/);
  assert.ok(script.includes('/"defaultValue"\\s*:\\s*"([A-Z0-9]{2})"/'));
  assert.ok(script.includes('/ccpg=([A-Z]{2})(?:\\\\u0026|&)/'));
  assert.match(script, /const forceUiFallback = Boolean\(options\.forceUiFallback\)/);
  assert.match(script, /if \(!forceUiFallback && shouldNavigateHostedCountryUrl/);
  assert.match(script, /setTimeout\(\(\) => \{\n\s+window\.location\.assign\(targetUrl\);/);
  assert.match(script, /function getHostedCountrySelectTrigger\(\)/);
  assert.match(script, /function findHostedCountryUnitedStatesOption\(\)/);
  assert.ok(script.includes('if (!/^\\/pay\\/?$/i.test(getPayPalHostedPathname())) {'));
  assert.match(script, /scheduleHostedCountrySwitchAutoRun\(\);/);
});

test('PayPal approval step switches hosted login country before submitting credentials', () => {
  const script = readProjectFile('background/steps/paypal-approve.js');

  assert.match(script, /async function ensurePayPalUsCountry\(tabId, options = \{\}\)/);
  assert.match(script, /type: 'PAYPAL_ENSURE_US_COUNTRY'/);
  assert.match(script, /countryCode: 'US'/);
  assert.match(script, /PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS = 15000/);
  assert.match(script, /async function waitForPayPalUrlUntilStopped\(tabId, options = \{\}\)/);
  assert.match(script, /async function waitForPayPalUsCountryNavigation\(tabId\)/);
  assert.match(script, /async function ensurePayPalUsCountryWithUiFallback\(tabId, reason = ''\)/);
  assert.match(script, /forceUiFallback: Boolean\(options\.forceUiFallback\)/);
  assert.match(script, /ensurePayPalUsCountry\(tabId, \{ forceUiFallback: true \}\)/);
  assert.match(script, /countryResult\?\.navigationStarted/);
  assert.match(script, /country\\.x=US/);
  assert.match(script, /timeoutMs: PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS/);
  assert.match(script, /PayPal 已跳转离开授权页，无法确认国家\/地区是否已切换为美国/);
  assert.match(script, /attemptedLoginCountryFallback/);
  assert.match(script, /pageState\.hostedCountryUnitedStates === false/);
  assert.match(script, /未能切换为美国，已停止登录以避免误用地区/);

  const switchIndex = script.indexOf('const countryResult = await ensurePayPalUsCountry(tabId)');
  const submitIndex = script.indexOf('const submitResult = await submitLogin(tabId, state)');
  assert.ok(switchIndex > -1, 'country switch should be invoked in login branch');
  assert.ok(submitIndex > -1, 'login submission should be present');
  assert.ok(switchIndex < submitIndex, 'country switch should run before login submission');
});

test('PayPal HAR shows country.x=US drives the hosted country state', () => {
  const har = readHar();
  const entries = har.log.entries || [];
  const payDocument = entries.find((entry) => {
    const url = entry.request?.url || '';
    return entry.request?.method === 'GET'
      && entry.response?.status === 200
      && /^https:\/\/www\.paypal\.com\/pay\/?\?/i.test(url)
      && /[?&]country\.x=US(?:&|$)/i.test(url);
  });
  assert.ok(payDocument, 'HAR should contain a PayPal pay document loaded with country.x=US');

  const html = payDocument.response?.content?.text || '';
  assert.match(html, /\\"defaultValue\\"\s*:\s*\\"US\\"/);
  assert.match(html, /ccpg=US(?:\\u0026|&)/);

  const countriesApi = entries.find((entry) => {
    const url = entry.request?.url || '';
    return /\/pay\/api\/countries\?/i.test(url)
      && /[?&]country\.x=US(?:&|$)/i.test(url)
      && entry.response?.status === 200;
  });
  assert.ok(countriesApi, 'HAR should contain the countries API requested with country.x=US');
  const countries = JSON.parse(countriesApi.response?.content?.text || '[]');
  assert.ok(countries.some((country) => country?.key === 'US' && country?.labelText === '美国'));
});

test('PayPal Hermes review click is button driven and covered by HAR02 labels', () => {
  const script = readProjectFile('content/paypal-flow.js');

  assert.match(script, /function findHostedReviewConsentButton\(\)/);
  assert.match(script, /button\[data-testid="consentButton"\], button\[name="agreeAndContinueQL"\]/);
  assert.match(script, /agree\\s\*\(\?:&\|and\)\?\\s\*continue\|agreeAndContinueQL\|agreeContinue/);
  assert.match(script, /同意并继续\|同意\.\*继续\|接受并继续/);
  assert.match(script, /isEnabledControl\(el\) && patterns\.some/);
  assert.doesNotMatch(script, /includes\('Set up once\. Pay faster next time'\)/);
  assert.match(script, /dispatchHostedGenericClick\(button\)/);
  assert.match(script, /clickedButtonText: buttonText/);
  assert.match(script, /stillReviewPage: isPayPalHostedReviewPage\(\)/);
  assert.match(script, /data-multipage-paypal-hosted-hermes-autorun/);
  assert.match(script, /document\.documentElement\.getAttribute\(PAYPAL_HOSTED_HERMES_AUTORUN_SENTINEL\) === '1'/);

  const backgroundScript = readProjectFile('background/steps/create-plus-checkout.js');
  assert.match(backgroundScript, /const hermesResult = await runHostedCheckoutPayPalStep/);
  assert.match(backgroundScript, /hermesResult\?\.submitted && hermesResult\?\.stillReviewPage/);
  assert.match(backgroundScript, /仍在复核页，继续等待页面跳转/);

  const har = readHar('www.paypal.com02.har');
  const entries = har.log.entries || [];
  const hermesDocument = entries.find((entry) => {
    const url = entry.request?.url || '';
    return /\/webapps\/hermes\?/i.test(url)
      && /[?&]country\.x=US(?:&|$)/i.test(url)
      && entry.response?.status === 200;
  });
  assert.ok(hermesDocument, 'HAR02 should contain the Hermes review document');
  const html = hermesDocument.response?.content?.text || '';
  assert.match(html, /"agreeAndContinueQL":"Agree and Continue"/);
  assert.match(html, /"headlineForQL":"Set up once\. Pay faster next time\."/);

  const hagridScript = entries.find((entry) => {
    const url = entry.request?.url || '';
    return /\/checkoutweb\/release\/hagrid\/static\/js\/main\..*\.chunk\.js/i.test(url);
  });
  assert.ok(hagridScript, 'HAR02 should contain the Hagrid main script');
  const hagridText = hagridScript.response?.content?.text || '';
  assert.match(hagridText, /id:"consentButton"/);
  assert.match(hagridText, /"data-testid":"consentButton"/);

  const translation = entries.find((entry) => /translateHtml/i.test(entry.request?.url || ''));
  assert.ok(translation, 'HAR02 should contain translated review labels');
  assert.match(translation.response?.content?.text || '', /同意并继续/);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

require('../background/steps/create-plus-checkout.js');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

test('cloud checkout conversion settings are exposed in sidepanel and persisted', () => {
  const html = readProjectFile('sidepanel/sidepanel.html');
  const sidepanelScript = readProjectFile('sidepanel/sidepanel.js');
  const backgroundScript = readProjectFile('background.js');

  assert.match(html, /id="row-plus-checkout-cloud-conversion-params"/);
  assert.match(html, /id="select-plus-checkout-cloud-conversion-payment-method"/);
  assert.match(html, /id="input-plus-checkout-cloud-conversion-country"/);
  assert.match(html, /id="input-plus-checkout-cloud-conversion-currency"/);

  assert.match(sidepanelScript, /plusCheckoutCloudConversionPaymentMethod:/);
  assert.match(sidepanelScript, /plusCheckoutCloudConversionCountry:/);
  assert.match(sidepanelScript, /plusCheckoutCloudConversionCurrency:/);
  assert.match(sidepanelScript, /normalizePlusCheckoutCloudConversionPaymentMethodValue/);
  assert.match(sidepanelScript, /normalizePlusCheckoutCloudConversionCountryValue/);
  assert.match(sidepanelScript, /normalizePlusCheckoutCloudConversionCurrencyValue/);

  assert.match(backgroundScript, /plusCheckoutCloudConversionPaymentMethod: PLUS_PAYMENT_METHOD_PAYPAL/);
  assert.match(backgroundScript, /plusCheckoutCloudConversionCountry: 'US'/);
  assert.match(backgroundScript, /plusCheckoutCloudConversionCurrency: 'USD'/);
  assert.match(backgroundScript, /case 'plusCheckoutCloudConversionPaymentMethod':/);
  assert.match(backgroundScript, /case 'plusCheckoutCloudConversionCountry':/);
  assert.match(backgroundScript, /case 'plusCheckoutCloudConversionCurrency':/);
});

test('cloud checkout API request uses configured payment method, country, and currency', async () => {
  const logs = [];
  const requests = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    addLog: async (message, level) => logs.push({ message, level }),
    fetch: async (url, options = {}) => {
      requests.push({
        url: String(url || ''),
        headers: options.headers,
        body: JSON.parse(options.body),
      });
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({
          preferredCheckoutUrl: 'https://pay.openai.com/c/pay/test',
          country: 'SG',
          currency: 'SGD',
        }),
      };
    },
  });

  const result = await executor.__test.generateCloudCheckoutFromApi(
    'access_token_1234567890',
    'paypal',
    {
      plusCheckoutCloudConversionApiUrl: 'https://cloud.example.test/api/checkout',
      plusCheckoutCloudConversionApiKey: 'api_key_1234567890',
      plusCheckoutCloudConversionPaymentMethod: 'gopay',
      plusCheckoutCloudConversionCountry: 'sg',
      plusCheckoutCloudConversionCurrency: 'sgd',
    }
  );

  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, 'https://cloud.example.test/api/checkout');
  assert.equal(requests[0].headers['X-API-Key'], 'api_key_1234567890');
  assert.deepEqual(requests[0].body, {
    accessToken: 'access_token_1234567890',
    paymentMethod: 'gopay',
    country: 'SG',
    currency: 'SGD',
  });
  assert.equal(result.preferredCheckoutUrl, 'https://pay.openai.com/c/pay/test');
  assert.equal(result.country, 'SG');
  assert.equal(result.currency, 'SGD');

  const logText = logs.map((entry) => entry.message).join('\n');
  assert.match(logText, /云端支付转换请求报文/);
  assert.match(logText, /云端支付转换响应报文/);
  assert.match(logText, /"paymentMethod":"gopay"/);
  assert.match(logText, /"country":"SG"/);
  assert.match(logText, /"currency":"SGD"/);
  assert.doesNotMatch(logText, /access_token_1234567890/);
  assert.doesNotMatch(logText, /api_key_1234567890/);
});

test('cloud checkout request details fall back when country or currency is incomplete', () => {
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
  });

  const details = executor.__test.getCloudCheckoutRequestDetails('paypal', {
    plusCheckoutCloudConversionPaymentMethod: 'gopay',
    plusCheckoutCloudConversionCountry: 'u',
    plusCheckoutCloudConversionCurrency: 'us',
  });

  assert.deepEqual(details, {
    paymentMethod: 'gopay',
    country: 'ID',
    currency: 'IDR',
  });
});

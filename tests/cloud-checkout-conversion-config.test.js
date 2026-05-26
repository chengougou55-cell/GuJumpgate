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

  assert.match(html, /<option value="">跟随Plus支付<\/option>/);
  assert.match(backgroundScript, /plusCheckoutCloudConversionPaymentMethod: ''/);
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
      plusCheckoutCloudConversionPaymentMethod: 'paypal',
      plusCheckoutCloudConversionCountry: 'sg',
      plusCheckoutCloudConversionCurrency: 'sgd',
    }
  );

  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, 'https://cloud.example.test/api/checkout');
  assert.equal(requests[0].headers['X-API-Key'], 'api_key_1234567890');
  assert.deepEqual(requests[0].body, {
    accessToken: 'access_token_1234567890',
    paymentMethod: 'paypal',
    country: 'SG',
    currency: 'SGD',
  });
  assert.equal(result.preferredCheckoutUrl, 'https://pay.openai.com/c/pay/test');
  assert.equal(result.country, 'SG');
  assert.equal(result.currency, 'SGD');

  const logText = logs.map((entry) => entry.message).join('\n');
  assert.match(logText, /云端支付转换请求报文/);
  assert.match(logText, /云端支付转换响应报文/);
  assert.match(logText, /"paymentMethod":"paypal"/);
  assert.match(logText, /"country":"SG"/);
  assert.match(logText, /"currency":"SGD"/);
  assert.doesNotMatch(logText, /access_token_1234567890/);
  assert.doesNotMatch(logText, /api_key_1234567890/);
});

test('cloud checkout API request follows current plus payment method by default', async () => {
  const requests = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    addLog: async () => {},
    fetch: async (_url, options = {}) => {
      requests.push(JSON.parse(options.body));
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        json: async () => ({
          preferredCheckoutUrl: 'https://pay.openai.com/c/pay/test',
        }),
      };
    },
  });

  await executor.__test.generateCloudCheckoutFromApi(
    'access_token_1234567890',
    'gopay',
    {
      plusCheckoutCloudConversionApiUrl: 'https://cloud.example.test/api/checkout',
      plusCheckoutCloudConversionApiKey: 'api_key_1234567890',
      plusCheckoutCloudConversionPaymentMethod: '',
    }
  );

  assert.equal(requests.length, 1);
  assert.equal(requests[0].paymentMethod, 'gopay');
  assert.equal(requests[0].country, 'ID');
  assert.equal(requests[0].currency, 'IDR');
});

test('cloud checkout API rejects payment method that conflicts with the current plus flow', async () => {
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    addLog: async () => {},
    fetch: async () => {
      throw new Error('fetch should not run');
    },
  });

  await assert.rejects(
    () => executor.__test.generateCloudCheckoutFromApi(
      'access_token_1234567890',
      'paypal',
      {
        plusCheckoutCloudConversionApiUrl: 'https://cloud.example.test/api/checkout',
        plusCheckoutCloudConversionPaymentMethod: 'gopay',
      }
    ),
    /paymentMethod=gopay 与当前 Plus 支付=paypal 不一致/
  );
});

test('cloud checkout response logs redact secrets echoed inside string fields', async () => {
  const logs = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    addLog: async (message, level) => logs.push({ message, level }),
    fetch: async () => ({
      ok: false,
      status: 400,
      statusText: 'Bad Request',
      json: async () => ({
        error: 'bad token access_token_1234567890 and api key api_key_1234567890',
      }),
    }),
  });

  let thrownMessage = '';
  try {
    await executor.__test.generateCloudCheckoutFromApi(
      'access_token_1234567890',
      'paypal',
      {
        plusCheckoutCloudConversionApiUrl: 'https://cloud.example.test/api/checkout',
        plusCheckoutCloudConversionApiKey: 'api_key_1234567890',
      }
    );
  } catch (error) {
    thrownMessage = error?.message || String(error || '');
  }
  assert.match(thrownMessage, /云端支付转换失败/);
  assert.doesNotMatch(thrownMessage, /access_token_1234567890/);
  assert.doesNotMatch(thrownMessage, /api_key_1234567890/);

  const logText = logs.map((entry) => entry.message).join('\n');
  assert.match(logText, /bad token access/);
  assert.doesNotMatch(logText, /access_token_1234567890/);
  assert.doesNotMatch(logText, /api_key_1234567890/);
});

test('cloud checkout request logs redact sensitive URL query parameters', async () => {
  const logs = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    addLog: async (message, level) => logs.push({ message, level }),
    fetch: async () => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      json: async () => ({
        preferredCheckoutUrl: 'https://pay.openai.com/c/pay/test',
      }),
    }),
  });

  await executor.__test.generateCloudCheckoutFromApi(
    'access_token_1234567890',
    'paypal',
    {
      plusCheckoutCloudConversionApiUrl: 'https://cloud.example.test/api/checkout?api_key=url_api_key_1234567890&x=1',
      plusCheckoutCloudConversionApiKey: 'api_key_1234567890',
    }
  );

  const logText = logs.map((entry) => entry.message).join('\n');
  assert.match(logText, /api_key=/);
  assert.doesNotMatch(logText, /url_api_key_1234567890/);
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

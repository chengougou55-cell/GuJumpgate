const assert = require('node:assert/strict');
const test = require('node:test');

require('../background/steps/create-plus-checkout.js');

function createExecutorWithPayload(payload) {
  return globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    fetch: async () => ({
      text: async () => (typeof payload === 'string' ? payload : JSON.stringify(payload)),
    }),
  });
}

async function fetchManualCode(payload) {
  const executor = createExecutorWithPayload(payload);
  const result = await executor.fetchHostedCheckoutVerificationCodeManually({
    verificationUrl: 'http://example.test/api/get_sms?key=test',
  });
  return result.code;
}

test('manual hosted checkout code fetch extracts plain 62-us PayPal response', async () => {
  const code = await fetchManualCode(
    "yes|PayPal: 201412 is your security code. Don't share it.|(PayPal)|到期时间：2026-07-29 00:00:00"
  );

  assert.equal(code, '201412');
});

test('manual hosted checkout code fetch extracts nested tgflare PayPal response', async () => {
  const code = await fetchManualCode({
    code: 1,
    msg: 'ok',
    data: {
      code: "PayPal: 288652 is your security code. Don't share it.",
      code_time: '2026-05-22 12:25:10',
      expired_date: '2026-07-31 00:00:00',
    },
  });

  assert.equal(code, '288652');
});

test('manual hosted checkout code fetch extracts issue 29 nested data.code response', async () => {
  const code = await fetchManualCode({
    code: 1,
    msg: 'ok',
    data: {
      code: 'PayPal: 011119 is your security code. Don`t share it.',
      code_time: '2026-05-21 10:37:02',
      expired_date: '2026-06-14 00:00:00',
    },
  });

  assert.equal(code, '011119');
});

test('manual hosted checkout code fetch extracts separated security code digits', async () => {
  const code = await fetchManualCode(
    "yes|PayPal: 1 2 3 4 5 6 is your security code. Don't share it.|(PayPal)|到期时间：2026-07-29 00:00:00"
  );

  assert.equal(code, '123456');
});

test('manual hosted checkout code fetch ignores metadata phone before sms text', async () => {
  const code = await fetchManualCode({
    data: {
      phone: '+14155552671',
      sms: "PayPal: 288652 is your security code. Don't share it.",
    },
  });

  assert.equal(code, '288652');
});

test('manual hosted checkout code fetch ignores metadata order id before message text', async () => {
  const code = await fetchManualCode({
    data: {
      order_id: '123456',
      message: "PayPal: 288652 is your security code. Don't share it.",
    },
  });

  assert.equal(code, '288652');
});

test('manual hosted checkout code fetch ignores PayPal confirmation text with expiration date', async () => {
  const executor = createExecutorWithPayload(
    'yes|PayPal: Thanks for confirming your phone number. Log in or get the app to get transaction alerts: https://py.pl/24BgEk|(PayPal)|到期时间：2026-07-29 00:00:00'
  );

  await assert.rejects(
    () => executor.fetchHostedCheckoutVerificationCodeManually({
      verificationUrl: 'http://example.test/api/get_sms?key=test',
    }),
    /暂未返回有效验证码/
  );
});

test('manual hosted checkout code fetch times out stalled sms endpoint', async () => {
  let aborted = false;
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    fetch: async (_url, options = {}) => {
      await new Promise((resolve, reject) => {
        const timer = setTimeout(resolve, 120);
        options.signal?.addEventListener('abort', () => {
          clearTimeout(timer);
          aborted = true;
          reject(Object.assign(new Error('aborted'), { name: 'AbortError' }));
        }, { once: true });
      });
      return {
        text: async () => '',
      };
    },
  });

  await assert.rejects(
    () => executor.fetchHostedCheckoutVerificationCodeManually({
      verificationUrl: 'http://example.test/api/get_sms?key=test',
      timeoutMs: 10,
    }),
    /验证码接口请求超时/
  );
  assert.equal(aborted, true);
});

test('hosted checkout sms pool accepts phone pipe url entries', async () => {
  const patches = [];
  let requestedUrl = '';
  const verificationUrl = 'https://sms.699.chat/api/get_sms?key=251a4a4760a848ba4920cb179f589236';
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    fetch: async (url) => {
      requestedUrl = String(url || '');
      return {
        text: async () => "yes|PayPal: 201412 is your security code. Don't share it.|(PayPal)|到期时间：2026-07-29 00:00:00",
      };
    },
    getState: async () => ({
      hostedCheckoutSmsPoolText: `+15824441369|${verificationUrl}`,
      hostedCheckoutSmsPoolUsage: {},
    }),
    setState: async (patch) => {
      patches.push(patch);
    },
  });

  const result = await executor.fetchHostedCheckoutVerificationCodeManually();
  const selectedEntry = patches.find((patch) => patch.hostedCheckoutCurrentSmsEntry)?.hostedCheckoutCurrentSmsEntry;

  assert.equal(result.code, '201412');
  assert.equal(result.verificationUrl, verificationUrl);
  assert.match(requestedUrl, /^https:\/\/sms\.699\.chat\/api\/get_sms\?key=251a4a4760a848ba4920cb179f589236&t=\d+$/);
  assert.equal(selectedEntry.phone, '5824441369');
  assert.equal(selectedEntry.verificationUrl, verificationUrl);
});

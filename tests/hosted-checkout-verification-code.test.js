const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

require('../background/steps/create-plus-checkout.js');

const repoRoot = path.resolve(__dirname, '..');
const createPlusCheckoutSource = fs.readFileSync(
  path.join(repoRoot, 'background/steps/create-plus-checkout.js'),
  'utf8'
);
const paypalFlowSource = fs.readFileSync(
  path.join(repoRoot, 'content/paypal-flow.js'),
  'utf8'
);
const plusCheckoutSource = fs.readFileSync(
  path.join(repoRoot, 'content/plus-checkout.js'),
  'utf8'
);

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

test('hosted PayPal no-SMS polling clicks resend exactly three times', async () => {
  const logs = [];
  const messages = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    addLog: async (message, level) => logs.push({ message, level }),
    enableTestHooks: true,
    fetch: async () => ({
      status: 200,
      text: async () => JSON.stringify({ data: { sms: '' } }),
    }),
    getState: async () => ({
      hostedCheckoutSmsPoolText: '+15824441369|http://example.test/api/get_sms?key=test',
      hostedCheckoutSmsPoolUsage: {},
    }),
    setState: async () => {},
    sleepWithStop: async () => {},
    waitForTabCompleteUntilStopped: async () => {},
    ensureContentScriptReadyOnTabUntilStopped: async () => {},
    sendTabMessageUntilStopped: async (_tabId, _source, message = {}) => {
      messages.push(message);
      return { stage: 'verification', resendClicked: true };
    },
  });

  await assert.rejects(
    () => executor.__test.pollHostedCheckoutVerificationCodeWithResend(1),
    /自动点击重新发送 3 次后仍未获取验证码/
  );

  const resendMessages = messages.filter((message) => message.payload?.resendVerificationCode);
  assert.equal(resendMessages.length, 3);
  assert.equal(logs.filter((log) => /正在点击 PayPal “重新发送”验证码/.test(log.message)).length, 3);
});

test('hosted PayPal polling does not resend on sms endpoint HTTP failure', async () => {
  const messages = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    fetch: async () => ({
      status: 500,
      text: async () => JSON.stringify({ error: 'server error' }),
    }),
    getState: async () => ({
      hostedCheckoutSmsPoolText: '+15824441369|http://example.test/api/get_sms?key=test',
      hostedCheckoutSmsPoolUsage: {},
    }),
    setState: async () => {},
    sleepWithStop: async () => {},
    waitForTabCompleteUntilStopped: async () => {},
    ensureContentScriptReadyOnTabUntilStopped: async () => {},
    sendTabMessageUntilStopped: async (_tabId, _source, message = {}) => {
      messages.push(message);
      return { stage: 'verification', resendClicked: true };
    },
  });

  await assert.rejects(
    () => executor.__test.pollHostedCheckoutVerificationCodeWithResend(1),
    /HTTP 500/
  );

  assert.equal(messages.filter((message) => message.payload?.resendVerificationCode).length, 0);
});

test('hosted PayPal no-SMS flow resends verification code up to three times', () => {
  assert.match(
    createPlusCheckoutSource,
    /const HOSTED_CHECKOUT_VERIFICATION_NO_SMS_RESEND_MAX_ATTEMPTS = 3;/
  );
  assert.match(createPlusCheckoutSource, /enableTestHooks = false/);
  assert.match(createPlusCheckoutSource, /enableTestHooks \? \{ __test: \{/);
  assert.match(createPlusCheckoutSource, /buildHostedCheckoutReplacementCard,/);
  assert.match(createPlusCheckoutSource, /pollHostedCheckoutVerificationCodeWithResend,/);
  assert.match(createPlusCheckoutSource, /async function pollHostedCheckoutVerificationCodeWithResend\(tabId, options = \{\}\)/);
  assert.match(createPlusCheckoutSource, /if \(isHostedCheckoutVerificationNoSmsError\(lastError\)\)/);
  assert.match(createPlusCheckoutSource, /tracker\.noSmsResendAttempts \+= 1;/);
  assert.match(createPlusCheckoutSource, /const hostedNoSmsResendTracker = \{ noSmsResendAttempts: 0 \};/);
  assert.match(createPlusCheckoutSource, /isHostedCheckoutVerificationNoSmsError\(error\)/);
  assert.match(createPlusCheckoutSource, /if \(!isHostedCheckoutVerificationNoSmsError\(error\)\) \{\s*throw error;\s*\}/);
  assert.match(createPlusCheckoutSource, /hosted checkout 验证码接口请求失败（HTTP \$\{status\}）。/);
  assert.match(createPlusCheckoutSource, /clickHostedCheckoutVerificationResend\(/);
  assert.match(createPlusCheckoutSource, /resendVerificationCode: true/);
  assert.match(createPlusCheckoutSource, /本轮轮询未收到 PayPal 短信验证码/);
  assert.match(createPlusCheckoutSource, /const verificationCode = await pollHostedCheckoutVerificationCodeWithResend\(tabId, \{\s*tracker: hostedNoSmsResendTracker,\s*\}\);/);
});

test('PayPal HAR fixtures expose verification resend controls used by automation', () => {
  const har03 = JSON.parse(fs.readFileSync(path.join(repoRoot, 'www.paypal.com03.har'), 'utf8'));
  const har04 = JSON.parse(fs.readFileSync(path.join(repoRoot, 'www.paypal.com04.har'), 'utf8'));
  const har04Text = har04.log.entries
    .map((entry) => entry.response?.content?.text || '')
    .join('\n');

  assert.ok(
    har03.log.entries.some((entry) => /geo\.ddc\.paypal\.com\/captcha/i.test(entry.request?.url || '')),
    'HAR03 should capture the PayPal/DataDome captcha branch before the OTP page stabilizes'
  );
  assert.match(har04Text, /"resendButtonTitle":"重新发送"/);
  assert.match(har04Text, /data-testid["']?:["']link-get-new-code|data-testid=["']link-get-new-code/);
  assert.match(har04Text, /ci-ciBasic-0/);
  assert.match(paypalFlowSource, /link-get-new-code/);
  assert.match(paypalFlowSource, /#linkGetNewCode/);
});

test('hosted PayPal card decline replaces number expiry and cvv together', () => {
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
  });
  const previousCard = {
    cardNumber: '4147525524970946',
    cardExpiry: '04 / 28',
    cardCvv: '566',
  };
  for (let index = 0; index < 50; index += 1) {
    const card = executor.__test.buildHostedCheckoutReplacementCard(previousCard);
    assert.notEqual(card.number, previousCard.cardNumber);
    assert.notEqual(card.expiry, previousCard.cardExpiry);
    assert.notEqual(card.cvv, previousCard.cardCvv);
  }

  assert.match(createPlusCheckoutSource, /const HOSTED_CHECKOUT_CARD_DECLINE_MAX_REPLACEMENTS = 3;/);
  assert.match(createPlusCheckoutSource, /function buildHostedCheckoutReplacementCard\(previousCard = \{\}, excludedSignatures = new Set\(\)\)/);
  assert.match(createPlusCheckoutSource, /buildHostedCheckoutReplacementCard,/);
  assert.match(createPlusCheckoutSource, /const signature = getHostedCheckoutCardFieldsSignature\(card\.number, card\.expiry, card\.cvv\);/);
  assert.match(createPlusCheckoutSource, /!excludedSignatures\.has\(signature\)/);
  assert.match(createPlusCheckoutSource, /function buildHostedCheckoutFallbackReplacementCard\(/);
  assert.match(createPlusCheckoutSource, /return buildHostedCheckoutFallbackReplacementCard\(/);
  assert.match(createPlusCheckoutSource, /card\.number !== previousNumber/);
  assert.match(createPlusCheckoutSource, /card\.expiry !== previousExpiry/);
  assert.match(createPlusCheckoutSource, /card\.cvv !== previousCvv/);
  assert.match(createPlusCheckoutSource, /replaceHostedCheckoutGuestProfileCard\(guestProfile, attemptedCardSignatures\)/);
  assert.match(createPlusCheckoutSource, /const attemptedCardSignatures = new Set\(\);/);
  assert.match(createPlusCheckoutSource, /attemptedCardSignatures\.add\(currentCardSignature\);/);
  assert.match(createPlusCheckoutSource, /attemptedCardSignatures\.add\(lastSubmittedCardSignature\);/);
  assert.match(createPlusCheckoutSource, /function getHostedCheckoutCardSignature\(guestProfile = \{\}\)/);
  assert.match(createPlusCheckoutSource, /pageState\.hostedCardDeclined\s*&& pageState\.hostedGuestCheckoutFormVisible\s*&& !pageState\.verificationInputsVisible/);
  assert.match(createPlusCheckoutSource, /lastSubmittedCardSignature === currentCardSignature/);
  assert.match(createPlusCheckoutSource, /刚提交新卡后短暂等待页面刷新/);
  assert.match(createPlusCheckoutSource, /lastSubmittedCardAt = Date\.now\(\);/);
  assert.match(createPlusCheckoutSource, /PayPal 提示 We weren.t able to add this card，正在更换卡号\/有效期\/CVV/);
  assert.match(paypalFlowSource, /getPayPalHostedCardDeclinedMessage/);
  assert.doesNotMatch(paypalFlowSource, /\)\|try\\s\+a\\s\+different\\s\+card\|/);
  assert.match(paypalFlowSource, /Check\\s\+all\\s\+the\\s\+details\\s\+are\\s\+correct/);
  assert.match(paypalFlowSource, /const hasHostedGuestCheckout = isPayPalHostedGuestCheckoutPage\(\);/);
  assert.match(paypalFlowSource, /function hasPayPalHostedGuestCheckoutForm\(\)/);
  assert.match(paypalFlowSource, /hostedGuestCheckoutFormVisible/);
  assert.match(paypalFlowSource, /hostedCardDeclined: hasPayPalHostedCardDeclinedError\(\)/);
});

test('hosted checkout uses Hero email from state for PayPal guest profile', async () => {
  const patches = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    getState: async () => ({
      normalHeroModeEnabled: true,
      email: 'HeroUser@Example.COM',
      hostedCheckoutPhoneNumber: '5551112222',
    }),
    setState: async (patch) => patches.push(patch),
    broadcastDataUpdate: (patch) => patches.push({ broadcast: patch }),
    chrome: {
      storage: {
        local: {
          get: async () => ({}),
        },
      },
    },
  });

  const config = await executor.__test.getHostedCheckoutRuntimeConfig({ ensureEmail: true });
  const profile = executor.__test.buildHostedCheckoutGuestProfile({}, config);

  assert.equal(config.email, 'herouser@example.com');
  assert.equal(profile.email, 'herouser@example.com');
  assert.equal(patches.length, 0);
});

test('hosted checkout prompts for Hero email when payment page lacks email', async () => {
  const patches = [];
  const requests = [];
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    getState: async () => ({
      normalHeroModeEnabled: true,
      signupPhoneNumber: '+57 (324) 132 10 49',
    }),
    requestManualAddEmailInput: async (payload) => {
      requests.push(payload);
      return 'FallbackHero@Example.COM';
    },
    setState: async (patch) => patches.push(patch),
    broadcastDataUpdate: (patch) => patches.push({ broadcast: patch }),
    chrome: {
      storage: {
        local: {
          get: async () => ({}),
        },
      },
    },
  });

  const email = await executor.__test.ensureHostedCheckoutEmail({
    normalHeroModeEnabled: true,
    signupPhoneNumber: '+57 (324) 132 10 49',
  });

  assert.equal(email, 'fallbackhero@example.com');
  assert.equal(requests.length, 1);
  assert.equal(requests[0].phoneNumber, '+57 (324) 132 10 49');
  assert.deepEqual(JSON.parse(JSON.stringify(patches[0])), {
    email: 'fallbackhero@example.com',
    registrationEmailState: {
      current: 'fallbackhero@example.com',
      previous: 'fallbackhero@example.com',
      source: 'normal_hero_checkout',
      updatedAt: patches[0].registrationEmailState.updatedAt,
    },
    manualAddEmailInputRequired: true,
  });
});

test('hosted checkout automation sends Hero startup email to PayPal flow', async () => {
  const paypalPayloads = [];
  const openAiPayloads = [];
  let paypalSubmitted = false;
  const executor = globalThis.MultiPageBackgroundPlusCheckoutCreate.createPlusCheckoutCreateExecutor({
    enableTestHooks: true,
    enableHostedCheckoutAutomation: true,
    addLog: async () => {},
    fetch: async () => ({
      ok: true,
      status: 200,
      json: async () => ({
        address: {
          Address: '123 Main St',
          City: 'New York',
          State_Full: 'New York',
          Zip_Code: '10001',
        },
      }),
    }),
    getState: async () => ({
      normalHeroModeEnabled: true,
      email: 'HeroUser@Example.COM',
      hostedCheckoutPhoneNumber: '5551112222',
    }),
    chrome: {
      tabs: {
        get: async () => ({
          id: 7,
          url: paypalSubmitted
            ? 'https://chatgpt.com/payments/success'
            : 'https://www.paypal.com/checkoutweb/guest',
        }),
      },
      storage: {
        local: {
          get: async () => ({}),
        },
      },
    },
    completeNodeFromBackground: async () => {},
    ensureContentScriptReadyOnTabUntilStopped: async () => {},
    sendTabMessageUntilStopped: async (_tabId, source, message = {}) => {
      if (source === 'plus-checkout') {
        if (message.type === 'RUN_HOSTED_OPENAI_CHECKOUT_STEP') {
          openAiPayloads.push(message.payload);
        }
        return {};
      }
      if (message.type === 'PAYPAL_HOSTED_GET_STATE') {
        return { hostedStage: 'guest_checkout' };
      }
      if (message.type === 'PAYPAL_RUN_HOSTED_CHECKOUT_STEP') {
        paypalPayloads.push(message.payload);
        paypalSubmitted = true;
        return {};
      }
      return {};
    },
    sleepWithStop: async () => {},
    waitForTabCompleteUntilStopped: async () => {},
    waitForTabUrlMatchUntilStopped: async (_tabId, matcher) => {
      const url = 'https://www.paypal.com/checkoutweb/guest';
      return matcher(url) ? { id: 7, url } : null;
    },
  });

  await executor.__test.runHostedCheckoutAutomation(7, { checkoutSessionId: 'cs_test' });

  assert.equal(openAiPayloads.length, 1);
  assert.equal(openAiPayloads[0].email, 'herouser@example.com');
  assert.equal(paypalPayloads.length, 1);
  assert.equal(paypalPayloads[0].email, 'herouser@example.com');
});

test('hosted OpenAI checkout waits for and advances empty email step before address fill', () => {
  assert.match(plusCheckoutSource, /async function waitForHostedOpenAiEmailInput\(timeoutMs = 6000\)/);
  assert.match(plusCheckoutSource, /const input = await waitForHostedOpenAiEmailInput\(options\.timeoutMs\);/);
  assert.match(plusCheckoutSource, /function findHostedOpenAiEmailSubmitButton\(\)/);
  assert.match(plusCheckoutSource, /async function advanceHostedOpenAiEmailStepIfNeeded\(emailResult = \{\}\)/);
  assert.match(plusCheckoutSource, /const button = findHostedOpenAiEmailSubmitButton\(\);/);
  assert.match(plusCheckoutSource, /dispatchHostedOpenAiClick\(button\);/);
  assert.match(plusCheckoutSource, /waitForHostedOpenAiStructuredAddressInput\(10000\);/);
  assert.match(plusCheckoutSource, /emailStepAdvanced: Boolean\(emailStepResult\?\.emailStepAdvanced\)/);
});

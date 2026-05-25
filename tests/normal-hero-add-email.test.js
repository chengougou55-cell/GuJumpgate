const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');

function loadStep8Module() {
  const sandbox = {
    self: {},
    console,
    setTimeout,
    clearTimeout,
  };
  vm.runInNewContext(
    fs.readFileSync(path.join(repoRoot, 'background/steps/fetch-login-code.js'), 'utf8'),
    sandbox
  );
  return sandbox.self.MultiPageBackgroundStep8;
}

function createBaseDeps(calls = []) {
  return {
    addLog: async (message, level) => calls.push(['log', level || 'info', message]),
    chrome: {
      tabs: {
        update: async (...args) => calls.push(['tab.update', ...args]),
      },
    },
    getOAuthFlowStepTimeoutMs: async (fallback) => fallback,
    getState: async () => ({}),
    getTabId: async () => 101,
    setState: async (updates) => calls.push(['setState', updates]),
    throwIfStopped: () => {},
  };
}

test('Normal Hero add-email prompts manual email and retries when email is already used', async () => {
  const calls = [];
  let promptCount = 0;
  const executor = loadStep8Module().createStep8Executor({
    ...createBaseDeps(calls),
    getState: async () => ({
      normalHeroModeEnabled: true,
      signupPhoneNumber: '+57 (324) 132 10 49',
    }),
    sendToContentScriptResilient: async (_target, message) => {
      calls.push(['content', message.type, message.payload]);
      if (message.type === 'GET_LOGIN_AUTH_STATE') {
        return { state: 'add_email_page', url: 'https://auth.openai.com/add-email' };
      }
      if (message.type === 'SUBMIT_ADD_EMAIL' && message.payload.email === 'used@example.com') {
        return { error: 'STEP8_EMAIL_IN_USE::email_in_use' };
      }
      return {
        displayedEmail: message.payload.email,
        url: 'https://auth.openai.com/verify',
      };
    },
    requestManualAddEmailInput: async () => {
      promptCount += 1;
      return promptCount === 1 ? 'used@example.com' : 'fresh@example.com';
    },
    persistRegistrationEmailState: async (_state, email, options) => calls.push(['persistEmail', email, options]),
    completeNodeFromBackground: async (nodeId, payload) => calls.push(['complete', nodeId, payload]),
  });

  await executor.executeBindEmail({
    normalHeroModeEnabled: true,
    signupPhoneNumber: '+57 (324) 132 10 49',
    oauthUrl: 'https://example.test/oauth',
    nodeId: 'bind-email',
    visibleStep: 9,
  });

  assert.equal(promptCount, 2);
  assert.deepEqual(
    calls.filter((item) => item[0] === 'content' && item[1] === 'SUBMIT_ADD_EMAIL').map((item) => item[2].email),
    ['used@example.com', 'fresh@example.com']
  );
  assert.deepEqual(JSON.parse(JSON.stringify(calls.find((item) => item[0] === 'complete'))), [
    'complete',
    'bind-email',
    {
      bindEmailSubmitted: true,
      email: 'fresh@example.com',
      step8VerificationTargetEmail: 'fresh@example.com',
      manualAddEmailInputRequired: true,
    },
  ]);
});

test('Normal Hero bind-email verification prompts manual code again after invalid code', async () => {
  const calls = [];
  let promptCount = 0;
  const executor = loadStep8Module().createStep8Executor({
    ...createBaseDeps(calls),
    sendToContentScriptResilient: async (_target, message) => {
      calls.push(['content', message.type, message.payload]);
      if (message.type === 'GET_LOGIN_AUTH_STATE') {
        return {
          state: 'verification_page',
          displayedEmail: 'fresh@example.com',
          url: 'https://auth.openai.com/verify',
        };
      }
      return {};
    },
    requestManualEmailCodeInput: async () => {
      promptCount += 1;
      return promptCount === 1 ? '111111' : '222222';
    },
    resolveVerificationStep: async (_step, _state, _mail, options) => {
      calls.push(['resolve', options.manualCode, options.completionStep]);
      if (options.manualCode === '111111') {
        return { invalidCode: true, errorText: 'invalid otp' };
      }
      return { code: options.manualCode };
    },
  });

  const result = await executor.executeFetchBindEmailCode({
    normalHeroModeEnabled: true,
    manualAddEmailInputRequired: true,
    bindEmailSubmitted: true,
    email: 'fresh@example.com',
    step8VerificationTargetEmail: 'fresh@example.com',
    visibleStep: 10,
  });

  assert.equal(promptCount, 2);
  assert.deepEqual(calls.filter((item) => item[0] === 'resolve'), [
    ['resolve', '111111', 10],
    ['resolve', '222222', 10],
  ]);
  assert.deepEqual(result, { code: '222222' });
  assert.equal(calls.some((item) => item[0] === 'log' && String(item[2]).includes('111111')), false);
  assert.equal(calls.some((item) => item[0] === 'log' && String(item[2]).includes('222222')), false);
});

test('Normal Hero bound-email relogin keeps manual email-code mode through step 11', async () => {
  const calls = [];
  const executor = loadStep8Module().createStep8Executor({
    ...createBaseDeps(calls),
    ensureStep8VerificationPageReady: async () => ({
      state: 'verification_page',
      displayedEmail: 'fresh@example.com',
      url: 'https://auth.openai.com/verify',
    }),
    requestManualEmailCodeInput: async () => '654321',
    resolveVerificationStep: async (_step, state, mail, options) => {
      calls.push(['resolve', state.step8VerificationTargetEmail, mail.provider, options.manualCode, options.completionStep]);
      return { code: options.manualCode };
    },
  });

  const result = await executor.executeBoundEmailLoginCode({
    normalHeroModeEnabled: true,
    manualAddEmailInputRequired: true,
    email: 'fresh@example.com',
    step8VerificationTargetEmail: 'fresh@example.com',
    oauthUrl: 'https://example.test/oauth',
    nodeId: 'fetch-bound-email-login-code',
    visibleStep: 11,
  });

  assert.deepEqual(calls.find((item) => item[0] === 'resolve'), [
    'resolve',
    'fresh@example.com',
    'manual',
    '654321',
    11,
  ]);
  assert.deepEqual(result, { code: '654321' });
});

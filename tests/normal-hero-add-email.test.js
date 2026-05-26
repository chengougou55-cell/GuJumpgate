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

function loadMessageRouterModule() {
  const sandbox = {
    self: {},
    console,
    setTimeout,
    clearTimeout,
  };
  vm.runInNewContext(
    fs.readFileSync(path.join(repoRoot, 'background/message-router.js'), 'utf8'),
    sandbox
  );
  return sandbox.self.MultiPageBackgroundMessageRouter;
}

function loadSignupFlowHelpersModule() {
  const sandbox = {
    self: {},
    console,
    setTimeout,
    clearTimeout,
  };
  vm.runInNewContext(
    fs.readFileSync(path.join(repoRoot, 'background/signup-flow-helpers.js'), 'utf8'),
    sandbox
  );
  return sandbox.self.MultiPageSignupFlowHelpers;
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

test('Normal Hero auto-run payload stores startup email while preserving phone identity', async () => {
  let state = {};
  let started = null;
  const patches = [];
  const router = loadMessageRouterModule().createMessageRouter({
    addLog: async () => {},
    buildXiaohongshuRuntimeReset: () => ({}),
    clearStopRequest: () => {},
    getPendingAutoRunTimerPlan: () => null,
    getState: async () => state,
    isAutoRunLockedState: () => false,
    normalizeRunCount: (value) => Math.max(1, Math.floor(Number(value) || 1)),
    setState: async (patch) => {
      patches.push(patch);
      state = { ...state, ...patch };
    },
    startAutoRunLoop: (totalRuns, options) => {
      started = { totalRuns, options };
    },
    validateAutoRunStart: () => ({ ok: true, errors: [] }),
  });

  const response = await router.handleMessage({
    type: 'AUTO_RUN',
    source: 'sidepanel',
    payload: {
      totalRuns: 1,
      normalHeroModeEnabled: true,
      signupPhoneNumber: '+57 (324) 132 10 49',
      signupEmail: 'HeroUser@Example.COM',
    },
  }, {});

  assert.deepEqual(JSON.parse(JSON.stringify(response)), { ok: true });
  assert.equal(state.normalHeroModeEnabled, true);
  assert.equal(state.manualSignupPhoneSmsEnabled, true);
  assert.equal(state.signupMethod, 'phone');
  assert.equal(state.accountIdentifierType, 'phone');
  assert.equal(state.accountIdentifier, '+57 (324) 132 10 49');
  assert.equal(state.email, 'herouser@example.com');
  assert.equal(state.registrationEmailState.current, 'herouser@example.com');
  assert.equal(state.registrationEmailState.source, 'normal_hero_start');
  assert.equal(state.normalHeroEmailRuntime, true);
  assert.equal(state.manualAddEmailInputRequired, true);
  assert.equal(started.totalRuns, 1);
  assert.equal(started.options.normalHeroModeEnabled, true);
  assert.ok(patches.some((patch) => patch.email === 'herouser@example.com'));
});

test('ordinary auto-run clears stale Normal Hero startup email before start', async () => {
  let state = {
    normalHeroModeEnabled: true,
    manualSignupPhoneSmsEnabled: true,
    manualAddEmailInputRequired: true,
    signupMethod: 'phone',
    resolvedSignupMethod: 'phone',
    signupPhoneNumber: '+57 (324) 132 10 49',
    accountIdentifierType: 'phone',
    accountIdentifier: '+57 (324) 132 10 49',
    email: 'HeroUser@Example.COM',
    normalHeroEmailRuntime: true,
    registrationEmailState: {
      current: 'HeroUser@Example.COM',
      previous: 'HeroUser@Example.COM',
      source: 'normal_hero_start',
      updatedAt: 1,
    },
  };
  let started = null;
  const router = loadMessageRouterModule().createMessageRouter({
    addLog: async () => {},
    buildXiaohongshuRuntimeReset: () => ({}),
    clearStopRequest: () => {},
    getPendingAutoRunTimerPlan: () => null,
    getState: async () => state,
    isAutoRunLockedState: () => false,
    normalizeRunCount: (value) => Math.max(1, Math.floor(Number(value) || 1)),
    setState: async (patch) => {
      state = { ...state, ...patch };
    },
    startAutoRunLoop: (totalRuns, options) => {
      started = { totalRuns, options };
    },
    validateAutoRunStart: () => ({ ok: true, errors: [] }),
  });

  const response = await router.handleMessage({
    type: 'AUTO_RUN',
    source: 'sidepanel',
    payload: {
      totalRuns: 1,
      mode: 'restart',
    },
  }, {});

  assert.deepEqual(JSON.parse(JSON.stringify(response)), { ok: true });
  assert.equal(state.normalHeroModeEnabled, false);
  assert.equal(state.manualSignupPhoneSmsEnabled, false);
  assert.equal(state.manualAddEmailInputRequired, false);
  assert.equal(state.email, null);
  assert.equal(state.normalHeroEmailRuntime, false);
  assert.deepEqual(JSON.parse(JSON.stringify(state.registrationEmailState)), {
    current: '',
    previous: '',
    source: '',
    updatedAt: 0,
  });
  assert.equal(state.accountIdentifierType, null);
  assert.equal(state.accountIdentifier, '');
  assert.equal(state.signupPhoneNumber, '');
  assert.equal(started.options.normalHeroModeEnabled, false);
});

test('xiaohongshu auto-run clears stale Normal Hero identity before checkout start', async () => {
  let state = {
    normalHeroModeEnabled: true,
    manualSignupPhoneSmsEnabled: true,
    manualAddEmailInputRequired: true,
    signupMethod: 'phone',
    resolvedSignupMethod: 'phone',
    signupPhoneNumber: '+57 (324) 132 10 49',
    accountIdentifierType: 'phone',
    accountIdentifier: '+57 (324) 132 10 49',
    email: 'HeroUser@Example.COM',
    normalHeroEmailRuntime: true,
    registrationEmailState: {
      current: 'HeroUser@Example.COM',
      previous: 'HeroUser@Example.COM',
      source: 'normal_hero_start',
      updatedAt: 1,
    },
  };
  let started = null;
  const router = loadMessageRouterModule().createMessageRouter({
    addLog: async () => {},
    buildXiaohongshuRuntimeReset: () => ({}),
    clearStopRequest: () => {},
    getPendingAutoRunTimerPlan: () => null,
    getState: async () => state,
    isAutoRunLockedState: () => false,
    normalizeRunCount: (value) => Math.max(1, Math.floor(Number(value) || 1)),
    getNodeIdsForState: () => [
      'open-chatgpt',
      'submit-signup-email',
      'plus-checkout-create',
      'sub2api-session-import',
    ],
    setState: async (patch) => {
      state = { ...state, ...patch };
    },
    startAutoRunLoop: (totalRuns, options) => {
      started = { totalRuns, options };
    },
    validateAutoRunStart: () => ({ ok: true, errors: [] }),
  });

  const response = await router.handleMessage({
    type: 'AUTO_RUN_XIAOHONGSHU',
    source: 'sidepanel',
    payload: {
      accessToken: 'xiaohongshu-token',
    },
  }, {});

  assert.deepEqual(JSON.parse(JSON.stringify(response)), { ok: true });
  assert.equal(state.xiaohongshuModeEnabled, true);
  assert.equal(state.normalHeroModeEnabled, false);
  assert.equal(state.manualSignupPhoneSmsEnabled, false);
  assert.equal(state.email, null);
  assert.equal(state.normalHeroEmailRuntime, false);
  assert.deepEqual(JSON.parse(JSON.stringify(state.registrationEmailState)), {
    current: '',
    previous: '',
    source: '',
    updatedAt: 0,
  });
  assert.equal(state.accountIdentifierType, null);
  assert.equal(state.accountIdentifier, '');
  assert.equal(state.signupPhoneNumber, '');
  assert.equal(started.options.xiaohongshuModeEnabled, true);
  assert.equal(started.options.mode, 'continue');
});

test('ordinary signup email generation ignores stale Normal Hero startup email', async () => {
  const calls = [];
  const helpers = loadSignupFlowHelpersModule().createSignupFlowHelpers({
    fetchGeneratedEmail: async (state) => {
      calls.push(['fetchGeneratedEmail', state.email, state.registrationEmailState, state.manualAddEmailInputRequired]);
      return 'generated@example.com';
    },
    isGeneratedAliasProvider: () => false,
    isHotmailProvider: () => false,
    isLuckmailProvider: () => false,
    persistRegistrationEmailState: async (_state, email, options) => {
      calls.push(['persistRegistrationEmailState', email, options]);
    },
  });

  const email = await helpers.resolveSignupEmailForFlow({
    normalHeroModeEnabled: false,
    manualSignupPhoneSmsEnabled: false,
    manualAddEmailInputRequired: true,
    email: 'HeroUser@Example.COM',
    registrationEmailState: {
      current: 'HeroUser@Example.COM',
      previous: 'HeroUser@Example.COM',
      source: 'normal_hero_start',
      updatedAt: 1,
    },
    emailGenerator: 'duck',
  });

  assert.equal(email, 'generated@example.com');
  assert.equal(calls[0][0], 'fetchGeneratedEmail');
  assert.equal(calls[0][1], null);
  assert.deepEqual(JSON.parse(JSON.stringify(calls[0][2])), {
    current: '',
    previous: '',
    source: '',
    updatedAt: 0,
  });
  assert.equal(calls[0][3], false);
});

test('ordinary signup email generation ignores Hero email after bind-email source changes', async () => {
  const calls = [];
  const helpers = loadSignupFlowHelpersModule().createSignupFlowHelpers({
    fetchGeneratedEmail: async (state) => {
      calls.push(['fetchGeneratedEmail', state.email, state.registrationEmailState, state.normalHeroEmailRuntime]);
      return 'generated-after-bind@example.com';
    },
    isGeneratedAliasProvider: () => false,
    isHotmailProvider: () => false,
    isLuckmailProvider: () => false,
    persistRegistrationEmailState: async (_state, email, options) => {
      calls.push(['persistRegistrationEmailState', email, options]);
    },
  });

  const email = await helpers.resolveSignupEmailForFlow({
    normalHeroModeEnabled: false,
    manualSignupPhoneSmsEnabled: false,
    manualAddEmailInputRequired: false,
    normalHeroEmailRuntime: true,
    email: 'bound-hero@example.com',
    registrationEmailState: {
      current: 'bound-hero@example.com',
      previous: 'herouser@example.com',
      source: 'bind_email',
      updatedAt: 2,
    },
    emailGenerator: 'duck',
  });

  assert.equal(email, 'generated-after-bind@example.com');
  assert.equal(calls[0][0], 'fetchGeneratedEmail');
  assert.equal(calls[0][1], null);
  assert.equal(calls[0][3], false);
});

test('ordinary generated alias clears stale Hero runtime before persistence', async () => {
  const calls = [];
  const helpers = loadSignupFlowHelpersModule().createSignupFlowHelpers({
    buildGeneratedAliasEmail: (state) => {
      calls.push(['buildGeneratedAliasEmail', state.email, state.normalHeroEmailRuntime]);
      return 'alias-after-hero@example.com';
    },
    isGeneratedAliasProvider: () => true,
    isReusableGeneratedAliasEmail: () => false,
    isHotmailProvider: () => false,
    isLuckmailProvider: () => false,
    persistRegistrationEmailState: async (state, email, options) => {
      calls.push(['persistRegistrationEmailState', state.email, state.normalHeroEmailRuntime, email, options]);
    },
  });

  const email = await helpers.resolveSignupEmailForFlow({
    normalHeroModeEnabled: false,
    manualSignupPhoneSmsEnabled: false,
    manualAddEmailInputRequired: false,
    normalHeroEmailRuntime: true,
    email: 'bound-hero@example.com',
    registrationEmailState: {
      current: 'bound-hero@example.com',
      previous: 'herouser@example.com',
      source: 'bind_email',
      updatedAt: 2,
    },
    emailGenerator: 'alias',
  });

  assert.equal(email, 'alias-after-hero@example.com');
  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [
    ['buildGeneratedAliasEmail', null, false],
    [
      'persistRegistrationEmailState',
      null,
      false,
      'alias-after-hero@example.com',
      {
        source: 'flow',
        preserveAccountIdentity: false,
      },
    ],
  ]);
});

test('email state writers clear stale Hero runtime unless explicitly preserved', () => {
  const backgroundScript = fs.readFileSync(path.join(repoRoot, 'background.js'), 'utf8');
  const setEmailStateMatch = backgroundScript.match(/async function setEmailStateSilently[\s\S]*?\n}\n\nasync function setEmailState/);
  const persistEmailStateMatch = backgroundScript.match(/async function persistRegistrationEmailState[\s\S]*?\n}\n\nasync function setSignupPhoneStateSilently/);

  assert.ok(setEmailStateMatch, 'setEmailStateSilently block should be present');
  assert.match(
    setEmailStateMatch[0],
    /const nextNormalHeroEmailRuntime = options\?\.normalHeroEmailRuntime !== undefined[\s\S]*\? Boolean\(options\.normalHeroEmailRuntime\)[\s\S]*: false;/
  );
  assert.match(setEmailStateMatch[0], /updates\.normalHeroEmailRuntime = nextNormalHeroEmailRuntime;/);

  assert.ok(persistEmailStateMatch, 'persistRegistrationEmailState block should be present');
  assert.match(
    persistEmailStateMatch[0],
    /const nextNormalHeroEmailRuntime = options\?\.normalHeroEmailRuntime !== undefined[\s\S]*\? Boolean\(options\.normalHeroEmailRuntime\)[\s\S]*: false;/
  );
  assert.match(persistEmailStateMatch[0], /const updates = \{ normalHeroEmailRuntime: nextNormalHeroEmailRuntime \};/);
  assert.match(persistEmailStateMatch[0], /preservedUpdates\.normalHeroEmailRuntime = nextNormalHeroEmailRuntime;/);
  assert.match(persistEmailStateMatch[0], /updates\.normalHeroEmailRuntime = nextNormalHeroEmailRuntime;/);
});

test('xiaohongshu auto-run clears Hero email after bind-email source changes', async () => {
  let state = {
    normalHeroModeEnabled: false,
    manualSignupPhoneSmsEnabled: false,
    manualAddEmailInputRequired: false,
    normalHeroEmailRuntime: true,
    signupMethod: 'email',
    resolvedSignupMethod: 'email',
    email: 'bound-hero@example.com',
    registrationEmailState: {
      current: 'bound-hero@example.com',
      previous: 'herouser@example.com',
      source: 'bind_email',
      updatedAt: 2,
    },
  };
  const router = loadMessageRouterModule().createMessageRouter({
    addLog: async () => {},
    buildXiaohongshuRuntimeReset: () => ({}),
    clearStopRequest: () => {},
    getPendingAutoRunTimerPlan: () => null,
    getState: async () => state,
    isAutoRunLockedState: () => false,
    normalizeRunCount: (value) => Math.max(1, Math.floor(Number(value) || 1)),
    getNodeIdsForState: () => [
      'open-chatgpt',
      'submit-signup-email',
      'plus-checkout-create',
      'sub2api-session-import',
    ],
    setState: async (patch) => {
      state = { ...state, ...patch };
    },
    startAutoRunLoop: () => {},
    validateAutoRunStart: () => ({ ok: true, errors: [] }),
  });

  const response = await router.handleMessage({
    type: 'AUTO_RUN_XIAOHONGSHU',
    source: 'sidepanel',
    payload: {
      accessToken: 'xiaohongshu-token',
    },
  }, {});

  assert.deepEqual(JSON.parse(JSON.stringify(response)), { ok: true });
  assert.equal(state.email, null);
  assert.equal(state.normalHeroEmailRuntime, false);
  assert.deepEqual(JSON.parse(JSON.stringify(state.registrationEmailState)), {
    current: '',
    previous: '',
    source: '',
    updatedAt: 0,
  });
});

test('Normal Hero add-email uses startup email and only prompts after email is already used', async () => {
  const calls = [];
  let promptCount = 0;
  const executor = loadStep8Module().createStep8Executor({
    ...createBaseDeps(calls),
    getState: async () => ({
      normalHeroModeEnabled: true,
      signupPhoneNumber: '+57 (324) 132 10 49',
      email: 'used@example.com',
      registrationEmailState: {
        current: 'used@example.com',
        previous: 'used@example.com',
        source: 'normal_hero_start',
        updatedAt: 1,
      },
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
      return 'fresh@example.com';
    },
    persistRegistrationEmailState: async (_state, email, options) => calls.push(['persistEmail', email, options]),
    completeNodeFromBackground: async (nodeId, payload) => calls.push(['complete', nodeId, payload]),
  });

  await executor.executeBindEmail({
    normalHeroModeEnabled: true,
    signupPhoneNumber: '+57 (324) 132 10 49',
    email: 'used@example.com',
    registrationEmailState: {
      current: 'used@example.com',
      previous: 'used@example.com',
      source: 'normal_hero_start',
      updatedAt: 1,
    },
    oauthUrl: 'https://example.test/oauth',
    nodeId: 'bind-email',
    visibleStep: 9,
  });

  assert.equal(promptCount, 1);
  assert.deepEqual(
    calls.filter((item) => item[0] === 'content' && item[1] === 'SUBMIT_ADD_EMAIL').map((item) => item[2].email),
    ['used@example.com', 'fresh@example.com']
  );
  assert.equal(calls.find((item) => item[0] === 'persistEmail')?.[2]?.normalHeroEmailRuntime, true);
  assert.deepEqual(JSON.parse(JSON.stringify(calls.find((item) => item[0] === 'complete'))), [
    'complete',
    'bind-email',
    {
      bindEmailSubmitted: true,
      email: 'fresh@example.com',
      step8VerificationTargetEmail: 'fresh@example.com',
      manualAddEmailInputRequired: true,
      normalHeroEmailRuntime: true,
    },
  ]);
});

test('Normal Hero add-email prompts manual email when startup email is missing', async () => {
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
      return {
        displayedEmail: message.payload.email,
        url: 'https://auth.openai.com/verify',
      };
    },
    requestManualAddEmailInput: async () => {
      promptCount += 1;
      return 'fallback@example.com';
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

  assert.equal(promptCount, 1);
  assert.deepEqual(
    calls.filter((item) => item[0] === 'content' && item[1] === 'SUBMIT_ADD_EMAIL').map((item) => item[2].email),
    ['fallback@example.com']
  );
  assert.equal(calls.find((item) => item[0] === 'complete')[2].email, 'fallback@example.com');
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

// content/paypal-flow.js — PayPal login and approval helper.

console.log('[MultiPage:paypal-flow] Content script loaded on', location.href);

const PAYPAL_FLOW_LISTENER_SENTINEL = 'data-multipage-paypal-flow-listener';
const PAYPAL_HOSTED_STAGE_OUTSIDE = 'outside_paypal';
const PAYPAL_HOSTED_STAGE_LOGIN = 'pay_login';
const PAYPAL_HOSTED_STAGE_ACCOUNT_CREATE_EMAIL = 'account_create_email';
const PAYPAL_HOSTED_STAGE_GUEST_CHECKOUT = 'guest_checkout';
const PAYPAL_HOSTED_STAGE_VERIFICATION = 'verification';
const PAYPAL_HOSTED_STAGE_REVIEW = 'review_consent';
const PAYPAL_HOSTED_STAGE_APPROVAL = 'approval';
const PAYPAL_HOSTED_STAGE_GENERIC_ERROR = 'generic_error';
const PAYPAL_HOSTED_STAGE_UNKNOWN = 'unknown';
const PAYPAL_HOSTED_HERMES_AUTORUN_SENTINEL = 'data-multipage-paypal-hosted-hermes-autorun';
const PAYPAL_HOSTED_GUEST_SUBMIT_SENTINEL = '__MULTIPAGE_PAYPAL_HOSTED_GUEST_SUBMIT__';
const PAYPAL_HOSTED_COUNTRY_TARGET_CODE = 'US';
const PAYPAL_HOSTED_COUNTRY_AUTORUN_SENTINEL = '__MULTIPAGE_PAYPAL_HOSTED_US_COUNTRY_AUTORUN__';
const PAYPAL_HOSTED_COUNTRY_IN_PROGRESS_SENTINEL = '__MULTIPAGE_PAYPAL_HOSTED_US_COUNTRY_IN_PROGRESS__';

if (document.documentElement.getAttribute(PAYPAL_FLOW_LISTENER_SENTINEL) !== '1') {
  document.documentElement.setAttribute(PAYPAL_FLOW_LISTENER_SENTINEL, '1');

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (
      message.type === 'PAYPAL_GET_STATE'
      || message.type === 'PAYPAL_SUBMIT_LOGIN'
      || message.type === 'PAYPAL_DISMISS_PROMPTS'
      || message.type === 'PAYPAL_CLICK_APPROVE'
      || message.type === 'PAYPAL_ENSURE_US_COUNTRY'
      || message.type === 'PAYPAL_HOSTED_GET_STATE'
      || message.type === 'PAYPAL_RUN_HOSTED_CHECKOUT_STEP'
    ) {
      resetStopState();
      handlePayPalCommand(message).then((result) => {
        sendResponse({ ok: true, ...(result || {}) });
      }).catch((err) => {
        if (isStopError(err)) {
          sendResponse({ stopped: true, error: err.message });
          return;
        }
        sendResponse({ error: err.message });
      });
      return true;
    }
  });
} else {
  console.log('[MultiPage:paypal-flow] 消息监听已存在，跳过重复注册');
}

async function performPayPalOperationWithDelay(metadata, operation) {
  const rootScope = typeof window !== 'undefined' ? window : globalThis;
  const gate = rootScope?.CodexOperationDelay?.performOperationWithDelay;
  return typeof gate === 'function' ? gate(metadata, operation) : operation();
}

async function handlePayPalCommand(message) {
  switch (message.type) {
    case 'PAYPAL_GET_STATE':
      return inspectPayPalState();
    case 'PAYPAL_SUBMIT_LOGIN':
      return submitPayPalLogin(message.payload || {});
    case 'PAYPAL_DISMISS_PROMPTS':
      return dismissPayPalPrompts();
    case 'PAYPAL_CLICK_APPROVE':
      return clickPayPalApprove();
    case 'PAYPAL_ENSURE_US_COUNTRY':
      return ensureHostedCountryUnitedStates(message.payload || {});
    case 'PAYPAL_HOSTED_GET_STATE':
      return inspectPayPalState();
    case 'PAYPAL_RUN_HOSTED_CHECKOUT_STEP':
      return runHostedCheckoutStep(message.payload || {});
    default:
      throw new Error(`paypal-flow.js 不处理消息：${message.type}`);
  }
}

async function waitUntil(predicate, options = {}) {
  const intervalMs = Math.max(50, Math.floor(Number(options.intervalMs) || 250));
  const timeoutMs = Math.max(0, Math.floor(Number(options.timeoutMs) || 0));
  const startedAt = Date.now();
  while (true) {
    throwIfStopped();
    const value = await predicate();
    if (value) {
      return value;
    }
    if (timeoutMs > 0 && Date.now() - startedAt >= timeoutMs) {
      throw new Error(options.timeoutMessage || 'PayPal page timed out waiting for target state.');
    }
    await sleep(intervalMs);
  }
}

async function waitForDocumentComplete() {
  await waitUntil(() => document.readyState === 'complete', { intervalMs: 200 });
  await sleep(1000);
}

function isVisibleElement(el) {
  if (!el) return false;
  let node = el;
  while (node && node.nodeType === 1) {
    if (node.hidden || node.getAttribute?.('aria-hidden') === 'true' || node.getAttribute?.('inert') !== null) {
      return false;
    }
    const nodeStyle = window.getComputedStyle(node);
    if (
      nodeStyle.display === 'none'
      || nodeStyle.visibility === 'hidden'
      || nodeStyle.visibility === 'collapse'
      || Number(nodeStyle.opacity) === 0
    ) {
      return false;
    }
    node = node.parentElement;
  }
  const style = window.getComputedStyle(el);
  const rect = el.getBoundingClientRect();
  return style.display !== 'none'
    && style.visibility !== 'hidden'
    && Number(rect.width) > 0
    && Number(rect.height) > 0;
}

function normalizeText(text = '') {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function getActionText(el) {
  return normalizeText([
    el?.textContent,
    el?.value,
    el?.getAttribute?.('aria-label'),
    el?.getAttribute?.('title'),
    el?.getAttribute?.('placeholder'),
    el?.getAttribute?.('name'),
    el?.id,
  ].filter(Boolean).join(' '));
}

function getVisibleControls(selector) {
  return Array.from(document.querySelectorAll(selector)).filter(isVisibleElement);
}

function isEnabledControl(el) {
  return Boolean(el)
    && !el.disabled
    && !el.matches?.(':disabled')
    && el.getAttribute?.('aria-disabled') !== 'true'
    && !el.closest?.('[aria-disabled="true"], fieldset[disabled]');
}

function findClickableByText(patterns) {
  const normalizedPatterns = (Array.isArray(patterns) ? patterns : [patterns]).filter(Boolean);
  const candidates = getVisibleControls('button, a, [role="button"], input[type="button"], input[type="submit"]');
  return candidates.find((el) => {
    const text = getActionText(el);
    return normalizedPatterns.some((pattern) => pattern.test(text));
  }) || null;
}

function findInputByPatterns(patterns) {
  const inputs = getVisibleControls('input')
    .filter((input) => {
      const type = String(input.getAttribute('type') || input.type || '').trim().toLowerCase();
      return isEnabledControl(input) && !['hidden', 'checkbox', 'radio', 'submit', 'button', 'file'].includes(type);
    });
  return inputs.find((input) => {
    const text = getActionText(input);
    return patterns.some((pattern) => pattern.test(text));
  }) || null;
}

function findEmailInput() {
  const isPasswordCandidate = (input) => {
    const type = String(input?.getAttribute?.('type') || input?.type || '').trim().toLowerCase();
    const metadataText = normalizeText([
      input?.textContent,
      input?.getAttribute?.('aria-label'),
      input?.getAttribute?.('title'),
      input?.getAttribute?.('placeholder'),
      input?.getAttribute?.('name'),
      input?.id,
    ].filter(Boolean).join(' '));
    return type === 'password' || /password|pass|密码/i.test(metadataText);
  };
  const inputs = getVisibleControls('input')
    .filter((input) => {
      const type = String(input.getAttribute('type') || input.type || '').trim().toLowerCase();
      return isEnabledControl(input)
        && !['hidden', 'checkbox', 'radio', 'submit', 'button', 'file'].includes(type)
        && !isPasswordCandidate(input);
    });
  return inputs.find((input) => [
    /email|login|user|账号|邮箱/i,
  ].some((pattern) => pattern.test(getActionText(input))))
    || getVisibleControls('input[type="email"]').find((input) => isVisibleElement(input) && !isPasswordCandidate(input))
    || null;
}

function findPasswordInput() {
  const inputs = getVisibleControls('input')
    .filter((input) => {
      const type = String(input.getAttribute('type') || input.type || '').trim().toLowerCase();
      return isEnabledControl(input) && !['hidden', 'checkbox', 'radio', 'submit', 'button', 'file'].includes(type);
    });
  return inputs.find((input) => {
    const type = String(input.getAttribute('type') || input.type || '').trim().toLowerCase();
    const metadataText = normalizeText([
      input?.textContent,
      input?.getAttribute?.('aria-label'),
      input?.getAttribute?.('title'),
      input?.getAttribute?.('placeholder'),
      input?.getAttribute?.('name'),
      input?.id,
    ].filter(Boolean).join(' '));
    return type === 'password' || /password|pass|密码/i.test(metadataText);
  }) || getVisibleControls('input[type="password"]').find(isVisibleElement) || null;
}

function findLoginNextButton() {
  return findClickableByText([
    /next|continue|login|log\s*in|sign\s*in/i,
    /下一步|继续|登录|登入/i,
  ]);
}

function findEmailNextButton() {
  return findClickableByText([
    /next|btn\s*next|btnnext/i,
    /下一页|下一步/i,
  ]);
}

function findPasswordLoginButton() {
  const button = findClickableByText([
    /login|log\s*in|sign\s*in/i,
    /登录|登入/i,
  ]);
  return button && button !== findEmailNextButton() ? button : null;
}

function findApproveButton() {
  return findClickableByText([
    /同意并继续|同意|继续|授权|确认并继续/i,
    /agree\s*(?:and)?\s*continue|continue|accept|authorize|agree|pay\s*now/i,
  ]);
}

function getPayPalHostedPathname() {
  return String(location?.pathname || '').trim();
}

function isPayPalHostedLoginPage() {
  const pathname = getPayPalHostedPathname();
  return pathname === '/pay'
    || pathname === '/pay/'
    || Boolean(document.getElementById('email'));
}

function findHostedAccountCreateEmailContinueButton() {
  return findClickableByText([
    /continue\s+(?:to\s+)?pay(?:ment)?/i,
    /继续付款|继续支付/i,
  ]);
}

function isPayPalHostedAccountCreateEmailPage() {
  const bodyText = normalizeText(document.body?.innerText || '');
  const emailInput = document.getElementById('email') || findEmailInput();
  const hasCardOrAddressForm = hasPayPalHostedGuestCheckoutForm();
  return Boolean(emailInput)
    && !findPasswordInput()
    && !hasCardOrAddressForm
    && Boolean(findHostedAccountCreateEmailContinueButton())
    && (
      /创建\s*PayPal\s*账户|create\s+(?:a\s+)?paypal\s+account/i.test(bodyText)
      || /您已有账号了吗|already\s+have\s+an?\s+account/i.test(bodyText)
    );
}

function hasPayPalHostedGuestCheckoutForm() {
  return ['cardNumber', 'billingLine1', 'cardExpiry', 'cardCvv']
    .some((id) => isVisibleElement(document.getElementById(id)));
}

function isPayPalHostedGuestCheckoutPage() {
  const pathname = getPayPalHostedPathname();
  return hasPayPalHostedGuestCheckoutForm()
    || (/\/checkoutweb\//i.test(pathname) && !/\/checkoutweb\/genericError/i.test(pathname));
}

function getPayPalHostedGenericErrorMessage() {
  const bodyText = normalizeText(document.body?.innerText || '');
  const match = bodyText.match(
    /Things\s+don[’']?t\s+appear\s+to\s+be\s+working\s+at\s+the\s+moment\.?|Sorry,\s*something\s+went\s+wrong\.?\s*Please\s+try\s+again\.?|Something\s+went\s+wrong(?:\.?\s*Please\s+go\s+back\s+to\s+[^.]+?\s+and\s+choose\s+another\s+way\s+to\s+pay\.?\s*PayPal\s+isn[’']?t\s+available\s+at\s+this\s+time\.?)?/i
  );
  return match ? match[0] : '';
}

function getPayPalHostedCardDeclinedMessage() {
  const bodyText = normalizeText(document.body?.innerText || '');
  const match = bodyText.match(
    /We\s+weren[’']?t\s+able\s+to\s+add\s+this\s+card[\s\S]{0,220}?(?:try\s+again|try\s+a\s+different\s+card|different\s+card)|Check\s+all\s+the\s+details\s+are\s+correct[\s\S]{0,160}?try\s+a\s+different\s+card|无法添加(?:这|此|该)?张?卡|无法添加(?:这|此|该)?付款卡|不能添加(?:这|此|该)?张?卡|请检查[\s\S]{0,80}?(?:换|尝试)[\s\S]{0,40}?卡/i
  );
  return match ? normalizeText(match[0]) : '';
}

function hasPayPalHostedCardDeclinedError() {
  return Boolean(getPayPalHostedCardDeclinedMessage());
}

function isPayPalHostedGenericErrorPage() {
  const pathname = getPayPalHostedPathname();
  const bodyText = normalizeText(document.body?.innerText || '');
  return /\/checkoutweb\/genericError/i.test(pathname)
    || Boolean(getPayPalHostedGenericErrorMessage())
    || (
      /(?:sorry,\s*)?something\s+went\s+wrong/i.test(bodyText)
      && /return\s+to\s+merchant/i.test(bodyText)
    )
    || (
      /paypal\s+isn[’']?t\s+available\s+at\s+this\s+time/i.test(bodyText)
      && /choose\s+another\s+way\s+to\s+pay/i.test(bodyText)
    );
}

function isPayPalHostedReviewPage() {
  return /\/webapps\/hermes/i.test(getPayPalHostedPathname());
}

function findHostedVerificationInputs() {
  return Array.from({ length: 6 }, (_, index) => document.getElementById(`ci-ciBasic-${index}`))
    .filter((input) => isVisibleElement(input));
}

function hasHostedVerificationInputs() {
  return findHostedVerificationInputs().length >= 6;
}

function getHostedVerificationErrorText() {
  const errorPattern = /check\s+the\s+code\s+and\s+try\s+again|(?:sorry,\s*)?something\s+went\s+wrong\.?\s*get\s+a\s+new\s+code|get\s+a\s+new\s+code/i;
  const alert = document.getElementById('message_ciBasic')
    || getVisibleControls('[role="alert"]').find((node) => errorPattern.test(normalizeText(node.textContent || '')));
  return alert && isVisibleElement(alert) ? normalizeText(alert.textContent || '') : '';
}

function hasHostedInvalidVerificationCodeError() {
  return /check\s+the\s+code\s+and\s+try\s+again|(?:sorry,\s*)?something\s+went\s+wrong\.?\s*get\s+a\s+new\s+code|get\s+a\s+new\s+code/i.test(getHostedVerificationErrorText());
}

function findHostedVerificationResendButton() {
  const direct = document.querySelector(
    [
      'button[data-testid="resend-link"]',
      'a[data-testid="link-get-new-code"]',
      'button[data-testid="link-get-new-code"]',
      '#linkGetNewCode',
      '#link-get-new-code',
    ].join(', ')
  );
  if (direct && isVisibleElement(direct) && isEnabledControl(direct)) {
    return direct;
  }
  const patterns = [
    /resend/i,
    /重新发送|重发/i,
  ];
  const candidates = getVisibleControls('button, [role="button"], a, input[type="button"], input[type="submit"]');
  return candidates.find((el) => {
    const text = getActionText(el);
    return isEnabledControl(el) && patterns.some((pattern) => pattern.test(text));
  }) || null;
}

function findHostedReviewConsentButton() {
  const direct = document.getElementById('consentButton')
    || document.querySelector('button[data-testid="consentButton"], button[name="agreeAndContinueQL"]');
  if (direct && isVisibleElement(direct) && isEnabledControl(direct)) {
    return direct;
  }
  const patterns = [
    /agree\s*(?:&|and)?\s*continue|agreeAndContinueQL|agreeContinue/i,
    /同意并继续|同意.*继续|接受并继续/i,
  ];
  const candidates = getVisibleControls('button, [role="button"], input[type="button"], input[type="submit"]');
  return candidates.find((el) => {
    const text = getActionText(el);
    return isEnabledControl(el) && patterns.some((pattern) => pattern.test(text));
  }) || null;
}

function detectPayPalHostedCheckoutStage() {
  if (!/paypal\./i.test(String(location?.host || ''))) {
    return PAYPAL_HOSTED_STAGE_OUTSIDE;
  }
  if (hasHostedVerificationInputs()) {
    return PAYPAL_HOSTED_STAGE_VERIFICATION;
  }
  if (isPayPalHostedGenericErrorPage()) {
    return PAYPAL_HOSTED_STAGE_GENERIC_ERROR;
  }
  if (isPayPalHostedAccountCreateEmailPage()) {
    return PAYPAL_HOSTED_STAGE_ACCOUNT_CREATE_EMAIL;
  }
  if (isPayPalHostedGuestCheckoutPage()) {
    return PAYPAL_HOSTED_STAGE_GUEST_CHECKOUT;
  }
  if (isPayPalHostedReviewPage() && findHostedReviewConsentButton()) {
    return PAYPAL_HOSTED_STAGE_REVIEW;
  }
  if (isPayPalHostedLoginPage()) {
    return PAYPAL_HOSTED_STAGE_LOGIN;
  }
  if (Boolean(findApproveButton())) {
    return PAYPAL_HOSTED_STAGE_APPROVAL;
  }
  return PAYPAL_HOSTED_STAGE_UNKNOWN;
}

function fillHostedInputById(id, value) {
  const input = document.getElementById(String(id || '').trim());
  if (!input || !isVisibleElement(input) || !isEnabledControl(input)) {
    return false;
  }
  fillInput(input, String(value || ''));
  return true;
}

function selectHostedOptionByIdText(id, text) {
  const select = document.getElementById(String(id || '').trim());
  const expectedText = normalizeText(text);
  if (!select || !expectedText || !Array.isArray(Array.from(select.options || []))) {
    return false;
  }
  const match = Array.from(select.options || []).find((option) => {
    const label = normalizeText(option?.textContent || option?.label || '');
    const value = normalizeText(option?.value || '');
    return label.toLowerCase().includes(expectedText.toLowerCase())
      || value.toLowerCase().includes(expectedText.toLowerCase());
  });
  if (!match) {
    return false;
  }
  select.value = match.value;
  select.dispatchEvent(new Event('change', { bubbles: true }));
  return true;
}

function getHostedCountrySelectTrigger() {
  const controls = getVisibleControls('button, [role="combobox"], [aria-controls="country-select-sheet"]');
  return controls.find((control) => {
    const metadataText = getActionText(control);
    return control.getAttribute?.('aria-controls') === 'country-select-sheet'
      || /country\s*(?:or\s*region|select)|国家|地区/i.test(metadataText);
  }) || null;
}

function getHostedCountryHiddenSelect() {
  const container = document.querySelector('[data-testid="language-selector-container"]');
  const selects = [
    ...(container ? Array.from(container.querySelectorAll('select')) : []),
    ...Array.from(document.querySelectorAll('select')),
  ];
  return selects.find((select) => {
    const text = getActionText(select);
    return /country|region|国家|地区/i.test(text)
      || select.closest?.('[data-testid="language-selector-container"]');
  }) || null;
}

function getHostedCountrySelectionText() {
  const trigger = getHostedCountrySelectTrigger();
  const select = getHostedCountryHiddenSelect();
  const selectedOption = select?.selectedOptions?.[0];
  return normalizeText([
    trigger?.textContent,
    trigger?.getAttribute?.('aria-label'),
    trigger?.getAttribute?.('title'),
    select?.value,
    selectedOption?.value,
    selectedOption?.label,
    selectedOption?.textContent,
  ].filter(Boolean).join(' '));
}

function getHostedCountryServerDefaultCode() {
  const html = String(document.documentElement?.innerHTML || '');
  const patterns = [
    /"defaultValue"\s*:\s*"([A-Z0-9]{2})"/,
    /\\"defaultValue\\"\s*:\s*\\"([A-Z0-9]{2})\\"/,
    /"ccpg=([A-Z]{2})(?:\\u0026|&)/,
    /ccpg=([A-Z]{2})(?:\\u0026|&)/,
  ];
  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match?.[1]) {
      return normalizeText(match[1]).toUpperCase();
    }
  }
  return '';
}

function isHostedCountryUnitedStates() {
  const urlCountry = normalizeText(new URLSearchParams(location.search || '').get('country.x') || '').toUpperCase();
  const serverDefault = getHostedCountryServerDefaultCode();
  if (serverDefault) {
    return serverDefault === PAYPAL_HOSTED_COUNTRY_TARGET_CODE;
  }
  const text = getHostedCountrySelectionText();
  if (/united\s*states|美国|\busa?\b/i.test(text)) {
    return true;
  }
  if (urlCountry === PAYPAL_HOSTED_COUNTRY_TARGET_CODE && !text) {
    return false;
  }
  if (urlCountry && urlCountry !== PAYPAL_HOSTED_COUNTRY_TARGET_CODE) {
    return false;
  }
  const trigger = getHostedCountrySelectTrigger();
  const flagNode = trigger?.querySelector?.('[style*="flags/2x.png"], span[style*="paypal-ui/components/flags"]');
  const styleText = String(flagNode?.getAttribute?.('style') || '');
  return /\bUS\b/i.test(trigger?.getAttribute?.('data-value') || '')
    || /United\s*States/i.test(trigger?.getAttribute?.('data-label') || '')
    || /25\.862%|26\.437%/.test(styleText);
}

function getHostedCountryContextId() {
  const ctxIdInput = document.querySelector('input[name="ctxId"]');
  const inputValue = normalizeText(ctxIdInput?.value || '');
  if (inputValue) {
    return inputValue;
  }
  const match = String(document.documentElement?.innerHTML || '').match(/"ctxId":"([^"]+)"/);
  if (match?.[1]) {
    return normalizeText(match[1]);
  }
  const escapedMatch = String(document.documentElement?.innerHTML || '').match(/\\"ctxId\\"\s*:\s*\\"([^"\\]+)\\"/);
  return normalizeText(escapedMatch?.[1] || '');
}

function buildHostedCountryUrl(targetCountryCode = PAYPAL_HOSTED_COUNTRY_TARGET_CODE) {
  if (!/paypal\./i.test(String(location?.host || ''))) {
    return '';
  }
  if (!/^\/pay\/?$/i.test(getPayPalHostedPathname())) {
    return '';
  }
  const target = normalizeText(targetCountryCode).toUpperCase() || PAYPAL_HOSTED_COUNTRY_TARGET_CODE;
  const url = new URL(location.href);
  url.pathname = '/pay/';
  url.searchParams.set('country.x', target);
  const ctxId = getHostedCountryContextId();
  if (ctxId && !url.searchParams.get('ctxId')) {
    url.searchParams.set('ctxId', ctxId);
  }
  return url.href;
}

function shouldNavigateHostedCountryUrl(targetCountryCode = PAYPAL_HOSTED_COUNTRY_TARGET_CODE) {
  const target = normalizeText(targetCountryCode).toUpperCase() || PAYPAL_HOSTED_COUNTRY_TARGET_CODE;
  const current = normalizeText(new URLSearchParams(location.search || '').get('country.x') || '').toUpperCase();
  return current !== target && Boolean(buildHostedCountryUrl(target));
}

function navigateHostedCountryUrl(targetCountryCode = PAYPAL_HOSTED_COUNTRY_TARGET_CODE) {
  const targetUrl = buildHostedCountryUrl(targetCountryCode);
  if (!targetUrl || targetUrl === location.href) {
    return false;
  }
  setTimeout(() => {
    window.location.assign(targetUrl);
  }, 50);
  return true;
}

function isHostedCountryOptionCandidate(el) {
  if (!el || !isVisibleElement(el) || !isEnabledControl(el)) {
    return false;
  }
  const text = getActionText(el);
  const value = normalizeText(el.getAttribute?.('value') || el.getAttribute?.('data-value') || '');
  return /united\s*states|美国/i.test(text)
    || /^US$/i.test(value);
}

function findHostedCountryUnitedStatesOption() {
  const selectors = [
    '#country-select-sheet button',
    '#country-select-sheet [role="option"]',
    '#country-select-sheet [role="menuitem"]',
    '#country-select-sheet li',
    '[role="dialog"] button',
    '[role="dialog"] [role="option"]',
    '[role="listbox"] [role="option"]',
    '[role="menu"] [role="menuitem"]',
    'button',
    '[role="option"]',
  ];
  const candidates = selectors.flatMap((selector) => Array.from(document.querySelectorAll(selector)));
  return candidates.find(isHostedCountryOptionCandidate) || null;
}

function findHostedCountrySearchInput() {
  const containers = [
    document.getElementById('country-select-sheet'),
    ...Array.from(document.querySelectorAll('[role="dialog"], [role="listbox"], [role="menu"]')),
  ].filter(Boolean);
  const inputs = containers.flatMap((container) => Array.from(container.querySelectorAll('input')));
  return inputs.find((input) => {
    const type = String(input.getAttribute('type') || input.type || '').trim().toLowerCase();
    return isVisibleElement(input)
      && isEnabledControl(input)
      && !['hidden', 'checkbox', 'radio', 'submit', 'button', 'file'].includes(type);
  }) || null;
}

function scrollHostedCountryContainers() {
  const containers = [
    document.getElementById('country-select-sheet'),
    ...Array.from(document.querySelectorAll('[role="dialog"], [role="listbox"], [role="menu"]')),
    ...Array.from(document.querySelectorAll('div, ul')).filter((node) => {
      if (!isVisibleElement(node)) return false;
      const style = window.getComputedStyle(node);
      return /(auto|scroll)/i.test(`${style.overflowY} ${style.overflow}`);
    }),
  ].filter(Boolean);
  containers.forEach((container) => {
    try {
      container.scrollTop = Math.min(container.scrollTop + 600, container.scrollHeight);
      container.dispatchEvent(new Event('scroll', { bubbles: true }));
    } catch {
      // Non-scrollable containers can be ignored.
    }
  });
}

async function selectHostedCountryViaNativeSelect(targetCountryCode) {
  const select = getHostedCountryHiddenSelect();
  if (!select || !Array.from(select.options || []).length) {
    return false;
  }
  const target = String(targetCountryCode || PAYPAL_HOSTED_COUNTRY_TARGET_CODE).trim().toUpperCase();
  const option = Array.from(select.options || []).find((item) => {
    const value = normalizeText(item.value || '');
    const label = normalizeText(item.label || item.textContent || '');
    return value.toUpperCase() === target
      || /united\s*states|美国/i.test(label)
      || /^US$/i.test(label);
  });
  if (!option) {
    return false;
  }
  select.value = option.value;
  select.dispatchEvent(new Event('input', { bubbles: true }));
  select.dispatchEvent(new Event('change', { bubbles: true }));
  await sleep(800);
  return isHostedCountryUnitedStates();
}

async function ensureHostedCountryUnitedStates(options = {}) {
  await waitForDocumentComplete();
  if (!/paypal\./i.test(String(location?.host || ''))) {
    return {
      attempted: false,
      changed: false,
      reason: 'not_paypal',
    };
  }
  const targetCountryCode = normalizeText(options.countryCode || PAYPAL_HOSTED_COUNTRY_TARGET_CODE).toUpperCase() || PAYPAL_HOSTED_COUNTRY_TARGET_CODE;
  const forceUiFallback = Boolean(options.forceUiFallback);
  if (isHostedCountryUnitedStates()) {
    return {
      attempted: false,
      changed: false,
      alreadySelected: true,
      countryText: getHostedCountrySelectionText(),
    };
  }

  if (!forceUiFallback && shouldNavigateHostedCountryUrl(targetCountryCode)) {
    const targetUrl = buildHostedCountryUrl(targetCountryCode);
    const navigationStarted = navigateHostedCountryUrl(targetCountryCode);
    return {
      attempted: true,
      changed: Boolean(navigationStarted),
      method: 'country_url',
      navigationStarted,
      targetUrl,
      selected: false,
      countryText: getHostedCountrySelectionText(),
    };
  }

  const changedBySelect = await selectHostedCountryViaNativeSelect(targetCountryCode);
  if (changedBySelect) {
    await sleep(1200);
    return {
      attempted: true,
      changed: true,
      method: 'native_select',
      selected: isHostedCountryUnitedStates(),
      countryText: getHostedCountrySelectionText(),
    };
  }

  const trigger = getHostedCountrySelectTrigger();
  if (!trigger) {
    return {
      attempted: false,
      changed: false,
      reason: 'country_trigger_not_found',
      countryText: getHostedCountrySelectionText(),
    };
  }

  simulateClick(trigger);
  await sleep(500);
  const searchInput = findHostedCountrySearchInput();
  if (searchInput) {
    fillInput(searchInput, 'United States');
    await sleep(500);
  }
  const option = await waitUntil(() => {
    const direct = findHostedCountryUnitedStatesOption();
    if (direct) return direct;
    scrollHostedCountryContainers();
    return null;
  }, {
    intervalMs: 250,
    timeoutMs: Number(options.timeoutMs) || 8000,
    timeoutMessage: 'PayPal 页面底部国家选择器未找到 United States 选项。',
  });
  await sleep(250);
  simulateClick(option);
  await sleep(2000);

  return {
    attempted: true,
    changed: true,
    method: 'country_sheet',
    selected: isHostedCountryUnitedStates(),
    countryText: getHostedCountrySelectionText(),
  };
}

function removeHostedCaptchaArtifacts() {
  let removed = false;
  const selectors = [
    '#captcha-standalone',
    '.captcha-overlay',
    '.captcha-container',
  ];
  selectors.forEach((selector) => {
    document.querySelectorAll(selector).forEach((node) => {
      try {
        node.remove();
        removed = true;
      } catch {
        // Ignore non-removable overlays.
      }
    });
  });
  return removed;
}

function startHostedCaptchaCleanupObserver(timeoutMs = 15000) {
  const observer = new MutationObserver(() => {
    removeHostedCaptchaArtifacts();
  });
  observer.observe(document.documentElement || document.body, {
    childList: true,
    subtree: true,
  });
  setTimeout(() => observer.disconnect(), Math.max(1000, Number(timeoutMs) || 15000));
  return observer;
}

function findHostedGuestSubmitButton() {
  return document.querySelector('button[data-testid="submit-button"]')
    || document.querySelector('button[data-testid="hosted-payment-submit-button"]')
    || document.querySelector('button[data-atomic-wait-intent="Submit_Email"]')
    || document.querySelector('button.SubmitButton--complete')
    || findClickableByText([
      /pay|continue|next|agree|subscribe/i,
      /支付|继续|下一步|同意|订阅/i,
    ]);
}

function buildHostedRandomEmail() {
  const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let value = '';
  for (let index = 0; index < 16; index += 1) {
    value += alphabet[Math.floor(Math.random() * alphabet.length)];
  }
  return `${value}@gmail.com`;
}

const HOSTED_PAYPAL_PASSWORD_MIN_LENGTH = 8;
const HOSTED_PAYPAL_PASSWORD_MAX_LENGTH = 20;
const HOSTED_PAYPAL_PASSWORD_SYMBOLS = '!@#$%^';
const HOSTED_PAYPAL_PASSWORD_SEQUENCE_ROWS = [
  '0123456789',
  '!@#$%^',
  'abcdefghijklmnopqrstuvwxyz',
  'qwertyuiop',
  'asdfghjkl',
  'zxcvbnm',
];

function hasHostedPayPalPasswordConsecutiveKeys(password = '') {
  const value = String(password || '').toLowerCase();
  for (let index = 0; index <= value.length - 4; index += 1) {
    const chunk = value.slice(index, index + 4);
    if (/^(.)\1{3}$/.test(chunk)) {
      return true;
    }
    if (HOSTED_PAYPAL_PASSWORD_SEQUENCE_ROWS.some((row) => {
      const reversed = row.split('').reverse().join('');
      return row.includes(chunk) || reversed.includes(chunk);
    })) {
      return true;
    }
  }
  return false;
}

function isHostedPayPalPasswordCompliant(password = '') {
  const value = String(password || '');
  return value.length >= HOSTED_PAYPAL_PASSWORD_MIN_LENGTH
    && value.length <= HOSTED_PAYPAL_PASSWORD_MAX_LENGTH
    && new RegExp(`[0-9${HOSTED_PAYPAL_PASSWORD_SYMBOLS.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}]`).test(value)
    && !hasHostedPayPalPasswordConsecutiveKeys(value);
}

function shuffleHostedPasswordChars(chars) {
  const values = chars.slice();
  for (let index = values.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [values[index], values[swapIndex]] = [values[swapIndex], values[index]];
  }
  return values;
}

function buildHostedRandomPassword() {
  const lowercase = 'abcdefghijklmnopqrstuvwxyz';
  const uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const digits = '0123456789';
  const alphabet = `${lowercase}${uppercase}${digits}${HOSTED_PAYPAL_PASSWORD_SYMBOLS}`;
  for (let attempt = 0; attempt < 100; attempt += 1) {
    const value = [
      lowercase[Math.floor(Math.random() * lowercase.length)],
      uppercase[Math.floor(Math.random() * uppercase.length)],
      digits[Math.floor(Math.random() * digits.length)],
      HOSTED_PAYPAL_PASSWORD_SYMBOLS[Math.floor(Math.random() * HOSTED_PAYPAL_PASSWORD_SYMBOLS.length)],
    ];
    while (value.length < 14) {
      value.push(alphabet[Math.floor(Math.random() * alphabet.length)]);
    }
    const password = shuffleHostedPasswordChars(value).join('');
    if (isHostedPayPalPasswordCompliant(password)) {
      return password;
    }
  }
  return 'Pa9!Vx2@Lm7#Qr';
}

function buildHostedVisaCard() {
  const prefixes = [
    [4, 1, 4, 7],
    [4, 1, 0, 0],
  ];
  const digits = prefixes[Math.floor(Math.random() * prefixes.length)].slice();
  while (digits.length < 15) {
    digits.push(Math.floor(Math.random() * 10));
  }
  const reversed = digits.slice().reverse();
  let sum = 0;
  for (let index = 0; index < reversed.length; index += 1) {
    let digit = reversed[index];
    if (index % 2 === 0) {
      digit *= 2;
      if (digit > 9) {
        digit -= 9;
      }
    }
    sum += digit;
  }
  const checkDigit = (10 - (sum % 10)) % 10;
  digits.push(checkDigit);
  const month = String(Math.floor(Math.random() * 12) + 1).padStart(2, '0');
  const currentYear = new Date().getFullYear() % 100;
  const year = currentYear + Math.floor(Math.random() * 4) + 2;
  const cvv = String(Math.floor(100 + Math.random() * 900));
  return {
    number: digits.join(''),
    expiry: `${month} / ${year}`,
    cvv,
  };
}

function dispatchHostedGenericClick(button) {
  const rect = button.getBoundingClientRect();
  const clientX = rect.left + rect.width / 2;
  const clientY = rect.top + rect.height / 2;
  const eventInit = {
    bubbles: true,
    cancelable: true,
    view: window,
    clientX,
    clientY,
  };
  button.dispatchEvent(new PointerEvent('pointerdown', eventInit));
  button.dispatchEvent(new MouseEvent('mousedown', eventInit));
  button.dispatchEvent(new PointerEvent('pointerup', eventInit));
  button.dispatchEvent(new MouseEvent('mouseup', eventInit));
  button.dispatchEvent(new MouseEvent('click', eventInit));
}

async function clickHostedGenericSubmitButton(retries = 0) {
  removeHostedCaptchaArtifacts();
  const button = findHostedGuestSubmitButton() || findEmailNextButton() || findLoginNextButton();
  if (!button) {
    if (retries >= 10) {
      throw new Error('PayPal hosted checkout 未找到可点击的继续/提交按钮。');
    }
    await sleep(1000);
    return clickHostedGenericSubmitButton(retries + 1);
  }

  const buttonText = normalizeText(button.textContent || '');
  if (button.disabled) {
    if (retries >= 10) {
      throw new Error('PayPal hosted checkout 按钮长时间处于 disabled 状态。');
    }
    await sleep(1000);
    return clickHostedGenericSubmitButton(retries + 1);
  }

  const rect = button.getBoundingClientRect();
  if (rect.height === 0) {
    if (retries >= 10) {
      throw new Error('PayPal hosted checkout 按钮长时间不可见。');
    }
    await sleep(1000);
    return clickHostedGenericSubmitButton(retries + 1);
  }

  dispatchHostedGenericClick(button);
  await sleep(1000);
  removeHostedCaptchaArtifacts();

  if (hasHostedVerificationInputs()) {
    return {
      clicked: true,
      verificationRequired: true,
      buttonText,
    };
  }

  const currentText = normalizeText(button.textContent || '');
  if (!/processing/i.test(currentText) && currentText === buttonText) {
    if (retries >= 10) {
      return {
        clicked: true,
        verificationRequired: false,
        buttonText,
        retried: true,
      };
    }
    await sleep(2000);
    return clickHostedGenericSubmitButton(retries + 1);
  }

  return {
    clicked: true,
    verificationRequired: false,
    buttonText,
  };
}

function normalizeHostedVerificationCode(value = '') {
  const digits = String(value || '').replace(/\D+/g, '');
  return digits.slice(0, 6);
}

async function submitHostedPayLogin(payload = {}) {
  await waitForDocumentComplete();
  removeHostedCaptchaArtifacts();
  const email = normalizeText(payload.email || buildHostedRandomEmail());
  if (!email) {
    throw new Error('PayPal hosted checkout 缺少邮箱。');
  }
  const emailInput = document.getElementById('email') || findEmailInput();
  if (!emailInput) {
    throw new Error('PayPal hosted checkout 未找到邮箱输入框。');
  }
  await sleep(2000);
  refillPayPalEmailInput(emailInput, email);
  await sleep(1000);
  const clickResult = await clickHostedGenericSubmitButton(0);
  return {
    stage: PAYPAL_HOSTED_STAGE_LOGIN,
    submitted: true,
    generatedEmail: email,
    verificationRequired: Boolean(clickResult?.verificationRequired),
    nextExpected: 'guest_checkout_or_verification',
  };
}

async function submitHostedAccountCreateEmail(payload = {}) {
  await waitForDocumentComplete();
  removeHostedCaptchaArtifacts();
  const email = normalizeText(payload.email || buildHostedRandomEmail());
  if (!email) {
    throw new Error('PayPal 创建账户页缺少邮箱。');
  }
  const emailInput = document.getElementById('email') || findEmailInput();
  if (!emailInput) {
    throw new Error('PayPal 创建账户页未找到邮箱输入框。');
  }
  await sleep(1000);
  refillPayPalEmailInput(emailInput, email);
  await sleep(500);
  const button = findHostedAccountCreateEmailContinueButton();
  if (button && isVisibleElement(button) && isEnabledControl(button)) {
    dispatchHostedGenericClick(button);
    await sleep(1000);
    removeHostedCaptchaArtifacts();
  } else {
    await clickHostedGenericSubmitButton(0);
  }
  return {
    stage: PAYPAL_HOSTED_STAGE_ACCOUNT_CREATE_EMAIL,
    submitted: true,
    generatedEmail: email,
    nextExpected: 'guest_checkout_or_verification',
  };
}

async function fillHostedVerificationCode(payload = {}) {
  const delayOperation = typeof performPayPalOperationWithDelay === 'function'
    ? performPayPalOperationWithDelay
    : async (_metadata, operation) => operation();
  await waitForDocumentComplete();
  const code = normalizeHostedVerificationCode(payload.verificationCode || payload.code || '');
  if (code.length !== 6) {
    throw new Error('PayPal hosted checkout 验证码无效。');
  }
  const inputs = findHostedVerificationInputs();
  if (inputs.length < 6) {
    throw new Error('PayPal hosted checkout 当前页面未显示验证码输入框。');
  }
  await delayOperation({ stepKey: 'plus-checkout-create', kind: 'fill', label: 'hosted-paypal-verification-code' }, async () => {
    inputs.forEach((input, index) => {
      fillInput(input, code[index] || '');
    });
  });
  return {
    stage: PAYPAL_HOSTED_STAGE_VERIFICATION,
    codeSubmitted: true,
  };
}

async function clickHostedVerificationResend() {
  const delayOperation = typeof performPayPalOperationWithDelay === 'function'
    ? performPayPalOperationWithDelay
    : async (_metadata, operation) => operation();
  await waitForDocumentComplete();
  const button = await waitUntil(() => findHostedVerificationResendButton(), {
    intervalMs: 250,
    timeoutMs: 10000,
    timeoutMessage: 'PayPal hosted checkout 当前验证码页未找到可用的 Resend 按钮。',
  });
  await delayOperation({ stepKey: 'plus-checkout-create', kind: 'click', label: 'hosted-paypal-verification-resend' }, async () => {
    simulateClick(button);
  });
  return {
    stage: PAYPAL_HOSTED_STAGE_VERIFICATION,
    resendClicked: true,
    invalidCodeVisibleAfterClick: hasHostedInvalidVerificationCodeError(),
  };
}

async function fillHostedGuestCheckout(payload = {}) {
  await waitForDocumentComplete();
  startHostedCaptchaCleanupObserver();
  removeHostedCaptchaArtifacts();
  log(`PayPal guest checkout：收到 payload.phone=${String(payload?.phone || '').trim() || '(空)'}，payload.address=${JSON.stringify(payload?.address || {})}`, 'info');

  await sleep(2000);
  const countrySelect = document.getElementById('country');
  if (countrySelect && String(countrySelect.value || '').trim().toUpperCase() !== 'US') {
    countrySelect.value = 'US';
    countrySelect.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(3000);
  }

  const card = buildHostedVisaCard();
  const email = normalizeText(payload.email || buildHostedRandomEmail());
  const phone = normalizeText(payload.phone || '');
  const payloadPassword = String(payload.password || '');
  const password = isHostedPayPalPasswordCompliant(payloadPassword)
    ? payloadPassword
    : buildHostedRandomPassword();
  const firstName = normalizeText(payload.firstName || 'James');
  const lastName = normalizeText(payload.lastName || 'Smith');
  const cardNumber = String(payload.cardNumber || card.number).replace(/\s+/g, '');
  const cardExpiry = normalizeText(payload.cardExpiry || card.expiry);
  const cardCvv = normalizeText(payload.cardCvv || card.cvv);
  const address = payload.address && typeof payload.address === 'object' ? payload.address : {};

  if (!email || !phone || !password || !cardNumber || !cardExpiry || !cardCvv) {
    throw new Error('PayPal hosted checkout 缺少卡支付所需资料（请先填写 PayPal 电话(不带+1) 或导入 PayPal 接码池）。');
  }

  fillHostedInputById('email', email);
  fillHostedInputById('phone', phone);
  fillHostedInputById('cardNumber', cardNumber);
  fillHostedInputById('cardExpiry', cardExpiry);
  fillHostedInputById('cardCvv', cardCvv);
  fillHostedInputById('password', password);
  fillHostedInputById('firstName', firstName);
  fillHostedInputById('lastName', lastName);
  fillHostedInputById('billingLine1', address.street || '');
  fillHostedInputById('billingCity', address.city || '');
  fillHostedInputById('billingPostalCode', address.zip || '');
  fillHostedInputById('billingLine1', address.street || '');
  selectHostedOptionByIdText('billingState', address.state || '');

  const rootScope = typeof window !== 'undefined' ? window : globalThis;
  if (!rootScope[PAYPAL_HOSTED_GUEST_SUBMIT_SENTINEL]) {
    rootScope[PAYPAL_HOSTED_GUEST_SUBMIT_SENTINEL] = true;
    setTimeout(() => {
      clickHostedGenericSubmitButton(0).catch((error) => {
        log(`PayPal hosted checkout guest submit 失败：${error?.message || error}`, 'warn');
      }).finally(() => {
        rootScope[PAYPAL_HOSTED_GUEST_SUBMIT_SENTINEL] = false;
      });
    }, 500);
  }

  return {
    stage: PAYPAL_HOSTED_STAGE_GUEST_CHECKOUT,
    submitted: true,
    verificationRequired: Boolean(hasHostedVerificationInputs()),
    submitScheduled: true,
  };
}

async function clickHostedReviewConsent() {
  await waitForDocumentComplete();
  log(`PayPal Hermes：开始等待并点击同意继续按钮。当前 URL：${location.href}`, 'info');
  let waited = 0;
  while (waited < 30) {
    waited += 1;
    const button = findHostedReviewConsentButton();
    if (button) {
      const buttonText = getActionText(button);
      log(`PayPal Hermes：第 ${waited}/30 秒找到同意继续按钮，准备点击：${buttonText || button.id || button.tagName}`, 'info');
      dispatchHostedGenericClick(button);
      await sleep(1500);
      return {
        stage: PAYPAL_HOSTED_STAGE_REVIEW,
        submitted: true,
        clickedButtonText: buttonText,
        currentUrl: location.href,
        stillReviewPage: isPayPalHostedReviewPage(),
      };
    }
    if (waited === 1 || waited % 5 === 0) {
      const pageText = normalizeText(document.body?.innerText || '').slice(0, 180);
      log(`PayPal Hermes：尚未找到同意继续按钮，继续等待（${waited}/30）。页面预览：${pageText}`, 'info');
    }
    await sleep(1000);
  }
  log('PayPal Hermes：等待 30 秒后仍未找到同意继续按钮。', 'warn');
  throw new Error('PayPal hosted checkout 账单确认页超时，未检测到同意并继续按钮。');
}

async function runHostedCheckoutStep(payload = {}) {
  if (isPayPalHostedReviewPage()) {
    return clickHostedReviewConsent();
  }
  const stage = detectPayPalHostedCheckoutStage();
  if (stage === PAYPAL_HOSTED_STAGE_VERIFICATION) {
    if (payload.resendVerificationCode) {
      return clickHostedVerificationResend();
    }
    if (!payload.verificationCode && !payload.code) {
      return {
        stage,
        requiresVerificationCode: true,
      };
    }
    return fillHostedVerificationCode(payload);
  }
  if (stage === PAYPAL_HOSTED_STAGE_LOGIN) {
    return submitHostedPayLogin(payload);
  }
  if (stage === PAYPAL_HOSTED_STAGE_ACCOUNT_CREATE_EMAIL) {
    return submitHostedAccountCreateEmail(payload);
  }
  if (stage === PAYPAL_HOSTED_STAGE_GUEST_CHECKOUT) {
    return fillHostedGuestCheckout(payload);
  }
  if (stage === PAYPAL_HOSTED_STAGE_REVIEW) {
    return clickHostedReviewConsent();
  }
  return {
    stage,
    submitted: false,
    approveReady: Boolean(findApproveButton()),
  };
}

function shouldAutoRunHostedHermesReview() {
  if (!isPayPalHostedReviewPage()) {
    return false;
  }
  if (document.documentElement.getAttribute(PAYPAL_HOSTED_HERMES_AUTORUN_SENTINEL) === '1') {
    return false;
  }
  document.documentElement.setAttribute(PAYPAL_HOSTED_HERMES_AUTORUN_SENTINEL, '1');
  return true;
}

function scheduleHostedHermesAutoRun() {
  if (!shouldAutoRunHostedHermesReview()) {
    return;
  }
  log(`PayPal Hermes 页面已命中，按油猴脚本方式自动等待并点击 Agree and Continue。当前 URL：${location.href}`, 'info');
  setTimeout(() => {
    clickHostedReviewConsent().then(() => {
      log('PayPal Hermes：已按油猴脚本方式执行 Agree and Continue。', 'ok');
    }).catch((error) => {
      log(`PayPal Hermes：自动点击 Agree and Continue 失败：${error?.message || error}`, 'warn');
    });
  }, 0);
}

function shouldAutoRunHostedCountrySwitch() {
  const rootScope = typeof window !== 'undefined' ? window : globalThis;
  if (rootScope[PAYPAL_HOSTED_COUNTRY_AUTORUN_SENTINEL]) {
    return false;
  }
  if (!/paypal\./i.test(String(location?.host || ''))) {
    return false;
  }
  if (!/^\/pay\/?$/i.test(getPayPalHostedPathname())) {
    return false;
  }
  if (!isPayPalHostedLoginPage()) {
    return false;
  }
  if (!document.querySelector('[data-testid="language-selector-container"]')) {
    return false;
  }
  rootScope[PAYPAL_HOSTED_COUNTRY_AUTORUN_SENTINEL] = true;
  return true;
}

function scheduleHostedCountrySwitchAutoRun() {
  if (!shouldAutoRunHostedCountrySwitch()) {
    return;
  }
  const rootScope = typeof window !== 'undefined' ? window : globalThis;
  setTimeout(() => {
    if (rootScope[PAYPAL_HOSTED_COUNTRY_IN_PROGRESS_SENTINEL]) {
      return;
    }
    rootScope[PAYPAL_HOSTED_COUNTRY_IN_PROGRESS_SENTINEL] = true;
    ensureHostedCountryUnitedStates({ countryCode: PAYPAL_HOSTED_COUNTRY_TARGET_CODE })
      .then((result) => {
        if (result?.selected || result?.alreadySelected) {
          log('PayPal 登录页底部国家/地区已自动切换为美国。', 'ok');
        } else if (result?.changed) {
          log('已尝试将 PayPal 登录页底部国家/地区自动切换为美国。', 'info');
        } else if (result?.reason) {
          log(`PayPal 登录页底部国家/地区自动切换跳过：${result.reason}`, 'info');
        }
      })
      .catch((error) => {
        log(`PayPal 登录页底部国家/地区自动切换失败：${error?.message || error}`, 'warn');
      })
      .finally(() => {
        rootScope[PAYPAL_HOSTED_COUNTRY_IN_PROGRESS_SENTINEL] = false;
      });
  }, 1500);
}

function findPasskeyPromptButtons() {
  const promptPatterns = [
    /passkey|通行密钥|安全密钥|下次登录|faster|save/i,
  ];
  const bodyText = normalizeText(document.body?.innerText || '');
  const likelyPrompt = promptPatterns.some((pattern) => pattern.test(bodyText));
  if (!likelyPrompt) {
    return [];
  }

  const cancelOrClose = getVisibleControls('button, a, [role="button"]')
    .filter((el) => {
      const text = getActionText(el);
      return /取消|稍后|不保存|不用|关闭|cancel|not now|maybe later|skip|close|x/i.test(text)
        || el.getAttribute?.('aria-label')?.match(/close|关闭/i);
    });

  const iconCloseButtons = getVisibleControls('button, [role="button"]')
    .filter((el) => {
      const text = getActionText(el);
      const rect = el.getBoundingClientRect();
      return (/^×$|^x$/i.test(text) || /close|关闭/i.test(text))
        && rect.width <= 64
        && rect.height <= 64;
    });

  return [...cancelOrClose, ...iconCloseButtons];
}

function hasPasskeyPrompt() {
  return findPasskeyPromptButtons().length > 0;
}

function getPayPalLoginPhase(emailInput, passwordInput) {
  const emailNextButton = findEmailNextButton();
  const passwordLoginButton = findPasswordLoginButton();
  if (emailInput && emailNextButton && isEnabledControl(emailNextButton) && (!passwordInput || !passwordLoginButton)) {
    return 'email';
  }
  if (emailInput && passwordInput) return 'login_combined';
  if (passwordInput) return 'password';
  if (emailInput) return 'email';
  return '';
}

function refillPayPalEmailInput(emailInput, email) {
  if (!emailInput) return;
  if (typeof emailInput.focus === 'function') {
    emailInput.focus();
  }
  fillInput(emailInput, '');
  fillInput(emailInput, email);
  if (typeof emailInput.blur === 'function') {
    emailInput.blur();
  }
}

async function submitPayPalLogin(payload = {}) {
  const delayOperation = typeof performPayPalOperationWithDelay === 'function'
    ? performPayPalOperationWithDelay
    : async (metadata, operation) => {
        const rootScope = typeof window !== 'undefined' ? window : globalThis;
        const gate = rootScope?.CodexOperationDelay?.performOperationWithDelay;
        return typeof gate === 'function' ? gate(metadata, operation) : operation();
      };
  await waitForDocumentComplete();

  const email = normalizeText(payload.email || '');
  const password = String(payload.password || '');
  if (!password) {
    throw new Error('PayPal 密码为空，请先在侧边栏配置。');
  }

  let passwordInput = findPasswordInput();
  const emailInput = findEmailInput();
  const emailNextButton = findEmailNextButton();

  if (emailInput && emailNextButton && isEnabledControl(emailNextButton) && (!passwordInput || !findPasswordLoginButton())) {
    await delayOperation({ stepKey: 'paypal-approve', kind: 'submit', label: 'paypal-email' }, async () => {
      refillPayPalEmailInput(emailInput, email);
      simulateClick(emailNextButton);
    });
    return {
      submitted: false,
      phase: 'email_submitted',
      awaiting: 'password_page',
    };
  }

  if (!passwordInput && emailInput && email) {
    await delayOperation({ stepKey: 'paypal-approve', kind: 'submit', label: 'paypal-email' }, async () => {
      refillPayPalEmailInput(emailInput, email);
      const nextButton = await waitUntil(() => {
        const button = findEmailNextButton() || findLoginNextButton();
        return button && isEnabledControl(button) ? button : null;
      }, {
        intervalMs: 250,
        timeoutMs: 8000,
        timeoutMessage: 'PayPal email page did not expose a clickable next/continue button.',
      });
      simulateClick(nextButton);
    });
    return {
      submitted: false,
      phase: 'email_submitted',
      awaiting: 'password_page',
    };
  } else if (!passwordInput && emailInput && !email) {
    throw new Error('PayPal 账号为空，请先在侧边栏配置。');
  } else if (emailInput && email) {
    await delayOperation({ stepKey: 'paypal-approve', kind: 'fill', label: 'paypal-email' }, async () => {
      refillPayPalEmailInput(emailInput, email);
    });
  }

  passwordInput = passwordInput || await waitUntil(() => findPasswordInput(), {
    intervalMs: 250,
    timeoutMs: 8000,
    timeoutMessage: 'PayPal password page did not expose a password input.',
  });
  await delayOperation({ stepKey: 'paypal-approve', kind: 'submit', label: 'paypal-password' }, async () => {
    fillInput(passwordInput, password);
    await sleep(1000);

    const loginButton = await waitUntil(() => {
      const button = findClickableByText([
        /login|log\s*in|sign\s*in|continue/i,
        /登录|登入|继续/i,
      ]);
      return button && isEnabledControl(button) ? button : null;
    }, {
      intervalMs: 250,
      timeoutMs: 8000,
      timeoutMessage: 'PayPal password page did not expose a clickable login/continue button.',
    });

    simulateClick(loginButton);
  });
  return {
    submitted: true,
    phase: 'password_submitted',
    awaiting: 'redirect_or_approval',
  };
}

async function dismissPayPalPrompts() {
  const delayOperation = typeof performPayPalOperationWithDelay === 'function'
    ? performPayPalOperationWithDelay
    : async (metadata, operation) => {
        const rootScope = typeof window !== 'undefined' ? window : globalThis;
        const gate = rootScope?.CodexOperationDelay?.performOperationWithDelay;
        return typeof gate === 'function' ? gate(metadata, operation) : operation();
      };
  await waitForDocumentComplete();
  const buttons = findPasskeyPromptButtons();
  let clicked = 0;
  for (const button of buttons) {
    if (!isVisibleElement(button) || !isEnabledControl(button)) {
      continue;
    }
    await delayOperation({ stepKey: 'paypal-approve', kind: 'click', label: 'paypal-dismiss-prompt' }, async () => {
      simulateClick(button);
    });
    clicked += 1;
    await sleep(500);
  }
  return {
    clicked,
    hasPromptAfterClick: hasPasskeyPrompt(),
  };
}

async function clickPayPalApprove() {
  const delayOperation = typeof performPayPalOperationWithDelay === 'function'
    ? performPayPalOperationWithDelay
    : async (metadata, operation) => {
        const rootScope = typeof window !== 'undefined' ? window : globalThis;
        const gate = rootScope?.CodexOperationDelay?.performOperationWithDelay;
        return typeof gate === 'function' ? gate(metadata, operation) : operation();
      };
  await waitForDocumentComplete();
  await dismissPayPalPrompts().catch(() => ({ clicked: 0 }));

  const button = findApproveButton();
  if (!button || !isEnabledControl(button)) {
    return {
      clicked: false,
      state: inspectPayPalState(),
    };
  }

  await delayOperation({ stepKey: 'paypal-approve', kind: 'click', label: 'paypal-approve' }, async () => {
    simulateClick(button);
  });
  return {
    clicked: true,
    buttonText: getActionText(button),
  };
}

function inspectPayPalState() {
  const emailInput = findEmailInput();
  const passwordInput = findPasswordInput();
  const approveButton = findApproveButton();
  const loginPhase = getPayPalLoginPhase(emailInput, passwordInput);
  const hostedStage = detectPayPalHostedCheckoutStage();
  const hasHostedGuestCheckout = isPayPalHostedGuestCheckoutPage();
  const hostedGuestCheckoutFormVisible = hasPayPalHostedGuestCheckoutForm();
  return {
    url: location.href,
    readyState: document.readyState,
    hostedStage,
    needsLogin: Boolean(loginPhase),
    loginPhase,
    hasEmailInput: Boolean(emailInput),
    hasPasswordInput: Boolean(passwordInput),
    hostedAccountCreateEmail: hostedStage === PAYPAL_HOSTED_STAGE_ACCOUNT_CREATE_EMAIL,
    hostedAccountCreateEmailContinueReady: Boolean(findHostedAccountCreateEmailContinueButton()),
    hasHostedGuestCheckout,
    hostedGuestCheckoutFormVisible,
    hostedCardDeclined: hasPayPalHostedCardDeclinedError(),
    hostedCardDeclinedMessage: getPayPalHostedCardDeclinedMessage(),
    hostedGenericError: hostedStage === PAYPAL_HOSTED_STAGE_GENERIC_ERROR,
    hostedGenericErrorMessage: getPayPalHostedGenericErrorMessage(),
    verificationInputsVisible: hasHostedVerificationInputs(),
    hostedVerificationInvalidCode: hasHostedInvalidVerificationCodeError(),
    hostedVerificationErrorText: getHostedVerificationErrorText(),
    hostedVerificationResendReady: Boolean(findHostedVerificationResendButton()),
    reviewConsentReady: Boolean(findHostedReviewConsentButton()),
    hostedCountryText: getHostedCountrySelectionText(),
    hostedCountryUnitedStates: isHostedCountryUnitedStates(),
    approveReady: Boolean(approveButton && isEnabledControl(approveButton)),
    approveButtonText: approveButton ? getActionText(approveButton) : '',
    hasPasskeyPrompt: hasPasskeyPrompt(),
    bodyTextPreview: normalizeText(document.body?.innerText || '').slice(0, 240),
  };
}

scheduleHostedHermesAutoRun();
scheduleHostedCountrySwitchAutoRun();

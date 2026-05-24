(function attachBackgroundPayPalApprove(root, factory) {
  root.MultiPageBackgroundPayPalApprove = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createBackgroundPayPalApproveModule() {
  const PAYPAL_SOURCE = 'paypal-flow';
  const PLUS_CHECKOUT_SOURCE = 'plus-checkout';
  const PAYPAL_INJECT_FILES = ['content/utils.js', 'content/operation-delay.js', 'content/paypal-flow.js'];
  const PAYPAL_LOGIN_TRANSITION_TIMEOUT_MS = 30000;
  const PAYPAL_LOGIN_TRANSITION_POLL_MS = 500;
  const PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS = 15000;

  function createPayPalApproveExecutor(deps = {}) {
    const {
      addLog,
      chrome,
      completeNodeFromBackground,
      ensureContentScriptReadyOnTabUntilStopped,
      getTabId,
      isTabAlive,
      queryTabsInAutomationWindow = null,
      sendTabMessageUntilStopped,
      setState,
      sleepWithStop,
      waitForTabCompleteUntilStopped,
      waitForTabUrlMatchUntilStopped,
    } = deps;

    async function resolvePayPalTabId(state = {}) {
      const paypalTabId = await getTabId(PAYPAL_SOURCE);
      if (paypalTabId && await isTabAlive(PAYPAL_SOURCE)) {
        return paypalTabId;
      }
      const discoveredPayPalTabId = await findOpenPayPalTabId();
      if (discoveredPayPalTabId) {
        await addLog('步骤 8：已从当前浏览器标签中发现 PayPal 页面，正在接管继续执行。', 'info');
        return discoveredPayPalTabId;
      }
      const checkoutTabId = await getTabId(PLUS_CHECKOUT_SOURCE);
      if (checkoutTabId) {
        return checkoutTabId;
      }
      const storedTabId = Number(state.plusCheckoutTabId) || 0;
      if (storedTabId) {
        return storedTabId;
      }
      throw new Error('步骤 8：未找到 PayPal 标签页，请先完成步骤 7。');
    }

    async function findOpenPayPalTabId() {
      if (!chrome?.tabs?.query) {
        return 0;
      }

      const queryTabs = typeof queryTabsInAutomationWindow === 'function'
        ? queryTabsInAutomationWindow
        : (queryInfo) => chrome.tabs.query(queryInfo);
      const tabs = await queryTabs({}).catch(() => []);
      const candidates = (Array.isArray(tabs) ? tabs : [])
        .filter((tab) => Number.isInteger(tab?.id) && isPayPalUrl(tab.url || ''));
      if (!candidates.length) {
        return 0;
      }

      const match = candidates.find((tab) => tab.active && tab.currentWindow)
        || candidates.find((tab) => tab.active)
        || candidates[0];
      if (match?.id && chrome?.tabs?.update) {
        await chrome.tabs.update(match.id, { active: true }).catch(() => {});
      }
      return match?.id || 0;
    }

    async function waitForPayPalUrlUntilStopped(tabId, options = {}) {
      const timeoutMs = Math.max(0, Math.floor(Number(options.timeoutMs) || 0));
      if (!timeoutMs) {
        return waitForTabUrlMatchUntilStopped(tabId, (url) => /paypal\./i.test(url));
      }
      const startedAt = Date.now();
      while (Date.now() - startedAt < timeoutMs) {
        const tab = await chrome.tabs.get(tabId).catch(() => null);
        if (!tab) {
          throw new Error('步骤 8：PayPal 标签页已关闭，无法继续等待页面。');
        }
        if (isPayPalUrl(tab.url || '')) {
          return tab;
        }
        await sleepWithStop(PAYPAL_LOGIN_TRANSITION_POLL_MS);
      }
      throw new Error(`步骤 8：等待 PayPal 页面跳转超时（${Math.round(timeoutMs / 1000)}秒）。`);
    }

    async function ensurePayPalReady(tabId, logMessage = '', options = {}) {
      await waitForPayPalUrlUntilStopped(tabId, {
        timeoutMs: options.timeoutMs,
      });
      await waitForTabCompleteUntilStopped(tabId, {
        timeoutMs: options.timeoutMs,
      });
      await sleepWithStop(1000);
      await ensureContentScriptReadyOnTabUntilStopped(PAYPAL_SOURCE, tabId, {
        inject: PAYPAL_INJECT_FILES,
        injectSource: PAYPAL_SOURCE,
        logMessage: logMessage || '步骤 8：PayPal 页面仍在加载，等待脚本就绪...',
      });
    }

    async function getPayPalState(tabId) {
      const result = await sendTabMessageUntilStopped(tabId, PAYPAL_SOURCE, {
        type: 'PAYPAL_GET_STATE',
        source: 'background',
        payload: {},
      });
      if (result?.error) {
        throw new Error(result.error);
      }
      return result || {};
    }

    async function dismissPrompts(tabId) {
      const result = await sendTabMessageUntilStopped(tabId, PAYPAL_SOURCE, {
        type: 'PAYPAL_DISMISS_PROMPTS',
        source: 'background',
        payload: {},
      });
      if (result?.error) {
        throw new Error(result.error);
      }
      return result || {};
    }

    async function ensurePayPalUsCountry(tabId, options = {}) {
      const result = await sendTabMessageUntilStopped(tabId, PAYPAL_SOURCE, {
        type: 'PAYPAL_ENSURE_US_COUNTRY',
        source: 'background',
        payload: {
          countryCode: 'US',
          forceUiFallback: Boolean(options.forceUiFallback),
        },
      });
      if (result?.error) {
        throw new Error(result.error);
      }
      return result || {};
    }

    async function waitForPayPalUsCountryNavigation(tabId) {
      const startedAt = Date.now();
      while (Date.now() - startedAt < PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS) {
        const tab = await chrome.tabs.get(tabId).catch(() => null);
        if (!tab) {
          throw new Error('步骤 8：PayPal 标签页已关闭，无法等待国家/地区切换。');
        }
        const url = tab.url || '';
        if (isPayPalUrl(url) && /[?&]country\.x=US(?:&|$)/i.test(url)) {
          return tab;
        }
        await sleepWithStop(PAYPAL_LOGIN_TRANSITION_POLL_MS);
      }
      throw new Error('步骤 8：等待 PayPal country.x=US 重载超时。');
    }

    async function ensurePayPalUsCountryWithUiFallback(tabId, reason = '') {
      if (reason) {
        await addLog(`步骤 8：${reason}，改用底部国家/地区选择器兜底。`, 'warn');
      } else {
        await addLog('步骤 8：PayPal 页面仍未识别为美国，尝试使用底部国家/地区选择器兜底切换。', 'warn');
      }
      const currentTab = await chrome.tabs.get(tabId).catch(() => null);
      if (!currentTab) {
        throw new Error('步骤 8：PayPal 标签页已关闭，无法确认国家/地区是否已切换为美国。');
      }
      if (!isPayPalUrl(currentTab.url || '')) {
        throw new Error(`步骤 8：PayPal 已跳转离开授权页，无法确认国家/地区是否已切换为美国。当前 URL：${currentTab.url || 'unknown'}`);
      }
      await ensurePayPalReady(
        tabId,
        '步骤 8：PayPal 国家/地区兜底切换前正在确认页面就绪...',
        { timeoutMs: PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS }
      );
      const countryResult = await ensurePayPalUsCountry(tabId, { forceUiFallback: true }).catch((error) => ({
        error: error?.message || String(error || ''),
      }));
      if (countryResult?.selected || countryResult?.alreadySelected) {
        await addLog('步骤 8：PayPal 登录页底部国家/地区兜底切换为美国。', 'ok');
        await ensurePayPalReady(
          tabId,
          '步骤 8：PayPal 国家/地区兜底切换后正在确认页面就绪...',
          { timeoutMs: PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS }
        );
        return countryResult;
      }
      const details = countryResult?.error
        ? `诊断：${countryResult.error}`
        : '未收到已选中美国的确认结果。';
      throw new Error(`步骤 8：PayPal 国家/地区未能切换为美国，已停止登录以避免误用地区。${details}`);
    }

    function resolvePayPalCredentials(state = {}) {
      const currentId = String(state?.currentPayPalAccountId || '').trim();
      const accounts = Array.isArray(state?.paypalAccounts) ? state.paypalAccounts : [];
      const selectedAccount = currentId
        ? accounts.find((account) => String(account?.id || '').trim() === currentId) || null
        : null;
      return {
        email: String(selectedAccount?.email || state?.paypalEmail || '').trim(),
        password: String(selectedAccount?.password || state?.paypalPassword || ''),
      };
    }

    async function submitLogin(tabId, state = {}) {
      const credentials = resolvePayPalCredentials(state);
      if (!credentials.password) {
        throw new Error('步骤 8：未配置可用的 PayPal 账号，请先在侧边栏添加并选择账号。');
      }
      await addLog('步骤 8：正在填写 PayPal 登录信息并提交...', 'info');
      const result = await sendTabMessageUntilStopped(tabId, PAYPAL_SOURCE, {
        type: 'PAYPAL_SUBMIT_LOGIN',
        source: 'background',
        payload: {
          email: credentials.email,
          password: credentials.password,
        },
      });
      if (result?.error) {
        throw new Error(result.error);
      }
      return result || {};
    }

    function isPayPalUrl(url = '') {
      return /paypal\./i.test(String(url || ''));
    }

    function isPayPalPasswordState(pageState = {}) {
      return Boolean(pageState.hasPasswordInput)
        || pageState.loginPhase === 'password'
        || pageState.loginPhase === 'login_combined';
    }

    async function waitForPayPalPostLoginDecision(tabId, actionResult = {}) {
      const phase = String(actionResult?.phase || '').trim();
      const startedAt = Date.now();

      while (Date.now() - startedAt < PAYPAL_LOGIN_TRANSITION_TIMEOUT_MS) {
        const tab = await chrome.tabs.get(tabId).catch(() => null);
        if (!tab) {
          throw new Error('步骤 8：PayPal 标签页已关闭，无法继续识别登录后的页面。');
        }

        const currentUrl = tab.url || '';
        if (!currentUrl) {
          await sleepWithStop(PAYPAL_LOGIN_TRANSITION_POLL_MS);
          continue;
        }
        if (currentUrl && !isPayPalUrl(currentUrl)) {
          return {
            outcome: 'left_paypal',
            url: currentUrl,
          };
        }

        if (tab.status !== 'complete') {
          await sleepWithStop(PAYPAL_LOGIN_TRANSITION_POLL_MS);
          continue;
        }

        await ensurePayPalReady(
          tabId,
          phase === 'email_submitted'
            ? '步骤 8：PayPal 账号已提交，正在识别下一页...'
            : '步骤 8：PayPal 密码已提交，正在识别跳转结果...'
        );
        const pageState = await getPayPalState(tabId);

        if (pageState.hasPasskeyPrompt) {
          return {
            outcome: 'prompt',
            pageState,
          };
        }

        if (pageState.approveReady) {
          return {
            outcome: 'approve_ready',
            pageState,
          };
        }

        if (phase === 'email_submitted' && isPayPalPasswordState(pageState)) {
          return {
            outcome: 'password_ready',
            pageState,
          };
        }

        if (phase === 'password_submitted' && !pageState.needsLogin) {
          return {
            outcome: 'post_login_state',
            pageState,
          };
        }

        await sleepWithStop(PAYPAL_LOGIN_TRANSITION_POLL_MS);
      }

      return {
        outcome: 'timeout',
        phase,
      };
    }

    async function clickApprove(tabId) {
      const result = await sendTabMessageUntilStopped(tabId, PAYPAL_SOURCE, {
        type: 'PAYPAL_CLICK_APPROVE',
        source: 'background',
        payload: {},
      });
      if (result?.error) {
        throw new Error(result.error);
      }
      return Boolean(result?.clicked);
    }

    async function executePayPalApprove(state = {}) {
      const tabId = await resolvePayPalTabId(state);
      await ensurePayPalReady(tabId);
      await setState({ plusCheckoutTabId: tabId });

      let loggedWaiting = false;
      let attemptedLoginCountrySwitch = false;
      let attemptedLoginCountryFallback = false;
      while (true) {
        const currentUrl = (await chrome.tabs.get(tabId).catch(() => null))?.url || '';
        if (currentUrl && !isPayPalUrl(currentUrl)) {
          await addLog('步骤 8：PayPal 已跳转离开授权页，准备进入回跳确认。', 'ok');
          break;
        }

        await ensurePayPalReady(tabId, '步骤 8：PayPal 页面正在切换，等待脚本重新就绪...');
        const pageState = await getPayPalState(tabId);

        if (pageState.needsLogin) {
          if (!attemptedLoginCountrySwitch) {
            attemptedLoginCountrySwitch = true;
            const countryResult = await ensurePayPalUsCountry(tabId).catch((error) => ({
              error: error?.message || String(error || ''),
            }));
            if (countryResult?.selected || countryResult?.alreadySelected) {
              await addLog('步骤 8：PayPal 登录页底部国家/地区已切换为美国。', 'ok');
            } else if (countryResult?.changed) {
              await addLog('步骤 8：已尝试将 PayPal 登录页底部国家/地区切换为美国，继续登录。', 'info');
            } else if (countryResult?.error) {
              await addLog(`步骤 8：PayPal 国家/地区切换未完成，继续尝试登录：${countryResult.error}`, 'warn');
            }
            if (countryResult?.navigationStarted) {
              await addLog('步骤 8：PayPal 正在按 country.x=US 重新加载登录页，等待页面切换后继续。', 'info');
              try {
                await waitForPayPalUsCountryNavigation(tabId);
                await waitForTabCompleteUntilStopped(tabId, {
                  timeoutMs: PAYPAL_COUNTRY_SWITCH_NAVIGATION_TIMEOUT_MS,
                });
                await sleepWithStop(1000);
                loggedWaiting = false;
                continue;
              } catch (error) {
                attemptedLoginCountryFallback = true;
                await ensurePayPalUsCountryWithUiFallback(
                  tabId,
                  `PayPal country.x=US 重载未确认：${error?.message || error}`
                );
              }
            }
            await ensurePayPalReady(tabId, '步骤 8：PayPal 国家/地区切换后正在确认页面就绪...');
          } else if (!attemptedLoginCountryFallback && pageState.hostedCountryUnitedStates === false) {
            attemptedLoginCountryFallback = true;
            await ensurePayPalUsCountryWithUiFallback(tabId);
          }
          const submitResult = await submitLogin(tabId, state);
          const decision = await waitForPayPalPostLoginDecision(tabId, submitResult);
          if (decision.outcome === 'left_paypal') {
            await addLog('步骤 8：PayPal 登录后已跳转离开登录/授权页，继续进入回跳确认。', 'ok');
            break;
          }
          if (decision.outcome === 'password_ready') {
            await addLog('步骤 8：PayPal 账号页提交后已识别到密码页，继续填写密码。', 'info');
          } else if (decision.outcome === 'approve_ready') {
            await addLog('步骤 8：PayPal 登录后已识别到授权确认页，继续点击授权。', 'info');
          } else if (decision.outcome === 'prompt') {
            await addLog('步骤 8：PayPal 登录后已识别到提示弹窗，继续处理弹窗。', 'info');
          } else if (decision.outcome === 'timeout') {
            await addLog('步骤 8：PayPal 登录动作后暂未识别到新页面，重新检查当前页面状态。', 'warn');
          }
          loggedWaiting = false;
          continue;
        }

        if (pageState.hasPasskeyPrompt) {
          await addLog('步骤 8：检测到 PayPal 通行密钥提示，正在关闭...', 'info');
          await dismissPrompts(tabId);
          await sleepWithStop(1000);
          continue;
        }

        const dismissed = await dismissPrompts(tabId).catch(() => ({ clicked: 0 }));
        if (dismissed.clicked) {
          await sleepWithStop(1000);
          continue;
        }

        if (pageState.approveReady) {
          await addLog('步骤 8：正在点击 PayPal“同意并继续”...', 'info');
          const clicked = await clickApprove(tabId);
          if (clicked) {
            await setState({ plusPaypalApprovedAt: Date.now() });
            break;
          }
        }

        if (!loggedWaiting) {
          loggedWaiting = true;
          await addLog('步骤 8：等待 PayPal 授权按钮或下一步页面出现...', 'info');
        }
        await sleepWithStop(500);
      }

      await completeNodeFromBackground('paypal-approve', {
        plusPaypalApprovedAt: Date.now(),
      });
    }

    return {
      executePayPalApprove,
    };
  }

  return {
    createPayPalApproveExecutor,
  };
});

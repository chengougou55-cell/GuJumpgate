(function outlookEmailProviderModule(root, factory) {
  root.MultiPageBackgroundOutlookEmailProvider = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createOutlookEmailProviderModule() {
  const OUTLOOK_EMAIL_INBOX_LIST_TIMEOUT_MS = 12000;
  const OUTLOOK_EMAIL_JUNK_LIST_TIMEOUT_MS = 8000;
  const OUTLOOK_EMAIL_DETAIL_TIMEOUT_MS = 8000;
  const OUTLOOK_EMAIL_DETAIL_LIMIT = 3;

  function createOutlookEmailProvider(deps = {}) {
    const {
      addLog = async () => {},
      buildOutlookEmailHeaders,
      DEFAULT_OUTLOOK_EMAIL_API_KEY = '',
      DEFAULT_OUTLOOK_EMAIL_BASE_URL = '',
      OUTLOOK_EMAIL_DEFAULT_PAGE_SIZE = 20,
      OUTLOOK_EMAIL_GENERATOR = 'outlook-email',
      OUTLOOK_EMAIL_PROVIDER = 'outlook-email',
      fetchImpl = typeof fetch === 'function' ? fetch.bind(globalThis) : null,
      getState = async () => ({}),
      joinOutlookEmailUrl,
      normalizeOutlookEmailAccounts,
      normalizeOutlookEmailAddress,
      normalizeOutlookEmailBaseUrl,
      normalizeOutlookEmailMailApiDetail,
      normalizeOutlookEmailMailApiMessages,
      persistRegistrationEmailState = null,
      pickVerificationMessageWithTimeFallback,
      setEmailState = async () => {},
      setPersistentSettings = async () => {},
      sleepWithStop = async () => {},
      throwIfStopped = () => {},
    } = deps;

    async function persistResolvedEmailState(state = null, email, options = {}) {
      if (typeof persistRegistrationEmailState === 'function') {
        await persistRegistrationEmailState(state, email, options);
        return;
      }
      await setEmailState(email, options);
    }

    function normalizeOutlookEmailReceiveMailbox(value = '') {
      const normalized = normalizeOutlookEmailAddress(value);
      if (!normalized) return '';
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized) ? normalized : '';
    }

    function normalizeOutlookEmailUsedAddresses(value = []) {
      const source = Array.isArray(value) ? value : [];
      const addresses = [];
      const seen = new Set();
      for (const item of source) {
        const email = normalizeOutlookEmailReceiveMailbox(item);
        if (!email || seen.has(email)) continue;
        seen.add(email);
        addresses.push(email);
      }
      return addresses;
    }

    function getOutlookEmailConfig(state = {}) {
      return {
        baseUrl: normalizeOutlookEmailBaseUrl(state.outlookEmailBaseUrl || DEFAULT_OUTLOOK_EMAIL_BASE_URL),
        apiKey: String(state.outlookEmailApiKey || DEFAULT_OUTLOOK_EMAIL_API_KEY || '').trim(),
        receiveMailbox: normalizeOutlookEmailReceiveMailbox(state.outlookEmailReceiveMailbox),
        usedAddresses: normalizeOutlookEmailUsedAddresses(state.outlookEmailUsedAddresses),
      };
    }

    function ensureOutlookEmailConfig(state, options = {}) {
      const { requireApiKey = true } = options;
      const config = getOutlookEmailConfig(state);
      if (!config.baseUrl) {
        throw new Error('outlookEmail 服务地址为空或格式无效。');
      }
      if (requireApiKey && !config.apiKey) {
        throw new Error('outlookEmail API Key 为空。');
      }
      return config;
    }

    async function requestOutlookEmailJson(config, path, options = {}) {
      if (!fetchImpl) {
        throw new Error('outlookEmail 当前运行环境不支持 fetch。');
      }
      const {
        method = 'GET',
        payload,
        searchParams,
        signal,
        timeoutMs = 20000,
      } = options;
      const url = new URL(joinOutlookEmailUrl(config.baseUrl, path));
      if (searchParams && typeof searchParams === 'object') {
        for (const [key, value] of Object.entries(searchParams)) {
          if (value === undefined || value === null || value === '') continue;
          url.searchParams.set(key, String(value));
        }
      }

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(new Error('timeout')), timeoutMs);
      let externalAbortHandler = null;
      if (signal) {
        if (signal.aborted) {
          controller.abort(signal.reason || new Error('aborted'));
        } else {
          externalAbortHandler = () => controller.abort(signal.reason || new Error('aborted'));
          signal.addEventListener('abort', externalAbortHandler, { once: true });
        }
      }
      let response;
      try {
        response = await fetchImpl(url.toString(), {
          method,
          headers: buildOutlookEmailHeaders(config, {
            json: payload !== undefined,
          }),
          body: payload !== undefined ? JSON.stringify(payload) : undefined,
          signal: controller.signal,
        });
      } catch (err) {
        const aborted = err?.name === 'AbortError' || err?.message === 'timeout';
        const errorMessage = aborted
          ? `outlookEmail 请求超时（>${Math.round(timeoutMs / 1000)} 秒）`
          : `outlookEmail 请求失败：${err.message}`;
        throw new Error(errorMessage);
      } finally {
        clearTimeout(timeoutId);
        if (signal && externalAbortHandler) {
          signal.removeEventListener('abort', externalAbortHandler);
        }
      }

      const text = await response.text();
      let parsed;
      try {
        parsed = text ? JSON.parse(text) : {};
      } catch {
        parsed = text;
      }
      if (!response.ok) {
        const payloadError = typeof parsed === 'object' && parsed
          ? (parsed.message || parsed.error || parsed.msg)
          : '';
        throw new Error(`outlookEmail 请求失败：${payloadError || text || `HTTP ${response.status}`}`);
      }
      if (parsed && typeof parsed === 'object' && parsed.success === false) {
        throw new Error(`outlookEmail 业务错误：${parsed.message || parsed.error || parsed.msg || '请求失败'}`);
      }
      return parsed;
    }

    async function listOutlookEmailAccounts(state, options = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailConfig(latestState);
      const limit = Math.max(1, Number(options.limit) || 10000);
      const offset = Math.max(0, Number(options.offset) || 0);
      const payload = await requestOutlookEmailJson(config, '/api/external/accounts', {
        method: 'GET',
        searchParams: {
          limit,
          offset,
          sort_by: options.sortBy || 'created_at',
          sort_order: options.sortOrder || 'desc',
          group_id: options.groupId || '',
        },
      });
      return {
        config,
        accounts: normalizeOutlookEmailAccounts(payload),
      };
    }

    function pickOutlookEmailAccount(accounts = [], state = {}) {
      const config = getOutlookEmailConfig(state);
      const usedSet = new Set(config.usedAddresses);
      const currentEmail = normalizeOutlookEmailReceiveMailbox(state.email);
      if (currentEmail) {
        const currentAccount = accounts.find((account) => account.email === currentEmail) || null;
        if (currentAccount) return currentAccount;
      }
      return accounts.find((account) => account?.email && !usedSet.has(account.email)) || accounts[0] || null;
    }

    async function setOutlookEmailAddressUsed(email, options = {}) {
      const address = normalizeOutlookEmailReceiveMailbox(email);
      if (!address) return { updated: false };

      const state = options.state || await getState();
      const usedAddresses = normalizeOutlookEmailUsedAddresses([
        ...(state.outlookEmailUsedAddresses || []),
        address,
      ]);
      await setPersistentSettings({ outlookEmailUsedAddresses: usedAddresses });
      await addLog(`${options.logPrefix || 'outlookEmail'}：已将 ${address} 标记为已用。`, options.level || 'ok');
      return { updated: true, outlookEmailUsedAddresses: usedAddresses };
    }

    async function fetchOutlookEmailAddress(state, options = {}) {
      throwIfStopped();
      const latestState = state || await getState();
      const { accounts } = await listOutlookEmailAccounts(latestState, {
        limit: options.limit || 10000,
        groupId: options.groupId || latestState.outlookEmailGroupId || '',
      });
      const account = pickOutlookEmailAccount(accounts, latestState);
      if (!account?.email) {
        throw new Error('outlookEmail 没有返回可用邮箱。');
      }

      await persistResolvedEmailState(latestState, account.email, {
        source: 'generated:outlook-email',
        preserveAccountIdentity: Boolean(options?.preserveAccountIdentity),
      });
      await addLog(`outlookEmail：已取用 ${account.email}`, 'ok');
      return account.email;
    }

    function resolveOutlookEmailPollTargetEmail(state = {}, pollPayload = {}, config = getOutlookEmailConfig(state)) {
      const configuredReceiveMailbox = normalizeOutlookEmailReceiveMailbox(config.receiveMailbox);
      const mailProvider = String(state?.mailProvider || '').trim().toLowerCase();
      const emailGenerator = String(state?.emailGenerator || '').trim().toLowerCase();
      if (mailProvider === OUTLOOK_EMAIL_PROVIDER && emailGenerator !== OUTLOOK_EMAIL_GENERATOR && configuredReceiveMailbox) {
        return configuredReceiveMailbox;
      }

      const requestedTarget = normalizeOutlookEmailReceiveMailbox(pollPayload.targetEmail);
      if (requestedTarget) {
        return requestedTarget;
      }

      return normalizeOutlookEmailReceiveMailbox(state.email);
    }

    async function listOutlookEmailMessages(state, options = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailConfig(latestState);
      const address = normalizeOutlookEmailReceiveMailbox(options.address);
      if (!address) {
        throw new Error('outlookEmail 查信缺少目标邮箱地址。');
      }
      const top = Math.max(1, Math.min(50, Number(options.limit) || OUTLOOK_EMAIL_DEFAULT_PAGE_SIZE));
      const folder = String(options.folder || 'inbox').trim().toLowerCase();
      const payload = await requestOutlookEmailJson(config, '/api/external/emails', {
        method: 'GET',
        signal: options.signal,
        timeoutMs: options.timeoutMs || OUTLOOK_EMAIL_INBOX_LIST_TIMEOUT_MS,
        searchParams: {
          email: address,
          folder,
          top,
          skip: Math.max(0, Number(options.skip) || 0),
          subject_contains: options.subjectContains || '',
          from_contains: options.fromContains || '',
          keyword: options.keyword || '',
        },
      });
      const messages = normalizeOutlookEmailMailApiMessages(payload, folder)
        .filter((message) => !message.address || message.address === address);
      return { config, messages, payload };
    }

    async function getOutlookEmailMessageDetail(state, options = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailConfig(latestState);
      const address = normalizeOutlookEmailReceiveMailbox(options.address);
      const messageId = String(options.messageId || '').trim();
      if (!address || !messageId) {
        return { config, message: null, payload: null };
      }
      const payload = await requestOutlookEmailJson(
        config,
        `/api/external/email/${encodeURIComponent(address)}/${encodeURIComponent(messageId)}`,
        {
          method: 'GET',
          searchParams: {
            folder: options.folder || '',
            method: options.method || '',
          },
          signal: options.signal,
          timeoutMs: options.timeoutMs || 15000,
        }
      );
      const normalized = typeof normalizeOutlookEmailMailApiDetail === 'function'
        ? normalizeOutlookEmailMailApiDetail(payload, options.folder || '')
        : null;
      const base = options.baseMessage || {};
      const message = normalized
        ? {
          ...base,
          ...normalized,
          id: normalized.id || base.id,
          folder: normalized.folder || base.folder,
          mailbox: normalized.mailbox || base.mailbox,
          receivedDateTime: normalized.receivedDateTime || base.receivedDateTime,
        }
        : null;
      return { config, message, payload };
    }

    async function fetchOutlookEmailMessageDetails(state, messages = [], options = {}) {
      const details = [];
      const limit = Math.max(0, Math.min(10, Number(options.limit) || 0));
      const candidates = (messages || [])
        .slice()
        .sort((left, right) => {
          const leftTime = Date.parse(left?.receivedDateTime || '') || 0;
          const rightTime = Date.parse(right?.receivedDateTime || '') || 0;
          return rightTime - leftTime;
        })
        .slice(0, limit);
      const detailResults = await Promise.all(candidates.map(async (message) => {
        throwIfStopped();
        if (!message?.id) return null;
        try {
          const { message: detailMessage } = await getOutlookEmailMessageDetail(state, {
            address: options.address,
            messageId: message.id,
            folder: message.folder || options.folder || '',
            method: message.idMode || message.raw?.id_mode || '',
            baseMessage: message,
            timeoutMs: options.timeoutMs || OUTLOOK_EMAIL_DETAIL_TIMEOUT_MS,
          });
          if (detailMessage) {
            return detailMessage;
          }
        } catch (err) {
          await addLog(`步骤 ${options.step || ''}：outlookEmail 邮件详情读取失败：${err.message}`.replace(/^步骤\s+：/, 'outlookEmail：'), 'warn');
        }
        return null;
      }));
      details.push(...detailResults.filter(Boolean));
      return details;
    }

    async function findOutlookEmailCodeInDetails(state, detailCandidates = [], options = {}) {
      const candidates = Array.isArray(detailCandidates) ? detailCandidates : [];
      if (!candidates.length) {
        return { matchResult: null, match: null };
      }

      const controllers = new Map();
      let resolved = false;
      const matchOptions = buildOutlookEmailMatchOptions(options.pollPayload || {});
      const timeoutMs = options.timeoutMs || OUTLOOK_EMAIL_DETAIL_TIMEOUT_MS;

      const tasks = candidates.map((message, index) => {
        const controller = new AbortController();
        controllers.set(index, controller);
        return (async () => {
          throwIfStopped();
          try {
            const { message: detailMessage } = await getOutlookEmailMessageDetail(state, {
              address: options.address,
              messageId: message.id,
              folder: message.folder || options.folder || '',
              method: message.idMode || message.raw?.id_mode || '',
              baseMessage: message,
              signal: controller.signal,
              timeoutMs,
            });
            if (!detailMessage) {
              return { index, matchResult: null, match: null };
            }
            const detailMatchResult = pickVerificationMessageWithTimeFallback([detailMessage], matchOptions);
            return {
              index,
              matchResult: detailMatchResult,
              match: detailMatchResult.match,
            };
          } catch (err) {
            if (!resolved && err?.name !== 'AbortError') {
              await addLog(`步骤 ${options.step || ''}：outlookEmail 邮件详情读取失败：${err.message}`.replace(/^步骤\s+：/, 'outlookEmail：'), 'warn');
            }
            return {
              index,
              matchResult: null,
              match: null,
              error: err,
            };
          } finally {
            controllers.delete(index);
          }
        })();
      });

      for (let index = 0; index < tasks.length; index += 1) {
        const result = await tasks[index];
        if (result?.error && !resolved) {
          resolved = true;
          for (const [controllerIndex, controller] of controllers.entries()) {
            if (controllerIndex > index) {
              controller.abort(new Error('higher-priority-detail-failed'));
            }
          }
          return {
            matchResult: null,
            match: null,
            error: result.error,
          };
        }
        if (result?.match?.code) {
          resolved = true;
          for (const [controllerIndex, controller] of controllers.entries()) {
            if (controllerIndex > index) {
              controller.abort(new Error('matched'));
            }
          }
          await addLog(`步骤 ${options.step || ''}：列表邮件未命中，已通过 outlookEmail ${options.folderLabel || '邮件'}详情找到验证码。`, 'warn');
          return result;
        }
      }

      resolved = true;
      return { matchResult: null, match: null };
    }

    function buildOutlookEmailMatchOptions(pollPayload = {}) {
      return {
        afterTimestamp: pollPayload.filterAfterTimestamp || 0,
        senderFilters: pollPayload.senderFilters || [],
        subjectFilters: pollPayload.subjectFilters || [],
        requiredKeywords: pollPayload.requiredKeywords || [],
        requiredAnyKeywords: pollPayload.requiredAnyKeywords || [],
        codePatterns: pollPayload.codePatterns || [],
        excludeCodes: pollPayload.excludeCodes || [],
        preferredSubjectFilters: pollPayload.preferredSubjectFilters || [],
        preferredKeywords: pollPayload.preferredKeywords || [],
        excludedSenderFilters: pollPayload.excludedSenderFilters || [],
        excludedSubjectFilters: pollPayload.excludedSubjectFilters || [],
        excludedKeywords: pollPayload.excludedKeywords || [],
        disableTimeFallback: true,
        requireReceivedTimestamp: true,
      };
    }

    function normalizeFilterList(value = []) {
      return (Array.isArray(value) ? value : [])
        .map((item) => String(item || '').trim().toLowerCase())
        .filter(Boolean);
    }

    function isOutlookEmailDetailCandidate(message, pollPayload = {}) {
      const afterTimestamp = Number(pollPayload.filterAfterTimestamp || 0) || 0;
      const receivedAt = Date.parse(message?.receivedDateTime || '') || 0;
      if (afterTimestamp && (!receivedAt || receivedAt < afterTimestamp)) {
        return false;
      }

      const sender = String(message?.from?.emailAddress?.address || '').toLowerCase();
      const subject = String(message?.subject || '').toLowerCase();
      const preview = String(message?.bodyPreview || '').toLowerCase();
      const combinedText = [subject, sender, preview].filter(Boolean).join(' ');

      const requiredAnyKeywords = normalizeFilterList(pollPayload.requiredAnyKeywords);
      if (requiredAnyKeywords.length && !requiredAnyKeywords.some((item) => combinedText.includes(item))) {
        return false;
      }

      const excludedSenderFilters = normalizeFilterList(pollPayload.excludedSenderFilters);
      const excludedSubjectFilters = normalizeFilterList(pollPayload.excludedSubjectFilters);
      const excludedKeywords = normalizeFilterList(pollPayload.excludedKeywords);
      if (excludedSenderFilters.some((item) => sender.includes(item) || preview.includes(item))) {
        return false;
      }
      if (excludedSubjectFilters.some((item) => subject.includes(item) || preview.includes(item))) {
        return false;
      }
      if (excludedKeywords.some((item) => combinedText.includes(item))) {
        return false;
      }

      const senderFilters = normalizeFilterList(pollPayload.senderFilters);
      const subjectFilters = normalizeFilterList(pollPayload.subjectFilters);
      const requiredKeywords = normalizeFilterList(pollPayload.requiredKeywords);
      if (!senderFilters.length && !subjectFilters.length && !requiredKeywords.length) {
        return true;
      }

      return senderFilters.some((item) => combinedText.includes(item))
        || subjectFilters.some((item) => subject.includes(item) || preview.includes(item))
        || requiredKeywords.some((item) => combinedText.includes(item));
    }

    function scoreOutlookEmailDetailCandidate(message, pollPayload = {}) {
      const sender = String(message?.from?.emailAddress?.address || '').toLowerCase();
      const subject = String(message?.subject || '').toLowerCase();
      const preview = String(message?.bodyPreview || '').toLowerCase();
      const combinedText = [subject, sender, preview].filter(Boolean).join(' ');
      const preferredSubjectFilters = normalizeFilterList(pollPayload.preferredSubjectFilters);
      const preferredKeywords = normalizeFilterList(pollPayload.preferredKeywords);
      return (
        preferredSubjectFilters.some((item) => subject.includes(item) || preview.includes(item)) ? 2 : 0
      ) + (
        preferredKeywords.some((item) => combinedText.includes(item)) ? 1 : 0
      );
    }

    function pickOutlookEmailDetailCandidates(messages = [], pollPayload = {}) {
      const limit = Math.max(1, Math.min(10, Number(pollPayload.detailLimit) || OUTLOOK_EMAIL_DETAIL_LIMIT));
      return (Array.isArray(messages) ? messages : [])
        .filter((message) => message?.id && isOutlookEmailDetailCandidate(message, pollPayload))
        .sort((left, right) => {
          const leftScore = scoreOutlookEmailDetailCandidate(left, pollPayload);
          const rightScore = scoreOutlookEmailDetailCandidate(right, pollPayload);
          if (leftScore !== rightScore) {
            return rightScore - leftScore;
          }
          const leftTime = Date.parse(left?.receivedDateTime || '') || 0;
          const rightTime = Date.parse(right?.receivedDateTime || '') || 0;
          return rightTime - leftTime;
        })
        .slice(0, limit);
    }

    async function findOutlookEmailCodeInMessages(latestState, targetEmail, messages, step, pollPayload = {}, options = {}) {
      let matchResult = pickVerificationMessageWithTimeFallback(messages, buildOutlookEmailMatchOptions(pollPayload));
      let match = matchResult.match;
      if (!match?.code) {
        const detailCandidates = pickOutlookEmailDetailCandidates(messages, pollPayload);
        if (detailCandidates.length) {
          const detailMatch = await findOutlookEmailCodeInDetails(latestState, detailCandidates, {
            address: targetEmail,
            folderLabel: options.folderLabel,
            pollPayload,
            step,
            timeoutMs: pollPayload.detailTimeoutMs || OUTLOOK_EMAIL_DETAIL_TIMEOUT_MS,
          });
          if (detailMatch.match?.code) {
            matchResult = detailMatch.matchResult;
            match = detailMatch.match;
          } else if (detailMatch.error) {
            throw detailMatch.error;
          }
        }
      }
      return { matchResult, match };
    }

    function summarizeOutlookEmailMessagesForLog(messages) {
      return (messages || [])
        .slice()
        .sort((left, right) => {
          const leftTime = Date.parse(left.receivedDateTime || '') || 0;
          const rightTime = Date.parse(right.receivedDateTime || '') || 0;
          return rightTime - leftTime;
        })
        .slice(0, 3)
        .map((message) => {
          const receivedAt = message?.receivedDateTime || '未知时间';
          const sender = message?.from?.emailAddress?.address || '未知发件人';
          const subject = message?.subject || '（无主题）';
          const preview = String(message?.bodyPreview || '').replace(/\s+/g, ' ').trim().slice(0, 80);
          return `[${message.folder || 'inbox'}] ${receivedAt} | ${sender} | ${subject} | ${preview}`;
        })
        .join(' || ');
    }

    async function pollOutlookEmailVerificationCode(step, state, pollPayload = {}) {
      const latestState = state || await getState();
      const config = ensureOutlookEmailConfig(latestState);
      const targetEmail = resolveOutlookEmailPollTargetEmail(latestState, pollPayload, config);
      if (!targetEmail) {
        throw new Error('outlookEmail 轮询前缺少目标邮箱地址。');
      }
      await addLog(`步骤 ${step}：正在轮询 outlookEmail 邮件（${targetEmail}）...`, 'info');

      const maxAttempts = Number(pollPayload.maxAttempts) || 5;
      const intervalMs = Number(pollPayload.intervalMs) || 3000;
      let lastError = null;
      const folderConfigs = [
        {
          folder: 'inbox',
          label: '收件箱',
          timeoutMs: pollPayload.inboxListTimeoutMs || pollPayload.listTimeoutMs || OUTLOOK_EMAIL_INBOX_LIST_TIMEOUT_MS,
        },
        {
          folder: 'junkemail',
          label: '垃圾邮件',
          timeoutMs: pollPayload.junkListTimeoutMs || pollPayload.listTimeoutMs || OUTLOOK_EMAIL_JUNK_LIST_TIMEOUT_MS,
        },
      ];
      for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        throwIfStopped();
        try {
          const results = [];
          const folderTasks = folderConfigs.map((folderConfig) => {
            const controller = new AbortController();
            return {
              folderConfig,
              controller,
              promise: (async () => {
                if (Number(folderConfig.startDelayMs) > 0) {
                  await sleepWithStop(Number(folderConfig.startDelayMs));
                }
                throwIfStopped();
                try {
                  const result = await listOutlookEmailMessages(latestState, {
                    address: targetEmail,
                    folder: folderConfig.folder,
                    limit: pollPayload.limit || OUTLOOK_EMAIL_DEFAULT_PAGE_SIZE,
                    signal: controller.signal,
                    timeoutMs: folderConfig.timeoutMs,
                  });
                  return {
                    folderConfig,
                    messages: result.messages || [],
                    error: null,
                  };
                } catch (err) {
                  return {
                    folderConfig,
                    messages: [],
                    error: err,
                  };
                }
              })(),
            };
          });

          const abortLowerPriorityFolderTasks = (folderIndex) => {
            for (let index = folderIndex + 1; index < folderTasks.length; index += 1) {
              folderTasks[index].controller.abort(new Error('higher-priority-folder-matched'));
            }
          };

          const folderErrors = [];
          for (let folderIndex = 0; folderIndex < folderTasks.length; folderIndex += 1) {
            const folderResult = await folderTasks[folderIndex].promise;
            const { folderConfig, messages, error } = folderResult;
            if (error) {
              folderErrors.push({ label: folderConfig.label, error });
              await addLog(`步骤 ${step}：outlookEmail ${folderConfig.label}轮询失败：${error.message}`, 'warn');
              continue;
            }
            try {
              results.push(...messages);
              const { matchResult, match } = await findOutlookEmailCodeInMessages(
                latestState,
                targetEmail,
                messages,
                step,
                pollPayload,
                { folderLabel: folderConfig.label }
              );
              if (match?.code) {
                if (matchResult.usedRelaxedFilters) {
                  const fallbackLabel = matchResult.usedTimeFallback ? '宽松匹配 + 时间回退' : '宽松匹配';
                  await addLog(`步骤 ${step}：严格规则未命中，已改用 ${fallbackLabel} 并命中 outlookEmail 验证码。`, 'warn');
                }
                await addLog(`步骤 ${step}：已通过 outlookEmail 找到验证码：${match.code}`, 'ok');
                abortLowerPriorityFolderTasks(folderIndex);
                return {
                  ok: true,
                  code: match.code,
                  emailTimestamp: match.receivedAt || Date.now(),
                  mailId: match.message?.id || '',
                };
              }
            } catch (err) {
              folderErrors.push({ label: folderConfig.label, error: err });
              await addLog(`步骤 ${step}：outlookEmail ${folderConfig.label}匹配失败：${err.message}`, 'warn');
            }
          }
          if (!results.length && folderErrors.length) {
            throw new Error(folderErrors
              .map((item) => `${item.label}：${item.error.message}`)
              .join('；'));
          }

          lastError = new Error(`步骤 ${step}：暂未在 outlookEmail 中找到匹配验证码（${attempt}/${maxAttempts}）。`);
          await addLog(lastError.message, attempt === maxAttempts ? 'warn' : 'info');
          const sample = summarizeOutlookEmailMessagesForLog(results);
          if (sample) {
            await addLog(`步骤 ${step}：最近邮件样本：${sample}`, 'info');
          }
        } catch (err) {
          lastError = err;
          await addLog(`步骤 ${step}：outlookEmail 轮询失败：${err.message}`, 'warn');
        }
        if (attempt < maxAttempts) {
          await sleepWithStop(intervalMs);
        }
      }
      throw lastError || new Error(`步骤 ${step}：未在 outlookEmail 中找到新的匹配验证码。`);
    }

    return {
      ensureOutlookEmailConfig,
      fetchOutlookEmailAddress,
      getOutlookEmailConfig,
      getOutlookEmailMessageDetail,
      fetchOutlookEmailMessageDetails,
      listOutlookEmailAccounts,
      listOutlookEmailMessages,
      normalizeOutlookEmailReceiveMailbox,
      normalizeOutlookEmailUsedAddresses,
      pollOutlookEmailVerificationCode,
      requestOutlookEmailJson,
      resolveOutlookEmailPollTargetEmail,
      setOutlookEmailAddressUsed,
    };
  }

  return {
    createOutlookEmailProvider,
  };
});

(function outlookEmailProviderModule(root, factory) {
  root.MultiPageBackgroundOutlookEmailProvider = factory();
})(typeof self !== 'undefined' ? self : globalThis, function createOutlookEmailProviderModule() {
  const OUTLOOK_EMAIL_INBOX_LIST_TIMEOUT_MS = 12000;
  const OUTLOOK_EMAIL_JUNK_LIST_TIMEOUT_MS = 8000;
  const OUTLOOK_EMAIL_DETAIL_TIMEOUT_MS = 8000;
  const OUTLOOK_EMAIL_DETAIL_LIMIT = 3;
  const OUTLOOK_EMAIL_CLEAN_CHECK_TIMEOUT_MS = 4000;
  const OUTLOOK_EMAIL_CLEAN_CHECK_LIMIT = 10;
  const OUTLOOK_EMAIL_CLEAN_CHECK_DETAIL_LIMIT = 3;
  const OUTLOOK_EMAIL_CLEAN_CHECK_FAILURE_FALLBACK_THRESHOLD = 1;
  const OUTLOOK_EMAIL_CLEAN_CHECK_FOLDERS = Object.freeze(['inbox', 'junkemail']);
  const OUTLOOK_EMAIL_OPENAI_PROBES = Object.freeze([
    { fromContains: 'openai' },
    { keyword: 'openai' },
    { keyword: 'chatgpt' },
    { subjectContains: 'chatgpt' },
  ]);

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
      random = Math.random,
      setEmailState = async () => {},
      setPersistentSettings = async () => {},
      broadcastDataUpdate = async () => {},
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

    function createOutlookEmailError(message, details = {}) {
      const error = new Error(message);
      Object.assign(error, details);
      return error;
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
        throw createOutlookEmailError(errorMessage, {
          transient: aborted || /failed to fetch|network|timeout|timed out|econnreset|etimedout/i.test(String(err?.message || '')),
          cause: err,
        });
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
        const status = Number(response.status) || 0;
        const detail = payloadError || text || `HTTP ${status}`;
        const authError = status === 401 || status === 403;
        const transient = status === 408 || status === 425 || status === 429 || status >= 500;
        throw createOutlookEmailError(
          authError
            ? `outlookEmail API Key 或权限错误（HTTP ${status}）：${detail}`
            : `outlookEmail 请求失败：${detail}`,
          {
            status,
            transient,
            hard: authError,
            payload: parsed,
          }
        );
      }
      if (parsed && typeof parsed === 'object' && parsed.success === false) {
        const detail = parsed.message || parsed.error || parsed.msg || '请求失败';
        throw createOutlookEmailError(`outlookEmail 业务错误：${detail}`, {
          transient: /timeout|timed out|temporar|暂时|稍后|繁忙|重试|cannot access local variable 'detail'/i.test(String(detail || '')),
          payload: parsed,
        });
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

    function shuffleOutlookEmailAccounts(accounts = []) {
      const list = Array.isArray(accounts) ? accounts.slice() : [];
      for (let index = list.length - 1; index > 0; index -= 1) {
        const rawRandom = Number(typeof random === 'function' ? random() : Math.random());
        const boundedRandom = Number.isFinite(rawRandom)
          ? Math.min(0.999999999, Math.max(0, rawRandom))
          : Math.random();
        const swapIndex = Math.floor(boundedRandom * (index + 1));
        [list[index], list[swapIndex]] = [list[swapIndex], list[index]];
      }
      return list;
    }

    function getOutlookEmailAccountCandidates(accounts = [], state = {}) {
      const config = getOutlookEmailConfig(state);
      const usedSet = new Set(config.usedAddresses);
      const uniqueAccounts = [];
      const seen = new Set();
      for (const account of Array.isArray(accounts) ? accounts : []) {
        const email = normalizeOutlookEmailReceiveMailbox(account?.email);
        if (!email || seen.has(email)) continue;
        seen.add(email);
        uniqueAccounts.push({ ...account, email });
      }

      const preferred = [];
      const fallback = [];
      for (const account of uniqueAccounts) {
        if (usedSet.has(account.email)) {
          fallback.push(account);
        } else {
          preferred.push(account);
        }
      }
      return [
        ...shuffleOutlookEmailAccounts(preferred),
        ...shuffleOutlookEmailAccounts(fallback),
      ];
    }

    function isOutlookEmailOpenAiMessage(message = {}) {
      const text = [
        message?.from?.emailAddress?.address,
        message?.subject,
        message?.bodyPreview,
        message?.raw,
      ]
        .filter(Boolean)
        .join(' ')
        .toLowerCase();
      return /\bopenai\b|chatgpt|tm\.openai\.com|auth\.openai\.com|accounts\.openai\.com/.test(text);
    }

    function isOutlookEmailStrongOpenAiMessage(message = {}) {
      const sender = String(message?.from?.emailAddress?.address || '').trim().toLowerCase();
      return /(^|[.@_-])(openai|chatgpt)([.@_-]|$)|tm\.openai\.com|auth\.openai\.com|accounts\.openai\.com/.test(sender);
    }

    function isOutlookEmailSparseMessage(message = {}) {
      return !String(message?.from?.emailAddress?.address || '').trim()
        && !String(message?.subject || '').trim()
        && !String(message?.bodyPreview || '').trim()
        && Boolean(String(message?.id || '').trim());
    }

    function isOutlookEmailAllFolderFallbackError(err) {
      return /folder|mailbox|all|unsupported|invalid|未知|不支持|无效|文件夹|邮箱文件夹|cannot access local variable 'detail'|请求超时|timeout|timed out/i.test(String(err?.message || ''));
    }

    function isOutlookEmailTransientCleanCheckError(err) {
      if (err?.transient === true) return true;
      const status = Number(err?.status || 0) || 0;
      if (status === 408 || status === 425 || status === 429 || status >= 500) return true;
      return /请求超时|timeout|timed out|failed to fetch|network|abort|econnreset|etimedout|temporar|暂时|稍后|繁忙|重试|cannot access local variable 'detail'/i
        .test(String(err?.message || ''));
    }

    function isOutlookEmailRecoverableCleanCheckError(err) {
      return isOutlookEmailTransientCleanCheckError(err) || isOutlookEmailAllFolderFallbackError(err);
    }

    function isOutlookEmailHardCleanCheckError(err) {
      if (err?.hard === true) return true;
      const status = Number(err?.status || 0) || 0;
      if (status === 401 || status === 403) return true;
      return /api key|unauthorized|forbidden|权限|鉴权|认证|密钥/i.test(String(err?.message || ''));
    }

    function shouldUseOutlookEmailUnconfirmedFallback(err, options = {}) {
      if (options.allowUnconfirmedCleanCheckFallback === false) return false;
      if (isOutlookEmailHardCleanCheckError(err)) return false;
      if (err?.recoverable === true) return true;
      return isOutlookEmailRecoverableCleanCheckError(err);
    }

    function normalizeOutlookEmailCleanCheckFolders(value = []) {
      const source = Array.isArray(value) && value.length ? value : OUTLOOK_EMAIL_CLEAN_CHECK_FOLDERS;
      const folders = [];
      const seen = new Set();
      for (const item of source) {
        const folder = String(item || '').trim().toLowerCase();
        if (!folder || seen.has(folder)) continue;
        seen.add(folder);
        folders.push(folder);
      }
      return folders.length ? folders : OUTLOOK_EMAIL_CLEAN_CHECK_FOLDERS.slice();
    }

    function dedupeOutlookEmailMessages(messages = []) {
      const deduped = [];
      const seen = new Set();
      for (const message of Array.isArray(messages) ? messages : []) {
        if (!message) continue;
        const key = String(message.id || '').trim()
          || [
            message.receivedDateTime || '',
            message.from?.emailAddress?.address || '',
            message.subject || '',
            message.bodyPreview || '',
          ].join('|');
        if (!key || seen.has(key)) continue;
        seen.add(key);
        deduped.push(message);
      }
      return deduped;
    }

    function getOutlookEmailCleanCheckProbes(options = {}) {
      if (Array.isArray(options.openAiProbes)) {
        return options.openAiProbes.length ? options.openAiProbes : [{}];
      }
      return options.useProbeSearchForCleanCheck === true
        ? OUTLOOK_EMAIL_OPENAI_PROBES
        : [{}];
    }

    function shouldConfirmOutlookEmailMessageDetail(message = {}, probe = {}) {
      if (isOutlookEmailOpenAiMessage(message)) {
        return true;
      }
      if (isOutlookEmailSparseMessage(message)) {
        return true;
      }
      return Boolean(probe?.keyword || probe?.fromContains || probe?.subjectContains);
    }

    function buildOutlookEmailCleanCheckError(errors = []) {
      const source = Array.isArray(errors) ? errors : [];
      const message = source
        .map((item) => `${item.folder || 'unknown'}：${item.error?.message || item.error || '未知错误'}`)
        .join('；');
      return createOutlookEmailError(message || 'outlookEmail 清理检查失败。', {
        transient: source.length > 0 && source.every((item) => isOutlookEmailTransientCleanCheckError(item.error)),
        recoverable: source.length > 0 && source.every((item) => isOutlookEmailRecoverableCleanCheckError(item.error)),
        hard: source.some((item) => isOutlookEmailHardCleanCheckError(item.error)),
        cleanCheckErrors: source,
      });
    }

    function summarizeOutlookEmailOpenAiMessage(message = {}) {
      const sender = message?.from?.emailAddress?.address || '未知发件人';
      const subject = message?.subject || '（无主题）';
      const receivedAt = message?.receivedDateTime || '未知时间';
      return `${receivedAt} | ${sender} | ${subject}`;
    }

    async function getConfirmedOutlookEmailOpenAiMessage(state, address, message, options = {}) {
      if (!message) return null;
      if (isOutlookEmailStrongOpenAiMessage(message)) {
        return message;
      }
      if (!message.id) {
        return isOutlookEmailOpenAiMessage(message) ? message : null;
      }
      try {
        const { message: detailMessage } = await getOutlookEmailMessageDetail(state, {
          address,
          messageId: message.id,
          folder: message.folder || options.folder || '',
          method: message.idMode || message.raw?.id_mode || '',
          baseMessage: message,
          timeoutMs: options.cleanCheckTimeoutMs || OUTLOOK_EMAIL_CLEAN_CHECK_TIMEOUT_MS,
        });
        return isOutlookEmailOpenAiMessage(detailMessage) ? detailMessage : null;
      } catch (err) {
        await addLog(`outlookEmail：读取 ${address} 的可疑 OpenAI 邮件详情失败：${err.message}`, 'warn');
        throw err;
      }
    }

    async function collectOutlookEmailProbeMessages(state, targetEmail, probe, options = {}) {
      const results = [];
      const folders = normalizeOutlookEmailCleanCheckFolders(options.cleanCheckFolders);
      const timeoutMs = options.cleanCheckTimeoutMs || OUTLOOK_EMAIL_CLEAN_CHECK_TIMEOUT_MS;
      if (options.useAllFolderForCleanCheck === true) {
        try {
          const result = await listOutlookEmailMessages(state, {
            address: targetEmail,
            folder: 'all',
            limit: options.limit,
            timeoutMs,
            ...probe,
          });
          return { messages: result.messages || [], errors: [] };
        } catch (err) {
          if (!isOutlookEmailAllFolderFallbackError(err)) {
            throw err;
          }
          await addLog(`outlookEmail：folder=all 查询失败，回退检查${folders.join(' / ')}：${err.message}`, 'warn');
        }
      }

      const folderResults = await Promise.all(folders.map(async (folder) => {
        throwIfStopped();
        try {
          const { messages } = await listOutlookEmailMessages(state, {
            address: targetEmail,
            folder,
            limit: options.limit,
            timeoutMs,
            ...probe,
          });
          return { folder, messages: messages || [], error: null };
        } catch (error) {
          return { folder, messages: [], error };
        }
      }));

      const errors = [];
      for (const result of folderResults) {
        results.push(...(result.messages || []));
        if (result.error) {
          errors.push({ folder: result.folder, error: result.error });
        }
      }
      return {
        messages: dedupeOutlookEmailMessages(results),
        errors,
      };
    }

    async function listOutlookEmailProbeMessages(state, targetEmail, probe, options = {}) {
      const result = await collectOutlookEmailProbeMessages(state, targetEmail, probe, options);
      if (result.errors.length) {
        throw buildOutlookEmailCleanCheckError(result.errors);
      }
      return result.messages;
    }

    async function checkOutlookEmailHasOpenAiMail(state, address, options = {}) {
      const targetEmail = normalizeOutlookEmailReceiveMailbox(address);
      if (!targetEmail) {
        throw new Error('outlookEmail 检查 OpenAI 邮件时缺少目标邮箱地址。');
      }

      const probes = getOutlookEmailCleanCheckProbes(options);
      const limit = Math.max(1, Math.min(50, Number(options.cleanCheckLimit) || OUTLOOK_EMAIL_CLEAN_CHECK_LIMIT));
      const detailLimit = Math.max(0, Math.min(10, Number(options.cleanCheckDetailLimit) || OUTLOOK_EMAIL_CLEAN_CHECK_DETAIL_LIMIT));
      for (const probe of probes) {
        throwIfStopped();
        const { messages, errors } = await collectOutlookEmailProbeMessages(state, targetEmail, probe, {
          limit,
          timeoutMs: options.cleanCheckTimeoutMs || OUTLOOK_EMAIL_CLEAN_CHECK_TIMEOUT_MS,
          cleanCheckFolders: options.cleanCheckFolders,
          useAllFolderForCleanCheck: options.useAllFolderForCleanCheck,
        });
        let detailChecks = 0;
        let skippedDetailConfirmations = 0;
        for (const message of messages || []) {
          if (!shouldConfirmOutlookEmailMessageDetail(message, probe)) {
            continue;
          }
          const needsNetworkDetail = Boolean(message?.id) && !isOutlookEmailStrongOpenAiMessage(message);
          if (needsNetworkDetail) {
            if (detailChecks >= detailLimit) {
              skippedDetailConfirmations += 1;
              continue;
            }
            detailChecks += 1;
          }
          const confirmedMessage = await getConfirmedOutlookEmailOpenAiMessage(state, targetEmail, message, {
            folder: message.folder || 'all',
            cleanCheckTimeoutMs: options.cleanCheckTimeoutMs,
          });
          if (!confirmedMessage) {
            continue;
          }
          return {
            hasOpenAiMail: true,
            message: confirmedMessage,
            probe,
          };
        }
        if (skippedDetailConfirmations > 0) {
          throw createOutlookEmailError(
            `outlookEmail 清理检查有 ${skippedDetailConfirmations} 封可疑邮件未完成详情确认。`,
            { transient: true }
          );
        }
        if (errors.length) {
          throw buildOutlookEmailCleanCheckError(errors);
        }
      }
      return { hasOpenAiMail: false, message: null, probe: null };
    }

    async function pickOutlookEmailAccount(accounts = [], state = {}, options = {}) {
      const candidates = getOutlookEmailAccountCandidates(accounts, state);
      const requireCleanOpenAiMailbox = options.requireCleanOpenAiMailbox !== false;
      if (!requireCleanOpenAiMailbox) {
        return candidates[0] || null;
      }

      let checkedCount = 0;
      let dirtyCount = 0;
      let failedCount = 0;
      let fallbackAccount = null;
      let fallbackError = null;
      const fallbackFailureThreshold = Math.max(
        1,
        Number(options.cleanCheckFailureFallbackThreshold) || OUTLOOK_EMAIL_CLEAN_CHECK_FAILURE_FALLBACK_THRESHOLD
      );
      for (const account of candidates) {
        throwIfStopped();
        checkedCount += 1;
        try {
          const checkResult = await checkOutlookEmailHasOpenAiMail(state, account.email, options);
          if (!checkResult.hasOpenAiMail) {
            return {
              ...account,
              cleanCheck: {
                checkedCount,
                dirtyCount,
                failedCount,
              },
            };
          }
          dirtyCount += 1;
          await addLog(
            `outlookEmail：跳过 ${account.email}，邮箱内已有 OpenAI/ChatGPT 邮件（${summarizeOutlookEmailOpenAiMessage(checkResult.message)}）。`,
            'info'
          );
        } catch (err) {
          failedCount += 1;
          if (!shouldUseOutlookEmailUnconfirmedFallback(err, options)) {
            await addLog(`outlookEmail：跳过 ${account.email}，检查 OpenAI 邮件失败且不可作为备选：${err.message}`, 'warn');
            if (isOutlookEmailHardCleanCheckError(err)) {
              throw err;
            }
            continue;
          }
          if (!fallbackAccount) {
            fallbackAccount = account;
            fallbackError = err;
          }
          await addLog(`outlookEmail：${account.email} 的 OpenAI 邮件检查失败，已保留为备选：${err.message}`, 'warn');
          if (failedCount >= fallbackFailureThreshold && fallbackAccount) {
            await addLog(
              `outlookEmail：连续 ${failedCount} 个邮箱检查失败，改用未确认备选 ${fallbackAccount.email} 继续流程，避免长时间阻塞。`,
              'warn'
            );
            return {
              ...fallbackAccount,
              cleanCheck: {
                checkedCount,
                dirtyCount,
                failedCount,
                unconfirmed: true,
                error: fallbackError?.message || '',
              },
            };
          }
        }
      }

      if (fallbackAccount) {
        await addLog(
          `outlookEmail：未找到确认干净的邮箱，改用未确认备选 ${fallbackAccount.email} 继续流程。`,
          'warn'
        );
        return {
          ...fallbackAccount,
          cleanCheck: {
            checkedCount,
            dirtyCount,
            failedCount,
            unconfirmed: true,
            error: fallbackError?.message || '',
          },
        };
      }

      const reasonParts = [];
      if (dirtyCount) reasonParts.push(`${dirtyCount} 个已有 OpenAI/ChatGPT 邮件`);
      if (failedCount) reasonParts.push(`${failedCount} 个检查失败`);
      const reason = reasonParts.length ? `（${reasonParts.join('，')}）` : '';
      throw new Error(`outlookEmail 没有找到无 OpenAI 邮件的可用邮箱${reason}。`);
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
      const account = await pickOutlookEmailAccount(accounts, latestState, {
        cleanCheckLimit: options.cleanCheckLimit,
        cleanCheckTimeoutMs: options.cleanCheckTimeoutMs,
        cleanCheckDetailLimit: options.cleanCheckDetailLimit,
        cleanCheckFolders: options.cleanCheckFolders,
        cleanCheckFailureFallbackThreshold: options.cleanCheckFailureFallbackThreshold,
        allowUnconfirmedCleanCheckFallback: options.allowUnconfirmedCleanCheckFallback,
        useAllFolderForCleanCheck: options.useAllFolderForCleanCheck,
        useProbeSearchForCleanCheck: options.useProbeSearchForCleanCheck,
        openAiProbes: options.openAiProbes,
        requireCleanOpenAiMailbox: options.requireCleanOpenAiMailbox,
      });
      if (!account?.email) {
        throw new Error('outlookEmail 没有返回可用邮箱。');
      }

      const usedAddresses = normalizeOutlookEmailUsedAddresses([
        ...(latestState.outlookEmailUsedAddresses || []),
        account.email,
      ]);
      await setPersistentSettings({ outlookEmailUsedAddresses: usedAddresses });
      await broadcastDataUpdate({ outlookEmailUsedAddresses: usedAddresses });
      await persistResolvedEmailState(latestState, account.email, {
        source: 'generated:outlook-email',
        preserveAccountIdentity: Boolean(options?.preserveAccountIdentity),
      });
      if (account.cleanCheck?.unconfirmed) {
        await addLog(`outlookEmail：已取用邮箱 ${account.email}（OpenAI 邮件检查未完成：${account.cleanCheck.error || '远端检查失败'}）。`, 'warn');
      } else {
        await addLog(`outlookEmail：已随机取用无 OpenAI 邮件的邮箱 ${account.email}`, 'ok');
      }
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
          const abortPendingFolderTasks = (pendingFolderTasks) => {
            for (const task of pendingFolderTasks.values()) {
              task.controller.abort(new Error('folder-matched'));
            }
          };

          const folderErrors = [];
          const handleFolderResult = async (folderIndex, folderResult) => {
            const { folderConfig, messages, error } = folderResult;
            if (error) {
              folderErrors.push({ label: folderConfig.label, error });
              await addLog(`步骤 ${step}：outlookEmail ${folderConfig.label}轮询失败：${error.message}`, 'warn');
              return null;
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
            return null;
          };

          const pendingFolderTasks = new Map(folderTasks.map((task, folderIndex) => [folderIndex, task]));
          while (pendingFolderTasks.size) {
            const { folderIndex, folderResult } = await Promise.race(
              Array.from(pendingFolderTasks.entries(), ([currentIndex, task]) => task.promise.then(
                (result) => ({ folderIndex: currentIndex, folderResult: result }),
                (error) => ({
                  folderIndex: currentIndex,
                  folderResult: {
                    folderConfig: task.folderConfig,
                    messages: [],
                    error,
                  },
                })
              ))
            );
            pendingFolderTasks.delete(folderIndex);
            const match = await handleFolderResult(folderIndex, folderResult);
            if (match) {
              abortPendingFolderTasks(pendingFolderTasks);
              return match;
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

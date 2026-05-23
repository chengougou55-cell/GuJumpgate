const assert = require('node:assert/strict');
const test = require('node:test');

const outlookEmailUtils = require('../outlook-email-utils.js');
const hotmailUtils = require('../hotmail-utils.js');

globalThis.MultiPageBackgroundOutlookEmailProvider = null;
require('../background/outlook-email-provider.js');

function makeJsonResponse(payload) {
  return {
    ok: true,
    status: 200,
    async text() {
      return JSON.stringify(payload);
    },
  };
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test('outlookEmail provider reads detail body and does not reuse old code', async () => {
  const requests = [];
  const state = {
    email: 'kellyvicki8809@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    requests.push(parsed);
    if (parsed.pathname === '/api/external/emails') {
      const folder = parsed.searchParams.get('folder');
      return makeJsonResponse({
        emails: folder === 'inbox'
          ? [
            {
              id: '13',
              id_mode: 'imap',
              subject: '你的临时 ChatGPT 登录代码',
              from: 'ChatGPT <noreply@tm.openai.com>',
              body_preview: '你的临时 ChatGPT 登录代码 <style ...',
              date: '2026-05-23T12:21:53Z',
            },
            {
              id: 'old',
              subject: '你的 ChatGPT 临时验证码',
              from: 'noreply@tm.openai.com',
              body_preview: '输入此临时验证码以继续：895185',
              date: '2026-05-23T12:05:01Z',
            },
          ]
          : [],
      });
    }
    if (parsed.pathname.endsWith('/13')) {
      return makeJsonResponse({
        email: {
          id: '13',
          subject: '你的临时 ChatGPT 登录代码',
          from: 'ChatGPT <noreply@tm.openai.com>',
          body: '<html><body><p>输入此代码以登录：</p><div>220655</div></body></html>',
          date: '2026-05-23T12:21:53Z',
        },
      });
    }
    if (parsed.pathname.endsWith('/old')) {
      return makeJsonResponse({
        email: {
          id: 'old',
          subject: '你的 ChatGPT 临时验证码',
          from: 'noreply@tm.openai.com',
          body: '<html><body>输入此临时验证码以继续：895185</body></html>',
          date: '2026-05-23T12:05:01Z',
        },
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const logs = [];
  let junkAborted = false;
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async (message, level) => logs.push({ message, level }),
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl: async (url, options = {}) => {
      const parsed = new URL(url);
      if (parsed.pathname === '/api/external/emails' && parsed.searchParams.get('folder') === 'junkemail') {
        await new Promise((resolve, reject) => {
          const timer = setTimeout(resolve, 120);
          options.signal?.addEventListener('abort', () => {
            clearTimeout(timer);
            junkAborted = true;
            reject(Object.assign(new Error('aborted'), { name: 'AbortError' }));
          }, { once: true });
        });
      }
      return fetchImpl(url, options);
    },
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  const result = await provider.pollOutlookEmailVerificationCode(4, state, {
    filterAfterTimestamp: Date.UTC(2026, 4, 23, 12, 20, 58),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码', '代码'],
    requiredKeywords: ['chatgpt', '代码'],
    requiredAnyKeywords: ['openai', 'chatgpt'],
    preferredSubjectFilters: ['临时验证码'],
    preferredKeywords: ['输入此临时验证码以继续'],
    maxAttempts: 1,
  });

  assert.equal(result.code, '220655');
  assert.equal(result.mailId, '13');
  assert.equal(result.emailTimestamp, Date.UTC(2026, 4, 23, 12, 21, 53));
  const detailRequest = requests.find((item) => item.pathname.endsWith('/13'));
  assert.equal(detailRequest.searchParams.get('method'), 'imap');
  assert.equal(requests.some((item) => item.pathname.endsWith('/old')), false);
  await delay(0);
  assert.equal(junkAborted, true);
  assert.ok(logs.some((item) => item.message.includes('详情找到验证码')));
});

test('outlookEmail provider fails closed when only old codes are available', async () => {
  const state = {
    email: 'kellyvicki8809@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/emails') {
      return makeJsonResponse({
        emails: [
          {
            id: 'old',
            subject: '你的 ChatGPT 临时验证码',
            from: 'noreply@tm.openai.com',
            body_preview: '输入此临时验证码以继续：895185',
            date: '2026-05-23T12:05:01Z',
          },
        ],
      });
    }
    if (parsed.pathname.endsWith('/old')) {
      return makeJsonResponse({
        email: {
          id: 'old',
          subject: '你的 ChatGPT 临时验证码',
          from: 'noreply@tm.openai.com',
          body: '<html><body>输入此临时验证码以继续：895185</body></html>',
          date: '2026-05-23T12:05:01Z',
        },
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  await assert.rejects(
    () => provider.pollOutlookEmailVerificationCode(4, state, {
      filterAfterTimestamp: Date.UTC(2026, 4, 23, 12, 20, 58),
      senderFilters: ['openai', 'noreply'],
      subjectFilters: ['验证码', '代码'],
      requiredKeywords: ['chatgpt', '代码'],
      requiredAnyKeywords: ['openai', 'chatgpt'],
      maxAttempts: 1,
    }),
    /暂未在 outlookEmail 中找到匹配验证码/
  );
});

test('outlookEmail provider continues to junk only when inbox has no match', async () => {
  const requests = [];
  const state = {
    email: 'paulmr7658@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    requests.push(parsed);
    if (parsed.pathname === '/api/external/emails') {
      const folder = parsed.searchParams.get('folder');
      return makeJsonResponse({
        emails: folder === 'junkemail'
          ? [
            {
              id: 'junk-code',
              id_mode: 'sequence',
              subject: '你的 ChatGPT 临时验证码',
              from: 'noreply@tm.openai.com',
              body_preview: '输入此临时验证码以继续：665544',
              date: '2026-05-23T13:43:19Z',
            },
          ]
          : [],
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  const result = await provider.pollOutlookEmailVerificationCode(4, state, {
    filterAfterTimestamp: Date.UTC(2026, 4, 23, 13, 43, 0),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码', '代码'],
    requiredKeywords: ['chatgpt', '代码'],
    requiredAnyKeywords: ['openai', 'chatgpt'],
    maxAttempts: 1,
  });

  assert.equal(result.code, '665544');
  assert.deepEqual(
    requests
      .filter((item) => item.pathname === '/api/external/emails')
      .map((item) => item.searchParams.get('folder')),
    ['inbox', 'junkemail']
  );
});

test('outlookEmail provider lists inbox and junk in parallel', async () => {
  const state = {
    email: 'parallel@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/emails') {
      const folder = parsed.searchParams.get('folder');
      await delay(folder === 'inbox' ? 45 : 50);
      return makeJsonResponse({
        emails: folder === 'junkemail'
          ? [
            {
              id: 'junk-code',
              id_mode: 'sequence',
              subject: '你的 ChatGPT 临时验证码',
              from: 'noreply@tm.openai.com',
              body_preview: '输入此临时验证码以继续：998877',
              date: '2026-05-23T13:43:19Z',
            },
          ]
          : [],
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  const started = Date.now();
  const result = await provider.pollOutlookEmailVerificationCode(4, state, {
    filterAfterTimestamp: Date.UTC(2026, 4, 23, 13, 43, 0),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码', '代码'],
    requiredKeywords: ['chatgpt', '代码'],
    requiredAnyKeywords: ['openai', 'chatgpt'],
    maxAttempts: 1,
  });
  const elapsedMs = Date.now() - started;

  assert.equal(result.code, '998877');
  assert.ok(elapsedMs < 90, `expected parallel list requests, got ${elapsedMs}ms`);
});

test('outlookEmail provider returns on first matching detail and aborts slower details', async () => {
  const requests = [];
  let slowDetailAborted = false;
  let junkAborted = false;
  const state = {
    email: 'fastmatch@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url, options = {}) => {
    const parsed = new URL(url);
    requests.push(parsed);
    if (parsed.pathname === '/api/external/emails') {
      if (parsed.searchParams.get('folder') === 'junkemail') {
        await new Promise((resolve, reject) => {
          const timer = setTimeout(resolve, 120);
          options.signal?.addEventListener('abort', () => {
            clearTimeout(timer);
            junkAborted = true;
            reject(Object.assign(new Error('aborted'), { name: 'AbortError' }));
          }, { once: true });
        });
      }
      return makeJsonResponse({
        emails: [
          {
            id: 'fast-code',
            id_mode: 'sequence',
            subject: '你的临时 ChatGPT 登录代码',
            from: 'ChatGPT <noreply@tm.openai.com>',
            body_preview: '你的临时 ChatGPT 登录代码 <style ...',
            date: '2026-05-23T13:43:19Z',
          },
          {
            id: 'slow-code',
            id_mode: 'sequence',
            subject: '你的临时 ChatGPT 登录代码',
            from: 'ChatGPT <noreply@tm.openai.com>',
            body_preview: '你的临时 ChatGPT 登录代码 <style ...',
            date: '2026-05-23T13:43:18Z',
          },
        ],
      });
    }
    if (parsed.pathname.endsWith('/fast-code')) {
      return makeJsonResponse({
        email: {
          id: 'fast-code',
          subject: '你的临时 ChatGPT 登录代码',
          from: 'ChatGPT <noreply@tm.openai.com>',
          body: '<html><body><p>输入此代码以登录：</p><div>123987</div></body></html>',
          date: '2026-05-23T13:43:19Z',
        },
      });
    }
    if (parsed.pathname.endsWith('/slow-code')) {
      await new Promise((resolve, reject) => {
        const timer = setTimeout(resolve, 120);
        options.signal?.addEventListener('abort', () => {
          clearTimeout(timer);
          slowDetailAborted = true;
          reject(Object.assign(new Error('aborted'), { name: 'AbortError' }));
        }, { once: true });
      });
      return makeJsonResponse({
        email: {
          id: 'slow-code',
          subject: '你的临时 ChatGPT 登录代码',
          from: 'ChatGPT <noreply@tm.openai.com>',
          body: '<html><body><p>输入此代码以登录：</p><div>555666</div></body></html>',
          date: '2026-05-23T13:43:18Z',
        },
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  const started = Date.now();
  const result = await provider.pollOutlookEmailVerificationCode(4, state, {
    filterAfterTimestamp: Date.UTC(2026, 4, 23, 13, 43, 0),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码', '代码'],
    requiredKeywords: ['chatgpt', '代码'],
    requiredAnyKeywords: ['openai', 'chatgpt'],
    maxAttempts: 1,
  });
  const elapsedMs = Date.now() - started;
  await delay(0);

  assert.equal(result.code, '123987');
  assert.equal(result.mailId, 'fast-code');
  assert.equal(slowDetailAborted, true);
  assert.ok(elapsedMs < 100, `expected fast return, got ${elapsedMs}ms`);
  assert.equal(junkAborted, true);
});

test('outlookEmail provider keeps candidate order when lower priority detail returns first', async () => {
  const state = {
    email: 'orderedmatch@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/emails') {
      return makeJsonResponse({
        emails: [
          {
            id: 'newer-code',
            id_mode: 'sequence',
            subject: '你的临时 ChatGPT 登录代码',
            from: 'ChatGPT <noreply@tm.openai.com>',
            body_preview: '你的临时 ChatGPT 登录代码 <style ...',
            date: '2026-05-23T13:43:19Z',
          },
          {
            id: 'older-code',
            id_mode: 'sequence',
            subject: '你的临时 ChatGPT 登录代码',
            from: 'ChatGPT <noreply@tm.openai.com>',
            body_preview: '你的临时 ChatGPT 登录代码 <style ...',
            date: '2026-05-23T13:43:18Z',
          },
        ],
      });
    }
    if (parsed.pathname.endsWith('/newer-code')) {
      await delay(40);
      return makeJsonResponse({
        email: {
          id: 'newer-code',
          subject: '你的临时 ChatGPT 登录代码',
          from: 'ChatGPT <noreply@tm.openai.com>',
          body: '<html><body><p>输入此代码以登录：</p><div>222333</div></body></html>',
          date: '2026-05-23T13:43:19Z',
        },
      });
    }
    if (parsed.pathname.endsWith('/older-code')) {
      return makeJsonResponse({
        email: {
          id: 'older-code',
          subject: '你的临时 ChatGPT 登录代码',
          from: 'ChatGPT <noreply@tm.openai.com>',
          body: '<html><body><p>输入此代码以登录：</p><div>111222</div></body></html>',
          date: '2026-05-23T13:43:18Z',
        },
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  const result = await provider.pollOutlookEmailVerificationCode(4, state, {
    filterAfterTimestamp: Date.UTC(2026, 4, 23, 13, 43, 0),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码', '代码'],
    requiredKeywords: ['chatgpt', '代码'],
    requiredAnyKeywords: ['openai', 'chatgpt'],
    maxAttempts: 1,
  });

  assert.equal(result.code, '222333');
  assert.equal(result.mailId, 'newer-code');
});

test('outlookEmail provider does not fall back to older detail when newest detail fails', async () => {
  let olderDetailAborted = false;
  const state = {
    email: 'newestfail@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url, options = {}) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/emails') {
      return makeJsonResponse({
        emails: [
          {
            id: 'newest-code',
            id_mode: 'sequence',
            subject: '你的临时 ChatGPT 登录代码',
            from: 'ChatGPT <noreply@tm.openai.com>',
            body_preview: '你的临时 ChatGPT 登录代码 <style ...',
            date: '2026-05-23T13:43:19Z',
          },
          {
            id: 'older-code',
            id_mode: 'sequence',
            subject: '你的临时 ChatGPT 登录代码',
            from: 'ChatGPT <noreply@tm.openai.com>',
            body_preview: '你的临时 ChatGPT 登录代码 <style ...',
            date: '2026-05-23T13:43:18Z',
          },
        ],
      });
    }
    if (parsed.pathname.endsWith('/newest-code')) {
      throw Object.assign(new Error('detail timeout'), { name: 'AbortError' });
    }
    if (parsed.pathname.endsWith('/older-code')) {
      await new Promise((resolve, reject) => {
        const timer = setTimeout(resolve, 120);
        options.signal?.addEventListener('abort', () => {
          clearTimeout(timer);
          olderDetailAborted = true;
          reject(Object.assign(new Error('aborted'), { name: 'AbortError' }));
        }, { once: true });
      });
      return makeJsonResponse({
        email: {
          id: 'older-code',
          subject: '你的临时 ChatGPT 登录代码',
          from: 'ChatGPT <noreply@tm.openai.com>',
          body: '<html><body><p>输入此代码以登录：</p><div>444555</div></body></html>',
          date: '2026-05-23T13:43:18Z',
        },
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    sleepWithStop: async () => {},
  });

  await assert.rejects(
    () => provider.pollOutlookEmailVerificationCode(4, state, {
      filterAfterTimestamp: Date.UTC(2026, 4, 23, 13, 43, 0),
      senderFilters: ['openai', 'noreply'],
      subjectFilters: ['验证码', '代码'],
      requiredKeywords: ['chatgpt', '代码'],
      requiredAnyKeywords: ['openai', 'chatgpt'],
      maxAttempts: 1,
    }),
    /detail timeout|暂未在 outlookEmail/
  );
  await delay(0);
  assert.equal(olderDetailAborted, true);
});

test('outlookEmail provider randomly picks an account without OpenAI mail', async () => {
  const state = {
    email: 'old-current@hotmail.com',
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
    outlookEmailUsedAddresses: ['used@hotmail.com'],
  };
  const checkedEmails = [];
  let savedEmail = '';
  let persistedSettings = null;
  const logs = [];
  const randomValues = [0.8, 0.8, 0, 0];
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/accounts') {
      return makeJsonResponse({
        accounts: [
          { id: 'used', email: 'used@hotmail.com' },
          { id: 'dirty', email: 'dirty@hotmail.com' },
          { id: 'clean', email: 'clean@hotmail.com' },
          { id: 'old', email: 'old-current@hotmail.com' },
        ],
      });
    }
    if (parsed.pathname === '/api/external/emails') {
      const email = parsed.searchParams.get('email');
      checkedEmails.push(email);
      if (email === 'dirty@hotmail.com') {
        return makeJsonResponse({
          emails: [
            {
              id: 'openai-old',
              subject: '你的 ChatGPT 临时验证码',
              from: 'noreply@tm.openai.com',
              body_preview: 'OpenAI ChatGPT code',
              date: '2026-05-23T12:05:01Z',
            },
          ],
        });
      }
      return makeJsonResponse({ emails: [] });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async (message, level) => logs.push({ message, level }),
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    persistRegistrationEmailState: async (_state, email) => {
      savedEmail = email;
    },
    setPersistentSettings: async (updates) => {
      persistedSettings = updates;
    },
    random: () => randomValues.shift() ?? 0,
    sleepWithStop: async () => {},
  });

  const email = await provider.fetchOutlookEmailAddress(state);

  assert.equal(email, 'clean@hotmail.com');
  assert.equal(savedEmail, 'clean@hotmail.com');
  assert.deepEqual(persistedSettings, {
    outlookEmailUsedAddresses: ['used@hotmail.com', 'clean@hotmail.com'],
  });
  assert.equal(checkedEmails[0], 'dirty@hotmail.com');
  assert.equal(checkedEmails.at(-1), 'clean@hotmail.com');
  assert.ok(checkedEmails.includes('clean@hotmail.com'));
  assert.ok(logs.some((item) => item.message.includes('跳过 dirty@hotmail.com')));
  assert.ok(logs.some((item) => item.message.includes('已随机取用无 OpenAI 邮件的邮箱 clean@hotmail.com')));
});

test('outlookEmail provider rejects a pool when every account has OpenAI mail', async () => {
  const state = {
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/accounts') {
      return makeJsonResponse({
        accounts: [
          { id: 'first', email: 'first@hotmail.com' },
          { id: 'second', email: 'second@hotmail.com' },
        ],
      });
    }
    if (parsed.pathname === '/api/external/emails') {
      return makeJsonResponse({
        emails: [
          {
            id: 'openai-old',
            subject: '你的 ChatGPT 临时验证码',
            from: 'noreply@tm.openai.com',
            body_preview: 'OpenAI ChatGPT code',
            date: '2026-05-23T12:05:01Z',
          },
        ],
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    random: () => 0,
    sleepWithStop: async () => {},
  });

  await assert.rejects(
    () => provider.fetchOutlookEmailAddress(state),
    /没有找到无 OpenAI 邮件的可用邮箱/
  );
});

test('outlookEmail provider falls back when all folder is unsupported during clean check', async () => {
  const state = {
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const requestedFolders = [];
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/accounts') {
      return makeJsonResponse({
        accounts: [{ id: 'clean', email: 'clean@hotmail.com' }],
      });
    }
    if (parsed.pathname === '/api/external/emails') {
      const folder = parsed.searchParams.get('folder');
      requestedFolders.push(folder);
      if (folder === 'all') {
        return {
          ok: false,
          status: 400,
          async text() {
            return JSON.stringify({ message: 'unsupported folder all' });
          },
        };
      }
      return makeJsonResponse({ emails: [] });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    random: () => 0,
    sleepWithStop: async () => {},
  });

  const email = await provider.fetchOutlookEmailAddress(state);

  assert.equal(email, 'clean@hotmail.com');
  assert.deepEqual(requestedFolders.slice(0, 3), ['all', 'inbox', 'junkemail']);
});

test('outlookEmail provider confirms suspicious list rows before rejecting an account', async () => {
  const state = {
    emailGenerator: 'outlook-email',
    mailProvider: 'outlook-email',
    outlookEmailApiKey: 'test-key',
    outlookEmailBaseUrl: 'http://mail.test',
  };
  const detailRequests = [];
  const fetchImpl = async (url) => {
    const parsed = new URL(url);
    if (parsed.pathname === '/api/external/accounts') {
      return makeJsonResponse({
        accounts: [{ id: 'clean', email: 'clean@hotmail.com' }],
      });
    }
    if (parsed.pathname === '/api/external/emails') {
      if (parsed.searchParams.get('keyword') === 'openai') {
        return makeJsonResponse({
          emails: [
            {
              id: 'marketing-1',
              id_mode: 'imap',
              subject: 'Newsletter',
              from: 'news@example.com',
              body_preview: '',
              date: '2026-05-23T12:05:01Z',
            },
          ],
        });
      }
      return makeJsonResponse({ emails: [] });
    }
    if (parsed.pathname.includes('/api/external/email/')) {
      detailRequests.push(parsed.pathname);
      return makeJsonResponse({
        email: {
          id: 'marketing-1',
          subject: 'Newsletter',
          from: 'news@example.com',
          body: '<html><body>Ordinary newsletter body</body></html>',
          date: '2026-05-23T12:05:01Z',
        },
      });
    }
    throw new Error(`unexpected request: ${url}`);
  };
  const provider = globalThis.MultiPageBackgroundOutlookEmailProvider.createOutlookEmailProvider({
    addLog: async () => {},
    buildOutlookEmailHeaders: outlookEmailUtils.buildOutlookEmailHeaders,
    fetchImpl,
    getState: async () => state,
    joinOutlookEmailUrl: outlookEmailUtils.joinOutlookEmailUrl,
    normalizeOutlookEmailAccounts: outlookEmailUtils.normalizeOutlookEmailAccounts,
    normalizeOutlookEmailAddress: outlookEmailUtils.normalizeOutlookEmailAddress,
    normalizeOutlookEmailBaseUrl: outlookEmailUtils.normalizeOutlookEmailBaseUrl,
    normalizeOutlookEmailMailApiDetail: outlookEmailUtils.normalizeOutlookEmailMailApiDetail,
    normalizeOutlookEmailMailApiMessages: outlookEmailUtils.normalizeOutlookEmailMailApiMessages,
    pickVerificationMessageWithTimeFallback: hotmailUtils.pickVerificationMessageWithTimeFallback,
    random: () => 0,
    sleepWithStop: async () => {},
  });

  const email = await provider.fetchOutlookEmailAddress(state);

  assert.equal(email, 'clean@hotmail.com');
  assert.equal(detailRequests.length, 1);
});

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
  assert.ok(logs.some((item) => item.message.includes('邮件详情找到验证码')));
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

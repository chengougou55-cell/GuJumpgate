const assert = require('node:assert/strict');
const test = require('node:test');

const hotmailUtils = require('../hotmail-utils.js');

globalThis.MultiPageOpenAiMailRules = null;
require('../flows/openai/mail-rules.js');
const openAiMailRules = globalThis.MultiPageOpenAiMailRules.createOpenAiMailRules({
  getHotmailVerificationRequestTimestamp: () => Date.UTC(2026, 4, 23, 12, 4, 30),
});

test('OpenAI signup mail rule ignores login-code mail and picks signup verification code', () => {
  const state = {
    email: 'kellyvicki8809@hotmail.com',
    mailProvider: 'outlook-email',
  };
  const pollPayload = openAiMailRules.buildVerificationPollPayload(4, state);
  const messages = [
    {
      id: 'login-code',
      subject: '你的临时 ChatGPT 登录代码',
      from: { emailAddress: { address: 'ChatGPT <noreply@tm1.openai.com>' } },
      bodyPreview: '你的临时 ChatGPT 登录代码 输入此代码以登录：895185',
      receivedDateTime: '2026-05-23T12:05:43.000Z',
    },
    {
      id: 'signup-code',
      subject: '你的 ChatGPT 临时验证码',
      from: { emailAddress: { address: 'noreply@tm.openai.com' } },
      bodyPreview: '你的 ChatGPT 临时验证码 输入此临时验证码以继续：220655 如果并非你本人尝试创建 ChatGPT 帐户，请忽略此电子邮件。',
      receivedDateTime: '2026-05-23T12:05:01.000Z',
    },
  ];

  const matchResult = hotmailUtils.pickVerificationMessageWithTimeFallback(messages, {
    afterTimestamp: pollPayload.filterAfterTimestamp,
    senderFilters: pollPayload.senderFilters,
    subjectFilters: pollPayload.subjectFilters,
    requiredKeywords: pollPayload.requiredKeywords,
    requiredAnyKeywords: pollPayload.requiredAnyKeywords,
    codePatterns: pollPayload.codePatterns,
    preferredSubjectFilters: pollPayload.preferredSubjectFilters,
    preferredKeywords: pollPayload.preferredKeywords,
  });

  assert.equal(matchResult.match.code, '220655');
  assert.equal(matchResult.match.message.id, 'signup-code');
});

test('OpenAI signup mail rule accepts fresh login-code subject when it is the current code mail', () => {
  const pollPayload = openAiMailRules.buildVerificationPollPayload(4, {
    email: 'kellyvicki8809@hotmail.com',
    mailProvider: 'outlook-email',
  });
  const matchResult = hotmailUtils.pickVerificationMessageWithTimeFallback([
    {
      id: 'fresh-login-code',
      subject: '你的临时 ChatGPT 登录代码',
      from: { emailAddress: { address: 'ChatGPT <noreply@tm1.openai.com>' } },
      bodyPreview: '你的临时 ChatGPT 登录代码 输入此代码以登录：778899',
      receivedDateTime: '2026-05-23T12:21:53.000Z',
    },
  ], {
    afterTimestamp: Date.UTC(2026, 4, 23, 12, 20, 58),
    senderFilters: pollPayload.senderFilters,
    subjectFilters: pollPayload.subjectFilters,
    requiredKeywords: pollPayload.requiredKeywords,
    requiredAnyKeywords: pollPayload.requiredAnyKeywords,
    codePatterns: pollPayload.codePatterns,
    preferredSubjectFilters: pollPayload.preferredSubjectFilters,
    preferredKeywords: pollPayload.preferredKeywords,
    disableTimeFallback: true,
  });

  assert.equal(matchResult.match.code, '778899');
  assert.equal(matchResult.match.message.id, 'fresh-login-code');
});

test('mail matcher can disable time fallback so old codes are not reused', () => {
  const matchResult = hotmailUtils.pickVerificationMessageWithTimeFallback([
    {
      id: 'old-signup-code',
      subject: '你的 ChatGPT 临时验证码',
      from: { emailAddress: { address: 'noreply@tm.openai.com' } },
      bodyPreview: '输入此临时验证码以继续：220655',
      receivedDateTime: '2026-05-23T12:05:01.000Z',
    },
  ], {
    afterTimestamp: Date.UTC(2026, 4, 23, 12, 20, 58),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码'],
    requiredKeywords: ['chatgpt', '验证码'],
    codePatterns: openAiMailRules.buildVerificationPollPayload(4, { email: 'kellyvicki8809@hotmail.com' }).codePatterns,
    disableTimeFallback: true,
  });

  assert.equal(matchResult.match, null);
  assert.equal(matchResult.usedTimeFallback, false);
});

test('OpenAI mail rule rejects non-OpenAI verification code mail in the same window', () => {
  const pollPayload = openAiMailRules.buildVerificationPollPayload(4, {
    email: 'kellyvicki8809@hotmail.com',
    mailProvider: 'outlook-email',
  });
  const matchResult = hotmailUtils.pickVerificationMessageWithTimeFallback([
    {
      id: 'other-service-code',
      subject: '邮箱验证码',
      from: { emailAddress: { address: 'security@example-service.test' } },
      bodyPreview: '你的验证码是 334455，请在 10 分钟内使用。',
      receivedDateTime: '2026-05-23T12:21:53.000Z',
    },
  ], {
    afterTimestamp: Date.UTC(2026, 4, 23, 12, 20, 58),
    senderFilters: pollPayload.senderFilters,
    subjectFilters: pollPayload.subjectFilters,
    requiredKeywords: pollPayload.requiredKeywords,
    requiredAnyKeywords: pollPayload.requiredAnyKeywords,
    codePatterns: pollPayload.codePatterns,
    disableTimeFallback: true,
  });

  assert.equal(matchResult.match, null);
});

test('mail matcher rejects undated mail when timestamp is required', () => {
  const matchResult = hotmailUtils.pickVerificationMessageWithTimeFallback([
    {
      id: 'undated-code',
      subject: '你的临时 ChatGPT 登录代码',
      from: { emailAddress: { address: 'noreply@tm.openai.com' } },
      bodyPreview: '输入此代码以登录：112233',
      receivedDateTime: '',
    },
  ], {
    afterTimestamp: Date.UTC(2026, 4, 23, 12, 20, 58),
    senderFilters: ['openai', 'noreply'],
    subjectFilters: ['验证码', '代码'],
    requiredKeywords: ['chatgpt', '代码'],
    codePatterns: openAiMailRules.buildVerificationPollPayload(4, { email: 'kellyvicki8809@hotmail.com' }).codePatterns,
    disableTimeFallback: true,
    requireReceivedTimestamp: true,
  });

  assert.equal(matchResult.match, null);
});

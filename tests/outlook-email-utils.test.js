const assert = require('node:assert/strict');
const test = require('node:test');

const outlookEmailUtils = require('../outlook-email-utils.js');
const hotmailUtils = require('../hotmail-utils.js');

test('outlookEmail utils normalize external accounts from API payload', () => {
  const accounts = outlookEmailUtils.normalizeOutlookEmailAccounts({
    accounts: [
      {
        id: 305,
        email: 'KellyVicki8809@HOTMAIL.com',
        provider: 'outlook',
        group_name: '0522-hot',
      },
    ],
  });

  assert.equal(accounts.length, 1);
  assert.equal(accounts[0].email, 'kellyvicki8809@hotmail.com');
  assert.equal(accounts[0].id, '305');
  assert.equal(accounts[0].groupName, '0522-hot');
});

test('outlookEmail utils normalize external emails for code matching', () => {
  const messages = outlookEmailUtils.normalizeOutlookEmailMailApiMessages({
    emails: [
      {
        id: 'AAMk-test',
        subject: 'Your ChatGPT code',
        from: {
          emailAddress: {
            address: 'noreply@tm.openai.com',
          },
        },
        body_preview: 'Your ChatGPT code is 123456',
        date: '2026-05-23T10:22:33Z',
        id_mode: 'imap',
      },
    ],
  }, 'inbox');

  assert.equal(messages.length, 1);
  assert.equal(messages[0].folder, 'inbox');
  assert.equal(messages[0].idMode, 'imap');
  assert.equal(messages[0].from.emailAddress.address, 'noreply@tm.openai.com');
  assert.equal(messages[0].receivedDateTime, '2026-05-23T10:22:33.000Z');
  assert.equal(hotmailUtils.extractVerificationCodeFromMessage(messages[0]), '123456');
});

test('outlookEmail utils normalize external email detail body for code matching', () => {
  const message = outlookEmailUtils.normalizeOutlookEmailMailApiDetail({
    email: {
      id: '13',
      subject: '你的临时 ChatGPT 登录代码',
      from: 'ChatGPT <noreply@tm1.openai.com>',
      body: '<html><body><p>输入此代码以登录：</p><div>778899</div></body></html>',
      date: '23-May-2026 08:21:53 -0400',
    },
  }, 'inbox');

  assert.equal(message.id, '13');
  assert.equal(message.bodyPreview.includes('<html>'), false);
  assert.equal(hotmailUtils.extractVerificationCodeFromMessage(message), '778899');
});

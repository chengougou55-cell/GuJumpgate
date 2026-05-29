const assert = require('node:assert/strict');
const test = require('node:test');

const {
  getActivationStrategy,
  isEmailVerificationResendAction,
} = require('../content/activation-utils.js');

test('email-verification resend action dispatches click without form submit', () => {
  assert.deepEqual(getActivationStrategy({
    tagName: 'button',
    type: 'submit',
    hasForm: true,
    pathname: '/email-verification',
    text: '重新发送电子邮件',
  }), { method: 'dispatchClick' });

  assert.deepEqual(getActivationStrategy({
    tagName: 'button',
    type: 'submit',
    hasForm: true,
    pathname: '/email-verification',
    name: 'intent',
    value: 'resend',
  }), { method: 'dispatchClick' });
});

test('email-verification non-resend submit still uses requestSubmit', () => {
  assert.deepEqual(getActivationStrategy({
    tagName: 'button',
    type: 'submit',
    hasForm: true,
    pathname: '/email-verification',
    text: '继续',
  }), { method: 'requestSubmit' });
});

test('email-verification resend action detection accepts localized labels', () => {
  assert.equal(isEmailVerificationResendAction({ text: '重新发送验证码' }), true);
  assert.equal(isEmailVerificationResendAction({ text: 'Resend email' }), true);
  assert.equal(isEmailVerificationResendAction({ text: '继续' }), false);
});

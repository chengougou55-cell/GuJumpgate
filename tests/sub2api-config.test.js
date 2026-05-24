const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

test('SUB2API private password is not committed in tracked source', () => {
  assert.match(readProjectFile('.gitignore'), /^\/background\/local-settings\.private\.js$/m);

  const privateSettingsPath = path.join(repoRoot, 'background/local-settings.private.js');
  if (!fs.existsSync(privateSettingsPath)) {
    return;
  }

  const privateSettings = fs.readFileSync(privateSettingsPath, 'utf8');
  const sensitivePasswordMatch = privateSettings.match(/sub2apiPassword:\s*'([^']+)'/);
  if (!sensitivePasswordMatch) {
    return;
  }

  const sensitivePassword = sensitivePasswordMatch[1];
  const trackedFiles = [
    'background.js',
    'background/local-settings.js',
    'background/sub2api-api.js',
    'content/sub2api-panel.js',
    'sidepanel/sidepanel.js',
    'tests/sub2api-config.test.js',
  ];

  for (const file of trackedFiles) {
    assert.equal(readProjectFile(file).includes(sensitivePassword), false, `${file} should not contain private SUB2API password`);
  }
});

test('local SUB2API settings override stale persisted and session values', () => {
  const backgroundScript = readProjectFile('background.js');
  const helperMatch = backgroundScript.match(/function buildLocalSub2ApiSettingsOverride[\s\S]*?\n}\n\nfunction buildPersistentSettingsPayload/);
  assert.ok(helperMatch, 'expected buildLocalSub2ApiSettingsOverride helper');
  assert.match(backgroundScript, /importScripts\('background\/local-settings\.js'\);/);
  assert.match(backgroundScript, /importScripts\('background\/local-settings\.private\.js'\);/);
  assert.match(backgroundScript, /\.\.\.state,\n\s+\.\.\.localSub2ApiSettings,\n\s+accountRunHistory,/);

  const sandbox = {
    LOCAL_SUB2API_SETTING_KEYS: Object.freeze(['sub2apiUrl', 'sub2apiEmail', 'sub2apiPassword']),
    normalizePersistentSettingValue: (_key, value) => String(value || '').trim(),
    normalizePanelMode: (value) => value,
    normalizeLocalCpaStep9Mode: (value) => value,
    normalizeSub2ApiUrl: (value) => String(value || '').trim(),
    normalizeSub2ApiGroupNames: (value) => Array.isArray(value) ? value : [],
    normalizeSub2ApiAccountPriority: (value) => value,
    normalizeIpProxyServiceProfiles: (value) => value,
  };
  vm.runInNewContext(helperMatch[0].replace(/\nfunction buildPersistentSettingsPayload[\s\S]*$/, ''), sandbox);

  const override = sandbox.buildLocalSub2ApiSettingsOverride({
    sub2api: {
      sub2apiUrl: ' http://156.239.40.207:18080/admin/accounts ',
      sub2apiEmail: ' admin@example.test ',
      sub2apiPassword: 'test-private-password',
    },
  });
  const mergedState = {
    sub2apiUrl: 'http://old.example/admin/accounts',
    sub2apiEmail: 'old@example.test',
    sub2apiPassword: 'old-test-password',
    ...override,
  };

  assert.deepEqual(JSON.parse(JSON.stringify(override)), {
    sub2apiUrl: 'http://156.239.40.207:18080/admin/accounts',
    sub2apiEmail: 'admin@example.test',
    sub2apiPassword: 'test-private-password',
  });
  assert.equal(mergedState.sub2apiEmail, 'admin@example.test');
  assert.equal(mergedState.sub2apiPassword, 'test-private-password');
});

test('SUB2API login reports 2FA requirement explicitly', async () => {
  globalThis.MultiPageBackgroundSub2ApiApi = null;
  require('../background/sub2api-api.js');

  const api = globalThis.MultiPageBackgroundSub2ApiApi.createSub2ApiApi({
    normalizeSub2ApiUrl: (value) => value,
    fetchImpl: async () => ({
      ok: true,
      status: 200,
      async text() {
        return JSON.stringify({
          code: 0,
          data: {
            requires_2fa: true,
            temp_token: 'temporary-token',
            user_email_masked: 'c***5@gmail.com',
          },
        });
      },
    }),
  });

  await assert.rejects(
    () => api.loginSub2Api({
      sub2apiUrl: 'http://156.239.40.207:18080/admin/accounts',
      sub2apiEmail: 'admin@example.test',
      sub2apiPassword: 'test-private-password',
    }),
    /2FA\/TOTP/
  );
});

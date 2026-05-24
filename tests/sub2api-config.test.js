const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');

function readProjectFile(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), 'utf8');
}

function loadSub2ApiApiModule() {
  globalThis.MultiPageBackgroundSub2ApiApi = null;
  const modulePath = require.resolve('../background/sub2api-api.js');
  delete require.cache[modulePath];
  require('../background/sub2api-api.js');
  return globalThis.MultiPageBackgroundSub2ApiApi;
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

  const normalizeMatch = backgroundScript.match(/function normalizeSub2ApiGroupNames[\s\S]*?\n}\n\nfunction normalizeSub2ApiAccountPriority/);
  assert.ok(normalizeMatch, 'expected SUB2API group normalization helpers');
  const sandbox = {
    LOCAL_SUB2API_SETTING_KEYS: Object.freeze(['sub2apiUrl', 'sub2apiEmail', 'sub2apiPassword', 'sub2apiGroupName', 'sub2apiGroupNames']),
    XIAOHONGSHU_SUB2API_GROUP_NAME: 'xiaohongshu',
    DEFAULT_SUB2API_GROUP_NAME: 'codex',
    normalizePanelMode: (value) => value,
    normalizeLocalCpaStep9Mode: (value) => value,
    normalizeSub2ApiUrl: (value) => String(value || '').trim(),
    normalizeSub2ApiAccountPriority: (value) => value,
    normalizeIpProxyServiceProfiles: (value) => value,
  };
  vm.runInNewContext(`
${normalizeMatch[0].replace(/\nfunction normalizeSub2ApiAccountPriority[\s\S]*$/, '')}
function normalizePersistentSettingValue(key, value) {
  switch (key) {
    case 'sub2apiUrl':
      return normalizeSub2ApiUrl(value);
    case 'sub2apiEmail':
      return String(value || '').trim();
    case 'sub2apiPassword':
      return String(value || '');
    case 'sub2apiGroupName':
      return isXiaohongshuSub2ApiGroupName(value) ? DEFAULT_SUB2API_GROUP_NAME : String(value || '').trim();
    case 'sub2apiGroupNames':
      return normalizePublicSub2ApiGroupNames(value);
    default:
      return value;
  }
}
`, sandbox);
  vm.runInNewContext(helperMatch[0].replace(/\nfunction buildPersistentSettingsPayload[\s\S]*$/, ''), sandbox);

  const override = sandbox.buildLocalSub2ApiSettingsOverride({
    sub2api: {
      sub2apiUrl: ' http://156.239.40.207:18080/admin/accounts ',
      sub2apiEmail: ' admin@example.test ',
      sub2apiPassword: 'test-private-password',
      sub2apiGroupName: 'xiaohongshu',
      sub2apiGroupNames: ['openai-plus', 'xiaohongshu'],
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
    sub2apiGroupName: 'codex',
    sub2apiGroupNames: ['openai-plus'],
  });
  assert.equal(mergedState.sub2apiEmail, 'admin@example.test');
  assert.equal(mergedState.sub2apiPassword, 'test-private-password');
  assert.equal(mergedState.sub2apiGroupName, 'codex');
  assert.deepEqual(JSON.parse(JSON.stringify(mergedState.sub2apiGroupNames)), ['openai-plus']);
});

test('SUB2API login reports 2FA requirement explicitly', async () => {
  const apiModule = loadSub2ApiApiModule();
  const api = apiModule.createSub2ApiApi({
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

test('SUB2API xiaohongshu group is only selected by explicit mode flag', async () => {
  const apiModule = loadSub2ApiApiModule();

  async function runImport(stateOverrides = {}) {
    const seenGroupQueries = [];
    const seenImportBodies = [];
    const api = apiModule.createSub2ApiApi({
      normalizeSub2ApiUrl: (value) => value,
      XIAOHONGSHU_SUB2API_GROUP_NAME: 'xiaohongshu',
      fetchImpl: async (url, options = {}) => {
        const parsed = new URL(url);
        const body = options.body ? JSON.parse(options.body) : null;
        if (parsed.pathname === '/api/v1/auth/login') {
          return {
            ok: true,
            status: 200,
            async text() {
              return JSON.stringify({ code: 0, data: { access_token: 'admin-token' } });
            },
          };
        }
        if (parsed.pathname === '/api/v1/admin/groups/all') {
          seenGroupQueries.push(parsed.pathname);
          return {
            ok: true,
            status: 200,
            async text() {
              return JSON.stringify({
                code: 0,
                data: [
                  { id: 11, name: 'codex', platform: 'openai' },
                  { id: 22, name: 'openai-plus', platform: 'openai' },
                  { id: 33, name: 'xiaohongshu', platform: 'openai' },
                ],
              });
            },
          };
        }
        if (parsed.pathname === '/api/v1/admin/accounts/import/codex-session') {
          seenImportBodies.push(body);
          return {
            ok: true,
            status: 200,
            async text() {
              return JSON.stringify({ code: 0, data: { total: 1, created: 1, updated: 0, skipped: 0, failed: 0 } });
            },
          };
        }
        throw new Error(`unexpected SUB2API request: ${parsed.pathname}`);
      },
    });

    await api.importCurrentChatGptSession({
      sub2apiUrl: 'http://sub2api.example/admin/accounts',
      sub2apiEmail: 'admin@example.test',
      sub2apiPassword: 'test-private-password',
      sub2apiGroupName: 'codex',
      sub2apiAccountPriority: 1,
      accessToken: 'header.payload.signature',
      ...stateOverrides,
    });

    assert.equal(seenGroupQueries.length, 1);
    assert.equal(seenImportBodies.length, 1);
    return seenImportBodies[0].group_ids;
  }

  assert.deepEqual(await runImport(), [11]);
  assert.deepEqual(await runImport({
    xiaohongshuAccessToken: 'residual-token',
    directCheckoutAccessToken: 'residual-token',
  }), [11]);
  assert.deepEqual(await runImport({
    sub2apiGroupName: 'xiaohongshu',
    xiaohongshuAccessToken: 'residual-token',
    directCheckoutAccessToken: 'residual-token',
  }), [11]);
  assert.deepEqual(await runImport({
    xiaohongshuModeEnabled: true,
    xiaohongshuAccessToken: 'mode-token',
    directCheckoutAccessToken: 'mode-token',
  }), [33]);
});

test('SUB2API xiaohongshu OAuth callback ignores stale stored group ids', async () => {
  const apiModule = loadSub2ApiApiModule();
  const seenCreateBodies = [];
  let groupLookupCount = 0;
  const api = apiModule.createSub2ApiApi({
    normalizeSub2ApiUrl: (value) => value,
    XIAOHONGSHU_SUB2API_GROUP_NAME: 'xiaohongshu',
    fetchImpl: async (url, options = {}) => {
      const parsed = new URL(url);
      const body = options.body ? JSON.parse(options.body) : null;
      if (parsed.pathname === '/api/v1/auth/login') {
        return {
          ok: true,
          status: 200,
          async text() {
            return JSON.stringify({ code: 0, data: { access_token: 'admin-token' } });
          },
        };
      }
      if (parsed.pathname === '/api/v1/admin/groups/all') {
        groupLookupCount += 1;
        return {
          ok: true,
          status: 200,
          async text() {
            return JSON.stringify({
              code: 0,
              data: [
                { id: 11, name: 'codex', platform: 'openai' },
                { id: 33, name: 'xiaohongshu', platform: 'openai' },
              ],
            });
          },
        };
      }
      if (parsed.pathname === '/api/v1/admin/openai/exchange-code') {
        return {
          ok: true,
          status: 200,
          async text() {
            return JSON.stringify({ code: 0, data: { access_token: 'oauth-token', email: 'user@example.test' } });
          },
        };
      }
      if (parsed.pathname === '/api/v1/admin/accounts') {
        seenCreateBodies.push(body);
        return {
          ok: true,
          status: 200,
          async text() {
            return JSON.stringify({ code: 0, data: { id: 99 } });
          },
        };
      }
      throw new Error(`unexpected SUB2API request: ${parsed.pathname}`);
    },
  });

  await api.submitOpenAiCallback({
    sub2apiUrl: 'http://sub2api.example/admin/accounts',
    sub2apiEmail: 'admin@example.test',
    sub2apiPassword: 'test-private-password',
    localhostUrl: 'http://localhost:8000/auth/callback?code=callback-code&state=oauth-state',
    sub2apiSessionId: 'session-1',
    sub2apiOAuthState: 'oauth-state',
    sub2apiGroupId: 11,
    sub2apiGroupIds: [11],
    sub2apiGroupName: 'codex',
    sub2apiAccountPriority: 1,
    xiaohongshuModeEnabled: true,
  });

  assert.equal(groupLookupCount, 1);
  assert.equal(seenCreateBodies.length, 1);
  assert.deepEqual(seenCreateBodies[0].group_ids, [33]);
});

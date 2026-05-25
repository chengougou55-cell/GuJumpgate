const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const repoRoot = path.resolve(__dirname, '..');

function loadTabRuntime() {
  const sandbox = {};
  vm.runInNewContext(
    fs.readFileSync(path.join(repoRoot, 'background/tab-runtime.js'), 'utf8'),
    sandbox
  );
  return sandbox.MultiPageBackgroundTabRuntime;
}

test('getTabId clears stale registry entries for closed tabs', async () => {
  const tabRuntimeModule = loadTabRuntime();
  let state = {
    automationWindowId: 1,
    tabRegistry: {
      'signup-page': {
        tabId: 1274206602,
        ready: true,
        windowId: 1,
      },
    },
  };
  const patches = [];
  const runtime = tabRuntimeModule.createTabRuntime({
    chrome: {
      tabs: {
        get: async () => {
          throw new Error('No tab with id: 1274206602.');
        },
      },
    },
    getState: async () => state,
    setState: async (patch) => {
      patches.push(patch);
      state = { ...state, ...patch };
    },
  });

  const tabId = await runtime.getTabId('signup-page');

  assert.equal(tabId, null);
  assert.equal(state.tabRegistry['signup-page'], null);
  assert.equal(patches.length, 1);
  assert.equal(patches[0].tabRegistry['signup-page'], null);
});

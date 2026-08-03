'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { dispatchActions } = require('../src/03-action-management.js');

function makeFakeThread() {
  return {
    getId: function () {
      return 'thread-123';
    },
  };
}

// --- disabled action is fully skipped ---------------------------------------

test('dispatchActions: config.enabled === false skips the action before appliesTo/run', () => {
  const fakeAction = {
    name: 'disabled-action',
    config: { enabled: false },
    appliesTo: function () {
      throw new Error('appliesTo must not be called for a disabled action');
    },
    run: function () {
      throw new Error('run must not be called for a disabled action');
    },
  };

  const result = dispatchActions(makeFakeThread(), [fakeAction]);

  assert.equal(result.hadError, false);
});

// --- enabled: true runs normally ---------------------------------------------

test('dispatchActions: config.enabled === true runs appliesTo and run normally', () => {
  let appliesToCalls = 0;
  let runCalls = 0;
  const fakeAction = {
    name: 'enabled-action',
    config: { enabled: true },
    appliesTo: function () {
      appliesToCalls += 1;
      return true;
    },
    run: function () {
      runCalls += 1;
    },
  };

  dispatchActions(makeFakeThread(), [fakeAction]);

  assert.equal(appliesToCalls, 1);
  assert.equal(runCalls, 1);
});

// --- enabled key absent (back-compat) runs normally --------------------------

test('dispatchActions: config with no enabled key runs normally (back-compat)', () => {
  let appliesToCalls = 0;
  let runCalls = 0;
  const fakeAction = {
    name: 'no-enabled-key-action',
    config: {},
    appliesTo: function () {
      appliesToCalls += 1;
      return true;
    },
    run: function () {
      runCalls += 1;
    },
  };

  dispatchActions(makeFakeThread(), [fakeAction]);

  assert.equal(appliesToCalls, 1);
  assert.equal(runCalls, 1);
});

// --- appliesTo === false short-circuits run for an enabled action ------------

test('dispatchActions: appliesTo returning false short-circuits run (existing contract preserved)', () => {
  let appliesToCalls = 0;
  let runCalls = 0;
  const fakeAction = {
    name: 'no-match-action',
    config: { enabled: true },
    appliesTo: function () {
      appliesToCalls += 1;
      return false;
    },
    run: function () {
      runCalls += 1;
    },
  };

  dispatchActions(makeFakeThread(), [fakeAction]);

  assert.equal(appliesToCalls, 1);
  assert.equal(runCalls, 0);
});

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { ICS_CALENDAR_ACTION } = require('../src/05-action-ics-import.js');
const { ICS_ACTION_CONFIG } = require('../src/05-action-cfg-ics-import.js');
const { BOOKING_MANAGEMENT_ACTION } = require('../src/06-action-booking-com-management.js');
const { BOOKING_ACTION_CONFIG } = require('../src/06-action-cfg-booking-com-management.js');

// --- Action/config file split: getter-based config, load-order-independent ---
//
// Owner-requested refactor: split each action's *_CONFIG object out of its
// action file into its own sibling "{n}-action-cfg-{name}.js" file. This is
// only safe because each action descriptor's `config` property is an ES6
// GETTER (not a plain literal reference) — a getter is NOT evaluated when
// the object literal is constructed, only when something actually reads
// `action.config` (dispatchActions' enable check, the failure-notification
// builder), which happens inside function bodies long after every project
// file has finished loading. This is the same "resolve lazily, at the
// point of use" pattern already proven in this codebase for
// BOOKING_LANGUAGE_PACKS/getBookingLabels, applied here to make the config
// split safe regardless of which of the two files (action vs config) loads
// first alphabetically — unlike a plain literal `config: X_CONFIG`
// property, which would need X_CONFIG to already exist at
// object-construction time (GAS's shared-global-scope load-order category
// of bug this codebase has already been bitten by twice).

test('ICS_CALENDAR_ACTION.config is the SAME object reference as ICS_ACTION_CONFIG (getter identity, not a copy)', () => {
  assert.equal(ICS_CALENDAR_ACTION.config, ICS_ACTION_CONFIG);
});

test('BOOKING_MANAGEMENT_ACTION.config is the SAME object reference as BOOKING_ACTION_CONFIG (getter identity, not a copy)', () => {
  assert.equal(BOOKING_MANAGEMENT_ACTION.config, BOOKING_ACTION_CONFIG);
});

// --- Post quick-260726-spr update: individual FIELDS are now ES6 getters too ---
//
// The two tests below originally proved "the getter reads the live shared
// object, not a snapshot" by directly ASSIGNING to
// ICS_ACTION_CONFIG.notifyOnFailure / BOOKING_ACTION_CONFIG.notifyOnFailure.
// Since quick-260726-spr (Script Properties live-config override), every
// individual field on these config objects is ALSO an ES6 getter (resolving
// through PropertiesService-or-code-default — see src/01-setup.js's
// class-level JSDoc), so those fields no longer have a setter at all —
// `ICS_ACTION_CONFIG.notifyOnFailure = ...` now throws
// `TypeError: Cannot set property notifyOnFailure of #<Object> which has
// only a getter` under this file's 'use strict' mode. This is the correct,
// intended consequence of that refactor (a field's real value now lives in
// Script Properties, not as a directly-assignable JS property) — not a bug
// to route around.
//
// Rewritten below to prove the SAME underlying claim (config.notifyOnFailure
// is read live, never cached/snapshotted) through the ACTUAL mechanism that
// now makes it live: a minimal fake `global.PropertiesService`, exercising
// the real Script-Properties-or-code-default resolution path end-to-end.
// This is a strictly stronger regression proof than the original direct
// assignment ever was.

function withFakePropertiesService(initialProperties, fn) {
  const store = Object.assign({}, initialProperties);

  global.PropertiesService = {
    getScriptProperties: function () {
      return {
        getProperty: function (key) {
          return Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null;
        },
      };
    },
  };

  try {
    fn(store);
  } finally {
    delete global.PropertiesService;
  }
}

test('ICS_CALENDAR_ACTION.config.notifyOnFailure is read LIVE through Script Properties (the getter reads the live store, not a snapshot)', () => {
  // Baseline, PropertiesService absent (Node/test default): code default.
  assert.equal(ICS_CALENDAR_ACTION.config.notifyOnFailure, true);

  withFakePropertiesService({}, function (store) {
    // Still unset -> still the code default, even with PropertiesService now present.
    assert.equal(ICS_CALENDAR_ACTION.config.notifyOnFailure, true);

    // Set the live property AFTER already reading the getter once above —
    // proves it re-reads live each time, it did not cache the first read.
    store['05-action-ics-NOTIFY_ON_FAILURE'] = 'false';
    assert.equal(ICS_CALENDAR_ACTION.config.notifyOnFailure, false);
  });

  // PropertiesService removed again -> back to the code default, confirming
  // this test's fake store leaked no state into any other test.
  assert.equal(ICS_CALENDAR_ACTION.config.notifyOnFailure, true);
});

test('BOOKING_MANAGEMENT_ACTION.config.notifyOnFailure is read LIVE through Script Properties (the getter reads the live store, not a snapshot)', () => {
  assert.equal(BOOKING_MANAGEMENT_ACTION.config.notifyOnFailure, true);

  withFakePropertiesService({}, function (store) {
    assert.equal(BOOKING_MANAGEMENT_ACTION.config.notifyOnFailure, true);

    store['06-action-booking-com-NOTIFY_ON_FAILURE'] = 'false';
    assert.equal(BOOKING_MANAGEMENT_ACTION.config.notifyOnFailure, false);
  });

  assert.equal(BOOKING_MANAGEMENT_ACTION.config.notifyOnFailure, true);
});

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { resolveIcsCalendarId } = require('../src/05-action-ics-import.js');
const {
  resolveBookingCalendarId,
  listCalendarEventsPaginated,
  findEventByConfirmationTag,
  findOrTagMatchingEvent,
} = require('../src/06-action-booking-com-management.js');

// --- resolveIcsCalendarId ----------------------------------------------------
//
// Owner-requested feature: CONFIG.calendarId (src/01-setup.js) becomes a
// DEFAULT/fallback calendar; the ICS action can override it action-wide
// (ICS_ACTION_CONFIG.calendarId) and/or per-sender
// (ICS_ACTION_CONFIG.calendarIdBySender). Resolution order per
// ICS-carrying message, most-specific wins:
//   1. calendarIdBySender entry matching that message's sender
//   2. ICS_ACTION_CONFIG.calendarId (action-level override), if set
//   3. the global default (CONFIG.calendarId)
// Independent of importOnlyFrom (a separate allow-list gate for whether a
// message is processed at all) — this function is only ever consulted for
// messages that already passed that gate.

test('resolveIcsCalendarId: no mapping, no action override -> global default', () => {
  const config = { calendarId: null, calendarIdBySender: [] };
  assert.equal(resolveIcsCalendarId('someone@example.com', config, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('resolveIcsCalendarId: action override set, no matching mapping -> action override wins over global default', () => {
  const config = { calendarId: 'ACTION_CAL', calendarIdBySender: [] };
  assert.equal(resolveIcsCalendarId('someone@example.com', config, 'DEFAULT_CAL'), 'ACTION_CAL');
});

test('resolveIcsCalendarId: matching calendarIdBySender entry wins over BOTH the action override and the global default', () => {
  const config = {
    calendarId: 'ACTION_CAL',
    calendarIdBySender: [{ from: 'alice@example.com', calendarId: 'ALICE_CAL' }],
  };
  assert.equal(resolveIcsCalendarId('Alice <alice@example.com>', config, 'DEFAULT_CAL'), 'ALICE_CAL');
});

test('resolveIcsCalendarId: calendarIdBySender sender matching is case-insensitive (mirrors extractEmailAddress)', () => {
  const config = {
    calendarId: null,
    calendarIdBySender: [{ from: 'Alice@Example.com', calendarId: 'ALICE_CAL' }],
  };
  assert.equal(resolveIcsCalendarId('ALICE@EXAMPLE.COM', config, 'DEFAULT_CAL'), 'ALICE_CAL');
});

test('resolveIcsCalendarId: multiple matching calendarIdBySender entries, first match in list order wins', () => {
  const config = {
    calendarId: null,
    calendarIdBySender: [
      { from: 'bob@example.com', calendarId: 'FIRST_CAL' },
      { from: 'bob@example.com', calendarId: 'SECOND_CAL' },
    ],
  };
  assert.equal(resolveIcsCalendarId('bob@example.com', config, 'DEFAULT_CAL'), 'FIRST_CAL');
});

test('resolveIcsCalendarId: mapping present but sender does not match any entry, no action override -> falls through to global default', () => {
  const config = {
    calendarId: null,
    calendarIdBySender: [{ from: 'alice@example.com', calendarId: 'ALICE_CAL' }],
  };
  assert.equal(resolveIcsCalendarId('carol@example.com', config, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('resolveIcsCalendarId: mapping present but sender does not match any entry, action override set -> falls through to action override (not the global default)', () => {
  const config = {
    calendarId: 'ACTION_CAL',
    calendarIdBySender: [{ from: 'alice@example.com', calendarId: 'ALICE_CAL' }],
  };
  assert.equal(resolveIcsCalendarId('carol@example.com', config, 'DEFAULT_CAL'), 'ACTION_CAL');
});

test('resolveIcsCalendarId: shipped defaults (calendarId: null, calendarIdBySender: []) resolve to the global default, unchanged behavior', () => {
  const { ICS_ACTION_CONFIG } = require('../src/05-action-cfg-ics-import.js');
  assert.equal(resolveIcsCalendarId('anyone@example.com', ICS_ACTION_CONFIG, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

// --- resolveBookingCalendarId -------------------------------------------------
//
// Owner-requested feature: BOOKING_ACTION_CONFIG.calendarId — a single
// action-wide override (no per-sender map: booking.com only ever sends
// from one known address in practice). Same two-tier resolution as ICS's
// action-level override: config.calendarId || defaultCalendarId.

test('resolveBookingCalendarId: action override set -> used', () => {
  assert.equal(resolveBookingCalendarId({ calendarId: 'BOOKING_CAL' }, 'DEFAULT_CAL'), 'BOOKING_CAL');
});

test('resolveBookingCalendarId: action override unset (null) -> global default used', () => {
  assert.equal(resolveBookingCalendarId({ calendarId: null }, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('resolveBookingCalendarId: action override key missing entirely -> global default used', () => {
  assert.equal(resolveBookingCalendarId({}, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('resolveBookingCalendarId: shipped default (calendarId: null) resolves to the global default, unchanged behavior', () => {
  const { BOOKING_ACTION_CONFIG } = require('../src/06-action-cfg-booking-com-management.js');
  assert.equal(resolveBookingCalendarId(BOOKING_ACTION_CONFIG, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

// --- calendarId threading proof (booking's GAS-only Calendar call sites) ----
//
// listCalendarEventsPaginated/findEventByConfirmationTag/findOrTagMatchingEvent
// are GAS-only (Calendar global) and were not otherwise unit-tested before
// this change (see their own JSDoc — proven only by the live checkpoint).
// This does not newly unit-test their matching/patching BEHAVIOR; it proves
// the SIGNATURE change alone — that the passed-through calendarId parameter
// is what actually reaches Calendar.Events.list/patch, not a hardcoded
// value — using a minimal fake Calendar global (no real GAS dependency).

test('listCalendarEventsPaginated: forwards the given calendarId to Calendar.Events.list, not a hardcoded value', () => {
  const seenCalendarIds = [];
  global.Calendar = {
    Events: {
      list: function (calendarId) {
        seenCalendarIds.push(calendarId);
        return { items: [], nextPageToken: undefined };
      },
    },
  };

  try {
    listCalendarEventsPaginated({ timeMin: 'a', timeMax: 'b' }, 'CUSTOM_CAL');
    assert.deepEqual(seenCalendarIds, ['CUSTOM_CAL']);
  } finally {
    delete global.Calendar;
  }
});

test('findEventByConfirmationTag: forwards the given calendarId to Calendar.Events.list', () => {
  const seenCalendarIds = [];
  global.Calendar = {
    Events: {
      list: function (calendarId) {
        seenCalendarIds.push(calendarId);
        return { items: [], nextPageToken: undefined };
      },
    },
  };

  try {
    const result = findEventByConfirmationTag('a', 'b', '12345', 'CUSTOM_CAL');
    assert.equal(result, null);
    assert.deepEqual(seenCalendarIds, ['CUSTOM_CAL']);
  } finally {
    delete global.Calendar;
  }
});

test('findOrTagMatchingEvent: forwards the given calendarId to both list layers and to patch', () => {
  const seenListCalendarIds = [];
  let patchCalendarId = null;
  global.Calendar = {
    Events: {
      list: function (calendarId, params) {
        seenListCalendarIds.push(calendarId);
        if (params && params.privateExtendedProperty) {
          return { items: [], nextPageToken: undefined };
        }
        return {
          items: [
            {
              id: 'evt-1',
              summary: 'Hotel Foo',
              location: '',
              start: { date: '2026-09-11' },
              end: { date: '2026-09-14' },
            },
          ],
          nextPageToken: undefined,
        };
      },
      patch: function (resource, calendarId) {
        patchCalendarId = calendarId;
      },
    },
  };

  try {
    const checkIn = new Date(Date.UTC(2026, 8, 11));
    const checkOut = new Date(Date.UTC(2026, 8, 13));
    const result = findOrTagMatchingEvent('a', 'b', '99999', 'Hotel Foo', checkIn, checkOut, 'CUSTOM_CAL');
    assert.equal(result.id, 'evt-1');
    assert.deepEqual(seenListCalendarIds, ['CUSTOM_CAL', 'CUSTOM_CAL']);
    assert.equal(patchCalendarId, 'CUSTOM_CAL');
  } finally {
    delete global.Calendar;
  }
});

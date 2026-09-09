'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { orderThreadsForProcessing } = require('../src/02-main.js');

// --- orderThreadsForProcessing (debug/regiojet-cancel-not-deleted) -----------
//
// THE ORDERING DEFECT this helper exists to close.
//
// processEmails consumed GmailApp.search()'s result array in whatever order
// Gmail handed it back. That order is NOT DOCUMENTED anywhere in the Apps
// Script reference -- empirically it mirrors the Gmail UI, i.e. NEWEST FIRST.
//
// Every cancellation-capable action in this codebase (TRANSPORT_TICKETS_ACTION's
// cancelTransportTicketEvent, BOOKING_MANAGEMENT_ACTION's handleCancellation)
// needs the OPPOSITE: a cancellation email is by definition NEWER than the
// confirmation it cancels, so under newest-first the cancellation is dispatched
// BEFORE the event it is meant to delete has been created. It finds nothing,
// no-ops silently, and the confirmation -- processed second -- then creates an
// event that nothing will ever delete.
//
// Ordering oldest-first makes the pipeline CAUSAL: emails are processed in the
// order they actually arrived, so a confirmation is always seen before its own
// cancellation. The key is getLastMessageDate() (when the thread last became
// relevant), not the first message date.

function fakeThread(id, lastMessageDate) {
  return {
    getId: function () {
      return id;
    },
    getLastMessageDate: function () {
      return lastMessageDate;
    },
  };
}

function idsOf(threads) {
  return threads.map(function (thread) {
    return thread.getId();
  });
}

test('orderThreadsForProcessing: THE LOAD-BEARING CASE -- a newest-first array (as Gmail returns it) comes back oldest-first, so a confirmation thread is processed before its own cancellation thread', () => {
  const confirmation = fakeThread('confirmation', new Date('2026-09-08T13:05:00Z'));
  const cancellation = fakeThread('cancellation', new Date('2026-09-08T13:10:30Z'));

  // Gmail hands these back newest-first.
  const result = orderThreadsForProcessing([cancellation, confirmation]);

  assert.deepEqual(idsOf(result), ['confirmation', 'cancellation']);
});

test('orderThreadsForProcessing: an already-oldest-first array is left in that order (the helper is an ordering guarantee, not a reversal)', () => {
  const first = fakeThread('first', new Date('2026-09-01T08:00:00Z'));
  const second = fakeThread('second', new Date('2026-09-02T08:00:00Z'));
  const third = fakeThread('third', new Date('2026-09-03T08:00:00Z'));

  assert.deepEqual(idsOf(orderThreadsForProcessing([first, second, third])), ['first', 'second', 'third']);
});

test('orderThreadsForProcessing: sorts a shuffled array strictly by ascending last-message date', () => {
  const a = fakeThread('a', new Date('2026-09-03T08:00:00Z'));
  const b = fakeThread('b', new Date('2026-09-01T08:00:00Z'));
  const c = fakeThread('c', new Date('2026-09-05T08:00:00Z'));
  const d = fakeThread('d', new Date('2026-09-02T08:00:00Z'));

  assert.deepEqual(idsOf(orderThreadsForProcessing([a, b, c, d])), ['b', 'd', 'a', 'c']);
});

test('orderThreadsForProcessing: does not mutate the input array', () => {
  const older = fakeThread('older', new Date('2026-09-01T08:00:00Z'));
  const newer = fakeThread('newer', new Date('2026-09-02T08:00:00Z'));
  const input = [newer, older];

  orderThreadsForProcessing(input);

  assert.deepEqual(idsOf(input), ['newer', 'older']);
});

test('orderThreadsForProcessing: equal timestamps preserve their relative input order (stable)', () => {
  const sameInstant = new Date('2026-09-04T08:00:00Z');
  const first = fakeThread('first', sameInstant);
  const second = fakeThread('second', sameInstant);
  const third = fakeThread('third', sameInstant);

  assert.deepEqual(idsOf(orderThreadsForProcessing([first, second, third])), ['first', 'second', 'third']);
});

// --- boundary neighbours around the sort key --------------------------------

test('orderThreadsForProcessing: empty array in, empty array out', () => {
  assert.deepEqual(orderThreadsForProcessing([]), []);
});

test('orderThreadsForProcessing: null/undefined input returns an empty array, never throws', () => {
  assert.deepEqual(orderThreadsForProcessing(null), []);
  assert.deepEqual(orderThreadsForProcessing(undefined), []);
});

test('orderThreadsForProcessing: a single thread is returned unchanged', () => {
  const only = fakeThread('only', new Date('2026-09-04T08:00:00Z'));

  assert.deepEqual(idsOf(orderThreadsForProcessing([only])), ['only']);
});

test('orderThreadsForProcessing: a thread whose date is missing/unparseable sorts FIRST and never throws -- an unknown date must never silently push a thread behind a cancellation', () => {
  const noDate = { getId: function () { return 'no-date'; }, getLastMessageDate: function () { return null; } };
  const dated = fakeThread('dated', new Date('2026-09-04T08:00:00Z'));

  let result;
  assert.doesNotThrow(function () {
    result = orderThreadsForProcessing([dated, noDate]);
  });
  assert.deepEqual(idsOf(result), ['no-date', 'dated']);
});

test('orderThreadsForProcessing: a thread whose getLastMessageDate returns an INVALID Date sorts FIRST -- new Date("nonsense") is a real Date instance, so the instanceof check alone does not catch it', () => {
  const invalidDate = {
    getId: function () {
      return 'invalid-date';
    },
    getLastMessageDate: function () {
      return new Date('not a date');
    },
  };
  const dated = fakeThread('dated', new Date('2026-09-04T08:00:00Z'));

  assert.deepEqual(idsOf(orderThreadsForProcessing([dated, invalidDate])), ['invalid-date', 'dated']);
});

test('orderThreadsForProcessing: a thread whose getLastMessageDate returns a non-Date (string) sorts FIRST, never throws', () => {
  const stringDate = {
    getId: function () {
      return 'string-date';
    },
    getLastMessageDate: function () {
      return '2026-09-09T08:00:00Z';
    },
  };
  const dated = fakeThread('dated', new Date('2026-09-04T08:00:00Z'));

  assert.deepEqual(idsOf(orderThreadsForProcessing([dated, stringDate])), ['string-date', 'dated']);
});

test('orderThreadsForProcessing: a thread whose getLastMessageDate THROWS is tolerated (sorts first, the throw never escapes the sort)', () => {
  const thrower = {
    getId: function () {
      return 'thrower';
    },
    getLastMessageDate: function () {
      throw new Error('Gmail service unavailable');
    },
  };
  const dated = fakeThread('dated', new Date('2026-09-04T08:00:00Z'));

  let result;
  assert.doesNotThrow(function () {
    result = orderThreadsForProcessing([dated, thrower]);
  });
  assert.deepEqual(idsOf(result), ['thrower', 'dated']);
});

test('orderThreadsForProcessing: a thread with no getLastMessageDate method at all is tolerated (sorts first, no throw)', () => {
  const noAccessor = { getId: function () { return 'no-accessor'; } };
  const dated = fakeThread('dated', new Date('2026-09-04T08:00:00Z'));

  let result;
  assert.doesNotThrow(function () {
    result = orderThreadsForProcessing([dated, noAccessor]);
  });
  assert.deepEqual(idsOf(result), ['no-accessor', 'dated']);
});

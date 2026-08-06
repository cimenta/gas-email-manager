'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  hasGuestRelationship,
  findExistingEventByICalUid,
  importIcsEventWithSequenceRetry,
  parseIcs,
  buildEventResource,
} = require('../src/05-action-ics-import.js');

// ---------------------------------------------------------------------------
// PRESERVE-EXISTING-INVITE GUARD — regression coverage for the live-reported
// bug `ics-import-strips-rsvp`.
//
// THE BUG: Gmail's own native invite detection creates a genuine ATTENDEE COPY
// for an incoming Exchange/Teams invite (real organizer, owner as NEEDS-ACTION
// attendee, Accept/Decline UI). Minutes later the periodic trigger ran
// ICS_CALENDAR_ACTION over the same .ics and called Calendar.Events.import —
// a full-resource replace keyed by iCalUID, carrying NEITHER organizer NOR
// attendees (buildEventResource's deliberate T-03-05 firewall). That demoted
// the attendee copy to a plain self-owned private copy: guest list cleared,
// organizer reset to the owner, RSVP destroyed, organizer never notified.
//
// ORACLE TYPE: specified. The required behavior is stated directly — do not
// write over an event that already carries guests — so these assert the exact
// specified outcome (no import call at all), not merely "does not crash".
//
// The two constraints that must survive the fix are asserted here too:
//   * T-03-05 — the resource must still never carry attendees/organizer.
//   * quick-260723-gmk / quick-260731-seq — dedup-by-iCalUID and the
//     single-shot sequence-conflict recovery must be untouched whenever the
//     guard does not fire.
// ---------------------------------------------------------------------------

// --- hasGuestRelationship (pure) --------------------------------------------
//
// Boundary neighbours around the guard's equivalence class are covered
// explicitly: no event at all, event with no attendees key, 0 attendees
// (the off-by-one neighbour of the trigger point), 1 attendee (the trigger
// point itself), and many.

test('hasGuestRelationship: null (no existing event) -> false', () => {
  assert.equal(hasGuestRelationship(null), false);
});

test('hasGuestRelationship: undefined -> false', () => {
  assert.equal(hasGuestRelationship(undefined), false);
});

test('hasGuestRelationship: event with no attendees key at all -> false', () => {
  assert.equal(hasGuestRelationship({ id: 'evt-1', summary: 'Solo event' }), false);
});

test('hasGuestRelationship: BOUNDARY attendees: [] (exactly zero guests) -> false', () => {
  assert.equal(hasGuestRelationship({ id: 'evt-1', attendees: [] }), false);
});

test('hasGuestRelationship: BOUNDARY attendees length 1 (exactly one guest) -> true', () => {
  const event = {
    id: 'evt-1',
    attendees: [{ email: 'owner@example.com', responseStatus: 'needsAction', self: true }],
  };
  assert.equal(hasGuestRelationship(event), true);
});

test('hasGuestRelationship: many guests -> true', () => {
  const event = {
    id: 'evt-1',
    attendees: [
      { email: 'owner@example.com', responseStatus: 'needsAction', self: true },
      { email: 'jana@example.com', organizer: true },
      { email: 'carol@example.com', responseStatus: 'accepted' },
    ],
  };
  assert.equal(hasGuestRelationship(event), true);
});

test('hasGuestRelationship: an event the OWNER organized (guests are invitees, not self) is guarded too', () => {
  // Blowing away your own meeting's invitee list is just as destructive as
  // losing an RSVP you owed someone else — same array, same guard.
  const event = {
    id: 'evt-1',
    organizer: { email: 'owner@example.com', self: true },
    attendees: [{ email: 'someone.else@example.com', responseStatus: 'accepted' }],
  };
  assert.equal(hasGuestRelationship(event), true);
});

// --- findExistingEventByICalUid (faked Calendar global) ----------------------

test('findExistingEventByICalUid: queries the given calendarId by iCalUID with singleEvents:false, returns the first item', () => {
  const calls = [];
  const master = { id: 'evt-master', attendees: [{ email: 'a@example.com' }] };
  global.Calendar = {
    Events: {
      list: function (calendarId, params) {
        calls.push({ calendarId: calendarId, params: params });
        return { items: [master] };
      },
    },
  };

  try {
    const found = findExistingEventByICalUid('CUSTOM_CAL', 'uid-123');
    assert.equal(found, master);
    assert.equal(calls.length, 1);
    assert.equal(calls[0].calendarId, 'CUSTOM_CAL');
    assert.deepEqual(calls[0].params, { iCalUID: 'uid-123', singleEvents: false });
  } finally {
    delete global.Calendar;
  }
});

test('findExistingEventByICalUid: empty items / missing items / empty response all degrade to null (never throws)', () => {
  const responses = [{ items: [] }, {}, null, undefined];

  responses.forEach(function (response) {
    global.Calendar = { Events: { list: function () { return response; } } };
    try {
      assert.equal(findExistingEventByICalUid('CAL', 'uid-123'), null);
    } finally {
      delete global.Calendar;
    }
  });
});

// --- importIcsEventWithSequenceRetry: the guard ------------------------------
//
// Harness: fakes the Calendar global and captures console.log so the guard's
// diagnostic line does not pollute test output while still being assertable.

function withFakeCalendar(events, run) {
  const state = {
    importCalls: [],
    listCalls: [],
    logs: [],
  };
  const originalLog = console.log;

  global.Calendar = {
    Events: {
      list: function (calendarId, params) {
        state.listCalls.push({ calendarId: calendarId, params: params });
        return { items: events ? [events] : [] };
      },
      import: function (resource, calendarId, args) {
        state.importCalls.push({ resource: resource, calendarId: calendarId, args: args });
      },
    },
  };
  console.log = function (line) {
    state.logs.push(line);
  };

  try {
    run(state);
  } finally {
    console.log = originalLog;
    delete global.Calendar;
  }

  return state;
}

test('REGRESSION (ics-import-strips-rsvp): an existing Gmail-native invite with guests is NEVER written over — zero import calls', () => {
  // Exactly the reported situation: Gmail already created the attendee copy.
  const gmailNativeEvent = {
    id: 'evt-gmail-native',
    iCalUID: 'teams-uid-abc',
    sequence: 0,
    organizer: { email: 'jana@example.com', displayName: 'Jana Nováková' },
    attendees: [
      { email: 'owner@example.com', responseStatus: 'needsAction', self: true },
      { email: 'jana@example.com', organizer: true, responseStatus: 'accepted' },
    ],
  };

  let result;
  const state = withFakeCalendar(gmailNativeEvent, function () {
    const resource = { summary: 'Teams sync', iCalUID: 'teams-uid-abc', sequence: 0 };
    result = importIcsEventWithSequenceRetry(resource, 'primary', 'teams-uid-abc');
  });

  // THE load-bearing assertion, deliberately FIRST so it is what fails if the
  // guard is ever removed: no write of any kind may reach Calendar. Asserting
  // the return value first would let a fix that merely reports the right
  // outcome, while still writing, slip through.
  assert.equal(state.importCalls.length, 0, 'Calendar.Events.import must not be called over an event that carries guests');
  assert.equal(state.listCalls.length, 1);
  assert.deepEqual(result, { action: 'skipped-existing-invite', eventId: 'evt-gmail-native' });
  assert.equal(state.logs.length, 1);
  assert.match(state.logs[0], /teams-uid-abc/);
  assert.match(state.logs[0], /skipping import/);
});

test('no existing event for the iCalUID -> imports exactly once, unchanged create behavior', () => {
  const resource = { summary: 'Brand new', iCalUID: 'fresh-uid', sequence: 0 };

  const state = withFakeCalendar(null, function () {
    const result = importIcsEventWithSequenceRetry(resource, 'primary', 'fresh-uid');
    assert.deepEqual(result, { action: 'imported', eventId: null });
  });

  assert.equal(state.importCalls.length, 1);
  assert.equal(state.importCalls[0].resource, resource);
  assert.equal(state.importCalls[0].calendarId, 'primary');
  assert.deepEqual(state.importCalls[0].args, {});
  assert.equal(state.logs.length, 0, 'the guard must stay silent when it does not fire');
});

test('existing event WITHOUT guests (this script own prior import) -> still imported, dedup-by-iCalUID preserved', () => {
  // quick-260723-gmk must not regress: re-processing a thread whose event this
  // script itself created still updates that same event rather than skipping.
  const ownPriorEvent = { id: 'evt-ours', iCalUID: 'publish-uid', sequence: 0, attendees: [] };
  const resource = { summary: 'Updated title', iCalUID: 'publish-uid', sequence: 1 };

  const state = withFakeCalendar(ownPriorEvent, function () {
    const result = importIcsEventWithSequenceRetry(resource, 'primary', 'publish-uid');
    assert.deepEqual(result, { action: 'imported', eventId: null });
  });

  assert.equal(state.importCalls.length, 1);
  assert.equal(state.importCalls[0].resource.summary, 'Updated title');
});

test('optionalArgs (e.g. transport tickets supportsAttachments) still passes through to import unchanged', () => {
  const resource = { summary: 'RegioJet Praha - Brno', iCalUID: 'rj-uid', sequence: 0 };

  const state = withFakeCalendar(null, function () {
    importIcsEventWithSequenceRetry(resource, 'TRANSPORT_CAL', 'rj-uid', { supportsAttachments: true });
  });

  assert.equal(state.importCalls.length, 1);
  assert.deepEqual(state.importCalls[0].args, { supportsAttachments: true });
});

// --- importIcsEventWithSequenceRetry: pre-existing behavior untouched --------

test('sequence-conflict recovery (quick-260731-seq) still works when the guard does not fire', () => {
  const resource = { summary: 'Meeting', iCalUID: 'seq-uid', sequence: 0 };
  const importCalls = [];
  let listCount = 0;

  global.Calendar = {
    Events: {
      list: function () {
        listCount += 1;
        // 1st call: the guard's lookup -> an existing event with NO guests.
        // 2nd call: the recovery lookup -> that event's real sequence.
        return { items: [{ id: 'evt-x', sequence: 7, attendees: [] }] };
      },
      import: function (r) {
        importCalls.push({ sequence: r.sequence });
        if (importCalls.length === 1) {
          throw new Error('GoogleJsonResponseException: Invalid sequence value. Re-fetch the resource...');
        }
      },
    },
  };

  try {
    const result = importIcsEventWithSequenceRetry(resource, 'primary', 'seq-uid');
    assert.deepEqual(result, { action: 'imported', eventId: 'evt-x' });
  } finally {
    delete global.Calendar;
  }

  assert.equal(listCount, 2, 'one lookup for the guard, one for the sequence recovery');
  assert.deepEqual(importCalls, [{ sequence: 0 }, { sequence: 7 }], 'retry re-sends with the existing event sequence');
  assert.equal(resource.sequence, 7);
});

test('a non-sequence error is still rethrown unchanged, never retried', () => {
  const importCalls = [];
  global.Calendar = {
    Events: {
      list: function () { return { items: [] }; },
      import: function () {
        importCalls.push(1);
        throw new Error('Calendar usage limits exceeded.');
      },
    },
  };

  try {
    assert.throws(
      function () {
        importIcsEventWithSequenceRetry({ iCalUID: 'x' }, 'primary', 'x');
      },
      /Calendar usage limits exceeded/
    );
  } finally {
    delete global.Calendar;
  }

  assert.equal(importCalls.length, 1, 'no retry for a non-sequence error');
});

// --- T-03-05 firewall: still intact after the fix ---------------------------

test('T-03-05 still holds end to end: the resource actually sent to import carries no attendees and no organizer', () => {
  // A real METHOD:REQUEST invite with untrusted ORGANIZER/ATTENDEE lines. Those
  // must remain description text only — the fix must not have quietly started
  // forwarding them as real Calendar guests.
  const invite = [
    'BEGIN:VCALENDAR',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:teams-uid-abc',
    'SUMMARY:Teams sync',
    'DTSTART:20260810T090000Z',
    'DTEND:20260810T093000Z',
    'ORGANIZER;CN=Jana:mailto:jana@example.com',
    'ATTENDEE;CN=Owner;PARTSTAT=NEEDS-ACTION:mailto:owner@example.com',
    'ATTENDEE;CN=Carol:mailto:carol@example.com',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(invite);
  assert.equal(events.length, 1);
  const resource = buildEventResource(events[0]);

  const state = withFakeCalendar(null, function () {
    importIcsEventWithSequenceRetry(resource, 'primary', events[0].uid);
  });

  assert.equal(state.importCalls.length, 1);
  const sent = state.importCalls[0].resource;
  assert.equal(Object.prototype.hasOwnProperty.call(sent, 'attendees'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(sent, 'organizer'), false);
  // ...and the informational text is still there, so enrichment is unaffected.
  assert.match(sent.description, /Organizer: Jana <jana@example\.com>/);
  assert.match(sent.description, /carol@example\.com/);
});

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const {
  ICS_CALENDAR_ACTION,
  planIcsEventWrite,
  parseIcs,
} = require('../src/05-action-ics-import.js');

// --- ICS_CALENDAR_ACTION: STATUS:CANCELLED VEVENT hardening ----------------
//
// Live-reported bug regiojet-cancel-not-deleted, the deferred blind_spot,
// now owner-requested in scope.
//
// THE DEFECT. ICS_CALENDAR_ACTION is registry index 0 -- it runs BEFORE
// TRANSPORT_TICKETS_ACTION on the very same thread. Its `run` had NO branch
// for `event.status` at all: every parsed VEVENT went to
// Calendar.Events.import (uid present) or Calendar.Events.insert (uid
// absent), unconditionally. A RegioJet CANCELLATION .ics -- METHOD:CANCEL,
// one VEVENT carrying STATUS:CANCELLED -- was therefore imported as an
// ordinary LIVE event. The parser has carried `event.status` (trimmed,
// uppercased, single normalization point) since D-01/D-02 of
// quick-260813-dq2, but only TRANSPORT_TICKETS_ACTION ever read it.
//
// WHY THE EXISTING MITIGATION IS NOT ONE. The only thing that made this
// action stand down for RegioJet was the owner having set the
// `05-action-ics-EXCLUDE_FROM` Script Property. Its code default is `[]`
// (see ICS_ACTION_CONFIG.excludeFrom) -- so on a fresh install, or any
// install where that out-of-band owner-side step was missed, a cancellation
// email CREATED a live event, and the confirmation before it had created a
// second, UNTAGGED copy on the ICS-resolved calendar that
// TRANSPORT_TICKETS_ACTION's ticketIdentifier-keyed cancellation path can
// never reach. A correctness guarantee must not rest on a Script Property
// that defaults to off; the tests below therefore run with EXCLUDE_FROM at
// its DEFAULT (unset) throughout, and one explicitly asserts that default.
//
// THE FIX, in two parts, both at this action's own write site:
//   (1) a cancelled VEVENT is NEVER written as a live event -- no import,
//       no insert, in any configuration;
//   (2) a cancelled VEVENT that carries a UID actively CANCELS the event
//       already stored under that iCalUID on the RESOLVED calendar
//       (findExistingEventByICalUid -> Calendar.Events.remove), which is
//       what clears the orphaned copy described above. A cancelled VEVENT
//       with NO uid references nothing and is skipped outright.
//
// NOT DONE, deliberately: `buildEventResource` is left untouched. The
// D-01 firewall (it copies neither `status` nor `dtstamp` onto the Calendar
// API resource) is load-bearing for TRANSPORT_TICKETS_ACTION, which shares
// that pure builder and has its own, separate cancellation path; see
// test/transport-tickets.test.js's "entry.resource carries no status field"
// test, which stays green and unmodified. Cancellation is a WRITE-SITE
// decision here, not a resource-shape change to a shared builder.

// --- fixtures --------------------------------------------------------------

function buildIcs(veventLineGroups) {
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0'];

  veventLineGroups.forEach(function (group) {
    lines.push('BEGIN:VEVENT');
    group.forEach(function (line) {
      if (line !== null) {
        lines.push(line);
      }
    });
    lines.push('END:VEVENT');
  });

  lines.push('END:VCALENDAR');

  return lines.join('\r\n');
}

// A cancelled VEVENT shaped like RegioJet's real one (anonymized fictional
// ticket number and UID, same repo convention as the transport fixtures):
// METHOD:CANCEL-style payload, STATUS:CANCELLED, SEQUENCE already advanced.
function cancelledVevent(uidLine) {
  return [
    'DTSTAMP:20260908T131030Z',
    'DTSTART:20260908T160100Z',
    'DTEND:20260908T184200Z',
    'SUMMARY:#7788123456: Z Praha, hl.n., do Brno, hl.n., sedadla: [2/32]',
    'LOCATION:50.0830000, 14.4350000',
    uidLine,
    'STATUS:CANCELLED',
    'SEQUENCE:2',
  ];
}

function liveVevent(uidLine, statusLine) {
  return [
    'DTSTAMP:20260908T131240Z',
    'DTSTART:20260908T130100Z',
    'DTEND:20260908T150000Z',
    'SUMMARY:Elektronická jízdenka',
    'LOCATION:50.0830000, 14.4350000',
    uidLine,
    statusLine || null,
    'SEQUENCE:1',
  ];
}

const CANCEL_ICS = buildIcs([cancelledVevent('UID:-9876543210@regiojet.cz')]);
const CANCEL_ICS_NO_UID = buildIcs([cancelledVevent(null)]);

// --- GAS harness -----------------------------------------------------------

function fakeAttachment(name, data) {
  return {
    getName: function () {
      return name;
    },
    getContentType: function () {
      return 'text/calendar';
    },
    getDataAsString: function () {
      return data;
    },
  };
}

function fakeThread(icsTexts, fromHeader) {
  return {
    getId: function () {
      return 'thread-regiojet-cancel-1';
    },
    getMessages: function () {
      return [
        {
          getFrom: function () {
            return fromHeader || 'RegioJet <jizdenky@regiojet.cz>';
          },
          getAttachments: function () {
            return icsTexts.map(function (text, i) {
              return fakeAttachment('ticket' + i + '.ics', text);
            });
          },
        },
      ];
    },
  };
}

// Captures every Calendar write run() performs. `existingByUid` seeds the
// calendar with events already stored under a given iCalUID, so both
// findExistingEventByICalUid and the sequence-conflict recovery's own
// lookup resolve against it. Mirrors the harness convention in
// test/ics-mislabeled-encoding.test.js and test/existing-invite-guard.test.js.
function withFakeGasGlobals(options, body) {
  const opts = options || {};
  const existingByUid = opts.existingByUid || {};
  const state = { imported: [], inserted: [], removed: [], listCalls: [] };

  global.CONFIG = { calendarId: 'primary' };
  global.CalendarApp = {
    getCalendarById: function (id) {
      return { id: id };
    },
  };
  global.Calendar = {
    Events: {
      list: function (calendarId, params) {
        state.listCalls.push({ calendarId: calendarId, params: params });
        const uid = params && params.iCalUID;
        const hit = Object.prototype.hasOwnProperty.call(existingByUid, uid) ? existingByUid[uid] : null;
        return { items: hit ? [hit] : [] };
      },
      import: function (resource, calendarId) {
        state.imported.push({ resource: resource, calendarId: calendarId });
        return { id: 'evt-imported' };
      },
      insert: function (resource, calendarId) {
        state.inserted.push({ resource: resource, calendarId: calendarId });
        return { id: 'evt-inserted' };
      },
      remove: function (calendarId, eventId) {
        state.removed.push({ calendarId: calendarId, eventId: eventId });
      },
    },
  };

  try {
    body(state);
  } finally {
    delete global.CONFIG;
    delete global.CalendarApp;
    delete global.Calendar;
  }

  return state;
}

function withFakePropertiesService(properties, fn) {
  const store = Object.assign({}, properties);

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
    fn();
  } finally {
    delete global.PropertiesService;
  }
}

// --- THE DEFECT ------------------------------------------------------------

test('THE DEFECT: run() on a STATUS:CANCELLED VEVENT creates NO live calendar event (no import, no insert)', () => {
  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
  });

  assert.equal(state.imported.length, 0, 'a cancelled VEVENT must never be imported as a live event');
  assert.equal(state.inserted.length, 0, 'a cancelled VEVENT must never be inserted as a live event');
});

test('THE DEFECT, uid-less variant: a STATUS:CANCELLED VEVENT with NO UID is skipped outright -- it references nothing, so there is nothing to create and nothing to cancel', () => {
  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS_NO_UID]));
  });

  assert.equal(state.inserted.length, 0, 'the uid-less insert() fallback must not fire for a cancelled VEVENT');
  assert.equal(state.imported.length, 0);
  assert.equal(state.removed.length, 0, 'with no uid there is no event to reference -- nothing may be deleted');
});

test('THE ORPHAN CLEANUP: a cancelled VEVENT whose UID already has a live event on the calendar REMOVES that event', () => {
  const state = withFakeGasGlobals(
    { existingByUid: { '-9876543210@regiojet.cz': { id: 'evt-existing-live', attendees: [] } } },
    function () {
      ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
    }
  );

  assert.equal(state.removed.length, 1, 'the event stored under that iCalUID must be deleted');
  assert.equal(state.removed[0].eventId, 'evt-existing-live');
  assert.equal(state.imported.length, 0, 'and nothing may be written back in its place');
  assert.equal(state.inserted.length, 0);
});

test('a cancelled VEVENT whose UID has NO event on the calendar is a clean no-op: nothing created, nothing deleted, no throw', () => {
  const state = withFakeGasGlobals({}, function () {
    assert.doesNotThrow(function () {
      ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
    });
  });

  assert.equal(state.removed.length, 0, 'must not attempt a delete when the lookup found nothing');
  assert.equal(state.imported.length, 0);
  assert.equal(state.inserted.length, 0);
});

// --- BOUNDARY: cancelled and live VEVENTs in ONE .ics ----------------------

test('BOUNDARY: one .ics carrying a live AND a cancelled VEVENT imports ONLY the live one', () => {
  const mixed = buildIcs([
    liveVevent('UID:-1111111111@regiojet.cz', null),
    cancelledVevent('UID:-9876543210@regiojet.cz'),
  ]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([mixed]));
  });

  assert.equal(state.imported.length, 1, 'exactly the live VEVENT, never the cancelled one');
  assert.equal(state.imported[0].resource.iCalUID, '-1111111111@regiojet.cz');
  assert.equal(state.inserted.length, 0);
});

test('BOUNDARY: in a mixed .ics the cancelled VEVENT still cancels its own UID while the live one still imports -- the two decisions are independent', () => {
  const mixed = buildIcs([
    liveVevent('UID:-1111111111@regiojet.cz', null),
    cancelledVevent('UID:-9876543210@regiojet.cz'),
  ]);

  const state = withFakeGasGlobals(
    { existingByUid: { '-9876543210@regiojet.cz': { id: 'evt-existing-live', attendees: [] } } },
    function () {
      ICS_CALENDAR_ACTION.run(fakeThread([mixed]));
    }
  );

  assert.equal(state.imported.length, 1);
  assert.equal(state.imported[0].resource.iCalUID, '-1111111111@regiojet.cz');
  assert.equal(state.removed.length, 1);
  assert.equal(state.removed[0].eventId, 'evt-existing-live');
});

test('BOUNDARY: TWO cancelled VEVENTs in one .ics each cancel their own UID -- the guard is per VEVENT, not per attachment', () => {
  const twoCancels = buildIcs([
    cancelledVevent('UID:-9876543210@regiojet.cz'),
    cancelledVevent('UID:-5555555555@regiojet.cz'),
  ]);

  const state = withFakeGasGlobals(
    {
      existingByUid: {
        '-9876543210@regiojet.cz': { id: 'evt-a', attendees: [] },
        '-5555555555@regiojet.cz': { id: 'evt-b', attendees: [] },
      },
    },
    function () {
      ICS_CALENDAR_ACTION.run(fakeThread([twoCancels]));
    }
  );

  assert.deepEqual(
    state.removed.map(function (r) {
      return r.eventId;
    }),
    ['evt-a', 'evt-b']
  );
  assert.equal(state.imported.length, 0);
});

// --- BOUNDARY: the STATUS token itself ------------------------------------

test('BOUNDARY: mixed-case STATUS:Cancelled fires the guard too -- the guard compares the PARSER-normalized token, it does not re-normalize', () => {
  const mixedCase = buildIcs([
    [
      'DTSTAMP:20260908T131030Z',
      'DTSTART:20260908T160100Z',
      'DTEND:20260908T184200Z',
      'SUMMARY:Mixed-case status',
      'UID:-9876543210@regiojet.cz',
      'STATUS:Cancelled',
    ],
  ]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([mixedCase]));
  });

  assert.equal(state.imported.length, 0);
  assert.equal(state.inserted.length, 0);
});

// --- REGRESSION GUARDS: what this fix could itself break -------------------

test('REGRESSION GUARD: an ORDINARY VEVENT with no STATUS line at all still imports exactly as before -- same resource, same calendar', () => {
  const ordinary = buildIcs([liveVevent('UID:-1111111111@regiojet.cz', null)]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([ordinary]));
  });

  assert.equal(state.imported.length, 1);
  assert.equal(state.inserted.length, 0);
  assert.equal(state.removed.length, 0, 'the new cancel path must never fire for a non-cancelled event');
  assert.equal(state.imported[0].calendarId, 'primary');
  assert.deepEqual(state.imported[0].resource, {
    summary: 'Elektronická jízdenka',
    description: '',
    location: '50.0830000, 14.4350000',
    sequence: 1,
    iCalUID: '-1111111111@regiojet.cz',
    start: { dateTime: '2026-09-08T13:01:00.000Z' },
    end: { dateTime: '2026-09-08T15:00:00.000Z' },
  });
});

test('REGRESSION GUARD: STATUS:CONFIRMED still imports -- the guard keys on CANCELLED only, it is not a "has a status" test', () => {
  const confirmed = buildIcs([liveVevent('UID:-1111111111@regiojet.cz', 'STATUS:CONFIRMED')]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([confirmed]));
  });

  assert.equal(state.imported.length, 1);
  assert.equal(state.removed.length, 0);
});

test('REGRESSION GUARD: STATUS:TENTATIVE still imports', () => {
  const tentative = buildIcs([liveVevent('UID:-1111111111@regiojet.cz', 'STATUS:TENTATIVE')]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([tentative]));
  });

  assert.equal(state.imported.length, 1);
  assert.equal(state.removed.length, 0);
});

test('REGRESSION GUARD: an ordinary VEVENT with NO uid still takes the insert() fallback -- only the CANCELLED uid-less case is skipped', () => {
  const ordinaryNoUid = buildIcs([liveVevent(null, null)]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([ordinaryNoUid]));
  });

  assert.equal(state.inserted.length, 1);
  assert.equal(state.imported.length, 0);
  assert.equal(state.removed.length, 0);
});

test('REGRESSION GUARD: a valid VCALENDAR carrying ZERO VEVENTs remains a silent no-op (unchanged by the cancel path)', () => {
  const empty = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'END:VCALENDAR'].join('\r\n');

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([empty]));
  });

  assert.equal(state.imported.length, 0);
  assert.equal(state.inserted.length, 0);
  assert.equal(state.removed.length, 0);
});

test('REGRESSION GUARD: a cancelled VEVENT whose UID resolves to a GUEST-BEARING event leaves it untouched -- Gmail owns a real attendee copy (same stance as the preserve-existing-invite guard)', () => {
  const state = withFakeGasGlobals(
    {
      existingByUid: {
        '-9876543210@regiojet.cz': {
          id: 'evt-real-invite',
          attendees: [{ email: 'radek.simcik@example.com', responseStatus: 'accepted' }],
        },
      },
    },
    function () {
      ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
    }
  );

  assert.equal(state.removed.length, 0, 'never delete an event carrying a real guest relationship');
  assert.equal(state.imported.length, 0, 'and still never write it back as a live event');
  assert.equal(state.inserted.length, 0);
});

// --- BOUNDARY: EXCLUDE_FROM set vs unset ----------------------------------
//
// The whole point of this fix: the guarantee must hold at the SHIPPED
// DEFAULT, with the owner having configured nothing.

test('BOUNDARY, EXCLUDE_FROM UNSET (the shipped default): config.excludeFrom is [] AND the cancelled VEVENT is still not imported -- the guarantee does not depend on owner configuration', () => {
  assert.deepEqual(ICS_CALENDAR_ACTION.config.excludeFrom, [], 'the shipped code default excludes nobody');

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
  });

  assert.equal(state.imported.length, 0);
  assert.equal(state.inserted.length, 0);
});

test('BOUNDARY, EXCLUDE_FROM EMPTY STRING (property present but blank) behaves identically to unset', () => {
  withFakePropertiesService({ '05-action-ics-EXCLUDE_FROM': '' }, function () {
    assert.deepEqual(ICS_CALENDAR_ACTION.config.excludeFrom, []);

    const state = withFakeGasGlobals({}, function () {
      ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
    });

    assert.equal(state.imported.length, 0);
    assert.equal(state.inserted.length, 0);
  });
});

test('BOUNDARY, EXCLUDE_FROM SET to the sender: appliesTo() is false, so dispatchActions never runs this action at all -- the new cancel path can never reach a calendar another action owns', () => {
  const thread = fakeThread([CANCEL_ICS]);

  // Contrast, at the shipped default: the action DOES claim the thread --
  // which is exactly why the guard inside run() has to exist.
  assert.equal(ICS_CALENDAR_ACTION.appliesTo(thread), true, 'with EXCLUDE_FROM unset this action claims the RegioJet thread');

  withFakePropertiesService({ '05-action-ics-EXCLUDE_FROM': 'jizdenky@regiojet.cz' }, function () {
    assert.deepEqual(ICS_CALENDAR_ACTION.config.excludeFrom, ['jizdenky@regiojet.cz']);
    assert.equal(
      ICS_CALENDAR_ACTION.appliesTo(thread),
      false,
      'an excluded sender is not this action\'s business at all -- run() is never entered, so no import, insert or delete can happen'
    );
  });
});

// --- BOUNDARY: multi-calendar routing -------------------------------------

test('BOUNDARY: the cancellation targets the RESOLVED calendar, not CONFIG.calendarId -- both the lookup and the delete', () => {
  withFakePropertiesService({ '05-action-ics-CALENDAR_ID': 'routed@group.calendar.google.com' }, function () {
    const state = withFakeGasGlobals(
      { existingByUid: { '-9876543210@regiojet.cz': { id: 'evt-existing-live', attendees: [] } } },
      function () {
        ICS_CALENDAR_ACTION.run(fakeThread([CANCEL_ICS]));
      }
    );

    assert.equal(state.removed.length, 1);
    assert.equal(state.removed[0].calendarId, 'routed@group.calendar.google.com');
    assert.equal(state.listCalls[0].calendarId, 'routed@group.calendar.google.com');
  });
});

// --- planIcsEventWrite: the pure decision function -------------------------
//
// The write decision extracted as a pure, directly-testable function, the
// same house pattern as hasGuestRelationship (pure, unit-tested) sitting
// under importIcsEventWithSequenceRetry (GAS-only, tested through faked
// globals). Returning `resource: null` for both cancelled branches is a
// STRUCTURAL guarantee, not a convention: no resource is ever built from a
// cancelled VEVENT, so none can reach import/insert by any later edit.

test('planIcsEventWrite: a live event WITH a uid -> import, carrying the built resource', () => {
  const event = parseIcs(buildIcs([liveVevent('UID:-1111111111@regiojet.cz', null)]))[0];
  const plan = planIcsEventWrite(event);

  assert.equal(plan.action, 'import');
  assert.equal(plan.uid, '-1111111111@regiojet.cz');
  assert.equal(plan.resource.iCalUID, '-1111111111@regiojet.cz');
});

test('planIcsEventWrite: a live event WITHOUT a uid -> insert, resource present, uid null', () => {
  const event = parseIcs(buildIcs([liveVevent(null, null)]))[0];
  const plan = planIcsEventWrite(event);

  assert.equal(plan.action, 'insert');
  assert.equal(plan.uid, null);
  assert.ok(plan.resource, 'an ordinary uid-less event still needs a resource to insert');
  assert.ok(!('iCalUID' in plan.resource));
});

test('planIcsEventWrite: a CANCELLED event WITH a uid -> cancel, and resource is null (no resource is ever built from a cancelled VEVENT)', () => {
  const event = parseIcs(CANCEL_ICS)[0];
  const plan = planIcsEventWrite(event);

  assert.equal(plan.action, 'cancel');
  assert.equal(plan.uid, '-9876543210@regiojet.cz');
  assert.equal(plan.resource, null);
});

test('planIcsEventWrite: a CANCELLED event WITHOUT a uid -> skip, resource null, uid null', () => {
  const event = parseIcs(CANCEL_ICS_NO_UID)[0];
  const plan = planIcsEventWrite(event);

  assert.equal(plan.action, 'skip');
  assert.equal(plan.uid, null);
  assert.equal(plan.resource, null);
});

test('planIcsEventWrite: STATUS:CONFIRMED and STATUS:TENTATIVE both route to import, not cancel', () => {
  ['STATUS:CONFIRMED', 'STATUS:TENTATIVE'].forEach(function (statusLine) {
    const event = parseIcs(buildIcs([liveVevent('UID:-1111111111@regiojet.cz', statusLine)]))[0];
    assert.equal(planIcsEventWrite(event).action, 'import', statusLine + ' must not be treated as a cancellation');
  });
});

test('planIcsEventWrite: an event object with NO status key at all (hand-built, not from parseVeventBlock) routes to import -- never crashes on the missing field', () => {
  const plan = planIcsEventWrite({
    summary: 'S',
    description: 'D',
    location: 'L',
    start: new Date('2026-08-01T10:00:00Z'),
    end: new Date('2026-08-01T11:00:00Z'),
    isAllDay: false,
    recurrence: null,
    uid: 'u@x',
    sequence: 0,
  });

  assert.equal(plan.action, 'import');
});

test('planIcsEventWrite: an empty-string uid on a cancelled event is treated as NO uid -> skip (falsy uid, same rule as the live insert() fallback)', () => {
  const plan = planIcsEventWrite({
    summary: 'S',
    description: '',
    location: '',
    start: new Date('2026-09-08T16:01:00Z'),
    end: new Date('2026-09-08T18:42:00Z'),
    isAllDay: false,
    recurrence: null,
    uid: '',
    sequence: 2,
    status: 'CANCELLED',
  });

  assert.equal(plan.action, 'skip');
});

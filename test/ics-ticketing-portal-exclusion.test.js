'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

// LOAD ORDER (D-07, load-bearing for this file): 05 is required FIRST, then
// 07 -- this file is this repo's coverage for the 05-first load order. The
// companion 07-first order is covered by the child-process tests at the
// bottom of this file (a same-process require() only ever resolves a module
// once, so this static pair alone cannot prove the other order).
const {
  ICS_CALENDAR_ACTION,
  isTicketingPortalSender,
} = require('../src/05-action-ics-import.js');
const { TICKETING_PORTALS_ACTION } = require('../src/07-action-ticketing-portals.js');

// --- quick-261005-orv: structural ICS-import exclusion of every configured --
// --- ticketing-portal sender (D-07) ------------------------------------------
//
// THE DEFECT THIS CLOSES: Ticketportal.cz's order confirmation carries its
// event data in its OWN .ics attachment. Without this exclusion,
// ICS_CALENDAR_ACTION (this file) and TICKETING_PORTALS_ACTION (the sibling
// action) would BOTH see that same .ics and BOTH create a calendar event for
// it -- Events.import (iCalUID-deduped) and Events.insert
// (ticketIdentifier-deduped) each dedup against their OWN mechanism, so
// neither one's dedup can see the other's write. isTicketingPortalSender
// closes this STRUCTURALLY: it is applied UNCONDITIONALLY, independent of
// config.excludeFrom, which is why every test below runs with excludeFrom at
// its code DEFAULT ([]) -- the regiojet-cancel-not-deleted precedent this
// plan cites is exactly this: "a correctness guarantee resting on a Script
// Property that defaults to off is not a mitigation; it is an unset switch."
//
// The .ics text below is a byte-for-byte copy of the one in
// test/ticketing-portals.test.js (same real sample,
// "Potvrzení objednávky vstupenek.eml") -- duplicated here deliberately
// rather than shared via a require, so this file's own fixture integrity
// does not depend on the sibling test file's internals.

const REAL_TICKETPORTAL_CZ_ICS_TEXT = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'PRODID:-//TICKETPORTAL//Vstupenky na dosah//CZ',
  'METHOD:PUBLISH',
  'BEGIN:VTIMEZONE',
  'TZID:Europe/Prague',
  'LAST-MODIFIED:20231222T233358Z',
  'TZURL:https://www.tzurl.org/zoneinfo-outlook/Europe/Prague',
  'X-LIC-LOCATION:Europe/Prague',
  'BEGIN:DAYLIGHT',
  'TZNAME:CEST',
  'TZOFFSETFROM:+0100',
  'TZOFFSETTO:+0200',
  'DTSTART:19700329T020000',
  'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
  'END:DAYLIGHT',
  'BEGIN:STANDARD',
  'TZNAME:CET',
  'TZOFFSETFROM:+0200',
  'TZOFFSETTO:+0100',
  'DTSTART:19701025T030000',
  'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
  'END:STANDARD',
  'END:VTIMEZONE',
  'BEGIN:VEVENT',
  'ORGANIZER;CN=TICKETPORTAL:MAILTO:help@ticketportal.cz',
  'UID::ticketportal.performance.120009907',
  'DTSTAMP:20261005T162614Z',
  'DTSTART;TZID=Europe/Prague:20261019T190000',
  'DTEND;TZID=Europe/Prague:20261019T213000',
  'SUMMARY:HELENA Forever  19.10.2026 19:00',
  'LOCATION:BOBYHALL, Sportovní 559/2A, Brno',
  'GEO:49.212292;16.608061',
  'DESCRIPTION:Přidejte si do Vašeho kalendáře.',
  'URL:https://www.ticketportal.sk/event/12005754',
  'CATEGORIES:Koncerty - Pop',
  'BEGIN:VALARM',
  'ACTION:DISPLAY',
  'DESCRIPTION:Přidejte si připomínku do kalendáře.',
  'END:VALARM',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

const REAL_TICKETPORTAL_CZ_FROM_HEADER = '"Vstupenky Ticketportal" <vstupenky@ticketportal.cz>';

// --- local fakes/harness, mirroring test/ics-cancelled-vevent.test.js ------

function fakeAttachment(name, contentType, data) {
  return {
    getName: function () {
      return name;
    },
    getContentType: function () {
      return contentType || '';
    },
    getDataAsString: function () {
      return data;
    },
  };
}

function fakeIcsMessage(fromHeader, icsText) {
  return {
    getFrom: function () {
      return fromHeader;
    },
    getAttachments: function () {
      return [fakeAttachment('event.ics', 'text/calendar', icsText)];
    },
  };
}

function fakeThread(messages) {
  return {
    getId: function () {
      return 'thread-ticketportal-exclusion-1';
    },
    getMessages: function () {
      return messages;
    },
  };
}

function withFakeGasGlobals(options, body) {
  const opts = options || {};
  const existingByUid = opts.existingByUid || {};
  const state = { imported: [], inserted: [], removed: [], listCalls: [] };

  global.CONFIG = { calendarId: 'primary' };
  global.CalendarApp = {
    getCalendarById: function (id) {
      return { id: id, getTimeZone: function () { return 'Europe/Prague'; } };
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

// --- (g, D-07) isTicketingPortalSender --------------------------------------

test('quick-261005-orv: isTicketingPortalSender matches the shipped Ticketportal.cz sender case-insensitively, and is false for an unrelated sender, an empty list, or a null list (D-07)', () => {
  const shippedPortals = TICKETING_PORTALS_ACTION.config.ticketingPortals;

  assert.equal(isTicketingPortalSender(REAL_TICKETPORTAL_CZ_FROM_HEADER, shippedPortals), true);
  assert.equal(isTicketingPortalSender('"VSTUPENKY TICKETPORTAL" <VSTUPENKY@TICKETPORTAL.CZ>', shippedPortals), true);
  assert.equal(isTicketingPortalSender('RegioJet <jizdenky@regiojet.cz>', shippedPortals), false);
  assert.equal(isTicketingPortalSender(REAL_TICKETPORTAL_CZ_FROM_HEADER, []), false);
  assert.equal(isTicketingPortalSender(REAL_TICKETPORTAL_CZ_FROM_HEADER, null), false);
});

// --- (g) PRECONDITION + DEFECT ----------------------------------------------

test('quick-261005-orv: PRECONDITION -- ICS_CALENDAR_ACTION.config.excludeFrom is at its code default []; DEFECT -- appliesTo on a one-message Ticketportal.cz thread carrying the real .ics is false (D-07)', () => {
  assert.deepEqual(ICS_CALENDAR_ACTION.config.excludeFrom, []);

  const thread = fakeThread([fakeIcsMessage(REAL_TICKETPORTAL_CZ_FROM_HEADER, REAL_TICKETPORTAL_CZ_ICS_TEXT)]);
  assert.equal(ICS_CALENDAR_ACTION.appliesTo(thread), false);
});

// --- (g) run() on a mixed thread --------------------------------------------

test('quick-261005-orv: run() on a two-message thread (ticketportal + an ordinary sender, both carrying the real .ics) imports exactly ONE event -- the ordinary sender\'s (D-07)', () => {
  const ticketportalMessage = fakeIcsMessage(REAL_TICKETPORTAL_CZ_FROM_HEADER, REAL_TICKETPORTAL_CZ_ICS_TEXT);
  const aliceMessage = fakeIcsMessage('Alice Example <alice@example.com>', REAL_TICKETPORTAL_CZ_ICS_TEXT);
  const thread = fakeThread([ticketportalMessage, aliceMessage]);

  const state = withFakeGasGlobals({}, function () {
    ICS_CALENDAR_ACTION.run(thread);
  });

  assert.equal(state.imported.length, 1);
  assert.equal(state.inserted.length, 0);
});

// --- (g) a ticketportal-only thread ------------------------------------------

test('quick-261005-orv: run() on a ticketportal-ONLY thread throws the existing "No .ics attachment found on thread" error, with zero imports/inserts (D-07)', () => {
  const thread = fakeThread([fakeIcsMessage(REAL_TICKETPORTAL_CZ_FROM_HEADER, REAL_TICKETPORTAL_CZ_ICS_TEXT)]);

  let thrown = null;
  const state = withFakeGasGlobals({}, function () {
    try {
      ICS_CALENDAR_ACTION.run(thread);
    } catch (e) {
      thrown = e;
    }
  });

  assert.ok(thrown, 'run() must throw for a thread with no admitted .ics attachment');
  assert.match(String(thrown.message), /No \.ics attachment found on thread/);
  assert.equal(state.imported.length, 0);
  assert.equal(state.inserted.length, 0);
});

// --- (g, generality) every shipped portal, not just Ticketportal.cz --------

test('quick-261005-orv: EVERY shipped TICKETING_PORTALS entry is structurally excluded -- a one-message thread with the real .ics gives appliesTo false for each (D-07, generality)', () => {
  const shippedPortals = TICKETING_PORTALS_ACTION.config.ticketingPortals;
  assert.equal(shippedPortals.length, 6);

  shippedPortals.forEach(function (portal) {
    const thread = fakeThread([fakeIcsMessage('"Some Sender" <' + portal.identifyingEmail + '>', REAL_TICKETPORTAL_CZ_ICS_TEXT)]);
    assert.equal(ICS_CALENDAR_ACTION.appliesTo(thread), false, 'expected appliesTo false for ' + portal.identifyingEmail);
  });
});

// --- (g, config-keyed, not hardcoded) ---------------------------------------

test('quick-261005-orv: the exclusion follows the CONFIGURED TICKETING_PORTALS list, not a hardcoded address -- restricting that list to only enigoo.cz re-admits the Ticketportal.cz sender (D-07, config-keyed)', () => {
  withFakePropertiesService(
    { '07-action-ticketing-portals-TICKETING_PORTALS': '[{"identifyingEmail":"no-reply@enigoo.cz","calendarId":null,"insertPdfIntoEvent":false}]' },
    function () {
      const thread = fakeThread([fakeIcsMessage(REAL_TICKETPORTAL_CZ_FROM_HEADER, REAL_TICKETPORTAL_CZ_ICS_TEXT)]);
      assert.equal(ICS_CALENDAR_ACTION.appliesTo(thread), true);
    }
  );
});

// --- (g, REGRESSION) ordinary senders are unaffected ------------------------

test('quick-261005-orv: REGRESSION -- five ordinary (non-ticketing) senders still get their .ics imported normally, and the pre-existing excludeFrom path still works (D-07)', () => {
  const ordinaryFromHeaders = [
    'RegioJet <jizdenky@regiojet.cz>',
    'Alice Example <alice@example.com>',
    'boss@work.com',
    'Calendar Bot <calendar-noreply@example.org>',
    'Bob <BOB@EXAMPLE.NET>',
  ];

  ordinaryFromHeaders.forEach(function (fromHeader) {
    const thread = fakeThread([fakeIcsMessage(fromHeader, REAL_TICKETPORTAL_CZ_ICS_TEXT)]);
    assert.equal(ICS_CALENDAR_ACTION.appliesTo(thread), true, 'expected appliesTo true for ' + fromHeader);

    const state = withFakeGasGlobals({}, function () {
      ICS_CALENDAR_ACTION.run(thread);
    });
    assert.equal(state.imported.length, 1, 'expected exactly 1 import for ' + fromHeader);
  });

  // The pre-existing excludeFrom gate still works, independently of this
  // STRUCTURAL exclusion.
  withFakePropertiesService({ '05-action-ics-EXCLUDE_FROM': 'jizdenky@regiojet.cz' }, function () {
    const regiojetThread = fakeThread([fakeIcsMessage('RegioJet <jizdenky@regiojet.cz>', REAL_TICKETPORTAL_CZ_ICS_TEXT)]);
    assert.equal(ICS_CALENDAR_ACTION.appliesTo(regiojetThread), false);
  });
});

// --- (load order) both-load-order child-process tests -----------------------
//
// A same-process require() only ever resolves a module once, so the static
// requires at the top of this file prove ONLY the 05-first order. These two
// tests spawn a FRESH Node process for each order, resolving both src files
// by ABSOLUTE path, so each process genuinely exercises its own first-require
// winner for the 05<->07 circular require (see both files' end-of-file Node
// bridge comments for why the bridges MUST sit after module.exports).

function buildLoadOrderScript(firstPath, secondPath) {
  return [
    'const first = require(' + JSON.stringify(firstPath) + ');',
    'const second = require(' + JSON.stringify(secondPath) + ');',
    'const okTypes =',
    '  typeof globalThis.parseIcs === "function" &&',
    '  typeof globalThis.isIcsAttachment === "function" &&',
    '  typeof globalThis.resolveTicketingPortal === "function" &&',
    '  typeof globalThis.TICKETING_PORTALS_ACTION === "object";',
    'const ICS_CALENDAR_ACTION = (first && first.ICS_CALENDAR_ACTION) || (second && second.ICS_CALENDAR_ACTION);',
    'const fakeThread = {',
    '  getMessages: function () {',
    '    return [{',
    '      getFrom: function () { return ' + JSON.stringify(REAL_TICKETPORTAL_CZ_FROM_HEADER) + '; },',
    '      getAttachments: function () {',
    '        return [{',
    '          getName: function () { return "event.ics"; },',
    '          getContentType: function () { return "text/calendar"; },',
    '          getDataAsString: function () { return ' + JSON.stringify(REAL_TICKETPORTAL_CZ_ICS_TEXT) + '; },',
    '        }];',
    '      },',
    '    }];',
    '  },',
    '};',
    'const appliesToFalse = ICS_CALENDAR_ACTION.appliesTo(fakeThread) === false;',
    'if (okTypes && appliesToFalse) { process.stdout.write("OK"); } else { process.stdout.write("FAIL:" + JSON.stringify({ okTypes: okTypes, appliesToFalse: appliesToFalse })); }',
  ].join('\n');
}

test('quick-261005-orv: LOAD ORDER -- requiring 05 then 07 in a fresh process wires both cross-file bridges, and ICS_CALENDAR_ACTION.appliesTo correctly excludes a Ticketportal.cz thread (D-07, load order)', () => {
  const icsPath = path.join(__dirname, '..', 'src', '05-action-ics-import.js');
  const ticketingPath = path.join(__dirname, '..', 'src', '07-action-ticketing-portals.js');
  const script = buildLoadOrderScript(icsPath, ticketingPath);

  const stdout = require('node:child_process').execFileSync(process.execPath, ['-e', script], { encoding: 'utf8' });
  assert.equal(stdout.trim(), 'OK');
});

test('quick-261005-orv: LOAD ORDER -- requiring 07 then 05 in a fresh process wires both cross-file bridges, and ICS_CALENDAR_ACTION.appliesTo correctly excludes a Ticketportal.cz thread (D-07, load order)', () => {
  const ticketingPath = path.join(__dirname, '..', 'src', '07-action-ticketing-portals.js');
  const icsPath = path.join(__dirname, '..', 'src', '05-action-ics-import.js');
  const script = buildLoadOrderScript(ticketingPath, icsPath);

  const stdout = require('node:child_process').execFileSync(process.execPath, ['-e', script], { encoding: 'utf8' });
  assert.equal(stdout.trim(), 'OK');
});

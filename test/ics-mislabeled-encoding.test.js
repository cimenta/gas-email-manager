'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  parseIcs,
  normalizeIcsText,
  isIcsText,
  decodeBase64Utf8,
  ICS_CALENDAR_ACTION,
} = require('../src/05-action-ics-import.js');

// ---------------------------------------------------------------------------
// MISLABELED CONTENT-TRANSFER-ENCODING — regression coverage for the
// live-reported bug `linkedin-ics-not-imported`.
//
// THE BUG (two contributing causes, both required — see the debug session's
// and_gate note):
//
//   1. LinkedIn's "You're attending ..." event mail carries a genuine,
//      well-formed LinkedInEvent.ics, but its MIME part declares
//      `Content-Transfer-Encoding: 7bit` while the part body is actually
//      base64 text. Any spec-honouring MIME parser (Gmail included) takes the
//      declaration at its word and performs NO decoding, so
//      GmailAttachment#getDataAsString() hands back the literal base64 string
//      rather than iCalendar text.
//
//   2. parseIcs found no `BEGIN:VEVENT` line in that base64 text and returned
//      [] WITHOUT throwing, and ICS_CALENDAR_ACTION.run treated an empty event
//      array as success. So run() wrote nothing, threw nothing,
//      dispatchActions reported hadError=false, and the thread was labeled
//      processed. Silent, unreported data loss: no event, no failed label, no
//      owner notification.
//
// ORACLE TYPE: specified (both halves). The required behavior is stated
// directly — a valid iCalendar attachment must produce its event regardless of
// a sender's mislabeled transfer encoding, and an attachment that is not
// iCalendar at all must fail LOUDLY rather than vanish. These assert those
// exact outcomes, not merely "does not crash".
//
// THE CONSTRAINT THAT MUST SURVIVE THE FIX: a VALID VCALENDAR carrying zero
// VEVENTs must still parse to [] silently (test/ics-parser.test.js:132 —
// legitimate for cancellations and VTODO/VFREEBUSY-only calendars). The
// loud-failure guard therefore keys on "not iCalendar AT ALL", never on
// "empty event list".
// ---------------------------------------------------------------------------

// Reconstructed from the debug session's Evidence entry 1, which recorded the
// real part's structure: VERSION:2.0, CALSCALE:GREGORIAN, and one VEVENT with
// DTSTAMP/DTSTART/DTEND/SUMMARY/LOCATION/URL/UID. Identifying values replaced
// with fictional equivalents, per this repo's fixture convention.
const REAL_LINKEDIN_ICS = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'CALSCALE:GREGORIAN',
  'BEGIN:VEVENT',
  'DTSTAMP:20260901T090000Z',
  'DTSTART:20260910T160000Z',
  'DTEND:20260910T170000Z',
  'SUMMARY:Building AI Products That Ship',
  'LOCATION:LinkedIn Live',
  'URL:https://www.linkedin.com/events/7123456789012345678/',
  'UID:7123456789012345678@linkedin.com',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

// What the part ACTUALLY carries on the wire while declaring 7bit.
const LINKEDIN_WIRE_BODY = Buffer.from(REAL_LINKEDIN_ICS, 'utf8').toString('base64');

// --- decodeBase64Utf8 (pure) ------------------------------------------------
//
// Implemented in-repo rather than via Utilities.base64Decode so it stays
// Node-testable like the rest of the parser (Utilities is a GAS-only global).

test('decodeBase64Utf8: decodes plain ASCII base64', () => {
  assert.equal(decodeBase64Utf8(Buffer.from('BEGIN:VCALENDAR', 'utf8').toString('base64')), 'BEGIN:VCALENDAR');
});

test('decodeBase64Utf8: decodes multi-byte UTF-8 correctly (Czech diacritics — this codebase handles Czech senders)', () => {
  const source = 'SUMMARY:Schůzka s Evou Dvořákovou — příští týden';
  assert.equal(decodeBase64Utf8(Buffer.from(source, 'utf8').toString('base64')), source);
});

test('decodeBase64Utf8: tolerates the CRLF line wrapping real MIME base64 bodies use (76-char lines)', () => {
  const wrapped = LINKEDIN_WIRE_BODY.replace(/(.{76})/g, '$1\r\n');
  assert.equal(decodeBase64Utf8(wrapped), REAL_LINKEDIN_ICS);
});

test('decodeBase64Utf8: BOUNDARY both padding forms round-trip ("=" and "==")', () => {
  const onePad = Buffer.from('BEGIN:VCALENDA', 'utf8').toString('base64');
  const twoPad = Buffer.from('BEGIN:VCALEND', 'utf8').toString('base64');
  assert.ok(onePad.endsWith('='), 'fixture should carry one pad char');
  assert.ok(twoPad.endsWith('=='), 'fixture should carry two pad chars');
  assert.equal(decodeBase64Utf8(onePad), 'BEGIN:VCALENDA');
  assert.equal(decodeBase64Utf8(twoPad), 'BEGIN:VCALEND');
});

test('decodeBase64Utf8: returns null for text outside the base64 alphabet rather than emitting garbage', () => {
  assert.equal(decodeBase64Utf8('BEGIN:VCALENDAR\r\nVERSION:2.0'), null);
});

test('decodeBase64Utf8: BOUNDARY returns null for empty input', () => {
  assert.equal(decodeBase64Utf8(''), null);
});

test('decodeBase64Utf8: returns null when length is not a multiple of 4 (truncated base64)', () => {
  assert.equal(decodeBase64Utf8('QUJDRQ'), null);
});

// --- isIcsText (pure) -------------------------------------------------------

test('isIcsText: true for real iCalendar text', () => {
  assert.equal(isIcsText(REAL_LINKEDIN_ICS), true);
});

test('isIcsText: false for the base64 wire body (this is exactly what run() must catch)', () => {
  assert.equal(isIcsText(LINKEDIN_WIRE_BODY), false);
});

test('isIcsText: false for empty/null input', () => {
  assert.equal(isIcsText(''), false);
  assert.equal(isIcsText(null), false);
});

test('isIcsText: case-insensitive — BEGIN:vcalendar still counts', () => {
  assert.equal(isIcsText('begin:vcalendar\r\nend:vcalendar'), true);
});

// --- normalizeIcsText (pure) ------------------------------------------------

test('normalizeIcsText: already-decoded iCalendar passes through completely untouched', () => {
  assert.equal(normalizeIcsText(REAL_LINKEDIN_ICS), REAL_LINKEDIN_ICS);
});

test('normalizeIcsText: base64-encoded iCalendar is recovered to its decoded form', () => {
  assert.equal(normalizeIcsText(LINKEDIN_WIRE_BODY), REAL_LINKEDIN_ICS);
});

test('normalizeIcsText: SAFETY — valid base64 that decodes to something which is NOT iCalendar is left as-is, never substituted', () => {
  // "Hello world" is perfectly valid base64 input. The recovery must be gated
  // on the DECODED output being iCalendar, otherwise it would corrupt
  // unrelated attachment text.
  const notCalendar = Buffer.from('Hello world', 'utf8').toString('base64');
  assert.equal(normalizeIcsText(notCalendar), notCalendar);
});

test('normalizeIcsText: non-base64 garbage is left as-is', () => {
  assert.equal(normalizeIcsText('this is not a calendar at all'), 'this is not a calendar at all');
});

// --- parseIcs: the choke-point fix ------------------------------------------

test('THE BUG: parseIcs on the base64 wire body recovers the event instead of silently returning []', () => {
  const events = parseIcs(LINKEDIN_WIRE_BODY);

  assert.equal(events.length, 1, 'the mislabeled attachment must still yield its event');
  const event = events[0];
  assert.equal(event.summary, 'Building AI Products That Ship');
  assert.equal(event.location, 'LinkedIn Live');
  assert.equal(event.uid, '7123456789012345678@linkedin.com');
  assert.equal(event.start.toISOString(), '2026-09-10T16:00:00.000Z');
  assert.equal(event.end.toISOString(), '2026-09-10T17:00:00.000Z');
});

test('parseIcs on the base64 body produces byte-identical results to parsing the decoded text', () => {
  assert.deepEqual(parseIcs(LINKEDIN_WIRE_BODY), parseIcs(REAL_LINKEDIN_ICS));
});

test('parseIcs: CRLF-wrapped base64 (real MIME wrapping) is recovered too', () => {
  const wrapped = LINKEDIN_WIRE_BODY.replace(/(.{76})/g, '$1\r\n');
  assert.equal(parseIcs(wrapped).length, 1);
});

test('REGRESSION GUARD: parseIcs on ordinary decoded iCalendar is completely unaffected', () => {
  const events = parseIcs(REAL_LINKEDIN_ICS);
  assert.equal(events.length, 1);
  assert.equal(events[0].summary, 'Building AI Products That Ship');
});

test('REGRESSION GUARD: a valid VCALENDAR with zero VEVENTs still parses to [] without throwing', () => {
  const emptyIcs = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'END:VCALENDAR'].join('\r\n');
  assert.deepEqual(parseIcs(emptyIcs), []);
});

test('parseIcs stays tolerant: non-iCalendar text still returns [] rather than throwing (the LOUD failure belongs to run(), not the parser)', () => {
  assert.deepEqual(parseIcs('not a calendar'), []);
});

// --- ICS_CALENDAR_ACTION.run: the silent-failure guard ----------------------

function fakeAttachment(name, contentType, data) {
  return {
    getName: function () {
      return name;
    },
    getContentType: function () {
      return contentType;
    },
    getDataAsString: function () {
      return data;
    },
  };
}

function fakeThread(attachments) {
  return {
    getId: function () {
      return 'thread-linkedin-1';
    },
    getMessages: function () {
      return [
        {
          getFrom: function () {
            return 'LinkedIn <messages-noreply@linkedin.com>';
          },
          getAttachments: function () {
            return attachments;
          },
        },
      ];
    },
  };
}

// Runs `body` with the GAS globals run() touches faked out, capturing every
// Calendar write. Mirrors the harness convention in
// test/existing-invite-guard.test.js and test/calendar-routing.test.js.
function withFakeGasGlobals(body) {
  const imported = [];
  const inserted = [];

  global.CONFIG = { calendarId: 'primary' };
  global.CalendarApp = {
    getCalendarById: function () {
      return { id: 'primary' };
    },
  };
  global.Calendar = {
    Events: {
      list: function () {
        return { items: [] };
      },
      import: function (resource, calendarId) {
        imported.push({ resource: resource, calendarId: calendarId });
        return { id: 'evt-new' };
      },
      insert: function (resource, calendarId) {
        inserted.push({ resource: resource, calendarId: calendarId });
        return { id: 'evt-new' };
      },
    },
  };

  try {
    body();
  } finally {
    delete global.CONFIG;
    delete global.CalendarApp;
    delete global.Calendar;
  }

  return { imported: imported, inserted: inserted };
}

test('END TO END: run() on the mislabeled LinkedIn attachment actually creates the calendar event', () => {
  const thread = fakeThread([fakeAttachment('LinkedInEvent.ics', 'text/calendar', LINKEDIN_WIRE_BODY)]);

  assert.equal(ICS_CALENDAR_ACTION.appliesTo(thread), true, 'detection was never the problem');

  const state = withFakeGasGlobals(function () {
    ICS_CALENDAR_ACTION.run(thread);
  });

  assert.equal(state.imported.length, 1, 'exactly one event must reach the calendar');
  assert.equal(state.imported[0].resource.summary, 'Building AI Products That Ship');
  assert.equal(state.imported[0].resource.iCalUID, '7123456789012345678@linkedin.com');
});

test('THE SILENT FAILURE: run() THROWS when a matched .ics attachment is not iCalendar at all, instead of labeling the thread processed', () => {
  // Not iCalendar, and not recoverable base64 either — the class of failure
  // that used to vanish without a trace.
  const thread = fakeThread([fakeAttachment('LinkedInEvent.ics', 'text/calendar', 'totally unparseable payload')]);

  withFakeGasGlobals(function () {
    assert.throws(
      function () {
        ICS_CALENDAR_ACTION.run(thread);
      },
      /not iCalendar/i,
      'must surface loudly so dispatchActions applies the failed label and notifies the owner'
    );
  });
});

test('the loud guard names the thread so the owner notification is actionable', () => {
  const thread = fakeThread([fakeAttachment('LinkedInEvent.ics', 'text/calendar', 'totally unparseable payload')]);

  withFakeGasGlobals(function () {
    assert.throws(
      function () {
        ICS_CALENDAR_ACTION.run(thread);
      },
      /thread-linkedin-1/
    );
  });
});

test('CONSTRAINT: run() does NOT throw for a VALID VCALENDAR carrying zero VEVENTs — that stays a legitimate silent no-op', () => {
  const emptyIcs = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'END:VCALENDAR'].join('\r\n');
  const thread = fakeThread([fakeAttachment('cancelled.ics', 'text/calendar', emptyIcs)]);

  const state = withFakeGasGlobals(function () {
    ICS_CALENDAR_ACTION.run(thread);
  });

  assert.equal(state.imported.length, 0);
  assert.equal(state.inserted.length, 0);
});

test('CONSTRAINT: an ordinary well-formed .ics attachment still imports exactly as before', () => {
  const thread = fakeThread([fakeAttachment('invite.ics', 'text/calendar', REAL_LINKEDIN_ICS)]);

  const state = withFakeGasGlobals(function () {
    ICS_CALENDAR_ACTION.run(thread);
  });

  assert.equal(state.imported.length, 1);
  assert.equal(state.imported[0].resource.summary, 'Building AI Products That Ship');
});

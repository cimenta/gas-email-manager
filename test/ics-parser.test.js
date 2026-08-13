'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { parseIcs, buildEventResource } = require('../src/05-action-ics-import.js');

const SINGLE_VEVENT_ICS = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'BEGIN:VEVENT',
  'SUMMARY:Team offsite',
  'DTSTART:20260801T100000Z',
  'DTEND:20260801T110000Z',
  'DESCRIPTION:Quarterly planning session',
  'LOCATION:Conference Room A',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

test('single non-recurring VEVENT parses to one event with expected fields', () => {
  const events = parseIcs(SINGLE_VEVENT_ICS);

  assert.equal(events.length, 1);
  const event = events[0];
  assert.equal(event.summary, 'Team offsite');
  assert.ok(event.start instanceof Date);
  assert.ok(event.end instanceof Date);
  assert.equal(event.isAllDay, false);
  assert.equal(event.description, 'Quarterly planning session');
  assert.equal(event.location, 'Conference Room A');
  assert.equal(event.recurrence, null);
});

// --- SEQUENCE (RFC 5545 section 3.8.7.4) -------------------------------------
//
// Live-reported bug (quick-260731-seq): a genuine Calendly-generated
// Microsoft Exchange invite for a Teams meeting, sent to the owner's own
// gmail.com address, failed Calendar.Events.import with
// "GoogleJsonResponseException: Invalid sequence value. The specified
// sequence number is below the current sequence number of the resource."
// Root cause: this parser never extracted SEQUENCE at all, so
// buildEventResource never set a sequence field on the resource -
// Calendar.Events.import implicitly sent sequence 0, lower than the REAL
// sequence number (1) Gmail's own native Gmail-to-Calendar detection had
// already stored from this same invite (same family of "our own import
// collides with Google's native detection" issue as the original iCalUID
// dedup fix). Fixture below is MODELED on the real failing VEVENT's
// structure/shape (a Calendly-generated Exchange invite's UID/SUMMARY/
// LOCATION format), with all identifying values replaced by fictional
// equivalents.

const REAL_CALENDLY_EXCHANGE_ICS_WITH_SEQUENCE = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'BEGIN:VEVENT',
  'UID:040000008200E00074C5B7101A82E008000000001AB2345CD678EF0100000000000000',
  ' 0100000004D20E8BCE179574188BC5D2FE29825B2',
  'SUMMARY;LANGUAGE=en-US:Jan Novák and Eva Dvořáková',
  'DTSTART:20260803T150000Z',
  'DTEND:20260803T153000Z',
  'DTSTAMP:20260731T084147Z',
  'TRANSP:OPAQUE',
  'STATUS:CONFIRMED',
  'SEQUENCE:1',
  'LOCATION;LANGUAGE=en-US:https://calendly.com/events/example-meeting-id/microsoft_teams',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

test('a VEVENT carrying SEQUENCE:1 (real Calendly/Exchange invite fixture) parses to event.sequence === 1', () => {
  const events = parseIcs(REAL_CALENDLY_EXCHANGE_ICS_WITH_SEQUENCE);

  assert.equal(events.length, 1);
  assert.equal(events[0].sequence, 1);
});

test('a VEVENT with no SEQUENCE line at all parses to event.sequence === 0 (RFC 5545 3.8.7.4 documented default)', () => {
  const events = parseIcs(SINGLE_VEVENT_ICS);

  assert.equal(events[0].sequence, 0);
});

test('folded (continuation) lines are unfolded before parsing', () => {
  const foldedIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Folded test',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'DESCRIPTION:This is a long description that has been folded across',
    '  multiple continuation lines per RFC 5545 line folding rules.',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(foldedIcs);
  assert.equal(events.length, 1);
  assert.equal(
    events[0].description,
    'This is a long description that has been folded across multiple continuation lines per RFC 5545 line folding rules.'
  );
});

// Regression coverage for WR-02: an escaped literal backslash immediately
// followed by a literal "n" (raw bytes \, \, n) must not be misread as an
// escaped backslash plus an escaped newline. Applying the \n rule before the
// \\ rule (or any other multi-pass ordering) corrupts this case.
test('unescapeText handles an escaped backslash immediately followed by a literal "n" (WR-02)', () => {
  const backslashIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Backslash test',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'DESCRIPTION:C:\\\\notes\\\\nested',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(backslashIcs);
  assert.equal(events[0].description, 'C:\\notes\\nested');
});

test('start/end are correct absolute UTC instants for Z-suffixed times', () => {
  const events = parseIcs(SINGLE_VEVENT_ICS);
  const event = events[0];

  assert.equal(event.start.toISOString(), '2026-08-01T10:00:00.000Z');
  assert.equal(event.end.toISOString(), '2026-08-01T11:00:00.000Z');
});

test('an .ics with zero VEVENTs parses to an empty array without throwing', () => {
  const emptyIcs = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'END:VCALENDAR'].join('\r\n');

  assert.deepEqual(parseIcs(emptyIcs), []);
});

test('multiple VEVENTs parse into an array in source order', () => {
  const multiIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:First',
    'DTSTART:20260801T090000Z',
    'DTEND:20260801T093000Z',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'SUMMARY:Second',
    'DTSTART:20260802T090000Z',
    'DTEND:20260802T093000Z',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'SUMMARY:Third',
    'DTSTART:20260803T090000Z',
    'DTEND:20260803T093000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(multiIcs);
  assert.equal(events.length, 3);
  assert.deepEqual(
    events.map((e) => e.summary),
    ['First', 'Second', 'Third']
  );
});

test('two VEVENTs with identical DTSTART/DTEND parse to two separate objects (no merge)', () => {
  const adjacencyIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Duplicate slot A',
    'DTSTART:20260801T090000Z',
    'DTEND:20260801T093000Z',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'SUMMARY:Duplicate slot B',
    'DTSTART:20260801T090000Z',
    'DTEND:20260801T093000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(adjacencyIcs);
  assert.equal(events.length, 2);
  assert.notEqual(events[0], events[1]);
  assert.equal(events[0].summary, 'Duplicate slot A');
  assert.equal(events[1].summary, 'Duplicate slot B');
});

test('all-day VEVENT (DTSTART;VALUE=DATE) parses to isAllDay:true with date-only start/end', () => {
  const allDayIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Company holiday',
    'DTSTART;VALUE=DATE:20260901',
    'DTEND;VALUE=DATE:20260902',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(allDayIcs);
  assert.equal(events.length, 1);
  const event = events[0];
  assert.equal(event.isAllDay, true);
  assert.equal(event.start.toISOString(), '2026-09-01T00:00:00.000Z');
  assert.equal(event.end.toISOString(), '2026-09-02T00:00:00.000Z');
});

test('recurring VEVENT with a supported RRULE parses to a normalized recurrence descriptor', () => {
  const recurringIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Standup',
    'DTSTART:20260803T090000Z',
    'DTEND:20260803T091500Z',
    'RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=5;BYDAY=MO,WE',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(recurringIcs);
  assert.equal(events.length, 1);
  assert.deepEqual(events[0].recurrence, {
    freq: 'WEEKLY',
    interval: 2,
    count: 5,
    until: null,
    byDay: ['MO', 'WE'],
  });
});

test('recurring VEVENT with an unsupported sub-daily FREQ throws a controlled error', () => {
  const secondlyIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Pathological',
    'DTSTART:20260803T090000Z',
    'DTEND:20260803T091500Z',
    'RRULE:FREQ=SECONDLY;INTERVAL=1',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  assert.throws(() => parseIcs(secondlyIcs));

  const minutelyIcs = secondlyIcs.replace('FREQ=SECONDLY', 'FREQ=MINUTELY');
  assert.throws(() => parseIcs(minutelyIcs));
});

test('recurring VEVENT with UNTIL parses descriptor.until as a Date', () => {
  const untilIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Daily check-in',
    'DTSTART:20260803T090000Z',
    'DTEND:20260803T091500Z',
    'RRULE:FREQ=DAILY;UNTIL=20260810T090000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(untilIcs);
  assert.ok(events[0].recurrence.until instanceof Date);
  assert.equal(events[0].recurrence.until.toISOString(), '2026-08-10T09:00:00.000Z');
});

// Regression coverage for a real owner-reported bug (deviation from Plan
// 03-01's flagged_assumption): a TZID wall-clock DTSTART/DTEND was being
// treated as literal UTC, producing events 2 hours off during CEST. The
// VTIMEZONE block below is copied verbatim from the owner's real
// Exchange/Outlook .ics (a standard RFC 5545 STANDARD/DAYLIGHT pair with
// BYDAY/BYMONTH transition rules) — resolution must come from THIS embedded
// data, never from a hardcoded Windows-name mapping or the script's own
// timezone.
const CET_VTIMEZONE = [
  'BEGIN:VTIMEZONE',
  'TZID:Central Europe Standard Time',
  'BEGIN:STANDARD',
  'DTSTART:16010101T030000',
  'TZOFFSETFROM:+0200',
  'TZOFFSETTO:+0100',
  'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=10',
  'END:STANDARD',
  'BEGIN:DAYLIGHT',
  'DTSTART:16010101T020000',
  'TZOFFSETFROM:+0100',
  'TZOFFSETTO:+0200',
  'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=3',
  'END:DAYLIGHT',
  'END:VTIMEZONE',
].join('\r\n');

test('TZID wall-clock DTSTART during DST (DAYLIGHT period) resolves via embedded VTIMEZONE (owner-reported bug regression)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CET_VTIMEZONE,
    'BEGIN:VEVENT',
    'SUMMARY:CET meeting (summer)',
    'DTSTART;TZID=Central Europe Standard Time:20260722T140000',
    'DTEND;TZID=Central Europe Standard Time:20260722T150000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  assert.equal(events.length, 1);
  // 14:00 CEST (UTC+2) -> 12:00 UTC
  assert.equal(events[0].start.toISOString(), '2026-07-22T12:00:00.000Z');
  assert.equal(events[0].end.toISOString(), '2026-07-22T13:00:00.000Z');
});

test('TZID wall-clock DTSTART outside DST (STANDARD period) resolves via embedded VTIMEZONE', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CET_VTIMEZONE,
    'BEGIN:VEVENT',
    'SUMMARY:CET meeting (winter)',
    'DTSTART;TZID=Central Europe Standard Time:20261215T140000',
    'DTEND;TZID=Central Europe Standard Time:20261215T150000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  assert.equal(events.length, 1);
  // 14:00 CET (UTC+1) -> 13:00 UTC
  assert.equal(events[0].start.toISOString(), '2026-12-15T13:00:00.000Z');
  assert.equal(events[0].end.toISOString(), '2026-12-15T14:00:00.000Z');
});

// Regression coverage for CR-01: the DST-boundary window itself (roughly
// the 1-2 hours immediately surrounding each seasonal transition), where
// comparing a wall-clock-literal instant against a real (offset-corrected)
// UTC transition instant previously flipped the classification and produced
// a UTC result off by exactly the transition's UTC offset.
test('TZID wall-clock DTSTART just before the spring-forward transition resolves to STANDARD offset (CR-01 boundary)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CET_VTIMEZONE,
    'BEGIN:VEVENT',
    'SUMMARY:Just before spring-forward',
    'DTSTART;TZID=Central Europe Standard Time:20260329T013000',
    'DTEND;TZID=Central Europe Standard Time:20260329T013000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  // 01:30 local, before the 02:00 spring-forward -> still CET (+1)
  assert.equal(events[0].start.toISOString(), '2026-03-29T00:30:00.000Z');
});

test('TZID wall-clock DTSTART just after the spring-forward transition resolves to DAYLIGHT offset (CR-01 boundary)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CET_VTIMEZONE,
    'BEGIN:VEVENT',
    'SUMMARY:Just after spring-forward',
    'DTSTART;TZID=Central Europe Standard Time:20260329T033000',
    'DTEND;TZID=Central Europe Standard Time:20260329T033000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  // 03:30 local, after the 02:00 spring-forward -> already CEST (+2)
  assert.equal(events[0].start.toISOString(), '2026-03-29T01:30:00.000Z');
});

test('TZID wall-clock DTSTART just before the fall-back transition resolves to DAYLIGHT offset (CR-01 boundary)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CET_VTIMEZONE,
    'BEGIN:VEVENT',
    'SUMMARY:Just before fall-back',
    'DTSTART;TZID=Central Europe Standard Time:20261025T013000',
    'DTEND;TZID=Central Europe Standard Time:20261025T013000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  // 01:30 local, before the 03:00 fall-back -> still CEST (+2)
  assert.equal(events[0].start.toISOString(), '2026-10-24T23:30:00.000Z');
});

test('TZID wall-clock DTSTART just after the fall-back transition resolves to STANDARD offset (CR-01 boundary)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CET_VTIMEZONE,
    'BEGIN:VEVENT',
    'SUMMARY:Just after fall-back',
    'DTSTART;TZID=Central Europe Standard Time:20261025T033000',
    'DTEND;TZID=Central Europe Standard Time:20261025T033000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  // 03:30 local, after the 03:00 fall-back -> back to CET (+1)
  assert.equal(events[0].start.toISOString(), '2026-10-25T02:30:00.000Z');
});

// Regression coverage for WR-01: a VTIMEZONE with only a STANDARD sub-block
// (no DAYLIGHT), as commonly emitted by Exchange/Outlook for non-DST TZIDs
// (e.g. "China Standard Time"), must still resolve via its fixed offset
// rather than silently falling back to literal-as-UTC treatment.
const CST_VTIMEZONE_NO_DST = [
  'BEGIN:VTIMEZONE',
  'TZID:China Standard Time',
  'BEGIN:STANDARD',
  'DTSTART:16010101T000000',
  'TZOFFSETFROM:+0800',
  'TZOFFSETTO:+0800',
  'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=1',
  'END:STANDARD',
  'END:VTIMEZONE',
].join('\r\n');

test('TZID with a STANDARD-only VTIMEZONE (no DAYLIGHT) resolves via the fixed offset (WR-01)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    CST_VTIMEZONE_NO_DST,
    'BEGIN:VEVENT',
    'SUMMARY:Beijing meeting',
    'DTSTART;TZID=China Standard Time:20260722T140000',
    'DTEND;TZID=China Standard Time:20260722T150000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  assert.equal(events.length, 1);
  // 14:00 CST (UTC+8) -> 06:00 UTC
  assert.equal(events[0].start.toISOString(), '2026-07-22T06:00:00.000Z');
  assert.equal(events[0].end.toISOString(), '2026-07-22T07:00:00.000Z');
});

// Regression coverage for IN-02: a VTIMEZONE transition rule with an
// invalid BYDAY ordinal (0, not valid per RFC 5545) or an out-of-range nth
// occurrence (e.g. "5th Monday" in a month with only 4) must throw rather
// than silently mis-computing or rolling into an adjacent month.
function buildBadOrdinalVtimezoneIcs(byDay) {
  const vtimezone = [
    'BEGIN:VTIMEZONE',
    'TZID:Bad Ordinal TZ',
    'BEGIN:STANDARD',
    'DTSTART:16010101T030000',
    'TZOFFSETFROM:+0200',
    'TZOFFSETTO:+0100',
    'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=' + byDay + ';BYMONTH=10',
    'END:STANDARD',
    'BEGIN:DAYLIGHT',
    'DTSTART:16010101T020000',
    'TZOFFSETFROM:+0100',
    'TZOFFSETTO:+0200',
    'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=3',
    'END:DAYLIGHT',
    'END:VTIMEZONE',
  ].join('\r\n');

  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    vtimezone,
    'BEGIN:VEVENT',
    'SUMMARY:Bad ordinal test',
    'DTSTART;TZID=Bad Ordinal TZ:20261025T013000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');
}

test('VTIMEZONE transition rule with BYDAY ordinal 0 throws rather than silently mis-computing (IN-02)', () => {
  assert.throws(() => parseIcs(buildBadOrdinalVtimezoneIcs('0SU')));
});

test('VTIMEZONE transition rule with an out-of-range nth-weekday ordinal throws rather than rolling into another month (IN-02)', () => {
  assert.throws(() => parseIcs(buildBadOrdinalVtimezoneIcs('5MO')));
});

test('TZID with no matching VTIMEZONE block falls back to literal-as-UTC treatment (documented v1 limitation preserved)', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Unknown TZID',
    'DTSTART;TZID=Some/Unmapped:20260722T140000',
    'DTEND;TZID=Some/Unmapped:20260722T150000',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  assert.equal(events.length, 1);
  assert.equal(events[0].start.toISOString(), '2026-07-22T14:00:00.000Z');
});

// --- STATUS (RFC 5545 section 3.8.1.11, RegioJet cancellation detection, D-01/D-02) ----
//
// parseVeventBlock gains a `status` field: the VEVENT's own STATUS property
// value, trimmed and uppercased, or null when the property is absent or
// empty after trimming. This is the single normalization point every
// downstream consumer (cancellation detection in TRANSPORT_TICKETS_ACTION)
// compares against the uppercase token only. buildEventResource must NOT
// copy this field onto the Calendar API resource (verified in the
// event-resource suite, which stays untouched).

function buildStatusIcs(statusLine) {
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'BEGIN:VEVENT',
    'SUMMARY:Status test event',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    statusLine,
    'END:VEVENT',
    'END:VCALENDAR',
  ]
    .filter(function (line) {
      return line !== null;
    })
    .join('\r\n');
}

test('a VEVENT carrying STATUS:CANCELLED parses to event.status === "CANCELLED"', () => {
  const events = parseIcs(buildStatusIcs('STATUS:CANCELLED'));
  assert.equal(events.length, 1);
  assert.equal(events[0].status, 'CANCELLED');
});

test('a VEVENT with no STATUS line at all parses to event.status === null', () => {
  const events = parseIcs(buildStatusIcs(null));
  assert.equal(events.length, 1);
  assert.equal(events[0].status, null);
});

test('a VEVENT carrying mixed-case STATUS:Cancelled parses to event.status === "CANCELLED" (parser is the single normalization point)', () => {
  const events = parseIcs(buildStatusIcs('STATUS:Cancelled'));
  assert.equal(events.length, 1);
  assert.equal(events[0].status, 'CANCELLED');
});

test('a VEVENT carrying STATUS:CONFIRMED parses to event.status === "CONFIRMED" (general STATUS value, not a cancellation boolean)', () => {
  const events = parseIcs(buildStatusIcs('STATUS:CONFIRMED'));
  assert.equal(events.length, 1);
  assert.equal(events[0].status, 'CONFIRMED');
});

test('a VEVENT whose STATUS value is whitespace-only parses to event.status === null', () => {
  const events = parseIcs(buildStatusIcs('STATUS:   '));
  assert.equal(events.length, 1);
  assert.equal(events[0].status, null);
});

// --- DTSTAMP (RFC 5545 section 3.8.7.2, RegioJet cancel/rebook staleness ----
// --- detection, D-09/D-10/D-11 of quick-260813-dq2 Task 3) -----------------
//
// parseVeventBlock gains a `dtstamp` field: the VEVENT's own DTSTAMP
// property, parsed through the EXISTING parseIcsDate helper (RFC 5545
// DTSTAMP is always UTC Z-suffixed, so no TZID is ever passed) into a real
// Date when present, or null when the property is absent OR when
// parseIcsDate throws on an unrecognized value -- a malformed DTSTAMP must
// never crash an otherwise-valid import, same discipline as the SEQUENCE
// fallback. buildEventResource must NOT copy this field onto the Calendar
// API resource either (the D-01 firewall now covers both new parser
// fields). Reuses buildStatusIcs' inline-fixture shape above (it emits NO
// DTSTAMP line at all, so it is already the "absent" case).

function buildDtstampIcs(dtstampLine) {
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'BEGIN:VEVENT',
    'SUMMARY:Dtstamp test event',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    dtstampLine,
    'END:VEVENT',
    'END:VCALENDAR',
  ]
    .filter(function (line) {
      return line !== null;
    })
    .join('\r\n');
}

test('a VEVENT carrying DTSTAMP:20260813T073445Z parses to event.dtstamp as a real Date matching that instant', () => {
  const events = parseIcs(buildDtstampIcs('DTSTAMP:20260813T073445Z'));
  assert.equal(events.length, 1);
  assert.ok(events[0].dtstamp instanceof Date);
  assert.equal(events[0].dtstamp.toISOString(), '2026-08-13T07:34:45.000Z');
});

test('a VEVENT with no DTSTAMP line at all parses to event.dtstamp === null', () => {
  const events = parseIcs(buildDtstampIcs(null));
  assert.equal(events.length, 1);
  assert.equal(events[0].dtstamp, null);
});

test('a VEVENT carrying a malformed DTSTAMP parses to event.dtstamp === null and does NOT throw -- the rest of the event still parses correctly', () => {
  const events = parseIcs(buildDtstampIcs('DTSTAMP:not-a-date'));
  assert.equal(events.length, 1);
  assert.equal(events[0].dtstamp, null);
  assert.equal(events[0].summary, 'Dtstamp test event');
});

test('buildEventResource on a DTSTAMP-carrying parsed event produces a resource with no dtstamp key (D-01 firewall covers both new parser fields)', () => {
  const events = parseIcs(buildDtstampIcs('DTSTAMP:20260813T073445Z'));
  const resource = buildEventResource(events[0]);
  assert.equal(Object.prototype.hasOwnProperty.call(resource, 'dtstamp'), false);
});

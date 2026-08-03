'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { parseIcs, buildRRuleString, buildEventResource } = require('../src/05-action-ics-import.js');

// --- UID extraction (via parseIcs) ---------------------------------------

test('VEVENT carrying UID parses events[0].uid as the raw UID value', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Invite',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'UID:invite-abc@example.com',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  assert.equal(events[0].uid, 'invite-abc@example.com');
});

test('UID containing a backslash is taken RAW (opaque token), not run through unescapeText', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Invite',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'UID:a\\,b',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(ics);
  assert.equal(events[0].uid, 'a\\,b');
});

test('VEVENT with no UID property parses events[0].uid as null without throwing', () => {
  const ics = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:No UID',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  assert.doesNotThrow(() => {
    const events = parseIcs(ics);
    assert.equal(events[0].uid, null);
  });
});

// --- buildRRuleString -----------------------------------------------------

test('buildRRuleString: WEEKLY with interval, count, byDay', () => {
  const result = buildRRuleString({
    freq: 'WEEKLY',
    interval: 2,
    count: 5,
    until: null,
    byDay: ['MO', 'WE'],
  });
  assert.equal(result, 'RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=5;BYDAY=MO,WE');
});

test('buildRRuleString: DAILY with UNTIL (INTERVAL===1 omitted, UNTIL round-tripped to UTC form)', () => {
  const result = buildRRuleString({
    freq: 'DAILY',
    interval: 1,
    count: null,
    until: new Date('2026-08-10T09:00:00.000Z'),
    byDay: null,
  });
  assert.equal(result, 'RRULE:FREQ=DAILY;UNTIL=20260810T090000Z');
});

test('buildRRuleString: MONTHLY with ordinal-prefixed BYDAY preserved (-1FR, not stripped)', () => {
  const result = buildRRuleString({
    freq: 'MONTHLY',
    interval: 1,
    count: null,
    until: null,
    byDay: ['-1FR'],
  });
  assert.equal(result, 'RRULE:FREQ=MONTHLY;BYDAY=-1FR');
});

test('buildRRuleString: YEARLY with nothing but FREQ', () => {
  const result = buildRRuleString({
    freq: 'YEARLY',
    interval: 1,
    count: null,
    until: null,
    byDay: null,
  });
  assert.equal(result, 'RRULE:FREQ=YEARLY');
});

test('buildRRuleString: YEARLY with BYMONTH emitted when present', () => {
  const result = buildRRuleString({
    freq: 'YEARLY',
    interval: 1,
    count: null,
    until: null,
    byDay: null,
    byMonth: [3],
  });
  assert.equal(result, 'RRULE:FREQ=YEARLY;BYMONTH=3');
});

// --- buildEventResource -----------------------------------------------------

test('buildEventResource: timed event WITH uid -> iCalUID set, dateTime start/end, no recurrence key', () => {
  const event = {
    summary: 'S',
    description: 'D',
    location: 'L',
    start: new Date('2026-08-01T10:00:00Z'),
    end: new Date('2026-08-01T11:00:00Z'),
    isAllDay: false,
    recurrence: null,
    uid: 'u@x',
    sequence: 0,
  };

  const resource = buildEventResource(event);

  assert.deepEqual(resource, {
    summary: 'S',
    description: 'D',
    location: 'L',
    iCalUID: 'u@x',
    start: { dateTime: '2026-08-01T10:00:00.000Z' },
    end: { dateTime: '2026-08-01T11:00:00.000Z' },
    sequence: 0,
  });
  assert.ok(!('recurrence' in resource));
});

test('buildEventResource: timed event WITHOUT uid -> no iCalUID key, dateTime start/end', () => {
  const event = {
    summary: 'S',
    description: 'D',
    location: 'L',
    start: new Date('2026-08-01T10:00:00Z'),
    end: new Date('2026-08-01T11:00:00Z'),
    isAllDay: false,
    recurrence: null,
    uid: null,
    sequence: 0,
  };

  const resource = buildEventResource(event);

  assert.ok(!('iCalUID' in resource));
  assert.deepEqual(resource.start, { dateTime: '2026-08-01T10:00:00.000Z' });
  assert.deepEqual(resource.end, { dateTime: '2026-08-01T11:00:00.000Z' });
});

test('buildEventResource: all-day event -> date-only start/end, no dateTime key', () => {
  const event = {
    summary: 'S',
    description: 'D',
    location: 'L',
    isAllDay: true,
    start: new Date('2026-09-01T00:00:00Z'),
    end: new Date('2026-09-02T00:00:00Z'),
    uid: 'u@x',
    recurrence: null,
    sequence: 0,
  };

  const resource = buildEventResource(event);

  assert.deepEqual(resource.start, { date: '2026-09-01' });
  assert.deepEqual(resource.end, { date: '2026-09-02' });
  assert.ok(!('dateTime' in resource.start));
});

test('buildEventResource: recurring event -> resource.recurrence is an array with the RRULE string', () => {
  const event = {
    summary: 'S',
    description: 'D',
    location: 'L',
    start: new Date('2026-08-03T09:00:00Z'),
    end: new Date('2026-08-03T09:15:00Z'),
    isAllDay: false,
    uid: 'u@x',
    sequence: 0,
    recurrence: {
      freq: 'WEEKLY',
      interval: 2,
      count: 5,
      until: null,
      byDay: ['MO', 'WE'],
    },
  };

  const resource = buildEventResource(event);

  assert.deepEqual(resource.recurrence, ['RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=5;BYDAY=MO,WE']);
});

// --- sequence (RFC 5545 SEQUENCE, live-reported bug quick-260731-seq) -------
//
// See test/ics-parser.test.js's own "SEQUENCE" section for the full
// root-cause writeup: Calendar.Events.import previously omitted sequence
// entirely (implicitly 0), which Google rejected as stale/out-of-order
// whenever Gmail's own native detection had already stored a higher real
// sequence number from the same invite.

test('buildEventResource: a real (non-zero) event.sequence is set verbatim on the resource', () => {
  const event = {
    summary: 'Jan Novák and Eva Dvořáková',
    description: '',
    location: 'https://calendly.com/events/example-meeting-id/microsoft_teams',
    start: new Date('2026-08-03T15:00:00Z'),
    end: new Date('2026-08-03T15:30:00Z'),
    isAllDay: false,
    recurrence: null,
    uid: '040000008200E00074C5B7101A82E008000000001AB2345CD678EF01000000000000000100000004D20E8BCE179574188BC5D2FE29825B2',
    sequence: 1,
  };

  const resource = buildEventResource(event);

  assert.equal(resource.sequence, 1);
});

test('buildEventResource: event.sequence missing/undefined (e.g. a hand-built event not produced by parseVeventBlock) defensively defaults to 0', () => {
  const event = {
    summary: 'S',
    description: 'D',
    location: 'L',
    start: new Date('2026-08-01T10:00:00Z'),
    end: new Date('2026-08-01T11:00:00Z'),
    isAllDay: false,
    recurrence: null,
    uid: 'u@x',
    // sequence deliberately omitted
  };

  const resource = buildEventResource(event);

  assert.equal(resource.sequence, 0);
});

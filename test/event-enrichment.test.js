'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  parseIcs,
  buildEventResource,
  parseAddressProperty,
  formatAddress,
  buildOrganizerAttendeesText,
  collapseBlankLines,
} = require('../src/05-action-ics-import.js');

// ---------------------------------------------------------------------------
// Nesting regression (TR-1, TR-6): a nested VALARM's own DESCRIPTION must not
// clobber the VEVENT's real DESCRIPTION.
// ---------------------------------------------------------------------------

test('VEVENT own DESCRIPTION survives a nested VALARM DESCRIPTION:REMINDER (TR-1)', () => {
  const icsWithNestedAlarm = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Real meeting',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'DESCRIPTION:Real meeting details and Teams link',
    'BEGIN:VALARM',
    'DESCRIPTION:REMINDER',
    'TRIGGER;RELATED=START:-PT15M',
    'ACTION:DISPLAY',
    'END:VALARM',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(icsWithNestedAlarm);
  assert.equal(events.length, 1);
  assert.equal(events[0].description, 'Real meeting details and Teams link');
  assert.notEqual(events[0].description, 'REMINDER');
});

// ---------------------------------------------------------------------------
// parseAddressProperty (TR-2)
// ---------------------------------------------------------------------------

test('parseAddressProperty: CN param + mailto value splits into name/email', () => {
  const parsed = {
    name: 'ORGANIZER',
    params: { CN: 'someone' },
    value: 'mailto:someone@example.com',
  };
  assert.deepEqual(parseAddressProperty(parsed), {
    name: 'someone',
    email: 'someone@example.com',
  });
});

test('parseAddressProperty: no CN param yields null name', () => {
  const parsed = { name: 'ATTENDEE', params: {}, value: 'mailto:noname@example.com' };
  assert.deepEqual(parseAddressProperty(parsed), {
    name: null,
    email: 'noname@example.com',
  });
});

test('parseAddressProperty: case-insensitive mailto: prefix stripped', () => {
  const parsed = { name: 'ORGANIZER', params: {}, value: 'MAILTO:jana@x.com' };
  assert.deepEqual(parseAddressProperty(parsed), {
    name: null,
    email: 'jana@x.com',
  });
});

test('parseAddressProperty: value with no mailto: prefix used as-is', () => {
  const parsed = { name: 'ORGANIZER', params: {}, value: 'jana@x.com' };
  assert.deepEqual(parseAddressProperty(parsed), {
    name: null,
    email: 'jana@x.com',
  });
});

test('parseAddressProperty: null/undefined input returns null', () => {
  assert.equal(parseAddressProperty(null), null);
  assert.equal(parseAddressProperty(undefined), null);
});

// ---------------------------------------------------------------------------
// formatAddress (TR-3)
// ---------------------------------------------------------------------------

test('formatAddress: name differs from email -> "Name <email>"', () => {
  assert.equal(
    formatAddress({ name: 'Jana Nováková', email: 'jana@example.com' }),
    'Jana Nováková <jana@example.com>'
  );
});

test('formatAddress: name case-insensitively equals email -> bare email only', () => {
  assert.equal(formatAddress({ name: 'Jana@EXAMPLE.com', email: 'jana@example.com' }), 'jana@example.com');
});

test('formatAddress: name is null -> bare email only', () => {
  assert.equal(formatAddress({ name: null, email: 'jana@example.com' }), 'jana@example.com');
});

// ---------------------------------------------------------------------------
// buildOrganizerAttendeesText (TR-4)
// ---------------------------------------------------------------------------

test('buildOrganizerAttendeesText: organizer only -> single Organizer line', () => {
  const organizer = { name: 'Jana Nováková', email: 'jana@example.com' };
  assert.equal(
    buildOrganizerAttendeesText(organizer, []),
    'Organizer: Jana Nováková <jana@example.com>'
  );
});

test('buildOrganizerAttendeesText: attendees only, single -> single Attendees line', () => {
  const attendees = [{ name: 'Bob Smith', email: 'bob@example.com' }];
  assert.equal(
    buildOrganizerAttendeesText(null, attendees),
    'Attendees: Bob Smith <bob@example.com>'
  );
});

test('buildOrganizerAttendeesText: attendees only, multiple -> comma-joined', () => {
  const attendees = [
    { name: 'Bob Smith', email: 'bob@example.com' },
    { name: null, email: 'alice@example.com' },
  ];
  assert.equal(
    buildOrganizerAttendeesText(null, attendees),
    'Attendees: Bob Smith <bob@example.com>, alice@example.com'
  );
});

test('buildOrganizerAttendeesText: both -> two lines joined by one newline', () => {
  const organizer = { name: 'Jana Nováková', email: 'jana@example.com' };
  const attendees = [{ name: 'Bob Smith', email: 'bob@example.com' }];
  assert.equal(
    buildOrganizerAttendeesText(organizer, attendees),
    'Organizer: Jana Nováková <jana@example.com>\nAttendees: Bob Smith <bob@example.com>'
  );
});

test('buildOrganizerAttendeesText: neither organizer nor attendees -> exact empty string (no-op)', () => {
  assert.equal(buildOrganizerAttendeesText(null, []), '');
});

// ---------------------------------------------------------------------------
// collapseBlankLines
// ---------------------------------------------------------------------------

test('collapseBlankLines: 3 consecutive newlines collapse to 2', () => {
  assert.equal(collapseBlankLines('before\n\n\nafter'), 'before\n\nafter');
});

test('collapseBlankLines: 5 consecutive newlines collapse to 2', () => {
  assert.equal(collapseBlankLines('before\n\n\n\n\nafter'), 'before\n\nafter');
});

test('collapseBlankLines: a single blank line (\\n\\n) is left unchanged', () => {
  assert.equal(collapseBlankLines('before\n\nafter'), 'before\n\nafter');
});

test('collapseBlankLines: no blank lines is left unchanged', () => {
  assert.equal(collapseBlankLines('one line\nanother line'), 'one line\nanother line');
});

// ---------------------------------------------------------------------------
// Integration: description enrichment (TR-5)
// ---------------------------------------------------------------------------

const ORGANIZER_ATTENDEE_ICS = [
  'BEGIN:VCALENDAR',
  'BEGIN:VEVENT',
  'SUMMARY:Sprint planning',
  'DTSTART:20260801T100000Z',
  'DTEND:20260801T110000Z',
  'DESCRIPTION:Sprint planning agenda and notes',
  'ORGANIZER;CN=Jana Nováková:mailto:jana@example.com',
  'ATTENDEE;CN=Bob Smith:mailto:bob@example.com',
  'ATTENDEE:mailto:alice@example.com',
  'LOCATION:Room 4B',
  'X-MICROSOFT-SKYPETEAMSMEETINGURL:https://teams.microsoft.com/l/meetup-join/abc%3D',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

test('integration: description has Organizer/Attendees block PREPENDED before the original text (TR-5)', () => {
  const events = parseIcs(ORGANIZER_ATTENDEE_ICS);
  const event = events[0];

  assert.equal(
    event.description,
    'Organizer: Jana Nováková <jana@example.com>\n' +
      'Attendees: Bob Smith <bob@example.com>, alice@example.com' +
      '\n\n' +
      'Sprint planning agenda and notes'
  );
});

test('integration: description with organizer/attendees but NO original description text is just the block alone (no trailing separator)', () => {
  const icsNoDescription = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Sprint planning',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'ORGANIZER;CN=Jana Nováková:mailto:jana@example.com',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(icsNoDescription);
  assert.equal(events[0].description, 'Organizer: Jana Nováková <jana@example.com>');
});

test('integration: excess blank lines (3+) in the raw DESCRIPTION are collapsed to a single blank line in the final description', () => {
  const icsWithExcessBlankLines = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Interview',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'DESCRIPTION:Interview Confirmation\\n\\n\\n\\nHi Jan\\n\\n\\n\\n\\n\\n\\n\\nSee you then',
    'ORGANIZER;CN=Jana Nováková:mailto:jana@example.com',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(icsWithExcessBlankLines);
  assert.equal(
    events[0].description,
    'Organizer: Jana Nováková <jana@example.com>\n\n' +
      'Interview Confirmation\n\nHi Jan\n\nSee you then'
  );
});

// ---------------------------------------------------------------------------
// Integration: location enrichment (TR-5)
// ---------------------------------------------------------------------------

test('integration: meeting URL REPLACES the original LOCATION entirely when present (TR-5)', () => {
  const events = parseIcs(ORGANIZER_ATTENDEE_ICS);
  const event = events[0];

  assert.equal(event.location, 'https://teams.microsoft.com/l/meetup-join/abc%3D');
});

test('integration: no meeting URL -> location falls back to the original LOCATION value unchanged', () => {
  const icsNoMeetingUrl = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Sprint planning',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'LOCATION:Room 4B',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(icsNoMeetingUrl);
  assert.equal(events[0].location, 'Room 4B');
});

test('integration: meeting URL present with NO original LOCATION -> location is just the URL', () => {
  const icsMeetingUrlNoLocation = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'SUMMARY:Sprint planning',
    'DTSTART:20260801T100000Z',
    'DTEND:20260801T110000Z',
    'X-MICROSOFT-SKYPETEAMSMEETINGURL:https://teams.microsoft.com/l/meetup-join/abc%3D',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcs(icsMeetingUrlNoLocation);
  assert.equal(events[0].location, 'https://teams.microsoft.com/l/meetup-join/abc%3D');
});

// ---------------------------------------------------------------------------
// Safety assertion (T-03-05): buildEventResource must never carry
// attendees/organizer resource fields, no matter what the parsed event holds.
// ---------------------------------------------------------------------------

test('safety: buildEventResource has no attendees/organizer keys for an organizer+attendee event (T-03-05)', () => {
  const events = parseIcs(ORGANIZER_ATTENDEE_ICS);
  const resource = buildEventResource(events[0]);

  assert.equal(Object.prototype.hasOwnProperty.call(resource, 'attendees'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(resource, 'organizer'), false);
});

// ---------------------------------------------------------------------------
// No-op control (TR-6, TR-9): an event with neither organizer/attendees nor
// the meeting-URL property parses byte-identically to pre-change behavior.
// ---------------------------------------------------------------------------

test('no-op control: event without organizer/attendees/meeting-URL is unchanged (TR-9)', () => {
  const plainIcs = [
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

  const events = parseIcs(plainIcs);
  assert.equal(events.length, 1);
  assert.equal(events[0].description, 'Quarterly planning session');
  assert.equal(events[0].location, 'Conference Room A');
});

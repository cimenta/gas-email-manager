'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  resolveTransportSender,
  resolveTransportCalendarId,
  isTransportPdfAttachment,
  findTransportTicketPdfAttachment,
  extractTransportTicketIdentifier,
  buildTransportAttachmentFilename,
  resolveTransportProcessingJobs,
  TRANSPORT_TICKETS_ACTION,
} = require('../src/08-action-transport-tickets.js');
const { parseIcs, buildEventResource } = require('../src/05-action-ics-import.js');
const { TRANSPORT_TICKETS_ACTION_CONFIG } = require('../src/08-action-cfg-transport-tickets.js');

// --- REAL_REGIOJET_ICS -------------------------------------------------------
//
// quick-260803-us3: a real RegioJet ("jizdenky@regiojet.cz") train-ticket
// confirmation email carries its own text/calendar .ics attachment -- this
// action reuses the EXISTING parseIcs/buildEventResource parser
// (src/05-action-ics-import.js) rather than hand-rolling a second one (D-01).
//
// The VEVENT block below is MODELED on a real RegioJet .eml's VEVENT
// structure and TZID/VTIMEZONE shape (DTSTAMP, TZID-qualified DTSTART/DTEND,
// SUMMARY/DESCRIPTION/LOCATION/UID/SEQUENCE fields, seat-list formatting),
// with every trip-identifying and personal detail replaced by fictional
// equivalents (route, ticket number, seats, coordinates, UID, ATTENDEE
// identity) -- no field is copied verbatim from any real email.
//
// The two VTIMEZONE blocks (Europe/Vienna, Europe/Prague) are the STANDARD
// EU CET/CEST transition rules (STANDARD: TZOFFSETFROM +0200 -> TZOFFSETTO
// +0100, last Sunday of October; DAYLIGHT: TZOFFSETFROM +0100 -> TZOFFSETTO
// +0200, last Sunday of March) -- the VEVENT's DTSTART/DTEND TZID references
// are what actually matters for this test, and those only need to resolve
// against CORRECT EU rules, not a byte-for-byte copy of any source file.
const REAL_REGIOJET_ICS = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'METHOD:REQUEST',
  'BEGIN:VTIMEZONE',
  'TZID:Europe/Vienna',
  'BEGIN:STANDARD',
  'DTSTART:19701025T030000',
  'TZOFFSETFROM:+0200',
  'TZOFFSETTO:+0100',
  'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
  'END:STANDARD',
  'BEGIN:DAYLIGHT',
  'DTSTART:19700329T020000',
  'TZOFFSETFROM:+0100',
  'TZOFFSETTO:+0200',
  'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
  'END:DAYLIGHT',
  'END:VTIMEZONE',
  'BEGIN:VTIMEZONE',
  'TZID:Europe/Prague',
  'BEGIN:STANDARD',
  'DTSTART:19701025T030000',
  'TZOFFSETFROM:+0200',
  'TZOFFSETTO:+0100',
  'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
  'END:STANDARD',
  'BEGIN:DAYLIGHT',
  'DTSTART:19700329T020000',
  'TZOFFSETFROM:+0100',
  'TZOFFSETTO:+0200',
  'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
  'END:DAYLIGHT',
  'END:VTIMEZONE',
  'BEGIN:VEVENT',
  'DTSTAMP:20260802T174558Z',
  'DTSTART;TZID=Europe/Vienna:20260818T173600',
  'DTEND;TZID=Europe/Prague:20260818T191200',
  'CREATED:20260802T174558Z',
  'ATTENDEE;CN=jan.novak:mailto:jan.novak@example.com',
  'SUMMARY:#7788123456: Z Ostrava, hl.n., do Praha, hl.n., sedadla: [2/15,2/16]',
  'DESCRIPTION:Příjezd/Odjezd|Zastávka/Přestup|Nást.|Spoj|Vůz/sedadla\\nOdj:17:36|Ostrava, hl.n.|Ostrava → Praha (RJ, RJ 1036)|2/15, 16\\nPří:19:12|Praha, hl.n.\\n',
  'LOCATION:49.8400000, 18.2900000',
  'UID:-9876543210@regiojet.cz',
  'SEQUENCE:1',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

const REAL_SUMMARY = '#7788123456: Z Ostrava, hl.n., do Praha, hl.n., sedadla: [2/15,2/16]';

// --- duck-typed fake attachment/message factories ---------------------------
// Same style as test/ticketing-portals.test.js's fakeAttachment/fakeMessage.

function fakeAttachment(name, contentType, dataAsString) {
  return {
    getName: function () {
      return name;
    },
    getContentType: function () {
      return contentType || '';
    },
    getDataAsString: function () {
      return dataAsString || '';
    },
  };
}

function fakeMessage(fromHeader, attachments) {
  return {
    getFrom: function () {
      return fromHeader;
    },
    getAttachments: function () {
      return attachments || [];
    },
  };
}

const REAL_SENDERS = [{ identifyingEmail: 'jizdenky@regiojet.cz', calendarId: null, insertPdfIntoEvent: false }];

// --- resolveTransportSender --------------------------------------------------

test('resolveTransportSender: the real bare "jizdenky@regiojet.cz" From header resolves to the configured entry', () => {
  assert.equal(resolveTransportSender('jizdenky@regiojet.cz', REAL_SENDERS), REAL_SENDERS[0]);
});

test('resolveTransportSender: a "Name <jizdenky@regiojet.cz>" header resolves', () => {
  assert.equal(resolveTransportSender('RegioJet <jizdenky@regiojet.cz>', REAL_SENDERS), REAL_SENDERS[0]);
});

test('resolveTransportSender: a mixed-case header resolves (case-insensitive)', () => {
  assert.equal(resolveTransportSender('Jizdenky@RegioJet.CZ', REAL_SENDERS), REAL_SENDERS[0]);
});

test('resolveTransportSender: an unknown sender returns null', () => {
  assert.equal(resolveTransportSender('someone-else@example.com', REAL_SENDERS), null);
});

test('resolveTransportSender: a null senders list returns null without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(resolveTransportSender('jizdenky@regiojet.cz', null), null);
  });
});

test('resolveTransportSender: an empty senders list returns null without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(resolveTransportSender('jizdenky@regiojet.cz', []), null);
  });
});

// --- resolveTransportCalendarId ----------------------------------------------

test('resolveTransportCalendarId: entry calendarId wins when truthy', () => {
  assert.equal(resolveTransportCalendarId({ calendarId: 'CAL_REAL' }, 'CAL_DEFAULT'), 'CAL_REAL');
});

test('resolveTransportCalendarId: falls back to the passed default when calendarId is null (the shipped default) -- reproduces the real live crash class from quick-260731-tix round 6', () => {
  assert.equal(resolveTransportCalendarId({ calendarId: null }, 'CAL_DEFAULT'), 'CAL_DEFAULT');
});

// --- isTransportPdfAttachment / findTransportTicketPdfAttachment ------------

const REAL_ICS_ATTACHMENT = fakeAttachment('ticket.ics', 'text/calendar', REAL_REGIOJET_ICS);
const REAL_ETICKET_ATTACHMENT = fakeAttachment('eticket.pdf', 'application/pdf');
const REAL_INVOICE_ATTACHMENT = fakeAttachment('invoice.pdf', 'application/pdf');

test('isTransportPdfAttachment: matches by .pdf name extension', () => {
  assert.equal(isTransportPdfAttachment(REAL_ETICKET_ATTACHMENT), true);
});

test('isTransportPdfAttachment: a non-PDF attachment does not match', () => {
  assert.equal(isTransportPdfAttachment(REAL_ICS_ATTACHMENT), false);
});

test('findTransportTicketPdfAttachment: given the real 3-attachment message, returns eticket.pdf (not invoice.pdf, not the .ics)', () => {
  const message = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT, REAL_ETICKET_ATTACHMENT, REAL_INVOICE_ATTACHMENT]);
  assert.equal(findTransportTicketPdfAttachment(message), REAL_ETICKET_ATTACHMENT);
});

test('findTransportTicketPdfAttachment: returns null when the only PDF is invoice.pdf', () => {
  const message = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT, REAL_INVOICE_ATTACHMENT]);
  assert.equal(findTransportTicketPdfAttachment(message), null);
});

test('findTransportTicketPdfAttachment: returns null when there are no PDFs at all', () => {
  const message = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT]);
  assert.equal(findTransportTicketPdfAttachment(message), null);
});

test('findTransportTicketPdfAttachment: the invoice exclusion is case-insensitive on the filename', () => {
  const shoutingInvoice = fakeAttachment('INVOICE.PDF', 'application/pdf');
  const message = fakeMessage('jizdenky@regiojet.cz', [shoutingInvoice]);
  assert.equal(findTransportTicketPdfAttachment(message), null);
});

// --- extractTransportTicketIdentifier ----------------------------------------

test('extractTransportTicketIdentifier: the real SUMMARY yields "7788123456"', () => {
  assert.equal(extractTransportTicketIdentifier({ summary: REAL_SUMMARY, uid: '-9876543210@regiojet.cz' }), '7788123456');
});

test('extractTransportTicketIdentifier: a summary with no #<digits> prefix but a UID yields the raw UID', () => {
  assert.equal(extractTransportTicketIdentifier({ summary: 'No hash prefix here', uid: '-9876543210@regiojet.cz' }), '-9876543210@regiojet.cz');
});

test('extractTransportTicketIdentifier: neither present yields null (never throws)', () => {
  assert.doesNotThrow(() => {
    assert.equal(extractTransportTicketIdentifier({ summary: 'No hash prefix here', uid: null }), null);
  });
});

// --- buildTransportAttachmentFilename ----------------------------------------

test('buildTransportAttachmentFilename: the real summary, real start date, and real ticket identifier produce the exact expected filename', () => {
  const startDate = new Date(Date.UTC(2026, 7, 18, 15, 36));
  assert.equal(
    buildTransportAttachmentFilename(REAL_SUMMARY, startDate, '7788123456'),
    '#7788123456- Z Ostrava, hl.n., do Praha, hl.n., sedadla- [2-15,2-16] - 2026-08-18 - 7788123456.pdf'
  );
});

test('buildTransportAttachmentFilename: a null identifier omits that segment entirely and never embeds the literal word "null"', () => {
  const startDate = new Date(Date.UTC(2026, 7, 18, 15, 36));
  const filename = buildTransportAttachmentFilename(REAL_SUMMARY, startDate, null);
  assert.equal(filename, '#7788123456- Z Ostrava, hl.n., do Praha, hl.n., sedadla- [2-15,2-16] - 2026-08-18.pdf');
  assert.equal(filename.indexOf('null'), -1);
});

// --- resolveTransportProcessingJobs ------------------------------------------

test('resolveTransportProcessingJobs: a matching message carrying ticket.ics yields exactly ONE job with only the .ics attachment collected -- never one job per attachment', () => {
  const message = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT, REAL_ETICKET_ATTACHMENT, REAL_INVOICE_ATTACHMENT]);
  const jobs = resolveTransportProcessingJobs([message], REAL_SENDERS);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].sender, REAL_SENDERS[0]);
  assert.deepEqual(jobs[0].icsAttachments, [REAL_ICS_ATTACHMENT]);
});

test('resolveTransportProcessingJobs: a matching sender with NO .ics attachment yields zero jobs', () => {
  const message = fakeMessage('jizdenky@regiojet.cz', [REAL_ETICKET_ATTACHMENT, REAL_INVOICE_ATTACHMENT]);
  assert.deepEqual(resolveTransportProcessingJobs([message], REAL_SENDERS), []);
});

test('resolveTransportProcessingJobs: a non-matching sender yields zero jobs', () => {
  const message = fakeMessage('someone-else@example.com', [REAL_ICS_ATTACHMENT]);
  assert.deepEqual(resolveTransportProcessingJobs([message], REAL_SENDERS), []);
});

test('resolveTransportProcessingJobs: two matching messages on one thread yield two jobs', () => {
  const message1 = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT]);
  const message2 = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT]);
  const jobs = resolveTransportProcessingJobs([message1, message2], REAL_SENDERS);
  assert.equal(jobs.length, 2);
});

// --- Parser-reuse proof (D-01): the existing parseIcs/buildEventResource ----
// --- parser is reused verbatim, no second VEVENT parser exists --------------

test('parseIcs(REAL_REGIOJET_ICS): resolves the TZID-qualified DTSTART/DTEND through the real VTIMEZONE blocks, and extracts the rest of the VEVENT correctly', () => {
  const events = parseIcs(REAL_REGIOJET_ICS);
  assert.equal(events.length, 1);

  const event = events[0];
  assert.equal(event.start.toISOString(), '2026-08-18T15:36:00.000Z');
  assert.equal(event.end.toISOString(), '2026-08-18T17:12:00.000Z');
  assert.equal(event.sequence, 1);
  assert.equal(event.uid, '-9876543210@regiojet.cz');
  assert.equal(event.summary, REAL_SUMMARY);
  assert.equal(event.isAllDay, false);
  assert.ok(event.description.indexOf('Attendees: ') === 0, 'description should start with the Attendees: line');
  assert.ok(event.description.indexOf('Odj:17:36') !== -1, 'description should contain the real DESCRIPTION text');
});

test('buildEventResource on the parsed RegioJet event produces iCalUID, sequence, and matching start/end instants', () => {
  const event = parseIcs(REAL_REGIOJET_ICS)[0];
  const resource = buildEventResource(event);

  assert.equal(resource.iCalUID, '-9876543210@regiojet.cz');
  assert.equal(resource.sequence, 1);
  assert.equal(resource.start.dateTime, '2026-08-18T15:36:00.000Z');
  assert.equal(resource.end.dateTime, '2026-08-18T17:12:00.000Z');
});

// --- TRANSPORT_TICKETS_ACTION -------------------------------------------------

test('TRANSPORT_TICKETS_ACTION: name is "transport-tickets"', () => {
  assert.equal(TRANSPORT_TICKETS_ACTION.name, 'transport-tickets');
});

test('TRANSPORT_TICKETS_ACTION: .config is the SAME object reference as TRANSPORT_TICKETS_ACTION_CONFIG (getter identity, not a copy)', () => {
  assert.equal(TRANSPORT_TICKETS_ACTION.config, TRANSPORT_TICKETS_ACTION_CONFIG);
});

test('TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders: code default is the single jizdenky@regiojet.cz entry with calendarId null and insertPdfIntoEvent false', () => {
  assert.deepEqual(TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders, [
    { identifyingEmail: 'jizdenky@regiojet.cz', calendarId: null, insertPdfIntoEvent: false },
  ]);
});

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
  parseIdosTicketText,
  resolveTransportSenderMode,
  buildTransportBodyEntry,
  buildTransportIcsEntry,
  stripTransportSummaryIdentifierPrefix,
  partitionTransportEntriesByCancellation,
  cancelTransportTicketEvent,
  buildTransportEventPrivateProperties,
  isTransportCancellationStale,
  filterTransportEntriesToCreate,
  TRANSPORT_TICKETS_ACTION,
} = require('../src/08-action-transport-tickets.js');
const { parseIcs, buildEventResource } = require('../src/05-action-ics-import.js');
const { TRANSPORT_TICKETS_ACTION_CONFIG } = require('../src/08-action-cfg-transport-tickets.js');
const { isValidTicketingPortalsShape } = require('../src/01-setup.js');

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

// --- REAL_REGIOJET_CANCEL_ICS (D-06, quick-260813-dq2) -----------------------
//
// The matching CANCEL VEVENT for REAL_REGIOJET_ICS above -- same fictional
// ticket number (7788123456) and same fictional UID
// (-9876543210@regiojet.cz), so the pair reads as a natural
// confirm-then-cancel sequence for ONE ticket, exactly mirroring what a real
// RegioJet cancellation email carries: METHOD:CANCEL at VCALENDAR level,
// STATUS:CANCELLED on the VEVENT, a bumped SEQUENCE, and a later DTSTAMP
// (the cancellation was generated after the confirmation). Everything else
// (both VTIMEZONE blocks, DTSTART/DTEND, SUMMARY, DESCRIPTION, LOCATION,
// ATTENDEE, CREATED) is identical to REAL_REGIOJET_ICS -- written out as a
// full explicit array literal (not derived by string surgery) so the lines
// under test stay visible, per this codebase's fixture convention.
const REAL_REGIOJET_CANCEL_ICS = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'METHOD:CANCEL',
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
  'DTSTAMP:20260803T090000Z',
  'DTSTART;TZID=Europe/Vienna:20260818T173600',
  'DTEND;TZID=Europe/Prague:20260818T191200',
  'CREATED:20260802T174558Z',
  'ATTENDEE;CN=jan.novak:mailto:jan.novak@example.com',
  'SUMMARY:#7788123456: Z Ostrava, hl.n., do Praha, hl.n., sedadla: [2/15,2/16]',
  'DESCRIPTION:Příjezd/Odjezd|Zastávka/Přestup|Nást.|Spoj|Vůz/sedadla\\nOdj:17:36|Ostrava, hl.n.|Ostrava → Praha (RJ, RJ 1036)|2/15, 16\\nPří:19:12|Praha, hl.n.\\n',
  'LOCATION:49.8400000, 18.2900000',
  'UID:-9876543210@regiojet.cz',
  'STATUS:CANCELLED',
  'SEQUENCE:2',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

// --- REAL_REGIOJET_REBOOK_ICS (D-06, quick-260813-dq2 Task 3) ---------------
//
// The REBOOKING VEVENT for the SAME fictional ticket (7788123456) and SAME
// fictional UID (-9876543210@regiojet.cz) as REAL_REGIOJET_ICS and
// REAL_REGIOJET_CANCEL_ICS above -- the three fixtures read as one natural
// confirm -> cancel -> rebook sequence for ONE ticket, mirroring the real
// owner-reported emails (Problem A/B, D-08..D-12): DTSTAMP five minutes
// AFTER the CANCEL fixture's own DTSTAMP (the real emails' 5-minute gap),
// SEQUENCE RESET to 1 (RegioJet does not continue to 3 -- this is the whole
// reason SEQUENCE cannot be used for staleness detection, D-11/D-12),
// different fictional seats so "the rebooked event's details" is an
// observable difference, and NO STATUS line at all (a rebooking is an
// ordinary booking). Everything else (both VTIMEZONE blocks, DTSTART/DTEND,
// LOCATION, ATTENDEE, CREATED) stays identical -- written as a full explicit
// array literal, never derived from another fixture by string surgery.
const REAL_REGIOJET_REBOOK_ICS = [
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
  'DTSTAMP:20260803T090500Z',
  'DTSTART;TZID=Europe/Vienna:20260818T173600',
  'DTEND;TZID=Europe/Prague:20260818T191200',
  'CREATED:20260802T174558Z',
  'ATTENDEE;CN=jan.novak:mailto:jan.novak@example.com',
  'SUMMARY:#7788123456: Z Ostrava, hl.n., do Praha, hl.n., sedadla: [4/21,4/22]',
  'DESCRIPTION:Příjezd/Odjezd|Zastávka/Přestup|Nást.|Spoj|Vůz/sedadla\\nOdj:17:36|Ostrava, hl.n.|Ostrava → Praha (RJ, RJ 1036)|4/21, 22\\nPří:19:12|Praha, hl.n.\\n',
  'LOCATION:49.8400000, 18.2900000',
  'UID:-9876543210@regiojet.cz',
  'SEQUENCE:1',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

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

function fakeMessage(fromHeader, attachments, plainBody) {
  return {
    getFrom: function () {
      return fromHeader;
    },
    getAttachments: function () {
      return attachments || [];
    },
    getPlainBody: function () {
      return plainBody || '';
    },
  };
}

const REAL_SENDERS = [{ identifyingEmail: 'jizdenky@regiojet.cz', calendarId: null, insertPdfIntoEvent: false }];

// --- REAL_IDOS_BODY_TEXT -----------------------------------------------------
//
// quick-260804-bs7: an IDOS.cz ("jizdenky@idos.svt.cz") confirmation email
// carries NO .ics attachment at all -- route, dates, times, seats and both
// codes live only in the plain-text email body (message.getPlainBody()),
// next to a single ticket PDF. This is the SECOND transport carrier, and the
// FIRST one processed via the new 'body' mode (D-01) -- RegioJet's existing
// 'ics'-mode flow above is completely unaffected.
//
// The STRUCTURE below (paragraph order, Czech wording, the "- " bullets on
// the five e-ticket detail lines, the "D.M.YYYY H:MM" no-leading-zero
// date/time format, the "»" route separator, the "sedadlo A/B C/D" seat
// shape) is modelled on a real IDOS.cz confirmation email documented in
// 260804-bs7-CONTEXT.md, while EVERY trip-specific and personal value below
// is FICTIONAL (D-04, per the 2026-08-03 privacy-audit precedent -- see
// STATE.md's dedicated entry for the real incident that forced a full
// public-repo history squash): the route, both departure/arrival
// date/times, the seats, the e-ticket code, the IDOS.cz order code, the
// order number, the price, and -- critically -- the support-URL's `email=`
// query param, which in the real message carries the owner's own personal
// mail address. The sender identifiers (jizdenky@idos.svt.cz / idos.svt.cz
// / "IDOS.cz") are real business identifiers and stay exactly as-is, same
// rule already applied to RegioJet, enigoo.cz and Kino Art.
const IDOS_BODY_LINES = [
  'Vážený zákazníku,',
  '',
  'posíláme Vám potvrzení o nákupu e-jízdenky. U e-jízdenky je uvedeno, jestli je pro odbavení nutné ji zobrazit (obrazovka nebo tisk), nebo stačí jen znát její kód pro odbavení, nebo je nutné odbavení pouze v aplikaci IDOS (v případě takové jízdenky je její příloha jen daňový doklad). Detailní informace k požadavkům na odbavení s konkrétní e-jízdenkou jsou uvedeny na e-jízdence nebo ve Smluvních podmínkách. V příloze najdete potřebné PDF e-jízdenky respektive daňový doklad.',
  '',
  'Objednávka obsahuje tuto e-jízdenku:',
  '',
  '- kód e-jízdenky 7QKMR2,',
  '- 4.9.2026 6:05 Zelené Údolí hl.n. » Neustadt Hbf 4.9.2026 17:40, sedadlo 412/31 412/33,',
  '- 2x Včasná jízdenka Evropa, 2. třída, 1 osoba, Povinný vlak rj 51 v úseku Zelené Údolí hl.n. – Neustadt Hbf, 1x Rezervace, 2. třída, 2 osoby,',
  '- kód IDOS.cz PLTQ-DMVR-ZKBN,',
  '- e-jízdenku je nutné zobrazit nebo vytisknout',
  '',
  'Celková cena Vaší objednávky číslo 5120097364 je 689,- Kč. Číslo objednávky a kód IDOS.cz v žádném případě neslouží k odbavení!',
  '',
  'Podmínky pro vrácení e-jízdenky a případnou reklamaci při zrušení nebo opoždění spoje jsou uvedeny ve https://resources.crws.cz/content/conditions/spi.c.html.',
  '',
  'Pro případné dotazy využijte formulář na stránce Zákaznické podpory https://helpdesk.amsbus.cz/IDOS?lang=cs&email=zakaznik%40example.com&trnkod=PLTQ-DMVR-ZKBN&custom=401.',
  '',
  'Děkujeme Vám za využití služeb IDOS.cz.',
];

const REAL_IDOS_BODY_TEXT = IDOS_BODY_LINES.join('\n');
// CRLF variant proving the anchors are separator-agnostic (the round-4/5
// enigoo.cz lesson -- Apps Script's various text-extraction paths do not
// consistently join paragraphs with the same separator character).
const REAL_IDOS_BODY_TEXT_CRLF = IDOS_BODY_LINES.join('\r\n');

// Gmail real list-rendering marker variant (proven live in quick-260731-kar
// round 3): the five e-ticket detail lines (indices 6-10) prefixed with
// "* " instead of "- " -- the anchors must not depend on the bullet
// character.
const IDOS_BODY_LINES_BULLET_STAR = IDOS_BODY_LINES.map(function (line, index) {
  if (index >= 6 && index <= 10) {
    return line.replace(/^- /, '* ');
  }
  return line;
});
const REAL_IDOS_BODY_TEXT_BULLET_STAR = IDOS_BODY_LINES_BULLET_STAR.join('\n');

// Optional-anchor-missing variants (trip line present, one optional field
// absent) -- each must parse WITHOUT throwing (optional-anchors-never-throw
// rule).
const IDOS_BODY_LINES_NO_ORDER_CODE = IDOS_BODY_LINES.filter(function (_line, index) {
  return index !== 9; // '- kód IDOS.cz PLTQ-DMVR-ZKBN,'
});
const IDOS_BODY_LINES_NO_SEATS = IDOS_BODY_LINES.map(function (line, index) {
  if (index === 7) {
    return line.replace(', sedadlo 412/31 412/33', '');
  }
  return line;
});
const IDOS_BODY_LINES_NO_ETICKET_CODE = IDOS_BODY_LINES.filter(function (_line, index) {
  return index !== 6; // '- kód e-jízdenky 7QKMR2,'
});

// A text with no trip line at all (mandatory-anchor-missing -- must throw).
const IDOS_BODY_TEXT_NO_TRIP_LINE = 'Vážený zákazníku, toto je text bez žádné platné jízdenky.';

// Out-of-range hour/minute variants (mandatory validation -- must throw).
const IDOS_BODY_TEXT_HOUR_OUT_OF_RANGE = REAL_IDOS_BODY_TEXT.replace('4.9.2026 6:05', '4.9.2026 25:05');
const IDOS_BODY_TEXT_MINUTE_OUT_OF_RANGE = REAL_IDOS_BODY_TEXT.replace('4.9.2026 17:40', '4.9.2026 17:75');

const REAL_IDOS_SENDERS = [{ identifyingEmail: 'jizdenky@idos.svt.cz', calendarId: null, insertPdfIntoEvent: false, mode: 'body' }];

// The real ticket PDF's real MIME facts (D-03): Content-Type
// application/octet-stream (NOT application/pdf), name "ticket <order
// code>.pdf" -- matched by the EXISTING isTransportPdfAttachment/
// findTransportTicketPdfAttachment pair via the .pdf filename suffix, no
// new finder.
const REAL_IDOS_TICKET_PDF_ATTACHMENT = fakeAttachment('ticket PLTQ-DMVR-ZKBN.pdf', 'application/octet-stream');

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

// --- IDOS.cz ticket PDF attachment (D-03, quick-260804-bs7) -----------------
// Real MIME fact: Content-Type application/octet-stream, NOT application/pdf
// -- matched on the .pdf filename suffix by the EXISTING helpers, no new
// finder function.

test('isTransportPdfAttachment: matches the IDOS.cz ticket PDF despite its application/octet-stream content type', () => {
  assert.equal(isTransportPdfAttachment(REAL_IDOS_TICKET_PDF_ATTACHMENT), true);
});

test('findTransportTicketPdfAttachment: returns the IDOS.cz ticket PDF from a message whose only attachment it is', () => {
  const message = fakeMessage('jizdenky@idos.svt.cz', [REAL_IDOS_TICKET_PDF_ATTACHMENT], REAL_IDOS_BODY_TEXT);
  assert.equal(findTransportTicketPdfAttachment(message), REAL_IDOS_TICKET_PDF_ATTACHMENT);
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

// --- resolveTransportProcessingJobs, body mode (D-01, quick-260804-bs7) -----

test('resolveTransportProcessingJobs: an IDOS.cz message with ONLY the ticket PDF (no .ics at all) yields exactly ONE mode:"body" job', () => {
  const message = fakeMessage('jizdenky@idos.svt.cz', [REAL_IDOS_TICKET_PDF_ATTACHMENT], REAL_IDOS_BODY_TEXT);
  const jobs = resolveTransportProcessingJobs([message], REAL_IDOS_SENDERS);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].sender, REAL_IDOS_SENDERS[0]);
});

test('resolveTransportProcessingJobs: the same IDOS.cz message from an UNCONFIGURED sender yields zero jobs', () => {
  const message = fakeMessage('jizdenky@idos.svt.cz', [REAL_IDOS_TICKET_PDF_ATTACHMENT], REAL_IDOS_BODY_TEXT);
  assert.deepEqual(resolveTransportProcessingJobs([message], REAL_SENDERS), []);
});

test('resolveTransportProcessingJobs: a RegioJet message still yields exactly one mode:"ics" job carrying only the .ics attachment (regression guard)', () => {
  const message = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT, REAL_ETICKET_ATTACHMENT, REAL_INVOICE_ATTACHMENT]);
  const jobs = resolveTransportProcessingJobs([message], REAL_SENDERS);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'ics');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].sender, REAL_SENDERS[0]);
  assert.deepEqual(jobs[0].icsAttachments, [REAL_ICS_ATTACHMENT]);
});

test('resolveTransportProcessingJobs: a thread with one RegioJet message AND one IDOS.cz message yields two jobs, one of each mode, in message order', () => {
  const regioJetMessage = fakeMessage('jizdenky@regiojet.cz', [REAL_ICS_ATTACHMENT]);
  const idosMessage = fakeMessage('jizdenky@idos.svt.cz', [REAL_IDOS_TICKET_PDF_ATTACHMENT], REAL_IDOS_BODY_TEXT);
  const combinedSenders = REAL_SENDERS.concat(REAL_IDOS_SENDERS);
  const jobs = resolveTransportProcessingJobs([regioJetMessage, idosMessage], combinedSenders);

  assert.equal(jobs.length, 2);
  assert.equal(jobs[0].mode, 'ics');
  assert.equal(jobs[0].message, regioJetMessage);
  assert.equal(jobs[1].mode, 'body');
  assert.equal(jobs[1].message, idosMessage);
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

// --- parseIdosTicketText (D-01/D-02, quick-260804-bs7) -----------------------

test('parseIdosTicketText: the fictional fixture yields the exact route, both wall-clock components, seats, e-ticket code and order code (D-02)', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);

  assert.equal(parsed.from, 'Zelené Údolí hl.n.');
  assert.equal(parsed.to, 'Neustadt Hbf');
  assert.equal(parsed.from.indexOf('»'), -1);
  assert.equal(parsed.to.indexOf('»'), -1);
  assert.deepEqual(parsed.start, { year: 2026, month: 8, day: 4, hour: 6, minute: 5 });
  assert.deepEqual(parsed.end, { year: 2026, month: 8, day: 4, hour: 17, minute: 40 });
  assert.equal(parsed.seats, '412/31 412/33');
  assert.equal(parsed.eTicketCode, '7QKMR2');
  assert.equal(parsed.ticketIdentifier, 'PLTQ-DMVR-ZKBN');
  assert.notEqual(parsed.ticketIdentifier, parsed.eTicketCode);
});

test('parseIdosTicketText: the CRLF fixture parses identically to the LF one', () => {
  assert.deepEqual(parseIdosTicketText(REAL_IDOS_BODY_TEXT_CRLF), parseIdosTicketText(REAL_IDOS_BODY_TEXT));
});

test('parseIdosTicketText: a "* " Gmail bullet marker (instead of "- ") on the detail lines parses identically', () => {
  assert.deepEqual(parseIdosTicketText(REAL_IDOS_BODY_TEXT_BULLET_STAR), parseIdosTicketText(REAL_IDOS_BODY_TEXT));
});

test('parseIdosTicketText: a text with no trip line throws, and the thrown message contains the full raw text', () => {
  assert.throws(
    () => {
      parseIdosTicketText(IDOS_BODY_TEXT_NO_TRIP_LINE);
    },
    function (err) {
      return err instanceof Error && err.message.indexOf(IDOS_BODY_TEXT_NO_TRIP_LINE) !== -1;
    }
  );
});

test('parseIdosTicketText: trip line present but "kód IDOS.cz" absent returns ticketIdentifier null WITHOUT throwing (optional-anchor rule)', () => {
  assert.doesNotThrow(() => {
    const parsed = parseIdosTicketText(IDOS_BODY_LINES_NO_ORDER_CODE.join('\n'));
    assert.equal(parsed.ticketIdentifier, null);
  });
});

test('parseIdosTicketText: seats absent returns seats null WITHOUT throwing', () => {
  assert.doesNotThrow(() => {
    const parsed = parseIdosTicketText(IDOS_BODY_LINES_NO_SEATS.join('\n'));
    assert.equal(parsed.seats, null);
  });
});

test('parseIdosTicketText: e-ticket code absent returns eTicketCode null WITHOUT throwing', () => {
  assert.doesNotThrow(() => {
    const parsed = parseIdosTicketText(IDOS_BODY_LINES_NO_ETICKET_CODE.join('\n'));
    assert.equal(parsed.eTicketCode, null);
  });
});

test('parseIdosTicketText: an out-of-range departure hour throws, with the full text appended', () => {
  assert.throws(
    () => {
      parseIdosTicketText(IDOS_BODY_TEXT_HOUR_OUT_OF_RANGE);
    },
    function (err) {
      return err instanceof Error && err.message.indexOf(IDOS_BODY_TEXT_HOUR_OUT_OF_RANGE) !== -1;
    }
  );
});

test('parseIdosTicketText: an out-of-range arrival minute throws, with the full text appended', () => {
  assert.throws(
    () => {
      parseIdosTicketText(IDOS_BODY_TEXT_MINUTE_OUT_OF_RANGE);
    },
    function (err) {
      return err instanceof Error && err.message.indexOf(IDOS_BODY_TEXT_MINUTE_OUT_OF_RANGE) !== -1;
    }
  );
});

// --- resolveTransportSenderMode (D-01, quick-260804-bs7) --------------------

test('resolveTransportSenderMode: an entry with mode "body" resolves to "body"', () => {
  assert.equal(resolveTransportSenderMode({ mode: 'body' }), 'body');
});

test('resolveTransportSenderMode: an entry with mode "ics" resolves to "ics"', () => {
  assert.equal(resolveTransportSenderMode({ mode: 'ics' }), 'ics');
});

test('resolveTransportSenderMode: no mode field, RegioJet address (no registered body parser) resolves to "ics" -- the owner\'s already-live Script Property JSON back-compat guarantee', () => {
  assert.equal(resolveTransportSenderMode({ identifyingEmail: 'jizdenky@regiojet.cz' }), 'ics');
});

test('resolveTransportSenderMode: no mode field, IDOS.cz address (registered body parser) resolves to "body" -- the registered-parser fallback', () => {
  assert.equal(resolveTransportSenderMode({ identifyingEmail: 'jizdenky@idos.svt.cz' }), 'body');
});

test('resolveTransportSenderMode: a null/undefined sender resolves to "ics" without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(resolveTransportSenderMode(null), 'ics');
    assert.equal(resolveTransportSenderMode(undefined), 'ics');
  });
});

// WR-01 (260804-bs7 review): an INVALID mode (typo, unsupported value, or
// empty string) must throw loudly rather than silently falling back to
// 'ics' -- silently degrading is exactly the "Kino Art" silent-no-op class
// this function otherwise exists to prevent.
test('resolveTransportSenderMode: an invalid mode value throws rather than silently falling back to "ics" (WR-01)', () => {
  assert.throws(() => resolveTransportSenderMode({ mode: 'Body', identifyingEmail: 'x@example.com' }));
  assert.throws(() => resolveTransportSenderMode({ mode: 'pdf', identifyingEmail: 'x@example.com' }));
  assert.throws(() => resolveTransportSenderMode({ mode: '', identifyingEmail: 'x@example.com' }));
});

// --- buildTransportBodyEntry (D-05, quick-260804-bs7) ------------------------

test('buildTransportBodyEntry: the parsed fixture plus Europe/Prague produces zero-padded wall-clock start/end with no trailing Z/offset', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsed, 'Europe/Prague');

  assert.deepEqual(entry.resource.start, { dateTime: '2026-09-04T06:05:00', timeZone: 'Europe/Prague' });
  assert.deepEqual(entry.resource.end, { dateTime: '2026-09-04T17:40:00', timeZone: 'Europe/Prague' });
});

test('buildTransportBodyEntry: resource.summary is the route joined with the » separator, resource.location is the departure station', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsed, 'Europe/Prague');

  assert.equal(entry.resource.summary, 'Zelené Údolí hl.n. » Neustadt Hbf');
  assert.equal(entry.resource.location, 'Zelené Údolí hl.n.');
});

test('buildTransportBodyEntry: resource.description contains the e-ticket code, the seats and the order code', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsed, 'Europe/Prague');

  assert.ok(entry.resource.description.indexOf('7QKMR2') !== -1);
  assert.ok(entry.resource.description.indexOf('412/31 412/33') !== -1);
  assert.ok(entry.resource.description.indexOf('PLTQ-DMVR-ZKBN') !== -1);
});

test('buildTransportBodyEntry: a parse result with seats null omits the seat line entirely and never embeds the literal word "null"', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(Object.assign({}, parsed, { seats: null }), 'Europe/Prague');

  assert.equal(entry.resource.description.indexOf('null'), -1);
  assert.equal(entry.resource.description.indexOf('412/31 412/33'), -1);
});

test('buildTransportBodyEntry: entry.ticketIdentifier is the order code, entry.uid is null, entry.summary equals the resource summary', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsed, 'Europe/Prague');

  assert.equal(entry.ticketIdentifier, 'PLTQ-DMVR-ZKBN');
  assert.equal(entry.uid, null);
  assert.equal(entry.summary, entry.resource.summary);
});

test('buildTransportBodyEntry: filenameDate is the calendar date only, and feeding it through the EXISTING buildTransportAttachmentFilename produces the expected filename', () => {
  const parsed = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsed, 'Europe/Prague');

  assert.equal(entry.filenameDate.toISOString().slice(0, 10), '2026-09-04');
  assert.equal(
    buildTransportAttachmentFilename(entry.summary, entry.filenameDate, entry.ticketIdentifier),
    'Zelené Údolí hl.n. » Neustadt Hbf - 2026-09-04 - PLTQ-DMVR-ZKBN.pdf'
  );
});

// --- stripTransportSummaryIdentifierPrefix (v0.8.1) --------------------------

test('stripTransportSummaryIdentifierPrefix: strips a leading "#<digits>: " prefix', () => {
  assert.equal(stripTransportSummaryIdentifierPrefix(REAL_SUMMARY), 'Z Ostrava, hl.n., do Praha, hl.n., sedadla: [2/15,2/16]');
});

test('stripTransportSummaryIdentifierPrefix: a summary with no such prefix is returned unchanged', () => {
  assert.equal(stripTransportSummaryIdentifierPrefix('No hash prefix here'), 'No hash prefix here');
});

test('stripTransportSummaryIdentifierPrefix: null/undefined returns an empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(stripTransportSummaryIdentifierPrefix(null), '');
    assert.equal(stripTransportSummaryIdentifierPrefix(undefined), '');
  });
});

// --- buildTransportIcsEntry (the 'ics'-mode counterpart, quick-260804-bs7) --

test('buildTransportIcsEntry: given the real parsed RegioJet event, returns the shared entry shape through no new parsing (except the v0.8.1 summary-prefix strip)', () => {
  const event = parseIcs(REAL_REGIOJET_ICS)[0];
  const entry = buildTransportIcsEntry(event);
  const expectedSummary = stripTransportSummaryIdentifierPrefix(event.summary);

  assert.deepEqual(entry.resource, Object.assign({}, buildEventResource(event), { summary: expectedSummary }));
  assert.equal(entry.uid, event.uid);
  assert.equal(entry.ticketIdentifier, extractTransportTicketIdentifier(event));
  assert.equal(entry.summary, expectedSummary);
  assert.equal(entry.filenameDate, event.start);
});

test('buildTransportIcsEntry: the real RegioJet fixture\'s "#<digits>: " prefix is stripped from both entry.summary and entry.resource.summary, while ticketIdentifier still carries the number separately', () => {
  const event = parseIcs(REAL_REGIOJET_ICS)[0];
  const entry = buildTransportIcsEntry(event);

  assert.equal(entry.summary.indexOf('#'), -1, 'entry.summary should not contain the "#" ticket-number prefix');
  assert.equal(entry.resource.summary.indexOf('#'), -1, 'entry.resource.summary (the calendar event title) should not contain the "#" ticket-number prefix');
  assert.equal(entry.ticketIdentifier, '7788123456');
});

// --- RegioJet cancellation entry status (D-03, quick-260813-dq2) -------------

test('buildTransportIcsEntry: on the CANCEL fixture\'s parsed event, entry.status is "CANCELLED"', () => {
  const cancelEvent = parseIcs(REAL_REGIOJET_CANCEL_ICS)[0];
  const entry = buildTransportIcsEntry(cancelEvent);

  assert.equal(entry.status, 'CANCELLED');
});

test('buildTransportIcsEntry: on the existing confirmation fixture\'s parsed event (no STATUS line), entry.status is null', () => {
  const confirmationEvent = parseIcs(REAL_REGIOJET_ICS)[0];
  const entry = buildTransportIcsEntry(confirmationEvent);

  assert.equal(entry.status, null);
});

test('buildTransportIcsEntry: THE IDENTITY PROOF -- the CANCEL entry\'s ticketIdentifier equals the confirmation entry\'s ticketIdentifier exactly (D-03, this identity IS the cancellation mechanism)', () => {
  const confirmationEvent = parseIcs(REAL_REGIOJET_ICS)[0];
  const cancelEvent = parseIcs(REAL_REGIOJET_CANCEL_ICS)[0];

  const confirmationEntry = buildTransportIcsEntry(confirmationEvent);
  const cancelEntry = buildTransportIcsEntry(cancelEvent);

  assert.equal(cancelEntry.ticketIdentifier, confirmationEntry.ticketIdentifier);
  assert.equal(cancelEntry.ticketIdentifier, '7788123456');
});

test('buildTransportIcsEntry: entry.resource carries no status field for either the confirmation or the cancel fixture (D-01 firewall holds through the entry builder)', () => {
  const confirmationEvent = parseIcs(REAL_REGIOJET_ICS)[0];
  const cancelEvent = parseIcs(REAL_REGIOJET_CANCEL_ICS)[0];

  assert.equal(Object.prototype.hasOwnProperty.call(buildTransportIcsEntry(confirmationEvent).resource, 'status'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(buildTransportIcsEntry(cancelEvent).resource, 'status'), false);
});

test('buildTransportBodyEntry: returns status === null for the IDOS.cz fixture (no ICS STATUS concept, out of scope for cancellation via this mechanism)', () => {
  const parsedTicket = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsedTicket, 'Europe/Prague');

  assert.equal(entry.status, null);
});

// --- RegioJet dtstamp threading (D-10, quick-260813-dq2 Task 3) -------------

test('buildTransportIcsEntry: entry.dtstamp on each of the three fixtures strictly equals that event\'s own parsed dtstamp, ordering confirmation < cancellation < rebooking', () => {
  const confirmationEvent = parseIcs(REAL_REGIOJET_ICS)[0];
  const cancelEvent = parseIcs(REAL_REGIOJET_CANCEL_ICS)[0];
  const rebookEvent = parseIcs(REAL_REGIOJET_REBOOK_ICS)[0];

  const confirmationEntry = buildTransportIcsEntry(confirmationEvent);
  const cancelEntry = buildTransportIcsEntry(cancelEvent);
  const rebookEntry = buildTransportIcsEntry(rebookEvent);

  assert.equal(confirmationEntry.dtstamp.getTime(), confirmationEvent.dtstamp.getTime());
  assert.equal(cancelEntry.dtstamp.getTime(), cancelEvent.dtstamp.getTime());
  assert.equal(rebookEntry.dtstamp.getTime(), rebookEvent.dtstamp.getTime());

  assert.ok(confirmationEntry.dtstamp.getTime() < cancelEntry.dtstamp.getTime());
  assert.ok(cancelEntry.dtstamp.getTime() < rebookEntry.dtstamp.getTime());
});

test('buildTransportBodyEntry: returns dtstamp === null for the IDOS.cz fixture (no DTSTAMP concept, same treatment as status)', () => {
  const parsedTicket = parseIdosTicketText(REAL_IDOS_BODY_TEXT);
  const entry = buildTransportBodyEntry(parsedTicket, 'Europe/Prague');

  assert.equal(entry.dtstamp, null);
});

test('buildTransportIcsEntry: entry.resource carries no dtstamp key for either the confirmation or the cancel fixture', () => {
  const confirmationEvent = parseIcs(REAL_REGIOJET_ICS)[0];
  const cancelEvent = parseIcs(REAL_REGIOJET_CANCEL_ICS)[0];

  assert.equal(Object.prototype.hasOwnProperty.call(buildTransportIcsEntry(confirmationEvent).resource, 'dtstamp'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(buildTransportIcsEntry(cancelEvent).resource, 'dtstamp'), false);
});

test('buildTransportIcsEntry: the rebooking entry\'s ticketIdentifier AND uid both equal the confirmation entry\'s, while its summary differs -- the real-world shape that makes Problem A reachable (fixture-fidelity proof)', () => {
  const confirmationEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_ICS)[0]);
  const rebookEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_REBOOK_ICS)[0]);

  assert.equal(rebookEntry.ticketIdentifier, confirmationEntry.ticketIdentifier);
  assert.equal(rebookEntry.uid, confirmationEntry.uid);
  assert.notEqual(rebookEntry.summary, confirmationEntry.summary);
});

// --- TRANSPORT_TICKETS_ACTION -------------------------------------------------

test('TRANSPORT_TICKETS_ACTION: name is "transport-tickets"', () => {
  assert.equal(TRANSPORT_TICKETS_ACTION.name, 'transport-tickets');
});

test('TRANSPORT_TICKETS_ACTION: .config is the SAME object reference as TRANSPORT_TICKETS_ACTION_CONFIG (getter identity, not a copy)', () => {
  assert.equal(TRANSPORT_TICKETS_ACTION.config, TRANSPORT_TICKETS_ACTION_CONFIG);
});

test('TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders: code default is now TWO entries -- RegioJet (explicit mode "ics") plus IDOS.cz (mode "body")', () => {
  assert.deepEqual(TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders, [
    { identifyingEmail: 'jizdenky@regiojet.cz', calendarId: null, insertPdfIntoEvent: false, mode: 'ics' },
    { identifyingEmail: 'jizdenky@idos.svt.cz', calendarId: null, insertPdfIntoEvent: false, mode: 'body' },
  ]);
});

test('TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders: the new two-entry default still passes the shared isValidTicketingPortalsShape validator (the extra mode key is ignored by design)', () => {
  assert.equal(isValidTicketingPortalsShape(TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders), true);
});

// --- partitionTransportEntriesByCancellation (D-03, quick-260813-dq2) -------
//
// Pure, unit-tested split on entry.status === 'CANCELLED' -- placed BETWEEN
// entry construction and the existing seenInBatch/isDuplicateTransportTicket
// filter in processTransportTicketJob, so a cancel entry never gets
// dedup-dropped (its identifier deliberately matches the very event it is
// meant to delete).

test('partitionTransportEntriesByCancellation: splits a mixed array on status === "CANCELLED"', () => {
  const toCancelEntry = { status: 'CANCELLED', ticketIdentifier: 'a' };
  const toCreateEntry = { status: null, ticketIdentifier: 'b' };

  const result = partitionTransportEntriesByCancellation([toCancelEntry, toCreateEntry]);

  assert.deepEqual(result.toCancel, [toCancelEntry]);
  assert.deepEqual(result.toCreate, [toCreateEntry]);
});

test('partitionTransportEntriesByCancellation: routes status null, an absent status key, and any other status value to toCreate', () => {
  const nullStatusEntry = { status: null, ticketIdentifier: 'a' };
  const absentStatusEntry = { ticketIdentifier: 'b' };
  const otherStatusEntry = { status: 'CONFIRMED', ticketIdentifier: 'c' };

  const result = partitionTransportEntriesByCancellation([nullStatusEntry, absentStatusEntry, otherStatusEntry]);

  assert.deepEqual(result.toCancel, []);
  assert.deepEqual(result.toCreate, [nullStatusEntry, absentStatusEntry, otherStatusEntry]);
});

test('partitionTransportEntriesByCancellation: preserves relative order within each bucket and does not mutate the input array or its entries', () => {
  const entries = [
    { status: 'CANCELLED', ticketIdentifier: '1' },
    { status: null, ticketIdentifier: '2' },
    { status: 'CANCELLED', ticketIdentifier: '3' },
    { status: null, ticketIdentifier: '4' },
  ];
  const entriesCopy = entries.map(function (entry) {
    return Object.assign({}, entry);
  });

  const result = partitionTransportEntriesByCancellation(entries);

  assert.deepEqual(result.toCancel.map(function (e) { return e.ticketIdentifier; }), ['1', '3']);
  assert.deepEqual(result.toCreate.map(function (e) { return e.ticketIdentifier; }), ['2', '4']);
  assert.deepEqual(entries, entriesCopy, 'input array/entries must not be mutated');
});

test('partitionTransportEntriesByCancellation: returns two empty arrays for an empty input', () => {
  const result = partitionTransportEntriesByCancellation([]);

  assert.deepEqual(result.toCancel, []);
  assert.deepEqual(result.toCreate, []);
});

// --- cancelTransportTicketEvent (D-04, D-05, quick-260813-dq2) --------------
//
// GAS-only (Calendar global) -- proven under Node via a fake global.Calendar,
// same fake-global technique test/script-properties.test.js already uses for
// PropertiesService. Saves/restores the previous global.Calendar value
// around each test so no other test in the file is affected.

function withFakeCalendar(events, fn) {
  const listCalls = [];
  const removeCalls = [];
  const previousCalendar = global.Calendar;

  global.Calendar = {
    Events: {
      list: function (calendarId, options) {
        listCalls.push({ calendarId: calendarId, options: options });

        // Real Calendar API filters server-side on privateExtendedProperty --
        // the fake mirrors that so the race regression below is proven by
        // the SAME query-scoping the real API performs, not by a hand-rolled
        // find() the fake alone would need.
        const match = /^ticketIdentifier=(.*)$/.exec(options.privateExtendedProperty || '');
        const wantedIdentifier = match ? match[1] : null;
        const items = events.filter(function (event) {
          return (
            event.extendedProperties &&
            event.extendedProperties.private &&
            event.extendedProperties.private.ticketIdentifier === wantedIdentifier
          );
        });

        return { items: items };
      },
      remove: function (calendarId, eventId) {
        removeCalls.push({ calendarId: calendarId, eventId: eventId });
      },
    },
  };

  try {
    fn({ listCalls: listCalls, removeCalls: removeCalls });
  } finally {
    if (previousCalendar === undefined) {
      delete global.Calendar;
    } else {
      global.Calendar = previousCalendar;
    }
  }
}

test('cancelTransportTicketEvent: a match found calls Calendar.Events.remove exactly once with (calendarId, event.id), and the lookup is scoped via privateExtendedProperty', () => {
  const existingEvent = { id: 'event-1', extendedProperties: { private: { ticketIdentifier: '7788123456' } } };

  withFakeCalendar([existingEvent], function (calls) {
    cancelTransportTicketEvent('7788123456', 'calendar-a');

    assert.equal(calls.listCalls.length, 1);
    assert.equal(calls.listCalls[0].options.privateExtendedProperty, 'ticketIdentifier=7788123456');
    assert.equal(calls.removeCalls.length, 1);
    assert.deepEqual(calls.removeCalls[0], { calendarId: 'calendar-a', eventId: 'event-1' });
  });
});

test('cancelTransportTicketEvent: no matching event -> no Calendar.Events.remove call at all, no throw (D-05, accepted limitation)', () => {
  withFakeCalendar([], function (calls) {
    assert.doesNotThrow(function () {
      cancelTransportTicketEvent('does-not-exist', 'calendar-a');
    });

    assert.equal(calls.listCalls.length, 1);
    assert.equal(calls.removeCalls.length, 0);
  });
});

test('cancelTransportTicketEvent: a falsy ticketIdentifier (null) never reaches the Calendar API -- neither list nor remove is called, no throw', () => {
  withFakeCalendar([], function (calls) {
    assert.doesNotThrow(function () {
      cancelTransportTicketEvent(null, 'calendar-a');
    });

    assert.equal(calls.listCalls.length, 0);
    assert.equal(calls.removeCalls.length, 0);
  });
});

test('cancelTransportTicketEvent: a falsy ticketIdentifier (empty string) never reaches the Calendar API -- neither list nor remove is called, no throw', () => {
  withFakeCalendar([], function (calls) {
    assert.doesNotThrow(function () {
      cancelTransportTicketEvent('', 'calendar-a');
    });

    assert.equal(calls.listCalls.length, 0);
    assert.equal(calls.removeCalls.length, 0);
  });
});

// THE RACE REGRESSION (D-04): a cancelled ticket and a newly-purchased
// ticket for the SAME date/time but DIFFERENT ticket numbers must never
// cross-contaminate -- cancelling one issues a lookup scoped to that ticket
// number only and removes only that event's id. Holds regardless of which
// email is processed first (this test proves the STATIC guarantee: strict
// ticket-number matching, no date/time fallback -- processing order is
// irrelevant because the lookup never considers date/time at all).
test('cancelTransportTicketEvent: strict ticket-number matching -- cancelling one ticket never affects a different ticket for the SAME date/time (D-04 race regression, order-independent)', () => {
  const cancelledTicketEvent = {
    id: 'event-cancelled-ticket',
    start: { dateTime: '2026-08-18T17:36:00+02:00' },
    extendedProperties: { private: { ticketIdentifier: '7788123456' } },
  };
  const newlyPurchasedTicketEvent = {
    id: 'event-new-ticket',
    start: { dateTime: '2026-08-18T17:36:00+02:00' },
    extendedProperties: { private: { ticketIdentifier: '9900112233' } },
  };

  withFakeCalendar([cancelledTicketEvent, newlyPurchasedTicketEvent], function (calls) {
    cancelTransportTicketEvent('7788123456', 'calendar-a');

    assert.equal(calls.listCalls[0].options.privateExtendedProperty, 'ticketIdentifier=7788123456');
    assert.equal(calls.removeCalls.length, 1);
    assert.equal(calls.removeCalls[0].eventId, 'event-cancelled-ticket');
    assert.notEqual(calls.removeCalls[0].eventId, 'event-new-ticket');
  });
});

// --- buildTransportEventPrivateProperties (D-10, quick-260813-dq2 Task 3) ---
//
// New pure helper, the SINGLE writer of the extendedProperties.private tag
// isTransportCancellationStale later reads. Placed next to the entry
// builders in the production file.

test('buildTransportEventPrivateProperties: identifier + a real dtstamp -> { ticketIdentifier, dtstamp: <that entry\'s ISO string> }', () => {
  const dtstamp = new Date('2026-08-13T07:34:45.000Z');
  const result = buildTransportEventPrivateProperties({ ticketIdentifier: '7788123456', dtstamp: dtstamp });

  assert.deepEqual(result, { ticketIdentifier: '7788123456', dtstamp: '2026-08-13T07:34:45.000Z' });
});

test('buildTransportEventPrivateProperties: identifier + dtstamp: null -> { ticketIdentifier } with NO dtstamp key present at all', () => {
  const result = buildTransportEventPrivateProperties({ ticketIdentifier: '7788123456', dtstamp: null });

  assert.deepEqual(result, { ticketIdentifier: '7788123456' });
  assert.equal(Object.prototype.hasOwnProperty.call(result, 'dtstamp'), false);
  assert.equal(JSON.stringify(result).indexOf('null'), -1, 'serialized object must not contain the literal 4-character null token');
});

test('buildTransportEventPrivateProperties: falsy ticketIdentifier -> null, meaning the write loop sets no extendedProperties object at all (today\'s exact behavior, preserved)', () => {
  assert.equal(buildTransportEventPrivateProperties({ ticketIdentifier: null, dtstamp: new Date() }), null);
  assert.equal(buildTransportEventPrivateProperties({ ticketIdentifier: '', dtstamp: new Date() }), null);
});

test('buildTransportEventPrivateProperties: a non-Date / invalid-Date dtstamp -> the key is omitted rather than emitting an invalid string', () => {
  const withNonDate = buildTransportEventPrivateProperties({ ticketIdentifier: 'x', dtstamp: 'not-a-date' });
  assert.deepEqual(withNonDate, { ticketIdentifier: 'x' });

  const withInvalidDate = buildTransportEventPrivateProperties({ ticketIdentifier: 'x', dtstamp: new Date('not-a-date') });
  assert.deepEqual(withInvalidDate, { ticketIdentifier: 'x' });
});

// --- isTransportCancellationStale (D-11, quick-260813-dq2 Task 3) -----------
//
// True ONLY when the found event's stored dtstamp is present AND strictly
// newer than the cancellation entry's own dtstamp. Every missing/
// unparseable case returns false so the caller falls back to deleting -- an
// absent optional signal must never block D-05.

test('isTransportCancellationStale: stored dtstamp NEWER than the cancel dtstamp -> true', () => {
  const existingEvent = { extendedProperties: { private: { dtstamp: '2026-08-13T07:39:19.000Z' } } };
  assert.equal(isTransportCancellationStale(existingEvent, new Date('2026-08-13T07:34:45.000Z')), true);
});

test('isTransportCancellationStale: stored dtstamp OLDER than the cancel dtstamp -> false', () => {
  const existingEvent = { extendedProperties: { private: { dtstamp: '2026-08-02T17:45:58.000Z' } } };
  assert.equal(isTransportCancellationStale(existingEvent, new Date('2026-08-13T07:34:45.000Z')), false);
});

test('isTransportCancellationStale: stored dtstamp EQUAL to the cancel dtstamp -> false (strictly newer, not newer-or-equal)', () => {
  const same = '2026-08-13T07:34:45.000Z';
  const existingEvent = { extendedProperties: { private: { dtstamp: same } } };
  assert.equal(isTransportCancellationStale(existingEvent, new Date(same)), false);
});

test('isTransportCancellationStale: event has no stored dtstamp key -> false', () => {
  const existingEvent = { extendedProperties: { private: { ticketIdentifier: '7788123456' } } };
  assert.equal(isTransportCancellationStale(existingEvent, new Date('2026-08-13T07:34:45.000Z')), false);
});

test('isTransportCancellationStale: event has no extendedProperties at all, or a null one -> false, no throw', () => {
  assert.doesNotThrow(function () {
    assert.equal(isTransportCancellationStale({}, new Date()), false);
  });
  assert.doesNotThrow(function () {
    assert.equal(isTransportCancellationStale({ extendedProperties: null }, new Date()), false);
  });
  assert.doesNotThrow(function () {
    assert.equal(isTransportCancellationStale({ extendedProperties: { private: null } }, new Date()), false);
  });
});

test('isTransportCancellationStale: cancelDtstamp null/undefined -> false', () => {
  const existingEvent = { extendedProperties: { private: { dtstamp: '2026-08-13T07:39:19.000Z' } } };
  assert.equal(isTransportCancellationStale(existingEvent, null), false);
  assert.equal(isTransportCancellationStale(existingEvent, undefined), false);
});

test('isTransportCancellationStale: stored value unparseable -> false', () => {
  const existingEvent = { extendedProperties: { private: { dtstamp: 'not-a-date' } } };
  assert.equal(isTransportCancellationStale(existingEvent, new Date('2026-08-13T07:34:45.000Z')), false);
});

// --- cancelTransportTicketEvent, stale-cancellation guard (D-11, quick-260813-dq2 Task 3) ---
//
// The optional THIRD parameter is the cancellation entry's OWN dtstamp. All
// 5 pre-existing two-argument tests above still pass untouched -- the third
// parameter's absence means today's behavior (unchanged fallback: delete).

test('cancelTransportTicketEvent: found event tagged with a dtstamp NEWER than the cancel entry\'s -> Calendar.Events.remove is NEVER called, no throw (the stale cancellation is skipped)', () => {
  const existingEvent = {
    id: 'event-1',
    extendedProperties: { private: { ticketIdentifier: '7788123456', dtstamp: '2026-08-13T07:39:19.000Z' } },
  };

  withFakeCalendar([existingEvent], function (calls) {
    assert.doesNotThrow(function () {
      cancelTransportTicketEvent('7788123456', 'calendar-a', new Date('2026-08-13T07:34:45.000Z'));
    });

    assert.equal(calls.removeCalls.length, 0);
  });
});

test('cancelTransportTicketEvent: found event tagged with an OLDER dtstamp -> removed exactly as before (D-05\'s core case must not regress)', () => {
  const existingEvent = {
    id: 'event-1',
    extendedProperties: { private: { ticketIdentifier: '7788123456', dtstamp: '2026-08-02T17:45:58.000Z' } },
  };

  withFakeCalendar([existingEvent], function (calls) {
    cancelTransportTicketEvent('7788123456', 'calendar-a', new Date('2026-08-13T07:34:45.000Z'));

    assert.equal(calls.removeCalls.length, 1);
    assert.equal(calls.removeCalls[0].eventId, 'event-1');
  });
});

test('cancelTransportTicketEvent: found event with NO stored dtstamp (pre-feature event, or one written by a uid-less entry) -> removed, unchanged fallback', () => {
  const existingEvent = {
    id: 'event-1',
    extendedProperties: { private: { ticketIdentifier: '7788123456' } },
  };

  withFakeCalendar([existingEvent], function (calls) {
    cancelTransportTicketEvent('7788123456', 'calendar-a', new Date('2026-08-13T07:34:45.000Z'));

    assert.equal(calls.removeCalls.length, 1);
  });
});

test('cancelTransportTicketEvent: cancel entry with no dtstamp of its own (third argument null) against a tagged event -> removed, unchanged fallback', () => {
  const existingEvent = {
    id: 'event-1',
    extendedProperties: { private: { ticketIdentifier: '7788123456', dtstamp: '2026-08-13T07:39:19.000Z' } },
  };

  withFakeCalendar([existingEvent], function (calls) {
    cancelTransportTicketEvent('7788123456', 'calendar-a', null);

    assert.equal(calls.removeCalls.length, 1);
  });
});

// --- filterTransportEntriesToCreate (D-08, quick-260813-dq2 Task 3) --------
//
// The EXISTING seenInBatch + isDuplicateTransportTicket filter, extracted
// out of processTransportTicketJob so it is reachable under Node, with
// exactly one behavior change: isDuplicateTransportTicket is now consulted
// ONLY for an entry with no `uid` (D-08). seenInBatch still applies to
// EVERY entry, uid-bearing or not (deliberate scope limit).

test('filterTransportEntriesToCreate: THE PROBLEM A REGRESSION -- a uid-bearing entry whose ticketIdentifier MATCHES an existing event SURVIVES the filter, and Calendar.Events.list is never called for it', () => {
  const existingEvent = { id: 'event-1', extendedProperties: { private: { ticketIdentifier: '7788123456' } } };
  const reissueEntry = { ticketIdentifier: '7788123456', uid: '-9876543210@regiojet.cz' };

  withFakeCalendar([existingEvent], function (calls) {
    const result = filterTransportEntriesToCreate([reissueEntry], 'calendar-a');

    assert.deepEqual(result, [reissueEntry]);
    assert.equal(calls.listCalls.length, 0, 'isDuplicateTransportTicket must not be consulted at all for a uid-bearing entry');
  });
});

test('filterTransportEntriesToCreate: THE IDOS.cz PROTECTION IS INTACT -- a uid-less entry whose ticketIdentifier matches an existing event is DROPPED, and the lookup WAS issued', () => {
  const existingEvent = { id: 'event-1', extendedProperties: { private: { ticketIdentifier: 'PLTQ-DMVR-ZKBN' } } };
  const duplicateBodyEntry = { ticketIdentifier: 'PLTQ-DMVR-ZKBN', uid: null };

  withFakeCalendar([existingEvent], function (calls) {
    const result = filterTransportEntriesToCreate([duplicateBodyEntry], 'calendar-a');

    assert.deepEqual(result, []);
    assert.equal(calls.listCalls.length, 1, 'isDuplicateTransportTicket must still be consulted for a uid-less entry -- the narrowing is scoped, not global');
  });
});

test('filterTransportEntriesToCreate: WR-01 IS INTACT -- two uid-less same-batch entries sharing one ticketIdentifier collapse to the first only', () => {
  const first = { ticketIdentifier: 'PLTQ-DMVR-ZKBN', uid: null };
  const second = { ticketIdentifier: 'PLTQ-DMVR-ZKBN', uid: null };

  withFakeCalendar([], function () {
    const result = filterTransportEntriesToCreate([first, second], 'calendar-a');

    assert.deepEqual(result, [first]);
  });
});

test('filterTransportEntriesToCreate: seenInBatch still applies to uid-bearing entries too (D-08\'s deliberate scope limit) -- two same-batch uid-bearing entries sharing one ticketIdentifier still collapse to the first', () => {
  const first = { ticketIdentifier: '7788123456', uid: '-9876543210@regiojet.cz' };
  const second = { ticketIdentifier: '7788123456', uid: '-9876543210@regiojet.cz' };

  withFakeCalendar([], function (calls) {
    const result = filterTransportEntriesToCreate([first, second], 'calendar-a');

    assert.deepEqual(result, [first]);
    assert.equal(calls.listCalls.length, 0, 'uid-bearing entries never consult isDuplicateTransportTicket at all');
  });
});

test('filterTransportEntriesToCreate: a uid-less entry with a falsy ticketIdentifier survives (unchanged: the falsy guard makes isDuplicateTransportTicket return false and nothing is tracked in seenInBatch)', () => {
  const entry = { ticketIdentifier: null, uid: null };

  withFakeCalendar([], function (calls) {
    const result = filterTransportEntriesToCreate([entry], 'calendar-a');

    assert.deepEqual(result, [entry]);
    assert.equal(calls.listCalls.length, 0);
  });
});

test('filterTransportEntriesToCreate: relative order is preserved and the input array/entries are not mutated', () => {
  const entries = [
    { ticketIdentifier: 'a', uid: 'uid-a' },
    { ticketIdentifier: 'b', uid: null },
    { ticketIdentifier: 'c', uid: 'uid-c' },
  ];
  const entriesCopy = entries.map(function (entry) {
    return Object.assign({}, entry);
  });

  withFakeCalendar([], function () {
    const result = filterTransportEntriesToCreate(entries, 'calendar-a');

    assert.deepEqual(
      result.map(function (e) {
        return e.ticketIdentifier;
      }),
      ['a', 'b', 'c']
    );
    assert.deepEqual(entries, entriesCopy, 'input array/entries must not be mutated');
  });
});

// --- THE TWO ORDERING SCENARIOS (D-12, quick-260813-dq2 Task 3) ------------
//
// A shared, MUTATING fake calendar store per scenario (extends the
// withFakeCalendar shape above so remove actually deletes), plus a small
// local simulateTransportWrite(store, entry) test helper that upserts by
// entry.uid and tags the stored event via the EXPORTED
// buildTransportEventPrivateProperties -- i.e. the write loop's tagging is
// exercised through the real production helper, never a reimplementation of
// it. Each scenario runs the exported decision helpers in the SAME sequence
// processTransportTicketJob invokes them (partition -> cancel the toCancel
// bucket -> filter the toCreate bucket -> write what survives), once per
// message, in the stated order. These prove the COMPOSED DECISION LOGIC, NOT
// processTransportTicketJob's own GAS-only wiring, which only the live
// checkpoint can reach (D-12, Task 4's honest-scope note).

function withMutatingFakeCalendar(initialEvents, fn) {
  const store = initialEvents.slice();
  const removeCalls = [];
  const previousCalendar = global.Calendar;

  global.Calendar = {
    Events: {
      list: function (calendarId, options) {
        const match = /^ticketIdentifier=(.*)$/.exec(options.privateExtendedProperty || '');
        const wantedIdentifier = match ? match[1] : null;
        const items = store.filter(function (event) {
          return (
            event.extendedProperties &&
            event.extendedProperties.private &&
            event.extendedProperties.private.ticketIdentifier === wantedIdentifier
          );
        });

        return { items: items };
      },
      remove: function (calendarId, eventId) {
        removeCalls.push({ calendarId: calendarId, eventId: eventId });

        const index = store.findIndex(function (event) {
          return event.id === eventId;
        });
        if (index !== -1) {
          store.splice(index, 1);
        }
      },
    },
  };

  try {
    fn(store, removeCalls);
  } finally {
    if (previousCalendar === undefined) {
      delete global.Calendar;
    } else {
      global.Calendar = previousCalendar;
    }
  }
}

function simulateTransportWrite(store, entry) {
  const privateProperties = buildTransportEventPrivateProperties(entry);
  const existingIndex = store.findIndex(function (event) {
    return event.iCalUID === entry.uid;
  });

  const eventData = {
    id: existingIndex === -1 ? 'event-' + (store.length + 1) : store[existingIndex].id,
    iCalUID: entry.uid,
    summary: entry.summary,
  };
  if (privateProperties) {
    eventData.extendedProperties = { private: privateProperties };
  }

  if (existingIndex === -1) {
    store.push(eventData);
  } else {
    store[existingIndex] = eventData;
  }
}

function processTransportEntryForTest(store, calendarId, entry) {
  const partitioned = partitionTransportEntriesByCancellation([entry]);

  partitioned.toCancel.forEach(function (cancelEntry) {
    cancelTransportTicketEvent(cancelEntry.ticketIdentifier, calendarId, cancelEntry.dtstamp);
  });

  const toCreate = filterTransportEntriesToCreate(partitioned.toCreate, calendarId);
  toCreate.forEach(function (createEntry) {
    simulateTransportWrite(store, createEntry);
  });
}

test('ORDER 1 (cancel first, then rebooking): the rebooking survives and the store ends with exactly ONE event carrying the rebooking\'s summary and dtstamp tag', () => {
  const confirmationEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_ICS)[0]);
  const cancelEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_CANCEL_ICS)[0]);
  const rebookEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_REBOOK_ICS)[0]);

  withMutatingFakeCalendar([], function (store) {
    simulateTransportWrite(store, confirmationEntry);

    processTransportEntryForTest(store, 'calendar-a', cancelEntry);
    processTransportEntryForTest(store, 'calendar-a', rebookEntry);

    assert.equal(store.length, 1);
    assert.equal(store[0].summary, rebookEntry.summary);
    assert.equal(store[0].extendedProperties.private.dtstamp, rebookEntry.dtstamp.toISOString());
  });
});

test('ORDER 2 (rebooking first, then cancel -- THE EXACT SCENARIO THE OWNER ASKED ABOUT): the rebooked event survives untouched, and the stale cancellation records no remove call', () => {
  const confirmationEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_ICS)[0]);
  const cancelEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_CANCEL_ICS)[0]);
  const rebookEntry = buildTransportIcsEntry(parseIcs(REAL_REGIOJET_REBOOK_ICS)[0]);

  withMutatingFakeCalendar([], function (store, removeCalls) {
    simulateTransportWrite(store, confirmationEntry);

    processTransportEntryForTest(store, 'calendar-a', rebookEntry);
    processTransportEntryForTest(store, 'calendar-a', cancelEntry);

    assert.equal(store.length, 1);
    assert.equal(store[0].summary, rebookEntry.summary);
    assert.equal(store[0].extendedProperties.private.dtstamp, rebookEntry.dtstamp.toISOString());
    assert.equal(removeCalls.length, 0, 'the stale cancellation must not call Calendar.Events.remove');
  });
});

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

// --- buildTransportIcsEntry (the 'ics'-mode counterpart, quick-260804-bs7) --

test('buildTransportIcsEntry: given the real parsed RegioJet event, returns the shared entry shape through no new parsing', () => {
  const event = parseIcs(REAL_REGIOJET_ICS)[0];
  const entry = buildTransportIcsEntry(event);

  assert.deepEqual(entry.resource, buildEventResource(event));
  assert.equal(entry.uid, event.uid);
  assert.equal(entry.ticketIdentifier, extractTransportTicketIdentifier(event));
  assert.equal(entry.summary, event.summary);
  assert.equal(entry.filenameDate, event.start);
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

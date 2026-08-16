'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  parseEnigooTicketText,
  parseKinoArtTicketText,
  parseTicketmasterCzTicketText,
  resolveTicketingPortal,
  resolveTicketingCalendarId,
  addMinutesToWallClockComponents,
  formatWallClockComponentsIso,
  buildTicketAttachmentFilename,
  isTicketPdfAttachment,
  findTicketPdfAttachments,
  findKinoArtTicketPdfAttachment,
  findTicketmasterCzTicketPdfAttachment,
  resolveTicketProcessingJobs,
  TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL,
  TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL,
  TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL,
} = require('../src/07-action-ticketing-portals.js');

// --- parseEnigooTicketText ---------------------------------------------------
//
// v0.6.0 new action (quick-260731-tix): enigoo.cz confirmation emails carry
// NO usable event data in the body — everything (event name, date/time,
// venue) is inside an attached PDF ticket, extracted via Drive's PDF-to-
// Google-Docs OCR conversion pipeline.
//
// ROUND 5 (live-test-driven, the ACTUAL root-cause rewrite): the owner's
// round-3 full-extracted-text diagnostic finally paid off with a REAL live
// failure carrying the ACTUAL raw text `Body.getText()` returns for this
// ticket — reproduced VERBATIM below (REAL_ENIGOO_LIVE_FAILURE_TEXT_SINGLE /
// _DOUBLE). It conclusively disproved the original "one field per line"
// assumption this parser was built on (the event name and date/time are
// packed onto the SAME real paragraph; the location/price/discount onto a
// DIFFERENT single real paragraph) — see src/07-action-ticketing-portals.js's
// parseEnigooTicketText JSDoc for the full incident history and the new
// pattern-anchored extraction model (date/time regex + "Cena/price" label
// as anchors, not line position).
//
// The OLDER REAL_ENIGOO_TICKET_TEXT fixture below (one field per line) is
// now known to NOT reflect the real Body.getText() output shape — it was
// originally derived from reading the source PDF via a different
// extraction mechanism, not Google's own OCR. It is kept here (not deleted)
// specifically because the new anchor-based parser is verified to ALSO
// still handle this shape correctly (the anchors don't care how many real
// lines/paragraphs the fields are spread across, as long as whitespace
// separates them) — this is a valuable "works either way" regression proof,
// not evidence this was ever the true OCR shape.

const REAL_ENIGOO_TICKET_TEXT = [
  'Letní hudební festival',
  '15.08.2026 19:00',
  'Nádvoří kulturního domu',
  'Cena/price: 290 Kč',
  'Sleva/discount:',
  'Kulturní spolek z. s., Hlavní 123, 100 00 Praha, vedená u Krajského soudu v Praze (X99999), DUZP 2.8.2026',
  'Vstupenka je osvobozena od DPH dle ustanovení § 61 písm. e) zákona o dani z přidané hodnoty',
  '24601',
].join('\n');

test('parseEnigooTicketText: a one-field-per-line fixture (the ORIGINAL, now-disproven shape assumption) still parses the event name/location/date/time correctly under the new anchor-based extraction', () => {
  // Asserts individual fields (not a full deepEqual against a literal
  // object) because this fixture's field ORDER (legal/VAT text BEFORE the
  // trailing ticket number) differs from the REAL live-failure text's
  // order (ticket number immediately after "Sleva/discount:", legal text
  // after that) -- ticketIdentifier extraction is documented to depend on
  // that real ordering (see parseEnigooTicketText's own JSDoc, Anchor 4),
  // so it is deliberately not asserted against this non-representative
  // fixture's field order.
  const parsed = parseEnigooTicketText(REAL_ENIGOO_TICKET_TEXT);
  assert.equal(parsed.eventName, 'Letní hudební festival');
  assert.equal(parsed.location, 'Nádvoří kulturního domu');
  assert.equal(parsed.year, 2026);
  assert.equal(parsed.month, 7);
  assert.equal(parsed.day, 15);
  assert.equal(parsed.hour, 19);
  assert.equal(parsed.minute, 0);
});

test('parseEnigooTicketText: a second ticket page (same event, different trailing ticket number) parses to the SAME event name/location/date/time — proves one-event-per-purchase falls out naturally from anchoring on the FIRST date/time match (ticketIdentifier itself legitimately differs per physical page, since each page carries its own real ticket number)', () => {
  const page2Text = REAL_ENIGOO_TICKET_TEXT.replace('24601', '24600');
  const parsed1 = parseEnigooTicketText(REAL_ENIGOO_TICKET_TEXT);
  const parsed2 = parseEnigooTicketText(page2Text);
  assert.equal(parsed2.eventName, parsed1.eventName);
  assert.equal(parsed2.location, parsed1.location);
  assert.equal(parsed2.year, parsed1.year);
  assert.equal(parsed2.month, parsed1.month);
  assert.equal(parsed2.day, parsed1.day);
  assert.equal(parsed2.hour, parsed1.hour);
  assert.equal(parsed2.minute, parsed1.minute);
});

// --- OCR text modeled on a real live-failure shape (quick-260731-tix
// round 5) --------------------------------------------------------------------
//
// This models the shape `DocumentApp.getBody().getText()` actually returned
// during the round-5 investigation (identifying event/venue/purchase details
// replaced with fictional equivalents): event name + date/time packed onto
// ONE line/paragraph; location + price + discount packed onto a DIFFERENT
// single line/paragraph; the per-ticket number and the legal/VAT text each
// get their own line. This whole block repeats (except the trailing ticket
// number) for the second ticket page of the SAME purchase — proving these
// tests against both the single-occurrence text AND the full repeated text
// is what actually confirms the "one event per purchase, anchored on the
// FIRST date/time match" guarantee, not an assumption.

const REAL_ENIGOO_LIVE_FAILURE_TEXT_SINGLE = [
  'Letní hudební festival 15.08.2026 19:00 ',
  'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount: ',
  '24601',
  'Kulturní spolek z. s., Hlavní 123, 100 00 Praha, vedená u Krajského soudu v Praze (X99999), DUZP 2.8.2026 Vstupenka je osvobozena od DPH dle ustanovení § 61 písm. e) zákona o dani z přidané hodnoty ',
].join('\n');

const REAL_ENIGOO_LIVE_FAILURE_TEXT_DOUBLE = [
  'Letní hudební festival 15.08.2026 19:00 ',
  'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount: ',
  '24601',
  'Kulturní spolek z. s., Hlavní 123, 100 00 Praha, vedená u Krajského soudu v Praze (X99999), DUZP 2.8.2026 Vstupenka je osvobozena od DPH dle ustanovení § 61 písm. e) zákona o dani z přidané hodnoty ',
  'Letní hudební festival 15.08.2026 19:00 ',
  'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount: ',
  '24600',
  'Kulturní spolek z. s., Hlavní 123, 100 00 Praha, vedená u Krajského soudu v Praze (X99999), DUZP 2.8.2026 Vstupenka je osvobozena od DPH dle ustanovení § 61 písm. e) zákona o dani z přidané hodnoty',
].join('\n');

test('parseEnigooTicketText: the REAL raw live-failure text (single ticket occurrence) parses correctly — event name and date/time packed on the same real line, location packed with price/discount on a different real line', () => {
  assert.deepEqual(parseEnigooTicketText(REAL_ENIGOO_LIVE_FAILURE_TEXT_SINGLE), {
    eventName: 'Letní hudební festival',
    location: 'Nádvoří kulturního domu',
    year: 2026,
    month: 7,
    day: 15,
    hour: 19,
    minute: 0,
    ticketIdentifier: '24601',
  });
});

test('parseEnigooTicketText: the REAL raw live-failure text with BOTH ticket occurrences repeated parses to the SAME event data (INCLUDING the same ticketIdentifier, "24601" not "24600") — proves the parser anchors on the FIRST date/time match and does not get confused by the text repeating', () => {
  assert.deepEqual(parseEnigooTicketText(REAL_ENIGOO_LIVE_FAILURE_TEXT_DOUBLE), parseEnigooTicketText(REAL_ENIGOO_LIVE_FAILURE_TEXT_SINGLE));
});

// --- ticketIdentifier extraction (live-test-driven, quick-260731-tix round
// 8: the DEDUP SAFETY NET's stable key) -----------------------------------
//
// The owner hit a real duplicate-event bug: processing the enigoo.cz email
// ONE time created TWO calendar events for the same purchase, and manually
// re-running against the same email added a THIRD -- there was no
// idempotency protection at all (unlike the ICS action's iCalUID-based
// Events.import, or the booking.com action's confirmationNumber tag +
// findOrTagMatchingEvent safety-net). ticketIdentifier is the first
// per-ticket number found strictly after the "Sleva/discount:" label (the
// literal anchor immediately preceding it on real tickets), giving a
// stable, already-available key for the whole purchase -- confirming it
// resolves to "24601" (the FIRST ticket page's own number), not "24600"
// (the second page's), from the real repeated-ticket-text fixture.

test('parseEnigooTicketText: ticketIdentifier resolves to "24601" (the FIRST ticket page), not "24600", from the real repeated-ticket-text fixture', () => {
  const parsed = parseEnigooTicketText(REAL_ENIGOO_LIVE_FAILURE_TEXT_DOUBLE);
  assert.equal(parsed.ticketIdentifier, '24601');
  assert.notEqual(parsed.ticketIdentifier, '24600');
});

test('parseEnigooTicketText: ticketIdentifier is null (never throws) when "Sleva/discount:" is present but no digit run follows it', () => {
  const noTicketNumberText = ['Letní hudební festival', '15.08.2026 19:00', 'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount:'].join('\n');
  assert.equal(parseEnigooTicketText(noTicketNumberText).ticketIdentifier, null);
});

test('parseEnigooTicketText: ticketIdentifier is null (never throws) when "Sleva/discount:" is not present at all', () => {
  const noDiscountLabelText = ['Letní hudební festival', '15.08.2026 19:00', 'Nádvoří kulturního domu Cena/price: 290 Kč'].join('\n');
  assert.equal(parseEnigooTicketText(noDiscountLabelText).ticketIdentifier, null);
});

// --- separator-agnosticism (live-test-driven, quick-260731-tix round 4,
// preserved/adapted for round 5): Google Apps Script's Document Service is
// long-documented to join paragraphs with a CARRIAGE RETURN ('\r'), not a
// bare LINE FEED ('\n') — a well-known, stable Apps Script quirk. Round 4's
// specific line-array-splitting FIX has been entirely removed by the round-5
// rewrite (there is no line array left to split), but the underlying
// insight (don't assume one specific separator character) remains valid and
// is now expressed differently: every anchor regex uses `\s`/`[\s\S]`, which
// transparently absorbs `\r`, `\n`, or `\r\n` identically. These tests prove
// that separator-agnosticism survives the round-5 rewrite.

test('parseEnigooTicketText: a fixture joined with CARRIAGE RETURN (\\r), not \\n, still parses correctly (separator-agnostic anchoring)', () => {
  const crText = ['Letní hudební festival', '15.08.2026 19:00', 'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount:'].join('\r');
  assert.deepEqual(parseEnigooTicketText(crText), {
    eventName: 'Letní hudební festival',
    location: 'Nádvoří kulturního domu',
    year: 2026,
    month: 7,
    day: 15,
    hour: 19,
    minute: 0,
    ticketIdentifier: null,
  });
});

test('parseEnigooTicketText: a fixture joined with Windows-style CRLF (\\r\\n) also parses correctly (separator-agnostic anchoring)', () => {
  const crlfText = ['Letní hudební festival', '15.08.2026 19:00', 'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount:'].join('\r\n');
  assert.deepEqual(parseEnigooTicketText(crlfText), {
    eventName: 'Letní hudební festival',
    location: 'Nádvoří kulturního domu',
    year: 2026,
    month: 7,
    day: 15,
    hour: 19,
    minute: 0,
    ticketIdentifier: null,
  });
});

test('parseEnigooTicketText: stray blank lines around the fields do not disrupt anchor-based extraction', () => {
  const withBlankLines = ['', 'Letní hudební festival', '', '15.08.2026 19:00', '', 'Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount:', ''].join(
    '\n'
  );
  assert.deepEqual(parseEnigooTicketText(withBlankLines), {
    eventName: 'Letní hudební festival',
    location: 'Nádvoří kulturního domu',
    year: 2026,
    month: 7,
    day: 15,
    hour: 19,
    minute: 0,
    ticketIdentifier: null,
  });
});

test('parseEnigooTicketText: text with no location/"Cena/price" marker at all throws a controlled error', () => {
  assert.throws(() => parseEnigooTicketText('Letní hudební festival 15.08.2026 19:00'));
});

test('parseEnigooTicketText: text with no recognizable date/time pattern anywhere throws a controlled error', () => {
  assert.throws(() => parseEnigooTicketText('Event\nnot a date\nVenue Cena/price: 100 Kč'));
});

// --- diagnostic-on-failure: full extracted OCR text included in every thrown
// error (live-test-driven, quick-260731-tix round 3, still in effect after the
// round-5 rewrite) --------------------------------------------------------------
//
// It was THIS exact mechanism that supplied the real raw text motivating the
// round-5 rewrite above. Since this codebase's existing failure-notification
// path already emails the owner the full thrown error message on any action
// failure, including the COMPLETE extracted text in the error itself turns
// the next failure notification into the diagnostic artifact — no extra
// manual step needed. These tests prove the full raw text (not a
// truncated/excerpted snippet) is present in every throw site's message,
// regardless of which specific anchor failed to match.

test('parseEnigooTicketText: no-location-marker error includes the FULL raw extracted text (diagnostic-on-failure)', () => {
  const rawText = 'Letní hudební festival 15.08.2026 19:00';
  assert.throws(() => parseEnigooTicketText(rawText), (err) => {
    assert.ok(err.message.includes(rawText), 'error message should include the full raw extracted text');
    return true;
  });
});

test('parseEnigooTicketText: no-date/time-pattern error includes the FULL raw extracted text, not just an excerpt (diagnostic-on-failure)', () => {
  const rawText = 'Event\nnot a date\nVenue Cena/price: 100 Kč\nsome trailing OCR noise';
  assert.throws(() => parseEnigooTicketText(rawText), (err) => {
    assert.ok(err.message.includes(rawText), 'error message should include the full raw extracted text');
    return true;
  });
});

test('parseEnigooTicketText: hour-out-of-range error includes the FULL raw extracted text (diagnostic-on-failure)', () => {
  const rawText = 'Event 15.08.2026 25:00 Venue Cena/price: 100 Kč';
  assert.throws(() => parseEnigooTicketText(rawText), (err) => {
    assert.ok(err.message.includes(rawText), 'error message should include the full raw extracted text');
    assert.ok(err.message.includes('Hour out of range'), 'error message should still name the specific problem');
    return true;
  });
});

test('parseEnigooTicketText: minute-out-of-range error includes the FULL raw extracted text (diagnostic-on-failure)', () => {
  const rawText = 'Event 15.08.2026 19:75 Venue Cena/price: 100 Kč';
  assert.throws(() => parseEnigooTicketText(rawText), (err) => {
    assert.ok(err.message.includes(rawText), 'error message should include the full raw extracted text');
    assert.ok(err.message.includes('Minute out of range'), 'error message should still name the specific problem');
    return true;
  });
});

// --- resolveTicketingPortal --------------------------------------------------
//
// Mirrors resolveIcsCalendarId's per-sender lookup convention (src/05-action-
// ics-import.js) — case-insensitive match against each configured portal's
// identifyingEmail, first match in list order wins, no match returns null.

test('resolveTicketingPortal: matching sender (bare address) returns that portal entry', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  assert.deepEqual(resolveTicketingPortal('no-reply@enigoo.cz', portals), portals[0]);
});

test('resolveTicketingPortal: matching sender with a display name ("Name <email>") still resolves', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  assert.deepEqual(resolveTicketingPortal('Enigoo <no-reply@enigoo.cz>', portals), portals[0]);
});

test('resolveTicketingPortal: case-insensitive match', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  assert.deepEqual(resolveTicketingPortal('NO-REPLY@ENIGOO.CZ', portals), portals[0]);
});

test('resolveTicketingPortal: no matching sender returns null', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  assert.equal(resolveTicketingPortal('someone-else@example.com', portals), null);
});

test('resolveTicketingPortal: empty/null portals list returns null, never throws', () => {
  assert.doesNotThrow(() => {
    assert.equal(resolveTicketingPortal('no-reply@enigoo.cz', []), null);
    assert.equal(resolveTicketingPortal('no-reply@enigoo.cz', null), null);
  });
});

test('resolveTicketingPortal: multiple entries, first match in list order wins', () => {
  const portals = [
    { identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_FIRST', insertPdfIntoEvent: true },
    { identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_SECOND', insertPdfIntoEvent: false },
  ];
  assert.equal(resolveTicketingPortal('no-reply@enigoo.cz', portals).calendarId, 'CAL_FIRST');
});

// --- resolveTicketingCalendarId (live-test-driven, quick-260731-tix round 6:
// a genuine MISSING-FALLBACK bug, a distinct category from the prior three
// parsing-related rounds) --------------------------------------------------
//
// The owner hit `TypeError: Cannot read properties of null (reading
// 'getTimeZone')` live: the shipped enigoo.cz portal entry's default
// `calendarId` is `null` (the safe, unconfigured default, same as the other
// two actions' optional overrides), but this action originally read
// `portal.calendarId` DIRECTLY at every Calendar API call site, with no
// null-safety at all -- `CalendarApp.getCalendarById(null)` returns `null`,
// so `.getTimeZone()` threw. Both the ICS action (resolveIcsCalendarId) and
// the booking.com action (resolveBookingCalendarId) already fall back to
// CONFIG.calendarId (the global default) when their own action/portal-level
// override is null -- this is the same established two-tier resolution
// pattern, applied here for the first time. Mirrors
// test/calendar-routing.test.js's own resolveBookingCalendarId test suite
// shape exactly.

test('resolveTicketingCalendarId: portal override set -> used', () => {
  assert.equal(resolveTicketingCalendarId({ calendarId: 'PORTAL_CAL' }, 'DEFAULT_CAL'), 'PORTAL_CAL');
});

test('resolveTicketingCalendarId: portal override unset (null) -> global default used (the real live bug this fixes)', () => {
  assert.equal(resolveTicketingCalendarId({ calendarId: null }, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('resolveTicketingCalendarId: portal override key missing entirely -> global default used', () => {
  assert.equal(resolveTicketingCalendarId({}, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('resolveTicketingCalendarId: shipped default enigoo.cz portal entry (calendarId: null) resolves to the global default, reproducing the real live bug scenario', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const shippedPortal = TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals[0];
  assert.equal(shippedPortal.calendarId, null);
  assert.equal(resolveTicketingCalendarId(shippedPortal, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

// --- addMinutesToWallClockComponents (default-duration-when-no-end-time) ---
//
// The enigoo.cz ticket has no explicit end time, only a start
// (15.08.2026 19:00) — a sensible fixed default duration (2 hours) is added
// to compute the event's end. Documented as a per-portal-parser concern: a
// future portal whose PDF DOES include an end time should use it instead of
// this default.

test('addMinutesToWallClockComponents: 19:00 + 120 minutes (2-hour default) -> 21:00, same day', () => {
  assert.deepEqual(
    addMinutesToWallClockComponents({ year: 2026, month: 7, day: 15, hour: 19, minute: 0 }, 120),
    { year: 2026, month: 7, day: 15, hour: 21, minute: 0 }
  );
});

test('addMinutesToWallClockComponents: day rollover (23:30 + 120 minutes -> next day 01:30)', () => {
  assert.deepEqual(
    addMinutesToWallClockComponents({ year: 2026, month: 7, day: 15, hour: 23, minute: 30 }, 120),
    { year: 2026, month: 7, day: 16, hour: 1, minute: 30 }
  );
});

test('addMinutesToWallClockComponents: month rollover (31 Jan + a day-plus of minutes -> 1 Feb)', () => {
  assert.deepEqual(
    addMinutesToWallClockComponents({ year: 2026, month: 0, day: 31, hour: 23, minute: 0 }, 120),
    { year: 2026, month: 1, day: 1, hour: 1, minute: 0 }
  );
});

// --- formatWallClockComponentsIso -------------------------------------------

test('formatWallClockComponentsIso: formats zero-padded, no trailing Z or timezone offset', () => {
  assert.equal(
    formatWallClockComponentsIso({ year: 2026, month: 7, day: 15, hour: 19, minute: 0 }),
    '2026-08-15T19:00:00'
  );
});

test('formatWallClockComponentsIso: single-digit month/day/hour/minute are zero-padded', () => {
  assert.equal(
    formatWallClockComponentsIso({ year: 2026, month: 0, day: 1, hour: 1, minute: 5 }),
    '2026-01-01T01:05:00'
  );
});

// --- resolveTicketProcessingJobs (live-test-driven, quick-260731-tix round
// 8, CORRECTED in round 9) --------------------------------------------------
//
// The owner hit a real live bug: processing the enigoo.cz email ONE time
// created TWO calendar events for the same purchase. Round 8 claimed the
// real enigoo.cz email attaches its 2-ticket purchase as TWO SEPARATE PDF
// FILES and "fixed" this by processing only the FIRST qualifying PDF
// attachment per message. That claim was WRONG -- round 9 independently
// re-parsed the REAL raw .eml's actual MIME structure and confirmed exactly
// ONE application/pdf part (one PDF, 2 internal pages, not two file
// attachments). The "select only the first attachment" restriction has been
// REMOVED: it never addressed the real bug (there was only ever one
// attachment for this email) and would have silently dropped every
// attachment after the first for a legitimate FUTURE portal that emails
// multiple DIFFERENT purchases as separate PDFs in one message. These tests
// now prove resolveTicketProcessingJobs produces ONE JOB PER QUALIFYING PDF
// ATTACHMENT (its original, correct behavior) -- the actual duplicate-event
// protection for the SAME purchase comes entirely from the DEDUP SAFETY NET
// (ticketIdentifier tag + findTicketEventByIdentifier, kept unchanged from
// round 8, tested separately above), not from restricting which attachments
// get processed here.

function fakeAttachment(name, contentType) {
  return {
    getName: function () {
      return name;
    },
    getContentType: function () {
      return contentType || '';
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

test('isTicketPdfAttachment: matches by .pdf name extension (case-insensitive)', () => {
  assert.equal(isTicketPdfAttachment(fakeAttachment('tickets-3676-.PDF', 'application/octet-stream')), true);
});

test('isTicketPdfAttachment: matches by application/pdf content-type even without a .pdf name', () => {
  assert.equal(isTicketPdfAttachment(fakeAttachment('attachment.bin', 'application/pdf')), true);
});

test('isTicketPdfAttachment: a non-PDF attachment (neither name nor content-type) does not match', () => {
  assert.equal(isTicketPdfAttachment(fakeAttachment('image.png', 'image/png')), false);
});

test('findTicketPdfAttachments: returns every qualifying PDF attachment on a message, in source order', () => {
  const pdf1 = fakeAttachment('tickets-3676-a.pdf', 'application/pdf');
  const pdf2 = fakeAttachment('tickets-3676-b.pdf', 'application/pdf');
  const image = fakeAttachment('logo.png', 'image/png');
  const message = fakeMessage('no-reply@enigoo.cz', [image, pdf1, pdf2]);
  assert.deepEqual(findTicketPdfAttachments(message), [pdf1, pdf2]);
});

test('resolveTicketProcessingJobs: the REAL enigoo.cz scenario (a message with exactly ONE qualifying PDF attachment) produces exactly one job', () => {
  const pdf = fakeAttachment('tickets-3676-.pdf', 'application/pdf');
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  const message = fakeMessage('no-reply@enigoo.cz', [pdf]);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].attachment, pdf);
  assert.equal(jobs[0].portal, portals[0]);
});

test('resolveTicketProcessingJobs: a message with TWO qualifying PDF attachments (a hypothetical legitimate multi-purchase email) produces TWO jobs, not one -- corrected in round 9, this is no longer silently collapsed to the first attachment only', () => {
  const pdf1 = fakeAttachment('tickets-concert-a.pdf', 'application/pdf');
  const pdf2 = fakeAttachment('tickets-concert-b.pdf', 'application/pdf');
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  const message = fakeMessage('no-reply@enigoo.cz', [pdf1, pdf2]);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 2);
  assert.equal(jobs[0].attachment, pdf1);
  assert.equal(jobs[0].portal, portals[0]);
  assert.equal(jobs[1].attachment, pdf2);
  assert.equal(jobs[1].portal, portals[0]);
});

test('resolveTicketProcessingJobs: multiple messages on a thread each contribute jobs for their own qualifying attachments', () => {
  const pdf1a = fakeAttachment('tickets-a-1.pdf', 'application/pdf');
  const pdf1b = fakeAttachment('tickets-a-2.pdf', 'application/pdf');
  const pdf2 = fakeAttachment('tickets-b.pdf', 'application/pdf');
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message1 = fakeMessage('no-reply@enigoo.cz', [pdf1a, pdf1b]);
  const message2 = fakeMessage('no-reply@enigoo.cz', [pdf2]);

  const jobs = resolveTicketProcessingJobs([message1, message2], portals);

  assert.equal(jobs.length, 3);
  assert.equal(jobs[0].attachment, pdf1a);
  assert.equal(jobs[1].attachment, pdf1b);
  assert.equal(jobs[2].attachment, pdf2);
});

test('resolveTicketProcessingJobs: a message whose sender does not resolve to any configured portal contributes no job', () => {
  const pdf = fakeAttachment('tickets.pdf', 'application/pdf');
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage('someone-else@example.com', [pdf]);

  assert.deepEqual(resolveTicketProcessingJobs([message], portals), []);
});

test('resolveTicketProcessingJobs: a message that resolves to a portal but carries no qualifying PDF attachment contributes no job', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage('no-reply@enigoo.cz', [fakeAttachment('image.png', 'image/png')]);

  assert.deepEqual(resolveTicketProcessingJobs([message], portals), []);
});

test('resolveTicketProcessingJobs: empty/null messages list returns an empty jobs array, never throws', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  assert.deepEqual(resolveTicketProcessingJobs([], portals), []);
  assert.deepEqual(resolveTicketProcessingJobs(null, portals), []);
});

// --- parseKinoArtTicketText (quick-260731-kar: the SECOND, BODY-SOURCED
// processing mode) ------------------------------------------------------------
//
// Kino Art (kinoart.cz, a Czech cinema) is a genuinely different processing
// mode than enigoo.cz: its event data (movie, venue, date/time) comes
// entirely from the email BODY (message.getPlainBody()), never a PDF.
// Fixture below (HTML stripped to plain text, modeled on the real shape
// with the customer name/movie title/order number replaced by fictional
// equivalents) is a 2-seat, 1-purchase confirmation -- same "one event per
// purchase" rule as every other portal. The date format (`7. 8. 2026
// 17:45` -- no leading zeros, dot-space separated) is DIFFERENT from
// enigoo.cz's zero-padded no-space `15.08.2026`, hence its own distinct
// regex.

const REAL_KINO_ART_BODY_TEXT =
  'Potvrzení objednávky Potvrzení objednávky č. 900142 Informace Zákazník: Jan Novák Úhrada: Comgate Doručení: Elektronicky Prodejní místa Kino Art Položky Název Místo Datum a čas Umístění Cena Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek 4 / 2 150 Kč (1x Plná cena) Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek 4 / 1 150 Kč (1x Plná cena) Cena celkem 300 Kč Tento email není vstupenka.';

test('parseKinoArtTicketText: extracts the correct movie name/venue/date-time/ticketIdentifier from the real body text', () => {
  assert.deepEqual(parseKinoArtTicketText(REAL_KINO_ART_BODY_TEXT), {
    eventName: 'Tajný ostrov',
    location: 'Cihlářská - Malý sál',
    year: 2026,
    month: 7,
    day: 7,
    hour: 17,
    minute: 45,
    ticketIdentifier: '900142',
  });
});

test('parseKinoArtTicketText: the date-format regex correctly handles "7. 8. 2026 17:45" -- day 7, month 7 zero-indexed (August), year 2026, hour 17, minute 45', () => {
  const parsed = parseKinoArtTicketText(REAL_KINO_ART_BODY_TEXT);
  assert.equal(parsed.day, 7);
  assert.equal(parsed.month, 7);
  assert.equal(parsed.year, 2026);
  assert.equal(parsed.hour, 17);
  assert.equal(parsed.minute, 45);
});

test('parseKinoArtTicketText: text with no date/time pattern anywhere throws a controlled error', () => {
  assert.throws(() => parseKinoArtTicketText('Potvrzení objednávky č. 900142 no date here Cena Movie Cihlářská - Malý sál'));
});

test('parseKinoArtTicketText: text with no recognizable venue string throws a controlled error', () => {
  assert.throws(() => parseKinoArtTicketText('Potvrzení objednávky č. 900142 Cena Movie A Different Venue 7. 8. 2026 17:45'));
});

test('parseKinoArtTicketText: ticketIdentifier is null (never throws) when the order confirmation number is not present', () => {
  const noOrderNumberText = 'Cena Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek';
  const parsed = parseKinoArtTicketText(noOrderNumberText);
  assert.equal(parsed.ticketIdentifier, null);
  assert.equal(parsed.eventName, 'Tajný ostrov');
});

test('parseKinoArtTicketText: no-date/time-pattern error includes the FULL raw extracted text (diagnostic-on-failure, same convention as parseEnigooTicketText)', () => {
  const rawText = 'Potvrzení objednávky č. 900142 no date here Cena Movie Cihlářská - Malý sál';
  assert.throws(
    () => parseKinoArtTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes(rawText), 'error message should include the full raw extracted text');
      return true;
    }
  );
});

// REAL_KINO_ART_BODY_TEXT_WITH_BULLET_MARKERS (live-test-driven,
// quick-260731-kar round 3): the REAL message.getPlainBody() output Gmail
// actually produced for this email, observed live -- prefixes each data
// ROW with a literal "* " bullet-list marker (Gmail's real HTML-list-item
// plain-text rendering), which the hand-decoded REAL_KINO_ART_BODY_TEXT
// fixture above (a naive tag-strip approximation, not Gmail's actual
// conversion) did NOT reproduce -- same category of surprise as the
// enigoo.cz Body.getText() paragraph-separator lesson. Only the ROW START
// (immediately after the "Cena" column header, and again before the
// second repeated row) carries the marker -- the venue string
// ("Cihlářská - Malý sál", mid-row) does NOT, confirmed against the real
// observed value the owner/coordinator reported.
const REAL_KINO_ART_BODY_TEXT_WITH_BULLET_MARKERS =
  'Potvrzení objednávky Potvrzení objednávky č. 900142 Informace Zákazník: Jan Novák Úhrada: Comgate Doručení: Elektronicky Prodejní místa Kino Art Položky Název Místo Datum a čas Umístění Cena * Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek 4 / 2 150 Kč (1x Plná cena) * Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek 4 / 1 150 Kč (1x Plná cena) Cena celkem 300 Kč Tento email není vstupenka.';

test('parseKinoArtTicketText: strips a leading "* " bullet-list marker Gmail\'s real getPlainBody() rendering adds to the event name row', () => {
  const parsed = parseKinoArtTicketText(REAL_KINO_ART_BODY_TEXT_WITH_BULLET_MARKERS);
  assert.equal(parsed.eventName, 'Tajný ostrov');
});

test('parseKinoArtTicketText: the bullet-marker strip does not affect the venue/date-time/ticketIdentifier fields', () => {
  const parsed = parseKinoArtTicketText(REAL_KINO_ART_BODY_TEXT_WITH_BULLET_MARKERS);
  assert.equal(parsed.location, 'Cihlářská - Malý sál');
  assert.equal(parsed.day, 7);
  assert.equal(parsed.month, 7);
  assert.equal(parsed.year, 2026);
  assert.equal(parsed.hour, 17);
  assert.equal(parsed.minute, 45);
  assert.equal(parsed.ticketIdentifier, '900142');
});

// TICKET IDENTIFIER ROW-BOUNDARY MARKER (live-test-driven, quick-260731-kar
// round 4): a real live Kino Art attachment rename with insertPdfIntoEvent
// on came out as "Tajný ostrov - 2026-08-07 - null.pdf" -- the real
// order number (900142) failed to extract. Traced precisely: (1) the
// filename-builder call sites in BOTH processTicketPdfAttachment and
// processTicketFromMessageBody already pass parsedTicket.ticketIdentifier
// directly -- no threading/scoping bug (confirmed by direct code
// inspection, ruling out the coordinator's initial "variable scoping"
// hypothesis). (2) Independently re-decoded the real raw .eml's HTML
// (D:\download\KinoArtBrno.eml) around the order-number text and found the
// headline ("Potvrzení objednávky") and the order number ("č. 900142") sit
// in TWO SEPARATE <tr> table rows, not one contiguous block -- the exact
// same row-boundary structure round 3 already proved Gmail's real
// getPlainBody() renders with an inserted "* " marker (see the event-name
// bullet-marker fix above). The original ticketIdentifierMatch regex
// required "Potvrzení objednávky č." to be one CONTIGUOUS phrase separated
// only by whitespace -- an inserted "* " marker at this row boundary would
// have broken that literal match entirely, silently falling back to
// ticketIdentifier: null (this anchor is deliberately optional/never-
// throwing, so the parse otherwise succeeded with everything else intact,
// exactly matching the observed symptom). Fixed by tolerating the SAME
// noise (whitespace and/or a literal asterisk) between "objednávky" and
// "č." that round 3 already proved is real, observed Gmail rendering
// behavior -- not a new, unverified guess.
test('parseKinoArtTicketText: extracts the ticketIdentifier even when Gmail\'s real getPlainBody() rendering inserts a "* " row-boundary marker between "Potvrzení objednávky" and "č."', () => {
  const textWithMarkerBeforeOrderNumber =
    'Potvrzení objednávky Potvrzení objednávky * č. 900142 Informace Zákazník: Jan Novák Úhrada: Comgate Doručení: Elektronicky Prodejní místa Kino Art Položky Název Místo Datum a čas Umístění Cena * Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek 4 / 2 150 Kč (1x Plná cena) Cena celkem 150 Kč Tento email není vstupenka.';
  const parsed = parseKinoArtTicketText(textWithMarkerBeforeOrderNumber);
  assert.equal(parsed.ticketIdentifier, '900142');
  assert.equal(parsed.eventName, 'Tajný ostrov');
});

test('parseKinoArtTicketText: ticketIdentifier extraction is still correct with no marker present at all (regression guard for the original, unmarked fixture)', () => {
  const parsed = parseKinoArtTicketText(REAL_KINO_ART_BODY_TEXT);
  assert.equal(parsed.ticketIdentifier, '900142');
});

// --- findKinoArtTicketPdfAttachment -------------------------------------------
//
// Kino Art sends TWO PDF attachments per confirmation email: Vstupenky.pdf
// (the real ticket, one page per seat) and Doklad.pdf (a separate receipt/
// invoice, NOT ticket data -- the email body itself says "Tento email není
// vstupenka" confirming Vstupenky.pdf is the authoritative ticket).

test('findKinoArtTicketPdfAttachment: selects Vstupenky.pdf and excludes Doklad.pdf from the real filenames', () => {
  const vstupenky = fakeAttachment('Vstupenky.pdf', 'application/pdf');
  const doklad = fakeAttachment('Doklad.pdf', 'application/pdf');
  const message = fakeMessage('rezervace@kinoart.cz', [doklad, vstupenky]);

  assert.equal(findKinoArtTicketPdfAttachment(message), vstupenky);
});

test('findKinoArtTicketPdfAttachment: returns null when no attachment matches (e.g. only Doklad.pdf present)', () => {
  const doklad = fakeAttachment('Doklad.pdf', 'application/pdf');
  const message = fakeMessage('rezervace@kinoart.cz', [doklad]);

  assert.equal(findKinoArtTicketPdfAttachment(message), null);
});

test('findKinoArtTicketPdfAttachment: returns null when the message carries no PDF attachments at all', () => {
  const message = fakeMessage('rezervace@kinoart.cz', []);
  assert.equal(findKinoArtTicketPdfAttachment(message), null);
});

// --- buildTicketAttachmentFilename (quick-260731-kar: the cross-portal
// ATTACHMENT-RENAMING CONVENTION, retrofitted onto enigoo.cz too) -------------
//
// "{event name} - {YYYY-MM-DD} - {ticket identifier}.pdf", ISO-style date
// for consistent Drive-folder sorting. Since a Calendar event attachment's
// displayed title is derived from the file's name at attach time, this
// rename also directly determines what shows up on the calendar event.

test('buildTicketAttachmentFilename: produces the correct filename for the real enigoo.cz fixture', () => {
  assert.equal(
    buildTicketAttachmentFilename('Letní hudební festival', { year: 2026, month: 7, day: 15 }, '24601'),
    'Letní hudební festival - 2026-08-15 - 24601.pdf'
  );
});

test('buildTicketAttachmentFilename: produces the correct filename for the real Kino Art fixture', () => {
  assert.equal(
    buildTicketAttachmentFilename('Tajný ostrov', { year: 2026, month: 7, day: 7 }, '900142'),
    'Tajný ostrov - 2026-08-07 - 900142.pdf'
  );
});

test('buildTicketAttachmentFilename: zero-pads single-digit month/day in the ISO date', () => {
  assert.equal(buildTicketAttachmentFilename('Event', { year: 2026, month: 0, day: 1 }, '1'), 'Event - 2026-01-01 - 1.pdf');
});

test('buildTicketAttachmentFilename: accepts a full parsed-ticket object (with extra hour/minute/ticketIdentifier fields) unchanged', () => {
  const parsedTicket = { eventName: 'Letní hudební festival', location: 'Somewhere', year: 2026, month: 7, day: 15, hour: 19, minute: 0, ticketIdentifier: '24601' };
  assert.equal(buildTicketAttachmentFilename(parsedTicket.eventName, parsedTicket, parsedTicket.ticketIdentifier), 'Letní hudební festival - 2026-08-15 - 24601.pdf');
});

test('buildTicketAttachmentFilename: sanitizes filesystem-unsafe characters in the event name', () => {
  assert.equal(buildTicketAttachmentFilename('Movie: Part 2/3', { year: 2026, month: 0, day: 1 }, '1'), 'Movie- Part 2-3 - 2026-01-01 - 1.pdf');
});

// DEFENSIVE NULL-HANDLING (live-test-driven, quick-260731-kar round 4): a
// real live Kino Art attachment rename came out as
// "Tajný ostrov - 2026-08-07 - null.pdf" -- the literal 4-character
// string "null", from naive string concatenation coercing a JS `null`
// ticketIdentifier argument. This is a defensive secondary measure (the
// PRIMARY fix is making sure a portal's REAL ticketIdentifier is correctly
// extracted/threaded -- see parseKinoArtTicketText's own round-4 fix
// below) so that ANY portal whose parser genuinely cannot extract a
// ticketIdentifier (a documented, never-throwing possibility per every
// portal parser's own class-level doc) never embeds the literal word
// "null" in a filename -- the segment is omitted entirely instead.
test('buildTicketAttachmentFilename: omits the ticketIdentifier segment entirely (not the literal string "null") when ticketIdentifier is null', () => {
  assert.equal(buildTicketAttachmentFilename('Tajný ostrov', { year: 2026, month: 7, day: 7 }, null), 'Tajný ostrov - 2026-08-07.pdf');
});

test('buildTicketAttachmentFilename: omits the ticketIdentifier segment when ticketIdentifier is undefined too', () => {
  assert.equal(buildTicketAttachmentFilename('Event', { year: 2026, month: 0, day: 1 }, undefined), 'Event - 2026-01-01.pdf');
});

test('buildTicketAttachmentFilename: still includes the ticketIdentifier segment for a real, truthy value (regression guard)', () => {
  assert.equal(buildTicketAttachmentFilename('Tajný ostrov', { year: 2026, month: 7, day: 7 }, '900142'), 'Tajný ostrov - 2026-08-07 - 900142.pdf');
});

// --- resolveTicketProcessingJobs: BODY-SOURCED mode (Kino Art) ---------------
//
// A body-sourced portal produces EXACTLY ONE job per matching message,
// tagged `mode: 'body'`, with NO attachment/PDF requirement at all (unlike
// the PDF-sourced mode's `mode: 'pdf'` jobs, one per qualifying attachment).

test('resolveTicketProcessingJobs: a Kino Art message (body-sourced portal) produces exactly one "body"-mode job, with no PDF attachment requirement', () => {
  const portals = [{ identifyingEmail: 'rezervace@kinoart.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage('rezervace@kinoart.cz', []);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].portal, portals[0]);
});

test('resolveTicketProcessingJobs: a Kino Art message still produces exactly one "body"-mode job even when it also carries PDF attachments (Vstupenky.pdf/Doklad.pdf) -- never one job per attachment for a body-sourced portal', () => {
  const portals = [{ identifyingEmail: 'rezervace@kinoart.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  const message = fakeMessage('rezervace@kinoart.cz', [
    fakeAttachment('Vstupenky.pdf', 'application/pdf'),
    fakeAttachment('Doklad.pdf', 'application/pdf'),
  ]);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
});

test('resolveTicketProcessingJobs: an enigoo.cz message still produces "pdf"-mode jobs, unaffected by the new body-sourced mode', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const pdf = fakeAttachment('tickets.pdf', 'application/pdf');
  const message = fakeMessage('no-reply@enigoo.cz', [pdf]);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'pdf');
  assert.equal(jobs[0].attachment, pdf);
});

test('resolveTicketProcessingJobs: both an enigoo.cz message and a Kino Art message on the same thread each contribute their own correctly-moded job', () => {
  const portals = [
    { identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false },
    { identifyingEmail: 'rezervace@kinoart.cz', calendarId: 'CAL_B', insertPdfIntoEvent: false },
  ];
  const enigooMessage = fakeMessage('no-reply@enigoo.cz', [fakeAttachment('tickets.pdf', 'application/pdf')]);
  const kinoArtMessage = fakeMessage('rezervace@kinoart.cz', []);

  const jobs = resolveTicketProcessingJobs([enigooMessage, kinoArtMessage], portals);

  assert.equal(jobs.length, 2);
  assert.equal(jobs[0].mode, 'pdf');
  assert.equal(jobs[1].mode, 'body');
});

// --- parseTicketmasterCzTicketText (quick-260816-ocw: the THIRD supported
// portal, BODY-SOURCED like Kino Art) -----------------------------------------
//
// Ticketmaster CZ (noreply@ticketmaster.cz) confirmation emails carry their
// event data entirely in the plain-text body's "YOUR ORDER DETAILS" section
// (message.getPlainBody() — the SAME body-sourced processing mode Kino Art
// already proved), never a PDF, even though the real email DOES carry an
// eTicket.pdf attachment (D-01: deliberately never touched, no PDF-finder,
// no Drive/OCR pipeline of any kind for this portal). The date format
// ("Sunday 15 November 2026 at 20:00" -- full weekday, no-leading-zero day,
// full English month NAME, 4-digit year, the literal word "at", 24h HH:MM)
// is genuinely different from both enigoo.cz's zero-padded no-space
// "15.08.2026" and Kino Art's dot-space "7. 8. 2026", hence its own regex
// plus a local month-name-to-number lookup table (D-04) -- confirmed by grep
// that no such helper exists anywhere else in this codebase. No stable
// per-ticket or per-order confirmation number exists anywhere in the real
// body (checked directly against the real sample), so `ticketIdentifier` is
// `null` by design, a documented v1 limitation, never a substitute/invented
// value (D-05) -- the DEDUP SAFETY NET therefore cannot protect this portal
// against reprocessing duplicates, an accepted trade-off same as this file's
// other two parsers' own optional-anchor treatment.
//
// Fixture below reproduces the real sample's exact paragraph layout (one
// field per paragraph, separated by non-breaking-space-ONLY lines, per the
// PLAN's <real_sample_shape>) with a FICTIONAL event name/venue substituted
// for the real purchase (this file's established fixture convention) while
// preserving the real character classes that matter to the regex --
// non-ASCII Czech letters and an en-dash in the event name. The real date
// string ("Sunday 15 November 2026 at 20:00") and the real
// "Ticket Quantity: 2" line are kept verbatim.

const REAL_TICKETMASTER_CZ_BODY_TEXT = [
  'YOUR ORDER DETAILS',
  '\u00A0',
  'Léto v podzámčí – Koncertní večer',
  '\u00A0',
  'Sál Radost',
  '\u00A0',
  'Sunday 15 November 2026 at 20:00',
  '\u00A0',
  'Ticket Quantity: 2',
].join('\n');

test('parseTicketmasterCzTicketText: parses the real fixture into the full expected shape in one assertion, November proving the zero-indexed month', () => {
  assert.deepEqual(parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_BODY_TEXT), {
    eventName: 'Léto v podzámčí – Koncertní večer',
    location: 'Sál Radost',
    year: 2026,
    month: 10,
    day: 15,
    hour: 20,
    minute: 0,
    ticketIdentifier: null,
    ticketQuantity: 2,
    description: 'Léto v podzámčí – Koncertní večer\n\nSál Radost\n\nSunday 15 November 2026 at 20:00\n\nTicket Quantity: 2',
  });
});

test('parseTicketmasterCzTicketText: description reproduces the order-details block as eventName/location/date-time-line/quantity paragraphs, weekday name included (round 2)', () => {
  const parsed = parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_BODY_TEXT);
  assert.equal(
    parsed.description,
    'Léto v podzámčí – Koncertní večer\n\nSál Radost\n\nSunday 15 November 2026 at 20:00\n\nTicket Quantity: 2'
  );
});

test('parseTicketmasterCzTicketText: a non-numeric Ticket Quantity value does not throw -- ticketQuantity stays null and its line is omitted from description (round 2)', () => {
  const rawText = REAL_TICKETMASTER_CZ_BODY_TEXT.replace('Ticket Quantity: 2', 'Ticket Quantity: many');
  const parsed = parseTicketmasterCzTicketText(rawText);

  assert.equal(parsed.ticketQuantity, null);
  assert.equal(parsed.description, 'Léto v podzámčí – Koncertní večer\n\nSál Radost\n\nSunday 15 November 2026 at 20:00');
});

// TICKETMASTER_CZ_MONTH_TEST_NAMES -- all twelve full English month names in
// calendar order, used below to prove TICKETMASTER_CZ_MONTH_NAMES resolves
// every entry to the correct zero-indexed month number (D-04), not just
// November.
const TICKETMASTER_CZ_MONTH_TEST_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

function buildMinimalTicketmasterCzBody(monthName) {
  return ['YOUR ORDER DETAILS', 'Some Event', 'Some Venue', 'Sunday 1 ' + monthName + ' 2026 at 12:00', 'Ticket Quantity: 1'].join('\n');
}

test('parseTicketmasterCzTicketText: every one of the twelve English month names resolves to the correct zero-indexed month number (D-04)', () => {
  TICKETMASTER_CZ_MONTH_TEST_NAMES.forEach(function (monthName, index) {
    const parsed = parseTicketmasterCzTicketText(buildMinimalTicketmasterCzBody(monthName));
    assert.equal(parsed.month, index, monthName + ' should resolve to zero-indexed month ' + index);
  });
});

test('parseTicketmasterCzTicketText: the non-breaking-space-only separator lines in the real fixture do not disrupt eventName/location extraction', () => {
  const parsed = parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_BODY_TEXT);
  assert.equal(parsed.eventName, 'Léto v podzámčí – Koncertní večer');
  assert.equal(parsed.location, 'Sál Radost');
});

test('parseTicketmasterCzTicketText: a CRLF-joined variant of the same fixture parses identically -- separator-agnostic, same convention as the other two parsers', () => {
  const crlfText = REAL_TICKETMASTER_CZ_BODY_TEXT.replace(/\n/g, '\r\n');
  assert.deepEqual(parseTicketmasterCzTicketText(crlfText), parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_BODY_TEXT));
});

test('parseTicketmasterCzTicketText: ticketIdentifier is null and the parse still succeeds -- proves "documented limitation", not "parse failed" (D-05)', () => {
  const parsed = parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_BODY_TEXT);
  assert.equal(parsed.ticketIdentifier, null);
  assert.ok(parsed.eventName);
});

test('parseTicketmasterCzTicketText: Ticket Quantity 1, 2 and 5 all yield the SAME event identity (eventName/location/date-time/ticketIdentifier) -- quantity is never an event multiplier (D-06, amended round 2)', () => {
  const quantity2Result = parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_BODY_TEXT);
  const quantity1Text = REAL_TICKETMASTER_CZ_BODY_TEXT.replace('Ticket Quantity: 2', 'Ticket Quantity: 1');
  const quantity5Text = REAL_TICKETMASTER_CZ_BODY_TEXT.replace('Ticket Quantity: 2', 'Ticket Quantity: 5');
  const quantity1Result = parseTicketmasterCzTicketText(quantity1Text);
  const quantity5Result = parseTicketmasterCzTicketText(quantity5Text);

  // The event ITSELF (what determines whether this is one calendar event or
  // several) is identical across all three quantities -- D-06's core
  // guarantee, unchanged: one purchase always yields one event, regardless
  // of how many tickets it contains.
  ['eventName', 'location', 'year', 'month', 'day', 'hour', 'minute', 'ticketIdentifier'].forEach(function (field) {
    assert.equal(quantity1Result[field], quantity2Result[field], field + ' must not vary with quantity');
    assert.equal(quantity5Result[field], quantity2Result[field], field + ' must not vary with quantity');
  });

  // ROUND 2 AMENDMENT: unlike round 1, ticketQuantity and description NOW
  // deliberately vary with the input quantity, since description surfaces
  // the real "Ticket Quantity: N" line back to the owner -- this is new,
  // intentional data flow, not a regression of D-06's one-event guarantee.
  assert.equal(quantity1Result.ticketQuantity, 1);
  assert.equal(quantity2Result.ticketQuantity, 2);
  assert.equal(quantity5Result.ticketQuantity, 5);
  assert.ok(quantity1Result.description.includes('Ticket Quantity: 1'));
  assert.ok(quantity5Result.description.includes('Ticket Quantity: 5'));
});

test('parseTicketmasterCzTicketText: a body missing the "YOUR ORDER DETAILS" marker throws a controlled error carrying the FULL raw text', () => {
  const rawText = ['Léto v podzámčí – Koncertní večer', 'Sál Radost', 'Sunday 15 November 2026 at 20:00', 'Ticket Quantity: 2'].join('\n');
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: a body missing the "Ticket Quantity:" marker throws a controlled error carrying the FULL raw text', () => {
  const rawText = ['YOUR ORDER DETAILS', 'Léto v podzámčí – Koncertní večer', 'Sál Radost', 'Sunday 15 November 2026 at 20:00'].join('\n');
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: an order-details region with no recognizable date/time throws a controlled error carrying the FULL raw text', () => {
  const rawText = ['YOUR ORDER DETAILS', 'Léto v podzámčí – Koncertní večer', 'Sál Radost', 'no date here', 'Ticket Quantity: 2'].join('\n');
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: an unrecognized month name throws a controlled error carrying the FULL raw text', () => {
  const rawText = ['YOUR ORDER DETAILS', 'Léto v podzámčí – Koncertní večer', 'Sál Radost', 'Sunday 15 Smarch 2026 at 20:00', 'Ticket Quantity: 2'].join(
    '\n'
  );
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: hour out of range throws with "Hour out of range" and the FULL raw text', () => {
  const rawText = ['YOUR ORDER DETAILS', 'Léto v podzámčí – Koncertní večer', 'Sál Radost', 'Sunday 15 November 2026 at 25:00', 'Ticket Quantity: 2'].join(
    '\n'
  );
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes('Hour out of range'), 'error message should mention "Hour out of range"');
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: minute out of range throws with "Minute out of range" and the FULL raw text', () => {
  const rawText = ['YOUR ORDER DETAILS', 'Léto v podzámčí – Koncertní večer', 'Sál Radost', 'Sunday 15 November 2026 at 20:75', 'Ticket Quantity: 2'].join(
    '\n'
  );
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes('Minute out of range'), 'error message should mention "Minute out of range"');
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: an order-details region carrying only ONE non-empty paragraph before the date/time throws a controlled error carrying the FULL raw text (event name and venue not separable)', () => {
  const rawText = ['YOUR ORDER DETAILS', 'Léto v podzámčí – Koncertní večer', 'Sunday 15 November 2026 at 20:00', 'Ticket Quantity: 2'].join('\n');
  assert.throws(
    () => parseTicketmasterCzTicketText(rawText),
    (err) => {
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseTicketmasterCzTicketText: TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL is wired to the exported parser (D-01)', () => {
  assert.strictEqual(TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL['noreply@ticketmaster.cz'], parseTicketmasterCzTicketText);
});

test('Ticketmaster CZ has NO OCR/PDF-TEXT-parsing pipeline: TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL has no noreply@ticketmaster.cz key (D-01, still true after round 2)', () => {
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL, 'noreply@ticketmaster.cz'), false);
});

// ROUND 2 AMENDMENT to D-01: the owner asked the real eTicket.pdf to be
// attached to the created calendar event, the same OPTIONAL
// insertPdfIntoEvent attachment path Kino Art already proved -- this is
// NOT the OCR/PDF-TEXT pipeline (that stays untouched, see the test above),
// it is the separate find-the-PDF-and-attach-it-as-is mechanism the
// body-sourced mode already supports via
// TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL.
test('Ticketmaster CZ IS registered in TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL, wired to findTicketmasterCzTicketPdfAttachment (round 2)', () => {
  assert.strictEqual(TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL['noreply@ticketmaster.cz'], findTicketmasterCzTicketPdfAttachment);
});

// --- findTicketmasterCzTicketPdfAttachment (round 2) --------------------------
//
// The real Ticketmaster CZ email carries exactly one PDF attachment,
// filename "eTicket.pdf" (Content-Disposition: attachment, application/pdf).

test('findTicketmasterCzTicketPdfAttachment: selects the real eTicket.pdf attachment', () => {
  const eTicket = fakeAttachment('eTicket.pdf', 'application/pdf');
  const message = fakeMessage('noreply@ticketmaster.cz', [eTicket]);

  assert.equal(findTicketmasterCzTicketPdfAttachment(message), eTicket);
});

test('findTicketmasterCzTicketPdfAttachment: returns null when no attachment name matches "eTicket"', () => {
  const other = fakeAttachment('Invoice.pdf', 'application/pdf');
  const message = fakeMessage('noreply@ticketmaster.cz', [other]);

  assert.equal(findTicketmasterCzTicketPdfAttachment(message), null);
});

test('findTicketmasterCzTicketPdfAttachment: returns null when the message carries no PDF attachments at all', () => {
  const message = fakeMessage('noreply@ticketmaster.cz', []);
  assert.equal(findTicketmasterCzTicketPdfAttachment(message), null);
});

test('resolveTicketProcessingJobs: a noreply@ticketmaster.cz message with no attachments yields exactly one "body"-mode job (D-01)', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage('noreply@ticketmaster.cz', []);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].portal, portals[0]);
});

test('resolveTicketProcessingJobs: the SAME Ticketmaster CZ message carrying the real eTicket.pdf attachment STILL yields exactly one "body"-mode job -- never one job per attachment, never a "pdf"-mode job (D-01)', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage('noreply@ticketmaster.cz', [fakeAttachment('eTicket.pdf', 'application/pdf')]);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
});

test('TICKETING_PORTALS_ACTION_CONFIG: the shipped default seeds a THIRD entry for noreply@ticketmaster.cz, and resolveTicketingCalendarId falls back to the passed global default for it (D-02)', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const thirdPortal = TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals[2];

  assert.deepEqual(thirdPortal, { identifyingEmail: 'noreply@ticketmaster.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.equal(resolveTicketingCalendarId(thirdPortal, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

test('TICKETING_PORTALS_ACTION_CONFIG: regression guard -- entries 0 and 1 are still the unchanged enigoo.cz and Kino Art entries, array length is 3', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const portals = TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals;

  assert.equal(portals.length, 3);
  assert.deepEqual(portals[0], { identifyingEmail: 'no-reply@enigoo.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.deepEqual(portals[1], { identifyingEmail: 'rezervace@kinoart.cz', calendarId: null, insertPdfIntoEvent: false });
});

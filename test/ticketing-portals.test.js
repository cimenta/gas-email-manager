'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {
  parseEnigooTicketText,
  parseKinoArtTicketText,
  parseTicketmasterCzTicketText,
  parseEntradioTicketText,
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
  TICKETING_PORTALS_ACTION,
  processTicketFromMessageBody,
  // ROUND 2 (debug/entradio-portal-not-supported): the ticket-file + QR-code
  // Calendar-attachment pipeline.
  extractEntradioTicketCodes,
  findEntradioTicketDownloadUrl,
  buildEntradioQrCodeUrl,
  isEntradioTicketFileResponseAcceptable,
  isEntradioQrCodeResponseAcceptable,
  entradioFileExtensionForMimeType,
  buildEntradioTicketAttachmentFilename,
  buildEntradioQrCodeFilename,
  fetchEntradioAttachments,
  TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL,
  buildTicketCalendarEventResource,
  // debug/ticketmaster-cz-order-confirm: the body-content admission gate.
  ticketmasterCzTextHasOrderDetails,
  TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL,
  // quick-260921-gj0: the FIFTH portal, Fever (hello@feverup.com).
  parseFeverTicketText,
  feverTextHasPurchaseDetails,
  findFeverTicketPdfAttachment,
  feverResolveEventYear,
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

// `plainBody` was added by debug/ticketmaster-cz-order-confirm: before that
// fix this fake could not express body content AT ALL, which is precisely why
// the envelope-only admission gate went unnoticed (see that session's "why no
// gate caught it"). It defaults to '' rather than being required, so a portal
// with no registered body-content detector is unaffected.
//
// `receivedDate` (quick-260921-gj0) is an OPTIONAL fourth parameter exposed
// as `getDate()`, added because processTicketFromMessageBody now passes the
// message's own received date through to the registered body parser (D-07,
// Fever's year-inference source). Defaults to `undefined` -- a caller that
// never invokes processTicketFromMessageBody through this fake is unaffected.
//
// `subject` (quick-260921-gj0 round 2, D-27) is an OPTIONAL fifth parameter
// exposed as `getSubject()`, added because processTicketFromMessageBody now
// ALSO passes the message's own subject through to the registered body
// parser (D-25, Fever's primary event-name source). Same reasoning as
// `receivedDate` above: defaults to `undefined`, inert for any caller that
// never invokes processTicketFromMessageBody through this fake.
function fakeMessage(fromHeader, attachments, plainBody, receivedDate, subject) {
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
    getDate: function () {
      return receivedDate;
    },
    getSubject: function () {
      return subject;
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

// AMENDED by debug/ticketmaster-cz-order-confirm: both tests below now pass the
// real order-details body. They previously passed NO body at all and still
// asserted a job, which is exactly the envelope-only contract that produced the
// bug. The assertions are unchanged -- a Ticketmaster CZ message carrying real
// order details must still yield exactly one body-mode job.

test('resolveTicketProcessingJobs: a noreply@ticketmaster.cz message with no attachments yields exactly one "body"-mode job (D-01)', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage('noreply@ticketmaster.cz', [], REAL_TICKETMASTER_CZ_BODY_TEXT);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].portal, portals[0]);
});

test('resolveTicketProcessingJobs: the SAME Ticketmaster CZ message carrying the real eTicket.pdf attachment STILL yields exactly one "body"-mode job -- never one job per attachment, never a "pdf"-mode job (D-01)', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const message = fakeMessage(
    'noreply@ticketmaster.cz',
    [fakeAttachment('eTicket.pdf', 'application/pdf')],
    REAL_TICKETMASTER_CZ_BODY_TEXT
  );

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

test('TICKETING_PORTALS_ACTION_CONFIG: regression guard -- entries 0 and 1 are still the unchanged enigoo.cz and Kino Art entries, array length is 5 (updated quick-260921-gj0)', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const portals = TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals;

  assert.equal(portals.length, 5);
  assert.deepEqual(portals[0], { identifyingEmail: 'no-reply@enigoo.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.deepEqual(portals[1], { identifyingEmail: 'rezervace@kinoart.cz', calendarId: null, insertPdfIntoEvent: false });
});

// --- THE BODY-CONTENT ADMISSION GATE (debug/ticketmaster-cz-order-confirm) ---
//
// THE BUG: a real, legitimate Ticketmaster CZ PURCHASE CONFIRMATION (the email
// sent BEFORE the tickets themselves, which arrive in a separate follow-up)
// was admitted into the body-mode processing path on SENDER ALONE and then
// threw "Unrecognized Ticketmaster CZ ticket text: \"YOUR ORDER DETAILS\"
// marker not found", surfacing to the owner as an action-failure notification.
//
// Ticketmaster CZ sends at least TWO distinct templates from the one address
// noreply@ticketmaster.cz -- this confirmation, and the ticket-delivery email
// that already works. Config matches ADDRESSES, not templates, so no config
// change can separate them; only body content can. resolveTicketProcessingJobs
// and appliesTo both decided "this is a ticket" from the envelope and never
// called getPlainBody(), leaving the parser's marker THROW as the only way the
// system could say "this is not an order-details block".
//
// THE FIDELITY PROPERTY (established by debug/teamio-non-invite-error):
// ticketmasterCzTextHasOrderDetails is built on the SAME marker constant and
// the SAME NBSP normalization parseTicketmasterCzTicketText itself uses, so the
// detector returns false for EXACTLY the bodies the parser would have thrown
// the marker error on. It is not an independent heuristic and cannot drift.
//
// WHY NOT A POSITIVE "tickets follow separately" MARKER: the obvious candidate
// sentence is HARD-WRAPPED mid-phrase in the real body ("...with your\ntickets
// attached."), so a literal substring match finds NOTHING -- it would have
// looked like a fix while changing nothing. Matching it would need a
// whitespace-tolerant regex over MARKETING COPY, the exact drift-prone second
// marker the fidelity property exists to forbid.
//
// The fixture below is the REAL owner-supplied .eml's text/plain part
// (quoted-printable -> UTF-8), reproduced line for line INCLUDING its
// NBSP-only separator lines, its trailing-space lines and its mid-sentence
// wrap -- those are exactly the details a hand-approximated fixture would
// smooth away. Only the buyer's NAME and PHONE are redacted (PII, per
// push-public.bat's scrub convention); there is no buyer address in the body,
// and the only address present is Ticketmaster's own public registered office.
// Full fidelity matters here specifically because the whole claim under test is
// "the marker appears NOWHERE in the body" -- a truncated fixture could not
// support it.

const REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT = [
  '',
  '',
  '',
  'YOU GOT THE TICKETS!',
  '\u00A0',
  'ORDER NUMBER: 2387845',
  '',
  '',
  'ČESKÝ MEJDAN S IMPULSEM 2026',
  '\u00A0',
  '\u00A0',
  'O2 arena',
  '\u00A0',
  'Saturday 17 October 2026 at 18:00',
  '\u00A0',
  '2 x tickets',
  '\u00A0',
  'Promoter: Bestsport, a.s., IČ: 24214795',
  '\u00A0',
  '',
  '',
  '',
  '',
  'View Tickets [https://www.ticketmaster.cz/user/orders]',
  '',
  '',
  ' This email cannot be used for event entry.',
  ' ',
  '',
  'TICKET DELIVERY',
  '\u00A0',
  'eTicket',
  '\u00A0',
  'Delivery fee: 0 Kč',
  '',
  'Delivery information: You will receive a separate email with your',
  'tickets attached. You can also download your tickets at any time',
  'via your account, save them to your mobile wallet, or view them in',
  'the Ticketmaster mobile app for Android or iOS.',
  '',
  '',
  'ORDER SUMMARY',
  ' 2 ticket(s)',
  ' ',
  'Level',
  ' Section',
  ' Row',
  ' Seat',
  ' ',
  '4. POSCHODÍ',
  ' 421',
  ' 14',
  ' 13',
  ' ',
  '      ',
  '',
  '',
  '',
  '',
  '',
  'Level',
  ' Section',
  ' Row',
  ' Seat',
  ' ',
  '4. POSCHODÍ',
  ' 421',
  ' 14',
  ' 12',
  ' ',
  '      ',
  '',
  '',
  '',
  '',
  '',
  '',
  '      ',
  '',
  '',
  'PAYMENT SUMMARY',
  '',
  'Total',
  ' 2\u00A0010,00 Kč',
  ' ',
  'Payment Method',
  ' Mastercard',
  '',
  '',
  ' ',
  '      ',
  '',
  'YOUR CONTACT DETAILS',
  '\u00A0',
  'Jan Novák\u00A0\u00A0\u00A0',
  '\u00A0',
  'Phone: 420700000000',
  '',
  '',
  ' YOUR PHONE IS YOUR TICKET',
  '',
  ' Download the Ticketmaster App',
  ' ',
  ' Sign in to view your ticket(s)',
  ' ',
  ' For entry to the event, scan your ticket directly from your phone',
  ' ',
  '  ',
  '',
  ' ',
  '',
  'LET\'S CONNECT',
  '',
  '     ',
  'Need Help? Contact our Fan Support Team [https://help.ticketmaster.cz/hc/en-us]',
  '',
  '',
  'Ticketmaster [https://www.ticketmaster.cz/]\u00A0\u00A0\u00A0\u00A0\u00A0 Privacy Policy [https://privacy.ticketmaster.cz/en/privacy-policy]\u00A0\u00A0\u00A0\u00A0\u00A0 My Account [https://my.ticketmaster.cz/settings]',
  '',
  '© 2026 Ticketmaster Česká republika, a.s.',
  'Jungmannova 26/15, 110 00 Praha 1, Česká republika',
  'All rights reserved.',
].join('\n');

// FIXTURE INTEGRITY -- pins the three properties every test below depends on.
// If a future edit smooths the fixture, these fail FIRST and name the reason,
// rather than letting the regression tests pass for the wrong reason.

test('fixture integrity: the real order-confirmation body contains NO "YOUR ORDER DETAILS" and NO "Ticket Quantity:" marker anywhere, but DOES carry "ORDER SUMMARY"', () => {
  assert.equal(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('YOUR ORDER DETAILS'), -1);
  assert.equal(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('Ticket Quantity:'), -1);
  assert.notEqual(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('ORDER SUMMARY'), -1);
});

test('fixture integrity: the "tickets follow separately" sentence is hard-wrapped mid-phrase, so NO literal one-line marker can match it -- this is why the detector shares the parser marker instead', () => {
  assert.equal(
    REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('You will receive a separate email with your tickets attached'),
    -1
  );
  assert.notEqual(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('with your\ntickets attached'), -1);
});

test('fixture integrity: the confirmation carries a WELL-FORMED date/time the parser pattern matches -- so a date-shaped detector would wrongly admit it, and only the region marker separates the two templates', () => {
  const dateTimePattern = /(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\s+at\s+(\d{1,2}):(\d{2})/;
  const match = dateTimePattern.exec(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.replace(/\u00A0/g, ' '));

  assert.notEqual(match, null);
  assert.equal(match[0], '17 October 2026 at 18:00');
});

test('fixture integrity: the buyer name and phone number are redacted -- no real PII reaches the public repo', () => {
  assert.equal(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('Radek'), -1);
  assert.equal(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT.indexOf('420704145475'), -1);
});

// THE REGRESSION -- the two gates, against the SHIPPED default config rather
// than a hand-built portal list. Before the fix each admitted the message and
// the run then threw; both must now refuse it outright. These assert the
// action DOES NOT APPLY, not merely that it "does not throw": a run that still
// claimed and labelled the thread would satisfy a crash-freedom oracle while
// leaving the real defect in place.

test('THE REGRESSION -- resolveTicketProcessingJobs: the real Ticketmaster CZ ORDER CONFIRMATION (tickets follow separately) yields ZERO jobs against the shipped defaults', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const message = fakeMessage(
    'Ticketmaster <noreply@ticketmaster.cz>',
    [],
    REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT
  );

  assert.deepEqual(resolveTicketProcessingJobs([message], TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals), []);
});

test('THE REGRESSION -- TICKETING_PORTALS_ACTION.appliesTo: returns false for the real order-confirmation thread, so the action never claims or labels it', () => {
  const message = fakeMessage(
    'Ticketmaster <noreply@ticketmaster.cz>',
    [],
    REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT
  );
  const thread = {
    getMessages: function () {
      return [message];
    },
  };

  assert.equal(TICKETING_PORTALS_ACTION.appliesTo(thread), false);
});

test('THE REGRESSION: the order confirmation is refused even when it carries an eTicket.pdf attachment -- the attachment never re-opens the gate', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  const message = fakeMessage(
    'noreply@ticketmaster.cz',
    [fakeAttachment('eTicket.pdf', 'application/pdf')],
    REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT
  );

  assert.deepEqual(resolveTicketProcessingJobs([message], portals), []);
});

// THE HAPPY PATH -- the working ticket-details email must be entirely
// unaffected. This is the load-bearing counterweight: the fix adds a gate that
// could in principle refuse a genuine ticket email.

test('THE HAPPY PATH: the working Ticketmaster CZ ticket-details body still yields exactly one "body"-mode job against the shipped defaults', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const message = fakeMessage('Ticketmaster <noreply@ticketmaster.cz>', [], REAL_TICKETMASTER_CZ_BODY_TEXT);

  const jobs = resolveTicketProcessingJobs([message], TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].portal.identifyingEmail, 'noreply@ticketmaster.cz');
});

test('THE HAPPY PATH: appliesTo still claims a thread carrying the working ticket-details body', () => {
  const thread = {
    getMessages: function () {
      return [fakeMessage('noreply@ticketmaster.cz', [], REAL_TICKETMASTER_CZ_BODY_TEXT)];
    },
  };

  assert.equal(TICKETING_PORTALS_ACTION.appliesTo(thread), true);
});

test('THE HAPPY PATH: a thread mixing the confirmation AND the ticket-details email yields exactly ONE job -- for the ticket-details message only', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const confirmation = fakeMessage('noreply@ticketmaster.cz', [], REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT);
  const ticketDetails = fakeMessage('noreply@ticketmaster.cz', [], REAL_TICKETMASTER_CZ_BODY_TEXT);

  const jobs = resolveTicketProcessingJobs([confirmation, ticketDetails], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].message, ticketDetails);
});

// STRUCTURE PRESENT BUT UNPARSEABLE MUST STILL THROW. The gate only converts
// "no order-details region at all" into a silent skip; a genuinely malformed
// region is still admitted as a job so the parser can report it loudly.

test('a MALFORMED Ticketmaster CZ body (has "YOUR ORDER DETAILS" but no "Ticket Quantity:") is still ADMITTED as a job -- the gate never suppresses a real parse failure', () => {
  const portals = [{ identifyingEmail: 'noreply@ticketmaster.cz', calendarId: 'CAL_A', insertPdfIntoEvent: false }];
  const malformed = REAL_TICKETMASTER_CZ_BODY_TEXT.replace('Ticket Quantity: 2', 'Tickets: 2');
  const message = fakeMessage('noreply@ticketmaster.cz', [], malformed);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
});

test('a MALFORMED Ticketmaster CZ body still THROWS from the parser, unchanged -- structure-present-but-unparseable stays a loud error', () => {
  const malformed = REAL_TICKETMASTER_CZ_BODY_TEXT.replace('Ticket Quantity: 2', 'Tickets: 2');

  assert.throws(() => parseTicketmasterCzTicketText(malformed), /"Ticket Quantity:" marker not found/);
});

test('the parser STILL throws the marker error when called directly on the order confirmation -- the throw is a defensive invariant, now unreachable from the production path', () => {
  assert.throws(
    () => parseTicketmasterCzTicketText(REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT),
    /"YOUR ORDER DETAILS" marker not found/
  );
});

// THE FIDELITY PROPERTY -- the detector must return false for EXACTLY the
// bodies the parser throws the marker error on. Proved by construction over
// both real fixtures plus the boundary cases, so detector and parser cannot
// drift apart in a future edit.

test('FIDELITY PROPERTY: for every probe body, ticketmasterCzTextHasOrderDetails() === false IFF the parser throws the "YOUR ORDER DETAILS" marker error', () => {
  const probes = [
    REAL_TICKETMASTER_CZ_ORDER_CONFIRMATION_BODY_TEXT,
    REAL_TICKETMASTER_CZ_BODY_TEXT,
    REAL_TICKETMASTER_CZ_BODY_TEXT.replace('Ticket Quantity: 2', 'Tickets: 2'),
    REAL_TICKETMASTER_CZ_BODY_TEXT.replace('YOUR ORDER DETAILS', 'YOUR\u00A0ORDER DETAILS'),
    '',
    'nothing resembling a ticket',
    'YOUR ORDER DETAILS',
  ];

  probes.forEach(function (probe) {
    let threwMarkerError = false;
    try {
      parseTicketmasterCzTicketText(probe);
    } catch (error) {
      threwMarkerError = /"YOUR ORDER DETAILS" marker not found/.test(error.message);
    }

    assert.equal(ticketmasterCzTextHasOrderDetails(probe), !threwMarkerError, 'fidelity broken for probe: ' + JSON.stringify(probe.slice(0, 60)));
  });
});

test('ticketmasterCzTextHasOrderDetails: shares the parser NBSP normalization -- a marker split by non-breaking spaces is still detected', () => {
  assert.equal(ticketmasterCzTextHasOrderDetails('YOUR\u00A0ORDER\u00A0DETAILS\n\u00A0\nX\n\u00A0\nTicket Quantity: 1'), true);
});

test('ticketmasterCzTextHasOrderDetails: never throws on empty/null/undefined/non-string input, always returns a literal boolean', () => {
  [undefined, null, '', 0, 42, {}, []].forEach(function (input) {
    const result = ticketmasterCzTextHasOrderDetails(input);
    assert.equal(typeof result, 'boolean');
  });

  assert.equal(ticketmasterCzTextHasOrderDetails(undefined), false);
  assert.equal(ticketmasterCzTextHasOrderDetails(null), false);
  assert.equal(ticketmasterCzTextHasOrderDetails(''), false);
});

test('ticketmasterCzTextHasOrderDetails: the marker is matched case-SENSITIVELY and as a whole -- a lowercased or partial heading is not an order-details block', () => {
  assert.equal(ticketmasterCzTextHasOrderDetails('your order details'), false);
  assert.equal(ticketmasterCzTextHasOrderDetails('YOUR ORDER'), false);
  assert.equal(ticketmasterCzTextHasOrderDetails('ORDER DETAILS'), false);
  assert.equal(ticketmasterCzTextHasOrderDetails('preamble YOUR ORDER DETAILS trailing'), true);
});

// THE REGISTRY, and its deliberately FAIL-OPEN default. This gate diverges
// from debug/teamio-non-invite-error's FAIL-CLOSED one on purpose: the owner
// scoped this fix to Ticketmaster CZ only, and Kino Art / Entradio variants are
// unverified. Fail-closed would have silently disabled both. The two tests
// below pin that as a DECISION rather than leaving it an untested default.

test('TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL: noreply@ticketmaster.cz is wired to the exported detector', () => {
  assert.strictEqual(TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL['noreply@ticketmaster.cz'], ticketmasterCzTextHasOrderDetails);
});

test('TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL: has NO entry for Kino Art or Entradio -- the fix is scoped to Ticketmaster CZ (and, since quick-260921-gj0, Fever) only, the Kino Art/Entradio variants are unverified', () => {
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL, 'rezervace@kinoart.cz'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL, 'no-reply@app.entradio.cz'), false);
  assert.deepEqual(Object.keys(TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL), ['noreply@ticketmaster.cz', 'hello@feverup.com']);
});

test('FAIL-OPEN on a missing detector: a Kino Art message with an EMPTY body still yields exactly one "body"-mode job -- unchanged from before the fix', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const message = fakeMessage('rezervace@kinoart.cz', [], '');

  const jobs = resolveTicketProcessingJobs([message], TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].portal.identifyingEmail, 'rezervace@kinoart.cz');
});

test('FAIL-OPEN on a missing detector: an Entradio message with an EMPTY body still yields exactly one "body"-mode job -- unchanged from before the fix', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const message = fakeMessage('Kino Metropol <no-reply@app.entradio.cz>', [], '');

  const jobs = resolveTicketProcessingJobs([message], TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].portal.identifyingEmail, 'no-reply@app.entradio.cz');
});

test('FAIL-OPEN on a missing detector: appliesTo still claims Kino Art and Entradio threads with empty bodies', () => {
  ['rezervace@kinoart.cz', 'no-reply@app.entradio.cz'].forEach(function (sender) {
    const thread = {
      getMessages: function () {
        return [fakeMessage(sender, [], '')];
      },
    };

    assert.equal(TICKETING_PORTALS_ACTION.appliesTo(thread), true, 'fail-open broken for ' + sender);
  });
});

test('the gate never consults the body for a PDF/OCR-sourced portal: an enigoo.cz message with an empty body still yields its "pdf"-mode job', () => {
  const portals = [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'CAL_A', insertPdfIntoEvent: true }];
  const message = fakeMessage('no-reply@enigoo.cz', [fakeAttachment('tickets.pdf', 'application/pdf')], '');

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'pdf');
});

// --- parseEntradioTicketText (debug/entradio-portal-not-supported: the ------
// --- FOURTH supported portal, the THIRD body-sourced one) -------------------
//
// THE BUG THIS PORTAL EXISTS TO FIX: an Entradio confirmation
// (no-reply@app.entradio.cz) produced NO calendar event AND NO error. The
// sender matched no TICKETING_PORTALS entry, so resolveTicketingPortal
// returned null, resolveTicketProcessingJobs `continue`d past it without
// emitting a job, and appliesTo returned false -- a silent skip, which is the
// correct behaviour for an unknown sender. The defect was that this sender
// was unknown, not that any matching or parsing logic was broken. The
// resolveTicketProcessingJobs test at the bottom of this block is the direct
// reproduction: before the fix it yielded 0 jobs, after it yields exactly 1.
//
// Entradio is a white-label PLATFORM, not a venue -- the real sample came
// from Kino Metropol Olomouc, but the sender address is shared by every venue
// using the system, so one portal entry covers all of them and every anchor
// is on Entradio's own template structure rather than any venue's name.
//
// The fixture below is the REAL decoded .eml's text/plain part (owner-
// supplied sample, quoted-printable -> UTF-8), reproduced line for line
// INCLUDING its whitespace-only separator lines and the trailing spaces after
// the empty "Poschodí"/"Sleva" labels -- those are exactly the details a
// hand-approximated fixture would smooth away, and the Kino Art round-3
// incident is this codebase's standing proof that such smoothing hides real
// bugs. Two redactions only, neither touching a parse anchor: the buyer's own
// name/e-mail/phone are fictionalized (mandatory -- push-public.bat hard-
// aborts if the real address ever reaches src/ or test/), and the four
// per-recipient SendGrid tracking URLs are collapsed to one placeholder.
// Verified directly: the redacted fixture and the untouched real body parse
// to byte-identical results.

const REAL_ENTRADIO_BODY_TEXT = [
  ' Děkujeme za Vaši objednávku. Vaše platba byla úspěšně zaplacena.',
  '',
  'Zobrazit objednávku ve webovém prohlížeči ( https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-TRACKING-TOKEN )',
  '',
  '*Děkujeme za Vaši objednávku*',
  '',
  'Vaše platba byla úspěšně zaplacena.',
  '',
  'Číslo objednávky',
  '*2354152*',
  '',
  'Při vstupu na událost se prokážete vstupenkami, které jsou součástí tohoto e-mailu. Vstupenky není třeba tisknout, pokud je předložíte v mobilním telefonu.',
  'Pokud jste při nákupu využili slevu, prosíme, prokažte nárok na slevu při vstupu na událost. V případě, že nárok na slevu prokázat nedokážete, žádáme, vzniklý rozdíl uhraďte v pokladně.',
  '',
  'Událost',
  '---------------------------------------------------------------',
  '',
  '',
  '*ČERNO, VÍR*',
  '',
  '27. 9. 2026, 17:30',
  '',
  ' ',
  '',
  'Brána na událost se otevírá v 27. 9. 2026, od 17:00 hodin.',
  '',
  '',
  '',
  'Místo konání',
  '------------',
  '',
  '*Kino Metropol, Kino Metropol*',
  '',
  ' Sokolská 572/25, 77900 Olomouc, Česká republika',
  '',
  'Vstupenky',
  '---------',
  '',
  '',
  'TM5X59GM • 230 Kč',
  '',
  ' ',
  '',
  '',
  '',
  'Poschodí ',
  '',
  ' ',
  '',
  'Sekce vlevo',
  '',
  ' ',
  '',
  'Řada 3',
  '',
  ' ',
  '',
  'Místo 19',
  '',
  '',
  '',
  'Sleva ',
  '',
  '',
  '2ZKN9JXVT • 230 Kč',
  '',
  ' ',
  '',
  '',
  '',
  'Poschodí ',
  '',
  ' ',
  '',
  'Sekce vlevo',
  '',
  ' ',
  '',
  'Řada 3',
  '',
  ' ',
  '',
  'Místo 18',
  '',
  '',
  '',
  'Sleva ',
  '',
  '',
  'Platba',
  '------',
  '',
  'Číslo platby',
  '',
  '*2134174*',
  '',
  'Celkem',
  '',
  '460 Kč',
  '',
  'Uhrazeno',
  '',
  '14. 9. 2026, 12:28',
  '',
  'STÁHNOUT VSTUPENKY ( https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-TRACKING-TOKEN )',
  '',
  'STÁHNOUT JAKO DÁREK ( https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-TRACKING-TOKEN )',
  '',
  ' ',
  '',
  'Jméno *Jan Novák*',
  '',
  ' ',
  '',
  'E-mail *jan.novak@example.com*',
  '',
  ' ',
  '',
  'Telefon *+420 000 000 000*',
  '',
  'Společnost',
  '',
  '',
  'Fakturační adresa',
  '',
  '',
  '',
  'Adresa doručení',
  '',
  '',
  '',
  '',
  '',
  '',
  'Odesílatel e-mailu',
  '',
  'DCI KINO Olomouc s.r.o., Sokolská 572/25, 779 00 Olomouc, Česká republika, IČ: 29391709, DIČ: CZ29391709, DIČ: CZ29391709.',
  'Provozovna: Kino Metropol +420 722 955 466 info@kinometropol.cz',
  '',
  ' Zásady ochrany osobních údajů ( https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-TRACKING-TOKEN ) ',
  '',
  ' ',
  '',
  'Organizator události: DCI KINO Olomouc s.r.o.',
  '',
  'Tento e-mail byl vygenerovaný automaticky. Žádáme, neodpovídejte na něj.',
  '',
  '♥ VYUŽÍVÁME SYSTÉM ENTRADIO',
].join('\r\n');

const EXPECTED_ENTRADIO_DESCRIPTION = [
  'ČERNO, VÍR',
  '',
  'Kino Metropol, Sokolská 572/25, 77900 Olomouc, Česká republika',
  '',
  '27. 9. 2026, 17:30',
  '',
  'Číslo objednávky: 2354152',
  '',
  'Vstupenky (2):',
  'TM5X59GM • Sekce vlevo, Řada 3, Místo 19',
  '2ZKN9JXVT • Sekce vlevo, Řada 3, Místo 18',
].join('\n');

test('parseEntradioTicketText: parses the real fixture into the full expected shape in one assertion, September proving the zero-indexed month', () => {
  assert.deepEqual(parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT), {
    eventName: 'ČERNO, VÍR',
    location: 'Kino Metropol, Sokolská 572/25, 77900 Olomouc, Česká republika',
    year: 2026,
    month: 8,
    day: 27,
    hour: 17,
    minute: 30,
    ticketIdentifier: '2354152',
    ticketQuantity: 2,
    // ROUND 2 (debug/entradio-portal-not-supported): the RAW per-seat codes,
    // added so fetchEntradioAttachments can build one QR-code URL per seat.
    // A TIGHTENING of this assertion, not a loosening -- the expected shape
    // now pins one more field than it did before round 2.
    ticketCodes: ['TM5X59GM', '2ZKN9JXVT'],
    description: EXPECTED_ENTRADIO_DESCRIPTION,
  });
});

// THE decisive near-miss test. The "Událost" section carries TWO date/time-
// shaped strings for the same day: the real 17:30 start, and a gate-opening
// line at 17:00 ("Brána na událost se otevírá v 27. 9. 2026, od 17:00
// hodin."). Picking the right one must not depend on which happens to come
// first in the text, because "correct by ordering luck" and "correct by
// construction" are indistinguishable until the template shifts. The pattern
// requires digits where the gate line has the literal "od ", so it cannot
// match that line at any start offset -- proven twice below: once with the
// real line present, and once with the real line REMOVED, where a
// merely-first-wins parser would silently produce an event starting at 17:00.

test('parseEntradioTicketText: picks the real 17:30 start, never the 17:00 gate-opening time in the same section', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  assert.equal(parsed.hour, 17);
  assert.equal(parsed.minute, 30);
});

test('parseEntradioTicketText: with the real start-time line REMOVED, the gate-opening line is still not matched -- a controlled throw, never a silent 17:00 event', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('27. 9. 2026, 17:30', '');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /no date\/time pattern found after the event name/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseEntradioTicketText: the ticketIdentifier is the ORDER number, never the payment number sitting under a near-identical label in the same body', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  assert.equal(parsed.ticketIdentifier, '2354152');
  assert.notEqual(parsed.ticketIdentifier, '2134174');
});

test('parseEntradioTicketText: a body with no "Číslo objednávky" label still parses -- ticketIdentifier is null and its description line is omitted, never a throw', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('Číslo objednávky\r\n*2354152*', '');
  const parsed = parseEntradioTicketText(rawText);

  assert.equal(parsed.ticketIdentifier, null);
  assert.equal(parsed.eventName, 'ČERNO, VÍR');
  assert.equal(parsed.description.includes('Číslo objednávky:'), false);
});

// ONE EVENT PER ORDER (owner-confirmed, and the same one-event-per-purchase
// design every other portal in this file follows): a 2-seat order yields ONE
// parsed ticket carrying ONE ticketIdentifier. ticketQuantity reports the
// seat count for the description and is NEVER an event multiplier -- proven
// by parsing a 1-seat variant and asserting the event IDENTITY is unchanged.

test('parseEntradioTicketText: a 2-seat order yields ONE parsed ticket with ONE order-scoped identifier, not one per seat', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  assert.equal(parsed.ticketQuantity, 2);
  assert.equal(parsed.ticketIdentifier, '2354152');
});

test('parseEntradioTicketText: a 1-seat variant yields the SAME event identity as the 2-seat order -- only ticketQuantity and description differ', () => {
  // Drops the second seat block entirely, from its ticket code up to the
  // "Platba" section heading that bounds the tickets region.
  const oneSeatText = REAL_ENTRADIO_BODY_TEXT.replace(/2ZKN9JXVT[\s\S]*?(?=Platba\r\n-{3,})/, '');
  const oneSeat = parseEntradioTicketText(oneSeatText);
  const twoSeats = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  assert.equal(oneSeat.eventName, twoSeats.eventName);
  assert.equal(oneSeat.location, twoSeats.location);
  assert.equal(oneSeat.year, twoSeats.year);
  assert.equal(oneSeat.month, twoSeats.month);
  assert.equal(oneSeat.day, twoSeats.day);
  assert.equal(oneSeat.hour, twoSeats.hour);
  assert.equal(oneSeat.minute, twoSeats.minute);
  assert.equal(oneSeat.ticketIdentifier, twoSeats.ticketIdentifier);

  assert.equal(oneSeat.ticketQuantity, 1);
  assert.equal(oneSeat.description.includes('Vstupenky (1):'), true);
  assert.equal(oneSeat.description.includes('2ZKN9JXVT'), false);

  // ROUND 2: ticketCodes tracks the seat count the same way ticketQuantity
  // does -- one QR-code attachment per seat, never one per order.
  assert.deepEqual(oneSeat.ticketCodes, ['TM5X59GM']);
  assert.deepEqual(twoSeats.ticketCodes, ['TM5X59GM', '2ZKN9JXVT']);
});

// Section headings are anchored together with their dashes underline, not by
// a bare indexOf. This matters: the SAME body contains the prose sentence
// "Vstupenky není třeba tisknout..." well before the real "Vstupenky"
// section, and "Vaše platba byla úspěšně zaplacena" before the real "Platba"
// section. A bare indexOf on either word would scope the tickets region to
// the wrong span entirely.

test('parseEntradioTicketText: the earlier prose sentence containing the bare word "Vstupenky" is not mistaken for the tickets section heading', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  // Had the prose sentence won, the region would have started ~20 lines too
  // early and swept in no ticket-code lines at all.
  assert.equal(parsed.ticketQuantity, 2);
  assert.equal(parsed.description.includes('TM5X59GM'), true);
});

test('parseEntradioTicketText: with the whole tickets section removed, the parse still succeeds -- ticketQuantity is null and the seat block is omitted from the description', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace(/Vstupenky\r\n-{3,}[\s\S]*?(?=Platba\r\n-{3,})/, '');
  const parsed = parseEntradioTicketText(rawText);

  assert.equal(parsed.ticketQuantity, null);
  assert.equal(parsed.description.includes('Vstupenky ('), false);
  assert.equal(parsed.eventName, 'ČERNO, VÍR');
  assert.equal(parsed.location, 'Kino Metropol, Sokolská 572/25, 77900 Olomouc, Česká republika');

  // ROUND 2: no seat block means no codes to build QR URLs from -- an EMPTY
  // ARRAY, never null/undefined, so fetchEntradioAttachments can iterate it
  // unconditionally without a shape check.
  assert.deepEqual(parsed.ticketCodes, []);
});

// Seat fields are matched on their OWN line. The real sample leaves
// "Poschodí" and "Sleva" with an empty value, and a `\s+`-based pattern would
// jump the blank lines and capture the NEXT label's value -- silently
// rendering "Poschodí vlevo".

test('parseEntradioTicketText: the empty "Poschodí" label is omitted from the seat line, never filled in with the following field value', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  assert.equal(parsed.description.includes('Poschodí'), false);
  assert.equal(parsed.description.includes('TM5X59GM • Sekce vlevo, Řada 3, Místo 19'), true);
});

test('parseEntradioTicketText: a populated "Poschodí" value IS rendered, in the email\'s own field order', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('Poschodí \r\n\r\n \r\n\r\nSekce vlevo\r\n\r\n \r\n\r\nŘada 3\r\n\r\n \r\n\r\nMísto 19', 'Poschodí 1\r\n\r\n \r\n\r\nSekce vlevo\r\n\r\n \r\n\r\nŘada 3\r\n\r\n \r\n\r\nMísto 19');
  const parsed = parseEntradioTicketText(rawText);

  assert.equal(parsed.description.includes('TM5X59GM • Poschodí 1, Sekce vlevo, Řada 3, Místo 19'), true);
});

// The venue arrives as "<venue>, <hall>" and this tenant named its only hall
// after the venue, so the real value stutters: "Kino Metropol, Kino
// Metropol". De-stuttering is safe rather than speculative -- it can only
// change a value whose segments genuinely repeat.

test('parseEntradioTicketText: the stuttered venue name is collapsed once, then joined to the street address', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  assert.equal(parsed.location, 'Kino Metropol, Sokolská 572/25, 77900 Olomouc, Česká republika');
});

test('parseEntradioTicketText: a normal non-repeating "<venue>, <hall>" pair passes through completely untouched', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('*Kino Metropol, Kino Metropol*', '*Kino Metropol, Velký sál*');
  const parsed = parseEntradioTicketText(rawText);

  assert.equal(parsed.location, 'Kino Metropol, Velký sál, Sokolská 572/25, 77900 Olomouc, Česká republika');
});

test('parseEntradioTicketText: a venue with no address line still yields a usable location rather than a throw', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace(' Sokolská 572/25, 77900 Olomouc, Česká republika', '');
  const parsed = parseEntradioTicketText(rawText);

  assert.equal(parsed.location, 'Kino Metropol');
});

test('parseEntradioTicketText: LF-only and CR-only variants of the same fixture parse identically -- separator-agnostic, same convention as the other three parsers', () => {
  const lfOnly = REAL_ENTRADIO_BODY_TEXT.replace(/\r\n/g, '\n');
  const crOnly = REAL_ENTRADIO_BODY_TEXT.replace(/\r\n/g, '\r');

  assert.deepEqual(parseEntradioTicketText(lfOnly), parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT));
  assert.deepEqual(parseEntradioTicketText(crOnly), parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT));
});

// Controlled throws. Every one must carry the FULL raw text untruncated, per
// this file's diagnostic-on-failure convention -- the owner's failure
// notification email is the next diagnostic artifact.

test('parseEntradioTicketText: a body with no dash-underlined "Událost" section throws a controlled error carrying the FULL raw text', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('Událost\r\n', 'Udalost\r\n');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /no dash-underlined "Událost" section heading found/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseEntradioTicketText: a body with no dash-underlined "Místo konání" section throws a controlled error carrying the FULL raw text', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('Místo konání\r\n', 'Misto konani\r\n');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /no dash-underlined "Místo konání" section heading found/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseEntradioTicketText: a body with no bold event name throws a controlled error carrying the FULL raw text', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('*ČERNO, VÍR*', 'ČERNO, VÍR');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /no bold event name/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseEntradioTicketText: a body with no bold venue name throws a controlled error carrying the FULL raw text', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('*Kino Metropol, Kino Metropol*', 'Kino Metropol');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /no bold venue name/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseEntradioTicketText: hour out of range throws with "Hour out of range" and the FULL raw text', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('27. 9. 2026, 17:30', '27. 9. 2026, 25:30');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /Hour out of range \(0-23\)/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseEntradioTicketText: minute out of range throws with "Minute out of range" and the FULL raw text', () => {
  const rawText = REAL_ENTRADIO_BODY_TEXT.replace('27. 9. 2026, 17:30', '27. 9. 2026, 17:75');

  assert.throws(
    () => parseEntradioTicketText(rawText),
    (err) => {
      assert.match(err.message, /Minute out of range \(0-59\)/);
      assert.ok(err.message.includes(rawText), 'error message should include the full raw text');
      return true;
    }
  );
});

// --- Entradio wiring ---------------------------------------------------------

test('parseEntradioTicketText: TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL is wired to the exported parser', () => {
  assert.strictEqual(TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL['no-reply@app.entradio.cz'], parseEntradioTicketText);
});

test('Entradio has NO OCR/PDF-TEXT-parsing pipeline: TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL has no no-reply@app.entradio.cz key', () => {
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL, 'no-reply@app.entradio.cz'), false);
});

// A DELIBERATE absence, not an omission: an Entradio confirmation carries no
// ticket PDF at all -- its only attachment is the venue's terms and
// conditions (VOP_Metropol.pdf), so a finder registered here could only ever
// attach the wrong document to the owner's calendar.
test('Entradio is deliberately NOT registered in TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL -- its only PDF is the venue terms, never a ticket', () => {
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL, 'no-reply@app.entradio.cz'), false);
});

// THE DIRECT REPRODUCTION of the reported bug, against the SHIPPED default
// config rather than a hand-built portal list: before this fix the sender
// resolved to no portal and the message produced ZERO jobs -- silently, with
// no error anywhere. It must now produce exactly one body-mode job, and the
// terms-and-conditions PDF must not turn it into a pdf-mode job or a second
// job.

test('resolveTicketProcessingJobs: an Entradio message resolves against the SHIPPED default portals and yields exactly one "body"-mode job (was ZERO before this fix)', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const portals = TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals;
  const message = fakeMessage('Kino Metropol <no-reply@app.entradio.cz>', [fakeAttachment('VOP_Metropol.pdf', 'application/pdf')]);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].portal.identifyingEmail, 'no-reply@app.entradio.cz');
});

test('resolveTicketingPortal: the real "Kino Metropol <no-reply@app.entradio.cz>" From header now resolves against the shipped defaults (it returned null before this fix)', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const portal = resolveTicketingPortal('Kino Metropol <no-reply@app.entradio.cz>', TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals);

  assert.notEqual(portal, null);
  assert.equal(portal.identifyingEmail, 'no-reply@app.entradio.cz');
});

// appliesTo is a SECOND, independent gate -- dispatchActions consults it
// before run, and it repeats the portal+registry lookup rather than reusing
// resolveTicketProcessingJobs. A registration that satisfied only one of the
// two would still leave the email silently unprocessed, so both are pinned.
test('TICKETING_PORTALS_ACTION.appliesTo: claims an Entradio thread whose only attachment is the venue terms PDF (returned false before this fix)', () => {
  const message = fakeMessage('Kino Metropol <no-reply@app.entradio.cz>', [fakeAttachment('VOP_Metropol.pdf', 'application/pdf')]);
  const thread = {
    getMessages: function () {
      return [message];
    },
  };

  assert.equal(TICKETING_PORTALS_ACTION.appliesTo(thread), true);
});

test('TICKETING_PORTALS_ACTION.appliesTo: still returns false for an unrelated sender -- the new entry widened nothing else', () => {
  const thread = {
    getMessages: function () {
      return [fakeMessage('someone@unrelated-sender.example', [])];
    },
  };

  assert.equal(TICKETING_PORTALS_ACTION.appliesTo(thread), false);
});

test('TICKETING_PORTALS_ACTION_CONFIG: the shipped default seeds a FOURTH entry for no-reply@app.entradio.cz, and resolveTicketingCalendarId falls back to the passed global default for it', () => {
  const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
  const fourthPortal = TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals[3];

  assert.deepEqual(fourthPortal, { identifyingEmail: 'no-reply@app.entradio.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.equal(resolveTicketingCalendarId(fourthPortal, 'DEFAULT_CAL'), 'DEFAULT_CAL');
});

// ============================================================================
// ROUND 2 (debug/entradio-portal-not-supported): REAL TICKET-FILE AND QR-CODE
// CALENDAR ATTACHMENTS FOR ENTRADIO
// ============================================================================
//
// WHAT CHANGED AND WHY: round 1 shipped the portal registration + body parser,
// which fixed the reported "no event, no error" silent skip. Before live
// verification the owner expanded the scope: an Entradio confirmation carries
// no ticket file in the message at all, so the created event was going to be a
// bare event with nothing to show at the door. Round 2 fetches the real
// artifacts over HTTP -- the ticket file behind the "STÁHNOUT VSTUPENKY" link,
// and one QR-code image PER SEAT from Entradio's own qrcode endpoint -- and
// attaches them to the Calendar event.
//
// OWNER-SETTLED DECISIONS these tests encode (asked and answered, not
// re-derived here):
//   - Every seat's QR code is its OWN Calendar attachment, and it is ALWAYS
//     attempted -- never gated by insertPdfIntoEvent.
//   - The ticket-file download IS gated by insertPdfIntoEvent, the same toggle
//     the other three portals already use.
//   - Both land in the EXISTING shared CONFIG.ticketAttachmentDriveFolderName
//     folder. No new folder.
//   - A total attachment failure NEVER blocks event creation; it only sends a
//     separate notification email.
//
// THIS IS THIS CODEBASE'S FIRST-EVER OUTBOUND HTTP CALL, which is why the
// appsscript.json scope test below exists and why every fetch path is proven
// NON-THROWING rather than merely proven correct on the happy path.

// --- appsscript.json: the new OAuth scope -----------------------------------
//
// UrlFetchApp is unusable without script.external_request, and an Apps Script
// project that calls it without the scope declared fails at RUNTIME, inside the
// trigger, where the owner sees it only as a failed execution. Pinning it in
// the manifest is the only pre-deployment guard available.

function readAppsScriptManifest() {
  return JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'src', 'appsscript.json'), 'utf8'));
}

test('appsscript.json: declares script.external_request -- without it UrlFetchApp fails at runtime inside the trigger (ROUND 2, this codebase\'s first outbound HTTP)', () => {
  const manifest = readAppsScriptManifest();

  assert.equal(manifest.oauthScopes.includes('https://www.googleapis.com/auth/script.external_request'), true);
});

test('appsscript.json: every pre-round-2 scope is STILL declared -- the new scope is additive, it replaces nothing (a dropped scope would silently break another action)', () => {
  const manifest = readAppsScriptManifest();

  [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/script.send_mail',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
  ].forEach(function (scope) {
    assert.equal(manifest.oauthScopes.includes(scope), true, 'missing pre-existing scope: ' + scope);
  });
});

// --- extractEntradioTicketCodes ---------------------------------------------
//
// The RAW codes, as a SIBLING of extractEntradioTicketLines rather than a
// change to it: that function returns pre-formatted description LINES and is
// already covered by its own tests, so repurposing its return shape would have
// meant rewriting working assertions to serve a new caller.

test('extractEntradioTicketCodes: returns the raw per-seat codes from the tickets region, in the email\'s own order', () => {
  const region = [
    'TM5X59GM • 230 Kč',
    'Sekce vlevo',
    'Řada 3',
    'Místo 19',
    '',
    '2ZKN9JXVT • 230 Kč',
    'Sekce vlevo',
    'Řada 3',
    'Místo 18',
  ].join('\r\n');

  assert.deepEqual(extractEntradioTicketCodes(region), ['TM5X59GM', '2ZKN9JXVT']);
});

test('extractEntradioTicketCodes: a region with no ticket-code lines yields [] rather than null -- callers iterate it unconditionally', () => {
  assert.deepEqual(extractEntradioTicketCodes('Platba\r\n------\r\nCelkem 460 Kč'), []);
});

test('extractEntradioTicketCodes: agrees seat-for-seat with extractEntradioTicketLines on the real fixture -- the two extractors can never drift apart silently', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  // Every code must head its own description line. If one extractor's scan
  // changed and the other's did not, the QR attachments would silently stop
  // matching the seats listed in the event description.
  assert.equal(parsed.ticketCodes.length, parsed.ticketQuantity);
  parsed.ticketCodes.forEach(function (code) {
    const headedLines = parsed.description.split('\n').filter(function (line) {
      return line.indexOf(code) === 0;
    });
    assert.equal(headedLines.length, 1, 'code ' + code + ' must head exactly one description line');
  });
});

// --- findEntradioTicketDownloadUrl ------------------------------------------
//
// REAL HTML SHAPE, taken verbatim from the owner's sample .eml (text/html part,
// decoded quoted-printable -> UTF-8, lines 517 and 528). Only the
// per-recipient SendGrid tracking tokens are replaced with placeholders -- the
// tag structure, attribute order, styling and inner text are the real thing,
// because that structure IS what the pattern anchors on.
//
// THE NEAR-MISS THIS FIXTURE EXISTS TO PROVE: the very next button in the same
// email is "STÁHNOUT JAKO DÁREK" (download as a gift) -- an identically shaped
// anchor with a DIFFERENT URL. A pattern anchored on "STÁHNOUT" alone would
// fetch the gift artifact, and neither response validator could catch it: the
// gift link also answers 200 with a non-HTML body. Same class of trap as the
// gate-opening time round 1 found in the plain body.

const REAL_ENTRADIO_TICKETS_ANCHOR =
  '                  <a href="https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-TICKETS-TOKEN" ' +
  'style="background-color:#6A1B9A; border:1px solid #6A1B9A; border-color:#6A1B9A; border-radius:4px; border-width:1px; ' +
  'color:#FFFFFF; display:inline-block; font-size:14px; font-weight:bold; letter-spacing:0px; line-height:normal; ' +
  'padding:10px 16px 10px 16px; text-align:center; text-decoration:none; border-style:solid; ' +
  'font-family:verdana,geneva,sans-serif;" target="_blank">STÁHNOUT VSTUPENKY</a>';

const REAL_ENTRADIO_GIFT_ANCHOR =
  '                  <a href="https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-GIFT-TOKEN" ' +
  'style="background-color:#FFFFFF; border:1px solid #6A1B9A; border-color:#6A1B9A; border-radius:4px; border-width:1px; ' +
  'color:#6A1B9A; display:inline-block; font-size:14px; font-weight:bold; letter-spacing:0px; line-height:normal; ' +
  'padding:10px 16px 10px 16px; text-align:center; text-decoration:none; border-style:solid; ' +
  'font-family:verdana,geneva,sans-serif;" target="_blank">STÁHNOUT JAKO DÁREK</a>';

const REAL_ENTRADIO_HTML_BODY = [
  '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
  '<html><body>',
  '  <table><tr><td>',
  '    <img src="https://app.entradio.cz/qrcode?code=TM5X59GM&size=200" alt="QR" width="200" height="200" />',
  '    <img src="https://app.entradio.cz/qrcode?code=2ZKN9JXVT&size=200" alt="QR" width="200" height="200" />',
  '  </td></tr></table>',
  '  <table><tr><td align="center" style="padding:0;">',
  REAL_ENTRADIO_TICKETS_ANCHOR,
  '  </td></tr>',
  '  <tr><td align="center" style="padding:0;">',
  REAL_ENTRADIO_GIFT_ANCHOR,
  '  </td></tr></table>',
  '</body></html>',
].join('\r\n');

const EXPECTED_ENTRADIO_DOWNLOAD_URL = 'https://u00000000.ct.sendgrid.net/ls/click?upn=u001.EXAMPLE-TICKETS-TOKEN';

test('findEntradioTicketDownloadUrl: extracts the STÁHNOUT VSTUPENKY href from the real HTML body', () => {
  assert.equal(findEntradioTicketDownloadUrl(REAL_ENTRADIO_HTML_BODY), EXPECTED_ENTRADIO_DOWNLOAD_URL);
});

test('findEntradioTicketDownloadUrl: never returns the "STÁHNOUT JAKO DÁREK" gift link that sits 11 lines below it in the same email', () => {
  const url = findEntradioTicketDownloadUrl(REAL_ENTRADIO_HTML_BODY);

  assert.equal(url.includes('GIFT'), false);
  assert.equal(url, EXPECTED_ENTRADIO_DOWNLOAD_URL);
});

test('findEntradioTicketDownloadUrl: the tickets link is selected STRUCTURALLY, not by position -- it still wins when the gift anchor comes FIRST', () => {
  const reordered = [
    '<html><body><table>',
    '  <tr><td>' + REAL_ENTRADIO_GIFT_ANCHOR + '</td></tr>',
    '  <tr><td>' + REAL_ENTRADIO_TICKETS_ANCHOR + '</td></tr>',
    '</table></body></html>',
  ].join('\r\n');

  assert.equal(findEntradioTicketDownloadUrl(reordered), EXPECTED_ENTRADIO_DOWNLOAD_URL);
});

test('findEntradioTicketDownloadUrl: an HTML body with ONLY the gift anchor returns null -- a controlled miss, never the wrong URL', () => {
  const giftOnly = '<html><body><table><tr><td>' + REAL_ENTRADIO_GIFT_ANCHOR + '</td></tr></table></body></html>';

  assert.equal(findEntradioTicketDownloadUrl(giftOnly), null);
});

test('findEntradioTicketDownloadUrl: null, undefined and empty input all return null rather than throwing', () => {
  assert.equal(findEntradioTicketDownloadUrl(null), null);
  assert.equal(findEntradioTicketDownloadUrl(undefined), null);
  assert.equal(findEntradioTicketDownloadUrl(''), null);
});

test('findEntradioTicketDownloadUrl: tolerates arbitrary attributes on EITHER side of href, including none at all', () => {
  const attributesBefore = '<a class="btn" data-x="1" href="https://example.test/t" target="_blank">STÁHNOUT VSTUPENKY</a>';
  const noOtherAttributes = '<a href="https://example.test/t">STÁHNOUT VSTUPENKY</a>';

  assert.equal(findEntradioTicketDownloadUrl(attributesBefore), 'https://example.test/t');
  assert.equal(findEntradioTicketDownloadUrl(noOtherAttributes), 'https://example.test/t');
});

test('findEntradioTicketDownloadUrl: tolerates whitespace/newlines between the tag and its inner text', () => {
  const spaced = '<a href="https://example.test/t" target="_blank">\r\n  STÁHNOUT VSTUPENKY\r\n</a>';

  assert.equal(findEntradioTicketDownloadUrl(spaced), 'https://example.test/t');
});

test('findEntradioTicketDownloadUrl: an HTML-escaped &amp; in the href is decoded -- an un-decoded one would produce a URL that fetches nothing', () => {
  const escaped = '<a href="https://example.test/t?a=1&amp;b=2" target="_blank">STÁHNOUT VSTUPENKY</a>';

  assert.equal(findEntradioTicketDownloadUrl(escaped), 'https://example.test/t?a=1&b=2');
});

// --- buildEntradioQrCodeUrl -------------------------------------------------
//
// THE URL TEMPLATE IS NOT INVENTED: the real email's HTML renders each seat's
// QR inline as <img src="https://app.entradio.cz/qrcode?code=TM5X59GM&size=200">
// for exactly the two codes the body parser extracts. This function reproduces
// Entradio's own endpoint.

test('buildEntradioQrCodeUrl: reproduces the real per-seat QR endpoint observed in the sample email\'s own <img> tags', () => {
  assert.equal(buildEntradioQrCodeUrl('TM5X59GM'), 'https://app.entradio.cz/qrcode?code=TM5X59GM&size=200');
  assert.equal(buildEntradioQrCodeUrl('2ZKN9JXVT'), 'https://app.entradio.cz/qrcode?code=2ZKN9JXVT&size=200');
});

test('buildEntradioQrCodeUrl: percent-encodes the code -- an unencoded separator would silently truncate the query string', () => {
  assert.equal(buildEntradioQrCodeUrl('A&B=C D'), 'https://app.entradio.cz/qrcode?code=A%26B%3DC%20D&size=200');
});

// --- response validators ----------------------------------------------------
//
// muteHttpExceptions is required precisely so a non-200 arrives as a VALUE
// rather than a throw; these two functions are what turn that value into a
// decision. The text/html rejection is the important one: an expired SendGrid
// click wrapper, or a login wall, answers 200 with an HTML page -- saving that
// to Drive and attaching it to the calendar would look like a success and be
// worthless at the door.

test('isEntradioTicketFileResponseAcceptable: a 200 with a real file content-type is accepted', () => {
  assert.equal(isEntradioTicketFileResponseAcceptable(200, 'application/pdf'), true);
  assert.equal(isEntradioTicketFileResponseAcceptable(200, 'application/octet-stream'), true);
  assert.equal(isEntradioTicketFileResponseAcceptable(200, 'image/png'), true);
});

test('isEntradioTicketFileResponseAcceptable: a 200 that returns text/html is REJECTED -- that is a login/error page, not a ticket', () => {
  assert.equal(isEntradioTicketFileResponseAcceptable(200, 'text/html'), false);
  assert.equal(isEntradioTicketFileResponseAcceptable(200, 'text/html; charset=utf-8'), false);
  assert.equal(isEntradioTicketFileResponseAcceptable(200, 'TEXT/HTML; charset=UTF-8'), false);
});

test('isEntradioTicketFileResponseAcceptable: any non-200 is rejected, including a redirect followRedirects failed to resolve', () => {
  assert.equal(isEntradioTicketFileResponseAcceptable(302, 'application/pdf'), false);
  assert.equal(isEntradioTicketFileResponseAcceptable(403, 'application/pdf'), false);
  assert.equal(isEntradioTicketFileResponseAcceptable(404, 'application/pdf'), false);
  assert.equal(isEntradioTicketFileResponseAcceptable(500, 'application/pdf'), false);
});

test('isEntradioQrCodeResponseAcceptable: a 200 with an image/* content-type is accepted, case-insensitively', () => {
  assert.equal(isEntradioQrCodeResponseAcceptable(200, 'image/png'), true);
  assert.equal(isEntradioQrCodeResponseAcceptable(200, 'IMAGE/PNG'), true);
  assert.equal(isEntradioQrCodeResponseAcceptable(200, 'image/jpeg'), true);
});

test('isEntradioQrCodeResponseAcceptable: a QR endpoint is held to a STRICTER rule than the ticket file -- anything not image/* is rejected even at 200', () => {
  assert.equal(isEntradioQrCodeResponseAcceptable(200, 'text/html'), false);
  assert.equal(isEntradioQrCodeResponseAcceptable(200, 'application/pdf'), false);
  assert.equal(isEntradioQrCodeResponseAcceptable(200, 'application/octet-stream'), false);
  assert.equal(isEntradioQrCodeResponseAcceptable(200, ''), false);
  assert.equal(isEntradioQrCodeResponseAcceptable(500, 'image/png'), false);
});

// --- filenames --------------------------------------------------------------
//
// buildTicketAttachmentFilename hardcodes ".pdf", which is exactly why it is
// not reused here: the real format behind the Entradio download link is
// UNVERIFIED until live-tested, so the extension has to come from the fetched
// blob's own content-type rather than from an assumption.

test('entradioFileExtensionForMimeType: maps the formats a ticket download plausibly returns', () => {
  assert.equal(entradioFileExtensionForMimeType('application/pdf'), '.pdf');
  assert.equal(entradioFileExtensionForMimeType('image/png'), '.png');
  assert.equal(entradioFileExtensionForMimeType('image/jpeg'), '.jpg');
  assert.equal(entradioFileExtensionForMimeType('application/zip'), '.zip');
});

test('entradioFileExtensionForMimeType: strips content-type parameters and lowercases before matching', () => {
  assert.equal(entradioFileExtensionForMimeType('Application/PDF; charset=binary'), '.pdf');
  assert.equal(entradioFileExtensionForMimeType('  image/png  '), '.png');
});

test('entradioFileExtensionForMimeType: an unknown or missing content-type yields NO extension -- an honest missing suffix beats a confidently wrong one', () => {
  assert.equal(entradioFileExtensionForMimeType('application/octet-stream'), '');
  assert.equal(entradioFileExtensionForMimeType('application/x-unheard-of'), '');
  assert.equal(entradioFileExtensionForMimeType(''), '');
  assert.equal(entradioFileExtensionForMimeType(null), '');
  assert.equal(entradioFileExtensionForMimeType(undefined), '');
});

test('buildEntradioTicketAttachmentFilename: follows the shared "{event} - {YYYY-MM-DD} - {identifier}" convention, with the extension derived from the blob', () => {
  const components = { year: 2026, month: 8, day: 27, hour: 17, minute: 30 };

  assert.equal(
    buildEntradioTicketAttachmentFilename('ČERNO, VÍR', components, '2354152', 'application/pdf'),
    'ČERNO, VÍR - 2026-09-27 - 2354152.pdf'
  );
  assert.equal(
    buildEntradioTicketAttachmentFilename('ČERNO, VÍR', components, '2354152', 'application/zip'),
    'ČERNO, VÍR - 2026-09-27 - 2354152.zip'
  );
});

// THE DRIFT GUARD for the deliberate duplication: this function reimplements
// buildTicketAttachmentFilename's stem rather than delegating to it, so that
// the three already-live portals' naming path is not touched at all. That
// choice is only safe while the two agree -- so pin it.
test('buildEntradioTicketAttachmentFilename: produces the EXACT same name as buildTicketAttachmentFilename when the blob is a PDF -- the two must never drift apart', () => {
  const components = { year: 2026, month: 8, day: 27 };

  assert.equal(
    buildEntradioTicketAttachmentFilename('ČERNO, VÍR', components, '2354152', 'application/pdf'),
    buildTicketAttachmentFilename('ČERNO, VÍR', components, '2354152')
  );
  assert.equal(
    buildEntradioTicketAttachmentFilename('Tajný ostrov', { year: 2026, month: 7, day: 7 }, null, 'application/pdf'),
    buildTicketAttachmentFilename('Tajný ostrov', { year: 2026, month: 7, day: 7 }, null)
  );
});

test('buildEntradioTicketAttachmentFilename: a null ticketIdentifier omits the segment entirely, never the literal word "null" (the round-4 Kino Art incident)', () => {
  const name = buildEntradioTicketAttachmentFilename('ČERNO, VÍR', { year: 2026, month: 8, day: 27 }, null, 'application/pdf');

  assert.equal(name, 'ČERNO, VÍR - 2026-09-27.pdf');
  assert.equal(name.includes('null'), false);
});

test('buildEntradioTicketAttachmentFilename: filesystem-unsafe characters in the event name are sanitized', () => {
  assert.equal(
    buildEntradioTicketAttachmentFilename('AC/DC: Live?', { year: 2026, month: 0, day: 5 }, '9', 'application/pdf'),
    'AC-DC- Live- - 2026-01-05 - 9.pdf'
  );
});

test('buildEntradioQrCodeFilename: one file per seat, named "{event} - QR - {code}.png"', () => {
  assert.equal(buildEntradioQrCodeFilename('ČERNO, VÍR', 'TM5X59GM'), 'ČERNO, VÍR - QR - TM5X59GM.png');
  assert.equal(buildEntradioQrCodeFilename('ČERNO, VÍR', '2ZKN9JXVT'), 'ČERNO, VÍR - QR - 2ZKN9JXVT.png');
});

test('buildEntradioQrCodeFilename: the two seats of one order produce DISTINCT filenames -- the code is what disambiguates them', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);
  const names = parsed.ticketCodes.map(function (code) {
    return buildEntradioQrCodeFilename(parsed.eventName, code);
  });

  assert.equal(names.length, 2);
  assert.notEqual(names[0], names[1]);
});

test('buildEntradioQrCodeFilename: sanitizes both the event name and the code', () => {
  assert.equal(buildEntradioQrCodeFilename('AC/DC', 'A/B'), 'AC-DC - QR - A-B.png');
});

// --- buildTicketCalendarEventResource ---------------------------------------
//
// EXTRACTED FROM createTicketCalendarEvent so that round 2's third-parameter
// change (a single attachmentInfo object -> an ARRAY of attachments, each
// carrying its OWN mimeType) is PROVABLE rather than merely reviewed.
// createTicketCalendarEvent keeps only the two GAS calls it cannot shed
// (CalendarApp.getCalendarById().getTimeZone(), Calendar.Events.insert).
//
// THE BLAST-RADIUS TEST is the first one below: enigoo.cz, Kino Art and
// Ticketmaster CZ all now pass a ONE-ELEMENT array with mimeType
// 'application/pdf' where they used to pass a bare object, and the resource
// that reaches Calendar.Events.insert must be byte-for-byte what it was
// before round 2.

const ENIGOO_SHAPED_PARSED_TICKET = {
  eventName: 'Letní hudební festival',
  location: 'Nádvoří kulturního domu',
  year: 2026,
  month: 7,
  day: 15,
  hour: 19,
  minute: 0,
  ticketIdentifier: '24601',
};

test('buildTicketCalendarEventResource: a one-element PDF array produces EXACTLY the pre-round-2 resource -- the three live portals are provably unaffected by the signature change', () => {
  const built = buildTicketCalendarEventResource(ENIGOO_SHAPED_PARSED_TICKET, 'Europe/Prague', [
    {
      fileId: 'FILE_ID',
      fileUrl: 'https://drive.example/FILE_ID',
      title: 'Letní hudební festival - 2026-08-15 - 24601.pdf',
      mimeType: 'application/pdf',
    },
  ]);

  assert.deepEqual(built.resource, {
    summary: 'Letní hudební festival',
    location: 'Nádvoří kulturního domu',
    start: { dateTime: '2026-08-15T19:00:00', timeZone: 'Europe/Prague' },
    end: { dateTime: '2026-08-15T21:00:00', timeZone: 'Europe/Prague' },
    extendedProperties: { private: { ticketIdentifier: '24601' } },
    attachments: [
      {
        fileId: 'FILE_ID',
        fileUrl: 'https://drive.example/FILE_ID',
        title: 'Letní hudební festival - 2026-08-15 - 24601.pdf',
        mimeType: 'application/pdf',
      },
    ],
  });
  assert.deepEqual(built.optionalArgs, { supportsAttachments: true });
});

test('buildTicketCalendarEventResource: with NO attachments the resource has no attachments key and supportsAttachments is never set -- unchanged no-attachment behavior', () => {
  const forEmptyArray = buildTicketCalendarEventResource(ENIGOO_SHAPED_PARSED_TICKET, 'Europe/Prague', []);
  const forNull = buildTicketCalendarEventResource(ENIGOO_SHAPED_PARSED_TICKET, 'Europe/Prague', null);
  const forUndefined = buildTicketCalendarEventResource(ENIGOO_SHAPED_PARSED_TICKET, 'Europe/Prague', undefined);

  [forEmptyArray, forNull, forUndefined].forEach(function (built) {
    assert.equal(Object.prototype.hasOwnProperty.call(built.resource, 'attachments'), false);
    assert.deepEqual(built.optionalArgs, {});
  });
});

test('buildTicketCalendarEventResource: EVERY attachment keeps its OWN mimeType -- mimeType is no longer hardcoded to application/pdf', () => {
  const built = buildTicketCalendarEventResource(ENIGOO_SHAPED_PARSED_TICKET, 'Europe/Prague', [
    { fileId: 'F1', fileUrl: 'u1', title: 'ticket.pdf', mimeType: 'application/pdf' },
    { fileId: 'F2', fileUrl: 'u2', title: 'qr-a.png', mimeType: 'image/png' },
    { fileId: 'F3', fileUrl: 'u3', title: 'qr-b.png', mimeType: 'image/png' },
  ]);

  assert.equal(built.resource.attachments.length, 3);
  assert.deepEqual(
    built.resource.attachments.map(function (a) {
      return a.mimeType;
    }),
    ['application/pdf', 'image/png', 'image/png']
  );
  // Order is preserved: the ticket file first, then one QR per seat in seat
  // order, which is how they render on the calendar event.
  assert.deepEqual(
    built.resource.attachments.map(function (a) {
      return a.title;
    }),
    ['ticket.pdf', 'qr-a.png', 'qr-b.png']
  );
  assert.deepEqual(built.optionalArgs, { supportsAttachments: true });
});

test('buildTicketCalendarEventResource: description is set only when the parsed ticket carries one (enigoo.cz and Kino Art never do)', () => {
  const withoutDescription = buildTicketCalendarEventResource(ENIGOO_SHAPED_PARSED_TICKET, 'Europe/Prague', []);
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);
  const withDescription = buildTicketCalendarEventResource(parsed, 'Europe/Prague', []);

  assert.equal(Object.prototype.hasOwnProperty.call(withoutDescription.resource, 'description'), false);
  assert.equal(withDescription.resource.description, parsed.description);
});

test('buildTicketCalendarEventResource: a parser that could not extract a ticketIdentifier produces no extendedProperties tag at all', () => {
  const untagged = Object.assign({}, ENIGOO_SHAPED_PARSED_TICKET, { ticketIdentifier: null });
  const built = buildTicketCalendarEventResource(untagged, 'Europe/Prague', []);

  assert.equal(Object.prototype.hasOwnProperty.call(built.resource, 'extendedProperties'), false);
});

test('buildTicketCalendarEventResource: the real Entradio order becomes ONE 17:30-19:30 event tagged with the order number', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);
  const built = buildTicketCalendarEventResource(parsed, 'Europe/Prague', []);

  assert.equal(built.resource.summary, 'ČERNO, VÍR');
  assert.equal(built.resource.location, 'Kino Metropol, Sokolská 572/25, 77900 Olomouc, Česká republika');
  assert.equal(built.resource.start.dateTime, '2026-09-27T17:30:00');
  assert.equal(built.resource.end.dateTime, '2026-09-27T19:30:00');
  assert.deepEqual(built.resource.extendedProperties, { private: { ticketIdentifier: '2354152' } });
});

// --- the attachment-fetcher registry ----------------------------------------

test('TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL: Entradio is wired to fetchEntradioAttachments', () => {
  assert.strictEqual(
    TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL['no-reply@app.entradio.cz'],
    fetchEntradioAttachments
  );
});

// THE "EXISTING PORTALS UNCHANGED" GUARD at the registry level: the new
// fetcher path must be reachable ONLY from Entradio. A key appearing here for
// any other portal would start making outbound HTTP calls on their behalf.
test('TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL: has NO key for any other portal -- Entradio is the only sender that triggers an outbound fetch', () => {
  assert.deepEqual(Object.keys(TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL), ['no-reply@app.entradio.cz']);
});

test('the three pre-existing portals are still routed EXACTLY as before: same body parsers, same PDF finders, no attachment fetcher', () => {
  assert.strictEqual(TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL['no-reply@enigoo.cz'], parseEnigooTicketText);
  assert.strictEqual(TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL['rezervace@kinoart.cz'], parseKinoArtTicketText);
  assert.strictEqual(TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL['noreply@ticketmaster.cz'], parseTicketmasterCzTicketText);
  assert.strictEqual(TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL['rezervace@kinoart.cz'], findKinoArtTicketPdfAttachment);
  assert.strictEqual(
    TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL['noreply@ticketmaster.cz'],
    findTicketmasterCzTicketPdfAttachment
  );

  ['no-reply@enigoo.cz', 'rezervace@kinoart.cz', 'noreply@ticketmaster.cz'].forEach(function (sender) {
    assert.equal(Object.prototype.hasOwnProperty.call(TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL, sender), false);
  });
});

// --- fetchEntradioAttachments -----------------------------------------------
//
// The orchestrator IS unit-tested despite touching UrlFetchApp/DriveApp/CONFIG,
// using the same global-injection harness this repo already established for the
// transport-tickets and ICS actions (GAS concatenates every file into one
// shared global scope, so a bare `UrlFetchApp` reference resolves through
// globalThis under Node too). It is tested because its contract is a NEGATIVE
// one -- "never throws, always returns an array" -- and a negative contract
// cannot be verified by reading the happy path.

function entradioFakeBlob(contentType) {
  return {
    getContentType: function () {
      return contentType;
    },
  };
}

function withEntradioFetchGlobals(options, fn) {
  const previous = {
    CONFIG: global.CONFIG,
    UrlFetchApp: global.UrlFetchApp,
    DriveApp: global.DriveApp,
  };
  const realConsoleLog = console.log;
  const calls = { fetched: [], created: [], folders: [] };

  global.CONFIG = { ticketAttachmentDriveFolderName: 'GAS Email Manager - Tickets' };

  global.UrlFetchApp = {
    fetch: function (url, params) {
      calls.fetched.push({ url: url, params: params });
      const responder = options.respond(url);
      if (responder instanceof Error) {
        throw responder;
      }
      return {
        getResponseCode: function () {
          return responder.code;
        },
        getBlob: function () {
          return entradioFakeBlob(responder.contentType);
        },
      };
    },
  };

  let createdCount = 0;
  global.DriveApp = {
    getFoldersByName: function (name) {
      calls.folders.push(name);
      if (options.driveThrows) {
        throw new Error('Drive is unavailable');
      }
      return {
        hasNext: function () {
          return true;
        },
        next: function () {
          return {
            createFile: function (blob) {
              createdCount += 1;
              const record = { id: 'FILE_' + createdCount, name: null, contentType: blob.getContentType() };
              calls.created.push(record);
              return {
                getId: function () {
                  return record.id;
                },
                getUrl: function () {
                  return 'https://drive.example/' + record.id;
                },
                getName: function () {
                  return record.name;
                },
                setName: function (newName) {
                  record.name = newName;
                },
              };
            },
          };
        },
      };
    },
  };

  console.log = function () {};

  try {
    return fn(calls);
  } finally {
    console.log = realConsoleLog;
    Object.keys(previous).forEach(function (key) {
      if (previous[key] === undefined) {
        delete global[key];
      } else {
        global[key] = previous[key];
      }
    });
  }
}

function entradioFakeMessage(htmlBody) {
  return {
    getBody: function () {
      if (htmlBody instanceof Error) {
        throw htmlBody;
      }
      return htmlBody;
    },
  };
}

const ENTRADIO_PARSED_FOR_FETCH = {
  eventName: 'ČERNO, VÍR',
  year: 2026,
  month: 8,
  day: 27,
  ticketIdentifier: '2354152',
  ticketCodes: ['TM5X59GM', '2ZKN9JXVT'],
};

function entradioRespondOk(url) {
  if (url.indexOf('app.entradio.cz/qrcode') !== -1) {
    return { code: 200, contentType: 'image/png' };
  }
  return { code: 200, contentType: 'application/pdf' };
}

test('fetchEntradioAttachments: with insertPdfIntoEvent ON, returns the ticket file FIRST then one QR attachment per seat, all in the shared permanent folder', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function (calls) {
    const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
      identifyingEmail: 'no-reply@app.entradio.cz',
      insertPdfIntoEvent: true,
    });

    assert.deepEqual(
      attachments.map(function (a) {
        return a.title;
      }),
      ['ČERNO, VÍR - 2026-09-27 - 2354152.pdf', 'ČERNO, VÍR - QR - TM5X59GM.png', 'ČERNO, VÍR - QR - 2ZKN9JXVT.png']
    );
    assert.deepEqual(
      attachments.map(function (a) {
        return a.mimeType;
      }),
      ['application/pdf', 'image/png', 'image/png']
    );
    attachments.forEach(function (a) {
      assert.equal(typeof a.fileId, 'string');
      assert.equal(a.fileUrl, 'https://drive.example/' + a.fileId);
    });

    // The EXISTING shared folder, never a new Entradio-specific one.
    assert.deepEqual(
      calls.folders.filter(function (name, i, all) {
        return all.indexOf(name) === i;
      }),
      ['GAS Email Manager - Tickets']
    );
  });
});

test('fetchEntradioAttachments: every fetch uses followRedirects AND muteHttpExceptions -- a SendGrid click wrapper IS a redirect, and a non-200 must arrive as a value not a throw', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function (calls) {
    fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
      identifyingEmail: 'no-reply@app.entradio.cz',
      insertPdfIntoEvent: true,
    });

    assert.equal(calls.fetched.length, 3);
    calls.fetched.forEach(function (call) {
      assert.equal(call.params.followRedirects, true);
      assert.equal(call.params.muteHttpExceptions, true);
    });
    assert.equal(calls.fetched[0].url, EXPECTED_ENTRADIO_DOWNLOAD_URL);
    assert.equal(calls.fetched[1].url, 'https://app.entradio.cz/qrcode?code=TM5X59GM&size=200');
    assert.equal(calls.fetched[2].url, 'https://app.entradio.cz/qrcode?code=2ZKN9JXVT&size=200');
  });
});

// THE OWNER'S EXPLICIT DECISION, pinned: a QR code is not a PDF, so the PDF
// toggle has no business gating it. A regression here would silently strip the
// one artifact that actually gets the owner through the door.
test('fetchEntradioAttachments: with insertPdfIntoEvent OFF, the QR codes are STILL fetched -- only the ticket-file download is gated by that toggle', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function (calls) {
    const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
      identifyingEmail: 'no-reply@app.entradio.cz',
      insertPdfIntoEvent: false,
    });

    assert.deepEqual(
      attachments.map(function (a) {
        return a.title;
      }),
      ['ČERNO, VÍR - QR - TM5X59GM.png', 'ČERNO, VÍR - QR - 2ZKN9JXVT.png']
    );
    // The download link is never even fetched when the toggle is off.
    assert.equal(
      calls.fetched.some(function (call) {
        return call.url === EXPECTED_ENTRADIO_DOWNLOAD_URL;
      }),
      false
    );
  });
});

test('fetchEntradioAttachments: a ticket download answering 200 with text/html is rejected, but the QR codes still come through', () => {
  withEntradioFetchGlobals(
    {
      respond: function (url) {
        if (url.indexOf('app.entradio.cz/qrcode') !== -1) {
          return { code: 200, contentType: 'image/png' };
        }
        return { code: 200, contentType: 'text/html; charset=utf-8' };
      },
    },
    function () {
      const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
        identifyingEmail: 'no-reply@app.entradio.cz',
        insertPdfIntoEvent: true,
      });

      assert.equal(attachments.length, 2);
      assert.deepEqual(
        attachments.map(function (a) {
          return a.mimeType;
        }),
        ['image/png', 'image/png']
      );
    }
  );
});

test('fetchEntradioAttachments: ONE failing QR fetch does not take the other seat down with it', () => {
  withEntradioFetchGlobals(
    {
      respond: function (url) {
        if (url.indexOf('code=TM5X59GM') !== -1) {
          return { code: 500, contentType: 'text/plain' };
        }
        return entradioRespondOk(url);
      },
    },
    function () {
      const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
        identifyingEmail: 'no-reply@app.entradio.cz',
        insertPdfIntoEvent: true,
      });

      assert.deepEqual(
        attachments.map(function (a) {
          return a.title;
        }),
        ['ČERNO, VÍR - 2026-09-27 - 2354152.pdf', 'ČERNO, VÍR - QR - 2ZKN9JXVT.png']
      );
    }
  );
});

test('fetchEntradioAttachments: an HTML body with no STÁHNOUT VSTUPENKY link yields the QR codes only, never a throw', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function () {
    const attachments = fetchEntradioAttachments(
      entradioFakeMessage('<html><body>no button here</body></html>'),
      ENTRADIO_PARSED_FOR_FETCH,
      { identifyingEmail: 'no-reply@app.entradio.cz', insertPdfIntoEvent: true }
    );

    assert.equal(attachments.length, 2);
  });
});

// THE NEGATIVE CONTRACT. processTicketFromMessageBody calls this BEFORE
// creating the calendar event, so anything escaping here would abort the event
// -- exactly the outcome the owner ruled out ("vytvořit událost i tak").
test('fetchEntradioAttachments: UrlFetchApp itself throwing returns [] -- it NEVER propagates, so the calendar event is never blocked by a network failure', () => {
  withEntradioFetchGlobals(
    {
      respond: function () {
        return new Error('DNS lookup failed');
      },
    },
    function () {
      const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
        identifyingEmail: 'no-reply@app.entradio.cz',
        insertPdfIntoEvent: true,
      });

      assert.deepEqual(attachments, []);
    }
  );
});

test('fetchEntradioAttachments: Drive throwing on every write returns [] rather than propagating', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk, driveThrows: true }, function () {
    const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), ENTRADIO_PARSED_FOR_FETCH, {
      identifyingEmail: 'no-reply@app.entradio.cz',
      insertPdfIntoEvent: true,
    });

    assert.deepEqual(attachments, []);
  });
});

test('fetchEntradioAttachments: message.getBody() throwing returns [] rather than propagating', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function () {
    const attachments = fetchEntradioAttachments(
      entradioFakeMessage(new Error('body unavailable')),
      { eventName: 'ČERNO, VÍR', year: 2026, month: 8, day: 27, ticketIdentifier: '2354152', ticketCodes: [] },
      { identifyingEmail: 'no-reply@app.entradio.cz', insertPdfIntoEvent: true }
    );

    assert.deepEqual(attachments, []);
  });
});

test('fetchEntradioAttachments: a parsed ticket with no ticketCodes and the toggle off returns [] without any fetch at all', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function (calls) {
    const attachments = fetchEntradioAttachments(
      entradioFakeMessage(REAL_ENTRADIO_HTML_BODY),
      { eventName: 'ČERNO, VÍR', year: 2026, month: 8, day: 27, ticketIdentifier: '2354152', ticketCodes: [] },
      { identifyingEmail: 'no-reply@app.entradio.cz', insertPdfIntoEvent: false }
    );

    assert.deepEqual(attachments, []);
    assert.equal(calls.fetched.length, 0);
  });
});

test('fetchEntradioAttachments: a parsed ticket missing ticketCodes entirely (an older shape) degrades to the ticket file only, never a throw', () => {
  withEntradioFetchGlobals({ respond: entradioRespondOk }, function () {
    const attachments = fetchEntradioAttachments(
      entradioFakeMessage(REAL_ENTRADIO_HTML_BODY),
      { eventName: 'ČERNO, VÍR', year: 2026, month: 8, day: 27, ticketIdentifier: '2354152' },
      { identifyingEmail: 'no-reply@app.entradio.cz', insertPdfIntoEvent: true }
    );

    assert.equal(attachments.length, 1);
    assert.equal(attachments[0].mimeType, 'application/pdf');
  });
});

test('fetchEntradioAttachments: END-TO-END from the REAL fixture -- parse the real body, then attach one ticket file plus exactly one QR per real seat code', () => {
  const parsed = parseEntradioTicketText(REAL_ENTRADIO_BODY_TEXT);

  withEntradioFetchGlobals({ respond: entradioRespondOk }, function () {
    const attachments = fetchEntradioAttachments(entradioFakeMessage(REAL_ENTRADIO_HTML_BODY), parsed, {
      identifyingEmail: 'no-reply@app.entradio.cz',
      insertPdfIntoEvent: true,
    });

    assert.equal(attachments.length, 1 + parsed.ticketCodes.length);
    parsed.ticketCodes.forEach(function (code) {
      assert.equal(
        attachments.some(function (a) {
          return a.title === 'ČERNO, VÍR - QR - ' + code + '.png';
        }),
        true
      );
    });
  });
});

// --- processTicketFromMessageBody: the ROUND 2 WIRING ------------------------
//
// WHY THESE EXIST, and why they were written AFTER the first two round-2
// commits: the mutation pass over round 2's new code killed 12 of 15 mutants,
// and the three survivors were all in this function -- "never concatenate the
// fetched attachments onto the event", "never send the zero-attachment
// notification", and a redundant inner catch. The first two are not obscure
// edge cases, they are ROUND 2'S ENTIRE POINT: with either one applied, every
// pure function below still behaved perfectly and the whole feature silently
// did nothing. A signal that cannot see that is not a signal.
//
// This function is GAS-only (GmailMessage/DriveApp/CalendarApp/Calendar/
// UrlFetchApp), which is why it had never been unit-tested. It is tested here
// through the same global-injection harness the transport-tickets and ICS
// actions already use for their own GAS-only pipelines: Apps Script
// concatenates every project file into ONE shared global scope, so a bare
// `CONFIG` / `UrlFetchApp` / `notifyOwnerOfTicketAttachmentFailure` reference
// resolves through globalThis under Node too.
//
// The THIRD test is this round's "existing portals provably unchanged" signal
// in its strongest available form: a Kino Art message driven end-to-end
// through the SAME changed function, asserting the event it produces still
// carries exactly one application/pdf attachment and that ZERO outbound HTTP
// calls were made on its behalf.

function withTicketBodyRunGlobals(options, fn) {
  const previous = {
    CONFIG: global.CONFIG,
    Calendar: global.Calendar,
    CalendarApp: global.CalendarApp,
    DriveApp: global.DriveApp,
    UrlFetchApp: global.UrlFetchApp,
    notifyOwnerOfTicketAttachmentFailure: global.notifyOwnerOfTicketAttachmentFailure,
  };
  const realConsoleLog = console.log;
  const calls = { inserted: [], fetched: [], notified: [], created: [] };

  global.CONFIG = {
    calendarId: 'DEFAULT_CAL',
    ticketAttachmentDriveFolderName: 'GAS Email Manager - Tickets',
  };

  global.CalendarApp = {
    getCalendarById: function () {
      return {
        getTimeZone: function () {
          return 'Europe/Prague';
        },
      };
    },
  };

  global.Calendar = {
    Events: {
      // No pre-existing tagged event: the dedup safety net never short-circuits
      // in these tests, so the creation path is always the one exercised.
      list: function () {
        return { items: [] };
      },
      insert: function (resource, calendarId, optionalArgs) {
        calls.inserted.push({ resource: resource, calendarId: calendarId, optionalArgs: optionalArgs });
      },
    },
  };

  global.UrlFetchApp = {
    fetch: function (url, params) {
      calls.fetched.push({ url: url, params: params });
      const responder = options.respond ? options.respond(url) : { code: 404, contentType: 'text/plain' };
      if (responder instanceof Error) {
        throw responder;
      }
      return {
        getResponseCode: function () {
          return responder.code;
        },
        getBlob: function () {
          return entradioFakeBlob(responder.contentType);
        },
      };
    },
  };

  let createdCount = 0;
  global.DriveApp = {
    getFoldersByName: function () {
      return {
        hasNext: function () {
          return true;
        },
        next: function () {
          return {
            createFile: function () {
              createdCount += 1;
              const record = { id: 'FILE_' + createdCount, name: null };
              calls.created.push(record);
              return {
                getId: function () {
                  return record.id;
                },
                getUrl: function () {
                  return 'https://drive.example/' + record.id;
                },
                getName: function () {
                  return record.name;
                },
                setName: function (newName) {
                  record.name = newName;
                },
              };
            },
          };
        },
      };
    },
  };

  // Lives in src/02-main.js. Under GAS both files share one global scope; under
  // Node it has to be wired the way the runtime would, same as the transport
  // action's own harness does for importIcsEventWithSequenceRetry.
  global.notifyOwnerOfTicketAttachmentFailure = function (eventName, calendarId, ticketIdentifier) {
    calls.notified.push({ eventName: eventName, calendarId: calendarId, ticketIdentifier: ticketIdentifier });
  };

  console.log = function () {};

  try {
    return fn(calls);
  } finally {
    console.log = realConsoleLog;
    Object.keys(previous).forEach(function (key) {
      if (previous[key] === undefined) {
        delete global[key];
      } else {
        global[key] = previous[key];
      }
    });
  }
}

function entradioBodyModeMessage() {
  return {
    getFrom: function () {
      return 'Kino Metropol <no-reply@app.entradio.cz>';
    },
    getPlainBody: function () {
      return REAL_ENTRADIO_BODY_TEXT;
    },
    getBody: function () {
      return REAL_ENTRADIO_HTML_BODY;
    },
    getAttachments: function () {
      return [fakeAttachment('VOP_Metropol.pdf', 'application/pdf')];
    },
    // quick-260921-gj0: processTicketFromMessageBody now calls
    // message.getDate() unconditionally (D-07) -- parseEntradioTicketText
    // takes a single parameter and ignores the extra argument, so the exact
    // date value here is inert; it only needs to exist so the call doesn't
    // throw.
    getDate: function () {
      return new Date(2026, 8, 20);
    },
    // quick-260921-gj0 round 2: processTicketFromMessageBody now ALSO calls
    // message.getSubject() unconditionally (D-25/D-27) -- same reasoning as
    // getDate() above, inert for parseEntradioTicketText.
    getSubject: function () {
      return 'Potvrzení objednávky Entradio';
    },
  };
}

const ENTRADIO_PORTAL_PDF_ON = {
  identifyingEmail: 'no-reply@app.entradio.cz',
  calendarId: 'KINO_CAL',
  insertPdfIntoEvent: true,
};

// KILLS THE "never concatenate the fetched attachments" MUTANT. Everything
// upstream of this line can be perfect and the owner still gets a bare event.
test('processTicketFromMessageBody: an Entradio message creates ONE event carrying the downloaded ticket file AND both seats\' QR codes as real Calendar attachments', () => {
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(entradioBodyModeMessage(), ENTRADIO_PORTAL_PDF_ON);

    assert.equal(calls.inserted.length, 1);
    const insert = calls.inserted[0];

    assert.equal(insert.calendarId, 'KINO_CAL');
    assert.equal(insert.resource.summary, 'ČERNO, VÍR');
    assert.equal(insert.resource.start.dateTime, '2026-09-27T17:30:00');
    assert.equal(insert.resource.end.dateTime, '2026-09-27T19:30:00');
    assert.deepEqual(insert.resource.extendedProperties, { private: { ticketIdentifier: '2354152' } });

    assert.deepEqual(
      insert.resource.attachments.map(function (a) {
        return a.title;
      }),
      ['ČERNO, VÍR - 2026-09-27 - 2354152.pdf', 'ČERNO, VÍR - QR - TM5X59GM.png', 'ČERNO, VÍR - QR - 2ZKN9JXVT.png']
    );
    assert.deepEqual(
      insert.resource.attachments.map(function (a) {
        return a.mimeType;
      }),
      ['application/pdf', 'image/png', 'image/png']
    );
    // Without this the Calendar API silently drops the whole attachments array.
    assert.deepEqual(insert.optionalArgs, { supportsAttachments: true });

    // No attachment failure, so no notification.
    assert.equal(calls.notified.length, 0);
  });
});

test('processTicketFromMessageBody: with insertPdfIntoEvent OFF, the Entradio event still carries both QR codes and no ticket file', () => {
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(entradioBodyModeMessage(), {
      identifyingEmail: 'no-reply@app.entradio.cz',
      calendarId: 'KINO_CAL',
      insertPdfIntoEvent: false,
    });

    assert.deepEqual(
      calls.inserted[0].resource.attachments.map(function (a) {
        return a.title;
      }),
      ['ČERNO, VÍR - QR - TM5X59GM.png', 'ČERNO, VÍR - QR - 2ZKN9JXVT.png']
    );
    assert.equal(calls.notified.length, 0);
  });
});

// KILLS THE "never send the zero-attachment notification" MUTANT, and pins the
// owner's explicit rule in the same assertion: "Vytvořit událost i tak, jen
// upozornit e-mailem" -- create the event regardless, just send a warning.
test('processTicketFromMessageBody: when NOTHING can be fetched, the event is STILL created and the owner is notified exactly once', () => {
  withTicketBodyRunGlobals(
    {
      respond: function () {
        return new Error('the network is down');
      },
    },
    function (calls) {
      processTicketFromMessageBody(entradioBodyModeMessage(), ENTRADIO_PORTAL_PDF_ON);

      // The event is not sacrificed for the attachments.
      assert.equal(calls.inserted.length, 1);
      assert.equal(calls.inserted[0].resource.summary, 'ČERNO, VÍR');
      assert.equal(Object.prototype.hasOwnProperty.call(calls.inserted[0].resource, 'attachments'), false);
      assert.deepEqual(calls.inserted[0].optionalArgs, {});

      assert.deepEqual(calls.notified, [{ eventName: 'ČERNO, VÍR', calendarId: 'KINO_CAL', ticketIdentifier: '2354152' }]);
    }
  );
});

test('processTicketFromMessageBody: a PARTIAL attachment result is not a failure -- some QR codes and no ticket file sends NO notification', () => {
  withTicketBodyRunGlobals(
    {
      respond: function (url) {
        if (url.indexOf('app.entradio.cz/qrcode') !== -1) {
          return { code: 200, contentType: 'image/png' };
        }
        return { code: 404, contentType: 'text/plain' };
      },
    },
    function (calls) {
      processTicketFromMessageBody(entradioBodyModeMessage(), ENTRADIO_PORTAL_PDF_ON);

      assert.equal(calls.inserted[0].resource.attachments.length, 2);
      assert.equal(calls.notified.length, 0);
    }
  );
});

// THIS ROUND'S "EXISTING PORTALS PROVABLY UNCHANGED" SIGNAL, at the highest
// level available: the same changed function, driven end-to-end for Kino Art.
// The event must still carry exactly ONE application/pdf attachment named by
// the unchanged buildTicketAttachmentFilename convention, and NOT ONE outbound
// HTTP call may be made on its behalf.

// Deliberately the SAME REAL_KINO_ART_BODY_TEXT fixture the Kino Art parser
// tests already use, rather than a fresh approximation: the whole point of
// this test is that nothing about Kino Art changed, so it must be driven by
// the very data that proved Kino Art worked in the first place.

function kinoArtBodyModeMessage() {
  return {
    getFrom: function () {
      return 'Kino Art <rezervace@kinoart.cz>';
    },
    getPlainBody: function () {
      return REAL_KINO_ART_BODY_TEXT;
    },
    getBody: function () {
      return '<html><body>irrelevant</body></html>';
    },
    getAttachments: function () {
      return [
        Object.assign(fakeAttachment('Vstupenky.pdf', 'application/pdf'), {
          copyBlob: function () {
            return entradioFakeBlob('application/pdf');
          },
        }),
        Object.assign(fakeAttachment('Doklad.pdf', 'application/pdf'), {
          copyBlob: function () {
            return entradioFakeBlob('application/pdf');
          },
        }),
      ];
    },
    // quick-260921-gj0: same reasoning as entradioBodyModeMessage's own
    // getDate() above -- parseKinoArtTicketText ignores the extra argument.
    getDate: function () {
      return new Date(2026, 8, 20);
    },
    // quick-260921-gj0 round 2: same reasoning as entradioBodyModeMessage's
    // own getSubject() above -- parseKinoArtTicketText ignores it too.
    getSubject: function () {
      return 'Potvrzení rezervace Kino Art';
    },
  };
}

test('processTicketFromMessageBody: KINO ART IS UNTOUCHED BY ROUND 2 -- still exactly one application/pdf attachment, still named by the unchanged convention, and ZERO outbound HTTP calls', () => {
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    const parsed = parseKinoArtTicketText(REAL_KINO_ART_BODY_TEXT);

    processTicketFromMessageBody(kinoArtBodyModeMessage(), {
      identifyingEmail: 'rezervace@kinoart.cz',
      calendarId: 'ART_CAL',
      insertPdfIntoEvent: true,
    });

    assert.equal(calls.inserted.length, 1);
    const insert = calls.inserted[0];

    assert.equal(insert.resource.attachments.length, 1);
    assert.equal(insert.resource.attachments[0].mimeType, 'application/pdf');
    assert.equal(
      insert.resource.attachments[0].title,
      buildTicketAttachmentFilename(parsed.eventName, parsed, parsed.ticketIdentifier)
    );
    assert.deepEqual(insert.optionalArgs, { supportsAttachments: true });

    // THE POINT: Kino Art has no registered attachment fetcher, so round 2's
    // network path is unreachable for it.
    assert.deepEqual(calls.fetched, []);
    assert.equal(calls.notified.length, 0);
  });
});

test('processTicketFromMessageBody: a Kino Art message with insertPdfIntoEvent OFF still creates an event with no attachments and makes no outbound call', () => {
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(kinoArtBodyModeMessage(), {
      identifyingEmail: 'rezervace@kinoart.cz',
      calendarId: 'ART_CAL',
      insertPdfIntoEvent: false,
    });

    assert.equal(Object.prototype.hasOwnProperty.call(calls.inserted[0].resource, 'attachments'), false);
    assert.deepEqual(calls.inserted[0].optionalArgs, {});
    assert.deepEqual(calls.fetched, []);
    assert.equal(calls.notified.length, 0);
  });
});

// --- parseFeverTicketText (quick-260921-gj0: the FIFTH supported portal, ---
// --- the FOURTH body-sourced one) --------------------------------------------
//
// Fever (hello@feverup.com) confirmation emails carry ALL calendar-event data
// in the BODY -- event name, venue/address, the Czech-abbreviated date/time,
// and the per-seat ticket codes. Real sample inspected directly (269 KB
// .eml, "Potvrzení nákupu na Fever_ Candlelight..."): the message is
// multipart/mixed with exactly TWO parts, one text/html and one
// application/pdf ticket attachment -- there is NO text/plain part at all,
// unlike every other portal in this file. message.getPlainBody() therefore
// returns GMAIL'S OWN rendering of that HTML, whose exact line breaking
// cannot be observed from the .eml and must not be assumed. Every anchor
// below is therefore a LITERAL MARKER (never a line position), each verified
// to occur exactly once in the rendered text.
//
// The date format is Czech-ABBREVIATED ("so 19 pro - 08:00 odp.") with a
// 12-hour day-period marker (odp./dop.) and carries NO YEAR anywhere in the
// body -- the only four-digit year in the whole message is the footer
// copyright line. The event year is therefore INFERRED from an injected
// reference date (the message's own received date), never from a live
// clock read inside the parser.
//
// hello@feverup.com also sends ordinary marketing mail, so a content
// detector (feverTextHasPurchaseDetails) is registered -- the same
// debug/ticketmaster-cz-order-confirm pattern already established for
// Ticketmaster CZ, applied here from day one rather than discovered live.
//
// FIXTURE PROVENANCE: reproduces the real sample's rendered text sequence,
// with a FICTIONAL event name, venue, ticket ID and seat codes (the real
// values are live entry credentials and must never reach the repo -- see
// push-public.bat's leak guard). The real character classes that matter are
// preserved: non-ASCII Czech letters and a colon in the event name, and a
// hyphen plus non-ASCII Czech letters in the venue string. The real
// invisible preheader characters (U+034F, U+200C, U+00AD) and at least one
// U+00A0 (non-breaking space) are reproduced explicitly so the fixture
// exercises feverNormalizeBodyText, not just the visible text.

const FEVER_FIXTURE_CODES = [
  'QWERT12345ZXCVB67890',
  'ASDFG23456HJKLM78901',
  'POIUY34567LKJHG89012',
  'MNBVC45678TYUIO90123',
  'ZXCVB56789QWERT01234',
];

const FEVER_FIXTURE_EVENT_NAME = 'Vánoční trhy: Staroměstské náměstí';
const FEVER_FIXTURE_LOCATION = 'Obecní dům - Náměstí Republiky 5, Praha-Vinohrady';
const FEVER_FIXTURE_DATE_LINE = 'so 19 pro - 08:00 odp.';
const FEVER_FIXTURE_TICKET_ID = '512345678';

// The real sample's preheader padding, reproduced with the SAME invisible
// characters actually present (U+034F, U+200C, U+00AD -- all removed by
// feverNormalizeBodyText) plus a non-breaking space (U+00A0, folded to a
// plain space).
const FEVER_PREHEADER_PADDING = '͏͏‌‌­­ ';

const FEVER_BODY_LINES = [
  FEVER_PREHEADER_PADDING,
  'Zobrazit v prohlížeči',
  'Děkujeme! Tady jsou podrobnosti o tvém nákupu',
  FEVER_FIXTURE_EVENT_NAME,
  'Koupit znova',
  FEVER_FIXTURE_LOCATION,
  'Zobrazit na mapě',
  FEVER_FIXTURE_DATE_LINE,
  'Změnit datum nebo čas',
  '',
  '5 x Balkon',
  'ID vstupenky: ' + FEVER_FIXTURE_TICKET_ID,
].concat(FEVER_FIXTURE_CODES, ['', 'Shrnutí objednávky']);

// REAL_FEVER_BODY_TEXT -- LF-joined (a separate test below proves a
// CRLF-joined variant of this same fixture parses identically).
const REAL_FEVER_BODY_TEXT = FEVER_BODY_LINES.join('\n');

// FEVER_REFERENCE_DATE -- stands in for the message's own received date
// (message.getDate()). Chosen so the fixture's event month/day (19
// December) falls LATER in the SAME year as this reference date (15
// November), so year inference resolves to the reference year with no
// rollover (D-07).
const FEVER_REFERENCE_DATE = new Date(2026, 10, 15);

const EXPECTED_FEVER_DESCRIPTION = [
  FEVER_FIXTURE_EVENT_NAME,
  FEVER_FIXTURE_LOCATION,
  '19 pro - 08:00 odp.',
  'ID vstupenky: ' + FEVER_FIXTURE_TICKET_ID,
  '5 x Balkon',
  FEVER_FIXTURE_CODES.join('\n'),
].join('\n\n');

// buildFeverFixture -- a parametrized variant of REAL_FEVER_BODY_TEXT for the
// tests below that only need to vary ONE field (the date line, the quantity
// line, the ticket-ID line, or the seat codes) while keeping every other
// anchor at its real-shape default.
function buildFeverFixture(overrides) {
  const opts = overrides || {};
  const eventName = opts.eventName === undefined ? FEVER_FIXTURE_EVENT_NAME : opts.eventName;
  const location = opts.location === undefined ? FEVER_FIXTURE_LOCATION : opts.location;
  const dateLine = opts.dateLine === undefined ? FEVER_FIXTURE_DATE_LINE : opts.dateLine;
  const quantityLine = opts.quantityLine === undefined ? '5 x Balkon' : opts.quantityLine;
  const ticketIdLine = opts.ticketIdLine === undefined ? 'ID vstupenky: ' + FEVER_FIXTURE_TICKET_ID : opts.ticketIdLine;
  const codes = opts.codes === undefined ? FEVER_FIXTURE_CODES : opts.codes;

  const lines = [
    'Zobrazit v prohlížeči',
    'Děkujeme! Tady jsou podrobnosti o tvém nákupu',
    eventName,
    'Koupit znova',
    location,
    'Zobrazit na mapě',
    dateLine,
    'Změnit datum nebo čas',
    '',
  ];
  if (quantityLine) {
    lines.push(quantityLine);
  }
  if (ticketIdLine) {
    lines.push(ticketIdLine);
  }
  codes.forEach(function (code) {
    lines.push(code);
  });
  lines.push('');
  lines.push('Shrnutí objednávky');

  return lines.join('\n');
}

test('parseFeverTicketText: parses the real fixture into the full expected shape in one assertion -- December proving the zero-indexed month, and the afternoon marker proving the 12-hour conversion', () => {
  assert.deepEqual(parseFeverTicketText(REAL_FEVER_BODY_TEXT, FEVER_REFERENCE_DATE), {
    eventName: FEVER_FIXTURE_EVENT_NAME,
    location: FEVER_FIXTURE_LOCATION,
    year: 2026,
    month: 11,
    day: 19,
    hour: 20,
    minute: 0,
    ticketIdentifier: FEVER_FIXTURE_TICKET_ID,
    ticketQuantity: 5,
    description: EXPECTED_FEVER_DESCRIPTION,
  });
});

const FEVER_MONTH_TABLE_ORDER = [
  ['led', 0],
  ['úno', 1],
  ['bře', 2],
  ['dub', 3],
  ['kvě', 4],
  ['čvn', 5],
  ['čvc', 6],
  ['srp', 7],
  ['zář', 8],
  ['říj', 9],
  ['lis', 10],
  ['pro', 11],
];

test('parseFeverTicketText: all twelve Czech abbreviated month names resolve to the correct zero-indexed month number (D-04)', () => {
  FEVER_MONTH_TABLE_ORDER.forEach(function (entry) {
    const token = entry[0];
    const expectedMonth = entry[1];
    const parsed = parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 ' + token + ' - 08:00 odp.' }), FEVER_REFERENCE_DATE);
    assert.equal(parsed.month, expectedMonth, 'month abbreviation "' + token + '" should resolve to ' + expectedMonth);
  });
});

const FEVER_DIACRITIC_MONTH_PAIRS = [
  ['úno', 'uno', 1],
  ['bře', 'bre', 2],
  ['kvě', 'kve', 4],
  ['čvn', 'cvn', 5],
  ['čvc', 'cvc', 6],
  ['zář', 'zar', 8],
  ['říj', 'rij', 9],
];

test('parseFeverTicketText: each of the seven diacritic-bearing month abbreviations also resolves in its diacritic-stripped form, to the SAME month number (D-04)', () => {
  FEVER_DIACRITIC_MONTH_PAIRS.forEach(function (pair) {
    const withDiacritics = parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 ' + pair[0] + ' - 08:00 odp.' }), FEVER_REFERENCE_DATE);
    const withoutDiacritics = parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 ' + pair[1] + ' - 08:00 odp.' }), FEVER_REFERENCE_DATE);
    assert.equal(withDiacritics.month, pair[2]);
    assert.equal(withoutDiacritics.month, pair[2]);
  });
});

test('parseFeverTicketText: invisible preheader characters and non-breaking spaces used as literal separators do not disrupt extraction (D-03)', () => {
  const separator = '͏‌­ ';
  const body = [
    'Zobrazit v prohlížeči',
    'Tady jsou podrobnosti o tvém nákupu' + separator + FEVER_FIXTURE_EVENT_NAME,
    'Koupit znova',
    FEVER_FIXTURE_LOCATION + separator + 'Zobrazit na mapě',
    FEVER_FIXTURE_DATE_LINE,
    '5 x Balkon',
    'ID vstupenky: ' + FEVER_FIXTURE_TICKET_ID,
  ].join('\n');

  const parsed = parseFeverTicketText(body, FEVER_REFERENCE_DATE);
  assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.equal(parsed.location, FEVER_FIXTURE_LOCATION);
});

test('parseFeverTicketText: a CRLF-joined variant of the real fixture parses identically to the default LF-joined fixture -- separator-agnostic', () => {
  const crlfBody = FEVER_BODY_LINES.join('\r\n');
  assert.deepEqual(
    parseFeverTicketText(crlfBody, FEVER_REFERENCE_DATE),
    parseFeverTicketText(REAL_FEVER_BODY_TEXT, FEVER_REFERENCE_DATE)
  );
});

test('parseFeverTicketText: THE LINE-LAYOUT PROOF -- the purchase-details marker sharing a line with the event name, and the venue sharing a line with the map-link marker, still parses to the SAME eventName and location (D-02/D-03)', () => {
  const sharedLineBody = [
    'Zobrazit v prohlížeči',
    'Děkujeme! Tady jsou podrobnosti o tvém nákupu ' + FEVER_FIXTURE_EVENT_NAME,
    'Koupit znova',
    FEVER_FIXTURE_LOCATION + ' Zobrazit na mapě',
    FEVER_FIXTURE_DATE_LINE,
    '5 x Balkon',
    'ID vstupenky: ' + FEVER_FIXTURE_TICKET_ID,
  ].join('\n');

  const parsed = parseFeverTicketText(sharedLineBody, FEVER_REFERENCE_DATE);
  assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.equal(parsed.location, FEVER_FIXTURE_LOCATION);
});

test('parseFeverTicketText: the "Koupit znova" buy-again label is never mistaken for the event name or the venue, whether present between them or absent entirely', () => {
  const withLabel = parseFeverTicketText(buildFeverFixture({}), FEVER_REFERENCE_DATE);
  assert.equal(withLabel.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.equal(withLabel.location, FEVER_FIXTURE_LOCATION);

  const withoutLabelBody = [
    'Zobrazit v prohlížeči',
    'Tady jsou podrobnosti o tvém nákupu',
    FEVER_FIXTURE_EVENT_NAME,
    FEVER_FIXTURE_LOCATION,
    'Zobrazit na mapě',
    FEVER_FIXTURE_DATE_LINE,
    '5 x Balkon',
    'ID vstupenky: ' + FEVER_FIXTURE_TICKET_ID,
  ].join('\n');
  const withoutLabel = parseFeverTicketText(withoutLabelBody, FEVER_REFERENCE_DATE);
  assert.equal(withoutLabel.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.equal(withoutLabel.location, FEVER_FIXTURE_LOCATION);
});

test('parseFeverTicketText: "odp." adds 12 to a 1-11 hour and leaves 12 as 12; "dop." leaves a 1-11 hour and turns 12 into 0; no marker at all is read as 24-hour (D-05)', () => {
  assert.equal(parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 pro - 08:00 odp.' }), FEVER_REFERENCE_DATE).hour, 20);
  assert.equal(parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 pro - 12:00 odp.' }), FEVER_REFERENCE_DATE).hour, 12);
  assert.equal(parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 pro - 08:00 dop.' }), FEVER_REFERENCE_DATE).hour, 8);
  assert.equal(parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 pro - 12:00 dop.' }), FEVER_REFERENCE_DATE).hour, 0);
  assert.equal(parseFeverTicketText(buildFeverFixture({ dateLine: 'so 19 pro - 20:00' }), FEVER_REFERENCE_DATE).hour, 20);
});

test('parseFeverTicketText: year inference -- an earlier reference month/day keeps the reference year, a later one rolls forward, an equal month/day keeps the reference year (D-07)', () => {
  const body = buildFeverFixture({ dateLine: 'so 19 pro - 08:00 odp.' }); // 19 December

  // Reference EARLIER in the year (15 November) -- event still ahead, same year.
  assert.equal(parseFeverTicketText(body, new Date(2026, 10, 15)).year, 2026);

  // Reference LATER in the year (31 December) -- event already passed, rolls forward.
  assert.equal(parseFeverTicketText(body, new Date(2026, 11, 31)).year, 2027);

  // Reference on the SAME month/day (19 December) -- bought on the day of the event.
  assert.equal(parseFeverTicketText(body, new Date(2026, 11, 19)).year, 2026);
});

test('parseFeverTicketText: a missing reference date, or a non-Date second argument, throws with the full raw text -- never a silent fallback to the current clock (D-07)', () => {
  const body = buildFeverFixture({});

  assert.throws(
    () => parseFeverTicketText(body),
    (err) => {
      assert.match(err.message, /reference date|received date/i);
      assert.ok(err.message.includes(body), 'error message should include the full raw text');
      return true;
    }
  );

  assert.throws(
    () => parseFeverTicketText(body, '2026-12-19'),
    (err) => {
      assert.match(err.message, /reference date|received date/i);
      assert.ok(err.message.includes(body), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseFeverTicketText: ticketIdentifier is the "ID vstupenky:" number; a body with no such label parses successfully with ticketIdentifier null and a truthy eventName (D-09)', () => {
  const parsed = parseFeverTicketText(REAL_FEVER_BODY_TEXT, FEVER_REFERENCE_DATE);
  assert.equal(parsed.ticketIdentifier, FEVER_FIXTURE_TICKET_ID);

  const noIdBody = buildFeverFixture({ ticketIdLine: '', codes: [] });
  const noIdParsed = parseFeverTicketText(noIdBody, FEVER_REFERENCE_DATE);
  assert.equal(noIdParsed.ticketIdentifier, null);
  assert.ok(noIdParsed.eventName);
});

test('parseFeverTicketText: ONE-EVENT-PER-PURCHASE -- quantities of 1, 2 and 5 all produce identical eventName/location/date fields, differing only in ticketQuantity and description (D-10)', () => {
  const q1 = parseFeverTicketText(buildFeverFixture({ quantityLine: '1 x Balkon' }), FEVER_REFERENCE_DATE);
  const q2 = parseFeverTicketText(buildFeverFixture({ quantityLine: '2 x Balkon' }), FEVER_REFERENCE_DATE);
  const q5 = parseFeverTicketText(buildFeverFixture({ quantityLine: '5 x Balkon' }), FEVER_REFERENCE_DATE);

  [q1, q2, q5].forEach(function (parsed) {
    assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
    assert.equal(parsed.location, FEVER_FIXTURE_LOCATION);
    assert.equal(parsed.year, 2026);
    assert.equal(parsed.month, 11);
    assert.equal(parsed.day, 19);
    assert.equal(parsed.hour, 20);
    assert.equal(parsed.minute, 0);
  });

  assert.equal(q1.ticketQuantity, 1);
  assert.equal(q2.ticketQuantity, 2);
  assert.equal(q5.ticketQuantity, 5);
  assert.notEqual(q1.description, q5.description);
});

test('parseFeverTicketText: description contains the event name, location, raw date line, ticket ID and every per-seat code; no ticketCodes property is returned at all (D-11)', () => {
  const parsed = parseFeverTicketText(REAL_FEVER_BODY_TEXT, FEVER_REFERENCE_DATE);

  assert.ok(parsed.description.includes(FEVER_FIXTURE_EVENT_NAME));
  assert.ok(parsed.description.includes(FEVER_FIXTURE_LOCATION));
  assert.ok(parsed.description.includes('19 pro - 08:00 odp.'));
  assert.ok(parsed.description.includes(FEVER_FIXTURE_TICKET_ID));
  FEVER_FIXTURE_CODES.forEach(function (code) {
    assert.ok(parsed.description.includes(code));
  });

  assert.equal(Object.prototype.hasOwnProperty.call(parsed, 'ticketCodes'), false);
});

test('parseFeverTicketText: a body missing the purchase-details marker throws with the full raw text', () => {
  const body = REAL_FEVER_BODY_TEXT.replace('Tady jsou podrobnosti o tvém nákupu', '');
  assert.throws(
    () => parseFeverTicketText(body, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(err.message.includes(body));
      return true;
    }
  );
});

test('parseFeverTicketText: a body missing the map-link marker throws with the full raw text', () => {
  const body = REAL_FEVER_BODY_TEXT.replace('Zobrazit na mapě', '');
  assert.throws(
    () => parseFeverTicketText(body, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(err.message.includes(body));
      return true;
    }
  );
});

test('parseFeverTicketText: a body with no recognizable date/time throws with the full raw text', () => {
  const body = REAL_FEVER_BODY_TEXT.replace(FEVER_FIXTURE_DATE_LINE, 'termín bude upřesněn');
  assert.throws(
    () => parseFeverTicketText(body, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(err.message.includes(body));
      return true;
    }
  );
});

test('parseFeverTicketText: an unrecognized month abbreviation throws, naming the bad token, with the full raw text (D-04)', () => {
  const body = buildFeverFixture({ dateLine: 'so 19 xxx - 08:00 odp.' });
  assert.throws(
    () => parseFeverTicketText(body, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(err.message.includes('xxx'));
      assert.ok(err.message.includes(body));
      return true;
    }
  );
});

test('parseFeverTicketText: hour and minute out-of-range each throw with the full raw text, including an hour above 12 written WITH a day-period marker (D-05)', () => {
  const noMeridiemHourBody = buildFeverFixture({ dateLine: 'so 19 pro - 25:00' });
  assert.throws(
    () => parseFeverTicketText(noMeridiemHourBody, FEVER_REFERENCE_DATE),
    (err) => {
      assert.match(err.message, /Hour out of range \(0-23\)/);
      assert.ok(err.message.includes(noMeridiemHourBody));
      return true;
    }
  );

  const minuteBody = buildFeverFixture({ dateLine: 'so 19 pro - 08:75 odp.' });
  assert.throws(
    () => parseFeverTicketText(minuteBody, FEVER_REFERENCE_DATE),
    (err) => {
      assert.match(err.message, /Minute out of range \(0-59\)/);
      assert.ok(err.message.includes(minuteBody));
      return true;
    }
  );

  const meridiemHourBody = buildFeverFixture({ dateLine: 'so 19 pro - 15:00 odp.' });
  assert.throws(
    () => parseFeverTicketText(meridiemHourBody, FEVER_REFERENCE_DATE),
    (err) => {
      assert.match(err.message, /Hour out of range \(1-12\)/);
      assert.ok(err.message.includes(meridiemHourBody));
      return true;
    }
  );
});

test('parseFeverTicketText: a region yielding no usable event name throws with the full raw text; a region yielding no usable location throws with the full raw text', () => {
  const noEventNameBody = [
    'Zobrazit v prohlížeči',
    'Tady jsou podrobnosti o tvém nákupu',
    '',
    'Koupit znova',
    '',
    'Zobrazit na mapě',
    FEVER_FIXTURE_DATE_LINE,
  ].join('\n');
  assert.throws(
    () => parseFeverTicketText(noEventNameBody, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(err.message.includes(noEventNameBody));
      return true;
    }
  );

  const noLocationBody = [
    'Zobrazit v prohlížeči',
    'Tady jsou podrobnosti o tvém nákupu',
    FEVER_FIXTURE_EVENT_NAME,
    'Zobrazit na mapě',
    FEVER_FIXTURE_DATE_LINE,
  ].join('\n');
  assert.throws(
    () => parseFeverTicketText(noLocationBody, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(err.message.includes(noLocationBody));
      return true;
    }
  );
});

test('parseFeverTicketText: TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL["hello@feverup.com"] is strictly equal to the exported parseFeverTicketText (D-01)', () => {
  // Guarded with the typeof check so this assertion genuinely fails before
  // Task 2 (both sides would otherwise be `undefined` and trivially equal).
  assert.equal(typeof parseFeverTicketText, 'function');
  assert.equal(TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL['hello@feverup.com'], parseFeverTicketText);
});

test('Fever ABSENCE PROOF -- no key in TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL, and no key in TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL (D-01/D-11)', () => {
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL, 'hello@feverup.com'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL, 'hello@feverup.com'), false);
});

const FEVER_MARKETING_BODY = 'Podívej se na nové akce ve tvém městě! Nenech si ujít nadcházející koncerty a festivaly.';

test('feverTextHasPurchaseDetails: registered strictly, returns true on the real fixture, false on Fever marketing mail, and false (never throwing) on null/empty (D-08)', () => {
  assert.equal(TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL['hello@feverup.com'], feverTextHasPurchaseDetails);

  assert.equal(feverTextHasPurchaseDetails(REAL_FEVER_BODY_TEXT), true);
  assert.equal(feverTextHasPurchaseDetails(FEVER_MARKETING_BODY), false);
  assert.equal(feverTextHasPurchaseDetails(null), false);
  assert.equal(feverTextHasPurchaseDetails(''), false);
});

test('resolveTicketProcessingJobs / appliesTo: a Fever purchase confirmation yields exactly one mode:"body" job and applies; a Fever marketing email yields ZERO jobs and does not apply (D-08)', () => {
  const portals = [{ identifyingEmail: 'hello@feverup.com', calendarId: 'FEVER_CAL', insertPdfIntoEvent: false }];

  const purchaseMessage = fakeMessage('Fever <hello@feverup.com>', [], REAL_FEVER_BODY_TEXT);
  const jobs = resolveTicketProcessingJobs([purchaseMessage], portals);
  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
  assert.equal(jobs[0].message, purchaseMessage);
  assert.equal(jobs[0].portal, portals[0]);

  const marketingMessage = fakeMessage('Fever <hello@feverup.com>', [], FEVER_MARKETING_BODY);
  assert.equal(resolveTicketProcessingJobs([marketingMessage], portals).length, 0);

  // appliesTo consults the REAL shipped TICKETING_PORTALS_ACTION_CONFIG
  // default (D-15 seeds the Fever entry there), not the local `portals`
  // array above.
  const purchaseThread = {
    getMessages: function () {
      return [purchaseMessage];
    },
  };
  const marketingThread = {
    getMessages: function () {
      return [marketingMessage];
    },
  };
  assert.equal(TICKETING_PORTALS_ACTION.appliesTo(purchaseThread), true);
  assert.equal(TICKETING_PORTALS_ACTION.appliesTo(marketingThread), false);
});

test('resolveTicketProcessingJobs: a Fever message carrying the real PDF attachment still yields exactly one mode:"body" job -- never one job per attachment, never a pdf-mode job (D-01)', () => {
  const portals = [{ identifyingEmail: 'hello@feverup.com', calendarId: 'FEVER_CAL', insertPdfIntoEvent: true }];
  const pdfAttachment = fakeAttachment('order_998877665.pdf', 'application/pdf');
  const message = fakeMessage('Fever <hello@feverup.com>', [pdfAttachment], REAL_FEVER_BODY_TEXT);

  const jobs = resolveTicketProcessingJobs([message], portals);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].mode, 'body');
});

test('findFeverTicketPdfAttachment: returns the order PDF, null on an unrelated PDF name, null on a matching name that is not a PDF content type, and is registered strictly (D-12)', () => {
  const orderPdf = fakeAttachment('order_998877665.pdf', 'application/pdf');
  const message = fakeMessage('Fever <hello@feverup.com>', [orderPdf], REAL_FEVER_BODY_TEXT);
  assert.equal(findFeverTicketPdfAttachment(message), orderPdf);

  const unrelatedPdf = fakeAttachment('terms.pdf', 'application/pdf');
  const messageUnrelated = fakeMessage('Fever <hello@feverup.com>', [unrelatedPdf], REAL_FEVER_BODY_TEXT);
  assert.equal(findFeverTicketPdfAttachment(messageUnrelated), null);

  const nonPdfNamedOrder = fakeAttachment('order_998877665.txt', 'text/plain');
  const messageNonPdf = fakeMessage('Fever <hello@feverup.com>', [nonPdfNamedOrder], REAL_FEVER_BODY_TEXT);
  assert.equal(findFeverTicketPdfAttachment(messageNonPdf), null);

  assert.equal(TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL['hello@feverup.com'], findFeverTicketPdfAttachment);
});

// `subject` (quick-260921-gj0 round 2) is an OPTIONAL second parameter
// exposed as `getSubject()`, defaulting to `undefined` -- unchanged
// round-1 callers that invoke this factory with only `receivedDate` get a
// message whose subject is missing, which is exactly the D-24 fallback
// condition and keeps every round-1 assertion against this factory valid.
function feverBodyModeMessage(receivedDate, subject) {
  return {
    getFrom: function () {
      return 'Fever <hello@feverup.com>';
    },
    getPlainBody: function () {
      return REAL_FEVER_BODY_TEXT;
    },
    getBody: function () {
      return '<html><body>irrelevant</body></html>';
    },
    getAttachments: function () {
      return [];
    },
    getDate: function () {
      return receivedDate;
    },
    getSubject: function () {
      return subject;
    },
  };
}

test('processTicketFromMessageBody: passes message.getDate() through to the registered body parser as the year-inference reference date (D-07); the existing Kino Art and Entradio harness paths still produce their unchanged resources, proving the extra argument is inert for them', () => {
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(feverBodyModeMessage(FEVER_REFERENCE_DATE), {
      identifyingEmail: 'hello@feverup.com',
      calendarId: 'FEVER_CAL',
      insertPdfIntoEvent: false,
    });

    assert.equal(calls.inserted.length, 1);
    assert.equal(calls.inserted[0].resource.start.dateTime.slice(0, 4), '2026');
    assert.equal(calls.inserted[0].resource.summary, FEVER_FIXTURE_EVENT_NAME);
  });

  // Companion assertion: the existing Kino Art and Entradio harness paths are
  // untouched by this change -- same assertions their own dedicated tests
  // above already make, re-run here to prove the extra getDate() argument is
  // inert for parsers that do not read it.
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(kinoArtBodyModeMessage(), {
      identifyingEmail: 'rezervace@kinoart.cz',
      calendarId: 'ART_CAL',
      insertPdfIntoEvent: true,
    });
    assert.equal(calls.inserted[0].resource.attachments.length, 1);
    assert.equal(calls.inserted[0].resource.attachments[0].mimeType, 'application/pdf');
  });

  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(entradioBodyModeMessage(), ENTRADIO_PORTAL_PDF_ON);
    assert.equal(calls.inserted[0].resource.summary, 'ČERNO, VÍR');
    assert.equal(calls.inserted[0].resource.attachments.length, 3);
  });
});

test('TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals: the shipped default fifth entry deepEquals the Fever entry, and resolveTicketingCalendarId on it returns the passed global default (D-15)', () => {
  const portals = TICKETING_PORTALS_ACTION.config.ticketingPortals;
  assert.deepEqual(portals[4], { identifyingEmail: 'hello@feverup.com', calendarId: null, insertPdfIntoEvent: false });
  assert.equal(resolveTicketingCalendarId(portals[4], 'GLOBAL_DEFAULT_CAL'), 'GLOBAL_DEFAULT_CAL');
});

test('TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals: regression guard -- entries 0-3 (enigoo.cz, Kino Art, Ticketmaster CZ, Entradio) are unchanged and the array length is 5', () => {
  const portals = TICKETING_PORTALS_ACTION.config.ticketingPortals;
  assert.equal(portals.length, 5);
  assert.deepEqual(portals[0], { identifyingEmail: 'no-reply@enigoo.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.deepEqual(portals[1], { identifyingEmail: 'rezervace@kinoart.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.deepEqual(portals[2], { identifyingEmail: 'noreply@ticketmaster.cz', calendarId: null, insertPdfIntoEvent: false });
  assert.deepEqual(portals[3], { identifyingEmail: 'no-reply@app.entradio.cz', calendarId: null, insertPdfIntoEvent: false });
});

// --- ROUND 2 (live-test-driven): subject-sourced event name, plus a Gmail --
// --- image-placeholder skip in the body fallback -----------------------------
//
// These come from the owner's FIRST live run against a real Fever email
// (Task 3), plus the message's own Subject header decoded from that same
// real sample.
//
// Gmail rendered an <img> in the plain-text view as a bracketed placeholder
// line ("[image: <alt text>]") sitting BETWEEN the purchase-details marker
// and the real event-name line -- the first direct observation of what
// message.getPlainBody() actually returns for a Fever email, and it became
// the created event's summary verbatim (D-18). feverFirstNonEmptyLine and
// feverLastNonEmptyLine now also skip such a line (D-19), narrowly -- a line
// merely CONTAINING bracketed text stays eligible (D-20).
//
// Separately, the real Subject header decodes to a fixed template prefix
// followed by the event name VERBATIM, including the event name's own
// internal colon -- so the subject is now the PRIMARY event-name source
// (matched as a literal prefix, never split on a colon) and the
// placeholder-hardened body read above becomes the FALLBACK (D-24). The
// location stays body-sourced always (D-26).
//
// Neither the owner's real event name, venue, subject, ticket ID nor seat
// codes appear here -- fictional values only, per this suite's existing
// FIXTURE PROVENANCE note above.

// The real prefix, decoded from the real sample's own RFC 2047 subject
// header (D-24). Reused everywhere below, exactly like FEVER_SUBJECT_PREFIX
// inside the source file.
const FEVER_FIXTURE_SUBJECT_PREFIX = 'Potvrzení nákupu na Fever: ';
const FEVER_FIXTURE_SUBJECT = FEVER_FIXTURE_SUBJECT_PREFIX + FEVER_FIXTURE_EVENT_NAME;

// A fictional Gmail image-placeholder line, in the exact bracketed shape
// directly observed on the owner's live run (D-18).
const FEVER_FIXTURE_IMAGE_PLACEHOLDER = '[image: Event cover photo]';

test('parseFeverTicketText: THE LIVE REGRESSION -- a Gmail image-placeholder line rendered ABOVE the real event-name line is skipped, never returned as the event name (D-18)', () => {
  const body = buildFeverFixture({ eventName: FEVER_FIXTURE_IMAGE_PLACEHOLDER + '\n' + FEVER_FIXTURE_EVENT_NAME });
  const parsed = parseFeverTicketText(body, FEVER_REFERENCE_DATE);

  assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.equal(parsed.location, FEVER_FIXTURE_LOCATION);
  assert.equal(parsed.description.split('\n\n')[0], FEVER_FIXTURE_EVENT_NAME);
});

test('parseFeverTicketText: the LOCATION half of the same defect -- a placeholder line rendered between the venue and the map-link marker is skipped, never returned as the location (D-19)', () => {
  const body = buildFeverFixture({ location: FEVER_FIXTURE_LOCATION + '\n' + FEVER_FIXTURE_IMAGE_PLACEHOLDER });
  const parsed = parseFeverTicketText(body, FEVER_REFERENCE_DATE);

  assert.equal(parsed.location, FEVER_FIXTURE_LOCATION);
  assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
});

test('parseFeverTicketText: the placeholder skip tolerates spacing/casing variance and two placeholders on one line, but a line that merely BEGINS with unrelated bracketed text is still returned intact -- narrow skip, not a blanket bracket filter (D-19/D-20)', () => {
  const noSpace = parseFeverTicketText(
    buildFeverFixture({ eventName: '[image:Event cover photo]\n' + FEVER_FIXTURE_EVENT_NAME }),
    FEVER_REFERENCE_DATE
  );
  assert.equal(noSpace.eventName, FEVER_FIXTURE_EVENT_NAME);

  const differentCase = parseFeverTicketText(
    buildFeverFixture({ eventName: '[IMAGE: Event cover photo]\n' + FEVER_FIXTURE_EVENT_NAME }),
    FEVER_REFERENCE_DATE
  );
  assert.equal(differentCase.eventName, FEVER_FIXTURE_EVENT_NAME);

  const twoOnOneLine = parseFeverTicketText(
    buildFeverFixture({ eventName: '[image: icon one][image: icon two]\n' + FEVER_FIXTURE_EVENT_NAME }),
    FEVER_REFERENCE_DATE
  );
  assert.equal(twoOnOneLine.eventName, FEVER_FIXTURE_EVENT_NAME);

  // NEGATIVE (D-20): a line that merely BEGINS with bracketed text, but is
  // not an image-placeholder shape, is still fully eligible as the event
  // name -- proving the skip is narrow rather than a blanket bracket filter.
  const bracketedButReal = '[VIP] ' + FEVER_FIXTURE_EVENT_NAME;
  const notAPlaceholder = parseFeverTicketText(buildFeverFixture({ eventName: bracketedButReal }), FEVER_REFERENCE_DATE);
  assert.equal(notAPlaceholder.eventName, bracketedButReal);
});

test('parseFeverTicketText: a region containing NOTHING but placeholder lines still raises the EXISTING event-name throw carrying the full raw text -- degrades to a diagnostic email, never to a silently-wrong summary (D-20)', () => {
  const body = buildFeverFixture({
    eventName: FEVER_FIXTURE_IMAGE_PLACEHOLDER,
    location: FEVER_FIXTURE_IMAGE_PLACEHOLDER,
  });

  assert.throws(
    () => parseFeverTicketText(body, FEVER_REFERENCE_DATE),
    (err) => {
      assert.ok(
        err.message.includes('could not extract the event name between the purchase-details marker and the map-link marker'),
        'error message should be the EXISTING event-name-throw wording, unreworded'
      );
      assert.ok(err.message.includes(body), 'error message should include the full raw text');
      return true;
    }
  );
});

test('parseFeverTicketText: SUBJECT WINS, AND WINS ON PURPOSE -- a real Fever subject overrides the body-derived event name even when the body would produce a different, equally plausible name, and the internal colon survives in full (D-24)', () => {
  const differentCleanEventName = 'Novoroční ohňostroj: Staroměstské náměstí';
  const subject = FEVER_FIXTURE_SUBJECT_PREFIX + differentCleanEventName;
  const body = buildFeverFixture({}); // body's own event name is FEVER_FIXTURE_EVENT_NAME -- deliberately different from `subject`'s

  const parsed = parseFeverTicketText(body, FEVER_REFERENCE_DATE, subject);

  assert.equal(parsed.eventName, differentCleanEventName);
  assert.notEqual(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.ok(parsed.eventName.includes(':'), 'the event name\'s own internal colon must survive -- proves prefix-strip, not colon-split');
  assert.equal(parsed.description.split('\n\n')[0], differentCleanEventName);
});

test('parseFeverTicketText: FALLBACK CONDITIONS -- a subject that is missing entirely, undefined, null, a non-string, empty, prefix-less (a plausible marketing subject), or the prefix alone with nothing after it all fall back to the BODY event name WITHOUT throwing (D-24/D-25)', () => {
  const body = buildFeverFixture({});

  // No third argument at all -- this IS today's only call shape and must
  // keep working unchanged (D-25's round-1 compatibility guarantee).
  assert.equal(parseFeverTicketText(body, FEVER_REFERENCE_DATE).eventName, FEVER_FIXTURE_EVENT_NAME);

  [
    undefined,
    null,
    12345,
    '',
    'Podívej se na nové akce ve tvém městě!', // plausible Fever MARKETING subject, no prefix match
    FEVER_FIXTURE_SUBJECT_PREFIX, // the prefix alone, nothing after it
  ].forEach(function (subject) {
    const parsed = parseFeverTicketText(body, FEVER_REFERENCE_DATE, subject);
    assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
  });
});

test('parseFeverTicketText: THE REAL PRODUCTION SHAPE -- a body carrying the placeholder-above-the-event-name defect AND a real subject together resolve to the SUBJECT-derived name, proving the two halves of this round compose rather than fight (D-19/D-24)', () => {
  const bodyOnlyEventName = 'Zimní trhy na náměstí';
  const body = buildFeverFixture({ eventName: FEVER_FIXTURE_IMAGE_PLACEHOLDER + '\n' + bodyOnlyEventName });

  const parsed = parseFeverTicketText(body, FEVER_REFERENCE_DATE, FEVER_FIXTURE_SUBJECT);

  assert.equal(parsed.eventName, FEVER_FIXTURE_EVENT_NAME);
  assert.notEqual(parsed.eventName, bodyOnlyEventName);
  assert.equal(parsed.location, FEVER_FIXTURE_LOCATION);
});

test('processTicketFromMessageBody: passes message.getSubject() through to the registered body parser as the THIRD argument (D-25/D-27) -- the inserted Calendar resource\'s summary is the SUBJECT-derived event name, not the body\'s, and the start still carries the year round 1\'s D-07 threading implies; the Kino Art and Entradio harness paths stay unchanged, proving the third argument is inert for them too', () => {
  const differentSubjectEventName = 'Vítání jara: Karlův most';
  const subject = FEVER_FIXTURE_SUBJECT_PREFIX + differentSubjectEventName;

  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(feverBodyModeMessage(FEVER_REFERENCE_DATE, subject), {
      identifyingEmail: 'hello@feverup.com',
      calendarId: 'FEVER_CAL',
      insertPdfIntoEvent: false,
    });

    assert.equal(calls.inserted.length, 1);
    assert.equal(calls.inserted[0].resource.summary, differentSubjectEventName);
    assert.notEqual(calls.inserted[0].resource.summary, FEVER_FIXTURE_EVENT_NAME);
    assert.equal(calls.inserted[0].resource.start.dateTime.slice(0, 4), '2026');
  });

  // Companion assertion: the existing Kino Art and Entradio harness paths are
  // untouched by this change -- same assertions their own dedicated tests
  // above already make, re-run here to prove the third argument is inert for
  // parsers that do not read it, exactly as round 1 already proved for the
  // second one.
  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(kinoArtBodyModeMessage(), {
      identifyingEmail: 'rezervace@kinoart.cz',
      calendarId: 'ART_CAL',
      insertPdfIntoEvent: true,
    });
    assert.equal(calls.inserted[0].resource.attachments.length, 1);
    assert.equal(calls.inserted[0].resource.attachments[0].mimeType, 'application/pdf');
  });

  withTicketBodyRunGlobals({ respond: entradioRespondOk }, function (calls) {
    processTicketFromMessageBody(entradioBodyModeMessage(), ENTRADIO_PORTAL_PDF_ON);
    assert.equal(calls.inserted[0].resource.summary, 'ČERNO, VÍR');
    assert.equal(calls.inserted[0].resource.attachments.length, 3);
  });
});

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  BOOKING_LANGUAGE_PACKS,
  getBookingLabels,
  bookingExtractEmailAddress,
  bookingIsAllowedSender,
  subjectContainsAny,
  extractLabeledNumber,
  parse12HourTime,
  parseBookingDateLine,
  findLabeledLine,
  extractCheckInOutDate,
  extractHotelName,
  nextNonEmptyLineAfterLabel,
  computeSearchWindow,
  eventDateRangeOverlaps,
  eventSummaryOrLocationContainsHotelName,
  findFuzzyMatchingEvent,
  formatLocalWallClockIso,
  extractReservationDetailsSection,
  findLabeledLineAcrossLanguagePacks,
  extractCheckInOutDateAcrossLanguagePacks,
} = require('../src/06-action-booking-com-management.js');
const { BOOKING_ACTION_CONFIG } = require('../src/06-action-cfg-booking-com-management.js');
const { EN_LANGUAGE_PACK } = require('../src/06-lang-en.js');
const { CS_LANGUAGE_PACK, parseCzechBookingDateLine } = require('../src/06-lang-cs.js');

// --- Fixtures ----------------------------------------------------------
//
// Minimal Node fixture strings built from the real example field values
// (NOT full raw MIME) — shared across this file's batches.

const CONFIRMATION_SUBJECT = '🛄 Thanks! Your booking is confirmed at Hotel U Modré hvězdy';
const CANCELLATION_SUBJECT = 'Booking canceled for Hotel U Modré hvězdy';

const CONFIRMATION_BODY = [
  'Confirmation: 8471602593',
  'PIN: 6194 (Confidential)',
  '',
  'Check-in       Friday, September 11, 2026 (2:00 PM - 8:00 PM)',
  'Check-out      Sunday, September 13, 2026 (7:00 AM - 10:00 AM)',
  '',
  'Location',
  'Horni namesti 21, Olomouc, 77900, Czech Republic',
].join('\n');

const CANCELLATION_BODY = [
  'Confirmation Number: 8471602593',
  'PIN code: 6194',
  '',
  'Booking number 8471602593',
  '',
  'Check-in       Friday, September 11, 2026',
].join('\n');

// Real verbatim "Reservation details" section text decoded from the real
// confirmation .eml's plain-text body, leading whitespace preserved
// exactly as booking.com sends it — used to test extractReservationDetailsSection.
const CONFIRMATION_BODY_WITH_RESERVATION_DETAILS = [
  'Confirmation: 8471602593',
  'PIN: 6194 (Confidential)',
  '',
  'Reservation details',
  '',
  '   Check-in Friday, September 11, 2026 (2:00 PM - 8:00 PM)',
  '   Check-out Sunday, September 13, 2026 (7:00 AM - 10:00 AM)',
  '   Your reservation 2 nights, Double or Twin Room',
  '   You booked for 1 adult',
  '   Location',
  '',
  '',
  '    Horni namesti 21, Olomouc, 77900, Czech Republic',
  '',
  '   Phone +420 608 314 927',
  '   Contact Email property',
  '   Cancellation policy',
  "   You can cancel for free until 3 days before arrival. If you cancel",
  '   within 3 days of arrival, the cancellation fee will be the total price',
  "   of the reservation. If you don't show up, the no-show fee will be the",
  '   total price of the reservation.',
  '   Cancellation cost',
  '     * until September 7, 2026 11:59 PM: 0 Kč',
  '     * from September 8, 2026 12:00 AM: 3,360 Kč',
  '',
  "   Cancellation deadlines are in the property's local time.",
  '',
  '                                Price details',
].join('\n');

const CONFIRMATION_NUMBER_LABELS = ['Confirmation Number', 'Confirmation', 'Booking number', 'Confirmation number'];
const PIN_LABELS = ['PIN code', 'PIN Code', 'PIN'];
const HOTEL_NAME = 'Hotel U Modré hvězdy';

// --- Czech fixtures -----------------------------------------------------
//
// Real verbatim text decoded from two real booking.com emails (confirmation
// + cancellation, "Hotel Nikol - Free On-site Parking", reservation
// 6203948175), leading whitespace preserved exactly as booking.com sends
// it. First real-world exercise of the language-pack architecture — Czech's
// date/time format and hotel-name subject position turned out to be
// STRUCTURALLY different from English (see src/06-lang-cs.js), not just
// lexically different.

const CZ_HOTEL_NAME = 'Hotel Nikol - Free On-site Parking';

const CZ_CONFIRMATION_SUBJECT = '🛄 Hotel Nikol - Free On-site Parking – Děkujeme! Vaše rezervace je potvrzena';
const CZ_CANCELLATION_SUBJECT = 'Zrušení rezervace – Hotel Nikol - Free On-site Parking';

const CZ_CONFIRMATION_BODY = [
  '   Booking.com',
  '   Potvrzení rezervace: 6203948175',
  '   PIN kód: 9885 (Důvěrné)',
  '',
  'Děkujeme! Vaše rezervace v destinaci Olomouc je potvrzena.',
  '',
  'Informace o rezervaci',
  '',
  '   Příjezd pátek 11. září 2026 (od 14:00)',
  '   Odjezd neděle 13. září 2026 (do 10:00)',
  '   Vaše rezervace 2 noci, Dvoulůžkový pokoj Deluxe s manželskou postelí',
  '   nebo oddělenými postelemi',
  '   Rezervace pro: 2 dospělí',
  '   Místo',
  '',
  '',
  '    Wellnerova 318, Olomouc, 77900, Česká republika',
  '',
  '   Telefon +420 733 219 481',
  '   Kontakt Pošlete do ubytování e-mail',
  '   Podmínky zrušení rezervace',
  '   Rezervaci můžete zrušit zdarma do 7 dní před příjezdem. V případě',
  '   zrušení rezervace méně než 7 dní před příjezdem bude storno poplatek',
  '   činit 100 % ceny první noci. Pokud se k pobytu nedostavíte, bude',
  '   poplatek za nedojezd činit 100 % celkové ceny.',
  '   Poplatek za zrušení rezervace',
  '     * do 3. září 2026 23:59: 0 Kč',
  '     * od 4. září 2026 0:00: 2 871 Kč',
  '     * od 12. září 2026 0:00: 5 742 Kč',
  '',
  '   Termíny pro zrušení rezervace odpovídají časovému pásmu ubytování.',
  '',
  '                              Informace o ceně',
  '',
  '   1 Dvoulůžkový pokoj Deluxe s manželskou postelí nebo oddělenými',
  '   postelemi',
  '   5 126,79 Kč',
].join('\n');

const CZ_CANCELLATION_BODY = [
  '   Booking.com',
  '',
  '   Číslo rezervace: 6203948175',
  '',
  '   Kód PIN: 9885',
  '',
  'Vaše rezervace byla úspěšně zdarma zrušena',
  '',
  '   Vážená paní / Vážený pane,',
  '',
  '   Rádi bychom Vás informovali, že Vaše rezervace (Hotel Nikol - Free',
  '   On-site Parking) byla zrušena.',
  '   Hotel Nikol - Free On-site Parking',
  '',
  '',
  '    Wellnerova 318, Olomouc, 77900, Česká republika',
  '',
  '   ZRUŠENO',
  '',
  '',
  '    Telefon: +420 733 219 481',
  '',
  '   Pošlete do ubytování e-mail',
  '   Rezervace pro:  2 dospělí',
  '   Příjezd         pátek, 11. září 2026',
  '   Odjezd          neděle, 13. září 2026',
  '   Číslo rezervace 6203948175',
  '   Kód PIN         9885',
  '   Zrušení zdarma CZK 0',
].join('\n');

// A real fromGmail-type auto-created "stay" event, confirmed via live
// Calendar.Events.get() dumps: description is a useless Google-generic
// placeholder ("To see detailed information... use the official Google
// Calendar app"), never the real booking details — but summary/location/
// start/end ARE readable and correct. end.date is EXCLUSIVE (the day after
// the last night), same convention already used elsewhere in this codebase
// for all-day events.
const FROM_GMAIL_STAY_EVENT = {
  id: 'fromgmail-event-id',
  summary: 'Stay at ' + HOTEL_NAME,
  description: 'To see detailed information for this event, use the official Google Calendar app.',
  location: 'Horni namesti 21, Olomouc, 77900, Czech Republic',
  start: { date: '2026-09-11' },
  end: { date: '2026-09-14' },
};

// --- bookingExtractEmailAddress (LOCAL copy) -----------------------------
//
// Renamed from a plain extractEmailAddress after a live-reported bug: GAS
// concatenates every project file into one shared global scope, so this
// name collided at runtime with src/05-action-ics-import.js's identically
// named helper (the Apps Script trigger-picker surfaced this as a "same
// name... undefined behaviour" warning). Same logic, name-spaced name.

test('bookingExtractEmailAddress: angle-bracket header returns the bare lowercased address', () => {
  assert.equal(bookingExtractEmailAddress('Booking.com <noreply@booking.com>'), 'noreply@booking.com');
});

test('bookingExtractEmailAddress: bare address (no display name) returns itself', () => {
  assert.equal(bookingExtractEmailAddress('noreply@booking.com'), 'noreply@booking.com');
});

test('bookingExtractEmailAddress: mixed-case input is lowercased', () => {
  assert.equal(bookingExtractEmailAddress('Foo <Foo.Bar@Booking.COM>'), 'foo.bar@booking.com');
});

test('bookingExtractEmailAddress: null/undefined/empty return empty string without throwing', () => {
  assert.equal(bookingExtractEmailAddress(null), '');
  assert.equal(bookingExtractEmailAddress(undefined), '');
  assert.equal(bookingExtractEmailAddress(''), '');
});

// --- bookingIsAllowedSender (LOCAL copy) ----------------------------------
//
// Renamed from a plain isAllowedSender for the same real cross-file
// naming-collision reason as bookingExtractEmailAddress above.

test('bookingIsAllowedSender: empty/null/undefined list allows any sender', () => {
  assert.equal(bookingIsAllowedSender('anyone@x.com', []), true);
  assert.equal(bookingIsAllowedSender('anyone@x.com', null), true);
  assert.equal(bookingIsAllowedSender('anyone@x.com', undefined), true);
});

test('bookingIsAllowedSender: bare list entry match', () => {
  assert.equal(bookingIsAllowedSender('noreply@booking.com', ['noreply@booking.com']), true);
});

test('bookingIsAllowedSender: full "Name <email>" list entry match', () => {
  assert.equal(bookingIsAllowedSender('noreply@booking.com', ['Booking.com <noreply@booking.com>']), true);
});

test('bookingIsAllowedSender: no match returns false', () => {
  assert.equal(bookingIsAllowedSender('bob@other.com', ['noreply@booking.com']), false);
});

test('bookingIsAllowedSender: case-insensitive match', () => {
  assert.equal(bookingIsAllowedSender('NOREPLY@BOOKING.COM', ['noreply@booking.com']), true);
});

test('bookingIsAllowedSender: default allow-list value passes the default sender', () => {
  assert.equal(bookingIsAllowedSender('noreply@booking.com', ['noreply@booking.com']), true);
});

// --- subjectContainsAny ---------------------------------------------------

test('subjectContainsAny: case-insensitive match on the confirmation subject', () => {
  assert.equal(subjectContainsAny(CONFIRMATION_SUBJECT, ['booking is confirmed']), true);
});

test('subjectContainsAny: case-insensitive match on the cancellation subject', () => {
  assert.equal(subjectContainsAny(CANCELLATION_SUBJECT, ['Booking canceled for']), true);
});

test('subjectContainsAny: non-matching subject returns false', () => {
  assert.equal(subjectContainsAny('Some unrelated subject', ['booking is confirmed']), false);
});

test('subjectContainsAny: empty/null substrings list returns false', () => {
  assert.equal(subjectContainsAny(CONFIRMATION_SUBJECT, []), false);
  assert.equal(subjectContainsAny(CONFIRMATION_SUBJECT, null), false);
});

// --- extractLabeledNumber -------------------------------------------------

test('extractLabeledNumber: cancellation body matches via "Confirmation Number" (first present label)', () => {
  assert.equal(extractLabeledNumber(CANCELLATION_BODY, CONFIRMATION_NUMBER_LABELS), '8471602593');
});

test('extractLabeledNumber: confirmation body (no "Confirmation Number") matches via "Confirmation"', () => {
  assert.equal(extractLabeledNumber(CONFIRMATION_BODY, CONFIRMATION_NUMBER_LABELS), '8471602593');
});

test('extractLabeledNumber: first label in list order wins, and "Confirmation" does not falsely match inside "Confirmation Number:"', () => {
  const synthetic = 'Confirmation Number: 111\nsome other line\nConfirmation: 222';
  assert.equal(extractLabeledNumber(synthetic, CONFIRMATION_NUMBER_LABELS), '111');
});

test('extractLabeledNumber: "Confirmation" label alone does not falsely match inside "Confirmation Number:" (immediate-colon requirement)', () => {
  assert.equal(extractLabeledNumber('Confirmation Number: 111', ['Confirmation']), null);
});

test('extractLabeledNumber: no labeled number present returns null', () => {
  assert.equal(extractLabeledNumber('nothing relevant here', CONFIRMATION_NUMBER_LABELS), null);
});

test('extractLabeledNumber: PIN extraction from the confirmation body', () => {
  assert.equal(extractLabeledNumber(CONFIRMATION_BODY, PIN_LABELS), '6194');
});

test('extractLabeledNumber: PIN extraction from the cancellation body', () => {
  assert.equal(extractLabeledNumber(CANCELLATION_BODY, PIN_LABELS), '6194');
});

// --- extractLabeledNumber (Czech) -------------------------------------------
//
// Czech's confirmation-number/PIN labels differ between the confirmation
// email ('Potvrzení rezervace' / 'PIN kód') and the cancellation email
// ('Číslo rezervace' / 'Kód PIN') — both variants belong in one flat list,
// same reasoning as English's multiple confirmationNumberLabels entries.

test('extractLabeledNumber: Czech confirmation body matches via "Potvrzení rezervace"', () => {
  assert.equal(extractLabeledNumber(CZ_CONFIRMATION_BODY, CS_LANGUAGE_PACK.confirmationNumberLabels), '6203948175');
});

test('extractLabeledNumber: Czech cancellation body matches via "Číslo rezervace" (different wording, same field)', () => {
  assert.equal(extractLabeledNumber(CZ_CANCELLATION_BODY, CS_LANGUAGE_PACK.confirmationNumberLabels), '6203948175');
});

test('extractLabeledNumber: Czech confirmation body PIN matches via "PIN kód"', () => {
  assert.equal(extractLabeledNumber(CZ_CONFIRMATION_BODY, CS_LANGUAGE_PACK.pinLabels), '9885');
});

test('extractLabeledNumber: Czech cancellation body PIN matches via "Kód PIN" (different word order, same field)', () => {
  assert.equal(extractLabeledNumber(CZ_CANCELLATION_BODY, CS_LANGUAGE_PACK.pinLabels), '9885');
});

// --- subjectContainsAny (Czech) ----------------------------------------------

test('subjectContainsAny: Czech confirmation subject matches "rezervace je potvrzena"', () => {
  assert.equal(subjectContainsAny(CZ_CONFIRMATION_SUBJECT, CS_LANGUAGE_PACK.addToCalendarSubjectContains), true);
});

test('subjectContainsAny: Czech cancellation subject matches "Zrušení rezervace"', () => {
  assert.equal(subjectContainsAny(CZ_CANCELLATION_SUBJECT, CS_LANGUAGE_PACK.removeFromCalendarSubjectContains), true);
});

// --- parse12HourTime -------------------------------------------------------

test('parse12HourTime: "2:00 PM" -> 14:00', () => {
  assert.deepEqual(parse12HourTime('2:00 PM'), { hour: 14, minute: 0 });
});

test('parse12HourTime: "8:00 PM" -> 20:00', () => {
  assert.deepEqual(parse12HourTime('8:00 PM'), { hour: 20, minute: 0 });
});

test('parse12HourTime: "7:00 AM" -> 07:00', () => {
  assert.deepEqual(parse12HourTime('7:00 AM'), { hour: 7, minute: 0 });
});

test('parse12HourTime: "10:00 AM" -> 10:00', () => {
  assert.deepEqual(parse12HourTime('10:00 AM'), { hour: 10, minute: 0 });
});

test('parse12HourTime: "12:00 AM" -> midnight (00:00)', () => {
  assert.deepEqual(parse12HourTime('12:00 AM'), { hour: 0, minute: 0 });
});

test('parse12HourTime: "12:00 PM" -> noon (12:00)', () => {
  assert.deepEqual(parse12HourTime('12:00 PM'), { hour: 12, minute: 0 });
});

test('parse12HourTime: "11:59 PM" -> 23:59', () => {
  assert.deepEqual(parse12HourTime('11:59 PM'), { hour: 23, minute: 59 });
});

test('parse12HourTime: malformed input throws ("banana")', () => {
  assert.throws(() => parse12HourTime('banana'));
});

test('parse12HourTime: malformed input throws (out-of-range hour "25:00 PM")', () => {
  assert.throws(() => parse12HourTime('25:00 PM'));
});

// --- parseBookingDateLine ---------------------------------------------------

test('parseBookingDateLine: confirmation check-in line value parses date + time window', () => {
  assert.deepEqual(parseBookingDateLine('Friday, September 11, 2026 (2:00 PM - 8:00 PM)'), {
    year: 2026,
    month: 8,
    day: 11,
    startTime: { hour: 14, minute: 0 },
    endTime: { hour: 20, minute: 0 },
  });
});

test('parseBookingDateLine: confirmation check-out line value parses date + time window', () => {
  assert.deepEqual(parseBookingDateLine('Sunday, September 13, 2026 (7:00 AM - 10:00 AM)'), {
    year: 2026,
    month: 8,
    day: 13,
    startTime: { hour: 7, minute: 0 },
    endTime: { hour: 10, minute: 0 },
  });
});

test('parseBookingDateLine: cancellation date-only line (no parens) parses with null time window', () => {
  assert.deepEqual(parseBookingDateLine('Friday, September 11, 2026'), {
    year: 2026,
    month: 8,
    day: 11,
    startTime: null,
    endTime: null,
  });
});

test('parseBookingDateLine: malformed date text throws', () => {
  assert.throws(() => parseBookingDateLine('not a date at all'));
});

// --- parseCzechBookingDateLine (src/06-lang-cs.js) ---------------------------
//
// Czech's date/time format is STRUCTURALLY different from English's, not
// just lexically: day number FIRST with a trailing period, then a GENITIVE
// Czech month name, no comma before the year; EACH line carries its OWN
// single "od" (from/start) or "do" (until/end) 24-hour time, never a single
// "-"-separated range on one line like English's "(2:00 PM - 8:00 PM)".

test('parseCzechBookingDateLine: check-in line with "od" time -> startTime set, endTime null', () => {
  assert.deepEqual(parseCzechBookingDateLine('Příjezd pátek 11. září 2026 (od 14:00)'), {
    year: 2026,
    month: 8,
    day: 11,
    startTime: { hour: 14, minute: 0 },
    endTime: null,
  });
});

test('parseCzechBookingDateLine: check-out line with "do" time -> endTime set, startTime null', () => {
  assert.deepEqual(parseCzechBookingDateLine('Odjezd neděle 13. září 2026 (do 10:00)'), {
    year: 2026,
    month: 8,
    day: 13,
    startTime: null,
    endTime: { hour: 10, minute: 0 },
  });
});

test('parseCzechBookingDateLine: cancellation-email comma variant, date-only, no time -> both times null', () => {
  assert.deepEqual(parseCzechBookingDateLine('Příjezd         pátek, 11. září 2026'), {
    year: 2026,
    month: 8,
    day: 11,
    startTime: null,
    endTime: null,
  });
});

test('parseCzechBookingDateLine: malformed line throws', () => {
  assert.throws(() => parseCzechBookingDateLine('not a date at all'));
});

// --- Format 2: HH:MM-HH:MM time RANGE (real second Czech email, live-reported bug quick-260731-cs2) ---
//
// A second real Czech confirmation email ("Apartmány Lipová")
// imported with the correct hotel name/date but check-in/check-out TIMES
// both came out as midnight — Czech booking.com emails use AT LEAST TWO
// different time-window formats depending on the property: Format 1
// (Hotel Nikol, above) carries ONE "od"/"do"-prefixed single value per
// line; Format 2 (below, real verbatim text from the new email) carries a
// single HH:MM-HH:MM RANGE per line instead, with NO "od"/"do" keywords
// at all — structurally identical to how the English parser already
// handles a range, not like Format 1. The real separator is an EN DASH
// (–, U+2013); a plain ASCII hyphen is also accepted, same defensive
// posture already used elsewhere in this codebase for punctuation
// variants.

test('parseCzechBookingDateLine: Format 2 range (real "Apartmány Lipová" check-in line, en dash) -> both startTime and endTime set', () => {
  assert.deepEqual(parseCzechBookingDateLine('Příjezd sobota 15. srpna 2026 (15:00–22:00)'), {
    year: 2026,
    month: 7,
    day: 15,
    startTime: { hour: 15, minute: 0 },
    endTime: { hour: 22, minute: 0 },
  });
});

test('parseCzechBookingDateLine: Format 2 range (real "Apartmány Lipová" check-out line, en dash, midnight start) -> both startTime and endTime set', () => {
  assert.deepEqual(parseCzechBookingDateLine('Odjezd neděle 16. srpna 2026 (0:00–10:00)'), {
    year: 2026,
    month: 7,
    day: 16,
    startTime: { hour: 0, minute: 0 },
    endTime: { hour: 10, minute: 0 },
  });
});

test('parseCzechBookingDateLine: Format 2 range is also recognized with a plain ASCII hyphen instead of an en dash', () => {
  assert.deepEqual(parseCzechBookingDateLine('Příjezd sobota 15. srpna 2026 (15:00-22:00)'), {
    year: 2026,
    month: 7,
    day: 15,
    startTime: { hour: 15, minute: 0 },
    endTime: { hour: 22, minute: 0 },
  });
});

test('parseCzechBookingDateLine: Format 1 (od/do, Hotel Nikol) is UNAFFECTED by Format 2 support — zero regression', () => {
  assert.deepEqual(parseCzechBookingDateLine('Příjezd pátek 11. září 2026 (od 14:00)'), {
    year: 2026,
    month: 8,
    day: 11,
    startTime: { hour: 14, minute: 0 },
    endTime: null,
  });
  assert.deepEqual(parseCzechBookingDateLine('Odjezd neděle 13. září 2026 (do 10:00)'), {
    year: 2026,
    month: 8,
    day: 13,
    startTime: null,
    endTime: { hour: 10, minute: 0 },
  });
});

// --- findLabeledLine --------------------------------------------------------

test('findLabeledLine: cancellation body, checkInLabels -> date-only value with label+whitespace stripped', () => {
  assert.equal(
    findLabeledLine(CANCELLATION_BODY, EN_LANGUAGE_PACK.checkInLabels),
    'Friday, September 11, 2026'
  );
});

test('findLabeledLine: confirmation body, checkInLabels -> value with time window', () => {
  assert.equal(
    findLabeledLine(CONFIRMATION_BODY, EN_LANGUAGE_PACK.checkInLabels),
    'Friday, September 11, 2026 (2:00 PM - 8:00 PM)'
  );
});

test('findLabeledLine: confirmation body, checkOutLabels -> value with time window', () => {
  assert.equal(
    findLabeledLine(CONFIRMATION_BODY, EN_LANGUAGE_PACK.checkOutLabels),
    'Sunday, September 13, 2026 (7:00 AM - 10:00 AM)'
  );
});

test('findLabeledLine: no matching line for the given labels returns null', () => {
  assert.equal(findLabeledLine(CANCELLATION_BODY, ['Nonexistent Label']), null);
});

// --- extractCheckInOutDate ---------------------------------------------------

test('extractCheckInOutDate: confirmation body check-in, useStartTime=true -> earliest arrival instant', () => {
  const result = extractCheckInOutDate(CONFIRMATION_BODY, EN_LANGUAGE_PACK.checkInLabels, true);
  assert.equal(result.getTime(), Date.UTC(2026, 8, 11, 14, 0));
});

test('extractCheckInOutDate: confirmation body check-out, useStartTime=false -> latest departure instant', () => {
  const result = extractCheckInOutDate(CONFIRMATION_BODY, EN_LANGUAGE_PACK.checkOutLabels, false);
  assert.equal(result.getTime(), Date.UTC(2026, 8, 13, 10, 0));
});

test('extractCheckInOutDate: cancellation date-only check-in line, useStartTime=true -> midnight UTC fallback', () => {
  const result = extractCheckInOutDate(CANCELLATION_BODY, EN_LANGUAGE_PACK.checkInLabels, true);
  assert.equal(result.getTime(), Date.UTC(2026, 8, 11, 0, 0));
});

test('extractCheckInOutDate: label missing entirely returns null (not a throw)', () => {
  assert.equal(extractCheckInOutDate(CANCELLATION_BODY, ['Nonexistent Label'], true), null);
});

// --- findLabeledLineAcrossLanguagePacks / extractCheckInOutDateAcrossLanguagePacks
//
// The real GAS call sites (handleConfirmation/handleCancellation) need to
// find WHICH language pack's label matched a line, since the label match
// and the date FORMAT are coupled (if Czech's "Příjezd" matched, the line
// MUST be in Czech date format) — getBookingLabels' flattened cross-language
// union can't express this, since it loses which pack a value came from.
// BOOKING_LANGUAGE_PACKS is a shared mutable module-level object — these
// tests register the real EN_LANGUAGE_PACK/CS_LANGUAGE_PACK temporarily and
// clean up afterward so they never leak state into any other test.

test('findLabeledLineAcrossLanguagePacks: finds the Czech pack for a Czech email body', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;

  try {
    const found = findLabeledLineAcrossLanguagePacks(CZ_CONFIRMATION_BODY, 'checkInLabels');
    assert.notEqual(found, null);
    assert.equal(found.pack, CS_LANGUAGE_PACK);
    assert.ok(found.lineValue.indexOf('září') !== -1);
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
    delete BOOKING_LANGUAGE_PACKS.cs;
  }
});

test('findLabeledLineAcrossLanguagePacks: finds the English pack for an English email body', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;

  try {
    const found = findLabeledLineAcrossLanguagePacks(CONFIRMATION_BODY, 'checkInLabels');
    assert.notEqual(found, null);
    assert.equal(found.pack, EN_LANGUAGE_PACK);
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
    delete BOOKING_LANGUAGE_PACKS.cs;
  }
});

test('findLabeledLineAcrossLanguagePacks: no registered pack matches returns null', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;

  try {
    const found = findLabeledLineAcrossLanguagePacks('nothing relevant here', 'checkInLabels');
    assert.equal(found, null);
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
  }
});

test('extractCheckInOutDateAcrossLanguagePacks: Czech confirmation body check-in resolves to 2026-09-11 14:00 UTC (not shifted)', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;

  try {
    const result = extractCheckInOutDateAcrossLanguagePacks(CZ_CONFIRMATION_BODY, 'checkInLabels', true);
    assert.equal(result.getTime(), Date.UTC(2026, 8, 11, 14, 0));
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
    delete BOOKING_LANGUAGE_PACKS.cs;
  }
});

test('extractCheckInOutDateAcrossLanguagePacks: Czech confirmation body check-out resolves to 2026-09-13 10:00 UTC', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;

  try {
    const result = extractCheckInOutDateAcrossLanguagePacks(CZ_CONFIRMATION_BODY, 'checkOutLabels', false);
    assert.equal(result.getTime(), Date.UTC(2026, 8, 13, 10, 0));
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
    delete BOOKING_LANGUAGE_PACKS.cs;
  }
});

test('extractCheckInOutDateAcrossLanguagePacks: Czech cancellation body check-in (comma variant, no time) falls back to midnight UTC', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;

  try {
    const result = extractCheckInOutDateAcrossLanguagePacks(CZ_CANCELLATION_BODY, 'checkInLabels', true);
    assert.equal(result.getTime(), Date.UTC(2026, 8, 11, 0, 0));
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
    delete BOOKING_LANGUAGE_PACKS.cs;
  }
});

test('extractCheckInOutDateAcrossLanguagePacks: English confirmation body still resolves correctly (zero regression across the new cross-language pipeline)', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;

  try {
    const checkIn = extractCheckInOutDateAcrossLanguagePacks(CONFIRMATION_BODY, 'checkInLabels', true);
    const checkOut = extractCheckInOutDateAcrossLanguagePacks(CONFIRMATION_BODY, 'checkOutLabels', false);
    assert.equal(checkIn.getTime(), Date.UTC(2026, 8, 11, 14, 0));
    assert.equal(checkOut.getTime(), Date.UTC(2026, 8, 13, 10, 0));
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
    delete BOOKING_LANGUAGE_PACKS.cs;
  }
});

test('extractCheckInOutDateAcrossLanguagePacks: no registered pack matches returns null (not a throw)', () => {
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;

  try {
    const result = extractCheckInOutDateAcrossLanguagePacks('nothing relevant here', 'checkInLabels', true);
    assert.equal(result, null);
  } finally {
    delete BOOKING_LANGUAGE_PACKS.en;
  }
});

// --- extractHotelName --------------------------------------------------------
//
// Breaking signature change (2nd time): separators are now { separator,
// side } entries instead of bare strings, driven by real Czech data — the
// SAME separator (' – ') needs OPPOSITE sides depending on whether it's a
// confirmation or cancellation subject, so the old single-shared-list
// design (and the old bare-string entries) could not express this. English
// now carries confirmationHotelNameSeparators/cancellationHotelNameSeparators
// instead of one combined hotelNameSubjectSeparators list.

test('extractHotelName: English confirmation subject, " at " separator, hotel name AFTER', () => {
  assert.equal(extractHotelName(CONFIRMATION_SUBJECT, EN_LANGUAGE_PACK.confirmationHotelNameSeparators), 'Hotel U Modré hvězdy');
});

test('extractHotelName: English cancellation subject, " for " separator, hotel name AFTER', () => {
  assert.equal(extractHotelName(CANCELLATION_SUBJECT, EN_LANGUAGE_PACK.cancellationHotelNameSeparators), 'Hotel U Modré hvězdy');
});

test('extractHotelName: subject with none of the separators present returns null', () => {
  assert.equal(extractHotelName('No separator here', EN_LANGUAGE_PACK.confirmationHotelNameSeparators), null);
});

test('extractHotelName: leading emoji/pictographic noise before the hotel name is stripped (before-side real-data fix)', () => {
  const entries = [{ separator: ' – ', side: 'before' }];
  assert.equal(extractHotelName('🛄 Hotel Example – Something', entries), 'Hotel Example');
});

test('extractHotelName: Czech confirmation subject, " – " separator, hotel name BEFORE', () => {
  assert.equal(extractHotelName(CZ_CONFIRMATION_SUBJECT, CS_LANGUAGE_PACK.confirmationHotelNameSeparators), CZ_HOTEL_NAME);
});

test('extractHotelName: Czech cancellation subject, " – " separator, hotel name AFTER (same separator, opposite side from confirmation)', () => {
  assert.equal(extractHotelName(CZ_CANCELLATION_SUBJECT, CS_LANGUAGE_PACK.cancellationHotelNameSeparators), CZ_HOTEL_NAME);
});

// --- nextNonEmptyLineAfterLabel ----------------------------------------------

test('nextNonEmptyLineAfterLabel: confirmation body, locationLabels -> the address line', () => {
  assert.equal(
    nextNonEmptyLineAfterLabel(CONFIRMATION_BODY, EN_LANGUAGE_PACK.locationLabels),
    'Horni namesti 21, Olomouc, 77900, Czech Republic'
  );
});

test('nextNonEmptyLineAfterLabel: no matching label returns null', () => {
  assert.equal(nextNonEmptyLineAfterLabel(CONFIRMATION_BODY, ['Nonexistent Label']), null);
});

// --- computeSearchWindow -----------------------------------------------------
//
// Live-test-driven fix: Calendar.Events.list's `q` free-text search is
// unreliable for a bare numeric confirmation number, so findMatchingCalendarEvent
// switched to a bounded date-range enumeration instead. computeSearchWindow is
// the pure helper that derives that bounded window.

const ONE_DAY_MS = 24 * 60 * 60 * 1000;

test('computeSearchWindow: both dates present, default padding (7 days) -> checkIn-7d .. checkOut+7d', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11, 14, 0));
  const checkOut = new Date(Date.UTC(2026, 8, 13, 10, 0));

  const result = computeSearchWindow(checkIn, checkOut);

  assert.equal(result.timeMin, new Date(Date.UTC(2026, 8, 4, 14, 0)).toISOString());
  assert.equal(result.timeMax, new Date(Date.UTC(2026, 8, 20, 10, 0)).toISOString());
});

test('computeSearchWindow: both dates present, explicit paddingDays=1 -> checkIn-1d .. checkOut+1d', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11, 14, 0));
  const checkOut = new Date(Date.UTC(2026, 8, 13, 10, 0));

  const result = computeSearchWindow(checkIn, checkOut, 1);

  assert.equal(result.timeMin, new Date(Date.UTC(2026, 8, 10, 14, 0)).toISOString());
  assert.equal(result.timeMax, new Date(Date.UTC(2026, 8, 14, 10, 0)).toISOString());
});

test('computeSearchWindow: padding math is exactly paddingDays * 24h in milliseconds', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11, 14, 0));
  const checkOut = new Date(Date.UTC(2026, 8, 13, 10, 0));

  const result = computeSearchWindow(checkIn, checkOut, 3);

  assert.equal(new Date(result.timeMin).getTime(), checkIn.getTime() - 3 * ONE_DAY_MS);
  assert.equal(new Date(result.timeMax).getTime(), checkOut.getTime() + 3 * ONE_DAY_MS);
});

test('computeSearchWindow: result fields are ISO 8601 strings', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11, 14, 0));
  const checkOut = new Date(Date.UTC(2026, 8, 13, 10, 0));

  const result = computeSearchWindow(checkIn, checkOut);

  assert.equal(typeof result.timeMin, 'string');
  assert.equal(typeof result.timeMax, 'string');
  assert.match(result.timeMin, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
  assert.match(result.timeMax, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
});

test('computeSearchWindow: checkInDate null falls back to a wide (~2-year) window bracketing now', () => {
  const before = Date.now();
  const result = computeSearchWindow(null, new Date(Date.UTC(2026, 8, 13, 10, 0)));
  const after = Date.now();

  const timeMin = new Date(result.timeMin).getTime();
  const timeMax = new Date(result.timeMax).getTime();

  assert.ok(timeMin <= before, 'wide-window timeMin should be at or before "now"');
  assert.ok(timeMax >= after, 'wide-window timeMax should be at or after "now"');
  const spanDays = (timeMax - timeMin) / ONE_DAY_MS;
  assert.ok(spanDays >= 728 && spanDays <= 732, 'wide-window span should be approximately 2 years, got ' + spanDays + ' days');
});

test('computeSearchWindow: checkOutDate null falls back to the wide window', () => {
  const before = Date.now();
  const result = computeSearchWindow(new Date(Date.UTC(2026, 8, 11, 14, 0)), null);
  const after = Date.now();

  const timeMin = new Date(result.timeMin).getTime();
  const timeMax = new Date(result.timeMax).getTime();

  assert.ok(timeMin <= before);
  assert.ok(timeMax >= after);
});

test('computeSearchWindow: both dates null falls back to the wide window', () => {
  const before = Date.now();
  const result = computeSearchWindow(null, null);
  const after = Date.now();

  const timeMin = new Date(result.timeMin).getTime();
  const timeMax = new Date(result.timeMax).getTime();

  assert.ok(timeMin <= before);
  assert.ok(timeMax >= after);
});

test('computeSearchWindow: invalid Date objects (NaN time) are treated the same as null (wide fallback)', () => {
  const before = Date.now();
  const result = computeSearchWindow(new Date(NaN), new Date(NaN));
  const after = Date.now();

  const timeMin = new Date(result.timeMin).getTime();
  const timeMax = new Date(result.timeMax).getTime();

  assert.ok(timeMin <= before);
  assert.ok(timeMax >= after);
});

// --- eventDateRangeOverlaps --------------------------------------------------
//
// Second live-test-driven fix: fromGmail auto-created "stay" events never
// expose the real confirmation number via the public API (description is a
// generic Google placeholder), so matching falls back to hotel-name +
// date-range overlap against the real, readable summary/location/start/end
// fields. eventStart/eventEnd are raw Calendar API objects shaped
// {date: 'YYYY-MM-DD'} (all-day); eventEnd.date is EXCLUSIVE (the day after
// the last night).

test('eventDateRangeOverlaps: real numbers, default tolerance (1 day) -> true (obviously overlapping)', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11));
  const checkOut = new Date(Date.UTC(2026, 8, 13));

  assert.equal(
    eventDateRangeOverlaps(FROM_GMAIL_STAY_EVENT.start, FROM_GMAIL_STAY_EVENT.end, checkIn, checkOut),
    true
  );
});

test('eventDateRangeOverlaps: clearly non-overlapping date range -> false', () => {
  const checkIn = new Date(Date.UTC(2026, 9, 1)); // October 1 - far from the September event
  const checkOut = new Date(Date.UTC(2026, 9, 3));

  assert.equal(
    eventDateRangeOverlaps(FROM_GMAIL_STAY_EVENT.start, FROM_GMAIL_STAY_EVENT.end, checkIn, checkOut, 1),
    false
  );
});

test('eventDateRangeOverlaps: exact exclusive-end boundary (event end date === checkOutDate+1) -> true, not an off-by-one exclusion', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11));
  const checkOut = new Date(Date.UTC(2026, 8, 13));
  // Event end.date "2026-09-14" is exactly checkOutDate+1 (the exclusive-end
  // convention for an event whose last night is 2026-09-13).
  const eventStart = { date: '2026-09-11' };
  const eventEnd = { date: '2026-09-14' };

  assert.equal(eventDateRangeOverlaps(eventStart, eventEnd, checkIn, checkOut, 0), true);
});

test('eventDateRangeOverlaps: missing/invalid checkInDate or checkOutDate returns false (conservative, no false positives)', () => {
  const checkOut = new Date(Date.UTC(2026, 8, 13));
  assert.equal(eventDateRangeOverlaps(FROM_GMAIL_STAY_EVENT.start, FROM_GMAIL_STAY_EVENT.end, null, checkOut), false);
});

// --- eventSummaryOrLocationContainsHotelName ---------------------------------

test('eventSummaryOrLocationContainsHotelName: hotel name found in summary ("Stay at " + hotel name)', () => {
  assert.equal(eventSummaryOrLocationContainsHotelName(FROM_GMAIL_STAY_EVENT, HOTEL_NAME), true);
});

test('eventSummaryOrLocationContainsHotelName: hotel name found in location only', () => {
  const event = { summary: 'Some unrelated summary', location: 'Near ' + HOTEL_NAME };
  assert.equal(eventSummaryOrLocationContainsHotelName(event, HOTEL_NAME), true);
});

test('eventSummaryOrLocationContainsHotelName: case-insensitive match', () => {
  const event = { summary: 'STAY AT ' + HOTEL_NAME.toUpperCase(), location: '' };
  assert.equal(eventSummaryOrLocationContainsHotelName(event, HOTEL_NAME), true);
});

test('eventSummaryOrLocationContainsHotelName: no match in either field returns false', () => {
  const event = { summary: 'Unrelated meeting', location: 'Some other place' };
  assert.equal(eventSummaryOrLocationContainsHotelName(event, HOTEL_NAME), false);
});

test('eventSummaryOrLocationContainsHotelName: null/empty hotelName never matches (no accidental match-everything)', () => {
  assert.equal(eventSummaryOrLocationContainsHotelName(FROM_GMAIL_STAY_EVENT, null), false);
  assert.equal(eventSummaryOrLocationContainsHotelName(FROM_GMAIL_STAY_EVENT, ''), false);
});

// --- findFuzzyMatchingEvent ---------------------------------------------------

test('findFuzzyMatchingEvent: returns the first event matching both hotel name and date-range overlap', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11));
  const checkOut = new Date(Date.UTC(2026, 8, 13));
  const nonMatchingByName = {
    summary: 'Stay at Some Other Hotel',
    location: 'Elsewhere',
    start: { date: '2026-09-11' },
    end: { date: '2026-09-14' },
  };
  const nonMatchingByDate = {
    summary: 'Stay at ' + HOTEL_NAME,
    location: FROM_GMAIL_STAY_EVENT.location,
    start: { date: '2026-01-01' },
    end: { date: '2026-01-03' },
  };
  const events = [nonMatchingByName, nonMatchingByDate, FROM_GMAIL_STAY_EVENT];

  assert.equal(findFuzzyMatchingEvent(events, HOTEL_NAME, checkIn, checkOut), FROM_GMAIL_STAY_EVENT);
});

test('findFuzzyMatchingEvent: no matching event in the array returns null', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11));
  const checkOut = new Date(Date.UTC(2026, 8, 13));
  const nonMatchingByName = {
    summary: 'Stay at Some Other Hotel',
    location: 'Elsewhere',
    start: { date: '2026-09-11' },
    end: { date: '2026-09-14' },
  };

  assert.equal(findFuzzyMatchingEvent([nonMatchingByName], HOTEL_NAME, checkIn, checkOut), null);
});

test('findFuzzyMatchingEvent: empty events array returns null', () => {
  const checkIn = new Date(Date.UTC(2026, 8, 11));
  const checkOut = new Date(Date.UTC(2026, 8, 13));

  assert.equal(findFuzzyMatchingEvent([], HOTEL_NAME, checkIn, checkOut), null);
});

// --- formatLocalWallClockIso --------------------------------------------------
//
// Third live-test-driven fix: the owner's real created event showed
// check-in/check-out times 2 hours off (2pm reported as 4pm, 10am reported
// as noon) because handleConfirmation stamped extractCheckInOutDate's
// Date.UTC-built instant with `.toISOString()` (a literal 'Z'/UTC suffix)
// even though the underlying digits are wall-clock LOCAL time at the
// property, not UTC. formatLocalWallClockIso re-extracts the literal
// year/month/day/hour/minute digits via the UTC getters (which exactly
// recover the original numbers, since the Date was built via Date.UTC with
// no timezone shift ever applied) and formats them WITHOUT a trailing Z or
// offset, so the result can be paired with an explicit Calendar API
// `timeZone` field instead.

test('formatLocalWallClockIso: known instant (2026-09-11 14:00) formats as literal wall-clock digits, no Z suffix', () => {
  const date = new Date(Date.UTC(2026, 8, 11, 14, 0));
  assert.equal(formatLocalWallClockIso(date), '2026-09-11T14:00:00');
});

test('formatLocalWallClockIso: the real check-out instant (2026-09-13 10:00)', () => {
  const date = new Date(Date.UTC(2026, 8, 13, 10, 0));
  assert.equal(formatLocalWallClockIso(date), '2026-09-13T10:00:00');
});

test('formatLocalWallClockIso: zero-pads single-digit month/day/hour/minute', () => {
  const date = new Date(Date.UTC(2026, 0, 5, 9, 5));
  assert.equal(formatLocalWallClockIso(date), '2026-01-05T09:05:00');
});

test('formatLocalWallClockIso: midnight (no time window fallback case) formats with 00:00:00', () => {
  const date = new Date(Date.UTC(2026, 8, 11));
  assert.equal(formatLocalWallClockIso(date), '2026-09-11T00:00:00');
});

test('formatLocalWallClockIso: result contains no trailing Z or timezone offset', () => {
  const date = new Date(Date.UTC(2026, 8, 11, 14, 0));
  const result = formatLocalWallClockIso(date);
  assert.equal(/Z|[+-]\d{2}:\d{2}$/.test(result), false);
});

// --- extractReservationDetailsSection -----------------------------------------
//
// New feature: the created event's description should include the WHOLE
// "Reservation details" section of a confirmation email, not just the
// confirmation number + PIN. Config-driven start/end heading label lists
// (English-only for now, Czech equivalents to be appended once the owner
// reports real observed heading text from a Czech-language email).

test('extractReservationDetailsSection: real verbatim section text extracts correctly, bounded start/end headings excluded', () => {
  const result = extractReservationDetailsSection(
    CONFIRMATION_BODY_WITH_RESERVATION_DETAILS,
    EN_LANGUAGE_PACK.reservationDetailsHeadingLabels,
    EN_LANGUAGE_PACK.reservationDetailsEndHeadingLabels
  );

  assert.notEqual(result, null);
  assert.ok(
    result.indexOf('Check-in Friday, September 11, 2026 (2:00 PM - 8:00 PM)') !== -1,
    'expected check-in line to appear in the extracted section'
  );
  assert.ok(
    result.indexOf('Horni namesti 21, Olomouc, 77900, Czech Republic') !== -1,
    'expected the address line to appear in the extracted section'
  );
  assert.ok(result.indexOf('Reservation details') === -1, 'the start heading itself should not appear in the result');
  assert.ok(result.indexOf('Price details') === -1, 'the end heading should not appear in the result');
});

test('extractReservationDetailsSection: no start label present returns null', () => {
  const result = extractReservationDetailsSection(
    CONFIRMATION_BODY,
    EN_LANGUAGE_PACK.reservationDetailsHeadingLabels,
    EN_LANGUAGE_PACK.reservationDetailsEndHeadingLabels
  );

  assert.equal(result, null);
});

test('extractReservationDetailsSection: start label present but no end label captures to end of text (no crash)', () => {
  const text = ['Reservation details', '', 'Check-in Friday, September 11, 2026', 'Some other trailing line'].join('\n');

  const result = extractReservationDetailsSection(
    text,
    EN_LANGUAGE_PACK.reservationDetailsHeadingLabels,
    EN_LANGUAGE_PACK.reservationDetailsEndHeadingLabels
  );

  assert.notEqual(result, null);
  assert.ok(result.indexOf('Check-in Friday, September 11, 2026') !== -1);
  assert.ok(result.indexOf('Some other trailing line') !== -1);
});

test('extractReservationDetailsSection: heading is the last line with nothing meaningful after it returns null', () => {
  const text = ['Some preamble', 'Reservation details', '   ', ''].join('\n');

  const result = extractReservationDetailsSection(
    text,
    EN_LANGUAGE_PACK.reservationDetailsHeadingLabels,
    EN_LANGUAGE_PACK.reservationDetailsEndHeadingLabels
  );

  assert.equal(result, null);
});

test('extractReservationDetailsSection: Czech confirmation body captures check-in/out lines and room type, stops before "Informace o ceně"', () => {
  const result = extractReservationDetailsSection(
    CZ_CONFIRMATION_BODY,
    CS_LANGUAGE_PACK.reservationDetailsHeadingLabels,
    CS_LANGUAGE_PACK.reservationDetailsEndHeadingLabels
  );

  assert.notEqual(result, null);
  assert.ok(result.indexOf('Příjezd pátek 11. září 2026 (od 14:00)') !== -1, 'expected the Czech check-in line to appear');
  assert.ok(result.indexOf('Dvoulůžkový pokoj Deluxe') !== -1, 'expected the room type to appear');
  assert.ok(result.indexOf('Informace o rezervaci') === -1, 'the start heading itself should not appear in the result');
  assert.ok(result.indexOf('Informace o ceně') === -1, 'the end heading should not appear in the result');
});

// --- getBookingLabels / BOOKING_LANGUAGE_PACKS --------------------------------
//
// New language-pack architecture: language-dependent label arrays no
// longer live on BOOKING_ACTION_CONFIG. Instead, each language (e.g.
// src/06-lang-en.js) registers a pack into the shared
// BOOKING_LANGUAGE_PACKS registry, and getBookingLabels(fieldName) unions
// a given field's values across every registered pack at the point of
// use — so adding a new language ("copy 06-lang-en.js to
// 06-lang-<code>.js, translate the values") is enough, with zero changes
// to the action's own logic.

test('getBookingLabels: unions a field across multiple registered language packs', () => {
  // BOOKING_LANGUAGE_PACKS is a shared mutable module-level object — use
  // unique test-only keys and clean them up afterward so this test never
  // leaks state into any other test.
  BOOKING_LANGUAGE_PACKS.testLangA = { checkInLabels: ['Foo'] };
  BOOKING_LANGUAGE_PACKS.testLangB = { checkInLabels: ['Bar'] };

  try {
    const result = getBookingLabels('checkInLabels');
    assert.ok(result.indexOf('Foo') !== -1, 'expected "Foo" (testLangA) in the unioned result');
    assert.ok(result.indexOf('Bar') !== -1, 'expected "Bar" (testLangB) in the unioned result');
  } finally {
    delete BOOKING_LANGUAGE_PACKS.testLangA;
    delete BOOKING_LANGUAGE_PACKS.testLangB;
  }
});

test('getBookingLabels: a field absent from a registered pack contributes nothing (no throw)', () => {
  BOOKING_LANGUAGE_PACKS.testLangC = { someOtherField: ['X'] };

  try {
    const result = getBookingLabels('checkInLabels');
    assert.ok(Array.isArray(result));
  } finally {
    delete BOOKING_LANGUAGE_PACKS.testLangC;
  }
});

test('getBookingLabels: no registered packs returns an empty array (no throw)', () => {
  const savedKeys = Object.keys(BOOKING_LANGUAGE_PACKS);
  const saved = {};
  savedKeys.forEach(function (key) {
    saved[key] = BOOKING_LANGUAGE_PACKS[key];
    delete BOOKING_LANGUAGE_PACKS[key];
  });

  try {
    assert.deepEqual(getBookingLabels('checkInLabels'), []);
  } finally {
    savedKeys.forEach(function (key) {
      BOOKING_LANGUAGE_PACKS[key] = saved[key];
    });
  }
});

// --- Language pack completeness check -----------------------------------------
//
// Safety net for future contributed language packs: every pack MUST carry
// exactly this set of keys — this test defines that contract explicitly so
// a future pack missing a key, or misspelling one, is caught immediately
// rather than silently degrading matching for that language only. REVISED
// (Czech real-world contribution): hotelNameSubjectSeparators split into
// confirmationHotelNameSeparators/cancellationHotelNameSeparators (before/
// after-aware entries), and a new parseDateLine function-reference field
// was added — most fields are still non-empty arrays, but parseDateLine is
// a function, checked separately.

const REQUIRED_LANGUAGE_PACK_KEYS = [
  'confirmationNumberLabels',
  'pinLabels',
  'checkInLabels',
  'checkOutLabels',
  'locationLabels',
  'confirmationHotelNameSeparators',
  'cancellationHotelNameSeparators',
  'reservationDetailsHeadingLabels',
  'reservationDetailsEndHeadingLabels',
  'addToCalendarSubjectContains',
  'removeFromCalendarSubjectContains',
  'parseDateLine',
];

const HOTEL_NAME_SEPARATOR_KEYS = ['confirmationHotelNameSeparators', 'cancellationHotelNameSeparators'];

function assertLanguagePackShape(pack, packName) {
  const actualKeys = Object.keys(pack).sort();
  const expectedKeys = REQUIRED_LANGUAGE_PACK_KEYS.slice().sort();

  assert.deepEqual(actualKeys, expectedKeys, packName + ' must carry exactly the required keys, no more, no fewer');

  REQUIRED_LANGUAGE_PACK_KEYS.forEach(function (key) {
    if (key === 'parseDateLine') {
      assert.equal(typeof pack[key], 'function', packName + '.' + key + ' must be a function');
      return;
    }

    assert.ok(Array.isArray(pack[key]), packName + '.' + key + ' must be an array');
    assert.ok(pack[key].length > 0, packName + '.' + key + ' must be non-empty');

    if (HOTEL_NAME_SEPARATOR_KEYS.indexOf(key) !== -1) {
      pack[key].forEach(function (entry) {
        assert.ok(
          typeof entry.separator === 'string' && entry.separator.length > 0,
          packName + '.' + key + ' entries must have a non-empty separator string'
        );
        assert.ok(
          entry.side === 'before' || entry.side === 'after',
          packName + '.' + key + ' entries must have side "before" or "after"'
        );
      });
    }
  });
}

test('EN_LANGUAGE_PACK: has exactly the required keys, each correctly shaped', () => {
  assertLanguagePackShape(EN_LANGUAGE_PACK, 'EN_LANGUAGE_PACK');
});

test('CS_LANGUAGE_PACK: has exactly the required keys, each correctly shaped', () => {
  assertLanguagePackShape(CS_LANGUAGE_PACK, 'CS_LANGUAGE_PACK');
});

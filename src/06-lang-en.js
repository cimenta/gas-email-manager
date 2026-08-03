/**
 * EN_LANGUAGE_PACK — the English language pack for the booking.com
 * management action (src/06-action-booking-com-management.js).
 *
 * CONTRIBUTION CONTRACT (how to add a new language, e.g. Czech):
 *   1. Copy this file to src/06-lang-<code>.js (e.g. src/06-lang-cs.js).
 *   2. Translate every array's string VALUES to the real, observed text
 *      from a real booking.com email in that language — never guess or
 *      machine-translate a value that hasn't been confirmed against a
 *      real email; a wrong label silently fails to match rather than
 *      throwing, so an unverified guess is worse than no entry at all.
 *   3. Keep every KEY NAME identical to this file (do not rename, add,
 *      or remove keys) — src/06-action-booking-com-management.js reads
 *      these fields by name via getBookingLabels('fieldName') and has no
 *      other way to discover a pack's shape.
 *   4. Register the new pack under BOOKING_LANGUAGE_PACKS.<code> in the
 *      `else` branch at the bottom of this file (see below) — replace
 *      `<code>` with your language's short code (e.g. 'cs').
 *   5. No other file needs to change. getBookingLabels (in
 *      src/06-action-booking-com-management.js) automatically unions
 *      every registered pack's values for a given field at the point of
 *      use, so English + your new language both work simultaneously with
 *      zero code changes to the action's own logic.
 *
 * STRUCTURAL VS LEXICAL DIFFERENCES (learned from the first real-world
 * contribution, Czech — see src/06-lang-cs.js's own header for the full
 * story): most fields above are lexical-only (translate the words, keep
 * the mechanism). Two fields can differ STRUCTURALLY between languages,
 * not just lexically:
 *   - `confirmationHotelNameSeparators`/`cancellationHotelNameSeparators`:
 *     each entry is `{ separator, side: 'before' | 'after' }` — some
 *     languages put the hotel name BEFORE the separator in the subject
 *     line, not after (English always uses 'after'; do not assume every
 *     language does).
 *   - `parseDateLine`: a FUNCTION reference, not an array, parsing that
 *     language's check-in/check-out date-and-time line. English's is
 *     `parseBookingDateLine` (defined in the action file, referenced
 *     below — NOT duplicated here). Most future languages will NOT need
 *     a new parser — only write one when a language's date/time format
 *     genuinely differs from an already-registered one (e.g. different
 *     month-name grammar/position, a different time-window notation). A
 *     language whose emails already match an EXISTING format (English's
 *     or another registered pack's) can just reference THAT pack's
 *     `parseDateLine` from its own pack instead of writing a new one.
 *
 * This file registers itself into the shared BOOKING_LANGUAGE_PACKS
 * registry (declared in src/06-action-booking-com-management.js) via
 * GAS's shared-global-scope side effect — see that file's own
 * BOOKING_LANGUAGE_PACKS doc comment for the full load-order guarantee
 * this depends on (alphabetical-by-filename GAS execution order; do not
 * rename this file to something that would sort before "06-action-...js").
 */
const EN_LANGUAGE_PACK = {
  // Label phrasings observed across real booking.com confirmation and
  // cancellation emails — order matters: extractLabeledNumber returns the
  // first label (in this list's order) that matches anywhere in the text.
  confirmationNumberLabels: ['Confirmation Number', 'Confirmation', 'Booking number', 'Confirmation number'],
  pinLabels: ['PIN code', 'PIN Code', 'PIN'],
  checkInLabels: ['Check-in'],
  checkOutLabels: ['Check-out'],
  locationLabels: ['Location'],

  // Separators between the hotel name and the rest of a subject line —
  // English always puts the hotel name AFTER the separator (see
  // extractHotelName): the confirmation subject uses " at " (e.g.
  // "...confirmed at Hotel Name"), the cancellation subject uses " for "
  // (e.g. "Booking canceled for Hotel Name") — both real, confirmed
  // formats. Kept as two SEPARATE fields (not one shared list) because a
  // real-world language (Czech) needed the SAME separator string to take
  // OPPOSITE sides depending on confirmation vs. cancellation — a single
  // shared list cannot express that, so this file follows the same
  // two-field shape for consistency even though English's own two entries
  // happen to both be 'after'.
  confirmationHotelNameSeparators: [{ separator: ' at ', side: 'after' }],
  cancellationHotelNameSeparators: [{ separator: ' for ', side: 'after' }],

  // Bound the "Reservation details" section of a confirmation email's body
  // (see extractReservationDetailsSection) — everything between the
  // heading matching reservationDetailsHeadingLabels and the NEXT heading
  // matching reservationDetailsEndHeadingLabels is copied verbatim into
  // the created event's description (check-in/out, room type, guest
  // count, address, cancellation policy, etc.).
  reservationDetailsHeadingLabels: ['Reservation details'],
  reservationDetailsEndHeadingLabels: ['Price details'],

  // Subject substrings identifying a confirmation vs. a cancellation
  // email (matched case-insensitively via subjectContainsAny).
  addToCalendarSubjectContains: ['booking is confirmed'],
  removeFromCalendarSubjectContains: ['Booking canceled for'],

  // parseDateLine is assigned BELOW, in the environment-specific branch —
  // deliberately NOT as a plain object-literal property here. English's
  // date-line format is parsed by the action file's own
  // parseBookingDateLine (referenced, not duplicated, so it stays the
  // single source of truth for English's format) — but under Node, this
  // file is a standalone required module with no access to the action
  // file's top-level `parseBookingDateLine` identifier (GAS's implicit
  // shared-global-scope resolution, where every project file's top-level
  // declarations are mutually visible, has no Node equivalent for a
  // `require()`d module). See the branch below for how each environment
  // sources it correctly.
};

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Under Node this file is required directly by tests (e.g. to source a
// concrete label array for a pure-helper test) and does NOT mutate the
// real BOOKING_LANGUAGE_PACKS registry — that registration is GAS-only
// wiring (the `else` branch below), proven only by the live checkpoint,
// same category as the rest of the action's Calendar/Gmail-touching code.
if (typeof module !== 'undefined' && module.exports) {
  // Node: explicitly require the action file to source parseBookingDateLine
  // (see the "parseDateLine is assigned BELOW" comment above for why this
  // can't be a plain object-literal reference). Node caches modules by
  // resolved path, so this returns the SAME function reference the test
  // suite's own direct require of the action file already has — no
  // duplicate side effects, no circular-require issue (the action file
  // itself never requires any 06-lang-*.js file).
  EN_LANGUAGE_PACK.parseDateLine = require('./06-action-booking-com-management.js').parseBookingDateLine;
  module.exports = { EN_LANGUAGE_PACK: EN_LANGUAGE_PACK };
} else {
  EN_LANGUAGE_PACK.parseDateLine = parseBookingDateLine;
  BOOKING_LANGUAGE_PACKS.en = EN_LANGUAGE_PACK;
}

/**
 * CS_LANGUAGE_PACK — the Czech language pack for the booking.com
 * management action (src/06-action-booking-com-management.js). See
 * src/06-lang-en.js's own header for the full step-by-step contribution
 * contract every language pack follows.
 *
 * FIRST REAL-WORLD LANGUAGE-PACK CONTRIBUTION: building this pack from two
 * real Czech booking.com emails (confirmation + cancellation, "Hotel
 * Nikol - Free On-site Parking", reservation number 6203948175) revealed
 * that Czech's date/time format and hotel-name subject position are
 * STRUCTURALLY different from English, not just lexically different:
 *
 *   - HOTEL NAME POSITION: the SAME subject separator (' – ', an en dash)
 *     needs OPPOSITE sides depending on the email type — Czech
 *     confirmation: "{hotel} – Děkujeme!..." (hotel BEFORE the separator);
 *     Czech cancellation: "Zrušení rezervace – {hotel}" (hotel AFTER). This
 *     is why confirmationHotelNameSeparators/cancellationHotelNameSeparators
 *     exist as two separate, side-aware fields (see extractHotelName in
 *     the action file) instead of one shared separator list.
 *   - DATE/TIME LINE FORMAT: Czech's check-in/check-out lines are shaped
 *     entirely differently from English's, hence parseCzechBookingDateLine
 *     below (see its own JSDoc for the full breakdown) instead of
 *     referencing EN_LANGUAGE_PACK.parseDateLine.
 *
 * Most future languages will NOT need a new parseDateLine — only when a
 * language's date/time format genuinely differs from one already
 * registered, which is the case here.
 *
 * Every value in this pack is decoded verbatim from the two real emails
 * described above — never guessed or machine-translated (see
 * src/06-lang-en.js's contribution contract, point 2).
 */

/**
 * CZECH_MONTH_GENITIVE_TO_INDEX — lowercase Czech month name, GENITIVE
 * grammatical case (the form booking.com's Czech emails use in a date
 * like "11. září 2026"), -> zero-indexed month (0 = leden/January .. 11 =
 * prosinec/December), matching Date.UTC's month convention.
 *
 * VERIFICATION STATUS: only 'září' (September — which happens to be
 * grammatically invariant between nominative and genitive in Czech) is
 * EMPIRICALLY VERIFIED against a real booking.com email at the time this
 * pack was written. The remaining 11 genitive forms below are standard,
 * unambiguous Czech grammar (leden->ledna, únor->února, březen->března,
 * duben->dubna, květen->května, červen->června, červenec->července,
 * srpen->srpna, říjen->října, listopad->listopadu, prosinec->prosince)
 * and are included on that grammatical basis, but have NOT yet been
 * observed in a real email — flag/fix this table if a future real Czech
 * email surfaces an unexpected month-name mismatch.
 */
const CZECH_MONTH_GENITIVE_TO_INDEX = {
  ledna: 0,
  února: 1,
  března: 2,
  dubna: 3,
  května: 4,
  června: 5,
  července: 6,
  srpna: 7,
  září: 8,
  října: 9,
  listopadu: 10,
  prosince: 11,
};

/**
 * parseCzechTime — parses a 24-hour "H:MM" token into { hour, minute }. No
 * AM/PM conversion is needed — Czech booking.com emails use 24-hour time
 * directly (e.g. "14:00", never "2:00 PM"). Validates hour 0-23, minute
 * 0-59; throws a controlled Error otherwise, mirroring the action file's
 * parse12HourTime validation style for English. Pure, no GAS globals.
 */
function parseCzechTime(text) {
  const match = /^\s*(\d{1,2}):(\d{2})\s*$/.exec(text);
  if (!match) {
    throw new Error('Unrecognized Czech 24-hour time value: ' + text);
  }

  const hour = Number(match[1]);
  const minute = Number(match[2]);

  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (0-23) in Czech time value: ' + text);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in Czech time value: ' + text);
  }

  return { hour: hour, minute: minute };
}

/**
 * parseCzechBookingDateLine — the Czech-specific `parseDateLine`
 * implementation for this pack (see BOOKING_LANGUAGE_PACKS's language-pack
 * architecture in the action file: a language whose date/time format
 * genuinely differs from an already-registered one provides its own
 * parseDateLine function reference — Czech's does). Returns the SAME
 * shape as the action file's English-only parseBookingDateLine
 * ({ year, month, day, startTime, endTime }, month zero-indexed) so
 * calling code (extractCheckInOutDateAcrossLanguagePacks and friends)
 * never needs to know which language produced it.
 *
 * Czech's format differs STRUCTURALLY, not just lexically, from
 * English's:
 *   - The day number comes FIRST with a trailing period ("11."), then a
 *     GENITIVE Czech month name (see CZECH_MONTH_GENITIVE_TO_INDEX), then
 *     the year — with NO comma between day and year (unlike English's
 *     "September 11, 2026"). The date pattern is searched UNANCHORED
 *     anywhere in the line (the same technique the English parser uses to
 *     skip past a leading day-name word), so this works whether the line
 *     carries a leading day name with or without a trailing comma — both
 *     real, observed variants: "pátek 11. září 2026" (confirmation email)
 *     and "pátek, 11. září 2026" (cancellation email).
 *   - TIME WINDOW: Czech booking.com emails use AT LEAST TWO different
 *     time-window formats depending on the property (live-reported bug
 *     quick-260731-cs2, discovered via a SECOND real Czech confirmation
 *     email, "Apartmány Lipová", whose times came out as midnight
 *     under the original single-format implementation):
 *       - FORMAT 1 (Hotel Nikol, the first real email this pack was
 *         built from): EACH line carries its OWN single time, prefixed
 *         "od" (from/start) or "do" (until/end) — NEVER a range on one
 *         line. A check-in line has "od" only (-> startTime set, endTime
 *         null); a check-out line has "do" only (-> endTime set, startTime
 *         null).
 *       - FORMAT 2 ("Apartmány Lipová"): a SINGLE "HH:MM–HH:MM"
 *         time RANGE per line, with NO "od"/"do" keywords at all —
 *         structurally identical to how the English parser's own range
 *         handling works ("(2:00 PM - 8:00 PM)"), NOT like Format 1. Both
 *         startTime AND endTime are set from the SAME line (the caller's
 *         existing `useStartTime ? parsed.startTime : parsed.endTime`
 *         logic then picks the right one, exactly as it already does for
 *         English). The real separator is an EN DASH (–, U+2013); a plain
 *         ASCII hyphen (-) is also accepted, the same defensive posture
 *         already used elsewhere in this codebase for punctuation
 *         variants.
 *     Resolution order: the range pattern (Format 2) is tried FIRST; if it
 *     matches, both times are set from it and the od/do pattern is never
 *     consulted. If the range pattern does NOT match, the ORIGINAL od/do
 *     pattern (Format 1) is tried exactly as before — zero change to
 *     Hotel-Nikol-style parsing. If NEITHER pattern matches (a
 *     genuinely date-only line, e.g. a cancellation email's check-in/
 *     check-out lines with no time at all), both startTime/endTime stay
 *     null, exactly as before.
 *   - Either way, this exact returned shape means the EXISTING caller
 *     logic (`useStartTime ? parsed.startTime : parsed.endTime`) continues
 *     to work correctly unchanged — only the parser function differs per
 *     language, not the code that consumes its output.
 *   - Times are already 24-hour (parseCzechTime; no AM/PM conversion) in
 *     both formats.
 *
 * Throws a controlled Error if no day/month/year match is found anywhere
 * in the line, mirroring parseBookingDateLine's existing behavior for an
 * unparseable line. Pure, no GAS globals.
 */
function parseCzechBookingDateLine(lineText) {
  const text = String(lineText || '');
  const monthNamesPattern = Object.keys(CZECH_MONTH_GENITIVE_TO_INDEX).join('|');
  const dateMatch = new RegExp('(\\d{1,2})\\.\\s*(' + monthNamesPattern + ')\\s+(\\d{4})', 'i').exec(text);

  if (!dateMatch) {
    throw new Error('Unrecognized Czech booking date line: ' + lineText);
  }

  const day = Number(dateMatch[1]);
  const month = CZECH_MONTH_GENITIVE_TO_INDEX[dateMatch[2].toLowerCase()];
  const year = Number(dateMatch[3]);

  // FORMAT 2 first: a single "HH:MM[en dash or hyphen]HH:MM" range on the
  // line (see this function's class-level "TIME WINDOW" doc paragraph).
  const rangeMatch = /(\d{1,2}:\d{2})\s*[–-]\s*(\d{1,2}:\d{2})/.exec(text);

  let startTime = null;
  let endTime = null;

  if (rangeMatch) {
    startTime = parseCzechTime(rangeMatch[1]);
    endTime = parseCzechTime(rangeMatch[2]);
  } else {
    // FORMAT 1 fallback: the original od/do single-value pattern,
    // unchanged from before this fix.
    const startMatch = /\bod\s+(\d{1,2}:\d{2})\b/i.exec(text);
    const endMatch = /\bdo\s+(\d{1,2}:\d{2})\b/i.exec(text);

    startTime = startMatch ? parseCzechTime(startMatch[1]) : null;
    endTime = endMatch ? parseCzechTime(endMatch[1]) : null;
  }

  return {
    year: year,
    month: month,
    day: day,
    startTime: startTime,
    endTime: endTime,
  };
}

const CS_LANGUAGE_PACK = {
  // 'Potvrzení rezervace' (confirmation email's wording) and 'Číslo
  // rezervace' (cancellation email's DIFFERENT wording for the SAME
  // field) both belong in one flat list — extractLabeledNumber already
  // tries every label in list order and this field is read from both
  // email types, exactly like English's multiple confirmationNumberLabels
  // entries.
  confirmationNumberLabels: ['Potvrzení rezervace', 'Číslo rezervace'],
  // 'PIN kód' (confirmation) / 'Kód PIN' (cancellation) — same
  // word-order-differs-by-email-type reasoning as above.
  pinLabels: ['PIN kód', 'Kód PIN'],
  checkInLabels: ['Příjezd'],
  checkOutLabels: ['Odjezd'],
  locationLabels: ['Místo'],

  // See this file's class-level comment: the SAME separator (' – ') needs
  // OPPOSITE sides depending on confirmation vs. cancellation.
  confirmationHotelNameSeparators: [{ separator: ' – ', side: 'before' }],
  cancellationHotelNameSeparators: [{ separator: ' – ', side: 'after' }],

  reservationDetailsHeadingLabels: ['Informace o rezervaci'],
  reservationDetailsEndHeadingLabels: ['Informace o ceně'],

  addToCalendarSubjectContains: ['rezervace je potvrzena'],
  removeFromCalendarSubjectContains: ['Zrušení rezervace'],

  // parseCzechBookingDateLine is defined in THIS SAME file, so — unlike
  // EN_LANGUAGE_PACK.parseDateLine, which needs an explicit cross-file
  // require() under Node to reach a function defined in the action file —
  // it can be referenced directly as a plain object-literal property
  // here; no environment-specific branching is needed for a language
  // pack's OWN parser function, only for one borrowed from another file.
  parseDateLine: parseCzechBookingDateLine,
};

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Under Node this file is required directly by tests and does NOT mutate
// the real BOOKING_LANGUAGE_PACKS registry — that registration is GAS-only
// wiring (the `else` branch below), proven only by the live checkpoint.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CS_LANGUAGE_PACK: CS_LANGUAGE_PACK,
    parseCzechBookingDateLine: parseCzechBookingDateLine,
    CZECH_MONTH_GENITIVE_TO_INDEX: CZECH_MONTH_GENITIVE_TO_INDEX,
  };
} else {
  BOOKING_LANGUAGE_PACKS.cs = CS_LANGUAGE_PACK;
}

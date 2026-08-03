/**
 * BookingManagementAction — self-contained booking.com management action.
 *
 * On a booking.com CONFIRMATION email, adds a safety-net calendar event
 * ONLY when Google's own native detection has not already created one for
 * that confirmation number. On a booking.com CANCELLATION email, DELETES
 * the matching calendar event by confirmation number (silent no-op if none
 * matches).
 *
 * No GAS globals are referenced by the pure helpers in this file (see the
 * guarded Node export at the bottom) — they are exercised by the Node test
 * suite in test/. This action is fully self-contained: it re-implements its
 * own bookingExtractEmailAddress/bookingIsAllowedSender locally rather than
 * cross-requiring another action file (per the established one-file-per-action
 * pattern).
 *
 * Trust boundary: `text`/`subject` originate from UNTRUSTED inbound email
 * content. The remove path DELETES a real calendar event selected by a
 * confirmation number matched against that content — see T-lqi-01 in this
 * feature's threat register (accepted risk, mitigated by the blocking live
 * verification checkpoint; no dry-run mode).
 *
 * CONFIG SPLIT: BOOKING_ACTION_CONFIG (this action's tunable settings)
 * lives in the sibling src/06-action-cfg-booking-com-management.js, not in
 * this file — GAS's shared global scope means it's still visible by bare
 * name here, no require/import needed. This is DIFFERENT from
 * BOOKING_LANGUAGE_PACKS/getBookingLabels/the *_LANGUAGE_PACK files, which
 * remain in this file/their own 06-lang-*.js files — those are the
 * language-pack MECHANISM (how any action's language-dependent text gets
 * resolved), not this action's own tunable settings.
 * BOOKING_MANAGEMENT_ACTION's `config` property below is an ES6 GETTER
 * (not a plain literal reference), which is what makes the config split
 * safe regardless of which of the two files loads first alphabetically —
 * see BOOKING_ACTION_CONFIG's own doc comment in that sibling file for the
 * full load-order explanation (the same pattern already proven by
 * BOOKING_LANGUAGE_PACKS/getBookingLabels for the language-pack registry).
 *
 * GLOBALLY-UNIQUE NAMING WARNING (live-reported bug, real not cosmetic):
 * Apps Script concatenates every project file into ONE shared global
 * scope at runtime (unlike Node, where each file is its own module) — so
 * a top-level `function`/`const` name is NOT scoped to its own file, it
 * is scoped to the WHOLE PROJECT. Two files independently declaring a
 * same-named top-level function (e.g. this file's original
 * `extractEmailAddress`/`isAllowedSender` colliding with
 * src/05-action-ics-import.js's identically-named helpers) is a genuine
 * runtime collision — whichever file's declaration is evaluated last
 * (alphabetically) silently overwrites the other, and Apps Script's own
 * trigger-picker UI surfaces this as a "same name... undefined
 * behaviour" warning. This was fixed here by renaming this file's copies
 * to `bookingExtractEmailAddress`/`bookingIsAllowedSender` (see those
 * functions' own JSDoc for the full incident). ANY future action file OR
 * `06-lang-*.js` language-pack file MUST pick globally-unique top-level
 * names across the WHOLE project — "one file per action, local
 * reimplementation" does NOT imply per-file name scoping under GAS.
 *
 * MATCHING MECHANISM — full investigation history (findOrTagMatchingEvent):
 * the original design matched Google's own native fromGmail-detected event
 * by looking for the confirmation number inside its `description` field.
 * Three live-test rounds against the owner's real calendar found this
 * approach fundamentally could not work, for two independent reasons:
 *   1. Calendar.Events.list's `q` free-text search parameter turned out
 *      unreliable for a bare numeric token like a confirmation number
 *      (search indexes are known to handle pure-number tokens poorly) —
 *      fixed by switching to a bounded timeMin/timeMax date-range
 *      enumeration instead of `q` (see computeSearchWindow).
 *   2. Even with bounded enumeration, the description-substring match
 *      STILL failed: a fromGmail auto-created "stay" event's `description`
 *      is a generic, useless Google-authored placeholder — literally "To
 *      see detailed information... use the official Google Calendar app"
 *      — and NEVER contains the real confirmation number. Confirmed via
 *      raw Calendar.Events.get() JSON dumps. `description` is also NOT
 *      among the small set of fields Google's own docs list as patchable
 *      on a fromGmail event (colorId/reminders/visibility/transparency/
 *      status/attendees/extendedProperties are; description is not).
 *      A UI-driven "Duplicate" workaround was tested and ruled out too: a
 *      script-driven duplicate (Events.get + Events.insert with the same
 *      fields) only copies the same generic placeholder — the Calendar
 *      UI's own rich duplicate must reach data via privileged internal
 *      access the public API cannot reach.
 * RESOLUTION: findOrTagMatchingEvent now uses a two-layer lookup — an
 * EXACT match via a private extendedProperty tag (`confirmationNumber=...`,
 * settable via Calendar.Events.patch, which the docs DO allow), with a
 * FUZZY fallback (hotel-name substring in summary/location + date-range
 * overlap against the real, readable start/end fields — see
 * findFuzzyMatchingEvent) for the FIRST encounter of an untagged
 * (typically fromGmail) event. Once a fuzzy match is found, it is
 * immediately PATCHED with the tag, so every SUBSEQUENT lookup for that
 * same booking hits the exact-match path — the fuzzy fallback is only ever
 * needed once per booking. RESIDUAL RISK (documented, not eliminated): the
 * fuzzy fallback is necessarily lower-precision than an exact match for
 * that first untagged encounter — a coincidental false match would require
 * another booking with a similarly-named/matching hotel overlapping in the
 * same date window, which is a narrow, low-likelihood collision. This is a
 * necessary compromise given fromGmail events permanently do not expose
 * their confirmation number via the public API.
 *
 * MULTI-CALENDAR ROUTING (owner-requested): CONFIG.calendarId
 * (src/01-setup.js) is now a DEFAULT/fallback, not the sole target.
 * BOOKING_ACTION_CONFIG.calendarId (in the sibling config file) can
 * override it action-wide; resolveBookingCalendarId resolves the
 * two-tier choice. handleConfirmation/handleCancellation each resolve
 * this ONCE, as their very first statement, and thread it explicitly as a
 * parameter through every downstream Calendar API call site
 * (listCalendarEventsPaginated, findEventByConfirmationTag,
 * findOrTagMatchingEvent, CalendarApp.getTimeZone(), Calendar.Events.
 * insert/remove/patch) rather than each independently re-reading config —
 * see resolveBookingCalendarId's own JSDoc and each of those functions'
 * "MULTI-CALENDAR ROUTING" notes for the full threading. No per-sender
 * map is offered here (unlike the ICS action) — booking.com only ever
 * sends from one known address in practice.
 *
 * LANGUAGE PACKS: every natural-language-dependent piece of config (label
 * phrasings, subject substrings, etc.) lives OUTSIDE this file, in
 * per-language "06-lang-<code>.js" files (e.g. src/06-lang-en.js) that
 * register into the BOOKING_LANGUAGE_PACKS registry declared below —
 * BOOKING_ACTION_CONFIG itself now carries only non-language-dependent
 * settings. See BOOKING_LANGUAGE_PACKS's own doc comment for the full
 * contribution contract and the load-order guarantee this depends on.
 *
 * FILENAME WARNING: this file's name ("06-action-booking-com-
 * management.js") MUST keep sorting alphabetically BEFORE every
 * "06-lang-*.js" file (currently true: 'a' < 'l') — Apps Script executes
 * every project file's top-level code in alphabetical-by-filename order,
 * and BOOKING_LANGUAGE_PACKS must exist before any lang file's top-level
 * registration assignment runs, or that assignment throws a
 * ReferenceError. NEVER rename this file to something that would sort
 * after "06-lang-*.js" alphabetically.
 */

/**
 * BOOKING_LANGUAGE_PACKS — the shared registry every "06-lang-*.js" file
 * registers a language pack into via a GAS-shared-global-scope side
 * effect (the SAME mechanism this codebase already relies on for
 * 03-action-management.js referencing ICS_CALENDAR_ACTION/
 * BOOKING_MANAGEMENT_ACTION across files with zero explicit import —
 * Apps Script concatenates every file in a project into ONE global
 * execution scope at runtime; there is no import/require).
 *
 * CONTRIBUTION CONTRACT (adding a new language, e.g. Czech): copy
 * src/06-lang-en.js to src/06-lang-<code>.js, translate every array's
 * string VALUES to real, observed text from a real booking.com email in
 * that language (never guess/machine-translate an unverified value —
 * see 06-lang-en.js's own header comment for the full contract), keep
 * every key name identical, and register under
 * `BOOKING_LANGUAGE_PACKS.<code> = ...` in that new file. NO OTHER FILE
 * NEEDS TO CHANGE — getBookingLabels (below) automatically unions every
 * registered pack's values for a given field at the point of use.
 *
 * STRUCTURAL VS LEXICAL DIFFERENCES (learned from the first real-world
 * contribution, Czech — see src/06-lang-cs.js): most fields are plain
 * translated string arrays (lexical differences only — same matching
 * mechanism, different words). Two fields can differ STRUCTURALLY and
 * need real per-language logic, not just data:
 *   - `confirmationHotelNameSeparators`/`cancellationHotelNameSeparators`:
 *     each entry is `{ separator, side: 'before' | 'after' }` (see
 *     extractHotelName below) — some languages put the hotel name BEFORE
 *     the separator in the subject, not after (Czech's confirmation
 *     subject does; English's does not).
 *   - `parseDateLine`: a FUNCTION reference (not an array) parsing that
 *     language's check-in/check-out date-and-time line into
 *     { year, month, day, startTime, endTime } (month zero-indexed). Most
 *     future languages will NOT need a new one — only write one when a
 *     language's date/time format genuinely differs from an
 *     already-registered one (as Czech's does from English's: different
 *     month-name grammar/position, no comma before the year, one time
 *     per line via an "od"/"do"-style prefix instead of a single
 *     "-"-separated range) — a language whose emails already match an
 *     existing format can just reference that pack's `parseDateLine`
 *     directly (see src/06-lang-en.js's own header for the Node/GAS
 *     cross-require subtlety this involves for a REFERENCED function
 *     defined in ANOTHER file, vs. one defined in the same pack file).
 *
 * SCOPE NOTE: this language-pack architecture is deliberately
 * booking-action-only. The ICS action (src/05-action-ics-import.js)
 * parses a structured protocol format (RFC 5545), not natural-language
 * text, so it has no language-pack need — this project has an
 * established, deliberate precedent against centralizing config/
 * mechanism across actions until a second action actually needs it (see
 * this file's own one-file-per-action self-containment pattern, e.g.
 * bookingExtractEmailAddress/bookingIsAllowedSender re-implemented
 * locally rather than cross-required from the ICS action).
 *
 * LOAD-ORDER GUARANTEE — THE REAL MECHANISM (verified, not
 * `.clasp.json`'s `filePushOrder`): Google Apps Script's server-side
 * runtime executes each project file's top-level code in ALPHABETICAL
 * ORDER BY FILENAME, independent of clasp's `filePushOrder` setting.
 * `filePushOrder` only controls the order clasp UPLOADS files to the
 * remote project (historically relevant for HTML `include()` ordering in
 * web apps) — it does NOT change the Apps Script runtime's own execution
 * order once files are on the server. This project ALREADY relies on
 * this exact alphabetical-by-filename mechanism for its top-level
 * numbered files (01-setup.js .. 06-...js — see the 04-----------.js
 * empty separator file, and the project's own prior decision:
 * "getRegisteredActions() builds its array inline per call... to avoid
 * GAS file-load-order fragility"). "06-action-booking-com-management.js"
 * alphabetically precedes "06-lang-en.js" ('a' < 'l'), so this
 * declaration is guaranteed to exist before any lang file's registration
 * runs, with NO `.clasp.json` change needed — `filePushOrder` is
 * intentionally left as an empty array; changing it would be a red
 * herring fix that does not address the actual mechanism.
 */
const BOOKING_LANGUAGE_PACKS = {};

/**
 * getBookingLabels — the lazy accessor for a language-dependent config
 * field, returning the UNION of that field's values across EVERY
 * registered language pack (e.g. English + Czech confirmation-number
 * labels, concatenated, so extractLabeledNumber tries all of them
 * regardless of which language the inbound email happens to be in).
 *
 * Called AT THE POINT OF USE inside handleConfirmation/handleCancellation/
 * appliesTo/run (i.e. inside function bodies, which only execute long
 * after every file has finished loading) — deliberately NEVER cached
 * into a top-level `const` computed once at file-load time, because at
 * THAT point other files' registrations may not have run yet depending
 * on load order. Calling it lazily inside function bodies sidesteps the
 * load-order question entirely for correctness (though load order still
 * needs to be right for the BOOKING_LANGUAGE_PACKS declaration ITSELF to
 * exist before any lang file assigns to it — see the guarantee above).
 *
 * A field absent from a given pack contributes nothing (not a throw); no
 * registered packs at all returns []. Pure with respect to its own
 * logic (no GAS globals), but reads the GAS-shared-global-scope-populated
 * BOOKING_LANGUAGE_PACKS registry.
 */
function getBookingLabels(fieldName) {
  return Object.keys(BOOKING_LANGUAGE_PACKS).reduce(function (acc, lang) {
    const pack = BOOKING_LANGUAGE_PACKS[lang];
    const values = (pack && pack[fieldName]) || [];
    return acc.concat(values);
  }, []);
}

// Node/GAS environment bridge for BOOKING_ACTION_CONFIG (now defined in
// the sibling src/06-action-cfg-booking-com-management.js — see this
// file's class-level "CONFIG SPLIT" note). Under GAS's shared global
// scope, the sibling file's top-level `const BOOKING_ACTION_CONFIG` is
// ALREADY visible here by bare name — no action needed, and this
// `if` block never executes there (`typeof module` is always
// `'undefined'` under GAS). Under Node, each `require()`d file is its own
// isolated module with its own scope, so the several EXISTING internal
// bare references to `BOOKING_ACTION_CONFIG` elsewhere in this file
// (computeSearchWindow, eventDateRangeOverlaps,
// BOOKING_MANAGEMENT_ACTION.appliesTo/run) would otherwise throw
// ReferenceError. Assigning to `globalThis.BOOKING_ACTION_CONFIG` (rather
// than redeclaring a conflicting top-level `const`/`let`/`var` of the
// same name, which GAS's own concatenated scope would reject as a
// duplicate declaration) makes the bare identifier resolve correctly from
// any function in this file under Node too, via ordinary global-object
// fallback identifier resolution — with zero changes needed to any of
// those existing internal references.
if (typeof module !== 'undefined' && module.exports) {
  globalThis.BOOKING_ACTION_CONFIG = require('./06-action-cfg-booking-com-management.js').BOOKING_ACTION_CONFIG;
}

/**
 * bookingExtractEmailAddress — LOCAL copy of the ICS action's helper of the
 * same underlying logic (re-implemented here per the one-file-per-action
 * pattern, not cross-required). NAME-SPACED with a `booking` prefix
 * (renamed from a plain `extractEmailAddress` after a live-reported bug:
 * Apps Script concatenates every project file into ONE shared global
 * scope, so this file's original `extractEmailAddress` and the ICS
 * action's `extractEmailAddress` — both plain top-level function
 * declarations with the identical name — were a genuine RUNTIME
 * collision, not two independent "local copies" as the source-level
 * per-file layout implied. Whichever file's declaration evaluated last
 * (alphabetically, this file after 05-...) silently overwrote the other,
 * which is exactly what triggered the Apps Script trigger-picker's "same
 * name... undefined behaviour" warning. See this file's class-level
 * JSDoc for the general warning to future contributors). Extracts the
 * bare, trimmed, lowercased email address from a Gmail "From" header
 * value, or from a bare address with no display name. Pure, no GAS
 * globals. Never throws: a null/undefined/empty input returns ''.
 */
function bookingExtractEmailAddress(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const angleBracketMatch = /<([^>]*)>/.exec(fromHeader);
  const raw = angleBracketMatch ? angleBracketMatch[1] : fromHeader;

  return raw.trim().toLowerCase();
}

/**
 * bookingIsAllowedSender — LOCAL copy of the ICS action's helper of the
 * same underlying logic (renamed for the same real cross-file naming
 * collision reason as bookingExtractEmailAddress above — see that
 * function's JSDoc). Returns true when `allowList` is null, undefined, or
 * has zero length (allow-all); otherwise true only if the sender
 * extracted from `fromHeader` strictly equals the sender extracted from
 * at least one entry in `allowList` (each entry also run through
 * bookingExtractEmailAddress, so a bare address or a full 'Name <email>'
 * entry both match). Matching is case-insensitive because
 * bookingExtractEmailAddress lowercases both sides. Pure, no GAS globals.
 */
function bookingIsAllowedSender(fromHeader, allowList) {
  if (!allowList || allowList.length === 0) {
    return true;
  }

  const sender = bookingExtractEmailAddress(fromHeader);

  return allowList.some(function (entry) {
    return bookingExtractEmailAddress(entry) === sender;
  });
}

/**
 * resolveBookingCalendarId — resolves which calendar ID this action's
 * Calendar API calls should target. MULTI-CALENDAR ROUTING (owner-requested):
 * CONFIG.calendarId (src/01-setup.js) is a DEFAULT/fallback; this action
 * can override it action-wide via BOOKING_ACTION_CONFIG.calendarId. Simple
 * two-tier resolution: bookingConfig.calendarId when truthy, else
 * defaultCalendarId. No per-sender map is offered here (unlike the ICS
 * action's resolveIcsCalendarId) — booking.com only ever sends from one
 * known address in practice, so it was not requested and is not needed.
 * Pure, no GAS globals.
 */
function resolveBookingCalendarId(bookingConfig, defaultCalendarId) {
  return (bookingConfig && bookingConfig.calendarId) || defaultCalendarId;
}

/**
 * subjectContainsAny — returns a literal boolean: true when `subject`
 * case-insensitively contains any of the strings in `substrings`, false
 * when `substrings` is null/undefined/empty or none matches. Pure, no GAS
 * globals.
 */
function subjectContainsAny(subject, substrings) {
  if (!substrings || substrings.length === 0) {
    return false;
  }

  const lowerSubject = String(subject || '').toLowerCase();

  return substrings.some(function (substring) {
    return lowerSubject.indexOf(String(substring).toLowerCase()) !== -1;
  });
}

/**
 * escapeRegExp — escapes regex-special characters in a plain string so it
 * can be safely embedded inside a dynamically-built RegExp. Pure, no GAS
 * globals.
 */
function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * extractLabeledNumber — iterates `labels` in list order; for each, builds
 * a case-insensitive RegExp of the regex-escaped label followed by
 * optional whitespace, an IMMEDIATE colon, optional whitespace, then a
 * captured run of digits. Returns the first captured digit string for the
 * first label that matches anywhere in `text`, or null if none match.
 *
 * The immediate-colon requirement (label + '\s*:\s*(\d+)') is what stops a
 * shorter label like "Confirmation" from falsely matching inside a longer
 * label's occurrence like "Confirmation Number:" — "Confirmation" followed
 * by " Number:" is NOT "\s*:", so it does not match there. Pure, no GAS
 * globals.
 */
function extractLabeledNumber(text, labels) {
  if (!text || !labels) {
    return null;
  }

  for (let i = 0; i < labels.length; i++) {
    const pattern = new RegExp(escapeRegExp(labels[i]) + '\\s*:\\s*(\\d+)', 'i');
    const match = pattern.exec(text);
    if (match) {
      return match[1];
    }
  }

  return null;
}

/**
 * MONTH_NAME_TO_INDEX — lowercase full English month name -> zero-indexed
 * month (0 = January .. 11 = December), matching Date.UTC's month
 * convention. Structured as a plain lookup object (with this documenting
 * comment) so supporting another language is just adding more key/value
 * pairs here, with no change to any parsing logic that consumes it.
 */
const MONTH_NAME_TO_INDEX = {
  january: 0,
  february: 1,
  march: 2,
  april: 3,
  may: 4,
  june: 5,
  july: 6,
  august: 7,
  september: 8,
  october: 9,
  november: 10,
  december: 11,
};

/**
 * parse12HourTime — parses an "H:MM AM"/"H:MM PM" token into 24-hour
 * {hour, minute} using standard 12-hour rules (12 AM -> hour 0 (midnight),
 * 12 PM -> hour 12 (noon)). Validates the hour is 1-12 and the minute is
 * 0-59 before converting; throws a controlled Error on any unparseable or
 * out-of-range input. Pure, no GAS globals.
 */
function parse12HourTime(text) {
  const match = /^\s*(\d{1,2}):(\d{2})\s*(AM|PM)\s*$/i.exec(text);
  if (!match) {
    throw new Error('Unrecognized 12-hour time value: ' + text);
  }

  const hour12 = Number(match[1]);
  const minute = Number(match[2]);
  const meridiem = match[3].toUpperCase();

  if (hour12 < 1 || hour12 > 12) {
    throw new Error('Hour out of range (1-12) in 12-hour time value: ' + text);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in 12-hour time value: ' + text);
  }

  let hour24 = hour12 % 12; // 12 -> 0 (as a base), 1-11 -> itself
  if (meridiem === 'PM') {
    hour24 += 12;
  }

  return { hour: hour24, minute: minute };
}

/**
 * parseBookingDateLine — parses a booking.com date-line value into
 * { year, month, day, startTime, endTime }. `month` is zero-indexed
 * (Date.UTC convention). Matches the first "MonthName DD, YYYY" occurrence
 * (case-insensitive month name via MONTH_NAME_TO_INDEX), which naturally
 * skips any leading day-name word (e.g. "Friday, "). Throws a controlled
 * Error if that date portion is absent/unparseable. A parenthesized
 * "(H:MM AM/PM - H:MM AM/PM)" time window is OPTIONAL: when present,
 * startTime/endTime are set via parse12HourTime; when absent, both are
 * null (not an error — a missing time window is expected for date-only
 * lines). Pure, no GAS globals.
 */
function parseBookingDateLine(lineText) {
  const text = String(lineText || '');
  const monthNamesPattern = Object.keys(MONTH_NAME_TO_INDEX).join('|');
  const dateMatch = new RegExp('(' + monthNamesPattern + ')\\s+(\\d{1,2}),\\s*(\\d{4})', 'i').exec(text);

  if (!dateMatch) {
    throw new Error('Unrecognized booking date line: ' + lineText);
  }

  const month = MONTH_NAME_TO_INDEX[dateMatch[1].toLowerCase()];
  const day = Number(dateMatch[2]);
  const year = Number(dateMatch[3]);

  const windowMatch = /\(\s*(\d{1,2}:\d{2}\s*(?:AM|PM))\s*-\s*(\d{1,2}:\d{2}\s*(?:AM|PM))\s*\)/i.exec(text);

  return {
    year: year,
    month: month,
    day: day,
    startTime: windowMatch ? parse12HourTime(windowMatch[1]) : null,
    endTime: windowMatch ? parse12HourTime(windowMatch[2]) : null,
  };
}

/**
 * findLabeledLine — splits `text` into lines; for each line (left-trimmed),
 * tests each label in `labels` via a case-insensitive anchored pattern of
 * the regex-escaped label followed by one-or-more whitespace characters
 * then a captured remainder. Returns the trimmed remainder of the first
 * line that matches any label, or null if no line matches any label. Pure,
 * no GAS globals.
 */
function findLabeledLine(text, labels) {
  if (!text || !labels) {
    return null;
  }

  const lines = String(text).split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/^\s+/, '');

    for (let j = 0; j < labels.length; j++) {
      const pattern = new RegExp('^' + escapeRegExp(labels[j]) + '\\s+(.+)$', 'i');
      const match = pattern.exec(line);
      if (match) {
        return match[1].trim();
      }
    }
  }

  return null;
}

/**
 * buildInstantFromParsedDateLine — shared final step for both
 * extractCheckInOutDate and extractCheckInOutDateAcrossLanguagePacks:
 * given a parsed date-line object ({ year, month, day, startTime,
 * endTime }, month zero-indexed — the shape every parseDateLine
 * implementation, English or Czech or otherwise, returns) and the chosen
 * time ({hour, minute} or null), builds the resulting UTC Date. When the
 * chosen time is present: new Date(Date.UTC(year, month, day, hour,
 * minute)); when absent (a date-only line — no time window/prefix
 * present): falls back to midnight UTC: new Date(Date.UTC(year, month,
 * day)). Pure, no GAS globals.
 */
function buildInstantFromParsedDateLine(parsed, chosenTime) {
  if (chosenTime) {
    return new Date(Date.UTC(parsed.year, parsed.month, parsed.day, chosenTime.hour, chosenTime.minute));
  }

  return new Date(Date.UTC(parsed.year, parsed.month, parsed.day));
}

/**
 * extractCheckInOutDate — resolves a check-in or check-out instant from
 * `text` for the given `labels` (checkInLabels or checkOutLabels). Finds
 * the labeled line via findLabeledLine; if none is found, returns null (a
 * missing label is not an error). Otherwise parses the line via
 * `parseLineFn` (defaults to parseBookingDateLine — English's own parser —
 * when omitted, so every EXISTING call/test of this function is
 * byte-identical in behavior) and picks startTime when `useStartTime` is
 * true, else endTime, then builds the instant via
 * buildInstantFromParsedDateLine. A malformed date line's throw from
 * `parseLineFn` is allowed to propagate.
 *
 * `parseLineFn` exists so this function stays the shared LOWER-LEVEL pure
 * helper usable by any language's parser — see
 * extractCheckInOutDateAcrossLanguagePacks below for the cross-language
 * wrapper the real GAS handlers actually call, which determines the
 * correct pack (and therefore the correct parseLineFn) per email rather
 * than assuming English. Pure, no GAS globals.
 */
function extractCheckInOutDate(text, labels, useStartTime, parseLineFn) {
  const parseLine = parseLineFn || parseBookingDateLine;

  const lineValue = findLabeledLine(text, labels);
  if (lineValue === null) {
    return null;
  }

  const parsed = parseLine(lineValue);
  const chosenTime = useStartTime ? parsed.startTime : parsed.endTime;

  return buildInstantFromParsedDateLine(parsed, chosenTime);
}

/**
 * findLabeledLineAcrossLanguagePacks — finds which REGISTERED language
 * pack's `fieldName` labels (e.g. 'checkInLabels') match a line in `text`,
 * trying each pack in `Object.keys(BOOKING_LANGUAGE_PACKS)` order and
 * returning `{ lineValue, pack }` for the FIRST pack whose labels find a
 * match via the existing pure findLabeledLine, or null if no registered
 * pack matches at all.
 *
 * This exists because the label match and the date-line FORMAT are
 * coupled: if Czech's "Příjezd" label matched a line, that line MUST be
 * in Czech date format, and only the Czech pack's own `parseDateLine` can
 * parse it correctly — getBookingLabels' flattened cross-language union
 * (used elsewhere for label-only fields) would lose which pack a matched
 * label came from, making it useless here. In practice, a real email body
 * is written in exactly one language, so only one pack's labels will ever
 * actually match a given body — iteration order among packs does not
 * affect correctness, only which pack is found (correctly) first.
 *
 * Pure with respect to its own logic (no GAS globals), but reads the
 * GAS-shared-global-scope-populated BOOKING_LANGUAGE_PACKS registry (same
 * caveat as getBookingLabels).
 */
function findLabeledLineAcrossLanguagePacks(text, fieldName) {
  const langs = Object.keys(BOOKING_LANGUAGE_PACKS);

  for (let i = 0; i < langs.length; i++) {
    const pack = BOOKING_LANGUAGE_PACKS[langs[i]];
    if (!pack) {
      continue;
    }

    const labels = pack[fieldName] || [];
    const lineValue = findLabeledLine(text, labels);
    if (lineValue !== null) {
      return { lineValue: lineValue, pack: pack };
    }
  }

  return null;
}

/**
 * extractCheckInOutDateAcrossLanguagePacks — the CROSS-LANGUAGE wrapper
 * the real GAS handlers (handleConfirmation/handleCancellation) call
 * instead of the lower-level extractCheckInOutDate directly. Finds which
 * registered pack's `fieldName` labels (checkInLabels/checkOutLabels)
 * match a line in `text` via findLabeledLineAcrossLanguagePacks; if none
 * match, returns null (a missing label is not an error, same as
 * extractCheckInOutDate). Otherwise parses the matched line via THAT
 * SAME pack's own `parseDateLine` (falling back to English's
 * parseBookingDateLine only if a pack is somehow registered without one —
 * defensive, should not happen for a pack passing the completeness
 * check), picks startTime/endTime per `useStartTime`, and builds the
 * instant via the same buildInstantFromParsedDateLine helper
 * extractCheckInOutDate uses — so the "date-only line falls back to
 * midnight UTC" behavior is identical regardless of which language
 * produced the match. Pure with respect to its own logic (no GAS
 * globals), but reads BOOKING_LANGUAGE_PACKS (see
 * findLabeledLineAcrossLanguagePacks).
 */
function extractCheckInOutDateAcrossLanguagePacks(text, fieldName, useStartTime) {
  const found = findLabeledLineAcrossLanguagePacks(text, fieldName);
  if (!found) {
    return null;
  }

  const parseLine = found.pack.parseDateLine || parseBookingDateLine;
  const parsed = parseLine(found.lineValue);
  const chosenTime = useStartTime ? parsed.startTime : parsed.endTime;

  return buildInstantFromParsedDateLine(parsed, chosenTime);
}

/**
 * extractHotelName — tries each entry in `separatorEntries`, IN ORDER; each
 * entry is `{ separator: string, side: 'before' | 'after' }`. For the
 * first entry whose `separator` is actually present in `subject` (found
 * via the LAST occurrence), returns the trimmed text either BEFORE the
 * separator (`side: 'before'`) or AFTER it (`side: 'after'`); returns null
 * if none of the entries' separators are present.
 *
 * BREAKING SIGNATURE CHANGE (2nd time, real-world-data-driven): entries
 * used to be bare separator strings, always taking the "after" side. Real
 * Czech data proved this insufficient — the SAME separator (' – ', an en
 * dash) needs OPPOSITE sides depending on whether it's the confirmation
 * subject ("{hotel} – Děkujeme!...", hotel BEFORE) or the cancellation
 * subject ("Zrušení rezervace – {hotel}", hotel AFTER); a flat list of
 * bare separator strings cannot express this, since you cannot tell from
 * the separator string alone which side to take.
 *
 * After slicing, strips any leading run of non-letter/non-digit
 * characters (Unicode-aware) before the final trim — real confirmation
 * subjects start with a pictographic emoji (e.g. "🛄 "), which lands
 * inside the "before" slice for a before-side entry (English's
 * after-side entries never had this problem, since the emoji is always
 * in the discarded "before" portion for them). Applied uniformly
 * regardless of side so the behavior is consistent and future-proof; a
 * no-op for any hotel name that already starts with a letter/digit (i.e.
 * every existing English case). Pure, no GAS globals.
 */
function extractHotelName(subject, separatorEntries) {
  const text = String(subject || '');
  const list = separatorEntries || [];

  for (let i = 0; i < list.length; i++) {
    const entry = list[i];
    const index = text.lastIndexOf(entry.separator);

    if (index !== -1) {
      const raw = entry.side === 'before' ? text.slice(0, index) : text.slice(index + entry.separator.length);
      return raw.trim().replace(/^[^\p{L}\p{N}]+/u, '');
    }
  }

  return null;
}

/**
 * nextNonEmptyLineAfterLabel — finds the first line in `text` whose
 * left-trimmed text starts with any of `labels` (case-insensitive), then
 * returns the next SUBSEQUENT non-empty (trimmed) line, or null if no
 * label line is found (or no non-empty line follows it). Pure, no GAS
 * globals.
 *
 * UNICODE FIX (real Czech-data-driven bug): the "starts with this label"
 * check uses a Unicode-aware negative lookahead
 * (`(?![\p{L}\p{N}])`, with the `u` flag) rather than a plain `\b` word
 * boundary — JS's `\b` only recognizes ASCII word characters
 * (`[A-Za-z0-9_]`), so a label ending in a non-ASCII letter (e.g. Czech's
 * diacritic vowels) sits at a position `\b` does NOT treat as a boundary,
 * silently failing to match even though the label text is exactly
 * present. The lookahead form generalizes correctly to any language.
 */
function nextNonEmptyLineAfterLabel(text, labels) {
  if (!text || !labels) {
    return null;
  }

  const lines = String(text).split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/^\s+/, '');
    const matchesLabel = labels.some(function (label) {
      return new RegExp('^' + escapeRegExp(label) + '(?![\\p{L}\\p{N}])', 'iu').test(line);
    });

    if (matchesLabel) {
      for (let j = i + 1; j < lines.length; j++) {
        const trimmed = lines[j].trim();
        if (trimmed !== '') {
          return trimmed;
        }
      }
      return null;
    }
  }

  return null;
}

/**
 * extractReservationDetailsSection — captures the WHOLE "Reservation
 * details" section of a confirmation email body VERBATIM (check-in/out,
 * room type, guest count, address, cancellation policy, etc.), for
 * inclusion in the created event's description. LOCAL re-implementation
 * of blank-line collapsing (not a cross-require of the ICS action's
 * collapseBlankLines), per this file's established one-file-per-action
 * pattern.
 *
 * Splits `text` into lines; finds the FIRST line (left-trimmed) that
 * case-insensitively STARTS WITH any label in `startLabels` (same
 * anchored-regex matching style as findLabeledLine/
 * nextNonEmptyLineAfterLabel above). If no start label is found anywhere,
 * returns null. From the line AFTER that heading, collects every
 * subsequent line (each trimmed) until — EXCLUSIVE — the first line whose
 * TRIMMED text starts with any label in `endLabels` (same matching
 * style); a real "Price details" end heading is heavily left-padded in
 * the raw email, which is exactly why the END check trims first. If no
 * end label is found in the remaining text, collects through the end of
 * `text` — a truncated or differently-formatted email must never crash
 * the action, it should just capture everything remaining.
 *
 * Runs of 2+ consecutive blank (post-trim) lines collapse to a single
 * blank line; leading/trailing blank lines are trimmed from the final
 * result. Returns null (not an empty string) if nothing meaningful was
 * captured — e.g. the heading was the very last line with nothing after
 * it. Pure, no GAS globals.
 *
 * UNICODE FIX (real Czech-data-driven bug — see nextNonEmptyLineAfterLabel's
 * own JSDoc for the full explanation): both the start- and end-heading
 * checks use a Unicode-aware negative lookahead instead of a plain `\b`
 * word boundary, since `\b` silently fails to match a label ending in a
 * non-ASCII letter (e.g. Czech's "Informace o ceně", ending in "ě") even
 * when the label text is exactly present.
 */
function extractReservationDetailsSection(text, startLabels, endLabels) {
  if (!text || !startLabels) {
    return null;
  }

  const lines = String(text).split('\n');
  let startIndex = -1;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/^\s+/, '');
    const matchesStart = startLabels.some(function (label) {
      return new RegExp('^' + escapeRegExp(label) + '(?![\\p{L}\\p{N}])', 'iu').test(line);
    });
    if (matchesStart) {
      startIndex = i;
      break;
    }
  }

  if (startIndex === -1) {
    return null;
  }

  const collected = [];
  for (let j = startIndex + 1; j < lines.length; j++) {
    const trimmed = lines[j].trim();
    const matchesEnd = (endLabels || []).some(function (label) {
      return new RegExp('^' + escapeRegExp(label) + '(?![\\p{L}\\p{N}])', 'iu').test(trimmed);
    });
    if (matchesEnd) {
      break;
    }
    collected.push(trimmed);
  }

  // Collapse runs of 2+ consecutive blank lines down to a single blank line.
  const deduped = [];
  let prevBlank = false;
  collected.forEach(function (line) {
    const isBlank = line === '';
    if (isBlank && prevBlank) {
      return;
    }
    deduped.push(line);
    prevBlank = isBlank;
  });

  // Trim leading/trailing blank lines from the final result.
  while (deduped.length > 0 && deduped[0] === '') {
    deduped.shift();
  }
  while (deduped.length > 0 && deduped[deduped.length - 1] === '') {
    deduped.pop();
  }

  const joined = deduped.join('\n');
  return joined === '' ? null : joined;
}

/**
 * zeroPad — left-pads `value` with '0' to `length` digits. Pure, no GAS
 * globals. Tiny internal formatting helper for formatLocalWallClockIso.
 */
function zeroPad(value, length) {
  return String(value).padStart(length, '0');
}

/**
 * formatLocalWallClockIso — takes a Date built via
 * Date.UTC(year, month, day, hour, minute) (as extractCheckInOutDate
 * produces) and re-extracts the literal year/month/day/hour/minute digits
 * via the UTC getters (getUTCFullYear/getUTCMonth/getUTCDate/
 * getUTCHours/getUTCMinutes) — these exactly recover the original literal
 * wall-clock numbers, since the Date was built via Date.UTC in the first
 * place with no timezone shift ever applied anywhere. Formats them as a
 * zero-padded literal string 'YYYY-MM-DDTHH:MM:SS' — DELIBERATELY with NO
 * trailing 'Z' and NO timezone offset.
 *
 * This is the fix for a live-reported bug: `date.toISOString()` stamps a
 * literal UTC ('Z') suffix onto an instant whose digits are actually
 * wall-clock LOCAL time at the property (booking.com emails carry no
 * timezone indicator at all — see handleConfirmation's own JSDoc for how
 * the assumed timezone is now derived), which silently shifted every
 * created event by the difference between UTC and the owner's actual
 * calendar timezone (2 hours, for Central European Summer Time). The
 * string this function returns is meant to be paired with an explicit
 * Calendar API `timeZone` field (see handleConfirmation) so the API
 * interprets these digits as wall-clock local time in that zone, not UTC.
 * Pure, no GAS globals — this function itself is timezone-source-agnostic:
 * it only formats digits, it never decides WHICH timezone string is paired
 * with them (see handleConfirmation for that).
 */
function formatLocalWallClockIso(date) {
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  const hour = date.getUTCHours();
  const minute = date.getUTCMinutes();

  return (
    zeroPad(year, 4) +
    '-' +
    zeroPad(month, 2) +
    '-' +
    zeroPad(day, 2) +
    'T' +
    zeroPad(hour, 2) +
    ':' +
    zeroPad(minute, 2) +
    ':00'
  );
}

// ONE_DAY_MS is a pure unit-conversion constant (structural, not tunable
// matching policy), so — unlike the config-editable defaults below — it
// stays a plain module-level constant, alongside MONTH_NAME_TO_INDEX
// (a fixed language table, not a config value either).
const ONE_DAY_MS = 24 * 60 * 60 * 1000;

/**
 * isValidDate — true only for a real Date instance whose time value is not
 * NaN (i.e. not an "Invalid Date"). Pure, no GAS globals.
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

/**
 * computeSearchWindow — derives the bounded { timeMin, timeMax } (ISO 8601
 * strings) date-range window findMatchingCalendarEvent enumerates, replacing
 * a prior reliance on Calendar.Events.list's `q` free-text search parameter
 * (found live to be unreliable for a bare numeric confirmation number —
 * search indexes are known to handle pure-number tokens poorly, independent
 * of the matching event's eventType).
 *
 * When both `checkInDate` and `checkOutDate` are valid Dates: timeMin is
 * checkInDate minus `paddingDays` days, timeMax is checkOutDate plus
 * `paddingDays` days (default BOOKING_ACTION_CONFIG.searchWindowPaddingDays,
 * currently 7 — a tunable, owner-editable matching-behavior policy value,
 * not a structural constant). When either is missing/invalid (a real
 * possibility — a date line may not be present or parseable in a given
 * email; this is not always the "no dates at all" case), falls back to a
 * WIDE window (now minus/plus BOOKING_ACTION_CONFIG.wideFallbackWindowYears,
 * currently 1) so the search still has a chance to succeed rather than
 * silently failing to search at all. `paddingDays` remains an optional
 * explicit override — the config value is only the default. Pure, no GAS
 * globals.
 */
function computeSearchWindow(checkInDate, checkOutDate, paddingDays) {
  const padding = paddingDays === undefined ? BOOKING_ACTION_CONFIG.searchWindowPaddingDays : paddingDays;

  if (isValidDate(checkInDate) && isValidDate(checkOutDate)) {
    const timeMin = new Date(checkInDate.getTime() - padding * ONE_DAY_MS);
    const timeMax = new Date(checkOutDate.getTime() + padding * ONE_DAY_MS);
    return { timeMin: timeMin.toISOString(), timeMax: timeMax.toISOString() };
  }

  const now = new Date();
  const timeMin = new Date(now.getTime());
  timeMin.setUTCFullYear(timeMin.getUTCFullYear() - BOOKING_ACTION_CONFIG.wideFallbackWindowYears);
  const timeMax = new Date(now.getTime());
  timeMax.setUTCFullYear(timeMax.getUTCFullYear() + BOOKING_ACTION_CONFIG.wideFallbackWindowYears);

  return { timeMin: timeMin.toISOString(), timeMax: timeMax.toISOString() };
}

/**
 * parseAllDayEventBoundary — parses a raw Calendar API all-day event
 * boundary object (shaped { date: 'YYYY-MM-DD' }, as returned for a
 * fromGmail "stay" event's start/end) into a UTC Date, or null if the
 * boundary is missing or its date string is unparseable. Pure, no GAS
 * globals.
 */
function parseAllDayEventBoundary(boundary) {
  if (!boundary || !boundary.date) {
    return null;
  }

  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(boundary.date);
  if (!match) {
    return null;
  }

  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

/**
 * eventDateRangeOverlaps — true when the (tolerance-padded) booking window
 * [checkInDate - toleranceDays, checkOutDate + toleranceDays] overlaps the
 * calendar event's [eventStart, eventEnd) window, using a standard
 * half-open interval overlap check (a1 < b2 && b1 < a2).
 *
 * `eventStart`/`eventEnd` are the RAW Calendar API start/end objects
 * (each shaped { date: 'YYYY-MM-DD' } for the all-day "stay" events
 * fromGmail creates — see parseAllDayEventBoundary). CRITICAL:
 * `eventEnd.date` is EXCLUSIVE (the day AFTER the last night — e.g. a
 * checkout on Sept 13 shows as end.date "2026-09-14"), matching the same
 * exclusive-end convention this codebase already uses elsewhere for
 * all-day events (see buildEventResource in src/05-action-ics-import.js).
 * This is treated as the natural exclusive upper bound of the event's
 * half-open interval — never re-interpreted as an inclusive checkout date.
 *
 * `toleranceDays` defaults to BOOKING_ACTION_CONFIG.eventOverlapToleranceDays
 * (currently 1) — a tight, owner-tunable tolerance for parsing/rounding
 * slop only, NOT the earlier wide fallback window padding (see
 * computeSearchWindow above, a different, unrelated use of "padding").
 * `toleranceDays` remains an optional explicit override — the config value
 * is only the default.
 *
 * Returns false (conservative — never a false positive) if either
 * boundary is unparseable or either check-in/check-out date is missing or
 * invalid. Pure, no GAS globals.
 */
function eventDateRangeOverlaps(eventStart, eventEnd, checkInDate, checkOutDate, toleranceDays) {
  const tolerance = toleranceDays === undefined ? BOOKING_ACTION_CONFIG.eventOverlapToleranceDays : toleranceDays;

  const eventStartDate = parseAllDayEventBoundary(eventStart);
  const eventEndDate = parseAllDayEventBoundary(eventEnd);

  if (!eventStartDate || !eventEndDate || !isValidDate(checkInDate) || !isValidDate(checkOutDate)) {
    return false;
  }

  const ourLowMs = checkInDate.getTime() - tolerance * ONE_DAY_MS;
  const ourHighMs = checkOutDate.getTime() + tolerance * ONE_DAY_MS;

  return eventStartDate.getTime() < ourHighMs && ourLowMs < eventEndDate.getTime();
}

/**
 * eventSummaryOrLocationContainsHotelName — case-insensitive substring
 * check: true if `hotelName` (non-null, non-empty) is contained in
 * `event.summary` OR `event.location`, false otherwise. A null/empty
 * `hotelName` always returns false (never match everything by accident —
 * an empty needle would otherwise trivially "match" any string). Pure, no
 * GAS globals.
 */
function eventSummaryOrLocationContainsHotelName(event, hotelName) {
  if (!hotelName) {
    return false;
  }

  const needle = String(hotelName).toLowerCase();
  const summary = String((event && event.summary) || '').toLowerCase();
  const location = String((event && event.location) || '').toLowerCase();

  return summary.indexOf(needle) !== -1 || location.indexOf(needle) !== -1;
}

/**
 * findFuzzyMatchingEvent — the fuzzy-matching fallback for a fromGmail
 * "stay" event, whose description never carries the real confirmation
 * number (see the class-level JSDoc's fuzzy-matching note). Takes an
 * ALREADY-FETCHED array of event objects (the caller does the live
 * Calendar.Events.list enumeration — this function only filters/picks
 * from an in-memory array, so it is fully unit-testable without any GAS
 * global). Returns the FIRST event in `events` where BOTH
 * eventSummaryOrLocationContainsHotelName AND eventDateRangeOverlaps are
 * true, or null if none match. Pure, no GAS globals.
 */
function findFuzzyMatchingEvent(events, hotelName, checkInDate, checkOutDate) {
  const list = events || [];

  for (let i = 0; i < list.length; i++) {
    const event = list[i];
    if (
      eventSummaryOrLocationContainsHotelName(event, hotelName) &&
      eventDateRangeOverlaps(event.start, event.end, checkInDate, checkOutDate)
    ) {
      return event;
    }
  }

  return null;
}

/**
 * listCalendarEventsPaginated — enumerates ALL items of
 * Calendar.Events.list(calendarId, params) across every page, following
 * `nextPageToken` (passed back in as `pageToken`) until a page returns
 * none, so an event near a page boundary is never missed. Returns the
 * flattened array of every page's items.
 *
 * MULTI-CALENDAR ROUTING: `calendarId` is an EXPLICIT parameter (not a
 * bare read of CONFIG.calendarId internally) — the caller resolves it
 * once (resolveBookingCalendarId, called once at the top of
 * handleConfirmation/handleCancellation) and threads it down through
 * every downstream call, so every Calendar API call made while processing
 * ONE message consistently targets the SAME resolved calendar. GAS-only
 * (Calendar global) — signature-threading is unit-tested via a fake
 * Calendar global (test/calendar-routing.test.js); the matching/pagination
 * behavior itself is proven only by the live checkpoint.
 */
function listCalendarEventsPaginated(params, calendarId) {
  const allItems = [];
  let pageToken = undefined;

  do {
    const requestParams = Object.assign({}, params);
    if (pageToken) {
      requestParams.pageToken = pageToken;
    }

    const response = Calendar.Events.list(calendarId, requestParams);
    const items = (response && response.items) || [];
    allItems.push.apply(allItems, items);

    pageToken = response && response.nextPageToken;
  } while (pageToken);

  return allItems;
}

/**
 * findEventByConfirmationTag — the EXACT-match lookup: enumerates (paginated,
 * via listCalendarEventsPaginated, against the given `calendarId` — see
 * that function's own JSDoc for the multi-calendar-routing threading this
 * participates in) Calendar.Events.list filtered by
 * `privateExtendedProperty: 'confirmationNumber=' + confirmationNumber`
 * over the bounded [timeMin, timeMax) window, and returns the first item
 * found, or null. This filter is precise by construction (the API itself
 * matches on the tag) — no further substring/fuzzy check is needed.
 * GAS-only (Calendar global) — signature-threading is unit-tested via a
 * fake Calendar global; the matching behavior itself is proven only by
 * the live checkpoint.
 */
function findEventByConfirmationTag(timeMin, timeMax, confirmationNumber, calendarId) {
  const items = listCalendarEventsPaginated(
    {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      privateExtendedProperty: 'confirmationNumber=' + confirmationNumber,
    },
    calendarId
  );

  return items.length > 0 ? items[0] : null;
}

/**
 * findOrTagMatchingEvent — the single shared dedup mechanism for both the
 * add (search-before-create) and remove (search-before-delete) paths,
 * replacing the prior findMatchingCalendarEvent (which relied on a
 * description-substring match — since removed as dead code — first via a
 * `q` free-text search, found unreliable for bare numeric confirmation
 * numbers, THEN via a bounded-enumeration description-substring match,
 * found to still fail because a fromGmail auto-created "stay" event's
 * description is a useless Google-generic placeholder that NEVER carries
 * the real confirmation number — see the class-level JSDoc for the full
 * investigation).
 *
 * Two-layer lookup:
 *   1. EXACT tag match first (findEventByConfirmationTag): if any event in
 *      the window already carries a private extendedProperty
 *      `confirmationNumber=<confirmationNumber>`, return it immediately —
 *      no fuzzy matching or patching needed, this is already precise.
 *   2. FUZZY fallback: if no exact tag match, enumerate the SAME bounded
 *      window with a plain listing (no `q`, no privateExtendedProperty
 *      filter) and hand the collected items to the pure
 *      findFuzzyMatchingEvent(items, hotelName, checkInDate, checkOutDate)
 *      helper. If it finds a match, PATCH that event to add the
 *      confirmationNumber tag (Calendar.Events.patch) so every FUTURE
 *      lookup for this same booking hits the exact-match path instead —
 *      then return the (now-tagged) event.
 *
 * Returns null if neither layer finds anything — both call sites already
 * handle a null match as documented (silent no-op for cancellation,
 * create-new for confirmation).
 *
 * MULTI-CALENDAR ROUTING: `calendarId` is an EXPLICIT parameter, threaded
 * down to findEventByConfirmationTag/listCalendarEventsPaginated and to
 * the Calendar.Events.patch call below — the same resolved calendar
 * (resolveBookingCalendarId, resolved ONCE at the top of the calling
 * handler) is used consistently for every Calendar API call made while
 * processing one message. GAS-only (Calendar global) — signature-threading
 * is unit-tested via a fake Calendar global (test/calendar-routing.test.js);
 * the layering/patching wiring itself is proven only by the live
 * checkpoint; the pure matching logic within each layer IS unit-tested
 * (findFuzzyMatchingEvent and its helpers, above).
 */
function findOrTagMatchingEvent(timeMin, timeMax, confirmationNumber, hotelName, checkInDate, checkOutDate, calendarId) {
  const exactMatch = findEventByConfirmationTag(timeMin, timeMax, confirmationNumber, calendarId);
  if (exactMatch) {
    return exactMatch;
  }

  const candidates = listCalendarEventsPaginated({ timeMin: timeMin, timeMax: timeMax, singleEvents: true }, calendarId);
  const fuzzyMatch = findFuzzyMatchingEvent(candidates, hotelName, checkInDate, checkOutDate);
  if (!fuzzyMatch) {
    return null;
  }

  Calendar.Events.patch(
    { extendedProperties: { private: { confirmationNumber: confirmationNumber } } },
    calendarId,
    fuzzyMatch.id
  );

  return fuzzyMatch;
}

/**
 * handleCancellation — reads the plain body of `message`, extracts the
 * confirmation number via
 * extractLabeledNumber(body, getBookingLabels('confirmationNumberLabels')).
 * A null extraction is a SILENT no-op (console.log diagnostic, no throw) —
 * an un-matchable cancellation must never fail the action. Otherwise parses
 * check-in/check-out via extractCheckInOutDateAcrossLanguagePacks (which
 * finds BOTH the matching label AND the correct per-language date parser
 * together — see that function's JSDoc; the cancellation email's real
 * format has no time window on its Check-in/Check-out lines in any
 * language, just dates — the shared midnight-UTC fallback already handles
 * that, still yielding valid dates for windowing), computes the search
 * window via computeSearchWindow and the hotel name via extractHotelName
 * (getBookingLabels('cancellationHotelNameSeparators') — the cancellation
 * subject's separator/side may differ by language, e.g. English's " for "
 * vs. Czech's " – "), then finds the matching event via
 * findOrTagMatchingEvent (exact confirmationNumber tag first, fuzzy
 * hotel-name+date-overlap fallback second — see that function's JSDoc); if
 * found, deletes it via Calendar.Events.remove; if not found, this is ALSO
 * a silent no-op (owner-confirmed: no error, no failure notification —
 * see must_haves).
 *
 * MULTI-CALENDAR ROUTING: `calendarId` is resolved ONCE, as the very first
 * statement (resolveBookingCalendarId(BOOKING_ACTION_CONFIG,
 * CONFIG.calendarId)), and threaded explicitly through
 * findOrTagMatchingEvent and Calendar.Events.remove — never re-read from
 * config at each call site — so every Calendar API call made while
 * processing this one message consistently targets the same resolved
 * calendar. GAS-only (GmailMessage/Calendar globals) — not unit-tested,
 * proven only by the live checkpoint.
 */
function handleCancellation(message) {
  const calendarId = resolveBookingCalendarId(BOOKING_ACTION_CONFIG, CONFIG.calendarId);
  const body = message.getPlainBody();
  const confirmationNumber = extractLabeledNumber(body, getBookingLabels('confirmationNumberLabels'));

  if (confirmationNumber === null) {
    console.log('Booking cancellation: no confirmation number found, skipping (silent no-op).');
    return;
  }

  const checkIn = extractCheckInOutDateAcrossLanguagePacks(body, 'checkInLabels', true);
  const checkOut = extractCheckInOutDateAcrossLanguagePacks(body, 'checkOutLabels', false);
  const window = computeSearchWindow(checkIn, checkOut);
  const hotelName = extractHotelName(message.getSubject(), getBookingLabels('cancellationHotelNameSeparators'));

  const event = findOrTagMatchingEvent(window.timeMin, window.timeMax, confirmationNumber, hotelName, checkIn, checkOut, calendarId);
  if (!event) {
    console.log('Booking cancellation: no matching calendar event for confirmation number ' + confirmationNumber + ', skipping (silent no-op).');
    return;
  }

  Calendar.Events.remove(calendarId, event.id);
}

/**
 * handleConfirmation — extracts the confirmation number from the message
 * body via extractLabeledNumber; a null extraction is a SILENT no-op (never
 * create an un-matchable event). Parses check-in (earliest arrival,
 * useStartTime=true) and check-out (latest departure, useStartTime=false)
 * via extractCheckInOutDateAcrossLanguagePacks FIRST — a date-parsing
 * throw PROPAGATES here, before any calendar read/write (dispatch
 * isolation + notifyOnFailure at the framework level handle it like any
 * other action failure; it is never swallowed mid-way). The dates (and
 * the hotel name, derived from the subject via
 * extractHotelName/getBookingLabels('confirmationHotelNameSeparators'),
 * falling back to a body location line if absent) are then used both for the
 * safety-net existence check (findOrTagMatchingEvent — exact
 * confirmationNumber tag first, fuzzy hotel-name+date-overlap fallback
 * second) AND, if no existing event is found, as the new event's
 * start/end/summary. If a matching event already exists (typically
 * Google's own native fromGmail detection, now matched even though its
 * description never carries the real confirmation number — see
 * findOrTagMatchingEvent's JSDoc), logs and returns WITHOUT creating a
 * second event — this is not a duplicate path. Otherwise derives the
 * address from the body's location line (falling back to the hotel name),
 * extracts the optional PIN, appends the WHOLE "Reservation details"
 * section (see extractReservationDetailsSection) verbatim when present
 * (silently skipped when the heading is not found — never throws over a
 * missing/differently-formatted section), and inserts a plain (no
 * iCalUID) Advanced Calendar event resource, itself immediately tagged
 * with extendedProperties.private.confirmationNumber (so an event WE
 * create is exact-tagged from birth — never needs fuzzy-matching later).
 *
 * TIMEZONE: booking.com confirmation emails carry NO timezone indicator
 * anywhere in their check-in/check-out lines — the parsed wall-clock time
 * (e.g. "2:00 PM") is the LOCAL time at the property, and there is no way
 * to derive the correct IANA zone from the email content itself (unlike
 * the ICS action, which resolves TZID from the .ics's own embedded
 * VTIMEZONE data — see the class-level JSDoc in
 * src/05-action-ics-import.js; booking.com emails have no equivalent
 * embedded timezone data at all). Rather than a hardcoded config
 * assumption, this is derived live from the TARGET CALENDAR's own
 * configured timezone via
 * `CalendarApp.getCalendarById(calendarId).getTimeZone()` —
 * correct by construction for the common case (the calendar and the stay
 * are in the same region). KNOWN LIMITATION (accepted, no per-booking
 * geocoding): a stay booked in a materially different timezone than the
 * calendar's own setting cannot be detected from the email content alone.
 * formatLocalWallClockIso's re-extracted wall-clock digits are paired with
 * this live-derived timezone string on the resource's start/end.
 *
 * MULTI-CALENDAR ROUTING: `calendarId` is resolved ONCE, as the very first
 * statement (resolveBookingCalendarId(BOOKING_ACTION_CONFIG,
 * CONFIG.calendarId)), and threaded explicitly through
 * findOrTagMatchingEvent, the CalendarApp.getTimeZone() lookup, and
 * Calendar.Events.insert — never re-read from config at each call site —
 * so every Calendar API call made while processing this one message
 * consistently targets the same resolved calendar. GAS-only
 * (GmailMessage/Calendar/CalendarApp globals) — not unit-tested, proven
 * only by the live checkpoint.
 */
function handleConfirmation(message) {
  const calendarId = resolveBookingCalendarId(BOOKING_ACTION_CONFIG, CONFIG.calendarId);
  const body = message.getPlainBody();
  const confirmationNumber = extractLabeledNumber(body, getBookingLabels('confirmationNumberLabels'));

  if (confirmationNumber === null) {
    console.log('Booking confirmation: no confirmation number found, skipping (silent no-op).');
    return;
  }

  // Date parsing throws PROPAGATE intentionally — see dispatch isolation.
  // Computed FIRST: both the existence check below and the eventual
  // Calendar.Events.insert (if no existing event is found) need them.
  const checkIn = extractCheckInOutDateAcrossLanguagePacks(body, 'checkInLabels', true);
  const checkOut = extractCheckInOutDateAcrossLanguagePacks(body, 'checkOutLabels', false);

  const hotelNameFromSubject = extractHotelName(message.getSubject(), getBookingLabels('confirmationHotelNameSeparators'));
  const locationLine = nextNonEmptyLineAfterLabel(body, getBookingLabels('locationLabels'));
  const hotelName = hotelNameFromSubject || locationLine || 'Booking';

  const window = computeSearchWindow(checkIn, checkOut);
  const existingEvent = findOrTagMatchingEvent(window.timeMin, window.timeMax, confirmationNumber, hotelName, checkIn, checkOut, calendarId);
  if (existingEvent) {
    console.log('Booking confirmation: event for confirmation number ' + confirmationNumber + ' already exists, skipping (safety-net, not a duplicate path).');
    return;
  }

  const address = locationLine || hotelName;
  const pin = extractLabeledNumber(body, getBookingLabels('pinLabels'));

  const descriptionLines = ['Confirmation number: ' + confirmationNumber];
  if (pin !== null) {
    descriptionLines.push('PIN: ' + pin);
  }

  // Append the whole "Reservation details" section verbatim (check-in/out,
  // room type, guest count, address, cancellation policy, etc.) when
  // present. A null result (heading not found — e.g. a differently
  // formatted or non-English email) is skipped silently, same
  // graceful-degradation style as the PIN check above — never throw over
  // a missing section.
  const reservationDetails = extractReservationDetailsSection(
    body,
    getBookingLabels('reservationDetailsHeadingLabels'),
    getBookingLabels('reservationDetailsEndHeadingLabels')
  );
  if (reservationDetails !== null) {
    descriptionLines.push('');
    descriptionLines.push(reservationDetails);
  }

  // Derived live from the TARGET CALENDAR's own configured timezone —
  // NOT a hardcoded config assumption (see this function's JSDoc,
  // "TIMEZONE"). GAS-only call, not unit-testable in Node.
  const timeZone = CalendarApp.getCalendarById(calendarId).getTimeZone();

  const resource = {
    summary: hotelName,
    description: descriptionLines.join('\n'),
    location: address,
    // formatLocalWallClockIso + an explicit timeZone (NOT checkIn/checkOut's
    // own .toISOString(), which would stamp a literal UTC 'Z' suffix onto
    // digits that are actually wall-clock LOCAL time at the property —
    // see formatLocalWallClockIso's JSDoc for the live-reported 2-hour-off
    // bug this fixes).
    start: { dateTime: formatLocalWallClockIso(checkIn), timeZone: timeZone },
    end: { dateTime: formatLocalWallClockIso(checkOut), timeZone: timeZone },
    // Tag this event from birth so it is found via the exact-match path
    // (findEventByConfirmationTag) on any future lookup — never needs the
    // fuzzy fallback, since we know its confirmation number directly.
    extendedProperties: {
      private: {
        confirmationNumber: confirmationNumber,
      },
    },
  };

  Calendar.Events.insert(resource, calendarId);
}

/**
 * BOOKING_MANAGEMENT_ACTION — the booking.com management action
 * descriptor. Carries its own config block (BOOKING_ACTION_CONFIG),
 * independent of CONFIG and of any other action's config.
 *
 * appliesTo(thread): returns a literal boolean — true when any thread
 * message is from an allowed sender (bookingIsAllowedSender against
 * config.senderAllowList) AND its subject matches an add-subject when
 * config.addToCalendar.enabled (getBookingLabels('addToCalendarSubjectContains')),
 * OR matches a remove-subject when config.removeFromCalendar.enabled
 * (getBookingLabels('removeFromCalendarSubjectContains')). dispatchActions
 * only skips on a strict `=== false`, so a literal boolean is required.
 *
 * run(thread): iterates every message on the thread; for each, skips
 * unless the sender is allowed; if the subject matches the add list and
 * addToCalendar.enabled, dispatches to handleConfirmation; else if it
 * matches the remove list and removeFromCalendar.enabled, dispatches to
 * handleCancellation; a message matching neither is skipped gracefully
 * (never throws for a non-matching message — a thread can carry a message
 * that doesn't match either sub-behavior if a sub-behavior was disabled
 * after appliesTo already returned true for a DIFFERENT message on the
 * same thread).
 */
const BOOKING_MANAGEMENT_ACTION = {
  name: 'booking-management',

  // GETTER, not a plain literal property — see this file's class-level
  // "CONFIG SPLIT" note. Not evaluated at object-construction time, only
  // when something reads `.config`, which happens lazily inside function
  // bodies (dispatchActions, notifyOwnerOfFailure) long after every
  // project file has loaded — so this is safe regardless of whether
  // 06-action-cfg-booking-com-management.js or this file loads first
  // alphabetically (and here, this file actually loads FIRST — the
  // getter is what makes that direction safe too).
  get config() {
    return BOOKING_ACTION_CONFIG;
  },

  appliesTo: function (thread) {
    const config = BOOKING_ACTION_CONFIG;
    const messages = thread.getMessages();

    for (let i = 0; i < messages.length; i++) {
      const message = messages[i];
      if (!bookingIsAllowedSender(message.getFrom(), config.senderAllowList)) {
        continue;
      }

      const subject = message.getSubject();
      const matchesAdd = config.addToCalendar.enabled && subjectContainsAny(subject, getBookingLabels('addToCalendarSubjectContains'));
      const matchesRemove = config.removeFromCalendar.enabled && subjectContainsAny(subject, getBookingLabels('removeFromCalendarSubjectContains'));

      if (matchesAdd || matchesRemove) {
        return true;
      }
    }

    return false;
  },

  run: function (thread) {
    const config = BOOKING_ACTION_CONFIG;
    const messages = thread.getMessages();

    messages.forEach(function (message) {
      if (!bookingIsAllowedSender(message.getFrom(), config.senderAllowList)) {
        return;
      }

      const subject = message.getSubject();

      if (config.addToCalendar.enabled && subjectContainsAny(subject, getBookingLabels('addToCalendarSubjectContains'))) {
        handleConfirmation(message);
        return;
      }

      if (config.removeFromCalendar.enabled && subjectContainsAny(subject, getBookingLabels('removeFromCalendarSubjectContains'))) {
        handleCancellation(message);
        return;
      }

      // Matches neither sub-behavior for this message — skip gracefully.
    });
  },
};

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Single merged export — extended by later tasks in this same file as more
// pure helpers are implemented.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    BOOKING_LANGUAGE_PACKS: BOOKING_LANGUAGE_PACKS,
    getBookingLabels: getBookingLabels,
    bookingExtractEmailAddress: bookingExtractEmailAddress,
    bookingIsAllowedSender: bookingIsAllowedSender,
    resolveBookingCalendarId: resolveBookingCalendarId,
    subjectContainsAny: subjectContainsAny,
    extractLabeledNumber: extractLabeledNumber,
    MONTH_NAME_TO_INDEX: MONTH_NAME_TO_INDEX,
    parse12HourTime: parse12HourTime,
    parseBookingDateLine: parseBookingDateLine,
    findLabeledLine: findLabeledLine,
    extractCheckInOutDate: extractCheckInOutDate,
    findLabeledLineAcrossLanguagePacks: findLabeledLineAcrossLanguagePacks,
    extractCheckInOutDateAcrossLanguagePacks: extractCheckInOutDateAcrossLanguagePacks,
    extractHotelName: extractHotelName,
    nextNonEmptyLineAfterLabel: nextNonEmptyLineAfterLabel,
    computeSearchWindow: computeSearchWindow,
    eventDateRangeOverlaps: eventDateRangeOverlaps,
    eventSummaryOrLocationContainsHotelName: eventSummaryOrLocationContainsHotelName,
    findFuzzyMatchingEvent: findFuzzyMatchingEvent,
    formatLocalWallClockIso: formatLocalWallClockIso,
    extractReservationDetailsSection: extractReservationDetailsSection,
    listCalendarEventsPaginated: listCalendarEventsPaginated,
    findEventByConfirmationTag: findEventByConfirmationTag,
    findOrTagMatchingEvent: findOrTagMatchingEvent,
    BOOKING_MANAGEMENT_ACTION: BOOKING_MANAGEMENT_ACTION,
  };
}

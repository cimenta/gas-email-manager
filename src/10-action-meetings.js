/**
 * MEETINGS_ACTION — quick-260820-g4r NEW ACTION: recognizes general
 * meeting-invitation emails whose date/time/location live entirely in the
 * plain-text BODY (`message.getPlainBody()`), never a structured `.ics`
 * attachment, and creates ONE calendar event per matching message.
 *
 * NO-.ICS SCOPE (D-04): this action never claims a message that carries a
 * `.ics` attachment — ICS_CALENDAR_ACTION (src/05-action-ics-import.js)
 * already owns those, and its own config/files are deliberately untouched
 * by this task. The gate lives entirely in this file's own
 * meetingsMessageHasIcsAttachment / findMeetingProcessingJobs.
 *
 * PLUGGABLE "SYSTEMS" ARCHITECTURE (D-01, D-11), mirroring
 * TICKETING_PORTALS_ACTION (src/07-action-ticketing-portals.js): ONE action
 * file plus one sibling config file, NOT one file per system. Every
 * supported meeting-invitation "system" (Teamio ships first) is a config
 * entry (`{ domainPattern, calendarId }`) plus a dedicated parser function
 * inside THIS file, matched through MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN,
 * a registry object keyed by the SAME `domainPattern` string used in the
 * config entry — the identical pattern-is-the-registry-key convention
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL already establishes. Adding a
 * further system needs ONE new config entry, ONE new parser function
 * registered in MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN, and ONE new detector
 * registered in MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN (debug
 * teamio-non-invite-error — a sender match alone is NOT evidence that an email
 * is an invitation; see that registry's own class-level doc) — no change to the
 * matching logic itself. Both language
 * packs (Czech and English) also live in this one file (D-11) — no
 * `10-lang-*.js` files: src/06-lang-*.js are documented as deliberately
 * booking-action-only, and this action's pluggable-systems architecture
 * follows the ticketing action's one-file model instead.
 *
 * DETECTION BY SENDER DOMAIN, NOT EXACT ADDRESS (D-02) — the one respect in
 * which this action's matching differs from every prior action in this
 * codebase. A per-system `domainPattern` supports a leading `*.` subdomain
 * wildcard: `*.teamio.com` matches `recruit.teamio.com` (a real subdomain)
 * AND the bare apex `teamio.com` (the useful reading for an owner who types
 * one pattern and expects the whole domain covered) — but matches only on a
 * REAL DOT BOUNDARY, so `notteamio.com` (a different domain that merely
 * ends with the same letters) and `teamio.com.evil.net` (a longer domain
 * that merely starts with the pattern's suffix) do NOT match. A pattern
 * with no `*.` prefix is exact, case-insensitive equality only. See
 * meetingsDomainMatchesPattern below for the implementation, and this
 * project's ICS action for the established "From-header match is a
 * non-cryptographic convenience filter, not a security boundary" precedent
 * this action's own domain match equally relies on.
 *
 * DEDUP (D-07): free text carries no `iCalUID`, so this action reuses
 * TICKETING_PORTALS_ACTION's proven pattern — tag the created event with
 * `extendedProperties.private.meetingIdentifier`, and before creating,
 * query the target calendar for an event already carrying that exact tag
 * (find-before-create, see findMeetingEventByIdentifier /
 * isDuplicateMeetingInvite). Unlike Ticketmaster CZ's own documented v1
 * limitation (a ticket-purchase confirmation with no stable identifier
 * simply gets NO dedup protection at all), this action's identifier is
 * built by a PURE, deterministic function (buildMeetingIdentifier) from
 * sender + normalized subject + parsed start wall-clock — ALWAYS produced,
 * never null, so every meeting processed by this action gets the dedup
 * safety net, closing that gap for the general case a body-sourced parser
 * can hit.
 *
 * TIMEZONE: wall-clock digits parsed from the body are resolved against the
 * TARGET CALENDAR'S LIVE timezone (`CalendarApp.getCalendarById(calendarId)
 * .getTimeZone()`), never a hardcoded guess — the same "resolve wall-clock
 * digits, pair with a live-derived timeZone field" pattern already proven
 * by every prior action in this codebase (see
 * TICKETING_PORTALS_ACTION's own class-level "TIMEZONE" doc paragraph).
 *
 * GLOBALLY-UNIQUE NAMING WARNING (see the booking.com/ticketing-portals
 * action files' own class-level JSDoc for the full incident this warning
 * originates from): Apps Script concatenates every project file into ONE
 * shared global scope, and `07-action-ticketing-portals.js` — a file that
 * loads BEFORE this one alphabetically ("07-" sorts before "10-") — already
 * defines UN-namespaced `addMinutesToWallClockComponents`,
 * `formatWallClockComponentsIso`, `zeroPadTicketComponent` and
 * `DEFAULT_EVENT_DURATION_MINUTES`. A same-named definition in THIS file
 * would silently take over every call site that resolves those bare names
 * (since this file loads AFTER, its own top-level declarations would win
 * the last-one-wins collision) — a real, already-experienced failure class
 * in this codebase (see the booking.com action's own class-level JSDoc for
 * the first such incident). Every helper this file introduces is therefore
 * `meetings`/`MEETINGS`-prefixed, including its OWN local copies of the
 * wall-clock helpers (meetingsAddMinutesToWallClockComponents /
 * meetingsFormatWallClockComponentsIso / meetingsZeroPad) — never imported
 * or reused across files, per this codebase's established
 * per-action self-containment convention.
 */

// --- pure helpers (D-02, D-10) ----------------------------------------------

/**
 * meetingsExtractEmailAddress — LOCAL copy of the same underlying logic
 * already re-implemented independently in every other action file (see
 * ticketingExtractEmailAddress/mojemenickaExtractEmailAddress for the
 * identical shape). Extracts the bare, trimmed, lowercased email address
 * from a Gmail "From" header value, or from a bare address with no display
 * name. Pure, no GAS globals. Never throws: a null/undefined/empty input
 * returns ''.
 */
function meetingsExtractEmailAddress(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const angleBracketMatch = /<([^>]*)>/.exec(fromHeader);
  const raw = angleBracketMatch ? angleBracketMatch[1] : fromHeader;

  return raw.trim().toLowerCase();
}

/**
 * meetingsExtractSenderDomain — the part of the sender address after the
 * last `@`, lowercased. Built on meetingsExtractEmailAddress, so a
 * display-name `From` header resolves the same domain as a bare address.
 * Pure, no GAS globals. Never throws: an address with no `@` (or an empty/
 * null input) returns ''.
 */
function meetingsExtractSenderDomain(fromHeader) {
  const address = meetingsExtractEmailAddress(fromHeader);
  const atIndex = address.lastIndexOf('@');
  return atIndex === -1 ? '' : address.slice(atIndex + 1);
}

/**
 * meetingsDomainMatchesPattern — D-02's core matching rule. A `domainPattern`
 * beginning with the literal `*.` matches when `domain` equals the
 * remainder (the bare apex) OR ends with a dot plus the remainder (a real
 * subdomain, dot-boundary-anchored) — this is what makes `*.teamio.com`
 * match `recruit.teamio.com` and `teamio.com`, but NOT `notteamio.com`
 * (fails both the equality check and the dot-boundary check: `notteamio.com`
 * does not end with `.teamio.com`) and NOT `teamio.com.evil.net` (that
 * domain ends with `.evil.net`, not `.teamio.com`, so the suffix check
 * anchors correctly on the END of the string, not merely a substring
 * anywhere within it). Any pattern with no `*.` prefix is exact,
 * case-insensitive equality only. Both sides are trimmed and lowercased.
 * Pure, no GAS globals. Never throws on null/empty input — returns false.
 */
function meetingsDomainMatchesPattern(domain, domainPattern) {
  const normalizedDomain = String(domain || '').trim().toLowerCase();
  const normalizedPattern = String(domainPattern || '').trim().toLowerCase();

  if (!normalizedDomain || !normalizedPattern) {
    return false;
  }

  if (normalizedPattern.slice(0, 2) === '*.') {
    const suffix = normalizedPattern.slice(2);
    if (!suffix) {
      return false;
    }
    return normalizedDomain === suffix || normalizedDomain.slice(-(suffix.length + 1)) === '.' + suffix;
  }

  return normalizedDomain === normalizedPattern;
}

/**
 * resolveMeetingSystem — finds the FIRST entry in `meetingSystems` (list
 * order) whose `domainPattern` matches `fromHeader`'s sender domain (via
 * meetingsExtractSenderDomain + meetingsDomainMatchesPattern), mirroring
 * resolveTicketingPortal's own list-order-first-match convention. Returns
 * `null` on no match, or when `meetingSystems` is null/empty. Pure, no GAS
 * globals. Never throws.
 */
function resolveMeetingSystem(fromHeader, meetingSystems) {
  const list = meetingSystems || [];
  const domain = meetingsExtractSenderDomain(fromHeader);

  for (let i = 0; i < list.length; i++) {
    if (meetingsDomainMatchesPattern(domain, list[i].domainPattern)) {
      return list[i];
    }
  }

  return null;
}

/**
 * resolveMeetingsCalendarId — the two-tier `system.calendarId ||
 * defaultCalendarId` resolution, mirroring resolveTicketingCalendarId's own
 * shape and the real live-null-calendar incident that established this
 * fallback pattern project-wide (`CalendarApp.getCalendarById(null)` throws
 * a real error, not merely returns null-ish). Pure, no GAS globals.
 */
function resolveMeetingsCalendarId(system, defaultCalendarId) {
  return (system && system.calendarId) || defaultCalendarId;
}

/**
 * meetingsZeroPad — left-pads `value` with '0' to `length` digits. Pure, no
 * GAS globals. Deliberately namespaced (not a bare `zeroPad`), per this
 * file's own globally-unique-naming convention — see this file's
 * class-level "GLOBALLY-UNIQUE NAMING WARNING".
 */
function meetingsZeroPad(value, length) {
  return String(value).padStart(length, '0');
}

/**
 * meetingsAddMinutesToWallClockComponents — NAMESPACED local copy of
 * addMinutesToWallClockComponents (src/07-action-ticketing-portals.js) —
 * see this file's class-level "GLOBALLY-UNIQUE NAMING WARNING" for why a
 * bare-named copy here would be a real collision hazard, not merely a
 * style choice. Pure: adds `minutes` to a `{ year, month, day, hour,
 * minute }` wall-clock components object (month zero-indexed), returning a
 * NEW components object of the same shape, correctly handling
 * hour/day/month/year rollover, via `Date.UTC` arithmetic used purely as a
 * NEUTRAL zero-offset calculation space (never a real UTC instant — the
 * input components carry no timezone information at all). No GAS globals.
 */
function meetingsAddMinutesToWallClockComponents(components, minutes) {
  const asMs = Date.UTC(components.year, components.month, components.day, components.hour, components.minute) + minutes * 60000;
  const rolled = new Date(asMs);

  return {
    year: rolled.getUTCFullYear(),
    month: rolled.getUTCMonth(),
    day: rolled.getUTCDate(),
    hour: rolled.getUTCHours(),
    minute: rolled.getUTCMinutes(),
  };
}

/**
 * meetingsFormatWallClockComponentsIso — NAMESPACED local copy of
 * formatWallClockComponentsIso (src/07-action-ticketing-portals.js). Formats
 * a `{ year, month, day, hour, minute }` wall-clock components object
 * (month zero-indexed) as a zero-padded literal string
 * `'YYYY-MM-DDTHH:MM:00'` — DELIBERATELY with NO trailing `Z` and NO
 * timezone offset, meant to be paired with an explicit Calendar API
 * `timeZone` field (see this file's class-level "TIMEZONE" doc). Pure, no
 * GAS globals.
 */
function meetingsFormatWallClockComponentsIso(components) {
  return (
    meetingsZeroPad(components.year, 4) +
    '-' +
    meetingsZeroPad(components.month + 1, 2) +
    '-' +
    meetingsZeroPad(components.day, 2) +
    'T' +
    meetingsZeroPad(components.hour, 2) +
    ':' +
    meetingsZeroPad(components.minute, 2) +
    ':00'
  );
}

// MEETINGS_MAX_DESCRIPTION_LINKS (D-09) — the cap applied by
// meetingsHarvestBodyLinks below, bounding the blast radius of a
// link-stuffed body (see this task's own threat register, T-g4r-04).
const MEETINGS_MAX_DESCRIPTION_LINKS = 10;

/**
 * meetingsHarvestBodyLinks — D-09: returns every distinct `http(s)` URL
 * found in `bodyText`, in source order, excluding `excludeUrl` (the URL
 * already used as the event location, so it is not repeated inside the
 * description's own `Links:` block), capped at
 * MEETINGS_MAX_DESCRIPTION_LINKS. A trailing run of common sentence/markup
 * punctuation (`) , . ;`) is stripped from each match, since a URL is
 * frequently followed immediately by such a character in free-text prose.
 * Pure, no GAS globals. Never throws: a null/empty bodyText yields [].
 */
function meetingsHarvestBodyLinks(bodyText, excludeUrl) {
  const text = String(bodyText || '');
  const urlPattern = /https?:\/\/[^\s]+/g;
  const seen = {};
  const links = [];
  let match = urlPattern.exec(text);

  while (match !== null) {
    const url = match[0].replace(/[),.;]+$/, '');

    if (url && url !== excludeUrl && !seen[url]) {
      seen[url] = true;
      links.push(url);
      if (links.length >= MEETINGS_MAX_DESCRIPTION_LINKS) {
        break;
      }
    }

    match = urlPattern.exec(text);
  }

  return links;
}

// MEETINGS_IDENTIFIER_MAX_SUBJECT_LENGTH — the documented bound
// buildMeetingIdentifier truncates a normalized subject to, so an
// attacker-controlled (or merely very long) subject cannot produce an
// unbounded identifier string.
const MEETINGS_IDENTIFIER_MAX_SUBJECT_LENGTH = 80;

/**
 * buildMeetingIdentifier — D-07's dedup key. A PURE, deterministic
 * composite of the lowercased `senderEmail`, `subject` lowercased with
 * whitespace runs collapsed to a single space and truncated to
 * MEETINGS_IDENTIFIER_MAX_SUBJECT_LENGTH, and the formatted start
 * wall-clock (via meetingsFormatWallClockComponentsIso) — ALWAYS produced,
 * never null, deliberately closing the gap Ticketmaster CZ's own
 * `ticketIdentifier: null` documents as a v1 limitation (see this file's
 * class-level "DEDUP" doc). Every `=` character is stripped from the
 * result, because the identifier is interpolated into a
 * `privateExtendedProperty: 'meetingIdentifier=' + value` Calendar API
 * query string (see findMeetingEventByIdentifier below) — an unstripped
 * `=` in an attacker-controlled subject could otherwise smuggle a second
 * `=` into that query and corrupt the lookup (this task's own threat
 * register, T-g4r-06). No `Utilities.computeDigest` is used deliberately:
 * that is a GAS global, and hashing here would make this function
 * untestable under Node for no real benefit — a deterministic composite
 * string built from already-bounded inputs is sufficient and simpler. Pure,
 * no GAS globals. Never throws.
 */
function buildMeetingIdentifier(senderEmail, subject, startComponents) {
  const normalizedSender = String(senderEmail || '').trim().toLowerCase();
  const normalizedSubject = String(subject || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .slice(0, MEETINGS_IDENTIFIER_MAX_SUBJECT_LENGTH);
  const startText = meetingsFormatWallClockComponentsIso(startComponents);

  return (normalizedSender + '|' + normalizedSubject + '|' + startText).replace(/=/g, '');
}

// --- language packs (D-05, D-11) --------------------------------------------

/**
 * meetingsFindLabelLine — searches `lines` (already trimmed) for the first
 * line that starts with any of `labels` (case-insensitive) immediately
 * followed by `:`, returning `{ label, value, lineText }` (value: the text
 * after the colon, trimmed) or `null` if none match. Deliberately anchored
 * at the START of each line (not a `\b`-bounded regex scan across the whole
 * body) — several of this pack's labels (`Kdy`, `Čas`, `Kde`) begin or end
 * with a non-ASCII letter, and JS's `\b` word-boundary only recognizes
 * ASCII word characters by default, so a `\b`-based match can silently fail
 * exactly where a label starts/ends in a non-ASCII letter (the same real
 * incident documented in src/06-action-booking-com-management.js's own
 * history, "Informace o ceně"). Anchoring on the line's own start avoids
 * `\b` entirely. Pure, no GAS globals.
 */
function meetingsFindLabelLine(lines, labels) {
  for (let i = 0; i < lines.length; i++) {
    for (let j = 0; j < labels.length; j++) {
      const prefix = labels[j] + ':';
      if (lines[i].slice(0, prefix.length).toLowerCase() === prefix.toLowerCase()) {
        return { label: labels[j], value: lines[i].slice(prefix.length).trim(), lineText: lines[i] };
      }
    }
  }

  return null;
}

/**
 * meetingsExtractHtmlLabelValue — (round 2, live-test-driven bug fix,
 * quick-260820-g4r) reads a label's value out of an HTML body's own
 * markup, for the ONE real case where the plain-text body's own version of
 * that value is not what should end up on the calendar event.
 *
 * REAL BUG THIS FIXES: Teamio's ESP rewrites every URL-shaped string into
 * an opaque `track.teamio.com` click-tracking REDIRECT when generating the
 * plain-text alternative from the HTML — confirmed directly against the
 * owner's real `.eml` (quoted-printable, multipart/alternative). The
 * text/plain part's own "Kde:" line therefore never carries the real,
 * human-readable meeting URL at all; that URL exists ONLY in the
 * text/html part's own "Kde:" table cell, as plain (non-hyperlinked) text
 * wrapped in a `<strong>` tag — e.g.
 * `<td>Kde:</td><td><strong>https://teams.microsoft.com/meet/...</strong></td>`.
 * A parser reading `getPlainBody()` alone (this file's original, round-1
 * design) therefore extracted the opaque tracking redirect as the event's
 * location, exactly what the owner reported live.
 *
 * Searches `htmlBodyText` for the FIRST occurrence of any of `labels`
 * immediately followed by `:` (same case-insensitive label-matching
 * philosophy as meetingsFindLabelLine, generalized to any of a pack's
 * label lists — e.g. `pack.whereLabels` — so this is not hardcoded to
 * "Kde" alone), then returns the text content of the NEXT `<strong>...
 * </strong>` element found after that point (any remaining inner tags
 * stripped, defensively, in case a future real email nests further
 * markup there), trimmed. Returns `null` when the label is not found, no
 * `<strong>` element follows it, or the extracted text is empty — NEVER
 * throws, and never partially applies: a caller falling back to the
 * plain-text value on `null` is the correct, existing behavior for a
 * system whose HTML body was not supplied or does not carry this
 * particular table-cell shape. Pure, no GAS globals.
 */
function meetingsExtractHtmlLabelValue(htmlBodyText, labels) {
  const html = String(htmlBodyText || '');
  if (!html) {
    return null;
  }

  for (let i = 0; i < labels.length; i++) {
    const label = labels[i] + ':';
    const labelIndex = html.indexOf(label);
    if (labelIndex === -1) {
      continue;
    }

    const afterLabel = html.slice(labelIndex + label.length);
    const strongMatch = /<strong>([\s\S]*?)<\/strong>/i.exec(afterLabel);
    if (!strongMatch) {
      continue;
    }

    const text = strongMatch[1].replace(/<[^>]*>/g, '').trim();
    if (text) {
      return text;
    }
  }

  return null;
}

/**
 * MEETINGS_CS_MONTH_GENITIVE_TO_INDEX — lowercase Czech month name,
 * GENITIVE grammatical case (the form the real Teamio email uses in a date
 * like "24. srpna 2026"), -> zero-indexed month, matching Date.UTC's
 * convention. NAMESPACED copy of the same 12-entry table already
 * established in src/06-lang-cs.js's CZECH_MONTH_GENITIVE_TO_INDEX (never
 * imported across files, per this file's own self-containment convention —
 * see the class-level "GLOBALLY-UNIQUE NAMING WARNING").
 *
 * VERIFICATION STATUS: only 'srpna' (August) is EMPIRICALLY VERIFIED
 * against the real owner-supplied Teamio email this pack was built from.
 * The remaining 11 genitive forms are standard, unambiguous Czech grammar
 * and are included on that grammatical basis, but have NOT yet been
 * observed in a real Teamio invitation — flag/fix this table if a future
 * real email surfaces an unexpected month-name mismatch, the same
 * treatment CZECH_MONTH_GENITIVE_TO_INDEX itself documents for its own
 * 11 unverified entries.
 */
const MEETINGS_CS_MONTH_GENITIVE_TO_INDEX = {
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
 * meetingsParseCsDateText — the Czech pack's `parseDateText`. Searches
 * `value` UNANCHORED for `D. <genitive month name> YYYY` (e.g.
 * "24. srpna 2026") via a Unicode-aware `[\p{L}]+` letter-run for the month
 * token (the `u` regex flag, rather than a `\b` word boundary — see
 * meetingsFindLabelLine's own doc for why `\b` is unsafe here: Czech month
 * names can both start and end in non-ASCII letters). Being unanchored is
 * exactly what lets this skip a leading weekday name (with or without a
 * trailing comma, e.g. "Pondělí, 24. srpna 2026" or plain "24. srpna 2026")
 * without needing to parse or recognize the weekday at all. Returns
 * `{ year, month, day }` (month zero-indexed). Throws a controlled Error
 * (message only — the caller wraps it with the full raw body text) when no
 * day/month/year shape is found at all, or when the matched month token is
 * not a recognized Czech genitive month name (naming the bad token). Pure,
 * no GAS globals.
 */
function meetingsParseCsDateText(value) {
  const text = String(value || '');
  const match = /(\d{1,2})\.\s*([\p{L}]+)\s+(\d{4})/u.exec(text);

  if (!match) {
    throw new Error('no recognizable date found in "' + text + '"');
  }

  const monthToken = match[2].toLowerCase();
  if (!Object.prototype.hasOwnProperty.call(MEETINGS_CS_MONTH_GENITIVE_TO_INDEX, monthToken)) {
    throw new Error('unrecognized Czech month name "' + match[2] + '"');
  }

  return { year: Number(match[3]), month: MEETINGS_CS_MONTH_GENITIVE_TO_INDEX[monthToken], day: Number(match[1]) };
}

/**
 * meetingsParseCsTimeText — the Czech pack's `parseTimeText`. Searches
 * `value` UNANCHORED for an `H:MM` 24-hour time (e.g. "13:30, délka 30
 * minut" or plain "13:30"), then separately searches for the literal
 * duration phrase `délka <N> minut` (case-insensitive) anywhere in the
 * same value. Returns `{ hour, minute, durationMinutes }`, with
 * `durationMinutes` `null` when no duration phrase is present — the
 * fallback to a configured default is the CALLING pipeline's job, not this
 * parser's (D-08). Throws a controlled Error (message only, wrapped by the
 * caller) when no recognizable `H:MM` time is found at all. Range-guarding
 * hour/minute is deliberately NOT done here — see parseTeamioMeetingText's
 * own top-level range guard, shared across every language pack so both
 * carry the identical "Hour out of range" / "Minute out of range" wording.
 * Pure, no GAS globals.
 */
function meetingsParseCsTimeText(value) {
  const text = String(value || '');
  const timeMatch = /(\d{1,2}):(\d{2})/.exec(text);

  if (!timeMatch) {
    throw new Error('no recognizable time found in "' + text + '"');
  }

  const durationMatch = /délka\s+(\d+)\s*minut/i.exec(text);

  return {
    hour: Number(timeMatch[1]),
    minute: Number(timeMatch[2]),
    durationMinutes: durationMatch ? Number(durationMatch[1]) : null,
  };
}

const MEETINGS_CS_LANGUAGE_PACK = {
  whenLabels: ['Kdy'],
  timeLabels: ['Čas'],
  whereLabels: ['Kde'],
  parseDateText: meetingsParseCsDateText,
  parseTimeText: meetingsParseCsTimeText,
};

/**
 * MEETINGS_EN_MONTH_NAMES — lowercase full English month name -> zero-
 * indexed month. NAMESPACED, distinct from
 * TICKETMASTER_CZ_MONTH_NAMES (src/07-action-ticketing-portals.js) per this
 * file's self-containment convention — never imported across files.
 */
const MEETINGS_EN_MONTH_NAMES = {
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
 * meetingsParseEnDateText — the English pack's `parseDateText`, structurally
 * the mirror of meetingsParseCsDateText: searches `value` UNANCHORED for
 * `D <full month name> YYYY` (e.g. "24 August 2026"), which naturally skips
 * a leading weekday name ("Monday, 24 August 2026") the same way. Returns
 * `{ year, month, day }` (month zero-indexed). Throws a controlled Error
 * (message only, wrapped by the caller) on no match, or an unrecognized
 * month name (naming the bad token). Pure, no GAS globals.
 */
function meetingsParseEnDateText(value) {
  const text = String(value || '');
  const match = /(\d{1,2})\s+([\p{L}]+)\s+(\d{4})/u.exec(text);

  if (!match) {
    throw new Error('no recognizable date found in "' + text + '"');
  }

  const monthToken = match[2].toLowerCase();
  if (!Object.prototype.hasOwnProperty.call(MEETINGS_EN_MONTH_NAMES, monthToken)) {
    throw new Error('unrecognized English month name "' + match[2] + '"');
  }

  return { year: Number(match[3]), month: MEETINGS_EN_MONTH_NAMES[monthToken], day: Number(match[1]) };
}

/**
 * meetingsParseEnTimeText — the English pack's `parseTimeText`, structurally
 * the mirror of meetingsParseCsTimeText: an `H:MM` 24-hour time plus an
 * optional `duration <N> minute(s)` phrase (case-insensitive,
 * singular/plural tolerant). Returns `{ hour, minute, durationMinutes }`
 * (`durationMinutes` null when absent). Throws a controlled Error (message
 * only, wrapped by the caller) on no recognizable time. Range-guarding is
 * the caller's job, same as the Czech pack. Pure, no GAS globals.
 */
function meetingsParseEnTimeText(value) {
  const text = String(value || '');
  const timeMatch = /(\d{1,2}):(\d{2})/.exec(text);

  if (!timeMatch) {
    throw new Error('no recognizable time found in "' + text + '"');
  }

  const durationMatch = /duration\s+(\d+)\s*minutes?/i.exec(text);

  return {
    hour: Number(timeMatch[1]),
    minute: Number(timeMatch[2]),
    durationMinutes: durationMatch ? Number(durationMatch[1]) : null,
  };
}

/**
 * MEETINGS_EN_LANGUAGE_PACK
 *
 * VERIFICATION STATUS: this ENTIRE language pack — every label spelling
 * (`When` / `Time` / `Where`), the date/time shape ("D Month YYYY" /
 * "H:MM" / "duration N minutes"), and the month-name table above — is
 * REQUIRED by a locked user decision (D-05) but has NOT YET been observed
 * in a real meeting-invitation email. It is implemented as a standard,
 * reasonable English equivalent of the empirically-verified Czech pack, on
 * the same explicit "flag rather than silently present an unverified guess
 * as real data" basis CZECH_MONTH_GENITIVE_TO_INDEX already carries in
 * src/06-lang-cs.js (and MEETINGS_CS_MONTH_GENITIVE_TO_INDEX mirrors above)
 * for its own 11 unobserved genitive forms — correct this pack against real
 * data the first time an English meeting invitation actually arrives.
 */
const MEETINGS_EN_LANGUAGE_PACK = {
  whenLabels: ['When'],
  timeLabels: ['Time'],
  whereLabels: ['Where'],
  parseDateText: meetingsParseEnDateText,
  parseTimeText: meetingsParseEnTimeText,
};

// MEETINGS_LANGUAGE_PACKS — registered cs THEN en (insertion order,
// preserved by JS for string keys), so parseTeamioMeetingText's language-
// pack selection tries Czech first, matching the real, empirically-verified
// system this feature was built from.
const MEETINGS_LANGUAGE_PACKS = {
  cs: MEETINGS_CS_LANGUAGE_PACK,
  en: MEETINGS_EN_LANGUAGE_PACK,
};

// MEETINGS_FALLBACK_SUMMARY — used as the created event's summary when a
// matched message's subject is empty/missing, so this action never creates
// an event with a blank title.
const MEETINGS_FALLBACK_SUMMARY = 'Meeting invitation';

/**
 * meetingsDescribeSearchedLabelPairs — the `when:time` label-pair descriptor
 * for every registered language pack, in registration order (e.g.
 * `['Kdy:Čas', 'When:Time']`). Extracted so parseTeamioMeetingText's
 * no-pack-match error message can still name exactly what was searched, now
 * that the search itself lives in meetingsSelectLanguagePack below. Pure.
 */
function meetingsDescribeSearchedLabelPairs() {
  const packKeys = Object.keys(MEETINGS_LANGUAGE_PACKS);
  const pairs = [];

  for (let i = 0; i < packKeys.length; i++) {
    const pack = MEETINGS_LANGUAGE_PACKS[packKeys[i]];
    pairs.push(pack.whenLabels.join('/') + ':' + pack.timeLabels.join('/'));
  }

  return pairs;
}

/**
 * meetingsSelectLanguagePack — (debug teamio-non-invite-error, NEW) the
 * SINGLE SHARED SOURCE OF TRUTH for "does this body carry meeting-invitation
 * structure, and if so under which language pack?". Extracted verbatim from
 * what used to be parseTeamioMeetingText's own inline pack-selection loop, so
 * that the same decision can now ALSO be asked BEFORE committing to a parse
 * (see teamioTextLooksLikeMeetingInvitation and findMeetingProcessingJobs).
 *
 * WHY THIS EXISTS — the bug it fixes: before this, the only invitation-shaped-
 * content check in the whole action lived DOWNSTREAM, inside the parser, and
 * its only way of saying "this is not an invitation" was to THROW. Meanwhile
 * findMeetingProcessingJobs admitted messages on ENVELOPE EVIDENCE ALONE
 * (sender domain + registered parser + no `.ics`), never reading the body.
 * Since `*.teamio.com` is a multi-purpose ATS domain that also sends
 * rejections and status updates, every such email was admitted past the gate
 * and then necessarily threw, surfacing to the owner as an action-failure
 * notification (real incident, 2026-08-24: a candidate-rejection notice).
 *
 * THE FIDELITY PROPERTY that makes the new gate safe: because the detector and
 * the parser now consult THIS ONE function, `meetingsSelectLanguagePack(text)
 * === null` holds for EXACTLY the bodies on which the parser would have raised
 * its no-pack-match error — never a wider or narrower set. The two can never
 * drift apart, which a second, independently-written "looks like an invitation"
 * heuristic would eventually have done. Locked by the FIDELITY PROPERTY test in
 * test/meetings.test.js.
 *
 * Selection rule (unchanged): the FIRST pack in MEETINGS_LANGUAGE_PACKS order
 * (cs, then en) for which BOTH a when-label line AND a time-label line are
 * found. A where-label line is looked up for the winning pack only, and is
 * OPTIONAL — a meeting with no stated place is a real meeting. Labels from two
 * different packs never combine to satisfy one pack.
 *
 * Input handling mirrors the parser's: `String(bodyText || '')`, U+00A0
 * normalized to a regular space, split on `\r\n` / bare `\r` / bare `\n`, each
 * line trimmed. Returns `{ pack, whenMatch, timeMatch, whereMatch }` (the
 * matches being meetingsFindLabelLine's own `{ label, value, lineText }`
 * shape; `whereMatch` may be `null`), or `null` when no pack matched. Pure, no
 * GAS globals. NEVER THROWS — a null/undefined/empty/non-string input simply
 * yields `null`, which is what lets it be called from the job-resolution gate.
 */
function meetingsSelectLanguagePack(bodyText) {
  const normalizedText = String(bodyText || '').replace(/\u00A0/g, ' ');
  const lines = normalizedText.split(/\r\n|\r|\n/).map(function (line) {
    return line.trim();
  });

  const packKeys = Object.keys(MEETINGS_LANGUAGE_PACKS);
  for (let i = 0; i < packKeys.length; i++) {
    const pack = MEETINGS_LANGUAGE_PACKS[packKeys[i]];

    const foundWhen = meetingsFindLabelLine(lines, pack.whenLabels);
    const foundTime = meetingsFindLabelLine(lines, pack.timeLabels);

    if (foundWhen && foundTime) {
      return {
        pack: pack,
        whenMatch: foundWhen,
        timeMatch: foundTime,
        whereMatch: meetingsFindLabelLine(lines, pack.whereLabels),
      };
    }
  }

  return null;
}

/**
 * teamioTextLooksLikeMeetingInvitation — (debug teamio-non-invite-error, NEW)
 * Teamio's INVITATION DETECTOR: the cheap, never-throwing predicate
 * findMeetingProcessingJobs consults to decide whether a `*.teamio.com`
 * message is a meeting invitation at all, BEFORE any job is created for it.
 *
 * Deliberately delegates the entire decision to meetingsSelectLanguagePack —
 * see that function's own FIDELITY PROPERTY paragraph for why this must not be
 * re-implemented as an independent heuristic.
 *
 * DETECTS ON THE PLAIN-TEXT BODY ONLY, which is sound rather than a shortcut:
 * parseTeamioMeetingText REQUIRES the when/time pair in the plain-text body
 * (the HTML body is consulted only for the where VALUE, via
 * meetingsExtractHtmlLabelValue), so a message whose plain text lacks the pair
 * could not be parsed even if its HTML carried it.
 *
 * ACCEPTED LIMITATION, recorded deliberately: a GENUINE invitation written in
 * an unsupported language is indistinguishable from a non-invitation by this
 * predicate, and is now silently skipped rather than reported. The cost
 * asymmetry decides it — skipping costs one missed automation while the owner
 * still receives the email and can act manually, whereas throwing on routine
 * ATS mail floods the owner with failure notifications and thereby trains them
 * to ignore the very channel that reports real failures. Adding a language pack
 * to MEETINGS_LANGUAGE_PACKS extends detection and parsing together, since both
 * read the same registry.
 *
 * Pure, no GAS globals. Never throws.
 */
function teamioTextLooksLikeMeetingInvitation(bodyText) {
  return meetingsSelectLanguagePack(bodyText) !== null;
}

/**
 * parseTeamioMeetingText — the Teamio-specific (`*.teamio.com`) meeting-body
 * parser (D-05, D-08, D-09). Real reference body this parser was built
 * from (Czech, opaque tracking tokens and the real Teams meeting ID
 * fictionalized per this codebase's established fixture convention — see
 * test/meetings.test.js's own REAL_TEAMIO_CS_BODY_TEXT /
 * REAL_TEAMIO_CS_HTML_TEXT fixtures for the exact substitution):
 *
 *   Kdy: Pondělí, 24. srpna 2026
 *   Čas: 13:30, délka 30 minut
 *   Kde: https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345
 *
 *   Přejít na potvrzení pohovoru: https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890
 *
 * ROUND 2 CORRECTION (live-test-driven, quick-260820-g4r — see
 * meetingsExtractHtmlLabelValue's own class-level doc for the full incident
 * writeup): the ORIGINAL (round 1) design read the "Kde:" value straight
 * from the plain-text body above, which the owner's live pass proved wrong
 * — that value is a `track.teamio.com` click-tracking redirect Teamio's ESP
 * generates when rendering the plain-text alternative from the HTML, NEVER
 * the real, human-readable meeting URL. The clean URL exists ONLY in the
 * HTML body's own "Kde:" table cell. This function now accepts an OPTIONAL
 * third `htmlBodyText` argument and prefers the HTML-sourced value for
 * location when it can be extracted (see step 6 below); omitting it (or a
 * pack/HTML shape that yields nothing) falls back to the plain-text value
 * exactly as before — never a throw either way.
 *
 * EXTRACTION IS LABEL-ANCHORED, not line-position-based (see
 * meetingsFindLabelLine's own doc): this codebase has a documented incident
 * (parseEnigooTicketText's own class-level "round 5" rewrite,
 * src/07-action-ticketing-portals.js) establishing that a fixed-line-
 * position model is fragile against real-world body rendering variance —
 * label-anchored extraction does not depend on which line a field happens
 * to land on.
 *
 * Algorithm, in order:
 *   1. `rawText` is `String(bodyText || '')`; every U+00A0 (non-breaking
 *      space) is replaced with a regular space in a WORKING copy only —
 *      every thrown message below still reports the ORIGINAL rawText (same
 *      convention as parseTicketmasterCzTicketText's own NBSP handling).
 *   2. Split on `\r\n`, a bare `\r`, or a bare `\n` (separator-agnostic),
 *      trim each line.
 *   3. Select the language pack: the FIRST pack in MEETINGS_LANGUAGE_PACKS
 *      order (cs, then en) for which BOTH a when-label line and a
 *      time-label line are found. No pack matching both is a controlled
 *      throw naming the label pairs searched, ending with the full
 *      rawText.
 *   4. Parse the date from the when-line's value via the selected pack's
 *      `parseDateText`, and the time+duration from the time-line's value
 *      via `parseTimeText`. A failure in either is a controlled throw
 *      ending with the full rawText (an unrecognized month name names the
 *      bad value, since that is embedded in parseDateText's own thrown
 *      message).
 *   5. Range-guard hour 0-23 and minute 0-59, with the SAME
 *      "Hour out of range" / "Minute out of range" wording the ticketing
 *      parsers already use (see parseEnigooTicketText), each ending with
 *      the full rawText. Deliberately done HERE, once, rather than
 *      per-pack — so every language pack gets the identical wording without
 *      duplicating the check.
 *   6. Location (ROUND 2 CORRECTED, D-09/threat T-g4r-04 unaffected):
 *      `meetingsExtractHtmlLabelValue(htmlBodyText, selectedPack.whereLabels)`
 *      first — a real, human-readable URL when the HTML body carries the
 *      matched pack's where-label in its own table-cell shape; else the
 *      plain-text where-line's own value, trimmed; else `''` when neither
 *      source has one — a meeting with no stated place is a real meeting,
 *      not a parse failure.
 *   7. Summary: the trimmed `subject` argument, or MEETINGS_FALLBACK_SUMMARY
 *      when empty/missing.
 *   8. Description (D-09, ROUND 3 CORRECTED — bugs 3 and 4): the matched
 *      when/time/where label lines, in source order (only the lines that
 *      were actually found). The when/time lines are the raw PLAIN-TEXT
 *      lines, verbatim. The where line is RECONSTRUCTED as
 *      `<label>: <location>` using step 6's FINAL resolved `location`
 *      value (HTML-preferred) — NOT the raw plain-text line — so the same
 *      field never shows two different values in two different places on
 *      the same generated event (bug 3: the created event's `location`
 *      field and the description's own "Kde:" line must always agree).
 *      Then — when meetingsHarvestBodyLinks(rawText, location) returns
 *      anything — a blank line, the literal heading `Links:`, and one URL
 *      per line, each pair of entries separated by a BLANK line (bug 4,
 *      `links.join('\n\n')`, not `'\n'`) for readability. `location` here
 *      is step 6's FINAL resolved value (HTML-preferred), so when the HTML
 *      supplied a different, non-matching URL than every plain-text link
 *      (the real Teamio case), NEITHER the "Kde:" line's own tracked
 *      redirect NOR the confirm-CTA's tracked redirect gets excluded from
 *      the Links: block — both legitimately surface here (round 2's fix),
 *      even though the description's own "Kde:" line above now shows the
 *      clean HTML-sourced URL instead (round 3's fix) — the SAME raw
 *      plain-text "Kde:" URL simply becomes a Links: entry instead of
 *      being echoed twice.
 *
 * Returns `{ summary, location, description, year, month, day, hour,
 * minute, durationMinutes }` (month ZERO-INDEXED, matching every other
 * date-components object in this codebase; `durationMinutes` is `null`
 * when the body states none — the DEFAULT_DURATION_MINUTES fallback is
 * applied by processMeetingFromMessageBody, not here). Pure, no GAS
 * globals.
 */
function parseTeamioMeetingText(bodyText, subject, htmlBodyText) {
  const rawText = String(bodyText || '');

  // Pack selection is delegated to meetingsSelectLanguagePack (see its own
  // class-level doc): the SAME function findMeetingProcessingJobs's invitation
  // detector consults, so the gate and the parser can never disagree about
  // what counts as invitation-shaped text.
  const selection = meetingsSelectLanguagePack(rawText);

  // NOTE (debug teamio-non-invite-error): this throw is now a DEFENSIVE
  // INVARIANT, not a routine outcome. findMeetingProcessingJobs refuses to
  // create a job at all for a body this function could not select a pack for,
  // so reaching here means the gate and the parser disagreed \u2014 a real bug
  // worth surfacing loudly. Non-invitation emails from a matched sender domain
  // (the 2026-08-24 Teamio rejection-notice incident) no longer arrive here.
  if (!selection) {
    throw new Error(
      'Unrecognized meeting invitation: no registered language pack matched (searched label pairs: ' +
        meetingsDescribeSearchedLabelPairs().join(', ') +
        '). Full extracted text:\n' +
        rawText
    );
  }

  const selectedPack = selection.pack;
  const whenMatch = selection.whenMatch;
  const timeMatch = selection.timeMatch;
  const whereMatch = selection.whereMatch;

  let dateComponents;
  try {
    dateComponents = selectedPack.parseDateText(whenMatch.value);
  } catch (dateError) {
    throw new Error('Unrecognized meeting invitation date (' + dateError.message + '). Full extracted text:\n' + rawText);
  }

  let timeComponents;
  try {
    timeComponents = selectedPack.parseTimeText(timeMatch.value);
  } catch (timeError) {
    throw new Error('Unrecognized meeting invitation time (' + timeError.message + '). Full extracted text:\n' + rawText);
  }

  if (timeComponents.hour < 0 || timeComponents.hour > 23) {
    throw new Error('Hour out of range (0-23) in meeting invitation time. Full extracted text:\n' + rawText);
  }
  if (timeComponents.minute < 0 || timeComponents.minute > 59) {
    throw new Error('Minute out of range (0-59) in meeting invitation time. Full extracted text:\n' + rawText);
  }

  // ROUND 2 CORRECTED (D-09 location resolution) -- see this function's own
  // class-level "ROUND 2 CORRECTION" doc: HTML-sourced value preferred
  // (the real, human-readable URL), falling back to the plain-text
  // where-line's own value (which for a real Teamio email is an opaque
  // track.teamio.com click-tracking redirect, not what should end up on the
  // calendar event), falling back to '' when neither source has one.
  const htmlLocation = meetingsExtractHtmlLabelValue(htmlBodyText, selectedPack.whereLabels);
  const location = htmlLocation || (whereMatch ? whereMatch.value : '');
  const summary = subject && subject.trim() ? subject.trim() : MEETINGS_FALLBACK_SUMMARY;

  // ROUND 3 CORRECTED (live-test-driven bug 3): the rendered "where" label
  // line in the description must echo the SAME resolved `location` value
  // used above (HTML-preferred), not a second, independently-sourced copy
  // of the plain-text line -- otherwise the same field (e.g. "Kde:") shows
  // two different values in two different places on the same generated
  // event (the correct HTML-sourced URL as `location`, a stale
  // plain-text-sourced tracked redirect echoed into the description). The
  // when/time lines are unaffected -- only the where line has an
  // HTML-preferred alternate source at all.
  const matchedLabelLines = [whenMatch.lineText, timeMatch.lineText];
  if (whereMatch) {
    matchedLabelLines.push(whereMatch.label + ': ' + location);
  }

  // ROUND 3 CORRECTED (live-test-driven bug 4): a blank line between each
  // Links: entry, for readability -- `links.join('\n\n')`, not `'\n'`.
  const links = meetingsHarvestBodyLinks(rawText, location);
  let description = matchedLabelLines.join('\n');
  if (links.length > 0) {
    description += '\n\nLinks:\n' + links.join('\n\n');
  }

  return {
    summary: summary,
    location: location,
    description: description,
    year: dateComponents.year,
    month: dateComponents.month,
    day: dateComponents.day,
    hour: timeComponents.hour,
    minute: timeComponents.minute,
    durationMinutes: timeComponents.durationMinutes,
  };
}

// --- registry, job resolution and the no-.ics gate (D-01, D-04) ------------

/**
 * MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN — the local (single-file) registry
 * mapping a meeting system's `domainPattern` (the SAME string used in its
 * MEETINGS_ACTION_CONFIG.meetingSystems entry) to its email-body parser
 * function — the mechanism that lets every system's parser live in this
 * one file (D-01, D-11) while still cleanly routing a resolved
 * MEETINGS_ACTION_CONFIG entry to the right parser. Consequence, deliberate
 * and documented (the identical, already-accepted property of
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL): editing a shipped entry's
 * `domainPattern` to a value with no registered parser here resolves to a
 * system (resolveMeetingSystem still finds it) but produces NO processing
 * job (findMeetingProcessingJobs below requires a registered parser too).
 * Adding a brand-new system needs a parser function added here, in
 * addition to the config entry AND a detector in the sibling
 * MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN registry below — the same
 * fail-closed "no registration, no job" rule applies to both registries.
 */
const MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN = {
  '*.teamio.com': parseTeamioMeetingText,
};

/**
 * MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN — (debug
 * teamio-non-invite-error, NEW) the companion registry to
 * MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN, keyed by the SAME `domainPattern`
 * string, mapping a system to its INVITATION DETECTOR: a pure, never-throwing
 * `(plainBodyText) => boolean` predicate answering "is this message a meeting
 * invitation at all?" WITHOUT attempting a parse.
 *
 * WHY A SECOND REGISTRY EXISTS. Matching a sender is not the same claim as
 * "this email is a meeting invitation", and before this fix the action
 * conflated the two: findMeetingProcessingJobs admitted a message on envelope
 * evidence alone, and the parser — the only thing that ever looked at the body
 * — could say "not an invitation" only by THROWING. That is fine for a system
 * whose sender address is transactional-only (the precedent this action
 * inherited from TICKETING_PORTALS_ACTION, which matches EXACT addresses), but
 * this action is the first to match a whole DOMAIN (D-02), and `*.teamio.com`
 * is an ATS platform that sends rejections and status updates from the same
 * domain as its invitations. The envelope simply cannot carry the distinction,
 * so the content must — before any job is created.
 *
 * FAIL-CLOSED BY DESIGN: findMeetingProcessingJobs requires a REGISTERED
 * detector, exactly as it already requires a registered parser. A system with
 * a parser but no detector produces NO jobs rather than falling back to the old
 * envelope-only behaviour. Defaulting the other way ("no detector means admit
 * everything") is precisely the shape of the bug this fixes, and would silently
 * reintroduce it for the next system added. The cost is that adding a system
 * now needs a config entry + a parser + a detector; the registry-key equality
 * test in test/meetings.test.js enforces that the two registries stay in step.
 */
const MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN = {
  '*.teamio.com': teamioTextLooksLikeMeetingInvitation,
};

/**
 * meetingsIsIcsAttachment — true when `attachment`'s name ends in `.ics`
 * (case-insensitive) or its content-type is `text/calendar`. Reproduces
 * isIcsAttachment's (src/05-action-ics-import.js) exact rule LOCALLY,
 * per-action self-containment convention (D-04) — this file must never
 * cross-require the ICS action's own detection helper, since the whole
 * point of this gate living here (rather than as an ICS_ACTION_CONFIG
 * exclusion) is that this action's no-.ics scope is entirely self-owned.
 */
function meetingsIsIcsAttachment(attachment) {
  const name = (attachment.getName() || '').toLowerCase();
  const contentType = attachment.getContentType() || '';

  return name.slice(-4) === '.ics' || contentType === 'text/calendar';
}

/**
 * meetingsMessageHasIcsAttachment — true when any of `message`'s
 * attachments qualifies via meetingsIsIcsAttachment.
 */
function meetingsMessageHasIcsAttachment(message) {
  const attachments = message.getAttachments() || [];

  for (let i = 0; i < attachments.length; i++) {
    if (meetingsIsIcsAttachment(attachments[i])) {
      return true;
    }
  }

  return false;
}

/**
 * findMeetingProcessingJobs — the pure, TESTABLE extraction of `run`'s
 * per-message orchestration decision, mirroring
 * resolveTicketProcessingJobs's own shape. Given `messages` (an array of
 * message-like objects exposing `getFrom()`/`getAttachments()` — real
 * GmailMessage objects in production, plain duck-typed fakes in tests) and
 * `meetingSystems` (the MEETINGS_ACTION_CONFIG.meetingSystems array),
 * returns one job `{ message, system, parser }` per message that (a)
 * resolves to a configured system (resolveMeetingSystem), (b) has a
 * registered parser for that system's `domainPattern`
 * (MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN), (c) carries NO `.ics`
 * attachment (D-04, meetingsMessageHasIcsAttachment), AND (d) — debug
 * teamio-non-invite-error, NEW — has a registered INVITATION DETECTOR for that
 * `domainPattern` (MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN) which
 * returns true for the message's PLAIN-TEXT BODY. Every other message
 * contributes nothing. Pure in the sense that matters here — it touches no
 * GAS global, only methods on the passed-in message objects, so it is
 * fully unit-testable under Node with fake message/attachment objects.
 * Never throws.
 *
 * CONDITION (d) IS THE FIX for the 2026-08-24 incident: conditions (a)-(c) are
 * all ENVELOPE evidence, and a sender domain is not proof that an email is a
 * meeting invitation. `*.teamio.com` is a multi-purpose ATS domain that also
 * sends candidate-rejection notices, which sailed through (a)-(c) and then made
 * the parser throw, reaching the owner as an action-failure notification. The
 * check belongs HERE rather than inside `run` precisely because this resolver
 * also backs MEETINGS_ACTION.appliesTo — gating only `run` would leave
 * appliesTo true, so the action would still CLAIM the thread and mark it
 * processed while doing nothing. Returning no job means the action correctly
 * does not apply at all.
 */
function findMeetingProcessingJobs(messages, meetingSystems) {
  const list = messages || [];
  const jobs = [];

  for (let i = 0; i < list.length; i++) {
    const message = list[i];
    const system = resolveMeetingSystem(message.getFrom(), meetingSystems);
    if (!system) {
      continue;
    }

    const parser = MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN[system.domainPattern];
    if (!parser) {
      continue;
    }

    if (meetingsMessageHasIcsAttachment(message)) {
      continue;
    }

    // Fail-closed: an unregistered detector yields no job, mirroring the
    // unregistered-parser rule above. See
    // MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN's own class-level doc.
    const looksLikeInvitation = MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN[system.domainPattern];
    if (!looksLikeInvitation || !looksLikeInvitation(message.getPlainBody())) {
      continue;
    }

    jobs.push({ message: message, system: system, parser: parser });
  }

  return jobs;
}

// Node/GAS environment bridge for MEETINGS_ACTION_CONFIG (defined in the
// sibling src/10-action-cfg-meetings.js — see the 260724-lqi config-split
// refactor for the full load-order/getter rationale this mirrors). Under
// GAS's shared global scope this is ALREADY visible here by bare name — no
// action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare `MEETINGS_ACTION_CONFIG` reference inside MEETINGS_ACTION's
// `config` getter below would otherwise throw ReferenceError. Same
// `globalThis` bridge technique already established by every other action
// file's own equivalent bridge.
if (typeof module !== 'undefined' && module.exports) {
  globalThis.MEETINGS_ACTION_CONFIG = require('./10-action-cfg-meetings.js').MEETINGS_ACTION_CONFIG;
}

// --- the GAS pipeline (D-03, D-07), GAS-only, not unit-tested --------------

/**
 * findMeetingEventByIdentifier — the DEDUP SAFETY NET's lookup, mirroring
 * findTicketEventByIdentifier's (src/07-action-ticketing-portals.js) exact
 * query shape: `Calendar.Events.list(calendarId, { privateExtendedProperty:
 * 'meetingIdentifier=' + meetingIdentifier, singleEvents: true })`.
 * Deliberately NOT paginated or time-windowed — a `privateExtendedProperty`
 * filter against a near-certainly-unique per-meeting identifier is already
 * an exact match expected to return 0 or 1 events. Returns the first
 * matching event, or `null`. GAS-only (Calendar global) — not unit-tested,
 * proven only by the live checkpoint, same category as every other
 * GAS-only function in this codebase.
 */
function findMeetingEventByIdentifier(meetingIdentifier, calendarId) {
  const response = Calendar.Events.list(calendarId, {
    privateExtendedProperty: 'meetingIdentifier=' + meetingIdentifier,
    singleEvents: true,
  });
  const items = (response && response.items) || [];
  return items.length > 0 ? items[0] : null;
}

/**
 * isDuplicateMeetingInvite — the DEDUP SAFETY NET's decision. Returns
 * `false` on a falsy `meetingIdentifier` (buildMeetingIdentifier never
 * actually produces one, but this mirrors isDuplicateTicketPurchase's own
 * defensive shape), otherwise `true` (plus a `Meetings:`-prefixed
 * already-exists log line) when an event already carries that exact tag on
 * `calendarId`. GAS-only (calls findMeetingEventByIdentifier) — not
 * unit-tested, proven only by the live checkpoint.
 */
function isDuplicateMeetingInvite(meetingIdentifier, calendarId) {
  if (!meetingIdentifier) {
    return false;
  }

  const existingEvent = findMeetingEventByIdentifier(meetingIdentifier, calendarId);
  if (existingEvent) {
    console.log(
      'Meetings: event for meeting identifier ' + meetingIdentifier + ' already exists, skipping (safety-net, not a duplicate path).'
    );
    return true;
  }

  return false;
}

/**
 * createMeetingCalendarEvent — builds and inserts the Calendar event
 * resource from `parsedMeeting` (`{ summary, location, description, year,
 * month, day, hour, minute }`), tagged at creation with
 * `extendedProperties.private.meetingIdentifier` (D-07). TIMEZONE resolved
 * LIVE from the target calendar (`CalendarApp.getCalendarById(calendarId)
 * .getTimeZone()`), never a hardcoded assumption. GAS-only
 * (CalendarApp/Calendar globals) — not unit-tested, proven only by the live
 * checkpoint; the pure logic it depends on
 * (meetingsAddMinutesToWallClockComponents,
 * meetingsFormatWallClockComponentsIso) IS fully unit-tested.
 */
function createMeetingCalendarEvent(parsedMeeting, calendarId, durationMinutes, meetingIdentifier) {
  const timeZone = CalendarApp.getCalendarById(calendarId).getTimeZone();

  const startComponents = {
    year: parsedMeeting.year,
    month: parsedMeeting.month,
    day: parsedMeeting.day,
    hour: parsedMeeting.hour,
    minute: parsedMeeting.minute,
  };
  const endComponents = meetingsAddMinutesToWallClockComponents(startComponents, durationMinutes);

  const resource = {
    summary: parsedMeeting.summary,
    location: parsedMeeting.location,
    description: parsedMeeting.description,
    start: { dateTime: meetingsFormatWallClockComponentsIso(startComponents), timeZone: timeZone },
    end: { dateTime: meetingsFormatWallClockComponentsIso(endComponents), timeZone: timeZone },
    extendedProperties: { private: { meetingIdentifier: meetingIdentifier } },
  };

  Calendar.Events.insert(resource, calendarId);
}

/**
 * processMeetingFromMessageBody — the per-job pipeline (D-03, D-07, D-08):
 * reads `message.getPlainBody()`/`getSubject()`/`getBody()` (the last one
 * ROUND 2 NEW — see parseTeamioMeetingText's own class-level "ROUND 2
 * CORRECTION" doc: the real, human-readable meeting URL for a Teamio
 * invitation exists ONLY in the HTML body, never the plain-text one), parses
 * via `parser(bodyText, subject, htmlBodyText)`, resolves the target
 * calendar (resolveMeetingsCalendarId(system, CONFIG.calendarId)), builds
 * the dedup identifier from the sender + parsed summary + parsed start
 * wall-clock, returns early (no writes at all) when isDuplicateMeetingInvite
 * says so, otherwise resolves the duration as `parsedMeeting.durationMinutes
 * || MEETINGS_ACTION.config.defaultDurationMinutes` (D-08 — the body's own
 * stated duration wins when present) and creates the event, logging a
 * `Meetings:`-prefixed line naming the summary and the resolved
 * identifier. GAS-only (GmailMessage/CalendarApp/Calendar globals) — not
 * unit-tested, proven only by the live checkpoint; the pure logic it
 * depends on (the parser, buildMeetingIdentifier) IS fully unit-tested.
 */
function processMeetingFromMessageBody(message, system, parser) {
  const bodyText = message.getPlainBody();
  const subject = message.getSubject();
  const htmlBodyText = message.getBody();
  const parsedMeeting = parser(bodyText, subject, htmlBodyText);

  const calendarId = resolveMeetingsCalendarId(system, CONFIG.calendarId);
  const senderEmail = meetingsExtractEmailAddress(message.getFrom());
  const startComponents = {
    year: parsedMeeting.year,
    month: parsedMeeting.month,
    day: parsedMeeting.day,
    hour: parsedMeeting.hour,
    minute: parsedMeeting.minute,
  };
  const meetingIdentifier = buildMeetingIdentifier(senderEmail, parsedMeeting.summary, startComponents);

  if (isDuplicateMeetingInvite(meetingIdentifier, calendarId)) {
    return;
  }

  const durationMinutes = parsedMeeting.durationMinutes || MEETINGS_ACTION.config.defaultDurationMinutes;

  console.log(
    'Meetings: creating calendar event for "' + parsedMeeting.summary + '" (meetingIdentifier=' + meetingIdentifier + ') on calendar ' +
      calendarId +
      '.'
  );
  createMeetingCalendarEvent(parsedMeeting, calendarId, durationMinutes, meetingIdentifier);
}

// --- the descriptor (D-01) --------------------------------------------------

/**
 * MEETINGS_ACTION — the meetings action descriptor, built on the same shape
 * as every other action (see TICKETING_PORTALS_ACTION,
 * src/07-action-ticketing-portals.js).
 */
const MEETINGS_ACTION = {
  name: 'meetings',

  // GETTER, not a plain literal property — see this file's class-level
  // JSDoc and the sibling config file's own class-level JSDoc. Not
  // evaluated at object-construction time, only when something reads
  // `.config`, which happens lazily inside function bodies (dispatchActions,
  // notifyOwnerOfFailure, processMeetingFromMessageBody) long after every
  // project file has loaded.
  get config() {
    return MEETINGS_ACTION_CONFIG;
  },

  /**
   * appliesTo — returns a LITERAL boolean (dispatchActions only skips on a
   * strict `=== false`). True when findMeetingProcessingJobs finds at
   * least one matching, no-.ics job on the thread.
   */
  appliesTo: function (thread) {
    return findMeetingProcessingJobs(thread.getMessages(), MEETINGS_ACTION.config.meetingSystems).length > 0;
  },

  /**
   * run — resolves the job list via the same pure resolver appliesTo uses,
   * then runs processMeetingFromMessageBody once per job.
   */
  run: function (thread) {
    const jobs = findMeetingProcessingJobs(thread.getMessages(), MEETINGS_ACTION.config.meetingSystems);

    jobs.forEach(function (job) {
      processMeetingFromMessageBody(job.message, job.system, job.parser);
    });
  },
};

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Exports every pure helper/parser/registry (meetingsExtractEmailAddress,
// meetingsExtractSenderDomain, meetingsDomainMatchesPattern,
// resolveMeetingSystem, resolveMeetingsCalendarId, meetingsZeroPad,
// meetingsAddMinutesToWallClockComponents,
// meetingsFormatWallClockComponentsIso, meetingsHarvestBodyLinks,
// meetingsExtractHtmlLabelValue (ROUND 2 NEW — the HTML-sourced label-value
// extractor bug 1's fix depends on), buildMeetingIdentifier, the two
// language packs and their month tables,
// meetingsDescribeSearchedLabelPairs / meetingsSelectLanguagePack /
// teamioTextLooksLikeMeetingInvitation /
// MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN (debug
// teamio-non-invite-error NEW — the shared pack-selection helper, the
// invitation detector built on it, and the detector registry
// findMeetingProcessingJobs's content gate consults),
// parseTeamioMeetingText, MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN,
// meetingsIsIcsAttachment, meetingsMessageHasIcsAttachment,
// findMeetingProcessingJobs, MEETINGS_FALLBACK_SUMMARY,
// MEETINGS_MAX_DESCRIPTION_LINKS) and MEETINGS_ACTION (action registry).
// Deliberately exports NO bare `addMinutesToWallClockComponents` or
// `formatWallClockComponentsIso` key (D-10 collision guard — see this
// file's class-level "GLOBALLY-UNIQUE NAMING WARNING"; only the namespaced
// `meetings*` versions are exported).
// findMeetingEventByIdentifier/createMeetingCalendarEvent/
// processMeetingFromMessageBody remain genuinely GAS-only (reference
// CalendarApp/Calendar/GmailMessage globals directly) and are NOT
// exported — they are never invoked under Node.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    meetingsExtractEmailAddress: meetingsExtractEmailAddress,
    meetingsExtractSenderDomain: meetingsExtractSenderDomain,
    meetingsDomainMatchesPattern: meetingsDomainMatchesPattern,
    resolveMeetingSystem: resolveMeetingSystem,
    resolveMeetingsCalendarId: resolveMeetingsCalendarId,
    meetingsZeroPad: meetingsZeroPad,
    meetingsAddMinutesToWallClockComponents: meetingsAddMinutesToWallClockComponents,
    meetingsFormatWallClockComponentsIso: meetingsFormatWallClockComponentsIso,
    meetingsHarvestBodyLinks: meetingsHarvestBodyLinks,
    meetingsExtractHtmlLabelValue: meetingsExtractHtmlLabelValue,
    buildMeetingIdentifier: buildMeetingIdentifier,
    MEETINGS_CS_MONTH_GENITIVE_TO_INDEX: MEETINGS_CS_MONTH_GENITIVE_TO_INDEX,
    MEETINGS_EN_MONTH_NAMES: MEETINGS_EN_MONTH_NAMES,
    MEETINGS_CS_LANGUAGE_PACK: MEETINGS_CS_LANGUAGE_PACK,
    MEETINGS_EN_LANGUAGE_PACK: MEETINGS_EN_LANGUAGE_PACK,
    MEETINGS_LANGUAGE_PACKS: MEETINGS_LANGUAGE_PACKS,
    MEETINGS_FALLBACK_SUMMARY: MEETINGS_FALLBACK_SUMMARY,
    MEETINGS_MAX_DESCRIPTION_LINKS: MEETINGS_MAX_DESCRIPTION_LINKS,
    meetingsDescribeSearchedLabelPairs: meetingsDescribeSearchedLabelPairs,
    meetingsSelectLanguagePack: meetingsSelectLanguagePack,
    teamioTextLooksLikeMeetingInvitation: teamioTextLooksLikeMeetingInvitation,
    parseTeamioMeetingText: parseTeamioMeetingText,
    MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN: MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN,
    MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN: MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN,
    meetingsIsIcsAttachment: meetingsIsIcsAttachment,
    meetingsMessageHasIcsAttachment: meetingsMessageHasIcsAttachment,
    findMeetingProcessingJobs: findMeetingProcessingJobs,
    MEETINGS_ACTION: MEETINGS_ACTION,
  };
}

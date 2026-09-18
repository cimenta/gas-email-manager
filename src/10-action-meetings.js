/**
 * MEETINGS_ACTION — recognizes general meeting-invitation emails whose
 * date/time/location live entirely in the plain-text BODY
 * (`message.getPlainBody()`), never a structured `.ics` attachment, and
 * creates ONE calendar event per matching message.
 *
 * NO-.ICS SCOPE: this action never claims a message that carries a `.ics`
 * attachment — ICS_CALENDAR_ACTION (src/05-action-ics-import.js) already
 * owns those. The gate lives entirely in this file's own
 * meetingsMessageHasIcsAttachment / findMeetingProcessingJobs.
 *
 * PLUGGABLE "SYSTEMS" ARCHITECTURE: ONE action file plus one sibling config
 * file, NOT one file per system. Every supported meeting-invitation
 * "system" (Teamio ships first) is a config entry
 * (`{ domainPattern, calendarId }`) plus a parser function registered in
 * MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN plus a detector registered in
 * MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN, all three keyed by the
 * SAME `domainPattern` string — adding a further system needs no change to
 * the matching logic itself. Both language packs (Czech and English) also
 * live in this one file — no `10-lang-*.js` files.
 *
 * DETECTION BY SENDER DOMAIN, NOT EXACT ADDRESS. A per-system
 * `domainPattern` supports a leading `*.` subdomain wildcard:
 * `*.teamio.com` matches `recruit.teamio.com` (a real subdomain) AND the
 * bare apex `teamio.com`, but matches only on a REAL DOT BOUNDARY, so
 * `notteamio.com` (a different domain that merely ends with the same
 * letters) and `teamio.com.evil.net` (a longer domain that merely starts
 * with the pattern's suffix) do NOT match. A pattern with no `*.` prefix is
 * exact, case-insensitive equality only. See meetingsDomainMatchesPattern
 * below for the implementation, and this project's ICS action for the
 * established "From-header match is a non-cryptographic convenience
 * filter, not a security boundary" precedent this action's own domain
 * match equally relies on.
 *
 * DEDUP: free text carries no `iCalUID`, so the created event is tagged
 * with `extendedProperties.private.meetingIdentifier`, and before creating,
 * the target calendar is queried for an event already carrying that exact
 * tag (find-before-create, see findMeetingEventByIdentifier /
 * isDuplicateMeetingInvite). The identifier is built by a PURE,
 * deterministic function (buildMeetingIdentifier) from sender + normalized
 * subject + parsed start wall-clock — ALWAYS produced, never null, so every
 * meeting processed by this action gets the dedup safety net.
 *
 * TIMEZONE: wall-clock digits parsed from the body are resolved against the
 * TARGET CALENDAR'S LIVE timezone (`CalendarApp.getCalendarById(calendarId)
 * .getTimeZone()`), never a hardcoded guess.
 *
 * GLOBALLY-UNIQUE NAMING: Apps Script concatenates every project file into
 * ONE shared global scope, and `07-action-ticketing-portals.js` — a file
 * that loads BEFORE this one alphabetically ("07-" sorts before "10-") —
 * already defines UN-namespaced `addMinutesToWallClockComponents`,
 * `formatWallClockComponentsIso`, `zeroPadTicketComponent` and
 * `DEFAULT_EVENT_DURATION_MINUTES`. A same-named definition in THIS file
 * would silently take over every call site that resolves those bare names
 * (since this file loads AFTER, its own top-level declarations would win
 * the last-one-wins collision). Every helper this file introduces is
 * therefore `meetings`/`MEETINGS`-prefixed, including its OWN local copies
 * of the wall-clock helpers (meetingsAddMinutesToWallClockComponents /
 * meetingsFormatWallClockComponentsIso / meetingsZeroPad) — never imported
 * or reused across files.
 */

// --- pure helpers (D-02, D-10) ----------------------------------------------

/**
 * meetingsExtractEmailAddress — extracts the bare, trimmed, lowercased
 * email address from a Gmail "From" header value, or from a bare address
 * with no display name. Pure, no GAS globals. Never throws: a
 * null/undefined/empty input returns ''.
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
 * meetingsDomainMatchesPattern — the core matching rule. A `domainPattern`
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

// MEETINGS_SENDER_ATTRIBUTION_HEADING — the language-neutral literal
// heading meetingsFormatSenderAttribution renders, consistent with the
// English `Links:` heading the same description already emits into an
// otherwise-Czech body.
const MEETINGS_SENDER_ATTRIBUTION_HEADING = 'From:';

/**
 * meetingsExtractSenderDisplayName — returns the SANITIZED display-name
 * portion of a Gmail `From` header, or the empty string. The text
 * preceding the first `<` is the display name; a header with no `<` at all
 * has no display name, so this returns '' rather than treating the whole
 * header as a name. A single leading and trailing double-quote is stripped
 * when both are present (the common `"Display Name" <addr>` shape).
 *
 * SANITIZATION IS A SECURITY CONTROL, NOT COSMETICS: the From header's
 * display name is attacker-controllable and unauthenticated, so every run
 * of CR/LF/tab collapses to a single space and every `<`/`>` character is
 * removed before the result is used anywhere — without this, a hostile
 * display name could forge an extra label line inside a Calendar event
 * description (CR/LF injection) or smuggle angle-bracket markup into it.
 * Pure, no GAS globals. Never throws: a null/undefined/empty input returns
 * ''.
 */
function meetingsExtractSenderDisplayName(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const str = String(fromHeader);
  const angleIndex = str.indexOf('<');
  if (angleIndex === -1) {
    return '';
  }

  let namePart = str.slice(0, angleIndex).trim();
  if (namePart.length >= 2 && namePart.charAt(0) === '"' && namePart.charAt(namePart.length - 1) === '"') {
    namePart = namePart.slice(1, -1);
  }

  return namePart
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[<>]/g, '')
    .trim();
}

/**
 * meetingsFormatSenderAttribution — builds the ONE-LINE sender attribution
 * rendered as the description's first line. Resolves the address via the
 * EXISTING meetingsExtractEmailAddress — deliberately NOT a second
 * extractor, so the description's address and buildMeetingIdentifier's
 * dedup address can never disagree — and the sanitized name via
 * meetingsExtractSenderDisplayName. With both present: heading, space,
 * name, space, address wrapped in angle brackets. With an address only:
 * heading, space, bare address (no angle brackets — there is no name to
 * disambiguate from). With a name only: heading, space, name. With
 * neither: the empty string (the caller treats this as "nothing to
 * attribute"). Heading is MEETINGS_SENDER_ATTRIBUTION_HEADING. Pure, no
 * GAS globals. Never throws.
 */
function meetingsFormatSenderAttribution(fromHeader) {
  const name = meetingsExtractSenderDisplayName(fromHeader);
  const address = meetingsExtractEmailAddress(fromHeader);

  if (name && address) {
    return MEETINGS_SENDER_ATTRIBUTION_HEADING + ' ' + name + ' <' + address + '>';
  }
  if (address) {
    return MEETINGS_SENDER_ATTRIBUTION_HEADING + ' ' + address;
  }
  if (name) {
    return MEETINGS_SENDER_ATTRIBUTION_HEADING + ' ' + name;
  }
  return '';
}

/**
 * meetingsApplySenderAttribution — THE GENERAL GUARANTEE: every calendar
 * event MEETINGS_ACTION creates identifies its sender, unconditionally, for
 * every registered parser present and future. Returns a NEW shallow copy of
 * `parsedMeeting` whose `description` is the attribution, ONE blank line,
 * then the original description verbatim — identifying info leads, so the
 * variable-length `Links:` block at the bottom can never push it out of
 * view. When the attribution is empty (e.g. a falsy/unparseable
 * `fromHeader`) returns a copy whose description is left byte-for-byte
 * untouched — no bare `From:` with no value is ever rendered. When the
 * original description is empty, the attribution alone is the description
 * (no trailing blank line). A falsy `parsedMeeting` is returned as-is
 * rather than throwing.
 *
 * DELIBERATELY PARSER-AGNOSTIC: this function only ever reads/writes
 * `description` on whatever shape `parsedMeeting` happens to be — it does
 * not know or care which parser produced it. That is precisely what lets
 * processMeetingFromMessageBody apply this ONE function to every registered
 * parser's output at its single choke point, with zero per-parser change.
 * Pure, no GAS globals. Never throws.
 */
function meetingsApplySenderAttribution(parsedMeeting, fromHeader) {
  if (!parsedMeeting) {
    return parsedMeeting;
  }

  const attribution = meetingsFormatSenderAttribution(fromHeader);
  if (!attribution) {
    return Object.assign({}, parsedMeeting);
  }

  const originalDescription = parsedMeeting.description || '';
  const description = originalDescription ? attribution + '\n\n' + originalDescription : attribution;

  return Object.assign({}, parsedMeeting, { description: description });
}

/**
 * resolveMeetingSystem — finds the FIRST entry in `meetingSystems` (list
 * order) whose `domainPattern` matches `fromHeader`'s sender domain (via
 * meetingsExtractSenderDomain + meetingsDomainMatchesPattern). Returns
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
 * defaultCalendarId` resolution: `CalendarApp.getCalendarById(null)` throws
 * a real error rather than returning something null-ish, so a
 * `calendarId: null` default read directly with no fallback would be a
 * live crash. Pure, no GAS globals.
 */
function resolveMeetingsCalendarId(system, defaultCalendarId) {
  return (system && system.calendarId) || defaultCalendarId;
}

/**
 * meetingsZeroPad — left-pads `value` with '0' to `length` digits. Pure, no
 * GAS globals. Namespaced per this file's naming convention.
 */
function meetingsZeroPad(value, length) {
  return String(value).padStart(length, '0');
}

/**
 * meetingsAddMinutesToWallClockComponents — namespaced per this file's
 * naming convention. Pure: adds `minutes` to a `{ year, month, day, hour,
 * minute }` wall-clock components object (month zero-indexed), returning a
 * NEW components object of the same shape, correctly handling
 * hour/day/month/year rollover, via `Date.UTC` arithmetic used purely as a
 * NEUTRAL zero-offset calculation space — never a real UTC instant, since
 * the input components carry no timezone information at all. No GAS
 * globals.
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
 * meetingsFormatWallClockComponentsIso — namespaced per this file's naming
 * convention. Formats a `{ year, month, day, hour, minute }` wall-clock
 * components object (month zero-indexed) as a zero-padded literal string
 * `'YYYY-MM-DDTHH:MM:00'` — DELIBERATELY with NO trailing `Z` and NO
 * timezone offset, meant to be paired with an explicit Calendar API
 * `timeZone` field. Pure, no GAS globals.
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
 * meetingsHarvestBodyLinks — returns every distinct `http(s)` URL found in
 * `bodyText`, in source order, excluding `excludeUrl` (the URL already used
 * as the event location, so it is not repeated inside the description's
 * own `Links:` block), capped at MEETINGS_MAX_DESCRIPTION_LINKS. A trailing
 * run of common sentence/markup punctuation (`) , . ;`) is stripped from
 * each match, since a URL is frequently followed immediately by such a
 * character in free-text prose. Pure, no GAS globals. Never throws: a
 * null/empty bodyText yields [].
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
 * buildMeetingIdentifier — the dedup key. A PURE, deterministic composite
 * of the lowercased `senderEmail`, `subject` lowercased with whitespace
 * runs collapsed to a single space and truncated to
 * MEETINGS_IDENTIFIER_MAX_SUBJECT_LENGTH, and the formatted start
 * wall-clock (via meetingsFormatWallClockComponentsIso) — ALWAYS produced,
 * never null, so every meeting this action processes gets the dedup safety
 * net. Every `=` character is stripped from the result, because the
 * identifier is interpolated into a `privateExtendedProperty:
 * 'meetingIdentifier=' + value` Calendar API query string (see
 * findMeetingEventByIdentifier below) — an unstripped `=` in an
 * attacker-controlled subject could otherwise smuggle a second `=` into
 * that query and corrupt the lookup. No `Utilities.computeDigest` is used
 * deliberately: that is a GAS global, and hashing here would make this
 * function untestable under Node for no real benefit — a deterministic
 * composite string built from already-bounded inputs is sufficient and
 * simpler. Pure, no GAS globals. Never throws.
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
 * at the START of each line, not a `\b`-bounded regex scan across the whole
 * body: several of this pack's labels (`Kdy`, `Čas`, `Kde`) begin or end
 * with a non-ASCII letter, and JS's `\b` word-boundary only recognizes
 * ASCII word characters by default, so a `\b`-based match can silently fail
 * exactly where a label starts/ends in a non-ASCII letter (see
 * src/06-action-booking-com-management.js's own "Informace o ceně" for the
 * same class of case). Anchoring on the line's own start avoids `\b`
 * entirely. Pure, no GAS globals.
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
 * meetingsExtractHtmlLabelValue — reads a label's value out of an HTML
 * body's own markup, for the ONE real case where the plain-text body's own
 * version of that value is not what should end up on the calendar event.
 *
 * Teamio's ESP rewrites every URL-shaped string into an opaque
 * `track.teamio.com` click-tracking REDIRECT when generating the
 * plain-text alternative from the HTML. The text/plain part's own "Kde:"
 * line therefore never carries the real, human-readable meeting URL at
 * all; that URL exists ONLY in the text/html part's own "Kde:" table cell,
 * as plain (non-hyperlinked) text wrapped in a `<strong>` tag — e.g.
 * `<td>Kde:</td><td><strong>https://teams.microsoft.com/meet/...</strong></td>`.
 *
 * Searches `htmlBodyText` for the FIRST occurrence of any of `labels`
 * immediately followed by `:` (same case-insensitive label-matching
 * philosophy as meetingsFindLabelLine, generalized to any of a pack's
 * label lists — e.g. `pack.whereLabels` — so this is not hardcoded to
 * "Kde" alone), then returns the text content of the NEXT `<strong>...
 * </strong>` element found after that point (any remaining inner tags
 * stripped defensively), trimmed. Returns `null` when the label is not
 * found, no `<strong>` element follows it, or the extracted text is empty
 * — NEVER throws, and never partially applies: a caller falling back to
 * the plain-text value on `null` is the correct behavior for a body that
 * was not supplied or does not carry this particular table-cell shape.
 * Pure, no GAS globals.
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
 * GENITIVE grammatical case (the form a real Teamio date uses, e.g.
 * "24. srpna 2026"), -> zero-indexed month, matching Date.UTC's
 * convention. Namespaced copy of CZECH_MONTH_GENITIVE_TO_INDEX
 * (src/06-lang-cs.js), never imported across files.
 *
 * VERIFICATION STATUS: only 'srpna' (August) is empirically verified
 * against a real Teamio email. The remaining 11 genitive forms are
 * standard, unambiguous Czech grammar and are included on that basis, but
 * have not been observed in a real invitation — fix this table if a real
 * email surfaces an unexpected month-name mismatch.
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
 * parser's. Throws a controlled Error (message only, wrapped by the
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
 * shipped deliberately, rather than omitted, but has NOT YET been observed
 * in a real meeting-invitation email. It is a standard English equivalent
 * of the empirically-verified Czech pack — correct it against real data
 * the first time an English meeting invitation actually arrives.
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
 * meetingsSelectLanguagePack — the SINGLE SHARED SOURCE OF TRUTH for "does
 * this body carry meeting-invitation structure, and if so under which
 * language pack?", consulted both by the job-admission detector and by the
 * parser.
 *
 * THE FIDELITY PROPERTY: because both consult THIS ONE function,
 * `meetingsSelectLanguagePack(text) === null` holds for EXACTLY the bodies
 * on which the parser would raise its no-pack-match error — never a wider
 * or narrower set — so the two can never drift apart, which a second,
 * independently written "looks like an invitation" heuristic eventually
 * would. Locked by the FIDELITY PROPERTY test in test/meetings.test.js.
 *
 * `*.teamio.com` is a multi-purpose ATS domain that also sends rejections
 * and status updates, so a sender-domain match alone is not evidence that
 * an email is an invitation — that is WHY this content gate exists at all.
 *
 * Selection rule: the FIRST pack in MEETINGS_LANGUAGE_PACKS order (cs, then
 * en) for which BOTH a when-label line AND a time-label line are found. A
 * where-label line is looked up for the winning pack only, and is
 * OPTIONAL — a meeting with no stated place is a real meeting. Labels from
 * two different packs never combine to satisfy one pack.
 *
 * Input handling mirrors the parser's: `String(bodyText || '')`, U+00A0
 * normalized to a regular space, split on `\r\n` / bare `\r` / bare `\n`,
 * each line trimmed. Returns `{ pack, whenMatch, timeMatch, whereMatch }`
 * (the matches being meetingsFindLabelLine's own
 * `{ label, value, lineText }` shape; `whereMatch` may be `null`), or
 * `null` when no pack matched. Pure, no GAS globals. NEVER THROWS — a
 * null/undefined/empty/non-string input simply yields `null`, which is
 * what lets it be called from the job-resolution gate.
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
 * teamioTextLooksLikeMeetingInvitation — Teamio's INVITATION DETECTOR: the
 * cheap, never-throwing predicate findMeetingProcessingJobs consults to
 * decide whether a `*.teamio.com` message is a meeting invitation at all,
 * BEFORE any job is created for it.
 *
 * Deliberately delegates the entire decision to meetingsSelectLanguagePack
 * — see that function's own FIDELITY PROPERTY paragraph for why this must
 * not be re-implemented as an independent heuristic.
 *
 * DETECTS ON THE PLAIN-TEXT BODY ONLY, which is sound rather than a
 * shortcut: parseTeamioMeetingText REQUIRES the when/time pair in the
 * plain-text body (the HTML body is consulted only for the where VALUE,
 * via meetingsExtractHtmlLabelValue), so a message whose plain text lacks
 * the pair could not be parsed even if its HTML carried it.
 *
 * ACCEPTED LIMITATION, recorded deliberately: a GENUINE invitation written
 * in an unsupported language is indistinguishable from a non-invitation by
 * this predicate, and is silently skipped rather than reported. The cost
 * asymmetry decides it — skipping costs one missed automation while the
 * email still arrives and can be acted on manually, whereas throwing on
 * routine ATS mail floods the failure-notification channel and trains its
 * reader to ignore the very channel that reports real failures. Adding a
 * language pack to MEETINGS_LANGUAGE_PACKS extends detection and parsing
 * together, since both read the same registry.
 *
 * Pure, no GAS globals. Never throws.
 */
function teamioTextLooksLikeMeetingInvitation(bodyText) {
  return meetingsSelectLanguagePack(bodyText) !== null;
}

/**
 * parseTeamioMeetingText — the Teamio-specific (`*.teamio.com`) meeting-body
 * parser. Real reference body this parser was built from (Czech, opaque
 * tracking tokens and the real Teams meeting ID fictionalized — see
 * test/meetings.test.js's own REAL_TEAMIO_CS_BODY_TEXT /
 * REAL_TEAMIO_CS_HTML_TEXT fixtures for the exact substitution):
 *
 *   Kdy: Pondělí, 24. srpna 2026
 *   Čas: 13:30, délka 30 minut
 *   Kde: https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345
 *
 *   Přejít na potvrzení pohovoru: https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890
 *
 * EXTRACTION IS LABEL-ANCHORED, not line-position-based (see
 * meetingsFindLabelLine's own doc): a fixed-line-position model is fragile
 * against real-world body rendering variance; label-anchored extraction
 * does not depend on which line a field happens to land on.
 *
 * Algorithm, in order:
 *   1. `rawText` is `String(bodyText || '')`; every U+00A0 (non-breaking
 *      space) is replaced with a regular space in a WORKING copy only —
 *      every thrown message below still reports the ORIGINAL rawText
 *      (same convention as ticketmasterCzNormalizeTicketText's own NBSP
 *      handling).
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
 *   5. Range-guard hour 0-23 and minute 0-59, HERE, once, rather than
 *      per-pack, with the SAME "Hour out of range" / "Minute out of range"
 *      wording the ticketing parsers already use (see parseEnigooTicketText),
 *      each ending with the full rawText, so every language pack gets the
 *      identical wording without duplicating the check.
 *   6. Location, in order: `meetingsExtractHtmlLabelValue(htmlBodyText,
 *      selectedPack.whereLabels)` first — a real, human-readable URL when
 *      the HTML body carries the matched pack's where-label in its own
 *      table-cell shape; else the plain-text where-line's own value,
 *      trimmed; else `''` when neither source has one — a meeting with no
 *      stated place is a real meeting, not a parse failure.
 *   7. Summary: the trimmed `subject` argument, or MEETINGS_FALLBACK_SUMMARY
 *      when empty/missing.
 *   8. Description: the matched when/time/where label lines, in source
 *      order (only the lines that were actually found). The when/time
 *      lines are the raw PLAIN-TEXT lines, verbatim. The where line is
 *      RECONSTRUCTED as `<label>: <location>` using step 6's FINAL
 *      resolved `location` value — NOT the raw plain-text line — so the
 *      event's `location` field and the description's own where line can
 *      never show two different values for the same field. Then, when
 *      meetingsHarvestBodyLinks(rawText, location) returns anything, a
 *      blank line, the literal heading `Links:`, and one URL per line,
 *      each pair of entries separated by a BLANK line for readability.
 *      Because `location` here is the HTML-preferred value, neither the
 *      plain-text where-line's own tracked redirect nor the confirm-CTA's
 *      tracked redirect is excluded from the Links: block — both
 *      legitimately surface there, while the description's own where line
 *      shows the clean HTML-sourced URL.
 *
 * Returns `{ summary, location, description, year, month, day, hour,
 * minute, durationMinutes }` (month ZERO-INDEXED, matching every other
 * date-components object in this codebase; `durationMinutes` is `null`
 * when the body states none — the default is applied by
 * processMeetingFromMessageBody, not here). The third `htmlBodyText`
 * argument is OPTIONAL: omitting it, or a pack/HTML shape that yields
 * nothing, falls back to the plain-text value, and never throws either
 * way. Pure, no GAS globals.
 */
function parseTeamioMeetingText(bodyText, subject, htmlBodyText) {
  const rawText = String(bodyText || '');

  // Pack selection is delegated to meetingsSelectLanguagePack (see its own
  // class-level doc): the SAME function findMeetingProcessingJobs's invitation
  // detector consults, so the gate and the parser can never disagree about
  // what counts as invitation-shaped text.
  const selection = meetingsSelectLanguagePack(rawText);

  // This throw is a DEFENSIVE INVARIANT, not a routine outcome:
  // findMeetingProcessingJobs refuses to create a job at all for a body this
  // function could not select a pack for, so reaching here means the gate
  // and the parser disagreed \u2014 a real bug worth surfacing loudly.
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

  // Location resolution: HTML-sourced value preferred (see
  // meetingsExtractHtmlLabelValue's own doc for why -- the plain-text
  // value is an opaque tracking redirect, not what should end up on the
  // calendar event), falling back to the plain-text where-line's own
  // value, falling back to '' when neither source has one.
  const htmlLocation = meetingsExtractHtmlLabelValue(htmlBodyText, selectedPack.whereLabels);
  const location = htmlLocation || (whereMatch ? whereMatch.value : '');
  const summary = subject && subject.trim() ? subject.trim() : MEETINGS_FALLBACK_SUMMARY;

  // The rendered "where" label line in the description must echo the SAME
  // resolved `location` value used above, not a second, independently-
  // sourced copy of the plain-text line -- otherwise the same field shows
  // two different values in two different places on the same generated
  // event. The when/time lines are unaffected -- only the where line has
  // an HTML-preferred alternate source at all.
  const matchedLabelLines = [whenMatch.lineText, timeMatch.lineText];
  if (whereMatch) {
    matchedLabelLines.push(whereMatch.label + ': ' + location);
  }

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
 * one file while still cleanly routing a resolved MEETINGS_ACTION_CONFIG
 * entry to the right parser. Consequence, deliberate
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
 * MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN — the companion registry to
 * MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN, keyed by the SAME `domainPattern`
 * string, mapping a system to its INVITATION DETECTOR: a pure,
 * never-throwing `(plainBodyText) => boolean` predicate answering "is this
 * message a meeting invitation at all?" WITHOUT attempting a parse.
 *
 * WHY A SECOND REGISTRY EXISTS. Matching a sender is not the same claim as
 * "this email is a meeting invitation". This action is the first to match a
 * whole DOMAIN, and `*.teamio.com` is an ATS platform that sends rejections
 * and status updates from the same domain as its invitations. The envelope
 * simply cannot carry the distinction, so the content must — before any job
 * is created.
 *
 * FAIL-CLOSED BY DESIGN: findMeetingProcessingJobs requires a REGISTERED
 * detector, exactly as it already requires a registered parser. A system
 * with a parser but no detector produces NO jobs rather than falling back
 * to envelope-only admission. Defaulting the other way ("no detector means
 * admit everything") is precisely the shape of the bug this prevents, and
 * would silently reintroduce it for the next system added. The cost is
 * that adding a system now needs a config entry + a parser + a detector;
 * the registry-key equality test in test/meetings.test.js enforces that
 * the two registries stay in step.
 *
 * Each detector's agreement with its parser on what counts as an
 * invitation is guaranteed by meetingsSelectLanguagePack's own FIDELITY
 * PROPERTY, not by anything in this registry itself.
 */
const MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN = {
  '*.teamio.com': teamioTextLooksLikeMeetingInvitation,
};

/**
 * meetingsIsIcsAttachment — true when `attachment`'s name ends in `.ics`
 * (case-insensitive) or its content-type is `text/calendar`. Reproduces
 * isIcsAttachment's (src/05-action-ics-import.js) exact rule LOCALLY —
 * this file must never cross-require the ICS action's own detection
 * helper, since the whole point of this gate living here (rather than as
 * an ICS_ACTION_CONFIG exclusion) is that this action's no-.ics scope is
 * entirely self-owned.
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
 * attachment (meetingsMessageHasIcsAttachment), AND (d) has a registered
 * INVITATION DETECTOR for that `domainPattern`
 * (MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN) which returns true for
 * the message's PLAIN-TEXT BODY. Every other message contributes nothing.
 * Pure in the sense that matters here — it touches no GAS global, only
 * methods on the passed-in message objects, so it is fully unit-testable
 * under Node with fake message/attachment objects. Never throws.
 *
 * Conditions (a)-(c) are all ENVELOPE evidence, and a sender domain is not
 * proof that an email is a meeting invitation, so condition (d)'s content
 * check belongs HERE rather than inside `run` precisely because this
 * resolver also backs MEETINGS_ACTION.appliesTo — gating only `run` would
 * leave appliesTo true, so the action would still CLAIM the thread and
 * mark it processed while doing nothing. Returning no job means the
 * action correctly does not apply at all.
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
// sibling src/10-action-cfg-meetings.js — see that file's own class-level
// JSDoc for the full load-order/getter rationale this mirrors, which
// points onward to src/07-action-cfg-ticketing-portals.js). Under GAS's
// shared global scope this is ALREADY visible here by bare name — no
// action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare `MEETINGS_ACTION_CONFIG` reference inside MEETINGS_ACTION's
// `config` getter below would otherwise throw ReferenceError.
if (typeof module !== 'undefined' && module.exports) {
  globalThis.MEETINGS_ACTION_CONFIG = require('./10-action-cfg-meetings.js').MEETINGS_ACTION_CONFIG;
}

// --- the GAS pipeline (D-03, D-07), GAS-only, not unit-tested --------------

/**
 * findMeetingEventByIdentifier — the DEDUP SAFETY NET's lookup:
 * `Calendar.Events.list(calendarId, { privateExtendedProperty:
 * 'meetingIdentifier=' + meetingIdentifier, singleEvents: true })`.
 * Deliberately NOT paginated or time-windowed — a `privateExtendedProperty`
 * filter against a near-certainly-unique per-meeting identifier is already
 * an exact match expected to return 0 or 1 events. Returns the first
 * matching event, or `null`. GAS-only (Calendar global). Mirrors
 * findTicketEventByIdentifier's (src/07-action-ticketing-portals.js) exact
 * query shape.
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
 * actually produces one, but this is a defensive branch), otherwise `true`
 * (plus a `Meetings:`-prefixed already-exists log line) when an event
 * already carries that exact tag on `calendarId`. GAS-only (calls
 * findMeetingEventByIdentifier).
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
 * `extendedProperties.private.meetingIdentifier`. TIMEZONE resolved LIVE
 * from the target calendar (`CalendarApp.getCalendarById(calendarId)
 * .getTimeZone()`), never a hardcoded assumption. GAS-only
 * (CalendarApp/Calendar globals).
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
 * processMeetingFromMessageBody — the per-job pipeline: reads
 * `message.getPlainBody()`/`getSubject()`/`getBody()` — the HTML body is
 * read because, for a Teamio invitation, the real human-readable meeting
 * URL exists only there (see meetingsExtractHtmlLabelValue's own doc) —
 * parses via `parser(bodyText, subject, htmlBodyText)`, resolves the
 * target calendar (resolveMeetingsCalendarId(system, CONFIG.calendarId)),
 * builds the dedup identifier from the sender + parsed summary + parsed
 * start wall-clock, returns early (no writes at all) when
 * isDuplicateMeetingInvite says so, otherwise resolves the duration as
 * `parsedMeeting.durationMinutes ||
 * MEETINGS_ACTION.config.defaultDurationMinutes` (the body's own stated
 * duration wins when present) and creates the event, logging a
 * `Meetings:`-prefixed line naming the summary and the resolved
 * identifier. GAS-only (GmailMessage/CalendarApp/Calendar globals).
 *
 * SENDER ATTRIBUTION: immediately after the parser call, the parsed
 * meeting is passed through meetingsApplySenderAttribution before
 * anything downstream ever sees it. This is THE single choke point every
 * meeting event travels — present and future — regardless of which parser
 * produced it, which is what makes the "every event identifies its
 * sender" guarantee hold unconditionally rather than per-parser. See
 * meetingsApplySenderAttribution's own class-level doc for the full
 * rationale.
 */
function processMeetingFromMessageBody(message, system, parser) {
  const bodyText = message.getPlainBody();
  const subject = message.getSubject();
  const htmlBodyText = message.getBody();
  const parsedMeeting = meetingsApplySenderAttribution(parser(bodyText, subject, htmlBodyText), message.getFrom());

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
// Deliberately exports NO bare `addMinutesToWallClockComponents` or
// `formatWallClockComponentsIso` key — see this file's class-level
// "GLOBALLY-UNIQUE NAMING" doc — only the namespaced `meetings*` versions
// are exported. findMeetingEventByIdentifier/createMeetingCalendarEvent/
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
    meetingsExtractSenderDisplayName: meetingsExtractSenderDisplayName,
    meetingsFormatSenderAttribution: meetingsFormatSenderAttribution,
    meetingsApplySenderAttribution: meetingsApplySenderAttribution,
    MEETINGS_ACTION: MEETINGS_ACTION,
  };
}

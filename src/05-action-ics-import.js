/**
 * IcsParser — pure iCalendar (RFC 5545) VEVENT parser.
 *
 * No GAS globals are referenced here so this file loads unmodified both in
 * the Apps Script runtime and under plain Node (see the guarded export at
 * the bottom) — it is exercised by the Node test suite in test/.
 *
 * Trust boundary: the input `text` originates from an UNTRUSTED inbound
 * email attachment. This parser must never execute, evaluate, or persist
 * raw attacker bytes beyond structural field extraction, and must fail
 * closed (throw) on malformed/pathological input rather than silently
 * accepting it.
 *
 * CONFIG SPLIT: ICS_ACTION_CONFIG (this action's tunable settings) lives
 * in the sibling src/05-action-cfg-ics-import.js, not in this file — GAS's
 * shared global scope means it's still visible by bare name here, no
 * require/import needed. ICS_CALENDAR_ACTION's `config` property below is
 * an ES6 GETTER (not a plain literal reference), which is what makes this
 * split safe regardless of which of the two files loads first
 * alphabetically — see ICS_ACTION_CONFIG's own doc comment in that sibling
 * file for the full load-order explanation. (Under Node, where each
 * `require()`d file is its own isolated module with no shared global
 * scope, a `globalThis` bridge near the bottom of this file makes the
 * getter's bare `ICS_ACTION_CONFIG` reference resolve correctly too — see
 * that bridge's own comment for why `globalThis` and not a redeclared
 * `const`/`let`/`var`.)
 */

/**
 * Normalizes line endings and un-folds RFC 5545 folded lines: a line that
 * begins with a single space or horizontal tab is a continuation of the
 * previous line (the leading whitespace character is removed on rejoin).
 */
function unfoldLines(text) {
  const normalized = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const rawLines = normalized.split('\n');
  const lines = [];

  rawLines.forEach(function (line) {
    if ((line.charAt(0) === ' ' || line.charAt(0) === '\t') && lines.length > 0) {
      lines[lines.length - 1] += line.slice(1);
    } else {
      lines.push(line);
    }
  });

  return lines;
}

/**
 * Splits unfolded lines into one array of raw property lines per
 * BEGIN:VEVENT..END:VEVENT block, preserving source order. Lines outside
 * any VEVENT block (VCALENDAR/VTIMEZONE/etc.) are ignored.
 *
 * NESTING FIX (TR-1/T-gks-02): a VEVENT may itself contain nested
 * sub-components (BEGIN:VALARM..END:VALARM being the common case, but this
 * is handled generically for any sub-component type, matched by prefix, not
 * a hardcoded name). Such a sub-component's own properties (e.g. a VALARM's
 * own DESCRIPTION:REMINDER) must NOT be collected into the VEVENT's
 * propertyLines, or they silently overwrite the VEVENT's real properties in
 * the later last-wins props map (parseVeventBlock). A skipDepth counter
 * tracks nested BEGIN:/END: pairs seen while already inside an open VEVENT;
 * lines are only pushed into `current` while skipDepth === 0. This nesting
 * logic applies only to sub-components found INSIDE an open VEVENT — the
 * VEVENT delimiters themselves are handled by the existing branches above
 * and are entirely unaffected.
 */
function extractVeventBlocks(lines) {
  const blocks = [];
  let current = null;
  let skipDepth = 0;

  lines.forEach(function (line) {
    const trimmed = line.trim().toUpperCase();

    if (trimmed === 'BEGIN:VEVENT') {
      current = [];
      skipDepth = 0;
    } else if (trimmed === 'END:VEVENT') {
      if (current) {
        blocks.push(current);
      }
      current = null;
      skipDepth = 0;
    } else if (current && trimmed.indexOf('BEGIN:') === 0) {
      // Nested sub-component (e.g. VALARM) opening inside an open VEVENT.
      skipDepth += 1;
    } else if (current && trimmed.indexOf('END:') === 0 && skipDepth > 0) {
      // Matching nested sub-component close.
      skipDepth -= 1;
    } else if (current && skipDepth === 0) {
      current.push(line);
    }
  });

  return blocks;
}

/**
 * Extracts every BEGIN:VTIMEZONE..END:VTIMEZONE block into a map keyed by
 * TZID, each value shaped { standard, daylight } (either may be null if the
 * sub-block is absent/incomplete). Each populated sub-block is
 * { dtstartTime: {hour,minute,second}, offsetFromMinutes, offsetToMinutes,
 * byMonth, byDay } — the normalized transition rule used to resolve which
 * offset applies to a given wall-clock instant (see resolveTzidDate).
 *
 * TZID wall-clock times are resolved from the .ics's OWN embedded
 * VTIMEZONE data, never from a hardcoded Windows-name mapping table or the
 * script's own appsscript.json timeZone (owner-directed fix for a live
 * timezone bug).
 */
function extractVtimezoneBlocks(lines) {
  const vtimezones = {};
  let inVtimezone = false;
  let currentTzid = null;
  let currentSub = null;
  let subLines = null;

  lines.forEach(function (line) {
    const trimmed = line.trim().toUpperCase();

    if (trimmed === 'BEGIN:VTIMEZONE') {
      inVtimezone = true;
      currentTzid = null;
      return;
    }
    if (trimmed === 'END:VTIMEZONE') {
      inVtimezone = false;
      currentTzid = null;
      currentSub = null;
      subLines = null;
      return;
    }
    if (!inVtimezone) {
      return;
    }
    if (trimmed === 'BEGIN:STANDARD' || trimmed === 'BEGIN:DAYLIGHT') {
      currentSub = trimmed === 'BEGIN:STANDARD' ? 'standard' : 'daylight';
      subLines = [];
      return;
    }
    if (trimmed === 'END:STANDARD' || trimmed === 'END:DAYLIGHT') {
      if (currentTzid) {
        vtimezones[currentTzid] = vtimezones[currentTzid] || { standard: null, daylight: null };
        vtimezones[currentTzid][currentSub] = parseTransitionSubBlock(subLines);
      }
      currentSub = null;
      subLines = null;
      return;
    }
    if (currentSub) {
      subLines.push(line);
      return;
    }

    const parsed = parsePropertyLine(line);
    if (parsed && parsed.name === 'TZID') {
      currentTzid = parsed.value;
    }
  });

  return vtimezones;
}

/**
 * Parses one STANDARD/DAYLIGHT sub-block's raw property lines into a
 * normalized transition rule, or null if the sub-block is missing any of
 * the fields needed to resolve a transition (caller falls back to
 * literal-UTC treatment for that TZID, preserving the documented v1
 * limitation for that narrower case).
 */
function parseTransitionSubBlock(propertyLines) {
  const props = {};
  propertyLines.forEach(function (line) {
    const parsed = parsePropertyLine(line);
    if (parsed) {
      props[parsed.name] = parsed;
    }
  });

  if (!props.DTSTART || !props.TZOFFSETFROM || !props.TZOFFSETTO || !props.RRULE) {
    return null;
  }

  const dtstartMatch = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$/.exec(props.DTSTART.value);
  if (!dtstartMatch) {
    return null;
  }

  let rrule;
  try {
    rrule = parseRRule(props.RRULE.value);
  } catch (e) {
    return null;
  }

  if (!rrule.byMonth || !rrule.byDay) {
    return null;
  }

  return {
    dtstartTime: {
      hour: Number(dtstartMatch[4]),
      minute: Number(dtstartMatch[5]),
      second: Number(dtstartMatch[6]),
    },
    offsetFromMinutes: parseUtcOffset(props.TZOFFSETFROM.value),
    offsetToMinutes: parseUtcOffset(props.TZOFFSETTO.value),
    byMonth: rrule.byMonth,
    byDay: rrule.byDay,
  };
}

/**
 * Parses an RFC 5545 UTC-offset value ("+0200"/"-0500") into signed minutes.
 */
function parseUtcOffset(value) {
  const match = /^([+-])(\d{2})(\d{2})$/.exec(value);
  if (!match) {
    throw new Error('Unrecognized UTC offset value: ' + value);
  }
  const sign = match[1] === '-' ? -1 : 1;
  return sign * (Number(match[2]) * 60 + Number(match[3]));
}

// Weekday codes used by RFC 5545 BYDAY tokens, index-compatible with
// Date#getUTCDay() (0 = Sunday).
const BYDAY_WEEKDAY_CODES = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

/**
 * Parses a single BYDAY token ("-1SU", "1MO", "SU") into
 * { ordinal, weekdayIndex }. ordinal defaults to 1 (first) when no signed
 * prefix is present; a negative ordinal counts from the end of the month
 * (-1 = last). Ordinal `0` is not valid per RFC 5545 and is rejected
 * explicitly rather than silently mis-computed as a negative ordinal.
 */
function parseByDayToken(token) {
  const match = /^(-?\d+)?(SU|MO|TU|WE|TH|FR|SA)$/.exec(token);
  if (!match) {
    throw new Error('Unsupported BYDAY token: ' + token);
  }
  const ordinal = match[1] ? parseInt(match[1], 10) : 1;
  if (ordinal === 0) {
    throw new Error('Invalid BYDAY ordinal (0 is not valid per RFC 5545): ' + token);
  }
  return {
    ordinal: ordinal,
    weekdayIndex: BYDAY_WEEKDAY_CODES.indexOf(match[2]),
  };
}

/**
 * Computes the UTC calendar date (midnight) of the nth (or, for a negative
 * ordinal, the nth-from-last) occurrence of a weekday in a given UTC
 * year/month. month0 is zero-indexed (0 = January), matching Date.UTC.
 * Throws if the requested ordinal has no such occurrence in the month
 * (e.g. "5th Monday" in a month with only 4) rather than silently rolling
 * into an adjacent month via Date.UTC's overflow behavior.
 */
function nthWeekdayOfMonthUTC(year, month0, weekdayIndex, ordinal) {
  let day;

  if (ordinal > 0) {
    const firstOfMonth = new Date(Date.UTC(year, month0, 1));
    const firstWeekday = firstOfMonth.getUTCDay();
    const dayOffset = (weekdayIndex - firstWeekday + 7) % 7;
    day = 1 + dayOffset + (ordinal - 1) * 7;
  } else {
    const nextMonthFirst = new Date(Date.UTC(year, month0 + 1, 1));
    const lastDayOfMonth = new Date(nextMonthFirst.getTime() - 24 * 60 * 60 * 1000);
    const lastWeekday = lastDayOfMonth.getUTCDay();
    const dayOffset = (lastWeekday - weekdayIndex + 7) % 7;
    day = lastDayOfMonth.getUTCDate() - dayOffset - (Math.abs(ordinal) - 1) * 7;
  }

  const daysInMonth = new Date(Date.UTC(year, month0 + 1, 0)).getUTCDate();
  if (day < 1 || day > daysInMonth) {
    throw new Error(
      'BYDAY ordinal ' + ordinal + ' has no matching occurrence in ' + year + '-' + (month0 + 1)
    );
  }

  return new Date(Date.UTC(year, month0, day));
}

/**
 * Computes the yearly STANDARD/DAYLIGHT transition instant for a given year,
 * expressed in the SAME frame as `provisionalUtcMs` in
 * `resolveTzidOffsetMinutes`: the transition's wall-clock digits
 * reinterpreted as a literal UTC timestamp (i.e. NOT a real UTC instant —
 * deliberately not offset-corrected). This must stay in that frame so the
 * comparison in `resolveTzidOffsetMinutes` is like-with-like: mixing a real
 * UTC instant against a wall-clock-literal produces an answer that is wrong
 * by exactly the transition's UTC offset for any event whose wall-clock
 * time falls between the two.
 */
function computeTransitionWallInstantUTC(year, transition) {
  const parsed = parseByDayToken(transition.byDay[0]);
  const month0 = transition.byMonth[0] - 1;
  const transitionDate = nthWeekdayOfMonthUTC(year, month0, parsed.weekdayIndex, parsed.ordinal);
  return Date.UTC(
    transitionDate.getUTCFullYear(),
    transitionDate.getUTCMonth(),
    transitionDate.getUTCDate(),
    transition.dtstartTime.hour,
    transition.dtstartTime.minute,
    transition.dtstartTime.second
  );
}

/**
 * Resolves which offset (in minutes) applies to a wall-clock instant
 * (expressed as a Date.UTC() epoch computed directly from the wall-clock
 * digits, i.e. NOT yet offset-corrected), given a VTIMEZONE's
 * standard/daylight transition rules for that instant's year. Handles both
 * northern-hemisphere (DST is the inner window) and southern-hemisphere
 * (STANDARD is the inner window) transition orderings.
 *
 * Both `provisionalUtcMs` and the transition instants below are wall-clock
 * digits reinterpreted as literal UTC (never real UTC instants), so this
 * comparison is like-with-like.
 */
function resolveTzidOffsetMinutes(provisionalUtcMs, vtimezone) {
  const year = new Date(provisionalUtcMs).getUTCFullYear();
  const daylightStart = computeTransitionWallInstantUTC(year, vtimezone.daylight);
  const standardStart = computeTransitionWallInstantUTC(year, vtimezone.standard);

  if (daylightStart < standardStart) {
    if (provisionalUtcMs >= daylightStart && provisionalUtcMs < standardStart) {
      return vtimezone.daylight.offsetToMinutes;
    }
    return vtimezone.standard.offsetToMinutes;
  }

  if (provisionalUtcMs >= standardStart && provisionalUtcMs < daylightStart) {
    return vtimezone.standard.offsetToMinutes;
  }
  return vtimezone.daylight.offsetToMinutes;
}

/**
 * Resolves a TZID-qualified wall-clock date/time into a correct UTC Date
 * using the matching VTIMEZONE's embedded transition rules, or returns null
 * if that VTIMEZONE is incomplete (caller falls back to literal-UTC
 * treatment for that TZID). When only a STANDARD sub-block is present (a
 * VTIMEZONE for a TZID with no DST, e.g. "China Standard Time"), resolves
 * directly against its fixed offsetToMinutes rather than requiring both
 * STANDARD and DAYLIGHT sub-blocks.
 */
function resolveTzidDate(year, month, day, hour, minute, second, vtimezone) {
  if (!vtimezone) {
    return null;
  }
  const provisionalUtcMs = Date.UTC(year, month, day, hour, minute, second);

  if (vtimezone.standard && vtimezone.daylight) {
    const offsetMinutes = resolveTzidOffsetMinutes(provisionalUtcMs, vtimezone);
    return new Date(provisionalUtcMs - offsetMinutes * 60000);
  }
  if (vtimezone.standard && !vtimezone.daylight) {
    return new Date(provisionalUtcMs - vtimezone.standard.offsetToMinutes * 60000);
  }
  return null;
}

/**
 * Parses one raw "NAME;PARAM=VALUE;...:VALUE" property line into
 * { name, params, value }. Returns null for blank/unparseable lines.
 */
function parsePropertyLine(line) {
  if (!line) {
    return null;
  }

  const colonIndex = line.indexOf(':');
  if (colonIndex === -1) {
    return null;
  }

  const left = line.slice(0, colonIndex);
  const value = line.slice(colonIndex + 1);
  const segments = left.split(';');
  const name = segments[0].toUpperCase();
  const params = {};

  for (let i = 1; i < segments.length; i++) {
    const eqIndex = segments[i].indexOf('=');
    if (eqIndex !== -1) {
      const paramName = segments[i].slice(0, eqIndex).toUpperCase();
      const paramValue = segments[i].slice(eqIndex + 1);
      params[paramName] = paramValue;
    }
  }

  return { name: name, params: params, value: value };
}

/**
 * Parses an RFC 5545 DATE or DATE-TIME value (YYYYMMDD or
 * YYYYMMDDTHHMMSS[Z]) into a UTC Date instant. Z-suffixed values are exact
 * UTC instants. A non-Z-suffixed date-time whose property carried a `tzid`
 * matching an entry in `vtimezones` is resolved against that VTIMEZONE's
 * embedded STANDARD/DAYLIGHT transition rules (see extractVtimezoneBlocks).
 * Any other floating/date-only value, or a TZID with no matching VTIMEZONE
 * data, is interpreted literally in the script/calendar timezone (Etc/UTC)
 * per the documented v1 timezone bound for that narrower case.
 */
function parseIcsDate(value, tzid, vtimezones) {
  const match = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/.exec(value);
  if (!match) {
    throw new Error('Unrecognized ICS date/date-time value: ' + value);
  }

  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);

  if (match[4] === undefined) {
    return { date: new Date(Date.UTC(year, month, day)), isAllDay: true };
  }

  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const isUtcSuffixed = match[7] === 'Z';

  if (!isUtcSuffixed && tzid && vtimezones && vtimezones[tzid]) {
    const resolved = resolveTzidDate(year, month, day, hour, minute, second, vtimezones[tzid]);
    if (resolved) {
      return { date: resolved, isAllDay: false };
    }
  }

  return { date: new Date(Date.UTC(year, month, day, hour, minute, second)), isAllDay: false };
}

/**
 * Un-escapes RFC 5545 TEXT value escaping (\n, \, , \; , \\) so DESCRIPTION/
 * SUMMARY/LOCATION are rendered as intended, never left with literal
 * backslash-escapes.
 */
function unescapeText(value) {
  return String(value).replace(/\\(.)/g, function (match, ch) {
    if (ch === 'n' || ch === 'N') {
      return '\n';
    }
    if (ch === ',' || ch === ';' || ch === '\\') {
      return ch;
    }
    return match;
  });
}

// RRULE FREQ values this parser accepts. Sub-daily frequencies (SECONDLY/
// MINUTELY/HOURLY) are rejected outright — they bound the amount of
// recurrence metadata the script has to reason about and guard against a
// pathological/hostile .ics driving unbounded work. Recurrence expansion
// itself is always delegated to the Advanced Calendar Service server-side
// (via buildRRuleString + Calendar.Events.import/insert); this parser only
// produces a normalized descriptor, never enumerates occurrences.
const SUPPORTED_RRULE_FREQS = ['DAILY', 'WEEKLY', 'MONTHLY', 'YEARLY'];

/**
 * Parses an RRULE property value into a normalized recurrence descriptor:
 * { freq, interval, count, until, byDay }, plus an optional `byMonth`
 * (array of 1-12 integers) present only when the source RRULE included
 * BYMONTH — added to support VTIMEZONE STANDARD/DAYLIGHT transition rules
 * without changing the descriptor shape for ordinary VEVENT recurrences
 * that never set BYMONTH. Throws a controlled Error for an
 * unsupported/pathological FREQ rather than silently accepting it.
 */
function parseRRule(value) {
  const rule = {};

  value.split(';').forEach(function (part) {
    const eqIndex = part.indexOf('=');
    if (eqIndex === -1) {
      return;
    }
    rule[part.slice(0, eqIndex).toUpperCase()] = part.slice(eqIndex + 1);
  });

  if (!rule.FREQ || SUPPORTED_RRULE_FREQS.indexOf(rule.FREQ) === -1) {
    throw new Error('Unsupported recurrence FREQ: ' + rule.FREQ);
  }

  const descriptor = {
    freq: rule.FREQ,
    interval: rule.INTERVAL ? parseInt(rule.INTERVAL, 10) : 1,
    count: rule.COUNT ? parseInt(rule.COUNT, 10) : null,
    until: rule.UNTIL ? parseIcsDate(rule.UNTIL).date : null,
    byDay: rule.BYDAY ? rule.BYDAY.split(',') : null,
  };

  if (rule.BYMONTH) {
    descriptor.byMonth = rule.BYMONTH.split(',').map(function (n) {
      return parseInt(n, 10);
    });
  }

  return descriptor;
}

/**
 * parseAddressProperty — normalizes a parsed ORGANIZER/ATTENDEE property
 * (see parsePropertyLine) into { name, email }. `name` is the CN parameter
 * when present, or null otherwise. `email` is the property value with any
 * leading case-insensitive 'mailto:' prefix stripped and the result
 * trimmed; a value with no such prefix is used as-is. Returns null for a
 * null/undefined input (no ORGANIZER/ATTENDEE property present). Pure, no
 * GAS globals.
 */
function parseAddressProperty(parsedProp) {
  if (!parsedProp) {
    return null;
  }

  const rawValue = parsedProp.value || '';
  const email = rawValue.replace(/^mailto:/i, '').trim();

  return {
    name: parsedProp.params && parsedProp.params.CN ? parsedProp.params.CN : null,
    email: email,
  };
}

/**
 * formatAddress — renders a { name, email } address (see
 * parseAddressProperty) as display text: 'Name <email>' when a name is
 * present and case-insensitively different from the email (the common
 * case), or the bare email alone when name is null OR case-insensitively
 * equal to the email (a real Exchange pattern where CN duplicates the
 * mailto address, in which case a second copy of the address adds no
 * information). Pure, no GAS globals.
 */
function formatAddress(address) {
  const name = address.name;
  const email = address.email;

  if (name && name.toLowerCase() !== email.toLowerCase()) {
    return name + ' <' + email + '>';
  }

  return email;
}

/**
 * buildOrganizerAttendeesText — renders organizer/attendee informational
 * text to append to a VEVENT's description (TR-2..TR-5). `organizer` is a
 * single { name, email } address or null; `attendees` is a (possibly empty)
 * array of { name, email } addresses. Emits an 'Organizer: ...' line only
 * when organizer is non-null, and an 'Attendees: ...' line (comma-joined)
 * only when attendees is non-empty; the two lines are joined by a single
 * newline when both are present. Returns '' (a strict no-op) when organizer
 * is null AND attendees is empty, so callers can append unconditionally
 * without special-casing the absent case. Pure, no GAS globals.
 */
function buildOrganizerAttendeesText(organizer, attendees) {
  const lines = [];

  if (organizer) {
    lines.push('Organizer: ' + formatAddress(organizer));
  }
  if (attendees && attendees.length > 0) {
    lines.push('Attendees: ' + attendees.map(formatAddress).join(', '));
  }

  return lines.join('\n');
}

/**
 * collapseBlankLines — replaces any run of 3-or-more consecutive `\n`
 * characters with exactly `\n\n` (i.e. two-or-more consecutive empty lines
 * collapse down to a single blank-line separator). A single blank line
 * (`\n\n`, exactly two newlines) is left alone and NOT reduced further to
 * zero. Real Exchange/Outlook invites often carry long runs of blank lines
 * in their raw DESCRIPTION; this keeps the rendered event description
 * readable. Pure, no GAS globals.
 */
function collapseBlankLines(text) {
  return text.replace(/\n{3,}/g, '\n\n');
}

/**
 * Parses one VEVENT block's raw property lines into a normalized event
 * object: { summary, start, end, isAllDay, description, location,
 * recurrence, sequence, uid, status, dtstamp }. recurrence is null for non-recurring events,
 * or a normalized descriptor (see parseRRule) when RRULE is present.
 * `vtimezones` (see extractVtimezoneBlocks) is threaded through to resolve
 * any TZID-qualified DTSTART/DTEND against the .ics's own embedded timezone
 * data.
 *
 * SEQUENCE (RFC 5545 section 3.8.7.4, live-reported bug quick-260731-seq):
 * `sequence` is ALWAYS a real non-negative integer on the returned event
 * object, never `undefined`/`null` — parsed from the VEVENT's own SEQUENCE
 * property when present (`SEQUENCE:1` -> `1`), or defaulting to `0` when
 * the property is absent, exactly matching RFC 5545's documented default
 * for an absent SEQUENCE. A malformed (non-numeric) SEQUENCE value ALSO
 * falls back to `0` rather than propagating `NaN` — this is optional
 * scheduling metadata, not something a malformed value should crash the
 * parser over. This field exists because `Calendar.Events.import()`
 * previously received no `sequence` at all (see buildEventResource) and
 * therefore implicitly sent `0`; when a genuine invite sent to the
 * owner's own Gmail address had ALREADY been detected and stored by
 * Google's native Gmail-to-Calendar detection with a real, higher
 * sequence number embedded in the SAME .ics, our own import was rejected
 * as a stale/out-of-order update (`GoogleJsonResponseException: Invalid
 * sequence value...`) — the exact same family of "our own import
 * collides with Google's native detection" issue as the original iCalUID
 * dedup fix, surfacing through a different field.
 *
 * STATUS (RFC 5545 section 3.8.1.11, RegioJet cancellation detection, D-01/
 * D-02 of quick-260813-dq2): `status` is the VEVENT's own STATUS property
 * value, trimmed and uppercased, or `null` when the property is absent OR
 * when the trimmed value is empty (so no caller ever has to distinguish an
 * empty value from an absent property). This exists so TRANSPORT_TICKETS_ACTION
 * can detect a RegioJet cancellation purely from `event.status === 'CANCELLED'`
 * — a fixed, language-independent RFC 5545 token, deliberately never from
 * email subject/body text (which vary per RegioJet locale). This field is
 * parser-level only and purely additive: `buildEventResource` deliberately
 * does NOT copy it onto the Calendar API resource, so every other caller
 * (ICS_CALENDAR_ACTION included) is bit-for-bit unchanged.
 *
 * DTSTAMP (RFC 5545 section 3.8.7.2, RegioJet cancel/rebook staleness
 * detection, D-09/D-10/D-11 of quick-260813-dq2 Task 3): `dtstamp` is the
 * VEVENT's own DTSTAMP property, parsed through the EXISTING parseIcsDate
 * helper (RFC 5545 DTSTAMP is always UTC Z-suffixed, so no TZID is ever
 * passed) into a real Date when present, or `null` when the property is
 * absent OR when parseIcsDate throws on an unrecognized value. A malformed
 * DTSTAMP must never propagate a throw out of this function — exactly the
 * same "optional scheduling metadata must not crash an otherwise-valid
 * import" discipline the SEQUENCE fallback above already documents. This
 * field exists so TRANSPORT_TICKETS_ACTION can detect a stale cancellation
 * (one superseded by a later rebooking): RegioJet RESETS SEQUENCE across a
 * cancel+rebook pair, but DTSTAMP — real send time — stays monotonic.
 * Parser-level only and purely additive, same D-01 firewall as `status`:
 * `buildEventResource` deliberately does NOT copy this field onto the
 * Calendar API resource either.
 *
 * ENRICHMENT (TR-2..TR-5): ORGANIZER (single-valued per RFC 5545) is read
 * via the last-wins `props` map; ATTENDEE (multi-valued) is collected in a
 * SEPARATE pass over the raw `propertyLines` so multiple attendees all
 * survive (the last-wins props map would otherwise keep only the final
 * ATTENDEE line). Both are rendered as pure informational TEXT PREPENDED to
 * description (before the original description text, separated by a blank
 * line) — never surfaced as resource-level attendees/organizer fields (see
 * buildEventResource, the T-03-05 safety firewall). The Teams/meeting URL
 * (X-MICROSOFT-SKYPETEAMSMEETINGURL) REPLACES location entirely when
 * present (owner preference: the original LOCATION text and the meeting
 * URL together were confusing; only the URL is kept), taken RAW (never run
 * through unescapeText — a URI must not be TEXT-unescaped, same treatment
 * as UID below). When no meeting URL is present, location falls back to
 * the original LOCATION value unchanged. The final assembled description
 * is passed through collapseBlankLines so long runs of blank lines common
 * in real Exchange/Outlook DESCRIPTION values render as a single
 * blank-line separator.
 */
function parseVeventBlock(propertyLines, vtimezones) {
  const props = {};

  propertyLines.forEach(function (line) {
    const parsed = parsePropertyLine(line);
    if (parsed) {
      props[parsed.name] = parsed;
    }
  });

  if (!props.DTSTART) {
    throw new Error('VEVENT missing required DTSTART');
  }

  const startInfo = parseIcsDate(props.DTSTART.value, props.DTSTART.params.TZID, vtimezones);
  const endInfo = props.DTEND
    ? parseIcsDate(props.DTEND.value, props.DTEND.params.TZID, vtimezones)
    : startInfo;

  const organizer = props.ORGANIZER ? parseAddressProperty(props.ORGANIZER) : null;
  const attendees = propertyLines
    .map(parsePropertyLine)
    .filter(function (parsed) {
      return parsed && parsed.name === 'ATTENDEE';
    })
    .map(parseAddressProperty);

  const baseDescription = props.DESCRIPTION ? unescapeText(props.DESCRIPTION.value) : '';
  const organizerAttendeesText = buildOrganizerAttendeesText(organizer, attendees);
  const description = collapseBlankLines(
    organizerAttendeesText
      ? baseDescription
        ? organizerAttendeesText + '\n\n' + baseDescription
        : organizerAttendeesText
      : baseDescription
  );

  const baseLocation = props.LOCATION ? unescapeText(props.LOCATION.value) : '';
  const meetingUrl = props['X-MICROSOFT-SKYPETEAMSMEETINGURL']
    ? props['X-MICROSOFT-SKYPETEAMSMEETINGURL'].value
    : null;
  // Owner preference: when a meeting URL is present it REPLACES location
  // entirely (not combined with the original LOCATION text); absent, location
  // falls back to the original LOCATION value unchanged.
  const location = meetingUrl ? meetingUrl : baseLocation;

  // SEQUENCE: absent -> 0 (RFC 5545 3.8.7.4's documented default);
  // malformed (non-numeric) -> also 0, never NaN — see this function's
  // class-level "SEQUENCE" doc paragraph above for the full rationale.
  const parsedSequence = props.SEQUENCE ? parseInt(props.SEQUENCE.value, 10) : 0;
  const sequence = Number.isNaN(parsedSequence) ? 0 : parsedSequence;

  // STATUS: absent, or empty after trimming -> null; otherwise trimmed and
  // uppercased so every downstream consumer compares against a single
  // normalized token (see this function's class-level "STATUS" doc
  // paragraph above).
  const trimmedStatus = props.STATUS ? String(props.STATUS.value).trim() : '';
  const status = trimmedStatus ? trimmedStatus.toUpperCase() : null;

  // DTSTAMP: absent -> null; malformed/unrecognized -> also null, never
  // throw — see this function's class-level "DTSTAMP" doc paragraph above.
  // RFC 5545 DTSTAMP is always UTC Z-suffixed, so no TZID is passed.
  let dtstamp = null;
  if (props.DTSTAMP) {
    try {
      dtstamp = parseIcsDate(props.DTSTAMP.value, undefined, vtimezones).date;
    } catch (e) {
      dtstamp = null;
    }
  }

  return {
    summary: props.SUMMARY ? unescapeText(props.SUMMARY.value) : '',
    start: startInfo.date,
    end: endInfo.date,
    isAllDay: startInfo.isAllDay,
    description: description,
    location: location,
    recurrence: props.RRULE ? parseRRule(props.RRULE.value) : null,
    sequence: sequence,
    // UID is an opaque RFC 5545 identity token, taken RAW (never passed
    // through unescapeText) — it becomes the Advanced Calendar API resource's
    // iCalUID, the key Google Calendar's own dedup logic matches against.
    uid: props.UID ? props.UID.value : null,
    status: status,
    dtstamp: dtstamp,
  };
}

/**
 * parseIcs — parses raw .ics text into an ordered array of normalized event
 * objects, one per VEVENT block, in source order. Returns [] for input with
 * zero VEVENT blocks. Any BEGIN:VTIMEZONE blocks present are parsed once up
 * front and used to resolve TZID-qualified wall-clock times in every VEVENT.
 */
function parseIcs(text) {
  const lines = unfoldLines(text);
  const vtimezones = extractVtimezoneBlocks(lines);
  const blocks = extractVeventBlocks(lines);
  return blocks.map(function (block) {
    return parseVeventBlock(block, vtimezones);
  });
}

/**
 * formatIcsUtcTimestamp — turns a UTC Date into RFC 5545 UTC form
 * (YYYYMMDDTHHMMSSZ), derived from Date#toISOString() by stripping the '-'
 * and ':' separators and the fractional-seconds portion, leaving the
 * trailing Z. Pure, no GAS globals.
 */
function formatIcsUtcTimestamp(date) {
  return date.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');
}

/**
 * buildRRuleString — turns a normalized recurrence descriptor (see
 * parseRRule) into a single RFC 5545 'RRULE:'-prefixed string, suitable for
 * the Advanced Calendar API's resource.recurrence array. Always emits FREQ.
 * INTERVAL is omitted when absent or === 1 (matches the parser's own
 * convention). COUNT/UNTIL/BYDAY/BYMONTH are emitted only when present.
 * BYDAY tokens are joined with commas and preserved RAW (ordinal prefixes
 * such as '-1FR' are kept — the Advanced API's native RRULE parser supports
 * ordinal BYDAY, unlike the old CalendarApp.Recurrence.onlyOnWeekdays path).
 * Pure, no GAS globals.
 */
function buildRRuleString(descriptor) {
  const parts = ['FREQ=' + descriptor.freq];

  if (descriptor.interval && descriptor.interval !== 1) {
    parts.push('INTERVAL=' + descriptor.interval);
  }
  if (descriptor.count) {
    parts.push('COUNT=' + descriptor.count);
  }
  if (descriptor.until) {
    parts.push('UNTIL=' + formatIcsUtcTimestamp(descriptor.until));
  }
  if (descriptor.byDay && descriptor.byDay.length > 0) {
    parts.push('BYDAY=' + descriptor.byDay.join(','));
  }
  if (descriptor.byMonth && descriptor.byMonth.length > 0) {
    parts.push('BYMONTH=' + descriptor.byMonth.join(','));
  }

  return 'RRULE:' + parts.join(';');
}

/**
 * buildEventResource — turns a normalized parsed event (see parseVeventBlock)
 * into an Advanced Calendar API event resource (the shape accepted by
 * Calendar.Events.import/insert). summary/description/location are copied
 * directly. iCalUID is set only when event.uid is truthy (omitted entirely
 * otherwise — Events.insert does not need one). All-day events get
 * date-only {date:'YYYY-MM-DD'} start/end (exclusive-end, same convention
 * the old CalendarApp.createAllDayEvent used); timed events get
 * {dateTime: ISO string} start/end. recurrence is set to a one-element
 * array [buildRRuleString(event.recurrence)] only when the event recurs.
 *
 * SEQUENCE (live-reported bug quick-260731-seq): `resource.sequence` is
 * ALWAYS set — a real, documented field on the Calendar API v3 Event
 * resource — from `event.sequence` when it's a real number (the normal
 * case, coming from parseVeventBlock, which always produces one), or
 * defensively defaulting to `0` when `event.sequence` is missing/undefined
 * (e.g. a hand-built event object not produced by parseVeventBlock).
 * Applies identically to both the Events.import path (where a MISSING
 * sequence used to be silently treated as `0` by Google, causing a real
 * live "Invalid sequence value" rejection whenever Gmail's own native
 * detection had already stored a higher real sequence number from the
 * same invite) and the Events.insert fallback path (harmless there too,
 * since there is no existing event to conflict with on a plain insert).
 *
 * Pure, no GAS globals — references neither CalendarApp nor Calendar so it
 * stays Node-testable like the rest of the parser.
 */
function buildEventResource(event) {
  const resource = {
    summary: event.summary,
    description: event.description,
    location: event.location,
    sequence: typeof event.sequence === 'number' ? event.sequence : 0,
  };

  if (event.uid) {
    resource.iCalUID = event.uid;
  }

  if (event.isAllDay) {
    resource.start = { date: event.start.toISOString().slice(0, 10) };
    resource.end = { date: event.end.toISOString().slice(0, 10) };
  } else {
    resource.start = { dateTime: event.start.toISOString() };
    resource.end = { dateTime: event.end.toISOString() };
  }

  if (event.recurrence) {
    resource.recurrence = [buildRRuleString(event.recurrence)];
  }

  return resource;
}

/**
 * extractEmailAddress — extracts the bare, trimmed, lowercased email address
 * from a Gmail "From" header value (e.g. 'Jana Nováková <jana@example.com>') or
 * from a bare address with no display name (e.g. 'jana@example.com'). Pure, no
 * GAS globals. Never throws: a null/undefined/empty input returns ''.
 */
function extractEmailAddress(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const angleBracketMatch = /<([^>]*)>/.exec(fromHeader);
  const raw = angleBracketMatch ? angleBracketMatch[1] : fromHeader;

  return raw.trim().toLowerCase();
}

/**
 * isAllowedSender — returns true when `importOnlyFrom` is null, undefined,
 * or has zero length (the import-all default that preserves current
 * behavior); otherwise returns true only if the sender extracted from
 * `fromHeader` strictly equals the sender extracted from at least one entry
 * in `importOnlyFrom` (each list entry is run through extractEmailAddress
 * too, so a bare address or a full 'Name <email>' list entry both match).
 * Matching is case-insensitive because extractEmailAddress lowercases both
 * sides. Pure, no GAS globals.
 */
function isAllowedSender(fromHeader, importOnlyFrom) {
  if (!importOnlyFrom || importOnlyFrom.length === 0) {
    return true;
  }

  const sender = extractEmailAddress(fromHeader);

  return importOnlyFrom.some(function (entry) {
    return extractEmailAddress(entry) === sender;
  });
}

/**
 * isExcludedSender — quick-260803-us3 (D-03): the inverse of isAllowedSender
 * above, same shape and comparison convention exactly. Returns false when
 * `excludeFrom` is null, undefined, or has zero length (the "exclude
 * nobody" default that preserves current behavior); otherwise returns true
 * only if the sender extracted from `fromHeader` strictly equals the sender
 * extracted from at least one entry in `excludeFrom` (each list entry is
 * run through extractEmailAddress too, so a bare address or a full
 * 'Name <email>' list entry both match). Matching is case-insensitive
 * because extractEmailAddress lowercases both sides. Independent of
 * isAllowedSender — the two gates never affect each other; a sender can
 * appear in neither, either, or (in a pathological config) both lists.
 * Pure, no GAS globals.
 */
function isExcludedSender(fromHeader, excludeFrom) {
  if (!excludeFrom || excludeFrom.length === 0) {
    return false;
  }

  const sender = extractEmailAddress(fromHeader);

  return excludeFrom.some(function (entry) {
    return extractEmailAddress(entry) === sender;
  });
}

/**
 * resolveIcsCalendarId — resolves which calendar ID a given ICS-carrying
 * message's event(s) should be written into. Three-tier resolution,
 * most-specific wins:
 *   1. icsConfig.calendarIdBySender: an array of { from, calendarId }
 *      entries; if fromHeader's extracted sender case-insensitively
 *      matches an entry's `from` (via extractEmailAddress on both sides —
 *      the same comparison convention as isAllowedSender above), that
 *      entry's calendarId wins. List order, FIRST match wins (same "list,
 *      first match wins" convention used elsewhere in this codebase, e.g.
 *      extractLabeledNumber in the booking action).
 *   2. icsConfig.calendarId: a single action-wide override, used when no
 *      per-sender mapping matched (or none is configured).
 *   3. defaultCalendarId: the global CONFIG.calendarId fallback, used when
 *      neither of the above is set/matched.
 *
 * INDEPENDENT of importOnlyFrom (a separate allow-list gate for whether a
 * message is processed at all) — a sender can be allowed to import without
 * having a calendar mapping (falls through to the default); this function
 * is only ever consulted for a message that already passed that gate.
 * Genuinely PER-MESSAGE, not per-thread or per-run: if a thread carries
 * .ics attachments from two different senders, each is resolved
 * independently and can route to a different calendar within the same run
 * (see ICS_CALENDAR_ACTION.run / getIcsAttachmentTextsByMessage). Pure, no
 * GAS globals.
 */
function resolveIcsCalendarId(fromHeader, icsConfig, defaultCalendarId) {
  const mapping = (icsConfig && icsConfig.calendarIdBySender) || [];
  const sender = extractEmailAddress(fromHeader);

  for (let i = 0; i < mapping.length; i++) {
    if (extractEmailAddress(mapping[i].from) === sender) {
      return mapping[i].calendarId;
    }
  }

  if (icsConfig && icsConfig.calendarId) {
    return icsConfig.calendarId;
  }

  return defaultCalendarId;
}

// Node/GAS environment bridge for ICS_ACTION_CONFIG (now defined in the
// sibling src/05-action-cfg-ics-import.js — see this file's class-level
// "CONFIG SPLIT" note). Under GAS's shared global scope, the sibling
// file's top-level `const ICS_ACTION_CONFIG` is ALREADY visible here by
// bare name — no action needed, and this `if` block never executes there
// (`typeof module` is always `'undefined'` under GAS). Under Node, each
// `require()`d file is its own isolated module with its own scope, so the
// bare `ICS_ACTION_CONFIG` reference inside ICS_CALENDAR_ACTION's `config`
// getter below would otherwise throw ReferenceError. Assigning to
// `globalThis.ICS_ACTION_CONFIG` (rather than redeclaring a conflicting
// top-level `const`/`let`/`var` of the same name, which GAS's own
// concatenated scope would reject as a duplicate declaration) makes the
// bare identifier resolve correctly under Node too, via ordinary
// global-object fallback identifier resolution.
if (typeof module !== 'undefined' && module.exports) {
  globalThis.ICS_ACTION_CONFIG = require('./05-action-cfg-ics-import.js').ICS_ACTION_CONFIG;
}

/**
 * ICS_CALENDAR_ACTION — the ICS-to-Calendar import action. Carries its own
 * config block, independent of CONFIG and of any other action's config.
 *
 * Detects `.ics` attachments, parses every VEVENT (via parseIcs), and
 * imports one matching calendar event per VEVENT into a RESOLVED calendar
 * (see MULTI-CALENDAR ROUTING below) using parse-then-create ordering:
 * every attachment across the whole thread is parsed FIRST; only after a
 * clean parse do we begin writing ANY event, so a malformed .ics throws
 * before any calendar write (fail closed) — this still holds even though
 * writes may now target more than one calendar within a single run.
 *
 * SENDER ALLOW-LIST: config.importOnlyFrom (see below) narrows which
 * senders' .ics attachments are imported via findIcsAttachments' single
 * shared filter point (isAllowedSender). Empty (the default) imports from
 * any sender, matching original behavior exactly.
 *
 * MULTI-CALENDAR ROUTING (owner-requested): CONFIG.calendarId
 * (src/01-setup.js) is now a DEFAULT/fallback, not the sole target.
 * resolveIcsCalendarId (see its own JSDoc above) resolves, PER MESSAGE, a
 * 3-tier priority: a config.calendarIdBySender entry matching that
 * message's sender > config.calendarId (an action-wide override) >
 * CONFIG.calendarId (the global default). run() groups attachment text by
 * ORIGINATING MESSAGE (getIcsAttachmentTextsByMessage) specifically so
 * this resolution can happen once per message and be reused for every
 * event that message's attachment(s) produce — a thread carrying .ics
 * files from two different senders can therefore route to two different
 * calendars within the same run. With every override left at its shipped
 * default (null / []), every import targets CONFIG.calendarId exactly as
 * before this feature existed.
 *
 * DEDUP FIX (iCalUID): events are written via the Advanced Calendar
 * Service's Calendar.Events.import (not a simple CalendarApp create), so
 * the .ics UID becomes the calendar event's iCalUID identity key.
 * Google Calendar's own backend treats iCalUID as canonical: whichever
 * path — this script's Events.import, or the owner's native Gmail
 * "Yes"/RSVP-accept click on the same invite — reaches a given UID FIRST
 * creates the event; the other subsequently UPDATES that same event
 * instead of creating a duplicate. This holds in both temporal orders
 * (script-then-click and click-then-script), which is what resolves the
 * original duplicate-event bug. A UID-less VEVENT (rare, non-conformant
 * .ics) has no identity key to dedup against, so it falls back to a plain
 * Calendar.Events.insert — an ordinary create with no dedup guarantee,
 * scoped to that narrow edge case only.
 *
 * KNOWN LIMITATION: parse-then-create only protects against parse errors,
 * not partial-write errors. The Events.import/insert calls happen one at a
 * time in a loop; if event N of a multi-VEVENT .ics throws during the
 * write (e.g. a Calendar API quota/permission error), events 1..N-1 are
 * already committed to the calendar while the action as a whole is
 * recorded as failed. Since failed threads are excluded from future
 * search, this is low-frequency, but if the owner manually reprocesses the
 * thread (e.g. removes the failed label), the earlier events would be
 * re-imported — harmless for UID-bearing events (import is idempotent by
 * iCalUID) but would duplicate any UID-less event that used the insert
 * fallback. No compensating cleanup is performed today.
 *
 * KNOWN LIMITATION (sender allow-list): isAllowedSender/findIcsAttachments
 * check the Gmail "From" header exactly as GmailApp reports it — there is
 * no SPF/DKIM/DMARC verification of the sender in-script. This is a
 * documented, accepted limitation: it is a convenience filter, not a
 * security boundary, since it only narrows what already-Gmail-delivered
 * (already-spam-filtered) mail gets processed and widens no trust boundary.
 *
 * SENDER EXCLUDE-LIST (quick-260803-us3, D-03): config.excludeFrom (see
 * ICS_ACTION_CONFIG in the sibling src/05-action-cfg-ics-import.js) is
 * checked via isExcludedSender at the SAME two filter points as the
 * importOnlyFrom allow-list above (findIcsAttachments,
 * getIcsAttachmentTextsByMessage) — a sender listed there is skipped
 * entirely, silently, for this action. This is the hand-off switch that
 * lets TRANSPORT_TICKETS_ACTION (src/08-action-transport-tickets.js) own a
 * sender (e.g. jizdenky@regiojet.cz) whose confirmation email ALSO carries a
 * .ics attachment this action would otherwise import too, producing two
 * competing calendar events for the same email. Empty (the default)
 * excludes nobody, so behavior is unchanged for every existing user. The
 * owner sets the real value out-of-band via Script Properties — never
 * hardcoded into this action's shipped default.
 *

 * SEQUENCE-CONFLICT RECOVERY (live-reported bug quick-260731-seq): even
 * with buildEventResource now always setting a real `sequence` (see its
 * own doc comment), `Calendar.Events.import` can still throw
 * `GoogleJsonResponseException: Invalid sequence value...` in a narrower
 * race: Gmail's native Gmail-to-Calendar detection may have ALREADY
 * updated the same event to a HIGHER sequence number than the one our
 * own parsed .ics carries, between when the invite arrived and when this
 * script processes it. `importIcsEventWithSequenceRetry` (below) handles
 * exactly this — a bounded, single-shot recovery that implements Google's
 * OWN documented remediation instruction embedded in the error message
 * itself ("Re-fetch the resource and use its sequence number on the
 * following request"), not speculative retry logic: on that specific
 * error, look up the existing event by iCalUID, copy ITS sequence number
 * onto our resource, and retry Events.import exactly once. If the retry
 * also throws, or no existing event is found by iCalUID, the error
 * propagates normally — dispatch isolation (03-action-management.js)
 * already contains an action's throw, routes the thread to the failed
 * label, and (config.notifyOnFailure) notifies the owner, so no separate
 * handling is needed here for the give-up path.
 *
 * RSVP PRESERVATION (live-reported bug ics-import-strips-rsvp): the iCalUID
 * dedup behavior described above has one destructive edge. When Gmail's own
 * native detection reaches the UID FIRST — the normal order for a real
 * Exchange/Teams invite, since Gmail acts on delivery while this script runs
 * on a periodic trigger minutes later — the event Gmail creates is a genuine
 * ATTENDEE COPY: real organizer, the owner as a NEEDS-ACTION attendee,
 * Accept/Decline UI, responses routed back to the organizer. Our subsequent
 * Events.import is a full-resource replace carrying NEITHER organizer NOR
 * attendees (buildEventResource's T-03-05 firewall), which silently demoted
 * that attendee copy to a plain self-owned private copy: guest list cleared,
 * organizer reset to the owner, RSVP gone, organizer never notified.
 * importIcsEventWithSequenceRetry now guards this by looking the UID up
 * BEFORE writing and skipping the write entirely when the event already
 * there carries guests — see its own "PRESERVE-EXISTING-INVITE GUARD" doc
 * paragraph for the full rationale, including why skipping is correct rather
 * than trying to reconstruct the attendee copy.
 */
const ICS_CALENDAR_ACTION = {
  name: 'ics-calendar-import',

  // GETTER, not a plain literal property — see this file's class-level
  // "CONFIG SPLIT" note. Not evaluated at object-construction time, only
  // when something reads `.config`, which happens lazily inside function
  // bodies (dispatchActions, notifyOwnerOfFailure) long after every
  // project file has loaded — so this is safe regardless of whether
  // 05-action-cfg-ics-import.js or this file loads first alphabetically.
  get config() {
    return ICS_ACTION_CONFIG;
  },

  /**
   * appliesTo — returns a literal boolean. True when any message on the
   * thread carries an attachment whose name ends in .ics (case-insensitive)
   * or whose content-type is text/calendar, AND whose sender passes the
   * config.importOnlyFrom allow-list (see findIcsAttachments); false
   * otherwise. dispatchActions only skips on a strict `=== false`, so a
   * literal boolean is required.
   */
  appliesTo: function (thread) {
    return findIcsAttachments(thread).length > 0;
  },

  /**
   * run — groups every matching .ics attachment on the thread BY
   * ORIGINATING MESSAGE (getIcsAttachmentTextsByMessage; a thread may
   * carry more than one — an original invite plus an update, or a single
   * message with more than one .ics file — and every one of them must be
   * processed, not just the first), parses every group's attachments
   * (parse-then-create: EVERY attachment across the WHOLE thread is parsed
   * first; only after all parse cleanly does any calendar write begin, so
   * a malformed .ics throws before any write call — fail closed, still
   * true even though writes may target more than one calendar), then for
   * each message's group: resolves that message's calendar ONCE
   * (resolveIcsCalendarId — see MULTI-CALENDAR ROUTING above) and writes
   * one calendar event per parsed VEVENT into it via buildEventResource +
   * the Advanced Calendar Service: Calendar.Events.import (via
   * importIcsEventWithSequenceRetry, which adds a bounded single-shot
   * sequence-conflict recovery — see its own doc comment and the
   * class-level "SEQUENCE-CONFLICT RECOVERY" paragraph above) when the
   * event carries a UID (idempotent by iCalUID — the dedup fix),
   * Calendar.Events.insert otherwise (UID-less fallback, no dedup
   * guarantee). Throws a clear error if a resolved calendar cannot be
   * found — dispatch isolation (03-action-management.js) contains the
   * throw, routes the thread to the failed label, and
   * (config.notifyOnFailure) notifies the owner.
   */
  run: function (thread) {
    const messageGroups = getIcsAttachmentTextsByMessage(thread);

    // Parse EVERY group's attachments FIRST, across the whole thread,
    // before any calendar write begins anywhere (fail closed) — even
    // though different groups may end up targeting different calendars.
    const parsedGroups = messageGroups.map(function (group) {
      const events = group.texts.reduce(function (allEvents, attachmentText) {
        return allEvents.concat(parseIcs(attachmentText));
      }, []);
      return { fromHeader: group.fromHeader, events: events };
    });

    parsedGroups.forEach(function (group) {
      const calendarId = resolveIcsCalendarId(group.fromHeader, ICS_CALENDAR_ACTION.config, CONFIG.calendarId);
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) {
        throw new Error('Calendar not found for resolved calendarId: ' + calendarId);
      }

      group.events.forEach(function (event) {
        const resource = buildEventResource(event);

        if (event.uid) {
          // Idempotent by iCalUID — the actual dedup fix. See
          // importIcsEventWithSequenceRetry's own doc comment for the
          // bounded sequence-conflict recovery wrapped around this call.
          importIcsEventWithSequenceRetry(resource, calendarId, event.uid);
        } else {
          // No identity key to dedup against; ordinary create.
          Calendar.Events.insert(resource, calendarId);
        }
      });
    });
  },
};

/**
 * hasGuestRelationship — true when `existingEvent` is a real Calendar API
 * event resource that carries at least one entry in its `attendees` array,
 * i.e. a genuine guest/RSVP relationship (an invite the owner was invited
 * to, or a meeting the owner organized and invited others to). False for
 * null/undefined (no such event), for an event with no `attendees` key at
 * all, and for an event whose `attendees` array is empty.
 *
 * This is deliberately the SIMPLEST possible discriminator, and it is the
 * exact discriminator the bug calls for: `attendees` is precisely the field
 * whose loss destroys the Accept/Decline UI and the response path back to
 * the organizer. Notably it is correct in BOTH directions of ownership —
 * an event the owner organized has its invitees in the same array, and
 * blowing THOSE away would be just as destructive, so it is guarded too.
 *
 * Events this script itself created (via buildEventResource, which never
 * emits attendees — the T-03-05 firewall) always have an empty/absent
 * attendees array, so they return false and remain freely re-importable.
 * The same is true of METHOD:PUBLISH informational .ics events (booking
 * confirmations, transport tickets), which carry no ATTENDEE properties.
 * That is what keeps this guard inert for every pre-existing flow.
 *
 * Pure, no GAS globals — Node-testable.
 */
function hasGuestRelationship(existingEvent) {
  return !!(existingEvent && existingEvent.attendees && existingEvent.attendees.length > 0);
}

/**
 * findExistingEventByICalUid — looks up the event already stored on
 * `calendarId` under the iCalendar UID `uid`, returning the first item or
 * null when none exists. An iCalUID is expected to identify at most one
 * event per calendar.
 *
 * `singleEvents: false` (the API default, stated explicitly here for the
 * reader) so a RECURRING event resolves to its single master event rather
 * than being expanded into one item per instance — the master is what
 * carries the authoritative `attendees` array, and expanding a long series
 * into instances just to answer a yes/no question would be wasteful. This
 * differs deliberately from the `singleEvents: true` lookup inside
 * importIcsEventWithSequenceRetry's sequence-conflict recovery below, which
 * is live-proven in production for a different question (reading a
 * conflicting `.sequence`) and is left exactly as it is.
 *
 * Cancelled/deleted events are excluded because `showDeleted` defaults to
 * false — deliberate: an invite the owner already declined and deleted must
 * NOT block a fresh import of the same UID.
 *
 * Never throws on a malformed/empty API response (defensive `&&` chain), so
 * a surprising response shape degrades to "no existing event" — i.e. to the
 * previous, pre-guard behavior — rather than breaking the import path.
 *
 * GAS-only (Calendar global), same category as the rest of this file's
 * Calendar-API-touching code; exercised in tests via a faked Calendar
 * global (the harness convention established in test/calendar-routing.test.js).
 */
function findExistingEventByICalUid(calendarId, uid) {
  const response = Calendar.Events.list(calendarId, { iCalUID: uid, singleEvents: false });

  return response && response.items && response.items.length > 0 ? response.items[0] : null;
}

/**
 * importIcsEventWithSequenceRetry — FIRST applies the preserve-existing-invite
 * guard (see below), and only if that does not fire calls
 * `Calendar.Events.import(resource, calendarId)`. On failure, inspects the
 * thrown error's message: if it does NOT contain the substring "Invalid
 * sequence value" (Google's own fixed error text for this specific
 * failure — see ICS_CALENDAR_ACTION's class-level "SEQUENCE-CONFLICT
 * RECOVERY" doc paragraph for the full rationale), the error is
 * re-thrown immediately, unchanged.
 *
 * When the error IS a sequence conflict, performs ONE bounded recovery
 * attempt — implementing Google's own documented remediation instruction
 * embedded in the error message itself ("Re-fetch the resource and use
 * its sequence number on the following request"), not speculative
 * error-handling:
 *   1. Look up the existing event by iCalUID via
 *      `Calendar.Events.list(calendarId, { iCalUID, singleEvents: true })`,
 *      taking the first item if any (an iCalUID is expected to identify
 *      at most one event per calendar).
 *   2. If found, set `resource.sequence` to THAT existing event's own
 *      `.sequence` value (matching what the API is asking for — not
 *      incrementing further, just matching) and retry
 *      `Calendar.Events.import(resource, calendarId)` EXACTLY ONCE MORE.
 *   3. If the retry ALSO throws, or no existing event is found by the
 *      iCalUID lookup, the error propagates normally (this function does
 *      NOT catch that second attempt, and does not catch the "no existing
 *      event found" case at all — both fall through to the caller
 *      unchanged). This is a single-shot recovery, never a retry loop.
 *
 * PRESERVE-EXISTING-INVITE GUARD (live-reported bug ics-import-strips-rsvp):
 * before ANY write, findExistingEventByICalUid looks the UID up on the target
 * calendar; when hasGuestRelationship says the event already there carries
 * guests, this function logs one line and returns WITHOUT writing.
 *
 * WHY: `Calendar.Events.import` is a full-resource, non-patch upsert keyed by
 * iCalUID, and it is the one Calendar API operation where `organizer` is
 * writable. buildEventResource deliberately emits NEITHER `organizer` NOR
 * `attendees` (the T-03-05 firewall — untrusted .ics ATTENDEE values must
 * never become real guests; parseVeventBlock does not even carry them that
 * far, folding them into description text instead). So when Gmail's own
 * native invite detection had ALREADY created the event for that UID — with
 * the real organizer and the owner as a NEEDS-ACTION attendee — our import
 * rewrote that attendee copy as a plain self-owned private copy: guest list
 * cleared, organizer reset to the calendar owner, Accept/Decline gone, and
 * the organizer never notified. Google's own guidance is explicit that an
 * attendee's copy must "specify the organizer in the attendee's copy"; an
 * organizer-less import therefore cannot BE an attendee copy.
 *
 * WHY SKIP RATHER THAN RECONSTRUCT: reconstructing the attendee copy would
 * mean sending `organizer`/`attendees` on the import — reopening exactly the
 * T-03-05 hole, and depending on unverified assumptions about whether a
 * rebuilt private copy still round-trips RSVP to a foreign organizer. Gmail's
 * native event is strictly higher fidelity than anything this script can
 * build (it is a genuine attendee copy, and Gmail keeps it current as the
 * organizer sends updates), so the correct action is to leave it alone. The
 * action's goal — "the invite is on the calendar" — is already satisfied.
 *
 * SCOPE (why this is inert for every pre-existing flow): the guard fires ONLY
 * when an event with that exact iCalUID already exists on that exact calendar
 * AND carries at least one attendee. Events this script created itself never
 * have attendees; METHOD:PUBLISH informational .ics events (booking
 * confirmations, RegioJet/transport tickets) carry no ATTENDEE properties at
 * all; and a non-default calendarId route finds no event to collide with. In
 * every one of those cases the code below runs byte-for-byte as before, so
 * the iCalUID dedup guarantee (quick-260723-gmk) and the sequence-conflict
 * recovery (quick-260731-seq) are both fully preserved.
 *
 * RETURNS a small result object — `{ action: 'skipped-existing-invite' |
 * 'imported', eventId }` — so callers and tests can observe which branch
 * was taken. Both existing call sites (run below, and
 * processTransportTicketJob in src/08-action-transport-tickets.js) ignore the
 * return value, so this is backward compatible.
 *
 * GAS-only (Calendar global); the guard's two helpers are the exception —
 * hasGuestRelationship is pure and directly unit-tested, and the guard's
 * branching is unit-tested through a faked Calendar global (see
 * test/existing-invite-guard.test.js). The SEQUENCE parsing and
 * resource-building this recovery depends on IS fully unit-tested (see
 * parseVeventBlock/buildEventResource).
 *
 * `optionalArgs` (quick-260803-us3, EXTENSION): an optional trailing
 * argument, defaulting to `{}`, passed through UNCHANGED to BOTH
 * `Calendar.Events.import` calls (the initial attempt and the single-shot
 * retry). This function's own pre-existing 3-argument call site in `run`
 * below is UNCHANGED (it relies on the default). TRANSPORT_TICKETS_ACTION
 * (src/08-action-transport-tickets.js) is what actually needs this: it
 * reuses this SAME proven idempotent-by-iCalUID import path but also needs
 * to pass `{ supportsAttachments: true }` when the event carries a Drive
 * attachment — see that action's own JSDoc for the full rationale.
 */
function importIcsEventWithSequenceRetry(resource, calendarId, uid, optionalArgs) {
  const args = optionalArgs || {};

  // PRESERVE-EXISTING-INVITE GUARD (live-reported bug ics-import-strips-rsvp)
  // — see hasGuestRelationship / findExistingEventByICalUid above and this
  // function's own doc paragraph. Must run BEFORE the import call: the whole
  // point is that the import is what destroys the guest list, so the only
  // safe moment to check is while the existing event is still intact.
  const preExisting = findExistingEventByICalUid(calendarId, uid);
  if (hasGuestRelationship(preExisting)) {
    console.log(
      'ICS import: an event for iCalUID ' + uid + ' already exists on calendar ' + calendarId +
        ' with ' + preExisting.attendees.length + ' guest(s) — leaving it untouched to preserve its RSVP state (skipping import).'
    );
    return { action: 'skipped-existing-invite', eventId: preExisting.id };
  }

  try {
    Calendar.Events.import(resource, calendarId, args);
    return { action: 'imported', eventId: null };
  } catch (e) {
    const message = (e && e.message) || String(e);
    if (message.indexOf('Invalid sequence value') === -1) {
      throw e;
    }

    const existing = Calendar.Events.list(calendarId, { iCalUID: uid, singleEvents: true });
    const existingEvent = existing && existing.items && existing.items.length > 0 ? existing.items[0] : null;

    if (!existingEvent) {
      throw e;
    }

    resource.sequence = existingEvent.sequence;
    Calendar.Events.import(resource, calendarId, args);
    return { action: 'imported', eventId: existingEvent.id || null };
  }
}

/**
 * isIcsAttachment — true when `attachment`'s name ends in .ics
 * (case-insensitive) or its content-type is text/calendar. Shared by
 * findIcsAttachments and getIcsAttachmentTextsByMessage so the matching
 * rule lives in exactly one place.
 */
function isIcsAttachment(attachment) {
  const name = (attachment.getName() || '').toLowerCase();
  const contentType = attachment.getContentType() || '';

  return name.slice(-4) === '.ics' || contentType === 'text/calendar';
}

/**
 * findIcsAttachments — returns every GmailAttachment across every message on
 * the thread whose name ends in .ics (case-insensitive) or whose
 * content-type is text/calendar (via isIcsAttachment), in source order, or
 * [] if none match. Used by appliesTo (a flat count is all it needs);
 * getIcsAttachmentTextsByMessage does its own equivalent per-message walk
 * so it can preserve which message each attachment came from.
 *
 * SENDER ALLOW-LIST: before scanning a message's attachments, its "From"
 * header is checked against config.importOnlyFrom via isAllowedSender — the
 * single shared filter point. A message from a disallowed sender is skipped
 * entirely (its attachments never enter `matches`), so a thread whose only
 * .ics is from a disallowed sender yields [] here, making appliesTo return
 * false and the action a silent no-op for that thread (no error, no failed
 * label). With the default empty importOnlyFrom, every message is allowed,
 * so behavior is unchanged from before this filter existed.
 *
 * SENDER EXCLUDE-LIST (quick-260803-us3, D-03): a message's "From" header is
 * ALSO checked against config.excludeFrom via isExcludedSender — an excluded
 * sender's message is skipped the same way a disallowed sender's is. This is
 * the hand-off switch that lets another action (e.g. TRANSPORT_TICKETS_ACTION)
 * own a sender whose .ics attachments this action would otherwise also
 * claim. With the default empty excludeFrom, no sender is excluded, so
 * behavior is unchanged from before this gate existed.
 */
function findIcsAttachments(thread) {
  const messages = thread.getMessages();
  const matches = [];

  for (let i = 0; i < messages.length; i++) {
    if (!isAllowedSender(messages[i].getFrom(), ICS_CALENDAR_ACTION.config.importOnlyFrom)) {
      continue;
    }
    if (isExcludedSender(messages[i].getFrom(), ICS_CALENDAR_ACTION.config.excludeFrom)) {
      continue;
    }

    const attachments = messages[i].getAttachments();

    for (let j = 0; j < attachments.length; j++) {
      if (isIcsAttachment(attachments[j])) {
        matches.push(attachments[j]);
      }
    }
  }

  return matches;
}

/**
 * getIcsAttachmentTextsByMessage — returns one { fromHeader, texts } group
 * per thread message that (a) passes the sender allow-list
 * (isAllowedSender against config.importOnlyFrom, same shared filter point
 * as findIcsAttachments) and (b) has at least one .ics-matching attachment
 * (isIcsAttachment); a message with an allowed sender but no matching
 * attachment contributes no group at all (not an empty one). `texts` holds
 * that message's matching attachments' raw text, in source order — every
 * match is returned, not just the first, so a message with multiple .ics
 * files does not silently drop anything past the first.
 *
 * MULTI-CALENDAR ROUTING: grouping BY MESSAGE (rather than a single flat
 * list across the whole thread, as the prior getIcsAttachmentTexts did) is
 * what lets run() resolve a calendar ONCE PER MESSAGE via
 * resolveIcsCalendarId and reuse it for every event that message's
 * attachment(s) produce — see ICS_CALENDAR_ACTION's class-level JSDoc.
 *
 * Throws if the thread yields zero groups (should not happen when
 * appliesTo returned true, but guards against a future call-order bug).
 */
function getIcsAttachmentTextsByMessage(thread) {
  const messages = thread.getMessages();
  const groups = [];

  for (let i = 0; i < messages.length; i++) {
    const fromHeader = messages[i].getFrom();
    if (!isAllowedSender(fromHeader, ICS_CALENDAR_ACTION.config.importOnlyFrom)) {
      continue;
    }
    if (isExcludedSender(fromHeader, ICS_CALENDAR_ACTION.config.excludeFrom)) {
      continue;
    }

    const attachments = messages[i].getAttachments();
    const texts = [];

    for (let j = 0; j < attachments.length; j++) {
      if (isIcsAttachment(attachments[j])) {
        texts.push(attachments[j].getDataAsString());
      }
    }

    if (texts.length > 0) {
      groups.push({ fromHeader: fromHeader, texts: texts });
    }
  }

  if (groups.length === 0) {
    throw new Error('No .ics attachment found on thread: ' + thread.getId());
  }

  return groups;
}

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Single merged export carries parseIcs (parser tests), the two pure
// resource builders (buildRRuleString, buildEventResource), the two pure
// sender-allow-list helpers (extractEmailAddress, isAllowedSender), the
// pure multi-calendar-routing resolver (resolveIcsCalendarId), the four
// pure organizer/attendee/formatting enrichment helpers
// (parseAddressProperty, formatAddress, buildOrganizerAttendeesText,
// collapseBlankLines), the preserve-existing-invite guard's two pieces
// (hasGuestRelationship — pure; findExistingEventByICalUid and
// importIcsEventWithSequenceRetry — GAS-only, exported so the guard's
// branching can be exercised against a faked Calendar global), and
// ICS_CALENDAR_ACTION (action registry).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    hasGuestRelationship: hasGuestRelationship,
    findExistingEventByICalUid: findExistingEventByICalUid,
    importIcsEventWithSequenceRetry: importIcsEventWithSequenceRetry,
    parseIcs: parseIcs,
    buildRRuleString: buildRRuleString,
    buildEventResource: buildEventResource,
    extractEmailAddress: extractEmailAddress,
    isAllowedSender: isAllowedSender,
    isExcludedSender: isExcludedSender,
    isIcsAttachment: isIcsAttachment,
    resolveIcsCalendarId: resolveIcsCalendarId,
    parseAddressProperty: parseAddressProperty,
    formatAddress: formatAddress,
    buildOrganizerAttendeesText: buildOrganizerAttendeesText,
    collapseBlankLines: collapseBlankLines,
    ICS_CALENDAR_ACTION: ICS_CALENDAR_ACTION,
  };
}

/**
 * TRANSPORT_TICKETS_ACTION — detects train/bus ticket confirmation emails
 * (RegioJet first, extensible to other carriers via
 * TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders) and creates ONE
 * calendar event per email/ticket.
 *
 * UNLIKE every other action in this codebase, this action's event data
 * comes from the email's OWN `text/calendar` `.ics` attachment, parsed with
 * the EXISTING iCalendar parser in src/05-action-ics-import.js (parseIcs /
 * buildEventResource). There is DELIBERATELY no second VEVENT parser in
 * this project: hand-rolling one here would duplicate the VTIMEZONE
 * resolution logic (extractVtimezoneBlocks / resolveTzidDate) that file
 * already gets right.
 *
 * HAND-OFF FROM ICS_CALENDAR_ACTION: ICS_CALENDAR_ACTION already imports
 * `.ics` attachments from ANY sender by default, so without intervention it
 * would ALSO claim a RegioJet confirmation email, producing a second,
 * competing calendar event. A sender owned by this action is therefore
 * expected to ALSO be listed in ICS_ACTION_CONFIG.excludeFrom
 * (src/05-action-cfg-ics-import.js) so ICS_CALENDAR_ACTION stands down for
 * it. This is an OWNER-SIDE Script Properties step, not a code default —
 * this action's own detection (the presence of a `.ics` attachment from a
 * configured transport sender) is independent of ICS_ACTION_CONFIG
 * entirely: that is a separate action's config and must never gate this
 * one.
 *
 * ORDERING GUARANTEE: no Drive file is EVER created before the dedup
 * decision is made (see processTransportTicketJob's step order below) —
 * this action never uploads to a TEMP folder at all; the ticket PDF, when
 * insertPdfIntoEvent is true, is copied DIRECTLY into the PERMANENT
 * CONFIG.ticketAttachmentDriveFolderName folder only AFTER the dedup check
 * has already decided this run should proceed. This does NOT mean the
 * permanent file can never end up unattached to any event: if a Calendar
 * write later in the same run throws (quota/permission error, or the
 * FLAGGED ASSUMPTION below), the already-uploaded/renamed permanent file is
 * left with no reconciling try/finally, and no compensating cleanup is
 * performed today.
 *
 * DEDUP: the created event is tagged with
 * `extendedProperties.private.ticketIdentifier`, pre-checked via
 * `findTransportEventByIdentifier` before any write. The identifier's
 * PRIMARY source is the VEVENT SUMMARY's leading `#<digits>` (e.g.
 * `#7788123456`) — it matches the human-facing ticket number shown in the
 * email subject, unlike the opaque `UID` hash (a negative number for
 * RegioJet, e.g. `-9876543210@regiojet.cz`). `UID` is the fallback when no
 * `#<digits>` prefix is present; `null` (never throw) when neither is
 * present — a missing dedup key must never block calendar-event creation
 * (see extractTransportTicketIdentifier's own JSDoc).
 *
 * PDF ARCHIVE + ATTACH: when a resolved sender's `insertPdfIntoEvent` is
 * true, the ticket PDF (found via findTransportTicketPdfAttachment, which
 * EXCLUDES an accompanying `invoice.pdf` — that is not the ticket) is
 * copied into `CONFIG.ticketAttachmentDriveFolderName` — the SAME permanent
 * Drive folder src/07-action-ticketing-portals.js already uses (a single
 * shared destination, never a second folder setting) — renamed via
 * buildTransportAttachmentFilename, and attached to the created event using
 * the EventAttachment shape `{fileId, fileUrl, title, mimeType}` plus
 * `supportsAttachments: true` — `fileUrl` is REQUIRED, `fileId` alone is
 * NOT sufficient.
 *
 * GLOBALLY-UNIQUE NAMING: every pure helper in this file is
 * `transport`-prefixed to avoid colliding with the same-purpose helpers
 * already declared in src/05-*, src/06-*, and src/07-* — GAS concatenates
 * every project file into ONE shared global scope, so two files declaring
 * the same top-level function name silently collide.
 *
 * TWO PROCESSING MODES: a second carrier, IDOS.cz (`jizdenky@idos.svt.cz`),
 * sends confirmation emails with NO `.ics` attachment at all — route,
 * dates, times, seats and both codes live only in the plain-text email
 * body. Rather than forcing every future carrier through the `.ics`-VEVENT
 * model above, `transportSenders` entries carry a `mode` field
 * (`'ics' | 'body'`), mirroring the split TICKETING_PORTALS_ACTION already
 * proves for its own two processing modes (src/07-action-ticketing-
 * portals.js, resolveTicketProcessingJobs / processTicketFromMessageBody):
 *   - `'ics'` (RegioJet): parseIcs/buildEventResource ->
 *     buildTransportIcsEntry.
 *   - `'body'` (IDOS.cz): message.getPlainBody() -> a carrier-specific text
 *     parser (parseIdosTicketText) -> buildTransportBodyEntry.
 * `resolveTransportSenderMode` decides which mode a sender resolves to;
 * both modes then feed the SAME shared entry shape
 * (`{ resource, summary, filenameDate, ticketIdentifier, uid }`) into the
 * SAME dedup -> PDF-archive -> write pipeline in processTransportTicketJob
 * — only the extraction step differs, never the downstream
 * write/dedup/attach logic.
 */

/**
 * transportExtractEmailAddress — extracts the bare, trimmed, lowercased
 * email address from a Gmail "From" header value, or from a bare address
 * with no display name. Locally reimplemented per this file's naming
 * convention. Pure, no GAS globals. Never throws: a null/undefined/empty
 * input returns ''.
 */
function transportExtractEmailAddress(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const angleBracketMatch = /<([^>]*)>/.exec(fromHeader);
  const raw = angleBracketMatch ? angleBracketMatch[1] : fromHeader;

  return raw.trim().toLowerCase();
}

/**
 * resolveTransportSender — finds the TRANSPORT_SENDERS config entry whose
 * `identifyingEmail` case-insensitively matches `fromHeader`'s sender,
 * mirroring resolveTicketingPortal's per-sender lookup convention
 * (src/07-action-ticketing-portals.js): list order, FIRST match wins; a
 * null/empty `senders` list, or no match, returns `null`, never throws.
 * Pure, no GAS globals.
 */
function resolveTransportSender(fromHeader, senders) {
  const list = senders || [];
  const sender = transportExtractEmailAddress(fromHeader);

  for (let i = 0; i < list.length; i++) {
    if (transportExtractEmailAddress(list[i].identifyingEmail) === sender) {
      return list[i];
    }
  }

  return null;
}

/**
 * resolveTransportCalendarId — resolves which calendar ID this action's
 * Calendar API calls should target for a given `sender` (a resolved
 * TRANSPORT_SENDERS entry): `sender.calendarId` when truthy, else
 * `defaultCalendarId`. A `calendarId: null` default read DIRECTLY with no
 * fallback produces `TypeError: Cannot read properties of null (reading
 * 'getTimeZone')` — this two-tier resolution exists to prevent that. Pure,
 * no GAS globals.
 */
function resolveTransportCalendarId(sender, defaultCalendarId) {
  return (sender && sender.calendarId) || defaultCalendarId;
}

/**
 * isTransportPdfAttachment — true when `attachment`'s name ends in .pdf
 * (case-insensitive) or its content-type is application/pdf. Shared by
 * findTransportTicketPdfAttachment so the matching rule lives in exactly
 * one place, mirroring isTicketPdfAttachment (src/07-action-ticketing-
 * portals.js) and isIcsAttachment (src/05-action-ics-import.js).
 */
function isTransportPdfAttachment(attachment) {
  const name = (attachment.getName() || '').toLowerCase();
  const contentType = attachment.getContentType() || '';

  return name.slice(-4) === '.pdf' || contentType === 'application/pdf';
}

/**
 * findTransportTicketPdfAttachment — returns the FIRST qualifying PDF
 * attachment (via isTransportPdfAttachment) on `message` whose lowercased
 * name does NOT start with `invoice`, or `null` if none qualify. The
 * exclusion is scoped NARROWLY to the filename actually observed on the
 * real RegioJet fixture (`eticket.pdf` is the real ticket, `invoice.pdf` is
 * a separate accompanying invoice, NOT the ticket) — per this codebase's
 * "don't guess at an unobserved variant" discipline. Pure, no GAS globals
 * (operates only on the array message.getAttachments() already produces).
 */
function findTransportTicketPdfAttachment(message) {
  const pdfAttachments = message.getAttachments().filter(isTransportPdfAttachment);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = (pdfAttachments[i].getName() || '').toLowerCase();
    if (name.indexOf('invoice') !== 0) {
      return pdfAttachments[i];
    }
  }

  return null;
}

/**
 * extractTransportTicketIdentifier — the DEDUP SAFETY NET's stable key for
 * a parsed event object (as returned by parseIcs, src/05-action-ics-
 * import.js). PRIMARY source: the trimmed `event.summary`'s leading
 * `#<digits>` prefix (e.g. `#7788123456` -> `'7788123456'`) — this is the
 * HUMAN-FACING ticket number that also appears in the email subject, unlike
 * `event.uid` (an opaque, often negative, hash e.g.
 * `-9876543210@regiojet.cz`). FALLBACK: the raw `event.uid` when no
 * `#<digits>` prefix is present. `null` (never throws) when NEITHER is
 * present — a missing dedup key must never block calendar-event creation.
 * Pure, no GAS globals.
 */
function extractTransportTicketIdentifier(event) {
  const summary = event && event.summary ? String(event.summary).trim() : '';
  const summaryMatch = /^#(\d+)/.exec(summary);
  if (summaryMatch) {
    return summaryMatch[1];
  }

  if (event && event.uid) {
    return event.uid;
  }

  return null;
}

/**
 * sanitizeTransportFilenameComponent — replaces filesystem-unsafe
 * characters (`/ \ ? % * : | " < >`) with `-` and trims whitespace. Locally
 * reimplemented per this file's naming convention. Pure, no GAS globals.
 */
function sanitizeTransportFilenameComponent(value) {
  return String(value)
    .replace(/[/\\?%*:|"<>]/g, '-')
    .trim();
}

/**
 * buildTransportAttachmentFilename — the project-wide
 * `"{name} - {YYYY-MM-DD} - {identifier}.pdf"` attachment-renaming
 * convention: a Calendar event attachment's displayed `title` is derived
 * from the file's name AT ATTACH TIME, so renaming the file before it is
 * referenced improves both the Drive folder's browsability AND the
 * calendar event's displayed attachment name. `summary` is the VEVENT
 * SUMMARY text (e.g. the real `#7788123456: Z Ostrava, hl.n., do Praha,
 * hl.n., sedadla: [2/15,2/16]`), sanitized via
 * sanitizeTransportFilenameComponent. `startDate` is a real Date (the
 * parsed event's `.start`); the ISO date segment is taken from its UTC
 * calendar date (a knowingly minor simplification affecting only the
 * filename). `ticketIdentifier` (see extractTransportTicketIdentifier) is
 * OMITTED entirely — never string-coerced — when falsy, so no filename
 * ever embeds the literal 4-character word "null". Pure, no GAS globals.
 */
function buildTransportAttachmentFilename(summary, startDate, ticketIdentifier) {
  const isoDate = startDate.toISOString().slice(0, 10);
  const ticketIdentifierSegment = ticketIdentifier ? ' - ' + ticketIdentifier : '';

  return sanitizeTransportFilenameComponent(summary) + ' - ' + isoDate + ticketIdentifierSegment + '.pdf';
}

/**
 * parseIdosTicketText — the IDOS.cz-specific email-BODY parser. UNLIKE the
 * rest of this file, this parses `message.getPlainBody()` directly, never a
 * `.ics` VEVENT — IDOS.cz confirmation emails carry no `.ics` attachment at
 * all.
 *
 * FORMAT (its own shape, needs its own regex): `D.M.YYYY H:MM` — day, month
 * AND hour with NO leading zeros, dot-separated with no spaces between date
 * components. This is DIFFERENT from every other date format already in
 * this codebase (RegioJet's ISO-8601 `.ics` `DTSTART`, enigoo.cz's
 * zero-padded `15.08.2026`, Kino Art's dot-space `7. 8. 2026 17:45`). The
 * route separator is `»` (U+00BB), also new here.
 *
 * PATTERN-ANCHORED, NOT LINE-POSITION: ONE combined trip-line regex
 * captures departure date/time, the from-station (non-greedy up to `»`),
 * the to-station (non-greedy up to the arrival date), and arrival date/time
 * — a SINGLE regex (not four separate ones) guarantees all six fields come
 * from the SAME trip line. `\s+`/`\s*` between components absorbs whichever
 * line-separator convention is in play (CRLF vs LF), and the anchors are
 * deliberately independent of the `- `/`* ` bullet marker Gmail's real
 * getPlainBody() rendering may prefix each detail line with — the trip
 * anchor never looks at line starts at all. The customer-support URL line
 * (which, in a real message, embeds the recipient's own personal mail
 * address as an `email=` query param) is never touched by any anchor here.
 *
 * SCOPE LIMITATION (deliberate, same "don't guess at an unobserved
 * variant" discipline as KINO_ART_KNOWN_VENUES and
 * findKinoArtTicketPdfAttachment, src/07-action-ticketing-portals.js):
 * every anchor here uses `.exec` (no `/g`), matching only the FIRST trip
 * line in the body, so an order confirmation bundling more than one
 * e-jízdenka (e.g. a round trip) would have every ticket after the first
 * silently unprocessed — no event, no error, no log line. The real IDOS.cz
 * email this parser was built from, and every fixture in this codebase,
 * show a single-ticket order only; if a genuine multi-ticket IDOS.cz
 * confirmation is ever observed, this needs generalizing THEN, with real
 * data, not guessed now.
 *
 * OPTIONAL anchors, each `null` when absent, NEVER throwing (a missing
 * dedup key or descriptive field must never block event creation): `seats`
 * (the run of characters after the literal `sedadlo` up to the next comma
 * or line end), `eTicketCode` (after the literal `kód e-jízdenky`), and
 * `ticketIdentifier` (the PURCHASE-SCOPED order code after the literal
 * `kód IDOS.cz` — a character class of uppercase letters, digits and
 * hyphens, which naturally stops at the trailing comma; deliberately NOT
 * the shorter, per-ticket `kód e-jízdenky` value).
 *
 * Returns `{ from, to, start: {year, month, day, hour, minute}, end: {…},
 * seats, eTicketCode, ticketIdentifier }` (month ZERO-INDEXED, matching
 * every other date-components object in this codebase). Throws a
 * controlled `Error` — with the COMPLETE raw text appended
 * (diagnostic-on-failure, same convention as parseEnigooTicketText/
 * parseKinoArtTicketText) — when the trip anchor does not match, or when
 * an hour is outside 0-23 or a minute outside 0-59. Pure, no GAS globals.
 */
function parseIdosTicketText(text) {
  const rawText = String(text || '');

  // Both station-name groups are bounded to 200 characters and exclude
  // newlines: an unbounded [\s\S]+? pair separated only by the literal "»"
  // is polynomial-time (~O(n^2)) against a crafted body, and
  // resolveTransportSender trusts an unauthenticated From header -- a
  // spoofed sender could otherwise stall a shared processEmails() run for
  // seconds on a single message. A real station name is at most a few
  // dozen characters, so this bound is never reached by legitimate input.
  const tripMatch =
    /(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})\s+([^»\r\n]{1,200}?)\s*»\s*([^\r\n]{1,200}?)\s+(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})/.exec(
      rawText
    );
  if (!tripMatch) {
    throw new Error('Unrecognized IDOS.cz ticket text: no trip line found. Full raw text:\n' + rawText);
  }

  const start = {
    day: Number(tripMatch[1]),
    month: Number(tripMatch[2]) - 1,
    year: Number(tripMatch[3]),
    hour: Number(tripMatch[4]),
    minute: Number(tripMatch[5]),
  };
  const from = tripMatch[6].trim();
  const to = tripMatch[7].trim();
  const end = {
    day: Number(tripMatch[8]),
    month: Number(tripMatch[9]) - 1,
    year: Number(tripMatch[10]),
    hour: Number(tripMatch[11]),
    minute: Number(tripMatch[12]),
  };

  if (start.hour < 0 || start.hour > 23) {
    throw new Error('Departure hour out of range (0-23) in IDOS.cz ticket trip line. Full raw text:\n' + rawText);
  }
  if (start.minute < 0 || start.minute > 59) {
    throw new Error('Departure minute out of range (0-59) in IDOS.cz ticket trip line. Full raw text:\n' + rawText);
  }
  if (end.hour < 0 || end.hour > 23) {
    throw new Error('Arrival hour out of range (0-23) in IDOS.cz ticket trip line. Full raw text:\n' + rawText);
  }
  if (end.minute < 0 || end.minute > 59) {
    throw new Error('Arrival minute out of range (0-59) in IDOS.cz ticket trip line. Full raw text:\n' + rawText);
  }

  const seatsMatch = /sedadlo\s+([^,\r\n]+)/.exec(rawText);
  const seats = seatsMatch ? seatsMatch[1].trim() : null;

  const eTicketCodeMatch = /kód e-jízdenky\s+([^,\r\n]+)/.exec(rawText);
  const eTicketCode = eTicketCodeMatch ? eTicketCodeMatch[1].trim() : null;

  // D-02: the purchase-scoped ORDER code, not the shorter per-ticket
  // "kód e-jízdenky" value above.
  const ticketIdentifierMatch = /kód IDOS\.cz\s+([A-Z0-9-]+)/.exec(rawText);
  const ticketIdentifier = ticketIdentifierMatch ? ticketIdentifierMatch[1] : null;

  return {
    from: from,
    to: to,
    start: start,
    end: end,
    seats: seats,
    eTicketCode: eTicketCode,
    ticketIdentifier: ticketIdentifier,
  };
}

/**
 * TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL — the local (single-file)
 * registry mapping a BODY-SOURCED transport sender's `identifyingEmail`
 * (lowercased, via transportExtractEmailAddress) to its email-body parser
 * function, mirroring TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL's exact
 * keying convention (src/07-action-ticketing-portals.js). Which registry a
 * matching sender's address resolves against (this one, vs. simply having
 * no entry here at all) is what resolveTransportSenderMode's
 * registry-fallback branch below consults.
 */
const TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL = {
  'jizdenky@idos.svt.cz': parseIdosTicketText,
};

/**
 * resolveTransportSenderMode — decides which processing mode (`'ics'` or
 * `'body'`) a resolved `sender` (a TRANSPORT_SENDERS entry, or
 * `null`/`undefined`) should be processed through:
 *   1. An explicit `sender.mode` wins outright (`'body'` or `'ics'`).
 *   2. `sender.mode` PRESENT but neither `'body'` nor `'ics'` (a typo, an
 *      unsupported mode, an empty string) THROWS rather than silently
 *      falling through to the registry/`'ics'` fallback below — treating
 *      an invalid value the same as an ABSENT one reproduces the exact
 *      silent-no-op class this function otherwise exists to prevent: a
 *      garbage `mode` on a future body-sourced carrier with no registered
 *      parser would silently resolve to `'ics'`, find zero `.ics`
 *      attachments, and the thread would be labeled processed with no
 *      error and no log line.
 *   3. No `mode` field at all: `'body'` when a parser IS registered for
 *      this sender's address in TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL
 *      — this registry-fallback branch exists SPECIFICALLY to prevent a
 *      body-sourced sender silently falling through the `.ics`-attachment
 *      requirement and contributing zero jobs.
 *   4. Otherwise (no `mode`, no registered parser) — `'ics'`. This is
 *      DELIBERATE back-compat: the owner's already-live Script Property
 *      JSON (written before this `mode` field existed) has no `mode` on
 *      its RegioJet entry and must keep resolving to `'ics'` unchanged.
 *   5. A null/undefined `sender` resolves to `'ics'` without throwing (a
 *      defensive default; resolveTransportProcessingJobs never actually
 *      calls this with a null sender, since it already skips unmatched
 *      messages first, but this function stays null-safe on its own
 *      terms).
 * Pure, no GAS globals.
 */
function resolveTransportSenderMode(sender) {
  if (!sender) {
    return 'ics';
  }

  if (sender.mode === 'body') {
    return 'body';
  }
  if (sender.mode === 'ics') {
    return 'ics';
  }
  if (sender.mode !== undefined && sender.mode !== null) {
    throw new Error(
      'Unrecognized transportSenders mode "' + sender.mode + '" for sender: ' + sender.identifyingEmail
    );
  }

  const senderKey = transportExtractEmailAddress(sender.identifyingEmail);
  if (TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey]) {
    return 'body';
  }

  return 'ics';
}

/**
 * zeroPadTransportComponent — left-pads `value` with '0' to `length`
 * digits. Pure, no GAS globals. Locally reimplemented per this file's
 * naming convention.
 */
function zeroPadTransportComponent(value, length) {
  return String(value).padStart(length, '0');
}

/**
 * formatTransportWallClockIso — formats a `{ year, month, day, hour,
 * minute }` wall-clock components object (month zero-indexed) as a
 * zero-padded literal string `'YYYY-MM-DDTHH:MM:00'` — DELIBERATELY with
 * NO trailing `Z` and NO timezone offset, meant to be paired with an
 * explicit Calendar API `timeZone` field so the API interprets these
 * digits as wall-clock local time in that zone, not UTC. Locally
 * reimplemented per this file's naming convention. Pure, no GAS globals.
 */
function formatTransportWallClockIso(components) {
  return (
    zeroPadTransportComponent(components.year, 4) +
    '-' +
    zeroPadTransportComponent(components.month + 1, 2) +
    '-' +
    zeroPadTransportComponent(components.day, 2) +
    'T' +
    zeroPadTransportComponent(components.hour, 2) +
    ':' +
    zeroPadTransportComponent(components.minute, 2) +
    ':00'
  );
}

/**
 * buildTransportBodyEntry — builds the SHARED entry shape
 * (`{ resource, summary, filenameDate, ticketIdentifier, uid, status,
 * dtstamp }`) for a `'body'`-mode job from `parsedTicket`
 * (parseIdosTicketText's return shape) and an INJECTED `timeZone` (never
 * read from a GAS global here — its GAS-only caller,
 * processTransportTicketJob, resolves it live via
 * `CalendarApp.getCalendarById(calendarId).getTimeZone()`, keeping this
 * function pure and unit-testable). The timezone is always resolved live
 * from the target calendar, never hardcoded. `resource.description` lists
 * the e-ticket code, the seats and the order code, each line OMITTED
 * entirely — never string-coerced — when its value is falsy, so no
 * description ever embeds the literal 4-character word "null".
 * `entry.filenameDate` is built with `Date.UTC(start.year, start.month,
 * start.day)` used purely as NEUTRAL arithmetic space (never a real
 * instant — the parsed components carry no timezone information) so only
 * the calendar date reaches buildTransportAttachmentFilename. `entry.uid`
 * is always `null` for this mode (IDOS.cz has no VEVENT UID at all), which
 * is what routes this entry through the shared write loop's
 * `Calendar.Events.insert` branch rather than the `.ics`-only idempotent
 * import path. `entry.status` is always `null` for this mode — IDOS.cz
 * body-mode tickets have no ICS STATUS concept and are out of scope for
 * cancellation via this mechanism. `entry.dtstamp` is likewise always
 * `null` — IDOS.cz has no DTSTAMP concept either, same treatment as
 * `status`. Pure, no GAS globals.
 */
function buildTransportBodyEntry(parsedTicket, timeZone) {
  const summary = parsedTicket.from + ' » ' + parsedTicket.to;

  const descriptionLines = [];
  if (parsedTicket.eTicketCode) {
    descriptionLines.push('E-ticket code: ' + parsedTicket.eTicketCode);
  }
  if (parsedTicket.seats) {
    descriptionLines.push('Seats: ' + parsedTicket.seats);
  }
  if (parsedTicket.ticketIdentifier) {
    descriptionLines.push('IDOS.cz order code: ' + parsedTicket.ticketIdentifier);
  }

  const resource = {
    summary: summary,
    location: parsedTicket.from,
    description: descriptionLines.join('\n'),
    start: { dateTime: formatTransportWallClockIso(parsedTicket.start), timeZone: timeZone },
    end: { dateTime: formatTransportWallClockIso(parsedTicket.end), timeZone: timeZone },
  };

  return {
    resource: resource,
    summary: summary,
    filenameDate: new Date(Date.UTC(parsedTicket.start.year, parsedTicket.start.month, parsedTicket.start.day)),
    ticketIdentifier: parsedTicket.ticketIdentifier,
    uid: null,
    status: null,
    dtstamp: null,
  };
}

/**
 * stripTransportSummaryIdentifierPrefix — RegioJet's real VEVENT SUMMARY
 * leads with the same `#<digits>: ` ticket-number prefix
 * extractTransportTicketIdentifier already extracts as the dedup key — e.g.
 * `'#4400574546: Z Vídeň, Hbf, do Brno, hl.n., ...'`. Showing it again in
 * the CALENDAR EVENT'S TITLE is redundant (the identifier is already
 * tracked via extendedProperties.private.ticketIdentifier), so this strips
 * it from the display summary. Only strips the exact observed shape (`#`,
 * one or more digits, `:`, optional whitespace) at the very start; anything
 * else is left untouched, including a summary with no such prefix at all.
 * Pure, no GAS globals.
 */
function stripTransportSummaryIdentifierPrefix(summary) {
  return String(summary || '').replace(/^#\d+:\s*/, '');
}

/**
 * buildTransportIcsEntry — the `'ics'`-mode counterpart to
 * buildTransportBodyEntry, expressing RegioJet's EXISTING behavior through
 * the SAME shared entry shape — no new parsing at all: `resource` is the
 * REUSED buildEventResource(event), and
 * `uid`/`ticketIdentifier`/`filenameDate`/`status`/`dtstamp` are read
 * straight off the already-parsed `event` object (via the EXISTING
 * extractTransportTicketIdentifier for the identifier). `status` and
 * `dtstamp` come STRAIGHT off `event.status`/`event.dtstamp` — no new
 * parsing here either, and `buildEventResource` still does not copy either
 * onto `resource`. `summary` (and `resource.summary`, which otherwise
 * carries the raw VEVENT SUMMARY through unmodified) has RegioJet's
 * redundant `#<digits>: ` prefix stripped via
 * stripTransportSummaryIdentifierPrefix — this is the ONE deliberate
 * deviation from "no new parsing"; every other field is still the reused
 * ICS action's own value. Pure, no GAS globals.
 */
function buildTransportIcsEntry(event) {
  const summary = stripTransportSummaryIdentifierPrefix(event.summary);
  const resource = buildEventResource(event);
  resource.summary = summary;

  return {
    resource: resource,
    summary: summary,
    filenameDate: event.start,
    ticketIdentifier: extractTransportTicketIdentifier(event),
    uid: event.uid,
    status: event.status,
    dtstamp: event.dtstamp,
  };
}

/**
 * partitionTransportEntriesByCancellation — pure, unit-tested split of a
 * shared-shape entries array (see buildTransportIcsEntry/
 * buildTransportBodyEntry) into `{ toCancel, toCreate }`, strictly on
 * `entry.status === 'CANCELLED'`. Compares against the uppercase token ONLY
 * and does NOT re-normalize — the parser (`parseVeventBlock`,
 * src/05-action-ics-import.js) is the single normalization point for this
 * value, so this comparison never needs to trim/uppercase again.
 * `status: null`, an absent `status` key, and any other status value all
 * route to `toCreate`. Preserves relative order within each bucket and
 * never mutates the input array or its entries — this must sit BETWEEN
 * entry construction and the EXISTING seenInBatch/isDuplicateTransportTicket
 * filter in processTransportTicketJob: placing it after that filter would
 * let a cancel entry get dedup-dropped (its identifier deliberately matches
 * the very event it is meant to delete) and the cancellation would
 * silently vanish. Pure, no GAS globals.
 */
function partitionTransportEntriesByCancellation(entries) {
  const toCancel = [];
  const toCreate = [];

  (entries || []).forEach(function (entry) {
    if (entry.status === 'CANCELLED') {
      toCancel.push(entry);
    } else {
      toCreate.push(entry);
    }
  });

  return { toCancel: toCancel, toCreate: toCreate };
}

/**
 * resolveTransportProcessingJobs — the pure, TESTABLE extraction of `run`'s
 * per-message orchestration decision, mirroring resolveTicketProcessingJobs'
 * shape (src/07-action-ticketing-portals.js). Given `messages` (an array of
 * message-like objects exposing `getFrom()`/`getAttachments()`/
 * `getPlainBody()` — GAS `GmailMessage` objects in production, plain
 * duck-typed fakes in tests) and `senders` (the TRANSPORT_SENDERS config
 * array), returns an array of jobs, EACH TAGGED WITH A `mode`:
 *   - `{ mode: 'body', message, sender }` — EXACTLY ONE JOB PER MATCHING
 *     MESSAGE, decided via resolveTransportSenderMode. No `.ics`-attachment
 *     requirement at all — a body-sourced sender's event data is always
 *     present in the message body itself.
 *   - `{ mode: 'ics', message, sender, icsAttachments }` — the EXISTING
 *     RegioJet behavior, completely unchanged: collect the message's
 *     `.ics` attachments (via the REUSED isIcsAttachment,
 *     src/05-action-ics-import.js — no second matcher either); skip the
 *     message if none.
 * A message whose sender does not resolve against `senders` at all
 * contributes NO jobs. EXACTLY ONE JOB PER MATCHING MESSAGE for EITHER
 * mode (never one per attachment for `'ics'` mode either) — deliberate:
 * this action's real duplicate-event guarantee comes from the DEDUP SAFETY
 * NET (isDuplicateTransportTicket) applied per-ENTRY inside
 * processTransportTicketJob, not from restricting which
 * messages/attachments get processed here: a message can legitimately
 * carry more than one `.ics` attachment, and every one of them must still
 * be parsed and dedup-checked, just as one job, not silently dropped.
 * Pure, no GAS globals — every GAS-shaped method call here is invoked ON
 * THE PASSED-IN objects only, never a real global service, so this is
 * fully unit-testable under Node with fake message/attachment objects.
 */
function resolveTransportProcessingJobs(messages, senders) {
  const list = messages || [];
  const jobs = [];

  for (let i = 0; i < list.length; i++) {
    const message = list[i];
    const sender = resolveTransportSender(message.getFrom(), senders);
    if (!sender) {
      continue;
    }

    const mode = resolveTransportSenderMode(sender);

    if (mode === 'body') {
      jobs.push({ mode: 'body', message: message, sender: sender });
      continue;
    }

    const icsAttachments = message.getAttachments().filter(isIcsAttachment);
    if (icsAttachments.length === 0) {
      continue;
    }

    jobs.push({ mode: 'ics', message: message, sender: sender, icsAttachments: icsAttachments });
  }

  return jobs;
}

// Node/GAS environment bridge for parseIcs / buildEventResource /
// isIcsAttachment (defined in the sibling src/05-action-ics-import.js, this
// action reuses that file's parser rather than hand-rolling a second one)
// and for TRANSPORT_TICKETS_ACTION_CONFIG (defined in the sibling
// src/08-action-cfg-transport-tickets.js). Under GAS's shared global scope
// these are ALREADY visible here by bare name — no action needed, and this
// `if` block never executes there. Under Node, each `require()`d file is
// its own isolated module with its own scope, so the bare references
// inside this file's functions/getters would otherwise throw
// ReferenceError. Same `globalThis` bridge technique (not a redeclared
// `const`/`let`/`var`, which would collide under GAS's concatenated scope)
// already established by every other action file's own equivalent bridge.
if (typeof module !== 'undefined' && module.exports) {
  const icsModule = require('./05-action-ics-import.js');
  globalThis.parseIcs = icsModule.parseIcs;
  globalThis.buildEventResource = icsModule.buildEventResource;
  globalThis.isIcsAttachment = icsModule.isIcsAttachment;
  globalThis.TRANSPORT_TICKETS_ACTION_CONFIG = require('./08-action-cfg-transport-tickets.js').TRANSPORT_TICKETS_ACTION_CONFIG;
}

/**
 * getOrCreateTransportDriveFolder — finds a Drive folder by NAME (not ID)
 * via `DriveApp.getFoldersByName`, returning the FIRST match if one or more
 * exist, or creating a new folder via `DriveApp.createFolder` if none exist
 * yet. Used for the permanent CONFIG.ticketAttachmentDriveFolderName
 * folder — the SAME folder the ticketing-portals action uses. This
 * deliberately does NOT reuse getOrCreateDriveFolderByName (src/07-action-
 * ticketing-portals.js), despite the identical implementation: two files
 * declaring the same top-level function name would collide in GAS's single
 * shared global scope.
 */
function getOrCreateTransportDriveFolder(name) {
  const existing = DriveApp.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(name);
}

/**
 * findTransportEventsByIdentifier — the DEDUP SAFETY NET's lookup: searches
 * `calendarId` for existing events already tagged with
 * `extendedProperties.private.ticketIdentifier` equal to `ticketIdentifier`
 * via `Calendar.Events.list(calendarId, { privateExtendedProperty:
 * 'ticketIdentifier=' + ticketIdentifier, singleEvents: true })`. Not
 * paginated/time-windowed — a `privateExtendedProperty` equality filter is
 * already an EXACT match. Returns EVERY matching event (possibly empty),
 * never null.
 *
 * PLURAL BY CONSTRUCTION: a `ticketIdentifier` is NOT unique.
 * `filterTransportEntriesToCreate` skips the dedup pre-check for every
 * uid-bearing entry on the grounds that `Calendar.Events.import`'s iCalUID
 * keying already dedups — and it does, BY UID, NOT by ticket number. A
 * RegioJet reissue whose UID hash changes, or a multi-VEVENT / multi-leg
 * ticket, therefore produces TWO live events sharing ONE ticketIdentifier.
 * Callers that must act on ALL of them (cancelTransportTicketEvent) use
 * this plural function; callers that only ask "does any event already
 * carry this identifier?" use the singular wrapper
 * (findTransportEventByIdentifier) below.
 *
 * GAS-only (Calendar global) — not unit-tested directly, but fully
 * exercised through cancelTransportTicketEvent under a fake
 * global.Calendar.
 */
function findTransportEventsByIdentifier(ticketIdentifier, calendarId) {
  const response = Calendar.Events.list(calendarId, {
    privateExtendedProperty: 'ticketIdentifier=' + ticketIdentifier,
    singleEvents: true,
  });
  return (response && response.items) || [];
}

/**
 * findTransportEventByIdentifier — the single-event convenience wrapper over
 * findTransportEventsByIdentifier above: returns the first matching event, or
 * `null` if none found. This is the right shape for the DEDUP SAFETY NET's
 * existence question (isDuplicateTransportTicket), which only ever asks
 * "does any event already carry this identifier?" — never "which ones?".
 * GAS-only (Calendar global).
 */
function findTransportEventByIdentifier(ticketIdentifier, calendarId) {
  const items = findTransportEventsByIdentifier(ticketIdentifier, calendarId);
  return items.length > 0 ? items[0] : null;
}

/**
 * isDuplicateTransportTicket — the DEDUP SAFETY NET's decision, mirroring
 * isDuplicateTicketPurchase's exact shape (src/07-action-ticketing-
 * portals.js). Returns `false` for a falsy `ticketIdentifier` (an accepted
 * per-parse limitation, not a silent gap — a ticket whose identifier could
 * not be extracted simply does not get this protection), otherwise `true`
 * when findTransportEventByIdentifier finds a match on `calendarId`.
 */
function isDuplicateTransportTicket(ticketIdentifier, calendarId) {
  if (!ticketIdentifier) {
    return false;
  }

  const existingEvent = findTransportEventByIdentifier(ticketIdentifier, calendarId);
  if (existingEvent) {
    console.log(
      'Transport tickets: event for ticket identifier ' + ticketIdentifier + ' already exists, skipping (safety-net, not a duplicate path).'
    );
    return true;
  }

  return false;
}

/**
 * buildTransportEventPrivateProperties — the SINGLE writer of the
 * `extendedProperties.private` tag isTransportCancellationStale later
 * reads. Returns `null` when `entry.ticketIdentifier` is falsy — meaning
 * the write loop writes no `extendedProperties` object at all. Otherwise
 * returns `{ ticketIdentifier }`, plus a `dtstamp` key set to
 * `entry.dtstamp.toISOString()` ONLY when `entry.dtstamp` is a real, valid
 * Date — omitted entirely otherwise (never string-coerced), so the tagged
 * value is never the literal 4-character word "null". Pure, no GAS
 * globals.
 */
function buildTransportEventPrivateProperties(entry) {
  if (!entry.ticketIdentifier) {
    return null;
  }

  const properties = { ticketIdentifier: entry.ticketIdentifier };

  if (entry.dtstamp instanceof Date && !Number.isNaN(entry.dtstamp.getTime())) {
    properties.dtstamp = entry.dtstamp.toISOString();
  }

  return properties;
}

/**
 * isTransportCancellationStale — true ONLY when the found event's stored
 * `extendedProperties.private.dtstamp` (see
 * buildTransportEventPrivateProperties, the single writer of this tag) is
 * present AND strictly newer than the cancellation entry's OWN
 * `cancelDtstamp` — meaning a rebooking has already overwritten this event
 * since the cancellation was generated. EVERY missing/unparseable case
 * returns `false` so the caller falls back to the current, unchanged
 * behavior (delete): no stored dtstamp at all, no `extendedProperties`/
 * `private` at any level (never throws), an unparseable stored value, a
 * falsy `cancelDtstamp`, or an equal timestamp (strictly newer, not
 * newer-or-equal). An absent optional signal must never block the
 * cancellation guarantee. Pure, no GAS globals (operates only on the plain
 * `existingEvent` object already returned by
 * findTransportEventByIdentifier).
 */
function isTransportCancellationStale(existingEvent, cancelDtstamp) {
  if (!cancelDtstamp) {
    return false;
  }

  const cancelTime = new Date(cancelDtstamp).getTime();
  if (Number.isNaN(cancelTime)) {
    return false;
  }

  const storedValue =
    existingEvent &&
    existingEvent.extendedProperties &&
    existingEvent.extendedProperties.private &&
    existingEvent.extendedProperties.private.dtstamp;
  if (!storedValue) {
    return false;
  }

  const storedTime = new Date(storedValue).getTime();
  if (Number.isNaN(storedTime)) {
    return false;
  }

  return storedTime > cancelTime;
}

/**
 * cancelTransportTicketEvent — deletes the calendar event(s) a RegioJet
 * cancellation entry refers to, shaped and guarded like
 * isDuplicateTransportTicket above.
 *
 * 1. A falsy `ticketIdentifier` (null/empty string) is guarded FIRST — logs
 *    and returns WITHOUT ever reaching the Calendar API.
 * 2. No match is a SILENT no-op: logs and returns, never throws.
 *    `src/02-main.js`'s `orderThreadsForProcessing` rests its ordering
 *    rationale on this branch being documented here — mirrors
 *    booking.com's `handleCancellation`
 *    (src/06-action-booking-com-management.js), whose no-match branch is
 *    the same precedent.
 * 3. STALE-CANCELLATION GUARD: once a match is found and BEFORE the
 *    delete, the OPTIONAL third parameter `cancelDtstamp` (the
 *    cancellation entry's OWN dtstamp) is compared against the found
 *    event's stored dtstamp tag via isTransportCancellationStale. A
 *    cancellation and a rebooking for the same ticket are two INDEPENDENT
 *    messages/threads, so nothing guarantees which is processed first — if
 *    the found event's stored dtstamp is NEWER than this cancellation's
 *    own, a later booking has already superseded the ticket, and deleting
 *    would destroy the live rebooking. SEQUENCE cannot resolve this
 *    ordering — RegioJet RESETS it across a cancel+rebook pair (observed
 *    1 -> 2 -> 1) — whereas DTSTAMP, the real send time, is monotonic.
 *    Every missing-timestamp case (omitted third argument, no stored tag,
 *    unparseable value) falls back to the ORIGINAL unchanged behavior —
 *    delete — so an absent optional signal never blocks the cancellation
 *    guarantee.
 * 4. A match that is NOT stale calls `Calendar.Events.remove(calendarId,
 *    existingEvent.id)` — that EXACT argument order (calendar first, event
 *    id second) is the same convention used at `handleCancellation`
 *    (src/06-action-booking-com-management.js) — then logs success naming
 *    the identifier and the calendar.
 *
 * STRICT-IDENTIFIER-MATCH GUARANTEE: matching is STRICTLY on the exact
 * `ticketIdentifier` — RegioJet's own unique per-ticket number — via
 * findTransportEventsByIdentifier's `privateExtendedProperty` equality
 * query, NEVER on date, time or route. Deliberately NO fuzzy or
 * date-time-overlap fallback (unlike booking.com's hotel-name+date-overlap
 * fallback) — that would reintroduce exactly the cross-contamination race
 * a strict identifier match rules out. This holds regardless of which
 * email (the new confirmation or the old cancellation) is processed first,
 * since the lookup never considers date/time at all. The dtstamp
 * comparison above is a staleness guard on an ALREADY identifier-matched
 * event, never a matching mechanism of its own.
 *
 * GAS-only (calls findTransportEventsByIdentifier, which touches the
 * Calendar global, and calls Calendar.Events.remove directly); exercised
 * in tests through a faked global.Calendar, see
 * test/transport-tickets.test.js.
 */
function cancelTransportTicketEvent(ticketIdentifier, calendarId, cancelDtstamp) {
  if (!ticketIdentifier) {
    console.log('Transport tickets: cancellation entry has no ticketIdentifier, skipping (silent no-op, never reaches the Calendar API).');
    return;
  }

  const existingEvents = findTransportEventsByIdentifier(ticketIdentifier, calendarId);
  if (existingEvents.length === 0) {
    console.log(
      'Transport tickets: cancellation for ticket identifier ' + ticketIdentifier + ' has no matching calendar event, skipping (silent no-op, D-05 accepted limitation).'
    );
    return;
  }

  let removedCount = 0;

  existingEvents.forEach(function (existingEvent) {
    // Staleness is a property of the INDIVIDUAL event (its own stored
    // dtstamp tag), never of the identifier as a whole -- two events
    // sharing one ticket number can genuinely disagree about it.
    // Evaluating it per event keeps the guarantee exact instead of letting
    // one stale match veto every deletion.
    if (isTransportCancellationStale(existingEvent, cancelDtstamp)) {
      const storedDtstamp = existingEvent.extendedProperties.private.dtstamp;
      console.log(
        'Transport tickets: cancellation for ticket identifier ' + ticketIdentifier + ' is STALE for event ' + existingEvent.id + ' (event dtstamp ' + storedDtstamp + ' is newer than the cancellation\'s own dtstamp ' + cancelDtstamp + ') -- a rebooking already superseded this event, skipping its deletion (D-11).'
      );
      return;
    }

    Calendar.Events.remove(calendarId, existingEvent.id);
    removedCount += 1;
  });

  // The count distinguishes a complete cancellation from a partial one.
  if (removedCount === 0) {
    console.log(
      'Transport tickets: cancellation for ticket identifier ' + ticketIdentifier + ' matched ' + existingEvents.length + ' event(s) on calendar ' + calendarId + ', but every one was stale -- nothing deleted (D-11).'
    );
    return;
  }

  console.log(
    'Transport tickets: cancelled (deleted) ' + removedCount + ' of ' + existingEvents.length + ' matching calendar event(s) for ticket identifier ' + ticketIdentifier + ' on calendar ' + calendarId + '.'
  );
}

/**
 * filterTransportEntriesToCreate — the EXISTING seenInBatch +
 * isDuplicateTransportTicket dedup filter, extracted out of
 * processTransportTicketJob so it is reachable under Node with a fake
 * global.Calendar, mirroring the same reason cancelTransportTicketEvent is
 * already exported.
 *
 * `isDuplicateTransportTicket` is consulted ONLY for an entry with NO
 * `uid` — the call-site condition `!entry.uid && ...` below is the entire
 * rule. RegioJet reuses the SAME `ticketIdentifier` AND the SAME iCalUID
 * across a cancel+rebook pair; for a `uid`-bearing entry this pre-check is
 * not merely redundant on the happy path but ACTIVELY WRONG on a reissue,
 * since `Calendar.Events.import` (reached via
 * `importIcsEventWithSequenceRetry`, keyed on `entry.uid`, see
 * processTransportTicketJob's write loop) is ALREADY idempotent by iCalUID
 * and creates-or-updates the correct single event with no help from the
 * ticketIdentifier pre-check. The pre-check remains the ONLY protection
 * `uid`-less IDOS.cz entries have — `Calendar.Events.insert` has none of
 * its own — so it stays fully in force for them.
 *
 * `seenInBatch` is DELIBERATELY NOT narrowed alongside it — it keeps
 * applying to EVERY entry, `uid`-bearing or not: it is a WITHIN-ONE-PASS
 * guard against two same-batch entries sharing an identifier, and it
 * cannot cause the reissue problem above, because a confirmation and its
 * later cancel/rebook are separate `processTransportTicketJob` calls, each
 * with a fresh, empty `seenInBatch`.
 *
 * Preserves relative order, never mutates the input array or its entries.
 * GAS-only in the sense that it touches the Calendar global transitively
 * (via isDuplicateTransportTicket, for uid-less entries only); exercised
 * in tests through a faked global.Calendar, see
 * test/transport-tickets.test.js.
 */
function filterTransportEntriesToCreate(entries, calendarId) {
  const seenInBatch = {};

  return (entries || []).filter(function (entry) {
    if (entry.ticketIdentifier && seenInBatch[entry.ticketIdentifier]) {
      return false;
    }
    if (!entry.uid && isDuplicateTransportTicket(entry.ticketIdentifier, calendarId)) {
      return false;
    }
    if (entry.ticketIdentifier) {
      seenInBatch[entry.ticketIdentifier] = true;
    }
    return true;
  });
}

/**
 * processTransportTicketJob — the pipeline for ONE job (see
 * resolveTransportProcessingJobs' own JSDoc for the `{ mode, message,
 * sender, icsAttachments? }` shape), in this exact order so nothing is EVER
 * created in Drive before the decision to write is final (see this file's
 * class-level "ORDERING GUARANTEE" doc):
 *   1. Resolve the calendar ONCE via resolveTransportCalendarId(job.sender,
 *      CONFIG.calendarId) and thread it through every downstream call —
 *      never re-read `sender.calendarId` at a call site (see
 *      resolveTransportCalendarId's own JSDoc for the crash class this
 *      avoids).
 *   2. Build `entries` (the SHARED
 *      `{ resource, summary, filenameDate, ticketIdentifier, uid, status, dtstamp }`
 *      shape) — this is the ONLY step that differs per mode:
 *        - `mode: 'body'`: look up the registered body parser
 *          (TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL); a body-mode
 *          sender with NO registered parser throws a controlled Error
 *          naming the sender (never a silent no-op), parse
 *          `job.message.getPlainBody()`, and build ONE entry via
 *          buildTransportBodyEntry, using the calendar's LIVE-DERIVED
 *          timezone (`CalendarApp.getCalendarById(calendarId).getTimeZone()`
 *          — never hardcoded).
 *        - `mode: 'ics'`: parse EVERY `.ics` attachment's
 *          `getDataAsString()` through the REUSED parseIcs and map each
 *          resulting event through buildTransportIcsEntry — the EXISTING
 *          RegioJet behavior, now expressed through the shared entry shape,
 *          with no new parsing. All parsing completes before any write
 *          begins (fail-closed, same discipline as ICS_CALENDAR_ACTION.run).
 *      The calendar's timezone lookup and the pure buildEventResource call
 *      happening here (before the dedup filter below) are reads/pure
 *      calls, not writes — this file's "ORDERING GUARANTEE" doc concerns
 *      Drive/Calendar WRITES specifically, and remains unaffected.
 *   2.5. PARTITION: split `entries` via
 *      partitionTransportEntriesByCancellation into `{ toCancel, toCreate }`.
 *      This MUST sit here — AFTER entries is built, BEFORE the dedup filter
 *      in step 3 — because a cancel entry's ticketIdentifier deliberately
 *      matches the very event it is meant to delete; running the dedup
 *      filter on it first would drop it as an "already exists" duplicate
 *      and the cancellation would silently vanish. Only the SPLIT happens
 *      here; the cancellations themselves run LAST (step 6).
 *   3. Drop already-present `toCreate` entries via
 *      filterTransportEntriesToCreate(toCreate, calendarId) — see that
 *      function's own JSDoc for the full uid-less-only dedup rationale. If
 *      nothing remains, steps 4 and 5 are skipped entirely — no Drive
 *      upload, no write — but step 6 STILL RUNS (see
 *      writeTransportTicketEvents' own JSDoc: that early exit ends the
 *      WRITE PHASE, never the job; a cancellation-only message has nothing
 *      to create by definition and must still cancel).
 *   4. If `job.sender.insertPdfIntoEvent` is true: find the ticket PDF
 *      (findTransportTicketPdfAttachment, which excludes invoice.pdf and
 *      also matches the IDOS.cz ticket PDF despite its
 *      application/octet-stream content type); if found, copy its blob
 *      DIRECTLY into
 *      getOrCreateTransportDriveFolder(CONFIG.ticketAttachmentDriveFolderName)
 *      — no temp-folder hop, unlike the ticketing-portals action's OCR
 *      pipeline, since this action never needs the PDF's text — rename it
 *      via buildTransportAttachmentFilename(entriesToCreate[0].summary,
 *      .filenameDate, .ticketIdentifier) BEFORE reading its id/url/name
 *      (the displayed Calendar attachment title is derived from the file's
 *      name AT THAT POINT), and build `{ fileId, fileUrl, title }`. If no
 *      ticket PDF is found, log and continue without one — a missing
 *      attachment must never block the calendar event itself.
 *   5. For each remaining entry: set `resource.extendedProperties.private`
 *      to buildTransportEventPrivateProperties(entry)'s return value when it
 *      is truthy (`{ ticketIdentifier }`, plus a `dtstamp` ISO string when
 *      the entry carries a real one; `null` means no `extendedProperties`
 *      object is written at all), and, when an attachment exists,
 *      `resource.attachments = [{ fileId, fileUrl, title, mimeType:
 *      'application/pdf' }]` plus `optionalArgs.supportsAttachments = true`
 *      (the exact shape confirmed live by createTicketCalendarEvent,
 *      src/07-action-ticketing-portals.js). Writes via
 *      importIcsEventWithSequenceRetry(resource, calendarId, entry.uid,
 *      optionalArgs) when the entry carries a `uid` — REUSING
 *      ICS_CALENDAR_ACTION's proven idempotent-by-iCalUID path (only ever
 *      true for `'ics'`-mode entries), which also protects against Gmail's
 *      own native invite detection creating a second event from the same
 *      .ics — and Calendar.Events.insert(resource, calendarId,
 *      optionalArgs) otherwise (always true for `'body'`-mode entries,
 *      which never carry a UID).
 *   6. CANCELLATION, LAST: run cancelTransportTicketEvent over EVERY
 *      `toCancel` entry from step 2.5, with the already-resolved
 *      `calendarId` AND `entry.dtstamp` as the third argument, so a
 *      cancellation already superseded by a newer rebooking is detected and
 *      skipped. Running this last (rather than alongside step 2.5) makes
 *      `STATUS:CANCELLED` authoritative regardless of how a sender packages
 *      its VEVENTs — a cancel evaluated BEFORE the create it refers to
 *      would find no event, take its silent no-op branch, and the create
 *      would then win, leaving a CANCELLED ticket on the calendar with no
 *      error and no failure label. Mirrors the same causal ordering
 *      orderThreadsForProcessing (src/02-main.js) guarantees one level up,
 *      across threads.
 *
 * GAS-only (DriveApp/CalendarApp/Calendar globals); exercised in tests
 * through TRANSPORT_TICKETS_ACTION.run with faked globals, see
 * withTransportRunGlobals in test/transport-tickets.test.js.
 */
function processTransportTicketJob(job) {
  const calendarId = resolveTransportCalendarId(job.sender, CONFIG.calendarId);

  let entries;
  if (job.mode === 'body') {
    const senderKey = transportExtractEmailAddress(job.sender.identifyingEmail);
    const parseBodyText = TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey];
    if (!parseBodyText) {
      throw new Error('No transport body parser registered for sender: ' + job.sender.identifyingEmail);
    }

    const timeZone = CalendarApp.getCalendarById(calendarId).getTimeZone();
    const parsedTicket = parseBodyText(job.message.getPlainBody());
    entries = [buildTransportBodyEntry(parsedTicket, timeZone)];
  } else {
    const events = job.icsAttachments.reduce(function (allEvents, attachment) {
      return allEvents.concat(parseIcs(attachment.getDataAsString()));
    }, []);
    entries = events.map(buildTransportIcsEntry);
  }

  // Partition BEFORE the dedup filter below -- see this function's own
  // JSDoc, step 2.5, for why.
  const partitioned = partitionTransportEntriesByCancellation(entries);

  // D-08 of quick-260813-dq2 Task 3: filterTransportEntriesToCreate is the
  // EXISTING seenInBatch + isDuplicateTransportTicket filter (WR-01 of the
  // 260803-us3 review), extracted verbatim with exactly one behavior
  // change — see its own JSDoc for the full Problem A rationale.
  const entriesToCreate = filterTransportEntriesToCreate(partitioned.toCreate, calendarId);

  // CREATES SETTLE BEFORE CANCELS -- see this function's own JSDoc, step 6,
  // for why.
  writeTransportTicketEvents(job, entriesToCreate, calendarId);

  partitioned.toCancel.forEach(function (entry) {
    cancelTransportTicketEvent(entry.ticketIdentifier, calendarId, entry.dtstamp);
  });
}

/**
 * writeTransportTicketEvents — steps 4 and 5 of processTransportTicketJob
 * (the PDF archive/attach block and the calendar write loop). The early
 * `return` below is a LOCAL return from THIS function only — it ends the
 * write phase, never the job. Returning out of the whole job at that point
 * would silently skip every cancellation on a cancellation-only message:
 * the single most common shape this action sees.
 *
 * Behaviour of the two steps themselves is UNCHANGED — see
 * processTransportTicketJob's JSDoc (steps 4 and 5) for their full
 * contract. GAS-only (DriveApp/Calendar globals).
 */
function writeTransportTicketEvents(job, entriesToCreate, calendarId) {
  if (entriesToCreate.length === 0) {
    return;
  }

  let attachmentInfo = null;
  if (job.sender.insertPdfIntoEvent) {
    const pdfAttachment = findTransportTicketPdfAttachment(job.message);

    if (pdfAttachment) {
      const permanentFolder = getOrCreateTransportDriveFolder(CONFIG.ticketAttachmentDriveFolderName);
      const permanentPdfFile = permanentFolder.createFile(pdfAttachment.copyBlob());
      // Rename BEFORE reading getId()/getUrl()/getName() below -- the
      // Calendar event attachment's displayed title is derived from the
      // file's name AT THIS POINT (same convention as the ticketing-portals
      // action's own processTicketPdfAttachment).
      permanentPdfFile.setName(
        buildTransportAttachmentFilename(entriesToCreate[0].summary, entriesToCreate[0].filenameDate, entriesToCreate[0].ticketIdentifier)
      );
      attachmentInfo = {
        fileId: permanentPdfFile.getId(),
        fileUrl: permanentPdfFile.getUrl(),
        title: permanentPdfFile.getName(),
      };
    } else {
      console.log(
        'Transport tickets: insertPdfIntoEvent is true but no matching ticket PDF attachment was found on the message; creating the event(s) without an attachment.'
      );
    }
  }

  entriesToCreate.forEach(function (entry) {
    const resource = entry.resource;

    // D-10 of quick-260813-dq2 Task 3: buildTransportEventPrivateProperties
    // is the SINGLE writer of this tag -- isTransportCancellationStale is
    // its only reader.
    const privateProperties = buildTransportEventPrivateProperties(entry);
    if (privateProperties) {
      resource.extendedProperties = { private: privateProperties };
    }

    const optionalArgs = {};
    if (attachmentInfo) {
      resource.attachments = [
        { fileId: attachmentInfo.fileId, fileUrl: attachmentInfo.fileUrl, title: attachmentInfo.title, mimeType: 'application/pdf' },
      ];
      // A real, documented Calendar API v3 requirement (supportsAttachments,
      // default false) for the attachments array above to be accepted at
      // all -- never optional here when an attachment is present. See this
      // function's own "FLAGGED ASSUMPTION" doc above for the import-path
      // caveat.
      optionalArgs.supportsAttachments = true;
    }

    console.log(
      'Transport tickets: creating calendar event for "' + entry.summary + '" (ticketIdentifier=' + entry.ticketIdentifier + ') on calendar ' + calendarId + '.'
    );

    if (entry.uid) {
      importIcsEventWithSequenceRetry(resource, calendarId, entry.uid, optionalArgs);
    } else {
      Calendar.Events.insert(resource, calendarId, optionalArgs);
    }
  });
}

/**
 * TRANSPORT_TICKETS_ACTION — the transport-tickets action descriptor.
 * Carries its own config block (TRANSPORT_TICKETS_ACTION_CONFIG),
 * independent of CONFIG and of any other action's config (except for the
 * one shared cross-cutting CONFIG.ticketAttachmentDriveFolderName field —
 * same sharing as the ticketing-portals action).
 */
const TRANSPORT_TICKETS_ACTION = {
  name: 'transport-tickets',

  // GETTER, not a plain literal property — see this file's class-level
  // JSDoc and the sibling config file's own "CONFIG SPLIT" note. Not
  // evaluated at object-construction time, only when something reads
  // `.config`, which happens lazily inside function bodies (dispatchActions,
  // notifyOwnerOfFailure) long after every project file has loaded.
  get config() {
    return TRANSPORT_TICKETS_ACTION_CONFIG;
  },

  /**
   * appliesTo — returns a LITERAL boolean (dispatchActions only skips on a
   * strict `=== false`). True when resolveTransportProcessingJobs finds at
   * least one job on the thread — for an `'ics'`-mode sender, a message
   * carrying at least one `.ics` attachment; for a `'body'`-mode sender,
   * any matching message at all, no attachment required.
   */
  appliesTo: function (thread) {
    return resolveTransportProcessingJobs(thread.getMessages(), TRANSPORT_TICKETS_ACTION.config.transportSenders).length > 0;
  },

  /**
   * run — builds the processing job list via the pure, TESTABLE
   * resolveTransportProcessingJobs, then runs processTransportTicketJob for
   * each job.
   */
  run: function (thread) {
    const jobs = resolveTransportProcessingJobs(thread.getMessages(), TRANSPORT_TICKETS_ACTION.config.transportSenders);
    jobs.forEach(processTransportTicketJob);
  },
};

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// The Calendar-touching functions that ARE exported below
// (cancelTransportTicketEvent, isTransportCancellationStale,
// filterTransportEntriesToCreate) touch the Calendar global and are
// GAS-only in that sense, but are exported anyway specifically so their
// branching can be proven under Node with a fake global.Calendar, rather
// than deferred entirely to a live round. getOrCreateTransportDriveFolder/
// findTransportEventByIdentifier/findTransportEventsByIdentifier/
// isDuplicateTransportTicket/writeTransportTicketEvents/
// processTransportTicketJob remain genuinely GAS-only (reference
// DriveApp/CalendarApp/Calendar globals directly) and are NOT exported.
//
// They ARE reachable under Node INDIRECTLY, through the exported
// TRANSPORT_TICKETS_ACTION.run with faked Calendar/CalendarApp/DriveApp
// globals (see withTransportRunGlobals in test/transport-tickets.test.js).
// That helper must ALSO wire `globalThis.importIcsEventWithSequenceRetry`:
// GAS concatenates every project file into ONE shared global scope, so
// processTransportTicketJob's bare reference to it resolves there, but
// this file's Node bridge above only wires
// parseIcs/buildEventResource/isIcsAttachment from that same sibling
// module.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    transportExtractEmailAddress: transportExtractEmailAddress,
    resolveTransportSender: resolveTransportSender,
    resolveTransportCalendarId: resolveTransportCalendarId,
    isTransportPdfAttachment: isTransportPdfAttachment,
    findTransportTicketPdfAttachment: findTransportTicketPdfAttachment,
    extractTransportTicketIdentifier: extractTransportTicketIdentifier,
    sanitizeTransportFilenameComponent: sanitizeTransportFilenameComponent,
    buildTransportAttachmentFilename: buildTransportAttachmentFilename,
    resolveTransportProcessingJobs: resolveTransportProcessingJobs,
    parseIdosTicketText: parseIdosTicketText,
    resolveTransportSenderMode: resolveTransportSenderMode,
    buildTransportBodyEntry: buildTransportBodyEntry,
    buildTransportIcsEntry: buildTransportIcsEntry,
    stripTransportSummaryIdentifierPrefix: stripTransportSummaryIdentifierPrefix,
    formatTransportWallClockIso: formatTransportWallClockIso,
    partitionTransportEntriesByCancellation: partitionTransportEntriesByCancellation,
    cancelTransportTicketEvent: cancelTransportTicketEvent,
    buildTransportEventPrivateProperties: buildTransportEventPrivateProperties,
    isTransportCancellationStale: isTransportCancellationStale,
    filterTransportEntriesToCreate: filterTransportEntriesToCreate,
    TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL: TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL,
    TRANSPORT_TICKETS_ACTION: TRANSPORT_TICKETS_ACTION,
  };
}

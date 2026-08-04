/**
 * TRANSPORT_TICKETS_ACTION — quick-260803-us3 NEW ACTION: detects train/bus
 * ticket confirmation emails (RegioJet first, extensible to other carriers
 * via TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders — D-05) and creates
 * ONE calendar event per email/ticket.
 *
 * UNLIKE every other action in this codebase, this action's event data
 * comes from the email's OWN `text/calendar` `.ics` attachment, parsed with
 * the EXISTING iCalendar parser in src/05-action-ics-import.js (parseIcs /
 * buildEventResource) — D-01. There is DELIBERATELY no second VEVENT parser
 * in this project: hand-rolling one here would duplicate the VTIMEZONE
 * resolution logic (extractVtimezoneBlocks / resolveTzidDate) that file
 * already gets right, including the live-reported 2-hour TZID timezone bug
 * fix from Phase 3.
 *
 * HAND-OFF FROM ICS_CALENDAR_ACTION (D-03/D-06): ICS_CALENDAR_ACTION
 * already imports `.ics` attachments from ANY sender by default, so without
 * intervention it would ALSO claim a RegioJet confirmation email, producing
 * a second, competing calendar event. A sender owned by this action is
 * therefore expected to ALSO be listed in ICS_ACTION_CONFIG.excludeFrom
 * (src/05-action-cfg-ics-import.js) so ICS_CALENDAR_ACTION stands down for
 * it. This is an OWNER-SIDE Script Properties step (set live, after
 * `rebuildScriptProperties()`), not a code default — this action's own
 * detection (the presence of a `.ics` attachment from a configured
 * transport sender) is independent of ICS_ACTION_CONFIG entirely (D-06):
 * that is a separate action's config and must never gate this one.
 *
 * ORDERING GUARANTEE (scoped — 260803-us3 review WR-02): no Drive file is
 * EVER created before the dedup decision is made (see
 * processTransportTicketJob's step order below), so the SPECIFIC failure
 * mode round 10 fixed in ticketing-portals (a file's fate left undecided
 * because parsing/dedup happened AFTER upload, see that action's own
 * "ORPHANED TEMP-PDF FIX" doc, src/07-action-ticketing-portals.js) cannot
 * recur here — this action never uploads to a TEMP folder at all; the
 * ticket PDF, when insertPdfIntoEvent is true, is copied DIRECTLY into the
 * PERMANENT CONFIG.ticketAttachmentDriveFolderName folder only AFTER the
 * dedup check has already decided this run should proceed. This does NOT
 * mean the permanent file can never end up unattached to any event: if a
 * Calendar write later in the same run throws (quota/permission error, or
 * the FLAGGED ASSUMPTION below), the already-uploaded/renamed permanent
 * file is left with no reconciling try/finally — same pre-existing,
 * accepted gap as ticketing-portals' own processTicketFromMessageBody path,
 * not a new regression, not yet worth the added complexity of a
 * pdfFateResolved-style guard here.
 *
 * DEDUP (D-02): the created event is tagged with
 * `extendedProperties.private.ticketIdentifier`, pre-checked via
 * `findTransportEventByIdentifier` before any write — same proven pattern
 * as the ticketing-portals action's own `ticketIdentifier` safety net (see
 * that action's class-level "DEDUP SAFETY NET" doc). The identifier's
 * PRIMARY source is the VEVENT SUMMARY's leading `#<digits>` (e.g.
 * `#7788123456`) — it matches the human-facing ticket number shown in the
 * email subject, unlike the opaque `UID` hash (a negative number for
 * RegioJet, e.g. `-9876543210@regiojet.cz`). `UID` is the fallback when no
 * `#<digits>` prefix is present; `null` (never throw) when neither is
 * present — see extractTransportTicketIdentifier's own JSDoc.
 *
 * PDF ARCHIVE + ATTACH (D-04): when a resolved sender's `insertPdfIntoEvent`
 * is true, the ticket PDF (found via findTransportTicketPdfAttachment,
 * which EXCLUDES an accompanying `invoice.pdf` — that is not the ticket) is
 * copied into `CONFIG.ticketAttachmentDriveFolderName` — the SAME permanent
 * Drive folder src/07-action-ticketing-portals.js already uses (a single
 * shared destination, never a second folder setting) — renamed via
 * buildTransportAttachmentFilename, and attached to the created event using
 * the live-proven `{fileId, fileUrl, title, mimeType}` EventAttachment
 * shape plus `supportsAttachments: true` (see
 * src/07-action-ticketing-portals.js's class-level "round 7" doc for the
 * real live bug that established this exact shape — `fileUrl` is REQUIRED,
 * `fileId` alone is NOT sufficient).
 *
 * GLOBALLY-UNIQUE NAMING (see the ticketing-portals/booking actions' own
 * class-level JSDoc for the real cross-file naming-collision incident that
 * established this convention): every pure helper in this file is
 * `transport`-prefixed to avoid colliding with the same-purpose helpers
 * already declared in src/05-*, src/06-*, and src/07-* — GAS concatenates
 * every project file into ONE shared global scope, so two files declaring
 * the same top-level function name silently collide.
 *
 * src/07-action-ticketing-portals.js (the OCR/PDF-text pipeline) is NOT
 * modified at all by this task — this is a separate, .ics-driven action,
 * not a ticketing-portals mode.
 *
 * TWO PROCESSING MODES (quick-260804-bs7): a second carrier, IDOS.cz
 * (`jizdenky@idos.svt.cz`), sends confirmation emails with NO `.ics`
 * attachment at all — route, dates, times, seats and both codes live only
 * in the plain-text email body. Rather than forcing every future carrier
 * through the `.ics`-VEVENT model above, `transportSenders` entries now
 * carry a `mode` field (`'ics' | 'body'`), mirroring the EXACT
 * `mode: 'pdf'|'body'` split TICKETING_PORTALS_ACTION already proved
 * (src/07-action-ticketing-portals.js, resolveTicketProcessingJobs /
 * processTicketFromMessageBody):
 *   - `'ics'` (RegioJet, unchanged): parseIcs/buildEventResource ->
 *     buildTransportIcsEntry.
 *   - `'body'` (IDOS.cz, new): message.getPlainBody() -> a
 *     carrier-specific text parser (parseIdosTicketText) ->
 *     buildTransportBodyEntry.
 * `resolveTransportSenderMode` decides which mode a sender resolves to;
 * both modes then feed the SAME shared entry shape
 * (`{ resource, summary, filenameDate, ticketIdentifier, uid }`) into the
 * SAME dedup -> PDF-archive -> write pipeline in processTransportTicketJob
 * (D-05) — only the extraction step differs, never the downstream
 * write/dedup/attach logic.
 */

/**
 * transportExtractEmailAddress — LOCAL copy of the same underlying logic
 * already re-implemented independently in the ICS, booking.com, and
 * ticketing-portals action files (see the booking.com action's own JSDoc
 * for the real cross-file naming-collision incident that established the
 * "globally-unique name per action" convention this follows). Extracts the
 * bare, trimmed, lowercased email address from a Gmail "From" header value,
 * or from a bare address with no display name. Pure, no GAS globals. Never
 * throws: a null/undefined/empty input returns ''.
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
 * TRANSPORT_SENDERS entry). Two-tier resolution, mirroring
 * resolveTicketingCalendarId's exact shape and its documented live-crash
 * rationale (src/07-action-ticketing-portals.js, quick-260731-tix round 6:
 * a shipped `calendarId: null` default read DIRECTLY with no fallback
 * caused a real live `TypeError: Cannot read properties of null (reading
 * 'getTimeZone')`): `sender.calendarId` when truthy, else
 * `defaultCalendarId`. Pure, no GAS globals.
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
 * name does NOT start with `invoice`, or `null` if none qualify. This
 * exclusion is scoped NARROWLY to the filename actually observed on the
 * real RegioJet fixture (`eticket.pdf` is the real ticket, `invoice.pdf` is
 * a separate accompanying invoice, NOT the ticket) — per this codebase's
 * "don't guess at an unobserved variant" discipline (see
 * findKinoArtTicketPdfAttachment's own JSDoc, src/07-action-ticketing-
 * portals.js, for the same scoping discipline applied to a different real
 * fixture). Pure, no GAS globals (operates only on the array
 * message.getAttachments() already produces).
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
 * extractTransportTicketIdentifier — D-02: the DEDUP SAFETY NET's stable
 * key for a parsed event object (as returned by parseIcs, src/05-action-
 * ics-import.js). PRIMARY source: the trimmed `event.summary`'s leading
 * `#<digits>` prefix (e.g. `#7788123456` -> `'7788123456'`) — this is the
 * HUMAN-FACING ticket number that also appears in the email subject,
 * unlike `event.uid` (an opaque, often negative, hash e.g.
 * `-9876543210@regiojet.cz`). FALLBACK: the raw `event.uid` when no
 * `#<digits>` prefix is present. `null` (never throws) when NEITHER is
 * present — a missing dedup key must never block calendar-event creation,
 * same philosophy as parseEnigooTicketText's own optional ticketIdentifier
 * anchor (src/07-action-ticketing-portals.js). Pure, no GAS globals.
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
 * characters (`/ \ ? % * : | " < >`) with `-` and trims whitespace. Same
 * character set and behavior as sanitizeTicketAttachmentFilenameComponent
 * (src/07-action-ticketing-portals.js), locally reimplemented per this
 * codebase's globally-unique-naming convention. Pure, no GAS globals.
 */
function sanitizeTransportFilenameComponent(value) {
  return String(value)
    .replace(/[/\\?%*:|"<>]/g, '-')
    .trim();
}

/**
 * buildTransportAttachmentFilename — the project-wide
 * `"{name} - {YYYY-MM-DD} - {identifier}.pdf"` attachment-renaming
 * convention (see buildTicketAttachmentFilename's own JSDoc,
 * src/07-action-ticketing-portals.js, for the full rationale this mirrors:
 * a Calendar event attachment's displayed `title` is derived from the
 * file's name AT ATTACH TIME, so renaming the file before it is referenced
 * improves both the Drive folder's browsability AND the calendar event's
 * displayed attachment name). `summary` is the VEVENT SUMMARY text (e.g.
 * the real `#7788123456: Z Ostrava, hl.n., do Praha, hl.n., sedadla:
 * [2/15,2/16]`), sanitized via sanitizeTransportFilenameComponent.
 * `startDate` is a real Date (the parsed event's `.start`); the ISO date
 * segment is taken from its UTC calendar date (documented as a knowingly
 * minor simplification affecting only the filename, same caveat as every
 * other date-in-filename helper in this codebase). `ticketIdentifier`
 * (see extractTransportTicketIdentifier) is OMITTED entirely — never
 * string-coerced — when falsy, so no filename ever embeds the literal
 * 4-character word "null" (the defensive rule established in
 * quick-260731-kar round 4, see buildTicketAttachmentFilename's own
 * "DEFENSIVE NULL-HANDLING" doc). Pure, no GAS globals.
 */
function buildTransportAttachmentFilename(summary, startDate, ticketIdentifier) {
  const isoDate = startDate.toISOString().slice(0, 10);
  const ticketIdentifierSegment = ticketIdentifier ? ' - ' + ticketIdentifier : '';

  return sanitizeTransportFilenameComponent(summary) + ' - ' + isoDate + ticketIdentifierSegment + '.pdf';
}

/**
 * parseIdosTicketText — the IDOS.cz-specific email-BODY parser
 * (quick-260804-bs7, D-01/D-02). UNLIKE the rest of this file, this parses
 * `message.getPlainBody()` directly, never a `.ics` VEVENT — IDOS.cz
 * confirmation emails carry no `.ics` attachment at all.
 *
 * FORMAT (its own shape, needs its own regex): `D.M.YYYY H:MM` — day,
 * month AND hour with NO leading zeros, dot-separated with no spaces
 * between date components. This is DIFFERENT from every other date format
 * already in this codebase (RegioJet's ISO-8601 `.ics` `DTSTART`,
 * enigoo.cz's zero-padded `15.08.2026`, Kino Art's dot-space
 * `7. 8. 2026 17:45`). The route separator is `»` (U+00BB), also new here.
 *
 * PATTERN-ANCHORED, NOT LINE-POSITION (same discipline as
 * parseEnigooTicketText's round-5 rewrite and parseKinoArtTicketText, see
 * their own JSDoc in src/07-action-ticketing-portals.js): ONE combined
 * trip-line regex captures departure date/time, the from-station
 * (non-greedy up to `»`), the to-station (non-greedy up to the arrival
 * date), and arrival date/time — a SINGLE regex (not four separate ones)
 * guarantees all six fields come from the SAME trip line. `\s+`/`\s*`
 * between components absorbs whichever line-separator convention is in
 * play (CRLF vs LF), and the anchors are deliberately independent of the
 * `- `/`* ` bullet marker Gmail's real getPlainBody() rendering may prefix
 * each detail line with (quick-260731-kar round 3's proven noise class) —
 * the trip anchor never looks at line starts at all. The customer-support
 * URL line (which, in a real message, embeds the recipient's own personal
 * mail address as an `email=` query param) is never touched by any anchor
 * here.
 *
 * SCOPE LIMITATION (deliberate, same "don't guess at an unobserved
 * variant" discipline as KINO_ART_KNOWN_VENUE and
 * findKinoArtTicketPdfAttachment, src/07-action-ticketing-portals.js —
 * flagged by the 260804-bs7 review, WR-03): every anchor here uses `.exec`
 * (no `/g`), matching only the FIRST trip line in the body, so an order
 * confirmation bundling more than one e-jízdenka (e.g. a round trip) would
 * have every ticket after the first silently unprocessed — no event, no
 * error, no log line. The real IDOS.cz email this parser was built from,
 * and every fixture in this codebase, show a single-ticket order only; if
 * a genuine multi-ticket IDOS.cz confirmation is ever observed, this needs
 * generalizing THEN, with real data, not guessed now.
 *
 * OPTIONAL anchors, each `null` when absent, NEVER throwing (a missing
 * dedup key or descriptive field must never block event creation): `seats`
 * (the run of characters after the literal `sedadlo` up to the next comma
 * or line end), `eTicketCode` (after the literal `kód e-jízdenky`), and
 * `ticketIdentifier` (D-02: the PURCHASE-SCOPED order code after the
 * literal `kód IDOS.cz` — a character class of uppercase letters, digits
 * and hyphens, which naturally stops at the trailing comma; deliberately
 * NOT the shorter, per-ticket `kód e-jízdenky` value).
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
  // newlines (WR-02, 260804-bs7 review): an unbounded [\s\S]+? pair
  // separated only by the literal "»" measured as polynomial-time
  // (~O(n^2)) against a crafted body, since resolveTransportSender trusts
  // an unauthenticated From header -- a spoofed sender could otherwise
  // stall a shared processEmails() run for seconds on a single message. A
  // real station name is at most a few dozen characters, so this bound is
  // never reached by legitimate input.
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
 * function (quick-260804-bs7), mirroring
 * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL's exact keying convention
 * (src/07-action-ticketing-portals.js). Which registry a matching sender's
 * address resolves against (this one, vs. simply having no entry here at
 * all) is what resolveTransportSenderMode's registry-fallback branch below
 * consults.
 */
const TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL = {
  'jizdenky@idos.svt.cz': parseIdosTicketText,
};

/**
 * resolveTransportSenderMode — decides which processing mode (`'ics'` or
 * `'body'`, D-01) a resolved `sender` (a TRANSPORT_SENDERS entry, or
 * `null`/`undefined`) should be processed through:
 *   1. An explicit `sender.mode` wins outright (`'body'` or `'ics'`).
 *   2. `sender.mode` PRESENT but neither `'body'` nor `'ics'` (a typo, an
 *      unsupported mode, an empty string) THROWS rather than silently
 *      falling through to the registry/`'ics'` fallback below — a review
 *      finding (WR-01, 260804-bs7) caught that treating an invalid value
 *      the same as an ABSENT one reproduces the exact silent-no-op class
 *      this function otherwise exists to prevent: a garbage `mode` on a
 *      future body-sourced carrier with no registered parser would
 *      silently resolve to `'ics'`, find zero `.ics` attachments, and the
 *      thread would be labeled processed with no error and no log line.
 *   3. No `mode` field at all: `'body'` when a parser IS registered for
 *      this sender's address in TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL
 *      — this registry-fallback branch exists SPECIFICALLY to prevent the
 *      silent-no-op class quick-260731-kar round 2 already burned a live
 *      round on (a body-sourced sender silently falling through the
 *      `.ics`-attachment requirement and contributing zero jobs).
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
 * digits. Pure, no GAS globals. Locally reimplemented (not a bare
 * `zeroPad`/`zeroPadTicketComponent`) per this codebase's
 * globally-unique-naming convention — two files declaring the same
 * top-level function name would collide in GAS's shared global scope.
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
 * reimplemented per this codebase's one-file-per-action self-containment
 * convention (see formatWallClockComponentsIso, src/07-action-ticketing-
 * portals.js, for the pattern this mirrors). Pure, no GAS globals.
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
 * buildTransportBodyEntry — builds the SHARED entry shape (D-05:
 * `{ resource, summary, filenameDate, ticketIdentifier, uid }`) for a
 * `'body'`-mode job from `parsedTicket` (parseIdosTicketText's return
 * shape) and an INJECTED `timeZone` (never read from a GAS global here —
 * its GAS-only caller, processTransportTicketJob, resolves it live via
 * `CalendarApp.getCalendarById(calendarId).getTimeZone()`, keeping this
 * function pure and unit-testable, and keeping the live-derived-timezone
 * rule intact — two prior live 2-hour-offset bugs in this codebase were
 * caused by a hardcoded timezone assumption). `resource.description`
 * lists the e-ticket code, the seats and the order code, each line OMITTED
 * entirely — never string-coerced — when its value is falsy, so no
 * description ever embeds the literal 4-character word "null" (the
 * defensive rule established in quick-260731-kar round 4).
 * `entry.filenameDate` is built with `Date.UTC(start.year, start.month,
 * start.day)` used purely as NEUTRAL arithmetic space (never a real
 * instant — the parsed components carry no timezone information) so only
 * the calendar date reaches buildTransportAttachmentFilename.
 * `entry.uid` is always `null` for this mode (IDOS.cz has no VEVENT UID at
 * all), which is what routes this entry through the shared write loop's
 * `Calendar.Events.insert` branch rather than the `.ics`-only idempotent
 * import path. Pure, no GAS globals.
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
  };
}

/**
 * stripTransportSummaryIdentifierPrefix — v0.8.1: RegioJet's real VEVENT
 * SUMMARY leads with the same `#<digits>: ` ticket-number prefix
 * extractTransportTicketIdentifier already extracts as the dedup key (D-02)
 * — e.g. `'#4400574546: Z Vídeň, Hbf, do Brno, hl.n., ...'`. Showing it
 * again in the CALENDAR EVENT'S TITLE is redundant (the identifier is
 * already tracked via extendedProperties.private.ticketIdentifier), so this
 * strips it from the display summary. Only strips the exact observed shape
 * (`#`, one or more digits, `:`, optional whitespace) at the very start;
 * anything else is left untouched, including a summary with no such prefix
 * at all. Pure, no GAS globals.
 */
function stripTransportSummaryIdentifierPrefix(summary) {
  return String(summary || '').replace(/^#\d+:\s*/, '');
}

/**
 * buildTransportIcsEntry — the `'ics'`-mode counterpart to
 * buildTransportBodyEntry, expressing RegioJet's EXISTING behavior through
 * the SAME shared entry shape (D-05) — no new parsing at all: `resource`
 * is the REUSED buildEventResource(event) (D-01), `uid`/`ticketIdentifier`/
 * `filenameDate` are read straight off the already-parsed `event` object
 * (via the EXISTING extractTransportTicketIdentifier for the identifier).
 * `summary` (and `resource.summary`, which otherwise carries the raw VEVENT
 * SUMMARY through unmodified) has RegioJet's redundant `#<digits>: ` prefix
 * stripped via stripTransportSummaryIdentifierPrefix (v0.8.1) — this is the
 * ONE deliberate deviation from "no new parsing"; every other field is
 * still the reused ICS action's own value. Pure, no GAS globals.
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
  };
}

/**
 * resolveTransportProcessingJobs — the pure, TESTABLE extraction of `run`'s
 * per-message orchestration decision, mirroring resolveTicketProcessingJobs'
 * shape (src/07-action-ticketing-portals.js). Given `messages` (an array of
 * message-like objects exposing `getFrom()`/`getAttachments()`/
 * `getPlainBody()` — GAS `GmailMessage` objects in production, plain
 * duck-typed fakes in tests) and `senders` (the TRANSPORT_SENDERS config
 * array), returns an array of jobs, EACH TAGGED WITH A `mode`
 * (quick-260804-bs7 — see this file's class-level "TWO PROCESSING MODES"
 * doc):
 *   - `{ mode: 'body', message, sender }` — EXACTLY ONE JOB PER MATCHING
 *     MESSAGE, decided via resolveTransportSenderMode (D-01). No
 *     `.ics`-attachment requirement at all — a body-sourced sender's event
 *     data is always present in the message body itself.
 *   - `{ mode: 'ics', message, sender, icsAttachments }` — the EXISTING
 *     RegioJet behavior, completely unchanged: collect the message's
 *     `.ics` attachments (via the REUSED isIcsAttachment,
 *     src/05-action-ics-import.js — D-01, no second matcher either); skip
 *     the message if none.
 * A message whose sender does not resolve against `senders` at all
 * contributes NO jobs. EXACTLY ONE JOB PER MATCHING MESSAGE for EITHER
 * mode (never one per attachment for `'ics'` mode either) — deliberate:
 * this action's real duplicate-event guarantee comes from the DEDUP SAFETY
 * NET (isDuplicateTransportTicket) applied per-ENTRY inside
 * processTransportTicketJob, not from restricting which
 * messages/attachments get processed here — the corrected lesson from the
 * ticketing-portals action's own round-8/round-9 double-booking incident
 * (see resolveTicketProcessingJobs' own JSDoc for that full writeup: a
 * message could legitimately carry more than one `.ics` attachment, and
 * every one of them must still be parsed and dedup-checked, just as one
 * job, not silently dropped). Pure, no GAS globals — every GAS-shaped
 * method call here is invoked ON THE PASSED-IN objects only, never a real
 * global service, so this is fully unit-testable under Node with fake
 * message/attachment objects.
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

// Node/GAS environment bridge for parseIcs / buildEventResource / isIcsAttachment
// (defined in the sibling src/05-action-ics-import.js — D-01, this action
// reuses that file's parser rather than hand-rolling a second one) and for
// TRANSPORT_TICKETS_ACTION_CONFIG (defined in the sibling
// src/08-action-cfg-transport-tickets.js — see the 260724-lqi config-split
// refactor for the full load-order/getter rationale this mirrors). Under
// GAS's shared global scope these are ALREADY visible here by bare name —
// no action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare references inside this file's functions/getters would otherwise
// throw ReferenceError. Same `globalThis` bridge technique (not a
// redeclared `const`/`let`/`var`, which would collide under GAS's
// concatenated scope) already established by every other action file's own
// equivalent bridge.
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
 * folder — the SAME folder the ticketing-portals action uses (D-04). This
 * deliberately does NOT reuse getOrCreateDriveFolderByName (src/07-action-
 * ticketing-portals.js), despite the identical implementation: two files
 * declaring the same top-level function name would collide in GAS's single
 * shared global scope. GAS-only (DriveApp) — not unit-tested, proven only
 * by the live checkpoint.
 */
function getOrCreateTransportDriveFolder(name) {
  const existing = DriveApp.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(name);
}

/**
 * findTransportEventByIdentifier — the DEDUP SAFETY NET's lookup (D-02),
 * mirroring findTicketEventByIdentifier's exact query shape (src/07-action-
 * ticketing-portals.js): searches `calendarId` for an existing event
 * already tagged with `extendedProperties.private.ticketIdentifier` equal
 * to `ticketIdentifier` via `Calendar.Events.list(calendarId, {
 * privateExtendedProperty: 'ticketIdentifier=' + ticketIdentifier,
 * singleEvents: true })`. Not paginated/time-windowed — a
 * `privateExtendedProperty` filter against a near-certainly-unique
 * per-ticket number is already an EXACT match expected to return 0 or 1
 * events. Returns the first matching event, or `null` if none found.
 * GAS-only (Calendar global) — not unit-tested, proven only by the live
 * checkpoint.
 */
function findTransportEventByIdentifier(ticketIdentifier, calendarId) {
  const response = Calendar.Events.list(calendarId, {
    privateExtendedProperty: 'ticketIdentifier=' + ticketIdentifier,
    singleEvents: true,
  });
  const items = (response && response.items) || [];
  return items.length > 0 ? items[0] : null;
}

/**
 * isDuplicateTransportTicket — the DEDUP SAFETY NET's decision (D-02),
 * mirroring isDuplicateTicketPurchase's exact shape (src/07-action-
 * ticketing-portals.js). Returns `false` for a falsy `ticketIdentifier`
 * (documented as an accepted per-parse limitation, not a silent gap — a
 * ticket whose identifier could not be extracted simply does not get this
 * protection), otherwise `true` (and logs the same "already exists,
 * skipping (safety-net, not a duplicate path)" message this codebase has
 * used since the ticketing-portals action's own round 8) when
 * findTransportEventByIdentifier finds a match on `calendarId`. GAS-only
 * (calls findTransportEventByIdentifier, which touches the Calendar
 * global) — not unit-tested, proven only by the live checkpoint.
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
 * processTransportTicketJob — the pipeline for ONE job (see
 * resolveTransportProcessingJobs' own JSDoc for the `{ mode, message,
 * sender, icsAttachments? }` shape), in this exact order so nothing is EVER
 * created in Drive before the decision to write is final (see this file's
 * class-level "ORDERING GUARANTEE" doc):
 *   1. Resolve the calendar ONCE via resolveTransportCalendarId(job.sender,
 *      CONFIG.calendarId) and thread it through every downstream call
 *      (never re-read `sender.calendarId` at a call site — see
 *      resolveTransportCalendarId's own JSDoc for the real live-crash class
 *      this avoids).
 *   2. Build `entries` (D-05: the SHARED
 *      `{ resource, summary, filenameDate, ticketIdentifier, uid }` shape)
 *      — this is the ONLY step that differs per mode:
 *        - `mode: 'body'`: look up the registered body parser
 *          (TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL); a body-mode
 *          sender with NO registered parser throws a controlled Error
 *          naming the sender (never a silent no-op — the exact failure
 *          class quick-260731-kar round 2 already burned a live round on),
 *          parse `job.message.getPlainBody()`, and build ONE entry via
 *          buildTransportBodyEntry, using the calendar's LIVE-DERIVED
 *          timezone (`CalendarApp.getCalendarById(calendarId).getTimeZone()`
 *          — never hardcoded, the rule two prior live 2-hour-offset bugs
 *          established).
 *        - `mode: 'ics'`: parse EVERY `.ics` attachment's
 *          `getDataAsString()` through the REUSED parseIcs (D-01) and map
 *          each resulting event through buildTransportIcsEntry — the
 *          EXISTING RegioJet behavior, now expressed through the shared
 *          entry shape, with no new parsing. All parsing completes before
 *          any write begins (fail-closed, same discipline as
 *          ICS_CALENDAR_ACTION.run).
 *      The calendar's timezone lookup and the pure buildEventResource call
 *      now happening here (before the dedup filter below) are reads/pure
 *      calls, not writes — this file's "ORDERING GUARANTEE" doc concerns
 *      Drive/Calendar WRITES specifically, and remains unaffected.
 *   3. Compute each entry's dedup key (`entry.ticketIdentifier`, already
 *      set by whichever builder ran in step 2) and drop the ones
 *      isDuplicateTransportTicket reports as already present — the EXISTING
 *      seenInBatch + isDuplicateTransportTicket filter, kept verbatim. If
 *      nothing remains, return immediately — no Drive upload, no write.
 *   4. If `job.sender.insertPdfIntoEvent` is true: find the ticket PDF
 *      (findTransportTicketPdfAttachment, which excludes invoice.pdf; also
 *      matches the IDOS.cz ticket PDF despite its application/octet-stream
 *      content type, D-03 — no new finder needed); if found, copy its blob
 *      DIRECTLY into
 *      getOrCreateTransportDriveFolder(CONFIG.ticketAttachmentDriveFolderName)
 *      — no temp-folder hop, unlike the ticketing-portals action's OCR
 *      pipeline, since this action never needs the PDF's text — rename it
 *      via buildTransportAttachmentFilename(entriesToCreate[0].summary,
 *      .filenameDate, .ticketIdentifier) BEFORE reading its id/url/name
 *      (the displayed Calendar attachment title is derived from the file's
 *      name AT THAT POINT), and build `{ fileId, fileUrl, title }`. If no
 *      ticket PDF is found, log and continue without one — a missing
 *      attachment must never block the calendar event itself. This block
 *      is EXISTING, kept verbatim, now reading off the shared entry shape
 *      instead of the raw parsed `.ics` event.
 *   5. For each remaining entry: add `resource.extendedProperties = {
 *      private: { ticketIdentifier } }` when the identifier is truthy, and,
 *      when an attachment exists, `resource.attachments = [{ fileId,
 *      fileUrl, title, mimeType: 'application/pdf' }]` plus
 *      `optionalArgs.supportsAttachments = true` (the exact live-proven
 *      shape from createTicketCalendarEvent, src/07-action-ticketing-
 *      portals.js). Logs one line naming the summary, the identifier and
 *      the calendar before writing (the round-2 diagnostic convention
 *      established by the ticketing-portals action). Writes via
 *      importIcsEventWithSequenceRetry(resource, calendarId, entry.uid,
 *      optionalArgs) when the entry carries a `uid` — REUSING
 *      ICS_CALENDAR_ACTION's proven idempotent-by-iCalUID path (D-01, only
 *      ever true for `'ics'`-mode entries), which also protects against
 *      Gmail's own native invite detection creating a second event from
 *      the same .ics — and Calendar.Events.insert(resource, calendarId,
 *      optionalArgs) otherwise (always true for `'body'`-mode entries,
 *      which never carry a UID). This EXISTING write loop is kept
 *      verbatim, now operating on `entry.resource`/`entry.uid` instead of
 *      a raw parsed `.ics` event.
 *
 * FLAGGED ASSUMPTION: passing `supportsAttachments` to
 * `Calendar.Events.import` (as opposed to `Calendar.Events.insert`, where
 * this exact parameter is already live-proven by the ticketing-portals
 * action) is the documented Calendar API v3 parameter shape but has NOT
 * been confirmed against a live call in THIS project for the `import`
 * method specifically — this codebase has been bitten twice already by
 * unverified external API shapes (see src/07-action-ticketing-portals.js's
 * class-level "round 7" doc, and the paragraph-separator saga in rounds
 * 4-5). The live checkpoint for this feature verifies it.
 *
 * GAS-only (DriveApp/CalendarApp/Calendar globals) — not unit-tested,
 * proven only by the live checkpoint; the pure logic it depends on
 * (resolveTransportCalendarId, parseIdosTicketText,
 * resolveTransportSenderMode, buildTransportBodyEntry,
 * buildTransportIcsEntry, extractTransportTicketIdentifier,
 * findTransportTicketPdfAttachment, buildTransportAttachmentFilename,
 * parseIcs, buildEventResource) IS fully unit-tested.
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

  // WR-01 fix (260803-us3 review): a UID-less write goes through the
  // non-idempotent Calendar.Events.insert fallback (see the write loop
  // below), so two same-batch entries sharing a ticketIdentifier would BOTH
  // pass isDuplicateTransportTicket (neither exists in Calendar yet at
  // check time) and both get inserted — a real duplicate. seenInBatch
  // tracks identifiers already claimed earlier in this same pass so a
  // repeat is dropped before it ever reaches isDuplicateTransportTicket's
  // live-state check, independent of whether the entry carries a UID.
  const seenInBatch = {};
  const entriesToCreate = entries.filter(function (entry) {
    if (entry.ticketIdentifier && seenInBatch[entry.ticketIdentifier]) {
      return false;
    }
    if (isDuplicateTransportTicket(entry.ticketIdentifier, calendarId)) {
      return false;
    }
    if (entry.ticketIdentifier) {
      seenInBatch[entry.ticketIdentifier] = true;
    }
    return true;
  });

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

    if (entry.ticketIdentifier) {
      resource.extendedProperties = { private: { ticketIdentifier: entry.ticketIdentifier } };
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
   * carrying at least one `.ics` attachment; for a `'body'`-mode sender
   * (quick-260804-bs7), any matching message at all, no attachment
   * required. This function's own code needs no change for the two-mode
   * architecture — it already covers both modes through
   * resolveTransportProcessingJobs' own mode branching.
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
// Exports every pure helper (transportExtractEmailAddress,
// resolveTransportSender, resolveTransportCalendarId,
// isTransportPdfAttachment, findTransportTicketPdfAttachment,
// extractTransportTicketIdentifier, sanitizeTransportFilenameComponent,
// buildTransportAttachmentFilename, resolveTransportProcessingJobs,
// parseIdosTicketText, resolveTransportSenderMode, buildTransportBodyEntry,
// buildTransportIcsEntry, formatTransportWallClockIso,
// TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL — quick-260804-bs7) and
// TRANSPORT_TICKETS_ACTION (action registry). getOrCreateTransportDriveFolder/
// findTransportEventByIdentifier/isDuplicateTransportTicket/
// processTransportTicketJob remain genuinely GAS-only (reference
// DriveApp/CalendarApp/Calendar globals directly) and are NOT exported —
// they are never invoked under Node.
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
    TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL: TRANSPORT_BODY_PARSERS_BY_IDENTIFYING_EMAIL,
    TRANSPORT_TICKETS_ACTION: TRANSPORT_TICKETS_ACTION,
  };
}

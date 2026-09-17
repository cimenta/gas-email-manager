/**
 * TICKETING_PORTALS_ACTION — detects ticket-purchase confirmation emails
 * (concerts, theater, cinema, events) from configured ticketing portals and
 * creates ONE calendar event per email/purchase, never one per ticket
 * page/seat within that purchase: a purchase with multiple ticket pages for
 * the SAME event must yield exactly one event. A portal's text parser only
 * ever reads the FIRST occurrence of its anchored fields, and every ticket
 * page in a single-purchase PDF repeats the same event name/date/venue
 * (only a trailing per-ticket number differs).
 *
 * Two processing modes, depending on the portal:
 *   - PDF/OCR-sourced (e.g. enigoo.cz): the confirmation email's BODY
 *     carries no usable event data — everything (event name, date/time,
 *     venue) is inside an attached PDF ticket, extracted via Google
 *     Drive's PDF-to-Google-Docs OCR conversion, since Apps Script has no
 *     native PDF text parser.
 *   - Body-sourced (e.g. Kino Art, Ticketmaster CZ): the event data is
 *     parsed directly from the email's plain-text body.
 *
 * DEDUP SAFETY NET: a portal parser extracts a stable `ticketIdentifier`
 * where one exists; the created event is tagged at creation time via
 * `extendedProperties.private.ticketIdentifier`
 * (`processTicketPdfAttachment`), and BEFORE creating a new event,
 * `findTicketEventByIdentifier` searches the resolved calendar for an
 * existing event already carrying that same tag — if found, the run is a
 * silent no-op (temp PDF deleted, no re-upload/re-attach, no second
 * event). This layer checks real calendar state before every write,
 * rather than trying to guess ahead of time which message/attachment to
 * trust as the source. A portal parser that cannot extract any stable
 * identifier (documented per-parser limitation, not a silent gap) simply
 * does not get this protection.
 *
 * ALL PORTAL-SPECIFIC PARSING LOGIC LIVES IN THIS ONE FILE: every parser is
 * a clearly-separated, well-named function/section within THIS file (e.g.
 * parseEnigooTicketText), never a separate `07-portal-*.js` file. A portal
 * is matched to its parser via TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL /
 * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL (below), keyed by the SAME
 * identifyingEmail used to resolve which TICKETING_PORTALS config entry
 * (calendarId, insertPdfIntoEvent) applies.
 *
 * THE DRIVE/OCR PIPELINE (processTicketPdfAttachment, GAS-only —
 * DriveApp/Drive Advanced Service/DocumentApp/CalendarApp/Calendar globals
 * — not directly unit-tested, same category as every other GAS-only
 * function in this codebase):
 *   1. Get the PDF attachment Blob from the Gmail message
 *      (`attachment.copyBlob()`).
 *   2. Create the PDF as a file in a project-owned, fixed-name,
 *      auto-managed TEMP Drive folder (TICKETING_TEMP_DRIVE_FOLDER_NAME,
 *      found-or-created via getOrCreateDriveFolderByName — NOT
 *      user-configurable, no CONFIG field exists for it). EVERY processed
 *      ticket PDF passes through here first, regardless of
 *      insertPdfIntoEvent.
 *   3. Convert it to a Google Doc via the Drive ADVANCED Service
 *      (`Drive.Files.copy` with the target `mimeType` set to Google Docs)
 *      — this triggers Google's OCR/text-extraction pipeline. It works
 *      uniformly whether the source PDF has a real text layer or is a
 *      scanned image.
 *   4. Read the converted Doc's full text
 *      (`DocumentApp.openById(id).getBody().getText()`).
 *   5. Delete the converted Doc (a temp OCR artifact) — ALWAYS, wrapped in
 *      try/finally around the read, so a read failure still cleans up
 *      before its error propagates.
 *   6. If the matched portal's `insertPdfIntoEvent` is true: MOVE (not
 *      copy — avoids a duplicate lingering in the temp folder) the
 *      ORIGINAL uploaded PDF file from the temp folder into the
 *      PERMANENT `CONFIG.ticketAttachmentDriveFolderName` folder (also
 *      found-or-created via getOrCreateDriveFolderByName); its file ID is
 *      kept for the Calendar attachment step below. If `insertPdfIntoEvent`
 *      is false: delete the temp PDF file too — nothing is left in
 *      EITHER Drive folder when this toggle is off.
 *   7. Parse the extracted OCR text via the matched portal's parser
 *      function (TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL) to get
 *      `{ eventName, location, year, month, day, hour, minute }`.
 *   8. Build and insert a Calendar event resource (`summary`, `location`,
 *      `start`/`end`) into the portal's configured `calendarId`. When a
 *      portal's parsed ticket has no explicit end time, a fixed
 *      DEFAULT_EVENT_DURATION_MINUTES (2 hours) is added to the start via
 *      addMinutesToWallClockComponents — a per-portal-parser default to
 *      reach for, not a rule every future portal is forced through. If
 *      `insertPdfIntoEvent` was true, the Drive file is included as a real
 *      Calendar attachment: the resource's `attachments` array
 *      (`[{ fileId, fileUrl, title, mimeType }]`), AND
 *      `Calendar.Events.insert` is called with `{ supportsAttachments: true }`
 *      as its `optionalArgs`.
 *
 * CALENDAR API V3 EVENTATTACHMENT FACTS: `fileUrl` is a REQUIRED field on
 * every `attachments[]` entry per the documented `EventAttachment` schema;
 * `fileId` is actually READ-ONLY on that schema (the server derives it
 * FROM `fileUrl`) — providing `fileId` alone, without `fileUrl`, is not
 * sufficient even though it names a real Drive file (`fileUrl` comes from
 * `DriveApp`'s simple-service `File.getUrl()`). The `events.insert`
 * method's `supportsAttachments` query parameter defaults to `false` and
 * silently drops the `attachments` array rather than erroring when
 * omitted, so it is never optional here when an attachment is present.
 *
 * TIMEZONE: a portal's parsed wall-clock digits are treated as LOCAL time
 * at the venue, resolved live via
 * `CalendarApp.getCalendarById(calendarId).getTimeZone()` (the portal's own
 * configured calendar's timezone), never a hardcoded assumption. Formatted
 * via formatWallClockComponentsIso (deliberately no trailing `Z`/offset,
 * paired with an explicit Calendar API `timeZone` field).
 *
 * OAUTH SCOPES (see src/appsscript.json):
 *   - `https://www.googleapis.com/auth/drive` (the BROADER scope, not the
 *     narrower `drive.file`): `CONFIG.ticketAttachmentDriveFolderName` is
 *     a NAME-based lookup (`DriveApp.getFoldersByName`) deliberately
 *     designed to potentially match a folder the OWNER pre-created by
 *     hand, not necessarily one this script itself created —
 *     `drive.file`'s visibility restriction does not reliably cover that
 *     case.
 *   - `https://www.googleapis.com/auth/script.external_request`, required
 *     by `UrlFetchApp` for the Entradio attachment pipeline (see this
 *     file's "ENTRADIO ATTACHMENT PIPELINE" section). A project that calls
 *     UrlFetchApp without it fails at RUNTIME inside the trigger rather
 *     than at push time — which is why a test pins the manifest.
 *
 * GLOBALLY-UNIQUE NAMING: Apps Script concatenates every project file into
 * ONE shared global scope — every helper in this file is locally
 * reimplemented and namespaced to this action (e.g.
 * `ticketingExtractEmailAddress`, not a bare `extractEmailAddress`) to
 * avoid colliding with same-purpose helpers in other action files.
 */

/**
 * ticketingExtractEmailAddress — extracts the bare, trimmed, lowercased
 * email address from a Gmail "From" header value, or from a bare address
 * with no display name. Pure, no GAS globals. Never throws: a
 * null/undefined/empty input returns ''.
 */
function ticketingExtractEmailAddress(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const angleBracketMatch = /<([^>]*)>/.exec(fromHeader);
  const raw = angleBracketMatch ? angleBracketMatch[1] : fromHeader;

  return raw.trim().toLowerCase();
}

/**
 * resolveTicketingPortal — finds the TICKETING_PORTALS config entry whose
 * `identifyingEmail` case-insensitively matches `fromHeader`'s sender.
 * List order, FIRST match wins; no match (or a null/empty `portals` list)
 * returns `null`, never throws. Pure, no GAS globals.
 */
function resolveTicketingPortal(fromHeader, portals) {
  const list = portals || [];
  const sender = ticketingExtractEmailAddress(fromHeader);

  for (let i = 0; i < list.length; i++) {
    if (ticketingExtractEmailAddress(list[i].identifyingEmail) === sender) {
      return list[i];
    }
  }

  return null;
}

/**
 * resolveTicketingCalendarId — resolves which calendar ID this action's
 * Calendar API calls should target for a given `portal` (a resolved
 * TICKETING_PORTALS entry): `portal.calendarId` when truthy, else
 * `defaultCalendarId`. Pure, no GAS globals.
 */
function resolveTicketingCalendarId(portal, defaultCalendarId) {
  return (portal && portal.calendarId) || defaultCalendarId;
}

/**
 * parseEnigooTicketText — the enigoo.cz-specific ticket-text parser (see
 * this file's class-level JSDoc for why every portal's parser lives here,
 * not in its own file).
 *
 * PATTERN-ANCHORED EXTRACTION, NOT LINE-POSITION: what looks like "each
 * field on its own line" in the Google Docs editor's visual rendering does
 * NOT correspond to actual paragraph boundaries in what `Body.getText()`
 * returns, so extraction is anchored to literal patterns present
 * regardless of which paragraph/line they happen to share with other
 * fields:
 *
 *   1. DATE/TIME ANCHOR: the first `DD.MM.YYYY HH:MM`-shaped numeric
 *      match anywhere in the whole text. A multi-page purchase repeats
 *      this pattern once per page for the SAME event — only the FIRST
 *      occurrence is ever used, per this action's one-event-per-purchase
 *      design.
 *   2. EVENT NAME: everything from the ABSOLUTE START of the text (`^`,
 *      no multiline flag) up to (not including) that first date/time
 *      match — the real text always begins with the event name
 *      immediately followed by the date/time.
 *   3. LOCATION: everything strictly AFTER the date/time match's own end
 *      index, up to (not including) the literal label `Cena/price`.
 *      Searching from that specific end index (not the whole text
 *      blindly) keeps this scoped to THIS ticket occurrence's own price
 *      line, even across a multi-page purchase where the same labels
 *      repeat later.
 *   4. TICKET IDENTIFIER (OPTIONAL — never throws): the first run of
 *      digits found strictly AFTER the literal label `Sleva/discount:`,
 *      searched within the same already-scoped substring used for the
 *      location above. `null` when the label or a following digit run is
 *      not found — a missing dedup key must never block calendar-event
 *      creation, it only means the safety-net dedup check cannot run for
 *      this ticket.
 *
 * `\s`/`[\s\S]` are used throughout (never a line-array split) because
 * Apps Script's Document Service may join paragraphs with `\r`, not `\n`
 * — these character classes absorb `\r`, `\n`, and `\r\n` identically.
 *
 * Returns `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier }` (month zero-indexed, matching Date.UTC's convention
 * and every other date-parsing function in this codebase; `ticketIdentifier`
 * is a string or `null`). Throws a controlled Error if no date/time pattern
 * is found anywhere, if the event name or the "Cena/price"-bounded location
 * cannot be extracted, or if the matched hour/minute are out of range.
 * `ticketIdentifier` alone never causes a throw. Pure, no GAS globals.
 *
 * DIAGNOSTIC-ON-FAILURE: every throw below appends the COMPLETE raw `text`
 * argument (the whole DocumentApp-extracted OCR body text, untruncated)
 * after the specific problem description. This project's existing
 * failure-notification path (notifyOwnerOfFailure/composeFailureBody,
 * src/02-main.js) already emails the owner the full thrown error message
 * on any action failure, so this makes the failure notification itself
 * the diagnostic artifact, with no extra manual step required from the
 * owner. This diagnostic is generic to ANY future ticketing-portal parsing
 * failure, not specific to any one incident.
 */
function parseEnigooTicketText(text) {
  const rawText = String(text || '');

  // Anchor 1: the date/time pattern, first occurrence across the whole
  // text -- a multi-page purchase repeats it once per page for the same
  // event, and only the first occurrence is ever used (one event per
  // purchase).
  const dateTimeMatch = /(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})/.exec(rawText);
  if (!dateTimeMatch) {
    throw new Error('Unrecognized enigoo.cz ticket text: no date/time pattern found. Full extracted text:\n' + rawText);
  }

  const day = Number(dateTimeMatch[1]);
  const month = Number(dateTimeMatch[2]) - 1;
  const year = Number(dateTimeMatch[3]);
  const hour = Number(dateTimeMatch[4]);
  const minute = Number(dateTimeMatch[5]);

  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (0-23) in enigoo.cz ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in enigoo.cz ticket date/time match. Full extracted text:\n' + rawText);
  }

  // Anchor 2: the event name -- everything from the ABSOLUTE START of the
  // text up to the date/time pattern. `[\s\S]+?` (not `.+?`) lets the
  // non-greedy capture traverse a paragraph-separator character between
  // the two fields, regardless of which separator Google's OCR used.
  const eventNameMatch = /^\s*([\s\S]+?)\s+\d{1,2}\.\d{1,2}\.\d{4}\s+\d{1,2}:\d{2}/.exec(rawText);
  if (!eventNameMatch) {
    throw new Error(
      'Unrecognized enigoo.cz ticket text: could not extract the event name preceding the date/time. Full extracted text:\n' + rawText
    );
  }
  const eventName = eventNameMatch[1].trim();

  // Anchor 3: the location -- everything strictly after the date/time
  // match's own end index, up to the literal "Cena/price" label. Slicing
  // from that end index (not the whole text) keeps this scoped to THIS
  // ticket occurrence, even when the same labels repeat later.
  const afterDateTime = rawText.slice(dateTimeMatch.index + dateTimeMatch[0].length);
  const locationMatch = /([\s\S]+?)\s*Cena\/price/.exec(afterDateTime);
  if (!locationMatch) {
    throw new Error(
      'Unrecognized enigoo.cz ticket text: could not find "Cena/price" after the date/time to bound the location. Full extracted text:\n' +
        rawText
    );
  }
  const location = locationMatch[1].trim();

  // Anchor 4 (OPTIONAL -- never throws): the ticket identifier -- first
  // digit run after the literal label "Sleva/discount:", searched within
  // the same scoped `afterDateTime` substring as the location above.
  // `null` when not found; a missing dedup key never blocks event
  // creation.
  const ticketIdentifierMatch = /Sleva\/discount:[\s\S]*?(\d+)/.exec(afterDateTime);
  const ticketIdentifier = ticketIdentifierMatch ? ticketIdentifierMatch[1] : null;

  return {
    eventName: eventName,
    location: location,
    year: year,
    month: month,
    day: day,
    hour: hour,
    minute: minute,
    ticketIdentifier: ticketIdentifier,
  };
}

// KINO_ART_KNOWN_VENUE — the ONLY Kino Art venue/hall string observed in
// real data so far. Scope limitation: if Kino Art ever uses a different
// hall, this needs generalizing THEN, with real data, not guessed now.
const KINO_ART_KNOWN_VENUE = 'Cihlářská - Malý sál';

/**
 * parseKinoArtTicketText — the kinoart.cz-specific ticket-TEXT parser.
 * UNLIKE parseEnigooTicketText above, this parses the email BODY
 * (`message.getPlainBody()`), never a PDF.
 *
 * Extraction anchors (pattern-anchored, same philosophy as
 * parseEnigooTicketText — never a line-position/line-array approach):
 *   1. DATE/TIME: `D. M. YYYY HH:MM` — day/month WITHOUT leading zeros,
 *      dot-SPACE separated (e.g. `7. 8. 2026 17:45`) — a different numeric
 *      date format from enigoo.cz's zero-padded no-space `15.08.2026`,
 *      hence its own distinct regex. First occurrence only (the row
 *      repeats once per seat in a multi-seat purchase).
 *   2. VENUE: the literal `KINO_ART_KNOWN_VENUE` string — always
 *      immediately precedes the date/time in the flattened body text.
 *   3. EVENT (movie) NAME: everything between the LAST occurrence of the
 *      literal column-header word `Cena` (capital C — this
 *      case-SENSITIVE match is what distinguishes it from the lowercase
 *      "cena" inside "Plná cena" appearing later in the same text) that
 *      occurs BEFORE the venue, and the venue string itself.
 *   4. TICKET IDENTIFIER (OPTIONAL — never throws): the ORDER
 *      CONFIRMATION NUMBER, `Potvrzení objednávky č. <digits>`, found near
 *      the top of the body. Unlike enigoo.cz's per-TICKET number, this is
 *      naturally scoped to the WHOLE PURCHASE already (shared across
 *      every seat in a multi-seat purchase) — a BETTER dedup key, not
 *      merely an equivalent one.
 *
 * Returns the same `{ eventName, location, year, month, day, hour,
 * minute, ticketIdentifier }` shape as parseEnigooTicketText (month
 * zero-indexed). Throws a controlled Error if the date/time, venue, or
 * event name cannot be extracted; `ticketIdentifier` alone never causes a
 * throw. Pure, no GAS globals.
 */
function parseKinoArtTicketText(text) {
  const rawText = String(text || '');

  const dateTimeMatch = /(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})\s+(\d{1,2}):(\d{2})/.exec(rawText);
  if (!dateTimeMatch) {
    throw new Error('Unrecognized Kino Art ticket text: no date/time pattern found. Full extracted text:\n' + rawText);
  }

  const day = Number(dateTimeMatch[1]);
  const month = Number(dateTimeMatch[2]) - 1;
  const year = Number(dateTimeMatch[3]);
  const hour = Number(dateTimeMatch[4]);
  const minute = Number(dateTimeMatch[5]);

  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (0-23) in Kino Art ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in Kino Art ticket date/time match. Full extracted text:\n' + rawText);
  }

  const knownVenueIndex = rawText.indexOf(KINO_ART_KNOWN_VENUE);
  if (knownVenueIndex === -1) {
    throw new Error('Unrecognized Kino Art ticket text: known venue string not found. Full extracted text:\n' + rawText);
  }
  const location = KINO_ART_KNOWN_VENUE;

  const cenaIndex = rawText.indexOf('Cena');
  if (cenaIndex === -1 || cenaIndex >= knownVenueIndex) {
    throw new Error(
      'Unrecognized Kino Art ticket text: could not find the event name between the "Cena" column header and the venue. Full extracted text:\n' +
        rawText
    );
  }
  // BULLET-MARKER STRIP: Gmail's real message.getPlainBody() rendering of
  // this email's HTML data rows prefixes each row (the event name is the
  // row's first field) with a literal "* " bullet-list marker. Only that
  // literal marker is stripped -- a different bullet character would need
  // handling then, with real data, not guessed now.
  const eventName = rawText
    .slice(cenaIndex + 'Cena'.length, knownVenueIndex)
    .trim()
    .replace(/^\*\s+/, '');
  if (!eventName) {
    throw new Error('Unrecognized Kino Art ticket text: extracted event name was empty. Full extracted text:\n' + rawText);
  }

  // OPTIONAL (never throws): the order confirmation number, this portal's
  // dedup ticketIdentifier.
  //
  // ROW-BOUNDARY MARKER TOLERANCE: the "Potvrzení objednávky" headline and
  // the "č. <digits>" order number sit in TWO SEPARATE table rows, so
  // Gmail's real plain-text rendering can insert the same "* " marker
  // between them that the eventName bullet-marker strip above handles.
  // `[\s*]*` (whitespace and/or a literal asterisk) tolerates that noise
  // between "objednávky" and "č.".
  const ticketIdentifierMatch = /Potvrzení objednávky[\s*]*č\.\s*(\d+)/.exec(rawText);
  const ticketIdentifier = ticketIdentifierMatch ? ticketIdentifierMatch[1] : null;

  return {
    eventName: eventName,
    location: location,
    year: year,
    month: month,
    day: day,
    hour: hour,
    minute: minute,
    ticketIdentifier: ticketIdentifier,
  };
}

// TICKETMASTER_CZ_MONTH_NAMES: a local month-name-to-number lookup table,
// keyed by lowercased full English month name, mapping to the
// ZERO-INDEXED month number (matching every other date-components object
// in this file) -- Ticketmaster CZ confirmation emails render the date
// with a full English month NAME rather than a numeric month.
const TICKETMASTER_CZ_MONTH_NAMES = {
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

// TICKETMASTER_CZ_ORDER_DETAILS_MARKER — the literal heading opening the
// order-details region. Shared by parseTicketmasterCzTicketText and
// ticketmasterCzTextHasOrderDetails so the parser and the admission detector
// can never drift onto different strings.
const TICKETMASTER_CZ_ORDER_DETAILS_MARKER = 'YOUR ORDER DETAILS';

/**
 * ticketmasterCzNormalizeTicketText — U+00A0 (non-breaking space) -> a regular
 * space. The real body's separator lines are NBSP-only, so every marker search
 * runs against the normalized copy. Shared by the parser and the detector for
 * the same non-drift reason as the marker constant. Pure, never throws.
 */
function ticketmasterCzNormalizeTicketText(text) {
  return String(text || '').replace(/\u00A0/g, ' ');
}

/**
 * ticketmasterCzTextHasOrderDetails — the body-content admission predicate for
 * this portal: does `text` carry an order-details region at all?
 *
 * THE FIDELITY PROPERTY: built on the SAME marker constant and the SAME
 * normalization parseTicketmasterCzTicketText uses, so it returns false for
 * EXACTLY the bodies that parser would throw its marker error on — never an
 * independent heuristic that could drift from it.
 *
 * WHY THIS EXISTS (debug/ticketmaster-cz-order-confirm): Ticketmaster CZ sends
 * at least two templates from one address — the PURCHASE CONFIRMATION (tickets
 * follow in a separate email; carries "ORDER SUMMARY", no order-details region)
 * and the ticket-details email this parser is built for. Config matches
 * addresses, not templates, so only body content can separate them.
 *
 * Returns a literal boolean. Pure, never throws, no GAS globals.
 */
function ticketmasterCzTextHasOrderDetails(text) {
  return ticketmasterCzNormalizeTicketText(text).indexOf(TICKETMASTER_CZ_ORDER_DETAILS_MARKER) !== -1;
}

/**
 * parseTicketmasterCzTicketText — the noreply@ticketmaster.cz-specific
 * ticket-BODY parser. UNLIKE parseEnigooTicketText, and LIKE
 * parseKinoArtTicketText, this parses the email BODY
 * (`message.getPlainBody()`), never a PDF — no Drive upload, no OCR.
 *
 * This portal is registered in
 * TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL (via
 * findTicketmasterCzTicketPdfAttachment, below), the same optional
 * find-the-PDF-and-attach-it-as-is mechanism Kino Art uses — NOT the
 * OCR/PDF-TEXT pipeline (TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL stays
 * untouched). With `insertPdfIntoEvent` true, processTicketFromMessageBody
 * moves the real eTicket.pdf straight into the permanent
 * CONFIG.ticketAttachmentDriveFolderName folder and attaches it — no temp
 * folder, no Drive-to-Docs conversion, exactly Kino Art's own flow.
 *
 * Each field sits on its OWN paragraph in the real body, separated by
 * blank-or-NBSP-only lines (unlike Kino Art's single flattened line).
 * Extraction anchors on two LITERAL markers:
 *   1. `YOUR ORDER DETAILS` — marks the start of the order-details
 *      region. Missing -> controlled throw.
 *   2. `Ticket Quantity:` — marks the end of the order-details region.
 *      Missing, or found at/before the first marker -> controlled throw.
 *   3. DATE/TIME (within the region only): day digits, whitespace, an
 *      alphabetic month NAME, whitespace, 4-digit year, whitespace, the
 *      literal word `at`, whitespace, `HH:MM` — anchored on the day
 *      digits, which naturally skips the leading full weekday name
 *      ("Sunday"). No match -> controlled throw. The matched month name
 *      is resolved through TICKETMASTER_CZ_MONTH_NAMES
 *      case-insensitively; an unrecognized name -> controlled throw
 *      naming the bad value.
 *   4. EVENT NAME / VENUE: the region strictly BEFORE the date/time
 *      match, split on line breaks, each piece trimmed, empty pieces
 *      dropped. The FIRST surviving piece is the event name, the LAST is
 *      the venue. Fewer than two surviving pieces -> controlled throw.
 *
 * `Ticket Quantity: N` is read as the region-terminating marker AND as a
 * captured `ticketQuantity` number for the description — but it is NEVER
 * an event multiplier: quantities of 1, 2 or 5 all yield the exact same
 * eventName/location/date-time/ticketIdentifier, only ticketQuantity and
 * description differ. A non-numeric or missing value leaves
 * `ticketQuantity` `null` and simply omits its line from `description`.
 *
 * `description` reproduces the real order-details block back to the
 * owner as `eventName + '\n\n' + location + '\n\n' + <the real
 * weekday-prefixed date/time line> [+ '\n\n' + 'Ticket Quantity: N' when
 * ticketQuantity is not null]`. The date/time line reused is the ACTUAL
 * regionLines entry the date/time pattern matched against (`dateTimeText`
 * below), not a reconstruction from the parsed numeric components — this
 * deliberately keeps the owner-facing weekday name intact.
 *
 * `ticketIdentifier` is ALWAYS `null` for this portal: no stable
 * per-ticket or per-order confirmation number exists anywhere in the real
 * observed body. Consequence, documented as a v1 limitation: the DEDUP
 * SAFETY NET cannot protect this portal against reprocessing duplicates.
 *
 * Returns the same `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier }` shape as the other two parsers (month zero-indexed)
 * PLUS `ticketQuantity` (number|null) and `description` (string, see
 * above) — backward-compatible, since createTicketCalendarEvent only
 * reads `parsedTicket.description` when truthy. Every controlled throw
 * ends with the FULL raw `text` argument (untruncated), same
 * diagnostic-on-failure convention as the other parsers. Pure, no GAS
 * globals.
 */
function parseTicketmasterCzTicketText(text) {
  const rawText = String(text || '');
  // Only the NORMALIZED working copy is used for extraction; every thrown
  // message below still reports the ORIGINAL rawText. Shared with
  // ticketmasterCzTextHasOrderDetails -- see its JSDoc's FIDELITY PROPERTY.
  const normalizedText = ticketmasterCzNormalizeTicketText(rawText);

  const orderDetailsMarker = TICKETMASTER_CZ_ORDER_DETAILS_MARKER;
  // DEFENSIVE INVARIANT, not a production path: the job-admission gate
  // (TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL) already refused any
  // message failing this exact check, so `run` can no longer reach this throw.
  // Kept so a DIRECT call still fails loudly rather than silently.
  const orderDetailsIndex = normalizedText.indexOf(orderDetailsMarker);
  if (orderDetailsIndex === -1) {
    throw new Error(
      'Unrecognized Ticketmaster CZ ticket text: "YOUR ORDER DETAILS" marker not found. Full extracted text:\n' + rawText
    );
  }

  const ticketQuantityMarker = 'Ticket Quantity:';
  const ticketQuantityIndex = normalizedText.indexOf(ticketQuantityMarker, orderDetailsIndex + orderDetailsMarker.length);
  if (ticketQuantityIndex === -1) {
    throw new Error(
      'Unrecognized Ticketmaster CZ ticket text: "Ticket Quantity:" marker not found. Full extracted text:\n' + rawText
    );
  }

  // ticketQuantity: the digits immediately following the "Ticket
  // Quantity:" marker -- OUTSIDE `region` (which stops right before this
  // marker), so read directly from `normalizedText`. Never throws: a
  // missing or non-numeric value leaves `ticketQuantity` `null` rather
  // than blocking calendar-event creation.
  const ticketQuantityMatch = /^\s*(\d+)/.exec(normalizedText.slice(ticketQuantityIndex + ticketQuantityMarker.length));
  const ticketQuantity = ticketQuantityMatch ? Number(ticketQuantityMatch[1]) : null;

  const region = normalizedText.slice(orderDetailsIndex + orderDetailsMarker.length, ticketQuantityIndex);

  // Each field sits on its OWN paragraph in the region, separated by
  // blank/NBSP-only lines -- split into lines up front, trim each, and
  // drop the empty ones, so the WEEKDAY-PREFIXED date/time line is always
  // treated as ONE whole paragraph, never partially sliced mid-line.
  const regionLines = region
    .split(/\r\n|\r|\n/)
    .map(function (line) {
      return line.trim();
    })
    .filter(function (line) {
      return line.length > 0;
    });

  // Date/time anchor: day digits, alphabetic month NAME, 4-digit year, the
  // literal word "at", HH:MM -- anchored on the day digits, so the leading
  // full weekday name ("Sunday") is naturally skipped without needing to
  // be matched at all.
  const dateTimePattern = /(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\s+at\s+(\d{1,2}):(\d{2})/;
  const dateTimeMatch = dateTimePattern.exec(region);
  if (!dateTimeMatch) {
    throw new Error('Unrecognized Ticketmaster CZ ticket text: no date/time pattern found. Full extracted text:\n' + rawText);
  }

  const day = Number(dateTimeMatch[1]);
  const monthName = dateTimeMatch[2].toLowerCase();
  const year = Number(dateTimeMatch[3]);
  const hour = Number(dateTimeMatch[4]);
  const minute = Number(dateTimeMatch[5]);

  if (!Object.prototype.hasOwnProperty.call(TICKETMASTER_CZ_MONTH_NAMES, monthName)) {
    throw new Error(
      'Unrecognized Ticketmaster CZ ticket text: unrecognized month name "' + dateTimeMatch[2] + '". Full extracted text:\n' + rawText
    );
  }
  const month = TICKETMASTER_CZ_MONTH_NAMES[monthName];

  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (0-23) in Ticketmaster CZ ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in Ticketmaster CZ ticket date/time match. Full extracted text:\n' + rawText);
  }

  // The date/time-bearing PARAGRAPH (identified as the first regionLines
  // entry containing the exact matched date/time substring) marks the end
  // of the event-name/venue candidates -- everything from that paragraph
  // onward (the weekday-prefixed date/time line itself, plus anything
  // after it) is excluded, never treated as a partial trailing fragment.
  const dateTimeLineIndex = regionLines.findIndex(function (line) {
    return line.indexOf(dateTimeMatch[0]) !== -1;
  });
  const candidateLines = dateTimeLineIndex === -1 ? regionLines : regionLines.slice(0, dateTimeLineIndex);

  if (candidateLines.length < 2) {
    throw new Error(
      'Unrecognized Ticketmaster CZ ticket text: could not separate the event name and venue in the order-details region. Full extracted text:\n' +
        rawText
    );
  }

  const eventName = candidateLines[0];
  const location = candidateLines[candidateLines.length - 1];

  // dateTimeText: the ACTUAL regionLines entry the date/time pattern
  // matched against -- preserves the real weekday-prefixed line for
  // `description` below, rather than reconstructing it from the parsed
  // numeric components (which would lose the weekday name, since it is
  // never itself parsed or validated).
  const dateTimeText = dateTimeLineIndex === -1 ? dateTimeMatch[0] : regionLines[dateTimeLineIndex];

  // description (ROUND 2, NEW — see this function's class-level JSDoc):
  // reproduces the real order-details block back to the owner. The
  // "Ticket Quantity: N" line is included only when ticketQuantity was
  // actually captured above.
  const descriptionLines = [eventName, location, dateTimeText];
  if (ticketQuantity !== null) {
    descriptionLines.push('Ticket Quantity: ' + ticketQuantity);
  }

  return {
    eventName: eventName,
    location: location,
    year: year,
    month: month,
    day: day,
    hour: hour,
    minute: minute,
    ticketIdentifier: null,
    ticketQuantity: ticketQuantity,
    description: descriptionLines.join('\n\n'),
  };
}

// ENTRADIO_SECTION_HEADING_PATTERNS: Entradio's plain-text body is
// SECTION-HEADED -- a heading word on its own line, immediately underlined
// by a run of dashes ("Událost" / "Místo konání" / "Vstupenky" / "Platba").
// Anchoring on the heading TOGETHER WITH its dashes underline is not
// decoration: the bare heading words also occur in ordinary prose earlier
// in the body, so a plain substring search would land in the wrong place.
//
// The real underline lengths are NOT uniform, so the pattern requires a
// minimum of three dashes rather than an exact count. `[^\S\r\n]`
// (horizontal whitespace only) is used instead of a bare `\s` so a pattern
// can require "same line" where that matters; `[\r\n]+` between heading
// and underline keeps this separator-agnostic.
//
// Written out as four SEPARATE regex literals rather than built from a
// heading string via `new RegExp(...)`: a dynamically-built pattern must
// double-escape every backslash, which is exactly the kind of silent
// corruption that is invisible on review.
const ENTRADIO_SECTION_HEADING_PATTERNS = {
  event: /Událost[^\S\r\n]*[\r\n]+[^\S\r\n]*-{3,}[^\S\r\n]*(?=[\r\n]|$)/,
  venue: /Místo konání[^\S\r\n]*[\r\n]+[^\S\r\n]*-{3,}[^\S\r\n]*(?=[\r\n]|$)/,
  tickets: /Vstupenky[^\S\r\n]*[\r\n]+[^\S\r\n]*-{3,}[^\S\r\n]*(?=[\r\n]|$)/,
  payment: /Platba[^\S\r\n]*[\r\n]+[^\S\r\n]*-{3,}[^\S\r\n]*(?=[\r\n]|$)/,
};

// ENTRADIO_SEAT_FIELD_PATTERNS — the four per-seat labels Entradio renders
// inside each ticket block. Each value is matched on the SAME LINE as its
// label (`[^\S\r\n]+`, never `\s+`): a label can carry an empty value, and
// a `\s+`-based pattern would jump the blank line and capture the NEXT
// label's value instead. Every field here is optional.
const ENTRADIO_SEAT_FIELD_PATTERNS = [
  { label: 'Poschodí', pattern: /Poschodí[^\S\r\n]+([^\r\n]+)/ },
  { label: 'Sekce', pattern: /Sekce[^\S\r\n]+([^\r\n]+)/ },
  { label: 'Řada', pattern: /Řada[^\S\r\n]+([^\r\n]+)/ },
  { label: 'Místo', pattern: /Místo[^\S\r\n]+([^\r\n]+)/ },
];

// ENTRADIO_BOLD_VALUE_PATTERN — Entradio's plain-text rendering wraps the
// event name and the venue name in literal asterisks (`*ČERNO, VÍR*`), the
// conventional plain-text "bold" marker. Single-line by construction
// (`[^*\r\n]+`), since both real values are single-line.
const ENTRADIO_BOLD_VALUE_PATTERN = /\*([^*\r\n]+)\*/;

// ENTRADIO_DATE_TIME_PATTERN — `D. M. YYYY, HH:MM` (dot-SPACE separated,
// no leading zeros, comma before the time) — its own pattern rather than
// reusing Kino Art's `D. M. YYYY HH:MM`, because the comma does real work
// here: the same "Událost" section also carries a gate-opening line with
// the SAME date but a DIFFERENT time ("...27. 9. 2026, od 17:00 hodin.").
// The literal "od " between the comma and the digits means this pattern
// cannot match that line, so the correct start time is selected
// STRUCTURALLY rather than by ordering luck.
const ENTRADIO_DATE_TIME_PATTERN = /(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4}),?\s*(\d{1,2}):(\d{2})/;

// ENTRADIO_ORDER_NUMBER_PATTERN — the order number, this portal's
// ticketIdentifier: the label "Číslo objednávky" followed by the bold
// digits (`*2354152*`). `[\s*]*` spans the line break and bold markers
// between them. Deliberately NOT "Číslo platby" (the payment number, a
// different label with a different number in the same email).
const ENTRADIO_ORDER_NUMBER_PATTERN = /Číslo objednávky[\s*]*(\d+)/;

/**
 * findEntradioSection — locates one of Entradio's dash-underlined section
 * headings. Returns `{ headingIndex, bodyIndex }` (start of the heading;
 * just past the underline), or `null` when absent. Never throws. Pure, no
 * GAS globals.
 */
function findEntradioSection(text, pattern) {
  const match = pattern.exec(text);
  if (!match) {
    return null;
  }

  return { headingIndex: match.index, bodyIndex: match.index + match[0].length };
}

/**
 * firstEntradioNonEmptyLine — the first line of `text` that is non-empty
 * after trimming, or `''` when there is none. Used for the venue ADDRESS,
 * the first real line after the bold venue name. Splits on `\r\n`/`\r`/`\n`
 * alike. Pure, no GAS globals.
 */
function firstEntradioNonEmptyLine(text) {
  const lines = String(text).split(/\r\n|\r|\n/);

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (trimmed) {
      return trimmed;
    }
  }

  return '';
}

/**
 * dedupeEntradioVenueSegments — collapses repeated comma-separated
 * segments in Entradio's venue name, preserving first-seen order (some
 * venues repeat their own name as the hall name, e.g. "Kino Metropol,
 * Kino Metropol"). Only ever collapses segments that actually repeat, so
 * a normal venue-and-hall pair passes through untouched. Pure, no GAS
 * globals.
 */
function dedupeEntradioVenueSegments(venueName) {
  const seen = {};
  const kept = [];

  String(venueName)
    .split(',')
    .forEach(function (segment) {
      const trimmed = segment.trim();
      if (!trimmed || Object.prototype.hasOwnProperty.call(seen, trimmed)) {
        return;
      }
      seen[trimmed] = true;
      kept.push(trimmed);
    });

  return kept.join(', ');
}

/**
 * findEntradioTicketCodeMatches — the SINGLE scan for per-seat ticket-code
 * lines inside the "Vstupenky" region, returning `[{ code, index }]` in
 * source order (`[]` when there are none). Shared by BOTH
 * extractEntradioTicketLines (needs each match's INDEX to slice out its
 * seat block) and extractEntradioTicketCodes (needs only the codes) — one
 * shared scan means the two views can never disagree about which seats
 * exist.
 *
 * A ticket block starts at a TICKET CODE line — an uppercase alphanumeric
 * run at the start of a line, followed by the literal U+2022 bullet that
 * separates it from the price. The `g`-flagged pattern is declared INSIDE
 * this function deliberately: a regex literal creates a fresh object on
 * every evaluation, so its `lastIndex` can never leak between calls.
 *
 * SCOPE LIMITATION: matching is scoped to the uppercase-and-digits code
 * shape actually observed ("TM5X59GM", "2ZKN9JXVT"). A lowercase or
 * punctuated code would need handling THEN, with real data. Pure, no GAS
 * globals.
 */
function findEntradioTicketCodeMatches(region) {
  const codePattern = /^[^\S\r\n]*([A-Z0-9]{5,})[^\S\r\n]*•/gm;
  const found = [];
  let match;

  while ((match = codePattern.exec(region)) !== null) {
    found.push({ code: match[1], index: match.index });
  }

  return found;
}

/**
 * extractEntradioTicketCodes — the RAW per-seat ticket codes (e.g.
 * `["TM5X59GM", "2ZKN9JXVT"]`) from the "Vstupenky" region, in the email's
 * own order, or `[]` when there are none. Delegates its scan to
 * findEntradioTicketCodeMatches (shared with extractEntradioTicketLines),
 * so the two views cannot disagree about which seats exist.
 *
 * WHAT NEEDS THE RAW CODES: fetchEntradioAttachments builds one QR-code
 * URL per code — these codes ARE the join key between the parsed body and
 * the real Entradio QR endpoint.
 *
 * ALWAYS AN ARRAY, never null. Pure, no GAS globals, never throws.
 */
function extractEntradioTicketCodes(region) {
  return findEntradioTicketCodeMatches(region).map(function (entry) {
    return entry.code;
  });
}

/**
 * extractEntradioTicketLines — renders each per-seat ticket block in the
 * "Vstupenky" section as one human-readable summary line for the calendar
 * event's description, e.g. `"TM5X59GM • Sekce vlevo, Řada 3, Místo 19"`.
 *
 * ENTIRELY OPTIONAL AND NON-THROWING: this drives only `description` and
 * `ticketQuantity`, never the event's identity. An unreadable layout
 * yields an empty array rather than blocking calendar-event creation over
 * a presentational nicety.
 *
 * Seat blocks are located by findEntradioTicketCodeMatches (see its own
 * JSDoc); each block runs from its code line to the NEXT code line, or to
 * the end of the region. Pure, no GAS globals.
 */
function extractEntradioTicketLines(region) {
  const found = findEntradioTicketCodeMatches(region);

  return found.map(function (entry, i) {
    const block = region.slice(entry.index, i + 1 < found.length ? found[i + 1].index : region.length);
    const details = [];

    ENTRADIO_SEAT_FIELD_PATTERNS.forEach(function (field) {
      const valueMatch = field.pattern.exec(block);
      const value = valueMatch ? valueMatch[1].trim() : '';
      // A label with an EMPTY value is real and common ("Poschodí" and
      // "Sleva" are both blank on the real sample) -- such a field is simply
      // omitted rather than rendered as a dangling label.
      if (value) {
        details.push(field.label + ' ' + value);
      }
    });

    return entry.code + (details.length > 0 ? ' • ' + details.join(', ') : '');
  });
}

/**
 * parseEntradioTicketText — the no-reply@app.entradio.cz ticket-BODY
 * parser. Reads `message.getPlainBody()`, never a PDF, and never touches
 * the Drive/OCR pipeline: an Entradio confirmation carries NO ticket file
 * on the message at all (its single PDF attachment is the venue's terms
 * and conditions). Everything the calendar EVENT needs — name, date/time,
 * venue, order number, seats — is already in the body; the real tickets
 * live behind a "STÁHNOUT VSTUPENKY" download link, fetched separately by
 * fetchEntradioAttachments (see the "ENTRADIO ATTACHMENT PIPELINE" section
 * below). The `ticketCodes` field this parser returns exists solely to
 * feed that pipeline. This portal is deliberately ABSENT from
 * TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL: with no ticket PDF
 * among the message's attachments, registering a finder could only ever
 * attach the terms-and-conditions document to the owner's calendar.
 *
 * ENTRADIO IS A PLATFORM, NOT A VENUE: app.entradio.cz is a white-label
 * ticketing system used by many venues (the real sample's own venue name
 * appears only in the body). One portal entry covers all of them, so
 * every anchor below is on Entradio's own TEMPLATE structure (section
 * headings, bold markers, label words), never on any one venue's name —
 * the opposite of the KINO_ART_KNOWN_VENUE approach, deliberately so.
 *
 * Extraction anchors (pattern-anchored, never line-position):
 *   1. SECTIONS: each dash-underlined heading is located ONCE up front
 *      (findEntradioSection), and every subsequent extraction runs
 *      against a REGION bounded by two of them — this is what keeps the
 *      "Místo" SEAT label (inside "Vstupenky") from colliding with the
 *      "Místo konání" VENUE heading. "Událost" and "Místo konání" are
 *      REQUIRED (their regions carry the event's identity); "Vstupenky"
 *      and "Platba" are OPTIONAL bounds used only by the description.
 *   2. EVENT NAME: the first bold value in the "Událost" region.
 *   3. DATE/TIME: the first ENTRADIO_DATE_TIME_PATTERN match strictly
 *      after the event name's own match end (see that pattern's own
 *      comment for why a nearby gate-opening line cannot be matched by
 *      accident).
 *   4. LOCATION: the first bold value in the "Místo konání" region
 *      (de-stuttered via dedupeEntradioVenueSegments), joined to the
 *      first non-empty line after it (the street address). The address
 *      is OPTIONAL — a venue with no address line still yields a usable
 *      location.
 *   5. TICKET IDENTIFIER (OPTIONAL — never throws): the ORDER number,
 *      naturally scoped to the whole PURCHASE and shared by every seat in
 *      a multi-seat order — a 2-seat order still yields ONE event, and
 *      reprocessing it is caught by the shared DEDUP SAFETY NET.
 *   6. TICKET QUANTITY / DESCRIPTION (OPTIONAL): see
 *      extractEntradioTicketLines. `ticketQuantity` is the seat count and
 *      is NEVER an event multiplier.
 *   7. TICKET CODES (OPTIONAL): the RAW per-seat codes from the same
 *      region (extractEntradioTicketCodes). ALWAYS an array, `[]` when
 *      there are no seat blocks — never null.
 *
 * Returns parseTicketmasterCzTicketText's extended shape PLUS
 * `ticketCodes`: `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier, ticketQuantity, ticketCodes, description }` (month
 * zero-indexed). Throws a controlled Error if a REQUIRED section, the
 * event name, the date/time, or the venue cannot be extracted, or if the
 * matched hour/minute are out of range; every throw ends with the FULL
 * raw `text` untruncated, per this file's diagnostic-on-failure
 * convention. Pure, no GAS globals.
 */
function parseEntradioTicketText(text) {
  const rawText = String(text || '');

  const eventSection = findEntradioSection(rawText, ENTRADIO_SECTION_HEADING_PATTERNS.event);
  if (!eventSection) {
    throw new Error('Unrecognized Entradio ticket text: no dash-underlined "Událost" section heading found. Full extracted text:\n' + rawText);
  }

  const venueSection = findEntradioSection(rawText, ENTRADIO_SECTION_HEADING_PATTERNS.venue);
  if (!venueSection) {
    throw new Error(
      'Unrecognized Entradio ticket text: no dash-underlined "Místo konání" section heading found. Full extracted text:\n' + rawText
    );
  }

  // OPTIONAL bounds -- used only to scope the description's seat block, so a
  // missing one degrades the description rather than failing the parse.
  const ticketsSection = findEntradioSection(rawText, ENTRADIO_SECTION_HEADING_PATTERNS.tickets);
  const paymentSection = findEntradioSection(rawText, ENTRADIO_SECTION_HEADING_PATTERNS.payment);

  // --- Event region: name, then date/time ---
  const eventRegion = rawText.slice(eventSection.bodyIndex, venueSection.headingIndex);

  const eventNameMatch = ENTRADIO_BOLD_VALUE_PATTERN.exec(eventRegion);
  if (!eventNameMatch) {
    throw new Error('Unrecognized Entradio ticket text: no bold event name in the "Událost" section. Full extracted text:\n' + rawText);
  }
  const eventName = eventNameMatch[1].trim();

  const afterEventName = eventRegion.slice(eventNameMatch.index + eventNameMatch[0].length);
  const dateTimeMatch = ENTRADIO_DATE_TIME_PATTERN.exec(afterEventName);
  if (!dateTimeMatch) {
    throw new Error(
      'Unrecognized Entradio ticket text: no date/time pattern found after the event name in the "Událost" section. Full extracted text:\n' +
        rawText
    );
  }

  const day = Number(dateTimeMatch[1]);
  const month = Number(dateTimeMatch[2]) - 1;
  const year = Number(dateTimeMatch[3]);
  const hour = Number(dateTimeMatch[4]);
  const minute = Number(dateTimeMatch[5]);

  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (0-23) in Entradio ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in Entradio ticket date/time match. Full extracted text:\n' + rawText);
  }

  // The ACTUAL matched date/time substring, kept for `description` -- the
  // same "reproduce the real line back to the owner rather than rebuild it
  // from the parsed digits" choice parseTicketmasterCzTicketText makes.
  // Interior whitespace is collapsed so a line-broken match still renders on
  // one line.
  const dateTimeText = dateTimeMatch[0].replace(/\s+/g, ' ').trim();

  // --- Venue region: bold venue name + the address line under it ---
  const venueRegion = rawText.slice(venueSection.bodyIndex, ticketsSection ? ticketsSection.headingIndex : rawText.length);

  const venueNameMatch = ENTRADIO_BOLD_VALUE_PATTERN.exec(venueRegion);
  if (!venueNameMatch) {
    throw new Error('Unrecognized Entradio ticket text: no bold venue name in the "Místo konání" section. Full extracted text:\n' + rawText);
  }
  const venueName = dedupeEntradioVenueSegments(venueNameMatch[1].trim());
  const venueAddress = firstEntradioNonEmptyLine(venueRegion.slice(venueNameMatch.index + venueNameMatch[0].length));
  const location = venueAddress ? venueName + ', ' + venueAddress : venueName;

  // --- OPTIONAL: order number (the dedup key) ---
  const orderNumberMatch = ENTRADIO_ORDER_NUMBER_PATTERN.exec(rawText);
  const ticketIdentifier = orderNumberMatch ? orderNumberMatch[1] : null;

  // --- OPTIONAL: per-seat lines (description) and raw codes (QR attachments) ---
  // Both read the SAME region, and both delegate their seat scan to
  // findEntradioTicketCodeMatches, so the description's seat list and the
  // fetched QR codes can never disagree about which seats exist.
  const ticketsRegion = ticketsSection
    ? rawText.slice(ticketsSection.bodyIndex, paymentSection ? paymentSection.headingIndex : rawText.length)
    : '';
  const ticketLines = extractEntradioTicketLines(ticketsRegion);
  const ticketCodes = extractEntradioTicketCodes(ticketsRegion);
  const ticketQuantity = ticketLines.length > 0 ? ticketLines.length : null;

  const descriptionParagraphs = [eventName, location, dateTimeText];
  if (ticketIdentifier) {
    descriptionParagraphs.push('Číslo objednávky: ' + ticketIdentifier);
  }
  if (ticketLines.length > 0) {
    descriptionParagraphs.push('Vstupenky (' + ticketLines.length + '):\n' + ticketLines.join('\n'));
  }

  return {
    eventName: eventName,
    location: location,
    year: year,
    month: month,
    day: day,
    hour: hour,
    minute: minute,
    ticketIdentifier: ticketIdentifier,
    ticketQuantity: ticketQuantity,
    ticketCodes: ticketCodes,
    description: descriptionParagraphs.join('\n\n'),
  };
}

/* ===========================================================================
 * ENTRADIO ATTACHMENT PIPELINE
 * ===========================================================================
 *
 * WHAT GETS ATTACHED:
 *   - THE TICKET FILE behind the "STÁHNOUT VSTUPENKY" button, gated by the
 *     SAME `insertPdfIntoEvent` toggle every other portal uses.
 *   - ONE QR CODE PER SEAT, fetched from Entradio's own
 *     `app.entradio.cz/qrcode` endpoint, each saved and attached as its
 *     OWN file. ALWAYS attempted, NEVER gated by `insertPdfIntoEvent` — a
 *     QR code is not a PDF, and it is the artifact that actually gets the
 *     owner through the door.
 *   - Both land in the EXISTING shared
 *     `CONFIG.ticketAttachmentDriveFolderName` folder (the one enigoo.cz /
 *     Kino Art / Ticketmaster CZ already use). No new folder.
 *
 * NEGATIVE CONTRACT: every function in this section runs BEFORE the
 * Calendar event is created, and none of them may ever throw — an
 * escaping exception here would destroy the event the owner actually
 * needs in exchange for an attachment they could fetch by hand. Every
 * failure is caught, logged, and turned into "one fewer attachment". If
 * NOTHING could be attached at all, the Calendar event is still created
 * and a separate notification email is sent instead
 * (notifyOwnerOfTicketAttachmentFailure, src/02-main.js).
 *
 * TESTABILITY SPLIT, following this file's established convention: the
 * decisions (which URL, which filename, is this response acceptable) are
 * PURE functions with no GAS globals, unit-tested directly. Only the
 * three functions that genuinely touch UrlFetchApp/DriveApp are I/O
 * wrappers, and they are deliberately thin.
 */

// ENTRADIO_TICKET_DOWNLOAD_LINK_PATTERN — the "STÁHNOUT VSTUPENKY" button in
// the message's HTML body.
//
// THE NEAR-MISS THIS PATTERN EXISTS TO AVOID: the very next button in the
// same email is an IDENTICALLY shaped anchor
// (`...>STÁHNOUT JAKO DÁREK</a>`, "download as a gift") pointing at a
// DIFFERENT URL. Anchoring on "STÁHNOUT" alone would fetch the gift
// artifact, and NO response validator downstream could catch it: that
// link also answers 200 with a non-HTML body. The full literal inner text
// is therefore load-bearing, exactly like the comma in
// ENTRADIO_DATE_TIME_PATTERN — the tickets link is selected STRUCTURALLY,
// not by document order.
//
// `[^>]*` is tag-scoped by construction (it cannot cross a `>`), so
// attributes may appear on either side of `href` in any order. The `\s`
// before `href` is deliberate: it stops a hypothetical `data-href="…"`
// from being read as the href.
//
// SCOPE LIMITATION: double-quoted href only, and the inner text directly
// inside the anchor rather than nested in a child element. A variant
// would need handling THEN, with real data.
const ENTRADIO_TICKET_DOWNLOAD_LINK_PATTERN = /<a[^>]*\shref="([^"]*)"[^>]*>\s*STÁHNOUT VSTUPENKY\s*<\/a>/;

/**
 * findEntradioTicketDownloadUrl — returns the "STÁHNOUT VSTUPENKY" href
 * from an Entradio confirmation's HTML body, or `null` when there is none.
 * Runs against `message.getBody()` (HTML), never `getPlainBody()`.
 *
 * `&amp;` in the captured href is decoded back to `&`: an HTML attribute
 * value is REQUIRED to escape a bare ampersand, and an un-decoded one
 * would produce a URL that fetches nothing — a silent failure rather than
 * a visible one.
 *
 * Pure, no GAS globals. Never throws: null/undefined/empty input returns
 * null.
 */
function findEntradioTicketDownloadUrl(htmlBody) {
  const match = ENTRADIO_TICKET_DOWNLOAD_LINK_PATTERN.exec(String(htmlBody || ''));
  if (!match) {
    return null;
  }

  return match[1].replace(/&amp;/g, '&');
}

// ENTRADIO_QR_CODE_URL_PREFIX / _SUFFIX — Entradio's own per-seat QR
// endpoint, matching the codes parseEntradioTicketText extracts into
// `ticketCodes`.
const ENTRADIO_QR_CODE_URL_PREFIX = 'https://app.entradio.cz/qrcode?code=';
const ENTRADIO_QR_CODE_URL_SUFFIX = '&size=200';

// ENTRADIO_QR_CODE_MIME_TYPE — the fallback mimeType recorded on a QR-code
// Calendar attachment when the response carries no usable content-type of its
// own. The endpoint really does answer image/png; this only covers a blank.
const ENTRADIO_QR_CODE_MIME_TYPE = 'image/png';

/**
 * buildEntradioQrCodeUrl — the QR-image URL for ONE ticket code. The code is
 * percent-encoded (`encodeURIComponent`): an unencoded `&` or `=` inside a
 * code would silently truncate the query string and fetch the wrong image.
 * Pure, no GAS globals.
 */
function buildEntradioQrCodeUrl(code) {
  return ENTRADIO_QR_CODE_URL_PREFIX + encodeURIComponent(code) + ENTRADIO_QR_CODE_URL_SUFFIX;
}

/**
 * entradioNormalizedContentType — a response's content-type reduced to its
 * bare lowercased media type: parameters (`; charset=…`) stripped, whitespace
 * trimmed. A missing/empty content-type normalizes to `''`. Pure, never
 * throws.
 */
function entradioNormalizedContentType(contentType) {
  return String(contentType || '')
    .split(';')[0]
    .trim()
    .toLowerCase();
}

/**
 * isEntradioTicketFileResponseAcceptable — the ticket-file download's
 * accept/reject decision, kept pure and separate from the fetch itself.
 *
 * ACCEPT: HTTP 200 with any content-type that is not `text/html`. REJECT:
 * anything else.
 *
 * WHY text/html IS THE REJECTION: a SendGrid click wrapper whose token has
 * expired, or a portal that wants a login, answers 200 with an HTML PAGE.
 * Saving it to Drive and attaching it to the calendar would look like a
 * success and be worthless at the door.
 *
 * The content-type check is deliberately a BLOCKLIST rather than an
 * allowlist: the real format behind this link is unverified (it may be a
 * PDF, a zip, or an image), so rejecting the one known-bad answer beats
 * guessing at the set of good ones. Pure, no GAS globals.
 */
function isEntradioTicketFileResponseAcceptable(responseCode, contentType) {
  if (responseCode !== 200) {
    return false;
  }

  return entradioNormalizedContentType(contentType) !== 'text/html';
}

/**
 * isEntradioQrCodeResponseAcceptable — the QR-code fetch's accept/reject
 * decision. STRICTER than the ticket file's: this endpoint's format is
 * known, so it ACCEPTs only HTTP 200 with an `image/*` content-type and
 * rejects everything else, including an HTML error page served at 200.
 * Pure, no GAS globals.
 */
function isEntradioQrCodeResponseAcceptable(responseCode, contentType) {
  if (responseCode !== 200) {
    return false;
  }

  return entradioNormalizedContentType(contentType).indexOf('image/') === 0;
}

// ENTRADIO_TICKET_FILE_EXTENSIONS_BY_MIME_TYPE — filename extensions for the
// formats a ticket download plausibly returns. Consulted only by
// entradioFileExtensionForMimeType below; an unlisted type yields NO extension
// rather than a guessed one.
const ENTRADIO_TICKET_FILE_EXTENSIONS_BY_MIME_TYPE = {
  'application/pdf': '.pdf',
  'image/png': '.png',
  'image/jpeg': '.jpg',
  'application/zip': '.zip',
};

/**
 * entradioFileExtensionForMimeType — the filename extension for a fetched
 * blob's content-type, or `''` when the type is unknown, missing or
 * generic (`application/octet-stream`). Unlike buildTicketAttachmentFilename
 * (which hardcodes `.pdf` for the other three portals), Entradio's download
 * format is not known in advance, so the extension comes from the response
 * itself — an honest missing suffix beats a confidently wrong one. Pure, no
 * GAS globals, never throws.
 */
function entradioFileExtensionForMimeType(contentType) {
  const normalized = entradioNormalizedContentType(contentType);

  return ENTRADIO_TICKET_FILE_EXTENSIONS_BY_MIME_TYPE[normalized] || '';
}

/**
 * buildEntradioTicketAttachmentFilename — the downloaded ticket file's
 * name in the permanent Drive folder:
 * `"{event name} - {YYYY-MM-DD} - {order number}"` plus the extension
 * derived from the fetched blob's own content-type.
 *
 * DELIBERATE DUPLICATION of buildTicketAttachmentFilename's stem rather
 * than delegation to it, since that function is on the live attachment
 * path of the three already-working portals — a test asserts the two
 * produce the EXACT same name for a PDF, so they cannot drift apart
 * unnoticed.
 *
 * `dateComponents` is any `{ year, month, day }`-shaped object (month
 * zero-indexed). A falsy `ticketIdentifier` omits its segment entirely
 * rather than embedding the literal word "null". Pure, no GAS globals.
 */
function buildEntradioTicketAttachmentFilename(eventName, dateComponents, ticketIdentifier, contentType) {
  const isoDate =
    zeroPadTicketComponent(dateComponents.year, 4) +
    '-' +
    zeroPadTicketComponent(dateComponents.month + 1, 2) +
    '-' +
    zeroPadTicketComponent(dateComponents.day, 2);

  const ticketIdentifierSegment = ticketIdentifier ? ' - ' + ticketIdentifier : '';

  return (
    sanitizeTicketAttachmentFilenameComponent(eventName) +
    ' - ' +
    isoDate +
    ticketIdentifierSegment +
    entradioFileExtensionForMimeType(contentType)
  );
}

/**
 * buildEntradioQrCodeFilename — one QR image per SEAT, named
 * `"{event name} - QR - {ticket code}.png"`. The ticket CODE disambiguates
 * the seats (the only per-seat value guaranteed unique; row/seat numbers
 * repeat across orders). The `.png` extension is fixed rather than
 * derived — unlike the ticket download, this endpoint's format is known.
 * Pure, no GAS globals.
 */
function buildEntradioQrCodeFilename(eventName, code) {
  return (
    sanitizeTicketAttachmentFilenameComponent(eventName) +
    ' - QR - ' +
    sanitizeTicketAttachmentFilenameComponent(code) +
    '.png'
  );
}

/**
 * fetchEntradioResponseBlob — the SINGLE defensive UrlFetchApp call shared
 * by both fetchers below. Fetches `url`, judges the response with the
 * supplied pure predicate, and returns the response blob or `null`.
 *
 * `followRedirects: true` because a SendGrid click wrapper IS a redirect.
 * `muteHttpExceptions: true` because without it a 4xx/5xx THROWS, and this
 * function's entire contract is that it does not.
 *
 * NEVER THROWS. A transport failure, a rejected response, or a malformed
 * response object all become `null` plus one `console.log` line naming
 * the URL. GAS-only (UrlFetchApp), but every DECISION it makes lives in
 * the pure predicates it is handed.
 */
function fetchEntradioResponseBlob(url, isAcceptableResponse, description) {
  try {
    const response = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const blob = response.getBlob();
    const contentType = blob ? blob.getContentType() : '';

    if (!isAcceptableResponse(responseCode, contentType)) {
      console.log(
        'Ticketing portal (Entradio): ' + description + ' fetch REJECTED (HTTP ' + responseCode + ', content-type "' +
          contentType + '") for ' + url
      );
      return null;
    }

    return blob;
  } catch (fetchError) {
    console.log('Ticketing portal (Entradio): ' + description + ' fetch FAILED for ' + url + ': ' + fetchError);
    return null;
  }
}

/**
 * fetchEntradioTicketFileBlob — downloads the real ticket file from the
 * "STÁHNOUT VSTUPENKY" URL. Returns the blob, or `null` on any failure or
 * rejected response (notably an HTML login/error page served at 200).
 * Never throws. GAS-only (via fetchEntradioResponseBlob).
 */
function fetchEntradioTicketFileBlob(url) {
  return fetchEntradioResponseBlob(url, isEntradioTicketFileResponseAcceptable, 'ticket file');
}

/**
 * fetchEntradioQrCodeBlob — downloads ONE seat's QR-code image from Entradio's
 * own qrcode endpoint. Returns the blob, or `null` on any failure or on a
 * non-image response. Never throws. GAS-only (UrlFetchApp, via
 * fetchEntradioResponseBlob).
 */
function fetchEntradioQrCodeBlob(code) {
  return fetchEntradioResponseBlob(buildEntradioQrCodeUrl(code), isEntradioQrCodeResponseAcceptable, 'QR code ' + code);
}

/**
 * entradioSaveBlobAsAttachment — saves one fetched blob into the EXISTING
 * shared permanent Drive folder (`CONFIG.ticketAttachmentDriveFolderName`
 * — the same folder enigoo.cz / Kino Art / Ticketmaster CZ already use),
 * renames it, and returns `{ fileId, fileUrl, title, mimeType }`, or
 * `null` if anything went wrong.
 *
 * The file is RENAMED BEFORE `getName()`/`getUrl()` are read, because a
 * Calendar attachment's displayed title is derived from the Drive file's
 * name AT ATTACH TIME. `getUrl()`/`getName()` are called on the same
 * in-memory File handle rather than re-fetching it by ID.
 *
 * NEVER THROWS. GAS-only (DriveApp via getOrCreateDriveFolderByName).
 */
function entradioSaveBlobAsAttachment(blob, filename, mimeType) {
  try {
    const permanentFolder = getOrCreateDriveFolderByName(CONFIG.ticketAttachmentDriveFolderName);
    const file = permanentFolder.createFile(blob);
    file.setName(filename);

    return { fileId: file.getId(), fileUrl: file.getUrl(), title: file.getName(), mimeType: mimeType };
  } catch (saveError) {
    console.log('Ticketing portal (Entradio): failed to save "' + filename + '" to Drive: ' + saveError);
    return null;
  }
}

/**
 * fetchEntradioAttachments — the Entradio-specific attachment
 * ORCHESTRATOR, registered in
 * TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL and called by
 * processTicketFromMessageBody. Returns an ARRAY of
 * `{ fileId, fileUrl, title, mimeType }` EventAttachment infos — the
 * ticket file first (when fetched), then one QR code per seat in seat
 * order.
 *
 * Flow:
 *   1. IF `portal.insertPdfIntoEvent` — find the "STÁHNOUT VSTUPENKY" link
 *      in the message's HTML body, download it, save it to the shared
 *      permanent folder. Skipped entirely when the toggle is off.
 *   2. UNCONDITIONALLY — for EVERY code in `parsedTicket.ticketCodes`,
 *      fetch that seat's QR image and save it as its own file. Never
 *      gated by `insertPdfIntoEvent`: a QR code is not a PDF, and it is
 *      the artifact that actually gets the owner through the door.
 *
 * THE CONTRACT IS NEGATIVE AND ABSOLUTE: this function ALWAYS returns an
 * array, possibly `[]`, and NEVER throws. Every individual failure is
 * caught, logged, and costs exactly one attachment. This is load-bearing:
 * processTicketFromMessageBody calls this BEFORE creating the Calendar
 * event, so anything escaping here would trade the event the owner
 * actually needs for an attachment they can fetch by hand — when nothing
 * can be attached, the event is still created and a separate notification
 * is sent instead.
 *
 * GAS-only in its I/O (UrlFetchApp/DriveApp/CONFIG), but unit-tested
 * under Node through the same global-injection harness this file's other
 * GAS-only functions use — a negative contract cannot be verified by
 * reading the happy path.
 */
function fetchEntradioAttachments(message, parsedTicket, portal) {
  const attachments = [];

  try {
    if (portal && portal.insertPdfIntoEvent) {
      try {
        const downloadUrl = findEntradioTicketDownloadUrl(message.getBody());

        if (!downloadUrl) {
          console.log(
            'Ticketing portal (Entradio): insertPdfIntoEvent is true but no "STÁHNOUT VSTUPENKY" link was found in the HTML body; skipping the ticket-file download.'
          );
        } else {
          const ticketBlob = fetchEntradioTicketFileBlob(downloadUrl);
          if (ticketBlob) {
            const contentType = ticketBlob.getContentType();
            const saved = entradioSaveBlobAsAttachment(
              ticketBlob,
              buildEntradioTicketAttachmentFilename(
                parsedTicket.eventName,
                parsedTicket,
                parsedTicket.ticketIdentifier,
                contentType
              ),
              contentType
            );
            if (saved) {
              attachments.push(saved);
            }
          }
        }
      } catch (ticketFileError) {
        // message.getBody() itself throwing lands here. One missing
        // attachment, never a failed event.
        console.log('Ticketing portal (Entradio): the ticket-file step failed: ' + ticketFileError);
      }
    }

    const ticketCodes = (parsedTicket && parsedTicket.ticketCodes) || [];
    for (let i = 0; i < ticketCodes.length; i++) {
      try {
        const qrBlob = fetchEntradioQrCodeBlob(ticketCodes[i]);
        if (!qrBlob) {
          continue;
        }

        const saved = entradioSaveBlobAsAttachment(
          qrBlob,
          buildEntradioQrCodeFilename(parsedTicket.eventName, ticketCodes[i]),
          qrBlob.getContentType() || ENTRADIO_QR_CODE_MIME_TYPE
        );
        if (saved) {
          attachments.push(saved);
        }
      } catch (qrError) {
        // Per-seat isolation: one seat's QR failing must never cost the other
        // seats theirs.
        console.log('Ticketing portal (Entradio): the QR-code step failed for ' + ticketCodes[i] + ': ' + qrError);
      }
    }
  } catch (unexpectedError) {
    // The outermost net. Nothing above is expected to reach here; if the
    // parsed-ticket shape is ever something unanticipated, the event still
    // gets created.
    console.log('Ticketing portal (Entradio): attachment fetching failed unexpectedly: ' + unexpectedError);
  }

  return attachments;
}

/**
 * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL — the local (single-file)
 * registry mapping a BODY-SOURCED ticketing portal's `identifyingEmail`
 * (lowercased) to its email-body parser function — the body-sourced
 * counterpart to TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL below
 * (PDF/OCR-sourced). Which registry a portal's sender resolves against
 * determines which processing mode `resolveTicketProcessingJobs`/`run`
 * route it through.
 */
const TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL = {
  'rezervace@kinoart.cz': parseKinoArtTicketText,
  'noreply@ticketmaster.cz': parseTicketmasterCzTicketText,
  'no-reply@app.entradio.cz': parseEntradioTicketText,
};

/**
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL — the local (single-file)
 * registry mapping a ticketing portal's `identifyingEmail` (lowercased)
 * to its OCR-text parser function. This is what lets ALL portal parsers
 * live in one file while still routing a matched TICKETING_PORTALS entry
 * to the RIGHT parser — adding a future portal means adding one new
 * parser function plus one new entry here, nothing else.
 */
const TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL = {
  'no-reply@enigoo.cz': parseEnigooTicketText,
};

/**
 * DEFAULT_EVENT_DURATION_MINUTES — the fixed default duration (2 hours)
 * added to a parsed ticket's start time when a portal's PDF carries no
 * explicit end time. A per-portal-parser default to reach for, not a rule
 * every future portal is forced through — a FUTURE portal whose PDF DOES
 * include an end time should use it directly instead.
 */
const DEFAULT_EVENT_DURATION_MINUTES = 120;

/**
 * addMinutesToWallClockComponents — pure: adds `minutes` to a
 * `{ year, month, day, hour, minute }` wall-clock components object (month
 * zero-indexed), returning a NEW components object of the same shape,
 * correctly handling hour/day/month/year rollover. Implemented via
 * `Date.UTC` arithmetic purely as a NEUTRAL zero-offset calculation space
 * (never a real UTC instant — the input components carry no timezone
 * information at all): build a UTC-labeled millisecond timestamp, add the
 * offset, then re-extract the rolled-over digits via the UTC getters.
 * Pure, no GAS globals.
 */
function addMinutesToWallClockComponents(components, minutes) {
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
 * zeroPadTicketComponent — left-pads `value` with '0' to `length` digits.
 * Pure, no GAS globals. Internal formatting helper for
 * formatWallClockComponentsIso, namespaced per this file's
 * globally-unique-naming convention.
 */
function zeroPadTicketComponent(value, length) {
  return String(value).padStart(length, '0');
}

/**
 * formatWallClockComponentsIso — formats a `{ year, month, day, hour,
 * minute }` wall-clock components object (month zero-indexed) as a
 * zero-padded literal string `'YYYY-MM-DDTHH:MM:00'` — DELIBERATELY with
 * NO trailing `Z` and NO timezone offset, meant to be paired with an
 * explicit Calendar API `timeZone` field (see this file's class-level
 * "TIMEZONE" doc) so the API interprets these digits as wall-clock local
 * time in that zone, not UTC. Pure, no GAS globals.
 */
function formatWallClockComponentsIso(components) {
  return (
    zeroPadTicketComponent(components.year, 4) +
    '-' +
    zeroPadTicketComponent(components.month + 1, 2) +
    '-' +
    zeroPadTicketComponent(components.day, 2) +
    'T' +
    zeroPadTicketComponent(components.hour, 2) +
    ':' +
    zeroPadTicketComponent(components.minute, 2) +
    ':00'
  );
}

/**
 * sanitizeTicketAttachmentFilenameComponent — replaces filesystem-unsafe
 * characters (`/ \ ? % * : | " < >`) with `-` and trims whitespace, so
 * generated filenames stay unambiguous and safe to browse/sort. Pure, no
 * GAS globals.
 */
function sanitizeTicketAttachmentFilenameComponent(value) {
  return String(value)
    .replace(/[/\\?%*:|"<>]/g, '-')
    .trim();
}

/**
 * buildTicketAttachmentFilename — the ATTACHMENT-RENAMING CONVENTION
 * (applies to ALL portals): builds
 * `"{event name} - {YYYY-MM-DD} - {ticket identifier}.pdf"` for the
 * ticket PDF moved into the permanent
 * `CONFIG.ticketAttachmentDriveFolderName` folder. Uses ISO-style
 * `YYYY-MM-DD` (not a locale-specific date format) so files sort
 * consistently when browsing the Drive folder. Since a Calendar event
 * attachment's displayed `title` is derived from the file's name AT
 * ATTACH TIME, renaming the file before it is referenced improves BOTH
 * the Drive folder's browsability AND what shows up on the calendar
 * event. `dateComponents` is any `{ year, month, day }`-shaped object
 * (month zero-indexed) — a full parsed-ticket object works fine as-is.
 * A falsy `ticketIdentifier` omits its segment entirely rather than
 * embedding the literal word "null" — a naive string concatenation would
 * otherwise coerce it into that literal text. Pure, no GAS globals.
 */
function buildTicketAttachmentFilename(eventName, dateComponents, ticketIdentifier) {
  const isoDate =
    zeroPadTicketComponent(dateComponents.year, 4) +
    '-' +
    zeroPadTicketComponent(dateComponents.month + 1, 2) +
    '-' +
    zeroPadTicketComponent(dateComponents.day, 2);

  const ticketIdentifierSegment = ticketIdentifier ? ' - ' + ticketIdentifier : '';

  return sanitizeTicketAttachmentFilenameComponent(eventName) + ' - ' + isoDate + ticketIdentifierSegment + '.pdf';
}

// Node/GAS environment bridge for TICKETING_PORTALS_ACTION_CONFIG (defined
// in the sibling src/07-action-cfg-ticketing-portals.js). Under GAS's
// shared global scope this is ALREADY visible here by bare name -- no
// action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare `TICKETING_PORTALS_ACTION_CONFIG` reference inside
// TICKETING_PORTALS_ACTION's `config` getter below would otherwise throw
// ReferenceError.
if (typeof module !== 'undefined' && module.exports) {
  globalThis.TICKETING_PORTALS_ACTION_CONFIG = require('./07-action-cfg-ticketing-portals.js').TICKETING_PORTALS_ACTION_CONFIG;
}

/**
 * getOrCreateDriveFolderByName — finds a Drive folder by NAME (not ID) via
 * `DriveApp.getFoldersByName`, returning the FIRST match if one or more
 * exist, or creating a new folder via `DriveApp.createFolder` if none
 * exist yet. Used for BOTH the fixed-name auto-managed TEMP folder and
 * the permanent `CONFIG.ticketAttachmentDriveFolderName` folder. GAS-only
 * (DriveApp).
 */
function getOrCreateDriveFolderByName(name) {
  const existing = DriveApp.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(name);
}

// TICKETING_TEMP_DRIVE_FOLDER_NAME — the project-owned TEMP Drive folder's
// FIXED name, auto-managed, NOT user-configurable (no CONFIG field exists
// for it, deliberately — see this file's class-level "Drive/OCR pipeline"
// doc). EVERY processed ticket PDF is uploaded here first for OCR
// extraction, regardless of insertPdfIntoEvent.
const TICKETING_TEMP_DRIVE_FOLDER_NAME = 'GAS Email Manager - Temp';

/**
 * isTicketPdfAttachment — true when `attachment`'s name ends in .pdf
 * (case-insensitive) or its content-type is application/pdf. Shared by
 * findTicketPdfAttachments so the matching rule lives in exactly one
 * place.
 */
function isTicketPdfAttachment(attachment) {
  const name = (attachment.getName() || '').toLowerCase();
  const contentType = attachment.getContentType() || '';

  return name.slice(-4) === '.pdf' || contentType === 'application/pdf';
}

/**
 * findTicketPdfAttachments — every GmailAttachment on `message` whose
 * name ends in .pdf (case-insensitive) or whose content-type is
 * application/pdf (via isTicketPdfAttachment), in source order, or [] if
 * none match.
 */
function findTicketPdfAttachments(message) {
  return message.getAttachments().filter(isTicketPdfAttachment);
}

/**
 * findKinoArtTicketPdfAttachment — Kino Art sends TWO PDF attachments per
 * confirmation email: `Vstupenky.pdf` (the real ticket, one page per
 * seat) and `Doklad.pdf` (a separate receipt/invoice, NOT ticket data —
 * the email body itself says "Tento email není vstupenka", confirming
 * `Vstupenky.pdf` is authoritative). Returns the FIRST qualifying PDF
 * attachment whose name contains `"Vstupenky"`, or `null` — deliberately
 * EXCLUDES `Doklad.pdf`. SCOPE LIMITATION: scoped narrowly to the Czech
 * filename actually observed; a differently-named attachment would need
 * handling THEN, with real data. Pure, no GAS globals.
 */
function findKinoArtTicketPdfAttachment(message) {
  const pdfAttachments = findTicketPdfAttachments(message);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = pdfAttachments[i].getName() || '';
    if (name.indexOf('Vstupenky') !== -1) {
      return pdfAttachments[i];
    }
  }

  return null;
}

/**
 * findTicketmasterCzTicketPdfAttachment — Ticketmaster CZ's own
 * ticket-PDF finder for the OPTIONAL `insertPdfIntoEvent` attachment
 * path. The real observed email carries exactly one PDF attachment,
 * filename `eTicket.pdf`. Returns the FIRST qualifying PDF attachment
 * whose name contains `"eTicket"`, or `null` — same discipline as
 * findKinoArtTicketPdfAttachment above. SCOPE LIMITATION: a
 * differently-named attachment would need handling THEN, with real data.
 * Pure, no GAS globals.
 */
function findTicketmasterCzTicketPdfAttachment(message) {
  const pdfAttachments = findTicketPdfAttachments(message);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = pdfAttachments[i].getName() || '';
    if (name.indexOf('eTicket') !== -1) {
      return pdfAttachments[i];
    }
  }

  return null;
}

/**
 * TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL — the local
 * (single-file) registry mapping a BODY-SOURCED ticketing portal's
 * `identifyingEmail` to its own ticket-PDF-finder function, kept as a
 * SEPARATE registry (rather than hardcoding a finder directly inside
 * processTicketFromMessageBody) so a future body-sourced portal can
 * register its own finder the same way a future PDF-sourced portal
 * registers its own text parser.
 */
// DELIBERATE ABSENCE: there is NO 'no-reply@app.entradio.cz' key here, and
// that is a decision rather than an omission -- an Entradio confirmation
// has no ticket PDF to find. Its only attachment is the venue's terms and
// conditions, so a finder registered here could only ever attach the
// wrong document. A test pins this absence. Entradio's real tickets are
// not ON the message at all, which is why it has an entry in the SEPARATE
// TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL registry
// instead -- see that registry's JSDoc for why fetching and finding are
// kept apart.
const TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL = {
  'rezervace@kinoart.cz': findKinoArtTicketPdfAttachment,
  'noreply@ticketmaster.cz': findTicketmasterCzTicketPdfAttachment,
};

/**
 * TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL — the local
 * (single-file) registry mapping a BODY-SOURCED ticketing portal's
 * `identifyingEmail` to a function that FETCHES its Calendar attachments
 * from somewhere other than the message itself. Signature:
 * `(message, parsedTicket, portal) -> [{ fileId, fileUrl, title, mimeType }]`.
 *
 * DISTINCT FROM TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL above,
 * deliberately a second registry rather than an extension of that one: a
 * "finder" picks the right attachment OFF THE MESSAGE — a pure, offline,
 * always-cheap operation — while a "fetcher" goes out over the NETWORK and
 * needs an extra OAuth scope. Keeping them apart is what makes "which
 * portals make outbound calls" answerable via `Object.keys` on this
 * object. A portal may register in BOTH: processTicketFromMessageBody
 * runs the finder path first, then CONCATENATES this fetcher's results.
 *
 * A fetcher registered here MUST never throw and MUST always return an
 * array: it is called BEFORE the Calendar event is created, and an
 * attachment failure must never be able to cost the owner the event.
 */
const TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL = {
  'no-reply@app.entradio.cz': fetchEntradioAttachments,
};

/**
 * TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL — the local registry
 * mapping a BODY-SOURCED portal's `identifyingEmail` (lowercased) to a
 * never-throwing predicate over `message.getPlainBody()`, keyed the same way as
 * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL. Consulted by BOTH job-admission
 * gates (resolveTicketProcessingJobs and TICKETING_PORTALS_ACTION.appliesTo).
 *
 * WHY (debug/ticketmaster-cz-order-confirm): a sender address identifies a
 * PORTAL, not a TEMPLATE. Ticketmaster CZ sends a purchase confirmation AND a
 * ticket-details email from one address; admitting on sender alone meant the
 * confirmation was parsed as a ticket and threw, emailing the owner a spurious
 * failure notification for entirely routine mail.
 *
 * FAIL-OPEN BY DESIGN, and deliberately UNLIKE the fail-closed equivalent in
 * src/10-action-meetings.js: a portal with NO entry here is admitted exactly as
 * before. This fix is owner-scoped to Ticketmaster CZ; Kino Art and Entradio
 * have unverified variants, and fail-closed would silently stop processing them.
 * Registering a detector is therefore an OPT-IN per portal, pinned by tests so
 * the default stays a decision rather than an accident.
 */
const TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL = {
  'noreply@ticketmaster.cz': ticketmasterCzTextHasOrderDetails,
};

/**
 * ticketBodyLooksProcessable — the shared body-content gate behind both
 * admission points. Returns true when `senderKey` has no registered detector
 * (fail-open, see the registry's JSDoc), otherwise the detector's verdict on
 * `message.getPlainBody()`. Pure, no GAS globals.
 */
function ticketBodyLooksProcessable(message, senderKey) {
  const detector = TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL[senderKey];
  if (!detector) {
    return true;
  }

  return detector(message.getPlainBody()) === true;
}

/**
 * resolveTicketProcessingJobs — the pure, TESTABLE extraction of `run`'s
 * per-message orchestration decision. Given `messages` (an array of
 * message-like objects exposing `getFrom()`/`getAttachments()` — GAS
 * `GmailMessage` objects in production, plain duck-typed fakes in tests)
 * and `portals` (the TICKETING_PORTALS config array), returns an array of
 * processing jobs, EACH TAGGED WITH A `mode`:
 *   - `{ mode: 'pdf', attachment, portal }` — ONE JOB PER QUALIFYING PDF
 *     ATTACHMENT, for a portal whose sender resolves against
 *     TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL (PDF/OCR-sourced, e.g.
 *     enigoo.cz).
 *   - `{ mode: 'body', message, portal }` — EXACTLY ONE JOB PER MATCHING
 *     MESSAGE, for a portal whose sender resolves against
 *     TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL (body-sourced, e.g. Kino
 *     Art) — the ticket data comes from the body itself, once per
 *     message, never per attachment.
 * A message whose sender does not resolve to any configured portal, or
 * whose portal resolves to neither registry, or (for a PDF-sourced
 * portal) carries no qualifying PDF attachment, contributes NO jobs.
 * Pure, no GAS globals — every GAS-shaped method call here is invoked ON
 * THE PASSED-IN objects only, so this is fully unit-testable under Node
 * with fake message/attachment objects.
 *
 * The `mode: 'pdf'` branch processes EVERY qualifying PDF attachment on a
 * message, never just the first. The one guarantee this codebase actually
 * needs — "the SAME purchase never gets a second calendar event" — is
 * provided entirely by the DEDUP SAFETY NET (ticketIdentifier tag +
 * findTicketEventByIdentifier, shared by BOTH processing modes via
 * isDuplicateTicketPurchase), which checks real calendar state before
 * every write and is correct regardless of how many messages/attachments/
 * re-runs ever trigger processing.
 */
function resolveTicketProcessingJobs(messages, portals) {
  const list = messages || [];
  const jobs = [];

  for (let i = 0; i < list.length; i++) {
    const message = list[i];
    const portal = resolveTicketingPortal(message.getFrom(), portals);
    if (!portal) {
      continue;
    }

    const senderKey = ticketingExtractEmailAddress(portal.identifyingEmail);

    if (TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL[senderKey]) {
      const pdfAttachments = findTicketPdfAttachments(message);
      for (let j = 0; j < pdfAttachments.length; j++) {
        jobs.push({ mode: 'pdf', attachment: pdfAttachments[j], portal: portal });
      }
    } else if (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey] && ticketBodyLooksProcessable(message, senderKey)) {
      jobs.push({ mode: 'body', message: message, portal: portal });
    }
  }

  return jobs;
}

/**
 * findTicketEventByIdentifier — the DEDUP SAFETY NET's lookup: searches
 * `calendarId` for an existing event already tagged with
 * `extendedProperties.private.ticketIdentifier` equal to
 * `ticketIdentifier`:
 * `Calendar.Events.list(calendarId, { privateExtendedProperty:
 * 'ticketIdentifier=' + ticketIdentifier, singleEvents: true })`.
 * Deliberately NOT paginated and NOT time-windowed: a
 * `privateExtendedProperty` filter against a near-certainly-unique
 * per-purchase ticket number is already an EXACT match expected to return
 * 0 or 1 events. Returns the first matching event, or `null` if none
 * found. GAS-only (Calendar global) — not unit-tested, proven only by the
 * live checkpoint.
 */
function findTicketEventByIdentifier(ticketIdentifier, calendarId) {
  const response = Calendar.Events.list(calendarId, {
    privateExtendedProperty: 'ticketIdentifier=' + ticketIdentifier,
    singleEvents: true,
  });
  const items = (response && response.items) || [];
  return items.length > 0 ? items[0] : null;
}

/**
 * isDuplicateTicketPurchase — the DEDUP SAFETY NET's SHARED decision,
 * factored out so BOTH processing modes (PDF-sourced
 * `processTicketPdfAttachment` and body-sourced
 * `processTicketFromMessageBody`) use the EXACT same check, never
 * duplicated logic. Returns `true` (and logs "already exists, not a
 * duplicate path") when `ticketIdentifier` is truthy AND an event already
 * carries that exact tag on `calendarId`. Returns `false` when
 * `ticketIdentifier` is falsy (a documented per-parser limitation, not a
 * silent gap) or no matching event is found. GAS-only (calls
 * findTicketEventByIdentifier) — not unit-tested, proven only by the live
 * checkpoint.
 */
function isDuplicateTicketPurchase(ticketIdentifier, calendarId) {
  if (!ticketIdentifier) {
    return false;
  }

  const existingEvent = findTicketEventByIdentifier(ticketIdentifier, calendarId);
  if (existingEvent) {
    console.log(
      'Ticketing portal: event for ticket identifier ' + ticketIdentifier + ' already exists, skipping (safety-net, not a duplicate path).'
    );
    return true;
  }

  return false;
}

/**
 * buildTicketCalendarEventResource — the PURE half of
 * createTicketCalendarEvent: given a `parsedTicket`, an already-resolved
 * `timeZone` string and an ARRAY of attachment infos, returns
 * `{ resource, optionalArgs }` ready to hand to `Calendar.Events.insert`.
 * Touches no GAS global at all.
 *
 * ATTACHMENTS: `[]`, `null` and `undefined` all mean "no attachments" — no
 * `attachments` key is added to the resource and `supportsAttachments` is
 * never set. Each entry keeps its OWN `mimeType` (Entradio's QR codes are
 * `image/png`, its downloaded ticket file's format is not known until
 * fetched). Order is preserved.
 *
 * Two Calendar API v3 facts are load-bearing here: `fileUrl` is REQUIRED
 * on every `attachments[]` entry (`fileId` alone is not sufficient — it
 * is read-only on the EventAttachment schema and the server derives it
 * FROM the URL), and `events.insert` must be called with
 * `supportsAttachments: true` or the whole `attachments` array is
 * SILENTLY dropped rather than erroring.
 *
 * TIMEZONE is passed IN rather than resolved here, which is what keeps
 * this function pure — its caller does the one `CalendarApp` round-trip.
 */
function buildTicketCalendarEventResource(parsedTicket, timeZone, attachments) {
  const startComponents = {
    year: parsedTicket.year,
    month: parsedTicket.month,
    day: parsedTicket.day,
    hour: parsedTicket.hour,
    minute: parsedTicket.minute,
  };
  const endComponents = addMinutesToWallClockComponents(startComponents, DEFAULT_EVENT_DURATION_MINUTES);

  const resource = {
    summary: parsedTicket.eventName,
    location: parsedTicket.location,
    start: { dateTime: formatWallClockComponentsIso(startComponents), timeZone: timeZone },
    end: { dateTime: formatWallClockComponentsIso(endComponents), timeZone: timeZone },
  };

  // description: OPTIONAL, backward-compatible. Only Ticketmaster CZ's and
  // Entradio's parsers set `parsedTicket.description` -- enigoo.cz's and
  // Kino Art's own parsed-ticket objects never carry this field, so this
  // line is a no-op for them, leaving their events' description untouched.
  if (parsedTicket.description) {
    resource.description = parsedTicket.description;
  }

  if (parsedTicket.ticketIdentifier) {
    resource.extendedProperties = { private: { ticketIdentifier: parsedTicket.ticketIdentifier } };
  }

  const attachmentList = attachments || [];
  const optionalArgs = {};

  if (attachmentList.length > 0) {
    resource.attachments = attachmentList.map(function (attachment) {
      return {
        fileId: attachment.fileId,
        fileUrl: attachment.fileUrl,
        title: attachment.title,
        mimeType: attachment.mimeType,
      };
    });
    optionalArgs.supportsAttachments = true;
  }

  return { resource: resource, optionalArgs: optionalArgs };
}

/**
 * createTicketCalendarEvent — the SHARED Calendar event build+insert step,
 * factored out so BOTH processing modes share identical
 * event-shape/tagging/attachment logic, never duplicated. A THIN GAS
 * wrapper around buildTicketCalendarEventResource: the only two things
 * left here are the one live timezone lookup and the insert call.
 *
 * TIMEZONE derived live from the RESOLVED target calendar, never a
 * hardcoded assumption. `attachments` is an ARRAY — every call site
 * passes a list of `{ fileId, fileUrl, title, mimeType }` entries.
 *
 * GAS-only (CalendarApp/Calendar globals) — not unit-tested, proven only
 * by the live checkpoint; everything it decides IS unit-tested, in the
 * pure builder.
 */
function createTicketCalendarEvent(parsedTicket, calendarId, attachments) {
  const timeZone = CalendarApp.getCalendarById(calendarId).getTimeZone();
  const built = buildTicketCalendarEventResource(parsedTicket, timeZone, attachments);

  Calendar.Events.insert(built.resource, calendarId, built.optionalArgs);
}

/**
 * processTicketPdfAttachment — the Drive/OCR/Calendar pipeline for ONE PDF
 * attachment already known to belong to `portal` (a resolved
 * TICKETING_PORTALS config entry). See this file's class-level JSDoc,
 * steps 1-8, for the full flow this function implements end-to-end.
 * GAS-only (DriveApp/Drive Advanced Service/DocumentApp/CalendarApp/
 * Calendar globals) — not unit-tested, proven only by the live checkpoint;
 * the pure logic it depends on (parseEnigooTicketText and friends,
 * addMinutesToWallClockComponents, formatWallClockComponentsIso) IS fully
 * unit-tested.
 *
 * Parsing happens IMMEDIATELY after the OCR read, BEFORE the
 * move-to-permanent-folder-or-delete decision, because the DEDUP SAFETY
 * NET needs `parsedTicket.ticketIdentifier` to decide whether this run
 * should even move/attach the PDF at all. If a duplicate is found, the
 * temp PDF is simply deleted (never moved to the permanent folder, never
 * re-attached) and the function returns before any Calendar write.
 *
 * `uploadedPdfFile`'s fate (moved to the permanent folder, or trashed via
 * the dedup-skip / insertPdfIntoEvent-false paths) is tracked by a
 * `pdfFateResolved` flag set `true` at each of the three points where the
 * PDF's fate is explicitly decided, with a `finally` block wrapping
 * everything from upload through the Calendar API call that trashes the
 * file as a FALLBACK safety net only when that flag is still `false` —
 * this is what guarantees a temp PDF is never left orphaned when
 * `parseTicketText` throws or anything downstream (calendar resolution,
 * the Calendar API call) throws.
 */
function processTicketPdfAttachment(attachment, portal) {
  // Resolved ONCE at the top and threaded explicitly through every
  // downstream Calendar API call site below, rather than each call site
  // independently re-reading `portal.calendarId` -- see
  // resolveTicketingCalendarId's own JSDoc.
  const calendarId = resolveTicketingCalendarId(portal, CONFIG.calendarId);

  const tempFolder = getOrCreateDriveFolderByName(TICKETING_TEMP_DRIVE_FOLDER_NAME);
  const uploadedPdfFile = tempFolder.createFile(attachment.copyBlob());

  // pdfFateResolved -- set true at each of the three points below where
  // uploadedPdfFile's fate is explicitly decided (dedup-skip trash,
  // insertPdfIntoEvent move, insertPdfIntoEvent-false trash). The
  // surrounding try/finally's fallback cleanup only ever acts when this is
  // still false, i.e. something threw before any of those points were
  // reached.
  let pdfFateResolved = false;

  try {
    // Step 3-5: convert to a Google Doc via the Drive ADVANCED Service
    // (triggers Google's OCR pipeline), read its text, then ALWAYS delete
    // the converted Doc afterward — a temp OCR artifact, regardless of
    // whether reading its text succeeds or throws.
    const convertedDoc = Drive.Files.copy(
      { title: uploadedPdfFile.getName() + ' (OCR)', mimeType: MimeType.GOOGLE_DOCS },
      uploadedPdfFile.getId()
    );

    let extractedText;
    try {
      extractedText = DocumentApp.openById(convertedDoc.id).getBody().getText();
    } finally {
      DriveApp.getFileById(convertedDoc.id).setTrashed(true);
    }

    // Step 7 (MOVED UP -- see this function's own class-level doc above):
    // parse the extracted OCR text via the matched portal's parser, BEFORE
    // deciding the PDF's fate, since the dedup safety net immediately below
    // needs parsedTicket.ticketIdentifier.
    const parseTicketText = TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL[ticketingExtractEmailAddress(portal.identifyingEmail)];
    if (!parseTicketText) {
      throw new Error('No ticket-text parser registered for ticketing portal: ' + portal.identifyingEmail);
    }
    const parsedTicket = parseTicketText(extractedText);

    // DEDUP SAFETY NET: if this ticket's parser could extract a stable
    // ticketIdentifier, check whether an event already carries that exact
    // tag before doing anything else -- this is the layer that actually
    // prevents the duplicate, regardless of WHY a second processing
    // attempt occurs.
    if (isDuplicateTicketPurchase(parsedTicket.ticketIdentifier, calendarId)) {
      uploadedPdfFile.setTrashed(true);
      pdfFateResolved = true;
      return;
    }

    // Step 6: move the original PDF to the permanent folder (keeping its
    // file ID for the Calendar attachment below) when insertPdfIntoEvent is
    // true, or delete it from the temp folder entirely when false --
    // nothing is left in EITHER Drive folder when this toggle is off.
    // `attachments` is an array; this path produces AT MOST ONE entry.
    let attachments = [];
    if (portal.insertPdfIntoEvent) {
      const permanentFolder = getOrCreateDriveFolderByName(CONFIG.ticketAttachmentDriveFolderName);
      uploadedPdfFile.moveTo(permanentFolder);
      pdfFateResolved = true;
      // ATTACHMENT-RENAMING CONVENTION (quick-260731-kar, retrofitted onto
      // this already-shipped enigoo.cz path -- see buildTicketAttachmentFilename's
      // own JSDoc): rename BEFORE reading getName()/getUrl() below, since
      // the Calendar event attachment's displayed title is derived from
      // the file's name AT THIS POINT.
      uploadedPdfFile.setName(buildTicketAttachmentFilename(parsedTicket.eventName, parsedTicket, parsedTicket.ticketIdentifier));
      // getUrl()/getName() are called on the SAME in-memory File object
      // already moved (and now renamed) above, not a fresh
      // DriveApp.getFileById lookup -- the object reference remains valid
      // after moveTo()/setName(), so a second Drive round-trip is
      // unnecessary.
      attachments = [
        {
          fileId: uploadedPdfFile.getId(),
          fileUrl: uploadedPdfFile.getUrl(),
          title: uploadedPdfFile.getName(),
          mimeType: 'application/pdf',
        },
      ];
    } else {
      uploadedPdfFile.setTrashed(true);
      pdfFateResolved = true;
    }

    // Step 8: build and insert the Calendar event via the SHARED helper, so
    // both processing modes build the identical event
    // resource/tagging/attachment shape.
    //
    // DIAGNOSTIC VISIBILITY: symmetric log line so a "no event created, no
    // error" outcome is distinguishable in the Executions log the same way
    // for BOTH processing modes.
    console.log(
      'Ticketing portal: creating calendar event for "' + parsedTicket.eventName + '" (ticketIdentifier=' +
        parsedTicket.ticketIdentifier + ') on calendar ' + calendarId + '.'
    );
    createTicketCalendarEvent(parsedTicket, calendarId, attachments);
  } finally {
    if (!pdfFateResolved) {
      // Best-effort fallback cleanup -- do not let a cleanup failure mask
      // whatever original exception is already propagating (parsing
      // errors, missing OAuth scopes, calendar resolution failures, the
      // Calendar API call itself, etc.).
      try {
        uploadedPdfFile.setTrashed(true);
      } catch (cleanupError) {
        console.log('Ticketing portal: failed to clean up orphaned temp PDF ' + uploadedPdfFile.getId() + ': ' + cleanupError);
      }
    }
  }
}

/**
 * processTicketFromMessageBody — the BODY-SOURCED processing mode,
 * alongside the existing PDF/OCR-sourced processTicketPdfAttachment
 * above: for a portal whose event data comes entirely from the email BODY
 * (`message.getPlainBody()`), never from a PDF at all. The Drive/OCR
 * pipeline built for enigoo.cz is NOT used here in any way. Flow:
 *   1. Parse the plain body text via the matched portal's body parser
 *      (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL).
 *   2. DEDUP SAFETY NET (shared with the PDF-sourced mode via
 *      isDuplicateTicketPurchase) — if a duplicate, silent no-op, no
 *      Drive/Calendar writes at all.
 *   3. If `insertPdfIntoEvent` is true: find the portal's own ticket PDF
 *      among the message's attachments (via
 *      TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL), and MOVE IT
 *      DIRECTLY into the PERMANENT CONFIG.ticketAttachmentDriveFolderName
 *      folder — skipping BOTH the temp-folder upload AND the
 *      Drive-to-Docs OCR conversion entirely, since this mode never needs
 *      the PDF's TEXT, only the file itself as an attachment. Renamed via
 *      buildTicketAttachmentFilename, same convention as the PDF-sourced
 *      mode. If no matching PDF attachment is found even though the
 *      toggle is on, the event is still created, just without an
 *      attachment. If `insertPdfIntoEvent` is false: no PDF attachment is
 *      touched or uploaded at all.
 *   3b. If the portal has a registered attachment FETCHER
 *      (TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL — for a
 *      portal whose real ticket does not travel with the message at
 *      all), call it and concatenate its results. That call is
 *      guaranteed not to throw and to return an array, which is what
 *      makes it safe here, BEFORE the event exists. An EMPTY result sends
 *      notifyOwnerOfTicketAttachmentFailure and then carries on creating
 *      the event. Unlike step 3, this step is NOT gated by
 *      `insertPdfIntoEvent`: the fetcher itself decides what that toggle
 *      means for its portal (for Entradio it gates the ticket-file
 *      download only, never the per-seat QR codes).
 *   4. Build and insert the Calendar event via the SHARED
 *      createTicketCalendarEvent, passing the accumulated attachments
 *      ARRAY.
 * GAS-only (GmailMessage/DriveApp/CalendarApp/Calendar globals) — not
 * unit-tested, proven only by the live checkpoint; the pure logic it
 * depends on (parseKinoArtTicketText and friends) IS fully unit-tested.
 */
function processTicketFromMessageBody(message, portal) {
  const calendarId = resolveTicketingCalendarId(portal, CONFIG.calendarId);
  const bodyText = message.getPlainBody();

  const parseTicketBody = TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL[ticketingExtractEmailAddress(portal.identifyingEmail)];
  if (!parseTicketBody) {
    throw new Error('No ticket-body parser registered for ticketing portal: ' + portal.identifyingEmail);
  }
  const parsedTicket = parseTicketBody(bodyText);

  if (isDuplicateTicketPurchase(parsedTicket.ticketIdentifier, calendarId)) {
    return;
  }

  // `attachments` is an ARRAY, since a portal may contribute more than one
  // attachment. The PDF-finder path below still contributes at most one
  // entry, with `mimeType: 'application/pdf'`.
  const attachments = [];

  if (portal.insertPdfIntoEvent) {
    const findPortalTicketPdfAttachment = TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL[ticketingExtractEmailAddress(portal.identifyingEmail)];
    const pdfAttachment = findPortalTicketPdfAttachment ? findPortalTicketPdfAttachment(message) : null;

    if (pdfAttachment) {
      const permanentFolder = getOrCreateDriveFolderByName(CONFIG.ticketAttachmentDriveFolderName);
      const permanentPdfFile = permanentFolder.createFile(pdfAttachment.copyBlob());
      permanentPdfFile.setName(buildTicketAttachmentFilename(parsedTicket.eventName, parsedTicket, parsedTicket.ticketIdentifier));
      attachments.push({
        fileId: permanentPdfFile.getId(),
        fileUrl: permanentPdfFile.getUrl(),
        title: permanentPdfFile.getName(),
        mimeType: 'application/pdf',
      });
    } else {
      console.log(
        'Ticketing portal: insertPdfIntoEvent is true but no matching ticket PDF attachment was found on the message; any portal-specific attachment fetcher still runs below.'
      );
    }
  }

  // PORTAL-SPECIFIC ATTACHMENT FETCHING: a portal whose real ticket does
  // not travel WITH the message can register a fetcher that goes and gets
  // it -- see TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL
  // for why that is a separate registry from the PDF finders above.
  //
  // The fetcher's contract guarantees this call cannot throw and always
  // returns an array, which is what makes it safe to run HERE -- before
  // the Calendar event exists. An attachment failure NEVER blocks event
  // creation: an EMPTY result sends a separate notification email and
  // then carries straight on to create the event. A PARTIAL result is not
  // a failure worth an email.
  const fetchPortalAttachments = TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL[ticketingExtractEmailAddress(portal.identifyingEmail)];
  if (fetchPortalAttachments) {
    const fetchedAttachments = fetchPortalAttachments(message, parsedTicket, portal);

    if (fetchedAttachments.length === 0) {
      console.log(
        'Ticketing portal: no attachments could be fetched for "' + parsedTicket.eventName +
          '"; creating the event anyway and notifying the owner.'
      );
      notifyOwnerOfTicketAttachmentFailure(parsedTicket.eventName, calendarId, parsedTicket.ticketIdentifier);
    }

    for (let i = 0; i < fetchedAttachments.length; i++) {
      attachments.push(fetchedAttachments[i]);
    }
  }

  // DIAGNOSTIC VISIBILITY: this log line, paired with
  // isDuplicateTicketPurchase's existing "already exists, skipping" log,
  // makes it possible to distinguish in the Executions log between "no
  // matching portal/mode upstream" (this code never ran), "correctly
  // no-op'd via the dedup safety net", and "reached this point and called
  // Calendar.Events.insert" for a run that created no visible error.
  console.log(
    'Ticketing portal: creating calendar event for "' + parsedTicket.eventName + '" (ticketIdentifier=' +
      parsedTicket.ticketIdentifier + ') on calendar ' + calendarId + '.'
  );
  createTicketCalendarEvent(parsedTicket, calendarId, attachments);
}

/**
 * TICKETING_PORTALS_ACTION — the ticketing-portals action descriptor.
 * Carries its own config block (TICKETING_PORTALS_ACTION_CONFIG),
 * independent of CONFIG and of any other action's config, except for the
 * one shared cross-cutting CONFIG.ticketAttachmentDriveFolderName field
 * (see src/01-setup.js for why it lives there rather than here).
 */
const TICKETING_PORTALS_ACTION = {
  name: 'ticketing-portals',

  // GETTER, not a plain literal property — see this file's class-level
  // JSDoc and the sibling config file's own "CONFIG SPLIT" note. Not
  // evaluated at object-construction time, only when something reads
  // `.config`, which happens lazily inside function bodies (dispatchActions,
  // notifyOwnerOfFailure) long after every project file has loaded.
  get config() {
    return TICKETING_PORTALS_ACTION_CONFIG;
  },

  /**
   * appliesTo — returns a literal boolean. True when any message on the
   * thread is from a sender matching a configured TICKETING_PORTALS entry
   * AND either: (a) that portal resolves against
   * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL (PDF/OCR-sourced) AND the
   * message carries at least one qualifying PDF attachment, or (b) that
   * portal resolves against TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL
   * (body-sourced) AND the message body passes that portal's registered
   * content detector, if it has one (ticketBodyLooksProcessable). Otherwise
   * false. dispatchActions only skips on a strict `=== false`, so a literal
   * boolean is required.
   *
   * There is still NO attachment requirement for a body-sourced portal, but
   * the content check is not optional: this gate previously assumed a matching
   * sender meant the event data was "always present" in the body, which is
   * what let a Ticketmaster CZ purchase confirmation be claimed and then fail
   * (debug/ticketmaster-cz-order-confirm). appliesTo is a SECOND, independent
   * gate — dispatchActions consults it before run — so it must repeat the
   * check rather than rely on resolveTicketProcessingJobs.
   */
  appliesTo: function (thread) {
    const messages = thread.getMessages();

    for (let i = 0; i < messages.length; i++) {
      const portal = resolveTicketingPortal(messages[i].getFrom(), TICKETING_PORTALS_ACTION.config.ticketingPortals);
      if (!portal) {
        continue;
      }

      const senderKey = ticketingExtractEmailAddress(portal.identifyingEmail);

      if (TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL[senderKey] && findTicketPdfAttachments(messages[i]).length > 0) {
        return true;
      }
      if (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey] && ticketBodyLooksProcessable(messages[i], senderKey)) {
        return true;
      }
    }

    return false;
  },

  /**
   * run — builds the processing job list via the pure, TESTABLE
   * resolveTicketProcessingJobs (each job tagged with `mode: 'pdf'` or
   * `'body'`), then runs the appropriate pipeline exactly once per job:
   * processTicketPdfAttachment for `'pdf'` jobs (the Drive/OCR pipeline),
   * processTicketFromMessageBody for `'body'` jobs (body text only, no
   * Drive/OCR). Duplicate-event protection is provided entirely by the
   * DEDUP SAFETY NET shared by both pipelines (isDuplicateTicketPurchase),
   * not by restricting which attachments/messages get processed here. A
   * message matching neither a configured portal nor either processing
   * mode's requirements contributes no job and is skipped gracefully.
   */
  run: function (thread) {
    const messages = thread.getMessages();
    const jobs = resolveTicketProcessingJobs(messages, TICKETING_PORTALS_ACTION.config.ticketingPortals);

    jobs.forEach(function (job) {
      if (job.mode === 'pdf') {
        processTicketPdfAttachment(job.attachment, job.portal);
      } else if (job.mode === 'body') {
        processTicketFromMessageBody(job.message, job.portal);
      }
    });
  },
};

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Exports every pure function and TICKETING_PORTALS_ACTION. Also exports
// isTicketPdfAttachment/findTicketPdfAttachments/resolveTicketProcessingJobs:
// despite living alongside the GAS-only Drive/OCR/Calendar pipeline, none
// of these three reference a real GAS global directly -- they only invoke
// methods ON THE PASSED-IN message/attachment objects, so they are fully
// testable under Node with plain duck-typed fakes.
// getOrCreateDriveFolderByName/processTicketPdfAttachment/
// findTicketEventByIdentifier remain genuinely GAS-only (reference
// DriveApp/Drive/DocumentApp/CalendarApp/Calendar globals directly) and
// are NOT exported. TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL is
// exported despite living alongside the GAS-only Drive pipeline, since the
// registry object itself references no GAS global -- it is exported so a
// test can prove the ABSENCE of a noreply@ticketmaster.cz key in it, the
// same absence-proving coverage pattern used for
// TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL above.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    ticketingExtractEmailAddress: ticketingExtractEmailAddress,
    resolveTicketingPortal: resolveTicketingPortal,
    resolveTicketingCalendarId: resolveTicketingCalendarId,
    parseEnigooTicketText: parseEnigooTicketText,
    parseKinoArtTicketText: parseKinoArtTicketText,
    parseTicketmasterCzTicketText: parseTicketmasterCzTicketText,
    parseEntradioTicketText: parseEntradioTicketText,
    TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL: TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL,
    TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL: TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL,
    TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL: TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL,
    DEFAULT_EVENT_DURATION_MINUTES: DEFAULT_EVENT_DURATION_MINUTES,
    addMinutesToWallClockComponents: addMinutesToWallClockComponents,
    formatWallClockComponentsIso: formatWallClockComponentsIso,
    buildTicketAttachmentFilename: buildTicketAttachmentFilename,
    TICKETING_TEMP_DRIVE_FOLDER_NAME: TICKETING_TEMP_DRIVE_FOLDER_NAME,
    isTicketPdfAttachment: isTicketPdfAttachment,
    findTicketPdfAttachments: findTicketPdfAttachments,
    findKinoArtTicketPdfAttachment: findKinoArtTicketPdfAttachment,
    findTicketmasterCzTicketPdfAttachment: findTicketmasterCzTicketPdfAttachment,
    resolveTicketProcessingJobs: resolveTicketProcessingJobs,
    TICKETING_PORTALS_ACTION: TICKETING_PORTALS_ACTION,
    // processTicketFromMessageBody IS exported despite being GAS-only
    // (GmailMessage/DriveApp/CalendarApp/Calendar/UrlFetchApp), unlike its
    // sibling processTicketPdfAttachment: its coordination logic (which
    // lines attach fetched files, which notify on total failure) is not
    // itself covered by testing only the pure pieces it calls. It is
    // driven under Node through the same global-injection harness this
    // file's other GAS-only pipelines use.
    processTicketFromMessageBody: processTicketFromMessageBody,
    // The Entradio attachment pipeline. Everything here except
    // fetchEntradioAttachments is pure. fetchEntradioAttachments IS
    // exported despite touching UrlFetchApp/DriveApp/CONFIG, because its
    // contract is a NEGATIVE one -- never throws, always returns an array
    // -- and that cannot be verified by reading the happy path. Its two
    // thin I/O helpers (fetchEntradioResponseBlob/
    // entradioSaveBlobAsAttachment) stay unexported: they are covered
    // through it.
    extractEntradioTicketCodes: extractEntradioTicketCodes,
    findEntradioTicketDownloadUrl: findEntradioTicketDownloadUrl,
    buildEntradioQrCodeUrl: buildEntradioQrCodeUrl,
    isEntradioTicketFileResponseAcceptable: isEntradioTicketFileResponseAcceptable,
    isEntradioQrCodeResponseAcceptable: isEntradioQrCodeResponseAcceptable,
    entradioFileExtensionForMimeType: entradioFileExtensionForMimeType,
    buildEntradioTicketAttachmentFilename: buildEntradioTicketAttachmentFilename,
    buildEntradioQrCodeFilename: buildEntradioQrCodeFilename,
    fetchEntradioAttachments: fetchEntradioAttachments,
    TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL: TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL,
    buildTicketCalendarEventResource: buildTicketCalendarEventResource,
    // The body-content admission gate (debug/ticketmaster-cz-order-confirm).
    // The registry is exported so a test can prove the ABSENCE of Kino Art and
    // Entradio keys, the same absence-proving pattern used above -- that
    // absence IS the owner-scoped boundary of this fix.
    ticketmasterCzTextHasOrderDetails: ticketmasterCzTextHasOrderDetails,
    TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL: TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL,
  };
}

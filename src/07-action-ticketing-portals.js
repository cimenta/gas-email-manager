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
 * Three processing modes, depending on the portal:
 *   - PDF/OCR-sourced (e.g. enigoo.cz): the confirmation email's BODY
 *     carries no usable event data — everything (event name, date/time,
 *     venue) is inside an attached PDF ticket, extracted via Google
 *     Drive's PDF-to-Google-Docs OCR conversion, since Apps Script has no
 *     native PDF text parser.
 *   - Body-sourced (e.g. Kino Art, Ticketmaster CZ): the event data is
 *     parsed directly from the email's plain-text body. When that body is
 *     BLANK (smsticket.cz, quick-261026-ma6: its real text/plain MIME part
 *     is present but empty), the body text instead comes from
 *     ticketingHtmlToPlainText's own rendering of the HTML body, via the
 *     shared ticketingMessageBodyText helper (see this file's "SHARED
 *     BODY-TEXT SOURCE" section) — every portal whose plain body is
 *     non-blank is unaffected. Since quick-261029-pmx, a THIRD extraction
 *     TECHNIQUE exists within this same mode: for GoOut
 *     (info@mg.goout.net), the body input is the RAW HTML
 *     (TICKET_BODY_RAW_HTML_SENDERS / ticketingBodyInputForSender), and
 *     the parser reads structured schema.org EventReservation MICRODATA
 *     through the scope-aware ticketingReadMicrodataItems — structured
 *     data, not literal-marker text and not `.ics`. This is a
 *     TECHNIQUE, not a fourth processing mode: admission, dedup, the PDF
 *     attachment and event creation are all shared, unchanged. See this
 *     file's "SCHEMA.ORG MICRODATA READER" section.
 *   - ICS-sourced (Ticketportal.cz, quick-261005-orv; WebTicket.cz,
 *     quick-261028-h1n): the event data comes from the confirmation
 *     email's OWN `.ics` attachment, parsed by the EXISTING parseIcs
 *     (src/05-action-ics-import.js) — still never a SECOND `.ics` parser.
 *     Each ICS-sourced portal instead registers a small parsed-ticket
 *     BUILDER in TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL
 *     that only post-processes what parseIcs already returned. WebTicket.cz's
 *     own builder (buildWebticketParsedTicket) also reads the message body
 *     for its description, via parseWebticketOrderText — an ENRICHMENT
 *     source layered on top of the `.ics`'s core event fields, never a
 *     replacement for them.
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
 * does not get this protection. For smsticket.cz (quick-261026-ma6) that
 * identifier is the purchase's FIRST entry code (extractSmsticketTicketCodes).
 * WebTicket.cz's (quick-261028-h1n) identifier is the order number read
 * from its ticket PDF's own filename, falling back to the body's own
 * "Zakázka #" number, never the per-email `.ics` UID (WebTicket mints a
 * fresh UID for every sent email, so a re-send would otherwise duplicate
 * the event). GoOut's (quick-261029-pmx) identifier is the order's own
 * `reservationNumber`, identical across its per-ticket microdata blocks,
 * never a per-seat `ticketNumber` (which differs per block).
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
 * paired with an explicit Calendar API `timeZone` field). EXCEPTION
 * (quick-261029-pmx, D-05): GoOut's microdata start is itself
 * offset-qualified, so its parser sets `parsedTicket.utcOffset`, which
 * buildTicketCalendarEventResource appends verbatim to BOTH the start and
 * end `dateTime` strings — a no-op for every other portal, which never
 * sets this field.
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

// KINO_ART_KNOWN_VENUES — every Kino Art venue/hall string observed in real
// data. Kino Art is ONE cinema at ONE address (Cihlářská) with MULTIPLE
// halls ("více sálů"), so the HALL is the part that varies between orders.
//
// This was a single scalar string until 2026-09-21
// (debug/kino-art-venue-not-recognized): every order up to then happened to
// be in Malý sál, and the first Velký sál order (č. 673908) threw
// "known venue string not found" and was labeled failed. The scalar shape
// was a deliberately deferred generalization — the comment here used to say
// "if Kino Art ever uses a different hall, this needs generalizing THEN,
// with real data, not guessed now." That is what this list is; the rule
// still stands for hall #3.
//
// WHY AN ALLOWLIST AND NOT A PATTERN like /Cihlářská - .+ sál/: the venue
// string does DOUBLE DUTY in parseKinoArtTicketText — it is both the
// recognition gate AND the right-hand text anchor that bounds the event name
// (anchor 3 below slices from the "Cena" column header up to the venue's
// index). A loose pattern would move that boundary wherever it happened to
// match and could silently truncate or pad the movie title. Literal strings
// keep the anchor exact.
//
// Each entry is the real, accented, byte-verified string as it appears in
// Gmail's rendering of the confirmation email — plain ASCII hyphen-minus
// (U+002D) between ordinary U+0020 spaces, confirmed by a code-point dump of
// the raw HTML part. Matching is case-sensitive and diacritic-sensitive by
// design; see the tests pinning both.
const KINO_ART_KNOWN_VENUES = [
  'Cihlářská - Malý sál', // observed quick-260731-kar (order č. 900142)
  'Cihlářská - Velký sál', // observed 2026-09-21 (order č. 673908)
];

/**
 * resolveKinoArtKnownVenue — locates which known Kino Art hall appears in
 * the ticket text, and where.
 *
 * Returns `{ venue, index }` for the hall occurring EARLIEST in the text, or
 * `null` when no known hall is present at all.
 *
 * EARLIEST OCCURRENCE, NOT LIST ORDER. The naive
 * `KINO_ART_KNOWN_VENUES.find(v => text.indexOf(v) !== -1)` would return
 * whichever hall is listed first in the array, regardless of where it sits in
 * the text. Since the returned index is what bounds the event name, that
 * would mis-slice the movie title for any purchase whose text mentions two
 * halls. Scanning for the lowest index keeps the anchor on the FIRST data row
 * — the same row every other anchor in this parser reads.
 *
 * TIES BREAK LONGEST-FIRST. Neither current entry is a prefix of the other,
 * so on today's list this branch never decides anything. It is here because
 * the plausible next hall IS an overlapping one — a cinema that already
 * distinguishes Malý from Velký is exactly the kind that later adds a
 * "Cihlářská - Velký sál balkon". At a tie, the shorter entry would match
 * first and leave the tail of the hall name (" balkon") glued onto the event
 * name. Preferring the longest match at the same index prevents that.
 *
 * `knownVenues` is an INJECTION SEAM, not a caller-facing option: production
 * never passes it (the default IS the module list). It exists so the
 * longest-wins branch above can be exercised by a test with a deliberately
 * overlapping list — otherwise that branch is unreachable and untestable on
 * the real data, which is how speculative code rots into being wrong.
 *
 * Pure, no GAS globals.
 */
function resolveKinoArtKnownVenue(rawText, knownVenues) {
  const text = String(rawText || '');
  const venues = knownVenues || KINO_ART_KNOWN_VENUES;
  let best = null;

  for (let i = 0; i < venues.length; i++) {
    const venue = venues[i];
    // An empty entry would indexOf to 0 against ANY text, recognizing every
    // email and slicing the event name to nothing. Skipped defensively here
    // as well as asserted on the list itself in the tests.
    if (!venue) {
      continue;
    }
    const index = text.indexOf(venue);
    if (index === -1) {
      continue;
    }
    if (best === null || index < best.index || (index === best.index && venue.length > best.venue.length)) {
      best = { venue: venue, index: index };
    }
  }

  return best;
}

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
 *   2. VENUE (HALL): whichever literal string from `KINO_ART_KNOWN_VENUES`
 *      occurs EARLIEST in the text (see resolveKinoArtKnownVenue) — always
 *      immediately precedes the date/time in the flattened body text. Kino
 *      Art has several halls at the one address, so this is an allowlist,
 *      not a single string.
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

  // The hall is resolved against KINO_ART_KNOWN_VENUES rather than a single
  // constant (Kino Art has several halls). Fails CLOSED on an unknown hall:
  // without a located hall there is no right-hand boundary for the event
  // name, so there is nothing safe to return and guessing would write a
  // wrong calendar event. The error names the halls we DO know, so the next
  // new hall is diagnosable straight from the failure-notification email.
  const knownVenue = resolveKinoArtKnownVenue(rawText);
  if (!knownVenue) {
    throw new Error(
      'Unrecognized Kino Art ticket text: known venue string not found (known halls: ' +
        KINO_ART_KNOWN_VENUES.join(' | ') +
        '). Full extracted text:\n' +
        rawText
    );
  }
  const knownVenueIndex = knownVenue.index;
  // The hall that ACTUALLY matched — never a fixed constant, which is what
  // this was before multiple halls existed.
  const location = knownVenue.venue;

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
 * the opposite of the KINO_ART_KNOWN_VENUES approach, deliberately so.
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

/* ===========================================================================
 * FEVER (hello@feverup.com) — the FIFTH supported portal, the FOURTH
 * body-sourced one (quick-260921-gj0).
 * ===========================================================================
 *
 * THE STRUCTURAL FACT THAT SHAPES EVERYTHING BELOW: the real sample .eml
 * (269 KB, "Potvrzení nákupu na Fever_ Candlelight...") is multipart/mixed
 * with exactly TWO parts -- one text/html and one application/pdf ticket
 * attachment. There is NO text/plain part at all, unlike every other portal
 * in this file. message.getPlainBody() therefore returns GMAIL'S OWN
 * rendering of that HTML, whose exact line breaking cannot be observed from
 * the .eml and must not be assumed. Every extraction anchor below is
 * therefore a LITERAL MARKER, never a line position, and each of the two
 * required markers is verified to occur exactly once in the real rendered
 * text.
 *
 * The date format is Czech-ABBREVIATED with a 12-hour day-period marker
 * ("so 19 pro - 08:00 odp.") and carries NO YEAR anywhere in the body -- the
 * only four-digit year in the whole message is the footer copyright line.
 * The event year is therefore INFERRED from an injected reference date (the
 * message's own received date, see processTicketFromMessageBody below),
 * never from a live clock read inside the parser.
 */

// FEVER_PURCHASE_DETAILS_MARKER / FEVER_MAP_LINK_MARKER — the two REQUIRED
// literal anchors, each verified to occur exactly once in the real rendered
// text. The greeting prefix "Děkujeme! " is deliberately EXCLUDED from the
// purchase-details marker so a personalized greeting cannot break it.
const FEVER_PURCHASE_DETAILS_MARKER = 'Tady jsou podrobnosti o tvém nákupu';
const FEVER_MAP_LINK_MARKER = 'Zobrazit na mapě';

// FEVER_BUY_AGAIN_LABEL — the "Koupit znova" call-to-action label that can
// share either the event-name region or the venue region with its real
// neighbour, depending on Gmail's unknown line breaking. Skipped by
// feverFirstNonEmptyLine/feverLastNonEmptyLine below so it is never mistaken
// for either field.
const FEVER_BUY_AGAIN_LABEL = 'Koupit znova';

// FEVER_IMAGE_PLACEHOLDER_PATTERN (quick-260921-gj0 round 2, D-18/D-19) —
// Gmail renders an <img> in the plain-text view of an HTML email as a
// bracketed placeholder line of the form "[image: <alt text>]". This was
// OBSERVED LIVE, not assumed: on the owner's first real run, Fever's
// hero/cover image rendered as exactly such a line, sitting between the
// purchase-details marker and the real event-name line, and it became the
// created event's summary verbatim. The bracket-and-prefix shape
// ("[image:...]") is the DIRECTLY OBSERVED part; the tolerance around it --
// case-insensitive, optional whitespace after the opening bracket and around
// the prefix -- is the deliberate ASSUMPTION layered on that one
// observation, so a minor rendering difference in spacing or casing cannot
// re-open this bug. The inner class `[^\]]*` stops at the closing bracket,
// so two placeholders on one line (e.g. a social-icon row) are two separate
// matches rather than one greedy match swallowing the real text between
// them. `g`-flagged and used ONLY via String.replace (feverLineIsOnlyImagePlaceholders
// below) -- never .test/.exec -- so no `lastIndex` state can leak between
// calls, the same discipline FEVER_TICKET_CODE_PATTERN documents for itself.
const FEVER_IMAGE_PLACEHOLDER_PATTERN = /\[\s*image\s*:[^\]]*\]/gi;

// FEVER_SUBJECT_PREFIX (quick-260921-gj0 round 2, D-24) — the fixed template
// text Fever's confirmation-email Subject header starts with, decoded from
// the real sample's own RFC 2047 MIME-encoded-word Subject header. Everything
// after this prefix, verbatim, is the event name -- confirmed against the
// real sample, where the text after the prefix matched the owner's expected
// event name character-for-character, INCLUDING the event name's own
// internal colon. The subject is matched as a PREFIX ONLY (indexOf === 0)
// and NEVER split on a colon: the event name itself may contain one, so a
// colon-split (or "take the Nth segment") read would truncate it -- exactly
// the kind of plausible-looking bug this round exists to prevent. This exact
// prefix is evidenced by ONE real sample -- the same epistemic footing as the
// eleven un-observed entries in FEVER_MONTH_ABBREVIATIONS -- which is why the
// body-sourced fallback (feverFirstNonEmptyLine over the purchase-details/
// map-link region) still exists below rather than this becoming a hard
// requirement.
const FEVER_SUBJECT_PREFIX = 'Potvrzení nákupu na Fever: ';

// FEVER_ORDER_SUMMARY_MARKER — the optional RIGHT bound for the per-seat
// ticket-code scan (see parseFeverTicketText below). Its absence simply
// widens the scan to the end of the text rather than failing the parse.
const FEVER_ORDER_SUMMARY_MARKER = 'Shrnutí objednávky';

// FEVER_DATE_TIME_PATTERN — day digits, whitespace, a month token (any run
// of characters that is neither whitespace, a digit, nor common punctuation
// -- this is what lets Czech diacritics through without a Unicode property
// escape), whitespace, a hyphen/en-dash/em-dash separator, whitespace,
// HH:MM, then an OPTIONAL day-period token restricted to the two literal
// markers this table actually defines ("odp"/"dop", D-05) with an optional
// trailing period. Anchoring on the day digits is what skips the leading
// weekday abbreviation ("so") without needing to match it at all.
const FEVER_DATE_TIME_PATTERN = /(\d{1,2})\s+([^\s\d.,;:!?()'"-]+)\s+[-–—]\s+(\d{1,2}):(\d{2})(?:\s*(odp|dop)\.?)?/i;

// FEVER_TICKET_ID_PATTERN — "ID vstupenky: <digits>", this portal's dedup
// ticketIdentifier (D-09), scoped to the whole purchase.
const FEVER_TICKET_ID_PATTERN = /ID vstupenky:\s*(\d+)/;

// FEVER_TICKET_QUANTITY_PATTERN — "<N> x <label>" (e.g. "5 x Balkon").
// Informational only -- NEVER an event multiplier (D-10).
const FEVER_TICKET_QUANTITY_PATTERN = /(\d+)\s+x\s+([^\r\n]+)/;

// FEVER_TICKET_CODE_PATTERN — a word-bounded run of exactly twenty uppercase
// letters/digits, the per-seat ticket code shape actually observed. `g`
// flagged; parseFeverTicketText resets `lastIndex` to 0 before every scan so
// no state leaks between calls (the same discipline
// findEntradioTicketCodeMatches documents for its own locally-declared
// pattern).
const FEVER_TICKET_CODE_PATTERN = /\b[A-Z0-9]{20}\b/g;

// FEVER_MONTH_ABBREVIATIONS — the twelve Czech abbreviated month names,
// keyed by their DIACRITIC-FOLDED lowercase form, mapping to the
// ZERO-INDEXED month number. Only "pro" (prosinec/December) is DIRECTLY
// OBSERVED in the real sample; the other eleven come from the standard
// Czech (CLDR/ICU) abbreviated-month set, which the sample's own "so"
// weekday abbreviation and "odp." day-period abbreviation confirm Fever is
// formatting with. An unrecognized token is a controlled throw that NAMES
// the token and carries the full raw text, so the owner's failure email
// makes the next fix a one-line table addition.
const FEVER_MONTH_ABBREVIATIONS = {
  led: 0,
  uno: 1,
  bre: 2,
  dub: 3,
  kve: 4,
  cvn: 5,
  cvc: 6,
  srp: 7,
  zar: 8,
  rij: 9,
  lis: 10,
  pro: 11,
};

// FEVER_DIACRITIC_FOLDS — the small explicit character map covering exactly
// the accented characters that occur in the twelve abbreviations above
// (úno, bře, kvě, čvn, čvc, zář, říj). An explicit map is used rather than
// Unicode normalization so the behavior is identical and inspectable under
// both the Apps Script and Node runtimes.
const FEVER_DIACRITIC_FOLDS = {
  ú: 'u',
  ě: 'e',
  ř: 'r',
  č: 'c',
  á: 'a',
  í: 'i',
};

/**
 * feverNormalizeBodyText — pure; returns a copy of `text` with U+00A0
 * (non-breaking space) folded to a plain space, and U+034F, U+00AD, U+200B,
 * U+200C, U+200D, U+2060 and U+FEFF removed entirely. These are the
 * characters the real sample's preheader padding actually contains (U+00A0
 * x293, U+034F x288, U+200C x288, U+00AD x288, grep-verified). The
 * normalized copy is used ONLY for matching inside parseFeverTicketText;
 * every thrown message there still reports the ORIGINAL raw text. Never
 * throws; a null/undefined input normalizes to ''.
 */
function feverNormalizeBodyText(text) {
  return String(text || '')
    .replace(/ /g, ' ')
    .replace(/[͏­​‌‍⁠﻿]/g, '');
}

/**
 * feverLineIsOnlyImagePlaceholders (quick-260921-gj0 round 2, D-19) — pure;
 * `trimmedLine` is assumed ALREADY TRIMMED (the caller does that once, as
 * part of its own loop). Returns true when the line is non-empty AND
 * removing every FEVER_IMAGE_PLACEHOLDER_PATTERN match leaves nothing but
 * whitespace. A line that MIXES a placeholder with real text (e.g. "[image:
 * icon] Some Real Text") is deliberately NOT skipped (D-20) -- this is a
 * narrow skip for lines that are placeholder-only, not a blanket bracket
 * filter. A consequence: a region consisting only of such lines has nothing
 * left to return, which is exactly what drives parseFeverTicketText's
 * EXISTING full-raw-text throw below rather than a wrong summary.
 */
function feverLineIsOnlyImagePlaceholders(trimmedLine) {
  if (!trimmedLine) {
    return false;
  }
  return trimmedLine.replace(FEVER_IMAGE_PLACEHOLDER_PATTERN, '').trim() === '';
}

/**
 * feverFirstNonEmptyLine / feverLastNonEmptyLine — pure; split `text` on
 * carriage-return, newline or both, trim each line, skip empty lines, any
 * line equal to FEVER_BUY_AGAIN_LABEL, and (quick-260921-gj0 round 2, D-19)
 * any line that is ONLY a Gmail image-placeholder per
 * feverLineIsOnlyImagePlaceholders, and return the first/last surviving line
 * (or '' when there is none). These two are what make parseFeverTicketText
 * indifferent to whether a marker shares a line with its neighbour, immune
 * to the "Koupit znova" label landing between fields (D-02/D-03), and immune
 * to a Gmail-rendered image placeholder landing between fields (D-18/D-19).
 * Both still return the ORIGINAL trimmed line verbatim, which is what keeps
 * the parser's normalizedText.indexOf(...) anchors working unchanged.
 */
function feverFirstNonEmptyLine(text) {
  const lines = String(text).split(/\r\n|\r|\n/);

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (trimmed && trimmed !== FEVER_BUY_AGAIN_LABEL && !feverLineIsOnlyImagePlaceholders(trimmed)) {
      return trimmed;
    }
  }

  return '';
}

function feverLastNonEmptyLine(text) {
  const lines = String(text).split(/\r\n|\r|\n/);

  for (let i = lines.length - 1; i >= 0; i--) {
    const trimmed = lines[i].trim();
    if (trimmed && trimmed !== FEVER_BUY_AGAIN_LABEL && !feverLineIsOnlyImagePlaceholders(trimmed)) {
      return trimmed;
    }
  }

  return '';
}

/**
 * feverFoldMonthToken — pure; lowercases `token` and applies
 * FEVER_DIACRITIC_FOLDS character-by-character, so both the accented and
 * the diacritic-stripped spelling of a month abbreviation resolve to the
 * same FEVER_MONTH_ABBREVIATIONS key.
 */
function feverFoldMonthToken(token) {
  return String(token)
    .toLowerCase()
    .split('')
    .map(function (ch) {
      return FEVER_DIACRITIC_FOLDS[ch] || ch;
    })
    .join('');
}

/**
 * feverResolveEventYear — pure; per D-07, the event year is the reference
 * date's own year, UNLESS the parsed (month, day) falls strictly BEFORE the
 * reference date's own (month, day) -- in which case it rolls to the NEXT
 * year. Rests on the buy-before-the-event invariant: a ticket confirmation
 * is always dated before the event it is for. DOCUMENTED LIMITATION: a
 * ticket bought more than a year ahead of its event would resolve to the
 * wrong year; accepted, since Fever events are weeks-to-months out.
 */
function feverResolveEventYear(month, day, referenceDate) {
  const refMonth = referenceDate.getMonth();
  const refDay = referenceDate.getDate();
  const isBeforeReference = month < refMonth || (month === refMonth && day < refDay);

  return isBeforeReference ? referenceDate.getFullYear() + 1 : referenceDate.getFullYear();
}

/**
 * feverEventNameFromSubject (quick-260921-gj0 round 2, D-24/D-25) — pure;
 * NEVER THROWS. Returns the trimmed remainder of `subject` after
 * FEVER_SUBJECT_PREFIX when `subject` is a string that starts with that
 * exact prefix (case-sensitive, matched as a PREFIX via indexOf === 0 --
 * never split on a colon, since the event name itself may carry one) AND
 * something non-whitespace remains after stripping it. Returns `null` for a
 * subject that is missing, not a string, does not start with the prefix, or
 * leaves nothing but whitespace after it -- each of those is the documented
 * FALLBACK condition to the body-sourced read, not an error.
 */
function feverEventNameFromSubject(subject) {
  if (typeof subject !== 'string' || subject.indexOf(FEVER_SUBJECT_PREFIX) !== 0) {
    return null;
  }
  const remainder = subject.slice(FEVER_SUBJECT_PREFIX.length).trim();
  return remainder || null;
}

/**
 * parseFeverTicketText — the hello@feverup.com ticket-BODY parser.
 *
 * THE REAL RENDERED TEXT (fictionalized values substituted per this file's
 * fixture convention -- see the test suite):
 *
 *     Zobrazit v prohlížeči
 *     Děkujeme! Tady jsou podrobnosti o tvém nákupu
 *     <event name>                                  <-- event name
 *     Koupit znova
 *     <venue> - <address>                           <-- location
 *     Zobrazit na mapě
 *     so 19 pro - 08:00 odp.                         <-- date/time
 *     Změnit datum nebo čas
 *     ...
 *     5 x Balkon
 *     ID vstupenky: <digits>
 *     <20-char code> (one per seat)
 *     ...
 *     Shrnutí objednávky
 *
 * The real message has NO text/plain part at all (see this section's
 * class-level doc), so this parser reads Gmail's own rendering of the HTML
 * and its line breaking cannot be assumed -- precisely why every anchor
 * below is a literal marker with first/last-non-empty-line reads on either
 * side of it, never a line-position read.
 *
 * Extraction anchors, in order:
 *   1. PURCHASE-DETAILS MARKER (required) -- FEVER_PURCHASE_DETAILS_MARKER.
 *   2. MAP-LINK MARKER (required, strictly after #1) --
 *      FEVER_MAP_LINK_MARKER.
 *   3. EVENT NAME -- the first non-empty line between #1 and #2.
 *   4. LOCATION -- the LAST non-empty line between the end of the event-name
 *      line and #2. Bounding on both sides is what makes this correct
 *      whether or not the venue shares a line with either neighbour.
 *   5. DATE/TIME -- the first FEVER_DATE_TIME_PATTERN match strictly after
 *      #1. The matched substring, whitespace-collapsed, is kept verbatim for
 *      `description` -- the same "reproduce the real line back to the owner
 *      rather than rebuild it from the parsed digits" choice the
 *      Ticketmaster CZ and Entradio parsers make.
 *   6. TICKET IDENTIFIER (optional) -- FEVER_TICKET_ID_PATTERN.
 *   7. TICKET QUANTITY (optional) -- FEVER_TICKET_QUANTITY_PATTERN.
 *   8. PER-SEAT CODES (optional) -- every FEVER_TICKET_CODE_PATTERN match in
 *      the region from the ticket-ID match's end to FEVER_ORDER_SUMMARY_MARKER
 *      (or the end of text when that marker is absent).
 *
 * The date format is Czech-ABBREVIATED and genuinely different from all
 * four existing parsers' formats (enigoo.cz's DD.MM.YYYY, Kino Art's
 * D. M. YYYY, Ticketmaster CZ's full-English-name format, Entradio's
 * D. M. YYYY) -- it gets its OWN regex and its OWN table
 * (FEVER_MONTH_ABBREVIATIONS); TICKETMASTER_CZ_MONTH_NAMES holds full
 * ENGLISH names and is not reusable here (D-04).
 *
 * DAY-PERIOD RULES (D-05): "odp." (afternoon) adds 12 to a 1-11 hour and
 * leaves 12 as 12; "dop." (morning) leaves a 1-11 hour and turns 12 into 0.
 * Only the "odp." branch is directly observed in the real sample; "dop."
 * and the noon/midnight edges follow the standard 12-hour convention and are
 * documented as such. A missing day-period marker is read as 24-hour, never
 * treated as an error.
 *
 * THE BODY CARRIES NO YEAR (D-06/D-07): the only four-digit year anywhere in
 * the real message is the footer copyright line, which is never read as the
 * event year. The year comes entirely from the injected `referenceDate`
 * (the message's received date) via feverResolveEventYear. This parser
 * deliberately never reads a live clock -- a missing or non-Date
 * `referenceDate` is a controlled throw, not a silent fallback to the
 * current clock, since a silently-wrong calendar year is a worse failure
 * than a diagnostic email.
 *
 * `ticketIdentifier` makes the DEDUP SAFETY NET ACTIVE for this portal -- a
 * real difference from the Ticketmaster CZ parser's documented null
 * identifier. `ticketQuantity` is NEVER an event multiplier, per this
 * file's class-level ONE-EVENT-PER-PURCHASE design (D-10).
 *
 * D-11 NEGATIVE CONTRACT: the per-seat codes go into `description` ONLY; no
 * `ticketCodes` field is returned, no attachment fetcher is registered for
 * this portal, no outbound HTTP call is made, and no new OAuth scope is
 * needed.
 *
 * The real email carries one PDF attachment (`order_*.pdf`), which the
 * registered findFeverTicketPdfAttachment can attach when the owner turns
 * `insertPdfIntoEvent` on (D-12).
 *
 * No speculative handling of an unobserved rendering variant is applied
 * here: any such variant is handled THEN, with real data, per this file's
 * established discipline -- the full-raw-text throws below are what make
 * that possible. ROUND 2 (quick-260921-gj0, live-test-driven): the owner's
 * first live run revealed that Gmail renders an <img> in the plain-text view
 * as a bracketed "[image: <alt text>]" line, which landed between the
 * purchase-details marker and the real event-name line and became the
 * created event's summary. feverFirstNonEmptyLine/feverLastNonEmptyLine now
 * skip such a line (D-18/D-19) -- the one thing derived from that
 * observation, applied narrowly (D-20).
 *
 * `subject` (quick-260921-gj0 round 2, D-25) is an OPTIONAL third parameter,
 * NEVER validated and NEVER throws on its own. EVENT NAME resolution order
 * (D-24):
 *   1. SUBJECT PATH (primary) -- feverEventNameFromSubject(subject): the
 *      real Subject header decodes to FEVER_SUBJECT_PREFIX followed by the
 *      event name VERBATIM, including its own internal colon. Matched as a
 *      PREFIX ONLY, never split on a colon. Evidenced by exactly ONE real
 *      sample -- the same footing as FEVER_MONTH_ABBREVIATIONS' eleven
 *      un-observed entries.
 *   2. BODY PATH (fallback) -- used when the subject is missing, not a
 *      string, does not start with the prefix, or leaves nothing after it.
 *      This is the pre-existing feverFirstNonEmptyLine read, now hardened by
 *      the D-19 placeholder skip.
 *   3. Neither yields anything -> the EXISTING controlled event-name throw,
 *      wording unchanged, carrying the full raw body text.
 * LOCATION stays body-sourced ALWAYS (D-26): there is no venue signal in the
 * subject, so the body's own first surviving line is what anchors the
 * location region, regardless of which source won the event name. The
 * content detector (feverTextHasPurchaseDetails) stays body-based too --
 * D-08's marketing-mail admission gate is unaffected by this round.
 *
 * Returns `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier, ticketQuantity, description }` (month zero-indexed;
 * deliberately NO `ticketCodes` field, D-11). Throws a controlled Error,
 * always ending with the full raw `text` untruncated, if `referenceDate` is
 * missing/invalid, if either required marker is absent, if the event name
 * or location cannot be extracted, if no date/time pattern is found, if the
 * month abbreviation is unrecognized, or if the matched hour/minute are out
 * of range. Pure, no GAS globals, no clock read.
 */
function parseFeverTicketText(text, referenceDate, subject) {
  const rawText = String(text || '');
  const normalizedText = feverNormalizeBodyText(rawText);

  if (!(referenceDate instanceof Date) || isNaN(referenceDate.getTime())) {
    throw new Error(
      'Fever ticket text carries no year of its own: a valid reference Date (the message\'s received date) is required. Full extracted text:\n' +
        rawText
    );
  }

  const purchaseDetailsIndex = normalizedText.indexOf(FEVER_PURCHASE_DETAILS_MARKER);
  if (purchaseDetailsIndex === -1) {
    throw new Error(
      'Unrecognized Fever ticket text: "Tady jsou podrobnosti o tvém nákupu" marker not found. Full extracted text:\n' + rawText
    );
  }
  const afterPurchaseDetails = purchaseDetailsIndex + FEVER_PURCHASE_DETAILS_MARKER.length;

  const mapLinkIndex = normalizedText.indexOf(FEVER_MAP_LINK_MARKER, afterPurchaseDetails);
  if (mapLinkIndex === -1 || mapLinkIndex <= purchaseDetailsIndex) {
    throw new Error('Unrecognized Fever ticket text: "Zobrazit na mapě" marker not found. Full extracted text:\n' + rawText);
  }

  const eventNameRegion = normalizedText.slice(afterPurchaseDetails, mapLinkIndex);
  // quick-260921-gj0 round 2, D-24: the SUBJECT is the PRIMARY event-name
  // source; the body's own first surviving line (hardened against Gmail
  // image-placeholder lines by feverFirstNonEmptyLine, D-19) is computed
  // unconditionally and used as the FALLBACK -- and, per D-26, ALSO as the
  // sole anchor for the LOCATION region below, regardless of which source
  // won the event name.
  const bodyEventNameLine = feverFirstNonEmptyLine(eventNameRegion);
  const subjectEventName = feverEventNameFromSubject(subject);
  const eventName = subjectEventName || bodyEventNameLine;
  if (!eventName) {
    throw new Error(
      'Unrecognized Fever ticket text: could not extract the event name between the purchase-details marker and the map-link marker. Full extracted text:\n' +
        rawText
    );
  }

  // LOCATION anchors on the BODY's own first surviving line ALWAYS, never on
  // the resolved `eventName` (D-26): a SUBJECT-sourced eventName need not
  // appear in the body at all, in which case
  // normalizedText.indexOf(eventName, ...) would return -1 and silently
  // widen this region to start one character before the purchase-details
  // marker, corrupting the location read.
  const bodyEventNameAbsoluteIndex = normalizedText.indexOf(bodyEventNameLine, afterPurchaseDetails);
  const bodyEventNameLineEndIndex = bodyEventNameAbsoluteIndex + bodyEventNameLine.length;
  const locationRegion = normalizedText.slice(bodyEventNameLineEndIndex, mapLinkIndex);
  const location = feverLastNonEmptyLine(locationRegion);
  if (!location) {
    throw new Error(
      'Unrecognized Fever ticket text: could not extract the location between the event name and the map-link marker. Full extracted text:\n' +
        rawText
    );
  }

  const searchRegionForDateTime = normalizedText.slice(afterPurchaseDetails);
  const dateTimeMatch = FEVER_DATE_TIME_PATTERN.exec(searchRegionForDateTime);
  if (!dateTimeMatch) {
    throw new Error('Unrecognized Fever ticket text: no date/time pattern found. Full extracted text:\n' + rawText);
  }

  const day = Number(dateTimeMatch[1]);
  const monthToken = dateTimeMatch[2];
  const foldedMonthToken = feverFoldMonthToken(monthToken);
  if (!Object.prototype.hasOwnProperty.call(FEVER_MONTH_ABBREVIATIONS, foldedMonthToken)) {
    throw new Error(
      'Unrecognized Fever ticket text: unrecognized month abbreviation "' + monthToken + '". Full extracted text:\n' + rawText
    );
  }
  const month = FEVER_MONTH_ABBREVIATIONS[foldedMonthToken];

  let hour = Number(dateTimeMatch[3]);
  const minute = Number(dateTimeMatch[4]);
  const meridiem = dateTimeMatch[5] ? dateTimeMatch[5].toLowerCase() : null;

  if (meridiem) {
    if (hour < 1 || hour > 12) {
      throw new Error('Hour out of range (1-12) in Fever ticket date/time match. Full extracted text:\n' + rawText);
    }
    if (meridiem === 'odp') {
      hour = hour === 12 ? 12 : hour + 12;
    } else {
      hour = hour === 12 ? 0 : hour;
    }
  } else if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (0-23) in Fever ticket date/time match. Full extracted text:\n' + rawText);
  }

  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (0-59) in Fever ticket date/time match. Full extracted text:\n' + rawText);
  }

  const year = feverResolveEventYear(month, day, referenceDate);
  const dateTimeText = dateTimeMatch[0].replace(/\s+/g, ' ').trim();

  const ticketIdentifierMatch = FEVER_TICKET_ID_PATTERN.exec(normalizedText);
  const ticketIdentifier = ticketIdentifierMatch ? ticketIdentifierMatch[1] : null;

  const quantityMatch = FEVER_TICKET_QUANTITY_PATTERN.exec(normalizedText);
  const ticketQuantity = quantityMatch ? Number(quantityMatch[1]) : null;

  let codesRegion = '';
  if (ticketIdentifierMatch) {
    const codesStart = ticketIdentifierMatch.index + ticketIdentifierMatch[0].length;
    const orderSummaryIndex = normalizedText.indexOf(FEVER_ORDER_SUMMARY_MARKER, codesStart);
    codesRegion = normalizedText.slice(codesStart, orderSummaryIndex === -1 ? normalizedText.length : orderSummaryIndex);
  }

  FEVER_TICKET_CODE_PATTERN.lastIndex = 0;
  const ticketCodes = [];
  let codeMatch;
  while ((codeMatch = FEVER_TICKET_CODE_PATTERN.exec(codesRegion)) !== null) {
    ticketCodes.push(codeMatch[0]);
  }

  const descriptionParagraphs = [eventName, location, dateTimeText];
  if (ticketIdentifier) {
    descriptionParagraphs.push('ID vstupenky: ' + ticketIdentifier);
  }
  if (quantityMatch) {
    descriptionParagraphs.push(quantityMatch[0].replace(/\s+/g, ' ').trim());
  }
  if (ticketCodes.length > 0) {
    descriptionParagraphs.push(ticketCodes.join('\n'));
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
    description: descriptionParagraphs.join('\n\n'),
  };
}

/**
 * feverTextHasPurchaseDetails — the body-content admission predicate for
 * this portal (D-08): hello@feverup.com also sends ordinary marketing mail,
 * so admitting on sender alone would repeat the exact bug already fixed
 * once for Ticketmaster CZ (debug/ticketmaster-cz-order-confirm) -- routine
 * promotional mail gets claimed, fails to parse, and emails the owner a
 * spurious failure notification. Returns a LITERAL boolean for whether the
 * normalized text contains FEVER_PURCHASE_DETAILS_MARKER. Never throws,
 * tolerates a null/undefined/empty body. Pure, no GAS globals.
 */
function feverTextHasPurchaseDetails(text) {
  return feverNormalizeBodyText(text).indexOf(FEVER_PURCHASE_DETAILS_MARKER) !== -1;
}

// --- TICKETPORTAL.CZ (ICS-SOURCED, quick-261005-orv) ------------------------
//
// The SIXTH portal, and the first whose event data is NOT extracted by a
// portal-specific text/body parser at all: Ticketportal.cz's order
// confirmation carries the event's own `.ics` attachment, parsed by the
// EXISTING parseIcs/isIcsAttachment exports of src/05-action-ics-import.js
// (D-01) -- never a second ICS parser. The pure helpers below only
// post-process what parseIcs already returned.

// TICKETPORTAL_CZ_SUMMARY_DATE_SUFFIX_PATTERN — Ticketportal.cz bakes a
// "D(D).M(M).YYYY H(H):MM" date/time tail onto the end of the .ics SUMMARY
// (e.g. "HELENA Forever  19.10.2026 19:00"), separated from the real event
// name by 2+ spaces. That date/time is already available, correctly, from
// DTSTART -- this pattern strips the redundant, duplicated tail so the
// calendar event's summary is just the event name.
const TICKETPORTAL_CZ_SUMMARY_DATE_SUFFIX_PATTERN = /\s{2,}\d{1,2}\.\s*\d{1,2}\.\s*\d{4}\s+\d{1,2}:\d{2}\s*$/;

/**
 * stripTicketportalCzSummaryDateSuffix — removes the trailing baked-in
 * date/time tail (TICKETPORTAL_CZ_SUMMARY_DATE_SUFFIX_PATTERN) from a
 * Ticketportal.cz .ics SUMMARY, then trims the result. A summary with no
 * such tail is returned trimmed and otherwise unchanged. null/undefined/''
 * all return ''. Never throws. Pure, no GAS globals.
 */
function stripTicketportalCzSummaryDateSuffix(summary) {
  return String(summary || '')
    .replace(TICKETPORTAL_CZ_SUMMARY_DATE_SUFFIX_PATTERN, '')
    .trim();
}

// TICKETPORTAL_CZ_ORDER_NUMBER_FROM_PDF_NAME_PATTERN — the digit run
// immediately before the `.pdf` extension in the eTicket filename (e.g.
// "eTicket_13634244.pdf" -> "13634244").
const TICKETPORTAL_CZ_ORDER_NUMBER_FROM_PDF_NAME_PATTERN = /_(\d+)\.pdf$/i;

/**
 * extractTicketportalCzOrderNumber — reads the per-ORDER confirmation
 * number from the eTicket PDF's own FILENAME (D-04d). Returns the digit
 * string, or `null` when `pdfName` is null/undefined/empty or does not
 * match. Never throws. Pure, no GAS globals.
 *
 * WHY THE FILENAME, NEVER THE .ics UID: the .ics UID
 * (`:ticketportal.performance.120009907`) names the PERFORMANCE, not the
 * ORDER. Two genuine, separate orders for the same show share that UID, so
 * keying isDuplicateTicketPurchase on it would silently skip the second
 * order's calendar event -- a false "already exists" that drops a real
 * purchase. The order number baked into the eTicket filename is per-ORDER
 * (it equals the body's own "Variabilní symbol"), which is what this
 * portal's dedup safety net must key on instead.
 */
function extractTicketportalCzOrderNumber(pdfName) {
  const match = TICKETPORTAL_CZ_ORDER_NUMBER_FROM_PDF_NAME_PATTERN.exec(String(pdfName || ''));
  return match ? match[1] : null;
}

// =============================================================================
// SMSTICKET.CZ (info@smsticket.cz) — the SEVENTH supported portal, the FIFTH
// body-sourced one (quick-261026-ma6)
// =============================================================================
//
// STRUCTURAL FACTS (from the real sample .eml, decoded by the planner):
//   - The confirmation's text/plain MIME part is PRESENT but EMPTY (0 bytes).
//     message.getPlainBody() hands this parser NOTHING, which is exactly
//     what forces the generic D-02 fallback (ticketingMessageBodyText /
//     ticketingHtmlToPlainText, defined above ticketBodyLooksProcessable
//     below) — this portal is the reason that fallback exists at all.
//   - The template is in English (the owner's account language); only the
//     venue/seat data is Czech.
//   - The body carries its own four-digit YEAR, unlike Fever — no reference-
//     date inference is needed. There is no end time, so
//     DEFAULT_EVENT_DURATION_MINUTES (120) applies.
//   - Per-seat block layout: a seat label, then its price ("590 Kč"), then
//     its entry code ("NNN-NNN-N"), then an "Add to Apple Wallet" link, then
//     a QR <img> whose alt repeats the entry code.
//   - Entry codes are BEARER CREDENTIALS: whoever first shows a valid code
//     gets in. One purchase can cover multiple seats; ONE PDF holds every
//     ticket of the purchase.
//
// This parser must work on BOTH ticketingHtmlToPlainText's own rendering AND
// Gmail's own getPlainBody() rendering, should Gmail ever return one for this
// sender (D-03) — hence every read below is structural or anchored on
// "surviving lines" (smsticketSurvivingLines), never on a fixed line
// position.

// SMSTICKET_SUBJECT_PREFIX (D-04) — the fixed template text smsticket.cz's
// confirmation-email Subject header starts with, decoded from the real
// sample's own Subject header. Matched as a PREFIX ONLY (indexOf === 0) and
// NEVER split on a colon, the same discipline FEVER_SUBJECT_PREFIX documents
// for itself: the event name can carry its own internal colon.
const SMSTICKET_SUBJECT_PREFIX = 'smsticket: Tickets purchase ';

// SMSTICKET_DATE_TIME_PATTERN (D-05) — day digits, a dot, month digits, a
// dot, a four-digit year, each dot TOLERANT of surrounding whitespace;
// whitespace; ONE OPTIONAL non-digit word token followed by whitespace (the
// observed "at"; a Czech-template "v" passes the same shape, confirmed
// against the marketing fixture); then H(H):MM. Anchoring on the day digits
// is what skips a leading weekday word ("Sunday") without needing to match
// it. Non-global: a single `.exec()` always returns the FIRST match from the
// start of the string, which is what "use the FIRST match" (D-05) means in
// practice — no `g` flag, no `lastIndex` state to manage.
const SMSTICKET_DATE_TIME_PATTERN = /(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{4})\s+(?:[^\s\d]+\s+)?(\d{1,2}):(\d{2})/;

// SMSTICKET_TICKET_CODE_PATTERN (D-07) — a word-bounded "NNN-NNN-N" entry
// code. `g`-flagged and used ONLY via extractSmsticketTicketCodes below,
// which resets `lastIndex` to 0 before every scan — the same discipline
// FEVER_TICKET_CODE_PATTERN documents for itself, so no state leaks between
// calls.
const SMSTICKET_TICKET_CODE_PATTERN = /\b\d{3}-\d{3}-\d\b/g;

// SMSTICKET_TICKET_CODE_LINE_PATTERN (D-08) — a surviving line that is
// NOTHING BUT one entry code, anchored both ends. Used to recognize the
// per-seat code's OWN line so it can be enriched with its seat/price
// neighbours; a code sharing a line with other text (e.g. its price) simply
// never matches this, which is the documented graceful-degradation path
// (the bare code is still used, just without enrichment).
const SMSTICKET_TICKET_CODE_LINE_PATTERN = /^\d{3}-\d{3}-\d$/;

// SMSTICKET_PRICE_LINE_PATTERN (D-08) — a surviving line that is NOTHING BUT
// a Kč price: digits, optional inner spaces (a thousands separator),
// optional decimal part, optional whitespace, then the literal "Kč",
// anchored both ends.
const SMSTICKET_PRICE_LINE_PATTERN = /^\d[\d ]*(?:[.,]\d+)?\s*Kč$/;

// SMSTICKET_BOLD_WRAPPED_LINE_PATTERN (D-03) — ONE wrapping pair of asterisks
// around a non-asterisk run, Entradio's own plain-text bold convention
// (ENTRADIO_BOLD_VALUE_PATTERN), reused here because Gmail's own rendering of
// this portal's HTML is assumed to use the identical convention for the
// `<strong>`/`<b>` seat and code labels (D-03's "two renderings" contract).
const SMSTICKET_BOLD_WRAPPED_LINE_PATTERN = /^\*([^*]+)\*$/;

/**
 * smsticketSurvivingLines — pure; splits `text` on CR, LF or CRLF AFTER
 * normalizing it with feverNormalizeBodyText (folding U+00A0 and stripping
 * invisible padding characters — reused by name per this file's Fever
 * section, never duplicated). Each line is trimmed, has ONE wrapping pair of
 * SMSTICKET_BOLD_WRAPPED_LINE_PATTERN asterisks removed and is re-trimmed,
 * then dropped if it is empty or is ONLY a Gmail image-placeholder (via
 * feverLineIsOnlyImagePlaceholders, also reused by name). Returns the
 * surviving lines, in order. Never throws.
 */
function smsticketSurvivingLines(text) {
  const normalized = feverNormalizeBodyText(text);
  const lines = normalized.split(/\r\n|\r|\n/);
  const survivors = [];

  for (let i = 0; i < lines.length; i++) {
    let trimmed = lines[i].trim();
    const boldMatch = SMSTICKET_BOLD_WRAPPED_LINE_PATTERN.exec(trimmed);
    if (boldMatch) {
      trimmed = boldMatch[1].trim();
    }
    if (trimmed && !feverLineIsOnlyImagePlaceholders(trimmed)) {
      survivors.push(trimmed);
    }
  }

  return survivors;
}

/**
 * smsticketEventNameFromSubject (D-04) — pure; NEVER THROWS. Returns the
 * trimmed remainder of `subject` after SMSTICKET_SUBJECT_PREFIX when
 * `subject` is a string that starts with that exact prefix (matched as a
 * PREFIX via indexOf === 0 — never split on a colon) AND something
 * non-whitespace remains after stripping it. Returns `null` for a subject
 * that is missing, not a string, does not start with the prefix, or leaves
 * nothing but whitespace — each of those is the documented FALLBACK
 * condition to the body-sourced read in parseSmsticketTicketText below, not
 * an error. Same contract as feverEventNameFromSubject.
 */
function smsticketEventNameFromSubject(subject) {
  if (typeof subject !== 'string' || subject.indexOf(SMSTICKET_SUBJECT_PREFIX) !== 0) {
    return null;
  }
  const remainder = subject.slice(SMSTICKET_SUBJECT_PREFIX.length).trim();
  return remainder || null;
}

/**
 * extractSmsticketTicketCodes (D-07) — pure; resets
 * SMSTICKET_TICKET_CODE_PATTERN's `lastIndex` to 0, then collects every
 * DISTINCT match in `region` in FIRST-SEEN order (Gmail can repeat a code as
 * a QR placeholder's own alt text, which is why distinctness matters rather
 * than raw match count). Never throws; `null`, `undefined` or `''` all
 * return `[]`.
 */
function extractSmsticketTicketCodes(region) {
  if (!region) {
    return [];
  }

  SMSTICKET_TICKET_CODE_PATTERN.lastIndex = 0;
  const seen = [];
  const seenLookup = {};
  let match = SMSTICKET_TICKET_CODE_PATTERN.exec(region);
  while (match !== null) {
    const code = match[0];
    if (!seenLookup[code]) {
      seenLookup[code] = true;
      seen.push(code);
    }
    match = SMSTICKET_TICKET_CODE_PATTERN.exec(region);
  }

  return seen;
}

/**
 * parseSmsticketTicketText (D-03 through D-08) — the info@smsticket.cz
 * ticket-BODY parser. `text` is the rendered body text (either
 * ticketingHtmlToPlainText's own rendering of the HTML body, since the real
 * text/plain part is empty, or Gmail's own getPlainBody() rendering, should
 * Gmail ever return one for this sender — D-03). `referenceDate` is accepted
 * for the BODY-PARSER CONTRACT (processTicketFromMessageBody always passes
 * it) but IGNORED: the year comes from the body itself, unlike Fever.
 * `subject` is the message's Subject header, used as the PRIMARY event-name
 * source (D-04).
 *
 * EXTRACTION SEQUENCE:
 *   1. Normalize `text` (feverNormalizeBodyText) and find the FIRST
 *      SMSTICKET_DATE_TIME_PATTERN match — day/month/year/hour/minute, each
 *      range-validated (month 1-12, day 1-31, hour 0-23, minute 0-59).
 *   2. `dateTimeText` is the matched substring with internal whitespace
 *      collapsed to single spaces.
 *   3. `location` (D-06) is the FIRST surviving line STRICTLY AFTER the
 *      match's end index — correct whether or not Gmail ever joins the
 *      venue onto the date line, since smsticketSurvivingLines still finds
 *      it as the first surviving line of whatever remains.
 *   4. `eventName` (D-04): the subject-derived name
 *      (smsticketEventNameFromSubject) when available, otherwise the LAST
 *      surviving line of the text BEFORE THE START OF THE LINE containing
 *      the date/time match — cutting at the LINE START rather than the
 *      match index is what keeps a leading weekday token ("Sunday") from
 *      ever being read as the event name.
 *   5. Entry codes (D-07, extractSmsticketTicketCodes) are scanned ONLY in
 *      the region strictly after the date/time match. `ticketIdentifier` is
 *      the FIRST code, or `null` when there are none — this is the DEDUP
 *      SAFETY NET's key for this portal (an entry code is unique per
 *      ticket, so the first ticket's code identifies the whole purchase,
 *      stably across re-processing of the same email).
 *   6. Per-seat description lines (D-08): over the surviving lines after the
 *      date/time match, each DISTINCT code (in D-07's order) is enriched
 *      to `<seat> | <price> | <code>` when it occupies its OWN surviving
 *      line (SMSTICKET_TICKET_CODE_LINE_PATTERN) AND the line immediately
 *      before that is a bare price (SMSTICKET_PRICE_LINE_PATTERN) AND a
 *      line two before it exists (the seat label) — otherwise the BARE code
 *      is used (graceful degradation, e.g. a code sharing its line with its
 *      price). `description` is `[eventName, location, dateTimeText,
 *      ticketLines joined by '\n']` joined by `'\n\n'`, OMITTING the 4th
 *      paragraph entirely when there are no codes at all.
 *
 * Return shape: exactly `{ eventName, location, year, month, day, hour,
 * minute, ticketIdentifier, description }` — no `ticketCodes` field; the
 * seat count never multiplies the event (one event per purchase, D-08).
 *
 * Every thrown message starts with "Unrecognized smsticket.cz ticket text:"
 * (no date/time, no event name or no location) or "<Field> out of range
 * (<value>) in smsticket.cz ticket date/time match." (an invalid month, day,
 * hour or minute), and ALWAYS ends with ". Full extracted text:\n" followed
 * by the full ORIGINAL raw text — this file's DIAGNOSTIC-ON-FAILURE
 * convention, so the owner's failure email always carries enough context to
 * fix the next real-world variant.
 */
function parseSmsticketTicketText(text, referenceDate, subject) {
  const rawText = String(text || '');
  const normalizedText = feverNormalizeBodyText(rawText);

  const dateMatch = SMSTICKET_DATE_TIME_PATTERN.exec(normalizedText);
  if (!dateMatch) {
    throw new Error('Unrecognized smsticket.cz ticket text: no date/time found. Full extracted text:\n' + rawText);
  }

  const day = parseInt(dateMatch[1], 10);
  const month = parseInt(dateMatch[2], 10);
  const year = parseInt(dateMatch[3], 10);
  const hour = parseInt(dateMatch[4], 10);
  const minute = parseInt(dateMatch[5], 10);

  if (month < 1 || month > 12) {
    throw new Error('Month out of range (' + month + ') in smsticket.cz ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (day < 1 || day > 31) {
    throw new Error('Day out of range (' + day + ') in smsticket.cz ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (' + hour + ') in smsticket.cz ticket date/time match. Full extracted text:\n' + rawText);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (' + minute + ') in smsticket.cz ticket date/time match. Full extracted text:\n' + rawText);
  }

  const dateTimeText = dateMatch[0].replace(/\s+/g, ' ').trim();
  const matchEndIndex = dateMatch.index + dateMatch[0].length;
  const afterDateText = normalizedText.slice(matchEndIndex);
  const linesAfterDate = smsticketSurvivingLines(afterDateText);

  if (linesAfterDate.length === 0) {
    throw new Error('Unrecognized smsticket.cz ticket text: no location found. Full extracted text:\n' + rawText);
  }
  const location = linesAfterDate[0];

  let eventName = smsticketEventNameFromSubject(subject);
  if (!eventName) {
    const lastNewlineBeforeMatch = normalizedText.lastIndexOf('\n', Math.max(dateMatch.index - 1, 0));
    const startOfDateLineIndex = dateMatch.index === 0 ? 0 : lastNewlineBeforeMatch === -1 ? 0 : lastNewlineBeforeMatch + 1;
    const linesBeforeDate = smsticketSurvivingLines(normalizedText.slice(0, startOfDateLineIndex));

    if (linesBeforeDate.length === 0) {
      throw new Error('Unrecognized smsticket.cz ticket text: no event name found. Full extracted text:\n' + rawText);
    }
    eventName = linesBeforeDate[linesBeforeDate.length - 1];
  }

  const codes = extractSmsticketTicketCodes(afterDateText);
  const ticketIdentifier = codes.length > 0 ? codes[0] : null;

  const ticketLines = codes.map(function (code) {
    const codeLineIndex = linesAfterDate.indexOf(code);
    if (
      codeLineIndex >= 2 &&
      SMSTICKET_TICKET_CODE_LINE_PATTERN.test(linesAfterDate[codeLineIndex]) &&
      SMSTICKET_PRICE_LINE_PATTERN.test(linesAfterDate[codeLineIndex - 1])
    ) {
      return linesAfterDate[codeLineIndex - 2] + ' | ' + linesAfterDate[codeLineIndex - 1] + ' | ' + code;
    }
    return code;
  });

  const descriptionParagraphs = [eventName, location, dateTimeText];
  if (ticketLines.length > 0) {
    descriptionParagraphs.push(ticketLines.join('\n'));
  }

  return {
    eventName: eventName,
    location: location,
    year: year,
    month: month - 1,
    day: day,
    hour: hour,
    minute: minute,
    ticketIdentifier: ticketIdentifier,
    description: descriptionParagraphs.join('\n\n'),
  };
}

/**
 * smsticketTextHasPurchaseDetails (D-09) — the REQUIRED content detector for
 * info@smsticket.cz: `info@` is the sender's general address, and the real
 * confirmation's own footer advertises its newsletter and weekly
 * promotions, so sender alone cannot distinguish a purchase confirmation
 * from ordinary marketing mail (the same risk debug/ticketmaster-cz-order-
 * confirm first caught). Returns a LITERAL boolean: `true` only when the
 * normalized `text` has a SMSTICKET_DATE_TIME_PATTERN match AND
 * extractSmsticketTicketCodes finds at least one code strictly AFTER that
 * match. Structural and language-neutral (never an English phrase marker,
 * unlike ticketmasterCzTextHasOrderDetails), so it survives the owner's
 * English-template account and a hypothetical Czech-template one alike.
 * Never throws on `null`, `undefined` or `''`.
 */
function smsticketTextHasPurchaseDetails(text) {
  if (!text) {
    return false;
  }

  const normalizedText = feverNormalizeBodyText(text);
  const dateMatch = SMSTICKET_DATE_TIME_PATTERN.exec(normalizedText);
  if (!dateMatch) {
    return false;
  }

  const afterDateText = normalizedText.slice(dateMatch.index + dateMatch[0].length);
  return extractSmsticketTicketCodes(afterDateText).length > 0;
}

// =============================================================================
// WEBTICKET.CZ (automat@webticket.cz) — the EIGHTH supported portal, the
// SECOND ICS-sourced one (quick-261028-h1n)
// =============================================================================
//
// STRUCTURAL FACTS (from the real sample .eml, decoded by the planner):
//   - The confirmation's text/plain MIME part is NON-EMPTY but renders the
//     event name/date line in ALL CAPS ("PO 28. 12. 2026 19:00 - BIOGRAF
//     LÁSKA 2: NOVÝ PŘÍBĚH") — correct Czech casing cannot be recovered
//     from capitals, which is why this portal is ICS-sourced rather than
//     body-sourced (D-01), despite the body being non-empty.
//   - The confirmation ALSO carries its own genuine `.ics` attachment,
//     whose SUMMARY and LOCATION keep the proper case, and whose DTSTART is
//     an exact UTC instant. That `.ics`'s DTEND equals its DTSTART (a
//     zero-length VEVENT) — this is also why, before this change, the
//     generic ICS_CALENDAR_ACTION was already importing this sender's
//     `.ics` as a zero-length event (importOnlyFrom defaults to every
//     sender); registering this portal structurally ends that (D-01, D-07).
//   - The `.ics` UID is a PHP `uniqid()` minted PER SENT EMAIL (its 8-hex
//     seconds prefix equals the message's own DKIM `t=` timestamp) — a
//     re-sent confirmation would carry a fresh UID, so it is never used for
//     dedup (D-05).
//   - ONE plain-ASCII-named ticket PDF (`webticket_<order>.pdf`) carries the
//     order number in its FILENAME — the primary dedup source, same
//     precedent as Ticketportal.cz's eTicket PDF (D-05, D-06).
//   - The body's "Zobraz zakázku" order-status link and its two Wallet
//     links share one URL that is a BEARER LINK (opens the order page,
//     which hands out wallet passes) — never followed, never fetched
//     (D-01, D-09, D-11).
//   - The body also carries a weekday-prefixed date format ("PO 28. 12.
//     2026 19:00") — observed, but deliberately NEVER parsed: the `.ics`
//     DTSTART is already the exact instant (D-10).
//   - WebTicket sells tickets ON BEHALF OF MANY VENUES (this sample:
//     Divadlo Kalich) — the ONE sender address `automat@webticket.cz`
//     covers every venue, the same platform-sender shape as Entradio.

// WEBTICKET_ORDER_NUMBER_FROM_PDF_NAME_PATTERN — the digit run between the
// literal `webticket_` prefix and the `.pdf` extension in the real ticket
// filename (e.g. "webticket_3064785.pdf" -> "3064785"). Case-insensitive,
// since nothing about the filename's casing is guaranteed to survive every
// mail client. OBSERVED (the real sample) VERSUS TOLERANCE: only this one
// filename shape has been seen live; a differently-shaped WebTicket filename
// would need handling then, with real data — see
// extractWebticketOrderNumberFromPdfName's own SCOPE LIMITATION below.
const WEBTICKET_ORDER_NUMBER_FROM_PDF_NAME_PATTERN = /^webticket_(\d+)\.pdf$/i;

/**
 * extractWebticketOrderNumberFromPdfName (D-05) — reads the per-ORDER
 * WebTicket number from the ticket PDF's own FILENAME. Returns the digit
 * string, or `null` when `pdfName` is null/undefined/empty or does not
 * match. Never throws. Pure, no GAS globals.
 *
 * WHY THE FILENAME, NEVER THE .ics UID, NEVER kal: — the .ics UID is a PHP
 * `uniqid()` minted PER SENT EMAIL (see this section's banner): a re-sent
 * confirmation would carry a brand-new UID, so keying dedup on it would let
 * a re-send create a second event for the SAME purchase. The body's
 * "kal:00795453" number is the VENUE's own booking number (its prefix names
 * the venue, e.g. "kal" for Divadlo Kalich) — it is venue-scoped, not
 * order-scoped, so two different orders at different venues could
 * legitimately collide on their own "kal:"-equivalent numbers. The order
 * number baked into the ticket PDF's filename (same digits as the Subject's
 * "#3064785" and the order-status URL's path segment) is per-ORDER and
 * stable across re-processing and re-sends, which is what this portal's
 * dedup safety net must key on instead.
 */
function extractWebticketOrderNumberFromPdfName(pdfName) {
  const match = WEBTICKET_ORDER_NUMBER_FROM_PDF_NAME_PATTERN.exec(String(pdfName || ''));
  return match ? match[1] : null;
}

// WEBTICKET_ORDER_NUMBER_PATTERN (D-08, quick-261028-h1n Task 2) — the
// literal "ZAKÁZKA #<digits>" marker (case-insensitive, so it matches the
// real plain part's "ZAKÁZKA #" and the HTML rendering's "zakázka #"/
// "Zakázka #" alike), accepting either the diacritic Á or the plain ASCII A
// (observed shapes only; a differently-rendered variant would need handling
// then, with real data). Anchored on the literal `#` is what keeps this from
// ever matching the order-status URL's "zakazka/3064785" path segment,
// which has `/` where this pattern requires `#`.
const WEBTICKET_ORDER_NUMBER_PATTERN = /ZAK[AÁ]ZKA\s*#\s*(\d+)/i;

// WEBTICKET_ORDER_REFERENCE_PATTERN (D-08) — the venue's own booking
// reference immediately after the order number, separated by " / "
// (e.g. "ZAKÁZKA #3064785 / KAL:00795453" -> "kal:00795453", lowercased for
// a stable, predictable description). Observed shape only: one
// alphabetic-prefix-then-colon-then-digits token.
const WEBTICKET_ORDER_REFERENCE_PATTERN = /ZAK[AÁ]ZKA\s*#\s*\d+\s*\/\s*([A-Za-z]+:\d+)/i;

// WEBTICKET_WRAPPED_LINE_PATTERN (D-08) — ONE wrapping pair of `*` or `_`
// around a non-`*`/`_` run, the real plain part's own markdown-ish
// emphasis convention around the seat-count and the "V případě
// nesrovnalostí..." disclaimer lines (e.g. "_1 ks, 849 Kč_ ").
const WEBTICKET_WRAPPED_LINE_PATTERN = /^([*_])([^*_]+)\1$/;

/**
 * webticketSurvivingLines (D-08) — pure; splits `text` on CR, LF or CRLF
 * AFTER normalizing it with feverNormalizeBodyText (folding U+00A0 and
 * stripping invisible padding characters — reused by name per this file's
 * Fever section, never duplicated). Each line is trimmed, has ONE wrapping
 * pair of WEBTICKET_WRAPPED_LINE_PATTERN `*`/`_` removed and is re-trimmed,
 * then dropped if it is empty. Returns the surviving lines, in order. Never
 * throws.
 */
function webticketSurvivingLines(text) {
  const normalized = feverNormalizeBodyText(text);
  const lines = normalized.split(/\r\n|\r|\n/);
  const survivors = [];

  for (let i = 0; i < lines.length; i++) {
    let trimmed = lines[i].trim();
    const wrappedMatch = WEBTICKET_WRAPPED_LINE_PATTERN.exec(trimmed);
    if (wrappedMatch) {
      trimmed = wrappedMatch[2].trim();
    }
    if (trimmed) {
      survivors.push(trimmed);
    }
  }

  return survivors;
}

/**
 * parseWebticketOrderText (D-08, quick-261028-h1n Task 2) — the
 * automat@webticket.cz ticket-BODY ENRICHMENT parser: unlike every
 * body-sourced portal's parser, this does NOT supply the core event fields
 * (those come from the `.ics`, see buildWebticketParsedTicket) — it only
 * extracts the description/dedup-fallback material the body carries. Works
 * identically on the real all-caps plain part, its CRLF-ending variant, and
 * ticketingHtmlToPlainText's own rendering of the HTML part — every anchor
 * below is a literal marker, never a line position. NEVER throws: `null`,
 * `undefined`, a non-string or `''` all return
 * `{ orderNumber: null, orderReference: null, seatLines: [], totalPriceLine: null }`.
 *
 * EXTRACTION SEQUENCE:
 *   1. Normalize `text` (feverNormalizeBodyText) and find the FIRST
 *      WEBTICKET_ORDER_NUMBER_PATTERN match -> `orderNumber`.
 *   2. `orderReference` is group 1 of the first
 *      WEBTICKET_ORDER_REFERENCE_PATTERN match, LOWERCASED, or `null` when
 *      there is none.
 *   3. `seatLines` is every webticketSurvivingLines(text) line containing
 *      the literal separator " ... " whose text after the LAST " ... "
 *      fully matches the reused SMSTICKET_PRICE_LINE_PATTERN (by name, same
 *      reuse discipline smsticket.cz already established for the Fever
 *      helpers) — kept verbatim, in order.
 *   4. `totalPriceLine` is the first surviving line that starts with
 *      "Cena:" and whose remainder, trimmed, fully matches
 *      SMSTICKET_PRICE_LINE_PATTERN — kept verbatim, or `null` when none
 *      qualifies.
 *
 * Return shape: exactly `{ orderNumber, orderReference, seatLines,
 * totalPriceLine }`.
 */
function parseWebticketOrderText(text) {
  if (typeof text !== 'string' || !text) {
    return { orderNumber: null, orderReference: null, seatLines: [], totalPriceLine: null };
  }

  const normalizedText = feverNormalizeBodyText(text);

  const orderNumberMatch = WEBTICKET_ORDER_NUMBER_PATTERN.exec(normalizedText);
  const orderNumber = orderNumberMatch ? orderNumberMatch[1] : null;

  const orderReferenceMatch = WEBTICKET_ORDER_REFERENCE_PATTERN.exec(normalizedText);
  const orderReference = orderReferenceMatch ? orderReferenceMatch[1].toLowerCase() : null;

  const survivingLines = webticketSurvivingLines(normalizedText);

  const seatLines = survivingLines.filter(function (line) {
    const lastSeparatorIndex = line.lastIndexOf(' ... ');
    if (lastSeparatorIndex === -1) {
      return false;
    }
    return SMSTICKET_PRICE_LINE_PATTERN.test(line.slice(lastSeparatorIndex + ' ... '.length).trim());
  });

  let totalPriceLine = null;
  for (let i = 0; i < survivingLines.length; i++) {
    const line = survivingLines[i];
    if (line.indexOf('Cena:') === 0 && SMSTICKET_PRICE_LINE_PATTERN.test(line.slice('Cena:'.length).trim())) {
      totalPriceLine = line;
      break;
    }
  }

  return {
    orderNumber: orderNumber,
    orderReference: orderReference,
    seatLines: seatLines,
    totalPriceLine: totalPriceLine,
  };
}

/**
 * buildWebticketParsedTicket (D-03, D-04, D-05; body enrichment added in
 * quick-261028-h1n Task 2, D-08) — builds the `parsedTicket` shape from a
 * single parseIcs-returned event (`icsEvent`), the matched ticket PDF's
 * filename (`ticketPdfName`, nullable), the target calendar's own
 * `timeZone`, an injected `formatDate` (see
 * ticketInstantToWallClockComponents), and `bodyText` (the shared
 * ticketingMessageBodyText(message) rendering — ICS BUILDER CONTRACT, see
 * processTicketFromIcsAttachment's own JSDoc). Pure.
 *
 * `eventName` = `String(icsEvent.summary || '').trim()` VERBATIM — UNLIKE
 * Ticketportal.cz's own builder, this does NOT strip any baked-in date
 * suffix: WebTicket's SUMMARY carries no such tail (D-03). Throws a clear
 * Error naming the portal when that comes back empty — a blank SUMMARY is
 * not a usable event name.
 * `location` (D-04) = the text of `icsEvent.location` BEFORE its first `|`,
 * trimmed — the venue/address segment only, since Google Calendar turns
 * `location` into a Maps lookup and the `| website | tel: ...` tail is
 * contact info, not an address. Falls back to the whole trimmed LOCATION
 * when that first segment is empty, and to `''` when LOCATION is absent
 * entirely. The FULL trimmed LOCATION line is still kept in the
 * description (D-08), so nothing is lost.
 * Start comes from `ticketInstantToWallClockComponents(icsEvent.start,
 * timeZone, formatDate)`.
 * `end` is set ONLY when `icsEvent.end` is a Date strictly LATER than
 * `icsEvent.start` — the same rule as buildTicketportalCzParsedTicket. The
 * real DTEND equals DTSTART, so `end` is omitted here and
 * buildTicketCalendarEventResource's existing DEFAULT_EVENT_DURATION_MINUTES
 * fallback applies (19:00-21:00) — copying the zero-length span would
 * reproduce the very zero-length event the generic ICS import used to
 * create.
 * `ticketIdentifier` (D-05) = extractWebticketOrderNumberFromPdfName(ticketPdfName)
 * when truthy, else `parseWebticketOrderText(bodyText).orderNumber` (the
 * body fallback), else `null`.
 * `description` (D-08) — built from parseWebticketOrderText(bodyText),
 * paragraphs joined by `'\n\n'`:
 *   1. eventName
 *   2. the FULL trimmed LOCATION, only if non-empty
 *   3. seatLines joined by `'\n'`, only if any
 *   4. the order lines joined by `'\n'`, only if any — `'Zakázka #' +
 *      ticketIdentifier + (orderReference ? ' / ' + orderReference : '')`
 *      when ticketIdentifier is truthy, then totalPriceLine when present.
 *      The "Zakázka" label is written in proper case regardless of which
 *      body rendering supplied the digits.
 *
 * Return shape: `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier, description }`, plus `end` only per the rule above.
 */
function buildWebticketParsedTicket(icsEvent, ticketPdfName, timeZone, formatDate, bodyText) {
  const eventName = String(icsEvent.summary || '').trim();
  if (!eventName) {
    throw new Error('WebTicket.cz: .ics SUMMARY is empty, no usable event name. Raw summary: "' + icsEvent.summary + '"');
  }

  const rawLocation = String(icsEvent.location || '');
  const firstSegment = rawLocation.split('|')[0].trim();
  const location = firstSegment || rawLocation.trim();
  const trimmedFullLocation = rawLocation.trim();

  const startComponents = ticketInstantToWallClockComponents(icsEvent.start, timeZone, formatDate);

  const orderText = parseWebticketOrderText(bodyText);
  const ticketIdentifier = extractWebticketOrderNumberFromPdfName(ticketPdfName) || orderText.orderNumber || null;

  const descriptionParagraphs = [eventName];
  if (trimmedFullLocation) {
    descriptionParagraphs.push(trimmedFullLocation);
  }
  if (orderText.seatLines.length > 0) {
    descriptionParagraphs.push(orderText.seatLines.join('\n'));
  }
  const orderLines = [];
  if (ticketIdentifier) {
    orderLines.push('Zakázka #' + ticketIdentifier + (orderText.orderReference ? ' / ' + orderText.orderReference : ''));
  }
  if (orderText.totalPriceLine) {
    orderLines.push(orderText.totalPriceLine);
  }
  if (orderLines.length > 0) {
    descriptionParagraphs.push(orderLines.join('\n'));
  }

  const parsedTicket = {
    eventName: eventName,
    location: location,
    year: startComponents.year,
    month: startComponents.month,
    day: startComponents.day,
    hour: startComponents.hour,
    minute: startComponents.minute,
    ticketIdentifier: ticketIdentifier,
    description: descriptionParagraphs.join('\n\n'),
  };

  if (icsEvent.end instanceof Date && icsEvent.end.getTime() > icsEvent.start.getTime()) {
    parsedTicket.end = ticketInstantToWallClockComponents(icsEvent.end, timeZone, formatDate);
  }

  return parsedTicket;
}

// =============================================================================
// GOOUT (info@mg.goout.net) — the NINTH supported portal, BODY-SOURCED from
// schema.org microdata (quick-261029-pmx)
// =============================================================================
//
// STRUCTURAL FACTS (from the real sample .eml, decoded by the planner):
//   - The confirmation has NO text/plain part at all (the Fever situation).
//     The raw HTML body is the ONLY input, read via TICKET_BODY_RAW_HTML_SENDERS
//     / ticketingBodyInputForSender above (D-01).
//   - The event data lives entirely in FIVE `<div itemscope
//     itemtype=".../EventReservation">` blocks (one per ticket), each with
//     THREE `itemprop="name"` properties (buyer, event, venue, in that
//     order) — read via the generic ticketingReadMicrodataItems (D-02).
//   - The start (`reservationFor.startDate`) is OFFSET-QUALIFIED
//     (`2026-12-19T17:00:00+01:00`) and there is NO `endDate` anywhere.
//   - `address.postalCode`, `address.addressRegion` and
//     `address.addressCountry` are all filled with the literal `CZE`
//     (deliberately ignored, D-04).
//   - `venueSeat`, `venueRow` and `venueSection` are all EMPTY — general
//     admission; the real section/price exist only in the VISIBLE HTML
//     (D-09).
//   - Each per-ticket `ticketNumber` IS the QR payload printed on the
//     per-ticket PNG — a LIVE entry credential (D-13).
//   - Attachments: five `image/png` QR codes (attachment disposition), the
//     real ticket PDF (filename carries NO order number, only a
//     date/time/random-token tail), and a SEPARATE `receipt.pdf` (D-11).
//   - The `.pkpass` Apple Wallet links are BEARER links, never followed
//     (D-18).
//   - The visible `Sobota 19. 12. 17:00–18:05` line uses U+00A0 no-break
//     spaces and a U+2013 en dash — ticketingHtmlToPlainText folds the
//     NBSPs to plain spaces (D-08).
//   - GoOut is sent via Mailgun and markets through this same pipeline (a
//     promo block rides along with every real confirmation) — the same
//     incident class as Ticketmaster CZ/Fever/smsticket.cz (D-10).
//   - GoOut is a PLATFORM selling for many venues (this sample: Červený
//     kostel, Brno) — the one sender covers every venue on it, the same
//     Entradio-style platform shape.

/**
 * gooutCleanValue (D-04, quick-261029-pmx) — pure; whitespace-collapses
 * (`\s` includes U+00A0) and trims a string value, or returns `''` for a
 * non-string. Used throughout this section to normalize raw microdata
 * values before comparing or joining them.
 */
function gooutCleanValue(value) {
  if (typeof value !== 'string') {
    return '';
  }
  return value.replace(/\s+/g, ' ').trim();
}

/**
 * gooutNestedItem (D-04, quick-261029-pmx) — pure; returns `value`
 * unchanged when it is a non-null object, else `{}` — a defensive
 * default so a missing nested microdata item (reservationFor, location,
 * address, underName) never throws when ITS OWN properties are read.
 */
function gooutNestedItem(value) {
  return value && typeof value === 'object' ? value : {};
}

/**
 * gooutLocationFromPlace (D-04, quick-261029-pmx) — pure; builds the
 * Calendar-event `location` string from a schema.org Place item: the
 * DISTINCT non-empty cleaned `[place.name, address.streetAddress,
 * address.addressLocality]`, joined with `', '`, or `''` when none are
 * present. `postalCode`, `addressRegion` and `addressCountry` are
 * deliberately IGNORED — the real sample fills all three with the literal
 * `'CZE'`. A missing location is not an error (the same WebTicket-style
 * absent-LOCATION rule).
 */
function gooutLocationFromPlace(place) {
  const placeItem = gooutNestedItem(place);
  const address = gooutNestedItem(placeItem.address);

  const candidates = [gooutCleanValue(placeItem.name), gooutCleanValue(address.streetAddress), gooutCleanValue(address.addressLocality)];

  const seen = {};
  const parts = [];
  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    if (candidate && !seen[candidate]) {
      seen[candidate] = true;
      parts.push(candidate);
    }
  }

  return parts.join(', ');
}

// GOOUT_START_DATE_PATTERN (D-04, quick-261029-pmx) — only the
// offset-qualified shape OBSERVED in the real sample
// ("2026-12-19T17:00:00+01:00") is observed live; a trailing 'Z', missing
// seconds/fractional seconds, and a missing offset entirely are TOLERANCE
// (never seen live, but accepted defensively — ISO 8601 allows all of
// them).
const GOOUT_START_DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:\d{2})?$/;

// GOOUT_CONFIRMED_STATUS_PATTERN (D-10, quick-261029-pmx) — the OBSERVED
// status is the literal 'http://schema.org/Confirmed' (a `<link href>`);
// a bare 'Confirmed' or a 'ReservationConfirmed' suffix are TOLERANCE
// (both spellings exist in schema.org reservation vocabularies), anchored
// so a 'NotConfirmed'/'Cancelled'-shaped value never matches.
const GOOUT_CONFIRMED_STATUS_PATTERN = /(?:^|\/)(?:Reservation)?Confirmed$/;

// GOOUT_VISIBLE_DATE_TIME_PATTERN (D-08, quick-261029-pmx) — the OBSERVED
// shape is "D. M. HH:MM–HH:MM" with no-break spaces around the dots
// (folded to plain spaces by ticketingHtmlToPlainText, which this pattern
// is applied AFTER) and a U+2013 EN DASH between the two times. An
// optional four-digit year between the date and the start time, and a
// plain hyphen or em dash time separator, are TOLERANCE — never observed
// live, but harmless to accept. Non-global; a single `.exec()` always
// returns the FIRST match.
const GOOUT_VISIBLE_DATE_TIME_PATTERN = /(\d{1,2})\.\s*(\d{1,2})\.\s*(?:\d{4}\s+)?(\d{1,2}):(\d{2})(?:\s*[-–—]\s*(\d{1,2}):(\d{2}))?/;

// GOOUT_TICKET_DETAIL_WINDOW_LINES (D-09, quick-261029-pmx) — the maximum
// number of surviving lines scanned AFTER each ticket number's own line
// for a bare Kč price line; 3 is wide enough to span the empty seat/row
// line and the "Sekce C2" line in the real fixture, while staying well
// short of the next ticket's own price line.
const GOOUT_TICKET_DETAIL_WINDOW_LINES = 3;

/**
 * gooutVisibleDateTime (D-08, quick-261029-pmx) — pure; cross-validates
 * the microdata start against the VISIBLE text's own date/time line —
 * the only place GoOut's confirmation states an end time at all (there is
 * no `endDate` anywhere in the microdata). `lines` is
 * `ticketingHtmlToPlainText(rawHtml).split('\n')`; `start` is the
 * microdata-derived `{ year, month (ZERO-indexed), day, hour, minute }`.
 *
 * Returns `{ dateTimeText, end }`: the FIRST line whose
 * GOOUT_VISIBLE_DATE_TIME_PATTERN match has day/month(1-indexed)/hour/
 * minute ALL equal to `start`'s own digits wins, and `dateTimeText` is
 * that WHOLE LINE (not just the matched substring). When that line's own
 * end-time group is also present, `delta` is the end minus the start in
 * minutes: `delta < 0` rolls past midnight (+1440), `delta === 0` yields
 * NO `end` (a zero-length span is never copied — the WebTicket lesson),
 * `delta > 0` yields `end = addMinutesToWallClockComponents(start,
 * delta)`.
 *
 * When NO line validates, `dateTimeText` is synthesized as
 * `'<day>. <month>. <year> <HH>:<MM>'`, and `end` stays `undefined` (the
 * shared 120-minute DEFAULT_EVENT_DURATION_MINUTES fallback then applies
 * in buildTicketCalendarEventResource).
 */
function gooutVisibleDateTime(lines, start) {
  const list = lines || [];
  const startMonthOneIndexed = start.month + 1;

  for (let i = 0; i < list.length; i++) {
    const match = GOOUT_VISIBLE_DATE_TIME_PATTERN.exec(list[i]);
    if (!match) {
      continue;
    }

    const matchedDay = Number(match[1]);
    const matchedMonth = Number(match[2]);
    const matchedHour = Number(match[3]);
    const matchedMinute = Number(match[4]);

    if (matchedDay !== start.day || matchedMonth !== startMonthOneIndexed || matchedHour !== start.hour || matchedMinute !== start.minute) {
      continue;
    }

    const result = { dateTimeText: list[i] };

    if (match[5] !== undefined && match[6] !== undefined) {
      const endHour = Number(match[5]);
      const endMinute = Number(match[6]);
      let delta = endHour * 60 + endMinute - (start.hour * 60 + start.minute);
      if (delta < 0) {
        delta += 1440;
      }
      if (delta > 0) {
        result.end = addMinutesToWallClockComponents(
          { year: start.year, month: start.month, day: start.day, hour: start.hour, minute: start.minute },
          delta
        );
      }
    }

    return result;
  }

  return {
    dateTimeText:
      start.day +
      '. ' +
      startMonthOneIndexed +
      '. ' +
      start.year +
      ' ' +
      zeroPadTicketComponent(start.hour, 2) +
      ':' +
      zeroPadTicketComponent(start.minute, 2),
  };
}

/**
 * gooutTicketLines (D-09, quick-261029-pmx) — pure; builds the
 * description's per-ticket detail lines. `lines` is the same
 * ticketingHtmlToPlainText-rendered array gooutVisibleDateTime reads;
 * `ticketNumbers` are the DISTINCT cleaned `ticketNumber` values, in item
 * order (venueSection is empty in the microdata — the section and price
 * exist only in the VISIBLE HTML, reusing this file's shared
 * SMSTICKET_PRICE_LINE_PATTERN by name).
 *
 * For each ticket number, finds its own EXACT line in `lines` and scans
 * the next GOOUT_TICKET_DETAIL_WINDOW_LINES lines for the first line
 * matching SMSTICKET_PRICE_LINE_PATTERN. When found, the returned line is
 * every line strictly between the ticket-number line and the price line,
 * plus the price and the ticket number itself, all joined by `' | '`
 * (giving `'Sekce C2 | 490 Kč | 1111111111111'` for the real fixture, or
 * `'Řada 5, místo 12 | Sekce C2 | 490 Kč | …'` for a seated layout).
 * Otherwise the BARE ticket number is used (graceful degradation). The
 * 3-line window can never reach the next ticket's own price line.
 */
function gooutTicketLines(lines, ticketNumbers) {
  const list = lines || [];

  return (ticketNumbers || []).map(function (ticketNumber) {
    const lineIndex = list.indexOf(ticketNumber);
    if (lineIndex === -1) {
      return ticketNumber;
    }

    for (let offset = 1; offset <= GOOUT_TICKET_DETAIL_WINDOW_LINES; offset++) {
      const priceLineIndex = lineIndex + offset;
      if (priceLineIndex >= list.length) {
        break;
      }
      if (SMSTICKET_PRICE_LINE_PATTERN.test(list[priceLineIndex])) {
        const between = list.slice(lineIndex + 1, priceLineIndex);
        return between.concat([list[priceLineIndex], ticketNumber]).join(' | ');
      }
    }

    return ticketNumber;
  });
}

/**
 * parseGooutTicketHtml (D-04, D-06, D-07, D-10; D-08/D-09 enrichment added
 * quick-261029-pmx Task 2) — the info@mg.goout.net ticket-BODY parser.
 * `html` is the message's RAW HTML body (TICKET_BODY_RAW_HTML_SENDERS /
 * ticketingBodyInputForSender, D-01) — GoOut's event data lives entirely
 * in schema.org EventReservation microdata ATTRIBUTES, which every
 * plain-text rendering strips. `referenceDate` and `subject` exist only
 * for the shared BODY-PARSER CONTRACT and are IGNORED: the year comes
 * from the microdata's own offset-qualified startDate, so there is no
 * year inference and no month-name table, unlike Fever.
 *
 * CORE STEPS (from the microdata):
 *   1. `items = ticketingReadMicrodataItems(rawHtml, 'EventReservation')`.
 *      Throws when empty.
 *   2. CONSISTENCY (D-06): every item must agree with item 0 on the
 *      cleaned reservationNumber, reservationFor.name and
 *      reservationFor.startDate — otherwise throws naming "more than one
 *      distinct reservation"; blocks are NEVER merged and a disagreement
 *      is never silently dropped.
 *   3. `event = gooutNestedItem(items[0].reservationFor)`; `eventName =
 *      gooutCleanValue(event.name)` is REQUIRED.
 *   4. The start is matched against GOOUT_START_DATE_PATTERN on
 *      `gooutCleanValue(event.startDate)`; month/day/hour/minute are each
 *      range-validated. `month` in the RETURN shape is ZERO-indexed, and
 *      seconds are dropped. `utcOffset` (D-05) is set as an OWN property
 *      ONLY when the matched offset group is present.
 *   5. `location` = gooutLocationFromPlace(event.location) (D-04).
 *   6. `ticketIdentifier` (D-07) = `gooutCleanValue(items[0].reservationNumber)
 *      || null` — the ORDER number, identical across every per-ticket
 *      block, NEVER a per-seat ticketNumber (which differs per block),
 *      the PDF filename (carries no order number) or ticketToken (a
 *      credential).
 *
 * VISIBLE-TEXT ENRICHMENT LAYER (D-08, D-09, quick-261029-pmx Task 2) —
 * read ONLY for the end time, the date/time line and each ticket's
 * section/price, NEVER for any core field above:
 *   7. `lines = ticketingHtmlToPlainText(rawHtml).split('\n')`.
 *   8. `gooutVisibleDateTime(lines, <the Task 1 start components>)` gives
 *      `{ dateTimeText, end }`, cross-validated against the microdata
 *      start so a mismatched or absent visible line can never produce a
 *      wrong end time (D-08).
 *   9. The DISTINCT cleaned `ticketNumber` values, in item order, are
 *      passed to gooutTicketLines(lines, ...) for the per-ticket
 *      section/price detail lines (D-09).
 *  10. `description` (D-09) joins, with `'\n\n'`: eventName; location,
 *      only if non-empty; dateTimeText (ALWAYS present, synthesized when
 *      no visible line validates); the ticket lines joined by `'\n'`,
 *      only if any; `'Objednávka č. ' + ticketIdentifier`, only if
 *      truthy.
 *
 * Every thrown message starts with 'Unrecognized GoOut ticket HTML: '
 * (or is the D-06 consistency message, or a '<Field> out of range'
 * message) and ALWAYS ends with '. Full raw HTML body:\n' + rawHtml —
 * this file's DIAGNOSTIC-ON-FAILURE convention, with the raw HTML because
 * that is this parser's own input.
 *
 * Return shape: `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier, description }`, plus `utcOffset` and/or `end` only
 * when present.
 */
function parseGooutTicketHtml(html, referenceDate, subject) {
  const rawHtml = String(html || '');
  const diagnosticSuffix = '. Full raw HTML body:\n' + rawHtml;

  const items = ticketingReadMicrodataItems(rawHtml, 'EventReservation');
  if (items.length === 0) {
    throw new Error('Unrecognized GoOut ticket HTML: no schema.org EventReservation microdata found' + diagnosticSuffix);
  }

  const firstReservationNumber = gooutCleanValue(items[0].reservationNumber);
  const firstEvent = gooutNestedItem(items[0].reservationFor);
  const firstEventName = gooutCleanValue(firstEvent.name);
  const firstStartDate = gooutCleanValue(firstEvent.startDate);

  for (let i = 1; i < items.length; i++) {
    const event = gooutNestedItem(items[i].reservationFor);
    if (
      gooutCleanValue(items[i].reservationNumber) !== firstReservationNumber ||
      gooutCleanValue(event.name) !== firstEventName ||
      gooutCleanValue(event.startDate) !== firstStartDate
    ) {
      throw new Error(
        'Unrecognized GoOut ticket HTML: the email carries more than one distinct reservation (reservation number, event name or start differ between EventReservation blocks); one event per purchase cannot be built safely' +
          diagnosticSuffix
      );
    }
  }

  const eventName = firstEventName;
  if (!eventName) {
    throw new Error(
      'Unrecognized GoOut ticket HTML: the EventReservation microdata carries no event name (reservationFor.name)' + diagnosticSuffix
    );
  }

  const startMatch = GOOUT_START_DATE_PATTERN.exec(firstStartDate);
  if (!startMatch) {
    throw new Error('Unrecognized GoOut ticket HTML: unrecognized reservationFor.startDate "' + firstStartDate + '"' + diagnosticSuffix);
  }

  const year = Number(startMatch[1]);
  const month = Number(startMatch[2]);
  const day = Number(startMatch[3]);
  const hour = Number(startMatch[4]);
  const minute = Number(startMatch[5]);
  const utcOffset = startMatch[6];

  if (month < 1 || month > 12) {
    throw new Error('Month out of range (' + month + ') in GoOut reservationFor.startDate' + diagnosticSuffix);
  }
  if (day < 1 || day > 31) {
    throw new Error('Day out of range (' + day + ') in GoOut reservationFor.startDate' + diagnosticSuffix);
  }
  if (hour < 0 || hour > 23) {
    throw new Error('Hour out of range (' + hour + ') in GoOut reservationFor.startDate' + diagnosticSuffix);
  }
  if (minute < 0 || minute > 59) {
    throw new Error('Minute out of range (' + minute + ') in GoOut reservationFor.startDate' + diagnosticSuffix);
  }

  const location = gooutLocationFromPlace(firstEvent.location);
  const ticketIdentifier = firstReservationNumber || null;
  const zeroIndexedMonth = month - 1;

  // VISIBLE-TEXT ENRICHMENT LAYER (D-08, D-09) -- read only for the end
  // time, the date/time line and each ticket's section/price, never for
  // any core field above.
  const lines = ticketingHtmlToPlainText(rawHtml).split('\n');
  const visible = gooutVisibleDateTime(lines, { year: year, month: zeroIndexedMonth, day: day, hour: hour, minute: minute });

  const ticketNumbers = [];
  const seenTicketNumbers = {};
  for (let i = 0; i < items.length; i++) {
    const cleanedTicketNumber = gooutCleanValue(items[i].ticketNumber);
    if (cleanedTicketNumber && !seenTicketNumbers[cleanedTicketNumber]) {
      seenTicketNumbers[cleanedTicketNumber] = true;
      ticketNumbers.push(cleanedTicketNumber);
    }
  }
  const ticketLines = gooutTicketLines(lines, ticketNumbers);

  const descriptionParagraphs = [eventName];
  if (location) {
    descriptionParagraphs.push(location);
  }
  descriptionParagraphs.push(visible.dateTimeText);
  if (ticketLines.length > 0) {
    descriptionParagraphs.push(ticketLines.join('\n'));
  }
  if (ticketIdentifier) {
    descriptionParagraphs.push('Objednávka č. ' + ticketIdentifier);
  }

  const parsedTicket = {
    eventName: eventName,
    location: location,
    year: year,
    month: zeroIndexedMonth,
    day: day,
    hour: hour,
    minute: minute,
    ticketIdentifier: ticketIdentifier,
    description: descriptionParagraphs.join('\n\n'),
  };

  if (utcOffset) {
    parsedTicket.utcOffset = utcOffset;
  }
  if (visible.end) {
    parsedTicket.end = visible.end;
  }

  return parsedTicket;
}

/**
 * gooutHtmlHasConfirmedReservation (D-10, quick-261029-pmx) — the
 * REQUIRED content detector for info@mg.goout.net: GoOut is a bulk ESP
 * (Mailgun), and the real confirmation itself carries GoOut's own promo
 * block, so GoOut markets through this same pipeline (the same incident
 * class as Ticketmaster CZ/Fever/smsticket.cz). Returns a LITERAL
 * boolean, never throws: `true` when
 * `ticketingReadMicrodataItems(html, 'EventReservation')` is non-empty
 * AND every item's `reservationStatus` is absent, `''`, or matches
 * GOOUT_CONFIRMED_STATUS_PATTERN.
 *
 * Marketing event listings carry schema.org `Event`, not
 * `EventReservation`, and are rejected by the exact type match. A
 * cancellation carrying an EventReservation with an explicit
 * non-confirmed status must NEVER create an event, so it is rejected
 * here.
 *
 * Completeness (a broken name or start) is deliberately NOT checked here:
 * a reservation-marked mail with missing/inconsistent fields is ADMITTED
 * and fails LOUDLY in parseGooutTicketHtml instead, carrying the raw HTML
 * in the failure email — rejecting it here would skip it silently,
 * against this file's core "no silent failures" value. An ABSENT status
 * is also admitted, for the same reason: it is likelier a template
 * variant of a confirmation than a cancellation.
 *
 * SCOPE LIMITATION: cancellation handling (deleting an already-created
 * event) and a changed-date notice for an already-processed order (dedup
 * silently skips it, keeping the event's old time) both need real data to
 * design.
 */
function gooutHtmlHasConfirmedReservation(html) {
  const items = ticketingReadMicrodataItems(String(html || ''), 'EventReservation');
  if (items.length === 0) {
    return false;
  }

  for (let i = 0; i < items.length; i++) {
    const status = items[i].reservationStatus;
    if (status && !GOOUT_CONFIRMED_STATUS_PATTERN.test(status)) {
      return false;
    }
  }

  return true;
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
// DELIBERATE ABSENCE: there is NO 'vstupenky@ticketportal.cz' key here
// (quick-261005-orv). Ticketportal.cz is ICS-SOURCED -- its event data comes
// from its order confirmation's own .ics attachment, parsed by the generic
// parseIcs, never by a per-portal body parser. See TICKET_ICS_MODE_SENDERS
// below. smsticket.cz (quick-261026-ma6) IS body-sourced and DOES have a key
// here (see below) -- it is NOT ICS-sourced, since the real confirmation
// carries no .ics attachment at all. Also NO 'automat@webticket.cz' key
// (quick-261028-h1n, D-01) -- WebTicket.cz is ALSO ICS-SOURCED, like
// Ticketportal.cz; its body is read only as an ENRICHMENT source inside
// buildWebticketParsedTicket (parseWebticketOrderText), never via this
// registry, which would otherwise route it to the body branch FIRST.
const TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL = {
  'rezervace@kinoart.cz': parseKinoArtTicketText,
  'noreply@ticketmaster.cz': parseTicketmasterCzTicketText,
  'no-reply@app.entradio.cz': parseEntradioTicketText,
  'hello@feverup.com': parseFeverTicketText,
  // The seventh portal, smsticket.cz (quick-261026-ma6).
  'info@smsticket.cz': parseSmsticketTicketText,
  // The ninth portal, GoOut (quick-261029-pmx) -- BODY-SOURCED, but read
  // from the RAW HTML via TICKET_BODY_RAW_HTML_SENDERS (D-01), never a
  // plain-text rendering.
  'info@mg.goout.net': parseGooutTicketHtml,
};

/**
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL — the local (single-file)
 * registry mapping a ticketing portal's `identifyingEmail` (lowercased)
 * to its OCR-text parser function. This is what lets ALL portal parsers
 * live in one file while still routing a matched TICKETING_PORTALS entry
 * to the RIGHT parser — adding a future portal means adding one new
 * parser function plus one new entry here, nothing else.
 */
// DELIBERATE ABSENCE: there is NO 'hello@feverup.com' key here (D-01). Fever
// is body-sourced -- everything the calendar event needs is already in the
// email body, so no Drive upload and no OCR conversion is needed for it.
// Also NO 'vstupenky@ticketportal.cz' key (quick-261005-orv) -- it is
// ICS-SOURCED, not OCR-sourced; see TICKET_ICS_MODE_SENDERS below. Also NO
// 'info@smsticket.cz' key (D-01, quick-261026-ma6) -- same reason as Fever:
// everything smsticket.cz's event needs is already in its body (or its HTML
// body's own rendering, see ticketingMessageBodyText), so no OCR is needed.
// Also NO 'automat@webticket.cz' key (quick-261028-h1n, D-01) -- same reason
// as Ticketportal.cz: WebTicket.cz's event data comes from its own .ics
// attachment, never Drive/OCR. Also NO 'info@mg.goout.net' key
// (quick-261029-pmx) -- no OCR is needed; the HTML's own schema.org
// microdata carries everything (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL
// above).
const TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL = {
  'no-reply@enigoo.cz': parseEnigooTicketText,
};

/**
 * TICKET_ICS_MODE_SENDERS — the ICS-SOURCED counterpart to the two parser
 * registries above (quick-261005-orv, D-02; a second portal, WebTicket.cz,
 * added quick-261028-h1n, D-01). A portal in this list has its event data
 * extracted from its own `.ics` attachment via the generic,
 * already-existing parseIcs (src/05-action-ics-import.js) — there is
 * deliberately NO per-portal parser FUNCTION for this mode, since parseIcs
 * is itself the parser. A message's sender resolving into this list AND
 * carrying at least one qualifying `.ics` attachment
 * (findTicketIcsAttachments) is what admits the `mode: 'ics'` job (see
 * resolveTicketProcessingJobs / TICKETING_PORTALS_ACTION.appliesTo/run).
 */
// DELIBERATE ABSENCE: there is NO 'info@mg.goout.net' entry here
// (quick-261029-pmx) -- the real confirmation carries no .ics attachment
// at all; its event data comes entirely from the raw HTML's own
// schema.org microdata (TICKET_BODY_RAW_HTML_SENDERS above).
const TICKET_ICS_MODE_SENDERS = ['vstupenky@ticketportal.cz', 'automat@webticket.cz'];

/**
 * TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL (quick-261028-h1n,
 * D-02) — the local (single-file) registry mapping an ICS-SOURCED ticketing
 * portal's `identifyingEmail` (lowercased) to its own parsed-ticket BUILDER
 * function — the per-sender post-processing step that turns a single
 * parseIcs-returned event into this file's shared `parsedTicket` shape.
 *
 * WHY A REGISTRY, NOT A HARDCODED CALL: before quick-261028-h1n,
 * processTicketFromIcsAttachment hardcoded a single call to
 * buildTicketportalCzParsedTicket — correct when Ticketportal.cz was the
 * ONLY ics-mode portal, but wrong the moment a second one (WebTicket.cz)
 * needed its own builder. This registry is keyed the SAME way as
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL / TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL
 * above, so adding a future ICS-sourced portal means registering a new
 * builder here, nothing else.
 *
 * parseIcs REMAINS THE ONLY `.ics` PARSER: every registered builder only
 * POST-PROCESSES what parseIcs already returned (the `icsEvent` argument);
 * none of them re-parses the `.ics` text itself.
 *
 * INVARIANT: `Object.keys(TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL)`
 * ALWAYS deepEquals TICKET_ICS_MODE_SENDERS — a test pins this, so the
 * admission array and the builder registry can never drift apart.
 *
 * Signature (quick-261028-h1n Task 2 extends this to FIVE arguments, D-08):
 * `(icsEvent, ticketPdfName, timeZone, formatDate, bodyText) -> parsedTicket`.
 * buildTicketportalCzParsedTicket takes only the first four and ignores a
 * fifth — "extra trailing argument is inert" per this file's existing
 * BODY-PARSER CONTRACT pattern.
 */
const TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL = {
  'vstupenky@ticketportal.cz': buildTicketportalCzParsedTicket,
  'automat@webticket.cz': buildWebticketParsedTicket,
};

/**
 * DEFAULT_EVENT_DURATION_MINUTES — the fixed default duration (2 hours)
 * added to a parsed ticket's start time when a portal's PDF carries no
 * explicit end time. A per-portal-parser default to reach for, not a rule
 * every future portal is forced through — a FUTURE portal whose PDF DOES
 * include an end time should use it directly instead. Ticketportal.cz
 * (quick-261005-orv) is exactly that future portal: its .ics DTEND supplies
 * a real end time via `parsedTicket.end` (see buildTicketCalendarEventResource
 * below), so this default never applies to it.
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

// TICKET_WALL_CLOCK_FORMAT_PATTERN (quick-261005-orv, D-04c) — the Java
// SimpleDateFormat pattern GAS's `Utilities.formatDate(date, timeZone,
// pattern)` expects, used to convert a .ics-parsed UTC Date instant into
// wall-clock digits in a given IANA timeZone. Shared by
// ticketInstantToWallClockComponents and its production caller
// (processTicketFromIcsAttachment) so the pattern string and its parser
// below can never drift apart.
const TICKET_WALL_CLOCK_FORMAT_PATTERN = 'yyyy-MM-dd HH:mm';

/**
 * ticketInstantToWallClockComponents — converts a UTC Date instant
 * (`date`, e.g. a parseIcs event's `start`/`end`) into a `{ year, month
 * (ZERO-indexed), day, hour, minute }` wall-clock components object in
 * `timeZone`, via an INJECTED `formatDate(date, timeZone, pattern)`
 * function — production passes a thin wrapper around GAS's
 * `Utilities.formatDate` (unavailable under Node); tests pass an
 * Intl-based stand-in. This injection seam is what keeps this function
 * itself pure and unit-testable (quick-261005-orv, D-04c).
 *
 * FAILS CLOSED: throws a clear Error quoting the raw formatter output if
 * it does not match `TICKET_WALL_CLOCK_FORMAT_PATTERN`'s own shape
 * (`/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/`), and throws if `date` is
 * not a valid Date — silently returning NaN-laced components would risk
 * writing a wrong calendar event rather than failing the action loudly.
 * Pure given its injected `formatDate`; no direct GAS global reference.
 */
function ticketInstantToWallClockComponents(date, timeZone, formatDate) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    throw new Error('ticketInstantToWallClockComponents: "date" is not a valid Date: ' + date);
  }

  const formatted = formatDate(date, timeZone, TICKET_WALL_CLOCK_FORMAT_PATTERN);
  const match = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(String(formatted));
  if (!match) {
    throw new Error(
      'ticketInstantToWallClockComponents: formatDate returned an unrecognized shape (expected "' +
        TICKET_WALL_CLOCK_FORMAT_PATTERN +
        '"): "' +
        formatted +
        '"'
    );
  }

  return {
    year: Number(match[1]),
    month: Number(match[2]) - 1,
    day: Number(match[3]),
    hour: Number(match[4]),
    minute: Number(match[5]),
  };
}

/**
 * buildTicketportalCzParsedTicket — builds the `parsedTicket` shape
 * (quick-261005-orv, D-04, D-05) from a single parseIcs-returned event
 * (`icsEvent`), the matched eTicket PDF's filename (`ticketPdfName`,
 * nullable), the target calendar's own `timeZone`, and an injected
 * `formatDate` (see ticketInstantToWallClockComponents). Pure.
 *
 * Returns `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier }`, PLUS `end: { year, month, day, hour, minute }` ONLY
 * when `icsEvent.end` is a Date strictly LATER than `icsEvent.start` —
 * parseVeventBlock sets `end = start` when the .ics carries no DTEND, and
 * omitting `end` in that case is what lets buildTicketCalendarEventResource's
 * existing DEFAULT_EVENT_DURATION_MINUTES fallback apply exactly as it does
 * for every other portal (D-05).
 *
 * `eventName` = stripTicketportalCzSummaryDateSuffix(icsEvent.summary) —
 * throws a clear Error if that comes back empty (a stripped-to-nothing
 * summary is not a usable event name).
 * `location` = icsEvent.location, or '' when absent.
 * `ticketIdentifier` = extractTicketportalCzOrderNumber(ticketPdfName) —
 * `null` when absent, NEVER icsEvent.uid in any fallback (D-04d; see that
 * function's own JSDoc for why the UID is unsafe to use here).
 *
 * Deliberately sets no `description`: the .ics DESCRIPTION is boilerplate
 * ("Přidejte si do Vašeho kalendáře.") plus parseIcs's own organizer line,
 * which adds nothing an owner reading the calendar event needs.
 */
function buildTicketportalCzParsedTicket(icsEvent, ticketPdfName, timeZone, formatDate) {
  const eventName = stripTicketportalCzSummaryDateSuffix(icsEvent.summary);
  if (!eventName) {
    throw new Error('Ticketportal.cz: stripped .ics SUMMARY is empty, no usable event name. Raw summary: "' + icsEvent.summary + '"');
  }

  const startComponents = ticketInstantToWallClockComponents(icsEvent.start, timeZone, formatDate);

  const parsedTicket = {
    eventName: eventName,
    location: icsEvent.location || '',
    year: startComponents.year,
    month: startComponents.month,
    day: startComponents.day,
    hour: startComponents.hour,
    minute: startComponents.minute,
    ticketIdentifier: extractTicketportalCzOrderNumber(ticketPdfName),
  };

  if (icsEvent.end instanceof Date && icsEvent.end.getTime() > icsEvent.start.getTime()) {
    parsedTicket.end = ticketInstantToWallClockComponents(icsEvent.end, timeZone, formatDate);
  }

  return parsedTicket;
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
 * findTicketIcsAttachments (quick-261005-orv, D-02) — every GmailAttachment
 * on `message` matching isIcsAttachment (the EXISTING matcher from
 * src/05-action-ics-import.js, reused via this file's end-of-file Node
 * bridge under Node / the shared GAS global scope — never a second .ics
 * matcher), in source order, or `[]` if none match. Pure over the passed-in
 * `message`, so this is fully unit-testable under Node with a duck-typed
 * fake.
 */
function findTicketIcsAttachments(message) {
  return message.getAttachments().filter(isIcsAttachment);
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
 * findFeverTicketPdfAttachment — Fever's own ticket-PDF finder for the
 * OPTIONAL `insertPdfIntoEvent` attachment path (D-12). The real observed
 * email carries exactly one PDF attachment, filename
 * `order_125529770.pdf`. Returns the FIRST qualifying PDF attachment whose
 * name contains `"order_"`, or `null` — same discipline as
 * findKinoArtTicketPdfAttachment and findTicketmasterCzTicketPdfAttachment
 * above. SCOPE LIMITATION: a differently-named attachment would need
 * handling THEN, with real data. Pure, no GAS globals.
 */
function findFeverTicketPdfAttachment(message) {
  const pdfAttachments = findTicketPdfAttachments(message);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = pdfAttachments[i].getName() || '';
    if (name.indexOf('order_') !== -1) {
      return pdfAttachments[i];
    }
  }

  return null;
}

/**
 * findTicketportalCzTicketPdfAttachment (quick-261005-orv, D-03) —
 * Ticketportal.cz's own ticket-PDF finder, for BOTH the OPTIONAL
 * `insertPdfIntoEvent` attachment path AND the per-ORDER
 * extractTicketportalCzOrderNumber dedup key. The real observed order
 * confirmation carries THREE PDF attachments: `eTicket_<order>.pdf` (the
 * real ticket), `Doklad_poplatek_<order>.pdf` (a service-fee receipt, NOT
 * ticket data), and a venue terms-and-conditions PDF with a diacritic
 * filename (also NOT ticket data). Returns the FIRST qualifying PDF
 * attachment whose name contains `"eTicket"`, or `null` — same discipline
 * as findKinoArtTicketPdfAttachment/findTicketmasterCzTicketPdfAttachment/
 * findFeverTicketPdfAttachment above. SCOPE LIMITATION: a differently-named
 * attachment would need handling THEN, with real data. Pure, no GAS
 * globals.
 */
function findTicketportalCzTicketPdfAttachment(message) {
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
 * findSmsticketTicketPdfAttachment (D-10, quick-261026-ma6) —
 * smsticket.cz's own ticket-PDF finder, for the OPTIONAL
 * `insertPdfIntoEvent` attachment path. The real observed confirmation
 * carries exactly ONE PDF attachment, filename "smsticket - <event name>
 * .pdf" (note the real space before ".pdf") — holding EVERY ticket of the
 * purchase, not one PDF per seat. Returns the FIRST qualifying PDF
 * attachment whose name contains the literal "smsticket", or `null` — same
 * discipline as findKinoArtTicketPdfAttachment/
 * findTicketmasterCzTicketPdfAttachment/findFeverTicketPdfAttachment/
 * findTicketportalCzTicketPdfAttachment above.
 *
 * LIVE CAVEAT: the real filename arrives only via an RFC 2047
 * `Content-Type name=` MIME parameter (the `Content-Disposition` is a bare
 * `attachment` with no filename there at all). If GAS's `getName()` ever
 * returned that undecoded, this finder returns `null` and
 * processTicketFromMessageBody already logs and creates the event without
 * an attachment — never a thrown error.
 *
 * SCOPE LIMITATION: a differently-named attachment would need handling
 * THEN, with real data. Pure, no GAS globals.
 */
function findSmsticketTicketPdfAttachment(message) {
  const pdfAttachments = findTicketPdfAttachments(message);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = pdfAttachments[i].getName() || '';
    if (name.indexOf('smsticket') !== -1) {
      return pdfAttachments[i];
    }
  }

  return null;
}

/**
 * findWebticketTicketPdfAttachment (D-06, quick-261028-h1n) —
 * WebTicket.cz's own ticket-PDF finder. The real observed confirmation
 * carries exactly ONE PDF attachment, filename "webticket_<order>.pdf", a
 * plain ASCII filename (no RFC 2047 encoding, unlike smsticket.cz's own
 * finder's LIVE CAVEAT above) — holding the purchase's real ticket(s).
 * Returns the FIRST qualifying PDF attachment whose name contains the
 * literal "webticket_", or `null` — same discipline as
 * findKinoArtTicketPdfAttachment/findTicketmasterCzTicketPdfAttachment/
 * findFeverTicketPdfAttachment/findTicketportalCzTicketPdfAttachment/
 * findSmsticketTicketPdfAttachment above.
 *
 * ICS-MODE CONSULTS THIS FOR BOTH THE DEDUP IDENTIFIER AND THE ATTACHMENT:
 * processTicketFromIcsAttachment looks this up via
 * TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL REGARDLESS of
 * `insertPdfIntoEvent`, since extractWebticketOrderNumberFromPdfName needs
 * the matched PDF's filename even when the PDF itself will not be
 * attached — the Ticketportal.cz D-08 precedent.
 *
 * SCOPE LIMITATION: a differently-named attachment would need handling
 * THEN, with real data. Pure, no GAS globals.
 */
function findWebticketTicketPdfAttachment(message) {
  const pdfAttachments = findTicketPdfAttachments(message);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = pdfAttachments[i].getName() || '';
    if (name.indexOf('webticket_') !== -1) {
      return pdfAttachments[i];
    }
  }

  return null;
}

// GOOUT_RECEIPT_PDF_NAME_MARKER (D-11, quick-261029-pmx) -- the observed
// receipt filename is the literal 'receipt.pdf'; matched
// case-insensitively as a SUBSTRING so a differently-cased or
// differently-suffixed receipt name is still excluded.
const GOOUT_RECEIPT_PDF_NAME_MARKER = 'receipt';

// GOOUT_TICKET_PDF_NAME_PATTERN (D-11, quick-261029-pmx) -- the real
// ticket filename's STABLE tail is "(D.M.YYYY H:MM <token>).pdf" -- the
// event's own date and time, which parseGooutTicketHtml already knows,
// followed by a random lowercase token tolerated as any non-parenthesis
// run. The random token itself is NEVER keyed on.
const GOOUT_TICKET_PDF_NAME_PATTERN = /\(\s*\d{1,2}\.\s*\d{1,2}\.\s*\d{4}\s+\d{1,2}:\d{2}\b[^()]*\)\s*\.pdf$/i;

/**
 * findGooutTicketPdfAttachment (D-11, quick-261029-pmx) — GoOut's own
 * ticket-PDF finder for the OPTIONAL `insertPdfIntoEvent` attachment
 * path. The real confirmation carries the ticket PDF (filename template
 * "<event> (<date> <time> <token>).pdf", carrying NO order number at
 * all) AND a SEPARATE `receipt.pdf`.
 *
 * HARD RECEIPT EXCLUSION: any PDF attachment whose lowercased name
 * contains GOOUT_RECEIPT_PDF_NAME_MARKER is skipped outright, regardless
 * of its position among the attachments.
 *
 * Among the remaining PDFs, returns the FIRST whose name matches
 * GOOUT_TICKET_PDF_NAME_PATTERN (the template's own date/time tail), or
 * `null`. DELIBERATELY NO "only remaining PDF" FALLBACK: the attach path
 * renames the matched file to `"<event> - <date> - <order>.pdf"`, so a
 * wrong file would masquerade as the real ticket — a missing attachment
 * (the event is still created; the existing "no matching ticket PDF" log
 * line fires) is the cheaper failure. The positive match keys on the
 * event's own date and time, the template's STABLE part, never on the
 * random token, so any miss degrades gracefully.
 *
 * The five per-ticket QR PNGs are not PDFs and are never attached by
 * this finder — the ticket PDF alone carries every ticket. The filename
 * is plain ASCII, so there is no RFC 2047 caveat (unlike smsticket.cz's
 * own finder's LIVE CAVEAT).
 *
 * SCOPE LIMITATION: a differently-named attachment would need handling
 * THEN, with real data. Pure, no GAS globals.
 */
function findGooutTicketPdfAttachment(message) {
  const pdfAttachments = findTicketPdfAttachments(message);

  for (let i = 0; i < pdfAttachments.length; i++) {
    const name = pdfAttachments[i].getName() || '';
    if (name.toLowerCase().indexOf(GOOUT_RECEIPT_PDF_NAME_MARKER) !== -1) {
      continue;
    }
    if (GOOUT_TICKET_PDF_NAME_PATTERN.test(name)) {
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
// Ticketportal.cz (quick-261005-orv, D-08) and WebTicket.cz
// (quick-261028-h1n, D-06) are registered here too, despite being
// ICS-SOURCED rather than body-sourced: this registry's attach-step
// consumer (processTicketFromIcsAttachment) reuses it exactly as the
// body-sourced pipeline does, regardless of processing mode.
const TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL = {
  'rezervace@kinoart.cz': findKinoArtTicketPdfAttachment,
  'noreply@ticketmaster.cz': findTicketmasterCzTicketPdfAttachment,
  'hello@feverup.com': findFeverTicketPdfAttachment,
  'vstupenky@ticketportal.cz': findTicketportalCzTicketPdfAttachment,
  'info@smsticket.cz': findSmsticketTicketPdfAttachment,
  'automat@webticket.cz': findWebticketTicketPdfAttachment,
  // The ninth portal, GoOut (quick-261029-pmx, D-11) -- a hard receipt
  // exclusion, never an "only remaining PDF" fallback.
  'info@mg.goout.net': findGooutTicketPdfAttachment,
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
// DELIBERATE ABSENCE: there is NO 'hello@feverup.com' key here (D-11). The
// per-seat codes fold into `description` only; this portal makes no
// outbound HTTP call and needs no new OAuth scope. Also NO
// 'info@smsticket.cz' key (D-01, quick-261026-ma6) -- same reason: its
// entry codes fold into `description` only, its ticket PDF (when wanted)
// travels WITH the message and is found by findSmsticketTicketPdfAttachment
// above, and it makes no outbound HTTP call and needs no new OAuth scope.
// Also NO 'automat@webticket.cz' key (quick-261028-h1n) -- its ticket PDF
// travels WITH the message (findWebticketTicketPdfAttachment above); the
// order-status/wallet links in its body are bearer links and are NEVER
// followed, so this portal makes no outbound HTTP call and needs no new
// OAuth scope either. Also NO 'info@mg.goout.net' key (quick-261029-pmx,
// D-18) -- its ticket PDF travels WITH the message
// (findGooutTicketPdfAttachment above); the .pkpass Apple Wallet links in
// its body are bearer links and are NEVER followed or fetched; no new
// OAuth scope is needed.
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
  // hello@feverup.com also sends ordinary marketing mail (D-08) -- applied
  // here from day one rather than discovered live, unlike Ticketmaster CZ.
  'hello@feverup.com': feverTextHasPurchaseDetails,
  // info@smsticket.cz also sends newsletters/weekly-promo mail from the same
  // address (its own real confirmation's footer advertises them) -- applied
  // here from day one too (D-09, quick-261026-ma6).
  'info@smsticket.cz': smsticketTextHasPurchaseDetails,
  // info@mg.goout.net (quick-261029-pmx, D-10) -- REQUIRED: GoOut is a
  // bulk ESP (Mailgun) and markets through this same pipeline.
  'info@mg.goout.net': gooutHtmlHasConfirmedReservation,
};
// DELIBERATE ABSENCE: there is NO 'automat@webticket.cz' key here
// (D-09, quick-261028-h1n). This registry gates BODY-mode admission only;
// WebTicket.cz is ICS-sourced, and ICS-mode admission already requires a
// `.ics` attachment from this sender (TICKET_ICS_MODE_SENDERS +
// findTicketIcsAttachments) -- that structural requirement is ITSELF the
// gate. `automat@webticket.cz` is a transactional robot address
// (X-Mailer: WebTicket Mailer); the real sample shows no marketing
// callout, unlike smsticket.cz's own footer. Marketing mail does not carry
// a `.ics`, so it is simply never admitted at all -- no detector is needed
// to tell it apart from a real purchase confirmation.

// =============================================================================
// SHARED BODY-TEXT SOURCE (quick-261026-ma6)
// =============================================================================
//
// Built for smsticket.cz (D-02), whose real confirmation's text/plain MIME
// part is PRESENT but EMPTY — but applied GENERICALLY, to every body-sourced
// portal's two existing message.getPlainBody() read sites below, because it
// can only ever turn a guaranteed failure (an empty string handed to a
// parser/detector that expects real content) into an attempt. Every portal
// whose plain body is non-blank (all six pre-existing portals, as observed
// live) sees BYTE-IDENTICAL input to before this change.

/**
 * ticketingStripHtmlScriptStyleAndComments (D-03, quick-261029-pmx) — pure;
 * extracted BYTE-IDENTICALLY from ticketingHtmlToPlainText's own style,
 * script and comment-removal steps, so the schema.org microdata reader
 * (ticketingReadMicrodataItems, see this file's "SCHEMA.ORG MICRODATA
 * READER" section) can reuse the exact same stripping without
 * duplicating it. Replaces each `<style>...</style>` and
 * `<script>...</script>` element with a newline (case-insensitive,
 * non-greedy), then removes every HTML comment entirely. Never throws;
 * `null`/`undefined`/`''` all return `''`.
 */
function ticketingStripHtmlScriptStyleAndComments(html) {
  let text = String(html || '');

  text = text.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '\n');
  text = text.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '\n');
  text = text.replace(/<!--[\s\S]*?-->/g, '');

  return text;
}

/**
 * ticketingDecodeHtmlEntities (D-03, quick-261029-pmx) — pure; extracted
 * BYTE-IDENTICALLY from ticketingHtmlToPlainText's own entity-decoding
 * steps, so the schema.org microdata reader decodes attribute values with
 * the exact same rules. Decodes numeric character references —
 * hexadecimal (`&#xHHHH;`) THEN decimal (`&#DDDD;`) — via
 * `String.fromCodePoint`, leaving any reference above `0x10FFFF`
 * untouched rather than throwing, then the named entities `nbsp` (to
 * U+00A0), `lt`, `gt`, `quot` and `apos`, THEN `amp` LAST — so an
 * already-decoded `&lt;` et al. is never double-decoded (e.g. the
 * literal source text `&amp;lt;` decodes to the literal text `&lt;`,
 * not `<`). Never throws; `null`/`undefined`/`''` all return `''`.
 */
function ticketingDecodeHtmlEntities(text) {
  let decoded = String(text || '');

  decoded = decoded.replace(/&#x([0-9a-f]+);/gi, function (match, hex) {
    const codePoint = parseInt(hex, 16);
    return codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : match;
  });
  decoded = decoded.replace(/&#(\d+);/g, function (match, dec) {
    const codePoint = parseInt(dec, 10);
    return codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : match;
  });

  decoded = decoded.replace(/&nbsp;/gi, ' ');
  decoded = decoded.replace(/&lt;/gi, '<');
  decoded = decoded.replace(/&gt;/gi, '>');
  decoded = decoded.replace(/&quot;/gi, '"');
  decoded = decoded.replace(/&apos;/gi, "'");
  decoded = decoded.replace(/&amp;/gi, '&');

  return decoded;
}

/**
 * ticketingHtmlToPlainText (D-02; its style/script/comment-removal and
 * entity-decoding steps extracted into ticketingStripHtmlScriptStyleAndComments
 * and ticketingDecodeHtmlEntities, quick-261029-pmx, D-03) — pure;
 * deterministically renders an HTML email body down to plain text, for
 * the body-sourced parsing pipeline's fallback when a portal's own
 * message.getPlainBody() comes back blank. Never throws;
 * `null`/`undefined`/`''` all return `''`. Algorithm, in order (every
 * pattern below is non-greedy, bounded by a literal closing tag, or
 * built from a negated character class — never a nested quantifier, per
 * this file's threat model, T-261026-03):
 *   1. String-coerce `html`, with `null`/`undefined` becoming `''`.
 *   2. Replace the whole `<head>...</head>` element with a newline
 *      (case-insensitive, non-greedy) — this step stays INLINE here (the
 *      microdata reader deliberately does NOT strip `<head>`, since it
 *      never holds an itemscope).
 *   3. ticketingStripHtmlScriptStyleAndComments: replace each
 *      `<style>...</style>` and `<script>...</script>` element with a
 *      newline, then remove every HTML comment.
 *   4. Remove every `<img>` tag ENTIRELY — alt text is never rendered.
 *   5. Replace every `<br>` tag (any form) with a newline.
 *   6. Replace every opening/closing tag of p, div, tr, td, th, li, ul,
 *      ol, table, tbody, thead, tfoot and h1-h6 with a newline.
 *   7. Remove every remaining tag.
 *   8. ticketingDecodeHtmlEntities: decode numeric character references
 *      (hex then decimal) via `String.fromCodePoint`, then the named
 *      entities nbsp/lt/gt/quot/apos, THEN amp LAST.
 *   9. Split on CR, LF or CRLF; collapse each run of spaces, tabs and
 *      U+00A0 to ONE space; trim; drop lines left empty by that
 *      collapse; join the survivors with a single newline.
 */
function ticketingHtmlToPlainText(html) {
  let text = String(html || '');

  text = text.replace(/<head\b[^>]*>[\s\S]*?<\/head>/gi, '\n');
  text = ticketingStripHtmlScriptStyleAndComments(text);
  text = text.replace(/<img\b[^>]*>/gi, '');
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/?(?:p|div|tr|td|th|li|ul|ol|table|tbody|thead|tfoot|h[1-6])\b[^>]*>/gi, '\n');
  text = text.replace(/<[^>]+>/g, '');

  text = ticketingDecodeHtmlEntities(text);

  const lines = text
    .split(/\r\n|\r|\n/)
    .map(function (line) {
      return line.replace(/[ \t\u00A0]+/g, ' ').trim();
    })
    .filter(function (line) {
      return line !== '';
    });

  return lines.join('\n');
}

/**
 * ticketingMessageBodyText (D-02) — pure-over-its-argument (calls only
 * methods ON `message`, no GAS global). Returns `message.getPlainBody()`
 * UNCHANGED when that has any non-whitespace content. Otherwise, when
 * `typeof message.getBody === 'function'`, returns
 * `ticketingHtmlToPlainText(message.getBody())`; otherwise returns `''`.
 * The `typeof` guard keeps this never-throwing for a duck-typed fake (or a
 * hypothetical future GAS object) that lacks `getBody` entirely. Never
 * calls `message.getBody()` when the plain body already has content — the
 * fallback is BLANK-TRIGGERED ONLY.
 */
function ticketingMessageBodyText(message) {
  const plainBody = message.getPlainBody();
  if (plainBody && String(plainBody).trim() !== '') {
    return plainBody;
  }

  if (typeof message.getBody === 'function') {
    return ticketingHtmlToPlainText(message.getBody());
  }

  return '';
}

/**
 * TICKET_BODY_RAW_HTML_SENDERS (D-01, quick-261029-pmx) — the list of
 * BODY-SOURCED portals whose confirmation must be read from the RAW HTML
 * body (`message.getBody()`) rather than any plain-text rendering. Built
 * for GoOut (info@mg.goout.net): its event data lives entirely in
 * schema.org EventReservation microdata (itemprop/content/href
 * attributes), and EVERY plain-text rendering throws those attributes
 * away — both Gmail's own getPlainBody() of an HTML-only message (the
 * Fever situation: no text/plain part at all) and this file's own
 * ticketingHtmlToPlainText (D-02, quick-261026-ma6), which renders tag
 * TEXT content, never attribute values.
 *
 * This is a per-sender TECHNIQUE, not a fourth processing mode (see
 * ticketingBodyInputForSender below): admission, dedup, the PDF finder,
 * the attachment fetcher and event creation are all identical to
 * ordinary body mode. A mode would duplicate processTicketFromMessageBody
 * for no reason — only the parser's INPUT differs.
 *
 * Every non-listed sender's input stays BYTE-IDENTICAL to before this
 * list existed (ticketingMessageBodyText is untouched).
 *
 * INVARIANT (tested): every entry here also has a
 * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL entry.
 */
const TICKET_BODY_RAW_HTML_SENDERS = ['info@mg.goout.net'];

/**
 * ticketingMessageRawHtml (D-01, quick-261029-pmx) — pure-over-its-argument
 * (calls only methods ON `message`, no GAS global). Returns
 * `String(message.getBody() || '')` when `typeof message.getBody ===
 * 'function'`, else `''`. NEVER calls `message.getPlainBody()` — unlike
 * ticketingMessageBodyText, this is the raw-HTML counterpart for a sender
 * whose event data lives only in tag ATTRIBUTES, which every plain-text
 * rendering discards.
 */
function ticketingMessageRawHtml(message) {
  if (typeof message.getBody === 'function') {
    return String(message.getBody() || '');
  }
  return '';
}

/**
 * ticketingBodyInputForSender (D-01, quick-261029-pmx) — the single
 * switch point deciding WHICH body-text source a body-sourced portal's
 * parser and content detector read: `ticketingMessageRawHtml(message)`
 * for a sender in TICKET_BODY_RAW_HTML_SENDERS (GoOut), else
 * `ticketingMessageBodyText(message)` UNCHANGED for every other
 * body-sourced sender. `senderKey` is the already-lowercased
 * identifyingEmail (ticketingExtractEmailAddress's own output). Pure,
 * given a pure `message`.
 */
function ticketingBodyInputForSender(message, senderKey) {
  if (TICKET_BODY_RAW_HTML_SENDERS.indexOf(senderKey) !== -1) {
    return ticketingMessageRawHtml(message);
  }
  return ticketingMessageBodyText(message);
}

// =============================================================================
// SCHEMA.ORG MICRODATA READER (quick-261029-pmx)
// =============================================================================
//
// Built for GoOut (info@mg.goout.net, D-02): a THIRD extraction TECHNIQUE
// for this file, distinct from free-text literal-marker parsing (every
// other body parser below) and `.ics` parsing (parseIcs, reused via this
// file's end-of-file Node/GAS bridge). GoOut's confirmation carries FIVE
// `<div itemscope itemtype=".../EventReservation">` blocks, each with
// THREE `itemprop="name"` properties (the buyer, the event and the venue,
// in that order) — a flat `itemprop="name"` regex would grab the buyer's
// own name as the event name, which is exactly the trap this scope-aware
// reader avoids.
//
// WHY NOT a flat itemprop="name" regex: every block holds three `name`
// properties, buyer first.
// WHY NOT slicing per block and taking "the first name after
// reservationFor": that is only correct on the real sample because the
// Event's own name happens to precede its location div — microdata
// guarantees no property order.
// WHY NOT XmlService: this HTML is not well-formed XML (a void `<br>`;
// unclosed `<meta ...>` tags on addressRegion/postalCode/addressCountry
// in the real sample).
// WHY NOT a DOM library: Apps Script has no runtime npm dependencies.
//
// DOCUMENTED LIMITATION: an unescaped `>` inside a quoted attribute value
// ends that tag early; a properly escaped `&gt;` is fine.

// TICKETING_MICRODATA_TAG_PATTERN (D-02, quick-261029-pmx) — matches ANY
// HTML tag (opening or closing): group 1 the optional leading '/' for a
// closing tag, group 2 the tag name, group 3 the remaining (possibly
// empty) attribute text. Linear: a negated class `[^>]*` bounded by a
// literal `>` — never a nested quantifier (T-261029-03). `g`-flagged;
// every caller resets `lastIndex` to 0 first.
const TICKETING_MICRODATA_TAG_PATTERN = /<(\/?)([A-Za-z][A-Za-z0-9-]*)([^>]*)>/g;

// TICKETING_HTML_ATTRIBUTE_PATTERN (D-02, quick-261029-pmx) — matches one
// HTML attribute: a name (any run of characters excluding whitespace,
// quotes, angle brackets, `/` and `=`), optionally followed by `=` and a
// double-quoted, single-quoted or bare (unquoted) value. Each alternative
// is disjoint and bounded by its own closing delimiter or a negated
// class — never a nested quantifier (T-261029-03). `g`-flagged; every
// caller resets `lastIndex` to 0 first.
const TICKETING_HTML_ATTRIBUTE_PATTERN = /([^\s"'<>\/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;

// TICKETING_HTML_VOID_ELEMENTS (D-02, quick-261029-pmx) — HTML5 void
// element names (lowercase), each of which can never have a matching
// closing tag and so never gets pushed onto the reader's own element
// stack.
const TICKETING_HTML_VOID_ELEMENTS = [
  'area',
  'base',
  'br',
  'col',
  'embed',
  'hr',
  'img',
  'input',
  'link',
  'meta',
  'param',
  'source',
  'track',
  'wbr',
];

// TICKETING_MICRODATA_PROPERTY_NAME_PATTERN (D-02, quick-261029-pmx) — an
// itemprop token must be a plain identifier (a letter, then letters/
// digits only) to be assigned onto a scope object — this is what rejects
// `__proto__` and its kin (T-261029-08).
const TICKETING_MICRODATA_PROPERTY_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9]*$/;

/**
 * ticketingParseHtmlAttributes (D-02, quick-261029-pmx) — pure; parses an
 * HTML tag's raw attribute text (everything between the tag name and the
 * closing `>`, INCLUDING a trailing self-closing `/` if present — callers
 * decide what that trailing slash means) into an `Object.create(null)`
 * map of LOWERCASED attribute names to their (still entity-ENCODED)
 * string values. A value-less (boolean) attribute maps to `''`. The FIRST
 * occurrence of a given name wins — a later duplicate on the same tag is
 * ignored. Resets TICKETING_HTML_ATTRIBUTE_PATTERN's `lastIndex` to 0
 * first. Never throws.
 */
function ticketingParseHtmlAttributes(attributeText) {
  const attributes = Object.create(null);
  const text = String(attributeText || '');

  TICKETING_HTML_ATTRIBUTE_PATTERN.lastIndex = 0;
  let match = TICKETING_HTML_ATTRIBUTE_PATTERN.exec(text);
  while (match !== null) {
    const name = match[1].toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(attributes, name)) {
      const value = match[2] !== undefined ? match[2] : match[3] !== undefined ? match[3] : match[4] !== undefined ? match[4] : '';
      attributes[name] = value;
    }
    match = TICKETING_HTML_ATTRIBUTE_PATTERN.exec(text);
  }

  return attributes;
}

/**
 * ticketingAssignMicrodataProperty (D-02, quick-261029-pmx) — pure; step 7
 * of ticketingReadMicrodataItems's algorithm (see its own JSDoc below).
 * Splits `itempropAttr` on whitespace and, for each token matching
 * TICKETING_MICRODATA_PROPERTY_NAME_PATTERN that is NOT already an OWN
 * property of `scope` (FIRST WINS), sets `scope[token] = value`. A token
 * failing the pattern — including `__proto__`, which starts with `_` — is
 * silently skipped (T-261029-08).
 */
function ticketingAssignMicrodataProperty(scope, itempropAttr, value) {
  const tokens = String(itempropAttr || '').split(/\s+/);
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token && TICKETING_MICRODATA_PROPERTY_NAME_PATTERN.test(token) && !Object.prototype.hasOwnProperty.call(scope, token)) {
      scope[token] = value;
    }
  }
}

/**
 * ticketingReadMicrodataItems (D-02, quick-261029-pmx) — a GENERIC,
 * scope-aware schema.org MICRODATA reader, regex/string-based like every
 * other parser in this file (see this section's banner for the why-not
 * list: a flat regex, block-slicing, XmlService, a DOM library).
 *
 * ALGORITHM:
 *   1. String-coerce `html`, then run it through
 *      ticketingStripHtmlScriptStyleAndComments (D-03) — `<head>` is NOT
 *      stripped, since it never holds an itemscope.
 *   2. Scan tags via TICKETING_MICRODATA_TAG_PATTERN (`lastIndex` reset to
 *      0 first); `<!DOCTYPE>` and comments never match (comments are
 *      already gone by step 1).
 *   3. Parse each tag's attribute text via ticketingParseHtmlAttributes.
 *   4. Maintain an element stack of `{ tagName, item }`. A tag is VOID
 *      when its lowercased name is in TICKETING_HTML_VOID_ELEMENTS, or
 *      its attribute text ends in `/` (self-closing). A closing tag pops
 *      back to and including the NEAREST stack entry with the same
 *      tagName, or is ignored if none matches. The current SCOPE is the
 *      `item` of the topmost stack entry that has one, or `null`.
 *   5. An opening tag WITH an `itemscope` attribute creates
 *      `{ '@type': <the last '/'-segment of the FIRST whitespace-
 *      separated itemtype token, or ''> }`.
 *      - If it ALSO has an `itemprop` attribute and a current scope
 *        exists, the new item is assigned INTO that scope (step 7)
 *        instead of becoming a top-level result.
 *      - If it has NO `itemprop` and ANY of its itemtype tokens' last
 *        segment equals `itemTypeName`, it is collected as a TOP-LEVEL
 *        RESULT (regardless of its own DOM nesting depth).
 *      - It is pushed onto the stack (as a scope for its own
 *        descendants) unless it is void.
 *   6. An opening tag WITHOUT `itemscope`, but WITH `itemprop` and a
 *      current scope, takes its value from the FIRST present of
 *      `content`, `datetime`, `href`, `src`, decoded via
 *      ticketingDecodeHtmlEntities (D-03) — `''` when none is present (a
 *      text-CONTENT value is UNSUPPORTED, since it is unobserved). It is
 *      assigned (step 7) and pushed unless void.
 *   7. ASSIGN (ticketingAssignMicrodataProperty): split `itemprop` on
 *      whitespace; set each token matching
 *      TICKETING_MICRODATA_PROPERTY_NAME_PATTERN that is not already an
 *      OWN property of the scope object (FIRST WINS).
 *
 * RETURN VALUE: an array of plain objects (nested items are plain objects
 * too) in DOCUMENT ORDER. NEVER throws; `null`/`undefined`/`''` all give
 * `[]`.
 */
function ticketingReadMicrodataItems(html, itemTypeName) {
  const text = ticketingStripHtmlScriptStyleAndComments(String(html || ''));
  const results = [];
  const stack = [];

  TICKETING_MICRODATA_TAG_PATTERN.lastIndex = 0;
  let match = TICKETING_MICRODATA_TAG_PATTERN.exec(text);

  while (match !== null) {
    const isClosing = match[1] === '/';
    const tagName = match[2].toLowerCase();
    const attributeText = match[3] || '';

    if (isClosing) {
      let popIndex = -1;
      for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i].tagName === tagName) {
          popIndex = i;
          break;
        }
      }
      if (popIndex !== -1) {
        stack.length = popIndex;
      }
      match = TICKETING_MICRODATA_TAG_PATTERN.exec(text);
      continue;
    }

    const attributes = ticketingParseHtmlAttributes(attributeText);
    const isVoid = TICKETING_HTML_VOID_ELEMENTS.indexOf(tagName) !== -1 || /\/\s*$/.test(attributeText);

    let currentScope = null;
    for (let i = stack.length - 1; i >= 0; i--) {
      if (stack[i].item) {
        currentScope = stack[i].item;
        break;
      }
    }

    const hasItemscope = Object.prototype.hasOwnProperty.call(attributes, 'itemscope');
    const itempropValue = Object.prototype.hasOwnProperty.call(attributes, 'itemprop') ? attributes['itemprop'] : null;

    if (hasItemscope) {
      const itemtypeAttr = attributes['itemtype'] || '';
      const itemtypeTokens = itemtypeAttr.split(/\s+/).filter(function (token) {
        return token;
      });
      const firstTokenSegments = (itemtypeTokens[0] || '').split('/');
      const newItem = { '@type': firstTokenSegments[firstTokenSegments.length - 1] || '' };

      if (itempropValue && currentScope) {
        ticketingAssignMicrodataProperty(currentScope, itempropValue, newItem);
      } else if (!itempropValue) {
        const matchesRequestedType = itemtypeTokens.some(function (token) {
          const segments = token.split('/');
          return segments[segments.length - 1] === itemTypeName;
        });
        if (matchesRequestedType) {
          results.push(newItem);
        }
      }

      if (!isVoid) {
        stack.push({ tagName: tagName, item: newItem });
      }
    } else if (itempropValue && currentScope) {
      const rawValue =
        attributes['content'] !== undefined
          ? attributes['content']
          : attributes['datetime'] !== undefined
            ? attributes['datetime']
            : attributes['href'] !== undefined
              ? attributes['href']
              : attributes['src'] !== undefined
                ? attributes['src']
                : '';
      ticketingAssignMicrodataProperty(currentScope, itempropValue, ticketingDecodeHtmlEntities(rawValue));
      if (!isVoid) {
        stack.push({ tagName: tagName, item: null });
      }
    } else if (!isVoid) {
      stack.push({ tagName: tagName, item: null });
    }

    match = TICKETING_MICRODATA_TAG_PATTERN.exec(text);
  }

  return results;
}

/**
 * ticketBodyLooksProcessable — the shared body-content gate behind both
 * admission points. Returns true when `senderKey` has no registered detector
 * (fail-open, see the registry's JSDoc), otherwise the detector's verdict on
 * `ticketingBodyInputForSender(message, senderKey)` (quick-261026-ma6:
 * previously `message.getPlainBody()` directly; quick-261029-pmx, D-01:
 * now routes through ticketingBodyInputForSender, which reads the RAW
 * HTML for a sender in TICKET_BODY_RAW_HTML_SENDERS — every other sender
 * still sees byte-identical input via ticketingMessageBodyText). Pure, no
 * GAS globals.
 */
function ticketBodyLooksProcessable(message, senderKey) {
  const detector = TICKET_BODY_CONTENT_DETECTORS_BY_IDENTIFYING_EMAIL[senderKey];
  if (!detector) {
    return true;
  }

  return detector(ticketingBodyInputForSender(message, senderKey)) === true;
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
 *   - `{ mode: 'ics', message, portal }` — EXACTLY ONE JOB PER MATCHING
 *     MESSAGE, for a portal whose sender resolves against
 *     TICKET_ICS_MODE_SENDERS (ICS-sourced, Ticketportal.cz,
 *     quick-261005-orv, D-02) AND the message carries at least one
 *     qualifying `.ics` attachment (findTicketIcsAttachments) — the ticket
 *     data comes from the .ics itself, once per message, never per
 *     attachment.
 * A message whose sender does not resolve to any configured portal, or
 * whose portal resolves to none of the three registries, or (for a
 * PDF/ICS-sourced portal) carries no qualifying attachment, contributes NO
 * jobs.
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
    } else if (TICKET_ICS_MODE_SENDERS.indexOf(senderKey) !== -1 && findTicketIcsAttachments(message).length > 0) {
      jobs.push({ mode: 'ics', message: message, portal: portal });
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

// TICKET_UTC_OFFSET_PATTERN (D-05, quick-261029-pmx) — validates an
// explicit UTC offset suffix ('Z' or '+HH:MM'/'-HH:MM') before it is
// appended VERBATIM to a Calendar API dateTime string in
// buildTicketCalendarEventResource below. Anchored both ends so a
// malformed or unexpected value (e.g. '+1:00', a bare 'UTC', or an IANA
// zone name) is silently ignored rather than producing an invalid
// dateTime string.
const TICKET_UTC_OFFSET_PATTERN = /^(?:Z|[+-]\d{2}:\d{2})$/;

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
 *
 * EXPLICIT END (quick-261005-orv, D-05): when `parsedTicket.end` is truthy
 * (a `{ year, month, day, hour, minute }` components object — currently
 * only buildTicketportalCzParsedTicket ever sets this), it is used
 * VERBATIM as the end components instead of the
 * DEFAULT_EVENT_DURATION_MINUTES fallback below. No pre-existing parser
 * sets `parsedTicket.end`, so this is a no-op for every other portal.
 *
 * EXPLICIT UTC OFFSET (quick-261029-pmx, D-05): when `parsedTicket.utcOffset`
 * is a string matching TICKET_UTC_OFFSET_PATTERN (only GoOut's parser
 * ever sets this — its microdata startDate is itself offset-qualified),
 * that offset is appended VERBATIM to BOTH `start.dateTime` and
 * `end.dateTime`, alongside the still-present `timeZone` field (the
 * Calendar API's own events.insert sample sends both together). This
 * makes the created instant EXACT regardless of the resolved calendar's
 * own zone, at the cost of a documented limitation: an event spanning a
 * DST transition keeps the START's offset for its END too (accepted). A
 * malformed value (e.g. `'+1:00'`, an IANA zone name, or a non-string)
 * is silently ignored, producing today's unsuffixed strings — a no-op
 * for every portal that never sets this field.
 */
function buildTicketCalendarEventResource(parsedTicket, timeZone, attachments) {
  const startComponents = {
    year: parsedTicket.year,
    month: parsedTicket.month,
    day: parsedTicket.day,
    hour: parsedTicket.hour,
    minute: parsedTicket.minute,
  };
  const endComponents = parsedTicket.end
    ? {
        year: parsedTicket.end.year,
        month: parsedTicket.end.month,
        day: parsedTicket.end.day,
        hour: parsedTicket.end.hour,
        minute: parsedTicket.end.minute,
      }
    : addMinutesToWallClockComponents(startComponents, DEFAULT_EVENT_DURATION_MINUTES);

  const utcOffsetSuffix =
    typeof parsedTicket.utcOffset === 'string' && TICKET_UTC_OFFSET_PATTERN.test(parsedTicket.utcOffset) ? parsedTicket.utcOffset : '';

  const resource = {
    summary: parsedTicket.eventName,
    location: parsedTicket.location,
    start: { dateTime: formatWallClockComponentsIso(startComponents) + utcOffsetSuffix, timeZone: timeZone },
    end: { dateTime: formatWallClockComponentsIso(endComponents) + utcOffsetSuffix, timeZone: timeZone },
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
 * above: for a portal whose event data comes entirely from the email BODY,
 * never from a PDF at all. The Drive/OCR pipeline built for enigoo.cz is NOT
 * used here in any way. Flow:
 *   1. Parse the body text — `ticketingMessageBodyText(message)`
 *      (quick-261026-ma6, D-02): `message.getPlainBody()` when that has any
 *      content, otherwise a plain-text rendering of `message.getBody()` for
 *      a portal whose plain-text part is empty (smsticket.cz) — via the
 *      matched portal's body parser (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL).
 *      When the plain body was blank and the HTML rendering was used
 *      instead, one console.log line names the portal, so the Executions
 *      log shows which path ran live.
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
 *
 * BODY-PARSER CONTRACT (quick-260921-gj0): the registered body parser is
 * always called with a SECOND argument, `message.getDate()` — the message's
 * received date (D-07) — and, as of round 2, a THIRD argument,
 * `message.getSubject()` — the message's subject (D-25). Both exist for
 * portals (Fever) whose body carries no year of its own (date) or whose
 * subject is a cleaner event-name source than the body (subject); the three
 * pre-existing body parsers each take fewer parameters and simply ignore the
 * extra arguments, so their behavior is unchanged.
 *
 * RAW-HTML BODY SOURCE (quick-261029-pmx, D-01): `senderKey` is computed
 * ONCE at the top and reused by every lookup below. For a sender in
 * TICKET_BODY_RAW_HTML_SENDERS (GoOut), step 1 logs a dedicated line and
 * reads `ticketingBodyInputForSender(message, senderKey)` directly —
 * `message.getPlainBody()` is NEVER called for such a sender, since its
 * event data lives only in HTML tag attributes that every plain-text
 * rendering discards. Every other sender keeps the EXACT pre-existing
 * blank-plain-body-then-HTML-fallback behavior.
 *
 * GAS-only (GmailMessage/DriveApp/CalendarApp/Calendar globals) — not
 * unit-tested, proven only by the live checkpoint; the pure logic it
 * depends on (parseKinoArtTicketText and friends) IS fully unit-tested.
 */
function processTicketFromMessageBody(message, portal) {
  const calendarId = resolveTicketingCalendarId(portal, CONFIG.calendarId);
  const senderKey = ticketingExtractEmailAddress(portal.identifyingEmail);

  let bodyText;
  if (TICKET_BODY_RAW_HTML_SENDERS.indexOf(senderKey) !== -1) {
    console.log('Ticketing portal: reading the raw HTML body (schema.org microdata) for ' + portal.identifyingEmail + '.');
    bodyText = ticketingBodyInputForSender(message, senderKey);
  } else {
    const rawPlainBody = message.getPlainBody();
    const plainBodyIsBlank = !rawPlainBody || String(rawPlainBody).trim() === '';

    if (plainBodyIsBlank && typeof message.getBody === 'function') {
      console.log(
        'Ticketing portal: message.getPlainBody() was blank for ' + portal.identifyingEmail +
          "; parsing the HTML body's own text rendering instead."
      );
    }

    bodyText = ticketingBodyInputForSender(message, senderKey);
  }

  const parseTicketBody = TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey];
  if (!parseTicketBody) {
    throw new Error('No ticket-body parser registered for ticketing portal: ' + portal.identifyingEmail);
  }
  const parsedTicket = parseTicketBody(bodyText, message.getDate(), message.getSubject());

  if (isDuplicateTicketPurchase(parsedTicket.ticketIdentifier, calendarId)) {
    return;
  }

  // `attachments` is an ARRAY, since a portal may contribute more than one
  // attachment. The PDF-finder path below still contributes at most one
  // entry, with `mimeType: 'application/pdf'`.
  const attachments = [];

  if (portal.insertPdfIntoEvent) {
    const findPortalTicketPdfAttachment = TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL[senderKey];
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
  const fetchPortalAttachments = TICKET_BODY_MODE_ATTACHMENT_FETCHERS_BY_IDENTIFYING_EMAIL[senderKey];
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
 * processTicketFromIcsAttachment (quick-261005-orv, D-04, D-08; generalized
 * to a per-sender builder registry in quick-261028-h1n, D-02) — the
 * ICS-SOURCED processing mode, shared by BOTH ICS-sourced portals
 * (Ticketportal.cz and, since quick-261028-h1n, WebTicket.cz): `message`'s
 * own `.ics` attachment supplies the event data (via the EXISTING parseIcs,
 * reused through this file's end-of-file Node bridge / the shared GAS
 * global scope — D-01), never a portal-specific text/body parser. GAS-only
 * (CalendarApp/Calendar/DriveApp/Utilities globals) — not unit-tested
 * directly beyond this function's own GAS-shaped coordination; the pure
 * logic it depends on (the registered builders and friends) IS fully
 * unit-tested. Exported for the Node global-injection harness, same as
 * processTicketFromMessageBody.
 *
 * Flow:
 *   (a) icsAttachments = findTicketIcsAttachments(message). Throws if
 *       empty (should not happen when appliesTo/resolveTicketProcessingJobs
 *       already gated on this, but guards against a future call-order
 *       bug).
 *   (b) events = parseIcs(icsAttachments[0].getDataAsString()). Throws a
 *       clear Error naming the portal if zero VEVENTs parsed. SCOPE
 *       LIMITATION: the real sample carries exactly one `.ics` attachment
 *       and one VEVENT; if more than one of either is ever seen, only the
 *       first is used and a console.log notes it — generalizing to more
 *       is deferred until real data shows it is needed.
 *   (c) Resolve the target calendar (resolveTicketingCalendarId + the one
 *       CalendarApp.getCalendarById round-trip, same pattern as
 *       createTicketCalendarEvent) and its live timeZone. Throws if the
 *       resolved calendarId does not resolve to a real calendar.
 *   (d) Resolve the matched ticket PDF via the EXISTING
 *       TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL registry (D-08) —
 *       REGARDLESS of `portal.insertPdfIntoEvent`, since the dedup
 *       ticketIdentifier depends on the PDF's filename even when the PDF
 *       itself will not be attached.
 *   (e) Look up this portal's builder in
 *       TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL (D-02) by
 *       `senderKey`, throwing a clear Error naming the portal when none is
 *       registered — this should not happen when TICKET_ICS_MODE_SENDERS
 *       already gated admission, but guards against the two registries
 *       drifting apart. parsedTicket = builder(events[0], ticketPdf ?
 *       ticketPdf.getName() : null, timeZone, <Utilities.formatDate
 *       wrapper>, ticketingMessageBodyText(message)). A null
 *       ticketIdentifier (no matched PDF, and for WebTicket.cz also no body
 *       fallback) is logged — the dedup safety net simply cannot run for
 *       this message, same documented limitation as every other portal
 *       missing a stable identifier.
 *
 *       ICS BUILDER CONTRACT (quick-261028-h1n, D-08): every registered
 *       builder is called with this SAME fifth argument,
 *       `ticketingMessageBodyText(message)` — the shared body-text source
 *       (the plain body, or the HTML rendering fallback when that is
 *       blank). Ticketportal.cz's own builder
 *       (buildTicketportalCzParsedTicket) takes only FOUR parameters and
 *       simply ignores a fifth, so its behavior is UNCHANGED — the same
 *       "extra trailing argument is inert" pattern this file's
 *       BODY-PARSER CONTRACT already establishes for body-sourced parsers.
 *       WebTicket.cz's own builder (buildWebticketParsedTicket) reads this
 *       argument to enrich its description and supply its dedup fallback
 *       (see that function's own JSDoc).
 *   (f) DEDUP SAFETY NET (shared via isDuplicateTicketPurchase) — if a
 *       duplicate, silent no-op, no Drive/Calendar writes at all.
 *   (g) If `portal.insertPdfIntoEvent` AND a matched PDF exists: copy it
 *       into the PERMANENT CONFIG.ticketAttachmentDriveFolderName folder
 *       (via getOrCreateDriveFolderByName), rename it via the shared
 *       buildTicketAttachmentFilename convention, and push one
 *       `{ fileId, fileUrl, title, mimeType: 'application/pdf' }` entry —
 *       the exact same shape processTicketFromMessageBody's own PDF-finder
 *       path produces.
 *   (h) Build and insert the Calendar event via the SHARED
 *       createTicketCalendarEvent, passing the accumulated attachments
 *       ARRAY (same diagnostic log convention as the other two modes).
 */
function processTicketFromIcsAttachment(message, portal) {
  const calendarId = resolveTicketingCalendarId(portal, CONFIG.calendarId);
  const senderKey = ticketingExtractEmailAddress(portal.identifyingEmail);

  const icsAttachments = findTicketIcsAttachments(message);
  if (icsAttachments.length === 0) {
    throw new Error('Ticketing portal (ics mode): no .ics attachment found on message for portal: ' + portal.identifyingEmail);
  }
  if (icsAttachments.length > 1) {
    console.log('Ticketing portal (ics mode): more than one .ics attachment found; only the first is used.');
  }

  const events = parseIcs(icsAttachments[0].getDataAsString());
  if (events.length === 0) {
    throw new Error('Ticketing portal (ics mode): the .ics attachment carried zero VEVENTs, for portal: ' + portal.identifyingEmail);
  }
  if (events.length > 1) {
    console.log('Ticketing portal (ics mode): more than one VEVENT found in the .ics; only the first is used.');
  }

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('Ticketing portal (ics mode): Calendar not found for resolved calendarId: ' + calendarId);
  }
  const timeZone = calendar.getTimeZone();

  const findPortalTicketPdfAttachment = TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL[senderKey];
  const ticketPdf = findPortalTicketPdfAttachment ? findPortalTicketPdfAttachment(message) : null;

  const buildParsedTicket = TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL[senderKey];
  if (!buildParsedTicket) {
    throw new Error('No .ics parsed-ticket builder registered for ticketing portal: ' + portal.identifyingEmail);
  }

  const parsedTicket = buildParsedTicket(
    events[0],
    ticketPdf ? ticketPdf.getName() : null,
    timeZone,
    function (date, timeZoneForFormat, pattern) {
      return Utilities.formatDate(date, timeZoneForFormat, pattern);
    },
    ticketingMessageBodyText(message)
  );
  if (!parsedTicket.ticketIdentifier) {
    console.log('Ticketing portal (ics mode): no ticketIdentifier available (no matched ticket PDF); dedup protection is unavailable for this message.');
  }

  if (isDuplicateTicketPurchase(parsedTicket.ticketIdentifier, calendarId)) {
    return;
  }

  const attachments = [];
  if (portal.insertPdfIntoEvent && ticketPdf) {
    const permanentFolder = getOrCreateDriveFolderByName(CONFIG.ticketAttachmentDriveFolderName);
    const permanentPdfFile = permanentFolder.createFile(ticketPdf.copyBlob());
    permanentPdfFile.setName(buildTicketAttachmentFilename(parsedTicket.eventName, parsedTicket, parsedTicket.ticketIdentifier));
    attachments.push({
      fileId: permanentPdfFile.getId(),
      fileUrl: permanentPdfFile.getUrl(),
      title: permanentPdfFile.getName(),
      mimeType: 'application/pdf',
    });
  } else if (portal.insertPdfIntoEvent) {
    console.log(
      'Ticketing portal (ics mode): insertPdfIntoEvent is true but no matching ticket PDF attachment was found on the message.'
    );
  }

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
   * AND any of: (a) that portal resolves against
   * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL (PDF/OCR-sourced) AND the
   * message carries at least one qualifying PDF attachment, (b) that
   * portal resolves against TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL
   * (body-sourced) AND the message body passes that portal's registered
   * content detector, if it has one (ticketBodyLooksProcessable), or (c)
   * that portal resolves against TICKET_ICS_MODE_SENDERS (ICS-sourced,
   * Ticketportal.cz, quick-261005-orv, D-02) AND the message carries at
   * least one qualifying `.ics` attachment. Otherwise false.
   * dispatchActions only skips on a strict `=== false`, so a literal
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
      if (TICKET_ICS_MODE_SENDERS.indexOf(senderKey) !== -1 && findTicketIcsAttachments(messages[i]).length > 0) {
        return true;
      }
    }

    return false;
  },

  /**
   * run — builds the processing job list via the pure, TESTABLE
   * resolveTicketProcessingJobs (each job tagged with `mode: 'pdf'`,
   * `'body'` or `'ics'`), then runs the appropriate pipeline exactly once
   * per job: processTicketPdfAttachment for `'pdf'` jobs (the Drive/OCR
   * pipeline), processTicketFromMessageBody for `'body'` jobs (body text
   * only, no Drive/OCR), processTicketFromIcsAttachment for `'ics'` jobs
   * (the message's own .ics attachment, quick-261005-orv). Duplicate-event
   * protection is provided entirely by the DEDUP SAFETY NET shared by all
   * three pipelines (isDuplicateTicketPurchase), not by restricting which
   * attachments/messages get processed here. A message matching neither a
   * configured portal nor any processing mode's requirements contributes
   * no job and is skipped gracefully.
   */
  run: function (thread) {
    const messages = thread.getMessages();
    const jobs = resolveTicketProcessingJobs(messages, TICKETING_PORTALS_ACTION.config.ticketingPortals);

    jobs.forEach(function (job) {
      if (job.mode === 'pdf') {
        processTicketPdfAttachment(job.attachment, job.portal);
      } else if (job.mode === 'body') {
        processTicketFromMessageBody(job.message, job.portal);
      } else if (job.mode === 'ics') {
        processTicketFromIcsAttachment(job.message, job.portal);
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
    // Exported so the tests can assert the hall list's SHAPE directly (array,
    // non-empty, unique, no blank/padded entries) and drive every listed hall
    // through the parser. The scalar-vs-list shape IS what the
    // kino-art-venue-not-recognized bug was, so it is pinned, not trusted.
    KINO_ART_KNOWN_VENUES: KINO_ART_KNOWN_VENUES,
    resolveKinoArtKnownVenue: resolveKinoArtKnownVenue,
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
    // The fifth portal, Fever (quick-260921-gj0).
    parseFeverTicketText: parseFeverTicketText,
    feverTextHasPurchaseDetails: feverTextHasPurchaseDetails,
    findFeverTicketPdfAttachment: findFeverTicketPdfAttachment,
    feverResolveEventYear: feverResolveEventYear,
    // The sixth portal, Ticketportal.cz (quick-261005-orv).
    TICKET_ICS_MODE_SENDERS: TICKET_ICS_MODE_SENDERS,
    findTicketIcsAttachments: findTicketIcsAttachments,
    findTicketportalCzTicketPdfAttachment: findTicketportalCzTicketPdfAttachment,
    stripTicketportalCzSummaryDateSuffix: stripTicketportalCzSummaryDateSuffix,
    extractTicketportalCzOrderNumber: extractTicketportalCzOrderNumber,
    ticketInstantToWallClockComponents: ticketInstantToWallClockComponents,
    buildTicketportalCzParsedTicket: buildTicketportalCzParsedTicket,
    processTicketFromIcsAttachment: processTicketFromIcsAttachment,
    // The seventh portal, smsticket.cz (quick-261026-ma6).
    parseSmsticketTicketText: parseSmsticketTicketText,
    smsticketTextHasPurchaseDetails: smsticketTextHasPurchaseDetails,
    extractSmsticketTicketCodes: extractSmsticketTicketCodes,
    findSmsticketTicketPdfAttachment: findSmsticketTicketPdfAttachment,
    // The shared D-02 body-text source (also used by every pre-existing
    // body-sourced portal's two message.getPlainBody() read sites).
    ticketingHtmlToPlainText: ticketingHtmlToPlainText,
    ticketingMessageBodyText: ticketingMessageBodyText,
    // The eighth portal, WebTicket.cz (quick-261028-h1n).
    extractWebticketOrderNumberFromPdfName: extractWebticketOrderNumberFromPdfName,
    buildWebticketParsedTicket: buildWebticketParsedTicket,
    findWebticketTicketPdfAttachment: findWebticketTicketPdfAttachment,
    TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL: TICKET_ICS_PARSED_TICKET_BUILDERS_BY_IDENTIFYING_EMAIL,
    parseWebticketOrderText: parseWebticketOrderText,
    // The ninth portal, GoOut (quick-261029-pmx) -- BODY-SOURCED from
    // schema.org EventReservation microdata in its RAW HTML.
    parseGooutTicketHtml: parseGooutTicketHtml,
    gooutHtmlHasConfirmedReservation: gooutHtmlHasConfirmedReservation,
    ticketingReadMicrodataItems: ticketingReadMicrodataItems,
    ticketingDecodeHtmlEntities: ticketingDecodeHtmlEntities,
    ticketingStripHtmlScriptStyleAndComments: ticketingStripHtmlScriptStyleAndComments,
    TICKET_BODY_RAW_HTML_SENDERS: TICKET_BODY_RAW_HTML_SENDERS,
    ticketingBodyInputForSender: ticketingBodyInputForSender,
    ticketingMessageRawHtml: ticketingMessageRawHtml,
    findGooutTicketPdfAttachment: findGooutTicketPdfAttachment,
  };
}

// Node/GAS environment bridge for src/05-action-ics-import.js's parseIcs and
// isIcsAttachment (quick-261005-orv, D-01). Deliberately placed AFTER the
// module.exports block above, i.e. as the LAST statement in this file: 05
// requires THIS file back (its own end-of-file bridge, for
// resolveTicketingPortal/TICKETING_PORTALS_ACTION), so whichever of 05/07
// loads FIRST would otherwise see the OTHER file's module.exports still at
// its original, empty `{}` -- a mid-file placement here would hand 05 this
// file's own PARTIAL exports during that file's own require() of this one.
// Same end-of-file-after-module.exports technique both bridges use, which
// is what makes either load order safe (see 05's own matching bridge
// comment). Under GAS's shared global scope this block never runs
// (`typeof module` is always 'undefined' there); the bare `parseIcs`/
// `isIcsAttachment` names resolve through the concatenated global scope at
// call time instead.
if (typeof module !== 'undefined' && module.exports) {
  const icsModule = require('./05-action-ics-import.js');
  globalThis.parseIcs = icsModule.parseIcs;
  globalThis.isIcsAttachment = icsModule.isIcsAttachment;
}

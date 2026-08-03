/**
 * TICKETING_PORTALS_ACTION — v0.6.0 NEW ACTION (quick-260731-tix): detects
 * ticket-purchase confirmation emails (concerts, theater, cinema, events)
 * from configured ticketing portals and creates ONE calendar event per
 * email/purchase.
 *
 * FIRST-EVER Drive/PDF/OCR usage and FIRST-EVER Calendar event attachment
 * usage in this codebase. Unlike the ICS action (a structured .ics
 * attachment) and the booking.com action (a parseable plain-text email
 * body), a ticketing-portal confirmation email's BODY carries NO usable
 * event data at all — everything (event name, date/time, venue) is inside
 * an attached PDF ticket, which must be extracted via Google Drive's
 * PDF-to-Google-Docs OCR conversion, since Apps Script has no native PDF
 * text parser.
 *
 * ONE-EVENT-PER-PURCHASE (confirmed explicitly with the owner, matches how
 * the booking.com action already treats a multi-guest reservation as one
 * event): a purchase with multiple ticket pages for the SAME event (e.g. a
 * 2-ticket purchase) must yield exactly ONE calendar event, never one per
 * page/ticket. A portal's text parser (e.g. parseEnigooTicketText) only
 * ever reads the FIRST occurrence of its anchored fields within ONE PDF's
 * extracted text, and every ticket page in a single-purchase PDF repeats
 * the same event name/date/venue (only a trailing per-ticket number
 * differs) — this alone is sufficient when a purchase's tickets are pages
 * of ONE PDF file.
 *
 * REAL DOUBLE-BOOKING INCIDENT (live-test-driven, quick-260731-tix round
 * 8, CORRECTED in round 9 — see the CORRECTION note below before trusting
 * anything about the attachment-count claim in an earlier version of this
 * comment): the owner reported a real live bug — processing the enigoo.cz
 * email ONE time created TWO calendar events for the same purchase, and
 * manually reprocessing the SAME email added a THIRD, since there was NO
 * idempotency protection at all (unlike the ICS action's `iCalUID`-based
 * `Events.import`, or the booking.com action's
 * `extendedProperties.private.confirmationNumber` tag +
 * `findOrTagMatchingEvent` safety-net check).
 *
 * ROUND 8's ORIGINAL (WRONG) DIAGNOSIS: round 8 claimed the real enigoo.cz
 * email attaches its 2-ticket purchase as TWO SEPARATE PDF FILES, and
 * "fixed" this by having `run` process only the FIRST qualifying PDF
 * attachment per message (`selectPrimaryTicketPdfAttachment`). That claim
 * was WRONG. Round 9 independently re-parsed the REAL raw `.eml`'s actual
 * MIME structure (not a guess) and confirmed exactly ONE `application/pdf`
 * MIME part — one PDF, 2 internal pages, not two file attachments.
 * `selectPrimaryTicketPdfAttachment` has been REMOVED: it never actually
 * addressed the real single-run duplicate for this email (which only ever
 * had one attachment, so selecting "the first of one" changed nothing),
 * and it would have been a genuine correctness regression for a
 * legitimate FUTURE scenario — a portal that emails multiple DIFFERENT
 * purchases as separate PDF attachments in one message would have had
 * every attachment after the first silently dropped, never processed.
 * `resolveTicketProcessingJobs` now processes EVERY qualifying PDF
 * attachment again (restored to its original, pre-round-8 behavior — see
 * its own JSDoc for the corrected reasoning).
 *
 * THE ACTUAL ROOT CAUSE OF THE ORIGINAL SINGLE-RUN DUPLICATE REMAINS
 * UNCONFIRMED at the exact mechanism level (this session has no live
 * access to the owner's actual Gmail thread/message structure to trace it
 * further) — plausible candidates include multiple qualifying MESSAGES
 * within the same Gmail thread (`run` iterates `thread.getMessages()`,
 * and a thread carrying more than one email from the same portal sender,
 * each independently resolving to a portal and each independently
 * carrying its own single PDF, would produce one processing job PER
 * MESSAGE) or a duplicate message delivery. `isTicketPdfAttachment`/
 * `findTicketPdfAttachments` were independently re-verified via direct
 * code inspection (`Array.prototype.filter`'s OR-based predicate cannot
 * ever duplicate a single source array element into two output entries —
 * proven with a standalone script, not assumed) to rule out a
 * duplicate-detection bug at that layer.
 *
 * WHAT ACTUALLY PREVENTS THE DUPLICATE, REGARDLESS OF THE EXACT
 * MECHANISM: the DEDUP SAFETY NET (kept from round 8, unaffected by the
 * round-9 correction above, mirroring the booking.com action's own proven
 * pattern): `parseEnigooTicketText` extracts a stable `ticketIdentifier`
 * (the first per-ticket number found, e.g. `"24601"`); the created event
 * is tagged at creation time via
 * `extendedProperties.private.ticketIdentifier`
 * (`processTicketPdfAttachment`), and BEFORE creating a new event,
 * `findTicketEventByIdentifier` searches the resolved calendar for an
 * existing event already carrying that same tag — if found, the run is a
 * silent no-op (temp PDF deleted, no re-upload/re-attach, no second
 * event), logged the same "already exists, not a duplicate path" style as
 * the booking.com action's own `handleConfirmation`. This layer is correct
 * regardless of WHY a second processing attempt for the same purchase
 * occurs (multiple messages in a thread, a duplicate delivery, a manual
 * re-run, a future retry mechanism, or anything else) — it checks real
 * calendar state before every write, rather than trying to guess ahead of
 * time which message/attachment to trust as "the one true source." A
 * portal parser that cannot extract any stable identifier (documented
 * per-parser limitation, not a silent gap) simply does not get this
 * protection.
 *
 * ALL PORTAL-SPECIFIC PARSING LOGIC LIVES IN THIS ONE FILE (owner-directed,
 * explicitly overriding an initially-proposed one-file-per-portal design
 * that would have mirrored the booking.com action's language-pack
 * architecture): for enigoo.cz now, and any future portal later, every
 * parser is a clearly-separated, well-named function/section within THIS
 * file (e.g. parseEnigooTicketText) — never a separate `07-portal-*.js`
 * file, unless the owner says otherwise for a specific future portal. A
 * portal is matched to its parser via TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL
 * (below), keyed by the SAME identifyingEmail used to resolve which
 * TICKETING_PORTALS config entry (calendarId, insertPdfIntoEvent) applies.
 *
 * THE DRIVE/OCR PIPELINE (processTicketPdfAttachment, GAS-only —
 * DriveApp/Drive Advanced Service/DocumentApp/CalendarApp/Calendar globals
 * — not directly unit-tested, same category as every other GAS-only
 * function in this codebase), exactly as confirmed with the owner:
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
 *      uniformly whether the source PDF has a real text layer (like
 *      enigoo.cz's) or is a scanned image, which is exactly why this is
 *      the right general-purpose mechanism here rather than, say, a
 *      text-layer-only shortcut.
 *   4. Read the converted Doc's full text
 *      (`DocumentApp.openById(id).getBody().getText()`).
 *   5. Delete the converted Doc (a temp OCR artifact) — ALWAYS, regardless
 *      of anything else (wrapped in try/finally around the read, so a
 *      read failure still cleans up before its error propagates).
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
 *      portal's parsed ticket has no explicit end time (enigoo.cz never
 *      does), a sensible fixed DEFAULT_EVENT_DURATION_MINUTES (2 hours) is
 *      added to the start via addMinutesToWallClockComponents — documented
 *      as a per-portal-parser concern: a FUTURE portal whose PDF DOES
 *      include an end time should use it directly instead of this
 *      default, rather than this shared default becoming implicitly
 *      mandatory for every portal. If `insertPdfIntoEvent` was true, the
 *      Drive file is included as a real Calendar attachment: the resource's
 *      `attachments` array (`[{ fileId, fileUrl, title, mimeType }]` — see
 *      the live-test-driven correction to this exact shape below), AND
 *      `Calendar.Events.insert` is called with `{ supportsAttachments: true }`
 *      as its `optionalArgs` — this is a REAL, DOCUMENTED Calendar API v3
 *      requirement (the `events.insert` method's `supportsAttachments`
 *      query parameter, default `false`) for attachments to be accepted at
 *      all; omitting it would silently drop the `attachments` array rather
 *      than erroring, so it is never optional here when an attachment is
 *      present. CALENDAR API EVENTATTACHMENT SHAPE (live-test-driven fix,
 *      quick-260731-tix round 7): the original design under-specified this
 *      resource's shape as just `{ fileId, title }` — the owner hit a real
 *      live `GoogleJsonResponseException: ... Missing attachment URL.`
 *      `fileUrl` is a REQUIRED field on every `attachments[]` entry per the
 *      Calendar API v3's documented `EventAttachment` schema; `fileId` is
 *      actually READ-ONLY on that schema (the server derives/confirms it
 *      FROM `fileUrl` when the URL points to a Drive resource the caller
 *      can access) — providing `fileId` alone, without `fileUrl`, is not
 *      sufficient even though it names a real Drive file. Fixed by adding
 *      `fileUrl` (via `DriveApp`'s simple-service `File.getUrl()`, already
 *      in use elsewhere in this file for Drive operations) and `mimeType`
 *      (`'application/pdf'` — a real, optional EventAttachment field,
 *      included as a defensive completeness improvement since this action
 *      always attaches a PDF, not because it is itself required).
 *
 * TIMEZONE: enigoo.cz tickets carry no timezone indicator, same documented
 * limitation as the booking.com action — the parsed wall-clock digits are
 * treated as LOCAL time at the venue, resolved live via
 * `CalendarApp.getCalendarById(calendarId).getTimeZone()` (the portal's own
 * configured calendar's timezone), never a hardcoded assumption. Formatted
 * via formatWallClockComponentsIso (deliberately no trailing `Z`/offset,
 * paired with an explicit Calendar API `timeZone` field) — the same
 * "resolve wall-clock digits, pair with a live-derived timeZone field"
 * pattern already proven by the booking.com action's formatLocalWallClockIso,
 * re-implemented locally here per this codebase's established
 * one-file-per-action self-containment convention (no cross-file require
 * for a helper this small).
 *
 * OAUTH SCOPE (see src/appsscript.json): uses the BROADER
 * `https://www.googleapis.com/auth/drive` scope rather than the narrower
 * `drive.file`, because `CONFIG.ticketAttachmentDriveFolderName` is a
 * NAME-based lookup (`DriveApp.getFoldersByName`) deliberately designed to
 * potentially match a folder the OWNER pre-created by hand — not
 * necessarily one this script itself created. `drive.file`'s documented
 * restriction (visibility limited to files/folders the app itself created
 * or the user explicitly opened) does not reliably cover that case, so the
 * broader scope was chosen deliberately, not by default. Flagged
 * explicitly for the owner in this feature's live-verification checkpoint.
 *
 * GLOBALLY-UNIQUE NAMING WARNING (see the booking.com action's own
 * class-level JSDoc for the full incident this warning originates from):
 * Apps Script concatenates every project file into ONE shared global
 * scope — `ticketingExtractEmailAddress` (not a bare `extractEmailAddress`)
 * is this file's own locally-reimplemented sender-parsing helper, per the
 * one-file-per-action self-containment pattern, deliberately namespaced to
 * avoid colliding with the ICS and booking.com actions' own same-purpose
 * helpers.
 */

/**
 * ticketingExtractEmailAddress — LOCAL copy of the same underlying logic
 * already re-implemented independently in the ICS and booking.com action
 * files (see the booking.com action's own JSDoc for the real cross-file
 * naming-collision incident that established the "globally-unique name per
 * action" convention this follows). Extracts the bare, trimmed, lowercased
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
 * `identifyingEmail` case-insensitively matches `fromHeader`'s sender,
 * mirroring resolveIcsCalendarId's per-sender lookup convention
 * (src/05-action-ics-import.js): list order, FIRST match wins; no match
 * (or a null/empty `portals` list) returns `null`, never throws. Pure, no
 * GAS globals.
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
 * TICKETING_PORTALS entry). MULTI-CALENDAR ROUTING (live-test-driven fix,
 * quick-260731-tix round 6): a REAL missing-fallback bug — the shipped
 * enigoo.cz portal entry's default `calendarId` is `null` (the safe,
 * unconfigured default, same as the other two actions' optional
 * overrides), but this action originally read `portal.calendarId`
 * DIRECTLY everywhere, with no null-safety at all. Owner hit
 * `TypeError: Cannot read properties of null (reading 'getTimeZone')`
 * live, since `CalendarApp.getCalendarById(null)` returns `null`. Both the
 * ICS action (`resolveIcsCalendarId`) and the booking.com action
 * (`resolveBookingCalendarId`) already fall back to `CONFIG.calendarId`
 * (the global default) when their own action/portal-level override is
 * null — this action was simply missing that same established fallback
 * pattern. Simple two-tier resolution, mirroring
 * `resolveBookingCalendarId`'s exact shape (no per-sender map complexity
 * needed here, since `portal` is already the single resolved entry):
 * `portal.calendarId` when truthy, else `defaultCalendarId`. Pure, no GAS
 * globals.
 */
function resolveTicketingCalendarId(portal, defaultCalendarId) {
  return (portal && portal.calendarId) || defaultCalendarId;
}

/**
 * parseEnigooTicketText — the enigoo.cz-specific ticket-text parser (see
 * this file's class-level JSDoc for why every portal's parser lives here,
 * not in its own file).
 *
 * PATTERN-ANCHORED EXTRACTION, NOT LINE-POSITION (live-test-driven,
 * quick-260731-tix round 5 — the ACTUAL root-cause rewrite): a fixed-line-
 * position model (event name = line 1, date/time = line 2, location =
 * line 3) was this parser's ORIGINAL design and survived two earlier
 * live-test-driven fixes (a missing OAuth scope, then a paragraph-
 * separator splitting bug — see below) before the owner's THIRD live
 * failure, now armed with the round-3 full-extracted-text diagnostic,
 * finally supplied the real ground truth and disproved the model
 * entirely. Example modeled on the raw text `DocumentApp.getBody().getText()`
 * actually returned during that live failure (identifying details replaced
 * with fictional equivalents, 2 pages of the same purchase):
 *
 *   Letní hudební festival 15.08.2026 19:00
 *   Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount:
 *   24601
 *   Kulturní spolek z. s., ... DUZP 2.8.2026 Vstupenka je osvobozena...
 *   Letní hudební festival 15.08.2026 19:00
 *   Nádvoří kulturního domu Cena/price: 290 Kč Sleva/discount:
 *   24600
 *   Kulturní spolek z. s., ... DUZP 2.8.2026 Vstupenka je osvobozena...
 *
 * The event name and the date/time are packed onto the SAME real
 * paragraph ("Letní hudební festival 15.08.2026 19:00"); the location/price/
 * discount are packed onto a DIFFERENT single real paragraph. What LOOKS
 * like "each field on its own line" in the Google Docs editor's visual
 * rendering (text wrapping around the ticket's inline QR code image) does
 * NOT correspond to actual paragraph boundaries in what `Body.getText()`
 * returns — there is no reliable line position to key off of at all.
 * "Line-position-based parsing cannot work here" (confirmed live);
 * extraction is now ANCHORED to two literal patterns that are present
 * regardless of which paragraph/line they happen to share with other
 * fields:
 *
 *   1. DATE/TIME ANCHOR: the first `DD.MM.YYYY HH:MM`-shaped numeric
 *      match anywhere in the whole text. A multi-page purchase (like the
 *      example above) repeats this pattern once per page for the SAME
 *      event — only the FIRST occurrence is ever used, consistent with
 *      this action's established one-calendar-event-per-purchase design
 *      (see this file's class-level JSDoc).
 *   2. EVENT NAME: everything from the ABSOLUTE START of the text up to
 *      (not including) that first date/time match — the real text always
 *      begins with the event name immediately followed by the date/time,
 *      with nothing else before it. Anchored at `^` (no multiline flag),
 *      so this can only ever match the text's very first occurrence.
 *   3. LOCATION: everything strictly AFTER the date/time match's end, up
 *      to (not including) the literal label `Cena/price` — the price
 *      label always immediately follows the location on real tickets.
 *      Searched starting from the date/time match's OWN end index (not
 *      the whole text blindly), so a non-greedy capture naturally stops
 *      at the FIRST "Cena/price" reachable from there — which is always
 *      THIS ticket occurrence's own price line, never a later
 *      repetition's, even across a multi-page purchase.
 *   4. TICKET IDENTIFIER (added live-test-driven, quick-260731-tix round 8
 *      — see this file's class-level JSDoc "DEDUP" section for the full
 *      double-booking incident this resolves): the first run of digits
 *      found strictly AFTER the literal label `Sleva/discount:` (which
 *      always immediately follows the price on real tickets, right
 *      before the per-ticket number). This is the SAME per-page ticket
 *      number (e.g. `"24601"`) that differs between otherwise-identical
 *      repeated pages of a multi-ticket purchase — using the FIRST
 *      occurrence gives a real, stable, already-available identifier for
 *      the whole purchase without inventing a new field. Unlike the other
 *      three anchors, this one is OPTIONAL: if `Sleva/discount:` is not
 *      found, or no digit run follows it, `ticketIdentifier` is `null`
 *      rather than throwing — a missing dedup key should never block
 *      calendar-event creation, it only means the safety-net dedup check
 *      (see `findTicketEventByIdentifier`) cannot run for this ticket.
 *
 * None of the four anchors above ever splits the text into a line array —
 * `\s`/`[\s\S]` in each pattern already transparently absorbs whichever
 * paragraph-separator character(s) Google's OCR pipeline actually used
 * (this parser genuinely does not need to know or care), so this rewrite
 * is immune to the exact question the round-4 fix below was trying to
 * answer.
 *
 * Returns `{ eventName, location, year, month, day, hour, minute,
 * ticketIdentifier }` (month zero-indexed, matching Date.UTC's convention
 * and every other date-parsing function in this codebase; `ticketIdentifier`
 * is a string or `null`). Throws a controlled Error if no date/time pattern
 * is found anywhere, if the event name or the "Cena/price"-bounded location
 * cannot be extracted, or if the matched hour/minute are out of range — a
 * malformed/unexpected ticket layout should fail visibly, not silently
 * produce a wrong event. `ticketIdentifier` alone never causes a throw (see
 * point 4 above). Pure, no GAS globals.
 *
 * DIAGNOSTIC-ON-FAILURE (live-test-driven, quick-260731-tix round 3): every
 * throw below appends the COMPLETE raw `text` argument (the whole
 * DocumentApp-extracted OCR body text, untruncated) after the specific
 * problem description — never just the one line/segment that looked wrong
 * in isolation. This project's existing failure-notification path
 * (notifyOwnerOfFailure/composeFailureBody, src/02-main.js) already emails
 * the owner the full thrown error message on any action failure, so this
 * makes the NEXT failure notification email itself the diagnostic
 * artifact, with no extra manual step (e.g. a manual Drive "Open with
 * Google Docs" conversion) required from the owner — it was THIS exact
 * mechanism that supplied the real raw text motivating the round-5
 * rewrite above. This diagnostic is intentionally generic to ANY future
 * ticketing-portal parsing failure, not specific to any one incident.
 *
 * ROUND 4 HISTORY (superseded by round 5 above, but the underlying
 * insight remains true and is worth preserving): the owner's SECOND live
 * failure was initially attributed to a paragraph-separator splitting bug
 * — Apps Script's Document Service is long-documented to join paragraphs
 * with a CARRIAGE RETURN (`\r`), not a line feed (`\n`), and this file's
 * then-current `.split('\n')` call did not account for that. That fix
 * (matching `\r\n`, a bare `\r`, or a bare `\n`) was a REAL, VALID
 * discovery about Apps Script's actual behavior — it is NOT wrong, it was
 * simply INSUFFICIENT, since it was still built on the now-disproven
 * fixed-line-position extraction model underneath. The round-4 line-array
 * splitting code itself has been entirely removed by the round-5 rewrite
 * above (there is no line array left to split at all), but the knowledge
 * that Google's OCR pipeline may join fields with `\r` rather than `\n`
 * is exactly why every anchor above uses `\s`/`[\s\S]` (which absorb
 * `\r`, `\n`, and `\r\n` identically) rather than assuming one specific
 * separator character.
 */
function parseEnigooTicketText(text) {
  const rawText = String(text || '');

  // Anchor 1: the date/time pattern, searched across the WHOLE text,
  // FIRST occurrence only (regex .exec with no 'g' flag always returns
  // the first match) -- a multi-page/multi-ticket PDF repeats this
  // pattern once per page for the SAME purchase, and only the first
  // occurrence is ever used, per this action's one-event-per-purchase
  // design.
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
  // text (`^`, no multiline flag, so this only ever matches starting at
  // the very first character) up to (not including) the date/time
  // pattern. `[\s\S]+?` (not a bare `.+?`) so the non-greedy capture can
  // still traverse a real paragraph-separator character if one somehow
  // ends up between the event name and the date/time, without depending
  // on which separator convention Google's OCR pipeline used.
  const eventNameMatch = /^\s*([\s\S]+?)\s+\d{1,2}\.\d{1,2}\.\d{4}\s+\d{1,2}:\d{2}/.exec(rawText);
  if (!eventNameMatch) {
    throw new Error(
      'Unrecognized enigoo.cz ticket text: could not extract the event name preceding the date/time. Full extracted text:\n' + rawText
    );
  }
  const eventName = eventNameMatch[1].trim();

  // Anchor 3: the location -- everything strictly AFTER the date/time
  // match's own end index, up to (not including) the literal "Cena/price"
  // label. Slicing from that specific end index (not searching the whole
  // text blindly) is what keeps the non-greedy capture scoped to THIS
  // ticket occurrence's own location/price line, even when the same
  // labels repeat later in the text for a multi-page purchase.
  const afterDateTime = rawText.slice(dateTimeMatch.index + dateTimeMatch[0].length);
  const locationMatch = /([\s\S]+?)\s*Cena\/price/.exec(afterDateTime);
  if (!locationMatch) {
    throw new Error(
      'Unrecognized enigoo.cz ticket text: could not find "Cena/price" after the date/time to bound the location. Full extracted text:\n' +
        rawText
    );
  }
  const location = locationMatch[1].trim();

  // Anchor 4 (OPTIONAL -- never throws, see this function's own class-level
  // doc above): the ticket identifier -- the first run of digits found
  // strictly AFTER the literal label "Sleva/discount:", searched within
  // the SAME already-scoped `afterDateTime` substring used for the
  // location above (so this, too, stays confined to THIS ticket
  // occurrence, never a later repetition's). `null` when the label or a
  // following digit run is not found, rather than throwing -- a missing
  // dedup key must never block calendar-event creation.
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
// real data so far (a literal anchor, mirroring enigoo.cz's "Cena/price"
// label anchor above). SCOPE LIMITATION (deliberate, matches this
// project's established "don't guess at an unobserved variant" discipline
// — e.g. Kino Art's PDF-attachment filename matching below is scoped the
// same way): if Kino Art ever uses a different hall in a future
// confirmation email, this needs generalizing THEN, with real data, not
// guessed now.
const KINO_ART_KNOWN_VENUE = 'Cihlářská - Malý sál';

/**
 * parseKinoArtTicketText — the kinoart.cz-specific ticket-TEXT parser
 * (quick-260731-kar). UNLIKE parseEnigooTicketText above, this parses the
 * email BODY (`message.getPlainBody()`), never a PDF — see this file's
 * class-level JSDoc for the full BODY-SOURCED PROCESSING MODE this
 * introduces alongside the existing PDF/OCR-sourced mode. Example fixture
 * (HTML stripped to plain text, modeled on the real observed shape, a
 * single flattened line with a 2-seat/1-purchase repeating row):
 *
 *   Potvrzení objednávky Potvrzení objednávky č. 900142 Informace
 *   Zákazník: Jan Novák Úhrada: Comgate Doručení: Elektronicky
 *   Prodejní místa Kino Art Položky Název Místo Datum a čas Umístění Cena
 *   Tajný ostrov Cihlářská - Malý sál 7. 8. 2026 17:45 pátek 4 / 2
 *   150 Kč (1x Plná cena) Tajný ostrov Cihlářská - Malý sál
 *   7. 8. 2026 17:45 pátek 4 / 1 150 Kč (1x Plná cena) Cena celkem 300 Kč
 *   Tento email není vstupenka.
 *
 * Extraction anchors (pattern-anchored, same philosophy as
 * parseEnigooTicketText — never a line-position/line-array approach):
 *   1. DATE/TIME: `D. M. YYYY HH:MM` — day/month WITHOUT leading zeros,
 *      dot-SPACE separated (e.g. `7. 8. 2026 17:45`) — a GENUINELY
 *      DIFFERENT numeric date format from enigoo.cz's zero-padded
 *      no-space `15.08.2026`, hence its own distinct regex rather than
 *      reusing enigoo's pattern. First occurrence only (the row repeats
 *      once per seat in a multi-seat purchase).
 *   2. VENUE: the literal `KINO_ART_KNOWN_VENUE` string (see its own doc
 *      above) — always immediately precedes the date/time in the
 *      flattened body text.
 *   3. EVENT (movie) NAME: everything between the LAST occurrence of the
 *      literal column-header word `Cena` (capital C — this
 *      case-SENSITIVE match is what distinguishes it from the lowercase
 *      "cena" inside "Plná cena" appearing later in the same text) that
 *      occurs BEFORE the venue, and the venue string itself. The real
 *      body's table header row ("Položky Název Místo Datum a čas
 *      Umístění Cena") ends with this exact word immediately before the
 *      first data row begins.
 *   4. TICKET IDENTIFIER (OPTIONAL — never throws, same philosophy as
 *      parseEnigooTicketText's own ticketIdentifier anchor): the ORDER
 *      CONFIRMATION NUMBER, `Potvrzení objednávky č. <digits>`, found
 *      near the top of the body. Unlike enigoo.cz's per-TICKET number,
 *      this is naturally scoped to the WHOLE PURCHASE already (shared
 *      across every seat in a multi-seat purchase) — a BETTER dedup key,
 *      not merely an equivalent one.
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
  // BULLET-MARKER STRIP (live-test-driven, quick-260731-kar round 3): Gmail's
  // REAL message.getPlainBody() rendering of this email's HTML data rows
  // prefixes each row (the event name is the row's first field) with a
  // literal "* " bullet-list marker -- observed live (the owner's first
  // created event came out as "* Tajný ostrov"), NOT reproduced by the
  // hand-decoded fixture text this parser was originally built/tested
  // against (a naive HTML-tag-strip approximation, not Gmail's actual
  // list-rendering conversion) -- same category of surprise as the
  // enigoo.cz Body.getText() paragraph-separator lesson (round 4-5, see
  // parseEnigooTicketText's own class-level doc). SCOPE LIMITATION
  // (deliberate, same "don't guess at an unobserved variant" discipline as
  // KINO_ART_KNOWN_VENUE and findKinoArtTicketPdfAttachment above): only
  // the literal "* " marker actually observed is stripped -- a different
  // bullet character (e.g. "-", "•") would need handling THEN, with real
  // data, not speculatively now.
  const eventName = rawText
    .slice(cenaIndex + 'Cena'.length, knownVenueIndex)
    .trim()
    .replace(/^\*\s+/, '');
  if (!eventName) {
    throw new Error('Unrecognized Kino Art ticket text: extracted event name was empty. Full extracted text:\n' + rawText);
  }

  // OPTIONAL (never throws, see this function's own class-level doc
  // above): the order confirmation number, this portal's dedup
  // ticketIdentifier.
  //
  // ROW-BOUNDARY MARKER TOLERANCE (live-test-driven, quick-260731-kar
  // round 4): a real live insertPdfIntoEvent attachment rename came out as
  // "... - null.pdf" -- the ticketIdentifier failed to extract. Traced
  // precisely (not a threading/scoping bug -- both call sites already pass
  // parsedTicket.ticketIdentifier directly): the real raw .eml's HTML
  // (D:\download\KinoArtBrno.eml) puts the "Potvrzení objednávky" headline
  // and the "č. 900142" order number in TWO SEPARATE <tr> table rows -- the
  // exact same row-boundary structure round 3 already proved Gmail's real
  // getPlainBody() renders with an inserted "* " marker (see the eventName
  // bullet-marker strip above). The original regex required "Potvrzení
  // objednávky č." to be one contiguous, whitespace-only-separated phrase,
  // which such a marker would break. `[\s*]*` (whitespace AND/OR a literal
  // asterisk, any number of times) tolerates the SAME real, already-proven
  // noise class round 3 established -- not a new, unverified guess.
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

/**
 * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL — the local (single-file)
 * registry mapping a BODY-SOURCED ticketing portal's `identifyingEmail`
 * (lowercased, via ticketingExtractEmailAddress) to its email-body parser
 * function (quick-260731-kar) — the body-sourced counterpart to
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL below (PDF/OCR-sourced). Which
 * registry a matching portal's sender resolves against determines which
 * processing mode `resolveTicketProcessingJobs`/`run` route it through
 * (see this file's class-level JSDoc).
 */
const TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL = {
  'rezervace@kinoart.cz': parseKinoArtTicketText,
};

/**
 * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL — the local (single-file)
 * registry mapping a ticketing portal's `identifyingEmail` (lowercased,
 * via ticketingExtractEmailAddress) to its OCR-text parser function. This
 * is the mechanism that lets ALL portal parsers live in one file (see this
 * file's class-level JSDoc) while still cleanly routing a matched
 * TICKETING_PORTALS config entry to the RIGHT parser — adding a future
 * portal means adding one new parser function plus one new entry here,
 * nothing else.
 */
const TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL = {
  'no-reply@enigoo.cz': parseEnigooTicketText,
};

/**
 * DEFAULT_EVENT_DURATION_MINUTES — the sensible fixed default duration (2
 * hours) added to a parsed ticket's start time when a portal's PDF carries
 * no explicit end time (enigoo.cz never does). This is a per-portal-parser
 * CONCERN, not a mandatory rule: a FUTURE portal whose PDF DOES include an
 * end time should use it directly instead of this shared default — this
 * constant exists so every portal that genuinely needs a default has one
 * consistent, documented value to reach for, not so every future portal is
 * forced through it.
 */
const DEFAULT_EVENT_DURATION_MINUTES = 120;

/**
 * addMinutesToWallClockComponents — pure: adds `minutes` to a
 * `{ year, month, day, hour, minute }` wall-clock components object (month
 * zero-indexed), returning a NEW components object of the same shape,
 * correctly handling hour/day/month/year rollover. Implemented via
 * `Date.UTC` arithmetic purely as a NEUTRAL zero-offset calculation space
 * (never a real UTC instant — the input components carry no timezone
 * information at all, same "wall-clock digits, no real timezone" treatment
 * already established by the booking.com action's formatLocalWallClockIso/
 * buildInstantFromParsedDateLine): build a UTC-labeled millisecond
 * timestamp from the components, add the offset, then re-extract the
 * (rolled-over, if applicable) digits via the UTC getters — this exactly
 * mirrors real calendar/clock arithmetic (23:30 + 2h -> next day 01:30)
 * without ever needing to know or guess a real timezone. Pure, no GAS
 * globals.
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
 * Pure, no GAS globals. Tiny internal formatting helper for
 * formatWallClockComponentsIso, deliberately namespaced (not a bare
 * `zeroPad`) per this file's globally-unique-naming convention (the
 * booking.com action already has its own same-purpose `zeroPad`).
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
 * "TIMEZONE" doc paragraph) so the API interprets these digits as
 * wall-clock local time in that zone, not UTC. Re-implements the same
 * shape/intent as the booking.com action's formatLocalWallClockIso
 * (which formats FROM a Date, not from already-separated components,
 * since this file's parsers never build an intermediate Date object at
 * all) locally, per this codebase's established one-file-per-action
 * self-containment convention. Pure, no GAS globals.
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
 * characters (`/ \ ? % * : | " < >`) with `-` and trims whitespace.
 * Google Drive itself doesn't strictly forbid most of these in a
 * filename, but this keeps generated filenames unambiguous and safe to
 * browse/sort regardless. Pure, no GAS globals.
 */
function sanitizeTicketAttachmentFilenameComponent(value) {
  return String(value)
    .replace(/[/\\?%*:|"<>]/g, '-')
    .trim();
}

/**
 * buildTicketAttachmentFilename — the ATTACHMENT-RENAMING CONVENTION
 * (quick-260731-kar, applies to ALL portals, including a retrofit onto
 * the already-shipped enigoo.cz path — see this file's class-level JSDoc):
 * builds `"{event name} - {YYYY-MM-DD} - {ticket identifier}.pdf"` for
 * the ticket PDF moved into the permanent
 * `CONFIG.ticketAttachmentDriveFolderName` folder, e.g.
 * `"Letní hudební festival - 2026-08-15 - 24601.pdf"` (enigoo.cz) or
 * `"Tajný ostrov - 2026-08-07 - 900142.pdf"` (Kino Art). Uses
 * ISO-style `YYYY-MM-DD` (not a locale-specific date format) so files
 * sort consistently when browsing the Drive folder. Since a Calendar
 * event attachment's displayed `title` is already derived from the
 * file's name AT ATTACH TIME (`uploadedPdfFile.getName()`), renaming the
 * file before it is referenced automatically improves BOTH the Drive
 * folder's browsability AND what shows up on the calendar event — no
 * separate title-setting logic needed, just rename the actual Drive file
 * before its name/URL are read. `dateComponents` is any `{ year, month,
 * day }`-shaped object (month zero-indexed, matching every other
 * date-components object in this file) — a full parsed-ticket object
 * (which also carries `hour`/`minute`) works fine as-is, those extra
 * fields are simply ignored. Pure, no GAS globals.
 *
 * DEFENSIVE NULL-HANDLING (live-test-driven, quick-260731-kar round 4): a
 * real live attachment rename came out as "... - null.pdf" -- naive string
 * concatenation coerced a JS `null` ticketIdentifier argument into the
 * literal 4-character text "null". This is a defensive SECONDARY measure
 * (the primary fix, for THIS incident, was correcting parseKinoArtTicketText's
 * own extraction -- see its own class-level doc) so that any portal parser
 * that genuinely cannot extract a ticketIdentifier (a documented,
 * never-throwing possibility for every portal parser in this file) never
 * embeds the literal word "null" -- the segment is omitted entirely
 * instead.
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

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers this file's config getter depends on indirectly via
// TICKETING_PORTALS_ACTION.config, and for TICKETING_PORTALS_ACTION_CONFIG
// itself (defined in the sibling src/07-action-cfg-ticketing-portals.js —
// see the 260724-lqi config-split refactor for the full load-order/getter
// rationale this mirrors). Under GAS's shared global scope this is ALREADY
// visible here by bare name — no action needed, and this `if` block never
// executes there. Under Node, each `require()`d file is its own isolated
// module with its own scope, so the bare `TICKETING_PORTALS_ACTION_CONFIG`
// reference inside TICKETING_PORTALS_ACTION's `config` getter below would
// otherwise throw ReferenceError. Same `globalThis` bridge technique
// already established by the ICS and booking.com action files' own
// equivalent bridges.
if (typeof module !== 'undefined' && module.exports) {
  globalThis.TICKETING_PORTALS_ACTION_CONFIG = require('./07-action-cfg-ticketing-portals.js').TICKETING_PORTALS_ACTION_CONFIG;
}

/**
 * getOrCreateDriveFolderByName — finds a Drive folder by NAME (not ID) via
 * `DriveApp.getFoldersByName`, returning the FIRST match if one or more
 * exist, or creating a new folder via `DriveApp.createFolder` if none
 * exist yet. Used for BOTH the fixed-name auto-managed TEMP folder and the
 * owner-configured permanent `CONFIG.ticketAttachmentDriveFolderName`
 * folder — same lookup pattern for both, per the owner's own explicit
 * design. GAS-only (DriveApp) — not unit-tested, proven only by the live
 * checkpoint.
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
 * place, mirroring the ICS action's own isIcsAttachment.
 */
function isTicketPdfAttachment(attachment) {
  const name = (attachment.getName() || '').toLowerCase();
  const contentType = attachment.getContentType() || '';

  return name.slice(-4) === '.pdf' || contentType === 'application/pdf';
}

/**
 * findTicketPdfAttachments — returns every GmailAttachment on `message`
 * whose name ends in .pdf (case-insensitive) or whose content-type is
 * application/pdf (via isTicketPdfAttachment), in source order, or [] if
 * none match.
 */
function findTicketPdfAttachments(message) {
  return message.getAttachments().filter(isTicketPdfAttachment);
}

/**
 * findKinoArtTicketPdfAttachment — Kino Art sends TWO PDF attachments per
 * confirmation email (quick-260731-kar): `Vstupenky.pdf` (the real
 * ticket, one page per seat) and `Doklad.pdf` (a separate receipt/
 * invoice, NOT ticket data — the email body itself says "Tento email
 * není vstupenka", i.e. "this email is not the ticket", confirming
 * `Vstupenky.pdf` is the authoritative ticket). Returns the FIRST
 * qualifying PDF attachment (via findTicketPdfAttachments) whose name
 * contains `"Vstupenky"`, or `null` if none match — this deliberately
 * EXCLUDES `Doklad.pdf` and anything else. SCOPE LIMITATION (deliberate,
 * same "don't guess at an unobserved variant" discipline as
 * KINO_ART_KNOWN_VENUE above): scoped narrowly to the Czech filename
 * actually observed; a hypothetical future English-language Kino Art
 * confirmation with a differently-named attachment would need handling
 * THEN, with real data, not guessed now. Pure, no GAS globals (operates
 * only on the array already produced by findTicketPdfAttachments).
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
 * TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL — the local
 * (single-file) registry mapping a BODY-SOURCED ticketing portal's
 * `identifyingEmail` to its own ticket-PDF-finder function
 * (quick-260731-kar) — mirrors TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL's
 * own per-portal keying convention exactly, kept as a SEPARATE registry
 * (rather than hardcoding Kino Art's own finder directly inside
 * processTicketFromMessageBody) so a future body-sourced portal can
 * register its own finder the same way a future PDF-sourced portal
 * registers its own text parser.
 */
const TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL = {
  'rezervace@kinoart.cz': findKinoArtTicketPdfAttachment,
};

/**
 * resolveTicketProcessingJobs — the pure, TESTABLE extraction of `run`'s
 * per-message orchestration decision. Given `messages` (an array of
 * message-like objects exposing `getFrom()`/`getAttachments()` — GAS
 * `GmailMessage` objects in production, plain duck-typed fakes in tests)
 * and `portals` (the TICKETING_PORTALS config array), returns an array of
 * processing jobs, EACH TAGGED WITH A `mode` (quick-260731-kar — see this
 * file's class-level JSDoc for the full TWO-PROCESSING-MODE
 * architecture):
 *   - `{ mode: 'pdf', attachment, portal }` — ONE JOB PER QUALIFYING PDF
 *     ATTACHMENT, for a portal whose sender resolves against
 *     TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL (PDF/OCR-sourced, e.g.
 *     enigoo.cz).
 *   - `{ mode: 'body', message, portal }` — EXACTLY ONE JOB PER MATCHING
 *     MESSAGE, for a portal whose sender resolves against
 *     TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL (body-sourced, e.g. Kino
 *     Art) — the ticket data comes from the body itself, once per
 *     message, never per attachment (a body-sourced portal has no
 *     "per-attachment" concept for its actual event data at all).
 * A message whose sender does not resolve to any configured portal, or
 * whose portal resolves to neither registry, or (for a PDF-sourced
 * portal) carries no qualifying PDF attachment, contributes NO jobs.
 * Pure, no GAS globals — every GAS-shaped method call here is invoked ON
 * THE PASSED-IN objects only, never a real global service, so this is
 * fully unit-testable under Node with fake message/attachment objects.
 *
 * CORRECTED (live-test-driven, quick-260731-tix round 9 — see this file's
 * class-level "ONE-EVENT-PER-PURCHASE" doc for the full corrected
 * incident writeup): round 8 introduced `selectPrimaryTicketPdfAttachment`
 * (processing only the FIRST qualifying PDF attachment per message), on
 * the claimed basis that the real enigoo.cz email attaches its 2-ticket
 * purchase as TWO SEPARATE PDF files. That claim was WRONG — independently
 * re-parsing the real raw `.eml`'s actual MIME structure confirmed exactly
 * ONE `application/pdf` part (a single PDF with 2 internal pages, not two
 * separate file attachments). `selectPrimaryTicketPdfAttachment` has been
 * REMOVED: restricting to "only the first attachment" was never the right
 * fix for this real email (which only ever had one attachment to begin
 * with, so it changed nothing here) and would have been a REAL
 * correctness regression for a legitimate future scenario — a portal
 * that genuinely emails multiple DIFFERENT purchases as separate PDF
 * attachments in ONE message would have silently had every attachment
 * after the first dropped entirely, never processed at all. The `mode:
 * 'pdf'` branch below processes EVERY qualifying attachment again (its
 * original, pre-round-8 behavior). The one guarantee this codebase
 * actually needs — "the SAME purchase never gets a second calendar
 * event" — is provided entirely by the DEDUP SAFETY NET (ticketIdentifier
 * tag + findTicketEventByIdentifier, shared by BOTH processing modes via
 * isDuplicateTicketPurchase), which checks real calendar state before
 * every write and is correct regardless of how many messages/attachments/
 * re-runs ever trigger processing — unlike trying to guess which
 * attachment(s) to trust as "primary" ahead of time.
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
    } else if (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey]) {
      jobs.push({ mode: 'body', message: message, portal: portal });
    }
  }

  return jobs;
}

/**
 * findTicketEventByIdentifier — the DEDUP SAFETY NET's lookup (live-test-
 * driven, quick-260731-tix round 8 — see this file's class-level
 * "ONE-EVENT-PER-PURCHASE" doc for the full double-booking incident this
 * resolves): searches `calendarId` for an existing event already tagged
 * with `extendedProperties.private.ticketIdentifier` equal to
 * `ticketIdentifier`, mirroring the booking.com action's own
 * findEventByConfirmationTag EXACT-match query shape
 * (src/06-action-booking-com-management.js) — self-contained per this
 * codebase's one-file-per-action convention (no cross-file reuse):
 * `Calendar.Events.list(calendarId, { privateExtendedProperty:
 * 'ticketIdentifier=' + ticketIdentifier, singleEvents: true })`.
 * Deliberately NOT paginated (unlike booking.com's own
 * listCalendarEventsPaginated) and NOT time-windowed (unlike booking.com's
 * date-range bounding): a `privateExtendedProperty` filter against a
 * near-certainly-unique per-purchase ticket number is already an EXACT
 * match expected to return 0 or 1 events, so neither booking.com's
 * pagination helper nor its fuzzy hotel-name+date-overlap fallback layer
 * is needed for this portal's simpler use case. Returns the first matching
 * event, or `null` if none found. GAS-only (Calendar global) — not
 * unit-tested, proven only by the live checkpoint, same category as every
 * other GAS-only function in this file.
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
 * isDuplicateTicketPurchase — the DEDUP SAFETY NET's SHARED decision
 * (factored out, quick-260731-kar, so BOTH processing modes — PDF-sourced
 * `processTicketPdfAttachment` and body-sourced
 * `processTicketFromMessageBody` — use the EXACT same check, never
 * duplicated logic, per the coordinator's explicit "don't duplicate that
 * logic, factor it so both paths share it" instruction). Returns `true`
 * (and logs the same "already exists, not a duplicate path" message this
 * codebase has used since round 8) when `ticketIdentifier` is truthy AND
 * an event already carries that exact tag on `calendarId`. Returns
 * `false` when `ticketIdentifier` is falsy (a portal parser that could
 * not extract one simply does not get this protection — a documented
 * per-parser limitation, not a silent gap) or no matching event is
 * found. GAS-only (calls findTicketEventByIdentifier, which touches the
 * Calendar global) — not unit-tested, proven only by the live checkpoint.
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
 * createTicketCalendarEvent — the SHARED Calendar event resource
 * build+insert step (factored out, quick-260731-kar, so BOTH processing
 * modes share identical event-shape/tagging/attachment logic, never
 * duplicated): builds the event resource from `parsedTicket` (`{
 * eventName, location, year, month, day, hour, minute, ticketIdentifier
 * }`), tags it with `extendedProperties.private.ticketIdentifier` when
 * available (same "tag at creation" idea as the booking.com action's
 * `confirmationNumber`), and conditionally attaches `attachmentInfo`
 * (`{ fileId, fileUrl, title }`, or `null` for no attachment) using the
 * exact `EventAttachment` shape + the REQUIRED `supportsAttachments: true`
 * insert option (see this file's class-level "round 7" doc for the real
 * live bug that established this exact shape). TIMEZONE derived live from
 * the RESOLVED target calendar, never a hardcoded assumption, same
 * principle as the booking.com action. GAS-only (CalendarApp/Calendar
 * globals) — not unit-tested, proven only by the live checkpoint.
 */
function createTicketCalendarEvent(parsedTicket, calendarId, attachmentInfo) {
  const timeZone = CalendarApp.getCalendarById(calendarId).getTimeZone();

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

  if (parsedTicket.ticketIdentifier) {
    resource.extendedProperties = { private: { ticketIdentifier: parsedTicket.ticketIdentifier } };
  }

  const insertOptionalArgs = {};
  if (attachmentInfo) {
    // A real, documented Calendar API v3 EventAttachment resource shape
    // (live-test-driven fix, quick-260731-tix round 7): `fileUrl` is a
    // REQUIRED field on every attachments[] entry -- see this file's
    // class-level "round 7" doc for the full real-live-bug writeup this
    // shape resolves (fileId alone, without fileUrl, is NOT sufficient).
    resource.attachments = [
      { fileId: attachmentInfo.fileId, fileUrl: attachmentInfo.fileUrl, title: attachmentInfo.title, mimeType: 'application/pdf' },
    ];
    // A real, documented Calendar API v3 requirement (events.insert's own
    // `supportsAttachments` query parameter, default false) for the
    // `attachments` array above to be accepted at all — never optional
    // here when an attachment is actually present.
    insertOptionalArgs.supportsAttachments = true;
  }

  Calendar.Events.insert(resource, calendarId, insertOptionalArgs);
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
 * STEP ORDER CHANGED (live-test-driven, quick-260731-tix round 8): parsing
 * (originally step 7) now happens IMMEDIATELY after the OCR read, BEFORE
 * the move-to-permanent-folder-or-delete decision (originally step 6) —
 * necessary because the DEDUP SAFETY NET below needs `parsedTicket.
 * ticketIdentifier` to decide whether this run should even move/attach the
 * PDF at all. If a duplicate is found, the temp PDF is simply deleted
 * (never moved to the permanent folder, never re-attached) and the
 * function returns before any Calendar write — "do NOT re-upload/
 * re-attach the PDF again" for an already-existing purchase.
 *
 * ORPHANED TEMP-PDF FIX (live-test-driven, quick-260731-tix round 10): the
 * owner reported ticket PDFs staying in the TEMP Drive folder forever. Root
 * cause: `uploadedPdfFile`'s fate (moved to the permanent folder, or
 * trashed via the dedup-skip / insertPdfIntoEvent-false paths) was only
 * ever decided on the explicit success paths below — unlike the
 * OCR-converted Doc (which already had a `try/finally` guaranteeing
 * cleanup regardless of outcome), nothing guaranteed `uploadedPdfFile`
 * itself got cleaned up if `parseTicketText` threw (a real, frequently-hit
 * failure mode during this feature's live testing), or if anything
 * downstream (calendar resolution, the Calendar API call) threw. Fixed
 * with a `pdfFateResolved` flag set `true` at each of the three points
 * where the PDF's fate is explicitly decided, and a `finally` block
 * wrapping everything from upload through the Calendar API call that
 * trashes the file as a FALLBACK safety net only when that flag is still
 * `false` — never touching a file whose fate (moved to the permanent
 * folder as a live Calendar attachment, or already trashed) was already
 * correctly decided. This does NOT retroactively clean up PDFs already
 * orphaned by EARLIER failed live-test runs before this fix landed — see
 * the checkpoint notes for that manual cleanup step.
 */
function processTicketPdfAttachment(attachment, portal) {
  // Resolved ONCE at the top and threaded explicitly through every
  // downstream Calendar API call site below (mirroring the booking.com
  // action's own established `calendarId` resolution convention) rather
  // than each call site independently re-reading `portal.calendarId` —
  // see resolveTicketingCalendarId's own JSDoc for the real live bug this
  // fixes.
  const calendarId = resolveTicketingCalendarId(portal, CONFIG.calendarId);

  const tempFolder = getOrCreateDriveFolderByName(TICKETING_TEMP_DRIVE_FOLDER_NAME);
  const uploadedPdfFile = tempFolder.createFile(attachment.copyBlob());

  // pdfFateResolved -- set true at each of the three points below where
  // uploadedPdfFile's fate is explicitly decided (dedup-skip trash,
  // insertPdfIntoEvent move, insertPdfIntoEvent-false trash). The
  // surrounding try/finally's fallback cleanup only ever acts when this is
  // still false, i.e. something threw before any of those points were
  // reached -- see this function's own class-level "ORPHANED TEMP-PDF FIX"
  // doc above.
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

    // DEDUP SAFETY NET (defense in depth, live-test-driven, quick-260731-tix
    // round 8, kept unchanged by the round-9 correction, and factored out
    // in quick-260731-kar so BOTH processing modes share the exact same
    // check -- mirrors the booking.com action's own proven
    // findOrTagMatchingEvent/handleConfirmation pattern): if this ticket's
    // parser could extract a stable ticketIdentifier, check whether an
    // event already carries that exact tag before doing anything else --
    // this is the layer that actually prevents the duplicate, regardless
    // of WHY a second processing attempt occurs (see this file's
    // class-level "REAL DOUBLE-BOOKING INCIDENT" doc for the corrected
    // root-cause writeup).
    if (isDuplicateTicketPurchase(parsedTicket.ticketIdentifier, calendarId)) {
      uploadedPdfFile.setTrashed(true);
      pdfFateResolved = true;
      return;
    }

    // Step 6: move the original PDF to the permanent folder (keeping its
    // file ID for the Calendar attachment below) when insertPdfIntoEvent is
    // true, or delete it from the temp folder entirely when false — nothing
    // is left in EITHER Drive folder when this toggle is off.
    let attachmentInfo = null;
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
      // already moved (and now renamed) above (not a fresh
      // DriveApp.getFileById lookup) -- moveTo()/setName() only change the
      // file's parent folder/name, the object reference itself remains
      // valid, so a second Drive round-trip to re-fetch a file we already
      // have a live handle to is unnecessary.
      attachmentInfo = {
        fileId: uploadedPdfFile.getId(),
        fileUrl: uploadedPdfFile.getUrl(),
        title: uploadedPdfFile.getName(),
      };
    } else {
      uploadedPdfFile.setTrashed(true);
      pdfFateResolved = true;
    }

    // Step 8: build and insert the Calendar event via the SHARED helper
    // (factored out, quick-260731-kar, so both processing modes build the
    // identical event resource/tagging/attachment shape).
    //
    // DIAGNOSTIC VISIBILITY (quick-260731-kar round 2 -- see
    // processTicketFromMessageBody's own matching log for the full
    // rationale): symmetric log line so a "no event created, no error"
    // outcome is distinguishable in the Executions log the same way for
    // BOTH processing modes.
    console.log(
      'Ticketing portal: creating calendar event for "' + parsedTicket.eventName + '" (ticketIdentifier=' +
        parsedTicket.ticketIdentifier + ') on calendar ' + calendarId + '.'
    );
    createTicketCalendarEvent(parsedTicket, calendarId, attachmentInfo);
  } finally {
    if (!pdfFateResolved) {
      // Best-effort fallback cleanup -- do not let a cleanup failure mask
      // whatever original exception is already propagating (parsing
      // errors, missing OAuth scopes, calendar resolution failures, the
      // Calendar API call itself, etc. -- exactly the kind of real
      // failures this feature hit repeatedly during live testing, each of
      // which previously left an orphaned PDF behind in the temp folder
      // forever).
      try {
        uploadedPdfFile.setTrashed(true);
      } catch (cleanupError) {
        console.log('Ticketing portal: failed to clean up orphaned temp PDF ' + uploadedPdfFile.getId() + ': ' + cleanupError);
      }
    }
  }
}

/**
 * processTicketFromMessageBody — the BODY-SOURCED processing mode
 * (quick-260731-kar, alongside the existing PDF/OCR-sourced
 * processTicketPdfAttachment above): for a portal whose event data comes
 * entirely from the email BODY (`message.getPlainBody()` — the SAME
 * method the booking.com action already uses), never from a PDF at all.
 * The Drive/OCR pipeline built for enigoo.cz is NOT used here in any way
 * — see this file's class-level JSDoc for the full two-processing-mode
 * architecture. Flow:
 *   1. Parse the plain body text via the matched portal's body parser
 *      (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL).
 *   2. DEDUP SAFETY NET (shared with the PDF-sourced mode via
 *      isDuplicateTicketPurchase) — if a duplicate, silent no-op, no
 *      Drive/Calendar writes at all.
 *   3. If `insertPdfIntoEvent` is true: find the portal's own ticket PDF
 *      among the message's attachments (via
 *      TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL, e.g.
 *      findKinoArtTicketPdfAttachment), and MOVE IT DIRECTLY into the
 *      PERMANENT CONFIG.ticketAttachmentDriveFolderName folder —
 *      skipping BOTH the temp-folder upload AND the Drive-to-Docs OCR
 *      conversion entirely, since this mode never needs the PDF's TEXT,
 *      only the file itself as an attachment. Renamed via
 *      buildTicketAttachmentFilename, same convention as the PDF-sourced
 *      mode. If no matching PDF attachment is found even though the
 *      toggle is on, the event is still created, just without an
 *      attachment (a missing expected attachment should never block the
 *      calendar event itself, logged for visibility). If
 *      `insertPdfIntoEvent` is false: no PDF attachment is touched or
 *      uploaded at all — nothing to clean up, since nothing was ever
 *      created in Drive.
 *   4. Build and insert the Calendar event via the SHARED
 *      createTicketCalendarEvent.
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

  let attachmentInfo = null;
  if (portal.insertPdfIntoEvent) {
    const findPortalTicketPdfAttachment = TICKET_BODY_MODE_PDF_FINDERS_BY_IDENTIFYING_EMAIL[ticketingExtractEmailAddress(portal.identifyingEmail)];
    const pdfAttachment = findPortalTicketPdfAttachment ? findPortalTicketPdfAttachment(message) : null;

    if (pdfAttachment) {
      const permanentFolder = getOrCreateDriveFolderByName(CONFIG.ticketAttachmentDriveFolderName);
      const permanentPdfFile = permanentFolder.createFile(pdfAttachment.copyBlob());
      permanentPdfFile.setName(buildTicketAttachmentFilename(parsedTicket.eventName, parsedTicket, parsedTicket.ticketIdentifier));
      attachmentInfo = {
        fileId: permanentPdfFile.getId(),
        fileUrl: permanentPdfFile.getUrl(),
        title: permanentPdfFile.getName(),
      };
    } else {
      console.log(
        'Ticketing portal: insertPdfIntoEvent is true but no matching ticket PDF attachment was found on the message; creating the event without an attachment.'
      );
    }
  }

  // DIAGNOSTIC VISIBILITY (quick-260731-kar, round 2 -- see this function's
  // caller-side investigation for the real incident this responds to): a
  // "no calendar event created" outcome with ZERO thrown error was, until
  // now, indistinguishable in the Executions log between "this action's
  // run() was never entered at all for this message" (e.g. a sender/config
  // mismatch upstream) and "this ran all the way through and either
  // correctly no-op'd via the dedup safety net (which already logs) or
  // reached this exact point and called Calendar.Events.insert." This log
  // line, paired with isDuplicateTicketPurchase's existing "already exists,
  // skipping" log, makes those three outcomes trivially distinguishable by
  // reading the Executions log for the run in question -- same "make the
  // NEXT diagnostic artifact visible up front" philosophy already
  // established by parseEnigooTicketText's own diagnostic-on-failure
  // rewrite (quick-260731-tix round 3).
  console.log(
    'Ticketing portal: creating calendar event for "' + parsedTicket.eventName + '" (ticketIdentifier=' +
      parsedTicket.ticketIdentifier + ') on calendar ' + calendarId + '.'
  );
  createTicketCalendarEvent(parsedTicket, calendarId, attachmentInfo);
}

/**
 * TICKETING_PORTALS_ACTION — the ticketing-portals action descriptor.
 * Carries its own config block (TICKETING_PORTALS_ACTION_CONFIG),
 * independent of CONFIG and of any other action's config (except for the
 * one shared cross-cutting CONFIG.ticketAttachmentDriveFolderName field —
 * see that field's own doc comment in src/01-setup.js for why it lives
 * there rather than here).
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
   * appliesTo — returns a literal boolean (quick-260731-kar: now checks
   * BOTH processing modes). True when any message on the thread is from a
   * sender matching a configured TICKETING_PORTALS entry
   * (resolveTicketingPortal) AND either: (a) that portal resolves against
   * TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL (PDF/OCR-sourced) AND the
   * message carries at least one qualifying PDF attachment
   * (findTicketPdfAttachments), or (b) that portal resolves against
   * TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL (body-sourced) — no
   * attachment requirement at all, since a body-sourced portal's event
   * data comes from the message body itself, always present. Otherwise
   * false. dispatchActions only skips on a strict `=== false`, so a
   * literal boolean is required.
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
      if (TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL[senderKey]) {
        return true;
      }
    }

    return false;
  },

  /**
   * run — builds the processing job list via the pure, TESTABLE
   * resolveTicketProcessingJobs (each job tagged with a `mode`, one of
   * `'pdf'` or `'body'` — see that function's own JSDoc for the full
   * two-processing-mode architecture introduced in quick-260731-kar),
   * then runs the appropriate pipeline exactly once per job:
   * processTicketPdfAttachment for `mode: 'pdf'` jobs (the Drive/OCR
   * pipeline), processTicketFromMessageBody for `mode: 'body'` jobs (body
   * text only, no Drive/OCR at all). Duplicate-event protection for the
   * SAME purchase is provided entirely by the DEDUP SAFETY NET shared by
   * both pipelines (isDuplicateTicketPurchase — see this file's
   * class-level "ONE-EVENT-PER-PURCHASE" doc), not by restricting which
   * attachments/messages get processed here. A message matching neither a
   * configured portal nor either processing mode's requirements
   * contributes no job at all and is skipped gracefully (never throws for
   * a non-matching message).
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
// Exports every pure function (resolveTicketingPortal, parseEnigooTicketText,
// addMinutesToWallClockComponents, formatWallClockComponentsIso) and
// TICKETING_PORTALS_ACTION (action registry). Also exports
// isTicketPdfAttachment/findTicketPdfAttachments/resolveTicketProcessingJobs
// (live-test-driven, quick-260731-tix rounds 8-9): despite living alongside
// the GAS-only Drive/OCR/Calendar pipeline, none of these three reference a
// real GAS global directly -- they only invoke methods ON THE PASSED-IN
// message/attachment objects, so they are fully testable under Node with
// plain duck-typed fakes (see resolveTicketProcessingJobs's own JSDoc).
// getOrCreateDriveFolderByName/processTicketPdfAttachment/
// findTicketEventByIdentifier remain genuinely GAS-only (reference
// DriveApp/Drive/DocumentApp/CalendarApp/Calendar globals directly) and
// are NOT exported — they are never invoked under Node.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    ticketingExtractEmailAddress: ticketingExtractEmailAddress,
    resolveTicketingPortal: resolveTicketingPortal,
    resolveTicketingCalendarId: resolveTicketingCalendarId,
    parseEnigooTicketText: parseEnigooTicketText,
    parseKinoArtTicketText: parseKinoArtTicketText,
    TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL: TICKET_TEXT_PARSERS_BY_IDENTIFYING_EMAIL,
    TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL: TICKET_BODY_PARSERS_BY_IDENTIFYING_EMAIL,
    DEFAULT_EVENT_DURATION_MINUTES: DEFAULT_EVENT_DURATION_MINUTES,
    addMinutesToWallClockComponents: addMinutesToWallClockComponents,
    formatWallClockComponentsIso: formatWallClockComponentsIso,
    buildTicketAttachmentFilename: buildTicketAttachmentFilename,
    TICKETING_TEMP_DRIVE_FOLDER_NAME: TICKETING_TEMP_DRIVE_FOLDER_NAME,
    isTicketPdfAttachment: isTicketPdfAttachment,
    findTicketPdfAttachments: findTicketPdfAttachments,
    findKinoArtTicketPdfAttachment: findKinoArtTicketPdfAttachment,
    resolveTicketProcessingJobs: resolveTicketProcessingJobs,
    TICKETING_PORTALS_ACTION: TICKETING_PORTALS_ACTION,
  };
}

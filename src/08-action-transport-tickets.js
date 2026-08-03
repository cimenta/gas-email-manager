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
 * resolveTransportProcessingJobs — the pure, TESTABLE extraction of `run`'s
 * per-message orchestration decision, mirroring resolveTicketProcessingJobs'
 * shape (src/07-action-ticketing-portals.js). Given `messages` (an array of
 * message-like objects exposing `getFrom()`/`getAttachments()` — GAS
 * `GmailMessage` objects in production, plain duck-typed fakes in tests)
 * and `senders` (the TRANSPORT_SENDERS config array), returns an array of
 * `{ message, sender, icsAttachments }` jobs: for each message, resolve its
 * sender; skip if no match; collect its `.ics` attachments (via the REUSED
 * isIcsAttachment, src/05-action-ics-import.js — D-01, no second matcher
 * either); skip if none. EXACTLY ONE JOB PER MATCHING MESSAGE (never one
 * per attachment) — deliberate: this action's real duplicate-event
 * guarantee comes from the DEDUP SAFETY NET (isDuplicateTransportTicket)
 * applied per-EVENT inside processTransportTicketJob, not from restricting
 * which messages/attachments get processed here — the corrected lesson
 * from the ticketing-portals action's own round-8/round-9 double-booking
 * incident (see resolveTicketProcessingJobs' own JSDoc for that full
 * writeup: "select only the first attachment" was the WRONG fix there, and
 * the same reasoning applies here — a message could legitimately carry
 * more than one `.ics` attachment, and every one of them must still be
 * parsed and dedup-checked, just as one job, not silently dropped). Pure,
 * no GAS globals — every GAS-shaped method call here is invoked ON THE
 * PASSED-IN objects only, never a real global service, so this is fully
 * unit-testable under Node with fake message/attachment objects.
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

    const icsAttachments = message.getAttachments().filter(isIcsAttachment);
    if (icsAttachments.length === 0) {
      continue;
    }

    jobs.push({ message: message, sender: sender, icsAttachments: icsAttachments });
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
 * resolveTransportProcessingJobs' own JSDoc for the `{ message, sender,
 * icsAttachments }` shape), in this exact order so nothing is EVER created
 * in Drive before the decision to write is final (see this file's
 * class-level "ORDERING GUARANTEE" doc):
 *   1. Resolve the calendar ONCE via resolveTransportCalendarId(job.sender,
 *      CONFIG.calendarId) and thread it through every downstream call
 *      (never re-read `sender.calendarId` at a call site — see
 *      resolveTransportCalendarId's own JSDoc for the real live-crash class
 *      this avoids).
 *   2. Parse EVERY `.ics` attachment's `getDataAsString()` through the
 *      REUSED parseIcs (D-01) and concatenate the resulting events — all
 *      parsing completes before any write begins (fail-closed, same
 *      discipline as ICS_CALENDAR_ACTION.run).
 *   3. Compute each event's `ticketIdentifier` (extractTransportTicketIdentifier)
 *      and drop the ones isDuplicateTransportTicket reports as already
 *      present. If nothing remains, return immediately — no Drive upload,
 *      no write.
 *   4. If `job.sender.insertPdfIntoEvent` is true: find the ticket PDF
 *      (findTransportTicketPdfAttachment, which excludes invoice.pdf); if
 *      found, copy its blob DIRECTLY into
 *      getOrCreateTransportDriveFolder(CONFIG.ticketAttachmentDriveFolderName)
 *      — no temp-folder hop, unlike the ticketing-portals action's OCR
 *      pipeline, since this action never needs the PDF's text — rename it
 *      via buildTransportAttachmentFilename BEFORE reading its id/url/name
 *      (the displayed Calendar attachment title is derived from the file's
 *      name AT THAT POINT), and build `{ fileId, fileUrl, title }`. If no
 *      ticket PDF is found, log and continue without one — a missing
 *      attachment must never block the calendar event itself.
 *   5. For each remaining event: buildEventResource(event) (REUSED, D-01),
 *      then add `resource.extendedProperties = { private: {
 *      ticketIdentifier } }` when the identifier is truthy, and, when an
 *      attachment exists, `resource.attachments = [{ fileId, fileUrl,
 *      title, mimeType: 'application/pdf' }]` plus
 *      `optionalArgs.supportsAttachments = true` (the exact live-proven
 *      shape from createTicketCalendarEvent, src/07-action-ticketing-
 *      portals.js). Logs one line naming the summary, the identifier and
 *      the calendar before writing (the round-2 diagnostic convention
 *      established by the ticketing-portals action). Writes via
 *      importIcsEventWithSequenceRetry(resource, calendarId, event.uid,
 *      optionalArgs) when the event has a UID — REUSING
 *      ICS_CALENDAR_ACTION's proven idempotent-by-iCalUID path (D-01),
 *      which also protects against Gmail's own native invite detection
 *      creating a second event from the same .ics — and
 *      Calendar.Events.insert(resource, calendarId, optionalArgs)
 *      otherwise.
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
 * (resolveTransportCalendarId, extractTransportTicketIdentifier,
 * findTransportTicketPdfAttachment, buildTransportAttachmentFilename,
 * parseIcs, buildEventResource) IS fully unit-tested.
 */
function processTransportTicketJob(job) {
  const calendarId = resolveTransportCalendarId(job.sender, CONFIG.calendarId);

  const events = job.icsAttachments.reduce(function (allEvents, attachment) {
    return allEvents.concat(parseIcs(attachment.getDataAsString()));
  }, []);

  // WR-01 fix (260803-us3 review): a UID-less write goes through the
  // non-idempotent Calendar.Events.insert fallback (see the write loop
  // below), so two same-batch entries sharing a ticketIdentifier would BOTH
  // pass isDuplicateTransportTicket (neither exists in Calendar yet at
  // check time) and both get inserted — a real duplicate. seenInBatch
  // tracks identifiers already claimed earlier in this same pass so a
  // repeat is dropped before it ever reaches isDuplicateTransportTicket's
  // live-state check, independent of whether the event carries a UID.
  const seenInBatch = {};
  const eventsToCreate = events
    .map(function (event) {
      return { event: event, ticketIdentifier: extractTransportTicketIdentifier(event) };
    })
    .filter(function (entry) {
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

  if (eventsToCreate.length === 0) {
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
        buildTransportAttachmentFilename(eventsToCreate[0].event.summary, eventsToCreate[0].event.start, eventsToCreate[0].ticketIdentifier)
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

  eventsToCreate.forEach(function (entry) {
    const event = entry.event;
    const resource = buildEventResource(event);

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
      'Transport tickets: creating calendar event for "' + event.summary + '" (ticketIdentifier=' + entry.ticketIdentifier + ') on calendar ' + calendarId + '.'
    );

    if (event.uid) {
      importIcsEventWithSequenceRetry(resource, calendarId, event.uid, optionalArgs);
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
   * least one job on the thread (a message from a configured transport
   * sender carrying at least one `.ics` attachment); false otherwise.
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
// buildTransportAttachmentFilename, resolveTransportProcessingJobs) and
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
    TRANSPORT_TICKETS_ACTION: TRANSPORT_TICKETS_ACTION,
  };
}

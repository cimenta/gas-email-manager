/**
 * TICKETING_PORTALS_ACTION_CONFIG — the config block for
 * TICKETING_PORTALS_ACTION (defined in the sibling
 * src/07-action-ticketing-portals.js). Same "*-action-cfg-*.js" split
 * pattern already established for the ICS and booking.com actions — see
 * their own config files' class-level JSDoc for the full load-order/getter
 * rationale (an ES6 getter is not evaluated at object-construction time,
 * only when something actually reads the property, which is what makes
 * this split safe regardless of which of the two sibling files loads first
 * alphabetically). Every field below is an ES6 GETTER, PropertiesService-
 * backed with a code-default fallback, per the Script Properties
 * live-config override feature (quick-260726-spr) — see
 * src/01-setup.js's class-level JSDoc for the full mechanism.
 *
 * v0.6.0 NEW ACTION (quick-260731-tix): detects ticket-purchase
 * confirmation emails (concerts, theater, cinema, events) from configured
 * ticketing portals. See the sibling action file's class-level JSDoc for
 * the full architecture (the Drive/OCR pipeline, one-event-per-purchase
 * design, why all portal-specific parsers live in that ONE file rather
 * than one file per portal, etc.) — including, since quick-260731-kar,
 * the SECOND body-sourced processing mode (Kino Art) alongside the
 * original PDF/OCR-sourced mode (enigoo.cz).
 */
const TICKETING_PORTALS_ACTION_CONFIG = {
  // Cross-cutting per-action enable/disable toggle, same convention as
  // every other action's config. Script Property override:
  // 07-action-ticketing-portals-ENABLED (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('07-action-ticketing-portals-ENABLED', true);
  },

  // Notify the script owner when this action throws. Script Property
  // override: 07-action-ticketing-portals-NOTIFY_ON_FAILURE (boolean).
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('07-action-ticketing-portals-NOTIFY_ON_FAILURE', true);
  },

  // TICKETING_PORTALS: array of { identifyingEmail, calendarId,
  // insertPdfIntoEvent }. Each entry identifies one supported ticketing
  // portal by its confirmation email's sender address (case-insensitive
  // match via resolveTicketingPortal in the sibling action file — same
  // comparison convention as every other sender-matching helper in this
  // codebase), the calendar its events should be created on, and whether
  // the original ticket PDF should be attached to the created Calendar
  // event (see the action file's class-level "Drive/OCR pipeline" doc for
  // exactly what this toggle controls: the PDF is ALWAYS uploaded to the
  // temp Drive folder for OCR either way — this toggle only decides
  // whether it then survives in the permanent folder + becomes a real
  // Calendar attachment, or is deleted).
  //
  // The shipped default seeds FIVE entries: enigoo.cz (the original portal
  // this feature was built from, PDF/OCR-sourced — see the sibling action
  // file's parseEnigooTicketText), Kino Art (kinoart.cz, a Czech cinema,
  // added quick-260731-kar, BODY-SOURCED — see the sibling action file's
  // parseKinoArtTicketText and its class-level "TWO PROCESSING MODES" doc
  // for the full architecture), Ticketmaster CZ (ticketmaster.cz,
  // added quick-260816-ocw, also BODY-SOURCED — see the sibling action
  // file's parseTicketmasterCzTicketText), Entradio
  // (no-reply@app.entradio.cz, added debug/entradio-portal-not-supported,
  // also BODY-SOURCED — see the sibling action file's
  // parseEntradioTicketText), and Fever (hello@feverup.com, added
  // quick-260921-gj0, also BODY-SOURCED — see the sibling action file's
  // parseFeverTicketText). All five entries ship with
  // calendarId left null and insertPdfIntoEvent left false — the owner
  // fills in the real calendar ID and decides the attachment toggle live,
  // per entry, via rebuildScriptProperties() + Script Properties, matching
  // the now-established settings workflow (never committed to git — same
  // placeholder-calendar-ID convention as CONFIG.calendarId itself).
  //
  // FEVER (hello@feverup.com): the sender also sends ordinary marketing
  // mail, which is why a content detector (feverTextHasPurchaseDetails) is
  // registered for it in the sibling action file — applied from day one
  // rather than discovered live, per debug/ticketmaster-cz-order-confirm.
  // Unlike Entradio, this portal DOES have a registered PDF finder
  // (findFeverTicketPdfAttachment), so turning insertPdfIntoEvent on
  // attaches the real ticket PDF the confirmation email carries.
  //
  // ENTRADIO IS A PLATFORM, NOT A VENUE (worth knowing before adding a
  // "missing" venue here): app.entradio.cz is a white-label ticketing system
  // that many venues send through — the sample this entry was built from
  // came from Kino Metropol Olomouc, but the SENDER address is shared across
  // every venue on the platform, so this ONE entry already covers all of
  // them. Its parser anchors on Entradio's own template structure, never on
  // any single venue's name. A new Entradio venue needs no config change at
  // all; it only needs its own calendarId here if the owner wants it routed
  // somewhere other than this entry's calendar (a per-venue split this
  // sender-keyed config shape cannot express — it would need a different
  // mechanism, and no such need exists today).
  //
  // ENTRADIO AND insertPdfIntoEvent (REWRITTEN IN ROUND 2 of
  // debug/entradio-portal-not-supported — round 1's version of this note said
  // leaving it false was "the only meaningful setting", which is NO LONGER
  // TRUE and is corrected here rather than left to contradict the code):
  //
  // For this portal the toggle means "ALSO DOWNLOAD THE TICKET FILE". An
  // Entradio confirmation carries no ticket PDF among its attachments (only
  // the venue's terms and conditions), so the sibling action file still
  // registers no PDF finder for it — but it DOES register an attachment
  // FETCHER (fetchEntradioAttachments), which follows the "STÁHNOUT
  // VSTUPENKY" link over HTTP when this toggle is ON and attaches the
  // downloaded file to the event.
  //
  // THE PER-SEAT QR CODES ARE NOT GATED BY THIS TOGGLE AT ALL. They are
  // fetched from Entradio's own qrcode endpoint and attached on EVERY
  // Entradio event regardless, because a QR code is not a PDF and it is the
  // artifact that actually gets the owner through the door. So the shipped
  // `false` below is a real, usable setting rather than a placeholder: the
  // event still arrives with one QR attachment per seat.
  //
  // NOTE: this portal is the reason src/appsscript.json now declares the
  // script.external_request OAuth scope — the first outbound HTTP in this
  // project. See the sibling action file's "ENTRADIO ATTACHMENT PIPELINE"
  // section.
  //
  // Script Property override: 07-action-ticketing-portals-TICKETING_PORTALS
  // (json — array of {identifyingEmail, calendarId, insertPdfIntoEvent}
  // objects).
  // Script Properties value example (MUST be valid JSON — double-quoted
  // keys AND double-quoted string values, unlike a JS object literal; see
  // ICS_ACTION_CONFIG.calendarIdBySender's own comment in
  // src/05-action-cfg-ics-import.js for the exact same JSON-vs-JS-object-
  // literal pitfall a real owner mistake already hit once for that other
  // JSON-typed setting):
  // [{"identifyingEmail":"no-reply@enigoo.cz","calendarId":"abc123@group.calendar.google.com","insertPdfIntoEvent":true},{"identifyingEmail":"rezervace@kinoart.cz","calendarId":"def456@group.calendar.google.com","insertPdfIntoEvent":false},{"identifyingEmail":"noreply@ticketmaster.cz","calendarId":"ghi789@group.calendar.google.com","insertPdfIntoEvent":false},{"identifyingEmail":"no-reply@app.entradio.cz","calendarId":"jkl012@group.calendar.google.com","insertPdfIntoEvent":false},{"identifyingEmail":"hello@feverup.com","calendarId":"mno345@group.calendar.google.com","insertPdfIntoEvent":false}]
  get ticketingPortals() {
    return getJsonSetting(
      '07-action-ticketing-portals-TICKETING_PORTALS',
      [
        { identifyingEmail: 'no-reply@enigoo.cz', calendarId: null, insertPdfIntoEvent: false },
        { identifyingEmail: 'rezervace@kinoart.cz', calendarId: null, insertPdfIntoEvent: false },
        { identifyingEmail: 'noreply@ticketmaster.cz', calendarId: null, insertPdfIntoEvent: false },
        { identifyingEmail: 'no-reply@app.entradio.cz', calendarId: null, insertPdfIntoEvent: false },
        { identifyingEmail: 'hello@feverup.com', calendarId: null, insertPdfIntoEvent: false },
      ],
      isValidTicketingPortalsShape
    );
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getBooleanSetting/getJsonSetting) and the shape validator
// (isValidTicketingPortalsShape) — all defined in the sibling
// src/01-setup.js. Under GAS's shared global scope these are ALREADY
// visible here by bare name (01-setup.js loads alphabetically first) — no
// action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare references inside TICKETING_PORTALS_ACTION_CONFIG's getters
// above would otherwise throw ReferenceError. Same `globalThis` bridge
// technique already established by the ICS and booking config files'
// equivalent bridges for these same kinds of helpers.
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getJsonSetting = settingsHelpers.getJsonSetting;
  globalThis.isValidTicketingPortalsShape = settingsHelpers.isValidTicketingPortalsShape;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { TICKETING_PORTALS_ACTION_CONFIG: TICKETING_PORTALS_ACTION_CONFIG };
}

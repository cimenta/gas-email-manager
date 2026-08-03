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
  // The shipped default seeds TWO entries: enigoo.cz (the original portal
  // this feature was built from, PDF/OCR-sourced — see the sibling action
  // file's parseEnigooTicketText) and, added quick-260731-kar, Kino Art
  // (kinoart.cz, a Czech cinema, BODY-SOURCED — see the sibling action
  // file's parseKinoArtTicketText and its class-level "TWO PROCESSING
  // MODES" doc for the full architecture). Both entries ship with
  // calendarId left null and insertPdfIntoEvent left false — the owner
  // fills in the real calendar ID and decides the attachment toggle live,
  // per entry, via rebuildScriptProperties() + Script Properties, matching
  // the now-established settings workflow (never committed to git — same
  // placeholder-calendar-ID convention as CONFIG.calendarId itself).
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
  // [{"identifyingEmail":"no-reply@enigoo.cz","calendarId":"abc123@group.calendar.google.com","insertPdfIntoEvent":true},{"identifyingEmail":"rezervace@kinoart.cz","calendarId":"def456@group.calendar.google.com","insertPdfIntoEvent":false}]
  get ticketingPortals() {
    return getJsonSetting(
      '07-action-ticketing-portals-TICKETING_PORTALS',
      [
        { identifyingEmail: 'no-reply@enigoo.cz', calendarId: null, insertPdfIntoEvent: false },
        { identifyingEmail: 'rezervace@kinoart.cz', calendarId: null, insertPdfIntoEvent: false },
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

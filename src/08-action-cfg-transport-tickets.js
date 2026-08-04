/**
 * TRANSPORT_TICKETS_ACTION_CONFIG — the config block for
 * TRANSPORT_TICKETS_ACTION (defined in the sibling
 * src/08-action-transport-tickets.js). Same "*-action-cfg-*.js" split
 * pattern already established for the ICS, booking.com, and
 * ticketing-portals actions — see their own config files' class-level
 * JSDoc for the full load-order/getter rationale (an ES6 getter is not
 * evaluated at object-construction time, only when something actually
 * reads the property, which is what makes this split safe regardless of
 * which of the two sibling files loads first alphabetically). Every field
 * below is an ES6 GETTER, PropertiesService-backed with a code-default
 * fallback, per the Script Properties live-config override feature
 * (quick-260726-spr) — see src/01-setup.js's class-level JSDoc for the
 * full mechanism.
 *
 * quick-260803-us3 NEW ACTION: detects train/bus ticket confirmation
 * emails (RegioJet first, extensible to other carriers) whose event data
 * comes from the email's OWN text/calendar .ics VEVENT — see the sibling
 * action file's class-level JSDoc for the full architecture (why this
 * reuses the existing ICS parser instead of hand-rolling a second one, the
 * PDF archive/attach flow, the ticketIdentifier dedup safety net).
 *
 * quick-260804-bs7: IDOS.cz added as a SECOND carrier, via a new `mode`
 * field (`'ics' | 'body'`) on each `transportSenders` entry — see the
 * sibling action file's class-level JSDoc for the full two-processing-mode
 * architecture this mirrors from TICKETING_PORTALS_ACTION_CONFIG's own
 * `mode: 'pdf'|'body'` split.
 */
const TRANSPORT_TICKETS_ACTION_CONFIG = {
  // Cross-cutting per-action enable/disable toggle, same convention as
  // every other action's config. Script Property override:
  // 08-action-transport-tickets-ENABLED (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('08-action-transport-tickets-ENABLED', true);
  },

  // Notify the script owner when this action throws. Script Property
  // override: 08-action-transport-tickets-NOTIFY_ON_FAILURE (boolean).
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('08-action-transport-tickets-NOTIFY_ON_FAILURE', true);
  },

  // TRANSPORT_SENDERS: array of { identifyingEmail, calendarId,
  // insertPdfIntoEvent, mode } — the SAME {identifyingEmail, calendarId,
  // insertPdfIntoEvent} shape as TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals
  // (src/07-action-cfg-ticketing-portals.js), validated with the SAME
  // shared isValidTicketingPortalsShape (see src/01-setup.js's own JSDoc on
  // that function for why it is shared, not cloned -- it ignores the extra
  // `mode` key by design, so no fork was needed for this new field). Each
  // entry identifies one supported transport carrier by its confirmation
  // email's sender address (case-insensitive match via
  // resolveTransportSender in the sibling action file), the calendar its
  // events should be created on, whether the ticket PDF should be archived
  // + attached to the created Calendar event, and its processing `mode`
  // (quick-260804-bs7, D-01):
  //   - `'ics'` — event data comes from the email's own text/calendar .ics
  //     VEVENT (RegioJet).
  //   - `'body'` — event data is parsed directly from the email's plain-
  //     text body, no .ics attachment at all (IDOS.cz).
  // An entry with NO `mode` field resolves to `'ics'` UNLESS a body parser
  // is registered for that sender's address (see
  // resolveTransportSenderMode's own JSDoc in the sibling action file) —
  // this is what keeps the owner's already-live Script Property JSON
  // (written before this `mode` field existed) working with zero edits.
  //
  // The shipped default seeds TWO entries: jizdenky@regiojet.cz (mode
  // 'ics', explicit for clarity even though it is also the fallback) and
  // jizdenky@idos.svt.cz (mode 'body', quick-260804-bs7 -- D-05, extensible
  // sender registry, future carriers are added the same way). calendarId
  // left null and insertPdfIntoEvent left false for the new entry — the
  // owner fills in the real calendar ID and decides the attachment toggle
  // live, per entry, via rebuildScriptProperties() + Script Properties,
  // matching the now-established settings workflow (never committed to
  // git — same placeholder-calendar-ID convention as CONFIG.calendarId
  // itself).
  //
  // A sender listed here is expected to ALSO be listed in
  // ICS_ACTION_CONFIG.excludeFrom (src/05-action-cfg-ics-import.js, D-03)
  // so ICS_CALENDAR_ACTION stands down for it — an owner-side Script
  // Properties step, not a code default; see this action's own class-level
  // JSDoc for the full rationale. This only matters for 'ics'-mode senders
  // — a 'body'-mode sender never carries a .ics attachment ICS_CALENDAR_ACTION
  // could otherwise compete over.
  //
  // Script Property override:
  // 08-action-transport-tickets-TRANSPORT_SENDERS (json — array of
  // {identifyingEmail, calendarId, insertPdfIntoEvent, mode} objects).
  // Script Properties value example (MUST be valid JSON — double-quoted
  // keys AND double-quoted string values, unlike a JS object literal; see
  // ICS_ACTION_CONFIG.calendarIdBySender's own comment in
  // src/05-action-cfg-ics-import.js for the exact same JSON-vs-JS-object-
  // literal pitfall a real owner mistake already hit once for that other
  // JSON-typed setting):
  // [{"identifyingEmail":"jizdenky@regiojet.cz","calendarId":"abc123@group.calendar.google.com","insertPdfIntoEvent":true,"mode":"ics"},{"identifyingEmail":"jizdenky@idos.svt.cz","calendarId":"def456@group.calendar.google.com","insertPdfIntoEvent":true,"mode":"body"}]
  get transportSenders() {
    return getJsonSetting(
      '08-action-transport-tickets-TRANSPORT_SENDERS',
      [
        { identifyingEmail: 'jizdenky@regiojet.cz', calendarId: null, insertPdfIntoEvent: false, mode: 'ics' },
        { identifyingEmail: 'jizdenky@idos.svt.cz', calendarId: null, insertPdfIntoEvent: false, mode: 'body' },
      ],
      isValidTicketingPortalsShape
    );
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getBooleanSetting/getJsonSetting) and the shared shape validator
// (isValidTicketingPortalsShape) — all defined in the sibling
// src/01-setup.js. Under GAS's shared global scope these are ALREADY
// visible here by bare name (01-setup.js loads alphabetically first) — no
// action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare references inside TRANSPORT_TICKETS_ACTION_CONFIG's getters
// above would otherwise throw ReferenceError. Same `globalThis` bridge
// technique already established by the ICS, booking, and ticketing-portals
// config files' equivalent bridges for these same kinds of helpers.
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getJsonSetting = settingsHelpers.getJsonSetting;
  globalThis.isValidTicketingPortalsShape = settingsHelpers.isValidTicketingPortalsShape;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { TRANSPORT_TICKETS_ACTION_CONFIG: TRANSPORT_TICKETS_ACTION_CONFIG };
}

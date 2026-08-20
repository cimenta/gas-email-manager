/**
 * MEETINGS_ACTION_CONFIG — the config block for MEETINGS_ACTION (defined in
 * the sibling src/10-action-meetings.js). Same "*-action-cfg-*.js" split
 * pattern already established for every other action — see
 * src/07-action-cfg-ticketing-portals.js's own class-level JSDoc for the
 * full load-order/getter rationale (an ES6 getter is not evaluated at
 * object-construction time, only when something actually reads the
 * property, which is what makes this split safe regardless of which of the
 * two sibling files loads first alphabetically). Every field below is an ES6
 * GETTER, PropertiesService-backed with a code-default fallback, per the
 * Script Properties live-config override feature — see src/01-setup.js's
 * class-level JSDoc for the full mechanism.
 *
 * quick-260820-g4r NEW ACTION: turns general meeting-invite emails carrying
 * NO `.ics` attachment into calendar events, matched to a supported "system"
 * (Teamio ships first) by SENDER DOMAIN pattern. See the sibling action
 * file's class-level JSDoc for the full architecture.
 */
const MEETINGS_ACTION_CONFIG = {
  // Cross-cutting per-action enable/disable toggle, same convention as
  // every other action's config. Script Property override:
  // 10-action-meetings-ENABLED (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('10-action-meetings-ENABLED', true);
  },

  // Notify the script owner when this action throws. Script Property
  // override: 10-action-meetings-NOTIFY_ON_FAILURE (boolean).
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('10-action-meetings-NOTIFY_ON_FAILURE', true);
  },

  // The fallback event duration (minutes) used ONLY when a matched
  // invitation's own body states no explicit duration (D-08) -- Teamio's
  // real body ("délka 30 minut") always states one, so this fallback exists
  // for a future system/variant that doesn't. Script Property override:
  // 10-action-meetings-DEFAULT_DURATION_MINUTES (number).
  // Script Properties value example: 60
  get defaultDurationMinutes() {
    return getNumberSetting('10-action-meetings-DEFAULT_DURATION_MINUTES', 60);
  },

  // MEETING_SYSTEMS: array of { domainPattern, calendarId }. Each entry
  // identifies one supported meeting-invitation "system" by a SENDER DOMAIN
  // pattern (D-02 -- see the sibling action file's
  // meetingsDomainMatchesPattern for the exact `*.` subdomain-wildcard/
  // apex/dot-boundary matching semantics) and the calendar its events should
  // be created on. Adding a further system needs ONE new entry here PLUS one
  // new parser function registered in the sibling action file's
  // MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN (keyed by the SAME domainPattern
  // string) -- the matching logic itself never changes.
  //
  // The shipped default seeds ONE entry: Teamio (*.teamio.com, the real
  // system this feature was built from -- see the sibling action file's
  // parseTeamioMeetingText). Ships with calendarId left null -- the owner
  // fills in the real calendar ID live via Script Properties or the admin
  // web app, never committed to git, same placeholder-calendar-ID
  // convention as every other action's own optional calendar override.
  //
  // Script Property override: 10-action-meetings-MEETING_SYSTEMS (json --
  // array of {domainPattern, calendarId} objects).
  // Script Properties value example (MUST be valid JSON -- double-quoted
  // keys AND double-quoted string values, unlike a JS object literal; see
  // ICS_ACTION_CONFIG.calendarIdBySender's own comment in
  // src/05-action-cfg-ics-import.js for the exact same JSON-vs-JS-object-
  // literal pitfall a real owner mistake already hit once for that other
  // JSON-typed setting):
  // [{"domainPattern":"*.teamio.com","calendarId":"abc123@group.calendar.google.com"}]
  get meetingSystems() {
    return getJsonSetting('10-action-meetings-MEETING_SYSTEMS', [{ domainPattern: '*.teamio.com', calendarId: null }], isValidMeetingSystemsShape);
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getBooleanSetting/getNumberSetting/getJsonSetting) and the shape
// validator (isValidMeetingSystemsShape) — all defined in the sibling
// src/01-setup.js. Under GAS's shared global scope these are ALREADY
// visible here by bare name (01-setup.js loads alphabetically first) — no
// action needed, and this `if` block never executes there. Under Node,
// each `require()`d file is its own isolated module with its own scope, so
// the bare references inside MEETINGS_ACTION_CONFIG's getters above would
// otherwise throw ReferenceError. Same `globalThis` bridge technique
// already established by every other action's config file.
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getNumberSetting = settingsHelpers.getNumberSetting;
  globalThis.getJsonSetting = settingsHelpers.getJsonSetting;
  globalThis.isValidMeetingSystemsShape = settingsHelpers.isValidMeetingSystemsShape;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { MEETINGS_ACTION_CONFIG: MEETINGS_ACTION_CONFIG };
}

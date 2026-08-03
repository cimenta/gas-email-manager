/**
 * BOOKING_ACTION_CONFIG — the config block for BOOKING_MANAGEMENT_ACTION
 * (defined in the sibling src/06-action-booking-com-management.js). Split
 * out into its own file per owner-directed refactor: each action's
 * tunable settings live in a sibling "{n}-action-cfg-{name}.js" file,
 * separate from the action's logic. GAS's shared global scope means this
 * constant is still visible by bare name inside
 * 06-action-booking-com-management.js — no require/import needed for the
 * GAS runtime path.
 *
 * SCOPE NOTE: this file carries ONLY this action's own tunable settings.
 * It does NOT carry BOOKING_LANGUAGE_PACKS/getBookingLabels or any
 * *_LANGUAGE_PACK — those are the language-pack MECHANISM (shared
 * infrastructure for resolving this action's language-dependent text),
 * not owner-tunable settings, and remain in
 * src/06-action-booking-com-management.js / the 06-lang-*.js files.
 *
 * LOAD-ORDER SAFETY: BOOKING_MANAGEMENT_ACTION's own `config` property is
 * an ES6 GETTER (`get config() { return BOOKING_ACTION_CONFIG; }`), not a
 * plain literal reference — a getter is NOT evaluated when the action's
 * object literal is constructed, only when something actually reads
 * `action.config`, which happens inside function bodies long after every
 * project file has finished loading. This is the SAME "resolve lazily, at
 * the point of use" pattern already proven in this file for
 * BOOKING_LANGUAGE_PACKS/getBookingLabels, applied here so this file and
 * src/06-action-booking-com-management.js can load in EITHER alphabetical
 * order safely — which matters concretely here, since
 * "06-action-booking-com-management.js" alphabetically precedes
 * "06-action-cfg-booking-com-management.js" ('booking' < 'cfg'), the
 * OPPOSITE of what a plain literal `config: BOOKING_ACTION_CONFIG`
 * property would have required (the config would need to already exist
 * before the action file's own top-level object-literal construction
 * runs, which it would not, under naive alphabetical load order).
 *
 * SCRIPT PROPERTIES LIVE-CONFIG OVERRIDE (owner-requested, quick-260726-spr):
 * every field below (including the two NESTED addToCalendar.enabled/
 * removeFromCalendar.enabled fields) is now an ES6 GETTER (same
 * load-order-safety reasoning as `config` above, applied one level
 * deeper), resolving through PropertiesService via the shared
 * get*Setting/getStringSetting/etc. accessor helpers defined in
 * src/01-setup.js, with the CURRENT literal value preserved as each
 * getter's code-default fallback. See src/01-setup.js's own class-level
 * JSDoc ("SCRIPT PROPERTIES LIVE-CONFIG OVERRIDE") for the full mechanism
 * and the approved Script Properties key list. Every existing reader of
 * these fields needs ZERO changes.
 */
const BOOKING_ACTION_CONFIG = {
  // Cross-cutting per-action enable/disable toggle (QT-260724-lqi-TOGGLE).
  // Strictly false skips this action entirely, before appliesTo/run.
  // Script Property override: 06-action-booking-com-ENABLED (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('06-action-booking-com-ENABLED', true);
  },

  // Notify the script owner when this action throws. Script Property
  // override: 06-action-booking-com-NOTIFY_ON_FAILURE (boolean).
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('06-action-booking-com-NOTIFY_ON_FAILURE', true);
  },

  // Sender allow-list: array of sender email addresses (case-insensitive;
  // a bare address or a full 'Name <email>' entry both work). An empty
  // array would allow any sender; booking.com's known sending address is
  // seeded here as the default. Script Property override:
  // 06-action-booking-com-SENDER_ALLOW_LIST (list, comma-separated).
  // Script Properties value example: alice@example.com, bob@example.com
  get senderAllowList() {
    return getListSetting('06-action-booking-com-SENDER_ALLOW_LIST', ['noreply@booking.com']);
  },

  // Sub-behavior toggles, each independently gated on the top-level
  // `enabled` above. The subject-substring matching itself is
  // language-dependent (see getBookingLabels('addToCalendarSubjectContains')
  // / getBookingLabels('removeFromCalendarSubjectContains') at the point
  // of use) — only the enable/disable flags live here. These stay plain
  // literal OBJECTS (addToCalendar/removeFromCalendar themselves are not
  // getters — only their own `enabled` property below is), since the
  // owner-approved Script Properties key list has one key per NESTED
  // boolean, not per parent object.
  addToCalendar: {
    // Script Property override: 06-action-booking-com-ADD_TO_CALENDAR_ENABLED
    // (boolean).
    // Script Properties value example: true
    get enabled() {
      return getBooleanSetting('06-action-booking-com-ADD_TO_CALENDAR_ENABLED', true);
    },
  },
  removeFromCalendar: {
    // Script Property override:
    // 06-action-booking-com-REMOVE_FROM_CALENDAR_ENABLED (boolean).
    // Script Properties value example: true
    get enabled() {
      return getBooleanSetting('06-action-booking-com-REMOVE_FROM_CALENDAR_ENABLED', true);
    },
  },

  // Days padded on each side of a real check-in/check-out pair when
  // building the Calendar search window (see computeSearchWindow). Script
  // Property override: 06-action-booking-com-SEARCH_WINDOW_PADDING_DAYS
  // (number).
  // Script Properties value example: 7
  get searchWindowPaddingDays() {
    return getNumberSetting('06-action-booking-com-SEARCH_WINDOW_PADDING_DAYS', 7);
  },

  // Fallback search-window half-width, in years each direction, used when
  // check-in/check-out cannot be derived at all (see computeSearchWindow).
  // Script Property override:
  // 06-action-booking-com-WIDE_FALLBACK_WINDOW_YEARS (number).
  // Script Properties value example: 1
  get wideFallbackWindowYears() {
    return getNumberSetting('06-action-booking-com-WIDE_FALLBACK_WINDOW_YEARS', 1);
  },

  // Tolerance (in days) padding a booking's date range when checking for
  // overlap against a candidate calendar event during fuzzy matching (see
  // eventDateRangeOverlaps). Deliberately small — once real check-in/
  // check-out dates are available, this only needs to absorb date-
  // parsing/rounding slop, not stand in for a missing date (contrast with
  // wideFallbackWindowYears above, an unrelated, much wider fallback).
  // Script Property override:
  // 06-action-booking-com-EVENT_OVERLAP_TOLERANCE_DAYS (number).
  // Script Properties value example: 1
  get eventOverlapToleranceDays() {
    return getNumberSetting('06-action-booking-com-EVENT_OVERLAP_TOLERANCE_DAYS', 1);
  },

  // MULTI-CALENDAR ROUTING (owner-requested): CONFIG.calendarId
  // (src/01-setup.js) is the DEFAULT/fallback calendar for every action.
  // This field, when truthy, overrides that default for ALL of this
  // action's Calendar API calls (list/patch/insert/remove), for both the
  // add and remove paths. null (the default) preserves original behavior
  // exactly — every call targets CONFIG.calendarId. No per-sender map is
  // offered here (unlike the ICS action) — booking.com only ever sends
  // from one known address in practice, so it was not requested and is not
  // needed. See resolveBookingCalendarId in the sibling
  // src/06-action-booking-com-management.js for the 2-tier resolution this
  // field participates in. Script Property override:
  // 06-action-booking-com-CALENDAR_ID (string).
  // Script Properties value example: abc123@group.calendar.google.com
  get calendarId() {
    return getStringSetting('06-action-booking-com-CALENDAR_ID', null);
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getStringSetting/getBooleanSetting/getNumberSetting/
// getListSetting — all defined in the sibling src/01-setup.js). Under
// GAS's shared global scope these are ALREADY visible here by bare name
// (01-setup.js loads alphabetically first) — no action needed, and this
// `if` block never executes there. Under Node, each `require()`d file is
// its own isolated module with its own scope, so the bare references
// inside BOOKING_ACTION_CONFIG's getters above would otherwise throw
// ReferenceError. Same `globalThis` bridge technique already established
// by src/06-action-booking-com-management.js's own BOOKING_ACTION_CONFIG
// bridge (from the 260724-lqi config-split refactor) and mirrored by the
// sibling ICS config file's equivalent bridge for these same helpers.
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getStringSetting = settingsHelpers.getStringSetting;
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getNumberSetting = settingsHelpers.getNumberSetting;
  globalThis.getListSetting = settingsHelpers.getListSetting;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { BOOKING_ACTION_CONFIG: BOOKING_ACTION_CONFIG };
}

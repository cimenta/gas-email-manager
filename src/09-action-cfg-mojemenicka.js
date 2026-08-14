/**
 * MOJEMENICKA_ACTION_CONFIG — the config block for MOJEMENICKA_ACTION
 * (defined in the sibling src/09-action-mojemenicka.js). Same
 * "*-action-cfg-*.js" split pattern already established for every other
 * action — see the ticketing-portals config file's own class-level JSDoc
 * for the full load-order/getter rationale (an ES6 getter is not evaluated
 * at object-construction time, only when something actually reads the
 * property, which is what makes this split safe regardless of which of the
 * two sibling files loads first alphabetically). Every field below is an
 * ES6 GETTER, PropertiesService-backed with a code-default fallback, per
 * the Script Properties live-config override feature (quick-260726-spr) —
 * see src/01-setup.js's class-level JSDoc for the full mechanism.
 *
 * Phase 4 (04-01) NEW ACTION: replies in-thread with a freshly-fetched menu
 * to an allow-listed sender whose SUBJECT carries any one of the configured
 * trigger strings (D-08, amending 04-01's single-string D-01). `allowedSenders`
 * and `triggerStrings` both defaulting to `[]` are FAIL-CLOSED BY DESIGN —
 * with either left unconfigured, MOJEMENICKA_ACTION.appliesTo can never
 * return true, so an unconfigured install must never reply to anyone. This
 * is the deliberate opposite of ICS_ACTION_CONFIG.importOnlyFrom's
 * empty-means-everyone rule, because this action — unlike every other
 * action in this codebase — sends outbound mail to a third party, not just
 * internal calendar/labeling side effects. `weeklyTriggerStrings` (D-09) is
 * a subordinate list, only consulted once a message has already matched
 * `triggerStrings` above — see its own getter comment below.
 */
const MOJEMENICKA_ACTION_CONFIG = {
  // Cross-cutting per-action enable/disable toggle, same convention as
  // every other action's config. Script Property override:
  // 09-action-mojemenicka-ENABLED (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('09-action-mojemenicka-ENABLED', true);
  },

  // Notify the script owner when this action throws. Script Property
  // override: 09-action-mojemenicka-NOTIFY_ON_FAILURE (boolean). Per D-05,
  // this only fires for genuinely unexpected errors (e.g. the unset-URL
  // misconfiguration guard, or the reply call itself throwing) — the
  // graceful fetch-failure fallback path (Task 2) deliberately does NOT
  // throw, so it never reaches this notification.
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('09-action-mojemenicka-NOTIFY_ON_FAILURE', true);
  },

  // The BARE base URL MOJEMENICKA_ACTION.run fetches via UrlFetchApp — no
  // query string of its own (D-10). The action appends its own
  // `?format=html&range=today` or `?format=html&range=week` (see
  // buildMojemenickaMenuRequestUrl in the sibling action file); a
  // configured value that already carries a query string is a
  // misconfiguration and assertMojemenickaMenuUrlIsFetchable throws. The
  // response body is still used directly, unmodified, as the reply's
  // htmlBody (D-02). No safe code default exists (there is no universal
  // "menu" URL), so this defaults to null; run() throws a clear,
  // owner-notified error naming this setting's key when it is left unset.
  // Script Property override: 09-action-mojemenicka-MENU_URL (string).
  // Script Properties value example: https://example-restaurant.cz/menu
  get menuUrl() {
    return getStringSetting('09-action-mojemenicka-MENU_URL', null);
  },

  // Allow-list of sender addresses (exact match, case-insensitive on both
  // sides via mojemenickaExtractEmailAddress) permitted to trigger a menu
  // reply (D-01). FAIL-CLOSED: an empty list means NO sender may trigger
  // this action — see this file's class-level JSDoc for why this diverges
  // from ICS_ACTION_CONFIG.importOnlyFrom's empty-means-everyone default.
  // Script Property override: 09-action-mojemenicka-ALLOWED_SENDERS (list).
  // Script Properties value example: jana@example.com, petr@example.com
  get allowedSenders() {
    return getListSetting('09-action-mojemenicka-ALLOWED_SENDERS', []);
  },

  // List of substrings MOJEMENICKA_ACTION looks for in a message's SUBJECT
  // line only (never the body) to decide whether it is a menu request
  // (D-08, amending 04-01's single-string D-01): the message matches when
  // the subject contains ANY ONE of these entries, matched case-
  // insensitively via subjectContainsAnyMojemenickaString. FAIL-CLOSED: an
  // empty or unset list never matches any subject — same rationale as
  // allowedSenders above, preserved unchanged from the retired singular
  // setting's fail-closed behavior. Script Property override:
  // 09-action-mojemenicka-TRIGGER_STRINGS (list).
  // Script Properties value example: menu, jidelnicek
  get triggerStrings() {
    return getListSetting('09-action-mojemenicka-TRIGGER_STRINGS', []);
  },

  // List of substrings that select the WEEK range instead of the default
  // TODAY range (D-09) — same case-insensitive substring matching as
  // triggerStrings above, via the same subjectContainsAnyMojemenickaString
  // matcher. SUBORDINATE to triggerStrings: this list is only ever
  // consulted for a message that has ALREADY matched triggerStrings above
  // — a weekly indicator present in a subject that carries no trigger
  // string never makes the action apply on its own. An empty or unset list
  // means every matched request uses the today range. Script Property
  // override: 09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS (list).
  // Script Properties value example: tyden, na tyden, weekly
  get weeklyTriggerStrings() {
    return getListSetting('09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS', []);
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getStringSetting/getBooleanSetting/getListSetting) — all
// defined in the sibling src/01-setup.js. Under GAS's shared global scope
// these are ALREADY visible here by bare name (01-setup.js loads
// alphabetically first) — no action needed, and this `if` block never
// executes there. Under Node, each `require()`d file is its own isolated
// module with its own scope, so the bare references inside
// MOJEMENICKA_ACTION_CONFIG's getters above would otherwise throw
// ReferenceError. Same `globalThis` bridge technique already established
// by every other action config file's equivalent bridge (see
// src/07-action-cfg-ticketing-portals.js).
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getStringSetting = settingsHelpers.getStringSetting;
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getListSetting = settingsHelpers.getListSetting;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { MOJEMENICKA_ACTION_CONFIG: MOJEMENICKA_ACTION_CONFIG };
}

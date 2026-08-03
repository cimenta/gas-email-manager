/**
 * ICS_ACTION_CONFIG — the config block for ICS_CALENDAR_ACTION (defined in
 * the sibling src/05-action-ics-import.js). Split out into its own file per
 * owner-directed refactor: each action's tunable settings live in a
 * sibling "{n}-action-cfg-{name}.js" file, separate from the action's
 * logic. GAS's shared global scope means this constant is still visible by
 * bare name inside 05-action-ics-import.js — no require/import needed for
 * the GAS runtime path.
 *
 * LOAD-ORDER SAFETY: ICS_CALENDAR_ACTION's own `config` property is an ES6
 * GETTER (`get config() { return ICS_ACTION_CONFIG; }`), not a plain
 * literal reference — a getter is NOT evaluated when the action's object
 * literal is constructed, only when something actually reads
 * `action.config` (dispatchActions' enable check, the failure-notification
 * builder), which happens inside function bodies long after every project
 * file has finished loading. This means this file and
 * 05-action-ics-import.js can load in EITHER alphabetical order safely —
 * "05-action-cfg-ics-import.js" happens to sort before
 * "05-action-ics-import.js" ('c' < 'i'), so ICS_ACTION_CONFIG is actually
 * defined first here regardless, but the getter means that ordering isn't
 * load-bearing the way a plain literal property would have required.
 *
 * SCRIPT PROPERTIES LIVE-CONFIG OVERRIDE (owner-requested, quick-260726-spr):
 * every field below is now an ES6 GETTER (same load-order-safety reasoning
 * as `config` above, applied one level deeper — to the individual fields),
 * resolving through PropertiesService via the shared get*Setting/
 * getStringSetting/etc. accessor helpers defined in src/01-setup.js, with
 * the CURRENT literal value preserved as each getter's code-default
 * fallback. See src/01-setup.js's own class-level JSDoc ("SCRIPT
 * PROPERTIES LIVE-CONFIG OVERRIDE") for the full mechanism and the
 * approved Script Properties key list. Every existing reader of these
 * fields (findIcsAttachments, resolveIcsCalendarId's caller, etc.) needs
 * ZERO changes.
 */
const ICS_ACTION_CONFIG = {
  // Cross-cutting per-action enable/disable toggle (QT-260724-lqi-TOGGLE).
  // A no-op default: dispatchActions only skips an action when this is
  // STRICTLY false, so leaving it true preserves this action's existing
  // behavior exactly while documenting the flag's presence on this action.
  // Script Property override: 05-action-ics-ENABLED (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('05-action-ics-ENABLED', true);
  },

  // Notify the script owner when this action throws.
  // Script Property override: 05-action-ics-NOTIFY_ON_FAILURE (boolean).
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('05-action-ics-NOTIFY_ON_FAILURE', true);
  },

  // Sender allow-list: array of sender email addresses (case-insensitive;
  // a bare address or a full 'Name <email>' entry both work — comparison
  // normalizes each side via extractEmailAddress). An empty array (the
  // default) imports .ics attachments from any sender, matching original
  // behavior exactly. Script Property override: 05-action-ics-IMPORT_ONLY_FROM
  // (list, comma-separated).
  // Script Properties value example: alice@example.com, bob@example.com
  get importOnlyFrom() {
    return getListSetting('05-action-ics-IMPORT_ONLY_FROM', []);
  },

  // Sender EXCLUDE-list (quick-260803-us3, D-03): the inverse hand-off gate
  // to importOnlyFrom above. An empty array (the default) excludes nobody,
  // so behavior is UNCHANGED for every existing user. This is the switch
  // that lets another action (e.g. TRANSPORT_TICKETS_ACTION) own a sender
  // that happens to send .ics attachments this action would otherwise also
  // claim — see isExcludedSender (src/05-action-ics-import.js) for the
  // matching logic (mirrors isAllowedSender's shape and its
  // case-insensitive extractEmailAddress comparison exactly) and its two
  // wiring points (findIcsAttachments, getIcsAttachmentTextsByMessage).
  // Never hardcode a real sender into this action's shipped default — the
  // owner sets the real value out-of-band via Script Properties.
  // Script Property override: 05-action-ics-EXCLUDE_FROM (list,
  // comma-separated).
  // Script Properties value example: jizdenky@regiojet.cz
  get excludeFrom() {
    return getListSetting('05-action-ics-EXCLUDE_FROM', []);
  },

  // MULTI-CALENDAR ROUTING (owner-requested): CONFIG.calendarId
  // (src/01-setup.js) is the DEFAULT/fallback calendar for every action.
  // This field, when truthy, overrides that default for EVERY ICS import
  // regardless of sender, UNLESS a more specific calendarIdBySender entry
  // (below) also matches that message's sender. null (the default)
  // preserves original behavior exactly — every import targets
  // CONFIG.calendarId. See resolveIcsCalendarId in the sibling
  // src/05-action-ics-import.js for the full 3-tier resolution this field
  // participates in. Script Property override: 05-action-ics-CALENDAR_ID
  // (string).
  // Script Properties value example: abc123@group.calendar.google.com
  get calendarId() {
    return getStringSetting('05-action-ics-CALENDAR_ID', null);
  },

  // MULTI-CALENDAR ROUTING (owner-requested): per-sender calendar overrides,
  // each entry shaped { from: '<sender email>', calendarId: '<calendar id>' }.
  // Resolved PER MESSAGE (not per-thread/per-run) against that message's
  // "From" header via resolveIcsCalendarId — case-insensitive, same
  // comparison convention as isAllowedSender/extractEmailAddress. Takes
  // priority over BOTH this file's own calendarId above and the global
  // CONFIG.calendarId default. If multiple entries match the same sender,
  // the FIRST match in list order wins (same "list, first match wins"
  // convention already used by importOnlyFrom-adjacent matching elsewhere
  // in this codebase). Independent of importOnlyFrom — that is a separate
  // gate for whether a message is processed at all; this map is only ever
  // consulted for a message that already passed that gate. Empty (the
  // default) means no per-sender routing. Script Property override:
  // 05-action-ics-CALENDAR_ID_BY_SENDER (json — array of
  // {from, calendarId} string-pair objects).
  //
  // Script Properties value example (MUST be valid JSON — double-quoted
  // keys AND double-quoted string values, unlike a JS object literal):
  // [{"from":"boss@work.com","calendarId":"abc123@group.calendar.google.com"}]
  //
  // REAL LIVE MISTAKE (owner-hit): pasting a JS object-literal-style value
  // instead — e.g. [{from: 'boss@work.com', calendarId: 'abc123...'}], with
  // unquoted keys and/or single-quoted strings — is valid JAVASCRIPT (which
  // is why it looks correct when copied from a .gs file) but INVALID JSON.
  // JSON.parse rejects it, so this setting SILENTLY falls back to its code
  // default ([], no per-sender routing) with NO visible error — every ICS
  // import then lands on the action-level/global default calendar instead,
  // with nothing in the logs to explain why. Run checkScriptPropertiesSyntax()
  // after editing this property to catch exactly this class of mistake
  // before it causes a silent misroute.
  get calendarIdBySender() {
    return getJsonSetting('05-action-ics-CALENDAR_ID_BY_SENDER', []);
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getStringSetting/getBooleanSetting/getListSetting/
// getJsonSetting — all defined in the sibling src/01-setup.js). Under
// GAS's shared global scope these are ALREADY visible here by bare name
// (01-setup.js loads alphabetically first) — no action needed, and this
// `if` block never executes there. Under Node, each `require()`d file is
// its own isolated module with its own scope, so the bare references
// inside ICS_ACTION_CONFIG's getters above would otherwise throw
// ReferenceError. Same `globalThis` bridge technique (not a redeclared
// `const`/`let`/`var`, which would collide with 01-setup.js's own
// top-level declarations under GAS's concatenated scope) already
// established by src/05-action-ics-import.js's ICS_ACTION_CONFIG bridge
// (from the 260724-lqi config-split refactor) and mirrored by the sibling
// booking config file's equivalent bridge for these same helpers.
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getStringSetting = settingsHelpers.getStringSetting;
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getListSetting = settingsHelpers.getListSetting;
  globalThis.getJsonSetting = settingsHelpers.getJsonSetting;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { ICS_ACTION_CONFIG: ICS_ACTION_CONFIG };
}

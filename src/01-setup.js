/**
 * CONFIG — single source of truth for this project's runtime configuration.
 *
 * Google Apps Script shares one global namespace across all files in the
 * project, so every other file reads this `CONFIG` constant directly — no
 * import/require needed or available in the GAS runtime.
 *
 * IMPORTANT: `calendarId` below is INTENTIONALLY left as a placeholder in
 * this git-tracked file — the owner does not want their real Calendar ID
 * committed to source control. The live value is set out-of-band, directly
 * in the Apps Script browser editor, after `clasp push`. Do NOT replace the
 * placeholder here and do NOT push again after the owner's manual edit, or
 * the next `clasp push` will overwrite their live value with this
 * placeholder. `labelName` and `failedLabelName` are free-text — edit them
 * to whatever Gmail label names you prefer; there is no "correct" default.
 *
 * MULTI-CALENDAR ROUTING (owner-requested, quick-260726-mcr): `calendarId`
 * is now a DEFAULT/FALLBACK, not the sole target every action writes to.
 * Each action can override it: the ICS action via
 * ICS_ACTION_CONFIG.calendarId (action-wide) and/or
 * ICS_ACTION_CONFIG.calendarIdBySender (per-sender, see
 * resolveIcsCalendarId in src/05-action-ics-import.js); the booking action
 * via BOOKING_ACTION_CONFIG.calendarId (action-wide, see
 * resolveBookingCalendarId in src/06-action-booking-com-management.js).
 * With every override left at its shipped default (null / []), every
 * action's Calendar API calls still target this `calendarId` exactly as
 * before this feature existed.
 *
 * SCRIPT PROPERTIES LIVE-CONFIG OVERRIDE (owner-requested, quick-260726-spr):
 * every field below (and on ICS_ACTION_CONFIG / BOOKING_ACTION_CONFIG in
 * their own sibling files) is now an ES6 GETTER instead of a plain literal.
 * Each getter resolves through PropertiesService.getScriptProperties() FIRST
 * — a GAS-native key/value store that `clasp push` NEVER touches — falling
 * back to the literal value documented in that field's own comment (its
 * "code default") only when no Script Property is set for that key. This
 * solves the actual problem that motivated this feature: `clasp push`
 * overwrites every src/*.js file wholesale, so any live setting the owner
 * previously edited directly in THIS file (e.g. a real calendarId) was
 * silently clobbered back to its git-tracked placeholder on the next push.
 * Now the live value lives in Script Properties instead, and this file's
 * literal values serve only as fallback defaults AND as living
 * documentation of what each setting does and its out-of-the-box value —
 * they are never themselves clobbered, because nothing here ever WRITES a
 * real value into this file.
 *
 * Because this reuses the SAME getter pattern already proven in this
 * codebase for `action.config` (see the 260724-lqi config-split refactor —
 * an ES6 getter is not evaluated at object-construction time, only when
 * something actually reads the property, which happens deep inside
 * function bodies long after every project file has loaded), EVERY
 * EXISTING CALL SITE that reads `CONFIG.calendarId`, `action.config.enabled`,
 * `ICS_ACTION_CONFIG.importOnlyFrom`, etc. anywhere else in the codebase
 * needs ZERO changes — they keep reading what looks like a plain property,
 * unaware it is now backed by a getter.
 *
 * The full mechanism (typed accessor helpers, pure parse/format functions,
 * SETTINGS_REGISTRY, rebuildScriptProperties, checkScriptPropertiesSyntax)
 * lives below, all in this file, since it's shared via GAS's global scope
 * by the sibling config files' own getters (see each sibling file's own
 * Node/GAS bridge for how they reach these helpers under Node, where there
 * is no shared global scope).
 *
 * CRITICAL GAS/NODE DUALITY WARNING: `PropertiesService` does not exist
 * under Node — referencing it throws `ReferenceError`, it does NOT
 * gracefully evaluate to `undefined` the way a merely-unset variable would.
 * `readScriptProperty` (below) is the SINGLE guard point every one of these
 * getters flows through; see its own doc comment for why centralizing the
 * guard there, rather than repeating it in every wrapper, is what keeps
 * this safe. A dedicated test (test/script-properties.test.js) proves every
 * field on every one of these three config objects still resolves to its
 * exact pre-refactor default under Node with PropertiesService absent —
 * this is the single most important regression check introduced by this
 * feature, since a miss here would silently break the entire pre-existing
 * test suite's ability to run at all.
 */
const CONFIG = {
  // DEFAULT/FALLBACK calendar for every action — see MULTI-CALENDAR
  // ROUTING above; each action can override this per-action and (ICS
  // only) per-sender. Placeholder — intentionally NOT replaced in git. The
  // real Calendar ID is set as a Script Property (01-setup-CALENDAR_ID),
  // never committed here. Find it via Calendar UI (Settings and sharing ->
  // Integrate calendar) or a one-off CalendarApp.getAllCalendars() listing
  // script.
  // Script Properties value example: abc123@group.calendar.google.com
  get calendarId() {
    return getStringSetting('01-setup-CALENDAR_ID', 'REPLACE_WITH_CALENDAR_ID');
  },

  // Gmail label applied once an email has been processed (success or fail).
  // Flat (non-nested) label name — owner manages label naming/renaming
  // directly in Gmail; no "GAS Email Manager" parent required.
  // Script Properties value example: II
  get labelName() {
    return getStringSetting('01-setup-LABEL_NAME', 'II');
  },

  // Gmail label applied in addition to labelName when processing fails.
  // Script Properties value example: FF
  get failedLabelName() {
    return getStringSetting('01-setup-FAILED_LABEL_NAME', 'FF');
  },

  // How many days back to search for unlabeled emails.
  // Script Properties value example: 2
  get daysBack() {
    return getNumberSetting('01-setup-DAYS_BACK', 2);
  },

  // Whether setup() should install the time-driven trigger. Toggle this and
  // re-run setup() to turn the schedule on/off.
  // Script Properties value example: true
  get installTrigger() {
    return getBooleanSetting('01-setup-INSTALL_TRIGGER', true);
  },

  // Interval, in minutes, for the time-driven trigger. Must be one of the
  // GAS-supported everyMinutes values: 1, 5, 10, 15, 30.
  // Script Properties value example: 5
  get triggerIntervalMinutes() {
    return getNumberSetting('01-setup-TRIGGER_INTERVAL_MINUTES', 5);
  },

  // v0.6.0 NEW GLOBAL SETTING (ticketing-portals action, quick-260731-tix):
  // the PERMANENT Drive folder NAME (looked up by name via
  // DriveApp.getFoldersByName, or created via DriveApp.createFolder if none
  // exists yet — NOT a folder ID) where a ticket PDF is moved when a
  // ticketing portal's insertPdfIntoEvent is true. Deliberately a CROSS-
  // CUTTING property on CONFIG (not on TICKETING_PORTALS_ACTION_CONFIG),
  // since it names a single shared destination folder used across every
  // configured portal, not a per-portal setting. See
  // src/07-action-ticketing-portals.js's class-level JSDoc for the full
  // Drive/OCR pipeline this participates in (the separate, fixed-name,
  // fully-auto-managed TEMP folder every ticket PDF passes through first is
  // NOT user-configurable and has no equivalent CONFIG field).
  // Script Properties value example: GAS Email Manager - Ticket Attachments
  get ticketAttachmentDriveFolderName() {
    return getStringSetting('01-setup-TICKET_ATTACHMENT_DRIVE_FOLDER_NAME', 'GAS Email Manager - Ticket Attachments');
  },
};

// FUNCTION-PICKER ORDERING (owner-hit real usability issue): the Apps
// Script editor's function-picker dropdown lists functions in their
// SOURCE POSITION within each file (not alphabetically project-wide, nor
// by any other ordering) — so the 4 manually-run entry points this file
// defines (rebuildScriptProperties, checkScriptPropertiesSyntax, setup,
// reconcileProcessEmailsTrigger) are deliberately placed here, immediately
// after CONFIG, BEFORE every internal helper function below (the typed
// accessors, pure parse/format functions, SETTINGS_REGISTRY, etc.) — so
// they appear at the TOP of that dropdown, ahead of anything a user would
// never want to run directly. Pure source-position reorder: JS function
// declarations are hoisted within their scope, and none of these 4
// functions reference anything at PARSE time (only inside their own
// bodies, evaluated lazily when called), so this placement is safe
// regardless of where CONFIG, SETTINGS_REGISTRY, or the internal helpers
// end up relative to them — same reasoning already established in this
// codebase for the getter-based config split. src/02-main.js needs no
// equivalent reorder: processEmails() is already the first function
// declared in that file.

/**
 * rebuildScriptProperties — real, top-level, MANUALLY-INVOKED function
 * (run directly from the Apps Script editor's function picker, same
 * category as setup()). Builds the full 20-key map via
 * buildRebuiltPropertiesMap(SETTINGS_REGISTRY) (see that function's own
 * doc comment for the empty-string/sentinel handling — every one of the
 * 20 keys is ALWAYS included, never omitted), then writes it in ONE
 * atomic call:
 * `PropertiesService.getScriptProperties().setProperties(rebuiltMap, true)`
 * — the `true` (deleteAllOthers) means the property store ends up holding
 * EXACTLY the canonical 20 keys, nothing stale left over; no separate
 * deleteAllProperties() call is needed.
 *
 * This ONE function naturally serves BOTH purposes: first-time migration
 * (nothing stored yet -> every entry's read() falls through to its code
 * default, so the store gets seeded entirely from today's code defaults,
 * with the sentinel standing in for any that are empty) AND the "add a
 * newly-introduced setting / fix key ordering" tool for later (every
 * ALREADY-SET value passes through unchanged via read(), only
 * missing/newly-added keys get freshly seeded) — these are the same
 * operation, not two different ones.
 *
 * GAS-only (PropertiesService) — not unit-tested directly; the map it
 * writes IS fully covered via buildRebuiltPropertiesMap's own pure-function
 * tests (which in turn exercise formatListSettingValue/formatJsonSettingValue/
 * formatSettingValueForStorage).
 */
function rebuildScriptProperties() {
  const rebuiltMap = buildRebuiltPropertiesMap(SETTINGS_REGISTRY);

  PropertiesService.getScriptProperties().setProperties(rebuiltMap, true);

  console.log(
    'rebuildScriptProperties: wrote ' + Object.keys(rebuiltMap).length + ' Script Properties (deleteAllOthers=true).'
  );
}

/**
 * checkScriptPropertiesSyntax — real, top-level, MANUALLY-INVOKED
 * READ-ONLY diagnostic (run directly from the Apps Script editor's
 * function picker). Walks SETTINGS_REGISTRY; for each entry, reads the
 * RAW stored value DIRECTLY
 * (`PropertiesService.getScriptProperties().getProperty(entry.key)`) —
 * deliberately NOT through the get*Setting getter/fallback chain, since
 * this function specifically wants to validate what is actually typed
 * there, not the already-recovered effective value. An unset key OR a key
 * holding exactly SETTING_DEFAULT_SENTINEL (the literal text
 * rebuildScriptProperties writes in place of an empty value — see its own
 * doc comment) are BOTH fine, not an error (informational only: "using
 * code default") — describeSettingSyntaxIssue itself carries this
 * sentinel-awareness (bypassing readScriptProperty's own translation is
 * exactly why this function needs its own check, rather than inheriting
 * it for free). A set-but-genuinely-malformed value (bad boolean token,
 * non-numeric, invalid JSON, wrong JSON shape — see
 * describeSettingSyntaxIssue) is collected into an issues list with the
 * exact key name and a specific description of what's wrong. Prints a
 * clear pass/fail summary via console.log — no email, this is a
 * manually-invoked tool, not unattended processEmails() failure handling,
 * and it NEVER throws (a malformed value is reported, not treated as
 * fatal — checkScriptPropertiesSyntax is a diagnostic, the runtime
 * get*Setting wrappers are what guarantee processEmails() itself never
 * crashes over a bad stored value).
 *
 * GAS-only (PropertiesService) — not unit-tested directly; the malformed-
 * value detection logic it depends on (describeSettingSyntaxIssue, and
 * transitively the pure parse functions) IS fully unit-tested.
 *
 * Returns the issues array (empty on a clean pass) so a caller could
 * inspect it programmatically if ever needed; the primary interface is the
 * console.log summary for a human running this from the editor.
 */
function checkScriptPropertiesSyntax() {
  const issues = [];

  SETTINGS_REGISTRY.forEach(function (entry) {
    const raw = PropertiesService.getScriptProperties().getProperty(entry.key);
    const issue = describeSettingSyntaxIssue(entry.type, raw, entry.jsonShapeValidator, entry.jsonShapeDescription);
    if (issue) {
      issues.push({ key: entry.key, issue: issue });
    }
  });

  if (issues.length === 0) {
    console.log('checkScriptPropertiesSyntax: PASS — all set Script Properties are syntactically valid (unset keys use their code default).');
  } else {
    console.log('checkScriptPropertiesSyntax: FAIL — ' + issues.length + ' issue(s) found:');
    issues.forEach(function (item) {
      console.log('  - ' + item.key + ': ' + item.issue);
    });
  }

  return issues;
}

/**
 * setup — one-time, manually-invoked function that establishes this
 * project's persistent Apps Script artifacts: the processed Gmail label
 * and the time-driven trigger targeting processEmails(). Run this once
 * from the Apps Script editor after deploying; re-running it is safe
 * and idempotent, and is also how you toggle the schedule on/off or pick
 * up a changed CONFIG.triggerIntervalMinutes.
 *
 * This function must never read, label, archive, or otherwise modify any
 * Gmail message, and must never delete or rename any pre-existing Gmail
 * label. It only creates the configured processed and failed labels when
 * they don't already exist (both labels are provisioned eagerly here, not
 * lazily on first use). Trigger reconciliation is scoped strictly to
 * triggers whose handler is 'processEmails' — it must never touch triggers
 * belonging to any other handler.
 */

// GAS ScriptApp.newTrigger(...).timeBased().everyMinutes(n) only accepts
// these interval values; anything else fails cryptically at create().
const SUPPORTED_TRIGGER_INTERVALS_MINUTES = [1, 5, 10, 15, 30];

function setup() {
  getOrCreateLabel(CONFIG.labelName);
  console.log('Processed label ready: ' + CONFIG.labelName);

  getOrCreateLabel(CONFIG.failedLabelName);
  console.log('Failed label ready: ' + CONFIG.failedLabelName);

  reconcileProcessEmailsTrigger();
}

/**
 * Reconciles the time-driven trigger targeting processEmails() against
 * CONFIG.installTrigger / CONFIG.triggerIntervalMinutes. Operates only on
 * triggers whose handler function is 'processEmails' — never touches any
 * other handler's triggers.
 */
function reconcileProcessEmailsTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'processEmails';
  });

  if (!CONFIG.installTrigger) {
    existingTriggers.forEach(function (t) {
      ScriptApp.deleteTrigger(t);
    });
    console.log(
      'installTrigger is false: removed ' + existingTriggers.length + ' processEmails trigger(s)'
    );
    return;
  }

  if (SUPPORTED_TRIGGER_INTERVALS_MINUTES.indexOf(CONFIG.triggerIntervalMinutes) === -1) {
    throw new Error(
      'Unsupported CONFIG.triggerIntervalMinutes: ' +
        CONFIG.triggerIntervalMinutes +
        '. Allowed values are ' +
        SUPPORTED_TRIGGER_INTERVALS_MINUTES.join(', ') +
        '.'
    );
  }

  // Reconcile to exactly one trigger at the current configured interval:
  // remove every existing processEmails trigger, then create one fresh.
  // This guarantees no duplicates AND picks up interval changes, since a
  // Trigger object exposes no public getter for its configured interval.
  existingTriggers.forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .everyMinutes(CONFIG.triggerIntervalMinutes)
    .create();

  console.log(
    'Removed ' +
      existingTriggers.length +
      ' stale processEmails trigger(s); installed 1 at ' +
      CONFIG.triggerIntervalMinutes +
      ' minute interval'
  );
}

/**
 * SETTING_DEFAULT_SENTINEL — the literal string written to Script
 * Properties in place of an empty value (see buildRebuiltPropertiesMap
 * below for WHY: Script Properties rejects an empty string outright).
 * Recognized as equivalent to "absent"/"use code default" by every read
 * path this feature has: readScriptProperty (and therefore every one of
 * the 5 get*Setting wrappers, for free) and describeSettingSyntaxIssue
 * (checkScriptPropertiesSyntax's pure core). Exactly one literal value,
 * defined once here, referenced everywhere else — never re-typed as a
 * bare string literal elsewhere in this feature.
 *
 * TO REVERT ANY SETTING TO ITS CODE DEFAULT: either edit its Script
 * Properties row to hold exactly this literal text ("DefaultValue"), or
 * delete that row entirely — both are treated identically by every read
 * path. The code default shown for each setting is documented right next
 * to its getter, in this file (CONFIG, above) or in the matching
 * ICS_ACTION_CONFIG/BOOKING_ACTION_CONFIG getter's own comment in their
 * respective sibling config files.
 */
const SETTING_DEFAULT_SENTINEL = 'DefaultValue';

/**
 * isSettingDefaultSentinel — pure: true iff `raw` is EXACTLY (case- and
 * whitespace-sensitive — this is a literal marker string, not a
 * user-facing token like the boolean type's "true"/"false") the
 * SETTING_DEFAULT_SENTINEL value. No GAS globals.
 */
function isSettingDefaultSentinel(raw) {
  return raw === SETTING_DEFAULT_SENTINEL;
}

/**
 * readScriptProperty — the SINGLE GAS/Node duality guard point for the
 * entire Script Properties feature. Every one of the 5 get*Setting
 * wrappers below flows through this one function, so guarding here once is
 * sufficient — there is no other path to PropertiesService anywhere in
 * this feature (checkScriptPropertiesSyntax is the sole deliberate
 * exception: it reads RAW values directly, bypassing this wrapper on
 * purpose — see its own doc comment — but it is GAS-only and never called
 * under Node).
 *
 * `PropertiesService` does not exist under Node: referencing the bare
 * identifier throws `ReferenceError` immediately (unlike a merely-unset
 * variable, which evaluates to `undefined`). This mirrors this codebase's
 * established `typeof module !== 'undefined'` GAS/Node duality guard
 * pattern already used throughout every `*_CONFIG` file's Node export
 * block — same technique, applied here to the opposite direction (guarding
 * a GAS-only global instead of a Node-only one).
 *
 * Returns the raw stored string for `key`, or `null` when ANY of: (a) no
 * Script Property is set for that key, (b) PropertiesService itself is
 * unavailable (Node), or (c) the stored value IS the
 * SETTING_DEFAULT_SENTINEL (see buildRebuiltPropertiesMap) — all three
 * cases are treated identically by every `get*Setting` wrapper: "no live
 * override, use the code default." This is the ONE place the sentinel is
 * translated back to "absent" for the read side, so every wrapper gets
 * correct sentinel-as-absent behavior for free, with zero changes needed
 * to each of them individually. Pure with respect to Node-safety; still
 * calls the live GAS API when available.
 */
function readScriptProperty(key) {
  if (typeof PropertiesService === 'undefined') {
    return null;
  }
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  return isSettingDefaultSentinel(raw) ? null : raw;
}

/**
 * getStringSetting — string-typed accessor: passthrough of the raw stored
 * value when set, else `codeDefault` unchanged. `codeDefault` itself need
 * not be a string (e.g. ICS_ACTION_CONFIG.calendarId's code default is
 * `null`) — this function never re-types its default argument, only the
 * live stored value (which, once actually set, is always a real string).
 */
function getStringSetting(key, codeDefault) {
  const raw = readScriptProperty(key);
  return raw === null || raw === undefined ? codeDefault : raw;
}

/**
 * getBooleanSetting — boolean-typed accessor. Delegates to the pure
 * parseBooleanSettingValue (see its own doc comment for the exact token
 * rules) so its parsing logic is independently unit-testable without any
 * GAS global.
 */
function getBooleanSetting(key, codeDefault) {
  return parseBooleanSettingValue(readScriptProperty(key), codeDefault);
}

/**
 * getNumberSetting — number-typed accessor. Delegates to the pure
 * parseNumberSettingValue.
 */
function getNumberSetting(key, codeDefault) {
  return parseNumberSettingValue(readScriptProperty(key), codeDefault);
}

/**
 * getListSetting — list-typed accessor (comma-separated string storage
 * form). Delegates to the pure parseListSettingValue for the SET case;
 * unlike the other wrappers, an explicitly-set-but-empty/whitespace-only
 * stored value is a valid, well-formed list (parses to `[]`, not a parse
 * failure) — parseListSettingValue itself embodies that rule, so this
 * wrapper only needs to special-case the genuinely UNSET case (raw is
 * null/undefined) as "use codeDefault", exactly like every other wrapper.
 */
function getListSetting(key, codeDefault) {
  const raw = readScriptProperty(key);
  return raw === null || raw === undefined ? codeDefault : parseListSettingValue(raw);
}

/**
 * getJsonSetting — JSON-typed accessor. Delegates to the pure
 * parseJsonSettingValue, which also validates the parsed shape (see its
 * own doc comment) — a syntactically valid but wrong-shaped JSON value is
 * treated the same as a parse failure: silent fallback to `codeDefault`.
 *
 * `isValidShape` (optional, quick-260731-tix generalization): the shape
 * validator to use for THIS setting — see parseJsonSettingValue's own doc
 * comment for why this parameter exists and its backward-compatible
 * default when omitted.
 */
function getJsonSetting(key, codeDefault, isValidShape) {
  return parseJsonSettingValue(readScriptProperty(key), codeDefault, isValidShape);
}

/**
 * parseBooleanSettingValue — pure. `raw` is the value exactly as returned
 * by readScriptProperty (a real string when set, `null`/`undefined` when
 * unset). Unset -> `codeDefault`. Set: trimmed and lowercased, then
 * matched case-insensitively against the literal tokens `"true"`/`"false"`
 * — any other token (including an empty/whitespace-only string, which is
 * NOT treated as "unset" here, unlike the list type) is a malformed
 * boolean and falls back to `codeDefault`, NEVER throws. No GAS globals.
 */
function parseBooleanSettingValue(raw, codeDefault) {
  if (raw === null || raw === undefined) {
    return codeDefault;
  }

  const normalized = String(raw).trim().toLowerCase();
  if (normalized === 'true') {
    return true;
  }
  if (normalized === 'false') {
    return false;
  }

  return codeDefault;
}

/**
 * parseNumberSettingValue — pure. Unset (`null`/`undefined`) ->
 * `codeDefault`. Set: parsed via the built-in `Number(...)`; a genuinely
 * unparseable value produces `NaN`, which falls back to `codeDefault`
 * (never throws). DOCUMENTED JS QUIRK, deliberately preserved rather than
 * special-cased: `Number('')` (and `Number('   ')`) is `0`, NOT `NaN` — an
 * explicitly-set empty/whitespace-only stored value therefore resolves to
 * the number `0`, not `codeDefault`. This is exactly what the approved
 * spec's literal wording ("parsed via Number(...), invalid -> NaN -> falls
 * back") produces; a dedicated test locks this in explicitly so it reads
 * as an intentional decision, not an oversight, if ever revisited. No GAS
 * globals.
 */
function parseNumberSettingValue(raw, codeDefault) {
  if (raw === null || raw === undefined) {
    return codeDefault;
  }

  const parsed = Number(raw);
  return isNaN(parsed) ? codeDefault : parsed;
}

/**
 * parseListSettingValue — pure. `null`/`undefined` -> `[]` (this function
 * has no `codeDefault` parameter by design — the get*Setting wrapper
 * handles the "genuinely unset" case itself before ever calling this;
 * this function's own job is purely "turn a stored string into a list").
 * Splits on `,`, trims each entry, then filters out empty entries — this
 * is what makes an explicitly-set empty or whitespace-only stored string
 * correctly resolve to `[]` rather than `['']`: a single comma-free empty
 * string splits to `['']`, trims to `['']`, and the empty-entry filter
 * removes that lone blank entry, leaving `[]`. No GAS globals.
 */
function parseListSettingValue(raw) {
  if (raw === null || raw === undefined) {
    return [];
  }

  return String(raw)
    .split(',')
    .map(function (entry) {
      return entry.trim();
    })
    .filter(function (entry) {
      return entry !== '';
    });
}

/**
 * isValidCalendarIdBySenderShape — pure shape validator for the ICS
 * action's JSON-typed setting (`05-action-ics-CALENDAR_ID_BY_SENDER` /
 * ICS_ACTION_CONFIG.calendarIdBySender): an array where every entry is a
 * plain object with a string `from` and a string `calendarId`. Also
 * serves as parseJsonSettingValue/getJsonSetting's BACKWARD-COMPATIBLE
 * DEFAULT shape validator when no explicit one is passed (see those
 * functions' own doc comments) — this preserves every pre-existing call
 * site's exact behavior unchanged.
 *
 * GENERALIZED (quick-260731-tix): originally documented as needing
 * generalizing "if a future JSON-typed setting with a different shape is
 * ever added" — that moment arrived with the ticketing-portals action's
 * `ticketingPortals` setting (a differently-shaped JSON array; see
 * isValidTicketingPortalsShape below). Rather than hardcoding a single
 * shape check inside parseJsonSettingValue, that function now accepts an
 * explicit `isValidShape` validator per call site, with THIS function
 * remaining the default for backward compatibility. No GAS globals.
 */
function isValidCalendarIdBySenderShape(value) {
  if (!Array.isArray(value)) {
    return false;
  }

  return value.every(function (entry) {
    return (
      entry !== null &&
      typeof entry === 'object' &&
      typeof entry.from === 'string' &&
      typeof entry.calendarId === 'string'
    );
  });
}

/**
 * isValidTicketingPortalsShape — pure shape validator for the
 * ticketing-portals action's JSON-typed setting
 * (`07-action-ticketing-portals-TICKETING_PORTALS` /
 * TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals, quick-260731-tix): an
 * array where every entry is a plain object with a string
 * `identifyingEmail`, a `calendarId` that is either a string or `null`
 * (the shipped default entry's code-default value, before the owner fills
 * in a real calendar), and a boolean `insertPdfIntoEvent`. See
 * src/07-action-ticketing-portals.js's class-level JSDoc for what each
 * field controls. Not a general-purpose JSON-schema validator, same scope
 * caveat as isValidCalendarIdBySenderShape. No GAS globals.
 *
 * SHARED (quick-260803-us3): also validates the transport-tickets action's
 * `08-action-transport-tickets-TRANSPORT_SENDERS` setting
 * (TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders) — that setting's
 * entries have the IDENTICAL `{identifyingEmail, calendarId,
 * insertPdfIntoEvent}` shape, so it reuses this validator rather than
 * cloning an equivalent one.
 */
function isValidTicketingPortalsShape(value) {
  if (!Array.isArray(value)) {
    return false;
  }

  return value.every(function (entry) {
    return (
      entry !== null &&
      typeof entry === 'object' &&
      typeof entry.identifyingEmail === 'string' &&
      (entry.calendarId === null || typeof entry.calendarId === 'string') &&
      typeof entry.insertPdfIntoEvent === 'boolean'
    );
  });
}

/**
 * parseJsonSettingValue — pure. Unset (`null`/`undefined`) -> `codeDefault`.
 * Set: parsed via `JSON.parse`; a syntax error (invalid JSON) is caught and
 * falls back to `codeDefault`, never throws. A syntactically VALID JSON
 * value that does not match the expected shape is ALSO treated as a parse
 * failure and falls back to `codeDefault` — "valid JSON, wrong shape" is
 * exactly as unusable to the caller as "not JSON at all".
 *
 * `isValidShape` (optional, quick-260731-tix generalization): the shape
 * validator to apply to the parsed value for THIS setting — each JSON-typed
 * setting has its own shape (compare isValidCalendarIdBySenderShape and
 * isValidTicketingPortalsShape). Defaults to isValidCalendarIdBySenderShape
 * when omitted, preserving every pre-existing call site's exact behavior
 * unchanged (this was the only JSON-typed setting before quick-260731-tix).
 * No GAS globals.
 */
function parseJsonSettingValue(raw, codeDefault, isValidShape) {
  if (raw === null || raw === undefined) {
    return codeDefault;
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    return codeDefault;
  }

  const isValid = isValidShape || isValidCalendarIdBySenderShape;
  return isValid(parsed) ? parsed : codeDefault;
}

/**
 * formatListSettingValue — pure, the write-side inverse of
 * parseListSettingValue: joins an array into this setting type's
 * comma-and-space-separated storage form (e.g. `'a@x.com, b@y.com'`).
 * `null`/`undefined`/`[]` all format to `''` — never throws. Used by
 * rebuildScriptProperties to serialize a list-typed setting's current
 * effective value for storage. No GAS globals.
 */
function formatListSettingValue(array) {
  return (array || []).join(', ');
}

/**
 * formatJsonSettingValue — pure, the write-side inverse of
 * parseJsonSettingValue: `JSON.stringify`s a value for storage (`undefined`
 * is normalized to `null` first, since `JSON.stringify(undefined)` itself
 * returns `undefined`, not a string, which would break storage). Used by
 * rebuildScriptProperties to serialize a json-typed setting's current
 * effective value. No GAS globals.
 */
function formatJsonSettingValue(value) {
  return JSON.stringify(value === undefined ? null : value);
}

/**
 * formatSettingValueForStorage — dispatches to the correct formatter for
 * `type` (as used by SETTINGS_REGISTRY entries), producing the exact
 * string storage form for a setting's current effective value:
 * boolean -> `"true"`/`"false"`; number -> `String(n)`; list ->
 * formatListSettingValue; json -> formatJsonSettingValue; string (and any
 * other/unrecognized type, defensively) -> the value itself coerced to a
 * string, or `''` for `null`/`undefined`. Used only by
 * rebuildScriptProperties. No GAS globals.
 */
function formatSettingValueForStorage(type, value) {
  if (type === 'boolean') {
    return String(Boolean(value));
  }
  if (type === 'number') {
    return String(value);
  }
  if (type === 'list') {
    return formatListSettingValue(value);
  }
  if (type === 'json') {
    return formatJsonSettingValue(value);
  }

  return value === null || value === undefined ? '' : String(value);
}

/**
 * describeSettingSyntaxIssue — the pure core of checkScriptPropertiesSyntax
 * (below): given a setting `type` and its RAW stored value, returns `null`
 * when the value is fine (either genuinely unset, i.e. `null`/`undefined`
 * — informational, not an issue — or set and well-formed for its type), or
 * a short human-readable description of what's wrong when it's set but
 * malformed.
 *
 * DELIBERATELY reuses the EXACT SAME pure parse functions the runtime
 * getters use, via a sentinel-default technique, rather than duplicating
 * any validation logic: each type's raw value is parsed with a unique
 * sentinel object as the codeDefault; if the parse result IS that sentinel
 * (by reference, `===`), the parse fell through to its fallback path,
 * meaning the raw value was malformed for that type. This guarantees
 * checkScriptPropertiesSyntax and the runtime getters can never disagree
 * about what counts as malformed, by construction — there is no second
 * copy of "what a valid boolean/number/JSON value looks like" to drift out
 * of sync with the parsers. `string` and `list` settings are NEVER flagged
 * — any raw string is a well-formed string, and parseListSettingValue
 * never fails (a blank stored value is a valid empty list, not an error).
 *
 * This function reads the RAW value DIRECTLY (see checkScriptPropertiesSyntax's
 * own doc comment for why), so it does NOT go through readScriptProperty's
 * sentinel-to-null translation — it therefore needs its OWN sentinel check
 * here, treating a raw value of exactly SETTING_DEFAULT_SENTINEL identically
 * to a genuinely absent (null/undefined) one: informational, never an
 * issue, regardless of `type`.
 *
 * `jsonShapeValidator`/`jsonShapeDescription` (optional, quick-260731-tix
 * generalization): for a `type === 'json'` entry, the specific shape
 * validator to check against and the human-readable description to quote
 * in the issue message — see SETTINGS_REGISTRY's `json`-typed entries for
 * where these come from (checkScriptPropertiesSyntax passes
 * `entry.jsonShapeValidator`/`entry.jsonShapeDescription` through). Both
 * default to the original ICS shape/description when omitted, preserving
 * every pre-existing caller's exact behavior unchanged. No GAS globals.
 */
const SETTING_SYNTAX_CHECK_SENTINEL = { settingSyntaxCheckSentinel: true };

function describeSettingSyntaxIssue(type, raw, jsonShapeValidator, jsonShapeDescription) {
  if (raw === null || raw === undefined || isSettingDefaultSentinel(raw)) {
    return null;
  }

  if (type === 'boolean') {
    const parsed = parseBooleanSettingValue(raw, SETTING_SYNTAX_CHECK_SENTINEL);
    return parsed === SETTING_SYNTAX_CHECK_SENTINEL
      ? 'expected "true" or "false" (case-insensitive), got: ' + JSON.stringify(raw)
      : null;
  }

  if (type === 'number') {
    const parsed = parseNumberSettingValue(raw, SETTING_SYNTAX_CHECK_SENTINEL);
    return parsed === SETTING_SYNTAX_CHECK_SENTINEL
      ? 'expected a numeric value, got: ' + JSON.stringify(raw)
      : null;
  }

  if (type === 'json') {
    const validator = jsonShapeValidator || isValidCalendarIdBySenderShape;
    const description = jsonShapeDescription || 'an array of {from, calendarId} string-pair objects';
    const parsed = parseJsonSettingValue(raw, SETTING_SYNTAX_CHECK_SENTINEL, validator);
    return parsed === SETTING_SYNTAX_CHECK_SENTINEL
      ? 'expected valid JSON: ' + description + ', got: ' + JSON.stringify(raw)
      : null;
  }

  // 'string' and 'list' settings never fail to parse — any raw string is a
  // well-formed value for either type.
  return null;
}

/**
 * SETTINGS_REGISTRY — single source of truth for the 20 approved Script
 * Properties keys (exact names, types, and live-value accessors), avoiding
 * duplicating that list three times (rebuildScriptProperties,
 * checkScriptPropertiesSyntax, and the completeness test all walk this
 * SAME array). Each entry: `{ key, type, read }`, where `read()` returns
 * the setting's CURRENT EFFECTIVE value by calling straight through the
 * SAME getter chain described above (e.g. `CONFIG.calendarId`,
 * `BOOKING_ACTION_CONFIG.addToCalendar.enabled`) — already correctly
 * resolved as "existing Script Property if set, else code default", so
 * rebuildScriptProperties needs no separate hardcoded-defaults table to
 * keep in sync with the getters above.
 *
 * LOAD-ORDER NOTE: each `read` closure references CONFIG/ICS_ACTION_CONFIG/
 * BOOKING_ACTION_CONFIG by BARE NAME. Under GAS's alphabetical-by-filename
 * load order, "01-setup.js" (this file) loads BEFORE
 * "05-action-cfg-ics-import.js" and "06-action-cfg-booking-com-management.js"
 * — so at the moment THIS ARRAY is constructed, ICS_ACTION_CONFIG and
 * BOOKING_ACTION_CONFIG do not exist yet. This is safe ONLY because a
 * closure's BODY is not evaluated at construction time, only when actually
 * CALLED (`entry.read()`), which happens exclusively inside
 * rebuildScriptProperties/checkScriptPropertiesSyntax — both manually
 * invoked by the owner long after every project file has finished loading.
 * This is the exact same "resolve lazily, at the point of use" pattern
 * this codebase has relied on repeatedly (BOOKING_LANGUAGE_PACKS/
 * getBookingLabels, every `action.config` getter).
 */
const SETTINGS_REGISTRY = [
  { key: '01-setup-CALENDAR_ID', type: 'string', read: function () { return CONFIG.calendarId; } },
  { key: '01-setup-LABEL_NAME', type: 'string', read: function () { return CONFIG.labelName; } },
  { key: '01-setup-FAILED_LABEL_NAME', type: 'string', read: function () { return CONFIG.failedLabelName; } },
  { key: '01-setup-DAYS_BACK', type: 'number', read: function () { return CONFIG.daysBack; } },
  { key: '01-setup-INSTALL_TRIGGER', type: 'boolean', read: function () { return CONFIG.installTrigger; } },
  { key: '01-setup-TRIGGER_INTERVAL_MINUTES', type: 'number', read: function () { return CONFIG.triggerIntervalMinutes; } },
  // v0.6.0 (quick-260731-tix): the ticketing-portals action's shared
  // permanent-attachment-folder name — see CONFIG.ticketAttachmentDriveFolderName's
  // own doc comment above for why this lives on CONFIG, not on
  // TICKETING_PORTALS_ACTION_CONFIG.
  { key: '01-setup-TICKET_ATTACHMENT_DRIVE_FOLDER_NAME', type: 'string', read: function () { return CONFIG.ticketAttachmentDriveFolderName; } },

  { key: '05-action-ics-ENABLED', type: 'boolean', read: function () { return ICS_ACTION_CONFIG.enabled; } },
  { key: '05-action-ics-NOTIFY_ON_FAILURE', type: 'boolean', read: function () { return ICS_ACTION_CONFIG.notifyOnFailure; } },
  { key: '05-action-ics-IMPORT_ONLY_FROM', type: 'list', read: function () { return ICS_ACTION_CONFIG.importOnlyFrom; } },
  // quick-260803-us3 (D-03): the inverse hand-off gate — see
  // ICS_ACTION_CONFIG.excludeFrom's own doc comment in
  // src/05-action-cfg-ics-import.js.
  { key: '05-action-ics-EXCLUDE_FROM', type: 'list', read: function () { return ICS_ACTION_CONFIG.excludeFrom; } },
  { key: '05-action-ics-CALENDAR_ID', type: 'string', read: function () { return ICS_ACTION_CONFIG.calendarId; } },
  {
    key: '05-action-ics-CALENDAR_ID_BY_SENDER',
    type: 'json',
    jsonShapeValidator: isValidCalendarIdBySenderShape,
    jsonShapeDescription: 'an array of {from, calendarId} string-pair objects',
    read: function () { return ICS_ACTION_CONFIG.calendarIdBySender; },
  },

  { key: '06-action-booking-com-ENABLED', type: 'boolean', read: function () { return BOOKING_ACTION_CONFIG.enabled; } },
  { key: '06-action-booking-com-NOTIFY_ON_FAILURE', type: 'boolean', read: function () { return BOOKING_ACTION_CONFIG.notifyOnFailure; } },
  { key: '06-action-booking-com-SENDER_ALLOW_LIST', type: 'list', read: function () { return BOOKING_ACTION_CONFIG.senderAllowList; } },
  { key: '06-action-booking-com-ADD_TO_CALENDAR_ENABLED', type: 'boolean', read: function () { return BOOKING_ACTION_CONFIG.addToCalendar.enabled; } },
  { key: '06-action-booking-com-REMOVE_FROM_CALENDAR_ENABLED', type: 'boolean', read: function () { return BOOKING_ACTION_CONFIG.removeFromCalendar.enabled; } },
  { key: '06-action-booking-com-SEARCH_WINDOW_PADDING_DAYS', type: 'number', read: function () { return BOOKING_ACTION_CONFIG.searchWindowPaddingDays; } },
  { key: '06-action-booking-com-WIDE_FALLBACK_WINDOW_YEARS', type: 'number', read: function () { return BOOKING_ACTION_CONFIG.wideFallbackWindowYears; } },
  { key: '06-action-booking-com-EVENT_OVERLAP_TOLERANCE_DAYS', type: 'number', read: function () { return BOOKING_ACTION_CONFIG.eventOverlapToleranceDays; } },
  { key: '06-action-booking-com-CALENDAR_ID', type: 'string', read: function () { return BOOKING_ACTION_CONFIG.calendarId; } },

  // v0.6.0 (quick-260731-tix): TICKETING_PORTALS_ACTION_CONFIG's own
  // fields — registered here for the exact same reason every other
  // action's config fields are: rebuildScriptProperties/
  // checkScriptPropertiesSyntax manage them through the ONE established
  // mechanism, not a parallel one.
  { key: '07-action-ticketing-portals-ENABLED', type: 'boolean', read: function () { return TICKETING_PORTALS_ACTION_CONFIG.enabled; } },
  { key: '07-action-ticketing-portals-NOTIFY_ON_FAILURE', type: 'boolean', read: function () { return TICKETING_PORTALS_ACTION_CONFIG.notifyOnFailure; } },
  {
    key: '07-action-ticketing-portals-TICKETING_PORTALS',
    type: 'json',
    jsonShapeValidator: isValidTicketingPortalsShape,
    jsonShapeDescription: 'an array of {identifyingEmail, calendarId, insertPdfIntoEvent} objects',
    read: function () { return TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals; },
  },

  // quick-260803-us3: TRANSPORT_TICKETS_ACTION_CONFIG's own fields —
  // registered here for the exact same reason every other action's config
  // fields are: rebuildScriptProperties/checkScriptPropertiesSyntax manage
  // them through the ONE established mechanism, not a parallel one.
  { key: '08-action-transport-tickets-ENABLED', type: 'boolean', read: function () { return TRANSPORT_TICKETS_ACTION_CONFIG.enabled; } },
  { key: '08-action-transport-tickets-NOTIFY_ON_FAILURE', type: 'boolean', read: function () { return TRANSPORT_TICKETS_ACTION_CONFIG.notifyOnFailure; } },
  {
    key: '08-action-transport-tickets-TRANSPORT_SENDERS',
    type: 'json',
    // The transport senders list has the IDENTICAL {identifyingEmail,
    // calendarId, insertPdfIntoEvent} shape as the ticketing-portals
    // setting above, so this reuses isValidTicketingPortalsShape rather
    // than cloning an equivalent validator.
    jsonShapeValidator: isValidTicketingPortalsShape,
    jsonShapeDescription: 'an array of {identifyingEmail, calendarId, insertPdfIntoEvent} objects',
    read: function () { return TRANSPORT_TICKETS_ACTION_CONFIG.transportSenders; },
  },
];

/**
 * buildRebuiltPropertiesMap — pure: walks `registry` (an array of
 * SETTINGS_REGISTRY-shaped `{ key, type, read }` entries) and returns the
 * plain `{key: stringValue, ...}` map rebuildScriptProperties (below)
 * writes to Script Properties in one atomic call. For each entry, formats
 * its CURRENT EFFECTIVE value (`entry.read()`, already correctly resolved
 * by the getter chain: existing Script Property if set, else code
 * default) via formatSettingValueForStorage.
 *
 * SCRIPT PROPERTIES EMPTY-STRING CONSTRAINT (real, owner-hit): Google's
 * PropertiesService rejects an empty string value outright — attempting to
 * store `''` throws. Several settings' natural formatted value CAN be
 * empty (an unset optional string override like ICS_ACTION_CONFIG.calendarId
 * formats `null` to `''`; an empty list override formats `[]` to `''`).
 * Whenever a formatted value would be `''`, this function substitutes
 * SETTING_DEFAULT_SENTINEL instead — so EVERY ONE of the registry's
 * entries ALWAYS gets a visible row in the built map, none ever omitted.
 * This is a deliberate discoverability choice (not just a workaround for
 * the empty-string constraint): a setting whose row never exists until
 * someone actively overrides it would be invisible to a new user browsing
 * Project Settings -> Script Properties, who would have no way to
 * discover that setting even exists to configure. The sentinel is
 * recognized as equivalent to "absent"/"use code default" by every read
 * path (readScriptProperty, and transitively describeSettingSyntaxIssue)
 * — see SETTING_DEFAULT_SENTINEL's own doc comment. No GAS globals —
 * `registry`'s entries themselves don't touch PropertiesService
 * (`entry.read()` reads the ALREADY-CORRECTLY-RESOLVED getter value, not
 * a raw property), so this whole function is independently unit-testable
 * under Node with a small hand-built fake registry.
 */
function buildRebuiltPropertiesMap(registry) {
  const map = {};

  registry.forEach(function (entry) {
    const formatted = formatSettingValueForStorage(entry.type, entry.read());
    map[entry.key] = formatted === '' ? SETTING_DEFAULT_SENTINEL : formatted;
  });

  return map;
}

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Exports CONFIG (for the duality regression test and for the sibling
// config files' bridges — see their own Node export blocks), the 5 typed
// get*Setting wrappers, every pure parse/format function, SETTINGS_REGISTRY,
// and SUPPORTED_TRIGGER_INTERVALS_MINUTES. setup()/reconcileProcessEmailsTrigger()/
// rebuildScriptProperties()/checkScriptPropertiesSyntax() are GAS-only
// (reference ScriptApp/PropertiesService/GmailApp-adjacent globals directly)
// and are not exported — they are never invoked under Node, only from the
// Apps Script editor.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CONFIG: CONFIG,
    SUPPORTED_TRIGGER_INTERVALS_MINUTES: SUPPORTED_TRIGGER_INTERVALS_MINUTES,
    SETTING_DEFAULT_SENTINEL: SETTING_DEFAULT_SENTINEL,
    isSettingDefaultSentinel: isSettingDefaultSentinel,
    getStringSetting: getStringSetting,
    getBooleanSetting: getBooleanSetting,
    getNumberSetting: getNumberSetting,
    getListSetting: getListSetting,
    getJsonSetting: getJsonSetting,
    parseBooleanSettingValue: parseBooleanSettingValue,
    parseNumberSettingValue: parseNumberSettingValue,
    parseListSettingValue: parseListSettingValue,
    parseJsonSettingValue: parseJsonSettingValue,
    formatListSettingValue: formatListSettingValue,
    formatJsonSettingValue: formatJsonSettingValue,
    formatSettingValueForStorage: formatSettingValueForStorage,
    describeSettingSyntaxIssue: describeSettingSyntaxIssue,
    buildRebuiltPropertiesMap: buildRebuiltPropertiesMap,
    isValidCalendarIdBySenderShape: isValidCalendarIdBySenderShape,
    isValidTicketingPortalsShape: isValidTicketingPortalsShape,
    SETTINGS_REGISTRY: SETTINGS_REGISTRY,
  };
}

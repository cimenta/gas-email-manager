/**
 * JOB_HUNTERS_ACTION_CONFIG — the config block for the STANDALONE JobHunters
 * capture step (defined in the sibling src/11-action-job-hunters.js). Same
 * "*-action-cfg-*.js" split pattern as every other action -- see
 * src/10-action-cfg-meetings.js's own class-level JSDoc for the full
 * load-order/getter rationale.
 *
 * STANDALONE, NOT A REGISTERED ACTION (D-01): JobHunters is never added to
 * getRegisteredActions() (src/03-action-management.js). processEmails'
 * shared search (`newer_than:<daysBack>d -label:"<labelName>"`) does not fit
 * JobHunters' selection, which is `label:"jobHunter"` with no age bound and
 * no regard for the II/FF labels -- so it runs its OWN search and loop,
 * called directly from processEmails() (see src/11-action-job-hunters.js's
 * runJobHuntersCapture). Its spreadsheet is OWNED BY ANOTHER APP and is
 * append-only -- this script never reads or writes columns 13-20 or any
 * existing row.
 *
 * quick-261016-j0d NEW STEP.
 */

// The committed sentinel default for spreadsheetId below -- deliberately
// never a real spreadsheet ID (D-13 privacy), the same convention as
// CONFIG.calendarId's REPLACE_WITH_CALENDAR_ID. The real ID is set only as
// the Script Property 11-action-job-hunters-SPREADSHEET_ID and is never
// committed. jobHuntersIsSpreadsheetIdConfigured (sibling action file)
// treats this exact string as "not configured".
const JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER = 'REPLACE_WITH_JOBHUNTERS_SPREADSHEET_ID';

// Internal bookkeeping key (epoch ms of the last ATTEMPTED run, scheduled or
// manual) -- written/read directly via PropertiesService by
// runJobHuntersCapture, deliberately absent from SETTINGS_REGISTRY (D-03).
// It must never be reset by rebuildScriptProperties() or flagged by
// checkScriptPropertiesSyntax() -- see INTERNAL_STATE_PROPERTY_KEYS and
// mergePreservedInternalStateProperties in src/01-setup.js, which preserve
// it across the registry's own deleteAllOthers write.
const JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY = '11-action-job-hunters-LAST_RUN_AT_MS';

// quick-261016-nu9: the code default for ownAddresses below. For the common
// case (no Gmail "send mail as" alias), the owner's own address is already
// knowable at runtime -- Session.getActiveUser().getEmail(), the same call
// notifyOwnerOfFailure (src/02-main.js) already uses to resolve its
// recipient -- so no manual Script Property is needed at all. Returns []
// under Node (no Session global) and on any GAS-side failure, so this can
// never throw and never blocks the getter it backs.
function jobHuntersDefaultOwnAddresses() {
  if (typeof Session === 'undefined') {
    return [];
  }

  try {
    const email = Session.getActiveUser().getEmail();
    return email ? [email] : [];
  } catch (e) {
    return [];
  }
}

const JOB_HUNTERS_ACTION_CONFIG = {
  // Cross-cutting enable/disable toggle, same convention as every other
  // action. Checked BEFORE the due-check, and `force` does NOT bypass it
  // (P-9). Script Property override: 11-action-job-hunters-ENABLED
  // (boolean).
  // Script Properties value example: true
  get enabled() {
    return getBooleanSetting('11-action-job-hunters-ENABLED', true);
  },

  // Notify the script owner when a thread or the whole run fails. Script
  // Property override: 11-action-job-hunters-NOTIFY_ON_FAILURE (boolean).
  // Script Properties value example: true
  get notifyOnFailure() {
    return getBooleanSetting('11-action-job-hunters-NOTIFY_ON_FAILURE', true);
  },

  // The external "JobHunters" spreadsheet ID. Committed sentinel default
  // below (JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER) -- the same convention
  // as CONFIG.calendarId's REPLACE_WITH_CALENDAR_ID. The real ID is set
  // only as the Script Property 11-action-job-hunters-SPREADSHEET_ID and is
  // never committed to git.
  // Script Properties value example: 1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789
  get spreadsheetId() {
    return getStringSetting('11-action-job-hunters-SPREADSHEET_ID', JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER);
  },

  // The target sheet NAME (not GID) within that spreadsheet. Script
  // Property override: 11-action-job-hunters-SHEET_NAME (string).
  // Script Properties value example: Emails
  get sheetName() {
    return getStringSetting('11-action-job-hunters-SHEET_NAME', 'Emails');
  },

  // The Gmail label the owner applies manually to a thread they want
  // captured. Script Property override:
  // 11-action-job-hunters-JOB_HUNTER_LABEL_NAME (string).
  // Script Properties value example: jobHunter
  get jobHunterLabelName() {
    return getStringSetting('11-action-job-hunters-JOB_HUNTER_LABEL_NAME', 'jobHunter');
  },

  // The label swapped onto a thread once every message has been captured
  // (or found already captured). Script Property override:
  // 11-action-job-hunters-PROCESSED_LABEL_NAME (string).
  // Script Properties value example: jh
  get processedLabelName() {
    return getStringSetting('11-action-job-hunters-PROCESSED_LABEL_NAME', 'jh');
  },

  // Addresses treated as the owner's own, for the row's "direction" column
  // (outgoing vs incoming) -- see jobHuntersParseFromHeader/buildJobHunterRow
  // in the sibling action file. Code default (quick-261016-nu9) is the
  // running account's own address via jobHuntersDefaultOwnAddresses() above
  // -- no manual setup needed for the common case. Setting this Script
  // Property REPLACES that default entirely (it does not merge with it), so
  // if you also send from a Gmail "send mail as" alias, include your
  // primary address too. Script Property override:
  // 11-action-job-hunters-OWN_ADDRESSES (list, comma-separated).
  // Script Properties value example: you@example.invalid, you.alt@example.invalid
  get ownAddresses() {
    return getListSetting('11-action-job-hunters-OWN_ADDRESSES', jobHuntersDefaultOwnAddresses());
  },

  // Minimum minutes between runs -- gates the WHOLE step (search +
  // processing), independent of CONFIG.triggerIntervalMinutes (which only
  // governs how often processEmails itself fires). Script Property
  // override: 11-action-job-hunters-SEARCH_INTERVAL_MINUTES (number).
  // Script Properties value example: 60
  get searchIntervalMinutes() {
    return getNumberSetting('11-action-job-hunters-SEARCH_INTERVAL_MINUTES', 60);
  },
};

// Node/GAS environment bridge for the Script Properties typed accessor
// helpers (getStringSetting/getBooleanSetting/getNumberSetting/
// getListSetting — all defined in the sibling src/01-setup.js). Under GAS's
// shared global scope these are ALREADY visible here by bare name
// (01-setup.js loads alphabetically first) — no action needed, and this
// `if` block never executes there. Under Node, each `require()`d file is
// its own isolated module with its own scope, so the bare references
// inside JOB_HUNTERS_ACTION_CONFIG's getters above would otherwise throw
// ReferenceError. Same `globalThis` bridge technique already established by
// every other action's config file.
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.getStringSetting = settingsHelpers.getStringSetting;
  globalThis.getBooleanSetting = settingsHelpers.getBooleanSetting;
  globalThis.getNumberSetting = settingsHelpers.getNumberSetting;
  globalThis.getListSetting = settingsHelpers.getListSetting;
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    JOB_HUNTERS_ACTION_CONFIG: JOB_HUNTERS_ACTION_CONFIG,
    JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER: JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER,
    JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY: JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY,
    jobHuntersDefaultOwnAddresses: jobHuntersDefaultOwnAddresses,
  };
}

/**
 * 00-webapp.js — the owner-only Apps Script web app admin UI (doGet), added
 * in quick-260804-kmo: three buttons that run processEmails() /
 * rebuildScriptProperties() / checkScriptPropertiesSyntax() with their
 * results rendered IN THE PAGE, plus a typed Script Properties editor
 * generated from the existing SETTINGS_REGISTRY (src/01-setup.js) so every
 * tunable setting -- including the 3 JSON-typed ones, edited as structured
 * repeating rows rather than hand-typed JSON -- can be viewed and changed
 * without opening the Apps Script editor.
 *
 * OWNER-ONLY (D-01): this file's own assertWebappOwner() guard is a
 * defense-in-depth layer on top of the appsscript.json webapp block
 * (executeAs USER_DEPLOYING, access MYSELF) -- the deployment's own access
 * setting can diverge from the manifest, so neither one alone is
 * sufficient.
 *
 * LOGS OUT OF SCOPE (D-02): this file never reads or persists the Apps
 * Script execution log -- it only renders the immediate RETURN VALUE of
 * each invoked function.
 *
 * CRITICAL: this file is require()d DIRECTLY by
 * test/webapp-settings-form.test.js under Node, where
 * HtmlService/Session/PropertiesService/GmailApp/CalendarApp do not exist.
 * Every GAS global reference in this file MUST stay inside a function body
 * -- never at top level -- or the Node require() itself throws
 * ReferenceError before a single test can run. This is the same "GAS/Node
 * duality" discipline already established in src/01-setup.js for
 * PropertiesService (see readScriptProperty's own doc comment there).
 *
 * All new top-level names in this file are `webapp`-prefixed (except the
 * GAS-reserved `doGet`) to stay collision-free in GAS's one shared global
 * namespace across every src/*.js file.
 */

/**
 * assertWebappOwner — the runtime owner-only guard invoked FIRST in doGet
 * and in every one of the 5 RPC handlers below (T-kmo-01). The
 * appsscript.json webapp block (executeAs USER_DEPLOYING, access MYSELF,
 * D-01/D-08) is NECESSARY but NOT SUFFICIENT on its own: a GAS deployment
 * carries its OWN access setting, configured separately in the Deploy
 * dialog, that can diverge from -- or be created independently of -- the
 * manifest (see this task's blocking Task 4 checkpoint, which reads that
 * setting back explicitly). Under the intended Execute-as-Me / Only-myself
 * deployment, Session.getActiveUser().getEmail() (who is actually running
 * this request) and Session.getEffectiveUser().getEmail() (whose authority
 * the script executes with) are BOTH the owner's own address. Under any
 * looser deployment those two values diverge, or the active user's email
 * comes back empty (an anonymous/unauthenticated caller) -- either case
 * throws here, refusing the request before any Gmail/Calendar/Drive
 * authority is exercised.
 */
function assertWebappOwner() {
  const activeEmail = Session.getActiveUser().getEmail();
  const effectiveEmail = Session.getEffectiveUser().getEmail();

  if (!activeEmail || activeEmail !== effectiveEmail) {
    throw new Error('Access denied: this web app is owner-only.');
  }
}

/**
 * doGet — the web app's single entry point (GAS-reserved name). Calls
 * assertWebappOwner() first, before anything else runs, then serves the
 * static admin page.
 *
 * Uses HtmlService.createHtmlOutputFromFile, NOT createTemplateFromFile:
 * the page is served as a plain static file and the client fetches every
 * setting value AFTERWARD over google.script.run (see webappGetSettings
 * below) -- so no stored setting value is EVER interpolated into HTML
 * server-side (T-kmo-05, D-05's "no free-text JSON entry surface"
 * discipline extends to the page-serving mechanism itself). A template
 * would tempt exactly the pattern this avoids.
 */
function doGet(e) {
  assertWebappOwner();

  return HtmlService.createHtmlOutputFromFile('00-webapp-ui')
    .setTitle('GAS Email Manager - Admin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * webappGetSettings — RPC handler: assertWebappOwner() first, then builds
 * webappBuildSettingsFormModel(SETTINGS_REGISTRY) and attaches, per model,
 * that entry's CURRENT EFFECTIVE value via its own read() closure (D-07) --
 * Script Property override if set, else code default, already correctly
 * resolved by the getter chain -- plus an `isOverridden` boolean (true iff
 * a raw Script Property exists for that key and is not
 * SETTING_DEFAULT_SENTINEL). Returns a plain, JSON-serializable array:
 * google.script.run cannot ferry functions across the RPC boundary, so
 * `value` here is always a plain string/number/boolean/array, never
 * anything closure-bearing.
 */
function webappGetSettings() {
  assertWebappOwner();

  const models = webappBuildSettingsFormModel(SETTINGS_REGISTRY);
  const rawProperties = PropertiesService.getScriptProperties().getProperties();

  return models.map(function (model) {
    const entry = SETTINGS_REGISTRY.filter(function (e) {
      return e.key === model.key;
    })[0];
    const raw = rawProperties[model.key];
    const isOverridden = raw !== undefined && raw !== null && !isSettingDefaultSentinel(raw);

    return {
      key: model.key,
      type: model.type,
      fieldKind: model.fieldKind,
      rowFields: model.rowFields,
      value: entry.read(),
      isOverridden: isOverridden,
    };
  });
}

/**
 * webappSaveSettings — RPC handler: assertWebappOwner() first, then walks
 * SETTINGS_REGISTRY and, for every entry whose key IS present in
 * `submittedByKey`, runs webappCoerceSubmittedSettingValue inside a
 * try/catch, collecting `{key, error}` failures. If ANY key failed, writes
 * NOTHING and returns `{ok:false, savedKeys:[], errors}` -- the same
 * parse-everything-before-any-write discipline the ICS action already
 * follows for its own attachments. Otherwise writes the whole map in ONE
 * `setProperties(map, false)` call; the `false` (deleteAllOthers) is
 * load-bearing -- the UI must never delete a key it did not submit.
 */
function webappSaveSettings(submittedByKey) {
  assertWebappOwner();

  const submitted = submittedByKey || {};
  const toWrite = {};
  const errors = [];

  SETTINGS_REGISTRY.forEach(function (entry) {
    if (!Object.prototype.hasOwnProperty.call(submitted, entry.key)) {
      return;
    }
    try {
      toWrite[entry.key] = webappCoerceSubmittedSettingValue(entry, submitted[entry.key]);
    } catch (err) {
      errors.push({ key: entry.key, error: err.message });
    }
  });

  if (errors.length > 0) {
    return { ok: false, savedKeys: [], errors: errors };
  }

  PropertiesService.getScriptProperties().setProperties(toWrite, false);

  return { ok: true, savedKeys: Object.keys(toWrite), errors: [] };
}

/**
 * webappRunProcessEmails / webappRunRebuildScriptProperties /
 * webappRunCheckScriptPropertiesSyntax — the 3 "Run" RPC handlers.
 * assertWebappOwner() first, then invoke the corresponding EXISTING
 * top-level function (processEmails / rebuildScriptProperties /
 * checkScriptPropertiesSyntax, all defined elsewhere in src/02-main.js and
 * src/01-setup.js -- none of the three is modified by this task) inside a
 * try/catch, returning a plain result object rather than rethrowing, so the
 * page always renders an outcome instead of a bare
 * google.script.run failure-handler stack trace. The syntax checker
 * specifically returns `{ok:true, issues: <the array
 * checkScriptPropertiesSyntax returned>}` so the client renders the
 * STRUCTURED array per D-06 -- never a console scrape.
 */
function webappRunProcessEmails() {
  assertWebappOwner();

  try {
    processEmails();
    return { ok: true, message: 'processEmails completed.' };
  } catch (err) {
    return { ok: false, message: 'processEmails failed: ' + err.message };
  }
}

function webappRunRebuildScriptProperties() {
  assertWebappOwner();

  try {
    rebuildScriptProperties();
    return {
      ok: true,
      message: 'rebuildScriptProperties completed: wrote ' + SETTINGS_REGISTRY.length + ' Script Properties.',
    };
  } catch (err) {
    return { ok: false, message: 'rebuildScriptProperties failed: ' + err.message };
  }
}

function webappRunCheckScriptPropertiesSyntax() {
  assertWebappOwner();

  try {
    const issues = checkScriptPropertiesSyntax();
    return { ok: true, issues: issues };
  } catch (err) {
    return { ok: false, issues: [], message: 'checkScriptPropertiesSyntax failed: ' + err.message };
  }
}

/**
 * webappGetVersion — RPC handler: assertWebappOwner() first, same 6th entry
 * point into the owner-only boundary as the 5 handlers above (T-ou8-01) --
 * it is independently reachable over google.script.run regardless of doGet,
 * so the guard is not optional here either. Exists purely so the page can
 * display the running build next to its title (D-01); deliberately returns
 * nothing but the version.
 */
function webappGetVersion() {
  assertWebappOwner();

  return { version: APP_VERSION };
}

/**
 * webappFieldKindForType — pure: maps a SETTINGS_REGISTRY entry's `type` to
 * the UI-facing `fieldKind` webappBuildSettingsFormModel attaches to each
 * model (D-04/D-05's rendering table). Returns `undefined` for an
 * unrecognized type, which webappBuildSettingsFormModel turns into a
 * descriptive throw naming the offending key -- an unmapped type is a hard
 * error, never a silently-blank field.
 */
function webappFieldKindForType(type) {
  if (type === 'string') return 'text';
  if (type === 'number') return 'number';
  if (type === 'boolean') return 'checkbox';
  if (type === 'list') return 'csv';
  if (type === 'json') return 'rows';
  return undefined;
}

/**
 * webappBuildSettingsFormModel — pure: the ONLY bridge between
 * SETTINGS_REGISTRY (the settings source of truth, src/01-setup.js) and the
 * admin UI (D-03) -- maps every registry entry to `{key, type, fieldKind,
 * rowFields}`. NEVER calls `entry.read()`: this function builds the FORM
 * SHAPE only, not any current value (webappGetSettings attaches each
 * model's live value separately via entry.read(), per D-07).
 *
 * A json-typed entry with no usable `jsonRowFields` row-form descriptor
 * (see that field's own doc comment on SETTINGS_REGISTRY's 3 json entries
 * in src/01-setup.js) THROWS, naming the offending key -- same
 * silent-no-op aversion as resolveTransportSenderMode's invalid-mode throw
 * elsewhere in this codebase: a future json setting shipped without a row
 * form must be a hard error, not an uneditable or unsafe field. An entry
 * whose `type` maps to no known fieldKind ALSO throws, naming the key.
 */
function webappBuildSettingsFormModel(registry) {
  return registry.map(function (entry) {
    const fieldKind = webappFieldKindForType(entry.type);
    if (fieldKind === undefined) {
      throw new Error(
        'webappBuildSettingsFormModel: setting "' + entry.key + '" has an unrecognized type "' + entry.type + '".'
      );
    }

    let rowFields = null;
    if (entry.type === 'json') {
      if (!Array.isArray(entry.jsonRowFields) || entry.jsonRowFields.length === 0) {
        throw new Error(
          'webappBuildSettingsFormModel: json-typed setting "' + entry.key + '" has no jsonRowFields row-form descriptor -- see D-05.'
        );
      }
      rowFields = entry.jsonRowFields;
    }

    return { key: entry.key, type: entry.type, fieldKind: fieldKind, rowFields: rowFields };
  });
}

/**
 * webappNormalizeSubmittedString — pure: trims a submitted row-field value
 * to its string form, treating `null`/`undefined` as `''` -- the shared
 * normalization every string-ish rowField kind (`string`, `stringOrNull`,
 * `enum`) below builds on.
 */
function webappNormalizeSubmittedString(rawValue) {
  if (rawValue === null || rawValue === undefined) {
    return '';
  }
  return String(rawValue).trim();
}

/**
 * webappBuildJsonRowsValue — pure: rebuilds a typed array from a submitted
 * repeating-row form (D-05), one row-object per rowFields descriptor
 * (`{name, kind, options, allowUnset, unsetLabel}` -- see
 * SETTINGS_REGISTRY's 3 json entries in src/01-setup.js for the real
 * descriptors). This is what makes "the editor cannot produce malformed
 * JSON" a proven property: every field is coerced to its declared kind
 * here, BEFORE webappCoerceSubmittedSettingValue ever hands the result to
 * the entry's own jsonShapeValidator.
 *
 * Row-drop rule: a row whose every string-ish field (`string`/
 * `stringOrNull`) trims to empty is a blank throwaway row (e.g. an "Add
 * row" the owner never filled in) and is DROPPED entirely, never included
 * in the output array.
 *
 * Kind rules:
 * - `string` -> trimmed string, `''` on empty (NEVER `null` -- the ICS
 *   shape validator, isValidCalendarIdBySenderShape, requires a real
 *   string).
 * - `stringOrNull` -> trimmed string, or JSON `null` when empty/whitespace.
 * - `boolean` -> `true` for `true`/`'on'`/`'true'`, else `false` -- covers
 *   both a real checkbox boolean and a form-submission `'on'` token.
 * - `enum` -> the exact matched option string; an out-of-`options` value
 *   THROWS naming the field and the bad value. An `allowUnset` field left
 *   blank OMITS that key from the row object entirely -- NEVER an empty
 *   string -- so resolveTransportSenderMode's registry-fallback branch (an
 *   absent `mode` key) keeps working and its invalid-mode throw is never
 *   triggered by the editor itself (see TRANSPORT_SENDERS' own
 *   jsonRowFields comment in src/01-setup.js for the full back-compat
 *   rationale).
 * - any other kind THROWS naming the field.
 */
function webappBuildJsonRowsValue(rows, rowFields) {
  const safeRows = Array.isArray(rows) ? rows : [];

  return safeRows.reduce(function (builtRows, submittedRow) {
    const row = {};
    let hasNonEmptyStringishField = false;

    rowFields.forEach(function (field) {
      const rawValue = submittedRow ? submittedRow[field.name] : undefined;

      if (field.kind === 'string') {
        const value = webappNormalizeSubmittedString(rawValue);
        if (value !== '') {
          hasNonEmptyStringishField = true;
        }
        row[field.name] = value;
        return;
      }

      if (field.kind === 'stringOrNull') {
        const trimmed = webappNormalizeSubmittedString(rawValue);
        if (trimmed !== '') {
          hasNonEmptyStringishField = true;
        }
        row[field.name] = trimmed === '' ? null : trimmed;
        return;
      }

      if (field.kind === 'boolean') {
        row[field.name] = rawValue === true || rawValue === 'on' || rawValue === 'true';
        return;
      }

      if (field.kind === 'enum') {
        const trimmed = webappNormalizeSubmittedString(rawValue);
        if (trimmed === '') {
          if (!field.allowUnset) {
            throw new Error(
              'webappBuildJsonRowsValue: field "' + field.name + '" requires one of [' + field.options.join(', ') + '], got an empty value.'
            );
          }
          // allowUnset: the key is deliberately OMITTED below (never set to
          // '') -- see this function's own doc comment.
          return;
        }
        if (field.options.indexOf(trimmed) === -1) {
          throw new Error(
            'webappBuildJsonRowsValue: field "' +
              field.name +
              '" got invalid value "' +
              trimmed +
              '", expected one of [' +
              field.options.join(', ') +
              '].'
          );
        }
        row[field.name] = trimmed;
        return;
      }

      throw new Error(
        'webappBuildJsonRowsValue: unrecognized rowField kind "' + field.kind + '" for field "' + field.name + '".'
      );
    });

    if (hasNonEmptyStringishField) {
      builtRows.push(row);
    }

    return builtRows;
  }, []);
}

/**
 * webappCoerceSubmittedSettingValue — pure: returns the exact Script
 * Properties storage STRING for `submitted`, per `entry.type`. Reuses the
 * EXISTING pure formatting helpers from src/01-setup.js
 * (parseListSettingValue/formatListSettingValue/formatJsonSettingValue)
 * rather than re-implementing any of them (D-03's "no parallel list"
 * discipline extends to formatting, not just the settings list itself).
 *
 * DELIBERATE ASYMMETRY with the runtime get*Setting wrappers: those never
 * throw on a malformed value (silent fallback to the code default, so
 * processEmails() never crashes over a bad stored value) -- this function
 * THROWS on a non-numeric `number` submission and on a json submission that
 * fails its own `entry.jsonShapeValidator`. The whole point of an editor is
 * that a bad value is refused AT THE DOOR, not quietly discarded: an owner
 * who mistypes a number should see an error, not a silent revert to
 * whatever the code default happens to be.
 *
 * For `type: 'json'`, the submitted rows are first rebuilt via
 * webappBuildJsonRowsValue(submitted, entry.jsonRowFields), then validated
 * against entry.jsonShapeValidator BEFORE formatJsonSettingValue ever runs
 * -- this ordering is what makes "the editor cannot produce malformed JSON"
 * a proven property rather than a hope (see webappBuildJsonRowsValue's own
 * doc comment).
 *
 * EMPTY-STRING SUBSTITUTION: whenever the formatted result comes out as the
 * empty string, this returns SETTING_DEFAULT_SENTINEL instead (Script
 * Properties rejects an empty string outright -- see that sentinel's own
 * doc comment in src/01-setup.js). A json setting with zero rows formats to
 * the literal `'[]'`, which is non-empty, so it is NEVER substituted.
 */
function webappCoerceSubmittedSettingValue(entry, submitted) {
  let formatted;

  if (entry.type === 'boolean') {
    formatted = String(Boolean(submitted));
  } else if (entry.type === 'number') {
    const parsed = Number(submitted);
    if (isNaN(parsed)) {
      throw new Error(
        'webappCoerceSubmittedSettingValue: "' + entry.key + '" requires a numeric value, got: ' + JSON.stringify(submitted)
      );
    }
    formatted = String(parsed);
  } else if (entry.type === 'string') {
    formatted = submitted === null || submitted === undefined ? '' : String(submitted);
  } else if (entry.type === 'list') {
    formatted = formatListSettingValue(parseListSettingValue(submitted));
  } else if (entry.type === 'json') {
    const rows = webappBuildJsonRowsValue(submitted, entry.jsonRowFields);
    if (typeof entry.jsonShapeValidator === 'function' && !entry.jsonShapeValidator(rows)) {
      throw new Error(
        'webappCoerceSubmittedSettingValue: "' + entry.key + '" produced a value that fails its own jsonShapeValidator: ' + JSON.stringify(rows)
      );
    }
    formatted = formatJsonSettingValue(rows);
  } else {
    throw new Error(
      'webappCoerceSubmittedSettingValue: setting "' + entry.key + '" has an unrecognized type "' + entry.type + '".'
    );
  }

  return formatted === '' ? SETTING_DEFAULT_SENTINEL : formatted;
}

// Node/GAS environment bridge for the shared pure helpers this file's core
// reuses (parseListSettingValue/formatListSettingValue/formatJsonSettingValue/
// SETTING_DEFAULT_SENTINEL), all defined in the sibling src/01-setup.js.
// LOAD-ORDER NOTE: under GAS's alphabetical-by-filename load order,
// "00-webapp.js" (this file) loads BEFORE "01-setup.js" -- the reverse of
// what its own name might suggest. This is still safe: every bare reference
// to these names above sits inside a FUNCTION BODY (webappCoerceSubmittedSettingValue
// etc.), never evaluated at file-load time, only when actually CALLED --
// which happens exclusively via google.script.run RPCs long after every
// project file has finished loading. Same "resolve lazily, at the point of
// use" pattern already relied on elsewhere in this codebase (SETTINGS_REGISTRY's
// own LOAD-ORDER NOTE in src/01-setup.js, BOOKING_LANGUAGE_PACKS/getBookingLabels,
// every `action.config` getter). Under Node, each require()d file is its own
// isolated module with its own scope, so the bare references above would
// otherwise throw ReferenceError -- hence this bridge. Same `globalThis`
// bridge technique already established by every *-action-cfg-*.js file's
// equivalent bridge (see src/08-action-cfg-transport-tickets.js).
if (typeof module !== 'undefined' && module.exports) {
  const settingsHelpers = require('./01-setup.js');
  globalThis.parseListSettingValue = settingsHelpers.parseListSettingValue;
  globalThis.formatListSettingValue = settingsHelpers.formatListSettingValue;
  globalThis.formatJsonSettingValue = settingsHelpers.formatJsonSettingValue;
  globalThis.SETTING_DEFAULT_SENTINEL = settingsHelpers.SETTING_DEFAULT_SENTINEL;
}

// GAS-safe Node export (inert under the Apps Script runtime). Only the pure
// core is exported here in Task 1 -- doGet/assertWebappOwner/the 5 RPC
// handlers (Task 2) are GAS-only and never invoked under Node.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    webappFieldKindForType: webappFieldKindForType,
    webappBuildSettingsFormModel: webappBuildSettingsFormModel,
    webappBuildJsonRowsValue: webappBuildJsonRowsValue,
    webappCoerceSubmittedSettingValue: webappCoerceSubmittedSettingValue,
  };
}

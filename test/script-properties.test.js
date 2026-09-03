'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  CONFIG,
  SETTING_DEFAULT_SENTINEL,
  isSettingDefaultSentinel,
  parseBooleanSettingValue,
  parseNumberSettingValue,
  parseListSettingValue,
  parseJsonSettingValue,
  formatListSettingValue,
  formatJsonSettingValue,
  describeSettingSyntaxIssue,
  getStringSetting,
  getBooleanSetting,
  getNumberSetting,
  getListSetting,
  getJsonSetting,
  buildRebuiltPropertiesMap,
  SETTINGS_REGISTRY,
} = require('../src/01-setup.js');

// --- Script Properties cannot store an empty string ("Property value ------
// --- must not be empty") — owner-hit real constraint, fixed via a --------
// --- literal sentinel written in place of any empty value ----------------
//
// A setting's row must ALSO always be visible in Project Settings -> Script
// Properties (discoverability), even before anyone overrides it — so
// rebuildScriptProperties can't simply omit empty-valued keys either. The
// fix: write the literal SETTING_DEFAULT_SENTINEL ('DefaultValue') instead
// of '', for every entry whose formatted value would otherwise be empty;
// every read path treats that sentinel exactly like "absent".

function withFakePropertiesService(initialProperties, fn) {
  const store = Object.assign({}, initialProperties);

  global.PropertiesService = {
    getScriptProperties: function () {
      return {
        getProperty: function (key) {
          return Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null;
        },
      };
    },
  };

  try {
    fn(store);
  } finally {
    delete global.PropertiesService;
  }
}

test('SETTING_DEFAULT_SENTINEL: is exactly the approved literal text "DefaultValue"', () => {
  assert.equal(SETTING_DEFAULT_SENTINEL, 'DefaultValue');
});

test('isSettingDefaultSentinel: true only for the exact sentinel string', () => {
  assert.equal(isSettingDefaultSentinel('DefaultValue'), true);
});

test('isSettingDefaultSentinel: false for near-misses (case, whitespace, empty, unset)', () => {
  assert.equal(isSettingDefaultSentinel('defaultvalue'), false);
  assert.equal(isSettingDefaultSentinel('DefaultValue '), false);
  assert.equal(isSettingDefaultSentinel(''), false);
  assert.equal(isSettingDefaultSentinel(null), false);
  assert.equal(isSettingDefaultSentinel(undefined), false);
});

test('getStringSetting: a raw stored value of exactly "DefaultValue" is treated as absent -> returns codeDefault', () => {
  withFakePropertiesService({ 'some-key': 'DefaultValue' }, function () {
    assert.equal(getStringSetting('some-key', 'CODE_DEFAULT'), 'CODE_DEFAULT');
  });
});

test('getListSetting: a raw stored value of exactly "DefaultValue" is treated as absent -> returns codeDefault', () => {
  withFakePropertiesService({ 'some-key': 'DefaultValue' }, function () {
    assert.deepEqual(getListSetting('some-key', ['fallback@x.com']), ['fallback@x.com']);
  });
});

test('getBooleanSetting: a raw stored value of exactly "DefaultValue" is treated as absent -> returns codeDefault', () => {
  withFakePropertiesService({ 'some-key': 'DefaultValue' }, function () {
    assert.equal(getBooleanSetting('some-key', true), true);
  });
});

test('describeSettingSyntaxIssue: the DefaultValue sentinel is never flagged, for every type', () => {
  assert.equal(describeSettingSyntaxIssue('boolean', SETTING_DEFAULT_SENTINEL), null);
  assert.equal(describeSettingSyntaxIssue('number', SETTING_DEFAULT_SENTINEL), null);
  assert.equal(describeSettingSyntaxIssue('string', SETTING_DEFAULT_SENTINEL), null);
  assert.equal(describeSettingSyntaxIssue('list', SETTING_DEFAULT_SENTINEL), null);
  assert.equal(describeSettingSyntaxIssue('json', SETTING_DEFAULT_SENTINEL), null);
});

// --- buildRebuiltPropertiesMap ------------------------------------------------

test('buildRebuiltPropertiesMap: an entry whose formatted value would be empty gets the sentinel instead of ""', () => {
  const registry = [
    { key: 'K_STRING_EMPTY', type: 'string', read: function () { return null; } },
    { key: 'K_LIST_EMPTY', type: 'list', read: function () { return []; } },
  ];
  const map = buildRebuiltPropertiesMap(registry);
  assert.equal(map.K_STRING_EMPTY, SETTING_DEFAULT_SENTINEL);
  assert.equal(map.K_LIST_EMPTY, SETTING_DEFAULT_SENTINEL);
});

test('buildRebuiltPropertiesMap: an entry with a real non-empty value is written verbatim, not sentineled', () => {
  const registry = [
    { key: 'K_STRING_SET', type: 'string', read: function () { return 'CAL_123'; } },
    { key: 'K_LIST_SET', type: 'list', read: function () { return ['a@x.com', 'b@y.com']; } },
  ];
  const map = buildRebuiltPropertiesMap(registry);
  assert.equal(map.K_STRING_SET, 'CAL_123');
  assert.equal(map.K_LIST_SET, 'a@x.com, b@y.com');
});

test('buildRebuiltPropertiesMap: falsy-but-non-empty formatted values (boolean false, number 0) are NOT sentineled', () => {
  const registry = [
    { key: 'K_BOOL_FALSE', type: 'boolean', read: function () { return false; } },
    { key: 'K_NUMBER_ZERO', type: 'number', read: function () { return 0; } },
  ];
  const map = buildRebuiltPropertiesMap(registry);
  assert.equal(map.K_BOOL_FALSE, 'false');
  assert.equal(map.K_NUMBER_ZERO, '0');
});

test('buildRebuiltPropertiesMap: every registry entry ALWAYS gets a key in the built map — none are ever omitted', () => {
  const registry = [
    { key: 'K_EMPTY', type: 'string', read: function () { return null; } },
    { key: 'K_SET', type: 'string', read: function () { return 'value'; } },
  ];
  const map = buildRebuiltPropertiesMap(registry);
  assert.equal(Object.keys(map).length, 2);
  assert.ok(Object.prototype.hasOwnProperty.call(map, 'K_EMPTY'));
  assert.ok(Object.prototype.hasOwnProperty.call(map, 'K_SET'));
});

test('buildRebuiltPropertiesMap: the real SETTINGS_REGISTRY, with no live overrides, always yields all 32 rows — empty-default fields get the sentinel, non-empty ones their real value', () => {
  // SETTINGS_REGISTRY's read() closures for the ICS/booking/ticketing/
  // transport-tickets/meetings entries reference
  // ICS_ACTION_CONFIG/BOOKING_ACTION_CONFIG/TICKETING_PORTALS_ACTION_CONFIG/
  // TRANSPORT_TICKETS_ACTION_CONFIG/MEETINGS_ACTION_CONFIG by bare name from
  // within src/01-setup.js's own module scope. Under GAS's shared global
  // scope this resolves for free (all five objects are genuinely global);
  // under Node's per-file module isolation, this test supplies them itself
  // via the exact same globalThis-bridge technique the sibling config files
  // already use in the OTHER direction (pulling get*Setting FROM
  // 01-setup.js) — see src/05-action-cfg-ics-import.js's own bridge for the
  // precedent. This is a TEST-ONLY bridge (not added to src/01-setup.js
  // itself, to avoid a circular require: 01-setup.js is what those files
  // themselves require to reach get*Setting).
  global.ICS_ACTION_CONFIG = require('../src/05-action-cfg-ics-import.js').ICS_ACTION_CONFIG;
  global.BOOKING_ACTION_CONFIG = require('../src/06-action-cfg-booking-com-management.js').BOOKING_ACTION_CONFIG;
  global.TICKETING_PORTALS_ACTION_CONFIG = require('../src/07-action-cfg-ticketing-portals.js').TICKETING_PORTALS_ACTION_CONFIG;
  global.TRANSPORT_TICKETS_ACTION_CONFIG = require('../src/08-action-cfg-transport-tickets.js').TRANSPORT_TICKETS_ACTION_CONFIG;
  global.MEETINGS_ACTION_CONFIG = require('../src/10-action-cfg-meetings.js').MEETINGS_ACTION_CONFIG;

  try {
    const map = buildRebuiltPropertiesMap(SETTINGS_REGISTRY);
    assert.equal(Object.keys(map).length, 32);

    // Optional-override fields whose code default is null/[] -> sentinel.
    assert.equal(map['05-action-ics-CALENDAR_ID'], SETTING_DEFAULT_SENTINEL);
    assert.equal(map['06-action-booking-com-CALENDAR_ID'], SETTING_DEFAULT_SENTINEL);
    assert.equal(map['05-action-ics-IMPORT_ONLY_FROM'], SETTING_DEFAULT_SENTINEL);

    // A list field with a non-empty code default -> its real value, not the sentinel.
    assert.equal(map['06-action-booking-com-SENDER_ALLOW_LIST'], 'noreply@booking.com');

    // A json field whose empty-array code default formats to the non-empty
    // string "[]" -> untouched (never sentineled, since "[]" !== "").
    assert.equal(map['05-action-ics-CALENDAR_ID_BY_SENDER'], '[]');

    // The new v0.6.0 string field -> its real code default, not the sentinel.
    assert.equal(map['01-setup-TICKET_ATTACHMENT_DRIVE_FOLDER_NAME'], 'GAS Email Manager - Ticket Attachments');

    // The new v0.6.0 json field (three seeded portal entries as of
    // quick-260816-ocw: enigoo.cz, Kino Art, and Ticketmaster CZ, non-empty)
    // -> its real formatted value, not the sentinel.
    assert.equal(
      map['07-action-ticketing-portals-TICKETING_PORTALS'],
      '[{"identifyingEmail":"no-reply@enigoo.cz","calendarId":null,"insertPdfIntoEvent":false},{"identifyingEmail":"rezervace@kinoart.cz","calendarId":null,"insertPdfIntoEvent":false},{"identifyingEmail":"noreply@ticketmaster.cz","calendarId":null,"insertPdfIntoEvent":false}]'
    );

    // quick-260803-us3: ICS_ACTION_CONFIG.excludeFrom's empty-array code
    // default -> sentinel (same treatment as importOnlyFrom above).
    assert.equal(map['05-action-ics-EXCLUDE_FROM'], SETTING_DEFAULT_SENTINEL);

    // quick-260804-bs7: TRANSPORT_TICKETS_ACTION_CONFIG's json field now
    // seeds TWO entries (RegioJet, explicit mode "ics", plus IDOS.cz, mode
    // "body") -> its real formatted value, not the sentinel.
    assert.equal(
      map['08-action-transport-tickets-TRANSPORT_SENDERS'],
      '[{"identifyingEmail":"jizdenky@regiojet.cz","calendarId":null,"insertPdfIntoEvent":false,"mode":"ics"},{"identifyingEmail":"jizdenky@idos.svt.cz","calendarId":null,"insertPdfIntoEvent":false,"mode":"body"}]'
    );

    // quick-260820-g4r: MEETINGS_ACTION_CONFIG's 4 fields -- ENABLED and
    // NOTIFY_ON_FAILURE code default true, DEFAULT_DURATION_MINUTES code
    // default 60, MEETING_SYSTEMS ships the single default Teamio entry
    // (never sentineled, since the code default is non-empty).
    assert.equal(map['10-action-meetings-ENABLED'], 'true');
    assert.equal(map['10-action-meetings-NOTIFY_ON_FAILURE'], 'true');
    assert.equal(map['10-action-meetings-DEFAULT_DURATION_MINUTES'], '60');
    assert.equal(map['10-action-meetings-MEETING_SYSTEMS'], '[{"domainPattern":"*.teamio.com","calendarId":null}]');
  } finally {
    delete global.ICS_ACTION_CONFIG;
    delete global.BOOKING_ACTION_CONFIG;
    delete global.TICKETING_PORTALS_ACTION_CONFIG;
    delete global.TRANSPORT_TICKETS_ACTION_CONFIG;
    delete global.MEETINGS_ACTION_CONFIG;
  }
});
const { ICS_ACTION_CONFIG } = require('../src/05-action-cfg-ics-import.js');
const { BOOKING_ACTION_CONFIG } = require('../src/06-action-cfg-booking-com-management.js');
const { TICKETING_PORTALS_ACTION_CONFIG } = require('../src/07-action-cfg-ticketing-portals.js');
const { MEETINGS_ACTION_CONFIG } = require('../src/10-action-cfg-meetings.js');

// --- parseBooleanSettingValue ------------------------------------------------

test('parseBooleanSettingValue: "true" (any case) -> true', () => {
  assert.equal(parseBooleanSettingValue('true', false), true);
  assert.equal(parseBooleanSettingValue('TRUE', false), true);
  assert.equal(parseBooleanSettingValue('True', false), true);
});

test('parseBooleanSettingValue: "false" (any case) -> false', () => {
  assert.equal(parseBooleanSettingValue('false', true), false);
  assert.equal(parseBooleanSettingValue('FALSE', true), false);
});

test('parseBooleanSettingValue: surrounding whitespace tolerated', () => {
  assert.equal(parseBooleanSettingValue('  true  ', false), true);
});

test('parseBooleanSettingValue: null/undefined (unset) -> codeDefault', () => {
  assert.equal(parseBooleanSettingValue(null, true), true);
  assert.equal(parseBooleanSettingValue(undefined, false), false);
});

test('parseBooleanSettingValue: malformed token (not "true"/"false") -> codeDefault, never throws', () => {
  assert.doesNotThrow(() => {
    assert.equal(parseBooleanSettingValue('yes', 'DEFAULT'), 'DEFAULT');
    assert.equal(parseBooleanSettingValue('1', 'DEFAULT'), 'DEFAULT');
    assert.equal(parseBooleanSettingValue('', 'DEFAULT'), 'DEFAULT');
  });
});

// --- parseNumberSettingValue -------------------------------------------------

test('parseNumberSettingValue: valid numeric string -> Number', () => {
  assert.equal(parseNumberSettingValue('7', 0), 7);
  assert.equal(parseNumberSettingValue('-3.5', 0), -3.5);
});

test('parseNumberSettingValue: null/undefined (unset) -> codeDefault', () => {
  assert.equal(parseNumberSettingValue(null, 42), 42);
  assert.equal(parseNumberSettingValue(undefined, 42), 42);
});

test('parseNumberSettingValue: non-numeric (NaN) -> codeDefault, never throws', () => {
  assert.doesNotThrow(() => {
    assert.equal(parseNumberSettingValue('not-a-number', 99), 99);
    assert.equal(parseNumberSettingValue('5 days', 99), 99);
  });
});

test('parseNumberSettingValue: empty string is a documented JS quirk (Number("") === 0, NOT NaN) -> resolves to 0, not codeDefault', () => {
  // Deliberate, not a bug: Number('') === 0 per JS semantics; the spec's
  // "invalid -> NaN -> falls back" rule only fires on genuine NaN.
  assert.equal(parseNumberSettingValue('', 99), 0);
});

// --- parseListSettingValue ---------------------------------------------------

test('parseListSettingValue: comma-separated entries, each trimmed', () => {
  assert.deepEqual(parseListSettingValue('a@x.com, b@y.com,c@z.com'), ['a@x.com', 'b@y.com', 'c@z.com']);
});

test('parseListSettingValue: single entry, no commas', () => {
  assert.deepEqual(parseListSettingValue('only@one.com'), ['only@one.com']);
});

test('parseListSettingValue: empty string -> [] not [""]', () => {
  assert.deepEqual(parseListSettingValue(''), []);
});

test('parseListSettingValue: whitespace-only string -> []', () => {
  assert.deepEqual(parseListSettingValue('   '), []);
});

test('parseListSettingValue: trailing comma / empty entries filtered out', () => {
  assert.deepEqual(parseListSettingValue('a@x.com, ,b@y.com,'), ['a@x.com', 'b@y.com']);
});

test('parseListSettingValue: null/undefined -> []', () => {
  assert.deepEqual(parseListSettingValue(null), []);
  assert.deepEqual(parseListSettingValue(undefined), []);
});

// --- parseJsonSettingValue ----------------------------------------------------

test('parseJsonSettingValue: valid array of {from, calendarId} objects -> parsed array', () => {
  const raw = '[{"from":"a@x.com","calendarId":"CAL_A"},{"from":"b@y.com","calendarId":"CAL_B"}]';
  assert.deepEqual(parseJsonSettingValue(raw, []), [
    { from: 'a@x.com', calendarId: 'CAL_A' },
    { from: 'b@y.com', calendarId: 'CAL_B' },
  ]);
});

test('parseJsonSettingValue: valid empty array -> []', () => {
  assert.deepEqual(parseJsonSettingValue('[]', ['SENTINEL']), []);
});

test('parseJsonSettingValue: null/undefined (unset) -> codeDefault', () => {
  assert.deepEqual(parseJsonSettingValue(null, ['DEFAULT']), ['DEFAULT']);
});

test('parseJsonSettingValue: invalid JSON syntax -> codeDefault, never throws', () => {
  assert.doesNotThrow(() => {
    assert.deepEqual(parseJsonSettingValue('{not valid json', ['DEFAULT']), ['DEFAULT']);
  });
});

test('parseJsonSettingValue: valid JSON but wrong shape (not an array) -> codeDefault', () => {
  assert.deepEqual(parseJsonSettingValue('{"from":"a@x.com"}', ['DEFAULT']), ['DEFAULT']);
});

test('parseJsonSettingValue: valid JSON array but entries missing required string fields -> codeDefault', () => {
  assert.deepEqual(parseJsonSettingValue('[{"from":"a@x.com"}]', ['DEFAULT']), ['DEFAULT']);
  assert.deepEqual(parseJsonSettingValue('[{"from":123,"calendarId":"CAL_A"}]', ['DEFAULT']), ['DEFAULT']);
});

// --- parseJsonSettingValue: generalized shape-validator parameter -----------
// (quick-260731-tix) — a second, differently-shaped JSON setting
// (ticketing portals: {identifyingEmail, calendarId, insertPdfIntoEvent})
// is exactly the extension point isValidCalendarIdBySenderShape's own doc
// comment anticipated. When the 3rd (shape-validator) argument is omitted,
// behavior is UNCHANGED (defaults to the original ICS shape check) — every
// test above this section proves that backward-compatible default still
// holds. These tests prove the NEW explicit-validator path works
// independently, using a representative differently-shaped validator (not
// the real isValidTicketingPortalsShape, to keep this test file decoupled
// from the ticketing action's own file — see test/ticketing-portals.test.js
// for that function's own dedicated coverage).

function isValidPersonAgeShape(value) {
  return Array.isArray(value) && value.every(function (entry) {
    return entry !== null && typeof entry === 'object' && typeof entry.name === 'string' && typeof entry.age === 'number';
  });
}

test('parseJsonSettingValue: an explicit custom shape validator is used INSTEAD of the default ICS shape check when provided', () => {
  const raw = '[{"name":"Alice","age":30}]';
  assert.deepEqual(parseJsonSettingValue(raw, [], isValidPersonAgeShape), [{ name: 'Alice', age: 30 }]);
});

test('parseJsonSettingValue: a value valid under the DEFAULT (ICS) shape but NOT the explicit custom shape falls back to codeDefault', () => {
  const raw = '[{"from":"a@x.com","calendarId":"CAL_A"}]';
  assert.deepEqual(parseJsonSettingValue(raw, ['DEFAULT'], isValidPersonAgeShape), ['DEFAULT']);
});

test('parseJsonSettingValue: omitting the shape-validator argument still uses the original ICS shape check (zero regression)', () => {
  const raw = '[{"from":"a@x.com","calendarId":"CAL_A"}]';
  assert.deepEqual(parseJsonSettingValue(raw, ['DEFAULT']), [{ from: 'a@x.com', calendarId: 'CAL_A' }]);
});

// --- formatListSettingValue / formatJsonSettingValue (write-side, for rebuildScriptProperties) ---

test('formatListSettingValue: array -> comma-and-space-joined string', () => {
  assert.equal(formatListSettingValue(['a@x.com', 'b@y.com']), 'a@x.com, b@y.com');
});

test('formatListSettingValue: empty array -> empty string', () => {
  assert.equal(formatListSettingValue([]), '');
});

test('formatListSettingValue: null/undefined -> empty string (never throws)', () => {
  assert.doesNotThrow(() => {
    assert.equal(formatListSettingValue(null), '');
    assert.equal(formatListSettingValue(undefined), '');
  });
});

test('formatJsonSettingValue: array -> valid JSON string round-tripping via parseJsonSettingValue', () => {
  const value = [{ from: 'a@x.com', calendarId: 'CAL_A' }];
  const formatted = formatJsonSettingValue(value);
  assert.deepEqual(parseJsonSettingValue(formatted, 'SENTINEL'), value);
});

test('formatJsonSettingValue: empty array -> "[]"', () => {
  assert.equal(formatJsonSettingValue([]), '[]');
});

// --- describeSettingSyntaxIssue (checkScriptPropertiesSyntax's pure core) ----

test('describeSettingSyntaxIssue: unset (null) is never an issue, for every type', () => {
  assert.equal(describeSettingSyntaxIssue('boolean', null), null);
  assert.equal(describeSettingSyntaxIssue('number', null), null);
  assert.equal(describeSettingSyntaxIssue('string', null), null);
  assert.equal(describeSettingSyntaxIssue('list', null), null);
  assert.equal(describeSettingSyntaxIssue('json', null), null);
});

test('describeSettingSyntaxIssue: valid boolean/number/json values are never flagged', () => {
  assert.equal(describeSettingSyntaxIssue('boolean', 'true'), null);
  assert.equal(describeSettingSyntaxIssue('number', '5'), null);
  assert.equal(describeSettingSyntaxIssue('json', '[]'), null);
});

test('describeSettingSyntaxIssue: malformed boolean is flagged with a specific description', () => {
  const issue = describeSettingSyntaxIssue('boolean', 'yes');
  assert.notEqual(issue, null);
  assert.match(issue, /true|false/i);
});

test('describeSettingSyntaxIssue: malformed number is flagged with a specific description', () => {
  const issue = describeSettingSyntaxIssue('number', 'not-a-number');
  assert.notEqual(issue, null);
  assert.match(issue, /numeric/i);
});

test('describeSettingSyntaxIssue: invalid JSON is flagged with a specific description', () => {
  const issue = describeSettingSyntaxIssue('json', '{not valid json');
  assert.notEqual(issue, null);
});

test('describeSettingSyntaxIssue: wrong JSON shape is flagged with a specific description', () => {
  const issue = describeSettingSyntaxIssue('json', '[{"from":"a@x.com"}]');
  assert.notEqual(issue, null);
});

test('describeSettingSyntaxIssue: an explicit custom shape validator + description is used INSTEAD of the default ICS one (quick-260731-tix generalization)', () => {
  const validPass = describeSettingSyntaxIssue('json', '[{"name":"Alice","age":30}]', isValidPersonAgeShape, 'an array of {name, age} objects');
  assert.equal(validPass, null);

  const issue = describeSettingSyntaxIssue('json', '[{"from":"a@x.com","calendarId":"CAL_A"}]', isValidPersonAgeShape, 'an array of {name, age} objects');
  assert.notEqual(issue, null);
  assert.match(issue, /name, age/);
});

test('describeSettingSyntaxIssue: string/list types are never flagged (any string is valid)', () => {
  assert.equal(describeSettingSyntaxIssue('string', 'anything at all'), null);
  assert.equal(describeSettingSyntaxIssue('list', ''), null);
  assert.equal(describeSettingSyntaxIssue('list', 'a@x.com, b@y.com'), null);
});

// --- get*Setting wrappers (GAS/Node duality guard, exercised directly) ------

test('getStringSetting: PropertiesService absent under Node -> returns codeDefault', () => {
  assert.equal(getStringSetting('some-key', 'DEFAULT_STRING'), 'DEFAULT_STRING');
});

test('getBooleanSetting: PropertiesService absent under Node -> returns codeDefault', () => {
  assert.equal(getBooleanSetting('some-key', true), true);
  assert.equal(getBooleanSetting('some-key', false), false);
});

test('getNumberSetting: PropertiesService absent under Node -> returns codeDefault', () => {
  assert.equal(getNumberSetting('some-key', 7), 7);
});

test('getListSetting: PropertiesService absent under Node -> returns codeDefault', () => {
  assert.deepEqual(getListSetting('some-key', ['noreply@booking.com']), ['noreply@booking.com']);
});

test('getJsonSetting: PropertiesService absent under Node -> returns codeDefault', () => {
  assert.deepEqual(getJsonSetting('some-key', []), []);
});

test('getJsonSetting: accepts an explicit shape-validator argument, passed through to parseJsonSettingValue (PropertiesService absent -> codeDefault regardless, but must not throw)', () => {
  assert.doesNotThrow(() => {
    assert.deepEqual(getJsonSetting('some-key', ['DEFAULT'], isValidPersonAgeShape), ['DEFAULT']);
  });
});

// --- CRITICAL: the Node/GAS-duality regression test -------------------------
//
// This is the single most important test in this file. PropertiesService
// does not exist under Node — referencing it throws ReferenceError, not
// "evaluates to undefined". Every field below is read by dozens of existing
// tests throughout the suite; a missed guard anywhere in the getter chain
// would break the ENTIRE 201-test baseline's ability to run at all, not just
// this file. This test proves every field still resolves to its EXACT
// pre-refactor documented default with PropertiesService absent.

test('CRITICAL: with PropertiesService absent (Node), every CONFIG field matches its exact pre-refactor default', () => {
  assert.equal(typeof PropertiesService, 'undefined');
  assert.equal(CONFIG.calendarId, 'REPLACE_WITH_CALENDAR_ID');
  assert.equal(CONFIG.labelName, 'II');
  assert.equal(CONFIG.failedLabelName, 'FF');
  assert.equal(CONFIG.daysBack, 2);
  assert.equal(CONFIG.installTrigger, true);
  assert.equal(CONFIG.triggerIntervalMinutes, 5);
  assert.equal(CONFIG.ticketAttachmentDriveFolderName, 'GAS Email Manager - Ticket Attachments');
});

test('CRITICAL: with PropertiesService absent (Node), every ICS_ACTION_CONFIG field matches its exact pre-refactor default', () => {
  assert.equal(typeof PropertiesService, 'undefined');
  assert.equal(ICS_ACTION_CONFIG.enabled, true);
  assert.equal(ICS_ACTION_CONFIG.notifyOnFailure, true);
  assert.deepEqual(ICS_ACTION_CONFIG.importOnlyFrom, []);
  assert.equal(ICS_ACTION_CONFIG.calendarId, null);
  assert.deepEqual(ICS_ACTION_CONFIG.calendarIdBySender, []);
});

test('CRITICAL: with PropertiesService absent (Node), every BOOKING_ACTION_CONFIG field (including nested) matches its exact pre-refactor default', () => {
  assert.equal(typeof PropertiesService, 'undefined');
  assert.equal(BOOKING_ACTION_CONFIG.enabled, true);
  assert.equal(BOOKING_ACTION_CONFIG.notifyOnFailure, true);
  assert.deepEqual(BOOKING_ACTION_CONFIG.senderAllowList, ['noreply@booking.com']);
  assert.equal(BOOKING_ACTION_CONFIG.addToCalendar.enabled, true);
  assert.equal(BOOKING_ACTION_CONFIG.removeFromCalendar.enabled, true);
  assert.equal(BOOKING_ACTION_CONFIG.searchWindowPaddingDays, 7);
  assert.equal(BOOKING_ACTION_CONFIG.wideFallbackWindowYears, 1);
  assert.equal(BOOKING_ACTION_CONFIG.eventOverlapToleranceDays, 1);
  assert.equal(BOOKING_ACTION_CONFIG.calendarId, null);
});

test('CRITICAL: with PropertiesService absent (Node), every TICKETING_PORTALS_ACTION_CONFIG field matches its exact pre-refactor default (quick-260731-tix, updated quick-260731-kar for the Kino Art entry, updated quick-260816-ocw for the Ticketmaster CZ entry)', () => {
  assert.equal(typeof PropertiesService, 'undefined');
  assert.equal(TICKETING_PORTALS_ACTION_CONFIG.enabled, true);
  assert.equal(TICKETING_PORTALS_ACTION_CONFIG.notifyOnFailure, true);
  assert.deepEqual(TICKETING_PORTALS_ACTION_CONFIG.ticketingPortals, [
    { identifyingEmail: 'no-reply@enigoo.cz', calendarId: null, insertPdfIntoEvent: false },
    { identifyingEmail: 'rezervace@kinoart.cz', calendarId: null, insertPdfIntoEvent: false },
    { identifyingEmail: 'noreply@ticketmaster.cz', calendarId: null, insertPdfIntoEvent: false },
  ]);
});

test('CRITICAL: with PropertiesService absent (Node), every MEETINGS_ACTION_CONFIG field matches its exact code default (quick-260820-g4r)', () => {
  assert.equal(typeof PropertiesService, 'undefined');
  assert.equal(MEETINGS_ACTION_CONFIG.enabled, true);
  assert.equal(MEETINGS_ACTION_CONFIG.notifyOnFailure, true);
  assert.equal(MEETINGS_ACTION_CONFIG.defaultDurationMinutes, 60);
  assert.deepEqual(MEETINGS_ACTION_CONFIG.meetingSystems, [{ domainPattern: '*.teamio.com', calendarId: null }]);
});

// --- SETTINGS_REGISTRY completeness (a completeness guard, same spirit as ---
// --- the language-pack completeness test) -----------------------------------

const EXPECTED_SETTINGS_REGISTRY = [
  { key: '01-setup-CALENDAR_ID', type: 'string' },
  { key: '01-setup-LABEL_NAME', type: 'string' },
  { key: '01-setup-FAILED_LABEL_NAME', type: 'string' },
  { key: '01-setup-DAYS_BACK', type: 'number' },
  { key: '01-setup-INSTALL_TRIGGER', type: 'boolean' },
  { key: '01-setup-TRIGGER_INTERVAL_MINUTES', type: 'number' },
  { key: '01-setup-TICKET_ATTACHMENT_DRIVE_FOLDER_NAME', type: 'string' },
  { key: '05-action-ics-ENABLED', type: 'boolean' },
  { key: '05-action-ics-NOTIFY_ON_FAILURE', type: 'boolean' },
  { key: '05-action-ics-IMPORT_ONLY_FROM', type: 'list' },
  { key: '05-action-ics-EXCLUDE_FROM', type: 'list' },
  { key: '05-action-ics-CALENDAR_ID', type: 'string' },
  { key: '05-action-ics-CALENDAR_ID_BY_SENDER', type: 'json' },
  { key: '06-action-booking-com-ENABLED', type: 'boolean' },
  { key: '06-action-booking-com-NOTIFY_ON_FAILURE', type: 'boolean' },
  { key: '06-action-booking-com-SENDER_ALLOW_LIST', type: 'list' },
  { key: '06-action-booking-com-ADD_TO_CALENDAR_ENABLED', type: 'boolean' },
  { key: '06-action-booking-com-REMOVE_FROM_CALENDAR_ENABLED', type: 'boolean' },
  { key: '06-action-booking-com-SEARCH_WINDOW_PADDING_DAYS', type: 'number' },
  { key: '06-action-booking-com-WIDE_FALLBACK_WINDOW_YEARS', type: 'number' },
  { key: '06-action-booking-com-EVENT_OVERLAP_TOLERANCE_DAYS', type: 'number' },
  { key: '06-action-booking-com-CALENDAR_ID', type: 'string' },
  { key: '07-action-ticketing-portals-ENABLED', type: 'boolean' },
  { key: '07-action-ticketing-portals-NOTIFY_ON_FAILURE', type: 'boolean' },
  { key: '07-action-ticketing-portals-TICKETING_PORTALS', type: 'json' },
  { key: '08-action-transport-tickets-ENABLED', type: 'boolean' },
  { key: '08-action-transport-tickets-NOTIFY_ON_FAILURE', type: 'boolean' },
  { key: '08-action-transport-tickets-TRANSPORT_SENDERS', type: 'json' },
  { key: '10-action-meetings-ENABLED', type: 'boolean' },
  { key: '10-action-meetings-NOTIFY_ON_FAILURE', type: 'boolean' },
  { key: '10-action-meetings-DEFAULT_DURATION_MINUTES', type: 'number' },
  { key: '10-action-meetings-MEETING_SYSTEMS', type: 'json' },
];

test('SETTINGS_REGISTRY: exactly the 32 approved keys, in the exact approved names and types, no more no fewer', () => {
  assert.equal(SETTINGS_REGISTRY.length, 32);
  const actual = SETTINGS_REGISTRY.map(function (entry) {
    return { key: entry.key, type: entry.type };
  });
  assert.deepEqual(actual, EXPECTED_SETTINGS_REGISTRY);
});

test('SETTINGS_REGISTRY: every json-typed entry carries its own jsonShapeValidator and jsonShapeDescription (quick-260731-tix generalization)', () => {
  const jsonEntries = SETTINGS_REGISTRY.filter(function (entry) {
    return entry.type === 'json';
  });
  assert.equal(jsonEntries.length, 4);
  jsonEntries.forEach(function (entry) {
    assert.equal(typeof entry.jsonShapeValidator, 'function');
    assert.equal(typeof entry.jsonShapeDescription, 'string');
  });
});

test('SETTINGS_REGISTRY: every entry has a callable read() function', () => {
  SETTINGS_REGISTRY.forEach(function (entry) {
    assert.equal(typeof entry.read, 'function');
  });
});

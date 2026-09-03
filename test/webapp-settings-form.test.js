'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const {
  SETTINGS_REGISTRY,
  SETTING_DEFAULT_SENTINEL,
  formatListSettingValue,
  parseListSettingValue,
  formatJsonSettingValue,
  parseJsonSettingValue,
  isValidCalendarIdBySenderShape,
  isValidTicketingPortalsShape,
} = require('../src/01-setup.js');

const {
  webappBuildSettingsFormModel,
  webappBuildJsonRowsValue,
  webappCoerceSubmittedSettingValue,
} = require('../src/00-webapp.js');

// --- webappBuildSettingsFormModel -------------------------------------------

test('webappBuildSettingsFormModel: over the real 32-entry registry, returns 32 models with {key, type, fieldKind, rowFields}', () => {
  const models = webappBuildSettingsFormModel(SETTINGS_REGISTRY);
  assert.equal(models.length, 32);
  models.forEach(function (model) {
    assert.equal(typeof model.key, 'string');
    assert.equal(typeof model.type, 'string');
    assert.equal(typeof model.fieldKind, 'string');
    assert.ok('rowFields' in model);
  });
});

test('webappBuildSettingsFormModel: fieldKind mapping is exactly string->text, number->number, boolean->checkbox, list->csv, json->rows', () => {
  const models = webappBuildSettingsFormModel(SETTINGS_REGISTRY);
  const byKey = {};
  models.forEach(function (model) {
    byKey[model.key] = model;
  });

  assert.equal(byKey['01-setup-CALENDAR_ID'].fieldKind, 'text');
  assert.equal(byKey['01-setup-DAYS_BACK'].fieldKind, 'number');
  assert.equal(byKey['01-setup-INSTALL_TRIGGER'].fieldKind, 'checkbox');
  assert.equal(byKey['05-action-ics-IMPORT_ONLY_FROM'].fieldKind, 'csv');
  assert.equal(byKey['05-action-ics-CALENDAR_ID_BY_SENDER'].fieldKind, 'rows');

  // quick-260820-g4r: Meetings' 4 keys.
  assert.equal(byKey['10-action-meetings-ENABLED'].fieldKind, 'checkbox');
  assert.equal(byKey['10-action-meetings-NOTIFY_ON_FAILURE'].fieldKind, 'checkbox');
  assert.equal(byKey['10-action-meetings-DEFAULT_DURATION_MINUTES'].fieldKind, 'number');
  assert.equal(byKey['10-action-meetings-MEETING_SYSTEMS'].fieldKind, 'rows');
});

test('webappBuildSettingsFormModel: the 4 json keys get fieldKind "rows" and rowFields with the exact expected field names', () => {
  const models = webappBuildSettingsFormModel(SETTINGS_REGISTRY);
  const byKey = {};
  models.forEach(function (model) {
    byKey[model.key] = model;
  });

  const icsNames = byKey['05-action-ics-CALENDAR_ID_BY_SENDER'].rowFields.map(function (f) {
    return f.name;
  });
  assert.deepEqual(icsNames, ['from', 'calendarId']);

  const portalNames = byKey['07-action-ticketing-portals-TICKETING_PORTALS'].rowFields.map(function (f) {
    return f.name;
  });
  assert.deepEqual(portalNames, ['identifyingEmail', 'calendarId', 'insertPdfIntoEvent']);

  const transportNames = byKey['08-action-transport-tickets-TRANSPORT_SENDERS'].rowFields.map(function (f) {
    return f.name;
  });
  assert.deepEqual(transportNames, ['identifyingEmail', 'calendarId', 'insertPdfIntoEvent', 'mode']);

  const meetingNames = byKey['10-action-meetings-MEETING_SYSTEMS'].rowFields.map(function (f) {
    return f.name;
  });
  assert.deepEqual(meetingNames, ['domainPattern', 'calendarId']);
});

test('webappBuildSettingsFormModel: completeness guard -- every json-typed entry in the real registry carries a non-empty jsonRowFields array', () => {
  const jsonEntries = SETTINGS_REGISTRY.filter(function (entry) {
    return entry.type === 'json';
  });
  assert.equal(jsonEntries.length, 4);
  jsonEntries.forEach(function (entry) {
    assert.ok(Array.isArray(entry.jsonRowFields));
    assert.ok(entry.jsonRowFields.length > 0);
  });
});

test('webappBuildSettingsFormModel: a hand-built registry with a json entry lacking jsonRowFields THROWS, naming the offending key', () => {
  const badRegistry = [{ key: 'fake-json-key', type: 'json', read: function () { return []; } }];
  assert.throws(function () {
    webappBuildSettingsFormModel(badRegistry);
  }, /fake-json-key/);
});

test('webappBuildSettingsFormModel: a hand-built registry entry with an unrecognized type THROWS', () => {
  const badRegistry = [{ key: 'fake-weird-key', type: 'wat', read: function () { return null; } }];
  assert.throws(function () {
    webappBuildSettingsFormModel(badRegistry);
  }, /fake-weird-key/);
});

test("webappBuildSettingsFormModel: never invokes the real registry's read() closures (proven by success under Node with no fake GAS globals installed)", () => {
  assert.equal(typeof PropertiesService, 'undefined');
  assert.doesNotThrow(function () {
    webappBuildSettingsFormModel(SETTINGS_REGISTRY);
  });
});

// --- webappBuildJsonRowsValue ------------------------------------------------

const ICS_ROW_FIELDS = SETTINGS_REGISTRY.find(function (e) {
  return e.key === '05-action-ics-CALENDAR_ID_BY_SENDER';
}).jsonRowFields;
const PORTALS_ROW_FIELDS = SETTINGS_REGISTRY.find(function (e) {
  return e.key === '07-action-ticketing-portals-TICKETING_PORTALS';
}).jsonRowFields;
const TRANSPORT_ROW_FIELDS = SETTINGS_REGISTRY.find(function (e) {
  return e.key === '08-action-transport-tickets-TRANSPORT_SENDERS';
}).jsonRowFields;

test('webappBuildJsonRowsValue: happy path for the ICS calendarIdBySender shape', () => {
  const result = webappBuildJsonRowsValue(
    [{ from: 'a@x.com', calendarId: 'cal-1@group.calendar.google.com' }],
    ICS_ROW_FIELDS
  );
  assert.deepEqual(result, [{ from: 'a@x.com', calendarId: 'cal-1@group.calendar.google.com' }]);
  assert.equal(isValidCalendarIdBySenderShape(result), true);
});

test('webappBuildJsonRowsValue: happy path for the ticketing-portals shape', () => {
  const result = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'cal-2@group.calendar.google.com', insertPdfIntoEvent: true }],
    PORTALS_ROW_FIELDS
  );
  assert.deepEqual(result, [
    { identifyingEmail: 'no-reply@enigoo.cz', calendarId: 'cal-2@group.calendar.google.com', insertPdfIntoEvent: true },
  ]);
  assert.equal(isValidTicketingPortalsShape(result), true);
});

test('webappBuildJsonRowsValue: happy path for the transport-senders shape (with an explicit mode)', () => {
  const result = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'jizdenky@regiojet.cz', calendarId: 'cal-3@group.calendar.google.com', insertPdfIntoEvent: false, mode: 'ics' }],
    TRANSPORT_ROW_FIELDS
  );
  assert.deepEqual(result, [
    { identifyingEmail: 'jizdenky@regiojet.cz', calendarId: 'cal-3@group.calendar.google.com', insertPdfIntoEvent: false, mode: 'ics' },
  ]);
  assert.equal(isValidTicketingPortalsShape(result), true);
});

test('webappBuildJsonRowsValue: a row whose every string-ish field trims to empty is dropped; a row with any non-empty string-ish field is kept', () => {
  const result = webappBuildJsonRowsValue(
    [
      { from: '', calendarId: '   ' },
      { from: 'kept@x.com', calendarId: '' },
    ],
    ICS_ROW_FIELDS
  );
  assert.equal(result.length, 1);
  assert.equal(result[0].from, 'kept@x.com');
});

test('webappBuildJsonRowsValue: kind stringOrNull with "", whitespace, or null yields JSON null; kind string with "" yields "" (never null)', () => {
  const emptyResult = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'x@y.com', calendarId: '', insertPdfIntoEvent: false }],
    PORTALS_ROW_FIELDS
  );
  assert.equal(emptyResult[0].calendarId, null);

  const whitespaceResult = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'x@y.com', calendarId: '   ', insertPdfIntoEvent: false }],
    PORTALS_ROW_FIELDS
  );
  assert.equal(whitespaceResult[0].calendarId, null);

  const nullResult = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'x@y.com', calendarId: null, insertPdfIntoEvent: false }],
    PORTALS_ROW_FIELDS
  );
  assert.equal(nullResult[0].calendarId, null);

  const stringResult = webappBuildJsonRowsValue([{ from: 'nonempty@x.com', calendarId: '' }], ICS_ROW_FIELDS);
  assert.equal(stringResult[0].calendarId, '');
  assert.notEqual(stringResult[0].calendarId, null);
});

test('webappBuildJsonRowsValue: kind boolean coerces true/false/"on"/undefined/null to a real boolean', () => {
  const rowFieldsBoolOnly = [
    { name: 'anchor', kind: 'string' },
    { name: 'flag', kind: 'boolean' },
  ];
  assert.equal(webappBuildJsonRowsValue([{ anchor: 'a', flag: true }], rowFieldsBoolOnly)[0].flag, true);
  assert.equal(webappBuildJsonRowsValue([{ anchor: 'a', flag: false }], rowFieldsBoolOnly)[0].flag, false);
  assert.equal(webappBuildJsonRowsValue([{ anchor: 'a', flag: 'on' }], rowFieldsBoolOnly)[0].flag, true);
  assert.equal(webappBuildJsonRowsValue([{ anchor: 'a', flag: undefined }], rowFieldsBoolOnly)[0].flag, false);
  assert.equal(webappBuildJsonRowsValue([{ anchor: 'a', flag: null }], rowFieldsBoolOnly)[0].flag, false);
});

test('webappBuildJsonRowsValue: kind enum accepts "ics" and "body"; rejects invalid values by throwing, naming the field and the bad value', () => {
  assert.equal(
    webappBuildJsonRowsValue(
      [{ identifyingEmail: 'a@b.com', calendarId: null, insertPdfIntoEvent: false, mode: 'ics' }],
      TRANSPORT_ROW_FIELDS
    )[0].mode,
    'ics'
  );
  assert.equal(
    webappBuildJsonRowsValue(
      [{ identifyingEmail: 'a@b.com', calendarId: null, insertPdfIntoEvent: false, mode: 'body' }],
      TRANSPORT_ROW_FIELDS
    )[0].mode,
    'body'
  );

  ['ICS', 'pdf', 'nonsense'].forEach(function (badValue) {
    assert.throws(function () {
      webappBuildJsonRowsValue(
        [{ identifyingEmail: 'a@b.com', calendarId: null, insertPdfIntoEvent: false, mode: badValue }],
        TRANSPORT_ROW_FIELDS
      );
    }, new RegExp('mode.*' + badValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  });
});

test('webappBuildJsonRowsValue: an enum field declared allowUnset with a submitted value of "" or null OMITS that key entirely (never mode: "")', () => {
  const resultEmpty = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'a@b.com', calendarId: null, insertPdfIntoEvent: false, mode: '' }],
    TRANSPORT_ROW_FIELDS
  );
  assert.equal(Object.prototype.hasOwnProperty.call(resultEmpty[0], 'mode'), false);

  const resultNull = webappBuildJsonRowsValue(
    [{ identifyingEmail: 'a@b.com', calendarId: null, insertPdfIntoEvent: false, mode: null }],
    TRANSPORT_ROW_FIELDS
  );
  assert.equal(Object.prototype.hasOwnProperty.call(resultNull[0], 'mode'), false);
});

test('webappBuildJsonRowsValue: an unrecognized rowField kind THROWS', () => {
  assert.throws(function () {
    webappBuildJsonRowsValue([{ weird: 'x' }], [{ name: 'weird', kind: 'wat' }]);
  }, /weird/);
});

// --- webappCoerceSubmittedSettingValue ---------------------------------------

test('webappCoerceSubmittedSettingValue: boolean -> "true"/"false"', () => {
  const entry = { key: 'fake-bool', type: 'boolean' };
  assert.equal(webappCoerceSubmittedSettingValue(entry, true), 'true');
  assert.equal(webappCoerceSubmittedSettingValue(entry, false), 'false');
});

test('webappCoerceSubmittedSettingValue: number 5 and "5" both -> "5"; a non-numeric value THROWS naming the key', () => {
  const entry = { key: 'fake-number', type: 'number' };
  assert.equal(webappCoerceSubmittedSettingValue(entry, 5), '5');
  assert.equal(webappCoerceSubmittedSettingValue(entry, '5'), '5');
  assert.throws(function () {
    webappCoerceSubmittedSettingValue(entry, 'abc');
  }, /fake-number/);
});

test('webappCoerceSubmittedSettingValue: string passthrough', () => {
  const entry = { key: 'fake-string', type: 'string' };
  assert.equal(webappCoerceSubmittedSettingValue(entry, 'hello'), 'hello');
});

test('webappCoerceSubmittedSettingValue: list accepts both an array and a comma-separated string, output equals formatListSettingValue(parseListSettingValue(input)) exactly', () => {
  const entry = { key: 'fake-list', type: 'list' };
  const arrayInput = ['a@x.com', 'b@y.com'];
  const stringInput = 'a@x.com, b@y.com';
  assert.equal(webappCoerceSubmittedSettingValue(entry, arrayInput), formatListSettingValue(parseListSettingValue(arrayInput)));
  assert.equal(webappCoerceSubmittedSettingValue(entry, stringInput), formatListSettingValue(parseListSettingValue(stringInput)));
});

test('webappCoerceSubmittedSettingValue: json -> formatJsonSettingValue of the rebuilt rows', () => {
  const icsEntry = SETTINGS_REGISTRY.find(function (e) {
    return e.key === '05-action-ics-CALENDAR_ID_BY_SENDER';
  });
  const submittedRows = [{ from: 'a@x.com', calendarId: 'cal-1@group.calendar.google.com' }];
  const result = webappCoerceSubmittedSettingValue(icsEntry, submittedRows);
  const expectedRows = webappBuildJsonRowsValue(submittedRows, icsEntry.jsonRowFields);
  assert.equal(result, formatJsonSettingValue(expectedRows));
});

test("webappCoerceSubmittedSettingValue: json rows that fail the entry's jsonShapeValidator THROW before any string is produced", () => {
  const badEntry = {
    key: 'fake-bad-json',
    type: 'json',
    jsonRowFields: [
      { name: 'from', kind: 'string' },
      { name: 'calendarId', kind: 'string' },
    ],
    jsonShapeValidator: function () {
      return false;
    },
  };
  assert.throws(function () {
    webappCoerceSubmittedSettingValue(badEntry, [{ from: 'a@x.com', calendarId: 'cal@x.com' }]);
  }, /fake-bad-json/);
});

test('webappCoerceSubmittedSettingValue: empty-string constraint -- empty string/list results return the sentinel; a json setting with zero rows returns the literal "[]"', () => {
  assert.equal(webappCoerceSubmittedSettingValue({ key: 'fake-string', type: 'string' }, ''), SETTING_DEFAULT_SENTINEL);
  assert.equal(webappCoerceSubmittedSettingValue({ key: 'fake-list', type: 'list' }, []), SETTING_DEFAULT_SENTINEL);
  assert.equal(webappCoerceSubmittedSettingValue({ key: 'fake-list-2', type: 'list' }, ''), SETTING_DEFAULT_SENTINEL);

  const icsEntry = SETTINGS_REGISTRY.find(function (e) {
    return e.key === '05-action-ics-CALENDAR_ID_BY_SENDER';
  });
  assert.equal(webappCoerceSubmittedSettingValue(icsEntry, []), '[]');
});

test('webappCoerceSubmittedSettingValue: THE round-trip proof -- for each of the 4 json entries, the produced storage string is always readable by the live runtime getter (never falls back to a sentinel default)', () => {
  const UNIQUE_SENTINEL = { uniqueRoundTripSentinel: true };

  const cases = [
    {
      entry: SETTINGS_REGISTRY.find(function (e) {
        return e.key === '05-action-ics-CALENDAR_ID_BY_SENDER';
      }),
      rows: [{ from: 'a@x.com', calendarId: 'cal-1@group.calendar.google.com' }],
    },
    {
      entry: SETTINGS_REGISTRY.find(function (e) {
        return e.key === '07-action-ticketing-portals-TICKETING_PORTALS';
      }),
      rows: [{ identifyingEmail: 'no-reply@enigoo.cz', calendarId: '', insertPdfIntoEvent: true }],
    },
    {
      entry: SETTINGS_REGISTRY.find(function (e) {
        return e.key === '08-action-transport-tickets-TRANSPORT_SENDERS';
      }),
      rows: [{ identifyingEmail: 'jizdenky@regiojet.cz', calendarId: 'cal-2@group.calendar.google.com', insertPdfIntoEvent: false, mode: '' }],
    },
    {
      entry: SETTINGS_REGISTRY.find(function (e) {
        return e.key === '10-action-meetings-MEETING_SYSTEMS';
      }),
      rows: [{ domainPattern: '*.teamio.com', calendarId: '' }],
    },
  ];

  cases.forEach(function (testCase) {
    const stored = webappCoerceSubmittedSettingValue(testCase.entry, testCase.rows);
    const roundTripped = parseJsonSettingValue(stored, UNIQUE_SENTINEL, testCase.entry.jsonShapeValidator);
    assert.notEqual(roundTripped, UNIQUE_SENTINEL);
  });
});

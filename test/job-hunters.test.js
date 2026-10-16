'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

// quick-261016-j0d: the JobHunters capture step is STANDALONE (D-01), so its
// contract is pinned entirely by this one file plus the static wiring
// guards at the bottom. Require at top level so a missing src/ module fails
// the WHOLE file at require time -- acceptable in Task 1 (RED), see that
// task's own <action> note.
const {
  JOB_HUNTERS_COLUMN_NAMES,
  JOB_HUNTERS_BODY_MAX_CHARS,
  JOB_HUNTERS_DRY_RUN_MESSAGE_COUNT,
  JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS,
  JOB_HUNTERS_LOCK_TIMEOUT_MS,
  jobHuntersIsDue,
  jobHuntersDescribeNextRun,
  jobHuntersGetNextRunDescription,
  jobHuntersParseFromHeader,
  jobHuntersNormalizeAddressList,
  cleanJobHunterBodyText,
  jobHuntersIsBulkEmail,
  buildJobHunterRow,
  resolveJobHuntersColumnIndexes,
  jobHuntersColumnSegments,
  jobHuntersRowsForSegment,
  jobHuntersBuildExistingIdSet,
  jobHuntersFindLastDataRow,
  selectLatestJobHunterMessages,
  jobHuntersFormatDryRunRow,
  jobHuntersDescribeRunResult,
  jobHuntersIsSpreadsheetIdConfigured,
  runJobHuntersCapture,
} = require('../src/11-action-job-hunters.js');

const {
  JOB_HUNTERS_ACTION_CONFIG,
  JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY,
  JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER,
} = require('../src/11-action-cfg-job-hunters.js');

// --- Fakes -------------------------------------------------------------
//
// Mirrors test/ticketing-portals.test.js's duck-typed fakeMessage/
// fakeAttachment precedent (D-14), extended with getThread().getId() and a
// recording getHeader(name), since buildJobHunterRow reads exactly 3 raw
// headers and must never call getRawContent (D-06).

function jobHunterFakeAttachment(name) {
  return {
    getName: function () {
      return name;
    },
  };
}

function jobHunterFakeMessage(overrides) {
  const opts = overrides || {};
  const headers = opts.headers || {};
  const headerCallLog = [];
  const attachments =
    opts.attachments !== undefined
      ? opts.attachments
      : [jobHunterFakeAttachment('CV.pdf'), jobHunterFakeAttachment('Cover letter.pdf')];

  return {
    getId: function () {
      return opts.id !== undefined ? opts.id : 'msg-0001';
    },
    getThread: function () {
      const threadId = opts.threadId !== undefined ? opts.threadId : 'thr-0001';
      return {
        getId: function () {
          return threadId;
        },
      };
    },
    getDate: function () {
      return opts.date !== undefined ? opts.date : new Date('2026-10-16T09:30:00.000Z');
    },
    getFrom: function () {
      return opts.from !== undefined ? opts.from : '"Recruiter Example" <Recruiter@Example.invalid>';
    },
    getTo: function () {
      return opts.to !== undefined ? opts.to : 'You@Example.invalid';
    },
    getCc: function () {
      return opts.cc !== undefined ? opts.cc : 'HR-Team@Example.invalid';
    },
    getSubject: function () {
      return opts.subject !== undefined ? opts.subject : 'Interview invitation';
    },
    getPlainBody: function () {
      if (opts.plainBodyThrows) {
        throw new Error('getPlainBody failed');
      }
      return opts.plainBody !== undefined
        ? opts.plainBody
        : "Hi,\n\n> previous quoted line\nWe'd like to schedule an interview.\n\nOn Mon, 12 Oct 2026 at 09:00, Jane Example <jane@example.invalid> wrote:\n> Thanks for applying.";
    },
    getAttachments: function () {
      return attachments;
    },
    getHeader: function (name) {
      headerCallLog.push(name);
      return headers[name] || '';
    },
    getHeaderCallLog: function () {
      return headerCallLog.slice();
    },
    getRawContent: function () {
      throw new Error('getRawContent must never be called');
    },
  };
}

function jobHunterFakeThread(id, lastMessageDate, messages) {
  let sharedLog = null;
  return {
    getId: function () {
      return id;
    },
    getLastMessageDate: function () {
      return lastMessageDate;
    },
    getMessages: function () {
      return messages || [];
    },
    getFirstMessageSubject: function () {
      const msgs = messages || [];
      return msgs.length > 0 ? msgs[0].getSubject() : '';
    },
    removeLabel: function (label) {
      if (sharedLog) {
        sharedLog.push(['removeLabel', id, label.getName()]);
      }
    },
    addLabel: function (label) {
      if (sharedLog) {
        sharedLog.push(['addLabel', id, label.getName()]);
      }
    },
    __attachLog: function (log) {
      sharedLog = log;
    },
  };
}

// The exact 20-name header from D-06, used by every harness test's default
// sheet shape and by the resolveJobHuntersColumnIndexes/segments tests.
const OWNER_HEADER = [
  'id',
  'thread_id',
  'received_at',
  'direction',
  'from_email',
  'from_name',
  'to',
  'cc',
  'subject',
  'body',
  'attachments',
  'is_bulk',
  'type',
  'type_confidence',
  'signals',
  'application_id',
  'job_id',
  'match_method',
  'state',
  'communication_id',
];

function defaultJobHuntersSheetRows() {
  return [
    OWNER_HEADER.slice(),
    [
      'msg-existing-1', 'thr-existing', '2026-10-01T00:00:00.000Z', 'incoming',
      'someone@example.invalid', 'Someone Example', 'you@example.invalid', '',
      'Existing message', 'Existing body', '', false,
      'application', 0.9, '[]', 'app-1', 'job-1', 'manual', 'active', 'comm-1',
    ],
    [
      'msg-0003', 'thr-0003', '2026-10-10T00:00:00.000Z', 'incoming',
      'other@example.invalid', 'Other Example', 'you@example.invalid', '',
      'Another message', 'Another body', '', false,
      'application', 0.5, '[]', 'app-2', 'job-2', 'manual', 'active', 'comm-2',
    ],
  ];
}

function defaultJobHuntersProperties(overrides) {
  return Object.assign(
    {
      '11-action-job-hunters-SPREADSHEET_ID': 'TEST-SPREADSHEET-ID',
      '11-action-job-hunters-OWN_ADDRESSES': 'you@example.invalid',
    },
    overrides || {}
  );
}

// A fake GAS Range + Sheet backing store (TEST HARNESS CONTRACT). getRange
// throws on numRows/numCols < 1 (real GAS behavior); setValues/setNumberFormat
// are both recorded into `calls` AND pushed onto the shared ordered `log`, so
// ordering against label calls is assertable.
//
// quick-261016-nu9 round 2: `maxRowsOverride` simulates the live-reported bug
// -- a real GAS sheet can have MORE rows (getMaxRows()) than rows that
// actually hold data (store.length), e.g. checkbox/data-validation
// formatting pre-applied to rows the owner's other app provisioned ahead of
// time. getRange already returns '' for any cell beyond `store`'s real
// content, so a getMaxRows() larger than store.length is enough to
// reproduce "padding rows read as empty" without modelling real sheet
// formatting at all. Defaults to store.length (no padding), so every
// pre-existing test is unaffected.
function createJobHuntersFakeSheet(sheetRows, log, calls, maxRowsOverride) {
  const store = sheetRows.map(function (row) {
    return row.slice();
  });
  let maxRows = typeof maxRowsOverride === 'number' ? maxRowsOverride : store.length;

  function getLastRow() {
    return store.length;
  }
  function getLastColumn() {
    return store.length > 0 ? store[0].length : 0;
  }
  function getMaxRows() {
    return maxRows;
  }
  function insertRowsAfter(afterRow, howMany) {
    calls.insertRowsAfter.push({ afterRow: afterRow, howMany: howMany });
    log.push(['insertRowsAfter', afterRow, howMany]);
    maxRows += howMany;
  }

  function getRange(row, col, numRows, numCols) {
    if (numRows < 1 || numCols < 1) {
      throw new Error('getRange: numRows and numCols must be >= 1');
    }
    calls.getRange.push({ row: row, col: col, numRows: numRows, numCols: numCols });

    return {
      getValues: function () {
        const result = [];
        for (let r = 0; r < numRows; r++) {
          const sourceRow = store[row - 1 + r] || [];
          const values = [];
          for (let c = 0; c < numCols; c++) {
            values.push(sourceRow[col - 1 + c] !== undefined ? sourceRow[col - 1 + c] : '');
          }
          result.push(values);
        }
        return result;
      },
      setValues: function (values) {
        calls.setValues.push({ row: row, col: col, numRows: numRows, numCols: numCols, values: values });
        log.push(['setValues', row, col, numRows, numCols]);
        for (let r = 0; r < values.length; r++) {
          const targetRowIndex = row - 1 + r;
          while (store.length <= targetRowIndex) {
            store.push([]);
          }
          for (let c = 0; c < values[r].length; c++) {
            store[targetRowIndex][col - 1 + c] = values[r][c];
          }
        }
      },
      setNumberFormat: function (format) {
        calls.setNumberFormat.push({ row: row, col: col, numRows: numRows, numCols: numCols, format: format });
        log.push(['setNumberFormat', row, col, numRows, numCols, format]);
      },
    };
  }

  return {
    getLastRow: getLastRow,
    getLastColumn: getLastColumn,
    getMaxRows: getMaxRows,
    insertRowsAfter: insertRowsAfter,
    getRange: getRange,
    __getStore: function () {
      return store;
    },
  };
}

// withJobHuntersRunGlobals -- the TEST HARNESS CONTRACT. Injects
// PropertiesService/LockService/GmailApp/SpreadsheetApp/getOrCreateLabel/
// notifyOwnerOfFailure as globals, saving/restoring (or deleting) each in
// `finally`, and hands `fn` a context exposing the shared ordered `log` and
// per-call recordings.
function withJobHuntersRunGlobals(options, fn) {
  const opts = options || {};
  const log = [];
  const calls = {
    propertyWrites: [],
    getScriptLockCalls: 0,
    tryLockMs: [],
    releaseLockCalls: 0,
    searchQueries: [],
    openByIdCalls: [],
    getUserLabelByNameCalls: [],
    getOrCreateLabelCalls: [],
    notifyOwnerOfFailureCalls: [],
    getRange: [],
    setValues: [],
    setNumberFormat: [],
    insertRowsAfter: [],
  };

  const propertiesStore = Object.assign({}, opts.properties || {});

  const keysToSet = ['PropertiesService', 'LockService', 'GmailApp', 'SpreadsheetApp', 'getOrCreateLabel', 'notifyOwnerOfFailure'];
  const savedGlobals = {};
  keysToSet.forEach(function (key) {
    savedGlobals[key] = { existed: Object.prototype.hasOwnProperty.call(global, key), value: global[key] };
  });

  global.PropertiesService = {
    getScriptProperties: function () {
      return {
        getProperty: function (key) {
          return Object.prototype.hasOwnProperty.call(propertiesStore, key) ? propertiesStore[key] : null;
        },
        setProperty: function (key, value) {
          propertiesStore[key] = value;
          calls.propertyWrites.push({ key: key, value: value });
        },
        getProperties: function () {
          return Object.assign({}, propertiesStore);
        },
      };
    },
  };

  global.LockService = {
    getScriptLock: function () {
      calls.getScriptLockCalls++;
      return {
        tryLock: function (ms) {
          calls.tryLockMs.push(ms);
          return opts.lockAcquired !== false;
        },
        releaseLock: function () {
          calls.releaseLockCalls++;
        },
      };
    },
  };

  const fakeLabelsByName = {};
  function getFakeLabel(name) {
    if (!fakeLabelsByName[name]) {
      fakeLabelsByName[name] = { getName: function () { return name; } };
    }
    return fakeLabelsByName[name];
  }

  (opts.threads || []).forEach(function (thread) {
    if (thread && typeof thread.__attachLog === 'function') {
      thread.__attachLog(log);
    }
  });

  global.GmailApp = {
    search: function (q) {
      calls.searchQueries.push(q);
      return opts.threads || [];
    },
    getUserLabelByName: function (name) {
      calls.getUserLabelByNameCalls.push(name);
      return getFakeLabel(name);
    },
  };

  let fakeSheet = null;
  if (!opts.sheetMissing && opts.sheetRows) {
    fakeSheet = createJobHuntersFakeSheet(opts.sheetRows, log, calls, opts.sheetMaxRows);
  }

  global.SpreadsheetApp = {
    openById: function (id) {
      calls.openByIdCalls.push(id);
      return {
        getSheetByName: function () {
          return opts.sheetMissing ? null : fakeSheet;
        },
      };
    },
  };

  global.getOrCreateLabel = function (name) {
    calls.getOrCreateLabelCalls.push(name);
    return getFakeLabel(name);
  };

  global.notifyOwnerOfFailure = function (actionName, thread, error) {
    calls.notifyOwnerOfFailureCalls.push({ actionName: actionName, thread: thread, error: error });
  };

  try {
    return fn({
      log: log,
      calls: calls,
      getProperties: function () {
        return Object.assign({}, propertiesStore);
      },
      getSheetStore: function () {
        return fakeSheet ? fakeSheet.__getStore() : null;
      },
    });
  } finally {
    keysToSet.forEach(function (key) {
      if (savedGlobals[key].existed) {
        global[key] = savedGlobals[key].value;
      } else {
        delete global[key];
      }
    });
  }
}

// --- jobHuntersIsDue (D-04, P-7) --------------------------------------
//
// Pure due-check gate: sanitizes all three inputs so a corrupt/garbage
// stored value can never block runs forever (P-7), and a future-dated last
// run or a non-finite "now" always counts as due.

test('quick-261016-j0d (D-04, P-7): jobHuntersIsDue is true at exactly interval*60000 elapsed, false 1ms earlier', () => {
  assert.equal(jobHuntersIsDue(1000000, 60, 4600000), true);
  assert.equal(jobHuntersIsDue(1000000, 60, 4599999), false);
});

test('quick-261016-j0d (D-04, P-7): jobHuntersIsDue treats lastRunAtMs 0 as due', () => {
  assert.equal(jobHuntersIsDue(0, 60, 1760000000000), true);
});

test('quick-261016-j0d (P-7): jobHuntersIsDue treats garbage/negative last as 0 (always due against a realistic "now")', () => {
  ['abc', NaN, -5, undefined, null].forEach(function (last) {
    assert.equal(jobHuntersIsDue(last, 60, 1760000000000), true);
  });
});

test('quick-261016-j0d (P-7): jobHuntersIsDue accepts string inputs via Number() coercion', () => {
  assert.equal(jobHuntersIsDue('1000000', '60', 4600000), true);
});

test('quick-261016-j0d (P-7): jobHuntersIsDue treats negative/garbage interval as 0', () => {
  assert.equal(jobHuntersIsDue(0, -10, 1000), true);
  assert.equal(jobHuntersIsDue(0, 'abc', 1000), true);
});

test('quick-261016-j0d (P-7): jobHuntersIsDue treats a future last-run as due', () => {
  assert.equal(jobHuntersIsDue(5000000, 60, 1000000), true);
});

test('quick-261016-j0d (P-7): jobHuntersIsDue treats a non-finite "now" as due', () => {
  assert.equal(jobHuntersIsDue(0, 60, NaN), true);
  assert.equal(jobHuntersIsDue(0, 60, Infinity), true);
});

// --- jobHuntersDescribeNextRun (quick-261016-nu9 round 3) --------------
//
// Built directly on jobHuntersIsDue, so the "due now" branch here must
// agree with jobHuntersIsDue's own verdict for every input -- these tests
// pin the human-readable wording, not a second due-check implementation.

test('quick-261016-nu9: jobHuntersDescribeNextRun -- never run yet (lastRunAtMs 0) says so, not just "due now"', () => {
  assert.equal(jobHuntersDescribeNextRun(0, 60, 1760000000000), 'Never run yet -- due now.');
});

test('quick-261016-nu9: jobHuntersDescribeNextRun -- overdue (past the interval) says "Due now."', () => {
  assert.equal(jobHuntersDescribeNextRun(1000000, 60, 4600000), 'Due now.');
});

test('quick-261016-nu9: jobHuntersDescribeNextRun -- not yet due gives the exact next-due instant and a rounded-up minute count', () => {
  // last=1000000, interval=60 -> nextRunAtMs=1000000+3600000=4600000. At
  // now=4600000-90000 (90s early), 90000ms remain -> ceil(90000/60000)=2.
  assert.equal(
    jobHuntersDescribeNextRun(1000000, 60, 4600000 - 90000),
    'Next due ' + new Date(4600000).toISOString() + ' (in about 2 minutes).'
  );
});

test('quick-261016-nu9: jobHuntersDescribeNextRun -- singular "minute" wording at exactly 1 minute remaining', () => {
  assert.equal(
    jobHuntersDescribeNextRun(1000000, 60, 4600000 - 60000),
    'Next due ' + new Date(4600000).toISOString() + ' (in about 1 minute).'
  );
});

test('quick-261016-nu9: jobHuntersDescribeNextRun -- a future last-run is due now, same as jobHuntersIsDue', () => {
  assert.equal(jobHuntersDescribeNextRun(5000000, 60, 1000000), 'Due now.');
});

test('quick-261016-nu9: jobHuntersDescribeNextRun -- garbage interval is due now, same as jobHuntersIsDue', () => {
  assert.equal(jobHuntersDescribeNextRun(1000000, 'abc', 1760000000000), 'Due now.');
});

// --- jobHuntersGetNextRunDescription (GAS-only read, harness test) -----

test('quick-261016-nu9: jobHuntersGetNextRunDescription reads the real LAST_RUN_AT_MS property and searchIntervalMinutes setting', () => {
  const lastRunAtMs = 1000000;
  withJobHuntersRunGlobals(
    {
      properties: defaultJobHuntersProperties({
        '11-action-job-hunters-SEARCH_INTERVAL_MINUTES': '60',
        [JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY]: String(lastRunAtMs),
      }),
    },
    () => {
      const before = Date.now();
      const message = jobHuntersGetNextRunDescription();
      assert.equal(message, 'Due now.');
      assert.ok(Date.now() >= before);
    }
  );
});

test('quick-261016-nu9: jobHuntersGetNextRunDescription -- LAST_RUN_AT_MS absent (never run) says so', () => {
  withJobHuntersRunGlobals({ properties: defaultJobHuntersProperties() }, () => {
    assert.equal(jobHuntersGetNextRunDescription(), 'Never run yet -- due now.');
  });
});

// --- jobHuntersParseFromHeader (D-06, P-1) ----------------------------

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- quoted name', () => {
  assert.deepEqual(
    jobHuntersParseFromHeader('"Jane Example" <jane@example.invalid>'),
    { name: 'Jane Example', email: 'jane@example.invalid' }
  );
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- unquoted name', () => {
  assert.deepEqual(
    jobHuntersParseFromHeader('Jane Example <jane@example.invalid>'),
    { name: 'Jane Example', email: 'jane@example.invalid' }
  );
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- bare address, no name', () => {
  assert.deepEqual(jobHuntersParseFromHeader('jane@example.invalid'), { name: '', email: 'jane@example.invalid' });
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- upper-case email is lowercased', () => {
  assert.deepEqual(
    jobHuntersParseFromHeader('"Jane Example" <JANE@EXAMPLE.INVALID>'),
    { name: 'Jane Example', email: 'jane@example.invalid' }
  );
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- comma inside the quoted name', () => {
  assert.deepEqual(
    jobHuntersParseFromHeader('"Example, Jane" <jane@example.invalid>'),
    { name: 'Example, Jane', email: 'jane@example.invalid' }
  );
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- the LAST angle pair wins when the name itself contains angle brackets', () => {
  assert.deepEqual(
    jobHuntersParseFromHeader('"Jane <HR team>" <jane@example.invalid>'),
    { name: 'Jane <HR team>', email: 'jane@example.invalid' }
  );
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- angle pair with no preceding name', () => {
  assert.deepEqual(jobHuntersParseFromHeader('<jane@example.invalid>'), { name: '', email: 'jane@example.invalid' });
});

test('quick-261016-j0d (D-06): jobHuntersParseFromHeader -- null/empty/undefined all give {name:"", email:""}', () => {
  assert.deepEqual(jobHuntersParseFromHeader(null), { name: '', email: '' });
  assert.deepEqual(jobHuntersParseFromHeader(''), { name: '', email: '' });
  assert.deepEqual(jobHuntersParseFromHeader(undefined), { name: '', email: '' });
});

// --- jobHuntersNormalizeAddressList ------------------------------------

test('quick-261016-j0d: jobHuntersNormalizeAddressList trims, joins with ", " and lowercases', () => {
  assert.equal(
    jobHuntersNormalizeAddressList('Jane <Jane@Example.invalid>,Bob@Example.invalid'),
    'jane <jane@example.invalid>, bob@example.invalid'
  );
});

test('quick-261016-j0d: jobHuntersNormalizeAddressList preserves order and duplicates', () => {
  assert.equal(
    jobHuntersNormalizeAddressList('bob@example.invalid, bob@example.invalid, alice@example.invalid'),
    'bob@example.invalid, bob@example.invalid, alice@example.invalid'
  );
});

test('quick-261016-j0d: jobHuntersNormalizeAddressList null/"" give ""', () => {
  assert.equal(jobHuntersNormalizeAddressList(null), '');
  assert.equal(jobHuntersNormalizeAddressList(''), '');
});

// --- cleanJobHunterBodyText (D-06, P-5) --------------------------------

test('quick-261016-j0d (D-06): cleanJobHunterBodyText drops ">" lines, keeps surrounding lines', () => {
  assert.equal(cleanJobHunterBodyText('Hello\n> quoted\nWorld'), 'Hello\nWorld');
});

test('quick-261016-j0d (D-06, P-5): cleanJobHunterBodyText cuts at an English "On ... wrote:" line', () => {
  const result = cleanJobHunterBodyText(
    'Keep this.\nOn Mon, 12 Oct 2026 at 09:00, Jane Example <jane@example.invalid> wrote:\nDrop this.'
  );
  assert.equal(result, 'Keep this.');
});

test('quick-261016-j0d (D-06, P-5): cleanJobHunterBodyText cuts at a Czech "Dne ... napsal(a):" line', () => {
  const result = cleanJobHunterBodyText(
    'Keep this.\nDne po 12. 10. 2026 v 9:00 odesílatel Jana Example <jana@example.invalid> napsal(a):\nDrop this.'
  );
  assert.equal(result, 'Keep this.');
});

test('quick-261016-j0d (D-06, P-5): cleanJobHunterBodyText cuts at a "-----Original Message-----" line', () => {
  const result = cleanJobHunterBodyText('Keep this.\n-----Original Message-----\nDrop this.');
  assert.equal(result, 'Keep this.');
});

test('quick-261016-j0d (P-5): cleanJobHunterBodyText -- "-- " (trailing space) cuts the signature, a bare "--" does not', () => {
  assert.equal(cleanJobHunterBodyText('Keep this.\n-- \nSignature dropped.'), 'Keep this.');
  assert.equal(cleanJobHunterBodyText('Keep this.\n--\nKept too.'), 'Keep this.\n--\nKept too.');
});

test('quick-261016-j0d (P-5): cleanJobHunterBodyText normalizes CRLF/CR to \\n, and the cut still works on a CRLF body', () => {
  const result = cleanJobHunterBodyText(
    'Keep this.\r\nOn Mon, 12 Oct 2026 at 09:00, Jane Example <jane@example.invalid> wrote:\r\nDrop this.\r'
  );
  assert.equal(result, 'Keep this.');
});

test('quick-261016-j0d: cleanJobHunterBodyText slices to exactly JOB_HUNTERS_BODY_MAX_CHARS', () => {
  const longBody = 'x'.repeat(10050);
  const result = cleanJobHunterBodyText(longBody);
  assert.equal(result.length, JOB_HUNTERS_BODY_MAX_CHARS);
  assert.equal(result.length, 10000);
});

test('quick-261016-j0d: cleanJobHunterBodyText on a marker-free body is trimmed and otherwise identical', () => {
  assert.equal(cleanJobHunterBodyText('  Hello world  '), 'Hello world');
});

test('quick-261016-j0d: cleanJobHunterBodyText -- a marker phrase mid-line (not starting the line) is NOT cut', () => {
  const body = 'He said On Monday he wrote: something else\nSecond line.';
  assert.equal(cleanJobHunterBodyText(body), body);
});

test('quick-261016-j0d: cleanJobHunterBodyText null gives ""', () => {
  assert.equal(cleanJobHunterBodyText(null), '');
});

// --- jobHuntersIsBulkEmail (D-06, P-6) ---------------------------------

test('quick-261016-j0d (D-06): jobHuntersIsBulkEmail -- List-Unsubscribe non-empty is true, whitespace-only is false', () => {
  assert.equal(jobHuntersIsBulkEmail('<mailto:unsub@example.invalid>', '', '', 'someone@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('   ', '', '', 'someone@example.invalid'), false);
});

test('quick-261016-j0d (D-06): jobHuntersIsBulkEmail -- Precedence bulk/list variants are true, junk is false', () => {
  assert.equal(jobHuntersIsBulkEmail('', 'bulk', '', 'someone@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', 'LIST', '', 'someone@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', ' Bulk ', '', 'someone@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', 'junk', '', 'someone@example.invalid'), false);
});

test('quick-261016-j0d (D-06): jobHuntersIsBulkEmail -- Auto-Submitted auto-generated is true, no/No/"" are false', () => {
  assert.equal(jobHuntersIsBulkEmail('', '', 'auto-generated', 'someone@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', '', 'no', 'someone@example.invalid'), false);
  assert.equal(jobHuntersIsBulkEmail('', '', 'No', 'someone@example.invalid'), false);
  assert.equal(jobHuntersIsBulkEmail('', '', '', 'someone@example.invalid'), false);
});

test('quick-261016-j0d (D-06, P-6): jobHuntersIsBulkEmail -- noreply/no-reply/donotreply local-parts are true (exact match), jobs-noreply and reply are false', () => {
  assert.equal(jobHuntersIsBulkEmail('', '', '', 'noreply@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', '', '', 'No-Reply@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', '', '', 'DoNotReply@example.invalid'), true);
  assert.equal(jobHuntersIsBulkEmail('', '', '', 'jobs-noreply@example.invalid'), false);
  assert.equal(jobHuntersIsBulkEmail('', '', '', 'reply@example.invalid'), false);
});

test('quick-261016-j0d: jobHuntersIsBulkEmail -- all-empty/null arguments give false without throwing, result is always typeof boolean', () => {
  const result = jobHuntersIsBulkEmail(null, null, null, null);
  assert.equal(result, false);
  assert.equal(typeof result, 'boolean');
  assert.equal(typeof jobHuntersIsBulkEmail('', '', '', ''), 'boolean');
});

// --- JOB_HUNTERS_COLUMN_NAMES / buildJobHunterRow (D-06) ---------------

test('quick-261016-j0d (D-06): JOB_HUNTERS_COLUMN_NAMES is the 12 names id..is_bulk, in order', () => {
  assert.deepEqual(JOB_HUNTERS_COLUMN_NAMES, [
    'id', 'thread_id', 'received_at', 'direction', 'from_email', 'from_name',
    'to', 'cc', 'subject', 'body', 'attachments', 'is_bulk',
  ]);
});

test('quick-261016-j0d (D-06): buildJobHunterRow on the default fake produces the exact expected 12-value row', () => {
  const message = jobHunterFakeMessage({});
  const row = buildJobHunterRow(message);

  assert.deepEqual(row, [
    'msg-0001',
    'thr-0001',
    '2026-10-16T09:30:00.000Z',
    'incoming',
    'recruiter@example.invalid',
    'Recruiter Example',
    'you@example.invalid',
    'hr-team@example.invalid',
    'Interview invitation',
    "Hi,\n\nWe'd like to schedule an interview.",
    'CV.pdf | Cover letter.pdf',
    false,
  ]);
});

test('quick-261016-j0d (D-06): buildJobHunterRow -- is_bulk is strictEqual boolean true when List-Unsubscribe is set', () => {
  const message = jobHunterFakeMessage({ headers: { 'List-Unsubscribe': '<mailto:unsub@example.invalid>' } });
  const row = buildJobHunterRow(message);
  assert.strictEqual(row[11], true);
});

test('quick-261016-j0d (D-06): buildJobHunterRow -- attachments is "" with no attachments', () => {
  const message = jobHunterFakeMessage({ attachments: [] });
  const row = buildJobHunterRow(message);
  assert.equal(row[10], '');
});

test('quick-261016-j0d (D-06): buildJobHunterRow -- direction is "outgoing" when from matches an ownAddresses entry', () => {
  const message = jobHunterFakeMessage({ from: 'you@example.invalid' });
  const row = buildJobHunterRow(message, ['You@Example.invalid']);
  assert.equal(row[3], 'outgoing');
});

test('quick-261016-j0d (D-06): buildJobHunterRow -- ownAddresses undefined gives "incoming"', () => {
  const message = jobHunterFakeMessage({});
  const row = buildJobHunterRow(message, undefined);
  assert.equal(row[3], 'incoming');
});

test('quick-261016-j0d (D-06): buildJobHunterRow -- reads exactly List-Unsubscribe/Precedence/Auto-Submitted, never getRawContent', () => {
  const message = jobHunterFakeMessage({});
  buildJobHunterRow(message);
  assert.deepEqual(message.getHeaderCallLog(), ['List-Unsubscribe', 'Precedence', 'Auto-Submitted']);
  // getRawContent throws if ever called -- buildJobHunterRow not throwing proves it wasn't.
});

// --- resolveJobHuntersColumnIndexes (D-06) -----------------------------

test('quick-261016-j0d (D-06): resolveJobHuntersColumnIndexes on OWNER_HEADER gives id 0 .. is_bulk 11', () => {
  const indexes = resolveJobHuntersColumnIndexes(OWNER_HEADER);
  assert.equal(indexes.id, 0);
  assert.equal(indexes.thread_id, 1);
  assert.equal(indexes.received_at, 2);
  assert.equal(indexes.direction, 3);
  assert.equal(indexes.from_email, 4);
  assert.equal(indexes.from_name, 5);
  assert.equal(indexes.to, 6);
  assert.equal(indexes.cc, 7);
  assert.equal(indexes.subject, 8);
  assert.equal(indexes.body, 9);
  assert.equal(indexes.attachments, 10);
  assert.equal(indexes.is_bulk, 11);
});

test('quick-261016-j0d (D-06): resolveJobHuntersColumnIndexes resolves correctly regardless of order, with foreign columns interleaved', () => {
  const shuffled = ['foo', 'is_bulk', 'id', 'bar', 'attachments', 'cc', 'to', 'baz', 'subject', 'body', 'from_name', 'from_email', 'direction', 'thread_id', 'received_at'];
  const indexes = resolveJobHuntersColumnIndexes(shuffled);
  assert.equal(indexes.is_bulk, 1);
  assert.equal(indexes.id, 2);
  assert.equal(indexes.attachments, 4);
  assert.equal(indexes.cc, 5);
  assert.equal(indexes.to, 6);
  assert.equal(indexes.subject, 8);
  assert.equal(indexes.body, 9);
  assert.equal(indexes.from_name, 10);
  assert.equal(indexes.from_email, 11);
  assert.equal(indexes.direction, 12);
  assert.equal(indexes.thread_id, 13);
  assert.equal(indexes.received_at, 14);
});

test('quick-261016-j0d (D-06): resolveJobHuntersColumnIndexes trims cell text (" id " matches "id")', () => {
  const header = OWNER_HEADER.slice();
  header[0] = ' id ';
  const indexes = resolveJobHuntersColumnIndexes(header);
  assert.equal(indexes.id, 0);
});

test('quick-261016-j0d (D-06): resolveJobHuntersColumnIndexes throws naming both missing columns', () => {
  const header = OWNER_HEADER.filter(function (name) {
    return name !== 'is_bulk' && name !== 'cc';
  });
  assert.throws(function () {
    resolveJobHuntersColumnIndexes(header);
  }, function (err) {
    return err.message.includes('is_bulk') && err.message.includes('cc');
  });
});

test('quick-261016-j0d (D-06): resolveJobHuntersColumnIndexes throws on an empty header', () => {
  assert.throws(function () {
    resolveJobHuntersColumnIndexes([]);
  });
});

// --- jobHuntersColumnSegments / jobHuntersRowsForSegment (P-3) ---------

test('quick-261016-j0d (P-3): jobHuntersColumnSegments on OWNER_HEADER gives one segment covering all 12 fields at startColumn 1', () => {
  const indexes = resolveJobHuntersColumnIndexes(OWNER_HEADER);
  const segments = jobHuntersColumnSegments(indexes);
  assert.deepEqual(segments, [{ startColumn: 1, fields: JOB_HUNTERS_COLUMN_NAMES.slice() }]);
});

test('quick-261016-j0d (P-3): jobHuntersColumnSegments shifts startColumn when 2 foreign leading columns precede the block', () => {
  const header = ['foo', 'bar'].concat(OWNER_HEADER);
  const indexes = resolveJobHuntersColumnIndexes(header);
  const segments = jobHuntersColumnSegments(indexes);
  assert.equal(segments.length, 1);
  assert.equal(segments[0].startColumn, 3);
});

test('quick-261016-j0d (P-3): a foreign column between body and attachments produces two segments, neither covering it', () => {
  const header = ['id', 'thread_id', 'received_at', 'direction', 'from_email', 'from_name', 'to', 'cc', 'subject', 'body', 'notes', 'attachments', 'is_bulk'];
  const indexes = resolveJobHuntersColumnIndexes(header);
  const segments = jobHuntersColumnSegments(indexes);

  assert.equal(segments.length, 2);
  segments.forEach(function (segment) {
    const colSpan = [segment.startColumn, segment.startColumn + segment.fields.length - 1];
    assert.equal(colSpan[0] <= 11 && 11 <= colSpan[1], false, 'segment must not cover the foreign "notes" column (11)');
  });
});

test('quick-261016-j0d (P-3): jobHuntersRowsForSegment picks exactly that segment\'s fields, in segment order', () => {
  const segment = { startColumn: 12, fields: ['attachments', 'is_bulk'] };
  const rows = [
    ['id-1', 'tid-1', 'rec-1', 'dir-1', 'fe-1', 'fn-1', 'to-1', 'cc-1', 'sub-1', 'body-1', 'att-1', true],
  ];
  assert.deepEqual(jobHuntersRowsForSegment(rows, segment), [['att-1', true]]);
});

// --- jobHuntersBuildExistingIdSet --------------------------------------

test('quick-261016-j0d: jobHuntersBuildExistingIdSet trims, skips empty/null, and stringifies', () => {
  const set = jobHuntersBuildExistingIdSet([['a'], [''], [null], [' b '], [123]]);
  assert.deepEqual(Array.from(set).sort(), ['123', 'a', 'b']);
});

test('quick-261016-j0d: jobHuntersBuildExistingIdSet on [] gives an empty Set', () => {
  assert.equal(jobHuntersBuildExistingIdSet([]).size, 0);
});

// --- jobHuntersFindLastDataRow (quick-261016-nu9 round 2 fix) ----------
//
// Live-reported bug: sheet.getLastRow() counts a checkbox-formatted but
// otherwise empty cell as non-empty, so on the owner's real sheet (checkbox
// data validation pre-applied to is_bulk on rows 2-1000) new rows landed at
// row 1001 instead of row 2. jobHuntersFindLastDataRow instead finds the
// last row whose id CELL VALUE is non-empty after trimming -- formatting
// alone can never satisfy that.

test('quick-261016-nu9: jobHuntersFindLastDataRow on no data rows at all gives 1 (header row; next write is row 2)', () => {
  assert.equal(jobHuntersFindLastDataRow([]), 1);
});

test('quick-261016-nu9: jobHuntersFindLastDataRow -- all-empty-string data rows give 1, same as no rows', () => {
  assert.equal(jobHuntersFindLastDataRow([[''], [''], ['']]), 1);
});

test('quick-261016-nu9: jobHuntersFindLastDataRow -- two real rows then 997 checkbox-but-empty rows gives sheet row 3', () => {
  const rows = [['msg-existing-1'], ['msg-0003']];
  for (let i = 0; i < 997; i++) {
    rows.push(['']);
  }
  assert.equal(jobHuntersFindLastDataRow(rows), 3);
});

test('quick-261016-nu9: jobHuntersFindLastDataRow -- null and undefined cells count as empty, not as data', () => {
  assert.equal(jobHuntersFindLastDataRow([['msg-1'], [null], [undefined], ['']]), 2);
});

test('quick-261016-nu9: jobHuntersFindLastDataRow -- whitespace-only id trims to empty', () => {
  assert.equal(jobHuntersFindLastDataRow([['msg-1'], ['   ']]), 2);
});

test('quick-261016-nu9: jobHuntersFindLastDataRow on [] (no rows passed at all) gives 1', () => {
  assert.equal(jobHuntersFindLastDataRow(undefined), 1);
});

// --- selectLatestJobHunterMessages (D-09, P-16) ------------------------

test('quick-261016-j0d (D-09, P-16): selectLatestJobHunterMessages returns the 3 newest across threads, newest-first', () => {
  const m1 = jobHunterFakeMessage({ id: 'm1', date: new Date('2026-10-01T00:00:00Z') });
  const m2 = jobHunterFakeMessage({ id: 'm2', date: new Date('2026-10-02T00:00:00Z') });
  const m3 = jobHunterFakeMessage({ id: 'm3', date: new Date('2026-10-03T00:00:00Z') });
  const m4 = jobHunterFakeMessage({ id: 'm4', date: new Date('2026-10-04T00:00:00Z') });
  const threadA = jobHunterFakeThread('a', m2.getDate(), [m1, m2]);
  const threadB = jobHunterFakeThread('b', m4.getDate(), [m3, m4]);

  const result = selectLatestJobHunterMessages([threadA, threadB], 3);
  assert.deepEqual(result.map(function (m) { return m.getId(); }), ['m4', 'm3', 'm2']);
});

test('quick-261016-j0d (D-09): selectLatestJobHunterMessages returns both when only 2 messages exist', () => {
  const m1 = jobHunterFakeMessage({ id: 'm1', date: new Date('2026-10-01T00:00:00Z') });
  const m2 = jobHunterFakeMessage({ id: 'm2', date: new Date('2026-10-02T00:00:00Z') });
  const threadA = jobHunterFakeThread('a', m2.getDate(), [m1, m2]);

  const result = selectLatestJobHunterMessages([threadA], 3);
  assert.deepEqual(result.map(function (m) { return m.getId(); }), ['m2', 'm1']);
});

test('quick-261016-j0d (D-09): selectLatestJobHunterMessages sorts an invalid-date message last', () => {
  const m1 = jobHunterFakeMessage({ id: 'm1', date: new Date('2026-10-01T00:00:00Z') });
  const bad = jobHunterFakeMessage({ id: 'bad', date: new Date('not-a-date') });
  const threadA = jobHunterFakeThread('a', m1.getDate(), [bad, m1]);

  const result = selectLatestJobHunterMessages([threadA], 2);
  assert.deepEqual(result.map(function (m) { return m.getId(); }), ['m1', 'bad']);
});

test('quick-261016-j0d (D-09): selectLatestJobHunterMessages count 0 gives []', () => {
  const m1 = jobHunterFakeMessage({ id: 'm1' });
  const threadA = jobHunterFakeThread('a', m1.getDate(), [m1]);
  assert.deepEqual(selectLatestJobHunterMessages([threadA], 0), []);
});

// --- jobHuntersFormatDryRunRow (D-09) ----------------------------------

test('quick-261016-j0d (D-09): jobHuntersFormatDryRunRow -- 12 lines, body truncated to 200 chars, subject complete, is_bulk formatted', () => {
  const longBody = 'b'.repeat(300);
  const longSubject = 's'.repeat(300);
  const row = ['id-1', 'tid-1', 'rec-1', 'incoming', 'fe@example.invalid', 'fn', 'to@example.invalid', 'cc@example.invalid', longSubject, longBody, 'att-1', false];
  const lines = jobHuntersFormatDryRunRow(row);

  assert.equal(lines.length, 12);
  assert.equal(lines[9], 'body: ' + longBody.slice(0, 200));
  assert.equal(lines[8], 'subject: ' + longSubject);
  assert.equal(lines[11], 'is_bulk: false');
});

// --- jobHuntersDescribeRunResult (P-15) --------------------------------

test('quick-261016-j0d (P-15): jobHuntersDescribeRunResult -- the two exact result strings', () => {
  assert.equal(jobHuntersDescribeRunResult({ ran: false, reason: 'lock busy' }), 'JobHunters did not run: lock busy.');
  assert.equal(
    jobHuntersDescribeRunResult({ ran: true, rowsWritten: 3, threadsSucceeded: 2, threadsFailed: 1 }),
    'JobHunters ran: 3 row(s) written, 2 thread(s) relabeled, 1 thread(s) failed.'
  );
});

// --- jobHuntersIsSpreadsheetIdConfigured (P-10) ------------------------

test('quick-261016-j0d (P-10): jobHuntersIsSpreadsheetIdConfigured', () => {
  assert.equal(jobHuntersIsSpreadsheetIdConfigured(JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER), false);
  assert.equal(jobHuntersIsSpreadsheetIdConfigured(''), false);
  assert.equal(jobHuntersIsSpreadsheetIdConfigured('   '), false);
  assert.equal(jobHuntersIsSpreadsheetIdConfigured(null), false);
  assert.equal(jobHuntersIsSpreadsheetIdConfigured('TEST-SPREADSHEET-ID'), true);
});

// --- JOB_HUNTERS_ACTION_CONFIG / key constants (D-02) ------------------

test('quick-261016-j0d (D-02): JOB_HUNTERS_ACTION_CONFIG.spreadsheetId defaults to the committed sentinel under Node', () => {
  assert.strictEqual(JOB_HUNTERS_ACTION_CONFIG.spreadsheetId, JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER);
});

test('quick-261016-j0d (D-02, D-03): JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY is the exact internal-state key', () => {
  assert.equal(JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY, '11-action-job-hunters-LAST_RUN_AT_MS');
});

// --- runJobHuntersCapture harness tests (D-03, D-04, D-05, D-07) -------
//
// Exercises the real GAS-only entry point end-to-end against the injected
// fakes, proving dedup, the lock/due gates, the write-before-relabel
// ordering, the per-thread failure isolation (P-11/P-12), and the
// fail-closed column resolution (D-06) -- all without a single real GAS
// global.

test('quick-261016-j0d (D-05, D-07): happy path -- oldest-first dedup, single batched write, relabel after write, formula-injection guard', () => {
  const msgA1 = jobHunterFakeMessage({ id: 'msg-0001', threadId: 'thr-A', from: '"Recruiter A" <recruiterA@example.invalid>', attachments: [] });
  const msgA2 = jobHunterFakeMessage({ id: 'msg-0002', threadId: 'thr-A', from: 'you@example.invalid', attachments: [] });
  const threadA = jobHunterFakeThread('thr-A', new Date('2026-10-10T08:00:00Z'), [msgA1, msgA2]);

  const msgB1 = jobHunterFakeMessage({ id: 'msg-0003', threadId: 'thr-B', from: '"Dup Sender" <dup@example.invalid>', attachments: [] });
  const threadB = jobHunterFakeThread('thr-B', new Date('2026-10-15T08:00:00Z'), [msgB1]);

  const timeBeforeCall = Date.now();

  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetRows: defaultJobHuntersSheetRows(), threads: [threadB, threadA] },
    (ctx) => {
      const result = runJobHuntersCapture(false);

      assert.deepEqual(result, { ran: true, rowsWritten: 2, threadsSucceeded: 2, threadsFailed: 0 });
      assert.equal(ctx.calls.searchQueries[0], 'label:"jobHunter"');
      assert.equal(ctx.calls.tryLockMs[0], 30000);

      assert.equal(ctx.calls.setValues.length, 1);
      const writeCall = ctx.calls.setValues[0];
      assert.equal(writeCall.row, 4);
      assert.equal(writeCall.col, 1);
      assert.equal(writeCall.values.length, 2);
      assert.equal(writeCall.values[0][0], 'msg-0001');
      assert.equal(writeCall.values[1][0], 'msg-0002');
      assert.equal(writeCall.values[1][3], 'outgoing');
      assert.equal(typeof writeCall.values[1][11], 'boolean');

      const removeLabelEntries = ctx.log.filter(function (e) { return e[0] === 'removeLabel'; });
      const addLabelEntries = ctx.log.filter(function (e) { return e[0] === 'addLabel'; });
      assert.equal(removeLabelEntries.length, 2);
      assert.equal(addLabelEntries.length, 2);

      const setValuesLogIndex = ctx.log.findIndex(function (e) { return e[0] === 'setValues'; });
      ctx.log.forEach(function (entry, idx) {
        if (entry[0] === 'removeLabel' || entry[0] === 'addLabel') {
          assert.ok(idx > setValuesLogIndex, 'label change must come after the sheet write');
        }
      });

      assert.ok(ctx.calls.getOrCreateLabelCalls.includes('jh'));
      assert.equal(ctx.calls.releaseLockCalls, 1);

      const lastRunWrite = ctx.calls.propertyWrites.filter(function (w) { return w.key === JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY; })[0];
      assert.ok(lastRunWrite);
      assert.ok(Number(lastRunWrite.value) >= timeBeforeCall);

      const store = ctx.getSheetStore();
      assert.deepEqual(store[1], defaultJobHuntersSheetRows()[1]);
      assert.deepEqual(store[2], defaultJobHuntersSheetRows()[2]);

      ctx.calls.setValues.concat(ctx.calls.setNumberFormat).forEach(function (call) {
        assert.ok(call.col >= 1 && call.col + call.numCols - 1 <= 12, 'every write must stay within columns 1-12');
      });

      const idFormatCalls = ctx.calls.setNumberFormat.filter(function (c) { return c.col === 1 && c.format === '@'; });
      const isBulkFormatCalls = ctx.calls.setNumberFormat.filter(function (c) { return c.col === 12; });
      assert.ok(idFormatCalls.length >= 1);
      assert.equal(isBulkFormatCalls.length, 0);
    }
  );
});

test('quick-261016-nu9: live-reported bug -- checkbox-formatted-but-empty rows 2-1000 no longer push the write to row 1001', () => {
  const message = jobHunterFakeMessage({ id: 'msg-0001', threadId: 'thr-A', attachments: [] });
  const thread = jobHunterFakeThread('thr-A', new Date('2026-10-10T08:00:00Z'), [message]);

  withJobHuntersRunGlobals(
    {
      properties: defaultJobHuntersProperties(),
      sheetRows: defaultJobHuntersSheetRows(),
      sheetMaxRows: 1000,
      threads: [thread],
    },
    (ctx) => {
      const result = runJobHuntersCapture(false);

      assert.deepEqual(result, { ran: true, rowsWritten: 1, threadsSucceeded: 1, threadsFailed: 0 });
      assert.equal(ctx.calls.setValues.length, 1);
      assert.equal(ctx.calls.setValues[0].row, 4, 'must write right after the real data (row 3), not at maxRows+1 (1001)');
      assert.equal(ctx.calls.insertRowsAfter.length, 0, 'the sheet already has enough rows, so it must not be grown');
    }
  );
});

test('quick-261016-nu9: sheet grown with insertRowsAfter when new rows would exceed getMaxRows(), format+write still land after the insert', () => {
  const msg1 = jobHunterFakeMessage({ id: 'msg-0001', threadId: 'thr-A', attachments: [] });
  const msg2 = jobHunterFakeMessage({ id: 'msg-0002', threadId: 'thr-A', attachments: [] });
  const thread = jobHunterFakeThread('thr-A', new Date('2026-10-10T08:00:00Z'), [msg1, msg2]);

  withJobHuntersRunGlobals(
    {
      properties: defaultJobHuntersProperties(),
      sheetRows: defaultJobHuntersSheetRows(),
      sheetMaxRows: 3,
      threads: [thread],
    },
    (ctx) => {
      const result = runJobHuntersCapture(false);

      assert.deepEqual(result, { ran: true, rowsWritten: 2, threadsSucceeded: 1, threadsFailed: 0 });
      assert.equal(ctx.calls.insertRowsAfter.length, 1);
      assert.deepEqual(ctx.calls.insertRowsAfter[0], { afterRow: 3, howMany: 2 });
      assert.equal(ctx.calls.setValues[0].row, 4);

      const insertIdx = ctx.log.findIndex(function (e) { return e[0] === 'insertRowsAfter'; });
      const formatIdx = ctx.log.findIndex(function (e) { return e[0] === 'setNumberFormat'; });
      const writeIdx = ctx.log.findIndex(function (e) { return e[0] === 'setValues'; });
      assert.ok(insertIdx < formatIdx, 'rows must be inserted before the plain-text format is applied');
      assert.ok(formatIdx < writeIdx, 'the plain-text format must be applied before the values are written');
    }
  );
});

test('quick-261016-j0d (P-3): non-contiguous header -- two setValues calls, the id-bearing segment written last, no write touches the foreign column', () => {
  const header = ['id', 'thread_id', 'received_at', 'direction', 'from_email', 'from_name', 'to', 'cc', 'subject', 'body', 'notes', 'attachments', 'is_bulk'];
  const message = jobHunterFakeMessage({ id: 'msg-nc-1', threadId: 'thr-nc', attachments: [] });
  const thread = jobHunterFakeThread('thr-nc', new Date('2026-10-10T08:00:00Z'), [message]);

  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetRows: [header], threads: [thread] },
    (ctx) => {
      runJobHuntersCapture(false);

      assert.equal(ctx.calls.setValues.length, 2);
      assert.equal(ctx.calls.setValues[0].col, 12);
      assert.equal(ctx.calls.setValues[1].col, 1);

      ctx.calls.setValues.concat(ctx.calls.setNumberFormat).forEach(function (call) {
        const touchesNotes = call.col <= 11 && 11 <= call.col + call.numCols - 1;
        assert.equal(touchesNotes, false, 'no write may touch the foreign "notes" column');
      });
    }
  );
});

test('quick-261016-j0d (D-07): an all-deduped thread writes nothing but is still relabeled', () => {
  const message = jobHunterFakeMessage({ id: 'msg-existing-1', threadId: 'thr-dedup' });
  const thread = jobHunterFakeThread('thr-dedup', new Date('2026-10-10T08:00:00Z'), [message]);

  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetRows: defaultJobHuntersSheetRows(), threads: [thread] },
    (ctx) => {
      const result = runJobHuntersCapture(false);
      assert.deepEqual(result, { ran: true, rowsWritten: 0, threadsSucceeded: 1, threadsFailed: 0 });
      assert.equal(ctx.calls.setValues.length, 0);
      assert.ok(ctx.log.some(function (e) { return e[0] === 'addLabel' && e[1] === 'thr-dedup'; }));
    }
  );
});

test('quick-261016-j0d (D-07): an empty search still runs, writes nothing, writes LAST_RUN', () => {
  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetRows: defaultJobHuntersSheetRows(), threads: [] },
    (ctx) => {
      const result = runJobHuntersCapture(false);
      assert.deepEqual(result, { ran: true, rowsWritten: 0, threadsSucceeded: 0, threadsFailed: 0 });
      assert.equal(ctx.calls.setValues.length, 0);
      assert.equal(ctx.log.length, 0);
      assert.ok(ctx.calls.propertyWrites.some(function (w) { return w.key === JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY; }));
    }
  );
});

test('quick-261016-j0d (D-06): a sheet missing the is_bulk column aborts the whole run -- zero writes, zero labels, search never called, lock released, LAST_RUN written', () => {
  const headerMissingIsBulk = OWNER_HEADER.filter(function (name) { return name !== 'is_bulk'; });

  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetRows: [headerMissingIsBulk], threads: [] },
    (ctx) => {
      assert.throws(function () {
        runJobHuntersCapture(false);
      }, /is_bulk/);

      assert.equal(ctx.calls.setValues.length, 0);
      assert.equal(ctx.log.length, 0);
      assert.equal(ctx.calls.searchQueries.length, 0);
      assert.equal(ctx.calls.releaseLockCalls, 1);
      assert.ok(ctx.calls.propertyWrites.some(function (w) { return w.key === JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY; }));
    }
  );
});

test('quick-261016-j0d (D-07): a missing sheet throws naming "Emails", with zero writes', () => {
  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetMissing: true, threads: [] },
    (ctx) => {
      assert.throws(function () {
        runJobHuntersCapture(false);
      }, /Emails/);
      assert.equal(ctx.calls.setValues.length, 0);
    }
  );
});

test('quick-261016-j0d (P-10): an unset spreadsheet ID throws naming the key, openById never called, lock released', () => {
  withJobHuntersRunGlobals({ properties: {}, threads: [] }, (ctx) => {
    assert.throws(function () {
      runJobHuntersCapture(false);
    }, /11-action-job-hunters-SPREADSHEET_ID/);
    assert.equal(ctx.calls.openByIdCalls.length, 0);
    assert.equal(ctx.calls.releaseLockCalls, 1);
  });
});

test('quick-261016-j0d (P-11, P-12): a per-thread failure isolates the failing thread -- its rows are never written, it keeps its label, the owner is notified', () => {
  const failingMessage = jobHunterFakeMessage({ id: 'msg-a-fail', threadId: 'thr-A-fail', plainBodyThrows: true });
  const threadA = jobHunterFakeThread('thr-A-fail', new Date('2026-10-10T08:00:00Z'), [failingMessage]);
  const okMessage = jobHunterFakeMessage({ id: 'msg-0004', threadId: 'thr-B-ok', attachments: [] });
  const threadB = jobHunterFakeThread('thr-B-ok', new Date('2026-10-11T08:00:00Z'), [okMessage]);

  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), sheetRows: defaultJobHuntersSheetRows(), threads: [threadB, threadA] },
    (ctx) => {
      const result = runJobHuntersCapture(false);
      assert.deepEqual(result, { ran: true, rowsWritten: 1, threadsSucceeded: 1, threadsFailed: 1 });

      assert.equal(ctx.calls.setValues.length, 1);
      assert.equal(ctx.calls.setValues[0].values.length, 1);
      assert.equal(ctx.calls.setValues[0].values[0][0], 'msg-0004');

      assert.equal(ctx.log.filter(function (e) { return e[1] === 'thr-A-fail'; }).length, 0);

      assert.equal(ctx.calls.notifyOwnerOfFailureCalls.length, 1);
      assert.equal(ctx.calls.notifyOwnerOfFailureCalls[0].actionName, 'JobHunters');
      assert.equal(ctx.calls.notifyOwnerOfFailureCalls[0].thread, threadA);
    }
  );

  // Rerun with NOTIFY_ON_FAILURE 'false': the spy is not called.
  const failingMessage2 = jobHunterFakeMessage({ id: 'msg-a-fail-2', threadId: 'thr-A-fail-2', plainBodyThrows: true });
  const threadA2 = jobHunterFakeThread('thr-A-fail-2', new Date('2026-10-10T08:00:00Z'), [failingMessage2]);

  withJobHuntersRunGlobals(
    {
      properties: defaultJobHuntersProperties({ '11-action-job-hunters-NOTIFY_ON_FAILURE': 'false' }),
      sheetRows: defaultJobHuntersSheetRows(),
      threads: [threadA2],
    },
    (ctx) => {
      runJobHuntersCapture(false);
      assert.equal(ctx.calls.notifyOwnerOfFailureCalls.length, 0);
    }
  );
});

test('quick-261016-j0d (D-04): not due -- no lock/search/openById/writes; force=true bypasses the due-check and rewrites LAST_RUN', () => {
  const lastRun = String(Date.now() - 5 * 60000);

  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties({ '11-action-job-hunters-LAST_RUN_AT_MS': lastRun }), threads: [] },
    (ctx) => {
      const result = runJobHuntersCapture(false);
      assert.deepEqual(result, { ran: false, reason: 'not due' });
      assert.equal(ctx.calls.getScriptLockCalls, 0);
      assert.equal(ctx.calls.searchQueries.length, 0);
      assert.equal(ctx.calls.openByIdCalls.length, 0);
      assert.equal(ctx.calls.propertyWrites.length, 0);
    }
  );

  withJobHuntersRunGlobals(
    {
      properties: defaultJobHuntersProperties({ '11-action-job-hunters-LAST_RUN_AT_MS': lastRun }),
      sheetRows: defaultJobHuntersSheetRows(),
      threads: [],
    },
    (ctx) => {
      const result = runJobHuntersCapture(true);
      assert.equal(result.ran, true);
      assert.ok(ctx.calls.propertyWrites.some(function (w) { return w.key === JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY; }));
    }
  );
});

test('quick-261016-j0d (P-9): disabled -- ran:false reason:disabled for both force values, nothing else touched', () => {
  [false, true].forEach(function (force) {
    withJobHuntersRunGlobals(
      { properties: defaultJobHuntersProperties({ '11-action-job-hunters-ENABLED': 'false' }), threads: [] },
      (ctx) => {
        const result = runJobHuntersCapture(force);
        assert.deepEqual(result, { ran: false, reason: 'disabled' });
        assert.equal(ctx.calls.getScriptLockCalls, 0);
        assert.equal(ctx.calls.searchQueries.length, 0);
        assert.equal(ctx.calls.openByIdCalls.length, 0);
        assert.equal(ctx.calls.propertyWrites.length, 0);
      }
    );
  });
});

test('quick-261016-j0d (D-05): lock busy -- ran:false reason "lock busy", no search, zero propertyWrites, no releaseLock', () => {
  withJobHuntersRunGlobals(
    { properties: defaultJobHuntersProperties(), threads: [], lockAcquired: false },
    (ctx) => {
      const result = runJobHuntersCapture(false);
      assert.deepEqual(result, { ran: false, reason: 'lock busy' });
      assert.equal(ctx.calls.tryLockMs.length, 1);
      assert.equal(ctx.calls.tryLockMs[0], 30000);
      assert.equal(ctx.calls.searchQueries.length, 0);
      assert.equal(ctx.calls.propertyWrites.length, 0);
      assert.equal(ctx.calls.releaseLockCalls, 0);
    }
  );
});

test('quick-261016-j0d: a custom JOB_HUNTER_LABEL_NAME and SEARCH_INTERVAL_MINUTES are honored', () => {
  withJobHuntersRunGlobals(
    {
      properties: defaultJobHuntersProperties({
        '11-action-job-hunters-JOB_HUNTER_LABEL_NAME': 'jobs-test',
        '11-action-job-hunters-SEARCH_INTERVAL_MINUTES': '5',
        '11-action-job-hunters-LAST_RUN_AT_MS': String(Date.now() - 6 * 60000),
      }),
      sheetRows: defaultJobHuntersSheetRows(),
      threads: [],
    },
    (ctx) => {
      const result = runJobHuntersCapture(false);
      assert.equal(result.ran, true);
      assert.equal(ctx.calls.searchQueries[0], 'label:"jobs-test"');
    }
  );
});

// --- Static wiring guards ------------------------------------------------
//
// Read the compiled source text with fs so these checks hold even before
// Task 2/3/4 exist -- the RED-phase failures below are expected to be
// "false" assertions, not thrown TypeErrors. Slices a function body from
// its declaration to the first "\n}" after it, stripping "//" line comments
// first (quick-261016-j0d, D-14).

function readSrcFile(relPath) {
  return fs.readFileSync(path.join(__dirname, '..', relPath), 'utf8');
}

function stripLineComments(text) {
  return text
    .split('\n')
    .map(function (line) {
      return line.replace(/\/\/.*$/, '');
    })
    .join('\n');
}

function sliceFunctionBody(source, functionName) {
  const startMatch = source.match(new RegExp('function\\s+' + functionName + '\\s*\\('));
  if (!startMatch) {
    return '';
  }
  const afterStart = source.slice(startMatch.index);
  const endIndex = afterStart.indexOf('\n}');
  const rawSlice = endIndex === -1 ? afterStart : afterStart.slice(0, endIndex);
  return stripLineComments(rawSlice);
}

test('quick-261016-j0d wiring guard (D-11): processEmails calls runJobHuntersCapture(false) after threads.forEach, in its own null-tolerant notify path, with the query line unchanged', () => {
  const source = readSrcFile('src/02-main.js');
  const body = sliceFunctionBody(source, 'processEmails');

  const forEachIndex = body.indexOf('threads.forEach');
  const captureIndex = body.indexOf('runJobHuntersCapture(false)');
  assert.ok(forEachIndex !== -1, 'processEmails must still contain threads.forEach');
  assert.ok(captureIndex !== -1, 'processEmails must call runJobHuntersCapture(false)');
  assert.ok(captureIndex > forEachIndex, 'the JobHunters call must come after the dispatch loop');

  assert.ok(body.includes("notifyOwnerOfFailure('JobHunters', null,"));
  assert.ok(body.includes('const query = \'newer_than:\' + CONFIG.daysBack + \'d -label:"\' + CONFIG.labelName + \'"\';'));
});

test('quick-261016-j0d wiring guard (D-01): getRegisteredActions is byte-unchanged -- JobHunters is never registered', () => {
  const source = readSrcFile('src/03-action-management.js');
  assert.ok(source.includes('return [ICS_CALENDAR_ACTION, BOOKING_MANAGEMENT_ACTION, TICKETING_PORTALS_ACTION, TRANSPORT_TICKETS_ACTION, MEETINGS_ACTION];'));
});

test('quick-261016-j0d wiring guard (D-09): testJobHuntersDryRun is the first function declared, and never touches Sheets/labels/lock/properties', () => {
  const source = readSrcFile('src/11-action-job-hunters.js');
  const match = source.match(/^function (\w+)/m);
  assert.ok(match, 'expected at least one top-level function declaration');
  assert.equal(match[1], 'testJobHuntersDryRun');

  const body = sliceFunctionBody(source, 'testJobHuntersDryRun');
  ['SpreadsheetApp', 'setValues', 'addLabel', 'removeLabel', 'LockService', 'setProperty'].forEach(function (token) {
    assert.equal(body.includes(token), false, 'testJobHuntersDryRun body must not reference ' + token);
  });
});

test('quick-261016-j0d wiring guard (D-12): appsscript.json declares the spreadsheets scope additively, and enabledAdvancedServices stays at 2 entries', () => {
  const manifest = JSON.parse(readSrcFile('src/appsscript.json'));
  assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/spreadsheets'));

  [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/script.send_mail',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/script.external_request',
  ].forEach(function (scope) {
    assert.ok(manifest.oauthScopes.includes(scope), 'missing pre-existing scope: ' + scope);
  });

  assert.equal(manifest.dependencies.enabledAdvancedServices.length, 2);
});

test('quick-261016-j0d wiring guard [settings]: (D-08) setup() provisions both JobHunters labels', () => {
  const source = readSrcFile('src/01-setup.js');
  const body = sliceFunctionBody(source, 'setup');
  assert.ok(body.includes('getOrCreateLabel(JOB_HUNTERS_ACTION_CONFIG.jobHunterLabelName)'));
  assert.ok(body.includes('getOrCreateLabel(JOB_HUNTERS_ACTION_CONFIG.processedLabelName)'));
});

test('quick-261016-j0d wiring guard [webapp]: (D-10, T-261016-03) webappRunJobHuntersNow asserts ownership first and calls the real capture + describer', () => {
  const source = readSrcFile('src/00-webapp.js');
  assert.ok(/function webappRunJobHuntersNow\(\)\s*\{\s*assertWebappOwner\(\);/.test(source));

  const body = sliceFunctionBody(source, 'webappRunJobHuntersNow');
  assert.ok(body.includes('runJobHuntersCapture(true)'));
  assert.ok(body.includes('jobHuntersDescribeRunResult('));
});

test('quick-261016-j0d wiring guard [webapp]: (D-02, D-10, P-17) the admin UI wires the Job Hunters sub-tab, Run Now button and its own result box; zero innerHTML', () => {
  const html = readSrcFile('src/00-webapp-ui.html');
  assert.ok(html.includes("'11-action-job-hunters-': 'Job Hunters'"));

  const groupPrefixesMatch = html.match(/GROUP_PREFIXES\s*=\s*\[([\s\S]*?)\]/);
  assert.ok(groupPrefixesMatch, 'expected a GROUP_PREFIXES array literal');
  assert.ok(groupPrefixesMatch[1].includes("'11-action-job-hunters-'"));

  assert.ok(html.includes('.webappRunJobHuntersNow()'));
  assert.ok(html.includes('job-hunters-run-result'));

  const innerHtmlCount = (html.match(/innerHTML/g) || []).length;
  assert.equal(innerHtmlCount, 0);
});

test('quick-261016-nu9 wiring guard [webapp]: webappGetJobHuntersNextRun asserts ownership first and delegates to the pure describer', () => {
  const source = readSrcFile('src/00-webapp.js');
  assert.ok(/function webappGetJobHuntersNextRun\(\)\s*\{\s*assertWebappOwner\(\);/.test(source));

  const body = sliceFunctionBody(source, 'webappGetJobHuntersNextRun');
  assert.ok(body.includes('jobHuntersGetNextRunDescription()'));
});

test('quick-261016-nu9 wiring guard [webapp]: the Job Hunters sub-tab displays and refreshes the next-run time around Run Now', () => {
  const html = readSrcFile('src/00-webapp-ui.html');
  assert.ok(html.includes('job-hunters-next-run'));
  assert.ok(html.includes('.webappGetJobHuntersNextRun()'));

  const innerHtmlCount = (html.match(/innerHTML/g) || []).length;
  assert.equal(innerHtmlCount, 0);
});

/**
 * JobHunters capture step -- src/11-action-job-hunters.js (quick-261016-j0d).
 *
 * STANDALONE (D-01): this is NOT a registered action. getRegisteredActions()
 * (src/03-action-management.js) never includes it; it is invoked directly
 * from processEmails() (src/02-main.js) via runJobHuntersCapture(false),
 * after the existing dispatch loop, in its own try/catch.
 *
 * APPEND-ONLY, FIRST-12-COLUMNS, RESOLVE-BY-NAME (D-06): the target sheet
 * belongs to ANOTHER APP. This script writes ONLY the 12 columns
 * id..is_bulk (JOB_HUNTERS_COLUMN_NAMES), resolved by header NAME via
 * resolveJobHuntersColumnIndexes -- never a fixed A:L range. Columns 13-20
 * and every existing row are never read, written, or formatted.
 *
 * WHY NOT sheet.getLastRow() (live-reported, quick-261016-nu9 round 2): the
 * owner's real "Emails" sheet carries checkbox DATA VALIDATION on the
 * is_bulk column for rows 2-1000 (the other app's own pre-provisioning).
 * getLastRow() counts a cell as non-empty when it carries a checkbox
 * control, even with no id and no real data in that row -- so it returned
 * 1000 on an otherwise-empty sheet, and the first real rows landed at 1001
 * instead of row 2. The fix reads the WHOLE id column in one call
 * (jobHuntersFindLastDataRow, below) and finds the last row whose id CELL
 * VALUE is actually non-empty, which checkbox formatting alone can never
 * satisfy -- the same single read also seeds the dedup Set
 * (jobHuntersBuildExistingIdSet), so there is still only one id-column read
 * per run.
 *
 * CRASH-SAFETY ORDERING: the sheet write happens BEFORE any label swap. A
 * thread is only relabeled jobHunter -> jh after its rows are durably
 * written, so a failure between those two steps leaves the thread exactly
 * where dedup will pick it up again next run.
 *
 * GLOBALLY-UNIQUE NAMING (P-1): Apps Script concatenates every project file
 * into ONE shared global scope. Every helper here is `jobHunters`/
 * `JOB_HUNTERS`-prefixed to stay collision-free with the rest of the
 * codebase.
 *
 * FORMULA-INJECTION GUARD (P-4, threat T-261016-01): before writing a new
 * row, every one of its cells except is_bulk gets setNumberFormat('@')
 * (plain text), so an inbound subject/body/address starting with "=" is
 * never evaluated as a Sheets formula, and an all-digit message id is never
 * coerced to a number (which would break dedup).
 */

// FUNCTION-PICKER ORDERING (D-09, same convention as src/01-setup.js's own
// note): testJobHuntersDryRun must be the FIRST function declaration in
// this file so it appears at the top of the Apps Script editor's function
// picker, ahead of every internal helper. runJobHuntersCapture is second,
// for the same reason -- it is the other manually-triggerable entry point
// (also called by webappRunJobHuntersNow).

/**
 * testJobHuntersDryRun — read-only preview, runnable directly from the Apps
 * Script editor's function picker. A SEPARATE code path from
 * runJobHuntersCapture (D-09), never sharing logic with the real thread
 * loop (small duplication accepted by design). Never touches Sheets, the
 * lock, labels, the dedup read, or any Script Property write. Logs the
 * would-be rows of the latest JOB_HUNTERS_DRY_RUN_MESSAGE_COUNT
 * jobHunter-labeled messages across all matching threads, with each row's
 * body cut to JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS chars, and returns the
 * array of built rows. Ignores `enabled` (P-16) since this is a read-only
 * preview; logs a note instead when it is false.
 */
function testJobHuntersDryRun() {
  const config = JOB_HUNTERS_ACTION_CONFIG;

  console.log('=== JobHunters DRY RUN -- nothing is written, no label changes ===');
  if (config.enabled !== true) {
    console.log('Note: JobHunters is currently DISABLED (enabled=false); this preview still runs.');
  }

  const threads = orderThreadsForProcessing(GmailApp.search('label:"' + config.jobHunterLabelName + '"'));
  const messages = selectLatestJobHunterMessages(threads, JOB_HUNTERS_DRY_RUN_MESSAGE_COUNT);

  if (messages.length === 0) {
    console.log('No message currently carries the "' + config.jobHunterLabelName + '" label.');
    return [];
  }

  const rows = messages.map(function (message) {
    return buildJobHunterRow(message, config.ownAddresses);
  });

  rows.forEach(function (row) {
    jobHuntersFormatDryRunRow(row).forEach(function (line) {
      console.log(line);
    });
    console.log('---');
  });

  return rows;
}

/**
 * runJobHuntersCapture — the GAS-only entry point (D-07). `force` bypasses
 * ONLY the due-check (D-04) -- `enabled` and the lock are never bypassed
 * (P-9). Exported for the Node global-injection test harness, the same
 * precedent as processTicketFromMessageBody (src/07-action-ticketing-portals.js).
 *
 * Order: enabled check -> due check (unless force) -> lock -> (inside
 * try/finally) open spreadsheet+sheet -> resolve columns (fail closed,
 * D-06) -> read the dedup Set -> search+order threads -> per-thread
 * try/catch buffering rows (P-11) -> ONE batched write per contiguous
 * column segment, id-bearing segment last (P-3) -> relabel only the
 * succeeded threads, each swap in its own try/catch (P-12) -> finally:
 * write LAST_RUN_AT_MS (P-8) then release the lock.
 *
 * Returns one of:
 *   {ran:false, reason:'disabled'}
 *   {ran:false, reason:'not due'}
 *   {ran:false, reason:'lock busy'}
 *   {ran:true, rowsWritten, threadsSucceeded, threadsFailed}
 */

/**
 * jobHuntersGetNextRunInfo — GAS-only (quick-261016-nu9 round 4, replaces
 * round 3's jobHuntersGetNextRunDescription). Reads the SAME LAST_RUN_AT_MS
 * property and searchIntervalMinutes setting runJobHuntersCapture's own
 * due-check consults, and returns the STRUCTURED jobHuntersNextRunInfo --
 * deliberately NOT a pre-rendered string, so the webapp RPC can hand
 * `nextRunAtMs` straight to the browser for local-timezone rendering
 * instead of baking in a server-side (UTC) time string. Read-only -- never
 * touches PropertiesService beyond a single getProperty, never acquires the
 * lock, never searches Gmail. Exported for the Node global-injection
 * harness, same precedent as runJobHuntersCapture itself.
 */
function jobHuntersGetNextRunInfo() {
  const lastRunRaw = PropertiesService.getScriptProperties().getProperty(JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY);
  return jobHuntersNextRunInfo(lastRunRaw, JOB_HUNTERS_ACTION_CONFIG.searchIntervalMinutes, Date.now());
}

function runJobHuntersCapture(force) {
  const config = JOB_HUNTERS_ACTION_CONFIG;

  if (config.enabled === false) {
    return { ran: false, reason: 'disabled' };
  }

  const scriptProperties = PropertiesService.getScriptProperties();

  if (!force) {
    const lastRunRaw = scriptProperties.getProperty(JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY);
    if (!jobHuntersIsDue(lastRunRaw, config.searchIntervalMinutes, Date.now())) {
      return { ran: false, reason: 'not due' };
    }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(JOB_HUNTERS_LOCK_TIMEOUT_MS)) {
    console.log('runJobHuntersCapture: could not acquire the script lock -- skipping this run.');
    return { ran: false, reason: 'lock busy' };
  }

  try {
    if (!jobHuntersIsSpreadsheetIdConfigured(config.spreadsheetId)) {
      throw new Error('JobHunters: Script Property 11-action-job-hunters-SPREADSHEET_ID is not configured.');
    }

    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    } catch (openError) {
      throw new Error('JobHunters: could not open the JobHunters spreadsheet: ' + openError.message);
    }

    const sheet = spreadsheet.getSheetByName(config.sheetName);
    if (sheet === null) {
      throw new Error('JobHunters: target sheet "' + config.sheetName + '" was not found in the JobHunters spreadsheet.');
    }

    const lastColumn = sheet.getLastColumn();
    const header = lastColumn < 1 ? [] : sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const columnIndexes = resolveJobHuntersColumnIndexes(header);
    const segments = jobHuntersColumnSegments(columnIndexes);

    // quick-261016-nu9 round 2: read the id column's FULL range (header
    // through getMaxRows(), one call), not getLastRow() -- see this file's
    // class-level "WHY NOT getLastRow()" note. The header row is sliced off
    // before feeding the same data rows to BOTH jobHuntersBuildExistingIdSet
    // (dedup) and jobHuntersFindLastDataRow (next write row), so there is
    // still only one id-column read per run.
    const maxRows = sheet.getMaxRows();
    const idColumnDataRows = sheet.getRange(1, columnIndexes.id + 1, maxRows, 1).getValues().slice(1);
    const existingIds = jobHuntersBuildExistingIdSet(idColumnDataRows);
    const lastDataRow = jobHuntersFindLastDataRow(idColumnDataRows);

    const threads = orderThreadsForProcessing(GmailApp.search('label:"' + config.jobHunterLabelName + '"'));

    const newRows = [];
    const succeededThreads = [];
    let threadsFailed = 0;

    threads.forEach(function (thread) {
      const threadRows = [];
      const threadIds = [];

      try {
        thread.getMessages().forEach(function (message) {
          const id = String(message.getId());
          if (existingIds.has(id) || threadIds.indexOf(id) !== -1) {
            return;
          }
          threadRows.push(buildJobHunterRow(message, config.ownAddresses));
          threadIds.push(id);
        });

        threadRows.forEach(function (row) {
          newRows.push(row);
        });
        threadIds.forEach(function (id) {
          existingIds.add(id);
        });
        succeededThreads.push(thread);
      } catch (threadError) {
        threadsFailed++;
        console.error('JobHunters: thread ' + thread.getId() + ' failed: ' + threadError);
        if (config.notifyOnFailure === true) {
          notifyOwnerOfFailure('JobHunters', thread, threadError);
        }
      }
    });

    if (newRows.length > 0) {
      const startRow = lastDataRow + 1;
      const neededRows = startRow + newRows.length - 1 - maxRows;
      if (neededRows > 0) {
        sheet.insertRowsAfter(maxRows, neededRows);
      }

      const orderedSegments = segments.slice().sort(function (a, b) {
        const aHasId = a.fields.indexOf('id') !== -1 ? 1 : 0;
        const bHasId = b.fields.indexOf('id') !== -1 ? 1 : 0;
        return aHasId - bHasId;
      });

      orderedSegments.forEach(function (segment) {
        segment.fields.forEach(function (fieldName, fieldOffset) {
          if (fieldName === 'is_bulk') {
            return;
          }
          sheet.getRange(startRow, segment.startColumn + fieldOffset, newRows.length, 1).setNumberFormat('@');
        });

        sheet
          .getRange(startRow, segment.startColumn, newRows.length, segment.fields.length)
          .setValues(jobHuntersRowsForSegment(newRows, segment));
      });
    }

    let threadsSucceeded = 0;
    if (succeededThreads.length > 0) {
      const jobHunterLabel = GmailApp.getUserLabelByName(config.jobHunterLabelName);
      const processedLabel = getOrCreateLabel(config.processedLabelName);

      succeededThreads.forEach(function (thread) {
        try {
          if (jobHunterLabel !== null) {
            thread.removeLabel(jobHunterLabel);
          }
          thread.addLabel(processedLabel);
          threadsSucceeded++;
        } catch (relabelError) {
          threadsFailed++;
          console.error('JobHunters: relabeling thread ' + thread.getId() + ' failed: ' + relabelError);
          if (config.notifyOnFailure === true) {
            notifyOwnerOfFailure('JobHunters', thread, relabelError);
          }
        }
      });
    }

    return { ran: true, rowsWritten: newRows.length, threadsSucceeded: threadsSucceeded, threadsFailed: threadsFailed };
  } finally {
    scriptProperties.setProperty(JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY, String(Date.now()));
    lock.releaseLock();
  }
}

// --- Constants -----------------------------------------------------------

// The 12 names id..is_bulk, in D-06 order -- the exact set of columns this
// script is ever allowed to touch.
const JOB_HUNTERS_COLUMN_NAMES = [
  'id', 'thread_id', 'received_at', 'direction', 'from_email', 'from_name',
  'to', 'cc', 'subject', 'body', 'attachments', 'is_bulk',
];

const JOB_HUNTERS_BODY_MAX_CHARS = 10000;
const JOB_HUNTERS_DRY_RUN_MESSAGE_COUNT = 3;
const JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS = 200;
const JOB_HUNTERS_LOCK_TIMEOUT_MS = 30000;

// --- Pure helpers ----------------------------------------------------------

/**
 * jobHuntersIsDue — pure due-check gate (D-04, P-7). Sanitizes every input
 * so a corrupt/garbage stored value can never block runs forever: a
 * non-finite or negative `lastRunAtMs` becomes 0, a non-finite or negative
 * `intervalMinutes` becomes 0, a future-dated last run always counts as
 * due, and a non-finite `nowMs` always counts as due. Never throws.
 */
function jobHuntersIsDue(lastRunAtMs, intervalMinutes, nowMs) {
  let last = Number(lastRunAtMs);
  if (!Number.isFinite(last) || last < 0) {
    last = 0;
  }

  let interval = Number(intervalMinutes);
  if (!Number.isFinite(interval) || interval < 0) {
    interval = 0;
  }

  const now = Number(nowMs);
  if (!Number.isFinite(now)) {
    return true;
  }

  if (last > now) {
    return true;
  }

  return now - last >= interval * 60000;
}

/**
 * jobHuntersNextRunInfo — pure (quick-261016-nu9 round 4). The STRUCTURED
 * form behind both jobHuntersDescribeNextRun (below, a UTC-text form kept
 * for log/console use) and the webapp RPC (which sends `nextRunAtMs` to the
 * BROWSER to render in the viewer's OWN local timezone -- see round 4's fix
 * note on jobHuntersDescribeNextRun: a server-rendered time string can only
 * ever be in ONE fixed zone, never the viewer's). Built on the exact same
 * due-check jobHuntersIsDue already uses, so the two can never disagree.
 * `nextRunAtMs` is `null` exactly when `hasRunBefore` is false (never run
 * yet -- there is no previous run to add the interval to). Never throws.
 */
function jobHuntersNextRunInfo(lastRunAtMs, intervalMinutes, nowMs) {
  const last = Number(lastRunAtMs);
  const hasRunBefore = Number.isFinite(last) && last > 0;
  const due = jobHuntersIsDue(lastRunAtMs, intervalMinutes, nowMs);
  const nextRunAtMs = hasRunBefore ? last + Number(intervalMinutes) * 60000 : null;

  return { hasRunBefore: hasRunBefore, due: due, nextRunAtMs: nextRunAtMs };
}

/**
 * jobHuntersDescribeNextRun — pure (quick-261016-nu9 round 3: "next search
 * time" display). Human-readable, UTC-text string for when the NEXT
 * scheduled run becomes due, built on jobHuntersNextRunInfo above.
 *
 * ROUND 4 FIX NOTE: this is NO LONGER what the webapp's Job Hunters tab
 * displays -- a live-reported issue found that its `toISOString()` (always
 * UTC) read as "the script's timezone" to the owner, which is misleading:
 * it is unrelated to appsscript.json's configured timeZone and, more to the
 * point, is not the OWNER's own timezone either. The webapp RPC
 * (webappGetJobHuntersNextRun, src/00-webapp.js) now sends the raw
 * `nextRunAtMs` from jobHuntersNextRunInfo to the browser instead, which
 * renders it via the viewer's own local Date#toLocaleString(). This
 * function is kept as a pure, independently tested UTC-text form for any
 * future server-side log/console use, where UTC is the right and
 * unambiguous choice. Never throws.
 */
function jobHuntersDescribeNextRun(lastRunAtMs, intervalMinutes, nowMs) {
  const info = jobHuntersNextRunInfo(lastRunAtMs, intervalMinutes, nowMs);

  if (info.due) {
    return info.hasRunBefore ? 'Due now.' : 'Never run yet -- due now.';
  }

  const minutesRemaining = Math.max(1, Math.ceil((info.nextRunAtMs - Number(nowMs)) / 60000));

  return (
    'Next due ' +
    new Date(info.nextRunAtMs).toISOString() +
    ' (in about ' +
    minutesRemaining +
    (minutesRemaining === 1 ? ' minute).' : ' minutes).')
  );
}

/**
 * jobHuntersParseFromHeader — pure (D-06, P-1). Returns {name, email}. The
 * email is the LAST "<...>" pair anchored at the end of the header (else
 * the whole trimmed value), always trimmed and lowercased. The name is the
 * text before that pair, trimmed, with one surrounding pair of double
 * quotes stripped, then trimmed again -- '' when absent. null/''/undefined
 * gives {name:'', email:''}. Never throws.
 */
function jobHuntersParseFromHeader(fromHeader) {
  if (!fromHeader) {
    return { name: '', email: '' };
  }

  const trimmedHeader = String(fromHeader).trim();
  const match = trimmedHeader.match(/^([\s\S]*)<([^>]*)>\s*$/);

  if (!match) {
    return { name: '', email: trimmedHeader.toLowerCase() };
  }

  let name = match[1].trim();
  if (name.length >= 2 && name.charAt(0) === '"' && name.charAt(name.length - 1) === '"') {
    name = name.slice(1, -1).trim();
  }

  return { name: name, email: match[2].trim().toLowerCase() };
}

/**
 * jobHuntersNormalizeAddressList — pure. Splits on ',', trims each part,
 * drops empties, joins with ', ', then lowercases the whole result.
 * Preserves order and duplicates. null/'' give ''. Never throws.
 */
function jobHuntersNormalizeAddressList(raw) {
  if (!raw) {
    return '';
  }

  const parts = String(raw)
    .split(',')
    .map(function (part) {
      return part.trim();
    })
    .filter(function (part) {
      return part.length > 0;
    });

  return parts.join(', ').toLowerCase();
}

// Quote-boundary line patterns (P-5): a line matching any of these, or
// exactly "-- " (trailing space, the Usenet/email signature delimiter),
// ends the kept body -- everything from that line onward is dropped. A
// bare "--" (no trailing space) is NOT a signature delimiter.
const JOB_HUNTERS_QUOTE_CUT_PATTERNS = [/^On .+wrote:$/, /^Dne .+napsal\(a\):$/, /^-{5}Original Message-{5}$/];

/**
 * cleanJobHunterBodyText — pure (D-06, P-5). Normalizes \r\n/\r to \n FIRST
 * (so the `$`-anchored cut patterns below match reliably), then walks lines
 * top to bottom: a line starting with '>' is dropped; a line matching a
 * quote-boundary pattern or exactly "-- " stops processing entirely (every
 * line from there on is dropped); every other line is kept. Joins the kept
 * lines, trims, and slices to JOB_HUNTERS_BODY_MAX_CHARS. null/undefined
 * give ''. Never throws.
 */
function cleanJobHunterBodyText(rawPlainBody) {
  if (rawPlainBody === null || rawPlainBody === undefined) {
    return '';
  }

  const normalized = String(rawPlainBody).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');
  const kept = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (
      line === '-- ' ||
      JOB_HUNTERS_QUOTE_CUT_PATTERNS.some(function (pattern) {
        return pattern.test(line);
      })
    ) {
      break;
    }

    if (line.indexOf('>') === 0) {
      continue;
    }

    kept.push(line);
  }

  return kept.join('\n').trim().slice(0, JOB_HUNTERS_BODY_MAX_CHARS);
}

/**
 * jobHuntersIsBulkEmail — pure (D-06, P-6). True when ANY of: List-Unsubscribe
 * is non-empty (after trim); Precedence is 'bulk' or 'list' (case-insensitive,
 * trimmed); Auto-Submitted is non-empty and not 'no' (case-insensitive,
 * trimmed); or the from-email local-part is EXACTLY (case-insensitive)
 * 'noreply', 'no-reply' or 'donotreply' -- a prefixed local-part like
 * 'jobs-noreply' does NOT match (P-6). Null-safe on every argument, never
 * throws, always returns a real boolean.
 */
function jobHuntersIsBulkEmail(listUnsubscribeHeader, precedenceHeader, autoSubmittedHeader, fromEmail) {
  const listUnsubscribe = listUnsubscribeHeader ? String(listUnsubscribeHeader).trim() : '';
  if (listUnsubscribe.length > 0) {
    return true;
  }

  const precedence = precedenceHeader ? String(precedenceHeader).trim().toLowerCase() : '';
  if (precedence === 'bulk' || precedence === 'list') {
    return true;
  }

  const autoSubmitted = autoSubmittedHeader ? String(autoSubmittedHeader).trim().toLowerCase() : '';
  if (autoSubmitted.length > 0 && autoSubmitted !== 'no') {
    return true;
  }

  const email = fromEmail ? String(fromEmail).trim().toLowerCase() : '';
  const atIndex = email.indexOf('@');
  const localPart = atIndex === -1 ? email : email.slice(0, atIndex);

  return localPart === 'noreply' || localPart === 'no-reply' || localPart === 'donotreply';
}

/**
 * buildJobHunterRow — pure (D-06). Builds one 12-value row, in
 * JOB_HUNTERS_COLUMN_NAMES order, from a GAS GmailMessage-shaped object.
 * Calls message.getHeader exactly for 'List-Unsubscribe', 'Precedence' and
 * 'Auto-Submitted' -- never getRawContent. `ownAddresses` null/undefined is
 * treated as [].
 */
function buildJobHunterRow(message, ownAddresses) {
  const parsedFrom = jobHuntersParseFromHeader(message.getFrom());
  const fromEmail = parsedFrom.email;

  const normalizedOwnAddresses = (ownAddresses || []).map(function (addr) {
    return String(addr).trim().toLowerCase();
  });
  const direction = normalizedOwnAddresses.indexOf(fromEmail) !== -1 ? 'outgoing' : 'incoming';

  const attachments = (message.getAttachments() || [])
    .map(function (attachment) {
      return attachment.getName();
    })
    .join(' | ');

  const isBulk = jobHuntersIsBulkEmail(
    message.getHeader('List-Unsubscribe'),
    message.getHeader('Precedence'),
    message.getHeader('Auto-Submitted'),
    fromEmail
  );

  return [
    message.getId(),
    message.getThread().getId(),
    message.getDate().toISOString(),
    direction,
    fromEmail,
    parsedFrom.name,
    jobHuntersNormalizeAddressList(message.getTo()),
    jobHuntersNormalizeAddressList(message.getCc()),
    message.getSubject(),
    cleanJobHunterBodyText(message.getPlainBody()),
    attachments,
    isBulk,
  ];
}

/**
 * resolveJobHuntersColumnIndexes — pure (D-06). Resolves each of the 12
 * JOB_HUNTERS_COLUMN_NAMES to its 0-based index in `headerRow`, matching
 * cell text via `String(cell).trim()`; on a duplicate name, the FIRST
 * occurrence wins. Throws naming every missing column when one or more of
 * the 12 names is absent -- this is the deliberate fail-closed exception,
 * the one function in this file that throws.
 */
function resolveJobHuntersColumnIndexes(headerRow) {
  const header = headerRow || [];
  const indexes = {};

  JOB_HUNTERS_COLUMN_NAMES.forEach(function (name) {
    for (let i = 0; i < header.length; i++) {
      if (String(header[i]).trim() === name) {
        indexes[name] = i;
        break;
      }
    }
  });

  const missing = JOB_HUNTERS_COLUMN_NAMES.filter(function (name) {
    return !(name in indexes);
  });

  if (missing.length > 0) {
    throw new Error('JobHunters: the target sheet header row is missing required column(s): ' + missing.join(', '));
  }

  return indexes;
}

/**
 * jobHuntersColumnSegments — pure (P-3). Groups the 12 resolved column
 * indexes into maximal CONTIGUOUS segments (sorted ascending by position),
 * returning `[{startColumn, fields}]` with 1-based `startColumn` and each
 * segment's `fields` listed in sheet order. A reordered or interleaved
 * header never merges two fields that aren't adjacent in the real sheet.
 */
function jobHuntersColumnSegments(columnIndexes) {
  const entries = JOB_HUNTERS_COLUMN_NAMES
    .map(function (name) {
      return { name: name, index: columnIndexes[name] };
    })
    .sort(function (a, b) {
      return a.index - b.index;
    });

  const segments = [];
  let current = null;

  entries.forEach(function (entry) {
    if (current && entry.index === current.lastIndex + 1) {
      current.fields.push(entry.name);
      current.lastIndex = entry.index;
      return;
    }

    current = { startColumn: entry.index + 1, fields: [entry.name], lastIndex: entry.index };
    segments.push(current);
  });

  return segments.map(function (segment) {
    return { startColumn: segment.startColumn, fields: segment.fields };
  });
}

/**
 * jobHuntersRowsForSegment — pure (P-3). Maps canonical 12-value rows (in
 * JOB_HUNTERS_COLUMN_NAMES order) to a 2D array holding only `segment`'s
 * fields, in that segment's own order.
 */
function jobHuntersRowsForSegment(rows, segment) {
  const fieldIndexes = segment.fields.map(function (name) {
    return JOB_HUNTERS_COLUMN_NAMES.indexOf(name);
  });

  return (rows || []).map(function (row) {
    return fieldIndexes.map(function (idx) {
      return row[idx];
    });
  });
}

/**
 * jobHuntersBuildExistingIdSet — pure. Takes a 2D getValues()-shaped array
 * (one column) and returns a Set of `String(v).trim()`, skipping empty,
 * null and undefined cells.
 */
function jobHuntersBuildExistingIdSet(idColumnValues) {
  const set = new Set();

  (idColumnValues || []).forEach(function (row) {
    const raw = row && row[0];
    if (raw === null || raw === undefined) {
      return;
    }
    const trimmed = String(raw).trim();
    if (trimmed.length > 0) {
      set.add(trimmed);
    }
  });

  return set;
}

/**
 * jobHuntersFindLastDataRow — pure (quick-261016-nu9 round 2 fix). Takes the
 * id column's DATA rows only (a 2D getValues()-shaped array for sheet rows
 * 2..2+N-1 -- the header row at sheet row 1 must already be excluded by the
 * caller) and returns the SHEET ROW NUMBER of the last row whose id cell is
 * actually non-empty after trimming, scanning from the bottom so a long
 * mostly-empty tail (e.g. checkbox-formatted but dataless rows) costs no
 * more than finding that one real row. Returns 1 (the header row) when NO
 * row has a non-empty id, meaning the caller should start writing at row 2.
 * Deliberately independent of sheet.getLastRow(), which also counts a cell
 * carrying checkbox/data-validation formatting as non-empty even with no
 * real value in it -- see this file's class-level "WHY NOT getLastRow()"
 * note.
 */
function jobHuntersFindLastDataRow(idColumnDataRows) {
  const rows = idColumnDataRows || [];

  for (let i = rows.length - 1; i >= 0; i--) {
    const raw = rows[i] && rows[i][0];
    if (raw !== null && raw !== undefined && String(raw).trim().length > 0) {
      return i + 2;
    }
  }

  return 1;
}

/**
 * jobHuntersMessageTimeForSort — pure sort-key helper for
 * selectLatestJobHunterMessages: a message's date in epoch ms, or
 * -Infinity when unreadable/invalid -- sorts such a message LAST under a
 * newest-first sort. Never throws.
 */
function jobHuntersMessageTimeForSort(message) {
  try {
    const date = message.getDate();
    if (!(date instanceof Date)) {
      return -Infinity;
    }
    const time = date.getTime();
    return Number.isNaN(time) ? -Infinity : time;
  } catch (e) {
    return -Infinity;
  }
}

/**
 * selectLatestJobHunterMessages — pure (D-09, P-16). Flattens every
 * thread's getMessages() into one array, sorts newest-first by date
 * (unreadable dates sort last), and returns the first `count`.
 */
function selectLatestJobHunterMessages(threads, count) {
  const allMessages = [];

  (threads || []).forEach(function (thread) {
    (thread.getMessages() || []).forEach(function (message) {
      allMessages.push(message);
    });
  });

  allMessages.sort(function (a, b) {
    return jobHuntersMessageTimeForSort(b) - jobHuntersMessageTimeForSort(a);
  });

  return allMessages.slice(0, count);
}

/**
 * jobHuntersFormatDryRunRow — pure (D-09). Returns 12 "<name>: <value>"
 * lines for a canonical row, in JOB_HUNTERS_COLUMN_NAMES order. The body
 * line is cut to JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS chars.
 */
function jobHuntersFormatDryRunRow(row) {
  return JOB_HUNTERS_COLUMN_NAMES.map(function (name, idx) {
    const value = row[idx];
    if (name === 'body') {
      return 'body: ' + String(value).slice(0, JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS);
    }
    return name + ': ' + String(value);
  });
}

/**
 * jobHuntersDescribeRunResult — pure (P-15). Formats a runJobHuntersCapture
 * result into the one human-readable summary line the webapp Run Now
 * handler and testJobHuntersDryRun's caller both use.
 */
function jobHuntersDescribeRunResult(result) {
  if (result.ran !== true) {
    return 'JobHunters did not run: ' + result.reason + '.';
  }

  return (
    'JobHunters ran: ' +
    result.rowsWritten +
    ' row(s) written, ' +
    result.threadsSucceeded +
    ' thread(s) relabeled, ' +
    result.threadsFailed +
    ' thread(s) failed.'
  );
}

/**
 * jobHuntersIsSpreadsheetIdConfigured — pure (P-10). false for null,
 * undefined, '', whitespace-only, or the committed sentinel marker; true
 * for anything else.
 */
function jobHuntersIsSpreadsheetIdConfigured(spreadsheetId) {
  if (spreadsheetId === null || spreadsheetId === undefined) {
    return false;
  }

  const trimmed = String(spreadsheetId).trim();
  if (trimmed.length === 0 || trimmed === JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER) {
    return false;
  }

  return true;
}

// Node/GAS environment bridge. Under GAS's shared global scope,
// JOB_HUNTERS_ACTION_CONFIG/JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY/
// JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER (sibling cfg file) and
// orderThreadsForProcessing (src/02-main.js) are already visible here by
// bare name -- this `if` block never executes there. Under Node, each
// require()d file is its own isolated module, so this bridges them onto
// globalThis. Deliberately does NOT bridge getOrCreateLabel or
// notifyOwnerOfFailure: the Node test harness injects its own spies for
// those, and that injection must stay authoritative.
if (typeof module !== 'undefined' && module.exports) {
  const cfg = require('./11-action-cfg-job-hunters.js');
  globalThis.JOB_HUNTERS_ACTION_CONFIG = cfg.JOB_HUNTERS_ACTION_CONFIG;
  globalThis.JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY = cfg.JOB_HUNTERS_LAST_RUN_AT_MS_PROPERTY_KEY;
  globalThis.JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER = cfg.JOB_HUNTERS_SPREADSHEET_ID_UNSET_MARKER;

  const mainExports = require('./02-main.js');
  globalThis.orderThreadsForProcessing = mainExports.orderThreadsForProcessing;
}

// GAS-safe Node export (inert under the Apps Script runtime). Every
// INTERFACE CONTRACT name for this file. runJobHuntersCapture is exported
// for the Node harness even though it is GAS-only, the same precedent as
// processTicketFromMessageBody. testJobHuntersDryRun is NOT exported --
// it is GAS-editor-only and shares no logic with the tested path.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    JOB_HUNTERS_COLUMN_NAMES: JOB_HUNTERS_COLUMN_NAMES,
    JOB_HUNTERS_BODY_MAX_CHARS: JOB_HUNTERS_BODY_MAX_CHARS,
    JOB_HUNTERS_DRY_RUN_MESSAGE_COUNT: JOB_HUNTERS_DRY_RUN_MESSAGE_COUNT,
    JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS: JOB_HUNTERS_DRY_RUN_BODY_PREVIEW_CHARS,
    JOB_HUNTERS_LOCK_TIMEOUT_MS: JOB_HUNTERS_LOCK_TIMEOUT_MS,
    jobHuntersIsDue: jobHuntersIsDue,
    jobHuntersParseFromHeader: jobHuntersParseFromHeader,
    jobHuntersNormalizeAddressList: jobHuntersNormalizeAddressList,
    cleanJobHunterBodyText: cleanJobHunterBodyText,
    jobHuntersIsBulkEmail: jobHuntersIsBulkEmail,
    buildJobHunterRow: buildJobHunterRow,
    resolveJobHuntersColumnIndexes: resolveJobHuntersColumnIndexes,
    jobHuntersColumnSegments: jobHuntersColumnSegments,
    jobHuntersRowsForSegment: jobHuntersRowsForSegment,
    jobHuntersBuildExistingIdSet: jobHuntersBuildExistingIdSet,
    jobHuntersFindLastDataRow: jobHuntersFindLastDataRow,
    selectLatestJobHunterMessages: selectLatestJobHunterMessages,
    jobHuntersFormatDryRunRow: jobHuntersFormatDryRunRow,
    jobHuntersDescribeRunResult: jobHuntersDescribeRunResult,
    jobHuntersIsSpreadsheetIdConfigured: jobHuntersIsSpreadsheetIdConfigured,
    runJobHuntersCapture: runJobHuntersCapture,
    jobHuntersNextRunInfo: jobHuntersNextRunInfo,
    jobHuntersDescribeNextRun: jobHuntersDescribeNextRun,
    jobHuntersGetNextRunInfo: jobHuntersGetNextRunInfo,
  };
}

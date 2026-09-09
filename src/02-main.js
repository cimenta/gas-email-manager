// Version: 0.16.2
/**
 * APP_VERSION — the running application version, rendered next to the admin
 * web app's page title via webappGetVersion (src/00-webapp.js, D-01). This
 * const and the literal `// Version:` comment on line 1 above are TWO
 * INDEPENDENT REPRESENTATIONS of the same fact: push-public.bat parses that
 * comment text via `findstr /C:"// Version:" src\02-main.js` (token 3) to
 * detect the release tag, while this const is what the web app actually
 * reads and displays at runtime. They MUST be bumped together on every
 * future release -- test/app-version.test.js is the guard that fails the
 * whole suite the moment they ever drift apart.
 */
const APP_VERSION = '0.16.2';

/**
 * orderThreadsForProcessing — CAUSAL ORDERING GUARANTEE (live-reported bug
 * regiojet-cancel-not-deleted). Returns a NEW array of the same threads
 * sorted OLDEST-FIRST by `getLastMessageDate()`, i.e. in the order the mail
 * actually arrived. Never mutates the input; a null/undefined input returns
 * [].
 *
 * WHY THIS EXISTS: processEmails used to consume `GmailApp.search()`'s result
 * array in whatever order Gmail handed it back. That order is NOT DOCUMENTED
 * anywhere in the Apps Script reference — empirically it mirrors the Gmail UI,
 * i.e. NEWEST FIRST. Nothing in this codebase ever established it, and nothing
 * would notice if Google changed it.
 *
 * That silence was load-bearing. Every cancellation-capable action here
 * (TRANSPORT_TICKETS_ACTION's cancelTransportTicketEvent,
 * BOOKING_MANAGEMENT_ACTION's handleCancellation) deletes an event that an
 * EARLIER email created. A cancellation is by definition NEWER than the
 * confirmation it cancels, so under newest-first the cancellation is
 * dispatched FIRST — it finds no event to delete, takes its documented silent
 * no-op branch, and the confirmation, processed second, then creates an event
 * that nothing will ever delete. The cancelled trip stays on the calendar
 * forever, with no error and no failure label. This is exactly the live
 * RegioJet report: a ticket bought and cancelled inside ONE trigger window
 * left both emails unprocessed for the same run.
 *
 * Sorting by LAST message date (not first) is deliberate: it is when the
 * thread last became relevant, which is what determines whether its newest
 * message is a confirmation or a cancellation.
 *
 * A thread whose date cannot be read (missing accessor, null, unparseable)
 * sorts FIRST rather than last — fail toward "create before cancel", never
 * toward letting an unknown date push a thread behind a cancellation. Pure
 * (touches only the passed-in objects), so it is Node-testable — see
 * test/thread-processing-order.test.js.
 */
function orderThreadsForProcessing(threads) {
  const list = (threads || []).slice();

  return list.sort(function (a, b) {
    return threadLastMessageTime(a) - threadLastMessageTime(b);
  });
}

/**
 * threadLastMessageTime — orderThreadsForProcessing's sort key: the thread's
 * last-message time in epoch milliseconds, or 0 when it cannot be determined
 * (no accessor, a throwing accessor, a null/invalid Date). 0 sorts such a
 * thread first — see orderThreadsForProcessing's own JSDoc for why the
 * unknown case must fail toward EARLIER, never later. Pure, never throws.
 */
function threadLastMessageTime(thread) {
  try {
    if (!thread || typeof thread.getLastMessageDate !== 'function') {
      return 0;
    }

    const date = thread.getLastMessageDate();
    if (!(date instanceof Date)) {
      return 0;
    }

    const time = date.getTime();
    return Number.isNaN(time) ? 0 : time;
  } catch (e) {
    return 0;
  }
}

/**
 * processEmails — the named target of the time-driven trigger installed by
 * setup() (see 01-setup.js). Searches Gmail for recent, unprocessed threads,
 * dispatches each through the pluggable action registry, and labels each
 * thread by outcome. Already-labeled threads are excluded from selection, so
 * a handled thread (success or failure) is never reprocessed.
 *
 * Threads are dispatched OLDEST-FIRST via orderThreadsForProcessing (see its
 * JSDoc for the live cancellation bug that made this explicit rather than
 * inherited from GmailApp.search's undocumented order). Ordering changes only
 * the sequence, never the membership: every selected thread is still
 * processed exactly once and labeled exactly as before.
 */
function processEmails() {
  const query = 'newer_than:' + CONFIG.daysBack + 'd -label:"' + CONFIG.labelName + '"';
  const threads = orderThreadsForProcessing(GmailApp.search(query));

  threads.forEach(function (thread) {
    const result = dispatchActions(thread);

    thread.addLabel(getOrCreateLabel(CONFIG.labelName));

    if (result.hadError) {
      thread.addLabel(getOrCreateLabel(CONFIG.failedLabelName));
    }
  });
}

/**
 * getOrCreateLabel — idempotent get-or-create for a Gmail user label.
 *
 * Shared by 01-setup.js (eager provisioning at setup() time) and
 * processEmails() above (runtime labeling), so labeling stays robust even if
 * a label is later deleted out-of-band. Follows a create-if-absent pattern
 * (GmailApp.getUserLabelByName null-check before createLabel()).
 */
function getOrCreateLabel(name) {
  const existingLabel = GmailApp.getUserLabelByName(name);

  if (existingLabel === null) {
    return GmailApp.createLabel(name);
  }

  return existingLabel;
}

/**
 * Notifications — owner failure notification.
 *
 * composeFailureBody is pure (no GAS globals) so it is Node-testable per the
 * guarded export below; notifyOwnerOfFailure is GAS-only wiring (Session +
 * MailApp).
 *
 * Trust boundary: `errorMessage` and `threadSubject` may be influenced by
 * untrusted inbound email content (e.g. an attacker-chosen subject, or a
 * parser error string). The body is PLAIN TEXT ONLY (no HTML body) so there
 * is no HTML/script execution context for that content to land in, and the
 * recipient/subject are never attacker-controlled.
 */

/**
 * composeFailureBody — pure plain-text description of an action failure:
 * which action failed, on which thread, and the error. No HTML markup.
 */
function composeFailureBody(actionName, threadSubject, errorMessage) {
  return [
    'A GAS Email Manager action failed while processing an email.',
    '',
    'Action: ' + actionName,
    'Email subject: ' + threadSubject,
    'Error: ' + errorMessage,
    '',
    'The email has been labeled as failed. No further automatic retry will occur.',
  ].join('\n');
}

/**
 * notifyOwnerOfFailure — resolves the script owner via
 * Session.getActiveUser().getEmail() and sends a fixed-subject, plain-text
 * failure email. Never sends HTML, never sends to any recipient other than
 * the owner. The send is wrapped in its own try/catch so a mail failure can
 * never break dispatchActions' per-action isolation invariant; if the
 * recipient resolves empty, the send is skipped.
 */
function notifyOwnerOfFailure(actionName, thread, error) {
  try {
    const recipient = Session.getActiveUser().getEmail();
    if (!recipient) {
      return;
    }

    const threadSubject = thread.getFirstMessageSubject();
    const subject = 'GAS Email Manager: action failed';
    const body = composeFailureBody(actionName, threadSubject, String(error));

    MailApp.sendEmail(recipient, subject, body);
  } catch (notifyError) {
    console.error('notifyOwnerOfFailure failed to send: ' + notifyError);
  }
}

// GAS-safe Node export (inert under the Apps Script runtime).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    composeFailureBody: composeFailureBody,
    orderThreadsForProcessing: orderThreadsForProcessing,
    APP_VERSION: APP_VERSION,
  };
}

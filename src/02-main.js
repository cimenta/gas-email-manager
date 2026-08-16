// Version: 0.13.0
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
const APP_VERSION = '0.13.0';

/**
 * processEmails — the named target of the time-driven trigger installed by
 * setup() (see 01-setup.js). Searches Gmail for recent, unprocessed threads,
 * dispatches each through the pluggable action registry, and labels each
 * thread by outcome. Already-labeled threads are excluded from selection, so
 * a handled thread (success or failure) is never reprocessed.
 */
function processEmails() {
  const query = 'newer_than:' + CONFIG.daysBack + 'd -label:"' + CONFIG.labelName + '"';
  const threads = GmailApp.search(query);

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
  module.exports = { composeFailureBody: composeFailureBody, APP_VERSION: APP_VERSION };
}

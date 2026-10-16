'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { composeFailureBody, composeTicketAttachmentFailureBody, notifyOwnerOfFailure } = require('../src/02-main.js');

test('composeFailureBody includes action name, thread subject, and error text as plain text', () => {
  const body = composeFailureBody(
    'ics-calendar-import',
    'Team offsite invite',
    'Error: Missing DTSTART in VEVENT 2'
  );

  assert.equal(typeof body, 'string');
  assert.ok(body.includes('ics-calendar-import'));
  assert.ok(body.includes('Team offsite invite'));
  assert.ok(body.includes('Error: Missing DTSTART in VEVENT 2'));
  assert.equal(/[<>]/.test(body), false, 'body must not contain HTML angle brackets');
});

// --- composeTicketAttachmentFailureBody -------------------------------------
//
// ROUND 2 (debug/entradio-portal-not-supported). A SEPARATE notification from
// composeFailureBody, not a reuse of it, for one concrete reason: that body
// says the email "has been labeled as failed" and that "no further automatic
// retry will occur". Both sentences would be FALSE here -- the calendar event
// WAS created and the thread WILL be labeled processed. Only the attachments
// are missing. Telling the owner their event failed when it did not is worse
// than sending nothing, because it invites them to re-run a thread that the
// dedup safety net will then correctly no-op.

test('composeTicketAttachmentFailureBody: names the event, the calendar and the order/ticket identifier so the owner can find the event and download manually', () => {
  const body = composeTicketAttachmentFailureBody('ČERNO, VÍR', 'kino@group.calendar.google.com', '2354152');

  assert.equal(typeof body, 'string');
  assert.ok(body.includes('ČERNO, VÍR'));
  assert.ok(body.includes('kino@group.calendar.google.com'));
  assert.ok(body.includes('2354152'));
});

test('composeTicketAttachmentFailureBody: states the event WAS created -- it must never read as a failure of the event itself', () => {
  const body = composeTicketAttachmentFailureBody('ČERNO, VÍR', 'kino@group.calendar.google.com', '2354152');

  assert.equal(body.includes('was created'), true);
  // The two sentences from composeFailureBody that would be untrue here.
  assert.equal(body.includes('labeled as failed'), false);
  assert.equal(body.includes('No further automatic retry'), false);
});

test('composeTicketAttachmentFailureBody: tells the owner to download the tickets manually', () => {
  const body = composeTicketAttachmentFailureBody('ČERNO, VÍR', 'kino@group.calendar.google.com', '2354152');

  assert.equal(body.toLowerCase().includes('manual'), true);
});

// The SAME assertion composeFailureBody's own test makes, on purpose. The
// security property is that this composer emits NO HTML MARKUP OF ITS OWN, so
// that notifyOwnerOfTicketAttachmentFailure's three-argument MailApp.sendEmail
// (no htmlBody) really is a plain-text send -- there is then no HTML execution
// context for untrusted inbound content to land in.
//
// It is deliberately NOT an input-sanitizing assertion: an earlier draft of
// this test demanded that angle brackets in the EVENT NAME be stripped, which
// is both stronger than anything else in this codebase does and actively
// wrong, since it would corrupt a legitimate event title containing "<" or
// ">". Plain-text delivery is the guarantee; escaping the payload is not.
test('composeTicketAttachmentFailureBody: emits no HTML markup of its own -- the body is a plain-text send, with no HTML context for untrusted content to land in', () => {
  const body = composeTicketAttachmentFailureBody('ČERNO, VÍR', 'kino@group.calendar.google.com', '2354152');

  assert.equal(/[<>]/.test(body), false, 'body must not contain HTML angle brackets');
});

test('composeTicketAttachmentFailureBody: a parser that extracted no ticket identifier still yields a usable body, never the literal word "null"', () => {
  const body = composeTicketAttachmentFailureBody('ČERNO, VÍR', 'kino@group.calendar.google.com', null);

  assert.ok(body.includes('ČERNO, VÍR'));
  assert.equal(body.includes('null'), false);
});

// --- notifyOwnerOfFailure: null-thread tolerance (quick-261016-j0d, D-11, --
// --- P-14) -- JobHunters' run-level failure (D-11) has no single thread to
// --- blame, so notifyOwnerOfFailure must accept thread=null and still send,
// --- using a fixed placeholder subject line instead of crashing on
// --- thread.getFirstMessageSubject(). ---------------------------------------

function notificationFakeThread(subject) {
  return {
    getFirstMessageSubject: function () {
      return subject;
    },
  };
}

function withNotifyOwnerOfFailureGlobals(fn) {
  const sentEmails = [];

  const savedSession = Object.prototype.hasOwnProperty.call(global, 'Session');
  const previousSession = global.Session;
  const savedMailApp = Object.prototype.hasOwnProperty.call(global, 'MailApp');
  const previousMailApp = global.MailApp;

  global.Session = {
    getActiveUser: function () {
      return { getEmail: function () { return 'owner@example.invalid'; } };
    },
  };
  global.MailApp = {
    sendEmail: function (recipient, subject, body) {
      sentEmails.push({ recipient: recipient, subject: subject, body: body });
    },
  };

  try {
    fn(sentEmails);
  } finally {
    if (savedSession) {
      global.Session = previousSession;
    } else {
      delete global.Session;
    }
    if (savedMailApp) {
      global.MailApp = previousMailApp;
    } else {
      delete global.MailApp;
    }
  }
}

test('quick-261016-j0d (D-11, P-14): notifyOwnerOfFailure tolerates a null thread -- sends one email with the fixed run-level placeholder subject text', () => {
  withNotifyOwnerOfFailureGlobals((sentEmails) => {
    notifyOwnerOfFailure('JobHunters', null, new Error('Spreadsheet unreachable'));

    assert.equal(sentEmails.length, 1);
    assert.equal(sentEmails[0].recipient, 'owner@example.invalid');
    assert.equal(sentEmails[0].subject, 'GAS Email Manager: action failed');
    assert.ok(sentEmails[0].body.includes('Action: JobHunters'));
    assert.ok(sentEmails[0].body.includes('Spreadsheet unreachable'));
    assert.ok(sentEmails[0].body.includes('Email subject: (none -- run-level failure, not tied to a single email)'));
  });
});

test('quick-261016-j0d (D-11): notifyOwnerOfFailure regression -- a real thread still reports its own subject', () => {
  withNotifyOwnerOfFailureGlobals((sentEmails) => {
    const thread = notificationFakeThread('Interview invitation');
    notifyOwnerOfFailure('JobHunters', thread, new Error('boom'));

    assert.equal(sentEmails.length, 1);
    assert.ok(sentEmails[0].body.includes('Email subject: Interview invitation'));
  });
});

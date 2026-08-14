'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  mojemenickaExtractEmailAddress,
  isMojemenickaAllowedSender,
  subjectContainsAnyMojemenickaString,
  resolveMojemenickaRequestedRange,
  assertMojemenickaMenuUrlIsFetchable,
  buildMojemenickaMenuRequestUrl,
  resolveMojemenickaRequestMessages,
  fetchMojemenickaMenuHtml,
  buildMojemenickaFallbackHtml,
  MOJEMENICKA_ACTION,
} = require('../src/09-action-mojemenicka.js');
const { MOJEMENICKA_ACTION_CONFIG } = require('../src/09-action-cfg-mojemenicka.js');
const { dispatchActions } = require('../src/03-action-management.js');

// --- test helpers: fake PropertiesService / UrlFetchApp, installed on -------
// --- global and always torn down in a try/finally (order-independence), ----
// --- mirroring test/script-properties.test.js's withFakePropertiesService --
// --- and test/action-config-split.test.js's equivalent. --------------------

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

function withFakeUrlFetchApp(fetchImpl, fn) {
  const calls = [];
  global.UrlFetchApp = {
    fetch: function (url, options) {
      calls.push({ url: url, options: options });
      return fetchImpl(url, options);
    },
  };

  try {
    fn(calls);
  } finally {
    delete global.UrlFetchApp;
  }
}

function makeFakeResponse(code, body) {
  return {
    getResponseCode: function () {
      return code;
    },
    getContentText: function () {
      return body;
    },
  };
}

function makeFakeMessage(from, subject) {
  const replyCalls = [];
  return {
    getFrom: function () {
      return from;
    },
    getSubject: function () {
      return subject;
    },
    reply: function (body, options) {
      replyCalls.push({ body: body, options: options });
    },
    getReplyCalls: function () {
      return replyCalls;
    },
  };
}

function makeFakeThread(messages) {
  return {
    getMessages: function () {
      return messages;
    },
    getId: function () {
      return 'thread-mojemenicka-test';
    },
  };
}

// --- mojemenickaExtractEmailAddress ------------------------------------------

test('mojemenickaExtractEmailAddress: angle-bracket header returns the bare lowercased address', () => {
  assert.equal(mojemenickaExtractEmailAddress('Jana Nováková <J.Novakova@Example.COM>'), 'j.novakova@example.com');
});

test('mojemenickaExtractEmailAddress: bare address (no display name) returns itself lowercased', () => {
  assert.equal(mojemenickaExtractEmailAddress('Jana@Example.com'), 'jana@example.com');
});

test('mojemenickaExtractEmailAddress: surrounding whitespace trimmed', () => {
  assert.equal(mojemenickaExtractEmailAddress('  jana@example.com  '), 'jana@example.com');
});

test('mojemenickaExtractEmailAddress: whitespace inside brackets trimmed', () => {
  assert.equal(mojemenickaExtractEmailAddress('Name < jana@example.com >'), 'jana@example.com');
});

test('mojemenickaExtractEmailAddress: null returns empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(mojemenickaExtractEmailAddress(null), '');
  });
});

test('mojemenickaExtractEmailAddress: undefined returns empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(mojemenickaExtractEmailAddress(undefined), '');
  });
});

test('mojemenickaExtractEmailAddress: empty string returns empty string without throwing', () => {
  assert.doesNotThrow(() => {
    assert.equal(mojemenickaExtractEmailAddress(''), '');
  });
});

// --- isMojemenickaAllowedSender (FAIL-CLOSED — the opposite of --------------
// --- ICS_ACTION_CONFIG.importOnlyFrom's empty-means-everyone rule) ----------

test('isMojemenickaAllowedSender: exact address match, case-insensitive on the LIST side', () => {
  assert.equal(isMojemenickaAllowedSender('a@x.com', ['A@X.com']), true);
});

test('isMojemenickaAllowedSender: exact address match, case-insensitive on the SENDER side', () => {
  assert.equal(isMojemenickaAllowedSender('A@X.com', ['a@x.com']), true);
});

test('isMojemenickaAllowedSender: display-name-wrapped From header matches', () => {
  assert.equal(isMojemenickaAllowedSender('Jana Nováková <jana@example.com>', ['jana@example.com']), true);
});

test('isMojemenickaAllowedSender: sender NOT in the list returns false', () => {
  assert.equal(isMojemenickaAllowedSender('stranger@example.com', ['jana@example.com']), false);
});

test('isMojemenickaAllowedSender: empty array returns false (fail-closed)', () => {
  assert.equal(isMojemenickaAllowedSender('jana@example.com', []), false);
});

test('isMojemenickaAllowedSender: null list returns false (fail-closed)', () => {
  assert.equal(isMojemenickaAllowedSender('jana@example.com', null), false);
});

test('isMojemenickaAllowedSender: undefined list returns false (fail-closed)', () => {
  assert.equal(isMojemenickaAllowedSender('jana@example.com', undefined), false);
});

// --- subjectContainsAnyMojemenickaString (D-08/D-09 shared OR matcher) ------

test('subjectContainsAnyMojemenickaString: matches a single-entry list, case-insensitive substring (D-08)', () => {
  assert.equal(subjectContainsAnyMojemenickaString('dnesni menu prosim', ['menu']), true);
});

test('subjectContainsAnyMojemenickaString: OR across a list — the SECOND entry matches', () => {
  assert.equal(subjectContainsAnyMojemenickaString('Prosim JIDELNICEK', ['menu', 'jidelnicek']), true);
});

test('subjectContainsAnyMojemenickaString: no entry in the list matches returns false', () => {
  assert.equal(subjectContainsAnyMojemenickaString('ahoj jak se mas', ['menu', 'jidelnicek']), false);
});

test('subjectContainsAnyMojemenickaString: empty candidates list is fail-closed', () => {
  assert.equal(subjectContainsAnyMojemenickaString('dnesni menu', []), false);
});

test('subjectContainsAnyMojemenickaString: null candidates is fail-closed', () => {
  assert.equal(subjectContainsAnyMojemenickaString('dnesni menu', null), false);
});

test('subjectContainsAnyMojemenickaString: undefined candidates is fail-closed', () => {
  assert.equal(subjectContainsAnyMojemenickaString('dnesni menu', undefined), false);
});

test('subjectContainsAnyMojemenickaString: blank entries are skipped, never treated as a match-everything wildcard', () => {
  assert.equal(subjectContainsAnyMojemenickaString('dnesni menu', ['', '   ']), false);
});

test('subjectContainsAnyMojemenickaString: a blank entry does not poison a list that also holds a real entry', () => {
  assert.equal(subjectContainsAnyMojemenickaString('dnesni menu', ['', 'menu']), true);
});

test('subjectContainsAnyMojemenickaString: null subject returns false', () => {
  assert.equal(subjectContainsAnyMojemenickaString(null, ['menu']), false);
});

test('subjectContainsAnyMojemenickaString: undefined subject returns false', () => {
  assert.equal(subjectContainsAnyMojemenickaString(undefined, ['menu']), false);
});

test('subjectContainsAnyMojemenickaString: empty subject returns false', () => {
  assert.equal(subjectContainsAnyMojemenickaString('', ['menu']), false);
});

// --- resolveMojemenickaRequestedRange (D-09) ---------------------------------

test('resolveMojemenickaRequestedRange: subject carrying a weekly-trigger string returns "week"', () => {
  assert.equal(resolveMojemenickaRequestedRange('menu na tyden prosim', ['tyden']), 'week');
});

test('resolveMojemenickaRequestedRange: case-insensitive, same matcher as D-08', () => {
  assert.equal(resolveMojemenickaRequestedRange('menu na TYDEN prosim', ['tyden']), 'week');
});

test('resolveMojemenickaRequestedRange: subject without a weekly-trigger string returns "today"', () => {
  assert.equal(resolveMojemenickaRequestedRange('menu prosim', ['tyden']), 'today');
});

test('resolveMojemenickaRequestedRange: an unset (empty) weekly list means every request is a today request', () => {
  assert.equal(resolveMojemenickaRequestedRange('menu na tyden prosim', []), 'today');
});

test('resolveMojemenickaRequestedRange: a null weekly list means every request is a today request', () => {
  assert.equal(resolveMojemenickaRequestedRange('menu na tyden prosim', null), 'today');
});

test('resolveMojemenickaRequestedRange: OR across the weekly list', () => {
  assert.equal(resolveMojemenickaRequestedRange('menu prosim', ['tyden', 'weekly']), 'today');
  assert.equal(resolveMojemenickaRequestedRange('menu weekly', ['tyden', 'weekly']), 'week');
});

// --- assertMojemenickaMenuUrlIsFetchable (D-10) ------------------------------

test('assertMojemenickaMenuUrlIsFetchable: a bare base URL does not throw', () => {
  assert.doesNotThrow(() => assertMojemenickaMenuUrlIsFetchable('https://example.com/menu'));
});

test('assertMojemenickaMenuUrlIsFetchable: null/undefined/empty/whitespace-only each throw, naming the setting key', () => {
  [null, undefined, '', '   '].forEach(function (badValue) {
    assert.throws(() => assertMojemenickaMenuUrlIsFetchable(badValue), /09-action-mojemenicka-MENU_URL/);
  });
});

test('assertMojemenickaMenuUrlIsFetchable: a URL already carrying a query string throws, naming the setting key', () => {
  assert.throws(() => assertMojemenickaMenuUrlIsFetchable('https://example.com/menu?lang=cs'), /09-action-mojemenicka-MENU_URL/);
});

test('assertMojemenickaMenuUrlIsFetchable: a trailing bare "?" also throws', () => {
  assert.throws(() => assertMojemenickaMenuUrlIsFetchable('https://example.com/menu?'), /09-action-mojemenicka-MENU_URL/);
});

// --- buildMojemenickaMenuRequestUrl (D-10) -----------------------------------

test('buildMojemenickaMenuRequestUrl: today range appends the exact today query string', () => {
  assert.equal(buildMojemenickaMenuRequestUrl('https://example.com/menu', 'today'), 'https://example.com/menu?format=html&range=today');
});

test('buildMojemenickaMenuRequestUrl: week range appends the exact week query string', () => {
  assert.equal(buildMojemenickaMenuRequestUrl('https://example.com/menu', 'week'), 'https://example.com/menu?format=html&range=week');
});

test('buildMojemenickaMenuRequestUrl: delegates to the assert — a URL with an existing query string throws', () => {
  assert.throws(() => buildMojemenickaMenuRequestUrl('https://example.com/menu?lang=cs', 'today'), /09-action-mojemenicka-MENU_URL/);
});

// --- resolveMojemenickaRequestMessages ---------------------------------------

test('resolveMojemenickaRequestMessages: returns only messages satisfying BOTH conditions', () => {
  const matching = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
  const wrongSender = makeFakeMessage('stranger@example.com', 'dnesni menu prosim');
  const wrongSubject = makeFakeMessage('jana@example.com', 'ahoj');
  const messages = [wrongSender, matching, wrongSubject];

  const result = resolveMojemenickaRequestMessages(messages, ['jana@example.com'], ['menu']);

  assert.deepEqual(result, [matching]);
});

test('resolveMojemenickaRequestMessages: returns [] when none match', () => {
  const messages = [makeFakeMessage('stranger@example.com', 'ahoj')];
  assert.deepEqual(resolveMojemenickaRequestMessages(messages, ['jana@example.com'], ['menu']), []);
});

test('resolveMojemenickaRequestMessages: preserves input order', () => {
  const first = makeFakeMessage('jana@example.com', 'menu prosim');
  const second = makeFakeMessage('jana@example.com', 'jeste jednou menu');
  const result = resolveMojemenickaRequestMessages([first, second], ['jana@example.com'], ['menu']);
  assert.deepEqual(result, [first, second]);
});

test('resolveMojemenickaRequestMessages: a message matching only the sender is excluded', () => {
  const message = makeFakeMessage('jana@example.com', 'ahoj, jak se mas');
  assert.deepEqual(resolveMojemenickaRequestMessages([message], ['jana@example.com'], ['menu']), []);
});

test('resolveMojemenickaRequestMessages: a message matching only the subject is excluded', () => {
  const message = makeFakeMessage('stranger@example.com', 'dnesni menu prosim');
  assert.deepEqual(resolveMojemenickaRequestMessages([message], ['jana@example.com'], ['menu']), []);
});

test('resolveMojemenickaRequestMessages: a message whose subject matches the SECOND trigger string is returned', () => {
  const message = makeFakeMessage('jana@example.com', 'prosim jidelnicek');
  assert.deepEqual(resolveMojemenickaRequestMessages([message], ['jana@example.com'], ['menu', 'jidelnicek']), [message]);
});

test('resolveMojemenickaRequestMessages: a message whose subject contains only a weekly-trigger string and no trigger string is NOT returned (D-09 subordination)', () => {
  const message = makeFakeMessage('jana@example.com', 'na tyden prosim');
  assert.deepEqual(resolveMojemenickaRequestMessages([message], ['jana@example.com'], ['menu']), []);
});

// --- MOJEMENICKA_ACTION descriptor: name / config identity -------------------

test('MOJEMENICKA_ACTION.name equals "mojemenicka"', () => {
  assert.equal(MOJEMENICKA_ACTION.name, 'mojemenicka');
});

test('MOJEMENICKA_ACTION.config is the SAME object reference as MOJEMENICKA_ACTION_CONFIG (getter identity, not a copy)', () => {
  assert.equal(MOJEMENICKA_ACTION.config, MOJEMENICKA_ACTION_CONFIG);
});

// --- MOJEMENICKA_ACTION.appliesTo: literal boolean ---------------------------

test('MOJEMENICKA_ACTION.appliesTo: returns the LITERAL boolean true when a matching message exists (second trigger string in the list)', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu, jidelnicek',
    },
    function () {
      const message = makeFakeMessage('jana@example.com', 'prosim jidelnicek');
      const thread = makeFakeThread([message]);
      assert.equal(MOJEMENICKA_ACTION.appliesTo(thread), true);
    }
  );
});

test('MOJEMENICKA_ACTION.appliesTo: returns the LITERAL boolean false when no matching message exists', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
    },
    function () {
      const message = makeFakeMessage('stranger@example.com', 'ahoj');
      const thread = makeFakeThread([message]);
      assert.equal(MOJEMENICKA_ACTION.appliesTo(thread), false);
    }
  );
});

// --- fetchMojemenickaMenuHtml (D-10: takes menuUrl + range) ------------------

test('fetchMojemenickaMenuHtml: returns the exact response body on a 200, today range, and fetches the exact today URL', () => {
  withFakeUrlFetchApp(
    function () {
      return makeFakeResponse(200, '<h1>Dnesni menu</h1>');
    },
    function (calls) {
      const result = fetchMojemenickaMenuHtml('https://example.com/menu', 'today');
      assert.equal(result, '<h1>Dnesni menu</h1>');
      assert.equal(calls.length, 1);
      assert.equal(calls[0].url, 'https://example.com/menu?format=html&range=today');
      assert.equal(calls[0].options.muteHttpExceptions, true);
    }
  );
});

test('fetchMojemenickaMenuHtml: week range fetches the exact week URL', () => {
  withFakeUrlFetchApp(
    function () {
      return makeFakeResponse(200, '<h1>Tydenni menu</h1>');
    },
    function (calls) {
      const result = fetchMojemenickaMenuHtml('https://example.com/menu', 'week');
      assert.equal(result, '<h1>Tydenni menu</h1>');
      assert.equal(calls.length, 1);
      assert.equal(calls[0].url, 'https://example.com/menu?format=html&range=week');
    }
  );
});

test('fetchMojemenickaMenuHtml: throws when the response code is 500', () => {
  withFakeUrlFetchApp(
    function () {
      return makeFakeResponse(500, '<h1>Dnesni menu</h1>');
    },
    function () {
      assert.throws(() => fetchMojemenickaMenuHtml('https://example.com/menu', 'today'));
    }
  );
});

test('fetchMojemenickaMenuHtml: throws when the response code is 404', () => {
  withFakeUrlFetchApp(
    function () {
      return makeFakeResponse(404, '<h1>Dnesni menu</h1>');
    },
    function () {
      assert.throws(() => fetchMojemenickaMenuHtml('https://example.com/menu', 'today'));
    }
  );
});

test('fetchMojemenickaMenuHtml: throws when the body is empty', () => {
  withFakeUrlFetchApp(
    function () {
      return makeFakeResponse(200, '');
    },
    function () {
      assert.throws(() => fetchMojemenickaMenuHtml('https://example.com/menu', 'today'));
    }
  );
});

test('fetchMojemenickaMenuHtml: throws when the body is whitespace-only', () => {
  withFakeUrlFetchApp(
    function () {
      return makeFakeResponse(200, '   ');
    },
    function () {
      assert.throws(() => fetchMojemenickaMenuHtml('https://example.com/menu', 'today'));
    }
  );
});

// --- THE END-TO-END TRACER TEST -----------------------------------------------

test('MOJEMENICKA_ACTION: end-to-end — allow-listed sender asks for the menu and gets it verbatim (D-01, D-02, D-03)', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Dnesni menu</h1>');
        },
        function () {
          const message = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
          const thread = makeFakeThread([message]);

          assert.equal(MOJEMENICKA_ACTION.appliesTo(thread), true);

          MOJEMENICKA_ACTION.run(thread);

          const replyCalls = message.getReplyCalls();
          assert.equal(replyCalls.length, 1);
          assert.equal(replyCalls[0].options.htmlBody, '<h1>Dnesni menu</h1>');
          assert.equal(replyCalls[0].body, '');
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION: end-to-end today — weekly list configured but subject has no weekly indicator fetches range=today (D-09)', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS': 'tyden',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Dnesni menu</h1>');
        },
        function (calls) {
          const message = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
          const thread = makeFakeThread([message]);

          MOJEMENICKA_ACTION.run(thread);

          const replyCalls = message.getReplyCalls();
          assert.equal(replyCalls.length, 1);
          assert.equal(replyCalls[0].options.htmlBody, '<h1>Dnesni menu</h1>');
          assert.ok(calls[0].url.indexOf('range=today') !== -1);
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION: end-to-end week — subject carrying a weekly-trigger string fetches range=week (D-09)', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS': 'tyden',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Tydenni menu</h1>');
        },
        function (calls) {
          const message = makeFakeMessage('jana@example.com', 'menu na tyden prosim');
          const thread = makeFakeThread([message]);

          MOJEMENICKA_ACTION.run(thread);

          const replyCalls = message.getReplyCalls();
          assert.equal(replyCalls.length, 1);
          assert.equal(replyCalls[0].options.htmlBody, '<h1>Tydenni menu</h1>');
          assert.ok(calls[0].url.indexOf('range=week') !== -1);
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION: D-09 subordination end-to-end — a weekly-string-only subject never applies, UrlFetchApp.fetch is never called', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS': 'tyden',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          throw new Error('UrlFetchApp.fetch should never be called');
        },
        function (calls) {
          const message = makeFakeMessage('jana@example.com', 'na tyden prosim');
          const thread = makeFakeThread([message]);

          assert.equal(MOJEMENICKA_ACTION.appliesTo(thread), false);
          assert.equal(calls.length, 0);
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION.run: range is resolved from the LAST matching message subject, not the first', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS': 'tyden',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Menu</h1>');
        },
        function (calls) {
          const first = makeFakeMessage('jana@example.com', 'menu prosim');
          const second = makeFakeMessage('jana@example.com', 'menu na tyden prosim');
          const thread = makeFakeThread([first, second]);

          MOJEMENICKA_ACTION.run(thread);

          assert.ok(calls[0].url.indexOf('range=week') !== -1);
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION.run: reversed message order fetches range=today from the now-last message', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS': 'tyden',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Menu</h1>');
        },
        function (calls) {
          const first = makeFakeMessage('jana@example.com', 'menu na tyden prosim');
          const second = makeFakeMessage('jana@example.com', 'menu prosim');
          const thread = makeFakeThread([first, second]);

          MOJEMENICKA_ACTION.run(thread);

          assert.ok(calls[0].url.indexOf('range=today') !== -1);
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION.run: multi-message thread — two matching messages, replies exactly ONCE, to the LAST matching message', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Dnesni menu</h1>');
        },
        function () {
          const first = makeFakeMessage('jana@example.com', 'menu prosim');
          const second = makeFakeMessage('jana@example.com', 'jeste jednou menu');
          const thread = makeFakeThread([first, second]);

          MOJEMENICKA_ACTION.run(thread);

          assert.equal(first.getReplyCalls().length, 0);
          assert.equal(second.getReplyCalls().length, 1);
        }
      );
    }
  );
});

test('MOJEMENICKA_ACTION.run: unset menuUrl throws an Error naming the 09-action-mojemenicka-MENU_URL setting', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
    },
    function () {
      const message = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
      const thread = makeFakeThread([message]);

      assert.throws(() => MOJEMENICKA_ACTION.run(thread), /09-action-mojemenicka-MENU_URL/);
    }
  );
});

test('MOJEMENICKA_ACTION.run: NEW misconfiguration — a menu URL with an existing query string THROWS, UrlFetchApp.fetch is never called, no reply is recorded (D-10)', () => {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu?lang=cs',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          throw new Error('UrlFetchApp.fetch should never be called');
        },
        function (calls) {
          const message = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
          const thread = makeFakeThread([message]);

          assert.throws(() => MOJEMENICKA_ACTION.run(thread), /09-action-mojemenicka-MENU_URL/);
          assert.equal(calls.length, 0);
          assert.equal(message.getReplyCalls().length, 0);
        }
      );
    }
  );
});

// --- buildMojemenickaFallbackHtml (Task 2, D-04) -----------------------------

test('buildMojemenickaFallbackHtml: returns a non-empty string containing at least one HTML tag', () => {
  const html = buildMojemenickaFallbackHtml();
  assert.equal(typeof html, 'string');
  assert.ok(html.length > 0);
  assert.match(html, /<[a-zA-Z][^>]*>/);
});

test('buildMojemenickaFallbackHtml: returns the SAME string on every call (fixed constant)', () => {
  assert.equal(buildMojemenickaFallbackHtml(), buildMojemenickaFallbackHtml());
});

test('buildMojemenickaFallbackHtml: information-disclosure guard — contains neither "http" nor the configured URL/error text (T-04-02)', () => {
  const html = buildMojemenickaFallbackHtml();
  const configuredUrl = 'https://example.com/secret-menu-endpoint';
  const errorText = 'ECONNREFUSED could not connect';

  assert.equal(html.toLowerCase().indexOf('http'), -1);
  assert.equal(html.indexOf(configuredUrl), -1);
  assert.equal(html.indexOf(errorText), -1);
});

// --- run fallback: four failure modes, D-04's graceful fallback -------------

function withMojemenickaRunFixture(fetchImpl, fn) {
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-WEEKLY_TRIGGER_STRINGS': 'tyden',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
    },
    function () {
      withFakeUrlFetchApp(fetchImpl, function () {
        const message = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
        const thread = makeFakeThread([message]);
        fn(message, thread);
      });
    }
  );
}

test('MOJEMENICKA_ACTION.run: fallback when UrlFetchApp.fetch itself throws — replies once with the fallback body, does not throw', () => {
  withMojemenickaRunFixture(
    function () {
      throw new Error('network unreachable');
    },
    function (message, thread) {
      assert.doesNotThrow(() => MOJEMENICKA_ACTION.run(thread));
      const replyCalls = message.getReplyCalls();
      assert.equal(replyCalls.length, 1);
      assert.equal(replyCalls[0].options.htmlBody, buildMojemenickaFallbackHtml());
    }
  );
});

test('MOJEMENICKA_ACTION.run: fallback on response code 500 — replies once with the fallback body, does not throw', () => {
  withMojemenickaRunFixture(
    function () {
      return makeFakeResponse(500, 'Internal Server Error');
    },
    function (message, thread) {
      assert.doesNotThrow(() => MOJEMENICKA_ACTION.run(thread));
      const replyCalls = message.getReplyCalls();
      assert.equal(replyCalls.length, 1);
      assert.equal(replyCalls[0].options.htmlBody, buildMojemenickaFallbackHtml());
    }
  );
});

test('MOJEMENICKA_ACTION.run: fallback on response code 404 — replies once with the fallback body, does not throw', () => {
  withMojemenickaRunFixture(
    function () {
      return makeFakeResponse(404, 'Not Found');
    },
    function (message, thread) {
      assert.doesNotThrow(() => MOJEMENICKA_ACTION.run(thread));
      const replyCalls = message.getReplyCalls();
      assert.equal(replyCalls.length, 1);
      assert.equal(replyCalls[0].options.htmlBody, buildMojemenickaFallbackHtml());
    }
  );
});

test('MOJEMENICKA_ACTION.run: fallback on 200 with empty body — replies once with the fallback body, does not throw', () => {
  withMojemenickaRunFixture(
    function () {
      return makeFakeResponse(200, '');
    },
    function (message, thread) {
      assert.doesNotThrow(() => MOJEMENICKA_ACTION.run(thread));
      const replyCalls = message.getReplyCalls();
      assert.equal(replyCalls.length, 1);
      assert.equal(replyCalls[0].options.htmlBody, buildMojemenickaFallbackHtml());
    }
  );
});

// --- D-05 integration proof: fallback path never notifies the owner --------

test('dispatchActions: a failing menu fetch (fallback path) returns hadError false and an empty errors array (D-05 — notifyOwnerOfFailure unreachable)', () => {
  withMojemenickaRunFixture(
    function () {
      throw new Error('network unreachable');
    },
    function (message, thread) {
      const result = dispatchActions(thread, [MOJEMENICKA_ACTION]);
      assert.equal(result.hadError, false);
      assert.deepEqual(result.errors, []);
    }
  );
});

test('dispatchActions: contrast case — when reply() ITSELF throws, the error propagates and dispatchActions reports hadError true naming mojemenicka', () => {
  // NOTIFY_ON_FAILURE forced false here purely as Node test isolation:
  // notifyOwnerOfFailure (src/02-main.js) is a GAS-only global not bridged
  // into src/03-action-management.js under Node (no test in this codebase
  // bridges it — see test/action-dispatch.test.js). This does not weaken
  // the assertion below: hadError/errors are populated by dispatchActions'
  // catch block BEFORE it checks notifyOnFailure, so the propagation claim
  // this test exists to prove is unaffected either way.
  withFakePropertiesService(
    {
      '09-action-mojemenicka-ALLOWED_SENDERS': 'jana@example.com',
      '09-action-mojemenicka-TRIGGER_STRINGS': 'menu',
      '09-action-mojemenicka-MENU_URL': 'https://example.com/menu',
      '09-action-mojemenicka-NOTIFY_ON_FAILURE': 'false',
    },
    function () {
      withFakeUrlFetchApp(
        function () {
          return makeFakeResponse(200, '<h1>Dnesni menu</h1>');
        },
        function () {
          const message = makeFakeMessage('jana@example.com', 'dnesni menu prosim');
          message.reply = function () {
            throw new Error('reply failed');
          };
          const thread = makeFakeThread([message]);

          const result = dispatchActions(thread, [MOJEMENICKA_ACTION]);

          assert.equal(result.hadError, true);
          assert.equal(result.errors.length, 1);
          assert.equal(result.errors[0].action, 'mojemenicka');
        }
      );
    }
  );
});

// --- Happy path unchanged: Task 1's tracer test still passes with the -------
// --- real fetched HTML, not the fallback body (regression guard) ------------

test('MOJEMENICKA_ACTION.run: happy path unchanged — real fetched HTML is used, not the fallback body', () => {
  withMojemenickaRunFixture(
    function () {
      return makeFakeResponse(200, '<h1>Dnesni menu</h1>');
    },
    function (message, thread) {
      MOJEMENICKA_ACTION.run(thread);
      const replyCalls = message.getReplyCalls();
      assert.equal(replyCalls.length, 1);
      assert.equal(replyCalls[0].options.htmlBody, '<h1>Dnesni menu</h1>');
      assert.notEqual(replyCalls[0].options.htmlBody, buildMojemenickaFallbackHtml());
    }
  );
});

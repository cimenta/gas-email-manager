/**
 * MOJEMENICKA_ACTION — Phase 4 (04-01) NEW ACTION, amended 04-04. Recognizes
 * menu-request emails from an allow-listed set of senders and replies
 * in-thread with a freshly-fetched menu.
 *
 * D-08 SUBJECT TRIGGER (amends 04-01's single-string D-01): appliesTo/run
 * only treat a message as a menu request when BOTH are true: the sender is
 * in config.allowedSenders (exact address match, case-insensitive both
 * sides — see mojemenickaExtractEmailAddress/isMojemenickaAllowedSender),
 * AND the message's SUBJECT line (never the body) contains ANY ONE of
 * config.triggerStrings, an OR-across-a-list check via
 * subjectContainsAnyMojemenickaString.
 *
 * D-09 TODAY/WEEK RANGE (new): once a message has already matched D-08
 * above, resolveMojemenickaRequestedRange checks that SAME message's
 * subject for any of config.weeklyTriggerStrings (via the same OR matcher)
 * to decide between the "week" and "today" range. This check is strictly
 * subordinate to D-08 — a weekly indicator alone, without a trigger string
 * also present, never makes the action apply.
 *
 * D-02/D-10 VERBATIM HTML, PER-RANGE FETCH: fetchMojemenickaMenuHtml builds
 * its request URL via buildMojemenickaMenuRequestUrl (menuUrl +
 * '?format=html&range=' + range) and its returned body is used directly,
 * unmodified, as the reply's htmlBody — no template wrapping, no text
 * extraction. A configured menuUrl that already carries a query string is a
 * misconfiguration, asserted against by assertMojemenickaMenuUrlIsFetchable.
 *
 * D-03 IN-THREAD REPLY: this is the first action in this codebase to reply
 * to the original sender (every other "email out" call is
 * notifyOwnerOfFailure's MailApp.sendEmail to the OWNER only — a
 * structurally different, non-thread-scoped, plain-text-only, fixed-
 * recipient call). run() uses GmailMessage.reply(body, options) directly.
 *
 * GLOBALLY-UNIQUE NAMING WARNING (see the booking.com/ticketing-portals
 * action files' own class-level JSDoc for the full incident this warning
 * originates from): Apps Script concatenates every project file into ONE
 * shared global scope — every top-level symbol here is prefixed
 * `mojemenicka`/`Mojemenicka`/`MOJEMENICKA`, this file's own
 * locally-reimplemented copy of the sender-parsing/matching helpers every
 * other action independently re-implements too, per the
 * one-file-per-action self-containment pattern.
 */

/**
 * mojemenickaExtractEmailAddress — LOCAL copy of the same underlying logic
 * already re-implemented independently in every other action file (see
 * ticketingExtractEmailAddress/transportExtractEmailAddress for the
 * identical shape). Extracts the bare, trimmed, lowercased email address
 * from a Gmail "From" header value, or from a bare address with no display
 * name. Pure, no GAS globals. Never throws: a null/undefined/empty input
 * returns ''.
 */
function mojemenickaExtractEmailAddress(fromHeader) {
  if (!fromHeader) {
    return '';
  }

  const angleBracketMatch = /<([^>]*)>/.exec(fromHeader);
  const raw = angleBracketMatch ? angleBracketMatch[1] : fromHeader;

  return raw.trim().toLowerCase();
}

/**
 * isMojemenickaAllowedSender — D-01's sender check. FAIL-CLOSED: unlike
 * ICS_ACTION_CONFIG.importOnlyFrom's empty-means-everyone rule, an empty,
 * null, or undefined `allowedSenders` list here returns false for EVERY
 * sender — this action sends outbound mail to a third party, so an
 * unconfigured install must never reply to anyone. Comparison is exact
 * (not substring), case-insensitive on both sides via
 * mojemenickaExtractEmailAddress. Pure, no GAS globals.
 */
function isMojemenickaAllowedSender(fromHeader, allowedSenders) {
  if (!allowedSenders || allowedSenders.length === 0) {
    return false;
  }

  const sender = mojemenickaExtractEmailAddress(fromHeader);

  for (let i = 0; i < allowedSenders.length; i++) {
    if (mojemenickaExtractEmailAddress(allowedSenders[i]) === sender) {
      return true;
    }
  }

  return false;
}

/**
 * subjectContainsAnyMojemenickaString — D-08's subject check (amends and
 * renames the retired single-string matcher from 04-01), also reused
 * as-is for D-09's weekly-range check (resolveMojemenickaRequestedRange
 * below) — ONE matcher deliberately serves both so their semantics can
 * never drift apart. Iterates `candidates`, returning true on the first
 * case-insensitive substring hit against `subject`. FAIL-CLOSED: a
 * null/undefined/non-array/empty `candidates`, or a null/undefined/empty
 * `subject`, returns false; individual null/empty/whitespace-only entries
 * in `candidates` are skipped rather than treated as matches (never a
 * match-everything wildcard). Pure, no GAS globals.
 */
function subjectContainsAnyMojemenickaString(subject, candidates) {
  if (!Array.isArray(candidates) || candidates.length === 0) {
    return false;
  }
  if (!subject) {
    return false;
  }

  const lowerSubject = subject.toLowerCase();

  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    if (!candidate || candidate.trim() === '') {
      continue;
    }
    if (lowerSubject.indexOf(candidate.toLowerCase()) !== -1) {
      return true;
    }
  }

  return false;
}

/**
 * resolveMojemenickaRequestedRange — D-09's today/week range selector.
 * Returns the literal string 'week' when `subject` contains any of
 * `weeklyTriggerStrings` (via subjectContainsAnyMojemenickaString above),
 * otherwise the literal string 'today'. Pure, no GAS globals.
 *
 * SUBORDINATION: this function is only ever called (see run() below) on a
 * subject that has ALREADY matched config.triggerStrings via
 * resolveMojemenickaRequestMessages — it is a range selector for an
 * already-applying request, never a second gate on whether the action
 * applies at all.
 */
function resolveMojemenickaRequestedRange(subject, weeklyTriggerStrings) {
  return subjectContainsAnyMojemenickaString(subject, weeklyTriggerStrings) ? 'week' : 'today';
}

/**
 * assertMojemenickaMenuUrlIsFetchable — D-10's SINGLE source of truth for
 * the menu-URL misconfiguration rule, called both by run() (before the
 * try/catch, so a misconfiguration reaches the owner per D-05) and by
 * buildMojemenickaMenuRequestUrl below. Throws an Error naming
 * '09-action-mojemenicka-MENU_URL' when `menuUrl` is null/undefined or
 * trims to empty, and separately when it already carries a query string
 * (`indexOf('?') !== -1`) — the action always builds its own clean query
 * string and must never merge into or shadow an existing one. Pure, no GAS
 * globals.
 */
function assertMojemenickaMenuUrlIsFetchable(menuUrl) {
  if (!menuUrl || menuUrl.trim() === '') {
    throw new Error('MojeMenicka: 09-action-mojemenicka-MENU_URL is not configured');
  }
  if (menuUrl.indexOf('?') !== -1) {
    throw new Error('MojeMenicka: 09-action-mojemenicka-MENU_URL must be a bare base URL with no query string of its own');
  }
}

/**
 * buildMojemenickaMenuRequestUrl — D-10. Asserts `menuUrl` is fetchable
 * (see assertMojemenickaMenuUrlIsFetchable above), then returns
 * `menuUrl + '?format=html&range=' + range` — e.g.
 * 'https://example.com/menu?format=html&range=today' or
 * '...&range=week'. Pure, no GAS globals.
 */
function buildMojemenickaMenuRequestUrl(menuUrl, range) {
  assertMojemenickaMenuUrlIsFetchable(menuUrl);
  return menuUrl + '?format=html&range=' + range;
}

/**
 * resolveMojemenickaRequestMessages — the single pure resolver both
 * appliesTo and run call, so the two can never disagree about which
 * messages match (D-08's two-condition trigger, AND of sender + subject).
 * `messages` is any array of objects exposing getFrom()/getSubject() (a
 * real GmailMessage[] in GAS, a fake array of plain objects under Node).
 * `triggerStrings` is now a LIST (D-08, amending 04-01's single string),
 * forwarded to subjectContainsAnyMojemenickaString. Preserves input order.
 * Pure w.r.t. its inputs.
 */
function resolveMojemenickaRequestMessages(messages, allowedSenders, triggerStrings) {
  return (messages || []).filter(function (message) {
    return (
      isMojemenickaAllowedSender(message.getFrom(), allowedSenders) &&
      subjectContainsAnyMojemenickaString(message.getSubject(), triggerStrings)
    );
  });
}

/**
 * fetchMojemenickaMenuHtml — GAS-only (UrlFetchApp), fakeable under Node
 * via a fake global.UrlFetchApp (see the "No Analog Found / Gaps" table in
 * 04-PATTERNS.md — no prior action in this codebase calls UrlFetchApp).
 * D-10: builds its own request URL via buildMojemenickaMenuRequestUrl
 * (menuUrl + range), rather than fetching `menuUrl` directly.
 * `muteHttpExceptions: true` is LOAD-BEARING, not a style choice: an
 * unmuted non-2xx response throws inside the GAS call itself, before
 * Task 2's fallback body can be produced — that would defeat D-04. Treats
 * a response code outside 200-299, or a body that is empty or
 * whitespace-only, as a failure by throwing an Error naming the response
 * code. Returns the response body verbatim (D-02) on success.
 */
function fetchMojemenickaMenuHtml(menuUrl, range) {
  const url = buildMojemenickaMenuRequestUrl(menuUrl, range);
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('MojeMenicka: menu fetch failed with HTTP status ' + code);
  }
  if (!body || body.trim() === '') {
    throw new Error('MojeMenicka: menu fetch returned an empty body (HTTP ' + code + ')');
  }

  return body;
}

/**
 * buildMojemenickaFallbackHtml — D-04's graceful-fallback reply body. Pure,
 * zero-argument, returns a FIXED brief HTML paragraph telling the requester
 * the menu could not be retrieved right now and to try again later — the
 * SAME string on every call (not formatted per-error).
 *
 * Deliberately carries NO internal detail — no URL, no HTTP status, no
 * error text, no stack. This is a deliberate mitigation of threat T-04-02
 * (information disclosure): the recipient here is an external third party,
 * not the owner, and the owner already has the full detail in the
 * Stackdriver log (see run()'s console.log call below).
 */
function buildMojemenickaFallbackHtml() {
  return '<p>Sorry, the menu could not be retrieved right now. Please try again later.</p>';
}

/**
 * MOJEMENICKA_ACTION — the mojemenicka action descriptor, built on the
 * same shape as every other action (see TRANSPORT_TICKETS_ACTION,
 * src/08-action-transport-tickets.js).
 */
const MOJEMENICKA_ACTION = {
  name: 'mojemenicka',

  // GETTER, not a plain literal property — see this file's class-level
  // JSDoc and the sibling config file's own class-level JSDoc. Not
  // evaluated at object-construction time, only when something reads
  // `.config`, which happens lazily inside function bodies (dispatchActions,
  // notifyOwnerOfFailure) long after every project file has loaded.
  get config() {
    return MOJEMENICKA_ACTION_CONFIG;
  },

  /**
   * appliesTo — returns a LITERAL boolean (dispatchActions only skips on a
   * strict `=== false`). True when resolveMojemenickaRequestMessages finds
   * at least one matching message on the thread. Reads
   * config.triggerStrings (D-08, a list — resolveMojemenickaRequestMessages
   * forwards it to the OR matcher).
   */
  appliesTo: function (thread) {
    return (
      resolveMojemenickaRequestMessages(thread.getMessages(), MOJEMENICKA_ACTION.config.allowedSenders, MOJEMENICKA_ACTION.config.triggerStrings)
        .length > 0
    );
  },

  /**
   * run — resolves the matching messages via the same pure resolver
   * appliesTo uses, takes the LAST one (exactly one reply per thread per
   * run — replying per matching message would multiply replies on a long
   * thread), resolves the today/week range from THAT message's subject
   * (D-09 — the exact message that will receive the reply, not the thread
   * or the first message), fetches the menu for that range, and replies
   * in-thread with the verbatim HTML (D-02, D-03, D-10).
   *
   * The menu-URL misconfiguration guard below (assertMojemenickaMenuUrlIs-
   * Fetchable) stays OUTSIDE the try/catch and keeps propagating: an unset
   * or query-string-bearing URL is a MISCONFIGURATION, not one of D-04's
   * three enumerated fetch failures (network error, non-2xx, empty body) —
   * the owner must be told about it, so it deliberately throws and reaches
   * dispatchActions' catch / notifyOnFailure (D-05).
   *
   * The fetch itself is wrapped in a try/catch (D-04): on success the
   * reply carries the fetched HTML unchanged; on ANY fetch failure
   * (thrown error, non-2xx status, empty/whitespace body — all three
   * raised as a thrown Error by fetchMojemenickaMenuHtml), the underlying
   * reason is logged via console.log (prefixed "MojeMenicka:", matching
   * this codebase's action-level logging convention — only the owner can
   * read Stackdriver logs) and the reply body falls back to
   * buildMojemenickaFallbackHtml(). Critically, this catch does NOT
   * rethrow: per D-05, dispatchActions' catch/notifyOnFailure gate stays
   * reserved for genuinely unexpected errors, and a menu-fetch failure is
   * by definition not one — the deliberate consequence is that a menu-URL
   * outage produces replies-with-apology plus log lines and NO owner
   * email, while the misconfiguration guard above still throws and still
   * notifies. A failure in message.reply() itself is NOT caught here (it
   * happens outside this try block on the success path, and fallback
   * replies that themselves fail are equally genuine, unexpected
   * failures) — it propagates to dispatchActions' catch exactly like any
   * other unexpected error.
   */
  run: function (thread) {
    const messages = resolveMojemenickaRequestMessages(
      thread.getMessages(),
      MOJEMENICKA_ACTION.config.allowedSenders,
      MOJEMENICKA_ACTION.config.triggerStrings
    );
    const message = messages[messages.length - 1];
    const range = resolveMojemenickaRequestedRange(message.getSubject(), MOJEMENICKA_ACTION.config.weeklyTriggerStrings);
    const menuUrl = MOJEMENICKA_ACTION.config.menuUrl;

    assertMojemenickaMenuUrlIsFetchable(menuUrl);

    let html;
    try {
      html = fetchMojemenickaMenuHtml(menuUrl, range);
    } catch (fetchError) {
      console.log('MojeMenicka: menu fetch failed, replying with fallback body: ' + fetchError);
      html = buildMojemenickaFallbackHtml();
    }

    message.reply('', { htmlBody: html });
  },
};

// Node/GAS environment bridge for MOJEMENICKA_ACTION_CONFIG (defined in the
// sibling src/09-action-cfg-mojemenicka.js — see the 260724-lqi config-
// split refactor for the full load-order/getter rationale this mirrors).
// Under GAS's shared global scope this is ALREADY visible here by bare
// name — no action needed, and this `if` block never executes there.
// Under Node, each `require()`d file is its own isolated module with its
// own scope, so the bare reference inside this file's descriptor above
// would otherwise throw ReferenceError. Same `globalThis` bridge technique
// (not a redeclared `const`/`let`/`var`, which would collide under GAS's
// concatenated scope) already established by every other action file's own
// equivalent bridge (see src/08-action-transport-tickets.js line 735).
if (typeof module !== 'undefined' && module.exports) {
  globalThis.MOJEMENICKA_ACTION_CONFIG = require('./09-action-cfg-mojemenicka.js').MOJEMENICKA_ACTION_CONFIG;
}

// GAS-safe Node export: `typeof module` is safely "undefined" in the Apps
// Script runtime, so this line is inert there and only active under Node.
// Exports every pure helper (mojemenickaExtractEmailAddress,
// isMojemenickaAllowedSender, subjectContainsAnyMojemenickaString (D-08/
// D-09's shared OR matcher, replacing the retired single-string matcher),
// resolveMojemenickaRequestedRange (D-09), assertMojemenickaMenuUrlIs-
// Fetchable / buildMojemenickaMenuRequestUrl (D-10),
// resolveMojemenickaRequestMessages, buildMojemenickaFallbackHtml),
// fetchMojemenickaMenuHtml (GAS-only/UrlFetchApp, but exported so D-04's
// three fetch-failure modes can be proven under Node with a fake
// global.UrlFetchApp — same precedent set when cancelTransportTicketEvent
// was exported for fake-global.Calendar testing in quick-260813-dq2), and
// MOJEMENICKA_ACTION (action registry).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    mojemenickaExtractEmailAddress: mojemenickaExtractEmailAddress,
    isMojemenickaAllowedSender: isMojemenickaAllowedSender,
    subjectContainsAnyMojemenickaString: subjectContainsAnyMojemenickaString,
    resolveMojemenickaRequestedRange: resolveMojemenickaRequestedRange,
    assertMojemenickaMenuUrlIsFetchable: assertMojemenickaMenuUrlIsFetchable,
    buildMojemenickaMenuRequestUrl: buildMojemenickaMenuRequestUrl,
    resolveMojemenickaRequestMessages: resolveMojemenickaRequestMessages,
    fetchMojemenickaMenuHtml: fetchMojemenickaMenuHtml,
    buildMojemenickaFallbackHtml: buildMojemenickaFallbackHtml,
    MOJEMENICKA_ACTION: MOJEMENICKA_ACTION,
  };
}

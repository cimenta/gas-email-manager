'use strict';

// --- MEETINGS_ACTION test suite (quick-260820-g4r) --------------------------
//
// THIS ACTION IS BODY-SOURCED and scoped to messages carrying NO `.ics`
// attachment (D-04): a meeting invitation whose date/time/location live in
// the plain-text BODY, never a structured calendar attachment.
// ICS_CALENDAR_ACTION already owns every `.ics`-carrying message.
//
// UNLIKE EVERY PRIOR ACTION IN THIS CODEBASE, detection here is by SENDER
// DOMAIN pattern (with a `*.` subdomain-wildcard), not by an exact sender
// address (D-02) -- see meetingsDomainMatchesPattern's own tests below for
// the dot-boundary/suffix-anchor semantics this requires.
//
// Teamio (`*.teamio.com`) is the first supported system. REAL_TEAMIO_CS_BODY_TEXT
// below is modeled on a real owner-supplied Teamio interview-invitation email
// (Czech): the three label lines (`Kdy:` / `Čas:` / `Kde:`), the leading
// weekday-plus-comma, the genitive month name, and the `délka 30 minut`
// duration phrase are all verbatim from that real body -- only the meeting
// URL and the confirmation-link URL have been substituted with fictional
// example.test domains (this project's established fixture convention; see
// this file's own domains below, never the real recruit.teamio.com link).
// EQUIVALENT_EN_BODY_TEXT is the LOCKED-BUT-UNOBSERVED English counterpart
// (D-05): required by decision, but not yet seen in a real email -- see
// src/10-action-meetings.js's own MEETINGS_EN_LANGUAGE_PACK VERIFICATION
// STATUS note for the same treatment CZECH_MONTH_GENITIVE_TO_INDEX already
// carries in src/06-lang-cs.js for its own unobserved genitive forms.
//
// Requiring ../src/10-action-meetings.js and ../src/10-action-cfg-meetings.js
// throws at require() time under Node while those files do not exist yet --
// THAT throw is this task's intended RED signal, not a broken test file. Do
// not "fix" it by stubbing either file; Task 2 creates them for real.

const test = require('node:test');
const assert = require('node:assert/strict');

const {
  meetingsExtractEmailAddress,
  meetingsExtractSenderDomain,
  meetingsDomainMatchesPattern,
  resolveMeetingSystem,
  resolveMeetingsCalendarId,
  meetingsAddMinutesToWallClockComponents,
  meetingsFormatWallClockComponentsIso,
  meetingsHarvestBodyLinks,
  meetingsExtractHtmlLabelValue,
  buildMeetingIdentifier,
  MEETINGS_CS_MONTH_GENITIVE_TO_INDEX,
  parseTeamioMeetingText,
  MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN,
  meetingsIsIcsAttachment,
  meetingsMessageHasIcsAttachment,
  findMeetingProcessingJobs,
  MEETINGS_ACTION,
  MEETINGS_FALLBACK_SUMMARY,
  meetingsSelectLanguagePack,
  teamioTextLooksLikeMeetingInvitation,
  MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN,
  meetingsExtractSenderDisplayName,
  meetingsFormatSenderAttribution,
  meetingsApplySenderAttribution,
} = require('../src/10-action-meetings.js');
const { MEETINGS_ACTION_CONFIG } = require('../src/10-action-cfg-meetings.js');
const ticketingModule = require('../src/07-action-ticketing-portals.js');

// --- fixtures ----------------------------------------------------------------
//
// ROUND 2 CORRECTION (live-test-driven, quick-260820-g4r): round 1's fixture
// was built from an oversimplified paraphrase, not the real .eml -- the
// owner's live pass against the actual Teamio interview-invitation email
// found two real bugs it never exercised. Re-derived here directly from the
// real .eml's own MIME parts (quoted-printable, multipart/alternative), with
// only the long opaque tracking tokens and the real Teams meeting ID
// fictionalized -- every label word, the date string ("24. srpna 2026"), the
// leading weekday ("Pondělí,"), the duration phrase ("délka 30 minut"), and
// the STRUCTURE of both bugs below are verbatim/faithful to the real source:
//
//   BUG 1 (location): the real .eml's text/plain part's "Kde:" line is NOT
//   the clean, human-readable Teams URL -- Teamio's ESP rewrites every
//   URL-shaped string into a track.teamio.com click-tracking REDIRECT when
//   generating the plain-text alternative from the HTML. The literal
//   `https://teams.microsoft.com/meet/...` URL only exists in the
//   text/html part's own "Kde:" table cell, as plain (non-hyperlinked) text
//   inside a <strong> tag. REAL_TEAMIO_CS_BODY_TEXT's own "Kde:" line below
//   is therefore a (fictionalized) track-style redirect, matching the real
//   plain-text part; REAL_TEAMIO_CS_HTML_TEXT below supplies the
//   (fictionalized) clean URL the way the real HTML part does.
//
//   BUG 2 (Links: coverage): the real .eml's plain-text part carries TWO
//   distinct URLs, not one -- the "Kde:" line's own redirect, AND a
//   separate "Přejít na potvrzení pohovoru: <url>" (confirm-attendance CTA)
//   line further down, on a DIFFERENT track.teamio.com redirect token. Once
//   the location is HTML-sourced (bug 1's fix), neither plain-text URL
//   equals the resolved location, so both now legitimately surface in the
//   description's Links: block -- see the dedicated round-2 tests below.

const REAL_TEAMIO_CS_BODY_TEXT =
  'Kdy: Pondělí, 24. srpna 2026\n' +
  'Čas: 13:30, délka 30 minut\n' +
  'Kde: https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345\n' +
  '\n' +
  'Přejít na potvrzení pohovoru: https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890\n';

// REAL_TEAMIO_CS_HTML_TEXT (round 2, NEW): the fictionalized counterpart of
// the real .eml's text/html MIME part -- the ONLY place the real, clean
// Teams meeting URL exists (as plain text inside a <strong> tag in the
// "Kde:" table cell, never hyperlinked). The confirm-CTA <a href> carries
// the SAME track.teamio.com-style redirect token as the plain-text part's
// own "Přejít na potvrzení pohovoru:" line -- matching the real .eml, where
// that particular link has no clean alternative anywhere.
const REAL_TEAMIO_CS_HTML_TEXT =
  '<html><body>' +
  '<table><tbody>' +
  '<tr><td>Kdy:</td><td><strong>Pondělí, 24. srpna 2026</strong></td></tr>' +
  '<tr><td>Čas:</td><td><strong>13:30, délka 30 minut</strong></td></tr>' +
  '<tr><td>Kde:</td><td><strong>https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL</strong></td></tr>' +
  '</tbody></table>' +
  '<a href="https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890">Přejít na potvrzení pohovoru</a>' +
  '</body></html>';

// EQUIVALENT_EN_BODY_TEXT (D-05): locked-but-unobserved English variant of
// the same shape -- same date/time/duration, English labels and month name.
// No real English .eml exists, so this stays a direct clean-URL fixture
// (no HTML/tracking-redirect complexity invented for an unverified pack).
const EQUIVALENT_EN_BODY_TEXT =
  'When: Monday, 24 August 2026\n' +
  'Time: 13:30, duration 30 minutes\n' +
  'Where: https://teams.example-meet.test/324015781582590?p=mGynz87a3H6pQs6ySL\n' +
  '\n' +
  'Go to interview confirmation: https://recruit.example-teamio.test/confirm/abc123\n';

// REAL_TEAMIO_CS_REJECTION_BODY_TEXT (debug teamio-non-invite-error, NEW):
// the text/plain part of a REAL owner-supplied Teamio email that is NOT an
// invitation -- a candidate-rejection notice sent from the very same
// `*.teamio.com` domain the shipped default config matches. Verbatim from
// the real .eml (`Výběrové řízení na pozici AI Implementation Specialist –
// LLM, RAG & Agents.eml`, 2026-08-24) except the recruiter's direct phone
// and email, which are fictionalized per this file's fixture convention.
//
// THE LOAD-BEARING PROPERTY of this fixture: it contains NONE of the six
// label words (`Kdy:`/`Čas:`/`Kde:`/`When:`/`Time:`/`Where:`) in either MIME
// part -- verified against the real .eml. It is a plain prose letter. This
// is what makes it a true member of the "matched sender, no invitation
// structure" class, which had ZERO coverage before this session: every prior
// findMeetingProcessingJobs test below holds an INVITATION body constant
// while varying only the sender and the attachments, so the gate's total
// blindness to body content was structurally invisible to the suite.
const REAL_TEAMIO_CS_REJECTION_BODY_TEXT =
  'Vážený pane Šimčíku,\n' +
  '\n' +
  'děkujeme Vám za zájem pracovat v naší společnosti na pozici AI Implementation Specialist – LLM, RAG & Agents.\n' +
  'Velmi si toho vážíme.\n' +
  '\n' +
  'Vaše zkušenosti a přístup nás zaujaly, nicméně jsme se rozhodli dát přednost kandidátovi, který lépe odpovídá specifickému profilu, který pro tuto roli aktuálně hledáme.\n' +
  '\n' +
  'Pevně věříme, že najdete příležitost, která bude odpovídat Vašim kvalitám a profesnímu zaměření, a přejeme Vám hodně úspěchů v dalším kariérním směřování.\n' +
  '\n' +
  'S pozdravem,\n' +
  '\n' +
  'JUDr. Simona Pažinková\n' +
  'HR Manager\n' +
  '+420 000 000 000\n' +
  'simona.pazinkova@example-prochazkapartners.test\n' +
  '\n' +
  'Prochazka & Partners\n' +
  'Václavské náměstí 841/3, 110 00 Praha 1\n';

// The real rejection .eml's From header shape: Teamio randomizes the LOCAL
// PART per recipient (the opaque `uaktdd3ygto` token), which is exactly why
// no narrower exact-address config could ever separate this email from a
// genuine invitation -- both arrive from the same domain with unpredictable
// local parts. Only body content can distinguish them.
const REAL_TEAMIO_REJECTION_FROM =
  'Simona Pažinková <simona.pazinkova.prochazkapartnerssro.uaktdd3ygto@recruit.teamio.com>';

// REAL_TEAMIO_CS_INVITATION_FROM (quick-260824-hva, NEW): the reported
// invitation's own From header -- display name `Denisa Čerevková`, address
// on `recruit.teamio.com`, matching REAL_TEAMIO_REJECTION_FROM's exact shape
// (display name, space, angle-bracketed address) and the SAME
// opaque-local-part fictionalization convention that constant already
// documents (a real per-recipient Teamio tracking token, fictionalized here).
const REAL_TEAMIO_CS_INVITATION_FROM =
  'Denisa Čerevková <denisa.cerevkova.somecompanysro.xk29fjq8a1@recruit.teamio.com>';

// The reported Czech subject line for the invitation this quick task fixes
// (quick-260824-hva) -- used only by the regression-proof test below.
const REAL_TEAMIO_CS_INVITATION_SUBJECT = 'Pozvánka na pohovor – AI Implementation Specialist';

function fakeMeetingAttachment(name, contentType) {
  return {
    getName: function () {
      return name;
    },
    getContentType: function () {
      return contentType || '';
    },
  };
}

function fakeMeetingMessage(fromHeader, bodyText, subject, attachments) {
  return {
    getFrom: function () {
      return fromHeader;
    },
    getPlainBody: function () {
      return bodyText;
    },
    getSubject: function () {
      return subject;
    },
    getAttachments: function () {
      return attachments || [];
    },
  };
}

// --- domain matching (D-02) --------------------------------------------------

test('meetingsDomainMatchesPattern: "*.teamio.com" matches "recruit.teamio.com" (subdomain wildcard)', () => {
  assert.equal(meetingsDomainMatchesPattern('recruit.teamio.com', '*.teamio.com'), true);
});

test('meetingsDomainMatchesPattern: "*.teamio.com" also matches the bare apex "teamio.com"', () => {
  assert.equal(meetingsDomainMatchesPattern('teamio.com', '*.teamio.com'), true);
});

test('meetingsDomainMatchesPattern: "*.teamio.com" does NOT match "notteamio.com" -- dot-boundary enforced', () => {
  assert.equal(meetingsDomainMatchesPattern('notteamio.com', '*.teamio.com'), false);
});

test('meetingsDomainMatchesPattern: "*.teamio.com" does NOT match "teamio.com.evil.net" -- suffix-anchored, not substring', () => {
  assert.equal(meetingsDomainMatchesPattern('teamio.com.evil.net', '*.teamio.com'), false);
});

test('meetingsDomainMatchesPattern: a pattern with no "*." prefix is exact equality only', () => {
  assert.equal(meetingsDomainMatchesPattern('recruit.teamio.com', 'teamio.com'), false);
  assert.equal(meetingsDomainMatchesPattern('teamio.com', 'teamio.com'), true);
});

test('meetingsDomainMatchesPattern: case-insensitive on both sides', () => {
  assert.equal(meetingsDomainMatchesPattern('Recruit.Teamio.COM', '*.TEAMIO.com'), true);
});

test('meetingsDomainMatchesPattern: never throws on null/empty input, returns false', () => {
  assert.equal(meetingsDomainMatchesPattern('', '*.teamio.com'), false);
  assert.equal(meetingsDomainMatchesPattern('teamio.com', ''), false);
  assert.equal(meetingsDomainMatchesPattern(null, null), false);
});

test('meetingsExtractSenderDomain: a display-name From header resolves the same domain as a bare address', () => {
  assert.equal(
    meetingsExtractSenderDomain('Denisa Čeřevková <denisa.x@recruit.teamio.com>'),
    meetingsExtractSenderDomain('denisa.x@recruit.teamio.com')
  );
  assert.equal(meetingsExtractSenderDomain('denisa.x@recruit.teamio.com'), 'recruit.teamio.com');
});

test('resolveMeetingSystem: returns the FIRST matching config entry, list order', () => {
  const systems = [
    { domainPattern: '*.teamio.com', calendarId: 'first' },
    { domainPattern: 'recruit.teamio.com', calendarId: 'second' },
  ];
  const result = resolveMeetingSystem('a@recruit.teamio.com', systems);
  assert.equal(result.calendarId, 'first');
});

test('resolveMeetingSystem: returns null for no match, a null list, and an empty list -- never throws', () => {
  assert.equal(resolveMeetingSystem('a@unrelated.example', [{ domainPattern: '*.teamio.com', calendarId: null }]), null);
  assert.equal(resolveMeetingSystem('a@recruit.teamio.com', null), null);
  assert.equal(resolveMeetingSystem('a@recruit.teamio.com', []), null);
});

test('resolveMeetingSystem: an entry whose domainPattern has no registered parser resolves to a system but produces no processing job', () => {
  const systems = [{ domainPattern: '*.unregistered-system.example', calendarId: null }];
  const system = resolveMeetingSystem('a@sub.unregistered-system.example', systems);
  assert.ok(system);

  const message = fakeMeetingMessage('a@sub.unregistered-system.example', REAL_TEAMIO_CS_BODY_TEXT, 'Interview', []);
  assert.equal(findMeetingProcessingJobs([message], systems).length, 0);
});

// --- Teamio parser, Czech (D-05, D-08, D-09) ---------------------------------

test('parseTeamioMeetingText: the Czech fixture with a subject and the HTML body parses to the exact expected components, location HTML-sourced (bug 1 fix)', () => {
  const result = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Interview with Denisa', REAL_TEAMIO_CS_HTML_TEXT);
  assert.deepEqual(
    {
      summary: result.summary,
      year: result.year,
      month: result.month,
      day: result.day,
      hour: result.hour,
      minute: result.minute,
      durationMinutes: result.durationMinutes,
      location: result.location,
    },
    {
      summary: 'Interview with Denisa',
      year: 2026,
      month: 7,
      day: 24,
      hour: 13,
      minute: 30,
      durationMinutes: 30,
      location: 'https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL',
    }
  );
});

test('parseTeamioMeetingText: BUG 1 REGRESSION -- without an HTML body, location falls back to the plain-text "Kde:" value, which for a real Teamio email is an opaque track.teamio.com click-tracking redirect, NOT a human-readable URL', () => {
  const result = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject');
  assert.equal(result.location, 'https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345');
  // documents exactly why bug 1 mattered: this fallback value is NOT the
  // recognizable Teams meeting URL a human (or the calendar event) wants.
  assert.notEqual(result.location, 'https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL');
});

test('MEETINGS_CS_MONTH_GENITIVE_TO_INDEX: month-table completeness -- every Czech genitive month name resolves to 0-11 in order', () => {
  const expectedOrder = [
    'ledna',
    'února',
    'března',
    'dubna',
    'května',
    'června',
    'července',
    'srpna',
    'září',
    'října',
    'listopadu',
    'prosince',
  ];
  expectedOrder.forEach(function (name, index) {
    assert.equal(MEETINGS_CS_MONTH_GENITIVE_TO_INDEX[name], index);
  });
});

test('parseTeamioMeetingText: leading weekday-plus-comma is skipped, and a variant with no weekday prefix parses identically', () => {
  const withWeekday = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject');
  const noWeekdayBody = REAL_TEAMIO_CS_BODY_TEXT.replace('Pondělí, 24. srpna 2026', '24. srpna 2026');
  const withoutWeekday = parseTeamioMeetingText(noWeekdayBody, 'Subject');

  assert.deepEqual(
    { year: withWeekday.year, month: withWeekday.month, day: withWeekday.day },
    { year: withoutWeekday.year, month: withoutWeekday.month, day: withoutWeekday.day }
  );
});

test('parseTeamioMeetingText: a body with no duration phrase yields durationMinutes null (fallback is the pipeline\'s job, not the parser\'s)', () => {
  const noDurationBody = REAL_TEAMIO_CS_BODY_TEXT.replace('Čas: 13:30, délka 30 minut', 'Čas: 13:30');
  const result = parseTeamioMeetingText(noDurationBody, 'Subject');
  assert.equal(result.durationMinutes, null);
});

test('parseTeamioMeetingText: a CRLF-joined variant of the fixture parses identically -- separator-agnostic', () => {
  const crlfBody = REAL_TEAMIO_CS_BODY_TEXT.replace(/\n/g, '\r\n');
  const lfResult = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject');
  const crlfResult = parseTeamioMeetingText(crlfBody, 'Subject');

  assert.deepEqual(
    { year: lfResult.year, month: lfResult.month, day: lfResult.day, hour: lfResult.hour, minute: lfResult.minute, durationMinutes: lfResult.durationMinutes },
    { year: crlfResult.year, month: crlfResult.month, day: crlfResult.day, hour: crlfResult.hour, minute: crlfResult.minute, durationMinutes: crlfResult.durationMinutes }
  );
});

test('parseTeamioMeetingText: an empty or missing subject yields the MEETINGS_FALLBACK_SUMMARY value rather than an empty summary or a throw', () => {
  assert.equal(parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, '').summary, MEETINGS_FALLBACK_SUMMARY);
  assert.equal(parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, undefined).summary, MEETINGS_FALLBACK_SUMMARY);
});

test('parseTeamioMeetingText: a missing "Kde:" line yields an empty location and still parses -- a meeting with no stated place is real, not a parse failure', () => {
  const noWhereBody = 'Kdy: Pondělí, 24. srpna 2026\nČas: 13:30, délka 30 minut\n';
  const result = parseTeamioMeetingText(noWhereBody, 'Subject');
  assert.equal(result.location, '');
  assert.equal(result.year, 2026);
});

test('parseTeamioMeetingText: without an HTML body, description reproduces the matched label lines verbatim AND carries a Links: block with the confirm-CTA URL but NOT the (plain-text-sourced) location URL', () => {
  const result = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject');

  assert.ok(result.description.indexOf('Kdy: Pondělí, 24. srpna 2026') !== -1);
  assert.ok(result.description.indexOf('Čas: 13:30, délka 30 minut') !== -1);
  assert.ok(
    result.description.indexOf('Kde: https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345') !== -1
  );
  assert.ok(result.description.indexOf('Links:') !== -1);
  assert.ok(
    result.description.indexOf('https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890') !== -1
  );

  // location (plain-text-sourced here, since no html arg was passed) equals
  // the "Kde:" line's own URL, so it is excluded from the Links: block --
  // only the confirm-CTA URL appears there (1 link).
  const linksBlock = result.description.slice(result.description.indexOf('Links:'));
  assert.equal(linksBlock.indexOf('https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345'), -1);
});

test('parseTeamioMeetingText: BUG 2 FIX -- with the HTML body provided, the Links: block carries BOTH real plain-text URLs (the "Kde:" line\'s own tracked redirect AND the confirm-CTA tracked redirect), since neither matches the now HTML-sourced location', () => {
  const result = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject', REAL_TEAMIO_CS_HTML_TEXT);

  // location is now the clean HTML-sourced Teams URL (bug 1's fix) -- see
  // this file's own "location HTML-sourced" test above.
  assert.equal(result.location, 'https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL');

  const linksBlock = result.description.slice(result.description.indexOf('Links:'));
  const linksLines = linksBlock
    .split('\n')
    .slice(1)
    .filter(function (line) {
      return line.trim() !== '';
    });

  assert.deepEqual(linksLines, [
    'https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345',
    'https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890',
  ]);
});

test('parseTeamioMeetingText: BUG 3 FIX -- with the HTML body provided, the description\'s echoed "Kde:" line uses the SAME resolved location value as the location field, not a second, independently-sourced plain-text copy', () => {
  const result = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject', REAL_TEAMIO_CS_HTML_TEXT);

  assert.equal(result.location, 'https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL');
  assert.ok(result.description.indexOf('Kde: https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL') !== -1);

  // The stale plain-text-sourced tracked URL must NOT appear as the
  // rendered "Kde:" label line -- it may still legitimately appear later,
  // as a Links: entry (round 2's fix, proven by the test above), but the
  // "Kde:" line itself must be consistent with the resolved location.
  const kdeLineStart = result.description.indexOf('Kde:');
  const kdeLineEnd = result.description.indexOf('\n', kdeLineStart);
  const kdeLine = result.description.slice(kdeLineStart, kdeLineEnd === -1 ? undefined : kdeLineEnd);
  assert.equal(kdeLine, 'Kde: https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL');
  assert.equal(kdeLine.indexOf('track.example-teamio.test'), -1);
});

test('parseTeamioMeetingText: BUG 4 FIX -- the Links: block\'s two entries are separated by a blank line', () => {
  const result = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject', REAL_TEAMIO_CS_HTML_TEXT);

  const expectedLinksBlock =
    'Links:\n' +
    'https://track.example-teamio.test/f/a/IXJFmY0v_kdeToken~~/AAJwnRA~/kdeRedirectPath12345\n' +
    '\n' +
    'https://track.example-teamio.test/f/a/XO9VvLPu_ctaToken~~/AAJwnRA~/ctaRedirectPath67890';

  const linksBlock = result.description.slice(result.description.indexOf('Links:'));
  assert.equal(linksBlock, expectedLinksBlock);
});

test('meetingsHarvestBodyLinks: dedupes repeated URLs, preserves source order, and caps at 10', () => {
  const manyExtra = Array.from({ length: 12 }, function (_, i) {
    return 'https://example.test/link' + i;
  }).join(' ');
  const bodyWithDupes = 'https://a.test/one https://a.test/one https://a.test/two ' + manyExtra;

  const links = meetingsHarvestBodyLinks(bodyWithDupes, null);
  assert.equal(links.length, 10);
  assert.deepEqual(links.slice(0, 3), ['https://a.test/one', 'https://a.test/two', 'https://example.test/link0']);
});

// --- meetingsExtractHtmlLabelValue (round 2, bug 1 fix) ----------------------

test('meetingsExtractHtmlLabelValue: extracts the <strong>-wrapped value from the real Teamio HTML table structure for a given label', () => {
  assert.equal(
    meetingsExtractHtmlLabelValue(REAL_TEAMIO_CS_HTML_TEXT, ['Kde']),
    'https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL'
  );
  assert.equal(meetingsExtractHtmlLabelValue(REAL_TEAMIO_CS_HTML_TEXT, ['Kdy']), 'Pondělí, 24. srpna 2026');
});

test('meetingsExtractHtmlLabelValue: returns null when the label is not present, or when htmlBodyText is empty/undefined -- never throws', () => {
  assert.equal(meetingsExtractHtmlLabelValue(REAL_TEAMIO_CS_HTML_TEXT, ['Where']), null);
  assert.equal(meetingsExtractHtmlLabelValue('', ['Kde']), null);
  assert.equal(meetingsExtractHtmlLabelValue(undefined, ['Kde']), null);
});

// --- Teamio parser, English (D-05) -------------------------------------------

test('parseTeamioMeetingText: the English fixture parses to the same date/time components as the Czech one', () => {
  const csResult = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, 'Subject');
  const enResult = parseTeamioMeetingText(EQUIVALENT_EN_BODY_TEXT, 'Subject');

  assert.deepEqual(
    { year: enResult.year, month: enResult.month, day: enResult.day, hour: enResult.hour, minute: enResult.minute, durationMinutes: enResult.durationMinutes },
    { year: csResult.year, month: csResult.month, day: csResult.day, hour: csResult.hour, minute: csResult.minute, durationMinutes: csResult.durationMinutes }
  );
});

// --- parser failure modes -----------------------------------------------------

test('parseTeamioMeetingText: a body matching no registered language pack throws, message carries the full raw text', () => {
  const bogusBody = 'Some unrelated meeting text with no recognizable labels at all.';
  assert.throws(
    function () {
      parseTeamioMeetingText(bogusBody, 'Subject');
    },
    function (err) {
      return err.message.indexOf(bogusBody) !== -1;
    }
  );
});

test('parseTeamioMeetingText: labels present but no recognizable date throws, message carries the full raw text', () => {
  const badDateBody = 'Kdy: not-a-date-at-all\nČas: 13:30\n';
  assert.throws(
    function () {
      parseTeamioMeetingText(badDateBody, 'Subject');
    },
    function (err) {
      return err.message.indexOf(badDateBody) !== -1;
    }
  );
});

test('parseTeamioMeetingText: labels present but no recognizable time throws, message carries the full raw text', () => {
  const badTimeBody = 'Kdy: 24. srpna 2026\nČas: not-a-time\n';
  assert.throws(
    function () {
      parseTeamioMeetingText(badTimeBody, 'Subject');
    },
    function (err) {
      return err.message.indexOf(badTimeBody) !== -1;
    }
  );
});

test('parseTeamioMeetingText: an unrecognized month name names the bad value in the thrown message', () => {
  const badMonthBody = 'Kdy: 24. neexistujícího 2026\nČas: 13:30\n';
  assert.throws(function () {
    parseTeamioMeetingText(badMonthBody, 'Subject');
  }, /neexistujícího/);
});

test('parseTeamioMeetingText: hour out of range (25:30) throws "Hour out of range"', () => {
  const badHourBody = 'Kdy: 24. srpna 2026\nČas: 25:30\n';
  assert.throws(function () {
    parseTeamioMeetingText(badHourBody, 'Subject');
  }, /Hour out of range/);
});

test('parseTeamioMeetingText: minute out of range (13:75) throws "Minute out of range"', () => {
  const badMinuteBody = 'Kdy: 24. srpna 2026\nČas: 13:75\n';
  assert.throws(function () {
    parseTeamioMeetingText(badMinuteBody, 'Subject');
  }, /Minute out of range/);
});

// --- no-.ics gate and job resolution (D-04) ----------------------------------

test('findMeetingProcessingJobs: a matching-domain message with NO attachments yields exactly one processing job', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const message = fakeMeetingMessage('a@recruit.teamio.com', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', []);
  const jobs = findMeetingProcessingJobs([message], systems);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].message, message);
  assert.equal(jobs[0].system, systems[0]);
});

test('findMeetingProcessingJobs: the SAME message carrying an invite.ics attachment yields ZERO jobs (name rule)', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const message = fakeMeetingMessage('a@recruit.teamio.com', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', [
    fakeMeetingAttachment('invite.ics', 'application/octet-stream'),
  ]);
  assert.equal(findMeetingProcessingJobs([message], systems).length, 0);
});

test('findMeetingProcessingJobs: an attachment named x.dat with content type text/calendar also yields ZERO jobs (content-type rule)', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const message = fakeMeetingMessage('a@recruit.teamio.com', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', [
    fakeMeetingAttachment('x.dat', 'text/calendar'),
  ]);
  assert.equal(findMeetingProcessingJobs([message], systems).length, 0);
});

test('findMeetingProcessingJobs: a matching-domain message carrying an unrelated photo.jpg attachment still yields exactly one job -- the gate is .ics-specific, not attachment-phobic', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const message = fakeMeetingMessage('a@recruit.teamio.com', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', [
    fakeMeetingAttachment('photo.jpg', 'image/jpeg'),
  ]);
  assert.equal(findMeetingProcessingJobs([message], systems).length, 1);
});

test('findMeetingProcessingJobs: a non-matching-domain message yields zero jobs and never throws', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const message = fakeMeetingMessage('a@unrelated.example', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', []);
  assert.equal(findMeetingProcessingJobs([message], systems).length, 0);
});

test('MEETINGS_ACTION.appliesTo returns a LITERAL true/false (strict equality, not truthiness) across a matching thread, an .ics-carrying thread and a non-matching thread', () => {
  const matchingThread = {
    getMessages: function () {
      return [fakeMeetingMessage('a@recruit.teamio.com', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', [])];
    },
  };
  const icsThread = {
    getMessages: function () {
      return [fakeMeetingMessage('a@recruit.teamio.com', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', [fakeMeetingAttachment('invite.ics', '')])];
    },
  };
  const nonMatchingThread = {
    getMessages: function () {
      return [fakeMeetingMessage('a@unrelated.example', REAL_TEAMIO_CS_BODY_TEXT, 'Subject', [])];
    },
  };

  assert.strictEqual(MEETINGS_ACTION.appliesTo(matchingThread), true);
  assert.strictEqual(MEETINGS_ACTION.appliesTo(icsThread), false);
  assert.strictEqual(MEETINGS_ACTION.appliesTo(nonMatchingThread), false);
});

// --- dedup identifier (D-07) --------------------------------------------------

test('buildMeetingIdentifier: the same sender + subject + start components always produce the identical string', () => {
  const start = { year: 2026, month: 7, day: 24, hour: 13, minute: 30 };
  const id1 = buildMeetingIdentifier('a@recruit.teamio.com', 'Interview', start);
  const id2 = buildMeetingIdentifier('a@recruit.teamio.com', 'Interview', start);
  assert.deepEqual(id1, id2);
});

test('buildMeetingIdentifier: changing the sender, the subject, or any start component produces a different string', () => {
  const start = { year: 2026, month: 7, day: 24, hour: 13, minute: 30 };
  const base = buildMeetingIdentifier('a@recruit.teamio.com', 'Interview', start);

  assert.notEqual(buildMeetingIdentifier('b@recruit.teamio.com', 'Interview', start), base);
  assert.notEqual(buildMeetingIdentifier('a@recruit.teamio.com', 'Different Subject', start), base);
  assert.notEqual(buildMeetingIdentifier('a@recruit.teamio.com', 'Interview', Object.assign({}, start, { hour: 14 })), base);
});

test('buildMeetingIdentifier: subject case and surrounding/internal whitespace differences normalize to the SAME identifier', () => {
  const start = { year: 2026, month: 7, day: 24, hour: 13, minute: 30 };
  const id1 = buildMeetingIdentifier('a@recruit.teamio.com', 'Interview   With   Denisa', start);
  const id2 = buildMeetingIdentifier('a@recruit.teamio.com', '  interview with denisa  ', start);
  assert.equal(id1, id2);
});

test('buildMeetingIdentifier: a very long subject produces a bounded-length identifier, and the result contains no "=" character', () => {
  const start = { year: 2026, month: 7, day: 24, hour: 13, minute: 30 };
  const longSubject = 'A'.repeat(500) + '=suspicious=query=injection';
  const id = buildMeetingIdentifier('a@recruit.teamio.com', longSubject, start);

  assert.ok(id.length < 200);
  assert.equal(id.indexOf('='), -1);
});

// --- wiring, config and naming ------------------------------------------------

test('MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN["*.teamio.com"] is strictly equal to the exported parseTeamioMeetingText', () => {
  assert.strictEqual(MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN['*.teamio.com'], parseTeamioMeetingText);
});

test('MEETINGS_ACTION_CONFIG.meetingSystems ships the single default Teamio entry, and resolveMeetingsCalendarId falls back to the global default', () => {
  assert.deepEqual(MEETINGS_ACTION_CONFIG.meetingSystems, [{ domainPattern: '*.teamio.com', calendarId: null }]);

  const system = MEETINGS_ACTION_CONFIG.meetingSystems[0];
  assert.equal(resolveMeetingsCalendarId(system, 'default-calendar@group.calendar.google.com'), 'default-calendar@group.calendar.google.com');
});

test('MEETINGS_ACTION_CONFIG code defaults: defaultDurationMinutes 60, enabled true, notifyOnFailure true', () => {
  assert.equal(MEETINGS_ACTION_CONFIG.defaultDurationMinutes, 60);
  assert.equal(MEETINGS_ACTION_CONFIG.enabled, true);
  assert.equal(MEETINGS_ACTION_CONFIG.notifyOnFailure, true);
});

test('MEETINGS_ACTION.config is the SAME object reference as MEETINGS_ACTION_CONFIG (getter identity, not a copy)', () => {
  assert.equal(MEETINGS_ACTION.config, MEETINGS_ACTION_CONFIG);
});

test('D-10 collision guard: the meetings module exports no bare addMinutesToWallClockComponents/formatWallClockComponentsIso key, and the ticketing module\'s own is NOT strictly equal to the meetings module\'s namespaced equivalent', () => {
  const meetingsModule = require('../src/10-action-meetings.js');

  assert.equal(Object.prototype.hasOwnProperty.call(meetingsModule, 'addMinutesToWallClockComponents'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(meetingsModule, 'formatWallClockComponentsIso'), false);
  assert.notStrictEqual(meetingsModule.meetingsAddMinutesToWallClockComponents, ticketingModule.addMinutesToWallClockComponents);
  assert.notStrictEqual(ticketingModule.addMinutesToWallClockComponents, meetingsModule.meetingsAddMinutesToWallClockComponents);
});

test('meetingsAddMinutesToWallClockComponents: rolls over correctly across an hour, a day, a month and a year boundary', () => {
  assert.deepEqual(meetingsAddMinutesToWallClockComponents({ year: 2026, month: 7, day: 24, hour: 23, minute: 45 }, 30), {
    year: 2026,
    month: 7,
    day: 25,
    hour: 0,
    minute: 15,
  });
  assert.deepEqual(meetingsAddMinutesToWallClockComponents({ year: 2026, month: 7, day: 31, hour: 23, minute: 45 }, 30), {
    year: 2026,
    month: 8,
    day: 1,
    hour: 0,
    minute: 15,
  });
  assert.deepEqual(meetingsAddMinutesToWallClockComponents({ year: 2026, month: 11, day: 31, hour: 23, minute: 45 }, 30), {
    year: 2027,
    month: 0,
    day: 1,
    hour: 0,
    minute: 15,
  });
});

test('meetingsFormatWallClockComponentsIso: emits a zero-padded YYYY-MM-DDTHH:MM:SS string with no trailing Z and no offset', () => {
  assert.equal(meetingsFormatWallClockComponentsIso({ year: 2026, month: 7, day: 24, hour: 13, minute: 30 }), '2026-08-24T13:30:00');
});

test('meetingsExtractEmailAddress: never throws on null/undefined/empty input, returns empty string', () => {
  assert.equal(meetingsExtractEmailAddress(null), '');
  assert.equal(meetingsExtractEmailAddress(undefined), '');
  assert.equal(meetingsExtractEmailAddress(''), '');
});

// --- INVITATION-STRUCTURE GATE (debug teamio-non-invite-error) ---------------
//
// THE BUG THIS SECTION LOCKS OUT: findMeetingProcessingJobs used to admit a
// message as a meeting-processing job on ENVELOPE EVIDENCE ALONE -- sender
// domain matches a configured system, a parser is registered for that
// domainPattern, and no `.ics` attachment is present. It never read the body.
// The system's ONLY invitation-shaped-content check lived DOWNSTREAM inside
// parseTeamioMeetingText's language-pack selection, whose sole way of saying
// "this is not an invitation" was to THROW. Because `*.teamio.com` is a
// multi-purpose ATS domain that also sends rejections and status updates, any
// such email was admitted past the gate and then necessarily threw, reaching
// the owner as an action-failure notification (real incident, 2026-08-24).
//
// The fix adds the missing predicate to the SAME resolver appliesTo and run
// both use, so a body with no invitation structure means the action does not
// apply AT ALL -- no job, no thread claimed, no throw.
//
// ORACLE TYPE: specified (an explicit job-count/appliesTo contract), not
// implicit crash-freedom -- asserting merely "run does not throw" would be
// satisfied by an action that still wrongly claims and labels the thread.

test('findMeetingProcessingJobs: THE REGRESSION -- the real Teamio REJECTION email (matching domain, no .ics, NO invitation structure) yields ZERO jobs', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const message = fakeMeetingMessage(
    REAL_TEAMIO_REJECTION_FROM,
    REAL_TEAMIO_CS_REJECTION_BODY_TEXT,
    'Výběrové řízení na pozici AI Implementation Specialist – LLM, RAG & Agents',
    []
  );

  assert.equal(findMeetingProcessingJobs([message], systems).length, 0);
});

test('MEETINGS_ACTION.appliesTo: returns LITERAL false for the real Teamio rejection thread -- the action must not claim a thread it cannot process', () => {
  const rejectionThread = {
    getMessages: function () {
      return [
        fakeMeetingMessage(REAL_TEAMIO_REJECTION_FROM, REAL_TEAMIO_CS_REJECTION_BODY_TEXT, 'Výběrové řízení', []),
      ];
    },
  };

  assert.equal(MEETINGS_ACTION.appliesTo(rejectionThread), false);
});

test('MEETINGS_ACTION.run: the real Teamio rejection thread is a silent no-op -- never throws (this exact throw was the reported failure notification)', () => {
  const rejectionThread = {
    getMessages: function () {
      return [
        fakeMeetingMessage(REAL_TEAMIO_REJECTION_FROM, REAL_TEAMIO_CS_REJECTION_BODY_TEXT, 'Výběrové řízení', []),
      ];
    },
  };

  assert.doesNotThrow(function () {
    MEETINGS_ACTION.run(rejectionThread);
  });
});

test('findMeetingProcessingJobs: a genuine INVITATION from the same domain still yields exactly one job -- the new gate must not break the happy path', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const invitation = fakeMeetingMessage(REAL_TEAMIO_REJECTION_FROM, REAL_TEAMIO_CS_BODY_TEXT, 'Interview', []);

  assert.equal(findMeetingProcessingJobs([invitation], systems).length, 1);
});

test('findMeetingProcessingJobs: mixed thread -- one invitation plus one rejection from the same sender yields exactly ONE job, for the invitation only', () => {
  const systems = [{ domainPattern: '*.teamio.com', calendarId: null }];
  const invitation = fakeMeetingMessage(REAL_TEAMIO_REJECTION_FROM, REAL_TEAMIO_CS_BODY_TEXT, 'Interview', []);
  const rejection = fakeMeetingMessage(REAL_TEAMIO_REJECTION_FROM, REAL_TEAMIO_CS_REJECTION_BODY_TEXT, 'Rejection', []);

  const jobs = findMeetingProcessingJobs([invitation, rejection], systems);

  assert.equal(jobs.length, 1);
  assert.equal(jobs[0].message, invitation);
});

test('findMeetingProcessingJobs: FAIL-CLOSED -- a system whose domainPattern has a parser but NO registered detector yields zero jobs', () => {
  const patternWithParserButNoDetector = '*.detectorless-fixture.test';
  MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN[patternWithParserButNoDetector] = parseTeamioMeetingText;

  try {
    const systems = [{ domainPattern: patternWithParserButNoDetector, calendarId: null }];
    const message = fakeMeetingMessage('a@sub.detectorless-fixture.test', REAL_TEAMIO_CS_BODY_TEXT, 'Interview', []);

    assert.equal(findMeetingProcessingJobs([message], systems).length, 0);
  } finally {
    delete MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN[patternWithParserButNoDetector];
  }
});

// --- the detector predicate itself -------------------------------------------

test('teamioTextLooksLikeMeetingInvitation: true for the real Czech invitation body and for the English equivalent', () => {
  assert.equal(teamioTextLooksLikeMeetingInvitation(REAL_TEAMIO_CS_BODY_TEXT), true);
  assert.equal(teamioTextLooksLikeMeetingInvitation(EQUIVALENT_EN_BODY_TEXT), true);
});

test('teamioTextLooksLikeMeetingInvitation: false for the real Teamio rejection body', () => {
  assert.equal(teamioTextLooksLikeMeetingInvitation(REAL_TEAMIO_CS_REJECTION_BODY_TEXT), false);
});

test('teamioTextLooksLikeMeetingInvitation: BOUNDARY -- requires BOTH a when-label and a time-label, exactly mirroring the parser pack-selection rule', () => {
  // when only
  assert.equal(teamioTextLooksLikeMeetingInvitation('Kdy: Pondělí, 24. srpna 2026\n'), false);
  // time only
  assert.equal(teamioTextLooksLikeMeetingInvitation('Čas: 13:30, délka 30 minut\n'), false);
  // where only -- a where-label alone is never sufficient
  assert.equal(teamioTextLooksLikeMeetingInvitation('Kde: https://example.test/meet\n'), false);
  // both -> true, with no where-label at all (a meeting with no stated place is still a meeting)
  assert.equal(teamioTextLooksLikeMeetingInvitation('Kdy: Pondělí, 24. srpna 2026\nČas: 13:30\n'), true);
  // cross-language pairing must NOT satisfy a single pack
  assert.equal(teamioTextLooksLikeMeetingInvitation('Kdy: Monday\nTime: 13:30\n'), false);
});

test('teamioTextLooksLikeMeetingInvitation: never throws on empty/null/undefined/non-string input, returns false', () => {
  assert.equal(teamioTextLooksLikeMeetingInvitation(''), false);
  assert.equal(teamioTextLooksLikeMeetingInvitation(null), false);
  assert.equal(teamioTextLooksLikeMeetingInvitation(undefined), false);
  assert.equal(teamioTextLooksLikeMeetingInvitation(12345), false);
});

test('MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN: keyed by the SAME domainPattern string as the parser registry, so every shipped parser has a detector', () => {
  const parserPatterns = Object.keys(MEETING_BODY_PARSERS_BY_DOMAIN_PATTERN);
  const detectorPatterns = Object.keys(MEETING_INVITATION_DETECTORS_BY_DOMAIN_PATTERN);

  assert.deepEqual(parserPatterns.sort(), detectorPatterns.sort());
});

// --- detector/parser fidelity: the property that makes the gate safe ---------

test('FIDELITY PROPERTY: the detector returns false for exactly the bodies on which the parser throws its no-pack-match error, and true otherwise', () => {
  const bodies = [
    REAL_TEAMIO_CS_BODY_TEXT,
    EQUIVALENT_EN_BODY_TEXT,
    REAL_TEAMIO_CS_REJECTION_BODY_TEXT,
    '',
    'Kdy: Pondělí, 24. srpna 2026\n',
    'Čas: 13:30\n',
    'nothing structured here at all',
    'Kdy: Pondělí, 24. srpna 2026\nČas: 13:30\n',
  ];

  bodies.forEach(function (body) {
    const detected = teamioTextLooksLikeMeetingInvitation(body);

    let threwNoPackMatch = false;
    try {
      parseTeamioMeetingText(body, 'Subject');
    } catch (err) {
      threwNoPackMatch = /no registered language pack matched/.test(err.message);
    }

    // detector false  <=>  parser raises the no-pack-match error
    assert.equal(
      detected,
      !threwNoPackMatch,
      'detector/parser disagreed on body: ' + JSON.stringify(body.slice(0, 60))
    );
  });
});

test('meetingsSelectLanguagePack: returns the matched pack plus its label lines, or null -- the single shared source of truth both the detector and the parser consult', () => {
  const cs = meetingsSelectLanguagePack(REAL_TEAMIO_CS_BODY_TEXT);
  assert.ok(cs);
  assert.equal(cs.whenMatch.value, 'Pondělí, 24. srpna 2026');
  assert.equal(cs.timeMatch.value, '13:30, délka 30 minut');
  assert.ok(cs.whereMatch);

  const en = meetingsSelectLanguagePack(EQUIVALENT_EN_BODY_TEXT);
  assert.ok(en);
  assert.equal(en.whenMatch.value, 'Monday, 24 August 2026');

  assert.equal(meetingsSelectLanguagePack(REAL_TEAMIO_CS_REJECTION_BODY_TEXT), null);
  assert.equal(meetingsSelectLanguagePack(''), null);
});

test('PARSER CONTRACT UNCHANGED: a structure-PRESENT but value-unparseable body still THROWS -- the gate must not silence genuine parse bugs', () => {
  // Labels present (so the detector admits it), month name is nonsense.
  const structuredButBroken = 'Kdy: Pondělí, 24. nonexistentmonth 2026\nČas: 13:30\n';

  assert.equal(teamioTextLooksLikeMeetingInvitation(structuredButBroken), true);
  assert.throws(
    function () {
      parseTeamioMeetingText(structuredButBroken, 'Subject');
    },
    function (err) {
      return /Unrecognized meeting invitation date/.test(err.message);
    }
  );
});

// --- sender attribution (quick-260824-hva) ------------------------------------
//
// THE BUG THIS SECTION LOCKS OUT: every calendar event MEETINGS_ACTION creates
// carried NOTHING identifying who sent it -- parseTeamioMeetingText assembles
// `description` from the matched Kdy:/Čas:/Kde: label lines plus a harvested
// Links: block and NOTHING else, and it is never even passed the From header.
// processMeetingFromMessageBody DOES read message.getFrom(), but only feeds it
// to buildMeetingIdentifier for the invisible dedup tag. Real incident: the
// owner forwarded a real Teamio interview invitation (v0.15.1) whose generated
// event carried no sender info at all.
//
// THE FIX IS PIPELINE-LEVEL, NOT A TEAMIO PATCH (D-01): three pure helpers
// applied at processMeetingFromMessageBody's single choke point, between the
// parser call and createMeetingCalendarEvent -- so the guarantee holds for
// every registered meeting system, present and future, with zero per-parser
// change. parseTeamioMeetingText's own signature and body are untouched.

test('meetingsExtractSenderDisplayName: a quoted display name returns the name with the surrounding double quotes stripped and trimmed', () => {
  assert.equal(meetingsExtractSenderDisplayName('"Denisa Čerevková" <denisa@recruit.teamio.com>'), 'Denisa Čerevková');
});

test('meetingsExtractSenderDisplayName: an unquoted display name returns the trimmed name', () => {
  assert.equal(meetingsExtractSenderDisplayName('Denisa Čerevková <denisa@recruit.teamio.com>'), 'Denisa Čerevková');
});

test('meetingsExtractSenderDisplayName: a bare address with no display name, and an angle-bracket-only header, both return the empty string', () => {
  assert.equal(meetingsExtractSenderDisplayName('denisa@recruit.teamio.com'), '');
  assert.equal(meetingsExtractSenderDisplayName('<denisa@recruit.teamio.com>'), '');
});

test('meetingsExtractSenderDisplayName: null, undefined and the empty string all return the empty string and never throw', () => {
  assert.equal(meetingsExtractSenderDisplayName(null), '');
  assert.equal(meetingsExtractSenderDisplayName(undefined), '');
  assert.equal(meetingsExtractSenderDisplayName(''), '');
});

test('meetingsExtractSenderDisplayName SANITIZATION (T-hva-02): CR/LF/tab collapse to a single space and every "<"/">" is removed from a hostile display name attempting to forge a second label line and smuggle an angle-bracket tag', () => {
  const hostileFrom = '"Evil\r\nKde:\thttp://fake.test>Attacker" <real@recruit.teamio.com>';
  const result = meetingsExtractSenderDisplayName(hostileFrom);

  assert.equal(result.indexOf('\n'), -1);
  assert.equal(result.indexOf('\r'), -1);
  assert.equal(result.indexOf('<'), -1);
  assert.equal(result.indexOf('>'), -1);
  assert.equal(result, 'Evil Kde: http://fake.testAttacker');
});

test('meetingsFormatSenderAttribution: the real invitation header renders as one line, starting with "From:", containing BOTH the display name and the lowercased address in angle brackets', () => {
  const result = meetingsFormatSenderAttribution(REAL_TEAMIO_CS_INVITATION_FROM);

  assert.equal(result, 'From: Denisa Čerevková <denisa.cerevkova.somecompanysro.xk29fjq8a1@recruit.teamio.com>');
  assert.equal(result.indexOf('\n'), -1);
});

test('meetingsFormatSenderAttribution: a bare-address header (no display name) renders the heading plus the address with NO angle brackets', () => {
  assert.equal(meetingsFormatSenderAttribution('denisa@recruit.teamio.com'), 'From: denisa@recruit.teamio.com');
});

test('meetingsFormatSenderAttribution: a header carrying a display name but no parseable address renders the heading plus the name only', () => {
  assert.equal(meetingsFormatSenderAttribution('Denisa Čerevková <>'), 'From: Denisa Čerevková');
});

test('meetingsFormatSenderAttribution: null/undefined/empty input returns the empty string', () => {
  assert.equal(meetingsFormatSenderAttribution(null), '');
  assert.equal(meetingsFormatSenderAttribution(undefined), '');
  assert.equal(meetingsFormatSenderAttribution(''), '');
});

test('meetingsApplySenderAttribution: given a parser-shaped object and the real invitation header, the description\'s FIRST line is the attribution, SECOND line is empty, original description follows verbatim from line 3', () => {
  const parsedMeeting = {
    summary: 'Interview with Denisa',
    description: 'Kdy: Pondělí, 24. srpna 2026\nČas: 13:30, délka 30 minut',
    location: 'https://teams.example-meet.test/meet/324015781582590',
    year: 2026,
    month: 7,
    day: 24,
    hour: 13,
    minute: 30,
    durationMinutes: 30,
  };

  const result = meetingsApplySenderAttribution(parsedMeeting, REAL_TEAMIO_CS_INVITATION_FROM);
  const lines = result.description.split('\n');

  assert.equal(lines[0], 'From: Denisa Čerevková <denisa.cerevkova.somecompanysro.xk29fjq8a1@recruit.teamio.com>');
  assert.equal(lines[1], '');
  assert.equal(lines.slice(2).join('\n'), parsedMeeting.description);
});

test('meetingsApplySenderAttribution: returns a NEW object -- input is unmutated, and every other field survives the copy untouched', () => {
  const originalDescription = 'Kdy: Pondělí, 24. srpna 2026\nČas: 13:30, délka 30 minut';
  const parsedMeeting = {
    summary: 'Interview with Denisa',
    description: originalDescription,
    location: 'https://teams.example-meet.test/meet/324015781582590',
    year: 2026,
    month: 7,
    day: 24,
    hour: 13,
    minute: 30,
    durationMinutes: 30,
  };

  const result = meetingsApplySenderAttribution(parsedMeeting, REAL_TEAMIO_CS_INVITATION_FROM);

  assert.notEqual(result, parsedMeeting);
  assert.equal(parsedMeeting.description, originalDescription);
  assert.equal(result.summary, parsedMeeting.summary);
  assert.equal(result.location, parsedMeeting.location);
  assert.equal(result.year, parsedMeeting.year);
  assert.equal(result.month, parsedMeeting.month);
  assert.equal(result.day, parsedMeeting.day);
  assert.equal(result.hour, parsedMeeting.hour);
  assert.equal(result.minute, parsedMeeting.minute);
  assert.equal(result.durationMinutes, parsedMeeting.durationMinutes);
});

test('meetingsApplySenderAttribution: an empty/missing From header leaves the description byte-for-byte identical and still returns a usable object', () => {
  const originalDescription = 'Kdy: Pondělí, 24. srpna 2026\nČas: 13:30, délka 30 minut';
  const parsedMeeting = { summary: 'Interview with Denisa', description: originalDescription, location: '', year: 2026, month: 7, day: 24, hour: 13, minute: 30, durationMinutes: 30 };

  [null, undefined, ''].forEach(function (emptyFrom) {
    const result = meetingsApplySenderAttribution(parsedMeeting, emptyFrom);
    assert.equal(result.description, originalDescription);
    assert.ok(result);
    assert.equal(result.summary, parsedMeeting.summary);
  });
});

test('THE REGRESSION PROOF (quick-260824-hva): the real Teamio interview invitation\'s description contains BOTH the display name and the address, and the pre-existing Kdy:/Čas:/Kde: lines and Links: block all survive intact', () => {
  const parsed = parseTeamioMeetingText(REAL_TEAMIO_CS_BODY_TEXT, REAL_TEAMIO_CS_INVITATION_SUBJECT, REAL_TEAMIO_CS_HTML_TEXT);
  const result = meetingsApplySenderAttribution(parsed, REAL_TEAMIO_CS_INVITATION_FROM);

  assert.ok(result.description.indexOf('Denisa Čerevková') !== -1);
  assert.ok(result.description.indexOf('denisa.cerevkova.somecompanysro.xk29fjq8a1@recruit.teamio.com') !== -1);

  // pre-existing content survives intact, verbatim
  assert.ok(result.description.indexOf('Kdy: Pondělí, 24. srpna 2026') !== -1);
  assert.ok(result.description.indexOf('Čas: 13:30, délka 30 minut') !== -1);
  assert.ok(result.description.indexOf('Kde: https://teams.example-meet.test/meet/324015781582590?p=mGynz87a3H6pQs6ySL') !== -1);
  assert.ok(result.description.indexOf('Links:') !== -1);

  // the attribution leads, per D-03 placement
  const lines = result.description.split('\n');
  assert.ok(lines[0].indexOf('From:') === 0);
  assert.equal(lines[1], '');
});

test('THE GENERALITY PROOF (D-01): a synthetic parsed-meeting object with nothing to do with Teamio gets the attribution applied identically -- the guarantee lives above the parser layer', () => {
  const futureSystemParsedMeeting = { summary: 'Some other system meeting', description: 'Body line' };

  const result = meetingsApplySenderAttribution(futureSystemParsedMeeting, REAL_TEAMIO_CS_INVITATION_FROM);
  const lines = result.description.split('\n');

  assert.equal(lines[0], 'From: Denisa Čerevková <denisa.cerevkova.somecompanysro.xk29fjq8a1@recruit.teamio.com>');
  assert.equal(lines[1], '');
  assert.equal(lines.slice(2).join('\n'), 'Body line');
});

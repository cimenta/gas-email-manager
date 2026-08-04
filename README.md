# GAS Email Manager

A Google Apps Script application that periodically scans a Gmail account for
unlabeled emails within a configurable time window, runs a pluggable set of
"actions" against each match, and labels the email once processed.

Four actions ship today:

- **ICS-to-Calendar import** — any sender, generic `.ics` calendar attachments
- **Booking.com reservation management** — booking.com
- **Ticketing-portal PDF import** — enigoo.cz, Kino Art (kinoart.cz)
- **Transport ticket import** — RegioJet (regiojet.cz), IDOS.cz (idos.svt.cz)

Every matching email reliably gets its action(s) applied exactly once — no
duplicate processing, no silent failures: if any action throws, the owner
gets a plain-text failure notification and the thread is labeled as failed
in addition to processed, so nothing is lost silently.

## Features

- **ICS-to-Calendar import** — detects `.ics` attachments (by filename or
  `text/calendar` content-type) from any sender by default, parses every
  `VEVENT` (single, multiple, all-day, and recurring), and imports each into
  a configured Google Calendar with idempotent dedup and full timezone
  resolution. See [ICS-to-Calendar import](#ics-to-calendar-import).
- **Booking.com reservation management** — on a booking confirmation email,
  adds a safety-net calendar event with the check-in/check-out times, room
  type, guest count, address, and the full "Reservation details" section
  from the email in the description — but only if Google's own native
  Gmail-to-Calendar detection hasn't already created one for that
  reservation. On a cancellation email, deletes the matching event. See
  [Booking.com matching](#bookingcom-matching) for how a match is found.
- **Ticketing-portal PDF import** — for a configured portal (identified by
  its confirmation email's sender address), extracts the event name,
  date/time, and venue from the attached PDF ticket via Google Drive's
  PDF-to-Docs OCR conversion (Apps Script has no native PDF text parser),
  creates one calendar event per purchase regardless of how many tickets it
  contains, and optionally attaches the original PDF to the event. See
  [Ticketing portals](#ticketing-portals).
- **Transport ticket import (train/bus)** — for a configured carrier
  (identified by its confirmation email's sender address), builds the
  calendar event either from the `.ics` invite the confirmation email
  already carries (RegioJet) or, for a carrier with no `.ics` at all, by
  parsing the plain-text email body directly (IDOS.cz) — and optionally
  archives + attaches the accompanying ticket PDF to the created event
  either way. See [Transport tickets](#transport-tickets).
- **Sender allow-list (opt-in) and hand-off (exclude-list)** — restrict
  either action to specific sender addresses; unlisted senders are silently
  skipped (no error, thread still marked processed). Empty by default,
  meaning any sender's email is processed by the ICS action; the booking.com
  action defaults to `noreply@booking.com` only. The ICS action also
  supports the inverse — an `excludeFrom` list (empty by default) that hands
  a sender off to a more specific action instead, e.g. so the transport
  ticket action can own a carrier's `.ics` invites without the generic ICS
  action also importing them as a second, competing event.
- **Per-action enable/disable toggle** — every action (and, for the
  booking.com action, its add and remove sub-behaviors independently) can be
  turned off without deleting code — a disabled action is skipped before its
  `appliesTo`/`run` are ever called.
- **Multi-calendar routing** — a single `CONFIG.calendarId` default, with each
  action able to override it, and the ICS action additionally able to route
  different senders to entirely different calendars. See
  [Configuration reference](#configuration-reference).
- **Settings survive redeploys** — every setting can be set live via
  `PropertiesService`, which `clasp push` never touches, instead of editing
  the checked-in code. See
  [Live settings override](#live-settings-override-script-properties).
- **Fail-closed parsing** — a malformed `.ics` is fully parsed before any
  calendar write; a parse error produces zero calendar writes.
- **Owner failure notifications** — if an action fails, a plain-text email
  describes the failing action, the affected thread, and the error.
- **No duplicate processing** — once a thread is labeled (success or
  failure), it's excluded from all future runs.
- **Pluggable action framework** — each action lives in its own file with
  its own config block kept at the top of the file (so settings are the
  first thing visible when opening it); adding an unrelated future action
  requires no changes to the dispatch/labeling core.

## ICS-to-Calendar import

Detects `.ics` calendar attachments (by filename or `text/calendar`
content-type) from any sender by default — unlike the other actions, no
specific sender needs to be registered ahead of time; any `.ics`-carrying
email qualifies unless explicitly excluded (see
[Configuration reference](#configuration-reference)). It:

1. Parses every `VEVENT` in the attachment (single, multiple, all-day, and
   recurring), resolving `TZID`-based wall-clock times through the `.ics`
   file's own embedded `VTIMEZONE` rules (standard RFC 5545 practice for
   Outlook/Exchange, Apple Calendar, Google Calendar, etc.) rather than
   assuming a fixed offset.
2. Imports via the Advanced Calendar Service's `Events.import`, which is
   idempotent by the `.ics` file's own `UID` (`iCalUID`). If you also accept
   the same invite natively in Gmail (clicking "Yes"), Google Calendar
   resolves both paths to the *same* event instead of creating two — correct
   regardless of which happens first. The `.ics`'s own `SEQUENCE` revision
   number is also parsed and carried through, so importing an invite Gmail's
   native detection has already processed doesn't get rejected as a stale
   update; a genuine conflict is recovered with one bounded
   re-fetch-and-retry, following the Calendar API's own documented recovery
   instructions.
3. Carries the organizer/attendee list into the event description as plain
   informational text (never as real Calendar guests — see
   [Security notes](#security-notes)), along with the meeting description
   and an online-meeting URL (e.g. a Microsoft Teams link). Nested
   sub-components (like a VALARM reminder) are correctly excluded from the
   event's own fields, and excessive blank lines common in real
   Outlook/Exchange invites are collapsed to a single blank-line separator.

**Hand-off to a more specific action:** a sender with a more specialized
action registered for it (e.g. the transport-tickets action) can be added to
`ICS_ACTION_CONFIG.excludeFrom` so this generic action stands down for that
sender instead of creating a second, competing event for the same `.ics`
invite. See [Transport tickets](#transport-tickets).

## Booking.com matching

Google's own Gmail-to-Calendar detection (`eventType: "fromGmail"`) auto-
creates a calendar event from a booking confirmation email before this
script ever sees it — but the Calendar API only ever exposes a generic,
useless placeholder for that event's `description` field, never the real
reservation details, and `description` isn't one of the fields Google allows
patching on that event type either. So matching can't rely on API-readable
text at all for that first encounter. Instead, this action:

1. Tries an **exact match** first, via a private `extendedProperties` tag
   (`confirmationNumber=...`) — instant and precise once an event has been
   tagged.
2. Falls back to a **fuzzy match** (hotel name found in the event's
   `summary`/`location`, combined with a check-in/check-out date-range
   overlap against the event's `start`/`end`) for an untagged event's first
   encounter — this is necessarily lower-precision than an exact match, but
   it's the only information the Calendar API actually exposes for a
   Google-native event.
3. Once a fuzzy match succeeds, the event is immediately tagged with the
   confirmation number, so every future lookup for that same booking hits
   the fast, exact path instead.

An event this script creates itself is tagged from birth and never needs the
fuzzy fallback. The residual risk of the fuzzy fallback is a false match
between two different bookings at similarly-named properties with
overlapping stay dates — narrow, but not impossible; there is currently no
dry-run mode, so the remove path deletes a matched event outright.

## Language packs

The booking.com action's email-text matching (labels like "Check-in",
"Confirmation Number", "Reservation details", and the confirmation/
cancellation subject phrases) is driven by per-language "packs" instead of
hardcoded strings, so it isn't tied to English-language booking.com emails.
**English (`src/06-lang-en.js`) and Czech (`src/06-lang-cs.js`) ship today.**

Each language lives in exactly one file, `src/06-lang-<code>.js`, which
registers into a shared registry; every registered language's labels are
unioned together at match time, so the action doesn't need to detect or
guess which language a given email is written in.

Most of a language pack is just label arrays — copy an existing pack,
translate the string values, keep every key name identical, register under
the new language code. **Two pieces can require actual code, not just
translated strings, if your language's booking.com emails differ
structurally from an existing pack:**
- **Date/time format** — a `parseDateLine` function reference on the pack.
  Building the Czech pack surfaced that its date lines are day-first with a
  genitive month name and no comma (`"11. září 2026"`). It also surfaced
  that booking.com's own templates aren't consistent even within one
  language: some Czech properties render the time window as two separate
  `od`/`do`-prefixed 24-hour clauses (`"(od 14:00)"` / `"(do 10:00)"`),
  others as a single `HH:MM–HH:MM` range on each line (`"(15:00–22:00)"`) —
  structurally the same shape English already uses. `parseCzechBookingDateLine`
  in `src/06-lang-cs.js` tries the range form first, falling back to the
  `od`/`do` form, and is a worked example for handling more than one
  real-world variant within a single language. If your language's format
  matches an existing pack's, you can just reference that pack's
  `parseDateLine` instead of writing a new one.
- **Hotel name position in the subject** — `confirmationHotelNameSeparators`
  and `cancellationHotelNameSeparators`, each an array of
  `{ separator, side: 'before' | 'after' }`. Czech needed this because the
  same separator (an en dash) puts the hotel name *before* it in a
  confirmation subject but *after* it in a cancellation subject — the
  opposite of English's fixed "after" assumption on both.

See `src/06-lang-cs.js`'s and `src/06-lang-en.js`'s header comments for the
exact required keys.

## Ticketing portals

Unlike the ICS and booking.com actions, a ticket-purchase confirmation
email's body is typically pure boilerplate ("thanks for your purchase, your
e-ticket is attached") — every real detail (event name, date/time, venue)
lives only inside the attached PDF. Since Apps Script has no native PDF text
parser, this action:

1. Uploads the PDF to a project-owned, auto-managed temp Drive folder
   (`GAS Email Manager - Temp`, created automatically if missing).
2. Converts it to a Google Doc via the Drive Advanced Service — this
   triggers Google's OCR pipeline, which works whether the source PDF has a
   real text layer or is a scanned image.
3. Reads the converted Doc's text, deletes the temp Doc, and parses the
   result with the matching portal's own parser.
4. Creates ONE calendar event for the whole purchase, no matter how many
   tickets/pages the PDF contains.
5. If that portal's `insertPdfIntoEvent` is on, moves the original PDF into
   your permanent configured Drive folder and attaches it to the event as a
   real Calendar attachment; otherwise deletes it. Nothing lingers in
   either Drive folder unless you asked for it.

**Duplicate protection:** each created event is tagged with a stable
identifier extracted from the ticket (e.g. its ticket number) via a private
`extendedProperties` tag, checked against existing calendar events before
creating a new one — the same tag-before-create pattern the booking.com
action already uses, so reprocessing the same email can never create a
second event.

**Per-portal parsers, one shared file.** A portal's PDF layout is
platform-specific code, not just configuration — same principle as a
language pack's `parseDateLine`. Unlike language packs, all portal parsers
live together in one file, `src/07-action-ticketing-portals.js` (an explicit
project choice, not a limitation) — adding a new portal means adding a
config entry plus a parser function in that same file.

**New OAuth scopes:** this action needs `https://www.googleapis.com/auth/drive`
(folder/file management — the broader scope, not `drive.file`, since it
needs to find folders you may have created yourself, not just ones the
script created) and `https://www.googleapis.com/auth/documents` (reading the
OCR-converted Doc's text). You'll be prompted to re-authorize the first time
you run anything after adding this action.

## Transport tickets

A train/bus carrier's confirmation email either already carries a `.ics`
invite with the trip's real data (RegioJet), or has no `.ics` at all and
carries every detail only in the plain-text body (IDOS.cz). Each registered
carrier declares which shape its email has via a `mode` field
(`'ics'` or `'body'`) — both modes share the exact same dedup / PDF-archive /
Calendar-write pipeline downstream; only how the event's data is extracted
differs.

**`mode: 'ics'`** (RegioJet):

1. Reuses the ICS action's own iCalendar parser to read the invite's `.ics`
   attachment: departure/arrival stops and times (resolved through the
   invite's own `VTIMEZONE` rules, same as the ICS action), route summary,
   and the invite's `UID`/`SEQUENCE`.

**`mode: 'body'`** (IDOS.cz):

1. Parses `message.getPlainBody()` directly with a carrier-specific text
   parser (no `.ics`, no PDF text extraction) — station names, departure/
   arrival date and time, seats, and the order code, all pattern-anchored
   (never dependent on line position, since Gmail's rendering of a message
   body isn't guaranteed to preserve line boundaries).

**Both modes then:**

2. If the carrier's `insertPdfIntoEvent` is on, copy the accompanying ticket
   PDF into your permanent configured Drive folder (the same one the
   ticketing-portals action uses) and attach it to the event as a real
   Calendar attachment.
3. Tag the created event with a stable identifier extracted from the ticket
   (its ticket/order number — purchase-scoped when the carrier sells
   multi-seat orders as one ticket, e.g. IDOS.cz's own order code, rather
   than a narrower per-seat number) via a private `extendedProperties` tag,
   checked against existing calendar events before creating a new one —
   same tag-before-create pattern the booking.com and ticketing-portals
   actions already use.

**Adding a carrier with no `mode` field:** an entry with no `mode` set
resolves to `'ics'` unless a body-mode parser has been registered for that
carrier's sender address in the source code — this is a back-compat
guarantee, not something to rely on for a new carrier. Always set `mode`
explicitly when registering one.

**Why a separate action instead of just using the ICS action for `'ics'`-mode
carriers?** The generic ICS action would already import a `.ics`-carrying
carrier's invite on its own — but without archiving or attaching the ticket
PDF, and without the ticket's own identifier as the dedup key. Registering an
`'ics'`-mode carrier here is only useful alongside adding that carrier's
sender address to the ICS action's `excludeFrom` list (see
[Configuration reference](#configuration-reference)), so the two actions
don't both create a competing event for the same email. This doesn't apply
to `'body'`-mode carriers like IDOS.cz — the generic ICS action has nothing
to import for them in the first place, since there's no `.ics` attachment at
all.

## Security notes

- **Attendees are never real Calendar guests.** Organizer/attendee names
  from an `.ics` are surfaced as plain text in the event description only.
  The script never adds them as actual Calendar participants and never
  sends a calendar invitation on your behalf — the `.ics` attachment is
  untrusted, attacker-controllable input, and doing so would let a
  malicious invite make your account contact or notify arbitrary third
  parties.
- **Sender allow-lists are a convenience filter, not an authentication
  boundary.** They check the Gmail "From" header as GmailApp reports it —
  there is no SPF/DKIM/DMARC verification in-script.
- **The booking.com action's remove path deletes a real calendar event with
  no dry-run mode.** Matching is exact when a confirmation-number tag is
  present, and fuzzy (hotel name + date overlap) otherwise — see
  [Booking.com matching](#bookingcom-matching). Review the matching logic
  and test against your own real emails before relying on it unattended.
- **The ticketing-portals action's Drive scope is broader than strictly
  needed for defense-in-depth (`drive`, not `drive.file`).** This is
  necessary so it can find a permanent attachment folder you created
  yourself, rather than only folders the script created — be aware this
  scope grants the script access to your Drive generally, not just the
  folders this feature uses.

## Project layout

Apps Script's editor lists files alphabetically with no manual reordering,
so files are numbered to keep related code grouped:

| File | Contents |
|------|----------|
| `src/01-setup.js` | Global runtime `CONFIG`, one-time `setup()` (labels + trigger) |
| `src/02-main.js` | `processEmails()` entry point, label helpers, failure notification |
| `src/03-action-management.js` | Action registry + per-action dispatch/isolation (incl. the enable/disable toggle) |
| `src/04-----------.js` | Empty — visual separator before the actions section |
| `src/05-action-cfg-ics-import.js` | `ICS_ACTION_CONFIG` — the ICS action's settings |
| `src/05-action-ics-import.js` | The ICS-to-Calendar action (parser, action descriptor) |
| `src/06-action-cfg-booking-com-management.js` | `BOOKING_ACTION_CONFIG` — the booking.com action's settings |
| `src/06-action-booking-com-management.js` | The booking.com management action (parsers, matching, action descriptor) |
| `src/06-lang-en.js` | English language pack for the booking.com action's email-text matching |
| `src/06-lang-cs.js` | Czech language pack (see [Language packs](#language-packs)) |
| `src/07-action-cfg-ticketing-portals.js` | `TICKETING_PORTALS_ACTION_CONFIG` — the ticketing-portals action's settings |
| `src/07-action-ticketing-portals.js` | The ticketing-portals action (Drive/OCR pipeline, every portal's parser, action descriptor) |
| `src/08-action-cfg-transport-tickets.js` | `TRANSPORT_TICKETS_ACTION_CONFIG` — the transport-tickets action's settings |
| `src/08-action-transport-tickets.js` | The transport-tickets action (.ics-VEVENT-sourced, reuses the ICS action's parser, action descriptor) |
| `src/appsscript.json` | Apps Script manifest (timezone, OAuth scopes, Advanced Calendar/Drive/Documents Services) |

Each action's own settings live in a `const *_CONFIG` object in its own
sibling `*-cfg-*.js` file — global, cross-action settings live in
`01-setup.js`. An action descriptor reads its config through a `get
config()` getter rather than a plain object property, so which of the two
files Apps Script happens to load first never matters.

## Setup

1. Install [clasp](https://github.com/google/clasp) and log in:
   ```
   npm install -g @google/clasp
   clasp login
   ```
2. Create a new Apps Script project (or use an existing one) and copy
   `.clasp.json.example` to `.clasp.json`, filling in your script ID:
   ```
   cp .clasp.json.example .clasp.json
   ```
3. Push the source:
   ```
   clasp push
   ```
4. Run `rebuildScriptProperties()` once from the function picker, then set
   your real values in Project Settings → Script Properties (recommended —
   see [Live settings override](#live-settings-override-script-properties),
   these survive every future `clasp push`). At minimum, set
   `01-setup-CALENDAR_ID` to your target Google Calendar's ID (find it under
   Calendar Settings → Settings and sharing → Integrate calendar). Editing
   `CONFIG`/`ICS_ACTION_CONFIG`/`BOOKING_ACTION_CONFIG` directly in the code
   files still works too, but any value set there gets overwritten on your
   next `clasp push`.
5. In the Apps Script editor, confirm the **Calendar** Advanced Service
   shows as enabled (Services panel, left sidebar) — required for
   `Events.import`. If it's not listed, add it (its scope is already
   declared in the manifest, so this should already be linked after a
   `clasp push`).
6. Run `setup()` once from the editor and authorize the requested scopes
   (Gmail modify, Calendar, Drive, Documents, mail send, user email — Drive
   and Documents are needed only by the ticketing-portals action's PDF/OCR
   pipeline). This creates the processed/failed Gmail labels and installs
   the time-driven trigger.
7. `processEmails()` now runs automatically on the configured interval, or
   you can run it manually to test.

## Configuration reference

Every field below can also be set live via Script Properties instead of
editing code — see
[Live settings override](#live-settings-override-script-properties) for the
exact property key each one maps to.

**Global** — `CONFIG` at the top of `src/01-setup.js` (no bound spreadsheet):

| Field | Meaning |
|-------|---------|
| `calendarId` | DEFAULT/fallback Google Calendar ID — used by both actions unless an action-level or per-sender override below applies |
| `labelName` | Gmail label applied once a thread has been processed |
| `failedLabelName` | Additional label applied if any action failed |
| `daysBack` | How many days back to search for unlabeled threads |
| `installTrigger` | Whether `setup()` installs the time-driven trigger |
| `triggerIntervalMinutes` | Trigger interval — must be one of `1, 5, 10, 15, 30` |
| `ticketAttachmentDriveFolderName` | Drive folder name (looked up or created, not a folder ID) where a ticket PDF is stored when a portal's `insertPdfIntoEvent` is on |

**ICS import action** — `ICS_ACTION_CONFIG` in
`src/05-action-cfg-ics-import.js`:

| Field | Meaning |
|-------|---------|
| `enabled` | Cross-cutting toggle — `false` skips this action entirely |
| `notifyOnFailure` | Whether a failure notification email is sent when this action throws |
| `importOnlyFrom` | Array of sender email addresses to restrict import to; empty (default) imports from any sender |
| `excludeFrom` | Array of sender email addresses to skip entirely, regardless of `importOnlyFrom` — empty (default) excludes nobody. Use this to hand a sender off to a more specific action (e.g. the transport-tickets action) so the two don't both create an event for the same email |
| `calendarId` | Overrides `CONFIG.calendarId` for every ICS import, regardless of sender, unless a `calendarIdBySender` entry below also matches |
| `calendarIdBySender` | Array of `{ from, calendarId }` — routes a specific sender's `.ics` invites to a specific calendar, taking priority over both this action's own `calendarId` and the global default. First match in list order wins. Independent of `importOnlyFrom` (a separate gate for whether a message is processed at all) |

**Booking.com management action** — `BOOKING_ACTION_CONFIG` in
`src/06-action-cfg-booking-com-management.js` (language-dependent label text
lives separately — see [Language packs](#language-packs)):

| Field | Meaning |
|-------|---------|
| `enabled` | Cross-cutting toggle — `false` skips this action entirely |
| `notifyOnFailure` | Whether a failure notification email is sent when this action throws |
| `senderAllowList` | Sender addresses this action processes; defaults to `noreply@booking.com` only |
| `addToCalendar.enabled` | Whether the confirmation-email safety-net event creation is active |
| `removeFromCalendar.enabled` | Whether the cancellation-email event deletion is active |
| `searchWindowPaddingDays` | Days padded around check-in/check-out when enumerating candidate calendar events |
| `wideFallbackWindowYears` | Fallback search window half-width (years) when check-in/check-out dates can't be parsed |
| `eventOverlapToleranceDays` | Slop tolerance (days) for the fuzzy date-overlap match |
| `calendarId` | Overrides `CONFIG.calendarId` for ALL of this action's Calendar API calls (add and remove paths alike). No per-sender map — booking.com only ever sends from one known address in practice |

**Ticketing-portals action** — `TICKETING_PORTALS_ACTION_CONFIG` in
`src/07-action-cfg-ticketing-portals.js`:

| Field | Meaning |
|-------|---------|
| `enabled` | Cross-cutting toggle — `false` skips this action entirely |
| `notifyOnFailure` | Whether a failure notification email is sent when this action throws |
| `ticketingPortals` | Array of `{ identifyingEmail, calendarId, insertPdfIntoEvent }` — one entry per supported ticketing portal, matched by its confirmation email's sender address. `calendarId` falls back to `CONFIG.calendarId` when unset. See [Ticketing portals](#ticketing-portals) |

**Transport-tickets action** — `TRANSPORT_TICKETS_ACTION_CONFIG` in
`src/08-action-cfg-transport-tickets.js`:

| Field | Meaning |
|-------|---------|
| `enabled` | Cross-cutting toggle — `false` skips this action entirely |
| `notifyOnFailure` | Whether a failure notification email is sent when this action throws |
| `transportSenders` | Array of `{ identifyingEmail, calendarId, insertPdfIntoEvent, mode }` — one entry per supported carrier, matched by its confirmation email's sender address. `mode` is `'ics'` or `'body'` (see [Transport tickets](#transport-tickets)); always set it explicitly for a new carrier — an entry with no `mode` falls back to `'ics'` unless a body parser is registered for that address in the source code. `calendarId` falls back to `CONFIG.calendarId` when unset. Ships with a RegioJet (`mode: 'ics'`) and an IDOS.cz (`mode: 'body'`) entry by default. For an `'ics'`-mode carrier, remember to also add its sender address to `ICS_ACTION_CONFIG.excludeFrom` above |

## Live settings override (Script Properties)

`clasp push` always overwrites `src/*.js` with your local copy — so any
setting you edit directly in a code file gets silently reset on the next
push. Every setting above can instead be set in **Script Properties**
(Apps Script editor → Project Settings → Script Properties), which `clasp
push` never touches at all. The code file's own value becomes just the
fallback default and living documentation (comments, examples) — the
*actual* live value, once set, survives every future push.

**Setup:** run `rebuildScriptProperties()` once from the function picker.
It writes all 28 settings as visible rows — anything you've already
customized in code is preserved as-is, and everything else appears with the
literal placeholder text `DefaultValue` (Script Properties can't store an
empty value, so this stands in for "not overridden — use the code
default"). Edit any row's value to activate that override; set it back to
the literal text `DefaultValue`, or delete the row, to revert to the code
default. Re-run `rebuildScriptProperties()` any time a future update adds a
new setting — it only fills in what's missing, never touches an existing
value, and keeps every row in a consistent order.

**Property key naming:** each key is prefixed with the file it comes
from — `01-setup-*` (from `CONFIG`), `05-action-ics-*` (from
`ICS_ACTION_CONFIG`), `06-action-booking-com-*` (from `BOOKING_ACTION_CONFIG`),
`07-action-ticketing-portals-*` (from `TICKETING_PORTALS_ACTION_CONFIG`),
`08-action-transport-tickets-*` (from `TRANSPORT_TICKETS_ACTION_CONFIG`)
— followed by the setting name in caps, e.g. `01-setup-CALENDAR_ID`,
`06-action-booking-com-CALENDAR_ID`. Simple list settings
(`importOnlyFrom`, `excludeFrom`, `senderAllowList`) are entered as a
comma-separated string (`alice@example.com, bob@example.com`);
`calendarIdBySender`, `ticketingPortals`, and `transportSenders` are the
settings that need valid JSON — **note that JSON requires double-quoted
keys and string values** (`[{"from":"boss@work.com","calendarId":"abc@group.calendar.google.com"}]`),
unlike a plain JavaScript object literal (single quotes / unquoted keys),
which will silently fail to parse and fall back to the code default with no
visible error. Each setting's own code-file comment shows a correct example
value ready to paste in.

**`checkScriptPropertiesSyntax()`** — run manually any time to validate
every stored value against its expected type (boolean, number, or the
`calendarIdBySender` JSON shape) without changing anything. Reports the
exact key and problem for anything malformed. At runtime, a malformed value
never crashes `processEmails()` — it's used as the trigger to catch a typo
ahead of time, not a hard requirement; a bad value just falls back to the
code default silently.

## Testing

The parser, builder, matcher, and notification-body logic are pure functions
(no GAS globals), exported via a guarded `module.exports` that's inert under
the Apps Script runtime, so they're testable under plain Node:

```
node --test test/*.test.js
```

## License

MIT — see [LICENSE](./LICENSE).

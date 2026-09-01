# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single Google Apps Script file (`newsletter_autolabel.gs`) that gets pasted into
the Apps Script editor of a Google Sheet. It replaces a native Gmail filter that
OR-matches sender addresses and has outgrown Gmail's filter character limit.
"Add an OR clause to a filter" becomes "add a row to a sheet."

The Sheet is the datastore — a `Senders` tab with `email` / `added` columns. There
is no database and no server; the Sheet, the Gmail mailbox, and a 15-minute
Apps Script time trigger are the entire runtime.

## Repository state

`newsletter_autolabel.gs` is the only source and the only deployable. Tests live
in `test/` and have **zero dependencies** — they use Node's built-in test runner,
so there is no `node_modules`, no linter config, and no `appsscript.json` in the
repo (the manifest lives only in the deployed Apps Script project). `clasp` and
`eslint` are **not** installed. Do not document or invoke a command without
checking it exists first.

## Commands

```bash
npm test     # node --test test/  — 39 unit tests, no deps
npm run check  # syntax-only parse of the .gs file
```

`npm run check` exists because `node --check` rejects the `.gs` extension; it
goes through `vm.Script` instead, which needs no temp copy. It proves the file
parses as JavaScript and nothing more.

Run a single test file or one case:

```bash
node --test test/address.test.js
node --test --test-name-pattern 'quote' test/
```

## How the tests reach a non-module

`.gs` has no exports and no `require` — it is pasted verbatim into the Apps
Script editor and runs there against globals that do not exist locally.
`test/harness.js` therefore reads the real file and evaluates it in a `vm`
context with `Gmail`, `SpreadsheetApp`, `ScriptApp`, and `Logger` stubbed.

The property this buys, and the reason to keep it: **the file you edit is the
file that runs is the file that is tested.** There is no build step and no
generated copy to drift. Un-stubbed globals are `Proxy` traps that throw, so a
test which accidentally reaches the Gmail API fails loudly instead of silently
passing.

Two gotchas: only `function` declarations land on the sandbox object (top-level
`const` like `SHEET_ID` is block-scoped and unreachable — pass such values as
arguments), and `test/harness.js` must be updated if the script starts using a
new Apps Script global.

## Deploying a change

Editing this file does not deploy it. The deployed copy is whatever is pasted in
the Sheet's Apps Script editor, so a change here reaches production only when
someone re-pastes it (or, if the project is later wired to clasp, `clasp push`).
Treat repo and deployed script as two copies that drift.

`setup()` is run once by hand from the editor. It is idempotent by construction:
`initSheet_()` only writes the header when the sheet is empty, and
`removeExistingTriggers_()` deletes existing `runNewsletterAutoLabel` triggers
before creating a new one — so re-running `setup()` does not accumulate duplicate
triggers. Preserve that property when editing either function.

## Architecture

`runNewsletterAutoLabel()` is the trigger entry point and runs two passes in order:

1. **`processQueue_()` — learn senders.** Finds messages the user tagged with the
   `to-be-filtered` label, reads the `From` header, appends new addresses to the
   Sheet, and strips the trigger label. This is the human input channel: tagging a
   message is how a user "adds a filter rule."
2. **`labelFromKnownSenders_()` — apply.** For every address in the Sheet, finds
   mail from that sender `in:anywhere` lacking the target label and applies
   `.Newsletters` + `CATEGORY_UPDATES` while removing `INBOX` + `SPAM`.

The two-pass ordering is deliberate and load-bearing: pass 1 removes the trigger
label without labeling the message, because pass 2 in the *same run* catches it
along with the sender's entire history. That is also why there is no separate
backfill routine — pass 2 is simultaneously the steady-state path and the
one-time backfill for a newly added sender.

**The file is split into a pure section and an I/O section.** Everything under
`// ---- pure logic (unit-tested; see test/) ----` touches no API and is covered
by tests: `normalizeAddressPadding_`, `isValidSenderAddress_`, `extractAddress_`, `buildSenderQuery_`,
`shouldContinuePaging_`. The functions below it are thin wrappers that call
those. Keep new logic on the pure side of that line wherever it can go there.

**`extractAddress_` returns `null` on anything malformed, and callers must
handle it.** It is deliberately stricter than RFC 5322: whatever it returns is
interpolated into a Gmail `from:"..."` query, so a value containing a quote
could break out of the clause and match unintended mail — which pass 2 would
then archive and relabel. `isValidSenderAddress_` also filters rows read from
the Sheet, because the Sheet is hand-editable and junk can arrive without ever
passing through `extractAddress_`. `buildSenderQuery_` throws rather than
building a query from a bad address.

**The per-sender cap counts messages examined, not messages modified.** This is
why `shouldContinuePaging_` takes `fetched`. Counting modifications meant a
sender whose mail was already labelled never incremented the counter and
paginated through their whole archive at one API read per message — the most
likely way to hit the ~6 minute limit.

**Two label ID spaces coexist in the same `labelIds` array, and this is correct.**
User labels (`.Newsletters`, `to-be-filtered`) have opaque IDs that must be
resolved by name through `getLabelIdByName_()`. System labels (`INBOX`, `SPAM`,
`CATEGORY_UPDATES`) *are* their own IDs and are correctly hardcoded as string
literals. Do not "normalize" the hardcoded ones into name lookups — the lookup
would fail.

Everything uses the **Gmail advanced service** (`Gmail.Users.Messages.*`), not the
built-in `GmailApp`. It requires `enabledAdvancedServices` in the manifest and is
enabled via Services (+) in the editor. The two are not interchangeable: message
IDs, label semantics, and quota accounting differ. Do not mix them in one code path.

## Constraints that shape the code

- **Quota is the dominant cost.** Every message costs a `messages.get` plus,
  when it needs changing, a `messages.modify`. `MAX_MESSAGES_PER_ADDRESS_PER_RUN`
  (200) caps per-sender work per run; `BACKFILL_QUERY_SUFFIX` caps history depth.
  Any change that adds a per-message API call multiplies against a sender's whole
  archive on first backfill.
- **Apps Script kills executions at ~6 minutes.** Work must stay resumable across
  runs. The pagination loops and the per-sender cap exist for this reason — a run
  that hits the cap simply finishes the rest on the next 15-minute tick.
- **Runs overlap.** The script is not re-entrant-safe; a slow run can still be
  going when the next trigger fires. Operations are written to be idempotent
  (already-labeled messages are skipped, known addresses are not re-appended) so
  overlap is harmless rather than prevented.
- `SHEET_ID` is a placeholder (`PUT_YOUR_SHEET_ID_HERE`) and is expected to stay
  that way in the repo. Do not commit a real Sheet ID.

## Reviewing changes here

The file is syntactically JavaScript but semantically Apps Script, so ordinary JS
review instincts mislead in specific ways worth checking against:

- Apps Script runs V8 but is not Node — no modules, no `require`, no `async`/`await`
  against the Gmail service. Top-level `const`s are genuine globals shared across
  every function.
- The trailing-underscore convention (`initSheet_`, `readKnownSenders_`) is not
  cosmetic: Apps Script treats it as private and hides those functions from the
  editor's Run dropdown and from being callable as a trigger handler. Renaming
  `runNewsletterAutoLabel` or adding an underscore to it breaks the trigger, which
  references the handler by name string.
- Errors are surfaced via `Logger.log` and visible in the editor's execution log.
  There is no other error channel — a thrown exception in a triggered run is
  reported only by Apps Script's own failure email.

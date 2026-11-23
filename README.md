# gmail-newsletter-filter

A Google Apps Script that labels newsletters in Gmail by sender, using a Google
Sheet as the list of senders.

It exists because Gmail filters have a character limit. A filter that
OR-matches sender addresses eventually stops accepting new ones, and the fix —
splitting it across several filters — makes the rule set harder to reason about
than the problem it solves. This replaces "add another OR clause to a filter"
with "add a row to a sheet," and the row gets added for you.

More about the why: [projectinsomnia.com/blog/gmail-filter-character-limit](https://projectinsomnia.com/blog/gmail-filter-character-limit/)

## How it works

Tag a stray newsletter with the Gmail label `.to-be-filtered`. Within 15
minutes, that message and **everything else that sender has ever sent you** is
labeled `.Newsletters`, categorized under Updates, and removed from the inbox.

To stop filtering a sender, tag one of their messages `.remove-from-filter`.
The next run deletes that sender's row from the sheet and removes the tag.
Mail that is already labeled stays where it is; only future mail is affected.

A time-driven trigger runs three passes in order:

1. **Learn.** Find messages you tagged `.to-be-filtered` and read the `From`
   header. Add the address to the `Senders` sheet if it is new, then remove
   the trigger label.
2. **Forget.** Find messages you tagged `.remove-from-filter`, delete every
   sheet row for that sender, and remove the removal label. This runs after
   Learn, so a sender tagged with both labels ends up removed.
3. **Apply.** For every address in the sheet, find mail from that sender
   `in:anywhere` that lacks the target label. Add `.Newsletters` and
   `CATEGORY_UPDATES`; remove `INBOX` and `SPAM`.

The ordering matters. Pass 1 removes the trigger label without labeling the
message. Pass 3 in the *same run* picks it up, along with the sender's entire
history. That is also why there is no separate backfill routine. Pass 3 is both
the steady-state path and the one-time backfill for a newly added sender.

## Setup

You need a Google account, and nothing else. There is no server, no database,
and nothing to install locally.

1. Create a Google Sheet. Note its ID — the long string in the URL between
   `/d/` and `/edit`.
2. In the Sheet, open **Extensions → Apps Script**. Delete the default code and
   paste the entire contents of [`newsletter_autolabel.gs`](newsletter_autolabel.gs).
3. Set `SHEET_ID` at the top of the file to the ID from step 1.
4. In the Apps Script editor, click **Services (+)** and add **Gmail API**.
   Leave the version as-is. This is the *advanced* Gmail service, which the
   script requires; the built-in `GmailApp` is a different API and is not
   interchangeable.
5. Create the Gmail labels `.to-be-filtered` and `.remove-from-filter`, plus
   your target label (`.Newsletters` by default) if it does not exist yet.
6. Select `setup` in the function dropdown and click **Run**. Approve the
   permission prompts.

`setup` creates the sheet header row and a trigger that runs every 15 minutes.
It is idempotent — running it again does not create duplicate triggers.

## Configuration

All settings live at the top of `newsletter_autolabel.gs`.

| Constant | Default | What it does |
| --- | --- | --- |
| `SHEET_ID` | `PUT_YOUR_SHEET_ID_HERE` | The Sheet that stores sender addresses. |
| `SHEET_TAB_NAME` | `Senders` | Tab within that Sheet. |
| `TRIGGER_LABEL_NAME` | `.to-be-filtered` | The label you apply to teach it a sender. |
| `REMOVAL_LABEL_NAME` | `.remove-from-filter` | The label you apply to make it forget a sender. |
| `TARGET_LABEL_NAME` | `.Newsletters` | The label it applies. |
| `BACKFILL_QUERY_SUFFIX` | `in:anywhere` | How far back to search. Narrow it to cap history. |
| `MAX_MESSAGES_PER_ADDRESS_PER_RUN` | `200` | Messages examined per sender per run. |
| `MAX_REMOVALS_PER_RUN` | `100` | Removal requests examined per run. The rest wait for the next run. |

The first backfill for a high-volume sender — a newspaper with years of
archived mail — burns Gmail API quota faster than a low-volume one. If that
matters to you, narrow `BACKFILL_QUERY_SUFFIX` to something like
`in:anywhere newer_than:1y` before the first run.

This script does not touch any existing native Gmail filter. Both can run at
the same time, and an address covered by both is harmless — matching mail
already carries the label, so the script does nothing to it.

## Development

The script is one file with no build step. It is pasted verbatim into the Apps
Script editor, so the file you edit is the file that runs.

```bash
npm test      # 71 unit tests via node --test, zero dependencies
npm run check # syntax-only parse of the .gs file
```

There is no `node_modules`. The tests use Node's built-in test runner.
`test/harness.js` reads the real `.gs` file and evaluates it in a `vm` context.
It stubs `Gmail`, `SpreadsheetApp`, `ScriptApp`, `LockService`, and `Logger`. Un-stubbed
globals are `Proxy` traps that throw, so a test that accidentally reaches the
Gmail API fails loudly instead of quietly passing.

Tests cover the pure-logic section of the file: address extraction and
validation, query construction, pagination, sender-row matching, and the
`SHEET_ID` configuration check. They also run the removal pass against stubbed
Gmail and Sheets. The pure section is kept apart from the I/O below it. New
logic belongs on the pure side of that line wherever it can go there.

Because the tests stub all I/O, they cannot catch a mistake in how the Gmail
API itself is called. One such bug did reach production. See
[`CLAUDE.md`](CLAUDE.md) for the architecture notes and the specific traps that
Apps Script sets for ordinary JavaScript instincts.

### Deploying a change

Editing this repository does not deploy anything. The deployed copy is whatever
is currently pasted into the Sheet's Apps Script editor, so a change here
reaches production only when someone re-pastes it. Treat the repository and the
deployed script as two copies that drift.

### Upgrading from `to-be-filtered`

Earlier versions used the label `to-be-filtered` (no leading dot). Before you
re-paste this version:

1. In Gmail, rename `to-be-filtered` to `.to-be-filtered`. Renaming keeps the
   label on messages that already have it.
2. Create the label `.remove-from-filter`.

If you re-paste first, each run logs `label ".to-be-filtered" not found` and
skips new senders until the label is renamed. Nothing is lost: tagged messages
are picked up on the first run after the rename.

## License

MIT — see [LICENSE.md](LICENSE.md).

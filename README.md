# gmail-newsletter-filter

A Google Apps Script that labels newsletters in Gmail by sender, using a Google
Sheet as the list of senders.

It exists because Gmail filters have a character limit. A filter that
OR-matches sender addresses eventually stops accepting new ones, and the fix —
splitting it across several filters — makes the rule set harder to reason about
than the problem it solves. This replaces "add another OR clause to a filter"
with "add a row to a sheet," and the row gets added for you.

## How it works

Tag a stray newsletter with the Gmail label `to-be-filtered`. Within 15
minutes, that message and **everything else that sender has ever sent you** is
labeled `.Newsletters`, categorized under Updates, and removed from the inbox.

A time-driven trigger runs two passes in order:

1. **Learn.** Find messages you tagged `to-be-filtered`, read the `From`
   header, append the address to the `Senders` sheet if it is new, and remove
   the trigger label.
2. **Apply.** For every address in the sheet, find mail from that sender
   `in:anywhere` that lacks the target label. Add `.Newsletters` and
   `CATEGORY_UPDATES`; remove `INBOX` and `SPAM`.

The ordering matters. Pass 1 removes the trigger label without labeling the
message, because pass 2 in the *same run* picks it up along with the sender's
entire history. That is also why there is no separate backfill routine — pass 2
is both the steady-state path and the one-time backfill for a newly added
sender.

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
5. Create the Gmail label `to-be-filtered`, plus your target label
   (`.Newsletters` by default) if it does not exist yet.
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
| `TRIGGER_LABEL_NAME` | `to-be-filtered` | The label you apply to teach it a sender. |
| `TARGET_LABEL_NAME` | `.Newsletters` | The label it applies. |
| `BACKFILL_QUERY_SUFFIX` | `in:anywhere` | How far back to search. Narrow it to cap history. |
| `MAX_MESSAGES_PER_ADDRESS_PER_RUN` | `200` | Messages examined per sender per run. |

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
npm test      # 39 unit tests via node --test, zero dependencies
npm run check # syntax-only parse of the .gs file
```

There is no `node_modules`. The tests use Node's built-in test runner, and
`test/harness.js` reads the real `.gs` file and evaluates it in a `vm` context
with `Gmail`, `SpreadsheetApp`, `ScriptApp`, and `Logger` stubbed. Un-stubbed
globals are `Proxy` traps that throw, so a test that accidentally reaches the
Gmail API fails loudly instead of quietly passing.

Tests cover the pure-logic section of the file — address extraction and
validation, query construction, and pagination. That section is deliberately
separated from the I/O below it, and new logic belongs on the pure side of the
line wherever it can go there.

Because the tests stub all I/O, they cannot catch a mistake in how the Gmail
API itself is called. One such bug did reach production. See
[`CLAUDE.md`](CLAUDE.md) for the architecture notes and the specific traps that
Apps Script sets for ordinary JavaScript instincts.

### Deploying a change

Editing this repository does not deploy anything. The deployed copy is whatever
is currently pasted into the Sheet's Apps Script editor, so a change here
reaches production only when someone re-pastes it. Treat the repository and the
deployed script as two copies that drift.

## License

MIT — see [LICENSE.md](LICENSE.md).

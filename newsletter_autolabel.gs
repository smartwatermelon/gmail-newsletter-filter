/**
 * Newsletter auto-labeling script.
 *
 * Problem this solves: the native Gmail filter that OR-matches sender
 * addresses hits Gmail's character limit. This script replaces "add a new
 * OR clause to a filter" with "add a row to a sheet."
 *
 * Workflow:
 *   1. You get mail from a sender that should be labeled .Newsletters but
 *      isn't. You apply the Gmail label "to-be-filtered" to that message.
 *   2. This script runs on a timer. It finds messages tagged
 *      "to-be-filtered", pulls the sender address, adds it to the "Senders"
 *      sheet (if new), and removes the "to-be-filtered" tag.
 *   3. On every run, for every address in the sheet, the script finds mail
 *      from that address anywhere (inbox, spam, trash) that does not yet
 *      have .Newsletters, and applies: add .Newsletters + CATEGORY_UPDATES,
 *      remove INBOX + SPAM. This covers both new mail going forward and a
 *      one-time backfill the first time an address is added.
 *
 * SETUP
 *   1. Create a Google Sheet. Note its ID (the long string in the URL
 *      between /d/ and /edit).
 *   2. In the Sheet: Extensions > Apps Script. Delete the default code,
 *      paste this whole file.
 *   3. Fill in SHEET_ID below.
 *   4. In the Apps Script editor: Services (+) > add "Gmail API" (advanced
 *      service). Leave the version as-is.
 *   5. Run `setup` once from the editor (select it in the function dropdown,
 *      click Run). Approve the permission prompts. This creates the sheet
 *      header row and the 15-minute trigger.
 *   6. Done. From now on: tag a stray newsletter "to-be-filtered" and it
 *      (and everything else from that sender) gets labeled within 15
 *      minutes.
 *
 * NOTES
 *   - First backfill for a high-volume sender (e.g. a newspaper with years
 *     of mail) can take a while and burns through Gmail API read/write
 *     quota faster than a low-volume one. If that's a problem, narrow
 *     BACKFILL_QUERY_SUFFIX below (e.g. "newer_than:1y") to cap history.
 *   - This script does not touch or replace the existing native
 *     ".Newsletters" filter. Both can run at the same time. Addresses
 *     already covered by the native filter don't need to be added here too
 *     — but there's no harm if they are (matching mail already has the
 *     label, so the script does nothing).
 */

// ---- CONFIG ----
const SHEET_ID = 'PUT_YOUR_SHEET_ID_HERE';
const SHEET_TAB_NAME = 'Senders';
const TRIGGER_LABEL_NAME = 'to-be-filtered';
const TARGET_LABEL_NAME = '.Newsletters';
const BACKFILL_QUERY_SUFFIX = 'in:anywhere'; // add e.g. ' newer_than:2y' to cap history
const MAX_MESSAGES_PER_ADDRESS_PER_RUN = 200; // safety cap per sender per run

/**
 * Fail loudly if SHEET_ID was never filled in.
 *
 * Throws rather than logging: both callers would otherwise go on to do real
 * work against a Sheet that does not exist, and a triggered run reports a
 * thrown exception through Apps Script's own failure email. A Logger.log here
 * would be visible only to someone already reading the execution log.
 */
function assertConfigured_() {
  if (!isConfiguredSheetId_(SHEET_ID)) {
    throw new Error(
      'SHEET_ID is not set. Open this script in the Apps Script editor and ' +
        'replace the SHEET_ID placeholder with your Sheet id — the long ' +
        'string in the Sheet URL between /d/ and /edit.',
    );
  }
}

// ---- SETUP (run once manually) ----
function setup() {
  assertConfigured_();
  initSheet_();
  removeExistingTriggers_();
  ScriptApp.newTrigger('runNewsletterAutoLabel')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Setup complete. Trigger runs every 15 minutes.');
}

function initSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_TAB_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['email', 'added']);
    sheet.setFrozenRows(1);
  }
}

function removeExistingTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runNewsletterAutoLabel') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ---- MAIN ENTRY POINT (runs on the timer) ----
function runNewsletterAutoLabel() {
  assertConfigured_();
  const targetLabelId = getLabelIdByName_(TARGET_LABEL_NAME);
  if (!targetLabelId) {
    Logger.log('ERROR: label "%s" not found.', TARGET_LABEL_NAME);
    return;
  }

  processQueue_(targetLabelId);
  labelFromKnownSenders_(targetLabelId);
}

// Step 1: pick up messages tagged "to-be-filtered", learn their senders.
function processQueue_(targetLabelId) {
  const triggerLabelId = getLabelIdByName_(TRIGGER_LABEL_NAME);
  if (!triggerLabelId) {
    Logger.log('ERROR: label "%s" not found.', TRIGGER_LABEL_NAME);
    return;
  }

  const known = readKnownSenders_();
  let pageToken;
  do {
    const resp = Gmail.Users.Messages.list('me', {
      labelIds: [triggerLabelId],
      maxResults: 100,
      pageToken: pageToken,
    });
    const messages = resp.messages || [];
    messages.forEach(m => {
      const full = Gmail.Users.Messages.get('me', m.id, { format: 'metadata', metadataHeaders: ['From'] });
      const fromHeader = (full.payload.headers.find(h => h.name === 'From') || {}).value || '';
      const address = extractAddress_(fromHeader);
      if (!address) {
        // Unparseable From header. Leave the trigger label in place so the
        // message stays visible as unprocessed — clearing it here would
        // destroy the user's signal and silently drop the request.
        Logger.log('SKIP: could not extract an address from From header %s', JSON.stringify(fromHeader));
        return;
      }
      if (!known.has(address)) {
        appendSender_(address);
        known.add(address);
      }
      // Safe to clear now that the sender is recorded; the per-sender pass
      // below will label this message (and everything else from that sender)
      // on this same run.
      // Advanced service signature is modify(resource, userId, id): body first.
      Gmail.Users.Messages.modify({ removeLabelIds: [triggerLabelId] }, 'me', m.id);
    });
    pageToken = resp.nextPageToken;
  } while (pageToken);
}

// Step 2: for every known sender, label matching mail that isn't labeled yet.
function labelFromKnownSenders_(targetLabelId) {
  // The caller checks this and returns early, so this is defensive. It stays
  // because of the specific damage a null id would do: labelIds.indexOf(null)
  // never matches, so the "already labelled" skip stops working and every
  // message from every known sender is modified again on every run — with
  // null in addLabelIds — pulling mail out of the inbox on a loop.
  if (!targetLabelId) {
    Logger.log('ERROR: labelFromKnownSenders_ called with no targetLabelId; skipping.');
    return;
  }

  const addresses = readKnownSenders_();
  const inboxId = 'INBOX';
  const spamId = 'SPAM';
  const updatesId = 'CATEGORY_UPDATES';

  addresses.forEach(address => {
    let pageToken;
    let fetched = 0; // messages examined — what the cap actually bounds
    do {
      const resp = Gmail.Users.Messages.list('me', {
        q: buildSenderQuery_(address, BACKFILL_QUERY_SUFFIX),
        maxResults: 100,
        pageToken: pageToken,
      });
      const messages = resp.messages || [];
      messages.forEach(m => {
        if (fetched >= MAX_MESSAGES_PER_ADDRESS_PER_RUN) return;
        fetched++;
        const full = Gmail.Users.Messages.get('me', m.id, { format: 'minimal' });
        const labelIds = full.labelIds || [];
        if (labelIds.indexOf(targetLabelId) !== -1) return; // already done
        Gmail.Users.Messages.modify(
          { addLabelIds: [targetLabelId, updatesId], removeLabelIds: [inboxId, spamId] },
          'me',
          m.id,
        );
      });
      pageToken = shouldContinuePaging_(fetched, MAX_MESSAGES_PER_ADDRESS_PER_RUN, resp.nextPageToken)
        ? resp.nextPageToken
        : undefined;
    } while (pageToken);
  });
}

// ---- pure logic (unit-tested; see test/) ----
// These functions touch no Gmail or Sheets API. They are the parts worth
// testing, and test/harness.js loads this file to reach them.

/**
 * Strip the padding a real From header or a hand-edited sheet cell may carry,
 * and nothing else.
 *
 * Deliberately not `String.prototype.trim`, which also strips U+FEFF and the
 * other Unicode spaces. Using trim() here would let a value carrying an
 * invisible character be normalised into one that then passes validation —
 * laundering the anomaly instead of reporting it.
 */
function normalizeAddressPadding_(value) {
  return value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '').toLowerCase();
}

/**
 * True if `value` is something we are willing to put inside a Gmail
 * `from:"..."` clause.
 *
 * ALLOWLIST, not a denylist. An earlier version rejected a set of dangerous
 * characters and passed everything else; that missed Gmail's other query
 * operators — `{}` (OR group), `-` (negate), `*` (wildcard), `:` (operator
 * separator) — each of which reached a live query through a crafted From
 * header. Enumerating what is permitted is the only version of this that
 * stays correct as Gmail's query language grows.
 *
 * Deliberately narrower than RFC 5321, which legitimately permits `{`, `}`,
 * `*` and `-` in a local part. The consequence of accepting a bad value here
 * is mail being archived (pass 2 removes INBOX from whatever the query
 * matches), so this errs toward rejection: a real address wrongly rejected
 * costs one unfiltered newsletter, which is recoverable and visible.
 * ASCII-only by construction, which also excludes fullwidth confusables,
 * zero-width marks, and control characters.
 */
function isValidSenderAddress_(value) {
  if (typeof value !== 'string') return false;
  const v = normalizeAddressPadding_(value);
  // `-` is legal inside a local part (no-reply@) but is Gmail's negation
  // operator at the start of a term, so it is barred from first position.
  return /^[a-z0-9._%+][a-z0-9._%+-]*@[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*\.[a-z]{2,}$/.test(v);
}

/**
 * Extract the address from an RFC 5322 From header.
 *
 * @param {string} fromHeader raw header value, e.g. 'NYT <news@nytimes.com>'
 * @returns {?string} lowercased address, or null if the header does not
 *   contain exactly one well-formed address, as judged solely by
 *   isValidSenderAddress_. Callers MUST handle null: a header with no address
 *   at all used to fall through to the whole header text, which then reached
 *   the sheet and, from there, a Gmail query.
 */
function extractAddress_(fromHeader) {
  if (typeof fromHeader !== 'string') return null;
  const angleMatch = fromHeader.match(/<([^>]+)>/);
  const raw = angleMatch ? angleMatch[1] : fromHeader;
  // Deliberately no trim() here. isValidSenderAddress_ is the single arbiter
  // of what counts as an address, and it forgives only ASCII padding.
  // Normalising first would let trim()'s Unicode-space handling launder a
  // value (e.g. one carrying U+FEFF) into something that then validates.
  if (!isValidSenderAddress_(raw)) return null;
  return normalizeAddressPadding_(raw);
}

/**
 * Build the Gmail search query for one sender.
 *
 * Single home for the from:"..." construction so the escaping rule has one
 * place to live and one place to be tested.
 *
 * @throws {Error} if the address is not valid. Throwing beats returning a
 *   narrower query, which would silently label the wrong mail.
 */
function buildSenderQuery_(address, suffix) {
  if (!isValidSenderAddress_(address)) {
    throw new Error('invalid sender address: ' + JSON.stringify(address));
  }
  const clause = 'from:"' + address + '"';
  return suffix ? clause + ' ' + suffix : clause;
}

/**
 * Whether the per-sender pagination loop should request another page.
 *
 * Counts messages *examined*, not messages *changed*. Counting changes meant
 * a sender whose mail was already labelled never incremented the counter and
 * paginated through their entire archive, one API read per message — the most
 * likely way to hit the ~6 minute execution limit.
 */
function shouldContinuePaging_(fetched, cap, nextPageToken) {
  if (!nextPageToken) return false;
  return fetched < cap;
}

/**
 * True if SHEET_ID has actually been filled in.
 *
 * The file ships with the placeholder and is expected to keep it in the repo,
 * so "pasted the script, skipped the SHEET_ID step, ran setup()" is a normal
 * mistake rather than an exotic one. Without this check the first symptom is
 * whatever SpreadsheetApp.openById() throws for a non-existent document, which
 * says nothing about the step that was missed.
 *
 * Takes the id as an argument, not from the global, so it stays on the tested
 * side of the line — top-level const is unreachable from the sandbox.
 */
function isConfiguredSheetId_(sheetId) {
  if (typeof sheetId !== 'string') return false;
  const trimmed = sheetId.trim();
  if (!trimmed) return false;
  return trimmed.toUpperCase() !== 'PUT_YOUR_SHEET_ID_HERE';
}

// ---- sheet + Gmail helpers ----
function readKnownSenders_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TAB_NAME);
  const lastRow = sheet.getLastRow();
  const set = new Set();
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    values.forEach(row => {
      const raw = String(row[0] || '');
      // The sheet is hand-editable, so junk can arrive without ever passing
      // through extractAddress_. Drop it here rather than at query time.
      // Validate the raw cell for the same reason extractAddress_ does: the
      // validator, not a prior trim(), decides what an address is.
      if (isValidSenderAddress_(raw)) set.add(normalizeAddressPadding_(raw));
    });
  }
  return set;
}

function appendSender_(address) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TAB_NAME);
  sheet.appendRow([address, new Date()]);
}

function getLabelIdByName_(name) {
  const labels = Gmail.Users.Labels.list('me').labels || [];
  const found = labels.find(l => l.name === name);
  return found ? found.id : null;
}

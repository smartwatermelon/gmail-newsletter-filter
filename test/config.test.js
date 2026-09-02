"use strict";

const test = require("node:test");
const assert = require("node:assert");
const { loadScript } = require("./harness");

// The placeholder guard. SHEET_ID ships as PUT_YOUR_SHEET_ID_HERE and is meant
// to stay that way in the repo, so the failure mode this catches is a real one:
// paste the file, skip step 3 of the setup, and run setup(). Without the guard
// SpreadsheetApp.openById() throws something about a missing document, which
// does not point at the actual mistake.
//
// This takes the id as an argument rather than reading the global, because
// top-level const is not reachable from the sandbox. See test/harness.js.
test("isConfiguredSheetId_", async (t) => {
  const { isConfiguredSheetId_ } = loadScript();

  await t.test("rejects the shipped placeholder", () => {
    assert.strictEqual(isConfiguredSheetId_("PUT_YOUR_SHEET_ID_HERE"), false);
  });

  await t.test("rejects the placeholder regardless of case", () => {
    assert.strictEqual(isConfiguredSheetId_("put_your_sheet_id_here"), false);
  });

  await t.test("rejects the placeholder with stray whitespace", () => {
    assert.strictEqual(
      isConfiguredSheetId_("  PUT_YOUR_SHEET_ID_HERE  "),
      false,
    );
  });

  await t.test("rejects empty and whitespace-only ids", () => {
    assert.strictEqual(isConfiguredSheetId_(""), false);
    assert.strictEqual(isConfiguredSheetId_("   "), false);
  });

  await t.test("rejects a non-string id", () => {
    assert.strictEqual(isConfiguredSheetId_(null), false);
    assert.strictEqual(isConfiguredSheetId_(undefined), false);
    assert.strictEqual(isConfiguredSheetId_(42), false);
  });

  await t.test("accepts a real-looking Sheet id", () => {
    assert.strictEqual(
      isConfiguredSheetId_("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"),
      true,
    );
  });
});

// The defensive guard in labelFromKnownSenders_. The caller already checks
// targetLabelId and returns early, so this is unreachable in normal operation.
// It is here because of what happens if it ever stops being unreachable: a
// null id makes the "already labelled" check (labelIds.indexOf(null)) never
// match, so every message from every known sender is re-modified on every run
// — with null in addLabelIds — archiving mail out of the inbox on a loop.
// Cheap guard, expensive absence.
test("labelFromKnownSenders_ refuses a missing target label", async (t) => {
  await t.test("returns without touching Gmail when the id is null", () => {
    // Gmail is left un-stubbed: the harness turns it into a trap that throws
    // on any property access, so reaching the API fails the test loudly.
    const { labelFromKnownSenders_, __logged } = loadScript();

    assert.doesNotThrow(() => labelFromKnownSenders_(null));

    const messages = __logged.map((entry) => entry.join(" "));
    assert.ok(
      messages.some((m) => m.includes("targetLabelId")),
      `expected a log naming targetLabelId, got: ${JSON.stringify(messages)}`,
    );
  });

  await t.test("also refuses undefined and empty string", () => {
    const { labelFromKnownSenders_ } = loadScript();
    assert.doesNotThrow(() => labelFromKnownSenders_(undefined));
    assert.doesNotThrow(() => labelFromKnownSenders_(""));
  });
});

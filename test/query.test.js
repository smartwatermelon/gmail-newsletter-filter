"use strict";

const test = require("node:test");
const assert = require("node:assert");
const { loadScript } = require("./harness");

// buildSenderQuery_ is the single place a sender address becomes Gmail query
// syntax. Isolating it means the escaping rule has one home and one test.
test("buildSenderQuery_", async (t) => {
  const { buildSenderQuery_ } = loadScript();

  await t.test("builds a quoted from: clause with the suffix", () => {
    assert.strictEqual(
      buildSenderQuery_("news@nytimes.com", "in:anywhere"),
      'from:"news@nytimes.com" in:anywhere',
    );
  });

  await t.test("supports a narrowed backfill window", () => {
    assert.strictEqual(
      buildSenderQuery_("news@nytimes.com", "in:anywhere newer_than:1y"),
      'from:"news@nytimes.com" in:anywhere newer_than:1y',
    );
  });

  await t.test("tolerates an empty suffix without trailing space", () => {
    assert.strictEqual(buildSenderQuery_("a@b.com", ""), 'from:"a@b.com"');
  });

  await t.test("refuses to build a query from an invalid address", () => {
    // Defence in depth: even if a bad value reaches here, it must not become
    // a query. Throwing is correct — silently returning a narrower query
    // would hide the bug.
    assert.throws(
      () => buildSenderQuery_('a" OR from:"boss@work.com', "in:anywhere"),
      /invalid/i,
    );
    assert.throws(() => buildSenderQuery_("", "in:anywhere"), /invalid/i);
    assert.throws(() => buildSenderQuery_(null, "in:anywhere"), /invalid/i);
  });
});

// The per-sender cap exists to keep a run inside Apps Script's ~6 minute
// execution limit. It previously counted modifications, so a sender whose mail
// was already labelled paginated forever while the counter stayed at zero.
test("shouldContinuePaging_", async (t) => {
  const { shouldContinuePaging_ } = loadScript();

  await t.test("continues when under the cap and a page remains", () => {
    assert.strictEqual(shouldContinuePaging_(50, 200, "token"), true);
  });

  await t.test("stops when there is no next page", () => {
    assert.strictEqual(shouldContinuePaging_(50, 200, undefined), false);
    assert.strictEqual(shouldContinuePaging_(50, 200, null), false);
  });

  await t.test("stops once fetches reach the cap", () => {
    assert.strictEqual(shouldContinuePaging_(200, 200, "token"), false);
  });

  await t.test("stops if fetches somehow exceed the cap", () => {
    assert.strictEqual(shouldContinuePaging_(250, 200, "token"), false);
  });

  await t.test("counts fetches, not modifications", () => {
    // The regression: 200 messages examined, none needing a change. The old
    // logic saw modified === 0, concluded it was under the cap, and kept
    // paging through the sender's entire archive.
    assert.strictEqual(
      shouldContinuePaging_(200, 200, "token"),
      false,
      "a run that examined 200 messages must stop even if it changed none",
    );
  });
});

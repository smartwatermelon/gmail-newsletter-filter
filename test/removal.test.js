"use strict";

const test = require("node:test");
const assert = require("node:assert");
const { loadScript, fakeSpreadsheetApp } = require("./harness");

// findSenderRows_ decides which sheet rows a removal request deletes.
test("findSenderRows_", async (t) => {
  const { findSenderRows_ } = loadScript();

  await t.test("returns 1-based sheet rows, highest first", () => {
    const values = [["a@x.com"], ["b@x.com"], ["a@x.com"]];
    assert.deepStrictEqual(
      Array.from(findSenderRows_(values, new Set(["a@x.com"]))),
      [4, 2],
    );
  });

  await t.test("matches padded and mixed-case cells", () => {
    const values = [["  A@X.com \t"], ["b@x.com"]];
    assert.deepStrictEqual(
      Array.from(findSenderRows_(values, new Set(["a@x.com"]))),
      [2],
    );
  });

  await t.test("ignores junk cells and non-matching senders", () => {
    const values = [[""], [null], ["not an address"], ["b@x.com"]];
    assert.deepStrictEqual(
      Array.from(findSenderRows_(values, new Set(["a@x.com"]))),
      [],
    );
  });

  await t.test("handles several addresses at once", () => {
    const values = [["a@x.com"], ["b@x.com"], ["c@x.com"]];
    assert.deepStrictEqual(
      Array.from(findSenderRows_(values, new Set(["a@x.com", "c@x.com"]))),
      [4, 2],
    );
  });
});

/**
 * Fake of the Gmail advanced service surface processRemovals_ touches.
 * `messages` maps id -> From header for messages carrying the removal label.
 */
function fakeGmail({ labels, messages }) {
  const modified = [];
  return {
    __modified: modified,
    Users: {
      Labels: { list: () => ({ labels }) },
      Messages: {
        list: () => ({ messages: Object.keys(messages).map((id) => ({ id })) }),
        get: (_user, id) => ({
          payload: { headers: [{ name: "From", value: messages[id] }] },
        }),
        modify: (resource, user, id) => modified.push({ resource, user, id }),
      },
    },
  };
}

/** Script lock fake. `acquired` controls what tryLock returns. */
function fakeLockService(acquired = true) {
  const events = [];
  return {
    __events: events,
    getScriptLock: () => ({
      tryLock: () => {
        events.push("tryLock");
        return acquired;
      },
      releaseLock: () => events.push("release"),
    }),
  };
}

test("processRemovals_", async (t) => {
  const removalLabel = { name: ".remove-from-filter", id: "Label_9" };

  await t.test("deletes every row for the sender and clears the label", () => {
    const sheet = fakeSpreadsheetApp([
      ["email", "added"],
      ["a@x.com", "d1"],
      ["b@x.com", "d2"],
      [" A@x.com", "d3"],
    ]);
    const Gmail = fakeGmail({
      labels: [removalLabel],
      messages: { m1: "News <a@x.com>" },
    });
    const LockService = fakeLockService();
    const ctx = loadScript({ Gmail, SpreadsheetApp: sheet, LockService });
    ctx.processRemovals_();

    assert.deepStrictEqual(LockService.__events, ["tryLock", "release"]);
    assert.deepStrictEqual(sheet.__rows, [
      ["email", "added"],
      ["b@x.com", "d2"],
    ]);
    // Body first: modify(resource, userId, id).
    assert.deepStrictEqual(JSON.parse(JSON.stringify(Gmail.__modified)), [
      { resource: { removeLabelIds: ["Label_9"] }, user: "me", id: "m1" },
    ]);
  });

  await t.test("clears the label when the sender is not in the sheet", () => {
    const sheet = fakeSpreadsheetApp([
      ["email", "added"],
      ["b@x.com", "d"],
    ]);
    const Gmail = fakeGmail({
      labels: [removalLabel],
      messages: { m1: "z@x.com" },
    });
    loadScript({
      Gmail,
      SpreadsheetApp: sheet,
      LockService: fakeLockService(),
    }).processRemovals_();

    assert.strictEqual(sheet.__rows.length, 2);
    assert.strictEqual(Gmail.__modified.length, 1);
  });

  await t.test("keeps the label on an unparseable From header", () => {
    const sheet = fakeSpreadsheetApp([
      ["email", "added"],
      ["b@x.com", "d"],
    ]);
    const Gmail = fakeGmail({
      labels: [removalLabel],
      messages: { m1: "no address here" },
    });
    const ctx = loadScript({ Gmail, SpreadsheetApp: sheet });
    ctx.processRemovals_();

    assert.strictEqual(sheet.__rows.length, 2);
    assert.strictEqual(Gmail.__modified.length, 0);
    assert.match(String(ctx.__logged[0][0]), /SKIP/);
  });

  await t.test("logs and returns when the removal label is missing", () => {
    // SpreadsheetApp left as a trap: a missing label must not touch the sheet.
    const Gmail = fakeGmail({ labels: [], messages: {} });
    const ctx = loadScript({ Gmail });
    ctx.processRemovals_();

    assert.strictEqual(Gmail.__modified.length, 0);
    assert.match(String(ctx.__logged[0][1]), /remove-from-filter/);
  });

  await t.test("dedups several tagged messages from one sender", () => {
    const sheet = fakeSpreadsheetApp([
      ["email", "added"],
      ["a@x.com", "d"],
      ["b@x.com", "d"],
    ]);
    const Gmail = fakeGmail({
      labels: [removalLabel],
      messages: { m1: "a@x.com", m2: "A <a@x.com>" },
    });
    loadScript({
      Gmail,
      SpreadsheetApp: sheet,
      LockService: fakeLockService(),
    }).processRemovals_();

    assert.deepStrictEqual(sheet.__rows, [
      ["email", "added"],
      ["b@x.com", "d"],
    ]);
    assert.strictEqual(Gmail.__modified.length, 2);
  });

  await t.test("handles a header-only sheet", () => {
    const sheet = fakeSpreadsheetApp([["email", "added"]]);
    const Gmail = fakeGmail({
      labels: [removalLabel],
      messages: { m1: "a@x.com" },
    });
    loadScript({
      Gmail,
      SpreadsheetApp: sheet,
      LockService: fakeLockService(),
    }).processRemovals_();

    assert.strictEqual(sheet.__rows.length, 1);
    assert.strictEqual(Gmail.__modified.length, 1);
  });

  await t.test(
    "deletes the last data row of a sheet trimmed to content",
    () => {
      // maxRows 2 = header + one row: a plain deleteRow would leave only the
      // frozen header, which Sheets refuses.
      const sheet = fakeSpreadsheetApp(
        [
          ["email", "added"],
          ["a@x.com", "d"],
        ],
        { maxRows: 2 },
      );
      const Gmail = fakeGmail({
        labels: [removalLabel],
        messages: { m1: "a@x.com" },
      });
      const ctx = loadScript({
        Gmail,
        SpreadsheetApp: sheet,
        LockService: fakeLockService(),
      });
      ctx.processRemovals_();

      assert.deepStrictEqual(sheet.__rows, [["email", "added"]]);
      assert.strictEqual(Gmail.__modified.length, 1);
      assert.strictEqual(ctx.__logged.length, 0);
    },
  );

  await t.test("leaves labels in place when the lock is busy", () => {
    const sheet = fakeSpreadsheetApp([
      ["email", "added"],
      ["a@x.com", "d"],
    ]);
    const Gmail = fakeGmail({
      labels: [removalLabel],
      messages: { m1: "a@x.com" },
    });
    const ctx = loadScript({
      Gmail,
      SpreadsheetApp: sheet,
      LockService: fakeLockService(false),
    });
    ctx.processRemovals_();

    assert.strictEqual(sheet.__rows.length, 2);
    assert.strictEqual(Gmail.__modified.length, 0);
    assert.match(String(ctx.__logged[0][0]), /lock/);
  });

  await t.test(
    "logs a sheet failure instead of throwing, and keeps labels",
    () => {
      const Gmail = fakeGmail({
        labels: [removalLabel],
        messages: { m1: "a@x.com" },
      });
      const SpreadsheetApp = {
        openById: () => {
          throw new Error("Sheets is down");
        },
      };
      const LockService = fakeLockService();
      const ctx = loadScript({ Gmail, SpreadsheetApp, LockService });
      assert.doesNotThrow(() => ctx.processRemovals_());

      assert.strictEqual(Gmail.__modified.length, 0);
      assert.deepStrictEqual(LockService.__events, ["tryLock", "release"]);
      assert.match(String(ctx.__logged[0][0]), /ERROR/);
    },
  );
});

// Removals must run after additions (so remove wins) and before apply (so
// the removed sender is not relabelled in the same run).
test("runNewsletterAutoLabel runs learn, forget, apply in that order", () => {
  const ctx = loadScript();
  const calls = [];
  ctx.assertConfigured_ = () => {};
  ctx.getLabelIdByName_ = () => "Label_1";
  ctx.processQueue_ = () => calls.push("learn");
  ctx.processRemovals_ = () => calls.push("forget");
  ctx.labelFromKnownSenders_ = () => calls.push("apply");
  ctx.runNewsletterAutoLabel();
  assert.deepStrictEqual(calls, ["learn", "forget", "apply"]);
});

test("runNewsletterAutoLabel still applies labels when removals throw", () => {
  const ctx = loadScript();
  const calls = [];
  ctx.assertConfigured_ = () => {};
  ctx.getLabelIdByName_ = () => "Label_1";
  ctx.processQueue_ = () => calls.push("learn");
  ctx.processRemovals_ = () => {
    throw new Error("Gmail is down");
  };
  ctx.labelFromKnownSenders_ = () => calls.push("apply");
  assert.doesNotThrow(() => ctx.runNewsletterAutoLabel());
  assert.deepStrictEqual(calls, ["learn", "apply"]);
  assert.match(String(ctx.__logged[0][0]), /ERROR/);
});

test("processRemovals_ examines at most 100 requests per run", () => {
  const messages = {};
  for (let i = 0; i < 150; i++) messages["m" + i] = `s${i}@x.com`;
  const gets = [];
  const Gmail = fakeGmail({
    labels: [{ name: ".remove-from-filter", id: "Label_9" }],
    messages,
  });
  // Serve two pages so the cap has to stop pagination, not just one page.
  const ids = Object.keys(messages);
  Gmail.Users.Messages.list = (_user, opts) =>
    opts.pageToken
      ? { messages: ids.slice(100).map((id) => ({ id })) }
      : {
          messages: ids.slice(0, 100).map((id) => ({ id })),
          nextPageToken: "p2",
        };
  const realGet = Gmail.Users.Messages.get;
  Gmail.Users.Messages.get = (user, id, opts) => {
    gets.push(id);
    return realGet(user, id, opts);
  };
  const sheet = fakeSpreadsheetApp([["email", "added"]]);
  loadScript({
    Gmail,
    SpreadsheetApp: sheet,
    LockService: fakeLockService(),
  }).processRemovals_();

  assert.strictEqual(gets.length, 100);
  assert.strictEqual(Gmail.__modified.length, 100);
});

"use strict";

const test = require("node:test");
const assert = require("node:assert");
const { loadScript, fakeScriptProperties } = require("./harness");

const HOUR = 60 * 60 * 1000;

test("shouldSendAlert_", async (t) => {
  const { shouldSendAlert_ } = loadScript();
  const now = 1_000 * HOUR;

  await t.test("sends when nothing was sent before", () => {
    assert.strictEqual(shouldSendAlert_(0, now, 6), true);
  });

  await t.test("suppresses inside the interval", () => {
    assert.strictEqual(shouldSendAlert_(now - 5 * HOUR, now, 6), false);
  });

  await t.test("sends once the interval has passed", () => {
    assert.strictEqual(shouldSendAlert_(now - 6 * HOUR, now, 6), true);
  });

  await t.test("treats a last-sent time in the future as long ago", () => {
    assert.strictEqual(shouldSendAlert_(now + HOUR, now, 6), true);
  });
});

test("buildAlert_ includes the error, a link, and the throttle", () => {
  const { buildAlert_ } = loadScript();
  const alert = buildAlert_(new Error("Quota exceeded"), "SCRIPT123", 6);
  assert.match(alert.subject, /failed: Quota exceeded$/);
  assert.match(alert.body, /Error: Quota exceeded/);
  assert.match(
    alert.body,
    /https:\/\/script\.google\.com\/home\/projects\/SCRIPT123\/executions/,
  );
  assert.match(alert.body, /every 6 hours/);
});

/** Stubs for the mail path; `props` is the backing store. */
function alertStubs({ lastAlertAt, mailThrows = false } = {}) {
  const props = {};
  if (lastAlertAt !== undefined) props.lastAlertAt = String(lastAlertAt);
  const sent = [];
  return {
    __props: props,
    __sent: sent,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in props ? props[k] : null),
        setProperty: (k, v) => {
          props[k] = v;
        },
      }),
    },
    MailApp: {
      sendEmail: (to, subject, body) => {
        if (mailThrows) throw new Error("mail quota");
        sent.push({ to, subject, body });
      },
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => "me@x.com" }) },
    ScriptApp: { getScriptId: () => "SCRIPT123" },
  };
}

function failingRun(stubs) {
  const ctx = loadScript(stubs);
  ctx.runPasses_ = () => {
    throw new Error("boom");
  };
  return ctx;
}

test("runNewsletterAutoLabel emails on failure and rethrows", () => {
  const stubs = alertStubs();
  const ctx = failingRun(stubs);
  assert.throws(() => ctx.runNewsletterAutoLabel(), /boom/);
  assert.strictEqual(stubs.__sent.length, 1);
  assert.strictEqual(stubs.__sent[0].to, "me@x.com");
  assert.ok(stubs.__props.lastAlertAt);
});

test("runNewsletterAutoLabel does not email again inside the interval", () => {
  const stubs = alertStubs({ lastAlertAt: Date.now() - 60 * 1000 });
  const ctx = failingRun(stubs);
  assert.throws(() => ctx.runNewsletterAutoLabel(), /boom/);
  assert.strictEqual(stubs.__sent.length, 0);
});

test("a mail failure does not replace the run's own error", () => {
  const stubs = alertStubs({ mailThrows: true });
  const ctx = failingRun(stubs);
  assert.throws(() => ctx.runNewsletterAutoLabel(), /boom/);
  assert.strictEqual(stubs.__props.lastAlertAt, undefined);
  assert.match(String(ctx.__logged[0][0]), /could not send/);
});

test("a successful run sends nothing", () => {
  // MailApp left as a trap: touching it would throw.
  const props = fakeScriptProperties({});
  const ctx = loadScript({ PropertiesService: props });
  ctx.runPasses_ = () => {};
  assert.doesNotThrow(() => ctx.runNewsletterAutoLabel());
});

test("a successful run clears the throttle, so the next failure emails", () => {
  const props = fakeScriptProperties({ lastAlertAt: String(Date.now()) });
  const ctx = loadScript({ PropertiesService: props });
  ctx.runPasses_ = () => {};
  ctx.runNewsletterAutoLabel();
  assert.strictEqual("lastAlertAt" in props.__values, false);
});

test("a failing run keeps the throttle", () => {
  const stubs = alertStubs({ lastAlertAt: Date.now() - 60 * 1000 });
  const before = stubs.__props.lastAlertAt;
  const ctx = failingRun(stubs);
  assert.throws(() => ctx.runNewsletterAutoLabel(), /boom/);
  assert.strictEqual(stubs.__props.lastAlertAt, before);
});

test("a property error after a successful run does not fail the run", () => {
  const ctx = loadScript({
    PropertiesService: {
      getScriptProperties: () => {
        throw new Error("properties down");
      },
    },
  });
  ctx.runPasses_ = () => {};
  assert.doesNotThrow(() => ctx.runNewsletterAutoLabel());
  assert.match(String(ctx.__logged[0][0]), /could not clear/);
});

test("a missing target label is thrown, so it alerts", () => {
  const ctx = loadScript();
  ctx.assertConfigured_ = () => {};
  ctx.getLabelIdByName_ = () => null;
  assert.throws(() => ctx.runPasses_(), /not found/);
});

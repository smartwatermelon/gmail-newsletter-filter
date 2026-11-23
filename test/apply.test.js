"use strict";

const test = require("node:test");
const assert = require("node:assert");
const { loadScript, fakeSpreadsheetApp } = require("./harness");

// The -label: exclusion keeps pass 3 inside Gmail's per-minute quota.
// Testing buildSenderQuery_ alone does not prove the pass uses it.
test("labelFromKnownSenders_ excludes already-labelled mail in its query", () => {
  const queries = [];
  const Gmail = {
    Users: {
      Messages: {
        list: (_user, opts) => {
          queries.push(opts.q);
          return { messages: [] };
        },
      },
    },
  };
  const SpreadsheetApp = fakeSpreadsheetApp([
    ["email", "added"],
    ["a@x.com", "d"],
  ]);
  loadScript({ Gmail, SpreadsheetApp }).labelFromKnownSenders_("Label_1");

  assert.deepStrictEqual(queries, [
    'from:"a@x.com" -label:.newsletters in:anywhere',
  ]);
});

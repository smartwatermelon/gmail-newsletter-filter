"use strict";

/**
 * Loads newsletter_autolabel.gs into a sandbox so its functions can be
 * unit-tested under plain Node.
 *
 * Why this exists: the .gs file is not a Node module. It has no exports and
 * no require, because it is pasted verbatim into the Apps Script editor and
 * runs there against globals (Gmail, SpreadsheetApp, ScriptApp, Logger) that
 * do not exist locally. Rather than keep a second, importable copy of the
 * logic — which would drift from the deployed one — the tests read the real
 * file and evaluate it in a vm context with those globals stubbed.
 *
 * The property this buys: the file you edit is the file that runs is the file
 * that is tested. There is no build step and no generated artifact.
 *
 * Note: only `function` declarations land on the sandbox object. Top-level
 * `const` (SHEET_ID, TARGET_LABEL_NAME, ...) is block-scoped to the script and
 * is NOT reachable from tests. If a test needs one of those values, pass it in
 * as an argument rather than reaching for it here.
 */

const fs = require("fs");
const path = require("path");
const vm = require("vm");

const SCRIPT_PATH = path.join(__dirname, "..", "newsletter_autolabel.gs");

/**
 * Build a fresh sandbox with the script loaded into it.
 *
 * @param {object} [stubs] Apps Script globals to expose. Anything omitted gets
 *   a stub that throws on use, so a test that accidentally reaches the Gmail
 *   or Sheets API fails loudly instead of silently doing nothing.
 * @returns {object} the vm context; script functions hang off it directly.
 */
function loadScript(stubs = {}) {
  const logged = [];

  const context = {
    // Logger is safe to call in any test; capture instead of printing.
    Logger: { log: (...args) => logged.push(args) },
    console,

    // Anything the test did not deliberately stub becomes a trap.
    Gmail: stubs.Gmail || trap("Gmail"),
    SpreadsheetApp: stubs.SpreadsheetApp || trap("SpreadsheetApp"),
    ScriptApp: stubs.ScriptApp || trap("ScriptApp"),

    ...stubs,
  };

  context.__logged = logged;
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(SCRIPT_PATH, "utf8"), context, {
    filename: "newsletter_autolabel.gs",
  });
  return context;
}

/**
 * A stand-in object whose every property access throws. Used for Apps Script
 * globals a given test has no business touching.
 */
function trap(name) {
  return new Proxy(
    {},
    {
      get(_target, prop) {
        throw new Error(
          `test touched un-stubbed global ${name}.${String(prop)} — ` +
            "stub it explicitly if the test needs it",
        );
      },
    },
  );
}

/**
 * Minimal fake of the SpreadsheetApp surface the script uses.
 * Backed by a plain array of rows so assertions can read it directly.
 *
 * @param {Array<Array<*>>} rows initial sheet contents, header row included.
 */
function fakeSpreadsheetApp(rows) {
  const data = rows.map((r) => r.slice());

  const sheet = {
    getLastRow: () => data.length,
    appendRow: (row) => data.push(row.slice()),
    setFrozenRows: () => {},
    getRange: (startRow, startCol, numRows, numCols) => ({
      getValues: () =>
        data
          .slice(startRow - 1, startRow - 1 + numRows)
          .map((r) => r.slice(startCol - 1, startCol - 1 + numCols)),
    }),
  };

  return {
    __rows: data,
    openById: () => ({
      getSheetByName: () => sheet,
      insertSheet: () => sheet,
    }),
  };
}

module.exports = { loadScript, fakeSpreadsheetApp, SCRIPT_PATH };

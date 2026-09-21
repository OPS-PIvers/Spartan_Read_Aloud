/**
 * Regression tests for the case manager picker backend.
 *
 * Apps Script cannot be unit tested directly, so this loads Code.js into a
 * Node VM with fakes for the Apps Script globals getCaseManagers() touches:
 * SpreadsheetApp (a workbook of named sheets), CacheService (an in-memory
 * cache) and the session token validator.
 *
 *   node tests/case-manager-picker.test.js
 *
 * A case manager is any unique email in column D of the "Case Managers" sheet;
 * the "Student Directory" sheet carries the same value in column D, so both
 * are scanned and the union is returned.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');

// ---------------------------------------------------------------- sandbox --

/** In-memory stand-in for CacheService.getScriptCache(). */
function makeCache() {
  const store = {};
  return {
    get: key => (key in store ? store[key] : null),
    put: (key, value) => { store[key] = value; },
    remove: key => { delete store[key]; },
    _store: store
  };
}

const scriptCache = makeCache();

const sandbox = {
  console,
  Logger: { log: () => {} },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  MimeType: { PDF: 'application/pdf', GOOGLE_DOCS: 'application/vnd.google-apps.document' },
  SpreadsheetApp: {}, DriveApp: {}, Drive: {}, UrlFetchApp: {},
  ScriptApp: {}, Utilities: {}, HtmlService: {},
  CacheService: { getScriptCache: () => scriptCache, getUserCache: () => makeCache() },
  DocumentApp: {}, MailApp: {}, Session: {}
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'Constants.js'), 'utf8'), sandbox, { filename: 'Constants.js' });
vm.runInContext(fs.readFileSync(path.join(ROOT, 'Code.js'), 'utf8'), sandbox, { filename: 'Code.js' });

// Every test drives a valid staff session unless it says otherwise.
sandbox.validateAdminToken = token => (token === 'good-token' ? { email: 'staff@school.org', role: 'teacher' } : null);

/** Installs a workbook: { 'Sheet Name': [[row], [row]] }. */
function setSheets(sheets) {
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: name => {
      if (!(name in sheets)) return null;
      return { getDataRange: () => ({ getValues: () => sheets[name] }) };
    }
  });
  // The roster is cached script-wide; start each case from a clean slate.
  scriptCache.remove('case_managers');
}

// --------------------------------------------------------------- fixtures --

const DIRECTORY_HEADER = ['First', 'Last', 'Student Email', 'Case Manager Email'];
const TEACHERS_HEADER = ['First', 'Last', 'Email', 'Password', 'Role', 'Beta'];

function emailsOf(result) {
  return result.caseManagers.map(cm => cm.email);
}

// ------------------------------------------------------------------ runner --

let passed = 0;
const failures = [];

function test(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  ok  ${name}`);
  } catch (err) {
    failures.push({ name, err });
    console.log(`FAIL  ${name}\n      ${err.message}`);
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function assertDeepEqual(actual, expected, message) {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a !== b) throw new Error(`${message}\n      expected ${b}\n      actual   ${a}`);
}

// ------------------------------------------------------------------- tests --

test('column D of the Case Managers sheet becomes the roster', () => {
  setSheets({
    'Case Managers': [
      DIRECTORY_HEADER,
      ['Ann', 'Ames', 'ann.s@school.org', 'cm.one@school.org'],
      ['Bob', 'Best', 'bob.s@school.org', 'cm.two@school.org']
    ]
  });
  const res = sandbox.getCaseManagers('good-token');
  assert(res.success, 'call should succeed');
  assertDeepEqual(emailsOf(res), ['cm.one@school.org', 'cm.two@school.org'], 'both column D emails should be listed');
});

test('duplicate case managers collapse to one entry', () => {
  setSheets({
    'Case Managers': [
      DIRECTORY_HEADER,
      ['Ann', 'Ames', 'ann.s@school.org', 'cm.one@school.org'],
      ['Bob', 'Best', 'bob.s@school.org', 'CM.One@School.org'],
      ['Cal', 'Cole', 'cal.s@school.org', '  cm.one@school.org  ']
    ]
  });
  const res = sandbox.getCaseManagers('good-token');
  assertDeepEqual(emailsOf(res), ['cm.one@school.org'], 'case and whitespace variants are the same person');
});

test('the Student Directory is scanned too, and the union is returned', () => {
  setSheets({
    'Case Managers': [DIRECTORY_HEADER, ['Ann', 'Ames', 'ann.s@school.org', 'cm.one@school.org']],
    'Student Directory': [
      DIRECTORY_HEADER,
      ['Dee', 'Dunn', 'dee.s@school.org', 'cm.two@school.org'],
      ['Eli', 'Epps', 'eli.s@school.org', 'cm.one@school.org']
    ]
  });
  const res = sandbox.getCaseManagers('good-token');
  assertDeepEqual(emailsOf(res), ['cm.one@school.org', 'cm.two@school.org'], 'both sheets contribute, without duplicates');
});

test('a deployment with only a Student Directory still gets a roster', () => {
  setSheets({
    'Student Directory': [DIRECTORY_HEADER, ['Dee', 'Dunn', 'dee.s@school.org', 'cm.two@school.org']]
  });
  const res = sandbox.getCaseManagers('good-token');
  assertDeepEqual(emailsOf(res), ['cm.two@school.org'], 'Student Directory column D is the same field');
});

test('blank cells and stray labels are skipped', () => {
  setSheets({
    'Case Managers': [
      DIRECTORY_HEADER,
      ['Ann', 'Ames', 'ann.s@school.org', ''],
      ['Bob', 'Best', 'bob.s@school.org', 'none assigned'],
      ['Cal', 'Cole', 'cal.s@school.org', 'cm.one@school.org']
    ]
  });
  const res = sandbox.getCaseManagers('good-token');
  assertDeepEqual(emailsOf(res), ['cm.one@school.org'], 'only real addresses become picker entries');
});

test('names come from the Teachers sheet, otherwise from the email', () => {
  setSheets({
    'Case Managers': [
      DIRECTORY_HEADER,
      ['Ann', 'Ames', 'ann.s@school.org', 'jane.doe@school.org'],
      ['Bob', 'Best', 'bob.s@school.org', 'sam.roe@school.org']
    ],
    'Teachers': [TEACHERS_HEADER, ['Jane', 'Doe-Smith', 'jane.doe@school.org', '', 'Sp.Ed.', '']]
  });
  const res = sandbox.getCaseManagers('good-token');
  const byEmail = {};
  res.caseManagers.forEach(cm => { byEmail[cm.email] = cm.name; });
  assert(byEmail['jane.doe@school.org'] === 'Jane Doe-Smith', 'staff name should win');
  assert(byEmail['sam.roe@school.org'] === 'Sam Roe', 'non-staff falls back to a prettified local part');
});

test('results are sorted by display name', () => {
  setSheets({
    'Case Managers': [
      DIRECTORY_HEADER,
      ['A', 'A', 'a@school.org', 'zoe.zane@school.org'],
      ['B', 'B', 'b@school.org', 'amy.able@school.org'],
      ['C', 'C', 'c@school.org', 'max.moss@school.org']
    ]
  });
  const res = sandbox.getCaseManagers('good-token');
  assertDeepEqual(res.caseManagers.map(cm => cm.name), ['Amy Able', 'Max Moss', 'Zoe Zane'], 'alphabetical by name');
});

test('a missing sheet yields an empty roster, not an error', () => {
  setSheets({});
  const res = sandbox.getCaseManagers('good-token');
  assert(res.success, 'call should still succeed');
  assertDeepEqual(res.caseManagers, [], 'no sheets means no case managers');
});

test('a non-staff token is rejected', () => {
  setSheets({ 'Case Managers': [DIRECTORY_HEADER, ['A', 'A', 'a@school.org', 'cm.one@school.org']] });
  const res = sandbox.getCaseManagers('bad-token');
  assert(res.error === 'Unauthorized', 'the roster must not leak to non-staff');
  assert(!res.caseManagers, 'no data on an unauthorized call');
});

test('the cache is reused, and forceRefresh re-reads the sheets', () => {
  setSheets({ 'Case Managers': [DIRECTORY_HEADER, ['A', 'A', 'a@school.org', 'cm.one@school.org']] });
  assertDeepEqual(emailsOf(sandbox.getCaseManagers('good-token')), ['cm.one@school.org'], 'first read');

  // Swap the sheet contents without clearing the cache.
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ({
    getSheetByName: name => (name === 'Case Managers'
      ? { getDataRange: () => ({ getValues: () => [DIRECTORY_HEADER, ['B', 'B', 'b@school.org', 'cm.two@school.org']] }) }
      : null)
  });

  assertDeepEqual(emailsOf(sandbox.getCaseManagers('good-token')), ['cm.one@school.org'], 'cached roster is served');
  assertDeepEqual(emailsOf(sandbox.getCaseManagers('good-token', true)), ['cm.two@school.org'], 'forceRefresh bypasses the cache');
});

// ------------------------------------------------------------------ report --

console.log(`\n${passed} passed, ${failures.length} failed\n`);
process.exit(failures.length === 0 ? 0 : 1);

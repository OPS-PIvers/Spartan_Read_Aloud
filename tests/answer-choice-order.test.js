/**
 * Regression tests for answer-choice layout handling.
 *
 * Apps Script cannot be unit tested directly, so this loads Code.js into a
 * Node VM with the handful of Apps Script globals the pure text/HTML helpers
 * touch. Zero dependencies:
 *
 *   node tests/answer-choice-order.test.js
 *
 * Background: test generators such as ExamView lay multiple-choice options out
 * in a multi-column table filled COLUMN-major (a, b, c down the left column;
 * d, e down the right) but stored ROW-major, so every converter that walks the
 * document in reading order yields "a, d, b, e, c". ExamView also puts the
 * marker ("a.") in a cell of its own and prefixes each question with an answer
 * blank ("____ 1."), which used to hide the question number from the chunker.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');

// ---------------------------------------------------------------- sandbox --

const sandbox = {
  console,
  Logger: { log: () => {} },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  MimeType: { PDF: 'application/pdf', GOOGLE_DOCS: 'application/vnd.google-apps.document' },
  SpreadsheetApp: {}, DriveApp: {}, Drive: {}, UrlFetchApp: {},
  ScriptApp: {}, Utilities: {}, HtmlService: {}, CacheService: {},
  DocumentApp: {}, MailApp: {}, Session: {}
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'Constants.js'), 'utf8'), sandbox, { filename: 'Constants.js' });
vm.runInContext(fs.readFileSync(path.join(ROOT, 'Code.js'), 'utf8'), sandbox, { filename: 'Code.js' });

// --------------------------------------------------------------- fixtures --

/** Wraps cell content the way a Drive HTML export does. */
function cell(inner) {
  return `<td class="c2" colspan="1" rowspan="1"><p class="c1"><span class="c0">${inner}</span></p></td>`;
}
function tableRow(cells) { return `<tr class="c4">${cells.map(cell).join('')}</tr>`; }
function table(rows) { return `<table class="c9">${rows.map(tableRow).join('')}</table>`; }
function para(text) { return `<p class="c1"><span class="c0">${text}</span></p>`; }
function doc(body) { return `<html><body>${body}</body></html>`; }

/** Runs the real pipeline and returns the chunk texts. */
function chunkTexts(html) {
  return sandbox.parseHtmlToChunks(sandbox.sanitizeHtml(html)).map(c => c.text);
}

/**
 * Pulls the answer-option lines out of every chunk, in document order.
 * Lowercase markers only, so uppercase roman-numeral stems ("I. producer")
 * are not counted as options.
 */
function optionLines(html) {
  const lines = [];
  chunkTexts(html).forEach(text => {
    text.split('\n').forEach(line => {
      if (/^\(?[a-z][.)\]]\s/.test(line.trim())) lines.push(line.trim());
    });
  });
  return lines;
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

function assertEqual(actual, expected, message) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) throw new Error(`${message}\n      expected: ${e}\n      actual:   ${a}`);
}

// ------------------------------------------------------------------- tests --

console.log('\nAnswer-choice ordering\n');

test('5 options split a,b,c | d,e across two columns are read a-e', () => {
  const html = doc(
    para('____ 1. Approximately how many calories can be supported by this field?') +
    table([['a.', '357.4', 'd.', '35,740'], ['b.', '35.74', 'e.', '357,400'], ['c.', '3.5']])
  );
  assertEqual(optionLines(html),
    ['a. 357.4', 'b. 35.74', 'c. 3.5', 'd. 35,740', 'e. 357,400'],
    'options should be alphabetical, each joined to its own text');
});

test('4 options split a,b | c,d across two columns are read a-d', () => {
  const html = doc(
    para('____ 5. Which is reused because of rock erosion and sedimentation?') +
    table([['a.', 'nitrogen', 'c.', 'phosphorus'], ['b.', 'energy', 'd.', 'Carbon']])
  );
  assertEqual(optionLines(html),
    ['a. nitrogen', 'b. energy', 'c. phosphorus', 'd. Carbon'],
    'a 2x2 answer grid should linearize alphabetically');
});

test('a single-column table that is already ordered is left in order', () => {
  const html = doc(
    para('____ 6. Which statement is true?') +
    table([['a.', 'energy and nutrients are recycled'], ['b.', 'only energy is recycled'],
           ['c.', 'only nutrients are recycled'], ['d.', 'neither is recycled']])
  );
  assertEqual(optionLines(html),
    ['a. energy and nutrients are recycled', 'b. only energy is recycled',
     'c. only nutrients are recycled', 'd. neither is recycled'],
    'an already-ordered table must survive unchanged');
});

test('markers merged into the same cell as their text are handled', () => {
  const html = doc(
    para('____ 7. Which abiotic factor most influences succession?') +
    table([['a. amount of light', 'd. how windy the area is'],
           ['b. type of nutrients in soil', 'e. air temperatures'],
           ['c. height of plants']])
  );
  assertEqual(optionLines(html),
    ['a. amount of light', 'b. type of nutrients in soil', 'c. height of plants',
     'd. how windy the area is', 'e. air temperatures'],
    'options should reorder whether or not the marker has its own cell');
});

test('column-major order in flat paragraphs (no table) is corrected', () => {
  const html = doc(
    para('____ 2. A bird that has eaten an insect that fed on a plant is considered a') +
    para('a. producer.') + para('d. tertiary consumer.') + para('b. primary consumer.') +
    para('e. decomposer') + para('c. secondary consumer.')
  );
  assertEqual(optionLines(html),
    ['a. producer.', 'b. primary consumer.', 'c. secondary consumer.',
     'd. tertiary consumer.', 'e. decomposer'],
    'the flat-paragraph fallback should reorder too');
});

test('answer options carrying images keep their <img> and are reordered', () => {
  const html = doc(
    para('____ 9. Which diagram shows commensalism?') +
    table([['a.', '<img src="data:image/png;base64,AAAA" alt="A">',
            'c.', '<img src="data:image/png;base64,CCCC" alt="C">'],
           ['b.', '<img src="data:image/png;base64,BBBB" alt="B">',
            'd.', '<img src="data:image/png;base64,DDDD" alt="D">']])
  );
  const out = sandbox.sanitizeHtml(html);
  assertEqual((out.match(/<img/g) || []).length, 4, 'all four images must survive');
  const order = (out.match(/alt="[A-D]"/g) || []).map(s => s.slice(5, 6));
  assertEqual(order, ['A', 'B', 'C', 'D'], 'image options must end up alphabetical');
});

console.log('\nGuards against false positives\n');

test('a genuine data table is not flattened', () => {
  const html = doc(
    para('____ 12. Use the population data below.') +
    table([['Year', 'Deer', 'Wolves'], ['1995', '1,200', '14'],
           ['2000', '900', '31'], ['2005', '1,050', '22']])
  );
  const out = sandbox.sanitizeHtml(html);
  assert(/<table/i.test(out), 'a data table must remain a table');
  assert(out.indexOf('Wolves') < out.indexOf('1,200'), 'row order must be preserved');
});

test('roman-numeral stems are not mistaken for answer markers', () => {
  const html = doc(
    para('____ 3. At which trophic level(s) would you be considered?') +
    para('I. producer') + para('II. primary consumer') +
    para('III. secondary consumer') + para('IV. tertiary consumer') +
    table([['a.', 'IV only', 'd.', 'I and II'], ['b.', 'II only', 'e.', 'II and III'],
           ['c.', 'III only']])
  );
  const text = chunkTexts(html).join('\n');
  assert(text.indexOf('I. producer') < text.indexOf('II. primary consumer'),
    'roman numerals must keep their original order');
  assertEqual(optionLines(html),
    ['a. IV only', 'b. II only', 'c. III only', 'd. I and II', 'e. II and III'],
    'lettered options alongside roman numerals still reorder');
});

test('an incomplete letter set is left untouched', () => {
  const html = doc(para('Contact b. Jones or d. Smith or f. Adams for details.'));
  const out = sandbox.sanitizeHtml(html);
  assert(/b\. Jones/.test(out) && out.indexOf('b. Jones') < out.indexOf('d. Smith'),
    'letters that do not form an a,b,c... run must not be reordered');
});

console.log('\nChunk boundaries\n');

test('an answer blank before the number still starts a new chunk', () => {
  const html = doc(
    para('____ 1. First question?') + table([['a.', 'one', 'c.', 'three'], ['b.', 'two', 'd.', 'four']]) +
    para('____ 2. Second question?') + table([['a.', 'five', 'c.', 'seven'], ['b.', 'six', 'd.', 'eight']])
  );
  const chunks = chunkTexts(html);
  assertEqual(chunks.length, 2, 'each question needs its own audio chunk');
  assert(/First question/.test(chunks[0]) && /Second question/.test(chunks[1]),
    'questions must not bleed across chunks');
  assertEqual(chunks[0].split('\n').slice(1), ['a. one', 'b. two', 'c. three', 'd. four'],
    'first chunk carries its own options, in order');
});

console.log('\nSpoken answer markers\n');

test('every option a-e is announced, not just a-d', () => {
  const ssml = sandbox.addPausesToText('1. Pick one?\na. one\nb. two\nc. three\nd. four\ne. five');
  ['A', 'B', 'C', 'D', 'E'].forEach(letter => {
    assert(ssml.indexOf(`<say-as interpret-as="characters">${letter}</say-as>`) !== -1,
      `option ${letter} should be spelled out`);
  });
});

test('prose abbreviations are not read as answer choices', () => {
  const ssml = sandbox.addPausesToText('The test begins at 9 a.m. sharp, e.g. in room 4.');
  assert(ssml.indexOf('say-as') === -1, '"9 a.m." and "e.g." must be left alone');
  assert(ssml.indexOf('9 a.m. sharp, e.g. in room 4.') !== -1, 'the sentence must survive intact');
});

test('parenthesised markers are announced', () => {
  const ssml = sandbox.addPausesToText('1. Pick one?\n(a) yes\n(b) no\n(c) maybe');
  assert(/\(<say-as interpret-as="characters">A<\/say-as>\)/.test(ssml),
    'a "(a)" marker should keep its parentheses and be spelled out');
});

// ------------------------------------------------------------------ report --

console.log(`\n${passed} passed, ${failures.length} failed\n`);
process.exit(failures.length === 0 ? 0 : 1);

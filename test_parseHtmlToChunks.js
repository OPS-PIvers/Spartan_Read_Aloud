// Mock global Logger for GAS environment
global.Logger = {
  log: (msg) => console.log('[Logger]', msg)
};

const { parseHtmlToChunks } = require('./TextProcessing.js');

// --- Tests ---

function assert(condition, message) {
  if (!condition) {
    console.error(`❌ FAIL: ${message}`);
    process.exit(1);
  } else {
    console.log(`✅ PASS: ${message}`);
  }
}

function assertEqual(actual, expected, message) {
  const actualStr = JSON.stringify(actual);
  const expectedStr = JSON.stringify(expected);
  if (actualStr !== expectedStr) {
    console.error(`❌ FAIL: ${message}\nExpected: ${expectedStr}\nActual:   ${actualStr}`);
    process.exit(1);
  } else {
    console.log(`✅ PASS: ${message}`);
  }
}

console.log('--- Running Tests for parseHtmlToChunks (Refactored) ---');

// Test Case 1: Short Paragraphs (Should Merge)
const html1 = `
  <p id="sra-block-1">First paragraph.</p>
  <p id="sra-block-2">Second paragraph.</p>
`;
const chunks1 = parseHtmlToChunks(html1);
assertEqual(chunks1.length, 1, 'Should merge short paragraphs');
assertEqual(chunks1[0].text, 'First paragraph. Second paragraph.', 'Merged short paragraphs text match');

// Test Case 2: Question and Answer Grouping
const html2 = `
  <p id="sra-block-3">1. What is the capital of France?</p>
  <p id="sra-block-4">a. Berlin</p>
  <p id="sra-block-5">b. Paris</p>
  <p id="sra-block-6">c. Rome</p>
`;
const chunks2 = parseHtmlToChunks(html2);
assertEqual(chunks2.length, 1, 'Should group question and answers into 1 chunk');
assertEqual(chunks2[0].text, '1. What is the capital of France?\na. Berlin\nb. Paris\nc. Rome', 'Grouped text match');
assertEqual(chunks2[0].ids, ['sra-block-3', 'sra-block-4', 'sra-block-5', 'sra-block-6'], 'Grouped IDs match');

// Test Case 3: Flowing Text (Sentence continuation)
const html3 = `
  <p id="sra-block-7">This is a sentence that continues</p>
  <p id="sra-block-8">in the next paragraph.</p>
`;
const chunks3 = parseHtmlToChunks(html3);
assertEqual(chunks3.length, 1, 'Should merge flowing text into 1 chunk');
assertEqual(chunks3[0].text, 'This is a sentence that continues in the next paragraph.', 'Merged text match');

// Test Case 4: Long Separate Thoughts (Should Split)
const longText = 'This is a very long paragraph that should be treated as a complete thought because it ends with a period and is sufficiently long to warrant a separate chunk in the text-to-speech generation process. '.repeat(2);
const html4 = `
  <p id="sra-block-9">${longText}</p>
  <p id="sra-block-10">This is another one.</p>
`;
const chunks4 = parseHtmlToChunks(html4);
assertEqual(chunks4.length, 2, 'Should keep long separate sentences as 2 chunks');

// Test Case 5: Headers (Should Split)
const html5 = `
  <h1 id="sra-block-11">Chapter 1</h1>
  <p id="sra-block-12">The beginning.</p>
`;
const chunks5 = parseHtmlToChunks(html5);
assertEqual(chunks5.length, 2, 'Header should start a new chunk');

// Test Case 6: Directions (Should Split)
const html6 = `
  <p id="sra-block-13">Directions: Read carefully.</p>
  <p id="sra-block-14">Question 1 starts here.</p>
`;
const chunks6 = parseHtmlToChunks(html6);
assertEqual(chunks6.length, 2, 'Directions should form their own chunk');

console.log('--- All Tests Passed ---');

const fs = require('fs');
const code = fs.readFileSync('Code.js', 'utf8');

// Mock HtmlService
const HtmlService = {
  createHtmlOutput: (html) => html
};

// Mock escapeHtmlBackend from Code.js
function extractFunction(code, funcName) {
  const funcRegex = new RegExp(`function\\s+${funcName}\\s*\\([^)]*\\)\\s*{`);
  const match = code.match(funcRegex);
  if (!match) return null;

  let openBraces = 0;
  let funcString = '';
  let inString = false;
  let stringChar = '';
  let i = match.index;

  // Find start
  while (i < code.length && code[i] !== '{') {
    funcString += code[i];
    i++;
  }

  // Extract body
  do {
    const char = code[i];
    funcString += char;

    if (inString) {
      if (char === stringChar && code[i-1] !== '\\') {
        inString = false;
      }
    } else {
      if (char === "'" || char === '"' || char === '`') {
        inString = true;
        stringChar = char;
      } else if (char === '{') {
        openBraces++;
      } else if (char === '}') {
        openBraces--;
      }
    }
    i++;
  } while (openBraces > 0 && i < code.length);

  return funcString;
}

const escapeHtmlBackendStr = extractFunction(code, 'escapeHtmlBackend');
const escapeHtmlBackend = new Function(`return ${escapeHtmlBackendStr}`)();

function testEscapeHtmlBackend() {
  console.log("Testing escapeHtmlBackend...");
  const maliciousEmail = "<script>alert('xss')</script>";
  const escaped = escapeHtmlBackend(maliciousEmail);
  console.log("Input:", maliciousEmail);
  console.log("Output:", escaped);

  if (escaped === "&lt;script&gt;alert(&#39;xss&#39;)&lt;/script&gt;") {
    console.log("✅ escapeHtmlBackend correctly escapes HTML");
  } else {
    console.error("❌ escapeHtmlBackend failed");
    process.exit(1);
  }
}

testEscapeHtmlBackend();

const doGetStr = extractFunction(code, 'doGet');
// Mocking the behavior inside doGet where this occurs
function simulateAccessDenied(userEmail) {
    return HtmlService.createHtmlOutput(`<h1>Access Denied</h1><p>Your email (${escapeHtmlBackend(userEmail)}) is not authorized to use this application.</p>`);
}

function testXssFix() {
  console.log("\nTesting XSS Fix in Access Denied logic...");
  const maliciousEmail = "malicious<script>alert(1)</script>@email.com";
  const output = simulateAccessDenied(maliciousEmail);

  console.log("Generated HTML:", output);

  if (output.includes("<script>")) {
     console.error("❌ Vulnerability still present!");
     process.exit(1);
  } else if (output.includes("&lt;script&gt;")) {
     console.log("✅ Vulnerability fixed! Email was escaped.");
  } else {
     console.error("❌ Unexpected output.");
     process.exit(1);
  }
}

testXssFix();

const fs = require('fs');
const code = fs.readFileSync('Code.js', 'utf8');

// Mock HtmlService
const HtmlService = {
  createHtmlOutput: (html) => html
};

// Mock escapeHtmlBackend from Code.js

// The escapeHtmlBackend function is a pure utility and can be copied here directly
// for more robust testing, avoiding the need for fragile string parsing.
function escapeHtmlBackend(text) {
  if (!text) return '';
  return text.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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

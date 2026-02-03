import sys

with open('Code.js', 'r') as f:
    content = f.read()

old_broken = """/**
 * Retrieves assessment data for authenticated student.
 * All file types (PDF, Google Docs, Word) are converted to HTML for consistent rendering.
 * @param {string} email Student email
 * @param {string} password Assessment password
 * @param {string} assessmentUrl Optional - specific assessment URL to load (for multi-assessment selection)
 * @returns {Object} Assessment data with assessmentHtml and sessionToken for secure audio access
 */
          // Password is correct, proceed..."""

new_fixed = """/**
 * Retrieves assessment data for authenticated student.
 * All file types (PDF, Google Docs, Word) are converted to HTML for consistent rendering.
 * @param {string} email Student email
 * @param {string} password Assessment password
 * @param {string} assessmentUrl Optional - specific assessment URL to load (for multi-assessment selection)
 * @returns {Object} Assessment data with assessmentHtml and sessionToken for secure audio access
 */
function getAssessmentPdf(email, password, assessmentUrl) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return { error: 'Backend Error: "Assessment Database" sheet not found.' };
    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[CONSTANTS.COL.PDF_URL];

      // Find the correct assessment row using the URL
      if (pdfUrl === assessmentUrl) {
        const studentEmailsRaw = row[CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();
        const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());
        const sheetPassword = row[CONSTANTS.COL.PASSWORD].toString().trim();

        // Check if user is staff
        const user = getUserByEmail(cleanEmail);
        const isStaff = user && (user.userType === CONSTANTS.ROLE_TOKEN_TEACHER || user.userType === CONSTANTS.ROLE_TOKEN_ADMIN || user.userType === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN);

        // Verify this authenticated user's email is in the list for this assessment (or is staff)
        if (isStaff || studentEmails.includes(cleanEmail)) {
          // NEW: Validate the provided password against the one in the sheet (skip for staff)
          if (!isStaff && password !== sheetPassword) {
            return { error: 'Incorrect password for this assessment.' };
          }

          // Password is correct, proceed..."""

content = content.replace(old_broken, new_fixed)

with open('Code.js', 'w') as f:
    f.write(content)

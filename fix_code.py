import sys

with open('Code.js', 'r') as f:
    content = f.read()

# Fix step1_AnalyzePdfsAndCountChunks
old_step1 = """  for (let i = 1; i < data.length; i++) {
    const pdfUrl = data[i][CONSTANTS.COL.PDF_URL];
    if (pdfUrl && !chunkCount) {
      const readAloudEnabled = data[i][CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
      if (!readAloudEnabled) {
        // Just mark as complete with 0 chunks
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(0);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
        sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue("NO_AUDIO_REQUIRED");
        continue;
      }

      const fileId = getFileIdFromUrl(pdfUrl);
    if (pdfUrl && !chunkCount) {
      const fileId = getFileIdFromUrl(pdfUrl);
      if (!fileId) {
        Logger.log(`Invalid Drive URL in row ${i + 1}. Skipping.`);
        continue;
      }"""

new_step1 = """  for (let i = 1; i < data.length; i++) {
    const pdfUrl = data[i][CONSTANTS.COL.PDF_URL];
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];

    if (pdfUrl && !chunkCount) {
      const readAloudEnabled = data[i][CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
      if (!readAloudEnabled) {
        // Just mark as complete with 0 chunks
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(0);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
        sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue("NO_AUDIO_REQUIRED");
        continue;
      }

      const fileId = getFileIdFromUrl(pdfUrl);
      if (!fileId) {
        Logger.log(`Invalid Drive URL in row ${i + 1}. Skipping.`);
        continue;
      }"""

content = content.replace(old_step1, new_step1)

# Fix getAllAssessments
old_getall = """      } catch (e) {
        fileName = 'Unknown file';
        password: row[CONSTANTS.COL.PASSWORD] || "",
        studentEmails: row[CONSTANTS.COL.STUDENT_EMAILS] || "",
        readAloudEnabled: row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false
      });

      assessments.push({
        rowIndex: i,
        fileName: fileName,
        pdfUrl: pdfUrl,
        chunkCount: row[CONSTANTS.COL.CHUNK_COUNT] || 0,"""

new_getall = """      } catch (e) {
        fileName = 'Unknown file';
      }

      assessments.push({
        rowIndex: i,
        fileName: fileName,
        pdfUrl: pdfUrl,
        chunkCount: row[CONSTANTS.COL.CHUNK_COUNT] || 0,
        audioJson: row[CONSTANTS.COL.AUDIO_JSON] || '',
        isComplete: row[CONSTANTS.COL.IS_COMPLETE] === true,
        className: row[CONSTANTS.COL.CLASS_NAME] || '',
        instructor: row[CONSTANTS.COL.INSTRUCTOR] || '',
        password: row[CONSTANTS.COL.PASSWORD] || '',
        studentEmails: row[CONSTANTS.COL.STUDENT_EMAILS] || '',
        readAloudEnabled: row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false
      });"""

content = content.replace(old_getall, new_getall)

# Fix updateAssessmentRow
old_update = """    if (data.studentEmails !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.STUDENT_EMAILS + 1).setValue(parseStudentEmails(data.studentEmails));
    }
    if (data.readAloudEnabled !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.READ_ALOUD_ENABLED + 1).setValue(data.readAloudEnabled);
    }

    SpreadsheetApp.flush();
    if (data.password !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.PASSWORD + 1).setValue(data.password);
    }
    if (data.studentEmails !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.STUDENT_EMAILS + 1).setValue(parseStudentEmails(data.studentEmails));
    }

    SpreadsheetApp.flush();"""

new_update = """    if (data.password !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.PASSWORD + 1).setValue(data.password);
    }
    if (data.studentEmails !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.STUDENT_EMAILS + 1).setValue(parseStudentEmails(data.studentEmails));
    }
    if (data.readAloudEnabled !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.READ_ALOUD_ENABLED + 1).setValue(data.readAloudEnabled);
    }

    SpreadsheetApp.flush();"""

content = content.replace(old_update, new_update)

# Fix addNewAssessment
old_addnew = """    newRow[CONSTANTS.COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || "");
    newRow[CONSTANTS.COL.READ_ALOUD_ENABLED] = metadata.readAloudEnabled !== false;
      if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl) {
        return { error: 'This file is already in the database.' };
      }
    }

    // Add new row with file URL and metadata
    const newRow = new Array(8).fill(''); // 8 columns
    newRow[CONSTANTS.COL.PDF_URL] = fileUrl;
    newRow[CONSTANTS.COL.CLASS_NAME] = metadata.className || '';
    newRow[CONSTANTS.COL.INSTRUCTOR] = metadata.instructor || '';
    newRow[CONSTANTS.COL.PASSWORD] = metadata.password || '';
    newRow[CONSTANTS.COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || '');

    sheet.appendRow(newRow);"""

new_addnew = """      if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl) {
        return { error: 'This file is already in the database.' };
      }
    }

    // Add new row with file URL and metadata
    const newRow = new Array(13).fill(''); // 13 columns
    newRow[CONSTANTS.COL.PDF_URL] = fileUrl;
    newRow[CONSTANTS.COL.CLASS_NAME] = metadata.className || '';
    newRow[CONSTANTS.COL.INSTRUCTOR] = metadata.instructor || '';
    newRow[CONSTANTS.COL.PASSWORD] = metadata.password || '';
    newRow[CONSTANTS.COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || '');
    newRow[CONSTANTS.COL.READ_ALOUD_ENABLED] = metadata.readAloudEnabled !== false;

    sheet.appendRow(newRow);"""

content = content.replace(old_addnew, new_addnew)

# Fix getAssessmentPdf (this one is the worst)
# I'll just replace the whole function if possible, or large chunks of it.
import re
# Find the start of the function and end
# Note: This is fragile if there are multiple functions with same name but let's hope not.
# Actually, I'll just use a simpler replacement for the broken part.

old_getpdf_broken = """          // Password is correct, proceed...
          const audioDataJson = row[CONSTANTS.COL.AUDIO_JSON];
          const readAloudEnabled = row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;

          if (readAloudEnabled && !audioDataJson) {
            return { error: "Audio for this assessment has not been generated yet. Please try again later." };
          }

          const fileId = getFileIdFromUrl(pdfUrl);
          if (!fileId) return { error: "Invalid Google Drive URL in sheet." };

          const file = DriveApp.getFileById(fileId);
          const mimeType = file.getMimeType();
          const fileName = file.getName();
          const audioChunks = audioDataJson ? JSON.parse(audioDataJson) : [];
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
          return {
            fileType: "html",
            assessmentHtml: sanitizeHtml(conversionResult.html),
            fileName: fileName,
            audioChunks: audioChunks,
            sessionToken: sessionToken,
            readAloudEnabled: readAloudEnabled
          };
            return { error: 'Audio for this assessment has not been generated yet. Please try again later.' };
          }

          const fileId = getFileIdFromUrl(pdfUrl);
          if (!fileId) return { error: 'Invalid Google Drive URL in sheet.' };

          const file = DriveApp.getFileById(fileId);
          const mimeType = file.getMimeType();
          const fileName = file.getName();
          const audioChunks = JSON.parse(audioDataJson);

          Logger.log(`Serving assessment: ${fileName} (${mimeType}) to ${email}`);

          // Generate session token for secure audio access
          // If staff, use their role to bypass list validation. If student, use pdfUrl to bind token to assessment.
          const sessionToken = isStaff ?
            generateSessionToken(cleanEmail, user.userType) :
            generateSessionToken(cleanEmail, pdfUrl);

          // Convert all files (PDF, Docs, Word) to HTML with embedded images
          Logger.log('→ Converting to HTML for native rendering');
          const conversionResult = convertFileToHtml(fileId);
          if (conversionResult.error) {
            Logger.log(`✗ Conversion error: ${conversionResult.error}`);
            return { error: `Could not load assessment: ${conversionResult.error}` };
          }

          return {
            fileType: 'html',
            assessmentHtml: sanitizeHtml(conversionResult.html),
            fileName: fileName,
            audioChunks: audioChunks,
            sessionToken: sessionToken
          };"""

new_getpdf_fixed = """          // Password is correct, proceed...
          const audioDataJson = row[CONSTANTS.COL.AUDIO_JSON];
          const readAloudEnabled = row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;

          if (readAloudEnabled && !audioDataJson) {
            return { error: "Audio for this assessment has not been generated yet. Please try again later." };
          }

          const fileId = getFileIdFromUrl(pdfUrl);
          if (!fileId) return { error: "Invalid Google Drive URL in sheet." };

          const file = DriveApp.getFileById(fileId);
          const mimeType = file.getMimeType();
          const fileName = file.getName();
          const audioChunks = audioDataJson ? JSON.parse(audioDataJson) : [];

          Logger.log(`Serving assessment: ${fileName} (${mimeType}) to ${email}`);

          // Generate session token for secure audio access
          const sessionToken = isStaff ?
            generateSessionToken(cleanEmail, user.userType) :
            generateSessionToken(cleanEmail, pdfUrl);

          // Convert all files to HTML
          Logger.log('→ Converting to HTML for native rendering');
          const conversionResult = convertFileToHtml(fileId);
          if (conversionResult.error) {
            Logger.log(`✗ Conversion error: ${conversionResult.error}`);
            return { error: `Could not load assessment: ${conversionResult.error}` };
          }

          return {
            fileType: "html",
            assessmentHtml: sanitizeHtml(conversionResult.html),
            fileName: fileName,
            audioChunks: audioChunks,
            sessionToken: sessionToken,
            readAloudEnabled: readAloudEnabled
          };"""

# Use re.escape and sub for more robustness with large blocks?
# Or just hope string replace works if it's exact.
content = content.replace(old_getpdf_broken, new_getpdf_fixed)

# Fix processNewAssessment
old_process_broken = """        // Skip audio generation, mark as complete immediately
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(0);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
        sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue("NO_AUDIO_REQUIRED");
        SpreadsheetApp.flush();
        Logger.log("Read Aloud disabled. Marked as complete.");
        break;
      }
 * Internal function called after adding new assessment.
 * @param {string} fileUrl File URL to process
 */
function processNewAssessment(fileUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl && !data[i][CONSTANTS.COL.CHUNK_COUNT]) {
      // Found the new row - analyze it
      const fileId = getFileIdFromUrl(fileUrl);
      if (!fileId) continue;

      Logger.log(`Processing new assessment: ${fileUrl}`);"""

new_process_fixed = """/**
 * Processes a specific assessment (runs steps 1 and 2).
 * Internal function called after adding new assessment.
 * @param {string} fileUrl File URL to process
 */
function processNewAssessment(fileUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl && !data[i][CONSTANTS.COL.CHUNK_COUNT]) {
      // Found the new row - analyze it
      const fileId = getFileIdFromUrl(fileUrl);
      if (!fileId) continue;

      const readAloudEnabled = data[i][CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
      Logger.log(`Processing new assessment: ${fileUrl} (Read Aloud: ${readAloudEnabled})`);

      if (!readAloudEnabled) {
        // Skip audio generation, mark as complete immediately
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(0);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
        sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue("NO_AUDIO_REQUIRED");
        SpreadsheetApp.flush();
        Logger.log("Read Aloud disabled. Marked as complete.");
        break;
      }"""

content = content.replace(old_process_broken, new_process_fixed)

with open('Code.js', 'w') as f:
    f.write(content)

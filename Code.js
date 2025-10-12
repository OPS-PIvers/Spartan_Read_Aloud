/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const AUDIO_DRIVE_FOLDER_NAME = "Assessment Audio Files"; 

// --- NEW SPREADSHEET COLUMN MAPPING ---
// A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7
const COL = {
  PDF_URL: 0,
  CHUNK_COUNT: 1,
  AUDIO_JSON: 2,
  IS_COMPLETE: 3,
  CLASS_NAME: 4,
  INSTRUCTOR: 5,
  PASSWORD: 6,
  STUDENT_EMAILS: 7
};

// --- SUPPORTED FILE FORMATS ---
const SUPPORTED_MIME_TYPES = {
  PDF: MimeType.PDF,
  GOOGLE_DOCS: MimeType.GOOGLE_DOCS,
  MS_WORD: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  MS_WORD_OLD: 'application/msword'
};

// --- TRIGGER & MENU --- 

/**
 * Adds a custom menu to the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Spartan Read Aloud')
      .addItem('Run All Steps', 'runAllSteps')
      .addToUi();
}

/**
 * Runs all the processing steps in sequence.
 */
function runAllSteps() {
  step0_addNewPdfs();
  step1_AnalyzePdfsAndCountChunks();
  step2_GenerateMissingAudioAndFinalize();
}

// --- MAIN CONTROL FUNCTIONS ---

/**
 * STEP 0: Finds new PDFs in the designated Drive folder and adds them to the sheet.
 */
function step0_addNewPdfs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) {
    Logger.log('ERROR: "Assessment Database" sheet not found.');
    return;
  }

  const mainAudioFolder = getOrCreateFolder(AUDIO_DRIVE_FOLDER_NAME);
  if (!mainAudioFolder) return;

  const pdfSourceFolderName = "Assessment PDFs";
  const pdfFolders = mainAudioFolder.getFoldersByName(pdfSourceFolderName);
  if (!pdfFolders.hasNext()) {
    Logger.log(`ERROR: Source folder "${pdfSourceFolderName}" not found inside "${AUDIO_DRIVE_FOLDER_NAME}".`);
    return;
  }
  const pdfFolder = pdfFolders.next();

  // Get existing URLs to prevent duplicates
  const data = sheet.getDataRange().getValues();
  const existingUrls = new Set(data.map(row => row[COL.PDF_URL]));

  const allFiles = pdfFolder.getFiles();
  let addedCount = 0;

  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const mimeType = file.getMimeType();

    // Check if supported format
    const isSupported = Object.values(SUPPORTED_MIME_TYPES).includes(mimeType);

    if (isSupported) {
      const fileUrl = file.getUrl();
      if (!existingUrls.has(fileUrl)) {
        sheet.appendRow([fileUrl]);
        Logger.log(`Added new file: ${file.getName()} (${mimeType})`);
        addedCount++;
      }
    }
  }

  if (addedCount > 0) {
    SpreadsheetApp.flush();
    Logger.log(`Step 0 finished. Added ${addedCount} new PDFs.`);
  } else {
    Logger.log('Step 0 finished. No new PDFs found.');
  }
}


/**
 * STEP 1: Analyzes new PDFs to count their text chunks.
 */
function step1_AnalyzePdfsAndCountChunks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) {
    Logger.log('ERROR: "Assessment Database" sheet not found.');
    return;
  }
  const data = sheet.getDataRange().getValues();
  Logger.log('Starting Step 1: Analyzing new PDFs...');

  for (let i = 1; i < data.length; i++) {
    const pdfUrl = data[i][COL.PDF_URL];
    const chunkCount = data[i][COL.CHUNK_COUNT];

    if (pdfUrl && !chunkCount) {
      const fileId = getFileIdFromUrl(pdfUrl);
      if (!fileId) {
        Logger.log(`Invalid Drive URL in row ${i + 1}. Skipping.`);
        continue;
      }
      const fileName = DriveApp.getFileById(fileId).getName();
      Logger.log(`-> Analyzing '${fileName}'...`);

      const textChunks = extractTextFromFile(fileId);
      if (textChunks && textChunks.length > 0) {
        sheet.getRange(i + 1, COL.CHUNK_COUNT + 1).setValue(textChunks.length);
        sheet.getRange(i + 1, COL.IS_COMPLETE + 1).setValue(false);
        Logger.log(`--> Found ${textChunks.length} chunks. Updated sheet.`);
      } else {
        Logger.log(`--> No text chunks found for '${fileName}'.`);
      }
    }
  }
  SpreadsheetApp.flush();
  Logger.log('Step 1 Analysis finished.');
}


/**
 * STEP 2: Generates missing audio files and finalizes the JSON data.
 * Now uses descriptive filenames based on the chunk's text.
 */
function step2_GenerateMissingAudioAndFinalize() {
  const SCRIPT_START_TIME = new Date();
  const SCRIPT_TIMEOUT_MS = 5 * 60 * 1000;

  const mainAudioFolder = getOrCreateFolder(AUDIO_DRIVE_FOLDER_NAME);
  if (!mainAudioFolder) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  Logger.log('Starting Step 2: Generating missing audio...');

  for (let i = 1; i < data.length; i++) {
    const elapsedTime = new Date() - SCRIPT_START_TIME;
    if (elapsedTime > SCRIPT_TIMEOUT_MS) {
      Logger.log(`Approaching 6-minute execution limit. Stopping gracefully.`);
      break;
    }

    const isComplete = data[i][COL.IS_COMPLETE];
    const pdfUrl = data[i][COL.PDF_URL];
    const totalChunks = data[i][COL.CHUNK_COUNT];

    if (pdfUrl && totalChunks > 0 && !isComplete) {
      const fileId = getFileIdFromUrl(pdfUrl);
      if (!fileId) continue;

      const file = DriveApp.getFileById(fileId);
      const fileName = file.getName(); 
      
      const baseName = fileName.replace(/\.pdf$/i, '').trim();
      const assessmentSubfolder = getOrCreateSubfolder(mainAudioFolder, baseName);
      if (!assessmentSubfolder) continue;

      Logger.log(`Processing '${fileName}' (Row ${i + 1}). Total chunks: ${totalChunks}`);

      const textChunks = extractTextFromFile(fileId);
      if (!textChunks || textChunks.length !== totalChunks) {
          Logger.log(`--> ERROR: Mismatch in chunk count for '${fileName}'. Expected ${totalChunks}, found ${textChunks ? textChunks.length : 0}. Skipping.`);
          continue;
      }
      
      const audioFileObjects = [];
      let allChunksProcessed = true;

      for (let j = 0; j < totalChunks; j++) {
         const chunkText = textChunks[j];
         // --- NEW: Generate the descriptive filename ---
         const newChunkName = generateSafeFilenameFromText(chunkText, j);

         // Define legacy names for backwards compatibility
         const cleanLegacyName = `${baseName}-chunk-${j + 1}.wav`;
         const legacyFullName = `${fileName}-chunk-${j + 1}.wav`;

         // Check for new name, then the two old formats
         let existingFiles = assessmentSubfolder.getFilesByName(newChunkName);
         if (!existingFiles.hasNext()) {
            existingFiles = assessmentSubfolder.getFilesByName(cleanLegacyName);
         }
         if (!existingFiles.hasNext()) {
            existingFiles = assessmentSubfolder.getFilesByName(legacyFullName);
         }

         let audioFile = null;

         if (existingFiles.hasNext()) {
            audioFile = existingFiles.next();
         } else {
            Logger.log(`--> Generating new audio for chunk ${j + 1} with name "${newChunkName}"...`);
            // Always generate new files with the new descriptive name
            audioFile = generateAudioFromTextChunk(chunkText, newChunkName, assessmentSubfolder);
         }

         if (audioFile) {
            audioFileObjects.push(audioFile);
         } else {
            Logger.log(`--> FAILED to process chunk ${j + 1}. Will retry on next run.`);
            allChunksProcessed = false;
            break; 
         }
      }
      
      if (allChunksProcessed && audioFileObjects.length === totalChunks) {
        Logger.log(`--> All ${totalChunks} audio chunks accounted for. Finalizing...`);
        const audioDataForSheet = [];
        for(let j = 0; j < totalChunks; j++) {
           const chunkText = textChunks[j];
           const audioFile = audioFileObjects[j];
           // Generate searchWords: first 8 words to match frontend display
           const words = chunkText.trim().split(/\s+/);
           const searchWords = words.slice(0, 8).join(' ') + (words.length > 8 ? '...' : '');

           audioDataForSheet.push({
             text: chunkText,
             searchWords: searchWords,
             audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
             audioFilename: audioFile.getName()
           });
        }
        sheet.getRange(i + 1, COL.AUDIO_JSON + 1).setValue(JSON.stringify(audioDataForSheet, null, 2));
        sheet.getRange(i + 1, COL.IS_COMPLETE + 1).setValue(true);
        Logger.log(`--> Successfully created JSON and marked as complete.`);
      } else {
        Logger.log(`--> Process for '${fileName}' partially complete. Will resume on next run.`);
      }
    }
  }
  SpreadsheetApp.flush();
  Logger.log('Step 2 processing finished.');
}


// --- HELPER FUNCTIONS ---

/**
 * NEW: Creates a safe, descriptive filename from the first few words of a text chunk.
 * @param {string} text The text of the chunk.
 * @param {number} chunkIndex The zero-based index of the chunk.
 * @returns {string} A sanitized, unique filename.
 */
function generateSafeFilenameFromText(text, chunkIndex) {
  // Get first 6 words
  const firstWords = text.split(/\s+/).slice(0, 6).join(' ');
  // Sanitize: remove non-alphanumerics (but keep hyphens), and replace spaces with hyphens
  const sanitized = firstWords.replace(/[^\w\s-]/g, '').replace(/\s+/g, '-');
  // Add chunk index for uniqueness and the extension, ensuring it's not too long
  const fullName = `${sanitized}-chunk-${chunkIndex + 1}.wav`;
  return fullName.substring(0, 250); // Trim to a safe length
}

/**
 * Converts a file to HTML format based on its MIME type.
 * Handles Google Docs, Word docs, and PDFs (via OCR).
 * @param {string} fileId The Drive file ID
 * @returns {Object} { html: string, mimeType: string, fileName: string } or { error: string }
 */
function convertFileToHtml(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    const fileSize = file.getSize();

    Logger.log(`Converting file: ${file.getName()} (${mimeType}, ${fileSize} bytes)`);

    // Check file size (warn if >10MB, reject if >45MB to stay under GAS limits)
    if (fileSize > 45 * 1024 * 1024) {
      Logger.log(`✗ File too large: ${fileSize} bytes`);
      return { error: 'File too large (>45MB). Please use a smaller file.' };
    }

    if (fileSize > 10 * 1024 * 1024) {
      Logger.log(`⚠ Large file warning: ${fileSize} bytes - may be slow`);
    }

    let htmlContent = null;

    // Handle Google Docs
    if (mimeType === SUPPORTED_MIME_TYPES.GOOGLE_DOCS) {
      Logger.log('→ Converting Google Doc to HTML via Drive API export');
      htmlContent = exportDocToHtml(fileId);
    }
    // Handle Microsoft Word (.docx and .doc)
    else if (mimeType === SUPPORTED_MIME_TYPES.MS_WORD || mimeType === SUPPORTED_MIME_TYPES.MS_WORD_OLD) {
      Logger.log('→ Converting Word doc: First to Google Doc, then to HTML');
      htmlContent = convertWordToHtml(fileId, file);
    }
    // Handle PDFs
    else if (mimeType === SUPPORTED_MIME_TYPES.PDF) {
      Logger.log('→ Converting PDF: OCR to Google Doc, then to HTML');
      htmlContent = convertPdfToHtml(fileId, file);
    }
    // Unsupported format
    else {
      Logger.log(`✗ Unsupported MIME type: ${mimeType}`);
      return { error: `Unsupported file format: ${mimeType}` };
    }

    if (!htmlContent) {
      Logger.log('✗ HTML conversion returned null');
      return { error: 'Failed to convert file to HTML' };
    }

    Logger.log(`✓ Successfully converted to HTML (${htmlContent.length} chars)`);
    return {
      html: htmlContent,
      mimeType: mimeType,
      fileName: file.getName()
    };

  } catch (e) {
    Logger.log(`✗ Error in convertFileToHtml: ${e.toString()}`);
    return { error: `Conversion failed: ${e.toString()}` };
  }
}

/**
 * Exports a Google Doc to HTML using Drive API v3 REST endpoint.
 * @param {string} fileId The Google Doc file ID
 * @returns {string|null} HTML content or null on failure
 */
function exportDocToHtml(fileId) {
  try {
    const token = ScriptApp.getOAuthToken();
    const url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/html`;

    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    };

    Logger.log(`→ Calling Drive API v3 export: ${url}`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      Logger.log('✓ Drive API export successful');
      return response.getContentText();
    }

    Logger.log(`✗ Drive API export failed: ${responseCode}`);
    Logger.log(`Response: ${response.getContentText()}`);
    return null;

  } catch (e) {
    Logger.log(`✗ Exception in exportDocToHtml: ${e.toString()}`);
    return null;
  }
}

/**
 * Converts Word file to HTML by first converting to Google Doc, then exporting.
 * @param {string} fileId The Word file ID
 * @param {GoogleAppsScript.Drive.File} file The file object
 * @returns {string|null} HTML content or null on failure
 */
function convertWordToHtml(fileId, file) {
  let tempDocId = null;
  try {
    const blob = file.getBlob();
    const resource = {
      title: blob.getName(),
      mimeType: MimeType.GOOGLE_DOCS // Convert to Google Doc
    };

    Logger.log('→ Converting Word file to temporary Google Doc');
    // Use Drive API v2 to convert Word → Google Doc
    const tempDoc = Drive.Files.insert(resource, blob);
    tempDocId = tempDoc.id;
    Logger.log(`→ Created temporary Google Doc: ${tempDocId}`);

    // Export the Google Doc to HTML using Drive API v3
    const html = exportDocToHtml(tempDocId);

    if (!html) {
      Logger.log('✗ Failed to export temporary doc to HTML');
      return null;
    }

    Logger.log('✓ Word file successfully converted to HTML');
    return html;

  } catch (e) {
    Logger.log(`✗ Word conversion failed: ${e.toString()}`);
    return null;
  } finally {
    // Always clean up temporary doc
    if (tempDocId) {
      try {
        Drive.Files.remove(tempDocId);
        Logger.log('→ Deleted temporary Google Doc');
      } catch (cleanupError) {
        Logger.log(`⚠ Failed to delete temp doc ${tempDocId}: ${cleanupError.toString()}`);
      }
    }
  }
}

/**
 * Converts PDF to HTML using OCR (existing logic) + HTML export.
 * @param {string} fileId The PDF file ID
 * @param {GoogleAppsScript.Drive.File} file The file object
 * @returns {string|null} HTML content or null on failure
 */
function convertPdfToHtml(fileId, file) {
  let tempDocId = null;
  try {
    const blob = file.getBlob();
    const resource = {
      title: blob.getName(),
      mimeType: blob.getContentType()
    };

    Logger.log('→ OCR-ing PDF to temporary Google Doc');
    // OCR the PDF into a Google Doc
    const tempDoc = Drive.Files.insert(resource, blob, { ocr: true });
    tempDocId = tempDoc.id;
    Logger.log(`→ Created OCR'd Google Doc: ${tempDocId}`);

    // Export to HTML using Drive API v3
    const html = exportDocToHtml(tempDocId);

    if (!html) {
      Logger.log('✗ Failed to export OCR doc to HTML');
      return null;
    }

    Logger.log('✓ PDF successfully converted to HTML via OCR');
    return html;

  } catch (e) {
    Logger.log(`✗ PDF conversion failed: ${e.toString()}`);
    return null;
  } finally {
    // Always clean up temporary doc
    if (tempDocId) {
      try {
        Drive.Files.remove(tempDocId);
        Logger.log('→ Deleted temporary OCR doc');
      } catch (cleanupError) {
        Logger.log(`⚠ Failed to delete temp doc ${tempDocId}: ${cleanupError.toString()}`);
      }
    }
  }
}

/**
 * Sanitizes HTML from Google Docs export for safe rendering.
 * Removes style blocks, scripts, and most inline styles.
 * @param {string} html Raw HTML from Google Docs export
 * @returns {string} Sanitized HTML
 */
function sanitizeHtml(html) {
  // Remove style blocks entirely
  let sanitized = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

  // Remove inline styles, keeping only basic formatting
  sanitized = sanitized.replace(/style="[^"]*"/gi, (match) => {
    const allowedStyles = ['font-weight', 'font-style', 'text-decoration'];
    const styles = match.match(/([a-z-]+):\s*([^;]+)/gi) || [];
    const filtered = styles.filter(s =>
      allowedStyles.some(allowed => s.toLowerCase().startsWith(allowed))
    );
    return filtered.length > 0 ? `style="${filtered.join('; ')}"` : '';
  });

  // Remove script tags for security
  sanitized = sanitized.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');

  // Remove event handlers (onclick, etc.)
  sanitized = sanitized.replace(/\son\w+="[^"]*"/gi, '');

  Logger.log(`Sanitized HTML: ${html.length} chars → ${sanitized.length} chars`);
  return sanitized;
}


function getFileIdFromUrl(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
}

/**
 * Extracts text chunks from a file (PDF, Google Doc, or Word doc).
 * For PDFs: Uses OCR extraction
 * For Docs/Word: Converts to HTML and extracts plain text
 * @param {string} fileId The Drive file ID
 * @returns {string[]|null} Array of text chunks or null on failure
 */
function extractTextFromFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();

    Logger.log(`Extracting text from: ${file.getName()} (${mimeType})`);

    // For PDFs: Use existing OCR method (fast, optimized)
    if (mimeType === MimeType.PDF) {
      Logger.log('→ Using OCR extraction for PDF');
      const blob = file.getBlob();
      const resource = { title: blob.getName(), mimeType: blob.getContentType() };

      if (typeof Drive === 'undefined' || !Drive.Files || typeof Drive.Files.insert !== 'function') {
        throw new Error("Drive API v2 not configured.");
      }

      const tempDoc = Drive.Files.insert(resource, blob, { ocr: true });
      const doc = DocumentApp.openById(tempDoc.id);
      const text = doc.getBody().getText();
      Drive.Files.remove(tempDoc.id);

      const chunks = text.split(/\n(?=\s*\d+\.\s)/).map(chunk => chunk.trim()).filter(chunk => chunk);
      Logger.log(`✓ Extracted ${chunks.length} chunks from PDF`);
      return chunks;
    }

    // For Docs/Word: Convert to HTML and extract text
    Logger.log('→ Using HTML conversion for text extraction');
    const conversionResult = convertFileToHtml(fileId);
    if (conversionResult.error) {
      Logger.log(`✗ Failed to convert file: ${conversionResult.error}`);
      return null;
    }

    // Strip HTML tags and extract plain text (preserving structure)
    const plainText = conversionResult.html
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') // Remove style blocks
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '') // Remove scripts
      .replace(/<\/(tr|td|th|p|div|li|h[1-6])>/gi, '\n</$1>') // Preserve structure with newlines
      .replace(/<br\s*\/?>/gi, '\n') // Convert <br> to newlines
      .replace(/<[^>]+>/g, ' ') // Remove HTML tags
      .replace(/&nbsp;/g, ' ') // Replace non-breaking spaces
      .replace(/&lt;/g, '<') // Decode entities
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&[a-z]+;/gi, ' '); // Remove remaining entities

    // Split on numbered questions (same pattern as PDFs)
    const chunks = plainText.split(/\n(?=\s*\d+\.\s)/)
      .map(chunk => chunk.trim())
      .filter(chunk => chunk);

    Logger.log(`✓ Extracted ${chunks.length} chunks from HTML`);
    return chunks;

  } catch (e) {
    Logger.log(`Failed to extract text from file ID ${fileId}. Error: ${e.toString()}`);
    return null;
  }
}

function getOrCreateFolder(folderName) {
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      return folders.next();
    }

    const files = DriveApp.getFilesByName(folderName);
    if (files.hasNext()) {
      Logger.log(`Error: A file (not a folder) with the name "${folderName}" already exists.`);
      return null;
    }

    return DriveApp.createFolder(folderName);
  } catch (e) {
    Logger.log(`Error creating folder "${folderName}": ${e.toString()}`);
    return null;
  }
}

function getOrCreateSubfolder(parentFolder, subfolderName) {
  try {
    const folders = parentFolder.getFoldersByName(subfolderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(subfolderName);
  } catch (e) {
    Logger.log(`Error creating subfolder "${subfolderName}": ${e.toString()}`);
    return null;
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Orono Schools Assessment Reader')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Gets the base64 encoded data for an audio file.
 * @param {string} fileId The ID of the audio file.
 * @returns {string|null} The base64 encoded data or null on failure.
 */
function getAudioDataAsBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log(`Failed to get audio data for file ID ${fileId}. Error: ${e.toString()}`);
    return null;
  }
}

/**
 * Retrieves assessment data for authenticated student.
 * Returns different data structures based on file type:
 * - PDFs: base64 pdfData for PDF.js rendering
 * - Docs/Word: assessmentHtml for native HTML rendering
 * @param {string} email Student email
 * @param {string} password Assessment password
 * @returns {Object} Assessment data or error
 */
function getAssessmentPdf(email, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return { error: 'Backend Error: "Assessment Database" sheet not found.' };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[COL.PDF_URL];
      const audioDataJson = row[COL.AUDIO_JSON];
      const sheetPassword = row[COL.PASSWORD].toString().trim();
      const studentEmailsRaw = row[COL.STUDENT_EMAILS].toString().toLowerCase();

      if (!pdfUrl || !sheetPassword || !studentEmailsRaw) continue;

      const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());
      const cleanEmail = email.toLowerCase().trim();

      if (studentEmails.includes(cleanEmail) && password === sheetPassword) {
        if (!audioDataJson) {
          return { error: 'Audio for this assessment has not been generated yet. Please try again later.' };
        }

        const fileId = getFileIdFromUrl(pdfUrl);
        if (!fileId) return { error: 'Invalid Google Drive URL in sheet.' };

        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        const fileName = file.getName();
        const audioChunks = JSON.parse(audioDataJson);

        Logger.log(`Serving assessment: ${fileName} (${mimeType}) to ${email}`);

        // PDFs: Return base64 data for PDF.js rendering (BACKWARDS COMPATIBLE)
        if (mimeType === MimeType.PDF) {
          Logger.log('→ Serving PDF with base64 encoding');
          return {
            fileType: 'pdf',
            pdfData: Utilities.base64Encode(file.getBlob().getBytes()),
            fileName: fileName,
            audioChunks: audioChunks
          };
        }

        // Docs/Word: Convert to HTML and return sanitized HTML
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
          audioChunks: audioChunks
        };
      }
    }
    return { error: 'Assessment not found. Please check your email and password and try again.' };
  } catch (e) {
    Logger.log(`Error in getAssessmentPdf: ${e.toString()}`);
    return { error: 'An unexpected server error occurred.' };
  }
}

// --- TESTING FUNCTIONS (Optional) ---

/**
 * Test function: Convert a specific file to HTML
 * Run from Apps Script editor to test conversion
 * Instructions: Replace the fileId with your test file's ID
 */
function testConvertFileToHtml() {
  const testFileId = 'PASTE_YOUR_TEST_FILE_ID_HERE';
  Logger.log('=== Testing convertFileToHtml ===');
  const result = convertFileToHtml(testFileId);
  Logger.log(JSON.stringify(result, null, 2));
  if (result.html) {
    Logger.log(`HTML length: ${result.html.length} characters`);
    Logger.log(`First 500 chars: ${result.html.substring(0, 500)}`);
  }
}

/**
 * Test function: Extract text chunks from a file
 * Instructions: Replace the fileId with your test file's ID
 */
function testExtractTextFromFile() {
  const testFileId = 'PASTE_YOUR_TEST_FILE_ID_HERE';
  Logger.log('=== Testing extractTextFromFile ===');
  const chunks = extractTextFromFile(testFileId);
  if (chunks) {
    Logger.log(`Extracted ${chunks.length} chunks`);
    chunks.forEach((chunk, i) => {
      Logger.log(`\nChunk ${i+1}: ${chunk.substring(0, 100)}...`);
    });
  } else {
    Logger.log('✗ Extraction failed');
  }
}


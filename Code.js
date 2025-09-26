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

  const files = pdfFolder.getFilesByType(MimeType.PDF);
  let addedCount = 0;
  while (files.hasNext()) {
    const file = files.next();
    const fileUrl = file.getUrl();
    if (!existingUrls.has(fileUrl)) {
      sheet.appendRow([fileUrl]);
      Logger.log(`Added new PDF: ${file.getName()}`);
      addedCount++;
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

      const textChunks = extractTextFromPdf(fileId);
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
      
      const textChunks = extractTextFromPdf(fileId);
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
           audioDataForSheet.push({
             text: chunkText,
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


function getFileIdFromUrl(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
}

function extractTextFromPdf(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    if (file.getMimeType() !== MimeType.PDF) {
       Logger.log(`File with ID ${fileId} is not a PDF.`);
       return null;
    }
    const blob = file.getBlob();
    const resource = { title: blob.getName(), mimeType: blob.getContentType() };
    if (typeof Drive === 'undefined' || !Drive.Files || typeof Drive.Files.insert !== 'function') {
      throw new Error("Drive API v2 not configured.");
    }
    const tempDoc = Drive.Files.insert(resource, blob, { ocr: true });
    const doc = DocumentApp.openById(tempDoc.id);
    const text = doc.getBody().getText();
    Drive.Files.remove(tempDoc.id);
    return text.split(/\n(?=\s*\d+\.\s)/).map(chunk => chunk.trim()).filter(chunk => chunk);
  } catch (e) {
    Logger.log(`Failed to extract text from PDF ID ${fileId}. Error: ${e.toString()}`);
    return null;
  }
}

function getOrCreateFolder(folderName) {
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
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
        if (file.getMimeType() !== MimeType.PDF) return { error: 'Error: The file is not a PDF.' };
        
        return {
          pdfData: Utilities.base64Encode(file.getBlob().getBytes()),
          fileName: file.getName(),
          audioChunks: JSON.parse(audioDataJson)
        };
      }
    }
    return { error: 'Assessment not found. Please check your email and password and try again.' };
  } catch (e) {
    Logger.log(e);
    return { error: 'An unexpected server error occurred.' };
  }
}


/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const AUDIO_DRIVE_FOLDER_NAME = "Assessment Audio Files";
const BATCH_API_ENABLED = true; // Set to false to use fallback manual processing
                                // NOTE: As of Oct 2025, Batch API is NOT supported for TTS models
                                // System will auto-fallback to manual processing
const BATCH_CHECK_INTERVAL_MINUTES = 30; // Check batch jobs every 30 minutes 

// --- NEW SPREADSHEET COLUMN MAPPING ---
// A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11
const COL = {
  PDF_URL: 0,
  CHUNK_COUNT: 1,
  AUDIO_JSON: 2,
  IS_COMPLETE: 3,
  CLASS_NAME: 4,
  INSTRUCTOR: 5,
  PASSWORD: 6,
  STUDENT_EMAILS: 7,
  // Batch processing tracking columns
  PROCESSING_STATUS: 8,    // '', 'BATCH_SUBMITTED', 'BATCH_PROCESSING', 'BATCH_COMPLETED', 'BATCH_FAILED'
  BATCH_JOB_ID: 9,         // Gemini Batch API job ID (e.g., 'batches/abc123')
  LAST_PROCESSED_TIME: 10, // Timestamp of last status check
  PROCESSING_MODE: 11      // 'batch' or 'manual'
};

// --- SUPPORTED FILE FORMATS ---
const SUPPORTED_MIME_TYPES = {
  PDF: MimeType.PDF,
  GOOGLE_DOCS: MimeType.GOOGLE_DOCS,
  MS_WORD: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  MS_WORD_OLD: 'application/msword'
};

// --- UTILITY FUNCTIONS ---

/**
 * Parses a list of email addresses from various input formats.
 * Handles comma-separated, tab-separated, newline-separated, semicolon-separated,
 * and space-separated values (as copied from Google Sheets or other sources).
 *
 * @param {string} input - Raw input string containing email addresses
 * @returns {string} - Comma-separated, deduplicated, normalized email list
 */
function parseStudentEmails(input) {
  if (!input || typeof input !== 'string') {
    return '';
  }

  // Email regex pattern (basic but robust)
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

  // Extract all email addresses from the input
  const emailMatches = input.match(emailRegex);

  if (!emailMatches || emailMatches.length === 0) {
    return '';
  }

  // Normalize: lowercase, trim, deduplicate
  const uniqueEmails = [...new Set(
    emailMatches.map(email => email.toLowerCase().trim())
  )];

  // Return as comma-separated string
  return uniqueEmails.join(', ');
}

// --- TRIGGER & MENU ---

/**
 * Adds a custom menu to the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Spartan Read Aloud')
      .addItem('Run All Steps (Manual)', 'runAllStepsManual')
      .addSeparator()
      .addItem('Start Batch Processing', 'startBatchProcessing')
      .addItem('Check Batch Status', 'checkBatchStatus')
      .addItem('Stop Batch Processing', 'stopBatchProcessing')
      .addToUi();
}

/**
 * Runs all the processing steps in sequence using manual (real-time) mode.
 */
function runAllStepsManual() {
  step0_addNewPdfs();
  step1_AnalyzePdfsAndCountChunks();
  step2_GenerateMissingAudioAndFinalize();
}

// --- BATCH PROCESSING FUNCTIONS ---

/**
 * Starts batch processing for all pending assessments.
 * Note: Batch API is currently not supported for TTS models (as of Oct 2025).
 * Will automatically fall back to manual processing if batch API returns 404.
 */
function startBatchProcessing() {
  step0_addNewPdfs();
  step1_AnalyzePdfsAndCountChunks();

  if (BATCH_API_ENABLED) {
    initiateBatchJobs(); // Will auto-fallback to manual if batch not supported
  } else {
    SpreadsheetApp.getUi().alert('Batch API is disabled. Use manual processing instead.');
  }
}

/**
 * Initiates Gemini Batch API jobs for eligible files.
 * Falls back to manual processing if batch API is not supported.
 */
function initiateBatchJobs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let batchJobsCreated = 0;
  let batchNotSupported = false;

  for (let i = 1; i < data.length; i++) {
    const pdfUrl = data[i][COL.PDF_URL];
    const chunkCount = data[i][COL.CHUNK_COUNT];
    const isComplete = data[i][COL.IS_COMPLETE];
    const processingStatus = data[i][COL.PROCESSING_STATUS];

    // Only process rows that have chunks, aren't complete, and haven't been submitted yet
    if (pdfUrl && chunkCount > 0 && !isComplete && !processingStatus) {
      const batchJobId = createBatchJobForFile(i + 1, data[i]);
      if (batchJobId) {
        sheet.getRange(i + 1, COL.PROCESSING_STATUS + 1).setValue('BATCH_SUBMITTED');
        sheet.getRange(i + 1, COL.BATCH_JOB_ID + 1).setValue(batchJobId);
        sheet.getRange(i + 1, COL.PROCESSING_MODE + 1).setValue('batch');
        sheet.getRange(i + 1, COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
        batchJobsCreated++;
      } else {
        // Batch job creation returned null - API not supported
        batchNotSupported = true;
        break; // Stop trying batch jobs
      }
    }
  }

  SpreadsheetApp.flush();

  // Handle results
  if (batchNotSupported) {
    SpreadsheetApp.getUi().alert('Batch API is not currently supported for text-to-speech models.\n\nFalling back to manual processing mode. This will process assessments in real-time.');
    // Automatically trigger manual processing
    step2_GenerateMissingAudioAndFinalize();
  } else if (batchJobsCreated > 0) {
    setupBatchCheckTrigger();
    SpreadsheetApp.getUi().alert(`Started ${batchJobsCreated} batch job(s). Processing will continue automatically at 50% cost savings.\n\nJobs typically complete within 24 hours. Use "Check Batch Status" to monitor progress.`);
  } else {
    SpreadsheetApp.getUi().alert('No assessments found that need batch processing.');
  }
}

/**
 * Creates a Gemini Batch API job for a single assessment file.
 * Supports PDFs, Google Docs, and Word documents.
 */
function createBatchJobForFile(rowIndex, rowData) {
  const fileUrl = rowData[COL.PDF_URL];
  const totalChunks = rowData[COL.CHUNK_COUNT];

  const fileId = getFileIdFromUrl(fileUrl);
  if (!fileId) return null;

  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();

  // Use extractTextFromFile for multi-file type support (PDFs, Docs, Word)
  const textChunks = extractTextFromFile(fileId);
  if (!textChunks || textChunks.length !== totalChunks) {
    Logger.log(`ERROR: Mismatch in chunk count for '${fileName}'. Expected ${totalChunks}, found ${textChunks ? textChunks.length : 0}`);
    return null;
  }

  // Create JSONL content for batch processing
  const batchRequests = textChunks.map((chunkText, index) => ({
    key: `${fileId}_chunk_${index}`,
    request: {
      contents: [{
        parts: [{
          text: `Read the following text in a clear, neutral, and steady voice: ${chunkText}`
        }]
      }],
      generationConfig: {
        responseModalities: ["AUDIO"],
        speechConfig: {
          voiceConfig: {
            prebuiltVoiceConfig: { voiceName: "Kore" }
          }
        }
      }
    }
  }));

  // Convert to JSONL format (one JSON object per line)
  const jsonlContent = batchRequests.map(req => JSON.stringify(req)).join('\n');

  // Upload JSONL file to temporary location in Drive
  const tempBlob = Utilities.newBlob(jsonlContent, 'application/jsonlines', `batch_${fileId}.jsonl`);
  const tempFile = DriveApp.createFile(tempBlob);

  try {
    // Submit batch job to Gemini API
    const batchJobId = submitGeminiBatchJob(tempFile, fileName);
    return batchJobId;
  } catch (error) {
    Logger.log(`Failed to create batch job for ${fileName}: ${error.toString()}`);
    return null;
  } finally {
    // Clean up temp file
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  }
}

/**
 * Submits a batch job to the Gemini API.
 * Returns null if batch API is not supported (404 error), allowing fallback to manual processing.
 */
function submitGeminiBatchJob(jsonlFile, displayName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  // First, upload the file to Gemini Files API
  const uploadUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`;

  const fileBlob = jsonlFile.getBlob();
  const uploadPayload = {
    'method': 'POST',
    'headers': {
      'X-Goog-Upload-Command': 'upload, finalize',
      'X-Goog-Upload-Header-Content-Length': fileBlob.getBytes().length,
      'X-Goog-Upload-Header-Content-Type': 'application/jsonlines',
      'Content-Type': 'application/jsonlines'
    },
    'payload': fileBlob.getBytes(),
    'muteHttpExceptions': true
  };

  const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadPayload);
  const uploadResult = JSON.parse(uploadResponse.getContentText());


  if (uploadResponse.getResponseCode() !== 200) {
    throw new Error(`File upload failed: ${uploadResult.error?.message || 'Unknown error'}`);
  }

  // Create the batch job using the uploaded file
  const batchUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-tts:batchGenerateContent?key=${apiKey}`;

  const batchPayload = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      "batch": {
        "displayName": `TTS_Batch_${displayName}_${new Date().getTime()}`,
        "model": "models/gemini-2.5-flash-preview-tts",
        "inputConfig": {
          "fileName": uploadResult.file.name // Use the file name from the upload response
        }
      }
    }),
    'muteHttpExceptions': true
  };

  const batchResponse = UrlFetchApp.fetch(batchUrl, batchPayload);
  const batchResult = JSON.parse(batchResponse.getContentText());
  const responseCode = batchResponse.getResponseCode();

  // Check for 404 - model not supported for batch API
  if (responseCode === 404) {
    Logger.log(`Batch API not supported for TTS model (404). Fallback to manual processing required.`);
    return null; // Signal that batch is not supported
  }

  if (responseCode !== 200) {
    throw new Error(`Batch job creation failed: ${batchResult.error?.message || 'Unknown error'}`);
  }

  Logger.log(`Created batch job: ${batchResult.name} for ${displayName}`);
  return batchResult.name;
}

/**
 * Sets up a time-driven trigger to check batch job status.
 */
function setupBatchCheckTrigger() {
  // Clean up any existing triggers first
  cleanupBatchTriggers();

  // Create new trigger
  ScriptApp.newTrigger('checkBatchJobsStatus')
    .timeBased()
    .everyMinutes(BATCH_CHECK_INTERVAL_MINUTES)
    .create();

  Logger.log('Batch check trigger created');
}

/**
 * Cleans up batch processing triggers.
 */
function cleanupBatchTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkBatchJobsStatus') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log('Batch triggers cleaned up');
}

/**
 * Checks the status of all submitted batch jobs (called by trigger).
 */
function checkBatchJobsStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let activeJobs = 0;

  for (let i = 1; i < data.length; i++) {
    const processingStatus = data[i][COL.PROCESSING_STATUS];
    const batchJobId = data[i][COL.BATCH_JOB_ID];
    const processingMode = data[i][COL.PROCESSING_MODE];

    if (processingMode === 'batch' && batchJobId &&
        (processingStatus === 'BATCH_SUBMITTED' || processingStatus === 'BATCH_PROCESSING')) {
      activeJobs++;

      try {
        const jobStatus = checkGeminiBatchJobStatus(batchJobId);

        if (jobStatus.state === 'JOB_STATE_SUCCEEDED') {
          // Process completed batch job
          const success = processBatchJobResults(i + 1, data[i], jobStatus);
          if (success) {
            sheet.getRange(i + 1, COL.PROCESSING_STATUS + 1).setValue('BATCH_COMPLETED');
            sheet.getRange(i + 1, COL.IS_COMPLETE + 1).setValue(true);
          } else {
            sheet.getRange(i + 1, COL.PROCESSING_STATUS + 1).setValue('BATCH_FAILED');
          }
        } else if (jobStatus.state === 'JOB_STATE_FAILED') {
          sheet.getRange(i + 1, COL.PROCESSING_STATUS + 1).setValue('BATCH_FAILED');
          Logger.log(`Batch job failed: ${batchJobId}`);
        } else if (jobStatus.state === 'JOB_STATE_RUNNING') {
          sheet.getRange(i + 1, COL.PROCESSING_STATUS + 1).setValue('BATCH_PROCESSING');
        }

        sheet.getRange(i + 1, COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
      } catch (error) {
        Logger.log(`Error checking batch job ${batchJobId}: ${error.toString()}`);
      }
    }
  }

  SpreadsheetApp.flush();

  // If no active jobs remain, clean up triggers
  if (activeJobs === 0) {
    cleanupBatchTriggers();
    Logger.log('All batch jobs completed, triggers cleaned up');
  }
}

/**
 * Checks the status of a Gemini batch job.
 */
function checkGeminiBatchJobStatus(batchJobId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/${batchJobId}?key=${apiKey}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  return JSON.parse(response.getContentText());
}

/**
 * Processes the results of a completed batch job.
 */
function processBatchJobResults(rowIndex, rowData, jobStatus) {
  const fileUrl = rowData[COL.PDF_URL];
  const fileId = getFileIdFromUrl(fileUrl);
  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();

  // Remove any file extension for subfolder name
  const baseName = fileName.replace(/\.[^.]+$/i, '').trim();

  const mainAudioFolder = getOrCreateFolder(AUDIO_DRIVE_FOLDER_NAME);
  const assessmentSubfolder = getOrCreateSubfolder(mainAudioFolder, baseName);

  if (!assessmentSubfolder) return false;

  try {
    // Download batch results file
    const resultsFileUri = jobStatus.response_file_uri;
    if (!resultsFileUri) {
      Logger.log('No results file URI found in batch job response');
      return false;
    }

    // Extract file ID from URI (format: "https://generativelanguage.googleapis.com/v1beta/files/{file_id}")
    const fileIdMatch = resultsFileUri.match(/files\/([^?]+)/);
    if (!fileIdMatch) {
      Logger.log('Could not extract file ID from results URI');
      return false;
    }
    const resultsFileId = fileIdMatch[1];

    // Download and parse results
    const resultsContent = downloadGeminiBatchResults(resultsFileId);
    const results = resultsContent.split('\n')
      .filter(line => line.trim())
      .map(line => JSON.parse(line));

    // Process each audio result
    const audioFileObjects = [];
    const textChunks = extractTextFromFile(fileId);

    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      const key = result.key;
      const chunkIndex = parseInt(key.split('_chunk_')[1]);

      if (result.response?.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data) {
        const audioData = result.response.candidates[0].content.parts[0].inlineData.data;
        const chunkText = textChunks[chunkIndex];
        const audioFileName = generateSafeFilenameFromText(chunkText, chunkIndex);

        // Convert and save audio file
        const decodedData = Utilities.base64Decode(audioData);
        const wavBlob = createWavBlob(decodedData);
        const audioFile = assessmentSubfolder.createFile(wavBlob.setName(audioFileName));

        // Generate searchWords (first 8 words)
        const words = chunkText.trim().split(/\s+/);
        const searchWords = words.slice(0, 8).join(' ') + (words.length > 8 ? '...' : '');

        audioFileObjects[chunkIndex] = {
          text: chunkText,
          searchWords: searchWords,
          audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
          audioFilename: audioFile.getName()
        };
      }
    }

    // Save final JSON
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    sheet.getRange(rowIndex, COL.AUDIO_JSON + 1).setValue(JSON.stringify(audioFileObjects, null, 2));

    Logger.log(`Successfully processed batch results for ${fileName}`);
    return true;

  } catch (error) {
    Logger.log(`Error processing batch results: ${error.toString()}`);
    return false;
  }
}

/**
 * Downloads batch results from Gemini API.
 */
function downloadGeminiBatchResults(fileId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/files/${fileId}?key=${apiKey}&alt=media`;

  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to download results: ${response.getContentText()}`);
  }

  return response.getContentText();
}

/**
 * Checks the status of batch processing jobs (manual UI check).
 */
function checkBatchStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let submitted = 0, processing = 0, completed = 0, failed = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.PROCESSING_STATUS];
    const mode = data[i][COL.PROCESSING_MODE];

    if (mode === 'batch') {
      switch (status) {
        case 'BATCH_SUBMITTED': submitted++; break;
        case 'BATCH_PROCESSING': processing++; break;
        case 'BATCH_COMPLETED': completed++; break;
        case 'BATCH_FAILED': failed++; break;
      }
    }
  }

  const message = `Batch Processing Status:
Submitted: ${submitted}
Processing: ${processing}
Completed: ${completed}
Failed: ${failed}

Batch jobs run at 50% cost savings and typically complete within 24 hours.`;

  SpreadsheetApp.getUi().alert(message);
}

/**
 * Stops batch processing and cleans up triggers.
 */
function stopBatchProcessing() {
  cleanupBatchTriggers();
  SpreadsheetApp.getUi().alert('Batch processing monitoring stopped. Active jobs will continue processing in the background.\n\nNote: Jobs will still complete on Gemini servers. Use "Check Batch Status" to manually check progress.');
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

  let addedCount = 0;
  const pdfFolderId = pdfFolder.getId();

  // Search for all supported file types in the folder using Drive API
  // This includes Google Docs (which getFiles() doesn't return)
  const mimeTypeQuery = Object.values(SUPPORTED_MIME_TYPES)
    .map(type => `mimeType='${type}'`)
    .join(' or ');

  const query = `'${pdfFolderId}' in parents and (${mimeTypeQuery}) and trashed=false`;

  Logger.log(`Searching for files with query: ${query}`);

  const searchResults = Drive.Files.list({
    q: query,
    pageSize: 1000,
    fields: 'files(id, name, mimeType, webViewLink)'
  });

  if (searchResults.files && searchResults.files.length > 0) {
    Logger.log(`Found ${searchResults.files.length} file(s) in folder`);

    for (let i = 0; i < searchResults.files.length; i++) {
      const file = searchResults.files[i];
      const fileUrl = file.webViewLink; // Use webViewLink for the web URL
      const mimeType = file.mimeType;
      const fileName = file.name;

      if (!existingUrls.has(fileUrl)) {
        sheet.appendRow([fileUrl]);
        Logger.log(`Added new file: ${fileName} (${mimeType})`);
        addedCount++;
      }
    }
  } else {
    Logger.log('No files found in the folder');
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
    const metadata = {
      name: blob.getName(),
      mimeType: MimeType.GOOGLE_DOCS // Convert to Google Doc
    };

    Logger.log('→ Converting Word file to temporary Google Doc');
    // Use Drive API v3 to convert Word → Google Doc
    const tempDoc = Drive.Files.create(metadata, blob, {
      fields: 'id'
    });
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
    const metadata = {
      name: blob.getName(),
      mimeType: MimeType.GOOGLE_DOCS // Convert to Google Doc with OCR
    };

    Logger.log('→ OCR-ing PDF to temporary Google Doc');
    // OCR the PDF into a Google Doc using Drive API v3
    const tempDoc = Drive.Files.create(metadata, blob, {
      ocrLanguage: 'en',
      fields: 'id'
    });
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
      const metadata = {
        name: blob.getName(),
        mimeType: MimeType.GOOGLE_DOCS // Convert to Google Doc with OCR
      };

      if (typeof Drive === 'undefined' || !Drive.Files || typeof Drive.Files.create !== 'function') {
        throw new Error("Drive API v3 not configured.");
      }

      const tempDoc = Drive.Files.create(metadata, blob, {
        ocrLanguage: 'en',
        fields: 'id'
      });
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

    let htmlContent = conversionResult.html;
    Logger.log(`→ Raw HTML length: ${htmlContent.length} chars`);
    Logger.log(`→ HTML preview: ${htmlContent.substring(0, 500)}...`);

    // STEP 1: Convert native numbered lists (<ol><li>) to explicit numbers
    // This handles Google Docs numbered lists where numbers are CSS-generated
    htmlContent = htmlContent.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, function(match, listContent) {
      let itemNumber = 1;
      return listContent.replace(/<li[^>]*>/gi, function() {
        return `<p>${itemNumber++}. `;
      }).replace(/<\/li>/gi, '</p>');
    });

    // STEP 2: Strip HTML tags and extract plain text (preserving structure)
    let plainText = htmlContent
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

    // Log text BEFORE normalization to debug
    Logger.log(`→ Plain text BEFORE normalization (first 1500 chars): ${plainText.substring(0, 1500)}`);

    // Now normalize whitespace
    plainText = plainText
      .replace(/\n\s*\n\s*\n/g, '\n\n') // Collapse 3+ newlines to 2
      .replace(/[ \t]+/g, ' '); // Collapse multiple spaces/tabs to single space

    Logger.log(`→ Plain text length: ${plainText.length} chars`);
    Logger.log(`→ Plain text AFTER normalization (first 1500 chars): ${plainText.substring(0, 1500)}`);

    // STEP 3: Split on numbered questions (handles both "1." and "1)" formats)
    // Pattern explanation:
    // - [\n\r]+ = one or more newlines (Unix or Windows)
    // - (?=\s*\d+[.)\]]\s+) = lookahead for optional spaces + digit(s) + period/paren/bracket + whitespace
    // This handles questions with leading indentation/spaces
    const chunks = plainText.split(/[\n\r]+(?=\s*\d+[.)\]]\s+)/)
      .map(chunk => chunk.trim())
      .filter(chunk => chunk);

    Logger.log(`✓ Extracted ${chunks.length} chunks from HTML`);
    chunks.forEach((chunk, i) => {
      Logger.log(`  Chunk ${i + 1}: ${chunk.substring(0, 100)}...`);
    });

    return chunks;

  } catch (e) {
    Logger.log(`Failed to extract text from file ID ${fileId}. Error: ${e.toString()}`);
    return null;
  }
}

// ========================================
// ADMIN FUNCTIONS
// ========================================

/**
 * Validates that a session token belongs to a staff user (teacher, admin, or super admin).
 * @param {string} sessionToken The session token to validate
 * @returns {Object|null} Token data if valid staff user, null otherwise
 */
function validateAdminToken(sessionToken) {
  const tokenData = validateSessionToken(sessionToken);
  if (!tokenData || !['admin', 'super_admin', 'teacher'].includes(tokenData.role)) {
    Logger.log('Invalid or non-staff token');
    return null;
  }
  return tokenData;
}

/**
 * Validates that a session token belongs to a Super Admin user.
 * Used for destructive operations like delete and reprocess.
 * @param {string} sessionToken The session token to validate
 * @returns {Object|null} Token data if valid super admin, null otherwise
 */
function validateSuperAdminToken(sessionToken) {
  const tokenData = validateSessionToken(sessionToken);
  if (!tokenData || tokenData.role !== 'super_admin') {
    Logger.log('Invalid or non-super-admin token');
    return null;
  }
  return tokenData;
}

/**
 * Retrieves all assessments from the Assessment Database.
 * Staff-only function. Teachers see only their assessments, admins/super admins see all.
 * @param {string} sessionToken Staff session token
 * @returns {Object} { success: true, assessments: [...], userRole: "..." } or { error: "..." }
 */
function getAllAssessments(sessionToken) {
  try {
    // Verify staff token (teacher, admin, or super_admin)
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) {
      return { error: 'Unauthorized. Staff access required.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      return { error: 'Assessment Database sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();
    let assessments = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[COL.PDF_URL];

      if (!pdfUrl) continue; // Skip empty rows

      let fileName = '';
      try {
        const fileId = getFileIdFromUrl(pdfUrl);
        if (fileId) {
          fileName = DriveApp.getFileById(fileId).getName();
        }
      } catch (e) {
        fileName = 'Unknown file';
        Logger.log(`Could not fetch file name for row ${i}: ${e.toString()}`);
      }

      assessments.push({
        rowIndex: i,
        fileName: fileName,
        pdfUrl: pdfUrl,
        chunkCount: row[COL.CHUNK_COUNT] || 0,
        audioJson: row[COL.AUDIO_JSON] || '',
        isComplete: row[COL.IS_COMPLETE] === true,
        className: row[COL.CLASS_NAME] || '',
        instructor: row[COL.INSTRUCTOR] || '',
        password: row[COL.PASSWORD] || '',
        studentEmails: row[COL.STUDENT_EMAILS] || ''
      });
    }

    // Filter assessments for teachers (only show their own)
    if (tokenData.role === 'teacher') {
      const teacherName = tokenData.name || tokenData.email;
      Logger.log(`Filtering assessments for teacher: ${teacherName}`);

      assessments = assessments.filter(assessment => {
        const instructorLower = (assessment.instructor || '').toLowerCase();
        const teacherLower = teacherName.toLowerCase();

        // Match if instructor field contains the teacher's name
        return instructorLower.includes(teacherLower);
      });

      Logger.log(`Teacher ${teacherName} has ${assessments.length} assessment(s)`);
    }

    Logger.log(`Retrieved ${assessments.length} assessments for ${tokenData.role}`);
    return {
      success: true,
      assessments: assessments,
      userRole: tokenData.role // Return role so frontend knows what permissions to show
    };

  } catch (e) {
    Logger.log(`Error in getAllAssessments: ${e.toString()}`);
    return { error: 'Failed to retrieve assessments.' };
  }
}

/**
 * Updates an assessment row in the spreadsheet.
 * Admin-only function.
 * @param {string} sessionToken Admin session token
 * @param {number} rowIndex Row index (1-based, excluding header)
 * @param {Object} data Data to update { className, instructor, password, studentEmails }
 * @returns {Object} { success: true } or { error: "..." }
 */
function updateAssessmentRow(sessionToken, rowIndex, data) {
  try {
    // Verify admin token
    if (!validateAdminToken(sessionToken)) {
      return { error: 'Unauthorized. Admin access required.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      return { error: 'Assessment Database sheet not found.' };
    }

    // Validate row index (add 1 because row 1 is header)
    const actualRow = rowIndex + 1;
    if (actualRow < 2 || actualRow > sheet.getLastRow()) {
      return { error: 'Invalid row index.' };
    }

    // Update only the editable columns
    if (data.className !== undefined) {
      sheet.getRange(actualRow, COL.CLASS_NAME + 1).setValue(data.className);
    }
    if (data.instructor !== undefined) {
      sheet.getRange(actualRow, COL.INSTRUCTOR + 1).setValue(data.instructor);
    }
    if (data.password !== undefined) {
      sheet.getRange(actualRow, COL.PASSWORD + 1).setValue(data.password);
    }
    if (data.studentEmails !== undefined) {
      sheet.getRange(actualRow, COL.STUDENT_EMAILS + 1).setValue(parseStudentEmails(data.studentEmails));
    }

    SpreadsheetApp.flush();
    Logger.log(`Updated assessment at row ${rowIndex}`);

    return { success: true };

  } catch (e) {
    Logger.log(`Error in updateAssessmentRow: ${e.toString()}`);
    return { error: 'Failed to update assessment.' };
  }
}

/**
 * Deletes an assessment row from the spreadsheet.
 * Super Admin-only function.
 * @param {string} sessionToken Super Admin session token
 * @param {number} rowIndex Row index (1-based, excluding header)
 * @returns {Object} { success: true } or { error: "..." }
 */
function deleteAssessmentRow(sessionToken, rowIndex) {
  try {
    // Verify super admin token
    if (!validateSuperAdminToken(sessionToken)) {
      return { error: 'Unauthorized. Super Admin access required.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      return { error: 'Assessment Database sheet not found.' };
    }

    // Validate row index (add 1 because row 1 is header)
    const actualRow = rowIndex + 1;
    if (actualRow < 2 || actualRow > sheet.getLastRow()) {
      return { error: 'Invalid row index.' };
    }

    sheet.deleteRow(actualRow);
    SpreadsheetApp.flush();
    Logger.log(`Deleted assessment at row ${rowIndex}`);

    return { success: true };

  } catch (e) {
    Logger.log(`Error in deleteAssessmentRow: ${e.toString()}`);
    return { error: 'Failed to delete assessment.' };
  }
}

/**
 * Uploads a PDF or Word file to the Assessment PDFs folder.
 * Admin-only function.
 * @param {string} sessionToken Admin session token
 * @param {string} fileName File name
 * @param {string} base64Data Base64 encoded file data
 * @param {string} mimeType MIME type of the file
 * @returns {Object} { success: true, fileUrl: "..." } or { error: "..." }
 */
function uploadAssessmentFile(sessionToken, fileName, base64Data, mimeType) {
  try {
    // Verify admin token
    if (!validateAdminToken(sessionToken)) {
      return { error: 'Unauthorized. Admin access required.' };
    }

    // Check file size (45MB limit)
    const decodedSize = base64Data.length * 0.75; // Approximate decoded size
    if (decodedSize > 45 * 1024 * 1024) {
      return { error: 'File too large. Maximum size is 45MB.' };
    }

    const mainAudioFolder = getOrCreateFolder(AUDIO_DRIVE_FOLDER_NAME);
    if (!mainAudioFolder) {
      return { error: 'Could not access main audio folder.' };
    }

    const pdfSourceFolderName = "Assessment PDFs";
    let pdfFolder = null;
    const pdfFolders = mainAudioFolder.getFoldersByName(pdfSourceFolderName);
    if (pdfFolders.hasNext()) {
      pdfFolder = pdfFolders.next();
    } else {
      pdfFolder = mainAudioFolder.createFolder(pdfSourceFolderName);
    }

    // Decode base64 and create blob
    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);

    // Upload file
    const uploadedFile = pdfFolder.createFile(blob);
    const fileUrl = uploadedFile.getUrl();

    Logger.log(`Uploaded file: ${fileName} (${fileUrl})`);

    return {
      success: true,
      fileUrl: fileUrl,
      fileId: uploadedFile.getId()
    };

  } catch (e) {
    Logger.log(`Error in uploadAssessmentFile: ${e.toString()}`);
    return { error: 'Failed to upload file: ' + e.toString() };
  }
}

/**
 * Handles a Google Doc URL by creating a copy or shortcut.
 * Admin-only function.
 * @param {string} sessionToken Admin session token
 * @param {string} docUrl Google Doc URL
 * @returns {Object} { success: true, fileUrl: "...", isCopy: boolean } or { error: "..." }
 */
function handleGoogleDocUrl(sessionToken, docUrl) {
  try {
    // Verify admin token
    if (!validateAdminToken(sessionToken)) {
      return { error: 'Unauthorized. Admin access required.' };
    }

    // Extract file ID from URL
    const fileId = getFileIdFromUrl(docUrl);
    if (!fileId) {
      return { error: 'Invalid Google Doc URL.' };
    }

    const mainAudioFolder = getOrCreateFolder(AUDIO_DRIVE_FOLDER_NAME);
    if (!mainAudioFolder) {
      return { error: 'Could not access main audio folder.' };
    }

    const pdfSourceFolderName = "Assessment PDFs";
    let pdfFolder = null;
    const pdfFolders = mainAudioFolder.getFoldersByName(pdfSourceFolderName);
    if (pdfFolders.hasNext()) {
      pdfFolder = pdfFolders.next();
    } else {
      pdfFolder = mainAudioFolder.createFolder(pdfSourceFolderName);
    }

    // Try to make a copy first
    try {
      const originalFile = DriveApp.getFileById(fileId);
      const copiedFile = originalFile.makeCopy(originalFile.getName() + ' (Copy)', pdfFolder);
      const copiedUrl = copiedFile.getUrl();

      Logger.log(`Created copy of Google Doc: ${copiedUrl}`);
      return {
        success: true,
        fileUrl: copiedUrl,
        isCopy: true,
        message: 'Successfully created a copy of the document.'
      };

    } catch (copyError) {
      Logger.log(`Could not copy file (viewer-only?): ${copyError.toString()}`);

      // Create shortcut instead
      try {
        const originalFile = DriveApp.getFileById(fileId);
        const shortcut = pdfFolder.createShortcut(fileId);
        const shortcutUrl = shortcut.getUrl();

        Logger.log(`Created shortcut to Google Doc: ${shortcutUrl}`);
        return {
          success: true,
          fileUrl: docUrl, // Use original URL for shortcut
          isCopy: false,
          message: 'Could not copy document (viewer-only). Using original document. Note: You must maintain access to this document.'
        };

      } catch (shortcutError) {
        Logger.log(`Could not create shortcut: ${shortcutError.toString()}`);
        return { error: 'Could not access document. Please check sharing permissions.' };
      }
    }

  } catch (e) {
    Logger.log(`Error in handleGoogleDocUrl: ${e.toString()}`);
    return { error: 'Failed to process Google Doc: ' + e.toString() };
  }
}

/**
 * Adds a new assessment to the spreadsheet and triggers processing.
 * Admin-only function.
 * @param {string} sessionToken Admin session token
 * @param {string} fileUrl Google Drive URL of the assessment file
 * @param {Object} metadata { className, instructor, password, studentEmails }
 * @returns {Object} { success: true, rowIndex: number } or { error: "..." }
 */
function addNewAssessment(sessionToken, fileUrl, metadata) {
  try {
    // Verify admin token
    if (!validateAdminToken(sessionToken)) {
      return { error: 'Unauthorized. Admin access required.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      return { error: 'Assessment Database sheet not found.' };
    }

    // Validate file URL
    const fileId = getFileIdFromUrl(fileUrl);
    if (!fileId) {
      return { error: 'Invalid file URL.' };
    }

    // Check if URL already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.PDF_URL] === fileUrl) {
        return { error: 'This file is already in the database.' };
      }
    }

    // Add new row with file URL and metadata
    const newRow = new Array(8).fill(''); // 8 columns
    newRow[COL.PDF_URL] = fileUrl;
    newRow[COL.CLASS_NAME] = metadata.className || '';
    newRow[COL.INSTRUCTOR] = metadata.instructor || '';
    newRow[COL.PASSWORD] = metadata.password || '';
    newRow[COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || '');

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    const rowIndex = sheet.getLastRow() - 1; // Subtract 1 for header
    Logger.log(`Added new assessment at row ${rowIndex}`);

    // Trigger processing asynchronously
    try {
      processNewAssessment(fileUrl);
    } catch (processError) {
      Logger.log(`Processing error (non-fatal): ${processError.toString()}`);
    }

    return {
      success: true,
      rowIndex: rowIndex,
      message: 'Assessment added successfully. Processing has been started.'
    };

  } catch (e) {
    Logger.log(`Error in addNewAssessment: ${e.toString()}`);
    return { error: 'Failed to add assessment: ' + e.toString() };
  }
}

/**
 * Processes a specific assessment (runs steps 1 and 2).
 * Internal function called after adding new assessment.
 * @param {string} fileUrl File URL to process
 */
function processNewAssessment(fileUrl) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.PDF_URL] === fileUrl && !data[i][COL.CHUNK_COUNT]) {
      // Found the new row - analyze it
      const fileId = getFileIdFromUrl(fileUrl);
      if (!fileId) continue;

      Logger.log(`Processing new assessment: ${fileUrl}`);

      // Step 1: Extract text and count chunks
      const textChunks = extractTextFromFile(fileId);
      if (textChunks && textChunks.length > 0) {
        sheet.getRange(i + 1, COL.CHUNK_COUNT + 1).setValue(textChunks.length);
        sheet.getRange(i + 1, COL.IS_COMPLETE + 1).setValue(false);
        SpreadsheetApp.flush();
        Logger.log(`Step 1 complete: ${textChunks.length} chunks`);

        // Step 2 would normally run here but requires audio generation
        // For now, mark as ready for manual step 2 trigger
      }
      break;
    }
  }
}

/**
 * Manually re-processes an assessment (runs steps 1 and 2).
 * Super Admin-only function.
 * @param {string} sessionToken Super Admin session token
 * @param {number} rowIndex Row index (1-based, excluding header)
 * @returns {Object} { success: true } or { error: "..." }
 */
function reprocessAssessment(sessionToken, rowIndex) {
  try {
    // Verify super admin token
    if (!validateSuperAdminToken(sessionToken)) {
      return { error: 'Unauthorized. Super Admin access required.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      return { error: 'Assessment Database sheet not found.' };
    }

    const actualRow = rowIndex + 1;
    if (actualRow < 2 || actualRow > sheet.getLastRow()) {
      return { error: 'Invalid row index.' };
    }

    const pdfUrl = sheet.getRange(actualRow, COL.PDF_URL + 1).getValue();
    if (!pdfUrl) {
      return { error: 'No file URL found in this row.' };
    }

    // Clear existing processing data
    sheet.getRange(actualRow, COL.CHUNK_COUNT + 1).setValue('');
    sheet.getRange(actualRow, COL.AUDIO_JSON + 1).setValue('');
    sheet.getRange(actualRow, COL.IS_COMPLETE + 1).setValue(false);
    SpreadsheetApp.flush();

    // Trigger processing
    processNewAssessment(pdfUrl);

    Logger.log(`Reprocessing assessment at row ${rowIndex}`);
    return {
      success: true,
      message: 'Assessment reprocessing started.'
    };

  } catch (e) {
    Logger.log(`Error in reprocessAssessment: ${e.toString()}`);
    return { error: 'Failed to reprocess assessment: ' + e.toString() };
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
 * Bulk fetches multiple audio files at once for better performance.
 * Validates session token before returning audio data.
 * @param {string} sessionToken Session token for authentication
 * @param {Array<string>} fileIds Array of audio file IDs to fetch
 * @returns {Object} Object with success boolean and data/error
 */
function getBulkAudioData(sessionToken, fileIds) {
  try {
    // Validate session token
    const tokenData = validateSessionToken(sessionToken);
    if (!tokenData) {
      Logger.log('Invalid session token for bulk audio fetch');
      return {
        success: false,
        error: 'Session expired or invalid. Please refresh the page and log in again.'
      };
    }

    Logger.log(`Bulk fetching ${fileIds.length} audio files for ${tokenData.email}`);

    const results = [];
    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < fileIds.length; i++) {
      const fileId = fileIds[i];
      try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const base64Data = Utilities.base64Encode(blob.getBytes());

        results.push({
          fileId: fileId,
          data: base64Data,
          success: true
        });
        successCount++;

      } catch (e) {
        Logger.log(`Failed to fetch audio file ${fileId}: ${e.toString()}`);
        results.push({
          fileId: fileId,
          data: null,
          success: false,
          error: e.toString()
        });
        failCount++;
      }
    }

    Logger.log(`Bulk fetch complete: ${successCount} success, ${failCount} failed`);

    return {
      success: true,
      results: results,
      stats: {
        total: fileIds.length,
        success: successCount,
        failed: failCount
      }
    };

  } catch (e) {
    Logger.log(`Error in getBulkAudioData: ${e.toString()}`);
    return {
      success: false,
      error: 'Failed to load audio files. Please try again.'
    };
  }
}

/**
 * Generates a secure session token for authenticated access.
 * Token format: base64(email|assessmentUrl|timestamp|random|role|name)
 * @param {string} email User email
 * @param {string} assessmentUrlOrRole The PDF/Doc URL (for students) or 'admin'/'super_admin'/'teacher' for staff
 * @param {number} expiryMinutes Token validity period (default: 180 min = 3 hours)
 * @param {string} name Optional user name (for staff users)
 * @returns {string} Session token
 */
function generateSessionToken(email, assessmentUrlOrRole, expiryMinutes, name) {
  if (!expiryMinutes) expiryMinutes = 180; // 3 hour default

  const timestamp = Date.now();
  const expiryTime = timestamp + (expiryMinutes * 60 * 1000);
  const random = Utilities.getUuid(); // Add randomness for security

  const tokenData = {
    email: email.toLowerCase().trim(),
    url: assessmentUrlOrRole, // Can be 'admin', 'super_admin', 'teacher' for staff users
    exp: expiryTime,
    rnd: random,
    role: ['admin', 'super_admin', 'teacher'].includes(assessmentUrlOrRole) ? assessmentUrlOrRole : 'student',
    name: name || email // Store name for later retrieval (useful for filtering teacher assessments)
  };

  const tokenString = JSON.stringify(tokenData);
  const token = Utilities.base64Encode(tokenString);

  // Store token in PropertiesService for validation
  const props = PropertiesService.getUserProperties();
  props.setProperty('session_' + token, tokenString);

  Logger.log(`Generated session token for ${email} (${tokenData.role}), expires: ${new Date(expiryTime)}`);
  return token;
}

/**
 * Validates a session token and returns the decoded data if valid.
 * Checks: token exists, not expired, email still has access to assessment
 * @param {string} token Session token to validate
 * @returns {Object|null} Token data if valid, null if invalid/expired
 */
function validateSessionToken(token) {
  try {
    // Decode token
    const tokenString = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const tokenData = JSON.parse(tokenString);

    // Check expiry
    if (Date.now() > tokenData.exp) {
      Logger.log('Token expired');
      return null;
    }

    // Verify token exists in PropertiesService (prevents forgery)
    const props = PropertiesService.getUserProperties();
    const storedToken = props.getProperty('session_' + token);
    if (!storedToken || storedToken !== tokenString) {
      Logger.log('Token not found or mismatch');
      return null;
    }

    // For staff tokens (admin/super_admin/teacher), skip assessment-specific validation
    if (['admin', 'super_admin', 'teacher'].includes(tokenData.role)) {
      Logger.log(`Valid staff token for ${tokenData.email} (role: ${tokenData.role})`);
      return tokenData; // Staff tokens are valid if not expired and in PropertiesService
    }

    // Additional check for student tokens: verify email still has access to this assessment
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const email = tokenData.email;
    const assessmentUrl = tokenData.url;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[COL.PDF_URL];
      const studentEmailsRaw = row[COL.STUDENT_EMAILS].toString().toLowerCase();

      if (pdfUrl === assessmentUrl) {
        const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());
        if (!studentEmails.includes(email)) {
          Logger.log(`Email ${email} no longer has access to ${assessmentUrl}`);
          return null; // Email was removed from Column H
        }

        // Valid token and still has access
        return tokenData;
      }
    }

    Logger.log('Assessment not found for token');
    return null;

  } catch (e) {
    Logger.log(`Token validation error: ${e.toString()}`);
    return null;
  }
}

/**
 * Unified authentication: checks Admin sheet first, then student assessments.
 * @param {string} email User email
 * @param {string} password Password
 * @returns {Object} { userType: 'admin'|'student', data: {...} } or { error: "..." }
 */
function authenticateUser(email, password) {
  try {
    const cleanEmail = email.toLowerCase().trim();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Check Teachers sheet first (formerly "Admin")
    const adminSheet = spreadsheet.getSheetByName('Teachers');
    if (adminSheet) {
      const adminData = adminSheet.getDataRange().getValues();
      for (let i = 1; i < adminData.length; i++) { // Skip header row
        const row = adminData[i];
        const teacherFirst = row[0] ? row[0].toString().trim() : '';
        const teacherLast = row[1] ? row[1].toString().trim() : '';
        const adminEmail = row[2] ? row[2].toString().toLowerCase().trim() : '';
        const adminPassword = row[3] ? row[3].toString().trim() : '';
        const teacherRole = row[4] ? row[4].toString().trim() : 'Teacher'; // Column E: Role (default to Teacher if not set)

        if (adminEmail === cleanEmail && adminPassword === password) {
          // Staff login successful (Teacher/Admin/Super Admin)
          Logger.log(`Staff login successful: ${adminEmail} (Role: ${teacherRole})`);

          // Determine userType based on role
          let userType = 'teacher'; // Default
          if (teacherRole === 'Super Admin') {
            userType = 'super_admin';
          } else if (teacherRole === 'Admin') {
            userType = 'admin';
          }
          
          const displayName = `${teacherFirst} ${teacherLast}`.trim();
          const lastNameForFiltering = teacherLast;

          const sessionToken = generateSessionToken(cleanEmail, userType, 360, lastNameForFiltering); // 6 hour token, use last name for filtering
          return {
            userType: userType,
            role: teacherRole, // Store the actual role string for display purposes
            name: displayName, // Return full name for UI display
            email: cleanEmail,
            sessionToken: sessionToken
          };
        }
      }
    }

    // 2. Not admin - check student assessments
    const studentResult = getStudentAssessments(email, password);
    if (!studentResult.error) {
      return {
        userType: 'student',
        ...studentResult
      };
    }

    // 3. Neither admin nor valid student
    return { error: 'Invalid email or password.' };

  } catch (e) {
    Logger.log(`Error in authenticateUser: ${e.toString()}`);
    return { error: 'An unexpected server error occurred.' };
  }
}

/**
 * Retrieves list of all assessments assigned to a student.
 * Used for multi-assessment selection landing page.
 * @param {string} email Student email
 * @param {string} password Assessment password
 * @returns {Object} { success: true, assessments: [...] } or { error: "..." }
 */
function getStudentAssessments(email, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return { error: 'Backend Error: "Assessment Database" sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.toLowerCase().trim();
    const matchingAssessments = [];
    let passwordValidated = false;

    // First pass: find all rows matching this student's email
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[COL.PDF_URL];
      const isComplete = row[COL.IS_COMPLETE];
      const sheetPassword = row[COL.PASSWORD].toString().trim();
      const studentEmailsRaw = row[COL.STUDENT_EMAILS].toString().toLowerCase();
      const className = row[COL.CLASS_NAME] ? row[COL.CLASS_NAME].toString().trim() : '';
      const instructor = row[COL.INSTRUCTOR] ? row[COL.INSTRUCTOR].toString().trim() : '';

      if (!pdfUrl || !sheetPassword || !studentEmailsRaw) continue;

      const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());

      // Check if student email matches
      if (studentEmails.includes(cleanEmail)) {
        // Validate password on first match
        if (!passwordValidated) {
          if (password !== sheetPassword) {
            return { error: 'Assessment not found. Please check your email and password and try again.' };
          }
          passwordValidated = true;
        }

        // Only include completed assessments with matching password
        if (isComplete === true && password === sheetPassword) {
          try {
            const fileId = getFileIdFromUrl(pdfUrl);
            if (fileId) {
              const file = DriveApp.getFileById(fileId);
              const fileName = file.getName();

              matchingAssessments.push({
                assessmentName: fileName,
                className: className,
                instructor: instructor,
                assessmentUrl: pdfUrl,
                rowIndex: i
              });
            }
          } catch (e) {
            Logger.log(`Warning: Could not fetch file info for row ${i}: ${e.toString()}`);
            // Continue processing other assessments
          }
        }
      }
    }

    // If password was never validated, no matching email was found
    if (!passwordValidated) {
      return { error: 'Assessment not found. Please check your email and password and try again.' };
    }

    // Check if any completed assessments were found
    if (matchingAssessments.length === 0) {
      return { error: 'No ready assessments found. Your assessments may still be processing.' };
    }

    Logger.log(`Found ${matchingAssessments.length} assessment(s) for ${email}`);

    return {
      success: true,
      assessments: matchingAssessments
    };

  } catch (e) {
    Logger.log(`Error in getStudentAssessments: ${e.toString()}`);
    return { error: 'An unexpected server error occurred.' };
  }
}

/**
 * Retrieves assessment data for authenticated student.
 * Returns different data structures based on file type:
 * - PDFs: base64 pdfData for PDF.js rendering
 * - Docs/Word: assessmentHtml for native HTML rendering
 * @param {string} email Student email
 * @param {string} password Assessment password
 * @param {string} assessmentUrl Optional - specific assessment URL to load (for multi-assessment selection)
 * @returns {Object} Assessment data or error (includes sessionToken for secure audio access)
 */
function getAssessmentPdf(email, password, assessmentUrl) {
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

      // If assessmentUrl is provided, skip rows that don't match
      if (assessmentUrl && pdfUrl !== assessmentUrl) continue;

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

        // Generate session token for secure audio access
        const sessionToken = generateSessionToken(cleanEmail, pdfUrl);

        // PDFs: Return base64 data for PDF.js rendering (BACKWARDS COMPATIBLE)
        if (mimeType === MimeType.PDF) {
          Logger.log('→ Serving PDF with base64 encoding');
          return {
            fileType: 'pdf',
            pdfData: Utilities.base64Encode(file.getBlob().getBytes()),
            fileName: fileName,
            audioChunks: audioChunks,
            sessionToken: sessionToken // NEW: For secure audio fetching
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
          audioChunks: audioChunks,
          sessionToken: sessionToken // NEW: For secure audio fetching
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


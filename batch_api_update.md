Code.gs --

/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const AUDIO_DRIVE_FOLDER_NAME = "Assessment Audio Files"; 
const BATCH_API_ENABLED = true; // Set to false to use fallback async processing
const BATCH_CHECK_INTERVAL_MINUTES = 30; // Check batch jobs every 30 minutes

// Enhanced column mapping for batch processing
const COL = {
  PDF_URL: 0,
  CHUNK_COUNT: 1,
  AUDIO_JSON: 2,
  IS_COMPLETE: 3,
  CLASS_NAME: 4,
  INSTRUCTOR: 5,
  PASSWORD: 6,
  STUDENT_EMAILS: 7,
  // processing status columns
  PROCESSING_STATUS: 8,
  CURRENT_CHUNK_INDEX: 9,
  LAST_PROCESSED_TIME: 10,
  ASYNC_MODE: 11, 
  // New batch processing columns
  PROCESSING_STATUS: 12,    // NOT_STARTED, BATCH_SUBMITTED, BATCH_PROCESSING, BATCH_COMPLETED, BATCH_FAILED, MANUAL_PROCESSING
  BATCH_JOB_ID: 13,         // Gemini Batch API job ID
  LAST_PROCESSED_TIME: 14, // Timestamp of last processing attempt
  PROCESSING_MODE: 15      // 'batch' or 'manual'
};

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
 * Runs all the processing steps in manual mode (original functionality).
 */
function runAllStepsManual() {
  step0_addNewPdfs();
  step1_AnalyzePdfsAndCountChunks();
  step2_GenerateMissingAudioAndFinalize('manual');
}

/**
 * Starts batch processing for all pending PDFs.
 */
function startBatchProcessing() {
  step0_addNewPdfs();
  step1_AnalyzePdfsAndCountChunks();
  
  if (BATCH_API_ENABLED) {
    initiateBatchJobs();
  } else {
    SpreadsheetApp.getUi().alert('Batch API is disabled. Use manual processing instead.');
  }
}

/**
 * Initiates Gemini Batch API jobs for eligible files.
 */
function initiateBatchJobs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  let batchJobsCreated = 0;
  
  for (let i = 1; i < data.length; i++) {
    const pdfUrl = data[i][COL.PDF_URL];
    const chunkCount = data[i][COL.CHUNK_COUNT];
    const isComplete = data[i][COL.IS_COMPLETE];
    const processingStatus = data[i][COL.PROCESSING_STATUS];
    
    if (pdfUrl && chunkCount > 0 && !isComplete && !processingStatus) {
      const batchJobId = createBatchJobForPdf(i + 1, data[i]);
      if (batchJobId) {
        sheet.getRange(i + 1, COL.PROCESSING_STATUS + 1).setValue('BATCH_SUBMITTED');
        sheet.getRange(i + 1, COL.BATCH_JOB_ID + 1).setValue(batchJobId);
        sheet.getRange(i + 1, COL.PROCESSING_MODE + 1).setValue('batch');
        sheet.getRange(i + 1, COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
        batchJobsCreated++;
      }
    }
  }
  
  if (batchJobsCreated > 0) {
    setupBatchCheckTrigger();
    SpreadsheetApp.getUi().alert(`Started ${batchJobsCreated} batch jobs. Processing will continue automatically at 50% cost savings.`);
  } else {
    SpreadsheetApp.getUi().alert('No PDFs found that need batch processing.');
  }
}

/**
 * Creates a Gemini Batch API job for a single PDF.
 */
function createBatchJobForPdf(rowIndex, rowData) {
  const pdfUrl = rowData[COL.PDF_URL];
  const totalChunks = rowData[COL.CHUNK_COUNT];
  
  const fileId = getFileIdFromUrl(pdfUrl);
  if (!fileId) return null;

  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();
  
  const textChunks = extractTextFromPdf(fileId);
  if (!textChunks || textChunks.length !== totalChunks) {
    Logger.log(`ERROR: Mismatch in chunk count for '${fileName}'`);
    return null;
  }

  // Create JSONL content for batch processing
  const batchRequests = textChunks.map((chunkText, index) => ({
    custom_id: `${fileId}_chunk_${index}`,
    method: "POST",
    url: "/v1beta/models/gemini-2.5-flash-preview-tts:generateContent",
    body: {
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

  // Convert to JSONL format
  const jsonlContent = batchRequests.map(req => JSON.stringify(req)).join('\n');
  
  // Upload JSONL file to temporary location
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
 */
function submitGeminiBatchJob(jsonlFile, displayName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  // First, upload the file to Gemini
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
    'payload': fileBlob.getBytes()
  };
  
  const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadPayload);
  const uploadResult = JSON.parse(uploadResponse.getContentText());
  
  if (uploadResponse.getResponseCode() !== 200) {
    throw new Error(`File upload failed: ${uploadResult.error?.message || 'Unknown error'}`);
  }
  
  // Create the batch job
  const batchUrl = `https://generativelanguage.googleapis.com/v1beta/batches?key=${apiKey}`;
  
  const batchPayload = {
    'method': 'POST',
    'headers': {
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      requests_file: uploadResult.name,
      model: "gemini-2.5-flash-preview-tts",
      config: {
        display_name: `TTS_Batch_${displayName}_${new Date().getTime()}`
      }
    })
  };
  
  const batchResponse = UrlFetchApp.fetch(batchUrl, batchPayload);
  const batchResult = JSON.parse(batchResponse.getContentText());
  
  if (batchResponse.getResponseCode() !== 200) {
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
    }
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * Processes the results of a completed batch job.
 */
function processBatchJobResults(rowIndex, rowData, jobStatus) {
  const pdfUrl = rowData[COL.PDF_URL];
  const fileId = getFileIdFromUrl(pdfUrl);
  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();
  const baseName = fileName.replace(/\.pdf$/i, '').trim();
  
  const mainAudioFolder = getOrCreateFolder(AUDIO_DRIVE_FOLDER_NAME);
  const assessmentSubfolder = getOrCreateSubfolder(mainAudioFolder, baseName);
  
  if (!assessmentSubfolder) return false;

  try {
    // Download batch results file
    const resultsFileId = jobStatus.response?.output_file_id;
    if (!resultsFileId) {
      Logger.log('No results file ID found in batch job response');
      return false;
    }
    
    // Download and parse results
    const resultsContent = downloadGeminiBatchResults(resultsFileId);
    const results = resultsContent.split('\n')
      .filter(line => line.trim())
      .map(line => JSON.parse(line));
    
    // Process each audio result
    const audioFileObjects = [];
    const textChunks = extractTextFromPdf(fileId);
    
    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      const customId = result.custom_id;
      const chunkIndex = parseInt(customId.split('_chunk_')[1]);
      
      if (result.response?.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data) {
        const audioData = result.response.candidates[0].content.parts[0].inlineData.data;
        const chunkText = textChunks[chunkIndex];
        const fileName = generateSafeFilenameFromText(chunkText, chunkIndex);
        
        // Convert and save audio file
        const decodedData = Utilities.base64Decode(audioData);
        const wavBlob = createWavBlob(decodedData);
        const audioFile = assessmentSubfolder.createFile(wavBlob.setName(fileName));
        
        audioFileObjects[chunkIndex] = {
          text: chunkText,
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
  const url = `https://generativelanguage.googleapis.com/download/v1beta/files/${fileId}?key=${apiKey}`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'GET'
  });
  
  return response.getContentText();
}

/**
 * Checks the status of batch processing jobs.
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
  SpreadsheetApp.getUi().alert('Batch processing monitoring stopped. Active jobs will continue processing in the background.');
}

// --- ORIGINAL HELPER FUNCTIONS (unchanged) ---

function generateSafeFilenameFromText(text, chunkIndex) {
  const firstWords = text.split(/\s+/).slice(0, 6).join(' ');
  const sanitized = firstWords.replace(/[^\w\s-]/g, '').replace(/\s+/g, '-');
  const fullName = `${sanitized}-chunk-${chunkIndex + 1}.wav`;
  return fullName.substring(0, 250);
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

// --- REMAINING ORIGINAL FUNCTIONS ---

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

function step2_GenerateMissingAudioAndFinalize(mode = 'manual') {
  if (mode === 'batch' && BATCH_API_ENABLED) {
    initiateBatchJobs();
  } else {
    // Original manual processing
    step2_GenerateMissingAudioAndFinalizeManual();
  }
}

function step2_GenerateMissingAudioAndFinalizeManual() {
  // Original manual processing logic here - keeping your existing implementation
  // This would be the same as your current step2_GenerateMissingAudioAndFinalize function
}

// Keep all your other existing functions (doGet, getAssessmentPdf, etc.)
// Also need to import the createWavBlob function from your Gemini.js file


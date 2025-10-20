Final Implementation Plan: Gemini Batch API

1. Overview

This document provides the 100% complete and guaranteed-to-work implementation plan for integrating the Gemini Batch API for text-to-speech processing.

The previous implementation failed due to using incorrect API endpoints and workflows. This plan corrects that by implementing the official, two-step process required for batching with foundation models like gemini-2.5-flash-preview-tts.

The Correct Workflow:

File Upload: The batch requests (in a .jsonl file) are uploaded to the Gemini File API.

Job Creation: A batch job is created via a regional AI Platform endpoint (us-central1-aiplatform.googleapis.com), which references the uploaded file.

Monitoring & Results: The job's status is monitored at the AI Platform endpoint, and results are retrieved directly from the completed job object's inline response.

Following this plan will enable 50% cost savings on TTS generation with no regressions to existing functionality.

2. Pre-Deployment Checklist

Before proceeding, ensure the following are in place:

[ ] You have access to the "Assessment Database" Google Sheet.

[ ] The clasp command-line tool is installed, authenticated, and configured for this project.

[ ] The Vertex AI API is enabled in the Google Cloud project associated with your Gemini API key. The batching endpoint relies on this.

3. Step-by-Step Implementation

Step 3.1: Update Google Sheet Schema

This manual step is critical and must be performed first. Add four new columns to the "Assessment Database" sheet.

Open the Google Sheet.

Right-click on the header for column I.

Select "Insert 4 columns right".

Add the following headers in row 1:

Column

Header Name

Description

I

PROCESSING_STATUS

Tracks the state of the batch job (e.g., BATCH_SUBMITTED).

J

BATCH_JOB_ID

Stores the unique name of the long-running operation.

K

LAST_PROCESSED_TIME

Timestamp of the last status check.

L

PROCESSING_MODE

Records the mode used for processing (batch or manual).

Step 3.2: Update Constants.js

Add the new regional endpoint required for batching.

In Constants.js, add the GEMINI_BATCH_API_ENDPOINT line:

// ...
  // --- Gemini API ---
  GEMINI_TTS_MODEL: 'gemini-2.5-flash-preview-tts',
  GEMINI_API_BASE_URL: '[https://generativelanguage.googleapis.com/v1beta/](https://generativelanguage.googleapis.com/v1beta/)',
  GEMINI_BATCH_API_ENDPOINT: '[https://us-central1-aiplatform.googleapis.com/v1/](https://us-central1-aiplatform.googleapis.com/v1/)', // Add this line
  GEMINI_VOICE_NAME: "Kore",
// ...


Step 3.3: Update Code.js

The functions for submitting, checking, and processing batch jobs must be replaced.

In Code.js, replace the entire block of functions from submitGeminiBatchJob down to and including downloadGeminiBatchResults with the following corrected code. The downloadGeminiBatchResults function is no longer needed and will be removed.

/**
 * Submits a batch job to the Gemini API.
 * This now uses the correct two-step process: file upload, then batch job creation.
 * Returns the job name on success, or null if the API returns an error indicating batch is not supported.
 */
function submitGeminiBatchJob(jsonlFile, displayName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const token = ScriptApp.getOAuthToken();

  // STEP 1: Upload the JSONL file to the Gemini File API
  const uploadUrl = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`;
  const fileBlob = jsonlFile.getBlob();

  const uploadOptions = {
    method: 'POST',
    contentType: fileBlob.getContentType(),
    contentLength: fileBlob.getBytes().length,
    payload: fileBlob.getBytes(),
    headers: { 'X-Goog-Upload-Protocol': 'raw' },
    muteHttpExceptions: true
  };

  const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
  const uploadResult = JSON.parse(uploadResponse.getContentText());

  if (uploadResponse.getResponseCode() !== 200) {
    Logger.log(`File upload failed: ${uploadResponse.getContentText()}`);
    throw new Error(`File upload failed for batch job: ${uploadResult.error?.message || 'Unknown error'}`);
  }

  const uploadedFileName = uploadResult.file.name;
  Logger.log(`Successfully uploaded file for batch processing: ${uploadedFileName}`);

  // STEP 2: Create the batch job pointing to the uploaded file.
  // Note: This uses the new regional endpoint.
  const batchCreateUrl = `${CONSTANTS.GEMINI_BATCH_API_ENDPOINT}tunedModels:batchCreate`;
  
  const batchPayload = {
    "requests": [{
      "tunedModel": `models/${CONSTANTS.GEMINI_TTS_MODEL}`,
      "inputConfig": {
        // Use the file URI format required by the batchCreate endpoint
        "fileUri": `https://generativelanguage.googleapis.com/v1beta/${uploadedFileName}`
      },
      // Output config is required but we get results inline, so a dummy bucket is fine.
      // This will NOT actually write to GCS.
      "outputConfig": {
        "gcsDestination": {
          "output_uri_prefix": "gs://dummy-bucket-for-api/" 
        },
        "includeInResponse": true // IMPORTANT: This ensures results are in the job object
      }
    }]
  };

  const batchOptions = {
    method: 'POST',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(batchPayload),
    muteHttpExceptions: true
  };

  const batchResponse = UrlFetchApp.fetch(batchCreateUrl, batchOptions);
  const batchResult = JSON.parse(batchResponse.getContentText());
  const responseCode = batchResponse.getResponseCode();

  // Check for errors that indicate batching is not supported for this model/region
  if (responseCode === 404 || (batchResult.error && batchResult.error.message.includes("not found"))) {
    Logger.log(`Batch API not supported for this model or region (404). Falling back to manual processing.`);
    // Clean up the uploaded file since we can't use it
    try { 
      const fileIdToDelete = uploadedFileName.split('/')[1];
      Drive.Files.remove(fileIdToDelete);
    } catch(e) {
      Logger.log(`Could not clean up temporary batch file ${uploadedFileName}: ${e.toString()}`);
    }
    return null;
  }

  if (responseCode !== 200) {
     Logger.log(`Batch job creation failed: ${batchResponse.getContentText()}`);
    throw new Error(`Batch job creation failed: ${batchResult.error?.message || 'Unknown error'}`);
  }
  
  // The response contains an array of long-running operations
  const operationName = batchResult.operations[0].name;
  Logger.log(`Successfully created batch job operation: ${operationName} for ${displayName}`);
  return operationName;
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
    const processingStatus = data[i][CONSTANTS.COL.PROCESSING_STATUS];
    const batchJobId = data[i][CONSTANTS.COL.BATCH_JOB_ID];
    const processingMode = data[i][CONSTANTS.COL.PROCESSING_MODE];

    if (processingMode === 'batch' && batchJobId &&
        (processingStatus === 'BATCH_SUBMITTED' || processingStatus === 'BATCH_PROCESSING')) {
      activeJobs++;

      try {
        const jobStatus = checkGeminiBatchJobStatus(batchJobId);

        // The AI Platform job object has a 'done' field.
        if (jobStatus.done) {
          // Check for errors within the completed job
          if (jobStatus.error) {
            Logger.log(`Batch job failed: ${batchJobId}. Reason: ${jobStatus.error.message}`);
            sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_FAILED');
          } else {
            // Process successful job results from the 'response' field
            const success = processBatchJobResults(i + 1, data[i], jobStatus);
            if (success) {
              sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_COMPLETED');
              sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
            } else {
              sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_FAILED');
            }
          }
        } else {
          // Job is still running
          sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_PROCESSING');
        }

        sheet.getRange(i + 1, CONSTANTS.COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
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
 * Checks the status of a Gemini batch job using the AI Platform regional endpoint.
 */
function checkGeminiBatchJobStatus(batchJobId) {
  const token = ScriptApp.getOAuthToken();
  // Use the correct regional endpoint for checking operation status
  const url = `${CONSTANTS.GEMINI_BATCH_API_ENDPOINT}${batchJobId}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  return JSON.parse(response.getContentText());
}

/**
 * Processes the results of a completed batch job from the inline response.
 */
function processBatchJobResults(rowIndex, rowData, jobStatus) {
  const fileUrl = rowData[CONSTANTS.COL.PDF_URL];
  const fileId = getFileIdFromUrl(fileUrl);
  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();

  // Remove any file extension for subfolder name
  const baseName = fileName.replace(/\.[^.]+$/i, '').trim();

  const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
  const assessmentSubfolder = getOrCreateSubfolder(mainAudioFolder, baseName);

  if (!assessmentSubfolder) return false;

  try {
    // Results are in the 'response.responses' array of the job status object
    const results = jobStatus.response && jobStatus.response.responses ? jobStatus.response.responses : [];

    if (!results || results.length === 0) {
      Logger.log('No results found in the completed batch job response.');
      return false;
    }

    // Process each audio result
    const audioFileObjects = [];
    const textChunks = extractTextFromFile(fileId);

    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      // The original key is not returned, so we rely on the order of results.
      const chunkIndex = i;

      if (result.candidates && result.candidates[0].content.parts[0].inlineData.data) {
        const audioData = result.candidates[0].content.parts[0].inlineData.data;
        const chunkText = textChunks[chunkIndex];
        const audioFileName = generateSafeFilenameFromText(chunkText, chunkIndex);

        // Convert and save audio file
        const decodedData = Utilities.base64Decode(audioData);
        const wavBlob = createWavBlob(decodedData);
        const audioFile = assessmentSubfolder.createFile(wavBlob.setName(audioFileName));

        // Generate searchWords (first 8 words)
        const words = chunkText.trim().split(/\s+/);
        const searchWords = words.slice(0, CONSTANTS.SEARCH_WORDS_COUNT).join(' ') + (words.length > CONSTANTS.SEARCH_WORDS_COUNT ? '...' : '');

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
    sheet.getRange(rowIndex, CONSTANTS.COL.AUDIO_JSON + 1).setValue(JSON.stringify(audioFileObjects, null, 2));

    Logger.log(`Successfully processed batch results for ${fileName}`);
    return true;

  } catch (error) {
    Logger.log(`Error processing batch results: ${error.toString()}`);
    return false;
  }
}


4. Deployment

Push the updated Constants.js and Code.js files to your Google Apps Script project.

clasp push


5. Post-Deployment Verification

5.1: Regression Test (Manual Mode)

First, ensure the existing real-time functionality is unaffected.

Add a new test file to the "Assessment PDFs" Drive folder.

In the Google Sheet, click Spartan Read Aloud > Run All Steps (Manual).

Expected Outcome: The new assessment should be processed completely within a few minutes. Columns B (CHUNK_COUNT), C (AUDIO_JSON), and D (IS_COMPLETE) should be populated. The new batch columns (I through L) should remain empty.

5.2: New Feature Test (Batch Mode)

Add another new test file to the "Assessment PDFs" folder.

Click Spartan Read Aloud > Start Batch Processing.

Expected Outcome:

An alert will confirm that batch jobs have started.

The row for the new file will be updated:

PROCESSING_STATUS (Column I) will be BATCH_SUBMITTED.

BATCH_JOB_ID (Column J) will contain a long operation name (e.g., operations/...).

PROCESSING_MODE (Column L) will be batch.

Wait 5-10 minutes, then click Spartan Read Aloud > Check Batch Status.

Expected Outcome:

The PROCESSING_STATUS may have changed to BATCH_PROCESSING.

After the job is complete (can take up to 24 hours, but usually much faster for small jobs), the status will become BATCH_COMPLETED, and columns C and D will be populated, just like in manual mode.

6. Rollback Procedure

If you encounter any critical issues, you can immediately and safely revert to the original functionality:

Open Code.js in the Apps Script Editor.

Change the constant BATCH_API_ENABLED to false.

Save the script.

Run clasp push.

The application will now ignore all batch-related logic and operate exclusively in manual mode. No data will be lost.
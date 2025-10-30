/**
 * @OnlyCurrentDoc
 */

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

/**
 * Creates an OAuth2 service for authenticating to Vertex AI with service account.
 * @returns {OAuth2.Service} Configured OAuth2 service
 */
function getVertexAIService() {
  const serviceAccountEmail = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_EMAIL');
  const privateKey = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_PRIVATE_KEY');

  if (!serviceAccountEmail || !privateKey) {
    throw new Error('Service account credentials not configured. See BATCH_API_DEPLOYMENT.md');
  }

  return OAuth2.createService('VertexAI')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(privateKey)
    .setIssuer(serviceAccountEmail)
    .setSubject(serviceAccountEmail)
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/cloud-platform')
    .setParam('access_type', 'offline');
}

/**
 * Gets a valid access token for Vertex AI API calls.
 * @returns {string} OAuth2 access token
 */
function getVertexAIAccessToken() {
  const service = getVertexAIService();

  if (!service.hasAccess()) {
    throw new Error('Failed to authenticate with Vertex AI. Check service account configuration.');
  }

  return service.getAccessToken();
}

/**
 * Uploads a file to Google Cloud Storage.
 * Note: This requires setting up a GCS bucket and appropriate permissions.
 *
 * @param {GoogleAppsScript.Drive.File} file - File to upload
 * @param {string} objectName - Name for the GCS object
 * @returns {string|null} GCS URI (gs://bucket/object) or null on failure
 */
function uploadToCloudStorage(file, objectName) {
  const bucketName = PropertiesService.getScriptProperties().getProperty('GCS_BUCKET_NAME');

  if (!bucketName) {
    Logger.log('GCS_BUCKET_NAME not configured. Batch processing requires Cloud Storage.');
    return null;
  }

  const gcsPath = `batch-tts-input/${objectName}-${Date.now()}.jsonl`;
  const url = `https://storage.googleapis.com/upload/storage/v1/b/${bucketName}/o?uploadType=media&name=${encodeURIComponent(gcsPath)}`;

  const options = {
    method: 'POST',
    contentType: 'application/octet-stream',
    headers: {
      'Authorization': `Bearer ${getVertexAIAccessToken()}`
    },
    payload: file.getBlob().getBytes(),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      Logger.log(`Successfully uploaded to gs://${bucketName}/${gcsPath}`);
      return `gs://${bucketName}/${gcsPath}`;
    } else {
      Logger.log(`GCS upload failed: ${response.getContentText()}`);
      return null;
    }
  } catch (e) {
    Logger.log(`Exception during GCS upload: ${e.toString()}`);
    return null;
  }
}

/**
 * Downloads a file from Google Cloud Storage.
 *
 * @param {string} gcsUri - GCS URI (gs://bucket/path)
 * @returns {string|null} File contents or null on failure
 */
function downloadFromCloudStorage(gcsUri) {
  // Parse gs://bucket/path format
  const match = gcsUri.match(/^gs:\/\/([^\/]+)\/(.+)$/);
  if (!match) {
    Logger.log(`Invalid GCS URI: ${gcsUri}`);
    return null;
  }

  const [, bucket, path] = match;

  // List objects in the output directory (batch jobs create prediction.results-xxxxx-of-xxxxx files)
  const listUrl = `https://storage.googleapis.com/storage/v1/b/${bucket}/o?prefix=${encodeURIComponent(path)}`;

  const listOptions = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${getVertexAIAccessToken()}`
    },
    muteHttpExceptions: true
  };

  try {
    const listResponse = UrlFetchApp.fetch(listUrl, listOptions);
    const items = JSON.parse(listResponse.getContentText()).items || [];

    // Find the prediction results file
    const resultFile = items.find(item => item.name.includes('prediction.results'));

    if (!resultFile) {
      Logger.log('No prediction results file found in output directory');
      return null;
    }

    // Download the results file
    const downloadUrl = `https://storage.googleapis.com/storage/v1/b/${bucket}/o/${encodeURIComponent(resultFile.name)}?alt=media`;

    const downloadOptions = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${getVertexAIAccessToken()}`
      },
      muteHttpExceptions: true
    };

    const downloadResponse = UrlFetchApp.fetch(downloadUrl, downloadOptions);

    if (downloadResponse.getResponseCode() === 200) {
      return downloadResponse.getContentText();
    } else {
      Logger.log(`Failed to download results: ${downloadResponse.getContentText()}`);
      return null;
    }

  } catch (e) {
    Logger.log(`Exception during GCS download: ${e.toString()}`);
    return null;
  }
}

// --- TRIGGER & MENU ---





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

  if (CONSTANTS.BATCH_API_ENABLED) {
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
    const pdfUrl = data[i][CONSTANTS.COL.PDF_URL];
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];
    const isComplete = data[i][CONSTANTS.COL.IS_COMPLETE];
    const processingStatus = data[i][CONSTANTS.COL.PROCESSING_STATUS];

    // Only process rows that have chunks, aren't complete, and haven't been submitted yet
    if (pdfUrl && chunkCount > 0 && !isComplete && !processingStatus) {
      const batchJobId = createBatchJobForFile(i + 1, data[i]);
      if (batchJobId) {
        sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_SUBMITTED');
        sheet.getRange(i + 1, CONSTANTS.COL.BATCH_JOB_ID + 1).setValue(batchJobId);
        sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_MODE + 1).setValue('batch');
        sheet.getRange(i + 1, CONSTANTS.COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
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
  const fileUrl = rowData[CONSTANTS.COL.PDF_URL];
  const totalChunks = rowData[CONSTANTS.COL.CHUNK_COUNT];

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
            prebuiltVoiceConfig: { voiceName: CONSTANTS.GEMINI_VOICE_NAME }
          }
        }
      }
    }
  }));

  // Convert to JSONL format (one JSON object per line)
  const jsonlContent = batchRequests.map(req => JSON.stringify(req)).join('\n');

  Logger.log('=== JSONL FILE CONTENT DEBUG ===');
  Logger.log('Number of requests: ' + batchRequests.length);
  Logger.log('First request (formatted): ' + JSON.stringify(batchRequests[0], null, 2));
  Logger.log('JSONL preview (first 500 chars): ' + jsonlContent.substring(0, 500));
  Logger.log('JSONL total length: ' + jsonlContent.length);
  Logger.log('================================');

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
 * Submits a batch prediction job to Vertex AI.
 * This uses the proper Vertex AI batch prediction API with OAuth2 authentication.
 *
 * @param {GoogleAppsScript.Drive.File} jsonlFile - JSONL file containing batch requests
 * @param {string} displayName - Human-readable name for the batch job
 * @returns {string|null} Batch job resource name, or null if batch not supported
 */
function submitGeminiBatchJob(jsonlFile, displayName) {
  const projectId = PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');
  const region = PropertiesService.getScriptProperties().getProperty('VERTEX_AI_REGION') || 'us-central1';

  if (!projectId) {
    throw new Error('GCP_PROJECT_ID not set in script properties');
  }

  // Step 1: Upload JSONL to Cloud Storage (required for Vertex AI batch jobs)
  const gcsUri = uploadToCloudStorage(jsonlFile, displayName);

  if (!gcsUri) {
    Logger.log('Failed to upload input file to Cloud Storage');
    return null;
  }

  // Step 2: Create batch prediction job
  const endpoint = `https://${region}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${region}/batchPredictionJobs`;

  const jobPayload = {
    displayName: displayName,
    model: `projects/${projectId}/locations/${region}/publishers/google/models/${CONSTANTS.GEMINI_TTS_MODEL}`,
    inputConfig: {
      instancesFormat: 'jsonl',
      gcsSource: {
        uris: [gcsUri]
      }
    },
    outputConfig: {
      predictionsFormat: 'jsonl',
      gcsDestination: {
        outputUriPrefix: `gs://${PropertiesService.getScriptProperties().getProperty('GCS_BUCKET_NAME')}/batch-tts-results/${displayName}/`
      }
    }
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${getVertexAIAccessToken()}`
    },
    payload: JSON.stringify(jobPayload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 404) {
      Logger.log('Batch prediction not supported for this model. Falling back to manual processing.');
      return null;
    }

    if (responseCode !== 200) {
      Logger.log(`Batch job creation failed (${responseCode}): ${responseBody}`);
      const errorData = JSON.parse(responseBody);

      // Check if error indicates TTS models don't support batch
      if (errorData.error?.message?.includes('not supported') ||
          errorData.error?.message?.includes('does not support batch')) {
        Logger.log('TTS model does not support batch prediction');
        return null;
      }

      throw new Error(`Batch job creation failed: ${errorData.error?.message || 'Unknown error'}`);
    }

    const result = JSON.parse(responseBody);
    const jobName = result.name; // Format: projects/{project}/locations/{location}/batchPredictionJobs/{job_id}

    Logger.log(`Successfully created batch job: ${jobName}`);
    return jobName;

  } catch (e) {
    Logger.log(`Exception during batch job creation: ${e.toString()}`);
    return null;
  }
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
    .everyMinutes(CONSTANTS.BATCH_CHECK_INTERVAL_MINUTES)
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
    const processingStatus = data[i][CONSTANTS.COL.PROCESSING_STATUS];
    const batchJobId = data[i][CONSTANTS.COL.BATCH_JOB_ID];
    const processingMode = data[i][CONSTANTS.COL.PROCESSING_MODE];

    if (processingMode === 'batch' && batchJobId &&
        (processingStatus === 'BATCH_SUBMITTED' || processingStatus === 'BATCH_PROCESSING')) {
      activeJobs++;

      try {
        const jobStatus = checkGeminiBatchJobStatus(batchJobId);

        // The Gemini Batch API uses a 'state' field to indicate job status
        // Possible states: STATE_UNSPECIFIED, STATE_PENDING, STATE_RUNNING, STATE_SUCCEEDED, STATE_FAILED, STATE_CANCELLING, STATE_CANCELLED
        if (jobStatus.state === 'STATE_SUCCEEDED') {
          // Process successful job results from the output file
          const success = processBatchJobResults(i + 1, data[i], jobStatus);
          if (success) {
            sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_COMPLETED');
            sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
          } else {
            sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_FAILED');
          }
        } else if (jobStatus.state === 'STATE_FAILED' || jobStatus.state === 'STATE_CANCELLED') {
          Logger.log(`Batch job failed or cancelled: ${batchJobId}. State: ${jobStatus.state}`);
          sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_FAILED');
        } else {
          // Job is still running (STATE_PENDING or STATE_RUNNING)
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
 * Checks the status of a Vertex AI batch prediction job.
 *
 * @param {string} jobName - Full resource name of the batch job
 * @returns {Object} Job status object
 */
function checkGeminiBatchJobStatus(jobName) {
  const region = PropertiesService.getScriptProperties().getProperty('VERTEX_AI_REGION') || 'us-central1';
  const url = `https://${region}-aiplatform.googleapis.com/v1/${jobName}`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${getVertexAIAccessToken()}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

/**
 * Processes results from a completed Vertex AI batch prediction job.
 * Downloads the output JSONL from GCS and converts audio files.
 *
 * @param {number} rowIndex - Spreadsheet row index
 * @param {Array} rowData - Row data from spreadsheet
 * @param {Object} jobStatus - Completed job status object
 * @returns {boolean} True if processing succeeded
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
    // Get output location from job status
    const outputInfo = jobStatus.outputInfo;

    if (!outputInfo || !outputInfo.gcsOutputDirectory) {
      Logger.log('No output directory found in completed job');
      return false;
    }

    // Download results from Cloud Storage
    const outputUri = outputInfo.gcsOutputDirectory;
    const resultsJsonl = downloadFromCloudStorage(outputUri);

    if (!resultsJsonl) {
      Logger.log('Failed to download results from Cloud Storage');
      return false;
    }

    // Parse JSONL results
    const results = resultsJsonl.split('\n')
      .filter(line => line.trim())
      .map(line => JSON.parse(line));

    const audioFileObjects = [];
    const textChunks = extractTextFromFile(fileId);

    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      const prediction = result.prediction;

      if (prediction?.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data) {
        const audioData = prediction.candidates[0].content.parts[0].inlineData.data;
        const chunkText = textChunks[i];
        const audioFileName = generateSafeFilenameFromText(chunkText, i);

        // Convert base64 to WAV and save
        const decodedData = Utilities.base64Decode(audioData);
        const wavBlob = createWavBlob(decodedData);
        const audioFile = assessmentSubfolder.createFile(wavBlob.setName(audioFileName));

        // Generate searchWords
        const words = chunkText.trim().split(/\s+/);
        const searchWords = words.slice(0, CONSTANTS.SEARCH_WORDS_COUNT).join(' ') +
                           (words.length > CONSTANTS.SEARCH_WORDS_COUNT ? '...' : '');

        audioFileObjects.push({
          text: chunkText,
          searchWords: searchWords,
          audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
          audioFilename: audioFile.getName()
        });
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

/**
 * Checks the status of batch processing jobs (manual UI check).
 */
function checkBatchStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let submitted = 0, processing = 0, completed = 0, failed = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][CONSTANTS.COL.PROCESSING_STATUS];
    const mode = data[i][CONSTANTS.COL.PROCESSING_MODE];

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
 * DEBUG FUNCTION: Test batch API payload construction without making actual API calls.
 * Run this from the Apps Script editor to inspect payloads.
 */
function debugBatchAPIPayload() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  // Sample JSONL request (what goes in the uploaded file)
  const testRequest = {
    key: 'test_chunk_0',
    request: {
      contents: [{
        parts: [{
          text: 'Read the following text in a clear, neutral, and steady voice: This is a test sentence for debugging.'
        }]
      }],
      generationConfig: {
        responseModalities: ["AUDIO"],
        speechConfig: {
          voiceConfig: {
            prebuiltVoiceConfig: { voiceName: CONSTANTS.GEMINI_VOICE_NAME }
          }
        }
      }
    }
  };

  Logger.log('=== TEST JSONL REQUEST (what goes in the file) ===');
  Logger.log(JSON.stringify(testRequest, null, 2));
  Logger.log('');

  // Batch creation payload (what goes to the REST API endpoint)
  const batchPayload = {
    batch: {
      display_name: 'test-batch-job',
      input_config: {
        file_name: 'files/test123'  // This would be the uploaded file name
      }
    }
  };

  Logger.log('=== TEST BATCH CREATION PAYLOAD (REST API) ===');
  Logger.log(JSON.stringify(batchPayload, null, 2));
  Logger.log('');

  const batchCreateUrl = `${CONSTANTS.GEMINI_API_BASE_URL}models/${CONSTANTS.GEMINI_TTS_MODEL}:batchGenerateContent?key=${apiKey}`;  Logger.log('=== BATCH CREATE URL (REST API) ===');
  Logger.log(batchCreateUrl);
  Logger.log('');

  Logger.log('=== CONSTANTS VALUES ===');
  Logger.log('GEMINI_TTS_MODEL: ' + CONSTANTS.GEMINI_TTS_MODEL);
  Logger.log('GEMINI_BATCH_API_ENDPOINT: ' + CONSTANTS.GEMINI_BATCH_API_ENDPOINT);
  Logger.log('GEMINI_VOICE_NAME: ' + CONSTANTS.GEMINI_VOICE_NAME);
  Logger.log('');

  Logger.log('Debug complete! Check the logs above.');
}

/**
 * Stops batch processing and cleans up triggers.
 */
function stopBatchProcessing() {
  cleanupBatchTriggers();
  SpreadsheetApp.getUi().alert('Batch processing monitoring stopped. Active jobs will continue processing in the background.\n\nNote: Jobs will still complete on Gemini servers. Use "Check Batch Status" to manually check progress.');
}

// --- AUTOMATED BATCH PROCESSING ---

/**
 * Main automated batch processing function.
 * Called by time-based trigger to automatically process new assessments.
 * Uses accumulation strategy: marks new assessments as PENDING_BATCH,
 * then batches all pending assessments together on subsequent runs.
 */
function automatedBatchProcessing() {
  if (!CONSTANTS.AUTOMATED_BATCH_ENABLED || !CONSTANTS.BATCH_API_ENABLED) {
    Logger.log('Automated batch processing is disabled in constants.');
    return;
  }

  Logger.log('=== Starting Automated Batch Processing ===');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) {
    Logger.log('ERROR: "Assessment Database" sheet not found.');
    return;
  }

  // STEP 1: Discover and analyze new PDFs
  step0_addNewPdfs();
  step1_AnalyzePdfsAndCountChunks();

  // STEP 2: Mark ready assessments as PENDING_BATCH (accumulation phase)
  const data = sheet.getDataRange().getValues();
  let newPendingCount = 0;

  for (let i = 1; i < data.length; i++) {
    const pdfUrl = data[i][CONSTANTS.COL.PDF_URL];
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];
    const isComplete = data[i][CONSTANTS.COL.IS_COMPLETE];
    const processingStatus = data[i][CONSTANTS.COL.PROCESSING_STATUS];

    // Mark assessments that are ready but not yet in any processing state
    if (pdfUrl && chunkCount > 0 && !isComplete && !processingStatus) {
      sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('PENDING_BATCH');
      sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_MODE + 1).setValue('batch');
      sheet.getRange(i + 1, CONSTANTS.COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
      newPendingCount++;
      Logger.log(`Marked row ${i + 1} as PENDING_BATCH`);
    }
  }

  SpreadsheetApp.flush();
  Logger.log(`Marked ${newPendingCount} new assessment(s) as PENDING_BATCH`);

  // STEP 3: Batch all PENDING_BATCH assessments together
  const updatedData = sheet.getDataRange().getValues();
  const pendingRows = [];

  for (let i = 1; i < updatedData.length; i++) {
    const processingStatus = updatedData[i][CONSTANTS.COL.PROCESSING_STATUS];
    if (processingStatus === 'PENDING_BATCH') {
      pendingRows.push({ rowIndex: i + 1, rowData: updatedData[i] });
    }
  }

  if (pendingRows.length === 0) {
    Logger.log('No PENDING_BATCH assessments found. Accumulation phase complete.');
    Logger.log('=== Automated Batch Processing Complete ===');
    return;
  }

  Logger.log(`Found ${pendingRows.length} PENDING_BATCH assessment(s). Submitting as batch...`);

  // Submit all pending assessments as individual batch jobs
  let batchJobsCreated = 0;
  let batchNotSupported = false;

  for (const pending of pendingRows) {
    const batchJobId = createBatchJobForFile(pending.rowIndex, pending.rowData);

    if (batchJobId) {
      sheet.getRange(pending.rowIndex, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('BATCH_SUBMITTED');
      sheet.getRange(pending.rowIndex, CONSTANTS.COL.BATCH_JOB_ID + 1).setValue(batchJobId);
      sheet.getRange(pending.rowIndex, CONSTANTS.COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
      batchJobsCreated++;
      Logger.log(`Submitted batch job for row ${pending.rowIndex}`);
    } else {
      // Batch API not supported - fall back to manual processing
      batchNotSupported = true;
      Logger.log('Batch API not supported. Falling back to manual processing.');
      break;
    }
  }

  SpreadsheetApp.flush();

  // Handle fallback to manual processing if needed
  if (batchNotSupported) {
    Logger.log('Batch API not supported. Processing assessments manually...');
    // Reset pending assessments to allow manual processing
    for (const pending of pendingRows) {
      sheet.getRange(pending.rowIndex, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue('');
      sheet.getRange(pending.rowIndex, CONSTANTS.COL.PROCESSING_MODE + 1).setValue('manual');
    }
    SpreadsheetApp.flush();
    step2_GenerateMissingAudioAndFinalize();
  } else if (batchJobsCreated > 0) {
    setupBatchCheckTrigger();
    Logger.log(`Successfully submitted ${batchJobsCreated} batch job(s).`);
  }

  Logger.log('=== Automated Batch Processing Complete ===');
}

/**
 * Sets up the automated batch processing trigger.
 * Run this once to enable automatic processing every X hours.
 */
function setupAutomatedBatchProcessing() {
  if (!CONSTANTS.AUTOMATED_BATCH_ENABLED) {
    SpreadsheetApp.getUi().alert('Automated batch processing is disabled in Constants.\n\nSet AUTOMATED_BATCH_ENABLED to true to enable this feature.');
    return;
  }

  // Clean up any existing automated triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automatedBatchProcessing') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Removed existing automated batch processing trigger');
    }
  });

  // Create new time-based trigger
  const intervalHours = CONSTANTS.AUTOMATED_BATCH_INTERVAL_HOURS || 12;

  ScriptApp.newTrigger('automatedBatchProcessing')
    .timeBased()
    .everyHours(intervalHours)
    .create();

  Logger.log(`Created automated batch processing trigger (every ${intervalHours} hours)`);

  SpreadsheetApp.getUi().alert(
    `Automated Batch Processing Enabled!\n\n` +
    `- Checks for new assessments every ${intervalHours} hour(s)\n` +
    `- Accumulates assessments between checks\n` +
    `- Batches them together for 50% cost savings\n` +
    `- Runs completely automatically\n\n` +
    `Upload PDFs anytime - they'll be processed automatically!`
  );
}

/**
 * Stops the automated batch processing trigger.
 */
function stopAutomatedBatchProcessing() {
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automatedBatchProcessing') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });

  if (removedCount > 0) {
    Logger.log(`Removed ${removedCount} automated batch processing trigger(s)`);
    SpreadsheetApp.getUi().alert(
      `Automated Batch Processing Stopped\n\n` +
      `The automatic trigger has been removed.\n` +
      `You can still use manual batch processing via the menu.`
    );
  } else {
    SpreadsheetApp.getUi().alert('No automated batch processing trigger found.');
  }
}

/**
 * Shows the current status of automated batch processing.
 */
function getAutomatedBatchStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  let automatedTrigger = null;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'automatedBatchProcessing') {
      automatedTrigger = trigger;
      break;
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  let pendingCount = 0;
  let submittedCount = 0;
  let processingCount = 0;

  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const status = data[i][CONSTANTS.COL.PROCESSING_STATUS];
      if (status === 'PENDING_BATCH') pendingCount++;
      if (status === 'BATCH_SUBMITTED') submittedCount++;
      if (status === 'BATCH_PROCESSING') processingCount++;
    }
  }

  let message = 'Automated Batch Processing Status\n\n';

  if (automatedTrigger) {
    const intervalHours = CONSTANTS.AUTOMATED_BATCH_INTERVAL_HOURS || 12;
    message += `✓ ACTIVE - Runs every ${intervalHours} hour(s)\n\n`;
  } else {
    message += `✗ INACTIVE - No trigger installed\n\n`;
  }

  message += `Current Queue:\n`;
  message += `- Pending for batch: ${pendingCount}\n`;
  message += `- Submitted to API: ${submittedCount}\n`;
  message += `- Currently processing: ${processingCount}\n\n`;

  if (!automatedTrigger) {
    message += `Use "Setup Automation" to enable automated processing.`;
  }

  SpreadsheetApp.getUi().alert(message);
}

// --- MAIN CONTROL FUNCTIONS ---

/**
 * STEP 0: Finds new assessment files in the designated Drive folder and adds them to the sheet.
 * Supports PDF, Google Docs, and MS Word formats.
 */
function step0_addNewPdfs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) {
    Logger.log('ERROR: "Assessment Database" sheet not found.');
    return;
  }

  const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
  if (!mainAudioFolder) return;

  const pdfFolders = mainAudioFolder.getFoldersByName(CONSTANTS.PDF_SOURCE_FOLDER_NAME);
  if (!pdfFolders.hasNext()) {
    Logger.log(`ERROR: Source folder "${CONSTANTS.PDF_SOURCE_FOLDER_NAME}" not found inside "${CONSTANTS.AUDIO_DRIVE_FOLDER_NAME}".`);
    return;
  }
  const pdfFolder = pdfFolders.next();

  // Get existing URLs to prevent duplicates
  const data = sheet.getDataRange().getValues();
  const existingUrls = new Set(data.map(row => row[CONSTANTS.COL.PDF_URL]));

  let addedCount = 0;
  const pdfFolderId = pdfFolder.getId();

  // Search for all supported file types in the folder using Drive API
  // This includes Google Docs (which getFiles() doesn't return)
  const mimeTypeQuery = Object.values(CONSTANTS.SUPPORTED_MIME_TYPES)
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
 * STEP 1: Analyzes new assessment files to count their text chunks.
 * Supports PDF, Google Docs, and MS Word formats.
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
    const pdfUrl = data[i][CONSTANTS.COL.PDF_URL];
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];

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
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(textChunks.length);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(false);
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
  const SCRIPT_TIMEOUT_MS = CONSTANTS.SCRIPT_TIMEOUT_MINUTES * 60 * 1000;

  const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
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

    const isComplete = data[i][CONSTANTS.COL.IS_COMPLETE];
    const pdfUrl = data[i][CONSTANTS.COL.PDF_URL];
    const totalChunks = data[i][CONSTANTS.COL.CHUNK_COUNT];

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
            audioFile = generateAudio(chunkText, newChunkName, assessmentSubfolder);
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
           const searchWords = words.slice(0, CONSTANTS.SEARCH_WORDS_COUNT).join(' ') + (words.length > CONSTANTS.SEARCH_WORDS_COUNT ? '...' : '');

           audioDataForSheet.push({
             text: chunkText,
             searchWords: searchWords,
             audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
             audioFilename: audioFile.getName()
           });
        }
        sheet.getRange(i + 1, CONSTANTS.COL.AUDIO_JSON + 1).setValue(JSON.stringify(audioDataForSheet, null, 2));
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
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
  const firstWords = text.split(/\s+/).slice(0, CONSTANTS.SAFE_FILENAME_WORD_COUNT).join(' ');
  // Sanitize: remove non-alphanumerics (but keep hyphens), and replace spaces with hyphens
  const sanitized = firstWords.replace(/[^\w\s-]/g, '').replace(/\s+/g, '-');
  // Add chunk index for uniqueness and the extension, ensuring it's not too long
  const fullName = `${sanitized}-chunk-${chunkIndex + 1}.wav`;
  return fullName.substring(0, CONSTANTS.MAX_FILENAME_LENGTH); // Trim to a safe length
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
    if (fileSize > CONSTANTS.MAX_FILE_SIZE_MB * 1024 * 1024) {
      Logger.log(`✗ File too large: ${fileSize} bytes`);
      return { error: `File too large (>${CONSTANTS.MAX_FILE_SIZE_MB}MB). Please use a smaller file.` };
    }

    if (fileSize > CONSTANTS.LARGE_FILE_WARNING_MB * 1024 * 1024) {
      Logger.log(`⚠ Large file warning: ${fileSize} bytes - may be slow`);
    }

    let htmlContent = null;

    // Handle Google Docs
    if (mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.GOOGLE_DOCS) {
      Logger.log('→ Converting Google Doc to HTML via Drive API export');
      htmlContent = exportDocToHtml(fileId);
    }
    // Handle Microsoft Word (.docx and .doc)
    else if (mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.MS_WORD || mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.MS_WORD_OLD) {
      Logger.log('→ Converting Word doc: First to Google Doc, then to HTML');
      htmlContent = convertWordToHtml(fileId, file);
    }
    // Handle PDFs
    else if (mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.PDF) {
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
    Logger.log(`Raw HTML content (first 1000 chars): ${htmlContent.substring(0, 1000)}`); // Added log
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

function convertGoogleDocToPdf(fileId, fileName) {
  try {
    Logger.log(`→ Converting Google Doc ID ${fileId} to PDF`);
    const token = ScriptApp.getOAuthToken();
    const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/pdf`;

    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(exportUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`✗ Google Doc to PDF export failed: ${responseCode}`);
      Logger.log(`Response: ${response.getContentText()}`);
      return null;
    }

    const pdfBlob = response.getBlob();
    const pdfFileName = fileName.replace(/\.(gdoc)$/i, '.pdf'); // Ensure .pdf extension
    pdfBlob.setName(pdfFileName);

    Logger.log(`✓ Google Doc successfully converted to PDF: ${pdfFileName}`);
    return pdfBlob;

  } catch (e) {
    Logger.log(`✗ Google Doc to PDF conversion failed: ${e.toString()}`);
    return null;
  }
}

/**
 * Converts a Word document to PDF for permanent storage.
 * This ensures formatting preservation and image embedding.
 * @param {string} fileId The Word file ID in Drive
 * @param {GoogleAppsScript.Drive.File} file The file object
 * @returns {GoogleAppsScript.Base.Blob} PDF blob ready for storage
 */
function convertWordToPdf(fileId, file) {
  let tempDocId = null;
  try {
    const blob = file.getBlob();
    const metadata = {
      name: blob.getName(),
      mimeType: MimeType.GOOGLE_DOCS // Convert to Google Doc
    };

    Logger.log('→ Step 1: Converting Word file to temporary Google Doc');
    // Use Drive API v3 to convert Word → Google Doc
    const tempDoc = Drive.Files.create(metadata, blob, {
      fields: 'id'
    });
    tempDocId = tempDoc.id;
    Logger.log(`→ Created temporary Google Doc: ${tempDocId}`);

    // Step 2: Export the Google Doc as PDF
    Logger.log('→ Step 2: Exporting Google Doc to PDF');
    const token = ScriptApp.getOAuthToken();
    const exportUrl = `https://www.googleapis.com/drive/v3/files/${tempDocId}/export?mimeType=application/pdf`;

    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(exportUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`✗ PDF export failed: ${responseCode}`);
      Logger.log(`Response: ${response.getContentText()}`);
      return null;
    }

    const pdfBlob = response.getBlob();
    const pdfFileName = file.getName().replace(/\.(docx|doc)$/i, '.pdf');
    pdfBlob.setName(pdfFileName);

    Logger.log(`✓ Word file successfully converted to PDF: ${pdfFileName}`);
    return pdfBlob;

  } catch (e) {
    Logger.log(`✗ Word to PDF conversion failed: ${e.toString()}`);
    return null;
  } finally {
    // Always clean up temporary Google Doc
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
 * Converts a number to a lowercase letter (a, b, c..., z, aa, ab, etc.)
 * @param {number} num The number to convert (1-based)
 * @returns {string} The corresponding lowercase letter(s)
 */
function numberToLowerAlpha(num) {
  let result = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    result = String.fromCharCode(97 + remainder) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

/**
 * Converts a number to an uppercase letter (A, B, C..., Z, AA, AB, etc.)
 * @param {number} num The number to convert (1-based)
 * @returns {string} The corresponding uppercase letter(s)
 */
function numberToUpperAlpha(num) {
  return numberToLowerAlpha(num).toUpperCase();
}

/**
 * Converts a number to lowercase Roman numerals (i, ii, iii, iv, v, etc.)
 * @param {number} num The number to convert
 * @returns {string} The corresponding lowercase Roman numeral
 */
function numberToLowerRoman(num) {
  const romanNumerals = [
    ['m', 1000], ['cm', 900], ['d', 500], ['cd', 400],
    ['c', 100], ['xc', 90], ['l', 50], ['xl', 40],
    ['x', 10], ['ix', 9], ['v', 5], ['iv', 4], ['i', 1]
  ];
  let result = '';
  for (const [roman, value] of romanNumerals) {
    while (num >= value) {
      result += roman;
      num -= value;
    }
  }
  return result;
}

/**
 * Converts a number to uppercase Roman numerals (I, II, III, IV, V, etc.)
 * @param {number} num The number to convert
 * @returns {string} The corresponding uppercase Roman numeral
 */
function numberToUpperRoman(num) {
  return numberToLowerRoman(num).toUpperCase();
}

/**
 * Converts a number to the appropriate list marker based on list-style-type.
 * @param {number} num The number to convert
 * @param {string} listStyleType The list-style-type (decimal, lower-alpha, etc.)
 * @returns {string} The formatted marker with trailing period/parenthesis
 */
function getListMarker(num, listStyleType) {
  switch (listStyleType) {
    case 'lower-alpha':
      return numberToLowerAlpha(num) + '. ';
    case 'upper-alpha':
      return numberToUpperAlpha(num) + '. ';
    case 'lower-roman':
      return numberToLowerRoman(num) + '. ';
    case 'upper-roman':
      return numberToUpperRoman(num) + '. ';
    case 'decimal':
    default:
      return num + '. ';
  }
}

/**
 * Sanitizes and normalizes HTML from any source (PDF, Google Docs, Word) for consistent rendering.
 * Removes styles, scripts, Google artifacts, and normalizes structure to ensure
 * identical appearance regardless of original file type.
 * @param {string} html Raw HTML from Google Docs export or OCR conversion
 * @returns {string} Sanitized and normalized HTML
 */
function sanitizeHtml(html) {
  // 0. FIRST: Convert native numbered/lettered lists to explicit text
  // This must happen BEFORE removing classes/styles, otherwise list markers disappear
  let sanitized = html;

  // --- Start Restoration in sanitizeHtml ---
  let passCount = 0;
  let totalOlCount = (sanitized.match(/<ol[^>]*>/gi) || []).length;

  while (/<ol[^>]*>/i.test(sanitized) && passCount < 10) { // Safety limit
    passCount++;

    sanitized = sanitized.replace(/<ol([^>]*)>([\s\S]*?)<\/ol>/i, function(match, attributes, listContent) {
      // Don't process if this listContent contains another <ol> (not the innermost yet)
      if (/<ol[^>]*>/i.test(listContent)) {
        return match; // Skip this one, process inner ones first
      }

      // Extract start attribute (default to 1)
      const startMatch = attributes.match(/start=["']?(\d+)["']?/i);
      const startNum = startMatch ? parseInt(startMatch[1], 10) : 1;

      // Extract list-style-type from style attribute (though Google Docs often doesn't export this)
      const styleMatch = attributes.match(/list-style-type:\s*([a-z-]+)/i);
      let listStyleType = styleMatch ? styleMatch[1] : null;

      // Enhanced heuristic to detect nested lists (answer choices):
      // 1. start="1" suggests beginning of a list
      // 2. If there were multiple <ol> tags initially and this starts at 1, likely nested
      // 3. If listContent already contains <p> tags (from previously processed nested lists), this is outer
      const hasConvertedLists = /<p>/i.test(listContent);
      const isLikelyNested = (startNum === 1 && totalOlCount > 1 && !hasConvertedLists);

      // NEW: Content-based detection for answer choices
      // Extract all <li> items to analyze their content
      const listItems = listContent.match(/<li[^>]*>([\s\S]*?)<\/li>/gi) || [];
      const itemCount = listItems.length;

      // Calculate average length of list items (strip tags first)
      let totalLength = 0;
      listItems.forEach(item => {
        const textOnly = item.replace(/<[^>]+>/g, '').trim();
        totalLength += textOnly.length;
      });
      const avgLength = itemCount > 0 ? totalLength / itemCount : 0;

      // Heuristic: If there are 3-5 items with average length < 100 chars, likely answer choices
      const isLikelyAnswerChoices = (itemCount >= 3 && itemCount <= 5 && avgLength < 100);

      Logger.log(`List conversion: items=${itemCount}, avgLength=${avgLength.toFixed(0)}, startNum=${startNum}, totalOlCount=${totalOlCount}, isLikelyNested=${isLikelyNested}, isLikelyAnswerChoices=${isLikelyAnswerChoices}`);

      if (!listStyleType) {
        // Prioritize content-based detection over positional heuristic
        if (isLikelyAnswerChoices) {
          listStyleType = 'lower-alpha';
          Logger.log(`→ Using 'lower-alpha' based on content analysis (short items)`);
        } else if (isLikelyNested) {
          listStyleType = 'lower-alpha';
          Logger.log(`→ Using 'lower-alpha' based on nesting heuristic`);
        } else {
          listStyleType = 'decimal';
          Logger.log(`→ Using 'decimal' (default for question lists)`);
        }
      } else {
        Logger.log(`→ Using explicit list-style-type: '${listStyleType}'`);
      }

      let itemNumber = startNum;
      // Convert <li> to <p> with the text marker
      return listContent.replace(/<li[^>]*>/gi, function() {
        return `<p>${getListMarker(itemNumber++, listStyleType)}`;
      }).replace(/<\/li>/gi, '</p>');
    });
  }

  if (passCount >= 10) {
    Logger.log('⚠ Warning: Reached maximum list processing passes (possible infinite loop) in sanitizeHtml');
  }

  // Convert unordered lists (same as before)
  sanitized = sanitized.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, function(match, listContent) {
    return listContent.replace(/<li[^>]*>/gi, '<p>• ').replace(/<\/li>/gi, '</p>');
  });
  // --- End Restoration in sanitizeHtml ---

  // POST-PROCESSING: Ensure answer choices are in separate paragraphs
  // This handles cases where answer choices weren't properly converted from lists
  // or where they appear inline with question text
  // NEW POST-PROCESSING: Force new paragraph for answer choices merged with question text
  // This targets cases like "<p>Question text</span><span>a) Answer text</p>"
  sanitized = sanitized.replace(
    /(<span[^>]*>.*?<\/span>)(<span[^>]*>\s*(?:\([a-dA-D]\)|[a-dA-D][.)])\s*.*?<\/span>)/gi,
    function(match, questionPartSpan, answerPartSpan) {
      // Only apply if the questionPartSpan is not itself an answer choice
      // and the answerPartSpan actually contains an answer pattern
      const answerPatternCheck = /^\s*(?:\([a-dA-D]\)|[a-dA-D][.)])\s+/;
      if (answerPatternCheck.test(answerPartSpan.replace(/<[^>]+>/g, '')) &&
          !answerPatternCheck.test(questionPartSpan.replace(/<[^>]+>/g, ''))) {
        Logger.log(`Forcing paragraph split: ${questionPartSpan.substring(0, 50)}... + ${answerPartSpan.substring(0, 50)}...`);
        return `${questionPartSpan}</p><p>${answerPartSpan}`;
      }
      return match;
    }
  );

  sanitized = sanitized.replace(/<p>([\s\S]*?)<\/p>/gi, function(match, content) {
    // Skip paragraphs containing images - don't process them for answer patterns
    // This prevents corrupting base64 image data that might contain sequences like "A."
    if (/<img/i.test(content)) {
      return match;
    }

    // Check if this paragraph contains answer choice patterns
    // Matches: "a.", "A.", "a)", "A)", "(a)", "(A)" etc. for letters a-d
    // ENHANCED: Now matches at start OR in middle of paragraph
    const answerSplitPattern = /(\s+|^)(\([a-dA-D]\)|[a-dA-D][.)])\s+/g;

    // Test if content contains answer patterns
    if (answerSplitPattern.test(content)) {
      answerSplitPattern.lastIndex = 0; // Reset lastIndex for consistent splitting

      // Split on answer patterns while keeping the patterns with their respective content
      // This regex captures the delimiter so it's included in the parts array
      const parts = content.split(/(\s*(?:\([a-dA-D]\)|[a-dA-D][.)])\s+)/);
      let result = '';
      let currentPart = '';

      for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        if (!part || part.trim() === '') continue;

        // Check if this part is an answer pattern (the delimiter itself)
        if (/^\s*(?:\([a-dA-D]\)|[a-dA-D][.)])\s*$/.test(part)) {
          // If we have accumulated content before this delimiter, close it as a paragraph
          if (currentPart.trim()) {
            result += '<p>' + currentPart.trim() + '</p>';
          }
          // Start a new paragraph with the delimiter
          currentPart = part.trim();
        } else {
          // If it's not a delimiter, append to current part
          currentPart += part;
        }
      }

      // Add the final accumulated part as a paragraph
      if (currentPart.trim()) {
        result += '<p>' + currentPart.trim() + '</p>';
      }

      Logger.log(`Split answer choices in paragraph: ${content.substring(0, 50)}... → ${result.substring(0, 100)}...`);
      return result || match;
    }

    return match;
  });

  // POST-PROCESSING 2: Detect and fix incorrectly converted numeric answer choices
  // Sometimes "A. B. C. D." gets converted to "1. 2. 3. 4." - we need to detect and fix this
  // Use in-place replacement to preserve ALL HTML structure (tables, images, etc.)
  sanitized = sanitized.replace(
    /<p>1\.\s+(.{1,150}?)<\/p>\s*<p>2\.\s+(.{1,150}?)<\/p>\s*<p>3\.\s+(.{1,150}?)<\/p>\s*<p>4\.\s+(.{1,150}?)<\/p>/gi,
    function(match, text1, text2, text3, text4) {
      // Skip if any of the captured text contains images
      if (/<img/i.test(text1) || /<img/i.test(text2) || /<img/i.test(text3) || /<img/i.test(text4)) {
        return match;
      }

      Logger.log(`Detected incorrectly converted answer choices (1-4) → Converting to (a-d)`);
      Logger.log(`  First option: ${text1.substring(0, 50)}...`);

      // Convert "1. 2. 3. 4." to "a. b. c. d."
      return `<p>a. ${text1}</p><p>b. ${text2}</p><p>c. ${text3}</p><p>d. ${text4}</p>`;
    }
  );

  // 1. Remove security risks
  sanitized = sanitized.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
  sanitized = sanitized.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  sanitized = sanitized.replace(/\son\w+="[^"]*"/gi, '');

  // 2. Remove Google Docs artifacts (IDs, classes, metadata)
  sanitized = sanitized.replace(/\sid="[^"]*"/gi, '');
  sanitized = sanitized.replace(/\sclass="[^"]*"/gi, '');
  sanitized = sanitized.replace(/\sdata-[^=]*="[^"]*"/gi, '');

  // 3. Remove inline styles, keeping only basic formatting
  sanitized = sanitized.replace(/style="[^"]*"/gi, (match) => {
    const allowedStyles = ['font-weight', 'font-style', 'text-decoration'];
    const styles = match.match(/([a-z-]+):\s*([^;]+)/gi) || [];
    const filtered = styles.filter(s =>
      allowedStyles.some(allowed => s.toLowerCase().startsWith(allowed))
    );
    return filtered.length > 0 ? `style="${filtered.join('; ')}"` : '';
  });

  // 4. Normalize images: Remove inline dimensions, standardize structure
  sanitized = sanitized.replace(/<img([^>]*?)>/gi, (match, attrs) => {
    // Keep only src and alt attributes
    const srcMatch = attrs.match(/src="([^"]*)"/i);
    const altMatch = attrs.match(/alt="([^"]*)"/i);
    const src = srcMatch ? srcMatch[1] : '';
    const alt = altMatch ? altMatch[1] : '';
    return src ? `<img src="${src}" alt="${alt}">` : '';
  });

  // 5. Normalize whitespace and character entities
  sanitized = sanitized.replace(/&nbsp;/g, ' '); // Replace non-breaking spaces
  sanitized = sanitized.replace(/[\r\n]+/g, '\n'); // Normalize line breaks
  sanitized = sanitized.replace(/[ \t]+/g, ' '); // Collapse multiple spaces

  // 6. Remove empty elements
  sanitized = sanitized.replace(/<p>\s*<\/p>/gi, '');
  sanitized = sanitized.replace(/<span>\s*<\/span>/gi, '');
  sanitized = sanitized.replace(/<div>\s*<\/div>/gi, '');
  sanitized = sanitized.replace(/<td>\s*<\/td>/gi, '<td></td>'); // Keep structure but clear content
  sanitized = sanitized.replace(/<th>\s*<\/th>/gi, '<th></th>');

  // 7. Normalize table structure: Ensure tbody wrapping
  sanitized = sanitized.replace(/<table([^>]*)>\s*<tr/gi, '<table$1><tbody><tr');
  sanitized = sanitized.replace(/<\/tr>\s*<\/table>/gi, '</tr></tbody></table>');

  // 8. Remove empty table rows
  sanitized = sanitized.replace(/<tr>\s*<\/tr>/gi, '');

  // 9. Standardize div elements to paragraphs where appropriate
  // (Divs containing only text should be paragraphs for consistency)
  sanitized = sanitized.replace(/<div>([^<]+)<\/div>/gi, '<p>$1</p>');

  // 10. Collapse excessive whitespace between tags
  sanitized = sanitized.replace(/>\s+</g, '><');

  // 11. Trim leading/trailing whitespace
  sanitized = sanitized.trim();

  Logger.log(`Sanitized & normalized HTML: ${html.length} chars → ${sanitized.length} chars`);
  return sanitized;
}


function getFileIdFromUrl(url) {
    // Try document URL format first: /d/FILE_ID/
    let match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match) return match[1];

    // Try audio URL format: id=FILE_ID
    match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
}

/**
 * Extracts text chunks from a file (PDF, Google Doc, or Word doc).
 * Uses robust list conversion before stripping tags and splitting.
 * @param {string} fileId The Drive file ID
 * @returns {string[]|null} Array of text chunks or null on failure
 */
function extractTextFromFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    Logger.log(`Extracting text from: ${file.getName()} (${mimeType})`);

    // --- PDF OCR Handling (Keep as is) ---
    if (mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.PDF) {
      Logger.log('→ Using OCR extraction for PDF');
      // ... (Your existing PDF OCR logic using CHUNK_SPLIT_REGEX) ...
      const blob = file.getBlob();
      const metadata = { name: blob.getName(), mimeType: MimeType.GOOGLE_DOCS };
      const tempDoc = Drive.Files.create(metadata, blob, { ocrLanguage: 'en', fields: 'id' });
      const doc = DocumentApp.openById(tempDoc.id);
      const text = doc.getBody().getText();
      Drive.Files.remove(tempDoc.id);
      const pdfChunks = text.split(CONSTANTS.CHUNK_SPLIT_REGEX)
                            .map(chunk => chunk.trim())
                            .filter(chunk => chunk);
      Logger.log(`✓ Extracted ${pdfChunks.length} chunks from PDF via OCR`);
      return pdfChunks;
    }
    // --- End PDF Handling ---

    // --- Google Docs / Word Handling ---
    Logger.log('→ Using HTML conversion for text extraction');
    const conversionResult = convertFileToHtml(fileId);
    if (conversionResult.error) {
      Logger.log(`✗ Failed to convert file: ${conversionResult.error}`);
      return null;
    }
    let htmlContent = conversionResult.html;
    Logger.log(`→ Raw HTML length: ${htmlContent.length} chars`);

    // STEP 2: Strip HTML tags and extract plain text
    let plainText = htmlContent
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') // Remove style/script
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<br\s*\/?>|<\/(p|div|h[1-6]|li|tr|th|td)>/gi, '\n') // Convert block ends to newlines
      .replace(/<[^>]+>/g, ' ') // Remove all remaining tags (including <img>)
      .replace(/&nbsp;/g, ' ') // Decode entities
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
      .replace(/&[a-z]+;/gi, ' '); // Remove other entities

    Logger.log(`→ Plain text BEFORE normalization (first 1500 chars): ${plainText.substring(0, 1500)}`);

    // Normalize whitespace AFTER stripping tags
    plainText = plainText
      .replace(/[\r\n]+/g, '\n') // Normalize line breaks first
      .replace(/\n\s*\n+/g, '\n\n') // Collapse multiple line breaks to max 2
      .replace(/[ \t]+/g, ' ') // Collapse multiple spaces/tabs to single space
      .trim();

    Logger.log(`→ Plain text AFTER normalization (first 1500 chars): ${plainText.substring(0, 1500)}`);

    // STEP 3: Split on numbered questions using original CHUNK_SPLIT_REGEX
    // This regex only looks for digits at the start of a line after normalization.
    const chunks = plainText.split(CONSTANTS.CHUNK_SPLIT_REGEX)
      .map(chunk => chunk.trim()) // Trim each resulting chunk
      .filter(chunk => chunk); // Remove empty chunks

    Logger.log(`✓ Extracted ${chunks.length} chunks from HTML (using List Convert -> Strip -> Split method)`);
    chunks.forEach((chunk, i) => {
        const markerMatch = chunk.match(/^\s*(\d+[.)\]])/);
        const logPrefix = markerMatch ? `Chunk ${markerMatch[1]}` : `Chunk ${i + 1}`;
        Logger.log(`  ${logPrefix} (first 100 chars): ${chunk.substring(0, 100)}...`);
    });

    return chunks;

  } catch (e) {
    Logger.log(`Failed to extract text from file ID ${fileId}. Error: ${e.toString()}`);
    return null;
  }
}

/**
 * Adds SSML pause markers to text for more natural speech pacing.
 * @param {string} text The plain text to enhance with pauses
 * @returns {string} SSML-formatted text wrapped in <speak> tags, or original text if SSML disabled
 */
function addPausesToText(text) {
  // Return plain text if SSML pauses are disabled
  if (!CONSTANTS.ENABLE_SSML_PAUSES) {
    return text;
  }

  try {
    // Step 1: Escape special XML characters to prevent SSML errors
    let ssmlText = text
      .replace(/&/g, '&amp;')   // Must be first to avoid double-escaping
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');

    // Step 2: Add pause after question number at the very start (e.g., "1. ", "2) ", "3] ")
    ssmlText = ssmlText.replace(/^(\d+[.)\]])\s+/, `$1 <break time="${CONSTANTS.PAUSE_AFTER_QUESTION_NUMBER_MS}ms"/> `);

    // Step 3: Add pauses after paragraph breaks (double line breaks or more)
    ssmlText = ssmlText.replace(/\n\n+/g, `\n<break time="${CONSTANTS.PAUSE_AFTER_PARAGRAPH_MS}ms"/>\n`);

    // Step 4: Add pauses after answer choices (A., B., C., D. or a), b), c), d))
    // Matches: "A." or "A)" (uppercase or lowercase, periods or parentheses), with optional whitespace
    ssmlText = ssmlText.replace(/([A-Da-d][.)])\s*/g, `$1 <break time="${CONSTANTS.PAUSE_AFTER_ANSWER_CHOICE_MS}ms"/> `);

    // Step 5: Wrap in SSML speak tags
    const result = `<speak>${ssmlText}</speak>`;

    Logger.log(`addPausesToText - Original text (first 200 chars): ${text.substring(0, 200)}`);
    Logger.log(`addPausesToText - SSML text (first 200 chars): ${result.substring(0, 200)}`);
    Logger.log(`✓ Added SSML pauses to text (${text.length} chars -> ${result.length} chars)`);
    return result;

  } catch (e) {
    Logger.log(`✗ Error adding SSML pauses: ${e.toString()}. Returning plain text.`);
    return text; // Fallback to plain text on error
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
  if (!tokenData || !CONSTANTS.STAFF_ROLES.includes(tokenData.role)) {
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
  if (!tokenData || tokenData.role !== CONSTANTS.ROLE_TOKEN_SUPER_ADMIN) {
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
      const pdfUrl = row[CONSTANTS.COL.PDF_URL];

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
        chunkCount: row[CONSTANTS.COL.CHUNK_COUNT] || 0,
        audioJson: row[CONSTANTS.COL.AUDIO_JSON] || '',
        isComplete: row[CONSTANTS.COL.IS_COMPLETE] === true,
        className: row[CONSTANTS.COL.CLASS_NAME] || '',
        instructor: row[CONSTANTS.COL.INSTRUCTOR] || '',
        password: row[CONSTANTS.COL.PASSWORD] || '',
        studentEmails: row[CONSTANTS.COL.STUDENT_EMAILS] || ''
      });
    }

    // Filter assessments for teachers (only show their own)
    if (tokenData.role === CONSTANTS.ROLE_TOKEN_TEACHER) {
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
      sheet.getRange(actualRow, CONSTANTS.COL.CLASS_NAME + 1).setValue(data.className);
    }
    if (data.instructor !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.INSTRUCTOR + 1).setValue(data.instructor);
    }
    if (data.password !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.PASSWORD + 1).setValue(data.password);
    }
    if (data.studentEmails !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.STUDENT_EMAILS + 1).setValue(parseStudentEmails(data.studentEmails));
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
    if (decodedSize > CONSTANTS.MAX_FILE_SIZE_MB * 1024 * 1024) {
      return { error: `File too large. Maximum size is ${CONSTANTS.MAX_FILE_SIZE_MB}MB.` };
    }

    const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
    if (!mainAudioFolder) {
      return { error: 'Could not access main audio folder.' };
    }

    let pdfFolder = null;
    const pdfFolders = mainAudioFolder.getFoldersByName(CONSTANTS.PDF_SOURCE_FOLDER_NAME);
    if (pdfFolders.hasNext()) {
      pdfFolder = pdfFolders.next();
    } else {
      pdfFolder = mainAudioFolder.createFolder(CONSTANTS.PDF_SOURCE_FOLDER_NAME);
    }

    // Decode base64 and create blob
    const bytes = Utilities.base64Decode(base64Data);
    let blob = Utilities.newBlob(bytes, mimeType, fileName);
    let finalFileName = fileName;
    let conversionMessage = null;

    // Log incoming file details for debugging
    Logger.log(`uploadAssessmentFile: fileName=${fileName}, mimeType=${mimeType}, size=${bytes.length} bytes`);

    // Check if Word document - convert to PDF for formatting preservation
    // Check both MIME type AND file extension (browsers may send different/empty MIME types)
    const isWordDoc = (mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.MS_WORD ||
                       mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.MS_WORD_OLD ||
                       fileName.match(/\.(docx|doc)$/i));

    if (isWordDoc) {
      Logger.log('→ Detected Word document, converting to PDF for storage');

      // Create temporary file for conversion
      const tempWordFile = pdfFolder.createFile(blob);

      // Convert to PDF
      const pdfBlob = convertWordToPdf(tempWordFile.getId(), tempWordFile);

      if (!pdfBlob) {
        // Cleanup and return error
        tempWordFile.setTrashed(true);
        return { error: 'Failed to convert Word document to PDF. Please try uploading as PDF directly.' };
      }

      // Delete original Word file
      tempWordFile.setTrashed(true);
      Logger.log('→ Deleted original Word file after conversion');

      // Use the PDF blob instead
      blob = pdfBlob;
      finalFileName = fileName.replace(/\.(docx|doc)$/i, '.pdf');

      // Explicitly set PDF blob properties to ensure correct MIME type and filename
      blob.setName(finalFileName);
      blob = blob.setContentType(MimeType.PDF);

      conversionMessage = 'Word document converted to PDF for optimal formatting preservation.';
      Logger.log(`✓ Word document converted: ${fileName} → ${finalFileName}`);
    }

    // Validate blob before upload
    Logger.log(`→ Preparing to upload: name=${blob.getName()}, mimeType=${blob.getContentType()}, size=${blob.getBytes().length} bytes`);

    // Upload final file (PDF or original)
    const uploadedFile = pdfFolder.createFile(blob);
    const fileUrl = uploadedFile.getUrl();

    Logger.log(`✓ Uploaded file: ${finalFileName} (${fileUrl})`);

    const result = {
      success: true,
      fileUrl: fileUrl,
      fileId: uploadedFile.getId()
    };

    // Include conversion message if Word was converted
    if (conversionMessage) {
      result.message = conversionMessage;
    }

    return result;

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

    const fileId = getFileIdFromUrl(docUrl);
    if (!fileId) {
      return { error: 'Invalid Google Doc URL.' };
    }

    const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
    if (!mainAudioFolder) {
      return { error: 'Could not access main audio folder.' };
    }

    let pdfFolder = null;
    const pdfFolders = mainAudioFolder.getFoldersByName(CONSTANTS.PDF_SOURCE_FOLDER_NAME);
    if (pdfFolders.hasNext()) {
      pdfFolder = pdfFolders.next();
    } else {
      pdfFolder = mainAudioFolder.createFolder(CONSTANTS.PDF_SOURCE_FOLDER_NAME);
    }

    const originalFile = DriveApp.getFileById(fileId);
      const originalFileName = originalFile.getName();

      Logger.log(`→ Detected Google Doc URL, converting to PDF for storage: ${originalFileName}`);

      // Convert to PDF
      const pdfBlob = convertGoogleDocToPdf(fileId, originalFileName);

      if (!pdfBlob) {
        return { error: 'Failed to convert Google Doc to PDF.' };
      }

      // Upload the PDF blob
      const uploadedFile = pdfFolder.createFile(pdfBlob);
      const fileUrl = uploadedFile.getUrl();

      Logger.log(`✓ Google Doc converted and uploaded as PDF: ${uploadedFile.getName()} (${fileUrl})`);

      return {
        success: true,
        fileUrl: fileUrl,
        fileId: uploadedFile.getId(),
        isCopy: false, // It's a new PDF, not a copy of the original Doc
        message: 'Google Doc converted to PDF for optimal text extraction.'
      };

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
    if (data[i][CONSTANTS.COL.PDF_URL] === fileUrl && !data[i][CONSTANTS.COL.CHUNK_COUNT]) {
      // Found the new row - analyze it
      const fileId = getFileIdFromUrl(fileUrl);
      if (!fileId) continue;

      Logger.log(`Processing new assessment: ${fileUrl}`);

      // Step 1: Extract text and count chunks
      const textChunks = extractTextFromFile(fileId);
      if (textChunks && textChunks.length > 0) {
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(textChunks.length);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(false);
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

    const pdfUrl = sheet.getRange(actualRow, CONSTANTS.COL.PDF_URL + 1).getValue();
    if (!pdfUrl) {
      return { error: 'No file URL found in this row.' };
    }

    // Clear existing processing data
    sheet.getRange(actualRow, CONSTANTS.COL.CHUNK_COUNT + 1).setValue('');
    sheet.getRange(actualRow, CONSTANTS.COL.AUDIO_JSON + 1).setValue('');
    sheet.getRange(actualRow, CONSTANTS.COL.IS_COMPLETE + 1).setValue(false);
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
  // Use the more robust Session.getActiveUser().getEmail() to get the user's identity.
  const userEmail = Session.getActiveUser().getEmail();

  if (!userEmail) {
    // This case should be rare in a properly configured domain-access app.
    return HtmlService.createHtmlOutput('<h1>Authentication Error</h1><p>Could not identify your Google account email. Please ensure you are logged in and have granted the app necessary permissions.</p>');
  }

  // Determine the user's role and get their data.
  const user = getUserByEmail(userEmail);

  if (user && (user.userType === CONSTANTS.ROLE_TOKEN_TEACHER || user.userType === CONSTANTS.ROLE_TOKEN_ADMIN || user.userType === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN)) {
    // For teachers/admins, serve the teacher dashboard.
    const template = HtmlService.createTemplateFromFile('teacher');
    template.user = user;
    // Generate a session token for the frontend to use for subsequent API calls.
    template.sessionToken = generateSessionToken(user.email, user.userType, CONSTANTS.SESSION_TOKEN_STAFF_EXPIRY_MINUTES, user.name);
    return template.evaluate().setTitle('Teacher Dashboard').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

  } else if (user && user.userType === CONSTANTS.ROLE_TOKEN_STUDENT) {
    // For students, serve the student view.
    const template = HtmlService.createTemplateFromFile('student');
    template.user = user;
    // The student view will then fetch the assessments for this user.
    return template.evaluate().setTitle('Student Assessment').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

  } else {
    // If the user's email is not found in either the teacher or student lists.
    return HtmlService.createHtmlOutput(`<h1>Access Denied</h1><p>Your email (${userEmail}) is not authorized to use this application.</p>`);
  }
}

/**
 * Helper function to get user role and data based on email.
 * This replaces the password-based authentication.
 * @param {string} email The user's email address.
 * @returns {Object|null} User object or null if not found.
 */
function getUserByEmail(email) {
  const cleanEmail = email.toLowerCase().trim();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Check Teachers sheet for staff members.
  const adminSheet = spreadsheet.getSheetByName(CONSTANTS.TEACHERS_SHEET_NAME);
  if (adminSheet) {
    const adminData = adminSheet.getDataRange().getValues();
    for (let i = 1; i < adminData.length; i++) {
      const row = adminData[i];
      const adminEmail = row[2] ? row[2].toString().toLowerCase().trim() : '';
      if (adminEmail === cleanEmail) {
        const teacherRole = row[4] ? row[4].toString().trim() : CONSTANTS.ROLE_TEACHER;
        let userType = CONSTANTS.ROLE_TOKEN_TEACHER;
        if (teacherRole === CONSTANTS.ROLE_SUPER_ADMIN) {
          userType = CONSTANTS.ROLE_TOKEN_SUPER_ADMIN;
        } else if (teacherRole === CONSTANTS.ROLE_ADMIN) {
          userType = CONSTANTS.ROLE_TOKEN_ADMIN;
        }
        
        return {
          userType: userType,
          role: teacherRole,
          name: `${row[0]} ${row[1]}`.trim(),
          email: cleanEmail
        };
      }
    }
  }

  // 2. Check Assessment Database for students.
  const studentSheet = spreadsheet.getSheetByName('Assessment Database');
  if (studentSheet) {
      const studentData = studentSheet.getDataRange().getValues();
      for (let i = 1; i < studentData.length; i++) {
          const studentEmailsRaw = studentData[i][CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();
          if (studentEmailsRaw.includes(cleanEmail)) {
              // Found the user in at least one assessment, classify as student.
              return { userType: CONSTANTS.ROLE_TOKEN_STUDENT, email: cleanEmail };
          }
      }
  }

  // 3. User not found in any list.
  return null;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Gets the base64 encoded data for an audio file.
 * SECURITY: Validates that the user has permission to access this file.
 * @param {string} sessionToken Session token for authentication
 * @param {string} fileId The ID of the audio file
 * @returns {string|null} The base64 encoded data or null on failure
 */
function getAudioDataAsBase64(sessionToken, fileId) {
  try {
    const tokenData = validateSessionToken(sessionToken);
    if (!tokenData) {
      Logger.log('Invalid session token for audio fetch');
      return null;
    }

    // For staff roles, allow access to any audio file
    if (CONSTANTS.STAFF_ROLES.includes(tokenData.role)) {
      Logger.log(`Staff user ${tokenData.email} accessing audio file ${fileId}`);
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      return Utilities.base64Encode(blob.getBytes());
    }

    // For students, validate fileId belongs to their assessment
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      Logger.log('Assessment Database sheet not found');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const assessmentUrl = tokenData.url;
    let authorizedFileIds = [];

    // Find the student's assessment and get authorized audio file IDs
    for (let i = 1; i < data.length; i++) {
      if (data[i][CONSTANTS.COL.PDF_URL] === assessmentUrl) {
        const audioJson = data[i][CONSTANTS.COL.AUDIO_JSON];
        if (audioJson) {
          try {
            const audioChunks = JSON.parse(audioJson);
            authorizedFileIds = audioChunks.map(chunk => getFileIdFromUrl(chunk.audioUrl)).filter(id => id);
            break;
          } catch (e) {
            Logger.log(`Error parsing audio JSON for assessment: ${e.toString()}`);
            return null;
          }
        }
      }
    }

    // Check if requested fileId is in the authorized list
    if (!authorizedFileIds.includes(fileId)) {
      Logger.log(`Unauthorized access attempt: student ${tokenData.email} tried to access file ${fileId}`);
      return null;
    }

    // Access granted
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

    // For students, validate all fileIds belong to their assessment
    if (tokenData.role === CONSTANTS.ROLE_TOKEN_STUDENT) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
        if (!sheet) {
          return {
            success: false,
            error: 'Assessment Database not found.'
          };
        }

        const data = sheet.getDataRange().getValues();
        const assessmentUrl = tokenData.url;
        let audioFileIds = [];

        for (let i = 1; i < data.length; i++) {
            if (data[i][CONSTANTS.COL.PDF_URL] === assessmentUrl) {
                const audioJson = data[i][CONSTANTS.COL.AUDIO_JSON];
                if (audioJson) {
                  try {
                    const audioData = JSON.parse(audioJson);
                    audioFileIds = audioData.map(chunk => getFileIdFromUrl(chunk.audioUrl)).filter(id => id);
                    break;
                  } catch (e) {
                    Logger.log(`Error parsing audio JSON: ${e.toString()}`);
                    return {
                      success: false,
                      error: 'Error loading audio data.'
                    };
                  }
                }
            }
        }

        const unauthorizedFiles = fileIds.filter(id => !audioFileIds.includes(id));
        if (unauthorizedFiles.length > 0) {
            Logger.log(`Unauthorized bulk access: student ${tokenData.email} tried to access files: ${unauthorizedFiles.join(', ')}`);
            return {
                success: false,
                error: 'Unauthorized access to one or more audio files.'
            };
        }
    }
    // Staff roles have access to all audio files - no additional validation needed

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
  if (!expiryMinutes) expiryMinutes = CONSTANTS.SESSION_TOKEN_DEFAULT_EXPIRY_MINUTES; // 3 hour default

  const timestamp = Date.now();
  const expiryTime = timestamp + (expiryMinutes * 60 * 1000);
  const random = Utilities.getUuid(); // Add randomness for security

  const tokenData = {
    email: email.toLowerCase().trim(),
    url: assessmentUrlOrRole, // Can be 'admin', 'super_admin', 'teacher' for staff users
    exp: expiryTime,
    rnd: random,
    role: CONSTANTS.STAFF_ROLES.includes(assessmentUrlOrRole) ? assessmentUrlOrRole : CONSTANTS.ROLE_TOKEN_STUDENT,
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
    if (CONSTANTS.STAFF_ROLES.includes(tokenData.role)) {
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
      const pdfUrl = row[CONSTANTS.COL.PDF_URL];
      const studentEmailsRaw = row[CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();

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
 * Retrieves list of all assessments assigned to a student by email only.
 * @param {string} email Student email
 * @returns {Object} { success: true, assessments: [...] } or { error: "..." }
 */
function getStudentAssessmentsForEmail(email) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return { error: 'Backend Error: "Assessment Database" sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const cleanEmail = email.toLowerCase().trim();
    const matchingAssessments = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[CONSTANTS.COL.PDF_URL];
      const isComplete = row[CONSTANTS.COL.IS_COMPLETE];
      const studentEmailsRaw = row[CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();
      const className = row[CONSTANTS.COL.CLASS_NAME] ? row[CONSTANTS.COL.CLASS_NAME].toString().trim() : '';
      const instructor = row[CONSTANTS.COL.INSTRUCTOR] ? row[CONSTANTS.COL.INSTRUCTOR].toString().trim() : '';

      if (!pdfUrl || isComplete !== true || !studentEmailsRaw) continue;

      const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());

      if (studentEmails.includes(cleanEmail)) {
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
        }
      }
    }

    Logger.log(`Found ${matchingAssessments.length} assessment(s) for ${email}`);

    return {
      success: true,
      assessments: matchingAssessments
    };

  } catch (e) {
    Logger.log(`Error in getStudentAssessmentsForEmail: ${e.toString()}`);
    return { error: 'An unexpected server error occurred.' };
  }
}

/**
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

        // Verify this authenticated user's email is in the list for this assessment.
        if (studentEmails.includes(cleanEmail)) {
          // NEW: Validate the provided password against the one in the sheet.
          if (password !== sheetPassword) {
            return { error: 'Incorrect password for this assessment.' };
          }

          // Password is correct, proceed...
          const audioDataJson = row[CONSTANTS.COL.AUDIO_JSON];
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
          };
        } else {
          // User is trying to access an assessment they are not assigned to.
          return { error: 'You are not authorized to access this assessment.' };
        }
      }
    }
    // If loop finishes, the assessmentUrl was not found.
    return { error: 'Assessment not found.' };
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
function testExtractTextFromFile(testFileId) {
  if (!testFileId) {
    Logger.log('ERROR: No file ID provided for testing.');
    return;
  }
  Logger.log(`=== Testing extractTextFromFile for file ID: ${testFileId} ===`);
  const chunks = extractTextFromFile(testFileId);
  if (chunks) {
    Logger.log(`Extracted ${chunks.length} chunks`);
    chunks.forEach((chunk, i) => {
      Logger.log(`\nChunk ${i+1}: ${chunk}`);
    });
  } else {
    Logger.log('✗ Extraction failed');
  }
}

/**
 * Run all test functions.
 * @param {string} listBasedFileId - File ID for list-based test file
 * @param {string} tableBasedFileId - File ID for table-based test file
 * Usage: runTests('YOUR_LIST_FILE_ID', 'YOUR_TABLE_FILE_ID');
 */
function runTests(listBasedFileId, tableBasedFileId) {
  if (!listBasedFileId || !tableBasedFileId) {
    Logger.log('ERROR: Please provide both file IDs as arguments to runTests.');
    Logger.log('Usage: runTests("YOUR_LIST_FILE_ID", "YOUR_TABLE_FILE_ID")');
    return;
  }

  Logger.log('--- STARTING TEST RUN ---');
  testExtractTextFromFile(listBasedFileId);
  testExtractTextFromFile(tableBasedFileId);
  Logger.log('--- FINISHED TEST RUN ---');
}

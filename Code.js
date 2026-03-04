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
  const batchRequests = textChunks.map((chunkObj, index) => ({
    key: `${fileId}_chunk_${index}`,
    request: {
      contents: [{
        parts: [{
          text: `Read the following text in a clear, neutral, and steady voice: ${chunkObj.text}`
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
        const chunkObj = textChunks[i];
        const chunkText = chunkObj.text;
        const chunkIds = chunkObj.ids;
        
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
          elementIds: chunkIds, // NEW: Add IDs
          searchWords: searchWords,
          audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
          audioFilename: audioFile.getName()
        });
      }
    }

    // Save final JSON
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    setLargeDataInCell(sheet.getRange(rowIndex, CONSTANTS.COL.AUDIO_JSON + 1), JSON.stringify(audioFileObjects, null, 2), fileName);

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

  // Get existing IDs to prevent duplicates
  const data = sheet.getDataRange().getValues();
  const existingIds = new Set();
  for (let i = 1; i < data.length; i++) {
    const url = data[i][CONSTANTS.COL.PDF_URL];
    if (url) {
      const id = getFileIdFromUrl(url);
      if (id) existingIds.add(id);
    }
  }

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
      const fileId = file.id;
      const mimeType = file.mimeType;
      const fileName = file.name;

      if (!existingIds.has(fileId)) {
        sheet.appendRow([fileUrl]);
        Logger.log(`Added new file: ${fileName} (${mimeType})`);
        addedCount++;
        existingIds.add(fileId); // Prevent adding same file twice in one run
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
    const processingStatus = data[i][CONSTANTS.COL.PROCESSING_STATUS];

    // Only analyze if chunkCount is unset and no processing has started
    if (pdfUrl && chunkCount === '' && !processingStatus) {
      const readAloudEnabled = data[i][CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
      if (!readAloudEnabled) {
        markAssessmentAsNoAudioRequired(sheet, i + 1, pdfUrl);
        continue;
      }

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
        // Clear cached HTML to force regeneration if chunking changed
        setLargeDataInCell(sheet.getRange(i + 1, CONSTANTS.COL.ASSESSMENT_HTML + 1), '');
        Logger.log(`--> Found ${textChunks.length} chunks. Updated sheet and cleared HTML cache.`);
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
         const chunkObj = textChunks[j];
         const chunkText = chunkObj.text;
         const chunkIds = chunkObj.ids;
         
         // --- NEW: Generate the descriptive filename ---
         const newChunkName = generateSafeFilenameFromText(chunkText, j);

         // ONLY look for the descriptive filename that includes the text representation
         // This prevents picking up old, mismatched audio files if chunking changes
         let existingFiles = assessmentSubfolder.getFilesByName(newChunkName);

         let audioFile = null;

         if (existingFiles.hasNext()) {
            audioFile = existingFiles.next();
         } else {
            Logger.log(`--> Generating new audio for chunk ${j + 1} with name "${newChunkName}"...`);
            // Always generate new files with the new descriptive name
            audioFile = generateAudio(chunkText, newChunkName, assessmentSubfolder);
         }

         if (audioFile) {
            // Store file plus the IDs for this chunk
            audioFileObjects.push({ file: audioFile, ids: chunkIds });
         } else {
            Logger.log(`--> FAILED to process chunk ${j + 1}. Will retry on next run.`);
            allChunksProcessed = false;
            break; 
         }
      }
      
      if (allChunksProcessed && audioFileObjects.length === totalChunks) {
        Logger.log(`--> All ${totalChunks} audio chunks accounted for. Finalizing...`);
        
        // Generate and cache HTML for fast student loading
        Logger.log(`--> Generating and caching HTML for '${fileName}'...`);
        const conversionResult = convertFileToHtml(fileId);
        if (!conversionResult.error) {
          const sanitizedHtml = sanitizeHtml(conversionResult.html);
          setLargeDataInCell(sheet.getRange(i + 1, CONSTANTS.COL.ASSESSMENT_HTML + 1), sanitizedHtml, fileName + "_html");
        } else {
          Logger.log(`--> WARNING: Failed to cache HTML: ${conversionResult.error}`);
        }

        const audioDataForSheet = [];
        for(let j = 0; j < totalChunks; j++) {
           const chunkText = textChunks[j].text;
           const chunkData = audioFileObjects[j];
           const audioFile = chunkData.file;
           
           // Generate searchWords: first 8 words to match frontend display
           const words = chunkText.trim().split(/\s+/);
           const searchWords = words.slice(0, CONSTANTS.SEARCH_WORDS_COUNT).join(' ') + (words.length > CONSTANTS.SEARCH_WORDS_COUNT ? '...' : '');

           audioDataForSheet.push({
             text: chunkText,
             elementIds: chunkData.ids, // NEW: Add IDs for frontend mapping
             searchWords: searchWords,
             audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
             audioFilename: audioFile.getName()
           });
        }
        setLargeDataInCell(sheet.getRange(i + 1, CONSTANTS.COL.AUDIO_JSON + 1), JSON.stringify(audioDataForSheet, null, 2), fileName);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
        sheet.getRange(i + 1, CONSTANTS.COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
        Logger.log(`--> Successfully created JSON, cached HTML, and marked as complete.`);
      } else {
        Logger.log(`--> Process for '${fileName}' partially complete. Will resume on next run.`);
      }
    }
  }
  SpreadsheetApp.flush();
  Logger.log('Step 2 processing finished.');
  
  // Clean up any immediate triggers that might have called this
  cleanupImmediateTriggers();
}

/**
 * Cleans up any one-time immediate triggers for step2_GenerateMissingAudioAndFinalize.
 */
function cleanupImmediateTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'step2_GenerateMissingAudioAndFinalize' && 
        trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      // We only want to delete it if it's not a recurring trigger
      // Note: CLOCK triggers created with .after() are not recurring
      try {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('Cleaned up one-time background trigger');
      } catch (e) {
        // Ignore errors if trigger already deleted
      }
    }
  });
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

    // Embed images as base64 data URIs to fix authentication issues
    Logger.log('→ Embedding images as base64');
    htmlContent = embedImagesAsBase64(htmlContent);

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
 * Embeds images in HTML as base64 data URIs by downloading them from Google CDN.
 * This fixes broken images that require OAuth authentication.
 * @param {string} html The HTML content containing image URLs
 * @returns {string} HTML with images embedded as base64 data URIs
 */
function embedImagesAsBase64(html) {
  try {
    Logger.log('→ Scanning HTML for images to embed');

    const token = ScriptApp.getOAuthToken();
    let modifiedHtml = html;
    let imageCount = 0;
    let failedCount = 0;

    // Find all <img> tags with src attributes
    const imgRegex = /<img[^>]+src=["']([^"']+)["'][^>]*>/gi;
    const matches = [];
    let match;

    // Collect all matches first
    while ((match = imgRegex.exec(html)) !== null) {
      matches.push({
        fullTag: match[0],
        url: match[1]
      });
    }

    Logger.log(`→ Found ${matches.length} images to process`);

    // Process each image
    for (let i = 0; i < matches.length; i++) {
      try {
        const imgMatch = matches[i];
        const imageUrl = imgMatch.url;

        // Skip if already a data URI
        if (imageUrl.startsWith('data:')) {
          Logger.log(`→ [Image ${i+1}/${matches.length}] Already embedded, skipping`);
          continue;
        }

        Logger.log(`→ [Image ${i+1}/${matches.length}] Processing: ${imageUrl.substring(0, 80)}...`);

        // Try to download with retries
        const blob = downloadImageWithRetry(imageUrl, token, 3);

        if (blob) {
          const base64Data = Utilities.base64Encode(blob.getBytes());
          const mimeType = blob.getContentType() || 'image/png';
          const dataUri = `data:${mimeType};base64,${base64Data}`;

          // Replace the URL in the original tag
          const newTag = imgMatch.fullTag.replace(imageUrl, dataUri);
          modifiedHtml = modifiedHtml.replace(imgMatch.fullTag, newTag);

          imageCount++;
          Logger.log(`✓ [Image ${i+1}/${matches.length}] SUCCESS: Embedded (${mimeType}, ${Math.round(base64Data.length / 1024)}KB)`);
        } else {
          failedCount++;
          Logger.log(`✗ [Image ${i+1}/${matches.length}] FAILED: Could not download after retries`);
        }

      } catch (imgError) {
        failedCount++;
        Logger.log(`✗ [Image ${i+1}/${matches.length}] ERROR: ${imgError.toString()}`);
      }
    }

    Logger.log(`✓ Image embedding complete: ${imageCount} succeeded, ${failedCount} failed out of ${matches.length} total`);

    if (failedCount > 0) {
      Logger.log(`⚠ WARNING: ${failedCount} image(s) failed to embed and will appear broken in the student view`);
    }

    return modifiedHtml;

  } catch (e) {
    Logger.log(`✗ Exception in embedImagesAsBase64: ${e.toString()}`);
    // Return original HTML if embedding fails
    return html;
  }
}

/**
 * Downloads an image with retry logic and multiple authentication methods.
 * @param {string} url The image URL
 * @param {string} token OAuth token
 * @param {number} maxRetries Maximum number of retry attempts
 * @returns {GoogleAppsScript.Base.Blob|null} The image blob or null if failed
 */
function downloadImageWithRetry(url, token, maxRetries) {
  const authMethods = [
    // Method 1: OAuth Bearer token
    { name: 'OAuth Bearer', headers: { 'Authorization': `Bearer ${token}` } },
    // Method 2: No authentication (for public images)
    { name: 'No Auth', headers: {} },
    // Method 3: OAuth with cookies
    { name: 'OAuth with cookies', headers: { 'Authorization': `Bearer ${token}`, 'Cookie': '' } }
  ];

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    for (const method of authMethods) {
      try {
        Logger.log(`  → Attempt ${attempt}/${maxRetries} using ${method.name}`);

        const response = UrlFetchApp.fetch(url, {
          headers: method.headers,
          muteHttpExceptions: true,
          followRedirects: true
        });

        const responseCode = response.getResponseCode();
        Logger.log(`  → HTTP ${responseCode}`);

        if (responseCode === 200) {
          const blob = response.getBlob();
          const contentType = blob.getContentType();

          // Verify it's actually an image
          if (contentType && contentType.startsWith('image/')) {
            Logger.log(`  → Success with ${method.name}`);
            return blob;
          } else {
            Logger.log(`  → Unexpected content type: ${contentType}`);
          }
        } else if (responseCode === 302 || responseCode === 301) {
          // Handle redirects manually
          const redirectUrl = response.getHeaders()['Location'];
          if (redirectUrl) {
            Logger.log(`  → Following redirect to: ${redirectUrl.substring(0, 60)}...`);
            return downloadImageWithRetry(redirectUrl, token, 1); // One retry for redirect
          }
        }

      } catch (fetchError) {
        Logger.log(`  → ${method.name} failed: ${fetchError.toString()}`);
      }
    }

    // Wait before retry (exponential backoff)
    if (attempt < maxRetries) {
      const waitMs = Math.pow(2, attempt) * 100; // 200ms, 400ms, 800ms
      Logger.log(`  → Waiting ${waitMs}ms before retry...`);
      Utilities.sleep(waitMs);
    }
  }

  return null;
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
    // Keep only src and alt attributes, handling both single and double quotes
    const srcMatch = attrs.match(/src=(?:"([^"]*)"|'([^']*)')/i);
    const altMatch = attrs.match(/alt=(?:"([^"]*)"|'([^']*)')/i);
    const src = srcMatch ? (srcMatch[1] || srcMatch[2] || '') : '';
    let alt = altMatch ? (altMatch[1] || altMatch[2] || '') : '';

    // Ensure alt is never undefined or just whitespace - provide default
    if (!alt || alt.trim() === '') {
      alt = 'Image';
    }

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

  // NEW: Split merged paragraphs (common in PDF extraction)
  // Fixes "text b." and "text28. Which" issues
  sanitized = sanitized.replace(/<p([^>]*)>([\s\S]*?)<\/p>/gi, (match, attrs, content) => {
    // 1. Split on answer options in the middle of text: "text b. more text"
    // Matches a space, then an option marker like a. b. c. d. e.
    let splitContent = content.replace(/(\s)([a-e][.)]\s+)/g, '</p><p$1>$2');
    
    // 2. Split on question numbers merged with text: "text28. Which" or "text 28. Which"
    // Matches text followed by optional space, then a number and period/parenthesis
    splitContent = splitContent.replace(/([a-zA-Z])(\s?\d+[.)\]]\s+)/g, '$1</p><p$1>$2');
    
    if (splitContent !== content) {
      return `<p${attrs}>${splitContent}</p>`;
    }
    return match;
  });

  // NEW: Ensure ALL table cell content is wrapped in <p> if it's not already block-level
  // This version handles mixed content (text + block elements) more robustly
  sanitized = sanitized.replace(/<(td|th)([^>]*)>([\s\S]*?)<\/\1>/gi, (match, tag, attrs, content) => {
    // If there are no block tags at all, wrap the whole content
    if (!/<(p|div|ul|ol|table|h[1-6])/i.test(content)) {
      return `<${tag}${attrs}><p>${content}</p></${tag}>`;
    }

    // If there ARE block tags, we need to wrap "orphan" text nodes
    // Split by block tags, keeping the tags in the array
    const parts = content.split(/(<p[^>]*>[\s\S]*?<\/p>|<div[^>]*>[\s\S]*?<\/div>|<ul[^>]*>[\s\S]*?<\/ul>|<ol[^>]*>[\s\S]*?<\/ol>|<table[^>]*>[\s\S]*?<\/table>|<h[1-6][^>]*>[\s\S]*?<\/h[1-6]>)/gi);
    
    const wrappedParts = parts.map(part => {
      if (!part.trim()) return ''; // Ignore empty/whitespace-only parts
      if (part.startsWith('<')) return part; // Already a block element
      return `<p>${part.trim()}</p>`; // Wrap orphan text
    });

    return `<${tag}${attrs}>${wrappedParts.join('')}</${tag}>`;
  });

  // 12. Assign unique IDs to all block elements for precise TTS mapping
  // This enables the frontend to highlight exactly what is being read
  let blockCounter = 0;
  // EXCLUDE div and blockquote to avoid nested tag parsing issues with simple regex
  const blockTags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li'];
  const blockRegex = new RegExp(`<(${blockTags.join('|')})([^>]*)>`, 'gi');

  sanitized = sanitized.replace(blockRegex, (match, tag, attrs) => {
    // Remove any existing ID to ensure our sequential ones are used consistently
    const cleanAttrs = attrs.replace(/\bid=["'][^"']*["']\s?/gi, '').trim();
    const space = cleanAttrs ? ' ' : '';
    return `<${tag}${space}${cleanAttrs} id="sra-block-${blockCounter++}">`;
  });

  // 13. Add question-start class for visual separation
  // Matches: <p id="..." > 1.  or <p id="..." > 1) or <p id="..." > Question 1.
  // Enhanced to match "Question 1.", "Q1.", "Section 1.", etc.
  sanitized = sanitized.replace(/(<p[^>]*>)(?:\s|&nbsp;)*((?:(?:Question|Q|Section|Part)\s*)?\d+[.)\]]\s+)/gi, '$1<span class="question-marker">$2</span>');
  sanitized = sanitized.replace(/<p([^>]*id="sra-block-[^"]*"[^>]*)>(?:\s|&nbsp;)*(?=<span class="question-marker">)/gi, '<p$1 class="question-start">');

  // 14. Add question-text class to paragraphs that end in a question mark
  // This catches questions that don't start with a number (e.g. "What is the capital of France?")
  // but we explicitly exclude answer patterns (e.g. "a) Answer?")
  sanitized = sanitized.replace(/<p([^>]*id="sra-block-[^"]*"[^>]*)>([\s\S]*?\?\s*)<\/p>/gi, (match, attrs, content) => {
    // If it's already a question start, don't double up (though CSS would handle it)
    if (attrs.includes('question-start')) return match;
    
    // Check if the content (stripped of HTML tags) starts with an answer choice pattern
    const plainText = content.replace(/<[^>]+>/g, '').trim();
    const answerPattern = /^\s*(?:\([a-dA-D]\)|[a-dA-D][.)])\s+/i;
    
    if (answerPattern.test(plainText)) {
      return match; // It's an answer option that happens to end in a question mark
    }
    
    // It's likely a question without a number
    const space = attrs.trim() ? ' ' : '';
    return `<p${attrs}${space}class="question-text">${content}</p>`;
  });

  Logger.log(`Sanitized & normalized HTML: ${html.length} chars → ${sanitized.length} chars (Added ${blockCounter} IDs)`);
  return sanitized;
}

/**
 * Parses HTML with 'sra-block-N' IDs into structured TTS chunks.
 * Uses improved logic to group questions with answers and merge flowing text.
 * @param {string} html The sanitized HTML with IDs
 * @returns {Array<{text: string, ids: string[]}>} Array of chunk objects
 */
function parseHtmlToChunks(html) {
  const chunks = [];
  
  // Extract all blocks with their text and ID using regex
  // Only target the specific tags we added IDs to in sanitizeHtml
  const blockPattern = /<(p|h[1-6]|li)[^>]*id="(sra-block-\d+)"[^>]*>([\s\S]*?)<\/\1>/gi;
  
  let match;
  const blocks = [];
  
  while ((match = blockPattern.exec(html)) !== null) {
    const tag = match[1].toLowerCase();
    const id = match[2];
    let content = match[3];
    
    // Strip tags from content to get plain text for analysis
    const plainText = content
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
      
    if (plainText) {
      blocks.push({ id, tag, text: plainText });
    }
  }
  
  Logger.log(`Found ${blocks.length} content blocks for chunking`);
  
  if (blocks.length === 0) return [];

  let currentChunk = { text: '', ids: [] };
  
  // Helper to finish current chunk and start a new one
  const commitChunk = () => {
    if (currentChunk.text.trim()) {
      currentChunk.text = currentChunk.text.trim();
      chunks.push(currentChunk);
    }
    currentChunk = { text: '', ids: [] };
  };

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    
    // --- Detection Logic ---
    
    // 1. Question Start: "1.", "1)", "Q1", "(1)"
    const isQuestionStart = /^(?:\(?\d+|Q\d+)[.)\]]/.test(block.text);
    
    // 2. Answer Option: "a.", "b.", "A)", "(a)", "A. ", etc.
    // Improved regex: handles optional leading space, optional parentheses, single letter, period or closing paren, and trailing space or end of string
    const isAnswerOption = /^\s*\(?[a-zA-Z][.)]\s*(?:\s|$)/.test(block.text);
    
    // 3. Header
    const isHeader = /^h[1-6]/.test(block.tag);
    
    // 4. Metadata/Directions (e.g., "Directions:", "Read the following...")
    const isDirections = /^(directions|instructions|read|note):/i.test(block.text);

    // --- Grouping Decision Matrix ---

    if (currentChunk.text === '') {
      // Start of new chunk
      currentChunk.text = block.text;
      currentChunk.ids.push(block.id);
    } 
    else if (isAnswerOption) {
      // Rule: Merge answer options with the preceding chunk (likely the question or previous option)
      // This MUST come before isQuestionStart to ensure options are grouped even if the question was long
      currentChunk.text += '\n' + block.text;
      currentChunk.ids.push(block.id);
    }
    else if (isQuestionStart || isHeader) {
      // Rule: Questions and Headers usually start a NEW thought/context
      // Priority 1: Force break before a new question or header
      commitChunk();
      currentChunk.text = block.text;
      currentChunk.ids.push(block.id);
    }
    else if (isDirections) {
      // Rule: Directions usually precede content. 
      commitChunk();
      currentChunk.text = block.text;
      currentChunk.ids.push(block.id);
    }
    else {
      // Rule: Standard Paragraph / Text Continuation
      
      // Check if previous chunk ended with a sentence terminator
      const sentenceEndRegex = /[.!?]"?$/;
      const previousEndedSentence = sentenceEndRegex.test(currentChunk.text.trim());
      
      // Check length of previous chunk
      const isPreviousShort = currentChunk.text.length < 150;
      
      // Check if current block is a continuation (sentence fragment flow)
      // or if previous was just a short intro line
      if (!previousEndedSentence || isPreviousShort) {
         // Merge for flow
         currentChunk.text += ' ' + block.text; // Use space for flowing text
         currentChunk.ids.push(block.id);
      } else {
         // Previous was a complete thought and long enough -> Start new
         commitChunk();
         currentChunk.text = block.text;
         currentChunk.ids.push(block.id);
      }
    }
  }
  
  commitChunk(); // Final commit
  
  return chunks;
}


function getFileIdFromUrl(url) {
    if (!url) return null;
    // Try document URL format first: /d/FILE_ID/
    let match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match) return match[1];

    // Try audio URL format: id=FILE_ID
    match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
}

/**
 * Checks if two Drive URLs point to the same file.
 * @param {string} url1 First URL
 * @param {string} url2 Second URL
 * @returns {boolean} True if they have the same ID
 */
function areUrlsSameFile(url1, url2) {
  if (url1 === url2) return true;
  if (!url1 || !url2) return false;
  const id1 = getFileIdFromUrl(url1);
  const id2 = getFileIdFromUrl(url2);
  return id1 !== null && id2 !== null && id1 === id2;
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

    // --- PDF Handling: Fall through to HTML conversion ---
    // We previously had a separate OCR text-only path here. 
    // We now route PDFs through convertFileToHtml -> parseHtmlToChunks 
    // to ensure we generate the element IDs required for highlighting.

    // --- Google Docs / Word / PDF Handling ---
    Logger.log('→ Using HTML conversion for text extraction');
    const conversionResult = convertFileToHtml(fileId);
    if (conversionResult.error) {
      Logger.log(`✗ Failed to convert file: ${conversionResult.error}`);
      return null;
    }
    
    // NEW: Use sanitizeHtml to ensure we work with the exact same structure the student sees
    // This adds the IDs and normalizes lists
    const htmlContent = sanitizeHtml(conversionResult.html);
    Logger.log(`→ Sanitized HTML length: ${htmlContent.length} chars`);

    // NEW: Parse HTML into structured chunks (grouping questions with answers)
    const structuredChunks = parseHtmlToChunks(htmlContent);
    
    Logger.log(`✓ Extracted ${structuredChunks.length} structured chunks from HTML`);
    
    // For now, we return just the text to match existing contract
    // In Step 2, we will access the full object (text + IDs)
    // We store the full structured data in a cache or property? 
    // actually, extractTextFromFile is called in Step 2 again. 
    // So we can just return the objects and let the caller handle it?
    // Existing callers expect array of strings. Let's attach the metadata property to the string object
    // or just return the array of objects and update the callers.
    // Let's update the callers.
    
    return structuredChunks;

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
  if (!text) return '';
  text = text.trim();

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
      .replace(/\'/g, '&apos;');

    // Step 2: Add pause after question number at the very start (e.g., "1. ", "2) ", "3] ")
    ssmlText = ssmlText.replace(/^(\d+[.)\]])\s+/, `$1 <break time="${CONSTANTS.PAUSE_AFTER_QUESTION_NUMBER_MS}ms"/> `);

    // Step 3: Add pauses after paragraph breaks (double line breaks or more)
    ssmlText = ssmlText.replace(/\n\n+/g, `\n<break time="${CONSTANTS.PAUSE_AFTER_PARAGRAPH_MS}ms"/>\n`);

    // Step 4: Add pause between question and first answer choice
    // Matches a newline that is immediately followed by "a.", "b)", etc.
    ssmlText = ssmlText.replace(/\n(?=[a-dA-D][.)])/g, ` <break time="${CONSTANTS.PAUSE_BEFORE_ANSWER_BLOCK_MS}ms"/>\n`);

    // Step 4b: Add pause before inline answer choices (e.g. "a. Paris b. Madrid")
    // Use word boundary to ensure we only match standalone choice markers
    ssmlText = ssmlText.replace(/(^|[^\n])[ \t]+(?=\b[a-dA-D][.)])/g, `$1 <break time="${CONSTANTS.PAUSE_BEFORE_INLINE_ANSWER_MS}ms"/> `);

    // Step 5: Add pauses after answer choices (A., B., C., D. or a), b), c), d))
    // Matches: "A." or "A)" (uppercase or lowercase, periods or parentheses), with optional whitespace
    // ENHANCED: Wrap the choice letter in <say-as interpret-as="characters"> to ensure correct pronunciation
    ssmlText = ssmlText.replace(/(\b)([A-Da-d])([.)])\s*/g, (match, boundary, letter, punctuation) => {
      const upperLetter = letter.toUpperCase();
      return `${boundary}<say-as interpret-as="characters">${upperLetter}</say-as>${punctuation} <break time="${CONSTANTS.PAUSE_AFTER_ANSWER_CHOICE_MS}ms"/> `;
    });

    // Step 6: Wrap in SSML speak tags
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
/**
 * Retrieves all assessments from the spreadsheet, filtered by user role.
 * Includes performance optimizations: CacheService and local file name storage.
 * @param {string} sessionToken Staff session token
 * @param {boolean} fetchAll If true, returns the full history; if false, limits to most recent 20 (for dashboard)
 * @returns {Object} List of assessments and user role
 */
function getAllAssessments(sessionToken, fetchAll = false) {
  try {
    // Verify staff token (teacher, admin, or super_admin)
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) {
      return { error: 'Unauthorized. Staff access required.' };
    }

    const cacheVersion = PropertiesService.getScriptProperties().getProperty('CACHE_VERSION_ASSESSMENTS') || '1';
    const cacheKey = `assessments_v2_${cacheVersion}_${tokenData.email}_${tokenData.role}`;
    const cache = CacheService.getScriptCache();
    let cachedData = cache.get(cacheKey);

    let assessments = [];
    if (cachedData) {
      try {
        assessments = JSON.parse(cachedData);
        Logger.log(`[CACHE] Retrieved ${assessments.length} assessments from cache for ${tokenData.email}`);
      } catch (e) {
        Logger.log(`[CACHE] Parse error: ${e.toString()}`);
        cachedData = null;
      }
    }

    if (!cachedData) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
      if (!sheet) {
        return { error: 'Assessment Database sheet not found.' };
      }

      const data = sheet.getDataRange().getValues();
      
      // Skip header row
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const pdfUrl = row[CONSTANTS.COL.PDF_URL];

        if (!pdfUrl) continue; // Skip empty rows

        // PERFORMANCE: Read filename from local column instead of Drive API
        let fileName = row[CONSTANTS.COL.FILE_NAME] || '';
        
        // Fallback for legacy rows (only run if name is truly missing)
        if (!fileName) {
          try {
            const fileId = getFileIdFromUrl(pdfUrl);
            if (fileId) {
              fileName = DriveApp.getFileById(fileId).getName();
              // Back-fill the sheet so it's faster next time
              sheet.getRange(i + 1, CONSTANTS.COL.FILE_NAME + 1).setValue(fileName);
            }
          } catch (e) {
            fileName = 'Unknown file';
          }
        }

        assessments.push({
          rowIndex: i,
          fileName: fileName,
          pdfUrl: pdfUrl,
          chunkCount: row[CONSTANTS.COL.CHUNK_COUNT] || 0,
          hasAudio: !!row[CONSTANTS.COL.AUDIO_JSON],
          isComplete: row[CONSTANTS.COL.IS_COMPLETE] === true,
          className: row[CONSTANTS.COL.CLASS_NAME] || '',
          instructor: row[CONSTANTS.COL.INSTRUCTOR] || '',
          password: row[CONSTANTS.COL.PASSWORD] || '',
          studentEmails: row[CONSTANTS.COL.STUDENT_EMAILS] || '',
          readAloudEnabled: row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false,
          submissionEnabled: row[CONSTANTS.COL.SUBMISSION_ENABLED] === true,
          submissionDeliveryMode: row[CONSTANTS.COL.SUBMISSION_DELIVERY_MODE] || 'email',
          hasSubmissions: !!(row[CONSTANTS.COL.SUBMISSION_TIMESTAMPS] && row[CONSTANTS.COL.SUBMISSION_TIMESTAMPS].length > 2),
          accessExpires: row[CONSTANTS.COL.ACCESS_EXPIRES] ? Utilities.formatDate(new Date(row[CONSTANTS.COL.ACCESS_EXPIRES]), "America/Chicago", "yyyy-MM-dd'T'HH:mm") : ''
        });
      }

      // Filter assessments for teachers (only show their own)
      if (tokenData.role === CONSTANTS.ROLE_TOKEN_TEACHER) {
        const teacherName = (tokenData.name || '').toLowerCase();
        const teacherEmail = (tokenData.email || '').toLowerCase();
        assessments = assessments.filter(assessment => {
          const instructorField = (assessment.instructor || '').toLowerCase();
          const instructors = instructorField.split(/[,\/]/).map(name => name.trim()).filter(name => name);
          if (instructors.length === 0) return false;
          return instructors.some(instructorPart => 
            teacherName.includes(instructorPart) || 
            teacherEmail.includes(instructorPart) ||
            instructorPart.includes(teacherEmail)
          );
        });
      }

      // Filter assessments for Sp.Ed. users
      if (tokenData.role === CONSTANTS.ROLE_TOKEN_SPED) {
        const teacherName = (tokenData.name || '').toLowerCase();
        const teacherEmail = (tokenData.email || '').toLowerCase();
        const caseloadStudents = getStudentsByCaseManager(tokenData.email);
        assessments = assessments.filter(assessment => {
          const instructorField = (assessment.instructor || '').toLowerCase();
          const instructors = instructorField.split(/[,\/]/).map(name => name.trim()).filter(name => name);
          const isInstructor = instructors.some(inst => 
            teacherName.includes(inst) || teacherEmail.includes(inst) || inst.includes(teacherEmail)
          );
          if (isInstructor) return true;
          const studentEmailsField = (assessment.studentEmails || '').toLowerCase();
          return caseloadStudents.some(studentEmail => studentEmailsField.includes(studentEmail));
        });
      }

      // Sort assessments by rowIndex descending (most recent first)
      assessments.sort((a, b) => b.rowIndex - a.rowIndex);

      // Save to cache (limit string size to avoid CacheService errors)
      try {
        const jsonStr = JSON.stringify(assessments);
        if (jsonStr.length < 95000) { // 100KB limit
          cache.put(cacheKey, jsonStr, 900); // 15 minutes
        }
      } catch (cacheError) {
        Logger.log(`[CACHE] Failed to save to cache: ${cacheError.toString()}`);
      }
    }

    const totalAvailable = assessments.length;
    const readyCount = assessments.filter(a => a.isComplete).length;
    const processingCount = assessments.filter(a => a.chunkCount > 0 && !a.isComplete).length;
    
    // Slice if fetchAll is false (recent 20)
    if (!fetchAll && assessments.length > 20) {
      assessments = assessments.slice(0, 20);
    }

    return {
      success: true,
      assessments: assessments,
      totalCount: totalAvailable,
      readyCount: readyCount,
      processingCount: processingCount,
      isTruncated: !fetchAll && totalAvailable > 20,
      userRole: tokenData.role
    };

  } catch (e) {
    Logger.log(`Error in getAllAssessments: ${e.toString()}`);
    return { error: 'Failed to retrieve assessments.' };
  }
}

/**
 * Checks if a user is authorized to modify or delete a specific assessment.
 * Super Admins and Admins can access everything.
 * Teachers can only access assessments where they are listed as an instructor.
 * @param {Object} tokenData Decoded session token data
 * @param {Array} rowData The spreadsheet row data for the assessment
 * @returns {boolean} True if authorized
 */
function isAuthorizedForAssessment(tokenData, rowData) {
  if (tokenData.role === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN || tokenData.role === CONSTANTS.ROLE_TOKEN_ADMIN) {
    return true;
  }
  
  if (tokenData.role === CONSTANTS.ROLE_TOKEN_TEACHER) {
    const instructorField = (rowData[CONSTANTS.COL.INSTRUCTOR] || '').toLowerCase();
    const teacherName = (tokenData.name || '').toLowerCase();
    const teacherEmail = (tokenData.email || '').toLowerCase();
    
    // Split by comma or slash
    const instructors = instructorField.split(/[,\/]/).map(name => name.trim()).filter(name => name);
    
    if (instructors.length === 0) return false;
    
    // Authorized if instructor field contains teacher's name or email
    return instructors.some(inst => 
      (teacherName && inst.includes(teacherName)) || 
      (teacherEmail && inst.includes(teacherEmail)) ||
      (teacherName && teacherName.includes(inst))
    );
  }

  if (tokenData.role === CONSTANTS.ROLE_TOKEN_SPED) {
    // 1. Check if they are the instructor
    const instructorField = (rowData[CONSTANTS.COL.INSTRUCTOR] || '').toLowerCase();
    const teacherName = (tokenData.name || '').toLowerCase();
    const teacherEmail = (tokenData.email || '').toLowerCase();
    const instructors = instructorField.split(/[,\/]/).map(name => name.trim()).filter(name => name);
    
    if (instructors.some(inst => 
      (teacherName && inst.includes(teacherName)) || 
      (teacherEmail && inst.includes(teacherEmail)) ||
      (teacherName && teacherName.includes(inst))
    )) {
      return true;
    }

    // 2. Check caseload (original Sp.Ed. logic)
    const caseloadStudents = getStudentsByCaseManager(tokenData.email);
    const studentEmailsField = (rowData[CONSTANTS.COL.STUDENT_EMAILS] || '').toLowerCase();
    
    return caseloadStudents.some(studentEmail => studentEmailsField.includes(studentEmail));
  }
  
  return false;
}

/**
 * Updates an assessment row in the spreadsheet.
 * Staff-only function. Teachers can only update their own assessments.
 * @param {string} sessionToken Staff session token
 * @param {number} rowIndex Row index (1-based, excluding header)
 * @param {Object} data Data to update { className, instructor, password, studentEmails }
 * @returns {Object} { success: true } or { error: "..." }
 */
function updateAssessmentRow(sessionToken, rowIndex, data) {
  try {
    // Verify staff token
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) {
      return { error: 'Unauthorized. Staff access required.' };
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

    // Check ownership for teachers
    const rowData = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!isAuthorizedForAssessment(tokenData, rowData)) {
      return { error: 'Unauthorized. You can only update your own assessments.' };
    }

    if (data.assessmentTitle !== undefined) {
      const fileName = data.assessmentTitle || '';
      sheet.getRange(actualRow, CONSTANTS.COL.FILE_NAME + 1).setValue(fileName);
      
      try {
        const fileId = getFileIdFromUrl(rowData[CONSTANTS.COL.PDF_URL]);
        if (fileId) {
          const file = DriveApp.getFileById(fileId);
          let newName = fileName;
          // Maintain extension if it's a PDF
          if (file.getMimeType() === MimeType.PDF && !newName.toLowerCase().endsWith('.pdf')) {
            newName += '.pdf';
          }
          if (file.getName() !== newName) {
            file.setName(newName);
            Logger.log(`Renamed assessment file to: ${newName}`);
          }
        }
      } catch (renameError) {
        Logger.log(`Warning: Failed to rename file: ${renameError.toString()}`);
      }
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
      if (!data.studentEmails || data.studentEmails.toString().trim() === '') {
        return { error: 'At least one student email is required.' };
      }
      sheet.getRange(actualRow, CONSTANTS.COL.STUDENT_EMAILS + 1).setValue(parseStudentEmails(data.studentEmails));
    }
    if (data.readAloudEnabled !== undefined) {
      sheet.getRange(actualRow, CONSTANTS.COL.READ_ALOUD_ENABLED + 1).setValue(data.readAloudEnabled);
    }
    if (data.submissionEnabled !== undefined) {
      const canManageSubmissions = CONSTANTS.SUBMISSION_FEATURE_ENABLED && 
        (!CONSTANTS.SUBMISSION_ADMIN_ONLY || CONSTANTS.SUBMISSION_ADMIN_ROLES.includes(tokenData.role));
      
      if (canManageSubmissions) {
        sheet.getRange(actualRow, CONSTANTS.COL.SUBMISSION_ENABLED + 1).setValue(data.submissionEnabled);
      }
    }
    if (data.submissionDeliveryMode !== undefined) {
      const canManageSubmissions = CONSTANTS.SUBMISSION_FEATURE_ENABLED && 
        (!CONSTANTS.SUBMISSION_ADMIN_ONLY || CONSTANTS.SUBMISSION_ADMIN_ROLES.includes(tokenData.role));
      
      if (canManageSubmissions) {
        sheet.getRange(actualRow, CONSTANTS.COL.SUBMISSION_DELIVERY_MODE + 1).setValue(data.submissionDeliveryMode);
      }
    }
    if (data.accessExpires !== undefined) {
      const expiryDate = data.accessExpires ? new Date(data.accessExpires.replace('T', ' ')) : '';
      sheet.getRange(actualRow, CONSTANTS.COL.ACCESS_EXPIRES + 1).setValue(expiryDate);
    }

    SpreadsheetApp.flush();
    
    // Invalidate dashboard cache
    invalidateAssessmentsCache();
    
    Logger.log(`Updated assessment at row ${rowIndex}`);

    return { success: true };

  } catch (e) {
    Logger.log(`Error in updateAssessmentRow: ${e.toString()}`);
    return { error: 'Failed to update assessment.' };
  }
}

/**
 * Deletes an assessment row from the spreadsheet.
 * Staff-only function. Teachers can only delete their own assessments.
 * @param {string} sessionToken Staff session token
 * @param {number} rowIndex Row index (1-based, excluding header)
 * @returns {Object} { success: true } or { error: "..." }
 */
function deleteAssessmentRow(sessionToken, rowIndex) {
  try {
    // Verify staff token
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) {
      return { error: 'Unauthorized. Staff access required.' };
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

    // Check ownership for teachers
    const rowData = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!isAuthorizedForAssessment(tokenData, rowData)) {
      return { error: 'Unauthorized. You can only delete your own assessments.' };
    }

    // 1. Get File URL to trash the file (prevent re-discovery)
    const pdfUrl = rowData[CONSTANTS.COL.PDF_URL];
    
    if (pdfUrl) {
      try {
        const fileId = getFileIdFromUrl(pdfUrl);
        if (fileId) {
          DriveApp.getFileById(fileId).setTrashed(true);
          Logger.log(`Trashed source file for row ${rowIndex} (ID: ${fileId})`);
        }
      } catch (fileError) {
        Logger.log(`Warning: Could not trash file for row ${rowIndex}: ${fileError.toString()}`);
        // Continue with row deletion even if file trash fails
      }
    }

    sheet.deleteRow(actualRow);
    SpreadsheetApp.flush();
    
    // Invalidate dashboard cache
    invalidateAssessmentsCache();
    
    Logger.log(`Deleted assessment at row ${rowIndex} by ${tokenData.email}`);

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
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) {
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

    // Check if file already exists in the database
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (areUrlsSameFile(data[i][CONSTANTS.COL.PDF_URL], fileUrl)) {
        return { error: 'This file is already in the database.' };
      }
    }

    // Rename file in Drive if title provided
    if (metadata.assessmentTitle) {
      try {
        const file = DriveApp.getFileById(fileId);
        // Ensure it keeps its extension if it's a PDF
        let newName = metadata.assessmentTitle;
        if (file.getMimeType() === MimeType.PDF && !newName.toLowerCase().endsWith('.pdf')) {
          newName += '.pdf';
        }
        file.setName(newName);
      } catch (e) {
        Logger.log(`Warning: Could not rename file: ${e.toString()}`);
      }
    }

    // Add new row with file URL and metadata
    const newRow = new Array(19).fill(''); // 19 columns (0-18)
    newRow[CONSTANTS.COL.PDF_URL] = fileUrl;
    newRow[CONSTANTS.COL.CLASS_NAME] = metadata.className || '';
    newRow[CONSTANTS.COL.INSTRUCTOR] = metadata.instructor || '';
    newRow[CONSTANTS.COL.PASSWORD] = metadata.password || '';
    newRow[CONSTANTS.COL.STUDENT_EMAILS] = parseStudentEmails(metadata.studentEmails || '');
    newRow[CONSTANTS.COL.READ_ALOUD_ENABLED] = metadata.readAloudEnabled !== false;
    newRow[CONSTANTS.COL.ACCESS_EXPIRES] = metadata.accessExpires ? new Date(metadata.accessExpires.replace('T', ' ')) : '';
    newRow[CONSTANTS.COL.FILE_NAME] = metadata.assessmentTitle || ''; // Column 18: Display Name
    
    const canManageSubmissions = CONSTANTS.SUBMISSION_FEATURE_ENABLED && 
      (!CONSTANTS.SUBMISSION_ADMIN_ONLY || CONSTANTS.SUBMISSION_ADMIN_ROLES.includes(tokenData.role));
    newRow[CONSTANTS.COL.SUBMISSION_ENABLED] = canManageSubmissions && metadata.submissionEnabled === true;
    newRow[CONSTANTS.COL.SUBMISSION_DELIVERY_MODE] = (canManageSubmissions && metadata.submissionDeliveryMode) || 'email';

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    // Invalidate dashboard cache
    invalidateAssessmentsCache();

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
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];
    const processingStatus = data[i][CONSTANTS.COL.PROCESSING_STATUS];

    if (areUrlsSameFile(data[i][CONSTANTS.COL.PDF_URL], fileUrl) && chunkCount === '' && !processingStatus) {
      // Found the new row - analyze it
      const fileId = getFileIdFromUrl(fileUrl);
      if (!fileId) continue;

      const readAloudEnabled = data[i][CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
      Logger.log(`Processing new assessment: ${fileUrl} (Read Aloud: ${readAloudEnabled})`);

      if (!readAloudEnabled) {
        markAssessmentAsNoAudioRequired(sheet, i + 1, fileUrl);
        break;
      }

      // Step 1: Extract text and count chunks
      const textChunks = extractTextFromFile(fileId);
      if (textChunks && textChunks.length > 0) {
        sheet.getRange(i + 1, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(textChunks.length);
        sheet.getRange(i + 1, CONSTANTS.COL.IS_COMPLETE + 1).setValue(false);
        SpreadsheetApp.flush();
        Logger.log(`Step 1 complete: ${textChunks.length} chunks`);

        // Step 2: Trigger audio generation in the background
        // We use a trigger instead of direct call to prevent UI timeouts
        Logger.log('→ Scheduling background audio generation (Step 2)...');
        triggerImmediateProcessing();
      }
      break;
    }
  }
}

/**
 * Creates a one-time trigger to run audio generation immediately (in the background).
 * This allows the main request to return to the user while processing continues.
 */
function triggerImmediateProcessing() {
  // Check if a trigger is already scheduled to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  const alreadyScheduled = triggers.some(t => t.getHandlerFunction() === 'step2_GenerateMissingAudioAndFinalize');
  
  if (!alreadyScheduled) {
    ScriptApp.newTrigger('step2_GenerateMissingAudioAndFinalize')
      .timeBased()
      .after(1000) // Run in 1 second
      .create();
    Logger.log('Created background trigger for Step 2');
  } else {
    Logger.log('Background trigger for Step 2 already exists, skipping');
  }
}
/**
 * Marks an assessment as complete when Read Aloud is disabled.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The assessment database sheet
 * @param {number} rowIndex The 1-based row index in the sheet
 */
function markAssessmentAsNoAudioRequired(sheet, rowIndex, pdfUrl) {
  let audioChunksJson = '[]';
  let chunkCount = 0;

  if (pdfUrl) {
    const fileId = getFileIdFromUrl(pdfUrl);
    if (fileId) {
      const structuredChunks = extractTextFromFile(fileId);
      if (structuredChunks && structuredChunks.length > 0) {
        chunkCount = structuredChunks.length;
        // Map to a format the student viewer expects, but without audio URLs
        const simplifiedChunks = structuredChunks.map(chunk => ({
          text: chunk.text,
          elementIds: chunk.ids,
          audioUrl: '' // No audio
        }));
        audioChunksJson = JSON.stringify(simplifiedChunks);
      }
    }
  }

  sheet.getRange(rowIndex, CONSTANTS.COL.CHUNK_COUNT + 1).setValue(chunkCount);
  setLargeDataInCell(sheet.getRange(rowIndex, CONSTANTS.COL.AUDIO_JSON + 1), audioChunksJson, pdfUrl);
  sheet.getRange(rowIndex, CONSTANTS.COL.IS_COMPLETE + 1).setValue(true);
  sheet.getRange(rowIndex, CONSTANTS.COL.PROCESSING_STATUS + 1).setValue("NO_AUDIO_REQUIRED");
  SpreadsheetApp.flush();
  Logger.log(`Row ${rowIndex}: Read Aloud disabled. Extracted ${chunkCount} chunks for highlighting. Marked as complete.`);
}


/**
 * Manually re-processes an assessment (runs steps 1 and 2).
 * Staff-only function. Teachers can only reprocess their own assessments.
 * @param {string} sessionToken Staff session token
 * @param {number} rowIndex Row index (1-based, excluding header)
 * @returns {Object} { success: true } or { error: "..." }
 */
function reprocessAssessment(sessionToken, rowIndex) {
  try {
    // Verify staff token
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) {
      return { error: 'Unauthorized. Staff access required.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) {
      return { error: 'Assessment Database sheet not found.' };
    }

    const actualRow = rowIndex + 1;
    if (actualRow < 2 || actualRow > sheet.getLastRow()) {
      return { error: 'Invalid row index.' };
    }

    // Check ownership for teachers
    const rowData = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!isAuthorizedForAssessment(tokenData, rowData)) {
      return { error: 'Unauthorized. You can only reprocess your own assessments.' };
    }

    const pdfUrl = rowData[CONSTANTS.COL.PDF_URL];
    if (!pdfUrl) {
      return { error: 'No file URL found in this row.' };
    }

    // Clear existing processing data
    sheet.getRange(actualRow, CONSTANTS.COL.CHUNK_COUNT + 1).setValue('');
    sheet.getRange(actualRow, CONSTANTS.COL.AUDIO_JSON + 1).setValue('');
    sheet.getRange(actualRow, CONSTANTS.COL.IS_COMPLETE + 1).setValue(false);
    SpreadsheetApp.flush();

    // Force regeneration: Delete existing audio folder
    try {
      const fileId = getFileIdFromUrl(pdfUrl);
      const file = DriveApp.getFileById(fileId);
      const fileName = file.getName();
      // Remove extension to get folder name
      // Matches logic in step2_GenerateMissingAudioAndFinalize
      const baseName = fileName.replace(/\.[^.]+$/i, '').trim(); 
      
      const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
      if (mainAudioFolder) {
        const subfolders = mainAudioFolder.getFoldersByName(baseName);
        while (subfolders.hasNext()) {
          const folder = subfolders.next();
          folder.setTrashed(true);
          Logger.log(`Trashed existing audio folder: ${baseName}`);
        }
      }
    } catch (folderError) {
      Logger.log(`Warning: Could not delete audio folder: ${folderError.toString()}`);
    }

    // Trigger processing
    processNewAssessment(pdfUrl);
    
    // Invalidate dashboard cache
    invalidateAssessmentsCache();

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

/**
 * Invalidates the dashboard assessment cache for ALL users.
 * Called when an assessment is added, updated, or deleted.
 */
function invalidateAssessmentsCache() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  
  try {
    // Wait for up to 5 seconds for other processes to finish
    if (lock.tryLock(5000)) {
      const currentVersion = parseInt(props.getProperty('CACHE_VERSION_ASSESSMENTS') || '1');
      props.setProperty('CACHE_VERSION_ASSESSMENTS', (currentVersion + 1).toString());
      Logger.log(`[CACHE] Global assessments cache version incremented to ${currentVersion + 1}`);
      lock.releaseLock();
    }
  } catch (e) {
    Logger.log(`[CACHE] Warning: Failed to increment cache version: ${e.toString()}`);
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

/**
 * Sets a value in a spreadsheet cell, automatically offloading to a Drive file
 * if the content exceeds the 50,000 character limit for Google Sheets.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The cell range to set
 * @param {string} content The content to set
 * @param {string} fileName Optional filename prefix for the Drive file
 */
function setLargeDataInCell(range, content, fileName) {
  // Check if current value is a Drive pointer and clean it up if so
  const currentValue = range.getValue();
  if (typeof currentValue === 'string' && currentValue.startsWith("DRIVE_FILE_ID:")) {
    const oldFileId = currentValue.replace("DRIVE_FILE_ID:", "");
    try {
      DriveApp.getFileById(oldFileId).setTrashed(true);
      Logger.log(`[DATA_OFFLOAD] Deleted old Drive file: ${oldFileId}`);
    } catch (e) {
      Logger.log(`[DATA_OFFLOAD] Warning: Could not delete old file ${oldFileId}: ${e.toString()}`);
    }
  }

  if (!content) {
    range.setValue('');
    return;
  }
  
  const limit = 45000; // Safe limit below 50,000
  if (content.length <= limit) {
    range.setValue(content);
    return;
  }
  
  Logger.log(`[DATA_OFFLOAD] Content length (${content.length}) exceeds limit. Offloading to Drive...`);
  
  try {
    const mainAudioFolder = getOrCreateFolder(CONSTANTS.AUDIO_DRIVE_FOLDER_NAME);
    const jsonFolder = getOrCreateSubfolder(mainAudioFolder, "JSON Metadata");
    
    if (!jsonFolder) throw new Error("Could not create JSON Metadata folder");
    
    const safeFileName = (fileName || "Data").substring(0, 50) + "_" + new Date().getTime() + ".json";
    const file = jsonFolder.createFile(safeFileName, content, MimeType.PLAIN_TEXT);
    
    // Store as a pointer
    const pointer = "DRIVE_FILE_ID:" + file.getId();
    range.setValue(pointer);
    Logger.log(`[DATA_OFFLOAD] Content offloaded to Drive file: ${file.getId()}`);
  } catch (e) {
    Logger.log(`[DATA_OFFLOAD] ERROR: Failed to offload to Drive: ${e.toString()}`);
    throw e;
  }
}

/**
 * Retrieves content from a cell, automatically fetching from a Drive file
 * if the content is a pointer (starts with "DRIVE_FILE_ID:").
 * @param {string} cellValue The raw value from the spreadsheet cell
 * @returns {string} The full content
 */
function getLargeDataFromCell(cellValue) {
  if (!cellValue || typeof cellValue !== 'string') return cellValue || '';
  
  if (cellValue.startsWith("DRIVE_FILE_ID:")) {
    const fileId = cellValue.replace("DRIVE_FILE_ID:", "");
    Logger.log(`[DATA_LOAD] Fetching content from Drive file: ${fileId}`);
    try {
      const file = DriveApp.getFileById(fileId);
      return file.getBlob().getDataAsString();
    } catch (e) {
      Logger.log(`[DATA_LOAD] ERROR: Failed to fetch from Drive (${fileId}): ${e.toString()}`);
      return ''; 
    }
  }
  
  return cellValue;
}

/**
 * Exposes specific public constants to the frontend.
 * @returns {Object} Public constants
 */
function getConstants() {
  return {
    SUBMISSION_FEATURE_ENABLED: CONSTANTS.SUBMISSION_FEATURE_ENABLED,
    SUBMISSION_ADMIN_ONLY: CONSTANTS.SUBMISSION_ADMIN_ONLY,
    SUBMISSION_ADMIN_ROLES: CONSTANTS.SUBMISSION_ADMIN_ROLES
  };
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

  // Check if staff user wants to preview student view
  if (user && e.parameter.preview === 'true' && 
      (user.userType === CONSTANTS.ROLE_TOKEN_TEACHER || 
       user.userType === CONSTANTS.ROLE_TOKEN_SPED || 
       user.userType === CONSTANTS.ROLE_TOKEN_ADMIN || 
       user.userType === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN)) {
    Logger.log(`doGet: Entering student preview mode. Preview: ${e.parameter.preview}, Assessment URL: ${e.parameter.assessmentUrl}`);
    const template = HtmlService.createTemplateFromFile('student');
    template.user = user;
    // Pass the specific assessment URL if provided (for direct preview from table)
    template.targetAssessmentUrl = e.parameter.assessmentUrl || null;
    return template.evaluate().setTitle('Student Preview').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }

  if (user && (user.userType === CONSTANTS.ROLE_TOKEN_TEACHER || 
               user.userType === CONSTANTS.ROLE_TOKEN_SPED || 
               user.userType === CONSTANTS.ROLE_TOKEN_ADMIN || 
               user.userType === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN)) {
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
  if (!email) return null;
  const cleanEmail = email.toLowerCase().trim();
  
  // Try CacheService first
  const cache = CacheService.getUserCache();
  const cachedUser = cache.get('user_info_' + cleanEmail);
  if (cachedUser) {
    try {
      return JSON.parse(cachedUser);
    } catch (e) {
      // If parsing fails, proceed to re-fetch
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Check Teachers sheet for staff members.
  const adminSheet = spreadsheet.getSheetByName(CONSTANTS.TEACHERS_SHEET_NAME);
  let userResult = null;

  if (adminSheet) {
    const adminData = adminSheet.getDataRange().getValues();
    for (let i = 1; i < adminData.length; i++) {
      const row = adminData[i];
      const adminEmail = row[2] ? row[2].toString().toLowerCase().trim() : '';
      if (adminEmail === cleanEmail) {
        const teacherRole = row[4] ? row[4].toString().trim() : CONSTANTS.ROLE_TEACHER;
        const betaFeaturesRaw = row[CONSTANTS.COL_TEACHERS_BETA_FEATURES] ? row[CONSTANTS.COL_TEACHERS_BETA_FEATURES].toString().trim() : '';
        const betaFeatures = betaFeaturesRaw.split(',').map(f => f.trim()).filter(f => f);

        let userType = CONSTANTS.ROLE_TOKEN_TEACHER;
        if (teacherRole === CONSTANTS.ROLE_SUPER_ADMIN) {
          userType = CONSTANTS.ROLE_TOKEN_SUPER_ADMIN;
        } else if (teacherRole === CONSTANTS.ROLE_ADMIN) {
          userType = CONSTANTS.ROLE_TOKEN_ADMIN;
        } else if (teacherRole === CONSTANTS.ROLE_SPED) {
          userType = CONSTANTS.ROLE_TOKEN_SPED;
        }
        
        userResult = {
          userType: userType,
          role: userType, // Use token role consistently for frontend logic
          displayName: teacherRole, // Keep display name separately if needed
          name: `${row[0]} ${row[1]}`.trim(),
          lastName: row[1] ? row[1].toString().trim() : '',
          email: cleanEmail,
          betaFeatures: betaFeatures
        };
        break;
      }
    }
  }

  if (!userResult) {
    // 2. Check Assessment Database for students.
    const studentSheet = spreadsheet.getSheetByName('Assessment Database');
    if (studentSheet) {
        const studentData = studentSheet.getDataRange().getValues();
        for (let i = 1; i < studentData.length; i++) {
            const studentEmailsRaw = studentData[i][CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();
            if (studentEmailsRaw.includes(cleanEmail)) {
                // Found the user in at least one assessment, classify as student.
                userResult = { userType: CONSTANTS.ROLE_TOKEN_STUDENT, email: cleanEmail };
                break;
            }
        }
    }
  }

  // Cache the result for 30 minutes (1800 seconds) if found
  if (userResult) {
    cache.put('user_info_' + cleanEmail, JSON.stringify(userResult), 1800);
  }

  return userResult;
}

/**
 * Retrieves a list of student emails assigned to a specific case manager.
 * Matches Case Manager Email in 'Student Directory' sheet.
 * @param {string} caseManagerEmail The email of the Sp.Ed. teacher
 * @returns {string[]} Array of student emails
 */
function getStudentsByCaseManager(caseManagerEmail) {
  if (!caseManagerEmail) return [];
  const cleanCMEmail = caseManagerEmail.toLowerCase().trim();
  
  // Try CacheService
  const cache = CacheService.getUserCache();
  const cachedStudents = cache.get('caseload_' + cleanCMEmail);
  if (cachedStudents) {
    try {
      return JSON.parse(cachedStudents);
    } catch (e) {
      // Proceed to fetch
    }
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const directorySheet = spreadsheet.getSheetByName('Student Directory');
    if (!directorySheet) {
      Logger.log('Student Directory sheet not found.');
      return [];
    }

    const data = directorySheet.getDataRange().getValues();
    const students = [];

    // Directory Structure: A) First, B) Last, C) Student Email, D) Case Manager Email
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentEmail = row[2] ? row[2].toString().toLowerCase().trim() : '';
      const cmEmail = row[3] ? row[3].toString().toLowerCase().trim() : '';

      if (cmEmail === cleanCMEmail && studentEmail) {
        students.push(studentEmail);
      }
    }

    // Cache for 1 hour (3600 seconds)
    cache.put('caseload_' + cleanCMEmail, JSON.stringify(students), 3600);
    return students;
  } catch (e) {
    Logger.log(`Error in getStudentsByCaseManager: ${e.toString()}`);
    return [];
  }
}


/**
 * Retrieves mappings of emails that have access to specific beta features.
 * Only accessible to Super Admins.
 * @param {string} sessionToken Super Admin session token
 * @returns {Object} { success: true, mappings: { featureId: [emails] } } or { error: "..." }
 */
function getBetaFeatureMappings(sessionToken) {
  try {
    const tokenData = validateSuperAdminToken(sessionToken);
    if (!tokenData) return { error: 'Unauthorized. Super Admin access required.' };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTANTS.TEACHERS_SHEET_NAME);
    if (!sheet) return { error: 'Teachers sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const mappings = {
      'Student Submissions': []
    };

    for (let i = 1; i < data.length; i++) {
      const email = data[i][2] ? data[i][2].toString().toLowerCase().trim() : '';
      const featuresRaw = data[i][CONSTANTS.COL_TEACHERS_BETA_FEATURES] ? data[i][CONSTANTS.COL_TEACHERS_BETA_FEATURES].toString().trim() : '';
      const features = featuresRaw.split(',').map(f => f.trim()).filter(f => f);

      if (email) {
        features.forEach(f => {
          if (!mappings[f]) mappings[f] = [];
          mappings[f].push(email);
        });
      }
    }

    return { success: true, mappings: mappings };
  } catch (e) {
    Logger.log(`Error in getBetaFeatureMappings: ${e.toString()}`);
    return { error: e.toString() };
  }
}

/**
 * Updates beta feature access for users in the Teachers sheet.
 * Only accessible to Super Admins.
 * @param {string} sessionToken Super Admin session token
 * @param {Object} mappings { featureId: [emails] }
 * @returns {Object} { success: true } or { error: "..." }
 */
function updateBetaFeatureMappings(sessionToken, mappings) {
  try {
    const tokenData = validateSuperAdminToken(sessionToken);
    if (!tokenData) return { error: 'Unauthorized. Super Admin access required.' };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTANTS.TEACHERS_SHEET_NAME);
    if (!sheet) return { error: 'Teachers sheet not found.' };

    const range = sheet.getDataRange();
    const data = range.getValues();
    
    // Create a per-user map for new features
    const userFeatureMap = {};
    for (const featureId in mappings) {
      const emails = mappings[featureId];
      emails.forEach(email => {
        const cleanEmail = email.toLowerCase().trim();
        if (!userFeatureMap[cleanEmail]) userFeatureMap[cleanEmail] = new Set();
        userFeatureMap[cleanEmail].add(featureId);
      });
    }

    // Update the data array
    for (let i = 1; i < data.length; i++) {
      const email = data[i][2] ? data[i][2].toString().toLowerCase().trim() : '';
      if (email) {
        const features = Array.from(userFeatureMap[email] || []);
        data[i][CONSTANTS.COL_TEACHERS_BETA_FEATURES] = features.join(', ');
      }
    }

    // Save back to sheet
    range.setValues(data);
    
    // Clear user cache to ensure roles are re-fetched
    const cache = CacheService.getUserCache();
    for (let i = 1; i < data.length; i++) {
       const email = data[i][2] ? data[i][2].toString().toLowerCase().trim() : '';
       if (email) cache.remove('user_info_' + email);
    }

    return { success: true };
  } catch (e) {
    Logger.log(`Error in updateBetaFeatureMappings: ${e.toString()}`);
    return { error: e.toString() };
  }
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
      if (areUrlsSameFile(data[i][CONSTANTS.COL.PDF_URL], assessmentUrl)) {
        const audioJson = getLargeDataFromCell(data[i][CONSTANTS.COL.AUDIO_JSON]);
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
        const assessmentUrl = tokenData.url;
        const cache = CacheService.getUserCache();
        // Use hash of URL for a safe and unique cache key
        const urlHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assessmentUrl));
        const cacheKey = 'audio_ids_' + urlHash;
        let audioFileIds = [];
        
        const cachedIds = cache.get(cacheKey);
        if (cachedIds) {
          audioFileIds = JSON.parse(cachedIds);
        } else {
          const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
          if (!sheet) {
            return {
              success: false,
              error: 'Assessment Database not found.'
            };
          }

          const data = sheet.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
              if (areUrlsSameFile(data[i][CONSTANTS.COL.PDF_URL], assessmentUrl)) {
                  const audioJson = getLargeDataFromCell(data[i][CONSTANTS.COL.AUDIO_JSON]);
                  if (audioJson) {
                    try {
                      const audioData = JSON.parse(audioJson);
                      audioFileIds = audioData.map(chunk => getFileIdFromUrl(chunk.audioUrl)).filter(id => id);
                      // Cache for 30 minutes
                      cache.put(cacheKey, JSON.stringify(audioFileIds), 1800);
                      break;
                    } catch (e) {
                      Logger.log(`Error parsing audio JSON: ${e.toString()}`);
                    }
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

  // Periodic cleanup of expired tokens to prevent UserProperties storage exhaustion (500KB limit)
  // Run approximately 10% of the time to minimize overhead
  if (Math.random() < 0.1) {
    cleanupExpiredTokens();
  }

  Logger.log(`Generated session token for ${email} (${tokenData.role}), expires: ${new Date(expiryTime)}`);
  return token;
}

/**
 * Purges expired session tokens from UserProperties to free up storage.
 */
function cleanupExpiredTokens() {
  const props = PropertiesService.getUserProperties();
  const allProps = props.getProperties();
  let deletedCount = 0;
  const now = Date.now();

  for (const key in allProps) {
    if (key.startsWith('session_')) {
      try {
        const tokenData = JSON.parse(allProps[key]);
        if (tokenData.exp && now > tokenData.exp) {
          props.deleteProperty(key);
          deletedCount++;
        }
      } catch (e) {
        // If it's not valid JSON, it's likely corrupt or old, so delete it
        props.deleteProperty(key);
        deletedCount++;
      }
    }
  }

  if (deletedCount > 0) {
    Logger.log(`[CLEANUP] Purged ${deletedCount} expired session tokens from UserProperties.`);
  }
}

/**
 * Validates a session token and returns the decoded data if valid.
 * Checks: token exists, not expired, email still has access to assessment
 * @param {string} token Session token to validate
 * @returns {Object|null} Token data if valid, null if invalid/expired
 */
function validateSessionToken(token) {
  if (!token) return null;

  // Try CacheService first
  const cache = CacheService.getUserCache();
  const tokenHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, token));
  const cachedData = cache.get('session_valid_' + tokenHash);
  if (cachedData) {
    try {
      return JSON.parse(cachedData);
    } catch (e) {
      // Proceed to full validation
    }
  }

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
      cache.put('session_valid_' + tokenHash, JSON.stringify(tokenData), 600); // Cache for 10 minutes
      return tokenData;
    }

    // Additional check for student tokens: verify email still has access to this assessment
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const email = tokenData.email;
    const assessmentUrl = tokenData.url;
    let hasAccess = false;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[CONSTANTS.COL.PDF_URL];
      const studentEmailsRaw = row[CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();

      if (areUrlsSameFile(pdfUrl, assessmentUrl)) {
        const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());
        if (studentEmails.includes(email)) {
          hasAccess = true;
          break;
        }
      }
    }

    if (hasAccess) {
      // Valid token and still has access
      cache.put('session_valid_' + tokenHash, JSON.stringify(tokenData), 600); // Cache for 10 minutes
      return tokenData;
    } else {
      Logger.log(`Email ${email} no longer has access to ${assessmentUrl}`);
      return null;
    }

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

    // Check if user is staff (for preview mode)
    const user = getUserByEmail(cleanEmail);
    const isStaff = user && (user.userType === CONSTANTS.ROLE_TOKEN_TEACHER || 
                             user.userType === CONSTANTS.ROLE_TOKEN_SPED || 
                             user.userType === CONSTANTS.ROLE_TOKEN_ADMIN || 
                             user.userType === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN);
    const isTeacher = user && user.userType === CONSTANTS.ROLE_TOKEN_TEACHER;
    const isSpEd = user && user.userType === CONSTANTS.ROLE_TOKEN_SPED;
    const staffName = user ? user.name : '';
    const caseloadStudents = isSpEd ? getStudentsByCaseManager(cleanEmail) : [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pdfUrl = row[CONSTANTS.COL.PDF_URL];
      const isComplete = row[CONSTANTS.COL.IS_COMPLETE];
      const studentEmailsRaw = row[CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();
      const className = row[CONSTANTS.COL.CLASS_NAME] ? row[CONSTANTS.COL.CLASS_NAME].toString().trim() : '';
      const instructor = row[CONSTANTS.COL.INSTRUCTOR] ? row[CONSTANTS.COL.INSTRUCTOR].toString().trim() : '';

      if (!pdfUrl || isComplete !== true) continue;

      // Determine if this assessment should be shown
      let shouldShow = false;

      if (isStaff) {
        // Staff see all, or teachers see only their own, or Sp.Ed. see caseload
        if (isTeacher) {
           // Split by comma or slash and check if any part is included in staff's name/email
           const instructors = instructor.toLowerCase().split(/[,\/]/).map(name => name.trim()).filter(name => name);
           shouldShow = instructors.some(name => staffName.toLowerCase().includes(name));
        } else if (isSpEd) {
           // Sp.Ed. see if any student in assessment is on their caseload
           shouldShow = caseloadStudents.some(studentEmail => studentEmailsRaw.includes(studentEmail));
        } else {
           shouldShow = true; // Admins see all
        }
      } else if (studentEmailsRaw) {
        // Students only see assigned
        const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());
        shouldShow = studentEmails.includes(cleanEmail);
      }

      if (shouldShow) {
        try {
          const fileId = getFileIdFromUrl(pdfUrl);
          if (fileId) {
            const file = DriveApp.getFileById(fileId);
            const fileName = file.getName();

            const isBetaUser = user && user.betaFeatures && user.betaFeatures.includes('Student Submissions');

            matchingAssessments.push({
              assessmentName: fileName,
              className: className,
              instructor: instructor,
              assessmentUrl: pdfUrl,
              readAloudEnabled: row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false,
              submissionEnabled: CONSTANTS.SUBMISSION_FEATURE_ENABLED && 
                                row[CONSTANTS.COL.SUBMISSION_ENABLED] === true && 
                                isBetaUser,
              requiresPassword: !!(row[CONSTANTS.COL.PASSWORD] && row[CONSTANTS.COL.PASSWORD].toString().trim()),
              rowIndex: i
            });
          }
        } catch (e) {
          Logger.log(`Warning: Could not fetch file info for row ${i}: ${e.toString()}`);
        }
      }
    }

    Logger.log(`Found ${matchingAssessments.length} assessment(s) for ${email} (Staff: ${isStaff})`);

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
      if (areUrlsSameFile(pdfUrl, assessmentUrl)) {
        const studentEmailsRaw = row[CONSTANTS.COL.STUDENT_EMAILS].toString().toLowerCase();
        const studentEmails = studentEmailsRaw.split(',').map(e => e.trim());
        const sheetPassword = row[CONSTANTS.COL.PASSWORD].toString().trim();

        // Check if user is staff
        const user = getUserByEmail(cleanEmail);
        const isStaff = user && (user.userType === CONSTANTS.ROLE_TOKEN_TEACHER || 
                                 user.userType === CONSTANTS.ROLE_TOKEN_SPED || 
                                 user.userType === CONSTANTS.ROLE_TOKEN_ADMIN || 
                                 user.userType === CONSTANTS.ROLE_TOKEN_SUPER_ADMIN);

        // Verify this authenticated user's email is in the list for this assessment (or is staff)
        if (isStaff || studentEmails.includes(cleanEmail)) {
          // NEW: Validate the provided password against the one in the sheet (skip for staff or if no password is set)
          if (!isStaff && sheetPassword && password !== sheetPassword) {
            return { error: 'Incorrect password for this assessment.' };
          }

          // Password is correct, proceed...
          const audioDataJson = getLargeDataFromCell(row[CONSTANTS.COL.AUDIO_JSON]);
          const readAloudEnabled = row[CONSTANTS.COL.READ_ALOUD_ENABLED] !== false;
          const accessExpires = row[CONSTANTS.COL.ACCESS_EXPIRES];

          // Check for expiration
          if (accessExpires && accessExpires instanceof Date && !isNaN(accessExpires)) {
            if (new Date() > accessExpires) {
              return { error: 'Access to this assessment has expired. Please contact your instructor.' };
            }
          }

          if (readAloudEnabled && !audioDataJson) {
            return { error: "Audio for this assessment has not been generated yet. Please try again later." };
          }

          const fileId = getFileIdFromUrl(pdfUrl);
          if (!fileId) return { error: "Invalid Google Drive URL in sheet." };

          const file = DriveApp.getFileById(fileId);
          const mimeType = file.getMimeType();
          const fileName = file.getName();
          const audioChunks = (audioDataJson && audioDataJson.trim()) ? JSON.parse(audioDataJson) : [];

          Logger.log(`Serving assessment: ${fileName} (${mimeType}) to ${email}`);

          // Generate session token for secure audio access
          const sessionToken = isStaff ? 
            generateSessionToken(cleanEmail, user.userType) : 
            generateSessionToken(cleanEmail, pdfUrl);

          // Get cached HTML if available, otherwise fall back to real-time conversion
          let assessmentHtml = getLargeDataFromCell(row[CONSTANTS.COL.ASSESSMENT_HTML]);
          const lastProcessed = row[CONSTANTS.COL.LAST_PROCESSED_TIME];
          const fileLastUpdated = file.getLastUpdated();
          
          // Smart Refresh: If file was updated in Drive AFTER we last processed it, force re-conversion
          let forceReconversion = false;
          if (assessmentHtml && lastProcessed instanceof Date && fileLastUpdated > lastProcessed) {
            Logger.log(`→ Smart Refresh: Drive file is newer than cache (${fileLastUpdated} > ${lastProcessed}). Forcing re-conversion.`);
            forceReconversion = true;
          }

          if (!assessmentHtml || forceReconversion) {
            Logger.log(forceReconversion ? '→ Re-converting HTML...' : '→ No cached HTML found. Converting in real-time...');
            const conversionResult = convertFileToHtml(fileId);
            if (conversionResult.error) {
              Logger.log(`✗ Conversion error: ${conversionResult.error}`);
              return { error: `Could not load assessment: ${conversionResult.error}` };
            }
            assessmentHtml = sanitizeHtml(conversionResult.html);
            
            // Update the cache and the last processed time
            setLargeDataInCell(sheet.getRange(i + 1, CONSTANTS.COL.ASSESSMENT_HTML + 1), assessmentHtml, fileName + "_html");
            sheet.getRange(i + 1, CONSTANTS.COL.LAST_PROCESSED_TIME + 1).setValue(new Date());
          } else {
            Logger.log('→ Using cached HTML from spreadsheet');
          }

          const submissionEnabled = CONSTANTS.SUBMISSION_FEATURE_ENABLED && row[CONSTANTS.COL.SUBMISSION_ENABLED] === true;

          return {
            fileType: "html",
            assessmentHtml: assessmentHtml,
            fileName: fileName,
            audioChunks: audioChunks,
            sessionToken: sessionToken,
            readAloudEnabled: readAloudEnabled,
            submissionEnabled: submissionEnabled
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

// --- SUBMISSION FUNCTIONS ---

/**
 * Looks up an instructor's email from the Teachers sheet by name.
 * Matches against "First Last" name format.
 * @param {string} instructorName The instructor's display name (e.g., "John Smith")
 * @returns {string|null} The instructor's email or null if not found
 */
/**
 * Looks up instructor emails from the Teachers sheet by name or returns them directly if they are emails.
 * Supports multiple comma-separated inputs.
 * @param {string} instructorName The instructor's display name or email (e.g., "John Smith" or "jsmith@example.com")
 * @returns {string|null} Comma-separated instructor emails or null if none found
 */
function getInstructorEmail(instructorName) {
  if (!instructorName) return null;

  // Support comma-separated instructor names/emails or slash-separated
  const instructorNames = instructorName.toLowerCase().trim().split(/[,\/]/).map(function(s) { return s.trim(); });
  const results = [];

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teacherSheet = spreadsheet.getSheetByName(CONSTANTS.TEACHERS_SHEET_NAME);
  
  // Cache teacher data if available
  let teacherData = null;
  if (teacherSheet) {
    teacherData = teacherSheet.getDataRange().getValues();
  }

  instructorNames.forEach(name => {
    if (!name) return;

    // 0. If the input is already an email, add it directly
    if (name.includes('@')) {
      results.push(name);
      return;
    }

    if (!teacherData) return;

    // 1. Exact full name match
    for (let i = 1; i < teacherData.length; i++) {
      const row = teacherData[i];
      const fullName = (row[0] + ' ' + row[1]).toLowerCase().trim();
      if (fullName === name) {
        if (row[2]) results.push(row[2].toString().trim());
        return;
      }
    }

    // 2. Full name contains the search name (e.g., search "Smith" matches "John Smith")
    for (let i = 1; i < teacherData.length; i++) {
      const row = teacherData[i];
      const fullName = (row[0] + ' ' + row[1]).toLowerCase().trim();
      if (fullName.includes(name)) {
        if (row[2]) results.push(row[2].toString().trim());
        return;
      }
    }

    // 3. Last name fallback
    const fallbackMatches = [];
    for (let i = 1; i < teacherData.length; i++) {
      const row = teacherData[i];
      const lastName = row[1] ? row[1].toString().toLowerCase().trim() : '';
      if (lastName && name.includes(lastName)) {
        fallbackMatches.push(row);
      }
    }

    if (fallbackMatches.length === 1) {
      const match = fallbackMatches[0];
      if (match[2]) results.push(match[2].toString().trim());
    }
  });

  // Return unique emails as a comma-separated string
  const uniqueEmails = [...new Set(results)].filter(e => e);
  return uniqueEmails.length > 0 ? uniqueEmails.join(', ') : null;
}

/**
 * Submits student assessment responses. Generates a PDF from the responses
 * and emails it to the instructor.
 * @param {string} sessionToken Session token for authentication
 * @param {string} assessmentUrl The assessment URL identifier
 * @param {Array<Object>} responses Array of { chunkIndex, questionLabel, questionType, answer }
 * @returns {Object} { success: true } or { error: "..." }
 */
function submitAssessmentResponses(sessionToken, assessmentUrl, responses) {
  const lock = LockService.getScriptLock();
  try {
    // Wait for up to 120 seconds for the lock to handle high concurrency (e.g. 90 students)
    lock.waitLock(120000);

    // Validate session
    const tokenData = validateSessionToken(sessionToken);
    if (!tokenData) {
      return { error: 'Session expired. Please log in again.' };
    }

    // Ensure student tokens are only used for their authorized assessment
    if (tokenData.role === CONSTANTS.ROLE_TOKEN_STUDENT && !areUrlsSameFile(tokenData.url, assessmentUrl)) {
      return { error: 'Unauthorized for this assessment.' };
    }

    const studentEmail = tokenData.email;
    // Attempt to pull student name from Student Directory (Column A + B)
    // Fall back to token name (if provided) or student email
    const directoryName = getStudentNameFromDirectory(studentEmail);
    const studentName = directoryName || tokenData.name || studentEmail;

    // Look up assessment metadata
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
    if (!sheet) return { error: 'Assessment Database not found.' };

    const data = sheet.getDataRange().getValues();
    let assessmentName = 'Assessment';
    let className = '';
    let instructorName = '';
    let assessmentFound = false;
    let assessmentRowIndex = -1;
    let storedChunks = [];
    let submissionDeliveryMode = 'email';

    for (let i = 1; i < data.length; i++) {
      if (areUrlsSameFile(data[i][CONSTANTS.COL.PDF_URL], assessmentUrl)) {
        assessmentFound = true;
        assessmentRowIndex = i;
        // Verify submissions are enabled for this assessment
        if (data[i][CONSTANTS.COL.SUBMISSION_ENABLED] !== true) {
          return { error: 'Submissions are not enabled for this assessment.' };
        }
        
        submissionDeliveryMode = data[i][CONSTANTS.COL.SUBMISSION_DELIVERY_MODE] || 'email';

        // Check for duplicate submission using getLargeDataFromCell to handle Drive-offloaded pointers
        const rawTimestamps = data[i][CONSTANTS.COL.SUBMISSION_TIMESTAMPS];
        const submissionTimestampsJson = getLargeDataFromCell(rawTimestamps);
        
        if (submissionTimestampsJson && submissionTimestampsJson.trim()) {
          try {
            const submissionTimestamps = JSON.parse(submissionTimestampsJson);
            if (submissionTimestamps[studentEmail]) {
              const previousSubmission = new Date(submissionTimestamps[studentEmail]);
              return { 
                error: 'You have already submitted this assessment on ' + 
                  previousSubmission.toLocaleString('en-US', { timeZone: 'America/Chicago' }) + 
                  '. Multiple submissions are not allowed.' 
              };
            }
          } catch (e) {
            Logger.log('Error parsing submission timestamps: ' + e.toString());
          }
        }

        // Load stored chunks for validation
        try {
          const audioDataJson = getLargeDataFromCell(data[i][CONSTANTS.COL.AUDIO_JSON]);
          if (audioDataJson) {
            storedChunks = JSON.parse(audioDataJson);
          }
        } catch (e) {
          Logger.log('Error parsing audio chunks for validation: ' + e.toString());
        }

        className = data[i][CONSTANTS.COL.CLASS_NAME] ? data[i][CONSTANTS.COL.CLASS_NAME].toString().trim() : '';
        instructorName = data[i][CONSTANTS.COL.INSTRUCTOR] ? data[i][CONSTANTS.COL.INSTRUCTOR].toString().trim() : '';
        try {
          const fileId = getFileIdFromUrl(assessmentUrl);
          if (fileId) assessmentName = DriveApp.getFileById(fileId).getName();
        } catch (e) {
          Logger.log('Could not get assessment filename: ' + e.toString());
        }
        break;
      }
    }

    if (!assessmentFound) {
      return { error: 'Assessment not found.' };
    }

    // Validate that submission contains at least some answers
    const answeredCount = responses.filter(r => r.answer && r.answer.toString().trim()).length;
    if (answeredCount === 0) {
      return { error: 'Cannot submit an empty assessment. Please answer at least one question.' };
    }

    // Server-side validation: overwrite client-provided labels and types with stored metadata
    // This prevents students from spoofing question labels or types
    for (let i = 0; i < responses.length; i++) {
      const response = responses[i];
      if (response.chunkIndex < 0 || response.chunkIndex >= storedChunks.length) {
        Logger.log('Warning: Invalid chunk index ' + response.chunkIndex + ' in submission from ' + studentEmail);
        return { error: 'Invalid response data. Please refresh and try again.' };
      }

      const chunkMetadata = storedChunks[response.chunkIndex];
      // Use the server-side searchWords as the label if not provided, or regenerate it
      // Actually, buildSidebarFields uses getQuestionLabel(chunk.text, index)
      // We should replicate that logic here for consistency
      const numMatch = chunkMetadata.text.match(/^(\d+)[.)\]]/);
      const serverLabel = numMatch ? 'Question ' + numMatch[1] : 'Part ' + (response.chunkIndex + 1);
      
      // Determine type based on the text (same as detectQuestionType in student.html)
      // This is a duplicate of the logic, maybe should be a shared utility
      const mcPattern = /(?:^|[\n\r])\s*\(?([a-eA-E])\s*[.)](?:\s+|$)/g;
      const matches = [...chunkMetadata.text.matchAll(mcPattern)];
      let serverType = (matches.length >= 2) ? 'mc' : 'text';

      // Check for horizontal layout as well
      if (serverType === 'text') {
        const horizontalPattern = /\s{2,}\(?([b-eB-E])\s*[.)](?:\s+|$)/g;
        const hMatches = [...chunkMetadata.text.matchAll(horizontalPattern)];
        if (hMatches.length > 0 && /^\s*\(?[aA]\s*[.)]/.test(chunkMetadata.text)) {
          serverType = 'mc';
        }
      }

      response.questionLabel = serverLabel;
      response.questionType = serverType;
    }

    // --- STEP 1: Save to Submissions Sheet (ALWAYS) ---
    try {
      Logger.log('[SUBMIT] Saving to Submissions sheet...');
      const submissionsSheet = getOrCreateSubmissionsSheet();
      const responsesJson = JSON.stringify(responses);
      
      // Build row data using constants for column mapping
      const rowData = [];
      rowData[CONSTANTS.COL_SUBMISSIONS.TIMESTAMP] = new Date();
      rowData[CONSTANTS.COL_SUBMISSIONS.ASSESSMENT_URL] = assessmentUrl;
      rowData[CONSTANTS.COL_SUBMISSIONS.ASSESSMENT_NAME] = assessmentName;
      rowData[CONSTANTS.COL_SUBMISSIONS.STUDENT_EMAIL] = studentEmail;
      rowData[CONSTANTS.COL_SUBMISSIONS.STUDENT_NAME] = studentName;
      rowData[CONSTANTS.COL_SUBMISSIONS.RESPONSES_JSON] = ''; // Placeholder for offloaded data
      
      submissionsSheet.appendRow(rowData);
      
      // Use setLargeDataInCell for the responses JSON to handle potential 50k char limit
      const lastRow = submissionsSheet.getLastRow();
      const responsesRange = submissionsSheet.getRange(lastRow, CONSTANTS.COL_SUBMISSIONS.RESPONSES_JSON + 1);
      setLargeDataInCell(responsesRange, responsesJson, `Submission_${assessmentName}_${studentName}`);
      
    } catch (e) {
      Logger.log('[SUBMIT] Error saving to Submissions sheet: ' + e.toString());
      // For bulk mode, this is critical
      if (submissionDeliveryMode === 'bulk') {
        throw e; // Rethrow to be caught by main catch block
      }
    }

    // RELEASE LOCK during slow PDF generation and emailing phase to allow other students to submit concurrently
    try {
      lock.releaseLock();
      Logger.log('[SUBMIT] Released lock for slow operations.');
    } catch (lockErr) {
      Logger.log('[SUBMIT] Warning: Error releasing lock: ' + lockErr.toString());
    }

    const timestamp = new Date().toLocaleString('en-US', { timeZone: 'America/Chicago' });

    // --- STEP 2: Emailing (CONDITIONAL) ---
    if (submissionDeliveryMode === 'email') {
        Logger.log('[SUBMIT] Generating HTML content for PDF...');
        
        let htmlBody = '<html><head><style>';
        htmlBody += 'body { font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }';
        htmlBody += '.header { text-align: center; border-bottom: 2px solid #2d3f89; padding-bottom: 20px; margin-bottom: 30px; }';
        htmlBody += '.header h1 { color: #2d3f89; margin: 0; font-size: 24px; }';
        htmlBody += '.header p { color: #666; font-size: 14px; margin: 5px 0 0; }';
        htmlBody += '.metadata { background: #f8f9fc; padding: 20px; border-radius: 8px; margin-bottom: 30px; border: 1px solid #e1e4e8; }';
        htmlBody += '.metadata-table { width: 100%; border-collapse: collapse; }';
        htmlBody += '.metadata-table td { padding: 8px 0; vertical-align: top; border-bottom: 1px solid #edf0f2; }';
        htmlBody += '.metadata-table td:first-child { width: 140px; font-weight: bold; color: #555; }';
        htmlBody += '.responses-title { color: #2d3f89; font-size: 20px; border-bottom: 1px solid #e1e4e8; padding-bottom: 10px; margin: 40px 0 20px; }';
        htmlBody += '.question-block { margin-bottom: 30px; page-break-inside: avoid; border-left: 3px solid #eaecf5; padding-left: 20px; }';
        htmlBody += '.question-label { font-weight: bold; color: #2d3f89; margin-bottom: 10px; font-size: 16px; }';
        htmlBody += '.answer-container { padding: 15px; background: #fff; border: 1px solid #e1e4e8; border-radius: 6px; min-height: 40px; }';
        htmlBody += '.answer-mc { background: #f5f7fa; border-left: 4px solid #2d3f89; font-weight: bold; font-size: 17px; }';
        htmlBody += '.answer-text { line-height: 1.6; font-size: 15px; }';
        htmlBody += '.empty-answer { color: #999; font-style: italic; }';
        htmlBody += '.footer { margin-top: 50px; padding-top: 20px; border-top: 1px solid #eee; text-align: center; color: #999; font-size: 11px; }';
        htmlBody += '</style></head><body>';

        htmlBody += '<div class="header"><h1>Assessment Submission</h1><p>Spartan Assessment Portal</p></div>';

        htmlBody += '<div class="metadata"><table class="metadata-table">';
        htmlBody += '<tr><td>Student</td><td>' + escapeHtmlBackend(studentName) + ' (' + escapeHtmlBackend(studentEmail) + ')</td></tr>';
        htmlBody += '<tr><td>Assessment</td><td>' + escapeHtmlBackend(assessmentName) + '</td></tr>';
        if (className) htmlBody += '<tr><td>Class</td><td>' + escapeHtmlBackend(className) + '</td></tr>';
        if (instructorName) htmlBody += '<tr><td>Instructor</td><td>' + escapeHtmlBackend(instructorName) + '</td></tr>';
        htmlBody += '<tr><td>Submitted At</td><td>' + timestamp + '</td></tr>';
        htmlBody += '</table></div>';

        htmlBody += '<h2 class="responses-title">Student Responses</h2>';

        // Add each response
        for (let i = 0; i < responses.length; i++) {
          const r = responses[i];
          const answerText = r.answer ? r.answer.toString().trim() : '';
          const hasAnswer = answerText.length > 0;
          
          htmlBody += '<div class="question-block">';
          htmlBody += '<div class="question-label">' + escapeHtmlBackend(r.questionLabel) + '</div>';
          
          if (r.questionType === 'mc') {
            htmlBody += '<div class="answer-container answer-mc">';
            htmlBody += hasAnswer ? escapeHtmlBackend(answerText) : '<span class="empty-answer">No answer selected</span>';
            htmlBody += '</div>';
          } else {
            htmlBody += '<div class="answer-container answer-text">';
            htmlBody += hasAnswer ? renderSubmissionAnswer(answerText) : '<span class="empty-answer">No response provided</span>';
            htmlBody += '</div>';
          }
          htmlBody += '</div>';
        }

        htmlBody += '<div class="footer">Generated by Spartan Assessment Portal on ' + timestamp + '</div>';
        htmlBody += '</body></html>';

        Logger.log('[SUBMIT] Creating temporary Google Doc via Drive API...');
        // Create a Google Doc from the HTML content
        const htmlBlob = Utilities.newBlob(htmlBody, 'text/html', 'submission.html');
        
        let tempDocMeta;
        try {
          tempDocMeta = Drive.Files.create(
            { name: 'Submission - ' + assessmentName + ' - ' + studentName, mimeType: 'application/vnd.google-apps.document' },
            htmlBlob
          );
          Logger.log('[SUBMIT] Temp Doc created: ' + tempDocMeta.id);
        } catch (driveErr) {
          Logger.log('[SUBMIT] Drive API error: ' + driveErr.toString());
          throw driveErr;
        }

        Logger.log('[SUBMIT] Converting Doc to PDF...');
        // Export as PDF
        const pdfBlob = DriveApp.getFileById(tempDocMeta.id).getAs('application/pdf');
        pdfBlob.setName(assessmentName + ' - ' + studentName + ' - Submission.pdf');
        Logger.log('[SUBMIT] PDF generated successfully.');

        // Clean up temp doc
        Logger.log('[SUBMIT] Trashing temporary Doc...');
        DriveApp.getFileById(tempDocMeta.id).setTrashed(true);

        // Find instructor email
        Logger.log('[SUBMIT] Finding instructor email for: ' + instructorName);
        const instructorEmail = getInstructorEmail(instructorName);
        Logger.log('[SUBMIT] Instructor email: ' + (instructorEmail || 'NOT FOUND'));

        // Build email subject from template
        const subject = CONSTANTS.SUBMISSION_EMAIL_SUBJECT
          .replace('{assessmentName}', assessmentName)
          .replace('{studentName}', studentName);

        // Send email
        if (instructorEmail) {
          Logger.log('[SUBMIT] Sending email to instructor: ' + instructorEmail);
          MailApp.sendEmail(instructorEmail, subject,
            'Please see the attached assessment submission from ' + studentName + '.', {
            attachments: [pdfBlob],
            name: 'Spartan Assessment Portal',
            replyTo: studentEmail
          });
          Logger.log('[SUBMIT] Email sent to instructor.');
        } else {
          // Fallback: send to the script owner / deployer
          const fallbackEmail = Session.getEffectiveUser().getEmail();
          Logger.log('[SUBMIT] Falling back to script owner email: ' + fallbackEmail);
          MailApp.sendEmail(fallbackEmail, subject + ' [Instructor Not Found]',
            'Could not find email for instructor "' + instructorName + '". Submission from ' + studentName + ' attached.', {
            attachments: [pdfBlob],
            name: 'Spartan Assessment Portal',
            replyTo: studentEmail
          });
          Logger.log('[SUBMIT] Email sent to fallback.');
        }
    } else {
        Logger.log('[SUBMIT] Delivery mode is BULK. Skipping individual email.');
    }

    // Record submission timestamp
    Logger.log('[SUBMIT] Recording timestamp in spreadsheet...');
    try {
      // Re-acquire lock to ensure safe update of the shared timestamp object in the database
      try {
        lock.waitLock(120000);
        Logger.log('[SUBMIT] Re-acquired lock for final timestamp update.');
      } catch (lockErr) {
        Logger.log('[SUBMIT] Warning: Could not re-acquire lock for timestamp: ' + lockErr.toString());
        // Continue anyway as the primary data is already saved in the Submissions sheet
      }

      // Re-read current timestamps from sheet as they may have changed during the slow PDF generation phase
      const timestampsRange = sheet.getRange(assessmentRowIndex + 1, CONSTANTS.COL.SUBMISSION_TIMESTAMPS + 1);
      const currentTimestampsJson = getLargeDataFromCell(timestampsRange.getValue());
      
      let submissionTimestamps = {};
      if (currentTimestampsJson && currentTimestampsJson.trim()) {
        try {
          submissionTimestamps = JSON.parse(currentTimestampsJson);
        } catch (e) {
          Logger.log('[SUBMIT] Error parsing existing timestamps: ' + e.toString());
        }
      }
      submissionTimestamps[studentEmail] = new Date().toISOString();
      
      setLargeDataInCell(timestampsRange, JSON.stringify(submissionTimestamps), `Timestamps_${assessmentName}`);
      
      Logger.log('[SUBMIT] Timestamp recorded.');

      // Invalidate the teacher dashboard cache so they see the "Download PDF" button immediately
      invalidateAssessmentsCache();
      
    } catch (e) {
      Logger.log('[SUBMIT] Error recording submission timestamp: ' + e.toString());
    }

    Logger.log('[SUBMIT] Submission process completed successfully.');
    return { success: true, sentTo: submissionDeliveryMode === 'email' ? 'instructor' : 'stored' };

  } catch (e) {
    Logger.log('[SUBMIT] CRITICAL ERROR: ' + e.toString());
    return { error: 'Submission failed: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Generates a consolidated PDF of all submissions for an assessment.
 * @param {string} sessionToken - Admin session token
 * @param {string} assessmentUrl - Assessment URL identifier
 * @returns {Object} { url: "..." } or { error: "..." }
 */
function generateConsolidatedSubmissionsPdf(sessionToken, assessmentUrl) {
  try {
    const tokenData = validateAdminToken(sessionToken);
    if (!tokenData) return { error: 'Unauthorized.' };

    const sheet = getOrCreateSubmissionsSheet();
    const data = sheet.getDataRange().getValues();
    
    // Filter rows for this assessment
    // CONSTANTS.COL_SUBMISSIONS: { TIMESTAMP: 0, ASSESSMENT_URL: 1, ASSESSMENT_NAME: 2, STUDENT_EMAIL: 3, STUDENT_NAME: 4, RESPONSES_JSON: 5 }
    const submissions = [];
    let assessmentName = 'Assessment';
    
    // Start from row 1 (skip header)
    for (let i = 1; i < data.length; i++) {
        if (areUrlsSameFile(data[i][CONSTANTS.COL_SUBMISSIONS.ASSESSMENT_URL], assessmentUrl)) {
            const rawResponses = data[i][CONSTANTS.COL_SUBMISSIONS.RESPONSES_JSON];
            const responsesJson = getLargeDataFromCell(rawResponses);
            
            try {
              submissions.push({
                  timestamp: data[i][CONSTANTS.COL_SUBMISSIONS.TIMESTAMP],
                  studentName: data[i][CONSTANTS.COL_SUBMISSIONS.STUDENT_NAME],
                  studentEmail: data[i][CONSTANTS.COL_SUBMISSIONS.STUDENT_EMAIL],
                  responses: JSON.parse(responsesJson || '[]')
              });
            } catch (parseErr) {
              Logger.log(`Error parsing responses for ${data[i][CONSTANTS.COL_SUBMISSIONS.STUDENT_EMAIL]}: ${parseErr.toString()}`);
              // Push with empty responses so the student is still listed in the report
              submissions.push({
                  timestamp: data[i][CONSTANTS.COL_SUBMISSIONS.TIMESTAMP],
                  studentName: data[i][CONSTANTS.COL_SUBMISSIONS.STUDENT_NAME],
                  studentEmail: data[i][CONSTANTS.COL_SUBMISSIONS.STUDENT_EMAIL],
                  responses: [],
                  error: 'Data corrupted'
              });
            }
            if (data[i][CONSTANTS.COL_SUBMISSIONS.ASSESSMENT_NAME]) {
              assessmentName = data[i][CONSTANTS.COL_SUBMISSIONS.ASSESSMENT_NAME];
            }
        }
    }

    if (submissions.length === 0) {
        return { error: 'No submissions found for this assessment.' };
    }

    // Sort by student name
    submissions.sort((a, b) => (a.studentName || '').localeCompare(b.studentName || ''));

    // Generate HTML
    let htmlBody = '<html><head><style>';
    htmlBody += 'body { font-family: "Helvetica Neue", Helvetica, Arial, sans-serif; color: #333; margin: 0; padding: 0; line-height: 1.0; }';
    htmlBody += '.title-page { text-align: center; margin-top: 0; margin-bottom: 20px; border-bottom: 2px solid #2d3f89; padding-bottom: 10px; }';
    htmlBody += '.title-page h1 { color: #2d3f89; font-size: 24px; margin: 0 0 5px 0; }';
    htmlBody += '.title-page h2 { color: #666; font-size: 16px; font-weight: normal; margin: 0 0 10px 0; }';
    htmlBody += '.stats { font-size: 12px; color: #444; }';
    htmlBody += '.page-break { page-break-after: always; }';
    
    htmlBody += '.student-section { border-top: none; padding-top: 0; margin-bottom: 15px; }';
    htmlBody += '.student-header { border-bottom: 1px solid #2d3f89; padding-bottom: 3px; margin-bottom: 5px; }';
    htmlBody += '.student-header h2 { color: #2d3f89; margin: 0; font-size: 16px; }';
    htmlBody += '.student-meta { display: flex; justify-content: space-between; margin-top: 2px; color: #666; font-size: 10px; }';
    
    htmlBody += '.response-table { width: 100%; table-layout: fixed; border-collapse: collapse; border: none; }';
    htmlBody += '.response-cell { width: 33.3%; vertical-align: top; padding: 1px 10px 4px 10px; box-sizing: border-box; border: none; }';
    htmlBody += '.question-block { margin: 0; padding: 0; border: none; }';
    htmlBody += '.question-label { font-weight: bold; margin-bottom: 2px; color: #555; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; }';
    htmlBody += '.answer-container { padding: 0 0 2px 0; border: none; background: transparent; font-size: 11px; }';
    htmlBody += '.answer-mc { font-weight: bold; color: #2d3f89; }';
    htmlBody += '.answer-text { line-height: 1.1; word-wrap: break-word; color: #333; }';
    htmlBody += '.empty-answer { color: #999; font-style: italic; font-size: 10px; }';
    htmlBody += '</style></head><body>';
    
    // Title Page
    htmlBody += `<div class="title-page">
        <h1>Submission Report</h1>
        <h2>${escapeHtmlBackend(assessmentName)}</h2>
        <div class="stats">
            <p><strong>Total Submissions:</strong> ${submissions.length} | <strong>Generated:</strong> ${new Date().toLocaleString()}</p>
        </div>
    </div>`;
    
    submissions.forEach((sub, index) => {
        // Add a marker for a native page break before every student except the very first one
        if (index > 0) {
            htmlBody += '<p>[[PAGE_BREAK]]</p>';
        }

        htmlBody += '<div class="student-section">';
        
        // Student Header
        htmlBody += `<div class="student-header">
            <h2>${escapeHtmlBackend(sub.studentName)}</h2>
            <div class="student-meta">
                <span>${escapeHtmlBackend(sub.studentEmail)}</span>
                <span style="float: right;">Submitted: ${new Date(sub.timestamp).toLocaleString()}</span>
            </div>
            <div style="clear: both;"></div>
        </div>`;
        
        if (sub.error) {
            htmlBody += `<div style="color: #d32f2f; padding: 10px; border: 1px solid #d32f2f; border-radius: 4px; margin: 10px 0; font-size: 11px; background-color: #fdf5f5;">
                <strong>⚠️ DATA ERROR:</strong> ${escapeHtmlBackend(sub.error)}. The detailed responses for this student could not be loaded from the database.
            </div>`;
        }
        
        // --- HYBRID LAYOUT: Group sequential MC questions into columns, keep Text questions full-width ---
        const blocks = [];
        let currentBlock = null;

        sub.responses.forEach(r => {
            const isMc = r.questionType === 'mc';
            
            if (isMc) {
                if (!currentBlock || currentBlock.type !== 'mc') {
                    currentBlock = { type: 'mc', items: [] };
                    blocks.push(currentBlock);
                }
                currentBlock.items.push(r);
            } else {
                blocks.push({ type: 'text', item: r });
                currentBlock = null;
            }
        });

        // Render blocks
        blocks.forEach((block, bIdx) => {
            if (block.type === 'mc') {
                // Render MC group in 3-column table
                const items = block.items;
                const itemsPerCol = Math.ceil(items.length / 3);
                
                htmlBody += '<table class="response-table">';
                for (let rowIdx = 0; rowIdx < itemsPerCol; rowIdx++) {
                    htmlBody += '<tr>';
                    for (let colIdx = 0; colIdx < 3; colIdx++) {
                        htmlBody += '<td class="response-cell" style="padding-bottom: 8px;">';
                        const itemIdx = rowIdx + (colIdx * itemsPerCol);
                        if (itemIdx < items.length) {
                            const r = items[itemIdx];
                            const answerText = r.answer ? r.answer.toString().trim() : '';
                            const hasAnswer = answerText.length > 0;
                            
                            htmlBody += '<div class="question-block">';
                            htmlBody += `<div class="question-label">${escapeHtmlBackend(r.questionLabel)}</div>`;
                            htmlBody += '<div class="answer-container answer-mc">';
                            htmlBody += hasAnswer ? escapeHtmlBackend(answerText) : '<span class="empty-answer">No answer</span>';
                            htmlBody += '</div></div>';
                        }
                        htmlBody += '</td>';
                    }
                    htmlBody += '</tr>';
                }
                htmlBody += '</table>';
            } else {
                // Render text response full-width
                const r = block.item;
                const answerText = r.answer ? r.answer.toString().trim() : '';
                const hasAnswer = answerText.length > 0;
                
                htmlBody += '<div style="width: 100%; padding: 0 10px; box-sizing: border-box;">';
                htmlBody += `<div class="question-label" style="border-bottom: 1px solid #f0f0f0; padding-bottom: 2px; margin-bottom: 4px;">${escapeHtmlBackend(r.questionLabel)}</div>`;
                htmlBody += '<div class="answer-container answer-text" style="min-height: 20px; font-size: 11px;">';
                htmlBody += hasAnswer ? renderSubmissionAnswer(answerText) : '<span class="empty-answer">No response provided</span>';
                htmlBody += '</div></div>';
            }
            
            // Add a "blank line" spacer between ALL blocks (MC tables or text responses)
            htmlBody += '<div style="height: 25px; font-size: 1px;">&nbsp;</div>';
        });
        
        htmlBody += '</div>';
    });
    
    htmlBody += '</body></html>';

    // Generate PDF
    Logger.log('[REPORT] Creating temporary Doc...');
    const htmlBlob = Utilities.newBlob(htmlBody, 'text/html', 'report.html');
    let tempDocMeta;
    try {
        tempDocMeta = Drive.Files.create(
            { name: `Submissions - ${assessmentName}`, mimeType: 'application/vnd.google-apps.document' },
            htmlBlob
        );
    } catch (e) {
        Logger.log('[REPORT] Drive API error: ' + e.toString());
        throw e;
    }
    
    Logger.log('[REPORT] Converting to PDF...');
    
    // Open the document to set margins and clean up top-of-page whitespace
    const doc = DocumentApp.openById(tempDocMeta.id);
    const body = doc.getBody();
    
    // Set 1-inch (72pt) margins
    body.setMarginTop(72);
    body.setMarginBottom(72);
    body.setMarginLeft(72);
    body.setMarginRight(72);
    
    // Add a native page break after the title page (which is at the top)
    // We find the first student header and insert before it
    const searchResult = body.findText("\\[\\[PAGE_BREAK\\]\\]");
    let currentMatch = searchResult;
    while (currentMatch) {
        const element = currentMatch.getElement();
        const parent = element.getParent();
        const container = parent.getParent();
        
        // Insert page break and remove the marker paragraph
        const index = container.getChildIndex(parent);
        container.insertPageBreak(index);
        container.removeChild(parent);
        
        currentMatch = body.findText("\\[\\[PAGE_BREAK\\]\\]", currentMatch);
    }
    
    // Remove any empty paragraphs at the very start
    while (body.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && 
           body.getChild(0).asParagraph().getText().trim() === "" &&
           body.getNumChildren() > 1) {
        body.removeChild(body.getChild(0));
    }
    
    doc.saveAndClose();
    
    const pdfBlob = DriveApp.getFileById(tempDocMeta.id).getAs('application/pdf');
    pdfBlob.setName(`Submissions - ${assessmentName} (${new Date().toISOString().substring(0,10)}).pdf`);
    
    // Trash temp doc
    DriveApp.getFileById(tempDocMeta.id).setTrashed(true);
    
    // Save PDF to Drive and return URL
    const pdfFile = DriveApp.createFile(pdfBlob);
    
    // Set public sharing so the user can download it immediately via the link
    // Note: This link might be short-lived if we want, but for now we keep it simple.
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { url: pdfFile.getDownloadUrl() };

  } catch (e) {
    Logger.log('Error generating bulk PDF: ' + e.toString());
    return { error: 'Failed to generate report: ' + e.toString() };
  }
}

/**
 * Safely renders a submission answer, allowing ONLY basic formatting and structural tags.
 * This prevents XSS while supporting rich text formatting from contenteditable fields.
 * @param {string} text The raw answer text which may contain HTML
 * @returns {string} The safe HTML string
 */
function renderSubmissionAnswer(text) {
  if (!text) return '';

  // First escape everything to be safe
  let escaped = escapeHtmlBackend(text);

  // Then selectively restore ONLY safe tags that were escaped
  // We handle both lowercase and uppercase, and strip any attributes for security
  return escaped
    .replace(/&lt;b&gt;/gi, '<b>')
    .replace(/&lt;\/b&gt;/gi, '</b>')
    .replace(/&lt;u&gt;/gi, '<u>')
    .replace(/&lt;\/u&gt;/gi, '</u>')
    .replace(/&lt;strong&gt;/gi, '<b>')
    .replace(/&lt;\/strong&gt;/gi, '</b>')
    .replace(/&lt;br\s*\/?&gt;/gi, '<br>')
    .replace(/&lt;div&gt;/gi, '<div>')
    .replace(/&lt;\/div&gt;/gi, '</div>')
    .replace(/&lt;p&gt;/gi, '<p>')
    .replace(/&lt;\/p&gt;/gi, '</p>')
    .replace(/\n/g, '<br>');
}

/**
 * Server-side HTML escaping utility for building email/doc HTML.
 * @param {string} text Text to escape
 * @returns {string} Escaped text
 */
function escapeHtmlBackend(text) {
  if (!text) return '';
  return text.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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

/**
 * Helper: Looks up a student's name from the 'Student Directory' sheet by email.
 * Directory Structure: A) First Name, B) Last Name, C) Student Email
 * @param {string} email The student email to search for
 * @returns {string|null} The combined "First Last" name or null if not found
 */
function getStudentNameFromDirectory(email) {
  if (!email) return null;
  const cleanEmail = email.toLowerCase().trim();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const directorySheet = spreadsheet.getSheetByName('Student Directory');
    if (!directorySheet) return null;
    
    const data = directorySheet.getDataRange().getValues();
    // Start from row 1 to skip headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentEmail = row[2] ? row[2].toString().toLowerCase().trim() : '';
      
      if (studentEmail === cleanEmail) {
        const firstName = row[0] ? row[0].toString().trim() : '';
        const lastName = row[1] ? row[1].toString().trim() : '';
        const fullName = (firstName + ' ' + lastName).trim();
        return fullName || null;
      }
    }
  } catch (e) {
    Logger.log('Error looking up student name in directory: ' + e.toString());
  }
  return null;
}

/**
 * Helper: Gets or creates the Submissions sheet.
 */
function getOrCreateSubmissionsSheet() {
  const lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for the lock
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONSTANTS.SUBMISSIONS_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONSTANTS.SUBMISSIONS_SHEET_NAME);
      // Add headers
      const headers = [
        'Timestamp',
        'Assessment URL',
        'Assessment Name',
        'Student Email',
        'Student Name',
        'Responses JSON'
      ];
      // Set headers
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      // Auto-resize columns appropriately
      sheet.autoResizeColumns(1, headers.length);
    }
    return sheet;
  } catch (e) {
    Logger.log('Error in getOrCreateSubmissionsSheet: ' + e.toString());
    throw e;
  } finally {
    lock.releaseLock();
  }
}

# Vertex AI Batch Processing Guide for Gemini TTS

## Table of Contents
1. [Critical Status Update](#critical-status-update)
2. [Understanding the Batch API Architecture](#understanding-the-batch-api-architecture)
3. [Why Your Current Implementation Fails](#why-your-current-implementation-fails)
4. [Proper Vertex AI Authentication Setup](#proper-vertex-ai-authentication-setup)
5. [Future-Ready Batch Implementation](#future-ready-batch-implementation)
6. [Current Recommended Approach](#current-recommended-approach)
7. [Migration Path](#migration-path)

---

## Critical Status Update

**As of October 2025, the `gemini-2.5-flash-preview-tts` model does NOT support batch prediction on Vertex AI.**

### Evidence

1. **User Reports**: Multiple developers have reported that while Google's documentation lists the TTS models as batch-compatible, actual API calls return 404 errors or "not supported" messages.

2. **Forum Discussions**: The Google AI Developers Forum contains confirmed reports (September 2025) that batch API is not available for `gemini-2.5-flash-preview-tts` despite documentation suggesting otherwise.

3. **Your Errors**: The "unexpected token" and 404 errors you're experiencing are consistent with attempting to use batch prediction on a model that doesn't support it.

### What This Means

- **Batch processing for TTS is not currently possible** - You cannot achieve the 50% cost savings advertised for batch prediction
- **Your current implementation cannot work** - No amount of endpoint or authentication changes will enable batch TTS at this time
- **Manual processing remains necessary** - The real-time API is currently the only way to generate TTS audio
- **Future availability unknown** - Google has not announced when (or if) batch support will be added for TTS models

### Recommendation

**Set `BATCH_API_ENABLED: false` in Constants.js and use manual processing exclusively until Google announces batch support for TTS models.**

The remainder of this document prepares your codebase for future batch support and explains the proper Vertex AI architecture for when it becomes available.

---

## Understanding the Batch API Architecture

### Two Distinct Google AI APIs

Google provides **two separate APIs** for accessing Gemini models, each with different capabilities:

#### 1. Gemini REST API (generativelanguage.googleapis.com)
- **Authentication**: API Key
- **Endpoint**: `https://generativelanguage.googleapis.com/v1beta/`
- **Use Case**: Simple, direct access to Gemini models
- **Batch Support**: Limited - some models support batch via `:batchGenerateContent` endpoint
- **TTS Batch Support**: **NO** - TTS models return 404 when attempting batch operations
- **Current Status**: This is what your code currently uses

#### 2. Vertex AI API (aiplatform.googleapis.com)
- **Authentication**: OAuth 2.0 with Service Account
- **Endpoint**: `https://REGION-aiplatform.googleapis.com/v1/`
- **Use Case**: Enterprise production workloads with advanced features
- **Batch Support**: Full support via batch prediction jobs (for supported models)
- **TTS Batch Support**: **NO** - Not yet available even on Vertex AI
- **IAM Permissions Required**: `roles/aiplatform.user` (includes `aiplatform.endpoints.predict`)
- **Future Path**: When batch TTS becomes available, it will likely be here first

### Why Vertex AI Is Different

The permission error you discovered (`aiplatform.endpoints.predict`) is specific to **Vertex AI**, not the Gemini REST API. This reveals a fundamental architectural difference:

- **Gemini REST API**: Uses simple API keys, no IAM permissions required
- **Vertex AI**: Uses Google Cloud IAM with service accounts, requiring specific roles and permissions

Your documentation mentioned needing the Vertex AI endpoint, which is correct for full batch support, but **also requires completely different authentication**.

---

## Why Your Current Implementation Fails

### The Authentication Mismatch

Your current code in [Code.js:252](Code.js#L252):

```javascript
const batchCreateUrl = `${CONSTANTS.GEMINI_BATCH_API_ENDPOINT}models/${CONSTANTS.GEMINI_TTS_MODEL}:batchGenerateContent?key=${apiKey}`;
```

This attempts to:
- Use the **Gemini REST API endpoint** format
- Authenticate with an **API key** (`?key=${apiKey}`)
- Call `:batchGenerateContent` on a **TTS model**

### Why This Cannot Work

1. **TTS models don't support batch**: The `gemini-2.5-flash-preview-tts` model has no `:batchGenerateContent` endpoint
2. **Wrong endpoint for Vertex AI**: If you were using Vertex AI, the endpoint format would be completely different
3. **Wrong authentication method**: Vertex AI batch jobs require OAuth 2.0 Bearer tokens, not API keys

### The "Unexpected Token" Error

This error in your recent commit likely comes from:
1. The API returning a 404 or error HTML page
2. Your code trying to parse it as JSON with `JSON.parse()`
3. Receiving HTML like `<!DOCTYPE html>` instead of `{`

---

## Proper Vertex AI Authentication Setup

**Note**: This section prepares you for when batch TTS becomes available. Do not implement this until Google confirms batch support for TTS models.

### Step 1: Enable Required APIs

In your Google Cloud Console (https://console.cloud.google.com):

1. Navigate to **APIs & Services** > **Library**
2. Enable the following APIs:
   - **Vertex AI API** (`aiplatform.googleapis.com`)
   - **Generative Language API** (`generativelanguage.googleapis.com`) - already enabled
3. Click **Enable** for each

### Step 2: Create Service Account

1. Go to **IAM & Admin** > **Service Accounts**
2. Click **Create Service Account**
3. Fill in details:
   - **Name**: `spartan-read-aloud-tts`
   - **ID**: Auto-generated
   - **Description**: "Service account for batch TTS processing"
4. Click **Create and Continue**

### Step 3: Grant IAM Roles

On the "Grant this service account access to project" step:

1. Add role: **Vertex AI User** (`roles/aiplatform.user`)
   - This grants:
     - `aiplatform.endpoints.predict` - Required for model inference
     - `aiplatform.batchPredictionJobs.create` - Required for batch jobs
     - `aiplatform.batchPredictionJobs.get` - Required for status checks
2. Add role: **Storage Object Viewer** (`roles/storage.objectViewer`)
   - Required if batch jobs use Cloud Storage for input/output
3. Click **Continue**, then **Done**

### Step 4: Create Service Account Key

1. Click on your newly created service account
2. Go to **Keys** tab
3. Click **Add Key** > **Create new key**
4. Choose **JSON** format
5. Click **Create**
6. **Save the downloaded JSON file securely** - you'll need its contents

### Step 5: Add OAuth2 Library to Apps Script

1. In Apps Script Editor, click the **+** next to **Libraries**
2. Enter Script ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`
3. Click **Look up**
4. Select the latest version
5. Set Identifier to: `OAuth2`
6. Click **Add**

### Step 6: Store Service Account Credentials

Open the downloaded JSON file. You'll see something like:

```json
{
  "type": "service_account",
  "project_id": "your-project-id",
  "private_key_id": "abc123...",
  "private_key": "-----BEGIN PRIVATE KEY-----\n...",
  "client_email": "spartan-read-aloud-tts@your-project.iam.gserviceaccount.com",
  "client_id": "123456789",
  ...
}
```

Add these to Script Properties:

1. In Apps Script Editor: **Project Settings** > **Script Properties**
2. Add properties:
   - `SERVICE_ACCOUNT_EMAIL`: The `client_email` value
   - `SERVICE_ACCOUNT_PRIVATE_KEY`: The entire `private_key` value (including `-----BEGIN/END PRIVATE KEY-----`)
   - `GCP_PROJECT_ID`: The `project_id` value
   - `VERTEX_AI_REGION`: `us-central1` (or your preferred region)

### Step 7: Update appsscript.json Scopes

Add the OAuth scope for Google Cloud Platform:

```json
{
  "timeZone": "America/New_York",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v2",
        "serviceId": "drive"
      }
    ],
    "libraries": [
      {
        "userSymbol": "OAuth2",
        "version": "43",
        "libraryId": "1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/cloud-platform"
  ]
}
```

The `https://www.googleapis.com/auth/cloud-platform` scope is required for Vertex AI access.

---

## Future-Ready Batch Implementation

**This section contains code to implement when Google enables batch support for TTS models.**

### Architecture Overview

When batch TTS becomes available, the workflow will be:

1. **Upload Input File**: Create JSONL file and upload to Google Cloud Storage
2. **Create Batch Job**: Submit batch prediction job to Vertex AI
3. **Poll Job Status**: Check job status periodically via trigger
4. **Retrieve Results**: Download output JSONL from Cloud Storage
5. **Process Audio**: Convert base64 audio to WAV files and store in Drive

### New Helper: OAuth2 Service

Add to `Code.js`:

```javascript
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
```

### Update Constants.js

Replace the batch endpoint:

```javascript
const CONSTANTS = {
  // ... existing constants ...

  // --- Vertex AI Configuration ---
  VERTEX_AI_PROJECT_ID: PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID'),
  VERTEX_AI_REGION: 'us-central1', // Or from script properties
  VERTEX_AI_ENDPOINT: 'https://us-central1-aiplatform.googleapis.com/v1',

  // --- Gemini API ---
  GEMINI_TTS_MODEL: 'gemini-2.5-flash-preview-tts',
  GEMINI_API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/',
  GEMINI_VOICE_NAME: "Kore",

  // ... rest of constants ...
};
```

### Updated Batch Job Submission

Replace `submitGeminiBatchJob()` in `Code.js`:

```javascript
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
        outputUriPrefix: `gs://your-output-bucket/batch-tts-results/${displayName}/`
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
```

### Updated Status Checking

Replace `checkGeminiBatchJobStatus()`:

```javascript
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
```

### Updated Result Processing

Replace `processBatchJobResults()`:

```javascript
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

    // Save JSON to spreadsheet
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
```

### Cloud Storage Bucket Setup

Before batch processing will work, you need to create a GCS bucket:

1. Go to Cloud Storage in Google Cloud Console
2. Click **Create Bucket**
3. Name it (e.g., `spartan-tts-batch-processing`)
4. Choose region matching your Vertex AI region (`us-central1`)
5. Set storage class to **Standard**
6. Set access control to **Uniform**
7. Click **Create**
8. Add to Script Properties:
   - Key: `GCS_BUCKET_NAME`
   - Value: `spartan-tts-batch-processing` (your bucket name)

---

## Current Recommended Approach

**Until Google enables batch support for TTS models, use this optimized manual processing approach:**

### 1. Disable Batch Processing

In `Constants.js`:

```javascript
BATCH_API_ENABLED: false,
```

### 2. Optimize Manual Processing

Consider these improvements to speed up manual processing:

#### A. Add Progress Logging

```javascript
function step2_GenerateMissingAudioAndFinalize() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessment Database');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  let totalToProcess = 0;
  let processedCount = 0;

  // Count how many need processing
  for (let i = 1; i < data.length; i++) {
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];
    const audioJson = data[i][CONSTANTS.COL.AUDIO_JSON];
    if (chunkCount && !audioJson) totalToProcess++;
  }

  Logger.log(`Starting manual processing: ${totalToProcess} assessments to process`);

  for (let i = 1; i < data.length; i++) {
    const chunkCount = data[i][CONSTANTS.COL.CHUNK_COUNT];
    const audioJson = data[i][CONSTANTS.COL.AUDIO_JSON];

    if (chunkCount && !audioJson) {
      processedCount++;
      Logger.log(`Processing ${processedCount}/${totalToProcess}...`);

      generateAudioForRow(i + 1, data[i]);

      // Update status in sheet
      sheet.getRange(i + 1, CONSTANTS.COL.PROCESSING_STATUS + 1)
           .setValue(`MANUAL (${processedCount}/${totalToProcess})`);
      SpreadsheetApp.flush();
    }
  }

  Logger.log(`Manual processing complete: ${processedCount} assessments processed`);
}
```

#### B. Add Retry Logic

```javascript
function generateAudioFromTextChunkWithRetry(text, fileName, folder, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = generateAudioFromTextChunk(text, fileName, folder);
      if (result) return result;

      Logger.log(`Attempt ${attempt}/${maxRetries} failed for ${fileName}, retrying...`);
      Utilities.sleep(1000 * attempt); // Exponential backoff
    } catch (e) {
      Logger.log(`Exception on attempt ${attempt}/${maxRetries}: ${e.toString()}`);
      if (attempt === maxRetries) throw e;
      Utilities.sleep(1000 * attempt);
    }
  }
  return null;
}
```

#### C. Add Cost Tracking

```javascript
function trackAPIUsage(chunkCount) {
  const props = PropertiesService.getScriptProperties();
  const currentCount = parseInt(props.getProperty('TOTAL_CHUNKS_PROCESSED') || '0');
  const newCount = currentCount + chunkCount;

  props.setProperty('TOTAL_CHUNKS_PROCESSED', newCount.toString());
  props.setProperty('LAST_PROCESSING_DATE', new Date().toISOString());

  // Rough cost estimate: $0.05 per 1000 characters (adjust based on actual pricing)
  // Assume average chunk is 500 characters
  const estimatedCost = (newCount * 500 * 0.05) / 1000;

  Logger.log(`Total chunks processed: ${newCount}`);
  Logger.log(`Estimated API cost to date: $${estimatedCost.toFixed(2)}`);
}
```

### 3. Monitor API Quotas

Add to your menu:

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu(CONSTANTS.MENU_NAME)
      .addItem(CONSTANTS.MENU_ITEMS.RUN_MANUAL, 'runAllStepsManual')
      .addSeparator()
      .addItem('View API Usage Stats', 'showAPIUsageStats')
      .addToUi();
}

function showAPIUsageStats() {
  const props = PropertiesService.getScriptProperties();
  const totalChunks = props.getProperty('TOTAL_CHUNKS_PROCESSED') || '0';
  const lastDate = props.getProperty('LAST_PROCESSING_DATE') || 'Never';

  const message = `
API Usage Statistics:
━━━━━━━━━━━━━━━━━━━━
Total chunks processed: ${totalChunks}
Last processing: ${lastDate}

Note: Batch API for TTS not yet available.
Estimated savings when batch becomes available: 50%
  `.trim();

  SpreadsheetApp.getUi().alert('API Usage', message, SpreadsheetApp.getUi().ButtonSet.OK);
}
```

---

## Migration Path

### When Google Announces Batch Support for TTS

1. **Monitor Official Channels**:
   - Google AI Developer Blog: https://developers.googleblog.com/
   - Vertex AI Release Notes: https://cloud.google.com/vertex-ai/docs/release-notes
   - Gemini API Changelog: https://ai.google.dev/gemini-api/docs/changelog

2. **Verification Checklist**:
   - [ ] Official documentation confirms `gemini-2.5-flash-preview-tts` supports batch
   - [ ] Release notes specify the correct endpoint format
   - [ ] Example code is provided by Google
   - [ ] Pricing for batch TTS is published

3. **Implementation Steps**:
   - [ ] Complete [Proper Vertex AI Authentication Setup](#proper-vertex-ai-authentication-setup)
   - [ ] Set up Cloud Storage bucket
   - [ ] Update code with [Future-Ready Batch Implementation](#future-ready-batch-implementation)
   - [ ] Test with a single small assessment
   - [ ] Verify cost savings in billing
   - [ ] Set `BATCH_API_ENABLED: true`
   - [ ] Roll out to production

4. **Rollback Plan**:
   - If any issues occur, simply set `BATCH_API_ENABLED: false`
   - All batch columns in spreadsheet can remain (they'll just be unused)
   - No data loss occurs

---

## Additional Resources

### Official Documentation
- [Vertex AI Batch Prediction](https://cloud.google.com/vertex-ai/docs/predictions/get-batch-predictions)
- [Gemini API Documentation](https://ai.google.dev/gemini-api/docs)
- [Service Account Authentication](https://cloud.google.com/docs/authentication/provide-credentials-adc#service-account)
- [Apps Script OAuth2 Library](https://github.com/googleworkspace/apps-script-oauth2)

### Community Resources
- [Google AI Developers Forum](https://discuss.ai.google.dev/)
- [Stack Overflow: vertex-ai tag](https://stackoverflow.com/questions/tagged/vertex-ai)

### Monitoring
- [Google Cloud Console - Vertex AI](https://console.cloud.google.com/vertex-ai)
- [API Usage Dashboard](https://console.cloud.google.com/apis/dashboard)
- [Billing Reports](https://console.cloud.google.com/billing)

---

## Summary

**Current Status**: Batch processing for `gemini-2.5-flash-preview-tts` is **not available** as of October 2025.

**Action Required**: Set `BATCH_API_ENABLED: false` and use optimized manual processing.

**Future Path**: This document provides a complete implementation ready for when Google enables batch TTS support, including proper Vertex AI authentication with service accounts and IAM permissions.

**Permission Context**: The `roles/aiplatform.user` role (with `aiplatform.endpoints.predict` permission) you discovered is indeed required for Vertex AI batch operations, confirming that proper batch support requires the full Vertex AI architecture, not just the Gemini REST API.

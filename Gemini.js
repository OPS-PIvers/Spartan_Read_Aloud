/**
 * Main entry point for audio generation. Routes to the configured TTS provider
 * with automatic fallback to Gemini if the primary provider fails.
 * @param {string} text The text to convert to speech.
 * @param {string} fileName The desired, complete filename for the output file.
 * @param {GoogleAppsScript.Drive.Folder} folder The Drive folder to save the file in.
 * @returns {GoogleAppsScript.Drive.File|null} The created audio file or null on failure.
 */
function generateAudio(text, fileName, folder) {
  let audioFile = null;

  // Try the configured provider first
  if (CONSTANTS.TTS_PROVIDER === 'GEMINI') {
    Logger.log(`-> Using Gemini TTS (configured provider)...`);
    audioFile = generateAudioFromTextChunk(text, fileName, folder);
  } else if (CONSTANTS.TTS_PROVIDER === 'GOOGLE_CLOUD') {
    Logger.log(`-> Using Google Cloud TTS (configured provider)...`);
    audioFile = generateAudioWithStandardVoice(text, fileName, folder);

    // Fallback to Gemini if Google Cloud TTS fails
    if (!audioFile) {
      Logger.log('-> Google Cloud TTS failed, falling back to Gemini TTS...');
      audioFile = generateAudioFromTextChunk(text, fileName, folder);
    }
  } else {
    Logger.log(`-> ERROR: Unknown TTS_PROVIDER "${CONSTANTS.TTS_PROVIDER}". Defaulting to Gemini.`);
    audioFile = generateAudioFromTextChunk(text, fileName, folder);
  }

  return audioFile;
}

/**
 * Calls the Gemini API to generate audio from a text chunk.
 * @param {string} text The text to convert to speech.
 * @param {string} fileName The desired, complete filename for the output file (e.g., 'assessment-chunk-1.wav').
 * @param {GoogleAppsScript.Drive.Folder} folder The Drive folder to save the file in.
 * @returns {GoogleAppsScript.Drive.File|null} The created audio file or null on failure.
 */
function generateAudioFromTextChunk(text, fileName, folder) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const model = CONSTANTS.GEMINI_TTS_MODEL;
  const url = `${CONSTANTS.GEMINI_API_BASE_URL}models/${model}:generateContent?key=${apiKey}`;

  // --- FIX: Corrected the payload structure ---
  const payload = {
    model: CONSTANTS.GEMINI_TTS_MODEL,
    contents: [{
      parts: [{
        text: `Read the following text in a clear, neutral, and steady voice: ${text}`
      }]
    }],
    generationConfig: {
      responseModalities: ["AUDIO"],
      speechConfig: { // Moved speechConfig inside generationConfig
        voiceConfig: {
          prebuiltVoiceConfig: { voiceName: CONSTANTS.GEMINI_VOICE_NAME }
        }
      }
    }
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    Logger.log(`-> Calling Gemini API for chunk: "${fileName}"...`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const audioData = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data;

      if (audioData) {
        const decodedData = Utilities.base64Decode(audioData);
        const wavBlob = createWavBlob(decodedData);
        const wavFile = folder.createFile(wavBlob.setName(fileName));
        return wavFile;
      } else {
        Logger.log('-> ERROR: Gemini API response was successful, but contained no audio data.');
        return null;
      }
    } else {
      Logger.log(`-> ERROR: Gemini API returned a non-200 response. Code: ${responseCode}. Body: ${responseBody}`);
      return null;
    }
  } catch (e) {
    Logger.log(`-> EXCEPTION during Gemini API call: ${e.toString()}`);
    return null;
  }
}

/**
 * Calls the Google Cloud Text-to-Speech API to generate audio using a Standard voice.
 * Uses the cheaper "Standard" voice models (4M free characters per month).
 * @param {string} text The text to convert to speech.
 * @param {string} fileName The desired, complete filename for the output file.
 * @param {GoogleAppsScript.Drive.Folder} folder The Drive folder to save the file in.
 * @returns {GoogleAppsScript.Drive.File|null} The created audio file or null on failure.
 */
function generateAudioWithStandardVoice(text, fileName, folder) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  // Use the Cloud Text-to-Speech API endpoint (not Gemini endpoint)
  const url = `https://texttospeech.googleapis.com/v1/text:synthesize?key=${apiKey}`;

  // Add SSML pauses to text for more natural pacing
  const processedText = addPausesToText(text);

  // Determine if we're using SSML or plain text
  const isSSML = processedText.startsWith('<speak>');

  // Payload structure for Cloud TTS API
  const payload = {
    "input": isSSML ? { "ssml": processedText } : { "text": processedText },
    "voice": {
      "languageCode": "en-US",
      "name": CONSTANTS.GOOGLE_CLOUD_TTS_VOICE
    },
    "audioConfig": {
      "audioEncoding": "LINEAR16", // Raw PCM data compatible with createWavBlob
      "sampleRateHertz": CONSTANTS.WAV_SAMPLE_RATE
    }
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    Logger.log(`-> Calling Google Cloud TTS API for chunk: "${fileName}"...`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);

      // Get the audio from 'audioContent' (it's a base64 string)
      const audioData = jsonResponse.audioContent;

      if (audioData) {
        // Decode the base64 string into raw bytes
        const decodedData = Utilities.base64Decode(audioData);

        // Pass the raw bytes to the existing WAV creation function
        const wavBlob = createWavBlob(decodedData);
        const wavFile = folder.createFile(wavBlob.setName(fileName));
        return wavFile;
      } else {
        Logger.log('-> ERROR: Google Cloud TTS API response was successful, but contained no audio data.');
        return null;
      }
    } else {
      Logger.log(`-> ERROR: Google Cloud TTS API returned a non-200 response. Code: ${responseCode}. Body: ${responseBody}`);
      return null;
    }
  } catch (e) {
    Logger.log(`-> EXCEPTION during Google Cloud TTS API call: ${e.toString()}`);
    return null;
  }
}

/**
 * Creates a valid WAV file blob from raw 16-bit PCM audio data.
 * @param {byte[]} pcmData The raw audio data from the API.
 * @return {GoogleAppsScript.Base.Blob} A blob representing the WAV file.
 */
function createWavBlob(pcmData) {
  const sampleRate = CONSTANTS.WAV_SAMPLE_RATE; // Gemini TTS standard sample rate
  const numChannels = CONSTANTS.WAV_NUM_CHANNELS;
  const bitsPerSample = CONSTANTS.WAV_BITS_PER_SAMPLE;
  const byteRate = sampleRate * numChannels * bitsPerSample / 8;
  const blockAlign = numChannels * bitsPerSample / 8;
  const dataSize = pcmData.length;
  const fileSize = 36 + dataSize;

  const buffer = new ArrayBuffer(44);
  const view = new DataView(buffer);

  // RIFF header
  writeString(view, 0, 'RIFF');
  view.setUint32(4, fileSize, true);
  writeString(view, 8, 'WAVE');
  
  // "fmt " sub-chunk
  writeString(view, 12, 'fmt ');
  view.setUint32(16, 16, true); // Sub-chunk size
  view.setUint16(20, 1, true);  // Audio format (1 for PCM)
  view.setUint16(22, numChannels, true);
  view.setUint32(24, sampleRate, true);
  view.setUint32(28, byteRate, true);
  view.setUint16(32, blockAlign, true);
  view.setUint16(34, bitsPerSample, true);

  // "data" sub-chunk
  writeString(view, 36, 'data');
  view.setUint32(40, dataSize, true);

  const headerBytes = Array.from(new Uint8Array(buffer));
  const wavBytes = headerBytes.concat(pcmData);

  return Utilities.newBlob(wavBytes, MimeType.WAV);
}

/**
 * Helper to write a string to a DataView.
 * @param {DataView} view The DataView to write to.
 * @param {number} offset The byte offset.
 * @param {string} string The string to write.
 */
function writeString(view, offset, string) {
  for (let i = 0; i < string.length; i++) {
    view.setUint8(offset + i, string.charCodeAt(i));
  }
}


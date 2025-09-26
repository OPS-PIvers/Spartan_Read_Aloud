/**
 * Calls the Gemini API to generate audio from a text chunk.
 * @param {string} text The text to convert to speech.
 * @param {string} fileName The desired, complete filename for the output file (e.g., 'assessment-chunk-1.wav').
 * @param {GoogleAppsScript.Drive.Folder} folder The Drive folder to save the file in.
 * @returns {GoogleAppsScript.Drive.File|null} The created audio file or null on failure.
 */
function generateAudioFromTextChunk(text, fileName, folder) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const model = 'gemini-2.5-flash-preview-tts';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  // --- FIX: Corrected the payload structure ---
  const payload = {
    model: "gemini-2.5-flash-preview-tts",
    contents: [{
      parts: [{
        text: `Read the following text in a clear, neutral, and steady voice: ${text}`
      }]
    }],
    generationConfig: {
      responseModalities: ["AUDIO"],
      speechConfig: { // Moved speechConfig inside generationConfig
        voiceConfig: {
          prebuiltVoiceConfig: { voiceName: "Kore" }
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
 * Creates a valid WAV file blob from raw 16-bit PCM audio data.
 * @param {byte[]} pcmData The raw audio data from the API.
 * @return {GoogleAppsScript.Base.Blob} A blob representing the WAV file.
 */
function createWavBlob(pcmData) {
  const sampleRate = 24000; // Gemini TTS standard sample rate
  const numChannels = 1;
  const bitsPerSample = 16;
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


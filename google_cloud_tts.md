Based on your app's code, here is how you can modify it to use the cheaper "Standard" voice models.

To use the "Standard" voices, you need to call the **Cloud Text-to-Speech API** endpoint, which is different from the Gemini API endpoint your app is currently using.

### Your Current Setup

Your `Constants.js` file specifies the `gemini-2.5-flash-preview-tts` model, and your `Gemini.js` file has a function `generateAudioFromTextChunk` that formats a request for that specific Gemini API.

The "Standard" voices you asked about are "Legacy TTS models" from the pricing page. They use a different API with a different request structure.

-----

### How to Modify Your App

You can add a new function to your `Gemini.js` file that calls the correct API. The main differences are the URL, the JSON payload, and the fact that the audio is returned as a base64 string (which needs decoding).

#### 1\. Add this New Function to `Gemini.js`

This function is designed to work with your existing `createWavBlob` function.

```javascript
/**
 * Calls the legacy Cloud Text-to-Speech API to generate audio using a Standard voice.
 * @param {string} text The text to convert to speech.
 * @param {string} fileName The desired, complete filename for the output file.
 * @param {GoogleAppsScript.Drive.Folder} folder The Drive folder to save the file in.
 * @returns {GoogleAppsScript.Drive.File|null} The created audio file or null on failure.
 */
function generateAudioWithStandardVoice(text, fileName, folder) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  // 1. Use the legacy Text-to-Speech API endpoint
  const url = `https://texttospeech.googleapis.com/v1/text:synthesize?key=${apiKey}`;

  // 2. Use the payload structure for the legacy API
  const payload = {
    "input": {
      "text": text
    },
    "voice": {
      "languageCode": "en-US",
      "name": "en-US-Standard-A" // Example Standard voice
    },
    "audioConfig": {
      "audioEncoding": "LINEAR16", // This is the raw PCM data your createWavBlob function expects
      "sampleRateHertz": CONSTANTS.WAV_SAMPLE_RATE // Use sample rate from Constants.js
    }
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    Logger.log(`-> Calling legacy TTS API for chunk: "${fileName}"...`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      
      // 3. Get the audio from 'audioContent' (it's a base64 string)
      const audioData = jsonResponse.audioContent;

      if (audioData) {
        // 4. Decode the base64 string into raw bytes
        const decodedData = Utilities.base64Decode(audioData);
        
        // 5. Pass the raw bytes to your existing WAV creation function
        const wavBlob = createWavBlob(decodedData); 
        const wavFile = folder.createFile(wavBlob.setName(fileName));
        return wavFile;
      } else {
        Logger.log('-> ERROR: Legacy TTS API response was successful, but contained no audio data.');
        return null;
      }
    } else {
      Logger.log(`-> ERROR: Legacy TTS API returned a non-200 response. Code: ${responseCode}. Body: ${responseBody}`);
      return null;
    }
  } catch (e) {
    Logger.log(`-> EXCEPTION during legacy TTS API call: ${e.toString()}`);
    return null;
  }
}
```

#### 2\. Update `Code.js` to Call the New Function

In your `Code.js` file, find the `step2_GenerateMissingAudioAndFinalize` function. Inside its `for` loop, change the line that calls `generateAudioFromTextChunk`:

**Change this:**

```javascript
audioFile = generateAudioFromTextChunk(chunkText, newChunkName, assessmentSubfolder);
```

**To this:**

```javascript
audioFile = generateAudioWithStandardVoice(chunkText, newChunkName, assessmentSubfolder);
```

-----

### STANDARD VOICE NAMES:

1st option: en-US-Standard-H (female)
2nd option: en-US-Standard-I (male)


### Possible GAS Library for Speech API:

#### Speech.gs

class Speech {

  constructor({
    projectNumber,
    tokenService,
    locationId = 'eu',
    fieldMask,
    apiKey,
    voice,
    gcsUri,
    audioConfig,
    noCache = false,
    noisy = false
  } = {}) {
    // sync
    // https://texttospeech.googleapis.com/v1/text:synthesize

    // operations
    // https://texttospeech.googleapis.com/v1/{parent=projects/*/locations/*}
    this.endpoint = "https://texttospeech.googleapis.com/v1"
    this.tokenService = tokenService
    this.projectNumber = projectNumber
    this.locationId = locationId
    this.voice = {
      languageCode: "en-US",
      ...voice
    }
    this.gcsUri = gcsUri
    this.audioConfig = {
      audioEncoding: 'MP3',
      ...audioConfig
    }
    this.noCache = noCache
    this.noisy = noisy
    this.basePayload = {
    }
    if (fieldMask) {
      if (!Array.isArray(fieldMask)) throw 'fieldmask must be an array'
      this.basePayload.fieldMask = fieldMask.join(",")
    }

    // use this fetcher with a shortish cache life for getting processors
    const fetchOptions = {
      endpoint: this.endpoint,
      tokenService: this.tokenService,
      defaultParams: apiKey ? { key: apiKey } : null
    }

    this.fetcher = Exports.newFetch(fetchOptions)

  }

  get syncProcessEndpoint() {
    return '/text:synthesize'
  }



  getVoice(voice = {}) {
    return {
      voice: {
        ...this.voice,
        ...voice
      }
    }
  }

  getOutputGcsUri(gcsUri = this.gcsUri) {
    return {
      gcsUri
    }
  }


  getSynthesesInput({ ssml, text }) {
    if (ssml && text) throw 'Provide either ssml or text property only as input'
    if (!ssml && !text) throw 'Provide either ssml or text property as input - there was neither'
    const input = ssml ? { ssml } : { text }
    return {
      input
    }
  }

  getAudioConfig(audioConfig = {}) {
    // this ping is really just to allow any defaults to be set from the constructor as a future enhancement
    return {
      audioConfig: {
        audioEncoding: "MP3",
        ...this.audioConfig,
        ...audioConfig
      }
    }
  }

  /**
   * synthesize audio 
   * @return {Operation}  and operation response
   */
  synthesize(input, { audioConfig, voice } = {}, { noCache = this.noCache, noisy = this.noisy } = {}) {
    const payload = {
      ...this.getSynthesesInput(input),
      ...this.getAudioConfig(audioConfig),
      ...this.getVoice(voice)
    }

    // https://texttospeech.googleapis.com/v1/{parent=projects/*/locations/*}:synthesizeLongAudio
    const result = this._post({
      noCache,
      noisy,
      processEndpoint: this.endpoint + this.syncProcessEndpoint,
      payload
    })

    return {
      ...result,
      mp3: Exports.Utils.b64ToBlob(result.data.audioContent, 'audio/mpeg')
    }

  }

  // not implelemted yet

  /**
   * synthesize audio in batch
   * @return {Operation}  and operation response
   */
  synthesizeLongAudio(input, { audioConfig, gcsUri, voice } = {}, { noCache, noisy } = {}) {
    const payload = {
      ...this.getSynthesesInput(input),
      ...this.getAudioConfig(audioConfig),
      ...this.getOutputGcsUri(gcsUri),
      ...this.getVoice(voice)
    }

    // https://texttospeech.googleapis.com/v1/{parent=projects/*/locations/*}:synthesizeLongAudio
    const result = this._post({
      noCache,
      noisy,
      processEndpoint: '',
      payload,
      type: "synthesizeLongAudio"
    })

    return result

  }


  _post({
    noCache = false,
    noisy = false,
    processEndpoint,
    payload,
    type
  }) {

    const cacher = this.fetcher.cacher
    const keyer = (url, options) => {
      return cacher.keyer(
        options.payload,
        options.method + url
      )
    }
    // fetcher options
    const options = this.getOptions({
      payload: JSON.stringify(payload)
    })

    // tweak the processor path
    const url = type ? processEndpoint.replace(/(.*):(.*$)/, `$1${type}`) : processEndpoint

    // args for the fetcher
    const args = {
      options,
      noCache,
      noisy,
      keyer,
      url
    }


    // do the fetch - the result is base64 encoded
    return this.fetcher.fetch(args)
  }



  /**
   * get the fetch options
   */
  getOptions(options = {}) {
    return {
      method: "POST",
      contentType: "application/json; charset=utf-8",
      ...options
    }
  }

}
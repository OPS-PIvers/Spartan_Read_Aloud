/**
 * Tests for createWavHeader function.
 * This function can be run manually in the Apps Script editor.
 */
function testCreateWavHeader() {
  Logger.log('=== Starting testCreateWavHeader ===');

  const sampleRate = CONSTANTS.WAV_SAMPLE_RATE;
  const numChannels = CONSTANTS.WAV_NUM_CHANNELS;
  const bitsPerSample = CONSTANTS.WAV_BITS_PER_SAMPLE;
  const byteRate = sampleRate * numChannels * bitsPerSample / 8;
  const blockAlign = numChannels * bitsPerSample / 8;

  const dataSize = 1000;
  const header = createWavHeader(dataSize);

  // 1. Check Header Length
  if (header.length !== 44) {
    Logger.log('FAIL: Header length is not 44 bytes.');
  } else {
    Logger.log('PASS: Header length is 44 bytes.');
  }

  const view = new DataView(new Uint8Array(header).buffer);

  // Helper to read string from DataView
  function readString(v, offset, length) {
    let str = '';
    for (let i = 0; i < length; i++) {
      str += String.fromCharCode(v.getUint8(offset + i));
    }
    return str;
  }

  // 2. Check RIFF
  const riff = readString(view, 0, 4);
  if (riff !== 'RIFF') {
    Logger.log(`FAIL: Expected RIFF, got ${riff}`);
  } else {
    Logger.log('PASS: RIFF tag correct.');
  }

  // 3. Check File Size (36 + dataSize)
  const fileSize = view.getUint32(4, true);
  if (fileSize !== 36 + dataSize) {
    Logger.log(`FAIL: Expected file size ${36 + dataSize}, got ${fileSize}`);
  } else {
    Logger.log('PASS: File size correct.');
  }

  // 4. Check WAVE
  const wave = readString(view, 8, 4);
  if (wave !== 'WAVE') {
    Logger.log(`FAIL: Expected WAVE, got ${wave}`);
  } else {
    Logger.log('PASS: WAVE tag correct.');
  }

  // 5. Check fmt
  const fmt = readString(view, 12, 4);
  if (fmt !== 'fmt ') {
    Logger.log(`FAIL: Expected 'fmt ', got '${fmt}'`);
  } else {
    Logger.log('PASS: fmt tag correct.');
  }

  // 6. Check Subchunk1Size (16 for PCM)
  const subchunk1Size = view.getUint32(16, true);
  if (subchunk1Size !== 16) {
    Logger.log(`FAIL: Expected Subchunk1Size 16, got ${subchunk1Size}`);
  } else {
    Logger.log('PASS: Subchunk1Size correct.');
  }

  // 7. Check AudioFormat (1 for PCM)
  const audioFormat = view.getUint16(20, true);
  if (audioFormat !== 1) {
    Logger.log(`FAIL: Expected AudioFormat 1, got ${audioFormat}`);
  } else {
    Logger.log('PASS: AudioFormat correct.');
  }

  // 8. Check NumChannels
  const nc = view.getUint16(22, true);
  if (nc !== numChannels) {
    Logger.log(`FAIL: Expected NumChannels ${numChannels}, got ${nc}`);
  } else {
    Logger.log('PASS: NumChannels correct.');
  }

  // 9. Check SampleRate
  const sr = view.getUint32(24, true);
  if (sr !== sampleRate) {
    Logger.log(`FAIL: Expected SampleRate ${sampleRate}, got ${sr}`);
  } else {
    Logger.log('PASS: SampleRate correct.');
  }

  // 10. Check ByteRate
  const br = view.getUint32(28, true);
  if (br !== byteRate) {
    Logger.log(`FAIL: Expected ByteRate ${byteRate}, got ${br}`);
  } else {
    Logger.log('PASS: ByteRate correct.');
  }

  // 11. Check BlockAlign
  const ba = view.getUint16(32, true);
  if (ba !== blockAlign) {
    Logger.log(`FAIL: Expected BlockAlign ${blockAlign}, got ${ba}`);
  } else {
    Logger.log('PASS: BlockAlign correct.');
  }

  // 12. Check BitsPerSample
  const bps = view.getUint16(34, true);
  if (bps !== bitsPerSample) {
    Logger.log(`FAIL: Expected BitsPerSample ${bitsPerSample}, got ${bps}`);
  } else {
    Logger.log('PASS: BitsPerSample correct.');
  }

  // 13. Check data tag
  const dataTag = readString(view, 36, 4);
  if (dataTag !== 'data') {
    Logger.log(`FAIL: Expected 'data', got '${dataTag}'`);
  } else {
    Logger.log('PASS: data tag correct.');
  }

  // 14. Check Data Size
  const ds = view.getUint32(40, true);
  if (ds !== dataSize) {
    Logger.log(`FAIL: Expected DataSize ${dataSize}, got ${ds}`);
  } else {
    Logger.log('PASS: DataSize correct.');
  }

  Logger.log('=== Finished testCreateWavHeader ===');
}

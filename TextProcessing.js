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

  // Initialize with isHeader flag
  let currentChunk = { text: '', ids: [], isHeader: false };

  // Helper to finish current chunk and start a new one
  const commitChunk = () => {
    const trimmedText = currentChunk.text.trim();
    if (trimmedText) {
      // Create clean chunk object without internal flags
      const chunkToPush = {
        text: trimmedText,
        ids: [...currentChunk.ids] // Clone array
      };
      chunks.push(chunkToPush);
    }
    // Reset chunk state
    currentChunk = { text: '', ids: [], isHeader: false };
  };

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];

    // --- Detection Logic ---

    // 1. Question Start: "1.", "1)", "Q1", "(1)"
    const isQuestionStart = /^(?:\(?\d+|Q\d+)[.)\]]/.test(block.text);

    // 2. Answer Option: "a.", "b.", "A)", "(a)"
    const isAnswerOption = /^(?:\(?[a-zA-Z][.)]|[a-zA-Z]\.)(?:\s|$)/.test(block.text);

    // 3. Header
    const isHeader = /^h[1-6]/.test(block.tag);

    // 4. Metadata/Directions (e.g., "Directions:", "Read the following...")
    const isDirections = /^(directions|instructions|read|note):/i.test(block.text);

    // --- Grouping Decision Matrix ---

    if (currentChunk.text === '') {
      // Start of new chunk
      currentChunk.text = block.text;
      currentChunk.ids.push(block.id);
      if (isHeader || isDirections) currentChunk.isHeader = true;
    }
    else if (isAnswerOption) {
      // Rule: Merge answer options with the preceding chunk (likely the question or previous option)
      // This MUST come before isQuestionStart to ensure options are grouped even if the question was long
      currentChunk.text += '\n' + block.text;
      currentChunk.ids.push(block.id);
      // Answer options don't inherit header status
      currentChunk.isHeader = false;
    }
    else if (isQuestionStart || isHeader || isDirections) {
      // Rule: Questions, Headers, and Directions start a new thought/context
      // Priority 1: Force break before a new question or header
      commitChunk();
      currentChunk.text = block.text;
      currentChunk.ids.push(block.id);
      if (isHeader || isDirections) {
        currentChunk.isHeader = true;
      }
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
      // EXCEPTION: If previous was a header (or directions), always split
      if (!currentChunk.isHeader && (!previousEndedSentence || isPreviousShort)) {
         // Merge for flow
         currentChunk.text += ' ' + block.text; // Use space for flowing text
         currentChunk.ids.push(block.id);
      } else {
         // Previous was a complete thought and long enough (or was a header) -> Start new
         commitChunk();
         currentChunk.text = block.text;
         currentChunk.ids.push(block.id);
      }
    }
  }

  commitChunk(); // Final commit

  return chunks;
}

if (typeof module !== 'undefined') {
  module.exports = { parseHtmlToChunks };
}

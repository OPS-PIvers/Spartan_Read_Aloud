Here’s a breakdown of the issues and how to fix them for a robust mapping.

-----

### The Core Problems

1.  **Inconsistent Search Terms**: There's a mismatch in how the search text is generated between your backend and frontend code.

      * In **`Code.js`**, the `searchWords` property is created from the first **6 words** of the text chunk.
      * In **`index.html`**, the `getCleanedChunkText` function generates a search term from the first **8 words** and appends "...".
        This inconsistency is a primary source of mapping failures.

2.  **OCR vs. PDF.js Discrepancies**: The text extracted by the Google Drive OCR in `extractTextFromPdf` can have minor differences (e.g., spacing, line breaks, special characters) from the text layer rendered by PDF.js in the browser. This makes exact string matching fragile.

3.  **Fragile Search Logic**: The `searchTextInPDF` function in `index.html` has multiple fallbacks, which indicates that the primary matching methods are not consistently working. Relying on a short, non-unique snippet of text is inherently unreliable.

-----

### The Solution: A More Robust Mapping Strategy

To fix this, you need to standardize the text and use the full text chunk as the identifier. This will create a much more reliable mapping between the audio and the text.

Here are the steps to implement this solution:

#### 1\. Use the Full Text for Searching

Instead of a snippet, use the entire text of the chunk for searching. This dramatically increases the uniqueness of the search term.

  * **In `Code.js`**, you can remove the `searchWords` property from the JSON payload. The frontend will use the full `text` property of the chunk object. Your `audioDataForSheet` creation would look like this:

    ```javascript
    // In step2_GenerateMissingAudioAndFinalize() in Code.js
    audioDataForSheet.push({
      text: chunkText,
      audioUrl: `https://drive.google.com/uc?id=${audioFile.getId()}&export=media`,
      audioFilename: audioFile.getName()
      // The 'searchWords' property is no longer needed
    });
    ```

#### 2\. Normalize Text on Both Frontend and Backend (for future robustness)

While we are using the full text, small discrepancies can still exist. To make the matching even more robust, you can normalize the text by converting it to a consistent format.

Here is a normalization function you can use in your frontend JavaScript:

```javascript
// In index.html
function normalizeText(text) {
  return text.trim().toLowerCase().replace(/\s+/g, ' ').replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,"");
}
```

#### 3\. Update the Frontend JavaScript

Now, update your `index.html` to use this new strategy.

  * In the `playAudio` function, call `highlightChunkInPDF` with the full, original text from the server.

    ```javascript
    // In playAudio() in index.html
    function playAudio(chunk, index) {
      // ... (existing code) ...

      // Use the full text for highlighting
      highlightChunkInPDF(chunk.text); 

      // ... (rest of the function) ...
    }
    ```

  * Modify `highlightChunkInPDF` and `searchTextInPDF` to use the normalized text for comparison. This will make the matching much more reliable.

    ```javascript
    // In searchTextInPDF() in index.html
    function searchTextInPDF(searchText) {
      // Normalize the search text from the server chunk
      const searchPhrase = normalizeText(searchText);
      console.log('=== SEARCH DEBUG ===');
      console.log('Searching for normalized phrase:', searchPhrase);

      const textLayers = document.querySelectorAll('.textLayer');
      
      for (let layer of textLayers) {
        const spans = layer.querySelectorAll('span');
        let combinedText = '';
        let consecutiveSpans = [];

        for (let i = 0; i < spans.length; i++) {
          // Logic to combine and normalize text from spans
          // and compare with searchPhrase
        }
      }
      // ...
    }
    ```

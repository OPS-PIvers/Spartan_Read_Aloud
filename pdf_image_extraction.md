When you use Google Drive's built-in conversion feature to turn a PDF into a Google Doc (which is what your convertPdfToHtml function does), it performs two actions simultaneously:

It runs Optical Character Recognition (OCR) on any text in the PDF.

It extracts the images from the PDF and places them inline with the text in the new Google Doc, attempting to preserve the original layout.

Your convertPdfToHtml function in Code.js already uses this process. When it exports that temporary Google Doc to HTML, the final HTML file contains both the extracted text and the embedded images.

Yes, absolutely. Your code is already perfectly set up to do this, but you are currently *bypassing* this functionality for PDF files.

The `convertPdfToHtml` function in your `Code.js` file (line 1341) already does exactly what you're asking for:

1.  It converts the PDF to a temporary Google Doc using OCR (`Drive.Files.create(metadata, blob, {ocrLanguage: 'en'})`).
2.  This conversion process *preserves images* and places them alongside the extracted text.
3.  It then calls `exportDocToHtml` (line 1357) to export that new Google Doc (with images) into a single HTML file.

The *only* reason you are still using PDF.js is because your `getAssessmentPdf` function (line 1970) explicitly checks if the file is a PDF and, if so, sends the raw PDF data to the frontend instead of converting it to HTML.

To fix this and use the HTML-with-images conversion for *all* file types (including PDFs), you just need to remove that check.

### How to Implement the Change

In `Code.js`, modify the `getAssessmentPdf` function:

**1. Find this block of code (around line 2008):**

```javascript
          // PDFs: Return base64 data for PDF.js rendering
          if (mimeType === CONSTANTS.SUPPORTED_MIME_TYPES.PDF) {
            Logger.log('→ Serving PDF with base64 encoding');
            return {
              fileType: 'pdf',
              pdfData: Utilities.base64Encode(file.getBlob().getBytes()),
              fileName: fileName,
              audioChunks: audioChunks,
              sessionToken: sessionToken
            };
          }

          // Docs/Word: Convert to HTML and return sanitized HTML
          Logger.log('→ Converting to HTML for native rendering');
```

**2. Delete the entire `if` block.** Your modified function should look like this:

```javascript
// ... (start of getAssessmentPdf function) ...

          const fileId = getFileIdFromUrl(pdfUrl);
          if (!fileId) return { error: 'Invalid Google Drive URL in sheet.' };

          const file = DriveApp.getFileById(fileId);
          const mimeType = file.getMimeType();
          const fileName = file.getName();
          const audioChunks = JSON.parse(audioDataJson);

          Logger.log(`Serving assessment: ${fileName} (${mimeType}) to ${email}`);

          // Generate session token for secure audio access
          const sessionToken = generateSessionToken(cleanEmail, pdfUrl);

          // ALWAYS Convert to HTML and return sanitized HTML
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
// ... (rest of function) ...
```

By removing this check, PDFs will no longer be treated as a special case. They will be processed by the `convertFileToHtml` function just like Google Docs and Word files, resulting in an HTML file with the images included, which is exactly what you wanted. Your `student.html` file will now *always* receive `fileType: 'html'`.
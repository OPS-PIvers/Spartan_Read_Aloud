# Spartan Assessment Portal - System Audit Report

This report outlines identified bugs, performance optimization opportunities, and recommended UI enhancements for the Spartan Assessment Portal Google Apps Script project.

## 1. Bugs & Technical Issues

| Category | Issue Description | Impact |
| :--- | :--- | :--- |
| **Logic** | **Chunk Count Mismatch Loop:** In `step2_GenerateMissingAudioAndFinalize`, processing skips if `textChunks.length !== totalChunks`. If regex logic or parsing changes, assessments can get "stuck" in a perpetual processing loop. | Assessments remain in "Processing" state indefinitely. |
| **API** | **Gemini TTS Payload Structure:** `generateAudioFromTextChunk` uses a fixed payload. Gemini API updates to `responseModalities` or `speechConfig` requirements may break this hardcoded structure. | Audio generation failure. |
| **Storage** | **Google Sheets Cell Limit:** `AUDIO_JSON` is stored in a single cell. Google Sheets has a 50,000 character limit per cell. Long assessments with many chunks will exceed this limit. | Data truncation and broken student views. |
| **Frontend** | **Browser Autoplay Policies:** The `globalAudioPlayer` may be blocked by browsers if the student attempts to play audio before interacting with the page. | Audio fails to play on first click. |
| **Backend** | **OCR Execution Time:** PDF OCR via `convertPdfToHtml` is resource-heavy. Large batches or large files risk hitting the 6-minute Google Apps Script execution limit. | Partial processing and incomplete data. |

## 2. Low-Burden, High-Reward Performance Optimizations

### A. Batch Spreadsheet Operations
*   **Current:** `step1` and `step2` perform `setValue()` inside loops.
*   **Optimization:** Read the entire range into an array using `getValues()`, modify the array in memory, and write back once using `setValues()`. This can reduce processing time by 80%+.

### B. Parallel Audio Fetching
*   **Current:** Audio chunks are often fetched sequentially.
*   **Optimization:** Use `UrlFetchApp.fetchAll()` for internal API calls and fetching Drive blobs to reduce network latency during generation and student preloading.

### C. Text Extraction Memoization
*   **Current:** `extractTextFromFile` may re-run OCR or conversion multiple times for the same file.
*   **Optimization:** Implement a caching layer (using `CacheService` or a hidden sheet) to store the extracted HTML/JSON mapping keyed by the file's MD5 checksum or `lastUpdated` timestamp.

### D. Asset Lazy Loading
*   **Current:** Images are embedded as massive base64 strings in the HTML payload.
*   **Optimization:** Use temporary signed Drive URLs or implement a lazy-loading mechanism where base64 data is fetched only when the image enters the viewport.

## 3. UI & UX Enhancements

### Student Experience
*   **Focus Mode "Auto-Scroll":** Automatically scroll the active chunk to the top 1/3 of the screen (instead of center) to provide better reading context for subsequent text.
*   **Playback Presets:** Add "Slow" (0.8x), "Normal" (1.0x), and "Fast" (1.2x) buttons to the toolbar for faster accessibility than a slider.
*   **Keyboard Controls:** Map `Space` to Play/Pause and `Arrow keys` to navigation.

### Teacher Dashboard
*   **Live Processing Logs:** Add a "View Logs" button that displays the most recent `Logger.log` entries, allowing teachers to diagnose why an assessment is "Processing."
*   **Bulk Actions:** Enable checkboxes in the database table to allow bulk setting of passwords, expiry dates, or student lists.
*   **Student Engagement Tracking:** If submissions are enabled, show a "Last Seen" timestamp or progress percentage for students assigned to an assessment.

### Accessibility
*   **High Contrast Mode:** A toggle for a simplified, high-contrast UI for students with visual impairments.
*   **Dyslexic-Friendly Fonts:** Add an option to switch the assessment text to fonts like OpenDyslexic.


const pdfContainer = document.getElementById('pdf-container');
const assessmentSelectionContainer = document.getElementById('assessment-selection-container');
const viewerContainer = document.getElementById('viewer-container');

let globalAudioPlayer = new Audio();
let currentlyPlayingChunk = null;
let serverChunks = [];
let currentChunkIndex = 0;
let isLooping = false;
let isFocusMode = false;
let currentPlaybackRate = 1.0;
let currentHighlight = null;
let pageTextContent = []; // Store extracted text content for each page
let isProgrammaticScroll = false; // Flag to ignore auto-scroll during highlighting
let audioCache = new Map(); // Cache loaded audio chunks (key: fileId, value: base64 data)
let isLoadingAudio = false; // Track loading state
let sessionToken = null; // Secure session token for audio access
let totalChunks = 0; // Total number of audio chunks
let chunksLoaded = 0; // Number of chunks currently loaded/cached

// Session storage for multi-assessment flow
let storedEmail = null;
let storedPassword = null;
let availableAssessments = [];

// Loading state for assessment preloading
let currentLoadingCard = null; // Reference to the card currently being loaded
let currentLoadingOverlay = null; // Reference to the loading overlay element
let pendingAssessmentData = null; // Stores PDF data while preloading audio

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.11.338/pdf.worker.min.js';

function initializeStudentView(studentData) {
  storedEmail = studentData.email;
  storedPassword = studentData.password;
  onAssessmentsLoaded({ assessments: studentData.assessments });
}


function onAssessmentsLoaded(result) {
  console.log('onAssessmentsLoaded called with result:', result);

  if (result.error) {
    onAssessmentLoadError(result.error);
    return;
  }

  availableAssessments = result.assessments;
  console.log(`Loaded ${availableAssessments.length} assessment(s)`);
  console.log('Assessment data:', availableAssessments);

  // Always show assessment selection screen for consistent UX
  if (availableAssessments.length > 0) {
    console.log('Showing assessment selection screen...');
    showAssessmentSelection();
  } else {
    console.error('No assessments in array - this should not happen');
    onAssessmentLoadError('No assessments found in response');
  }
}

/**
 * Creates and displays a loading overlay on an assessment card
 * @param {HTMLElement} cardElement - The assessment card to show loading on
 * @returns {Object} - Object containing overlay element and update function
 */
function showLoadingOverlay(cardElement) {
  // Create overlay container
  const overlay = document.createElement('div');
  overlay.className = 'assessment-card-loading-overlay';

  // Create spinner
  const spinner = document.createElement('div');
  spinner.className = 'skeleton-loader';

  // Create progress meter container
  const meterContainer = document.createElement('div');
  meterContainer.className = 'progress-meter-container';

  // Create progress meter fill
  const meterFill = document.createElement('div');
  meterFill.className = 'progress-meter-fill';
  meterContainer.appendChild(meterFill);

  // Append elements
  overlay.appendChild(spinner);
  overlay.appendChild(meterContainer);
  cardElement.appendChild(overlay);

  return {
    overlay: overlay,
    spinner: spinner,
    meterFill: meterFill,
    updateProgress: (percent) => {
      meterFill.style.width = `${Math.min(100, Math.max(0, percent))}%`;
    },
    showBeginButton: (onClickCallback) => {
      // Remove spinner and meter
      spinner.remove();
      meterContainer.remove();

      // Create Begin Assessment button
      const beginBtn = document.createElement('button');
      beginBtn.className = 'begin-assessment-btn';
      beginBtn.textContent = 'Begin Assessment';
      beginBtn.onclick = (e) => {
        e.stopPropagation(); // Prevent card click event
        onClickCallback();
      };

      overlay.appendChild(beginBtn);
    },
    remove: () => {
      overlay.remove();
    }
  };
}

function showAssessmentSelection() {
  assessmentSelectionContainer.style.display = 'block';

  const assessmentGrid = document.getElementById('assessment-grid');
  assessmentGrid.innerHTML = '';

  availableAssessments.forEach((assessment, index) => {
    const card = document.createElement('div');
    card.className = 'assessment-card';
    card.onclick = () => loadSpecificAssessment(assessment.assessmentUrl, card);

    let metaHtml = '';
    if (assessment.instructor) {
      metaHtml += `<p class="assessment-meta-item"><span class="assessment-meta-label">Teacher:</span> ${escapeHtml(assessment.instructor)}</p>`;
    }
    if (assessment.className) {
      metaHtml += `<p class="assessment-meta-item"><span class="assessment-meta-label">Class:</span> ${escapeHtml(assessment.className)}</p>`;
    }

    card.innerHTML = `
      <h3 class="assessment-title">${escapeHtml(assessment.assessmentName)}</h3>
      <div class="assessment-meta">
        ${metaHtml}
        <span class="assessment-status-badge">Ready</span>
      </div>
    `;

    assessmentGrid.appendChild(card);
  });
}

function loadSpecificAssessment(assessmentUrl, cardElement) {
  console.log(`Loading specific assessment: ${assessmentUrl}`);

  // Store reference to card and show loading overlay
  currentLoadingCard = cardElement;
  currentLoadingOverlay = showLoadingOverlay(cardElement);

  // Disable other cards while loading
  const allCards = document.querySelectorAll('.assessment-card');
  allCards.forEach(card => {
    if (card !== cardElement) {
      card.style.pointerEvents = 'none';
      card.style.opacity = '0.5';
    }
  });

  // Fetch PDF data but don't show viewer yet
  google.script.run
    .withSuccessHandler(onPdfLoadedForPreload)
    .withFailureHandler(onPdfLoadError)
    .getAssessmentPdf(storedEmail, storedPassword, assessmentUrl);
}

function backToLogin() {
  google.script.run.withSuccessHandler(function(html) {
    document.getElementById('app-container').innerHTML = html;
  }).getLoginView();
}

function onAssessmentLoadError(error) {
    console.error('Assessment load error:', error);
    // Show error inline in the viewer container (using textContent to prevent XSS)
    viewerContainer.style.display = 'block';
    assessmentSelectionContainer.style.display = 'none';
    viewerContainer.innerHTML = '';
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.style.color = 'red';
    errorDiv.style.fontWeight = 'bold';
    errorDiv.style.margin = '2em';
    errorDiv.textContent = 'Error loading PDF: ' + String(error);
    viewerContainer.appendChild(errorDiv);
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

/**
 * Handles PDF load for preload flow - stores data and starts audio preloading
 * without displaying the viewer yet.
 */
function onPdfLoadedForPreload(result) {
  if (result.error) {
    onPdfLoadError(result.error);
    return;
  }

  console.log('PDF data received, starting audio preload...');

  // Store data for later rendering
  pendingAssessmentData = result;
  serverChunks = result.audioChunks;
  totalChunks = serverChunks.length;
  sessionToken = result.sessionToken;
  chunksLoaded = 0; // Reset counter
  audioCache.clear(); // Clear any old cache

  console.log(`Starting preload for ${totalChunks} audio chunks`);

  // Start preloading with progress callback
  bulkPreloadAudioWithProgress((progress) => {
    // Update progress meter
    if (currentLoadingOverlay) {
      currentLoadingOverlay.updateProgress(progress);
    }
  }, () => {
    // On complete: show Begin Assessment button
    console.log('✓ All audio preloaded!');
    if (currentLoadingOverlay) {
      currentLoadingOverlay.showBeginButton(() => {
        // When Begin Assessment is clicked
        beginAssessment();
      });
    }
  });
}

/**
 * Called when user clicks "Begin Assessment" button after preloading completes
 */
function beginAssessment() {
  console.log('Beginning assessment with preloaded data...');

  // Hide assessment selection, show viewer
  assessmentSelectionContainer.style.display = 'none';
  viewerContainer.style.display = 'block';

  // Remove loading overlay
  if (currentLoadingOverlay) {
    currentLoadingOverlay.remove();
    currentLoadingOverlay = null;
  }

  // Now render the PDF/HTML using stored data
  if (pendingAssessmentData) {
    onPdfLoaded(pendingAssessmentData);
    pendingAssessmentData = null; // Clear stored data
  }
}

function onPdfLoaded(result) {
  if (result.error) {
    onPdfLoadError(result.error);
    return;
  }

  serverChunks = result.audioChunks;
  totalChunks = serverChunks.length;
  sessionToken = result.sessionToken; // Store session token for audio fetching
  console.log(`Loaded ${serverChunks.length} audio chunks`);
  console.log(`Session token received: ${sessionToken ? 'Yes' : 'No'}`);

  // Dual rendering based on fileType from backend
  if (result.fileType === 'pdf') {
    console.log('Rendering PDF using PDF.js');
    // Existing PDF.js rendering (BACKWARDS COMPATIBLE)
    const pdfData = atob(result.pdfData);
    const loadingTask = pdfjsLib.getDocument({ data: pdfData });

    loadingTask.promise.then(pdfDocument => {
      pdfContainer.innerHTML = '';
      const pagePromises = [];
      for (let pageNum = 1; pageNum <= pdfDocument.numPages; pageNum++) {
        pagePromises.push(processPage(pdfDocument, pageNum));
      }

      Promise.all(pagePromises).then(() => {
        console.log("✓ PDF rendered successfully");
        initializeAudioToolbar();
        setupEventListeners();
        // Only preload audio if not already loaded (backward compatibility for direct loads)
        if (chunksLoaded < totalChunks) {
          console.log('Starting background audio preload...');
          bulkPreloadAudio();
        } else {
          console.log('✓ Audio already preloaded, ready to play!');
        }
      });
    }).catch(onPdfLoadError);

  } else if (result.fileType === 'html') {
    console.log('Rendering HTML assessment');
    // New HTML rendering for Docs/Word
    renderHtmlAssessment(result.assessmentHtml);
    initializeAudioToolbar();
    setupEventListeners();
    // Only preload audio if not already loaded (backward compatibility for direct loads)
    if (chunksLoaded < totalChunks) {
      console.log('Starting background audio preload...');
      bulkPreloadAudio();
    } else {
      console.log('✓ Audio already preloaded, ready to play!');
    }

  } else {
    onPdfLoadError(`Unknown file type received: ${result.fileType}`);
  }
}

function processPage(pdfDocument, pageNum) {
  return pdfDocument.getPage(pageNum).then(page => {
    // Use scale 1.5 for high-quality rendering with good layout balance
    const viewport = page.getViewport({ scale: 1.5 });

    // Create page container
    const pageContainer = document.createElement('div');
    pageContainer.className = 'page-container';
    pageContainer.id = 'page-container-' + pageNum;

    // Create canvas
    const canvas = document.createElement('canvas');
    canvas.id = 'page-' + pageNum;
    const context = canvas.getContext('2d');
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    // Create text layer div
    const textLayerDiv = document.createElement('div');
    textLayerDiv.className = 'textLayer';
    textLayerDiv.style.width = viewport.width + 'px';
    textLayerDiv.style.height = viewport.height + 'px';

    // Append canvas and text layer to page container
    pageContainer.appendChild(canvas);
    pageContainer.appendChild(textLayerDiv);
    pdfContainer.appendChild(pageContainer);

    const renderContext = {
      canvasContext: context,
      viewport: viewport
    };

    // Render canvas and text layer
    const renderPromise = page.render(renderContext).promise;

    const textPromise = page.getTextContent().then(textContent => {
      const pageText = textContent.items.map(item => item.str).join(' ');
      pageTextContent[pageNum - 1] = pageText; // Store 0-indexed

      // Render text layer
      return pdfjsLib.renderTextLayer({
        textContent: textContent,
        container: textLayerDiv,
        viewport: viewport,
        textDivs: []
      }).promise;
    }).catch(err => {
      console.warn(`Failed to extract/render text for page ${pageNum}:`, err);
      pageTextContent[pageNum - 1] = ''; // Empty string fallback
    });

    return Promise.all([renderPromise, textPromise]);
  });
}

/**
 * Renders an HTML-based assessment (from Google Docs or Word).
 * Replaces PDF.js rendering for non-PDF formats.
 * @param {string} htmlContent Sanitized HTML from backend
 */
function renderHtmlAssessment(htmlContent) {
  console.log("Rendering HTML assessment");

  // Create container with CSS reset
  const htmlContainer = document.createElement('div');
  htmlContainer.className = 'html-assessment-container';
  htmlContainer.innerHTML = htmlContent;

  pdfContainer.innerHTML = '';
  pdfContainer.appendChild(htmlContainer);

  // Phase 2: Pattern detection and enhancement
  detectAndEnhanceAssessmentPatterns(htmlContainer);
  groupRelatedElements(htmlContainer);

  // Extract text content structure for highlighting
  extractTextContentFromHtml(htmlContainer);

  console.log("✓ HTML assessment rendered with pattern enhancement");
}

/**
 * Extracts text content from HTML elements for chunk matching.
 * Stores element references and normalized text for highlighting.
 * @param {HTMLElement} container The HTML assessment container
 */
function extractTextContentFromHtml(container) {
  pageTextContent = []; // Reset global array

  // Select all text-containing elements
  const elements = container.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, div');

  elements.forEach((element, index) => {
    const text = element.textContent.trim();
    if (text) {
      pageTextContent.push({
        element: element,
        text: text,
        normalizedText: normalizeText(text)
      });
    }
  });

  console.log(`✓ Extracted ${pageTextContent.length} text elements from HTML`);
}

/**
 * Detects assessment patterns and adds CSS classes for enhanced styling.
 * Patterns detected:
 * - Questions: paragraphs starting with "1.", "2.", etc.
 * - Answer choices: text starting with "A.", "B.", "C.", "D."
 * - Passages: longer paragraphs without numbering
 * @param {HTMLElement} container The HTML assessment container
 */
function detectAndEnhanceAssessmentPatterns(container) {
  console.log('=== PATTERN DETECTION ===');

  // Patterns to detect
  const questionPattern = /^\s*(\d+)\.\s+/; // Matches "1. ", "2. ", etc.
  const answerPattern = /^\s*([A-D])\.\s+/i; // Matches "A. ", "B. ", "C. ", "D. "
  const passageLengthThreshold = 150; // Characters

  let questionsDetected = 0;
  let answersDetected = 0;
  let passagesDetected = 0;

  // Get all potential elements (paragraphs, list items, table cells)
  const elements = container.querySelectorAll('p, li, td, div');

  elements.forEach((element) => {
    const text = element.textContent.trim();

    // Skip empty elements
    if (!text) return;

    // Skip if already classified or has children with patterns
    if (element.classList.contains('detected-question') ||
        element.classList.contains('detected-answer-choice') ||
        element.classList.contains('detected-passage')) {
      return;
    }

    // Detect questions
    if (questionPattern.test(text)) {
      element.classList.add('detected-question');
      questionsDetected++;
      console.log(`→ Detected question: ${text.substring(0, 50)}...`);
    }
    // Detect answer choices
    else if (answerPattern.test(text)) {
      element.classList.add('detected-answer-choice');
      answersDetected++;
    }
    // Detect passages (longer text without numbering)
    else if (text.length > passageLengthThreshold &&
             !questionPattern.test(text) &&
             !answerPattern.test(text) &&
             element.tagName.toLowerCase() === 'p') {
      element.classList.add('detected-passage');
      passagesDetected++;
      console.log(`→ Detected passage: ${text.substring(0, 50)}...`);
    }
  });

  console.log(`✓ Pattern detection complete:`);
  console.log(`  - Questions: ${questionsDetected}`);
  console.log(`  - Answer choices: ${answersDetected}`);
  console.log(`  - Passages: ${passagesDetected}`);

  return {
    questions: questionsDetected,
    answers: answersDetected,
    passages: passagesDetected
  };
}

/**
 * Improves text chunking by grouping related elements.
 * Groups question + answer choices together for better highlighting.
 * @param {HTMLElement} container The HTML assessment container
 */
function groupRelatedElements(container) {
  const questions = container.querySelectorAll('.detected-question');

  questions.forEach((question) => {
    // Find all answer choices immediately following this question
    let nextElement = question.nextElementSibling;
    const relatedAnswers = [];

    while (nextElement &&
           (nextElement.classList.contains('detected-answer-choice') ||
            nextElement.tagName === 'BR' ||
            (nextElement.textContent.trim() === ''))) {

      if (nextElement.classList.contains('detected-answer-choice')) {
        relatedAnswers.push(nextElement);
      }

      nextElement = nextElement.nextElementSibling;

      // Stop after finding 4-5 answers (typical multiple choice)
      if (relatedAnswers.length >= 5) break;
    }

    // Store relationship for improved highlighting
    if (relatedAnswers.length > 0) {
      question.setAttribute('data-answer-count', relatedAnswers.length);
      relatedAnswers.forEach((answer, index) => {
        answer.setAttribute('data-question-ref', questions.length);
        answer.setAttribute('data-answer-index', index);
      });
    }
  });

  console.log(`✓ Grouped ${questions.length} questions with their answer choices`);
}

function onPdfLoadError(error) {
  console.error('PDF Load Error:', error);

  // Clean up loading overlay if present
  if (currentLoadingOverlay) {
    currentLoadingOverlay.remove();
    currentLoadingOverlay = null;
  }

  // Re-enable all cards
  const allCards = document.querySelectorAll('.assessment-card');
  allCards.forEach(card => {
    card.style.pointerEvents = '';
    card.style.opacity = '';
  });

  // Reset state
  currentLoadingCard = null;
  pendingAssessmentData = null;

    // Instead of manipulating login DOM (not present in student view), trigger error UI
  if (typeof onAssessmentLoadError === 'function') {
    onAssessmentLoadError(error);
  }
}


function getCleanedChunkText(text) {
  const words = text.trim().split(/\s+/);
  return words.slice(0, 8).join(' ') + (words.length > 8 ? '...' : '');
}

/**
 * Extracts file ID from Google Drive URL.
 * @param {string} audioUrl Drive URL containing file ID
 * @returns {string|null} File ID or null if not found
 */
function getFileIdFromAudioUrl(audioUrl) {
  const fileIdMatch = audioUrl.match(/id=([a-zA-Z0-9_-]+)/);
  return fileIdMatch ? fileIdMatch[1] : null;
}

/**
 * Enhanced bulk preload with progress tracking.
 * @param {Function} onProgressCallback - Called with progress percentage (0-100)
 * @param {Function} onCompleteCallback - Called when all audio is loaded
 */
function bulkPreloadAudioWithProgress(onProgressCallback, onCompleteCallback) {
  if (!sessionToken || serverChunks.length === 0) {
    console.warn('Cannot preload audio: missing session token or chunks');
    onCompleteCallback && onCompleteCallback();
    return;
  }

  console.log('=== BULK AUDIO PRELOAD WITH PROGRESS ===');

  // Get all file IDs
  const allFileIds = serverChunks.map(c => getFileIdFromAudioUrl(c.audioUrl)).filter(id => id);

  if (allFileIds.length === 0) {
    console.warn('No valid file IDs found');
    onCompleteCallback && onCompleteCallback();
    return;
  }

  // Report initial progress
  onProgressCallback && onProgressCallback(0);

  // Load all chunks in batches of 5
  loadInBatchesWithProgress(allFileIds, 5, onProgressCallback, onCompleteCallback);
}

/**
 * Loads file IDs in batches with progress reporting.
 * @param {Array<string>} fileIds - Array of file IDs
 * @param {number} batchSize - Number of files per batch
 * @param {Function} onProgressCallback - Progress callback
 * @param {Function} onCompleteCallback - Completion callback
 */
function loadInBatchesWithProgress(fileIds, batchSize, onProgressCallback, onCompleteCallback) {
  if (fileIds.length === 0) {
    console.log('✓ All chunks loaded with progress tracking');
    onProgressCallback && onProgressCallback(100);
    onCompleteCallback && onCompleteCallback();
    return;
  }

  const batch = fileIds.slice(0, batchSize);
  const remaining = fileIds.slice(batchSize);

  bulkFetchAndCache(batch)
    .then(() => {
      // Calculate and report progress
      const progress = (chunksLoaded / totalChunks) * 100;
      console.log(`→ Progress: ${chunksLoaded}/${totalChunks} (${Math.round(progress)}%)`);
      onProgressCallback && onProgressCallback(progress);

      // Continue with next batch
      setTimeout(() => {
        loadInBatchesWithProgress(remaining, batchSize, onProgressCallback, onCompleteCallback);
      }, 200); // Small delay between batches
    })
    .catch(err => {
      console.error('Batch loading failed:', err);
      // Continue anyway
      setTimeout(() => {
        loadInBatchesWithProgress(remaining, batchSize, onProgressCallback, onCompleteCallback);
      }, 500);
    });
}

/**
 * Bulk preloads audio chunks in smart batches.
 * Strategy: Load first 5 immediately, then rest in background.
 */
function bulkPreloadAudio() {
  if (!sessionToken || serverChunks.length === 0) {
    console.warn('Cannot preload audio: missing session token or chunks');
    return;
  }

  console.log('=== BULK AUDIO PRELOAD START ===');
  updateLoadingProgress(); // Show initial progress

  // Phase 1: Load first 5 chunks immediately (priority for playback)
  const priorityCount = Math.min(5, serverChunks.length);
  const priorityChunks = serverChunks.slice(0, priorityCount);
  const priorityFileIds = priorityChunks.map(c => getFileIdFromAudioUrl(c.audioUrl)).filter(id => id);

  console.log(`→ Phase 1: Loading first ${priorityFileIds.length} priority chunks`);

  bulkFetchAndCache(priorityFileIds)
    .then(() => {
      console.log(`✓ Phase 1 complete: ${chunksLoaded}/${totalChunks} chunks ready`);
      updateLoadingProgress();

      // Phase 2: Load remaining chunks in background (if any)
      if (serverChunks.length > priorityCount) {
        const remainingChunks = serverChunks.slice(priorityCount);
        const remainingFileIds = remainingChunks.map(c => getFileIdFromAudioUrl(c.audioUrl)).filter(id => id);

        console.log(`→ Phase 2: Background loading ${remainingFileIds.length} remaining chunks`);

        // Load in batches of 5 to avoid timeouts
        loadInBatches(remainingFileIds, 5);
      }
    })
    .catch(err => {
      console.error('Priority chunk loading failed:', err);
    });
}

/**
 * Loads file IDs in batches to avoid timeouts.
 * @param {Array<string>} fileIds Array of file IDs
 * @param {number} batchSize Number of files per batch
 */
function loadInBatches(fileIds, batchSize) {
  if (fileIds.length === 0) {
    console.log('✓ All chunks loaded');
    updateLoadingProgress();
    return;
  }

  const batch = fileIds.slice(0, batchSize);
  const remaining = fileIds.slice(batchSize);

  bulkFetchAndCache(batch)
    .then(() => {
      console.log(`→ Batch complete: ${chunksLoaded}/${totalChunks} chunks ready`);
      updateLoadingProgress();

      // Continue with next batch after short delay
      setTimeout(() => loadInBatches(remaining, batchSize), 500);
    })
    .catch(err => {
      console.error('Batch loading failed:', err);
      // Continue anyway
      setTimeout(() => loadInBatches(remaining, batchSize), 1000);
    });
}

/**
 * Fetches multiple audio files via getBulkAudioData and caches results.
 * @param {Array<string>} fileIds Array of file IDs to fetch
 * @returns {Promise} Resolves when all files are cached
 */
function bulkFetchAndCache(fileIds) {
  if (!fileIds || fileIds.length === 0) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler(result => {
        if (!result.success) {
          console.error('Bulk fetch failed:', result.error);
          reject(result.error);
          return;
        }

        // Cache all successful results
        result.results.forEach(item => {
          if (item.success && item.data) {
            audioCache.set(item.fileId, item.data);
            chunksLoaded++;
          } else {
            console.warn(`Failed to load audio: ${item.fileId}`, item.error);
          }
        });

        console.log(`✓ Cached ${result.stats.success} files (${audioCache.size} total in cache)`);
        resolve();
      })
      .withFailureHandler(err => {
        console.error('Bulk fetch error:', err);
        reject(err);
      })
      .getBulkAudioData(sessionToken, fileIds);
  });
}

/**
 * Fetches and caches audio data for a single chunk (fallback).
 * Used when bulk loading is not available or fails.
 * @param {string} audioUrl The audio URL from the chunk
 * @returns {Promise<string>} Base64-encoded audio data
 */
function fetchAndCacheAudio(audioUrl) {
  const fileId = getFileIdFromAudioUrl(audioUrl);
  if (!fileId) {
    return Promise.reject('Could not parse audio file ID from URL');
  }

  // Check cache first
  if (audioCache.has(fileId)) {
    console.log('✓ Using cached audio for:', fileId);
    return Promise.resolve(audioCache.get(fileId));
  }

  console.log('→ Fetching single audio from server:', fileId);

  // Fallback to single fetch if not in cache
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler(base64Data => {
        if (!base64Data) {
          reject('Failed to load audio data from server.');
          return;
        }

        // Cache the audio data
        audioCache.set(fileId, base64Data);
        chunksLoaded++;
        console.log(`✓ Cached audio (${chunksLoaded}/${totalChunks})`);
        resolve(base64Data);
      })
      .withFailureHandler(err => {
        reject(err);
      })
          .getAudioDataAsBase64(sessionToken, fileId);
  });
}

/**
 * Updates the loading progress indicator in the UI.
 * Shows how many chunks are ready to play.
 */
function updateLoadingProgress() {
  const chunkCounter = document.getElementById('chunk-counter');
  if (!chunkCounter) return;

  if (chunksLoaded < totalChunks) {
    chunkCounter.textContent = `Loading audio: ${chunksLoaded}/${totalChunks} ready`;
    chunkCounter.style.color = '#f59e0b'; // Orange while loading
  } else {
    chunkCounter.textContent = `Chunk ${currentChunkIndex + 1} of ${totalChunks}`;
    chunkCounter.style.color = '#1a73e8'; // Blue when complete
  }
}

function normalizeText(text) {
  // Normalize text to handle OCR vs PDF.js differences
  return text.trim().toLowerCase()
    .replace(/\s+/g, ' ')  // Normalize whitespace
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, '');  // Remove punctuation
}

function initializeAudioToolbar() {
  document.getElementById('audio-toolbar').style.display = 'flex';
  populateTitleDropdown();
  createChunkMarkers();
  updateToolbarDisplay();

  // Initialize with first chunk ready to play
  if (serverChunks.length > 0) {
    currentChunkIndex = 0;
    updateToolbarDisplay();
  }
}

function createChunkMarkers() {
  const markersContainer = document.getElementById('chunk-markers');
  if (!markersContainer || serverChunks.length <= 1) return;

  markersContainer.innerHTML = ''; // Clear existing markers

  // Calculate total duration (estimate based on first chunk if needed)
  // For now, we'll space markers evenly as we don't know exact durations upfront
  const markerCount = serverChunks.length - 1; // markers between chunks

  for (let i = 1; i < serverChunks.length; i++) {
    const marker = document.createElement('div');
    marker.className = 'chunk-marker';
    // Position evenly across the progress bar
    marker.style.left = `${(i / serverChunks.length) * 100}%`;
    marker.title = `Chunk ${i + 1}: ${getCleanedChunkText(serverChunks[i].text)}`;
    markersContainer.appendChild(marker);
  }
}

function populateTitleDropdown() {
  const dropdown = document.getElementById('title-dropdown');
  dropdown.innerHTML = '';

  serverChunks.forEach((chunk, index) => {
    const item = document.createElement('div');
    item.className = 'dropdown-item';
    item.textContent = getCleanedChunkText(chunk.text);
    item.onclick = () => jumpToChunk(index);
    dropdown.appendChild(item);
  });
}

function updateToolbarDisplay() {
  if (serverChunks.length === 0) return;

  const currentChunk = serverChunks[currentChunkIndex];
  if (currentChunk) {
    const titleText = document.getElementById('current-title-text');
    if (titleText) {
      titleText.textContent = getCleanedChunkText(currentChunk.text);
    }
  }

  // Update chunk counter
  const chunkCounter = document.getElementById('chunk-counter');
  if (chunkCounter) {
    chunkCounter.textContent = `Chunk ${currentChunkIndex + 1} of ${serverChunks.length}`;
  }

  // Update speed preset buttons
  document.querySelectorAll('.speed-preset').forEach(btn => {
    const speed = parseFloat(btn.getAttribute('data-speed'));
    btn.classList.toggle('active', Math.abs(speed - currentPlaybackRate) < 0.01);
  });

  // Update speed slider
  const speedSlider = document.getElementById('speed-slider');
  if (speedSlider) {
    speedSlider.value = currentPlaybackRate;
  }

  // Update dropdown current item
  document.querySelectorAll('.dropdown-item').forEach((item, index) => {
    item.classList.toggle('current', index === currentChunkIndex);
  });

  // Update navigation button states
  document.getElementById('prev-btn').disabled = currentChunkIndex === 0;
  document.getElementById('next-btn').disabled = currentChunkIndex === serverChunks.length - 1;
}

function playAudio(chunk, index) {
  if (currentlyPlayingChunk === chunk) {
    if (globalAudioPlayer.paused) {
      globalAudioPlayer.play();
      updatePlayPauseButton(true);
    } else {
      globalAudioPlayer.pause();
      updatePlayPauseButton(false);
    }
    return;
  }

  globalAudioPlayer.pause();
  currentlyPlayingChunk = chunk;
  currentChunkIndex = index;

  // Show loading state
  isLoadingAudio = true;
  updatePlayPauseButton(true);
  setLoadingState(true);

  updateToolbarDisplay();

  // Use searchWords for current chunk and next chunk as boundaries
  // Fallback: generate searchWords from text for backward compatibility with old assessments
  const currentSearchWords = chunk.searchWords ||
    (chunk.text ? chunk.text.trim().split(/\s+/).slice(0, 8).join(' ') + '...' : '');
  const nextChunk = serverChunks[index + 1];
  const nextSearchWords = nextChunk ?
    (nextChunk.searchWords || (nextChunk.text ? nextChunk.text.trim().split(/\s+/).slice(0, 8).join(' ') + '...' : ''))
    : null;

  console.log('Audio title displayed:', getCleanedChunkText(chunk.text));
  console.log('Current chunk searchWords:', currentSearchWords);
  console.log('Next chunk searchWords:', nextSearchWords || '(last chunk)');

  highlightChunkInPDF(currentSearchWords, nextSearchWords);

  // Check if audio is already cached (from bulk preload)
  const fileId = getFileIdFromAudioUrl(chunk.audioUrl);
  if (!fileId) {
    onPdfLoadError('Invalid audio URL format');
    isLoadingAudio = false;
    setLoadingState(false);
    return;
  }

  let audioDataPromise;
  if (audioCache.has(fileId)) {
    console.log(`✓ Using preloaded audio for chunk ${index + 1}`);
    audioDataPromise = Promise.resolve(audioCache.get(fileId));
  } else {
    console.log(`→ Audio not preloaded, fetching chunk ${index + 1}`);
    audioDataPromise = fetchAndCacheAudio(chunk.audioUrl);
  }

  // Play the audio
  audioDataPromise
    .then(base64Data => {
      // Clear loading state
      isLoadingAudio = false;
      setLoadingState(false);

      // Set audio source and play
      globalAudioPlayer.src = `data:audio/wav;base64,${base64Data}`;
      globalAudioPlayer.playbackRate = currentPlaybackRate;

      return globalAudioPlayer.play();
    })
    .then(() => {
      // Audio started playing successfully
      console.log('✓ Audio playing');
      // Ensure play/pause button shows pause icon after loading completes
      updatePlayPauseButton(true);
    })
    .catch(err => {
      // Handle errors
      console.error('Audio playback error:', err);
      onPdfLoadError(err.toString());
      currentlyPlayingChunk = null;
      isLoadingAudio = false;
      setLoadingState(false);
      updatePlayPauseButton(false);
    });
}

function updatePlayPauseButton(isPlaying) {
  const btn = document.getElementById('play-pause-btn');
  const playIcon = btn.querySelector('.play-icon');
  const pauseIcon = btn.querySelector('.pause-icon');
  const loadingSpinner = btn.querySelector('.loading-spinner');

  // Hide loading spinner when updating play/pause state
  if (loadingSpinner) {
    loadingSpinner.style.display = 'none';
  }

  if (isPlaying) {
    playIcon.style.display = 'none';
    pauseIcon.style.display = 'block';
  } else {
    playIcon.style.display = 'block';
    pauseIcon.style.display = 'none';
  }
}

/**
 * Updates UI to show loading state during audio fetch.
 * @param {boolean} loading True to show loading state, false to clear it
 */
function setLoadingState(loading) {
  const playPauseBtn = document.getElementById('play-pause-btn');
  const playIcon = playPauseBtn.querySelector('.play-icon');
  const pauseIcon = playPauseBtn.querySelector('.pause-icon');
  const loadingSpinner = playPauseBtn.querySelector('.loading-spinner');

  if (loading) {
    // Hide play/pause icons, show spinner
    playIcon.style.display = 'none';
    pauseIcon.style.display = 'none';
    loadingSpinner.style.display = 'block';

    // Disable button while loading
    playPauseBtn.disabled = true;
    playPauseBtn.style.opacity = '0.8';
  } else {
    // Hide spinner, button state will be set by updatePlayPauseButton
    loadingSpinner.style.display = 'none';

    // Re-enable controls
    playPauseBtn.disabled = false;
    playPauseBtn.style.opacity = '1';
  }
}

function updateProgressBar() {
  if (globalAudioPlayer.duration) {
    const progress = (globalAudioPlayer.currentTime / globalAudioPlayer.duration) * 100;
    const progressFill = document.getElementById('progress-bar-fill');
    if (progressFill) {
      progressFill.style.width = progress + '%';
    }
  }
}

function setupEventListeners() {
  // Core controls
  document.getElementById('prev-btn').onclick = () => navigateChunk(-1);
  document.getElementById('next-btn').onclick = () => navigateChunk(1);
  document.getElementById('play-pause-btn').onclick = togglePlayPause;
  document.getElementById('skip-back-btn').onclick = () => skipTime(-10);
  document.getElementById('skip-forward-btn').onclick = () => skipTime(10);

  // Timeline scrubbing
  document.getElementById('timeline').oninput = scrubTimeline;

  // Speed preset buttons
  document.querySelectorAll('.speed-preset').forEach(btn => {
    btn.onclick = () => {
      const speed = parseFloat(btn.getAttribute('data-speed'));
      setPlaybackSpeed(speed);
    };
  });

  // Speed slider
  const speedSlider = document.getElementById('speed-slider');
  if (speedSlider) {
    speedSlider.oninput = (e) => {
      setPlaybackSpeed(parseFloat(e.target.value));
    };
  }

  // Title click to open chunk list
  const toolbarTitle = document.querySelector('.toolbar-title');
  if (toolbarTitle) {
    toolbarTitle.onclick = (e) => {
      e.stopPropagation();
      toggleChunkMenu();
    };
  }

  // Chunk list button (old compatibility)
  const oldTitleBtn = document.getElementById('current-title-btn');
  if (oldTitleBtn) {
    oldTitleBtn.onclick = toggleTitleDropdown;
  }

  // New menu buttons
  const speedMenuBtn = document.getElementById('speed-menu-btn');
  if (speedMenuBtn) {
    speedMenuBtn.onclick = (e) => {
      e.stopPropagation();
      toggleSpeedMenu();
    };
  }

  // Feature toggles
  const focusModeBtn = document.getElementById('focus-mode-btn');
  if (focusModeBtn) focusModeBtn.onclick = toggleFocusMode;

  const loopBtn = document.getElementById('loop-btn');
  if (loopBtn) loopBtn.onclick = toggleLoop;

  // Toggle secondary panel
  const toggleSecBtn = document.getElementById('toggle-secondary-btn');
  if (toggleSecBtn) toggleSecBtn.onclick = toggleSecondaryPanel;

  // Keyboard shortcuts
  document.addEventListener('keydown', handleKeyboard);

  // Audio event listeners
  globalAudioPlayer.ontimeupdate = () => {
    updateTimeline();
    updateProgressBar();
  };
  globalAudioPlayer.onended = handleAudioEnd;
  globalAudioPlayer.onloadedmetadata = updateDuration;

  // Close dropdowns when clicking outside
  document.addEventListener('click', (e) => {
    if (!e.target.closest('#speed-menu-btn') && !e.target.closest('#speed-popup')) {
      const speedPopup = document.getElementById('speed-popup');
      if (speedPopup) speedPopup.style.display = 'none';
    }
    if (!e.target.closest('.toolbar-title') && !e.target.closest('#chunk-popup')) {
      const chunkPopup = document.getElementById('chunk-popup');
      if (chunkPopup) chunkPopup.style.display = 'none';
    }
    // Old dropdown compatibility
    if (!e.target.closest('.chunk-nav')) {
      const titleDropdown = document.getElementById('title-dropdown');
      if (titleDropdown) titleDropdown.style.display = 'none';
    }
  });
}

let isSecondaryPanelExpanded = false;

function toggleSecondaryPanel() {
  const panel = document.getElementById('toolbar-secondary');
  const toggleBtn = document.getElementById('toggle-secondary-btn');
  const expandIcon = toggleBtn.querySelector('.expand-icon');
  const collapseIcon = toggleBtn.querySelector('.collapse-icon');

  isSecondaryPanelExpanded = !isSecondaryPanelExpanded;

  if (isSecondaryPanelExpanded) {
    panel.classList.add('expanded');
    expandIcon.style.display = 'none';
    collapseIcon.style.display = 'block';
    toggleBtn.setAttribute('aria-expanded', 'true');
    toggleBtn.title = 'Hide additional controls';
  } else {
    panel.classList.remove('expanded');
    expandIcon.style.display = 'block';
    collapseIcon.style.display = 'none';
    toggleBtn.setAttribute('aria-expanded', 'false');
    toggleBtn.title = 'Show more controls';
  }
}

function setPlaybackSpeed(speed) {
  currentPlaybackRate = Math.max(0.5, Math.min(2.0, speed));
  globalAudioPlayer.playbackRate = currentPlaybackRate;
  updateToolbarDisplay();

  // Update speed display
  const speedDisplay = document.getElementById('speed-display');
  if (speedDisplay) {
    speedDisplay.textContent = speed + 'x';
  }

  // Close speed popup
  const speedPopup = document.getElementById('speed-popup');
  if (speedPopup) {
    speedPopup.style.display = 'none';
  }
}

function toggleSpeedMenu() {
  const popup = document.getElementById('speed-popup');
  const chunkPopup = document.getElementById('chunk-popup');

  // Close chunk menu if open
  if (chunkPopup) chunkPopup.style.display = 'none';

  // Toggle speed menu
  if (popup) {
    popup.style.display = popup.style.display === 'none' ? 'block' : 'none';
  }
}

function toggleChunkMenu() {
  const popup = document.getElementById('chunk-popup');
  const speedPopup = document.getElementById('speed-popup');

  // Close speed menu if open
  if (speedPopup) speedPopup.style.display = 'none';

  // Toggle chunk menu
  if (popup) {
    popup.style.display = popup.style.display === 'none' ? 'block' : 'none';
  }
}

function navigateChunk(direction) {
  const newIndex = currentChunkIndex + direction;
  if (newIndex >= 0 && newIndex < serverChunks.length) {
    jumpToChunk(newIndex);
  }
}

function jumpToChunk(index) {
  currentChunkIndex = index;
  const chunk = serverChunks[index];
  playAudio(chunk, index);

  // Close all menus
  const titleDropdown = document.getElementById('title-dropdown');
  if (titleDropdown) titleDropdown.style.display = 'none';

  const chunkPopup = document.getElementById('chunk-popup');
  if (chunkPopup) chunkPopup.style.display = 'none';
}

function togglePlayPause() {
  if (globalAudioPlayer.paused) {
    if (!currentlyPlayingChunk) {
      jumpToChunk(currentChunkIndex);
    } else {
      globalAudioPlayer.play();
    }
    updatePlayPauseButton(true);
  } else {
    globalAudioPlayer.pause();
    updatePlayPauseButton(false);
  }
}

function skipTime(seconds) {
  // Only skip if audio is loaded and has duration
  if (!globalAudioPlayer.duration) {
    return;
  }

  // Calculate new time and clamp to valid range
  const newTime = Math.max(0, Math.min(
    globalAudioPlayer.currentTime + seconds,
    globalAudioPlayer.duration
  ));

  // Update audio position
  globalAudioPlayer.currentTime = newTime;

  // Update UI immediately
  updateTimeline();
  updateProgressBar();
}

function scrubTimeline() {
  const timeline = document.getElementById('timeline');
  const seekTime = (timeline.value / 100) * globalAudioPlayer.duration;
  globalAudioPlayer.currentTime = seekTime;
}

function updateTimeline() {
  if (globalAudioPlayer.duration) {
    const progress = (globalAudioPlayer.currentTime / globalAudioPlayer.duration) * 100;
    document.getElementById('timeline').value = progress;
    document.getElementById('current-time').textContent = formatTime(globalAudioPlayer.currentTime);
  }
}

function updateDuration() {
  document.getElementById('total-time').textContent = formatTime(globalAudioPlayer.duration);
}

function formatTime(seconds) {
  const mins = Math.floor(seconds / 60);
  const secs = Math.floor(seconds % 60);
  return `${mins}:${secs.toString().padStart(2, '0')}`;
}

function handleAudioEnd() {
  updatePlayPauseButton(false);

  if (isLooping) {
    globalAudioPlayer.currentTime = 0;
    globalAudioPlayer.play();
    updatePlayPauseButton(true);
  } else if (currentChunkIndex < serverChunks.length - 1) {
    // Auto-advance to next chunk
    navigateChunk(1);
  } else {
    currentlyPlayingChunk = null;
  }
}

function toggleLoop() {
  isLooping = !isLooping;
  const loopBtn = document.getElementById('loop-btn');
  loopBtn.classList.toggle('active', isLooping);
  loopBtn.setAttribute('aria-pressed', isLooping.toString());
}

function toggleTitleDropdown() {
  const dropdown = document.getElementById('title-dropdown');
  dropdown.style.display = dropdown.style.display === 'none' ? 'block' : 'none';
}

/**
 * Highlights text chunks in HTML-rendered assessments.
 * Uses element-based matching instead of PDF.js text layer spans.
 * @param {string} currentSearchWords First few words of current chunk
 * @param {string} nextSearchWords First few words of next chunk (boundary)
 * @returns {boolean} True if highlighting succeeded
 */
function highlightChunkInHTML(currentSearchWords, nextSearchWords) {
  // Remove previous highlights and has-highlight classes
  document.querySelectorAll('.highlighted-element').forEach(el => {
    el.classList.remove('highlighted-element');
  });
  document.querySelectorAll('.has-highlight').forEach(el => {
    el.classList.remove('has-highlight');
  });

  const normalizedSearch = normalizeText(currentSearchWords);
  console.log('=== HTML HIGHLIGHTING ===');
  console.log('Searching for:', currentSearchWords);
  console.log('Normalized:', normalizedSearch);

  // Find matching element in extracted text content
  for (let i = 0; i < pageTextContent.length; i++) {
    const item = pageTextContent[i];

    if (item.normalizedText.includes(normalizedSearch)) {
      console.log(`✓ Found match at element ${i}`);

      // Highlight current element
      item.element.classList.add('highlighted-element');

      // Add has-highlight class to all parent elements up to html-assessment-container
      let parent = item.element.parentElement;
      while (parent && !parent.classList.contains('html-assessment-container')) {
        parent.classList.add('has-highlight');
        parent = parent.parentElement;
      }

      // If there's a next chunk, highlight all elements up to it
      if (nextSearchWords) {
        const normalizedNext = normalizeText(nextSearchWords);
        console.log('Highlighting range to next chunk:', nextSearchWords);

        for (let j = i + 1; j < pageTextContent.length; j++) {
          if (pageTextContent[j].normalizedText.includes(normalizedNext)) {
            console.log(`✓ Found end boundary at element ${j}`);
            break; // Stop before next chunk
          }
          // Highlight all elements between current and next
          const elem = pageTextContent[j].element;
          elem.classList.add('highlighted-element');

          // Add has-highlight to parents
          let p = elem.parentElement;
          while (p && !p.classList.contains('html-assessment-container')) {
            p.classList.add('has-highlight');
            p = p.parentElement;
          }
        }
      }

      // Scroll to first highlighted element (with programmatic scroll flag)
      isProgrammaticScroll = true;
      const rect = item.element.getBoundingClientRect();
      const offset = window.innerHeight * 0.3; // 30% from top
      window.scrollTo({
        top: window.scrollY + rect.top - offset,
        behavior: 'smooth'
      });

      // Clear flag after scroll animation completes (~500ms for smooth scroll)
      setTimeout(() => {
        isProgrammaticScroll = false;
      }, 800);

      console.log('✓ HTML highlighting complete');
      return true;
    }
  }

  console.log('✗ No match found in HTML elements');

  // Fallback: highlight first element
  if (pageTextContent.length > 0) {
    const fallbackIndex = Math.min(currentChunkIndex, pageTextContent.length - 1);
    const fallbackElem = pageTextContent[fallbackIndex].element;
    fallbackElem.classList.add('highlighted-element');

    // Add has-highlight to parent elements
    let parent = fallbackElem.parentElement;
    while (parent && !parent.classList.contains('html-assessment-container')) {
      parent.classList.add('has-highlight');
      parent = parent.parentElement;
    }

    // Programmatic scroll with flag
    isProgrammaticScroll = true;
    fallbackElem.scrollIntoView({ behavior: 'smooth', block: 'center' });
    setTimeout(() => {
      isProgrammaticScroll = false;
    }, 800);

    console.log(`Using fallback highlighting at element ${fallbackIndex}`);
    return true;
  }

  return false;
}

function highlightChunkInPDF(currentSearchWords, nextSearchWords) {
  // Detect rendering mode and route to appropriate highlighting function
  const isHtmlMode = document.querySelector('.html-assessment-container') !== null;

  if (isHtmlMode) {
    // Use HTML-based highlighting
    highlightChunkInHTML(currentSearchWords, nextSearchWords);
    return;
  }

  // EXISTING PDF HIGHLIGHTING LOGIC BELOW
  // Remove previous highlights and has-highlight classes
  if (currentHighlight) {
    currentHighlight.classList.remove('highlighted-chunk');
  }
  document.querySelectorAll('.highlighted-text-span').forEach(span => {
    span.classList.remove('highlighted-text-span');
  });
  document.querySelectorAll('.has-highlight').forEach(el => {
    el.classList.remove('has-highlight');
  });

  // Use range-based search with current and next chunk boundaries
  console.log('Using range-based highlighting with chunk boundaries...');
  const found = searchTextRangeInPDF(currentSearchWords, nextSearchWords);

  if (found) {
    if (found.spans && found.spans.length > 0) {
      // Highlight all spans in the range
      found.spans.forEach(span => {
        span.classList.add('highlighted-text-span');
      });
      currentHighlight = found.container;

      // Add has-highlight class to page container (parent of textLayer)
      if (found.container && found.container.classList.contains('page-container')) {
        found.container.classList.add('has-highlight');
      }

      // Scroll to first highlighted span (with programmatic scroll flag)
      isProgrammaticScroll = true;
      const rect = found.spans[0].getBoundingClientRect();
      const offset = window.innerHeight * 0.3; // 30% from top
      window.scrollTo({
        top: window.scrollY + rect.top - offset,
        behavior: 'smooth'
      });

      // Clear flag after scroll animation completes
      setTimeout(() => {
        isProgrammaticScroll = false;
      }, 800);

      console.log(`✓ Highlighted ${found.spans.length} spans in the chunk range`);
    } else {
      // Fallback to container highlighting
      found.classList.add('highlighted-chunk');
      found.classList.add('has-highlight');
      currentHighlight = found;

      // Programmatic scroll with flag
      isProgrammaticScroll = true;
      const rect = found.getBoundingClientRect();
      const offset = window.innerHeight * 0.3; // 30% from top
      window.scrollTo({
        top: window.scrollY + rect.top - offset,
        behavior: 'smooth'
      });

      setTimeout(() => {
        isProgrammaticScroll = false;
      }, 800);

      console.log('→ Using page-level highlighting fallback');
    }
  } else {
    console.log('✗ Could not find or highlight chunk');
  }
}

function searchTextInPDF(searchText) {
  if (!searchText) return null;

  // Normalize the search text
  const searchPhrase = normalizeText(searchText);
  console.log('=== SEARCH DEBUG ===');
  console.log('Original search text:', searchText);
  console.log('Normalized search phrase:', searchPhrase);

  const textLayers = document.querySelectorAll('.textLayer');
  console.log('Found', textLayers.length, 'text layers');

  // Try exact normalized phrase match first
  for (let layer of textLayers) {
    const spans = layer.querySelectorAll('span');

    // Method 1: Check each individual span for the complete phrase
    for (let span of spans) {
      const spanText = normalizeText(span.textContent);
      if (spanText && spanText.includes(searchPhrase)) {
        console.log('✓ FOUND normalized phrase in single span');
        return {
          container: layer.closest('.page-container'),
          spans: [span]
        };
      }
    }

    // Method 2: Check combinations of consecutive spans
    for (let i = 0; i < spans.length; i++) {
      let combinedText = '';
      let consecutiveSpans = [];

      // Try combining up to 15 consecutive spans for longer text
      for (let j = i; j < Math.min(i + 15, spans.length); j++) {
        const span = spans[j];
        const spanText = span.textContent.trim();

        if (spanText.length > 0) {
          if (combinedText.length > 0) combinedText += ' ';
          combinedText += spanText;
          consecutiveSpans.push(span);

          // Check normalized text
          const normalizedCombined = normalizeText(combinedText);
          if (normalizedCombined.includes(searchPhrase)) {
            console.log('✓ FOUND normalized phrase across', consecutiveSpans.length, 'spans');
            console.log('  Combined text:', combinedText.substring(0, 100));
            return {
              container: layer.closest('.page-container'),
              spans: consecutiveSpans
            };
          }
        }

        // Stop if combined text is much longer than search phrase
        if (combinedText.length > searchPhrase.length * 4) break;
      }
    }
  }

  // Fallback: Try first 4 normalized words only
  const words = searchPhrase.split(/\s+/).filter(w => w.length > 0);
  if (words.length > 4) {
    const firstFourWords = words.slice(0, 4).join(' ');
    console.log('Trying first 4 normalized words:', firstFourWords);

    for (let layer of textLayers) {
      const spans = layer.querySelectorAll('span');

      for (let i = 0; i < spans.length; i++) {
        let combinedText = '';
        let consecutiveSpans = [];

        for (let j = i; j < Math.min(i + 8, spans.length); j++) {
          const span = spans[j];
          const spanText = span.textContent.trim();

          if (spanText.length > 0) {
            if (combinedText.length > 0) combinedText += ' ';
            combinedText += spanText;
            consecutiveSpans.push(span);

            const normalizedCombined = normalizeText(combinedText);
            if (normalizedCombined.includes(firstFourWords)) {
              console.log('✓ FOUND first 4 normalized words in:', combinedText.substring(0, 100));
              return {
                container: layer.closest('.page-container'),
                spans: consecutiveSpans
              };
            }
          }
        }
      }
    }
  }

  console.log('✗ No text matches found, falling back to page-level highlighting');

  // Final fallback to page container based on chunk index
  const pageContainers = document.querySelectorAll('.page-container');
  if (pageContainers.length > 0) {
    const fallbackPageIndex = Math.min(currentChunkIndex, pageContainers.length - 1);
    console.log('Using fallback page:', fallbackPageIndex + 1);
    return pageContainers[fallbackPageIndex];
  }

  return null;
}

function getAllSpansFromPageRange(startPageContainer, maxPages = 3) {
  /**
   * Collects spans from multiple consecutive pages starting from startPageContainer.
   * Returns an object with all spans and metadata for multi-page highlighting.
   */
  const pageContainers = Array.from(document.querySelectorAll('.page-container'));
  const startPageIndex = pageContainers.indexOf(startPageContainer);

  if (startPageIndex === -1) {
    console.warn('Start page container not found in document');
    return { allSpans: [], spanToPage: new Map(), pages: [], startPageIndex: -1 };
  }

  const result = {
    allSpans: [],
    spanToPage: new Map(), // Track which page each span belongs to
    pages: [],
    startPageIndex: startPageIndex
  };

  // Collect spans from current page and up to maxPages-1 subsequent pages
  for (let i = 0; i < maxPages && (startPageIndex + i) < pageContainers.length; i++) {
    const pageContainer = pageContainers[startPageIndex + i];
    const textLayer = pageContainer.querySelector('.textLayer');
    if (textLayer) {
      const spans = Array.from(textLayer.querySelectorAll('span'));
      spans.forEach(span => result.spanToPage.set(span, pageContainer));
      result.allSpans.push(...spans);
      result.pages.push(pageContainer);
    }
  }

  console.log(`→ Collected ${result.allSpans.length} spans from ${result.pages.length} pages`);
  return result;
}

function findEndBoundaryAfterStart(allSpans, startIndex, endSearchText) {
  /**
   * Searches for endSearchText only in spans AFTER startIndex.
   * This ensures sequential ordering and includes all text between boundaries.
   * Returns the index where the end boundary starts, or -1 if not found.
   */
  const normalizedEndSearch = normalizeText(endSearchText);
  console.log('→ Searching for end boundary in spans after index', startIndex);

  // Search through remaining spans, trying consecutive combinations
  for (let i = startIndex + 1; i < allSpans.length; i++) {
    let combinedText = '';
    let consecutiveSpans = [];

    // Try combining up to 15 consecutive spans (consistent with searchTextInPDF)
    for (let j = i; j < Math.min(i + 15, allSpans.length); j++) {
      const span = allSpans[j];
      const spanText = span.textContent.trim();

      if (spanText.length > 0) {
        if (combinedText.length > 0) combinedText += ' ';
        combinedText += spanText;
        consecutiveSpans.push(span);

        // Check if we've found the end boundary
        const normalizedCombined = normalizeText(combinedText);
        if (normalizedCombined.includes(normalizedEndSearch)) {
          console.log(`→ Found end boundary at span index ${i} (matched across ${consecutiveSpans.length} spans)`);
          return i; // Return where the search started (first span in range)
        }
      }

      // Stop if combined text is much longer than search phrase
      // Increased multiplier to handle large gaps (blank paragraphs, images)
      if (combinedText.length > normalizedEndSearch.length * 10) break;
    }
  }

  // Fallback: Try first 4 words of end search text
  const words = normalizedEndSearch.split(/\s+/).filter(w => w.length > 0);
  if (words.length > 4) {
    const firstFourWords = words.slice(0, 4).join(' ');
    console.log('→ Trying first 4 words of end boundary:', firstFourWords);

    for (let i = startIndex + 1; i < allSpans.length; i++) {
      let combinedText = '';
      let consecutiveSpans = [];

      // Check up to 8 consecutive spans (consistent with searchTextInPDF)
      for (let j = i; j < Math.min(i + 8, allSpans.length); j++) {
        const span = allSpans[j];
        const spanText = span.textContent.trim();

        if (spanText.length > 0) {
          if (combinedText.length > 0) combinedText += ' ';
          combinedText += spanText;
          consecutiveSpans.push(span);

          const normalizedCombined = normalizeText(combinedText);
          if (normalizedCombined.includes(firstFourWords)) {
            console.log(`→ Found end boundary using first 4 words at span index ${i}`);
            return i; // Return where the search started
          }
        }
      }
    }
  }

  return -1; // Not found
}

function searchTextRangeInPDF(startSearchText, endSearchText) {
  /**
   * Searches for a range of text between two searchWords boundaries.
   * Returns all spans from the start position to the end position.
   * If no end boundary, highlights to end of page/document.
   */
  if (!startSearchText) return null;

  console.log('=== RANGE SEARCH DEBUG ===');
  console.log('Start search text:', startSearchText);
  console.log('End search text:', endSearchText || '(end of page)');

  // Find the starting position using the proven searchWords method
  const startResult = searchTextInPDF(startSearchText);
  if (!startResult) {
    console.log('✗ Could not find start boundary');
    return null;
  }

  console.log('✓ Found start boundary');

  // If we only have a container (page-level fallback), return it as-is
  if (!startResult.spans || startResult.spans.length === 0) {
    return startResult;
  }

  // Get spans from current page AND subsequent pages (multi-page support)
  const pageContainer = startResult.container;
  const multiPageData = getAllSpansFromPageRange(pageContainer, 3);

  if (multiPageData.allSpans.length === 0) {
    console.log('✗ No spans found in page range');
    return startResult;
  }

  const allSpans = multiPageData.allSpans;
  const startSpanIndex = allSpans.indexOf(startResult.spans[0]);

  if (startSpanIndex === -1) {
    console.log('✗ Could not locate start span in multi-page range');
    return startResult;
  }

  console.log(`→ Searching across ${multiPageData.pages.length} pages for end boundary`);

  // If no end search text, highlight from start to end of page
  if (!endSearchText) {
    console.log('→ Highlighting from start to end of page');
    return {
      container: pageContainer,
      spans: allSpans.slice(startSpanIndex)
    };
  }

  // Search for end boundary ONLY in spans after the start position
  // This ensures we find the NEXT occurrence and include all text (even after images)
  const endSpanIndex = findEndBoundaryAfterStart(allSpans, startSpanIndex, endSearchText);

  if (endSpanIndex === -1) {
    console.log('→ End boundary not found after start position, highlighting to end of page');
    return {
      container: pageContainer,
      spans: allSpans.slice(startSpanIndex)
    };
  }

  // Return the precise range between boundaries
  console.log(`✓ Precise range: ${endSpanIndex - startSpanIndex} spans (including text after images)`);
  return {
    container: pageContainer,
    spans: allSpans.slice(startSpanIndex, endSpanIndex)
  };
}

function toggleFocusMode() {
  isFocusMode = !isFocusMode;
  document.body.classList.toggle('focus-mode', isFocusMode);

  const focusBtn = document.getElementById('focus-mode-btn');
  focusBtn.classList.toggle('active', isFocusMode);
  focusBtn.setAttribute('aria-pressed', isFocusMode.toString());
}

function handleKeyboard(event) {
  // Don't interfere if user is typing in input fields
  if (event.target.tagName === 'INPUT' || event.target.tagName === 'TEXTAREA') {
    return;
  }

  // Prevent default only for our shortcuts
  switch(event.code) {
    case 'Space':
      event.preventDefault();
      togglePlayPause();
      break;
    case 'ArrowLeft':
      event.preventDefault();
      if (event.ctrlKey || event.metaKey) {
        navigateChunk(-1);
      } else {
        skipTime(-5);
      }
      break;
    case 'ArrowRight':
      event.preventDefault();
      if (event.ctrlKey || event.metaKey) {
        navigateChunk(1);
      } else {
        skipTime(5);
      }
      break;
    case 'KeyF':
      if (!event.ctrlKey && !event.metaKey) {
        event.preventDefault();
        toggleFocusMode();
      }
      break;
    case 'KeyL':
      if (!event.ctrlKey && !event.metaKey) {
        event.preventDefault();
        toggleLoop();
      }
      break;
    case 'ArrowUp':
      if (event.ctrlKey || event.metaKey) {
        event.preventDefault();
        setPlaybackSpeed(currentPlaybackRate + 0.25);
      }
      break;
    case 'ArrowDown':
      if (event.ctrlKey || event.metaKey) {
        event.preventDefault();
        setPlaybackSpeed(currentPlaybackRate - 0.25);
      }
      break;
    case 'Slash':
      if (event.shiftKey) { // '?' key
        event.preventDefault();
        toggleKeyboardHints();
      }
      break;
  }
}

function toggleKeyboardHints() {
  const hints = document.getElementById('keyboard-hints');
  hints.style.display = hints.style.display === 'none' ? 'block' : 'none';
}

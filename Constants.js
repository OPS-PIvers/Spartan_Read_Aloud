const CONSTANTS = {
  // --- General Configuration ---
  AUDIO_DRIVE_FOLDER_NAME: "Assessment Audio Files",
  PDF_SOURCE_FOLDER_NAME: "Assessment PDFs",
  TEACHERS_SHEET_NAME: "Teachers",

  // --- Batch Processing ---
  BATCH_API_ENABLED: false, // KEEP FALSE until Google confirms TTS batch support
  BATCH_CHECK_INTERVAL_MINUTES: 30,

  // --- Automated Batch Processing ---
  AUTOMATED_BATCH_ENABLED: true,
  AUTOMATED_BATCH_INTERVAL_HOURS: 12,
  BATCH_ACCUMULATION_ENABLED: true,

  // --- NEW: Vertex AI Configuration ---
  GCP_PROJECT_ID: PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID'),
  VERTEX_AI_REGION: PropertiesService.getScriptProperties().getProperty('VERTEX_AI_REGION') || 'us-central1', // Default if not set
  GCS_BUCKET_NAME: PropertiesService.getScriptProperties().getProperty('GCS_BUCKET_NAME'),
  VERTEX_AI_ENDPOINT: `https://${PropertiesService.getScriptProperties().getProperty('VERTEX_AI_REGION') || 'us-central1'}-aiplatform.googleapis.com/v1`,

  // --- Spreadsheet Column Mapping (0-based) ---
  COL: {
    PDF_URL: 0,
    CHUNK_COUNT: 1,
    AUDIO_JSON: 2,
    IS_COMPLETE: 3,
    CLASS_NAME: 4,
    INSTRUCTOR: 5,
    PASSWORD: 6,
    STUDENT_EMAILS: 7,
    PROCESSING_STATUS: 8,
    BATCH_JOB_ID: 9,
    LAST_PROCESSED_TIME: 10,
    PROCESSING_MODE: 11,
    READ_ALOUD_ENABLED: 12,
    ACCESS_EXPIRES: 13
  },

  // --- File Handling ---
  SUPPORTED_MIME_TYPES: {
    PDF: MimeType.PDF,
    GOOGLE_DOCS: MimeType.GOOGLE_DOCS,
    MS_WORD: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    MS_WORD_OLD: 'application/msword'
  },
  MAX_FILE_SIZE_MB: 45,
  LARGE_FILE_WARNING_MB: 10,
  MAX_FILENAME_LENGTH: 250,

  // --- Script Execution ---
  SCRIPT_TIMEOUT_MINUTES: 5,

  // --- Text & Audio Processing ---
  // Robust regex that handles numbered list formats: "1.", "1)", "1]" with varying whitespace and line endings
  // Limited to max 3 leading spaces to prevent splitting on deeply indented lists
  CHUNK_SPLIT_REGEX: /[\n\r]+(?=\s{0,3}\d+[.)\]]\s+)/,
  SEARCH_WORDS_COUNT: 8,
  SAFE_FILENAME_WORD_COUNT: 6,

  // --- SSML Pause Configuration ---
  ENABLE_SSML_PAUSES: true,
  PAUSE_AFTER_PARAGRAPH_MS: 800,
  PAUSE_AFTER_ANSWER_CHOICE_MS: 500,
  PAUSE_AFTER_QUESTION_NUMBER_MS: 400,
  PAUSE_BEFORE_ANSWER_BLOCK_MS: 800,
  PAUSE_BEFORE_INLINE_ANSWER_MS: 500,

  // --- TTS Provider Configuration ---
  TTS_PROVIDER: 'GOOGLE_CLOUD',
  GEMINI_TTS_MODEL: 'gemini-2.5-flash-preview-tts', // Model *used* for batch job
  GEMINI_API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/', // Keep for potential fallback/other uses
  GEMINI_VOICE_NAME: "Kore",
  GOOGLE_CLOUD_TTS_VOICE: 'en-US-Standard-H',

  // --- Audio Generation (WAV Header) ---
  WAV_SAMPLE_RATE: 24000,
  WAV_NUM_CHANNELS: 1,
  WAV_BITS_PER_SAMPLE: 16,

  // --- Authentication & Roles ---
  SESSION_TOKEN_DEFAULT_EXPIRY_MINUTES: 180, // 3 hours
  SESSION_TOKEN_STAFF_EXPIRY_MINUTES: 360,  // 6 hours

  // Display roles (used in UI)
  ROLE_SUPER_ADMIN: 'Super Admin',
  ROLE_ADMIN: 'Admin',
  ROLE_TEACHER: 'Teacher',
  ROLE_STUDENT: 'student',

  // Token roles (used in session tokens and validation)
  ROLE_TOKEN_SUPER_ADMIN: 'super_admin',
  ROLE_TOKEN_ADMIN: 'admin',
  ROLE_TOKEN_TEACHER: 'teacher',
  ROLE_TOKEN_STUDENT: 'student',

  // List of staff token roles (for validation)
  STAFF_ROLES: ['admin', 'super_admin', 'teacher'],

  // --- UI & Menu ---
  MENU_NAME: 'Spartan Assessment Portal',
  MENU_ITEMS: {
    RUN_MANUAL: 'Run All Steps (Manual)',
    START_BATCH: 'Start Batch Processing',
    CHECK_BATCH: 'Check Batch Status',
    STOP_BATCH: 'Stop Batch Processing'
  }
};

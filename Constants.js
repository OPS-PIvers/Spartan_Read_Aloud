const CONSTANTS = {
  // --- General Configuration ---
  AUDIO_DRIVE_FOLDER_NAME: "Assessment Audio Files",
  PDF_SOURCE_FOLDER_NAME: "Assessment PDFs",
  TEACHERS_SHEET_NAME: "Teachers",

  // --- Batch Processing ---
  BATCH_API_ENABLED: true,
  BATCH_CHECK_INTERVAL_MINUTES: 30,

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
    PROCESSING_MODE: 11
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
  CHUNK_SPLIT_REGEX: /[\n\r]+(?=\s*\d+[.)\]]\s+)/,
  SEARCH_WORDS_COUNT: 8,
  SAFE_FILENAME_WORD_COUNT: 6,

  // --- Gemini API ---
  GEMINI_TTS_MODEL: 'gemini-2.5-flash-preview-tts',
  GEMINI_API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/',
  GEMINI_VOICE_NAME: "Kore",

  // --- Audio Generation (WAV Header) ---
  WAV_SAMPLE_RATE: 24000,
  WAV_NUM_CHANNELS: 1,
  WAV_BITS_PER_SAMPLE: 16,

  // --- Authentication & Roles ---
  SESSION_TOKEN_DEFAULT_EXPIRY_MINUTES: 180, // 3 hours
  SESSION_TOKEN_STAFF_EXPIRY_MINUTES: 360,  // 6 hours
  STAFF_ROLES: ['admin', 'super_admin', 'teacher'],
  ROLE_SUPER_ADMIN: 'Super Admin',
  ROLE_ADMIN: 'Admin',
  ROLE_TEACHER: 'Teacher',
  ROLE_STUDENT: 'student',

  // --- UI & Menu ---
  MENU_NAME: 'Spartan Read Aloud',
  MENU_ITEMS: {
    RUN_MANUAL: 'Run All Steps (Manual)',
    START_BATCH: 'Start Batch Processing',
    CHECK_BATCH: 'Check Batch Status',
    STOP_BATCH: 'Stop Batch Processing'
  }
};

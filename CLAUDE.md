# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Spartan Read Aloud is a Google Apps Script-based web application that provides text-to-speech functionality for educational assessments. The system processes assessments in multiple formats (PDF, Google Docs, MS Word), generates audio using Google Cloud Text-to-Speech API, and presents both the document and synchronized audio through a web interface for students.

## Development Commands

Since this is a Google Apps Script project, traditional build commands don't apply. Use `clasp` (Google Apps Script CLI) for development:

- **Push local changes to Apps Script**: `clasp push`
- **Pull changes from Apps Script**: `clasp pull`
- **Open project in Apps Script editor**: `clasp open`
- **Deploy as web app**: Use the `/deploy` slash command, which runs `clasp push` followed by `clasp deploy --deploymentId AKfycbwbnej8CBXrgSt7YFbpkAs9uj2f4OYB5518KRjjhP2a6N5RdWNwxVmzUuF54xslyOt6Ww`

The active deployment ID for the web app is: `AKfycbwbnej8CBXrgSt7YFbpkAs9uj2f4OYB5518KRjjhP2a6N5RdWNwxVmzUuF54xslyOt6Ww`

## Architecture

### Core Components

**Backend (Google Apps Script)**:
- `Code.js`: Main application logic for assessment processing, spreadsheet management, web app endpoints, and authentication
- `Constants.js`: Centralized configuration management for TTS providers, file handling, batch processing, roles, and system settings
- `Gemini.js`: Text-to-speech generation supporting both Google Cloud TTS and legacy Gemini API (configurable via `TTS_PROVIDER` in Constants.js)
- `appsscript.json`: Project configuration including OAuth scopes and enabled services

**Frontend (Modular HTML/JavaScript Architecture)**:
- `index.html`: Main application wrapper that includes styles and dynamically loads views based on user role
- `login.html`: Authentication interface for students and staff (teachers, admins, super admins)
- `student.html`: Assessment viewer with PDF.js and HTML rendering, synchronized audio playback, highlighting, and accessibility features
- `teacher.html`: Admin dashboard for uploading assessments, managing metadata, and viewing assessment status
- `styles.html`: Shared CSS styles for the entire application
- `teacher-styles.html`: Admin-specific styles for the dashboard interface

**Data Storage**:
- Google Sheets ("Assessment Database"): Primary database storing assessment URLs, audio metadata, credentials, processing status, batch job IDs, and processing modes
  - Main sheet: Assessment data with columns defined in `Constants.js COL` mapping
  - "Teachers" sheet: Staff authentication and role management
- Google Drive: Organized folder structure for source files and generated audio
  - "Assessment Audio Files" → "Assessment PDFs" (source documents) + individual assessment subfolders (generated audio chunks)

### Processing Pipeline

The system operates in three sequential steps for assessment processing:

1. **Step 0** (`step0_addNewPdfs`): Scans Drive folder for new documents (PDF, Google Docs, MS Word) and adds them to the spreadsheet
2. **Step 1** (`step1_AnalyzePdfsAndCountChunks`):
   - Extracts text from documents using OCR for PDFs or native text extraction for Docs/Word
   - Splits text into numbered question chunks using regex pattern (`CHUNK_SPLIT_REGEX`)
   - Counts chunks and prepares metadata
3. **Step 2** (`step2_GenerateMissingAudioAndFinalize`):
   - Generates audio files via Google Cloud TTS API for each text chunk
   - Creates descriptive filenames based on text content (first 6 words)
   - Stores audio metadata as JSON in spreadsheet
   - Marks assessment as complete

### Batch Processing (Experimental)

The system includes batch processing capabilities configured in Constants.js:

- **Batch API**: `BATCH_API_ENABLED: false` - Batch processing for cost savings (currently disabled)
- **Automated Batch**: Runs every 12 hours (`AUTOMATED_BATCH_INTERVAL_HOURS: 12`)
- **Accumulation**: Collects multiple assessments before batching (`BATCH_ACCUMULATION_ENABLED: true`)
- **Current Limitation**: TTS models don't support batch API on Vertex AI yet

### File Format Support

The system supports multiple assessment file formats:

**Supported Formats**:
- PDF documents (native support)
- Google Docs (converted to HTML for rendering)
- MS Word (.docx, .doc) - automatically converted to PDF

**File Constraints**:
- Maximum file size: 45MB (`MAX_FILE_SIZE_MB`)
- Large file warning threshold: 10MB (`LARGE_FILE_WARNING_MB`)
- Maximum filename length: 250 characters (`MAX_FILENAME_LENGTH`)

**MIME Types** (defined in `Constants.js SUPPORTED_MIME_TYPES`):
- `application/pdf`
- `application/vnd.google-apps.document`
- `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- `application/msword`

### Text-to-Speech (TTS) Provider

**Current Provider**: Google Cloud Text-to-Speech API

The system was switched from Gemini TTS to Google Cloud TTS for cost optimization:

- **API**: Google Cloud Text-to-Speech v1 (`https://texttospeech.googleapis.com/v1/text:synthesize`)
- **Voice**: `en-US-Standard-H` (female) - configurable in Constants.js
- **Pricing**: 4 million free characters per month (Standard voices)
- **Alternative Voice**: `en-US-Standard-I` (male)
- **Audio Format**: LINEAR16 PCM, 24kHz sample rate, mono channel, 16-bit depth
- **Output**: WAV files with proper headers

**Configuration** (in Constants.js):
```javascript
TTS_PROVIDER: 'GOOGLE_CLOUD', // Options: 'GOOGLE_CLOUD' or 'GEMINI'
GOOGLE_CLOUD_TTS_VOICE: 'en-US-Standard-H'
```

**Legacy Support**: Gemini TTS (`gemini-2.5-flash-preview-tts`) still available by changing `TTS_PROVIDER` setting

### Authentication & Role System

The application implements role-based access control with session tokens:

**Roles**:
- **Super Admin** - Full system access, can manage all staff
- **Admin** - Can manage teachers and assessments
- **Teacher** - Can upload and manage own assessments
- **Student** - Can view assigned assessments only

**Session Management**:
- Student sessions: 180 minutes (3 hours)
- Staff sessions: 360 minutes (6 hours)
- Token-based authentication for secure audio file access
- Role verification on all sensitive endpoints

**Staff Management**:
- Teachers stored in "Teachers" sheet with email and role
- Password-based authentication for staff
- Student authentication via email/password per assessment

### Student Interface Features

The student assessment viewer (`student.html`) provides:

**Assessment Selection**:
- Grid view of available assessments for logged-in student
- Assessment cards showing class name, instructor, and status
- Audio preloading with progress bar before starting assessment

**Dual Rendering Engine**:
- **PDF Rendering**: PDF.js with text layer for PDFs
- **HTML Rendering**: Clean, accessible HTML for Google Docs and Word documents
- Automatic detection and routing based on file type

**Audio Playback Controls**:
- Play/pause, skip backward/forward (10 seconds)
- Previous/next chunk navigation
- Playback speed control (0.5x to 2.0x)
- Progress bar with chunk markers
- Visual timeline with time display
- Chunk/question dropdown menu

**Synchronized Highlighting**:
- Real-time text highlighting synchronized with audio playback
- Precise boundary detection (highlights end at last word of chunk, not whitespace)
- Unified highlighting style for both PDF and HTML rendering
- Automatic scrolling to highlighted content

**Accessibility Features**:
- Focus mode: Dims non-highlighted content for better concentration
- High-contrast support for users with visual impairments
- Reduced motion support for users sensitive to animations
- Keyboard shortcuts for all controls
- Generous line spacing (2.0) and limited line length (65 characters) for readability
- Dyslexia-friendly typography and color choices

**Audio Optimization**:
- Bulk preloading in batches (5 chunks at a time)
- Smart caching to prevent redundant downloads
- Loading indicators during audio fetch
- Background preloading of remaining chunks

### Teacher Interface Features

The admin dashboard (`teacher.html`) enables staff to:

**Assessment Upload**:
- **File Upload**: Drag & drop or browse for PDF/Word files (up to 45MB)
- **Google Doc URL**: Direct linking to Google Docs
- File type detection and validation
- Upload progress indicators
- Automatic conversion of Word documents to PDF

**Assessment Management**:
- Set class name and instructor name
- Generate or specify assessment password
- Add student email addresses (comma-separated)
- Processing mode selection (manual vs. batch)

**Assessment Monitoring**:
- List all assessments with status indicators
- View processing status (pending, processing, complete)
- Edit assessment metadata
- Delete assessments
- Batch job tracking (when batch processing enabled)

**Role-Based Access**:
- Super Admins can manage all staff accounts
- Admins can manage teachers
- Teachers can only manage their own assessments

### Key Data Flow

**Assessment Upload Flow**:
1. Staff uploads file or provides Google Doc URL via teacher dashboard
2. File is validated (size, format, permissions)
3. Document is copied to "Assessment PDFs" folder in Drive
4. Metadata row added to spreadsheet with status "Pending"
5. Processing pipeline triggered (manual or batch mode)

**Audio Generation Flow**:
1. Text extracted from document (OCR for PDF, native for Docs/Word)
2. Text split into chunks using numbered question detection
3. For each chunk:
   - Generate descriptive filename (first 6 words)
   - Call Google Cloud TTS API with chunk text
   - Create WAV file with proper headers
   - Upload to assessment-specific Drive subfolder
   - Cache audio file ID and metadata
4. Store complete audio metadata as JSON in spreadsheet column
5. Mark assessment as complete

**Student Authentication Flow**:
1. Student enters email and password on login page
2. Server validates credentials against assessment passwords and student email lists
3. Generate session token with limited expiry
4. Return list of available assessments for that student
5. Student selects assessment, triggers audio preloading
6. Session token required for all audio file fetches (secure access)

**Rendering Flow**:
1. Server determines file type (PDF vs. Google Doc/Word)
2. For PDF: Send base64-encoded PDF data to client
3. For Docs/Word: Convert to sanitized HTML and send to client
4. Client renders using appropriate engine (PDF.js or HTML container)
5. Client extracts text content for chunk matching
6. Audio chunks mapped to document text using searchWords
7. Highlighting applied via DOM manipulation (spans for PDF, classes for HTML)

### Google APIs Used

- **Drive v2**: File storage, management, OCR text extraction, and MIME type detection
- **Sheets**: Database operations for assessment metadata and staff management
- **Documents**: Temporary document creation for OCR processing and Google Docs conversion
- **Google Cloud Text-to-Speech**: Audio generation with Standard voices
- **Gemini API** (optional): Legacy TTS provider if configured

### Configuration Requirements

**Script Properties** (set in Apps Script project settings):
- `GEMINI_API_KEY`: API key for Google Cloud TTS and/or Gemini API

**OAuth Scopes** (defined in appsscript.json):
- `https://www.googleapis.com/auth/spreadsheets` - Read/write spreadsheet data
- `https://www.googleapis.com/auth/drive` - Access Drive files and folders
- `https://www.googleapis.com/auth/documents` - Create temporary documents for OCR
- `https://www.googleapis.com/auth/script.external_request` - Call external APIs (TTS)

**Drive Folder Structure**:
```
Assessment Audio Files/
├── Assessment PDFs/           (source documents uploaded by teachers)
│   ├── assessment1.pdf
│   ├── assessment2.docx
│   └── assessment3 (Google Doc shortcut)
└── [Assessment Name]/         (generated audio chunks, one folder per assessment)
    ├── chunk_0_One_part_of.wav
    ├── chunk_1_In_a_hunter.wav
    └── chunk_2_Damage_to_the.wav
```

**Spreadsheet Structure**:
- Main sheet: One row per assessment with columns defined in `Constants.js COL`:
  - Column 0: PDF_URL (Drive file ID or URL)
  - Column 1: CHUNK_COUNT (number of question chunks)
  - Column 2: AUDIO_JSON (metadata for all audio chunks)
  - Column 3: IS_COMPLETE (boolean flag)
  - Column 4: CLASS_NAME
  - Column 5: INSTRUCTOR
  - Column 6: PASSWORD (for student access)
  - Column 7: STUDENT_EMAILS (comma-separated list)
  - Column 8: PROCESSING_STATUS (pending/processing/complete/error)
  - Column 9: BATCH_JOB_ID (for batch processing)
  - Column 10: LAST_PROCESSED_TIME (timestamp)
  - Column 11: PROCESSING_MODE (manual/batch)
- "Teachers" sheet: Staff authentication with email and role columns

### Recent Improvements

**Highlighting System** (October 2024):
- Fixed highlighting to end precisely at last word of chunk (no trailing whitespace)
- Unified highlighting appearance between PDF and HTML rendering
- Removed visual inconsistencies (extra shadows, opacity differences)

**Focus Mode** (October 2024):
- Auto-dimming of inactive questions to reduce cognitive load
- Smooth opacity transitions for better UX
- Proper parent element handling to avoid stacking issues

**TTS Provider Switch** (October 2024):
- Migrated from Gemini TTS to Google Cloud Standard voices
- Cost reduction: 4M free characters/month vs. paid-only Gemini
- Maintained audio quality with Standard voice models

**File Upload Improvements** (October 2024):
- Enhanced Word-to-PDF conversion reliability
- Better error messages for file upload failures
- Password display in textarea (not concealed) for easier copying

**Batch API Integration** (October 2024):
- Infrastructure prepared for batch processing
- Note: Currently disabled due to lack of TTS model support on Vertex AI

### Known Limitations

1. **Batch Processing**: TTS models don't currently support Vertex AI batch prediction despite infrastructure being ready
2. **File Size**: Very large PDFs (>45MB) cannot be processed due to Apps Script memory limits
3. **OCR Quality**: Text extraction from scanned PDFs depends on image quality
4. **Session Tokens**: Fixed expiry times (no "remember me" option)
5. **Audio Format**: WAV files are larger than MP3; future optimization possible

### Related Documentation

- `GEMINI.md` - Original Gemini API documentation (legacy TTS provider)
- `.claude/commands/deploy.md` - Deployment automation slash command

### Development Tips

1. **Testing Audio**: Use the teacher dashboard to upload a small test assessment (1-2 questions) before processing large batches
2. **Debugging Highlighting**: Check browser console logs - extensive logging shows chunk matching and span selection
3. **Role Testing**: Create test accounts in the "Teachers" sheet with different roles to verify permissions
4. **File Formats**: When testing Word uploads, ensure documents are numbered lists (1., 2., etc.) for proper chunk detection
5. **Cost Monitoring**: Track Google Cloud TTS usage via Google Cloud Console to stay within free tier (4M chars/month)

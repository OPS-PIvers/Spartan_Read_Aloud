# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Spartan Read Aloud is a Google Apps Script-based web application that provides text-to-speech functionality for educational assessments. The system processes PDF assessments, generates audio using the Gemini API, and presents both the PDF and audio through a web interface for students.

## Development Commands

Since this is a Google Apps Script project, traditional build commands don't apply. Use `clasp` (Google Apps Script CLI) for development:

- **Push local changes to Apps Script**: `clasp push`
- **Pull changes from Apps Script**: `clasp pull`
- **Open project in Apps Script editor**: `clasp open`
- **Deploy as web app**: `clasp deploy` (after configuring in Apps Script editor)

## Architecture

### Core Components

**Backend (Google Apps Script)**:
- `Code.js`: Main application logic for PDF processing, spreadsheet management, and web app endpoints
- `Gemini.js`: Gemini API integration for text-to-speech generation
- `appsscript.json`: Project configuration including OAuth scopes and enabled services

**Frontend**:
- `index.html`: Single-page web application with embedded CSS/JavaScript for student interface

**Data Storage**:
- Google Sheets ("Assessment Database"): Acts as the primary database storing PDF URLs, audio metadata, credentials, and processing status
- Google Drive: Stores source PDFs and generated audio files in organized folder structure

### Processing Pipeline

The system operates in three sequential steps (Code.js:36-40):

1. **Step 0** (`step0_addNewPdfs`): Scans Drive folder for new PDFs and adds them to the spreadsheet
2. **Step 1** (`step1_AnalyzePdfsAndCountChunks`): Extracts text from PDFs using OCR and counts text chunks
3. **Step 2** (`step2_GenerateMissingAudioAndFinalize`): Generates audio files via Gemini API and creates final JSON metadata

### Key Data Flow

- PDF URLs and metadata stored in spreadsheet columns (COL constants in Code.js:10-19)
- Text extraction uses Drive API with OCR (Code.js:266-287)
- Audio generation creates descriptive filenames based on text content (Code.js:250-258)
- Web app authenticates students via email/password and serves PDFs with synchronized audio

### Google APIs Used

- **Drive v2**: PDF storage, file management, and OCR text extraction
- **Sheets**: Database operations for assessment metadata
- **Documents**: Temporary document creation for OCR processing
- **Gemini API**: Text-to-speech audio generation

### Configuration Requirements

- Gemini API key stored as script property (`GEMINI_API_KEY`)
- OAuth scopes defined in appsscript.json for Drive, Sheets, Documents access
- Drive folder structure: "Assessment Audio Files" → "Assessment PDFs" (source) + individual assessment subfolders (generated audio)
- The active deployment ID for the GAS web app is "AKfycbylBrrX4PQLzbDhLv_5OjXw6weF0oNVMaMtr7g5WJj4rlxukqHx81qlY2FvWXVLAwLHvw"
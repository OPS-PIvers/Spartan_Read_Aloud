## Purpose

Short, targeted instructions to help an AI code agent be productive in this repository.

## Big picture (one-paragraph)

This repo is a Google Apps Script web app (Apps Script V8 runtime) that ingests assessment documents (PDF/Docs/Word) from Drive, extracts text, splits the text into numbered question "chunks", generates audio for each chunk (Google Cloud TTS or legacy Gemini), stores audio in Drive and metadata in a Google Sheet, and serves a student-facing web UI with synchronized highlighting and playback. Server logic lives in `Code.js` + `Gemini.js`; configuration is centralized in `Constants.js`; UIs are `student.html` and `teacher.html`.

## Key files & responsibilities

- `Code.js` — main server-side application logic and pipeline. Important exported functions to reference: `runAllStepsManual()`, `step0_addNewPdfs()`, `step1_AnalyzePdfsAndCountChunks()`, `step2_GenerateMissingAudioAndFinalize()`, `startBatchProcessing()`, `initiateBatchJobs()`, `createBatchJobForFile()`.
- `Gemini.js` — TTS-specific logic and batch submission helpers; used as fallback/provider for Gemini model.
- `Constants.js` — single source of truth for names, column indices, regexes, limits, and TTS provider selection. Useful constants: `CONSTANTS.COL` mapping, `CHUNK_SPLIT_REGEX`, `TTS_PROVIDER`, `MAX_FILE_SIZE_MB`.
- `student.html`, `teacher.html`, `teacher-styles.html` — front-end views. `student.html` uses PDF.js for PDF rendering and implements synchronized highlighting.
- `appsscript.json` — manifest: OAuth scopes, enabled Advanced Services (Drive v3) and the OAuth2 library. Always check this when adding scopes.
- `README.md`, `CLAUDE.md`, `GEMINI.md` — design notes and developer hints (clasp commands, deployment id) — source these for deploy guidance.

## Data flow (short)

UI (student/teacher HTML) -> google.script.run -> server functions in `Code.js` -> Drive/Sheets/Docs + TTS APIs -> audio files back to Drive -> client preloads audio via signed session tokens. Key mapping: sheet "Assessment Database" columns follow `CONSTANTS.COL` (0-based indices).

## Developer workflows / commands

- Local editing and sync: use clasp (documented in `CLAUDE.md`): `clasp push`, `clasp pull`, `clasp open`.
- Deploy from Apps Script editor or via `clasp deploy` (see `CLAUDE.md` for a project deployment id example).
- No build step required — changes to .js/.html are deployed to Apps Script runtime.

## Project-specific conventions & gotchas

- Spreadsheet schema is positional (0-based indices in `CONSTANTS.COL`). When adding/removing columns, update `Constants.js` instead of searching for string column names.
- Chunk splitting relies on `CHUNK_SPLIT_REGEX` to find numbered questions; adjust it cautiously — tests and UI highlighting depend on consistent chunk boundaries.
- TTS provider selection is controlled via `CONSTANTS.TTS_PROVIDER` (values: `'GOOGLE_CLOUD'` or `'GEMINI'`). Google Cloud TTS is the default for cost reasons; `Gemini.js` is kept for legacy/fallback and batch flows.
- Batch processing is experimental: `CONSTANTS.BATCH_API_ENABLED` is false by default. Code handles fallbacks from batch -> manual processing; preserve those checks when changing batch logic.
- WAV header creation is implemented in `createWavBlob()` (in `Code.js`). Preserve the byte-level behavior when modifying audio paths — downstream code expects standard PCM WAV files.

## Integration & secrets

- Script properties referenced in code (must be set in Apps Script project settings):
  - `GEMINI_API_KEY` (also used as general TTS API key in this repo)
  - `GCP_PROJECT_ID`, `VERTEX_AI_REGION`, `GCS_BUCKET_NAME` (for Vertex/Batch/GCS flows)
  - `SERVICE_ACCOUNT_EMAIL`, `SERVICE_ACCOUNT_PRIVATE_KEY` (for Vertex OAuth2 flows)
- Advanced services / libraries: Drive v3 is enabled; OAuth2 library is used (see `appsscript.json`). When adding new APIs, update `appsscript.json` scopes.

## Example edits an AI may be asked to make (how to be safe)

- Add a new metadata column: update `CONSTANTS.COL` mapping, update any logic that reads/writes the sheet (use column index constants), and update front-end where that metadata is shown.
- Change chunk detection: modify `CHUNK_SPLIT_REGEX` and run an end-to-end check with a sample PDF/Doc to verify chunk counts and highlighting alignment.
- Swap voice or adjust SSML pauses: update `CONSTANTS.GOOGLE_CLOUD_TTS_VOICE` and `PAUSE_*_MS` values; verify generated audio length and client sync.

## Minimal checks before PR

1. Run a small local validation: open project with `clasp open` and test an upload of a tiny, 1–2 question PDF through the teacher UI.
2. Confirm `CONSTANTS.COL` indexes match Sheet columns. Changing indexes without updating sheet code is the most common break.
3. If touching TTS or batch logic, ensure `CONSTANTS.BATCH_API_ENABLED` behavior is preserved and fallbacks remain.
4. Do not print secrets in logs; use `PropertiesService` and the Apps Script project settings.

## Where to look for more context

- `CLAUDE.md` and `GEMINI.md` for deploy notes and historical reasoning (why Google Cloud TTS was chosen).
- `Code.js` (server pipeline) and `Constants.js` (config) are the most important files to read first.

If any section above is unclear or you want examples expanded (examples of sheet rows, a sample chunk JSON, or a quick checklist for swapping TTS providers), tell me which part and I will iterate.
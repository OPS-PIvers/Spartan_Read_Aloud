# Spartan Assessment Portal — Codebase Audit

_Consolidated audit of `Code.js`, `Gemini.js`, `Constants.js`, `student.html`, `teacher.html`,
`teacher-styles.html`, and project configuration. Produced by a multi-agent review covering
security, backend correctness/performance, frontend, and repository hygiene._

This report supersedes the previous shallow audit. Note two corrections to earlier findings:
- The "Google Sheets 50,000-char cell limit" is **already handled** — `setLargeDataInCell`/
  `getLargeDataFromCell` (`Code.js:3464-3531`) offload large values to Drive with a
  `DRIVE_FILE_ID:` pointer, and every read site resolves it. Do not "re-fix" this.
- The documented "Dual Rendering Engine (PDF.js + HTML)" no longer exists on the frontend.
  `student.html` has zero PDF.js references; the server always sends sanitized HTML. Naming like
  `pdf-container`/`getAssessmentPdf`/`onPdfLoadError` is misleading legacy naming.

Findings are ordered by severity. Each item cites `file:line`, the failure scenario, and a fix.

---

## CRITICAL — fix before anything else

These are exploitable by **any** authenticated district user (the web app runs
`executeAs: USER_DEPLOYING`, `access: DOMAIN`, so every top-level function is a callable endpoint
running with the owner's full privileges, and any student can reach it).

### SEC-1 — Client-supplied `email` is trusted as the caller's identity (auth bypass)
`getAssessmentPdf(email, password, assessmentUrl)` (`Code.js:4451`) and
`getStudentAssessmentsForEmail(email)` (`Code.js:4351`) accept the caller's email as a **parameter**
and never verify it against `Session.getActiveUser().getEmail()`.
**Scenario:** a student runs `google.script.run.getAssessmentPdf('teacher@orono.k12.mn.us','',url)`
from the browser console and is treated as that teacher.
**Fix:** derive identity server-side (`const caller = Session.getActiveUser().getEmail()`); ignore any
client-supplied email; key all authorization off `caller`.

### SEC-2 — Privilege escalation: `getAssessmentPdf` mints a staff session token from a client email
Because `email` is attacker-controlled (SEC-1), passing any staff email sets `isStaff = true`,
skips the password gate (`Code.js:4478`), and returns a valid staff `sessionToken` with that person's
real role (`generateSessionToken(cleanEmail, user.userType)`, `Code.js:4509-4511`).
**Scenario:** a student requests a `super_admin` token, then calls every staff endpoint
(`getAllAssessments`, `updateAssessmentRow`, `deleteAssessmentRow`, `reprocessAssessment`,
`updateBetaFeatureMappings`, `generateConsolidatedSubmissionsPdf`). Complete student → super-admin
escalation.
**Fix:** issue tokens only in `doGet` from the verified `Session.getActiveUser()` identity (already
done at `Code.js:3579`). Remove token issuance from `getAssessmentPdf`, or re-derive `isStaff` from
the verified caller.

### BE-1 — Pipeline OCRs each document up to 3× per run, which is also the "stuck Processing" root cause
`extractTextFromFile` (`Code.js:2291`) never caches; one "Run All Steps" pass OCRs the same file three
times: `step1` (`Code.js:1081`), `step2` text (`Code.js:1139`), and `step2` HTML (`Code.js:1185`).
OCR output is **not** byte-identical across runs, so if pass-2's `textChunks.length` differs from the
`CHUNK_COUNT` written by pass-1, `step2` hits `textChunks.length !== totalChunks` (`Code.js:1140`) and
**silently skips the row forever** — no error state, permanent "Processing". Divergent `sra-block-N`
IDs between passes also silently break highlighting even when counts match.
**Fix:** OCR/convert **once** per row per run — have `step1` store the sanitized HTML / structured
chunks via `setLargeDataInCell`, and have `step2` read that instead of re-extracting. Eliminates the
mismatch failure class and cuts OCR/image work ~3×.

---

## HIGH

### Security
- **SEC-3 — Arbitrary Drive file read via staff token.** `getAudioDataAsBase64` (`Code.js:4009`) and
  `getBulkAudioData` (`Code.js:4130`) skip file-allow-listing for staff roles ("Staff roles have
  access to all audio files"). Combined with SEC-2, a forged staff token exfiltrates **any** Drive
  file the owner can read, as base64. **Fix:** validate `fileId` membership in the assessment's
  audio set even for staff.
- **SEC-4 — Unauthenticated PII enumeration.** `getUserByEmail` (`Code.js:3601`),
  `getStudentsByCaseManager` (`Code.js:3685`), `getStudentNameFromDirectory` (`Code.js:5371`),
  `getInstructorEmail` (`Code.js:4580`), and `getStudentAssessmentsForEmail` (`Code.js:4351`, no token
  param) have no session/role check but are directly callable. A domain user can harvest student
  names, emails, roles, and SpEd caseload membership over a wordlist. **Fix:** gate each with a
  validated token, or refactor them so they aren't reachable as endpoints.
- **SEC-5 — Unauthenticated data-mutating / cost-abuse endpoints.** `processNewAssessment`
  (`Code.js:3230`), `triggerImmediateProcessing` (`Code.js:3274`), `step0/1/2`,
  `syncExistingAssessmentPermissions` (`Code.js:3132`), `runAllStepsManual` (`Code.js:204`) validate no
  token; a student can drive up TTS/Vertex billing or churn Drive ACLs. **Fix:** gate every
  state-changing global with a role check.

### Backend correctness / reliability
- **BE-2 — `reprocessAssessment` leaves rows permanently stuck.** It clears `CHUNK_COUNT`/`AUDIO_JSON`/
  `IS_COMPLETE` (`Code.js:3361-3363`) but not `PROCESSING_STATUS`; `processNewAssessment`'s gate
  (`Code.js:3239`) requires `!processingStatus`, so the call silently no-ops while returning
  `{success:true}`. **Fix:** also clear `PROCESSING_STATUS`, `BATCH_JOB_ID`, and reset
  `PROCESSING_MODE`.
- **BE-3 — The only periodic retry is silently dead.** `automatedBatchProcessing` early-returns unless
  **both** `AUTOMATED_BATCH_ENABLED` and `BATCH_API_ENABLED` are true (`Code.js:749`), but
  `setupAutomatedBatchProcessing` checks only the former and shows a success alert (`Code.js:878`). The
  installed 12-hour trigger does nothing, so stuck rows have no automatic recovery. **Fix:** give the
  periodic trigger an unconditional `step2` retry path for manual-mode rows, independent of the batch
  flag.
- **BE-4 — No retry/backoff on TTS calls.** `generateAudioFromTextChunk` (`Gemini.js:40`) and
  `generateAudioWithStandardVoice` (`Gemini.js:107`) are single-attempt; a transient 429/5xx returns
  `null`, which breaks the whole row's `step2` pass (`Code.js:1173`) and forces a full re-OCR to retry
  one chunk. **Fix:** 2-3 retries with exponential backoff (mirror `downloadImageWithRetry`,
  `Code.js:1465`). Same gap on the Drive `Files.create`/`exportDocToHtml` conversion calls.
- **BE-5 — No `LockService` in the pipeline.** `step0/1/2` (`Code.js:970-1225`) take no lock, yet
  `triggerImmediateProcessing` fires a 1-second trigger for `step2` on every upload that can race a
  manual "Run All Steps". Two concurrent `step2` runs can both see a chunk as missing and generate
  **duplicate TTS audio** (double billing) before racing to overwrite the row. **Fix:** wrap the
  row-processing body in a script/row lock.

### Frontend
- **FE-1 — Stale-response race in audio playback.** `playAudio()` (`student.html:2023`) sets the
  current-chunk state synchronously but overwrites `globalAudioPlayer.src` in an async `.then()` with
  no generation guard, so rapid chunk switching can play chunk N-1 while the UI shows chunk N (or
  throws `AbortError`). **Fix:** capture a request token and ignore resolutions for chunks that are no
  longer current.
- **FE-2 — Malformed `AUDIO_JSON` hangs the player.** `playAudio()` sets `isLoadingAudio = true` then
  calls `.match()` on `chunk.audioUrl` (`student.html:2039`) **before** the promise chain; a missing
  `audioUrl` throws synchronously, the `.catch()` never runs, and the play button spins forever. **Fix:**
  validate `chunk.audioUrl` up front and route failures through the existing error path.
- **FE-3 — Missing `withFailureHandler` on critical calls.** `getConstants()` (`teacher.html:1291`) has
  no failure handler and is the only path that hides `#page-loader`; any error leaves the teacher on an
  infinite spinner. Same gap on `deleteRow`/`reprocessRow` (`teacher.html:2187`, `2205`) — destructive
  actions that give no feedback on server failure. **Fix:** add failure handlers that restore UI state
  and toast an error.
- **FE-4 — Accessibility gaps contradicting documented features.** `student.html` has **zero** ARIA
  usage and **no** keyboard shortcuts despite CLAUDE.md claiming both; `teacher.html:68` applies a
  global `* { outline: none }` stripping focus rings from most buttons (WCAG 2.4.7 failure).
  Clickable chunks, the play/pause button, sliders, modals, and toasts are all unusable/announced-less
  by keyboard/screen-reader users. **Fix:** remove the blanket `outline:none`, add `aria-*`/`role`
  attributes, a modal focus trap, `aria-live` on toasts, and implement (or stop documenting) keyboard
  shortcuts.
- **FE-5 — Dead code shipped to production.** `teacher-styles.html` is entirely unused (the `include()`
  helper at `Code.js:3989` is never called; its class names don't match `teacher.html`; it has a
  malformed double `</style>`). **Fix:** delete it.

---

## MEDIUM

### Security
- **SEC-6 — Session tokens are unsigned base64 JSON** (`Code.js:4201`); integrity depends solely on a
  `UserProperties` lookup. Any future path that trusts the decoded payload (e.g. the 10-minute cache at
  `Code.js:4269`) becomes a role-forgery vector. **Fix:** HMAC-sign tokens with a Script-Properties
  secret and verify before trusting `role`/`email`.
- **SEC-7 — `GEMINI_API_KEY` logged.** `debugBatchAPIPayload` builds `...?key=${apiKey}` and
  `Logger.log`s the URL (`Code.js:719`); the key lands in Stackdriver. **Fix:** never log the key;
  remove/guard the debug function.
- **SEC-8 — Stored-DOM XSS via blocklist sanitizer.** `sanitizeHtml` (`Code.js:1814`) only strips
  double-quoted `on*` handlers and `<script>`/`<style>`; single-quoted/unquoted handlers,
  `javascript:`/`data:` URIs, and `<svg onload>` survive and are injected via
  `innerHTML` (`student.html:1921`), then cached and re-served. A crafted source document runs script
  in every viewer. **Fix:** use an allow-list sanitizer (or client-side DOMPurify) and a strict CSP.
- **SEC-9 — Plaintext passwords** stored (`Code.js:2709`, `3083`), compared non-constant-time
  (`Code.js:4478`), and returned to staff clients (`Code.js:2507`). **Fix:** stop returning passwords
  to the client; use constant-time comparison.
- **SEC-10 — Spreadsheet formula injection.** `updateAssessmentRow`/`addNewAssessment` write
  `className`/`instructor`/`password` via `setValue` without neutralizing a leading `= + - @`
  (`Code.js:2703`, `3081`). **Fix:** prefix such values with `'` or reject them.

### Backend
- **BE-6 — Un-batched Sheets writes in loops.** `step0` `appendRow` per file (`Code.js:1020`); `step1`
  three `setValue`s per row (`Code.js:1083`); `step2` four range writes per row (`Code.js:1188`,
  `1211`). N rows → 3-4N API calls plus inconsistent-state risk if the 6-minute limit hits mid-row.
  **Fix:** accumulate into an array and write once with `setValues()` (see `updateBetaFeatureMappings`,
  `Code.js:3941`, as the in-repo reference). Add a time-budget check to `step1` (it has none).
- **BE-7 — `embedImagesAsBase64` is O(images × html-length).** It rebuilds the whole multi-MB HTML
  string per image (`Code.js:1380`). **Fix:** build a replacement map and do one pass.
- **BE-8 — Fragile date parsing.** `updateAssessmentRow` parses `ACCESS_EXPIRES` with the bare `Date`
  constructor on a non-ISO string (`Code.js:2736`); a V8 change could reinterpret it and lock students
  out. **Fix:** use `Utilities.parseDate(str, "America/Chicago", ...)`.
- **BE-9 — Consolidated submissions PDF is public-by-link forever.** `generateConsolidatedSubmissionsPdf`
  writes a PII PDF to Drive root with `ANYONE_WITH_LINK` and no cleanup (`Code.js:5249`). **Fix:** save
  to a restricted subfolder, use domain-restricted sharing, and trash old reports.
- **BE-10 — `step0` drops files past the first 1000.** `Drive.Files.list({pageSize:1000})` has no
  `pageToken` loop (`Code.js:1011`). **Fix:** paginate.

### Frontend
- **FE-6 — Unbounded audio cache / base64 data URIs.** `audioCache` holds base64 WAV strings for the
  whole session with no eviction (`student.html:1486`), and audio is delivered as `data:` URIs. **Fix:**
  use `Blob` + `URL.createObjectURL` (revoke on chunk change) and an LRU window around the current
  chunk.
- **FE-7 — Client uploads aren't size/type pre-validated.** `submitAssessment` reads the whole file to
  base64 before any check (`teacher.html:1754`); a 100 MB file round-trips before server rejection.
  **Fix:** mirror `MAX_FILE_SIZE_MB` client-side and check `file.size`/type before reading.
- **FE-8 — Accumulating document click listener.** `setupEventListeners()` adds a `document` click
  handler on every assessment load with no removal (`student.html:2280`), unlike the resize handlers
  that dedupe. **Fix:** remove the old listener first (or guard).
- **FE-9 — Preload chains keep running across assessment switches** using stale closures, repopulating
  a just-cleared `audioCache` with the old assessment's audio (`student.html:2123`). **Fix:** track an
  active-generation counter and ignore stale callbacks.
- **FE-10 — Duplicated UI code.** `showToast`/`showModal`/`closeModal`/`escapeHtml` and their CSS are
  byte-for-byte duplicated across `student.html` and `teacher.html` (~250 lines). **Fix:** extract a
  shared include and use the already-present (but unused) `include()` helper.
- **FE-11 — Playback/network errors are silent.** Autoplay blocks and fetch failures are only
  `console.error`'d (`student.html:2060`); the student sees the button revert with no explanation.
  **Fix:** surface via `showToast`. Session expiry mid-fetch also isn't wired into failure handlers.

---

## LOW / hygiene

- **REPO-1 — `.clasp.json` committed and not git-ignored.** Contains `scriptId`/`projectId` in a public
  repo. Not secret material, but should be ignored going forward. **Fix:** add `.clasp.json` to
  `.gitignore` (note: it's already tracked, so also `git rm --cached` it if you want it untracked).
- **REPO-2 — CLAUDE.md is stale.** References non-existent `index.html`/`login.html`/`styles.html`;
  documents only 12 of the 19 columns in `Constants.js` (missing `READ_ALOUD_ENABLED`, `ACCESS_EXPIRES`,
  `SUBMISSION_*`, `ASSESSMENT_HTML`, `FILE_NAME`) and omits the `Submissions` sheet; says "Drive v2"
  but `appsscript.json` uses v3; documents 4 OAuth scopes but 7 are declared. **Fix:** update to match
  reality.
- **REPO-3 — `CHUNK_SPLIT_REGEX` is dead** (`Constants.js:74`) — referenced only in docs, never in
  code; real chunking is in `parseHtmlToChunks` (`Code.js:2143`). Three docs tell maintainers to tune a
  constant that does nothing. **Fix:** remove it and correct the docs.
- **REPO-4 — ~750 lines (~14%) of dead batch subsystem** gated behind `BATCH_API_ENABLED:false`
  (`Code.js:217-963`), still wired into the `onOpen` menu, referencing a non-existent constant
  `GEMINI_BATCH_API_ENDPOINT` (`Code.js:725`). **Fix:** move to `BatchProcessing.js` and fix the stale
  reference.
- **REPO-5 — `Code.js` is 5,436 lines / 210 KB in one file.** Apps Script shares one global scope, so a
  pure file split (no code changes) is safe: `Pipeline.js`, `FileConversion.js`, `HtmlSanitizer.js`,
  `Auth.js`, `AssessmentAdmin.js`, `StudentApi.js`, `Submissions.js`, `BatchProcessing.js`, `Utils.js`,
  `Main.js`. Do it incrementally with `clasp push` + smoke test between moves.
- **REPO-6 — Stale docs.** `GEMINI.md` and `assessment_rendering_enhancements.md` describe old/proposed
  architecture as if current; `README.md` is a single title line; no linting/CI. **Fix:** label
  proposals as future work, expand the README, and add a minimal `.eslintrc` + a `clasp status` CI check.
- **BE-11 — Substring name matching.** `isAuthorizedForAssessment` (`Code.js:2597`) and
  `getInstructorEmail` (`Code.js:4580`) use bidirectional `.includes()`, so short/overlapping names can
  authorize the wrong teacher. **Fix:** compare on token boundaries.
- **SEC-11 — Token validation cached 10 min** (`Code.js:4300`), so de-assignment/role changes lag.
  **Fix:** shorten TTL or add a revocation flag.

---

## Suggested remediation order

1. **SEC-1 / SEC-2** — the auth-bypass/privilege-escalation pair. Everything else is secondary until a
   student can no longer mint a super-admin token.
2. **SEC-3, SEC-4, SEC-5** — lock down the remaining ungated endpoints.
3. **BE-1** — cache OCR once per run; this fixes the stuck-Processing class and cuts cost ~3×.
4. **BE-2, BE-3, BE-4, BE-5** — recovery paths, TTS/Drive retry-backoff, and the pipeline lock.
5. **SEC-6, SEC-7, SEC-8, SEC-9, SEC-10** — token signing, API-key logging, sanitizer, password
   handling, formula injection.
6. **FE-1 → FE-5** — playback races, failure handlers, accessibility, and dead-code removal.
7. **BE-6 → BE-10, FE-6 → FE-11** — backend and frontend medium-severity performance/robustness work.
8. Hygiene (**REPO-1 → REPO-6**, **BE-11**, **SEC-11**) and structure as ongoing cleanup.

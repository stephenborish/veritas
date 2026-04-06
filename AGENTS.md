# Veritas Assess — AI Developer Master Guidelines

You are developing **Veritas Assess**, a highly concurrent, secure, and polished K-12 standardized assessment platform built natively on Google Apps Script (GAS). This is a fully deployed, production-grade system — not a prototype. Treat every change with that level of care.

GAS has severe architectural quirks that differ fundamentally from Node.js, React, or standard web environments. Read and internalize every rule in this document before writing a single line of code. Violating any of these constraints risks breaking concurrency safety, corrupting student data, or introducing security vulnerabilities.

---

> **Living Document Rule:** After every major feature addition, bug fix, or refactor, **update this file** to reflect what you discovered. This is especially important when you uncover non-obvious behaviors — platform quirks, implicit data contracts, subtle security constraints, or anything you would not have known without reading the code deeply. Future AI sessions depend on this knowledge.

---

## 1. Core Architecture & Environment Constraints

### Shared Global Namespace
All backend `.gs` files (`Code.gs`, `Data.gs`, `Grading.gs`) share a single global scope. There is no module system.
* 🛑 **NEVER** use ES6 `import`/`export` syntax.
* 🛑 **NEVER** use `require()` or attempt to install npm packages at runtime.
* ✅ Define shared state via top-level `const` objects (e.g., `const DB = {...}`, `const Grader = {...}`).

### Frontend Delivery
The client UI is served via `HtmlService.createHtmlOutputFromFile()`. There is no build step, no bundler, no CDN of your own.
* 🛑 **NEVER** create separate `.css` or `.js` files. All styles and scripts must be inline within `StudentApp.html` and `TeacherApp.html`.
* The app uses `XFrameOptionsMode.ALLOWALL` and a mobile viewport meta tag. Do not alter these.
* External resources (fonts, KaTeX, DOMPurify) are loaded via CDN `<link>`/`<script>` tags at the top of the HTML files.

### External HTTP
All external HTTP requests (including Gemini API calls) must use `UrlFetchApp.fetch()`. `fetch()`, `XMLHttpRequest`, and `axios` are not available on the server side.

### `google.script.run` Constraint (Critical)
The client calls backend functions via `google.script.run`. **This mechanism can ONLY call top-level global functions** — it cannot reach into object methods like `DB.foo()` or `Grader.bar()`. This is why `Code.gs` exists entirely as a thin layer of top-level wrapper functions that delegate into `DB` and `Grader`. Do not bypass this pattern.

### `doGet` Routing
The `doGet(e)` function in `Code.gs` is the single entry point for all web requests. It routes to `StudentApp.html` if the URL contains `?page=student`, a `?code=`, or a `?studentToken=` parameter. All other requests go to `TeacherApp.html`. Do not add additional routing logic without understanding this.

---

## 2. Database Architecture (`Data.gs`)

### Google Sheets as a Database
The app uses a Google Sheets spreadsheet (named `'Veritas Assess — Data'`) as its database. The spreadsheet ID is cached in `PropertiesService` under `VA_SHEET_ID` after the first `initSystem()` call to avoid expensive Drive searches on every request.

### Sheet Schema
The following named sheets exist. Each column maps to a specific field — do not reorder:
* **Courses:** `ID | Name | Blocks | CreatedAt`
* **QSets:** `ID | Name | CourseID | CreatedAt | UpdatedAt | QuestionsJSON | StimuliJSON`
* **Rosters:** `Block | CourseID | StudentsJSON | UpdatedAt`
* **Sessions:** *(active and recently ended sessions)*
* **Archive:** *(ended sessions promoted here by `archiveSession()`)*
* **Responses:** `SessionID | StudentID | StudentName | QuestionID | Answer | IsCorrect | ... | AIScore`
* **Meta:** *(metacognition / confidence data)*
* **Violations:** *(anti-cheat event log)*
* **AIGrades:** `SessionID | StudentID | StudentName | QuestionID | Score | MaxPts | Feedback | Answer | GradedAt | IsOverride | OverrideScore | OverrideFeedback | ContextUsed`

### The `DB` Object
All reads and writes route through the `DB` object in `Data.gs`. Never call `SpreadsheetApp` directly from `Code.gs` or `Grading.gs`. `DB.ss()` resolves the spreadsheet; `DB.sh(name)` gets a named sheet.

### LockService (Mandatory for All Writes)
Because 30+ students may submit answers simultaneously, every single write operation **must** be wrapped in `DB.withLock(() => { ... })`. This uses `LockService.getScriptLock().waitLock(10000)` to queue simultaneous executions. If the lock isn't obtained, `withLock` returns `{ error: 'System is busy...' }` — the client `SyncQueue` handles this gracefully via retry.

### Data Serialization
Complex objects (arrays of answers, session configs, student rosters) are stored as JSON strings in cells. Always `JSON.parse()` on read and `JSON.stringify()` on write. The `QSets` sheet stores `QuestionsJSON` and `StimuliJSON` in separate columns.

### Auto-Formatting Gotcha
When writing student answers, prepend with a single quote (`'`) if the answer could be misinterpreted by Sheets (numbers, booleans, percent strings like `+100%` → `1`).

---

## 3. Client-Server Communication (`StudentApp.html`)

### The `SyncQueue` Object (Never Bypass)
Do not call `google.script.run` directly for student submissions. Always use:
```js
SyncQueue.add('studentSubmitAnswer', [S.sid, S.stuId, qId, answer]);
SyncQueue.add('studentSubmitMeta',   [S.sid, S.stuId, qId, confVal]);
```
`SyncQueue` implements exponential-backoff retry for GAS "System is busy" lock errors. Bypassing it risks silent data loss.

### The `S` (Session State) Object
All student-side state lives in the global `S` object:
* `S.sid` — session ID
* `S.stuId` — student identifier
* `S.qs` — questions array
* `S.ans` — answers map `{ qId: answerValue }`
* `S.conf` — confidence ratings map
* `S.cur` — current question index
* `S.locked` — whether the student is locked out (anti-cheat)
* `S.done` — whether the student has submitted
* `S.mode` — `'lockstep'` or free-navigation
* `S.metaEnabled` — whether confidence rating is active
* `S.lockedQs` / `S.revealedQs` — lockstep control arrays

Keep all DOM updates in sync with this object.

### Lockstep Mode
When `S.mode === 'lockstep'`, navigation is entirely teacher-controlled. Students cannot advance questions themselves. The teacher's dashboard sends `advanceQuestion()` / `goToQuestion()` / `revealAnswer()` / `revealAllAnswers()` / `toggleLockQuestion()` server calls. The student UI polls `studentCheckStatus()` to detect teacher-driven state changes and re-renders accordingly. In lockstep, the previous/next buttons are hidden and the student UI shows a "waiting" state for locked questions.

### Polling Architecture
The student UI uses `setInterval` to poll `studentCheckStatus()` (and `getGradingStatus()` during grading). The teacher UI polls `getLiveResults()` for the live dashboard. These are the primary mechanisms for real-time updates — GAS has no WebSocket or push capability.

---

## 4. Security, Identity & Anti-Cheat

### Student Access Tokens (JWT-like)
Student entry links embed a signed token (`?studentToken=...`). The token is a Base64-URL-encoded JSON payload signed with HMAC-SHA256 using `STUDENT_LINK_SECRET` (stored in `PropertiesService`, auto-generated on first use).

Token format: `<base64url_payload>.<base64url_signature>`

Key functions in `Code.gs`:
* `createStudentAccessToken(session, student)` — issues a signed token (valid 7 days)
* `verifyStudentAccessToken(token)` — validates signature AND expiry, returns decoded payload or `null`
* `constantTimeEquals_(a, b)` — timing-safe string comparison (prevents timing attacks on HMAC verification)
* `signStudentAccessPayload_(payloadSegment)` — HMAC signing
* `rotateStudentLinkSecret()` — immediately invalidates ALL existing tokens by regenerating `STUDENT_LINK_SECRET`; run from Apps Script editor if secret is compromised

The token payload includes: `v`, `sid`, `code`, `block`, `email`, `name`, `firstName`, `lastName`, `normalizedName`, `issuedAt`, `expiresAt` (7 days from issue).

🛑 **NEVER** use `===` for HMAC comparison — always use `constantTimeEquals_()`.
🛑 **NEVER** remove the `expiresAt` check in `verifyStudentAccessToken` — expired tokens must be rejected server-side.

Legacy tokens without `expiresAt` are accepted for backward-compatibility (they pass the check because `payload.expiresAt` is falsy).

### Identity Normalization
Student names are normalized via `normalizeStudentIdentityName_()` (lowercased, whitespace-collapsed) for fuzzy roster matching. This normalization is stored in the token payload itself (`normalizedName` field) so re-computation is consistent.

### Anti-Cheat: Browser Lockout
The student view enforces exam integrity by:
* Requiring the browser to enter fullscreen before the exam begins (with a hard block that re-prompts if they haven't entered).
* Listening for `fullscreenchange` — exiting fullscreen triggers `triggerLockout('fullscreen_exit')`.
* Listening for `visibilitychange` — switching tabs triggers a lockout.
* Intercepting screenshot keystrokes (Mac: Cmd+Shift+3/4/5; Win: Win+PrtSc; Chrome OS: CtrlWindow+PrtSc, etc.).
* Violations call `studentReportViolation()` → `DB.studentReportViolation()`, which logs to the Violations sheet and sets a `locked` flag on the student's session row.
* Teachers can `readmitStudent()` from the live dashboard. Teachers can also `dismissViolation()` from the archived session view.
* The `S.done` flag is set **before** exiting fullscreen on final submission to prevent a false lockout on the finish flow.

### XSS Prevention (DOMPurify)
Because UI rendering relies heavily on template literals injecting server-provided strings (question text, student names, answer choices, AI feedback), **all user-controlled data must be wrapped in `DOMPurify.sanitize(variable)`** before being injected into the DOM. DOMPurify is loaded via CDN in both HTML files.

🛑 This is non-negotiable. Missing a single sanitization call on question text or student answers is a critical XSS vulnerability.

### Input Validation for Teacher Data (`Data.gs`)
* `DB.validateName_(name)` — validates course/question-set names: rejects empty, whitespace-only, and names >200 characters. Called at the top of `createCourse`, `updateCourse`, `createQSet`, `updateQSet`. Returns an error string on failure or `null` if valid.
* Question arrays in `createQSet` / `updateQSet` are capped at **100 questions** to prevent oversized payloads from being stored in Sheets cells.
* `overrideScore(sessId, stuId, qId, score, fb)` — validates: `score` must be a finite non-negative number ≤ `maxPoints`; `fb` is truncated to 2000 chars. Invalid scores return `{ error: '...' }` before touching any data.
* `regradeWithContext(sessId, qId, ctx)` — the teacher-supplied context string is capped at **500 characters** to prevent API abuse and prompt injection via the Gemini context field.

### Image Upload Security (`Code.gs`)
* `uploadImage(base64, filename, mimeType)` enforces a **MIME type whitelist**: only `image/jpeg`, `image/png`, `image/gif`, `image/webp` are accepted. Any other type returns `{ error: 'Invalid image type...' }`.
* Base64 payload is capped at **~5 MB** (`MAX_IMAGE_BASE64_LENGTH_ = 6_700_000` chars) to prevent Drive quota exhaustion.

### Audit Log (`Data.gs`)
All irreversible teacher operations are logged to an `AuditLog` sheet (columns: `Timestamp | Action | Target | Details`). The sheet is created by `initSystem()`.

Logged actions:
| Action | Trigger |
|--------|---------|
| `DELETE_COURSE` | `DB.deleteCourse()` |
| `DELETE_QSET` | `DB.deleteQSet()` |
| `DELETE_ARCHIVE_SESSION` | `deleteArchiveSession()` in Code.gs |
| `END_SESSION` | `DB.endSession()` |
| `READMIT_STUDENT` | `DB.readmitStudent()` |
| `OVERRIDE_SCORE` | `Grader.overrideScore()` |
| `REGRADE_WITH_CONTEXT` | `Grader.regradeWithContext()` |
| `ROTATE_STUDENT_LINK_SECRET` | `rotateStudentLinkSecret()` |

🛑 Do not remove audit log calls from these functions — they are the only server-side record of destructive operations.

---

## 5. Email Delivery (`Code.gs`)

### `MailApp` vs `GmailApp`
The app uses `MailApp.sendEmail()` — **not** `GmailApp`. `MailApp` handles Google Workspace EDU domain restrictions far more gracefully and is the correct choice for school deployments. Do not switch to `GmailApp`.

### URL Resolution Complexity
Generating correct student join URLs is non-trivial due to GAS deployment quirks:
* `ScriptApp.getService().getUrl()` — canonical deployment URL (preferred, may be unavailable in some execution contexts).
* `clientBaseUrl` — passed from `window.location` in the browser (reliable runtime value, used as fallback).
* `PropertiesService.getProperty('DEPLOY_URL')` — stored property (last resort, may be stale after re-deploy).
* Editor preview URLs (containing `userCodeAppPanel` or `script.googleusercontent.com`) are **invalid** for student links and must be detected and rejected by `isPreviewEditorUrl()`.
* Legacy Workspace URLs (`/a/<domain>/macros/s/<id>/exec`) are rewritten to the canonical format (`/a/macros/<domain>/s/<id>/exec`) by `buildStudentJoinUrl()`.

Key URL helpers: `resolveWebAppBaseUrl()`, `resolveStudentLandingUrl()`, `buildStudentJoinUrl()`, `isCanonicalExecUrl()`, `isPreviewEditorUrl()`.

### Personalized Tokens in Emails
Each student's email contains a **unique, pre-signed `studentToken`** embedded in their join URL. This auto-fills their name and session code on the student login screen, eliminating manual entry. The email also displays the session code as a plaintext backup in case the token pre-fill fails.

---

## 6. AI Integration

### Gemini Models in Use
* **`gemini-2.5-flash`** — used for short-answer grading (`Grader.MODEL`). Chosen for speed and cost on per-response batch calls.
* **`gemini-2.5-pro`** — used for the full-class AI analysis report (`generateAIClassReport()`). Higher token budget, more analytical depth.

All calls use `UrlFetchApp.fetch()` to the `v1beta` REST endpoint with `muteHttpExceptions: true`. Always check `response.getResponseCode()` before parsing.

### System Instruction (Grading Persona)
Both `callGeminiBatch` and `callGemini` use the Gemini API's `systemInstruction` field (same pattern as `generateAIClassReport` in Code.gs) to separate the grading persona from the task prompt. The persona is defined once as `Grader.SYSTEM_INSTRUCTION` and enforces:
* Direct, specific, concise feedback — "like margin notes from an expert."
* No filler phrases ("Great job", "Good effort") — these waste tokens and add no pedagogical value.
* Never restate the question in feedback.
* Focus exclusively on scientific accuracy; ignore spelling/grammar unless it changes factual meaning.
* **Anti-hallucination:** "Only reference concepts the student actually wrote — never infer, assume, or fabricate content not present in the answer."
* **Anti-confabulation:** "If the answer is blank or nonsensical, score 0 and say so."
* **Anti-Lost-in-Middle:** "Grade every student with equal care regardless of their position in the list."

🛑 These directives are critical for grading accuracy. Do not weaken or remove them.

### Prompt Design Principles
The grading prompts use numbered `GRADING INSTRUCTIONS` to enforce structured evaluation:
1. Compare answer against rubric and ideal answer.
2. Award points only for correct scientific content.
3. Write **brief, specific feedback (1-3 sentences or fragments)** naming the exact concept correct, partially correct, or missing.
4. (Batch only) Return results for ALL student IDs listed — reinforced with explicit count.

Feedback length is flexible (1-3 sentences or fragments) to allow the AI to be as helpful as appropriate for each answer. The key constraint is specificity — the model must name the actual concept rather than hedging with generalities.

### Batch Grading (Primary Path) & Lost-in-Middle Mitigation
`Grader.callGeminiBatch()` sends up to **5 student responses per API call** (`Grader.BATCH_SIZE`) in a single prompt. The model returns a JSON **array** of `{ id, score, feedback }` objects.

**Batch size history:** 15 (original) → 10 → **5** (current). The aggressive reduction to 5 mitigates the "Lost in the Middle" phenomenon (Liu et al., 2023) where LLMs allocate attention in a U-shaped curve — students in the middle of larger batches receive measurably less accurate grading. At batch size 5, the attention curve is effectively flat. The cost is more API calls (6 for a class of 30 instead of 2), which is an acceptable tradeoff for grading accuracy and consistency.

**Positional anchoring:** Each student in the batch is prefixed with `[N/total]` markers (e.g., `[1/5]`, `[2/5]`) and the prompt header includes the explicit total count (`STUDENT RESPONSES (5 total)`). This gives the model explicit positional awareness and anchors equal attention across all students.

### Single-Response Fallback (Secondary Path)
If a batch call fails (network error, malformed JSON) or a student ID is missing from the batch result, the grader falls back to `Grader.callGemini()` for individual student re-grading. This fallback-within-fallback is intentional for resilience.

### Robust Response Parsing (Never Simplify)
LLMs frequently hallucinate markdown fences (` ```json `) around their response. The parsing pipeline strips these before attempting `JSON.parse()`. For single-response calls, a further regex fallback extracts `score` and `feedback` from partially malformed or truncated JSON. **Do not remove these fallbacks** — they are the difference between graceful degradation and a hard failure.

For batch calls, the parser attempts:
1. Extract the first `[...]` array literal from the raw text.
2. If that fails, parse the full `text` as JSON directly.
3. If both fail, throw to trigger the per-student fallback path.

### MAX_TOKENS Truncation Handling
When the Gemini API returns `finishReason: 'MAX_TOKENS'` for a **batch** call, the `_salvageBatchJSON(text, maxPts)` helper uses regex to extract all complete `{id, score, feedback}` objects from the truncated response. Successfully extracted entries are returned as partial results; student IDs missing from the partial results naturally fall back to single grading via the existing `resultMap` check in `gradeSession()`. If zero entries can be salvaged, an error is thrown to trigger the full per-student fallback path.

For **single** calls, the existing partial-extraction regex fallback (lines 449-456) already handles truncated JSON gracefully — no additional handling needed.

🛑 **Do not remove `_salvageBatchJSON`** — it is critical for preventing total data loss when batches are truncated.

### Token Budget Configuration
| Call Type | `maxOutputTokens` | Temperature | Rationale |
|-----------|-------------------|-------------|-----------|
| Batch grading | 8192 | 0.1 | ~1600 tokens per student (5 students), generous headroom for multi-sentence feedback + JSON |
| Single grading | 4096 | 0.1 | Ample room for 1-3 sentence feedback; ensures feedback is never truncated |
| Class report | 8192 | 0.2 | Longer analytical output with multiple sections |

Before sending data to `generateAIClassReport()`, strip the full session config (which contains bloated question-set metadata) and send only clean arrays of `questions`, `responses`, `metacognition`, and `violations`. This prevents token exhaustion and reduces latency.

### Regrade with Context
`regradeWithContext(sessId, qId, ctx)` re-grades all responses for a single question using a teacher-supplied `ctx` string appended to the prompt as `ADDITIONAL TEACHER CONTEXT`. This context string is also stored in the `AIGrades` sheet (`ContextUsed` column) for auditability.

---

## 7. Background Grading Trigger Pattern

### Why It Exists
GAS web app executions (`doGet`/`doPost` context) have a strict 6-minute execution timeout. Grading a full class of short-answer responses can exceed this limit. Additionally, `ScriptApp.newTrigger()` is **blocked** inside web app executions — you cannot create triggers dynamically from a web request.

### The Solution (Queue + Permanent Trigger)
1. **One-time setup (teacher does this once):** Run `setupGradingTrigger()` from the Apps Script editor. This creates a permanent 1-minute time-based trigger for `checkGradeQueue()`.
2. **On teacher request:** `startAIGrading(sessId)` writes `sessId` into a JSON array stored under `VA_GRADE_QUEUE` in `PropertiesService` and returns immediately.
3. **Every minute:** `checkGradeQueue()` dequeues the first session ID and calls `Grader.gradeSession()` synchronously.
4. **Progress polling:** The teacher UI polls `getGradingStatus(sessId)` via `setInterval` and updates a progress bar based on the `VA_GRADE_STATUS_<sessId>` property that the grader continuously writes.

### Grading Status States
`{ state: 'idle' | 'queued' | 'running' | 'partial' | 'done' | 'error', gradedCount, totalToGrade, errors, message }`

`'partial'` means the 5-minute timeout was hit mid-session — the teacher can trigger grading again to resume from where it stopped (already-graded responses are skipped via the `done` Set).

### Trigger Guard
`startGradingAsync()` checks that `checkGradeQueue` trigger exists **before** writing to the queue. If the trigger is missing, it returns a clear error message instructing the teacher to run `setupGradingTrigger()`. Do not remove this guard.

---

## 8. UI/UX Design System

### Aesthetic
Veritas Assess has a premium glassmorphism aesthetic. Do not introduce Tailwind, Bootstrap, or any external CSS framework. Use only the established CSS variables and component classes.

### Design Tokens (CSS Variables)
* **Backgrounds:** `--bg` (soft gradient page background), `--s` (glassmorphism pane — `rgba(255,255,255,0.85)`)
* **Borders & Shadows:** `--bd`, `--bd2`, `--sh` (subtle shadow), `--sh2` (hover shadow)
* **Brand Colors:** `--teal` (primary actions), `--red` (violations/destructive), `--amb` (warnings/unsure confidence), `--grn` (correct answers), `--blue` (MC question tags), `--pur` (SA question tags)
* **Text:** `--tx`, `--tx2`, `--tx3` (primary, secondary, muted)

### Typography (Google Fonts, loaded via CDN)
* `Outfit` — headers, branding, buttons, large numbers
* `Inter` — body text, inputs, textareas
* `JetBrains Mono` — timers, session codes (monospace precision)

### Standardized Components
* **Buttons:** `.btn.bp` (primary, teal gradient), `.btn.bg` (ghost/secondary), `.btn.bd` (destructive/red)
* **Cards:** `.card`, `.q-card` — glassmorphism blur effect with hover transitions
* **Modals:** `.sys-modal-bg` / `.sys-modal-card`. 🛑 **NEVER** use `alert()`, `confirm()`, or `prompt()` — they trigger browser security warnings (especially in fullscreen). Always use `sysConfirm(title, desc, confirmText, callback, isDestructive)`.
* **Toasts:** `toast('Message', 'ok' | 'err')` — transient notifications

### KaTeX Math Rendering
Both `StudentApp.html` and `TeacherApp.html` load KaTeX (v0.16.9) from CDN. Math expressions are rendered as `<span class="katex-inline" data-latex="...">` elements. After any dynamic DOM update that may contain math, call `document.querySelectorAll('.katex-inline').forEach(el => katex.render(...))`. KaTeX render errors are caught and logged with `console.warn()` — never let a KaTeX error block UI rendering.

The teacher question editor has a rich-text toolbar (`rteTB()`) with math insertion shortcuts for fractions (`frac`), square roots (`sqrt`), and other common expressions.

---

## 9. Question & Session Data Model

### Question Types
* `'mc'` — Multiple choice. Has `choices[]` array and `correctIndices[]` (supports multiple correct answers) or legacy `correctIndex`.
* `'sa'` — Short answer. Has `rubric` (grading criteria for AI), `sampleAnswer` (ideal answer for AI), and `points` (max score).

### Session Config
A session object includes: `sessionId`, `code` (student join code), `block`, `setId`, `setName`, `config` (JSON blob with `qSet`, `courseId`, `metacognitionEnabled`, `summaryConfig`, etc.), and runtime state like `currentQuestionIndex`, `revealedAnswers[]`, `lockedQuestions[]`.

### Summary Config
`updateSummaryConfig(id, cfg)` controls what students see on their post-submission summary screen (e.g., whether to show correct answers, AI feedback, scores). This is a teacher-controlled toggle.

### Stimuli (Question Sets)
Question sets support an optional `stimuli` array (stored in `StimuliJSON` column of `QSets`). Stimuli are reference materials (e.g., diagrams, passages) displayed alongside questions. The `stimuli` array is passed through to the student client in `studentGetQuestions()`.

### Score Syncing
After AI grading or a manual override, `Grader._syncResponseScore()` updates the `Responses` sheet directly (column 8 = AI score, column 7 = `isCorrect` boolean) to keep analytics consistent. This dual-write is intentional — the `Responses` sheet is the source of truth for analytics; `AIGrades` is the source of truth for grading details/feedback.

---

## 10. Drive Image Uploads

Images are uploaded from the teacher's question editor via `uploadImage(base64, filename, mimeType)`. The server:
1. Decodes the base64 string and creates a Drive blob.
2. Stores it in a folder named `'Veritas Assess — Images'` (auto-created on first use via `getOrCreateImageFolder()`).
3. Sets sharing to `ANYONE_WITH_LINK / VIEW`.
4. Returns a `drive.google.com/thumbnail?id=...&sz=w1000` URL for direct embedding.

Images are embedded directly in question text as `<img>` tags within the rich-text editor.

---

## 11. Permission & Initialization

### `AUTHORIZE_SYSTEM()`
This is a one-time setup function that the deployer runs manually from the Apps Script editor. It touches every GAS API the app uses (SpreadsheetApp, DriveApp, MailApp, PropertiesService, LockService) to trigger Google's OAuth consent flow for all required scopes at once. It also sends a confirmation email to the deployer. After running it, the deployer must re-deploy the web app with a new version.

🛑 Do not remove or simplify `AUTHORIZE_SYSTEM()` — it is the primary onboarding step.

### `initSystem()`
Called from the teacher UI on first setup. Creates the `'Veritas Assess — Data'` spreadsheet (if it doesn't exist), creates all required named sheets with their header rows, and stores the spreadsheet ID in `PropertiesService`.

---

## 12. Code Health & Best Practices

### Performance & API Limits
* **Avoid N+1 Queries in Loops:** Never use `getDataRange().getValues()` or `.setValue()` inside of an O(N) loop (like iterating over students or responses). This will quickly exhaust execution time limits.
* **O(1) Map Lookups:** When updating rows in Google Sheets, load the entire data range once before the loop, build a `Map` to index row indices by a unique key (e.g., `studentId`), update the array in memory, and write back using batched `setValues()` or single-row `setValues([row])`.
* **Batched Writes:** Always prefer a single `getRange(...).setValues(2D_ARRAY)` over multiple adjacent `.setValue()` calls.

### Exception Handling
* **Never leave `catch` blocks empty.** Log unexpected errors.
* **Backend:** Use `console.error(e)` or `Logger.log(e)` so errors appear in the Apps Script execution log.

### Spreadsheet Read Performance (Targeted Range Fetch)
* **Avoid `getDataRange().getValues()` for ID lookups.** Downloading the entire sheet into memory is extremely slow and acts as an N+1 equivalent on the payload size.
* **Use Targeted Ranges:** When searching for rows by an ID (e.g., `SessionID`, `StudentID`), fetch only the specific columns needed for matching to drastically reduce data transfer payload and memory overhead.
* **Example:**
  ```javascript
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // only first 2 columns
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === targetId) { ... }
    }
  }
  ```
* **Frontend:** Use `console.error(e)` for critical failures, `console.warn(e)` for non-critical (e.g., KaTeX render failures).
* **Expected fallbacks (do not log in O(N) loops):** `JSON.parse()` failures on plain-string student answers are expected and normal — do not emit a warning per student. Logging in tight loops causes thousands of avoidable noise entries and buries real errors.

### Google Sheets API Optimization
* **Batching `appendRow`**: Never call `sheet.appendRow()` (or a wrapper like `_batchAppendRows`) sequentially inside a loop, especially in a fallback scenario like AI grading. Accumulate the rows in memory and perform a single batched insert outside the loop.
* **Batching `setValue`**: When updating multiple adjacent columns in the same row, consolidate multiple `.setValue()` calls into a single `.setValues([[val1, val2]])` call. This significantly reduces the overhead of API roundtrips.

### API Contract (Code.gs as Single Source of Truth)
`Code.gs` contains a full API contract comment block listing every server-callable function, its parameters, and its return type. Keep this block updated whenever you add or modify server functions. It is the canonical contract between the frontend and backend.

### Backward Compatibility
The codebase maintains a few backward-compatible aliases (e.g., `getGradingStatus` → `getStatus`). Do not remove these without auditing all client callsites first.

---

## 13. Deployment Workflow

When making any change:
1. Edit the relevant `.gs` or `.html` file(s).
2. Review your changes against every applicable rule in this document.
3. Run `npx clasp push` to deploy to the live Apps Script project.
4. Notify the user the push is complete and ready for testing.

You have full authorization to deploy directly via `clasp`. Do not ask for permission to push — just push when the code is ready.

### Performance Note: Fast Row Lookups
To quickly find a matching row in a large sheet without pulling the entire dataset into memory via `getValues()`, use `createTextFinder`. Example: `const matches = sheet.createTextFinder(searchString).matchEntireCell(true).findAll();`. This operates natively within the Sheets API and runs orders of magnitude faster with ~0 runtime memory overhead compared to `getValues()` arrays in App Script's V8 engine.

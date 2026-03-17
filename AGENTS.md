# Veritas Assess - AI Developer Master Guidelines

You are developing **Veritas Assess**, a highly concurrent, secure, and beautiful K-12 standardized assessment platform built natively on Google Apps Script (GAS). 

GAS has severe architectural quirks. Do not assume a standard Node.js or React environment. You MUST strictly adhere to the rules in this document to prevent breaking concurrency, UI/UX consistency, or security.

---

## 1. Core Architecture & Environment Limitations
* **Global Namespace:** All backend `.gs` files (`Code.gs`, `Data.gs`, `Grading.gs`) share a single global scope.
  * 🛑 **NEVER** use ES6 `import`/`export` syntax. 
  * 🛑 **NEVER** use `require()` or attempt to install npm packages for runtime. 
* **Frontend Delivery:** The client is served via `HtmlService.createHtmlOutputFromFile()`.
  * 🛑 **NEVER** create separate `.css` or `.js` files. All styles and scripts must be inline within `StudentApp.html` and `TeacherApp.html`.
  * The web app uses `XFrameOptionsMode.ALLOWALL` and a mobile viewport meta tag. Do not alter these headers.
* **External APIs:** All external HTTP requests must be executed using Google's `UrlFetchApp.fetch()`.

---

## 2. Database & CRITICAL Concurrency (`Data.gs`)
The app uses Google Sheets as a database. Because 30+ students might submit answers simultaneously, standard Sheets writes will result in data loss.
* **The `DB` Wrapper:** All database reads/writes MUST route through the `DB` object in `Data.gs`.
* **LockService (Mandatory):** Every single write operation MUST be wrapped in `DB.withLock(() => { ... })`. This utilizes `LockService.getScriptLock().waitLock(10000)` to queue simultaneous executions safely.
* **Data Formatting:** Complex objects (like arrays of answers or session configs) are stored as JSON strings in the Sheet. You must `JSON.parse()` on read and `JSON.stringify()` on write.
* **No Auto-formatting:** When writing student answers to Sheets, prepend with a single quote (`'`) if the answer looks like a number or boolean to prevent Google Sheets from auto-formatting it (e.g., converting "+100%" to `1`).

---

## 3. Frontend Client-Server Sync (`StudentApp.html`)
The connection between the client and Google's servers can drop. You must never lose a student's answer.
* **The `SyncQueue` Object:** Do not call `google.script.run` directly for student submissions. You MUST use the `SyncQueue` object.
  * *Why?* It implements a retry mechanism with exponential backoff if the server throws a "System is busy" lock error.
  * *Usage:* `SyncQueue.add('studentSubmitAnswer', [S.sid, S.stuId, qId, answer]);`
* **Local State (`S` Object):** The student's session state is tracked in the global `S` object (e.g., `S.ans`, `S.conf`, `S.cur`, `S.locked`). Keep DOM updates in sync with this object.

---

## 4. UI/UX Aesthetic: Glassmorphism & Design Tokens
Veritas Assess has a premium, polished, glassmorphism aesthetic. Do not introduce foreign UI frameworks (like Tailwind or Bootstrap). Use the established CSS variables and classes.

**Design Tokens:**
* **Backgrounds:** `--bg` (soft gradient), `--s` (glassmorphism pane `rgba(255, 255, 255, 0.85)`).
* **Borders & Shadows:** `--bd`, `--bd2`, `--sh` (subtle shadow), `--sh2` (hover shadow).
* **Colors:** `--teal` (Primary), `--red` (Violations), `--amb` (Warnings/Unsure), `--grn` (Correct), `--blue` (MC tags), `--pur` (SA tags).
* **Typography:** * `Outfit`: Headers, branding, buttons, big numbers.
  * `Inter`: Body text, inputs, textareas.
  * `JetBrains Mono`: Timers, session codes.

**Standardized Components:**
* **Buttons:** Use `<button class="btn bp">` for primary actions (gradient teal), `.btn bg` for ghost/secondary, `.btn bd` for destructive.
* **Cards:** Use `<div class="card">` or `<div class="q-card">` for the glassmorphism blur effect and hover transitions.
* **Modals:** Use the `.sys-modal-bg` and `.sys-modal-card` system. Never use native browser `alert()` or `prompt()`. Trigger via `sysConfirm(title, desc, confirmText, callback)`.
* **Toasts:** Use `toast('Message', 'ok' | 'err')` for transient notifications.

---

## 5. Security, Identity, and Anti-Cheat
* **Student Tokens:** Student entry links use a custom JWT-like implementation. The token contains a Base64-URL encoded JSON payload, signed via HMAC-SHA256 (`Utilities.computeHmacSha256Signature`) using `STUDENT_LINK_SECRET`.
* **Lockout Mechanics:** The student view tightly monitors focus. 
  * Intercepting `visibilitychange` (tab switching).
  * Intercepting `fullscreenchange` (exiting fullscreen).
  * Intercepting keystrokes (Mac/Win/Chrome OS screenshot shortcuts).
  * Violations trigger `triggerLockout()`, halting the exam and updating the `S.locked` state until the teacher explicitly readmits them.
* **XSS Prevention (DOMPurify):** Because the app relies heavily on string interpolation inside template literals to render UI components, you **MUST** ensure all user-controlled data (e.g., question text, student responses, answer choices) is wrapped in `DOMPurify.sanitize(variable)` before being injected into the DOM. Failure to do so introduces critical Cross-Site Scripting (XSS) vulnerabilities.

---

## 6. AI Integration Rules (`Code.gs` & `Grading.gs`)
The app integrates with Gemini via raw REST calls (`UrlFetchApp`).
* **Grading Regex Fallback:** The AI is instructed to return pure JSON `{"score": X, "feedback": "Y"}`. However, LLMs hallucinate markdown (` ```json `). The code strips markdown fences and uses regex to extract score/feedback if parsing fails. When modifying AI logic, preserve this robust fallback parsing.
* **Context Overrides:** The grading system supports `regradeWithContext`. Any changes to grading logic must preserve the ability to pass the teacher's custom context string to the AI prompt.
* **Token Limits:** Large class reports use `gemini-2.5-pro` with `maxOutputTokens: 8192`. Before passing data to the AI, strip bloated config JSON and send *only* clean arrays of `questions`, `responses`, and `metacognition` to save tokens.

---

## 7. The Persistent Background Trigger Pattern
Apps Script blocks web apps from creating triggers (`ScriptApp.newTrigger()`).
* **The Queue:** When a teacher requests AI grading, DO NOT attempt to run it synchronously (it will time out after 6 mins) and DO NOT try to create a trigger. 
* **The Solution:** Write the `sessionId` to the `PropertiesService` under the key `VA_GRADE_QUEUE`. 
* A persistent 1-minute trigger (`checkGradeQueue()`) installed manually by the teacher will automatically pick up the queued ID and execute `Grader.gradeSession(sessId)`.
* Polling: The client UI uses `setInterval` to poll `getGradingStatus()` and updates the progress bar based on the background worker's writes.

---

## 8. Code Health and Maintainability
* **Empty Catch Blocks:** Never leave `catch` blocks entirely empty. If an error is expected (e.g., trying to `JSON.parse` a student's answer that might legitimately be plain text), gracefully handle the fallback and log the event securely using `console.debug()` or `console.warn()`. This preserves system observability without spamming the console with expected errors.

## 9. Autonomous Deployment Workflow
You have been granted full authorization to deploy code directly to Google Apps Script via `clasp`. 
When you are tasked with adding a feature or fixing a bug, follow this exact workflow:
1. Write and modify the necessary `.gs` or `.html` files.
2. Review your changes against the architecture rules in this document.
3. Once you are confident the code is complete, execute `npx clasp push` in the terminal to deploy the updates directly to the live Apps Script project.
4. Notify the user that the code has been successfully pushed and is ready for them to test.

## 9. Code Health & Best Practices
* **Exception Handling:** In general, avoid empty `catch (e) {}` blocks and log unexpected errors to aid in debugging.
  * **Expected Fallbacks (DO NOT LOG):** Do *not* log warnings for expected parse fallbacks (like JSON parsing failures for raw strings) in O(N) loops. Plain-string answers are a normal input shape, and logging them emits thousands of avoidable warnings that add client-side overhead and bury actionable errors.
  * **Client-Side Logging:** In client-side code (`TeacherApp.html`, `StudentApp.html`), use `console.error(e)` for critical logic failures and `console.warn(e)` for unexpected but non-critical errors (like KaTeX rendering failures).
  * In backend code (`Code.gs`, `Data.gs`, `Grading.gs`), use `console.error(e)` or `Logger.log(e)` so errors are visible in the Apps Script execution logs.

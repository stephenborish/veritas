// ═══════════════════════════════════════════════════════════════════
//  VERITAS ASSESS v5.3 — Code.gs
//  Optimized for Concurrency, Security, UI/UX, and Custom Email
// ═══════════════════════════════════════════════════════════════════

// ⚠️ RUN THIS FUNCTION ONCE FROM THE EDITOR TO GRANT ALL PERMISSIONS
// This touches every API the app uses so Google prompts for all required scopes.
// After running, re-deploy the web app as "Execute as: Me" / "Anyone" access.
function AUTHORIZE_SYSTEM() {
  try {
    // Spreadsheet access (Data.gs uses SpreadsheetApp extensively)
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty('SHEET_ID') || 'none'
    );
    Logger.log("SpreadsheetApp authorized. Sheet: " + (ss ? ss.getName() : 'N/A'));
  } catch(e) { Logger.log("SpreadsheetApp scope triggered: " + e.message); }

  try {
    // Drive access (image uploads use DriveApp)
    DriveApp.getRootFolder();
    Logger.log("DriveApp authorized.");
  } catch(e) { Logger.log("DriveApp scope triggered: " + e.message); }

  try {
    // Mail access (assessment email delivery)
    MailApp.sendEmail(Session.getActiveUser().getEmail(), "Veritas Assess: Authorized", "All permissions successfully granted. You can now send assessment links to your students.");
    Logger.log("MailApp authorized.");
  } catch(e) { Logger.log("MailApp scope triggered: " + e.message); }

  try {
    // Properties and Lock (used by Data.gs for concurrency)
    PropertiesService.getScriptProperties().getProperty('_auth_check');
    LockService.getScriptLock();
    Logger.log("PropertiesService + LockService authorized.");
  } catch(e) { Logger.log("Properties/Lock scope triggered: " + e.message); }

  Logger.log("AUTHORIZE_SYSTEM complete. Now re-deploy: Deploy > Manage Deployments > Edit > New Version > Execute as Me > Anyone.");
}

function doGet(e) {
  const p = (e.parameter.page || '').toLowerCase();
  const code = e.parameter.code || '';
  const studentToken = normalizeStudentToken(e.parameter.studentToken || e.parameter.stk || '');
  if (p === 'student' || code || studentToken) {
    return HtmlService.createHtmlOutputFromFile('StudentApp')
      .setTitle('Veritas Assess')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
  return HtmlService.createHtmlOutputFromFile('TeacherApp')
    .setTitle('Veritas Assess — Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function normalizeStudentCode(rawCode) {
  if (!rawCode) return '';
  const normalized = String(rawCode).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  const codePattern = /^[A-Z0-9]{2,20}$/;
  return codePattern.test(normalized) ? normalized : '';
}

function normalizeStudentToken(rawToken) {
  return String(rawToken || '').trim().replace(/\s+/g, '').slice(0, 2000);
}

function normalizeStudentIdentityName_(name) {
  return String(name || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeStudentEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function generateInvisibleNonce_() {
  // Invisible Unicode zero-width chars — visually identical but unique per send.
  // Prevents Gmail from threading repeated session link emails together.
  const zwChars = ['\u200B', '\u200C', '\u200D', '\u2060'];
  let nonce = '';
  for (let i = 0; i < 8; i++) {
    nonce += zwChars[Math.floor(Math.random() * 4)];
  }
  return nonce;
}

function getStudentNameParts_(student) {
  const firstName = String((student && student.firstName) || '').trim();
  const lastName = String((student && student.lastName) || '').trim();
  if (firstName || lastName) {
    return { firstName, lastName, fullName: [firstName, lastName].join(' ').trim() };
  }

  const fullName = String((student && student.name) || '').trim();
  const parts = fullName ? fullName.split(/\s+/) : [];
  return {
    firstName: parts[0] || '',
    lastName: parts.length > 1 ? parts.slice(1).join(' ') : '',
    fullName
  };
}

function base64UrlEncode_(value) {
  const encoded = typeof value === 'string'
    ? Utilities.base64EncodeWebSafe(value, Utilities.Charset.UTF_8)
    : Utilities.base64EncodeWebSafe(value);
  return encoded.replace(/=+$/g, '');
}

function base64UrlDecodeToString_(value) {
  const normalized = String(value || '');
  if (!normalized) return '';
  const padded = normalized + '==='.slice((normalized.length + 3) % 4);
  const bytes = Utilities.base64DecodeWebSafe(padded);
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}

function constantTimeEquals_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function getStudentLinkSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('STUDENT_LINK_SECRET');
  if (secret) return secret;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    secret = props.getProperty('STUDENT_LINK_SECRET');
    if (secret) return secret;
    secret = Utilities.getUuid() + '.' + new Date().getTime().toString(36);
    props.setProperty('STUDENT_LINK_SECRET', secret);
    return secret;
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('Error releasing lock in getStudentLinkSecret_: ' + e.toString());
    }
  }
}

function signStudentAccessPayload_(payloadSegment) {
  const sigBytes = Utilities.computeHmacSha256Signature(payloadSegment, getStudentLinkSecret_());
  return base64UrlEncode_(sigBytes);
}

// Rotate the HMAC signing secret — immediately invalidates ALL existing student tokens.
// Run from the Apps Script editor when the secret may have been compromised.
function rotateStudentLinkSecret() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    // Generate inline — do NOT call getStudentLinkSecret_() here because it would
    // attempt to re-acquire the same script lock (GAS locks are not reentrant).
    props.deleteProperty('STUDENT_LINK_SECRET');
    const newSecret = Utilities.getUuid() + '.' + new Date().getTime().toString(36);
    props.setProperty('STUDENT_LINK_SECRET', newSecret);
    DB.logAuditEvent('ROTATE_STUDENT_LINK_SECRET', 'STUDENT_LINK_SECRET', 'Secret rotated at ' + new Date().toISOString());
    Logger.log('rotateStudentLinkSecret: new secret generated. All existing student tokens are now invalid.');
    return { ok: true, message: 'Student link secret rotated. All previously issued student tokens are now invalid.' };
  } catch (e) {
    Logger.log('rotateStudentLinkSecret error: ' + e.toString());
    return { error: 'Failed to rotate secret: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) { Logger.log('Error releasing lock in rotateStudentLinkSecret: ' + e.toString()); }
  }
}

function createStudentAccessToken(session, student) {
  const email = normalizeStudentEmail_((student && student.email) || '');
  const nameParts = getStudentNameParts_(student);
  const fullName = nameParts.fullName || String((student && student.name) || '').trim();
  const now = new Date();
  const payload = {
    v: 1,
    sid: String((session && session.sessionId) || ''),
    code: normalizeStudentCode((session && session.code) || ''),
    block: String((session && session.block) || ''),
    email: email,
    name: fullName,
    firstName: nameParts.firstName,
    lastName: nameParts.lastName,
    normalizedName: normalizeStudentIdentityName_(fullName),
    issuedAt: now.toISOString(),
    expiresAt: new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString()
  };
  const payloadSegment = base64UrlEncode_(JSON.stringify(payload));
  return payloadSegment + '.' + signStudentAccessPayload_(payloadSegment);
}

function verifyStudentAccessToken(token) {
  const normalizedToken = normalizeStudentToken(token);
  if (!normalizedToken) return null;

  const parts = normalizedToken.split('.');
  if (parts.length !== 2) return null;

  const payloadSegment = parts[0];
  const providedSig = parts[1];
  const expectedSig = signStudentAccessPayload_(payloadSegment);
  if (!constantTimeEquals_(providedSig, expectedSig)) return null;

  try {
    const payload = JSON.parse(base64UrlDecodeToString_(payloadSegment));
    if (!payload || Number(payload.v) !== 1 || !payload.sid) return null;
    if (payload.expiresAt && new Date() > new Date(payload.expiresAt)) return null;
    payload.code = normalizeStudentCode(payload.code || '');
    payload.email = normalizeStudentEmail_(payload.email || '');
    payload.name = String(payload.name || '').trim();
    payload.firstName = String(payload.firstName || '').trim();
    payload.lastName = String(payload.lastName || '').trim();
    payload.normalizedName = normalizeStudentIdentityName_(payload.normalizedName || payload.name || [payload.firstName, payload.lastName].join(' '));
    return payload;
  } catch (e) {
    return null;
  }
}

function initSystem() { return DB.init(); }

// ── Courses ──
function createCourse(name, blocks) { return DB.createCourse(name, blocks); }
function getCourses() { return DB.getCourses(); }
function updateCourse(id, name, blocks) { return DB.updateCourse(id, name, blocks); }
function updateCourseBlocks(id, blocks) { return DB.updateCourseBlocks(id, blocks); }
function deleteCourse(id) { return DB.deleteCourse(id); }

// ── Question Sets ──
function createQuestionSet(name, courseId, questions, stimuli) { return DB.createQSet(name, courseId, questions, stimuli || []); }
function getQuestionSets(courseId) { return DB.getQSets(courseId); }
function getQuestionSet(id) { return DB.getQSet(id); }
function updateQuestionSet(id, name, courseId, questions, stimuli) { return DB.updateQSet(id, name, courseId, questions, stimuli || []); }
function deleteQuestionSet(id) { return DB.deleteQSet(id); }

// ── Rosters ──
function saveRoster(block, courseId, students) { return DB.saveRoster(block, courseId, students); }
function getRosters() { return DB.getRosters(); }
function getRoster(block, courseId) { return DB.getRoster(block, courseId); }
function getRostersByCourse(courseId) { return DB.getRostersByCourse(courseId); }
function addStudentToRoster(block, student, courseId) { return DB.addStudent(block, student, courseId); }
function removeStudentFromRoster(block, name, courseId) { return DB.removeStudent(block, name, courseId); }

// ── Sessions ──
function activateSession(config) { return DB.activateSession(config); }
function getActiveSession() { return DB.getActiveSession(); }
function endSession(id) { return DB.endSession(id); }
function getSessionHistory() { return DB.getSessionHistory(); }
function regenerateCode(id) { return DB.regenerateCode(id); }
function setSessionCode(id, code) { return DB.setSessionCode(id, code); }
function advanceQuestion(id) { return DB.advanceQuestion(id); }
function goToQuestion(id, qIndex) { return DB.goToQuestion(id, qIndex); }
function revealAnswer(id, qId) { return DB.revealAnswer(id, qId); }
function revealAllAnswers(id) { return DB.revealAllAnswers(id); }
function setTimer(id, config) { return DB.setTimer(id, config); }
function dismissQuestionTimer(sessId) { try { return DB.dismissQuestionTimer(sessId); } catch(e) { return {error:e.message}; } }
function updateSessionConfig(id, key, val) { return DB.updateSessionConfig(id, key, val); }
function updateSummaryConfig(id, cfg) { return DB.updateSummaryConfig(id, cfg); }
function archiveSession(id) { return DB.archiveSession(id); }
function getArchivedSessions() { return DB.getArchivedSessions(); }
function getArchivedSessionData(id) { return DB.getArchivedSessionData(id); }
function rescoreQuestion(sessId, qId, newAnswerText) { return DB.rescoreQuestion(sessId, qId, newAnswerText); }
function rescoreQuestionFull(sessId, qId, qJsonStr) { return DB.rescoreQuestionFull(sessId, qId, qJsonStr); }
function deleteArchiveSession(id) {
  return DB.withLock(() => {
    const a = DB.sh('Archive');
    if (!a) return false;
    const d = a.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        a.deleteRow(i + 1);
        DB.logAuditEvent('DELETE_ARCHIVE_SESSION', id, '');
        return true;
      }
    }
    return false;
  });
}
function dismissViolation(sessId, timestamp) { return DB.dismissViolation(sessId, timestamp); }

// ── Live ──
function getLiveResults(id) { return DB.getLiveResults(id); }
function getLiveQuestionDetail(sessId, qId) { return DB.getLiveQuestionDetail(sessId, qId); }
function readmitStudent(sessId, stuId) { return DB.readmitStudent(sessId, stuId); }

// ── Student ──
function studentJoin(code, first, last, clientToken, studentToken) { return DB.studentJoin(code, first, last, clientToken, studentToken); }
function studentGetQuestions(sessId, stuId) { return DB.studentGetQuestions(sessId, stuId); }
function studentSubmitAnswer(sessId, stuId, qId, answer) { return DB.studentSubmitAnswer(sessId, stuId, qId, answer); }
function studentSubmitMeta(sessId, stuId, qId, confidence) { return DB.studentSubmitMeta(sessId, stuId, qId, confidence); }
function studentReportViolation(sessId, stuId, type) { return DB.studentReportViolation(sessId, stuId, type); }
function studentCheckStatus(sessId, stuId) { return DB.studentCheckStatus(sessId, stuId); }
function studentFinish(sessId, stuId) { return DB.studentFinish(sessId, stuId); }
function studentGetSummary(sessId, stuId) { return DB.studentGetSummary(sessId, stuId); }

// ── AI Analysis ──
function generateAIClassReport(sessId) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return "Error: GEMINI_API_KEY not set. Go to Apps Script → Project Settings → Script Properties and add your key.";

    const data = DB.getArchivedSessionData(sessId);
    if (!data) return "Error: Could not retrieve session data for AI Analysis.";

    // Strip bulky config to save tokens, extract only what the AI needs for analysis
    const rawQuestions = (data.session.config && data.session.config.qSet ? data.session.config.qSet.questions : (data.session.qSet ? data.session.qSet.questions : [])) || [];
    
    const cleanData = {
      questions: rawQuestions.map(q => ({
        id: q.id,
        text: q.text,
        type: q.type,
        choices: q.choices,
        correctAnswer: q.type === 'mc' ? (q.choices ? q.choices[(q.correctIndices || [q.correctIndex || 0])[0]] : null) : (q.sampleAnswer || q.rubric)
      })),
      responses: data.responses.map(r => ({ questionId: r.questionId, answer: String(r.answer), isCorrect: r.isCorrect === true || r.isCorrect === 'TRUE' })),
      metacognition: data.meta.map(m => ({ questionId: m.questionId, confidence: m.confidence })),
      violations: data.violations.map(v => ({ type: v.type }))
    };

    const sysPrompt = "You are an expert K-12 pedagogy specialist, an assessment researcher, and a learning scientist. You have deep technical expertise in human behavior, metacognition, standardized assessment design, psychometrics, and science communication. Analyze this class dataset.";
    const userPrompt = `
      Please analyze this assessment session data and return a markdown-formatted report containing exactly these three sections:
      
      ## Class-Wide Strengths
      (Identify what the class universally understood well based on high scores and high confidence).
      
      ## Primary Misconceptions Identified
      (Identify specific distractors or common wrong answers chosen by multiple students. Explain what fundamental misunderstanding likely caused them to choose that distractor. 
      CRITICAL: For every misconception you discuss, you MUST explicitly write out the Full Question Text and the Exact Distractor Text inline so the teacher knows exactly what you are referring to without scrolling down. Include exactly how many students chose that distractor.)
      
      ## Recommended Next Steps / Warm-ups
      (Provide actionable suggestions for what the teacher should specifically address in the first 5 minutes of the next class to resolve these misconceptions).
      
      Here is the JSON dataset:
      ${JSON.stringify(cleanData)}
    `;

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=' + key;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: sysPrompt }] },
        contents: [{ parts: [{ text: userPrompt }] }],
        generationConfig: { temperature: 0.2, maxOutputTokens: 8192 }
      }),
      muteHttpExceptions: true
    });

    const httpCode = response.getResponseCode();
    const body = response.getContentText();
    if (httpCode !== 200) {
      throw new Error('Gemini API HTTP ' + httpCode + ': ' + body.substring(0, 400));
    }

    const resData = JSON.parse(body);
    if (!resData.candidates || !resData.candidates.length) {
      throw new Error('No candidates returned from Gemini.');
    }

    let text = ((resData.candidates[0].content && resData.candidates[0].content.parts && resData.candidates[0].content.parts[0].text) || '').trim();
    DB.saveAIClassReport(sessId, text);
    return text;

  } catch (e) {
    Logger.log("AI Analysis Error: " + e.toString());
    return "An error occurred while generating the AI report: " + e.toString();
  }
}
function getStudentDetail(sessId, stuId) { return DB.getStudentDetail(sessId, stuId); }

// ── Analytics ──
function getItemAnalysis(id) { return DB.getItemAnalysis(id); }
function getStudentAnalysis(id) { return DB.getStudentAnalysis(id); }
function getMetacognitionData(id) { return DB.getMetacognitionData(id); }

// ── AI Grading ──
function startAIGrading(sessId) { return Grader.startGradingAsync(sessId); }  // async – returns immediately
function runAIGrading(sessId) { return Grader.gradeSession(sessId); }          // sync  – kept for compatibility
function getStatus(sessId) { return Grader.getStatus(sessId); }
function getGradingStatus(sessId) { return getStatus(sessId); }
function overrideScore(sessId, stuId, qId, score, fb) { return Grader.overrideScore(sessId, stuId, qId, score, fb); }
function regradeWithContext(sessId, qId, ctx) { return Grader.regradeWithContext(sessId, qId, ctx); }

// ── Drive Image Uploads ──
const ALLOWED_IMAGE_MIME_TYPES_ = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
const MAX_IMAGE_BASE64_LENGTH_ = 6700000; // ≈5 MB after base64 encoding overhead

function uploadImage(base64, filename, mimeType) {
  try {
    if (!ALLOWED_IMAGE_MIME_TYPES_.includes(String(mimeType || ''))) {
      return { error: 'Invalid image type. Allowed types: JPEG, PNG, GIF, WebP.' };
    }
    const rawBase64 = base64.split(',')[1] || base64;
    if (rawBase64.length > MAX_IMAGE_BASE64_LENGTH_) {
      return { error: 'Image too large. Maximum size is 5 MB.' };
    }
    const blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), mimeType, filename);
    const folder = getOrCreateImageFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000', fileId: file.getId() };
  } catch(e) {
    return { error: 'Failed to upload image: ' + e.message };
  }
}

function getOrCreateImageFolder() {
  const name = 'Veritas Assess — Images';
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// ── Gmail Integration ──
function sendAssessmentEmails(sessId, clientBaseUrl) {
  const sess = DB.getSessionById(sessId);
  if (!sess) return { error: 'Session not found' };
  
  const courseId = (sess.config && sess.config.courseId) || '';
  const roster = DB.getRoster(sess.block, courseId);
  if (!roster.length) return { error: 'No students in roster for Block ' + sess.block };
  
  let baseUrl = resolveWebAppBaseUrl(clientBaseUrl);
  if (!baseUrl) return { error: 'System could not identify the Web App URL. Please deploy properly.' };
  
  // Ensure the URL is clean before appending query params
  baseUrl = baseUrl.split('?')[0];
  const studentUrl = resolveStudentLandingUrl(baseUrl);
  if (!studentUrl) return { error: 'System could not identify the student landing URL.' };
  
  let sent = 0, skipped = 0, errors = [];
  const subject = 'Your VERITAS Assess Link – ' + new Date().toLocaleDateString('en-US', { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' }) + generateInvisibleNonce_();

  roster.forEach(student => {
    if (!student.email || !student.email.includes('@')) {
      skipped++; return;
    }
    try {
      const fname = student.firstName || student.name.split(' ')[0] || 'Student';
      const studentToken = createStudentAccessToken(sess, student);
      const joinUrl = buildStudentJoinUrl(studentUrl, sess.code, studentToken);
      const html = buildAssessmentEmail(fname, joinUrl, student.email, sess.setName, sess.code);
      
      // Using MailApp. It handles EDU restrictions much better than GmailApp
      MailApp.sendEmail({
        to: student.email,
        subject: subject,
        htmlBody: html,
        name: 'VERITAS Assess'
      });
      
      sent++;
    } catch(e) {
      Logger.log('Email error for ' + student.email + ': ' + e.toString());
      errors.push(student.name + ': ' + e.message);
      skipped++;
    }
  });
  
  if (errors.length > 0 && sent === 0) {
    return { error: 'Failed to send. Google blocked the script. Please run AUTHORIZE_SYSTEM from the Apps Script Editor. E.g.: ' + errors[0] };
  }
  return { sent, skipped, total: roster.length, errors: errors.slice(0, 3) };
}

function sendIndividualAssessmentEmail(sessId, stuId, clientBaseUrl) {
  const sess = DB.getSessionById(sessId);
  if (!sess) return { error: 'Session not found' };
  
  const courseId = (sess.config && sess.config.courseId) || '';
  const roster = DB.getRoster(sess.block, courseId);
  const student = roster.find(s => DB.normalizeStudentName(s.name) === DB.normalizeStudentName(stuId) || s.email === stuId);
  if (!student) return { error: 'Student not found in the block roster.' };
  if (!student.email || !student.email.includes('@')) return { error: 'Student does not have a valid email address.' };
  
  let baseUrl = resolveWebAppBaseUrl(clientBaseUrl);
  if (!baseUrl) return { error: 'System could not identify the Web App URL.' };
  
  baseUrl = baseUrl.split('?')[0];
  const studentUrl = resolveStudentLandingUrl(baseUrl);
  if (!studentUrl) return { error: 'System could not identify the student landing URL.' };
  
  try {
    const fname = student.firstName || student.name.split(' ')[0] || 'Student';
    const studentToken = createStudentAccessToken(sess, student);
    const joinUrl = buildStudentJoinUrl(studentUrl, sess.code, studentToken);
    const html = buildAssessmentEmail(fname, joinUrl, student.email, sess.setName, sess.code);
    const subject = 'Your VERITAS Assess Link – ' + new Date().toLocaleDateString('en-US', { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' }) + generateInvisibleNonce_();

    MailApp.sendEmail({
      to: student.email,
      subject: subject,
      htmlBody: html,
      name: 'VERITAS Assess'
    });

    return { ok: true, sentTo: student.email };
  } catch(e) {
    Logger.log('Email error for single student ' + student.email + ': ' + e.toString());
    return { error: 'Failed to send email: ' + e.message };
  }
}

function resolveWebAppBaseUrl(clientBaseUrl) {
  let baseUrl = '';

  // 1. Canonical deployment URL from the runtime — always correct when available.
  try {
    baseUrl = ScriptApp.getService().getUrl() || '';
  } catch (e) {
    Logger.log('Error getting Web App URL from ScriptApp: ' + e.toString());
  }

  // 2. Caller-provided URL (window.location from the UI) — trustworthy runtime value.
  if (!baseUrl) {
    baseUrl = clientBaseUrl || '';
  }

  // 3. Stored property as last resort (may be stale after re-deploy).
  if (!baseUrl) {
    baseUrl = PropertiesService.getScriptProperties().getProperty('DEPLOY_URL') || '';
  }

  // Guard against internal editor preview links that break for students.
  if (baseUrl.indexOf('userCodeAppPanel') > -1) {
    try {
      baseUrl = ScriptApp.getService().getUrl() || baseUrl;
    } catch (e) {
      Logger.log('Error resolving Web App URL for editor preview: ' + e.toString());
    }
  }

  return String(baseUrl).split('?')[0];
}


function resolveStudentLandingUrl(fallbackBaseUrl) {
  const props = PropertiesService.getScriptProperties();
  const configuredLandingUrl = props.getProperty('STUDENT_LANDING_URL') || '';
  const canonicalExecUrl = isCanonicalExecUrl(fallbackBaseUrl) ? fallbackBaseUrl : '';
  const candidates = [
    canonicalExecUrl,
    configuredLandingUrl,
    fallbackBaseUrl || ''
  ];

  for (let i = 0; i < candidates.length; i++) {
    const candidate = String(candidates[i] || '').trim().split('?')[0];
    if (!candidate) continue;
    if (isPreviewEditorUrl(candidate)) continue;
    if (isCanonicalExecUrl(candidate)) return candidate;
    return candidate.endsWith('/') ? candidate : candidate + '/';
  }

  return '';
}

function buildStudentJoinUrl(baseUrl, sessionCode, studentToken) {
  let normalizedBase = String(baseUrl || '').trim();
  
  // Fix Google Workspace legacy execution URLs which break the web app iframe
  const badFormatMatch = normalizedBase.match(/^https:\/\/script\.google\.com\/a\/([^\/]+)\/macros\/s\/([^\/]+)\/exec/i);
  if (badFormatMatch) {
    normalizedBase = 'https://script.google.com/a/macros/' + badFormatMatch[1] + '/s/' + badFormatMatch[2] + '/exec';
  }

  const normalizedCode = normalizeStudentCode(sessionCode);
  const normalizedToken = normalizeStudentToken(studentToken);
  if (!normalizedBase) return '';

  const params = [];
  const hasQuery = normalizedBase.indexOf('?') > -1;

  if (isCanonicalExecUrl(normalizedBase)) {
    params.push('page=student');
  }
  if (normalizedCode) {
    params.push('code=' + encodeURIComponent(normalizedCode));
  }
  if (normalizedToken) {
    params.push('studentToken=' + encodeURIComponent(normalizedToken));
  }

  if (!params.length) {
    return normalizedBase;
  }

  return normalizedBase + (hasQuery ? '&' : '?') + params.join('&');
}


function isPreviewEditorUrl(url) {
  return String(url).indexOf('userCodeAppPanel') > -1 || String(url).indexOf('script.googleusercontent.com') > -1;
}

function isCanonicalExecUrl(url) {
  const normalized = String(url || '').trim();
  return /^https:\/\/script\.google\.com\/.+\/exec$/i.test(normalized);
}



/*
API CONTRACT — SINGLE SOURCE OF TRUTH (Code.gs wrappers -> backend owner)

Each function below is guaranteed to be server-callable and forwards to the listed backend implementation.

DB wrappers
- initSystem() -> DB.init() => { url }
- createCourse(name, blocks) -> DB.createCourse() => { id, name, blocks }
- getCourses() -> DB.getCourses() => Course[]
- updateCourse(id, name, blocks) -> DB.updateCourse() => boolean
- deleteCourse(id) -> DB.deleteCourse() => boolean
- createQuestionSet(name, courseId, questions, stimuli) -> DB.createQSet() => { id, name, questionCount }
- getQuestionSets(courseId) -> DB.getQSets() => QuestionSetSummary[]
- getQuestionSet(id) -> DB.getQSet() => QuestionSet|null
- updateQuestionSet(id, name, courseId, questions, stimuli) -> DB.updateQSet() => boolean
- deleteQuestionSet(id) -> DB.deleteQSet() => boolean
- saveRoster(block, courseId, students) -> DB.saveRoster() => { block, count }
- getRosters() -> DB.getRosters() => Record<string, Roster>
- getRoster(block, courseId) -> DB.getRoster() => Student[]
- getRostersByCourse(courseId) -> DB.getRostersByCourse() => Roster[]
- addStudentToRoster(block, student, courseId) -> DB.addStudent() => boolean
- removeStudentFromRoster(block, name, courseId) -> DB.removeStudent() => boolean
- activateSession(config) -> DB.activateSession() => Session
- getActiveSession() -> DB.getActiveSession() => Session|null
- endSession(id) -> DB.endSession() => boolean
- getSessionHistory() -> DB.getSessionHistory() => ArchiveSummary[]
- regenerateCode(id) -> DB.regenerateCode() => { sessionId, code }|{ error }
- setSessionCode(id, code) -> DB.setSessionCode() => { sessionId, code }|{ error }
- advanceQuestion(id) -> DB.advanceQuestion() => { ok, session }|{ error }
- goToQuestion(id, qIndex) -> DB.goToQuestion() => { ok, session }|{ error }
- revealAnswer(id, qId) -> DB.revealAnswer() => { ok, session }|{ error }
- revealAllAnswers(id) -> DB.revealAllAnswers() => { ok, session }|{ error }
- setTimer(id, config) -> DB.setTimer() => { ok, timer }|{ error }
- updateSessionConfig(id, key, val) -> DB.updateSessionConfig() => { ok, config }|{ error }
- updateSummaryConfig(id, cfg) -> DB.updateSummaryConfig() => { ok, summaryConfig }|{ error }
- archiveSession(id) -> DB.archiveSession() => boolean
- getLiveResults(id) -> DB.getLiveResults() => LiveResults
- getLiveQuestionDetail(sessId, qId) -> DB.getLiveQuestionDetail() => LiveQuestionDetail|null
- readmitStudent(sessId, stuId) -> DB.readmitStudent() => { ok }|{ error }
- studentJoin(code, first, last, clientToken, studentToken) -> DB.studentJoin() => JoinResult|{ error }
- studentGetQuestions(sessId, stuId) -> DB.studentGetQuestions() => StudentQuestionPayload|{ error }
- studentSubmitAnswer(sessId, stuId, qId, answer) -> DB.studentSubmitAnswer() => SaveResult|{ error }
- studentSubmitMeta(sessId, stuId, qId, confidence) -> DB.studentSubmitMeta() => boolean|{ error }
- studentReportViolation(sessId, stuId, type) -> DB.studentReportViolation() => { ok }|{ error }
- studentCheckStatus(sessId, stuId) -> DB.studentCheckStatus() => StatusPayload|{ error }
- studentFinish(sessId, stuId) -> DB.studentFinish() => { done, summary }|{ error }
- studentGetSummary(sessId, stuId) -> DB.studentGetSummary() => StudentSummary
- getStudentDetail(sessId, stuId) -> DB.getStudentDetail() => StudentDetail|{ error }
- getItemAnalysis(id) -> DB.getItemAnalysis() => LiveQuestionDetail[]
- getStudentAnalysis(id) -> DB.getStudentAnalysis() => LiveStudent[]
- getMetacognitionData(id) -> DB.getMetacognitionData() => MetaResponse[]

Grader wrappers
- runAIGrading(sessId) -> Grader.gradeSession() => { gradedCount?, errors?, message? }|{ error }
- getStatus(sessId) -> Grader.getStatus() => { state, sessionId, ... }
- getGradingStatus(sessId) -> getStatus(sessId) (backward-compatible alias)
- overrideScore(sessId, stuId, qId, score, fb) -> Grader.overrideScore() => { ok }|{ error }
- regradeWithContext(sessId, qId, ctx) -> Grader.regradeWithContext() => { ok, updated, message }|{ error }
*/

function buildAssessmentEmail(name, url, email, assessName, sessionCode) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8"> 
    <meta name="viewport" content="width=device-width, initial-scale=1.0"> 
    <meta http-equiv="X-UA-Compatible" content="IE=edge"> 
    <title>VERITAS Assess</title>
    <style>
        body { margin: 0; padding: 0; width: 100% !important; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f3f4f6; color: #333333; line-height: 1.6; }
        .container { max-width: 600px; margin: 40px auto; padding: 0; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        .content-padding { padding: 30px; }
        h1, p, li, span, div, a { font-size: 16px; }
        h1 { margin: 0 0 16px 0; font-weight: 700; color: #12385d; letter-spacing: -0.5px; }
        p { margin: 0 0 16px 0; }
        .btn { display: inline-block; background-color: #12385d; color: #ffffff !important; text-decoration: none; padding: 14px 35px; border-radius: 4px; font-weight: 600; margin: 10px 0; }
        .btn:hover { background-color: #0f2f4d; }
        .divider { height: 1px; background-color: #e5e7eb; margin: 20px 0; border: none; }
        .footer { font-size: 12px; color: #9ca3af; margin-top: 30px; text-align: center; }
        .instructions-list { padding-left: 18px; color: #4b5563; margin-bottom: 0; }
        .instructions-list li { margin-bottom: 8px; }
        @media screen and (max-width: 600px) {
            .container { margin: 0; border-radius: 0; }
            .content-padding { padding: 20px; }
            .btn { display: block; text-align: center; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div style="background-color: #12385d; padding: 20px 30px;">
            <span style="font-weight: 700; font-size: 26px; color: #ffffff;">VERITAS</span> 
            <span style="color: #c5a05a; margin: 0 8px; font-size: 26px;">|</span>
            <span style="color: #e5e7eb; font-size: 22px;">Assessment</span>
        </div>
        <div class="content-padding">
            <p>Hello <strong>${name}</strong>,</p>
            <p>Your secure link for <strong>${assessName}</strong> is ready. Open it below and your session details should load automatically.</p>
            <div style="text-align: center;">
                <a href="${url}" class="btn">Open Veritas</a>
            </div>
            <div style="background:#f8fafc;border:1px solid #d1d5db;border-radius:6px;padding:16px;text-align:center;margin:16px 0;">
                <p style="margin:0 0 8px 0;font-size:12px;letter-spacing:.7px;text-transform:uppercase;color:#6b7280;font-weight:700;">Session Code</p>
                <p style="margin:0;font-size:30px;letter-spacing:2px;font-weight:800;color:#111827;font-family:ui-monospace, SFMono-Regular, Menlo, monospace;">${sessionCode}</p>
                <p style="margin:8px 0 0 0;font-size:13px;color:#6b7280;">If the code box is blank after opening the link, use this backup code.</p>
            </div>
            <hr class="divider">
            <div style="background-color: #ffdcdc; padding: 20px; border-radius: 6px; border: 1px solid #fed7d7;">
                <p style="font-weight: 600; margin-bottom: 12px; margin-top: 0; color: #c53030;">Important Instructions:</p>
                <ul class="instructions-list" style="margin: 0;">
                    <li>You must remain in fullscreen mode throughout the entire session.</li>
                    <li>Don't navigate to other tabs or applications.</li>
                    <li>Do not refresh the browser to prevent disconnection.</li>
                </ul>
            </div>
            <div class="footer">
                <p style="margin-bottom: 8px;">🔒 This is a unique link for <strong>${email}</strong>. Do not share it.</p>
                <p style="margin-bottom: 0;">&copy; 2026 VERITAS ASSESS</p>
            </div>
        </div>
    </div>
</body>
</html>`;
}

function toggleLockQuestion(sessId, qId) {
  try {
    return DB.toggleLockQuestion(sessId, qId);
  } catch (e) {
    return {error: e.message};
  }
}

// ── ADVANCED ANALYTICS WRAPPERS ──
function computeCTTMetrics(sessionId) {
  try { return DB.computeCTTMetrics(sessionId); } catch (e) { return { error: e.message }; }
}
function computeItemDiscrimination(sessionId) {
  try { return DB.computeItemDiscrimination(sessionId); } catch (e) { return { error: e.message }; }
}
function computeConfidenceCalibration(sessionId) {
  try { return DB.computeConfidenceCalibration(sessionId); } catch (e) { return { error: e.message }; }
}
function getCrossSessionRiskReport(courseId, blockFilter) {
  try { return DB.getCrossSessionRiskReport(courseId, blockFilter); } catch (e) { return { error: e.message }; }
}

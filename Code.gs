// ═══════════════════════════════════════════════════════════════════
//  VERITAS ASSESS v5.3 — Code.gs
//  Optimized for Concurrency, Security, UI/UX, and Custom Email
// ═══════════════════════════════════════════════════════════════════

// ⚠️ RUN THIS FUNCTION ONCE FROM THE EDITOR TO GRANT EMAIL PERMISSIONS
function AUTHORIZE_SYSTEM() {
  try {
    MailApp.sendEmail(Session.getActiveUser().getEmail(), "Veritas Assess: Authorized", "Permissions successfully granted. You can now send assessment links to your students.");
    Logger.log("Authorization successful.");
  } catch(e) {
    Logger.log("Authorization failed: " + e.message);
  }
}

function doGet(e) {
  const p = (e.parameter.page || '').toLowerCase();
  const code = e.parameter.code || '';
  if (p === 'student' || code) {
    let html = HtmlService.createHtmlOutputFromFile('StudentApp')
      .setTitle('Veritas Assess')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
      
    if (code) {
      const c = html.getContent().replace('/*@@CODE@@*/', 'window.__code="' + code + '";');
      return HtmlService.createHtmlOutput(c).setTitle('Veritas Assess')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
    return html;
  }
  return HtmlService.createHtmlOutputFromFile('TeacherApp')
    .setTitle('Veritas Assess — Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function initSystem() { return DB.init(); }

// ── Courses ──
function createCourse(name, blocks) { return DB.createCourse(name, blocks); }
function getCourses() { return DB.getCourses(); }
function updateCourse(id, name, blocks) { return DB.updateCourse(id, name, blocks); }
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
function getRoster(block) { return DB.getRoster(block); }
function getRostersByCourse(courseId) { return DB.getRostersByCourse(courseId); }
function addStudentToRoster(block, student) { return DB.addStudent(block, student); }
function removeStudentFromRoster(block, name) { return DB.removeStudent(block, name); }

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
function updateSessionConfig(id, key, val) { return DB.updateSessionConfig(id, key, val); }
function updateSummaryConfig(id, cfg) { return DB.updateSummaryConfig(id, cfg); }

// ── Live ──
function getLiveResults(id) { return DB.getLiveResults(id); }
function getLiveQuestionDetail(sessId, qId) { return DB.getLiveQuestionDetail(sessId, qId); }
function readmitStudent(sessId, stuId) { return DB.readmitStudent(sessId, stuId); }

// ── Student ──
function studentJoin(code, first, last, clientToken) { return DB.studentJoin(code, first, last, clientToken); }
function studentGetQuestions(sessId, stuId) { return DB.studentGetQuestions(sessId, stuId); }
function studentSubmitAnswer(sessId, stuId, qId, answer) { return DB.studentSubmitAnswer(sessId, stuId, qId, answer); }
function studentSubmitMeta(sessId, stuId, qId, confidence) { return DB.studentSubmitMeta(sessId, stuId, qId, confidence); }
function studentReportViolation(sessId, stuId, type) { return DB.studentReportViolation(sessId, stuId, type); }
function studentCheckStatus(sessId, stuId) { return DB.studentCheckStatus(sessId, stuId); }
function studentFinish(sessId, stuId) { return DB.studentFinish(sessId, stuId); }
function studentGetSummary(sessId, stuId) { return DB.studentGetSummary(sessId, stuId); }

// ── Analytics ──
function getItemAnalysis(id) { return DB.getItemAnalysis(id); }
function getStudentAnalysis(id) { return DB.getStudentAnalysis(id); }
function getMetacognitionData(id) { return DB.getMetacognitionData(id); }

// ── AI Grading ──
function runAIGrading(sessId) { return Grader.gradeSession(sessId); }
function getGradingStatus(sessId) { return Grader.getStatus(sessId); }
function overrideScore(sessId, stuId, qId, score, fb) { return Grader.overrideScore(sessId, stuId, qId, score, fb); }
function regradeWithContext(sessId, qId, ctx) { return Grader.regradeWithContext(sessId, qId, ctx); }

// ── Drive Image Uploads ──
function uploadImage(base64, filename, mimeType) {
  try {
    const rawBase64 = base64.split(',')[1] || base64;
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
  
  const roster = DB.getRoster(sess.block);
  if (!roster.length) return { error: 'No students in roster for Block ' + sess.block };
  
  let baseUrl = clientBaseUrl || PropertiesService.getScriptProperties().getProperty('DEPLOY_URL');
  if (!baseUrl) {
    try { baseUrl = ScriptApp.getService().getUrl(); } catch(e) {}
  }
  if (!baseUrl) return { error: 'System could not identify the Web App URL. Please deploy properly.' };
  
  // Ensure the URL is clean before appending query params
  baseUrl = baseUrl.split('?')[0];
  const studentUrl = baseUrl + '?code=' + sess.code;
  
  let sent = 0, skipped = 0, errors = [];
  
  roster.forEach(student => {
    if (!student.email || !student.email.includes('@')) { 
      skipped++; return; 
    }
    try {
      const fname = student.firstName || student.name.split(' ')[0] || 'Student';
      const html = buildAssessmentEmail(fname, studentUrl, student.email, sess.setName);
      const subject = 'Your VERITAS Assess Link – ' + new Date().toLocaleDateString('en-US', { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' });
      
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

function buildAssessmentEmail(name, url, email, assessName) {
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
            <p>Your personalized link for <strong>${assessName}</strong> is ready. Please click the button below to begin.</p>
            <div style="text-align: center;">
                <a href="${url}" class="btn">Begin Session</a>
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

// ═══════════════════════════════════════════════════════════════════
//  Grading.gs v6.0 — Permanent-Trigger + Queue Pattern
//
//  WHY: ScriptApp.newTrigger() is BLOCKED inside web app executions
//  (doGet/doPost context). GAS explicitly forbids creating installable
//  triggers from an API/web-app call.
//
//  HOW: Instead, a permanent 1-minute time-based trigger is installed
//  ONCE by the teacher (via setupGradingTrigger()). That trigger calls
//  checkGradeQueue() every minute. When the web app wants to grade, it
//  writes the session ID to Script Properties. checkGradeQueue() picks
//  it up, runs gradeSession(), and clears the queue.
//
//  SETUP (one-time, in Apps Script editor):
//    Run `setupGradingTrigger()` from the editor menu. That's it.
// ═══════════════════════════════════════════════════════════════════

const Grader = {
  MODEL: 'gemini-2.5-flash',
  QUEUE_KEY: 'VA_GRADE_QUEUE',

  statusKey(sessId) { return 'VA_GRADE_STATUS_' + sessId; },

  setStatus(sessId, obj) {
    PropertiesService.getScriptProperties().setProperty(
      this.statusKey(sessId),
      JSON.stringify(Object.assign({ updatedAt: new Date().toISOString() }, obj || {}))
    );
  },

  getStatus(sessId) {
    const raw = PropertiesService.getScriptProperties().getProperty(this.statusKey(sessId));
    if (!raw) return { state: 'idle', sessionId: sessId };
    try { return JSON.parse(raw); } catch (e) { console.error('Failed to parse status JSON:', e); return { state: 'idle', sessionId: sessId }; }
  },

  // ── ASYNC ENTRY POINT (called from client via google.script.run) ─
  // Just writes the job to the queue and returns immediately.
  // The permanent checkGradeQueue trigger will pick it up within 1 min.
  startGradingAsync(sessId) {
    if (!sessId) return { error: 'No session ID provided.' };

    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) {
      const msg = 'GEMINI_API_KEY not set. Go to Apps Script → Project Settings → Script Properties and add your key.';
      this.setStatus(sessId, { state: 'error', sessionId: sessId, message: msg });
      return { error: msg };
    }

    const current = this.getStatus(sessId);
    if (current.state === 'running') {
      return { queued: false, message: 'Grading is already in progress.', status: current };
    }

    // Check if the permanent trigger is installed
    const triggers = ScriptApp.getProjectTriggers();
    const hasTrigger = triggers.some(t => t.getHandlerFunction() === 'checkGradeQueue');
    if (!hasTrigger) {
      const msg = '⚠ Grading trigger not installed. Open the Apps Script editor and run setupGradingTrigger() once to enable background grading.';
      this.setStatus(sessId, { state: 'error', sessionId: sessId, message: msg });
      return { error: msg };
    }

    // Push sessId into queue array (avoids overwriting concurrent sessions)
    const props = PropertiesService.getScriptProperties();
    let queue;
    try { queue = JSON.parse(props.getProperty(this.QUEUE_KEY) || '[]'); } catch(e) { console.error('Failed to parse queue JSON:', e); queue = []; }
    if (!queue.includes(sessId)) queue.push(sessId);
    props.setProperty(this.QUEUE_KEY, JSON.stringify(queue));

    this.setStatus(sessId, {
      state: 'queued',
      sessionId: sessId,
      gradedCount: 0,
      errors: 0,
      message: 'Grading queued — will begin within 1 minute. This tab updates automatically.'
    });

    return { ok: true, queued: true, sessionId: sessId };
  },

  // ── SYNCHRONOUS GRADER (called by the permanent trigger) ──────────
  gradeSession(sessId) {
    this.setStatus(sessId, { state: 'running', sessionId: sessId, gradedCount: 0, errors: 0, message: 'AI grading in progress...' });
    const startTime = Date.now();
    const TIMEOUT_LIMIT = 5 * 60 * 1000;

    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) {
      const msg = 'GEMINI_API_KEY not set.';
      this.setStatus(sessId, { state: 'error', sessionId: sessId, message: msg });
      return { error: msg };
    }

    const sess = DB.getSessionById(sessId);
    if (!sess) {
      const msg = 'Session not found.';
      this.setStatus(sessId, { state: 'error', sessionId: sessId, message: msg });
      return { error: msg };
    }

    const qSet = DB.getQSet(sess.setId);
    if (!qSet) {
      const msg = 'Question set not found.';
      this.setStatus(sessId, { state: 'error', sessionId: sessId, message: msg });
      return { error: msg };
    }

    const saQs = qSet.questions.filter(q => q.type === 'sa');
    if (!saQs.length) {
      const out = { gradedCount: 0, errors: 0, message: 'No short-answer questions in this question set.' };
      this.setStatus(sessId, Object.assign({ state: 'done', sessionId: sessId }, out));
      return out;
    }

    const resps = DB.getAllResponses(sessId);
    const existing = DB.getAIGrades(sessId);
    const gradeSheet = DB.sh('AIGrades');
    const respSheet = DB.sh('Responses');
    const respData = respSheet.getDataRange().getValues();
    const responseRowByKey = {};
    for (let i = 1; i < respData.length; i++) {
      if (respData[i][0] !== sessId) continue;
      responseRowByKey[respData[i][0] + '|' + respData[i][1] + '|' + respData[i][3]] = i + 1;
    }

    const done = new Set(existing.map(g => g.studentId + '|' + g.questionId));
    let count = 0, errors = 0, lastError = '';
    const totalToGrade = saQs.reduce((sum, q) => {
      return sum + resps.filter(r => r.questionId === q.id && r.answer && String(r.answer).length > 3 && !done.has(r.studentId + '|' + q.id)).length;
    }, 0);

    // If nothing to grade, report done
    if (totalToGrade === 0) {
      const out = { gradedCount: 0, errors: 0, totalToGrade: 0, message: 'All responses already graded.' };
      this.setStatus(sessId, Object.assign({ state: 'done', sessionId: sessId }, out));
      return out;
    }

    this.setStatus(sessId, { state: 'running', sessionId: sessId, gradedCount: 0, errors: 0, totalToGrade, message: `Grading ${totalToGrade} response(s)...` });

    let newGradeRows = [];

    for (const q of saQs) {
      const qResps = resps.filter(r => r.questionId === q.id && r.answer && String(r.answer).length > 3);

      for (const r of qResps) {
        if (Date.now() - startTime > TIMEOUT_LIMIT) {
          if (newGradeRows.length > 0) {
            this._batchAppendRows(gradeSheet, newGradeRows);
          }
          const out = { gradedCount: count, errors, totalToGrade, message: `Graded ${count}/${totalToGrade}. Time limit reached. Run again to finish.` };
          this.setStatus(sessId, Object.assign({ state: 'partial', sessionId: sessId }, out));
          return out;
        }

        const k = r.studentId + '|' + q.id;
        if (done.has(k)) continue;

        try {
          const result = this.callGemini(key, q, r.answer, '');

          newGradeRows.push([
            sessId, r.studentId, r.studentName, q.id,
            result.score, q.points || 1, result.feedback,
            r.answer, new Date().toISOString(),
            false, '', '', ''
          ]);

          const responseRow = responseRowByKey[sessId + '|' + r.studentId + '|' + q.id];
          if (responseRow) respSheet.getRange(responseRow, 8).setValue(result.score);

          count++;
          done.add(k);
          this.setStatus(sessId, { state: 'running', sessionId: sessId, gradedCount: count, errors, totalToGrade, message: `Grading... (${count}/${totalToGrade} done)` });
          Utilities.sleep(1200);

        } catch (e) {
          lastError = e.message || e.toString();
          Logger.log('Grade error ' + r.studentName + '/' + q.id + ': ' + e.toString());
          errors++;
          this.setStatus(sessId, { state: 'running', sessionId: sessId, gradedCount: count, errors, totalToGrade, message: `Error grading ${r.studentName}: ${lastError.substring(0, 150)}` });
          Utilities.sleep(500);
          if (errors > 5) {
            if (newGradeRows.length > 0) {
              this._batchAppendRows(gradeSheet, newGradeRows);
            }
            const out = { gradedCount: count, errors, totalToGrade, message: 'Too many API errors. Last error: ' + lastError };
            this.setStatus(sessId, Object.assign({ state: 'error', sessionId: sessId }, out));
            return out;
          }
        }
      }
    }

    if (newGradeRows.length > 0) {
      this._batchAppendRows(gradeSheet, newGradeRows);
    }

    const out = {
      gradedCount: count,
      errors,
      totalToGrade,
      message: count === 0 && errors > 0
        ? 'Grading failed. Error: ' + lastError
        : 'Graded ' + count + ' response' + (count !== 1 ? 's' : '') + (errors ? ' (' + errors + ' error(s): ' + lastError.substring(0, 120) + ')' : '') + '. ✓ Complete.'
    };
    this.setStatus(sessId, Object.assign({ state: errors ? 'partial' : 'done', sessionId: sessId }, out));
    return out;
  },

  callGemini(key, question, answer, extraContext) {
    const maxPts = question.points || 1;
    let rubricText = '';
    if (question.rubric) rubricText += '\nRUBRIC:\n' + question.rubric;
    if (question.sampleAnswer) rubricText += '\nIDEAL ANSWER:\n' + question.sampleAnswer;
    if (extraContext) rubricText += '\nADDITIONAL TEACHER CONTEXT:\n' + extraContext;

    const prompt = 'You are a science teacher grading a student response. Be precise and direct.\n\n' +
      'QUESTION (' + maxPts + ' pt' + (maxPts > 1 ? 's' : '') + '):\n' + question.text + '\n' +
      rubricText + '\n\nSTUDENT ANSWER:\n"' + answer + '"\n\n' +
      'Reply with ONLY this JSON (no markdown, no prose):\n' +
      '{"score":<0–' + maxPts + '>,"feedback":"<exact notes: what they got right, key missing concept(s), one specific correction>"}';

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + this.MODEL + ':generateContent?key=' + key;

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 4096 }
      }),
      muteHttpExceptions: true
    });

    const httpCode = response.getResponseCode();
    const body = response.getContentText();
    if (httpCode !== 200) {
      throw new Error('Gemini API HTTP ' + httpCode + ': ' + body.substring(0, 400));
    }

    const data = JSON.parse(body);
    if (!data.candidates || !data.candidates.length) {
      throw new Error('No candidates. Response: ' + body.substring(0, 300));
    }

    const candidate = data.candidates[0];
    const finishReason = candidate.finishReason || '';
    if (finishReason === 'MAX_TOKENS') {
      Logger.log('⚠ Gemini hit MAX_TOKENS for: ' + question.text.substring(0, 60));
    }

    let text = ((candidate.content && candidate.content.parts && candidate.content.parts[0].text) || '').trim();
    // Strip markdown fences
    text = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/, '').trim();

    // Try a complete {…} match first
    let jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        const result = JSON.parse(jsonMatch[0]);
        if (result.score !== undefined) {
          return { score: Math.min(Math.max(Number(result.score), 0), maxPts), feedback: result.feedback || 'No feedback generated.' };
        }
      } catch(e) { console.warn('JSON parse failed, falling through to partial extraction:', e); }
    }

    // Fallback: extract score and whatever feedback exists from partial/malformed JSON
    const scoreMatch = text.match(/["']?score["']?\s*:\s*(\d+(?:\.\d+)?)/);
    const feedbackMatch = text.match(/["']?feedback["']?\s*:\s*["']([\s\S]*?)["']\s*(?:,|\})/);
    if (scoreMatch) {
      const score = Math.min(Math.max(Number(scoreMatch[1]), 0), maxPts);
      const feedback = feedbackMatch ? feedbackMatch[1].replace(/\\n/g, ' ') : 'AI feedback was truncated.';
      Logger.log('Partial JSON extraction — score: ' + score);
      return { score, feedback };
    }

    throw new Error('Could not parse Gemini response (finishReason=' + finishReason + '): ' + text.substring(0, 200));
  },


  overrideScore(sessId, stuId, qId, score, fb) {
    const sheet = DB.sh('AIGrades');
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === sessId && d[i][1] === stuId && d[i][3] === qId) {
        sheet.getRange(i + 1, 10).setValue(true);
        sheet.getRange(i + 1, 11).setValue(score);
        sheet.getRange(i + 1, 12).setValue(fb || '');
        this._syncResponseScore(sessId, stuId, qId, Number(score) || 0);
        return { ok: true };
      }
    }
    return { error: 'Grade record not found — run AI grading first.' };
  },

  regradeWithContext(sessId, qId, ctx) {
    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return { error: 'GEMINI_API_KEY not set.' };
    const sess = DB.getSessionById(sessId); if (!sess) return { error: 'Session not found' };
    const qSet = DB.getQSet(sess.setId); if (!qSet) return { error: 'Question set not found' };
    const q = (qSet.questions || []).find(x => x.id === qId); if (!q) return { error: 'Question not found' };
    const resps = DB.getAllResponses(sessId).filter(r => r.questionId === qId && r.answer && String(r.answer).length > 3);
    const gradeSheet = DB.sh('AIGrades');
    let updated = 0;
    let newGradeRows = [];
    resps.forEach(r => {
      try {
        const result = this.callGemini(key, q, r.answer, ctx || '');
        const rows = gradeSheet.getDataRange().getValues();
        let found = false;
        for (let i = 1; i < rows.length; i++) {
          if (rows[i][0] === sessId && rows[i][1] === r.studentId && rows[i][3] === qId) {
            gradeSheet.getRange(i + 1, 5).setValue(result.score);
            gradeSheet.getRange(i + 1, 7).setValue(result.feedback);
            gradeSheet.getRange(i + 1, 13).setValue(ctx || '');
            found = true; break;
          }
        }
        if (!found) {
          newGradeRows.push([sessId, r.studentId, r.studentName, qId, result.score, q.points || 1, result.feedback, r.answer, new Date().toISOString(), false, '', '', ctx || '']);
        }
        this._syncResponseScore(sessId, r.studentId, qId, result.score);
        updated++;
      } catch (e) { Logger.log('Regrade error: ' + e.toString()); }
    });
    if (newGradeRows.length > 0) {
      this._batchAppendRows(gradeSheet, newGradeRows);
    }
    const out = { ok: true, updated, message: 'Regraded ' + updated + ' responses.' };
    this.setStatus(sessId, Object.assign({ state: 'done', sessionId: sessId }, out));
    return out;
  },

  _batchAppendRows(sheet, rows) {
    if (!rows || !rows.length) return;
    return DB.withLock(() => {
      const lastRow = sheet.getLastRow();
      const maxRows = sheet.getMaxRows();
      const neededRows = (lastRow + rows.length) - maxRows;
      if (neededRows > 0) {
        sheet.insertRowsAfter(maxRows, neededRows);
      }
      sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    });
  },

  _syncResponseScore(sessId, stuId, qId, score) {
    const resp = DB.sh('Responses');
    const d = resp.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === sessId && d[i][1] === stuId && d[i][3] === qId) {
        const max = Number(d[i][8]) || 0;
        resp.getRange(i + 1, 8).setValue(score);
        resp.getRange(i + 1, 7).setValue(max > 0 && Number(score) >= max);
        break;
      }
    }
  }

};

// ── PERMANENT QUEUE CHECKER (runs every 1 minute via installable trigger) ─
// Install ONCE by running setupGradingTrigger() from the Apps Script editor.
function checkGradeQueue() {
  const props = PropertiesService.getScriptProperties();
  let queue;
  try { queue = JSON.parse(props.getProperty(Grader.QUEUE_KEY) || '[]'); } catch(e) { console.error('Failed to parse queue JSON:', e); queue = []; }
  if (!queue.length) return; // Nothing queued

  // Dequeue the first session and save remaining back
  const sessId = queue.shift();
  props.setProperty(Grader.QUEUE_KEY, JSON.stringify(queue));

  Logger.log('checkGradeQueue: processing session ' + sessId);
  Grader.gradeSession(sessId);
}

// ── ONE-TIME SETUP (run from Apps Script editor, not from web app) ───
// Creates the permanent every-1-minute trigger if it doesn't exist.
function setupGradingTrigger() {
  const existing = ScriptApp.getProjectTriggers();
  const alreadyExists = existing.some(t => t.getHandlerFunction() === 'checkGradeQueue');
  if (alreadyExists) {
    Logger.log('setupGradingTrigger: trigger already exists. Nothing to do.');
    return 'Trigger already installed.';
  }
  ScriptApp.newTrigger('checkGradeQueue')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('setupGradingTrigger: ✓ Trigger created successfully.');
  return 'Grading trigger installed. checkGradeQueue will run every 1 minute.';
}

// ── TOP-LEVEL WRAPPERS ─────────────────────────────────────────────────
// google.script.run can ONLY call top-level global functions — it cannot
// reach into object methods. These wrappers are the required bridge.
function startAIGrading(sessId)           { return Grader.startGradingAsync(sessId); }
function getGradingStatus(sessId)         { return Grader.getStatus(sessId); }
function overrideGradeScore(sessId, stuId, qId, score, fb) { return Grader.overrideScore(sessId, stuId, qId, score, fb); }
function regradeQuestion(sessId, qId, ctx){ return Grader.regradeWithContext(sessId, qId, ctx); }

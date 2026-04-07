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
  MODEL: 'gemini-3-flash-preview',
  BATCH_SIZE: 3,
  SYSTEM_INSTRUCTION: 'You are a rigorous, fair science teacher grading student short-answer responses. Your feedback style is direct, specific, and concise — like margin notes from an expert. Never restate the question. Never use filler phrases like "Great job" or "Good effort." Focus exclusively on the scientific accuracy of what the student wrote. Ignore spelling, grammar, and punctuation unless they change the factual meaning of a science concept. Only reference concepts the student actually wrote — never infer, assume, or fabricate content not present in the answer. If the answer is blank or nonsensical, score 0 and say so. Grade every student with equal care regardless of their position in the list. CRITICAL: You MUST read each student\'s answer EXACTLY as written — do NOT paraphrase, restate, or reinterpret the student\'s words. If a student wrote "reject the null hypothesis" do not change this to "fail to reject" or any other wording. Do NOT attribute one student\'s answer to another student under any circumstances. Every grade must reflect only what that specific student actually wrote. When grading, identify up to 3 key science concept labels (1-4 words each) that the student\'s answer addresses incorrectly or incompletely. Include these as a "concepts" array in your JSON output. If the answer is fully correct, return an empty concepts array.',
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
    // BUG 4 fix: prefer snapshot taken at launch; fall back to live QSet for pre-fix sessions
    const allQuestions = sess.snapshotQuestions || (qSet && qSet.questions);
    if (!allQuestions) {
      const msg = 'Question set not found.';
      this.setStatus(sessId, { state: 'error', sessionId: sessId, message: msg });
      return { error: msg };
    }

    const saQs = allQuestions.filter(q => q.type === 'sa');
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
      responseRowByKey[respData[i][0] + '|' + respData[i][1] + '|' + respData[i][3]] = { row: i + 1, maxPts: Number(respData[i][8]) || 0 };
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
      const pendingResps = qResps.filter(r => !done.has(r.studentId + '|' + q.id));
      const BATCH_SIZE = this.BATCH_SIZE;

      for (let i = 0; i < pendingResps.length; i += BATCH_SIZE) {
        if (Date.now() - startTime > TIMEOUT_LIMIT) {
          if (newGradeRows.length > 0) {
            this._batchAppendRows(gradeSheet, newGradeRows);
          }
          const out = { gradedCount: count, errors, totalToGrade, message: `Graded ${count}/${totalToGrade}. Time limit reached. Run again to finish.` };
          this.setStatus(sessId, Object.assign({ state: 'partial', sessionId: sessId }, out));
          return out;
        }

        const batch = pendingResps.slice(i, i + BATCH_SIZE);

        try {
          const results = this.callGeminiBatch(key, q, batch, '');

          const resultMap = {};
          results.forEach(res => {
            if (res && res.studentId) resultMap[res.studentId] = res;
          });

          for (const r of batch) {
            let finalResult = resultMap[r.studentId];
            if (!finalResult) {
              Logger.log('Batch missed ID ' + r.studentId + ', falling back to single');
              if (Date.now() - startTime > TIMEOUT_LIMIT) {
                if (newGradeRows.length > 0) {
                  this._batchAppendRows(gradeSheet, newGradeRows);
                }
                const out = { gradedCount: count, errors, totalToGrade, message: `Graded ${count}/${totalToGrade}. Time limit reached during fallback. Run again to finish.` };
                this.setStatus(sessId, Object.assign({ state: 'partial', sessionId: sessId }, out));
                return out;
              }
              try {
                const singleRes = this.callGemini(key, q, r.answer, '');
                finalResult = { studentId: r.studentId, score: singleRes.score, feedback: singleRes.feedback };
              } catch (fallbackErr) {
                Logger.log(`Failed to grade ${r.studentName} (fallback): ${fallbackErr.message}`);
                throw fallbackErr; // Trigger batch catch block
              }
            }

            newGradeRows.push([
              sessId, r.studentId, r.studentName, q.id,
              finalResult.score, q.points || 1, finalResult.feedback,
              r.answer, new Date().toISOString(),
              false, '', '', '', JSON.stringify(finalResult.concepts || []), ''
            ]);

            const responseEntry = responseRowByKey[sessId + '|' + r.studentId + '|' + q.id];
            if (responseEntry) {
              respSheet.getRange(responseEntry.row, 7, 1, 2).setValues([[
                responseEntry.maxPts > 0 && Number(finalResult.score) >= responseEntry.maxPts,
                finalResult.score
              ]]);
            }

            count++;
            done.add(r.studentId + '|' + q.id);
          }

          // Flush grade rows to AIGrades immediately after each batch so that AIGrades and
          // Responses stay in sync — if the job times out or throws after this point, any
          // already-scored responses will have a matching AIGrades record.
          if (newGradeRows.length > 0) {
            this._batchAppendRows(gradeSheet, newGradeRows);
            newGradeRows = [];
          }

          this.setStatus(sessId, { state: 'running', sessionId: sessId, gradedCount: count, errors, totalToGrade, message: `Grading... (${count}/${totalToGrade} done)` });
          Utilities.sleep(1500);

        } catch (e) {
          lastError = e.message || e.toString();
          Logger.log('Grade batch error on Q ' + q.id + ': ' + e.toString());

          Logger.log('Batch failed, falling back to individual grading for this batch.');
          for (const r of batch) {
             if (Date.now() - startTime > TIMEOUT_LIMIT) {
               if (newGradeRows.length > 0) {
                 this._batchAppendRows(gradeSheet, newGradeRows);
               }
               const out = { gradedCount: count, errors, totalToGrade, message: `Graded ${count}/${totalToGrade}. Time limit reached during fallback. Run again to finish.` };
               this.setStatus(sessId, Object.assign({ state: 'partial', sessionId: sessId }, out));
               return out;
             }

             const k = r.studentId + '|' + q.id;
             if (done.has(k)) continue;
             try {
                const singleRes = this.callGemini(key, q, r.answer, '');
                newGradeRows.push([
                  sessId, r.studentId, r.studentName, q.id,
                  singleRes.score, q.points || 1, singleRes.feedback,
                  r.answer, new Date().toISOString(),
                  false, '', '', ''
                ]);
                const responseEntry = responseRowByKey[sessId + '|' + r.studentId + '|' + q.id];
                if (responseEntry) {
                  respSheet.getRange(responseEntry.row, 7, 1, 2).setValues([[
                    responseEntry.maxPts > 0 && Number(singleRes.score) >= responseEntry.maxPts,
                    singleRes.score
                  ]]);
                }
                count++;
                done.add(k);
                this.setStatus(sessId, { state: 'running', sessionId: sessId, gradedCount: count, errors, totalToGrade, message: `Grading... (${count}/${totalToGrade} done)` });
                Utilities.sleep(1200);
             } catch (fallbackE) {
                lastError = fallbackE.message || fallbackE.toString();
                errors++;
             }
          }

          // Flush any buffered rows from the entire fallback loop at once
          if (newGradeRows.length > 0) {
            this._batchAppendRows(gradeSheet, newGradeRows);
            newGradeRows = [];
          }

          if (errors > 5) {
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

  callGeminiBatch(key, question, responses, extraContext) {
    if (!responses || responses.length === 0) return [];

    const maxPts = question.points || 1;
    let rubricText = '';
    if (question.rubric) rubricText += '\nRUBRIC:\n' + question.rubric;
    if (question.sampleAnswer) rubricText += '\nIDEAL ANSWER:\n' + question.sampleAnswer;
    if (extraContext) rubricText += '\nADDITIONAL TEACHER CONTEXT:\n' + extraContext;

    const total = responses.length;
    let studentsText = responses.map((r, idx) =>
      '━━━ STUDENT ' + (idx + 1) + ' OF ' + total + ' ━━━\nID: ' + r.studentId + '\nANSWER: "' + r.answer + '"'
    ).join('\n\n');

    const prompt = 'QUESTION (' + maxPts + ' pt' + (maxPts > 1 ? 's' : '') + '):\n' +
      question.text + '\n' +
      rubricText + '\n\n' +
      'STUDENT RESPONSES (' + total + ' total — grade EACH independently):\n' + studentsText + '\n\n' +
      'GRADING INSTRUCTIONS:\n' +
      '1. Compare each answer against the rubric and ideal answer above.\n' +
      '2. Award points only for correct scientific content that addresses the question.\n' +
      '3. For each student, write brief, specific feedback (1-3 sentences or fragments): name the exact concept correct, partially correct, or missing. No filler.\n' +
      '4. You MUST return results for ALL ' + total + ' student IDs listed above. Do NOT skip any student. Do NOT cut off mid-sentence.\n' +
      '5. CRITICAL: Grade each student ONLY on the EXACT words in their answer. Do NOT paraphrase or reword what the student wrote. Do NOT confuse one student\'s answer with another.\n\n' +
      'Reply with ONLY a JSON array (no markdown, no code fences, no commentary):\n' +
      '[{"id":"<student_id>","score":<0-' + maxPts + '>,"feedback":"<brief specific feedback>"}]';

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + this.MODEL + ':generateContent?key=' + key;

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: Grader.SYSTEM_INSTRUCTION }] },
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 8192 }
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
    let text = ((candidate.content && candidate.content.parts && candidate.content.parts[0].text) || '').trim();

    if (finishReason === 'MAX_TOKENS') {
      Logger.log('⚠ Gemini hit MAX_TOKENS for batch on: ' + question.text.substring(0, 60));
      const salvaged = this._salvageBatchJSON(text, maxPts);
      if (salvaged.length > 0) {
        Logger.log('Salvaged ' + salvaged.length + '/' + responses.length + ' results from truncated batch');
        return salvaged;
      }
      throw new Error('MAX_TOKENS: batch response fully truncated');
    }
    // Strip markdown fences
    text = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();

    try {
       // Attempt a direct parse or find the first array literal block
       let jsonMatch = text.match(/\[[\s\S]*\]/);
       let result = [];

       if (jsonMatch) {
         try {
           result = JSON.parse(jsonMatch[0]);
         } catch(innerE) {
           console.warn('Batch JSON parse failed on match, attempting raw parse fallback:', innerE);
           result = JSON.parse(text);
         }
       } else {
         result = JSON.parse(text);
       }

       if (!Array.isArray(result)) result = [result];

       return result.map(res => ({
         studentId: String(res.id),
         score: Math.min(Math.max(Number(res.score) || 0, 0), maxPts),
         feedback: res.feedback || 'No feedback generated.',
         concepts: Array.isArray(res.concepts) ? res.concepts.slice(0, 3) : []
       }));
    } catch(e) {
       console.error('Batch JSON parse completely failed:', e, 'Raw text:', text);
       throw new Error('Could not parse batch Gemini response: ' + text.substring(0, 200));
    }
  },

  callGemini(key, question, answer, extraContext) {
    const maxPts = question.points || 1;
    let rubricText = '';
    if (question.rubric) rubricText += '\nRUBRIC:\n' + question.rubric;
    if (question.sampleAnswer) rubricText += '\nIDEAL ANSWER:\n' + question.sampleAnswer;
    if (extraContext) rubricText += '\nADDITIONAL TEACHER CONTEXT:\n' + extraContext;

    const prompt = 'QUESTION (' + maxPts + ' pt' + (maxPts > 1 ? 's' : '') + '):\n' +
      question.text + '\n' +
      rubricText + '\n\n' +
      'STUDENT ANSWER:\n"' + answer + '"\n\n' +
      'GRADING INSTRUCTIONS:\n' +
      '1. Compare the answer against the rubric and ideal answer.\n' +
      '2. Award points only for correct scientific content.\n' +
      '3. Write brief, specific feedback (1-3 sentences or fragments): name the exact concept correct, partially correct, or missing.\n' +
      '4. CRITICAL: Grade ONLY based on the EXACT words the student wrote. Do NOT paraphrase or reword the student\'s answer.\n\n' +
      'Reply with ONLY this JSON (no markdown, no code fences):\n' +
      '{"score":<0-' + maxPts + '>,"feedback":"<brief specific feedback>"}';

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + this.MODEL + ':generateContent?key=' + key;

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: Grader.SYSTEM_INSTRUCTION }] },
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
      Logger.log('WARNING: Gemini hit MAX_TOKENS for single grading: ' + question.text.substring(0, 60) + ' — attempting partial extraction');
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
    const numScore = Number(score);
    if (!isFinite(numScore) || numScore < 0) {
      return { error: 'Score must be a non-negative number.' };
    }
    return DB.withLock(() => {
      const sheet = DB.sh('AIGrades');
      const d = sheet.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (d[i][0] === sessId && d[i][1] === stuId && d[i][3] === qId) {
          const maxPoints = Number(d[i][5]) || 0;
          if (maxPoints > 0 && numScore > maxPoints) {
            return { error: 'Score (' + numScore + ') exceeds max points (' + maxPoints + ').' };
          }
          const originalAIScore = d[i][4]; // col 5, index 4 — original AI score before override
          sheet.getRange(i + 1, 10).setValue(true);
          sheet.getRange(i + 1, 11).setValue(numScore);
          sheet.getRange(i + 1, 12).setValue(String(fb || '').slice(0, 2000));
          if (d[i][14] === '' || d[i][14] === undefined || d[i][14] === null) {
            sheet.getRange(i + 1, 15).setValue(originalAIScore); // OriginalAIScore — only set once
          }
          this._syncResponseScore(sessId, stuId, qId, numScore);
          DB.logAuditEvent('OVERRIDE_SCORE', sessId + ':' + stuId + ':' + qId, 'score=' + numScore);
          return { ok: true };
        }
      }
      return { error: 'Grade record not found — run AI grading first.' };
    });
  },

  revertToAIScore(sessId, stuId, qId) {
    return DB.withLock(() => {
      const sheet = DB.sh('AIGrades');
      const d = sheet.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (d[i][0] === sessId && d[i][1] === stuId && d[i][3] === qId) {
          const originalAIScore = d[i][14]; // col 15 — saved at override time
          if (originalAIScore === '' || originalAIScore === null || originalAIScore === undefined) {
            return { error: 'No original AI score saved — cannot revert.' };
          }
          const restoreScore = Number(originalAIScore);
          sheet.getRange(i + 1, 10).setValue(false); // overridden = false
          sheet.getRange(i + 1, 11).setValue('');    // overrideScore cleared
          sheet.getRange(i + 1, 12).setValue('');    // overrideFeedback cleared
          sheet.getRange(i + 1, 5).setValue(restoreScore); // restore original score
          this._syncResponseScore(sessId, stuId, qId, restoreScore);
          DB.logAuditEvent('REVERT_AI_SCORE', sessId + ':' + stuId + ':' + qId, 'restoredScore=' + restoreScore);
          return { ok: true, restoredScore: restoreScore };
        }
      }
      return { error: 'Grade record not found.' };
    });
  },

  regradeWithContext(sessId, qId, ctx) {
    // Cap the teacher-supplied context string to prevent API abuse and prompt injection.
    const safeCtx = String(ctx || '').slice(0, 500);
    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return { error: 'GEMINI_API_KEY not set.' };
    const sess = DB.getSessionById(sessId); if (!sess) return { error: 'Session not found' };
    const qSet = DB.getQSet(sess.setId); if (!qSet) return { error: 'Question set not found' };
    const q = (qSet.questions || []).find(x => x.id === qId); if (!q) return { error: 'Question not found' };
    const resps = DB.getAllResponses(sessId).filter(r => r.questionId === qId && r.answer && String(r.answer).length > 3);
    const gradeSheet = DB.sh('AIGrades');
    let updated = 0;
    let newGradeRows = [];

    const BATCH_SIZE = this.BATCH_SIZE;
    for (let i = 0; i < resps.length; i += BATCH_SIZE) {
      const batch = resps.slice(i, i + BATCH_SIZE);
      let resultMap = {};
      try {
        const results = this.callGeminiBatch(key, q, batch, safeCtx);
        results.forEach(res => {
          if (res && res.studentId) resultMap[res.studentId] = res;
        });

        // Fallback for missing items in batch map, before taking DB lock
        for (const r of batch) {
          if (!resultMap[r.studentId]) {
            try {
              resultMap[r.studentId] = this.callGemini(key, q, r.answer, safeCtx);
            } catch(e) { continue; }
          }
        }

        DB.withLock(() => {
          const rows = gradeSheet.getDataRange().getValues();

          // O(1) Index for existing grades in this session/question
          const rowIndexMap = new Map();
          for (let j = 1; j < rows.length; j++) {
            if (rows[j][0] === sessId && rows[j][3] === qId) {
              rowIndexMap.set(rows[j][1], j);
            }
          }

          for (const r of batch) {
            const result = resultMap[r.studentId];
            if (!result) continue; // skipped/failed

            const j = rowIndexMap.get(r.studentId);
            if (j !== undefined) {
              rows[j][4] = result.score;
              rows[j][6] = result.feedback;
              rows[j][12] = safeCtx;
              // Use rows[j].length for range width to match exactly what is in the row array
              gradeSheet.getRange(j + 1, 1, 1, rows[j].length).setValues([rows[j]]);
            } else {
              newGradeRows.push([sessId, r.studentId, r.studentName, qId, result.score, q.points || 1, result.feedback, r.answer, new Date().toISOString(), false, '', '', safeCtx]);
            }
          }
        });

        // Sync response scores outside DB lock for performance
        for (const r of batch) {
          const result = resultMap[r.studentId];
          if (result) {
            this._syncResponseScore(sessId, r.studentId, qId, result.score);
            updated++;
          }
        }
      } catch (e) {
         Logger.log('Regrade batch error: ' + e.toString());
         // Populate fallbacks for the entire batch
         for (const r of batch) {
           if (resultMap && resultMap[r.studentId]) continue;
           try {
             resultMap[r.studentId] = this.callGemini(key, q, r.answer, safeCtx);
           } catch (fallbackE) { Logger.log('Regrade single error: ' + fallbackE.toString()); }
         }

         DB.withLock(() => {
           const rows = gradeSheet.getDataRange().getValues();

           const rowIndexMap = new Map();
           for (let j = 1; j < rows.length; j++) {
             if (rows[j][0] === sessId && rows[j][3] === qId) {
               rowIndexMap.set(rows[j][1], j);
             }
           }

           for (const r of batch) {
             const result = resultMap[r.studentId];
             if (!result) continue; // Failed even fallback

             const j = rowIndexMap.get(r.studentId);
             if (j !== undefined) {
               rows[j][4] = result.score;
               rows[j][6] = result.feedback;
               rows[j][12] = safeCtx;
               gradeSheet.getRange(j + 1, 1, 1, rows[j].length).setValues([rows[j]]);
             } else {
               newGradeRows.push([sessId, r.studentId, r.studentName, qId, result.score, q.points || 1, result.feedback, r.answer, new Date().toISOString(), false, '', '', safeCtx]);
             }
           }
         });

         // Sync response scores
         for (const r of batch) {
           const result = resultMap[r.studentId];
           if (result) {
             this._syncResponseScore(sessId, r.studentId, qId, result.score);
             updated++;
           }
         }
      }
    }

    if (newGradeRows.length > 0) {
      this._batchAppendRows(gradeSheet, newGradeRows);
    }
    DB.logAuditEvent('REGRADE_WITH_CONTEXT', sessId + ':' + qId, 'ctx_length=' + safeCtx.length + ' updated=' + updated);
    const out = { ok: true, updated, message: 'Regraded ' + updated + ' responses.' };
    this.setStatus(sessId, Object.assign({ state: 'done', sessionId: sessId }, out));
    return out;
  },

  _salvageBatchJSON(text, maxPts) {
    // Extract complete JSON objects from a truncated batch response.
    // Each valid {id, score, feedback} entry is returned; missing student IDs
    // will naturally fall back to single grading via the existing resultMap check.
    const results = [];
    const regex = /\{\s*"id"\s*:\s*"([^"]+)"\s*,\s*"score"\s*:\s*(\d+(?:\.\d+)?)\s*,\s*"feedback"\s*:\s*"((?:[^"\\]|\\.)*)"\s*\}/g;
    let match;
    while ((match = regex.exec(text)) !== null) {
      results.push({
        studentId: match[1],
        score: Math.min(Math.max(Number(match[2]) || 0, 0), maxPts),
        feedback: match[3] || 'No feedback generated.'
      });
    }
    return results;
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

  // Dequeue the first session and save remaining back BEFORE processing so that
  // if the script times out mid-grading the entry is still removed (partial state
  // is recoverable by re-running grading from the UI).
  const sessId = queue.shift();
  props.setProperty(Grader.QUEUE_KEY, JSON.stringify(queue));

  Logger.log('checkGradeQueue: processing session ' + sessId);
  try {
    Grader.gradeSession(sessId);
  } catch (e) {
    Logger.log('checkGradeQueue: gradeSession threw for ' + sessId + ': ' + e.toString());
    // Re-queue the session for retry, up to a maximum of 3 attempts so a persistently
    // broken session does not loop forever.
    const status = Grader.getStatus(sessId);
    const attempts = (status.attempts || 0) + 1;
    const MAX_ATTEMPTS = 3;
    if (attempts <= MAX_ATTEMPTS) {
      let retryQueue;
      try { retryQueue = JSON.parse(props.getProperty(Grader.QUEUE_KEY) || '[]'); } catch(e2) { retryQueue = []; }
      if (!retryQueue.includes(sessId)) retryQueue.push(sessId);
      props.setProperty(Grader.QUEUE_KEY, JSON.stringify(retryQueue));
      Grader.setStatus(sessId, { state: 'queued', sessionId: sessId, attempts, message: 'Grading error — will retry automatically (attempt ' + attempts + '/' + MAX_ATTEMPTS + '). Error: ' + (e.message || e.toString()).substring(0, 200) });
      Logger.log('checkGradeQueue: re-queued ' + sessId + ' for retry (attempt ' + attempts + '/' + MAX_ATTEMPTS + ')');
    } else {
      Grader.setStatus(sessId, { state: 'error', sessionId: sessId, attempts, message: 'Grading failed after ' + MAX_ATTEMPTS + ' attempts. Last error: ' + (e.message || e.toString()).substring(0, 300) });
      Logger.log('checkGradeQueue: giving up on ' + sessId + ' after ' + MAX_ATTEMPTS + ' attempts');
    }
  }
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
// Returns {installed: bool} — called by TeacherApp before attempting to grade
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const installed = triggers.some(t => t.getHandlerFunction() === 'checkGradeQueue');
  return { installed };
}

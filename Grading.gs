// ═══════════════════════════════════════════════════════════════════
//  Grading.gs v5.3 — Features gemini-3-flash-preview & Timeout guards
// ═══════════════════════════════════════════════════════════════════

const Grader = {
  MODEL: 'gemini-3-flash-preview', 
  
  gradeSession(sessId) {
    const startTime = Date.now();
    const TIMEOUT_LIMIT = 5 * 60 * 1000;

    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return { error: 'GEMINI_API_KEY not set. Go to Project Settings > Script Properties.' };
    
    try { this.testKey(key); } catch(e) { return { error: 'API key test failed: ' + e.message }; }
    
    const sess = DB.getSessionById(sessId);
    if (!sess) return { error: 'Session not found' };
    const qSet = DB.getQSet(sess.setId);
    if (!qSet) return { error: 'Question set not found' };
    
    const saQs = qSet.questions.filter(q => q.type === 'sa');
    if (!saQs.length) return { message: 'No short-answer questions to grade.' };
    
    const resps = DB.getAllResponses(sessId);
    const existing = DB.getAIGrades(sessId);
    const gradeSheet = DB.sh('AIGrades');
    const respSheet = DB.sh('Responses');
    
    const done = new Set(existing.map(g => g.studentId + '|' + g.questionId));
    let count = 0, errors = 0, errorMsgs = [];
    
    for (const q of saQs) {
      const qResps = resps.filter(r => r.questionId === q.id && r.answer && String(r.answer).length > 3);
      
      for (const r of qResps) {
        if (Date.now() - startTime > TIMEOUT_LIMIT) {
          return { gradedCount: count, errors, message: `Graded ${count} responses. Execution time limit approaching. Run again to finish.` };
        }

        const k = r.studentId + '|' + q.id;
        if (done.has(k)) continue;
        
        try {
          const result = this.callGemini(key, q, r.answer, '');
          
          gradeSheet.appendRow([
            sessId, r.studentId, r.studentName, q.id,
            result.score, q.points || 1, result.feedback,
            r.answer, new Date().toISOString(),
            false, '', '', ''
          ]);
          
          const rd = respSheet.getDataRange().getValues();
          for (let i = 1; i < rd.length; i++) {
            if (rd[i][0] === sessId && rd[i][1] === r.studentId && rd[i][3] === q.id) {
              respSheet.getRange(i + 1, 8).setValue(result.score);
              break;
            }
          }
          
          count++;
          done.add(k);
          Utilities.sleep(1200);
          
        } catch (e) {
          Logger.log('Grade error ' + r.studentName + '/' + q.id + ': ' + e.toString());
          errorMsgs.push(r.studentName + ': ' + e.message);
          errors++;
          if (errors > 5) {
            return { gradedCount: count, errors, message: 'Stopped after ' + errors + ' errors. Last: ' + e.message };
          }
        }
      }
    }
    
    return { gradedCount: count, errors, message: 'Graded ' + count + ' responses' + (errors ? ' (' + errors + ' errors)' : '') + '.' };
  },
  
  testKey(key) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + this.MODEL + ':generateContent?key=' + key;
    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: 'Say hello' }] }], generationConfig: { maxOutputTokens: 10 } }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error('API returned status ' + code);
    return true;
  },
  
  callGemini(key, question, answer, extraContext) {
    const maxPts = question.points || 1;
    let rubricText = '';
    if (question.rubric) rubricText += '\nRUBRIC:\n' + question.rubric;
    if (question.sampleAnswer) rubricText += '\nIDEAL ANSWER:\n' + question.sampleAnswer;
    if (extraContext) rubricText += '\nADDITIONAL TEACHER CONTEXT:\n' + extraContext;
    
    const prompt = 'You are an expert science teacher grading a student response. Grade precisely against the rubric.\n\n' +
      'QUESTION (' + maxPts + ' point' + (maxPts > 1 ? 's' : '') + '):\n' + question.text + '\n' +
      rubricText + '\n\nSTUDENT ANSWER:\n"' + answer + '"\n\n' +
      'Return ONLY valid JSON with these fields:\n' +
      '{"score": <number 0 to ' + maxPts + '>, "feedback": "<2-3 sentences about what was right, wrong, and how to improve.>"}';

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + this.MODEL + ':generateContent?key=' + key;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.2, maxOutputTokens: 400, responseMimeType: 'application/json' }
      }),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) throw new Error('Gemini API error');
    if (!data.candidates || !data.candidates.length) throw new Error('No response candidates');
    
    let text = data.candidates[0].content.parts[0].text || '';
    text = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    
    let result = JSON.parse(text);
    return { score: Math.min(Math.max(result.score, 0), maxPts), feedback: result.feedback || 'No feedback generated.' };
  }
};

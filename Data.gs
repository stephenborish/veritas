// ═══════════════════════════════════════════════════════════════════
//  Data.gs v5.2 — Features LockService to prevent concurrency data loss
// ═══════════════════════════════════════════════════════════════════

const DB = {
  SS_NAME: 'Veritas Assess — Data',
  
  withLock(callback) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000); 
      return callback();
    } catch (e) {
      Logger.log("Lock exception: " + e.toString());
      return { error: 'System is busy processing other students. Please try again in a moment.' };
    } finally {
      lock.releaseLock();
    }
  },

  ss() {
    const id = PropertiesService.getScriptProperties().getProperty('VA_SHEET_ID');
    if (id) {
      try {
        return SpreadsheetApp.openById(id);
      } catch (e) {
        Logger.log('Error opening spreadsheet by ID: ' + e.toString());
      }
    }
    const f = DriveApp.getFilesByName(this.SS_NAME);
    if (f.hasNext()) { const s = SpreadsheetApp.open(f.next()); PropertiesService.getScriptProperties().setProperty('VA_SHEET_ID',s.getId()); return s; }
    throw new Error('Run initSystem() first.');
  },
  sh(n) { return this.ss().getSheetByName(n); },
  
  init() {
    let ss; const f = DriveApp.getFilesByName(this.SS_NAME);
    ss = f.hasNext() ? SpreadsheetApp.open(f.next()) : SpreadsheetApp.create(this.SS_NAME);
    PropertiesService.getScriptProperties().setProperty('VA_SHEET_ID', ss.getId());
    const sheets = {
      'Courses':['ID','Name','Blocks','CreatedAt'],
      'QSets':['ID','Name','CourseID','CreatedAt','UpdatedAt','QuestionsJSON','StimuliJSON'],
      'Rosters':['Block','CourseID','StudentsJSON','UpdatedAt'],
      'Sessions':['ID','Code','SetID','SetName','Block','Mode','RandQ','RandC','Status','CurrentQ','StartedAt','EndedAt','ConfigJSON','TimerJSON','RevealMode','SummaryConfigJSON','RevealedQs','CalcEnabled'],
      'Responses':['SessionID','StudentID','StudentName','QuestionID','QIndex','Answer','IsCorrect','Points','MaxPoints','SubmittedAt','PartialCredit'],
      'Metacognition':['SessionID','StudentID','StudentName','QuestionID','Confidence','SubmittedAt'],
      'StudentSessions':['SessionID','StudentID','StudentName','Email','Status','JoinedAt','FinishedAt','ViolationCount','LockedOut','QOrder','NeedsFullscreen', 'ClientToken', 'NormalizedName', 'IdentityKey'],
      'Violations':['SessionID','StudentID','StudentName','Type','Timestamp','Resolved'],
      'AIGrades':['SessionID','StudentID','StudentName','QuestionID','Score','MaxScore','Feedback','Answer','GradedAt','Overridden','OverrideScore','OverrideFeedback','Context'],
      'Archive':['SessionID','Code','SetName','Block','StartedAt','EndedAt','StudentCount','AvgPct','DataJSON']
    };
    for (const [name, headers] of Object.entries(sheets)) {
      let s = ss.getSheetByName(name);
      if (!s) s = ss.insertSheet(name);
      s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#0d7377').setFontColor('#fff');
      s.setFrozenRows(1);
    }
    const d = ss.getSheetByName('Sheet1');
    if (d && ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(d);
      } catch (e) {
        Logger.log('Error deleting default Sheet1: ' + e.toString());
      }
    }
    return {url:ss.getUrl()};
  },
  
  makeCode() {
    const w = ['PHOTON','NEURON','PROTON','ENZYME','GENOME','PLASMA','QUASAR','MITOSIS','ATOMIC','BORON','CARBON','HELIUM','KELVIN','NEUTRO','ORBITAL','REDOX','SOLUTE','TENSOR','VECTOR','VOLTAGE'];
    return w[Math.floor(Math.random()*w.length)] + String(Math.floor(Math.random()*90)+10);
  },
  
  // ── COURSES, QSETS, ROSTERS (Standard CRUD) ──
  createCourse(name, blocks) {
    const s = this.sh('Courses'); const id = 'crs_'+Utilities.getUuid().slice(0,8);
    s.appendRow([id, name, JSON.stringify(blocks||[]), new Date().toISOString()]);
    return {id, name, blocks};
  },
  getCourses() {
    const d = this.sh('Courses').getDataRange().getValues();
    return d.slice(1).map(r => ({id:r[0], name:r[1], blocks:JSON.parse(r[2]||'[]'), createdAt:r[3]}));
  },
  updateCourse(id, name, blocks) {
    const s=this.sh('Courses'); const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id){if(name!==undefined&&name!==null&&String(name).trim()!=='')s.getRange(i+1,2).setValue(name);if(blocks!==undefined)s.getRange(i+1,3).setValue(JSON.stringify(blocks));return true;}
    return false;
  },
  updateCourseBlocks(id, blocks) {
    return this.updateCourse(id, undefined, blocks);
  },
  deleteCourse(id) {
    const s=this.sh('Courses'); const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id){s.deleteRow(i+1);return true;}
    return false;
  },
  createQSet(name, courseId, questions, stimuli) {
    const s=this.sh('QSets'); const id='qs_'+Utilities.getUuid().slice(0,8); const now=new Date().toISOString();
    questions.forEach((q,i)=>{if(!q.id)q.id='q'+(i+1)+'_'+Date.now().toString(36)});
    s.appendRow([id,name,courseId||'',now,now,JSON.stringify(questions),JSON.stringify(stimuli||[])]);
    return {id,name,questionCount:questions.length};
  },
  getQSets(courseId) {
    const d=this.sh('QSets').getDataRange().getValues();
    return d.slice(1).filter(r=>!courseId||r[2]===courseId||!r[2]).map(r=>{
      const qs=JSON.parse(r[5]||'[]');
      return {id:r[0],name:r[1],courseId:r[2],createdAt:r[3],updatedAt:r[4],questionCount:qs.length,mcCount:qs.filter(q=>q.type==='mc').length,saCount:qs.filter(q=>q.type==='sa').length};
    });
  },
  getQSet(id) {
    const d=this.sh('QSets').getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id) return {id:d[i][0],name:d[i][1],courseId:d[i][2],questions:JSON.parse(d[i][5]||'[]'),stimuli:JSON.parse(d[i][6]||'[]')};
    return null;
  },
  updateQSet(id,name,courseId,questions,stimuli) {
    const s=this.sh('QSets'); const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id){
      s.getRange(i+1,2).setValue(name); s.getRange(i+1,3).setValue(courseId||'');
      s.getRange(i+1,5).setValue(new Date().toISOString());
      s.getRange(i+1,6).setValue(JSON.stringify(questions)); s.getRange(i+1,7).setValue(JSON.stringify(stimuli||[]));
      return {id,name};
    }
    return null;
  },
  deleteQSet(id) { const s=this.sh('QSets'); const d=s.getDataRange().getValues(); for(let i=1;i<d.length;i++) if(d[i][0]===id){s.deleteRow(i+1);return true;} return false; },
  saveRoster(block, courseId, students) {
    const s=this.sh('Rosters'); const d=s.getDataRange().getValues(); const now=new Date().toISOString();
    for(let i=1;i<d.length;i++) if(String(d[i][0])===String(block)){s.getRange(i+1,2).setValue(courseId||'');s.getRange(i+1,3).setValue(JSON.stringify(students));s.getRange(i+1,4).setValue(now);return {block,count:students.length};}
    s.appendRow([block,courseId||'',JSON.stringify(students),now]); return {block,count:students.length};
  },
  getRosters() {
    const d=this.sh('Rosters').getDataRange().getValues(); const r={};
    for(let i=1;i<d.length;i++){const stu=JSON.parse(d[i][2]||'[]');r[d[i][0]]={block:d[i][0],courseId:d[i][1],students:stu,count:stu.length,updatedAt:d[i][3]};}
    return r;
  },
  getRoster(block) { const d=this.sh('Rosters').getDataRange().getValues(); for(let i=1;i<d.length;i++) if(String(d[i][0])===String(block)) return JSON.parse(d[i][2]||'[]'); return []; },
  getRostersByCourse(courseId) {
    const out = this.getRosters();
    return Object.values(out).filter(r => String(r.courseId || '') === String(courseId || ''));
  },
  addStudent(block, student) {
    return this.withLock(() => {
      const roster = this.getRoster(block);
      const incomingName = String((student && student.name) || '').trim();
      if (!incomingName) return { error: 'Student name is required' };
      const normalized = this.normalizeStudentName(incomingName);
      const exists = roster.some(s => this.normalizeStudentName((s && s.name) || '') === normalized);
      if (exists) return { ok: true, added: false, count: roster.length };
      roster.push(student);
      const rosters = this.getRosters();
      const existing = rosters[String(block)] || { courseId: '' };
      this.saveRoster(block, existing.courseId || '', roster);
      return { ok: true, added: true, count: roster.length };
    });
  },
  removeStudent(block, name) {
    return this.withLock(() => {
      const roster = this.getRoster(block);
      const normalized = this.normalizeStudentName(name || '');
      const next = roster.filter(s => this.normalizeStudentName((s && s.name) || '') !== normalized);
      const removed = next.length !== roster.length;
      const rosters = this.getRosters();
      const existing = rosters[String(block)] || { courseId: '' };
      this.saveRoster(block, existing.courseId || '', next);
      return { ok: true, removed, count: next.length };
    });
  },
  
  // ── SESSIONS ──
  activateSession(config) {
    const s=this.sh('Sessions');
    const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][8]==='active'){s.getRange(i+1,9).setValue('ended');s.getRange(i+1,12).setValue(new Date().toISOString());}
    const qSet=this.getQSet(config.setId); if(!qSet) throw new Error('Question set not found');
    const id='sess_'+Utilities.getUuid().slice(0,8); const code=this.makeCode(); const now=new Date().toISOString();
    const mode = config.mode || 'self-paced';
    const initialQ = mode === 'lockstep' ? -1 : 0;
    s.appendRow([id,code,config.setId,qSet.name,config.block,mode,config.randomizeQuestions||false,config.randomizeChoices||false,'active',initialQ,now,'',JSON.stringify(config),JSON.stringify(config.timer||{type:'none'}),config.revealMode||'end',JSON.stringify(config.summaryConfig||{showScore:true}),'[]',config.calculatorEnabled||false]);
    return {sessionId:id,code,setName:qSet.name,block:config.block,mode:config.mode,questionCount:qSet.questions.length};
  },
  getActiveSession() {
    const d=this.sh('Sessions').getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][8]==='active') return this._parseSess(d[i],i+1);
    return null;
  },
  getSessionById(id) {
    const d=this.sh('Sessions').getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id) return this._parseSess(d[i],i+1);
    return null;
  },
  _parseSess(r,row) {
    return {sessionId:r[0],code:r[1],setId:r[2],setName:r[3],block:r[4],mode:r[5],randQ:r[6],randC:r[7],status:r[8],currentQ:r[9]||0,startedAt:r[10],endedAt:r[11],config:JSON.parse(r[12]||'{}'),timer:JSON.parse(r[13]||'{}'),revealMode:r[14]||'end',summaryConfig:JSON.parse(r[15]||'{}'),revealedQs:JSON.parse(r[16]||'[]'),calcEnabled:r[17]===true||r[17]==='TRUE',row};
  },
  _normalizeSessionState(sess, questionIds) {
    const ids = Array.isArray(questionIds) ? questionIds : [];
    const maxQ = Math.max(0, ids.length - 1);
    const requestedQ = Number(sess.currentQ);
    const minBound = sess.mode === 'lockstep' ? -1 : 0;
    const currentQ = Number.isFinite(requestedQ) ? Math.max(minBound, Math.min(requestedQ, maxQ)) : 0;
    const revealMode = sess.revealMode || 'end';
    let revealedQs = Array.isArray(sess.revealedQs) ? sess.revealedQs.filter(qId => ids.indexOf(qId) !== -1) : [];
    if (revealMode === 'never') revealedQs = [];
    if (revealMode === 'end' && sess.status === 'ended') revealedQs = ids.slice();
    const lockedQs = Array.isArray((sess.config || {}).lockedQs) ? sess.config.lockedQs : [];
    return {
      sessionId: sess.sessionId,
      sessionStatus: sess.status,
      mode: sess.mode,
      currentQ,
      timer: sess.timer,
      revealMode,
      revealedQs,
      lockedQs,
      calcEnabled: sess.calcEnabled,
      metacognitionEnabled: (sess.config || {}).metacognitionEnabled !== false
    };
  },
  endSession(id) {
    return this.withLock(() => {
      const s=this.sh('Sessions'); const d=s.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]===id){
        s.getRange(i+1,9).setValue('ended');s.getRange(i+1,12).setValue(new Date().toISOString());this.archiveSession(id);
        const ss=this.sh('StudentSessions'); const sd=ss.getDataRange().getValues();
        for(let j=1;j<sd.length;j++) if(sd[j][0]===id&&(sd[j][8]===true||sd[j][8]==='TRUE')){ss.getRange(j+1,9).setValue(false);}
        return true;
      }
      return false;
    });
  },
  
  // ── CONCURRENT SECURE STUDENT JOIN ──
  studentJoin(code, first, last, clientToken, studentToken) {
    return this.withLock(() => {
      const sess=this.getActiveSession(); if(!sess) return {error:'No active session.'};
      if(sess.code!==code.toUpperCase().trim()) return {error:'Invalid code.'};
      
      const name=(first.trim()+' '+last.trim()).trim();
      const normalizedName=this.normalizeStudentName(name);
      const roster=this.getRoster(sess.block);
      const accessToken = normalizeStudentToken(studentToken);
      const tokenData = accessToken ? verifyStudentAccessToken(accessToken) : null;
      if (accessToken && !tokenData) return {error:'Your secure link is invalid. Please reopen the email from your teacher.'};
      if (tokenData && String(tokenData.sid || '') !== String(sess.sessionId)) return {error:'This secure link is for a different session.'};
      if (tokenData && tokenData.code && tokenData.code !== String(sess.code || '')) return {error:'This secure link does not match the active session code.'};
      if (!normalizedName) return {error:'Please enter your full name.'};

      const matchedRoster=(roster||[]).filter(r=>this.normalizeStudentName((r&&r.name)||'')===normalizedName);
      let rosterEntry=null;

      if(tokenData){
        const tokenEmail=this.normalizeStudentEmail(tokenData.email || '');
        const tokenName=this.normalizeStudentName(tokenData.normalizedName || tokenData.name || [tokenData.firstName, tokenData.lastName].join(' '));
        rosterEntry=(roster||[]).find(r=>{
          const rosterEmail=this.normalizeStudentEmail((r&&r.email)||'');
          const rosterName=this.normalizeStudentName((r&&r.name)||'');
          if(tokenEmail && rosterEmail) return rosterEmail===tokenEmail;
          return !!tokenName && rosterName===tokenName;
        }) || null;
        if(!rosterEntry) return {error:'This secure link no longer matches the roster for this session.'};
        if(this.normalizeStudentName((rosterEntry&&rosterEntry.name)||'') !== normalizedName){
          return {error:'Please enter your name exactly as it appears in your teacher\'s email link.'};
        }
      } else {
        if(!matchedRoster.length) return {error:'Your name is not on the roster for this session.'};
        if(matchedRoster.length > 1) return {error:'Multiple students share this name. Please use your personalized email link.'};
        rosterEntry=matchedRoster[0];
        if(this.normalizeStudentEmail((rosterEntry&&rosterEntry.email)||'')){
          return {error:'Please use the secure assessment link from your email to join this session.'};
        }
      }

      const rosterEmail=this.normalizeStudentEmail((rosterEntry&&rosterEntry.email)||'');
      const storedName=String((rosterEntry&&rosterEntry.name)||name).trim();
      const storedNormalizedName=this.normalizeStudentName(storedName);
      const rosterKey=rosterEmail?('email:'+rosterEmail):('name:'+storedNormalizedName+'#1');
      
      const ssSheet=this.sh('StudentSessions'); const sd=ssSheet.getDataRange().getValues();
      const deriveIdentityKey=(row)=>{
        const existing=String(row[13]||'');
        if(existing) return existing;
        const rowEmail=this.normalizeStudentEmail(row[3]||'');
        if(rowEmail) return 'sess:'+row[0]+'|email:'+rowEmail;
        const rowNorm=this.normalizeStudentName(row[12]||row[2]||'');
        return 'sess:'+row[0]+'|name:'+rowNorm+'#1';
      };
      const identityKey='sess:'+sess.sessionId+'|'+rosterKey;
      
      for(let i=1;i<sd.length;i++) {
        const rowIdentity=deriveIdentityKey(sd[i]);
        if(sd[i][0]===sess.sessionId && rowIdentity===identityKey) {
          if(sd[i][11] && clientToken !== sd[i][11]) return {error:'This name is already in use by another device.'};
          if(sd[i][8]===true||sd[i][8]==='TRUE') return {error:'Locked out. Wait for teacher.'};
          
          const needsFS = sd[i][10]===true||sd[i][10]==='TRUE';
          const qSet=this.getQSet(sess.setId);
          return {sessionId:sess.sessionId,studentId:sd[i][1],studentName:sd[i][2],mode:sess.mode,questionCount:qSet?qSet.questions.length:0,rejoined:true,calcEnabled:sess.calcEnabled,timer:sess.timer,revealMode:sess.revealMode,needsFullscreen:needsFS,metacognitionEnabled:sess.config.metacognitionEnabled!==false};
        }
      }

      const stuId='stu_'+Utilities.getUuid().replace(/-/g,'').slice(0,12);
      
      // New Join
      let qOrder='';
      if((sess.randQ||sess.mode==='randomized') && sess.mode !== 'lockstep'){
        const qSet=this.getQSet(sess.setId);
        if(qSet){
          const idx=qSet.questions.map((_,i)=>i);
          for(let i=idx.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[idx[i],idx[j]]=[idx[j],idx[i]];}
          qOrder=JSON.stringify(idx);
        }
      }
      ssSheet.appendRow([sess.sessionId, stuId, storedName, rosterEmail, 'active', new Date().toISOString(), '', 0, false, qOrder, false, clientToken, storedNormalizedName, identityKey]);
      const qSet=this.getQSet(sess.setId);
      return {sessionId:sess.sessionId,studentId:stuId,studentName:storedName,mode:sess.mode,questionCount:qSet?qSet.questions.length:0,rejoined:false,calcEnabled:sess.calcEnabled,timer:sess.timer,revealMode:sess.revealMode,needsFullscreen:false,metacognitionEnabled:sess.config.metacognitionEnabled!==false};
    });
  },

  studentGetQuestions(sessId,stuId) {
    const sess=this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
    const qSet=this.getQSet(sess.setId); if(!qSet) return {error:'Questions not found'};
    let questions=JSON.parse(JSON.stringify(qSet.questions)); const stimuli=qSet.stimuli||[];
    const sd=this.sh('StudentSessions').getDataRange().getValues(); let qOrder=null;
    for(let i=1;i<sd.length;i++) {
      if(sd[i][0]===sessId&&sd[i][1]===stuId&&sd[i][9]){
        try {
          qOrder = JSON.parse(sd[i][9]);
        } catch (e) {
          Logger.log('Error parsing qOrder for student ' + stuId + ': ' + e.toString());
        }
        break;
      }
    }
    if(qOrder && qOrder.length && sess.mode !== 'lockstep') questions=qOrder.map(idx=>questions[idx]);
    if(sess.randC) questions.forEach(q=>{if(q.type==='mc'&&q.choices){const ci=q.correctIndices||[q.correctIndex];const ca=ci.map(x=>q.choices[x]);for(let i=q.choices.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[q.choices[i],q.choices[j]]=[q.choices[j],q.choices[i]];}q.correctIndices=ca.map(c=>q.choices.indexOf(c));if(q.correctIndices.length===1)q.correctIndex=q.correctIndices[0];}});
    const studentQs=questions.map(q=>{const sq={...q};
      // Mark multi-answer questions BEFORE stripping correctIndices
      if(q.type==='mc'){const ci=q.correctIndices||[q.correctIndex];sq.multiSelect=Array.isArray(ci)&&ci.length>1;}
      delete sq.correctIndex;delete sq.correctIndices;delete sq.correctAnswer;delete sq.rubric;delete sq.sampleAnswer;return sq;});
    const resps=this.getAllResponses(sessId).filter(r=>r.studentId===stuId);
    const existing={};resps.forEach(r=>{existing[r.questionId]=r.answer;});
    const metaResps=this.getAllMeta(sessId).filter(m=>m.studentId===stuId);
    const existingMeta={};metaResps.forEach(m=>{existingMeta[m.questionId]=m.confidence;});
    const sessionState = this._normalizeSessionState(sess, studentQs.map(q=>q.id));
    return {questions:studentQs,stimuli,existing,existingMeta,...sessionState};
  },
  
  // ── CONCURRENT SECURE SUBMISSION ──
  studentSubmitAnswer(sessId, stuId, qId, answer) {
    return this.withLock(() => {
      const sess=this.getSessionById(sessId); if(!sess||sess.status!=='active') return {error:'Session ended'};
      const qSet=this.getQSet(sess.setId); const q=qSet.questions.find(qq=>qq.id===qId); if(!q) return {error:'Q not found'};
      const lockedQs = Array.isArray((sess.config || {}).lockedQs) ? sess.config.lockedQs : [];
      if (lockedQs.includes(qId)) return {error:'This question has been locked by the teacher.'};
      const stuName=this.getStudentName(sessId,stuId); const maxPts=q.points||1;
      let isCorrect=null,points=0,partialCredit=false;
      if(q.type==='mc'){
        const stripHtml = s => String(s || '').replace(/<[^>]*>/g, '').replace(/&[^;]+;/g, ' ').replace(/\s+/g, '').toLowerCase();
        const ci=q.correctIndices||[q.correctIndex]; 
        const ca=ci.map(x=>q.choices[x]);
        const sel=Array.isArray(answer)?answer:[answer];
        
        let correctCount = 0;
        let incorrectCount = 0;
        
        sel.forEach(ans => {
           let cleanAns = stripHtml(ans);
           let matched = false;
           for (let i = 0; i < ca.length; i++) {
              let cleanKey = stripHtml(ca[i]);
              if (cleanAns === cleanKey || (cleanKey.length > 2 && cleanAns.includes(cleanKey)) || (cleanKey.length > 2 && cleanKey.includes(cleanAns))) {
                 matched = true;
                 break;
              }
           }
           if (matched) correctCount++;
           else incorrectCount++;
        });
        
        if (ca.length === 1) {
           isCorrect = (correctCount === 1 && incorrectCount === 0);
           points = isCorrect ? maxPts : 0;
        } else {
           isCorrect = (correctCount === ca.length && incorrectCount === 0);
           points = isCorrect ? maxPts : 0;
           partialCredit = false;
        }
      } else if (q.type === 'sa') {
          points = 0;
      }
      const qIdx=qSet.questions.findIndex(qq=>qq.id===qId);
      let ansStr=Array.isArray(answer)?JSON.stringify(answer):String(answer);
      // Force plain text in Sheets to prevent auto-formatting (e.g. 100% -> 1)
      if (ansStr && !ansStr.startsWith('{') && !ansStr.startsWith('[')) {
        ansStr = "'" + ansStr;
      }
      
      const rSheet=this.sh('Responses');
      const rd=rSheet.getDataRange().getValues();
      
      for(let i=1;i<rd.length;i++) {
        if(rd[i][0]===sessId && rd[i][1]===stuId && rd[i][3]===qId){
          rSheet.getRange(i+1,6).setNumberFormat('@').setValue(ansStr);rSheet.getRange(i+1,7).setValue(isCorrect);rSheet.getRange(i+1,8).setValue(points);rSheet.getRange(i+1,10).setValue(new Date().toISOString());rSheet.getRange(i+1,11).setValue(partialCredit);
          let ci2=null;if(sess.revealMode==='immediate')ci2=this._correctInfo(q);
          return {saved:true,isCorrect,points,maxPts,partialCredit,correctInfo:ci2};
        }
      }
      const newRow = rSheet.getLastRow() + 1;
      rSheet.getRange(newRow, 6).setNumberFormat('@');
      rSheet.appendRow([sessId,stuId,stuName,qId,qIdx,ansStr,isCorrect,points,maxPts,new Date().toISOString(),partialCredit]);
      let ci3=null;if(sess.revealMode==='immediate')ci3=this._correctInfo(q);
      return {saved:true,isCorrect,points,maxPts,partialCredit,correctInfo:ci3};
    });
  },
  _correctInfo(q){if(q.type==='mc'){const ci=q.correctIndices||[q.correctIndex];return {correctAnswers:ci.map(i=>q.choices[i]),correctIndices:ci};}return null;},
  
  studentSubmitMeta(sessId,stuId,qId,conf) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId);
      if (sess) {
         const lockedQs = Array.isArray((sess.config || {}).lockedQs) ? sess.config.lockedQs : [];
         if (lockedQs.includes(qId)) return {error: 'This question has been locked by the teacher.'};
      }
      const name=this.getStudentName(sessId,stuId);
      const s=this.sh('Metacognition');const d=s.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]===sessId&&d[i][1]===stuId&&d[i][3]===qId){s.getRange(i+1,5).setValue(conf);s.getRange(i+1,6).setValue(new Date().toISOString());return true;}
      s.appendRow([sessId,stuId,name,qId,conf,new Date().toISOString()]);return true;
    });
  },
  
  studentReportViolation(sessId,stuId,type) {
    return this.withLock(() => {
      const name=this.getStudentName(sessId,stuId);
      this.sh('Violations').appendRow([sessId,stuId,name,type,new Date().toISOString(),false]);
      const s=this.sh('StudentSessions');const d=s.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]===sessId&&d[i][1]===stuId){
        s.getRange(i+1,8).setValue((d[i][7]||0)+1);
        s.getRange(i+1,9).setValue(true); // Lock out
        s.getRange(i+1,11).setValue(true); // Needs fullscreen on re-admit
        break;
      }
      return {lockedOut:true};
    });
  },
  
  readmitStudent(sessId, stuId) {
    return this.withLock(() => {
      const s=this.sh('StudentSessions');const d=s.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]===sessId&&d[i][1]===stuId){
        s.getRange(i+1,9).setValue(false); // Unlock
        s.getRange(i+1,11).setValue(true); // Require fullscreen re-entry
        break;
      }
      // Resolve ALL violations for this student
      const v=this.sh('Violations');const vd=v.getDataRange().getValues();
      for(let i=1;i<vd.length;i++) if(vd[i][0]===sessId&&vd[i][1]===stuId&&!vd[i][5]){
        v.getRange(i+1,6).setValue(true);
      }
      return {readmitted:true};
    });
  },

  studentCheckStatus(sessId,stuId) {
    const sess=this.getSessionById(sessId); if(!sess) return {sessionStatus:'ended'};
    const qSet=this.getQSet(sess.setId);
    const validIds = qSet ? qSet.questions.map(q=>q.id) : [];
    const sessionState = this._normalizeSessionState(sess, validIds);
    
    // Attach details for revealed questions
    if (sess.revealedQs && sess.revealedQs.length > 0 && qSet) {
       sessionState.revealedDetails = {};
       const responses = this.getAllResponses(sessId).filter(r => r.studentId === stuId);
       const grades = this.getAIGrades(sessId).filter(g => g.studentId === stuId);
       
       sess.revealedQs.forEach(qId => {
          const q = qSet.questions.find(x => x.id === qId);
          if (q) {
             const r = responses.find(x => x.questionId === qId);
             const g = grades.find(x => x.questionId === qId);
             const ci = q.correctIndices || [q.correctIndex];
             sessionState.revealedDetails[qId] = {
                clientCorrectAnswer: q.type === 'mc' ? ci.map(idx => q.choices[idx]).join(', ') : null,
                clientRubric: q.rubric || '',
                clientSampleAnswer: q.sampleAnswer || '',
                aiScore: g ? g.score : null,
                aiFeedback: g ? g.feedback : null,
                ptsEarned: r ? (Number(r.points) || 0) : 0
             };
          }
       });
    }

    const d=this.sh('StudentSessions').getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===sessId&&d[i][1]===stuId){
      return {...sessionState,lockedOut:d[i][8]===true||d[i][8]==='TRUE',
        needsFullscreen:d[i][10]===true||d[i][10]==='TRUE'};
    }
    return {sessionStatus:'not-found'};
  },

  studentFinish(sessId,stuId) {
    return this.withLock(() => {
      const sess=this.getSessionById(sessId);
      const s=this.sh('StudentSessions');const d=s.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]===sessId&&d[i][1]===stuId){
        s.getRange(i+1,5).setValue('finished');
        s.getRange(i+1,7).setValue(new Date().toISOString());
        const responses=this.getAllResponses(sessId).filter(r=>r.studentId===stuId);
        return this.buildStudentSummary(sess, responses);
      }
      return this.buildStudentSummary(sess, []);
    });
  },

  // Read functions do not strictly require locks for performance reasons.
  getLiveResults(sessId) {
    const sess=this.getSessionById(sessId);if(!sess)return null;
    const qSet=this.getQSet(sess.setId);const questions=qSet?qSet.questions:[];
    const snapshot={};
    const stuSess=this.getStudentSessions(sessId,snapshot);const resps=this.getAllResponses(sessId,snapshot);
    const viols=this.getActiveViolations(sessId,snapshot);const meta=this.getAllMeta(sessId,snapshot);const roster=this.getRoster(sess.block);
    const grades=this.getAIGrades(sessId,snapshot);

    const questionById={};
    questions.forEach(q=>{questionById[q.id]=q;});
    const responsesByStudent={};
    const responsesByQuestion={};
    resps.forEach(r=>{
      if(!responsesByStudent[r.studentId]) responsesByStudent[r.studentId]=[];
      responsesByStudent[r.studentId].push(r);
      if(!responsesByQuestion[r.questionId]) responsesByQuestion[r.questionId]=[];
      responsesByQuestion[r.questionId].push(r);
    });
    const metaByStudent={};
    const metaByQuestion={};
    const metaByStudentQuestion={};
    meta.forEach(m=>{
      if(!metaByStudent[m.studentId]) metaByStudent[m.studentId]=[];
      metaByStudent[m.studentId].push(m);
      if(!metaByQuestion[m.questionId]) metaByQuestion[m.questionId]=[];
      metaByQuestion[m.questionId].push(m);
      metaByStudentQuestion[m.studentId+'|'+m.questionId]=m;
    });
    const gradeByStudentQuestion={};
    grades.forEach(g=>{gradeByStudentQuestion[g.studentId+'|'+g.questionId]=g;});
    
    const nowMs=Date.now();
    const students=stuSess.map(ss=>{
      const sr=responsesByStudent[ss.studentId]||[];
      const sm=metaByStudent[ss.studentId]||[];
      let mcC=0,mcT=0;
      sr.forEach(r=>{const q=questionById[r.questionId];if(q&&q.type==='mc'){mcT++;if(r.isCorrect===true||r.isCorrect==='TRUE')mcC++;}});
      const lastTs=[ss.joinedAt,ss.finishedAt].concat(sr.map(x=>x.submittedAt)).concat(sm.map(x=>x.submittedAt)).filter(Boolean).map(x=>new Date(x).getTime()).sort((a,b)=>b-a)[0]||0;
      const activeNow=ss.status==='active' && !ss.lockedOut && (nowMs-lastTs)<(2*60*1000);
      return {studentId:ss.studentId,name:ss.studentName,status:ss.status,activeNow,answered:sr.length,total:questions.length,mcCorrect:mcC,mcTotal:mcT,lockedOut:ss.lockedOut,violationCount:ss.violationCount,avgConf:sm.length>0?(sm.reduce((s,m)=>s+m.confidence,0)/sm.length).toFixed(1):null,responses:sr};
    });
    
    const qStats=questions.map((q,idx)=>{
      const qr=responsesByQuestion[q.id]||[];
      const qm=metaByQuestion[q.id]||[];
      if(q.type==='mc'){
        const correct=qr.filter(r=>r.isCorrect===true||r.isCorrect==='TRUE').length;
        const dist={};(q.choices||[]).forEach(c=>dist[c]=0);
        qr.forEach(r => {
          try {
            const a = JSON.parse(r.answer);
            if (Array.isArray(a)) {
              a.forEach(x => { if (dist[x] !== undefined) dist[x]++; });
            } else if (dist[r.answer] !== undefined) {
              dist[r.answer]++;
            }
          } catch (e) {
            // Not JSON, treat as plain string
            if (dist[r.answer] !== undefined) dist[r.answer]++;
          }
        });
        const ci=q.correctIndices||[q.correctIndex];
        return {id:q.id,idx,text:q.text,type:'mc',points:q.points||1,choices:q.choices||[],correctIndices:ci,total:qr.length,correct,pct:qr.length>0?Math.round((correct/qr.length)*100):0,dist,avgConf:qm.length>0?(qm.reduce((s,m)=>s+m.confidence,0)/qm.length).toFixed(1):null,
          studentResponses:qr.map(r=>({studentId:r.studentId,name:r.studentName,answer:r.answer,isCorrect:r.isCorrect,confidence:(metaByStudentQuestion[r.studentId+'|'+q.id]||{}).confidence||null}))};
      }else{
        return {id:q.id,idx,text:q.text,type:'sa',points:q.points||1,total:qr.length,
          studentResponses:qr.map(r=>{
            const g=gradeByStudentQuestion[r.studentId+'|'+q.id];
            const m=metaByStudentQuestion[r.studentId+'|'+q.id];
            // aiScore from AIGrades (g.score); r.points is written back to Responses sheet by grader
            const aiScore = g ? (g.overridden ? g.overrideScore : g.score) : null;
            return {studentId:r.studentId,name:r.studentName,answer:r.answer, score: aiScore, feedback: g ? g.feedback : null, overridden: g ? !!g.overridden : false, confidence:m?m.confidence:null};
          })};
      }
    });
    
    const joinedRosterKeys=new Set(stuSess.filter(s=>s.identityKey&&String(s.identityKey).indexOf('|email:')>-1).map(s=>s.identityKey));
    const joinedNameCounts={};
    stuSess.filter(s=>!s.identityKey||String(s.identityKey).indexOf('|email:')===-1).forEach(s=>{
      const n=s.normalizedName||this.normalizeStudentName(s.studentName||'');
      joinedNameCounts[n]=(joinedNameCounts[n]||0)+1;
    });
    const missing=[];
    const seenNameCounts={};
    (roster||[]).forEach(r=>{
      const rn=this.normalizeStudentName((r&&r.name)||'');
      const email=(r&&r.email)?String(r.email).trim().toLowerCase():'';
      if(email){
        const rosterIdentity='sess:'+sess.sessionId+'|email:'+email;
        if(joinedRosterKeys.has(rosterIdentity)) return;
      }else{
        seenNameCounts[rn]=(seenNameCounts[rn]||0)+1;
        if((joinedNameCounts[rn]||0)>=seenNameCounts[rn]) return;
      }
      missing.push({
        studentId: null, name: (r&&r.name)||'', email: (r&&r.email)||'', status: 'not-joined',
        answered: 0, total: questions.length, mcCorrect: 0, mcTotal: 0,
        lockedOut: false, violationCount: 0, avgConf: null, responses: []
      });
    });
    
    return {session:{...sess,questionCount:questions.length},students:[...students,...missing],qStats,violations:viols,totalJoined:students.length,totalFinished:students.filter(s=>s.status==='finished').length,totalActiveNow:students.filter(s=>s.activeNow).length,rosterSize:roster.length,missingCount:missing.length};
  },


  getSessionHistory() {
    const d=this.sh('Archive').getDataRange().getValues();
    return d.slice(1).map(r=>({sessionId:r[0],code:r[1],setName:r[2],block:r[3],startedAt:r[4],endedAt:r[5],studentCount:r[6],avgPct:r[7]})).reverse();
  },
  regenerateCode(id) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const code=this.makeCode();
      this.sh('Sessions').getRange(sess.row,2).setValue(code);
      return {sessionId:id,code};
    });
  },
  setSessionCode(id, code) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const c=String(code||'').toUpperCase().replace(/[^A-Z0-9]/g,'').slice(0,20);
      if(!c) return {error:'Invalid code'};
      this.sh('Sessions').getRange(sess.row,2).setValue(c);
      return {sessionId:id,code:c};
    });
  },
  goToQuestion(id, qIndex) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const qSet=this.getQSet(sess.setId); if(!qSet) return {error:'Question set not found'};
      const maxQ = Math.max(0, qSet.questions.length - 1);
      const q=Math.max(0,Math.min(Number(qIndex)||0,maxQ));
      this.sh('Sessions').getRange(sess.row,10).setValue(q);
      const updatedSess = this.getSessionById(id);
      return {ok:true,session:this._normalizeSessionState(updatedSess,qSet.questions.map(qq=>qq.id))};
    });
  },

  toggleLockQuestion(id, qId) {
    return this.withLock(() => {
      const sess = this.getSessionById(id); 
      if (!sess) return {error:'Session not found'};
      
      const cfg = sess.config || {};
      const lockedQs = Array.isArray(cfg.lockedQs) ? cfg.lockedQs : [];
      
      const idx = lockedQs.indexOf(qId);
      if (idx > -1) {
        lockedQs.splice(idx, 1);
      } else {
        lockedQs.push(qId);
      }
      
      cfg.lockedQs = lockedQs;
      this.sh('Sessions').getRange(sess.row, 13).setValue(JSON.stringify(cfg));
      
      const updatedSess = this.getSessionById(id);
      const qSet = this.getQSet(sess.setId);
      return {ok: true, session: this._normalizeSessionState(updatedSess, qSet ? qSet.questions.map(q=>q.id) : [])};
    });
  },
  advanceQuestion(id) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const qSet=this.getQSet(sess.setId); if(!qSet) return {error:'Question set not found'};
      const maxQ = Math.max(0, qSet.questions.length - 1);
      const nextQ = Math.max(0, Math.min((Number(sess.currentQ)||0) + 1, maxQ));
      this.sh('Sessions').getRange(sess.row,10).setValue(nextQ);
      const updatedSess = this.getSessionById(id);
      return {ok:true,session:this._normalizeSessionState(updatedSess,qSet.questions.map(q=>q.id))};
    });
  },
  revealAnswer(id, qId) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const qSet=this.getQSet(sess.setId); if(!qSet) return {error:'Question set not found'};
      const validIds = (qSet.questions||[]).map(q=>q.id);
      if (validIds.indexOf(qId) === -1) return {error:'Question not found'};
      const cur = Array.isArray(sess.revealedQs) ? sess.revealedQs.filter(id2 => validIds.indexOf(id2) !== -1) : [];
      if (!cur.includes(qId)) cur.push(qId);
      this.sh('Sessions').getRange(sess.row,17).setValue(JSON.stringify(cur));
      const updatedSess = this.getSessionById(id);
      return {ok:true,session:this._normalizeSessionState(updatedSess,validIds)};
    });
  },
  revealAllAnswers(id) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const qSet=this.getQSet(sess.setId); if(!qSet) return {error:'Question set not found'};
      const all=(qSet.questions||[]).map(q=>q.id);
      this.sh('Sessions').getRange(sess.row,17).setValue(JSON.stringify(all));
      const updatedSess = this.getSessionById(id);
      return {ok:true,session:this._normalizeSessionState(updatedSess,all)};
    });
  },
  setTimer(id, config) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      this.sh('Sessions').getRange(sess.row,14).setValue(JSON.stringify(config||{type:'none'}));
      return {ok:true,timer:config||{type:'none'}};
    });
  },
  updateSessionConfig(id, key, val) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const cfg=sess.config||{}; cfg[key]=val;
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      return {ok:true,config:cfg};
    });
  },
  updateSummaryConfig(id, cfg) {
    return this.withLock(() => {
      const sess=this.getSessionById(id); if(!sess) return {error:'Session not found'};
      const merged=Object.assign({},sess.summaryConfig||{},cfg||{});
      this.sh('Sessions').getRange(sess.row,16).setValue(JSON.stringify(merged));
      return {ok:true,summaryConfig:merged};
    });
  },
  studentGetSummary(sessId, stuId) {
    const sess=this.getSessionById(sessId);
    const responses=sess?this.getAllResponses(sessId).filter(r=>r.studentId===stuId):[];
    return this.buildStudentSummary(sess, responses);
  },
  buildStudentSummary(sess, responses) {
    const cfg=sess?(sess.summaryConfig||{}):{showScore:false};
    const safeResponses=Array.isArray(responses)?responses:[];
    const totalPts=safeResponses.reduce((sum,r)=>sum+(Number(r.points)||0),0);
    const totalMax=safeResponses.reduce((sum,r)=>sum+(Number(r.maxPoints)||0),0);
    const pct=totalMax>0?Math.round((totalPts/totalMax)*100):0;
    return {
      showScore:cfg.showScore!==false,
      pct,
      totalPts,
      totalMax,
      responses:safeResponses.length,
      revealMode:(sess&&sess.revealMode)||'end'
    };
  },
  getLiveQuestionDetail(sessId, qId) {
    const live=this.getLiveResults(sessId); if(!live) return null;
    return (live.qStats||[]).find(q=>q.id===qId)||null;
  },
  getItemAnalysis(id) {
    const live=this.getLiveResults(id); if(!live) return [];
    return live.qStats||[];
  },
  getStudentAnalysis(id) {
    const live=this.getLiveResults(id); if(!live) return [];
    return live.students||[];
  },
  getMetacognitionData(id) {
    return this.getAllMeta(id);
  },
  getStudentDetail(sessId, stuId) {
    const sess=this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
    const qSet=this.getQSet(sess.setId)||{questions:[]};
    const responses=this.getAllResponses(sessId).filter(r=>r.studentId===stuId);
    const meta=this.getAllMeta(sessId).filter(m=>m.studentId===stuId);
    const grades=this.getAIGrades(sessId).filter(g=>g.studentId===stuId);
    const student=this.getStudentSessions(sessId).find(s=>s.studentId===stuId);
    const details=(qSet.questions||[]).map((q,idx)=>{
      const r=responses.find(x=>x.questionId===q.id)||null;
      const m=meta.find(x=>x.questionId===q.id)||null;
      const g=grades.find(x=>x.questionId===q.id)||null;
      return {questionId:q.id,qIndex:idx,questionText:q.text,type:q.type,answer:r?r.answer:'',isCorrect:r?r.isCorrect:null,points:r?Number(r.points)||0:0,maxPoints:q.points||1,confidence:m?m.confidence:null,aiScore:g?g.score:null,aiFeedback:g?g.feedback:null};
    });
    const totalPts=details.reduce((a,b)=>a+(Number(b.points)||0),0);
    const totalMax=details.reduce((a,b)=>a+(Number(b.maxPoints)||0),0);
    return {student,session:{sessionId:sessId,setName:sess.setName,code:sess.code},totalPts,totalMax,pct:totalMax?Math.round((totalPts/totalMax)*100):0,details};
  },
  archiveSession(id) {
    const sess=this.getSessionById(id); if(!sess) return false;
    const archive=this.sh('Archive');
    const existing=archive.getDataRange().getValues().slice(1).find(r=>r[0]===id);
    if(existing) return true;
    const students=this.getStudentSessions(id);
    const responses=this.getAllResponses(id);
    const totalByStudent={};
    students.forEach(st=>{const sr=responses.filter(r=>r.studentId===st.studentId); const pts=sr.reduce((a,r)=>a+(Number(r.points)||0),0); const max=sr.reduce((a,r)=>a+(Number(r.maxPoints)||0),0); totalByStudent[st.studentId]=max?Math.round((pts/max)*100):0;});
    const pcts=Object.values(totalByStudent);
    const avgPct=pcts.length?Math.round(pcts.reduce((a,b)=>a+b,0)/pcts.length):0;
    archive.appendRow([id,sess.code,sess.setName,sess.block,sess.startedAt,sess.endedAt||new Date().toISOString(),students.length,avgPct,JSON.stringify({session:sess,students,responses,meta:this.getAllMeta(id),grades:this.getAIGrades(id),violations:this.getActiveViolations(id)})]);
    return true;
  },

  getArchivedSessions() {
    const archive = this.sh('Archive');
    if (!archive) return [];
    const values = archive.getDataRange().getValues();
    if (values.length <= 1) return [];
    
    return values.slice(1).map(r => ({
      id: r[0],
      code: r[1],
      setName: r[2],
      block: r[3],
      startedAt: r[4],
      endedAt: r[5],
      studentCount: r[6],
      avgPct: r[7]
    })).sort((a,b) => new Date(b.startedAt).getTime() - new Date(a.startedAt).getTime());
  },

  getArchivedSessionData(sessionId) {
    const archive = this.sh('Archive');
    if (!archive) return null;
    const values = archive.getDataRange().getValues();
    const row = values.slice(1).find(r => r[0] === sessionId);
    if (!row || !row[8]) return null;
    try {
      const data = JSON.parse(row[8]);
      if (!data.session) data.session = {};
      data.session.id = sessionId; // Ensure id is always present
      if (!data.session.config) data.session.config = {};
      
      // Legacy support: inject missing qSet
      if (!data.session.config.qSet && !data.session.qSet && data.session.setId) {
        data.session.config.qSet = this.getQSet(data.session.setId) || {};
      }
      return data;
    } catch(e) {
      return null;
    }
  },

  saveAIClassReport(sessionId, reportMarkdown) {
    const archive = this.sh('Archive');
    if (!archive) return false;
    const values = archive.getDataRange().getValues();
    const rowIdx = values.slice(1).findIndex(r => r[0] === sessionId);
    if (rowIdx === -1) return false;
    
    const rowNum = rowIdx + 2; // +1 for header, +1 for 0-index
    try {
      const data = JSON.parse(values[rowNum - 1][8]);
      data.aiReport = reportMarkdown;
      archive.getRange(rowNum, 9).setValue(JSON.stringify(data));
      return true;
    } catch (e) {
      return false;
    }
  },

  rescoreQuestionFull(sessionId, questionId, qJsonStr) {
    return this.withLock(() => {
      let globalUpdate = false;
      const qNew = JSON.parse(qJsonStr);
      qNew.points = Number(qNew.points) || 1;

      // 1. UPDATE LIVE DATABASES (Sessions & Responses)
      try {
        const respSh = this.sh('Responses');
        if (respSh) {
          const respRows = respSh.getDataRange().getValues();
          const stripHtml = s => String(s || '').replace(/<[^>]*>/g, '').replace(/&[^;]+;/g, ' ').replace(/\s+/g, '').toLowerCase();
          
          for (let i = 1; i < respRows.length; i++) {
            if (respRows[i][0] === sessionId && respRows[i][3] === questionId) {
              const rowNum = i + 1;
              const rawAns = String(respRows[i][5]);
              let ansToMatch = rawAns;
              try {
                const p = JSON.parse(rawAns);
                if (Array.isArray(p)) ansToMatch = p.join(' ');
              } catch (e) {
                // If it's not JSON, we use the raw string. No need to log excessively here as it's common.
              }
              ansToMatch = stripHtml(ansToMatch);
              let pts = 0; let isCorrect = false;

              if (qNew.type === 'mc') {
                const correctIndices = qNew.correctIndices || [qNew.correctIndex || 0];
                const matchedIdx = (qNew.choices || []).findIndex(ch => {
                  const c = stripHtml(ch);
                  return c === ansToMatch || (c.length > 2 && ansToMatch.includes(c)) || (c.length > 2 && c.includes(ansToMatch));
                });
                // Fallback: percentage-conversion (Sheets converts "+100%" → 1)
                let finalIdx = matchedIdx;
                if (finalIdx === -1) {
                  const numAns = parseFloat(ansToMatch);
                  if (!isNaN(numAns)) {
                    const pctVal = Math.round(numAns * 100);
                    const candidates = [String(pctVal)+'%',(pctVal>0?'+':'')+String(pctVal)+'%',String(pctVal),(pctVal>0?'+':'')+String(pctVal)].map(s=>stripHtml(s));
                    finalIdx = (qNew.choices||[]).findIndex(ch => { const c=stripHtml(ch); return candidates.some(cd=>c===cd); });
                    if (finalIdx === -1) {
                      finalIdx = (qNew.choices||[]).findIndex(ch => { const c=stripHtml(ch); if(c.length>20) return false; const ns=c.replace(/[^0-9.\-+]/g,''); if(!ns) return false; return Math.abs(parseFloat(ns)-pctVal)<0.5; });
                    }
                  }
                }
                if (finalIdx !== -1 && correctIndices.includes(finalIdx)) {
                  pts = qNew.points; isCorrect = true;
                }
              }
              // SA rescoring requires an AI call so we wipe its grading
              respSh.getRange(rowNum, 7, 1, 3).setValues([[isCorrect, pts, qNew.points]]);
              globalUpdate = true;
            }
          }
        }

        const sessSh = this.sh('Sessions');
        if (sessSh) {
          const sessRows = sessSh.getDataRange().getValues();
          for (let i = 1; i < sessRows.length; i++) {
            if (sessRows[i][0] === sessionId) {
              const cfg = JSON.parse(sessRows[i][12] || '{}');
              if (cfg.qSet && cfg.qSet.questions) {
                const qIdx = cfg.qSet.questions.findIndex(q => q.id === questionId);
                if (qIdx !== -1) {
                  Object.assign(cfg.qSet.questions[qIdx], qNew);
                  sessSh.getRange(i + 1, 13).setValue(JSON.stringify(cfg));
                  globalUpdate = true;
                }
              }
            }
          }
        }
      } catch(e) { /* Live update failed, continue to archive */ }

      // 2. UPDATE ARCHIVE DATABASE
      const archive = this.sh('Archive');
      if (!archive) return globalUpdate;
      const values = archive.getDataRange().getValues();
      const rowIdx = values.slice(1).findIndex(r => r[0] === sessionId);
      if (rowIdx === -1) return globalUpdate;
      
      const rowNum = rowIdx + 2; 
      try {
        const data = JSON.parse(values[rowNum - 1][8]);
        const qSet = Object.assign({}, (data.session.config && data.session.config.qSet) ? data.session.config.qSet : data.session.qSet);
        if (!qSet || !qSet.questions) return globalUpdate;
        
        const qOriginal = qSet.questions.find(qx => qx.id === questionId);
        if (!qOriginal) return globalUpdate;
        
        Object.assign(qOriginal, qNew);
        
        if (qOriginal.type === 'mc') {
          const stripHtml = s => String(s || '').replace(/<[^>]*>/g, '').replace(/&[^;]+;/g, ' ').replace(/\s+/g, '').toLowerCase();
          data.responses.forEach(r => {
             if (r.questionId === questionId) {
                let ansToMatch = String(r.answer || '');
                try {
                  const p = JSON.parse(ansToMatch);
                  if (Array.isArray(p)) ansToMatch = p.join(' ');
                } catch (e) {
                  // Not JSON, use raw
                }
                ansToMatch = stripHtml(ansToMatch);
                const correctIndices = qOriginal.correctIndices || [qOriginal.correctIndex || 0];
                const matchedIdx = (qOriginal.choices || []).findIndex(ch => {
                  const c = stripHtml(ch);
                  return c === ansToMatch || (c.length > 2 && ansToMatch.includes(c)) || (c.length > 2 && c.includes(ansToMatch));
                });
                // Fallback: percentage-conversion (Sheets converts "+100%" → 1)
                let finalIdx = matchedIdx;
                if (finalIdx === -1) {
                  const numAns = parseFloat(ansToMatch);
                  if (!isNaN(numAns)) {
                    const pctVal = Math.round(numAns * 100);
                    const candidates = [String(pctVal)+'%',(pctVal>0?'+':'')+String(pctVal)+'%',String(pctVal),(pctVal>0?'+':'')+String(pctVal)].map(s=>stripHtml(s));
                    finalIdx = (qOriginal.choices||[]).findIndex(ch => { const c=stripHtml(ch); return candidates.some(cd=>c===cd); });
                    if (finalIdx === -1) {
                      finalIdx = (qOriginal.choices||[]).findIndex(ch => { const c=stripHtml(ch); if(c.length>20) return false; const ns=c.replace(/[^0-9.\-+]/g,''); if(!ns) return false; return Math.abs(parseFloat(ns)-pctVal)<0.5; });
                    }
                  }
                }
                
                if (finalIdx !== -1 && correctIndices.includes(finalIdx)) {
                   r.points = qOriginal.points; 
                   r.isCorrect = true;
                } else {
                   r.points = 0;
                   r.isCorrect = false;
                }
                r.maxPoints = qOriginal.points;
             }
          });
        } else {
           data.responses.forEach(r => {
             if (r.questionId === questionId) {
               r.points = 0;
               r.isCorrect = false;
               r.maxPoints = qOriginal.points;
               r.score = 0;
               r.feedback = '';
             }
           });
           if (data.grades) {
             data.grades = data.grades.filter(g => g.questionId !== questionId);
           }
        }
        
        const students = data.students || [];
        const responses = data.responses || [];
        const totalByStudent={};
        students.forEach(st=>{
          const sr=responses.filter(r=>r.studentId===st.studentId);
          const pts=sr.reduce((a,r)=>a+(Number(r.points)||0),0); 
          const max=sr.reduce((a,r)=>a+(Number(r.maxPoints)||0),0); 
          totalByStudent[st.studentId]=max?Math.round((pts/max)*100):0;
        });
        const pcts=Object.values(totalByStudent);
        const avgPct=pcts.length?Math.round(pcts.reduce((a,b)=>a+b,0)/pcts.length):0;
        
        if (data.session.config && data.session.config.qSet) data.session.config.qSet = qSet;
        else data.session.qSet = qSet;
        
        archive.getRange(rowNum, 8).setValue(avgPct); 
        archive.getRange(rowNum, 9).setValue(JSON.stringify(data)); 
        
        return true;
      } catch (e) {
        return false;
      }
    });
  },

  dismissViolation(sessionId, timestamp) {
    const archive = this.sh('Archive');
    const values = archive.getDataRange().getValues();
    const rowIdx = values.slice(1).findIndex(r => r[0] === sessionId);
    if (rowIdx === -1) return false;
    
    const rowNum = rowIdx + 2; 
    try {
      const data = JSON.parse(values[rowNum - 1][8]);
      if (data.violations) {
          data.violations = data.violations.filter(v => String(v.timestamp) !== String(timestamp));
          archive.getRange(rowNum, 9).setValue(JSON.stringify(data));
      }
      
      const violSheet = this.sh('Violations');
      const violData = violSheet.getDataRange().getValues();
      for (let i=1; i<violData.length; i++) {
         if (violData[i][0] === sessionId && String(violData[i][4]) === String(timestamp)) {
             violSheet.deleteRow(i+1);
             break;
         }
      }
      return true;
    } catch (e) {
      return false;
    }
  },

  // Helpers
  getStudentName(sessId,stuId){const d=this.sh('StudentSessions').getDataRange().getValues();for(let i=1;i<d.length;i++)if(d[i][0]===sessId&&d[i][1]===stuId)return d[i][2];return 'Unknown';},
  _getSheetValues(sheetName,snapshot){
    if(snapshot&&snapshot[sheetName]) return snapshot[sheetName];
    const values=this.sh(sheetName).getDataRange().getValues();
    if(snapshot) snapshot[sheetName]=values;
    return values;
  },
  _getSessionRows(sheetName,sessId,snapshot){
    const d=this._getSheetValues(sheetName,snapshot);
    return d.slice(1).filter(r=>r[0]===sessId);
  },
  getStudentSessions(sessId,snapshot){const rows=this._getSessionRows('StudentSessions',sessId,snapshot);return rows.map(r=>({studentId:r[1],studentName:r[2],email:r[3],status:r[4],joinedAt:r[5],finishedAt:r[6],violationCount:r[7]||0,lockedOut:r[8]===true||r[8]==='TRUE',qOrder:r[9],needsFS:r[10]===true||r[10]==='TRUE',clientToken:r[11],normalizedName:r[12]||this.normalizeStudentName(r[2]||''),identityKey:r[13]||''}));},
  getAllResponses(sessId,snapshot){const rows=this._getSessionRows('Responses',sessId,snapshot);return rows.map(r=>({studentId:r[1],studentName:r[2],questionId:r[3],qIndex:r[4],answer:r[5],isCorrect:r[6],points:r[7],maxPoints:r[8],submittedAt:r[9],partialCredit:r[10]}));},
  getAllMeta(sessId,snapshot){const rows=this._getSessionRows('Metacognition',sessId,snapshot);return rows.map(r=>({studentId:r[1],studentName:r[2],questionId:r[3],confidence:r[4],submittedAt:r[5]}));},
  getActiveViolations(sessId,snapshot){const rows=this._getSessionRows('Violations',sessId,snapshot);return rows.map(r=>({studentId:r[1],studentName:r[2],type:r[3],timestamp:r[4],resolved:r[5]===true||r[5]==='TRUE'}));},
  getAIGrades(sessId,snapshot){const rows=this._getSessionRows('AIGrades',sessId,snapshot);return rows.map(r=>({studentId:r[1],studentName:r[2],questionId:r[3],score:r[4],maxScore:r[5],feedback:r[6],answer:r[7],gradedAt:r[8],overridden:r[9]===true||r[9]==='TRUE',overrideScore:r[10],overrideFeedback:r[11],context:r[12]}));},
  normalizeStudentName(name){return String(name||'').toLowerCase().replace(/\s+/g,' ').trim();},
  normalizeStudentEmail(email){return String(email||'').trim().toLowerCase();}
};

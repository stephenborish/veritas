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
      'Archive':['SessionID','Code','SetName','Block','StartedAt','EndedAt','StudentCount','AvgPct','DataJSON'],
      'AuditLog':['Timestamp','Action','Target','Details']
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

  // Returns an error string if name is invalid, or null if valid.
  validateName_(name) {
    const s = String(name || '').trim();
    if (!s) return 'Name cannot be empty.';
    if (s.length > 200) return 'Name too long (max 200 characters).';
    return null;
  },

  // Appends a row to the AuditLog sheet. Fire-and-forget — no lock needed for
  // append-only operations (Sheets appends are atomic at the row level).
  logAuditEvent(action, target, details) {
    try {
      const sheet = this.sh('AuditLog');
      if (!sheet) return;
      sheet.appendRow([new Date().toISOString(), String(action || ''), String(target || ''), String(details || '')]);
    } catch (e) {
      Logger.log('logAuditEvent error: ' + e.toString());
    }
  },

  // ── COURSES, QSETS, ROSTERS (Standard CRUD) ──
  createCourse(name, blocks) {
    const nameErr = this.validateName_(name);
    if (nameErr) return { error: nameErr };
    const safeName = String(name).trim();
    const s = this.sh('Courses'); const id = 'crs_'+Utilities.getUuid().slice(0,8);
    s.appendRow([id, safeName, JSON.stringify(blocks||[]), new Date().toISOString()]);
    return {id, name: safeName, blocks};
  },
  getCourses() {
    const d = this.sh('Courses').getDataRange().getValues();
    return d.slice(1).map(r => ({id:r[0], name:r[1], blocks:JSON.parse(r[2]||'[]'), createdAt:r[3]}));
  },
  updateCourse(id, name, blocks) {
    if (name !== undefined && name !== null) {
      const nameErr = this.validateName_(name);
      if (nameErr) return { error: nameErr };
    }
    const s=this.sh('Courses'); const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id){if(name!==undefined&&name!==null&&String(name).trim()!=='')s.getRange(i+1,2).setValue(String(name).trim());if(blocks!==undefined)s.getRange(i+1,3).setValue(JSON.stringify(blocks));return true;}
    return false;
  },
  updateCourseBlocks(id, blocks) {
    return this.updateCourse(id, undefined, blocks);
  },
  deleteCourse(id) {
    const s=this.sh('Courses'); const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id){s.deleteRow(i+1);this.logAuditEvent('DELETE_COURSE', id, '');return true;}
    return false;
  },
  createQSet(name, courseId, questions, stimuli) {
    const nameErr = this.validateName_(name);
    if (nameErr) return { error: nameErr };
    if (!Array.isArray(questions)) return { error: 'Questions must be an array.' };
    if (questions.length > 100) return { error: 'Question set cannot exceed 100 questions.' };
    const safeName = String(name).trim();
    const s=this.sh('QSets'); const id='qs_'+Utilities.getUuid().slice(0,8); const now=new Date().toISOString();
    questions.forEach((q,i)=>{if(!q.id)q.id='q'+(i+1)+'_'+Date.now().toString(36)});
    s.appendRow([id,safeName,courseId||'',now,now,JSON.stringify(questions),JSON.stringify(stimuli||[])]);
    return {id,name:safeName,questionCount:questions.length};
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
    const nameErr = this.validateName_(name);
    if (nameErr) return { error: nameErr };
    if (!Array.isArray(questions)) return { error: 'Questions must be an array.' };
    if (questions.length > 100) return { error: 'Question set cannot exceed 100 questions.' };
    const safeName = String(name).trim();
    const s=this.sh('QSets'); const d=s.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(d[i][0]===id){
      s.getRange(i+1,2).setValue(safeName); s.getRange(i+1,3).setValue(courseId||'');
      s.getRange(i+1,5).setValue(new Date().toISOString());
      s.getRange(i+1,6).setValue(JSON.stringify(questions)); s.getRange(i+1,7).setValue(JSON.stringify(stimuli||[]));
      return {id,name:safeName};
    }
    return null;
  },
  deleteQSet(id) { const s=this.sh('QSets'); const d=s.getDataRange().getValues(); for(let i=1;i<d.length;i++) if(d[i][0]===id){s.deleteRow(i+1);this.logAuditEvent('DELETE_QSET', id, '');return true;} return false; },
  saveRoster(block, courseId, students) {
    const s=this.sh('Rosters'); const d=s.getDataRange().getValues(); const now=new Date().toISOString();
    for(let i=1;i<d.length;i++) if(String(d[i][0])===String(block) && String(d[i][1])===String(courseId||'')){s.getRange(i+1,3).setValue(JSON.stringify(students));s.getRange(i+1,4).setValue(now);return {block,count:students.length};}
    s.appendRow([block,courseId||'',JSON.stringify(students),now]); return {block,count:students.length};
  },
  getRosters() {
    const d=this.sh('Rosters').getDataRange().getValues(); const r={};
    for(let i=1;i<d.length;i++){
      const block=d[i][0],courseId=d[i][1],rawJSON=d[i][2]||'[]',updatedAt=d[i][3];
      const key = courseId ? courseId + '_' + block : block;
      let cached=null;
      Object.defineProperty(r,key,{
        get:()=>{if(!cached){const stu=JSON.parse(rawJSON);cached={block,courseId,students:stu,count:stu.length,updatedAt};}return cached;},
        enumerable:true,
        configurable:true
      });
    }
    return r;
  },
  getRoster(block, courseId) {
    const d=this.sh('Rosters').getDataRange().getValues();
    // Match both block and courseId. Treat undefined/null courseId as empty string.
    const targetCourseId = courseId || '';
    for(let i=1;i<d.length;i++) {
      if(String(d[i][0])===String(block) && String(d[i][1])===String(targetCourseId)) {
        return JSON.parse(d[i][2]||'[]');
      }
    }
    return [];
  },
  getRostersByCourse(courseId) {
    const out = this.getRosters();
    return Object.values(out).filter(r => String(r.courseId || '') === String(courseId || ''));
  },
  addStudent(block, student, courseId) {
    return this.withLock(() => {
      const roster = this.getRoster(block, courseId);
      const incomingName = String((student && student.name) || '').trim();
      if (!incomingName) return { error: 'Student name is required' };
      const normalized = this.normalizeStudentName(incomingName);
      const exists = roster.some(s => this.normalizeStudentName((s && s.name) || '') === normalized);
      if (exists) return { ok: true, added: false, count: roster.length };
      roster.push(student);
      const rosters = this.getRosters();
      const key = courseId ? courseId + '_' + block : block;
      const existing = rosters[key] || { courseId: courseId || '' };
      this.saveRoster(block, existing.courseId || '', roster);
      return { ok: true, added: true, count: roster.length };
    });
  },
  removeStudent(block, name, courseId) {
    return this.withLock(() => {
      const roster = this.getRoster(block, courseId);
      const normalized = this.normalizeStudentName(name || '');
      const next = roster.filter(s => this.normalizeStudentName((s && s.name) || '') !== normalized);
      const removed = next.length !== roster.length;
      const rosters = this.getRosters();
      const key = courseId ? courseId + '_' + block : block;
      const existing = rosters[key] || { courseId: courseId || '' };
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
    const initialQ = -1;
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
    const minBound = -1;
    const currentQ = Number.isFinite(requestedQ) ? Math.max(minBound, Math.min(requestedQ, maxQ)) : 0;
    const revealMode = sess.revealMode || 'end';
    let revealedQs = Array.isArray(sess.revealedQs) ? sess.revealedQs.filter(qId => ids.indexOf(qId) !== -1) : [];
    if (revealMode === 'never') revealedQs = [];
    if (revealMode === 'end' && sess.status === 'ended') revealedQs = ids.slice();
    const lockedQs = Array.isArray((sess.config || {}).lockedQs) ? sess.config.lockedQs : [];
    const cfg = sess.config || {};
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
      metacognitionEnabled: cfg.metacognitionEnabled !== false,
      qTimerSeconds: cfg.qTimerSeconds || 0,
      qTimerState:   cfg.qTimerState   || null,
      closeAt:       cfg.closeAt       || null
    };
  },
  endSession(id) {
    return this.withLock(() => {
      const s=this.sh('Sessions'); const d=s.getDataRange().getValues();
      for(let i=1;i<d.length;i++) if(d[i][0]===id){
        s.getRange(i+1,9).setValue('ended');s.getRange(i+1,12).setValue(new Date().toISOString());this.archiveSession(id);
        const ss=this.sh('StudentSessions'); const sd=ss.getDataRange().getValues();
        for(let j=1;j<sd.length;j++) if(sd[j][0]===id&&(sd[j][8]===true||sd[j][8]==='TRUE')){ss.getRange(j+1,9).setValue(false);}
        this.logAuditEvent('END_SESSION', id, '');
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

      // Enforce access window
      const _now = new Date();
      const _openAt  = (sess.config||{}).openAt  ? new Date(sess.config.openAt)  : null;
      const _closeAt = (sess.config||{}).closeAt ? new Date(sess.config.closeAt) : null;
      if (_openAt  && _now < _openAt)  return {error: 'This assessment has not opened yet. It opens at ' + _openAt.toLocaleString() + '.'};
      if (_closeAt && _now > _closeAt) return {error: 'This assessment is closed.'};

      const name=(first.trim()+' '+last.trim()).trim();
      const normalizedName=this.normalizeStudentName(name);
      const courseId=(sess.config && sess.config.courseId) || '';
      const roster=this.getRoster(sess.block, courseId);
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
      const _newJoinQSet=this.getQSet(sess.setId);
      if((sess.randQ||sess.mode==='randomized') && sess.mode !== 'lockstep'){
        if(_newJoinQSet){
          const idx=_newJoinQSet.questions.map((_,i)=>i);
          for(let i=idx.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[idx[i],idx[j]]=[idx[j],idx[i]];}
          qOrder=JSON.stringify(idx);
        }
      }
      // Generate per-student shuffled choice order for all MC questions (stable across page reloads)
      let choiceOrders='';
      if(_newJoinQSet){
        const co={};
        _newJoinQSet.questions.forEach(q=>{
          if(q.type==='mc'&&q.choices&&q.choices.length>1){
            const idx=q.choices.map((_,i)=>i);
            for(let i=idx.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[idx[i],idx[j]]=[idx[j],idx[i]];}
            co[q.id]=idx;
          }
        });
        choiceOrders=JSON.stringify(co);
      }
      ssSheet.appendRow([sess.sessionId, stuId, storedName, rosterEmail, 'active', new Date().toISOString(), '', 0, false, qOrder, false, clientToken, storedNormalizedName, identityKey, choiceOrders]);
      return {sessionId:sess.sessionId,studentId:stuId,studentName:storedName,mode:sess.mode,questionCount:_newJoinQSet?_newJoinQSet.questions.length:0,rejoined:false,calcEnabled:sess.calcEnabled,timer:sess.timer,revealMode:sess.revealMode,needsFullscreen:false,metacognitionEnabled:sess.config.metacognitionEnabled!==false};
    });
  },

  studentGetQuestions(sessId,stuId) {
    const sess=this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
    const qSet=this.getQSet(sess.setId); if(!qSet) return {error:'Questions not found'};
    let questions=JSON.parse(JSON.stringify(qSet.questions)); const stimuli=qSet.stimuli||[];
    const ssSheet=this.sh('StudentSessions'); const sd=ssSheet.getDataRange().getValues(); let qOrder=null; let choiceOrders=null; let stuRowIdx=-1;
    for(let i=1;i<sd.length;i++) {
      if(sd[i][0]===sessId&&sd[i][1]===stuId){
        stuRowIdx=i+1;
        if(sd[i][9]){try{qOrder=JSON.parse(sd[i][9]);}catch(e){Logger.log('Error parsing qOrder for student '+stuId+': '+e.toString());}}
        if(sd[i][14]){try{choiceOrders=JSON.parse(sd[i][14]);}catch(e){Logger.log('Error parsing choiceOrders for student '+stuId+': '+e.toString());}}
        break;
      }
    }
    // If no choiceOrders yet (student joined before this feature), generate and persist now
    if(!choiceOrders && stuRowIdx>0){
      const co={};
      qSet.questions.forEach(q=>{
        if(q.type==='mc'&&q.choices&&q.choices.length>1){
          const idx=q.choices.map((_,i)=>i);
          for(let i=idx.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[idx[i],idx[j]]=[idx[j],idx[i]];}
          co[q.id]=idx;
        }
      });
      choiceOrders=co;
      try{ssSheet.getRange(stuRowIdx,15).setValue(JSON.stringify(co));}catch(e){Logger.log('Error saving choiceOrders: '+e.toString());}
    }
    if(qOrder && qOrder.length && sess.mode !== 'lockstep') questions=qOrder.map(idx=>questions[idx]);
    // Apply per-student persisted choice order (generated once at join time for consistency across reloads)
    if(choiceOrders) questions.forEach(q=>{if(q.type==='mc'&&q.choices&&choiceOrders[q.id]){const order=choiceOrders[q.id];const orig=q.choices.slice();const origCi=q.correctIndices||[q.correctIndex];const correctTexts=origCi.map(i=>orig[i]);q.choices=order.map(i=>orig[i]);q.correctIndices=correctTexts.map(ca=>q.choices.indexOf(ca));if(q.correctIndices.length===1)q.correctIndex=q.correctIndices[0];}});
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
          rSheet.getRange(i+1,6).setNumberFormat('@');
          rSheet.getRange(i+1,6,1,6).setValues([[ansStr, isCorrect, points, rd[i][8], new Date().toISOString(), partialCredit]]);
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
      this.logAuditEvent('READMIT_STUDENT', sessId, 'stuId=' + stuId);
      return {readmitted:true};
    });
  },

  studentCheckStatus(sessId,stuId) {
    const sess=this.getSessionById(sessId); if(!sess) return {sessionStatus:'ended'};
    // Enforce access window close time — treat expired window as ended session
    const _closeAt = (sess.config||{}).closeAt ? new Date(sess.config.closeAt) : null;
    if (_closeAt && new Date() > _closeAt) return {sessionStatus:'ended'};
    const qSet=this.getQSet(sess.setId);
    const validIds = qSet ? qSet.questions.map(q=>q.id) : [];
    const sessionState = this._normalizeSessionState(sess, validIds);
    
    // Attach details for revealed questions
    if (sess.revealedQs && sess.revealedQs.length > 0 && qSet) {
       sessionState.revealedDetails = {};
       const responses = this.getAllResponses(sessId).filter(r => r.studentId === stuId);
       const grades = this.getAIGrades(sessId).filter(g => g.studentId === stuId);
       const respMap = (responses || []).reduce((acc, r) => { acc[r.questionId] = r; return acc; }, {});
       const gradeMap = (grades || []).reduce((acc, g) => { acc[g.questionId] = g; return acc; }, {});
       const qMap = (qSet.questions || []).reduce((acc, q) => { acc[q.id] = q; return acc; }, {});

       sess.revealedQs.forEach(qId => {
          const q = qMap[qId];
          if (q) {
             const r = respMap[qId];
             const g = gradeMap[qId];
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
        needsFullscreen:d[i][10]===true||d[i][10]==='TRUE',
        gradeStatus:Grader.getStatus(sessId).state||'idle'};
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
    const viols=this.getActiveViolations(sessId,snapshot);const meta=this.getAllMeta(sessId,snapshot);
    const courseId=(sess.config && sess.config.courseId) || '';
    const roster=this.getRoster(sess.block, courseId);
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
    
    return {session:{...sess,questionCount:questions.length},students:[...students,...missing],qStats,violations:viols,totalJoined:students.length,totalFinished:students.filter(s=>s.status==='finished').length,totalActiveNow:students.filter(s=>s.activeNow).length,rosterSize:roster.length,missingCount:missing.length,gradeStatus:Grader.getStatus(sessId)};
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
      const sheet = this.sh('Sessions');
      sheet.getRange(sess.row,10).setValue(q);
      // Auto-start question timer for lockstep sessions that have a question time limit
      const qTimerSeconds = (sess.config||{}).qTimerSeconds;
      if (qTimerSeconds && sess.mode === 'lockstep') {
        const qId = qSet.questions[q] ? qSet.questions[q].id : null;
        if (qId) {
          const cfg = sess.config || {};
          const _now = Date.now();
          cfg.qTimerState = {qId, startedAt: new Date(_now).toISOString(), endTime: _now + qTimerSeconds * 1000, originalSeconds: qTimerSeconds, paused: false, pausedRemaining: null, dismissed: false, cancelled: false};
          sheet.getRange(sess.row,13).setValue(JSON.stringify(cfg));
        }
      }
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
      const sheet = this.sh('Sessions');
      sheet.getRange(sess.row,10).setValue(nextQ);
      // Auto-start question timer for lockstep sessions that have a question time limit
      const qTimerSeconds = (sess.config||{}).qTimerSeconds;
      if (qTimerSeconds && sess.mode === 'lockstep') {
        const qId = qSet.questions[nextQ] ? qSet.questions[nextQ].id : null;
        if (qId) {
          const cfg = sess.config || {};
          const _now2 = Date.now();
          cfg.qTimerState = {qId, startedAt: new Date(_now2).toISOString(), endTime: _now2 + qTimerSeconds * 1000, originalSeconds: qTimerSeconds, paused: false, pausedRemaining: null, dismissed: false, cancelled: false};
          sheet.getRange(sess.row,13).setValue(JSON.stringify(cfg));
        }
      }
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
  dismissQuestionTimer(sessId) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
      const cfg = sess.config || {};
      if (cfg.qTimerState) { cfg.qTimerState.dismissed = true; cfg.qTimerState.paused = false; cfg.qTimerState.cancelled = false; }
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      const qSet = this.getQSet(sess.setId);
      const updated = this.getSessionById(sessId);
      return {ok:true, session:this._normalizeSessionState(updated, qSet ? qSet.questions.map(q=>q.id) : [])};
    });
  },
  pauseQTimer(sessId) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
      const cfg = sess.config || {};
      if (!cfg.qTimerState || cfg.qTimerState.dismissed || cfg.qTimerState.cancelled) return {error:'No active timer'};
      if (cfg.qTimerState.paused) return {ok:true};
      cfg.qTimerState.pausedRemaining = Math.max(0, (cfg.qTimerState.endTime || Date.now()) - Date.now());
      cfg.qTimerState.paused = true;
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      const qSet = this.getQSet(sess.setId);
      const updated = this.getSessionById(sessId);
      return {ok:true, session:this._normalizeSessionState(updated, qSet ? qSet.questions.map(q=>q.id) : [])};
    });
  },
  resumeQTimer(sessId) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
      const cfg = sess.config || {};
      if (!cfg.qTimerState || !cfg.qTimerState.paused) return {error:'Timer not paused'};
      cfg.qTimerState.endTime = Date.now() + (cfg.qTimerState.pausedRemaining || 0);
      cfg.qTimerState.paused = false;
      cfg.qTimerState.pausedRemaining = null;
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      const qSet = this.getQSet(sess.setId);
      const updated = this.getSessionById(sessId);
      return {ok:true, session:this._normalizeSessionState(updated, qSet ? qSet.questions.map(q=>q.id) : [])};
    });
  },
  extendQTimer(sessId, additionalSeconds) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
      const cfg = sess.config || {};
      if (!cfg.qTimerState) return {error:'No timer active'};
      const addMs = (Number(additionalSeconds) || 0) * 1000;
      if (cfg.qTimerState.paused) {
        cfg.qTimerState.pausedRemaining = (cfg.qTimerState.pausedRemaining || 0) + addMs;
      } else {
        cfg.qTimerState.endTime = (cfg.qTimerState.endTime || Date.now()) + addMs;
        // If timer had expired (cancelled/dismissed), restart it
        if (cfg.qTimerState.cancelled || cfg.qTimerState.dismissed) {
          cfg.qTimerState.cancelled = false;
          cfg.qTimerState.dismissed = false;
        }
      }
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      const qSet = this.getQSet(sess.setId);
      const updated = this.getSessionById(sessId);
      return {ok:true, session:this._normalizeSessionState(updated, qSet ? qSet.questions.map(q=>q.id) : [])};
    });
  },
  resetQTimer(sessId) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
      const cfg = sess.config || {};
      if (!cfg.qTimerState) return {error:'No timer active'};
      const origSec = cfg.qTimerState.originalSeconds || cfg.qTimerSeconds || 60;
      cfg.qTimerState.endTime = Date.now() + origSec * 1000;
      cfg.qTimerState.paused = false;
      cfg.qTimerState.pausedRemaining = null;
      cfg.qTimerState.dismissed = false;
      cfg.qTimerState.cancelled = false;
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      const qSet = this.getQSet(sess.setId);
      const updated = this.getSessionById(sessId);
      return {ok:true, session:this._normalizeSessionState(updated, qSet ? qSet.questions.map(q=>q.id) : [])};
    });
  },
  cancelQTimer(sessId) {
    return this.withLock(() => {
      const sess = this.getSessionById(sessId); if(!sess) return {error:'Session not found'};
      const cfg = sess.config || {};
      if (cfg.qTimerState) { cfg.qTimerState.cancelled = true; cfg.qTimerState.dismissed = true; cfg.qTimerState.paused = false; }
      this.sh('Sessions').getRange(sess.row,13).setValue(JSON.stringify(cfg));
      const qSet = this.getQSet(sess.setId);
      const updated = this.getSessionById(sessId);
      return {ok:true, session:this._normalizeSessionState(updated, qSet ? qSet.questions.map(q=>q.id) : [])};
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
    const respMap = (responses || []).reduce((acc, r) => { acc[r.questionId] = r; return acc; }, {});
    const metaMap = (meta || []).reduce((acc, m) => { acc[m.questionId] = m; return acc; }, {});
    const gradeMap = (grades || []).reduce((acc, g) => { acc[g.questionId] = g; return acc; }, {});
    const details=(qSet.questions||[]).map((q,idx)=>{
      const r=respMap[q.id]||null;
      const m=metaMap[q.id]||null;
      const g=gradeMap[q.id]||null;
      const effectiveAiScore = g ? (g.overridden ? Number(g.overrideScore) : g.score) : null;
      const effectiveAiFeedback = g ? (g.overridden && g.overrideFeedback ? g.overrideFeedback : g.feedback) : null;
      return {questionId:q.id,qIndex:idx,questionText:q.text,type:q.type,answer:r?r.answer:'',isCorrect:r?r.isCorrect:null,points:r?Number(r.points)||0:0,maxPoints:q.points||1,confidence:m?m.confidence:null,aiScore:effectiveAiScore,aiFeedback:effectiveAiFeedback,overridden:g?!!g.overridden:false};
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
    const respsByStudent = (responses || []).reduce((acc, r) => {
      if (!acc[r.studentId]) acc[r.studentId] = [];
      acc[r.studentId].push(r);
      return acc;
    }, {});
    const totalByStudent={};
    students.forEach(st=>{
      const sr=respsByStudent[st.studentId]||[];
      const pts=sr.reduce((a,r)=>a+(Number(r.points)||0),0);
      const max=sr.reduce((a,r)=>a+(Number(r.maxPoints)||0),0);
      totalByStudent[st.studentId]=max?Math.round((pts/max)*100):0;
    });
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
      
      // Legacy support: inject missing qSet only when the archive genuinely has no question data
      // (sessions archived before question snapshots were stored). Flag the result so callers
      // know the question data may differ from what students saw, because it reflects the current
      // live question set rather than the version used during the session.
      if (!data.session.config.qSet && !data.session.qSet && data.session.setId) {
        data.session.config.qSet = this.getQSet(data.session.setId) || {};
        data._qSetIsLive = true;
      }
      return data;
    } catch(e) {
      Logger.log('Error parsing archived session data for sessionId ' + sessionId + ': ' + e.toString());
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
      Logger.log('Error saving AI class report for sessionId ' + sessionId + ': ' + e.toString());
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
                const ca = correctIndices.map(idx => stripHtml((qNew.choices || [])[idx] || ''));
                let sel;
                try {
                  const p = JSON.parse(rawAns);
                  sel = Array.isArray(p) ? p.map(a => stripHtml(a)) : [stripHtml(rawAns)];
                } catch (e) { sel = [stripHtml(rawAns)]; }

                if (ca.length === 1) {
                  // Single-select: index-based match with percentage-conversion fallback
                  const ansToMatch = sel[0];
                  const matchedIdx = (qNew.choices || []).findIndex(ch => {
                    const c = stripHtml(ch);
                    return c === ansToMatch || (c.length > 2 && ansToMatch.includes(c)) || (c.length > 2 && c.includes(ansToMatch));
                  });
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
                  if (finalIdx !== -1 && correctIndices.includes(finalIdx)) { pts = qNew.points; isCorrect = true; }
                } else {
                  // Multi-select: ALL selections must match correct answers, no extras
                  let correctCount = 0, incorrectCount = 0;
                  sel.forEach(cleanAns => {
                    let matched = false;
                    for (let j = 0; j < ca.length; j++) {
                      if (cleanAns === ca[j] || (ca[j].length > 2 && cleanAns.includes(ca[j])) || (ca[j].length > 2 && ca[j].includes(cleanAns))) {
                        matched = true; break;
                      }
                    }
                    if (matched) correctCount++; else incorrectCount++;
                  });
                  isCorrect = (correctCount === ca.length && incorrectCount === 0);
                  if (isCorrect) pts = qNew.points;
                }
              }
              // SA rescoring requires an AI call so we wipe its grading
              respSh.getRange(rowNum, 7, 1, 3).setValues([[isCorrect, pts, qNew.points]]);
              globalUpdate = true;
            }
          }
        }

        // Clear stale AIGrades rows for SA questions so gradeSession() will re-grade them.
        // Rows where the teacher has set an override are preserved so those scores are not lost.
        if (qNew.type === 'sa') {
          const aiSh = this.sh('AIGrades');
          if (aiSh) {
            const aiRows = aiSh.getDataRange().getValues();
            for (let i = aiRows.length - 1; i >= 1; i--) {
              if (aiRows[i][0] === sessionId && aiRows[i][3] === questionId) {
                const isOverridden = aiRows[i][9] === true || aiRows[i][9] === 'TRUE';
                if (!isOverridden) {
                  aiSh.deleteRow(i + 1);
                }
              }
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
      } catch(e) {
        Logger.log('Live update failed during rescoreQuestionFull for sessionId ' + sessionId + ': ' + e.toString() + '. Continuing to archive.');
      }

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
                const rawAns = String(r.answer || '');
                const correctIndices = qOriginal.correctIndices || [qOriginal.correctIndex || 0];
                const ca = correctIndices.map(idx => stripHtml((qOriginal.choices || [])[idx] || ''));
                let sel;
                try {
                  const p = JSON.parse(rawAns);
                  sel = Array.isArray(p) ? p.map(a => stripHtml(a)) : [stripHtml(rawAns)];
                } catch (e) { sel = [stripHtml(rawAns)]; }

                let pts = 0, isCorrect = false;
                if (ca.length === 1) {
                  // Single-select: index-based match with percentage-conversion fallback
                  const ansToMatch = sel[0];
                  const matchedIdx = (qOriginal.choices || []).findIndex(ch => {
                    const c = stripHtml(ch);
                    return c === ansToMatch || (c.length > 2 && ansToMatch.includes(c)) || (c.length > 2 && c.includes(ansToMatch));
                  });
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
                  if (finalIdx !== -1 && correctIndices.includes(finalIdx)) { pts = qOriginal.points; isCorrect = true; }
                } else {
                  // Multi-select: ALL selections must match correct answers, no extras
                  let correctCount = 0, incorrectCount = 0;
                  sel.forEach(cleanAns => {
                    let matched = false;
                    for (let j = 0; j < ca.length; j++) {
                      if (cleanAns === ca[j] || (ca[j].length > 2 && cleanAns.includes(ca[j])) || (ca[j].length > 2 && ca[j].includes(cleanAns))) {
                        matched = true; break;
                      }
                    }
                    if (matched) correctCount++; else incorrectCount++;
                  });
                  isCorrect = (correctCount === ca.length && incorrectCount === 0);
                  if (isCorrect) pts = qOriginal.points;
                }

                if (isCorrect) {
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
           // Build a lookup of teacher-overridden scores for this question so they are preserved.
           const overriddenScores = {};
           if (data.grades) {
             data.grades.forEach(g => {
               if (g.questionId === questionId && g.overridden) {
                 overriddenScores[g.studentId] = Number(g.overrideScore) || 0;
               }
             });
           }
           data.responses.forEach(r => {
             if (r.questionId === questionId) {
               r.maxPoints = qOriginal.points;
               if (overriddenScores.hasOwnProperty(r.studentId)) {
                 // Preserve the teacher's manually overridden score.
                 r.points = overriddenScores[r.studentId];
                 r.isCorrect = r.maxPoints > 0 && r.points >= r.maxPoints;
               } else {
                 r.points = 0;
                 r.isCorrect = false;
                 r.score = 0;
                 r.feedback = '';
               }
             }
           });
           if (data.grades) {
             // Remove non-overridden grades so re-grading will fill them in; keep overrides intact.
             data.grades = data.grades.filter(g => g.questionId !== questionId || !!g.overridden);
           }
        }
        
        const students = data.students || [];
        const responses = data.responses || [];
        const respsByStudent = (responses || []).reduce((acc, r) => {
          if (!acc[r.studentId]) acc[r.studentId] = [];
          acc[r.studentId].push(r);
          return acc;
        }, {});
        const totalByStudent={};
        students.forEach(st=>{
          const sr=respsByStudent[st.studentId]||[];
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
        Logger.log('Error updating archive database in rescoreQuestionFull for sessionId ' + sessionId + ': ' + e.toString());
        return false;
      }
    });
  },

  // ── ADVANCED ANALYTICS ──

  computeCTTMetrics(sessionId) {
    const data = this.getArchivedSessionData(sessionId);
    if (!data) return { error: 'Session not found in archive' };

    const students = data.students || [];
    const responses = data.responses || [];
    const qSet = (data.session.config && data.session.config.qSet) || data.session.qSet || {};
    const questions = qSet.questions || [];
    const mcQuestions = questions.filter(q => q.type === 'mc');

    // Per-student total score (%)
    const studentScores = students.map(st => {
      const sr = responses.filter(r => r.studentId === st.studentId);
      const pts = sr.reduce((a, r) => a + (Number(r.points) || 0), 0);
      const max = sr.reduce((a, r) => a + (Number(r.maxPoints) || 0), 0);
      const pct = max > 0 ? (pts / max) * 100 : 0;
      return { studentId: st.studentId, name: st.studentName, pts, max, pct };
    }).filter(s => s.max > 0);

    if (studentScores.length === 0) return { error: 'No student responses found' };

    const n = studentScores.length;
    const mean = studentScores.reduce((a, s) => a + s.pct, 0) / n;
    const variance = studentScores.reduce((a, s) => a + Math.pow(s.pct - mean, 2), 0) / n;
    const sd = Math.sqrt(variance);
    const sorted = [...studentScores].sort((a, b) => a.pct - b.pct);
    const median = n % 2 === 0
      ? (sorted[n / 2 - 1].pct + sorted[n / 2].pct) / 2
      : sorted[Math.floor(n / 2)].pct;

    // KR-20 (uses raw points variance, not % variance)
    let kr20 = null;
    if (mcQuestions.length >= 2 && n >= 3) {
      const rawPts = studentScores.map(s => s.pts);
      const rawMean = rawPts.reduce((a, b) => a + b, 0) / n;
      const rawVar = rawPts.reduce((a, x) => a + Math.pow(x - rawMean, 2), 0) / n;
      let sumPQ = 0;
      mcQuestions.forEach(q => {
        const qResps = responses.filter(r => r.questionId === q.id);
        const correct = qResps.filter(r => r.isCorrect === true || r.isCorrect === 'TRUE').length;
        if (students.length > 0) {
          const p = correct / students.length;
          sumPQ += p * (1 - p);
        }
      });
      const k = mcQuestions.length;
      if (rawVar > 0) {
        kr20 = (k / (k - 1)) * (1 - sumPQ / rawVar);
        kr20 = Math.round(Math.max(-1, Math.min(1, kr20)) * 1000) / 1000;
      }
    }

    const sem = (kr20 !== null && sd > 0) ? Math.round(sd * Math.sqrt(1 - kr20) * 10) / 10 : null;
    const skewness = sd > 0 ? Math.round((3 * (mean - median)) / sd * 100) / 100 : 0;

    // Histogram: 10 equal buckets (0–10, 10–20, …, 90–100)
    const histogram = Array.from({ length: 10 }, (_, i) => ({ label: `${i * 10}–${i * 10 + 10}`, count: 0 }));
    studentScores.forEach(s => { histogram[Math.min(9, Math.floor(s.pct / 10))].count++; });

    // Quartile cut-offs (index-based)
    const q1End = Math.floor(n * 0.25);
    const q2End = Math.floor(n * 0.50);
    const q3End = Math.floor(n * 0.75);

    const enriched = sorted.map((s, i) => ({
      ...s,
      pct: Math.round(s.pct * 10) / 10,
      zScore: sd > 0 ? Math.round(((s.pct - mean) / sd) * 100) / 100 : 0,
      percentileRank: Math.round(((i + 0.5) / n) * 100),
      quartile: i < q1End ? 1 : i < q2End ? 2 : i < q3End ? 3 : 4
    }));

    return {
      kr20,
      kr20Label: kr20 === null ? 'N/A (need ≥2 MC items & ≥3 students)'
        : kr20 < 0.5 ? 'Unacceptable (<0.50)' : kr20 < 0.7 ? 'Acceptable (0.50–0.70)'
        : kr20 < 0.9 ? 'Good (0.70–0.90)' : 'Excellent (≥0.90)',
      kr20Color: kr20 === null ? 'tx3' : kr20 < 0.5 ? 'red' : kr20 < 0.7 ? 'amb' : 'grn',
      sem,
      mean: Math.round(mean * 10) / 10,
      median: Math.round(median * 10) / 10,
      sd: Math.round(sd * 10) / 10,
      skewness,
      skewnessLabel: skewness > 0.5 ? 'Positively Skewed — many low scorers'
        : skewness < -0.5 ? 'Negatively Skewed — many high scorers' : 'Roughly Symmetric',
      n, histogram, students: enriched, mcItemCount: mcQuestions.length
    };
  },

  computeItemDiscrimination(sessionId) {
    const data = this.getArchivedSessionData(sessionId);
    if (!data) return { error: 'Session not found' };

    const students = data.students || [];
    const responses = data.responses || [];
    const qSet = (data.session.config && data.session.config.qSet) || data.session.qSet || {};
    const questions = qSet.questions || [];
    const n = students.length;
    if (n < 3) return { error: 'Need ≥ 3 students for discrimination analysis' };

    // Total raw score per student
    const studentTotals = {};
    students.forEach(st => {
      const sr = responses.filter(r => r.studentId === st.studentId);
      studentTotals[st.studentId] = sr.reduce((a, r) => a + (Number(r.points) || 0), 0);
    });
    const totalScores = Object.values(studentTotals);
    const totalMean = totalScores.reduce((a, b) => a + b, 0) / n;
    const totalSD = Math.sqrt(totalScores.reduce((a, x) => a + Math.pow(x - totalMean, 2), 0) / n);

    const stripHtml = s => String(s || '').replace(/<[^>]*>/g, '').replace(/&[^;]+;/g, ' ').replace(/\s+/g, '').toLowerCase();

    return questions.map((q, idx) => {
      const qResps = responses.filter(r => r.questionId === q.id);
      const correctCount = qResps.filter(r => r.isCorrect === true || r.isCorrect === 'TRUE').length;
      const p = n > 0 ? correctCount / n : 0;

      // Point-biserial for overall item
      const correctTotals = qResps.filter(r => r.isCorrect === true || r.isCorrect === 'TRUE').map(r => studentTotals[r.studentId] || 0);
      const incorrectTotals = students.filter(st => !qResps.some(r => r.studentId === st.studentId && (r.isCorrect === true || r.isCorrect === 'TRUE'))).map(st => studentTotals[st.studentId] || 0);
      const mCorrect = correctTotals.length > 0 ? correctTotals.reduce((a, b) => a + b, 0) / correctTotals.length : 0;
      const mIncorrect = incorrectTotals.length > 0 ? incorrectTotals.reduce((a, b) => a + b, 0) / incorrectTotals.length : 0;

      let discrimination = 0;
      if (totalSD > 0 && p > 0 && p < 1) {
        discrimination = ((mCorrect - mIncorrect) / totalSD) * Math.sqrt(p * (1 - p));
        discrimination = Math.round(Math.max(-1, Math.min(1, discrimination)) * 1000) / 1000;
      }
      const discLabel = discrimination >= 0.3 ? 'Excellent' : discrimination >= 0.2 ? 'Good' : discrimination >= 0.1 ? 'Fair' : 'Poor';
      const discColor = discrimination >= 0.3 ? 'grn' : discrimination >= 0.1 ? 'amb' : 'red';

      // Per-choice analysis (MC only)
      let choiceStats = null;
      let nonFunctioningCount = 0;
      if (q.type === 'mc' && q.choices) {
        const correctIdxs = q.correctIndices || [q.correctIndex || 0];
        choiceStats = q.choices.map((optText, optIdx) => {
          const isCorrect = correctIdxs.includes(optIdx);
          const cleanOpt = stripHtml(optText);
          const optResps = qResps.filter(r => {
            let a = String(r.answer || '');
            try { const p = JSON.parse(a); if (Array.isArray(p)) a = p.join(' '); } catch (e) {}
            const ca = stripHtml(a);
            return ca === cleanOpt || (cleanOpt.length > 2 && ca.includes(cleanOpt)) || (cleanOpt.length > 2 && cleanOpt.includes(ca) && ca.length > 2);
          });
          const chooseCount = optResps.length;
          const choosePct = n > 0 ? Math.round((chooseCount / n) * 100) : 0;

          // Point-biserial for this choice
          const chooserTotals = optResps.map(r => studentTotals[r.studentId] || 0);
          const nonChooserTotals = students.filter(st => !optResps.some(r => r.studentId === st.studentId)).map(st => studentTotals[st.studentId] || 0);
          const mChoose = chooserTotals.length > 0 ? chooserTotals.reduce((a, b) => a + b, 0) / chooserTotals.length : 0;
          const mNon = nonChooserTotals.length > 0 ? nonChooserTotals.reduce((a, b) => a + b, 0) / nonChooserTotals.length : 0;
          const pChoice = n > 0 ? chooseCount / n : 0;
          let choicePbis = 0;
          if (totalSD > 0 && pChoice > 0 && pChoice < 1) {
            choicePbis = Math.round(((mChoose - mNon) / totalSD) * Math.sqrt(pChoice * (1 - pChoice)) * 1000) / 1000;
          }
          const isFunctioning = choosePct >= 5 || isCorrect;
          if (!isCorrect && !isFunctioning) nonFunctioningCount++;
          return {
            text: optText, letter: String.fromCharCode(65 + optIdx),
            isCorrect, count: chooseCount, pct: choosePct, pointBiserial: choicePbis, isFunctioning,
            pbisFlag: isCorrect ? (choicePbis > 0 ? 'good' : 'warn') : (choicePbis < 0 ? 'good' : 'warn')
          };
        });
      }

      return {
        questionId: q.id, idx, text: q.text, type: q.type || 'sa',
        difficulty: Math.round(p * 100) / 100,
        difficultyPct: Math.round(p * 100),
        difficultyLabel: p >= 0.9 ? 'Very Easy' : p >= 0.7 ? 'Easy' : p >= 0.4 ? 'Moderate' : p >= 0.2 ? 'Difficult' : 'Very Difficult',
        difficultyColor: p >= 0.7 ? 'grn' : p >= 0.4 ? 'amb' : 'red',
        discrimination, discLabel, discColor, choiceStats, nonFunctioningCount,
        respondentCount: qResps.length
      };
    });
  },

  computeConfidenceCalibration(sessionId) {
    const data = this.getArchivedSessionData(sessionId);
    if (!data) return { error: 'Session not found' };

    const students = data.students || [];
    const responses = data.responses || [];
    const meta = data.meta || [];
    const qSet = (data.session.config && data.session.config.qSet) || data.session.qSet || {};
    const questions = qSet.questions || [];

    if (meta.length === 0) return { error: 'No confidence data for this session' };

    const respMap = {};
    responses.forEach(r => { respMap[r.studentId + '|' + r.questionId] = r; });

    // Per-student calibration
    const studentCalibration = students.map(st => {
      const stuMeta = meta.filter(m => m.studentId === st.studentId);
      if (stuMeta.length === 0) return null;
      let brierSum = 0, brierCount = 0;
      let accConf = 0, accUnconf = 0, inaccConf = 0, inaccUnconf = 0;
      stuMeta.forEach(m => {
        const r = respMap[st.studentId + '|' + m.questionId];
        if (!r || !m.confidence) return;
        const conf = Number(m.confidence);
        if (conf === 0) return;
        const confNorm = (conf - 1) / 4; // 1–5 → 0–1
        const isCorrect = r.isCorrect === true || r.isCorrect === 'TRUE' ? 1 : 0;
        brierSum += Math.pow(confNorm - isCorrect, 2);
        brierCount++;
        const highConf = conf >= 3;
        if (isCorrect && highConf) accConf++;
        else if (isCorrect && !highConf) accUnconf++;
        else if (!isCorrect && highConf) inaccConf++;
        else inaccUnconf++;
      });
      const brierScore = brierCount > 0 ? Math.round((brierSum / brierCount) * 1000) / 1000 : null;
      const total = accConf + accUnconf + inaccConf + inaccUnconf;
      const dangerPct = total > 0 ? Math.round((inaccConf / total) * 100) : 0;
      return {
        studentId: st.studentId, name: st.studentName, brierScore,
        brierLabel: brierScore === null ? 'N/A' : brierScore < 0.1 ? 'Excellent' : brierScore < 0.2 ? 'Good' : brierScore < 0.35 ? 'Fair' : 'Poor',
        brierColor: brierScore === null ? 'tx3' : brierScore < 0.1 ? 'grn' : brierScore < 0.2 ? 'grn' : brierScore < 0.35 ? 'amb' : 'red',
        quadrant: { accConf, accUnconf, inaccConf, inaccUnconf }, totalRated: brierCount, dangerPct
      };
    }).filter(Boolean).sort((a, b) => b.dangerPct - a.dangerPct);

    // Per-item overconfidence rate
    const itemOverconfidence = questions.map(q => {
      const qResps = responses.filter(r => r.questionId === q.id);
      const qMeta = meta.filter(m => m.questionId === q.id);
      let overconfident = 0, total = 0;
      qMeta.forEach(m => {
        const r = qResps.find(rx => rx.studentId === m.studentId);
        if (!r || !m.confidence) return;
        total++;
        const isCorrect = r.isCorrect === true || r.isCorrect === 'TRUE';
        if (!isCorrect && Number(m.confidence) >= 3) overconfident++;
      });
      return { questionId: q.id, text: q.text, overconfidenceCount: overconfident, overconfidenceRate: total > 0 ? Math.round((overconfident / total) * 100) : 0, total };
    }).filter(q => q.total > 0).sort((a, b) => b.overconfidenceRate - a.overconfidenceRate);

    // Class-wide quadrant summary
    const classSummary = studentCalibration.reduce(
      (acc, s) => { acc.accConf += s.quadrant.accConf; acc.accUnconf += s.quadrant.accUnconf; acc.inaccConf += s.quadrant.inaccConf; acc.inaccUnconf += s.quadrant.inaccUnconf; return acc; },
      { accConf: 0, accUnconf: 0, inaccConf: 0, inaccUnconf: 0 }
    );

    return { studentCalibration, itemOverconfidence, classSummary };
  },

  getCrossSessionRiskReport(courseId, blockFilter) {
    const archive = this.sh('Archive');
    if (!archive) return { error: 'Archive not found' };

    const values = archive.getDataRange().getValues();
    const courseSessions = [];
    values.slice(1).forEach(row => {
      if (!row[8]) return;
      try {
        const data = JSON.parse(row[8]);
        const cfg = (data.session && data.session.config) || {};
        if (!courseId || cfg.courseId === courseId) {
          let timerAlloc = null;
          try {
            const tj = data.session.timerJSON ? (typeof data.session.timerJSON === 'string' ? JSON.parse(data.session.timerJSON) : data.session.timerJSON) : null;
            if (tj && tj.duration) timerAlloc = Number(tj.duration) * 60 * 1000;
          } catch (e) {}
          courseSessions.push({ sessionId: row[0], setName: row[2], block: row[3], startedAt: row[4], students: data.students || [], responses: data.responses || [], violations: data.violations || [], timerAlloc, config: cfg });
        }
      } catch (e) { Logger.log('getCrossSessionRiskReport parse error: ' + e.toString()); }
    });

    if (courseSessions.length === 0) return { error: 'No archived sessions found for this course' };

    // Determine which blocks to analyze
    const useBlockFilter = blockFilter && blockFilter !== 'all';
    const allBlocks = useBlockFilter
      ? [blockFilter]
      : [...new Set(courseSessions.map(s => s.block).filter(Boolean))].sort();

    // Build roster map per block for "never accessed" detection
    const rosterByBlock = {};
    try {
      const rosterData = this.getRostersByCourse(courseId);
      (rosterData || []).forEach(roster => {
        const bk = roster.block || '';
        if (!rosterByBlock[bk]) rosterByBlock[bk] = {};
        (roster.students || []).forEach(s => {
          const key = this.normalizeStudentName(s.name);
          if (!rosterByBlock[bk][key]) rosterByBlock[bk][key] = { name: s.name };
        });
      });
    } catch (e) {}

    // Helper: compute risk categories for a set of sessions
    const computeBlockRisk = (sessions) => {
      const studentData = {};
      const getOrCreate = (id, name) => {
        if (!studentData[id]) studentData[id] = { studentId: id, name, sessions: [], violationSessions: [], totalJoined: 0, totalFinished: 0, scores: [] };
        return studentData[id];
      };

      sessions.forEach(sess => {
        sess.students.forEach(st => {
          const sd = getOrCreate(st.studentId, st.studentName);
          const sr = sess.responses.filter(r => r.studentId === st.studentId);
          const pts = sr.reduce((a, r) => a + (Number(r.points) || 0), 0);
          const max = sr.reduce((a, r) => a + (Number(r.maxPoints) || 0), 0);
          const score = max > 0 ? Math.round((pts / max) * 100) : null;
          let durationMs = null;
          if (st.joinedAt && st.finishedAt) durationMs = new Date(st.finishedAt).getTime() - new Date(st.joinedAt).getTime();
          const stuViolations = sess.violations.filter(v => v.studentId === st.studentId);
          if (stuViolations.length > 0) sd.violationSessions.push({ sessionId: sess.sessionId, setName: sess.setName, startedAt: sess.startedAt, count: stuViolations.length });
          const finished = st.status === 'finished';
          if (st.status !== 'pending') sd.totalJoined++;
          if (finished) { sd.totalFinished++; if (score !== null) sd.scores.push(score); }
          sd.sessions.push({ sessionId: sess.sessionId, setName: sess.setName, startedAt: sess.startedAt, block: sess.block, score, finished, joinedAt: st.joinedAt, finishedAt: st.finishedAt, durationMs, timerAlloc: sess.timerAlloc, answerCount: sr.length });
        });
      });

      // Class avg/SD for "consistently low" threshold — computed per block
      const classScores = [];
      sessions.forEach(sess => {
        sess.students.forEach(st => {
          if (st.status !== 'finished') return;
          const sr = sess.responses.filter(r => r.studentId === st.studentId);
          const pts = sr.reduce((a, r) => a + (Number(r.points) || 0), 0);
          const max = sr.reduce((a, r) => a + (Number(r.maxPoints) || 0), 0);
          if (max > 0) classScores.push((pts / max) * 100);
        });
      });
      const classAvg = classScores.length > 0 ? classScores.reduce((a, b) => a + b, 0) / classScores.length : 50;
      const classSD = classScores.length > 1 ? Math.sqrt(classScores.reduce((a, x) => a + Math.pow(x - classAvg, 2), 0) / classScores.length) : 15;
      const lowThreshold = classAvg - classSD;

      const neverAccessed = [], consistentlyLow = [], avoiders = [], violators = [], rapidSubmitters = [];

      const fmtDate = (d) => d ? new Date(d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '';

      Object.values(studentData).forEach(st => {
        if (st.totalJoined === 0) {
          neverAccessed.push({ name: st.name, sessionCount: st.sessions.length, sessions: st.sessions.map(s => ({ setName: s.setName, date: fmtDate(s.startedAt) })) });
          return;
        }
        const joinedNotFinished = st.sessions.filter(s => s.joinedAt && !s.finished).length;
        if (joinedNotFinished >= 2) avoiders.push({ name: st.name, avoidCount: joinedNotFinished, sessions: st.sessions.map(s => ({ setName: s.setName, date: fmtDate(s.startedAt) })) });
        if (st.scores.length >= 2) {
          const stuAvg = st.scores.reduce((a, b) => a + b, 0) / st.scores.length;
          if (stuAvg < lowThreshold) consistentlyLow.push({ name: st.name, avgScore: Math.round(stuAvg), threshold: Math.round(lowThreshold), classAvg: Math.round(classAvg), scores: st.scores.map(s => Math.round(s)), sessions: st.sessions.filter(s => s.score !== null).map(s => ({ setName: s.setName, date: fmtDate(s.startedAt), score: s.score })) });
        }
        if (st.violationSessions.length >= 1) violators.push({ name: st.name, sessionsWithViolations: st.violationSessions.map(v => ({ setName: v.setName, date: fmtDate(v.startedAt), count: v.count })) });
        const rapidSess = st.sessions.filter(s => s.finished && s.durationMs && s.timerAlloc && s.durationMs < s.timerAlloc * 0.25);
        if (rapidSess.length >= 2) rapidSubmitters.push({ name: st.name, rapidCount: rapidSess.length, sessions: rapidSess.map(s => ({ setName: s.setName, date: fmtDate(s.startedAt), durationMin: Math.round(s.durationMs / 60000), allocMin: Math.round(s.timerAlloc / 60000) })) });
      });

      return {
        sessionCount: sessions.length,
        classAvg: Math.round(classAvg), classSD: Math.round(classSD), lowThreshold: Math.round(lowThreshold),
        rosterSize: Object.keys(studentData).length,
        neverAccessed: neverAccessed.sort((a, b) => a.name.localeCompare(b.name)),
        consistentlyLow: consistentlyLow.sort((a, b) => a.avgScore - b.avgScore),
        avoiders: avoiders.sort((a, b) => b.avoidCount - a.avoidCount),
        violators: violators.sort((a, b) => b.sessionsWithViolations.length - a.sessionsWithViolations.length),
        rapidSubmitters: rapidSubmitters.sort((a, b) => b.rapidCount - a.rapidCount)
      };
    };

    // Compute per-block results
    const byBlock = {};
    let totalFlags = 0;
    let totalSessions = 0;
    allBlocks.forEach(block => {
      const blockSessions = courseSessions.filter(s => s.block === block);
      if (blockSessions.length === 0) return;
      const result = computeBlockRisk(blockSessions);

      // Add roster-only students who never appeared in any session for this block
      const blockRoster = rosterByBlock[block] || {};
      const blockStudentNames = new Set();
      blockSessions.forEach(sess => sess.students.forEach(st => blockStudentNames.add(this.normalizeStudentName(st.studentName))));
      Object.values(blockRoster).forEach(rs => {
        const normName = this.normalizeStudentName(rs.name);
        if (!blockStudentNames.has(normName)) result.neverAccessed.push({ name: rs.name, sessionCount: 0, sessions: [], fromRoster: true });
      });
      result.neverAccessed.sort((a, b) => a.name.localeCompare(b.name));

      byBlock[block] = result;
      totalFlags += result.neverAccessed.length + result.consistentlyLow.length + result.avoiders.length + result.violators.length + result.rapidSubmitters.length;
      totalSessions += result.sessionCount;
    });

    return { courseId, blockFilter: blockFilter || 'all', byBlock, totalSessions, totalFlags };
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
      Logger.log('Error dismissing violation for sessionId ' + sessionId + ' at ' + timestamp + ': ' + e.toString());
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

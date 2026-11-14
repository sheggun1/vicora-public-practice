/**
 * VICORA PUBLIC PRACTICE — standalone Google Apps Script backend
 * Google Sheets database; independent of the VICORA School CBT backend.
 *
 * FIRST RUN:
 * 1. Create a NEW Apps Script project.
 * 2. Replace Code.gs with this complete file and save.
 * 3. Select setupVICORAPublicPractice and click Run; approve permissions.
 * 4. Run initializeVICORAAdminLogin once to seed the username/password admin account.
 * 5. Deploy > New deployment > Web app; Execute as: Me; access: Anyone.
 * 6. Copy the /exec URL into Vercel as VICORA_APPS_SCRIPT_URL.
 *
 * The setup function is safe to run again: it reuses the same database and key.
 * It does not overwrite existing records.
 */

const VP = Object.freeze({
  VERSION: '1.3.0',
  DB_PROP: 'VICORA_PRACTICE_DATABASE_ID',
  ADMIN_PROP: 'VICORA_PRACTICE_ADMIN_KEY',
  SESSION_MS: 365 * 24 * 60 * 60 * 1000,
  TYPES: ['WAEC', 'NECO', 'JAMB', 'GENERAL'],
  LEVELS: ['PRIMARY', 'JUNIOR_SECONDARY', 'SENIOR_SECONDARY'],
  SHEETS: {
    PracticeUsers: ['user_id','full_name','email','password_salt','password_hash','active','created_at','last_login','school_level','class_level'],
    PracticeSessions: ['session_id','token_hash','user_id','expires_at','created_at','revoked'],
    PracticeQuestions: ['question_id','exam_type','school_level','class_level','subject','year','topic','difficulty','question_text','image_url','option_a','option_b','option_c','option_d','correct_option','explanation','marks','active','created_at','updated_at'],
    PracticeAttempts: ['attempt_id','user_id','exam_type','school_level','class_level','subject','exam_mode','proctor_enabled','proctor_events_json','question_ids_json','question_count','started_at','duration_minutes','deadline_at','submitted_at','score','total','status'],
    PracticeAnswers: ['answer_id','attempt_id','question_id','selected_option'],
    PracticeAdmins: ['admin_id','username','password_salt','password_hash','active','created_at','last_login'],
    PracticeAdminSessions: ['session_id','token_hash','admin_id','expires_at','created_at','revoked'],
    PracticePasswordResets: ['reset_id','user_id','token_hash','expires_at','used','created_at']
  }
});

/**
 * Run this function once from the Apps Script editor.
 * Creates a dedicated spreadsheet and all tabs, and generates the admin key.
 * Admin credentials are stored as salted password hashes in the database.
 */
function setupVICORAPublicPractice() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty(VP.DB_PROP);
  let ss;
  if (dbId) {
    ss = SpreadsheetApp.openById(dbId);
  } else {
    ss = SpreadsheetApp.create('VICORA Public Practice Database');
    dbId = ss.getId();
    props.setProperty(VP.DB_PROP, dbId);
  }

  Object.keys(VP.SHEETS).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const required = VP.SHEETS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, required.length).setValues([required]);
      sh.setFrozenRows(1);
    } else {
      const existing = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      required.forEach(function(h) {
        if (existing.indexOf(h) < 0) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
        }
      });
    }
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }

  const result = {
    status: 'READY',
    message: 'VICORA Public Practice database is ready.',
    spreadsheetUrl: ss.getUrl(),
    spreadsheetId: dbId,
    tabs: Object.keys(VP.SHEETS),
    nextSteps: [
      'Save the administrator key securely.',
      'Deploy this project as a Web app, executing as you, accessible to anyone.',
      'Set the Vercel environment variable VICORA_APPS_SCRIPT_URL to the deployed /exec URL.'
    ]
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'health');
  if (action === 'health') {
    return json_({ok:true, service:'VICORA Public Practice API', version:VP.VERSION, timestamp:new Date().toISOString()});
  }
  return json_({ok:false, error:'Unsupported GET action.'});
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = String(body.action || '');
    let data;
    switch (action) {
      case 'practice.register': data = register_(body); break;
      case 'practice.login': data = login_(body); break;
      case 'practice.admin.login': data = adminLogin_(body); break;
      case 'practice.admin.changePassword': requireAdmin_(body); data = adminChangePassword_(body); break;
      case 'practice.password.requestReset': data = requestPasswordReset_(body); break;
      case 'practice.password.completeReset': data = completePasswordReset_(body); break;
      case 'practice.logout': data = logout_(body); break;
      case 'practice.me': data = me_(body); break;
      case 'practice.catalog': data = catalog_(); break;
      case 'practice.history': data = history_(body); break;
      case 'practice.history.detail': data = historyDetail_(body); break;
      case 'practice.start': data = start_(body); break;
      case 'practice.resume': data = resume_(body); break;
      case 'practice.submit': data = submit_(body); break;
      case 'practice.admin.listQuestions': requireAdmin_(body); data = listQuestions_(); break;
      case 'practice.admin.importQuestions': requireAdmin_(body); data = importQuestions_(body); break;
      case 'practice.admin.listUsers': requireAdmin_(body); data = listUsers_(); break;
      case 'practice.admin.setUserStatus': requireAdmin_(body); data = setUserStatus_(body); break;
      case 'practice.admin.uploadImage': requireAdmin_(body); data = uploadQuestionImage_(body); break;
      case 'practice.admin.verifyQuestion': requireAdmin_(body); data = verifyQuestionWithAI_(body); break;
      case 'practice.admin.aiSettings': requireAdmin_(body); data = aiSettings_(); break;
      case 'practice.admin.saveAiSettings': requireAdmin_(body); data = saveAiSettings_(body); break;
      case 'practice.admin.chatAI': requireAdmin_(body); data = adminAIChat_(body); break;
      case 'practice.ai.tutor': data = studentTutor_(body); break;
      default: throw new Error('Unsupported action: ' + action);
    }
    return json_({ok:true, data:data});
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return json_({ok:false, error:safeError_(err)});
  }
}

// -------------------------- Authentication -------------------------------

function register_(b) {
  const name = required_(b.fullName, 'Full name', 120);
  const email = validEmail_(b.email);
  const password = String(b.password || '');
  const schoolLevel = normalLevel_(b.schoolLevel || 'SENIOR_SECONDARY');
  const classLevel = normalClass_(schoolLevel,b.classLevel || defaultClass_(schoolLevel));
  if (password.length < 10 || password.length > 128) throw new Error('Password must contain between 10 and 128 characters.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (findOne_('PracticeUsers', 'email', email)) throw new Error('An account with this email address already exists.');
    const now = new Date().toISOString();
    const id = Utilities.getUuid();
    const salt = randomToken_();
    const token = randomToken_();
    append_('PracticeUsers', {
      user_id:id, full_name:name, email:email, password_salt:salt,
      password_hash:hash_(salt + password), active:'true', created_at:now, last_login:now,
      school_level:schoolLevel, class_level:classLevel
    });
    append_('PracticeSessions', {
      session_id:Utilities.getUuid(), token_hash:hash_(token), user_id:id,
      expires_at:new Date(Date.now() + VP.SESSION_MS).toISOString(), created_at:now, revoked:'false'
    });
    return {sessionToken:token, user:{userId:id, fullName:name, email:email, schoolLevel:schoolLevel, classLevel:classLevel}};
  } finally {
    lock.releaseLock();
  }
}

function login_(b) {
  const email = validEmail_(b.email);
  const password = String(b.password || '');
  const user = findOne_('PracticeUsers', 'email', email);
  if (!user || user.active !== 'true' || user.password_hash !== hash_(String(user.password_salt) + password)) {
    throw new Error('Email or password is incorrect.');
  }
  const token = randomToken_();
  const now = new Date().toISOString();
  append_('PracticeSessions', {
    session_id:Utilities.getUuid(), token_hash:hash_(token), user_id:user.user_id,
    expires_at:new Date(Date.now() + VP.SESSION_MS).toISOString(), created_at:now, revoked:'false'
  });
  updateById_('PracticeUsers', 'user_id', user.user_id, {last_login:now});
  return {sessionToken:token, user:{userId:user.user_id, fullName:user.full_name, email:user.email, schoolLevel:user.school_level || 'SENIOR_SECONDARY', classLevel:user.class_level || 'SS 1'}};
}

function auth_(b) {
  const token = String(b.sessionToken || '');
  if (token.length < 40 || token.length > 200) throw new Error('Please sign in to your VICORA Practice account.');
  const session = findOne_('PracticeSessions', 'token_hash', hash_(token));
  if (!session || session.revoked === 'true' || Date.parse(session.expires_at) <= Date.now()) {
    throw new Error('Your practice session has expired. Please sign in again.');
  }
  const user = findOne_('PracticeUsers', 'user_id', session.user_id);
  if (!user || user.active !== 'true') throw new Error('Practice account is inactive.');
  return user;
}

function me_(b) {
  const u = auth_(b);
  let ownAttempts = read_('PracticeAttempts').filter(a => a.user_id === u.user_id);
  ownAttempts.filter(a => a.status === 'in_progress' && attemptDeadline_(a) <= Date.now()).forEach(a => {
    submit_({sessionToken:b.sessionToken, attemptId:a.attempt_id, answers:{}, proctorEvents:[]});
  });
  ownAttempts = read_('PracticeAttempts').filter(a => a.user_id === u.user_id);
  const active = ownAttempts.find(a => a.status === 'in_progress') || null;
  const canUseTutor = !active;
  return {
    user:{userId:u.user_id, fullName:u.full_name, email:u.email, schoolLevel:u.school_level || 'SENIOR_SECONDARY', classLevel:u.class_level || 'SS 1'},
    activeExam:active ? {attemptId:active.attempt_id, deadlineAt:new Date(attemptDeadline_(active)).toISOString(), durationMinutes:Number(active.duration_minutes)||60} : null,
    canUseTutor:canUseTutor
  };
}
function resume_(b) {
  const user = auth_(b);
  let attempt = read_('PracticeAttempts').find(a => a.user_id === user.user_id && a.status === 'in_progress');
  if (!attempt) throw new Error('There is no examination to resume.');
  if (attemptDeadline_(attempt) <= Date.now()) {
    const result = submit_({sessionToken:b.sessionToken,attemptId:attempt.attempt_id,answers:{},proctorEvents:[]});
    return {expired:true,result:result};
  }
  const ids = JSON.parse(attempt.question_ids_json || '[]');
  const byId = {};
  read_('PracticeQuestions').forEach(q => { byId[q.question_id] = q; });
  const questions = ids.map(id => byId[id]).filter(Boolean).map(q => safeQuestion_(q,false,''));
  return {
    expired:false,attemptId:attempt.attempt_id,examType:attempt.exam_type,
    schoolLevel:attempt.school_level,classLevel:attempt.class_level,
    examMode:attempt.exam_mode||'standard',proctorEnabled:attempt.proctor_enabled==='true',
    subject:attempt.subject,questionCount:questions.length,
    durationMinutes:Number(attempt.duration_minutes)||60,
    deadlineAt:new Date(attemptDeadline_(attempt)).toISOString(),questions:questions
  };
}

function attemptDeadline_(attempt) {
  const explicit = Date.parse(attempt.deadline_at || '');
  if (Number.isFinite(explicit)) return explicit;
  const started = Date.parse(attempt.started_at || '');
  return (Number.isFinite(started) ? started : Date.now()) + (Math.max(1,Number(attempt.duration_minutes)||60) * 60000);
}

function logout_(b) {
  const token = String(b.sessionToken || '');
  let submittedAttempt = null;
  if (token) {
    const user = auth_(b);
    const active = read_('PracticeAttempts').find(a => a.user_id === user.user_id && a.status === 'in_progress');
    if (active) {
      submittedAttempt = submit_({
        sessionToken:token,
        attemptId:active.attempt_id,
        answers:b.answers && typeof b.answers === 'object' ? b.answers : {},
        proctorEvents:Array.isArray(b.proctorEvents) ? b.proctorEvents : []
      });
    }
    const session = findOne_('PracticeSessions', 'token_hash', hash_(token));
    if (session) updateById_('PracticeSessions', 'session_id', session.session_id, {revoked:'true'});
  }
  return {loggedOut:true, autoSubmitted:!!submittedAttempt, result:submittedAttempt ? {
    attemptId:submittedAttempt.attemptId, percentage:submittedAttempt.percentage,
    score:submittedAttempt.score, total:submittedAttempt.total,
    timeExpired:submittedAttempt.timeExpired
  } : null};
}

// ---------------------------- Question bank -------------------------------

function catalog_() {
  const rows = read_('PracticeQuestions').filter(q => q.active === 'true');
  const subjects = {};
  const counts = {};
  const subjectsByLevel = {};
  rows.forEach(function(q) {
    if (!subjects[q.exam_type]) subjects[q.exam_type] = [];
    const level = q.school_level || 'SENIOR_SECONDARY';
    if (!subjectsByLevel[level]) subjectsByLevel[level] = {};
    if (!subjectsByLevel[level][q.exam_type]) subjectsByLevel[level][q.exam_type] = [];
    if (subjectsByLevel[level][q.exam_type].indexOf(q.subject) < 0) subjectsByLevel[level][q.exam_type].push(q.subject);
    if (subjects[q.exam_type].indexOf(q.subject) < 0) subjects[q.exam_type].push(q.subject);
    counts[q.exam_type + '|' + q.subject] = (counts[q.exam_type + '|' + q.subject] || 0) + 1;
  });
  Object.keys(subjects).forEach(k => subjects[k].sort());
  return {examTypes:VP.TYPES, levels:VP.LEVELS, subjects:subjects, subjectsByLevel:subjectsByLevel, counts:counts};
}

function start_(b) {
  const user = auth_(b);
  const type = String(b.examType || '').toUpperCase();
  const schoolLevel = normalLevel_(b.schoolLevel || user.school_level || 'SENIOR_SECONDARY');
  const classLevel = normalClass_(schoolLevel,b.classLevel || user.class_level || defaultClass_(schoolLevel));
  const examMode = String(b.examMode || 'standard').toLowerCase() === 'proctored' ? 'proctored' : 'standard';
  const proctorEnabled = examMode === 'proctored' && b.proctorEnabled !== false;
  const subject = required_(b.subject, 'Subject', 100);
  const count = Math.max(5, Math.min(60, Number(b.questionCount) || 20));
  const durationMinutes = Number(b.durationMinutes);
  if (!Number.isFinite(durationMinutes) || durationMinutes < 1 || durationMinutes > 300) throw new Error('Select an examination duration between 1 and 300 minutes.');
  read_('PracticeAttempts').filter(a => a.user_id === user.user_id && a.status === 'in_progress' && attemptDeadline_(a) <= Date.now()).forEach(a => {
    submit_({sessionToken:b.sessionToken, attemptId:a.attempt_id, answers:{}, proctorEvents:[]});
  });
  if (read_('PracticeAttempts').some(a => a.user_id === user.user_id && a.status === 'in_progress')) throw new Error('You already have an examination in progress. Resume or submit it before starting another.');
  if (VP.TYPES.indexOf(type) < 0) throw new Error('Choose a valid examination type.');
  const available = read_('PracticeQuestions').filter(q => q.active === 'true' && q.exam_type === type && q.subject === subject && (q.school_level || 'SENIOR_SECONDARY') === schoolLevel && (!q.class_level || q.class_level === classLevel));
  if (!available.length) throw new Error('There are no questions for this selection yet.');
  const shuffled = available.slice();
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = shuffled[i]; shuffled[i] = shuffled[j]; shuffled[j] = tmp;
  }
  const chosen = shuffled.slice(0, count);
  const ids = chosen.map(q => q.question_id);
  const attemptId = Utilities.getUuid();
  const now = new Date().toISOString();
  const deadlineAt = new Date(Date.now() + durationMinutes * 60000).toISOString();
  append_('PracticeAttempts', {
    attempt_id:attemptId, user_id:user.user_id, exam_type:type, school_level:schoolLevel, class_level:classLevel, subject:subject,
    exam_mode:examMode, proctor_enabled:String(proctorEnabled), proctor_events_json:'[]',
    question_ids_json:JSON.stringify(ids), question_count:ids.length,
    started_at:now, duration_minutes:String(durationMinutes), deadline_at:deadlineAt,
    submitted_at:'', score:'', total:'', status:'in_progress'
  });
  return {attemptId:attemptId, examType:type, schoolLevel:schoolLevel, classLevel:classLevel, examMode:examMode, proctorEnabled:proctorEnabled, subject:subject, questionCount:ids.length, durationMinutes:durationMinutes, deadlineAt:deadlineAt, questions:chosen.map(q => safeQuestion_(q, false, ''))};
}

function safeQuestion_(q, includeAnswer, selected) {
  const out = {
    id:q.question_id, examType:q.exam_type, subject:q.subject, year:q.year,
    topic:q.topic, difficulty:q.difficulty, questionText:q.question_text, imageUrl:q.image_url || '',
    options:{A:q.option_a, B:q.option_b, C:q.option_c, D:q.option_d}
  };
  if (includeAnswer) {
    out.selectedOption = selected || '';
    out.correctOption = q.correct_option;
    out.correct = !!selected && selected === q.correct_option;
    out.explanation = q.explanation || '';
  }
  return out;
}

function submit_(b) {
  const user = auth_(b);
  const id = required_(b.attemptId, 'Attempt ID', 80);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const attempt = findOne_('PracticeAttempts', 'attempt_id', id);
    if (!attempt || attempt.user_id !== user.user_id) throw new Error('Practice attempt not found.');
    if (attempt.status !== 'in_progress') throw new Error('This practice attempt has already been submitted.');
    const timeExpired = Date.now() >= attemptDeadline_(attempt);
    const ids = JSON.parse(attempt.question_ids_json);
    const questions = read_('PracticeQuestions');
    const byId = {};
    questions.forEach(q => { byId[q.question_id] = q; });
    const answers = !timeExpired && b.answers && typeof b.answers === 'object' ? b.answers : {};
    let score = 0, total = 0, correctCount = 0;
    const answerRows = [];
    ids.forEach(function(qid) {
      const q = byId[qid];
      if (!q) return;
      const selectedRaw = String(answers[qid] || '').toUpperCase();
      const selected = ['A','B','C','D'].indexOf(selectedRaw) >= 0 ? selectedRaw : '';
      const marks = Math.max(0.01, Math.min(100, Number(q.marks) || 1));
      total += marks;
      if (selected && selected === q.correct_option) { score += marks; correctCount++; }
      answerRows.push([Utilities.getUuid(), id, qid, selected]);
    });
    if (answerRows.length) {
      const sh = sheet_('PracticeAnswers');
      sh.getRange(sh.getLastRow() + 1, 1, answerRows.length, 4).setValues(answerRows);
    }
    updateById_('PracticeAttempts', 'attempt_id', id, {
      submitted_at:new Date().toISOString(), score:String(score), total:String(total), status:'submitted',
      proctor_events_json:JSON.stringify(Array.isArray(b.proctorEvents) ? b.proctorEvents.slice(0,100) : [])
    });
    const review = ids.map(qid => {
      const q = byId[qid];
      const selected = String(answers[qid] || '').toUpperCase();
      return q ? safeQuestion_(q, true, ['A','B','C','D'].indexOf(selected) >= 0 ? selected : '') : null;
    }).filter(Boolean);
    return {
      attemptId:id, examType:attempt.exam_type, subject:attempt.subject, timeExpired:timeExpired,
      score:score, total:total, percentage:total ? Math.round(score / total * 1000) / 10 : 0,
      correctCount:correctCount, questionCount:ids.length, review:review
    };
  } finally {
    lock.releaseLock();
  }
}

function history_(b) {
  const user = auth_(b);
  const rows = read_('PracticeAttempts').filter(a => a.user_id === user.user_id)
    .sort((a,b) => String(b.started_at).localeCompare(String(a.started_at))).slice(0,50);
  return {attempts:rows.map(function(a) {
    const total = Number(a.total) || 0, score = Number(a.score) || 0;
    return {
      id:a.attempt_id, exam_type:a.exam_type, school_level:a.school_level || 'SENIOR_SECONDARY', class_level:a.class_level || 'SS 1', subject:a.subject, status:a.status,
      percent:total ? Math.round(score / total * 1000) / 10 : 0,
      score:score, total:total, question_count:Number(a.question_count) || 0,
      submitted_at:a.submitted_at ? Math.floor(Date.parse(a.submitted_at) / 1000) : null
    };
  })};
}

function historyDetail_(b) {
  const user = auth_(b);
  const id = required_(b.attemptId, 'Attempt ID', 80);
  const a = findOne_('PracticeAttempts', 'attempt_id', id);
  if (!a || a.user_id !== user.user_id || a.status !== 'submitted') throw new Error('Submitted practice attempt not found.');
  const ids = JSON.parse(a.question_ids_json);
  const qs = read_('PracticeQuestions');
  const saved = read_('PracticeAnswers').filter(x => x.attempt_id === id);
  const selected = {};
  saved.forEach(x => { selected[x.question_id] = x.selected_option; });
  const byId = {};
  qs.forEach(q => { byId[q.question_id] = q; });
  const review = ids.map(qid => byId[qid] ? safeQuestion_(byId[qid], true, selected[qid] || '') : null).filter(Boolean);
  const total = Number(a.total) || 0, score = Number(a.score) || 0;
  return {
    result:{
      id:a.attempt_id, exam_type:a.exam_type, subject:a.subject, score:score, total:total,
      percent:total ? Math.round(score / total * 1000) / 10 : 0,
      correct_count:review.filter(q => q.correct).length,
      question_count:Number(a.question_count) || review.length
    },
    review:review.map(q => ({
      id:q.id, year:q.year, topic:q.topic, question_text:q.questionText, image_url:q.imageUrl || '',
      options:q.options, selected_option:q.selectedOption, correct_option:q.correctOption,
      correct:q.correct, explanation:q.explanation
    }))
  };
}

// ------------------------- Administrator actions -------------------------

function requireAdmin_(b){
 const token=String(b.adminToken||'');if(token.length<40||token.length>200)throw Error('Administrator session expired. Please sign in again.');
 const s=findOne_('PracticeAdminSessions','token_hash',hash_(token));if(!s||s.revoked==='true'||Date.parse(s.expires_at)<=Date.now())throw Error('Administrator session expired. Please sign in again.');
 const a=findOne_('PracticeAdmins','admin_id',s.admin_id);if(!a||a.active!=='true')throw Error('Administrator account is inactive.');return a;
}
function adminLogin_(b){
 const username=required_(b.username,'Username',80).toLowerCase(),password=String(b.password||''),a=findOne_('PracticeAdmins','username',username);
 if(!a||a.active!=='true'||a.password_hash!==hash_(String(a.password_salt)+password))throw Error('Username or password is incorrect.');
 const token=randomToken_(),now=new Date().toISOString();append_('PracticeAdminSessions',{session_id:Utilities.getUuid(),token_hash:hash_(token),admin_id:a.admin_id,expires_at:new Date(Date.now()+VP.SESSION_MS).toISOString(),created_at:now,revoked:'false'});
 updateById_('PracticeAdmins','admin_id',a.admin_id,{last_login:now});return {adminToken:token,username:a.username};
}
function adminChangePassword_(b){
 const a=requireAdmin_(b),old=String(b.currentPassword||''),next=String(b.newPassword||'');
 if(a.password_hash!==hash_(String(a.password_salt)+old))throw Error('Current password is incorrect.');
 if(next.length<10||next.length>128)throw Error('New password must contain between 10 and 128 characters.');
 const salt=randomToken_();updateById_('PracticeAdmins','admin_id',a.admin_id,{password_salt:salt,password_hash:hash_(salt+next)});
 read_('PracticeAdminSessions').filter(x=>x.admin_id===a.admin_id&&x.revoked!=='true').forEach(x=>updateById_('PracticeAdminSessions','session_id',x.session_id,{revoked:'true'}));return {changed:true};
}
function requestPasswordReset_(b){
 const email=validEmail_(b.email),u=findOne_('PracticeUsers','email',email);
 if(u&&u.active==='true'){
  const token=randomToken_(),now=new Date(),expires=new Date(now.getTime()+30*60*1000);
  append_('PracticePasswordResets',{reset_id:Utilities.getUuid(),user_id:u.user_id,token_hash:hash_(token),expires_at:expires.toISOString(),used:'false',created_at:now.toISOString()});
  const base='https://vicora-practice-new.vercel.app';
  MailApp.sendEmail({to:u.email,subject:'VICORA Practice password reset',body:'Hello '+u.full_name+',\n\nUse this link to reset your VICORA Practice password. It expires in 30 minutes and can only be used once:\n\n'+base+'/?reset='+encodeURIComponent(token)+'\n\nIf you did not request this, ignore this email.'});
 }
 return {message:'If the email belongs to an active account, a password reset link has been sent.'};
}
function completePasswordReset_(b){
 const token=String(b.token||''),password=String(b.password||'');if(token.length<40||token.length>100)throw Error('This password reset link is invalid or expired.');
 if(password.length<10||password.length>128)throw Error('Password must contain between 10 and 128 characters.');
 const r=findOne_('PracticePasswordResets','token_hash',hash_(token));if(!r||r.used==='true'||Date.parse(r.expires_at)<=Date.now())throw Error('This password reset link is invalid or expired.');
 const u=findOne_('PracticeUsers','user_id',r.user_id);if(!u||u.active!=='true')throw Error('This account is unavailable.');
 const salt=randomToken_();updateById_('PracticeUsers','user_id',u.user_id,{password_salt:salt,password_hash:hash_(salt+password)});
 updateById_('PracticePasswordResets','reset_id',r.reset_id,{used:'true'});
 read_('PracticeSessions').filter(x=>x.user_id===u.user_id&&x.revoked!=='true').forEach(x=>updateById_('PracticeSessions','session_id',x.session_id,{revoked:'true'}));return {changed:true};
}

function listQuestions_() {
  const rows = read_('PracticeQuestions').sort((a,b) => String(b.updated_at || b.created_at).localeCompare(String(a.updated_at || a.created_at)));
  return {questions:rows.map(q => ({
    id:q.question_id, exam_type:q.exam_type, school_level:q.school_level || 'SENIOR_SECONDARY', class_level:q.class_level || '', subject:q.subject, year:q.year,
    topic:q.topic, difficulty:q.difficulty, question_text:q.question_text, image_url:q.image_url || '',
    option_a:q.option_a, option_b:q.option_b, option_c:q.option_c, option_d:q.option_d,
    correct_option:q.correct_option, explanation:q.explanation,
    marks:Number(q.marks) || 1, active:q.active === 'true'
  }))};
}

function importQuestions_(b) {
  const list = b.questions;
  if (!Array.isArray(list) || list.length < 1 || list.length > 100) throw new Error('Import between 1 and 100 questions per request.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let saved = 0;
    list.forEach(function(q) {
      const type = String(q.exam_type || '').toUpperCase();
      const schoolLevel = normalLevel_(q.school_level || 'SENIOR_SECONDARY');
      const classLevel = normalClass_(schoolLevel,q.class_level || defaultClass_(schoolLevel));
      const subject = required_(q.subject, 'Subject', 100);
      const questionText = required_(q.question_text, 'Question', 5000);
      const correct = String(q.correct_option || '').toUpperCase();
      if (VP.TYPES.indexOf(type) < 0 || ['A','B','C','D'].indexOf(correct) < 0) throw new Error('Each question needs a valid exam type and correct option.');
      const options = ['a','b','c','d'].map(k => required_(q['option_' + k], 'Option ' + k.toUpperCase(), 1500));
      const all = read_('PracticeQuestions');
      const existing = (q.question_id ? findOne_('PracticeQuestions','question_id',String(q.question_id)) : null) ||
        all.find(x => x.exam_type === type && x.subject === subject && (x.school_level || 'SENIOR_SECONDARY') === schoolLevel && (x.class_level || defaultClass_(schoolLevel)) === classLevel && x.question_text === questionText);
      const now = new Date().toISOString();
      const values = {
        exam_type:type, school_level:schoolLevel, class_level:classLevel, subject:subject, year:clean_(q.year || '',12),
        topic:clean_(q.topic || '',160), difficulty:clean_(q.difficulty || 'medium',30),
        question_text:questionText, image_url:clean_(q.image_url || '',2000), option_a:options[0], option_b:options[1],
        option_c:options[2], option_d:options[3], correct_option:correct,
        explanation:clean_(q.explanation || '',5000),
        marks:String(Math.max(0.01,Math.min(100,Number(q.marks) || 1))),
        active:'true', updated_at:now
      };
      if (existing) {
        updateById_('PracticeQuestions','question_id',existing.question_id,values);
      } else {
        values.question_id = Utilities.getUuid();
        values.created_at = now;
        append_('PracticeQuestions',values);
      }
      saved++;
    });
    return {saved:saved};
  } finally {
    lock.releaseLock();
  }
}


function listUsers_() {
  const attempts=read_('PracticeAttempts'),rows=read_('PracticeUsers').sort((a,b)=>String(b.created_at).localeCompare(String(a.created_at)));
  return {users:rows.map(u=>{const own=attempts.filter(a=>a.user_id===u.user_id);return {user_id:u.user_id,full_name:u.full_name,email:u.email,active:u.active==='true',created_at:u.created_at,last_login:u.last_login,school_level:u.school_level||'SENIOR_SECONDARY',class_level:u.class_level||'SS 1',attempts:own.length,completed:own.filter(a=>a.status==='submitted').length};})};
}
function setUserStatus_(b) {
  const id=required_(b.userId,'User ID',80),active=b.active===true||String(b.active)==='true';
  updateById_('PracticeUsers','user_id',id,{active:String(active)});
  if(!active)read_('PracticeSessions').filter(x=>x.user_id===id&&x.revoked!=='true').forEach(x=>updateById_('PracticeSessions','session_id',x.session_id,{revoked:'true'}));
  return {updated:true,active:active};
}
function uploadQuestionImage_(b) {
  const m=String(b.dataUrl||'').match(/^data:(image\/(?:png|jpeg|jpg|webp|gif));base64,([A-Za-z0-9+/=]+)$/i);
  if(!m)throw Error('Upload a PNG, JPG, WEBP or GIF image.');
  if(m[2].length>5*1024*1024)throw Error('Image is too large. Please use an image under 3.5 MB.');
  const p=PropertiesService.getScriptProperties();let fid=p.getProperty('VICORA_PRACTICE_IMAGE_FOLDER_ID');
  const folder=fid?DriveApp.getFolderById(fid):DriveApp.createFolder('VICORA Public Practice Question Images');
  if(!fid)p.setProperty('VICORA_PRACTICE_IMAGE_FOLDER_ID',folder.getId());
  const mime=m[1].toLowerCase()==='image/jpg'?'image/jpeg':m[1].toLowerCase();
  const file=folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),mime,clean_(b.fileName||'question-image',120)));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return {imageUrl:'https://drive.google.com/uc?export=view&id='+file.getId()};
}
const AI_PROVIDERS_=['gemini','groq','openai','meta','claude','grok'];
const AI_DEFAULT_MODELS_={gemini:'gemini-3.5-flash-lite',groq:'openai/gpt-oss-120b',openai:'gpt-4.1-mini',meta:'muse-spark-1.3',claude:'claude-sonnet-4-5',grok:'grok-4.7'};
function aiSettings_(){const p=PropertiesService.getScriptProperties(),providers={};AI_PROVIDERS_.forEach(k=>{providers[k]={configured:!!p.getProperty('VICORA_'+k.toUpperCase()+'_API_KEY'),model:p.getProperty('VICORA_'+k.toUpperCase()+'_MODEL')||AI_DEFAULT_MODELS_[k]};});return {providers:providers,providerOrder:p.getProperty('VICORA_AI_PROVIDER_ORDER')||'gemini,groq,openai,meta,claude,grok'};}
function saveAiSettings_(b){const p=PropertiesService.getScriptProperties();AI_PROVIDERS_.forEach(k=>{const prefix='VICORA_'+k.toUpperCase();const key=b[k+'Key'];if(key&&String(key).trim())p.setProperty(prefix+'_API_KEY',String(key).trim());if(b[k+'Model']!==undefined)p.setProperty(prefix+'_MODEL',validModelId_(b[k+'Model'],k+' model'));});if(b.providerOrder!==undefined){const order=String(b.providerOrder).toLowerCase().split(',').map(x=>x.trim()).filter(Boolean);if(order.length!==AI_PROVIDERS_.length||new Set(order).size!==AI_PROVIDERS_.length||order.some(x=>AI_PROVIDERS_.indexOf(x)<0))throw Error('Fallback order must contain each provider exactly once: gemini, groq, openai, meta, claude, grok.');p.setProperty('VICORA_AI_PROVIDER_ORDER',order.join(','));}return aiSettings_();}
function validModelId_(value,label){const id=String(value||'').trim();if(!id||id.length>120||!/^[A-Za-z0-9._:/-]+$/.test(id))throw Error('Enter a valid '+label+' model ID.');return id;}
function verifyQuestionWithAI_(b){const q=b.question||{};const prompt='Review this Nigerian curriculum question for the stated school level and class. Check factual accuracy, ambiguity, all options, proposed correct answer, explanation and image relevance. Do not assume supplied answer is correct. Return a concise structured report with verdict PASS/REVIEW/FAIL, answerCheck, issues, suggestedCorrection, explanationCheck and levelFit. Question: '+JSON.stringify(q);return callPracticeAI_(prompt,q.image_url||'',true);}
function studentTutor_(b){const u=auth_(b);if(read_('PracticeAttempts').some(a=>a.user_id===u.user_id&&a.status==='in_progress'))throw Error('AI Study is locked while you have an active examination. Submit your examination before asking AI for help.');const messages=cleanChatMessages_(b.messages,b.question);if(!messages.length)throw Error('Enter a learning question.');const prompt='You are VICORA AI Study Tutor. Teach a student in '+(u.school_level||'SENIOR_SECONDARY')+', class '+(u.class_level||'SS 1')+'. Support Primary, Junior Secondary and Senior Secondary learning with age-appropriate, curriculum-aligned explanations. Teach step by step, use simple examples, check understanding, and never shame the learner. If asked to solve a problem, show the method and final answer. Keep answers clear and useful. The student is studying independently, not taking an active examination. Conversation:\\n'+formatChatMessages_(messages);return callPracticeAI_(prompt,String(b.image_url||''),false);}
function adminAIChat_(b){const messages=cleanChatMessages_(b.messages,b.question);if(!messages.length)throw Error('Enter a message for VICORA Admin AI.');const prompt='You are VICORA Admin AI, an assistant for the authorised administrator of VICORA Public CBT Practice. Help with question-bank planning, Nigerian curriculum alignment, question generation, answer verification, explanations, test preparation, platform operations and interpreting any figures the administrator explicitly provides. Be precise, flag uncertainty, and never claim to have changed database records or settings. When generating questions, include school level, class, subject, topic, four options A-D, correct answer and explanation. Conversation:\\n'+formatChatMessages_(messages);return callPracticeAI_(prompt,'',true);}
function cleanChatMessages_(raw,fallback){let list=Array.isArray(raw)?raw:[];list=list.slice(-12).map(m=>({role:m&&m.role==='assistant'?'assistant':'user',content:clean_(m&&m.content,3000)})).filter(m=>m.content);if(!list.length&&fallback)list=[{role:'user',content:clean_(fallback,3000)}];return list;}
function formatChatMessages_(messages){return messages.map(m=>(m.role==='assistant'?'VICORA AI':'User')+': '+m.content).join('\\n');}
function callPracticeAI_(prompt,imageUrl,isAdmin){
 const p=PropertiesService.getScriptProperties(),order=(p.getProperty('VICORA_AI_PROVIDER_ORDER')||'gemini,groq,openai,meta,claude,grok').split(',').map(x=>x.trim()).filter(x=>AI_PROVIDERS_.indexOf(x)>=0);
 let last='No AI provider is configured.',attempted=[];
 for(const provider of order){
  const prefix='VICORA_'+provider.toUpperCase(),key=p.getProperty(prefix+'_API_KEY');
  if(!key)continue;
  attempted.push(provider);
  try{
   let answer='';
   if(provider==='gemini'){
    const parts=[{text:prompt}];
    if(imageUrl){const fid=(String(imageUrl).match(/[?&]id=([^&]+)/)||[])[1];if(fid){const blob=DriveApp.getFileById(fid).getBlob();parts.push({inlineData:{mimeType:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())}});}else parts.push({text:'Image URL: '+imageUrl});}
    const model=p.getProperty(prefix+'_MODEL')||AI_DEFAULT_MODELS_.gemini;
    const url='https://generativelanguage.googleapis.com/v1beta/models/'+encodeURIComponent(model)+':generateContent?key='+encodeURIComponent(key);
    const res=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify({contents:[{parts:parts}],generationConfig:{temperature:0.2,maxOutputTokens:1200}}),muteHttpExceptions:true});
    const data=parseAIJson_(res,provider);if(res.getResponseCode()>=300)throw Error(aiErrorMessage_(data,res,provider));
    answer=(data.candidates&&data.candidates[0]&&data.candidates[0].content&&data.candidates[0].content.parts||[]).map(x=>x.text||'').join('');
   }else if(provider==='claude'){
    const model=p.getProperty(prefix+'_MODEL')||AI_DEFAULT_MODELS_.claude;
    const res=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{method:'post',contentType:'application/json',headers:{'x-api-key':key,'anthropic-version':'2023-06-01'},payload:JSON.stringify({model:model,max_tokens:1200,messages:[{role:'user',content:prompt}]}),muteHttpExceptions:true});
    const data=parseAIJson_(res,provider);if(res.getResponseCode()>=300)throw Error(aiErrorMessage_(data,res,provider));
    answer=(data.content||[]).map(x=>x.text||'').join('');
   }else{
    const base=provider==='groq'?'https://api.groq.com/openai/v1':provider==='openai'?'https://api.openai.com/v1':provider==='meta'?'https://api.meta.ai/v1':'https://api.x.ai/v1';
    const model=p.getProperty(prefix+'_MODEL')||AI_DEFAULT_MODELS_[provider];
    const res=UrlFetchApp.fetch(base+'/chat/completions',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+key},payload:JSON.stringify({model:model,messages:[{role:'user',content:prompt}],temperature:0.2,max_tokens:1200}),muteHttpExceptions:true});
    const data=parseAIJson_(res,provider);if(res.getResponseCode()>=300)throw Error(aiErrorMessage_(data,res,provider));
    answer=data.choices&&data.choices[0]&&data.choices[0].message&&data.choices[0].message.content;
   }
   if(!answer)throw Error(provider+' returned an empty answer.');
   return {provider:provider,answer:String(answer)};
  }catch(e){last=provider+': '+safeError_(e);}
 }
 if(!attempted.length)throw Error('No AI provider is configured. Add at least one provider API key in VICORA Admin → AI model settings.');
 throw Error('All configured AI providers failed. '+last+' VICORA tried: '+attempted.join(' → ')+'. Check API keys, model IDs, quotas and provider availability.');
}
function parseAIJson_(res,provider){try{return JSON.parse(res.getContentText());}catch(e){throw Error(provider+' returned an invalid response (HTTP '+res.getResponseCode()+').');}}
function aiErrorMessage_(data,res,provider){const err=data&&(data.error||data);const msg=err&&(err.message||err.type)||'request failed';return provider+' API error HTTP '+res.getResponseCode()+': '+String(msg).slice(0,260);}
function normalLevel_(v){const x=String(v||'').toUpperCase().replace(/[ -]+/g,'_');if(VP.LEVELS.indexOf(x)<0)throw Error('Choose Primary, Junior Secondary or Senior Secondary.');return x;}
function defaultClass_(l){return l==='PRIMARY'?'Primary 1':l==='JUNIOR_SECONDARY'?'JSS 1':'SS 1';}
function normalClass_(l,v){const x=clean_(v,30),ok=l==='PRIMARY'?/^Primary [1-6]$/i.test(x):l==='JUNIOR_SECONDARY'?/^JSS [1-3]$/i.test(x):/^SS [1-3]$/i.test(x);if(!ok)throw Error('Choose a class that matches the selected school level.');return l==='PRIMARY'?x.replace(/^primary/i,'Primary'):x.toUpperCase();}

function authorizeVICORAPasswordResetEmail(){
  const quota=MailApp.getRemainingDailyQuota();
  Logger.log('VICORA password reset email authorization is ready. Remaining daily recipient quota: '+quota);
  return {authorized:true,remainingDailyQuota:quota};
}
function initializeVICORAAdminLogin(){
 const id=PropertiesService.getScriptProperties().getProperty(VP.DB_PROP);if(!id)throw Error('Existing practice database was not found. Do not create a new database.');
 const ss=SpreadsheetApp.openById(id),specs={PracticeAdmins:['admin_id','username','password_salt','password_hash','active','created_at','last_login'],PracticeAdminSessions:['session_id','token_hash','admin_id','expires_at','created_at','revoked'],PracticePasswordResets:['reset_id','user_id','token_hash','expires_at','used','created_at']};
 Object.keys(specs).forEach(n=>{let sh=ss.getSheetByName(n);if(!sh)sh=ss.insertSheet(n);if(sh.getLastRow()===0){sh.getRange(1,1,1,specs[n].length).setValues([specs[n]]);sh.setFrozenRows(1);sh.getRange(1,1,1,specs[n].length).setFontWeight('bold');}});
 const sh=ss.getSheetByName('PracticeAdmins');if(sh.getLastRow()===1)sh.appendRow(['VICORA-ADMIN-1','@sheggun2026','dd3621825a14ca6ad7db29a32c44cc9db34ed3ac1f19364f','651acf10de365d4f26342bf0862d5a4624707c6e170a0b83ae7d772561ffb913','true',new Date().toISOString(),'']);
 return {ready:true,spreadsheetUrl:ss.getUrl(),adminUsername:sh.getRange(2,2).getValue(),message:'Admin login initialized. Edit PracticeAdmins column B to change username; change password in the dashboard.'};
}

// ------------------------------- Utilities --------------------------------

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(VP.DB_PROP);
  if (!id) throw new Error('Database is not initialized. Run setupVICORAPublicPractice() first.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const sh = db_().getSheetByName(name);
  if (!sh) throw new Error('Database tab missing: ' + name + '. Run setupVICORAPublicPractice().');
  return sh;
}

function read_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(function(row) {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i] instanceof Date ? row[i].toISOString() : String(row[i] == null ? '' : row[i]); });
    return obj;
  });
}

function append_(name, obj) {
  const sh = sheet_(name);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''));
}

function findOne_(name, field, value) {
  const target = String(value).toLowerCase();
  return read_(name).find(row => String(row[field] == null ? '' : row[field]).toLowerCase() === target) || null;
}

function updateById_(name, field, value, updates) {
  const sh = sheet_(name);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const col = headers.indexOf(field);
  if (col < 0) throw new Error('Database column missing: ' + field);
  const found = sh.getRange(2,col+1,Math.max(sh.getLastRow()-1,0),1)
    .createTextFinder(String(value)).matchEntireCell(true).findNext();
  if (!found) throw new Error('Record not found.');
  const rowNo = found.getRow();
  const range = sh.getRange(rowNo,1,1,headers.length);
  const row = range.getValues()[0];
  headers.forEach((h,i) => { if (Object.prototype.hasOwnProperty.call(updates,h)) row[i] = updates[h]; });
  range.setValues([row]);
}

function validEmail_(value) {
  const email = required_(value,'Email',190).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address.');
  return email;
}

function required_(value,label,max) {
  const out = clean_(value,max);
  if (!out) throw new Error(label + ' is required.');
  return out;
}

function clean_(value,max) {
  return String(value == null ? '' : value).replace(/[\u0000-\u001F\u007F]/g,' ').trim().slice(0,max || 500);
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
}

function hash_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value),Utilities.Charset.UTF_8);
  return bytes.map(b => {
    const n = b < 0 ? b + 256 : b;
    return ('0' + n.toString(16)).slice(-2);
  }).join('');
}

function safeError_(err) {
  return String(err && err.message ? err.message : err || 'Unknown error.').slice(0,400);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

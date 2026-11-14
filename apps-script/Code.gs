/**
 * VICORA PUBLIC PRACTICE — standalone Google Apps Script backend
 * Google Sheets database; independent of the VICORA School CBT backend.
 *
 * FIRST RUN:
 * 1. Create a NEW Apps Script project.
 * 2. Replace Code.gs with this complete file and save.
 * 3. Select setupVICORAPublicPractice and click Run; approve permissions.
 * 4. Open the execution log. Save the generated administrator key securely.
 * 5. Deploy > New deployment > Web app; Execute as: Me; access: Anyone.
 * 6. Share the deployed /exec URL with students and administrators.
 *
 * The setup function is safe to run again: it reuses the same database and key.
 * It does not overwrite existing records.
 */

const VP = Object.freeze({
  VERSION: '1.0.0',
  DB_PROP: 'VICORA_PRACTICE_DATABASE_ID',
  ADMIN_PROP: 'VICORA_PRACTICE_ADMIN_KEY',
  SESSION_MS: 30 * 24 * 60 * 60 * 1000,
  TYPES: ['WAEC', 'NECO', 'JAMB', 'GENERAL'],
  SHEETS: {
    PracticeUsers: ['user_id','full_name','email','password_salt','password_hash','active','created_at','last_login'],
    PracticeSessions: ['session_id','token_hash','user_id','expires_at','created_at','revoked'],
    PracticeQuestions: ['question_id','exam_type','subject','year','topic','difficulty','question_text','option_a','option_b','option_c','option_d','correct_option','explanation','marks','active','created_at','updated_at'],
    PracticeAttempts: ['attempt_id','user_id','exam_type','subject','question_ids_json','question_count','started_at','submitted_at','score','total','status'],
    PracticeAnswers: ['answer_id','attempt_id','question_id','selected_option']
  }
});

/**
 * Run this function once from the Apps Script editor.
 * Creates a dedicated spreadsheet and all tabs, and generates the admin key.
 * The key is saved in Script Properties and returned/logged only during setup.
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

  let adminKey = props.getProperty(VP.ADMIN_PROP);
  let keyWasCreated = false;
  if (!adminKey || adminKey.length < 40) {
    adminKey = 'VICORA-PRAC-' + randomToken_() + randomToken_().slice(0, 12);
    props.setProperty(VP.ADMIN_PROP, adminKey);
    keyWasCreated = true;
  }

  const result = {
    status: 'READY',
    message: 'VICORA Public Practice database is ready.',
    spreadsheetUrl: ss.getUrl(),
    spreadsheetId: dbId,
    tabs: Object.keys(VP.SHEETS),
    administratorKey: adminKey,
    administratorKeyWasGeneratedNow: keyWasCreated,
    nextSteps: [
      'Save the administrator key securely.',
      'Deploy this project as a Web app, executing as you, accessible to anyone.',
      'Use the deployed /exec URL as the student and administrator website.'
    ]
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function doGet(e) {
  // The deployed Apps Script URL serves the complete student/admin web app.
  // Optional health check: append ?api=health
  if (e && e.parameter && e.parameter.api === 'health') {
    return json_({ok:true, service:'VICORA Public Practice', version:VP.VERSION, timestamp:new Date().toISOString()});
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('VICORA Public Practice')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = String(body.action || '');
    let data;
    switch (action) {
      case 'practice.register': data = register_(body); break;
      case 'practice.login': data = login_(body); break;
      case 'practice.logout': data = logout_(body); break;
      case 'practice.me': data = me_(body); break;
      case 'practice.catalog': data = catalog_(); break;
      case 'practice.history': data = history_(body); break;
      case 'practice.history.detail': data = historyDetail_(body); break;
      case 'practice.start': data = start_(body); break;
      case 'practice.submit': data = submit_(body); break;
      case 'practice.admin.listQuestions': requireAdmin_(body); data = listQuestions_(); break;
      case 'practice.admin.importQuestions': requireAdmin_(body); data = importQuestions_(body); break;
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
      password_hash:hash_(salt + password), active:'true', created_at:now, last_login:now
    });
    append_('PracticeSessions', {
      session_id:Utilities.getUuid(), token_hash:hash_(token), user_id:id,
      expires_at:new Date(Date.now() + VP.SESSION_MS).toISOString(), created_at:now, revoked:'false'
    });
    return {sessionToken:token, user:{userId:id, fullName:name, email:email}};
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
  return {sessionToken:token, user:{userId:user.user_id, fullName:user.full_name, email:user.email}};
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
  return {user:{userId:u.user_id, fullName:u.full_name, email:u.email}};
}

function logout_(b) {
  const token = String(b.sessionToken || '');
  if (token) {
    const s = findOne_('PracticeSessions', 'token_hash', hash_(token));
    if (s) updateById_('PracticeSessions', 'session_id', s.session_id, {revoked:'true'});
  }
  return {loggedOut:true};
}

// ---------------------------- Question bank -------------------------------

function catalog_() {
  const rows = read_('PracticeQuestions').filter(q => q.active === 'true');
  const subjects = {};
  const counts = {};
  rows.forEach(function(q) {
    if (!subjects[q.exam_type]) subjects[q.exam_type] = [];
    if (subjects[q.exam_type].indexOf(q.subject) < 0) subjects[q.exam_type].push(q.subject);
    counts[q.exam_type + '|' + q.subject] = (counts[q.exam_type + '|' + q.subject] || 0) + 1;
  });
  Object.keys(subjects).forEach(k => subjects[k].sort());
  return {examTypes:VP.TYPES, subjects:subjects, counts:counts};
}

function start_(b) {
  const user = auth_(b);
  const type = String(b.examType || '').toUpperCase();
  const subject = required_(b.subject, 'Subject', 100);
  const count = Math.max(5, Math.min(60, Number(b.questionCount) || 20));
  if (VP.TYPES.indexOf(type) < 0) throw new Error('Choose a valid examination type.');
  const available = read_('PracticeQuestions').filter(q => q.active === 'true' && q.exam_type === type && q.subject === subject);
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
  append_('PracticeAttempts', {
    attempt_id:attemptId, user_id:user.user_id, exam_type:type, subject:subject,
    question_ids_json:JSON.stringify(ids), question_count:ids.length,
    started_at:now, submitted_at:'', score:'', total:'', status:'in_progress'
  });
  return {attemptId:attemptId, examType:type, subject:subject, questionCount:ids.length, questions:chosen.map(q => safeQuestion_(q, false, ''))};
}

function safeQuestion_(q, includeAnswer, selected) {
  const out = {
    id:q.question_id, examType:q.exam_type, subject:q.subject, year:q.year,
    topic:q.topic, difficulty:q.difficulty, questionText:q.question_text,
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
    const ids = JSON.parse(attempt.question_ids_json);
    const questions = read_('PracticeQuestions');
    const byId = {};
    questions.forEach(q => { byId[q.question_id] = q; });
    const answers = b.answers && typeof b.answers === 'object' ? b.answers : {};
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
      submitted_at:new Date().toISOString(), score:String(score), total:String(total), status:'submitted'
    });
    const review = ids.map(qid => {
      const q = byId[qid];
      const selected = String(answers[qid] || '').toUpperCase();
      return q ? safeQuestion_(q, true, ['A','B','C','D'].indexOf(selected) >= 0 ? selected : '') : null;
    }).filter(Boolean);
    return {
      attemptId:id, examType:attempt.exam_type, subject:attempt.subject,
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
      id:a.attempt_id, exam_type:a.exam_type, subject:a.subject, status:a.status,
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
      id:q.id, year:q.year, topic:q.topic, question_text:q.questionText,
      options:q.options, selected_option:q.selectedOption, correct_option:q.correctOption,
      correct:q.correct, explanation:q.explanation
    }))
  };
}

// ------------------------- Administrator actions -------------------------

function requireAdmin_(b) {
  const expected = PropertiesService.getScriptProperties().getProperty(VP.ADMIN_PROP) || '';
  const supplied = String(b.adminKey || '');
  if (expected.length < 40 || supplied.length < 40 || hash_(supplied) !== hash_(expected)) {
    throw new Error('Not authorised. Check the VICORA Practice Administrator Key.');
  }
}

function listQuestions_() {
  const rows = read_('PracticeQuestions').sort((a,b) => String(b.updated_at || b.created_at).localeCompare(String(a.updated_at || a.created_at)));
  return {questions:rows.map(q => ({
    id:q.question_id, exam_type:q.exam_type, subject:q.subject, year:q.year,
    topic:q.topic, difficulty:q.difficulty, question_text:q.question_text,
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
      const subject = required_(q.subject, 'Subject', 100);
      const questionText = required_(q.question_text, 'Question', 5000);
      const correct = String(q.correct_option || '').toUpperCase();
      if (VP.TYPES.indexOf(type) < 0 || ['A','B','C','D'].indexOf(correct) < 0) throw new Error('Each question needs a valid exam type and correct option.');
      const options = ['a','b','c','d'].map(k => required_(q['option_' + k], 'Option ' + k.toUpperCase(), 1500));
      const all = read_('PracticeQuestions');
      const existing = (q.question_id ? findOne_('PracticeQuestions','question_id',String(q.question_id)) : null) ||
        all.find(x => x.exam_type === type && x.subject === subject && x.question_text === questionText);
      const now = new Date().toISOString();
      const values = {
        exam_type:type, subject:subject, year:clean_(q.year || '',12),
        topic:clean_(q.topic || '',160), difficulty:clean_(q.difficulty || 'medium',30),
        question_text:questionText, option_a:options[0], option_b:options[1],
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


// ----------------------- Apps Script web-app interface --------------------
// These public wrappers are called by Index.html through google.script.run.
// The underlying validation and database logic remains in the functions above.

function uiRegister(payload) { return register_(payload || {}); }
function uiLogin(payload) { return login_(payload || {}); }
function uiLogout(payload) { return logout_(payload || {}); }
function uiMe(payload) { return me_(payload || {}); }
function uiCatalog() { return catalog_(); }
function uiHistory(payload) { return history_(payload || {}); }
function uiStart(payload) { return start_(payload || {}); }
function uiSubmit(payload) { return submit_(payload || {}); }

function uiAdminListQuestions(payload) {
  requireAdmin_(payload || {});
  return listQuestions_();
}

function uiAdminImportQuestions(payload) {
  payload = payload || {};
  requireAdmin_(payload);
  return importQuestions_(payload);
}

/**
 * VICORA PUBLIC PRACTICE — backend-only Google Apps Script API.
 * The website HTML is hosted separately. This project serves JSON only.
 *
 * Keep the existing Script Properties and spreadsheet database.
 * After updating this file in Apps Script, deploy a NEW web-app version.
 */

const VP = Object.freeze({
  VERSION: '1.0.1',
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
 * Existing setup function and all database/authentication/question functions
 * are retained below. Do not run setup again just to deploy this API change.
 */

function doGet(e) {
  return json_({
    ok: true,
    service: 'VICORA Public Practice API',
    version: VP.VERSION,
    mode: 'backend-only',
    timestamp: new Date().toISOString()
  });
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


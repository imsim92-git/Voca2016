/**
 * 우리반 단어장 — 기록 서버 (Google Apps Script)
 *
 * 이 코드는 선생님의 구글 스프레드시트에 붙어서 동작합니다.
 *  - '학생' 시트: 학번, 이름, 가입·최근 접속 시각 (비밀번호는 원문이 아니라 해시로만 저장)
 *  - '기록' 시트: 누가, 언제, 어느 학습지를, 어떤 연습으로, 몇 점 받았는지, 틀린 단어
 *  - '지문' 시트: 선생님이 등록한 지문 목록(교과서, 모의고사 1번 …)과 그 단어
 *
 * 단어표 사진 인식을 쓰려면: 편집기 왼쪽 '서비스' 옆 + → Drive API → 추가 (안내서 참고)
 *
 * 처음 한 번: 아래 ADMIN_PASSWORD_INIT에 관리자 비밀번호를 적고 setup 함수를 실행하세요.
 * 실행이 끝나면 비밀번호는 안전하게 저장되므로, 여기 적은 글자는 지워도 됩니다.
 */
const ADMIN_PASSWORD_INIT = '';

const SHEET_STUDENTS = '학생';
const SHEET_LOG = '기록';
const STUDENT_HEADERS = ['학번', '이름', '가입 시각', '최근 접속', 'salt', '비밀번호 해시'];
const LOG_HEADERS = ['시각', '학번', '이름', '학습지', '연습', '점수', '문항 수', '정답률', '틀린 단어', '페이지 ID'];
const SHEET_SETS = '지문';
const SET_HEADERS = ['ID', '이름', '순서', '단어 수', '단어(JSON)', '만든 시각', '고친 시각'];
const TOKEN_DAYS = 180;          // 학생 로그인 유지 기간
const ADMIN_TOKEN_HOURS = 12;    // 관리자 로그인 유지 시간
const MAX_FAILS = 8;             // 비밀번호를 이만큼 틀리면
const LOCK_MINUTES = 10;         // 이 시간 동안 로그인 잠금
const HASH_ROUNDS = 300;

/* ---------------- 설치 ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_STUDENTS, STUDENT_HEADERS);
  ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
  ensureSheet_(ss, SHEET_SETS, SET_HEADERS);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (ADMIN_PASSWORD_INIT) {
    const salt = Utilities.getUuid();
    props.setProperty('ADMIN_SALT', salt);
    props.setProperty('ADMIN_HASH', hashPw_(ADMIN_PASSWORD_INIT, salt));
  }
  if (!props.getProperty('ADMIN_HASH')) throw new Error('맨 위 ADMIN_PASSWORD_INIT에 관리자 비밀번호를 적고 다시 실행해 주세요.');
  return '설치 완료';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (name === SHEET_STUDENTS) sh.getRange('A:A').setNumberFormat('@'); // 학번 앞자리 0 보존
  return sh;
}

/* ---------------- 웹 앱 입구 ---------------- */
function doGet() {
  return json_({ok: true, app: 'word-practice'});
}

function doPost(e) {
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ok: false, error: '요청 형식이 올바르지 않아요.'}); }
  try {
    const fn = ACTIONS[req.action];
    if (!fn) return json_({ok: false, error: '알 수 없는 요청이에요.'});
    return json_(fn(req));
  } catch (err) {
    return json_({ok: false, code: (err && err.code) || '', error: (err && err.userMessage) || '서버에서 문제가 생겼어요. 잠시 후 다시 해 주세요.', detail: String(err && err.message || err)});
  }
}

const ACTIONS = {
  register: register_,
  login: login_,
  log: log_,
  me: me_,
  adminLogin: adminLogin_,
  admin: admin_,
  adminStudent: adminStudent_,
  adminReset: adminReset_,
  sets: sets_,
  adminSetSave: adminSetSave_,
  adminSetDelete: adminSetDelete_,
  adminSetOrder: adminSetOrder_,
  adminOcr: adminOcr_
};

/* ---------------- 학생 ---------------- */
function register_(req) {
  const sid = cleanSid_(req.sid), name = cleanName_(req.name), pw = cleanPw_(req.pw);
  return withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS);
    const row = findStudentRow_(sh, sid);
    if (row) {
      const hash = sh.getRange(row, 6).getValue();
      if (hash) fail_('이미 등록된 학번이에요. 로그인해 주세요. 비밀번호가 기억나지 않으면 선생님께 초기화를 부탁하세요.');
      // 선생님이 비밀번호를 초기화한 학생: 새 비밀번호로 다시 등록
      const salt = Utilities.getUuid(), now = new Date();
      sh.getRange(row, 2, 1, 5).setValues([[asText_(name), sh.getRange(row, 3).getValue() || now, now, asText_(salt), asText_(hashPw_(pw, salt))]]);
    } else {
      const salt = Utilities.getUuid(), now = new Date();
      sh.appendRow([asText_(sid), asText_(name), now, now, asText_(salt), asText_(hashPw_(pw, salt))]);
    }
    // 등록은 '기록' 탭에 남기지 않음 ('학생' 탭의 가입 시각으로 확인)
    return {ok: true, token: makeToken_('s', sid, TOKEN_DAYS * 24), student: {sid: sid, name: name}};
  });
}

function login_(req) {
  const sid = cleanSid_(req.sid), pw = String(req.pw || '');
  const cache = CacheService.getScriptCache(), key = 'fail_' + sid;
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) fail_(`비밀번호를 여러 번 틀려서 ${LOCK_MINUTES}분 동안 로그인할 수 없어요.`);
  const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
  if (!row) fail_('등록되지 않은 학번이에요. 처음이라면 "처음 등록"을 눌러 주세요.');
  const [, name, , , salt, hash] = sh.getRange(row, 1, 1, 6).getValues()[0];
  if (!hash) fail_('비밀번호가 초기화되었어요. "처음 등록"에서 새 비밀번호를 정해 주세요.');
  if (hashPw_(pw, salt) !== hash) {
    cache.put(key, String(fails + 1), LOCK_MINUTES * 60);
    fail_('학번이나 비밀번호가 맞지 않아요.');
  }
  cache.remove(key);
  sh.getRange(row, 4).setValue(new Date());   // 로그인은 '학생' 탭의 최근 접속만 갱신
  return {ok: true, token: makeToken_('s', sid, TOKEN_DAYS * 24), student: {sid: sid, name: String(name)}};
}

function log_(req) {
  const sid = checkToken_(req.token, 's');
  const st = studentInfo_(sid);
  const score = Math.max(0, Math.floor(Number(req.score) || 0));
  const total = Math.max(0, Math.floor(Number(req.total) || 0));
  const wrong = (Array.isArray(req.wrong) ? req.wrong : []).slice(0, 200).map(w => String(w).slice(0, 80).replace(/ ; /g, ' ')).filter(Boolean);
  const when = req.at ? new Date(req.at) : new Date();
  appendLog_([
    isNaN(when) ? new Date() : when, sid, st.name,
    String(req.pageTitle || '').slice(0, 60), String(req.mode || '').slice(0, 20),
    score, total, total ? Math.round(score / total * 100) / 100 : '',
    wrong.join(' ; '), String(req.page || '').slice(0, 30)
  ]);
  return {ok: true};
}

function me_(req) {
  const sid = checkToken_(req.token, 's');
  return Object.assign({ok: true, student: studentInfo_(sid)}, statsFor_(sid));
}

/* ---------------- 관리자 ---------------- */
function adminLogin_(req) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache(), key = 'fail_admin';
  if (Number(cache.get(key) || 0) >= MAX_FAILS) fail_(`비밀번호를 여러 번 틀려서 ${LOCK_MINUTES}분 동안 잠겼어요.`);
  if (hashPw_(String(req.pw || ''), props.getProperty('ADMIN_SALT')) !== props.getProperty('ADMIN_HASH')) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCK_MINUTES * 60);
    fail_('관리자 비밀번호가 맞지 않아요.');
  }
  cache.remove(key);
  return {ok: true, token: makeToken_('a', 'admin', ADMIN_TOKEN_HOURS)};
}

function admin_(req) {
  checkToken_(req.token, 'a');
  const students = readStudents_();
  const logs = readLogs_();
  const per = {};
  students.forEach(s => per[s.sid] = {sid: s.sid, name: s.name, joined: iso_(s.joined), lastSeen: iso_(s.lastSeen), hasPw: s.hasPw, sessions: 0, questions: 0, correct: 0, lastPractice: ''});
  const wrongCount = {};
  logs.forEach(l => {
    if (!isPractice_(l)) return;
    const p = per[l.sid] || (per[l.sid] = {sid: l.sid, name: l.name, sessions: 0, questions: 0, correct: 0, lastPractice: ''});
    p.sessions++; p.questions += l.total; p.correct += l.score;
    if (!p.lastPractice || l.at > p.lastPractice) p.lastPractice = l.at;
    l.wrong.forEach(w => { const k = l.page + '|' + w; (wrongCount[k] = wrongCount[k] || {en: w, page: l.page, count: 0, students: {}}).count++; wrongCount[k].students[l.sid] = 1; });
  });
  const wrongTop = Object.keys(wrongCount).map(k => wrongCount[k]).map(w => ({en: w.en, page: w.page, count: w.count, students: Object.keys(w.students).length}))
    .sort((a, b) => b.count - a.count).slice(0, 50);
  return {
    ok: true,
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    students: Object.keys(per).map(k => per[k]).sort((a, b) => String(a.sid).localeCompare(String(b.sid))),
    recent: logs.slice(-300).reverse(),
    wrongTop: wrongTop
  };
}

function adminStudent_(req) {
  checkToken_(req.token, 'a');
  const sid = cleanSid_(req.sid);
  return Object.assign({ok: true, student: studentInfo_(sid)}, statsFor_(sid));
}

function adminReset_(req) {
  checkToken_(req.token, 'a');
  const sid = cleanSid_(req.sid);
  return withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
    if (!row) fail_('그 학번의 학생이 없어요.');
    sh.getRange(row, 5, 1, 2).setValues([['', '']]);
    CacheService.getScriptCache().remove('fail_' + sid);
    return {ok: true};
  });
}

/* ---------------- 통계 ---------------- */
function statsFor_(sid) {
  const logs = readLogs_().filter(l => l.sid === sid && isPractice_(l));
  const wrong = {};
  let questions = 0, correct = 0;
  logs.forEach(l => {
    questions += l.total; correct += l.score;
    l.wrong.forEach(w => {
      const k = l.page + '|' + w;
      const o = wrong[k] || (wrong[k] = {en: w, page: l.page, count: 0, last: ''});
      o.count++; if (l.at > o.last) o.last = l.at;
    });
  });
  return {
    stats: {sessions: logs.length, questions: questions, correct: correct},
    recent: logs.slice(-30).reverse(),
    wrong: Object.keys(wrong).map(k => wrong[k]).sort((a, b) => b.count - a.count || (b.last > a.last ? 1 : -1)).slice(0, 60)
  };
}
function isPractice_(l) { return l.total > 0; }


/* ---------------- 지문 목록 (학생은 읽기만, 선생님만 등록·수정) ---------------- */
function readSets_() {
  const sh = sheet_(SHEET_SETS), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, SET_HEADERS.length).getValues().map((r, i) => {
    let words = []; try { words = JSON.parse(String(r[4] || '[]')); } catch (e) {}
    return {row: i + 2, id: String(r[0]).trim(), name: String(r[1]), order: Number(r[2]) || 0, words: words, updated: iso_(r[6])};
  }).filter(x => x.id).sort((a, b) => a.order - b.order || a.row - b.row);
}
function sets_() {
  return {ok: true, sets: readSets_().map(x => ({id: x.id, name: x.name, words: x.words, updated: x.updated}))};
}
function cleanWords_(list) {
  if (!Array.isArray(list)) fail_('단어 목록 형식이 올바르지 않아요.');
  const out = [];
  list.forEach(w => {
    if (!Array.isArray(w)) return;
    const en = String(w[0] || '').trim().slice(0, 80), ko = String(w[1] || '').trim().slice(0, 120), note = String(w[2] || '').trim().slice(0, 160);
    if (en && ko) out.push(note ? [en, ko, note] : [en, ko]);
  });
  if (out.length > 600) fail_('한 지문에 단어는 600개까지 넣을 수 있어요.');
  return out;
}
function adminSetSave_(req) {
  checkToken_(req.token, 'a');
  const name = String(req.name || '').trim().replace(/\s+/g, ' ');
  if (!name || name.length > 40) fail_('지문 이름을 1~40자로 써 주세요.');
  const words = cleanWords_(req.words || []);
  const json = JSON.stringify(words);
  if (json.length > 45000) fail_('단어가 너무 많아요. 지문을 둘로 나눠 주세요.');
  return withLock_(() => {
    const sh = sheet_(SHEET_SETS), now = new Date(), list = readSets_();
    const cur = req.id ? list.find(x => x.id === String(req.id)) : null;
    if (req.id && !cur) fail_('그 지문을 찾을 수 없어요. 새로고침해 주세요.');
    if (cur) {
      sh.getRange(cur.row, 2, 1, 4).setValues([[asText_(name), cur.order, words.length, json]]);
      sh.getRange(cur.row, 7).setValue(now);
      return {ok: true, id: cur.id};
    }
    const id = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 7);
    const order = list.reduce((m, x) => Math.max(m, x.order), 0) + 1;
    sh.appendRow([id, asText_(name), order, words.length, json, now, now]);
    return {ok: true, id: id};
  });
}
function adminSetDelete_(req) {
  checkToken_(req.token, 'a');
  return withLock_(() => {
    const cur = readSets_().find(x => x.id === String(req.id || ''));
    if (!cur) fail_('그 지문을 찾을 수 없어요.');
    sheet_(SHEET_SETS).deleteRow(cur.row);
    return {ok: true};
  });
}
function adminSetOrder_(req) {
  checkToken_(req.token, 'a');
  const ids = (Array.isArray(req.ids) ? req.ids : []).map(String);
  return withLock_(() => {
    const sh = sheet_(SHEET_SETS);
    readSets_().forEach(x => { const k = ids.indexOf(x.id); sh.getRange(x.row, 3).setValue(k < 0 ? ids.length + x.order : k + 1); });
    return {ok: true};
  });
}

/* 단어표 사진 → 글자: 선생님 구글 드라이브의 무료 글자 인식(OCR)을 쓰고, 임시 문서는 바로 휴지통으로 */
function adminOcr_(req) {
  checkToken_(req.token, 'a');
  const b64 = String(req.data || '');
  if (!b64) fail_('사진이 없어요.');
  if (b64.length > 14000000) fail_('사진이 너무 커요. 10MB보다 작은 사진으로 올려 주세요.');
  if (typeof Drive === 'undefined') fail_("사진 인식을 쓰려면 Apps Script 편집기 왼쪽 '서비스' 옆 + → Drive API → 추가를 한 뒤, 새 버전으로 다시 배포해 주세요.", 'no_drive');
  const mime = /^image\/(png|jpeg|gif|bmp|webp)$/.test(String(req.mime)) ? String(req.mime) : 'image/jpeg';
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, 'word-photo');
  let id = '';
  try {
    const meta = {mimeType: MimeType.GOOGLE_DOCS};
    if (Drive.Files.create) { meta.name = '단어장 사진 인식 (임시)'; id = Drive.Files.create(meta, blob, {ocrLanguage: 'ko'}).id; }
    else { meta.title = '단어장 사진 인식 (임시)'; id = Drive.Files.insert(meta, blob, {ocr: true, ocrLanguage: 'ko'}).id; }
    const text = DocumentApp.openById(id).getBody().getText();
    return {ok: true, text: text};
  } catch (err) {
    if (err && err.userMessage) throw err;
    fail_('사진에서 글자를 읽지 못했어요. 더 밝고 반듯하게 찍은 사진으로 다시 해 주세요.');
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
}

/* ---------------- 시트 읽기·쓰기 ---------------- */
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ensureSheet_(ss, name, name === SHEET_STUDENTS ? STUDENT_HEADERS : name === SHEET_SETS ? SET_HEADERS : LOG_HEADERS);
}
function findStudentRow_(sh, sid) {
  const n = sh.getLastRow(); if (n < 2) return 0;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === sid) return i + 2;
  return 0;
}
function studentInfo_(sid) {
  const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
  if (!row) fail_('학생 정보를 찾을 수 없어요. 다시 로그인해 주세요.', 'auth');
  return {sid: sid, name: String(sh.getRange(row, 2).getValue())};
}
function readStudents_() {
  const sh = sheet_(SHEET_STUDENTS), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 6).getValues().map(r => ({sid: String(r[0]).trim(), name: String(r[1]), joined: r[2], lastSeen: r[3], hasPw: !!r[5]})).filter(s => s.sid);
}
function readLogs_() {
  const sh = sheet_(SHEET_LOG), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, LOG_HEADERS.length).getValues().map(r => ({
    at: iso_(r[0]), sid: String(r[1]).trim(), name: String(r[2]), pageTitle: String(r[3]), mode: String(r[4]),
    score: Number(r[5]) || 0, total: Number(r[6]) || 0, wrong: String(r[8] || '').split(' ; ').filter(Boolean), page: String(r[9] || '')
  }));
}
function appendLog_(row) {
  const sh = sheet_(SHEET_LOG);
  withLock_(() => sh.appendRow(row.map(v => typeof v === 'string' ? asText_(v) : v)));
}
// 시트가 학번 앞자리 0을 지우거나 '='로 시작하는 글을 수식으로 바꾸지 않도록 글자로 고정
function asText_(v) { v = String(v); return /^[0-9=+\-@]/.test(v) ? "'" + v : v; }

/* ---------------- 보안 도구 ---------------- */
function hashPw_(pw, salt) {
  let h = String(salt) + ':' + pw;
  for (let i = 0; i < HASH_ROUNDS; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + pw, Utilities.Charset.UTF_8));
  return h;
}
function sign_(s) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) fail_('서버 설치가 끝나지 않았어요. 선생님께 알려 주세요.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret)).replace(/=+$/, '');
}
function makeToken_(kind, id, hours) {
  const body = [kind, Utilities.base64EncodeWebSafe(id, Utilities.Charset.UTF_8).replace(/=+$/, ''), Date.now() + hours * 3600 * 1000].join('.');
  return body + '.' + sign_(body);
}
function checkToken_(token, kind) {
  const p = String(token || '').split('.');
  if (p.length !== 4 || p[0] !== kind) fail_('다시 로그인해 주세요.', 'auth');
  const body = p.slice(0, 3).join('.');
  if (sign_(body) !== p[3] || Number(p[2]) < Date.now()) fail_('로그인이 만료되었어요. 다시 로그인해 주세요.', 'auth');
  let id = p[1]; while (id.length % 4) id += '=';
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(id)).getDataAsString('UTF-8');
}
function cleanSid_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!/^[0-9A-Za-z-]{2,20}$/.test(s)) fail_('학번은 숫자(또는 영문)로 2~20자로 써 주세요.');
  return s;
}
function cleanName_(v) {
  const s = String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
  if (s.length < 1 || s.length > 20 || /^[=+\-@]/.test(s)) fail_('이름을 1~20자로 써 주세요.');
  return s;
}
function cleanPw_(v) {
  const s = String(v == null ? '' : v);
  if (s.length < 4 || s.length > 40) fail_('비밀번호는 4~40자로 정해 주세요.');
  return s;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (lock.hasLock()) return fn();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function fail_(msg, code) { const e = new Error(msg); e.userMessage = msg; e.code = code; throw e; }
function iso_(d) { return d instanceof Date ? d.toISOString() : (d ? String(d) : ''); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

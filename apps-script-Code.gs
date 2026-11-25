/**
 * 김아근의 수업자료일람 — 저장 서버 (Google Apps Script)
 *
 * 하는 일
 *  - 사이트가 목록을 읽어 갈 때(GET) 시트에 저장된 목록을 돌려줍니다. 누구나 읽을 수 있습니다.
 *  - 저장(POST)은 관리자 비밀번호가 맞을 때만 받습니다. 비밀번호 확인은 이 서버(구글)에서 합니다.
 *  - 비밀번호는 시트 메뉴 [수업자료일람 > 관리자 비밀번호 설정]에서 정합니다.
 *    코드나 시트에는 비밀번호 대신 해시(되돌릴 수 없는 암호화 값)만 남습니다.
 *  - 10분 안에 5번 틀리면 10분 동안 로그인과 저장을 막습니다.
 *
 * 시트 구성
 *  - "_data": 사이트가 쓰는 원본(숨김). 직접 고치지 마세요.
 *  - "목록": 저장할 때마다 새로 써지는 읽기용 표(백업·확인용). 여기서 고쳐도 사이트에는 반영되지 않습니다.
 */

const DATA_SHEET = '_data';
const LIST_SHEET = '목록';
const CHUNK = 40000;          // 한 칸에 넣는 글자 수(시트 한 칸 한도 50,000자보다 작게)
const MARK = '~';             // 원본 칸 앞에 붙이는 표시: 시트가 숫자나 수식으로 바꿔 읽지 않게 합니다
const MAX_CHARS = 200000;     // 목록 전체 최대 크기
const MAX_FAILS = 5;
const WINDOW_MS = 10 * 60 * 1000;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('수업자료일람')
    .addItem('관리자 비밀번호 설정', 'setAdminPassword')
    .addItem('로그인 잠금 풀기', 'clearLock')
    .addToUi();
}

function setAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('관리자 비밀번호 설정', '사이트에서 쓸 관리자 비밀번호를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText().trim();
  if (pw.length < 4) { ui.alert('비밀번호는 4자 이상으로 입력하세요.'); return; }
  savePassword_(pw);
  ui.alert('저장했습니다. 이제 사이트에서 이 비밀번호로 관리자 로그인을 할 수 있습니다.');
}

function clearLock() {
  PropertiesService.getScriptProperties().setProperty('fails', '[]');
  SpreadsheetApp.getUi().alert('로그인 잠금을 풀었습니다.');
}

function savePassword_(pw) {
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ pwSalt: salt, pwHash: hash_(salt + pw), fails: '[]' });
}

function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 읽기: 누구나
function doGet() {
  return json_({ ok: true, data: readData_() });
}

// 쓰기: 비밀번호가 맞을 때만. 본문은 {"action": "check" | "save", "password": "...", "data": {...}}
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ result: 'invalid' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const v = verify_(String(body.password || ''));
    if (v !== 'ok' || body.action === 'check') return json_({ result: v });
    if (body.action !== 'save' || !valid_(body.data)) return json_({ result: 'invalid' });
    writeData_(body.data);
    return json_({ result: 'ok' });
  } finally {
    lock.releaseLock();
  }
}

// 'ok' | 'wrong' | 'locked' | 'nopassword'
function verify_(pw) {
  const props = PropertiesService.getScriptProperties();
  const p = props.getProperties();
  if (!p.pwHash) return 'nopassword';
  const now = Date.now();
  let fails = [];
  try { fails = JSON.parse(p.fails || '[]').filter(t => now - t < WINDOW_MS); } catch (err) { fails = []; }
  if (fails.length >= MAX_FAILS) return 'locked';
  if (hash_(p.pwSalt + pw) === p.pwHash) return 'ok';
  fails.push(now);
  props.setProperty('fails', JSON.stringify(fails));
  return 'wrong';
}

function valid_(d) {
  if (!d || typeof d !== 'object' || Array.isArray(d) || !Array.isArray(d.folders)) return false;
  return JSON.stringify(d).length <= MAX_CHARS;
}

function sheet_(name, hidden) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); if (hidden) sh.hideSheet(); }
  return sh;
}

function readData_() {
  const sh = sheet_(DATA_SHEET, true);
  const last = sh.getLastRow();
  if (!last) return null;
  const text = sh.getRange(1, 1, last, 1).getValues().map(r => String(r[0]).replace(/^~/, '')).join('');
  try { return JSON.parse(text); } catch (err) { return null; }
}

// 시트가 수식으로 읽지 않도록 = + - @ 로 시작하는 글자 앞에 ' 를 붙입니다
function cell_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function writeData_(d) {
  const text = JSON.stringify(d);
  const rows = [];
  for (let i = 0; i < text.length; i += CHUNK) rows.push([MARK + text.slice(i, i + CHUNK)]);
  const sh = sheet_(DATA_SHEET, true);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 1).setValues(rows);

  const list = [['구분', '폴더', '이름', '링크', '색']];
  (d.favorites || []).forEach(f => list.push(['즐겨찾기', '', cell_(f.name), cell_(f.url), cell_(f.color)]));
  d.folders.forEach(f => {
    list.push(['폴더', cell_(f.name), '', '', cell_(f.color)]);
    (f.items || []).forEach(i => list.push(['자료', cell_(f.name), cell_(i.name), cell_(i.url), cell_(i.color)]));
  });
  const ls = sheet_(LIST_SHEET, false);
  ls.clearContents();
  ls.getRange(1, 1, list.length, 5).setValues(list);
}

/**
 * 김아근의 수업자료일람 — 저장 서버 (Google Apps Script)
 *
 * 하는 일
 *  - 사이트가 목록을 읽어 갈 때(GET) 시트에 저장된 목록을 돌려줍니다. 누구나 읽을 수 있습니다.
 *  - 저장(POST)은 관리자 비밀번호가 맞을 때만 받습니다. 비밀번호 확인은 이 서버(구글)에서 합니다.
 *  - 비밀번호는 시트 메뉴 [수업자료일람 > 관리자 비밀번호 설정]에서 정합니다.
 *    코드나 시트에는 비밀번호 대신 해시(되돌릴 수 없는 암호화 값)만 남습니다.
 *  - PDF 올리기(action "upload")도 비밀번호가 맞을 때만 받습니다. 파일은 내 드라이브의
 *    "수업자료일람 PDF" 폴더에 저장되고, 링크가 있는 사람은 볼 수 있게 공유됩니다.
 *    처음 한 번은 시트 메뉴 [수업자료일람 > PDF 기능 승인]으로 드라이브 권한을 허용하세요.
 *  - 10분 안에 5번 틀리면 10분 동안 로그인과 저장을 막습니다.
 *  - "관리자만 보기"로 표시한 폴더(adminOnly)는 누구나 읽는 목록(GET)에서 빼고 보냅니다.
 *    비밀번호가 맞을 때(POST action "load")만 그 폴더까지 담은 전체 목록을 돌려줍니다.
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
const PDF_FOLDER = '수업자료일람 PDF';
const MAX_PDF_BYTES = 20 * 1024 * 1024;   // PDF 한 개 최대 20MB

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('수업자료일람')
    .addItem('관리자 비밀번호 설정', 'setAdminPassword')
    .addItem('로그인 잠금 풀기', 'clearLock')
    .addItem('PDF 기능 승인(처음 한 번)', 'setupPdf')
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

function setupPdf() {
  const folder = pdfFolder_();
  SpreadsheetApp.getUi().alert('PDF 기능을 쓸 준비가 됐습니다. 올린 PDF는 내 드라이브의 "' + folder.getName() + '" 폴더에 저장됩니다.\n\n이제 Apps Script에서 [배포 > 배포 관리 > 편집 > 새 버전 > 배포]를 눌러 주세요.');
}

// 내 드라이브의 PDF 폴더(없으면 만듦)
function pdfFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('pdfFolderId');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (err) {} }
  const folder = DriveApp.createFolder(PDF_FOLDER);
  props.setProperty('pdfFolderId', folder.getId());
  return folder;
}

// PDF 저장: {"name": "...", "b64": "..."} → 드라이브 파일 id와 보기 주소
function uploadPdf_(file) {
  if (!file || typeof file.b64 !== 'string' || !file.b64) return { result: 'invalid' };
  let bytes;
  try { bytes = Utilities.base64Decode(file.b64); } catch (err) { return { result: 'invalid' }; }
  if (bytes.length > MAX_PDF_BYTES) return { result: 'toolarge' };
  // 파일 앞부분이 %PDF 인지 확인(PDF가 아닌 파일은 받지 않음)
  if (!(bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)) return { result: 'notpdf' };
  let name = String(file.name || '자료.pdf').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
  if (!/\.pdf$/i.test(name)) name += '.pdf';
  const created = pdfFolder_().createFile(Utilities.newBlob(bytes, 'application/pdf', name));
  created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { result: 'ok', id: created.getId(), name: name, url: 'https://drive.google.com/file/d/' + created.getId() + '/view' };
}

// PDF 지우기(휴지통으로): 이 사이트가 올린 폴더 안의 파일만 지웁니다
function trashPdf_(id) {
  if (typeof id !== 'string' || !/^[\w-]{10,}$/.test(id)) return { result: 'invalid' };
  try {
    const file = DriveApp.getFileById(id);
    const folderId = pdfFolder_().getId();
    const parents = file.getParents();
    while (parents.hasNext()) { if (parents.next().getId() === folderId) { file.setTrashed(true); return { result: 'ok' }; } }
    return { result: 'invalid' };
  } catch (err) { return { result: 'ok' }; }   // 이미 없는 파일
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

// 읽기: 누구나. 관리자만 보기 폴더는 빼고 보냅니다
function doGet() {
  return json_({ ok: true, data: public_(readData_()) });
}

function public_(d) {
  if (!d || !Array.isArray(d.folders)) return d;
  const out = JSON.parse(JSON.stringify(d));
  out.folders = out.folders.filter(f => !(f && f.adminOnly === true));
  return out;
}

function hasHidden_(d) {
  return !!(d && Array.isArray(d.folders) && d.folders.some(f => f && f.adminOnly === true));
}

// 쓰기: 비밀번호가 맞을 때만. 본문은 {"action": "check" | "load" | "save" | "upload" | "trash", "password": "...", "data": {...}, "full": true}
//  - upload: {"file": {"name": "...", "b64": "..."}} PDF를 드라이브에 저장하고 주소를 돌려줍니다
//  - trash: {"fileId": "..."} 이 사이트가 올린 PDF를 휴지통으로 보냅니다
//  - load: 관리자만 보기 폴더까지 담은 전체 목록을 돌려줍니다
//  - save: 관리자만 보기 폴더가 저장돼 있는데 전체 목록을 받지 않은 화면(full이 아님)이 저장하면,
//          그 폴더가 지워지지 않도록 'stale'로 막습니다
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ result: 'invalid' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const v = verify_(String(body.password || ''));
    if (v !== 'ok' || body.action === 'check') return json_({ result: v });
    if (body.action === 'load') return json_({ result: 'ok', data: readData_() });
    if (body.action === 'upload') return json_(uploadPdf_(body.file));
    if (body.action === 'trash') return json_(trashPdf_(body.fileId));
    if (body.action !== 'save' || !valid_(body.data)) return json_({ result: 'invalid' });
    if (body.full !== true && hasHidden_(readData_())) return json_({ result: 'stale' });
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

  const list = [['구분', '폴더', '이름', '링크', '색', '관리자만']];
  (d.favorites || []).forEach(f => list.push(['즐겨찾기', '', cell_(f.name), cell_(f.url), cell_(f.color), '']));
  d.folders.forEach(f => {
    const only = f.adminOnly === true ? '예' : '';
    list.push(['폴더', cell_(f.name), '', '', cell_(f.color), only]);
    (f.items || []).forEach(i => list.push(['자료', cell_(f.name), cell_(i.name), cell_(i.url), cell_(i.color), only]));
  });
  const ls = sheet_(LIST_SHEET, false);
  ls.clearContents();
  ls.getRange(1, 1, list.length, 6).setValues(list);
}

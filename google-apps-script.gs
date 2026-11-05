/**
 * 시리아이 캠페인 지원서 → 구글 시트 접수
 *
 * 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고
 * "웹 앱"으로 배포하세요. (액세스 권한: 모든 사용자)
 *
 * 탭 두 개가 자동으로 생겨요.
 *  - 지원서 : 지원이 들어올 때마다 한 줄씩 쌓임
 *  - 고료표 : 페이지 코드 → 팔로워 구간 · 고료  (내부용, 인플루언서에게 안 보임)
 *             여기 금액을 채워두면 지원서에 고료가 자동으로 붙어요.
 */

const SHEET_NAME = '지원서';
const RATE_SHEET = '고료표';
const HEADERS = ['접수일시', '지원 브랜드', '이름', '연락처', '이메일', '인스타그램', '팔로워 수', '주소', '하고 싶은 말',
                 '페이지 코드', '팔로워 구간', '고료', '확인', '상태'];
const RATE_HEADERS = ['페이지 코드', '팔로워 구간', '최소 팔로워', '최대 팔로워', '고료', '메모'];
/* 이 파일은 공개돼요. 코드·금액은 여기 적지 말고, 허브에 보이는 값을 시트 '고료표'에 직접 입력하세요.
   (최소/최대 팔로워는 비워두면 확인을 건너뛰어요) */
const RATE_SEED = [];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const code = String(d.page || '').trim();
    const rate = lookupRate_(code);
    getSheet_().appendRow([
      new Date(),
      safe_((d.brands || []).join(', ')),
      safe_(d.name),
      safe_(d.phone),
      safe_(d.email),
      safe_(d.instagram),
      safe_(d.followers),
      safe_(d.address),
      safe_(d.message),
      safe_(code || '(코드 없음)'),
      safe_(rate ? rate[1] : ''),
      rate ? rate[4] : '',
      checkTier_(rate, d.followers),
      '신규',
    ]);
    return json_({ ok: true });   // 고료 정보는 응답에도 넣지 않음
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

/* 브라우저로 웹 앱 주소를 열었을 때 연결 확인용 + 탭 생성 */
function doGet() {
  getSheet_();
  getRateSheet_();
  return json_({ ok: true, message: '연결되었습니다' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#fde4e7');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

function getRateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RATE_SHEET);
    sheet.appendRow(RATE_HEADERS);
    RATE_SEED.forEach((r) => sheet.appendRow(r));
    sheet.getRange(1, 1, 1, RATE_HEADERS.length).setFontWeight('bold').setBackground('#f5f2ec');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function lookupRate_(code) {
  if (!code) return null;
  const rows = getRateSheet_().getDataRange().getValues().slice(1);
  return rows.find((r) => String(r[0]).trim() === code) || null;
}

/* 지원자가 적은 팔로워 수가 링크 구간 안에 있는지 — 다른 구간 링크로 지원한 경우를 표시 */
function checkTier_(rate, followersText) {
  if (!rate) return '⚠ 코드 확인';
  const n = parseFollowers_(followersText);
  if (n == null) return '⚠ 팔로워 확인';
  if (rate[2] === '' && rate[3] === '') return '';
  const min = Number(rate[2]) || 0;
  const max = rate[3] === '' ? Infinity : Number(rate[3]);
  return n >= min && n <= max ? 'OK' : '⚠ 구간 불일치';
}

/* "12,000" "1.2만" "12k" "1만2천" 같은 표기를 숫자로 */
function parseFollowers_(v) {
  let s = String(v || '').replace(/[\s,명]/g, '').toLowerCase();
  if (!s) return null;
  let total = 0, m;
  if ((m = s.match(/^([\d.]+)만/))) { total += parseFloat(m[1]) * 10000; s = s.slice(m[0].length); }
  if ((m = s.match(/^([\d.]+)천/))) { total += parseFloat(m[1]) * 1000; s = s.slice(m[0].length); }
  if ((m = s.match(/^([\d.]+)k$/))) { total += parseFloat(m[1]) * 1000; s = ''; }
  if ((m = s.match(/^([\d.]+)m$/))) { total += parseFloat(m[1]) * 1000000; s = ''; }
  if (s) { if (/^\d+(\.\d+)?$/.test(s)) total += parseFloat(s); else return null; }
  return Math.round(total);
}

/* 앞자리 0(연락처)·@(계정)이 깨지거나, =로 시작하는 값이 수식으로 실행되지 않게 텍스트로 고정 */
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@0]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

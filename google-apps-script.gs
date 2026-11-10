/**
 * 시리아이 캠페인 지원서 → 구글 시트 '지원서' 탭
 *
 * 설치: 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고
 *       배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자) → 나온 /exec 주소를
 *       어드민의 '지원서 접수 연결'에 붙여넣으세요.
 * 코드를 고친 뒤에는 배포 > 배포 관리 > 수정(연필) > 버전: 새 버전 으로 다시 배포해야 반영돼요.
 *
 * 지원한 브랜드마다 한 줄씩 '지원서' 탭 하나에만 쌓여요 (두 브랜드에 지원하면 두 줄).
 * 칸은 1행의 이름으로 찾아서 넣어요. 칸 순서를 바꾸거나 오른쪽에 메모 칸을 추가해도 돼요.
 */

/* 새 지원이 오면 이 주소로 메일 알림. 비워두면 알림 없음. 여러 명은 쉼표로. */
const NOTIFY_EMAIL = '';

/* 링크 코드 → 고료. 허브(PIN)에 보이는 코드와 금액을 적어요.
   ⚠ 깃허브에 올라가는 파일에는 비워두고, Apps Script 에 붙여넣은 코드에만 채워요. */
const RATES = {
  // '코드': '5만원',
};

/* 어드민에서 캠페인마다 적은 고료를 푸는 키. 깃허브 파일에는 비워두고, Apps Script 에 붙여넣은 코드에만 채워요. */
const SHEET_KEY = '';

const SHEET_NAME = '지원서';
const HEADERS = ['접수일시', '고료', '캠페인 고료', '브랜드', '이름', '연락처', '이메일', '인스타그램', '우편번호', '주소', '하고 싶은 말',
                 '공동 작업자 동의', '2차 활용 동의'];
const COL = (name) => HEADERS.indexOf(name) + 1;          // 열 번호 (1부터)
const TEXT_COLS = ['연락처', '우편번호'];                  // 앞자리 0 이 지워지지 않게 '일반 텍스트' 서식
const TEST_MARK = '[테스트] ';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const names = (d.brands || []).map(String);
    if (!names.length) return json_({ ok: false, error: 'no brand' });

    const code = String(d.page || '').trim();
    const fee = RATES[code] || (code ? '확인 필요 (' + code + ')' : '링크 없음');
    const insta = normInsta_(d.instagram);
    const instaUrl = insta ? 'https://www.instagram.com/' + insta + '/' : '';
    const phone = phone_(d.phone);
    const zip = String(d.zip || '').replace(/\D/g, '').slice(0, 5);
    const name = (d.test === true ? TEST_MARK : '') + String(d.name || '');
    const now = new Date();

    const fees = (d.fees || []).map(unsealFee_);
    const values = names.map((brand, i) => [
      now, fee, safe_(fees[i] || ''), safe_(brand), safe_(name), phone, safe_(d.email), instaUrl || safe_(d.instagram),
      zip, safe_(d.address), safe_(d.message),
      agree_(d.collabAgree), agree_(d.secondaryAgree),
    ]);
    /* 칸 이름으로 찾아서 넣어요 — 칸 순서를 바꾸거나 메모 칸을 추가해도 그대로 쌓여요 */
    const col = columnMap_(sheet);
    const width = sheet.getLastColumn();
    const rows = values.map((v) => { const r = new Array(width).fill(''); HEADERS.forEach((h, j) => { r[col[h] - 1] = v[j]; }); return r; });
    const start = sheet.getLastRow() + 1;
    TEXT_COLS.forEach((h) => sheet.getRange(start, col[h], rows.length, 1).setNumberFormat('@'));
    sheet.getRange(start, 1, rows.length, width).setValues(rows);
    if (instaUrl) {   // 인스타그램 칸을 눌러 바로 프로필로 이동
      const link = SpreadsheetApp.newRichTextValue().setText(instaUrl).setLinkUrl(instaUrl).build();
      sheet.getRange(start, col['인스타그램'], rows.length, 1).setRichTextValues(rows.map(() => [link]));
    }

    if (NOTIFY_EMAIL && d.test !== true) {
      try {
        MailApp.sendEmail(NOTIFY_EMAIL, `[시리아이] 새 지원 · ${names.join(', ')} · ${d.name}`,
          `브랜드: ${names.join(', ')}\n고료: ${fee}${fees.some(Boolean) ? ' / 캠페인 고료: ' + fees.join(', ') : ''}\n이름: ${d.name}\n인스타그램: ${instaUrl || d.instagram}\n연락처: ${phone}\n\n시트: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
      } catch (err) {}
    }
    return json_({ ok: true });   // 고료 정보는 응답에 넣지 않음
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

/* 브라우저로 /exec 주소를 열면 연결 확인 + 탭 생성 */
function doGet() {
  getSheet_();
  return json_({ ok: true, message: '연결되었습니다' });
}

/* '지원서' 탭은 하나만 써요. 탭을 새로 만들거나 옮기지 않아요. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet && sheet.getLastRow() > 0) return sheet;
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME, 0);
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#fde4e7');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  TEXT_COLS.forEach((h) => sheet.getRange(2, COL(h), sheet.getMaxRows() - 1, 1).setNumberFormat('@'));
  return sheet;
}

/* 1행 칸 이름 → 열 번호. 없는 칸은 지우지 않고 맨 오른쪽에 새로 만들어요 (메모 칸 등 다른 칸은 건드리지 않음) */
function columnMap_(sheet) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const head = sheet.getRange(1, 1, 1, width).getValues()[0].map((h) => String(h).trim());
  const col = {};
  head.forEach((h, i) => { if (h && !(h in col)) col[h] = i + 1; });
  let next = width + 1;
  HEADERS.forEach((h) => {
    if (col[h]) return;
    sheet.getRange(1, next).setValue(h).setFontWeight('bold').setBackground('#fde4e7');
    col[h] = next++;
  });
  return col;
}

/* 어드민이 잠가 둔 캠페인 고료(enc2:…) 풀기 — HMAC-SHA256 키스트림, 태그가 안 맞으면 '' */
function unsealFee_(s) {
  s = String(s || '');
  if (!SHEET_KEY || s.indexOf('enc2:') !== 0) return '';
  try {
    const key = Utilities.base64Decode(SHEET_KEY);
    const bytes = (t) => Utilities.newBlob(t).getBytes();
    const hm = (data) => Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, data, key);
    const [nonce, ct, tag] = s.slice(5).split('.').map((x) => Utilities.base64Decode(x));
    const want = hm([].concat(bytes('tag'), nonce, ct)).slice(0, 16);
    if (tag.length !== 16 || want.some((x, i) => x !== tag[i])) return '';
    const out = [];
    for (let i = 0; i * 32 < ct.length; i++) {
      const ks = hm([].concat(nonce, bytes('ks'), [i]));
      for (let j = 0; j < 32 && i * 32 + j < ct.length; j++) out.push(ct[i * 32 + j] ^ ks[j]);
    }
    return Utilities.newBlob(out).getDataAsString('UTF-8');
  } catch (err) { return ''; }
}

/* true → '동의', false → '미동의', 해당 없음 → '-' */
function agree_(v) { return v === true ? '동의' : v === false ? '미동의' : '-'; }

/* '@Abc', 'abc', 'instagram.com/abc', 'https://www.instagram.com/abc/?igsh=…' → 'abc'
   계정 이름으로 쓸 수 없는 값이면 '' (그때는 지원자가 적은 그대로 시트에 남김) */
function normInsta_(v) {
  const h = String(v || '').trim().replace(/^'/, '')
    .replace(/^(https?:\/\/)?(www\.|m\.)?instagram\.com\//i, '')
    .replace(/^@/, '').replace(/[/?#].*$/, '').toLowerCase();
  return /^[a-z0-9._]{1,30}$/.test(h) ? h : '';
}

/* 01012345678 / 010 1234 5678 / 010.1234.5678 → 010-1234-5678 (02 지역번호도 처리) */
function phone_(v) {
  const n = String(v || '').replace(/\D/g, '');
  if (/^02\d{7,8}$/.test(n)) return n.replace(/^(02)(\d{3,4})(\d{4})$/, '$1-$2-$3');
  if (/^0\d{9,10}$/.test(n)) return n.replace(/^(0\d{2})(\d{3,4})(\d{4})$/, '$1-$2-$3');
  return safe_(String(v || '').trim());
}

/* =, +, -, @ 로 시작하는 값이 수식으로 실행되지 않게 텍스트로 고정 */
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

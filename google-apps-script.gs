/**
 * 시리아이 캠페인 지원서 → 구글 시트 접수
 *
 * 설치: 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣고
 *       배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자) → 나온 /exec 주소를
 *       어드민의 '지원서 접수 연결'에 붙여넣으세요.
 * 코드를 고친 뒤에는 배포 > 배포 관리 > 수정(연필) > 버전: 새 버전 으로 다시 배포해야 반영돼요.
 *
 * 탭 두 개가 자동으로 생겨요.
 *  - 지원서 : 지원한 브랜드마다 한 줄씩 (여러 브랜드에 지원하면 여러 줄) — 브랜드로 필터하기 쉽게
 *  - 고료표 : 페이지 코드 → 고료 (허브에 보이는 값을 직접 입력, 공개되지 않음)
 */

/* 새 지원이 오면 이 주소로 메일 알림. 비워두면 알림 없음. 여러 명은 쉼표로. */
const NOTIFY_EMAIL = '';

const SHEET_NAME = '지원서';
const RATE_SHEET = '고료표';
const HEADERS = ['접수일시', '브랜드', '이름', '연락처', '이메일', '인스타그램', '우편번호', '주소', '하고 싶은 말',
                 '페이지 코드', '팔로워 구간', '고료', '확인', '상태', '브랜드 ID', '공동 작업자 동의', '2차 활용 동의'];
const COL = (name) => HEADERS.indexOf(name) + 1;          // 열 번호 (1부터)
const TEXT_COLS = ['연락처', '우편번호'];                  // 앞자리 0 이 지워지지 않게 '일반 텍스트' 서식
const RATE_HEADERS = ['페이지 코드', '팔로워 구간', '고료', '메모'];
/* 이 파일은 공개돼요. 코드·금액은 여기 적지 말고 시트 '고료표'에 직접 입력하세요. */
const RATE_SEED = [];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const isTest = d.test === true;
    const names = (d.brands || []).map(String);
    const ids = (d.brandIds || []).map(String);
    if (!names.length) return json_({ ok: false, error: 'no brand' });

    const code = String(d.page || '').trim();
    const rate = lookupRate_(code);
    const insta = normInsta_(d.instagram);
    const instaUrl = insta ? 'https://www.instagram.com/' + insta + '/' : '';
    const phone = phone_(d.phone);
    const zip = String(d.zip || '').replace(/\D/g, '').slice(0, 5);
    const seen = existingKeys_(sheet);
    const now = new Date();

    const rows = names.map((name, i) => {
      const id = ids[i] || name;
      const dup = !isTest && seen.has(insta + '|' + id);
      const checks = [rate ? '' : '⚠ 코드 확인', dup ? '⚠ 중복 지원' : ''].filter(Boolean).join(' / ');
      return [
        now, safe_(name), safe_(d.name), phone, safe_(d.email), instaUrl || safe_(d.instagram),
        zip, safe_(d.address), safe_(d.message),
        safe_(code || '(코드 없음)'), safe_(rate ? rate[1] : ''), rate ? rate[2] : '',
        checks || 'OK', isTest ? '테스트' : '신규', safe_(id),
        agree_(d.collabAgree), agree_(d.secondaryAgree),
      ];
    });
    const start = sheet.getLastRow() + 1;
    TEXT_COLS.forEach((h) => sheet.getRange(start, COL(h), rows.length, 1).setNumberFormat('@'));
    sheet.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
    if (instaUrl) {   // 인스타그램 칸을 눌러 바로 프로필로 이동
      const link = SpreadsheetApp.newRichTextValue().setText(instaUrl).setLinkUrl(instaUrl).build();
      sheet.getRange(start, COL('인스타그램'), rows.length, 1).setRichTextValues(rows.map(() => [link]));
    }

    if (NOTIFY_EMAIL && !isTest) {
      try {
        MailApp.sendEmail(NOTIFY_EMAIL, `[시리아이] 새 지원 · ${names.join(', ')} · ${d.name}`,
          `브랜드: ${names.join(', ')}\n이름: ${d.name}\n인스타그램: ${instaUrl || d.instagram}\n연락처: ${phone}\n\n시트: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
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
    TEXT_COLS.forEach((h) => sheet.getRange(2, COL(h), sheet.getMaxRows() - 1, 1).setNumberFormat('@'));
    sheet.getRange(2, COL('상태'), sheet.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['신규', '검토중', '선정', '탈락', '발송완료', '업로드완료', '테스트'], true).build());
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

/* 이미 지원한 (인스타그램 + 브랜드 ID) 조합 */
function existingKeys_(sheet) {
  const last = sheet.getLastRow();
  const set = new Set();
  if (last < 2) return set;
  const vals = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const iInsta = COL('인스타그램') - 1, iStatus = COL('상태') - 1, iId = COL('브랜드 ID') - 1;
  vals.forEach((r) => { if (r[iStatus] !== '테스트') set.add(normInsta_(r[iInsta]) + '|' + String(r[iId] || r[1])); });
  return set;
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

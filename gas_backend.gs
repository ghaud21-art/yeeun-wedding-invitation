/**
 * 모바일 청첩장 GAS 백엔드
 * 배포 방법:
 *   1) 아래 CONFIG 값을 실제 스프레드시트 ID / 드라이브 폴더 ID / Make Webhook URL로 채운다.
 *   2) 배포 > 웹 앱으로 배포, 액세스 권한: '전체' (누구나), 실행 계정: '나'
 *   3) 배포된 /exec URL을 index.html의 CFG.GAS_URL 에 넣는다.
 *   4) GALLERY_FOLDER 하위에 '봄' / '여름' / '가을' / '겨울' 이름의 서브폴더를 만든다.
 *
 * 시트 탭은 첫 조회/저장 시 자동으로 만들어지고, 그중 '설정'과 '계좌' 탭은
 * 처음 만들어질 때 예시 값이 자동으로 채워진다. 스프레드시트를 열어 그 값을
 * 원하는 내용으로 바꾸면(친구를 편집자로 초대해도 됨) 청첩장 페이지에 바로 반영된다.
 *   - 방명록: 타임스탬프 | 이름 | 메시지
 *   - 참석여부: 타임스탬프 | 이름 | 연락처 | 하객측 | 참석여부 | 식사여부 | 동행인원
 *   - 게스트사진: 타임스탬프 | 파일명 | 파일URL | 업로더
 *   - 설정: 항목 | 내용 | 입력 예시   (신랑/신부 이름, 날짜, 장소, 교통 안내 등)
 *   - 계좌: 측 | 관계 | 이름 | 은행 | 계좈번호
 */

const SHEET_ID       = 'YOUR_SHEET_ID';
const GALLERY_FOLDER = 'YOUR_GALLERY_FOLDER_ID';
const GUEST_FOLDER   = 'YOUR_GUEST_FOLDER_ID';
const MAKE_WEBHOOK   = 'YOUR_MAKE_WEBHOOK_URL';

const SHEET_GUESTBOOK = '방명록';
const SHEET_ATTEND    = '참석여부';
const SHEET_GUESTPIC  = '게스트사진';
const SHEET_CONFIG    = '설정';
const SHEET_ACCOUNTS  = '계좌';

const GALLERY_TAB_FOLDER = {
  spring: '봄',
  summer: '여름',
  fall: '가을',
  winter: '겨울'
};

/* ---------------- 공통 유틸 ---------------- */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// header를 넘기면, 시트가 없어서 새로 만들 때 그 헤더 행(들)을 함께 써준다.
function getSheet_(name, headerRows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headerRows) headerRows.forEach(row => sheet.appendRow(row));
  }
  return sheet;
}

// 숫자만 추출 후 010-1234-5678 형식으로 복원
function formatPhone_(raw) {
  let digits = String(raw || '').replace(/\D/g, '');
  if (digits.length > 0 && digits[0] !== '0') digits = '0' + digits;
  if (digits.length === 11) {
    return digits.slice(0, 3) + '-' + digits.slice(3, 7) + '-' + digits.slice(7);
  } else if (digits.length === 10) {
    return digits.slice(0, 3) + '-' + digits.slice(3, 6) + '-' + digits.slice(6);
  }
  return digits;
}

function digitsOnly_(raw) {
  return String(raw || '').replace(/\D/g, '');
}

function driveImageUrl_(fileId) {
  return 'https://drive.google.com/uc?export=view&id=' + fileId;
}

/* ---------------- doGet (조회) ---------------- */

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'guestbook') return jsonOut_(getGuestbook_());
    if (action === 'gallery') return jsonOut_(getGallery_(e.parameter.tab));
    if (action === 'guestPhotos') return jsonOut_(getGuestPhotos_());
    if (action === 'config') return jsonOut_({ ok: true, config: getConfigMap_(), accounts: getAccountsMap_() });
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getGuestbook_() {
  const sheet = getSheet_(SHEET_GUESTBOOK, [['타임스탬프', '이름', '메시지']]);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[1]) continue;
    rows.push({ timestamp: r[0], name: r[1], message: r[2] });
  }
  rows.reverse();
  return { ok: true, items: rows };
}

function getGallery_(tab) {
  const folderName = GALLERY_TAB_FOLDER[tab] || GALLERY_TAB_FOLDER.spring;
  const root = DriveApp.getFolderById(GALLERY_FOLDER);
  const subFolders = root.getFoldersByName(folderName);
  if (!subFolders.hasNext()) return { ok: true, items: [] };
  const folder = subFolders.next();
  const files = folder.getFiles();
  const items = [];
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') !== 0) continue;
    items.push({ id: f.getId(), name: f.getName(), url: driveImageUrl_(f.getId()) });
  }
  return { ok: true, items: items };
}

function getGuestPhotos_() {
  const folder = DriveApp.getFolderById(GUEST_FOLDER);
  const files = folder.getFiles();
  const items = [];
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') !== 0) continue;
    items.push({ id: f.getId(), name: f.getName(), url: driveImageUrl_(f.getId()) });
  }
  return { ok: true, items: items };
}

// '설정' 탭(항목 | 내용 | 입력 예시)을 { 항목: 내용 } 형태로 반환.
// 처음 조회되어 시트가 새로 만들어졌다면 예시 값을 함께 채워 넣는다.
function getConfigMap_() {
  const sheet = getSheet_(SHEET_CONFIG, [['항목', '내용', '입력 예시']]);
  if (sheet.getLastRow() <= 1) seedDefaultConfig_(sheet);
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (!key) continue;
    let val = values[i][1];
    if (val instanceof Date) {
      // 날짜/시간으로 인식된 셀은 타임존 오차 없이 그대로 쓸 수 있게 문자열로 변환
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
    }
    map[key] = (val === undefined || val === null) ? '' : String(val).trim();
  }
  return map;
}

// '계좌' 탭(측 | 관계 | 이름 | 은행 | 계좈번호)을 { groom:[...], bride:[...] } 형태로 반환.
function getAccountsMap_() {
  const sheet = getSheet_(SHEET_ACCOUNTS, [['측', '관계', '이름', '은행', '계좈번호']]);
  if (sheet.getLastRow() <= 1) seedDefaultAccounts_(sheet);
  const values = sheet.getDataRange().getValues();
  const accounts = { groom: [], bride: [] };
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const side = String(r[0] || '').trim();
    if (!side) continue;
    const entry = { who: String(r[1] || ''), name: String(r[2] || ''), bank: String(r[3] || ''), num: String(r[4] || '') };
    if (!entry.bank && !entry.num) continue; // 은행/계좌를 안 채운 행은 화면에 노출하지 않음
    if (side.indexOf('신랑') === 0) accounts.groom.push(entry);
    else if (side.indexOf('신부') === 0) accounts.bride.push(entry);
  }
  return accounts;
}

function seedDefaultConfig_(sheet) {
  [
    ['신랑 이름', '김도윤', ''],
    ['신부 이름', '이서연', ''],
    ['신랑 아버지', '김재현', ''],
    ['신랑 어머니', '박은주', ''],
    ['신랑측 순서', '장남', '장남 / 차남 / 삼남 등'],
    ['신부 아버지', '이성호', ''],
    ['신부 어머니', '최미란', ''],
    ['신부측 순서', '장녀', '장녀 / 차녀 / 삼녀 등'],
    ['예식 날짜', '2027-05-22', 'YYYY-MM-DD'],
    ['예식 시간', '14:00', '24시간제 HH:MM'],
    ['예식장 이름', '라온컨벤션 3층 그랜드홀', ''],
    ['주소', '서울특별시 강남구 테헤란로 123', ''],
    ['카카오맵 링크', '', '카카오맵 앱에서 장소 공유 링크 복사'],
    ['네이버지도 링크', '', ''],
    ['티맵 링크', '', ''],
    ['지하철 안내', '2호선 선릉역 4번 출구\n도보 5분', '줄바꿈은 셀 안에서 Alt+Enter'],
    ['버스 안내', '146 · 341 · 360\n라온컨벤션 앞 하차', ''],
    ['자가용 안내', '내비게이션에 "라온컨벤션" 검색', ''],
    ['주차 안내', '건물 지하 B1~B3\n2시간 무료', ''],
    ['카카오페이 링크', '', '카카오페이 송금 요청 링크']
  ].forEach(row => sheet.appendRow(row));
}

function seedDefaultAccounts_(sheet) {
  [
    ['신랑측', '신랑', '김도윤', '신한은행', '110-234-567890'],
    ['신랑측', '父', '김재현', '', ''],
    ['신랑측', '母', '박은주', '', ''],
    ['신부측', '신부', '이서연', '국민은행', '012-34-5678-901'],
    ['신부측', '父', '이성호', '', ''],
    ['신부측', '母', '최미란', '', '']
  ].forEach(row => sheet.appendRow(row));
}

/* ---------------- doPost (저장) ----------------
 * 브라우저 CORS 프리플라이트를 피하기 위해 프론트엔드에서
 * fetch(url, { method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body: JSON.stringify(payload) })
 * 형태로 호출해야 한다. body는 e.postData.contents로 수신, JSON.parse해서 사용.
 */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'guestbook') return jsonOut_(saveGuestbook_(body));
    if (action === 'attend') return jsonOut_(saveAttend_(body));
    if (action === 'uploadPhoto') return jsonOut_(saveGuestPhoto_(body));
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function saveGuestbook_(body) {
  const sheet = getSheet_(SHEET_GUESTBOOK, [['타임스탬프', '이름', '메시지']]);
  sheet.appendRow([new Date(), body.name || '', body.message || '']);
  return { ok: true };
}

function saveAttend_(body) {
  const sheet = getSheet_(SHEET_ATTEND, [['타임스탬프', '이름', '연락처', '하객측', '참석여부', '식사여부', '동행인원']]);
  const phoneFormatted = formatPhone_(body.phone);
  sheet.appendRow([
    new Date(),
    body.name || '',
    "'" + phoneFormatted, // 앞자리 0 유지를 위해 텍스트 강제
    body.side || '',
    body.attend || '',
    body.meal || '',
    body.companions || 0
  ]);

  if (MAKE_WEBHOOK && MAKE_WEBHOOK.indexOf('YOUR_') !== 0) {
    try {
      UrlFetchApp.fetch(MAKE_WEBHOOK, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          name: body.name || '',
          phone: digitsOnly_(body.phone),
          side: body.side || '',
          attend: body.attend || '',
          meal: body.meal || '',
          companions: body.companions || 0
        }),
        muteHttpExceptions: true
      });
    } catch (webhookErr) {
      // 웹훅 실패해도 참석 응답 저장은 성공 처리
    }
  }
  return { ok: true };
}

function saveGuestPhoto_(body) {
  const folder = DriveApp.getFolderById(GUEST_FOLDER);
  const base64 = body.fileData.split(',').pop();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', body.fileName || ('guest_' + Date.now() + '.jpg'));
  const file = folder.createFile(blob);

  const sheet = getSheet_(SHEET_GUESTPIC, [['타임스탬프', '파일명', '파일URL', '업로더']]);
  sheet.appendRow([new Date(), file.getName(), driveImageUrl_(file.getId()), body.uploader || '']);

  return { ok: true, url: driveImageUrl_(file.getId()) };
}

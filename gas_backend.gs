/**
 * 모바일 청첩장 GAS 백엔드
 * 배포 방법:
 *   1) 아래 SHEET_ID / GALLERY_FOLDER / GUEST_FOLDER는 이미 채워져 있음(구글 드라이브에 자동 생성됨).
 *      MAKE_WEBHOOK만 필요하면 채운다.
 *   2) script.google.com에서 새 프로젝트를 만들고 이 파일 내용을 붙여넣는다.
 *   3) 배포 > 웹 앱으로 배포, 액세스 권한: '전체' (누구나), 실행 계정: '나'
 *   4) 배포된 /exec URL을 index.html의 CFG.GAS_URL 에 넣는다.
 *
 * 시트 탭은 첫 조회/저장 시 자동으로 만들어지고, 그중 '설정'과 '계좌' 탭은
 * 처음 만들어질 때 예시 값이 자동으로 채워진다. 스프레드시트를 열어 그 값을
 * 원하는 내용으로 바꾸면(친구를 편집자로 초대해도 됨) 청첩장 페이지에 바로 반영된다.
 *   - 방명록: 타임스탬프 | 이름 | 메시지 | 비밀번호 (작성자가 나중에 수정할 때 사용, 조회 응답엔 노출 안 함)
 *   - 참석여부: 타임스탬프 | 이름 | 연락처 | 하객측 | 참석여부 | 식사여부 | 동행인원
 *   - 게스트사진: 타임스탬프 | 파일명 | 파일URL | 업로더
 *   - 설정: 항목 | 내용 | 입력 예시   (신랑/신부 이름, 날짜, 장소, 교통 안내 등)
 *   - 계좌: 측 | 관계 | 이름 | 은행 | 계좈번호
 */

const SHEET_ID       = '1ILxUsi45jWYk8zIt1UrmVUMmu4OIY5ghSEJBETHEARc'; // "예은이 청첩장 데이터" 스프레드시트
const GALLERY_FOLDER = '1WrPSohG8a26JBYpMrHVU_CWqMYVFw6eq'; // "예은이 청첩장 - 갤러리" (봄/여름/가을/겨울/메인 사진 하위 폴더 포함)
const GUEST_FOLDER   = '1O9aOtl2yP6gcg83p6m5wlFDbAuP6N5jH'; // "예은이 청첩장 - 하객사진"
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

// 갤러리 폴더 하위에 이 이름의 서브폴더를 만들고 사진 1장을 넣으면 히어로 메인 사진으로 쓰인다.
const MAIN_PHOTO_FOLDER_NAME = '메인 사진';

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

// 'uc?export=view' 형태는 광고/추적 차단 확장 프로그램에 자주 걸려 이미지가 안 뜨는
// 경우가 있어서, 구글 드라이브의 썸네일 전용 엔드포인트를 사용한다.
function driveImageUrl_(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';
}

/* ---------------- doGet (조회) ---------------- */

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'guestbook') return jsonOut_(getGuestbook_());
    if (action === 'gallery') return jsonOut_(getGallery_(e.parameter.tab));
    if (action === 'guestPhotos') return jsonOut_(getGuestPhotos_());
    if (action === 'mainPhoto') return jsonOut_(getMainPhoto_());
    if (action === 'config') return jsonOut_({ ok: true, config: getConfigMap_(), accounts: getAccountsMap_() });
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getGuestbook_() {
  const sheet = getSheet_(SHEET_GUESTBOOK, [['타임스탬프', '이름', '메시지', '비밀번호']]);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[1]) continue;
    // id는 실제 시트 행 번호(1-indexed) — 수정 요청 시 이 번호로 행을 다시 찾는다.
    // 비밀번호는 조회 응답에 절대 포함하지 않는다.
    rows.push({ id: i + 1, timestamp: r[0], name: r[1], message: r[2] });
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

// 갤러리 폴더 하위 '메인 사진' 폴더에서 이미지 1장을 찾아 반환한다(히어로 배경용).
function getMainPhoto_() {
  const root = DriveApp.getFolderById(GALLERY_FOLDER);
  const subFolders = root.getFoldersByName(MAIN_PHOTO_FOLDER_NAME);
  if (!subFolders.hasNext()) return { ok: true, url: null };
  const files = subFolders.next().getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (String(f.getMimeType()).indexOf('image/') !== 0) continue;
    return { ok: true, url: driveImageUrl_(f.getId()) };
  }
  return { ok: true, url: null };
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
  // '내용' 열을 텍스트 서식으로 고정해서 날짜/시간을 입력해도 구글 시트가
  // Date로 자동 변환하지 않게 한다(자동 변환 시 시간대 오차로 값이 틀어짐).
  sheet.getRange('B2:B1000').setNumberFormat('@');
  if (sheet.getLastRow() <= 1) seedDefaultConfig_(sheet);
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (!key) continue;
    let val = values[i][1];
    if (val instanceof Date) {
      // 그래도 Date로 인식된 셀(붙여넣기 등)은 UTC 기준으로 읽어 시간대 오차를 피한다.
      val = Utilities.formatDate(val, 'UTC', "yyyy-MM-dd'T'HH:mm");
    }
    map[key] = (val === undefined || val === null) ? '' : String(val).trim();
  }
  return map;
}

// '계좌' 탭(측 | 관계 | 이름 | 은행 | 계좈번호 | 카카오페이 링크)을 { groom:[...], bride:[...] } 형태로 반환.
// 카카오페이 링크는 각 행마다 있을 수도 없을 수도 있고, 비어 있으면 화면에 버튼이 뜨지 않는다.
function getAccountsMap_() {
  const sheet = getSheet_(SHEET_ACCOUNTS, [['측', '관계', '이름', '은행', '계좈번호', '카카오페이 링크']]);
  if (sheet.getLastRow() <= 1) seedDefaultAccounts_(sheet);
  const values = sheet.getDataRange().getValues();
  const accounts = { groom: [], bride: [] };
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const side = String(r[0] || '').trim();
    if (!side) continue;
    const entry = { who: String(r[1] || ''), name: String(r[2] || ''), bank: String(r[3] || ''), num: String(r[4] || ''), kakaopay: String(r[5] || '').trim() };
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
    ['소개 문구', '서로의 계절을 지나\n같은 방향을 바라보게 된 두 사람이\n이제 하나의 길을 걷고자 합니다.\n\n귀한 걸음으로 축복해 주시면\n더없는 기쁨으로 간직하겠습니다.', '줄바꿈은 셀 안에서 Alt+Enter'],
    ['예식장 이름', '호텔 라뷔포레', ''],
    ['주소', '경기 수원시 팔달구 인계동 1133-8', ''],
    ['카카오맵 링크', '', '카카오맵 앱에서 장소 공유 링크 복사'],
    ['네이버지도 링크', '', ''],
    ['티맵 링크', '', ''],
    ['지하철 안내', '지하철 안내 (추후 입력)', '줄바꿈은 셀 안에서 Alt+Enter'],
    ['버스 안내', '버스 안내 (추후 입력)', ''],
    ['자가용 안내', '내비게이션에 "호텔 라뷔포레" 검색', ''],
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
    if (action === 'editGuestbook') return jsonOut_(editGuestbook_(body));
    if (action === 'attend') return jsonOut_(saveAttend_(body));
    if (action === 'uploadPhoto') return jsonOut_(saveGuestPhoto_(body));
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function saveGuestbook_(body) {
  const sheet = getSheet_(SHEET_GUESTBOOK, [['타임스탬프', '이름', '메시지', '비밀번호']]);
  sheet.appendRow([new Date(), body.name || '', body.message || '', "'" + String(body.password || '').trim()]);
  return { ok: true };
}

// 작성자가 남긴 비밀번호가 일치할 때만 이름/메시지를 수정한다.
function editGuestbook_(body) {
  const sheet = getSheet_(SHEET_GUESTBOOK, [['타임스탬프', '이름', '메시지', '비밀번호']]);
  const id = Number(body.id);
  if (!id || id < 2 || id > sheet.getLastRow()) return { ok: false, error: '글을 찾을 수 없습니다' };
  const row = sheet.getRange(id, 1, 1, 4).getValues()[0];
  const savedPw = String(row[3] || '').trim();
  const inputPw = String(body.password || '').trim();
  if (!savedPw || savedPw !== inputPw) return { ok: false, error: '비밀번호가 일치하지 않습니다' };
  sheet.getRange(id, 2, 1, 2).setValues([[body.name || '', body.message || '']]);
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

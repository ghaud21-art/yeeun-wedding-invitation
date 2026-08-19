/**
 * 모바일 청첩장 GAS 백엔드
 * 배포 방법:
 *   1) 아래 CONFIG 값을 실제 스프레드시트 ID / 드라이브 폴더 ID / Make Webhook URL로 채운다.
 *   2) 스프레드시트에 '방명록', '참석여부', '게스트사진' 3개 탭(시트)을 만든다.
 *      - 방명록: 타임스탬프 | 이름 | 메시지
 *      - 참석여부: 타임스탬프 | 이름 | 연락처 | 하객측 | 참석여부 | 식사여부 | 동행인원
 *      - 게스트사진: 타임스탬프 | 파일명 | 파일URL | 업로더
 *   3) GALLERY_FOLDER 하위에 '스냅사진' / '일상' / '여행' 이름의 서브폴더를 만든다.
 *   4) 배포 > 웹 앱으로 배포, 액세스 권한: '전체' (누구나), 실행 계정: '나'
 *   5) 배포된 /exec URL을 index.html의 CFG.GAS_URL 에 넣는다.
 */

const SHEET_ID       = 'YOUR_SHEET_ID';
const GALLERY_FOLDER = 'YOUR_GALLERY_FOLDER_ID';
const GUEST_FOLDER   = 'YOUR_GUEST_FOLDER_ID';
const MAKE_WEBHOOK   = 'YOUR_MAKE_WEBHOOK_URL';

const SHEET_GUESTBOOK = '방명록';
const SHEET_ATTEND    = '참석여부';
const SHEET_GUESTPIC  = '게스트사진';

const GALLERY_TAB_FOLDER = {
  snap: '스냅사진',
  daily: '일상',
  travel: '여행'
};

/* ---------------- 공통 유틸 ---------------- */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
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
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getGuestbook_() {
  const sheet = getSheet_(SHEET_GUESTBOOK);
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
  const folderName = GALLERY_TAB_FOLDER[tab] || GALLERY_TAB_FOLDER.snap;
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
  const sheet = getSheet_(SHEET_GUESTBOOK);
  sheet.appendRow([new Date(), body.name || '', body.message || '']);
  return { ok: true };
}

function saveAttend_(body) {
  const sheet = getSheet_(SHEET_ATTEND);
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

  const sheet = getSheet_(SHEET_GUESTPIC);
  sheet.appendRow([new Date(), file.getName(), driveImageUrl_(file.getId()), body.uploader || '']);

  return { ok: true, url: driveImageUrl_(file.getId()) };
}

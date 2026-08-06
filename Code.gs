/**
 * 제작가이드 아카이브 — Google Sheets 연동용 Apps Script 웹앱
 * ------------------------------------------------------------------
 * 이 스크립트는 HTML 파일(제작가이드_아카이브)이 로그인 없이도 구글 시트를
 * 읽고 쓸 수 있게 해주는 "대리인" 역할을 합니다. HTML에서는 fetch로 이
 * 웹앱의 배포 URL만 호출하면 되고, 실제 시트 접근 권한은 이 스크립트를
 * 배포한 사람(당신) 계정으로 처리됩니다.
 *
 * ⚠️ 업데이트 내용 (항목 분리 + 이미지 URL 지원)
 * - "광고사이즈" 컬럼(L)은 이제 [{item, content}] 형태의 JSON을 저장합니다
 *   (기존에 자유 텍스트로 저장된 데이터는 HTML 쪽에서 자동으로 변환해서 읽습니다).
 * - 새 컬럼 "이미지URL"(Q)이 추가되었습니다.
 * - 기존 시트에 이미 헤더가 있어도 ensureHeader_가 자동으로 새 컬럼 헤더를
 *   채워 넣습니다 (직접 손댈 필요 없음).
 *
 * [배포 방법]
 * 1. script.google.com → "ㅈㅈ" 프로젝트 → 이 파일 내용을 Code.gs에 전체 붙여넣기 (덮어쓰기)
 * 2. 아래 SHEET_ID / SHEET_NAME이 HTML 쪽 GOOGLE_SHEET_ID / GOOGLE_SHEET_NAME과
 *    동일한지 확인 (이미 같은 값으로 맞춰져 있습니다)
 * 3. 우측 상단 "배포" → "배포 관리" → 연필(수정) 아이콘 → 버전: "새 버전" → "배포"
 *    (URL은 그대로 유지됩니다. 새로 만들 필요 없습니다.)
 */

// ====== 설정 (HTML 쪽 GOOGLE_SHEET_ID / GOOGLE_SHEET_NAME과 동일하게 유지) ======
const SHEET_ID = "1LTLh2Q7-KHs5k9rzlrpYxd6ZggDL6xmOR6Bb0TzFxoI";
const SHEET_NAME = "시트1";

// 시트 열 구조 (A~Q) — HTML의 HEADERS/COL 상수와 반드시 동일한 순서를 유지하세요.
const HEADERS = ["ID","구분","매체","가이드명","제공방식","링크URL","파일명","파일크기","파일데이터",
                  "지면","광고유형","광고사이즈","파일형식","비고","수정일시","정렬순서","이미지URL"];

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  ensureHeader_(sheet);
  return sheet;
}

function ensureHeader_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  const first = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let mismatch = false;
  for (let i = 0; i < HEADERS.length; i++) {
    if (first[i] !== HEADERS[i]) { mismatch = true; break; }
  }
  if (mismatch) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- GET: 데이터 조회 ----------
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "read";
    const sheet = getSheet_();

    if (action === "read") {
      const lastRow = sheet.getLastRow();
      const lastCol = HEADERS.length;
      let values = [];
      if (lastRow > 1) {
        values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      }
      return jsonOut_({ ok: true, values: values });
    }

    return jsonOut_({ ok: false, error: "UNKNOWN_ACTION: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ---------- POST: append(생성) / update(수정) / delete(삭제) ----------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const sheet = getSheet_();

    if (action === "append") {
      const rows = body.rows || [];
      if (rows.length > 0) {
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
      }
      return jsonOut_({ ok: true, appended: rows.length });
    }

    if (action === "update") {
      const updates = body.updates || []; // [{ rowNumber, row }]
      updates.forEach(function (u) {
        if (u.rowNumber > 1) {
          sheet.getRange(u.rowNumber, 1, 1, HEADERS.length).setValues([u.row]);
        }
      });
      return jsonOut_({ ok: true, updated: updates.length });
    }

    if (action === "delete") {
      const rowNumbers = (body.rowNumbers || [])
        .filter(function (n) { return n > 1; })
        .sort(function (a, b) { return b - a; }); // 뒤에서부터 삭제해야 앞 행 번호가 안 밀림
      rowNumbers.forEach(function (rn) {
        sheet.deleteRow(rn);
      });
      return jsonOut_({ ok: true, deleted: rowNumbers.length });
    }

    return jsonOut_({ ok: false, error: "UNKNOWN_ACTION: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

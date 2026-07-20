// ==========================================================
// 건의함 - 학생이 실명(학번+이름)으로 건의사항을 남기고,
// 체육교사 중 아이디가 정확히 "박정환"인 계정만 조회할 수 있다.
// 익명이 아니므로 학생 본인도 자신이 제출한 건의 이력을 볼 수 있게 한다.
// ==========================================================
const SUGGESTION_VIEWER_TEACHER_ID = '박정환';

function setupSuggestionSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('건의함');
  if (!sheet) {
    sheet = ss.insertSheet('건의함');
    sheet.appendRow(['건의ID', '학번', '이름', '내용', '등록일시']);
    sheet.getRange("A1:E1").setBackground("#455a64").setFontColor("white").setFontWeight("bold");
  }
}

function isSuggestionViewer_(session) {
  return session.role === '체육교사' && session.id === SUGGESTION_VIEWER_TEACHER_ID;
}

function submitSuggestion(content, token) {
  try {
    const session = requireSession(token);
    if (session.role !== '학생') return { success: false, message: "학생만 건의사항을 제출할 수 있습니다." };
    content = (content || '').toString().trim();
    if (!content) return { success: false, message: "건의 내용을 입력하세요." };
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('건의함').appendRow([id, session.id, session.name, content, dateStr]);
    return { success: true, message: "건의사항이 접수되었습니다. 소중한 의견 감사합니다!" };
  } catch (e) { return { success: false, message: e.message }; }
}

// 박정환 계정만 전체 목록 조회 가능. 다른 체육교사/일반교사는 권한 오류.
function getSuggestions(token) {
  try {
    const session = requireSession(token);
    if (!isSuggestionViewer_(session)) return { success: false, message: "이 건의함은 열람 권한이 없습니다." };
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('건의함').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push({ id: data[i][0], studentId: data[i][1], name: data[i][2], content: data[i][3], date: toDateTimeStr_(data[i][4]) });
    }
    list.reverse();
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

function deleteSuggestion(suggestionId, token) {
  try {
    const session = requireSession(token);
    if (!isSuggestionViewer_(session)) return { success: false, message: "이 건의함은 처리 권한이 없습니다." };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('건의함');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === suggestionId) { sheet.deleteRow(i + 1); return { success: true, message: "삭제했습니다." }; } }
    return { success: false, message: "건의 내역을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 학생 본인이 과거에 제출한 건의 이력 (실명제라 본인 것은 볼 수 있게)
function getMySuggestions(token) {
  try {
    const session = requireSession(token);
    if (session.role !== '학생') return { success: true, list: [] };
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('건의함').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === session.id.toString()) list.push({ id: data[i][0], content: data[i][3], date: toDateTimeStr_(data[i][4]) });
    }
    list.reverse();
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

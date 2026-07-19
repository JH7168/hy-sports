// ==========================================================
// 공지사항 로직
// ==========================================================
function getBoardData(token) {
  try {
    requireSession(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('공지사항');
    if (!sheet) { sheet = ss.insertSheet('공지사항'); sheet.appendRow(['날짜', '제목', '내용']); }
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue;
      let rawDate = data[i][0];
      let dateStr = rawDate instanceof Date ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : rawDate.toString();
      result.push({ date: dateStr, title: data[i][1], content: data[i][2] });
    }
    return result.reverse();
  } catch (e) { return []; }
}

function saveNotice(title, content, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('공지사항');
    if (!sheet) { sheet = ss.insertSheet('공지사항'); sheet.appendRow(['날짜', '제목', '내용']); sheet.getRange("A1:C1").setBackground("#1e3c72").setFontColor("white").setFontWeight("bold"); }
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([dateStr, title, content]);
    return { success: true, message: "공지사항이 성공적으로 등록되었습니다!" };
  } catch(e) { return { success: false, message: "공지 등록 실패: " + e.message }; }
}

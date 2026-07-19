// ==========================================================
// 물품 관리 데이터베이스 (Inventory & Purchase) 로직
// ==========================================================
function saveInventoryItem(loc, name, spec, qty, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('체육물품대장');
    if (!sheet) return { success: false, message: "물품대장 시트가 존재하지 않습니다." };
    sheet.appendRow([loc, name, spec, qty]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getInventoryData(location, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('체육물품대장');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      let rowLoc = data[i][0];
      if (!rowLoc) continue;
      if (location === '전체' || rowLoc === location) {
        result.push({ loc: rowLoc, name: data[i][1], spec: data[i][2], qty: data[i][3] });
      }
    }
    return result;
  } catch (e) { return []; }
}

function savePurchaseRequest(data, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('물품구입신청');
    if (!sheet) return {success: false, message: "신청 시트가 없습니다."};

    const seq = sheet.getLastRow();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    sheet.appendRow([seq, data.item, data.spec, data.qty, data.price, data.total, data.teacher, dateStr]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function getPurchaseRequests(token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('물품구입신청');
    let totals = { '박정환': 0, '문승연': 0, '양서경': 0 };
    if (!sheet) return { list: [], totals: totals };

    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (!row[0]) continue;
      list.push(row);

      let teacherName = row[6];
      let totalAmount = parseInt(row[5]) || 0;
      if (totals[teacherName] !== undefined) {
        totals[teacherName] += totalAmount;
      }
    }
    return { list: list.reverse(), totals: totals };
  } catch (e) { return { list: [], totals: { '박정환': 0, '문승연': 0, '양서경': 0 } }; }
}

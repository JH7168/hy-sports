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
    return sanitizeDates_({ list: list.reverse(), totals: totals });
  } catch (e) { return { list: [], totals: { '박정환': 0, '문승연': 0, '양서경': 0 } }; }
}

// ==========================================================
// 예산 관리 - 등록/삭제는 박정환 선생님만, 조회는 체육교사 전체
// ==========================================================
const BUDGET_MANAGER_ID = '박정환';
function isBudgetManager_(session) { return session.role === '체육교사' && session.id === BUDGET_MANAGER_ID; }

function saveBudgetItem(budgetType, allocated, used, token) {
  try {
    const session = requirePeTeacher(token);
    if (!isBudgetManager_(session)) return { success: false, message: "예산 등록은 박정환 선생님만 할 수 있습니다." };
    budgetType = (budgetType || '').toString().trim();
    if (!budgetType) return { success: false, message: "예산 종류를 입력하세요." };
    const allocatedNum = parseInt(allocated, 10) || 0;
    const usedNum = parseInt(used, 10) || 0;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('예산관리');
    if (!sheet) return { success: false, message: "예산관리 시트가 존재하지 않습니다." };
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([id, budgetType, allocatedNum, usedNum, allocatedNum - usedNum, dateStr, session.name]);
    return { success: true, message: "예산이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getBudgetList(token) {
  try {
    const session = requirePeTeacher(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('예산관리');
    const data = sheet ? sheet.getDataRange().getValues() : [];
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push({ id: data[i][0], type: data[i][1], allocated: data[i][2], used: data[i][3], balance: data[i][4], regDate: toDateTimeStr_(data[i][5]), regBy: data[i][6] });
    }
    return sanitizeDates_({ success: true, list: list, isManager: isBudgetManager_(session) });
  } catch (e) { return { success: false, message: e.message, list: [], isManager: false }; }
}

function removeBudgetItem(budgetId, token) {
  try {
    const session = requirePeTeacher(token);
    if (!isBudgetManager_(session)) return { success: false, message: "예산 삭제는 박정환 선생님만 할 수 있습니다." };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('예산관리');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === budgetId) { sheet.deleteRow(i + 1); return { success: true, message: "예산 항목을 삭제했습니다." }; } }
    return { success: false, message: "찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

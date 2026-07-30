// ==========================================================
// 물품 관리 데이터베이스 (Inventory & Purchase) 로직
// ==========================================================
function saveInventoryItem(loc, name, content, spec, qty, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('체육물품대장');
    if (!sheet) return { success: false, message: "물품대장 시트가 존재하지 않습니다." };
    sheet.appendRow([loc, name, content, spec, qty]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateInventoryItem(rowIndex, loc, name, content, spec, qty, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('체육물품대장');
    if (!sheet) return { success: false, message: "물품대장 시트가 존재하지 않습니다." };
    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, message: "대상 물품을 찾을 수 없습니다." };
    sheet.getRange(rowIndex, 1, 1, 5).setValues([[loc, name, content, spec, qty]]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
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
        result.push({ rowIndex: i + 1, loc: rowLoc, name: data[i][1], content: data[i][2], spec: data[i][3], qty: data[i][4] });
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

    sheet.appendRow([seq, data.item, data.spec, data.qty, data.price, data.total, data.teacher, dateStr, data.budgetId || '']);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function updatePurchaseRequest(rowIndex, item, spec, qty, price, teacher, budgetId, token) {
  try {
    requirePeTeacher(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('물품구입신청');
    if (!sheet) return { success: false, message: "신청 시트가 없습니다." };
    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, message: "대상 신청 내역을 찾을 수 없습니다." };
    const qtyNum = parseInt(qty, 10) || 0;
    const priceNum = parseInt(price, 10) || 0;
    sheet.getRange(rowIndex, 2, 1, 6).setValues([[item, spec, qtyNum, priceNum, qtyNum * priceNum, teacher]]);
    sheet.getRange(rowIndex, 9).setValue(budgetId || '');
    return { success: true, message: "구입 신청 내역을 수정했습니다." };
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
      list.push({ rowIndex: i + 1, seq: row[0], item: row[1], spec: row[2], qty: row[3], price: row[4], total: row[5], teacher: row[6], budgetId: row[8] || '' });

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

function saveBudgetItem(detail, costCategory, calcBasis, currentAmount, remainingAmount, token) {
  try {
    const session = requirePeTeacher(token);
    if (!isBudgetManager_(session)) return { success: false, message: "예산 등록은 박정환 선생님만 할 수 있습니다." };
    detail = (detail || '').toString().trim();
    if (!detail) return { success: false, message: "세부항목을 입력하세요." };
    costCategory = (costCategory || '').toString().trim();
    calcBasis = (calcBasis || '').toString().trim();
    const currentNum = parseInt(currentAmount, 10) || 0;
    const remainingNum = parseInt(remainingAmount, 10) || 0;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('예산관리');
    if (!sheet) return { success: false, message: "예산관리 시트가 존재하지 않습니다." };
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([id, detail, costCategory, calcBasis, currentNum, remainingNum, dateStr, session.name]);
    return { success: true, message: "예산이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateBudgetItem(budgetId, detail, costCategory, calcBasis, currentAmount, remainingAmount, token) {
  try {
    const session = requirePeTeacher(token);
    if (!isBudgetManager_(session)) return { success: false, message: "예산 수정은 박정환 선생님만 할 수 있습니다." };
    detail = (detail || '').toString().trim();
    if (!detail) return { success: false, message: "세부항목을 입력하세요." };
    costCategory = (costCategory || '').toString().trim();
    calcBasis = (calcBasis || '').toString().trim();
    const currentNum = parseInt(currentAmount, 10) || 0;
    const remainingNum = parseInt(remainingAmount, 10) || 0;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('예산관리');
    if (!sheet) return { success: false, message: "예산관리 시트가 존재하지 않습니다." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === budgetId) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[detail, costCategory, calcBasis, currentNum, remainingNum]]);
        return { success: true, message: "예산 항목을 수정했습니다." };
      }
    }
    return { success: false, message: "예산 항목을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 세부항목이 같은 것끼리 먼저 묶고, 그 안에서는 원가통계비목이 같은 것끼리 묶어서 보여준다.
function getBudgetList(token) {
  try {
    const session = requirePeTeacher(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('예산관리');
    const data = sheet ? sheet.getDataRange().getValues() : [];
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push({ id: data[i][0], detail: data[i][1], costCategory: data[i][2], calcBasis: data[i][3], currentAmount: data[i][4], remainingAmount: data[i][5], regDate: toDateTimeStr_(data[i][6]), regBy: data[i][7] });
    }
    list.sort((a, b) => a.detail.toString().localeCompare(b.detail.toString(), 'ko') || a.costCategory.toString().localeCompare(b.costCategory.toString(), 'ko'));
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

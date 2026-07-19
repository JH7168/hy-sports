// ==========================================================
// 체육한마당 로직
// ==========================================================
function setupSportsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['체육한마당_임시', '체육한마당_배포'].forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(['연도', '학년', '종목', '대진표_JSON', '일정_JSON']);
      sheet.getRange("A1:E1").setBackground("#1e3c72").setFontColor("white").setFontWeight("bold");
    }
  });
}

function saveSportsDraft(year, grade, sport, bracketData, scheduleData, token) {
  try {
    requirePeTeacher(token);
    setupSportsSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체육한마당_임시');
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == year && data[i][1] == grade && data[i][2] == sport) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 4, 1, 2).setValues([[JSON.stringify(bracketData), JSON.stringify(scheduleData)]]);
    } else {
      sheet.appendRow([year, grade, sport, JSON.stringify(bracketData), JSON.stringify(scheduleData)]);
    }
    return { success: true, message: "💾 임시저장 완료!" };
  } catch (e) { return { success: false, message: e.message }; }
}

function publishSportsData(year, grade, sport, token) {
  try {
    requirePeTeacher(token);
    setupSportsSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const draftSheet = ss.getSheetByName('체육한마당_임시');
    const liveSheet = ss.getSheetByName('체육한마당_배포');
    if (!draftSheet) return { success: false, message: "⚠️ 임시저장된 데이터가 없습니다. 먼저 [임시 저장] 버튼을 눌러주세요." };

    const draftData = draftSheet.getDataRange().getValues();
    let targetDraft = null;
    for (let i = 1; i < draftData.length; i++) {
      if (draftData[i][0] == year && draftData[i][1] == grade && draftData[i][2] == sport) { targetDraft = draftData[i]; break; }
    }
    if(!targetDraft) return { success: false, message: "⚠️ 해당 연도/학년/종목으로 임시저장된 대진표가 없습니다." };

    const liveData = liveSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < liveData.length; i++) {
      if (liveData[i][0] == year && liveData[i][1] == grade && liveData[i][2] == sport) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) {
      liveSheet.getRange(rowIndex, 4, 1, 2).setValues([[targetDraft[3], targetDraft[4]]]);
    } else {
      liveSheet.appendRow([year, grade, sport, targetDraft[3], targetDraft[4]]);
    }
    return { success: true, message: "🎉 대진표 배포 성공!" };
  } catch (e) { return { success: false, message: "시스템 오류가 발생했습니다: " + e.message }; }
}

function unpublishSportsData(year, grade, sport, token) {
  try {
    requirePeTeacher(token);
    setupSportsSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const liveSheet = ss.getSheetByName('체육한마당_배포');

    if (liveSheet.getLastRow() < 2) return { success: false, message: "⚠️ 현재 배포된 대진표 데이터가 없습니다." };

    const liveData = liveSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < liveData.length; i++) {
      if (liveData[i][0] == year && liveData[i][1] == grade && liveData[i][2] == sport) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) {
      liveSheet.deleteRow(rowIndex);
      return { success: true, message: "🛑 배포가 성공적으로 취소되었습니다. 학생들 화면에서 대진표가 내려갔습니다." };
    } else {
      return { success: false, message: "⚠️ 해당 조건으로 배포된 대진표를 찾을 수 없습니다." };
    }
  } catch (e) { return { success: false, message: "시스템 오류가 발생했습니다: " + e.message }; }
}

function getSportsLiveData(year, grade, sport, token) {
  try {
    requireSession(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const liveSheet = ss.getSheetByName('체육한마당_배포');
    if (!liveSheet || liveSheet.getLastRow() < 2) return { success: false, message: "아직 선생님께서 배포하신 대진표가 없습니다." };

    const data = liveSheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == year && data[i][1] == grade && data[i][2] == sport) {
        return { success: true, bracket: JSON.parse(data[i][3]), schedule: JSON.parse(data[i][4]) };
      }
    }
    return { success: false, message: "선택하신 학년 및 종목에 배포된 대진표가 없습니다." };
  } catch(e) { return { success: false, message: e.message }; }
}

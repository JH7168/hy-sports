// ==========================================================
// 부트스트랩: 페이지 진입점 및 초기 시트/설정 세팅
// 나머지 로직은 Auth.js / Board.js / Paps.js / Sports.js /
// Inventory.js / Utils.js 로 기능별 분리되어 있다.
// ==========================================================
function include(filename) { return HtmlService.createTemplateFromFile(filename).evaluate().getContent(); }

function setupSystemSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recentYears = getRecentAcademicYears(3);

  let oldSheet = ss.getSheetByName('PAPS_상세');
  if (oldSheet && !ss.getSheetByName('PAPS_상세_' + recentYears[0])) {
    oldSheet.setName('PAPS_상세_' + recentYears[0]);
  }

  recentYears.forEach(year => {
    let sheet = ss.getSheetByName('PAPS_상세_' + year);
    if (!sheet) {
      sheet = ss.insertSheet('PAPS_상세_' + year);
      const papsHeaders = ['학번', '이름', '종합점수', '종합등급', '왕오달_기록', '왕오달_점수', '왕오달_등급', '앉아윗몸_기록', '앉아윗몸_점수', '앉아윗몸_등급', '팔굽_기록', '팔굽_점수', '팔굽_등급', '제멀_기록', '제멀_점수', '제멀_등급', 'BMI_기록', 'BMI_점수', 'BMI_등급'];
      sheet.getRange(1, 1, 1, papsHeaders.length).setValues([papsHeaders]);
      sheet.getRange("A1:S1").setBackground("#1e3c72").setFontColor("white").setFontWeight("bold");
    }
  });

  ['체육한마당_임시', '체육한마당_배포'].forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(['연도', '학년', '종목', '대진표_JSON', '일정_JSON']);
      sheet.getRange("A1:E1").setBackground("#1e3c72").setFontColor("white").setFontWeight("bold");
    }
  });

  let memberSheet = ss.getSheetByName('회원정보');
  if (!memberSheet) {
    memberSheet = ss.insertSheet('회원정보');
    memberSheet.appendRow(['구분', '아이디', '비밀번호']);
    memberSheet.getRange("A1:C1").setBackground("#1e3c72").setFontColor("white").setFontWeight("bold");
    memberSheet.getRange("B:C").setNumberFormat("@");
  }

  // 체육물품대장 시트 세팅
  let inventorySheet = ss.getSheetByName('체육물품대장');
  if (!inventorySheet) {
    inventorySheet = ss.insertSheet('체육물품대장');
    inventorySheet.appendRow(['장소', '물품명', '규격', '수량']);
    inventorySheet.getRange("A1:D1").setBackground("#4caf50").setFontColor("white").setFontWeight("bold");
  }

  // 물품구입신청 시트 세팅
  let purchaseSheet = ss.getSheetByName('물품구입신청');
  if (!purchaseSheet) {
    purchaseSheet = ss.insertSheet('물품구입신청');
    purchaseSheet.appendRow(['순번', '내용', '규격', '수량', '예상단가', '예상금액', '신청교사', '신청일시']);
    purchaseSheet.getRange("A1:H1").setBackground("#2196f3").setFontColor("white").setFontWeight("bold");
  }

  // 소규모수업 - 체대입시반 시트 세팅
  setupPhysPrepSheets();

  // 교사 가입 인증 코드: 소스코드가 아닌 스크립트 속성에 보관 (없을 때만 최초 1회 초기화)
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PE_TEACHER_CODE')) props.setProperty('PE_TEACHER_CODE', 'PY4312');
  if (!props.getProperty('TEACHER_CODE')) props.setProperty('TEACHER_CODE', 'HY4312');
}

function doGet() {
  setupSystemSheets();
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('한영고 체육인성부').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

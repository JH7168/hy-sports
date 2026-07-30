// ==========================================================
// 부트스트랩: 페이지 진입점 및 초기 시트/설정 세팅
// 나머지 로직은 Auth.js / Board.js / Paps.js / Sports.js /
// Inventory.js / Utils.js 로 기능별 분리되어 있다.
// ==========================================================
function include(filename) { return HtmlService.createTemplateFromFile(filename).evaluate().getContent(); }

// setupSystemSheets()는 시트 존재 여부를 수십 번 확인하는 무거운 작업이라, 예전처럼
// 페이지를 열 때마다(doGet마다) 매번 실행하면 그만큼 로딩이 느려진다. 시트 구조는
// 한 번만 맞춰두면 되므로 스크립트 속성에 완료 버전을 기록해두고 그 이후로는 건너뛴다.
// 이후 시트/컬럼 구조를 바꾸는 코드를 추가할 때는 이 버전 숫자를 올려야 기존
// 스프레드시트에도 그 변경(마이그레이션)이 한 번 더 반영된다.
const SETUP_VERSION = '5';

function ensureSystemSheetsSetup_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SETUP_VERSION') === SETUP_VERSION) return;
  setupSystemSheets();
  props.setProperty('SETUP_VERSION', SETUP_VERSION);
}

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
    inventorySheet.appendRow(['장소', '물품명', '내용', '규격', '수량']);
    inventorySheet.getRange("A1:E1").setBackground("#4caf50").setFontColor("white").setFontWeight("bold");
  } else if (inventorySheet.getRange(1, 3).getValue() !== '내용') {
    // 예전에는 물품명 칸에 "물품명(내용)"처럼 괄호로 합쳐서 적었다. 내용 칸을 새로 끼워넣으면서
    // 기존 값도 그 괄호 패턴이면 자동으로 분리해준다(패턴이 아니면 물품명 칸에 그대로 둔다).
    inventorySheet.insertColumnAfter(2);
    inventorySheet.getRange(1, 3).setValue('내용');
    inventorySheet.getRange("A1:E1").setBackground("#4caf50").setFontColor("white").setFontWeight("bold");
    const lastRow = inventorySheet.getLastRow();
    if (lastRow > 1) {
      const names = inventorySheet.getRange(2, 2, lastRow - 1, 1).getValues();
      const newNames = []; const contents = [];
      names.forEach(r => {
        const raw = (r[0] || '').toString().trim();
        const match = raw.match(/^(.+?)\(([^()]+)\)$/);
        if (match) { newNames.push([match[1].trim()]); contents.push([match[2].trim()]); }
        else { newNames.push([raw]); contents.push(['']); }
      });
      inventorySheet.getRange(2, 2, newNames.length, 1).setValues(newNames);
      inventorySheet.getRange(2, 3, contents.length, 1).setValues(contents);
    }
  }

  // 물품구입신청 시트 세팅
  const purchaseHeaders = ['순번', '내용', '규격', '수량', '예상단가', '예상금액', '신청교사', '신청일시', '예산ID'];
  let purchaseSheet = ss.getSheetByName('물품구입신청');
  if (!purchaseSheet) {
    purchaseSheet = ss.insertSheet('물품구입신청');
    purchaseSheet.getRange(1, 1, 1, purchaseHeaders.length).setValues([purchaseHeaders]);
    purchaseSheet.getRange(1, 1, 1, purchaseHeaders.length).setBackground("#2196f3").setFontColor("white").setFontWeight("bold");
  } else if (purchaseSheet.getRange(1, 9).getValue() !== '예산ID') {
    // 이 신청이 어느 예산 항목에서 지출되는지 연결해 반영 후 예상 잔액을 보여주기 위한 컬럼.
    purchaseSheet.getRange(1, 1, 1, purchaseHeaders.length).setValues([purchaseHeaders]);
    purchaseSheet.getRange(1, 1, 1, purchaseHeaders.length).setBackground("#2196f3").setFontColor("white").setFontWeight("bold");
  }

  // 예산관리 시트 세팅 (박정환 선생님만 등록/삭제 가능)
  const budgetHeaders = ['예산ID', '세부항목', '원가통계비목', '산출내역', '예산현액', '예산잔액', '등록일시', '등록자'];
  let budgetSheet = ss.getSheetByName('예산관리');
  if (!budgetSheet) {
    budgetSheet = ss.insertSheet('예산관리');
    budgetSheet.getRange(1, 1, 1, budgetHeaders.length).setValues([budgetHeaders]);
    budgetSheet.getRange(1, 1, 1, budgetHeaders.length).setBackground("#1e3c72").setFontColor("white").setFontWeight("bold");
  } else if (budgetSheet.getRange(1, 2).getValue() !== '세부항목') {
    // 예전에는 예산종류/편성금액/사용금액/잔액 형식이었다. 세부항목=예산종류, 예산현액=편성금액,
    // 예산잔액=잔액으로 옮기고(사용금액 개념은 새 형식에 없어 버린다), 원가통계비목·산출내역은
    // 기존에 없던 항목이라 빈칸으로 시작한다.
    const oldData = budgetSheet.getDataRange().getValues();
    const migrated = [budgetHeaders];
    for (let i = 1; i < oldData.length; i++) {
      const row = oldData[i];
      if (!row[0]) continue;
      migrated.push([row[0], row[1], '', '', row[2], row[4], row[5], row[6]]);
    }
    budgetSheet.clear();
    budgetSheet.getRange(1, 1, migrated.length, budgetHeaders.length).setValues(migrated);
    budgetSheet.getRange(1, 1, 1, budgetHeaders.length).setBackground("#1e3c72").setFontColor("white").setFontWeight("bold");
  }

  // 소규모수업 - 체대입시반 시트 세팅
  setupPhysPrepSheets();

  // 점심시간 - 풋살장/체육관/운동장 시설 예약 시트 세팅
  setupFacilitySheets();

  // 스포츠클럽(축구·농구·배구·배드민턴부) 시트 세팅
  setupSportsClubSheets();

  // 건의함 시트 세팅
  setupSuggestionSheets();

  // 친사리그 시트 세팅
  setupLeagueSheets();

  // 교사 가입 인증 코드: 소스코드가 아닌 스크립트 속성에 보관 (없을 때만 최초 1회 초기화)
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PE_TEACHER_CODE')) props.setProperty('PE_TEACHER_CODE', 'PY4312');
  if (!props.getProperty('TEACHER_CODE')) props.setProperty('TEACHER_CODE', 'HY4312');
}

function doGet() {
  ensureSystemSheetsSetup_();
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('한영고 체육인성부').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

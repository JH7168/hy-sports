function include(filename) { return HtmlService.createTemplateFromFile(filename).evaluate().getContent(); }

function setupSystemSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let oldSheet = ss.getSheetByName('PAPS_상세');
  if (oldSheet && !ss.getSheetByName('PAPS_상세_2026')) {
    oldSheet.setName('PAPS_상세_2026');
  }

  ['2026', '2025', '2024'].forEach(year => {
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
}

function doGet() {
  setupSystemSheets();
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('한영고 체육부').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================================
// 공지사항 로직
// ==========================================================
function getBoardData() {
  try {
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

function saveNotice(title, content) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('공지사항');
    if (!sheet) { sheet = ss.insertSheet('공지사항'); sheet.appendRow(['날짜', '제목', '내용']); sheet.getRange("A1:C1").setBackground("#1e3c72").setFontColor("white").setFontWeight("bold"); }
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([dateStr, title, content]);
    return { success: true, message: "공지사항이 성공적으로 등록되었습니다!" };
  } catch(e) { return { success: false, message: "공지 등록 실패: " + e.message }; }
}

// ==========================================================
// 인증 및 회원가입 로직
// ==========================================================
function verifyLogin(role, userId, userPw) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (role === '학생') {
      const sheet = ss.getSheetByName('학생명렬표');
      if (!sheet) return { success: false, message: "'학생명렬표' 탭이 없습니다." };
      const data = sheet.getDataRange().getValues();
      const inputId = userId.toString().trim();
      const inputPw = userPw.toString().trim();
      for (let i = 20; i < data.length; i++) { 
        if (data[i][0] && data[i][0].toString().trim() === inputId) {
          if (data[i][2] && data[i][2].toString().trim() === inputPw) return { success: true, name: data[i][1].toString().trim(), role: '학생', id: inputId };
          else return { success: false, message: "비밀번호가 일치하지 않습니다." };
        }
      }
      return { success: false, message: "가입되지 않은 학생이거나 학번이 틀렸습니다." };
    } else {
      const sheet = ss.getSheetByName('회원정보');
      if (!sheet) return { success: false, message: "'회원정보' 탭이 없습니다." };
      const data = sheet.getDataRange().getValues();
      const inputId = userId.toString().trim();
      const inputPw = userPw.toString().trim();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === role && data[i][1].toString().trim() === inputId && data[i][2].toString().trim() === inputPw) return { success: true, name: data[i][1].toString().trim(), role: data[i][0], id: inputId };
      }
      return { success: false, message: "정보가 일치하지 않습니다." };
    }
  } catch (e) { return { success: false, message: "시스템 오류" }; }
}

function findStudent(grade, cls, num) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('학생명렬표');
    if(!sheet) return { success: false, message: "명렬표 탭이 없습니다." };
    const formatNum = (n) => n.toString().trim().padStart(2, '0');
    const targetId = grade.toString().trim() + formatNum(cls) + formatNum(num);
    const data = sheet.getDataRange().getValues();
    for(let i = 20; i < data.length; i++) {
      if(data[i][0] && data[i][0].toString().trim() === targetId) return { success: true, name: data[i][1].toString().trim() };
    }
    return { success: false, message: "명렬표에 없는 학생입니다." };
  } catch(e) { return { success: false, message: "검색 오류" }; }
}

function registerUser(role, id, pw, name, secretCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inputId = id.toString().trim();
    const inputPw = pw.toString().trim();
    if (role === '학생') {
      const sheet = ss.getSheetByName('학생명렬표');
      if (!sheet) return { success: false, message: "⚠️ '학생명렬표' 탭을 찾을 수 없습니다. 관리자에게 문의하세요." };
      sheet.getRange("C20").setValue("비밀번호");
      const data = sheet.getDataRange().getValues();
      if (data.length < 21) return { success: false, message: "⚠️ 명렬표에 등록된 학생 데이터가 없습니다." };
      for(let i = 20; i < data.length; i++) {
        if(data[i][0] && data[i][0].toString().trim() === inputId) {
          if(data[i][2] && data[i][2].toString().trim() !== "") return { success: false, message: "이미 가입된 학생입니다." };
          sheet.getRange(i + 1, 3).setNumberFormat("@").setValue(inputPw);
          return { success: true, message: "학생 회원가입이 완료되었습니다! 로그인해주세요." };
        }
      }
      return { success: false, message: "명렬표에서 학생을 찾을 수 없습니다." };
    } else {
      let expectedCode = (role === '체육교사') ? "PY4312" : "HY4312";
      if (secretCode.toString().trim() !== expectedCode) return { success: false, message: "학교 가입 인증 코드가 틀렸습니다." };
      const ssSheet = ss.getSheetByName('회원정보');
      const data = ssSheet.getDataRange().getValues();
      for(let i = 1; i < data.length; i++) {
        if(data[i][0] === role && data[i][1].toString().trim() === inputId) return { success: false, message: "이미 가입된 아이디입니다." };
      }
      const nextRow = ssSheet.getLastRow() + 1;
      ssSheet.getRange(nextRow, 1).setValue(role);
      ssSheet.getRange(nextRow, 2).setNumberFormat("@").setValue(inputId);
      ssSheet.getRange(nextRow, 3).setNumberFormat("@").setValue(inputPw);
      return { success: true, message: "교사 회원가입이 완료되었습니다! 로그인해주세요." };
    }
  } catch(e) { return { success: false, message: "가입 처리 중 오류 발생: " + e.message }; }
}

// ==========================================================
// PAPS 로직
// ==========================================================
function uploadPapsPdf(data, filename) {
  try {
    const FOLDER_ID = "1m31h6WC-EGuVHjaHi1ecBKdSDRBvOAvP"; 
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const existing = folder.getFilesByName(filename);
    while(existing.hasNext()) { existing.next().setTrashed(true); }
    const contentType = data.substring(5, data.indexOf(';'));
    const bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,') + 7));
    const blob = Utilities.newBlob(bytes, contentType, filename);
    folder.createFile(blob);
    return { success: true, message: filename + " 업로드 성공!" };
  } catch(e) { return { success: false, message: "업로드 실패: " + e.message }; }
}

function clearAllPapsPdf() {
  try {
    const FOLDER_ID = "1m31h6WC-EGuVHjaHi1ecBKdSDRBvOAvP"; 
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    let deleteCount = 0;
    while (files.hasNext()) { let file = files.next(); file.setTrashed(true); deleteCount++; }
    return { success: true, message: `지정 폴더 안의 PDF 파일 총 ${deleteCount}건을 깨끗하게 정리했습니다!` };
  } catch(e) { return { success: false, message: "드라이브 비우기 중 오류 발생: " + e.message }; }
}

function getAdminGradeStats(year) {
  try {
    if (!year) year = "2026";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PAPS_상세_' + year);
    if (!sheet) return { success: false, message: `${year}학년도 데이터가 없습니다.` };

    const data = sheet.getDataRange().getValues();
    let stats = {
      "1": { count: 0, grade1: 0, grade2: 0, grade3: 0, grade4: 0, grade5: 0, totalScore: 0, totalGrade: 0, cardioRec: 0, cardioScore: 0, cardioGrade: 0, flexRec: 0, flexScore: 0, flexGrade: 0, strengthRec: 0, strengthScore: 0, strengthGrade: 0, powerRec: 0, powerScore: 0, powerGrade: 0, bmiRec: 0, bmiScore: 0, bmiGrade: 0 },
      "2": { count: 0, grade1: 0, grade2: 0, grade3: 0, grade4: 0, grade5: 0, totalScore: 0, totalGrade: 0, cardioRec: 0, cardioScore: 0, cardioGrade: 0, flexRec: 0, flexScore: 0, flexGrade: 0, strengthRec: 0, strengthScore: 0, strengthGrade: 0, powerRec: 0, powerScore: 0, powerGrade: 0, bmiRec: 0, bmiScore: 0, bmiGrade: 0 },
      "3": { count: 0, grade1: 0, grade2: 0, grade3: 0, grade4: 0, grade5: 0, totalScore: 0, totalGrade: 0, cardioRec: 0, cardioScore: 0, cardioGrade: 0, flexRec: 0, flexScore: 0, flexGrade: 0, strengthRec: 0, strengthScore: 0, strengthGrade: 0, powerRec: 0, powerScore: 0, powerGrade: 0, bmiRec: 0, bmiScore: 0, bmiGrade: 0 }
    };

    function parseNum(val) {
      if (val === undefined || val === null || val === "") return 0;
      let cleaned = val.toString().replace(/[^\d.]/g, '');
      let num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i]; if (!row[0]) continue;
      let grade = row[0].toString().trim().charAt(0);
      if (!stats[grade]) continue;

      stats[grade].count++;
      let tGrade = parseNum(row[3]);
      stats[grade].totalGrade += tGrade;
      stats[grade].totalScore += parseNum(row[2]);
      
      if (tGrade === 1) stats[grade].grade1++;
      if (tGrade === 2) stats[grade].grade2++;
      if (tGrade === 3) stats[grade].grade3++;
      if (tGrade === 4) stats[grade].grade4++;
      if (tGrade === 5) stats[grade].grade5++;
      
      stats[grade].cardioRec += parseNum(row[4]); stats[grade].cardioScore += parseNum(row[5]); stats[grade].cardioGrade += parseNum(row[6]);
      stats[grade].flexRec += parseNum(row[7]); stats[grade].flexScore += parseNum(row[8]); stats[grade].flexGrade += parseNum(row[9]);
      stats[grade].strengthRec += parseNum(row[10]); stats[grade].strengthScore += parseNum(row[11]); stats[grade].strengthGrade += parseNum(row[12]);
      stats[grade].powerRec += parseNum(row[13]); stats[grade].powerScore += parseNum(row[14]); stats[grade].powerGrade += parseNum(row[15]);
      stats[grade].bmiRec += parseNum(row[16]); stats[grade].bmiScore += parseNum(row[17]); stats[grade].bmiGrade += parseNum(row[18]); 
    }

    let result = {};
    for (let g in stats) {
      let cnt = stats[g].count || 1;
      let targetCnt = stats[g].grade4 + stats[g].grade5;
      result[g] = {
        totalStudents: stats[g].count, grade1: stats[g].grade1, grade2: stats[g].grade2, grade3: stats[g].grade3, grade4: stats[g].grade4, grade5: stats[g].grade5,
        targetCount: targetCnt, targetRatio: stats[g].count > 0 ? ((targetCnt / stats[g].count) * 100).toFixed(1) : 0,
        totalScore: (stats[g].totalScore / cnt).toFixed(1), totalGrade: (stats[g].totalGrade / cnt).toFixed(1),
        cardio: { rec: (stats[g].cardioRec / cnt).toFixed(1), score: (stats[g].cardioScore / cnt).toFixed(1), grade: (stats[g].cardioGrade / cnt).toFixed(1) },
        flex: { rec: (stats[g].flexRec / cnt).toFixed(1), score: (stats[g].flexScore / cnt).toFixed(1), grade: (stats[g].flexGrade / cnt).toFixed(1) },
        strength: { rec: (stats[g].strengthRec / cnt).toFixed(1), score: (stats[g].strengthScore / cnt).toFixed(1), grade: (stats[g].strengthGrade / cnt).toFixed(1) },
        power: { rec: (stats[g].powerRec / cnt).toFixed(1), score: (stats[g].powerScore / cnt).toFixed(1), grade: (stats[g].powerGrade / cnt).toFixed(1) },
        bmi: { rec: (stats[g].bmiRec / cnt).toFixed(1), score: (stats[g].bmiScore / cnt).toFixed(1), grade: (stats[g].bmiGrade / cnt).toFixed(1) }
      };
    }
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function getStudentPapsDetail(studentId, year) {
  try {
    if (!year) year = "2026";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PAPS_상세_' + year); 
    if (!sheet) return { success: false, message: `${year}학년도 평가 데이터가 없습니다.` };
    
    const data = sheet.getDataRange().getValues();
    const idStr = studentId.toString().trim();
    
    let targetRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === idStr) { targetRow = data[i]; break; }
    }
    if (!targetRow) return { success: true, papsData: null, pdfUrl: null };

    function parseId(id) {
      const s = id.toString().trim(); if (!s || s.length < 4) return null;
      const g = s.charAt(0); let c = (s.length === 5) ? s.substring(1, 3) : s.substring(1, 2);
      return { grade: g, cls: parseInt(c, 10).toString() };
    }

    const targetParsed = parseId(idStr);
    if (!targetParsed) return { success: false, message: "학번 형식이 잘못되었습니다." };

    const metrics = [{ key: 'totalScore', idx: 2 }, { key: 'cardioRec', idx: 4 }, { key: 'flexRec', idx: 7 }, { key: 'strengthRec', idx: 10 }, { key: 'powerRec', idx: 13 }, { key: 'bmiRec', idx: 16 }];
    let schoolData = {}; let classData = {}; metrics.forEach(m => { schoolData[m.key] = []; classData[m.key] = []; });
    let bmiScoresSchool = []; let bmiScoresClass = []; let targetBmiScore = parseFloat(targetRow[17]);
    let totalSchoolStudents = 0; let totalClassStudents = 0;
    let classGroup = {}; let schoolScoreGroup = { cardio: [], flex: [], strength: [], power: [], bmi: [] };

    for (let i = 1; i < data.length; i++) {
      const row = data[i]; if (!row[0]) continue;
      const parsed = parseId(row[0]); if (!parsed) continue;

      if (parsed.grade === targetParsed.grade) {
        totalSchoolStudents++;
        const isSameClass = (parsed.cls === targetParsed.cls);
        if (isSameClass) totalClassStudents++; 

        metrics.forEach(m => { let val = parseFloat(row[m.idx]); if (!isNaN(val)) { schoolData[m.key].push(val); if (isSameClass) classData[m.key].push(val); } });
        let bScore = parseFloat(row[17]); if (!isNaN(bScore)) { bmiScoresSchool.push(bScore); if (isSameClass) bmiScoresClass.push(bScore); }

        const cStr = parsed.cls;
        if (!classGroup[cStr]) classGroup[cStr] = { total: [], cardio: [], flex: [], strength: [], power: [], bmi: [] };
        if (!isNaN(parseFloat(row[2]))) classGroup[cStr].total.push(parseFloat(row[2]));    
        if (!isNaN(parseFloat(row[5]))) classGroup[cStr].cardio.push(parseFloat(row[5]));  
        if (!isNaN(parseFloat(row[8]))) classGroup[cStr].flex.push(parseFloat(row[8]));    
        if (!isNaN(parseFloat(row[11]))) classGroup[cStr].strength.push(parseFloat(row[11])); 
        if (!isNaN(parseFloat(row[14]))) classGroup[cStr].power.push(parseFloat(row[14]));  
        if (!isNaN(parseFloat(row[17]))) classGroup[cStr].bmi.push(parseFloat(row[17]));    

        if (!isNaN(parseFloat(row[5]))) schoolScoreGroup.cardio.push(parseFloat(row[5]));
        if (!isNaN(parseFloat(row[8]))) schoolScoreGroup.flex.push(parseFloat(row[8]));
        if (!isNaN(parseFloat(row[11]))) schoolScoreGroup.strength.push(parseFloat(row[11]));
        if (!isNaN(parseFloat(row[14]))) schoolScoreGroup.power.push(parseFloat(row[14]));
        if (!isNaN(parseFloat(row[17]))) schoolScoreGroup.bmi.push(parseFloat(row[17]));
      }
    }

    function getStats(arr, targetVal) {
      if (arr.length === 0 || isNaN(targetVal)) return { avg: "-", rank: "-" };
      const avg = (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(1);
      let rank = 1; arr.forEach(val => { if (val > targetVal) rank++; });
      return { avg: avg, rank: rank };
    }

    function getClassRank(targetClass, metricKey) {
      let classAverages = [];
      for (let cStr in classGroup) {
        let scores = classGroup[cStr][metricKey];
        if (scores.length > 0) classAverages.push({ cls: cStr, avg: scores.reduce((a, b) => a + b, 0) / scores.length });
      }
      classAverages.sort((a, b) => b.avg - a.avg); 
      let rank = 1; let targetAvg = 0;
      let found = classAverages.find(a => a.cls === targetClass); if (found) targetAvg = found.avg;
      classAverages.forEach(a => { if (a.avg > targetAvg) rank++; });
      return { rank: rank, totalClasses: classAverages.length };
    }

    const targetClsStr = targetParsed.cls;
    const classRanks = { total: getClassRank(targetClsStr, 'total'), cardio: getClassRank(targetClsStr, 'cardio'), flex: getClassRank(targetClsStr, 'flex'), strength: getClassRank(targetClsStr, 'strength'), power: getClassRank(targetClsStr, 'power'), bmi: getClassRank(targetClsStr, 'bmi') };
    const calcAvg = (arr) => arr.length > 0 ? (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(1) : 0;
    
    let classScoreAvg = { cardio: 0, flex: 0, strength: 0, power: 0, bmi: 0 };
    if (classGroup[targetClsStr]) {
      classScoreAvg.cardio = calcAvg(classGroup[targetClsStr].cardio); classScoreAvg.flex = calcAvg(classGroup[targetClsStr].flex); classScoreAvg.strength = calcAvg(classGroup[targetClsStr].strength); classScoreAvg.power = calcAvg(classGroup[targetClsStr].power); classScoreAvg.bmi = calcAvg(classGroup[targetClsStr].bmi);
    }
    let schoolScoreAvg = { cardio: calcAvg(schoolScoreGroup.cardio), flex: calcAvg(schoolScoreGroup.flex), strength: calcAvg(schoolScoreGroup.strength), power: calcAvg(schoolScoreGroup.power), bmi: calcAvg(schoolScoreGroup.bmi) };

    const papsData = {
      totalScore: targetRow[2], totalGrade: targetRow[3], cardioRec: targetRow[4], cardioScore: targetRow[5], cardioGrade: targetRow[6], flexRec: targetRow[7], flexScore: targetRow[8], flexGrade: targetRow[9], strengthRec: targetRow[10], strengthScore: targetRow[11], strengthGrade: targetRow[12], powerRec: targetRow[13], powerScore: targetRow[14], powerGrade: targetRow[15], bmiRec: targetRow[16], bmiScore: targetRow[17], bmiGrade: targetRow[18],
      targetClass: targetClsStr, classRanks: classRanks, classScoreAvg: classScoreAvg, schoolScoreAvg: schoolScoreAvg, 
      stats: {
        totalSchoolStudents: totalSchoolStudents, totalClassStudents: totalClassStudents,
        totalScore: { school: getStats(schoolData['totalScore'], parseFloat(targetRow[2])), class: getStats(classData['totalScore'], parseFloat(targetRow[2])) },
        cardio: { school: getStats(schoolData['cardioRec'], parseFloat(targetRow[4])), class: getStats(classData['cardioRec'], parseFloat(targetRow[4])) },
        flex: { school: getStats(schoolData['flexRec'], parseFloat(targetRow[7])), class: getStats(classData['flexRec'], parseFloat(targetRow[7])) },
        strength: { school: getStats(schoolData['strengthRec'], parseFloat(targetRow[10])), class: getStats(classData['strengthRec'], parseFloat(targetRow[10])) },
        power: { school: getStats(schoolData['powerRec'], parseFloat(targetRow[13])), class: getStats(classData['powerRec'], parseFloat(targetRow[13])) },
        bmi: { school: { avg: getStats(schoolData['bmiRec'], parseFloat(targetRow[16])).avg, rank: getStats(bmiScoresSchool, targetBmiScore).rank }, class: { avg: getStats(classData['bmiRec'], parseFloat(targetRow[16])).avg, rank: getStats(bmiScoresClass, targetBmiScore).rank } }
      }
    };

    const FOLDER_ID = "1m31h6WC-EGuVHjaHi1ecBKdSDRBvOAvP"; const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFilesByName(studentId + ".pdf"); let pdfUrl = null;
    if (files.hasNext()) { const file = files.next(); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); pdfUrl = file.getUrl().replace("view?usp=drivesdk", "preview"); }
    return { success: true, papsData: papsData, pdfUrl: pdfUrl };
  } catch (e) { return { success: false, message: "오류 발생: " + e.message }; }
}

function saveNeisPapsData(parsedData, year) {
  try {
    if (!year) year = "2026";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('PAPS_상세_' + year);
    if(!sheet) { setupSystemSheets(); sheet = ss.getSheetByName('PAPS_상세_' + year); }
    const existingData = sheet.getDataRange().getValues();
    const idRowMap = {}; for(let i = 1; i < existingData.length; i++) { if(existingData[i][0]) idRowMap[existingData[i][0].toString()] = i + 1; }
    for(let i = 0; i < parsedData.length; i++) {
      const rowData = parsedData[i]; const studentId = rowData[0].toString();
      if(idRowMap[studentId]) { sheet.getRange(idRowMap[studentId], 1, 1, 19).setValues([rowData]); } else { sheet.appendRow(rowData); }
    }
    return { success: true, message: `${year}학년도 데이터가 자동 동기화되었습니다.` };
  } catch(e) { return { success: false, message: "DB 저장 오류: " + e.message }; }
}

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

function saveSportsDraft(year, grade, sport, bracketData, scheduleData) {
  try {
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

function publishSportsData(year, grade, sport) {
  try {
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

function unpublishSportsData(year, grade, sport) {
  try {
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

function getSportsLiveData(year, grade, sport) {
  try {
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

// ==========================================================
// ★ 신규: 물품 관리 데이터베이스 (Inventory & Purchase) 로직
// ==========================================================

function saveInventoryItem(loc, name, spec, qty) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('체육물품대장');
    if (!sheet) return { success: false, message: "물품대장 시트가 존재하지 않습니다." };
    sheet.appendRow([loc, name, spec, qty]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getInventoryData(location) {
  try {
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

function savePurchaseRequest(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('물품구입신청');
    if (!sheet) return {success: false, message: "신청 시트가 없습니다."};
    
    const seq = sheet.getLastRow(); 
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    
    sheet.appendRow([seq, data.item, data.spec, data.qty, data.price, data.total, data.teacher, dateStr]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function getPurchaseRequests() {
  try {
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
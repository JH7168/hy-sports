// ==========================================================
// PAPS 로직
// ==========================================================
function uploadPapsPdf(data, filename, token) {
  try {
    requirePeTeacher(token);
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

function clearAllPapsPdf(token) {
  try {
    requirePeTeacher(token);
    const FOLDER_ID = "1m31h6WC-EGuVHjaHi1ecBKdSDRBvOAvP";
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFiles();
    let deleteCount = 0;
    while (files.hasNext()) { let file = files.next(); file.setTrashed(true); deleteCount++; }
    return { success: true, message: `지정 폴더 안의 PDF 파일 총 ${deleteCount}건을 깨끗하게 정리했습니다!` };
  } catch(e) { return { success: false, message: "드라이브 비우기 중 오류 발생: " + e.message }; }
}

function getAdminGradeStats(year, token) {
  try {
    requirePeTeacher(token);
    if (!year) year = getCurrentAcademicYear();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PAPS_상세_' + year);
    if (!sheet) return { success: false, message: `${year}학년도 데이터가 없습니다.` };

    const data = sheet.getDataRange().getValues();
    let stats = {
      "1": { count: 0, grade1: 0, grade2: 0, grade3: 0, grade4: 0, grade5: 0, totalScore: 0, totalGrade: 0, cardioRec: 0, cardioScore: 0, cardioGrade: 0, flexRec: 0, flexScore: 0, flexGrade: 0, strengthRec: 0, strengthScore: 0, strengthGrade: 0, powerRec: 0, powerScore: 0, powerGrade: 0, bmiRec: 0, bmiScore: 0, bmiGrade: 0 },
      "2": { count: 0, grade1: 0, grade2: 0, grade3: 0, grade4: 0, grade5: 0, totalScore: 0, totalGrade: 0, cardioRec: 0, cardioScore: 0, cardioGrade: 0, flexRec: 0, flexScore: 0, flexGrade: 0, strengthRec: 0, strengthScore: 0, strengthGrade: 0, powerRec: 0, powerScore: 0, powerGrade: 0, bmiRec: 0, bmiScore: 0, bmiGrade: 0 },
      "3": { count: 0, grade1: 0, grade2: 0, grade3: 0, grade4: 0, grade5: 0, totalScore: 0, totalGrade: 0, cardioRec: 0, cardioScore: 0, cardioGrade: 0, flexRec: 0, flexScore: 0, flexGrade: 0, strengthRec: 0, strengthScore: 0, strengthGrade: 0, powerRec: 0, powerScore: 0, powerGrade: 0, bmiRec: 0, bmiScore: 0, bmiGrade: 0 }
    };

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

// 학년이 올라가면 학번이 바뀌므로(예: 1학년 10101 -> 2학년 20101), 학번이 아니라
// 이름으로 전년도 PAPS_상세 시트를 찾아 종합점수 증감을 계산한다. 동명이인이 있어
// 특정할 수 없거나 전년도 데이터 자체가 없으면 비교하지 않고 그 상태를 알려준다.
function buildPapsGrowthTrack_(name, year, currentTotalScore, currentTotalGrade) {
  const prevYear = (parseInt(year, 10) - 1).toString();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prevSheet = ss.getSheetByName('PAPS_상세_' + prevYear);
  if (!prevSheet) return { status: 'no_data' };
  const data = prevSheet.getDataRange().getValues();
  const matches = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === name) matches.push(data[i]);
  }
  if (matches.length === 0) return { status: 'no_data' };
  if (matches.length > 1) return { status: 'ambiguous' };
  const prevTotalScore = parseFloat(matches[0][2]);
  if (isNaN(prevTotalScore) || isNaN(currentTotalScore)) return { status: 'no_data' };
  const delta = Math.round((currentTotalScore - prevTotalScore) * 10) / 10;
  return {
    status: 'ok', prevYear: prevYear, prevTotalScore: prevTotalScore, prevTotalGrade: matches[0][3],
    currentTotalScore: currentTotalScore, currentTotalGrade: currentTotalGrade,
    delta: delta, improved: delta > 0
  };
}

// studentId는 반드시 본인(학생) 또는 교사만 조회 가능하도록 검증한다.
function getStudentPapsDetail(studentId, year, token) {
  try {
    requireOwnerOrTeacher(token, studentId);
    if (!year) year = getCurrentAcademicYear();
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
      },
      growthTrack: buildPapsGrowthTrack_(targetRow[1].toString().trim(), year, parseFloat(targetRow[2]), targetRow[3])
    };

    const FOLDER_ID = "1m31h6WC-EGuVHjaHi1ecBKdSDRBvOAvP"; const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFilesByName(studentId + ".pdf"); let pdfUrl = null;
    if (files.hasNext()) { const file = files.next(); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); pdfUrl = file.getUrl().replace("view?usp=drivesdk", "preview"); }
    return { success: true, papsData: papsData, pdfUrl: pdfUrl };
  } catch (e) { return { success: false, message: "오류 발생: " + e.message }; }
}

function saveNeisPapsData(parsedData, year, token) {
  try {
    requirePeTeacher(token);
    if (!year) year = getCurrentAcademicYear();
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

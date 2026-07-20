// ==========================================================
// 소규모 수업 - 체대입시반 로직
// 명단(체대입시반_명단) / 종목(체대입시반_종목) /
// 공식기록(체대입시반_공식기록, 3·6·9월 교사 입력) /
// 자율기록(체대입시반_자율기록, 학생 개인 연습 기록) /
// 배점표(체대입시반_배점표, 종목별 기록구간→점수 환산표) 5개 시트로 구성.
// 랭킹·최고기록은 공식기록만 집계 대상으로 삼는다(자율기록은 본인만 열람).
// 종목마다 만점을 설정하고(합이 1000점이 되도록), 배점표로 기록을 점수로
// 환산해 회차별 총점을 계산한다.
// ==========================================================
function setupPhysPrepSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    '체대입시반_명단': ['학번', '이름', '학년', '등록일'],
    '체대입시반_종목': ['종목ID', '종목명', '단위', '방향', '사용여부', '등록일', '만점'],
    '체대입시반_공식기록': ['학번', '이름', '학년', '연도', '회차', '종목ID', '기록값', '입력일시', '입력교사', '환산점수', '만점_스냅샷'],
    '체대입시반_자율기록': ['학번', '이름', '종목ID', '기록값', '입력일시', '측정회차', '기록ID']
  };
  Object.keys(sheets).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      const headers = sheets[name];
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#6a1b9a").setFontColor("white").setFontWeight("bold");
    } else {
      // 기존에 이미 만들어져 있던 시트에는 새로 추가된 컬럼이 없을 수 있으므로 보정한다.
      const headers = sheets[name];
      if (sheet.getLastColumn() < headers.length || sheet.getRange(1, headers.length).getValue() !== headers[headers.length - 1]) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setBackground("#6a1b9a").setFontColor("white").setFontWeight("bold");
      }
    }
  });

  // 자율기록에 기록ID가 없는(위 마이그레이션으로 컬럼만 막 추가된) 예전 데이터가 있으면
  // 학생이 개별 삭제할 수 있도록 UUID를 하나씩 채워준다. 새로 추가되는 기록은 항상
  // 기록ID를 갖고 생성되므로(addSelfRecord), 이 보정은 딱 한 번만 하면 충분하다.
  // 스크립트 속성으로 가드하지 않으면 전체 자율기록을 매 페이지 로드(doGet)마다
  // 통째로 읽어 훑는 꼴이 되어 기록이 쌓일수록 느려지므로, 완료 후 다시 실행하지 않는다.
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PHYSPREP_SELF_RECORD_ID_MIGRATED')) {
    const selfSheet = ss.getSheetByName('체대입시반_자율기록');
    if (selfSheet) {
      const selfData = selfSheet.getDataRange().getValues();
      for (let i = 1; i < selfData.length; i++) {
        if (selfData[i][0] && !selfData[i][6]) selfSheet.getRange(i + 1, 7).setValue(Utilities.getUuid());
      }
    }
    props.setProperty('PHYSPREP_SELF_RECORD_ID_MIGRATED', '1');
  }

  // 배점표는 "기준값 하나"가 아니라 "하한(이상)~상한(미만) 구간"으로 관리한다.
  // 예전 단일 기준값 버전으로 이미 만들어져 있으면, 방향(HIGH는 하한, LOW는 상한)에
  // 맞춰 구간으로 한 번만 변환해준다.
  const scoreHeaders = ['배점ID', '종목ID', '하한값(이상)', '상한값(미만)', '점수'];
  let scoreSheet = ss.getSheetByName('체대입시반_배점표');
  if (!scoreSheet) {
    scoreSheet = ss.insertSheet('체대입시반_배점표');
    scoreSheet.getRange(1, 1, 1, scoreHeaders.length).setValues([scoreHeaders]);
    scoreSheet.getRange(1, 1, 1, scoreHeaders.length).setBackground("#6a1b9a").setFontColor("white").setFontWeight("bold");
  } else if (scoreSheet.getRange(1, 3).getValue() === '기준값') {
    const data = scoreSheet.getDataRange().getValues();
    const eventsData = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
    const directionByEvent = {};
    for (let i = 1; i < eventsData.length; i++) { if (eventsData[i][0]) directionByEvent[eventsData[i][0]] = eventsData[i][3]; }
    const migrated = [scoreHeaders];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const dir = directionByEvent[data[i][1]] || 'HIGH';
      if (dir === 'LOW') migrated.push([data[i][0], data[i][1], '', data[i][2], data[i][3]]);
      else migrated.push([data[i][0], data[i][1], data[i][2], '', data[i][3]]);
    }
    scoreSheet.clear();
    scoreSheet.getRange(1, 1, migrated.length, scoreHeaders.length).setValues(migrated);
    scoreSheet.getRange(1, 1, 1, scoreHeaders.length).setBackground("#6a1b9a").setFontColor("white").setFontWeight("bold");
  }
}

// 학번 여러 개를 한 번에 조회할 때 쓴다. 반 대표 명단(주전+후보) 제출처럼 한 요청 안에서
// 학번을 여러 명 조회해야 할 때, lookupStudentInfo를 인원수만큼 반복 호출하면 전교생이
// 담긴 학생명렬표를 그 횟수만큼 통째로 다시 읽게 되어 느려진다. 시트는 한 번만 읽어서
// 조회 대상 학번들의 결과를 맵으로 한꺼번에 돌려준다.
function lookupStudentInfoBulk_(ids) {
  const wanted = {};
  ids.forEach(id => { wanted[id.toString().trim()] = true; });
  const result = {};
  if (Object.keys(wanted).length === 0) return result;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('학생명렬표');
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues();
  for (let i = 20; i < data.length; i++) {
    const idStr = data[i][0] ? data[i][0].toString().trim() : '';
    if (idStr && wanted[idStr]) result[idStr] = { name: data[i][1].toString().trim(), grade: idStr.charAt(0) };
  }
  return result;
}

// 학번으로 학생명렬표에서 이름/학년 조회 (명단 등록 시 사용)
function lookupStudentInfo(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('학생명렬표');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const idStr = studentId.toString().trim();
  for (let i = 20; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === idStr) {
      return { name: data[i][1].toString().trim(), grade: idStr.charAt(0) };
    }
  }
  return null;
}

function isPhysPrepMember(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('체대입시반_명단');
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const idStr = studentId.toString().trim();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === idStr) return true;
  }
  return false;
}

// 체육교사이거나, 명단에 등록된 학생 본인만 허용
function requirePhysPrepAccess(token) {
  const session = requireSession(token);
  const isTeacher = (session.role === '체육교사');
  const isMember = (session.role === '학생' && isPhysPrepMember(session.id));
  if (!isTeacher && !isMember) throw new Error("체대입시반 전용 수업입니다. 이용 권한이 없습니다.");
  return session;
}

// ==========================================================
// 명단 관리 (체육교사 전용)
// ==========================================================
function addPhysPrepStudent(studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_명단');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === idStr) return { success: false, message: "이미 등록된 학생입니다." };
    }
    const info = lookupStudentInfo(idStr);
    if (!info) return { success: false, message: "명렬표에서 학번을 찾을 수 없습니다." };
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([idStr, info.name, info.grade, dateStr]);
    return { success: true, message: info.name + " 학생을 명단에 등록했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removePhysPrepStudent(studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_명단');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === idStr) { sheet.deleteRow(i + 1); return { success: true, message: "명단에서 삭제했습니다." }; }
    }
    return { success: false, message: "명단에서 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getPhysPrepRoster(token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = ss.getSheetByName('체대입시반_명단').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) list.push({ id: data[i][0].toString(), name: data[i][1], grade: data[i][2], regDate: toDateStr_(data[i][3]) });
    }
    // 학번이 "학년+반(2자리)+번호(2자리)"로 구성되어 있어, 문자열 그대로 정렬하면
    // 학년 -> 반 -> 번호 순서가 자연스럽게 맞춰진다.
    list.sort((a, b) => a.id.localeCompare(b.id));
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// ==========================================================
// 종목 관리 (체육교사 전용 등록/수정, 조회는 로그인만 하면 가능)
// ==========================================================
function addPhysPrepEvent(name, unit, direction, maxScore, token) {
  try {
    requirePeTeacher(token);
    name = name.toString().trim();
    if (!name) return { success: false, message: "종목명을 입력하세요." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_종목');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === name && data[i][4] === 'Y') return { success: false, message: "이미 등록된 종목입니다." };
    }
    const eventId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([eventId, name, (unit || '').toString().trim(), direction === 'LOW' ? 'LOW' : 'HIGH', 'Y', dateStr, parseInt(maxScore, 10) || 0]);
    return { success: true, message: "종목이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updatePhysPrepEvent(eventId, name, unit, direction, maxScore, active, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_종목');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[name, unit, direction === 'LOW' ? 'LOW' : 'HIGH', active ? 'Y' : 'N']]);
        sheet.getRange(i + 1, 7).setValue(parseInt(maxScore, 10) || 0);
        return { success: true, message: "종목 정보를 수정했습니다." };
      }
    }
    return { success: false, message: "종목을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 종목명은 예시일 뿐 실제로는 자유롭게 추가되는 목록이라 활성/비활성만 구분해 반환한다.
// 화면 표시는 종목명 글자수가 긴 것부터, 같은 글자수면 가나다순으로 배열한다.
function getPhysPrepEvents(token) {
  try {
    requireSession(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) list.push({ id: data[i][0], name: data[i][1], unit: data[i][2], direction: data[i][3], active: data[i][4] === 'Y', maxScore: parseInt(data[i][6], 10) || 0 });
    }
    list.sort((a, b) => b.name.length - a.name.length || a.name.localeCompare(b.name, 'ko'));
    return { success: true, list: list };
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// 종목을 완전히 삭제한다(배점표도 같이 정리). 이미 입력된 측정 기록(공식/자율)은 그대로
// 남고, 종목명 조회 시 "(삭제된 종목)"으로 표시된다.
function removePhysPrepEvent(eventId, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_종목');
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) { if (data[i][0] === eventId) { sheet.deleteRow(i + 1); found = true; break; } }
    if (!found) return { success: false, message: "종목을 찾을 수 없습니다." };
    const scoreSheet = ss.getSheetByName('체대입시반_배점표');
    const scoreData = scoreSheet.getDataRange().getValues();
    for (let i = scoreData.length - 1; i >= 1; i--) { if (scoreData[i][1] === eventId) scoreSheet.deleteRow(i + 1); }
    return { success: true, message: "종목을 삭제했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getPhysPrepEventMeta(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) return { id: data[i][0], name: data[i][1], unit: data[i][2], direction: data[i][3], maxScore: parseInt(data[i][6], 10) || 0 };
  }
  return null;
}

// ==========================================================
// 배점표 관리 (종목별 기록 구간[하한 이상, 상한 미만) → 점수 환산표, 체육교사 전용 설정)
// 종목마다 만점을 정해두고(전체 합이 1000점이 되도록), 배점표로 기록을
// 점수로 환산한다. 배점표가 없는 종목은 총점 계산에서 제외된다.
// 구간은 방향과 무관하게 항상 "하한 이상 ~ 상한 미만"이며, 최고 구간은 상한을
// 비워두고(예: 300 이상) 최저 구간은 하한을 비워둘 수 있다(예: 200 미만).
// ==========================================================
function getPhysPrepScoreTable(eventId, token) {
  try {
    requireSession(token);
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('체대입시반_배점표').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) { if (data[i][1] === eventId) list.push({ id: data[i][0], lower: data[i][2], upper: data[i][3], score: data[i][4] }); }
    list.sort((a, b) => (parseInt(b.score, 10) || 0) - (parseInt(a.score, 10) || 0));
    return { success: true, list: list };
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// 화면에서 행을 자유롭게 추가/수정/삭제한 뒤 한 번에 저장한다 - 이 종목의 기존 배점표
// 전체를 지우고 새로 받은 목록으로 다시 만든다(부분 수정이 아니라 통째로 교체).
function savePhysPrepScoreTable(eventId, rows, token) {
  try {
    requirePeTeacher(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('체대입시반_배점표');
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) { if (data[i][1] === eventId) sheet.deleteRow(i + 1); }
    (rows || []).forEach(r => {
      const hasLower = r.lower !== '' && r.lower !== null && r.lower !== undefined;
      const hasUpper = r.upper !== '' && r.upper !== null && r.upper !== undefined;
      if (!hasLower && !hasUpper) return; // 이상/미만이 둘 다 비어있는 행은 무시
      const s = parseInt(r.score, 10);
      if (isNaN(s)) return;
      sheet.appendRow([Utilities.getUuid(), eventId, hasLower ? parseFloat(r.lower) : '', hasUpper ? parseFloat(r.upper) : '', s]);
    });
    return { success: true, message: "배점표를 저장했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 기록값 하나를 종목의 배점표(구간 목록)에 따라 점수로 환산한다. 배점표가 없으면 null(집계 제외).
// 구간은 [하한 이상, 상한 미만)이며 방향과 무관하다 — 어느 구간에 높은 점수를 매길지는
// 배점표를 만드는 선생님이 정한다.
function physPrepCalcScore_(ranges, value) {
  if (!ranges || ranges.length === 0) return null;
  let best = null;
  for (let i = 0; i < ranges.length; i++) {
    const r = ranges[i];
    const lowOk = r.lower === '' || r.lower === null || r.lower === undefined || value >= parseFloat(r.lower);
    const highOk = r.upper === '' || r.upper === null || r.upper === undefined || value < parseFloat(r.upper);
    if (lowOk && highOk) { const sc = parseInt(r.score, 10) || 0; if (best === null || sc > best) best = sc; }
  }
  return best === null ? 0 : best; // 어떤 구간에도 해당하지 않으면 0점
}

// 학생 본인의 회차(연도·3/6/9월)별 총점 추이. 배점표가 설정된 종목만 집계 대상이다.
// 각 기록은 저장 시점(saveOfficialRecords)에 계산해둔 환산점수/만점 스냅샷을 그대로
// 사용하므로, 이후 배점표를 수정하더라도 과거 회차에 이미 표시됐던 총점은 바뀌지
// 않는다(회차별로 "OOO/만점점" 형태로 반환). 이 기능 도입 이전에 저장돼 스냅샷이
// 없는 옛 기록만 예외적으로 현재 배점표를 이용해 한 번 보정 계산한다.
function getMyPhysPrepScoreSummary(token) {
  try {
    const session = requirePhysPrepAccess(token);
    if (session.role !== '학생') return { success: false, message: "학생 계정으로 로그인해야 합니다." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const eventsData = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
    const events = {};
    for (let i = 1; i < eventsData.length; i++) {
      if (eventsData[i][0] && eventsData[i][4] === 'Y') events[eventsData[i][0]] = { direction: eventsData[i][3], maxScore: parseInt(eventsData[i][6], 10) || 0 };
    }

    const tableData = ss.getSheetByName('체대입시반_배점표').getDataRange().getValues();
    const tablesByEvent = {};
    for (let i = 1; i < tableData.length; i++) {
      const eid = tableData[i][1];
      if (!events[eid]) continue;
      if (!tablesByEvent[eid]) tablesByEvent[eid] = [];
      tablesByEvent[eid].push({ lower: tableData[i][2], upper: tableData[i][3], score: tableData[i][4] });
    }

    const officialData = ss.getSheetByName('체대입시반_공식기록').getDataRange().getValues();
    const byRound = {};
    for (let i = 1; i < officialData.length; i++) {
      const row = officialData[i];
      if (!row[0] || row[0].toString().trim() !== session.id.toString()) continue;
      const eid = row[5];
      const storedScore = row[9], storedMax = row[10];
      let score, maxForEvent;
      if (storedScore !== '' && storedScore !== null && storedScore !== undefined) {
        score = parseInt(storedScore, 10) || 0;
        maxForEvent = parseInt(storedMax, 10) || 0;
      } else {
        if (!tablesByEvent[eid]) continue;
        const calc = physPrepCalcScore_(tablesByEvent[eid], parseFloat(row[6]));
        if (calc === null) continue;
        score = calc;
        maxForEvent = events[eid] ? events[eid].maxScore : 0;
      }
      const key = row[3] + '-' + row[4];
      if (!byRound[key]) byRound[key] = { year: parseInt(row[3], 10), round: parseInt(row[4], 10), score: 0, maxTotal: 0 };
      byRound[key].score += score;
      byRound[key].maxTotal += maxForEvent;
    }

    const trend = Object.values(byRound).sort((a, b) => (a.year * 100 + a.round) - (b.year * 100 + b.round));
    return { success: true, trend: trend };
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 공식측정 기록 입력 (3·6·9월, 체육교사 전용 일괄 입력)
// ==========================================================
function getOfficialRecordsForEntry(year, round, eventId, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterData = ss.getSheetByName('체대입시반_명단').getDataRange().getValues();
    const roster = [];
    for (let i = 1; i < rosterData.length; i++) { if (rosterData[i][0]) roster.push({ id: rosterData[i][0].toString(), name: rosterData[i][1] }); }

    const recData = ss.getSheetByName('체대입시반_공식기록').getDataRange().getValues();
    const valueMap = {};
    for (let i = 1; i < recData.length; i++) {
      const row = recData[i];
      if (row[0] && row[3] == year && row[4] == round && row[5] === eventId) valueMap[row[0].toString()] = row[6];
    }
    const result = roster.map(s => ({ id: s.id, name: s.name, value: valueMap[s.id] !== undefined ? valueMap[s.id] : '' }));
    return { success: true, roster: result };
  } catch (e) { return { success: false, message: e.message, roster: [] }; }
}

function saveOfficialRecords(year, round, eventId, records, token) {
  try {
    const session = requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_공식기록');
    const data = sheet.getDataRange().getValues();
    const rowMap = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[3] == year && row[4] == round && row[5] === eventId) rowMap[row[0].toString()] = i + 1;
    }

    // 저장 시점의 배점표로 점수를 미리 계산해 환산점수/만점 스냅샷으로 함께 저장해둔다 -
    // 나중에 배점표를 고쳐도 이미 저장된 회차의 총점 표시가 바뀌지 않도록 하기 위함
    // (getMyPhysPrepScoreSummary 참고).
    const eventMeta = getPhysPrepEventMeta(eventId);
    const scoreTableData = ss.getSheetByName('체대입시반_배점표').getDataRange().getValues();
    const ranges = [];
    for (let i = 1; i < scoreTableData.length; i++) { if (scoreTableData[i][1] === eventId) ranges.push({ lower: scoreTableData[i][2], upper: scoreTableData[i][3], score: scoreTableData[i][4] }); }

    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    let savedCount = 0;
    records.forEach(r => {
      if (r.value === '' || r.value === null || r.value === undefined) return; // 빈 값은 건너뜀 (미측정)
      const grade = r.id.toString().charAt(0);
      const calcScore = ranges.length > 0 ? physPrepCalcScore_(ranges, parseFloat(r.value)) : null;
      const scoreSnapshot = calcScore === null ? '' : calcScore;
      const maxSnapshot = (calcScore === null || !eventMeta) ? '' : eventMeta.maxScore;
      const rowData = [r.id, r.name, grade, year, round, eventId, r.value, dateStr, session.name, scoreSnapshot, maxSnapshot];
      if (rowMap[r.id]) { sheet.getRange(rowMap[r.id], 1, 1, 11).setValues([rowData]); } else { sheet.appendRow(rowData); }
      savedCount++;
    });
    return { success: true, message: `${savedCount}명의 기록을 저장했습니다.` };
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 자율기록 (학생 개인 연습 기록, 랭킹 미반영)
// 종목당 하루 최대 2회까지 기록할 수 있고, 1회만 측정하는 종목도 있으므로
// 1차 기록만 입력해도 등록되도록 2차는 선택 입력으로 둔다.
// ==========================================================
function addSelfRecord(eventId, value1, value2, token) {
  try {
    const session = requirePhysPrepAccess(token);
    if (session.role !== '학생') return { success: false, message: "학생만 자율 기록을 입력할 수 있습니다." };
    if (value1 === '' || value1 === null || value1 === undefined || isNaN(parseFloat(value1))) return { success: false, message: "1회 기록값을 정확히 입력하세요." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_자율기록');
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([session.id, session.name, eventId, value1, dateStr, 1, Utilities.getUuid()]);
    if (value2 !== '' && value2 !== null && value2 !== undefined && !isNaN(parseFloat(value2))) {
      sheet.appendRow([session.id, session.name, eventId, value2, dateStr, 2, Utilities.getUuid()]);
    }
    return { success: true, message: "기록이 저장되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 잘못 입력한 연습 기록을 본인이 직접 지울 수 있게 한다.
function removeSelfRecord(recordId, token) {
  try {
    const session = requirePhysPrepAccess(token);
    if (session.role !== '학생') return { success: false, message: "학생만 자율 기록을 삭제할 수 있습니다." };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('체대입시반_자율기록');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === recordId) {
        if (data[i][0].toString().trim() !== session.id.toString()) return { success: false, message: "본인의 기록만 삭제할 수 있습니다." };
        sheet.deleteRow(i + 1);
        return { success: true, message: "기록을 삭제했습니다." };
      }
    }
    return { success: false, message: "기록을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 본인의 공식기록(연도/회차 순 정렬 + 직전 대비 향상량) 및 자율기록 이력 조회
function getMyPhysPrepRecords(token) {
  try {
    const session = requirePhysPrepAccess(token);
    if (session.role !== '학생') return { success: false, message: "학생 계정으로 로그인해야 개인 기록을 볼 수 있습니다." };
    const studentId = session.id;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const eventsData = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
    const eventMeta = {};
    for (let i = 1; i < eventsData.length; i++) { if (eventsData[i][0]) eventMeta[eventsData[i][0]] = { name: eventsData[i][1], unit: eventsData[i][2], direction: eventsData[i][3] }; }

    const officialData = ss.getSheetByName('체대입시반_공식기록').getDataRange().getValues();
    const byEvent = {};
    for (let i = 1; i < officialData.length; i++) {
      const row = officialData[i];
      if (row[0] && row[0].toString().trim() === studentId) {
        const eid = row[5];
        if (!byEvent[eid]) byEvent[eid] = [];
        byEvent[eid].push({ year: parseInt(row[3], 10), round: parseInt(row[4], 10), value: parseFloat(row[6]) });
      }
    }
    const officialResult = {};
    Object.keys(byEvent).forEach(eid => {
      const list = byEvent[eid].sort((a, b) => (a.year * 100 + a.round) - (b.year * 100 + b.round));
      const meta = eventMeta[eid] || { name: '(삭제된 종목)', unit: '', direction: 'HIGH' };
      const withDelta = list.map((rec, idx) => {
        let delta = null, improved = null;
        if (idx > 0) {
          delta = Math.round((rec.value - list[idx - 1].value) * 100) / 100;
          improved = meta.direction === 'LOW' ? delta < 0 : delta > 0;
        }
        return { year: rec.year, round: rec.round, value: rec.value, delta: delta, improved: improved };
      });
      officialResult[eid] = { name: meta.name, unit: meta.unit, direction: meta.direction, records: withDelta };
    });

    // 자율기록은 종목별로 묶어서, 최고기록과 시간순 히스토리(추이 그래프용)를 함께 내려준다.
    const selfData = ss.getSheetByName('체대입시반_자율기록').getDataRange().getValues();
    const selfByEvent = {};
    for (let i = 1; i < selfData.length; i++) {
      const row = selfData[i];
      if (row[0] && row[0].toString().trim() === studentId) {
        const eid = row[2];
        if (!selfByEvent[eid]) selfByEvent[eid] = [];
        selfByEvent[eid].push({ id: row[6], value: parseFloat(row[3]), date: toDateTimeStr_(row[4]), round: row[5] || 1 });
      }
    }
    const selfResult = {};
    Object.keys(selfByEvent).forEach(eid => {
      const meta = eventMeta[eid] || { name: '(삭제된 종목)', unit: '', direction: 'HIGH' };
      const list = selfByEvent[eid].sort((a, b) => new Date(a.date) - new Date(b.date));
      let best = null;
      list.forEach(r => { if (best === null || (meta.direction === 'LOW' ? r.value < best : r.value > best)) best = r.value; });
      selfResult[eid] = { name: meta.name, unit: meta.unit, direction: meta.direction, best: best, records: list };
    });

    return sanitizeDates_({ success: true, official: officialResult, self: selfResult });
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 랭킹 (공식기록 전용). mode: 'all' | 'year:YYYY' | 'window:3' | 'window:5'
// 동일 학생은 해당 구간 내 본인 최고기록 1건만 집계한다.
// ==========================================================
function getPhysPrepRanking(eventId, mode, token) {
  try {
    requirePhysPrepAccess(token);
    const eventMeta = getPhysPrepEventMeta(eventId);
    if (!eventMeta) return { success: false, message: "종목을 찾을 수 없습니다." };

    let minYear = -Infinity, maxYear = Infinity;
    if (mode.indexOf('window:') === 0) {
      const n = parseInt(mode.split(':')[1], 10);
      minYear = parseInt(getCurrentAcademicYear(), 10) - n + 1;
    } else if (mode.indexOf('year:') === 0) {
      minYear = maxYear = parseInt(mode.split(':')[1], 10);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = ss.getSheetByName('체대입시반_공식기록').getDataRange().getValues();
    const bestByStudent = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || row[5] !== eventId) continue;
      const year = parseInt(row[3], 10);
      if (year < minYear || year > maxYear) continue;
      const value = parseFloat(row[6]);
      if (isNaN(value)) continue;
      const id = row[0].toString();
      const cur = bestByStudent[id];
      const better = !cur || (eventMeta.direction === 'LOW' ? value < cur.value : value > cur.value);
      // 이름·학년은 그 회차 당시 명렬표에서 조회해온 값을 기록 시점에 그대로 저장해둔 것을 쓴다.
      // 그래야 나중에 명렬표에서 학생이 사라지거나(졸업 등) 학번이 재사용되어도 과거 기록의
      // 이름·학년 표기가 절대 바뀌거나 유실되지 않는다.
      if (better) bestByStudent[id] = { id: id, name: row[1], grade: row[2], value: value, year: year, round: row[4] };
    }
    // 반평생(30년+) 쌓일 데이터라 랭킹 화면에는 상위 10명까지만 보여준다. 나머지 기록은
    // 삭제하지 않고 시트에 그대로 남아있으며, 조회 조건(연도/기간)을 바꾸면 다시 나타난다.
    const ranking = Object.values(bestByStudent).sort((a, b) => eventMeta.direction === 'LOW' ? a.value - b.value : b.value - a.value).slice(0, 10);
    return { success: true, event: eventMeta, ranking: ranking };
  } catch (e) { return { success: false, message: e.message }; }
}

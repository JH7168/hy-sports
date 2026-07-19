// ==========================================================
// 소규모 수업 - 체대입시반 로직
// 명단(체대입시반_명단) / 종목(체대입시반_종목) /
// 공식기록(체대입시반_공식기록, 3·6·9월 교사 입력) /
// 자율기록(체대입시반_자율기록, 학생 개인 연습 기록) 4개 시트로 구성.
// 랭킹·최고기록은 공식기록만 집계 대상으로 삼는다(자율기록은 본인만 열람).
// ==========================================================
function setupPhysPrepSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    '체대입시반_명단': ['학번', '이름', '학년', '등록일'],
    '체대입시반_종목': ['종목ID', '종목명', '단위', '방향', '사용여부', '등록일'],
    '체대입시반_공식기록': ['학번', '이름', '학년', '연도', '회차', '종목ID', '기록값', '입력일시', '입력교사'],
    '체대입시반_자율기록': ['학번', '이름', '종목ID', '기록값', '입력일시']
  };
  Object.keys(sheets).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      const headers = sheets[name];
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#6a1b9a").setFontColor("white").setFontWeight("bold");
    }
  });
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
      if (data[i][0]) list.push({ id: data[i][0].toString(), name: data[i][1], grade: data[i][2], regDate: data[i][3] });
    }
    return { success: true, list: list };
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// ==========================================================
// 종목 관리 (체육교사 전용 등록/수정, 조회는 로그인만 하면 가능)
// ==========================================================
function addPhysPrepEvent(name, unit, direction, token) {
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
    sheet.appendRow([eventId, name, (unit || '').toString().trim(), direction === 'LOW' ? 'LOW' : 'HIGH', 'Y', dateStr]);
    return { success: true, message: "종목이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updatePhysPrepEvent(eventId, name, unit, direction, active, token) {
  try {
    requirePeTeacher(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_종목');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[name, unit, direction === 'LOW' ? 'LOW' : 'HIGH', active ? 'Y' : 'N']]);
        return { success: true, message: "종목 정보를 수정했습니다." };
      }
    }
    return { success: false, message: "종목을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 종목명은 예시일 뿐 실제로는 자유롭게 추가되는 목록이라 활성/비활성만 구분해 반환한다.
function getPhysPrepEvents(token) {
  try {
    requireSession(token);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) list.push({ id: data[i][0], name: data[i][1], unit: data[i][2], direction: data[i][3], active: data[i][4] === 'Y' });
    }
    return { success: true, list: list };
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

function getPhysPrepEventMeta(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName('체대입시반_종목').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) return { id: data[i][0], name: data[i][1], unit: data[i][2], direction: data[i][3] };
  }
  return null;
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
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    let savedCount = 0;
    records.forEach(r => {
      if (r.value === '' || r.value === null || r.value === undefined) return; // 빈 값은 건너뜀 (미측정)
      const grade = r.id.toString().charAt(0);
      const rowData = [r.id, r.name, grade, year, round, eventId, r.value, dateStr, session.name];
      if (rowMap[r.id]) { sheet.getRange(rowMap[r.id], 1, 1, 9).setValues([rowData]); } else { sheet.appendRow(rowData); }
      savedCount++;
    });
    return { success: true, message: `${savedCount}명의 기록을 저장했습니다.` };
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 자율기록 (학생 개인 연습 기록, 랭킹 미반영)
// ==========================================================
function addSelfRecord(eventId, value, token) {
  try {
    const session = requirePhysPrepAccess(token);
    if (session.role !== '학생') return { success: false, message: "학생만 자율 기록을 입력할 수 있습니다." };
    if (value === '' || value === null || value === undefined || isNaN(parseFloat(value))) return { success: false, message: "기록값을 정확히 입력하세요." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('체대입시반_자율기록');
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([session.id, session.name, eventId, value, dateStr]);
    return { success: true, message: "기록이 저장되었습니다." };
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

    const selfData = ss.getSheetByName('체대입시반_자율기록').getDataRange().getValues();
    const selfList = [];
    for (let i = 1; i < selfData.length; i++) {
      const row = selfData[i];
      if (row[0] && row[0].toString().trim() === studentId) {
        const meta = eventMeta[row[2]] || { name: '(삭제된 종목)', unit: '' };
        selfList.push({ eventName: meta.name, unit: meta.unit, value: row[3], date: row[4] });
      }
    }
    selfList.reverse();

    return { success: true, official: officialResult, self: selfList };
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
      if (better) bestByStudent[id] = { id: id, name: row[1], value: value, year: year, round: row[4] };
    }
    const ranking = Object.values(bestByStudent).sort((a, b) => eventMeta.direction === 'LOW' ? a.value - b.value : b.value - a.value);
    return { success: true, event: eventMeta, ranking: ranking };
  } catch (e) { return { success: false, message: e.message }; }
}

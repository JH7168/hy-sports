// ==========================================================
// 점심시간 시설 예약 시스템 (풋살장 / 체육관 / 운동장 공용)
// 세 시설이 동일한 로직을 공유한다: 체육교사가 이용일/모드(학년매칭·
// 자유경쟁)/신청기간/발표시각/선발팀수를 정해 슬롯을 열면, 팀장이
// 팀명+초대 멤버(학번 또는 교사이름)를 등록하고, 초대된 멤버 전원이
// 각자 로그인해서 "가입 확정"해야만 그 팀이 추첨 대상이 된다.
// 신청마감(appEnd) 이후에는 더 이상 팀 생성/가입확정이 불가능하고,
// 발표시각(announceAt)이 지난 뒤 누군가 슬롯 정보를 처음 조회하는
// 시점에 전원 확정된 팀들 중에서 무작위로 선발팀이 정해진다(지연 추첨).
//
// 시설별 데이터는 시트명 접두어(풋살장_/체육관_/운동장_)로 물리적으로
// 분리해 스프레드시트에서 보기 쉽게 하되, 코드 로직은 facility
// 파라미터 하나로 공유한다.
// ==========================================================
const FACILITY_LABELS = { FUTSAL: '풋살장', GYM: '체육관', FIELD: '운동장' };

function facilitySheetName_(facility, suffix) {
  const label = FACILITY_LABELS[facility];
  if (!label) throw new Error('알 수 없는 시설입니다: ' + facility);
  return label + '_' + suffix;
}

function setupFacilitySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(FACILITY_LABELS).forEach(facility => {
    const sheets = {
      [facilitySheetName_(facility, '예약슬롯')]: ['슬롯ID', '이용일자', '모드', '대상학년', '신청시작', '신청마감', '발표시각', '선발팀수', '상태', '추첨결과JSON', '생성일시', '생성교사'],
      [facilitySheetName_(facility, '팀신청')]: ['신청ID', '슬롯ID', '팀명', '주장ID', '주장이름', '주장역할', '신청일시', '선정여부'],
      [facilitySheetName_(facility, '팀원')]: ['신청ID', '회원ID', '이름', '역할', '가입확정', '확정일시'],
      [facilitySheetName_(facility, '관리학생')]: ['학번', '이름', '학년', '등록일'],
      [facilitySheetName_(facility, '위반기록')]: ['기록ID', '학번', '이름', '사유', '위반일자', '등록일시', '등록자ID', '등록자이름']
    };
    Object.keys(sheets).forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        const headers = sheets[name];
        sheet = ss.insertSheet(name);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setBackground("#00838f").setFontColor("white").setFontWeight("bold");
        if (name.indexOf('_예약슬롯') > -1) {
          sheet.getRange("B:B").setNumberFormat("@");
          sheet.getRange("E:G").setNumberFormat("@");
        }
      }
    });
  });

  // 체육관 요일별 스포츠클럽 준비일 기본 배정 (체육교사가 이후 수정 가능)
  let gymSchedule = ss.getSheetByName('체육관_요일배정');
  if (!gymSchedule) {
    gymSchedule = ss.insertSheet('체육관_요일배정');
    gymSchedule.appendRow(['요일', '배정부서']);
    gymSchedule.getRange("A1:B1").setBackground("#8d6e63").setFontColor("white").setFontWeight("bold");
    // 토·일은 등교하지 않으므로 배정 대상에서 제외한다.
    const defaults = [['월', '배구부'], ['화', '농구부'], ['수', '배드민턴부'], ['목', '배구부'], ['금', '농구부']];
    gymSchedule.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }
}

function getFacilitySlotSheet_(facility) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(facilitySheetName_(facility, '예약슬롯')); }
function getFacilityAppSheet_(facility) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(facilitySheetName_(facility, '팀신청')); }
function getFacilityMemberSheet_(facility) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(facilitySheetName_(facility, '팀원')); }
function getFacilityManagerSheet_(facility) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(facilitySheetName_(facility, '관리학생')); }
function getFacilityViolationSheet_(facility) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(facilitySheetName_(facility, '위반기록')); }

function facilitySlotRowToObj_(row, rowIndex) {
  return {
    rowIndex: rowIndex, id: row[0], date: toDateStr_(row[1]), mode: row[2], grade: row[3],
    appStart: toDateTimeStr_(row[4]), appEnd: toDateTimeStr_(row[5]), announceAt: toDateTimeStr_(row[6]), teamCount: row[7], status: row[8],
    resultJson: row[9], createdAt: row[10], createdBy: row[11]
  };
}

function isFacilityTeamComplete_(applicationId, memberRows) {
  const members = memberRows.filter(r => r[0] === applicationId);
  if (members.length === 0) return false;
  return members.every(r => r[4] === true);
}

// 실제 추첨 로직(조건 확인 없이 무조건 수행). 전원 확정된 팀만 후보로 삼는다.
function performFacilityDraw_(facility, slotId, slotRowIndex, teamCount) {
  const appSheet = getFacilityAppSheet_(facility);
  const appData = appSheet.getDataRange().getValues();
  const memberData = getFacilityMemberSheet_(facility).getDataRange().getValues();
  const memberRows = memberData.slice(1);

  const applicants = [];
  for (let i = 1; i < appData.length; i++) {
    if (appData[i][1] === slotId) {
      appSheet.getRange(i + 1, 8).setValue(false);
      if (isFacilityTeamComplete_(appData[i][0], memberRows)) applicants.push({ rowIndex: i + 1, id: appData[i][0] });
    }
  }
  for (let i = applicants.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = applicants[i]; applicants[i] = applicants[j]; applicants[j] = tmp;
  }
  const n = parseInt(teamCount, 10) || 2;
  const selected = applicants.slice(0, n);
  selected.forEach(a => appSheet.getRange(a.rowIndex, 8).setValue(true));
  const selectedIds = selected.map(a => a.id);
  getFacilitySlotSheet_(facility).getRange(slotRowIndex, 9, 1, 2).setValues([['DRAWN', JSON.stringify(selectedIds)]]);
  return selectedIds;
}

function ensureFacilityDrawn_(facility, slotObj) {
  if (slotObj.status !== 'OPEN') return slotObj;
  if (new Date() < new Date(slotObj.announceAt)) return slotObj;
  const selectedIds = performFacilityDraw_(facility, slotObj.id, slotObj.rowIndex, slotObj.teamCount);
  slotObj.status = 'DRAWN';
  slotObj.resultJson = JSON.stringify(selectedIds);
  return slotObj;
}

function getAllFacilitySlots_(facility) {
  const data = getFacilitySlotSheet_(facility).getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push(ensureFacilityDrawn_(facility, facilitySlotRowToObj_(data[i], i + 1)));
  }
  return list;
}

// 학번(숫자형) 또는 교사 이름으로 회원을 찾는다.
function lookupFacilityMember_(idOrName) {
  const key = idOrName.toString().trim();
  if (!key) return null;
  if (/^\d{4,5}$/.test(key)) {
    const info = lookupStudentInfo(key);
    if (info) return { id: key, name: info.name, role: '학생', grade: info.grade };
    return null;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('회원정보');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] === '교사' || data[i][0] === '체육교사') && data[i][1] && data[i][1].toString().trim() === key) {
      return { id: key, name: key, role: data[i][0], grade: null };
    }
  }
  return null;
}

// ==========================================================
// 슬롯 관리 (체육교사 전용)
// ==========================================================
function createFacilitySlot(facility, date, mode, grade, appStart, appEnd, announceAt, teamCount, token) {
  try {
    const session = requirePeTeacher(token);
    if (!date || !appStart || !appEnd || !announceAt) return { success: false, message: "이용일자, 신청 시작/마감, 발표시각을 모두 입력하세요." };
    if (new Date(appStart) >= new Date(appEnd)) return { success: false, message: "신청 시작 시각이 마감 시각보다 빨라야 합니다." };
    if (new Date(appEnd) > new Date(announceAt)) return { success: false, message: "발표시각은 신청 마감 이후여야 합니다." };
    const isGrade = mode === 'GRADE';
    const sheet = getFacilitySlotSheet_(facility);
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([id, date, isGrade ? 'GRADE' : 'OPEN', isGrade ? grade : '', appStart, appEnd, announceAt, teamCount || 2, 'OPEN', '', dateStr, session.name]);
    return { success: true, message: "예약 슬롯이 생성되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateFacilitySlot(facility, slotId, date, mode, grade, appStart, appEnd, announceAt, teamCount, status, token) {
  try {
    requirePeTeacher(token);
    if (new Date(appStart) >= new Date(appEnd)) return { success: false, message: "신청 시작 시각이 마감 시각보다 빨라야 합니다." };
    if (new Date(appEnd) > new Date(announceAt)) return { success: false, message: "발표시각은 신청 마감 이후여야 합니다." };
    const isGrade = mode === 'GRADE';
    const sheet = getFacilitySlotSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slotId) {
        sheet.getRange(i + 1, 2, 1, 7).setValues([[date, isGrade ? 'GRADE' : 'OPEN', isGrade ? grade : '', appStart, appEnd, announceAt, teamCount || 2]]);
        if (status) sheet.getRange(i + 1, 9).setValue(status);
        return { success: true, message: "슬롯 정보를 수정했습니다." };
      }
    }
    return { success: false, message: "슬롯을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function cancelFacilitySlot(facility, slotId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getFacilitySlotSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slotId) { sheet.getRange(i + 1, 9).setValue('CANCELLED'); return { success: true, message: "슬롯을 취소했습니다." }; }
    }
    return { success: false, message: "슬롯을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function drawFacilitySlot(facility, slotId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getFacilitySlotSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === slotId) {
        performFacilityDraw_(facility, slotId, i + 1, data[i][7]);
        return { success: true, message: "추첨을 완료했습니다." };
      }
    }
    return { success: false, message: "슬롯을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function overrideFacilityResult(facility, slotId, selectedApplicationIds, token) {
  try {
    requirePeTeacher(token);
    const appSheet = getFacilityAppSheet_(facility);
    const appData = appSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < appData.length; i++) {
      if (appData[i][1] === slotId) { found = true; appSheet.getRange(i + 1, 8).setValue(selectedApplicationIds.indexOf(appData[i][0]) > -1); }
    }
    if (!found) return { success: false, message: "해당 슬롯에 신청 내역이 없습니다." };
    const slotSheet = getFacilitySlotSheet_(facility);
    const slotData = slotSheet.getDataRange().getValues();
    for (let i = 1; i < slotData.length; i++) {
      if (slotData[i][0] === slotId) {
        slotSheet.getRange(i + 1, 9, 1, 2).setValues([['DRAWN', JSON.stringify(selectedApplicationIds)]]);
        return { success: true, message: "발표 결과를 수동으로 지정했습니다." };
      }
    }
    return { success: false, message: "슬롯을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getFacilitySlotsAdmin(facility, token) {
  try {
    requirePeTeacher(token);
    const slots = getAllFacilitySlots_(facility);
    const appData = getFacilityAppSheet_(facility).getDataRange().getValues();
    const countMap = {};
    for (let i = 1; i < appData.length; i++) { const sid = appData[i][1]; countMap[sid] = (countMap[sid] || 0) + 1; }
    slots.forEach(s => { s.appCount = countMap[s.id] || 0; });
    slots.sort((a, b) => new Date(b.date) - new Date(a.date));
    return sanitizeDates_({ success: true, list: slots });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

function getFacilityApplicationsAdmin(facility, slotId, token) {
  try {
    requirePeTeacher(token);
    const appData = getFacilityAppSheet_(facility).getDataRange().getValues();
    const memberData = getFacilityMemberSheet_(facility).getDataRange().getValues();
    const list = [];
    for (let i = 1; i < appData.length; i++) {
      if (appData[i][1] === slotId) {
        const appId = appData[i][0];
        const members = [];
        for (let j = 1; j < memberData.length; j++) { if (memberData[j][0] === appId) members.push({ id: memberData[j][1], name: memberData[j][2], role: memberData[j][3], confirmed: memberData[j][4] === true }); }
        list.push({ id: appId, teamName: appData[i][2], leaderId: appData[i][3], leaderName: appData[i][4], leaderRole: appData[i][5], appliedAt: appData[i][6], selected: appData[i][7] === true, members: members });
      }
    }
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

function removeFacilityApplication(facility, applicationId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getFacilityAppSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId) { sheet.deleteRow(i + 1); break; }
    }
    const memberSheet = getFacilityMemberSheet_(facility);
    const memberData = memberSheet.getDataRange().getValues();
    for (let i = memberData.length - 1; i >= 1; i--) { if (memberData[i][0] === applicationId) memberSheet.deleteRow(i + 1); }
    return { success: true, message: "신청을 삭제했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 다듬어진 문구로 홈 공지사항에 우천취소 안내를 자동 등록한다.
function postFacilityRainNotice(facility, slotId, token) {
  try {
    requirePeTeacher(token);
    const data = getFacilitySlotSheet_(facility).getDataRange().getValues();
    let slot = null;
    for (let i = 1; i < data.length; i++) { if (data[i][0] === slotId) { slot = facilitySlotRowToObj_(data[i], i + 1); break; } }
    if (!slot) return { success: false, message: "슬롯을 찾을 수 없습니다." };
    const label = FACILITY_LABELS[facility];
    const title = `📢 [${label} 이용 안내] ${slot.date} 예약 관련 안내`;
    const content = `${slot.date} ${label} 예약·이용과 관련하여 안내드립니다.\n\n⚠️ 우천 등 기상 악화로 경기가 취소될 경우, 별도의 재신청 기회 없이 이번 회차 예약은 그대로 종료됩니다. 신청 및 참여에 참고해 주시기 바랍니다.`;
    return saveNotice(title, content, token);
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 학생/교사 이용 화면 - 팀 신청 및 개별 가입확정
// ==========================================================
function createFacilityTeam(facility, slotId, teamName, memberIdsText, token) {
  try {
    const session = requireSession(token);
    teamName = (teamName || '').toString().trim();
    if (!teamName) return { success: false, message: "팀명을 입력하세요." };

    const slotData = getFacilitySlotSheet_(facility).getDataRange().getValues();
    let slot = null;
    for (let i = 1; i < slotData.length; i++) { if (slotData[i][0] === slotId) { slot = facilitySlotRowToObj_(slotData[i], i + 1); break; } }
    if (!slot) return { success: false, message: "존재하지 않는 예약 슬롯입니다." };
    slot = ensureFacilityDrawn_(facility, slot);
    if (slot.status !== 'OPEN') return { success: false, message: "이미 신청이 마감되었습니다." };
    const now = new Date();
    if (now < new Date(slot.appStart) || now > new Date(slot.appEnd)) return { success: false, message: "지금은 신청 가능한 시간이 아닙니다." };

    if (slot.mode === 'GRADE') {
      if (session.role !== '학생' || session.id.toString().charAt(0) !== slot.grade.toString()) {
        return { success: false, message: `이 예약은 ${slot.grade}학년 학생만 팀을 만들 수 있습니다.` };
      }
    }

    const lines = (memberIdsText || '').toString().split(/\n|,/).map(s => s.trim()).filter(s => s);
    const invited = [];
    const seen = { [session.id]: true };
    for (let i = 0; i < lines.length; i++) {
      const key = lines[i];
      if (seen[key]) continue;
      const info = lookupFacilityMember_(key);
      if (!info) return { success: false, message: `"${key}"님을 명단(학생명렬표/회원정보)에서 찾을 수 없습니다. 학생은 학번(예: 10101), 교사는 정확한 이름으로 입력해주세요.` };
      if (slot.mode === 'GRADE') {
        if (info.role !== '학생' || info.grade !== slot.grade.toString()) return { success: false, message: `"${key}"님은 ${slot.grade}학년 학생이 아니라 이 팀에 초대할 수 없습니다.` };
      }
      seen[key] = true;
      invited.push(info);
    }
    if (invited.length === 0) return { success: false, message: "팀장 본인 외에 최소 1명 이상 팀원을 초대해야 합니다." };

    const allIds = [session.id].concat(invited.map(m => m.id));
    const appData = getFacilityAppSheet_(facility).getDataRange().getValues();
    const slotAppIds = {};
    for (let i = 1; i < appData.length; i++) { if (appData[i][1] === slotId) slotAppIds[appData[i][0]] = true; }
    const memberData = getFacilityMemberSheet_(facility).getDataRange().getValues();
    for (let i = 1; i < memberData.length; i++) {
      if (slotAppIds[memberData[i][0]] && allIds.indexOf(memberData[i][1].toString()) > -1) {
        return { success: false, message: `${memberData[i][2]}님은 이미 이 슬롯의 다른 팀에 참여 중입니다.` };
      }
    }

    const appId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    getFacilityAppSheet_(facility).appendRow([appId, slotId, teamName, session.id, session.name, session.role, dateStr, false]);

    const memberSheet = getFacilityMemberSheet_(facility);
    memberSheet.appendRow([appId, session.id, session.name, session.role, true, dateStr]);
    invited.forEach(m => memberSheet.appendRow([appId, m.id, m.name, m.role, false, '']));

    return { success: true, message: `팀 신청이 등록되었습니다. 초대한 ${invited.length}명 전원이 각자 로그인해서 가입 확정해야 팀이 완성됩니다.` };
  } catch (e) { return { success: false, message: e.message }; }
}

function confirmFacilityMembership(facility, applicationId, token) {
  try {
    const session = requireSession(token);
    const memberSheet = getFacilityMemberSheet_(facility);
    const data = memberSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === applicationId && data[i][1].toString() === session.id.toString()) {
        if (data[i][4] === true) return { success: false, message: "이미 가입 확정했습니다." };
        const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        memberSheet.getRange(i + 1, 5, 1, 2).setValues([[true, dateStr]]);
        return { success: true, message: "팀 가입이 확정되었습니다!" };
      }
    }
    return { success: false, message: "초대받은 팀 신청 내역을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function cancelMyFacilityApplication(facility, applicationId, token) {
  try {
    const session = requireSession(token);
    const appSheet = getFacilityAppSheet_(facility);
    const appData = appSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < appData.length; i++) { if (appData[i][0] === applicationId) { rowIndex = i + 1; break; } }
    if (rowIndex === -1) return { success: false, message: "신청 내역을 찾을 수 없습니다." };
    if (appData[rowIndex - 1][3].toString() !== session.id.toString()) return { success: false, message: "본인이 만든 팀만 취소할 수 있습니다." };
    appSheet.deleteRow(rowIndex);
    const memberSheet = getFacilityMemberSheet_(facility);
    const memberData = memberSheet.getDataRange().getValues();
    for (let i = memberData.length - 1; i >= 1; i--) { if (memberData[i][0] === applicationId) memberSheet.deleteRow(i + 1); }
    return { success: true, message: "팀 신청을 취소했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getFacilityBoard(facility, token) {
  try {
    const session = requireSession(token);
    const slots = getAllFacilitySlots_(facility);
    const now = new Date();

    const openSlots = slots.filter(s => s.status === 'OPEN' && now >= new Date(s.appStart) && now <= new Date(s.appEnd));
    const upcomingSlots = slots.filter(s => s.status === 'OPEN' && now < new Date(s.appStart));
    const resultSlots = slots.filter(s => s.status === 'DRAWN').sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 10);

    const eligibleOpenSlots = openSlots.filter(s => {
      if (s.mode === 'GRADE') return session.role === '학생' && session.id.toString().charAt(0) === s.grade.toString();
      return true;
    });

    const appData = getFacilityAppSheet_(facility).getDataRange().getValues();
    const memberData = getFacilityMemberSheet_(facility).getDataRange().getValues();
    const memberRows = memberData.slice(1);

    const myTeams = {};
    for (let i = 1; i < appData.length; i++) {
      if (appData[i][3].toString() === session.id.toString()) {
        const appId = appData[i][0];
        const members = memberRows.filter(r => r[0] === appId).map(r => ({ id: r[1], name: r[2], confirmed: r[4] === true }));
        myTeams[appData[i][1]] = { applicationId: appId, teamName: appData[i][2], members: members, complete: members.every(m => m.confirmed) };
      }
    }

    const pendingInvites = [];
    memberRows.forEach(r => {
      if (r[1].toString() === session.id.toString() && r[4] !== true) {
        const appRow = appData.find(a => a[0] === r[0]);
        if (appRow) pendingInvites.push({ applicationId: r[0], slotId: appRow[1], teamName: appRow[2], leaderName: appRow[4] });
      }
    });

    const resultDetails = {};
    resultSlots.forEach(s => {
      let ids = [];
      try { ids = JSON.parse(s.resultJson || '[]'); } catch (e) { ids = []; }
      resultDetails[s.id] = [];
      for (let i = 1; i < appData.length; i++) {
        if (ids.indexOf(appData[i][0]) > -1) {
          const members = memberRows.filter(r => r[0] === appData[i][0]).map(r => r[2]);
          resultDetails[s.id].push({ teamName: appData[i][2], members: members.join(', ') });
        }
      }
    });

    return sanitizeDates_({
      success: true,
      eligibleOpenSlots: eligibleOpenSlots,
      upcomingSlots: upcomingSlots,
      resultSlots: resultSlots,
      myTeams: myTeams,
      pendingInvites: pendingInvites,
      resultDetails: resultDetails,
      isManager: isFacilityManager_(facility, session.id)
    });
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 시설 관리학생(위반 관리 권한) 명단 - 체육교사 전용 관리
// ==========================================================
function isFacilityManager_(facility, studentId) {
  const sheet = getFacilityManagerSheet_(facility);
  const data = sheet.getDataRange().getValues();
  const idStr = studentId.toString().trim();
  for (let i = 1; i < data.length; i++) { if (data[i][0] && data[i][0].toString().trim() === idStr) return true; }
  return false;
}

function requireFacilityManager(facility, token) {
  const session = requireSession(token);
  const isTeacher = (session.role === '체육교사');
  const isManagerStudent = (session.role === '학생' && isFacilityManager_(facility, session.id));
  if (!isTeacher && !isManagerStudent) throw new Error("해당 시설의 관리 권한이 있는 학생만 이용할 수 있습니다.");
  return session;
}

function addFacilityManager(facility, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = getFacilityManagerSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] && data[i][0].toString().trim() === idStr) return { success: false, message: "이미 등록된 관리학생입니다." }; }
    const info = lookupStudentInfo(idStr);
    if (!info) return { success: false, message: "명렬표에서 학번을 찾을 수 없습니다." };
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([idStr, info.name, info.grade, dateStr]);
    return { success: true, message: info.name + " 학생을 " + FACILITY_LABELS[facility] + " 관리학생으로 등록했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeFacilityManager(facility, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = getFacilityManagerSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] && data[i][0].toString().trim() === idStr) { sheet.deleteRow(i + 1); return { success: true, message: "명단에서 삭제했습니다." }; } }
    return { success: false, message: "명단에서 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getFacilityManagers(facility, token) {
  try {
    requirePeTeacher(token);
    const data = getFacilityManagerSheet_(facility).getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) { if (data[i][0]) list.push({ id: data[i][0].toString(), name: data[i][1], grade: data[i][2], regDate: toDateStr_(data[i][3]) }); }
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// ==========================================================
// 시설 위반 기록 (관리학생 또는 체육교사)
// ==========================================================
function addFacilityViolation(facility, studentId, reason, violationDate, token) {
  try {
    const session = requireFacilityManager(facility, token);
    const idStr = studentId.toString().trim();
    reason = (reason || '').toString().trim();
    if (!reason) return { success: false, message: "위반 사유를 입력하세요." };
    if (!violationDate) return { success: false, message: "위반 발생일자를 입력하세요." };
    const info = lookupStudentInfo(idStr);
    if (!info) return { success: false, message: "명렬표에서 학번을 찾을 수 없습니다." };
    const recordId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    getFacilityViolationSheet_(facility).appendRow([recordId, idStr, info.name, reason, violationDate, dateStr, session.id, session.name]);
    return { success: true, message: `${info.name} 학생의 위반 기록이 등록되었습니다.` };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeFacilityViolation(facility, recordId, token) {
  try {
    requireFacilityManager(facility, token);
    const sheet = getFacilityViolationSheet_(facility);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === recordId) { sheet.deleteRow(i + 1); return { success: true, message: "위반 기록을 삭제했습니다." }; } }
    return { success: false, message: "기록을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function facilityViolationRowToObj_(row) {
  return { id: row[0], studentId: row[1], name: row[2], reason: row[3], violationDate: toDateStr_(row[4]), registeredAt: toDateTimeStr_(row[5]), recorderName: row[7] };
}

function getFacilityViolations(facility, token) {
  try {
    requireFacilityManager(facility, token);
    const data = getFacilityViolationSheet_(facility).getDataRange().getValues();
    const list = [];
    const countMap = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push(facilityViolationRowToObj_(data[i]));
      countMap[data[i][1]] = (countMap[data[i][1]] || 0) + 1;
    }
    list.reverse();
    return sanitizeDates_({ success: true, list: list, countMap: countMap });
  } catch (e) { return { success: false, message: e.message, list: [], countMap: {} }; }
}

// 풋살장·체육관·운동장 위반기록을 한 화면에서 모아 보기 위한 통합 조회 (체육교사 전용).
function getAllFacilityViolations(token) {
  try {
    requirePeTeacher(token);
    const list = [];
    const countMap = {};
    Object.keys(FACILITY_LABELS).forEach(facility => {
      const data = getFacilityViolationSheet_(facility).getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        const obj = facilityViolationRowToObj_(data[i]);
        obj.facility = facility;
        obj.facilityLabel = FACILITY_LABELS[facility];
        list.push(obj);
        countMap[obj.studentId] = (countMap[obj.studentId] || 0) + 1;
      }
    });
    list.sort((a, b) => new Date(b.violationDate) - new Date(a.violationDate));
    return sanitizeDates_({ success: true, list: list, countMap: countMap });
  } catch (e) { return { success: false, message: e.message, list: [], countMap: {} }; }
}

// ==========================================================
// 체육관 요일별 스포츠클럽 준비일 배정 (체육관 슬롯 생성 시 참고용)
// ==========================================================
function getGymWeekdaySchedule(token) {
  try {
    requireSession(token);
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('체육관_요일배정').getDataRange().getValues();
    const schedule = {};
    for (let i = 1; i < data.length; i++) { if (data[i][0]) schedule[data[i][0]] = data[i][1] || ''; }
    return { success: true, schedule: schedule };
  } catch (e) { return { success: false, message: e.message, schedule: {} }; }
}

function updateGymWeekdaySchedule(schedule, token) {
  try {
    requirePeTeacher(token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('체육관_요일배정');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const day = data[i][0];
      if (schedule[day] !== undefined) sheet.getRange(i + 1, 2).setValue(schedule[day]);
    }
    return { success: true, message: "요일별 배정을 저장했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

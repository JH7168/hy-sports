// ==========================================================
// 친사리그 (교내 스포츠리그) - 해마다 세부 규칙이 달라지므로 종목을
// 자유롭게 여러 개 만들 수 있게 설계했다. 종목마다 스포츠 종류(프리셋
// 또는 직접입력), 진행방식(토너먼트/리그전), 참가방식(반별/지원자모집),
// 대상학년, 기간(시작~종료일)을 따로 정하고, 팀을 등록한 뒤 랜덤으로
// 대진을 짤 수 있다. 날짜별 일정(대진+메모)은 달력 형태로 볼 수 있다.
// 1, 2학년만 이용 가능(3학년/일반 열람 불가).
// ==========================================================
const LEAGUE_SPORT_PRESETS = ['축구', '풋살', '농구', '족구', '배구'];

function setupLeagueSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    '친사리그_종목': ['종목ID', '종목명', '형식', '참가방식', '대상학년', '시작일', '종료일', '상태', '생성일시', '생성교사'],
    '친사리그_팀': ['종목ID', '팀ID', '팀명', '비고'],
    '친사리그_대진': ['종목ID', '매치ID', '회차', '팀A', '팀B', '예정일자', '장소', '결과'],
    '친사리그_일정메모': ['종목ID', '메모ID', '날짜', '내용']
  };
  Object.keys(sheets).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      const headers = sheets[name];
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#5e35b1").setFontColor("white").setFontWeight("bold");
      if (name === '친사리그_종목') { sheet.getRange("F:G").setNumberFormat("@"); }
      if (name === '친사리그_대진') { sheet.getRange("F:F").setNumberFormat("@"); }
      if (name === '친사리그_일정메모') { sheet.getRange("C:C").setNumberFormat("@"); }
    }
  });
}

function getLeagueEventSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('친사리그_종목'); }
function getLeagueTeamSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('친사리그_팀'); }
function getLeagueMatchSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('친사리그_대진'); }
function getLeagueNoteSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('친사리그_일정메모'); }

function leagueEventRowToObj_(row) {
  return { id: row[0], sportName: row[1], format: row[2], joinMode: row[3], targetGrade: row[4], startDate: toDateStr_(row[5]), endDate: toDateStr_(row[6]), status: row[7], createdAt: row[8], createdBy: row[9] };
}

// 체육교사이거나, 1·2학년 학생만 친사리그를 볼 수 있다(3학년 제외).
function requireLeagueViewAccess(token) {
  const session = requireSession(token);
  const isTeacher = (session.role === '체육교사' || session.role === '교사');
  const isEligibleStudent = (session.role === '학생' && ['1', '2'].indexOf(session.id.toString().charAt(0)) > -1);
  if (!isTeacher && !isEligibleStudent) throw new Error("친사리그는 1, 2학년만 이용할 수 있습니다.");
  return session;
}

// ==========================================================
// 종목 관리 (체육교사 전용 생성/수정, 조회는 열람권한만 있으면 가능)
// ==========================================================
function createLeagueEvent(sportName, format, joinMode, targetGrade, startDate, endDate, token) {
  try {
    const session = requirePeTeacher(token);
    sportName = (sportName || '').toString().trim();
    if (!sportName) return { success: false, message: "종목명을 입력하세요." };
    if (!startDate || !endDate) return { success: false, message: "시작일과 종료일을 입력하세요." };
    if (new Date(startDate) > new Date(endDate)) return { success: false, message: "시작일이 종료일보다 늦을 수 없습니다." };
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    getLeagueEventSheet_().appendRow([id, sportName, format === 'LEAGUE' ? 'LEAGUE' : 'TOURNAMENT', joinMode === 'REGISTRATION' ? 'REGISTRATION' : 'CLASS', targetGrade === '1' || targetGrade === '2' ? targetGrade : 'BOTH', startDate, endDate, 'ACTIVE', dateStr, session.name]);
    return { success: true, message: "친사리그 종목이 생성되었습니다.", eventId: id };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateLeagueEvent(eventId, sportName, format, joinMode, targetGrade, startDate, endDate, token) {
  try {
    requirePeTeacher(token);
    sportName = (sportName || '').toString().trim();
    if (!sportName) return { success: false, message: "종목명을 입력하세요." };
    if (new Date(startDate) > new Date(endDate)) return { success: false, message: "시작일이 종료일보다 늦을 수 없습니다." };
    const sheet = getLeagueEventSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        sheet.getRange(i + 1, 2, 1, 6).setValues([[sportName, format === 'LEAGUE' ? 'LEAGUE' : 'TOURNAMENT', joinMode === 'REGISTRATION' ? 'REGISTRATION' : 'CLASS', targetGrade === '1' || targetGrade === '2' ? targetGrade : 'BOTH', startDate, endDate]]);
        return { success: true, message: "종목 정보를 수정했습니다." };
      }
    }
    return { success: false, message: "종목을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function cancelLeagueEvent(eventId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getLeagueEventSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === eventId) { sheet.getRange(i + 1, 8).setValue('CANCELLED'); return { success: true, message: "종목을 취소(비공개) 처리했습니다." }; } }
    return { success: false, message: "종목을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteLeagueEvent(eventId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getLeagueEventSheet_();
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = data.length - 1; i >= 1; i--) { if (data[i][0] === eventId) { sheet.deleteRow(i + 1); found = true; break; } }
    if (!found) return { success: false, message: "종목을 찾을 수 없습니다." };
    [getLeagueTeamSheet_(), getLeagueMatchSheet_(), getLeagueNoteSheet_()].forEach(sh => {
      const d = sh.getDataRange().getValues();
      for (let i = d.length - 1; i >= 1; i--) { if (d[i][0] === eventId) sh.deleteRow(i + 1); }
    });
    return { success: true, message: "종목과 관련 데이터를 모두 삭제했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getLeagueEvents(token) {
  try {
    const session = requireLeagueViewAccess(token);
    const isTeacher = (session.role === '체육교사' || session.role === '교사');
    const data = getLeagueEventSheet_().getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const ev = leagueEventRowToObj_(data[i]);
      if (!isTeacher) {
        if (ev.status !== 'ACTIVE') continue;
        if (ev.targetGrade !== 'BOTH' && ev.targetGrade !== session.id.toString().charAt(0)) continue;
      }
      list.push(ev);
    }
    list.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// ==========================================================
// 팀 관리 (체육교사 전용)
// ==========================================================
function addLeagueTeam(eventId, teamName, note, token) {
  try {
    requirePeTeacher(token);
    teamName = (teamName || '').toString().trim();
    if (!teamName) return { success: false, message: "팀명을 입력하세요." };
    const teamId = Utilities.getUuid();
    getLeagueTeamSheet_().appendRow([eventId, teamId, teamName, note || '']);
    return { success: true, message: "팀이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 반별 참가방식 편의 기능: 1~n반 팀을 한 번에 생성한다.
function addLeagueClassTeams(eventId, classCount, token) {
  try {
    requirePeTeacher(token);
    const n = parseInt(classCount, 10) || 8;
    const sheet = getLeagueTeamSheet_();
    const existing = sheet.getDataRange().getValues();
    const existingNames = {};
    for (let i = 1; i < existing.length; i++) { if (existing[i][0] === eventId) existingNames[existing[i][2]] = true; }
    let added = 0;
    for (let i = 1; i <= n; i++) {
      const name = i + '반';
      if (existingNames[name]) continue;
      sheet.appendRow([eventId, Utilities.getUuid(), name, '']);
      added++;
    }
    return { success: true, message: `${added}개 반 팀이 추가되었습니다.` };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeLeagueTeam(eventId, teamId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getLeagueTeamSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === eventId && data[i][1] === teamId) { sheet.deleteRow(i + 1); return { success: true, message: "팀을 삭제했습니다." }; } }
    return { success: false, message: "팀을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getLeagueTeams(eventId, token) {
  try {
    requireLeagueViewAccess(token);
    const data = getLeagueTeamSheet_().getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) { if (data[i][0] === eventId) list.push({ id: data[i][1], name: data[i][2], note: data[i][3] }); }
    return { success: true, list: list };
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// ==========================================================
// 대진 관리 - 랜덤 배정(토너먼트 1회전 랜덤 페어링 / 리그전 라운드로빈)
// ==========================================================
function generateLeagueMatches(eventId, token) {
  try {
    requirePeTeacher(token);
    const eventData = getLeagueEventSheet_().getDataRange().getValues();
    let event = null;
    for (let i = 1; i < eventData.length; i++) { if (eventData[i][0] === eventId) { event = leagueEventRowToObj_(eventData[i]); break; } }
    if (!event) return { success: false, message: "종목을 찾을 수 없습니다." };

    const teamData = getLeagueTeamSheet_().getDataRange().getValues();
    let teams = [];
    for (let i = 1; i < teamData.length; i++) { if (teamData[i][0] === eventId) teams.push(teamData[i][2]); }
    if (teams.length < 2) return { success: false, message: "대진을 짜려면 팀이 최소 2개 이상 필요합니다." };

    // 기존 대진 삭제 후 재생성
    const matchSheet = getLeagueMatchSheet_();
    const matchData = matchSheet.getDataRange().getValues();
    for (let i = matchData.length - 1; i >= 1; i--) { if (matchData[i][0] === eventId) matchSheet.deleteRow(i + 1); }

    // 셔플
    for (let i = teams.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const tmp = teams[i]; teams[i] = teams[j]; teams[j] = tmp; }

    const rows = [];
    if (event.format === 'LEAGUE') {
      const pairs = [];
      for (let i = 0; i < teams.length; i++) { for (let j = i + 1; j < teams.length; j++) pairs.push([teams[i], teams[j]]); }
      for (let i = pairs.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const tmp = pairs[i]; pairs[i] = pairs[j]; pairs[j] = tmp; }
      pairs.forEach(p => rows.push([eventId, Utilities.getUuid(), '리그전', p[0], p[1], '', '', '']));
    } else {
      for (let i = 0; i < teams.length; i += 2) {
        if (i + 1 < teams.length) rows.push([eventId, Utilities.getUuid(), '1라운드', teams[i], teams[i + 1], '', '', '']);
        else rows.push([eventId, Utilities.getUuid(), '1라운드', teams[i], '', '', '', '부전승']);
      }
    }
    rows.forEach(r => matchSheet.appendRow(r));
    return { success: true, message: `${rows.length}개의 대진이 랜덤으로 편성되었습니다.` };
  } catch (e) { return { success: false, message: e.message }; }
}

function addLeagueMatch(eventId, round, teamA, teamB, scheduledDate, place, token) {
  try {
    requirePeTeacher(token);
    if (!teamA) return { success: false, message: "팀A는 필수입니다." };
    getLeagueMatchSheet_().appendRow([eventId, Utilities.getUuid(), round || '', teamA, teamB || '', scheduledDate || '', place || '', '']);
    return { success: true, message: "대진이 추가되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateLeagueMatch(eventId, matchId, scheduledDate, place, result, token) {
  try {
    requirePeTeacher(token);
    const sheet = getLeagueMatchSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId && data[i][1] === matchId) {
        sheet.getRange(i + 1, 6, 1, 3).setValues([[scheduledDate || '', place || '', result || '']]);
        return { success: true, message: "대진 정보를 저장했습니다." };
      }
    }
    return { success: false, message: "대진을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeLeagueMatch(eventId, matchId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getLeagueMatchSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === eventId && data[i][1] === matchId) { sheet.deleteRow(i + 1); return { success: true, message: "대진을 삭제했습니다." }; } }
    return { success: false, message: "대진을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getLeagueMatches(eventId, token) {
  try {
    requireLeagueViewAccess(token);
    const data = getLeagueMatchSheet_().getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) list.push({ id: data[i][1], round: data[i][2], teamA: data[i][3], teamB: data[i][4], scheduledDate: toDateStr_(data[i][5]), place: data[i][6], result: data[i][7] });
    }
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// ==========================================================
// 날짜별 일정메모 (대진 외에 자유롭게 남기는 공지성 일정, 예: 개회식)
// ==========================================================
function addLeagueDayNote(eventId, date, content, token) {
  try {
    requirePeTeacher(token);
    content = (content || '').toString().trim();
    if (!date || !content) return { success: false, message: "날짜와 내용을 모두 입력하세요." };
    getLeagueNoteSheet_().appendRow([eventId, Utilities.getUuid(), date, content]);
    return { success: true, message: "일정이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeLeagueDayNote(eventId, noteId, token) {
  try {
    requirePeTeacher(token);
    const sheet = getLeagueNoteSheet_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === eventId && data[i][1] === noteId) { sheet.deleteRow(i + 1); return { success: true, message: "일정을 삭제했습니다." }; } }
    return { success: false, message: "일정을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 달력 렌더링에 필요한 정보를 한 번에 묶어서 반환 (날짜별 대진+메모)
function getLeagueCalendar(eventId, token) {
  try {
    requireLeagueViewAccess(token);
    const eventData = getLeagueEventSheet_().getDataRange().getValues();
    let event = null;
    for (let i = 1; i < eventData.length; i++) { if (eventData[i][0] === eventId) { event = leagueEventRowToObj_(eventData[i]); break; } }
    if (!event) return { success: false, message: "종목을 찾을 수 없습니다." };

    const matchData = getLeagueMatchSheet_().getDataRange().getValues();
    const noteData = getLeagueNoteSheet_().getDataRange().getValues();
    const byDate = {};
    const ensure = (d) => { if (!byDate[d]) byDate[d] = { matches: [], notes: [] }; return byDate[d]; };
    for (let i = 1; i < matchData.length; i++) {
      if (matchData[i][0] === eventId && matchData[i][5]) {
        ensure(toDateStr_(matchData[i][5])).matches.push({ id: matchData[i][1], round: matchData[i][2], teamA: matchData[i][3], teamB: matchData[i][4], place: matchData[i][6], result: matchData[i][7] });
      }
    }
    for (let i = 1; i < noteData.length; i++) {
      if (noteData[i][0] === eventId) ensure(toDateStr_(noteData[i][2])).notes.push({ id: noteData[i][1], content: noteData[i][3] });
    }
    return sanitizeDates_({ success: true, event: event, byDate: byDate });
  } catch (e) { return { success: false, message: e.message }; }
}

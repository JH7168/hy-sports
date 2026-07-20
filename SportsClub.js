// ==========================================================
// 스포츠클럽 (축구부/농구부/배구부/배드민턴부) 명단 + 자유게시판
// 체육교사가 부원을 등록하면, 그 부원들은 학생 화면의 해당 스포츠클럽
// 탭에서 자유게시판(공지, 훈련일지, 잡담 등 자유 주제)을 직접 꾸밀 수
// 있다. 글 작성은 부원 누구나, 수정·삭제는 작성자 본인 또는 체육교사만
// 가능하다.
// ==========================================================
const CLUB_LABELS = { SOCCER: '축구부', BASKETBALL: '농구부', VOLLEYBALL: '배구부', BADMINTON: '배드민턴부' };

function setupSportsClubSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let memberSheet = ss.getSheetByName('스포츠클럽_명단');
  if (!memberSheet) {
    memberSheet = ss.insertSheet('스포츠클럽_명단');
    memberSheet.appendRow(['부서', '학번', '이름', '학년', '등록일']);
    memberSheet.getRange("A1:E1").setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  }
  let postSheet = ss.getSheetByName('스포츠클럽_게시글');
  if (!postSheet) {
    postSheet = ss.insertSheet('스포츠클럽_게시글');
    postSheet.appendRow(['부서', '글ID', '작성자ID', '작성자이름', '제목', '내용', '작성일시', '수정일시']);
    postSheet.getRange("A1:H1").setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  }
}

function isSportsClubMember_(club, studentId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
  const idStr = studentId.toString().trim();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) return true; }
  return false;
}

// 체육교사이거나, 해당 부에 등록된 학생 본인만 허용
function requireSportsClubAccess(club, token) {
  const session = requireSession(token);
  const isTeacher = (session.role === '체육교사');
  const isMember = (session.role === '학생' && isSportsClubMember_(club, session.id));
  if (!isTeacher && !isMember) throw new Error(CLUB_LABELS[club] + " 소속 학생만 이용할 수 있습니다.");
  return session;
}

// ==========================================================
// 부원 명단 관리 (체육교사 전용)
// ==========================================================
function addSportsClubMember(club, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) return { success: false, message: "이미 등록된 부원입니다." }; }
    const info = lookupStudentInfo(idStr);
    if (!info) return { success: false, message: "명렬표에서 학번을 찾을 수 없습니다." };
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([club, idStr, info.name, info.grade, dateStr]);
    return { success: true, message: info.name + "님을 " + CLUB_LABELS[club] + " 부원으로 등록했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeSportsClubMember(club, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) { sheet.deleteRow(i + 1); return { success: true, message: "명단에서 삭제했습니다." }; } }
    return { success: false, message: "명단에서 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getSportsClubMembers(club, token) {
  try {
    requirePeTeacher(token);
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) { if (data[i][0] === club) list.push({ id: data[i][1].toString(), name: data[i][2], grade: data[i][3], regDate: toDateStr_(data[i][4]) }); }
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// 로그인한 학생이 소속된 클럽 코드 목록 (학생 nav 접근 제어용)
function getMySportsClubs(token) {
  try {
    const session = requireSession(token);
    if (session.role !== '학생') return { success: true, clubs: [] };
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
    const clubs = [];
    for (let i = 1; i < data.length; i++) { if (data[i][1] && data[i][1].toString().trim() === session.id.toString()) clubs.push(data[i][0]); }
    return { success: true, clubs: clubs };
  } catch (e) { return { success: false, message: e.message, clubs: [] }; }
}

// ==========================================================
// 부원 자유게시판
// ==========================================================
function createSportsClubPost(club, title, content, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    title = (title || '').toString().trim();
    content = (content || '').toString().trim();
    if (!title || !content) return { success: false, message: "제목과 내용을 모두 입력하세요." };
    const postId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글').appendRow([club, postId, session.id, session.name, title, content, dateStr, dateStr]);
    return { success: true, message: "글이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateSportsClubPost(club, postId, title, content, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    title = (title || '').toString().trim();
    content = (content || '').toString().trim();
    if (!title || !content) return { success: false, message: "제목과 내용을 모두 입력하세요." };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] === postId) {
        if (data[i][2].toString() !== session.id.toString() && session.role !== '체육교사') return { success: false, message: "본인이 작성한 글만 수정할 수 있습니다." };
        const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        sheet.getRange(i + 1, 5, 1, 4).setValues([[title, content, data[i][6], dateStr]]);
        return { success: true, message: "글이 수정되었습니다." };
      }
    }
    return { success: false, message: "글을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteSportsClubPost(club, postId, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] === postId) {
        if (data[i][2].toString() !== session.id.toString() && session.role !== '체육교사') return { success: false, message: "본인이 작성한 글만 삭제할 수 있습니다." };
        sheet.deleteRow(i + 1);
        return { success: true, message: "글이 삭제되었습니다." };
      }
    }
    return { success: false, message: "글을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getSportsClubPosts(club, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club) list.push({ id: data[i][1], authorId: data[i][2], authorName: data[i][3], title: data[i][4], content: data[i][5], createdAt: toDateTimeStr_(data[i][6]), updatedAt: toDateTimeStr_(data[i][7]), mine: data[i][2].toString() === session.id.toString() });
    }
    list.reverse();
    return sanitizeDates_({ success: true, list: list });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

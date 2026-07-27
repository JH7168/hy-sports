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
    memberSheet.appendRow(['부서', '학번', '이름', '학년', '등록일', '주장여부', '부주장여부']);
    memberSheet.getRange("A1:G1").setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  } else {
    // 기존에 이미 만들어져 있던 시트에는 주장/부주장 컬럼이 없을 수 있으므로 최초 1회 보정한다.
    if (memberSheet.getRange(1, 6).getValue() !== '주장여부') memberSheet.getRange(1, 6).setValue('주장여부').setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
    if (memberSheet.getRange(1, 7).getValue() !== '부주장여부') memberSheet.getRange(1, 7).setValue('부주장여부').setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  }
  let postSheet = ss.getSheetByName('스포츠클럽_게시글');
  if (!postSheet) {
    postSheet = ss.insertSheet('스포츠클럽_게시글');
    postSheet.appendRow(['부서', '글ID', '작성자ID', '작성자이름', '제목', '내용', '작성일시', '수정일시', '설문여부', '일정시작', '일정종료', '마감일시']);
    postSheet.getRange("A1:L1").setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  } else {
    // 헤더 검사 셀을 한 번씩 따로 읽으면 이 함수가 매 요청마다 실행되므로 그만큼 호출이
    // 쌓이는데, 여기서는 필요한 두 셀(I열·J열)을 한 번의 읽기로 같이 가져와 비교한다.
    const hdr = postSheet.getRange(1, 9, 1, 2).getValues()[0];
    if (hdr[0] !== '설문여부') {
      // 기존에 이미 만들어져 있던 게시글 시트에는 모집 설문 관련 컬럼이 없을 수 있으므로 보정한다.
      postSheet.getRange(1, 9, 1, 4).setValues([['설문여부', '일정시작', '일정종료', '마감일시']]).setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
    } else if (hdr[1] === '요일') {
      // 요일+시간 2컬럼을 일정시작+일정종료로 재해석한다(컬럼 수는 그대로 유지). 이 기능은
      // 방금 도입되어 실제 데이터가 거의 없을 것으로 보고, 기존 값 보존 없이 헤더만 정리한다.
      postSheet.getRange(1, 10, 1, 2).setValues([['일정시작', '일정종료']]).setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
    } else if (hdr[1] === '일정일시') {
      // 일정을 시작~종료 구간으로 나눈다: 기존 일정일시 컬럼을 일정시작으로 삼고, 그 뒤에
      // 일정종료 컬럼을 새로 끼워넣는다(기존 종료값은 없었으므로 새로 입력받는다).
      postSheet.insertColumnAfter(10);
      postSheet.getRange(1, 10, 1, 2).setValues([['일정시작', '일정종료']]).setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
    }
  }
  // 첨부파일(홍보물 jpg·한글 파일 등) 컬럼이 없는 기존 시트는 한 번만 보정한다.
  if (postSheet.getRange(1, 13).getValue() !== '첨부파일ID') {
    postSheet.getRange(1, 13, 1, 3).setValues([['첨부파일ID', '첨부파일명', '첨부파일타입']]).setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  }
  let commentSheet = ss.getSheetByName('스포츠클럽_댓글');
  if (!commentSheet) {
    commentSheet = ss.insertSheet('스포츠클럽_댓글');
    commentSheet.appendRow(['부서', '댓글ID', '글ID', '작성자ID', '작성자이름', '내용', '작성일시']);
    commentSheet.getRange("A1:G1").setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  }
  // 모집 설문(글에 딸린 참여/불참 응답) - 부원이 응답하면 학번당 1행으로 upsert된다.
  let pollSheet = ss.getSheetByName('스포츠클럽_설문응답');
  if (!pollSheet) {
    pollSheet = ss.insertSheet('스포츠클럽_설문응답');
    pollSheet.appendRow(['부서', '글ID', '학번', '이름', '응답', '응답일시']);
    pollSheet.getRange("A1:F1").setBackground("#3949ab").setFontColor("white").setFontWeight("bold");
  }
}

function isSportsClubMember_(club, studentId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
  const idStr = studentId.toString().trim();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) return true; }
  return false;
}

// 해당 부에 등록된 학생이면서 주장(Y)으로 표시된 경우만 true.
function isSportsClubCaptain_(club, studentId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
  const idStr = studentId.toString().trim();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) return data[i][5] === 'Y'; }
  return false;
}

// 부원 모집(등록/제외)은 체육교사 또는 그 부의 주장만 할 수 있다. 부주장·일반 부원은 불가.
function requireSportsClubCaptain_(club, token) {
  const session = requireSession(token);
  if (session.role !== '학생' || !isSportsClubCaptain_(club, session.id)) throw new Error(CLUB_LABELS[club] + " 주장만 이용할 수 있는 기능입니다.");
  return session;
}

// 체육교사이거나, 해당 부에 등록된 학생 본인만 허용.
// 게시글/설문 관련 함수들은 모두 이 함수를 거치므로, 여기서 한 번 시트 존재를 보장해두면
// (예: 배포 시점 차이로 새로 추가된 시트인 스포츠클럽_설문응답이 아직 없는 경우) 개별
// 함수에서 "Cannot read properties of null (reading 'getDataRange')" 오류가 나는 것을 막을 수 있다.
function requireSportsClubAccess(club, token) {
  setupSportsClubSheets();
  const session = requireSession(token);
  const isTeacher = (session.role === '체육교사');
  const isMember = (session.role === '학생' && isSportsClubMember_(club, session.id));
  if (!isTeacher && !isMember) throw new Error(CLUB_LABELS[club] + " 소속 학생만 이용할 수 있습니다.");
  return session;
}

// ==========================================================
// 부원 명단 관리 (체육교사 또는 그 부의 주장)
// 실제 처리 로직은 core 함수에 모아두고, 교사용/주장용 진입점은 각자 서버에서
// 직접 권한을 검증한 뒤 core를 호출한다(클라이언트가 보낸 role은 신뢰하지 않는다).
// ==========================================================
function sportsClubAddMemberCore_(club, studentId) {
  const idStr = studentId.toString().trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) return { success: false, message: "이미 등록된 부원입니다." }; }
  const info = lookupStudentInfo(idStr);
  if (!info) return { success: false, message: "명렬표에서 학번을 찾을 수 없습니다." };
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  sheet.appendRow([club, idStr, info.name, info.grade, dateStr]);
  return { success: true, message: info.name + "님을 " + CLUB_LABELS[club] + " 부원으로 등록했습니다." };
}

function sportsClubRemoveMemberCore_(club, studentId) {
  const idStr = studentId.toString().trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) { sheet.deleteRow(i + 1); return { success: true, message: "명단에서 삭제했습니다." }; } }
  return { success: false, message: "명단에서 찾을 수 없습니다." };
}

function sportsClubMembersCore_(club) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) { if (data[i][0] === club) list.push({ id: data[i][1].toString(), name: data[i][2], grade: data[i][3], regDate: toDateStr_(data[i][4]), isCaptain: data[i][5] === 'Y', isViceCaptain: data[i][6] === 'Y' }); }
  // 학번이 "학년+반(2자리)+번호(2자리)"로 구성되어 있어, 문자열 그대로 정렬하면
  // 학년 -> 반 -> 번호 순서가 자연스럽게 맞춰진다.
  list.sort((a, b) => a.id.localeCompare(b.id));
  return list;
}

function addSportsClubMember(club, studentId, token) {
  try { requirePeTeacher(token); return sportsClubAddMemberCore_(club, studentId); } catch (e) { return { success: false, message: e.message }; }
}

function removeSportsClubMember(club, studentId, token) {
  try { requirePeTeacher(token); return sportsClubRemoveMemberCore_(club, studentId); } catch (e) { return { success: false, message: e.message }; }
}

function getSportsClubMembers(club, token) {
  try { requirePeTeacher(token); return sanitizeDates_({ success: true, list: sportsClubMembersCore_(club) }); } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// 주장은 부원을 자유롭게 모집·등록할 수 있지만, 주장·부주장 임명/해제는 여전히
// 체육교사 전용 기능이므로(임명 오·남용 방지) 이 경로로는 일반 부원만 제외할 수 있다.
function addSportsClubMemberByCaptain(club, studentId, token) {
  try { requireSportsClubCaptain_(club, token); return sportsClubAddMemberCore_(club, studentId); } catch (e) { return { success: false, message: e.message }; }
}

function removeSportsClubMemberByCaptain(club, studentId, token) {
  try {
    requireSportsClubCaptain_(club, token);
    const idStr = studentId.toString().trim();
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단').getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) {
        if (data[i][5] === 'Y' || data[i][6] === 'Y') return { success: false, message: "주장·부주장은 체육교사만 명단에서 제외할 수 있습니다." };
        break;
      }
    }
    return sportsClubRemoveMemberCore_(club, idStr);
  } catch (e) { return { success: false, message: e.message }; }
}

function getSportsClubMembersForCaptain(club, token) {
  try { requireSportsClubCaptain_(club, token); return sanitizeDates_({ success: true, list: sportsClubMembersCore_(club) }); } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// 부에 주장은 한 번에 한 명만 존재하도록, 새로 임명하면 기존 주장은 자동으로 해제한다.
// 부주장이 주장으로 승격되는 경우 부주장 겸직은 자동으로 해제한다.
function setSportsClubCaptain(club, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== club) continue;
      if (data[i][1] && data[i][1].toString().trim() === idStr) {
        sheet.getRange(i + 1, 6).setValue('Y');
        if (data[i][6] === 'Y') sheet.getRange(i + 1, 7).setValue('');
        found = true;
      } else if (data[i][5] === 'Y') { sheet.getRange(i + 1, 6).setValue(''); }
    }
    if (!found) return { success: false, message: "명단에서 학생을 찾을 수 없습니다." };
    return { success: true, message: "주장으로 임명했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeSportsClubCaptain(club, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) { sheet.getRange(i + 1, 6).setValue(''); return { success: true, message: "주장 임명을 해제했습니다." }; }
    }
    return { success: false, message: "명단에서 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// 부에 부주장도 한 번에 한 명만 존재하도록, 새로 임명하면 기존 부주장은 자동으로 해제한다.
// 이미 주장인 학생은 부주장을 겸직할 수 없다.
function setSportsClubViceCaptain(club, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) { if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) { targetRow = i; break; } }
    if (targetRow === -1) return { success: false, message: "명단에서 학생을 찾을 수 없습니다." };
    if (data[targetRow][5] === 'Y') return { success: false, message: "이미 주장으로 임명된 학생은 부주장으로 지정할 수 없습니다." };
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== club) continue;
      if (i === targetRow) sheet.getRange(i + 1, 7).setValue('Y');
      else if (data[i][6] === 'Y') sheet.getRange(i + 1, 7).setValue('');
    }
    return { success: true, message: "부주장으로 임명했습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function removeSportsClubViceCaptain(club, studentId, token) {
  try {
    requirePeTeacher(token);
    const idStr = studentId.toString().trim();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_명단');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] && data[i][1].toString().trim() === idStr) { sheet.getRange(i + 1, 7).setValue(''); return { success: true, message: "부주장 임명을 해제했습니다." }; }
    }
    return { success: false, message: "명단에서 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
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
// 대회 홍보물(jpg 이미지·한글 문서 등) 첨부파일을 저장할 드라이브 폴더.
// 최초 업로드 시 한 번만 만들고, 폴더 ID를 스크립트 속성에 저장해 재사용한다.
function sportsClubAttachFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPORTSCLUB_ATTACH_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) { /* 폴더가 지워졌으면 아래에서 새로 만든다 */ }
  }
  const folder = DriveApp.createFolder('스포츠클럽_첨부파일');
  props.setProperty('SPORTSCLUB_ATTACH_FOLDER_ID', folder.getId());
  return folder;
}

const SPORTSCLUB_IMAGE_EXTS = ['jpg', 'jpeg', 'png', 'gif', 'webp'];

// 첨부파일 데이터URL을 드라이브에 업로드하고 { fileId, name, type } 형태로 반환한다.
// type은 'image'(화면에 바로 미리보기) 또는 'file'(다운로드 링크만 제공, 한글 문서 등)이다.
function sportsClubSaveAttachment_(attachData, attachName) {
  if (attachData.length > 27000000) throw new Error("첨부파일 용량이 너무 큽니다(20MB 이하로 올려주세요).");
  const ext = (attachName.split('.').pop() || '').toLowerCase();
  const isImage = SPORTSCLUB_IMAGE_EXTS.indexOf(ext) !== -1;
  const bytes = Utilities.base64Decode(attachData.substr(attachData.indexOf('base64,') + 7));
  const mimeType = isImage ? ('image/' + (ext === 'jpg' ? 'jpeg' : ext)) : 'application/octet-stream';
  const blob = Utilities.newBlob(bytes, mimeType, attachName);
  const file = sportsClubAttachFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), name: attachName, type: isImage ? 'image' : 'file' };
}

// isPoll이 true면 일정시작/일정종료/마감일시를 함께 저장해 "모집 설문"이 딸린 글로 등록한다.
// attachData(데이터URL)와 attachName이 함께 오면 첨부파일(홍보물 이미지·한글 문서 등)도 저장한다.
function createSportsClubPost(club, title, content, isPoll, scheduleStart, scheduleEnd, deadline, attachData, attachName, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    title = (title || '').toString().trim();
    content = (content || '').toString().trim();
    const hasAttach = !!(attachData && attachName);
    if (!hasAttach && (!title || !content)) return { success: false, message: "제목과 내용을 모두 입력하세요." };
    const poll = !!isPoll;
    if (poll) {
      scheduleStart = (scheduleStart || '').toString().trim();
      scheduleEnd = (scheduleEnd || '').toString().trim();
      if (!scheduleStart || !scheduleEnd || !deadline) return { success: false, message: "모집 설문은 일정(시작·종료)과 마감일시를 모두 입력해야 합니다." };
      if (new Date(scheduleEnd) <= new Date(scheduleStart)) return { success: false, message: "종료 시간은 시작 시간보다 늦어야 합니다." };
    }
    let attachment = null;
    if (attachData && attachName) attachment = sportsClubSaveAttachment_(attachData, attachName);
    const postId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글').appendRow([
      club, postId, session.id, session.name, title, content, dateStr, dateStr, poll ? 'Y' : 'N',
      poll ? scheduleStart : '', poll ? scheduleEnd : '', poll ? deadline : '',
      attachment ? attachment.fileId : '', attachment ? attachment.name : '', attachment ? attachment.type : ''
    ]);
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
        const attachFileId = data[i][12];
        sheet.deleteRow(i + 1);
        if (attachFileId) { try { DriveApp.getFileById(attachFileId).setTrashed(true); } catch (e) { /* 이미 지워진 파일이면 무시 */ } }
        // 글이 삭제되면 그 글에 달린 댓글·설문 응답도 함께 정리한다.
        const commentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_댓글');
        const commentData = commentSheet.getDataRange().getValues();
        for (let j = commentData.length - 1; j >= 1; j--) { if (commentData[j][0] === club && commentData[j][2] === postId) commentSheet.deleteRow(j + 1); }
        const pollSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_설문응답');
        const pollData = pollSheet.getDataRange().getValues();
        for (let j = pollData.length - 1; j >= 1; j--) { if (pollData[j][0] === club && pollData[j][1] === postId) pollSheet.deleteRow(j + 1); }
        return { success: true, message: "글이 삭제되었습니다." };
      }
    }
    return { success: false, message: "글을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getSportsClubPosts(club, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    const amICaptain = session.role === '학생' && isSportsClubCaptain_(club, session.id);
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글').getDataRange().getValues();
    const commentData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_댓글').getDataRange().getValues();
    const commentsByPost = {};
    for (let i = 1; i < commentData.length; i++) {
      if (commentData[i][0] !== club) continue;
      const postId = commentData[i][2];
      if (!commentsByPost[postId]) commentsByPost[postId] = [];
      commentsByPost[postId].push({ id: commentData[i][1], authorId: commentData[i][3], authorName: commentData[i][4], content: commentData[i][5], createdAt: toDateTimeStr_(commentData[i][6]), mine: commentData[i][3].toString() === session.id.toString() });
    }
    const pollData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_설문응답').getDataRange().getValues();
    const responsesByPost = {};
    for (let i = 1; i < pollData.length; i++) {
      if (pollData[i][0] !== club) continue;
      const postId = pollData[i][1];
      if (!responsesByPost[postId]) responsesByPost[postId] = [];
      responsesByPost[postId].push({ studentId: pollData[i][2].toString(), name: pollData[i][3], response: pollData[i][4], respondedAt: toDateTimeStr_(pollData[i][5]) });
    }
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== club) continue;
      const postId = data[i][1];
      const item = { id: postId, authorId: data[i][2], authorName: data[i][3], title: data[i][4], content: data[i][5], createdAt: toDateTimeStr_(data[i][6]), updatedAt: toDateTimeStr_(data[i][7]), mine: data[i][2].toString() === session.id.toString(), comments: commentsByPost[postId] || [] };
      if (data[i][8] === 'Y') {
        const responses = responsesByPost[postId] || [];
        const mine = responses.find(r => r.studentId === session.id.toString());
        item.poll = {
          start: toDateTimeStr_(data[i][9]), end: toDateTimeStr_(data[i][10]), deadline: toDateTimeStr_(data[i][11]),
          closed: new Date() > new Date(data[i][11]),
          responses: responses,
          joinCount: responses.filter(r => r.response === '참여').length,
          declineCount: responses.filter(r => r.response === '불참').length,
          myResponse: mine ? mine.response : null
        };
      }
      const attachFileId = data[i][12];
      if (attachFileId) {
        const isImg = data[i][14] === 'image';
        item.attachment = {
          fileId: attachFileId, name: data[i][13], type: data[i][14],
          viewUrl: 'https://drive.google.com/file/d/' + attachFileId + '/view',
          // uc?export=view는 모바일 브라우저에서 깨진 이미지로 뜨는 경우가 많아 썸네일 엔드포인트를 사용한다.
          imageUrl: isImg ? ('https://drive.google.com/thumbnail?id=' + attachFileId + '&sz=w800') : null,
          fullImageUrl: isImg ? ('https://drive.google.com/thumbnail?id=' + attachFileId + '&sz=w1600') : null
        };
      }
      list.push(item);
    }
    list.reverse();
    return sanitizeDates_({ success: true, list: list, amICaptain: amICaptain });
  } catch (e) { return { success: false, message: e.message, list: [] }; }
}

// 모집 설문 참여/불참 응답 (마감 전까지 몇 번이든 바꿔서 다시 제출할 수 있다).
function submitSportsClubPollResponse(club, postId, response, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    if (response !== '참여' && response !== '불참') return { success: false, message: "참여 또는 불참 중 하나를 선택하세요." };
    const postData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_게시글').getDataRange().getValues();
    let post = null;
    for (let i = 1; i < postData.length; i++) { if (postData[i][0] === club && postData[i][1] === postId) { post = postData[i]; break; } }
    if (!post) return { success: false, message: "글을 찾을 수 없습니다." };
    if (post[8] !== 'Y') return { success: false, message: "모집 설문이 아닌 글입니다." };
    if (new Date() > new Date(post[11])) return { success: false, message: "설문 마감 시간이 지났습니다." };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_설문응답');
    const data = sheet.getDataRange().getValues();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] === postId && data[i][2].toString() === session.id.toString()) {
        sheet.getRange(i + 1, 5, 1, 2).setValues([[response, dateStr]]);
        return { success: true, message: "응답이 저장되었습니다." };
      }
    }
    sheet.appendRow([club, postId, session.id, session.name, response, dateStr]);
    return { success: true, message: "응답이 저장되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

// ==========================================================
// 부원 댓글 (부원 본인만 작성, 본인 또는 체육교사만 삭제)
// ==========================================================
function createSportsClubComment(club, postId, content, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    content = (content || '').toString().trim();
    if (!content) return { success: false, message: "댓글 내용을 입력하세요." };
    const commentId = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_댓글').appendRow([club, commentId, postId, session.id, session.name, content, dateStr]);
    return { success: true, message: "댓글이 등록되었습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteSportsClubComment(club, commentId, token) {
  try {
    const session = requireSportsClubAccess(club, token);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('스포츠클럽_댓글');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === club && data[i][1] === commentId) {
        if (data[i][3].toString() !== session.id.toString() && session.role !== '체육교사') return { success: false, message: "본인이 작성한 댓글만 삭제할 수 있습니다." };
        sheet.deleteRow(i + 1);
        return { success: true, message: "댓글이 삭제되었습니다." };
      }
    }
    return { success: false, message: "댓글을 찾을 수 없습니다." };
  } catch (e) { return { success: false, message: e.message }; }
}

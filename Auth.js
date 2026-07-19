// ==========================================================
// 세션 관리
// ==========================================================
// CacheService 최대 보관 시간은 21600초(6시간). getSession 호출 시마다
// 만료시간을 다시 늘려주므로("슬라이딩 세션"), 사용자가 앱을 계속
// 사용하는 동안에는 로그아웃되지 않고 6시간 이상 활동이 없을 때만 만료된다.
const SESSION_TTL_SECONDS = 21600;

function createSession(role, id, name) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify({ role: role, id: id, name: name }), SESSION_TTL_SECONDS);
  return token;
}

function getSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('sess_' + token);
  if (!raw) return null;
  cache.put('sess_' + token, raw, SESSION_TTL_SECONDS); // 활동 시 세션 연장
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function invalidateSession(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return { success: true };
}

// 로그인된 사용자만 허용 (역할 무관)
function requireSession(token) {
  const session = getSession(token);
  if (!session) throw new Error("로그인이 만료되었거나 유효하지 않습니다. 다시 로그인해주세요.");
  return session;
}

// 체육교사만 허용
function requirePeTeacher(token) {
  const session = requireSession(token);
  if (session.role !== '체육교사') throw new Error("체육교사 권한이 필요한 기능입니다.");
  return session;
}

// 본인(학생) 또는 교사만 허용 - PAPS 개인정보 조회용
function requireOwnerOrTeacher(token, studentId) {
  const session = requireSession(token);
  const isTeacher = (session.role === '교사' || session.role === '체육교사');
  const isSelf = (session.role === '학생' && session.id.toString().trim() === studentId.toString().trim());
  if (!isTeacher && !isSelf) throw new Error("조회 권한이 없습니다.");
  return session;
}

// ==========================================================
// 비밀번호 해싱 (SHA-256 + 랜덤 salt)
// 저장 형식: "salt$hashHex". 구버전 평문 비밀번호와도 호환되며,
// 평문으로 로그인에 성공하면 그 즉시 해시 형식으로 자동 전환된다.
// ==========================================================
function hashPassword(pw, salt) {
  salt = salt || Utilities.getUuid();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + salt, Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    let b = bytes[i]; if (b < 0) b += 256;
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return salt + '$' + hex;
}

function verifyPassword(inputPw, storedValue) {
  if (!storedValue) return false;
  const str = storedValue.toString();
  if (str.indexOf('$') === -1) return str === inputPw; // 레거시 평문 비밀번호 호환
  const salt = str.substring(0, str.indexOf('$'));
  return hashPassword(inputPw, salt) === str;
}

// ==========================================================
// 가입 인증 코드 (소스코드가 아닌 스크립트 속성에 보관)
// ==========================================================
function getSecretCode(role) {
  const props = PropertiesService.getScriptProperties();
  return role === '체육교사' ? props.getProperty('PE_TEACHER_CODE') : props.getProperty('TEACHER_CODE');
}

// ==========================================================
// 인증 및 회원가입 로직
// ==========================================================
function verifyLogin(role, userId, userPw) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inputId = userId.toString().trim();
    const inputPw = userPw.toString().trim();

    if (role === '학생') {
      const sheet = ss.getSheetByName('학생명렬표');
      if (!sheet) return { success: false, message: "'학생명렬표' 탭이 없습니다." };
      const data = sheet.getDataRange().getValues();
      for (let i = 20; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim() === inputId) {
          const stored = data[i][2];
          if (stored && verifyPassword(inputPw, stored)) {
            if (stored.toString().indexOf('$') === -1) {
              sheet.getRange(i + 1, 3).setNumberFormat("@").setValue(hashPassword(inputPw)); // 평문 -> 해시 자동 전환
            }
            const name = data[i][1].toString().trim();
            const token = createSession('학생', inputId, name);
            return { success: true, name: name, role: '학생', id: inputId, token: token };
          }
          return { success: false, message: "비밀번호가 일치하지 않습니다." };
        }
      }
      return { success: false, message: "가입되지 않은 학생이거나 학번이 틀렸습니다." };
    } else {
      const sheet = ss.getSheetByName('회원정보');
      if (!sheet) return { success: false, message: "'회원정보' 탭이 없습니다." };
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === role && data[i][1].toString().trim() === inputId) {
          const stored = data[i][2];
          if (stored && verifyPassword(inputPw, stored)) {
            if (stored.toString().indexOf('$') === -1) {
              sheet.getRange(i + 1, 3).setNumberFormat("@").setValue(hashPassword(inputPw));
            }
            const token = createSession(role, inputId, inputId);
            return { success: true, name: inputId, role: role, id: inputId, token: token };
          }
          return { success: false, message: "정보가 일치하지 않습니다." };
        }
      }
      return { success: false, message: "정보가 일치하지 않습니다." };
    }
  } catch (e) { return { success: false, message: "시스템 오류: " + e.message }; }
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
          sheet.getRange(i + 1, 3).setNumberFormat("@").setValue(hashPassword(inputPw));
          return { success: true, message: "학생 회원가입이 완료되었습니다! 로그인해주세요." };
        }
      }
      return { success: false, message: "명렬표에서 학생을 찾을 수 없습니다." };
    } else {
      const expectedCode = getSecretCode(role);
      if (!expectedCode || secretCode.toString().trim() !== expectedCode) return { success: false, message: "학교 가입 인증 코드가 틀렸습니다." };
      const ssSheet = ss.getSheetByName('회원정보');
      const data = ssSheet.getDataRange().getValues();
      for(let i = 1; i < data.length; i++) {
        if(data[i][0] === role && data[i][1].toString().trim() === inputId) return { success: false, message: "이미 가입된 아이디입니다." };
      }
      const nextRow = ssSheet.getLastRow() + 1;
      ssSheet.getRange(nextRow, 1).setValue(role);
      ssSheet.getRange(nextRow, 2).setNumberFormat("@").setValue(inputId);
      ssSheet.getRange(nextRow, 3).setNumberFormat("@").setValue(hashPassword(inputPw));
      return { success: true, message: "교사 회원가입이 완료되었습니다! 로그인해주세요." };
    }
  } catch(e) { return { success: false, message: "가입 처리 중 오류 발생: " + e.message }; }
}

// ==========================================================
// 공통 유틸리티 (연도 계산, 숫자 파싱 등)
// ==========================================================

// 한국 학사년도는 3월에 시작하므로 1~2월은 이전 학년도로 취급한다.
function getCurrentAcademicYear() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const y = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);
  const m = parseInt(Utilities.formatDate(now, tz, 'MM'), 10);
  return (m <= 2 ? y - 1 : y).toString();
}

// 현재 학년도부터 과거로 count개 연도를 문자열 배열로 반환 (기본 3개)
function getRecentAcademicYears(count) {
  count = count || 3;
  const current = parseInt(getCurrentAcademicYear(), 10);
  const years = [];
  for (let i = 0; i < count; i++) years.push((current - i).toString());
  return years;
}

function parseNum(val) {
  if (val === undefined || val === null || val === "") return 0;
  const cleaned = val.toString().replace(/[^\d.]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// 시트 컬럼을 "@"(텍스트) 서식으로 지정해도, appendRow로 날짜처럼 생긴 문자열을 쓰면
// 스프레드시트가 이를 실제 Date 객체로 자동 변환해버리는 경우가 있다. Date 객체가
// google.script.run 응답에 섞여 있으면 클라이언트에서 직렬화가 깨져 성공 핸들러에
// null이 전달되는 문제가 생길 수 있으므로, 시트에서 읽어온 값은 항상 이 함수로
// 순수 문자열로 되돌려서 반환한다.
function toDateStr_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return val.toString();
}

// <input type="datetime-local"> 값(신청시작/신청마감/발표시각 등)용. 마찬가지로
// Date 객체로 자동 변환됐을 경우 "yyyy-MM-ddTHH:mm" 형태의 문자열로 되돌린다.
function toDateTimeStr_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  return val.toString();
}

// google.script.run 응답에 시트에서 읽은 Date 객체가 하나라도 섞여 있으면 클라이언트
// 쪽 성공 핸들러가 null을 받는 문제가 생길 수 있다. 개별 컬럼마다 일일이
// toDateStr_/toDateTimeStr_를 적용하는 걸 놓치기 쉬우므로, 시트 데이터를 반환하는
// 함수들은 이 함수로 반환값 전체를 한 번 더 훑어서 남아있는 Date 객체를 전부
// 문자열로 안전하게 바꾼 뒤 return한다(마지막 안전망).
function sanitizeDates_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  if (Array.isArray(value)) return value.map(sanitizeDates_);
  if (value && typeof value === 'object') {
    const out = {};
    for (const k in value) { if (Object.prototype.hasOwnProperty.call(value, k)) out[k] = sanitizeDates_(value[k]); }
    return out;
  }
  return value;
}

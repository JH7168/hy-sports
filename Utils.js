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

/* =====================================================================
 *  csv-utils.js  —  단일 HTML 도구용 헬퍼 스니펫 모음
 *
 *  사용법: 새 도구 HTML의 <script> 안에 필요한 함수를 그대로 복붙.
 *  외부 의존성: 없음 (CSV 파싱이 필요하면 PapaParse CDN 별도 로드)
 *
 *  추출 출처:
 *   - sumin.github.io/report-generator.html (parseDate, readFileAsRows)
 *   - 수강데이터가공기/index.html (decode/HMS/parseNumberCell/escapeHtml/download)
 * ===================================================================== */


/* ---------------------------------------------------------------------
 *  1) 파일 인코딩 디코딩  —  UTF-8 우선, 실패 시 EUC-KR(CP949) 폴백
 *
 *  엑셀에서 저장한 CSV는 EUC-KR로 떨어지는 경우가 많아서 둘 다 처리해야 함.
 *  fatal:true 로 한 번 시도 → 깨지면 EUC-KR.
 *
 *  사용:
 *    const reader = new FileReader();
 *    reader.onload = () => {
 *      const text = decodeBuffer(reader.result);
 *      Papa.parse(text, { header: true, complete: (p) => ... });
 *    };
 *    reader.readAsArrayBuffer(file);
 * --------------------------------------------------------------------- */
function decodeBuffer(buf) {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buf);
  } catch {
    try {
      return new TextDecoder("euc-kr").decode(buf);
    } catch {
      return new TextDecoder("utf-8").decode(buf);
    }
  }
}


/* ---------------------------------------------------------------------
 *  2) 초 → "HH:MM:SS"
 *
 *  엑셀 시간 셀처럼 반올림. 음수/NaN은 0으로 보정.
 * --------------------------------------------------------------------- */
function secondsToHMS(seconds) {
  if (!Number.isFinite(seconds) || seconds < 0) seconds = 0;
  const s = Math.round(seconds);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(h)}:${pad(m)}:${pad(sec)}`;
}


/* ---------------------------------------------------------------------
 *  3) 셀 값 → 숫자
 *
 *  엑셀에서 "1,234" 같은 콤마/공백 포함 문자열, 빈 셀, "-" 모두 안전 처리.
 * --------------------------------------------------------------------- */
function parseNumberCell(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  if (s === "" || s === "-") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}


/* ---------------------------------------------------------------------
 *  4) HTML 이스케이프  —  innerHTML로 사용자 입력 넣을 때 항상 거쳐 갈 것
 * --------------------------------------------------------------------- */
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[c]);
}


/* ---------------------------------------------------------------------
 *  5) 날짜 값 통합 파싱
 *
 *  처리 가능한 입력:
 *   - JS Date 객체 (XLSX cellDates:true 옵션)
 *   - Excel 시리얼 넘버 (40000~60000 구간 = 대략 2009~2064)
 *   - 텍스트 "2026. 4. 29.", "2026-04-29", "2026/4/29" 등
 *
 *  반환: Date 객체 또는 null
 * --------------------------------------------------------------------- */
function parseDate(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  if (typeof val === "number" && val > 40000 && val < 60000) {
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  const m = String(val).match(/(\d{4})[.\-\/\s]+(\d{1,2})[.\-\/\s]+(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}


/* ---------------------------------------------------------------------
 *  6) CSV 다운로드  —  엑셀에서 한글 안 깨지게 UTF-8 BOM 부착
 *
 *  rows: 2차원 배열 ([[헤더...], [값...], ...])
 *  filename: 다운로드 파일명 (확장자 .csv 포함)
 *
 *  의존성: PapaParse (CDN으로 별도 로드)
 *    <script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"></script>
 * --------------------------------------------------------------------- */
function downloadCsvWithBom(rows, filename) {
  const csv = Papa.unparse(rows);
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}


/* ---------------------------------------------------------------------
 *  7) 오늘 날짜  —  파일명에 붙이기 좋은 "YYYY-MM-DD"
 * --------------------------------------------------------------------- */
function todayYmd() {
  return new Date().toISOString().slice(0, 10);
}


/* ---------------------------------------------------------------------
 *  8) 엑셀 → rows  —  XLSX(.xlsx/.xls) 와 CSV 둘 다 처리
 *
 *  의존성: XLSX (SheetJS) CDN
 *    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
 *
 *  cellDates:true 로 날짜 셀은 Date 객체로 받음.
 *  cb: (rows) => void  —  각 행은 헤더명을 키로 갖는 객체
 * --------------------------------------------------------------------- */
function readFileAsRows(file, cb) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    cb(XLSX.utils.sheet_to_json(ws, { defval: "" }));
  };
  reader.readAsArrayBuffer(file);
}

'use strict';

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// ───────────── 계정 설정 ─────────────
const USERS = {
  admin: process.env.ADMIN_PASSWORD || 'kpi1234',
};

// ───────────── 출발지/도착지 기준 데이터 로드 ─────────────
const FROM_FILE = path.join(__dirname, 'FROM_LIST.xlsx');
const TO_FILE   = path.join(__dirname, 'TO_LIST.xlsx');
const fromMap = {}; // 출발지코드 → 출발지건물
const toMap   = {}; // 도착지코드 → { 건물, 검사내용 }

(function loadReferenceData() {
  // FROM_LIST: col0=출발지코드, col2=출발지건물
  try {
    const wb = XLSX.readFile(FROM_FILE);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (r[0]) fromMap[String(r[0]).trim()] = r[2] ? String(r[2]).trim() : null;
    }
    console.log(`FROM_LIST 로드 완료: 출발지 ${Object.keys(fromMap).length}개`);
  } catch (e) {
    console.warn('FROM_LIST 로드 실패:', e.message);
  }

  // TO_LIST: col0=도착지코드, col2=도착지건물, col3=검사내용
  try {
    const wb = XLSX.readFile(TO_FILE);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (r[0]) toMap[String(r[0]).trim()] = {
        건물: r[2] ? String(r[2]).trim() : null,
        검사내용: r[3] ? String(r[3]).trim() : null,
      };
    }
    console.log(`TO_LIST 로드 완료: 도착지 ${Object.keys(toMap).length}개`);
  } catch (e) {
    console.warn('TO_LIST 로드 실패:', e.message);
  }
})();

// ───────────── 한국 공휴일 ─────────────
const KOREAN_HOLIDAYS = new Set([
  // 2025년
  '2025-01-01', '2025-01-28', '2025-01-29', '2025-01-30',
  '2025-03-01', '2025-05-05', '2025-05-06', '2025-06-06',
  '2025-08-15', '2025-10-03', '2025-10-05', '2025-10-06', '2025-10-07',
  '2025-10-09', '2025-12-25',
  // 2026년
  '2026-01-01',                               // 신정
  '2026-02-16', '2026-02-17', '2026-02-18',  // 설날 연휴
  '2026-03-01',                               // 삼일절
  '2026-05-05',                               // 어린이날
  '2026-05-24',                               // 부처님오신날
  '2026-06-06',                               // 현충일
  '2026-08-15',                               // 광복절
  '2026-09-24', '2026-09-25', '2026-09-26',  // 추석 연휴
  '2026-10-03',                               // 개천절
  '2026-10-09',                               // 한글날
  '2026-12-25',                               // 성탄절
]);

// ───────────── 중앙값 계산 ─────────────
function median(arr) {
  if (!arr || !arr.length) return null;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const val = sorted.length % 2 === 0
    ? (sorted[mid - 1] + sorted[mid]) / 2
    : sorted[mid];
  return +val.toFixed(1);
}

// ───────────── 평일 여부 (토·일·공휴일 제외) ─────────────
function isWorkingDay(dateStr) {
  if (!dateStr) return false;
  if (KOREAN_HOLIDAYS.has(dateStr)) return false;
  const [y, m, d] = dateStr.split('-').map(Number);
  const dow = new Date(y, m - 1, d).getDay(); // 0=일, 6=토
  return dow !== 0 && dow !== 6;
}

// 데이터 내 평일 근무일 수 (날짜 집합 기준)
function getWorkingDays(data) {
  const days = new Set(data.map(r => r.호출일).filter(isWorkingDay));
  return days.size || 1;
}

// ───────────── 업로드된 데이터 저장소 ─────────────
let cachedData = null;

// ───────────── 미들웨어 ─────────────
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: false, limit: '100mb' }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'kpi-secret-key-2025',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 },
}));

const upload = multer({ dest: path.join(__dirname, 'uploads') });

// ───────────── 인증 미들웨어 ─────────────
function requireAuth(req, res, next) {
  if (req.session && req.session.user) return next();
  res.redirect('/');
}

app.use(express.static(path.join(__dirname, 'public')));

// ───────────── 인증 라우트 ─────────────
app.post('/login', (req, res) => {
  const { username, password } = req.body;
  if (USERS[username] && USERS[username] === password) {
    req.session.user = username;
    res.json({ ok: true });
  } else {
    res.json({ ok: false, message: '아이디 또는 비밀번호가 틀렸습니다.' });
  }
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

app.get('/api/me', (req, res) => {
  if (req.session && req.session.user) {
    res.json({ ok: true, user: req.session.user });
  } else {
    res.json({ ok: false });
  }
});

// ───────────── 시간 파싱 헬퍼 ─────────────
function parseTimeToMinutes(t) {
  if (!t || typeof t !== 'string') return null;
  const parts = t.trim().split(':');
  if (parts.length < 2) return null;
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function diffMinutes(startStr, endStr) {
  const s = parseTimeToMinutes(startStr);
  const e = parseTimeToMinutes(endStr);
  if (s === null || e === null) return null;
  let diff = e - s;
  if (diff < -300) diff += 24 * 60; // 자정 넘김 보정
  if (diff < 0 || diff > 600) return null;
  return diff;
}

// ───────────── 이동수단 매핑 ─────────────
const TRANSPORT_MAP = { A: '도보', B: '휠체어', C: '침대', D: '이동카', Z: '도보' };
function getTransport(code) {
  if (!code) return '도보';
  return TRANSPORT_MAP[String(code).trim().toUpperCase()] || '도보';
}

// ───────────── JOBTP 매핑 ─────────────
const JOBTP_MAP = { D: '새벽', E: '긴급', R: '정규', RE: '예약호출' };
function getJobType(code) {
  if (!code) return '-';
  const c = String(code).trim().toUpperCase();
  return JOBTP_MAP[c] || c;
}

// ───────────── 행 처리 ─────────────
function processRow(r) {
  // 개인정보 마스킹
  const objno = r['objno'] ? String(r['objno']).slice(-4) : '';
  const objnm = r['objnm'] ? String(r['objnm']).slice(-1) : '';

  // 출발지 건물 조회
  const fromDetailCode = r['출발지세부코드'] ? String(r['출발지세부코드']).trim() : '';
  const fromCode       = r['출발지코드']   ? String(r['출발지코드']).trim()   : '';
  const fromBuilding   = fromMap[fromDetailCode] || fromMap[fromCode] || null;

  // 도착지 건물 조회
  const toCode     = r['도착지코드'] ? String(r['도착지코드']).trim() : '';
  const toEntry    = toMap[toCode] || {};
  const toBuilding = toEntry.건물 || null;
  const 검사내용   = toEntry.검사내용 || '';  // 도착지코드 → distance 파일만 사용

  // 장거리 여부
  const isLong = fromBuilding && toBuilding &&
    ['본', '별관', '암', '양성자'].includes(fromBuilding) &&
    ['본', '별관', '암', '양성자'].includes(toBuilding) &&
    fromBuilding !== toBuilding;

  // 방사선검사실 제외 여부 (출발지 또는 도착지에 포함 시 KPI 산출 제외)
  const 출발지명 = r['출발지'] || '';
  const 도착지명 = r['도착지'] || '';
  const 제외여부 = 출발지명.includes('방사선검사실') || 도착지명.includes('방사선검사실');

  // 이송시간 KPI (통합 중앙값 산출 기준)
  // 예약시간 있으면 예약→시작, 없으면 호출→시작
  const hasReserve = !!(r['예약시간'] && String(r['예약시간']).trim());
  const 이송kpi = hasReserve
    ? diffMinutes(r['예약시간'], r['시작시간'])
    : diffMinutes(r['호출시간'], r['시작시간']);

  // 호출 시간대
  const callHour = r['호출시간']
    ? parseInt(String(r['호출시간']).split(':')[0], 10)
    : null;

  // 동시누름: 시작시간~종료시간 3분 이내
  const 시작종료차 = diffMinutes(r['시작시간'], r['종료시간']);
  const 동시건 = 시작종료차 !== null && 시작종료차 <= 3;

  return {
    호출일:       r['호출일'] || '',
    환자번호:     objno,
    환자명:       objnm + '○',
    출발지:       출발지명,
    출발지건물:   fromBuilding || '',
    도착지:       도착지명,
    도착지건물:   toBuilding || '',
    업무내용:     r['업무내용'] || '',
    검사내용,
    장거리:       isLong,
    이동수단:     getTransport(r['이동수단코드']),
    호출유형:     getJobType(r['JOBTP']),
    jobtpCode:    String(r['JOBTP'] || '').trim().toUpperCase(),
    담당자:       r['AGENTNM'] || '미배정',
    현재상태:     r['현재상태코드'] || '',
    호출시간:     r['호출시간'] || '',
    시작시간:     r['시작시간'] || '',
    종료시간:     r['종료시간'] || '',
    예약시간:     r['예약시간'] || '',
    제외여부,
    이송kpi,
    호출시간대:   callHour,
    지연시간:     r['지연시간'] || 0,
    동시건,
  };
}

// ───────────── 파일 업로드 ─────────────
app.post('/upload', requireAuth, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, message: '파일이 없습니다.' });

  try {
    const wb = XLSX.readFile(req.file.path, {
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false,
      sheetStubs: false,
    });
    const sheetName = wb.SheetNames.includes('워크시트 익스포트')
      ? '워크시트 익스포트'
      : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { raw: true, defval: null });
    wb.Sheets = {};

    cachedData = rows.map(processRow);
    fs.unlink(req.file.path, () => {});

    res.json({
      ok: true,
      count: cachedData.length,
      dateRange: getDateRange(cachedData),
    });
  } catch (e) {
    fs.unlink(req.file.path, () => {});
    res.status(500).json({ ok: false, message: '파일 처리 오류: ' + e.message });
  }
});

// ───────────── JSON 업로드 (브라우저에서 파싱된 데이터 수신) ─────────────
app.post('/upload-json', requireAuth, (req, res) => {
  try {
    const { rows } = req.body;
    if (!rows || !Array.isArray(rows)) return res.status(400).json({ ok: false, message: '데이터가 없습니다.' });

    cachedData = rows.map(processRow);

    res.json({
      ok: true,
      count: cachedData.length,
      dateRange: getDateRange(cachedData),
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: '처리 오류: ' + e.message });
  }
});

// ───────────── 날짜 범위 ─────────────
function getDateRange(data) {
  const dates = data.map(r => r.호출일).filter(Boolean).sort();
  return dates.length ? { start: dates[0], end: dates[dates.length - 1] } : null;
}

// ───────────── 필터 적용 ─────────────
function applyFilter(data, query) {
  let d = data;
  if (query.dateFrom) d = d.filter(r => r.호출일 >= query.dateFrom);
  if (query.dateTo)   d = d.filter(r => r.호출일 <= query.dateTo);
  if (query.agent)    d = d.filter(r => r.담당자 === query.agent);
  return d;
}

// ───────────── 통계 API ─────────────
app.get('/api/stats', requireAuth, (req, res) => {
  if (!cachedData) return res.json({ ok: false, message: '업로드된 데이터가 없습니다.' });

  const data = applyFilter(cachedData, req.query);

  res.json({
    ok: true,
    count: data.length,
    dateRange: getDateRange(data),
    individual: getIndividualStats(data),
    hourly: getHourlyStats(data),
    timing: getTimingStats(data),
    agents: [...new Set(cachedData.map(r => r.담당자).filter(Boolean))].sort(),
    depts:  [...new Set(cachedData.map(r => r.출발지).filter(Boolean))].sort(),
  });
});

// ① 개인별 실적
function getIndividualStats(data) {
  const map = {};

  for (const r of data) {
    const name = r.담당자;
    if (!map[name]) map[name] = {
      담당자: name, 총건수: 0, 평일건수: 0, 평일날짜: new Set(), 장거리: 0,
      도보: 0, 휠체어: 0, 침대: 0, 이동카: 0,
      kpiArr: [], 업무: {}, 동시건: 0,
    };
    const s = map[name];
    s.총건수++;
    if (isWorkingDay(r.호출일)) {
      s.평일건수++;
      s.평일날짜.add(r.호출일);
    }
    if (r.동시건) s.동시건++;
    if (r.장거리) s.장거리++;
    if (r.이동수단 === '도보')   s.도보++;
    else if (r.이동수단 === '휠체어') s.휠체어++;
    else if (r.이동수단 === '침대')   s.침대++;
    else if (r.이동수단 === '이동카') s.이동카++;
    // 이송시간 KPI: 방사선검사실 제외
    if (!r.제외여부 && r.이송kpi !== null) s.kpiArr.push(r.이송kpi);
    const job = r.업무내용 || '기타';
    s.업무[job] = (s.업무[job] || 0) + 1;
  }

  return Object.values(map)
    .map(s => ({
      담당자:       s.담당자,
      총건수:       s.총건수,
      장거리:       s.장거리,
      도보:         s.도보,
      휠체어:       s.휠체어,
      침대:         s.침대,
      이동카:       s.이동카,
      업무:         s.업무,
      중간이송시간: median(s.kpiArr),
      일평균:       s.평일날짜.size > 0 ? +(s.평일건수 / s.평일날짜.size).toFixed(1) : 0,
      근무일수:     s.평일날짜.size,
      동시건:       s.동시건,
    }))
    .sort((a, b) => b.총건수 - a.총건수);
}

// ② 시간대별 업무량 (0~23시 전체, 차트/테이블 표시는 프론트에서 필터)
function getHourlyStats(data) {
  const hours = Array.from({ length: 24 }, (_, i) => ({
    시간대: i, 총건수: 0, 장거리: 0,
    도보: 0, 휠체어: 0, 침대: 0, 이동카: 0,
    긴급: 0, 예약: 0, 새벽엑스레이: 0, 정규엑스레이: 0,
  }));
  for (const r of data) {
    const h = r.호출시간대;
    if (h === null || h < 0 || h > 23) continue;
    hours[h].총건수++;
    if (r.장거리) hours[h].장거리++;
    if (r.이동수단 === '도보')   hours[h].도보++;
    else if (r.이동수단 === '휠체어') hours[h].휠체어++;
    else if (r.이동수단 === '침대')   hours[h].침대++;
    else if (r.이동수단 === '이동카') hours[h].이동카++;
    if (r.jobtpCode === 'E')  hours[h].긴급++;
    if (r.jobtpCode === 'RE') hours[h].예약++;
    if (r.jobtpCode === 'D')  hours[h].새벽엑스레이++;
    if (r.jobtpCode === 'R')  hours[h].정규엑스레이++;
  }
  return hours;
}

// ③ 이송시간 분석
function getTimingStats(data) {
  // 방사선검사실 제외 + 유효값만
  const kpiData = data.filter(r => !r.제외여부 && r.이송kpi !== null);

  // 전체 중앙값
  const allKpi = kpiData.map(r => r.이송kpi);

  // 시간대별 중앙값 (6~22시)
  const hourBuckets = {};
  for (let h = 6; h <= 22; h++) hourBuckets[h] = [];
  for (const r of kpiData) {
    const h = r.호출시간대;
    if (h !== null && h >= 6 && h <= 22) hourBuckets[h].push(r.이송kpi);
  }
  const hourlyTiming = Object.entries(hourBuckets).map(([h, arr]) => ({
    시간대: +h,
    중간이송시간: median(arr),
    건수: arr.length,
  }));

  // 담당자별: 평일 데이터만 사용, 중앙값 + 평일 일평균
  // 평일 전체 건수 + 실제 근무일 추적 (방사선 포함, 일평균 분모용)
  const agentTotalMap = {};
  const agentDaysMap = {}; // 담당자별 실제 근무일 Set
  for (const r of data) {
    if (!isWorkingDay(r.호출일)) continue;
    agentTotalMap[r.담당자] = (agentTotalMap[r.담당자] || 0) + 1;
    if (!agentDaysMap[r.담당자]) agentDaysMap[r.담당자] = new Set();
    agentDaysMap[r.담당자].add(r.호출일);
  }

  // 평일 KPI 배열 (방사선 제외, 이송시간 중앙값용)
  const agentKpiMap = {};
  for (const r of kpiData) {
    if (!isWorkingDay(r.호출일)) continue;
    if (!agentKpiMap[r.담당자]) agentKpiMap[r.담당자] = [];
    agentKpiMap[r.담당자].push(r.이송kpi);
  }

  const agentNames = new Set([
    ...Object.keys(agentTotalMap),
    ...Object.keys(agentKpiMap),
  ]);

  const agentTiming = [...agentNames].map(name => {
    const 평일건수 = agentTotalMap[name] || 0;
    const 근무일수 = agentDaysMap[name] ? agentDaysMap[name].size : 0;
    return {
      담당자:       name,
      중간이송시간: median(agentKpiMap[name] || []),
      평일건수,
      일평균:       근무일수 > 0 ? +(평일건수 / 근무일수).toFixed(1) : 0,
      근무일수,
    };
  }).sort((a, b) => b.평일건수 - a.평일건수);

  return {
    전체: {
      중간이송시간: median(allKpi),
      이송건수:     allKpi.length,
    },
    개인별: agentTiming,
    시간대별: hourlyTiming,
  };
}

// ───────────── 업무구분 매핑 ─────────────
const CATEGORY_ORDER = [
  'CT검사','MRI','감마나이프','검사','기관지내시경','기타',
  '내시경','방사선','방종','병동복귀','수술','심전도',
  '심혈관조영','외래','응급실','재활','중환자','초음파',
  '통원치료실','투석','투시조영','폐기능','핵의학','혈관조영','회복실',
];

const VALID_CATEGORY = new Set(CATEGORY_ORDER);

// 업무구분 3단계 판별
// 1단계: JOBTP 'D' 또는 'R' → 방사선
// 2단계: 업무내용='기타' → 기타
// 3단계: 도착지코드 → TO_LIST 검사내용 (없으면 기타)
function get업무구분(r) {
  if (r.jobtpCode === 'D' || r.jobtpCode === 'R') return '방사선';
  if ((r.업무내용 || '').trim() === '기타') return '기타';
  const 검사 = r.검사내용 ? String(r.검사내용).trim() : '';
  if (검사 && VALID_CATEGORY.has(검사)) return 검사;
  return '기타';
}

const DOW_KR = ['일','월','화','수','목','금','토'];

// ───────────── 월간 보고서 생성 ─────────────
app.get('/api/report', requireAuth, (req, res) => {
  if (!cachedData) return res.status(400).json({ ok: false, message: '업로드된 데이터가 없습니다.' });

  try {
    const data = applyFilter(cachedData, req.query);
    const wb = buildReportWorkbook(data);
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    const dateRange = getDateRange(data);
    const fileName = `이송KPI_${dateRange?.start?.slice(0,7) || 'report'}.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) {
    res.status(500).json({ ok: false, message: '보고서 생성 오류: ' + e.message });
  }
});

function buildReportWorkbook(data) {
  const wb = XLSX.utils.book_new();

  // 데이터를 월별로 그룹화
  const monthGroups = {};
  for (const r of data) {
    const ym = r.호출일 ? r.호출일.slice(0, 7) : null; // "2026-01"
    if (!ym) continue;
    if (!monthGroups[ym]) monthGroups[ym] = [];
    monthGroups[ym].push(r);
  }

  const months = Object.keys(monthGroups).sort();

  // ── Sheet1 ──
  const aoa = []; // array of arrays

  // ① 월별 요약 (sample의 B1:E5 구조)
  aoa.push([]); // 빈 행
  aoa.push(['월', '업무량(일평균)', '초과건(일평균)', '초과율']);

  for (const ym of months) {
    const rows = monthGroups[ym];
    // 평일 데이터가 없는 월은 요약에서도 제외
    const hasWorkday = rows.some(r => r.호출일 && isWorkingDay(r.호출일));
    if (!hasWorkday) continue;
    const wdays = getWorkingDays(rows);
    const total = rows.filter(r => r.호출일 && isWorkingDay(r.호출일)).length;
    const over30 = rows.filter(r => r.호출일 && isWorkingDay(r.호출일) && !r.제외여부 && r.이송kpi !== null && r.이송kpi > 30).length;
    const label = ym.slice(5, 7).replace(/^0/, '') + '월';
    aoa.push([
      label,
      wdays ? +(total / wdays).toFixed(1) : total,
      wdays ? +(over30 / wdays).toFixed(1) : over30,
      total ? Math.round(over30 / total * 100) + '%' : '0%',
    ]);
  }

  // ② 월별 상세 (month별 반복, 평일 데이터 있는 월만)
  for (const ym of months) {
    const rows = monthGroups[ym];

    // 해당 월의 평일 날짜 목록 추출 (데이터 기준)
    const wdaySet = new Set(
      rows.map(r => r.호출일).filter(d => d && isWorkingDay(d))
    );
    const wdays = [...wdaySet].sort(); // "2026-01-02", ...

    // 평일 데이터가 없는 월은 건너뜀
    if (wdays.length === 0) continue;

    const yearShort = ym.slice(2, 4); // "26"
    const monthNum  = ym.slice(5, 7).replace(/^0/, ''); // "1"

    aoa.push([]); // 빈 행

    // 제목 행
    const titleRow = [`${yearShort}년 ${monthNum}월 월간 업무보고`];
    // 날짜 수만큼 빈 칸 + "평일" 레이블
    for (let i = 1; i < wdays.length + 1; i++) titleRow.push(null);
    titleRow.push(null, '평일', wdays.length);
    aoa.push(titleRow);

    // 헤더 행: 업무구분, 일자, [날짜들], 합계, 평균, 비중
    aoa.push(['업무구분', '일자', ...wdays.map(d => +d.slice(8)), '합계', '평균', '비중']);

    // 요일 행
    aoa.push([null, '요일', ...wdays.map(d => {
      const [y, m, dd] = d.split('-').map(Number);
      return DOW_KR[new Date(y, m - 1, dd).getDay()];
    }), null, null, null]);

    // 카테고리별 일별 집계
    const catDayCount = {}; // { cat: { date: count } }
    for (const cat of CATEGORY_ORDER) catDayCount[cat] = {};

    for (const r of rows) {
      if (!r.호출일 || !isWorkingDay(r.호출일)) continue;
      const cat = get업무구분(r);
      const finalCat = VALID_CATEGORY.has(cat) ? cat : '기타';
      if (!catDayCount[finalCat][r.호출일]) catDayCount[finalCat][r.호출일] = 0;
      catDayCount[finalCat][r.호출일]++;
    }

    // 총합 계산 (비중 분모)
    let grandTotal = 0;
    for (const cat of CATEGORY_ORDER) {
      grandTotal += wdays.reduce((s, d) => s + (catDayCount[cat][d] || 0), 0);
    }

    // 카테고리 행 출력
    let sumGrandTotal = 0;
    for (const cat of CATEGORY_ORDER) {
      const dayCounts = wdays.map(d => catDayCount[cat][d] || 0);
      const total = dayCounts.reduce((a, b) => a + b, 0);
      const avg = wdays.length ? +(total / wdays.length).toFixed(1) : 0;
      const pct = grandTotal ? Math.round(total / grandTotal * 100) + '%' : '0%';
      aoa.push([cat, null, ...dayCounts, total, avg, pct]);
      sumGrandTotal += total;
    }

    // 총합계 행
    const totalDaySums = wdays.map(d =>
      CATEGORY_ORDER.reduce((s, cat) => s + (catDayCount[cat][d] || 0), 0)
    );
    const grandAvg = wdays.length ? +(sumGrandTotal / wdays.length).toFixed(1) : 0;
    aoa.push(['총합계', null, ...totalDaySums, sumGrandTotal, grandAvg, '100%']);

    // ③ 30분 초과 지연 상세
    aoa.push([]); // 빈 행
    aoa.push([`${monthNum}월 30분 초과 지연 업무량`]);
    aoa.push(['업무구분', '일자', ...wdays.map(d => +d.slice(8)), '총합계', '평균', '비중']);
    aoa.push([null, '요일', ...wdays.map(d => {
      const [y, m, dd] = d.split('-').map(Number);
      return DOW_KR[new Date(y, m - 1, dd).getDay()];
    }), null, null, null]);

    // 지연(이송kpi > 30) 카테고리별 집계
    const delayDayCount = {};
    for (const cat of CATEGORY_ORDER) delayDayCount[cat] = {};

    for (const r of rows) {
      if (!r.호출일 || !isWorkingDay(r.호출일)) continue;
      if (r.제외여부) continue;
      if (r.이송kpi === null || r.이송kpi <= 30) continue;
      const cat = get업무구분(r);
      const finalCat = VALID_CATEGORY.has(cat) ? cat : '기타';
      if (!delayDayCount[finalCat][r.호출일]) delayDayCount[finalCat][r.호출일] = 0;
      delayDayCount[finalCat][r.호출일]++;
    }

    let delayGrandTotal = 0;
    for (const cat of CATEGORY_ORDER) {
      delayGrandTotal += wdays.reduce((s, d) => s + (delayDayCount[cat][d] || 0), 0);
    }

    let sumDelayTotal = 0;
    for (const cat of CATEGORY_ORDER) {
      const dayCounts = wdays.map(d => delayDayCount[cat][d] || 0);
      const total = dayCounts.reduce((a, b) => a + b, 0);
      const avg = wdays.length ? +(total / wdays.length).toFixed(1) : 0;
      const pct = delayGrandTotal ? Math.round(total / delayGrandTotal * 100) + '%' : '0%';
      aoa.push([cat, null, ...dayCounts, total, avg, pct]);
      sumDelayTotal += total;
    }

    // 총합계
    const delayDaySums = wdays.map(d =>
      CATEGORY_ORDER.reduce((s, cat) => s + (delayDayCount[cat][d] || 0), 0)
    );
    const delayGrandAvg = wdays.length ? +(sumDelayTotal / wdays.length).toFixed(1) : 0;
    aoa.push(['총합계', null, ...delayDaySums, sumDelayTotal, delayGrandAvg, sumDelayTotal ? Math.round(sumDelayTotal / sumGrandTotal * 100) + '%' : '0%']);

    // 지연율 행
    const delayRateRow = ['지연율', null];
    for (const d of wdays) {
      const dayTotal = CATEGORY_ORDER.reduce((s, cat) => s + (catDayCount[cat][d] || 0), 0);
      const dayDelay = CATEGORY_ORDER.reduce((s, cat) => s + (delayDayCount[cat][d] || 0), 0);
      delayRateRow.push(dayTotal ? Math.round(dayDelay / dayTotal * 100) + '%' : '0%');
    }
    const overallRate = sumGrandTotal ? Math.round(sumDelayTotal / sumGrandTotal * 100) + '%' : '0%';
    delayRateRow.push(overallRate, overallRate, overallRate);
    aoa.push(delayRateRow);

    aoa.push(['이송시간 30분 초과 기준 (방사선검사실 제외 동일 기준 아님)', null, ...wdays.map(() => null), null, null, null]);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, '월간보고');

  // ── 지연상세 시트: 30분 초과 건 전체 목록 ──
  const detailHeader = [
    '호출일', '요일', '담당자', '업무구분',
    '출발지', '도착지',
    '호출시간', '예약시간', '시작시간', '종료시간',
    '이송KPI(분)', '이동수단', '이송유형', '환자번호', '환자명',
  ];
  const detailRows = [detailHeader];

  const delayAll = data
    .filter(r => r.호출일 && isWorkingDay(r.호출일) && !r.제외여부 && r.이송kpi !== null && r.이송kpi > 30)
    .sort((a, b) => {
      const catA = get업무구분(a), catB = get업무구분(b);
      if (catA !== catB) return catA.localeCompare(catB);
      if (a.호출일 !== b.호출일) return a.호출일.localeCompare(b.호출일);
      return (a.호출시간 || '').localeCompare(b.호출시간 || '');
    });

  for (const r of delayAll) {
    const [y, m, d] = r.호출일.split('-').map(Number);
    const dow = DOW_KR[new Date(y, m - 1, d).getDay()];
    detailRows.push([
      r.호출일, dow, r.담당자, get업무구분(r),
      r.출발지, r.도착지,
      r.호출시간, r.예약시간, r.시작시간, r.종료시간,
      r.이송kpi, r.이동수단, r.호출유형, r.환자번호, r.환자명,
    ]);
  }

  const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
  XLSX.utils.book_append_sheet(wb, wsDetail, '지연상세');

  // ── 개인별실적 시트 ──
  const indivStats = getIndividualStats(data);
  const indivHeader = [
    '순위', '담당자', '총건수', '평일 일평균', '장거리',
    '도보', '휠체어', '침대', '이동카', '동시누름', '이송시간 중앙값(분)',
  ];
  const indivRows = [indivHeader, ...indivStats.map((r, i) => [
    i + 1, r.담당자, r.총건수, r.일평균, r.장거리,
    r.도보, r.휠체어, r.침대, r.이동카, r.동시건, r.중간이송시간,
  ])];
  const wsIndiv = XLSX.utils.aoa_to_sheet(indivRows);
  XLSX.utils.book_append_sheet(wb, wsIndiv, '개인별실적');

  return wb;
}

// ───────────── 담당자 일자별 실적 ─────────────
app.get('/api/agent-daily', requireAuth, (req, res) => {
  if (!cachedData) return res.json({ ok: false, message: '업로드된 데이터가 없습니다.' });

  const { agent, dateFrom, dateTo } = req.query;
  if (!agent) return res.json({ ok: false, message: '담당자명을 입력하세요.' });

  let data = cachedData.filter(r => r.담당자 === agent);
  if (dateFrom) data = data.filter(r => r.호출일 >= dateFrom);
  if (dateTo)   data = data.filter(r => r.호출일 <= dateTo);

  const dayMap = {};
  for (const r of data) {
    const d = r.호출일;
    if (!d) continue;
    if (!dayMap[d]) dayMap[d] = { 날짜: d, 총건수: 0, 장거리: 0, 도보: 0, 휠체어: 0, 침대: 0, 이동카: 0 };
    const s = dayMap[d];
    s.총건수++;
    if (r.장거리) s.장거리++;
    if (r.이동수단 === '도보')   s.도보++;
    else if (r.이동수단 === '휠체어') s.휠체어++;
    else if (r.이동수단 === '침대')   s.침대++;
    else if (r.이동수단 === '이동카') s.이동카++;
  }

  const days = Object.values(dayMap).sort((a, b) => a.날짜.localeCompare(b.날짜));
  res.json({ ok: true, agent, days });
});

// ───────────── 부서별 이송현황 ─────────────
app.get('/api/dept-stats', requireAuth, (req, res) => {
  if (!cachedData) return res.json({ ok: false, message: '데이터가 없습니다.' });

  const { dept, dateFrom, dateTo } = req.query;
  if (!dept) return res.json({ ok: false, message: '부서명을 입력하세요.' });

  let data = cachedData.filter(r => r.출발지 && r.출발지.includes(dept));
  if (dateFrom) data = data.filter(r => r.호출일 >= dateFrom);
  if (dateTo)   data = data.filter(r => r.호출일 <= dateTo);

  const total = data.length;
  const over30 = data.filter(r => !r.제외여부 && r.이송kpi !== null && r.이송kpi > 30).length;
  const kpiArr = data.filter(r => !r.제외여부 && r.이송kpi !== null).map(r => r.이송kpi);

  const hours = Array.from({ length: 24 }, (_, i) => ({ 시간대: i, 총건수: 0, 초과: 0 }));
  for (const r of data) {
    const h = r.호출시간대;
    if (h === null || h < 0 || h > 23) continue;
    hours[h].총건수++;
    if (!r.제외여부 && r.이송kpi !== null && r.이송kpi > 30) hours[h].초과++;
  }

  res.json({ ok: true, dept, total, over30, 중간이송시간: median(kpiArr), hours });
});

// ───────────── 서버 시작 ─────────────
app.listen(PORT, () => {
  console.log(`KPI 서버 실행 중: http://localhost:${PORT}`);
});

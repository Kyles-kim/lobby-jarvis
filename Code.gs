/**
 * Lobby — 말로 기록하고, 말로 묻는 개인 비서
 * 백엔드: Google Apps Script  (설계서 v0.3 · 1차 MVP)
 *
 * [처음 한 번만]
 *  1) 프로젝트 설정 > 스크립트 속성에 추가
 *       ANTHROPIC_API_KEY = sk-ant-...
 *       APP_PIN           = 앱 접속 암호(예: 4~6자리 숫자)
 *  2) 편집기 상단 함수 목록에서 setup 선택 → 실행 (권한 승인)
 *  3) 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → URL 복사
 */

// ───────────────────────── 설정 ─────────────────────────
const CONFIG = {
  MODEL: 'claude-sonnet-4-5',   // 말로업무일지에서 쓰시는 모델명과 맞춰 주세요
  TZ: 'Asia/Seoul',
  TRASH_DAYS: 30,               // 삭제 후 복구 가능 기간(일)
  MAX_ROWS_TO_AI: 150,          // 답변 작성 시 Claude에게 넘길 최대 기록 수
};

const SH = {
  RECORD: '기록', PROFILE: '프로필', CATEGORY: '분류',
  LOG: '대화로그', HISTORY: '변경이력', TRASH: '휴지통',
};

const REC_HEADERS = ['ID', '입력일시', '대상일', '대상시각', '유형', '분류', '제목', '내용',
  '금액', '인물', '장소', '태그', '상태', '원문', '수정일시', '삭제일시'];
const COL = REC_HEADERS.reduce((o, h, i) => (o[h] = i, o), {});

const TYPES = ['일상', '지출', '일정', '업무일정', '정보', '특이점', '아이디어'];
const STATES = ['정상', '예정', '완료', '취소', '삭제'];
const EDITABLE = ['대상일', '대상시각', '유형', '분류', '제목', '내용', '금액', '인물', '장소', '태그', '상태'];

const DEFAULT_CATEGORIES = [
  ['차량', '정비·주유·보험·세차'],
  ['생활', '장보기·생활용품·공과금'],
  ['미용/자기관리', '이발·운동·자기계발'],
  ['외식/모임', '외식·친구·동창 모임'],
  ['가족행사', '생일·기념일·가족 모임'],
  ['여가/취미', '골프·영화·취미 활동'],
  ['여행', '국내외 여행·숙박'],
  ['업무', '회사 일정·거래처·생산'],
  ['재테크', '주식·저축·투자'],
  ['앱개발', '앱 아이디어·개발 작업'],
];

// ───────────────────────── 초기 설정 ─────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.TZ);

  ensureSheet_(ss, SH.RECORD, REC_HEADERS, ['A:D', 'O:P']);
  ensureSheet_(ss, SH.PROFILE, ['항목', '값', '갱신일'], ['C:C']);
  const cat = ensureSheet_(ss, SH.CATEGORY, ['분류명', '설명', '추가일'], ['C:C']);
  ensureSheet_(ss, SH.LOG, ['시각', '질문', '조회계획', '음성답변', '해당건수'], ['A:A']);
  ensureSheet_(ss, SH.HISTORY, ['시각', '기록ID', '칸', '이전값', '새값', '경로'], ['A:B']);
  ensureSheet_(ss, SH.TRASH, REC_HEADERS, ['A:D', 'O:P']);

  if (cat.getLastRow() < 2) {
    const today = fmt_(new Date(), 'yyyy-MM-dd');
    cat.getRange(2, 1, DEFAULT_CATEGORIES.length, 3)
      .setValues(DEFAULT_CATEGORIES.map(r => [r[0], r[1], today]));
  }

  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'purgeTrash');
  if (!has) ScriptApp.newTrigger('purgeTrash').timeBased().everyDays(1).atHour(4).create();

  const p = PropertiesService.getScriptProperties();
  const miss = ['ANTHROPIC_API_KEY', 'APP_PIN'].filter(k => !p.getProperty(k));
  Logger.log(miss.length ? '⚠ 스크립트 속성 누락: ' + miss.join(', ') : '✔ Lobby 설정 완료');
}

function ensureSheet_(ss, name, headers, textCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1F2530').setFontColor('#EDEFF3');
    sh.setFrozenRows(1);
  }
  (textCols || []).forEach(a1 => sh.getRange(a1).setNumberFormat('@'));
  return sh;
}

// ───────────────────────── 웹 요청 ─────────────────────────
function doGet() {
  return json_({ ok: true, app: 'Lobby', msg: 'Lobby 서버가 동작 중입니다.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const pin = PropertiesService.getScriptProperties().getProperty('APP_PIN');
    if (!pin || String(req.pin) !== String(pin)) return json_({ ok: false, error: '접속 암호가 맞지 않습니다.' });

    const handlers = {
      init: actInit_, parse: actParse_, save: actSave_, ask: actAsk_,
      list: actList_, update: actUpdate_, 'delete': actDelete_, restore: actRestore_,
    };
    const fn = handlers[req.action];
    if (!fn) return json_({ ok: false, error: '알 수 없는 요청: ' + req.action });
    return json_(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── init ─────────────────────────
function actInit_() {
  return { categories: getCategories_(), types: TYPES, states: STATES, today: fmt_(new Date(), 'yyyy-MM-dd'), recent: recent_(5) };
}

// ───────────────────────── 기록: 정리(parse) ─────────────────────────
function actParse_(req) {
  const text = mask_(String(req.text || '').trim());
  if (!text) throw new Error('말씀하신 내용이 비어 있습니다.');
  const cats = getCategories_();

  const system = [
    '너는 카일님의 개인 비서 Lobby다. 카일님의 음성 발화를 기록용 JSON으로 정리한다.',
    dateContext_(),
    '유형(7개 중 하나): ' + TYPES.join(', '),
    '  일상=있었던 일, 지출=돈을 쓴 기록(금액 필수), 일정=앞으로의 개인 약속, 업무일정=회사 약속·마감,',
    '  정보=기억해 둘 사실, 특이점=평소와 다른 일·주의할 일, 아이디어=떠오른 생각.',
    '분류(기존): ' + cats.map(c => c.name).join(', '),
    '규칙:',
    '- 한 발화에 여러 건이 있으면 각각 나눈다.',
    '- 대상일은 yyyy-MM-dd. "어제"=오늘-1, "그저께"=오늘-2, "다음 주 X요일"=다음 주(월~일)의 X요일, "이번 주말"=이번 주 토요일(내용에 "주말" 표기), "다음 달 초"=다음 달 1일(내용에 "초순" 표기).',
    '- 날짜 언급이 없으면 일상·지출·정보·특이점·아이디어는 오늘. 일정·업무일정인데 날짜가 없으면 대상일을 ""로 두고 확인필요에 질문을 적는다.',
    '- 대상시각은 HH:mm(24시간). 없으면 "".',
    '- 금액은 원 단위 정수. "8만5천"=85000. 없으면 null.',
    '- 제목은 20자 안팎 한 줄 요약, 내용은 정리된 본문(수치·조건 보존).',
    '- 상태: 일정·업무일정은 "예정", 나머지는 "정상".',
    '- 태그는 검색용 핵심어 1~4개.',
    '- 기존 분류에 맞지 않으면 새 분류명을 제안하고 새분류=true, 분류설명에 한 줄 설명. 억지로 새 분류를 만들지 말 것.',
    '- 유형이 정보이고 "항목=값"으로 정리 가능하면 프로필항목/프로필값을 채운다(예: "차 타이어 규격" / "235/55R19").',
    '- 카드번호·계좌번호·비밀번호는 절대 옮겨 적지 말고 "***"로 가린다.',
    '반드시 JSON만 출력:',
    '{"items":[{"유형":"","분류":"","새분류":false,"분류설명":"","대상일":"","대상시각":"","제목":"","내용":"","금액":null,"인물":[],"장소":"","태그":[],"상태":"","프로필항목":"","프로필값":"","확인필요":""}]}',
  ].join('\n');

  const out = extractJson_(callClaude_(system, text, 2000));
  const items = (out.items || []).map(it => normalizeItem_(it, cats));
  return { items: items, raw: text };
}

function normalizeItem_(it, cats) {
  const names = cats.map(c => c.name);
  const o = {
    유형: TYPES.indexOf(it.유형) >= 0 ? it.유형 : '일상',
    분류: String(it.분류 || '생활').trim(),
    대상일: String(it.대상일 || ''),
    대상시각: String(it.대상시각 || ''),
    제목: String(it.제목 || '').slice(0, 60),
    내용: String(it.내용 || ''),
    금액: toNum_(it.금액),
    인물: listStr_(it.인물),
    장소: String(it.장소 || ''),
    태그: listStr_(it.태그),
    상태: STATES.indexOf(it.상태) >= 0 ? it.상태 : (/일정/.test(it.유형) ? '예정' : '정상'),
    프로필항목: String(it.프로필항목 || ''),
    프로필값: String(it.프로필값 || ''),
    확인필요: String(it.확인필요 || ''),
  };
  o.새분류 = names.indexOf(o.분류) < 0;
  o.분류설명 = o.새분류 ? String(it.분류설명 || '') : '';
  return o;
}

// ───────────────────────── 기록: 저장(save) ─────────────────────────
function actSave_(req) {
  const items = req.items || [];
  if (!items.length) throw new Error('저장할 기록이 없습니다.');
  const raw = mask_(String(req.raw || ''));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SH.RECORD);
    const now = new Date();
    const nowStr = fmt_(now, 'yyyy-MM-dd HH:mm');
    const prefix = fmt_(now, 'yyyyMMdd');
    let seq = nextSeq_(sh, prefix);

    // 새 분류 승인분 추가
    const known = getCategories_().map(c => c.name);
    const catSh = ss.getSheetByName(SH.CATEGORY);
    items.forEach(it => {
      if (it.새분류승인 && it.분류 && known.indexOf(it.분류) < 0) {
        catSh.appendRow([it.분류, it.분류설명 || '', fmt_(now, 'yyyy-MM-dd')]);
        known.push(it.분류);
      }
    });

    const rows = [], ids = [];
    items.forEach(it => {
      const id = prefix + '-' + ('00' + (seq++)).slice(-3);
      ids.push(id);
      const cat = known.indexOf(it.분류) >= 0 ? it.분류 : '생활';
      rows.push([
        id, nowStr, it.대상일 || fmt_(now, 'yyyy-MM-dd'), it.대상시각 || '',
        it.유형, cat, mask_(it.제목 || ''), mask_(it.내용 || ''),
        toNum_(it.금액) === null ? '' : toNum_(it.금액),
        listStr_(it.인물), it.장소 || '', listStr_(it.태그),
        it.상태 || '정상', raw, '', '',
      ]);
      if (it.유형 === '정보' && it.프로필항목 && it.프로필값) upsertProfile_(it.프로필항목, mask_(it.프로필값));
    });
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, REC_HEADERS.length).setValues(rows);
    return { ids: ids, recent: recent_(5) };
  } finally {
    lock.releaseLock();
  }
}

function nextSeq_(sh, prefix) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 1;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  let max = 0;
  ids.forEach(id => { if (id.indexOf(prefix + '-') === 0) max = Math.max(max, parseInt(id.split('-')[1], 10) || 0); });
  return max + 1;
}

function upsertProfile_(key, val) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.PROFILE);
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const keys = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim());
    const i = keys.indexOf(key.trim());
    if (i >= 0) { sh.getRange(i + 2, 2, 1, 2).setValues([[val, today]]); return; }
  }
  sh.appendRow([key, val, today]);
}

// ───────────────────────── 질문(ask) ─────────────────────────
function actAsk_(req) {
  const q = String(req.question || '').trim();
  if (!q) throw new Error('질문이 비어 있습니다.');
  const cats = getCategories_().map(c => c.name);

  // 1) 질문 → 조회계획
  const planSys = [
    '너는 개인 비서 Lobby의 질문 해석기다. 카일님의 질문을 기록 시트 조회계획 JSON으로 바꾼다.',
    dateContext_(),
    '유형: ' + TYPES.join(', '),
    '분류: ' + cats.join(', '),
    '상태: ' + STATES.filter(s => s !== '삭제').join(', '),
    '종류: 조회(목록), 통계(합계·횟수·평균), 최근(가장 최근 이력), 요약(기간 회고), 계획(제안), 정보(프로필 사실 질문)',
    '규칙:',
    '- 기간은 대상일 기준. "이번 주"=이번 주 월~일, "이번 달"=1일~말일, "올해"=1/1~12/31, "최근"=지난 30일. 기간 언급이 없고 최근·정보 질문이면 기간 null(전체).',
    '- 검색어는 제목·내용·태그·인물·장소에서 찾을 핵심 단어. 동의어도 함께(예: 엔진오일 → ["엔진오일","오일"]).',
    '- 분류·유형은 확실할 때만 넣고, 애매하면 비워서 검색어로 찾는다.',
    '- 계산: 합계|횟수|평균|없음. 그룹: 월|분류|유형|없음.',
    '반드시 JSON만 출력:',
    '{"종류":"","기간":{"시작":"","끝":""},"유형":[],"분류":[],"상태":[],"검색어":[],"계산":"없음","그룹":"없음","정렬":"최신","개수":30}',
  ].join('\n');
  const plan = extractJson_(callClaude_(planSys, q, 600));

  // 2) 시트에서 필터·계산 (숫자는 프로그램이 계산)
  const result = runPlan_(plan);
  const profile = getProfile_();

  // 3) 답변 작성
  const ansSys = [
    '너는 카일님의 개인 비서 Lobby다. 아래 [조회결과]와 [프로필]만 근거로 답한다.',
    dateContext_(),
    '규칙:',
    '- 기록에 없는 내용은 지어내지 말고 "해당 기록이 없습니다"라고 말한 뒤, 지금 기록할지 묻는다.',
    '- 숫자(합계·건수·평균)는 [조회결과].통계 값을 그대로 쓴다. 직접 다시 계산하지 않는다.',
    '- 음성: 존댓말, 1~3문장, 첫 문장에 결론. 금액은 "48만 2천 원"처럼 읽기 쉽게. 목록이 4건 이상이면 3건까지만 말하고 "나머지는 화면에 정리해 두었습니다".',
    '- 상세: 화면용. 줄바꿈으로 구분한 짧은 줄들. 목록은 "• " 로 시작. 금액은 482,000원 형식.',
    '- 근거ID: 답변에 쓴 기록의 ID 배열(최대 10개).',
    '반드시 JSON만 출력: {"음성":"","상세":"","근거ID":[]}',
  ].join('\n');
  const ansUser = JSON.stringify({ 질문: q, 조회계획: plan, 조회결과: { 통계: result.stats, 기록: result.aiRows }, 프로필: profile });
  const ans = extractJson_(callClaude_(ansSys, ansUser, 1500));

  const evid = (ans.근거ID || []).map(String);
  const evidence = result.rows.filter(r => evid.indexOf(r.ID) >= 0).slice(0, 10);

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.LOG)
    .appendRow([fmt_(new Date(), 'yyyy-MM-dd HH:mm'), q, JSON.stringify(plan), ans.음성 || '', result.stats.건수]);

  return { speech: ans.음성 || '', detail: ans.상세 || '', evidence: evidence, stats: result.stats, plan: plan };
}

function runPlan_(plan) {
  let rows = getRecords_().filter(r => r.상태 !== '삭제');
  const p = plan || {};
  const from = p.기간 && p.기간.시작, to = p.기간 && p.기간.끝;
  if (from) rows = rows.filter(r => dayOf_(r) >= from);
  if (to) rows = rows.filter(r => dayOf_(r) <= to);
  if (p.유형 && p.유형.length) rows = rows.filter(r => p.유형.indexOf(r.유형) >= 0);
  if (p.분류 && p.분류.length) rows = rows.filter(r => p.분류.indexOf(r.분류) >= 0);
  if (p.상태 && p.상태.length) rows = rows.filter(r => p.상태.indexOf(r.상태) >= 0);
  if (p.검색어 && p.검색어.length) {
    const kws = p.검색어.map(k => String(k).toLowerCase()).filter(Boolean);
    rows = rows.filter(r => {
      const hay = [r.제목, r.내용, r.태그, r.인물, r.장소].join(' ').toLowerCase();
      return kws.some(k => hay.indexOf(k) >= 0);
    });
  }
  const asc = p.정렬 === '오래된';
  rows.sort((a, b) => (dayOf_(a) + a.대상시각).localeCompare(dayOf_(b) + b.대상시각) * (asc ? 1 : -1));

  // 통계
  const money = rows.map(r => r.금액).filter(v => typeof v === 'number');
  const sum = money.reduce((s, v) => s + v, 0);
  const stats = {
    건수: rows.length,
    금액건수: money.length,
    합계: sum,
    평균: money.length ? Math.round(sum / money.length) : 0,
    기간: { 시작: from || '', 끝: to || '' },
  };
  const g = p.그룹;
  if (g && g !== '없음') {
    const key = r => g === '월' ? dayOf_(r).slice(0, 7) : (g === '분류' ? r.분류 : r.유형);
    const groups = {};
    rows.forEach(r => {
      const k = key(r) || '(없음)';
      groups[k] = groups[k] || { 건수: 0, 합계: 0 };
      groups[k].건수++;
      if (typeof r.금액 === 'number') groups[k].합계 += r.금액;
    });
    stats.그룹 = groups;
  }

  const limit = Math.min(Number(p.개수) || CONFIG.MAX_ROWS_TO_AI, CONFIG.MAX_ROWS_TO_AI);
  const aiRows = rows.slice(0, limit).map(r => ({
    ID: r.ID, 날짜: dayOf_(r), 시각: r.대상시각, 유형: r.유형, 분류: r.분류,
    제목: r.제목, 내용: r.내용, 금액: r.금액, 인물: r.인물, 장소: r.장소, 상태: r.상태,
  }));
  return { rows: rows, aiRows: aiRows, stats: stats };
}

// ───────────────────────── 수정&삭제 탭 ─────────────────────────
function actList_(req) {
  const deleted = !!req.deleted;
  let rows = getRecords_().filter(r => deleted ? r.상태 === '삭제' : r.상태 !== '삭제');
  if (req.from || req.to) {
    const f = req.from || '0000-00-00', t = req.to || '9999-99-99';
    rows = rows.filter(r => {
      const d1 = dayOf_(r), d2 = String(r.입력일시).slice(0, 10);
      return (d1 >= f && d1 <= t) || (d2 >= f && d2 <= t);
    });
  }
  if (req.type) rows = rows.filter(r => r.유형 === req.type);
  if (req.q) {
    const k = String(req.q).toLowerCase();
    rows = rows.filter(r => [r.제목, r.내용, r.태그, r.인물, r.장소, r.원문, r.분류].join(' ').toLowerCase().indexOf(k) >= 0);
  }
  rows.sort((a, b) => String(b.입력일시 + b.ID).localeCompare(String(a.입력일시 + a.ID)));
  return { rows: rows.slice(0, 200), total: rows.length };
}

function actUpdate_(req) {
  const fields = req.fields || {};
  return withRow_(req.id, (sh, rowIdx, cur) => {
    const now = fmt_(new Date(), 'yyyy-MM-dd HH:mm');
    const hist = [];
    const next = cur.slice();
    EDITABLE.forEach(k => {
      if (!(k in fields)) return;
      let v = fields[k];
      if (k === '금액') v = toNum_(v) === null ? '' : toNum_(v);
      else if (k === '인물' || k === '태그') v = listStr_(v);
      else v = mask_(String(v == null ? '' : v));
      if (k === '유형' && TYPES.indexOf(v) < 0) return;
      if (k === '상태' && STATES.indexOf(v) < 0) return;
      const old = cellStr_(cur[COL[k]], k);
      if (String(old) !== String(v)) { next[COL[k]] = v; hist.push([now, req.id, k, old, v, '화면']); }
    });
    if (!hist.length) return { changed: 0 };
    next[COL['수정일시']] = now;
    sh.getRange(rowIdx, 1, 1, REC_HEADERS.length).setValues([next]);
    appendHistory_(hist);
    return { changed: hist.length, row: toObj_(next) };
  });
}

function actDelete_(req) {
  return withRow_(req.id, (sh, rowIdx, cur) => {
    const now = fmt_(new Date(), 'yyyy-MM-dd HH:mm');
    const old = cellStr_(cur[COL['상태']]);
    if (old === '삭제') return { already: true };
    sh.getRange(rowIdx, COL['상태'] + 1).setValue('삭제');
    sh.getRange(rowIdx, COL['삭제일시'] + 1).setValue(now);
    appendHistory_([[now, req.id, '상태', old, '삭제', '화면']]);
    return { deleted: req.id };
  });
}

function actRestore_(req) {
  return withRow_(req.id, (sh, rowIdx, cur) => {
    const now = fmt_(new Date(), 'yyyy-MM-dd HH:mm');
    // 삭제 직전 상태 찾기
    const hs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.HISTORY);
    const n = hs.getLastRow() - 1;
    let prev = '정상';
    if (n > 0) {
      const h = hs.getRange(2, 1, n, 6).getValues();
      for (let i = h.length - 1; i >= 0; i--) {
        if (String(h[i][1]) === String(req.id) && h[i][2] === '상태' && h[i][4] === '삭제') { prev = String(h[i][3]) || '정상'; break; }
      }
    }
    sh.getRange(rowIdx, COL['상태'] + 1).setValue(prev);
    sh.getRange(rowIdx, COL['삭제일시'] + 1).setValue('');
    appendHistory_([[now, req.id, '상태', '삭제', prev, '복구']]);
    return { restored: req.id, 상태: prev };
  });
}

function withRow_(id, fn) {
  if (!id) throw new Error('기록 ID가 없습니다.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.RECORD);
    const n = sh.getLastRow() - 1;
    if (n < 1) throw new Error('기록이 없습니다.');
    const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(id));
    if (i < 0) throw new Error('해당 기록을 찾지 못했습니다: ' + id);
    const rowIdx = i + 2;
    const cur = sh.getRange(rowIdx, 1, 1, REC_HEADERS.length).getValues()[0];
    return fn(sh, rowIdx, cur);
  } finally {
    lock.releaseLock();
  }
}

function appendHistory_(rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.HISTORY);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
}

/** 매일 새벽 4시: 삭제 후 30일 지난 기록을 휴지통 탭으로 옮김 */
function purgeTrash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH.RECORD), tr = ss.getSheetByName(SH.TRASH);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const limit = fmt_(new Date(Date.now() - CONFIG.TRASH_DAYS * 86400000), 'yyyy-MM-dd HH:mm');
  const data = sh.getRange(2, 1, n, REC_HEADERS.length).getValues();
  const move = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const d = cellStr_(data[i][COL['삭제일시']]);
    if (data[i][COL['상태']] === '삭제' && d && d < limit) { move.push(data[i]); sh.deleteRow(i + 2); }
  }
  if (move.length) tr.getRange(tr.getLastRow() + 1, 1, move.length, REC_HEADERS.length).setValues(move.reverse());
}

// ───────────────────────── 공통 유틸 ─────────────────────────
function getRecords_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.RECORD);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, REC_HEADERS.length).getValues().filter(r => r[0]).map(toObj_);
}

function toObj_(r) {
  const o = {};
  REC_HEADERS.forEach((h, i) => { o[h] = h === '금액' ? toNum_(r[i]) : cellStr_(r[i], h); });
  return o;
}

function cellStr_(v, h) {
  if (v instanceof Date) {
    if (h === '대상시각') return fmt_(v, 'HH:mm');
    if (h === '대상일') return fmt_(v, 'yyyy-MM-dd');
    return fmt_(v, 'yyyy-MM-dd HH:mm');
  }
  return v == null ? '' : String(v);
}

function dayOf_(r) { return r.대상일 || String(r.입력일시).slice(0, 10); }

function recent_(n) {
  return getRecords_().filter(r => r.상태 !== '삭제')
    .sort((a, b) => String(b.입력일시 + b.ID).localeCompare(String(a.입력일시 + a.ID))).slice(0, n);
}

function getCategories_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.CATEGORY);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 2).getValues().filter(r => r[0]).map(r => ({ name: String(r[0]).trim(), desc: String(r[1]) }));
}

function getProfile_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.PROFILE);
  const n = sh.getLastRow() - 1;
  if (n < 1) return {};
  const o = {};
  sh.getRange(2, 1, n, 2).getValues().forEach(r => { if (r[0]) o[String(r[0])] = cellStr_(r[1]); });
  return o;
}

function dateContext_() {
  const now = new Date();
  const wd = ['월', '화', '수', '목', '금', '토', '일'];
  const u = Number(fmt_(now, 'u'));            // 1=월 … 7=일
  const monday = new Date(now.getTime() - (u - 1) * 86400000);
  const sunday = new Date(monday.getTime() + 6 * 86400000);
  return '오늘: ' + fmt_(now, 'yyyy-MM-dd') + ' (' + wd[u - 1] + '요일), 현재 ' + fmt_(now, 'HH:mm') +
    '. 이번 주: ' + fmt_(monday, 'yyyy-MM-dd') + ' ~ ' + fmt_(sunday, 'yyyy-MM-dd') + ' (월~일).';
}

function fmt_(d, p) { return Utilities.formatDate(d, CONFIG.TZ, p); }

function toNum_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? Math.round(v) : null;
  const n = Number(String(v).replace(/[^\d.-]/g, ''));
  return String(v).replace(/[^\d]/g, '') === '' || !isFinite(n) ? null : Math.round(n);
}

function listStr_(v) {
  if (Array.isArray(v)) return v.map(s => String(s).trim()).filter(Boolean).join(', ');
  return String(v || '').trim();
}

/** 카드번호·계좌번호처럼 보이는 긴 숫자열 가리기 */
function mask_(s) {
  return String(s)
    .replace(/\b(\d{4})[-\s]?(\d{4})[-\s]?(\d{4})[-\s]?(\d{4})\b/g, '$1-****-****-****')
    .replace(/\b\d{2,6}-\d{2,6}-\d{4,8}\b/g, m => m.slice(0, 3) + '***');
}

function callClaude_(system, user, maxTokens) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('스크립트 속성에 ANTHROPIC_API_KEY가 없습니다.');
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CONFIG.MODEL, max_tokens: maxTokens || 1500, system: system,
      messages: [{ role: 'user', content: user }],
    }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (code !== 200) throw new Error('Claude API 오류(' + code + '): ' + (body.error && body.error.message || ''));
  return body.content.map(c => c.text || '').join('');
}

function extractJson_(text) {
  const m = String(text).match(/```(?:json)?\s*([\s\S]*?)```/);
  const s = m ? m[1] : String(text);
  const i = s.search(/[\[{]/);
  const j = Math.max(s.lastIndexOf('}'), s.lastIndexOf(']'));
  if (i < 0 || j < i) throw new Error('AI 응답을 해석하지 못했습니다. 다시 말씀해 주세요.');
  return JSON.parse(s.slice(i, j + 1));
}

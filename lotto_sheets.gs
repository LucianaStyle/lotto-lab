/**
 * 로또 6/45 + 연금복권 720+ 대시보드 — Google Sheets (Apps Script) · 엔진 v2 (2026-09)
 *
 * 설치 (시트 안에서):
 *  1. 확장 프로그램 > Apps Script → 이 파일 전체로 교체 → 저장
 *     (반드시 시트에 바인딩된 프로젝트. script.google.com의 독립형 프로젝트는 getUi() 오류)
 *  2. 시트 새로고침 → [복권분석 > 지금 동기화] 1회 실행 (권한 승인)
 *     첫 실행 때 자동 이전: 기존 추천 시트는 '로또추천_v1'·'연금추천_v1'로 보존되고,
 *     로또이력은 판매액 오류를 고친 새 스키마로 다시 받는다.
 *  3. [복권분석 > 자동 동기화 설치] → 3시간마다 자동 (추첨 결과 게시가 늦어도 다음 주기에 잡힌다)
 *
 * 구조 — 번호는 시트가 만들지 않는다:
 *   로컬 PC의 lotto_lab.py(엔진 v2)가 추천을 회차당 1번 생성해 GitHub 미러에 올리고,
 *   시트는 그것을 가져와 보여주고 채점한다. 그래서 리포트와 시트가 항상 같은 번호를 보여준다.
 *   (동행복권은 해외 IP를 차단해 Apps Script가 직접 수집할 수도 없다 — "사용할 수 없는 주소")
 *
 * 시트:
 *   로또추천 / 연금추천   이번 주 추천 + 지난 회차 채점(맞은 번호 색칠)
 *   성적                 엔진별 누적 성적 vs 이론 기대값
 *   편중감시              번호 구간·번호별·연금 끝자리 분포 (균등=1.00) — v1 편중 재발 감시
 *   번호인기도            5등 당첨자 수로 역산한 번호별 인기 (분할 회피의 근거)
 *   검증                 워크포워드 백테스트 + 로컬 엔진 실전 기록(대조군 포함)
 *   통계 / 로또이력 / 연금이력 / *_v1 (구 엔진 기록 보존)
 */

var MIRROR = 'https://raw.githubusercontent.com/LucianaStyle/lotto-lab/main/data/';
var FETCH_OPT = { muteHttpExceptions: true, headers: { 'Cache-Control': 'no-cache' } };

var SH = {
  LOTTO: '로또이력', PENSION: '연금이력', STATS: '통계',
  PICK: '로또추천', PPICK: '연금추천', PICK_V1: '로또추천_v1', PPICK_V1: '연금추천_v1',
  SCORE: '성적', BIAS: '편중감시', POP: '번호인기도', CHECK: '검증'
};
var TAB_ORDER = [SH.PICK, SH.PPICK, SH.SCORE, SH.BIAS, SH.POP, SH.CHECK, SH.STATS,
                 SH.LOTTO, SH.PENSION, SH.PICK_V1, SH.PPICK_V1];

var LOTTO_HEAD = ['회차', '추첨일', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', '보너스',
                  '1등당첨자', '1등금액', '판매액', '2등', '3등', '4등', '5등'];
var PENSION_HEAD = ['회차', '추첨일', '조', '번호', '보너스'];

// 추천 시트 레이아웃 (열 번호는 1부터)
var LV2 = { head: ['생성일', '대상회차', '추첨일', '세트', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6',
                   '분할배율', '적중', '등수'], target: 2, date: 3, n1: 5, hit: 12, rank: 13 };
var LV1 = { head: ['생성일', '대상회차', '추첨일', '세트', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6',
                   '합계', '홀수', '인기도', '적중', '등수'], target: 2, date: 3, n1: 5, hit: 14, rank: 15 };
var PL = { head: ['생성일', '대상회차', '추첨일', '순위', '조', '번호', '적중자리', '등수'],
           target: 2, date: 3, jo: 5, num: 6, hit: 7, rank: 8 };

var RANGES = [[1, 9], [10, 19], [20, 29], [30, 39], [40, 45]];

// 색상
var C_HIT_BG = '#c6efce', C_HIT_FG = '#0b6b3a';     // 번호 일치(초록)
var C_BONUS_BG = '#ffe08a', C_BONUS_FG = '#7a5200'; // 보너스 일치(호박)
var C_HEAD_BG = '#1f3864', C_HEAD_FG = '#ffffff';
var C_SUB_BG = '#e8eef7';
var C_GROUP_A = '#ffffff', C_GROUP_B = '#f2f6fc';   // 회차 그룹 교대 배경
var C_PENDING = '#9aa0a6';

// ───────────────────────── 메뉴 · 진입점 ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('복권분석')
    .addItem('지금 동기화 (데이터·추천·채점)', 'syncAll')
    .addItem('서식 다시 적용', 'beautifyAll')
    .addSeparator()
    .addItem('자동 동기화 설치 (3시간마다)', 'installTriggers')
    .addToUi();
}

// 구버전 트리거(토·목 22시 weeklyJob)가 남아 있어도 동작하도록 이름을 유지한다
function weeklyJob() { syncAll(); }

function syncAll() {
  migrateV2_();
  var changed = updateLotto() + updatePension() + importPicks_();
  importAux_();
  gradeAll_();
  buildStats();
  buildScore_();
  buildBias_();
  if (changed > 0) beautifyAll();
  toast_('동기화 완료' + (changed ? ' — 신규 ' + changed + '건' : ' — 변경 없음'));
}

function installTriggers() {
  var ss = SpreadsheetApp.getActive();
  if (ss.getSpreadsheetTimeZone() !== 'Asia/Seoul') ss.setSpreadsheetTimeZone('Asia/Seoul');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'weeklyJob' || f === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(3).create();
  toast_('자동 동기화 설치 완료 — 3시간마다 (Asia/Seoul)');
}

// 1회성 이전: 구 추천 시트 보존 + 이력 재수집 + 트리거 교체
function migrateV2_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('schema') === 'v2') return;
  var ss = SpreadsheetApp.getActive();
  [[SH.PICK, SH.PICK_V1], [SH.PPICK, SH.PPICK_V1]].forEach(function (p) {
    var old = ss.getSheetByName(p[0]);
    if (old && !ss.getSheetByName(p[1])) old.setName(p[1]);
  });
  var h = ss.getSheetByName(SH.LOTTO);
  if (h) h.clear();                     // 판매액이 절반으로 들어간 구 스키마 → 전량 재수집
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      var f = t.getHandlerFunction();
      if (f === 'weeklyJob' || f === 'syncAll') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('syncAll').timeBased().everyHours(3).create();
  } catch (e) { /* 트리거 권한이 없으면 메뉴의 [자동 동기화 설치]로 */ }
  props.setProperty('schema', 'v2');
}

// ───────────────────────── 유틸 ─────────────────────────

function sheet(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function toast_(msg) { SpreadsheetApp.getActive().toast(msg); }

// "20260711" / "2026-07-11" / Date → Date
function ymd_(v) {
  if (v instanceof Date) return v;
  var s = String(v).replace(/\D/g, '');
  if (s.length !== 8) return null;
  return new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8)));
}

// "2026-09-30 22:11" → Date
function ts_(v) {
  var m = String(v).match(/(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) : v;
}

function pad6_(v) {
  var s = String(v).replace(/\D/g, '');
  while (s.length < 6) s = '0' + s;
  return s;
}

function comb_(n, k) {
  if (k < 0 || k > n) return 0;
  var r = 1;
  for (var i = 0; i < k; i++) r = r * (n - i) / (i + 1);
  return r;
}

function getCsv_(file) {
  var res = UrlFetchApp.fetch(MIRROR + file, FETCH_OPT);
  if (res.getResponseCode() === 404) return [];     // 아직 생성 전인 보조 파일
  if (res.getResponseCode() !== 200) {
    throw new Error('미러 CSV 응답 HTTP ' + res.getResponseCode() + ' — 저장소 공개 여부 확인: ' + MIRROR + file);
  }
  return Utilities.parseCsv(res.getContentText()).slice(1).filter(function (r) { return r.length > 1; });
}

function lastValue_(sh, col) {
  return sh.getLastRow() > 1 ? Number(sh.getRange(sh.getLastRow(), col).getValue()) : 0;
}

function targetsIn_(sh, col) {
  var have = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[Number(r[0])] = 1; });
  }
  return have;
}

function append_(sh, rows, width) {
  if (!rows.length) return 0;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, width).setValues(rows);
  return rows.length;
}

// ───────────────────────── 데이터 (GitHub 미러) ─────────────────────────

function updateLotto() {
  var sh = sheet(SH.LOTTO, LOTTO_HEAD);
  var last = lastValue_(sh, 1);
  var rows = getCsv_('lotto_history.csv')
    .filter(function (r) { return r.length >= 16 && Number(r[0]) > last; })
    .map(function (r) { return [Number(r[0]), ymd_(r[1])].concat(r.slice(2, 16).map(Number)); })
    .sort(function (a, b) { return a[0] - b[0]; });
  return append_(sh, rows, LOTTO_HEAD.length);
}

function updatePension() {
  var sh = sheet(SH.PENSION, PENSION_HEAD);
  var last = lastValue_(sh, 1);
  var rows = getCsv_('pension_history.csv')
    .filter(function (r) { return r.length >= 5 && Number(r[0]) > last; })
    .map(function (r) { return [Number(r[0]), ymd_(r[1]), Number(r[2]), pad6_(r[3]), pad6_(r[4])]; })
    .sort(function (a, b) { return a[0] - b[0]; });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 4, rows.length, 2).setNumberFormat('@');  // 앞자리 0 보존
  return append_(sh, rows, PENSION_HEAD.length);
}

// 로컬 엔진 v2의 공식 추천을 가져온다 (회차당 1번 생성된 것 — 시트는 새로 뽑지 않는다)
function importPicks_() {
  var n = 0;
  var sh = sheet(SH.PICK, LV2.head);
  var have = targetsIn_(sh, LV2.target);
  // created,target,draw_date,engine,strategy,set,n1..n6,split_mult
  var rows = getCsv_('lotto_picks.csv')
    .filter(function (r) { return r[3] === 'v2' && r[4] === '공식' && !have[Number(r[1])]; })
    .map(function (r) {
      return [ts_(r[0]), Number(r[1]), ymd_(r[2]), r[5]].concat(r.slice(6, 12).map(Number), [Number(r[12]), '', '']);
    });
  n += append_(sh, rows, LV2.head.length);

  var psh = sheet(SH.PPICK, PL.head);
  var phave = targetsIn_(psh, PL.target);
  // created,target,draw_date,engine,rank,jo,num
  var prows = getCsv_('pension_picks.csv')
    .filter(function (r) { return r[3] === 'v2' && !phave[Number(r[1])]; })
    .map(function (r) { return [ts_(r[0]), Number(r[1]), ymd_(r[2]), r[4] + '순위', Number(r[5]), pad6_(r[6]), '', '']; });
  if (prows.length) psh.getRange(psh.getLastRow() + 1, PL.num, prows.length, 1).setNumberFormat('@');
  n += append_(psh, prows, PL.head.length);
  return n;
}

// 번호인기도 · 검증 시트 (로컬 엔진 산출물을 그대로 표시)
function importAux_() {
  var pop = getCsv_('popularity.csv');     // number,popularity,rank
  if (pop.length) buildPopularity_(pop);
  buildCheck_(getCsv_('backtest.csv'), getCsv_('scoreboard.csv'));
}

// ───────────────────────── 채점 ─────────────────────────

function gradeAll_() {
  var ss = SpreadsheetApp.getActive();
  var lres = lottoResults_(), pres = pensionResults_();
  [[SH.PICK, LV2], [SH.PICK_V1, LV1]].forEach(function (p) {
    var sh = ss.getSheetByName(p[0]);
    if (sh && sh.getLastRow() > 1) gradeLotto_(sh, p[1], lres);
  });
  [SH.PPICK, SH.PPICK_V1].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) gradePension_(sh, pres);
  });
}

function lottoResults_() {
  var res = {}, sh = SpreadsheetApp.getActive().getSheetByName(SH.LOTTO);
  if (!sh || sh.getLastRow() < 2) return res;
  sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(function (r) {
    res[Number(r[0])] = { date: r[1], nums: r.slice(2, 8).map(Number), bonus: Number(r[8]) };
  });
  return res;
}

function pensionResults_() {
  var res = {}, sh = SpreadsheetApp.getActive().getSheetByName(SH.PENSION);
  if (!sh || sh.getLastRow() < 2) return res;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
    res[Number(r[0])] = { date: r[1], jo: Number(r[2]), num: pad6_(r[3]), bonus: pad6_(r[4]) };
  });
  return res;
}

function lottoRank_(m, bonusHit) {
  return m === 6 ? '1등' : (m === 5 && bonusHit) ? '2등' : m === 5 ? '3등'
       : m === 4 ? '4등' : m === 3 ? '5등' : '낙첨';
}

function gradeLotto_(sh, L, res) {
  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, 1, n, L.head.length).getValues();
  var numRange = sh.getRange(2, L.n1, n, 6);
  var bgs = numRange.getBackgrounds(), fgs = numRange.getFontColors();
  var hits = [], ranks = [], dates = [];
  for (var i = 0; i < n; i++) {
    var r = res[Number(vals[i][L.target - 1])];
    if (!r) {
      hits.push([vals[i][L.hit - 1] || '대기']);
      ranks.push([vals[i][L.rank - 1] || '추첨 전']);
      dates.push([vals[i][L.date - 1]]);
      continue;
    }
    var m = 0, bonusHit = false;
    for (var c = 0; c < 6; c++) {
      var num = Number(vals[i][L.n1 - 1 + c]);
      if (r.nums.indexOf(num) >= 0) { m++; bgs[i][c] = C_HIT_BG; fgs[i][c] = C_HIT_FG; }
      else if (num === r.bonus) { bonusHit = true; bgs[i][c] = C_BONUS_BG; fgs[i][c] = C_BONUS_FG; }
      else { bgs[i][c] = '#ffffff'; fgs[i][c] = '#000000'; }
    }
    hits.push([m + '개' + (bonusHit ? '+보너스' : '')]);
    ranks.push([lottoRank_(m, bonusHit)]);
    dates.push([r.date]);
  }
  numRange.setBackgrounds(bgs).setFontColors(fgs);
  sh.getRange(2, L.hit, n, 1).setValues(hits);
  sh.getRange(2, L.rank, n, 1).setValues(ranks);
  sh.getRange(2, L.date, n, 1).setValues(dates);
}

// 연금 등수: 끝에서부터 연속 일치한 자리수
//   1등 조+6자리 / 2등 6자리 / 3등 뒤5 / 4등 뒤4 / 5등 뒤3 / 6등 뒤2 / 7등 뒤1
function gradePension_(sh, res) {
  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, 1, n, PL.head.length).getValues();
  var hitStyle = SpreadsheetApp.newTextStyle().setForegroundColor(C_HIT_FG).setBold(true).build();
  var plain = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setBold(false).build();
  var rts = [], hits = [], ranks = [], dates = [];
  for (var i = 0; i < n; i++) {
    var num = pad6_(vals[i][PL.num - 1]);
    var r = res[Number(vals[i][PL.target - 1])];
    var rt = SpreadsheetApp.newRichTextValue().setText(num).setTextStyle(0, 6, plain);
    if (!r) {
      rts.push([rt.build()]);
      hits.push([vals[i][PL.hit - 1] || '대기']);
      ranks.push([vals[i][PL.rank - 1] || '추첨 전']);
      dates.push([vals[i][PL.date - 1]]);
      continue;
    }
    var m = 0;
    while (m < 6 && num[5 - m] === r.num[5 - m]) m++;
    var joHit = Number(vals[i][PL.jo - 1]) === r.jo;
    var rank = (m === 6 && joHit) ? '1등' : m === 6 ? '2등' : m === 5 ? '3등' : m === 4 ? '4등'
             : m === 3 ? '5등' : m === 2 ? '6등' : m === 1 ? '7등' : '낙첨';
    if (num === r.bonus) rank += '(보너스 일치)';
    if (m > 0) rt.setTextStyle(6 - m, 6, hitStyle);
    rts.push([rt.build()]);
    hits.push([m + '자리' + (joHit ? '+조일치' : '')]);
    ranks.push([rank]);
    dates.push([r.date]);
  }
  sh.getRange(2, PL.num, n, 1).setRichTextValues(rts);
  sh.getRange(2, PL.hit, n, 1).setValues(hits);
  sh.getRange(2, PL.rank, n, 1).setValues(ranks);
  sh.getRange(2, PL.date, n, 1).setValues(dates);
}

// ───────────────────────── 성적 ─────────────────────────

function lottoHits_(name, L) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, L.hit, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
    var m = String(r[0]).match(/^(\d)개/);
    if (m) out.push(Number(m[1]));
  });
  return out;
}

function pensionTails_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, PL.hit, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
    var m = String(r[0]).match(/^(\d)자리/);
    if (m) out.push(Number(m[1]));
  });
  return out;
}

function buildScore_() {
  var sh = sheet(SH.SCORE);
  sh.clear();
  var W = 13, rows = [];
  var pad = function (r) { while (r.length < W) r.push(''); return r; };
  var pk = [];
  for (var k = 0; k <= 6; k++) pk.push(comb_(6, k) * comb_(39, 6 - k) / comb_(45, 6));

  rows.push(pad(['로또 — 세트별 적중 개수 (실제 / 이론 기대)']));
  rows.push(['구분', '채점 세트', '평균 적중', '이론 평균', '0개', '1개', '2개', '3개', '4개', '5개', '6개',
             '당첨(3개↑)', '이론 기대']);
  [['v2 현행', SH.PICK, LV2], ['v1 구엔진', SH.PICK_V1, LV1]].forEach(function (g) {
    var h = lottoHits_(g[1], g[2]), n = h.length;
    var dist = [0, 0, 0, 0, 0, 0, 0], sum = 0;
    h.forEach(function (x) { dist[x]++; sum += x; });
    rows.push([g[0], n, n ? Math.round(sum / n * 100) / 100 : '-', 0.8].concat(dist,
              [dist[3] + dist[4] + dist[5] + dist[6], Math.round(n * (pk[3] + pk[4] + pk[5] + pk[6]) * 100) / 100]));
    rows.push(['  └ 이론 기대', '', '', ''].concat(pk.map(function (p) { return Math.round(p * n * 10) / 10; }), ['', '']));
  });
  rows.push(pad(['']));
  rows.push(pad(['연금 — 끝자리 일치 (후보 전체 기준, 실제 / 이론 기대)']));
  rows.push(pad(['구분', '채점 후보', '평균 일치', '이론 평균', '낙첨', '7등', '6등', '5등', '4등', '3등', '2등↑',
                 '7등↑', '이론 기대']));
  var pt = [0.9, 0.09, 0.009, 0.0009, 0.00009, 0.000009, 0.000001];   // 끝 k자리 정확히 일치 확률
  [['v2 현행', SH.PPICK], ['v1 구엔진', SH.PPICK_V1]].forEach(function (g) {
    var t = pensionTails_(g[1]), n = t.length, d = [0, 0, 0, 0, 0, 0, 0], sum = 0;
    t.forEach(function (x) { d[x]++; sum += x; });
    rows.push([g[0], n, n ? Math.round(sum / n * 1000) / 1000 : '-', 0.111].concat(d.slice(0, 6), [d[6], n - d[0],
              Math.round(n * 0.1 * 10) / 10]));
    rows.push(['  └ 이론 기대', '', '', ''].concat(pt.map(function (p) { return Math.round(p * n * 10) / 10; }), ['', '']));
  });
  rows.push(pad(['']));
  rows.push(pad(['※ 각 추첨은 독립시행이라 어떤 엔진도 적중 확률을 바꾸지 못한다. 실제가 이론 기대 근처면 정상.']));
  rows.push(pad(['※ 연금 v1은 후보의 64%가 끝자리 4에 몰려, 그 주 끝자리가 4가 아니면 대부분 동시에 낙첨됐다.']));
  sh.getRange(1, 1, rows.length, W).setValues(rows.map(pad));

  [1, 9].forEach(function (r) { sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG); });
  [2, 10].forEach(function (r) { sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground(C_SUB_BG); });
  [4, 6, 12, 14].forEach(function (r) { sh.getRange(r, 1, 1, W).setFontColor(C_PENDING); });
  sh.getRange(1, 2, rows.length, W - 1).setHorizontalAlignment('center');
  sh.setColumnWidth(1, 130);
}

// ───────────────────────── 편중 감시 ─────────────────────────

function pickedNumbers_(name, L) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, L.n1, sh.getLastRow() - 1, 6).getValues().forEach(function (r) {
    r.forEach(function (v) { if (Number(v) >= 1 && Number(v) <= 45) out.push(Number(v)); });
  });
  return out;
}

function pensionTailDigits_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, PL.num, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
    out.push(Number(pad6_(r[0])[5]));
  });
  return out;
}

function buildBias_() {
  var sh = sheet(SH.BIAS);
  sh.clear();
  sh.clearConditionalFormatRules();
  var W = 11, rows = [], pad = function (r) { while (r.length < W) r.push(''); return r; };
  var groups = [['v2 현행', pickedNumbers_(SH.PICK, LV2)], ['v1 구엔진', pickedNumbers_(SH.PICK_V1, LV1)]];

  rows.push(pad(['① 로또 번호 구간 비중 (균등 = 1.00, 0.8 미만·1.2 초과면 편중)']));
  rows.push(pad(['구분', '1~9', '10~19', '20~29', '30~39', '40~45', '최대편차', '표본(번호 수)']));
  groups.forEach(function (g) {
    var ns = g[1], n = ns.length, row = [g[0]], dev = 0;
    RANGES.forEach(function (rg) {
      var cnt = ns.filter(function (x) { return x >= rg[0] && x <= rg[1]; }).length;
      var v = n ? Math.round(cnt / n / ((rg[1] - rg[0] + 1) / 45) * 100) / 100 : '';
      if (n) dev = Math.max(dev, Math.abs(v - 1));
      row.push(v);
    });
    rows.push(pad(row.concat([n ? Math.round(dev * 100) + '%' : '-', n])));
  });

  rows.push(pad(['']));
  var gridStarts = [];
  groups.forEach(function (g) {
    rows.push(pad(['② 번호별 추천 빈도 — ' + g[0] + ' (균등 = 1.00)']));
    rows.push(pad(['구간', '+0', '+1', '+2', '+3', '+4', '+5', '+6', '+7', '+8', '+9']));
    var cnt = {}, n = g[1].length;
    g[1].forEach(function (x) { cnt[x] = (cnt[x] || 0) + 1; });
    gridStarts.push(rows.length + 1);
    for (var d = 0; d < 5; d++) {
      var row = [(d * 10) + '번대'];
      for (var u = 0; u < 10; u++) {
        var num = d * 10 + u;
        row.push(num >= 1 && num <= 45 && n ? Math.round((cnt[num] || 0) / (n / 45) * 100) / 100 : '');
      }
      rows.push(row);
    }
    rows.push(pad(['']));
  });

  rows.push(pad(['③ 연금 후보 끝자리 분포 (균등 = 1.00) — 연금 등수는 끝자리부터 정해진다']));
  rows.push(pad(['구분', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']));
  var tailStart = rows.length + 1;
  [['v2 현행', pensionTailDigits_(SH.PPICK)], ['v1 구엔진', pensionTailDigits_(SH.PPICK_V1)]].forEach(function (g) {
    var n = g[1].length, cnt = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    g[1].forEach(function (x) { cnt[x]++; });
    rows.push([g[0]].concat(cnt.map(function (c) { return n ? Math.round(c / (n / 10) * 100) / 100 : ''; })));
  });
  sh.getRange(1, 1, rows.length, W).setValues(rows.map(pad));

  // 색: 1.00 흰색, 낮으면 빨강, 높으면 주황
  var ranges = [sh.getRange(3, 2, 2, 5), sh.getRange(tailStart, 2, 2, 10)];
  gridStarts.forEach(function (s) { ranges.push(sh.getRange(s, 2, 5, 10)); });
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#e67c73', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMaxpointWithValue('#f6b26b', SpreadsheetApp.InterpolationType.NUMBER, '2')
    .setRanges(ranges).build()]);
  sh.getRange(1, 1, rows.length, W).setHorizontalAlignment('center');
  sh.getRange(1, 1, rows.length, 1).setHorizontalAlignment('left');
  for (var i = 0; i < rows.length; i++) {
    if (/^[①②③]/.test(String(rows[i][0]))) {
      sh.getRange(i + 1, 1, 1, W).setFontWeight('bold').setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG);
    } else if (rows[i][0] === '구분' || rows[i][0] === '구간') {
      sh.getRange(i + 1, 1, 1, W).setFontWeight('bold').setBackground(C_SUB_BG);
    }
  }
  sh.setColumnWidth(1, 110);
}

// ───────────────────────── 번호인기도 · 검증 ─────────────────────────

function buildPopularity_(pop) {
  var sh = sheet(SH.POP);
  sh.clear();
  sh.clearConditionalFormatRules();
  var W = 11, rows = [], pad = function (r) { while (r.length < W) r.push(''); return r; };
  var v = {};
  pop.forEach(function (r) { v[Number(r[0])] = Number(r[1]); });
  rows.push(pad(['번호별 인기도 — 사람들이 얼마나 많이 고르는가 (평균 = 1.000, 빨강=인기·파랑=비인기)']));
  rows.push(pad(['구간', '+0', '+1', '+2', '+3', '+4', '+5', '+6', '+7', '+8', '+9']));
  for (var d = 0; d < 5; d++) {
    var row = [(d * 10) + '번대'];
    for (var u = 0; u < 10; u++) { var num = d * 10 + u; row.push(v[num] ? v[num] : ''); }
    rows.push(row);
  }
  var sorted = Object.keys(v).map(Number).sort(function (a, b) { return v[b] - v[a]; });
  rows.push(pad(['']));
  rows.push(pad(['가장 많이 찍힘', sorted.slice(0, 8).join(', ')]));
  rows.push(pad(['가장 덜 찍힘', sorted.slice(-8).reverse().join(', ')]));
  rows.push(pad(['']));
  rows.push(pad(['어떻게 구했나: 5등(3개 일치)은 회당 약 270만 명. 무작위 구매라면 기대 인원이 정확히 계산되므로,']));
  rows.push(pad(['  실제가 기대보다 많았던 회차의 당첨번호 = 많이 찍힌 번호. 최근 10년치로 번호별 기여를 역산했다.']));
  rows.push(pad(['  과거로 학습해 미래 200회를 예측한 상관 0.83 — 실재하는 효과다. 다만 차이는 ±3% 수준으로 작다.']));
  rows.push(pad(['쓰는 곳: 추천 엔진이 인기 번호를 조금 덜 골라, 1등 당첨 시 나눠 가질 사람을 줄인다(당첨 확률은 동일).']));
  sh.getRange(1, 1, rows.length, W).setValues(rows.map(pad));
  sh.getRange(3, 2, 5, 10).setNumberFormat('0.000').setHorizontalAlignment('center');
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#6fa8dc', SpreadsheetApp.InterpolationType.NUMBER, '0.98')
    .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMaxpointWithValue('#e67c73', SpreadsheetApp.InterpolationType.NUMBER, '1.025')
    .setRanges([sh.getRange(3, 2, 5, 10)]).build()]);
  sh.getRange(1, 1, 1, W).setFontWeight('bold').setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG);
  sh.getRange(2, 1, 1, W).setFontWeight('bold').setBackground(C_SUB_BG);
  sh.setColumnWidth(1, 110);
}

function buildCheck_(bt, sb) {
  var sh = sheet(SH.CHECK);
  sh.clear();
  var W = 8, rows = [], pad = function (r) { while (r.length < W) r.push(''); return r; };
  rows.push(pad(['① 워크포워드 백테스트 — 각 회차 직전까지의 데이터만으로 추천하고 실제 결과로 채점']));
  rows.push(pad(['엔진', '회차', '평균 적중', '3개↑ 세트', '주당 5등↑ 확률', '분할배율', '구간 최대편차', '20번대 비중']));
  // backtest.csv: engine,draws,sets,avg_hits,pct_prize,weekly_any_prize_mc,split_mult,range_dev,share_20s
  bt.forEach(function (r) {
    rows.push([r[0], Number(r[1]), Number(r[3]), Number(r[4]) + '%', Number(r[5]) + '%',
               Number(r[6]), Number(r[7]) + '%', Number(r[8]) + 'x']);
  });
  rows.push(pad(['이론', '', 0.8, '2.38%', '(무작위) 11.3%', 1, '0%', '1.00x']));
  rows.push(pad(['']));
  rows.push(pad(['읽는 법']));
  rows.push(pad(['· 평균 적중: 모든 엔진이 이론값 0.80 근처 — 적중 확률은 어떤 방법으로도 못 바꾼다.']));
  rows.push(pad(['· 주당 5등↑: 5세트 중 1개라도 3개 이상 맞힐 확률(4만 회 모의추첨). 30개 번호 분산으로 v2가 가장 높다.']));
  rows.push(pad(['· 분할배율: 1등 당첨 시 함께 당첨될 예상 인원(무작위=1). v1은 편중이 컸지만 효과가 거의 없었다.']));
  rows.push(pad(['· 구간 최대편차·20번대 비중: v1은 20번대를 절반만 뽑았다(0.50x). v2는 구조적으로 편중이 불가능하다.']));
  rows.push(pad(['']));
  var sbStart = rows.length + 1;
  rows.push(pad(['② 로컬 엔진 실전 기록 (대조군 = 같은 주 완전 무작위 5세트, 구매용 아님)']));
  rows.push(pad(['게임', '구분', '회차', '장수', '평균 적중', '이론', '당첨(실제/이론)', '분포']));
  // scoreboard.csv: game,group,draws,tickets,avg_hits,theory_avg,prize_count,theory_prize,detail
  sb.forEach(function (r) {
    rows.push([r[0], r[1], Number(r[2]), Number(r[3]), Number(r[4]), Number(r[5]), r[6] + ' / ' + r[7], r[8]]);
  });
  sh.getRange(1, 1, rows.length, W).setValues(rows.map(pad));
  [1, sbStart].forEach(function (r) { sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG); });
  [2, sbStart + 1].forEach(function (r) { sh.getRange(r, 1, 1, W).setFontWeight('bold').setBackground(C_SUB_BG); });
  sh.getRange(2, 2, rows.length - 1, W - 1).setHorizontalAlignment('center');
  sh.setColumnWidth(1, 110);
}

// ───────────────────────── 통계 ─────────────────────────

function buildStats() {
  var src = SpreadsheetApp.getActive().getSheetByName(SH.LOTTO);
  if (!src || src.getLastRow() < 2) return;
  var data = src.getRange(2, 1, src.getLastRow() - 1, 8).getValues();
  var freq = {}, rfreq = {}, lastSeen = {}, recentFrom = Math.max(0, data.length - 52);
  data.forEach(function (r, idx) {
    for (var c = 2; c <= 7; c++) {
      var n = r[c];
      freq[n] = (freq[n] || 0) + 1;
      lastSeen[n] = r[0];
      if (idx >= recentFrom) rfreq[n] = (rfreq[n] || 0) + 1;
    }
  });
  var latest = data[data.length - 1][0];
  var sh = sheet(SH.STATS);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, 1, 4).setValues([['번호', '역대출현', '최근52회', '미출현회차']]);
  var rows = [];
  for (var n = 1; n <= 45; n++) rows.push([n, freq[n] || 0, rfreq[n] || 0, latest - (lastSeen[n] || 0)]);
  sh.getRange(2, 1, 45, 4).setValues(rows);
  sh.getRange(1, 6, 3, 2).setValues([
    ['기준 회차', latest + '회'],
    ['갱신', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')],
    ['비고', '출현 빈도는 균등분포와 구별 불가 — 참고용 (예측력 없음)']
  ]);
  headerStyle_(sh, 4);
  sh.getRange(2, 1, 45, 4).setHorizontalAlignment('center');
  sh.getRange(1, 6, 3, 1).setFontWeight('bold');
  sh.setConditionalFormatRules([2, 4].map(function (col) {
    return SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.MIN, '')
      .setGradientMaxpointWithValue(col === 2 ? '#4a86c8' : '#f6b26b', SpreadsheetApp.InterpolationType.MAX, '')
      .setRanges([sh.getRange(2, col, 45, 1)]).build();
  }));
}

// ───────────────────────── 서식 ─────────────────────────

function beautifyAll() {
  var ss = SpreadsheetApp.getActive();
  var g = function (n) { return ss.getSheetByName(n); };
  if (g(SH.LOTTO)) formatHistory_(g(SH.LOTTO), LOTTO_HEAD.length, true);
  if (g(SH.PENSION)) formatHistory_(g(SH.PENSION), PENSION_HEAD.length, false);
  if (g(SH.PICK)) formatPicks_(g(SH.PICK), LV2, 6);
  if (g(SH.PICK_V1)) formatPicks_(g(SH.PICK_V1), LV1, 6);
  if (g(SH.PPICK)) formatPicks_(g(SH.PPICK), PL, 1);
  if (g(SH.PPICK_V1)) formatPicks_(g(SH.PPICK_V1), PL, 1);
  var pos = 1;
  TAB_ORDER.forEach(function (name) {
    var sh = g(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(pos++); }
  });
  if (g(SH.PICK)) ss.setActiveSheet(g(SH.PICK));
}

function headerStyle_(sh, nCols) {
  sh.getRange(1, 1, 1, nCols)
    .setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 30);
  sh.setFrozenRows(1);
}

// 이력: 날짜 서식 + 연도 경계 굵은 선 + 금액·인원 천단위
function formatHistory_(sh, nCols, isLotto) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  headerStyle_(sh, nCols);
  sh.getRange(2, 1, n, nCols).setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('center')
    .setBorder(null, null, null, null, false, false);
  sh.getRange(2, 2, n, 1).setNumberFormat('yyyy-mm-dd(ddd)');
  if (isLotto) sh.getRange(2, 10, n, 7).setNumberFormat('#,##0');
  var dates = sh.getRange(2, 2, n, 1).getValues(), prev = null;
  for (var i = 0; i < n; i++) {
    var d = dates[i][0];
    if (!(d instanceof Date)) continue;
    if (prev !== null && d.getFullYear() !== prev) {
      sh.getRange(i + 2, 1, 1, nCols).setBorder(true, null, null, null, null, null,
        C_HEAD_BG, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    prev = d.getFullYear();
  }
  sh.autoResizeColumns(1, nCols);
}

// 추천: 회차 그룹마다 배경 교대 + 경계선, 등수 강조. 번호 열 배경은 채점 색이라 건드리지 않는다.
function formatPicks_(sh, L, numCount) {
  var n = sh.getLastRow() - 1, W = L.head.length;
  if (n < 1) return;
  headerStyle_(sh, W);
  var numStart = L.n1 || L.num;
  var all = sh.getRange(2, 1, n, W);
  all.setFontFamily('Roboto Mono').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
     .setBorder(null, null, null, null, false, false);
  sh.getRange(2, 1, n, 1).setNumberFormat('yyyy-mm-dd HH:mm').setHorizontalAlignment('left');
  sh.getRange(2, L.date, n, 1).setNumberFormat('yyyy-mm-dd(ddd)');

  var vals = all.getValues(), bgs = all.getBackgrounds(), prevKey = null, grp = -1;
  for (var i = 0; i < n; i++) {
    var key = String(vals[i][L.target - 1]);
    if (key !== prevKey) {
      grp++;
      if (i > 0) sh.getRange(i + 2, 1, 1, W).setBorder(true, null, null, null, null, null,
        C_HEAD_BG, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      prevKey = key;
    }
    for (var c = 0; c < W; c++) {
      if (c + 1 >= numStart && c + 1 < numStart + numCount) continue;
      bgs[i][c] = grp % 2 === 0 ? C_GROUP_A : C_GROUP_B;
    }
  }
  all.setBackgrounds(bgs);

  var ranks = sh.getRange(2, L.rank, n, 1).getValues(), fg = [], fw = [];
  ranks.forEach(function (r) {
    var v = String(r[0]), win = /^[1-7]등/.test(v);
    fg.push([win ? C_HIT_FG : (v === '추첨 전' ? C_PENDING : '#000000')]);
    fw.push([win ? 'bold' : 'normal']);
  });
  sh.getRange(2, L.rank, n, 1).setFontColors(fg).setFontWeights(fw);
  sh.getRange(2, numStart, n, numCount).setFontWeight('bold').setFontSize(11);
  if (L === PL) sh.getRange(2, PL.num, n, 1).setNumberFormat('@');
  if (L === LV2) sh.getRange(2, 11, n, 1).setNumberFormat('0.00');
  sh.autoResizeColumns(1, W);
  sh.setColumnWidth(1, 130);
}

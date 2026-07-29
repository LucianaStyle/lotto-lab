/**
 * 로또 6/45 + 연금복권 720+ 자동 분석 — Google Sheets (Apps Script)
 *
 * 설치:
 *  1. 스프레드시트를 열고 그 안에서 확장 프로그램 > Apps Script (반드시 시트에 바인딩된 프로젝트여야
 *     메뉴가 동작한다. script.google.com에서 만든 독립형 프로젝트는 getUi() 오류가 난다)
 *  2. 이 파일 전체를 붙여넣고 저장 → 시트로 돌아와 새로고침 → 메뉴에 [복권분석]이 생김
 *  3. [복권분석 > 전체 업데이트+추천] 최초 1회 실행 (권한 승인 필요)
 *  4. [복권분석 > 주간 자동실행 설치] → 매주 토 22시(로또)·목 22시(연금) 자동 갱신
 *
 * 시트 구성:
 *  로또이력 / 연금이력 — 원본 당첨 데이터 (연도·월 경계 구분선)
 *  통계               — 번호별 출현·미출현 + 조합 프로파일
 *  로또추천 / 연금추천 — 회차별 추천, 추첨 후 자동 채점(적중 번호 색칠)
 *  성적               — 누적 적중 성적표 (이론 기대값 대비)
 *
 * ⚠️ 동행복권은 해외 IP를 차단하므로 Apps Script의 직접 수집은 실패한다("사용할 수 없는 주소").
 *    → 로컬 PC(lotto_lab.py)가 갱신·푸시한 GitHub 미러 CSV에서 가져온다.
 */

var MIRROR_LOTTO_CSV = 'https://raw.githubusercontent.com/LucianaStyle/lotto-lab/main/data/lotto_history.csv';
var MIRROR_PENSION_CSV = 'https://raw.githubusercontent.com/LucianaStyle/lotto-lab/main/data/pension_history.csv';

var BASE = 'https://www.dhlottery.co.kr';
var FETCH_OPT = {
  headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
  muteHttpExceptions: true
};

// 시트 이름
var SH_LOTTO = '로또이력', SH_PENSION = '연금이력', SH_STATS = '통계';
var SH_PICK = '로또추천', SH_PPICK = '연금추천', SH_SCORE = '성적';

// 색상 — 적중 표시
var C_HIT_BG = '#c6efce', C_HIT_FG = '#0b6b3a';   // 번호 일치(초록)
var C_BONUS_BG = '#ffe08a', C_BONUS_FG = '#7a5200'; // 보너스 일치(호박)
var C_MISS_BG = '#ffffff';
var C_HEAD_BG = '#1f3864', C_HEAD_FG = '#ffffff';
var C_GROUP_A = '#ffffff', C_GROUP_B = '#f2f6fc';   // 회차 그룹 교대 배경
var C_PENDING = '#9aa0a6';                          // 미추첨 회색

var PICK_HEAD = ['생성일', '대상회차', '추첨일', '세트', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6',
                 '합계', '홀수', '인기도', '적중', '등수'];
var PICK_N1 = 5;   // n1 열 위치(1-indexed)
var PICK_HIT = 14, PICK_RANK = 15;

var PPICK_HEAD = ['생성일', '대상회차', '추첨일', '순위', '조', '번호', '적중자리', '등수'];
var PPICK_NUM = 6, PPICK_HIT = 7, PPICK_RANK = 8;
var PENSION_N = 20;   // 연금 후보 수 — 조·번호 조합이 1장뿐이라 품절 대비로 넉넉히 뽑는다

// ───────────────────────── 메뉴 ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('복권분석')
    .addItem('전체 업데이트+추천', 'weeklyJob')
    .addSeparator()
    .addItem('데이터 갱신 (로또+연금)', 'updateAll')
    .addItem('통계 재계산', 'buildStats')
    .addItem('추천 번호 생성', 'buildPicks')
    .addItem('추천 재생성 (중복 허용)', 'buildPicksForce')
    .addSeparator()
    .addItem('결과 채점 (적중 색칠)', 'gradeResults')
    .addItem('서식 다시 적용', 'beautifyAll')
    .addSeparator()
    .addItem('주간 자동실행 설치', 'installTriggers')
    .addToUi();
}

function weeklyJob() {
  updateLotto();
  updatePension();
  buildStats();
  gradeResults();   // 새 결과가 들어왔으니 지난 추천부터 채점
  buildPicks();     // 그 다음 회차 추천 생성
  beautifyAll();
}

function updateAll() { updateLotto(); updatePension(); beautifyAll(); }

function installTriggers() {
  var ss = SpreadsheetApp.getActive();
  // 트리거 시각은 스크립트 표준시 기준 — 한국 시간과 어긋나지 않도록 고정
  if (ss.getSpreadsheetTimeZone() !== 'Asia/Seoul') ss.setSpreadsheetTimeZone('Asia/Seoul');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklyJob') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyJob').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(22).create();  // 로또 추첨(20:35) 후
  ScriptApp.newTrigger('weeklyJob').timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(22).create();  // 연금 추첨(19:05) 후
  ss.toast('토 22시 / 목 22시 자동실행 설치 완료 (Asia/Seoul)');
}

// ───────────────────────── 유틸 ─────────────────────────

function sheet(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers) sh.appendRow(headers);
  return sh;
}

function toast_(msg) { SpreadsheetApp.getActive().toast(msg); }

// "20260711" 또는 Date → Date
function ymd_(v) {
  if (v instanceof Date) return v;
  var s = String(v).replace(/\D/g, '');
  if (s.length !== 8) return null;
  return new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8)));
}

// 다음 추첨일 (로또=토, 연금=목). 당일이면 그 주로 간주
function nextDraw_(dow) {
  var d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + ((dow - d.getDay() + 7) % 7));
  return d;
}

function getJson(url) {
  var res;
  try {
    res = UrlFetchApp.fetch(url, FETCH_OPT);
  } catch (e) {
    throw new Error(
      '동행복권 접속 실패("사용할 수 없는 주소"): 동행복권이 해외 IP를 차단하므로 ' +
      'Apps Script에서는 직접 수집이 불가합니다. 파일 상단 MIRROR_* 상수의 미러 CSV를 사용하세요. ' +
      '원인: ' + e.message);
  }
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ' ' + url);
  return JSON.parse(res.getContentText());
}

function getCsv_(url) {
  var res = UrlFetchApp.fetch(url, FETCH_OPT);
  if (res.getResponseCode() !== 200) {
    throw new Error('미러 CSV 응답 HTTP ' + res.getResponseCode() + ' — 저장소가 공개 상태인지 확인: ' + url);
  }
  return Utilities.parseCsv(res.getContentText()).slice(1);
}

// ───────────────────────── 데이터 수집 ─────────────────────────

function updateLotto() {
  var sh = sheet(SH_LOTTO, ['회차', '추첨일', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', '보너스',
                            '1등당첨자', '1등금액', '판매액']);
  var last = sh.getLastRow() > 1 ? Number(sh.getRange(sh.getLastRow(), 1).getValue()) : 0;

  if (MIRROR_LOTTO_CSV) {
    var rows = getCsv_(MIRROR_LOTTO_CSV)
      .filter(function (r) { return r.length >= 12 && Number(r[0]) > last; })
      .map(function (r) { return [Number(r[0]), ymd_(r[1])].concat(r.slice(2, 12).map(Number)); })
      .sort(function (a, b) { return a[0] - b[0]; });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    toast_('로또 ' + rows.length + '회 추가 (미러 CSV)');
    return;
  }

  var est = 1 + Math.floor((Date.now() - new Date(2002, 11, 7).getTime()) / (7 * 864e5));
  var latest = 0;
  for (var e = est + 1; e > est - 6 && !latest; e--) {
    var probe = fetchLottoWindow_(e);
    if (probe.length) latest = Math.max.apply(null, probe.map(function (r) { return r[0]; }));
  }
  if (!latest) throw new Error('로또 API 응답 없음 — dhlottery 개편 여부 확인');
  if (latest <= last) { toast_('로또: 신규 회차 없음(' + last + '회)'); return; }

  var buf = {};
  for (var t = last + 1; t <= latest; t += 9) {
    fetchLottoWindow_(Math.min(latest, t + 4)).forEach(function (r) { buf[r[0]] = r; });
    Utilities.sleep(120);
  }
  for (var k = last + 1; k <= latest; k++) {
    if (!buf[k]) fetchLottoWindow_(k).forEach(function (r) { buf[r[0]] = r; });
  }
  var out = [];
  for (var i = last + 1; i <= latest; i++) if (buf[i]) out.push(buf[i]);
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  toast_('로또 ' + out.length + '회 추가 (최신 ' + latest + '회)');
}

function fetchLottoWindow_(epsd) {
  var j = getJson(BASE + '/lt645/selectPstLt645InfoNew.do?srchDir=center&srchLtEpsd=' + epsd);
  return (j.data && j.data.list || []).map(function (d) {
    return [d.ltEpsd, ymd_(d.ltRflYmd), d.tm1WnNo, d.tm2WnNo, d.tm3WnNo, d.tm4WnNo, d.tm5WnNo,
            d.tm6WnNo, d.bnsWnNo, d.rnk1WnNope, d.rnk1WnAmt, d.rlvtEpsdSumNtslAmt];
  }).sort(function (a, b) { return a[0] - b[0]; });
}

function updatePension() {
  var sh = sheet(SH_PENSION, ['회차', '추첨일', '조', '번호', '보너스']);
  var last = sh.getLastRow() > 1 ? Number(sh.getRange(sh.getLastRow(), 1).getValue()) : 0;

  if (MIRROR_PENSION_CSV) {
    var rows = getCsv_(MIRROR_PENSION_CSV)
      .filter(function (r) { return r.length >= 5 && Number(r[0]) > last; })
      .map(function (r) { return [Number(r[0]), ymd_(r[1]), Number(r[2]), pad6_(r[3]), pad6_(r[4])]; })
      .sort(function (a, b) { return a[0] - b[0]; });
    if (rows.length) {
      var at = sh.getLastRow() + 1;
      sh.getRange(at, 1, rows.length, 5).setValues(rows);
      sh.getRange(at, 4, rows.length, 2).setNumberFormat('@');  // 앞자리 0 보존
    }
    toast_('연금 ' + rows.length + '회 추가 (미러 CSV)');
    return;
  }
  var j = getJson(BASE + '/pt720/selectPstPt720WnList.do');
  var rows2 = j.data.result
    .filter(function (d) { return d.psltEpsd > last; })
    .sort(function (a, b) { return a.psltEpsd - b.psltEpsd; })
    .map(function (d) { return [d.psltEpsd, ymd_(d.psltRflYmd), Number(d.wnBndNo), pad6_(d.wnRnkVl), pad6_(d.bnsRnkVl)]; });
  if (rows2.length) {
    var at2 = sh.getLastRow() + 1;
    sh.getRange(at2, 1, rows2.length, 5).setValues(rows2);
    sh.getRange(at2, 4, rows2.length, 2).setNumberFormat('@');
  }
  toast_('연금 ' + rows2.length + '회 추가');
}

function pad6_(v) {
  var s = String(v).replace(/\D/g, '');
  while (s.length < 6) s = '0' + s;
  return s;
}

// ───────────────────────── 통계 ─────────────────────────

function readLotto_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH_LOTTO);
  if (!sh || sh.getLastRow() < 2) throw new Error('먼저 [데이터 갱신]을 실행하세요');
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
}

function buildStats() {
  var data = readLotto_();
  var freq = {}, rfreq = {}, lastSeen = {};
  var recentFrom = Math.max(0, data.length - 52);
  data.forEach(function (r, idx) {
    for (var c = 2; c <= 7; c++) {
      var n = r[c];
      freq[n] = (freq[n] || 0) + 1;
      lastSeen[n] = r[0];
      if (idx >= recentFrom) rfreq[n] = (rfreq[n] || 0) + 1;
    }
  });
  var latest = data[data.length - 1][0];
  var sh = sheet(SH_STATS);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['번호', '역대출현', '최근52회', '미출현회차']]);
  var rows = [];
  for (var n = 1; n <= 45; n++) rows.push([n, freq[n] || 0, rfreq[n] || 0, latest - (lastSeen[n] || 0)]);
  sh.getRange(2, 1, 45, 4).setValues(rows);

  var sums = data.map(function (r) { return r[2] + r[3] + r[4] + r[5] + r[6] + r[7]; })
                 .sort(function (a, b) { return a - b; });
  sh.getRange(1, 6, 6, 2).setValues([
    ['기준 회차', latest + '회'],
    ['갱신 시각', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')],
    ['합계 5% 하한', sums[Math.floor(sums.length * 0.05)]],
    ['합계 95% 상한', sums[Math.floor(sums.length * 0.95)]],
    ['평균 합계', Math.round(sums.reduce(function (a, b) { return a + b; }, 0) / sums.length)],
    ['비고', '출현 빈도는 균등분포와 구별 불가 — 참고용']
  ]);
  formatStats_(sh);
}

// ───────────────────────── 추천 생성 ─────────────────────────

function features_(c) {
  var s = c.slice().sort(function (a, b) { return a - b; });
  var consec = 0, odd = 0, low = 0, le31 = 0, dec = {}, sum = 0;
  for (var i = 0; i < 6; i++) {
    sum += s[i];
    if (s[i] % 2) odd++;
    if (s[i] <= 22) low++;
    if (s[i] <= 31) le31++;
    dec[Math.floor((s[i] - 1) / 10)] = 1;
    if (i && s[i] - s[i - 1] === 1) consec++;
  }
  return { sum: sum, odd: odd, low: low, le31: le31, consec: consec, decades: Object.keys(dec).length };
}

function buildPicks() { buildPicks_(false); }
function buildPicksForce() { buildPicks_(true); }

function buildPicks_(force) {
  var data = readLotto_();
  var latest = Number(data[data.length - 1][0]);
  var target = latest + 1;

  var sh = sheet(SH_PICK, PICK_HEAD);
  if (!force && hasTarget_(sh, 2, target)) {
    toast_(target + '회 추천이 이미 있습니다. 다시 만들려면 [추천 재생성]을 쓰세요.');
  } else {
    var lastDraw = data[data.length - 1].slice(2, 8);
    var histKeys = {};
    data.forEach(function (r) {
      histKeys[r.slice(2, 8).sort(function (a, b) { return a - b; }).join(',')] = 1;
    });
    var sums = data.map(function (r) { return r[2] + r[3] + r[4] + r[5] + r[6] + r[7]; })
                   .sort(function (a, b) { return a - b; });
    var lo = sums[Math.floor(sums.length * 0.05)], hi = sums[Math.floor(sums.length * 0.95)];

    var cands = [], guard = 0;
    while (cands.length < 4000 && guard++ < 200000) {
      var pool = [];
      for (var n = 1; n <= 45; n++) pool.push(n);
      var c = [];
      for (var k = 0; k < 6; k++) c.push(pool.splice(Math.floor(Math.random() * pool.length), 1)[0]);
      c.sort(function (a, b) { return a - b; });
      var f = features_(c);
      if (f.sum < lo || f.sum > hi) continue;
      if (f.odd < 2 || f.odd > 4) continue;
      if (f.low < 2 || f.low > 4) continue;
      if (f.consec > 1) continue;
      if (f.decades < 3) continue;
      if (f.le31 === 6) continue;                 // 전원 생일범위 → 분할 위험
      if (histKeys[c.join(',')]) continue;        // 역대 1등 조합 배제
      if (c.filter(function (x) { return lastDraw.indexOf(x) >= 0; }).length >= 4) continue;
      var pop = 0.039 * f.le31 - 0.019 * f.consec + 0.011 * f.odd + 0.010 * f.decades;
      cands.push({ c: c, f: f, pop: pop });
    }
    cands.sort(function (a, b) { return a.pop - b.pop; });
    var top = cands.slice(0, Math.max(200, Math.floor(cands.length / 10)));
    for (var i2 = top.length - 1; i2 > 0; i2--) {
      var j2 = Math.floor(Math.random() * (i2 + 1));
      var tmp = top[i2]; top[i2] = top[j2]; top[j2] = tmp;
    }
    var picked = [];
    for (var t = 0; t < top.length && picked.length < 5; t++) {
      var ok = picked.every(function (p) {
        return top[t].c.filter(function (x) { return p.c.indexOf(x) >= 0; }).length <= 2;
      });
      if (ok) picked.push(top[t]);
    }

    var now = new Date(), drawDate = nextDraw_(6);  // 토요일
    var out = picked.map(function (p, idx) {
      return [now, target, drawDate, String.fromCharCode(65 + idx)]
        .concat(p.c, [p.f.sum, p.f.odd, Math.round(p.pop * 100) / 100, '', '']);
    });
    sh.getRange(sh.getLastRow() + 1, 1, out.length, PICK_HEAD.length).setValues(out);
  }

  buildPensionPicks_(force);
  gradeResults();
  beautifyAll();
  toast_('추천 생성 완료 (로또 ' + target + '회)');
}

// 연금복권 추천 — 품절 대비 순위 리스트 PENSION_N건
// 조·번호 조합은 전국에 1장뿐이라 이미 팔린 것은 살 수 없다. 앞 순위 품절 시 다음 순위로 구매.
// 조는 1~5조에 라운드로빈으로 고르게 배정 — 한 조가 통째로 매진돼도 대안이 남는다.
// (실제 당첨 확률은 모든 조합이 동일 — 자리별 가중치는 재미 요소)
function buildPensionPicks_(force) {
  var pSh = SpreadsheetApp.getActive().getSheetByName(SH_PENSION);
  if (!pSh || pSh.getLastRow() < 2) return;
  var pLatest = Number(pSh.getRange(pSh.getLastRow(), 1).getValue());
  var target = pLatest + 1;
  var sh = sheet(SH_PPICK, PPICK_HEAD);
  if (!force && hasTarget_(sh, 2, target)) return;

  var pData = pSh.getRange(2, 4, pSh.getLastRow() - 1, 1).getValues();  // 번호
  var posCnt = [];
  for (var d = 0; d < 6; d++) posCnt.push([0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
  pData.forEach(function (r) {
    var s = pad6_(r[0]);
    for (var d2 = 0; d2 < 6; d2++) posCnt[d2][Number(s[d2])]++;
  });
  var posW = posCnt.map(function (cnt) {
    var mean = cnt.reduce(function (a, b) { return a + b; }, 0) / 10;
    return cnt.map(function (x) { return Math.max(1, 2 * mean - x); });
  });
  var wPick = function (w) {
    var tot = w.reduce(function (a, b) { return a + b; }, 0), r = Math.random() * tot;
    for (var i = 0; i < w.length; i++) { r -= w[i]; if (r <= 0) return i; }
    return w.length - 1;
  };

  // 번호(6자리)만 점수순으로 뽑고, 조는 뒤에서 균등 배정한다
  var seen = {}, list = [];
  for (var g = 0; g < Math.max(2000, PENSION_N * 100); g++) {
    var digits = [], score = 0;
    for (var d3 = 0; d3 < 6; d3++) { var dg = wPick(posW[d3]); digits.push(dg); score += posW[d3][dg]; }
    var key = digits.join('');
    if (seen[key]) continue;
    seen[key] = 1;
    list.push({ num: key, score: score });
  }
  list.sort(function (x, y) { return y.score - x.score; });

  var picked = [];
  list.forEach(function (cand) {   // 후보 간 자리 일치 ≤3
    if (picked.length >= PENSION_N) return;
    var tooClose = picked.some(function (q) {
      var same = 0;
      for (var i = 0; i < 6; i++) if (q.num[i] === cand.num[i]) same++;
      return same > 3;
    });
    if (!tooClose) picked.push(cand);
  });
  list.forEach(function (cand) {   // 모자라면 제약 완화해 채움
    if (picked.length < PENSION_N && picked.indexOf(cand) < 0) picked.push(cand);
  });

  // 조 라운드로빈 배정 (5개 단위로 1~5조 한 번씩, 순서는 매회 섞어 편향 방지)
  var jos = [];
  while (jos.length < picked.length) {
    var bag = [1, 2, 3, 4, 5];
    for (var b = bag.length - 1; b > 0; b--) {
      var s = Math.floor(Math.random() * (b + 1)), tp = bag[b]; bag[b] = bag[s]; bag[s] = tp;
    }
    jos = jos.concat(bag);
  }

  var now = new Date(), drawDate = nextDraw_(4);  // 목요일
  var out = picked.map(function (p, i) {
    return [now, target, drawDate, (i + 1) + '순위', jos[i], p.num, '', ''];
  });
  var at = sh.getLastRow() + 1;
  sh.getRange(at, 1, out.length, PPICK_HEAD.length).setValues(out);
  sh.getRange(at, PPICK_NUM, out.length, 1).setNumberFormat('@');
}

function hasTarget_(sh, col, target) {
  if (sh.getLastRow() < 2) return false;
  var vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  return vals.some(function (r) { return Number(r[0]) === target; });
}

// ───────────────────────── 결과 채점 ─────────────────────────

function gradeResults() {
  var a = gradeLotto_(), b = gradePension_();
  buildScore_();
  toast_('채점 완료 — 로또 ' + a + '건 / 연금 ' + b + '건');
}

function gradeLotto_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SH_PICK), hSh = ss.getSheetByName(SH_LOTTO);
  if (!sh || sh.getLastRow() < 2 || !hSh || hSh.getLastRow() < 2) return 0;

  var res = {};
  hSh.getRange(2, 1, hSh.getLastRow() - 1, 9).getValues().forEach(function (r) {
    res[Number(r[0])] = { date: r[1], nums: r.slice(2, 8).map(Number), bonus: Number(r[8]) };
  });

  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, 1, n, PICK_HEAD.length).getValues();
  var bgs = sh.getRange(2, PICK_N1, n, 6).getBackgrounds();
  var fgs = sh.getRange(2, PICK_N1, n, 6).getFontColors();
  var hits = [], ranks = [], dates = [], graded = 0;

  for (var i = 0; i < n; i++) {
    var target = Number(vals[i][1]);
    var r = res[target];
    if (!r) {                       // 아직 추첨 전
      hits.push([vals[i][PICK_HIT - 1] || '대기']);
      ranks.push([vals[i][PICK_RANK - 1] || '추첨 전']);
      dates.push([vals[i][2]]);
      continue;
    }
    var m = 0, bonusHit = false;
    for (var c = 0; c < 6; c++) {
      var num = Number(vals[i][PICK_N1 - 1 + c]);
      if (r.nums.indexOf(num) >= 0) {
        m++; bgs[i][c] = C_HIT_BG; fgs[i][c] = C_HIT_FG;
      } else if (num === r.bonus) {
        bonusHit = true; bgs[i][c] = C_BONUS_BG; fgs[i][c] = C_BONUS_FG;
      } else {
        bgs[i][c] = C_MISS_BG; fgs[i][c] = '#000000';
      }
    }
    var rank = m === 6 ? '1등' : (m === 5 && bonusHit) ? '2등' : m === 5 ? '3등'
             : m === 4 ? '4등' : m === 3 ? '5등' : '낙첨';
    hits.push([m + '개' + (bonusHit ? '+보너스' : '')]);
    ranks.push([rank]);
    dates.push([r.date]);
    graded++;
  }
  sh.getRange(2, PICK_N1, n, 6).setBackgrounds(bgs).setFontColors(fgs);
  sh.getRange(2, PICK_HIT, n, 1).setValues(hits);
  sh.getRange(2, PICK_RANK, n, 1).setValues(ranks);
  sh.getRange(2, 3, n, 1).setValues(dates);
  return graded;
}

// 연금복권 등수: 뒤에서부터 연속 일치한 자리수로 결정
//   1등 = 조+6자리 전부 / 2등 = 조 무관 6자리 / 3등 뒤5 / 4등 뒤4 / 5등 뒤3 / 6등 뒤2 / 7등 뒤1
function gradePension_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SH_PPICK), hSh = ss.getSheetByName(SH_PENSION);
  if (!sh || sh.getLastRow() < 2 || !hSh || hSh.getLastRow() < 2) return 0;

  var res = {};
  hSh.getRange(2, 1, hSh.getLastRow() - 1, 5).getValues().forEach(function (r) {
    res[Number(r[0])] = { date: r[1], jo: Number(r[2]), num: pad6_(r[3]), bonus: pad6_(r[4]) };
  });

  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, 1, n, PPICK_HEAD.length).getValues();
  var rts = [], hits = [], ranks = [], dates = [], graded = 0;
  var hitStyle = SpreadsheetApp.newTextStyle().setForegroundColor(C_HIT_FG).setBold(true).build();
  var missStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setBold(false).build();

  for (var i = 0; i < n; i++) {
    var target = Number(vals[i][1]);
    var num = pad6_(vals[i][PPICK_NUM - 1]);
    var r = res[target];
    if (!r) {
      rts.push([SpreadsheetApp.newRichTextValue().setText(num)
                 .setTextStyle(0, 6, missStyle).build()]);
      hits.push([vals[i][PPICK_HIT - 1] || '대기']);
      ranks.push([vals[i][PPICK_RANK - 1] || '추첨 전']);
      dates.push([vals[i][2]]);
      continue;
    }
    var m = 0;
    while (m < 6 && num[5 - m] === r.num[5 - m]) m++;
    var joHit = Number(vals[i][4]) === r.jo;
    var rank = (m === 6 && joHit) ? '1등' : m === 6 ? '2등' : m === 5 ? '3등' : m === 4 ? '4등'
             : m === 3 ? '5등' : m === 2 ? '6등' : m === 1 ? '7등' : '낙첨';
    if (num === r.bonus) rank += '(보너스 일치)';
    var rt = SpreadsheetApp.newRichTextValue().setText(num).setTextStyle(0, 6, missStyle);
    if (m > 0) rt.setTextStyle(6 - m, 6, hitStyle);
    rts.push([rt.build()]);
    hits.push([m + '자리' + (joHit ? '+조일치' : '')]);
    ranks.push([rank]);
    dates.push([r.date]);
    graded++;
  }
  sh.getRange(2, PPICK_NUM, n, 1).setRichTextValues(rts);
  sh.getRange(2, PPICK_HIT, n, 1).setValues(hits);
  sh.getRange(2, PPICK_RANK, n, 1).setValues(ranks);
  sh.getRange(2, 3, n, 1).setValues(dates);
  return graded;
}

// ───────────────────────── 성적표 ─────────────────────────

function comb_(n, k) {
  if (k < 0 || k > n) return 0;
  var r = 1;
  for (var i = 0; i < k; i++) r = r * (n - i) / (i + 1);
  return r;
}

function buildScore_() {
  var ss = SpreadsheetApp.getActive();
  var sh = sheet(SH_SCORE);
  var pk = ss.getSheetByName(SH_PICK);
  sh.clear();

  var rows = [['로또 추천 성적 (채점 완료분만)', '']];
  var total = 0, sumHit = 0, dist = [0, 0, 0, 0, 0, 0, 0], rankCnt = {};
  if (pk && pk.getLastRow() > 1) {
    pk.getRange(2, PICK_HIT, pk.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var mm = String(r[0]).match(/^(\d)개/);
      if (!mm) return;
      total++; var k = Number(mm[1]); dist[k]++; sumHit += k;
      rankCnt[r[1]] = (rankCnt[r[1]] || 0) + 1;
    });
  }
  rows.push(['채점한 세트 수', total]);
  rows.push(['평균 적중 개수', total ? Math.round(sumHit / total * 100) / 100 : 0]);
  rows.push(['이론 기대 적중', 0.8]);
  rows.push(['', '']);
  rows.push(['맞은 개수', '실제 건수', '이론 기대 건수']);
  for (var k = 0; k <= 6; k++) {
    var p = comb_(6, k) * comb_(39, 6 - k) / comb_(45, 6);
    rows.push([k + '개', dist[k], Math.round(p * total * 100) / 100]);
  }
  rows.push(['', '']);
  rows.push(['등수', '건수', '']);
  ['1등', '2등', '3등', '4등', '5등', '낙첨'].forEach(function (rk) {
    rows.push([rk, rankCnt[rk] || 0, '']);
  });

  var ppk = ss.getSheetByName(SH_PPICK);
  rows.push(['', '']);
  rows.push(['연금복권 성적 (채점 완료분만)', '']);
  var pTotal = 0, pRank = {};
  if (ppk && ppk.getLastRow() > 1) {
    ppk.getRange(2, PPICK_HIT, ppk.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (!/^\d자리/.test(String(r[0]))) return;
      pTotal++;
      var rk = String(r[1]).replace('(보너스 일치)', '');
      pRank[rk] = (pRank[rk] || 0) + 1;
    });
  }
  rows.push(['채점한 후보 수', pTotal]);
  ['1등', '2등', '3등', '4등', '5등', '6등', '7등', '낙첨'].forEach(function (rk) {
    rows.push([rk, pRank[rk] || 0, '']);
  });
  rows.push(['', '']);
  rows.push(['※ 각 추첨은 독립시행 — 성적이 좋아도 다음 회차 확률은 변하지 않는다.', '']);

  var width = 3;
  var padded = rows.map(function (r) { while (r.length < width) r.push(''); return r.slice(0, width); });
  sh.getRange(1, 1, padded.length, width).setValues(padded);
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG);
  sh.getRange(6, 1, 1, width).setFontWeight('bold').setBackground('#e8eef7');
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 130);
}

// ───────────────────────── 서식 ─────────────────────────

function beautifyAll() {
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SH_LOTTO)) formatHistory_(ss.getSheetByName(SH_LOTTO), 12, 2, 1);
  if (ss.getSheetByName(SH_PENSION)) formatHistory_(ss.getSheetByName(SH_PENSION), 5, 2, 1);
  if (ss.getSheetByName(SH_STATS)) formatStats_(ss.getSheetByName(SH_STATS));
  if (ss.getSheetByName(SH_PICK)) formatPicks_(ss.getSheetByName(SH_PICK), PICK_HEAD.length, 2, 3);
  if (ss.getSheetByName(SH_PPICK)) formatPicks_(ss.getSheetByName(SH_PPICK), PPICK_HEAD.length, 2, 3);
  toast_('서식 적용 완료');
}

function headerStyle_(sh, nCols) {
  sh.getRange(1, 1, 1, nCols)
    .setBackground(C_HEAD_BG).setFontColor(C_HEAD_FG).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 30);
  sh.setFrozenRows(1);
}

// 이력 시트: 날짜 서식 + 연도 경계 구분선 + 숫자 천단위
function formatHistory_(sh, nCols, dateCol, epsdCol) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  headerStyle_(sh, nCols);
  sh.getRange(2, dateCol, n, 1).setNumberFormat('yyyy-mm-dd(ddd)').setHorizontalAlignment('center');
  sh.getRange(2, epsdCol, n, 1).setHorizontalAlignment('center');
  sh.getRange(2, 1, n, nCols).setFontFamily('Roboto Mono').setFontSize(10);
  if (nCols >= 12) {   // 로또: 금액 열 천단위
    sh.getRange(2, 10, n, 3).setNumberFormat('#,##0');
    sh.getRange(2, 3, n, 7).setHorizontalAlignment('center');
  } else {             // 연금: 조·번호 가운데
    sh.getRange(2, 3, n, 3).setHorizontalAlignment('center');
  }
  // 연도가 바뀌는 행 위에 굵은 구분선 (날짜 구간을 눈으로 나눠 보기 위함)
  sh.getRange(2, 1, n, nCols).setBorder(null, null, null, null, false, false);
  var dates = sh.getRange(2, dateCol, n, 1).getValues();
  var prevYear = null;
  for (var i = 0; i < n; i++) {
    var d = dates[i][0];
    if (!(d instanceof Date)) continue;
    var y = d.getFullYear();
    if (prevYear !== null && y !== prevYear) {
      sh.getRange(i + 2, 1, 1, nCols)
        .setBorder(true, null, null, null, null, null, '#1f3864', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    prevYear = y;
  }
  sh.autoResizeColumns(1, nCols);
}

function formatStats_(sh) {
  headerStyle_(sh, 4);
  sh.getRange(2, 1, 45, 4).setHorizontalAlignment('center').setFontFamily('Roboto Mono');
  sh.getRange(1, 6, 6, 1).setFontWeight('bold');
  sh.setColumnWidth(5, 24); sh.setColumnWidth(6, 150); sh.setColumnWidth(7, 220);
  // 출현 빈도 색조 (많을수록 진하게)
  sh.getRange(2, 2, 45, 1).clearFormat();
  var rules = sh.getConditionalFormatRules().filter(function (r) {
    return r.getRanges().every(function (rg) { return rg.getSheet().getName() !== sh.getName(); });
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.MIN, '')
    .setGradientMaxpointWithValue('#4a86c8', SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([sh.getRange(2, 2, 45, 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.MIN, '')
    .setGradientMaxpointWithValue('#f6b26b', SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([sh.getRange(2, 4, 45, 1)]).build());
  sh.setConditionalFormatRules(rules);
  sh.getRange(2, 2, 45, 1).setNumberFormat('0');
}

// 추천 시트: 회차 그룹마다 배경 교대 + 경계 구분선, 등수 강조
function formatPicks_(sh, nCols, groupCol, dateCol) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  headerStyle_(sh, nCols);
  var vals = sh.getRange(2, 1, n, nCols).getValues();
  sh.getRange(2, 1, n, 1).setNumberFormat('yyyy-mm-dd HH:mm');
  sh.getRange(2, dateCol, n, 1).setNumberFormat('yyyy-mm-dd(ddd)');
  sh.getRange(2, 1, n, nCols).setFontFamily('Roboto Mono').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(2, 1, n, 1).setHorizontalAlignment('left');

  // 회차 그룹: 교대 배경(번호 열 제외 — 적중 색을 덮지 않도록) + 그룹 경계선
  var isLotto = (nCols === PICK_HEAD.length);
  var numStart = isLotto ? PICK_N1 : PPICK_NUM;
  var numCount = isLotto ? 6 : 1;
  var bgCols = [];
  for (var c = 1; c <= nCols; c++) {
    if (c >= numStart && c < numStart + numCount) continue;
    bgCols.push(c);
  }
  var groupIdx = -1, prevKey = null;
  for (var i = 0; i < n; i++) {
    var key = String(vals[i][groupCol - 1]);
    if (key !== prevKey) {
      groupIdx++;
      if (i > 0) {
        sh.getRange(i + 2, 1, 1, nCols).setBorder(
          true, null, null, null, null, null, '#1f3864', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      prevKey = key;
    }
    var bg = (groupIdx % 2 === 0) ? C_GROUP_A : C_GROUP_B;
    bgCols.forEach(function (c) { sh.getRange(i + 2, c).setBackground(bg); });
  }

  // 등수 열 강조 + 미추첨 회색
  var rankCol = isLotto ? PICK_RANK : PPICK_RANK;
  var ranks = sh.getRange(2, rankCol, n, 1).getValues();
  var fg = [], fw = [];
  for (var r2 = 0; r2 < n; r2++) {
    var v = String(ranks[r2][0]);
    var win = /^[1-7]등/.test(v);
    fg.push([win ? C_HIT_FG : (v === '추첨 전' ? C_PENDING : '#000000')]);
    fw.push([win ? 'bold' : 'normal']);
  }
  sh.getRange(2, rankCol, n, 1).setFontColors(fg).setFontWeights(fw);
  sh.getRange(2, numStart, n, numCount).setFontWeight('bold').setFontSize(11);
  sh.autoResizeColumns(1, nCols);
  sh.setColumnWidth(1, 130);
}

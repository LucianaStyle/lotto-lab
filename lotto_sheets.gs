/**
 * 로또 6/45 + 연금복권 720+ 자동 분석 — Google Sheets (Apps Script)
 *
 * 설치:
 *  1. 새 구글 시트 → 확장 프로그램 > Apps Script → 이 파일 전체를 붙여넣고 저장
 *  2. 시트로 돌아와 새로고침 → 메뉴에 [복권분석]이 생김
 *  3. [복권분석 > 전체 업데이트+추천] 최초 1회 실행 (권한 승인 필요, 2~3분 소요)
 *  4. [복권분석 > 주간 자동실행 설치] → 매주 토 22시(로또)·목 22시(연금) 자동 갱신
 *
 * 파이썬 버전(lotto_lab.py)과 동일한 로직:
 *  - 프로파일 필터(합계 중앙 90% 구간, 홀짝 2~4, 저구간 2~4, 연속쌍 ≤1 등)
 *  - 분할 회피: 생일범위(≤31) 숫자·연속수·정형 패턴을 피해 당첨 시 독식 확률↑
 *
 * ⚠️ 알려진 제약 — "사용할 수 없는 주소" 예외:
 *   동행복권은 해외 IP 접속을 차단한다. Apps Script(UrlFetchApp)는 구글 해외 서버에서
 *   요청이 나가므로 dhlottery 직접 수집이 불가능하다. 해결책 두 가지:
 *   (A) 권장: 로컬 PC의 lotto_lab.py 자동 갱신(작업 스케줄러 lotto-lab-update)을 쓰고,
 *       이 시트는 분석/기록용으로만 사용 — data/*.csv를 시트로 가져오기(Import).
 *   (B) 로컬 갱신본 CSV를 한국 외에서도 접근 가능한 URL(GitHub raw 등)에 올리고
 *       아래 MIRROR_* 상수에 그 주소를 넣으면 시트가 미러에서 자동 수집한다.
 */

var MIRROR_LOTTO_CSV = 'https://raw.githubusercontent.com/LucianaStyle/lotto-lab/main/data/lotto_history.csv';
var MIRROR_PENSION_CSV = 'https://raw.githubusercontent.com/LucianaStyle/lotto-lab/main/data/pension_history.csv';

var BASE = 'https://www.dhlottery.co.kr';
var FETCH_OPT = {
  headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
  muteHttpExceptions: true
};

// ───────────────────────── 메뉴 ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('복권분석')
    .addItem('전체 업데이트+추천', 'weeklyJob')
    .addItem('로또 데이터 갱신', 'updateLotto')
    .addItem('연금 데이터 갱신', 'updatePension')
    .addItem('통계 재계산', 'buildStats')
    .addItem('추천 번호 생성', 'buildPicks')
    .addSeparator()
    .addItem('주간 자동실행 설치', 'installTriggers')
    .addToUi();
}

function weeklyJob() {
  updateLotto();
  updatePension();
  buildStats();
  buildPicks();
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('weeklyJob').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(22).create();  // 로또 추첨(20:35) 후
  ScriptApp.newTrigger('weeklyJob').timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(22).create();  // 연금 추첨(19:05) 후
  SpreadsheetApp.getActive().toast('토 22시 / 목 22시 자동실행 설치 완료');
}

// ───────────────────────── 유틸 ─────────────────────────

function sheet(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers) sh.appendRow(headers);
  return sh;
}

function getJson(url) {
  var res;
  try {
    res = UrlFetchApp.fetch(url, FETCH_OPT);
  } catch (e) {
    throw new Error(
      '동행복권 접속 실패("사용할 수 없는 주소"): 동행복권이 해외 IP를 차단하므로 ' +
      'Apps Script에서는 직접 수집이 불가합니다. 로컬 자동 갱신(lotto-lab-update 작업)을 쓰거나 ' +
      '파일 상단 MIRROR_* 상수에 미러 CSV 주소를 설정하세요. 원인: ' + e.message);
  }
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ' ' + url);
  return JSON.parse(res.getContentText());
}

// 미러 CSV(쉼표 구분, 헤더 1줄) → 2차원 배열
function getCsv_(url) {
  var res = UrlFetchApp.fetch(url, FETCH_OPT);
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ' ' + url);
  return Utilities.parseCsv(res.getContentText()).slice(1);
}

// ───────────────────────── 데이터 수집 ─────────────────────────

function updateLotto() {
  var sh = sheet('로또이력', ['회차', '추첨일', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', '보너스', '1등당첨자', '1등금액', '판매액']);
  var last = sh.getLastRow() > 1 ? sh.getRange(sh.getLastRow(), 1).getValue() : 0;

  // 미러 CSV가 설정돼 있으면 그쪽에서 수집 (dhlottery 해외 IP 차단 우회)
  if (MIRROR_LOTTO_CSV) {
    var rows = getCsv_(MIRROR_LOTTO_CSV)
      .map(function (r) { return [Number(r[0]), r[1]].concat(r.slice(2, 12).map(Number)); })
      .filter(function (r) { return r[0] > last; })
      .sort(function (a, b) { return a[0] - b[0]; });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    SpreadsheetApp.getActive().toast('로또 ' + rows.length + '회 추가 (미러 CSV)');
    return;
  }

  // 최신 회차 탐색: 1회=2002-12-07, 미래 회차는 빈 리스트
  var est = 1 + Math.floor((Date.now() - new Date(2002, 11, 7).getTime()) / (7 * 864e5));
  var latest = 0;
  for (var e = est + 1; e > est - 6 && !latest; e--) {
    var rows = fetchLottoWindow_(e);
    if (rows.length) latest = Math.max.apply(null, rows.map(function (r) { return r[0]; }));
  }
  if (!latest) throw new Error('로또 API 응답 없음 — dhlottery 개편 여부 확인');
  if (latest <= last) { SpreadsheetApp.getActive().toast('로또: 신규 회차 없음(' + last + '회)'); return; }

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
  SpreadsheetApp.getActive().toast('로또 ' + out.length + '회 추가 (최신 ' + latest + '회)');
}

function fetchLottoWindow_(epsd) {
  var j = getJson(BASE + '/lt645/selectPstLt645InfoNew.do?srchDir=center&srchLtEpsd=' + epsd);
  return (j.data && j.data.list || []).map(function (d) {
    return [d.ltEpsd, d.ltRflYmd, d.tm1WnNo, d.tm2WnNo, d.tm3WnNo, d.tm4WnNo, d.tm5WnNo, d.tm6WnNo,
            d.bnsWnNo, d.rnk1WnNope, d.rnk1WnAmt, d.rlvtEpsdSumNtslAmt];
  }).sort(function (a, b) { return a[0] - b[0]; });
}

function updatePension() {
  var sh = sheet('연금이력', ['회차', '추첨일', '조', '번호', '보너스']);
  var last = sh.getLastRow() > 1 ? sh.getRange(sh.getLastRow(), 1).getValue() : 0;
  if (MIRROR_PENSION_CSV) {
    var rows = getCsv_(MIRROR_PENSION_CSV)
      .map(function (r) { return [Number(r[0]), r[1], Number(r[2]), "'" + r[3], "'" + r[4]]; })
      .filter(function (r) { return r[0] > last; })
      .sort(function (a, b) { return a[0] - b[0]; });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    SpreadsheetApp.getActive().toast('연금 ' + rows.length + '회 추가 (미러 CSV)');
    return;
  }
  var j = getJson(BASE + '/pt720/selectPstPt720WnList.do');
  var rows = j.data.result
    .filter(function (d) { return d.psltEpsd > last; })
    .sort(function (a, b) { return a.psltEpsd - b.psltEpsd; })
    .map(function (d) { return [d.psltEpsd, d.psltRflYmd, Number(d.wnBndNo), "'" + d.wnRnkVl, "'" + d.bnsRnkVl]; });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  SpreadsheetApp.getActive().toast('연금 ' + rows.length + '회 추가');
}

// ───────────────────────── 통계 ─────────────────────────

function readLotto_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('로또이력');
  if (!sh || sh.getLastRow() < 2) throw new Error('먼저 [로또 데이터 갱신]을 실행하세요');
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
}

function buildStats() {
  var data = readLotto_();
  var freq = {}, rfreq = {}, lastSeen = {};
  var recentFrom = data.length - 52;
  data.forEach(function (r, idx) {
    for (var c = 2; c <= 7; c++) {
      var n = r[c];
      freq[n] = (freq[n] || 0) + 1;
      lastSeen[n] = r[0];
      if (idx >= recentFrom) rfreq[n] = (rfreq[n] || 0) + 1;
    }
  });
  var latest = data[data.length - 1][0];
  var sh = sheet('통계');
  sh.clear();
  sh.appendRow(['번호', '역대출현', '최근52회', '미출현회차', '', '갱신: ' + new Date()]);
  var rows = [];
  for (var n = 1; n <= 45; n++) {
    rows.push([n, freq[n] || 0, rfreq[n] || 0, latest - (lastSeen[n] || 0)]);
  }
  sh.getRange(2, 1, 45, 4).setValues(rows);
  // 조합 프로파일 요약
  var sums = data.map(function (r) { return r[2] + r[3] + r[4] + r[5] + r[6] + r[7]; }).sort(function (a, b) { return a - b; });
  sh.getRange(1, 6, 3, 2).setValues([
    ['합계 5% 하한', sums[Math.floor(sums.length * 0.05)]],
    ['합계 95% 상한', sums[Math.floor(sums.length * 0.95)]],
    ['평균 합계', Math.round(sums.reduce(function (a, b) { return a + b; }, 0) / sums.length)]
  ]);
}

// ───────────────────────── 추천 생성 ─────────────────────────

function features_(c) {
  var s = c.slice().sort(function (a, b) { return a - b; });
  var consec = 0, odd = 0, low = 0, le31 = 0, dec = {};
  var sum = 0;
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

function buildPicks() {
  var data = readLotto_();
  var latest = data[data.length - 1][0];
  var lastDraw = data[data.length - 1].slice(2, 8);
  var histKeys = {};
  data.forEach(function (r) { histKeys[r.slice(2, 8).sort(function (a, b) { return a - b; }).join(',')] = 1; });
  var sums = data.map(function (r) { return r[2] + r[3] + r[4] + r[5] + r[6] + r[7]; }).sort(function (a, b) { return a - b; });
  var lo = sums[Math.floor(sums.length * 0.05)], hi = sums[Math.floor(sums.length * 0.95)];

  var cands = [];
  var guard = 0;
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
    if (f.le31 === 6) continue;                    // 전원 생일범위 → 분할 위험
    if (histKeys[c.join(',')]) continue;           // 역대 1등 조합 배제
    var overlap = c.filter(function (x) { return lastDraw.indexOf(x) >= 0; }).length;
    if (overlap >= 4) continue;
    // 인기도 근사(파이썬 회귀 계수의 부호 반영): 생일범위·연속수가 많을수록 대중적
    var pop = 0.039 * f.le31 - 0.019 * f.consec + 0.011 * f.odd + 0.010 * f.decades;
    cands.push({ c: c, f: f, pop: pop });
  }
  cands.sort(function (a, b) { return a.pop - b.pop; });
  var top = cands.slice(0, Math.max(200, Math.floor(cands.length / 10)));
  // 셔플 후 세트 간 중복 ≤2로 5세트 선별
  for (var i = top.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = top[i]; top[i] = top[j]; top[j] = tmp;
  }
  var picked = [];
  for (var t = 0; t < top.length && picked.length < 5; t++) {
    var ok = picked.every(function (p) {
      return top[t].c.filter(function (x) { return p.c.indexOf(x) >= 0; }).length <= 2;
    });
    if (ok) picked.push(top[t]);
  }

  var sh = sheet('추천', ['생성시각', '대상회차', '세트', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', '합계', '홀수', '인기도점수']);
  var now = new Date();
  picked.forEach(function (p, idx) {
    sh.appendRow([now, latest + 1, String.fromCharCode(65 + idx)].concat(p.c, [p.f.sum, p.f.odd, Math.round(p.pop * 100) / 100]));
  });

  // 연금복권 추천 — 품절 대비 순위 리스트 6건
  // (조·번호 조합은 1장씩만 판매되므로 앞 순위 품절 시 다음 순위로 구매.
  //  1~5순위는 서로 다른 조로 분산. 실제 당첨 확률은 모든 조합이 동일 — 가중치는 재미 요소)
  var pSh = SpreadsheetApp.getActive().getSheetByName('연금이력');
  if (pSh && pSh.getLastRow() > 1) {
    var pLatest = pSh.getRange(pSh.getLastRow(), 1).getValue();
    var pData = pSh.getRange(2, 3, pSh.getLastRow() - 1, 2).getValues(); // 조, 번호
    // 자리별·조별 미달빈도 가중치 계산
    var posCnt = [], joCnt = [0, 0, 0, 0, 0];
    for (var d = 0; d < 6; d++) { posCnt.push([0,0,0,0,0,0,0,0,0,0]); }
    pData.forEach(function (r) {
      joCnt[Number(r[0]) - 1]++;
      var s = String(r[1]).replace(/\D/g, ''); while (s.length < 6) s = '0' + s;
      for (var d2 = 0; d2 < 6; d2++) posCnt[d2][Number(s[d2])]++;
    });
    var posW = posCnt.map(function (cnt) {
      var mean = cnt.reduce(function (a, b) { return a + b; }, 0) / 10;
      return cnt.map(function (c) { return Math.max(1, 2 * mean - c); });
    });
    var joMean = joCnt.reduce(function (a, b) { return a + b; }, 0) / 5;
    var joW = joCnt.map(function (c) { return Math.max(1, 2 * joMean - c); });
    var wPick = function (w) {
      var tot = w.reduce(function (a, b) { return a + b; }, 0), r = Math.random() * tot;
      for (var i2 = 0; i2 < w.length; i2++) { r -= w[i2]; if (r <= 0) return i2; }
      return w.length - 1;
    };
    // 후보 풀 생성 → 점수순 정렬 → 조 미중복 우선으로 6건 선별
    var pPool = {}, pList = [];
    for (var g = 0; g < 400; g++) {
      var digits = [], score = 0;
      for (var d3 = 0; d3 < 6; d3++) { var dg = wPick(posW[d3]); digits.push(dg); score += posW[d3][dg]; }
      var jo = wPick(joW) + 1;
      var key = jo + '-' + digits.join('');
      if (pPool[key]) continue;
      pPool[key] = 1;
      pList.push({ jo: jo, num: digits.join(''), score: score + joW[jo - 1] });
    }
    pList.sort(function (x, y) { return y.score - x.score; });
    var pPicked = [];
    pList.forEach(function (cand) {
      if (pPicked.length >= 6) return;
      var joUsed = pPicked.some(function (q) { return q.jo === cand.jo; });
      if (pPicked.length < 5 && joUsed) return;
      pPicked.push(cand);
    });
    pList.forEach(function (cand) {
      if (pPicked.length < 6 && pPicked.indexOf(cand) < 0) pPicked.push(cand);
    });
    pPicked.forEach(function (cand, r2) {
      sh.appendRow([now, '연금 ' + (pLatest + 1) + '회', (r2 + 1) + '순위', cand.jo + '조', "'" + cand.num]);
    });
  }
  SpreadsheetApp.getActive().toast('추천 5세트 + 연금 순위 6건 생성 완료');
}

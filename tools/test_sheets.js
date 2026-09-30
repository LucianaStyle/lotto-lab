// lotto_sheets.gs 종단 테스트 — Sheets API 스텁 + 파이썬이 만든 실제 data/*.csv를 미러로 사용
// 실행: node tools/test_sheets.js   (lotto_sheets.gs를 고치면 반드시 실행)
const fs = require('fs');
const vm = require('vm');
const LAB = require('path').resolve(__dirname, '..');

// ── 스텁 ──────────────────────────────────────────────
function makeSheet(name, rows) {
  const data = rows.map(r => r.slice());
  const fmt = { bg: {}, fg: {}, rich: {}, nf: {} };
  const sh = { _data: data, _fmt: fmt, _name: name };
  const grow = (r, c) => {
    while (data.length < r) data.push([]);
    for (const row of data) while (row.length < c) row.push('');
  };
  function range(r, c, nr = 1, nc = 1) {
    let proxy;
    const each = (fn) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) fn(i, j); };
    const api = {
      getSheet: () => sh,
      getValues: () => Array.from({ length: nr }, (_, i) =>
        Array.from({ length: nc }, (_, j) => ((data[r - 1 + i] || [])[c - 1 + j] ?? ''))),
      setValues: v => {
        if (v.length !== nr || v.some(row => row.length !== nc)) {
          throw new Error(`setValues 크기 불일치 ${name}: 범위 ${nr}x${nc}, 값 ${v.length}x${v[0] && v[0].length}`);
        }
        grow(r + nr - 1, c + nc - 1);
        each((i, j) => { data[r - 1 + i][c - 1 + j] = v[i][j]; });
        return proxy;
      },
      getValue: () => ((data[r - 1] || [])[c - 1] ?? ''),
      getBackgrounds: () => Array.from({ length: nr }, (_, i) =>
        Array.from({ length: nc }, (_, j) => fmt.bg[`${r + i},${c + j}`] || '#ffffff')),
      setBackgrounds: v => { each((i, j) => { fmt.bg[`${r + i},${c + j}`] = v[i][j]; }); return proxy; },
      getFontColors: () => Array.from({ length: nr }, (_, i) =>
        Array.from({ length: nc }, (_, j) => fmt.fg[`${r + i},${c + j}`] || '#000000')),
      setFontColors: v => { each((i, j) => { fmt.fg[`${r + i},${c + j}`] = v[i][j]; }); return proxy; },
      setRichTextValues: v => {
        grow(r + nr - 1, c + nc - 1);
        each((i, j) => { fmt.rich[`${r + i},${c + j}`] = v[i][j]; data[r - 1 + i][c - 1 + j] = v[i][j]._text; });
        return proxy;
      },
      setNumberFormat: f => { each((i, j) => { fmt.nf[`${r + i},${c + j}`] = f; }); return proxy; },
    };
    proxy = new Proxy(api, { get: (t, p) => (p in t ? t[p] : () => proxy) });
    return proxy;
  }
  Object.assign(sh, {
    getName: () => sh._name,
    setName: n => { delete sheets[sh._name]; sh._name = n; sheets[n] = sh; return sh; },
    getLastRow: () => data.length,
    getRange: (r, c, nr, nc) => range(r, c, nr, nc),
    clear: () => { data.length = 0; fmt.bg = {}; fmt.fg = {}; fmt.rich = {}; fmt.nf = {}; return sh; },
    clearConditionalFormatRules: () => {}, setConditionalFormatRules: () => {},
    getConditionalFormatRules: () => [],
    setFrozenRows: () => {}, setRowHeight: () => {}, setColumnWidth: () => {}, autoResizeColumns: () => {},
  });
  return sh;
}

let sheets = {}, toasts = [], props = {}, triggers = [], tabOrder = [];
const mirror = {};
const ss = {
  getSheetByName: n => sheets[n] || null,
  insertSheet: n => (sheets[n] = makeSheet(n, [])),
  toast: m => toasts.push(m),
  getSpreadsheetTimeZone: () => 'Asia/Seoul', setSpreadsheetTimeZone: () => {},
  setActiveSheet: s => { ss._active = s; return s; },
  moveActiveSheet: pos => { tabOrder.push([pos, ss._active._name]); },
};
const chainBuilder = () => { const b = new Proxy({}, { get: (t, p) => (p === 'build' ? () => ({}) : () => b) }); return b; };
const SpreadsheetApp = {
  getActive: () => ss,
  getUi: () => ({ createMenu: () => chainBuilder() }),
  newTextStyle: chainBuilder,
  newRichTextValue: () => {
    const o = { _text: '', _styles: [] };
    const b = { setText: t => { o._text = t; return b; }, setTextStyle: (s, e) => { o._styles.push([s, e]); return b; }, build: () => o };
    return b;
  },
  newConditionalFormatRule: chainBuilder,
  BorderStyle: { SOLID_MEDIUM: 'SOLID_MEDIUM' },
  InterpolationType: { MIN: 'MIN', MAX: 'MAX', NUMBER: 'NUMBER' },
};
function parseCsv(text) {           // RFC4180 간이 파서 (따옴표 지원)
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(f); f = ''; }
    else if (ch === '\n') { row.push(f.replace(/\r$/, '')); rows.push(row); row = []; f = ''; }
    else f += ch;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}
const Utilities = { parseCsv, formatDate: d => d.toISOString(), sleep: () => {} };
const UrlFetchApp = {
  fetch: url => {
    const file = url.split('/data/')[1];
    const body = mirror[file];
    return { getResponseCode: () => (body === undefined ? 404 : 200), getContentText: () => body };
  },
};
const ScriptApp = {
  getProjectTriggers: () => triggers.slice(),
  deleteTrigger: t => { triggers = triggers.filter(x => x !== t); },
  newTrigger: fn => {
    const t = { fn, getHandlerFunction: () => fn };
    const b = new Proxy({}, { get: (_, p) => (p === 'create' ? () => { triggers.push(t); return t; } : () => b) });
    return b;
  },
};
const PropertiesService = {
  getDocumentProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }),
};

for (const f of ['lotto_history.csv', 'pension_history.csv', 'lotto_picks.csv', 'pension_picks.csv',
                 'popularity.csv', 'backtest.csv', 'scoreboard.csv']) {
  mirror[f] = fs.readFileSync(`${LAB}/data/${f}`, 'utf8');
}
const code = fs.readFileSync(`${LAB}/lotto_sheets.gs`, 'utf8');
const ctx = { SpreadsheetApp, Utilities, UrlFetchApp, ScriptApp, PropertiesService, console };
vm.createContext(ctx);
vm.runInContext(code, ctx);

let fail = 0;
const check = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}: ${JSON.stringify(got)}${ok ? '' : `  (기대: ${JSON.stringify(want)})`}`);
};
const D = (y, m, d) => new Date(y, m - 1, d);

// ── 구버전(v1) 시트 상태 재현 ─────────────────────────────
sheets['로또이력'] = makeSheet('로또이력', [
  ['회차', '추첨일', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', '보너스', '1등당첨자', '1등금액', '판매액'],
  [1242, D(2026, 9, 19), 2, 4, 10, 16, 31, 41, 9, 9, 3281029250, 61286928832],   // 판매액 절반 값(구 오류)
]);
sheets['로또추천'] = makeSheet('로또추천', [
  ctx.LV1.head,
  [D(2026, 9, 20), 1243, D(2026, 9, 26), 'A', 9, 18, 30, 31, 32, 33, 153, 4, 0.5, '', ''],   // 9,18 적중
  [D(2026, 9, 20), 1243, D(2026, 9, 26), 'B', 1, 2, 3, 35, 36, 37, 114, 4, 0.5, '', ''],     // 보너스 35
]);
sheets['연금추천'] = makeSheet('연금추천', [
  ctx.PL.head,
  [D(2026, 9, 18), 334, D(2026, 9, 24), '1순위', 2, '956029', '', ''],   // 조+6자리 = 1등
  [D(2026, 9, 18), 334, D(2026, 9, 24), '2순위', 1, '000129', '', ''],   // 끝2자리(29) = 6등
]);
triggers = [{ fn: 'weeklyJob', getHandlerFunction: () => 'weeklyJob' }];

console.log('── 1회차 동기화 (자동 이전 포함) ──');
ctx.syncAll();
const S = n => sheets[n] && sheets[n]._data;
check('구 로또추천 → 로또추천_v1 보존', S('로또추천_v1').length - 1, 2);
check('구 연금추천 → 연금추천_v1 보존', S('연금추천_v1').length - 1, 2);
check('스키마 표시', props.schema, 'v2');
check('구 트리거 제거 + 3시간 트리거', triggers.map(t => t.fn), ['syncAll']);

const hist = S('로또이력');
const lastCsv = mirror['lotto_history.csv'].trim().split('\n').slice(-1)[0].split(',');
check('로또이력 전량 재수집 (행 수)', hist.length - 1, Number(lastCsv[0]));
check('로또이력 16열 헤더', hist[0].length, 16);
check('마지막 회차 판매액 = 총판매액', hist[hist.length - 1][11], Number(lastCsv[11]));
check('추첨일은 Date', Object.prototype.toString.call(hist[hist.length - 1][1]), '[object Date]');

const picks = mirror['lotto_picks.csv'].trim().split('\n').slice(1).map(l => l.split(','))
  .filter(r => r[3] === 'v2' && r[4] === '공식');
const P = S('로또추천');
check('v2 공식 추천만 가져옴 (대조군·v1 제외)', P.length - 1, picks.length);
check('v2 헤더', P[0], ctx.LV2.head);
check('첫 세트 번호 = CSV와 동일', P[1].slice(4, 10), picks[0].slice(6, 12).map(Number));
check('분할배율 숫자', typeof P[1][10], 'number');
check('미추첨 표시', P[1][12], '추첨 전');

const pp = mirror['pension_picks.csv'].trim().split('\n').slice(1).map(l => l.split(',')).filter(r => r[3] === 'v2');
const PP = S('연금추천');
check('연금 v2 후보 수', PP.length - 1, pp.length);
check('연금 번호 6자리 문자열', PP.slice(1).every(r => /^\d{6}$/.test(String(r[5]))), true);
check('연금 순위 표기', PP[1][3], '1순위');

// v1 기록 채점 (1243회: 9 18 24 38 43 44 + 35)
const V1 = S('로또추천_v1');
check('v1 A세트 적중', V1[1][13], '2개');
check('v1 B세트 보너스', V1[2][13], '0개+보너스');
check('v1 A세트 9·18 초록', [sheets['로또추천_v1']._fmt.bg['2,5'], sheets['로또추천_v1']._fmt.bg['2,6']], ['#c6efce', '#c6efce']);
check('v1 추첨일 실제값으로 교정', V1[1][2].getDate(), 26);
const PV1 = S('연금추천_v1');
check('연금 v1 1등 판정', PV1[1][7], '1등');
check('연금 v1 6등 판정', PV1[2][7], '6등');
check('연금 끝2자리 강조 오프셋', sheets['연금추천_v1']._fmt.rich['3,6']._styles.slice(-1)[0], [4, 6]);

for (const n of ['성적', '편중감시', '번호인기도', '검증', '통계']) check(`${n} 시트 생성`, (S(n) || []).length > 3, true);
const bias = S('편중감시');
const v2row = bias.find(r => r[0] === 'v2 현행');
check('편중감시 v2 구간 비중 최대편차 ≤10%', parseInt(v2row[6]) <= 10, true);
const pop = S('번호인기도');
check('번호인기도 7번 칸 값', typeof pop[2][8], 'number');
const chk = S('검증');
check('검증: 백테스트 3개 엔진 행', chk.filter(r => /^(v1|v2|무작위)/.test(String(r[0]))).length, 3);
check('탭 순서 첫 번째 = 로또추천', tabOrder[0], [1, '로또추천']);

console.log('\n── 2회차 동기화 (변경 없음 → 중복 없음) ──');
toasts = [];
const before = [S('로또이력').length, S('로또추천').length, S('연금추천').length];
ctx.syncAll();
check('행 수 불변', [S('로또이력').length, S('로또추천').length, S('연금추천').length], before);
check('변경 없음 안내', /변경 없음/.test(toasts.slice(-1)[0]), true);

console.log('\n── 새 회차 도착 시뮬레이션 ──');
const target = Number(picks[0][1]);
const ours = picks.filter(r => Number(r[1]) === target);
const A = ours[0].slice(6, 12).map(Number);
const drawn = [A[0], A[1], A[2], A[3], ...[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45]
  .filter(x => !ours.some(r => r.slice(6, 12).map(Number).includes(x))).slice(0, 2)].sort((a, b) => a - b);
const bonus = A[4];
mirror['lotto_history.csv'] += `${target},20261003,${drawn.join(',')},${bonus},10,2500000000,120000000000,90,3000,150000,2500000\n`;
mirror['lotto_picks.csv'] += `2026-10-03 21:31,${target + 1},2026-10-10,v2,공식,A,1,12,23,34,40,45,0.99\n`;
ctx.syncAll();
const P2 = S('로또추천');
const rowA = P2.findIndex(r => r[1] === target && r[3] === 'A');
check('A세트 4개+보너스 → 적중', P2[rowA][11], '4개+보너스');
check('A세트 등수 4등', P2[rowA][12], '4등');
check('다음 회차 추천 추가', P2.filter(r => r[1] === target + 1).length, 1);
check('다음 회차는 추첨 전', P2.find(r => r[1] === target + 1)[12], '추첨 전');
const fmtA = sheets['로또추천']._fmt.bg;
check('A세트 5번째 칸 보너스 호박색', fmtA[`${rowA + 1},9`], '#ffe08a');
check('성적 v2 채점 세트 = 5', S('성적').find(r => r[0] === 'v2 현행')[1], 5);

console.log('\n── 구 트리거 호환 ──');
try { ctx.weeklyJob(); console.log('PASS  weeklyJob 별칭 동작'); } catch (e) { fail++; console.log('FAIL  weeklyJob', e.message); }

console.log('\n' + (fail ? `❌ 실패 ${fail}건` : '✅ 전체 통과'));
process.exit(fail ? 1 : 0);

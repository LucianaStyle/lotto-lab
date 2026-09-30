# -*- coding: utf-8 -*-
"""엔진 불변 조건 검사 — 규칙을 바꾼 뒤 반드시 실행: python tools/check_engine.py

v1에서 23~31번 기피(0.28배)·연금 끝자리 4 쏠림(64%)이 몇 달간 아무도 모르게 이어졌다.
같은 종류의 편중이 다시 생기면 여기서 바로 실패하게 한다.
"""
import os
import shutil
import sys
import tempfile
from collections import Counter

import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import lotto_lab as L  # noqa: E402

fails = 0


def check(label, ok, detail=""):
    global fails
    fails += not ok
    print(f"{'PASS' if ok else 'FAIL'}  {label}{'' if ok else '  ← ' + str(detail)}")


lotto = pd.read_csv(L.LOTTO_CSV)
pension = L.load_pension()
beta = L.fit_popularity(lotto)
hist = frozenset(tuple(sorted(r)) for r in lotto[L.NUM_COLS].to_numpy(dtype=int))
last = tuple(int(v) for v in lotto.iloc[-1][L.NUM_COLS])

# ── 로또: 5세트 구조 ──
rng = np.random.default_rng(1)
bad_distinct = bad_quota = bad_pattern = bad_hist = 0
allnums = []
for _ in range(300):
    sets = L.generate_lotto(beta, 5, L.STRENGTH_DEFAULT, rng, hist, last)
    flat = [n for s in sets for n in s]
    allnums += flat
    bad_distinct += len(set(flat)) != 30
    q = [sum(lo <= n <= hi for n in flat) for lo, hi in L.RANGES]
    expect = [(hi - lo + 1) * 30 / 45 for lo, hi in L.RANGES]
    bad_quota += any(abs(a - b) >= 1 for a, b in zip(q, expect))
    bad_pattern += any(L.is_pattern(s) for s in sets)
    bad_hist += any(s in hist for s in sets)
check("로또 5세트 = 서로 다른 번호 30개", bad_distinct == 0, bad_distinct)
check("로또 구간별 개수가 비례 할당과 1개 미만 차이", bad_quota == 0, bad_quota)
check("로또 패턴(등차·4연속) 없음", bad_pattern == 0, bad_pattern)
check("로또 역대 1등 조합 재사용 없음", bad_hist == 0, bad_hist)

cnt = Counter(allnums)
rel = np.array([cnt[n] for n in range(1, 46)]) / (len(allnums) / 45)
check(f"로또 번호별 빈도 0.5~1.6배 (실제 {rel.min():.2f}~{rel.max():.2f})", rel.min() >= 0.5 and rel.max() <= 1.6)
share = [sum(cnt[n] for n in range(lo, hi + 1)) / len(allnums) / ((hi - lo + 1) / 45) for lo, hi in L.RANGES]
check(f"로또 구간 비중 0.95~1.05 (실제 {min(share):.2f}~{max(share):.2f})", min(share) >= 0.95 and max(share) <= 1.05)
n10 = [len(L.generate_lotto(beta, 10, 0.5, rng, hist, last)) for _ in range(3)]
check("로또 10세트 요청도 처리", n10 == [10, 10, 10], n10)

# ── 연금: 끝자리 분산 ──
bad_tail = bad_last2 = bad_jo = bad_fmt = 0
tails = Counter()
for s in range(200):
    c = L.generate_pension(20, np.random.default_rng(s))
    nums = [x["num"] for x in c]
    bad_fmt += not all(len(n) == 6 and n.isdigit() for n in nums)
    bad_tail += any(len({n[5] for n in nums[i:i + 10]}) != 10 for i in (0, 10))
    bad_last2 += len({n[4:] for n in nums}) != 20
    bad_jo += any(len({x["jo"] for x in c[i:i + 5]}) != 5 for i in range(0, 20, 5))
    tails.update(n[5] for n in nums)
check("연금 번호 6자리", bad_fmt == 0, bad_fmt)
check("연금 10순위 단위로 끝자리 0~9 전부", bad_tail == 0, bad_tail)
check("연금 끝 2자리 20개 모두 다름", bad_last2 == 0, bad_last2)
check("연금 5순위 단위로 조 1~5 전부", bad_jo == 0, bad_jo)
check(f"연금 끝자리 최다 비중 ≤ 11% (실제 {max(tails.values()) / sum(tails.values()) * 100:.1f}%)",
      max(tails.values()) / sum(tails.values()) <= 0.11)

# ── 기록부: 회차당 1회만 생성 (임시 폴더에서) ──
tmp = tempfile.mkdtemp()
try:
    saved = (L.LOTTO_PICKS_CSV, L.PENSION_PICKS_CSV)
    L.LOTTO_PICKS_CSV, L.PENSION_PICKS_CSV = os.path.join(tmp, "l.csv"), os.path.join(tmp, "p.csv")
    lp1, pp1, ch1 = L.ensure_picks(lotto, pension, 5, 20, 0.5)
    lp2, pp2, ch2 = L.ensure_picks(lotto, pension, 5, 20, 0.5)
    check("첫 실행은 생성", ch1 and len(lp1) == 10 and len(pp1) == 20, (ch1, len(lp1), len(pp1)))
    same = ((lp1[L.NUM_COLS].astype(int).to_numpy() == lp2[L.NUM_COLS].astype(int).to_numpy()).all()
            and list(pp1["num"].astype(str)) == list(pp2["num"].astype(str)))
    check("재실행은 그대로 (같은 번호 유지)", (not ch2) and same, (ch2, same))
    tgt = int(lotto["epsd"].max()) + 1
    check("추첨일 = 직전 추첨일 + 7일", lp1["draw_date"].iloc[0] == L.draw_date_for(lotto, tgt).isoformat())
finally:
    L.LOTTO_PICKS_CSV, L.PENSION_PICKS_CSV = saved
    shutil.rmtree(tmp)

# ── 데이터: 판매액 필드 ──
recent = lotto.tail(100)
ratio5 = (recent["r5_n"] / (L.lotto_games(recent) * L.P_RANK5)).median()
check(f"5등 기대 대비 배율 ≈ 1 (실제 {ratio5:.3f}) — 판매액 필드가 총판매액인지", 0.9 < ratio5 < 1.1)

print("\n" + ("✅ 전체 통과" if not fails else f"❌ 실패 {fails}건"))
sys.exit(1 if fails else 0)

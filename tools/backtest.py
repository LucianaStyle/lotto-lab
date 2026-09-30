# -*- coding: utf-8 -*-
"""엔진 검증 — 워크포워드 백테스트 + 몬테카를로. 결과: data/backtest.csv

    python tools/backtest.py            # 최근 300회 (약 3분)
    python tools/backtest.py --draws 100

각 회차마다 '그 회차 이전 데이터만'으로 추천을 만들고 실제 당첨번호로 채점한다.
주당 5등 이상 확률은 300주 실측의 표준오차가 ±1.8%p라 엔진 간 차이를 가리지 못하므로,
균등 무작위 추첨 4만 회에 대한 몬테카를로로 따로 잰다(추첨이 균등하다는 것은 카이제곱으로 확인됨).

비교 대상 v1(2026-07~09 운영)은 재현을 위해 이 파일에만 원본 그대로 보존한다.
"""
import argparse
import itertools
import os
import random
import sys
import time

import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import lotto_lab as L  # noqa: E402


# ── v1 엔진 원본 (비교용 보존) ─────────────────────────────────────────
def _v1_features(nums):
    s = sorted(nums)
    return {"sum": sum(s), "odd": sum(x % 2 for x in s), "low": sum(x <= 22 for x in s),
            "le31": sum(x <= 31 for x in s), "consec": sum(b - a == 1 for a, b in zip(s, s[1:])),
            "decades": len({(x - 1) // 10 for x in s})}


def _v1_model(df):
    feats = pd.DataFrame([_v1_features(tuple(r)) for r in df[L.NUM_COLS].to_numpy()])
    latest = df["epsd"].max()
    m = df[(df["sales"] > 0) & (df["epsd"] >= latest - 520)]
    mf = pd.DataFrame([_v1_features(tuple(r)) for r in m[L.NUM_COLS].to_numpy()])
    expected = L.lotto_games(m).to_numpy() / L.TOTAL_COMBOS
    ratio = (m["rank1_winners"].to_numpy() + 0.5) / (expected + 0.5)
    X = np.column_stack([np.ones(len(mf)), mf["le31"], mf["consec"], mf["odd"],
                         np.abs(mf["sum"] - feats["sum"].mean()) / 10.0, mf["decades"]])
    coef, *_ = np.linalg.lstsq(X, np.log(ratio), rcond=None)
    return {"coef": coef, "sum_mean": feats["sum"].mean(),
            "sum_range": tuple(np.percentile(feats["sum"], [5, 95])),
            "hist": {tuple(sorted(r)) for r in df[L.NUM_COLS].to_numpy()},
            "last": tuple(sorted(df.iloc[-1][L.NUM_COLS]))}


def _v1_sets(a, n_sets=5, pool_size=15_000, seed=None):
    rng = random.Random(seed)
    lo, hi = a["sum_range"]

    def pop(c):
        f = _v1_features(c)
        return float(np.array([1.0, f["le31"], f["consec"], f["odd"],
                               abs(f["sum"] - a["sum_mean"]) / 10.0, f["decades"]]) @ a["coef"])
    cands, seen = [], set()
    while len(cands) < pool_size and len(seen) < pool_size * 20:
        c = tuple(sorted(rng.sample(range(1, 46), 6)))
        if c in seen:
            continue
        seen.add(c)
        f = _v1_features(c)
        if (lo <= f["sum"] <= hi and 2 <= f["odd"] <= 4 and 2 <= f["low"] <= 4 and f["consec"] <= 1
                and f["decades"] >= 3 and f["le31"] < 6 and c not in a["hist"]
                and len(set(c) & set(a["last"])) < 4):
            cands.append(c)
    top = sorted(cands, key=pop)[:max(n_sets * 40, len(cands) // 10)]
    rng.shuffle(top)
    picked = []
    for c in top:
        if all(len(set(c) & set(p)) <= 2 for p in picked):
            picked.append(c)
        if len(picked) == n_sets:
            break
    return picked


# ── 평가 ───────────────────────────────────────────────────────────
def mc_any_prize(portfolios, n_draws=40_000, seed=7):
    """포트폴리오(5세트)별로 '균등 무작위 추첨에서 3개 이상 일치 세트가 1개라도' 확률의 평균."""
    rng = np.random.default_rng(seed)
    D = np.zeros((n_draws, 45), dtype=np.int8)
    D[np.arange(n_draws)[:, None], np.argsort(rng.random((n_draws, 45)), axis=1)[:, :6]] = 1
    ps = []
    for sets in portfolios:
        S = np.zeros((len(sets), 45), dtype=np.int8)
        for j, s in enumerate(sets):
            S[j, np.array(s) - 1] = 1
        ps.append(((D @ S.T) >= 3).any(1).mean())
    return float(np.mean(ps)) * 100


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--draws", type=int, default=300)
    args = ap.parse_args()
    df = pd.read_csv(L.LOTTO_CSV)
    start = len(df) - args.draws
    nums = df[L.NUM_COLS].to_numpy(dtype=int)
    rng = np.random.default_rng(42)
    engines = ["v1 (구)", "v2 (신규)", "무작위"]
    res = {e: {"hits": [], "nums": [], "mult": [], "ports": []} for e in engines}
    t0 = time.time()
    for idx in range(start, len(df)):
        past = df.iloc[:idx]
        beta = L.fit_popularity(past)
        hist = frozenset(tuple(sorted(r)) for r in nums[:idx])
        win = set(nums[idx])
        sets = {
            "v1 (구)": _v1_sets(_v1_model(past), 5, seed=int(df.epsd[idx])),
            "v2 (신규)": L.generate_lotto(beta, 5, L.STRENGTH_DEFAULT, rng, hist, tuple(nums[idx - 1])),
            "무작위": L.generate_random(rng, 5),
        }
        for e, ss in sets.items():
            res[e]["hits"] += [len(win & set(s)) for s in ss]
            res[e]["nums"] += [n for s in ss for n in s]
            res[e]["mult"] += [L.split_mult(s, beta) for s in ss]
            res[e]["ports"].append(ss)
        if (idx - start) % 50 == 0:
            print(f"  {int(df.epsd[idx])}회 ... {time.time() - t0:.0f}s", flush=True)

    rand_mult = np.mean(res["무작위"]["mult"])
    rows = []
    for e in engines:
        h, ns = np.array(res[e]["hits"]), np.array(res[e]["nums"])
        rg = [((ns >= lo) & (ns <= hi)).mean() / ((hi - lo + 1) / 45) for lo, hi in L.RANGES]
        rows.append({
            "engine": e, "draws": args.draws, "sets": len(h),
            "avg_hits": round(h.mean(), 4),
            "pct_prize": round((h >= 3).mean() * 100, 3),
            "weekly_any_prize_mc": round(mc_any_prize(res[e]["ports"][:150]), 3),
            "split_mult": round(np.mean(res[e]["mult"]) / rand_mult, 4),
            "range_dev": round(max(abs(x - 1) for x in rg) * 100, 2),
            "share_20s": round(((ns >= 20) & (ns <= 29)).mean() / (10 / 45), 3),
        })
    out = pd.DataFrame(rows)
    out.to_csv(L.BACKTEST_CSV, index=False)
    print(out.to_string(index=False))
    print(f"\n→ {L.BACKTEST_CSV}  (이론: 평균 적중 0.800, 5등↑ 세트 2.38%, 무작위 5세트 주당 11.3%)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()

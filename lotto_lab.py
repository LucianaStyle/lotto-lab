# -*- coding: utf-8 -*-
"""
lotto_lab.py — 로또 6/45 + 연금복권 720+ 데이터 수집·통계 분석·번호 생성 도구

사용법:
    python lotto_lab.py            # 데이터 갱신 + 전체 분석 + 추천 번호 출력
    python lotto_lab.py --update   # 데이터 갱신만
    python lotto_lab.py --sets 10  # 추천 세트 수 지정

데이터 출처: 동행복권 공개 API (2026-07 개편 신규 엔드포인트)
  로또:   /lt645/selectPstLt645InfoNew.do?srchDir=center&srchLtEpsd=N  (회차당 ~10건)
  연금:   /pt720/selectPstPt720WnList.do                               (전체 이력 일괄)

방법론 요약 (정직 고지):
  * 각 추첨은 독립시행 — 어떤 통계도 '다음 번호'의 적중 확률을 높이지 못한다.
  * 이 도구가 실제로 최적화하는 것은 두 가지:
    (1) 프로파일 필터: 과거 당첨 조합의 통계적 형태(합계·홀짝·구간분산)에서
        극단적으로 벗어나는 조합을 배제 — 확률은 동일하나 "전형적" 조합만 남김.
    (2) 분할 회피(실질 기대값 개선): 회차별 1등 당첨자 수와 판매량 데이터로
        '조합 인기도 모형'을 적합, 남들이 많이 찍는 조합을 피해
        당첨 시 독식 확률을 높인다. 이것이 유일하게 수학적으로 유효한 엣지다.
"""
import argparse
import itertools
import json
import math
import os
import random
import sys
import time
from collections import Counter
from datetime import date, datetime

import numpy as np
import pandas as pd
import requests
from scipy import stats

BASE = "https://www.dhlottery.co.kr"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36"}
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
LOTTO_CSV = os.path.join(DATA_DIR, "lotto_history.csv")
PENSION_CSV = os.path.join(DATA_DIR, "pension_history.csv")
REPORT_MD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report.md")

FIRST_DRAW_DATE = date(2002, 12, 7)  # 로또 1회 추첨일
TOTAL_COMBOS = 8_145_060             # C(45,6)
TICKET_PRICE = 1000

NUM_COLS = ["n1", "n2", "n3", "n4", "n5", "n6"]


# ──────────────────────────────── 데이터 수집 ────────────────────────────────

def estimate_latest_epsd() -> int:
    return 1 + (date.today() - FIRST_DRAW_DATE).days // 7


def get_with_retry(sess, url: str, params: dict | None = None, tries: int = 3):
    """일시 장애(네트워크·5xx) 대비 3회 재시도. 마지막 실패는 그대로 올린다."""
    for i in range(tries):
        try:
            r = sess.get(url, params=params, headers=UA, timeout=15)
            r.raise_for_status()
            return r
        except (requests.RequestException, ValueError):
            if i == tries - 1:
                raise
            time.sleep(2 * (i + 1))


def fetch_lotto_window(sess: requests.Session, epsd: int) -> list[dict]:
    url = f"{BASE}/lt645/selectPstLt645InfoNew.do"
    r = get_with_retry(sess, url, {"srchDir": "center", "srchLtEpsd": epsd})
    rows = r.json().get("data", {}).get("list", []) or []
    out = []
    for d in rows:
        out.append({
            "epsd": d["ltEpsd"],
            "date": d["ltRflYmd"],
            "n1": d["tm1WnNo"], "n2": d["tm2WnNo"], "n3": d["tm3WnNo"],
            "n4": d["tm4WnNo"], "n5": d["tm5WnNo"], "n6": d["tm6WnNo"],
            "bonus": d["bnsWnNo"],
            "rank1_winners": d.get("rnk1WnNope"),
            "rank1_amount": d.get("rnk1WnAmt"),
            "sales": d.get("rlvtEpsdSumNtslAmt"),
        })
    return out


def update_lotto() -> pd.DataFrame:
    os.makedirs(DATA_DIR, exist_ok=True)
    have: dict[int, dict] = {}
    if os.path.exists(LOTTO_CSV):
        old = pd.read_csv(LOTTO_CSV)
        have = {int(r["epsd"]): r.to_dict() for _, r in old.iterrows()}

    sess = requests.Session()
    # 최신 회차 확정: 미래 회차는 빈 리스트가 오므로 추정치부터 아래로 탐색
    probe = []
    for e in range(estimate_latest_epsd() + 1, estimate_latest_epsd() - 6, -1):
        probe = fetch_lotto_window(sess, e)
        if probe:
            break
        time.sleep(0.1)
    if not probe:
        raise RuntimeError("로또 API 응답이 비어 있음 — 사이트 개편 여부 확인 필요")
    latest = max(r["epsd"] for r in probe)
    for r in probe:
        have[r["epsd"]] = r

    missing = [e for e in range(1, latest + 1) if e not in have]
    if missing:
        print(f"[로또] 최신 {latest}회 / 수집 필요 {len(missing)}회 다운로드 중...", flush=True)
        # center 윈도우가 ~10건씩 반환하므로 9회차 간격으로 순회
        targets = sorted({min(latest, e + 4) for e in missing})
        done = set()
        for t in targets:
            if t in done:
                continue
            for r in fetch_lotto_window(sess, t):
                have[r["epsd"]] = r
                done.add(r["epsd"])
            time.sleep(0.15)
        missing = [e for e in range(1, latest + 1) if e not in have]
        for e in missing:  # 윈도우 경계에서 빠진 회차 보충
            for r in fetch_lotto_window(sess, e):
                have[r["epsd"]] = r
            time.sleep(0.15)

    df = pd.DataFrame(sorted(have.values(), key=lambda r: r["epsd"]))
    df.to_csv(LOTTO_CSV, index=False)
    print(f"[로또] {len(df)}회분 저장 완료 → {LOTTO_CSV}")
    return df


def update_pension() -> pd.DataFrame:
    os.makedirs(DATA_DIR, exist_ok=True)
    r = get_with_retry(requests.Session(), f"{BASE}/pt720/selectPstPt720WnList.do")
    rows = r.json()["data"]["result"]
    df = pd.DataFrame([{
        "epsd": d["psltEpsd"], "date": d["psltRflYmd"],
        "jo": int(d["wnBndNo"]), "num": str(d["wnRnkVl"]).zfill(6),
        "bonus": str(d["bnsRnkVl"]).zfill(6),
    } for d in rows]).sort_values("epsd").reset_index(drop=True)
    df.to_csv(PENSION_CSV, index=False)
    print(f"[연금] {len(df)}회분 저장 완료 → {PENSION_CSV}")
    return df


# ──────────────────────────────── 조합 특성 ────────────────────────────────

def combo_features(nums: tuple[int, ...]) -> dict:
    s = sorted(nums)
    consec = sum(1 for a, b in zip(s, s[1:]) if b - a == 1)
    return {
        "sum": sum(s),
        "odd": sum(1 for x in s if x % 2),
        "low": sum(1 for x in s if x <= 22),          # 1~22 저구간
        "le31": sum(1 for x in s if x <= 31),         # 생일 범위(인기 구간)
        "consec": consec,
        "decades": len({(x - 1) // 10 for x in s}),   # 십의 자리 구간 다양성
        "range": s[-1] - s[0],
    }


# ──────────────────────────────── 통계 분석 ────────────────────────────────

def analyze_lotto(df: pd.DataFrame, recent_n: int = 52) -> dict:
    all_nums = df[NUM_COLS].to_numpy().ravel()
    freq = Counter(all_nums)
    n_draws = len(df)

    # 1) 균등성 검정 (카이제곱): p가 크면 "완전 무작위와 구별 불가"
    obs = np.array([freq.get(i, 0) for i in range(1, 46)])
    chi2, pval = stats.chisquare(obs)

    # 2) 최근 구간 hot/cold
    recent = df.tail(recent_n)
    rfreq = Counter(recent[NUM_COLS].to_numpy().ravel())

    # 3) 미출현 기간 (overdue)
    last_seen = {}
    for _, row in df.iterrows():
        for c in NUM_COLS:
            last_seen[row[c]] = row["epsd"]
    latest = df["epsd"].max()
    overdue = {i: latest - last_seen.get(i, 0) for i in range(1, 46)}

    # 4) 조합 프로파일 분포 (필터 경계 산출: 중앙 90%)
    feats = pd.DataFrame([combo_features(tuple(r)) for r in df[NUM_COLS].to_numpy()])
    sum_lo, sum_hi = np.percentile(feats["sum"], [5, 95])

    # 5) 짝 동반출현 top
    pair_cnt = Counter()
    for r in df[NUM_COLS].to_numpy():
        pair_cnt.update(itertools.combinations(sorted(r), 2))

    # 6) 인기도 모형: 1등 당첨자수(판매량 보정) ~ 조합 특성
    #    ratio = 관측 당첨자수 / (판매게임수 / 8,145,060) → 1보다 크면 대중적 조합
    m = df.dropna(subset=["rank1_winners", "sales"]).copy()
    m = m[(m["sales"] > 0) & (m["epsd"] >= latest - 520)]  # 최근 10년: 현재 구매 행태 반영
    mf = pd.DataFrame([combo_features(tuple(r)) for r in m[NUM_COLS].to_numpy()])
    expected = (m["sales"].to_numpy() / TICKET_PRICE) / TOTAL_COMBOS
    ratio = (m["rank1_winners"].to_numpy() + 0.5) / (expected + 0.5)
    X = np.column_stack([
        np.ones(len(mf)),
        mf["le31"], mf["consec"], mf["odd"],
        np.abs(mf["sum"] - feats["sum"].mean()) / 10.0,
        mf["decades"],
    ])
    coef, *_ = np.linalg.lstsq(X, np.log(ratio), rcond=None)
    # 회귀 유의성(대략): 잔차 대비 설명력
    pred = X @ coef
    ss_res = np.sum((np.log(ratio) - pred) ** 2)
    ss_tot = np.sum((np.log(ratio) - np.log(ratio).mean()) ** 2)
    r2 = 1 - ss_res / ss_tot

    return {
        "n_draws": n_draws, "latest": int(latest),
        "freq": freq, "rfreq": rfreq, "recent_n": recent_n,
        "chi2": chi2, "pval": pval,
        "overdue": overdue,
        "feats": feats, "sum_range": (sum_lo, sum_hi),
        "pairs_top": pair_cnt.most_common(10),
        "pop_coef": coef, "pop_r2": r2,
        "sum_mean": feats["sum"].mean(),
        "history_sets": {tuple(sorted(r)) for r in df[NUM_COLS].to_numpy()},
        "last_draw": tuple(sorted(df.iloc[-1][NUM_COLS])),
    }


def popularity_score(nums: tuple[int, ...], coef: np.ndarray, sum_mean: float) -> float:
    """예측 log(당첨자수 배율). 낮을수록 '남들이 안 찍는' 조합."""
    f = combo_features(nums)
    x = np.array([1.0, f["le31"], f["consec"], f["odd"],
                  abs(f["sum"] - sum_mean) / 10.0, f["decades"]])
    return float(x @ coef)


# ──────────────────────────────── 번호 생성 ────────────────────────────────

def generate_sets(a: dict, n_sets: int = 5, pool_size: int = 200_000,
                  seed: int | None = None) -> list[dict]:
    rng = random.Random(seed)
    sum_lo, sum_hi = a["sum_range"]
    cands = []
    seen = set()
    while len(cands) < pool_size:
        c = tuple(sorted(rng.sample(range(1, 46), 6)))
        if c in seen:
            continue
        seen.add(c)
        f = combo_features(c)
        # 프로파일 필터: 역대 1등 조합의 전형적 형태 (중앙 90% 구간)
        if not (sum_lo <= f["sum"] <= sum_hi):
            continue
        if not (2 <= f["odd"] <= 4):
            continue
        if not (2 <= f["low"] <= 4):
            continue
        if f["consec"] > 1:
            continue
        if f["decades"] < 3:
            continue
        if f["le31"] == 6:      # 전원 생일범위 → 분할 위험 최대, 배제
            continue
        if c in a["history_sets"]:  # 역대 1등 조합 재출현 배제(대중이 재구매하는 조합)
            continue
        if len(set(c) & set(a["last_draw"])) >= 4:
            continue
        cands.append(c)
        if len(seen) > pool_size * 20:
            break

    # 인기도(예측 분할 위험) 오름차순 정렬 → 상위 10% 후보군에서 무작위 선별
    # (엄격 최소화는 조합이 한 형태로 수렴하므로, 저인기 구간 내 다양성을 확보)
    scored = sorted(cands, key=lambda c: popularity_score(c, a["pop_coef"], a["sum_mean"]))
    top = scored[:max(n_sets * 40, len(scored) // 10)]
    rng.shuffle(top)
    picked: list[tuple] = []
    for c in top:
        if all(len(set(c) & set(p)) <= 2 for p in picked):
            picked.append(c)
        if len(picked) == n_sets:
            break
    return [{
        "nums": p,
        "pop": popularity_score(p, a["pop_coef"], a["sum_mean"]),
        **combo_features(p),
    } for p in picked]


def analyze_pension(df: pd.DataFrame) -> dict:
    digits = np.array([[int(ch) for ch in s] for s in df["num"]])
    pos_freq = [Counter(digits[:, i]) for i in range(6)]
    jo_freq = Counter(df["jo"])
    # 자리별 균등성 검정
    pvals = []
    for i in range(6):
        obs = np.array([pos_freq[i].get(d, 0) for d in range(10)])
        pvals.append(stats.chisquare(obs).pvalue)
    jo_obs = np.array([jo_freq.get(j, 0) for j in range(1, 6)])
    jo_p = stats.chisquare(jo_obs).pvalue
    return {"n": len(df), "latest": int(df["epsd"].max()),
            "pos_freq": pos_freq, "jo_freq": jo_freq, "pos_pvals": pvals, "jo_p": jo_p}


def generate_pension(a: dict, n: int = 6, seed: int | None = None) -> list[dict]:
    """순위가 매겨진 후보 n개 생성. 연금복권은 조·번호 조합이 1장씩만 판매되므로
    품절 대비 예비 후보가 필요하다. 1~5순위는 서로 다른 조로 분산(특정 조 매진 대비).
    순위 기준은 자리별 미달빈도 가중 점수 — 재미 요소이며 실제 확률은 모두 동일하다.
    """
    rng = random.Random(seed)
    weights = []
    for i in range(6):
        obs = np.array([a["pos_freq"][i].get(d, 0) for d in range(10)])
        weights.append((obs.mean() * 2 - obs).clip(min=1).astype(float))
    jo_obs = np.array([a["jo_freq"].get(j, 0) for j in range(1, 6)])
    jo_w = (jo_obs.mean() * 2 - jo_obs).clip(min=1).astype(float)

    pool, seen = [], set()
    for _ in range(600):
        num = tuple(rng.choices(range(10), weights=weights[i])[0] for i in range(6))
        jo = rng.choices(range(1, 6), weights=jo_w)[0]
        if (jo, num) in seen:
            continue
        seen.add((jo, num))
        score = sum(weights[i][num[i]] for i in range(6)) + jo_w[jo - 1]
        pool.append({"jo": jo, "num": num, "score": float(score)})
    pool.sort(key=lambda p: -p["score"])

    picked: list[dict] = []
    for p in pool:  # 1차: 조 미중복 + 자리 일치 ≤3 (후보 간 다양성)
        if len(picked) >= n:
            break
        if len(picked) < 5 and p["jo"] in [q["jo"] for q in picked]:
            continue
        if any(sum(1 for i in range(6) if p["num"][i] == q["num"][i]) > 3 for q in picked):
            continue
        picked.append(p)
    for p in pool:  # 2차: 모자라면 제약 완화해 채움
        if len(picked) >= n:
            break
        if p not in picked:
            picked.append(p)
    return [{"rank": i + 1, "jo": p["jo"], "num": "".join(map(str, p["num"]))}
            for i, p in enumerate(picked)]


# ──────────────────────────────── 리포트 ────────────────────────────────

def fmt_nums(nums) -> str:
    return " ".join(f"{n:2d}" for n in nums)


def run_report(n_sets: int, seed: int | None):
    lotto = pd.read_csv(LOTTO_CSV)
    pension = pd.read_csv(PENSION_CSV, dtype={"num": str, "bonus": str})
    a = analyze_lotto(lotto)
    p = analyze_pension(pension)

    lines = []
    w = lines.append
    w(f"# 로또 6/45 · 연금복권 720+ 통계 분석 리포트")
    w(f"생성: {datetime.now():%Y-%m-%d %H:%M} / 로또 {a['n_draws']}회(~{a['latest']}회) · 연금 {p['n']}회(~{p['latest']}회)\n")

    w("## 1. 무작위성 검증 (정직 고지)")
    w(f"- 로또 번호 균등성 카이제곱 검정: chi2={a['chi2']:.1f}, p={a['pval']:.3f}")
    w(f"  → p>0.05이면 역대 출현 빈도는 완전 무작위와 통계적으로 구별 불가 = '뜨거운 번호'는 노이즈.")
    w(f"- 연금복권 자리별 p값: {', '.join(f'{v:.2f}' for v in p['pos_pvals'])} / 조 p={p['jo_p']:.2f}\n")

    top10 = sorted(a["freq"].items(), key=lambda kv: -kv[1])[:10]
    bot10 = sorted(a["freq"].items(), key=lambda kv: kv[1])[:10]
    hot = sorted(a["rfreq"].items(), key=lambda kv: -kv[1])[:10]
    od = sorted(a["overdue"].items(), key=lambda kv: -kv[1])[:10]
    w("## 2. 빈도·미출현 (참고용 — 위 검정에 따라 예측력 없음)")
    w(f"- 역대 최다 출현: {', '.join(f'{n}({c})' for n, c in top10)}")
    w(f"- 역대 최소 출현: {', '.join(f'{n}({c})' for n, c in bot10)}")
    w(f"- 최근 {a['recent_n']}회 HOT: {', '.join(f'{n}({c})' for n, c in hot)}")
    w(f"- 최장 미출현: {', '.join(f'{n}번({c}회)' for n, c in od)}")
    w(f"- 동반출현 최다 짝: {', '.join(f'{a}-{b}({c})' for (a, b), c in a['pairs_top'][:6])}\n")

    f = a["feats"]
    w("## 3. 역대 1등 조합의 통계 프로파일 (생성 필터 근거)")
    w(f"- 합계: 평균 {f['sum'].mean():.0f}, 중앙 90% 구간 [{a['sum_range'][0]:.0f}, {a['sum_range'][1]:.0f}]")
    w(f"- 홀수 개수 분포: {dict(sorted(Counter(f['odd']).items()))}")
    w(f"- 연속수 쌍 분포: {dict(sorted(Counter(f['consec']).items()))}\n")

    w("## 4. 조합 인기도 모형 (분할 회피 = 유일한 실질 엣지)")
    w(f"- 회차별 1등 당첨자 수(판매량 보정)를 조합 특성으로 회귀. R²={a['pop_r2']:.3f}")
    c = a["pop_coef"]
    w(f"- 계수: 생일범위(≤31) 개수 {c[1]:+.3f}, 연속쌍 {c[2]:+.3f}, 홀수개수 {c[3]:+.3f}, "
      f"합계이탈 {c[4]:+.3f}, 구간다양성 {c[5]:+.3f}")
    w(f"  → 양(+)의 계수 특성이 많은 조합일수록 당첨자가 많았음(상금 분할↑). 생성기는 이 점수를 최소화.\n")

    sets = generate_sets(a, n_sets=n_sets, seed=seed)
    w(f"## 5. 이번 주 추천 조합 (로또 {a['latest'] + 1}회차용, {n_sets}세트)")
    w("프로파일 필터 통과 + 예측 인기도(분할 위험) 최소화 + 세트 간 중복 ≤2")
    for i, s in enumerate(sets, 1):
        w(f"- {chr(64+i)}세트: **{fmt_nums(s['nums'])}**  (합 {s['sum']}, 홀 {s['odd']}, 인기도 {s['pop']:+.2f})")
    w("")

    pns = generate_pension(p, n=max(6, n_sets), seed=seed)
    w(f"## 6. 연금복권 720+ 추천 ({p['latest'] + 1}회차용, 품절 대비 순위 리스트)")
    w("연금복권은 조·번호 조합이 1장씩만 판매되므로, 앞 순위가 품절이면 다음 순위로 구매.")
    w("1~5순위는 서로 다른 조로 분산(특정 조 매진 대비). 실제 당첨 확률은 모든 조합이 동일.")
    for pn in pns:
        w(f"- {pn['rank']}순위: **{pn['jo']}조 {pn['num']}**")
    w("")

    w("> ⚠️ 본 리포트의 어떤 항목도 당첨 확률 자체를 높이지 않는다. 5번 항목의 분할 회피만이")
    w("> '당첨됐을 때 더 많이 받는' 방향의 실질적 최적화다. 구매는 여유 자금 내에서.")

    text = "\n".join(lines)
    with open(REPORT_MD, "w", encoding="utf-8") as fp:
        fp.write(text)
    print(text)
    print(f"\n리포트 저장 → {REPORT_MD}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--update", action="store_true", help="데이터 갱신만 수행")
    ap.add_argument("--no-fetch", action="store_true", help="저장된 데이터로 분석만")
    ap.add_argument("--sets", type=int, default=5)
    ap.add_argument("--seed", type=int, default=None, help="재현 가능한 생성용 시드")
    args = ap.parse_args()

    if not args.no_fetch:
        update_lotto()
        update_pension()
    if not args.update:
        run_report(args.sets, args.seed)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()

# -*- coding: utf-8 -*-
"""
lotto_lab.py — 로또 6/45 + 연금복권 720+ 데이터 수집·분석·번호 생성 (엔진 v2)

사용법:
    python lotto_lab.py              # 데이터 갱신 + 추천(회차당 1회만 생성) + 채점 + 리포트
    python lotto_lab.py --update     # 데이터 갱신만
    python lotto_lab.py --no-fetch   # 저장 데이터로 추천·리포트만
    python lotto_lab.py --strength 1 # 분할 회피 강도 (0=무작위 ~ 2=강함, 기본 0.5)
    python tools/backtest.py         # 엔진 검증 (워크포워드 백테스트 → data/backtest.csv)

데이터 출처: 동행복권 공개 API (2026-07 개편 신규 엔드포인트)
  로또: /lt645/selectPstLt645InfoNew.do?srchDir=center&srchLtEpsd=N  (회차당 ~10건)
  연금: /pt720/selectPstPt720WnList.do                               (전체 이력 일괄)

엔진 v2 원칙 (2026-09 개편 — v1은 23~31번을 0.28배로 기피하는 편중이 있었다):
  1. 확률: 모든 조합의 1등 확률은 1/8,145,060으로 같다. 어떤 필터도 이것을 바꾸지 못한다.
  2. 분할 회피: 5등(회당 ~270만 명) 당첨자 수의 '기대 대비 배율'로 번호별 인기도를 역산한다
     (표본 외 예측 상관 0.84). 인기 번호를 조금 덜 고르면 1등 당첨 시 나눠 가질 사람이 준다.
  3. 분산: 5세트 30개 번호를 겹치지 않게, 번호 구간(1-9/10대/20대/30대/40대)별 비율대로 배치한다.
     구간 편중이 구조적으로 불가능하고, 5세트가 같이 맞고 같이 틀리는 일이 줄어
     '주당 5등 이상 1건이라도' 확률이 무작위보다 높다(11.8% vs 11.3%, 기대 상금은 동일).
"""
import argparse
import os
import random
import re
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta
from math import comb

import numpy as np
import pandas as pd
import requests
from scipy import stats

BASE = "https://www.dhlottery.co.kr"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126 Safari/537.36"}
ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data")
LOTTO_CSV = os.path.join(DATA_DIR, "lotto_history.csv")
PENSION_CSV = os.path.join(DATA_DIR, "pension_history.csv")
LOTTO_PICKS_CSV = os.path.join(DATA_DIR, "lotto_picks.csv")
PENSION_PICKS_CSV = os.path.join(DATA_DIR, "pension_picks.csv")
POPULARITY_CSV = os.path.join(DATA_DIR, "popularity.csv")
SCOREBOARD_CSV = os.path.join(DATA_DIR, "scoreboard.csv")
BACKTEST_CSV = os.path.join(DATA_DIR, "backtest.csv")
REPORT_MD = os.path.join(ROOT, "report.md")
LOG_FILE = os.path.join(ROOT, "logs", "update.log")

FIRST_DRAW_DATE = date(2002, 12, 7)   # 로또 1회 추첨일
TOTAL_COMBOS = comb(45, 6)            # 8,145,060
NUM_COLS = ["n1", "n2", "n3", "n4", "n5", "n6"]
LOTTO_COLS = ["epsd", "date", *NUM_COLS, "bonus", "rank1_winners", "rank1_amount", "sales",
              "r2_n", "r3_n", "r4_n", "r5_n"]
RANGES = [(1, 9), (10, 19), (20, 29), (30, 39), (40, 45)]
ENGINE = "v2"
STRENGTH_DEFAULT = 0.5   # 1.0 이상이면 7·12번이 거의 안 나와 또 다른 편중으로 보인다 (튜닝 결과)
PENSION_N = 20           # 연금은 조·번호 조합이 1장뿐이라 품절 대비로 넉넉히

# k개 일치 확률 (균등 무작위 구매 가정) — 등수별 '기대 당첨자 수'의 기준
P_MATCH = {k: comb(6, k) * comb(39, 6 - k) / TOTAL_COMBOS for k in range(7)}
P_RANK5 = P_MATCH[3]


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
    r = get_with_retry(sess, f"{BASE}/lt645/selectPstLt645InfoNew.do",
                       {"srchDir": "center", "srchLtEpsd": epsd})
    out = []
    for d in r.json().get("data", {}).get("list", []) or []:
        out.append({
            "epsd": d["ltEpsd"], "date": d["ltRflYmd"],
            **{f"n{i}": d[f"tm{i}WnNo"] for i in range(1, 7)},
            "bonus": d["bnsWnNo"],
            "rank1_winners": d.get("rnk1WnNope"),
            "rank1_amount": d.get("rnk1WnAmt"),
            # 총판매액. rlvtEpsdSumNtslAmt는 당첨금 배분액(약 400회 이후 판매액의 50%)이라 쓰면 안 된다.
            "sales": d.get("wholEpsdSumNtslAmt"),
            **{f"r{k}_n": d.get(f"rnk{k}WnNope") for k in range(2, 6)},
        })
    return out


def update_lotto() -> pd.DataFrame:
    os.makedirs(DATA_DIR, exist_ok=True)
    have: dict[int, dict] = {}
    if os.path.exists(LOTTO_CSV):
        old = pd.read_csv(LOTTO_CSV)
        if "r5_n" in old.columns:            # 구 스키마(판매액 오류·등수 누락)는 통째로 재수집
            have = {int(r["epsd"]): r.to_dict() for _, r in old.iterrows()}
        else:
            print("[로또] 데이터 스키마 갱신 — 전 회차 재수집", flush=True)

    sess = requests.Session()
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
        done = set()
        for t in sorted({min(latest, e + 4) for e in missing}):   # center 윈도우가 ~10건씩
            if t in done:
                continue
            for r in fetch_lotto_window(sess, t):
                have[r["epsd"]] = r
                done.add(r["epsd"])
            time.sleep(0.15)
        for e in [e for e in range(1, latest + 1) if e not in have]:
            for r in fetch_lotto_window(sess, e):
                have[r["epsd"]] = r
            time.sleep(0.15)

    df = pd.DataFrame(sorted(have.values(), key=lambda r: r["epsd"]))[LOTTO_COLS]
    df.to_csv(LOTTO_CSV, index=False)
    print(f"[로또] {len(df)}회분 저장 완료 (최신 {latest}회)")
    return df


def update_pension() -> pd.DataFrame:
    os.makedirs(DATA_DIR, exist_ok=True)
    r = get_with_retry(requests.Session(), f"{BASE}/pt720/selectPstPt720WnList.do")
    df = pd.DataFrame([{
        "epsd": d["psltEpsd"], "date": d["psltRflYmd"],
        "jo": int(d["wnBndNo"]), "num": str(d["wnRnkVl"]).zfill(6),
        "bonus": str(d["bnsRnkVl"]).zfill(6),
    } for d in r.json()["data"]["result"]]).sort_values("epsd").reset_index(drop=True)
    df.to_csv(PENSION_CSV, index=False)
    print(f"[연금] {len(df)}회분 저장 완료 (최신 {int(df['epsd'].max())}회)")
    return df


def load_pension() -> pd.DataFrame:
    return pd.read_csv(PENSION_CSV, dtype={"num": str, "bonus": str})


# ──────────────────────────────── 날짜 ────────────────────────────────

WEEKDAY_KR = "월화수목금토일"


def draw_date_for(df: pd.DataFrame, target: int) -> date:
    """대상 회차의 추첨일 = 마지막 추첨일 + 7일 × 회차 차이.
    로또·연금 모두 창설 이래 예외 없이 주 1회다(전 회차 간격 7일 검증됨).
    '다음 토요일'식 계산은 추첨 직후 실행 시 당일을 반환해 7일 어긋난다."""
    last = df.iloc[-1]
    d = datetime.strptime(str(int(last["date"])), "%Y%m%d").date()
    return d + timedelta(days=7 * (int(target) - int(last["epsd"])))


def fmt_date(d: date) -> str:
    return f"{d:%Y-%m-%d}({WEEKDAY_KR[d.weekday()]})"


# ──────────────────────────────── 번호 인기도 모형 ────────────────────────────────
#
# 균등 무작위 구매라면 5등(3개 일치) 당첨자 기대값 = 판매게임수 × P(3개 일치)로 정확히 정해진다.
# 실제 5등이 기대보다 많았다면 그 회차 당첨번호가 사람들이 많이 고른 번호였다는 뜻이다.
# 회당 ~270만 명이라 잡음이 거의 없어, '어느 번호가 포함되면 배율이 오르는가'를 능형회귀로
# 역산하면 번호별 인기도가 나온다. 조합의 1등 동반당첨 배율 ≈ exp(2 × Σ번호 인기도)
# (5등은 6개 중 3개 부분집합의 인기를 반영하므로 6개 전체에는 약 2배로 작용 — 검증 구간에서
# 기울기 2가 포아송 로그우도 최대였다).

def lotto_games(df: pd.DataFrame) -> pd.Series:
    return df["sales"] / np.where(df["epsd"] <= 87, 2000, 1000)   # 2004-08 이전 1게임 2,000원


def number_matrix(df: pd.DataFrame) -> np.ndarray:
    X = np.zeros((len(df), 45))
    nums = df[NUM_COLS].to_numpy(dtype=int)
    X[np.arange(len(df))[:, None], nums - 1] = 1
    return X


def fit_popularity(df: pd.DataFrame, window: int = 520, lam: float = 20.0) -> np.ndarray:
    """최근 window회로 번호별 인기도(log 배율 기여, 평균 0으로 중심화)를 추정."""
    d = df.dropna(subset=["r5_n", "sales"]).tail(window)
    y = np.log(d["r5_n"].to_numpy() / (lotto_games(d).to_numpy() * P_RANK5))
    X = number_matrix(d)
    beta = np.linalg.solve(X.T @ X + lam * np.eye(45), X.T @ (y - y.mean()))
    return beta - beta.mean()


def split_mult(combo, beta: np.ndarray) -> float:
    """예상 1등 동반당첨자 배율 (무작위 조합 평균 ≈ 1.00, 낮을수록 당첨 시 몫이 큼)."""
    return float(np.exp(2 * beta[np.asarray(combo) - 1].sum()))


def validate_popularity(df: pd.DataFrame, test_n: int = 200) -> dict:
    """표본 외 검증: 검증 구간 직전까지로 학습 → 검증 구간 5등 배율 예측 상관,
    그리고 당첨번호 인기도 5분위별 실제 1등 당첨자 배율."""
    d = df.dropna(subset=["r5_n", "sales"]).reset_index(drop=True)
    train, test = d.iloc[:-test_n], d.iloc[-test_n:]
    beta = fit_popularity(train)
    y = np.log(test["r5_n"] / (lotto_games(test) * P_RANK5))
    pred = number_matrix(test) @ beta
    r = stats.pearsonr(pred, y)

    recent = d.tail(544)
    b_all = fit_popularity(d)
    score = number_matrix(recent) @ b_all
    exp1 = lotto_games(recent) * P_MATCH[6]
    q = pd.qcut(score, 5, labels=False)
    quint = [float(recent["rank1_winners"][q == k].sum() / exp1[q == k].sum()) for k in range(5)]
    return {"r": float(r.statistic), "p": float(r.pvalue), "n_test": test_n,
            "quintile_r1": quint, "n_quint": len(recent)}


# ──────────────────────────────── 로또 엔진 v2 ────────────────────────────────

def is_pattern(c) -> bool:
    """사람들이 일부러 고르는 모양(분할 위험 큼): 등차수열(용지의 가로·세로·대각선 포함), 4연속 이상."""
    s = sorted(c)
    d = [b - a for a, b in zip(s, s[1:])]
    if len(set(d)) == 1:
        return True
    run = best = 1
    for x in d:
        run = run + 1 if x == 1 else 1
        best = max(best, run)
    return best >= 4


def generate_lotto(beta: np.ndarray, n_sets: int = 5, strength: float = STRENGTH_DEFAULT,
                   rng: np.random.Generator | None = None, hist: frozenset = frozenset(),
                   last: tuple = ()) -> list[tuple]:
    """5세트(30개 번호) 단위 블록마다 번호를 겹치지 않게, 구간 비율대로 뽑는다.
    구간 안에서는 인기 번호일수록 덜 뽑히게(가중치 exp(-strength·z)) 해서 분할을 피한다.
    30개를 6개씩 나누는 여러 방법 중 패턴·역대 1등 조합을 피하면서 분할위험 합이 최소인 것을 쓴다."""
    rng = rng or np.random.default_rng()
    z = (beta - beta.mean()) / (beta.std() + 1e-12)
    w = np.exp(-strength * z)
    out: list[tuple] = []
    while len(out) < n_sets:
        k = min(5, n_sets - len(out))
        need = 6 * k
        raw = np.array([(hi - lo + 1) * need / 45 for lo, hi in RANGES])
        quota = np.floor(raw).astype(int)
        frac = raw - quota
        rem = need - quota.sum()
        if rem:
            quota[rng.choice(len(RANGES), rem, replace=False, p=frac / frac.sum())] += 1
        chosen = []
        for (lo, hi), cnt in zip(RANGES, quota):
            nums = np.arange(lo, hi + 1)
            p = w[nums - 1] / w[nums - 1].sum()
            chosen += [int(v) for v in rng.choice(nums, cnt, replace=False, p=p)]
        best, best_cost = None, np.inf
        for _ in range(400):
            perm = rng.permutation(chosen)
            sets = [tuple(sorted(int(v) for v in perm[i * 6:(i + 1) * 6])) for i in range(k)]
            if any(is_pattern(s) or s in hist or len(set(s) & set(last)) >= 4 for s in sets):
                continue
            if any(len({(v - 1) // 10 for v in s}) < 3 for s in sets):
                continue
            cost = sum(split_mult(s, beta) for s in sets)
            if cost < best_cost:
                best, best_cost = sets, cost
        if best:
            out += best
    return out[:n_sets]


def generate_random(rng: np.random.Generator, n_sets: int = 5) -> list[tuple]:
    """대조군: 완전 무작위. 엔진이 '맞히는 능력'이 무작위와 같다는 것을 매주 실전으로 보여준다."""
    return [tuple(sorted(int(v) for v in rng.choice(45, 6, replace=False) + 1)) for _ in range(n_sets)]


# ──────────────────────────────── 연금 엔진 v2 ────────────────────────────────

def generate_pension(n: int = PENSION_N, rng: np.random.Generator | None = None) -> list[dict]:
    """품절 대비 순위 후보 n개.

    연금 1등은 1장당 고정 연금이라 나눠 갖지 않는다 → 인기도 모형이 필요 없고, 번호는 균등 무작위.
    대신 등수가 '끝자리부터 몇 자리 일치'로 정해지므로 끝자리를 분산한다:
      - 10순위 단위로 끝자리 0~9를 한 번씩 → 앞에서 10장을 사면 매주 7등(끝 1자리) 이상 확정
      - 끝 2자리는 전 후보가 서로 다름 → 6등 이상 확률 = 산 장수 × 1%
      - 조는 5순위 단위로 1~5조 한 번씩 → 한 조가 통째로 매진돼도 대안이 남는다
    (v1은 자리별 '미달빈도' 점수 정렬로 끝자리 4가 후보의 64%를 차지했다.)"""
    rng = rng or np.random.default_rng()
    tails: list[int] = []
    while len(tails) < n:
        tails += [int(v) for v in rng.permutation(10)]
    jos: list[int] = []
    while len(jos) < n:
        jos += [int(v) for v in rng.permutation(5) + 1]
    used2, out = set(), []
    for i in range(n):
        while True:
            head = [int(v) for v in rng.integers(0, 10, 5)]
            last2 = (head[4], tails[i])
            if last2 not in used2 or len(used2) >= 100:
                break
        used2.add(last2)
        out.append({"rank": i + 1, "jo": jos[i], "num": "".join(map(str, head + [tails[i]]))})
    return out


# ──────────────────────────────── 추천 기록부 ────────────────────────────────
# 회차당 한 번만 생성해 기록한다 (v1은 실행할 때마다 새로 뽑아 '무엇을 추천했는지'가 남지 않았다).
# 시트는 이 파일을 GitHub 미러에서 가져가므로 리포트·시트가 항상 같은 번호를 보여준다.

LOTTO_PICK_COLS = ["created", "target", "draw_date", "engine", "strategy", "set", *NUM_COLS, "split_mult"]
PENSION_PICK_COLS = ["created", "target", "draw_date", "engine", "rank", "jo", "num"]


def _read(path: str, cols: list[str], **kw) -> pd.DataFrame:
    if os.path.exists(path):
        return pd.read_csv(path, **kw)
    return pd.DataFrame(columns=cols)


def ensure_picks(lotto: pd.DataFrame, pension: pd.DataFrame, n_sets: int, n_pension: int,
                 strength: float) -> tuple[pd.DataFrame, pd.DataFrame, bool]:
    lp = _read(LOTTO_PICKS_CSV, LOTTO_PICK_COLS)
    pp = _read(PENSION_PICKS_CSV, PENSION_PICK_COLS, dtype={"num": str})
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    changed = False

    target = int(lotto["epsd"].max()) + 1
    have = lp[(lp["target"] == target) & (lp["engine"] == ENGINE)]
    if have.empty:
        beta = fit_popularity(lotto)
        rng = np.random.default_rng(target * 7919)
        hist = frozenset(tuple(sorted(r)) for r in lotto[NUM_COLS].to_numpy(dtype=int))
        last = tuple(int(v) for v in lotto.iloc[-1][NUM_COLS])
        dd = draw_date_for(lotto, target).isoformat()
        rows = []
        for strategy, sets in (("공식", generate_lotto(beta, n_sets, strength, rng, hist, last)),
                               ("대조군", generate_random(rng, n_sets))):
            for i, s in enumerate(sets):
                rows.append({"created": now, "target": target, "draw_date": dd, "engine": ENGINE,
                             "strategy": strategy, "set": chr(65 + i), **dict(zip(NUM_COLS, s)),
                             "split_mult": round(split_mult(s, beta), 4)})
        lp = pd.concat([lp, pd.DataFrame(rows)], ignore_index=True)
        changed = True

    ptarget = int(pension["epsd"].max()) + 1
    if pp[(pp["target"] == ptarget) & (pp["engine"] == ENGINE)].empty:
        dd = draw_date_for(pension, ptarget).isoformat()
        rows = [{"created": now, "target": ptarget, "draw_date": dd, "engine": ENGINE, **c}
                for c in generate_pension(n_pension, np.random.default_rng(ptarget * 104729))]
        pp = pd.concat([pp, pd.DataFrame(rows)], ignore_index=True)
        changed = True

    if changed:
        lp[LOTTO_PICK_COLS].to_csv(LOTTO_PICKS_CSV, index=False)
        pp[PENSION_PICK_COLS].to_csv(PENSION_PICKS_CSV, index=False)
    return lp, pp, changed


def backfill_v1_from_log(lotto: pd.DataFrame, pension: pd.DataFrame) -> None:
    """v1 시절 추천을 로그에서 복원해 기록부에 넣는다 (성적 비교용, 1회성).
    회차별로 '추첨 전에 생성된 마지막 추천'만 인정한다 — 추첨 뒤 생성분은 결과를 알고 만든 것이 아니지만
    사용자가 볼 수 없었던 번호이므로 제외한다."""
    if not os.path.exists(LOG_FILE):
        return
    lp = _read(LOTTO_PICKS_CSV, LOTTO_PICK_COLS)
    pp = _read(PENSION_PICKS_CSV, PENSION_PICK_COLS, dtype={"num": str})
    if (lp["engine"] == "v1").any() or (pp["engine"] == "v1").any():
        return
    log = open(LOG_FILE, encoding="utf-8", errors="replace").read()
    lrows, prows = {}, {}
    for block in re.split(r"^생성: ", log, flags=re.M)[1:]:
        created = datetime.strptime(block[:16], "%Y-%m-%d %H:%M")
        m = re.search(r"로또 (\d+)회차", block)
        if m:
            t = int(m.group(1))
            dd = draw_date_for(lotto, t)
            sets = re.findall(r"^- [A-E]세트: \*\*([\d ]+)\*\*", block, re.M)
            if len(sets) >= 5 and created < datetime.combine(dd, datetime.min.time()) + timedelta(hours=20):
                lrows[t] = [{"created": created.strftime("%Y-%m-%d %H:%M"), "target": t,
                             "draw_date": dd.isoformat(), "engine": "v1", "strategy": "공식",
                             "set": chr(65 + i), **dict(zip(NUM_COLS, map(int, s.split()))),
                             "split_mult": np.nan} for i, s in enumerate(sets[:5])]
        m = re.search(r"연금복권 720\+ 추천 \((\d+)회차", block)
        if m:
            t = int(m.group(1))
            dd = draw_date_for(pension, t)
            cands = re.findall(r"\| (\d+) \| (\d)조 \| \*\*(\d{6})\*\*", block)
            cands += re.findall(r"^- (\d+)순위: \*\*(\d)조 (\d{6})\*\*", block, re.M)
            if cands and created < datetime.combine(dd, datetime.min.time()) + timedelta(hours=19):
                prows[t] = [{"created": created.strftime("%Y-%m-%d %H:%M"), "target": t,
                             "draw_date": dd.isoformat(), "engine": "v1", "rank": int(r),
                             "jo": int(j), "num": n} for r, j, n in sorted(cands, key=lambda x: int(x[0]))]
    if lrows:
        lp = pd.concat([pd.DataFrame([r for t in sorted(lrows) for r in lrows[t]]), lp], ignore_index=True)
        lp[LOTTO_PICK_COLS].to_csv(LOTTO_PICKS_CSV, index=False)
    if prows:
        pp = pd.concat([pd.DataFrame([r for t in sorted(prows) for r in prows[t]]), pp], ignore_index=True)
        pp[PENSION_PICK_COLS].to_csv(PENSION_PICKS_CSV, index=False)
    print(f"[기록부] v1 추천 복원: 로또 {len(lrows)}회차 · 연금 {len(prows)}회차")


# ──────────────────────────────── 채점 ────────────────────────────────

def lotto_rank(hits: int, bonus_hit: bool) -> str:
    return {6: "1등", 5: "2등" if bonus_hit else "3등", 4: "4등", 3: "5등"}.get(hits, "낙첨")


def pension_rank(num: str, jo: int, win: dict) -> tuple[int, str]:
    m = 0
    while m < 6 and num[5 - m] == win["num"][5 - m]:
        m += 1
    rank = {6: "1등" if jo == win["jo"] else "2등", 5: "3등", 4: "4등", 3: "5등", 2: "6등", 1: "7등"}.get(m, "낙첨")
    return m, rank


def grade(lp: pd.DataFrame, pp: pd.DataFrame, lotto: pd.DataFrame, pension: pd.DataFrame):
    win = {int(r["epsd"]): (set(int(r[c]) for c in NUM_COLS), int(r["bonus"])) for _, r in lotto.iterrows()}
    g = lp[lp["target"].isin(win)].copy()
    picked = [{int(v) for v in row} for row in g[NUM_COLS].to_numpy()]
    g["hits"] = [len(win[int(t)][0] & s) for t, s in zip(g["target"], picked)]
    g["bonus_hit"] = [win[int(t)][1] in s for t, s in zip(g["target"], picked)]
    g["rank"] = [lotto_rank(h, b) for h, b in zip(g["hits"], g["bonus_hit"])]

    pwin = {int(r["epsd"]): {"jo": int(r["jo"]), "num": str(r["num"]).zfill(6)} for _, r in pension.iterrows()}
    pg = pp[pp["target"].isin(pwin)].copy()
    res = [pension_rank(str(n).zfill(6), int(j), pwin[int(t)]) for n, j, t in zip(pg["num"], pg["jo"], pg["target"])]
    pg["tail"] = [m for m, _ in res]
    pg["rank_result"] = [r for _, r in res]
    return g, pg


def scoreboard(g: pd.DataFrame, pg: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (eng, strat), d in g.groupby(["engine", "strategy"]):
        n = len(d)
        rows.append({"game": "로또", "group": f"{eng} {strat}", "draws": d["target"].nunique(), "tickets": n,
                     "avg_hits": round(d["hits"].mean(), 3), "theory_avg": 0.8,
                     "prize_count": int((d["hits"] >= 3).sum()),
                     "theory_prize": round(n * sum(P_MATCH[k] for k in (3, 4, 5, 6)), 2),
                     "detail": " ".join(f"{k}개:{int((d['hits'] == k).sum())}" for k in range(7) if (d["hits"] == k).any())})
    for eng, d in pg.groupby("engine"):
        n = len(d)
        rows.append({"game": "연금", "group": f"{eng} 후보 전체", "draws": d["target"].nunique(), "tickets": n,
                     "avg_hits": round(d["tail"].mean(), 3), "theory_avg": round(sum(0.1 ** k for k in range(1, 7)), 3),
                     "prize_count": int((d["tail"] >= 1).sum()), "theory_prize": round(n * 0.1, 2),
                     "detail": " ".join(f"끝{k}자리:{int((d['tail'] == k).sum())}" for k in range(1, 7) if (d["tail"] == k).any())})
    return pd.DataFrame(rows)


# ──────────────────────────────── 참고 통계 ────────────────────────────────

def reference_stats(df: pd.DataFrame) -> dict:
    nums = df[NUM_COLS].to_numpy(dtype=int)
    freq = Counter(nums.ravel())
    obs = np.array([freq.get(i, 0) for i in range(1, 46)])
    chi2, pval = stats.chisquare(obs)
    latest = int(df["epsd"].max())
    last_seen = {}
    for e, row in zip(df["epsd"], nums):
        for n in row:
            last_seen[int(n)] = int(e)
    return {"chi2": chi2, "pval": pval, "freq": freq,
            "overdue": {i: latest - last_seen.get(i, 0) for i in range(1, 46)}}


def pension_uniformity(df: pd.DataFrame) -> list[float]:
    digits = np.array([[int(ch) for ch in s] for s in df["num"]])
    return [stats.chisquare(np.bincount(digits[:, i], minlength=10)).pvalue for i in range(6)]


# ──────────────────────────────── 리포트 ────────────────────────────────

def fmt_nums(nums) -> str:
    return " ".join(f"{int(n):2d}" for n in nums)


def run_report(n_sets: int, n_pension: int, strength: float) -> None:
    lotto = pd.read_csv(LOTTO_CSV)
    pension = load_pension()
    backfill_v1_from_log(lotto, pension)
    lp, pp, _ = ensure_picks(lotto, pension, n_sets, n_pension, strength)
    g, pg = grade(lp, pp, lotto, pension)
    sb = scoreboard(g, pg)
    sb.to_csv(SCOREBOARD_CSV, index=False)

    beta = fit_popularity(lotto)
    pop = pd.DataFrame({"number": range(1, 46), "popularity": np.round(np.exp(beta), 4)})
    pop["rank"] = pop["popularity"].rank(ascending=False, method="min").astype(int)
    pop.to_csv(POPULARITY_CSV, index=False)
    val = validate_popularity(lotto)
    ref = reference_stats(lotto)

    lt = int(lotto["epsd"].max()) + 1
    pt = int(pension["epsd"].max()) + 1
    cur = lp[(lp["target"] == lt) & (lp["engine"] == ENGINE)]
    official = cur[cur["strategy"] == "공식"]
    control = cur[cur["strategy"] == "대조군"]
    pcur = pp[(pp["target"] == pt) & (pp["engine"] == ENGINE)].sort_values("rank")

    L = []
    w = L.append
    w("# 로또 6/45 · 연금복권 720+ 리포트 (엔진 v2)")
    w(f"생성: {datetime.now():%Y-%m-%d %H:%M} / 데이터 로또 ~{lt - 1}회 · 연금 ~{pt - 1}회\n")

    w(f"## 1. 로또 {lt}회 추천 — {fmt_date(draw_date_for(lotto, lt))} 추첨")
    w("30개 번호를 겹치지 않게, 번호 구간 비율대로 배치 + 인기 번호를 덜 골라 1등 분할 회피")
    w("")
    w("| 세트 | 번호 | 예상 분할배율 |")
    w("|:--:|:--|--:|")
    for _, r in official.iterrows():
        w(f"| {r['set']} | **{fmt_nums(r[NUM_COLS])}** | {r['split_mult']:.2f} |")
    w("")
    w("분할배율: 1등 당첨 시 함께 당첨될 것으로 예상되는 인원의 배율 (무작위 조합 평균 = 1.00, 낮을수록 몫이 큼)\n")

    w(f"## 2. 연금복권 {pt}회 추천 — {fmt_date(draw_date_for(pension, pt))} 추첨 (품절 대비 {len(pcur)}순위)")
    w("앞 순위가 품절이면 다음 순위로. 끝자리가 10순위 단위로 0~9를 한 번씩 덮어, 앞에서 10장을 사면 7등 이상 확정.")
    w("")
    w("| 순위 | 조 | 번호 | 순위 | 조 | 번호 |")
    w("|---:|:--:|:--|---:|:--:|:--|")
    rows = list(pcur.itertuples())
    half = (len(rows) + 1) // 2
    for i in range(half):
        a = rows[i]
        cells = [str(a.rank), f"{a.jo}조", f"**{str(a.num).zfill(6)}**"]
        if i + half < len(rows):
            b = rows[i + half]
            cells += [str(b.rank), f"{b.jo}조", f"**{str(b.num).zfill(6)}**"]
        else:
            cells += ["", "", ""]
        w("| " + " | ".join(cells) + " |")
    w("")

    w("## 3. 성적표 (실전 기록)")
    last_t = int(lotto["epsd"].max())
    lastg = g[(g["target"] == last_t) & (g["strategy"] == "공식")]
    if not lastg.empty:
        wn = sorted(int(v) for v in lotto.iloc[-1][NUM_COLS])
        w(f"- 직전 {last_t}회 당첨번호 **{fmt_nums(wn)}** + 보너스 {int(lotto.iloc[-1]['bonus'])}")
        for _, r in lastg.iterrows():
            mark = " ".join(f"[{int(n)}]" if int(n) in wn else str(int(n)) for n in r[NUM_COLS])
            w(f"  - {r['engine']} {r['set']}세트: {mark} → {r['hits']}개 {r['rank']}")
    w("")
    w("| 구분 | 회차 | 장수 | 평균 적중 | 이론 | 당첨(5등↑/7등↑) | 이론 기대 | 적중 분포 |")
    w("|---|--:|--:|--:|--:|--:|--:|---|")
    for _, r in sb.iterrows():
        w(f"| {r['game']} {r['group']} | {r['draws']} | {r['tickets']} | {r['avg_hits']:.2f} | {r['theory_avg']:.2f} "
          f"| {r['prize_count']} | {r['theory_prize']:.1f} | {r['detail']} |")
    w("")
    w("대조군 = 같은 주에 완전 무작위로 뽑은 5세트(구매용 아님). 공식 추천과 적중률이 같게 나오는 것이 정상이다 —")
    w("어떤 방법도 적중 확률 자체는 바꾸지 못한다는 것을 매주 실전으로 확인하는 장치다.\n")

    w("## 4. 엔진 검증")
    if os.path.exists(BACKTEST_CSV):
        bt = pd.read_csv(BACKTEST_CSV)
        w(f"워크포워드 백테스트 — 과거 {int(bt['draws'].iloc[0])}회, 각 회차 직전 데이터만으로 추천 후 채점:\n")
        w("| 엔진 | 평균 적중 | 5등↑ 세트 비율 | 주당 5등↑ 확률* | 분할배율 | 구간 최대편차 | 20번대 비중 |")
        w("|---|--:|--:|--:|--:|--:|--:|")
        for _, r in bt.iterrows():
            w(f"| {r['engine']} | {r['avg_hits']:.3f} | {r['pct_prize']:.2f}% | {r['weekly_any_prize_mc']:.2f}% "
              f"| {r['split_mult']:.3f} | {r['range_dev']:.1f}% | {r['share_20s']:.2f}x |")
        w("\n\\* 몬테카를로 4만 회 추첨 기준(300주 실측은 표준오차 ±1.8%p라 구별 불가). 이론 평균 적중 0.800.\n")
    w(f"번호 인기도 모형: 5등 배율 표본 외 예측 상관 **r={val['r']:.2f}** (검증 {val['n_test']}회, p={val['p']:.0e}).")
    q = val["quintile_r1"]
    w(f"당첨번호 인기도 5분위별 실제 1등 당첨자(기대 대비, 최근 {val['n_quint']}회): "
      f"가장 비인기 {q[0]:.2f}x → {q[1]:.2f}x → {q[2]:.2f}x → {q[3]:.2f}x → 가장 인기 {q[4]:.2f}x\n")
    top = pop.sort_values("popularity", ascending=False)
    hi8, lo8 = top.head(8), top.tail(8).iloc[::-1]
    w(f"- 가장 많이 찍히는 번호: {', '.join(f'{n}({v:.3f})' for n, v in zip(hi8['number'], hi8['popularity']))}")
    w(f"- 가장 덜 찍히는 번호: {', '.join(f'{n}({v:.3f})' for n, v in zip(lo8['number'], lo8['popularity']))}")
    w(f"- 분할 회피 강도 {strength} (0=무작위, 1 이상이면 인기 번호가 거의 안 나와 편중으로 보임)\n")

    w("## 5. 참고 통계")
    w(f"- 로또 번호 균등성 카이제곱: p={ref['pval']:.3f} → 역대 출현 빈도는 무작위와 구별 불가 ('뜨거운 번호'는 잡음)")
    w(f"- 연금 자리별 균등성 p값: {', '.join(f'{v:.2f}' for v in pension_uniformity(pension))}")
    od = sorted(ref["overdue"].items(), key=lambda kv: -kv[1])[:8]
    w(f"- 최장 미출현(참고용, 예측력 없음): {', '.join(f'{n}번 {c}회' for n, c in od)}\n")
    w("> ⚠️ 어떤 분석도 당첨 확률(1등 1/8,145,060)을 바꾸지 못한다. 이 엔진이 개선하는 것은 ① 당첨 시 몫(분할 회피)")
    w("> ② 5세트가 같이 맞고 같이 틀리는 일을 줄이는 분산, 두 가지다. 구매는 여유 자금 내에서.")

    text = "\n".join(L)
    with open(REPORT_MD, "w", encoding="utf-8") as fp:
        fp.write(text)
    print(text)
    print(f"\n리포트 저장 → {REPORT_MD}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--update", action="store_true", help="데이터 갱신만 수행")
    ap.add_argument("--no-fetch", action="store_true", help="저장된 데이터로 추천·리포트만")
    ap.add_argument("--sets", type=int, default=5, help="로또 추천 세트 수")
    ap.add_argument("--pension", type=int, default=PENSION_N, help="연금 후보 수(품절 대비)")
    ap.add_argument("--strength", type=float, default=STRENGTH_DEFAULT, help="분할 회피 강도 (0~2)")
    args = ap.parse_args()

    if not args.no_fetch:
        update_lotto()
        update_pension()
    if not args.update:
        run_report(args.sets, args.pension, args.strength)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()

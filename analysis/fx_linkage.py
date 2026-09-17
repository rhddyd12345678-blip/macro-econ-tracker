"""원/달러 환율이 달러지수·엔/달러·달러/유로에 얼마나 연동됐는지 분석해
site/data/fx_analysis.json으로 내보낸다.

- 원/달러 월간 '로그 변화율'을 달러지수·엔/달러·달러/유로 월간 로그 변화율에
  회귀(수준값 회귀는 가짜 회귀 위험이 있어 쓰지 않음). 단순회귀 3개 + 다중회귀
  1개(VIF 포함) + 롤링 24개월 상관계수·R².
- 판정 규칙(가장 강한 연동/경계 신호/유의한 연동 없음)과 표준오차(Newey-West
  HAC), 롤링 유의수준 5% 기준선은 analysis/rate_linkage.py의 함수를 그대로
  재사용해 금리 연동 분석과 동일한 기준을 쓴다.
- 한미 정책금리차(월말)와 원/달러 수준·변화의 상관도 함께 계산한다.

data/fx_daily.csv(일별, fx_collect.py가 만듦)와 로우데이터 엑셀(정책금리만,
읽기 전용)을 입력으로 쓴다. collect.py 흐름에서 fx_collect.py 다음에 자동
실행된다.
"""

import csv
import json
import os
import sys
import warnings
from datetime import date, datetime

import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

warnings.simplefilter("ignore")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rate_linkage as rl  # noqa: E402  (newey_west_lags, simple_regression, rolling_regression,
# r2_significance_threshold, classify_strongest, load_rawdata, build_policy_monthly 재사용)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DAILY_CSV_PATH = os.path.join(BASE_DIR, "data", "fx_daily.csv")
FX_ANALYSIS_JSON_PATH = os.path.join(BASE_DIR, "site", "data", "fx_analysis.json")

OBS_START = "2021-09-01"
ROLLING_WINDOW = rl.ROLLING_WINDOW

FX_PAIR_CODE = {
    "원/달러": "KRW",
    "달러지수(광의)": "DXY",
    "엔/달러": "JPY",
    "달러/유로": "EUR",
}
PREDICTOR_CODES = ["DXY", "JPY", "EUR"]


def load_fx_daily():
    if not os.path.exists(DAILY_CSV_PATH):
        raise SystemExit(f"오류: {DAILY_CSV_PATH} 파일이 없습니다. fx_collect.py를 먼저 실행하세요.")
    with open(DAILY_CSV_PATH, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def build_fx_monthly_levels(rows):
    """월평균 수준값 DataFrame(index=월초 Timestamp, columns=KRW/DXY/JPY/EUR).
    완결되지 않은 이번 달은 제외한다."""
    current_month = date.today().strftime("%Y-%m")
    by_pair_month = {}
    for r in rows:
        code = FX_PAIR_CODE.get(r["통화쌍"])
        if not code:
            continue
        if r["날짜"] < OBS_START:
            continue
        mk = r["날짜"][:7]
        by_pair_month.setdefault(code, {}).setdefault(mk, []).append(float(r["값"]))

    data = {}
    for code, months in by_pair_month.items():
        col = {}
        for mk, vals in months.items():
            if mk >= current_month:
                continue
            col[pd.Timestamp(mk + "-01")] = sum(vals) / len(vals)
        data[code] = col
    return pd.DataFrame(data).sort_index()


def log_returns(levels):
    return np.log(levels).diff()


def analyze_fx(levels):
    diff = log_returns(levels).dropna(how="all")
    y = diff["KRW"]

    simple = {}
    scatter = {}
    for code in PREDICTOR_CODES:
        simple[code] = rl.simple_regression(y, diff[code])
        pair = pd.concat([diff[code], y], axis=1, keys=["x", "y"]).dropna()
        scatter[code] = [
            {"date": idx.date().isoformat(), "x": round(float(row["x"]), 6), "y": round(float(row["y"]), 6)}
            for idx, row in pair.iterrows()
        ]

    strongest, strongest_boundary, all_insignificant = rl.classify_strongest(simple)

    multi_df = pd.concat([y] + [diff[c] for c in PREDICTOR_CODES], axis=1, keys=["y"] + PREDICTOR_CODES).dropna()
    n_multi = len(multi_df)
    X = sm.add_constant(multi_df[PREDICTOR_CODES])
    lags = rl.newey_west_lags(n_multi)
    multi_model = sm.OLS(multi_df["y"], X).fit(cov_type="HAC", cov_kwds={"maxlags": lags})
    vif = {}
    for i, c in enumerate(PREDICTOR_CODES, start=1):
        vif[c] = float(variance_inflation_factor(X.values, i))
    multi = {
        "included": PREDICTOR_CODES,
        "coeffs": {k: float(v) for k, v in multi_model.params.items()},
        "p_values": {k: float(v) for k, v in multi_model.pvalues.items()},
        "vif": vif,
        "adj_r2": float(multi_model.rsquared_adj),
        "r2": float(multi_model.rsquared),
        "n": int(n_multi),
        "hac_lags": lags,
    }

    rolling = {
        "window": ROLLING_WINDOW,
        "beta": {},
        "r2": {},
        "corr": {},
        "p": {},
        "p_ols": {},
        "dates": {},
        "r2_threshold": rl.r2_significance_threshold(ROLLING_WINDOW),
    }
    for code in PREDICTOR_CODES:
        rr = rl.rolling_regression(y, diff[code])
        rolling["dates"][code] = rr["dates"]
        rolling["beta"][code] = rr["beta"]
        rolling["r2"][code] = rr["r2"]
        rolling["corr"][code] = rr["corr"]
        rolling["p"][code] = rr["p"]
        rolling["p_ols"][code] = rr["p_ols"]

    recent = {}
    for code in PREDICTOR_CODES:
        if rolling["beta"][code]:
            recent[code] = {
                "beta": rolling["beta"][code][-1],
                "r2": rolling["r2"][code][-1],
                "corr": rolling["corr"][code][-1],
                "p": rolling["p"][code][-1],
                "p_ols": rolling["p_ols"][code][-1],
                "as_of": rolling["dates"][code][-1],
                "window": ROLLING_WINDOW,
            }
        else:
            recent[code] = None
    strongest_recent, strongest_recent_boundary, recent_all_insignificant = rl.classify_strongest(recent)

    return {
        "simple": simple,
        "strongest": strongest,
        "strongest_boundary": strongest_boundary,
        "strongest_all_insignificant": all_insignificant,
        "recent": recent,
        "strongest_recent": strongest_recent,
        "strongest_recent_boundary": strongest_recent_boundary,
        "recent_all_insignificant": recent_all_insignificant,
        "scatter": scatter,
        "multi": multi,
        "rolling": rolling,
        "n_months_available": int(diff.dropna(how="all").shape[0]),
    }


def analyze_kr_us_spread_corr(levels):
    """한미 정책금리차(월말)와 원/달러 수준·변화의 상관."""
    raw_rows = rl.load_rawdata()
    policy_monthly = rl.build_policy_monthly(raw_rows)  # index=월말, columns=KR/US/JP
    spread = (policy_monthly["KR"] - policy_monthly["US"]).dropna()
    spread.index = spread.index.to_period("M").to_timestamp()  # 월말 -> 월초로 맞춰 원/달러(월초 인덱스)와 정렬

    krw_level = levels["KRW"].dropna()
    krw_change = log_returns(levels)["KRW"].dropna()
    spread_change = spread.diff().dropna()

    level_pair = pd.concat([spread, krw_level], axis=1, keys=["spread", "krw"]).dropna()
    change_pair = pd.concat([spread_change, krw_change], axis=1, keys=["spread_chg", "krw_chg"]).dropna()

    level_corr = float(level_pair["spread"].corr(level_pair["krw"])) if len(level_pair) > 5 else None
    change_corr = float(change_pair["spread_chg"].corr(change_pair["krw_chg"])) if len(change_pair) > 5 else None

    return {
        "level_corr": level_corr,
        "level_n": int(len(level_pair)),
        "change_corr": change_corr,
        "change_n": int(len(change_pair)),
        "series": {
            "dates": [d.date().isoformat() for d in level_pair.index],
            "spread_kr_us": [round(float(v), 3) for v in level_pair["spread"]],
            "krw_usd": [round(float(v), 2) for v in level_pair["krw"]],
        },
    }


def main():
    rows = load_fx_daily()
    levels = build_fx_monthly_levels(rows)

    fx = analyze_fx(levels)
    spread_corr = analyze_kr_us_spread_corr(levels)

    last_date_by_pair = {}
    for r in rows:
        code = FX_PAIR_CODE.get(r["통화쌍"])
        if not code:
            continue
        if code not in last_date_by_pair or r["날짜"] > last_date_by_pair[code]:
            last_date_by_pair[code] = r["날짜"]
    as_of = min(last_date_by_pair.values()) if last_date_by_pair else None

    result = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "as_of": as_of,
        "sample": {
            "start": OBS_START,
            "end": levels.index.max().date().isoformat() if len(levels) else None,
            "n_months_level": int(len(levels)),
        },
        "fx": fx,
        "kr_us_spread_corr": spread_corr,
    }

    os.makedirs(os.path.dirname(FX_ANALYSIS_JSON_PATH), exist_ok=True)
    with open(FX_ANALYSIS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"fx_linkage.py: {FX_ANALYSIS_JSON_PATH} 생성")
    print(f"  가장 강한 연동(전체 기간): {fx['strongest']}")
    print(f"  가장 강한 연동(최근 24개월): {fx['strongest_recent']}")
    print(f"  다중회귀 R²: {fx['multi']['r2']:.3f}, VIF: {fx['multi']['vif']}")
    print(f"  한미 정책금리차 vs 원/달러: 수준 상관={spread_corr['level_corr']}, 변화 상관={spread_corr['change_corr']}")


if __name__ == "__main__":
    main()

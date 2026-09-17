"""한국 금리(국채10년·기준금리)가 미국·일본·독일·프랑스 금리에 얼마나 연동됐는지
분석해 site/data/analysis.json으로 내보낸다.

- 분석 A(시장금리): 국채10년 '월간 변화(차분, %p)'끼리 회귀(수준값 회귀는 가짜
  회귀 위험이 있어 쓰지 않음). 단순회귀 4개 + 다중회귀 1개(공선성 높으면 프랑스
  제외) + 롤링 24개월 + 동월/1개월 선행 비교.
- 분석 B(정책금리): 월말 기준 정책금리 시계열, 한미·한일 스프레드, 교차상관.
- 표준오차는 Newey-West(HAC).

엑셀은 읽기만 하며(openpyxl read_only), 전혀 수정하지 않는다. collect.py 실행
끝에 export.py 다음으로 자동 호출된다: `python3 analysis/rate_linkage.py`.
"""

import json
import os
import warnings
from datetime import date, datetime

import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

warnings.simplefilter("ignore")
import openpyxl  # noqa: E402

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(BASE_DIR, "매크로_트래커_로우데이터.xlsx")
ANALYSIS_JSON_PATH = os.path.join(BASE_DIR, "site", "data", "analysis.json")

OBS_START = "2021-09-01"
ROLLING_WINDOW = 24
FR_DE_COLLINEARITY_THRESHOLD = 0.8  # |corr| 또는 VIF>10 이면 프랑스를 다중회귀에서 제외
VIF_THRESHOLD = 10.0

BOND_COUNTRIES = {"한국": "KR", "미국": "US", "일본": "JP", "독일": "DE", "프랑스": "FR"}
POLICY_INDICATORS = {
    "한국": "기준금리",
    "미국": "정책금리(상단)",
    "일본": "정책금리",
}


def load_rawdata():
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    ws = wb["로우데이터"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        base_date, country, indicator, value = row[0], row[1], row[2], row[3]
        if hasattr(base_date, "date"):
            base_date = base_date.date()
        rows.append({"date": base_date, "country": country, "indicator": indicator, "value": value})
    wb.close()
    return rows


def build_bond_levels(rows):
    """월별 국채10년 수준값 DataFrame(index=월초 Timestamp, columns=KR/US/JP/DE/FR)."""
    data = {}
    for r in rows:
        if r["indicator"] != "국채10년":
            continue
        code = BOND_COUNTRIES.get(r["country"])
        if not code:
            continue
        if r["date"].isoformat() < OBS_START:
            continue
        data.setdefault(code, {})[pd.Timestamp(r["date"])] = float(r["value"])
    df = pd.DataFrame(data).sort_index()
    return df


def build_policy_monthly(rows):
    """월말 기준 정책금리 시계열(index=월말 Timestamp, columns=KR/US/JP). 각국 최초
    행부터 마지막 관측월까지, 변경 시점 값을 다음 변경 전까지 그대로 이어간다."""
    code_by_country = {"한국": "KR", "미국": "US", "일본": "JP"}
    events = {code: [] for code in code_by_country.values()}
    for r in rows:
        country = r["country"]
        if country not in POLICY_INDICATORS:
            continue
        if r["indicator"] != POLICY_INDICATORS[country]:
            continue
        events[code_by_country[country]].append((pd.Timestamp(r["date"]), float(r["value"])))
    for code in events:
        events[code].sort(key=lambda x: x[0])

    last_date = max(max(d for d, _ in ev) for ev in events.values())
    month_ends = pd.date_range(pd.Timestamp(OBS_START), last_date, freq="M")

    series = {}
    for code, ev in events.items():
        vals = []
        for me in month_ends:
            v = None
            for d, val in ev:
                if d <= me:
                    v = val
                else:
                    break
            vals.append(v)
        series[code] = vals
    return pd.DataFrame(series, index=month_ends)


def newey_west_lags(n):
    return max(1, int(np.floor(4 * (n / 100) ** (2 / 9))))


def simple_regression(y, x):
    """y, x: pandas Series(공통 인덱스, dropna 적용됨). 반환: dict(beta, se, t, p, r2, n)."""
    df = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    n = len(df)
    if n < 8:
        return {"n": n, "insufficient": True}
    X = sm.add_constant(df["x"])
    lags = newey_west_lags(n)
    model = sm.OLS(df["y"], X).fit(cov_type="HAC", cov_kwds={"maxlags": lags})
    return {
        "beta": float(model.params["x"]),
        "const": float(model.params["const"]),
        "se": float(model.bse["x"]),
        "t": float(model.tvalues["x"]),
        "p": float(model.pvalues["x"]),
        "r2": float(model.rsquared),
        "n": int(n),
        "hac_lags": lags,
    }


def rolling_regression(y, x, window=ROLLING_WINDOW):
    df = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    dates, betas, r2s = [], [], []
    for end in range(window, len(df) + 1):
        chunk = df.iloc[end - window : end]
        X = sm.add_constant(chunk["x"])
        try:
            model = sm.OLS(chunk["y"], X).fit()
        except Exception:  # noqa: BLE001
            continue
        dates.append(chunk.index[-1].date().isoformat())
        betas.append(float(model.params["x"]))
        r2s.append(float(model.rsquared))
    return {"dates": dates, "beta": betas, "r2": r2s}


def analyze_market(bond_levels):
    diff = bond_levels.diff().dropna(how="all")
    kr = diff["KR"]

    simple = {}
    scatter = {}
    for code in ["US", "JP", "DE", "FR"]:
        simple[code] = simple_regression(kr, diff[code])
        pair = pd.concat([diff[code], kr], axis=1, keys=["x", "y"]).dropna()
        scatter[code] = [
            {"date": idx.date().isoformat(), "x": round(float(row["x"]), 4), "y": round(float(row["y"]), 4)}
            for idx, row in pair.iterrows()
        ]

    valid_r2 = {c: v["r2"] for c, v in simple.items() if not v.get("insufficient")}
    strongest = max(valid_r2, key=valid_r2.get) if valid_r2 else None

    de_fr_corr_df = pd.concat([diff["DE"], diff["FR"]], axis=1, keys=["DE", "FR"]).dropna()
    de_fr_corr = float(de_fr_corr_df["DE"].corr(de_fr_corr_df["FR"])) if len(de_fr_corr_df) > 5 else None

    full_predictors = ["US", "JP", "DE", "FR"]
    full_df = pd.concat([kr] + [diff[c] for c in full_predictors], axis=1, keys=["y"] + full_predictors).dropna()
    full_vif = {}
    if len(full_df) > len(full_predictors) + 2:
        Xf = sm.add_constant(full_df[full_predictors])
        for i, c in enumerate(full_predictors, start=1):
            full_vif[c] = float(variance_inflation_factor(Xf.values, i))

    exclude_fr = (de_fr_corr is not None and abs(de_fr_corr) > FR_DE_COLLINEARITY_THRESHOLD) or (
        full_vif.get("FR", 0) > VIF_THRESHOLD
    )
    predictors = ["US", "JP", "DE"] if exclude_fr else ["US", "JP", "DE", "FR"]

    multi_df = pd.concat([kr] + [diff[c] for c in predictors], axis=1, keys=["y"] + predictors).dropna()
    n_multi = len(multi_df)
    X = sm.add_constant(multi_df[predictors])
    lags = newey_west_lags(n_multi)
    multi_model = sm.OLS(multi_df["y"], X).fit(cov_type="HAC", cov_kwds={"maxlags": lags})
    vif = {}
    for i, c in enumerate(predictors, start=1):
        vif[c] = float(variance_inflation_factor(X.values, i))

    multi = {
        "included": predictors,
        "excluded": (
            {"FR": {"reason": "독일과 공선성이 높아 제외", "corr_with_DE": de_fr_corr, "vif_in_full_model": full_vif.get("FR")}}
            if exclude_fr
            else {}
        ),
        "coeffs": {k: float(v) for k, v in multi_model.params.items()},
        "p_values": {k: float(v) for k, v in multi_model.pvalues.items()},
        "vif": vif,
        "adj_r2": float(multi_model.rsquared_adj),
        "r2": float(multi_model.rsquared),
        "n": int(n_multi),
        "hac_lags": lags,
    }

    rolling = {"window": ROLLING_WINDOW, "beta": {}, "r2": {}, "dates": {}}
    for code in ["US", "JP", "DE", "FR"]:
        rr = rolling_regression(kr, diff[code])
        rolling["dates"][code] = rr["dates"]
        rolling["beta"][code] = rr["beta"]
        rolling["r2"][code] = rr["r2"]

    lead_lag = {}
    for code in ["US", "JP", "DE", "FR"]:
        contemporaneous = simple[code]
        lead1 = simple_regression(kr, diff[code].shift(1))
        lead_lag[code] = {"contemporaneous": contemporaneous, "lead1": lead1}

    return {
        "simple": simple,
        "strongest": strongest,
        "scatter": scatter,
        "multi": multi,
        "rolling": rolling,
        "lead_lag": lead_lag,
        "n_months_available": int(diff.dropna(how="all").shape[0]),
    }


def cross_correlation(y, x, max_lag=6):
    """lag>0: x가 y보다 lag개월 선행(과거 x가 현재 y와 상관)."""
    lags = list(range(-max_lag, max_lag + 1))
    corrs = []
    df = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    for lag in lags:
        shifted = df["x"].shift(lag)
        pair = pd.concat([df["y"], shifted], axis=1).dropna()
        if len(pair) < 8:
            corrs.append(None)
            continue
        c = pair.iloc[:, 0].corr(pair.iloc[:, 1])
        corrs.append(None if pd.isna(c) else float(c))
    return {"lags": lags, "corr": corrs}


def analyze_policy(policy_monthly):
    kr_us_spread = (policy_monthly["KR"] - policy_monthly["US"]).dropna()
    kr_jp_spread = (policy_monthly["KR"] - policy_monthly["JP"]).dropna()

    d_kr = policy_monthly["KR"].diff()
    d_us = policy_monthly["US"].diff()
    cc = cross_correlation(d_kr, d_us, max_lag=6)

    series = {
        "dates": [d.date().isoformat() for d in policy_monthly.index],
        "KR": [None if pd.isna(v) else float(v) for v in policy_monthly["KR"]],
        "US": [None if pd.isna(v) else float(v) for v in policy_monthly["US"]],
        "JP": [None if pd.isna(v) else float(v) for v in policy_monthly["JP"]],
    }
    spread = {
        "dates": [d.date().isoformat() for d in kr_us_spread.index],
        "KR_US": [float(v) for v in kr_us_spread],
    }
    spread_jp = {
        "dates": [d.date().isoformat() for d in kr_jp_spread.index],
        "KR_JP": [float(v) for v in kr_jp_spread],
    }

    n_changes = {
        "KR": int((policy_monthly["KR"].diff().fillna(0) != 0).sum()),
        "US": int((policy_monthly["US"].diff().fillna(0) != 0).sum()),
        "JP": int((policy_monthly["JP"].diff().fillna(0) != 0).sum()),
    }

    return {
        "series": series,
        "spread_kr_us": spread,
        "spread_kr_jp": spread_jp,
        "cross_corr_kr_us": cc,
        "n_rate_changes": n_changes,
        "caveat": "정책금리는 추적 기간 중 변경 횟수가 적어(한국 " + str(n_changes["KR"]) + "회, 미국 "
        + str(n_changes["US"]) + "회, 일본 " + str(n_changes["JP"]) + "회) 통계적 신뢰도가 시장금리 분석보다 낮음.",
    }


def main():
    rows = load_rawdata()
    bond_levels = build_bond_levels(rows)
    policy_monthly = build_policy_monthly(rows)

    market = analyze_market(bond_levels)
    policy = analyze_policy(policy_monthly)

    last_date = max(r["date"] for r in rows if r["indicator"] == "국채10년" and r["country"] == "한국")

    result = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "as_of": last_date.isoformat(),
        "sample": {
            "start": OBS_START,
            "end": bond_levels.index.max().date().isoformat(),
            "n_months_level": int(len(bond_levels)),
        },
        "market": market,
        "policy": policy,
    }

    os.makedirs(os.path.dirname(ANALYSIS_JSON_PATH), exist_ok=True)
    with open(ANALYSIS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"rate_linkage.py: {ANALYSIS_JSON_PATH} 생성 (기준일 {last_date.isoformat()})")
    print(f"  가장 강한 연동: {market['strongest']} (R²={market['simple'].get(market['strongest'], {}).get('r2')})")
    print(f"  다중회귀 포함 변수: {market['multi']['included']}")


if __name__ == "__main__":
    main()

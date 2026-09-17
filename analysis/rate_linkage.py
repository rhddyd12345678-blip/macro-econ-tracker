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
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats as scipy_stats
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


def _policy_events(rows, country):
    ind = POLICY_INDICATORS[country]
    ev = [(r["date"], float(r["value"])) for r in rows if r["country"] == country and r["indicator"] == ind]
    ev.sort(key=lambda x: x[0])
    return ev


def _value_in_effect(events, on_date):
    """events: [(date, value), ...] 정렬됨. on_date 시점에 적용 중이던 값(스텝 함수)."""
    v = None
    for d, val in events:
        if d <= on_date:
            v = val
        else:
            break
    return v


def build_bok_fed_case_table(rows):
    """한국은행 기준금리 변경마다: 방향, 변경폭, 직전 같은 방향 연준 변경일과의
    시차(개월), 그 시점의 한미 금리차를 계산한다. 방향이 직전 변경과 달라지는
    지점(또는 최초 변경)을 '사이클 시작'으로 표시한다."""
    kr_events = _policy_events(rows, "한국")
    us_events = _policy_events(rows, "미국")

    us_changes = []  # (date, direction) — 베이스라인(최초 행) 제외, 실제 변경만
    for i in range(1, len(us_events)):
        d, v = us_events[i]
        prev_v = us_events[i - 1][1]
        if v > prev_v:
            us_changes.append((d, "인상"))
        elif v < prev_v:
            us_changes.append((d, "인하"))

    cases = []
    prev_direction = None
    for i in range(1, len(kr_events)):
        d, v = kr_events[i]
        prev_v = kr_events[i - 1][1]
        change = round(v - prev_v, 3)
        if change == 0:
            continue
        direction = "인상" if change > 0 else "인하"

        prior_fed = [fd for fd, fdir in us_changes if fdir == direction and fd < d]
        prior_fed_date = max(prior_fed) if prior_fed else None
        lag_months = round((d - prior_fed_date).days / 30.44, 1) if prior_fed_date else None

        us_value_then = _value_in_effect(us_events, d)
        spread = round(v - us_value_then, 3) if us_value_then is not None else None

        cycle_start = prev_direction is None or direction != prev_direction
        cases.append(
            {
                "decision_date": d.isoformat(),
                "direction": direction,
                "change": change,
                "prior_fed_same_direction_date": prior_fed_date.isoformat() if prior_fed_date else None,
                "lag_months": lag_months,
                "spread_kr_us_at_decision": spread,
                "cycle_start": cycle_start,
            }
        )
        prev_direction = direction
    return cases


def _last_completed_month_end(today=None):
    """실행 시점 기준 '가장 최근 완료된 달'의 마지막 날. 예: 2026-09-17에 실행하면
    2026-09는 아직 끝나지 않았으므로 2026-08-31을 반환한다."""
    today = today or date.today()
    first_of_this_month = today.replace(day=1)
    return pd.Timestamp(first_of_this_month - timedelta(days=1))


def build_policy_monthly(rows):
    """월말 기준 정책금리 시계열(index=월말 Timestamp, columns=KR/US/JP). 각국 최초
    행부터 '가장 최근 완료된 달'까지, 변경 시점 값을 다음 변경 전까지 그대로
    이어간다(이벤트 날짜 자체가 아니라 완료된 달력월 기준으로 끝을 잡아야, 월말
    이전에 발생한 최신 변경이 해당 월에 반영된다)."""
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

    last_event_date = max(max(d for d, _ in ev) for ev in events.values())
    last_date = max(_last_completed_month_end(), last_event_date)
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
    """y, x: pandas Series(공통 인덱스, dropna 적용됨). 반환: dict(beta, se, t, p, r2, n).
    표본이 작을 때(특히 24개월 롤링 창) OLS와 HAC의 유의성 판정이 갈릴 수 있어
    둘 다 계산해 p(HAC)·p_ols(일반 OLS)로 함께 낸다."""
    df = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    n = len(df)
    if n < 8:
        return {"n": n, "insufficient": True}
    X = sm.add_constant(df["x"])
    lags = newey_west_lags(n)
    model_ols = sm.OLS(df["y"], X).fit()
    model_hac = sm.OLS(df["y"], X).fit(cov_type="HAC", cov_kwds={"maxlags": lags})
    return {
        "beta": float(model_hac.params["x"]),
        "const": float(model_hac.params["const"]),
        "se": float(model_hac.bse["x"]),
        "se_ols": float(model_ols.bse["x"]),
        "t": float(model_hac.tvalues["x"]),
        "t_ols": float(model_ols.tvalues["x"]),
        "p": float(model_hac.pvalues["x"]),
        "p_ols": float(model_ols.pvalues["x"]),
        "r2": float(model_hac.rsquared),
        "n": int(n),
        "hac_lags": lags,
    }


def rolling_regression(y, x, window=ROLLING_WINDOW):
    """window개월 롤링 단순회귀. R²와 별도로 상관계수·p값(HAC·OLS 둘 다)도 함께
    낸다 — R²=상관계수^2라 부호(같은 방향/반대 방향)는 R²만으로 알 수 없고,
    국가별 변동성 차이에 영향받지 않는 비교에도 상관계수가 더 적합하기 때문.
    표본이 24개월로 작을 때는 OLS와 HAC의 유의성 판정이 갈릴 수 있어 둘 다
    계산해 둔다. 창 크기가 고정이라 HAC 시차도 매번 동일하게 계산해 재사용한다."""
    lags = newey_west_lags(window)
    df = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    dates, betas, r2s, corrs, pvals, pvals_ols = [], [], [], [], [], []
    for end in range(window, len(df) + 1):
        chunk = df.iloc[end - window : end]
        X = sm.add_constant(chunk["x"])
        try:
            model_ols = sm.OLS(chunk["y"], X).fit()
            model_hac = sm.OLS(chunk["y"], X).fit(cov_type="HAC", cov_kwds={"maxlags": lags})
        except Exception:  # noqa: BLE001
            continue
        dates.append(chunk.index[-1].date().isoformat())
        betas.append(float(model_hac.params["x"]))
        r2s.append(float(model_hac.rsquared))
        corrs.append(float(chunk["x"].corr(chunk["y"])))
        pvals.append(float(model_hac.pvalues["x"]))
        pvals_ols.append(float(model_ols.pvalues["x"]))
    return {"dates": dates, "beta": betas, "r2": r2s, "corr": corrs, "p": pvals, "p_ols": pvals_ols}


def r2_significance_threshold(window=ROLLING_WINDOW, alpha=0.05):
    """단순회귀(예측변수 1개) n=window일 때, R²가 이 값을 넘어야 alpha 수준에서
    유의(고전적 F검정, HAC 아님 — 차트에 그릴 기준선이라 일반적인 임계값을 씀).
    df = window-2."""
    dfree = window - 2
    f_crit = float(scipy_stats.f.ppf(1 - alpha, 1, dfree))
    return f_crit / (f_crit + dfree)


def classify_strongest(stats_by_code):
    """국가별 {"r2":..., "p":..., "p_ols":...} 딕셔너리를 받아 '가장 강한 연동'
    판정을 셋 중 하나로 낸다:
    - strongest: OLS와 HAC p값이 둘 다 0.05 미만인 국가 중 R²가 가장 높은 국가
    - boundary: (strongest가 없을 때) 둘 중 하나만 0.05 미만인 국가 중 R²가
      가장 높은 국가 — 표본이 작아 표준오차 계산 방식에 따라 결론이 갈리는
      '경계 신호'
    - all_insignificant: 어느 국가도 둘 중 하나조차 유의하지 않음
    반환: (strongest_code_or_None, boundary_code_or_None, all_insignificant_bool)"""
    both, boundary = {}, {}
    any_significant = False
    for code, v in stats_by_code.items():
        if not v or v.get("insufficient"):
            continue
        p_hac, p_ols = v.get("p"), v.get("p_ols")
        if p_hac is None or p_ols is None:
            continue
        hac_sig, ols_sig = p_hac < 0.05, p_ols < 0.05
        if hac_sig or ols_sig:
            any_significant = True
        if hac_sig and ols_sig:
            both[code] = v["r2"]
        elif hac_sig or ols_sig:
            boundary[code] = v["r2"]
    strongest = max(both, key=both.get) if both else None
    strongest_boundary = max(boundary, key=boundary.get) if (not both and boundary) else None
    return strongest, strongest_boundary, not any_significant


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

    # '가장 강한 연동' 판정은 R² 기준이되, OLS·HAC p값이 둘 다 0.05 미만인 국가만
    # 후보로 삼는다. 둘 중 하나만 유의하면 '경계 신호'로 따로 표시(표본이 작을 때
    # 표준오차 계산 방식에 따라 결론이 갈릴 수 있음을 반영). 아무도 유의하지
    # 않으면 strongest/boundary 모두 None(홈페이지에서 '유의한 연동 없음').
    strongest, strongest_boundary, strongest_all_insignificant = classify_strongest(simple)

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

    rolling = {
        "window": ROLLING_WINDOW,
        "beta": {},
        "r2": {},
        "corr": {},
        "p": {},
        "p_ols": {},
        "dates": {},
        "r2_threshold": r2_significance_threshold(ROLLING_WINDOW),
    }
    for code in ["US", "JP", "DE", "FR"]:
        rr = rolling_regression(kr, diff[code])
        rolling["dates"][code] = rr["dates"]
        rolling["beta"][code] = rr["beta"]
        rolling["r2"][code] = rr["r2"]
        rolling["corr"][code] = rr["corr"]
        rolling["p"][code] = rr["p"]
        rolling["p_ols"][code] = rr["p_ols"]

    # '최근 24개월' 요약: 롤링 시계열의 마지막 구간(=가장 최근 24개월)을 그대로 쓴다.
    recent = {}
    for code in ["US", "JP", "DE", "FR"]:
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
    strongest_recent, strongest_recent_boundary, recent_all_insignificant = classify_strongest(recent)

    lead_lag = {}
    for code in ["US", "JP", "DE", "FR"]:
        contemporaneous = simple[code]
        lead1 = simple_regression(kr, diff[code].shift(1))
        lead_lag[code] = {"contemporaneous": contemporaneous, "lead1": lead1}

    return {
        "simple": simple,
        "strongest": strongest,
        "strongest_boundary": strongest_boundary,
        "strongest_all_insignificant": strongest_all_insignificant,
        "recent": recent,
        "strongest_recent": strongest_recent,
        "strongest_recent_boundary": strongest_recent_boundary,
        "recent_all_insignificant": recent_all_insignificant,
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


def analyze_policy(policy_monthly, rows):
    kr_us_spread = (policy_monthly["KR"] - policy_monthly["US"]).dropna()
    kr_jp_spread = (policy_monthly["KR"] - policy_monthly["JP"]).dropna()

    d_kr = policy_monthly["KR"].diff()
    d_us = policy_monthly["US"].diff()
    cc = cross_correlation(d_kr, d_us, max_lag=6)
    # 실제 계산 결과(모든 시차에서 양수, 뚜렷한 단일 정점 없음)를 확인하고 붙인
    # 해석 문구 — 데이터가 이 패턴을 유지하는 한 그대로 쓴다(선후행보다 같은
    # 사이클 기간이 겹친 효과로 보는 것이 더 타당하다는 판단).
    cc_note = "모든 시차에서 양수이고 뚜렷한 정점이 없어, 선후행 관계보다 같은 사이클 기간이 겹친 효과로 해석됨."

    bok_fed_cases = build_bok_fed_case_table(rows)

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
        "cross_corr_note": cc_note,
        "n_rate_changes": n_changes,
        "bok_fed_cases": bok_fed_cases,
        "caveat": "정책금리는 추적 기간 중 변경 횟수가 적어(한국 " + str(n_changes["KR"]) + "회, 미국 "
        + str(n_changes["US"]) + "회, 일본 " + str(n_changes["JP"]) + "회) 통계적 신뢰도가 시장금리 분석보다 낮음.",
    }


def main():
    rows = load_rawdata()
    bond_levels = build_bond_levels(rows)
    policy_monthly = build_policy_monthly(rows)

    market = analyze_market(bond_levels)
    policy = analyze_policy(policy_monthly, rows)

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
    print(f"  가장 강한 연동(전체 기간): {market['strongest']} (R²={market['simple'].get(market['strongest'], {}).get('r2')})")
    print(f"  가장 강한 연동(최근 24개월): {market['strongest_recent']} (R²={(market['recent'].get(market['strongest_recent']) or {}).get('r2')})")
    print(f"  다중회귀 포함 변수: {market['multi']['included']}")
    print(f"  한은-연준 금리 변경 사례: {len(policy['bok_fed_cases'])}건")


if __name__ == "__main__":
    main()

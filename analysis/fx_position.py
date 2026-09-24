"""환율 탭 '환율 위치 분석' 섹션을 위한 사실 기반 통계 분석.

이 스크립트는 매수·매도 신호나 투자 권유를 만들지 않는다. 원/달러 환율의
과거 대비 위치(백분위·이격도·변동성)와 과거 통계(분포·유사 국면·단순 규칙
백테스트)를 있는 그대로 계산해 site/data/fx_position.json으로 내보낼 뿐이며,
판단은 이 데이터를 보는 사람의 몫이다.

data/fx_daily.csv(원/달러·달러지수·엔/달러 등 일별, fx_collect.py가 만듦)와
로우데이터(한미 정책금리차·두바이유, 읽기 전용)를 입력으로 쓴다. collect.py
흐름에서 fx_linkage.py 다음에 자동 실행된다.

방법론 메모(설계 판단, 최종 JSON에도 일부 요약해 넣음):
- data/fx_daily.csv는 2021-09-01부터 시작해 아직 만 5년이 안 됐다. '5년'
  구간을 요구하는 지표(백분위·z-score·유사 국면 판정)는 실제로는
  '데이터 시작일부터 현재까지'로 근사한 값이며, 과거 시점 각각에 대해서도
  그 시점까지 쌓인 데이터만으로 '5년 백분위'를 계산한다(완전한 5년 이동창이
  아니라 확장창에 가깝다 — 초기 시점일수록 표본이 적어 값이 불안정할 수 있다).
- 회귀·상관은 로그 변화율/차분만 쓴다(수준값은 가짜 회귀 위험).
- 모든 기간(1/3/6개월)은 거래일이 아니라 달력월 기준이며, 목표일 ±5일 이내
  가장 가까운 실제 관측치를 그 시점 값으로 쓴다.
"""

import csv
import json
import os
import sys
import warnings
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
from scipy import stats as scipy_stats

warnings.simplefilter("ignore")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rate_linkage as rl  # noqa: E402 (load_rawdata, build_policy_monthly 재사용)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DAILY_CSV_PATH = os.path.join(BASE_DIR, "data", "fx_daily.csv")
FX_POSITION_JSON_PATH = os.path.join(BASE_DIR, "site", "data", "fx_position.json")

OBS_START = "2021-09-01"
SWING_THRESHOLD = 0.03  # 스윙 고점·저점 판정 기준(직전 극값 대비 3% 되돌림)
TRADING_DAYS_PER_YEAR = 252
NEAREST_TOLERANCE_DAYS = 5  # 목표 달력일 ±5일 이내 관측치만 그 시점 값으로 인정
HORIZON_MONTHS = [1, 3, 6]
ROLLING_CORR_WINDOW = 12  # 개월

PAIR_CODE = {"원/달러": "KRW", "달러지수(광의)": "DXY", "엔/달러": "JPY"}


# ---------------------------------------------------------------------------
# 데이터 로드
# ---------------------------------------------------------------------------


def load_fx_daily_rows():
    if not os.path.exists(DAILY_CSV_PATH):
        raise SystemExit(f"오류: {DAILY_CSV_PATH} 파일이 없습니다. fx_collect.py를 먼저 실행하세요.")
    with open(DAILY_CSV_PATH, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def build_daily_series(rows):
    """{code: pd.Series(index=DatetimeIndex 오름차순, value=float)} — KRW/DXY/JPY."""
    data = {code: {} for code in PAIR_CODE.values()}
    for r in rows:
        code = PAIR_CODE.get(r["통화쌍"])
        if not code:
            continue
        data[code][pd.Timestamp(r["날짜"])] = float(r["값"])
    return {code: pd.Series(d).sort_index() for code, d in data.items()}


def build_dubai_oil_monthly(raw_rows):
    """두바이유 월별 시계열(index=월초 Timestamp). 한국 행 기준(일본 행과 값 동일,
    같은 국제 유가를 국가별로 복제 저장해 둔 구조)."""
    vals = {}
    for r in raw_rows:
        if r["country"] == "한국" and r["indicator"] == "두바이유" and r["value"] is not None:
            d = r["date"]
            vals[pd.Timestamp(d.isoformat()[:7] + "-01")] = float(r["value"])
    return pd.Series(vals).sort_index()


# ---------------------------------------------------------------------------
# 공통 유틸
# ---------------------------------------------------------------------------


def nearest_value(dates_arr, values_arr, target_date, tolerance_days=NEAREST_TOLERANCE_DAYS):
    """target_date에 가장 가까운 (날짜, 값). 허용오차(일) 밖이면 (None, None)."""
    if len(dates_arr) == 0:
        return None, None
    target64 = np.datetime64(target_date)
    diffs = np.abs((dates_arr - target64) / np.timedelta64(1, "D"))
    pos = int(diffs.argmin())
    if diffs[pos] > tolerance_days:
        return None, None
    return dates_arr[pos], float(values_arr[pos])


def percentile_rank(window_values, value):
    """value가 window_values(같은 시리즈의 과거값 포함) 내에서 차지하는 백분위(0~100)."""
    arr = np.asarray(window_values, dtype=float)
    arr = arr[~np.isnan(arr)]
    if len(arr) == 0:
        return None
    return float((arr <= value).sum()) / len(arr) * 100.0


def expanding_5y_percentile_series(krw):
    """각 날짜 t에 대해 [t-5년, t] 구간(데이터 시작일 이후로 잘림) 내에서 그날 값의
    백분위. 초기 구간은 표본이 적어(확장창) 5년 미만일 수 있음 — 방법론 메모 참고."""
    dates = krw.index
    values = krw.values
    out = np.full(len(values), np.nan)
    start_pos = 0
    cutoff_offset = pd.DateOffset(years=5)
    for i in range(len(values)):
        cutoff = dates[i] - cutoff_offset
        while dates[start_pos] < cutoff:
            start_pos += 1
        window = values[start_pos : i + 1]
        out[i] = percentile_rank(window, values[i])
    return pd.Series(out, index=dates)


# ---------------------------------------------------------------------------
# (1) 위치 지표
# ---------------------------------------------------------------------------


def compute_position_indicators(krw, pct5y_series):
    last_date = krw.index[-1]
    last_val = float(krw.iloc[-1])
    out = {"as_of": last_date.date().isoformat(), "value": round(last_val, 2)}

    for label, years in [("1y", 1), ("3y", 3), ("5y", 5)]:
        cutoff = last_date - pd.DateOffset(years=years)
        window = krw[krw.index >= cutoff]
        out[f"percentile_{label}"] = round(percentile_rank(window.values, last_val), 1)
        out[f"n_{label}"] = int(len(window))

    cutoff_52w = last_date - pd.DateOffset(weeks=52)
    window_52w = krw[krw.index >= cutoff_52w]
    hi52, lo52 = float(window_52w.max()), float(window_52w.min())
    out["high_52w"] = round(hi52, 2)
    out["low_52w"] = round(lo52, 2)
    out["dev_from_high_52w_pct"] = round((last_val / hi52 - 1) * 100, 3)
    out["dev_from_low_52w_pct"] = round((last_val / lo52 - 1) * 100, 3)

    for w in (20, 60, 120):
        ma = krw.rolling(w).mean().iloc[-1]
        out[f"ma{w}"] = round(float(ma), 2) if pd.notna(ma) else None
        out[f"dev_from_ma{w}_pct"] = round((last_val / ma - 1) * 100, 3) if pd.notna(ma) else None

    cutoff_5y = last_date - pd.DateOffset(years=5)
    window_5y = krw[krw.index >= cutoff_5y]
    mean5y, std5y = float(window_5y.mean()), float(window_5y.std())
    out["mean_5y"] = round(mean5y, 2)
    out["std_5y"] = round(std5y, 2)
    out["z_score_5y"] = round((last_val - mean5y) / std5y, 3) if std5y else None

    out["percentile_5y_expanding_latest"] = round(float(pct5y_series.iloc[-1]), 1)
    return out


def detect_swings(krw, threshold=SWING_THRESHOLD):
    """직전 극값 대비 threshold(기본 3%) 이상 되돌림이 있어야 스윙 고점·저점으로
    확정한다(confirmed=True) — 이 되돌림이 일어나기 전까지는 그 극값이 스윙인지
    알 수 없으므로 모두 '사후 확정'이다. 아직 되돌림이 안 나와 확정되지 않은
    가장 최근 극값은 confirmed=False로 참고용만 포함한다."""
    dates = krw.index
    values = krw.values
    n = len(values)
    if n < 2:
        return []
    pivots = []
    trend = None
    run_max_idx, run_max_val = 0, values[0]
    run_min_idx, run_min_val = 0, values[0]
    ext_idx, ext_val = 0, values[0]
    for i in range(1, n):
        v = values[i]
        if trend is None:
            if v > run_max_val:
                run_max_val, run_max_idx = v, i
            if v < run_min_val:
                run_min_val, run_min_idx = v, i
            if run_max_idx > run_min_idx and run_max_val >= run_min_val * (1 + threshold):
                trend = "up"
                ext_idx, ext_val = run_max_idx, run_max_val
            elif run_min_idx > run_max_idx and run_min_val <= run_max_val * (1 - threshold):
                trend = "down"
                ext_idx, ext_val = run_min_idx, run_min_val
        elif trend == "up":
            if v > ext_val:
                ext_val, ext_idx = v, i
            elif v <= ext_val * (1 - threshold):
                pivots.append({"date": dates[ext_idx].date().isoformat(), "value": round(float(ext_val), 2), "type": "high", "confirmed": True})
                trend = "down"
                ext_idx, ext_val = i, v
        else:  # trend == "down"
            if v < ext_val:
                ext_val, ext_idx = v, i
            elif v >= ext_val * (1 + threshold):
                pivots.append({"date": dates[ext_idx].date().isoformat(), "value": round(float(ext_val), 2), "type": "low", "confirmed": True})
                trend = "up"
                ext_idx, ext_val = i, v
    if trend is not None:
        pivots.append({
            "date": dates[ext_idx].date().isoformat(), "value": round(float(ext_val), 2),
            "type": "high" if trend == "up" else "low", "confirmed": False,
        })
    return pivots


# ---------------------------------------------------------------------------
# (2) 달러 요인 분해
# ---------------------------------------------------------------------------


def build_dollar_decomposition(krw, dxy):
    df = pd.concat([krw, dxy], axis=1, keys=["KRW", "DXY"]).dropna()
    base_krw, base_dxy = float(df["KRW"].iloc[0]), float(df["DXY"].iloc[0])
    idx_krw = df["KRW"] / base_krw * 100
    idx_dxy = df["DXY"] / base_dxy * 100
    won_strength = df["KRW"] / df["DXY"]  # 원/달러 ÷ 달러지수
    base_won = float(won_strength.iloc[0])
    idx_won = won_strength / base_won * 100

    series = {
        "dates": [d.date().isoformat() for d in df.index],
        "krw_usd_index": [round(float(v), 3) for v in idx_krw],
        "dollar_index_index": [round(float(v), 3) for v in idx_dxy],
        "won_strength_index": [round(float(v), 3) for v in idx_won],
    }

    contributions = {}
    last_date = df.index[-1]
    for label, months in [("1m", 1), ("3m", 3), ("6m", 6)]:
        target = last_date - pd.DateOffset(months=months)
        pos = df.index.searchsorted(target)
        if pos >= len(df.index):
            contributions[label] = None
            continue
        start_date = df.index[max(0, min(pos, len(df.index) - 1))]
        if abs((start_date - target).days) > NEAREST_TOLERANCE_DAYS:
            contributions[label] = None
            continue
        krw0, krw1 = float(df["KRW"].loc[start_date]), float(df["KRW"].iloc[-1])
        dxy0, dxy1 = float(df["DXY"].loc[start_date]), float(df["DXY"].iloc[-1])
        won0, won1 = float(won_strength.loc[start_date]), float(won_strength.iloc[-1])
        total_log = np.log(krw1 / krw0)
        dollar_log = np.log(dxy1 / dxy0)
        won_log = np.log(won1 / won0)
        contributions[label] = {
            "start_date": start_date.date().isoformat(),
            "end_date": last_date.date().isoformat(),
            "total_pct": round((np.exp(total_log) - 1) * 100, 3),
            "dollar_factor_pct": round((np.exp(dollar_log) - 1) * 100, 3),
            "won_factor_pct": round((np.exp(won_log) - 1) * 100, 3),
            "dollar_factor_share_of_log": round(dollar_log / total_log, 3) if total_log else None,
        }

    return {"series": series, "contributions": contributions}


# ---------------------------------------------------------------------------
# (3) 드라이버: 롤링 12개월 상관계수
# ---------------------------------------------------------------------------


def build_monthly_frame(krw, jpy, oil_monthly, spread_monthly):
    """월말 기준 KRW/JPY/oil 수준값 + spread(정책금리차, 월말)를 하나의
    DataFrame으로 합친다(index=월말)."""
    current_month = pd.Timestamp(date.today().strftime("%Y-%m") + "-01")

    def monthly_eom(series):
        s = series.copy()
        s.index = s.index.to_period("M")
        eom = s.groupby(level=0).last()
        eom.index = eom.index.to_timestamp(how="end").normalize()
        return eom

    krw_m = monthly_eom(krw)
    jpy_m = monthly_eom(jpy)
    krw_m = krw_m[krw_m.index.to_period("M").to_timestamp() < current_month]
    jpy_m = jpy_m[jpy_m.index.to_period("M").to_timestamp() < current_month]

    oil_m = oil_monthly.copy()
    oil_m.index = oil_m.index.to_period("M").to_timestamp(how="end").normalize()

    spread_m = spread_monthly.copy()  # 이미 월말 인덱스(rl.build_policy_monthly)

    df = pd.concat([krw_m, jpy_m, oil_m, spread_m], axis=1, keys=["KRW", "JPY", "OIL", "SPREAD"])
    return df.sort_index()


def build_driver_correlations(monthly_df):
    df = monthly_df.dropna(how="all").copy()
    krw_logret = np.log(df["KRW"]).diff()
    jpy_logret = np.log(df["JPY"]).diff()
    oil_logret = np.log(df["OIL"]).diff()
    spread_diff = df["SPREAD"].diff()

    pairs = {"spread_kr_us": spread_diff, "dubai_oil": oil_logret, "jpy_usd": jpy_logret}
    out = {}
    for name, series in pairs.items():
        merged = pd.concat([krw_logret, series], axis=1, keys=["krw", "x"]).dropna()
        roll = merged["krw"].rolling(ROLLING_CORR_WINDOW).corr(merged["x"])
        roll = roll.dropna()
        out[name] = {
            "dates": [d.date().isoformat() for d in roll.index],
            "corr": [round(float(v), 3) for v in roll.values],
            "latest": round(float(roll.iloc[-1]), 3) if len(roll) else None,
        }
    return out


# ---------------------------------------------------------------------------
# (4) 변동성
# ---------------------------------------------------------------------------


def build_volatility(krw):
    log_ret = np.log(krw).diff().dropna()
    last_date = krw.index[-1]

    vol_series = {}
    latest_vol = {}
    pct_5y = {}
    median_5y = {}
    cutoff_5y = last_date - pd.DateOffset(years=5)
    for w in (20, 60, 120):
        rv = log_ret.rolling(w).std() * np.sqrt(TRADING_DAYS_PER_YEAR) * 100  # 연율화 %, 연 표준편차
        rv = rv.dropna()
        vol_series[str(w)] = {"dates": [d.date().isoformat() for d in rv.index], "value": [round(float(v), 3) for v in rv.values]}
        latest_vol[str(w)] = round(float(rv.iloc[-1]), 3) if len(rv) else None
        rv_5y = rv[rv.index >= cutoff_5y]
        pct_5y[str(w)] = round(percentile_rank(rv_5y.values, rv.iloc[-1]), 1) if len(rv_5y) else None
        median_5y[str(w)] = round(float(rv_5y.median()), 3) if len(rv_5y) else None

    ret_5y = log_ret[log_ret.index >= cutoff_5y]
    daily_stats = {
        "std_pct": round(float(ret_5y.std()) * 100, 4),
        "abs_p95_pct": round(float(np.percentile(np.abs(ret_5y.values), 95)) * 100, 4),
        "abs_p99_pct": round(float(np.percentile(np.abs(ret_5y.values), 99)) * 100, 4),
        "n": int(len(ret_5y)),
    }

    last_month_ret = log_ret[log_ret.index >= last_date - pd.DateOffset(months=1)]
    recent_month_stats = {
        "min_pct": round(float(last_month_ret.min()) * 100, 3) if len(last_month_ret) else None,
        "max_pct": round(float(last_month_ret.max()) * 100, 3) if len(last_month_ret) else None,
        "n": int(len(last_month_ret)),
    }

    # 향후 1/3/6개월 변화율의 과거 분포(예측이 아니라 실현된 과거 사례들의 분포) — 5년 표본
    forward_dist = {}
    dates_arr = krw.index.values
    values_arr = krw.values
    for months in HORIZON_MONTHS:
        rets = []
        for i, d in enumerate(krw.index):
            if d < cutoff_5y:
                continue
            target = d + pd.DateOffset(months=months)
            _, v_end = nearest_value(dates_arr, values_arr, target)
            if v_end is None:
                continue
            rets.append((v_end / values_arr[i] - 1) * 100)
        if rets:
            arr = np.array(rets)
            forward_dist[f"{months}m"] = {
                "p5": round(float(np.percentile(arr, 5)), 3),
                "p25": round(float(np.percentile(arr, 25)), 3),
                "p50": round(float(np.percentile(arr, 50)), 3),
                "p75": round(float(np.percentile(arr, 75)), 3),
                "p95": round(float(np.percentile(arr, 95)), 3),
                "n": int(len(arr)),
            }
        else:
            forward_dist[f"{months}m"] = None

    return {
        "series": vol_series,
        "latest": latest_vol,
        "percentile_5y": pct_5y,
        "median_5y": median_5y,
        "daily_return_5y": daily_stats,
        "recent_month": recent_month_stats,
        "forward_return_distribution_5y": forward_dist,
    }


# ---------------------------------------------------------------------------
# (5) 과거 유사 국면 + 양방향 최대 역행폭(MAE)
# ---------------------------------------------------------------------------


def mae_bidirectional(krw, entry_date, end_date):
    entry_val = float(krw.loc[entry_date])
    path = krw[(krw.index >= entry_date) & (krw.index <= end_date)]
    if path.empty:
        return None
    path_vals = path.values
    path_dates = path.index
    min_pos = int(np.argmin(path_vals))
    max_pos = int(np.argmax(path_vals))
    return {
        # '원화 약세(환율 상승)에 베팅'했을 때 중간에 최대 몇 % 하락했는지(음수 또는 0)
        "weak_won_bet_adverse_pct": round((path_vals[min_pos] / entry_val - 1) * 100, 3),
        "weak_won_bet_adverse_day": int((path_dates[min_pos] - entry_date).days),
        # '원화 강세(환율 하락)에 베팅'했을 때 중간에 최대 몇 % 상승했는지(양수 또는 0)
        "strong_won_bet_adverse_pct": round((path_vals[max_pos] / entry_val - 1) * 100, 3),
        "strong_won_bet_adverse_day": int((path_dates[max_pos] - entry_date).days),
    }


def build_analogs_and_mae(krw, pct5y_series, spread_daily):
    last_date = krw.index[-1]
    last_pct = float(pct5y_series.iloc[-1])
    last_spread_sign = None
    if spread_daily is not None and last_date in spread_daily.index:
        v = spread_daily.loc[last_date]
        last_spread_sign = None if pd.isna(v) else (1 if v > 0 else (-1 if v < 0 else 0))

    dates_arr = krw.index.values
    values_arr = krw.values

    # 유사 국면: 5년 백분위가 현재 ±10%p 이내 + 한미 금리차 부호가 같은 과거 시점(월 1회 샘플링, 최근 6개월 이내 시점은 제외 — 순방향 수익률 계산이 완결되지 않아서)
    candidate_dates = []
    seen_months = set()
    for i, d in enumerate(krw.index):
        if d > last_date - pd.DateOffset(months=6):
            continue
        mk = (d.year, d.month)
        if mk in seen_months:
            continue
        pct = pct5y_series.iloc[i]
        if pd.isna(pct):
            continue
        if abs(pct - last_pct) > 10:
            continue
        if spread_daily is not None and d in spread_daily.index:
            sv = spread_daily.loc[d]
            sign = None if pd.isna(sv) else (1 if sv > 0 else (-1 if sv < 0 else 0))
            if last_spread_sign is None or sign != last_spread_sign:
                continue
        else:
            continue
        seen_months.add(mk)
        candidate_dates.append(d)

    horizons = {}
    for months in HORIZON_MONTHS:
        rets, maes = [], []
        for d in candidate_dates:
            entry_val = float(krw.loc[d])
            target = d + pd.DateOffset(months=months)
            end_d, end_v = nearest_value(dates_arr, values_arr, target)
            if end_v is None:
                continue
            rets.append((end_v / entry_val - 1) * 100)
            mae = mae_bidirectional(krw, d, end_d)
            if mae:
                maes.append(mae)
        entry = {"n": len(rets)}
        if rets:
            arr = np.array(rets)
            entry.update({
                "median_pct": round(float(np.median(arr)), 3),
                "p25_pct": round(float(np.percentile(arr, 25)), 3),
                "p75_pct": round(float(np.percentile(arr, 75)), 3),
            })
        if maes:
            weak = np.array([m["weak_won_bet_adverse_pct"] for m in maes])
            strong = np.array([m["strong_won_bet_adverse_pct"] for m in maes])
            weak_day = np.array([m["weak_won_bet_adverse_day"] for m in maes])
            strong_day = np.array([m["strong_won_bet_adverse_day"] for m in maes])
            entry["mae"] = {
                "weak_won_bet_median_pct": round(float(np.median(weak)), 3),
                "weak_won_bet_worst_pct": round(float(np.min(weak)), 3),
                "weak_won_bet_worst_day": int(weak_day[int(np.argmin(weak))]),
                "strong_won_bet_median_pct": round(float(np.median(strong)), 3),
                "strong_won_bet_worst_pct": round(float(np.max(strong)), 3),
                "strong_won_bet_worst_day": int(strong_day[int(np.argmax(strong))]),
            }
        horizons[f"{months}m"] = entry

    # 전체 기간 기준 MAE 분포(유사 국면으로 필터링하지 않은 모든 3개월·6개월 보유)
    baseline = {}
    for months in (3, 6):
        weak_list, strong_list = [], []
        seen = set()
        for i, d in enumerate(krw.index):
            mk = (d.year, d.month)
            if mk in seen:
                continue
            if d > last_date - pd.DateOffset(months=months):
                continue
            target = d + pd.DateOffset(months=months)
            end_d, end_v = nearest_value(dates_arr, values_arr, target)
            if end_v is None:
                continue
            seen.add(mk)
            mae = mae_bidirectional(krw, d, end_d)
            if mae:
                weak_list.append(mae["weak_won_bet_adverse_pct"])
                strong_list.append(mae["strong_won_bet_adverse_pct"])
        if weak_list:
            baseline[f"{months}m"] = {
                "n": len(weak_list),
                "weak_won_bet_median_pct": round(float(np.median(weak_list)), 3),
                "weak_won_bet_p10_pct": round(float(np.percentile(weak_list, 10)), 3),
                "weak_won_bet_worst_pct": round(float(np.min(weak_list)), 3),
                "strong_won_bet_median_pct": round(float(np.median(strong_list)), 3),
                "strong_won_bet_p90_pct": round(float(np.percentile(strong_list, 90)), 3),
                "strong_won_bet_worst_pct": round(float(np.max(strong_list)), 3),
            }
        else:
            baseline[f"{months}m"] = None

    return {
        "current": {"percentile_5y": round(last_pct, 1), "spread_kr_us_sign": last_spread_sign},
        "n_candidates": len(candidate_dates),
        "candidate_dates": [d.date().isoformat() for d in candidate_dates],
        "horizons": horizons,
        "baseline_mae_all_periods": baseline,
    }


# ---------------------------------------------------------------------------
# (6) 단순 규칙 백테스트
# ---------------------------------------------------------------------------


def monthly_sample_dates(krw, spread_daily=None):
    """매달 마지막 거래일을 표본점으로(현재 진행 중인 달은 제외)."""
    last_date = krw.index[-1]
    current_month = pd.Timestamp(date.today().strftime("%Y-%m") + "-01")
    s = krw.copy()
    s.index = pd.DatetimeIndex(s.index)
    monthly = s.groupby(s.index.to_period("M")).apply(lambda x: x.index[-1])
    out = [d for d in monthly.values if pd.Timestamp(d).to_period("M").to_timestamp() < current_month]
    return [pd.Timestamp(d) for d in out]


def evaluate_rule(krw, signal_dates_with_bet, horizon_months=3):
    """signal_dates_with_bet: [(date, bet)], bet in {'상승','하락'} — '상승'=원화
    약세(환율 상승)에 베팅, '하락'=원화 강세(환율 하락)에 베팅."""
    dates_arr = krw.index.values
    values_arr = krw.values
    hits, rets, maes = [], [], []
    for d, bet in signal_dates_with_bet:
        entry_val = float(krw.loc[d])
        target = d + pd.DateOffset(months=horizon_months)
        end_d, end_v = nearest_value(dates_arr, values_arr, target)
        if end_v is None:
            continue
        actual_ret = (end_v / entry_val - 1) * 100
        hit = (actual_ret > 0 and bet == "상승") or (actual_ret < 0 and bet == "하락")
        hits.append(hit)
        signed_ret = actual_ret if bet == "상승" else -actual_ret  # 베팅 방향 기준 수익률
        rets.append(signed_ret)
        mae = mae_bidirectional(krw, d, end_d)
        if mae:
            maes.append(mae["weak_won_bet_adverse_pct"] if bet == "상승" else -mae["strong_won_bet_adverse_pct"])
            # 베팅 방향 기준 역행폭(항상 0 이하 값 = 얼마나 반대로 움직였는지)
    n = len(hits)
    if n == 0:
        return {"n": 0}
    hit_rate = sum(hits) / n * 100
    binom = scipy_stats.binomtest(sum(hits), n, 0.5, alternative="two-sided")
    result = {
        "n": n,
        "hit_rate_pct": round(hit_rate, 1),
        "mean_directional_return_pct": round(float(np.mean(rets)), 3),
        "binom_p_value": round(float(binom.pvalue), 4),
        "low_sample_caveat": n < 30,
    }
    if maes:
        result["mean_adverse_excursion_pct"] = round(float(np.mean(maes)), 3)
        result["worst_adverse_excursion_pct"] = round(float(np.min(maes)), 3)
    return result


def buy_and_hold_baseline(krw, horizon_months=3):
    dates_arr = krw.index.values
    values_arr = krw.values
    rets = []
    for d in monthly_sample_dates(krw):
        entry_val = float(krw.loc[d])
        target = d + pd.DateOffset(months=horizon_months)
        _, end_v = nearest_value(dates_arr, values_arr, target)
        if end_v is None:
            continue
        rets.append((end_v / entry_val - 1) * 100)
    return round(float(np.mean(rets)), 3) if rets else None


def build_backtests(krw, pct5y_series, spread_daily):
    sample_dates = monthly_sample_dates(krw)
    pct_at = pct5y_series.reindex(krw.index).ffill()

    # a) 5년 백분위 80 이상 -> 하락(원화강세) 베팅 / 20 이하 -> 상승(원화약세) 베팅
    rule_a_signals = []
    for d in sample_dates:
        if d not in pct_at.index:
            continue
        p = pct_at.loc[d]
        if pd.isna(p):
            continue
        if p >= 80:
            rule_a_signals.append((d, "하락"))
        elif p <= 20:
            rule_a_signals.append((d, "상승"))
    rule_a = evaluate_rule(krw, rule_a_signals, horizon_months=3)
    rule_a["rule"] = "5년 백분위 80%↑ → 하락(원화강세) 베팅 / 20%↓ → 상승(원화약세) 베팅, 3개월 뒤 평가"
    rule_a["baseline_buy_and_hold_pct"] = buy_and_hold_baseline(krw, 3)

    # b) 120일 이동평균 위 -> 상승 추세 추종(상승 베팅) / 아래 -> 하락 추세 추종(하락 베팅)
    ma120 = krw.rolling(120).mean()
    rule_b_signals = []
    for d in sample_dates:
        if d not in ma120.index or pd.isna(ma120.loc[d]):
            continue
        bet = "상승" if krw.loc[d] > ma120.loc[d] else "하락"
        rule_b_signals.append((d, bet))
    rule_b = evaluate_rule(krw, rule_b_signals, horizon_months=3)
    rule_b["rule"] = "120일 이동평균 위 → 상승(원화약세) 추세추종 / 아래 → 하락(원화강세) 추세추종, 3개월 뒤 평가"
    rule_b["baseline_buy_and_hold_pct"] = buy_and_hold_baseline(krw, 3)

    # c) 한미 금리차 확대(전월 대비 상승) -> 상승 베팅 / 축소 -> 하락 베팅
    rule_c_signals = []
    if spread_daily is not None:
        spread_at = spread_daily.reindex(krw.index).ffill()
        for i, d in enumerate(sample_dates):
            if i == 0 or d not in spread_at.index:
                continue
            prev_d = sample_dates[i - 1]
            if prev_d not in spread_at.index:
                continue
            cur_s, prev_s = spread_at.loc[d], spread_at.loc[prev_d]
            if pd.isna(cur_s) or pd.isna(prev_s) or cur_s == prev_s:
                continue
            bet = "상승" if cur_s > prev_s else "하락"
            rule_c_signals.append((d, bet))
    rule_c = evaluate_rule(krw, rule_c_signals, horizon_months=3)
    rule_c["rule"] = "한미 정책금리차 확대(전월 대비) → 상승(원화약세) 베팅 / 축소 → 하락(원화강세) 베팅, 3개월 뒤 평가"
    rule_c["baseline_buy_and_hold_pct"] = buy_and_hold_baseline(krw, 3)

    return {"rule_a_percentile": rule_a, "rule_b_ma120": rule_b, "rule_c_spread": rule_c}


# ---------------------------------------------------------------------------
# 이벤트 캘린더(향후 30일)
# ---------------------------------------------------------------------------


def build_upcoming_events(days_ahead=30):
    """발표일정 시트에서 향후 days_ahead일 내 정책금리·CPI·고용 발표를 뽑는다.
    로우데이터 엑셀을 직접 읽는다(export.py가 만드는 calendar.json은 이 스크립트
    실행 시점에 아직 최신이 아닐 수 있어 원본에서 바로 읽음). export.py의
    export_calendar()와 같은 위치 기반 열 순서(일자·국가·지표·대상기간·출처·비고)를 쓴다."""
    import openpyxl

    xlsx_path = os.path.join(BASE_DIR, "매크로_트래커_로우데이터.xlsx")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if "발표일정" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["발표일정"]
    today = date.today()
    end = today + timedelta(days=days_ahead)
    target_indicators = {"기준금리", "정책금리(상단)", "정책금리", "CPI 헤드라인", "CPI 근원", "비농업고용", "실업률"}
    events = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        scheduled_date, country, indicator = (list(row) + [None] * 3)[:3]
        d = scheduled_date
        if hasattr(d, "date"):
            d = d.date()
        if not isinstance(d, date) or not (today <= d <= end):
            continue
        if indicator not in target_indicators:
            continue
        events.append({
            "date": d.isoformat(),
            "country": country,
            "indicator": indicator,
            "d_day": (d - today).days,
        })
    wb.close()
    events.sort(key=lambda e: e["date"])
    return events


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main():
    fx_rows = load_fx_daily_rows()
    series = build_daily_series(fx_rows)
    krw, dxy, jpy = series["KRW"], series["DXY"], series["JPY"]

    raw_rows = rl.load_rawdata()
    policy_monthly = rl.build_policy_monthly(raw_rows)  # index=월말, columns=KR/US/JP
    spread_monthly = (policy_monthly["KR"] - policy_monthly["US"]).dropna()
    oil_monthly = build_dubai_oil_monthly(raw_rows)

    # 정책금리차를 일별로 순방향 채움(월별 값이 그 달 전체에 적용됐다고 가정)
    spread_daily = spread_monthly.copy()
    spread_daily.index = spread_daily.index.to_period("M").to_timestamp()  # 월말 -> 월초로 이동(재인덱싱 편의)
    full_idx = pd.date_range(spread_daily.index.min(), krw.index.max(), freq="D")
    spread_daily = spread_daily.reindex(full_idx).ffill()
    spread_daily = spread_daily.reindex(krw.index)

    pct5y_series = expanding_5y_percentile_series(krw)

    position = compute_position_indicators(krw, pct5y_series)
    swings = detect_swings(krw)
    dollar_decomp = build_dollar_decomposition(krw, dxy)

    monthly_df = build_monthly_frame(krw, jpy, oil_monthly, spread_monthly)
    drivers = build_driver_correlations(monthly_df)

    volatility = build_volatility(krw)
    analogs = build_analogs_and_mae(krw, pct5y_series, spread_daily)
    backtests = build_backtests(krw, pct5y_series, spread_daily)
    events = build_upcoming_events()

    result = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "as_of": krw.index[-1].date().isoformat(),
        "sample": {"start": krw.index[0].date().isoformat(), "end": krw.index[-1].date().isoformat(), "n_days": int(len(krw))},
        "position": position,
        "swings": swings,
        "krw_daily": {"dates": [d.date().isoformat() for d in krw.index], "value": [round(float(v), 2) for v in krw.values]},
        "dollar_decomposition": dollar_decomp,
        "drivers": drivers,
        "volatility": volatility,
        "analogs": analogs,
        "backtests": backtests,
        "upcoming_events": events,
        "methodology_note": (
            "이 데이터는 원/달러 환율의 과거 대비 위치와 과거 통계를 보여줄 뿐 향후 방향을 예측하지 않습니다. "
            "단기 환율은 펀더멘털 모형이 단순 예측을 이기지 못한다는 연구 결과가 일반적이며, 고점·저점은 사후에만 확정됩니다. "
            "표본은 제한적(2021-09~현재)이고 독립적인 국면은 이보다 훨씬 적습니다. "
            "과거 최대 역행폭은 미래 손실의 상한이 아닙니다. 투자 판단과 그 결과는 본인 책임입니다."
        ),
    }

    os.makedirs(os.path.dirname(FX_POSITION_JSON_PATH), exist_ok=True)
    with open(FX_POSITION_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"fx_position.py: {FX_POSITION_JSON_PATH} 생성")
    print(f"  현재 5년 백분위: {position['percentile_5y']}%, z-score: {position['z_score_5y']}")
    print(f"  스윙 고점·저점: {len(swings)}개(확정 {sum(1 for p in swings if p['confirmed'])}개)")
    print(f"  유사 국면 후보: {analogs['n_candidates']}개")
    for name, r in backtests.items():
        print(f"  백테스트 {name}: n={r.get('n')}, 적중률={r.get('hit_rate_pct')}%, p={r.get('binom_p_value')}")


if __name__ == "__main__":
    main()

"""'환율 위치 분석' 안의 '선행지표 예측력 검정'.

목적은 특정 지표가 원/달러의 미래 방향을 예측하는지 데이터로 확인하는 것이다.
결과가 나쁘게 나와도(예측력이 없다는 결론이어도) 그대로 낸다 — 매수·매도
신호나 투자 권유는 만들지 않는다.

후보 변수(data/fx_predictors_raw.csv, fx_collect.py가 수집 + 로우데이터·
forecasts.csv에서 가져오는 것들): 두바이유, 브렌트유, 구리, 위안화, VIX,
한국 수출액(전년동월비), 경상수지, 미국 2년물 국채금리, 한미 정책금리차.

방법론:
- 동시점 상관: 원/달러 월간 로그변화율 vs 각 변수 월간 변화(가격류는 로그
  변화율, 금리·스프레드는 %p 차분, 이미 변화율인 수출액YoY는 그대로).
- 선행 회귀: 변수의 t월 변화 -> 원/달러의 t+1/t+2/t+3월 변화(HAC+OLS 둘 다).
- 그레인저 인과성(1~3개월 시차, statsmodels).
- 표본 외(out-of-sample) 검정이 핵심: 36개월 롤링 창으로 학습해 다음 1개월을
  예측 -> 방향 적중률(이항검정 p값) + RMSE를 '변화 없음(랜덤워크, 항상 0
  예측)'과 비교한 비율(1 미만이어야 의미 있음). look-ahead bias 방지를 위해
  각 시점의 학습 데이터는 그 시점 이전 값만 쓴다(아래 build_walk_forward_oos
  참고 — 예측에 쓰는 마지막 X는 t월 값, 목표는 t+1월 KRW 변화이며 t+1월의
  어떤 정보도 학습에 들어가지 않는다).
- 변수별 단순 규칙 백테스트: 그 변수의 월간 변화가 자기 분포 상위/하위
  1/3(tercile) 안에 들면 다음 달 원화약세/강세에 베팅(고정 %는 변수마다
  스케일이 달라 부적절해 분위 기준을 씀 — 설계 판단, 출력에도 명시).
- 다중회귀 표본외 검정: 개별 검정에서 '예측력 확인' 또는 '경계'로 나온
  변수만 모아 다중회귀(변수가 하나도 없으면 동시점 상관 절대값 상위 3개로
  대체) — 표본이 36개월뿐이라 변수를 다 넣으면 과적합 위험이 커서 하는
  선택, 이 로직도 출력에 기록한다.

collect.py 흐름에서 analysis/fx_position.py 다음에 자동 실행되며,
site/data/fx_predictors.json으로 결과를 낸다.
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
from scipy import stats as scipy_stats
from statsmodels.tsa.stattools import grangercausalitytests

warnings.simplefilter("ignore")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fx_position as fpos  # noqa: E402  (nearest_value, mae_bidirectional, evaluate_rule 등 재사용)
import rate_linkage as rl  # noqa: E402  (load_rawdata, build_policy_monthly, simple_regression 재사용)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DAILY_CSV_PATH = os.path.join(BASE_DIR, "data", "fx_daily.csv")
PREDICTOR_CSV_PATH = os.path.join(BASE_DIR, "data", "fx_predictors_raw.csv")
FX_PREDICTORS_JSON_PATH = os.path.join(BASE_DIR, "site", "data", "fx_predictors.json")

OBS_START = "2021-09-01"
OOS_TRAIN_WINDOW = 36  # 개월
HORIZONS = [1, 2, 3]
GRANGER_MAXLAG = 3

# 변수별 변화 계산 방식: "logret"(가격류, 로그 변화율) | "diff"(금리·스프레드,
# %p 차분) | "level"(이미 변화율/비율인 계열, 그대로 씀)
VARIABLE_TRANSFORM = {
    "두바이유": "logret",
    "브렌트유": "logret",
    "구리": "logret",
    "위안화": "logret",
    "VIX": "logret",
    "한국수출액YoY": "level",
    "경상수지": "diff",
    "미국국채2년": "diff",
    "한미정책금리차": "diff",
}
VARIABLE_LABEL = {
    "두바이유": "두바이유",
    "브렌트유": "브렌트유",
    "구리": "구리",
    "위안화": "위안화(CNY/USD)",
    "VIX": "VIX",
    "한국수출액YoY": "한국 수출액(전년동월비)",
    "경상수지": "경상수지",
    "미국국채2년": "미국 2년물 국채금리",
    "한미정책금리차": "한미 정책금리차",
}


# ---------------------------------------------------------------------------
# 데이터 로드 및 월별 정렬
# ---------------------------------------------------------------------------


def load_krw_monthly():
    """원/달러 월평균 수준값(index=월초 Timestamp), 완결된 달만."""
    rows = fpos.load_fx_daily_rows()
    series = fpos.build_daily_series(rows)
    krw = series["KRW"]
    current_month = pd.Timestamp(date.today().strftime("%Y-%m") + "-01")
    s = krw.copy()
    s.index = s.index.to_period("M")
    monthly = s.groupby(level=0).mean()
    monthly.index = monthly.index.to_timestamp()
    monthly = monthly[monthly.index < current_month]
    return monthly, krw  # (월평균, 일별원본 — 일별은 백테스트 entry/exit 가격 조회용)


def load_predictor_raw_rows():
    if not os.path.exists(PREDICTOR_CSV_PATH):
        raise SystemExit(f"오류: {PREDICTOR_CSV_PATH} 파일이 없습니다. fx_collect.py를 먼저 실행하세요.")
    with open(PREDICTOR_CSV_PATH, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def monthly_mean_series(rows, name):
    """{날짜, 변수명, 값} 형태 행들 중 name에 해당하는 값을 월평균으로(일간
    변수) 또는 그대로(이미 월간인 변수, 한 달에 1개 값이라 평균해도 동일)
    합쳐 pandas Series(index=월초 Timestamp)로 만든다."""
    by_month = {}
    for r in rows:
        if r["변수명"] != name:
            continue
        mk = r["날짜"][:7]
        by_month.setdefault(mk, []).append(float(r["값"]))
    current_month = date.today().strftime("%Y-%m")
    data = {}
    for mk, vals in by_month.items():
        if mk >= current_month:
            continue
        data[pd.Timestamp(mk + "-01")] = sum(vals) / len(vals)
    return pd.Series(data).sort_index()


def load_dubai_oil_monthly(raw_rows):
    vals = {}
    current_month = date.today().strftime("%Y-%m")
    for r in raw_rows:
        if r["country"] == "한국" and r["indicator"] == "두바이유" and r["value"] is not None:
            mk = r["date"].isoformat()[:7]
            if mk >= current_month:
                continue
            vals[pd.Timestamp(mk + "-01")] = float(r["value"])
    return pd.Series(vals).sort_index()


def load_dgs2_monthly(predictor_rows):
    """미국 2년물 국채금리 월평균. forecast_collect.py도 DGS2를 모으지만
    최근 760일(~25개월)만 유지해(시장금리-기준금리 스프레드 차트용) 표본외
    검정에 필요한 전체 히스토리가 부족하므로, fx_collect.py가 별도로 전체
    기간 수집해 둔 data/fx_predictors_raw.csv에서 읽는다."""
    return monthly_mean_series(predictor_rows, "미국국채2년")


def build_export_yoy_monthly(rows):
    level = monthly_mean_series(rows, "한국수출액")
    if len(level) < 13:
        return pd.Series(dtype=float)
    yoy = {}
    for i in range(12, len(level)):
        d = level.index[i]
        prior_d = d - pd.DateOffset(years=1)
        if prior_d not in level.index:
            continue
        yoy[d] = (level.loc[d] / level.loc[prior_d] - 1) * 100
    return pd.Series(yoy).sort_index()


def build_variable_frame():
    """모든 후보 변수를 월초 인덱스로 정렬한 DataFrame(수준값, 아직 변환 전)."""
    predictor_rows = load_predictor_raw_rows()
    raw_rows = rl.load_rawdata()
    policy_monthly = rl.build_policy_monthly(raw_rows)
    spread = (policy_monthly["KR"] - policy_monthly["US"]).dropna()
    spread.index = spread.index.to_period("M").to_timestamp()

    data = {
        "두바이유": load_dubai_oil_monthly(raw_rows),
        "브렌트유": monthly_mean_series(predictor_rows, "브렌트유"),
        "구리": monthly_mean_series(predictor_rows, "구리"),
        "위안화": monthly_mean_series(predictor_rows, "위안화"),
        "VIX": monthly_mean_series(predictor_rows, "VIX"),
        "한국수출액YoY": build_export_yoy_monthly(predictor_rows),
        "경상수지": monthly_mean_series(predictor_rows, "경상수지"),
        "미국국채2년": load_dgs2_monthly(predictor_rows),
        "한미정책금리차": spread,
    }
    return pd.DataFrame(data).sort_index()


def transform_variable(level_series, kind):
    if kind == "logret":
        return np.log(level_series).diff()
    if kind == "diff":
        return level_series.diff()
    if kind == "level":
        return level_series
    raise ValueError(f"알 수 없는 변환: {kind}")


# ---------------------------------------------------------------------------
# 동시점 상관 + 선행 회귀 + 그레인저
# ---------------------------------------------------------------------------


def contemporaneous_corr(y, x):
    pair = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    if len(pair) < 8:
        return {"corr": None, "n": len(pair)}
    return {"corr": round(float(pair["y"].corr(pair["x"])), 3), "n": int(len(pair))}


def leading_regressions(y, x):
    """x의 t월 변화 -> y의 t+h월 변화(h=1,2,3). y를 -h만큼 당겨서(shift(-h))
    x와 같은 t 인덱스에서 비교 — x[t]가 실제로 y[t+h]보다 시간상 앞선다."""
    out = {}
    for h in HORIZONS:
        y_shifted = y.shift(-h)
        r = rl.simple_regression(y_shifted, x)
        out[f"{h}m"] = r
    return out


def granger_tests(y, x):
    """statsmodels grangercausalitytests: x가 y를 그레인저 인과하는지.
    입력 배열은 [y, x] 순서(왼쪽 열이 그레인저 인과되는 대상)."""
    pair = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    if len(pair) < GRANGER_MAXLAG + 10:
        return {f"lag{k}_p": None for k in range(1, GRANGER_MAXLAG + 1)} | {"n": len(pair)}
    try:
        res = grangercausalitytests(pair[["y", "x"]].values, GRANGER_MAXLAG, verbose=False)
    except Exception:  # noqa: BLE001
        return {f"lag{k}_p": None for k in range(1, GRANGER_MAXLAG + 1)} | {"n": len(pair)}
    out = {"n": int(len(pair))}
    for lag in range(1, GRANGER_MAXLAG + 1):
        try:
            out[f"lag{lag}_p"] = round(float(res[lag][0]["ssr_ftest"][1]), 4)
        except Exception:  # noqa: BLE001
            out[f"lag{lag}_p"] = None
    return out


# ---------------------------------------------------------------------------
# 표본 외(out-of-sample) 검정 — 핵심
# ---------------------------------------------------------------------------


def build_walk_forward_oos(y, x, window=OOS_TRAIN_WINDOW):
    """36개월 롤링 창으로 (x[t] -> y[t+1]) 관계를 학습해 다음 달을 예측한다.

    look-ahead 방지: 학습쌍은 (x[t], y[t+1]) for t = i-window .. i-1이며,
    예측 시점 i의 x[i]는 이미 알려진 값(그 달이 끝나야 관측되므로 실제로는
    다음 달 초에나 쓸 수 있다는 실무적 한계는 있지만, 최소한 y[i+1]이라는
    '미래 정답'은 학습·예측 어느 단계에도 들어가지 않는다). 예측 대상
    y[i+1]은 예측이 끝난 뒤에만 실제값과 비교한다.
    """
    pair = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    n = len(pair)
    preds, actuals, dates = [], [], []
    for i in range(window, n - 1):
        train_x = pair["x"].iloc[i - window : i].values  # x[t], t=i-window..i-1
        train_y = pair["y"].iloc[i - window + 1 : i + 1].values  # y[t+1], 대응
        X = sm.add_constant(train_x)
        try:
            model = sm.OLS(train_y, X).fit()
        except Exception:  # noqa: BLE001
            continue
        x_pred_input = pair["x"].iloc[i]
        pred = float(model.params[0] + model.params[1] * x_pred_input)
        actual = float(pair["y"].iloc[i + 1])
        preds.append(pred)
        actuals.append(actual)
        dates.append(pair.index[i + 1].date().isoformat())

    if not preds:
        return None

    preds = np.array(preds)
    actuals = np.array(actuals)
    hits = (np.sign(preds) == np.sign(actuals)) & (actuals != 0)
    n_eval = len(preds)
    n_hits = int(hits.sum())
    hit_rate = n_hits / n_eval * 100
    binom = scipy_stats.binomtest(n_hits, n_eval, 0.5, alternative="two-sided")
    rmse_model = float(np.sqrt(np.mean((preds - actuals) ** 2)))
    rmse_naive = float(np.sqrt(np.mean(actuals ** 2)))  # '변화 없음(랜덤워크)' 예측 = 항상 0
    return {
        "n": n_eval,
        "hit_rate_pct": round(hit_rate, 1),
        "binom_p_value": round(float(binom.pvalue), 4),
        "rmse_model": round(rmse_model, 5),
        "rmse_naive_random_walk": round(rmse_naive, 5),
        "rmse_ratio": round(rmse_model / rmse_naive, 4) if rmse_naive else None,
        "dates": dates,
        "predicted": [round(float(v), 5) for v in preds],
        "actual": [round(float(v), 5) for v in actuals],
    }


def judge_predictive_power(oos):
    """판정 규칙: 표본외 적중률이 50%를 유의하게(p<0.05) 웃돌고 RMSE 비율이
    1 미만이면 '예측력 확인', 둘 중 하나만이면 '경계', 둘 다 아니면
    '예측력 확인 안 됨'."""
    if not oos:
        return "표본 부족"
    hit_significant = oos["hit_rate_pct"] > 50 and oos["binom_p_value"] < 0.05
    rmse_better = oos["rmse_ratio"] is not None and oos["rmse_ratio"] < 1
    if hit_significant and rmse_better:
        return "예측력 확인"
    if hit_significant or rmse_better:
        return "경계"
    return "예측력 확인 안 됨"


# ---------------------------------------------------------------------------
# 변수별 단순 규칙 백테스트(분위 기준)
# ---------------------------------------------------------------------------


def build_variable_rule_backtest(krw_daily, change_series):
    """change_series(월 변화, index=월초)의 상위/하위 1/3(tercile) 기준으로
    '다음 달' 원화약세/강세에 베팅하는 단순 규칙. 고정 %는 변수마다 스케일이
    달라(예: VIX vs 수출액YoY) 부적절해 분위 기준을 쓴다."""
    change_series = change_series.dropna()
    if len(change_series) < 12:
        return {"n": 0}
    q_hi = change_series.quantile(2 / 3)
    q_lo = change_series.quantile(1 / 3)
    sample_dates = fpos.monthly_sample_dates(krw_daily)
    signal_dates = set(change_series.index)
    signals = []
    for d in sample_dates:
        # change_series는 월초 인덱스, sample_dates는 그 달 마지막 거래일 -> 같은 달 매칭
        mk = pd.Timestamp(d.year, d.month, 1)
        if mk not in signal_dates:
            continue
        v = change_series.loc[mk]
        if v >= q_hi:
            signals.append((d, "상승"))
        elif v <= q_lo:
            signals.append((d, "하락"))
    result = fpos.evaluate_rule(krw_daily, signals, horizon_months=1)
    result["threshold_high"] = round(float(q_hi), 4)
    result["threshold_low"] = round(float(q_lo), 4)
    if result.get("n"):
        result["baseline_buy_and_hold_pct"] = fpos.buy_and_hold_baseline(krw_daily, 1)
    return result


# ---------------------------------------------------------------------------
# 롤링 12개월 상관계수(화면 차트용, fx_position.py의 drivers와 같은 방식)
# ---------------------------------------------------------------------------


def rolling_corr_series(y, x, window=12):
    pair = pd.concat([y, x], axis=1, keys=["y", "x"]).dropna()
    roll = pair["y"].rolling(window).corr(pair["x"]).dropna()
    return {"dates": [d.date().isoformat() for d in roll.index], "corr": [round(float(v), 3) for v in roll.values]}


# ---------------------------------------------------------------------------
# 다중회귀 표본외 검정
# ---------------------------------------------------------------------------


def build_multi_walk_forward_oos(y, x_frame, window=OOS_TRAIN_WINDOW):
    df = pd.concat([y] + [x_frame[c] for c in x_frame.columns], axis=1, keys=["y"] + list(x_frame.columns)).dropna()
    n = len(df)
    cols = list(x_frame.columns)
    if n < window + 5 or not cols:
        return None
    preds, actuals, dates = [], [], []
    for i in range(window, n - 1):
        train_X = df[cols].iloc[i - window : i].values
        train_y = df["y"].iloc[i - window + 1 : i + 1].values
        X = sm.add_constant(train_X)
        try:
            model = sm.OLS(train_y, X).fit()
        except Exception:  # noqa: BLE001
            continue
        x_pred_input = np.concatenate([[1.0], df[cols].iloc[i].values])
        pred = float(np.dot(model.params, x_pred_input))
        actual = float(df["y"].iloc[i + 1])
        preds.append(pred)
        actuals.append(actual)
        dates.append(df.index[i + 1].date().isoformat())
    if not preds:
        return None
    preds = np.array(preds)
    actuals = np.array(actuals)
    hits = (np.sign(preds) == np.sign(actuals)) & (actuals != 0)
    n_eval = len(preds)
    n_hits = int(hits.sum())
    hit_rate = n_hits / n_eval * 100
    binom = scipy_stats.binomtest(n_hits, n_eval, 0.5, alternative="two-sided")
    rmse_model = float(np.sqrt(np.mean((preds - actuals) ** 2)))
    rmse_naive = float(np.sqrt(np.mean(actuals ** 2)))
    return {
        "variables": cols,
        "n": n_eval,
        "hit_rate_pct": round(hit_rate, 1),
        "binom_p_value": round(float(binom.pvalue), 4),
        "rmse_model": round(rmse_model, 5),
        "rmse_naive_random_walk": round(rmse_naive, 5),
        "rmse_ratio": round(rmse_model / rmse_naive, 4) if rmse_naive else None,
    }


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main():
    krw_monthly, krw_daily = load_krw_monthly()
    y = np.log(krw_monthly).diff().dropna()  # 원/달러 월간 로그변화율

    level_frame = build_variable_frame()

    results = {}
    change_frame = {}
    for name, kind in VARIABLE_TRANSFORM.items():
        level = level_frame[name].dropna()
        change = transform_variable(level, kind)
        change_frame[name] = change

        entry = {
            "label": VARIABLE_LABEL[name],
            "transform": kind,
            "contemporaneous": contemporaneous_corr(y, change),
            "leading_regression": leading_regressions(y, change),
            "granger": granger_tests(y, change),
        }
        oos = build_walk_forward_oos(y, change)
        entry["out_of_sample"] = oos
        entry["judgment"] = judge_predictive_power(oos)
        entry["rolling_corr_12m"] = rolling_corr_series(y, change)
        entry["rule_backtest"] = build_variable_rule_backtest(krw_daily, change)
        results[name] = entry

    # 다중회귀 후보 선정: '예측력 확인' 또는 '경계'인 변수만, 없으면 동시점
    # 상관 절대값 상위 3개로 대체(표본 36개월에 변수를 다 넣으면 과적합 위험).
    qualified = [name for name, r in results.items() if r["judgment"] in ("예측력 확인", "경계")]
    selection_method = "개별 검정에서 '예측력 확인' 또는 '경계'로 나온 변수"
    if not qualified:
        ranked = sorted(
            results.items(),
            key=lambda kv: abs(kv[1]["contemporaneous"]["corr"] or 0),
            reverse=True,
        )
        qualified = [name for name, _ in ranked[:3]]
        selection_method = "위 기준을 만족하는 변수가 없어 동시점 상관 절대값 상위 3개로 대체"

    x_frame_for_multi = pd.DataFrame({name: change_frame[name] for name in qualified})
    multi_oos = build_multi_walk_forward_oos(y, x_frame_for_multi)

    result = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "sample": {
            "start": OBS_START,
            "end": krw_monthly.index.max().date().isoformat() if len(krw_monthly) else None,
            "n_months_krw": int(len(krw_monthly)),
        },
        "oos_train_window_months": OOS_TRAIN_WINDOW,
        "variables": results,
        "multi_variable": {
            "selected_variables": qualified,
            "selection_method": selection_method,
            "out_of_sample": multi_oos,
        },
        "candidate_variables_not_collected": {
            "발틱운임지수(BDI)": "발틱거래소(Baltic Exchange) 유료 구독 데이터라 자동 수집 대상에서 제외.",
            "상하이컨테이너운임지수(SCFI)": "상하이해운거래소 공식 페이지가 로그인 후에만 수치를 보여주는 유료성 페이지라 자동 수집 불가로 확인. 무료·구조화된 대체는 찾지 못함(참고: 뉴욕연은 Global Supply Chain Pressure Index는 BDI 등을 반영하는 합성지수이며 무료 xlsx로 공개돼 있으나, SCFI 자체는 아니며 이번 작업 범위에서 후보 변수로 추가하지는 않음).",
        },
        "look_ahead_bias_note": (
            "표본외 검정에서 t월 x값으로 t+1월 y값을 예측할 때, 학습 구간(과거 36개월)의 각 쌍도 "
            "반드시 (x[t], y[t+1]) 형태로만 구성했다(x[t+1]이나 그 뒤 값을 학습에 절대 포함하지 않음). "
            "예측 시점의 계수는 그 시점 이전 36개월 데이터로만 추정했고, 비교 대상 실제값(y[t+1])은 "
            "예측이 끝난 뒤에만 채점에 사용했다. 코드상 이는 build_walk_forward_oos()에서 학습 인덱스를 "
            "i-window..i-1(x) / i-window+1..i(y, 한 칸 밀림)로, 예측 입력은 x[i] 하나로 한정해 보장한다."
        ),
        "methodology_note": (
            "동시점 상관이 높아도 예측력이 있다는 뜻은 아닙니다. 표본외 방향 적중률과 RMSE 비율이 "
            "실제 예측력의 기준이며, 표본이 제한적(2021-09~현재)이라 우연히 좋게 나올 수 있습니다."
        ),
    }

    os.makedirs(os.path.dirname(FX_PREDICTORS_JSON_PATH), exist_ok=True)
    with open(FX_PREDICTORS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"fx_predictors.py: {FX_PREDICTORS_JSON_PATH} 생성")
    for name, r in results.items():
        oos = r["out_of_sample"]
        print(f"  {VARIABLE_LABEL[name]}: 판정={r['judgment']}" + (f", 표본외 n={oos['n']}, 적중률={oos['hit_rate_pct']}%, RMSE비율={oos['rmse_ratio']}" if oos else ", 표본 부족"))
    print(f"  다중회귀 선정 변수({selection_method}): {qualified}")
    if multi_oos:
        print(f"  다중회귀 표본외: n={multi_oos['n']}, 적중률={multi_oos['hit_rate_pct']}%, RMSE비율={multi_oos['rmse_ratio']}")


if __name__ == "__main__":
    main()

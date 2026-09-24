"""'환율' 탭을 위한 일별 환율 수집 + '선행지표 예측력 검정'용 후보 변수 수집.

ECOS(원/달러·원/100엔·원/유로 매매기준율)와 FRED(달러지수 광의·엔/달러·
달러/유로)를 일별로 받아 data/fx_daily.csv(긴 형식: 날짜|통화쌍|값|출처)에
누적하고, 월평균·월말값을 계산해 site/data/fx.json으로 내보낸다.

로우데이터 엑셀(매크로_트래커_로우데이터.xlsx)은 전혀 건드리지 않는다 —
환율은 별도 CSV로만 관리한다(요청대로).

증분 수집: data/fx_daily.csv에 이미 있는 마지막 날짜 다음날부터만 새로 받는다.
반복값 방지: 최신 값이 그 통화쌍의 직전 저장값과 완전히 같으면(주말/휴일에
데이터 소스가 이전 값을 그대로 반복하는 경우가 있어) 경고만 남기고 그대로
저장한다(ECOS/FRED 모두 vintage 정보가 없어 보류 대신 기록만 함 — 기존
Eurostat/ECOS 정책값과 동일한 원칙).

추가로 '선행지표 예측력 검정'(analysis/fx_predictors.py)이 쓰는 후보 변수
(브렌트유·구리·위안화·VIX·한국 수출액·경상수지)도 같은 증분·반복값 원칙으로
수집해 data/fx_predictors_raw.csv(긴 형식: 날짜|변수명|값|출처)에 저장한다.
통화쌍이 아니라 원자재·환율·거시 변수가 섞여 있어 fx_daily.csv와는 다른
파일에 둔다. 일간 변수(브렌트유·위안화·VIX)와 월간 변수(구리·수출액·
경상수지)가 섞여 있으며, 날짜는 각 변수의 원래 주기 그대로 저장한다
(월간 변수는 그 달 1일로 저장 — analysis/fx_predictors.py가 월별로만 쓴다).

발틱운임지수(BDI)는 발틱거래소(Baltic Exchange) 유료 구독 데이터라 수집
대상에서 제외했다. 무료 대체재로 상하이컨테이너운임지수(SCFI, 상하이해운
거래소)를 확인했으나 로그인 후에만 수치가 보이는 유료성 페이지라 마찬가지로
제외했다(작업보고서 참고).

collect.py 흐름 끝에서 자동 실행된다.
"""

import csv
import json
import os
import sys
import warnings
from datetime import date, datetime, timedelta

warnings.simplefilter("ignore")
from dotenv import load_dotenv  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import collect  # noqa: E402  (FredClient/EcosClient 재사용)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DAILY_CSV_PATH = os.path.join(DATA_DIR, "fx_daily.csv")
JSON_PATH = os.path.join(BASE_DIR, "site", "data", "fx.json")

DAILY_HEADER = ["날짜", "통화쌍", "값", "출처"]
OBS_START = date(2021, 9, 1)

ECOS_STAT_CODE = "731Y001"
ECOS_SERIES = {
    "원/달러": ("0000001", "https://ecos.bok.or.kr"),
    "원/100엔": ("0000002", "https://ecos.bok.or.kr"),
    "원/유로": ("0000003", "https://ecos.bok.or.kr"),
}
FRED_SERIES = {
    "달러지수(광의)": ("DTWEXBGS", "https://fred.stlouisfed.org/series/DTWEXBGS"),
    "엔/달러": ("DEXJPUS", "https://fred.stlouisfed.org/series/DEXJPUS"),
    "달러/유로": ("DEXUSEU", "https://fred.stlouisfed.org/series/DEXUSEU"),
}
ALL_PAIRS = list(ECOS_SERIES.keys()) + list(FRED_SERIES.keys())

# ---- 선행지표 예측력 검정용 후보 변수 ----
PREDICTOR_CSV_PATH = os.path.join(DATA_DIR, "fx_predictors_raw.csv")
PREDICTOR_HEADER = ["날짜", "변수명", "값", "출처"]
PREDICTOR_OBS_START = date(2021, 9, 1)

FRED_PREDICTOR_DAILY = {
    "브렌트유": ("DCOILBRENTEU", "https://fred.stlouisfed.org/series/DCOILBRENTEU"),
    "위안화": ("DEXCHUS", "https://fred.stlouisfed.org/series/DEXCHUS"),
    "VIX": ("VIXCLS", "https://fred.stlouisfed.org/series/VIXCLS"),
    # forecast_collect.py도 DGS2를 모으지만 최근 760일(~25개월)만 유지하도록
    # 설계돼 있어(시장금리-기준금리 스프레드 차트용) 표본외 검정에 필요한
    # 36개월 이상의 전체 히스토리를 확보하려고 여기서 별도로 전체 기간 수집한다.
    "미국국채2년": ("DGS2", "https://fred.stlouisfed.org/series/DGS2"),
}
FRED_PREDICTOR_MONTHLY = {
    "구리": ("PCOPPUSDM", "https://fred.stlouisfed.org/series/PCOPPUSDM"),
}
ECOS_BOP_STAT_CODE = "301Y013"
ECOS_PREDICTOR_MONTHLY = {
    "한국수출액": ("110000", "https://ecos.bok.or.kr"),
    "경상수지": ("000000", "https://ecos.bok.or.kr"),
}
ALL_PREDICTORS = list(FRED_PREDICTOR_DAILY.keys()) + list(FRED_PREDICTOR_MONTHLY.keys()) + list(ECOS_PREDICTOR_MONTHLY.keys())

RESULTS = []


def log_result(source, status, note=""):
    RESULTS.append((source, status, note))
    print(f"  [{status}] {source}" + (f" — {note}" if note else ""))


def read_daily_csv():
    if not os.path.exists(DAILY_CSV_PATH):
        return []
    with open(DAILY_CSV_PATH, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_daily_csv(rows):
    rows_sorted = sorted(rows, key=lambda r: (r["통화쌍"], r["날짜"]))
    with open(DAILY_CSV_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=DAILY_HEADER)
        writer.writeheader()
        for r in rows_sorted:
            writer.writerow(r)


def last_date_for_pair(rows, pair):
    dates = [r["날짜"] for r in rows if r["통화쌍"] == pair]
    return max(dates) if dates else None


def next_month_start(d):
    """d가 속한 달의 다음 달 1일. 월간 시리즈는 항상 그 달 1일로 오므로
    (timedelta로 대충 30여 일을 더하면 FRED observation_start가 그 달 1일보다
    뒤로 밀려 해당 달 관측치가 통째로 누락될 수 있어) 정확히 1일 단위로 옮긴다."""
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def read_predictor_csv():
    if not os.path.exists(PREDICTOR_CSV_PATH):
        return []
    with open(PREDICTOR_CSV_PATH, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_predictor_csv(rows):
    rows_sorted = sorted(rows, key=lambda r: (r["변수명"], r["날짜"]))
    with open(PREDICTOR_CSV_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=PREDICTOR_HEADER)
        writer.writeheader()
        for r in rows_sorted:
            writer.writerow(r)


def last_date_for_var(rows, name):
    dates = [r["날짜"] for r in rows if r["변수명"] == name]
    return max(dates) if dates else None


def collect_ecos_pairs(ecos_key, existing_rows):
    added = []
    if not ecos_key:
        log_result("ECOS 환율", "건너뜀", "ECOS_API_KEY 없음")
        return added
    client = collect.EcosClient(ecos_key)
    today = date.today()
    for pair, (item_code, source_url) in ECOS_SERIES.items():
        try:
            last = last_date_for_pair(existing_rows, pair)
            start = (datetime.strptime(last, "%Y-%m-%d").date() + timedelta(days=1)) if last else OBS_START
            if start > today:
                log_result(f"ECOS - {pair}", "성공", "0건(이미 최신)")
                continue
            values = client.observations(ECOS_STAT_CODE, "D", start, today, [item_code])
            prev_val = None
            if last:
                prev_rows = [r for r in existing_rows if r["통화쌍"] == pair and r["날짜"] == last]
                if prev_rows:
                    prev_val = float(prev_rows[0]["값"])
            n = 0
            for d in sorted(values.keys()):
                v = values[d]
                if prev_val is not None and v == prev_val:
                    log_result(f"ECOS - {pair}", "반복값(참고)", f"{d.isoformat()}: {v} (직전과 동일, 그대로 저장)")
                added.append({"날짜": d.isoformat(), "통화쌍": pair, "값": str(round(v, 4)), "출처": source_url})
                prev_val = v
                n += 1
            log_result(f"ECOS - {pair}", "성공", f"{n}건")
        except Exception as e:  # noqa: BLE001
            log_result(f"ECOS - {pair}", "실패", str(e))
    return added


def collect_fred_pairs(fred_key, existing_rows):
    added = []
    if not fred_key:
        log_result("FRED 환율", "건너뜀", "FRED_API_KEY 없음")
        return added
    client = collect.FredClient(fred_key)
    today = date.today()
    for pair, (series_id, source_url) in FRED_SERIES.items():
        try:
            last = last_date_for_pair(existing_rows, pair)
            start = (datetime.strptime(last, "%Y-%m-%d").date() + timedelta(days=1)) if last else OBS_START
            if start > today:
                log_result(f"FRED - {pair}", "성공", "0건(이미 최신)")
                continue
            values = client.latest_observations(series_id, start, today)
            prev_val = None
            if last:
                prev_rows = [r for r in existing_rows if r["통화쌍"] == pair and r["날짜"] == last]
                if prev_rows:
                    prev_val = float(prev_rows[0]["값"])
            n = 0
            for d in sorted(values.keys()):
                v = float(values[d])
                if prev_val is not None and v == prev_val:
                    log_result(f"FRED - {pair}", "반복값(참고)", f"{d}: {v} (직전과 동일, 그대로 저장)")
                added.append({"날짜": d, "통화쌍": pair, "값": str(round(v, 4)), "출처": source_url})
                prev_val = v
                n += 1
            log_result(f"FRED - {pair}", "성공", f"{n}건")
        except Exception as e:  # noqa: BLE001
            log_result(f"FRED - {pair}", "실패", str(e))
    return added


def collect_fred_predictor_daily(fred_key, existing_rows):
    added = []
    if not fred_key:
        log_result("FRED 후보변수(일간)", "건너뜀", "FRED_API_KEY 없음")
        return added
    client = collect.FredClient(fred_key)
    today = date.today()
    for name, (series_id, source_url) in FRED_PREDICTOR_DAILY.items():
        try:
            last = last_date_for_var(existing_rows, name)
            start = (datetime.strptime(last, "%Y-%m-%d").date() + timedelta(days=1)) if last else PREDICTOR_OBS_START
            if start > today:
                log_result(f"FRED - {name}", "성공", "0건(이미 최신)")
                continue
            values = client.latest_observations(series_id, start, today)
            prev_val = None
            if last:
                prev_rows = [r for r in existing_rows if r["변수명"] == name and r["날짜"] == last]
                if prev_rows:
                    prev_val = float(prev_rows[0]["값"])
            n = 0
            for d in sorted(values.keys()):
                v = float(values[d])
                if prev_val is not None and v == prev_val:
                    log_result(f"FRED - {name}", "반복값(참고)", f"{d}: {v} (직전과 동일, 그대로 저장)")
                added.append({"날짜": d, "변수명": name, "값": str(round(v, 4)), "출처": source_url})
                prev_val = v
                n += 1
            log_result(f"FRED - {name}", "성공", f"{n}건")
        except Exception as e:  # noqa: BLE001
            log_result(f"FRED - {name}", "실패", str(e))
    return added


def collect_fred_predictor_monthly(fred_key, existing_rows):
    added = []
    if not fred_key:
        log_result("FRED 후보변수(월간)", "건너뜀", "FRED_API_KEY 없음")
        return added
    client = collect.FredClient(fred_key)
    today = date.today()
    for name, (series_id, source_url) in FRED_PREDICTOR_MONTHLY.items():
        try:
            last = last_date_for_var(existing_rows, name)
            start = next_month_start(datetime.strptime(last, "%Y-%m-%d").date()) if last else PREDICTOR_OBS_START
            if start > today:
                log_result(f"FRED - {name}", "성공", "0건(이미 최신)")
                continue
            values = client.latest_observations(series_id, start, today)  # FRED가 월초 날짜로 월간 값을 내줌
            prev_val = None
            if last:
                prev_rows = [r for r in existing_rows if r["변수명"] == name and r["날짜"] == last]
                if prev_rows:
                    prev_val = float(prev_rows[0]["값"])
            n = 0
            for d in sorted(values.keys()):
                v = float(values[d])
                if prev_val is not None and v == prev_val:
                    log_result(f"FRED - {name}", "반복값(참고)", f"{d}: {v} (직전과 동일, 그대로 저장)")
                added.append({"날짜": d, "변수명": name, "값": str(round(v, 4)), "출처": source_url})
                prev_val = v
                n += 1
            log_result(f"FRED - {name}", "성공", f"{n}건")
        except Exception as e:  # noqa: BLE001
            log_result(f"FRED - {name}", "실패", str(e))
    return added


def collect_ecos_predictor_monthly(ecos_key, existing_rows):
    added = []
    if not ecos_key:
        log_result("ECOS 후보변수", "건너뜀", "ECOS_API_KEY 없음")
        return added
    client = collect.EcosClient(ecos_key)
    today = date.today()
    for name, (item_code, source_url) in ECOS_PREDICTOR_MONTHLY.items():
        try:
            last = last_date_for_var(existing_rows, name)
            start = next_month_start(datetime.strptime(last, "%Y-%m-%d").date()) if last else PREDICTOR_OBS_START
            if start > today:
                log_result(f"ECOS - {name}", "성공", "0건(이미 최신)")
                continue
            values = client.observations(ECOS_BOP_STAT_CODE, "M", start, today, [item_code])
            prev_val = None
            if last:
                prev_rows = [r for r in existing_rows if r["변수명"] == name and r["날짜"] == last]
                if prev_rows:
                    prev_val = float(prev_rows[0]["값"])
            n = 0
            for d in sorted(values.keys()):
                v = values[d]
                if prev_val is not None and v == prev_val:
                    log_result(f"ECOS - {name}", "반복값(참고)", f"{d.isoformat()}: {v} (직전과 동일, 그대로 저장)")
                added.append({"날짜": d.isoformat(), "변수명": name, "값": str(round(v, 4)), "출처": source_url})
                prev_val = v
                n += 1
            log_result(f"ECOS - {name}", "성공", f"{n}건")
        except Exception as e:  # noqa: BLE001
            log_result(f"ECOS - {name}", "실패", str(e))
    return added


def month_key(d_str):
    return d_str[:7]


def build_monthly(rows):
    """{통화쌍: {"dates": [...], "avg": [...], "eom": [...]}} — 완결된 달만 포함."""
    current_month = date.today().strftime("%Y-%m")
    out = {}
    for pair in ALL_PAIRS:
        by_month = {}
        pair_rows = sorted([r for r in rows if r["통화쌍"] == pair], key=lambda r: r["날짜"])
        for r in pair_rows:
            mk = month_key(r["날짜"])
            by_month.setdefault(mk, []).append((r["날짜"], float(r["값"])))
        dates, avgs, eoms = [], [], []
        for mk in sorted(by_month.keys()):
            if mk >= current_month:
                continue
            vals = by_month[mk]
            dates.append(mk + "-01")
            avgs.append(round(sum(v for _, v in vals) / len(vals), 4))
            eoms.append(round(vals[-1][1], 4))
        out[pair] = {"dates": dates, "avg": avgs, "eom": eoms}
    return out


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(os.path.dirname(JSON_PATH), exist_ok=True)
    load_dotenv(dotenv_path=os.path.join(BASE_DIR, ".env"))
    fred_key = os.environ.get("FRED_API_KEY")
    ecos_key = os.environ.get("ECOS_API_KEY")

    print(f"=== fx_collect.py 실행: {date.today().isoformat()} ===\n")

    existing_rows = read_daily_csv()
    print(f"기존 data/fx_daily.csv: {len(existing_rows)}행")

    new_ecos = collect_ecos_pairs(ecos_key, existing_rows)
    new_fred = collect_fred_pairs(fred_key, existing_rows)

    existing_keys = {(r["날짜"], r["통화쌍"]) for r in existing_rows}
    all_new = new_ecos + new_fred
    added = [r for r in all_new if (r["날짜"], r["통화쌍"]) not in existing_keys]
    existing_rows.extend(added)
    write_daily_csv(existing_rows)
    print(f"\n=== data/fx_daily.csv: 신규 {len(added)}행 추가, 총 {len(existing_rows)}행 ===")

    monthly = build_monthly(existing_rows)
    latest = {}
    daily = {}
    for pair in ALL_PAIRS:
        pair_rows = sorted([r for r in existing_rows if r["통화쌍"] == pair], key=lambda r: r["날짜"])
        if pair_rows:
            latest[pair] = {"date": pair_rows[-1]["날짜"], "value": float(pair_rows[-1]["값"])}
        # 프론트에서 1개월/6개월/1년/5년 구간 차트와 전일·1개월·1년 전 대비 계산에
        # 쓸 일별 전체 시계열(OBS_START=2021-09부터라 5년치 전부와 같음).
        daily[pair] = [{"date": r["날짜"], "value": float(r["값"])} for r in pair_rows]

    output = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "monthly": monthly,
        "daily": daily,
        "latest": latest,
        "collection_results": [{"출처": s, "상태": st, "비고": n} for s, st, n in RESULTS],
    }
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"=== {JSON_PATH} 생성 ===")

    print("\n-- 선행지표 예측력 검정용 후보 변수 --")
    existing_predictor_rows = read_predictor_csv()
    print(f"기존 data/fx_predictors_raw.csv: {len(existing_predictor_rows)}행")
    new_fred_daily = collect_fred_predictor_daily(fred_key, existing_predictor_rows)
    new_fred_monthly = collect_fred_predictor_monthly(fred_key, existing_predictor_rows)
    new_ecos_predictor = collect_ecos_predictor_monthly(ecos_key, existing_predictor_rows)
    existing_predictor_keys = {(r["날짜"], r["변수명"]) for r in existing_predictor_rows}
    all_new_predictor = new_fred_daily + new_fred_monthly + new_ecos_predictor
    added_predictor = [r for r in all_new_predictor if (r["날짜"], r["변수명"]) not in existing_predictor_keys]
    existing_predictor_rows.extend(added_predictor)
    write_predictor_csv(existing_predictor_rows)
    print(f"=== data/fx_predictors_raw.csv: 신규 {len(added_predictor)}행 추가, 총 {len(existing_predictor_rows)}행 ===")

    print("\n=== 수집 결과 요약 ===")
    for s, st, n in RESULTS:
        print(f"  [{st}] {s}" + (f" — {n}" if n else ""))


if __name__ == "__main__":
    main()

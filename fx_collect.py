"""'환율' 탭을 위한 일별 환율 수집.

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

    print("\n=== 수집 결과 요약 ===")
    for s, st, n in RESULTS:
        print(f"  [{st}] {s}" + (f" — {n}" if n else ""))


if __name__ == "__main__":
    main()

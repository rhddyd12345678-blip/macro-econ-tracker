"""매크로 트래커 엑셀 -> 홈페이지용 JSON 내보내기.

'매크로_트래커_로우데이터.xlsx'를 읽기 전용으로 열어 site/data/data.json(로우데이터
전체 + 지표목록의 단위·출처)과 site/data/calendar.json(발표일정)을 생성한다.
엑셀 파일은 전혀 수정하지 않는다(읽기만 함). collect.py 실행이 끝나면 자동으로
이어서 실행되지만, 단독으로도 실행할 수 있다: `python3 export.py`.
"""

import json
import os
import warnings
from datetime import date, datetime

warnings.simplefilter("ignore", UserWarning)
import openpyxl  # noqa: E402

XLSX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "매크로_트래커_로우데이터.xlsx")
SITE_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "site", "data")
DATA_JSON_PATH = os.path.join(SITE_DATA_DIR, "data.json")
CALENDAR_JSON_PATH = os.path.join(SITE_DATA_DIR, "calendar.json")


def _to_iso(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.date().isoformat() if isinstance(v, datetime) else v.isoformat()
    return v


def export_indicators(wb):
    ws = wb["지표목록"]
    indicators = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        country, category, indicator, freq, unit, source = (list(row) + [None] * 6)[:6]
        indicators.append(
            {
                "country": country,
                "category": category,
                "indicator": indicator,
                "freq": freq,
                "unit": unit,
                "source": source,
            }
        )
    return indicators


def export_observations(wb):
    ws = wb["로우데이터"]
    observations = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        base_date, country, indicator, value, release_date, freq = (list(row) + [None] * 6)[:6]
        observations.append(
            {
                "date": _to_iso(base_date),
                "country": country,
                "indicator": indicator,
                "value": value,
                "release_date": _to_iso(release_date),
                "freq": freq,
            }
        )
    return observations


def export_calendar(wb):
    if "발표일정" not in wb.sheetnames:
        return []
    ws = wb["발표일정"]
    events = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        scheduled_date, country, indicator, period, source, note = (list(row) + [None] * 6)[:6]
        events.append(
            {
                "date": _to_iso(scheduled_date),
                "country": country,
                "indicator": indicator,
                "period": period,
                "source": source,
                "note": note,
            }
        )
    events.sort(key=lambda e: e["date"] or "")
    return events


def main():
    if not os.path.exists(XLSX_PATH):
        raise SystemExit(f"오류: {XLSX_PATH} 파일을 찾을 수 없습니다.")

    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    indicators = export_indicators(wb)
    observations = export_observations(wb)
    calendar_events = export_calendar(wb)
    wb.close()

    os.makedirs(SITE_DATA_DIR, exist_ok=True)

    generated_at = datetime.now().isoformat(timespec="seconds")

    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"generated_at": generated_at, "indicators": indicators, "observations": observations},
            f,
            ensure_ascii=False,
            indent=2,
        )

    with open(CALENDAR_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump({"generated_at": generated_at, "events": calendar_events}, f, ensure_ascii=False, indent=2)

    print(f"export.py: {DATA_JSON_PATH} 생성 (지표 {len(indicators)}개, 관측치 {len(observations)}개)")
    print(f"export.py: {CALENDAR_JSON_PATH} 생성 (발표일정 {len(calendar_events)}개)")


if __name__ == "__main__":
    main()

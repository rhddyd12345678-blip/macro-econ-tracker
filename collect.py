"""매크로 트래커 로우데이터 수집 스크립트.

'지표목록' 시트를 설정표로 읽어 FRED에서 데이터를 받아 '로우데이터' 시트에 추가한다.
'로우데이터' 시트는 이미 전체 범위(A1:F5000)에 드롭다운·서식이 적용된 템플릿이 채워져
있으므로, openpyxl로 저장하면 x14 확장 데이터 유효성 검사(드롭다운)가 깨진다. 이를
피하기 위해 읽기는 openpyxl(read_only)로, 쓰기는 xlsx(zip) 내부의 시트 XML을 직접
문자열 치환하는 방식으로 처리한다.
"""

import os
import re
import shutil
import sys
import tempfile
import time
import warnings
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from xml.sax.saxutils import escape

import requests
from dotenv import load_dotenv

warnings.simplefilter("ignore", UserWarning)
import openpyxl  # noqa: E402  (경고 억제 이후 임포트)

XLSX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "매크로_트래커_로우데이터.xlsx")
SHEET_INDICATORS = "지표목록"
SHEET_RAWDATA = "로우데이터"
SHEET_LIST = "_목록"
RAWDATA_SHEET_XML = "xl/worksheets/sheet2.xml"  # 워크북 내 로우데이터 시트의 실제 파일명

# 로우데이터 B/C/F열 드롭다운이 참조하는 _목록 열: 국가(A)/지표(C)/주기(B)
LIST_COL_FOR_DROPDOWN = {"B": "A", "C": "C", "F": "B"}

FRED_BASE = "https://api.stlouisfed.org/fred/series/observations"
OBS_START = date(2021, 9, 1)

# 1단계 수집 대상 국가. 이후 단계에서 다른 국가(e-Stat, ECOS, Eurostat 등)를 추가할 때 확장.
TARGET_COUNTRIES = {"미국"}

EXCEL_EPOCH = date(1899, 12, 30)


def to_excel_serial(d):
    return (d - EXCEL_EPOCH).days


def month_start(d):
    return date(d.year, d.month, 1)


def quarter_start(d):
    q_month = ((d.month - 1) // 3) * 3 + 1
    return date(d.year, q_month, 1)


# ---------------------------------------------------------------------------
# 지표목록 설정 로딩
# ---------------------------------------------------------------------------

def load_indicator_config(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[SHEET_INDICATORS]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()

    indicators = []
    for row in rows:
        if not row or row[0] is None:
            continue
        country, category, indicator, freq, unit, source, series_field, note = (
            list(row) + [None] * (8 - len(row))
        )[:8]
        indicators.append(
            {
                "country": country,
                "category": category,
                "indicator": indicator,
                "freq": freq,
                "unit": unit,
                "source": source,
                "series_field": series_field,
                "note": note,
            }
        )
    return indicators


FRED_FIELD_RE = re.compile(r"FRED:\s*([A-Za-z0-9_]+)")


def parse_fred_directive(series_field):
    if not series_field:
        return None
    m = FRED_FIELD_RE.search(series_field)
    if not m:
        return None
    return {
        "series_id": m.group(1),
        "is_pc1": "units=pc1" in series_field,
        "daily_to_monthly": "일간" in series_field and "월평균" in series_field,
    }


# ---------------------------------------------------------------------------
# FRED API
# ---------------------------------------------------------------------------

class FredClient:
    def __init__(self, api_key):
        self.api_key = api_key
        self.session = requests.Session()

    def _get(self, params):
        params = dict(params, api_key=self.api_key, file_type="json")
        resp = self.session.get(FRED_BASE, params=params, timeout=30)
        time.sleep(0.1)
        return resp

    def latest_observations(self, series_id, start, end, units="lin"):
        """최신(수정 반영) 값. {date: value_str} 반환, 결측('.')은 제외."""
        resp = self._get(
            {
                "series_id": series_id,
                "observation_start": start.isoformat(),
                "observation_end": end.isoformat(),
                "units": units,
            }
        )
        resp.raise_for_status()
        obs = resp.json().get("observations", [])
        return {o["date"]: o["value"] for o in obs if o["value"] != "."}

    def first_vintage_dates(self, series_id, dates):
        """주어진 관측일들(ISO 문자열 집합)에 대해 ALFRED 첫 vintage(최초 공개) 날짜를 구한다.

        전체 기간을 한 번에 조회(output_type=3)하되, vintage 개수 제한(2000개)에
        걸리면 관측일 하나씩 좁은 실시간 구간으로 재조회한다. 반환값도 ISO 문자열.
        """
        if not dates:
            return {}
        start, end = min(dates), max(dates)
        result = self._first_vintage_wide(series_id, start, end)
        if result is not None:
            return {d: v for d, v in result.items() if d in dates}

        # 넓은 구간 조회 실패 -> 관측일별 개별 조회로 폴백
        result = {}
        for d in sorted(dates):
            v = self._first_vintage_single(series_id, d)
            if v:
                result[d] = v
        return result

    def _first_vintage_wide(self, series_id, start, end):
        resp = self._get(
            {
                "series_id": series_id,
                "observation_start": start,
                "observation_end": end,
                "output_type": 3,
                "realtime_start": "1776-07-04",
                "realtime_end": "9999-12-31",
            }
        )
        if resp.status_code != 200:
            return None
        obs = resp.json().get("observations", [])
        out = {}
        for o in obs:
            vintages = [k for k in o.keys() if k != "date"]
            if not vintages:
                continue
            first_key = min(vintages, key=lambda k: k.rsplit("_", 1)[-1])
            out[o["date"]] = first_key.rsplit("_", 1)[-1]  # YYYYMMDD
        return {d: _yyyymmdd_to_iso(v) for d, v in out.items()}

    def _first_vintage_single(self, series_id, obs_date_str, window_days=60):
        obs_date = datetime.strptime(obs_date_str, "%Y-%m-%d").date()
        rt_start = obs_date - timedelta(days=window_days)
        rt_end = obs_date + timedelta(days=window_days)
        resp = self._get(
            {
                "series_id": series_id,
                "observation_start": obs_date_str,
                "observation_end": obs_date_str,
                "output_type": 3,
                "realtime_start": rt_start.isoformat(),
                "realtime_end": rt_end.isoformat(),
            }
        )
        if resp.status_code != 200:
            return None
        obs = resp.json().get("observations", [])
        if not obs:
            return None
        vintages = [k for k in obs[0].keys() if k != "date"]
        if not vintages:
            return None
        first_key = min(vintages, key=lambda k: k.rsplit("_", 1)[-1])
        return _yyyymmdd_to_iso(first_key.rsplit("_", 1)[-1])


def _yyyymmdd_to_iso(s):
    return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"


# ---------------------------------------------------------------------------
# 지표별 수집 로직
# ---------------------------------------------------------------------------

def collect_indicator(client, cfg, today):
    """반환: (rows, error_message_or_None)

    rows: [{date, value, release_date_or_None}]
    """
    directive = parse_fred_directive(cfg["series_field"])
    if directive is None:
        return None, "FRED 시리즈 매핑 없음 (출처: %r)" % cfg["series_field"]

    series_id = directive["series_id"]
    freq = cfg["freq"]

    try:
        if freq == "수시":
            return _collect_adhoc(client, series_id, today), None
        if directive["daily_to_monthly"]:
            return _collect_daily_to_monthly(client, series_id, today), None
        units = "pc1" if directive["is_pc1"] else "lin"
        return _collect_direct(client, series_id, units, today), None
    except requests.HTTPError as e:
        return None, f"FRED API 오류: {e}"
    except Exception as e:  # noqa: BLE001
        return None, f"수집 실패: {e}"


def _collect_direct(client, series_id, units, today):
    values = client.latest_observations(series_id, OBS_START, today, units=units)
    if not values:
        return []
    release_map = client.first_vintage_dates(series_id, set(values.keys()))
    rows = []
    for d_str, v_str in sorted(values.items()):
        rows.append(
            {
                "date": datetime.strptime(d_str, "%Y-%m-%d").date(),
                "value": round(float(v_str), 4),
                "release_date": _parse_iso_or_none(release_map.get(d_str)),
            }
        )
    return rows


def _collect_daily_to_monthly(client, series_id, today):
    daily = client.latest_observations(series_id, OBS_START, today, units="lin")
    if not daily:
        return []
    current_month = month_start(today)
    buckets = defaultdict(list)
    for d_str, v_str in daily.items():
        d = datetime.strptime(d_str, "%Y-%m-%d").date()
        ms = month_start(d)
        if ms >= current_month:
            continue  # 진행 중인 달(불완전 평균)은 제외
        buckets[ms].append(float(v_str))

    rows = []
    for ms in sorted(buckets.keys()):
        avg = sum(buckets[ms]) / len(buckets[ms])
        rows.append({"date": ms, "value": round(avg, 2), "release_date": None})
    return rows


def _collect_adhoc(client, series_id, today):
    daily = client.latest_observations(series_id, OBS_START, today, units="lin")
    if not daily:
        return []
    ordered = sorted(daily.items())  # (date_str, value_str)

    segments = []
    prev_val = None
    for d_str, v_str in ordered:
        v = float(v_str)
        if prev_val is None or v != prev_val:
            segments.append((d_str, v))
        prev_val = v

    seg_dates = {d for d, _ in segments}
    release_map = client.first_vintage_dates(series_id, seg_dates)

    rows = []
    for d_str, v in segments:
        rows.append(
            {
                "date": datetime.strptime(d_str, "%Y-%m-%d").date(),
                "value": round(v, 2),
                "release_date": _parse_iso_or_none(release_map.get(d_str)),
            }
        )
    return rows


def _parse_iso_or_none(s):
    if not s:
        return None
    return datetime.strptime(s, "%Y-%m-%d").date()


# ---------------------------------------------------------------------------
# 로우데이터 시트 읽기 (dedup, 다음 행 위치) / 쓰기 (XML 직접 치환)
# ---------------------------------------------------------------------------

def read_existing_rawdata(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[SHEET_RAWDATA]
    existing_keys = set()
    next_row = 2
    for row in ws.iter_rows(min_row=2, max_row=5000, values_only=True):
        base_date, country, indicator = row[0], row[1], row[2]
        if base_date is None:
            break
        if hasattr(base_date, "date"):
            base_date = base_date.date()
        existing_keys.add((base_date, country, indicator))
        next_row += 1
    wb.close()
    return existing_keys, next_row


def _build_row_xml(row_num, base_date, country, indicator, value, release_date, freq):
    def date_cell(col, style, d):
        if d is None:
            return f'<c r="{col}{row_num}" s="{style}"/>'
        return f'<c r="{col}{row_num}" s="{style}"><v>{to_excel_serial(d)}</v></c>'

    def text_cell(col, style, text):
        return f'<c r="{col}{row_num}" s="{style}" t="inlineStr"><is><t>{escape(str(text))}</t></is></c>'

    def num_cell(col, style, num):
        return f'<c r="{col}{row_num}" s="{style}"><v>{num}</v></c>'

    a = date_cell("A", 6, base_date)
    b = text_cell("B", 7, country)
    c = text_cell("C", 7, indicator)
    d = num_cell("D", 7, value)
    e = date_cell("E", 6, release_date)
    f = text_cell("F", 7, freq)
    return f'<row r="{row_num}" spans="1:6">{a}{b}{c}{d}{e}{f}</row>'


def _empty_row_template(row_num):
    return (
        f'<row r="{row_num}" spans="1:6">'
        f'<c r="A{row_num}" s="6"/><c r="B{row_num}" s="7"/><c r="C{row_num}" s="7"/>'
        f'<c r="D{row_num}" s="7"/><c r="E{row_num}" s="6"/><c r="F{row_num}" s="7"/>'
        f"</row>"
    )


def _replace_sheet_xml(xlsx_path, sheet_path, new_xml):
    """xlsx(zip) 안의 sheet_path 항목만 new_xml로 교체하고 나머지는 그대로 둔다."""
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    contents[sheet_path] = new_xml.encode("utf-8")

    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(xlsx_path))
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                zout.writestr(n, contents[n])
        shutil.move(tmp_path, xlsx_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def write_new_rows(xlsx_path, new_records, start_row):
    """new_records: [(base_date, country, indicator, value, release_date, freq)]"""
    if not new_records:
        return

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_xml = zin.read(RAWDATA_SHEET_XML).decode("utf-8")

    for i, rec in enumerate(new_records):
        row_num = start_row + i
        if row_num > 5000:
            raise RuntimeError("로우데이터 시트의 여유 행(5000행)을 초과했습니다.")
        old = _empty_row_template(row_num)
        if old not in sheet_xml:
            raise RuntimeError(f"{row_num}행이 예상한 빈 템플릿 형태가 아닙니다. 수동 확인이 필요합니다.")
        new = _build_row_xml(row_num, *rec)
        sheet_xml = sheet_xml.replace(old, new, 1)

    _replace_sheet_xml(xlsx_path, RAWDATA_SHEET_XML, sheet_xml)


_VALIDATION_BLOCK_RE = re.compile(r"<x14:dataValidation\b.*?</x14:dataValidation>", re.S)
_VALIDATION_F_RE = re.compile(r"(<xm:f>)(.*?)(</xm:f>)")
_VALIDATION_SQREF_RE = re.compile(r"<xm:sqref>([A-Z]+)\d+:[A-Z]+\d+</xm:sqref>")


def refresh_dropdown_ranges(xlsx_path):
    """로우데이터 B/C/F열 드롭다운(x14 데이터 유효성 검사)을 _목록 시트의 실제
    데이터 범위에 맞춰 다시 계산한다. _목록 시트 자체나 sqref(행 범위)는 건드리지
    않고, 각 드롭다운이 참조하는 '_목록!$X$1:$X$N' 범위의 N만 갱신한다."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    list_rows = list(wb[SHEET_LIST].iter_rows(values_only=True))
    wb.close()

    def col_count(col_letter):
        idx = ord(col_letter) - ord("A")
        n = 0
        for row in list_rows:
            if idx < len(row) and row[idx] is not None:
                n += 1
            else:
                break
        return n

    counts = {col: col_count(col) for col in set(LIST_COL_FOR_DROPDOWN.values())}
    for col, n in counts.items():
        if n < 1:
            raise RuntimeError(f"_목록 {col}열에서 값을 찾지 못했습니다.")

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_xml = zin.read(RAWDATA_SHEET_XML).decode("utf-8")

    changed = False

    def _replace_block(m):
        nonlocal changed
        block = m.group(0)
        sq = _VALIDATION_SQREF_RE.search(block)
        if not sq:
            return block
        dropdown_col = sq.group(1)
        list_col = LIST_COL_FOR_DROPDOWN.get(dropdown_col)
        if not list_col:
            return block
        new_range = f"{SHEET_LIST}!${list_col}$1:${list_col}${counts[list_col]}"

        def _sub_f(fm):
            nonlocal changed
            if fm.group(2) != new_range:
                changed = True
            return fm.group(1) + new_range + fm.group(3)

        return _VALIDATION_F_RE.sub(_sub_f, block, count=1)

    new_sheet_xml = _VALIDATION_BLOCK_RE.sub(_replace_block, sheet_xml)

    if not changed:
        return False

    _replace_sheet_xml(xlsx_path, RAWDATA_SHEET_XML, new_sheet_xml)
    return True


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
    api_key = os.environ.get("FRED_API_KEY")
    if not api_key:
        print("오류: .env 파일에 FRED_API_KEY가 설정되어 있지 않습니다.", file=sys.stderr)
        sys.exit(1)

    today = date.today()
    client = FredClient(api_key)

    all_indicators = load_indicator_config(XLSX_PATH)
    indicators = [i for i in all_indicators if i["country"] in TARGET_COUNTRIES]
    print(f"'{SHEET_INDICATORS}' 시트에서 {len(indicators)}개 지표(대상 국가: {', '.join(sorted(TARGET_COUNTRIES))})를 확인했습니다.\n")

    existing_keys, next_row = read_existing_rawdata(XLSX_PATH)
    print(f"기존 '{SHEET_RAWDATA}' 행: {len(existing_keys)}개, 다음 입력 행: {next_row}\n")

    to_write = []
    summary = []  # (indicator_label, rows_for_print)
    failed = []

    for cfg in indicators:
        label = f"{cfg['country']} / {cfg['indicator']}"
        rows, err = collect_indicator(client, cfg, today)
        if err:
            failed.append(f"{label}: {err}")
            continue
        if not rows:
            failed.append(f"{label}: 수집된 값 없음")
            continue

        added = 0
        for r in rows:
            key = (r["date"], cfg["country"], cfg["indicator"])
            if key in existing_keys:
                continue
            to_write.append(
                (r["date"], cfg["country"], cfg["indicator"], r["value"], r["release_date"], cfg["freq"])
            )
            existing_keys.add(key)
            added += 1

        summary.append((label, rows, added))

    if to_write:
        write_new_rows(XLSX_PATH, to_write, next_row)

    dropdown_updated = refresh_dropdown_ranges(XLSX_PATH)
    print(f"=== 수집 결과: {len(to_write)}개 행 추가 ===")
    print(f"=== B/C/F열 드롭다운(_목록 기준) {'갱신함' if dropdown_updated else '이미 최신 상태'} ===\n")

    for label, rows, added in summary:
        print(f"[{label}] 신규 {added}행 / 전체 {len(rows)}개 관측치 중 최근 3개:")
        for r in rows[-3:]:
            rel = r["release_date"].isoformat() if r["release_date"] else "(발표일 없음)"
            print(f"  {r['date'].isoformat()} | 값={r['value']} | 발표일={rel}")
        print()

    if failed:
        print("=== 수집하지 못한 지표 ===")
        for f in failed:
            print(f"  - {f}")
    else:
        print("모든 대상 지표를 수집했습니다.")


if __name__ == "__main__":
    main()

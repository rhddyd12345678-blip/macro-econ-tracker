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
INDICATORS_SHEET_XML = "xl/worksheets/sheet3.xml"  # 워크북 내 지표목록 시트의 실제 파일명

# 로우데이터 B/C/F열 드롭다운이 참조하는 _목록 열: 국가(A)/지표(C)/주기(B)
LIST_COL_FOR_DROPDOWN = {"B": "A", "C": "C", "F": "B"}

FRED_BASE = "https://api.stlouisfed.org/fred/series/observations"
ECOS_BASE = "https://ecos.bok.or.kr/api"
ESTAT_BASE = "https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData"
OBS_START = date(2021, 9, 1)

# 3단계 수집 대상 국가에 일본 추가. 이후 단계에서 다른 국가(Eurostat 등)를 추가할 때 확장.
TARGET_COUNTRIES = {"미국", "한국", "일본"}

# 지표목록 시트 한국/일본 행의 '시리즈 ID' 칸(G열)에 채워 넣을 확정 코드.
# 행 번호는 지표목록 시트의 실제 위치(1~3단계 조사로 확정: 한국 30~34·36행, 일본 11~14·16행).
KOREA_SERIES_UPDATES = {
    30: "ECOS: 901Y009/0 (전년동월비 계산)",  # CPI 헤드라인
    31: "ECOS: 901Y010/QB (전년동월비 계산)",  # CPI 근원 (농산물 및 석유류 제외지수)
    32: "ECOS: 901Y027/I61BA+I28A (원계열)",  # 취업자수
    33: "ECOS: 901Y027/I61BC+I28B (계절조정)",  # 실업률
    34: "ECOS: 200Y108/10601 (전기비 계산)",  # 실질GDP 성장률
    36: "ECOS: 722Y001/0101000 (변경일만 기록)",  # 기준금리
}
JAPAN_SERIES_UPDATES = {
    11: "ESTAT: 0004052037/tab3/cat0001/area00000 (전년동월비, e-Stat 제공값)",  # CPI 헤드라인
    12: "ESTAT: 0004052037/tab3/cat0161/area00000 (전년동월비, e-Stat 제공값)",  # CPI 근원
    13: "FRED: LRUNTTTTJPM156S",  # 실업률 (계절조정, OECD 경유)
    14: "FRED: JPNRGDPEXP (전기비 계산)",  # 실질GDP 성장률
    16: (
        "BOJ 공표자료 (수동 확정: 2021-09-01=-0.10%, 2024-03-19=0.10%, "
        "2024-07-31=0.25%, 2025-01-24=0.50%, 2025-12-19=0.75%, 2026-06-16=1.00%)"
    ),  # 정책금리
}

# 일본은행 정책금리 변경 이력(BOJ 공표문 원문 대조로 확정, 2021-09~현재). 기준일=발표일=결정일.
JAPAN_POLICY_RATE_CHANGES = [
    (date(2021, 9, 1), -0.10),
    (date(2024, 3, 19), 0.10),
    (date(2024, 7, 31), 0.25),
    (date(2025, 1, 24), 0.50),
    (date(2025, 12, 19), 0.75),
    (date(2026, 6, 16), 1.00),
]

EXCEL_EPOCH = date(1899, 12, 30)


def to_excel_serial(d):
    return (d - EXCEL_EPOCH).days


def month_start(d):
    return date(d.year, d.month, 1)


def quarter_start(d):
    q_month = ((d.month - 1) // 3) * 3 + 1
    return date(d.year, q_month, 1)


def add_months(d, n):
    total = d.year * 12 + (d.month - 1) + n
    return date(total // 12, total % 12 + 1, 1)


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
        "qoq_from_level": "전기비" in series_field and "계산" in series_field,
    }


ECOS_FIELD_RE = re.compile(r"ECOS:\s*([A-Za-z0-9]+)/([A-Za-z0-9+]+)")


def parse_ecos_directive(series_field):
    if not series_field:
        return None
    m = ECOS_FIELD_RE.search(series_field)
    if not m:
        return None
    return {"stat_code": m.group(1), "item_codes": m.group(2).split("+")}


ESTAT_FIELD_RE = re.compile(r"ESTAT:\s*(\w+)/tab(\w+)/cat(\w+)/area(\w+)")


def parse_estat_directive(series_field):
    if not series_field:
        return None
    m = ESTAT_FIELD_RE.search(series_field)
    if not m:
        return None
    return {"stats_data_id": m.group(1), "tab": m.group(2), "cat01": m.group(3), "area": m.group(4)}


BOJ_MANUAL_PREFIX = "BOJ 공표자료"


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
# ECOS API
# ---------------------------------------------------------------------------

def _ecos_date(d, cycle):
    if cycle == "D":
        return f"{d.year:04d}{d.month:02d}{d.day:02d}"
    if cycle == "M":
        return f"{d.year:04d}{d.month:02d}"
    if cycle == "Q":
        return f"{d.year:04d}Q{(d.month - 1) // 3 + 1}"
    raise ValueError(f"지원하지 않는 ECOS 주기: {cycle}")


def _ecos_parse_time(t, cycle):
    if cycle == "D":
        return date(int(t[0:4]), int(t[4:6]), int(t[6:8]))
    if cycle == "M":
        return date(int(t[0:4]), int(t[4:6]), 1)
    if cycle == "Q":
        y, q = int(t[0:4]), int(t[5])
        return date(y, (q - 1) * 3 + 1, 1)
    raise ValueError(f"지원하지 않는 ECOS 주기: {cycle}")


class EcosClient:
    def __init__(self, api_key):
        self.api_key = api_key
        self.session = requests.Session()

    def observations(self, stat_code, cycle, start, end, item_codes, count=6000):
        """{date: float} 반환. ECOS는 vintage/발표일 정보를 제공하지 않는다."""
        parts = "/".join(item_codes)
        url = (
            f"{ECOS_BASE}/StatisticSearch/{self.api_key}/json/kr/1/{count}/"
            f"{stat_code}/{cycle}/{_ecos_date(start, cycle)}/{_ecos_date(end, cycle)}/{parts}"
        )
        resp = self.session.get(url, timeout=30)
        time.sleep(0.1)
        resp.raise_for_status()
        payload = resp.json()
        if "StatisticSearch" not in payload:
            err = payload.get("RESULT", payload)
            raise RuntimeError(f"ECOS API 오류: {err}")
        rows = payload["StatisticSearch"].get("row", [])
        out = {}
        for row in rows:
            try:
                v = float(row["DATA_VALUE"])
            except (TypeError, ValueError):
                continue
            out[_ecos_parse_time(row["TIME"], cycle)] = v
        return out


# ---------------------------------------------------------------------------
# e-Stat API
# ---------------------------------------------------------------------------

def _estat_time_code(d):
    return f"{d.year:04d}00{d.month:02d}{d.month:02d}"


class EstatClient:
    def __init__(self, app_id):
        self.app_id = app_id
        self.session = requests.Session()

    def observations(self, stats_data_id, tab, cat01, area, start, end):
        """{date: float} 반환(월간). e-Stat은 vintage/발표일 정보를 제공하지 않는다."""
        params = {
            "appId": self.app_id,
            "statsDataId": stats_data_id,
            "cdTab": tab,
            "cdCat01": cat01,
            "cdArea": area,
            "cdTimeFrom": _estat_time_code(start),
            "cdTimeTo": _estat_time_code(end),
        }
        resp = self.session.get(ESTAT_BASE, params=params, timeout=30)
        time.sleep(0.1)
        resp.raise_for_status()
        payload = resp.json()
        stat_data = payload.get("GET_STATS_DATA", {}).get("STATISTICAL_DATA")
        if stat_data is None:
            err = payload.get("GET_STATS_DATA", {}).get("RESULT", payload)
            raise RuntimeError(f"e-Stat API 오류: {err}")
        values = stat_data.get("DATA_INF", {}).get("VALUE", [])
        if isinstance(values, dict):
            values = [values]
        out = {}
        for v in values:
            t = v["@time"]
            d = date(int(t[0:4]), int(t[6:8]), 1)
            try:
                out[d] = float(v["$"])
            except (TypeError, ValueError):
                continue
        return out


# ---------------------------------------------------------------------------
# 지표별 수집 로직
# ---------------------------------------------------------------------------

def collect_indicator(clients, cfg, today):
    """반환: (rows, error_message_or_None)

    rows: [{date, value, release_date_or_None}]
    clients: {"fred": FredClient, "ecos": EcosClient, "estat": EstatClient}
    """
    fred_directive = parse_fred_directive(cfg["series_field"])
    if fred_directive is not None:
        return _collect_fred(clients["fred"], fred_directive, cfg["freq"], today)

    ecos_directive = parse_ecos_directive(cfg["series_field"])
    if ecos_directive is not None:
        return _collect_ecos(clients["ecos"], ecos_directive, cfg, today)

    estat_directive = parse_estat_directive(cfg["series_field"])
    if estat_directive is not None:
        return _collect_estat(clients["estat"], estat_directive, today)

    if cfg["series_field"] and cfg["series_field"].startswith(BOJ_MANUAL_PREFIX):
        return _collect_boj_policy_rate(), None

    return None, "지원하지 않는 시리즈 출처 (시리즈 ID: %r)" % cfg["series_field"]


def _collect_fred(client, directive, freq, today):
    series_id = directive["series_id"]
    try:
        if freq == "수시":
            return _collect_adhoc(client, series_id, today), None
        if directive["daily_to_monthly"]:
            return _collect_daily_to_monthly(client, series_id, today), None
        if freq == "분기" and directive["qoq_from_level"]:
            return _collect_fred_qoq(client, series_id, today), None
        units = "pc1" if directive["is_pc1"] else "lin"
        return _collect_direct(client, series_id, units, today), None
    except requests.HTTPError as e:
        return None, f"FRED API 오류: {e}"
    except Exception as e:  # noqa: BLE001
        return None, f"수집 실패: {e}"


def _collect_ecos(client, directive, cfg, today):
    stat_code = directive["stat_code"]
    item_codes = directive["item_codes"]
    freq = cfg["freq"]
    category = cfg["category"]
    try:
        if freq == "수시":
            return _collect_ecos_adhoc(client, stat_code, item_codes, today), None
        if category == "물가":
            return _collect_ecos_yoy(client, stat_code, item_codes, today), None
        if category == "성장":
            return _collect_ecos_qoq(client, stat_code, item_codes, today), None
        return _collect_ecos_level(client, stat_code, item_codes, today), None
    except requests.HTTPError as e:
        return None, f"ECOS API 오류: {e}"
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


def _collect_fred_qoq(client, series_id, today):
    """분기 레벨 시리즈를 받아 전기비 %로 변환(한국 GDP와 동일 방식). 직접 계산한
    파생값이라 ALFRED vintage가 의미 없으므로 발표일은 항상 빈칸."""
    target_start = quarter_start(OBS_START)
    lookback_start = add_months(target_start, -3)
    values = client.latest_observations(series_id, lookback_start, today, units="lin")
    ordered = sorted(values.items())  # (date_str, value_str)
    rows = []
    prev_v = None
    for d_str, v_str in ordered:
        d = datetime.strptime(d_str, "%Y-%m-%d").date()
        v = float(v_str)
        if prev_v is not None and d >= target_start:
            qoq = (v / prev_v - 1) * 100
            rows.append({"date": d, "value": round(qoq, 2), "release_date": None})
        prev_v = v
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


def _collect_ecos_yoy(client, stat_code, item_codes, today):
    """월간 지수를 받아 전년동월비 %로 변환. ECOS는 발표일을 주지 않으므로 항상 빈칸."""
    lookback_start = date(OBS_START.year - 1, OBS_START.month, 1)
    values = client.observations(stat_code, "M", lookback_start, today, item_codes)
    rows = []
    for d in sorted(values):
        if d < OBS_START:
            continue
        prev = values.get(add_months(d, -12))
        if prev is None or prev == 0:
            continue
        yoy = (values[d] / prev - 1) * 100
        rows.append({"date": d, "value": round(yoy, 2), "release_date": None})
    return rows


def _collect_ecos_qoq(client, stat_code, item_codes, today):
    """분기 레벨을 받아 전기비 %로 변환. OBS_START가 속한 분기부터 포함(FRED 분기 시리즈와
    동일한 정렬 방식). ECOS는 발표일을 주지 않으므로 항상 빈칸."""
    target_start = quarter_start(OBS_START)
    lookback_start = add_months(target_start, -3)
    values = client.observations(stat_code, "Q", lookback_start, today, item_codes)
    ordered = sorted(values.items())
    rows = []
    prev_v = None
    for d, v in ordered:
        if prev_v is not None and d >= target_start:
            qoq = (v / prev_v - 1) * 100
            rows.append({"date": d, "value": round(qoq, 2), "release_date": None})
        prev_v = v
    return rows


def _collect_ecos_level(client, stat_code, item_codes, today):
    """레벨 값을 그대로 저장(취업자수/실업률 등). ECOS는 발표일을 주지 않으므로 항상 빈칸."""
    values = client.observations(stat_code, "M", OBS_START, today, item_codes)
    rows = []
    for d in sorted(values):
        if d < OBS_START:
            continue
        rows.append({"date": d, "value": round(values[d], 2), "release_date": None})
    return rows


def _collect_ecos_adhoc(client, stat_code, item_codes, today):
    """기준금리: 값이 바뀐 날짜만 기록. 기준일=발표일=금통위 결정일(변경 시작일)로 동일하게
    저장. 목표 기간 이전의 변경 이력도 함께 받아, 목표 기간 첫 날이 우연히 '변경일'로
    오인되지 않도록 한다."""
    lookback_start = date(OBS_START.year - 2, OBS_START.month, OBS_START.day)
    daily = client.observations(stat_code, "D", lookback_start, today, item_codes)
    ordered = sorted(daily.items())

    segments = []
    prev_val = None
    for d, v in ordered:
        if prev_val is None or v != prev_val:
            segments.append((d, v))
        prev_val = v

    rows = []
    for d, v in segments:
        if d < OBS_START:
            continue
        rows.append({"date": d, "value": round(v, 2), "release_date": d})
    return rows


def _collect_estat(client, directive, today):
    try:
        return _collect_estat_direct(
            client, directive["stats_data_id"], directive["tab"], directive["cat01"], directive["area"], today
        ), None
    except requests.HTTPError as e:
        return None, f"e-Stat API 오류: {e}"
    except Exception as e:  # noqa: BLE001
        return None, f"수집 실패: {e}"


def _collect_estat_direct(client, stats_data_id, tab, cat01, area, today):
    """e-Stat이 이미 계산해 주는 값(전년동월비 등)을 그대로 저장. e-Stat은 발표일을
    제공하지 않으므로 항상 빈칸."""
    values = client.observations(stats_data_id, tab, cat01, area, OBS_START, today)
    rows = []
    for d in sorted(values):
        if d < OBS_START:
            continue
        rows.append({"date": d, "value": round(values[d], 2), "release_date": None})
    return rows


def _collect_boj_policy_rate():
    """일본은행 정책금리: BOJ 공표문을 대조해 수동으로 확정한 변경 이력(JAPAN_POLICY_RATE_CHANGES).
    기준일=발표일=금융정책결정회의일."""
    rows = []
    for d, v in JAPAN_POLICY_RATE_CHANGES:
        if d < OBS_START:
            continue
        rows.append({"date": d, "value": v, "release_date": d})
    return rows


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


def update_indicator_series_ids(xlsx_path, updates):
    """지표목록 시트 특정 행들의 '시리즈 ID'(G열)에 확정된 코드를 적는다. updates에
    정의된 행만 건드리고, 이미 갱신된 행은 건너뛴다. 행마다 기존 placeholder 텍스트가
    달라도(공유 문자열/스타일 무관) 정규식으로 G열 셀 하나만 정확히 찾아 통째로 교체한다."""
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        sheet_xml = zin.read(INDICATORS_SHEET_XML).decode("utf-8")

    changed = False
    for row_num, new_text in updates.items():
        new_cell = f'<c r="G{row_num}" s="8" t="inlineStr"><is><t>{escape(new_text)}</t></is></c>'
        if new_cell in sheet_xml:
            continue  # 이미 갱신됨
        cell_re = re.compile(r'<c r="G%d"(?:[^>]*?/>|[^>]*>.*?</c>)' % row_num)
        new_sheet_xml, n = cell_re.subn(new_cell, sheet_xml, count=1)
        if n != 1:
            raise RuntimeError(f"지표목록 G{row_num} 셀을 찾지 못했습니다. 수동 확인이 필요합니다.")
        sheet_xml = new_sheet_xml
        changed = True

    if not changed:
        return False

    _replace_sheet_xml(xlsx_path, INDICATORS_SHEET_XML, sheet_xml)
    return True


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
    fred_key = os.environ.get("FRED_API_KEY")
    if not fred_key:
        print("오류: .env 파일에 FRED_API_KEY가 설정되어 있지 않습니다.", file=sys.stderr)
        sys.exit(1)
    ecos_key = os.environ.get("ECOS_API_KEY")
    if not ecos_key:
        print("오류: .env 파일에 ECOS_API_KEY가 설정되어 있지 않습니다.", file=sys.stderr)
        sys.exit(1)
    estat_key = os.environ.get("ESTAT_APP_ID")
    if not estat_key:
        print("오류: .env 파일에 ESTAT_APP_ID가 설정되어 있지 않습니다.", file=sys.stderr)
        sys.exit(1)

    today = date.today()
    clients = {"fred": FredClient(fred_key), "ecos": EcosClient(ecos_key), "estat": EstatClient(estat_key)}

    series_ids_updated = update_indicator_series_ids(XLSX_PATH, {**KOREA_SERIES_UPDATES, **JAPAN_SERIES_UPDATES})
    print(f"=== 지표목록 한국·일본 행 시리즈 ID {'갱신함' if series_ids_updated else '이미 최신 상태'} ===\n")

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
        rows, err = collect_indicator(clients, cfg, today)
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

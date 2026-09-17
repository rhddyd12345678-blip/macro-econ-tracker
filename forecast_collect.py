"""'시장 전망' 탭을 위한 전망·컨센서스·시장가격 데이터 수집.

data/forecasts.csv(긴 형식, 발표일=vintage 기준 누적 — 새 전망이 나와도 과거
전망 행을 절대 덮어쓰지 않는다)에 자동 수집 결과를 저장하고, 사람이 채워 넣는
data/forecasts_manual.csv를 함께 읽어 site/data/forecasts.json으로 내보낸다.

수집 원칙(모든 수집 함수가 지켜야 함):
1. 사이트가 403 등으로 자동 접근을 막으면 User-Agent 변경 같은 우회를 시도하지
   않는다 — '차단'으로 기록하고 넘어간다. 모든 HTTP 요청은 requests 기본
   설정(커스텀 User-Agent 없음)만 쓴다.
2. 저작권이 있는 원문(한국은행·일본은행 PDF, BMSI 보도자료 등)은 이 스크립트가
   직접 긁지 않는다 — 그런 출처는 사람이 data/forecasts_manual.csv에 수치만
   입력한다(docs/수동전망입력.md 참고).
3. 유료·재게시 금지 출처(Bloomberg, 연합인포맥스, CME FedWatch)는 아예 다루지
   않는다.
4. 한 출처가 실패해도 나머지 수집과 이 스크립트 자체, 그리고 이 스크립트를
   호출한 collect.py 파이프라인은 계속 진행되어야 한다 — 모든 수집 함수는
   예외를 삼키고 결과 상태(성공/실패/차단)를 반환한다.
5. 기관 전망(IMF/OECD/Fed SEP/ECB projections/필라델피아 연준 SPF/ECB SPF)은
   매주 월요일에만 조회한다(요청량 절약). 시장금리(ECOS/FRED)는 매번 조회한다.

실행: `python3 forecast_collect.py` (collect.py 흐름 끝에서 자동 호출됨).
"""

import csv
import json
import os
import re
import sys
import time
import warnings
from datetime import date, datetime, timedelta

import requests

warnings.simplefilter("ignore")
import openpyxl  # noqa: E402
import pandas as pd  # noqa: E402
from dotenv import load_dotenv  # noqa: E402

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
CSV_PATH = os.path.join(DATA_DIR, "forecasts.csv")
MANUAL_CSV_PATH = os.path.join(DATA_DIR, "forecasts_manual.csv")
JSON_PATH = os.path.join(BASE_DIR, "site", "data", "forecasts.json")

CSV_HEADER = ["발표일", "기관", "종류", "국가", "지표", "대상기간", "전망치", "출처URL"]
MANUAL_HEADER = ["발표일", "기관", "국가", "지표", "대상기간", "전망치", "응답비율_동결", "응답비율_인상", "응답비율_인하", "출처URL", "비고"]

MANUAL_INSTITUTION_KIND = {
    "한국은행": "기관전망",
    "일본은행": "기관전망",
    "금융투자협회": "설문컨센서스",
    "JCER": "설문컨센서스",
    "ECB": "설문컨센서스",
}

IMF_COUNTRIES = {"한국": "KOR", "미국": "USA", "일본": "JPN", "독일": "DEU", "프랑스": "FRA"}
OECD_COUNTRIES = {"한국": "KOR", "미국": "USA", "일본": "JPN", "독일": "DEU", "프랑스": "FRA"}

REQUEST_TIMEOUT = 30

RESULTS = []  # [(출처, 상태, 비고)] — 상태: 성공/실패/차단/건너뜀


def log_result(source, status, note=""):
    RESULTS.append((source, status, note))
    print(f"  [{status}] {source}" + (f" — {note}" if note else ""))


def _is_blocked(resp_or_exc):
    """403/401/429, 또는 흔한 WAF 차단 문구가 보이면 '차단'으로 판정. 차단 페이지에
    base64 인라인 이미지가 앞부분을 채우는 경우가 있어(예: ECB Data Portal 차단
    페이지) 응답 본문 전체에서 검색한다(응답이 아주 크면 앞 200KB만)."""
    if isinstance(resp_or_exc, requests.Response):
        if resp_or_exc.status_code in (401, 403, 429):
            return True
        text = (resp_or_exc.text[:200_000] if resp_or_exc.text else "").lower()
        for phrase in (
            "access denied",
            "access has been blocked",
            "security concerns",
            "blocked due to",
            "just a moment",
            "cloudflare",
        ):
            if phrase in text:
                return True
    return False


def this_and_next_year():
    y = date.today().year
    return y, y + 1


# ---------------------------------------------------------------------------
# 1. IMF World Economic Outlook (DataMapper API)
# ---------------------------------------------------------------------------

IMF_INDICATORS = {
    "실질GDP성장률": "NGDP_RPCH",
    "소비자물가상승률": "PCPIPCH",
    "실업률": "LUR",
}


def collect_imf():
    rows = []
    try:
        meta_resp = requests.get("https://www.imf.org/external/datamapper/api/v1/indicators", timeout=REQUEST_TIMEOUT)
        if _is_blocked(meta_resp):
            log_result("IMF WEO DataMapper", "차단", f"HTTP {meta_resp.status_code}")
            return rows
        meta_resp.raise_for_status()
        meta = meta_resp.json()["indicators"]
    except requests.RequestException as e:
        log_result("IMF WEO DataMapper", "실패", str(e))
        return rows
    except Exception as e:  # noqa: BLE001
        log_result("IMF WEO DataMapper", "실패", f"메타데이터 파싱 오류: {e}")
        return rows

    this_year, next_year = this_and_next_year()
    n_rows = 0
    for label, code in IMF_INDICATORS.items():
        try:
            url = f"https://www.imf.org/external/datamapper/api/v1/{code}"
            resp = requests.get(url, timeout=REQUEST_TIMEOUT)
            if _is_blocked(resp):
                log_result(f"IMF WEO - {label}", "차단", f"HTTP {resp.status_code}")
                continue
            resp.raise_for_status()
            values = resp.json()["values"][code]
            vintage_iso = meta.get(code, {}).get("last-modified", "")[:10] or date.today().isoformat()
            for kr_name, iso3 in IMF_COUNTRIES.items():
                country_vals = values.get(iso3, {})
                for year in (this_year, next_year):
                    v = country_vals.get(str(year))
                    if v is None:
                        continue
                    rows.append(
                        {
                            "발표일": vintage_iso,
                            "기관": "IMF",
                            "종류": "기관전망",
                            "국가": kr_name,
                            "지표": label,
                            "대상기간": str(year),
                            "전망치": round(float(v), 3),
                            "출처URL": "https://www.imf.org/external/datamapper/datasets/WEO",
                        }
                    )
                    n_rows += 1
        except requests.RequestException as e:
            log_result(f"IMF WEO - {label}", "실패", str(e))
        except Exception as e:  # noqa: BLE001
            log_result(f"IMF WEO - {label}", "실패", f"파싱 오류: {e}")

    if n_rows:
        log_result("IMF WEO DataMapper", "성공", f"{n_rows}건")
    return rows


# ---------------------------------------------------------------------------
# 2. OECD Economic Outlook (SDMX API, dataflow 버전 자동 탐지)
# ---------------------------------------------------------------------------

OECD_MEASURES = {
    "실질GDP성장률": "GDPV_ANNPCT",
    "실업률": "UNR",
    "단기금리(정책금리 근사)": "IRS",
}
_OECD_VERSION_RE = re.compile(r'version="(\d+)\.(\d+)"')


def _oecd_latest_version():
    url = "https://sdmx.oecd.org/public/rest/dataflow/OECD.ECO.MAD/DSD_EO@DF_EO/all?references=none&detail=allstubs"
    resp = requests.get(url, timeout=REQUEST_TIMEOUT)
    if _is_blocked(resp):
        raise RuntimeError(f"차단됨 (HTTP {resp.status_code})")
    resp.raise_for_status()
    versions = [(int(a), int(b)) for a, b in _OECD_VERSION_RE.findall(resp.text)]
    if not versions:
        raise RuntimeError("dataflow 버전을 찾지 못함")
    major, minor = max(versions)
    return f"{major}.{minor}"


def collect_oecd():
    rows = []
    try:
        version = _oecd_latest_version()
    except requests.RequestException as e:
        log_result("OECD Economic Outlook", "실패", f"버전 탐지 실패: {e}")
        return rows
    except RuntimeError as e:
        status = "차단" if "차단됨" in str(e) else "실패"
        log_result("OECD Economic Outlook", status, str(e))
        return rows

    n_rows = 0
    for label, measure in OECD_MEASURES.items():
        for kr_name, iso3 in OECD_COUNTRIES.items():
            time.sleep(0.3)  # 짧은 시간에 15개 요청을 연달아 보내 429를 맞는 일을 줄임
            try:
                url = (
                    f"https://sdmx.oecd.org/public/rest/data/OECD.ECO.MAD,DSD_EO@DF_EO,{version}/"
                    f"{iso3}.{measure}.A?startPeriod={date.today().year - 1}&dimensionAtObservation=AllDimensions&format=jsondata"
                )
                resp = requests.get(url, timeout=REQUEST_TIMEOUT)
                if resp.status_code == 404:
                    continue  # 해당 국가·지표 조합이 없을 수 있음(정상)
                if _is_blocked(resp):
                    log_result(f"OECD - {label}/{kr_name}", "차단", f"HTTP {resp.status_code}")
                    continue
                resp.raise_for_status()
                data = resp.json()["data"]
                vintage_label = data["structure"]["name"]  # 예: "Economic Outlook 119"
                obs = data["dataSets"][0].get("observations", {})
                dims = data["structure"]["dimensions"]["observation"]
                time_values = next(d["values"] for d in dims if d["id"] == "TIME_PERIOD")
                times = [v["id"] for v in time_values]
                vintage_iso = date.today().isoformat()  # OECD 응답엔 발표 '일자'가 없어 조회일로 근사, 에디션명은 비고에 남김
                this_year, next_year = this_and_next_year()
                for key, val in obs.items():
                    idx = int(key.split(":")[-1])
                    year = times[idx]
                    if year not in (str(this_year), str(next_year)):
                        continue
                    if val[0] is None:
                        continue
                    rows.append(
                        {
                            "발표일": vintage_iso,
                            "기관": "OECD",
                            "종류": "기관전망",
                            "국가": kr_name,
                            "지표": label,
                            "대상기간": f"{year} ({vintage_label})",
                            "전망치": round(float(val[0]), 3),
                            "출처URL": "https://www.oecd.org/en/publications/oecd-economic-outlook_16097408.html",
                        }
                    )
                    n_rows += 1
            except requests.RequestException as e:
                log_result(f"OECD - {label}/{kr_name}", "실패", str(e))
            except Exception as e:  # noqa: BLE001
                log_result(f"OECD - {label}/{kr_name}", "실패", f"파싱 오류: {e}")

    if n_rows:
        log_result("OECD Economic Outlook", "성공", f"{n_rows}건 (dataflow v{version})")
    return rows


# ---------------------------------------------------------------------------
# 3. 연준 SEP (Summary of Economic Projections)
# ---------------------------------------------------------------------------

SEP_ROW_LABELS = {
    "Change in real GDP": "실질GDP성장률",
    "Unemployment rate": "실업률",
    "PCE inflation": "PCE물가상승률",
    "Federal funds rate": "정책금리(중간값)",
}


def collect_fed_sep():
    rows = []
    try:
        cal_resp = requests.get("https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm", timeout=REQUEST_TIMEOUT)
        if _is_blocked(cal_resp):
            log_result("연준 SEP", "차단", f"캘린더 페이지 HTTP {cal_resp.status_code}")
            return rows
        cal_resp.raise_for_status()
        sep_dates = sorted(set(re.findall(r"fomcprojtabl(\d{8})\.htm", cal_resp.text)))
        past_dates = [d for d in sep_dates if d <= date.today().strftime("%Y%m%d")]
        if not past_dates:
            log_result("연준 SEP", "실패", "SEP 발표일을 찾지 못함")
            return rows
        latest = past_dates[-1]
        sep_url = f"https://www.federalreserve.gov/monetarypolicy/fomcprojtabl{latest}.htm"
        resp = requests.get(sep_url, timeout=REQUEST_TIMEOUT)
        if _is_blocked(resp):
            log_result("연준 SEP", "차단", f"HTTP {resp.status_code}")
            return rows
        resp.raise_for_status()
    except requests.RequestException as e:
        log_result("연준 SEP", "실패", str(e))
        return rows

    try:
        tables = pd.read_html(resp.text)
        main_table = tables[0]
        vintage_iso = f"{latest[:4]}-{latest[4:6]}-{latest[6:]}"
        this_year, next_year = this_and_next_year()
        for _, row in main_table.iterrows():
            var_label = row.iloc[0]
            if not isinstance(var_label, str) or var_label not in SEP_ROW_LABELS:
                continue
            indicator = SEP_ROW_LABELS[var_label]
            for year in (this_year, next_year):
                col = ("Median1", str(year))
                if col not in main_table.columns:
                    continue
                val = row.get(col)
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    continue
                rows.append(
                    {
                        "발표일": vintage_iso,
                        "기관": "연준(FOMC SEP)",
                        "종류": "기관전망",
                        "국가": "미국",
                        "지표": indicator,
                        "대상기간": str(year),
                        "전망치": round(float(val), 3),
                        "출처URL": sep_url,
                    }
                )
    except Exception as e:  # noqa: BLE001
        log_result("연준 SEP", "실패", f"표 파싱 오류: {e}")
        return rows

    log_result("연준 SEP", "성공", f"{len(rows)}건 ({vintage_iso} 발표분)")
    return rows


# ---------------------------------------------------------------------------
# 4. ECB staff / Eurosystem staff macroeconomic projections
# ---------------------------------------------------------------------------

ECB_SHEET_INDICATOR = {
    "Real Economy (annual)": "실질GDP성장률",
    "Prices (annual)": "HICP상승률",
}


def collect_ecb_projections():
    rows = []
    index_url = "https://www.ecb.europa.eu/press/projections/html/index.en.html"
    try:
        resp = requests.get(index_url, timeout=REQUEST_TIMEOUT)
        if _is_blocked(resp):
            log_result("ECB staff projections", "차단", f"HTTP {resp.status_code}")
            return rows
        resp.raise_for_status()
    except requests.RequestException as e:
        log_result("ECB staff projections", "실패", str(e))
        return rows

    m = re.search(r'href="(/pub/pdf/other/ecb\.projections\d{6}_ecbstaff_annex[^"]*\.xlsx)"', resp.text)
    if not m:
        log_result("ECB staff projections", "실패", "Excel 다운로드 링크를 찾지 못함")
        return rows
    xlsx_url = "https://www.ecb.europa.eu" + m.group(1)
    vintage_month = re.search(r"projections(\d{6})_ecbstaff", xlsx_url)
    vintage_iso = f"{vintage_month.group(1)[:4]}-{vintage_month.group(1)[4:]}-01" if vintage_month else date.today().isoformat()

    try:
        xresp = requests.get(xlsx_url, timeout=REQUEST_TIMEOUT)
        if _is_blocked(xresp):
            log_result("ECB staff projections", "차단", f"Excel HTTP {xresp.status_code}")
            return rows
        xresp.raise_for_status()
        tmp_path = os.path.join(DATA_DIR, "_tmp_ecb_projections.xlsx")
        with open(tmp_path, "wb") as f:
            f.write(xresp.content)
        wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
        this_year, next_year = this_and_next_year()
        for sheet_name, indicator in ECB_SHEET_INDICATOR.items():
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            for r in ws.iter_rows(values_only=True):
                if not r or r[0] is None or not hasattr(r[0], "year"):
                    continue
                year = r[0].year
                if year not in (this_year, next_year):
                    continue
                if r[1] is None:
                    continue
                rows.append(
                    {
                        "발표일": vintage_iso,
                        "기관": "ECB",
                        "종류": "기관전망",
                        "국가": "유로존",
                        "지표": indicator,
                        "대상기간": str(year),
                        "전망치": round(float(r[1]), 3),
                        "출처URL": xlsx_url,
                    }
                )
        wb.close()
        os.remove(tmp_path)
    except requests.RequestException as e:
        log_result("ECB staff projections", "실패", str(e))
        return rows
    except Exception as e:  # noqa: BLE001
        log_result("ECB staff projections", "실패", f"Excel 파싱 오류: {e}")
        return rows

    log_result("ECB staff projections", "성공", f"{len(rows)}건 (유로존 집계, 국가별 분해 없음)")
    return rows


# ---------------------------------------------------------------------------
# 5. 필라델피아 연준 Survey of Professional Forecasters (SPF)
# ---------------------------------------------------------------------------

# CPIA/CPIB = 당해년/내년 연평균 CPI 인플레이션 전망, UNEMPA/UNEMPB = 당해년/내년
# 연평균 실업률 전망 (Philadelphia Fed SPF 문서화된 관례). 이 관례로 이번에
# 확인하지 못한 실질GDP 성장률 연간 전망 파일은 이번 수집에서 제외했다.
SPF_FILES = {
    "소비자물가상승률(CPI)": ("Median_CPI_Level.xlsx", "CPIA", "CPIB"),
    "실업률": ("Median_UNEMP_Level.xlsx", "UNEMPA", "UNEMPB"),
}
SPF_BASE = "https://www.philadelphiafed.org/-/media/FRBP/Assets/Surveys-And-Data/survey-of-professional-forecasters/data-files/files/"


def collect_philly_fed_spf():
    rows = []
    this_year, next_year = this_and_next_year()
    for label, (fname, col_this, col_next) in SPF_FILES.items():
        url = SPF_BASE + fname
        try:
            resp = requests.get(url, timeout=REQUEST_TIMEOUT)
            if _is_blocked(resp):
                log_result(f"필라델피아 연준 SPF - {label}", "차단", f"HTTP {resp.status_code}")
                continue
            resp.raise_for_status()
            tmp_path = os.path.join(DATA_DIR, "_tmp_spf.xlsx")
            with open(tmp_path, "wb") as f:
                f.write(resp.content)
            wb = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            data_rows = [r for r in ws.iter_rows(values_only=True) if r and r[0] is not None and str(r[0]).strip() not in ("YEAR",)]
            header = [c for c in ws.iter_rows(values_only=True)][0]
            os.remove(tmp_path)
            wb.close()
            if col_this not in header or col_next not in header:
                log_result(f"필라델피아 연준 SPF - {label}", "실패", f"열 {col_this}/{col_next}을 찾지 못함")
                continue
            idx_this, idx_next = header.index(col_this), header.index(col_next)
            # 마지막(가장 최근) 설문 회차 행 사용
            valid = [r for r in data_rows if isinstance(r[0], (int, float))]
            if not valid:
                continue
            last = valid[-1]
            survey_year, survey_quarter = int(last[0]), int(last[1])
            vintage_iso = _spf_vintage_date(survey_year, survey_quarter)
            for target_year, idx in ((this_year, idx_this), (next_year, idx_next)):
                v = last[idx] if idx < len(last) else None
                if v is None or v == "#N/A":
                    continue
                rows.append(
                    {
                        "발표일": vintage_iso,
                        "기관": "필라델피아 연준(SPF)",
                        "종류": "설문컨센서스",
                        "국가": "미국",
                        "지표": label,
                        "대상기간": str(target_year),
                        "전망치": round(float(v), 3),
                        "출처URL": "https://www.philadelphiafed.org/surveys-and-data/data-files",
                    }
                )
        except requests.RequestException as e:
            log_result(f"필라델피아 연준 SPF - {label}", "실패", str(e))
        except Exception as e:  # noqa: BLE001
            log_result(f"필라델피아 연준 SPF - {label}", "실패", f"파싱 오류: {e}")

    if rows:
        log_result("필라델피아 연준 SPF", "성공", f"{len(rows)}건")
    return rows


def _spf_vintage_date(year, quarter):
    month = {1: 2, 2: 5, 3: 8, 4: 11}.get(quarter, 1)
    return date(year, month, 15).isoformat()


# ---------------------------------------------------------------------------
# 6. ECB Survey of Professional Forecasters (SPF) — 접근 시도 후 차단 시 기록만
# ---------------------------------------------------------------------------


def collect_ecb_spf():
    rows = []
    url = "https://data-api.ecb.europa.eu/service/data/SPF/Q.U2.HICP.POINT.P1Y0?format=csvdata"
    try:
        resp = requests.get(url, timeout=REQUEST_TIMEOUT)
        if _is_blocked(resp):
            log_result("ECB SPF", "차단", "data-api.ecb.europa.eu WAF 차단 — forecasts_manual.csv에 수동 입력 필요")
            return rows
        resp.raise_for_status()
        log_result("ECB SPF", "실패", "응답은 받았으나 파서 미구현(이번 조사에서는 차단만 예상함)")
    except requests.RequestException as e:
        log_result("ECB SPF", "실패", str(e))
    return rows


# ---------------------------------------------------------------------------
# 7. 시장금리 (ECOS 통안증권 1년·국고채 3년, FRED DGS2·T5YIE·T10YIE) — 매번 실행
# ---------------------------------------------------------------------------

ECOS_MARKET_SERIES = {
    "통안증권(1년)": "010400001",
    "국고채(3년)": "010200000",
}
FRED_MARKET_SERIES = {
    "국채(2년)": "DGS2",
    "기대인플레이션(5년)": "T5YIE",
    "기대인플레이션(10년)": "T10YIE",
}


def _month_key(d):
    return d.strftime("%Y-%m")


def collect_ecos_market_rates(ecos_key):
    rows = []
    if not ecos_key:
        log_result("ECOS 시장금리", "건너뜀", "ECOS_API_KEY 없음")
        return rows
    start = (date.today() - timedelta(days=760)).strftime("%Y%m%d")
    end = date.today().strftime("%Y%m%d")
    for label, item_code in ECOS_MARKET_SERIES.items():
        try:
            url = f"https://ecos.bok.or.kr/api/StatisticSearch/{ecos_key}/json/kr/1/1000/817Y002/D/{start}/{end}/{item_code}"
            resp = requests.get(url, timeout=REQUEST_TIMEOUT)
            if _is_blocked(resp):
                log_result(f"ECOS - {label}", "차단", f"HTTP {resp.status_code}")
                continue
            resp.raise_for_status()
            data = resp.json()
            if "StatisticSearch" not in data:
                log_result(f"ECOS - {label}", "실패", str(data)[:200])
                continue
            by_month = {}
            for r in data["StatisticSearch"]["row"]:
                d = datetime.strptime(r["TIME"], "%Y%m%d").date()
                by_month.setdefault(_month_key(d), []).append(float(r["DATA_VALUE"]))
            current_month = _month_key(date.today())
            for mk, vals in sorted(by_month.items()):
                if mk == current_month:
                    continue  # 진행 중인 달은 제외(불완전 평균)
                rows.append(
                    {
                        "발표일": mk + "-01",
                        "기관": "한국은행(ECOS)",
                        "종류": "시장가격",
                        "국가": "한국",
                        "지표": label,
                        "대상기간": mk,
                        "전망치": round(sum(vals) / len(vals), 4),
                        "출처URL": "https://ecos.bok.or.kr",
                    }
                )
        except requests.RequestException as e:
            log_result(f"ECOS - {label}", "실패", str(e))
        except Exception as e:  # noqa: BLE001
            log_result(f"ECOS - {label}", "실패", f"파싱 오류: {e}")

    if rows:
        log_result("ECOS 시장금리", "성공", f"{len(rows)}건(월평균)")
    return rows


def collect_fred_market_rates(fred_key):
    rows = []
    if not fred_key:
        log_result("FRED 시장금리", "건너뜀", "FRED_API_KEY 없음")
        return rows
    start = (date.today() - timedelta(days=760)).strftime("%Y-%m-%d")
    end = date.today().strftime("%Y-%m-%d")
    for label, series_id in FRED_MARKET_SERIES.items():
        try:
            url = "https://api.stlouisfed.org/fred/series/observations"
            params = {
                "series_id": series_id,
                "api_key": fred_key,
                "file_type": "json",
                "observation_start": start,
                "observation_end": end,
            }
            resp = requests.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if _is_blocked(resp):
                log_result(f"FRED - {label}", "차단", f"HTTP {resp.status_code}")
                continue
            resp.raise_for_status()
            obs = resp.json().get("observations", [])
            by_month = {}
            for o in obs:
                if o["value"] == ".":
                    continue
                d = datetime.strptime(o["date"], "%Y-%m-%d").date()
                by_month.setdefault(_month_key(d), []).append(float(o["value"]))
            current_month = _month_key(date.today())
            for mk, vals in sorted(by_month.items()):
                if mk == current_month:
                    continue
                rows.append(
                    {
                        "발표일": mk + "-01",
                        "기관": "FRED",
                        "종류": "시장가격",
                        "국가": "미국",
                        "지표": label,
                        "대상기간": mk,
                        "전망치": round(sum(vals) / len(vals), 4),
                        "출처URL": f"https://fred.stlouisfed.org/series/{series_id}",
                    }
                )
        except requests.RequestException as e:
            log_result(f"FRED - {label}", "실패", str(e))
        except Exception as e:  # noqa: BLE001
            log_result(f"FRED - {label}", "실패", f"파싱 오류: {e}")

    if rows:
        log_result("FRED 시장금리", "성공", f"{len(rows)}건(월평균)")
    return rows


# ---------------------------------------------------------------------------
# CSV 입출력 / 병합
# ---------------------------------------------------------------------------


def read_csv_rows(path, header):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        return [row for row in reader]


def dedup_key(row):
    return (row["발표일"], row["기관"], row["국가"], row["지표"], row["대상기간"])


def append_new_rows(existing_rows, new_rows):
    existing_keys = {dedup_key(r) for r in existing_rows}
    added = []
    for r in new_rows:
        r_str = {k: str(v) for k, v in r.items()}
        if dedup_key(r_str) in existing_keys:
            continue
        existing_rows.append(r_str)
        existing_keys.add(dedup_key(r_str))
        added.append(r_str)
    return added


def write_csv(path, header, rows):
    rows_sorted = sorted(rows, key=lambda r: (r["발표일"], r["기관"], r["국가"], r["지표"], r["대상기간"]))
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=header)
        writer.writeheader()
        for r in rows_sorted:
            writer.writerow({k: r.get(k, "") for k in header})


def ensure_manual_csv():
    if os.path.exists(MANUAL_CSV_PATH):
        return
    with open(MANUAL_CSV_PATH, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(MANUAL_HEADER)


def load_manual_rows():
    rows = read_csv_rows(MANUAL_CSV_PATH, MANUAL_HEADER)
    out = []
    for r in rows:
        if not r.get("발표일") or not r.get("전망치"):
            continue  # 빈 행(헤더만 있는 상태 등)은 건너뜀
        out.append(
            {
                "발표일": r["발표일"],
                "기관": r["기관"],
                "종류": MANUAL_INSTITUTION_KIND.get(r["기관"], "수동입력"),
                "국가": r["국가"],
                "지표": r["지표"],
                "대상기간": r["대상기간"],
                "전망치": r["전망치"],
                "응답비율_동결": r.get("응답비율_동결") or None,
                "응답비율_인상": r.get("응답비율_인상") or None,
                "응답비율_인하": r.get("응답비율_인하") or None,
                "출처URL": r.get("출처URL", ""),
                "비고": r.get("비고", ""),
            }
        )
    return out


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(os.path.dirname(JSON_PATH), exist_ok=True)
    ensure_manual_csv()

    load_dotenv(dotenv_path=os.path.join(BASE_DIR, ".env"))
    fred_key = os.environ.get("FRED_API_KEY")
    ecos_key = os.environ.get("ECOS_API_KEY")

    is_monday = date.today().weekday() == 0
    force_full = "--full" in sys.argv  # 수동 테스트용(요일 무시하고 전체 실행)

    print(f"=== forecast_collect.py 실행: {date.today().isoformat()} (월요일: {is_monday}) ===\n")

    new_rows = []

    print("-- 기관 전망 / 설문 컨센서스 (매주 월요일만) --")
    if is_monday or force_full:
        new_rows += collect_imf()
        new_rows += collect_oecd()
        new_rows += collect_fed_sep()
        new_rows += collect_ecb_projections()
        new_rows += collect_philly_fed_spf()
        new_rows += collect_ecb_spf()
    else:
        log_result("기관 전망 전체", "건너뜀", "월요일이 아님(주 1회만 조회)")

    print("\n-- 시장금리 (매번 실행) --")
    new_rows += collect_ecos_market_rates(ecos_key)
    new_rows += collect_fred_market_rates(fred_key)

    existing_rows = read_csv_rows(CSV_PATH, CSV_HEADER)
    added = append_new_rows(existing_rows, new_rows)
    write_csv(CSV_PATH, CSV_HEADER, existing_rows)
    print(f"\n=== data/forecasts.csv: 기존 {len(existing_rows) - len(added)}건 + 신규 {len(added)}건 = 총 {len(existing_rows)}건 ===")

    manual_rows = load_manual_rows()
    print(f"=== data/forecasts_manual.csv: {len(manual_rows)}건 로드 ===")

    csv_rows_for_json = [
        {
            "발표일": r["발표일"],
            "기관": r["기관"],
            "종류": r["종류"],
            "국가": r["국가"],
            "지표": r["지표"],
            "대상기간": r["대상기간"],
            "전망치": float(r["전망치"]) if r.get("전망치") not in (None, "") else None,
            "응답비율_동결": None,
            "응답비율_인상": None,
            "응답비율_인하": None,
            "출처URL": r.get("출처URL", ""),
            "비고": "",
        }
        for r in existing_rows
    ]
    for r in manual_rows:
        csv_rows_for_json.append(
            {
                **r,
                "전망치": float(r["전망치"]) if r.get("전망치") not in (None, "") else None,
                "응답비율_동결": float(r["응답비율_동결"]) if r.get("응답비율_동결") not in (None, "") else None,
                "응답비율_인상": float(r["응답비율_인상"]) if r.get("응답비율_인상") not in (None, "") else None,
                "응답비율_인하": float(r["응답비율_인하"]) if r.get("응답비율_인하") not in (None, "") else None,
            }
        )

    result_table = [{"출처": s, "상태": st, "비고": n} for s, st, n in RESULTS]

    output = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "rows": csv_rows_for_json,
        "collection_results": result_table,
    }
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n=== {JSON_PATH} 생성 (rows {len(csv_rows_for_json)}건) ===")

    print("\n=== 수집 결과 요약 ===")
    for s, st, n in RESULTS:
        print(f"  [{st}] {s}" + (f" — {n}" if n else ""))


if __name__ == "__main__":
    main()

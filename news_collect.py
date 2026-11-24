"""크게 움직인 지표에 관련 뉴스를 붙이는 스크립트.

흐름 (매일 collect.py 다음에 실행):
  1. site/data/data.json 에서 직전 대비 크게 움직인 지표를 고른다(사이트 '이번에 달라진 것'과 같은 규칙).
  2. 지표 발표일 앞뒤 기간의 뉴스를 검색한다. 한국 지표는 네이버 뉴스 API,
     나머지는 구글 뉴스 RSS. 네이버 키가 없으면 한국도 구글 뉴스(한국어)로 찾는다.
  3. Claude 가 후보 뉴스 중 관련 있는 것만 고르고 한 줄로 정리한다.
     - 인과를 단정하지 않는다("~때문" 금지, "같은 시기에" 표현).
     - 후보 목록에 없는 기사는 고를 수 없다(번호로만 고르고, 링크는 코드가 붙인다).
     - 확신이 없거나 규칙 위반이면 '확인 필요'로 표시한다 → 사람이 확인.
  4. site/data/news.json 에 쓰고, site/data/news_log.csv 에 실행 기록(확인 필요 비율)을 남긴다.

같은 지표·같은 기준일은 이전 결과를 재사용해서 Claude 호출은 새로 움직인 지표에만 쓴다.
"""

import csv
import email.utils
import html
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote_plus

import requests
from dotenv import load_dotenv

load_dotenv()

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(ROOT, "site", "data", "data.json")
NEWS_JSON = os.path.join(ROOT, "site", "data", "news.json")
NEWS_LOG = os.path.join(ROOT, "site", "data", "news_log.csv")
TEXT_RULES_JS = os.path.join(ROOT, "site", "text_rules.js")

MODEL = "claude-opus-5-5"
MAX_MOVERS = 5            # 뉴스를 붙일 지표 수
MAX_CANDIDATES = 12       # 지표 하나당 Claude 에게 보여줄 후보 기사 수
MAX_PICK = 3              # 지표 하나당 최대 연결 기사 수
RECENT_RELEASE_DAYS = 60  # 발표된 지 이 기간이 지난 변화는 뉴스를 붙이지 않음
HTTP_TIMEOUT = 20
USER_AGENT = "Mozilla/5.0 (compatible; macro-tracker-news/1.0)"

# ---- 사이트(text_rules.js)와 같은 변화 계산 규칙 ----
PERCENT_POINT_INDICATORS = {
    "CPI 헤드라인", "CPI 근원", "HICP 헤드라인", "HICP 근원", "PCE 근원",
    "실질GDP 성장률", "실업률", "기준금리", "정책금리", "정책금리(상단)", "국채10년",
}
THRESHOLD_GROUP = {
    "기준금리": "rate_level", "정책금리": "rate_level", "정책금리(상단)": "rate_level",
    "WTI": "price_level", "두바이유": "price_level", "브렌트": "price_level",
}
BIG = {"default": 1.0, "rate_level": 0.5, "price_level": 5.0}
SMALL = {"default": 0.3, "rate_level": 0.25, "price_level": 1.5}
OIL = {"WTI", "두바이유", "브렌트"}

# ---- 검색어 ----
COUNTRY_EN = {"미국": "US", "일본": "Japan", "독일": "Germany", "프랑스": "France", "한국": "South Korea"}
INDICATOR_EN = {
    "CPI 헤드라인": "inflation CPI", "CPI 근원": "core inflation",
    "HICP 헤드라인": "inflation HICP", "HICP 근원": "core inflation",
    "PCE 근원": "core PCE inflation", "실질GDP 성장률": "GDP growth",
    "실업률": "unemployment rate", "비농업고용": "nonfarm payrolls", "취업자수": "employment",
    "국채10년": "10-year bond yield", "기준금리": "central bank interest rate decision",
    "정책금리": "central bank interest rate decision", "정책금리(상단)": "Fed interest rate decision",
}
OIL_EN = {"WTI": "WTI crude oil price", "두바이유": "Dubai crude oil price", "브렌트": "Brent crude oil price"}
INDICATOR_KO = {
    "CPI 헤드라인": "소비자물가", "CPI 근원": "근원물가", "실질GDP 성장률": "GDP 성장률",
    "실업률": "실업률", "취업자수": "취업자수 고용동향", "국채10년": "국고채 10년 금리",
    "기준금리": "한국은행 기준금리",
}


def log(msg):
    print(f"[news] {msg}", flush=True)


def load_banned_phrases():
    """사이트 금지 표현 목록을 text_rules.js 에서 그대로 읽는다(목록을 두 군데에 두지 않기 위해)."""
    try:
        src = open(TEXT_RULES_JS, encoding="utf-8").read()
        m = re.search(r"var BANNED_PHRASES = \[(.*?)\];", src, re.S)
        return re.findall(r'"([^"]+)"', m.group(1)) if m else []
    except OSError:
        return []


def group_of(indicator):
    return THRESHOLD_GROUP.get(indicator, "default")


def find_movers(observations, today):
    """지표별 마지막 두 값의 변화를 '많이' 기준으로 나눈 크기 순으로 고른다."""
    series = {}
    for o in observations:
        if o.get("value") is None:
            continue
        series.setdefault((o["country"], o["indicator"]), []).append(o)
    movers, seen = [], set()
    for (country, indicator), rows in series.items():
        rows.sort(key=lambda r: r["date"])
        if len(rows) < 2:
            continue
        last, prev = rows[-1], rows[-2]
        if indicator in PERCENT_POINT_INDICATORS:
            diff, unit = last["value"] - prev["value"], "%p"
        else:
            if not prev["value"]:
                continue
            diff, unit = (last["value"] / prev["value"] - 1) * 100, "%"
        g = group_of(indicator)
        if abs(diff) < SMALL[g]:
            continue
        rel = last.get("release_date") or last["date"]
        try:
            rel_d = date.fromisoformat(rel[:10])
        except ValueError:
            continue
        if (today - rel_d).days > RECENT_RELEASE_DAYS:
            continue
        # 유가처럼 여러 나라에 같은 시계열이 있으면 한 번만
        dup = (indicator, last["date"], round(last["value"], 6), round(prev["value"], 6))
        if dup in seen:
            continue
        seen.add(dup)
        movers.append({
            "key": f"{country}|{indicator}|{last['date']}",
            "country": country, "indicator": indicator,
            "date": last["date"], "release_date": rel[:10],
            "value": round(last["value"], 4), "prev": round(prev["value"], 4),
            "diff": round(diff, 4), "unit": unit,
            "scaled": abs(diff) / BIG[g],
        })
    movers.sort(key=lambda m: m["scaled"], reverse=True)
    return movers[:MAX_MOVERS]


def news_window(m):
    rel = date.fromisoformat(m["release_date"])
    return rel - timedelta(days=3), rel + timedelta(days=5)


# ---- 뉴스 검색 ----
def google_news(query, start, end, lang="en"):
    q = f"{query} after:{start.isoformat()} before:{end.isoformat()}"
    if lang == "ko":
        tail = "hl=ko&gl=KR&ceid=KR:ko"
    else:
        tail = "hl=en-US&gl=US&ceid=US:en"
    url = f"https://news.google.com/rss/search?q={quote_plus(q)}&{tail}"
    r = requests.get(url, timeout=HTTP_TIMEOUT, headers={"User-Agent": USER_AGENT})
    r.raise_for_status()
    root = ET.fromstring(r.content)
    out = []
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        src_el = item.find("source")
        source = (src_el.text or "").strip() if src_el is not None else ""
        if source and title.endswith(" - " + source):
            title = title[: -len(" - " + source)]
        out.append({
            "title": html.unescape(title),
            "url": (item.findtext("link") or "").strip(),
            "source": source,
            "published": _rfc822_to_date(item.findtext("pubDate")),
        })
    return out


def naver_news(query, start, end):
    cid, secret = os.environ.get("NAVER_CLIENT_ID"), os.environ.get("NAVER_CLIENT_SECRET")
    r = requests.get(
        "https://openapi.naver.com/v1/search/news.json",
        params={"query": query, "display": 100, "sort": "sim"},
        headers={"X-Naver-Client-Id": cid, "X-Naver-Client-Secret": secret},
        timeout=HTTP_TIMEOUT,
    )
    r.raise_for_status()
    out = []
    for it in r.json().get("items", []):
        pub = _rfc822_to_date(it.get("pubDate"))
        if pub and not (start.isoformat() <= pub <= end.isoformat()):
            continue
        link = it.get("originallink") or it.get("link") or ""
        out.append({
            "title": html.unescape(re.sub(r"</?b>", "", it.get("title", ""))),
            "url": link,
            "source": re.sub(r"^https?://(www\.)?", "", link).split("/")[0],
            "published": pub,
            "snippet": html.unescape(re.sub(r"</?b>", "", it.get("description", "")))[:200],
        })
    return out


def _rfc822_to_date(s):
    if not s:
        return ""
    try:
        return email.utils.parsedate_to_datetime(s).astimezone(timezone.utc).date().isoformat()
    except (TypeError, ValueError):
        return ""


def search_news(m):
    start, end = news_window(m)
    has_naver = bool(os.environ.get("NAVER_CLIENT_ID") and os.environ.get("NAVER_CLIENT_SECRET"))
    if m["indicator"] in OIL:
        return google_news(OIL_EN[m["indicator"]], start, end), "구글 뉴스"
    if m["country"] == "한국":
        q = INDICATOR_KO.get(m["indicator"], m["indicator"])
        if has_naver:
            return naver_news(q, start, end), "네이버 뉴스"
        return google_news(q, start, end, lang="ko"), "구글 뉴스(한국어)"
    q = f'{COUNTRY_EN.get(m["country"], m["country"])} {INDICATOR_EN.get(m["indicator"], m["indicator"])}'
    return google_news(q, start, end), "구글 뉴스"


def dedupe(articles):
    seen, out = set(), []
    for a in articles:
        k = re.sub(r"\W+", "", a["title"].lower())[:60]
        if not a["url"] or k in seen:
            continue
        seen.add(k)
        out.append(a)
    return out


# ---- Claude 로 관련 기사 고르기 ----
SYSTEM_PROMPT = """너는 거시경제 지표 사이트의 편집 보조야. 지표 하나의 변화와, 그 지표 발표일 앞뒤에 나온 뉴스 후보 목록을 받는다.

할 일:
1. 후보 중 이 지표 변화와 직접 관련 있는 기사만 골라 번호(id)로 답한다. 최대 3개. 관련 있는 게 없으면 빈 목록.
   - 같은 나라·같은 지표를 다룬 기사만 관련 있다고 본다. 다른 나라 지표나 일반 시황 기사는 고르지 않는다.
2. 고른 기사가 있으면, 지표 변화와 기사 내용을 잇는 한국어 한 문장(60자 안팎)을 쓴다.
   - 비전공자도 읽을 수 있게 쉬운 말로 쓰고 "~요"체로 끝낸다.
   - 인과를 단정하지 않는다. "~때문", "~로 인해" 대신 "같은 시기에 ~ 소식이 있었어요", "~ 보도가 함께 나왔어요"처럼 쓴다.
   - 앞으로의 방향을 예측하거나 투자 판단을 권하지 않는다.
   - 기사 제목에 없는 숫자나 사실을 지어내지 않는다.
3. status 를 정한다.
   - "linked": 고른 기사가 이 지표 변화를 분명히 다룬다.
   - "review": 관련이 애매하거나, 기사 날짜가 발표일과 멀거나, 제목만으로 판단하기 어렵다. 사람이 확인해야 한다.
   - "none": 관련 기사가 없다.
4. reason 에 그렇게 판단한 이유를 한 문장으로 쓴다(사람이 검토할 때 읽는다)."""

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "selected_ids": {"type": "array", "items": {"type": "integer"}},
        "summary": {"type": "string"},
        "status": {"type": "string", "enum": ["linked", "review", "none"]},
        "reason": {"type": "string"},
    },
    "required": ["selected_ids", "summary", "status", "reason"],
    "additionalProperties": False,
}


def describe_change(m):
    sign = "+" if m["diff"] > 0 else ""
    return (f'{m["country"]} {m["indicator"]}: {m["prev"]} → {m["value"]} '
            f'({sign}{round(m["diff"], 2)}{m["unit"]}), 기준월 {m["date"][:7]}, 발표일 {m["release_date"]}')


def ask_claude(client, m, candidates):
    import anthropic  # 키가 있을 때만 필요

    lines = [f"[{i}] {c['published']} | {c['source']} | {c['title']}"
             + (f" | {c['snippet']}" if c.get("snippet") else "")
             for i, c in enumerate(candidates)]
    user = "지표 변화:\n" + describe_change(m) + "\n\n뉴스 후보:\n" + "\n".join(lines)
    try:
        resp = client.beta.messages.create(
            model=MODEL,
            max_tokens=4000,
            system=SYSTEM_PROMPT,
            output_config={"effort": "low", "format": {"type": "json_schema", "schema": OUTPUT_SCHEMA}},
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            messages=[{"role": "user", "content": user}],
        )
    except anthropic.APIStatusError as e:
        log(f"Claude API 오류 {e.status_code}: {e.message}")
        return None
    except anthropic.APIConnectionError as e:
        log(f"Claude 연결 오류: {e}")
        return None
    if resp.stop_reason != "end_turn":
        log(f"Claude 응답 중단: {resp.stop_reason}")
        return None
    text = next((b.text for b in resp.content if b.type == "text"), "")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        log("Claude 응답 JSON 파싱 실패")
        return None


def ask_claude_cli(m, candidates):
    """Claude 구독(Max 등)으로 돌리는 경로: Claude Code CLI 헤드리스 모드.
    CLAUDE_CODE_OAUTH_TOKEN(`claude setup-token`으로 만든 1년 토큰)으로 인증한다.
    ANTHROPIC_API_KEY가 환경에 있으면 CLI가 그 키(API 과금)를 먼저 쓰므로 자식 프로세스에서는 지운다."""
    import subprocess
    import tempfile

    lines = [f"[{i}] {c['published']} | {c['source']} | {c['title']}"
             + (f" | {c['snippet']}" if c.get("snippet") else "")
             for i, c in enumerate(candidates)]
    user = "지표 변화:\n" + describe_change(m) + "\n\n뉴스 후보:\n" + "\n".join(lines)
    env = {k: v for k, v in os.environ.items() if k not in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")}
    env["DISABLE_AUTOUPDATER"] = "1"
    cmd = [
        "claude", "-p", "아래 지표 변화와 뉴스 후보를 규칙대로 판단해 JSON으로만 답해.",
        "--model", MODEL,
        "--system-prompt", SYSTEM_PROMPT,
        "--tools", "",
        "--max-turns", "3",
        "--output-format", "json",
        "--json-schema", json.dumps(OUTPUT_SCHEMA),
    ]
    try:
        # 저장소 밖(빈 임시 폴더)에서 실행해 프로젝트 설정·훅을 읽지 않게 한다
        with tempfile.TemporaryDirectory() as tmp:
            p = subprocess.run(cmd, input=user, capture_output=True, text=True,
                               timeout=300, env=env, cwd=tmp)
    except (OSError, subprocess.TimeoutExpired) as e:
        log(f"Claude CLI 실행 실패: {e!r}")
        return None
    try:
        out = json.loads(p.stdout)
    except json.JSONDecodeError:
        log(f"Claude CLI 출력 파싱 실패 (exit {p.returncode}): {p.stdout[:200]} {p.stderr[:200]}")
        return None
    result = out.get("structured_output")
    if not isinstance(result, dict):
        log(f"Claude CLI 결과 없음 (exit {p.returncode}): {str(out.get('result'))[:200]}")
        return None
    return result


def check_result(result, candidates, banned):
    """Claude 결과를 코드로 다시 검사한다. 어기면 '확인 필요'로 내린다."""
    flags = []
    ids = [i for i in result.get("selected_ids", []) if isinstance(i, int) and 0 <= i < len(candidates)]
    if len(ids) != len(result.get("selected_ids", [])):
        flags.append("후보에 없는 기사 번호를 골랐어요")
    ids = list(dict.fromkeys(ids))[:MAX_PICK]
    summary = (result.get("summary") or "").strip()
    status = result.get("status", "review")
    hit = [p for p in banned if p in summary]
    if hit:
        flags.append("금지 표현 포함: " + ", ".join(hit))
    if re.search(r"https?://", summary):
        flags.append("문장에 링크가 들어갔어요")
    if not ids:
        status = "none"
        summary = ""
    elif status == "none":
        status = "review"
        flags.append("기사를 골랐는데 관련 없음으로 표시했어요")
    if flags and status == "linked":
        status = "review"
    reason = (result.get("reason") or "").strip()
    if flags:
        reason = (reason + " / 자동 검사: " + "; ".join(flags)).strip(" /")
    return ids, summary, status, reason


# ---- 메인 ----
def main():
    today = date.today()
    data = json.load(open(DATA_JSON, encoding="utf-8"))
    movers = find_movers(data["observations"], today)
    log(f"뉴스를 붙일 지표 {len(movers)}개: " + ", ".join(m["key"] for m in movers))

    previous = {}
    if os.path.exists(NEWS_JSON):
        try:
            for it in json.load(open(NEWS_JSON, encoding="utf-8")).get("items", []):
                previous[it["key"]] = it
        except (OSError, ValueError, KeyError):
            previous = {}

    # AI 경로 선택: 구독 토큰(CLI)이 있으면 그걸 우선, 없으면 API 키(SDK), 둘 다 없으면 AI 없이
    client = None
    use_cli = bool(os.environ.get("CLAUDE_CODE_OAUTH_TOKEN")) or os.environ.get("NEWS_AI") == "cli"
    if use_cli:
        client = "cli"
        log("Claude 구독(Claude Code CLI)으로 연결해요")
    elif os.environ.get("ANTHROPIC_API_KEY"):
        import anthropic
        client = anthropic.Anthropic()
        log("Claude API 키로 연결해요")
    else:
        log("Claude 인증 정보 없음 → 기사 후보만 모으고 연결 문장은 만들지 않아요")

    banned = load_banned_phrases()
    items, calls = [], 0
    for m in movers:
        old = previous.get(m["key"])
        if old and old.get("ai") and old.get("status") in ("linked", "review", "none"):
            items.append(old)
            continue
        try:
            raw, provider = search_news(m)
        except (requests.RequestException, ET.ParseError) as e:
            log(f"{m['key']} 뉴스 검색 실패: {e}")
            raw, provider = [], "검색 실패"
        candidates = dedupe(raw)[:MAX_CANDIDATES]
        item = {k: m[k] for k in ("key", "country", "indicator", "date", "release_date", "value", "prev", "diff", "unit")}
        item.update({"provider": provider, "checked_at": today.isoformat(), "articles": [], "ai": False})
        if not candidates:
            item.update({"status": "none", "summary": "", "reason": "발표일 앞뒤 기간에 검색된 기사가 없어요"})
        elif client is None:
            item.update({"status": "review", "summary": "",
                         "reason": "AI 연결 전이라 관련 여부를 확인하지 않은 검색 결과예요",
                         "articles": candidates[:MAX_PICK]})
        else:
            calls += 1
            result = ask_claude_cli(m, candidates) if client == "cli" else ask_claude(client, m, candidates)
            if result is None:
                item.update({"status": "review", "summary": "", "reason": "AI 응답을 받지 못했어요",
                             "articles": candidates[:MAX_PICK]})
            else:
                ids, summary, status, reason = check_result(result, candidates, banned)
                item.update({"status": status, "summary": summary, "reason": reason,
                             "articles": [candidates[i] for i in ids], "ai": True})
        items.append(item)
        log(f"{m['key']}: {item['status']} ({len(item['articles'])}건, {provider})")

    out = {
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
        "model": MODEL if client else None,
        "items": items,
    }
    with open(NEWS_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    counts = {s: sum(1 for it in items if it["status"] == s) for s in ("linked", "review", "none")}
    new_log = not os.path.exists(NEWS_LOG)
    with open(NEWS_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_log:
            w.writerow(["날짜", "지표수", "연결", "확인필요", "관련없음", "새로_호출한_AI"])
        w.writerow([today.isoformat(), len(items), counts["linked"], counts["review"], counts["none"], calls])
    log(f"완료: 연결 {counts['linked']} · 확인 필요 {counts['review']} · 없음 {counts['none']} · AI 호출 {calls}회")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # 뉴스는 부가 기능 — 실패해도 지표 업데이트를 막지 않는다
        log(f"실패: {e!r}")
        sys.exit(0)

"""GitHub 이슈(사이트의 [승인]/[거절] 버튼으로 연 것)를 읽어 사람 판정으로 기록한다.

review.yml 이 저장소 주인이 연 '[검토]' 이슈에서만 실행한다.
이슈 본문·제목은 환경변수(ISSUE_BODY)로만 받는다 — 셸 명령에 끼워 넣지 않는다.

본문 형식(사이트가 미리 채워 줌):
    메모: <사용자가 적는 한 줄>
    <!-- review
    대상: 홈 | 달력
    지표키: 나라|지표|기준일
    판정: 승인 | 거절
    -->

결과 한 줄을 stdout 에 쓰고, 실패하면 종료코드 1 (이슈에 이유를 댓글로 남기기 위해).
"""

import csv
import json
import os
import re
import sys
from datetime import date

ROOT = os.path.dirname(os.path.abspath(__file__))
NEWS_JSON = os.path.join(ROOT, "site", "data", "news.json")
CAL_JSON = os.path.join(ROOT, "site", "data", "calendar_notes.json")
REVIEW_CSV = os.path.join(ROOT, "data", "news_review.csv")
FIELDS = ["대상", "지표키", "판정", "AI판정", "메모", "검토일"]
MEMO_PLACEHOLDER = "(여기에 판단 이유를 한 줄로 적어 주세요)"


def fail(msg):
    print(msg)
    sys.exit(1)


def parse(body):
    block = re.search(r"<!--\s*review\s*(.*?)-->", body, re.S)
    if not block:
        fail("판정 정보 블록을 찾지 못했어요. 사이트의 [승인]/[거절] 버튼으로 이슈를 열어 주세요.")
    fields = dict(re.findall(r"^\s*(대상|지표키|판정)\s*:\s*(.+?)\s*$", block.group(1), re.M))
    if fields.get("대상") not in ("홈", "달력"):
        fail("대상은 '홈' 또는 '달력'이어야 해요.")
    if fields.get("판정") not in ("승인", "거절"):
        fail("판정은 '승인' 또는 '거절'이어야 해요.")
    if not re.fullmatch(r"[^|\n]{1,20}\|[^|\n]{1,30}\|\d{4}-\d{2}-\d{2}", fields.get("지표키", "")):
        fail("지표키 형식이 올바르지 않아요.")
    memo_m = re.search(r"^\s*메모\s*:\s*(.*)$", body[: block.start()], re.M)
    memo = (memo_m.group(1).strip() if memo_m else "")
    if memo == MEMO_PLACEHOLDER:
        memo = ""
    memo = re.sub(r"[\r\n]+", " ", memo)[:200]
    return fields["대상"], fields["지표키"], fields["판정"], memo


def current_ai_status(target, key):
    """지금 사이트 데이터에 실제로 있는 항목인지 확인하고, 그때의 AI 판정을 돌려준다."""
    try:
        if target == "홈":
            for it in json.load(open(NEWS_JSON, encoding="utf-8")).get("items", []):
                if it["key"] == key:
                    return it.get("ai_status", it["status"])
        else:
            note = json.load(open(CAL_JSON, encoding="utf-8")).get("notes", {}).get(key)
            if note:
                return note.get("ai_status", note["status"])
    except (OSError, ValueError, KeyError):
        pass
    return None


def main():
    target, key, decision, memo = parse(os.environ.get("ISSUE_BODY", ""))
    ai = current_ai_status(target, key)
    if ai is None:
        fail(f"[{target}] {key} 항목을 지금 사이트 데이터에서 찾지 못했어요(이미 목록에서 빠졌을 수 있어요).")
    new_file = not os.path.exists(REVIEW_CSV)
    os.makedirs(os.path.dirname(REVIEW_CSV), exist_ok=True)
    with open(REVIEW_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        if new_file:
            w.writeheader()
        w.writerow({"대상": target, "지표키": key, "판정": decision, "AI판정": ai,
                    "메모": memo, "검토일": date.today().isoformat()})
    print(f"기록했어요: [{target}] {key} → {decision}" + (f" (메모: {memo})" if memo else ""))


if __name__ == "__main__":
    main()

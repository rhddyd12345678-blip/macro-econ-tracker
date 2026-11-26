"""뉴스 연결·달력 해석 사람 검토 도구 (HITL).

  python3 review.py list            # 검토할 항목 보기 ('확인 필요'와 이미 검토한 것)
  python3 review.py list --all      # AI가 '연결'로 판단한 것까지 전부 보기(표본 검토용)
  python3 review.py 승인 <번호> "메모"
  python3 review.py 거절 <번호> "메모"

검토 대상은 두 가지예요.
  [홈]  '이번에 달라진 것'에 붙은 같은 시기 뉴스   (site/data/news.json)
  [달력] 지난 발표를 눌렀을 때 나오는 기사·해석   (site/data/calendar_notes.json)

판정은 data/news_review.csv 에 쌓이고, news_collect.py 가 다음 실행 때 사이트에 반영한다.
바로 반영하려면 판정 후 `python3 news_collect.py` 를 한 번 돌린다(이미 판단한 항목은 AI를 다시 부르지 않는다).
"""

import csv
import json
import os
import sys
from datetime import date

ROOT = os.path.dirname(os.path.abspath(__file__))
NEWS_JSON = os.path.join(ROOT, "site", "data", "news.json")
CAL_JSON = os.path.join(ROOT, "site", "data", "calendar_notes.json")
REVIEW_CSV = os.path.join(ROOT, "data", "news_review.csv")
FIELDS = ["대상", "지표키", "판정", "AI판정", "메모", "검토일"]
STATUS_KO = {"linked": "연결", "review": "확인 필요", "none": "관련 없음"}


def _load(path, getter):
    try:
        return getter(json.load(open(path, encoding="utf-8")))
    except (OSError, ValueError):
        return []


def load_items():
    """번호는 list 와 승인/거절에서 똑같아야 하므로 항상 같은 순서의 전체 목록을 쓴다."""
    home = _load(NEWS_JSON, lambda d: [dict(it, _target="홈") for it in d.get("items", [])])
    cal = _load(CAL_JSON, lambda d: [dict(v, key=k, _target="달력") for k, v in d.get("notes", {}).items()])
    items = [it for it in home + cal if it.get("articles") or it.get("meaning")]
    items.sort(key=lambda it: (it["_target"] != "홈", it.get("ai_status", it["status"]) != "review", it["key"]))
    return items


def needs_look(it):
    return it.get("ai_status", it["status"]) == "review" or it.get("human")


def cmd_list(show_all):
    items = load_items()
    if not any(show_all or needs_look(it) for it in items):
        print("검토할 항목이 없어요. (AI가 '연결'로 판단한 것까지 보려면 --all)")
        return
    for n, it in enumerate(items, 1):
        if not (show_all or needs_look(it)):
            continue
        ai = STATUS_KO.get(it.get("ai_status", it["status"]), it["status"])
        done = f"  ← 검토함: {it['human']['decision']} ({it['human']['date']})" if it.get("human") else ""
        sign = "+" if (it.get("diff") or 0) > 0 else ""
        when = f" 발표 {it['release_date']}" if it["_target"] == "달력" and it.get("release_date") else ""
        period = it.get("date") or it.get("period") or ""
        print(f"\n[{n}] [{it['_target']}] {it['country']} {it['indicator']} {period[:7]}{when}: "
              f"{it.get('prev')} → {it.get('value')} ({sign}{round(it.get('diff') or 0, 2)}{it.get('unit', '')})"
              f"  AI 판정: {ai}{done}")
        if it.get("summary"):
            print(f"    AI 요약: {it['summary']}")
        if it.get("meaning"):
            print(f"    AI 해석: {it['meaning']}")
        print(f"    AI 이유: {it.get('reason', '')}")
        for a in it.get("articles", []):
            print(f"    - {a.get('published', '')} {a.get('source', '')} | {a['title']}")
            print(f"      {a['url']}")


def cmd_decide(decision, number, memo):
    items = load_items()
    if not (1 <= number <= len(items)):
        sys.exit(f"번호는 1~{len(items)} 사이여야 해요. 먼저 `python3 review.py list` 로 확인하세요.")
    it = items[number - 1]
    new_file = not os.path.exists(REVIEW_CSV)
    os.makedirs(os.path.dirname(REVIEW_CSV), exist_ok=True)
    with open(REVIEW_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        if new_file:
            w.writeheader()
        w.writerow({"대상": it["_target"], "지표키": it["key"], "판정": decision,
                    "AI판정": it.get("ai_status", it["status"]), "메모": memo,
                    "검토일": date.today().isoformat()})
    print(f"기록했어요: [{it['_target']}] {it['country']} {it['indicator']} → {decision}")


def main():
    args = sys.argv[1:]
    if not args or args[0] == "list":
        cmd_list("--all" in args)
    elif args[0] in ("승인", "거절") and len(args) >= 2 and args[1].isdigit():
        cmd_decide(args[0], int(args[1]), " ".join(args[2:]))
    else:
        print(__doc__)


if __name__ == "__main__":
    main()

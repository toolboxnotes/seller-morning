"""
계정 상태 · 이슈 대장

이슈마다 한 번만 수집. 대장은 reports/account_health/issue_registry.json 에 저장하고
보기 쉽게 issue_registry.xlsx 도 같이 만듦.

매일 실행할 때:
  - 항목 건수가 대장의 지난번 확인 건수와 같으면 → 들어가지 않고 상품도 다시 보지 않음
  - 건수가 바뀌면 → 들어가서 새 이슈만 대장에 추가, 목록에서 사라진 이슈는 ‘사라짐’으로 표시
  - 건수가 0이 되면 → 그 항목의 미해결 이슈를 모두 ‘사라짐’으로 표시
  - 상품 정보는 ASIN마다 한 번만 수집
"""

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

OPEN, GONE = "미해결", "사라짐"
DETAIL_FIELDS = ("review_details", "detail_text", "detail_file", "policy_title", "policy_url", "policy_text")


def _empty() -> dict:
    return {"categories": {}, "issues": {}, "products": {}}


# 예전 버전은 상태·항목 이름을 중국어로 저장 — 읽을 때 한국어로 바꿈
_OLD_TEXT = {"未解决": OPEN, "已消失": GONE,
             "涉嫌侵犯知识产权": "지식재산권 침해 의심", "知识产权投诉": "지식재산권 관련 불만",
             "商品真实性买家投诉": "상품 진위 관련 고객 불만", "商品状况买家投诉": "상품 상태 관련 고객 불만",
             "食品和商品安全问题": "식품 및 상품 안전 문제", "违反上架政策": "리스팅 정책 위반",
             "违反受限商品政策": "제한 상품 정책 위반", "违反买家商品评论政策": "고객 상품 리뷰 정책 위반",
             "其他违反政策的情况": "기타 정책 위반", "监管合规": "규정 준수"}


def _migrate(reg: dict) -> dict:
    for i in reg.get("issues", {}).values():
        for f_ in ("status", "cat_name"):
            if i.get(f_) in _OLD_TEXT:
                i[f_] = _OLD_TEXT[i[f_]]
    for c in reg.get("categories", {}).values():
        if isinstance(c, dict) and c.get("name") in _OLD_TEXT:
            c["name"] = _OLD_TEXT[c["name"]]
    return reg


def load(path: Path) -> dict:
    try:
        reg = json.loads(path.read_text(encoding="utf-8"))
        for k, v in _empty().items():
            reg.setdefault(k, v)
        return _migrate(reg)
    except Exception:
        return _empty()


def save(reg: dict, path: Path) -> None:
    path.write_text(json.dumps(reg, ensure_ascii=False, indent=2), encoding="utf-8")


def issue_id(cat_key: str, it: dict) -> str:
    """항목 + ASIN + SKU + 사유 + 날짜로 같은 이슈를 구분, 모두 없으면 원문 사용."""
    parts = [cat_key] + [str(it.get(k) or "").strip().lower() for k in ("asin", "sku", "reason", "date")]
    if not any(parts[1:]):
        parts.append(" | ".join(it.get("cells") or it.get("context") or []))
    return hashlib.sha1("␟".join(parts).encode("utf-8")).hexdigest()[:16]


def open_issues(reg: dict, cat_key: str = None) -> List[dict]:
    return [i for i in reg["issues"].values()
            if i["status"] == OPEN and (cat_key is None or i["cat_key"] == cat_key)]


def plan(reg: dict, policies: List[dict]) -> Tuple[List[dict], List[dict]]:
    """
    돌려줌 (들어가서 볼 항목, 변화 없어 건너뛴 항목).
    지난번에 상세를 제대로 읽었고 건수가 그대로일 때만 건너뜀.
    """
    visit, skip = [], []
    for p in policies:
        c = p.get("count")
        if not c:
            continue
        last = reg["categories"].get(p["key"], {}).get("last_count")
        missing_detail = any(not i.get("detail_ok") for i in open_issues(reg, p["key"]))
        (skip if last == c and not missing_detail else visit).append(p)
    return visit, skip


def apply(reg: dict, policies: List[dict], details: List[dict], now: datetime) -> Tuple[List[dict], List[dict]]:
    """이번 결과를 대장에 합침. 돌려줌 (새 이슈, 이번에 ‘사라짐’으로 표시한 이슈)."""
    ts = now.isoformat(timespec="minutes")
    new, gone = [], []
    visited = {d["key"]: d for d in details}

    for p in policies:
        key, c = p["key"], p.get("count")
        cat = reg["categories"].setdefault(key, {"name": p["name"]})
        cat["name"] = p["name"]

        if c == 0:
            # 건수 0: 전에 기록한 이슈는 모두 없어짐
            for i in open_issues(reg, key):
                i["status"], i["gone_at"] = GONE, ts
                gone.append(i)
            cat["last_count"], cat["checked_at"] = 0, ts
            continue

        d = visited.get(key)
        if d is None:
            if c:  # 건너뛴 항목: 아직 있음으로 확인
                for i in open_issues(reg, key):
                    i["last_seen"] = ts
            continue

        if not d["issues"]:
            # 상세를 못 읽음(페이지 변경·오류): 건수를 바꾸지 않아 다음에 다시 봄
            continue

        seen_ids = set()
        for it in d["issues"]:
            iid = issue_id(key, it)
            seen_ids.add(iid)
            if iid in reg["issues"]:
                rec = reg["issues"][iid]
                rec["last_seen"] = ts
                # 상태 / 조치는 바뀔 수 있음(예: 자료 제출 후) → 최신 값으로 맞춤
                for f_ in ("status", "action", "ahr_impact", "at_risk"):
                    if it.get(f_):
                        rec["page_status" if f_ == "status" else f_] = it[f_]
                # 전에 상세 사유를 못 읽었는데 이번에 읽었으면 채움
                if not rec.get("detail_ok") and it.get("detail_ok"):
                    for f_ in DETAIL_FIELDS:
                        rec[f_] = it.get(f_) or ""
                    rec["detail_ok"] = True
                if rec["status"] == GONE:  # 다시 나타남
                    rec["status"] = OPEN
                    rec.pop("gone_at", None)
                    new.append(rec)
                continue
            rec = {
                "id": iid, "cat_key": key, "cat_name": p["name"], "status": OPEN,
                "first_seen": ts, "last_seen": ts,
                "asin": it.get("asin") or "", "sku": it.get("sku") or "", "title": it.get("title") or "",
                "reason": it.get("reason") or "", "date": it.get("date") or "",
                "page_status": it.get("status") or "",
                "at_risk": it.get("at_risk") or "", "action": it.get("action") or "",
                "ahr_impact": it.get("ahr_impact") or "",
                **{f_: it.get(f_) or "" for f_ in DETAIL_FIELDS},
                "detail_ok": bool(it.get("detail_ok")),
                "raw": " | ".join(it.get("cells") or it.get("context") or []),
                "page_file": d.get("text_file"),
            }
            reg["issues"][iid] = rec
            new.append(rec)

        for i in open_issues(reg, key):
            if i["id"] not in seen_ids:
                i["status"], i["gone_at"] = GONE, ts
                gone.append(i)

        cat["last_count"], cat["checked_at"] = c, ts

    return new, gone


def asins_without_product(reg: dict, asins: List[str]) -> List[str]:
    return [a for a in dict.fromkeys(asins) if a and a not in reg["products"]]


def add_products(reg: dict, products: List[dict], now: datetime) -> None:
    ts = now.isoformat(timespec="minutes")
    for pr in products:
        if pr.get("error") and not pr.get("title"):
            continue  # 못 읽었으면 기록하지 않음, 다음에 새 이슈가 있을 때 다시 시도
        reg["products"][pr["asin"]] = {**pr, "collected_at": ts}


# ================================================================ 대장 Excel

def row_height(text: str, chars_per_line: int = 55, min_h: float = 42) -> float:
    """글자 길이로 Excel 행 높이를 정해 긴 설명도 다 보이게(Excel 행 높이 최대 409)."""
    lines = sum(max(1, -(-len(part) // chars_per_line)) for part in (text or "").split("\n"))
    return max(min_h, min(409, lines * 14 + 8))


def detail_of(i: dict) -> str:
    """Review details와 사유 상세 페이지의 설명을 합침."""
    policy = ""
    if i.get("policy_text"):
        policy = f"【정책 설명: {i.get('policy_title') or ''}】\n{i['policy_text']}"
        if i.get("policy_url"):
            policy += f"\n(원문: {i['policy_url']})"
    parts = [x for x in (i.get("review_details"), i.get("detail_text"), policy) if x]
    text = "\n\n".join(parts) or (i.get("raw") or "")
    return text[:3000]


def write_xlsx(reg: dict, out: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FONT = "Malgun Gothic"
    RED, GREY, GREEN = "C01C1C", "6B6B6B", "1E7B34"
    HEAD = PatternFill("solid", fgColor="DCE6F1")
    thin = Side(style="thin", color="BFBFBF")
    BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

    def sheet(wb, title, cols, widths, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.freeze_panes = "A2"
        for i, (h, w) in enumerate(zip(cols, widths), start=1):
            c = ws.cell(row=1, column=i, value=h)
            c.font, c.fill, c.border = Font(name=FONT, bold=True), HEAD, BOX
            c.alignment = Alignment(vertical="center", horizontal="center")
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        return ws

    def row(ws, r, values, colors=None, link_col=None):
        for i, v in enumerate(values, start=1):
            c = ws.cell(row=r, column=i, value=v)
            color = (colors or {}).get(i)
            c.font = Font(name=FONT, size=10, color=color, bold=bool(color and color == RED))
            c.border = BOX
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if link_col == i and v:
                c.hyperlink = v
                c.font = Font(name=FONT, size=10, color="0563C1", underline="single")
        ws.row_dimensions[r].height = 42

    fmt = lambda s: (s or "").replace("T", " ")

    wb = Workbook()
    ws = sheet(wb, "이슈 대장",
               ["상태", "분류", "이슈 사유", "셀러 센트럴 날짜", "ASIN", "SKU", "상품명", "판매 위험", "취한 조치",
                "점수 영향", "셀러 센트럴 상태", "상세 설명", "최초 발견", "마지막 확인", "사라진 시간"],
               [9, 16, 34, 13, 13, 15, 36, 20, 26, 12, 22, 60, 16, 16, 16], first=True)
    issues = sorted(reg["issues"].values(),
                    key=lambda i: (i["status"] != OPEN, i["cat_name"], i["first_seen"]))
    for r, i in enumerate(issues, start=2):
        title = i["title"] or reg["products"].get(i["asin"], {}).get("title", "")
        row(ws, r, [i["status"], i["cat_name"], i["reason"], i["date"], i["asin"], i["sku"], title,
                    i.get("at_risk", ""), i.get("action", ""), i.get("ahr_impact", ""), i["page_status"],
                    detail_of(i), fmt(i["first_seen"]), fmt(i["last_seen"]), fmt(i.get("gone_at"))],
            colors={1: RED if i["status"] == OPEN else GREEN})
        ws.row_dimensions[r].height = row_height(detail_of(i), chars_per_line=55)

    ps = sheet(wb, "상품 정보",
               ["ASIN", "상품 제목", "브랜드", "가격", "평점", "리뷰 수", "구매 상태", "판매 순위",
                "미해결 이슈", "수집 시간", "상품 링크"],
               [13, 48, 16, 10, 18, 14, 22, 40, 10, 16, 32])
    for r, (asin, pr) in enumerate(sorted(reg["products"].items()), start=2):
        n_open = sum(1 for i in open_issues(reg) if i["asin"] == asin)
        row(ps, r, [asin, pr.get("title", ""), pr.get("brand", ""), pr.get("price", ""), pr.get("rating", ""),
                    pr.get("reviews", ""), pr.get("availability", ""), pr.get("bsr", ""), n_open,
                    fmt(pr.get("collected_at")), pr.get("url", "")],
            colors={9: RED} if n_open else None, link_col=11)

    note_r = max(len(issues), 1) + 3
    c = ws.cell(row=note_r, column=1,
                value="이슈마다 한 번만 기록합니다. ‘마지막 확인’은 계정 상태 페이지에서 이슈가 아직 있는 것을 마지막으로 확인한 시간입니다. "
                      " ‘사라짐’은 셀러 센트럴 목록에 이 이슈가 더 이상 없다는 뜻입니다(보통 해결되었거나 삭제된 경우). ")
    c.font = Font(name=FONT, size=9, color=GREY)

    wb.save(out)

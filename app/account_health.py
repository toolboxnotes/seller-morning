"""
계정 상태 점검

흐름:
  1. 저장된 로그인으로 Seller Central 열기
  2. 사람처럼 Account Health 페이지로 이동해 한 번 둘러봄
  3. 계정 상태 등급(AHR), 고객 서비스 / 배송 지표, 정책 준수 항목별 건수 읽기
  4. 지난번 결과와 비교해 주의할 사항 정리
  5. reports/account_health/ 에 Excel 보고서 저장(오늘 보고서 + 기록 두 시트)

단독 실행:  python account_health.py
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import ah_issues
import ah_registry
import human
from sc_browser import open_session

# 계정 상태 페이지(2026년부터 /performance/dashboard, 예전 주소 /performance/account/health는 여기로 이동)
AH_URL = "https://sellercentral.amazon.com/performance/dashboard"
AH_URL_PARTS = ("/performance/dashboard", "/performance/account/health")
REPORT_DIR = Path(__file__).resolve().parent.parent / "reports" / "account_health"
HISTORY_FILE = REPORT_DIR / "history.json"
REGISTRY_FILE = REPORT_DIR / "issue_registry.json"   # 이슈 대장(이슈마다 한 번만 기록)
REGISTRY_XLSX = REPORT_DIR / "issue_registry.xlsx"

# 화면 언어가 영어 / 중국어 / 한국어일 수 있어 항목마다 세 가지 표기를 모두 적음

# 계정 상태 등급 라벨 — 페이지 데이터가 로드됐는지 판단할 때도 사용
AHR_LABELS = ["Account Health Rating", "账户状况评级", "账户健康评级", "어카운트 헬스 등급"]

# 고객 서비스 / 배송 지표:  key, 이름, 페이지에 나올 수 있는 라벨, 비교 방식, 목표값(%)
METRICS = [
    ("odr", "주문 결함률 (ODR)", ["Order Defect Rate", "订单缺陷率", "주문 하자 비율"], "<", 1.0),
    ("lsr", "배송 지연율", ["Late Shipment Rate", "迟发率", "지연된 배송 비율"], "<", 4.0),
    ("pfcr", "출고 전 취소율", ["Pre-fulfillment Cancel Rate", "配送前取消率", "사전 이행 취소 비율"], "<", 2.5),
    ("vtr", "유효 추적률", ["Valid Tracking Rate", "有效追踪率", "유효 추적율"], ">", 95.0),
    ("otdr", "정시 배송률", ["On-Time Delivery Rate", "准时交货率", "적시 배송율"], ">", 90.0),
    ("bhdr", "영업시간 내 배송률", ["Business Hour Delivery Rate", "营业时间送达率", "영업시간 배송률"], ">", 90.0),
]

# 정책 준수 항목:  key, 이름, 페이지에 나올 수 있는 라벨
POLICIES = [
    ("ip_suspected", "지식재산권 침해 의심",
     ["Suspected Intellectual Property Violations", "涉嫌侵犯知识产权", "지적 재산권 위반으로 의심되는 사항"]),
    ("ip_complaints", "지식재산권 관련 불만",
     ["Received Intellectual Property Complaints", "收到的知识产权投诉", "지적 재산권 관련 불만"]),
    ("authenticity", "상품 진위 관련 고객 불만",
     ["Product Authenticity Customer Complaints", "商品真实性买家投诉", "제품 신뢰성 고객 이의"]),
    ("condition", "상품 상태 관련 고객 불만",
     ["Product Condition Customer Complaints", "商品状况买家投诉", "제품 상태 고객 불만 사항"]),
    ("safety", "식품 및 상품 안전 문제",
     ["Food and Product Safety Issues", "食品和商品安全问题", "식품 및 상품 안전 문제"]),
    ("listing", "리스팅 정책 위반",
     ["Listing Policy Violations", "违反上架政策", "上架政策违规", "목록 정책 위반"]),
    ("restricted", "제한 상품 정책 위반",
     ["Restricted Product Policy Violations", "违反受限商品政策", "금지 제품 정책 위반"]),
    ("reviews", "고객 상품 리뷰 정책 위반",
     ["Customer Product Reviews Policy Violations", "违反买家商品评论政策", "고객 상품 리뷰 정책 위반"]),
    ("other", "기타 정책 위반",
     ["Other Policy Violations", "其他违反政策", "기타 정책 위반 사항"]),
    ("regulatory", "규정 준수", ["Regulatory Compliance", "监管合规", "규정 준수"]),
]

# 페이지에서 이 단어가 들어간 문장은 확인용으로 보고서에 옮겨 적음
NOTICE_KEYWORDS = ["action required", "take action", "at risk", "deactivat", "suspend",
                   "需要采取", "请采取", "有风险", "停用", "警告",
                   "조치 필요", "조치를 취", "위험", "비활성화", "정지"]
# 그중 아래는 ‘반드시 처리’ — 주의 사항에 넣음
STRONG_NOTICE = ["action required", "需要采取", "조치 필요"]
# 고정 안내 문구(예: Account Health Assurance 소개의 “won't deactivate”) — 알림으로 보지 않음
NOTICE_EXCLUDE = ["account health assurance", "you're enrolled", "won't deactivate", "will not deactivate",
                  "어카운트 헬스 보증", "등록되었습니다", "账户状况保障"]

ALL_LABELS = [l for _, _, ls, *_ in METRICS for l in ls] + [l for _, _, ls in POLICIES for l in ls]
TARGET_WORDS = ("target", "目标", "대상", "목표")
NA_PATTERN = re.compile(r"N/?A|--|不适用|无数据|해당 사항 없음", re.I)


# ================================================================ 페이지 조작

def _is_ah_url(url: str) -> bool:
    return any(p in url for p in AH_URL_PARTS)


def _go_to_account_health(page) -> None:
    """먼저 사람처럼 페이지의 Account Health 메뉴를 클릭하고, 못 찾으면 주소로 바로 엶."""
    selector = ", ".join(f'a[href*="{p}"]' for p in AH_URL_PARTS)
    links = page.locator(selector)
    visible = None
    for i in range(min(links.count(), 10)):
        if links.nth(i).is_visible():
            visible = links.nth(i)
            break

    if visible is not None:
        print("   ‘계정 상태’ 진입 버튼 클릭...")
        human.click(page, visible, wait_after=False)
        try:
            page.wait_for_url(_is_ah_url, timeout=20000)
        except Exception:
            human.goto(page, AH_URL, wait_after=False)
    else:
        print("   계정 상태 페이지 여는 중...")
        human.goto(page, AH_URL, wait_after=False)

    print("   데이터 로딩 대기 중...")
    pattern = "|".join(re.escape(l) for l in AHR_LABELS)
    try:
        page.wait_for_function(
            f"() => /{pattern}/i.test(document.body.innerText)", timeout=45000
        )
    except Exception:
        print("   ⚠ 45초 안에 점수 영역이 보이지 않아, 현재 내용으로 계속 읽습니다. ")
    human.pause()


# ================================================================ 분석

def _lines(text: str) -> List[str]:
    return [l.strip() for l in text.splitlines() if l.strip()]


def _label_index(lines: List[str], labels: List[str]) -> Optional[int]:
    """줄 전체가 라벨인 곳을 먼저 찾고(가장 정확), 없으면 라벨이 들어 있는 줄을 찾음."""
    low_labels = [lab.lower() for lab in labels]
    for i, line in enumerate(lines):
        if line.lower() in low_labels:
            return i
    for i, line in enumerate(lines):
        low = line.lower()
        if any(lab in low for lab in low_labels):
            return i
    return None


def _is_other_label(line: str, own: List[str]) -> bool:
    low = line.lower()
    return any(l.lower() == low for l in ALL_LABELS if l not in own)


def parse_ahr(lines: List[str]) -> dict:
    i = _label_index(lines, AHR_LABELS)
    if i is None:
        return {"score": None, "status": None}
    score = None
    for s in lines[i:i + 12]:
        m = re.fullmatch(r"(\d{1,4})(\s*/\s*1000)?", s)
        if m and int(m.group(1)) <= 1000:
            score = int(m.group(1))
            break
    return {"score": score, "status": ahr_status(score)}


def ahr_status(score: Optional[int]) -> Optional[str]:
    if score is None:
        return None
    if score >= 200:
        return "건강"
    if score >= 100:
        return "위험"
    return "위급"


def parse_metric(lines: List[str], labels: List[str], window: int = 6) -> dict:
    """
    라벨 뒤에서 값(백분율 또는 N/A)을 찾음. “Target: …” 줄은 건너뛰고,
    값이 아닌 첫 줄이 나오면 멈춤(아래 세부 지표 값까지 가져오지 않게).
    주문 결함률은 ‘판매자 배송 / Amazon 배송’ 두 열이라 값이 두 개.
    """
    i = _label_index(lines, labels)
    if i is None:
        return {"found": False, "values": [], "raw": []}
    values, raw = [], []
    for s in lines[i + 1:i + 1 + window]:
        if any(w in s.lower() for w in TARGET_WORDS):
            continue
        if NA_PATTERN.fullmatch(s):
            raw.append("N/A")
            continue
        m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%", s)
        if m:
            values.append(float(m.group(1)))
            raw.append(s)
            continue
        if raw or _is_other_label(s, labels):
            break
    return {"found": True, "values": values, "raw": raw}


def parse_policy(lines: List[str], labels: List[str], window: int = 3) -> Optional[int]:
    i = _label_index(lines, labels)
    if i is None:
        return None
    for s in lines[i + 1:i + 1 + window]:
        if _is_other_label(s, labels):
            break
        if re.fullmatch(r"\d+", s):
            return int(s)
    return None


def parse_notices(lines: List[str]) -> List[str]:
    found, seen = [], set()
    for s in lines:
        low = s.lower()
        if not (10 <= len(s) <= 300) or not any(k in low for k in NOTICE_KEYWORDS):
            continue
        if any(x in low for x in NOTICE_EXCLUDE):
            continue
        if re.search(r"\d+\s*[-–]\s*\d+", s):  # 등급 구간 설명, 예: “100-199 At Risk”
            continue
        if s not in seen:
            seen.add(s)
            found.append(s)
    return found[:15]


def metric_text(m: dict) -> str:
    """지표 값을 읽기 쉬운 글로. 주문 결함률 두 열은 판매자 배송과 FBA를 각각 표시."""
    if not m["found"]:
        return "확인 불가"
    if not m["raw"]:
        return "데이터 없음"
    if m["key"] == "odr" and len(m["raw"]) == 2:
        return f"판매자 배송 {m['raw'][0]} · FBA {m['raw'][1]}"
    return " / ".join(m["raw"])


def parse_page(text: str) -> dict:
    lines = _lines(text)
    result = {"ahr": parse_ahr(lines), "metrics": [], "policies": [], "notices": parse_notices(lines)}

    for key, name, labels, op, target in METRICS:
        m = parse_metric(lines, labels)
        ok = None
        if m["values"]:
            ok = all(v < target for v in m["values"]) if op == "<" else all(v > target for v in m["values"])
        result["metrics"].append({
            "key": key, "name": name, "found": m["found"], "values": m["values"],
            "raw": m["raw"], "target": f"{target:g}% {'미만' if op == '<' else '초과'}", "ok": ok,
        })

    for key, name, labels in POLICIES:
        result["policies"].append({"key": key, "name": name, "count": parse_policy(lines, labels)})

    return result


# ================================================================ 비교와 주의 사항

_OLD_STATUS = {"健康": "건강", "有风险": "위험", "危急": "위급"}   # 예전 버전은 중국어로 저장했음


def load_history() -> list:
    try:
        history = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except Exception:
        return []
    for h in history:
        ahr = h.get("ahr") or {}
        if ahr.get("status") in _OLD_STATUS:
            ahr["status"] = _OLD_STATUS[ahr["status"]]
    return history


def save_history(history: list) -> None:
    HISTORY_FILE.write_text(json.dumps(history[-365:], ensure_ascii=False, indent=2), encoding="utf-8")


def build_attention(cur: dict, prev: Optional[dict]) -> List[str]:
    items = []
    ahr = cur["ahr"]
    if ahr["score"] is None:
        items.append("계정 상태 점수를 확인하지 못했습니다. 보고서의 페이지 스크린샷을 확인해 주세요. ")
    elif ahr["status"] != "건강":
        items.append(f"계정 상태 점수가 {ahr['score']}({ahr['status']})로, 건강 기준선 200 미만입니다. 빠르게 조치해야 합니다. ")

    if prev and ahr["score"] is not None and prev["ahr"].get("score") is not None:
        diff = ahr["score"] - prev["ahr"]["score"]
        if diff < 0:
            items.append(f"계정 상태 점수가 지난번({prev['time'][:16].replace('T', ' ')})보다 {-diff}점 하락했습니다. ")

    for m in cur["metrics"]:
        if m["ok"] is False:
            items.append(f"{m['name']} 값: {metric_text(m)} — 목표 미달 (목표 {m['target']}). ")

    # 정책 준수: ‘새로 발견 / 변화 있음 / 상세를 못 읽음’ 항목만 알림, 이미 봤고 변화 없는 항목은 반복하지 않음
    new_by_cat, gone_by_cat = {}, {}
    for i in cur.get("new_issues", []):
        new_by_cat.setdefault(i["cat_key"], []).append(i)
    for i in cur.get("gone_issues", []):
        gone_by_cat.setdefault(i["cat_key"], []).append(i)
    failed = {d["key"]: d for d in cur.get("details", []) if not d["issues"]}
    skipped = set(cur.get("skipped", []))

    for p in cur["policies"]:
        key, c = p["key"], p["count"]
        if key in failed:
            why = failed[key].get("error") or "구체적인 이슈를 확인하지 못했습니다"
            items.append(f"정책 준수 · {p['name']}: {c}건, 상세 내용을 읽지 못했습니다({why}). 셀러 센트럴에서 확인해 주세요. ")
            continue
        new = new_by_cat.get(key, [])
        if new:
            asins = list(dict.fromkeys(i["asin"] for i in new if i["asin"]))
            shown = ", ".join(asins[:6]) + (" 등" if len(asins) > 6 else "")
            asin_txt = f", 관련 ASIN: {shown}" if asins else ""
            items.append(f"정책 준수 · {p['name']}: 새 이슈 {len(new)}건{asin_txt}, 자세한 내용은 ‘새 이슈’ 시트를 참고하세요. ")
        gone = gone_by_cat.get(key, [])
        if gone:
            items.append(f"정책 준수 · {p['name']}: 이슈 {len(gone)}건이 셀러 센트럴 목록에서 사라졌습니다(해결된 것으로 보임). 대장에 표시했습니다. ")
        prev_count = {p_["key"]: p_["count"] for p_ in (prev or {}).get("policies", [])}.get(key)
        if c and not new and not gone and key not in skipped and key not in failed and prev_count != c:
            items.append(f"정책 준수 · {p['name']}: {c}건, 셀러 센트럴에서 상세 내용을 확인해 주세요. ")

    for n in cur["notices"]:
        if any(k in n.lower() for k in STRONG_NOTICE):
            items.append(f"페이지 안내: {n}")

    return items


# ================================================================ Excel 보고서

def write_xlsx(cur: dict, prev: Optional[dict], attention: List[str], history: list,
               screenshot: Path, raw_txt: Path, out: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FONT = "Malgun Gothic"
    GREEN, ORANGE, RED, GREY = "1E7B34", "C26A00", "C01C1C", "6B6B6B"
    HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
    thin = Side(style="thin", color="BFBFBF")
    BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

    def f(size=10.5, bold=False, color=None):
        return Font(name=FONT, size=size, bold=bold, color=color)

    def put(ws, row, col, value, font=None, fill=None, border=None, wrap=False, num=None, align="left"):
        c = ws.cell(row=row, column=col, value=value)
        c.font = font or f()
        if fill:
            c.fill = fill
        if border:
            c.border = border
        c.alignment = Alignment(vertical="center", horizontal=align, wrap_text=wrap)
        if num:
            c.number_format = num
        return c

    def header(ws, row, headers):
        for i, h in enumerate(headers, start=1):
            put(ws, row, i, h, f(bold=True), HEAD_FILL, BOX)

    wb = Workbook()

    # ================================================= 오늘 보고서
    ws = wb.active
    ws.title = "오늘 보고서"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", (30, 22, 16, 12)):
        ws.column_dimensions[col].width = w

    when = datetime.fromisoformat(cur["time"])
    put(ws, 1, 1, "계정 상태 일일 보고서", f(18, True, "1F3864"))
    put(ws, 2, 1, f"{when:%Y-%m-%d %H:%M} · Amazon.com(미국)", f(9.5, color=GREY))

    r = 4
    put(ws, r, 1, "오늘의 결론", f(13, True, "1F3864"))
    r += 1
    ahr = cur["ahr"]
    put(ws, r, 1, "계정 상태 점수", f(bold=True))
    if ahr["score"] is None:
        put(ws, r, 2, "확인 불가", f(12, True, GREY))
    else:
        color = {"건강": GREEN, "위험": ORANGE, "위급": RED}[ahr["status"]]
        put(ws, r, 2, ahr["score"], f(14, True, color))
        put(ws, r, 3, ahr["status"], f(12, True, color))
    r += 1
    put(ws, r, 1, "지난번 대비", f(bold=True))
    if prev and prev["ahr"].get("score") is not None and ahr["score"] is not None:
        diff = ahr["score"] - prev["ahr"]["score"]
        put(ws, r, 2, diff, f(color=RED if diff < 0 else GREY), num='+0;-0;0')
        put(ws, r, 3, f"지난번 {prev['time'][:16].replace('T', ' ')}", f(9, color=GREY))
    else:
        put(ws, r, 2, "—", f(color=GREY))
    r += 1
    put(ws, r, 1, "주의 필요", f(bold=True))
    put(ws, r, 2, f"{len(attention)}개 항목" if attention else "없음 ✓", f(12, True, RED if attention else GREEN))
    r += 1
    known = sum((cur.get("known_open") or {}).values())
    if known:
        put(ws, r, 1, "알려진 미해결 이슈", f(bold=True))
        put(ws, r, 2, f"{known}건", f(bold=True, color=ORANGE))
        put(ws, r, 3, "이슈 대장에 기록되어 있어 다시 나열하지 않습니다", f(9, color=GREY))
        r += 1
    put(ws, r, 1, "점수 구간: 200–1000 건강, 100–199 위험, 0–99 위급", f(9, color=GREY))

    # ---- 주의 사항
    r += 2
    put(ws, r, 1, "주의가 필요한 사항", f(13, True, "1F3864"))
    r += 1
    if known:
        empty_msg = f"새로 주의가 필요한 사항은 없습니다. 이미 알려진 미해결 이슈 {known}건은 이전에 기록되었습니다. 이슈 대장 issue_registry.xlsx를 참고하세요. "
    else:
        empty_msg = "계정 상태 점수가 정상이며, 모든 지표가 목표 범위 안에 있고, 정책 준수 기록이 없습니다. "
    items = attention or [empty_msg]
    for i, a in enumerate(items, start=1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        put(ws, r, 1, f"{i}. {a}" if attention else a, f(color=None if attention else GREEN), wrap=True)
        ws.row_dimensions[r].height = 30 if len(a) > 45 else 18
        r += 1

    # ---- 성과 지표
    r += 1
    put(ws, r, 1, "고객 서비스 및 배송 성과", f(13, True, "1F3864"))
    r += 1
    header(ws, r, ["지표", "현재 값", "목표", "상태"])
    for m in cur["metrics"]:
        r += 1
        val = metric_text(m)
        if not m["values"]:
            status, color = "—", GREY
        else:
            status, color = ("정상", GREEN) if m["ok"] else ("미달", RED)
        put(ws, r, 1, m["name"], border=BOX)
        put(ws, r, 2, val, border=BOX)
        put(ws, r, 3, m["target"], border=BOX)
        put(ws, r, 4, status, f(bold=True, color=color), border=BOX)
    r += 1
    put(ws, r, 1, "배송 지표는 판매자 배송 주문만 집계합니다. 모두 FBA라면 N/A로 표시되는 것이 정상입니다. ", f(9, color=GREY))

    # ---- 정책 준수
    r += 2
    put(ws, r, 1, "정책 준수", f(13, True, "1F3864"))
    r += 1
    header(ws, r, ["분류", "수량", "지난번 대비", "상태"])
    prev_pol = {p_["key"]: p_["count"] for p_ in (prev or {}).get("policies", [])}
    for pol in cur["policies"]:
        r += 1
        c, before = pol["count"], prev_pol.get(pol["key"])
        put(ws, r, 1, pol["name"], border=BOX)
        if c is None:
            put(ws, r, 2, "확인 불가", f(color=GREY), border=BOX)
            put(ws, r, 3, "—", border=BOX)
            put(ws, r, 4, "—", f(color=GREY), border=BOX)
        else:
            put(ws, r, 2, c, border=BOX, align="center")
            if before is None:
                put(ws, r, 3, "—", border=BOX, align="center")
            else:
                put(ws, r, 3, c - before, f(color=RED if c > before else None), border=BOX, num='+0;-0;0', align="center")
            put(ws, r, 4, "정상" if c == 0 else "처리 필요", f(bold=True, color=GREEN if c == 0 else RED), border=BOX)

    # ---- 페이지 알림
    if cur["notices"]:
        r += 2
        put(ws, r, 1, "페이지 안내(자동 발췌, 확인용)", f(13, True, "1F3864"))
        for n in cur["notices"]:
            r += 1
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            put(ws, r, 1, n, wrap=True)
            ws.row_dimensions[r].height = 30 if len(n) > 60 else 18

    # ---- 설명
    r += 2
    for line in [
        "위 데이터는 스크립트가 페이지 텍스트를 자동으로 읽은 결과입니다. ‘확인 불가’는 페이지에서 해당 항목을 찾지 못했다는 뜻으로, 데이터가 없거나 페이지가 바뀌었거나, "
        " 셀러 센트럴 언어가 영어 / 중국어 / 한국어가 아닌 경우일 수 있습니다. 실제 셀러 센트럴 페이지를 기준으로 하세요. ",
        f"전체 페이지 스크린샷: {screenshot.name}   페이지 원문: {raw_txt.name}(이 파일과 같은 폴더)",
    ]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        put(ws, r, 1, line, f(9, color=GREY), wrap=True)
        ws.row_dimensions[r].height = 30
        r += 1

    # ================================================= 새 이슈(이번에 새로 발견한 것만, 이미 기록된 것은 이슈 대장)
    new_issues = cur.get("new_issues") or []
    failed = [d for d in (cur.get("details") or []) if not d["issues"]]
    products = cur.get("products") or []
    if new_issues or failed:
        ds = wb.create_sheet("새 이슈")
        ds.freeze_panes = "A2"
        cols = ["분류", "이슈 사유", "셀러 센트럴 날짜", "ASIN", "SKU", "상품명", "판매 위험", "취한 조치",
                "점수 영향", "셀러 센트럴 상태", "상세 설명"]
        widths = [16, 34, 13, 13, 15, 36, 20, 26, 12, 22, 70]
        header(ds, 1, cols)
        for i, w in enumerate(widths, start=1):
            ds.column_dimensions[get_column_letter(i)].width = w
        prod_titles = {p_["asin"]: p_.get("title") for p_ in products}
        r = 1
        for it in new_issues:
            r += 1
            title = it.get("title") or prod_titles.get(it.get("asin")) or ""
            detail = ah_registry.detail_of(it)
            vals = [it["cat_name"], it["reason"], it["date"], it["asin"], it["sku"], title, it.get("at_risk", ""),
                    it.get("action", ""), it.get("ahr_impact", ""), it["page_status"], detail]
            for c_i, v in enumerate(vals, start=1):
                put(ds, r, c_i, v, border=BOX, wrap=c_i in (2, 6, 7, 8, 10, 11))
                ds.cell(row=r, column=c_i).alignment = Alignment(vertical="top", wrap_text=c_i in (2, 6, 7, 8, 10, 11))
            ds.row_dimensions[r].height = ah_registry.row_height(detail, chars_per_line=60)
        for d in failed:
            r += 1
            note = d.get("error") or "분류 페이지에서 구체적인 이슈를 확인하지 못했습니다"
            put(ds, r, 1, d["name"], border=BOX)
            ds.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
            put(ds, r, 2, f"{note}, 스크린샷 {d.get('screenshot') or ''}에서 확인하세요", f(color=GREY), border=BOX, wrap=True)
        r += 2
        ds.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        put(ds, r, 1, "이번에 새로 발견된 이슈만 표시합니다. ‘상세 설명’은 셀러 센트럴의 Review details와 사유 상세 페이지에서 가져왔으며, 원문은 *_detail.txt로 따로 저장됩니다. 모든 이슈(이전에 기록된 것 포함)는 issue_registry.xlsx 이슈 대장에 있으며, "
                      " 이슈마다 한 번만 기록합니다. ", f(9, color=GREY), wrap=True)

    # ================================================= 새 상품 정보(ASIN마다 한 번만 수집)
    if products:
        ps = wb.create_sheet("영향 받는 상품")
        ps.freeze_panes = "B2"
        cols = ["ASIN", "상품 제목", "브랜드", "가격", "평점", "리뷰 수", "구매 상태", "판매 순위", "상품 링크", "비고"]
        widths = [13, 48, 16, 10, 18, 14, 22, 40, 32, 30]
        header(ps, 1, cols)
        for i, w in enumerate(widths, start=1):
            ps.column_dimensions[get_column_letter(i)].width = w
        for r, pr in enumerate(products, start=2):
            vals = [pr["asin"], pr.get("title", ""), pr.get("brand", ""), pr.get("price", ""), pr.get("rating", ""),
                    pr.get("reviews", ""), pr.get("availability", ""), pr.get("bsr", ""), pr["url"], pr.get("error") or ""]
            for c_i, v in enumerate(vals, start=1):
                c = put(ps, r, c_i, v, f(color="0563C1") if c_i == 9 else (f(color=RED) if c_i == 10 and v else None),
                        border=BOX, wrap=c_i in (2, 7, 8, 10))
                if c_i == 9:
                    c.hyperlink = v
            ps.row_dimensions[r].height = 45
        r = len(products) + 3
        ps.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        put(ps, r, 1, "이번에 처음 수집한 상품만 표시합니다. 상품 정보는 Amazon.com 상품 페이지에서 읽었으며, 가격과 구매 상태는 배송지에 따라 다를 수 있습니다. "
                      " 상품이 판매 중지된 경우 페이지에 재고 없음으로 표시되거나 열리지 않을 수 있습니다. ", f(9, color=GREY), wrap=True)

    # ================================================= 기록
    hs = wb.create_sheet("기록")
    cols = (["시간", "점수", "지난번 대비", "상태"]
            + [m[1] for m in METRICS]
            + [p[1] for p in POLICIES]
            + ["주의 필요(항목)"])
    header(hs, 1, cols)
    hs.freeze_panes = "B2"
    for i, h in enumerate(cols, start=1):
        hs.column_dimensions[get_column_letter(i)].width = 18 if i == 1 else max(10, len(h) * 2 + 2)
    hs.row_dimensions[1].height = 32
    for c in hs[1]:
        c.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)

    for i, h in enumerate(history, start=2):
        # 시각을 글자로 적어야 추세 그래프가 같은 날 여러 번 점검을 따로 보여 줌
        put(hs, i, 1, h["time"][:16].replace("T", " "), border=BOX)
        score = h["ahr"].get("score")
        put(hs, i, 2, score, border=BOX, align="center")
        # 지난번 대비: 수식으로 계산, 두 번 모두 등급이 있을 때만 표시
        diff = "" if i == 2 else f'=IF(AND(ISNUMBER(B{i}),ISNUMBER(B{i - 1})),B{i}-B{i - 1},"")'
        put(hs, i, 3, diff, border=BOX, num='+0;-0;0', align="center")
        st = h["ahr"].get("status") or "확인 불가"
        put(hs, i, 4, st, f(color={"건강": GREEN, "위험": ORANGE, "위급": RED}.get(st, GREY)), border=BOX)
        col = 5
        mets = {m["key"]: m for m in h.get("metrics", [])}
        for key, *_ in METRICS:
            m = mets.get(key)
            text = metric_text(m) if m else "—"
            put(hs, i, col, text, f(color=RED if m and m.get("ok") is False else None), border=BOX)
            col += 1
        pols = {p_["key"]: p_["count"] for p_ in h.get("policies", [])}
        for key, *_ in POLICIES:
            c = pols.get(key)
            put(hs, i, col, c, f(color=RED if c else None), border=BOX)
            col += 1
        n_att = len(h.get("attention", []))
        put(hs, i, col, n_att, f(bold=bool(n_att), color=RED if n_att else None), border=BOX)

    # 등급 추세 그래프(기록이 두 번 이상일 때)
    last = len(history) + 1
    if len(history) >= 2:
        ch = LineChart()
        ch.title = "계정 상태 점수 추이"
        ch.y_axis.title = "평점"
        ch.height, ch.width = 7, 16
        ch.legend = None
        ch.add_data(Reference(hs, min_col=2, min_row=1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(hs, min_col=1, min_row=2, max_row=last))
        for sr in ch.series:
            sr.smooth = False
            sr.marker.symbol = "circle"
        hs.add_chart(ch, f"A{last + 3}")

    # 인쇄: 가로 방향, 페이지 너비에 맞춤
    for sheet in wb.worksheets:
        if sheet.title in ("새 이슈", "영향 받는 상품", "기록"):
            sheet.page_setup.orientation = "landscape"
            sheet.sheet_properties.pageSetUpPr.fitToPage = True
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.fitToHeight = 0

    wb.save(out)


# ================================================================ 메인 흐름

def run(open_report: bool = True) -> Path:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    stamp = f"{now:%Y%m%d_%H%M}"
    shot = REPORT_DIR / f"account_health_{stamp}.png"
    raw_txt = REPORT_DIR / f"account_health_{stamp}_page.txt"
    xlsx_path = REPORT_DIR / f"account_health_{stamp}.xlsx"

    with open_session(headless=False) as page:
        _go_to_account_health(page)
        print("   페이지 살펴보는 중...")
        human.read_page(page)
        text = page.inner_text("body")
        raw_txt.write_text(f"URL: {page.url}\n\n{text}", encoding="utf-8")
        page.screenshot(path=str(shot), full_page=True)
        cur = parse_page(text)

        # ---- 이슈 상세: 건수가 바뀐 항목만, 상품 정보는 ASIN마다 한 번만 수집
        registry = ah_registry.load(REGISTRY_FILE)
        to_visit, skipped = ah_registry.plan(registry, cur["policies"])
        for p_ in skipped:
            print(f"   “{p_['name']}” {p_['count']}건, 대장과 같아 이미 확인했으므로 건너뜁니다. ")
        details, products = [], []
        if to_visit:
            label_map = {k: labels for k, _, labels in POLICIES}
            details = ah_issues.collect_issue_details(
                page, to_visit, label_map, REPORT_DIR / f"account_health_{stamp}",
                is_known=lambda k, it: registry["issues"].get(ah_registry.issue_id(k, it), {}).get("detail_ok", False))
        new_issues, gone_issues = ah_registry.apply(registry, cur["policies"], details, now)
        need = ah_registry.asins_without_product(registry, [i["asin"] for i in new_issues])
        if need:
            products = ah_issues.collect_product_info(page, need)
            ah_registry.add_products(registry, products, now)

        # ---- 클레임(차지백·A-to-z·SAFE-T)과 고객의 소리(VoC): 페이지만 읽음
        extra = []
        try:
            import sc_pages
            print("   클레임·고객의 소리 확인...")
            import sp_api
            import sp_data
            perf = sp_api.try_api("계정 성과 지표(ODR·A-to-z·차지백 등)", sp_data.performance)
            if perf:
                perf.update({"name": "계정 성과 지표(API)", "src": "SP-API", "time": datetime.now().isoformat(timespec="seconds")})
                extra = [perf, sc_pages.read(page, "voc")]
            else:
                extra = [sc_pages.read(page, "claims"), sc_pages.read(page, "voc")]
        except Exception as e:
            print(f"   ⚠ 클레임·고객의 소리를 읽지 못했습니다: {type(e).__name__}")

    cur["time"] = now.isoformat(timespec="seconds")
    cur["details"], cur["products"] = details, products
    cur["new_issues"], cur["gone_issues"] = new_issues, gone_issues
    cur["known_open"] = {p_["key"]: len(ah_registry.open_issues(registry, p_["key"])) for p_ in cur["policies"]}
    cur["skipped"] = [p_["key"] for p_ in skipped]
    cur["issue_asins"] = {}
    for i in ah_registry.open_issues(registry):
        cur["issue_asins"].setdefault(i["cat_key"], [])
        if i["asin"] and i["asin"] not in cur["issue_asins"][i["cat_key"]]:
            cur["issue_asins"][i["cat_key"]].append(i["asin"])

    history = load_history()
    prev = history[-1] if history else None
    attention = build_attention(cur, prev)
    cur["attention"] = attention

    # 기록에는 요약만, 상세는 그때의 Excel과 이슈 대장에
    history.append({k: v for k, v in cur.items()
                    if k not in ("details", "products", "new_issues", "gone_issues")})
    write_xlsx(cur, prev, attention, history, shot, raw_txt, xlsx_path)
    try:
        import sc_pages
        sc_pages.append_sheets(xlsx_path, extra)
    except Exception as e:
        print(f"   ⚠ 클레임·고객의 소리 시트를 추가하지 못했습니다: {type(e).__name__}")
    try:
        import daily_bus
        ahr_ = cur["ahr"]
        lv = "P0" if ahr_.get("status") not in ("건강", None) else "P1"
        daily_bus.record("계정 상태", {"계정 상태 점수": ahr_.get("score"), "상태": ahr_.get("status")},
                         [(lv, "계정 상태", a, "계정 상태 보고서 확인") for a in attention],
                         xlsx_path, [a for r in extra if r for a in r.get("alerts", [])])
        km = {}
        for r in extra:
            km.update({f"{r.get('name', '')}: {k}": v for k, v in (r or {}).get("kpis", {}).items()
                       if not isinstance(v, (list, dict))})
        if km:
            d = daily_bus.load()
            d["modules"]["계정 상태"]["kpis"].update(km)
            daily_bus._save(d)
    except Exception:
        pass
    save_history(history)  # 보고서를 만든 뒤에만 기록에 남김
    ah_registry.save(registry, REGISTRY_FILE)
    try:
        ah_registry.write_xlsx(registry, REGISTRY_XLSX)
    except PermissionError:
        print(f"⚠ {REGISTRY_XLSX.name} 파일이 Excel에서 열려 있어 이번에는 대장 표를 업데이트하지 못했습니다(데이터는 저장되었으며 다음에 함께 업데이트됩니다). ")

    # ---- 터미널 요약
    ahr = cur["ahr"]
    print()
    if ahr["score"] is None:
        print(f"계정 상태 점수: 확인 불가(스크린샷 확인: {shot.name})")
    else:
        print(f"계정 상태 점수: {ahr['score']}({ahr['status']})")
    if attention:
        print(f"주의 필요 {len(attention)}개 항목: ")
        for a in attention:
            print("  •", a)
    else:
        print("오늘은 주의가 필요한 사항이 없습니다 ✓")
    print(f"\n📊 보고서 저장 완료: {xlsx_path}")
    print(f"📒 이슈 대장: {REGISTRY_XLSX}")

    if open_report:
        try:
            __import__("platform_util").open_path(xlsx_path)  # Excel로 열기
        except Exception:
            pass
    return xlsx_path


if __name__ == "__main__":
    run()

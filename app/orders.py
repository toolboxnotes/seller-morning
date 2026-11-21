"""
주문 보고서

기간 선택: 1 전날 / 2 최근 7일 / 3 최근 30일
  - 날짜는 미국 중부 시간 기준(주문 시각도 중부 시간으로 바꿔 날짜에 넣음)
  - 비교용으로 바로 앞 같은 길이 기간도 받음: 최근 7일 = 이번 7일 + 이전 7일
  - 하루 한 번만 받음: 오늘(중부 시간) 이미 받았고 필요한 날짜를 덮는 데이터가 있으면 그대로 사용

데이터 출처:
  - 주문: 메뉴 주문 → 주문 보고서 → All Orders(주문마다: 주문 번호, 주문 시각, SKU, ASIN, 수량, 금액, 상태, 배송 방식, 주)
    (보고서 센터의 예전 All Orders 페이지는 Amazon이 없애서 버튼을 눌러도 다운로드되지 않음)
  - 반품: 보고서 센터 Customer Concessions → FBA customer returns(반품률 계산용, 못 받으면 보고서에 적고 다른 내용엔 영향 없음)

파일은 reports/orders/ 에 저장:
  orders_data_시작_끝_받은날.xlsx    주문 데이터
  returns_data_시작_끝_받은날.xlsx   반품 데이터
  orders_report_기간_날짜.xlsx          주문 보고서(비교 분석)
  last_run.json                          기간별 마지막 실행 시각
  raw/ debug/                            원본 다운로드 파일 / 오류 때 스크린샷과 페이지 구조
"""

import json
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import human
import inventory as inv
import sc_nav
import sc_reports
from sc_browser import open_session

REPORT_DIR = Path(__file__).resolve().parent.parent / "reports" / "orders"
RAW_DIR = REPORT_DIR / "raw"
DEBUG_DIR = REPORT_DIR / "debug"
LAST_RUN_FILE = REPORT_DIR / "last_run.json"

PERIODS = {
    "1": ("yesterday", "전날", 1),
    "2": ("7d", "최근 7일", 7),
    "3": ("14d", "최근 14일", 14),
    "4": ("30d", "최근 30일", 30),
}
# 관리 화면의 날짜 드롭다운은 미리 정한 범위만 고를 수 있음(날짜 직접 입력 불가), 최대 30일.
# 하루 한 번 “Last 30 days”만 요청해 세 옵션 모두 이 데이터를 씀: 7일·14일은 이전 기간과 비교 가능, 30일은 이전 기간 데이터가 없어 비교하지 않음.
FETCH_DAYS = 30
PRESET_30 = re.compile(r"^\s*(Date range:\s*)?Last 30 days\s*$", re.I)


def has_compare(days: int) -> bool:
    return days * 2 <= FETCH_DAYS

# ---------------------------------------------------------------- 평가 기준(조정 가능)
HANDLING_DAYS = 2           # 판매자 배송 준비 기간(일): 넘으면 지연, 1일도 안 남으면 곧 지연
CANCEL_RATE_ALERT = 0.05    # 취소율 5% 넘으면 알림
RETURN_RATE_ALERT = 0.10    # 반품률 10% 넘으면 알림
MIN_UNITS_FOR_RATE = 5      # 건수가 너무 적으면 비율 알림을 하지 않음
MAX_CHUNK_DAYS = 29         # 한 번 요청에 담을 최대 일수(주문 보고서는 최대 30일, 시작일은 하루 앞당김)

# ---------------------------------------------------------------- 페이지 글자
ALL_ORDERS = ["All Orders", "모든 주문", "所有订单"]
RETURNS = ["FBA customer returns", "FBA Customer Returns", "FBA 고객 반품", "亚马逊物流买家退货"]
EXACT_DATES = re.compile(r"^\s*(Exact dates?|Custom|Custom date range|정확한 날짜|사용자 지정|指定日期|自定义)\s*$", re.I)
FROM_LABELS = re.compile(r"^\s*(From|Start date|시작|시작 날짜|从|开始日期)\s*:?\s*$", re.I)
TO_LABELS = re.compile(r"^\s*(To|End date|종료|종료 날짜|至|到|结束日期)\s*:?\s*$", re.I)
PRESETS = [(1, ["Last day", "Yesterday"]), (3, ["Last 3 days"]), (7, ["Last 7 days", "Last week"]),
           (14, ["Last 14 days"]), (30, ["Last 30 days", "Last month"]), (60, ["Last 60 days"]),
           (90, ["Last 90 days"]), (180, ["Last 180 days"]), (365, ["Last 365 days", "Last year"])]


# ================================================================ 미국 중부 시간

def central_now() -> datetime:
    utc = datetime.now(timezone.utc).replace(tzinfo=None)
    return utc + timedelta(hours=sc_reports._us_offset(utc, -6))


def central_today() -> date:
    return central_now().date()


def to_central(value) -> Optional[datetime]:
    """주문 보고서의 시각(예: 2026-10-25T14:22:11+00:00) → 미국 중부 시간."""
    if value in (None, ""):
        return None
    s = str(value).strip().replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        m = re.match(r"(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})", s)
        if not m:
            return None
        dt = datetime.fromisoformat(f"{m.group(1)}T{m.group(2)}+00:00")
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt + timedelta(hours=sc_reports._us_offset(dt, -6))


def period_ranges(days: int, today: Optional[date] = None):
    """이번 기간 = 어제까지 days일, 이전 기간 = 그 앞 같은 길이. 돌려줌 ((이번 시작, 이번 끝), (이전 시작, 이전 끝))."""
    t = today or central_today()
    cur_end = t - timedelta(days=1)
    cur_start = cur_end - timedelta(days=days - 1)
    prev_end = cur_start - timedelta(days=1)
    prev_start = prev_end - timedelta(days=days - 1)
    return (cur_start, cur_end), (prev_start, prev_end)


# ================================================================ 마지막 실행 날짜

def load_last_run() -> Dict:
    try:
        return json.loads(LAST_RUN_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def mark_last_run(key: str, label: str, cur, prev) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    data = load_last_run()
    data[key] = {"label": label, "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                 "central_date": central_today().isoformat(),
                 "current": f"{cur[0]} ~ {cur[1]}", "previous": f"{prev[0]} ~ {prev[1]}"}
    LAST_RUN_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def last_run_text() -> str:
    """메뉴에 보이는 ‘마지막 실행’."""
    data = load_last_run()
    if not data:
        return "아직 실행하지 않음"
    parts = []
    for _, (key, label, _d) in PERIODS.items():
        if key in data:
            parts.append(f"{label} {data[key]['time']}")
    return "마지막 실행: " + "; ".join(parts)


# ================================================================ 브라우저 조작

def _dump(page, name: str) -> Path:
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    path = DEBUG_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{name}_dom.txt"
    try:
        text = page.evaluate(inv._DUMP_JS)
    except Exception as e:
        text = f"(내보내기 실패: {e})"
    path.write_text(f"URL: {page.url}\n\n{text}", encoding="utf-8")
    sc_nav.save_debug(page, DEBUG_DIR, name)
    return path


REPORT_SECTION = {"All Orders": "Sales", "FBA customer returns": "Customer Concessions"}
REPORT_URL = {"All Orders": "/reportcentral/FlatFileAllOrdersReport/1"}


def _left_nav_link(page, names: List[str]):
    """보고서 센터 왼쪽 목차의 보고서 링크(a#report-nav-link-url) — 가운데 “Most Popular Reports”의 같은 이름 링크는 아님."""
    pat = re.compile(r"^\s*(" + "|".join(re.escape(n) for n in names) + r")\s*$", re.I)
    el = sc_nav._first_visible(page.locator("a#report-nav-link-url, a[id^=report-nav-link]").filter(has_text=pat))
    if el is not None:
        return el
    # 페이지 구조가 다르면: 위치로 가장 왼쪽의 같은 이름 링크를 찾음
    for lab in names:
        loc = page.get_by_role("link", name=lab, exact=True)
        best = None
        for i in range(min(loc.count(), 10)):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                box = el.bounding_box()
            except Exception:
                continue
            href = (el.get_attribute("href") or "").strip()
            if href in ("", "#") or href.startswith("javascript"):
                continue
            if box and (best is None or box["x"] < best[1]):
                best = (el, box["x"])
        # 왼쪽 목차는 페이지 맨 왼쪽(실제 페이지 약 x=26), 가운데 Most Popular Reports는 x≈500 이후
        if best is not None and best[1] < 240:
            return best[0]
    return None


def _section_show_more(page, section: str):
    """왼쪽 목차의 어떤 분류(예: Sales) 아래 “Show more...”."""
    header = page.locator("h4").filter(has_text=re.compile(rf"^\s*{re.escape(section)}\s*$", re.I)).first
    if header.count():
        more = header.locator("xpath=following-sibling::div[.//label][1]//label").first
        try:
            if more.count() and more.is_visible() and inv.SHOW_MORE.search(more.inner_text() or ""):
                return more
        except Exception:
            pass
    return None


def open_report_page(page, names: List[str]) -> None:
    """메뉴 공급망 → 공급망 보고서(보고서 센터)에서 왼쪽 목차의 보고서를 엶, 필요하면 그 분류의 “Show more...”를 먼저 누름."""
    if "reportcentral" not in page.url:
        print("   공급망 → 공급망 보고서 여는 중...")
        try:
            sc_nav.menu_go(page, [sc_nav.SUPPLY_CHAIN, sc_nav.SUPPLY_CHAIN_REPORTS], DEBUG_DIR)
        except Exception as e:
            print(f"   메뉴가 열리지 않았습니다({str(e)[:60]}). 보고서 센터를 바로 엽니다. ")
            human.goto(page, sc_nav.SC_BASE + "/reportcentral/WelcomePage", wait_after=False)
            sc_nav._settle(page)
        if "reportcentral" not in page.url:
            tab = sc_nav.find_visible(page, inv.FBA_TAB, prefer_link=True)
            if tab is not None:
                page = inv._click_and_follow(page, tab)
    link = _left_nav_link(page, names)
    if link is None:
        section = REPORT_SECTION.get(names[0])
        more = _section_show_more(page, section) if section else None
        if more is not None:
            print(f"   펼치기: {section} 아래의 Show more...")
            human.click(page, more)
            link = _left_nav_link(page, names)
    if link is None:
        # 다른 분류의 Show more도 펼쳐 봄(펼친 뒤 목록이 바뀌므로 매번 다시 찾음)
        for _ in range(6):
            more = sc_nav._first_visible(page.locator("label.reports-nav-show-link").filter(has_text=inv.SHOW_MORE))
            if more is None:
                more = sc_nav._first_visible(page.get_by_text(inv.SHOW_MORE))
            if more is None:
                break
            print("   Show more 펼치는 중...")
            human.click(page, more)
            link = _left_nav_link(page, names)
            if link is not None:
                break
    if link is None:
        # 가운데 “Most Popular Reports”의 링크, 또는 보고서 주소로 바로 엶
        pop = sc_nav._first_visible(page.locator("kat-link").filter(
            has_text=re.compile(r"^\s*(" + "|".join(re.escape(n) for n in names) + r")\s*$", re.I)))
        if pop is not None:
            link = pop
        elif names[0] in REPORT_URL:
            print(f"   왼쪽 목록에서 “{names[0]}”을(를) 찾지 못해, 해당 보고서 페이지를 바로 엽니다. ")
            human.goto(page, sc_nav.SC_BASE + REPORT_URL[names[0]], wait_after=False)
            sc_nav._settle(page)
            human.read_page(page, max_screens=2)
            human.move_away(page)
            return
    if link is None:
        raise RuntimeError(f"보고서 센터 왼쪽 목록에서 찾지 못함: “{names[0]}”")
    print(f"   보고서 열기: “{names[0]}”...")
    inv._click_and_follow(page, link)
    human.read_page(page, max_screens=2)
    human.move_away(page)   # 마우스를 왼쪽 목차에서 치움(마우스를 올리면 뜨는 설명이 아래 누를 곳을 가리지 않게)
    _open_download_tab(page)


def _open_download_tab(page) -> None:
    """보고서 페이지에 View Online / Download / Schedule 탭이 있으면 Download로 전환(파일 요청은 거기서만 됨)."""
    tab = page.locator("a#reportpage_download_tab").first
    try:
        if not tab.count() or not tab.is_visible():
            return
        if "selected" in (tab.get_attribute("class") or ""):
            return
    except Exception:
        return
    print("   Download 탭으로 전환...")
    inv._click_and_follow(page, tab)
    human.read_page(page, max_screens=2)
    human.move_away(page)


def _choose_option(page, pattern) -> bool:
    """페이지 드롭다운에서 옵션 하나 선택(기본 select 또는 Amazon kat-dropdown)."""
    # 기본 select(Amazon 일부 페이지는 숨기고 꾸민 버튼을 대신 보여 줌: 그때는 그 버튼을 먼저 누르고 옵션 선택)
    sels = page.locator("select")
    for i in range(sels.count()):
        sel = sels.nth(i)
        try:
            cls = sel.get_attribute("class") or ""
            visible = sel.is_visible()
            if not visible and "a-native-dropdown" not in cls:
                continue
            options = sel.locator("option")
            pairs = [(options.nth(j).inner_text().strip(), options.nth(j).get_attribute("value") or "")
                     for j in range(options.count())]
        except Exception:
            continue
        for text, value in pairs:
            if pattern.search(text) and value != "":
                if visible:
                    box = sel.bounding_box()
                    if box:
                        human.move_mouse(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                else:
                    btn = sel.locator("xpath=following-sibling::*[contains(@class,'a-button-dropdown')][1]")
                    if btn.count() and btn.first.is_visible():
                        human.click(page, btn.first, wait_after=False)
                        human.pause(0.6, 1.4)
                sel.select_option(value=value, force=not visible)   # 숨겨진 기본 드롭다운: 값을 맞춤
                page.keyboard.press("Escape")   # 드롭다운 목록 닫기
                human.pause()
                return True
    # kat-dropdown
    dds = page.locator("kat-dropdown")
    for i in range(dds.count()):
        dd = dds.nth(i)
        try:
            if not dd.is_visible():
                continue
        except Exception:
            continue
        option = dd.locator("kat-option").filter(has_text=pattern).first
        if not option.count():
            continue
        header = dd.locator(".select-header").first
        human.click(page, header if header.count() else dd)
        try:
            option.wait_for(state="visible", timeout=5000)
        except Exception:
            page.keyboard.press("Escape")
            continue
        option.scroll_into_view_if_needed()
        human.click(page, option)
        return True
    # 라디오 버튼 / 일반 글자
    el = sc_nav._first_visible(page.get_by_role("radio", name=pattern))
    if el is None:
        el = sc_nav._first_visible(page.get_by_text(pattern))
    if el is not None:
        human.click(page, el)
        return True
    return False


def _date_inputs(page):
    """“From / To” 날짜 입력란 두 개."""
    found = []
    for lab in (FROM_LABELS, TO_LABELS):
        el = sc_nav._first_visible(page.get_by_label(lab))
        found.append(el)
    if all(found):
        return found
    # 주문 보고서 페이지: #all-orders-report-start-date / #all-orders-report-end-date 안의 입력란
    s_ = sc_nav._first_visible(page.locator('[id$="report-start-date"] input.a-cal-input, [id$="start-date"] input[type="text"]'))
    e_ = sc_nav._first_visible(page.locator('[id$="report-end-date"] input.a-cal-input, [id$="end-date"] input[type="text"]'))
    if s_ is not None and e_ is not None:
        return [s_, e_]
    # Amazon 날짜 범위 컴포넌트: kat-date-range-picker 안의 kat-date-picker 두 개(start / end)
    s_ = sc_nav._first_visible(page.locator("kat-date-range-picker kat-date-picker.start input, kat-date-picker.start input"))
    e_ = sc_nav._first_visible(page.locator("kat-date-range-picker kat-date-picker.end input, kat-date-picker.end input"))
    if s_ is not None and e_ is not None:
        return [s_, e_]
    # 형식 안내로 찾기: MM/DD/YYYY 입력란 두 개, 왼쪽이 시작
    loc = page.locator('input[placeholder*="MM" i], input[placeholder*="YYYY" i], kat-date-picker input, '
                       'input[type="date"], kat-date-picker')
    boxes = []
    for i in range(min(loc.count(), 10)):
        el = loc.nth(i)
        try:
            if el.is_visible():
                b = el.bounding_box()
                if b:
                    boxes.append((b["y"], b["x"], el))
        except Exception:
            pass
    boxes.sort(key=lambda t: (round(t[0] / 20), t[1]))
    if len(boxes) >= 2:
        return [boxes[0][2], boxes[1][2]]
    return [None, None]


def _fmt(d: date, el) -> str:
    try:
        if (el.get_attribute("type") or "") == "date":
            return d.isoformat()
    except Exception:
        pass
    return f"{d.month:02d}/{d.day:02d}/{d.year}"


def set_date_range(page, start: date, end: date) -> str:
    """
    보고서 날짜 범위 설정: “Exact dates”를 골라 시작 / 끝 날짜 입력을 먼저 시도,
    안 되면 이 기간을 덮는 미리 정한 범위(예: Last 30 days)를 고르고 나중에 중부 시간 날짜로 거름.
    실제로 쓴 방법 설명을 돌려줌.
    """
    _choose_option(page, EXACT_DATES)
    frm, to = _date_inputs(page)
    if frm is not None and to is not None:
        print(f"   날짜 입력 {start:%m/%d/%Y} ~ {end:%m/%d/%Y}...")
        for attempt in range(2):
            human.type_text(page, frm, _fmt(start, frm), wait_after=False)
            page.keyboard.press("Tab")
            human.pause(0.8, 1.6)
            human.type_text(page, to, _fmt(end, to), wait_after=False)
            page.keyboard.press("Tab")
            page.keyboard.press("Escape")   # 뜰 수 있는 달력 닫기
            human.pause()
            try:
                ok = frm.input_value() == _fmt(start, frm) and to.input_value() == _fmt(end, to)
            except Exception:
                ok = True
            if ok:
                break
            print("   날짜 입력에 실패해 다시 입력합니다...")
        return f"지정 날짜 {start} ~ {end}"
    need = (central_today() - start).days + 1
    for days, names in PRESETS:
        if days >= need:
            pat = re.compile(r"^\s*(" + "|".join(re.escape(n) for n in names) + r")\s*$", re.I)
            if _choose_option(page, pat):
                print(f"   날짜 입력란을 찾지 못해 프리셋 “{names[0]}”을(를) 사용하고, 이후 날짜로 필터링합니다. ")
                return f"프리셋 범위 {names[0]}"
    path = _dump(page, "date_range")
    raise RuntimeError(f"날짜 범위를 설정하는 곳을 찾지 못했습니다(페이지 구조 저장: {path.name})")


ORDER_REPORTS_URL = "/order-reports-and-feeds/reports/allOrders"
ORDER_REQUEST = re.compile(r"^\s*(Request|Request( a)? report|Request download|Request \.txt download|Generate report|"
                           r"보고서 요청|다운로드 요청|请求报告|请求下载)\s*$", re.I)


def open_order_reports(page) -> None:
    """메뉴 주문 → 주문 보고서, 그다음 All Orders 열기."""
    if "/order-reports-and-feeds/" not in page.url:
        print("   주문 → 주문 보고서 여는 중...")
        sc_nav.menu_go(page, [sc_nav.ORDERS, sc_nav.ORDER_REPORTS], DEBUG_DIR)
    if "allorders" not in page.url.lower():
        pat = re.compile(r"^\s*(" + "|".join(re.escape(n) for n in ALL_ORDERS) + r")\s*$", re.I)
        el = None
        for loc in (page.get_by_role("tab", name=pat), page.get_by_role("link", name=pat),
                    page.locator("kat-tab, kat-link, a, button, [role=tab]").filter(has_text=pat)):
            el = sc_nav._first_visible(loc)
            if el is not None:
                break
        if el is not None:
            print("   ‘All Orders’ 여는 중...")
            inv._click_and_follow(page, el)
        else:
            print("   페이지에서 ‘All Orders’를 찾지 못해, All Orders 보고서 페이지를 바로 엽니다. ")
            human.goto(page, sc_nav.SC_BASE + ORDER_REPORTS_URL, wait_after=False)
            sc_nav._settle(page)
    human.read_page(page, max_screens=2)
    human.move_away(page)
    _dump(page, "order_reports_page")   # 실제 페이지에 맞춰 조정할 수 있게 페이지 구조 저장


def _find_order_request_button(page):
    # 주문 보고서 페이지: 노란 “Request” 버튼(a-button 안의 제출 버튼)
    el = sc_nav._first_visible(page.locator('span.a-button:has(input[name="request"])'))
    if el is None:
        el = sc_nav._first_visible(page.locator('input[name="request"]'))
    if el is None:
        el = sc_nav._first_visible(page.locator("kat-button").filter(has_text=ORDER_REQUEST))
    if el is None:
        loc = page.locator("kat-button[label]")
        for i in range(min(loc.count(), 30)):
            b = loc.nth(i)
            try:
                if b.is_visible() and ORDER_REQUEST.search(b.get_attribute("label") or ""):
                    el = b
                    break
            except Exception:
                pass
    if el is None:
        el = sc_nav._first_visible(page.get_by_role("button", name=ORDER_REQUEST))
    return el or inv._find_request_button(page, wait_s=5)


def _row_has_range(text: str, start: date, end: date) -> bool:
    """보고서 목록의 이 줄이 이 날짜 범위인지(여러 날짜 표기 지원)."""
    def forms(d: date):
        mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.month - 1]
        full = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                "October", "November", "December"][d.month - 1]
        return [f"{d.month}/{d.day}/{d.year % 100:02d}", f"{d.month:02d}/{d.day:02d}/{d.year}",
                f"{d.month}/{d.day}/{d.year}", d.isoformat(), f"{mon} {d.day}, {d.year}", f"{full} {d.day}, {d.year}"]
    t = text or ""
    def find(needle, pos=0):
        m = re.compile(r"(?<![\d/])" + re.escape(needle) + r"(?![\d])").search(t, pos)
        return (m.start(), m.end()) if m else None
    # 시작일과 종료일이 모두 있고 시작이 종료보다 앞
    for a in forms(start):
        fa = find(a)
        if not fa:
            continue
        for b in forms(end):
            if find(b, fa[1] if start != end else fa[0]):
                return True
    return False


_ROW_DT = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2,4}),?\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP]M)", re.I)


def _row_covers(text: str, start: date, end: date) -> bool:
    """
    보고서 줄의 기간(예: 10/18/26, 1:01 PM - 10/26/26, 1:01 PM, Amazon이 매일 자동으로 만드는 것)이
    start 날 0시부터 end 날 24시까지를 다 덮는지.
    """
    found = _ROW_DT.findall(text or "")
    if len(found) < 2:
        return False

    def dt(m):
        mo, d, y, hh, mm, ap = int(m[0]), int(m[1]), int(m[2]), int(m[3]), int(m[4]), m[5].upper()
        y = y + 2000 if y < 100 else y
        return datetime(y, mo, d, hh % 12 + (12 if ap == "PM" else 0), mm)
    try:
        a, b = dt(found[0]), dt(found[1])
    except ValueError:
        return False
    return a <= datetime(start.year, start.month, start.day) and b >= datetime(end.year, end.month, end.day, 23, 59)


def request_report(page, names: List[str], start: date, end: date, save_name: str,
                   opener=None, preset=None, need=None) -> Tuple[Path, str]:
    """
    보고서 페이지 열기 → 날짜 설정 → 보고서 요청 → 생성 기다림 → 다운로드.
    preset: 날짜 드롭다운에서 고를 미리 정한 범위(예: Last 30 days), 없으면 Exact dates를 골라 날짜 입력.
    need: 분석에 실제로 필요한 날짜 범위, 목록에 이것을 덮는 보고서가 이미 있으면 바로 다운로드.
    """
    need = need or (start, end)
    try:
        if opener:
            opener(page)
        else:
            open_report_page(page, names)

        # 목록에 같은 날짜 범위, 또는 이 기간을 다 덮는 보고서가 이미 있으면 바로 다운로드
        for el in sc_reports._download_controls(page):
            text = sc_reports.row_text(el)
            if _row_has_range(text, start, end) or _row_covers(text, need[0], need[1]):
                print(f"   다음 기간을 포함하는 보고서가 이미 있어 바로 다운로드합니다: {need[0]} ~ {need[1]} ")
                return sc_reports._download(page, el, RAW_DIR, save_name, DEBUG_DIR), f"기존 보고서 {start} ~ {end}"
        old = sc_reports.existing_rows(page)
        if _pending_same_range(page, start, end):
            print(f"   {start} ~ {end} 보고서는 이미 요청되어 생성 중이므로, 다시 요청하지 않습니다. ")
            print("   보고서 생성 대기 중(최대 45분)...")
            path = sc_reports.wait_and_download(page, old, RAW_DIR, save_name, timeout_min=45, debug_dir=DEBUG_DIR,
                                                on_long_wait=lambda p_: _dump(p_, f"{save_name}_long_wait"))
            return path, f"요청한 보고서 {start} ~ {end}"

        _select_order_date(page)
        if preset is not None:
            print("   날짜 범위 ‘Last 30 days’ 선택...")
            if not _choose_option(page, preset):
                path = _dump(page, "date_preset")
                raise RuntimeError(f"날짜 드롭다운에서 ‘Last 30 days’를 찾지 못했습니다(페이지 구조 저장: {path.name})")
            how = "프리셋 범위 Last 30 days"
        else:
            how = set_date_range(page, start, end)
        old = sc_reports.existing_rows(page)
        btn = _find_order_request_button(page) if opener else inv._find_request_button(page)
        if btn is None:
            path = _dump(page, "request_button")
            raise RuntimeError(f"보고서 요청 버튼을 찾지 못했습니다(페이지 구조 저장: {path.name})")
        name = (btn.get_attribute("label") or btn.inner_text() or "").strip()
        print(f"   클릭: “{name[:30]}”...")
        human.click(page, btn)
        _dump(page, f"{save_name}_after_request")   # 요청 뒤 페이지 구조 저장(보고서 목록 확인용)
        print("   보고서 생성 대기 중(최대 45분)...")
        # 이런 보고서 목록에는 “.txt” 글자가 없을 수 있음: 요청 뒤 새로 나타난, 다운로드 가능한 줄이면 받음
        path = sc_reports.wait_and_download(page, old, RAW_DIR, save_name, row_must_contain=None,
                                            timeout_min=45, debug_dir=DEBUG_DIR,
                                            on_long_wait=lambda p_: _dump(p_, f"{save_name}_long_wait"))
        return path, how
    except Exception:
        _dump(page, save_name)
        raise


def _range_texts(start: date, end: date) -> List[str]:
    """보고서 목록의 ‘Date Range Covered’ 표기, 예: 10/24/26, 12:00 AM - 10/26/26, 11:59 PM."""
    f = lambda d: f"{d.month}/{d.day}/{d.year % 100:02d}"
    return [f"{f(start)}, 12:00 AM", f"{f(end)}, 11:59 PM"]


def _pending_same_range(page, start: date, end: date) -> bool:
    loc = page.get_by_text(sc_reports.STATUS_RE)
    for i in range(min(loc.count(), 20)):
        el = loc.nth(i)
        try:
            if el.is_visible() and _row_has_range(sc_reports.row_text(el), start, end):
                return True
        except Exception:
            pass
    return False


def _select_order_date(page) -> None:
    """“Type of report”를 Order Date로(주문일 기준, 기본값이라 선택 안 됐을 때만 누름)."""
    radio = page.locator('kat-radiobutton[label="Order Date"], kat-radiobutton[value="2400"]').first
    group = page.locator("kat-radiobutton-group").first
    try:
        if radio.count() and group.count() and (group.get_attribute("value") or "") != (radio.get_attribute("value") or ""):
            print("   Order Date 선택...")
            human.click(page, radio)
    except Exception:
        pass


def _chunks(start: date, end: date) -> List[Tuple[date, date]]:
    out, s = [], start
    while s <= end:
        e = min(end, s + timedelta(days=MAX_CHUNK_DAYS - 1))
        out.append((s, e))
        s = e + timedelta(days=1)
    return out


def _find_cached(prefix: str, start: date, end: date, fetched: date) -> Optional[Path]:
    """오늘 이미 받았고 start~end를 덮는 데이터 파일."""
    for f in sorted(REPORT_DIR.glob(f"{prefix}_*_*_{fetched.isoformat()}.xlsx")):
        m = re.match(rf"{prefix}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{4}}-\d{{2}}-\d{{2}})_", f.name)
        if m and date.fromisoformat(m.group(1)) <= start and date.fromisoformat(m.group(2)) >= end:
            return f
    return None


def fetch_data(page_factory, prefix: str, names: List[str], label: str,
               start: date, end: date, today: date, opener=None, preset=None,
               need=None) -> Tuple[Optional[Path], str]:
    """start~end 데이터 얻기: 그날 캐시가 있으면 사용, 없으면 받음(필요하면 여러 번 요청해 합침)."""
    cached = _find_cached(prefix, start, end, today)
    if cached:
        print(f"   오늘 이미 가져온 {label}({cached.name}). 그대로 사용합니다. ")
        return cached, f"{cached.name}(오늘 이미 가져옴, 그대로 사용)"
    headers, rows, hows = [], [], []
    # SP-API 우선(주문: All Orders 보고서, 반품: FBA 고객 반품 보고서). 안 되면 아래 브라우저 방식
    api_fn = {"orders_data": "orders", "returns_data": "returns"}.get(prefix)
    if api_fn:
        import sp_api
        import sp_data
        raw = sp_api.try_api(label, getattr(sp_data, api_fn), start, end, f"{start}_{end}_{today}")
        if raw:
            h, r = inv.read_report(raw)
            out = REPORT_DIR / f"{prefix}_{start}_{end}_{today}.xlsx"
            inv.save_report_xlsx(h, r, out, label)
            print(f"   ✅ {label}(API): {len(r)}행")
            return out, f"{out.name}(SP-API, {len(r)}행)"
    with page_factory() as page:
        for i, (s, e) in enumerate([(start, end)] if preset else _chunks(start, end), start=1):
            print(f"   가져오는 중: {label}: {s} ~ {e}...")
            raw, how = request_report(page, names, s, e, f"{prefix}_{s}_{e}_{today}", opener, preset, need)
            h, r = inv.read_report(raw)
            if h and not headers:
                headers = h
            rows += r
            hows.append(how)
    out = REPORT_DIR / f"{prefix}_{start}_{end}_{today}.xlsx"
    inv.save_report_xlsx(headers, rows, out, label)
    return out, f"{out.name}({'; '.join(hows)}, {len(rows)}행)"


# ================================================================ 데이터 정리

def _g(row: Dict, *names, default=""):
    return inv._get(row, *names, default=default)


def parse_orders(rows: List[Dict]) -> List[Dict]:
    """한 줄 = 주문 하나 안의 상품 하나."""
    out = []
    for r in rows:
        dt = to_central(_g(r, "purchase-date", "purchase_date", "order-date"))
        if dt is None:
            continue
        status = str(_g(r, "order-status", "order_status")).strip()
        item_status = str(_g(r, "item-status")).strip()
        channel = str(_g(r, "fulfillment-channel", "fulfillment_channel")).strip()
        qty = inv._num(_g(r, "quantity", "quantity-purchased"))
        out.append({
            "order_id": str(_g(r, "amazon-order-id", "order-id")).strip(),
            "time": dt, "day": dt.date(),
            "sku": str(_g(r, "sku")).strip(), "asin": str(_g(r, "asin")).strip(),
            "title": str(_g(r, "product-name")).strip(),
            "status": status, "item_status": item_status,
            "cancelled": status.lower().startswith("cancel") or item_status.lower().startswith("cancel"),
            "channel": "FBA" if channel.lower().startswith("amazon") or channel.upper() == "AFN" else "판매자 배송",
            "qty": qty,
            "sales": inv._num(_g(r, "item-price")),
            "state": str(_g(r, "ship-state")).strip().upper(),
            "city": str(_g(r, "ship-city")).strip(),
            "service": str(_g(r, "ship-service-level")).strip(),
            "market": str(_g(r, "sales-channel")).strip(),
            "currency": str(_g(r, "currency")).strip(),
        })
    return out


MARKETPLACE = "Amazon.com"


def keep_us(orders: List[Dict]) -> Tuple[List[Dict], Dict[str, int]]:
    """All Orders 보고서에는 북미 모든 사이트(미국·멕시코·캐나다)가 들어 있음. Amazon.com만 남기고 다른 사이트는 따로 셈."""
    kept, other = [], {}
    for x in orders:
        m = x.get("market") or ""
        if not m or m.lower() == MARKETPLACE.lower():
            kept.append(x)
        else:
            other[m] = other.get(m, 0) + 1
    return kept, other


def parse_returns(rows: List[Dict]) -> List[Dict]:
    out = []
    for r in rows:
        dt = to_central(_g(r, "return-date", "return_date"))
        if dt is None:
            continue
        out.append({"day": dt.date(), "order_id": str(_g(r, "order-id")).strip(),
                    "sku": str(_g(r, "sku")).strip(), "asin": str(_g(r, "asin")).strip(),
                    "title": str(_g(r, "product-name")).strip(), "qty": inv._num(_g(r, "quantity")),
                    "reason": str(_g(r, "reason")).strip(), "disposition": str(_g(r, "detailed-disposition")).strip()})
    return out


def in_range(d: date, rng) -> bool:
    return rng[0] <= d <= rng[1]


def kpis(orders: List[Dict], returns: Optional[List[Dict]], rng) -> Dict:
    o = [x for x in orders if in_range(x["day"], rng)]
    ok = [x for x in o if not x["cancelled"]]
    ids = {x["order_id"] for x in o}
    ok_ids = {x["order_id"] for x in ok}
    cancel_ids = ids - ok_ids
    units = sum(x["qty"] for x in ok)
    sales = sum(x["sales"] for x in ok)
    res = {"주문 수": len(ok_ids), "판매 수량": units, "매출": round(sales, 2),
           "객단가": round(sales / len(ok_ids), 2) if ok_ids else 0,
           "취소 주문 수": len(cancel_ids), "취소율": (len(cancel_ids) / len(ids)) if ids else 0,
           "FBA 수량": sum(x["qty"] for x in ok if x["channel"] == "FBA"),
           "판매자 배송 수량": sum(x["qty"] for x in ok if x["channel"] == "판매자 배송")}
    if returns is not None:
        rq = sum(x["qty"] for x in returns if in_range(x["day"], rng))
        res["반품 수량"] = rq
        res["반품률"] = (rq / units) if units else 0
    return res


def daily(orders: List[Dict], rng) -> List[Dict]:
    days = [rng[0] + timedelta(days=i) for i in range((rng[1] - rng[0]).days + 1)]
    out = []
    for d in days:
        ok = [x for x in orders if x["day"] == d and not x["cancelled"]]
        out.append({"day": d, "orders": len({x["order_id"] for x in ok}),
                    "units": sum(x["qty"] for x in ok), "sales": round(sum(x["sales"] for x in ok), 2)})
    return out


def by_sku(orders: List[Dict], returns: Optional[List[Dict]], cur, prev) -> List[Dict]:
    skus: Dict[str, Dict] = {}
    for x in orders:
        which = "cur" if in_range(x["day"], cur) else ("prev" if in_range(x["day"], prev) else None)
        if which is None or not x["sku"]:
            continue
        s = skus.setdefault(x["sku"], {"sku": x["sku"], "asin": x["asin"], "title": x["title"],
                                       "cur_units": 0, "prev_units": 0, "cur_sales": 0.0, "prev_sales": 0.0,
                                       "cur_lines": 0, "cur_cancel": 0, "cur_returns": 0})
        s["title"] = s["title"] or x["title"]
        if x["cancelled"]:
            if which == "cur":
                s["cur_lines"] += 1
                s["cur_cancel"] += 1
            continue
        s[f"{which}_units"] += x["qty"]
        s[f"{which}_sales"] += x["sales"]
        if which == "cur":
            s["cur_lines"] += 1
    if returns is not None:
        for r in returns:
            if in_range(r["day"], cur) and r["sku"]:
                s = skus.setdefault(r["sku"], {"sku": r["sku"], "asin": r["asin"], "title": r["title"],
                                               "cur_units": 0, "prev_units": 0, "cur_sales": 0.0,
                                               "prev_sales": 0.0, "cur_lines": 0, "cur_cancel": 0,
                                               "cur_returns": 0})
                s["cur_returns"] += r["qty"]
    out = list(skus.values())
    for s in out:
        s["cancel_rate"] = s["cur_cancel"] / s["cur_lines"] if s["cur_lines"] else 0
        s["return_rate"] = (s["cur_returns"] / s["cur_units"]) if (returns is not None and s["cur_units"]) else None
    out.sort(key=lambda s: (-s["cur_units"], -s["cur_sales"], s["sku"]))
    return out


def by_state(orders: List[Dict], cur, prev) -> List[Dict]:
    st: Dict[str, Dict] = {}
    for x in orders:
        if x["cancelled"] or not x["state"]:
            continue
        which = "cur" if in_range(x["day"], cur) else ("prev" if in_range(x["day"], prev) else None)
        if which is None:
            continue
        s = st.setdefault(x["state"], {"state": x["state"], "cur_orders": set(), "cur_units": 0, "cur_sales": 0.0,
                                       "prev_units": 0})
        if which == "cur":
            s["cur_orders"].add(x["order_id"])
            s["cur_units"] += x["qty"]
            s["cur_sales"] += x["sales"]
        else:
            s["prev_units"] += x["qty"]
    out = []
    for s in st.values():
        s["cur_orders"] = len(s["cur_orders"])
        out.append(s)
    out.sort(key=lambda s: (-s["cur_units"], s["state"]))
    return out


def fbm_pending(orders: List[Dict]) -> List[Dict]:
    """아직 출고하지 않은 판매자 배송 주문: 주문 후 지난 시간으로 곧 지연되는지 판단."""
    now = central_now()
    seen, out = set(), []
    for x in sorted(orders, key=lambda x: x["time"]):
        if x["channel"] != "판매자 배송" or x["cancelled"]:
            continue
        st = (x["status"] or "").lower()
        if st not in ("unshipped", "pending", "partiallyshipped", "partially shipped", "shipping"):
            continue
        key = (x["order_id"], x["sku"])
        if key in seen:
            continue
        seen.add(key)
        hours = (now - x["time"]).total_seconds() / 3600
        left = HANDLING_DAYS * 24 - hours
        if left < 0:
            level = "기한 초과"
        elif left < 24:
            level = "기한 임박"
        else:
            level = "정상"
        out.append({**x, "hours": round(hours, 1), "left": round(left, 1), "level": level})
    order = {"기한 초과": 0, "기한 임박": 1, "정상": 2}
    out.sort(key=lambda x: (order[x["level"]], -x["hours"]))
    return out


# ================================================================ Excel

def write_report(out: Path, label: str, cur, prev, orders, returns, sources: Dict[str, str],
                 notes: List[str], compare: bool = True) -> Dict:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FONT = inv.FONT
    NAVY, GREY, RED, GREEN, ORANGE = "1F3864", "6B6B6B", "C01C1C", "1E7B34", "C26A00"
    HEAD = PatternFill("solid", fgColor="DCE6F1")
    thin = Side(style="thin", color="BFBFBF")
    BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
    PCT_CHG = '+0.0%;[Red]-0.0%;0.0%'
    NUM_CHG = '+#,##0;[Red]-#,##0;0'
    MONEY = '"$"#,##0.00'
    MONEY_CHG = '+"$"#,##0.00;[Red]-"$"#,##0.00;"$"0.00'

    def f(size=10.5, bold=False, color=None):
        return Font(name=FONT, size=size, bold=bold, color=color)

    def cell(ws, r, c, v, font=None, num=None, fill=None, border=BOX, wrap=False, align=None):
        x = ws.cell(row=r, column=c, value=v)
        x.font = font or f(10)
        if num:
            x.number_format = num
        if fill:
            x.fill = fill
        if border:
            x.border = border
        x.alignment = Alignment(vertical="center", wrap_text=wrap, horizontal=align)
        return x

    def header(ws, r, cols, widths=None):
        for c, h in enumerate(cols, start=1):
            cell(ws, r, c, h, f(10, True), fill=HEAD, wrap=True, align="center")
            if widths:
                ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
        ws.row_dimensions[r].height = 30

    rng_txt = lambda r: f"{r[0]:%m/%d} ~ {r[1]:%m/%d}" if r[0] != r[1] else f"{r[0]:%m/%d}"
    k_cur, k_prev = kpis(orders, returns, cur), kpis(orders, returns, prev)
    skus = by_sku(orders, returns, cur, prev)
    states = by_state(orders, cur, prev)
    pending = fbm_pending(orders)
    d_cur, d_prev = daily(orders, cur), daily(orders, prev)

    wb = Workbook()

    # ================================================= 개요
    ws = wb.active
    ws.title = "요약"
    ws.sheet_view.showGridLines = False
    for c, w in zip("ABCDE", (22, 18, 18, 16, 14)):
        ws.column_dimensions[c].width = w
    cell(ws, 1, 1, f"주문 보고서 · {label}", f(18, True, NAVY), border=None)
    cell(ws, 2, 1, (f"이번 기간 {cur[0]} ~ {cur[1]}   비교: 이전 기간 {prev[0]} ~ {prev[1]}(미국 중부 시간) " if compare
                    else f"{cur[0]} ~ {cur[1]}(미국 중부 시간, 비교 없음) ") +
                   f"생성: {datetime.now():%Y-%m-%d %H:%M}", f(9.5, color=GREY), border=None)

    r = 4
    cell(ws, r, 1, "핵심 지표", f(13, True, NAVY), border=None)
    r += 1
    header(ws, r, ["지표", f"이번 기간 {rng_txt(cur)}", f"이전 기간 {rng_txt(prev)}" if compare else "이전 기간(데이터 없음)", "변화", "변화율"])
    keys = ["주문 수", "판매 수량", "매출", "객단가", "취소 주문 수", "취소율", "FBA 수량", "판매자 배송 수량"]
    if returns is not None:
        keys += ["반품 수량", "반품률"]
    for k in keys:
        r += 1
        is_rate = k.endswith(("율", "률"))
        is_money = k in ("매출", "객단가")
        num = "0.0%" if is_rate else (MONEY if is_money else "#,##0")
        cell(ws, r, 1, k, f(10, True))
        cell(ws, r, 2, k_cur[k], num=num)
        bad_up = k in ("취소 주문 수", "취소율", "반품 수량", "반품률")   # 이 지표는 오르면 나쁨: 오르면 빨강, 내리면 초록
        if bad_up:
            chg_fmt = "[Red]+0.0%;[Color10]-0.0%;0.0%" if is_rate else "[Red]+#,##0;[Color10]-#,##0;0"
            pct_fmt = "[Red]+0.0%;[Color10]-0.0%;0.0%"
        else:
            chg_fmt = "+0.0%;[Red]-0.0%;0.0%" if is_rate else (MONEY_CHG if is_money else NUM_CHG)
            pct_fmt = PCT_CHG
        if compare:
            cell(ws, r, 3, k_prev[k], num=num)
            cell(ws, r, 4, f"=B{r}-C{r}", num=chg_fmt)
            cell(ws, r, 5, "" if is_rate else f'=IF(C{r}=0,"",B{r}/C{r}-1)', num=pct_fmt)
        else:
            for c_ in (3, 4, 5):
                cell(ws, r, c_, "—", f(10, color=GREY), align="center")
    if returns is None:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell(ws, r, 1, "반품 데이터를 가져오지 못해 이번에는 반품률을 계산하지 않습니다. ", f(9, color=ORANGE), border=None)

    # ---- 결론
    r += 2
    cell(ws, r, 1, "요점", f(13, True, NAVY), border=None)
    points = []
    def chg(a, b):
        return (a / b - 1) if b else None
    c_units = chg(k_cur["판매 수량"], k_prev["판매 수량"]) if compare else None
    c_sales = chg(k_cur["매출"], k_prev["매출"]) if compare else None
    if not compare:
        points.append((None, f"판매 수량 {k_cur['판매 수량']:,.0f}, 매출 ${k_cur['매출']:,.2f}, "
                             f"주문 {k_cur['주문 수']:,}건, 객단가 ${k_cur['객단가']:,.2f}. "))
    if c_units is not None:
        points.append((RED if c_units < -0.1 else (GREEN if c_units > 0.1 else None),
                       f"판매 수량 {k_cur['판매 수량']:,.0f}, 이전 기간 대비 {'증가' if c_units >= 0 else '감소'} {abs(c_units):.1%}; "
                       f"매출 ${k_cur['매출']:,.2f}" + (f", 이전 기간 대비 {'증가' if c_sales >= 0 else '감소'} {abs(c_sales):.1%}" if c_sales is not None else "") + ". "))
    movers = [s for s in skus if s["cur_units"] or s["prev_units"]]
    ups = sorted((s for s in movers if s["cur_units"] > s["prev_units"]), key=lambda s: s["prev_units"] - s["cur_units"])[:3]
    downs = sorted((s for s in movers if s["cur_units"] < s["prev_units"]), key=lambda s: s["cur_units"] - s["prev_units"])[:3]
    if not compare:
        top = [s for s in skus if s["cur_units"]][:3]
        if top:
            points.append((None, "가장 많이 팔린 SKU: " + "; ".join(f"{s['sku']}({s['cur_units']:g}개)" for s in top)))
        ups, downs = [], []
    if ups:
        points.append((GREEN, "가장 많이 증가: " + "; ".join(f"{s['sku']}({s['prev_units']:g} → {s['cur_units']:g}개)" for s in ups)))
    if downs:
        points.append((RED, "가장 많이 감소: " + "; ".join(f"{s['sku']}({s['prev_units']:g} → {s['cur_units']:g}개)" for s in downs)))
    hi_cancel = [s for s in skus if s["cur_lines"] >= MIN_UNITS_FOR_RATE and s["cancel_rate"] > CANCEL_RATE_ALERT]
    if hi_cancel:
        points.append((ORANGE, f"취소율이 {CANCEL_RATE_ALERT:.0%} 초과인 SKU {len(hi_cancel)}개, ‘취소 및 반품’ 시트를 참고하세요. "))
    hi_ret = [s for s in skus if s["return_rate"] is not None and s["cur_units"] >= MIN_UNITS_FOR_RATE
              and s["return_rate"] > RETURN_RATE_ALERT]
    if hi_ret:
        points.append((ORANGE, f"반품률이 {RETURN_RATE_ALERT:.0%} 초과인 SKU {len(hi_ret)}개, ‘취소 및 반품’ 시트를 참고하세요. "))
    late = [p for p in pending if p["level"] != "정상"]
    if late:
        n_over = sum(1 for p in late if p["level"] == "기한 초과")
        points.append((RED, f"판매자 배송 미발송 주문 중 기한 초과 {n_over}건, 기한 임박 {len(late) - n_over}건입니다. 빨리 발송하세요(‘FBM 미발송’ 시트 참고). "))
    if states:
        top = states[:3]
        total = sum(s["cur_units"] for s in states) or 1
        points.append((None, "판매가 가장 많은 주: " + ", ".join(f"{s['state']}({s['cur_units'] / total:.0%})" for s in top) + ". "))
    if not points:
        points.append((GREY, "이번 기간 주문 데이터가 없습니다. "))
    for color, text in points:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell(ws, r, 1, "• " + text, f(10.5, color=color), border=None, wrap=True)
        ws.row_dimensions[r].height = 30 if len(text) > 55 else 18

    r += 2
    cell(ws, r, 1, "데이터 출처", f(13, True, NAVY), border=None)
    for k, v in sources.items():
        r += 1
        cell(ws, r, 1, k, f(10, True), border=None)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        cell(ws, r, 2, v, f(9.5, color=GREY), border=None)
    for n in notes:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell(ws, r, 1, n, f(9, color=GREY), border=None, wrap=True)
        ws.row_dimensions[r].height = 30

    # ================================================= 일별 추세
    ds = wb.create_sheet("일별 추이")
    header(ds, 1, ["일차", "이번 기간 날짜", "이번 기간 주문", "이번 기간 수량", "이번 기간 매출", "이전 기간 날짜", "이전 기간 주문", "이전 기간 수량", "이전 기간 매출"],
           [8, 12, 10, 10, 13, 12, 10, 10, 13])
    if not compare:   # 이전 기간 없음: 오른쪽 이전 기간 열을 뺌
        for c_ in range(6, 10):
            ds.cell(row=1, column=c_).value = None
            ds.cell(row=1, column=c_).fill = PatternFill()
            ds.cell(row=1, column=c_).border = Border()
    for i, (a, b) in enumerate(zip(d_cur, d_prev), start=2):
        vals = [i - 1, a["day"].isoformat(), a["orders"], a["units"], a["sales"]]
        if compare:
            vals += [b["day"].isoformat(), b["orders"], b["units"], b["sales"]]
        for c, v in enumerate(vals, start=1):
            cell(ds, i, c, v, num=MONEY if c in (5, 9) else None, align="center" if c in (1, 2, 6) else None)
    last = len(d_cur) + 1
    tr = last + 1
    cell(ds, tr, 1, "합계", f(10, True))
    for c in ((3, 4, 5, 7, 8, 9) if compare else (3, 4, 5)):
        col = get_column_letter(c)
        cell(ds, tr, c, f"=SUM({col}2:{col}{last})", f(10, True), num=MONEY if c in (5, 9) else "#,##0")
    if len(d_cur) >= 2:
        vs = ": 이번 기간 vs 이전 기간" if compare else ""
        for title, cols, pos in ((f"일별 판매 수량{vs}", (4, 8), "K2"), (f"일별 매출{vs}", (5, 9), "K20")):
            ch = LineChart()
            ch.title, ch.height, ch.width = title, 8, 18
            for c in (cols if compare else cols[:1]):
                ch.add_data(Reference(ds, min_col=c, min_row=1, max_row=last), titles_from_data=True)
            ch.set_categories(Reference(ds, min_col=2, min_row=2, max_row=last))
            for sr in ch.series:
                sr.smooth = False
            ds.add_chart(ch, pos)

    # ================================================= SKU 순위
    ss = wb.create_sheet("SKU 순위")
    cols = ["순위", "SKU", "ASIN", "상품명", "이번 기간 수량", "이전 기간 수량", "수량 변화", "변화율", "이번 기간 매출", "이전 기간 매출",
            "매출 변화", "추세", "취소율", "반품률"]
    header(ss, 1, cols, [6, 18, 13, 40, 10, 10, 10, 9, 13, 13, 13, 7, 9, 9])
    ss.freeze_panes = "C2"
    for i, s in enumerate(skus, start=2):
        trend = ("—" if not compare else
                 "신규" if not s["prev_units"] and s["cur_units"] else ("↑" if s["cur_units"] > s["prev_units"]
                                                                       else ("↓" if s["cur_units"] < s["prev_units"] else "—")))
        tcolor = GREEN if trend in ("↑", "신규") else (RED if trend == "↓" else GREY)
        vals = [i - 1, s["sku"], s["asin"], s["title"], s["cur_units"], s["prev_units"], f"=E{i}-F{i}",
                f'=IF(F{i}=0,"",E{i}/F{i}-1)', round(s["cur_sales"], 2), round(s["prev_sales"], 2), f"=I{i}-J{i}",
                trend, s["cancel_rate"] if s["cur_lines"] else "", "" if s["return_rate"] is None else s["return_rate"]]
        nums = {7: NUM_CHG, 8: PCT_CHG, 9: MONEY, 10: MONEY, 11: MONEY_CHG, 13: "0.0%", 14: "0.0%"}
        for c, v in enumerate(vals, start=1):
            cell(ss, i, c, v, f(10, c == 12, tcolor if c == 12 else None), num=nums.get(c),
                 align="center" if c in (1, 12) else None)
    if skus:
        ss.auto_filter.ref = f"A1:N{len(skus) + 1}"

    # ================================================= 취소와 반품
    cs = wb.create_sheet("취소 및 반품")
    header(cs, 1, ["SKU", "ASIN", "상품명", "이번 기간 주문 행 수", "취소 행 수", "취소율", "이번 기간 판매 수량", "반품 수량", "반품률", "알림"],
           [18, 13, 40, 12, 10, 9, 12, 10, 9, 30])
    rows = [s for s in skus if s["cur_cancel"] or s["cur_returns"]]
    rows.sort(key=lambda s: -(max(s["cancel_rate"], s["return_rate"] or 0)))
    for i, s in enumerate(rows, start=2):
        tips = []
        if s["cur_lines"] >= MIN_UNITS_FOR_RATE and s["cancel_rate"] > CANCEL_RATE_ALERT:
            tips.append("취소율 높음")
        if s["return_rate"] is not None and s["cur_units"] >= MIN_UNITS_FOR_RATE and s["return_rate"] > RETURN_RATE_ALERT:
            tips.append("반품률 높음")
        vals = [s["sku"], s["asin"], s["title"], s["cur_lines"], s["cur_cancel"], f'=IF(D{i}=0,"",E{i}/D{i})',
                s["cur_units"], s["cur_returns"] if returns is not None else "",
                f'=IF(OR(G{i}=0,H{i}=""),"",H{i}/G{i})', ", ".join(tips)]
        for c, v in enumerate(vals, start=1):
            cell(cs, i, c, v, f(10, c == 10, RED if (c == 10 and tips) else None),
                 num={6: "0.0%", 9: "0.0%"}.get(c))
    note_r = cs.max_row + 2
    cell(cs, note_r, 1, f"취소율 = 취소된 주문 행 ÷ 이번 기간 주문 행. 반품률 = 이번 기간 반품 수량 ÷ 이번 기간 판매 수량(반품은 더 이전 주문일 수 있으므로 참고용). "
                        f" 수량이 {MIN_UNITS_FOR_RATE}개 미만인 SKU는 알림을 표시하지 않습니다. ", f(9, color=GREY), border=None)

    if returns:
        rs = wb.create_sheet("반품 상세")
        header(rs, 1, ["반품일(중부)", "주문 번호", "SKU", "ASIN", "상품명", "수량", "반품 사유", "처리"],
               [14, 22, 18, 13, 40, 7, 22, 18])
        cur_ret = [x for x in returns if in_range(x["day"], cur)]
        for j, x in enumerate(sorted(cur_ret, key=lambda x: x["day"]), start=2):
            for c, v in enumerate([x["day"].isoformat(), x["order_id"], x["sku"], x["asin"], x["title"], x["qty"],
                                   x["reason"], x["disposition"]], start=1):
                cell(rs, j, c, v)
        if not cur_ret:
            cell(rs, 2, 1, "이번 기간 반품 없음 ✓", f(10, color=GREEN), border=None)

    # ================================================= 판매자 배송 처리 대기
    ps = wb.create_sheet("FBM 미발송")
    header(ps, 1, ["상태", "주문 번호", "주문 시간(중부)", "경과 시간", "남은 시간", "SKU", "상품명", "수량", "배송 방식", "주"],
           [10, 22, 18, 10, 10, 18, 40, 7, 16, 7])
    for i, p in enumerate(pending, start=2):
        color = RED if p["level"] == "기한 초과" else (ORANGE if p["level"] == "기한 임박" else GREEN)
        vals = [p["level"], p["order_id"], p["time"].strftime("%Y-%m-%d %H:%M"), p["hours"], p["left"],
                p["sku"], p["title"], p["qty"], p["service"], p["state"]]
        for c, v in enumerate(vals, start=1):
            cell(ps, i, c, v, f(10, c == 1, color if c == 1 else None))
    nr = len(pending) + 3
    ps.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=10)
    cell(ps, nr, 1, (f"처리 기간 {HANDLING_DAYS}일 기준 추정: 주문 후 {HANDLING_DAYS * 24}시간이 지나도 미발송이면 ‘기한 초과’, "
                     f"남은 시간이 24시간 미만이면 ‘기한 임박’. 실제 발송 기한은 셀러 센트럴 주문 관리 화면을 기준으로 하세요. 처리 기간은 orders.py 맨 위에서 수정할 수 있습니다. "
                     if pending else "현재 미발송 판매자 배송 주문이 없습니다 ✓"), f(9, color=GREY), border=None, wrap=True)
    ps.row_dimensions[nr].height = 30

    # ================================================= 지역 분포
    gs = wb.create_sheet("지역 분포")
    header(gs, 1, ["주", "이번 기간 주문", "이번 기간 수량", "수량 비중", "이번 기간 매출", "이전 기간 수량", "수량 변화"], [8, 10, 10, 10, 13, 10, 10])
    total_units = sum(s["cur_units"] for s in states) or 1
    for i, s in enumerate(states, start=2):
        vals = [s["state"], s["cur_orders"], s["cur_units"], s["cur_units"] / total_units, round(s["cur_sales"], 2),
                s["prev_units"], f"=C{i}-F{i}"]
        for c, v in enumerate(vals, start=1):
            cell(gs, i, c, v, num={4: "0.0%", 5: MONEY, 7: NUM_CHG}.get(c), align="center" if c == 1 else None)
    if states:
        n = min(15, len(states))
        bc = BarChart()
        bc.type, bc.title, bc.height, bc.width = "bar", f"판매 상위 {n}개 주(이번 기간 수량)", 9, 16
        bc.legend = None
        bc.add_data(Reference(gs, min_col=3, min_row=1, max_row=n + 1), titles_from_data=True)
        bc.set_categories(Reference(gs, min_col=1, min_row=2, max_row=n + 1))
        bc.y_axis.majorGridlines = None
        bc.x_axis.scaling.orientation = "maxMin"
        gs.add_chart(bc, "I2")

    # ================================================= 원본 주문
    os_ = wb.create_sheet("주문 상세")
    oh = ["중부 날짜", "주문 시간(중부)", "주문 번호", "상태", "배송", "SKU", "ASIN", "상품명", "수량", "금액", "주", "도시"]
    orows = [{"중부 날짜": x["day"].isoformat(), "주문 시간(중부)": x["time"].strftime("%Y-%m-%d %H:%M"),
              "주문 번호": x["order_id"], "상태": x["status"], "배송": x["channel"], "SKU": x["sku"], "ASIN": x["asin"],
              "상품명": x["title"], "수량": x["qty"], "금액": x["sales"], "주": x["state"], "도시": x["city"]}
             for x in sorted(orders, key=lambda x: x["time"], reverse=True) if in_range(x["day"], (prev[0], cur[1] + timedelta(days=1)))]
    inv.write_table_sheet(os_, oh, orows, {"상품명": 40, "주문 번호": 22, "주문 시간(중부)": 17})

    for sheet in wb.worksheets:
        sheet.page_setup.orientation = "portrait" if sheet.title == "요약" else "landscape"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0

    wb.save(out)
    return {"cur": k_cur, "prev": k_prev, "points": [p[1] for p in points], "late": late}


# ================================================================ 메인 흐름

def choose_period() -> Optional[Tuple[str, str, int]]:
    print("   날짜 범위를 선택하세요(미국 중부 시간): ")
    for k, (_, label, days) in PERIODS.items():
        (cs, ce), (ps_, pe) = period_ranges(days)
        cmp_txt = f"비교 {ps_} ~ {pe}" if has_compare(days) else "비교 없음, 셀러 센트럴에서 최대 30일까지만 가져올 수 있음"
        rng = f"{cs}" if cs == ce else f"{cs} ~ {ce}"
        if has_compare(days) and ps_ == pe:
            cmp_txt = f"비교 {ps_}"
        print(f"     {k}. {label}: {rng}({cmp_txt})")
    try:
        c = input("   번호를 입력하고 Enter를 누르세요(그냥 Enter = 최근 7일): ").strip() or "2"
    except (KeyboardInterrupt, EOFError):
        return None
    if c not in PERIODS:
        print("   해당 옵션이 없습니다. ")
        return None
    return PERIODS[c]


def run(period: Optional[str] = None, open_report: bool = True) -> Optional[Path]:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    print("   " + last_run_text())
    choice = PERIODS.get(period) if period else choose_period()
    if choice is None:
        return None
    key, label, days = choice
    today = central_today()
    cur, prev = period_ranges(days, today)
    compare = has_compare(days)
    if compare:
        print(f"   이번 기간 {cur[0]} ~ {cur[1]}, 이전 기간 {prev[0]} ~ {prev[1]}(미국 중부 시간, 오늘은 {today})")
    else:
        print(f"   이번 기간 {cur[0]} ~ {cur[1]}(미국 중부 시간, 오늘은 {today}; 비교 없음)")

    # 하루 한 번 최근 30일을 받음(오늘까지의 주문 포함, 오늘 데이터는 판매자 배송 미출고 확인에만 씀)
    start, end = today - timedelta(days=FETCH_DAYS), today
    need = (prev[0] if compare else cur[0], cur[1])
    sources, notes = {}, []

    from contextlib import ExitStack
    stack = ExitStack()
    page_holder = {}

    def page_factory():
        class _Ctx:
            def __enter__(self_):
                if "page" not in page_holder:
                    page_holder["page"] = stack.enter_context(open_session(headless=False))
                return page_holder["page"]

            def __exit__(self_, *a):
                return False
        return _Ctx()

    try:
        try:
            odata, src = fetch_data(page_factory, "orders_data", ALL_ORDERS, "주문 데이터", start, end, today,
                                    opener=open_order_reports, preset=PRESET_30, need=need)
            sources["주문"] = src
        except Exception as e:
            raise RuntimeError(f"주문 보고서 가져오기 실패: {type(e).__name__}: {str(e).splitlines()[0][:200]}"
                               f"(스크린샷과 페이지 구조는 reports/orders/debug에 있음)") from e
        returns = None
        try:
            # 같은 브라우저 창을 계속 사용: 이미 보고서 센터에 있으면 왼쪽 목차의 반품 보고서를 바로 누름
            rdata, rsrc = fetch_data(page_factory, "returns_data", RETURNS, "반품 데이터", start, end, today,
                                     preset=PRESET_30, need=need)
            sources["반품"] = rsrc
            returns = parse_returns(inv.read_xlsx(rdata)[1])
        except Exception as e:
            msg = f"반품 보고서를 가져오지 못했습니다({type(e).__name__}: {str(e).splitlines()[0][:150]}). 이번에는 반품률을 계산하지 않습니다. "
            print("   ⚠", msg)
            notes.append(msg)
    finally:
        stack.close()

    orders, other = keep_us(parse_orders(inv.read_xlsx(odata)[1]))
    if other:
        notes.append("Amazon.com(미국) 주문만 집계합니다. 보고서의 다른 마켓플레이스 주문 행은 제외했습니다: " +
                     ", ".join(f"{k} {v}행" for k, v in other.items()) + ". ")
    notes.append("주문 시간은 보고서의 UTC를 미국 중부 시간으로 변환한 뒤 날짜별로 집계했습니다. 취소된 주문은 판매 수량과 매출에 포함하지 않습니다. ")
    out = REPORT_DIR / f"orders_report_{key}_{today}.xlsx"
    if not compare:
        notes.append("셀러 센트럴의 날짜 범위는 최대 30일이라 ‘최근 30일’은 이전 기간 데이터가 없어 비교하지 않습니다. "
                     " 가장 이른 하루의 데이터는 불완전할 수 있습니다(프리셋 범위는 미국 태평양 시간 기준). ")
    res = write_report(out, label, cur, prev, orders, returns, sources, notes, compare=compare)
    mark_last_run(key, label, cur, prev)
    extra = []
    try:
        import sc_pages
        print("\n   반품 분석(Returns and recovery insights) 확인...")
        extra = [sc_pages.read_many(["returns"])["returns"]]
        sc_pages.append_sheets(out, extra)
        sc_pages.print_alerts(extra)
    except Exception as e:
        print(f"   ⚠ 반품 분석을 읽지 못했습니다: {type(e).__name__}: {str(e)[:100]}")
    try:
        import daily_bus
        kc_ = res["cur"]
        daily_bus.record("주문 보고서", {f"{label} {k}": kc_.get(k) for k in ("주문 수", "판매 수량", "매출") if k in kc_},
                         daily_bus.points(res.get("points", [])[1:10], "주문 보고서 확인"), out)
        daily_bus.merge_pages("주문 보고서", extra)
    except Exception:
        pass

    # ---- 터미널 요약
    print()
    kc, kp = res["cur"], res["prev"]
    for k in ("주문 수", "판매 수량", "매출"):
        a, b = kc[k], kp[k]
        fmt = (lambda v: f"${v:,.2f}") if k == "매출" else (lambda v: f"{v:,.0f}")
        if compare:
            ch = f"{(a / b - 1):+.1%}" if b else "—"
            print(f"{k}: 이번 기간 {fmt(a)}, 이전 기간 {fmt(b)}({ch})")
        else:
            print(f"{k}: {fmt(a)}")
    for p in res["points"][1:6]:
        print("  •", p)
    print(f"\n📊 주문 보고서 저장 완료: {out}")
    if open_report:
        try:
            __import__("platform_util").open_path(out)
        except Exception:
            pass
    return out


if __name__ == "__main__":
    run()

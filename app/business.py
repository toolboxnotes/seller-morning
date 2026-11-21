"""
비즈니스 보고서 · 상세 페이지 판매 및 트래픽(하위 ASIN별)

기간 선택: 1 전날 / 2 최근 7일 / 3 최근 14일 / 4 최근 30일(미국 중부 시간)
  - 비교용으로 바로 앞 같은 길이 기간도 받음(비즈니스 보고서는 최근 2년까지 조회 가능, 네 가지 모두 비교 가능)
  - 보고서는 기간 전체를 ASIN별로 합친 것이라 이번 기간·이전 기간을 따로 받음, 같은 기간은 하루 한 번만

데이터 출처:
  - 메뉴 마케팅(Marketing) → 워크스페이스(Workspace): 페이지에 보이는 내용 읽기(상단 지표, 전환율, 인기 검색어, 프로모션, 고객층)
  - 워크스페이스 Conversion Rate 카드의 “View details” → 비즈니스 보고서 → 상세 페이지 판매 및 트래픽(하위 ASIN별)
    세션, 페이지뷰, Buy Box 비율, 주문 수량, 주문 매출, 전환율(Unit Session Percentage)

파일은 reports/business/ 에 저장:
  business_data_시작_끝.xlsx           기간별 보고서 데이터
  business_report_기간_날짜.xlsx       비즈니스 보고서 분석
  last_run.json                         기간별 마지막 실행 시각
  raw/ debug/                           원본 다운로드 파일 / 페이지 구조(처음 실행·오류 때)
"""

import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import human
import sc_labels
import inventory as inv
import orders as ordm
import sc_nav
import sc_reports
from sc_browser import open_session

REPORT_DIR = Path(__file__).resolve().parent.parent / "reports" / "business"
RAW_DIR = REPORT_DIR / "raw"
DEBUG_DIR = REPORT_DIR / "debug"
LAST_RUN_FILE = REPORT_DIR / "last_run.json"

PERIODS = ordm.PERIODS  # 전날 / 7일 / 14일 / 30일

# ---------------------------------------------------------------- 평가 기준(조정 가능)
LOW_CONV_RATIO = 0.5       # 전환율이 매장 평균의 50% 미만이면 ‘전환 낮음’
HIGH_CONV_RATIO = 1.5      # 전환율이 매장 평균의 1.5배 넘으면 ‘전환 높음’
BUYBOX_ALERT = 90.0        # Buy Box 비율 90% 미만이면 알림
DROP_ALERT = 0.30          # 세션이나 전환율이 이전 기간보다 30% 넘게 떨어지면 알림
MIN_SESSIONS_PER_DAY = 3   # ‘트래픽 많음’‘크게 하락’ 판단 때 세션이 최소 일수 × 이 값은 돼야 함(표본이 너무 적지 않게)

# ---------------------------------------------------------------- 페이지
BR_URL = "/business-reports"
WORKSPACE_URL = "/amazonsell/marketing"
MARKETING = sc_labels.L("Marketing")
WORKSPACE = ["Workspace", "작업장", "工作区"]
REPORT_ID = "102:DetailSalesTrafficByChildItem"
MENU_PATHS = [
    [["Reports", "보고서", "报告"], ["Business reports", "Business Reports", "비즈니스 보고서", "业务报告"]],
    [["My business", "내 비즈니스", "我的业务"], ["Business reports", "Business Reports", "비즈니스 보고서", "业务报告"]],
]
CHILD_REPORT = re.compile(r"(Detail Page Sales and Traffic By Child Item|By Child Item|"
                          r"상세 페이지 판매 및 트래픽.*하위|详情页面销售和流量.*子)", re.I)
DOWNLOAD_RE = re.compile(r"^\s*(Download|Download \(\.csv\)|CSV|다운로드|下载)\s*$", re.I)


# ================================================================ 마지막 실행

def load_last_run() -> Dict:
    try:
        return json.loads(LAST_RUN_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def mark_last_run(key: str, label: str) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    data = load_last_run()
    data[key] = {"label": label, "time": datetime.now().strftime("%Y-%m-%d %H:%M")}
    LAST_RUN_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def last_run_text() -> str:
    data = load_last_run()
    if not data:
        return "아직 실행하지 않음"
    parts = [f"{label} {data[key]['time']}" for _, (key, label, _d) in PERIODS.items() if key in data]
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


def _report_url(start: date, end: date) -> str:
    """비즈니스 보고서 페이지 주소에 보고서 종류와 날짜를 바로 넣을 수 있음."""
    from urllib.parse import quote
    return (f"{sc_nav.SC_BASE}{BR_URL}#/report?id={quote(REPORT_ID)}"
            f"&chartCols=&columns=&fromDate={start.isoformat()}&toDate={end.isoformat()}")


def open_workspace(page) -> None:
    """메뉴 마케팅 → 워크스페이스. 메뉴의 분류마다 “Workspace”가 있어 마케팅을 연 뒤 그 아래 것만 누름."""
    if WORKSPACE_URL in page.url:
        return
    print("   마케팅 → 워크스페이스 여는 중...")
    try:
        if not sc_nav.open_menu(page):
            raise RuntimeError("왼쪽 위 메뉴 버튼을 찾지 못했습니다")
        grp = sc_nav.find_visible(page, MARKETING)
        if grp is None:
            raise RuntimeError("메뉴에서 ‘Marketing’을 찾지 못했습니다")
        human.click(page, grp)
        link = sc_nav._first_visible(page.locator(f'a[href*="{WORKSPACE_URL}"]'))
        if link is None:
            raise RuntimeError("마케팅 아래에서 ‘Workspace’를 찾지 못했습니다")
        human.click(page, link, wait_after=False)
        try:
            page.wait_for_url(lambda u: WORKSPACE_URL in u, timeout=20000)
        except Exception:
            pass
    except Exception as e:
        print(f"   메뉴 조작에 실패했습니다({str(e)[:60]}). 마케팅 워크스페이스를 바로 엽니다. ")
        page.keyboard.press("Escape")
        human.goto(page, sc_nav.SC_BASE + WORKSPACE_URL, wait_after=False)
    sc_nav._settle(page)
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    human.read_page(page, max_screens=4)
    human.move_away(page)


def open_business_reports(page) -> None:
    """워크스페이스 Conversion Rate 카드의 “View details”로 비즈니스 보고서에 들어감, 못 찾으면 비즈니스 보고서 페이지를 바로 엶."""
    if "/business-reports" in page.url:
        return
    card = page.locator("div, section, kat-card").filter(
        has_text=re.compile(r"Conversion Rate|전환율|转化率")).filter(has_text=re.compile(r"View details|세부 정보|查看详情"))
    link = None
    for loc in (card.last.get_by_role("link", name=re.compile(r"^\s*(View details|세부 정보 보기|查看详情)\s*$", re.I)),
                card.last.get_by_text(re.compile(r"^\s*(View details|세부 정보 보기|查看详情)\s*$", re.I))):
        try:
            link = sc_nav._first_visible(loc)
        except Exception:
            link = None
        if link is not None:
            break
    if link is not None:
        print("   Conversion Rate의 ‘View details’ 클릭...")
        target = inv._click_and_follow(page, link)
        if target is not page:          # 새 탭으로 열리면: 주소를 가져와 지금 탭에서 열고 이후 같은 페이지에서 작업
            url = target.url
            target.close()
            human.goto(page, url, wait_after=False)
            sc_nav._settle(page)
        if "/business-reports" in page.url:
            return
    print("   워크스페이스에서 비즈니스 보고서로 이동하지 못해, 비즈니스 보고서 페이지를 바로 엽니다. ")
    human.goto(page, sc_nav.SC_BASE + BR_URL, wait_after=False)
    sc_nav._settle(page)


# ---------------------------------------------------------------- 워크스페이스 내용

def read_workspace(page) -> Dict:
    """마케팅 워크스페이스 페이지에서 확인할 수 있는 내용 읽기(조회만, 버튼은 누르지 않음)."""
    try:
        lines = [l.strip() for l in page.inner_text("body").splitlines() if l.strip()]
    except Exception:
        return {}
    ws = {"tiles": [], "conversion": [], "queries": [], "promotions": [], "segments": [], "raw": lines}

    # 상단 지표: 분류 / 지표 이름 / 값
    for name in ("My business", "Products", "Supply chain", "Orders", "Marketing", "Finance"):
        for i, l in enumerate(lines[:-2]):
            if l == name and re.match(r"^[\$\d]", lines[i + 2]):
                ws["tiles"].append((name, lines[i + 1], lines[i + 2]))
                break

    def section(title, stop):
        if title not in lines:
            return []
        i = lines.index(title) + 1
        out = []
        while i < len(lines) and lines[i] not in stop and len(out) < 80:
            out.append(lines[i])
            i += 1
        return out

    heads = {"Advertising", "Promotions", "Customer Segments", "Conversion Rate", "Top Search Queries",
             "Site map", "Recommendations"}
    conv = section("Conversion Rate", heads)
    for k in ("Ordered revenue", "Sessions", "Average sales per order item", "Units ordered", "Unit session percentage"):
        if k in conv:
            j = conv.index(k)
            if j + 1 < len(conv):
                ws["conversion"].append((k, conv[j + 1]))
    upd = next((c for c in conv if c.startswith("Updated")), "")
    if upd:
        ws["conversion"].append(("업데이트 시간", upd.replace("Updated:", "").strip()))
    rng = next((c for c in conv if re.match(r"^(Last \d+ days|Last day|Yesterday)", c)), "")
    if rng:
        ws["conversion"].append(("범위", rng))

    q = section("Top Search Queries", heads)
    for j, l in enumerate(q):
        m = re.match(r"^#(\d+)\s*(.*)$", l)
        if not m:
            continue
        rest = [m.group(2)] if m.group(2) else []
        rest += q[j + 1:j + 4]
        vals = [x for x in rest if x][:3]
        if len(vals) >= 3 and vals[1].endswith("%"):
            ws["queries"].append((int(m.group(1)), vals[0], vals[1], vals[2]))
    rng = next((c for c in q if re.search(r"\d{2}/\d{2}/\d{4}\s*-\s*\d{2}/\d{2}/\d{4}", c)), "")
    ws["queries_range"] = rng

    promo = section("Promotions", heads)
    for j, l in enumerate(promo):
        if l in ("Sales", "Units sold", "Glance views") and j + 1 < len(promo):
            grp = "Deals" if "Coupons" not in promo[:j] else "Coupons"
            ws["promotions"].append((grp, l, promo[j + 1]))

    seg = section("Customer Segments", heads)
    for k in ("Total customers", "New customers", "Potential customers", "Total", "Sales", "Purchase rate",
              "Purchase interval"):
        for j, l in enumerate(seg):
            if l == k and j + 1 < len(seg):
                ws["segments"].append((k, seg[j + 1]))
    return ws


ASIN_RE = re.compile(r"\bB0[A-Z0-9]{8}\b")


def _fmt_us(d: date) -> str:
    return d.strftime("%m/%d/%Y")


def open_child_report(page) -> None:
    """비즈니스 보고서 왼쪽 목차 By ASIN → Detail Page Sales and Traffic By Child Item."""
    if "DetailSalesTrafficByChildItem" in page.url:
        return
    link = None
    for loc in (page.locator("kat-link").filter(has_text=CHILD_REPORT),
                page.get_by_role("link", name=CHILD_REPORT),
                page.get_by_text(CHILD_REPORT)):
        link = sc_nav._first_visible(loc)
        if link is not None:
            break
    if link is not None:
        print("   왼쪽 ‘Detail Page Sales and Traffic By Child Item’ 클릭...")
        human.click(page, link, wait_after=False)
        try:
            page.wait_for_url(lambda u: "DetailSalesTrafficByChildItem" in u, timeout=15000)
        except Exception:
            pass
    if "DetailSalesTrafficByChildItem" not in page.url:
        print("   ‘상세 페이지 판매 및 트래픽(하위 ASIN별)’ 보고서 페이지를 바로 엽니다. ")
        human.goto(page, f"{sc_nav.SC_BASE}{BR_URL}#/report?id={REPORT_ID.replace(':', '%3A')}", wait_after=False)
    sc_nav._settle(page)
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    human.pause()


def _kat_value(el) -> str:
    try:
        return el.evaluate("e => e.value || e.getAttribute('value') || ''") or ""
    except Exception:
        return ""


def _select_custom(page) -> None:
    """Date 드롭다운(kat-dropdown: Yesterday / Week to date / … / Custom)에서 Custom 선택."""
    dd = page.locator("kat-dropdown").filter(has=page.locator('kat-option[value="CUSTOM"]')).first
    if not dd.count():
        print("   Date 드롭다운을 찾지 못해 날짜를 바로 설정합니다. ")
        return
    if _kat_value(dd) == "CUSTOM":
        return
    option = dd.locator('kat-option[value="CUSTOM"]').first
    header = dd.locator(".select-header").first
    for attempt in range(1, 4):
        print("   Date 드롭다운 여는 중..." if attempt == 1 else f"   Custom 다시 선택( {attempt}번째)...")
        human.click(page, header if header.count() else dd)
        try:
            option.wait_for(state="visible", timeout=5000)
        except Exception:
            human.pause()
            continue
        print("   ‘Custom’ 선택...")
        human.click(page, option, wait_after=False)
        for _ in range(10):
            if _kat_value(dd) == "CUSTOM":
                human.pause()
                return
            page.wait_for_timeout(500)
    path = _dump(page, "date_dropdown")
    raise RuntimeError(f"Date 드롭다운에서 Custom을 선택하지 못했습니다(페이지 구조 저장: {path.name})")


def _set_picker(page, picker, d: date, name: str) -> None:
    want = _fmt_us(d)
    if _kat_value(picker) == want:
        return
    inp = picker.locator("input").first
    for attempt in range(1, 4):
        print(f"   입력: {name} {want}...")
        human.type_text(page, inp, want, wait_after=False)
        page.keyboard.press("Tab")
        page.wait_for_timeout(600)
        if _kat_value(picker) == want or (inp.input_value() if inp.count() else "") == want:
            page.keyboard.press("Escape")   # 뜰 수 있는 달력 닫기
            human.pause(1.0, 2.0)
            return
    path = _dump(page, "date_input")
    raise RuntimeError(f"{name} 입력에 실패했습니다(페이지 구조 저장: {path.name})")


def set_dates(page, start: date, end: date) -> None:
    """Date를 Custom으로 → 시작·끝 날짜 입력 → Apply."""
    _select_custom(page)
    pickers = page.locator("kat-date-picker")
    if pickers.count() < 2:
        path = _dump(page, "no_date_picker")
        raise RuntimeError(f"페이지에서 시작 / 종료 날짜 입력란을 찾지 못했습니다(페이지 구조 저장: {path.name})")
    _set_picker(page, pickers.nth(0), start, "시작일")
    _set_picker(page, pickers.nth(1), end, "종료일")
    apply_ = sc_nav._first_visible(page.locator('kat-button[label="Apply"], kat-button[label="적용"], kat-button[label="应用"]'))
    if apply_ is None:
        apply_ = sc_nav.find_visible(page, ["Apply", "적용", "应用"])
    if apply_ is None:
        path = _dump(page, "no_apply")
        raise RuntimeError(f"Apply 버튼을 찾지 못했습니다(페이지 구조 저장: {path.name})")
    print("   Apply 클릭...")
    human.click(page, apply_, wait_after=False)


def _table_asins(page) -> int:
    try:
        return len(set(ASIN_RE.findall(page.inner_text("body"))))
    except Exception:
        return 0


def wait_table(page, start: date, end: date, timeout_s: int = 90) -> int:
    """표에 데이터(ASIN)가 나올 때까지 기다림. 보이는 ASIN 수를 돌려줌(없으면 0)."""
    t0 = time.time()
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    n = 0
    while time.time() - t0 < timeout_s:
        n = _table_asins(page)
        if n:
            break
        page.wait_for_timeout(2000)
    pk = page.locator("kat-date-picker")
    shown = [_kat_value(pk.nth(i)) for i in range(min(2, pk.count()))]
    if shown and shown != [_fmt_us(start), _fmt_us(end)]:
        print(f"   ⚠ 페이지의 날짜가 {' ~ '.join(shown)}로, 원하는 날짜 {_fmt_us(start)} ~ {_fmt_us(end)}와(과) 다릅니다. ")
    human.read_page(page, max_screens=2)
    human.move_away(page)
    print(f"   표에 ASIN {n}개가 있습니다. " if n else "   표에 데이터가 없습니다. ")
    return n


def _download_button(page):
    for loc in (page.locator('kat-button[label^="Download" i]'),
                page.get_by_role("button", name=re.compile(r"^\s*download", re.I)),
                page.get_by_text(DOWNLOAD_RE)):
        el = sc_nav._first_visible(loc)
        if el is not None:
            return el
    return None


def fetch_range(page, start: date, end: date, save_name: str) -> Tuple[Path, int]:
    """하위 ASIN 상세 페이지 판매 및 트래픽 보고서를 열고 날짜 선택(Custom → 시작 / 끝 → Apply), 표가 나오면 CSV 다운로드."""
    try:
        if "/business-reports" not in page.url:
            open_workspace(page)
            open_business_reports(page)
        open_child_report(page)
        print(f"   날짜 설정 {start} ~ {end}...")
        set_dates(page, start, end)
        n = wait_table(page, start, end)
        _dump(page, f"{save_name}_page")   # 실제 페이지에 맞춰 조정할 수 있게 페이지 구조 저장
        btn = _download_button(page)
        if btn is None:
            path = _dump(page, f"{save_name}_no_download")
            raise RuntimeError(f"페이지에서 Download (.csv) 버튼을 찾지 못했습니다(페이지 구조 저장: {path.name})")
        print("   Download (.csv) 클릭...")
        return sc_reports._download(page, btn, RAW_DIR, save_name, DEBUG_DIR), n
    except Exception:
        _dump(page, save_name)
        raise


RECENT_DAYS = 3   # 끝 날짜가 최근 3일 안이면 데이터가 아직 다 안 들어왔을 수 있음: 그날은 그날 받은 것을 쓰고 다음 날 다시 받음


def get_range(page_factory, start: date, end: date, today: date) -> Tuple[Path, str]:
    """한 기간의 데이터: 이미 받았으면 그대로 사용. 최근 며칠은 Amazon이 아직 갱신 중이라 받은 날짜별로 따로 저장하고 하루 한 번만 받음."""
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    recent = (today - end).days <= RECENT_DAYS
    out = REPORT_DIR / (f"business_data_{start}_{end}" + (f"_수집일{today}" if recent else "") + ".xlsx")
    if out.exists():
        print(f"   {start} ~ {end} 데이터는 이미 다운로드했습니다({out.name}). 그대로 사용합니다. ")
        return out, f"{out.name}(이미 다운로드함, 그대로 사용)"
    import sp_api
    import sp_data
    api = sp_api.try_api(f"비즈니스 보고서 {start}~{end}", lambda: sp_data.sales_traffic(start, end, sp_data.listing_titles()))
    if api is not None:
        h, rows = api
        inv.save_report_xlsx(h, rows, out, "Child ASIN")
        print(f"   ✅ {start} ~ {end}(API): {len(rows)}개 ASIN")
        return out, f"{out.name}(SP-API, {len(rows)}행)"
    with page_factory() as page:
        raw, shown = fetch_range(page, start, end, f"business_{start}_{end}")
    h, rows = inv.read_report(raw)
    if not h:
        raise RuntimeError(f"다운로드한 파일이 비어 있습니다: {raw.name}")
    if not rows and shown:
        raise RuntimeError(f"페이지 표에는 ASIN {shown}개가 있는데, 다운로드한 파일에 데이터가 없습니다: {raw.name}")
    inv.save_report_xlsx(h, rows, out, "Child ASIN")
    if rows:
        print(f"   ✅ {start} ~ {end}: {len(rows)}개 ASIN")
        return out, f"{out.name}({len(rows)}행)"
    print(f"   ⚠ {start} ~ {end}: 셀러 센트럴에 이 기간의 데이터가 없습니다(아직 업데이트되지 않았을 수 있음). ")
    return out, f"{out.name}(이 기간은 셀러 센트럴에 데이터 없음)"


# ================================================================ 데이터 정리

def _num(v) -> float:
    if v in (None, ""):
        return 0.0
    s = str(v).replace(",", "").replace("$", "").replace("%", "").replace("US", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _col(row: Dict, *names):
    """열 이름으로 값 찾기: 먼저 정확히(대소문자·기호 무시), 그다음 ‘포함’."""
    idx = {inv.norm(k): k for k in row}
    for n in names:
        k = idx.get(inv.norm(n))
        if k is not None:
            return row[k]
    for n in names:
        nn = inv.norm(n)
        for nk, k in idx.items():
            if nn in nk and "b2b" not in nk:
                return row[k]
    return None


def parse(rows: List[Dict]) -> Dict[str, Dict]:
    out = {}
    for r in rows:
        asin = str(_col(r, "(Child) ASIN", "Child ASIN", "子 ASIN", "하위 ASIN") or "").strip()
        if not asin:
            continue
        sessions = _num(_col(r, "Sessions - Total", "Sessions", "会话数 - 总计", "세션 - 합계"))
        units = _num(_col(r, "Units Ordered", "已订购商品数量", "주문된 수량"))
        conv = _col(r, "Unit Session Percentage", "商品会话百分比", "단위 세션 비율")
        out[asin] = {
            "asin": asin,
            "parent": str(_col(r, "(Parent) ASIN", "Parent ASIN", "父 ASIN", "상위 ASIN") or "").strip(),
            "title": str(_col(r, "Title", "标题", "상품명") or "").strip(),
            "sku": str(_col(r, "SKU") or "").strip(),
            "sessions": sessions,
            "page_views": _num(_col(r, "Page Views - Total", "Page Views", "页面浏览量 - 总计", "페이지뷰 - 합계")),
            "buybox": _num(_col(r, "Featured Offer Percentage", "Featured Offer (Buy Box) Percentage", "Buy Box Percentage",
                                "推荐报价（购买按钮）百分比", "추천 제안(바이 박스) 비율")),
            "units": units,
            "sales": _num(_col(r, "Ordered Product Sales", "已订购商品销售额", "주문 상품 판매량")),
            "items": _num(_col(r, "Total Order Items", "订单商品总数", "총 주문 아이템")),
            "conv": _num(conv) if conv not in (None, "") else (units / sessions * 100 if sessions else 0.0),
        }
    return out


def totals(d: Dict[str, Dict]) -> Dict:
    s = sum(x["sessions"] for x in d.values())
    u = sum(x["units"] for x in d.values())
    pv = sum(x["page_views"] for x in d.values())
    bb = (sum(x["buybox"] * x["page_views"] for x in d.values()) / pv) if pv else 0.0
    return {"세션 수": s, "페이지뷰": pv, "주문 수량": u,
            "주문 매출": round(sum(x["sales"] for x in d.values()), 2),
            "전환율": (u / s * 100) if s else 0.0, "바이박스 비율": bb,
            "트래픽이 있는 ASIN": sum(1 for x in d.values() if x["sessions"] > 0)}


def analyze(cur: Dict[str, Dict], prev: Dict[str, Dict], days: int) -> Dict[str, List[Dict]]:
    t = totals(cur)
    avg = t["전환율"]
    min_s = days * MIN_SESSIONS_PER_DAY
    with_traffic = sorted((x["sessions"] for x in cur.values() if x["sessions"] > 0))
    median = with_traffic[len(with_traffic) // 2] if with_traffic else 0
    p75 = with_traffic[int(len(with_traffic) * 0.75)] if with_traffic else 0

    res = {"high_traffic_low_conv": [], "high_conv_low_traffic": [], "low_buybox": [], "drops": []}
    for a, x in cur.items():
        p = prev.get(a)
        if x["sessions"] >= max(min_s, p75) and avg and x["conv"] < avg * LOW_CONV_RATIO:
            res["high_traffic_low_conv"].append({**x, "avg": avg})
        if x["sessions"] >= max(3, days) and x["sessions"] <= median and avg and x["conv"] >= avg * HIGH_CONV_RATIO \
                and x["units"] >= 2:
            res["high_conv_low_traffic"].append({**x, "avg": avg})
        if x["page_views"] > 0 and x["buybox"] < BUYBOX_ALERT:
            res["low_buybox"].append(x)
        if p and p["sessions"] >= min_s:
            s_chg = (x["sessions"] / p["sessions"] - 1) if p["sessions"] else 0
            c_chg = (x["conv"] / p["conv"] - 1) if p["conv"] else 0
            why = []
            if s_chg <= -DROP_ALERT:
                why.append(f"세션 수 {p['sessions']:,.0f} → {x['sessions']:,.0f}({s_chg:+.0%})")
            if p["conv"] and x["sessions"] >= min_s and c_chg <= -DROP_ALERT:
                why.append(f"전환율 {p['conv']:.1f}% → {x['conv']:.1f}%({c_chg:+.0%})")
            if why:
                res["drops"].append({**x, "prev": p, "why": "; ".join(why), "s_chg": s_chg, "c_chg": c_chg})
    # 이전 기간엔 트래픽이 있었는데 이번 기간엔 아예 없는 ASIN
    for a, p in prev.items():
        if a not in cur and p["sessions"] >= min_s:
            gone = {"asin": a, "parent": p["parent"], "title": p["title"], "sku": p["sku"], "sessions": 0,
                    "page_views": 0, "buybox": 0, "units": 0, "sales": 0, "items": 0, "conv": 0}
            res["drops"].append({**gone, "prev": p, "why": f"이번 기간 데이터 없음(이전 기간 세션 수 {p['sessions']:,.0f})",
                                 "s_chg": -1, "c_chg": 0})
    res["high_traffic_low_conv"].sort(key=lambda x: -x["sessions"])
    res["high_conv_low_traffic"].sort(key=lambda x: -x["conv"])
    res["low_buybox"].sort(key=lambda x: (x["buybox"], -x["page_views"]))
    res["drops"].sort(key=lambda x: (x["s_chg"], x["c_chg"]))
    return res


# ================================================================ 워크스페이스 스냅숏(하루 한 번)

def _ws_file(today: date) -> Path:
    return REPORT_DIR / f"workspace_{today}.json"


def load_workspace(today: date) -> Optional[Dict]:
    try:
        return json.loads(_ws_file(today).read_text(encoding="utf-8"))
    except Exception:
        return None


def snapshot_workspace(page, today: date) -> Dict:
    open_workspace(page)
    ws = read_workspace(page)
    _dump(page, "workspace")      # 실제 페이지에 맞춰 읽기를 조정할 수 있게 페이지 구조 저장
    ws.pop("raw", None)
    ws["time"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    ws["url"] = page.url
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    _ws_file(today).write_text(json.dumps(ws, ensure_ascii=False, indent=2), encoding="utf-8")
    n = len(ws.get("conversion", [])) + len(ws.get("queries", [])) + len(ws.get("promotions", [])) + len(ws.get("segments", []))
    print(f"   ✅ 마케팅 워크스페이스: {n}개 항목을 읽었습니다")
    return ws


# ================================================================ Excel 보고서

def write_report(out: Path, label: str, cur_rng, prev_rng, cur: Dict[str, Dict], prev: Dict[str, Dict],
                 ws_data: Optional[Dict], sources: Dict[str, str], notes: List[str]) -> Dict:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
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
    days = (cur_rng[1] - cur_rng[0]).days + 1

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

    def note(ws, r, text, ncol, color=GREY):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        cell(ws, r, 1, text, f(9, color=color), border=None, wrap=True)
        ws.row_dimensions[r].height = 30

    rng_txt = lambda r: f"{r[0]:%m/%d} ~ {r[1]:%m/%d}" if r[0] != r[1] else f"{r[0]:%m/%d}"
    t_cur, t_prev = totals(cur), totals(prev)
    res = analyze(cur, prev, days)

    wb = Workbook()

    # ================================================= 개요
    ws = wb.active
    ws.title = "요약"
    ws.sheet_view.showGridLines = False
    for c, w in zip("ABCDE", (24, 18, 18, 16, 14)):
        ws.column_dimensions[c].width = w
    cell(ws, 1, 1, f"비즈니스 보고서 · {label}", f(18, True, NAVY), border=None)
    cell(ws, 2, 1, f"이번 기간 {cur_rng[0]} ~ {cur_rng[1]}   비교: 이전 기간 {prev_rng[0]} ~ {prev_rng[1]}(미국 중부 시간) "
                   f"생성: {datetime.now():%Y-%m-%d %H:%M}", f(9.5, color=GREY), border=None)

    r = 4
    cell(ws, r, 1, "핵심 지표(상세 페이지 판매 및 트래픽, 하위 ASIN 합계)", f(13, True, NAVY), border=None)
    r += 1
    header(ws, r, ["지표", f"이번 기간 {rng_txt(cur_rng)}", f"이전 기간 {rng_txt(prev_rng)}", "변화", "변화율"])
    for k in ("세션 수", "페이지뷰", "주문 수량", "주문 매출", "전환율", "바이박스 비율", "트래픽이 있는 ASIN"):
        r += 1
        rate = k in ("전환율", "바이박스 비율")
        money = k == "주문 매출"
        num = "0.00%" if rate else (MONEY if money else "#,##0")
        a, b = t_cur[k], t_prev[k]
        if rate:
            a, b = a / 100, b / 100
        cell(ws, r, 1, k, f(10, True))
        cell(ws, r, 2, a, num=num)
        cell(ws, r, 3, b, num=num)
        cell(ws, r, 4, f"=B{r}-C{r}", num='+0.00%;[Red]-0.00%;0.00%' if rate else (MONEY_CHG if money else NUM_CHG))
        cell(ws, r, 5, f'=IF(C{r}=0,"",B{r}/C{r}-1)', num=PCT_CHG)

    # ---- 요점
    r += 2
    cell(ws, r, 1, "요점", f(13, True, NAVY), border=None)
    points = []

    def chg(a, b):
        return (a / b - 1) if b else None
    cs_, cc_ = chg(t_cur["세션 수"], t_prev["세션 수"]), chg(t_cur["전환율"], t_prev["전환율"])
    txt = f"세션 수 {t_cur['세션 수']:,.0f}"
    if cs_ is not None:
        txt += f"(이전 기간 대비 {'증가' if cs_ >= 0 else '감소'} {abs(cs_):.1%})"
    txt += f", 전환율 {t_cur['전환율']:.2f}%"
    if cc_ is not None:
        txt += f"(이전 기간 {t_prev['전환율']:.2f}%)"
    txt += f", 주문 매출 ${t_cur['주문 매출']:,.2f}. "
    points.append((RED if (cs_ or 0) < -0.1 or (cc_ or 0) < -0.1 else (GREEN if (cs_ or 0) > 0.1 else None), txt))
    if res["high_traffic_low_conv"]:
        points.append((ORANGE, f"트래픽은 높고 전환이 낮은 ASIN {len(res['high_traffic_low_conv'])}개: 가격, 이미지, 리뷰 또는 상세 설명에 문제가 있을 수 있습니다. "
                               f"‘트래픽↑전환↓’ 시트를 참고하세요. "))
    if res["high_conv_low_traffic"]:
        points.append((GREEN, f"전환은 높고 트래픽이 낮은 ASIN {len(res['high_conv_low_traffic'])}개: 광고를 추가하거나 키워드를 최적화해 유입을 늘려 보세요. "
                              f"‘전환↑트래픽↓’ 시트를 참고하세요. "))
    if res["low_buybox"]:
        points.append((RED, f"바이박스 비율이 {BUYBOX_ALERT:.0f}% 미만인 ASIN {len(res['low_buybox'])}개: "
                            f"가격, 재고, 다른 판매자의 합승 여부를 확인하세요. ‘바이박스 90% 미만’ 시트를 참고하세요. "))
    if res["drops"]:
        points.append((RED, f"트래픽 또는 전환이 이전 기간보다 크게 하락한({DROP_ALERT:.0%} 이상) ASIN {len(res['drops'])}개, ‘크게 하락’ 시트를 참고하세요. "))
    if len(points) == 1:
        points.append((GREEN, "특별히 주의가 필요한 ASIN이 없습니다 ✓"))
    for color, text in points:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell(ws, r, 1, "• " + text, f(10.5, color=color), border=None, wrap=True)
        ws.row_dimensions[r].height = 30 if len(text) > 50 else 18

    # ---- 마케팅 워크스페이스
    if ws_data:
        r += 2
        cell(ws, r, 1, f"마케팅 워크스페이스(읽은 시간 {ws_data.get('time', '')})", f(13, True, NAVY), border=None)
        blocks = [("전환율 카드", ws_data.get("conversion", [])),
                  ("상단 지표", [(f"{a} · {b}", c) for a, b, c in ws_data.get("tiles", [])]),
                  ("프로모션(최근 7일)", [(f"{a} · {b}", c) for a, b, c in ws_data.get("promotions", [])]),
                  ("고객 세그먼트", ws_data.get("segments", []))]
        any_ = False
        for title, items in blocks:
            if not items:
                continue
            any_ = True
            r += 1
            cell(ws, r, 1, title, f(10.5, True), border=None)
            for k, v in items:
                r += 1
                cell(ws, r, 1, k)
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
                cell(ws, r, 2, v)
        if ws_data.get("queries"):
            any_ = True
            r += 1
            cell(ws, r, 1, f"인기 검색어 상위 {min(5, len(ws_data['queries']))}개(전체 목록은 ‘인기 검색어’ 시트)", f(10.5, True), border=None)
            for q in ws_data["queries"][:5]:
                r += 1
                cell(ws, r, 1, f"#{q[0]} {q[1]}")
                cell(ws, r, 2, f"클릭률 {q[2]}")
                cell(ws, r, 3, f"구매율 {q[3]}")
        if not any_:
            r += 1
            note(ws, r, "워크스페이스 페이지에서 확인할 수 있는 숫자를 읽지 못했습니다(페이지 구조는 debug에 저장되었으니 개발자에게 보내 주시면 조정합니다). ", 5, ORANGE)
        r += 1
        note(ws, r, "워크스페이스 카드의 데이터는 아마존 자체 기준(예: ‘최근 30일’)으로 집계되어, 위의 이번 기간 / 이전 기간 날짜와 다를 수 있습니다. 참고용입니다. ", 5)

    r += 2
    cell(ws, r, 1, "데이터 출처", f(13, True, NAVY), border=None)
    for k, v in sources.items():
        r += 1
        cell(ws, r, 1, k, f(10, True), border=None)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        cell(ws, r, 2, v, f(9.5, color=GREY), border=None)
    for n in notes:
        r += 1
        note(ws, r, n, 5)

    # ================================================= 상세 표
    def asin_sheet(name, items, extra_cols, extra_vals, tip, widths_extra):
        sh = wb.create_sheet(name)
        cols = ["하위 ASIN", "상위 ASIN", "SKU", "상품명", "세션 수", "페이지뷰", "주문 수량", "주문 매출", "전환율",
                "바이박스 비율"] + extra_cols
        header(sh, 1, cols, [13, 13, 18, 40, 10, 11, 10, 12, 9, 11] + widths_extra)
        sh.freeze_panes = "B2"
        for i, x in enumerate(items, start=2):
            vals = [x["asin"], x["parent"], x["sku"], x["title"], x["sessions"], x["page_views"], x["units"],
                    round(x["sales"], 2), x["conv"] / 100, x["buybox"] / 100] + extra_vals(x, i)
            for c, v in enumerate(vals, start=1):
                num = {5: "#,##0", 6: "#,##0", 7: "#,##0", 8: MONEY, 9: "0.00%", 10: "0.0%"}.get(c)
                if c > 10:
                    num = extra_num.get(cols[c - 1])
                cell(sh, i, c, v, num=num, wrap=(c == len(cols) and isinstance(v, str) and len(v) > 30))
        if items:
            sh.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(items) + 1}"
        nr = len(items) + 3
        if not items:
            cell(sh, 2, 1, "이번 기간 해당 없음 ✓", f(10, color=GREEN), border=None)
        note(sh, nr, tip, min(len(cols), 10))
        return sh

    extra_num = {"전체 평균 전환율": "0.00%", "이전 기간 세션 수": "#,##0", "세션 수 변화": PCT_CHG, "이전 기간 전환율": "0.00%",
                 "전환율 변화": PCT_CHG}
    avg = t_cur["전환율"]
    asin_sheet("트래픽↑전환↓", res["high_traffic_low_conv"], ["전체 평균 전환율", "제안"],
               lambda x, i: [avg / 100, "가격, 메인 이미지, 리뷰, 불릿 포인트를 확인하세요. 광고가 부정확한 트래픽을 가져오는지도 확인하세요"],
               f"세션 수가 전체 상위 25%이고 최소 {days * MIN_SESSIONS_PER_DAY}개 이상이며, 전환율 < 전체 평균 × {LOW_CONV_RATIO:.0%}. "
               f"임계값은 business.py 맨 위에서 수정할 수 있습니다. ", [12, 40])
    asin_sheet("전환↑트래픽↓", res["high_conv_low_traffic"], ["전체 평균 전환율", "제안"],
               lambda x, i: [avg / 100, "전환은 좋지만 트래픽이 적습니다: 광고 예산을 늘리거나 제목 키워드를 최적화하세요"],
               f"세션 수가 중앙값 이하이고, 전환율이 전체 평균의 {HIGH_CONV_RATIO:g}배 이상이며, 주문이 최소 2개인 ASIN. ", [12, 36])
    asin_sheet("바이박스 90% 미만", res["low_buybox"], ["제안"],
               lambda x, i: ["가격 경쟁력, 재고 충분 여부, 다른 판매자의 합승 여부를 확인하세요"],
               f"바이박스 비율이 {BUYBOX_ALERT:.0f}% 미만(페이지뷰가 있는 ASIN만). ", [40])
    asin_sheet("크게 하락", res["drops"], ["이전 기간 세션 수", "세션 수 변화", "이전 기간 전환율", "전환율 변화", "사유"],
               lambda x, i: [x["prev"]["sessions"], f'=IF(K{i}=0,"",E{i}/K{i}-1)', x["prev"]["conv"] / 100,
                             f'=IF(M{i}=0,"",I{i}/M{i}-1)', x["why"]],
               f"이전 기간 세션 수가 최소 {days * MIN_SESSIONS_PER_DAY}개 이상이고, 이번 기간 세션 수 또는 전환율이 {DROP_ALERT:.0%} 이상 하락. ", [11, 11, 11, 11, 40])

    # ================================================= ASIN 목록(이번 기간 vs 이전 기간)
    al = wb.create_sheet("ASIN 전체")
    cols = ["하위 ASIN", "SKU", "상품명", "이번 기간 세션 수", "이전 기간 세션 수", "세션 수 변화", "이번 기간 전환율", "이전 기간 전환율",
            "이번 기간 수량", "이전 기간 수량", "이번 기간 매출", "이전 기간 매출", "매출 변화", "바이박스 비율", "알림"]
    header(al, 1, cols, [13, 18, 40, 11, 11, 10, 10, 10, 9, 9, 12, 12, 12, 11, 24])
    al.freeze_panes = "B2"
    flags = {}
    for key_, tag in (("high_traffic_low_conv", "트래픽↑전환↓"), ("high_conv_low_traffic", "전환↑트래픽↓"),
                      ("low_buybox", "바이박스 낮음"), ("drops", "크게 하락")):
        for x in res[key_]:
            flags.setdefault(x["asin"], []).append(tag)
    all_asins = sorted(set(cur) | set(prev), key=lambda a: (-(cur.get(a) or {}).get("sessions", 0),
                                                           -(prev.get(a) or {}).get("sessions", 0)))
    zero = {"sessions": 0, "conv": 0, "units": 0, "sales": 0, "buybox": 0, "title": "", "sku": ""}
    for i, a in enumerate(all_asins, start=2):
        x, p = cur.get(a, zero), prev.get(a, zero)
        tips = flags.get(a, [])
        vals = [a, x["sku"] or p["sku"], x["title"] or p["title"], x["sessions"], p["sessions"],
                f'=IF(E{i}=0,"",D{i}/E{i}-1)', x["conv"] / 100, p["conv"] / 100, x["units"], p["units"],
                round(x["sales"], 2), round(p["sales"], 2), f"=K{i}-L{i}", x["buybox"] / 100 if a in cur else "",
                ", ".join(tips)]
        nums = {4: "#,##0", 5: "#,##0", 6: PCT_CHG, 7: "0.00%", 8: "0.00%", 11: MONEY, 12: MONEY, 13: MONEY_CHG, 14: "0.0%"}
        for c, v in enumerate(vals, start=1):
            cell(al, i, c, v, f(10, c == 15, RED if (c == 15 and tips) else None), num=nums.get(c))
    if all_asins:
        al.auto_filter.ref = f"A1:O{len(all_asins) + 1}"
        n = min(15, len(all_asins))
        bc = BarChart()
        bc.type, bc.title, bc.height, bc.width = "bar", f"세션 수 상위 {n} ASIN: 이번 기간 vs 이전 기간", 10, 18
        bc.add_data(Reference(al, min_col=4, max_col=5, min_row=1, max_row=n + 1), titles_from_data=True)
        bc.set_categories(Reference(al, min_col=1, min_row=2, max_row=n + 1))
        bc.x_axis.scaling.orientation = "maxMin"
        bc.y_axis.majorGridlines = None
        al.add_chart(bc, "Q2")

    # ================================================= 인기 검색어(워크스페이스)
    if ws_data and ws_data.get("queries"):
        qs = wb.create_sheet("인기 검색어")
        header(qs, 1, ["순위", "검색어", "클릭률", "구매율"], [7, 40, 12, 12])
        for i, q in enumerate(ws_data["queries"], start=2):
            for c, v in enumerate(q, start=1):
                cell(qs, i, c, v, align="center" if c != 2 else None)
        note(qs, len(ws_data["queries"]) + 3, f"마케팅 워크스페이스 Top Search Queries {ws_data.get('queries_range', '')}"
                                             f"(브랜드 분석 데이터, 읽은 시간 {ws_data.get('time', '')}). ", 4)

    for sheet in wb.worksheets:
        sheet.page_setup.orientation = "portrait" if sheet.title == "요약" else "landscape"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0

    wb.save(out)
    return {"cur": t_cur, "prev": t_prev, "res": res, "points": [p[1] for p in points]}


# ================================================================ 메인 흐름

def choose_period() -> Optional[Tuple[str, str, int]]:
    print("   날짜 범위를 선택하세요(미국 중부 시간, 네 가지 모두 같은 길이의 이전 기간과 비교): ")
    for k, (_, label, days) in PERIODS.items():
        (cs, ce), (ps_, pe) = ordm.period_ranges(days)
        rng = f"{cs}" if cs == ce else f"{cs} ~ {ce}"
        prv = f"{ps_}" if ps_ == pe else f"{ps_} ~ {pe}"
        print(f"     {k}. {label}: {rng}(비교 {prv})")
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
    today = ordm.central_today()
    cur_rng, prev_rng = ordm.period_ranges(days, today)
    print(f"   이번 기간 {cur_rng[0]} ~ {cur_rng[1]}, 이전 기간 {prev_rng[0]} ~ {prev_rng[1]}(미국 중부 시간, 오늘은 {today})")

    from contextlib import ExitStack
    stack = ExitStack()
    holder = {}

    def page_factory():
        class _Ctx:
            def __enter__(self_):
                if "page" not in holder:
                    holder["page"] = stack.enter_context(open_session(headless=False))
                return holder["page"]

            def __exit__(self_, *a):
                return False
        return _Ctx()

    sources, notes = {}, []
    ws_data = load_workspace(today)
    try:
        if ws_data is None:
            try:
                with page_factory() as page:
                    ws_data = snapshot_workspace(page, today)
                sources["마케팅 워크스페이스"] = f"메뉴 마케팅 → 워크스페이스({ws_data['time']} 읽음)"
            except Exception as e:
                msg = f"마케팅 워크스페이스를 읽지 못했습니다({type(e).__name__}: {str(e).splitlines()[0][:150]}). 비즈니스 보고서에는 영향이 없습니다. "
                print("   ⚠", msg)
                notes.append(msg)
        else:
            print(f"   오늘 이미 마케팅 워크스페이스를 읽었습니다({ws_data.get('time')}). 그대로 사용합니다. ")
            sources["마케팅 워크스페이스"] = f"메뉴 마케팅 → 워크스페이스(오늘 {ws_data.get('time')} 읽음, 그대로 사용)"
        data = {}
        for name, (s, e) in (("이번 기간", cur_rng), ("이전 기간", prev_rng)):
            try:
                path, src = get_range(page_factory, s, e, today)
            except Exception as ex:
                raise RuntimeError(f"{name}({s} ~ {e}) 비즈니스 보고서 가져오기 실패: {type(ex).__name__}: "
                                   f"{str(ex).splitlines()[0][:200]}(스크린샷과 페이지 구조는 reports/business/debug에 있음)") from ex
            data[name] = parse(inv.read_xlsx(path)[1])
            sources[f"{name} 비즈니스 보고서"] = f"상세 페이지 판매 및 트래픽(하위 ASIN별) {s} ~ {e}: {src}"
    finally:
        stack.close()

    for name, (s, e) in (("이번 기간", cur_rng), ("이전 기간", prev_rng)):
        if not data[name]:
            notes.append(f"{name} {s} ~ {e} 셀러 센트럴에 데이터가 없습니다(최근 1~2일 데이터는 아마존이 아직 업데이트하지 않았을 수 있음). 비교 결과는 참고용입니다. ")
    notes.append("전환율 = 주문 수량 ÷ 세션 수(Unit Session Percentage). 전체 전환율은 모든 ASIN 합계로 계산하고, "
                 " 바이박스 비율은 페이지뷰 가중 평균입니다. 비즈니스 보고서 데이터는 보통 1~2일 지연되며, 최근 하루는 아직 완전하지 않을 수 있습니다. ")
    out = REPORT_DIR / f"business_report_{key}_{today}.xlsx"
    res = write_report(out, label, cur_rng, prev_rng, data["이번 기간"], data["이전 기간"], ws_data, sources, notes)
    mark_last_run(key, label)
    try:
        import daily_bus
        tc_ = res["cur"]
        daily_bus.record("비즈니스 보고서", {f"{label} {k}": tc_.get(k) for k in ("세션 수", "주문 수량", "주문 매출", "전환율") if k in tc_},
                         daily_bus.points(res.get("points", [])[1:10], "비즈니스 보고서 확인"), out)
    except Exception:
        pass

    print()
    tc, tp = res["cur"], res["prev"]
    for k, fmt in (("세션 수", "{:,.0f}"), ("주문 수량", "{:,.0f}"), ("주문 매출", "${:,.2f}"), ("전환율", "{:.2f}%")):
        a, b = tc[k], tp[k]
        ch = f"{(a / b - 1):+.1%}" if b else "—"
        print(f"{k}: 이번 기간 {fmt.format(a)}, 이전 기간 {fmt.format(b)}({ch})")
    for p in res["points"][1:]:
        print("•", p)
    print(f"\n✅ 보고서 저장 완료: {out}")
    if open_report:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass
    return out


if __name__ == "__main__":
    run()

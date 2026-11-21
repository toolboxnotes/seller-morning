"""
재고(두 가지 독립 기능, 메뉴에서 따로)

  전체 상품 재고  run_listings(): 메뉴 Products → Product reports, All Listings Report
  FBA 재고      run_fba(): 메뉴 Supply chain → Supply chain reports → FBA inventory reports, FBA 재고 보고서

보고서마다 하루 한 번만 받음: 그날 파일이 이미 있으면 그대로 쓰고 관리 화면에 다시 들어가지 않음.
받은 뒤 Excel로 바꾸고 평가가 들어간 일일 보고서를 만듦.

파일은 reports/inventory/ 에 저장:
  all_listings_YYYY-MM-DD.xlsx      전체 상품 데이터
  listings_report_YYYY-MM-DD.xlsx   전체 상품 재고 일일 보고서(평가)
  fba_inventory_YYYY-MM-DD.xlsx     FBA 재고 데이터
  fba_report_YYYY-MM-DD.xlsx        FBA 재고 일일 보고서(평가)
  raw/                              Amazon에서 받은 원본 파일
  debug/                            어떤 단계에서 예상한 내용을 못 찾았을 때의 스크린샷과 페이지 글자
"""

import xlsafe  # noqa: F401  Excel 수식 주입 방지(Amazon에서 받은 글자가 수식으로 실행되지 않게)
import csv
import io
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import human
import sc_labels
import sc_nav
import sc_reports
from sc_browser import open_session

REPORT_DIR = Path(__file__).resolve().parent.parent / "reports" / "inventory"
RAW_DIR = REPORT_DIR / "raw"
DEBUG_DIR = REPORT_DIR / "debug"

# ---------------------------------------------------------------- 페이지 글자(영 / 한 / 중)
ALL_LISTINGS = ["All Listings Report", "모든 리스팅 보고서", "전체 리스팅 보고서", "所有商品报告", "全部商品报告"]
REPORT_TYPE_HINT = re.compile(r"report type|select report|보고서 유형|보고서 선택|报告类型|选择报告", re.I)
REQUEST_REPORT = ["Request Report", "Request report", "보고서 요청", "请求报告"]

INVENTORY_GROUP = ["Inventory", "재고", "库存"]
FBA_REPORT_NAMES = [
    ["FBA Inventory", "FBA 재고", "亚马逊物流库存"],
    ["Manage FBA Inventory", "FBA 재고 관리", "管理亚马逊库存"],
]
FBA_REQUEST = ["Request .txt Download", "Request .txt download", ".txt 다운로드 요청", "请求 .txt 下载",
               "Request .csv Download", "Request .csv download", ".csv 다운로드 요청", "请求 .csv 下载",
               "Request Download", "Generate report", "다운로드 요청", "请求下载"]
FBA_FORMAT = [".txt"]  # 요구 사항대로 .txt 사용(SKU / ASIN / FNSKU 앞의 0이 없어지지 않음)

# ---------------------------------------------------------------- 평가 기준(필요하면 조정)
DOS_URGENT = 14        # 판매 가능 일수가 이보다 적으면: 긴급 보충
DOS_WATCH = 30         # 판매 가능 일수가 이보다 적으면: 보충 주의
AGED_DAYS_LABEL = "180일"


# ================================================================ 브라우저 조작

DROPDOWN_PLACEHOLDER = re.compile(r"^\s*(Select Report Type|Select report type|보고서 유형 선택|선택|选择报告类型)\s*$")

# 요소가 보고서 목록 표의 줄 안에 있는지(목록에도 “All Listings Report”라는 글자가 있어 거기를 누르면 안 됨)
_IN_TABLE_ROW_JS = """
(el) => {
  let n = el;
  for (let i = 0; i < 12 && n; i++) {
    if (n.matches && n.matches('tr, td, [role="row"], [role="cell"], kat-table-row, kat-table-cell, table')) return true;
    n = n.parentElement || (n.getRootNode && n.getRootNode().host) || null;
  }
  return false;
}
"""


# 페이지 구조(웹 컴포넌트 안 포함)를 글자로 — 태그 이름·주요 속성·글자만 남겨 인식 규칙 조정에 씀
_DUMP_JS = """
() => {
  const keep = ['id','class','role','name','type','value','placeholder','label','aria-label','aria-haspopup',
                'aria-expanded','aria-selected','data-value','href','selected'];
  const out = [];
  const walk = (node, depth) => {
    if (out.length > 4000 || depth > 40) return;
    if (node.nodeType === 3) {
      const t = node.textContent.trim();
      if (t) out.push('  '.repeat(depth) + '"' + t.slice(0, 120) + '"');
      return;
    }
    if (node.nodeType === 11) { for (const c of node.childNodes) walk(c, depth); return; }
    if (node.nodeType !== 1) return;
    const tag = node.tagName.toLowerCase();
    if (['script','style','svg','noscript','path'].includes(tag)) return;
    const attrs = keep.filter(a => node.hasAttribute(a))
                      .map(a => a + '="' + String(node.getAttribute(a)).slice(0, 80) + '"').join(' ');
    out.push('  '.repeat(depth) + '<' + tag + (attrs ? ' ' + attrs : '') + '>');
    if (node.shadowRoot) { out.push('  '.repeat(depth + 1) + '#shadow-root'); walk(node.shadowRoot, depth + 2); }
    for (const c of node.childNodes) walk(c, depth + 1);
  };
  // 페이지 윗부분만 내보냄(보고서 목록이 길어서 앞 몇 줄만)
  const main = document.querySelector('main') || document.body;
  walk(main, 0);
  return out.join('\\n');
}
"""


def dump_dom(page, name: str) -> Path:
    from datetime import datetime as _dt
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    path = DEBUG_DIR / f"{_dt.now():%Y%m%d_%H%M%S}_{name}_dom.txt"
    try:
        text = page.evaluate(_DUMP_JS)
    except Exception as e:
        text = f"(내보내기 실패: {e})"
    path.write_text(f"URL: {page.url}\n\n{text}", encoding="utf-8")
    sc_nav.save_debug(page, DEBUG_DIR, name)
    return path


REQUEST_LABEL = re.compile(r"^\s*(Request an Inventory Report|재고 보고서 요청|请求库存报告)\s*$", re.I)

# 어떤 점에서 아래로 찾기: 웹 컴포넌트를 통과해 ‘입력란’ 크기의 요소(너비 > 200, 높이 20~70)를 찾아 위치를 돌려줌
_BOX_BELOW_JS = """
([x, y0, maxDy]) => {
  const deep = (x, y) => {
    let el = document.elementFromPoint(x, y);
    while (el && el.shadowRoot) {
      const inner = el.shadowRoot.elementFromPoint(x, y);
      if (!inner || inner === el) break;
      el = inner;
    }
    return el;
  };
  for (let dy = 4; dy <= maxDy; dy += 6) {
    let el = deep(x, y0 + dy);
    for (let i = 0; i < 8 && el; i++) {
      const r = el.getBoundingClientRect();
      if (r.width > 200 && r.height >= 20 && r.height <= 70) {
        return {x: r.left, y: r.top, width: r.width, height: r.height,
                tag: el.tagName.toLowerCase(), text: (el.innerText || el.value || '').trim().slice(0, 60)};
      }
      el = el.parentElement || (el.getRootNode && el.getRootNode().host) || null;
    }
  }
  return null;
}
"""


def _dropdown_box(page) -> Optional[dict]:
    """“Request an Inventory Report” 제목 바로 아래 드롭다운 위치."""
    label = sc_nav._first_visible(page.get_by_text(REQUEST_LABEL))
    if label is None:
        return None
    label.scroll_into_view_if_needed()
    page.wait_for_timeout(400)
    lb = label.bounding_box()
    if not lb:
        return None
    return page.evaluate(_BOX_BELOW_JS, [lb["x"] + min(150, lb["width"] / 2 + 60), lb["y"] + lb["height"], 70])


def _dropdown_trigger(page):
    """“Request an Inventory Report” 아래의 보고서 유형 드롭다운."""
    # 보이는 “Select Report Type” 상자 자체를 먼저 누름(컴포넌트 바깥 틀이 입력란보다 훨씬 넓을 수 있어 바깥을 무작위로 누르면 빈 곳을 누름)
    cands = (
        page.get_by_text(DROPDOWN_PLACEHOLDER),
        page.get_by_placeholder(DROPDOWN_PLACEHOLDER),
        page.get_by_role("combobox"),
        page.get_by_role("button", name=DROPDOWN_PLACEHOLDER),
        page.get_by_role("listbox"),
        page.locator("kat-dropdown, kat-select, kat-combobox"),
        page.locator("[aria-haspopup=listbox], [aria-haspopup=true], [aria-expanded]"),
        page.locator('[placeholder*="Report Type" i], [label*="Report" i], [data-placeholder*="Report" i]'),
    )
    for loc in cands:
        try:
            n = loc.count()
        except Exception:
            continue
        for i in range(min(n, 10)):
            el = loc.nth(i)
            try:
                if el.is_visible() and not el.evaluate(_IN_TABLE_ROW_JS):
                    return el
            except Exception:
                continue
    return None


def _visible_option(page, labels: List[str], trigger_box: Optional[dict]):
    """드롭다운 목록의 옵션: role=option 먼저, 없으면 드롭다운 아래·표 밖·드롭다운에서 가장 가까운 것."""
    for lab in labels:
        el = sc_nav._first_visible(page.get_by_role("option", name=lab, exact=True))
        if el is None:
            el = sc_nav._first_visible(page.locator("kat-option, [role=option], li").filter(has_text=lab))
        if el is not None and not el.evaluate(_IN_TABLE_ROW_JS):
            return el
    best, best_d = None, None
    for lab in labels:
        loc = page.get_by_text(lab, exact=True)
        for i in range(min(loc.count(), 60)):
            el = loc.nth(i)
            try:
                if not el.is_visible() or el.evaluate(_IN_TABLE_ROW_JS):
                    continue
                box = el.bounding_box()
            except Exception:
                continue
            if not box:
                continue
            if trigger_box:
                below = box["y"] - (trigger_box["y"] + trigger_box["height"])
                if below < -5 or below > 900:
                    continue
                d = below
            else:
                d = box["y"]
            if best_d is None or d < best_d:
                best, best_d = el, d
    return best


def _request_button(page):
    """보이는 “Request Report” 버튼. 보고서 유형을 골라야만 나타나서 선택됐는지 판단하는 믿을 만한 근거."""
    el = sc_nav._first_visible(page.locator('kat-button[label*="Request Report" i], kat-button[label*="보고서 요청"]'))
    if el is None:
        el = sc_nav.find_visible(page, REQUEST_REPORT)
    if el is None:
        el = sc_nav._first_visible(page.get_by_role("button", name=re.compile(r"^\s*Request Report\s*$", re.I)))
    return el


def _selected_ok(page, trigger, labels: List[str]) -> bool:
    """
    선택됐는지 확인. 드롭다운 컴포넌트의 전체 글자를 보면 안 됨 — 컴포넌트 안에 원래 모든 옵션(All Listings Report 포함)이 들어 있음.
    두 가지만 봄: 컴포넌트의 현재 값(value)이 이 보고서인지, “Request Report” 버튼이 나타났는지.
    """
    if _request_button(page) is not None:
        return True
    try:
        val = trigger.evaluate("(el) => { let n = el; for (let i = 0; i < 6 && n; i++) {"
                               " if (typeof n.value === 'string' && n.value) return n.value;"
                               " n = n.parentElement || (n.getRootNode && n.getRootNode().host) || null; }"
                               " return ''; }")
        return bool(val) and any(l.lower() in val.lower() for l in labels)
    except Exception:
        return False


KATAL_DROPDOWN = 'kat-dropdown[label*="Inventory Report" i], kat-dropdown[placeholder*="Report Type" i]'


def _katal_value(dd) -> str:
    try:
        return dd.evaluate("el => String(el.value || el.getAttribute('value') || '')")
    except Exception:
        return ""


def _select_katal(page, labels: List[str]) -> bool:
    """
    보고서 유형 드롭다운은 Amazon의 kat-dropdown 컴포넌트:
      컴포넌트 안의 .select-header(보이는 “Select Report Type” 상자) 클릭 → 목록의 kat-option → 선택 확인.
    페이지에 이 컴포넌트가 없으면 False를 돌려 일반 방법으로 처리.
    """
    dd = page.locator(KATAL_DROPDOWN).first
    if not dd.count():
        return False
    header = dd.locator(".select-header").first
    name_re = re.compile(r"^\s*(" + "|".join(re.escape(l) for l in labels) + r")\s*$", re.I)
    option = dd.locator("kat-option").filter(has=page.locator(".standard-option-name", has_text=name_re)).first
    if not option.count():
        option = dd.locator("kat-option").filter(has_text=name_re).first
    if not option.count():
        raise RuntimeError(f"드롭다운에 “{labels[0]}” 옵션이 없습니다")
    target_value = option.get_attribute("value") or ""

    for attempt in range(1, 4):
        if (target_value and _katal_value(dd) == target_value) or _request_button(page) is not None:
            return True
        print("   보고서 유형 드롭다운 여는 중..." if attempt == 1 else f"   보고서 유형 다시 선택( {attempt}번째)...")
        human.click(page, header if header.count() else dd)
        try:
            option.wait_for(state="visible", timeout=5000)
        except Exception:
            print("     드롭다운 목록이 펼쳐지지 않아 잠시 후 다시 시도합니다. ")
            human.pause()
            continue
        option.scroll_into_view_if_needed()   # 목록이 길면 옵션이 목록의 보이는 범위에 들어오게
        page.wait_for_timeout(300)
        print(f"   선택: “{labels[0]}”...")
        human.click(page, option, wait_after=False)
        for _ in range(10):  # 최대 5초 기다림
            if (target_value and _katal_value(dd) == target_value) or _request_button(page) is not None:
                human.pause()
                return True
            page.wait_for_timeout(500)
    path = dump_dom(page, "report_type_dropdown")
    raise RuntimeError(f"드롭다운에서 “{labels[0]}”을(를) 선택하지 못했습니다(페이지 구조 저장: {path.name})")


def _select_report_type(page, labels: List[str]) -> None:
    """“Request an Inventory Report” 드롭다운에서 보고서 유형 선택: 드롭다운 열기 → 목록 옵션 클릭 → 선택 확인."""
    # 기본 <select>
    for i in range(page.locator("select").count()):
        sel = page.locator("select").nth(i)
        options = [o.strip() for o in sel.locator("option").all_inner_texts()]
        for lab in labels:
            if lab in options and sel.is_visible():
                box = sel.bounding_box()
                if box:
                    human.move_mouse(page, box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                sel.select_option(label=lab)
                human.pause()
                return

    # Amazon Katal 드롭다운(kat-dropdown): 실제 페이지 구조대로 바로 조작
    if _select_katal(page, labels):
        return

    # 그 밖의 경우: “Request an Inventory Report” 아래 입력란 클릭 → 목록 옵션 클릭 → 확인
    for attempt in range(1, 4):
        if _request_button(page) is not None:
            return
        box = _dropdown_box(page)
        trigger = None
        if box is None:
            trigger = _dropdown_trigger(page)
            if trigger is None:
                path = dump_dom(page, "report_type_dropdown")
                raise RuntimeError(f"보고서 유형 드롭다운을 찾지 못했습니다(페이지 구조 저장: {path.name})")
            box = trigger.bounding_box()
        print("   보고서 유형 드롭다운 여는 중..." if attempt == 1 else f"   보고서 유형 다시 선택( {attempt}번째)...")
        if trigger is not None:
            human.click(page, trigger)
        else:
            human.click_box(page, box)
        option = _visible_option(page, labels, box)
        if option is None:
            print("     드롭다운 목록에 옵션이 보이지 않아 닫은 후 다시 시도합니다. ")
            page.keyboard.press("Escape")
            human.pause()
            continue
        print(f"   선택: “{option.inner_text().strip()}”...")
        human.click(page, option, wait_after=False)
        for _ in range(10):  # “Request Report” 버튼이 나올 때까지 최대 5초 기다림
            if _request_button(page) is not None:
                human.pause()
                return
            page.wait_for_timeout(500)
    path = dump_dom(page, "report_type_dropdown")
    raise RuntimeError(f"드롭다운에서 “{labels[0]}”을(를) 선택하지 못했습니다(페이지 구조 저장: {path.name})")


def fetch_all_listings(page, today: str) -> Path:
    print("   상품 → 상품 보고서 여는 중...")
    sc_nav.menu_go(page, [sc_nav.PRODUCTS, sc_nav.PRODUCT_REPORTS], DEBUG_DIR)
    human.read_page(page, max_screens=2)
    try:
        names = ALL_LISTINGS + ["All Listings"]
        ready = sc_reports.download_ready_today(page, RAW_DIR, f"all_listings_{today}", names, page_tz="local")
        if ready:
            print("   오늘 생성된 All Listings Report가 이미 있어 바로 다운로드합니다. ")
            return ready
        old = sc_reports.existing_rows(page)
        if sc_reports.pending_today(page, names, page_tz="local"):
            print("   오늘 이미 All Listings Report를 요청해 생성 대기 중이므로, 다시 요청하지 않습니다. ")
        else:
            print("   All Listings Report 선택 후 보고서 요청...")
            _select_report_type(page, ALL_LISTINGS)
            old = sc_reports.existing_rows(page)
            btn = None
            for _ in range(10):  # 선택한 뒤 버튼이 조금 늦게 나올 수 있음
                btn = _request_button(page)
                if btn is not None:
                    break
                page.wait_for_timeout(500)
            if btn is None:
                path = dump_dom(page, "request_report_button")
                raise RuntimeError(f"보고서 유형을 선택한 뒤 ‘Request Report’ 버튼이 나타나지 않았습니다(페이지 구조 저장: {path.name})")
            print("   ‘Request Report’ 클릭...")
            human.click(page, btn)
        print("   보고서 생성 대기 중(아마존 안내로는 최대 45분 걸릴 수 있음)...")
        return sc_reports.wait_and_download(page, old, RAW_DIR, f"all_listings_{today}",
                                            row_must_contain=names, timeout_min=45,
                                            debug_dir=DEBUG_DIR)
    except Exception:
        dump_dom(page, "all_listings")
        raise


FBA_TAB = sc_labels.L("FBA inventory reports")
SHOW_MORE = re.compile(r"^\s*(Show more|더 보기|더보기|显示更多|查看更多)\s*\.*…?\s*$", re.I)
REQUEST_RE = re.compile(r"request.*download|generate|다운로드 요청|보고서 요청|请求.*下载|生成报告", re.I)


def _click_and_follow(page, el, wait_s: int = 15):
    """링크를 누른 뒤 목적 페이지로 따라감: 새 탭으로 열릴 수도, 지금 탭에서 이동할 수도 있음. 목적 페이지를 돌려줌."""
    ctx = page.context
    start = page.url
    href = el.get_attribute("href")
    try:
        with ctx.expect_page(timeout=6000) as pi:
            human.click(page, el, wait_after=False)
        target = pi.value
        target.wait_for_load_state("domcontentloaded")
        sc_nav._settle(target)
        return target
    except Exception:
        pass
    try:
        page.wait_for_url(lambda u: u != start, timeout=wait_s * 1000)
    except Exception:
        if href and not href.startswith("javascript"):
            print("   클릭 후 페이지가 이동하지 않아, 해당 보고서 페이지를 바로 엽니다. ")
            human.goto(page, page.evaluate("h => new URL(h, location.href).href", href), wait_after=False)
    sc_nav._settle(page)
    return page


def _find_request_button(page, wait_s: int = 20):
    """보고서 페이지의 ‘다운로드 요청’ 버튼(페이지가 아직 로딩 중일 수 있어 최대 wait_s초 기다림)."""
    import time as _t
    end = _t.time() + wait_s
    while _t.time() < end:
        el = sc_nav._first_visible(page.locator('kat-button[label*="Request .txt" i]'))
        if el is None:
            el = sc_nav.find_visible(page, FBA_REQUEST)
        if el is None:
            loc = page.get_by_role("button", name=REQUEST_RE)
            el = sc_nav._first_visible(loc)
        if el is not None:
            return el
        page.wait_for_timeout(1500)
    return None


def fetch_fba_inventory(page, today: str) -> Tuple[Path, str]:
    print("   공급망 → 공급망 보고서 여는 중...")
    sc_nav.menu_go(page, [sc_nav.SUPPLY_CHAIN, sc_nav.SUPPLY_CHAIN_REPORTS], DEBUG_DIR)
    try:
        # “FBA inventory reports” 탭(보고서 센터)으로 전환
        if "reportcentral" not in page.url:
            tab = sc_nav.find_visible(page, FBA_TAB, prefer_link=True)
            if tab is not None:
                print("   FBA inventory reports로 전환...")
                page = _click_and_follow(page, tab)
        human.read_page(page, max_screens=2)

        # 왼쪽 Inventory 분류: FBA Inventory 먼저, 없으면 Manage FBA Inventory
        # 목록이 다 안 보이면 먼저 Inventory 아래 “Show more...” 클릭
        chosen, el = None, None
        expanded = False
        for names in FBA_REPORT_NAMES:
            el = sc_nav.find_visible(page, names, prefer_link=True)
            if el is None and not expanded:
                more = sc_nav._first_visible(page.get_by_text(SHOW_MORE))
                if more is not None:
                    print("   Inventory의 Show more 펼치는 중...")
                    human.click(page, more)
                    expanded = True
                    el = sc_nav.find_visible(page, names, prefer_link=True)
            if el is not None:
                chosen = names[0]
                break
        if el is None:
            raise RuntimeError("보고서 센터 왼쪽 목록에서 FBA Inventory / Manage FBA Inventory를 찾지 못했습니다")

        print(f"   보고서 열기: “{chosen}”...")
        page = _click_and_follow(page, el)
        human.read_page(page, max_screens=2)

        ready = sc_reports.download_ready_today(page, RAW_DIR, f"fba_inventory_{today}", FBA_FORMAT, page_tz="central")
        if ready:
            print("   오늘 생성된 FBA 재고 보고서(.txt)가 이미 있어 바로 다운로드합니다. ")
            return ready, chosen
        old = sc_reports.existing_rows(page)
        if sc_reports.pending_today(page, FBA_FORMAT, page_tz="central"):
            print("   오늘 이미 FBA 재고 보고서를 요청해 생성 중이므로, 다시 요청하지 않습니다. ")
        else:
            btn = _find_request_button(page)
            if btn is None:
                raise RuntimeError(f"“{chosen}” 보고서 페이지에서 다운로드 요청 버튼을 찾지 못했습니다")
            name = (btn.get_attribute("label") or btn.inner_text() or "").strip()
            print(f"   클릭: “{name[:30]}”...")
            human.click(page, btn)
        print("   보고서 생성 대기 중(최대 45분)...")
        path = sc_reports.wait_and_download(page, old, RAW_DIR, f"fba_inventory_{today}",
                                            row_must_contain=FBA_FORMAT, timeout_min=45, debug_dir=DEBUG_DIR)
        return path, chosen
    except Exception:
        dump_dom(page, "fba_inventory")
        raise


# ================================================================ 파일 읽고 쓰기

TEXT_COLS = {"sku", "sellersku", "asin", "asin1", "asin2", "asin3", "fnsku", "productid", "listingid",
             "productidtype", "upc", "ean", "merchantshippinggroup"}
NUM_HINTS = ("price", "quantity", "qty", "available", "units", "days", "age", "inbound", "reserved",
             "unfulfillable", "unsellable", "cost", "volume", "supply", "excess", "shipped", "fee")


def norm(h: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (h or "").lower())


def read_report(path: Path) -> Tuple[List[str], List[Dict]]:
    """Amazon에서 받은 보고서 읽기(탭 또는 쉼표 구분, 인코딩 자동 인식)."""
    data = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    first = text.split("\n", 1)[0]
    delim = "\t" if "\t" in first else ","
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    rows = [r for r in reader if any(c.strip() for c in r)]
    if not rows:
        return [], []
    headers = [h.strip() for h in rows[0]]
    out = []
    for r in rows[1:]:
        r = r + [""] * (len(headers) - len(r))
        out.append({h: _convert(h, v.strip()) for h, v in zip(headers, r)})
    return headers, out


def _convert(header: str, v: str):
    n = norm(header)
    if n in TEXT_COLS or not any(k in n for k in NUM_HINTS) or v == "":
        return v
    try:
        f = float(v.replace(",", ""))
        return int(f) if f.is_integer() else f
    except ValueError:
        return v


def read_xlsx(path: Path) -> Tuple[List[str], List[Dict]]:
    from openpyxl import load_workbook
    ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    headers = [str(h) if h is not None else "" for h in next(it, [])]
    rows = [{h: ("" if v is None else v) for h, v in zip(headers, r)} for r in it if any(v not in (None, "") for v in r)]
    return headers, rows


# ---- Excel 스타일
FONT = "Malgun Gothic"   # 한글 글꼴(맑은 고딕)


def _styles():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    thin = Side(style="thin", color="BFBFBF")
    return {
        "head_fill": PatternFill("solid", fgColor="DCE6F1"),
        "box": Border(left=thin, right=thin, top=thin, bottom=thin),
        "Font": Font, "Alignment": Alignment, "PatternFill": PatternFill,
    }


def write_table_sheet(ws, headers: List[str], rows: List[Dict], widths: Optional[Dict[str, int]] = None) -> None:
    from openpyxl.utils import get_column_letter
    st = _styles()
    Font, Alignment = st["Font"], st["Alignment"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font, cell.fill, cell.border = Font(name=FONT, bold=True), st["head_fill"], st["box"]
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        w = (widths or {}).get(h)
        if w is None:
            sample = [len(str(r.get(h, ""))) for r in rows[:200]]
            w = min(50, max(10, len(h) + 2, (max(sample) if sample else 0) + 2))
        ws.column_dimensions[get_column_letter(c)].width = w
    for r_i, row in enumerate(rows, start=2):
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=r_i, column=c, value=row.get(h, ""))
            cell.font = Font(name=FONT, size=10)
    ws.freeze_panes = "B2" if len(headers) > 3 else "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"


def save_report_xlsx(headers: List[str], rows: List[Dict], out: Path, title: str) -> None:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = title
    write_table_sheet(ws, headers, rows)
    wb.save(out)


# ================================================================ 평가

def _get(row: Dict, *names, default=None):
    """여러 가능한 열 이름으로 값 찾기(열 이름은 대소문자·기호 무시)."""
    idx = row.get("__norm__")
    if idx is None:
        idx = {norm(k): k for k in row if k != "__norm__"}
        row["__norm__"] = idx
    for n in names:
        k = idx.get(norm(n))
        if k is not None and row[k] not in (None, ""):
            return row[k]
    return default


def _num(v) -> float:
    if v in (None, ""):
        return 0.0
    try:
        return float(str(v).replace(",", "").replace("$", ""))
    except ValueError:
        return 0.0


def build_rows(al_rows: List[Dict], fba_rows: List[Dict]) -> List[Dict]:
    """두 보고서를 SKU로 합쳐 재고 전체 표로."""
    merged: Dict[str, Dict] = {}

    for r in al_rows:
        sku = str(_get(r, "seller-sku", "sku", default="")).strip()
        if not sku:
            continue
        channel = str(_get(r, "fulfillment-channel", default=""))
        merged[sku] = {
            "sku": sku,
            "asin": _get(r, "asin1", "asin", default=""),
            "title": _get(r, "item-name", "product-name", "title", default=""),
            "status": _get(r, "status", default=""),
            "channel": "FBA" if channel.upper().startswith("AMAZON") else ("판매자 배송" if channel else ""),
            "price": _num(_get(r, "price")),
            "fbm_qty": _num(_get(r, "quantity")) if not channel.upper().startswith("AMAZON") else None,
            "open_date": _get(r, "open-date", default=""),
            "in_all": True, "in_fba": False,
        }

    for r in fba_rows:
        sku = str(_get(r, "sku", "seller-sku", default="")).strip()
        if not sku:
            continue
        m = merged.setdefault(sku, {"sku": sku, "in_all": False, "status": "", "channel": "FBA",
                                    "price": 0.0, "fbm_qty": None, "open_date": ""})
        m["in_fba"] = True
        m["asin"] = m.get("asin") or _get(r, "asin", default="")
        m["title"] = m.get("title") or _get(r, "product-name", default="")
        if not m.get("price"):
            m["price"] = _num(_get(r, "your-price", "price"))
        m["available"] = _num(_get(r, "available", "afn-fulfillable-quantity"))
        inbound = _get(r, "inbound-quantity")
        if inbound is None:
            inbound = sum(_num(_get(r, n)) for n in ("afn-inbound-working-quantity",
                                                     "afn-inbound-shipped-quantity",
                                                     "afn-inbound-receiving-quantity"))
        m["inbound"] = _num(inbound)
        m["reserved"] = _num(_get(r, "Total Reserved Quantity", "reserved-quantity", "afn-reserved-quantity"))
        m["unfulfillable"] = _num(_get(r, "unfulfillable-quantity", "afn-unsellable-quantity"))
        t30 = _get(r, "units-shipped-t30", "units-shipped-last-30-days")
        m["t30"] = None if t30 is None else _num(t30)
        t90 = _get(r, "units-shipped-t90", "units-shipped-last-90-days")
        m["t90"] = None if t90 is None else _num(t90)
        dos = _get(r, "days-of-supply")
        if dos not in (None, ""):
            m["dos"] = _num(dos)
        elif m["t30"]:
            m["dos"] = round(m["available"] / (m["t30"] / 30), 1)
        else:
            m["dos"] = None
        m["aged_181"] = sum(_num(_get(r, n)) for n in (
            "inv-age-181-to-270-days", "inv-age-271-to-365-days", "inv-age-365-plus-days",
            "inv-age-181-to-330-days", "inv-age-331-to-365-days"))
        m["aged_365"] = _num(_get(r, "inv-age-365-plus-days"))
        m["excess"] = _num(_get(r, "estimated-excess-quantity"))
        m["storage_cost"] = _num(_get(r, "estimated-storage-cost-next-month"))
        m["recommended"] = _get(r, "recommended-action", default="")

    return list(merged.values())


def evaluate(items: List[Dict], scope: str) -> List[Dict]:
    """
    주의할 항목 만들기: [{level, type, sku, asin, title, detail, advice}]
    scope = "listings": 전체 상품(상품 상태, 판매자 배송 재고, 가격)
    scope = "fba": FBA 재고(품절, 보충, 안 팔림, 재고 기간, 판매 불가), 그날 상품 상태와 함께 봄
    """
    out = []

    def add(level, typ, it, detail, advice):
        out.append({"level": level, "type": typ, "sku": it["sku"], "asin": it.get("asin", ""),
                    "title": it.get("title", ""), "detail": detail, "advice": advice})

    for it in items:
        status = str(it.get("status") or "").lower()

        if scope == "listings":
            if not it.get("in_all"):
                continue
            if status.startswith("inactive"):
                add("중간", "상품 판매 불가", it, "상품 상태가 Inactive로, 구매자가 볼 수 없거나 구매할 수 없습니다. ",
                    "상품 관리에서 판매 불가 사유(재고 없음, 가격 이상, 검색 제한 등)를 확인하세요. ")
            elif status.startswith("incomplete"):
                add("중간", "상품 정보 미완성", it, "상품 상태가 Incomplete입니다. ", "필수 정보를 채워야 정상적으로 판매할 수 있습니다. ")
            if it.get("channel") == "판매자 배송" and status.startswith("active") and (it.get("fbm_qty") or 0) <= 0:
                add("중간", "판매자 배송 재고 없음", it, "판매자 배송 상품 수량이 0입니다. ", "계속 판매 중이라면 재고 수량을 업데이트하세요. ")
            if status.startswith("active") and not it.get("price"):
                add("낮음", "가격 누락", it, "판매 중인 상품에 가격이 없습니다. ", "가격 설정을 확인하세요. ")
            continue

        # ---------------- FBA
        if not it.get("in_fba"):
            continue
        avail, inbound, t30 = it.get("available", 0), it.get("inbound", 0), it.get("t30")
        dos = it.get("dos")

        if avail > 0 and (status.startswith("inactive") or status.startswith("incomplete")):
            add("높음", "재고는 있으나 판매 불가", it, f"FBA 판매 가능 {avail:g}개, 하지만 상품 상태가 {it.get('status')}. ",
                "창고에 재고가 있는데 팔리지 않습니다. 상품 관리에서 원인을 확인하고 빨리 판매를 재개하세요. ")

        if avail <= 0 and (t30 or 0) > 0:
            add("높음", "FBA 품절", it, f"판매 가능 0개, 최근 30일 판매 {t30:g}개" + (f", 입고 중 {inbound:g}개" if inbound else ", 입고 중인 재고 없음"),
                "빨리 재입고하세요. " if not inbound else "입고 중인 배송의 진행 상황을 확인하세요. ")
        elif avail <= 0 and inbound > 0:
            add("중간", "판매 가능 0(입고 중 있음)", it, f"판매 가능 0개, 입고 중 {inbound:g}개. ", "입고 중인 배송의 진행 상황을 확인하세요. ")
        elif dos is not None and avail > 0 and dos < DOS_URGENT:
            add("높음", "재고 부족", it, f"판매 가능 {avail:g}개, 최근 30일 판매량 기준 약 {dos:g}일분" + (f", 입고 중 {inbound:g}개" if inbound else ""),
                f"판매 가능 일수가 {DOS_URGENT}일 미만입니다. 바로 재입고를 진행하세요. ")
        elif dos is not None and avail > 0 and dos < DOS_WATCH:
            add("중간", "재입고 필요", it, f"판매 가능 {avail:g}개, 약 {dos:g}일분" + (f", 입고 중 {inbound:g}개" if inbound else ""),
                f"판매 가능 일수가 {DOS_WATCH}일 미만입니다. 재입고 계획을 세우세요. ")

        if avail > 0 and it.get("t90") is not None and it["t90"] == 0:
            add("중간", "장기 미판매", it, f"판매 가능 {avail:g}개, 최근 90일 판매 없음. ", "할인 프로모션, 광고 또는 재고 제거를 고려하세요. ")

        if it.get("aged_365", 0) > 0:
            add("높음", "초장기 재고", it, f"{it['aged_365']:g}개가 재고 기간 365일 초과. ", "장기 보관료가 높으니 재고 정리 또는 제거를 권장합니다. ")
        elif it.get("aged_181", 0) > 0:
            add("중간", "재고 기간 긴 편", it, f"{it['aged_181']:g}개가 재고 기간 초과: {AGED_DAYS_LABEL}. ", "보관료에 주의하고, 프로모션으로 정리를 고려하세요. ")

        if it.get("unfulfillable", 0) > 0:
            add("중간", "판매 불가 재고", it, f"{it['unfulfillable']:g}개 판매 불가(파손 등). ", "제거 또는 폐기 주문을 생성하세요. ")

        if it.get("excess", 0) > 0:
            add("낮음", "과잉 재고", it, f"아마존 예상 과잉 재고 {it['excess']:g}개. ", "재입고를 줄이거나 프로모션을 진행하세요. ")

        rec = str(it.get("recommended") or "").strip()
        if rec and rec.lower() not in ("no action required", "none", "normal", "조치 불필요"):
            add("낮음", "아마존 제안", it, f"아마존 제안: {rec}", "아마존 제안을 참고해 처리하세요. ")

    order = {"높음": 0, "중간": 1, "낮음": 2}
    out.sort(key=lambda x: (order[x["level"]], x["type"], x["sku"]))
    return out


def summarize_listings(items: List[Dict], n: int) -> Dict:
    all_items = [i for i in items if i.get("in_all")]
    status = lambda s: sum(1 for i in all_items if str(i.get("status", "")).lower().startswith(s))
    return {
        "전체 상품 수(All Listings)": n,
        "판매 중 Active": status("active"),
        "판매 불가 Inactive": status("inactive"),
        "정보 미완성 Incomplete": status("incomplete"),
        "FBA 상품": sum(1 for i in all_items if i.get("channel") == "FBA"),
        "판매자 배송 상품": sum(1 for i in all_items if i.get("channel") == "판매자 배송"),
        "판매자 배송 재고 수량": sum(i.get("fbm_qty") or 0 for i in all_items if i.get("channel") == "판매자 배송"),
    }


def summarize_fba(items: List[Dict], n: int) -> Dict:
    fba = [i for i in items if i.get("in_fba")]
    return {
        "FBA 보고서 SKU 수": n,
        "판매 가능 재고가 있는 SKU": sum(1 for i in fba if i.get("available", 0) > 0),
        "FBA 판매 가능 수량": sum(i.get("available", 0) for i in fba),
        "입고 중 수량": sum(i.get("inbound", 0) for i in fba),
        "예약 수량": sum(i.get("reserved", 0) for i in fba),
        "판매 불가 수량": sum(i.get("unfulfillable", 0) for i in fba),
        "최근 30일 판매량(개)": sum(i.get("t30") or 0 for i in fba),
        "판매 가능 재고 가치(판매가 기준)": round(sum(i.get("available", 0) * (i.get("price") or 0) for i in fba), 2),
        "다음 달 예상 보관료": round(sum(i.get("storage_cost", 0) for i in fba), 2),
    }


MONEY_KEYS = {"판매 가능 재고 가치(판매가 기준)", "다음 달 예상 보관료"}


# ================================================================ 보고서

def write_report(out: Path, title: str, today: str, sources: Dict[str, str], summary: Dict,
                 findings: List[Dict], overview: Tuple[List[str], List[Dict]], raw: Tuple[str, Tuple],
                 errors: List[str], rules: str) -> None:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    st = _styles()
    Font, Alignment = st["Font"], st["Alignment"]
    BOX = st["box"]
    RED, ORANGE, GREY, GREEN, NAVY = "C01C1C", "C26A00", "6B6B6B", "1E7B34", "1F3864"
    LEVEL_COLOR = {"높음": RED, "중간": ORANGE, "낮음": GREY}

    wb = Workbook()
    ws = wb.active
    ws.title = "평가 요약"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABC", (30, 18, 60)):
        ws.column_dimensions[col].width = w

    def put(r, c, v, size=10.5, bold=False, color=None, num=None, wrap=False):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name=FONT, size=size, bold=bold, color=color)
        cell.alignment = Alignment(vertical="center", wrap_text=wrap)
        if num:
            cell.number_format = num
        return cell

    put(1, 1, title, 18, True, NAVY)
    put(2, 1, f"{today} · Amazon.com(미국)", 9.5, color=GREY)
    r = 4
    put(r, 1, "데이터 출처", 13, True, NAVY)
    for k, v in sources.items():
        r += 1
        put(r, 1, k, bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        put(r, 2, v, color=GREY)
    for e in errors:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        put(r, 1, f"⚠ {e}", color=RED, wrap=True)
        ws.row_dimensions[r].height = 30

    r += 2
    put(r, 1, "개요", 13, True, NAVY)
    for k, v in summary.items():
        r += 1
        put(r, 1, k)
        put(r, 2, v, bold=True, num='"$"#,##0.00' if k in MONEY_KEYS else "#,##0")

    r += 2
    put(r, 1, "평가 결론", 13, True, NAVY)
    if not findings:
        r += 1
        put(r, 1, "확인이 필요한 문제가 없습니다 ✓", bold=True, color=GREEN)
    else:
        hi = sum(1 for f_ in findings if f_["level"] == "높음")
        r += 1
        put(r, 1, f"총 {len(findings)}개 항목 확인 필요, 그중 높은 우선순위 {hi}개 항목", bold=True, color=RED if hi else ORANGE)
        r += 1
        for c, h in enumerate(["등급 · 유형", "수량", "제안"], start=1):
            cell = put(r, c, h, bold=True)
            cell.fill, cell.border = st["head_fill"], BOX
        counts, advice = {}, {}
        for f_ in findings:
            counts[(f_["level"], f_["type"])] = counts.get((f_["level"], f_["type"]), 0) + 1
            advice[f_["type"]] = f_["advice"]
        for (lvl, typ), n in sorted(counts.items(), key=lambda x: ({"높음": 0, "중간": 1, "낮음": 2}[x[0][0]], -x[1])):
            r += 1
            put(r, 1, f"{lvl} · {typ}", bold=True, color=LEVEL_COLOR[lvl]).border = BOX
            put(r, 2, n, num="#,##0").border = BOX
            put(r, 3, advice[typ], wrap=True).border = BOX
        r += 1
        put(r, 1, "상품별 상세 내용은 ‘확인 필요’ 시트를 참고하세요. ", 9, color=GREY)

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    put(r, 1, rules, 9, color=GREY, wrap=True)
    ws.row_dimensions[r].height = 45

    # ---- 주의할 항목
    fs = wb.create_sheet("확인 필요")
    cols = ["등급", "유형", "SKU", "ASIN", "상품명", "상황", "제안"]
    for c, (h, w) in enumerate(zip(cols, [7, 16, 18, 13, 44, 46, 40]), start=1):
        cell = fs.cell(row=1, column=c, value=h)
        cell.font, cell.fill, cell.border = Font(name=FONT, bold=True), st["head_fill"], BOX
        fs.column_dimensions[get_column_letter(c)].width = w
    for r_i, f_ in enumerate(findings, start=2):
        vals = [f_["level"], f_["type"], f_["sku"], f_["asin"], f_["title"], f_["detail"], f_["advice"]]
        for c, v in enumerate(vals, start=1):
            cell = fs.cell(row=r_i, column=c, value=v)
            cell.font = Font(name=FONT, size=10, bold=(c == 1), color=LEVEL_COLOR[f_["level"]] if c == 1 else None)
            cell.border = BOX
            cell.alignment = Alignment(vertical="top", wrap_text=c in (5, 6, 7))
    fs.freeze_panes = "C2"
    if findings:
        fs.auto_filter.ref = f"A1:G{len(findings) + 1}"

    # ---- 목록표 + 원본 데이터
    oh, orows = overview
    ts = wb.create_sheet("전체 목록")
    write_table_sheet(ts, oh, orows, {"SKU": 18, "ASIN": 13, "상품명": 44, "아마존 제안": 22})
    if "확인 항목 수" in oh:
        ci = oh.index("확인 항목 수") + 1
        for r_i, row in enumerate(orows, start=2):
            if row["확인 항목 수"]:
                ts.cell(row=r_i, column=ci).font = Font(name=FONT, size=10, bold=True, color=RED)
    if "가격" in oh:
        ci = oh.index("가격") + 1
        for r_i in range(2, len(orows) + 2):
            ts.cell(row=r_i, column=ci).number_format = '"$"#,##0.00'
    raw_title, (rh, rrows) = raw
    if rh:
        write_table_sheet(wb.create_sheet(raw_title), rh, rrows)

    for sheet in wb.worksheets:
        sheet.page_setup.orientation = "portrait" if sheet.title == "평가 요약" else "landscape"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0

    wb.save(out)


def _count_by_sku(findings):
    n = {}
    for f_ in findings:
        n[f_["sku"]] = n.get(f_["sku"], 0) + 1
    return n


def listings_overview(items, findings):
    n = _count_by_sku(findings)
    h = ["SKU", "ASIN", "상품명", "상태", "배송 방식", "가격", "판매자 배송 수량", "등록일", "확인 항목 수"]
    rows = [{"SKU": i["sku"], "ASIN": i.get("asin", ""), "상품명": i.get("title", ""), "상태": i.get("status", ""),
             "배송 방식": i.get("channel", ""), "가격": i.get("price") or "",
             "판매자 배송 수량": "" if i.get("fbm_qty") is None else i["fbm_qty"], "등록일": i.get("open_date", ""),
             "확인 항목 수": n.get(i["sku"], 0)}
            for i in sorted((i for i in items if i.get("in_all")), key=lambda x: (-n.get(x["sku"], 0), x["sku"]))]
    return h, rows


def fba_overview(items, findings):
    n = _count_by_sku(findings)
    h = ["SKU", "ASIN", "상품명", "상품 상태", "가격", "FBA 판매 가능", "입고 중", "예약", "판매 불가", "최근 30일 판매량",
         "판매 가능 일수", "재고 기간>180일", "아마존 제안", "확인 항목 수"]
    rows = [{"SKU": i["sku"], "ASIN": i.get("asin", ""), "상품명": i.get("title", ""),
             "상품 상태": i.get("status", "") or "(오늘 상품 보고서 없음)", "가격": i.get("price") or "",
             "FBA 판매 가능": i.get("available", ""), "입고 중": i.get("inbound", ""), "예약": i.get("reserved", ""),
             "판매 불가": i.get("unfulfillable", ""), "최근 30일 판매량": "" if i.get("t30") is None else i["t30"],
             "판매 가능 일수": "" if i.get("dos") is None else i["dos"], "재고 기간>180일": i.get("aged_181", ""),
             "아마존 제안": i.get("recommended", ""), "확인 항목 수": n.get(i["sku"], 0)}
            for i in sorted((i for i in items if i.get("in_fba")), key=lambda x: (-n.get(x["sku"], 0), x["sku"]))]
    return h, rows


# ================================================================ 메인 흐름

def _err(what: str, e: Exception) -> str:
    return f"{what} 가져오기 실패: {type(e).__name__}: {str(e).splitlines()[0][:200]}(스크린샷과 페이지 텍스트는 debug 폴더에 있음)"


def _finish(out: Path, summary: Dict, findings: List[Dict], show_keys, open_report: bool,
            module: str = "", extra: Optional[List[Dict]] = None) -> Path:
    try:
        import daily_bus
        daily_bus.record(module or out.stem, {k: summary.get(k) for k in show_keys if k in summary},
                         [({"높음": "P1", "중간": "P2"}.get(f_["level"], "P3"), f_["type"],
                           f"{f_.get('sku', '')} {f_.get('detail', '')}".strip(), f_.get("advice", "")) for f_ in findings],
                         out, [a for r in (extra or []) if r for a in r.get("alerts", [])])
        km = {f"{r.get('name', '')}: {k}": v for r in (extra or []) if r for k, v in r.get("kpis", {}).items()
              if not isinstance(v, (list, dict))}
        if km:
            d = daily_bus.load()
            d["modules"][module or out.stem]["kpis"].update(km)
            daily_bus._save(d)
    except Exception:
        pass
    if extra:
        try:
            import sc_pages
            sc_pages.print_alerts(extra)
        except Exception:
            pass
    print()
    for k in show_keys:
        v = summary[k]
        print(f"{k}: {v:,.0f}" if isinstance(v, float) and v.is_integer() else f"{k}: {v:,}")
    hi = [f_ for f_ in findings if f_["level"] == "높음"]
    print(f"확인 필요 {len(findings)}개 항목, 그중 높은 우선순위 {len(hi)}개. ")
    for f_ in hi[:8]:
        print(f"  • [{f_['type']}] {f_['sku']}: {f_['detail']}")
    print(f"\n📊 보고서 저장 완료: {out}")
    if open_report:
        try:
            __import__("platform_util").open_path(out)
        except Exception:
            pass
    return out


def run_listings(open_report: bool = True) -> Path:
    """전체 상품 재고: All Listings Report(하루 한 번) → Excel → 평가."""
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    data = REPORT_DIR / f"all_listings_{today}.xlsx"
    out = REPORT_DIR / f"listings_report_{today}.xlsx"
    sources, errors = {}, []

    if data.exists():
        print(f"   오늘의 전체 상품 파일이 이미 있습니다({data.name}). 다시 가져오지 않습니다. ")
        sources["전체 상품"] = f"{data.name}(오늘 이미 가져옴, 그대로 사용)"
    else:
        import sp_api
        import sp_data
        raw = sp_api.try_api("전체 상품(All Listings)", sp_data.all_listings, today)
        if raw:
            h, rows = read_report(raw)
            save_report_xlsx(h, rows, data, "All Listings")
            sources["전체 상품"] = f"{data.name}(SP-API All Listings Report, {len(rows)}행)"
            print(f"   ✅ 전체 상품(API): {len(rows)}행, 저장 완료 {data.name}")
        else:
            with open_session(headless=False) as page:
                try:
                    raw = fetch_all_listings(page, today)
                    h, rows = read_report(raw)
                    save_report_xlsx(h, rows, data, "All Listings")
                    sources["전체 상품"] = f"{data.name}(All Listings Report, {len(rows)}행)"
                    print(f"   ✅ 전체 상품: {len(rows)}행, 저장 완료 {data.name}")
                except Exception as e:
                    raise RuntimeError(_err("전체 상품 보고서", e)) from e

    al = read_xlsx(data)
    items = build_rows([dict(r) for r in al[1]], [])
    findings = evaluate(items, "listings")
    summary = summarize_listings(items, len(al[1]))
    rules = "평가 기준: Inactive는 판매 불가, Incomplete는 정보 미완성. 판매 중인 판매자 배송 상품의 수량이 0이면 재고 보충 알림, 판매 중인 상품에 가격이 없으면 확인 알림. "
    write_report(out, "전체 상품 재고 일일 보고서", today, sources, summary, findings,
                 listings_overview(items, findings), ("전체 상품", al), errors, rules)
    return _finish(out, summary, findings, ("전체 상품 수(All Listings)", "판매 중 Active", "판매 불가 Inactive",
                                           "정보 미완성 Incomplete"), open_report, module="전체 상품 재고")


def run_fba(open_report: bool = True) -> Path:
    """FBA 재고: 공급망 보고서의 FBA 재고 보고서(하루 한 번) → Excel → 평가."""
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    data = REPORT_DIR / f"fba_inventory_{today}.xlsx"
    listings = REPORT_DIR / f"all_listings_{today}.xlsx"
    out = REPORT_DIR / f"fba_report_{today}.xlsx"
    sources, errors = {}, []

    if data.exists():
        print(f"   오늘의 FBA 재고 파일이 이미 있습니다({data.name}). 다시 가져오지 않습니다. ")
        sources["FBA 재고"] = f"{data.name}(오늘 이미 가져옴, 그대로 사용)"
    else:
        import sp_api
        import sp_data
        raw = sp_api.try_api("FBA 재고", sp_data.fba_inventory, today)
        if raw:
            h, rows = read_report(raw)
            save_report_xlsx(h, rows, data, "FBA Inventory")
            sources["FBA 재고"] = f"{data.name}(SP-API FBA 재고 보고서, {len(rows)}행)"
            print(f"   ✅ FBA 재고(API): {len(rows)}행, 저장 완료 {data.name}")
        else:
            with open_session(headless=False) as page:
                try:
                    raw, name = fetch_fba_inventory(page, today)
                    h, rows = read_report(raw)
                    save_report_xlsx(h, rows, data, "FBA Inventory")
                    sources["FBA 재고"] = f"{data.name}({name} 보고서, {len(rows)}행)"
                    print(f"   ✅ FBA 재고: {len(rows)}행, 저장 완료 {data.name}")
                except Exception as e:
                    raise RuntimeError(_err("FBA 재고 보고서", e)) from e

    fba = read_xlsx(data)
    al = read_xlsx(listings) if listings.exists() else ([], [])
    if al[1]:
        sources["상품 상태"] = f"{listings.name}(재고는 있으나 판매 불가인 상품 확인용)"
    else:
        sources["상품 상태"] = "오늘의 전체 상품 파일이 아직 없습니다. 먼저 ‘전체 상품 재고’를 실행하면 ‘재고는 있으나 판매 불가’도 함께 확인합니다"
    items = build_rows([dict(r) for r in al[1]], [dict(r) for r in fba[1]])
    findings = evaluate(items, "fba")
    summary = summarize_fba(items, len(fba[1]))
    rules = (f"평가 기준: 판매 가능 일수 = 판매 가능 수량 ÷ 최근 30일 일평균 판매량(보고서에 판매 가능 일수가 있으면 우선 사용). {DOS_URGENT}일 미만은 재고 부족, "
             f" {DOS_WATCH}일 미만은 재입고 필요. 최근 90일 판매 없음은 장기 미판매. 재고 기간 {AGED_DAYS_LABEL} / 365일 초과 시 보관료 위험 알림. "
             f"판매 가능 재고가 있는데 상품 상태가 Inactive / Incomplete이면 높은 우선순위. 임계값은 inventory.py 맨 위에서 수정할 수 있습니다. ")
    write_report(out, "FBA 재고 일일 보고서", today, sources, summary, findings,
                 fba_overview(items, findings), ("FBA 재고", fba), errors, rules)
    extra = []
    try:
        import sc_pages
        print("\n   FBA 대시보드·입고 배송 확인...")
        res = sc_pages.read_many(["fba_dashboard", "shipments"])
        extra = [res["fba_dashboard"], res["shipments"]]
    except Exception as e:
        print(f"   ⚠ FBA 대시보드·입고 배송을 읽지 못했습니다: {type(e).__name__}: {str(e)[:100]}")
    try:
        import restock
        from datetime import datetime as _dt
        rp = restock.plan(items)
        rp["time"] = _dt.now().isoformat(timespec="seconds")
        extra.insert(0, rp)
        print(f"   재입고 제안: {rp['kpis']['보충 제안 SKU']}개 SKU(지금 보내야 함 {rp['kpis']['지금 보내야 하는 SKU']}개)")
    except Exception as e:
        print(f"   ⚠ 재입고 제안을 만들지 못했습니다: {type(e).__name__}: {str(e)[:100]}")
    try:
        import sc_pages
        sc_pages.append_sheets(out, extra)
    except Exception as e:
        print(f"   ⚠ 추가 시트(재입고 제안·대시보드)를 넣지 못했습니다: {type(e).__name__}: {str(e)[:100]}")
    return _finish(out, summary, findings, ("FBA 보고서 SKU 수", "FBA 판매 가능 수량", "입고 중 수량", "최근 30일 판매량(개)"),
                   open_report, module="FBA 재고", extra=extra)


if __name__ == "__main__":
    run_listings()
    run_fba()

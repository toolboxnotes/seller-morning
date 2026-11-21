"""
Seller Central 메뉴 이동

사람처럼 조작: 왼쪽 위 메뉴 → 1단계 분류(예: Products) → 2단계 항목(예: Product reports).
메뉴 이름마다 영어 / 한국어 / 중국어 표기를 모두 적어 화면 언어가 달라도 찾을 수 있음.
메뉴가 안 열리면 메뉴의 그 링크 주소로 바로 엶.
"""

import re
from pathlib import Path
from typing import List, Optional

import human
import sc_labels

SC_BASE = "https://sellercentral.amazon.com"

MENU = sc_labels.L("Menu")
PRODUCTS = sc_labels.L("Products")
PRODUCT_REPORTS = sc_labels.L("Product reports")
ORDERS = sc_labels.L("Orders")
ORDER_REPORTS = sc_labels.L("Order reports")
SUPPLY_CHAIN = sc_labels.L("Supply chain")
SUPPLY_CHAIN_REPORTS = sc_labels.L("Supply chain reports")


def _first_visible(locator) -> Optional[object]:
    for i in range(min(locator.count(), 15)):
        el = locator.nth(i)
        try:
            if el.is_visible():
                return el
        except Exception:
            continue
    return None


def find_visible(page, labels: List[str], prefer_link: bool = False):
    """글자가 완전히 같은 첫 번째 보이는 요소 찾기(영·한·중 이름 사전으로 다른 언어 이름도 함께 찾음)."""
    for lab in sc_labels.expand(labels):
        roles = ("link", "button", "menuitem") if prefer_link else ("button", "menuitem", "link")
        for role in roles:
            el = _first_visible(page.get_by_role(role, name=lab, exact=True))
            if el is not None:
                return el
        el = _first_visible(page.get_by_text(lab, exact=True))
        if el is not None:
            return el
    return None


def link_href(page, labels: List[str]) -> Optional[str]:
    """메뉴의 어떤 링크 주소 읽기(메뉴가 안 펼쳐져도 읽힘)."""
    labels = sc_labels.expand(labels)
    href = page.evaluate(
        """(labels) => {
            for (const a of document.querySelectorAll('a[href]')) {
              const t = (a.textContent || '').trim();
              if (labels.includes(t)) return a.href;
            }
            return null;
        }""", labels)
    return href


def open_menu(page) -> bool:
    el = find_visible(page, MENU)
    if el is None:
        el = _first_visible(page.locator(
            '[aria-label="Menu"], [aria-label="메뉴"], [aria-label="菜单"], #navbar-menu-icon, .nav-hamburger'))
    if el is None:
        return False
    human.click(page, el)
    return True


def menu_go(page, steps: List[List[str]], debug_dir: Optional[Path] = None) -> str:
    """
    메뉴를 차례로 클릭, 예: menu_go(page, [PRODUCTS, PRODUCT_REPORTS]).
    돌려줌 "menu"(메뉴로 도착) 또는 "url"(링크 주소로 바로 엶).
    """
    target_href = link_href(page, steps[-1])
    try:
        if not open_menu(page):
            raise RuntimeError("왼쪽 위 메뉴 버튼을 찾지 못했습니다")
        for labels in steps[:-1]:
            el = find_visible(page, labels)
            if el is None:
                raise RuntimeError(f"메뉴에서 찾지 못함: “{labels[0]}”")
            human.click(page, el)
        el = find_visible(page, steps[-1], prefer_link=True)
        if el is None:
            raise RuntimeError(f"메뉴에서 찾지 못함: “{steps[-1][0]}”")
        start = page.url
        human.click(page, el, wait_after=False)
        try:
            page.wait_for_url(lambda u: u != start, timeout=15000)
        except Exception:
            pass
        _settle(page)
        return "menu"
    except Exception as e:
        if debug_dir:
            save_debug(page, debug_dir, f"menu_{steps[-1][0]}")
        if not target_href:
            raise RuntimeError(f"메뉴로 이동하지 못함: “{steps[-1][0]}”: {e}") from e
        print(f"   메뉴 조작에 실패했습니다({e}). 해당 페이지를 바로 엽니다. ")
        page.keyboard.press("Escape")
        human.goto(page, target_href, wait_after=False)
        _settle(page)
        return "url"


def _settle(page) -> None:
    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    human.pause()


def save_debug(page, debug_dir: Path, name: str) -> Path:
    """지금 페이지의 글자와 스크린샷 저장(인식 규칙 조정용)."""
    from datetime import datetime
    debug_dir.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w.-]+", "_", name)[:60]
    base = debug_dir / f"{datetime.now():%Y%m%d_%H%M%S}_{safe}"
    try:
        Path(f"{base}.txt").write_text(f"URL: {page.url}\n\n{page.inner_text('body')}", encoding="utf-8")
        page.screenshot(path=f"{base}.png", full_page=True)
    except Exception:
        pass
    return base

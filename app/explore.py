"""
페이지 구조 탐색(보기만, 바꾸지 않음)

광고 Agent 개발 준비: 사람처럼 아래 페이지를 둘러보고 페이지마다 무엇이 있는지(제목, 탭, 버튼, 드롭다운, 표 열, 링크) 기록하고,
스크린샷·페이지 글자·페이지 구조를 저장. 페이지 열기와 메뉴 링크만 누르고 ‘만들기 / 저장 / 수정 / 삭제’ 버튼은 누르지 않음.

  1. 메뉴: 모든 메뉴 분류와 그 아래 항목(이름 + 주소) 읽기
  2. 메뉴 마케팅 → Brand analytics와 그 아래 보고서 페이지들
  3. 메뉴 마케팅 → 광고 콘솔(Ads console)과 왼쪽 메뉴의 페이지들

결과는 reports/explore/날짜_시간/ 에 저장:
  explore_summary.xlsx   메뉴 구조, 페이지 목록, 페이지별 버튼 / 탭 / 표 열 등
  pages.json             같은 내용(개발용)
  NN_이름.png / .txt / _dom.txt   페이지별 스크린샷, 글자, 구조
"""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import urlparse

import human
import sc_labels
import inventory as inv
import sc_nav
from sc_browser import open_session

REPORT_DIR = Path(__file__).resolve().parent.parent / "reports" / "explore"

MARKETING = sc_labels.L("Marketing")
BRAND_ANALYTICS = sc_labels.L("Brand analytics")
ADS_CONSOLE = sc_labels.L("Ads console")

MAX_BA_PAGES = 15        # Brand Analytics 하위 페이지 최대 개수
MAX_ADS_PAGES = 25       # 광고 콘솔 하위 페이지 최대 개수
# 주소나 글자에 이 단어가 있는 링크는 열지 않음(만들기·수정·삭제 페이지에 들어가지 않게)
UNSAFE = re.compile(r"create|/cb\b|/cb/|/cb\?|campaign-builder|/new\b|new-|edit|delete|remove|upload|archive|logout|sign-?out|pause|enable|launch|"
                    r"submit|bulk|duplicate|copy|payment|billing|invite|settings/users|"
                    r"만들기|생성|삭제|수정|편집|업로드|로그아웃|创建|删除|编辑|上传|退出", re.I)

_MENU_JS = r"""
() => {
  const out = [];
  for (const cat of document.querySelectorAll('.ngstrim-nav-menu-category')) {
    const l1 = cat.querySelector('.ngstrim-nav-menu-l1');
    if (!l1) continue;
    const name = (cat.querySelector('.ngstrim-nav-menu-l1-label') || l1).textContent.trim();
    const items = [];
    if (l1.tagName === 'A') items.push({name: name, href: l1.href});
    for (const a of cat.querySelectorAll('a.ngstrim-nav-menu-l2')) {
      const t = Array.from(a.childNodes).filter(n => n.nodeType === 3).map(n => n.textContent).join(' ').trim()
                || a.textContent.trim();
      items.push({name: t, href: a.href});
    }
    out.push({group: name, items: items});
  }
  return out;
}
"""

# 페이지의 주요 요소 읽기(컴포넌트 안 shadow DOM 포함)
_PAGE_JS = r"""
() => {
  const vis = (e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const txt = (e) => (e.innerText || e.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  const all = [];
  const walk = (root) => {
    for (const e of root.querySelectorAll('*')) {
      all.push(e);
      if (e.shadowRoot) walk(e.shadowRoot);
    }
  };
  walk(document);
  const uniq = (arr) => Array.from(new Set(arr.filter(x => x)));
  const inNav = (e) => !!(e.closest && e.closest('#ngstrim-nav-menu, .ngstrim-nav-menu-sidebar, #sc-footer-container, .ngs-action-center'));
  const res = {title: document.title, headings: [], tabs: [], buttons: [], links: [], tables: [], inputs: [], dropdowns: [], iframes: []};
  for (const e of all) {
    const tag = e.tagName.toLowerCase();
    if (inNav(e)) continue;
    if (/^h[1-3]$/.test(tag) && vis(e)) res.headings.push(txt(e));
    if ((e.getAttribute('role') === 'tab' || tag === 'kat-tab') && vis(e))
      res.tabs.push(e.getAttribute('label') || txt(e));
    if ((tag === 'button' || tag === 'kat-button' || e.getAttribute('role') === 'button') && vis(e)) {
      const t = e.getAttribute('label') || e.getAttribute('aria-label') || txt(e);
      if (t && t.length < 80) res.buttons.push(t);
    }
    if (tag === 'a' && e.href && vis(e)) res.links.push({text: txt(e) || e.getAttribute('aria-label') || '', href: e.href});
    if (tag === 'th' || e.getAttribute('role') === 'columnheader') { const t = txt(e); if (t) res.tables.push(t); }
    if ((tag === 'input' || tag === 'textarea' || tag === 'kat-input') && vis(e)) {
      const t = e.getAttribute('aria-label') || e.getAttribute('placeholder') || e.getAttribute('label') || e.getAttribute('name') || '';
      if (t) res.inputs.push(t);
    }
    if (tag === 'select' || tag === 'kat-dropdown') {
      const opts = Array.from(e.querySelectorAll('option, kat-option')).map(o => txt(o)).filter(x => x).slice(0, 40);
      res.dropdowns.push({label: e.getAttribute('label') || e.getAttribute('aria-label') || e.getAttribute('name') || '',
                          value: e.getAttribute('value') || '', options: opts});
    }
    if (tag === 'iframe' && e.src) res.iframes.push(e.src);
  }
  res.headings = uniq(res.headings); res.tabs = uniq(res.tabs); res.buttons = uniq(res.buttons);
  res.tables = uniq(res.tables); res.inputs = uniq(res.inputs);
  const seen = new Set();
  res.links = res.links.filter(l => { if (seen.has(l.href)) return false; seen.add(l.href); return true; });
  return res;
}
"""


class Explorer:
    def __init__(self, page, out_dir: Path):
        self.page = page
        self.out = out_dir
        self.pages: List[Dict] = []
        self.visited = set()

    def _norm(self, url: str) -> str:
        u = urlparse(url)
        return f"{u.netloc}{u.path.rstrip('/')}#{u.fragment.split('?')[0]}"

    def capture(self, section: str, name: str) -> Dict:
        page = self.page
        sc_nav._settle(page)
        try:
            page.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass
        human.read_page(page, max_screens=4)
        human.move_away(page)
        n = len(self.pages) + 1
        safe = re.sub(r"[^\w.-]+", "_", name)[:50] or "page"
        base = self.out / f"{n:02d}_{safe}"
        info = {"no": n, "section": section, "name": name, "url": page.url}
        try:
            info.update(page.evaluate(_PAGE_JS))
        except Exception as e:
            info["error"] = f"{type(e).__name__}: {e}"
        try:
            Path(f"{base}.txt").write_text(f"URL: {page.url}\n\n{page.inner_text('body')}", encoding="utf-8")
        except Exception:
            pass
        try:
            Path(f"{base}_dom.txt").write_text(f"URL: {page.url}\n\n{page.evaluate(inv._DUMP_JS)}", encoding="utf-8")
        except Exception:
            pass
        try:
            page.screenshot(path=f"{base}.png", full_page=True)
        except Exception:
            try:
                page.screenshot(path=f"{base}.png")
            except Exception:
                pass
        info["files"] = base.name
        self.pages.append(info)
        self.visited.add(self._norm(page.url))
        print(f"   ✅ [{n:02d}] {name}: 제목 {len(info.get('headings', []))}개, 탭 {len(info.get('tabs', []))}개, "
              f"버튼 {len(info.get('buttons', []))}개, 표 열 {len(info.get('tables', []))}개")
        return info

    def menu_click(self, group: List[str], item: List[str]) -> bool:
        """메뉴 → 분류 → 항목. 새 탭으로 열리면 그 탭으로 전환."""
        page = self.page
        try:
            if not sc_nav.open_menu(page):
                raise RuntimeError("메뉴 버튼 없음")
            g = sc_nav.find_visible(page, group)
            if g is None:
                raise RuntimeError(f"메뉴에 “{group[0]}” 없음")
            human.click(page, g)
            link = sc_nav.find_visible(page, item, prefer_link=True)
            if link is None:
                raise RuntimeError(f"메뉴에 “{item[0]}” 없음")
            target = inv._click_and_follow(page, link)
            if target is not page:
                url = target.url
                target.close()
                human.goto(page, url, wait_after=False)
            return True
        except Exception as e:
            print(f"   메뉴 조작 실패({str(e)[:80]}).")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            return False

    def visit(self, url: str, section: str, name: str) -> Optional[Dict]:
        if self._norm(url) in self.visited:
            return None
        print(f"   열기: {name[:50]} ({url[:90]})")
        try:
            human.goto(self.page, url, wait_after=False)
        except Exception as e:
            print(f"   ⚠ 열지 못했습니다: {str(e).splitlines()[0][:120]}")
            return None
        return self.capture(section, name)

    def sub_links(self, info: Dict, must: List[str], limit: int) -> List[Dict]:
        """페이지에서 같은 기능 영역의 링크(나온 순서, 위험한 조작·이미 본 것 제외)."""
        out, seen = [], set()
        for l in info.get("links", []):
            href, text = l.get("href", ""), (l.get("text") or "").strip()
            if not href.startswith("http") or not any(m in href for m in must):
                continue
            if UNSAFE.search(href) or UNSAFE.search(text):
                continue
            key = self._norm(href)
            if key in self.visited or key in seen:
                continue
            seen.add(key)
            out.append({"href": href, "text": text or urlparse(href).path})
            if len(out) >= limit:
                break
        return out


def read_menu(page) -> List[Dict]:
    try:
        return page.evaluate(_MENU_JS)
    except Exception:
        return []


def write_summary(out_dir: Path, menu: List[Dict], pages: List[Dict]) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    FONT = inv.FONT
    HEAD = PatternFill("solid", fgColor="DCE6F1")
    wb = Workbook()

    def sheet(ws, headers, rows, widths):
        for c, h in enumerate(headers, start=1):
            x = ws.cell(row=1, column=c, value=h)
            x.font, x.fill = Font(name=FONT, bold=True), HEAD
            ws.column_dimensions[x.column_letter].width = widths[c - 1]
        for r, row in enumerate(rows, start=2):
            for c, v in enumerate(row, start=1):
                x = ws.cell(row=r, column=c, value=v)
                x.font = Font(name=FONT, size=10)
                x.alignment = Alignment(vertical="top", wrap_text=True)
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "메뉴 구조"
    sheet(ws, ["분류", "항목", "주소"], [(g["group"], i["name"], i["href"]) for g in menu for i in g["items"]], [18, 40, 70])

    ps = wb.create_sheet("페이지 목록")
    j = lambda xs: "\n".join(xs) if xs else ""
    sheet(ps, ["번호", "구역", "이름", "주소", "페이지 제목", "제목(h1~h3)", "탭", "버튼", "표 열", "입력란", "드롭다운", "iframe", "파일"],
          [(p["no"], p["section"], p["name"], p["url"], p.get("title", ""), j(p.get("headings")), j(p.get("tabs")),
            j(p.get("buttons")), j(p.get("tables")), j(p.get("inputs")),
            j([f"{d['label'] or '(이름 없음)'} = {d['value']}: {' / '.join(d['options'][:15])}" for d in p.get("dropdowns", [])]),
            j(p.get("iframes")), p.get("files", "")) for p in pages],
          [6, 12, 28, 50, 30, 30, 25, 35, 35, 25, 45, 30, 22])

    ls = wb.create_sheet("페이지 링크")
    sheet(ls, ["번호", "페이지", "링크 텍스트", "주소"],
          [(p["no"], p["name"], l.get("text", ""), l.get("href", "")) for p in pages for l in p.get("links", [])],
          [6, 28, 40, 80])
    out = out_dir / "explore_summary.xlsx"
    wb.save(out)
    return out


def run(open_report: bool = True) -> Optional[Path]:
    out_dir = REPORT_DIR / f"{datetime.now():%Y%m%d_%H%M}"
    out_dir.mkdir(parents=True, exist_ok=True)
    print("   페이지는 열어 보기만 하고, 만들기 / 저장 / 수정 / 삭제 버튼은 누르지 않습니다.")
    menu, ex = [], None
    try:
        with open_session(headless=False) as page:
            ex = Explorer(page, out_dir)
            # ---- 1. 전체 메뉴
            print("\n【1】 메뉴 구조 읽기...")
            menu = read_menu(page)
            print(f"   ✅ 메뉴 분류 {len(menu)}개, 항목 {sum(len(g['items']) for g in menu)}개")
            ex.capture("홈", "Seller Central 홈")

            # ---- 2. Brand Analytics
            print("\n【2】 마케팅 → Brand analytics...")
            if not ex.menu_click(MARKETING, BRAND_ANALYTICS):
                human.goto(page, sc_nav.SC_BASE + "/brand-analytics", wait_after=False)
            ba = ex.capture("브랜드 분석", "Brand analytics 메인")
            for l in ex.sub_links(ba, ["/brand-analytics", "/analytics/dashboard", "/brand-analytics/dashboard"], MAX_BA_PAGES):
                info = ex.visit(l["href"], "브랜드 분석", l["text"])
                # 하위 페이지에 새 보고서 링크가 또 있으면 그것도 봄(전체 상한 이내)
                if info and len([p for p in ex.pages if p["section"] == "브랜드 분석"]) < MAX_BA_PAGES:
                    for l2 in ex.sub_links(info, ["/brand-analytics"], 3):
                        if len([p for p in ex.pages if p["section"] == "브랜드 분석"]) >= MAX_BA_PAGES:
                            break
                        ex.visit(l2["href"], "브랜드 분석", l2["text"])

            # ---- 3. 광고 콘솔
            print("\n【3】 마케팅 → Ads console...")
            human.goto(page, sc_nav.SC_BASE + "/home", wait_after=False)
            sc_nav._settle(page)
            if not ex.menu_click(MARKETING, ADS_CONSOLE):
                human.goto(page, sc_nav.SC_BASE + "/cm?source=ngs", wait_after=False)
            ads = ex.capture("광고 콘솔", "Ads console 메인")
            host = urlparse(page.url).netloc
            must = ["/cm/", "/cm?", "advertising.amazon", "/ads/", "/sspa/", "/reports", "/measurement"]
            for l in ex.sub_links(ads, must, MAX_ADS_PAGES):
                if len([p for p in ex.pages if p["section"] == "광고 콘솔"]) > MAX_ADS_PAGES:
                    break
                if urlparse(l["href"]).netloc not in (host, "sellercentral.amazon.com", "advertising.amazon.com"):
                    continue
                info = ex.visit(l["href"], "광고 콘솔", l["text"])
                # 첫 캠페인 페이지의 다음 단계(광고 그룹 / 키워드 등), 조금만 봄
                if info and "/campaigns" in l["href"]:
                    for l2 in ex.sub_links(info, ["/cm/"], 4):
                        ex.visit(l2["href"], "광고 콘솔", l2["text"])
    finally:
        if ex is not None:
            (out_dir / "pages.json").write_text(json.dumps({"menu": menu, "pages": ex.pages}, ensure_ascii=False, indent=1),
                                                encoding="utf-8")
            out = write_summary(out_dir, menu, ex.pages)
            print(f"\n✅ 확인한 페이지 {len(ex.pages)}개. 결과 저장 위치: {out_dir}")
            if open_report:
                import os, sys
                if sys.platform.startswith("win"):
                    try:
                        os.startfile(str(out))
                    except Exception:
                        pass
    return out_dir


if __name__ == "__main__":
    run()

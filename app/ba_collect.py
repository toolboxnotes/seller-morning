"""
Brand Analytics 키워드 자료 수집(조회만)

메뉴 마케팅 → Brand analytics → Search Analytics:
  - Search Query Performance(브랜드 보기): 검색어의 검색량과 우리 브랜드의 노출 / 클릭 / 장바구니 / 구매 점유율
  - Search Catalog Performance: ASIN별 검색 노출·클릭·장바구니·구매
  - Top Search Terms: 인기 검색어 순위와 클릭 상위 상품(우리 키워드로 하나씩 조회)

데이터 받는 방법(순서대로):
  1. Generate Download 클릭: 바로 다운로드, 또는 Download Manager에서 생성될 때까지 기다렸다가 다운로드(2~3분마다 새로 고침)
  2. 다운로드를 못 받으면 페이지의 표를 직접 읽음(최대 몇 페이지)
같은 주 데이터는 한 번만 받음(Brand Analytics는 주 단위로 갱신).
"""

import random
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import ads_common as ac
import human
import sc_labels
import inventory as inv
import sc_nav
import sc_reports

BA_DIR = ac.REPORT_DIR / "brand_analytics"
BA_BASE = "/brand-analytics/dashboard/"
MARKETING = sc_labels.L("Marketing")
BRAND_ANALYTICS = sc_labels.L("Brand analytics")

DASHBOARDS = {
    "sqp": {"path": "query-performance", "label": "Search Query Performance", "view": "query-performance-brands-view",
            "brand": True, "dm_name": "Search Query Performance", "hint": ["Search Query"]},
    "scp": {"path": "brand-catalog-performance", "label": "Search Catalog Performance",
            "view": "brand-catalog-performance-default-view", "brand": False, "dm_name": "Search Catalog Performance",
            "hint": ["ASIN"]},
}
TST = {"path": "top-search-terms", "label": "Top Search Terms"}

DOWNLOAD_WAIT_MIN = 10      # Download Manager에서 최대 얼마나 기다릴지(그래도 안 되면 페이지 표를 읽음)
MAX_SCRAPE_PAGES = 8        # 다운로드가 안 될 때 최대 몇 페이지까지 표를 읽을지
MAX_TST_TERMS = 15          # Top Search Terms에서 최대 몇 개 단어를 조회할지

_TABLE_JS = r"""
() => {
  const txt = (e) => (e.innerText || e.textContent || '').replace(/\s+/g, ' ').trim();
  const leaf = Array.from(document.querySelectorAll('[role="columnheader"]'))
                    .filter(h => !h.querySelector('[role="columnheader"]')).map(txt);
  // 머리글 줄마다 따로 기록(위치 포함) — 묶음 머리글과 실제 열 이름을 구분하려고
  const headerRows = [];
  const rows = [];
  for (const r of document.querySelectorAll('[role="row"]')) {
    const hs = Array.from(r.querySelectorAll('[role="columnheader"]')).filter(h => !h.querySelector('[role="columnheader"]'));
    if (hs.length) {
      headerRows.push(hs.map(h => { const b = h.getBoundingClientRect(); return {t: txt(h), l: b.left, r: b.right}; }));
      continue;
    }
    let cells = Array.from(r.querySelectorAll(':scope > [role="cell"], :scope > [role="gridcell"], :scope > td'));
    if (!cells.length) cells = Array.from(r.children);
    const vals = cells.map(txt);
    if (vals.some(v => v)) rows.push(vals);
  }
  return {headers: leaf, headerRows: headerRows, rows: rows};
}
"""


# ================================================================ 페이지 열기

def _week_info(page) -> Tuple[Optional[str], List[str], Optional[str]]:
    """지금 선택된 주, 고를 수 있는 주(최근 → 예전), 브랜드 번호. 옵션이 컴포넌트 안(shadow DOM)에 있을 수 있음."""
    try:
        return page.evaluate("""() => {
            const w = document.querySelector('kat-dropdown#weekly-week');
            const b = document.querySelector('kat-dropdown#brand');
            let opts = [];
            if (w) {
              const all = Array.from(w.querySelectorAll('kat-option'))
                .concat(w.shadowRoot ? Array.from(w.shadowRoot.querySelectorAll('kat-option')) : []);
              opts = all.map(o => o.getAttribute('value')).filter(x => x);
            }
            return [w ? w.getAttribute('value') : null, Array.from(new Set(opts)), b ? b.getAttribute('value') : null];
        }""")
    except Exception:
        return None, [], None


def expected_week(today=None) -> str:
    """규정상 있어야 할 가장 최근 주: Amazon의 주는 일요일 ~ 토요일, 오늘(태평양 시간) 이전 가장 가까운 토요일."""
    from datetime import timedelta
    t = today or ac.pacific_today()
    d = t - timedelta(days=1)
    while d.weekday() != 5:          # 5 = 토요일
        d -= timedelta(days=1)
    return d.isoformat()


def week_text(week: str) -> str:
    from datetime import timedelta
    end = datetime.strptime(week, "%Y-%m-%d").date()
    return f"{end - timedelta(days=6)} ~ {end}"


def _page_error(page) -> str:
    try:
        t = page.inner_text("body")
    except Exception:
        return ""
    for pat in (r"Invalid Input", r"Please select a valid value", r"No data (found|available)", r"There is no data"):
        m = re.search(pat, t, re.I)
        if m:
            return m.group(0)
    return ""


def choose_week(page, label: str, path: str, base_params: str, view: str) -> str:
    """먼저 보고서를 열어 고를 수 있는 주 확인 → 가장 최근 주 사용, 그 주에 데이터가 없으면 이유를 적고 한 주 전(최대 3번)."""
    want = expected_week()
    week, weeks, _ = _week_info(page)
    cands = [w for w in weeks if w <= want] or weeks or [want]
    if cands[0] != want:
        print(f"   ⚠ 규정상 최근 주는 {week_text(want)}이지만, 아직 선택할 수 없습니다(아마존이 아직 공개하지 않음). "
              f"{week_text(cands[0])} 주를 사용합니다.")
    for w in cands[:3]:
        nav_to(page, label, path, f"{base_params}&weekly-week={w}&view-id={view}&country-id=us")
        err = _page_error(page)
        n = page.locator('[role="rowgroup"] [role="row"]').count()
        if not err and n:
            print(f"   기간: {week_text(w)} (주간)")
            return w
        why = err or "표에 데이터가 없음"
        print(f"   ⚠ {week_text(w)} 주를 열 수 없습니다({why}). 이전 주로 다시 시도합니다.")
    return cands[0]


def _brand_id(page) -> str:
    from urllib.parse import parse_qs, urlparse
    q = parse_qs(urlparse(page.url).query)
    if q.get("brand"):
        return q["brand"][0]
    try:
        return page.evaluate("() => { const b = document.querySelector('kat-dropdown#brand'); return b ? b.getAttribute('value') : ''; }") or ""
    except Exception:
        return ""


def open_ba(page) -> None:
    """메뉴 마케팅 → Brand analytics."""
    if "/brand-analytics" in page.url:
        return
    print("   마케팅 → Brand analytics 여는 중...")
    try:
        sc_nav.menu_go(page, [MARKETING, BRAND_ANALYTICS], ac.DEBUG_DIR)
    except Exception as e:
        print(f"   메뉴 조작 실패({str(e)[:60]}), Brand analytics를 바로 엽니다.")
        human.goto(page, sc_nav.SC_BASE + "/brand-analytics", wait_after=False)
    ac.settle(page)


def nav_to(page, label: str, path: str, params: str = "") -> None:
    """왼쪽 / 위쪽 Search Analytics 메뉴에서 보고서 이름 클릭, 안 되면 주소로 바로 엶."""
    target = sc_nav.SC_BASE + BA_BASE + path + (("?" + params) if params else "")
    if params:
        human.goto(page, target, wait_after=False)
        ac.settle(page)
        _wait_rows(page)
        return
    link = sc_nav._first_visible(page.locator(f'a[href*="{BA_BASE}{path}"]'))
    if link is None:
        sec = sc_nav._first_visible(page.get_by_text(re.compile(r"^\s*Search Analytics\s*$")))
        if sec is not None:
            human.click(page, sec)
            link = sc_nav._first_visible(page.locator(f'a[href*="{BA_BASE}{path}"]'))
    if link is not None and not params:
        print(f"   ‘{label}’ 클릭...")
        human.click(page, link, wait_after=False)
        try:
            page.wait_for_url(lambda u: path in u, timeout=20000)
        except Exception:
            pass
    else:
        human.goto(page, target, wait_after=False)
    ac.settle(page)
    _wait_rows(page)


def _wait_rows(page, timeout_s: int = 60) -> int:
    t0 = time.time()
    n = 0
    while time.time() - t0 < timeout_s:
        try:
            n = page.locator('[role="rowgroup"] [role="row"]').count()
        except Exception:
            n = 0
        if n:
            break
        page.wait_for_timeout(2000)
    human.read_page(page, max_screens=2)
    human.move_away(page)
    return n


# ================================================================ 다운로드

def _dm_rows(page) -> List[Dict]:
    """Download Manager 표의 줄(글자 + 줄 안의 다운로드 버튼 순번)."""
    try:
        return page.evaluate(r"""() => {
            const out = [];
            const rows = Array.from(document.querySelectorAll('[role="row"], tr'));
            rows.forEach((r, i) => {
              if (r.querySelector('[role="columnheader"], th')) return;
              const t = (r.innerText || '').replace(/\s+/g, ' ').trim();
              if (!t) return;
              const dl = Array.from(r.querySelectorAll('a, button, kat-button, kat-link'))
                  .some(b => /download/i.test((b.getAttribute('label') || '') + ' ' + (b.innerText || '') + ' ' + (b.getAttribute('aria-label') || '')))
                  || !!r.querySelector('kat-icon[name="file_download"], kat-icon[name="download"]')
                  || Array.from(r.querySelectorAll('span, div')).some(e => !e.children.length && /^\s*Download\s*$/i.test(e.textContent || ''));
              out.push({i: i, text: t, ready: dl});
            });
            return out;
        }""")
    except Exception:
        return []


def _dm_button(row):
    """Download Manager 한 줄의 Download(버튼일 수도, 글자 + 다운로드 아이콘일 수도 있음)."""
    for sel in ('kat-button[label*="Download" i] button', 'kat-button[label*="Download" i]', 'a:has-text("Download")',
                'button:has-text("Download")', 'kat-link:has-text("Download")'):
        btn = sc_nav._first_visible(row.locator(sel))
        if btn is not None:
            return btn
    btn = sc_nav._first_visible(row.get_by_text(re.compile(r"^\s*Download\s*$", re.I)))
    if btn is not None:
        return btn
    return sc_nav._first_visible(row.locator('kat-icon[name="file_download"]'))


def _click_wait_download(page, el, save_name: str, wait_s: int = 15) -> Optional[Path]:
    """누른 뒤 몇 초 기다림: 브라우저 다운로드가 있으면 저장."""
    got = []
    handler = lambda d: got.append(d)
    page.on("download", handler)
    try:
        human.click(page, el, wait_after=False)
        for _ in range(wait_s * 2):
            if got:
                break
            page.wait_for_timeout(500)
    finally:
        try:
            page.remove_listener("download", handler)
        except Exception:
            pass
    if not got:
        return None
    d = got[0]
    ac.RAW_DIR.mkdir(parents=True, exist_ok=True)
    out = ac.RAW_DIR / f"{save_name}{Path(d.suggested_filename).suffix or '.csv'}"
    d.save_as(str(out))
    return out


def _try_direct_download(page, btn, save_name: str) -> Optional[Path]:
    """Generate Download 클릭: 바로 받아질 때도, 옵션 창이 뜰 때도 있음(CSV 선택 후 확인)."""
    out = _click_wait_download(page, btn, save_name)
    if out is not None:
        return out
    modal = ac.first_visible(page, ['kat-modal[visible]', '[role="dialog"]'])
    if modal is not None:
        ac.dump(page, f"{save_name}_download_modal")
        csv_opt = ac.by_text(page, r"^\s*(CSV|\.csv)\s*$", roles=("radio", "option", "button", "menuitem"))
        if csv_opt is not None:
            human.click(page, csv_opt)
        confirm = ac.first_visible(page, ['kat-modal kat-button[variant="primary"] button', 'kat-modal kat-button[variant="primary"]',
                                          '[role="dialog"] button[type="submit"]'])
        if confirm is None:
            confirm = ac.by_text(page, r"^\s*(Generate|Generate Download|Download|Confirm|Submit|Request)\s*$", roles=("button",))
        if confirm is not None:
            return _click_wait_download(page, confirm, save_name)
    return None


def _from_download_manager(page, dm_name: str, week: str, save_name: str, wait_min: int) -> Optional[Path]:
    """Download Manager에서 이 보고서를 찾아 생성됐으면 다운로드, 아직이면 2~3분마다 새로 고침."""
    url = sc_nav.SC_BASE + "/brand-analytics/download-manager"
    print(f"   Download Manager에서 ‘{dm_name}’ 생성 대기...")
    human.goto(page, url, wait_after=False)
    ac.settle(page)
    t0 = time.time()
    while True:
        rows = _dm_rows(page)
        # 가장 최근 요청이 맨 위: 이름이 같은 줄만 봄
        match = [r for r in rows if dm_name.lower() in r["text"].lower()]
        if week:
            y, m, d_ = week.split("-")
            forms = [week, f"{int(m)}/{int(d_)}/{y}", f"{m}/{d_}/{y}",
                     datetime.strptime(week, "%Y-%m-%d").strftime("%b %d, %Y"),
                     datetime.strptime(week, "%Y-%m-%d").strftime("%b ") + str(int(d_)) + ", " + y]
            same_week = [r for r in match if any(f in r["text"] for f in forms)]
            if same_week or any(re.search(r"\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2}", r["text"]) for r in match):
                match = same_week
        ready = [r for r in match if r["ready"]]
        if ready:
            r = ready[0]
            row = page.locator('[role="row"], tr').nth(r["i"])
            btn = None
            btn = _dm_button(row)
            if btn is not None:
                print("   생성 완료, 다운로드...")
                return sc_reports._download(page, btn, ac.RAW_DIR, save_name, ac.DEBUG_DIR)
        mins = (time.time() - t0) / 60
        if mins >= wait_min:
            ac.dump(page, f"{save_name}_download_manager")
            return None
        wait = random.uniform(sc_reports.POLL_MIN_S, sc_reports.POLL_MAX_S)
        state = "생성 중" if match else "목록에 아직 없음"
        print(f"     {state}, 대기 {mins:.0f}분, {wait / 60:.1f}분 후 새로고침(최대 {wait_min}분)...")
        page.wait_for_timeout(int(wait * 1000))
        page.reload(wait_until="domcontentloaded")
        ac.settle(page)


def _scrape(page, key: str, max_pages: int = MAX_SCRAPE_PAGES) -> List[Dict]:
    """페이지 표 읽기(다운로드를 못 받을 때)."""
    # 페이지당 줄 수를 최대로
    try:
        dd = page.locator("kat-dropdown").filter(has=page.locator('kat-option[value="100"]')).first
        if dd.count() and dd.get_attribute("value") != "100":
            human.click(page, dd.locator(".select-header").first if dd.locator(".select-header").count() else dd)
            opt = dd.locator('kat-option[value="100"]').first
            opt.wait_for(state="visible", timeout=5000)
            human.click(page, opt)
            _wait_rows(page)
    except Exception:
        pass
    out: List[Dict] = []
    for p in range(1, max_pages + 1):
        data = page.evaluate(_TABLE_JS)
        rows = [v for v in data.get("rows", []) if v]
        if not rows:
            break
        n = max(len(v) for v in rows)
        headers = _headers_from_rows(key, data.get("headerRows") or [], n)
        if headers is None:
            leaf = data.get("headers", [])
            if len(leaf) > n:                     # 앞부분은 묶음 머리글(예: Top Clicked Product #1)
                leaf = leaf[-n:]
            headers = _dedupe(_name_headers(key, leaf))
        if p == 1 and headers and headers[0].lower().startswith("search funnel"):
            print("   ⚠ 표 머리글이 맞지 않아 페이지 표를 건너뜁니다(전체 다운로드 사용).")
            ac.dump(page, f"ba_{key}_header_mismatch")
            return []
        for vals in rows:
            if len(vals) == len(headers):
                out.append(dict(zip(headers, vals)))
        nxt = sc_nav._first_visible(page.get_by_role("button", name=re.compile(f"Go to page {p + 1}$")))
        if nxt is None:
            break
        human.click(page, nxt, wait_after=False)
        page.wait_for_timeout(1500)
        _wait_rows(page, 30)
    return out


def _clean_group(t: str) -> str:
    return re.sub(r"^\s*Search Funnel\s*-\s*", "", re.sub(r"\s+", " ", t)).strip()


def _headers_from_rows(key: str, hrows: List[List[Dict]], n: int) -> Optional[List[str]]:
    """두 줄 머리글: 마지막 줄이 실제 열 이름(데이터 칸 수와 같아야 함), 위 줄은 묶음(위치로 대응)."""
    if not hrows:
        return None
    idx = [i for i, r in enumerate(hrows) if len(r) == n]
    if not idx:
        return None
    li = idx[-1]
    leaf = hrows[li]
    groups = [g for row in hrows[max(0, li - 2):li] for g in row if g.get("t") and g["r"] > g["l"]]
    out = []
    for h in leaf:
        name = re.sub(r"\s+", " ", h["t"]).strip()
        c = (h["l"] + h["r"]) / 2
        grp = next((g for g in groups if g["l"] - 1 <= c <= g["r"] + 1), None)
        if key == "sqp" and grp is not None:
            name = f"{_clean_group(grp['t'])}: {name}"
        out.append(name)
    if key == "sqp" and not any(":" in h for h in out):
        return None                          # 위치를 못 얻으면(예: 너비 0) → 예전 방법
    return _dedupe(out)


def _bad_old_file(path: Path) -> bool:
    """예전 버전은 페이지 표를 전체 파일 이름으로 저장했고 머리글이 어긋났음(첫 열이 Search Funnel …)."""
    try:
        rows = ac.load_rows(path)
        return bool(rows) and str(next(iter(rows[0].keys()))).lower().startswith("search funnel")
    except Exception:
        return False


def _dedupe(headers: List[str]) -> List[str]:
    """같은 이름의 열(예: Top Clicked Product #1/#2/#3 아래 모두 ASIN): 두 번째부터 #2, #3를 붙임."""
    seen: Dict[str, int] = {}
    out = []
    for h in headers:
        seen[h] = seen.get(h, 0) + 1
        out.append(h if seen[h] == 1 else f"{h} #{seen[h]}")
    if any(h.endswith(" #2") for h in out):
        out = [(f"{h} #1" if f"{h} #2" in out else h) for h in out]
    return out


def _name_headers(key: str, leaf: List[str]) -> List[str]:
    """표 머리글이 두 줄이면(예: Impressions → Total Count / Brand Count) “Impressions: Total Count”로 합침."""
    groups = ["Impressions", "Clicks", "Cart Adds", "Purchases"]
    out, gi = [], -1
    for h in leaf:
        h = re.sub(r"\s+", " ", h).strip()
        if key == "sqp" and h.startswith("Total Count"):
            gi += 1
        if gi >= 0 and key == "sqp":
            out.append(f"{groups[min(gi, 3)]}: {h}")
        else:
            out.append(h)
    return out




def _load_pending() -> Dict:
    import json
    try:
        return json.loads((BA_DIR / "pending_downloads.json").read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_pending(p: Dict) -> None:
    import json
    BA_DIR.mkdir(parents=True, exist_ok=True)
    (BA_DIR / "pending_downloads.json").write_text(json.dumps(p, ensure_ascii=False, indent=1), encoding="utf-8")


def latest_file(key: str) -> Optional[Path]:
    """가장 최근 주 파일: 전체 다운로드(key_주.xlsx)가 있으면 그것, 없으면 페이지 표로 읽은 것(key_주_page.xlsx)."""
    weeks = {}
    for f in BA_DIR.glob(f"{key}_*.xlsx"):
        m = re.match(rf"{key}_(\d{{4}}-\d{{2}}-\d{{2}})(_page)?\.xlsx$", f.name)
        if m:
            weeks.setdefault(m.group(1), {})["page" if m.group(2) else "full"] = f
    if not weeks:
        return None
    w = weeks[max(weeks)]
    return w.get("full") or w.get("page")


def _request_download(page, btn) -> Optional[str]:
    """Generate Download → Select Download Type 창 → Comprehensive View 선택(없으면 Simple View) → Generate Download."""
    human.click(page, btn)
    page.wait_for_timeout(1200)
    modal = ac.first_visible(page, ['kat-modal[visible]', '[role="dialog"]', 'kat-modal'])
    for name in ("Comprehensive View", "Simple View"):
        opt = sc_nav._first_visible(page.get_by_role("radio", name=re.compile(rf"^\s*{name}", re.I)))
        if opt is None:
            opt = sc_nav._first_visible(page.locator("kat-radiobutton, label").filter(has_text=re.compile(rf"^\s*{name}", re.I)))
        if opt is None:
            continue
        try:
            disabled = opt.get_attribute("disabled") is not None or opt.get_attribute("aria-disabled") == "true"
        except Exception:
            disabled = False
        if disabled:
            continue
        print(f"   다운로드 유형: {name}")
        human.click(page, opt)
        break
    else:
        name = None
        if modal is None:
            return None
        ac.dump(page, "ba_download_type_modal")
    # 창 안의 Generate Download(페이지에 있는 것 말고)
    gen = None
    for sel in ('kat-modal kat-button[label="Generate Download"] button', 'kat-modal kat-button[label="Generate Download"]',
                '[role="dialog"] button:has-text("Generate Download")'):
        gen = sc_nav._first_visible(page.locator(sel))
        if gen is not None:
            break
    if gen is None:
        btns = page.get_by_role("button", name=re.compile(r"^\s*Generate Download\s*$"))
        if btns.count() >= 2:
            gen = btns.nth(btns.count() - 1)
    if gen is None:
        ac.dump(page, "ba_download_type_modal")
        return None
    human.click(page, gen)
    page.wait_for_timeout(1500)
    return name or "기본"


def fetch_api(key: str) -> Optional[Path]:
    """SP-API로 최근 주(없으면 그 전 주) SCP·SQP 받기 → {key}_{주}.xlsx(전체 데이터). SQP는 SCP에서 노출이 많은 ASIN 기준."""
    import sp_data
    from datetime import date as _d, timedelta as _td
    weeks = [expected_week()]
    weeks.append((_d.fromisoformat(weeks[0]) - _td(days=7)).isoformat())
    for wk in weeks:
        out = BA_DIR / f"{key}_{wk}.xlsx"
        if out.exists():
            print(f"   {DASHBOARDS[key]['label']} {week_text(wk)} 주 전체 데이터가 이미 있습니다.")
            return out
        if key == "scp":
            rows = sp_data.scp_week(wk)
        else:
            scp = BA_DIR / f"scp_{wk}.xlsx"
            src = ac.load_rows(scp) if scp.exists() else (sp_data.scp_week(wk) or [])
            asins = [str(r.get("ASIN", "")) for r in sorted(src, key=lambda r: -ac.num(r.get("Impressions")))
                     if ac.num(r.get("Impressions")) > 0]
            rows = sp_data.sqp_week(wk, asins)
        if rows:
            BA_DIR.mkdir(parents=True, exist_ok=True)
            ac.save_rows(rows, out, key)
            pend = _load_pending()
            if pend.pop(key, None):
                _save_pending(pend)
            print(f"   ✅ {DASHBOARDS[key]['label']} {week_text(wk)} 주(API): {len(rows)}행")
            return out
        print(f"   {week_text(wk)} 주 데이터가 아직 없어 그 전 주를 확인합니다.")
    return None


def fetch_dashboard(page, key: str) -> Tuple[Optional[Path], str]:
    """보고서 열기(최근 주) → 다운로드 요청(Comprehensive View, 생성되면 Download Manager에서 받음)
    → 그동안 페이지 표를 먼저 읽어 다운로드 전에는 그것으로 분석. SP-API를 먼저 시도."""
    import sp_api
    api = sp_api.try_api(f"Brand Analytics {key.upper()}", fetch_api, key)
    if api:
        return api, f"{api.name}(SP-API)"
    d = DASHBOARDS[key]
    open_ba(page)
    brand = _brand_id(page)
    base = (f"brand={brand}&" if d["brand"] and brand else "") + "reporting-range=weekly"
    print(f"   ‘{d['label']}’ 여는 중...")
    nav_to(page, d["label"], d["path"], f"{base}&weekly-week={expected_week()}&view-id={d['view']}&country-id=us")
    week = choose_week(page, d["label"], d["path"], base, d["view"])
    if d["path"] not in page.url:
        path = ac.dump(page, f"ba_{key}_wrong_page")
        raise RuntimeError(f"{d['label']} 페이지가 열리지 않았습니다(페이지 구조 저장: {path.name})")
    BA_DIR.mkdir(parents=True, exist_ok=True)
    out = BA_DIR / f"{key}_{week}.xlsx"
    if out.exists() and _bad_old_file(out):
        bad = out.with_name(f"{key}_{week}_old_mismatch.xlsx")
        out.replace(bad)
        print(f"   ⚠ 기존 {out.name}은 열이 어긋난 페이지 표라서 {bad.name}로 바꾸고 다시 받습니다.")
    if out.exists():
        print(f"   {d['label']} {week_text(week)} 주 데이터는 이미 있습니다. 그대로 사용합니다.")
        return out, f"{out.name}(이미 가져옴)"
    ac.dump(page, f"ba_{key}_page")
    pend = _load_pending()
    if pend.get(key, {}).get("week") != week:
        btn = ac.first_visible(page, ["kat-button#GenerateDownloadButton button", "kat-button#GenerateDownloadButton"])
        if btn is not None:
            print("   Generate Download 클릭...")
            kind = _request_download(page, btn)
            if kind:
                pend[key] = {"week": week, "dm_name": d["dm_name"], "kind": kind, "requested": datetime.now().isoformat()}
                _save_pending(pend)
                print("   다운로드 요청 완료(Download Manager에서 생성 후 받음, 최대 25분).")
            else:
                print("   ⚠ 다운로드 요청 창을 처리하지 못했습니다(페이지 구조를 debug에 저장).")
            nav_to(page, d["label"], d["path"], f"{base}&weekly-week={week}&view-id={d['view']}&country-id=us")
    else:
        print("   이 주의 다운로드는 이미 요청했습니다. Download Manager에서 받습니다.")
    # 먼저 페이지 표 읽기(다운로드가 아직일 때 사용)
    rows = _scrape(page, key)
    if rows:
        ac.save_rows(rows, BA_DIR / f"{key}_{week}_page.xlsx", key)
        print(f"   ✅ 페이지 표 {len(rows)}행 읽음(전체 다운로드 전 임시 데이터)")
    else:
        ac.dump(page, f"ba_{key}_no_table")
    return (BA_DIR / f"{key}_{week}_page.xlsx") if rows else None, f"{week_text(week)} 주, 페이지 표 {len(rows)}행 + 전체 다운로드 요청"


def collect_downloads(page, wait_min: float = 0) -> Dict[str, str]:
    """Download Manager：요청해 둔 다운로드가 생성되면 받기. wait_min=0이면 한 번만 확인."""
    pend = _load_pending()
    if not pend:
        return {}
    done: Dict[str, str] = {}
    t0 = time.time()
    human.goto(page, sc_nav.SC_BASE + "/brand-analytics/download-manager", wait_after=False)
    ac.settle(page)
    while pend:
        rows = _dm_rows(page)
        for key, info in list(pend.items()):
            week = info["week"]
            y, m, d_ = week.split("-")
            forms = [week, f"{int(m)}/{int(d_)}/{y}", f"{m}/{d_}/{y}"]
            match = [r for r in rows if info["dm_name"].lower() in r["text"].lower() and any(f in r["text"] for f in forms)]
            ready = [r for r in match if r["ready"]]
            if not ready:
                continue
            row = page.locator('[role="row"], tr').nth(ready[0]["i"])
            btn = None
            btn = _dm_button(row)
            if btn is None:
                continue
            print(f"   {info['dm_name']} ({week_text(week)}) 생성 완료, 다운로드...")
            raw = sc_reports._download(page, btn, ac.RAW_DIR, f"ba_{key}_{week}", ac.DEBUG_DIR)
            hint = DASHBOARDS.get(key, {}).get("hint")
            h, data = ac.read_any(raw, hint)
            if data:
                ac.save_rows(data, BA_DIR / f"{key}_{week}.xlsx", key)
                print(f"   ✅ {info['dm_name']} 전체 데이터 {len(data)}행")
                done[key] = f"{key}_{week}.xlsx({len(data)}행, {info.get('kind', '')})"
            pend.pop(key)
            _save_pending(pend)
        mins = (time.time() - t0) / 60
        if not pend or wait_min <= 0 or mins >= wait_min:
            break
        wait = random.uniform(sc_reports.POLL_MIN_S, sc_reports.POLL_MAX_S)
        print(f"     Download Manager 생성 중 {len(pend)}개, 대기 {mins:.0f}분, {wait / 60:.1f}분 후 새로고침...")
        page.wait_for_timeout(int(wait * 1000))
        page.reload(wait_until="domcontentloaded")
        ac.settle(page)
    if pend:
        ac.dump(page, "ba_download_manager")
    return done


def pending_text() -> str:
    p = _load_pending()
    return ", ".join(f"{v['dm_name']}({week_text(v['week'])})" for v in p.values())


# ================================================================ Top Search Terms

def fetch_top_search_terms(page, terms: List[str], brand: str = "") -> Tuple[Optional[Path], str]:
    """인기 검색어: 우리 키워드로 하나씩 조회해 순위와 클릭 상위 상품 기록."""
    if not brand:
        import sp_api
        brand = sp_api.brand()
    open_ba(page)
    print("   ‘Top Search Terms’ 여는 중...")
    nav_to(page, TST["label"], TST["path"],
           f"reporting-range=weekly&weekly-week={expected_week()}&view-id=top-search-terms-default-view&country-id=us")
    week = choose_week(page, TST["label"], TST["path"], "reporting-range=weekly", "top-search-terms-default-view")
    out = BA_DIR / f"tst_{week}.xlsx"
    if out.exists():
        print(f"   Top Search Terms {week} 주 데이터는 이미 있습니다. 그대로 사용합니다.")
        return out, f"{out.name}(이미 가져옴)"
    ac.dump(page, "ba_tst_page")
    rows: List[Dict] = []
    box = ac.first_visible(page, ['kat-input#search-term-freeform input', 'input[placeholder="Search by keyword"]'])
    brand_box = ac.first_visible(page, ['kat-input#brand-freeform input', 'input[placeholder="Search Brand name"]'])

    def refine():
        btn = ac.first_visible(page, ['kat-button[label="Refine results"] button', 'kat-button[label="Refine results"]'])
        if btn is not None:
            human.click(page, btn, wait_after=False)
        else:
            page.keyboard.press("Enter")
        page.wait_for_timeout(2500)
        _wait_rows(page, 30)

    if box is None:
        print("   ⚠ Top Search Terms 검색창을 찾지 못했습니다.")
    else:
        # 1) 우리 브랜드가 클릭 상위에 있는 검색어
        if brand_box is not None and brand:
            print(f"   브랜드 ‘{brand}’가 클릭 상위인 검색어 조회...")
            human.type_text(page, brand_box, brand, wait_after=False)
            refine()
            for r in _scrape(page, "tst", 2):
                r["조회 방법"] = f"브랜드 {brand}"
                rows.append(r)
            human.type_text(page, brand_box, "", wait_after=False)
        # 2) 키워드별 조회
        for t in terms[:MAX_TST_TERMS]:
            print(f"   검색어 조회: {t}")
            human.type_text(page, box, t, wait_after=False)
            refine()
            got = _scrape(page, "tst", 1)
            for r in got[:20]:
                r["조회 방법"] = f"키워드 {t}"
                r["조회한 키워드"] = t
            rows += got[:20]
            human.pause()
    if not rows:
        ac.dump(page, "ba_tst_no_data")
        return None, "Top Search Terms를 읽지 못함(페이지 구조를 debug에 저장)"
    ac.save_rows(rows, out, "tst")
    print(f"   ✅ Top Search Terms {week}: {len(rows)}행")
    return out, f"{out.name}({len(rows)}행, {week} 주)"


# ================================================================ 통일한 필드로 정리

def _pick(row: Dict, *pats) -> str:
    keys = list(row.keys())
    for p in pats:
        rx = re.compile(p, re.I)
        for k in keys:
            if rx.search(k):
                return row[k]
    return ""


def parse_sqp(rows: List[Dict]) -> List[Dict]:
    out = []
    for r in rows:
        q = _pick(r, r"^search query$", r"^query$")
        if not q:
            continue
        g = lambda grp, leaf: ac.num(_pick(r, rf"^{grp}\W*{leaf}", rf"{grp}.*{leaf}"))
        out.append({
            "query": q.strip(),
            "score": ac.num(_pick(r, r"search query score")),
            "volume": ac.num(_pick(r, r"search query volume")),
            "imp_total": g("impressions", "total"), "imp_brand": g("impressions", "(brand|asin) count"),
            "imp_share": g("impressions", "(brand|asin) share"),
            "clk_total": g("clicks", "total"), "clk_rate": g("clicks", "click rate"),
            "clk_brand": g("clicks", "(brand|asin) count"), "clk_share": g("clicks", "(brand|asin) share"),
            "cart_total": g("cart adds", "total"), "cart_brand": g("cart adds", "(brand|asin) count"),
            "cart_share": g("cart adds", "(brand|asin) share"),
            "pur_total": g("purchases", "total"), "pur_rate": g("purchases", "purchase rate"),
            "pur_brand": g("purchases", "(brand|asin) count"), "pur_share": g("purchases", "(brand|asin) share"),
        })
    return out


def parse_scp(rows: List[Dict]) -> List[Dict]:
    out = []
    for r in rows:
        asin = ""
        for v in r.values():
            m = re.search(r"\bB0[A-Z0-9]{8}\b", str(v))
            if m:
                asin = m.group(0)
                break
        if not asin:
            continue
        out.append({
            "asin": asin,
            "title": _pick(r, r"title", r"product name"),
            "impressions": ac.num(_pick(r, r"impressions?\W*(count|impressions)?$", r"impression")),
            "clicks": ac.num(_pick(r, r"^clicks?\W*(count|clicks)?$", r"clicks?:\s*clicks", r"clicks? count")),
            "ctr": ac.num(_pick(r, r"click rate", r"ctr")),
            "cart_adds": ac.num(_pick(r, r"^cart adds?\W*(count|cart adds)?$", r"cart adds?:\s*cart", r"cart add count")),
            "purchases": ac.num(_pick(r, r"^purchases?\W*(count|purchases)?$", r"purchases?:\s*purch", r"purchase count")),
            "sales": ac.num(_pick(r, r"search traffic sales", r"sales")),
            "cvr": ac.num(_pick(r, r"conversion rate")),
            "price": ac.num(_pick(r, r"price")),
        })
    return out

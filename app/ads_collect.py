"""
광고 콘솔(Ads console) 자료 수집(조회만: 광고 설정은 바꾸지 않음)

두 가지 출처로 서로 보완:
  A. 캠페인 관리 페이지의 표 → Export(바로 다운로드 가능)
     캠페인(그리고 왼쪽 메뉴의 광고 그룹 / 타겟 / 검색어 보기), 기간 전날·최근 7일
  B. Reporting(새 보고서) → 템플릿으로 보고서 만들기: 최근 30일, 하루 단위
     SP 검색어 / 타겟 / 캠페인 / 광고 상품, SB 검색어
     보고서 생성 시간이 일정하지 않아서: 먼저 요청 → 다른 자료 수집 → 돌아와 2~3분마다 확인 → 시간이 지나면 있는 데이터로 먼저 보고서,
     다음에 같은 메뉴를 실행하면 남은 것만 받음. 보고서를 매일 자동 실행으로 설정해 두면 다음부터는 열 때 이미 만들어져 있음.
"""

import random
import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, urlparse

import ads_common as ac
import human
import sc_labels
import inventory as inv
import sc_nav
import sc_reports

MARKETING = sc_labels.L("Marketing")
ADS_CONSOLE = sc_labels.L("Ads console")

# 보고서: key, 템플릿 이름, 광고 유형
# 보고서: 모두 “Search term” 템플릿에서 시작해 필요한 열(차원)로 바꿈. 광고 유형은 거르지 않고 “Ad product” 열로 SP / SB / SD 구분.
#   key, 선택할 차원, 해제할 차원
REPORTS = [
    ("sp_search_term", ["Date", "Ad product", "Targeting", "Targeting match type"], ["Date range"]),
    ("sp_targeting", ["Date", "Ad product", "Targeting", "Targeting match type", "Target bid"], ["Date range", "Search term"]),
    ("sp_campaign", ["Date", "Ad product", "Campaign budget amount", "Placement classification"], ["Date range", "Search term"]),
    ("sp_advertised_product", ["Date", "Ad product", "Advertised product ID", "Advertised product SKU"],
     ["Date range", "Search term"]),
]
TEMPLATE_ID = "search_term"
REPORT_PREFIX = "AIAGENT"
REPORT_RANGE = "Last 7 days"     # 보고서 기간(Report period): 최근 7일, 하루 단위(Date 열) → 전날과 7일 모두 여기서 계산
def _email() -> str:
    """광고 보고서 Send to 메일: 화면 프로그램 ‘설정’의 확인 메일(비어 있으면 넣지 않음)."""
    try:
        import sp_api
        return sp_api.report_email()
    except Exception:
        return ""
REPORT_WAIT_MIN = 40          # 보고서를 최대 얼마나 기다릴지(넘으면 있는 데이터로 먼저)
ADS_HOST = "advertising.amazon.com"

DATE_PRESETS = {"yesterday": "Yesterday", "7d": "Last 7 days"}
MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()


# ================================================================ 광고 콘솔 열기

def open_ads(page) -> str:
    """메뉴 마케팅 → Ads console. entityId를 돌려줌(다른 페이지를 열 때 필요)."""
    if ADS_HOST not in page.url:
        print("   마케팅 → Ads console 여는 중...")
        try:
            if not sc_nav.open_menu(page):
                raise RuntimeError("메뉴 버튼 없음")
            g = sc_nav.find_visible(page, MARKETING)
            if g is None:
                raise RuntimeError("메뉴에 Marketing 없음")
            human.click(page, g)
            link = sc_nav.find_visible(page, ADS_CONSOLE, prefer_link=True)
            if link is None:
                raise RuntimeError("메뉴에 Ads console 없음")
            target = inv._click_and_follow(page, link)
            if target is not page:
                url = target.url
                target.close()
                human.goto(page, url, wait_after=False)
        except Exception as e:
            print(f"   메뉴 조작 실패({str(e)[:60]}), Ads console을 바로 엽니다.")
            human.goto(page, sc_nav.SC_BASE + "/cm?source=ngs", wait_after=False)
        try:
            page.wait_for_url(lambda u: ADS_HOST in u, timeout=60000)
        except Exception:
            pass
        ac.settle(page)
    return entity_id(page)


def entity_id(page) -> str:
    q = parse_qs(urlparse(page.url).query)
    if q.get("entityId"):
        return q["entityId"][0]
    try:
        href = page.evaluate("() => { const a = document.querySelector('a[href*=\"entityId=\"]'); return a ? a.href : ''; }")
        q = parse_qs(urlparse(href).query)
        return q.get("entityId", [""])[0]
    except Exception:
        return ""


def ads_url(path: str, eid: str) -> str:
    sep = "&" if "?" in path else "?"
    return f"{ac.ADS_BASE}{path}{sep}entityId={eid}" if eid else f"{ac.ADS_BASE}{path}"


# ================================================================ A. 표 내보내기

def _fmt_range(a: date, b: date) -> List[str]:
    """페이지 날짜 버튼에 나올 수 있는 표기."""
    m = lambda d: f"{MONTHS[d.month - 1]} {d.day}"
    if a == b:
        return [f"{m(a)}, {a.year}"]
    return [f"{m(a)} - {m(b)}, {b.year}", f"{m(a)} – {m(b)}, {b.year}", f"{m(a)}, {a.year} - {m(b)}, {b.year}"]


def set_date_range(page, preset: str, rng: Tuple[date, date]) -> bool:
    """표 위의 날짜 버튼 → 미리 정해진 기간(Yesterday / Last 7 days) 선택 → 표시된 날짜 확인."""
    btn = ac.first_visible(page, ['button[id$="dateRangeFilter:openContainer"]', 'button[aria-label="Open date range picker"]'])
    if btn is None:
        ac.dump(page, "ads_no_date_button")
        return False
    want = _fmt_range(*rng)
    try:
        if btn.inner_text().strip() in want:
            return True
    except Exception:
        pass
    print(f"   날짜 범위 ‘{preset}’ 선택...")
    human.click(page, btn)
    opt = ac.by_text(page, rf"^\s*{re.escape(preset)}\s*$", roles=("option", "menuitem", "radio", "button", "tab"))
    if opt is None:
        ac.dump(page, "ads_date_popover")
        page.keyboard.press("Escape")
        return False
    human.click(page, opt)
    ok_btn = ac.by_text(page, r"^\s*(Save|Apply|Update|Done|저장|적용)\s*$", roles=("button",))
    if ok_btn is not None:
        human.click(page, ok_btn, wait_after=False)
    ac.settle(page, 15000)
    try:
        shown = btn.inner_text()
    except Exception:
        shown = ""
    if shown.strip() not in want:
        print(f"   ⚠ 날짜 표시가 예상과 다릅니다: {shown.strip()[:40]} (예상: {want[0]})")
    return True


def export_table(page, save_name: str) -> Optional[Path]:
    """표 위의 Export → CSV(또는 첫 번째 항목) 선택 → 다운로드."""
    btn = ac.first_visible(page, ['button[id$="udm-export-dropdown"]', 'button[aria-label="Export"]'])
    if btn is None:
        btn = ac.by_text(page, r"^\s*Export\s*$", roles=("button",))
    if btn is None:
        ac.dump(page, f"{save_name}_no_export")
        return None
    human.click(page, btn)
    opts = page.get_by_role("option")
    target = None
    try:
        texts = [opts.nth(i).inner_text().strip() for i in range(min(opts.count(), 10))]
    except Exception:
        texts = []
    for i, t in enumerate(texts):
        if re.search(r"csv|download|excel|xlsx", t, re.I) and not re.search(r"create", t, re.I):
            target = opts.nth(i)
            break
    if target is None:
        print(f"   Export 메뉴에 바로 받는 옵션이 없어 건너뜁니다({', '.join(texts)[:60]}).")
        page.keyboard.press("Escape")
        return None
    print(f"   Export → {(texts[0] if texts else 'Export')[:40]}...")
    try:
        return sc_reports._download(page, target, ac.RAW_DIR, save_name, ac.DEBUG_DIR)
    except Exception as e:
        print(f"   ⚠ 내보내기 실패: {str(e)[:100]}")
        return None


def export_views(page, eid: str, today: date) -> Dict[str, Path]:
    """캠페인 표(그리고 찾을 수 있는 광고 그룹 / 타겟 / 검색어 보기)를 전날·최근 7일로 내보내기."""
    rng = ac.ranges(today)
    dd = ac.day_dir(today)
    views = {"campaigns": ads_url("/campaign-manager/all-campaigns", eid)}
    views.update(_discover_views(page, eid))
    got: Dict[str, Path] = {}
    for view, url in views.items():
        for key, preset in DATE_PRESETS.items():
            out = dd / f"cm_{view}_{key}.xlsx"
            if out.exists():
                got[f"{view}_{key}"] = out
                continue
            if page.url.split("#")[0] != url.split("#")[0]:
                print(f"   광고 관리자 ‘{view}’ 여는 중...")
                human.goto(page, url, wait_after=False)
                ac.settle(page)
                human.read_page(page, max_screens=2)
                human.move_away(page)
            if not set_date_range(page, preset, rng[key]):
                continue
            raw = export_table(page, f"cm_{view}_{key}_{today}")
            if raw is None:
                continue
            h, rows = ac.read_any(raw, ["Campaign", "Search term", "Target", "Ad group"])
            if rows:
                ac.save_rows(rows, out, view)
                got[f"{view}_{key}"] = out
                print(f"   ✅ {view} ({preset}): {len(rows)}행")
    return got


def _discover_views(page, eid: str) -> Dict[str, str]:
    """왼쪽 메뉴의 캠페인 전체 보기(광고 그룹 / 타겟 / 검색어). 있으면 같이 내보냄."""
    found = {}
    try:
        if "/campaign-manager" not in page.url:
            human.goto(page, ads_url("/campaign-manager/all-campaigns", eid), wait_after=False)
            ac.settle(page)
        nav = ac.first_visible(page, ['button[aria-label="Open side navigation"]'])
        if nav is not None:
            human.click(page, nav)
        links = page.evaluate("""() => Array.from(document.querySelectorAll('a[href*="campaign-manager"]'))
            .map(a => [(a.innerText || a.getAttribute('aria-label') || '').trim(), a.href])""")
        for text, href in links:
            t = text.lower()
            for key, pat in (("ad_groups", r"^ad groups?$"), ("targeting", r"^(targeting|targets)$"),
                             ("search_terms", r"^search terms?$")):
                if re.match(pat, t) and key not in found:
                    found[key] = href
        page.keyboard.press("Escape")
    except Exception:
        pass
    if found:
        print("   광고 관리자에서 추가로 찾은 화면: " + ", ".join(found))
    else:
        ac.dump(page, "ads_side_nav")
    return found


# ================================================================ B. Reporting(새 보고서)

_ROWS_JS = r"""
() => {
  const out = [];
  const rows = Array.from(document.querySelectorAll('[role="row"], tr'));
  rows.forEach((r, i) => {
    if (r.querySelector('[role="columnheader"], th')) return;
    const t = (r.innerText || '').replace(/\s+/g, ' ').trim();
    if (t) out.push({i: i, text: t});
  });
  return out;
}
"""
DATE_RE = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (\d{1,2}), (\d{4})")
STATUS_RE = re.compile(r"\b(Completed|In progress|Pending|Processing|Queued|Running|Failed|Cancelled)\b", re.I)


def report_name(key: str, day: Optional[date] = None) -> str:
    """오늘 쓸 보고서 이름(한 번만 실행하므로 이름에 날짜를 넣음)."""
    return f"{REPORT_PREFIX}_{key.upper()}_{(day or date.today()):%Y%m%d}"


def _open_reporting(page, eid: str) -> None:
    if "/reporting" in page.url and "history" not in page.url and "/new" not in page.url:
        page.reload(wait_until="domcontentloaded")
    else:
        print("   Reporting 여는 중...")
        human.goto(page, ads_url("/reporting", eid), wait_after=False)
    ac.settle(page)
    human.read_page(page, max_screens=1)


def _report_rows(page) -> List[Dict]:
    out = []
    for r in page.evaluate(_ROWS_JS):
        dates = [date(int(y), MONTHS.index(m) + 1, int(d)) for m, d, y in DATE_RE.findall(r["text"])]
        st = STATUS_RE.search(r["text"])
        r["status"] = st.group(1).lower() if st else ""
        r["last_run"] = dates[2] if len(dates) >= 3 else (dates[-1] if dates else None)
        out.append(r)
    return out


def report_status(page, name: str) -> Optional[Dict]:
    for r in _report_rows(page):
        if name in r["text"]:
            return r
    return None


def _choose(page, label_pat: str, option_pat: str) -> bool:
    """설정 페이지의 한 항목: 라벨 근처의 입력 요소를 찾아 옵션 선택."""
    rx_opt = re.compile(option_pat, re.I)
    # 1) 바로 누를 수 있는 옵션(라디오 / 탭 버튼)
    for role in ("radio", "tab", "checkbox", "button"):
        el = sc_nav._first_visible(page.get_by_role(role, name=rx_opt))
        if el is not None:
            try:
                if el.get_attribute("aria-checked") == "true" or el.get_attribute("aria-selected") == "true":
                    return True
            except Exception:
                pass
            if role != "button":
                human.click(page, el)
                return True
    # 2) 라벨 뒤의 드롭다운
    lab = sc_nav._first_visible(page.get_by_text(re.compile(label_pat, re.I)))
    if lab is None:
        return False
    box = lab.locator("xpath=following::*[self::button or self::select or @role='combobox' or @role='listbox'][1]")
    if not box.count():
        return False
    box = box.first
    try:
        if box.evaluate("e => e.tagName") == "SELECT":
            box.select_option(label=box.locator("option").filter(has_text=rx_opt).first.inner_text())
            human.pause()
            return True
        if rx_opt.search(box.inner_text() or ""):
            return True
    except Exception:
        pass
    human.click(page, box)
    opt = ac.by_text(page, option_pat, roles=("option", "menuitem", "radio", "checkbox"))
    if opt is None:
        page.keyboard.press("Escape")
        return False
    human.click(page, opt)
    return True


def _dim_label(page, name: str):
    """열 설정 왼쪽 차원 트리에서 이름이 정확히 같은 항목(label, 안에 checkbox)."""
    rx = re.compile(rf"^\s*{re.escape(name)}\s*$")
    loc = page.locator('#column-picker li[role="treeitem"] label').filter(has=page.locator("p", has_text=rx))
    return loc.first if loc.count() else None


def _set_dim(page, name: str, on: bool) -> bool:
    lab = _dim_label(page, name)
    if lab is None:
        print(f"     ⚠ 열 ‘{name}’을(를) 찾지 못했습니다.")
        return False
    box = lab.locator('input[type="checkbox"]').first
    try:
        checked = box.is_checked()
    except Exception:
        checked = None
    if checked == on:
        return True
    try:
        lab.scroll_into_view_if_needed(timeout=5000)
    except Exception:
        pass
    print(f"     열 {'추가' if on else '제거'}: {name}")
    target = lab.locator("p").first if lab.locator("p").count() else lab     # 글자를 클릭(label 빈 곳을 누르면 안 될 때가 있음)
    human.click(page, target, wait_after=False)
    page.wait_for_timeout(600)
    try:
        return box.is_checked() == on
    except Exception:
        return True


def _pick_from_button(page, current_pat: str, option_pat: str) -> bool:
    """“Last 7 days”“Once” 같은 버튼 → 뜨는 목록에서 하나 선택(→ 확인 버튼이 있으면 누름)."""
    btn = sc_nav._first_visible(page.get_by_role("button", name=re.compile(current_pat, re.I)))
    if btn is None:
        return False
    human.click(page, btn)
    opt = ac.by_text(page, option_pat, roles=("option", "menuitem", "radio", "button", "tab", "listitem"))
    if opt is None:
        page.keyboard.press("Escape")
        return False
    human.click(page, opt)
    ok = ac.by_text(page, r"^\s*(Apply|Save|Done|Update|확인|적용)\s*$", roles=("button",))
    if ok is not None:
        human.click(page, ok)
    return True


def _set_period(page) -> bool:
    """Report period 날짜 버튼 → 왼쪽 미리 정한 기간(Today / Yesterday / Last 7 days / …) 선택 → Save."""
    btn = ac.first_visible(page, ['button[id^="report-range-form:date-range-row-control"]'])
    if btn is None:
        return _pick_from_button(page, r"Last \d+ days|Yesterday|Today|This|Custom", rf"^\s*{re.escape(REPORT_RANGE)}\s*$")
    try:
        if btn.inner_text().strip() == REPORT_RANGE:
            return True
    except Exception:
        pass
    human.click(page, btn)
    page.wait_for_timeout(800)
    ac.dump(page, "report_date_picker")
    rx = re.compile(rf"^\s*{re.escape(REPORT_RANGE)}\s*$")
    cands = page.get_by_text(rx)
    opt = None
    for i in range(min(cands.count(), 10)):
        c = cands.nth(i)
        try:
            inside_btn = c.evaluate("(e, b) => b.contains(e)", btn.element_handle())
        except Exception:
            inside_btn = False
        if not inside_btn and c.is_visible():
            opt = c
            break
    if opt is None:
        page.keyboard.press("Escape")
        return False
    human.click(page, opt)
    save = sc_nav._first_visible(page.get_by_role("button", name=re.compile(r"^\s*Save\s*$")))
    if save is not None:
        human.click(page, save)
    page.wait_for_timeout(600)
    try:
        return btn.inner_text().strip() == REPORT_RANGE
    except Exception:
        return True


def _add_email(page, email: str) -> bool:
    """Delivery settings → Send to: 메일 입력 → Add."""
    box = ac.first_visible(page, ['input[placeholder="Enter email addresses"]', 'input[placeholder*="email" i]'])
    if box is None:
        return False
    try:
        if email in page.inner_text("body").split("Send to", 1)[-1][:400]:
            return True
    except Exception:
        pass
    human.type_text(page, box, email, wait_after=False)
    add = box.locator("xpath=following::button[normalize-space()='Add'][1]")
    if add.count():
        human.click(page, add.first)
    else:
        page.keyboard.press("Enter")
    return True


def create_report(page, eid: str, key: str, daily_subscription: bool = False) -> bool:
    """Search term 템플릿 → 보고서 이름·열·기간(최근 30일)·주기(매일) 변경 → Submit. 보고서만 만들 뿐 광고는 바꾸지 않음."""
    add, remove = next((a, r) for k, a, r in REPORTS if k == key)
    name = report_name(key)
    print(f"   보고서 만들기: {name}")
    _open_reporting(page, eid)
    card_btn = _template_button(page, "Search term")
    if card_btn is not None:
        human.click(page, card_btn, wait_after=False)
    else:
        human.goto(page, ads_url(f"/reporting/new?templateId={TEMPLATE_ID}", eid), wait_after=False)
    try:
        page.wait_for_url(lambda u: "/reporting/new" in u, timeout=20000)
    except Exception:
        pass
    ac.settle(page)
    if "/reporting/new" not in page.url or ac.first_visible(page, ["#urc_frb_create_report_button"]) is None:
        ac.dump(page, f"report_form_missing_{key}")
        print("   ⚠ 보고서 만들기 화면이 열리지 않았습니다.")
        return False
    # 보고서 이름
    box = page.locator("xpath=//*[normalize-space()='Report name']/following::input[1]").last
    if box.count():
        human.type_text(page, box, name, wait_after=False)
    # 열: 먼저 해제(Date range와 Date는 같이 고를 수 없어 Date range를 빼고 Date 선택), 그다음 선택, 마지막에 확인
    for d in remove:
        _set_dim(page, d, False)
    for d in add:
        _set_dim(page, d, True)
    for d in add:
        lab = _dim_label(page, d)
        try:
            ok = lab is not None and lab.locator('input[type="checkbox"]').first.is_checked()
        except Exception:
            ok = True
        if not ok:
            page.wait_for_timeout(800)
            if not _set_dim(page, d, True):
                print(f"     ⚠ 열 ‘{d}’ 선택이 반영되지 않았습니다.")
    # 기간
    if not _set_period(page):
        print(f"   ⚠ 기간을 ‘{REPORT_RANGE}’로 바꾸지 못했습니다(페이지 구조를 debug에 저장).")
    # 메일로 보내기
    email = _email()
    if email and not _add_email(page, email):
        print("   ⚠ Send to 이메일 입력란을 찾지 못했습니다.")
    # 주기
    if daily_subscription and not _pick_from_button(page, r"^\s*Once\s*$", r"^\s*Daily\s*$"):
        print("   ⚠ 보고서 빈도를 ‘Daily’로 바꾸지 못했습니다(한 번만 실행).")
    ac.dump(page, f"report_form_filled_{key}")
    submit = ac.first_visible(page, ["#urc_frb_create_report_button"])
    human.click(page, submit, wait_after=False)
    try:
        page.wait_for_url(lambda u: "/reporting/new" not in u, timeout=30000)
    except Exception:
        ac.dump(page, f"report_submit_failed_{key}")
        print("   ⚠ Submit 후 화면이 바뀌지 않았습니다(입력 오류일 수 있음, 페이지 구조를 debug에 저장).")
        cancel = ac.first_visible(page, ["#urc_cancel_subscription_button"])
        if cancel is not None:
            human.click(page, cancel)
        return False
    ac.settle(page)
    print("   보고서 요청 완료.")
    return True


def _template_button(page, template: str):
    """템플릿 카드(제목 = 템플릿 이름)의 Use template 버튼."""
    head = page.locator("h1, h2, h3, h4, h5, [role='heading']").filter(has_text=re.compile(rf"^\s*{re.escape(template)}\s*$", re.I))
    for i in range(min(head.count(), 5)):
        h = head.nth(i)
        card = h.locator("xpath=ancestor::*[.//button[normalize-space()='Use template']][1]")
        if card.count():
            b = sc_nav._first_visible(card.first.get_by_role("button", name=re.compile(r"Use template", re.I)))
            if b is not None:
                return b
    return None


def download_report(page, eid: str, key: str, row: Dict, today_kst: date) -> Optional[Path]:
    """보고서 줄 → 가장 최근 실행 파일 다운로드."""
    name = report_name(key)
    rows_loc = page.locator('[role="row"], tr')
    loc = rows_loc.nth(row["i"])
    btn = None
    for sel in ('a:has-text("Download")', 'button:has-text("Download")', '[aria-label*="Download" i]'):
        btn = sc_nav._first_visible(loc.locator(sel))
        if btn is not None:
            break
    if btn is None:
        link = sc_nav._first_visible(loc.locator('a[href*="/reporting/history"], a[href*="/reports/history"]'))
        if link is None:
            link = sc_nav._first_visible(loc.get_by_role("link", name=re.compile(re.escape(name))))
        if link is not None:
            human.click(page, link, wait_after=False)
            ac.settle(page)
            btn = ac.by_text(page, r"^\s*Download\s*", roles=("button", "link", "menuitem"))
    if btn is None:
        actions = sc_nav._first_visible(loc.locator('button[aria-label*="ction" i], button[aria-haspopup]'))
        if actions is not None:
            human.click(page, actions)
            btn = ac.by_text(page, r"^\s*Download\s*", roles=("menuitem", "option", "button", "link"))
    if btn is None:
        ac.dump(page, f"report_no_download_{key}")
        return None
    return sc_reports._download(page, btn, ac.RAW_DIR, f"{key}_{today_kst}", ac.DEBUG_DIR)


def request_reports(page, eid: str, state: Dict, only: Optional[List[str]] = None) -> Dict:
    """보고서마다: 오늘 생성된 것이 있으면 → 다운로드 대기로 기록, 없으면 → 새로 만듦. only: 이 보고서만 처리(예: 전에 요청에 실패한 것)."""
    today_kst = date.today()
    _open_reporting(page, eid)
    ac.dump(page, "reporting_list")
    for key, *_ in REPORTS:
        if only is not None and key not in only:
            continue
        out = ac.day_dir() / f"rp_{key}.xlsx"
        if out.exists():
            state[key] = {"status": "done", "file": str(out)}
            continue
        row = report_status(page, report_name(key))
        if row and row["status"] == "completed":
            state[key] = {"status": "ready", "requested": None}
            continue
        if row and row["status"] in ("in progress", "pending", "processing", "queued", "running"):
            state[key] = {"status": "waiting", "requested": state.get(key, {}).get("requested") or datetime.now().isoformat()}
            continue
        if row and row["status"] == "failed":
            print(f"   ⚠ {report_name(key)} 생성이 실패했습니다. 새로 요청합니다.")
        ok = create_report(page, eid, key)
        state[key] = {"status": "waiting" if ok else "failed", "requested": datetime.now().isoformat()}
        _open_reporting(page, eid)
    return state


def collect_reports(page, eid: str, state: Dict, wait_min: int = REPORT_WAIT_MIN) -> Dict:
    """생성 상태 확인, 생성된 것은 다운로드. wait_min=0이면 한 번만 보고 기다리지 않음."""
    today_kst = date.today()
    t0 = time.time()
    while True:
        _open_reporting(page, eid)
        rows = _report_rows(page)
        pending = []
        for key, *_ in REPORTS:
            st = state.get(key, {})
            if (ac.day_dir() / f"rp_{key}.xlsx").exists():
                state[key] = {**st, "status": "done", "file": str(ac.day_dir() / f"rp_{key}.xlsx")}
                continue
            if st.get("status") in ("done", "failed"):
                continue
            cand = [r for r in rows if report_name(key) in r["text"]]
            cand.sort(key=lambda r: (r["last_run"] or date.min), reverse=True)
            row = cand[0] if cand else None
            if row and row["status"] == "completed":
                print(f"   {report_name(key)} 생성 완료, 다운로드...")
                raw = download_report(page, eid, key, row, today_kst)
                if raw is not None:
                    h, data = ac.read_any(raw, ["Campaign", "Date", "Search term", "Targeting", "ASIN"])
                    out = ac.day_dir() / f"rp_{key}.xlsx"
                    ac.save_rows(data, out, key)
                    state[key] = {"status": "done", "file": str(out), "rows": len(data)}
                    print(f"   ✅ {report_name(key)}: {len(data)}행")
                    ac.save_state({"date": ac.pacific_today().isoformat(), "reports": state})
                    _open_reporting(page, eid)
                    rows = _report_rows(page)
                    continue
            if row and row["status"] == "failed":
                state[key] = {"status": "failed", "note": "아마존에서 보고서 생성 실패"}
                continue
            pending.append(key)
        ac.save_state({"date": ac.pacific_today().isoformat(), "reports": state})
        mins = (time.time() - t0) / 60
        if not pending or mins >= wait_min or wait_min <= 0:
            return state
        wait = random.uniform(sc_reports.POLL_MIN_S, sc_reports.POLL_MAX_S)
        print(f"     생성 중인 보고서 {len(pending)}개, 대기 {mins:.0f}분, "
              f"{(datetime.now() + timedelta(seconds=wait)):%H:%M:%S}에 다시 확인(최대 {wait_min}분)...")
        page.wait_for_timeout(int(wait * 1000))

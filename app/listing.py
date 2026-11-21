"""
Listing 점검(1단계: 읽고 제안만, 리스팅은 바꾸지 않음)

자료 출처:
  1. All Listings Report(재고 메뉴가 하루 한 번 받음, 오늘 것이 없으면 다시 받음): 제목, 상태
  2. Category Listings Report(상품 보고서 페이지, 하루 한 번): 불릿, 백엔드 검색어, 이미지, Item Highlights
  3. 상품 상세 페이지(amazon.com/dp/ASIN, 광고비 많은 순, 한 번에 최대 MAX_DETAIL개, 3일 안에 본 것은 건너뜀):
     실제 표시 제목, 불릿, 이미지 수, 동영상, A+, 평점, 구매 버튼, 메인 이미지 흰 배경
  4. Seller Central 상품 관리 페이지: 검색 제한 / 수정 필요 / Amazon 제안 변경(Review Listing Changes) — 페이지만 읽음

규칙 근거: 2026-07-27부터 제목 최대 75자(미디어 제외), 넘으면 Amazon이 AI로 차례로 다시 씀;
같은 단어는 최대 2번; ! $ ? _ { } ^ ¬ ¦ 등 문자 금지; 홍보 문구 금지; Item Highlights 최대 125자.
"""

import base64
import json
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import ads_common as ac
import human
import inventory as inv
import sc_nav
import sc_reports

BASE_DIR = Path(__file__).resolve().parent.parent
REPORT_DIR = BASE_DIR / "reports" / "listing"
RAW_DIR = REPORT_DIR / "raw"
DEBUG_DIR = REPORT_DIR / "debug"
DETAIL_DIR = REPORT_DIR / "detail"          # ASIN별 상세 페이지 결과(json, 날짜 포함)
SNAP_DIR = REPORT_DIR / "snapshots"         # 날마다의 스냅숏(변경 기록용)
INV_DIR = BASE_DIR / "reports" / "inventory"

TITLE_MAX = 75
HIGHLIGHT_MAX = 125
BULLET_MAX = 255
SEARCH_TERM_BYTES = 249
MIN_IMAGES = 7
MAX_DETAIL = 20            # 한 번에 최대 몇 개 상품 상세 페이지를 볼지
DETAIL_DAYS = 3            # 며칠 안에 본 상세 페이지는 다시 보지 않음
BAD_CHARS = set("!$?_{}^¬¦")
STOPWORDS = {"a", "an", "and", "or", "for", "the", "with", "in", "on", "of", "to", "by", "at", "from", "as", "&", "is"}
PROMO = ["free shipping", "best seller", "bestseller", "best-selling", "hot sale", "on sale", "big sale", "cheap",
         "100% quality", "quality guaranteed", "satisfaction guaranteed", "money back", "guaranteed", "#1",
         "top rated", "top-rated", "new arrival", "limited time", "discount", "lowest price", "best price",
         "buy now"]
CLAIMS = ["eco-friendly", "eco friendly", "non-toxic", "non toxic", "antibacterial", "anti-bacterial", "fda approved",
          "biodegradable", "cure", "cures", "miracle", "100% safe"]
EMOJI_RE = re.compile("[™®©†‡♥✓✔★☆❤➤►●❖"
                      "\U0001F300-\U0001FAFF☀-➿]")
ASIN_RE = re.compile(r"\bB0[A-Z0-9]{8}\b")
URL_RE = re.compile(r"https?://|www\.|\.com\b", re.I)

CATEGORY_LISTINGS = ["Category Listings Report", "카테고리 리스팅 보고서", "카테고리 목록 보고서", "分类商品报告", "类别商品报告"]
LEVEL_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}
LEVEL_NOTE = {"P0": "판매·노출에 바로 영향(오늘 처리)", "P1": "Amazon 정책 위반 가능 / 자동 수정 대상",
              "P2": "전환율·검색 개선", "P3": "참고"}
# 읽기만: 이 링크 / 버튼은 누르지 않음
UNSAFE = re.compile(r"edit|save|submit|approve|accept|reject|delete|publish|apply|update|upload|bulk|add-products|"
                    r"/abis/|listing/edit|productUpdate", re.I)
SC_PAGES = [  # (이름, 주소) — 먼저 이것을 시도, 페이지에서 찾은 관련 링크도 봄
    ("상품 관리", "/amazonsell/manage-products"),
    ("Fix your products", "/fixyourproducts"),
    ("재고 관리(전체)", "/myinventory/inventory"),
]
LINK_HINT = re.compile(r"suppress|fix your|fix product|listing quality|improve listing|review listing change|"
                       r"listing change|enhancement|recommendation|issue|search suppressed|검색 제한|수정 필요|개선", re.I)


# ================================================================ 작은 도구

def _dump(page, name: str) -> Path:
    """페이지 구조 + 스크린샷을 reports/listing/debug 에 저장."""
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w.-]+", "_", name)[:50]
    path = DEBUG_DIR / f"{datetime.now():%Y%m%d_%H%M%S}_{safe}_dom.txt"
    try:
        text = page.evaluate(inv._DUMP_JS)
    except Exception as e:
        text = f"(내보내기 실패: {e})"
    path.write_text(f"URL: {page.url}\n\n{text}", encoding="utf-8")
    sc_nav.save_debug(page, DEBUG_DIR, name)
    return path


def _brand() -> str:
    """브랜드 이름(설정 → 광고 설정의 브랜드 키워드). 없으면 빈 값 — 제목 브랜드 점검을 건너뜀."""
    import sp_api
    return sp_api.brand()


def _latest(folder: Path, pattern: str) -> Optional[Path]:
    files = sorted(folder.glob(pattern))
    return files[-1] if files else None


def _short(t: str, n: int = 60) -> str:
    t = str(t or "")
    return t if len(t) <= n else t[:n - 1] + "…"


def _words(t: str) -> List[str]:
    return re.findall(r"[a-z0-9]+(?:['-][a-z0-9]+)*", (t or "").lower())


# ================================================================ 규칙: 제목

def check_title(title: str, brand: str) -> List[Tuple[str, str, str, str]]:
    """돌려줌 [(등급, 분류, 설명, 제안)]."""
    out = []
    t = (title or "").strip()
    if not t:
        return [("P0", "제목", "제목이 비어 있음", "제목을 입력하세요")]
    n = len(t)
    if n > TITLE_MAX:
        out.append(("P1", "제목 75자 초과", f"{n}자({n - TITLE_MAX}자 초과). 잘리는 부분: “{_short(t[TITLE_MAX:], 50)}”",
                    "75자 이내로 줄이세요. 그대로 두면 Amazon AI가 제목을 자동으로 바꿉니다(브랜드 소유자는 14일 동안 검토 가능)"))
    bad = sorted({c for c in t if c in BAD_CHARS})
    if bad:
        out.append(("P1", "제목 금지 문자", "사용한 문자: " + " ".join(bad), "! $ ? _ { } ^ ¬ ¦ 문자를 빼세요"))
    cnt = Counter(w for w in _words(t) if w not in STOPWORDS and not w.isdigit())
    rep = [f"{w}×{c}" for w, c in cnt.items() if c > 2]
    if rep:
        out.append(("P1", "제목 단어 반복", "3번 이상: " + ", ".join(rep), "같은 단어(브랜드 포함)는 2번까지만 쓰세요"))
    low = t.lower()
    promo = [p for p in PROMO if re.search(rf"(?<![a-z]){re.escape(p)}(?![a-z])", low)]
    if promo:
        out.append(("P1", "제목 홍보 문구", ", ".join(promo), "홍보·가격 관련 문구를 빼세요"))
    if brand and not low.startswith(brand.lower()):
        out.append(("P3", "제목 브랜드", f"제목이 브랜드명({brand})으로 시작하지 않음", "다른 브랜드 상품이면 무시하세요"))
    caps = [w for w in re.findall(r"\b[A-Z]{5,}\b", t) if w.lower() != (brand or "").lower()]
    if caps:
        out.append(("P3", "제목 대문자", "대문자만 쓴 단어: " + ", ".join(caps[:5]), "브랜드 외에는 첫 글자만 대문자로 쓰세요"))
    return out


# ================================================================ 규칙: 불릿 / 검색어 / Item Highlights

def check_bullets(bullets: List[str], source: str) -> List[Tuple[str, str, str, str]]:
    out = []
    b = [x.strip() for x in bullets if str(x or "").strip()]
    if not b:
        return [("P1", "불릿 포인트", f"불릿 포인트가 없음({source})", "핵심 특징 5개를 작성하세요")]
    if len(b) < 5:
        out.append(("P2", "불릿 포인트", f"{len(b)}개만 있음({source})", "5개를 모두 작성하세요"))
    long_ = [i + 1 for i, x in enumerate(b) if len(x) > BULLET_MAX]
    if long_:
        out.append(("P2", "불릿 길이", f"{BULLET_MAX}자 초과: {', '.join(map(str, long_))}번", f"각 불릿은 {BULLET_MAX}자 이내 권장"))
    short_ = [i + 1 for i, x in enumerate(b) if len(x) < 10]
    if short_:
        out.append(("P2", "불릿 길이", f"너무 짧음: {', '.join(map(str, short_))}번", "10자 이상으로 구체적으로 쓰세요"))
    text = " ".join(b)
    low = text.lower()
    em = sorted(set(EMOJI_RE.findall(text)))
    if em:
        out.append(("P2", "불릿 특수문자", "사용한 문자: " + " ".join(em), "이모지·™·®·체크 표시 등은 빼세요"))
    promo = [p for p in PROMO if re.search(rf"(?<![a-z]){re.escape(p)}(?![a-z])", low)]
    if promo:
        out.append(("P1", "불릿 홍보 문구", ", ".join(promo), "홍보·가격·보증 문구를 빼세요"))
    claims = [p for p in CLAIMS if re.search(rf"(?<![a-z]){re.escape(p)}(?![a-z])", low)]
    if claims:
        out.append(("P1", "불릿 근거 필요 주장", ", ".join(claims), "인증·근거가 없으면 빼세요(노출 제한 원인)"))
    if URL_RE.search(text) or ASIN_RE.search(text):
        out.append(("P1", "불릿 링크·ASIN", "URL 또는 ASIN이 들어 있음", "링크·ASIN·연락처는 빼세요"))
    if re.search(r"\bN/?A\b|\bTBD\b", text):
        out.append(("P2", "불릿 빈 값", "N/A·TBD 같은 빈 값이 있음", "실제 내용을 쓰세요"))
    return out


def check_search_terms(st: str, title: str, brand: str) -> List[Tuple[str, str, str, str]]:
    st = (st or "").strip()
    if not st:
        return [("P2", "검색어(백엔드)", "후면 검색어가 비어 있음", "제목·불릿에 없는 관련 단어를 249바이트 이내로 넣으세요")]
    out = []
    nb = len(st.encode("utf-8"))
    if nb > SEARCH_TERM_BYTES:
        out.append(("P1", "검색어(백엔드)", f"{nb}바이트({SEARCH_TERM_BYTES}바이트 초과)", "249바이트를 넘으면 검색어가 반영되지 않을 수 있습니다"))
    w = _words(st)
    if brand and brand.lower() in w:
        out.append(("P3", "검색어(백엔드)", "자기 브랜드명이 들어 있음", "브랜드명은 빼고 다른 단어를 넣으세요"))
    if ASIN_RE.search(st.upper()):
        out.append(("P2", "검색어(백엔드)", "ASIN이 들어 있음", "ASIN은 빼세요"))
    dup = [x for x, c in Counter(w).items() if c > 1 and x not in STOPWORDS]
    in_title = sorted(set(w) & set(_words(title)) - STOPWORDS)
    if dup or len(in_title) >= 3:
        d = []
        if dup:
            d.append("반복: " + ", ".join(dup[:6]))
        if in_title:
            d.append("제목에 이미 있음: " + ", ".join(in_title[:6]))
        out.append(("P3", "검색어(백엔드)", " / ".join(d), "중복 단어를 빼고 새 단어로 채우세요"))
    return out


def check_highlights(hl: Optional[str]) -> List[Tuple[str, str, str, str]]:
    if hl is None:
        return []
    hl = hl.strip()
    if not hl:
        return [("P2", "Item Highlights", "작성되지 않음", "재질·용도·비교 단어를 125자 이내로 작성하세요(검색 결과에도 보임)")]
    if len(hl) > HIGHLIGHT_MAX:
        return [("P1", "Item Highlights", f"{len(hl)}자({HIGHLIGHT_MAX}자 초과)", "125자 이내로 줄이세요")]
    return []


# ================================================================ 규칙: 상세 페이지

def check_detail(d: Dict, our_title: str) -> List[Tuple[str, str, str, str]]:
    out = []
    if d.get("error"):
        return [("P1", "상세 페이지", d["error"], "상품 페이지가 열리는지 직접 확인하세요")]
    shown = (d.get("title") or "").strip()
    if shown and our_title and _norm_title(shown) != _norm_title(our_title):
        out.append(("P2", "표시 제목 다름", f"상세 페이지: “{_short(shown, 80)}”",
                    "등록한 제목과 다릅니다. Amazon이 바꿨거나 다른 판매자가 수정했을 수 있습니다(Review Listing Changes 확인)"))
    n = d.get("images") or 0
    if n < MIN_IMAGES:
        out.append(("P2", "이미지 수", f"{n}장", f"{MIN_IMAGES}장 이상(사용 장면·크기·구성품) 권장"))
    if not d.get("videos"):
        out.append(("P3", "동영상", "동영상 없음", "사용 영상을 추가하면 전환율에 도움이 됩니다"))
    wb = d.get("main_white")
    if wb is not None and wb < 0.9:
        out.append(("P1", "메인 이미지 배경", f"가장자리 흰색 비율 {wb:.0%}", "메인 이미지는 순백(RGB 255) 배경이어야 합니다. 직접 확인하세요"))
    side = d.get("main_px") or 0
    if side and side < 1000:
        out.append(("P2", "메인 이미지 크기", f"긴 변 {side}px", "1000px 이상이어야 확대(줌)가 됩니다"))
    if not d.get("aplus"):
        out.append(("P2", "A+ 콘텐츠", "A+ 콘텐츠 없음", "A+ 콘텐츠를 추가하세요(브랜드 등록 상품)"))
    if not d.get("brand_story"):
        out.append(("P3", "브랜드 스토리", "브랜드 스토리 없음", "A+ 관리자에서 브랜드 스토리를 추가할 수 있습니다"))
    r = d.get("rating")
    if r is not None and r < 4.0:
        out.append(("P1", "평점", f"{r:.1f}점(리뷰 {d.get('reviews') or 0}개)", "낮은 평점 리뷰 내용을 확인하고 상품·설명을 개선하세요"))
    if not d.get("buy_button"):
        out.append(("P1", "구매 버튼", "장바구니 버튼 없음: " + _short(d.get("availability") or "", 60),
                    "재고·바이박스·가격을 확인하세요"))
    return out


def _norm_title(t: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (t or "").lower())


# ================================================================ 자료 1: All Listings

def load_all_listings(page=None) -> Tuple[List[Dict], str]:
    """오늘 All Listings가 있으면 사용, 없고 브라우저가 있으면 다시 받음, 둘 다 안 되면 가장 최근 파일."""
    today = date.today().isoformat()
    f = INV_DIR / f"all_listings_{today}.xlsx"
    if not f.exists():
        import sp_api
        import sp_data
        raw = sp_api.try_api("전체 상품(All Listings)", sp_data.all_listings, today)
        if raw:
            INV_DIR.mkdir(parents=True, exist_ok=True)
            h, rows = inv.read_report(raw)
            inv.save_report_xlsx(h, rows, f, "All Listings")
            print(f"   ✅ All Listings(API) {len(rows)}행")
    if not f.exists() and page is not None:
        try:
            INV_DIR.mkdir(parents=True, exist_ok=True)
            raw = inv.fetch_all_listings(page, today)
            h, rows = inv.read_report(raw)
            inv.save_report_xlsx(h, rows, f, "All Listings")
            print(f"   ✅ All Listings Report {len(rows)}행")
        except Exception as e:
            print(f"   ⚠ All Listings Report를 가져오지 못했습니다: {str(e).splitlines()[0][:120]}")
    if not f.exists():
        f = _latest(INV_DIR, "all_listings_*.xlsx")
    if not f:
        return [], "없음"
    return inv.read_xlsx(f)[1], f.name


def _prev_all_listings(cur_name: str) -> Optional[Path]:
    files = [p for p in sorted(INV_DIR.glob("all_listings_*.xlsx")) if p.name < cur_name]
    return files[-1] if files else None


# ================================================================ 자료 2: Category Listings Report

def fetch_category_listings(page) -> Optional[Path]:
    today = date.today().isoformat()
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    have = sorted(RAW_DIR.glob(f"category_listings_{today}.*"))
    if have:
        print(f"   오늘의 Category Listings Report가 이미 있습니다({have[-1].name}).")
        return have[-1]
    print("   상품 → 상품 보고서 여는 중...")
    human.goto(page, sc_nav.SC_BASE + "/listing/reports", wait_after=False)
    ac.settle(page)
    human.read_page(page, max_screens=2)
    names = CATEGORY_LISTINGS + ["Category Listings"]
    ready = sc_reports.download_ready_today(page, RAW_DIR, f"category_listings_{today}", names, page_tz="local")
    if ready:
        print("   오늘 생성된 Category Listings Report를 바로 다운로드합니다.")
        return ready
    if sc_reports.pending_today(page, names, page_tz="local"):
        print("   오늘 이미 요청해 생성 중입니다. 다시 요청하지 않고 기다립니다.")
        old = sc_reports.existing_rows(page)
    else:
        print("   Category Listings Report 선택 후 보고서 요청...")
        inv._select_report_type(page, CATEGORY_LISTINGS)
        old = sc_reports.existing_rows(page)
        btn = None
        for _ in range(10):
            btn = inv._request_button(page)
            if btn is not None:
                break
            page.wait_for_timeout(500)
        if btn is None:
            path = _dump(page, "category_listings_request")
            raise RuntimeError(f"‘Request Report’ 버튼을 찾지 못했습니다(페이지 구조 저장: {path.name})")
        human.click(page, btn)
    print("   보고서 생성 대기 중(최대 30분)...")
    return sc_reports.wait_and_download(page, old, RAW_DIR, f"category_listings_{today}",
                                        row_must_contain=names, timeout_min=30, debug_dir=DEBUG_DIR)


def _read_text_table(path: Path) -> List[List[str]]:
    import csv
    import io
    data = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    first = text.split("\n", 1)[0]
    delim = "\t" if "\t" in first or "\t" in text[:5000] else ","
    return [r for r in csv.reader(io.StringIO(text), delimiter=delim)]


def parse_category_listings(path: Path) -> Tuple[Dict[str, Dict], List[str]]:
    """평면 파일: 앞 몇 줄은 템플릿 설명, 실제 열 이름 줄에 item_sku / item_name(또는 sku / item_name#1.value 같은 새 형식)이 있음.
    돌려줌 {sku: {title, bullets[], search_terms, highlights(None=보고서에 이 열이 없음), main_image, images, brand}}"""
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        rows = _read_text_table(path)
    hi = None
    for i, r in enumerate(rows[:10]):
        low = [c.strip().lower() for c in r]
        if any(re.match(r"^(item_sku|sku|contribution_sku|seller[-_ ]sku)", c) for c in low) and \
                any(c.startswith("item_name") or c in ("title", "product name") for c in low):
            hi = i
    if hi is None:
        return {}, []
    head = [c.strip() for c in rows[hi]]
    low = [c.lower() for c in head]

    def cols(rx):
        return [i for i, c in enumerate(low) if re.search(rx, c)]

    c_sku = cols(r"^(item_sku|sku|contribution_sku|seller[-_ ]sku)")[:1]
    c_title = cols(r"^item_name|^title$|^product name$")[:1]
    c_bul = cols(r"^bullet_point")
    c_kw = cols(r"^generic_keyword")
    c_hl = cols(r"highlight")
    c_main = cols(r"^main_image_url|^main_product_image_locator")[:1]
    c_other = cols(r"^other_image_url|^other_product_image_locator")
    c_brand = cols(r"^brand_name|^brand(#|$)")[:1]
    out: Dict[str, Dict] = {}
    for r in rows[hi + 1:]:
        if not r or not c_sku:
            continue
        g = lambda i: (r[i].strip() if i < len(r) else "")
        sku = g(c_sku[0])
        if not sku or sku.lower() in ("item_sku", "sku") or sku.startswith("#"):
            continue
        out[sku] = {
            "title": g(c_title[0]) if c_title else "",
            "bullets": [g(i) for i in c_bul if g(i)],
            "search_terms": " ".join(g(i) for i in c_kw if g(i)),
            "highlights": (" ".join(g(i) for i in c_hl if g(i)) if c_hl else None),
            "main_image": g(c_main[0]) if c_main else "",
            "images": (1 if c_main and g(c_main[0]) else 0) + sum(1 for i in c_other if g(i)),
            "brand": g(c_brand[0]) if c_brand else "",
        }
    return out, head


# ================================================================ 자료 3: 상품 상세 페이지

_DETAIL_JS = r"""
() => {
  const q = s => document.querySelector(s);
  const t = e => e ? (e.innerText || e.textContent || '').replace(/\s+/g, ' ').trim() : '';
  const body = (document.body && document.body.innerText) || '';
  const out = {url: location.href};
  out.captcha = !!q('form[action*="validateCaptcha"]') || /Enter the characters you see|Type the characters you see/i.test(body.slice(0, 3000));
  out.not_found = /Page Not Found|Sorry! We couldn.t find that page|looking for something\?/i.test(document.title + ' ' + body.slice(0, 1500));
  out.title = t(q('#productTitle'));
  const seen = new Set();
  out.bullets = Array.from(document.querySelectorAll('#feature-bullets ul li span.a-list-item, #feature-bullets ul li'))
      .map(t).filter(x => x && !/^›?\s*See more product details/i.test(x) && !seen.has(x) && seen.add(x));
  // 이미지 수: 썸네일, 또는 페이지 스크립트의 colorImages
  let thumbs = document.querySelectorAll('#altImages li.imageThumbnail').length;
  let scr = 0;
  for (const s of document.querySelectorAll('script')) {
    const x = s.textContent || '';
    const i = x.indexOf("'colorImages'");
    if (i >= 0) {
      const seg = x.slice(i, x.indexOf("'colorToAsin'", i) > 0 ? x.indexOf("'colorToAsin'", i) : i + 60000);
      const init = seg.indexOf("'initial'");
      const part = init >= 0 ? seg.slice(init) : seg;
      const end = part.indexOf(']}');
      scr = ((end > 0 ? part.slice(0, end) : part).match(/"hiRes"|"large"/g) || []).length;
      const hi = ((end > 0 ? part.slice(0, end) : part).match(/"hiRes"/g) || []).length;
      scr = hi || scr;
      break;
    }
  }
  out.images = Math.max(thumbs, scr);
  const vtxt = t(q('#altImages .video-count')) + ' ' + t(q('#altImages li.videoThumbnail')) + ' ' + t(q('#videoCount'));
  const vm = vtxt.match(/(\d+)\s*VIDEO/i);
  out.videos = vm ? parseInt(vm[1]) : (document.querySelectorAll('#altImages li.videoThumbnail, #altImages li.videoBlockIngress, #altImages .videoCountTemplate').length);
  const img = q('#landingImage') || q('#imgBlkFront');
  out.main_image = '';
  if (img) {
    out.main_image = img.getAttribute('data-old-hires') || '';
    if (!out.main_image) {
      try {
        const dyn = JSON.parse(img.getAttribute('data-a-dynamic-image') || '{}');
        let best = '', w = 0;
        for (const [u, s] of Object.entries(dyn)) if (s[0] > w) { w = s[0]; best = u; }
        out.main_image = best;
      } catch (e) {}
    }
    if (!out.main_image) out.main_image = img.getAttribute('src') || '';
  }
  const ap = q('#aplus') || q('#aplus_feature_div');
  out.aplus = !!(ap && t(ap).length > 60);
  const bs = q('#aplusBrandStory_feature_div');
  out.brand_story = !!(bs && t(bs).length > 30);
  const rt = q('#acrPopover');
  const rm = rt ? ((rt.getAttribute('title') || t(rt)).match(/([\d.]+)\s*out of/)) : null;
  out.rating = rm ? parseFloat(rm[1]) : null;
  const rc = t(q('#acrCustomerReviewText')).replace(/[^\d]/g, '');
  out.reviews = rc ? parseInt(rc) : 0;
  out.price = t(q('#corePrice_feature_div .a-price .a-offscreen')) || t(q('#corePriceDisplay_desktop_feature_div .a-price .a-offscreen')) || t(q('.a-price .a-offscreen'));
  out.availability = t(q('#availability'));
  out.buy_button = !!(q('#add-to-cart-button') || q('#buy-now-button'));
  out.seller = t(q('#sellerProfileTriggerId')) || t(q('#merchantInfoFeature_feature_div .offer-display-feature-text'));
  const det = t(q('#productDetails_detailBullets_sections1')) + ' ' + t(q('#detailBulletsWrapper_feature_div')) + ' ' + t(q('#prodDetails'));
  const bm = det.match(/Best Sellers Rank:?\s*#?([\d,]+)\s+in\s+([^(#]+)/i);
  out.bsr = bm ? ('#' + bm[1] + ' ' + bm[2].trim()) : '';
  out.brand = t(q('#bylineInfo'));
  out.variations = document.querySelectorAll('#twister li, #inline-twister-row-color_name li, [id^="inline-twister-row"] li').length;
  return out;
}
"""

_WHITE_JS = r"""
async (src) => {
  const img = new Image();
  await new Promise((ok, bad) => { img.onload = ok; img.onerror = bad; img.src = src; });
  const w = img.naturalWidth, h = img.naturalHeight;
  const c = document.createElement('canvas');
  const k = Math.min(1, 600 / Math.max(w, h));
  c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(h * k));
  const g = c.getContext('2d');
  g.drawImage(img, 0, 0, c.width, c.height);
  const d = g.getImageData(0, 0, c.width, c.height).data;
  const band = Math.max(2, Math.round(Math.min(c.width, c.height) * 0.02));
  let white = 0, all = 0;
  for (let y = 0; y < c.height; y++) {
    for (let x = 0; x < c.width; x++) {
      if (x >= band && x < c.width - band && y >= band && y < c.height - band) continue;
      const i = (y * c.width + x) * 4;
      all++;
      if (d[i] >= 245 && d[i + 1] >= 245 && d[i + 2] >= 245) white++;
    }
  }
  return {w: w, h: h, white: all ? white / all : null};
}
"""


def _main_image_check(page, url: str) -> Dict:
    """메인 이미지를 받아(브라우저 네트워크 사용) 페이지에서 canvas로 가장자리 색을 읽음: 흰 배경인지, 크기는 얼마인지."""
    if not url or url.startswith("data:"):
        return {}
    try:
        r = page.request.get(url, timeout=20000)
        if not r.ok:
            return {}
        body = r.body()
        ctype = (r.headers.get("content-type") or "image/jpeg").split(";")[0]
        src = f"data:{ctype};base64," + base64.b64encode(body).decode()
        res = page.evaluate(_WHITE_JS, src)
        return {"main_white": res.get("white"), "main_px": max(res.get("w") or 0, res.get("h") or 0)}
    except Exception:
        return {}


def _detail_cached(asin: str) -> Optional[Dict]:
    p = DETAIL_DIR / f"{asin}.json"
    if not p.exists():
        return None
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
        return d
    except Exception:
        return None


def fetch_details(page, asins: List[str], limit: int = MAX_DETAIL) -> Tuple[Dict[str, Dict], str]:
    """상품 상세 페이지를 차례로 봄(보기만, 버튼은 누르지 않음). 3일 안에 본 것은 저장된 결과 사용."""
    DETAIL_DIR.mkdir(parents=True, exist_ok=True)
    res: Dict[str, Dict] = {}
    todo = []
    cutoff = (date.today() - timedelta(days=DETAIL_DAYS - 1)).isoformat()
    for a in asins:
        c = _detail_cached(a)
        if c and c.get("date", "") >= cutoff:
            res[a] = c
        else:
            todo.append(a)
            if c:
                res[a] = c          # 먼저 예전 것을 넣고 새로 보면 덮어씀
    todo = todo[:limit]
    cached = len([a for a in asins if a in res and a not in todo])
    note = f"상세 페이지 {len(todo)}개 확인, 최근 {DETAIL_DAYS}일 안에 확인한 {cached}개는 저장된 결과 사용"
    if not todo:
        return res, note
    print(f"   상품 상세 페이지 {len(todo)}개 확인(광고비 많은 순, 한 번에 최대 {limit}개)...")
    for i, a in enumerate(todo, 1):
        try:
            human.goto(page, f"https://www.amazon.com/dp/{a}?th=1&psc=1", wait_after=False)
            try:
                page.wait_for_load_state("load", timeout=20000)
            except Exception:
                pass
            human.read_page(page, max_screens=3)
            d = page.evaluate(_DETAIL_JS)
            if d.get("captcha"):
                print("   ⚠ Amazon이 문자 확인(CAPTCHA)을 요구합니다. 상세 페이지 확인을 여기서 멈춥니다.")
                _dump(page, "listing_detail_captcha")
                note += " · CAPTCHA로 중간에 멈춤"
                break
            if d.get("not_found"):
                d = {"error": "상품 페이지를 찾을 수 없음(Page Not Found)"}
            else:
                d.update(_main_image_check(page, d.get("main_image")))
            d["date"] = date.today().isoformat()
            d["checked"] = datetime.now().isoformat(timespec="seconds")
            (DETAIL_DIR / f"{a}.json").write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
            res[a] = d
            print(f"     {i}/{len(todo)} {a}: 이미지 {d.get('images', 0)}장, 불릿 {len(d.get('bullets') or [])}개, "
                  f"A+ {'있음' if d.get('aplus') else '없음'}, 평점 {d.get('rating') or '—'}")
        except Exception as e:
            print(f"     {i}/{len(todo)} {a}: 확인 실패 {type(e).__name__}")
            if i == 1:
                _dump(page, f"listing_detail_failed_{a}")
    return res, note


# ================================================================ 자료 4: Seller Central 상품 관리 페이지(조회만)

_LINKS_JS = r"""
() => {
  const out = [];
  const seen = new Set();
  for (const a of document.querySelectorAll('a[href], kat-link[href]')) {
    const t = (a.innerText || a.textContent || a.getAttribute('label') || '').replace(/\s+/g, ' ').trim();
    const h = a.href || a.getAttribute('href') || '';
    if (!t || seen.has(t + h)) continue;
    seen.add(t + h);
    out.push({text: t.slice(0, 120), href: h});
  }
  return out;
}
"""

_COUNTS_JS = r"""
() => {
  const body = (document.body && document.body.innerText) || '';
  const out = [];
  const rx = /(search suppressed|suppressed|fix your products|listing issues?|needs attention|incomplete|inactive|review listing changes|enhancements?|recommendations?)[^\n]{0,40}/ig;
  let m; const seen = new Set();
  while ((m = rx.exec(body)) && out.length < 40) {
    const s = m[0].trim();
    if (!seen.has(s)) { seen.add(s); out.push(s); }
  }
  return out;
}
"""


def explore_sc(page) -> List[Dict]:
    """상품 관리 관련 페이지 보기: 페이지 구조 저장, ‘검색 제한 / 수정 필요 / 제안 변경’ 글자와 링크 기록. 버튼은 누르지 않음."""
    rows: List[Dict] = []
    visited = set()
    queue = [(n, sc_nav.SC_BASE + u) for n, u in SC_PAGES]
    while queue and len(visited) < 7:
        name, url = queue.pop(0)
        key = url.split("#")[0]
        if key in visited or UNSAFE.search(url):
            continue
        visited.add(key)
        try:
            human.goto(page, url, wait_after=False)
            ac.settle(page, 15000)
            human.read_page(page, max_screens=3)
            path = _dump(page, f"listing_{name}")
            counts = page.evaluate(_COUNTS_JS)
            links = [l for l in page.evaluate(_LINKS_JS) if LINK_HINT.search(l["text"]) or LINK_HINT.search(l["href"])]
            rows.append({"페이지": name, "주소": page.url, "찾은 문구": " | ".join(counts[:12]),
                         "관련 링크": " | ".join(f"{l['text']} → {l['href']}" for l in links[:10]), "저장": path.name})
            print(f"   {name}: 관련 문구 {len(counts)}개, 관련 링크 {len(links)}개")
            for l in links:
                h = l["href"]
                if h.startswith(sc_nav.SC_BASE) and not UNSAFE.search(h) and h.split("#")[0] not in visited:
                    queue.append((_short(l["text"], 30), h))
        except Exception as e:
            rows.append({"페이지": name, "주소": url, "찾은 문구": f"열기 실패: {type(e).__name__}", "관련 링크": "", "저장": ""})
    return rows


# ================================================================ 광고비(정렬용)

def _ad_spend() -> Dict[str, float]:
    try:
        import ads_analyze as aa
        rows, _ = aa._from_days("rp_sp_advertised_product.xlsx", ac.pacific_today(), 5)
        sp: Dict[str, float] = defaultdict(float)
        for r in ac.normalize(rows):
            a = str(r.get("asin") or "").strip().upper()
            if a:
                sp[a] += r.get("spend", 0) or 0
        return dict(sp)
    except Exception:
        return {}


def _fba_available() -> Dict[str, float]:
    f = _latest(INV_DIR, "fba_inventory_*.xlsx")
    if not f:
        return {}
    out: Dict[str, float] = {}
    for it in inv.build_rows([], inv.read_xlsx(f)[1]):
        out[it["sku"]] = it.get("available") or 0
    return out


# ================================================================ 분석

def analyze(al: List[Dict], cat: Dict[str, Dict], details: Dict[str, Dict], brand: str) -> Dict:
    spend = _ad_spend()
    fba = _fba_available()
    issues: List[Dict] = []
    titles: List[Dict] = []
    content: List[Dict] = []
    per_asin: Dict[str, Dict] = {}

    def add(level, typ, sku, asin, title, detail, advice):
        issues.append({"level": level, "type": typ, "sku": sku, "asin": asin, "title": title, "detail": detail,
                       "advice": advice, "spend": spend.get(asin, 0)})

    for r in al:
        sku = str(inv._get(r, "seller-sku", "sku", default="")).strip()
        if not sku:
            continue
        asin = str(inv._get(r, "asin1", "asin", default="")).strip().upper()
        title = str(inv._get(r, "item-name", "product-name", default="") or "")
        status = str(inv._get(r, "status", default="") or "")
        st = status.lower()
        c = cat.get(sku) or {}
        a = per_asin.setdefault(asin, {"asin": asin, "skus": [], "status": status, "title": title, "spend": spend.get(asin, 0)})
        a["skus"].append(sku)
        # 상태
        if st == "incomplete":
            add("P0", "정보 미완성", sku, asin, title, "Incomplete: 필수 정보가 빠져 판매할 수 없음",
                "상품 편집 화면에서 빨간 표시 항목(필수 속성)을 채우세요")
        elif st == "inactive":
            avail = fba.get(sku, 0)
            if avail > 0:
                add("P0", "재고 있는데 판매 중지", sku, asin, title, f"Inactive인데 FBA 판매 가능 재고 {avail:.0f}개",
                    "가격·노출 제한·정책 문제를 확인하세요(상품 관리 → 문제 해결)")
            else:
                add("P3", "판매 중지", sku, asin, title, "Inactive(재고 없음)", "재입고 계획이 없으면 무시하세요")
        # 제목(판매 중·정보 미완성만 점검, Inactive는 조금만)
        tf = check_title(title, brand)
        if st == "inactive":
            tf = [x for x in tf if x[1] == "제목 75자 초과"]
            tf = [("P3",) + x[1:] for x in tf]
        for lv, typ, det, adv in tf:
            add(lv, typ, sku, asin, title, det, adv)
        titles.append({"SKU": sku, "ASIN": asin, "상태": status, "광고비(최근)": round(spend.get(asin, 0), 2),
                       "글자 수": len(title.strip()), "75자 초과": max(0, len(title.strip()) - TITLE_MAX),
                       "75자까지": title.strip()[:TITLE_MAX], "잘리는 부분": title.strip()[TITLE_MAX:],
                       "문제": " / ".join(f"[{x[0]}] {x[1]}: {x[2]}" for x in tf if x[1] != "제목 75자 초과") or "",
                       "현재 제목": title})
        if st == "inactive":
            continue
        # 불릿 / 검색어 / Item Highlights(보고서) + 상세 페이지
        d = details.get(asin)
        bullets = c.get("bullets") if c else None
        src = "카테고리 보고서"
        if not bullets and d and not d.get("error"):
            bullets, src = d.get("bullets") or [], "상세 페이지"
        if bullets is not None and (c or d):
            for lv, typ, det, adv in check_bullets(bullets, src):
                add(lv, typ, sku, asin, title, det, adv)
        if c:
            for lv, typ, det, adv in check_search_terms(c.get("search_terms", ""), title, brand):
                add(lv, typ, sku, asin, title, det, adv)
            for lv, typ, det, adv in check_highlights(c.get("highlights")):
                add(lv, typ, sku, asin, title, det, adv)
            content.append({"SKU": sku, "ASIN": asin, "불릿 수": len(c.get("bullets") or []),
                            "가장 긴 불릿(자)": max([len(x) for x in c.get("bullets") or []] or [0]),
                            "검색어(바이트)": len((c.get("search_terms") or "").encode("utf-8")),
                            "Item Highlights": ("(보고서에 없음)" if c.get("highlights") is None else (c.get("highlights") or "(비어 있음)")),
                            "이미지(보고서)": c.get("images", 0), "검색어": c.get("search_terms", ""),
                            "불릿 1": (c.get("bullets") or [""])[0] if c.get("bullets") else ""})
    # 상세 페이지(ASIN마다 한 번)
    detail_rows = []
    for asin, a in per_asin.items():
        d = details.get(asin)
        if not d:
            continue
        sku = a["skus"][0]
        for lv, typ, det, adv in check_detail(d, a["title"]):
            if typ == "불릿 포인트":
                continue
            add(lv, typ, sku, asin, a["title"], det, adv)
        detail_rows.append({"ASIN": asin, "SKU": ", ".join(a["skus"][:3]), "확인일": d.get("date", ""),
                            "광고비(최근)": round(a["spend"], 2), "이미지": d.get("images", 0), "동영상": d.get("videos", 0),
                            "메인 흰 배경": ("" if d.get("main_white") is None else f"{d['main_white']:.0%}"),
                            "메인 크기(px)": d.get("main_px", ""), "A+": "있음" if d.get("aplus") else "없음",
                            "브랜드 스토리": "있음" if d.get("brand_story") else "없음", "평점": d.get("rating"),
                            "리뷰 수": d.get("reviews"), "가격": d.get("price", ""), "구매 버튼": "있음" if d.get("buy_button") else "없음",
                            "판매자": d.get("seller", ""), "BSR": d.get("bsr", ""), "불릿 수(페이지)": len(d.get("bullets") or []),
                            "표시 제목": d.get("title", "") or d.get("error", "")})
    # ASIN별 점수
    score_rows = []
    by_asin = defaultdict(Counter)
    for it in issues:
        by_asin[it["asin"]][it["level"]] += 1
    for asin, a in per_asin.items():
        c = by_asin.get(asin, Counter())
        sc = max(0, 100 - 30 * c["P0"] - 10 * c["P1"] - 4 * c["P2"] - 1 * c["P3"])
        d = details.get(asin) or {}
        score_rows.append({"ASIN": asin, "SKU": ", ".join(a["skus"][:3]) + (" …" if len(a["skus"]) > 3 else ""),
                           "상태": a["status"], "점수": sc, "P0": c["P0"], "P1": c["P1"], "P2": c["P2"], "P3": c["P3"],
                           "광고비(최근)": round(a["spend"], 2), "제목 글자 수": len(a["title"].strip()),
                           "이미지": d.get("images", ""), "A+": ("" if not d else ("있음" if d.get("aplus") else "없음")),
                           "평점": d.get("rating", ""), "제목": a["title"]})
    score_rows.sort(key=lambda x: (x["상태"].lower() == "inactive", x["점수"], -x["광고비(최근)"]))
    issues.sort(key=lambda x: (LEVEL_ORDER[x["level"]], -x["spend"], x["type"], x["sku"]))
    titles.sort(key=lambda x: (-x["광고비(최근)"], -x["75자 초과"]))
    return {"issues": issues, "titles": titles, "content": content, "details": detail_rows, "scores": score_rows,
            "per_asin": per_asin}


# ================================================================ 스냅숏과 변경

def snapshot(al: List[Dict], cat: Dict[str, Dict], details: Dict[str, Dict]) -> Dict:
    snap = {}
    for r in al:
        sku = str(inv._get(r, "seller-sku", "sku", default="")).strip()
        if not sku:
            continue
        asin = str(inv._get(r, "asin1", "asin", default="")).strip().upper()
        c = cat.get(sku) or {}
        d = details.get(asin) or {}
        snap[sku] = {"asin": asin, "title": str(inv._get(r, "item-name", default="") or ""),
                     "status": str(inv._get(r, "status", default="") or ""),
                     "price": str(inv._get(r, "price", default="") or ""),
                     "bullets": c.get("bullets") or [], "search_terms": c.get("search_terms", ""),
                     "shown_title": d.get("title", ""), "images": d.get("images"), "detail_date": d.get("date", "")}
    return snap


def changes(prev: Dict, cur: Dict) -> List[Dict]:
    out = []
    labels = {"title": "등록 제목", "status": "상태", "price": "가격", "bullets": "불릿 포인트", "search_terms": "검색어",
              "shown_title": "상세 페이지 표시 제목", "images": "이미지 수"}
    for sku, c in cur.items():
        p = prev.get(sku)
        if not p:
            continue
        for k, lab in labels.items():
            a, b = p.get(k), c.get(k)
            if k in ("shown_title", "images") and (not p.get("detail_date") or p.get("detail_date") == c.get("detail_date")):
                continue
            if a in (None, "", []) or b in (None, "", []):
                continue
            if a != b:
                fa = " / ".join(a) if isinstance(a, list) else str(a)
                fb = " / ".join(b) if isinstance(b, list) else str(b)
                out.append({"SKU": sku, "ASIN": c.get("asin", ""), "항목": lab, "이전": _short(fa, 200), "지금": _short(fb, 200)})
    return out


def _prev_snapshot(today: str) -> Tuple[Dict, str]:
    files = [p for p in sorted(SNAP_DIR.glob("*.json")) if p.stem < today]
    if not files:
        return {}, ""
    try:
        return json.loads(files[-1].read_text(encoding="utf-8")), files[-1].stem
    except Exception:
        return {}, ""


# ================================================================ Excel

def write_report(out: Path, today: str, res: Dict, sources: Dict[str, str], explore_rows: List[Dict],
                 chg: List[Dict], prev_day: str, notes: List[str]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    FONT = inv.FONT
    RED, ORANGE, GREY = "C01C1C", "C26A00", "6B6B6B"
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 90
    r = 1
    ws.cell(r, 1, f"Listing 점검 보고서 · {today}").font = Font(name=FONT, size=15, bold=True, color="1F3864")
    r += 1
    ws.cell(r, 1, "1단계: 자료 확인과 수정 제안만 합니다. 상품 정보는 바꾸지 않습니다.").font = Font(name=FONT, size=10, color=GREY)
    r += 2
    lv = Counter(x["level"] for x in res["issues"])
    ws.cell(r, 1, "조치 필요 건수").font = Font(name=FONT, bold=True)
    r += 1
    for k in ("P0", "P1", "P2", "P3"):
        ws.cell(r, 1, f"{k} · {LEVEL_NOTE[k]}").font = Font(name=FONT, color=RED if k == "P0" else (ORANGE if k == "P1" else None))
        ws.cell(r, 2, lv.get(k, 0)).font = Font(name=FONT, bold=True)
        r += 1
    r += 1
    ws.cell(r, 1, "분류별(상위)").font = Font(name=FONT, bold=True)
    r += 1
    for (lvl, typ), n in Counter((x["level"], x["type"]) for x in res["issues"]).most_common(18):
        ws.cell(r, 1, f"[{lvl}] {typ}").font = Font(name=FONT)
        ws.cell(r, 2, n).font = Font(name=FONT)
        r += 1
    r += 1
    t = res["titles"]
    act = [x for x in t if x["상태"].lower() != "inactive"]
    over = [x for x in act if x["75자 초과"] > 0]
    ws.cell(r, 1, "제목 75자 기준").font = Font(name=FONT, bold=True)
    ws.cell(r, 2, f"판매 중·미완성 {len(act)}개 중 {len(over)}개가 75자 초과 "
                  f"(2026-07-27부터 Amazon AI가 순차적으로 자동 수정. 브랜드 소유자는 적용 전 14일 동안 Review Listing Changes에서 검토 가능)").font = Font(name=FONT)
    ws.cell(r, 2).alignment = Alignment(wrap_text=True)
    r += 2
    ws.cell(r, 1, "자료").font = Font(name=FONT, bold=True)
    r += 1
    for k, v in sources.items():
        ws.cell(r, 1, k).font = Font(name=FONT)
        ws.cell(r, 2, v).font = Font(name=FONT)
        r += 1
    if prev_day:
        ws.cell(r, 1, "변경 비교").font = Font(name=FONT)
        ws.cell(r, 2, f"{prev_day} 대비 {len(chg)}건 변경(‘변경 기록’ 시트)").font = Font(name=FONT)
        r += 1
    if notes:
        r += 1
        ws.cell(r, 1, "참고").font = Font(name=FONT, bold=True)
        r += 1
        for n in notes:
            ws.cell(r, 2, n).font = Font(name=FONT, color=GREY)
            ws.cell(r, 2).alignment = Alignment(wrap_text=True)
            r += 1
    r += 1
    ws.cell(r, 1, "점검 기준").font = Font(name=FONT, bold=True)
    for line in ("제목: 75자 이내, 같은 단어 2번까지, ! $ ? _ { } ^ ¬ ¦ 금지, 홍보 문구 금지",
                 "불릿: 5개, 각 10~255자 권장, 이모지·™·®·링크·ASIN·홍보/근거 없는 주장 금지",
                 "검색어(백엔드): 249바이트 이내, 브랜드명·ASIN·중복 단어 제외 / Item Highlights: 125자 이내",
                 "상세 페이지: 이미지 7장 이상, 메인 이미지 흰 배경·1000px 이상, 동영상·A+·브랜드 스토리, 평점 4.0 이상, 구매 버튼",
                 "점수: 100 − P0×30 − P1×10 − P2×4 − P3×1"):
        r += 1
        ws.cell(r, 2, line).font = Font(name=FONT, size=10)

    def sheet(name, headers, rows, widths=None):
        w = wb.create_sheet(name)
        inv.write_table_sheet(w, headers, rows, widths)
        return w

    fills = {"P0": PatternFill("solid", fgColor="F8D7DA"), "P1": PatternFill("solid", fgColor="FFE8CC")}
    iss = [{"등급": x["level"], "분류": x["type"], "SKU": x["sku"], "ASIN": x["asin"], "광고비(최근)": round(x["spend"], 2),
            "내용": x["detail"], "조치": x["advice"], "제목": _short(x["title"], 80)} for x in res["issues"]]
    w = sheet("조치 필요", ["등급", "분류", "SKU", "ASIN", "광고비(최근)", "내용", "조치", "제목"], iss,
              {"등급": 6, "분류": 18, "SKU": 16, "ASIN": 13, "광고비(최근)": 11, "내용": 60, "조치": 55, "제목": 50})
    for i, x in enumerate(iss, start=2):
        f = fills.get(x["등급"])
        if f:
            w.cell(i, 1).fill = f
    sheet("ASIN 점수", ["ASIN", "SKU", "상태", "점수", "P0", "P1", "P2", "P3", "광고비(최근)", "제목 글자 수", "이미지", "A+", "평점", "제목"],
          res["scores"], {"제목": 70, "SKU": 24})
    sheet("제목 점검", ["SKU", "ASIN", "상태", "광고비(최근)", "글자 수", "75자 초과", "75자까지", "잘리는 부분", "문제", "현재 제목"],
          res["titles"], {"75자까지": 60, "잘리는 부분": 50, "문제": 50, "현재 제목": 80})
    if res["content"]:
        sheet("불릿·검색어", ["SKU", "ASIN", "불릿 수", "가장 긴 불릿(자)", "검색어(바이트)", "Item Highlights", "이미지(보고서)", "검색어", "불릿 1"],
              res["content"], {"Item Highlights": 40, "검색어": 60, "불릿 1": 60})
    if res["details"]:
        sheet("상세 페이지", ["ASIN", "SKU", "확인일", "광고비(최근)", "이미지", "동영상", "메인 흰 배경", "메인 크기(px)", "A+", "브랜드 스토리",
                         "평점", "리뷰 수", "가격", "구매 버튼", "판매자", "BSR", "불릿 수(페이지)", "표시 제목"],
              res["details"], {"표시 제목": 70, "BSR": 30})
    sheet("변경 기록", ["SKU", "ASIN", "항목", "이전", "지금"], chg, {"이전": 70, "지금": 70})
    if explore_rows:
        sheet("페이지 확인", ["페이지", "주소", "찾은 문구", "관련 링크", "저장"], explore_rows,
              {"주소": 50, "찾은 문구": 70, "관련 링크": 80, "저장": 40})
    wb.save(out)


# ================================================================ 실행

def run(open_report: bool = True) -> Optional[Path]:
    from sc_browser import open_session
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    SNAP_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    brand = _brand()
    sources: Dict[str, str] = {}
    notes: List[str] = []
    cat: Dict[str, Dict] = {}
    details: Dict[str, Dict] = {}
    explore_rows: List[Dict] = []
    api_issues: List[Tuple] = []
    topic_rows: List[Dict] = []
    print("   1단계: 자료 확인과 수정 제안만 합니다. 상품 정보는 바꾸지 않습니다.")

    with open_session(headless=False) as page:
        print("\n【1】 All Listings Report")
        al, al_name = load_all_listings(page)
        sources["전체 상품(All Listings)"] = f"{al_name}, {len(al)}행"
        if not al:
            raise RuntimeError("All Listings Report가 없습니다. 먼저 ‘전체 상품 재고’를 실행하세요")

        print("\n【2】 리스팅 상세(불릿·검색어·Item Highlights·Amazon 지적 사항)")
        import sp_api
        import sp_data
        items = sp_api.try_api("리스팅 상세(Listings Items API)",
                               lambda: sp_data.listings_items(sp_api.seller_id()))
        if items:
            cat = items
            api_issues = sp_data.issue_findings(items)
            hl = "있음" if any(v.get("highlights") is not None for v in cat.values()) else "항목 없음"
            sources["리스팅 상세(SP-API)"] = f"{len(cat)}개 SKU, Amazon 지적 사항 {len(api_issues)}건(Item Highlights: {hl})"
            print(f"   ✅ {len(cat)}개 SKU, Amazon 지적 사항 {len(api_issues)}건")
        try:
            if items:
                raise StopIteration
            raw = fetch_category_listings(page)
            cat, head = parse_category_listings(raw) if raw else ({}, [])
            if cat:
                hl = "있음" if any(v.get("highlights") is not None for v in cat.values()) else "보고서에 항목 없음"
                sources["카테고리 보고서"] = f"{raw.name}, {len(cat)}개 SKU(Item Highlights 열: {hl})"
                print(f"   ✅ {len(cat)}개 SKU 읽음")
            else:
                sources["카테고리 보고서"] = f"{raw.name if raw else '없음'}: 열 이름을 알아보지 못함"
                notes.append("카테고리 보고서의 열 구조를 알아보지 못했습니다. raw 폴더의 파일을 개발자에게 보내 주세요.")
        except StopIteration:
            pass
        except Exception as e:
            msg = str(e).splitlines()[0][:150]
            print(f"   ⚠ Category Listings Report 실패: {msg}")
            sources["카테고리 보고서"] = f"가져오지 못함({msg})"
            notes.append("카테고리 보고서가 없어 불릿은 상세 페이지 내용으로 점검했고, 검색어·Item Highlights는 점검하지 못했습니다.")

        print("\n【3】 상품 워크스페이스: 리스팅 상태(검색 제한·상세 페이지 삭제 등, 보기만 함)")
        status_page = None
        try:
            import sc_pages
            status_page = sc_pages.read(page, "listing_status")
            st = status_page.get("status") or {}
            sources["리스팅 상태(상품 워크스페이스)"] = (", ".join(f"{a} {b}" for a, b in st.items()) if st
                                                 else status_page.get("error") or status_page.get("note", "값 없음"))
        except Exception as e:
            sources["리스팅 상태(상품 워크스페이스)"] = f"확인 실패: {type(e).__name__}"

        print("\n【4】 상품 상세 페이지(amazon.com)")
        spend = _ad_spend()
        active = []
        for r in al:
            st = str(inv._get(r, "status", default="") or "").lower()
            a = str(inv._get(r, "asin1", "asin", default="")).strip().upper()
            if a and st == "active" and a not in active:
                active.append(a)
        active.sort(key=lambda a: -spend.get(a, 0))
        try:
            details, note = fetch_details(page, active)
            sources["상세 페이지"] = note + f" · 판매 중 ASIN {len(active)}개 중 결과 있는 것 {len(details)}개"
        except Exception as e:
            sources["상세 페이지"] = f"확인 실패: {type(e).__name__}"
        if len(details) < len(active):
            notes.append(f"상세 페이지는 하루 최대 {MAX_DETAIL}개씩 확인합니다. 남은 {len(active) - len(details)}개는 다음 실행 때 확인합니다.")

    # 광고비 상위 ASIN의 부정 리뷰 주제(Customer Feedback API)
    try:
        import sp_api
        import sp_data
        top = [a for a in active[:15]] if 'active' in dir() else []
        topic_rows = sp_api.try_api("리뷰 불만 주제(Customer Feedback API)", sp_data.negative_topics, top) or []
        if topic_rows:
            sources["리뷰 불만 주제(SP-API)"] = f"ASIN {len({r['ASIN'] for r in topic_rows})}개, 주제 {len(topic_rows)}개"
    except Exception:
        topic_rows = []

    print("\n【5】 분석")
    res = analyze(al, cat, details, brand)
    if api_issues or topic_rows:
        spend_ = _ad_spend()
        titles_ = {a: x["title"] for a, x in res["per_asin"].items()}
        for sku, asin, lvl, typ, det, adv in api_issues:
            res["issues"].append({"level": lvl, "type": typ, "sku": sku, "asin": asin, "title": titles_.get(asin, ""),
                                  "detail": det, "advice": adv, "spend": spend_.get(asin, 0)})
        for t in topic_rows:
            if (t.get("언급 비율(%)") or 0) and float(t["언급 비율(%)"]) >= 10:
                res["issues"].append({"level": "P2", "type": "리뷰 불만 주제", "sku": "", "asin": t["ASIN"],
                                      "title": titles_.get(t["ASIN"], ""),
                                      "detail": f"{t['주제']}: 리뷰의 {t['언급 비율(%)']}%가 언급(별점 영향 {t['별점 영향']}) · {t['예시'][:80]}",
                                      "advice": "상세 페이지 설명·사진으로 기대치를 맞추거나 상품을 개선하세요",
                                      "spend": spend_.get(t["ASIN"], 0)})
        res["issues"].sort(key=lambda x: (LEVEL_ORDER[x["level"]], -x["spend"], x["type"], x["sku"]))
    # 75자 제목·Item Highlights 초안(제안만)
    drafts = []
    try:
        import title_draft
        import ba_collect as _ba
        p_ = _ba.latest_file("sqp")
        sqp_terms = [(x["query"], x["volume"]) for x in _ba.parse_sqp(ac.load_rows(p_))][:300] if p_ else []
        drafts = title_draft.build([x for x in res["titles"] if str(x.get("상태", "")).lower() != "inactive"], brand, sqp_terms)
        sources["제목 초안"] = f"{len(drafts)}개 SKU(75자 초과, 검색어 {len(sqp_terms)}개 참고)"
        print(f"   제목 초안 {len(drafts)}개(‘제목 초안’ 시트, 검토 후 직접 수정)")
    except Exception as e:
        print(f"   ⚠ 제목 초안을 만들지 못했습니다: {type(e).__name__}: {str(e)[:100]}")
    snap = snapshot(al, cat, details)
    prev, prev_day = _prev_snapshot(today)
    chg = changes(prev, snap) if prev else []
    (SNAP_DIR / f"{today}.json").write_text(json.dumps(snap, ensure_ascii=False), encoding="utf-8")
    out = REPORT_DIR / f"listing_check_{today}.xlsx"
    write_report(out, today, res, sources, explore_rows, chg, prev_day, notes)
    try:
        import sc_pages
        extra_ = [status_page] if status_page else []
        if drafts:
            extra_.insert(0, {"name": "제목 초안", "src": "자동 생성 초안", "time": datetime.now().isoformat(timespec="seconds"),
                              "kpis": {"대상": f"75자 넘는 판매 중·미완성 SKU {len(drafts)}개",
                                       "방법": "핵심 상품명 + 검색량 높은 구절(75자) / 나머지는 Item Highlights(125자)",
                                       "주의": "자동 적용하지 않습니다. 검토 후 Seller Central에서 직접 수정하세요"},
                              "rows": drafts, "alerts": []})
        if topic_rows:
            extra_.append({"name": "리뷰 불만 주제", "src": "SP-API", "time": datetime.now().isoformat(timespec="seconds"),
                           "kpis": {"출처": "SP-API Customer Feedback(최근 6개월 리뷰)"}, "rows": topic_rows, "alerts": []})
        if extra_:
            sc_pages.append_sheets(out, extra_)
    except Exception:
        pass
    lv = Counter(x["level"] for x in res["issues"])
    try:
        import daily_bus
        t_over = sum(1 for x in res["titles"] if x["상태"].lower() != "inactive" and x["75자 초과"] > 0)
        daily_bus.record("Listing 점검", {"P0": lv.get("P0", 0), "P1": lv.get("P1", 0), "P2": lv.get("P2", 0),
                                         "제목 75자 초과(판매 중·미완성)": t_over},
                         [(x["level"], x["type"], f"{x['sku']} {x['asin']}: {x['detail']}", x["advice"])
                          for x in res["issues"] if x["level"] in ("P0", "P1")][:150], out)
        daily_bus.merge_pages("Listing 점검", [status_page] if status_page else [])
    except Exception:
        pass
    print(f"\n조치 필요: P0 {lv.get('P0', 0)}건, P1 {lv.get('P1', 0)}건, P2 {lv.get('P2', 0)}건, P3 {lv.get('P3', 0)}건")
    for x in res["issues"][:8]:
        print(f"  [{x['level']}] {x['type']} · {x['sku']}: {_short(x['detail'], 70)}")
    print(f"\n✅ 보고서 저장 완료: {out}")
    if open_report:
        try:
            import os
            import sys
            if True:
                __import__("platform_util").open_path(out)
        except Exception:
            pass
    return out


def run_offline(open_report: bool = False) -> Optional[Path]:
    """브라우저 없이: 이미 있는 파일로 다시 분석(테스트용)."""
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    al, al_name = load_all_listings(None)
    raws = sorted(RAW_DIR.glob("category_listings_*.*"))
    cat = parse_category_listings(raws[-1])[0] if raws else {}
    details = {p.stem: json.loads(p.read_text(encoding="utf-8")) for p in DETAIL_DIR.glob("*.json")} if DETAIL_DIR.exists() else {}
    res = analyze(al, cat, details, _brand())
    out = REPORT_DIR / f"listing_check_{today}.xlsx"
    write_report(out, today, res, {"전체 상품(All Listings)": f"{al_name}, {len(al)}행"}, [], [], "", [])
    return out


def last_run_text() -> str:
    files = sorted(REPORT_DIR.glob("listing_check_*.xlsx"))
    return f"마지막 점검: {files[-1].stem.replace('listing_check_', '')}" if files else "아직 실행하지 않음"


if __name__ == "__main__":
    run()

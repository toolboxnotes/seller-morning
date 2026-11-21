"""
광고 Agent 공통: 파일 위치, 날짜(미국 태평양 시간), 보고서 읽기, 열 이름 통일, 페이지 도구.

1단계는 자료 수집·제안만 하며 광고 설정은 바꾸지 않음.
"""

import json
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import human
import inventory as inv
import sc_nav

BASE_DIR = Path(__file__).resolve().parent.parent
REPORT_DIR = BASE_DIR / "reports" / "ads"
RAW_DIR = REPORT_DIR / "raw"
DEBUG_DIR = REPORT_DIR / "debug"
DATA_DIR = REPORT_DIR / "data"          # data/<미국 날짜>/xxx.xlsx  날마다의 데이터 스냅숏
STATE_FILE = REPORT_DIR / "state.json"   # 요청했고 아직 생성 중인 보고서
CONFIG_FILE = REPORT_DIR / "ads_config.xlsx"

ADS_BASE = "https://advertising.amazon.com"


# ================================================================ 날짜(미국 광고 보고서는 태평양 시간 기준)

def _us_offset(d: date, std: int) -> int:
    """미국 서머타임: 3월 둘째 일요일 ~ 11월 첫째 일요일."""
    mar = date(d.year, 3, 8 + (6 - date(d.year, 3, 8).weekday()) % 7)
    nov = date(d.year, 11, 1 + (6 - date(d.year, 11, 1).weekday()) % 7)
    return std + 1 if mar <= d < nov else std


def pacific_now() -> datetime:
    utc = datetime.now(timezone.utc)
    off = _us_offset(utc.date(), -8)
    return (utc + timedelta(hours=off)).replace(tzinfo=None)


def pacific_today() -> date:
    return pacific_now().date()


def ranges(today: Optional[date] = None) -> Dict[str, Tuple[date, date]]:
    """전날, 최근 7일, 규칙 기간(최근 14일, 최근 3일 제외), 최근 30일."""
    t = today or pacific_today()
    y = t - timedelta(days=1)
    return {
        "yesterday": (y, y),
        "7d": (y - timedelta(days=6), y),
        "prev7d": (y - timedelta(days=13), y - timedelta(days=7)),
        "rule14": (y - timedelta(days=16), y - timedelta(days=3)),
        "30d": (y - timedelta(days=29), y),
    }


def day_dir(today: Optional[date] = None) -> Path:
    d = DATA_DIR / (today or pacific_today()).isoformat()
    d.mkdir(parents=True, exist_ok=True)
    return d


# ================================================================ 상태(생성 중인 보고서)

def load_state() -> Dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(state: Dict) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    if "collected" not in state:          # 같은 날이면 ‘수집 완료’ 표시는 유지
        old = load_state()
        if old.get("date") == state.get("date") and old.get("collected"):
            state = {**state, "collected": True}
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


# ================================================================ 받은 파일 읽기

def read_any(path: Path, header_hint: Optional[List[str]] = None) -> Tuple[List[str], List[Dict]]:
    """csv / txt / xlsx 읽기. 파일 앞에 설명 줄이 있으면(Brand Analytics 다운로드가 흔히 그럼) 실제 머리글 줄을 자동으로 찾음."""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        rows = _xlsx_rows(path)
    else:
        import csv
        import io
        data = path.read_bytes()
        text = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        first = text.split("\n", 1)[0]
        delim = "\t" if text.count("\t") > text.count(",") else ","
        rows = [r for r in csv.reader(io.StringIO(text), delimiter=delim)]
    rows = [[("" if c is None else str(c)).strip() for c in r] for r in rows if any(str(c or "").strip() for c in r)]
    if not rows:
        return [], []
    hi = 0
    if header_hint:
        hints = [inv.norm(h) for h in header_hint]
        for i, r in enumerate(rows[:15]):
            normed = [inv.norm(c) for c in r]
            if any(any(h in c for c in normed) for h in hints) and len([c for c in r if c]) >= 3:
                hi = i
                break
    headers = rows[hi]
    out = []
    for r in rows[hi + 1:]:
        r = r + [""] * (len(headers) - len(r))
        out.append({h: v for h, v in zip(headers, r)})
    return headers, out


def _xlsx_rows(path: Path) -> List[List]:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def save_rows(rows: List[Dict], path: Path, title: str = "data") -> None:
    headers = []
    for r in rows:
        for k in r:
            if k not in headers:
                headers.append(k)
    inv.save_report_xlsx(headers, rows, path, title[:30])


def load_rows(path: Path) -> List[Dict]:
    if not path.exists():
        return []
    try:
        return read_any(path)[1]
    except Exception:
        return []


# ================================================================ 열 이름 통일

def num(v) -> float:
    if v in (None, "", "—", "-", "--"):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("US$", "").replace("$", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", s)      # 예: “$5.00 - daily”
        return float(m.group(0)) if m else 0.0


# 통일한 필드 이름: [가능한 열 이름(우선순위 순)]. 비교할 때 대소문자·기호는 무시.
FIELDS = {
    "date": ["date", "day", "start date"],
    "campaign": ["campaign name", "campaign"],
    "campaign_id": ["campaign id"],
    "ad_group": ["ad group name", "ad group"],
    "ad_group_id": ["ad group id"],
    "portfolio": ["portfolio name", "portfolio"],
    "status": ["campaign status", "status", "state"],
    "targeting_type": ["targeting type", "campaign targeting type", "type"],
    "budget": ["campaign budget amount", "budget", "daily budget", "campaign budget"],
    "targeting": ["targeting", "keyword text", "keyword", "target", "targeting expression", "target value"],
    "match_type": ["match type", "target match type", "keyword match type", "keyword type"],
    "search_term": ["customer search term", "search term", "query"],
    "placement": ["placement", "placement type", "placement classification"],
    "asin": ["advertised asin", "advertised product id", "asin"],
    "sku": ["advertised sku", "advertised product sku", "sku"],
    "ad_product": ["ad product"],
    "bid": ["keyword bid", "target bid", "bid", "max bid"],
    "impressions": ["impressions"],
    "clicks": ["clicks"],
    "spend": ["spend", "total cost", "cost"],
    "sales": ["7 day total sales", "sales", "14 day total sales", "total sales", "attributed sales"],
    "orders": ["7 day total orders", "purchases", "orders", "14 day total orders", "total orders"],
    "units": ["7 day total units", "units sold", "units", "14 day total units"],
    "tos_is": ["top of search impression share", "top-of-search impression share", "top-of-search is"],
}
# 열 이름에 이 단어가 있으면 위 필드로 보지 않음(예: ‘신규 고객 매출’은 전체 매출이 아님)
EXCLUDE = ("new to brand", "ntb", "long term", "halo", "other sku", "same sku", "advertised sku sales",
           "advertised sku units", "brand halo", "percentage", "rate", "share of", "clicks rate", "b2b",
           "view through", "viewable", "per ", "cost per", "acos", "roas", "ctr", "cpc", "cvr")


def field_map(headers: List[str]) -> Dict[str, str]:
    """보고서 열 이름을 통일한 필드에 연결."""
    normed = {h: inv.norm(h) for h in headers}
    out = {}
    for field, names in FIELDS.items():
        cands = [inv.norm(n) for n in names]
        found = None
        for c in cands:                               # 먼저 완전히 같은 것
            for h, n in normed.items():
                if n == c and h not in out.values():
                    found = h
                    break
            if found:
                break
        if not found:                                 # 그다음 ‘앞부분 같음 / 포함’, 헷갈리는 열은 제외
            for c in cands:
                for h, n in normed.items():
                    if h in out.values():
                        continue
                    low = h.lower()
                    if field not in ("tos_is",) and any(x in low for x in EXCLUDE):
                        continue
                    if n.startswith(c) or (len(c) > 4 and c in n):
                        found = h
                        break
                if found:
                    break
        if found:
            out[field] = found
    return out


def normalize(rows: List[Dict], extra: Optional[Dict[str, str]] = None) -> List[Dict]:
    """필드 통일 + ACoS / CTR / CPC / CVR 계산. 원래 열(raw: 로 시작)도 남겨 확인할 수 있게."""
    if not rows:
        return []
    fm = field_map(list(rows[0].keys()))
    out = []
    numeric = ("impressions", "clicks", "spend", "sales", "orders", "units", "budget", "bid", "tos_is")
    for r in rows:
        x = {}
        for f, h in fm.items():
            v = r.get(h)
            x[f] = num(v) if f in numeric else ("" if v is None else str(v).strip())
        if "date" in x:
            x["date"] = parse_date(x["date"])
        for f in numeric:
            x.setdefault(f, 0.0)
        if extra:
            x.update(extra)
        out.append(x)
    return out


_DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y", "%d/%m/%Y", "%Y%m%d", "%m/%d/%y", "%Y/%m/%d")


def parse_date(v) -> Optional[str]:
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()[:12].strip()
    for f in _DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date().isoformat()
        except ValueError:
            continue
    m = re.match(r"(\d{4}-\d{2}-\d{2})", str(v))
    return m.group(1) if m else None


def kpi(rows: List[Dict]) -> Dict:
    imp = sum(r.get("impressions", 0) for r in rows)
    clk = sum(r.get("clicks", 0) for r in rows)
    sp = sum(r.get("spend", 0) for r in rows)
    sa = sum(r.get("sales", 0) for r in rows)
    od = sum(r.get("orders", 0) for r in rows)
    return {"impressions": imp, "clicks": clk, "spend": round(sp, 2), "sales": round(sa, 2), "orders": od,
            "ctr": clk / imp if imp else 0, "cpc": sp / clk if clk else 0, "cvr": od / clk if clk else 0,
            "acos": sp / sa if sa else (None if not sp else float("inf")), "roas": sa / sp if sp else 0}


def in_range(r: Dict, rng: Tuple[date, date]) -> bool:
    d = r.get("date")
    if not d:
        return False
    return rng[0].isoformat() <= d <= rng[1].isoformat()


# ================================================================ 페이지 도구

def dump(page, name: str) -> Path:
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


def settle(page, idle_ms: int = 20000) -> None:
    sc_nav._settle(page)
    try:
        page.wait_for_load_state("networkidle", timeout=idle_ms)
    except Exception:
        pass


def first_visible(page_or_loc, selectors: List[str]):
    for s in selectors:
        try:
            el = sc_nav._first_visible(page_or_loc.locator(s))
        except Exception:
            el = None
        if el is not None:
            return el
    return None


def by_text(page, pattern: str, roles=("button", "link", "menuitem", "option", "tab", "radio", "checkbox")):
    rx = re.compile(pattern, re.I)
    for role in roles:
        try:
            el = sc_nav._first_visible(page.get_by_role(role, name=rx))
        except Exception:
            el = None
        if el is not None:
            return el
    try:
        return sc_nav._first_visible(page.get_by_text(rx))
    except Exception:
        return None

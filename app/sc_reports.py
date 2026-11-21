"""
Seller Central 보고서: 요청 → 생성 기다림 → 다운로드

공통 방법:
  1. 요청 전에 페이지에 이미 있는 ‘다운로드’ 버튼마다 그 줄의 글자를 기록(예전 보고서)
  2. ‘보고서 요청’ 클릭
  3. 2~3분마다(무작위, 밀리초 단위) 새로 고침, 예전 보고서가 아닌 ‘다운로드’ 버튼이 나올 때까지
  4. 그것을 눌러 다운로드
"""

import random
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional

import human
from sc_nav import find_visible, save_debug

# 보고서를 기다리며 새로 고치는 간격(초): 2~3분 무작위
POLL_MIN_S, POLL_MAX_S = 120, 180

DOWNLOAD_RE = re.compile(r"^\s*(Download|다운로드|下载)\s*$", re.I)
REFRESH_LABELS = ["Refresh", "새로 고침", "새로고침", "刷新"]

# 다운로드 버튼이 있는 ‘줄’의 글자 — 새 보고서와 예전 보고서 구분용
_ROW_TEXT_JS = """
(el) => {
  // 표 안의 줄만 인정, 표 줄 밖의 “Download”(예: 페이지 위쪽 Download 탭)는 빈 값
  let n = el;
  for (let i = 0; i < 10 && n; i++) {
    if (n.matches && n.matches('tr, [role="row"], kat-table-row')) {
      return (n.innerText || '').replace(/\\s+/g, ' ').trim();
    }
    n = n.parentElement || (n.getRootNode && n.getRootNode().host) || null;
  }
  return '';
}
"""


def _download_controls(page) -> list:
    loc = (page.locator('kat-button[label^="Download" i], kat-button[label="다운로드"]')
           .or_(page.get_by_role("button", name=DOWNLOAD_RE))
           .or_(page.get_by_role("link", name=DOWNLOAD_RE)))
    out = []
    for i in range(min(loc.count(), 50)):
        el = loc.nth(i)
        try:
            if el.is_visible() and row_text(el):   # 반드시 보고서 표의 어떤 줄 안에 있어야 함
                out.append(el)
        except Exception:
            pass
    return out


def row_text(el) -> str:
    try:
        return el.evaluate(_ROW_TEXT_JS)
    except Exception:
        return ""


def existing_rows(page) -> set:
    """요청 전 페이지에 이미 있는, 다운로드할 수 있는 보고서."""
    return {row_text(el) for el in _download_controls(page)}


def click_request(page, labels: List[str]) -> None:
    el = find_visible(page, labels)
    if el is None:
        # 글자가 완전히 같지 않으면 포함 관계로 버튼을 찾음
        pattern = re.compile("|".join(re.escape(l) for l in labels), re.I)
        loc = page.get_by_role("button", name=pattern)
        el = next((loc.nth(i) for i in range(min(loc.count(), 10)) if loc.nth(i).is_visible()), None)
    if el is None:
        raise RuntimeError(f"“{labels[0]}” 버튼을 찾지 못했습니다")
    human.click(page, el)


STATUS_RE = re.compile(r"^\s*(In Queue|In Progress|Processing|Submitted|Pending|Requested|Waiting|대기 중|대기|진행 중|처리 중|排队中|处理中|进行中)\s*$", re.I)


def _report_days() -> list:
    """
    ‘오늘’ 날짜. Amazon 보고서 센터는 미국 중부 시간으로 날짜를 보여 줌(한국 정오에 미국은 아직 전날),
    그래서 이 컴퓨터 날짜와 태평양 시간 날짜를 모두 오늘로 봄.
    """
    from datetime import date, timezone
    utc = datetime.now(timezone.utc)
    days = {date.today()}
    for h in (5, 6):  # 미국 중부: 서머타임 UTC-5, 겨울 UTC-6
        days.add((utc - timedelta(hours=h)).date())
    return sorted(days)


def today_strings(d=None) -> List[str]:
    """페이지에 ‘오늘’이 나올 수 있는 몇 가지 표기."""
    months = ["January", "February", "March", "April", "May", "June", "July", "August",
              "September", "October", "November", "December"]
    out = []
    for d in ([d] if d else _report_days()):
        full, short = months[d.month - 1], months[d.month - 1][:3]
        out += [f"{full} {d.day}, {d.year}", f"{short} {d.day}, {d.year}", f"{short}. {d.day}, {d.year}",
                d.isoformat(), f"{d.month:02d}/{d.day:02d}/{d.year}", f"{d.month}/{d.day}/{d.year}",
                f"{d.month}/{d.day}/{d.year % 100:02d},", f"{d.year}. {d.month}. {d.day}",
                f"{d.year}년 {d.month}월 {d.day}일", f"{d.year}年{d.month}月{d.day}日"]
    return out


_MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug",
                                          "sep", "oct", "nov", "dec"], start=1)}
_DT_LONG = re.compile(r"\b([A-Z][a-z]+)\.? (\d{1,2}), (\d{4})(?: at)? (\d{1,2}):(\d{2}) ?([AP]M)", re.I)
_DT_SHORT = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4}),? (\d{1,2}):(\d{2}) ?([AP]M)", re.I)


def _us_offset(d, std: int) -> int:
    """미국 시간대와 UTC의 차이: 서머타임(3월 둘째 일요일 ~ 11월 첫째 일요일)은 std+1, 나머지는 std."""
    from datetime import date as _d
    def nth_sunday(y, m, n):
        first = _d(y, m, 1)
        return first + timedelta(days=(6 - first.weekday()) % 7 + 7 * (n - 1))
    start, end = nth_sunday(d.year, 3, 2), nth_sunday(d.year, 11, 1)
    return std + 1 if start <= d.date() < end else std


# 페이지가 보여 주는 시간대 → 표준시와 UTC의 차이
US_TZ = {"central": -6, "pacific": -8, "eastern": -5}


def row_local_time(text: str, page_tz: str = "central") -> Optional[datetime]:
    """
    보고서 줄 글자에서 시각을 읽어 이 컴퓨터 시각으로 바꿈. 못 읽으면 None.
    page_tz: 이 페이지가 보여 주는 시간대. 상품 보고서 페이지는 이 컴퓨터 시각("local"), 보고서 센터(FBA)는 미국 중부 시간("central").
    """
    m = _DT_LONG.search(text)
    if m and m.group(1)[:3].lower() in _MONTHS:
        mon, day, year, hh, mm, ap = _MONTHS[m.group(1)[:3].lower()], int(m.group(2)), int(m.group(3)), \
            int(m.group(4)), int(m.group(5)), m.group(6).upper()
    else:
        m = _DT_SHORT.search(text)
        if not m:
            return None
        mon, day, year, hh, mm, ap = int(m.group(1)), int(m.group(2)), int(m.group(3)), \
            int(m.group(4)), int(m.group(5)), m.group(6).upper()
        if year < 100:
            year += 2000
    hh = hh % 12 + (12 if ap == "PM" else 0)
    try:
        pt = datetime(year, mon, day, hh, mm)
    except ValueError:
        return None
    if page_tz == "local":
        return pt
    from datetime import timezone
    utc = (pt - timedelta(hours=_us_offset(pt, US_TZ.get(page_tz, -6)))).replace(tzinfo=timezone.utc)
    local = utc.astimezone().replace(tzinfo=None)
    # 바꾼 결과가 미래면 페이지가 원래 이 컴퓨터 시각을 보여 준 것
    if local > datetime.now() + timedelta(minutes=10):
        return pt
    return local


def is_today_row(text: str, page_tz: str = "central") -> bool:
    """이 줄의 보고서가 이 컴퓨터 시각으로 오늘 만들어졌는지."""
    from datetime import date as _d
    t = row_local_time(text, page_tz)
    if t is not None:
        return t.date() == _d.today()
    low = text.lower()
    return any(s.lower() in low for s in today_strings(_d.today()))


FILE_TYPES = ("text/plain", "text/csv", "text/tab-separated-values", "application/octet-stream",
              "application/vnd.ms-excel", "application/csv", "application/x-download", "binary/octet-stream")


def _ext_for(url: str, ctype: str, disposition: str = "") -> str:
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', disposition or "", re.I)
    if m and Path(m.group(1)).suffix:
        return Path(m.group(1)).suffix
    ext = Path(url.split("?")[0]).suffix
    if ext and len(ext) <= 5:
        return ext
    return ".csv" if "csv" in (ctype or "") else ".txt"


def _download(page, el, save_dir: Path, save_name: str, debug_dir: Optional[Path] = None) -> Path:
    """
    ‘다운로드’를 눌러 파일 저장. Amazon이 파일을 주는 방식이 여러 가지라 모두 처리:
      - 브라우저가 바로 다운로드(지금 탭, 새 탭, 숨은 프레임)
      - 페이지 스크립트가 파일 내용을 먼저 받아 옴(네트워크 응답이 파일)
      - 새 탭이나 지금 탭에 파일 내용이 바로 보임(.txt가 흔히 그럼)
    아무것도 못 받으면 클릭 뒤 네트워크 요청·새 탭·콘솔 정보를 debug 폴더에 기록해 원인을 찾을 수 있게.
    """
    ctx = page.context
    got, new_pages, files, events = [], [], [], []

    def on_download(d):
        got.append(d)
        events.append(f"download: {d.suggested_filename}")

    def on_page(p_):
        new_pages.append(p_)
        p_.on("download", on_download)
        p_.on("response", on_response)
        events.append(f"new page: {p_.url}")

    def on_response(r):
        try:
            h = r.headers
            ctype = (h.get("content-type") or "").lower()
            disp = h.get("content-disposition") or ""
            if r.request.resource_type in ("xhr", "fetch", "document", "other") and len(events) < 400:
                events.append(f"{r.status} {r.request.method} [{r.request.resource_type}] {ctype[:40]} {r.url[:180]}")
            if r.status == 200 and ("attachment" in disp.lower() or any(t in ctype for t in FILE_TYPES)):
                files.append((r, ctype, disp))
        except Exception:
            pass

    def on_console(m):
        if len(events) < 400:
            events.append(f"console {m.type}: {m.text[:200]}")

    for p_ in ctx.pages:
        p_.on("download", on_download)
        p_.on("response", on_response)
    page.on("console", on_console)
    ctx.on("page", on_page)

    def save_bytes(data: bytes, ext: str) -> Path:
        save_dir.mkdir(parents=True, exist_ok=True)
        out = save_dir / f"{save_name}{ext}"
        out.write_bytes(data)
        return out

    out = None
    try:
        start_url = page.url
        human.click(page, el, wait_after=False)
        t0 = time.time()
        end = t0 + 120
        clicked_again = False
        while time.time() < end and out is None:
            page.wait_for_timeout(1000)
            if got:
                d = got[0]
                out = save_dir / f"{save_name}{Path(d.suggested_filename).suffix or '.txt'}"
                save_dir.mkdir(parents=True, exist_ok=True)
                d.save_as(str(out))
                break
            # 네트워크 응답 자체가 파일
            for r, ctype, disp in list(files):
                try:
                    body = r.body()
                except Exception:
                    continue
                if body:
                    out = save_bytes(body, _ext_for(r.url, ctype, disp))
                    break
            if out is not None:
                break
            # 새 탭이나 지금 탭이 파일 주소로 이동: 같은 로그인 상태로 받아 옴
            for p_ in new_pages + [page]:
                try:
                    url = p_.url
                except Exception:
                    continue
                if not url or url == "about:blank" or url == start_url or "/reportcentral/" in url:
                    continue
                try:
                    resp = ctx.request.get(url)
                except Exception:
                    continue
                ctype = (resp.headers.get("content-type") or "").lower()
                if resp.ok and "text/html" not in ctype:
                    out = save_bytes(resp.body(), _ext_for(url, ctype, resp.headers.get("content-disposition", "")))
                    if p_ is page:
                        page.go_back(wait_until="domcontentloaded")
                    break
            # 20초 동안 아무 반응 없음: 한 번 더 누름(첫 클릭은 버튼에 초점만 주는 경우가 있음)
            if out is None and not clicked_again and time.time() - t0 > 20 and not new_pages:
                clicked_again = True
                events.append("---- 20초 동안 파일을 받지 못해 한 번 더 클릭 ----")
                human.click(page, el, wait_after=False)
        if out is None:
            if debug_dir:
                debug_dir.mkdir(parents=True, exist_ok=True)
                log = debug_dir / f"{datetime.now():%Y%m%d_%H%M%S}_{save_name}_download_events.txt"
                log.write_text("Download 클릭 후 발생한 일:\n\n" + "\n".join(events or ["(네트워크 요청, 새 탭, 다운로드가 전혀 없었음)"]),
                               encoding="utf-8")
            raise TimeoutError("다운로드 클릭 후 2분 안에 파일을 받지 못했습니다")
    finally:
        ctx.remove_listener("page", on_page)
        try:
            page.remove_listener("console", on_console)
        except Exception:
            pass
        for p_ in ctx.pages:
            for ev, fn in (("download", on_download), ("response", on_response)):
                try:
                    p_.remove_listener(ev, fn)
                except Exception:
                    pass
    for np_ in new_pages:  # 다운로드용 새 탭은 닫고 보고서 페이지로 돌아옴
        try:
            if not np_.is_closed():
                np_.close()
        except Exception:
            pass
    human.pause()
    return out


def download_ready_today(page, save_dir: Path, save_name: str,
                         row_must_contain: Optional[List[str]] = None, page_tz: str = "central") -> Optional[Path]:
    """
    페이지에 이미 ‘오늘 만들어져 다운로드할 수 있는’ 같은 종류 보고서가 있으면 그것을 받고 다시 요청하지 않음.
    없으면 None.
    """
    for el in _download_controls(page):
        text = row_text(el).lower()
        if not is_today_row(row_text(el), page_tz):
            continue
        if row_must_contain and not any(k.lower() in text for k in row_must_contain):
            continue
        return _download(page, el, save_dir, save_name)
    return None


def pending_today(page, row_must_contain: Optional[List[str]] = None, page_tz: str = "central") -> bool:
    """오늘 이미 요청했고 아직 대기·생성 중인 같은 종류 보고서."""
    loc = page.get_by_text(STATUS_RE)
    for i in range(min(loc.count(), 20)):
        el = loc.nth(i)
        try:
            if not el.is_visible():
                continue
        except Exception:
            continue
        text = row_text(el).lower()
        if is_today_row(row_text(el), page_tz) and (not row_must_contain or any(k.lower() in text for k in row_must_contain)):
            return True
    return False


NO_DATA_RE = re.compile(r"No Data Available|데이터 없음|无可用数据|没有数据", re.I)


def _no_data_count(page) -> int:
    try:
        return page.get_by_text(NO_DATA_RE).count()
    except Exception:
        return 0


def _current_status(page) -> str:
    loc = page.get_by_text(STATUS_RE)
    for i in range(min(loc.count(), 5)):
        try:
            if loc.nth(i).is_visible():
                return loc.nth(i).inner_text().strip()
        except Exception:
            pass
    return ""


def wait_and_download(page, old_rows: set, save_dir: Path, save_name: str,
                      row_must_contain: Optional[List[str]] = None,
                      timeout_min: int = 45, debug_dir: Optional[Path] = None,
                      on_long_wait=None, long_wait_min: float = 8) -> Path:
    """
    새 보고서가 만들어지길 기다려 다운로드하고 save_dir / save_name 으로 저장(확장자는 받은 파일 그대로).
    row_must_contain: 새 보고서 줄에 이 글자 중 하나가 있어야 함(예: 보고서 이름) — 다른 보고서를 받지 않게.
    Amazon 설명으로는 보고서가 최대 45분 걸릴 수 있음.
    """
    start = time.time()
    deadline = start + timeout_min * 60
    no_data_before = _no_data_count(page)
    long_wait_done = False
    while time.time() < deadline:
        if on_long_wait and not long_wait_done and time.time() - start > long_wait_min * 60:
            long_wait_done = True
            try:
                on_long_wait(page)   # 너무 오래 기다리면 페이지 구조를 한 번 저장(새 보고서를 못 알아봤는지 확인용)
            except Exception:
                pass
        if _no_data_count(page) > no_data_before:
            if debug_dir:
                save_debug(page, debug_dir, f"no_data_{save_name}")
            raise RuntimeError("아마존이 No Data Available을 반환했습니다(이번 보고서에 데이터 없음). 잠시 후 다시 실행하세요")
        for el in _download_controls(page):
            text = row_text(el)
            if text in old_rows:
                continue
            if row_must_contain and not any(k.lower() in text.lower() for k in row_must_contain):
                continue
            return _download(page, el, save_dir, save_name, debug_dir)

        mins = (time.time() - start) / 60
        status = _current_status(page)
        wait_s = round(random.uniform(POLL_MIN_S, POLL_MAX_S), 3)
        nxt = datetime.now() + timedelta(seconds=wait_s)
        print(f"     보고서 생성 중{f'({status})' if status else ''}, 대기 {mins:.0f}분, "
              f"{nxt:%H:%M:%S}에 다시 새로고침(최대 {timeout_min}분)...")
        time.sleep(wait_s)
        btn = find_visible(page, REFRESH_LABELS)
        if btn is not None:
            human.click(page, btn)
        else:
            page.reload(wait_until="domcontentloaded")
            human.pause()

    if debug_dir:
        save_debug(page, debug_dir, f"timeout_{save_name}")
    raise TimeoutError(f"대기 {timeout_min}분이 지나도 보고서가 생성되지 않았습니다. 잠시 후 다시 실행하세요(생성되면 바로 다운로드합니다)")

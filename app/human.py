"""
사람처럼 브라우저 조작

모든 기능은 여기 함수로 클릭·스크롤·이동해야 자연스러운 속도가 유지됨:
  - 조작 사이에 2~5초 무작위 대기(밀리초 단위)
  - 마우스는 무작위 곡선을 따라 목표로 움직이고 순간 이동하지 않음
  - 클릭 위치는 요소 안에서 무작위, 누르고 떼는 사이에 짧은 멈춤
  - 스크롤은 작게 여러 번 나눠서, 중간에 ‘읽는’ 멈춤

사용법:
    from human import pause, click, scroll_by, read_page, goto
"""

import math
import random
import time
from typing import Optional

from playwright.sync_api import Locator, Page

# 조작 간격(초). 요구 사항대로 2~5초, 밀리초까지 무작위.
ACTION_MIN = 2.0
ACTION_MAX = 5.0

# 페이지마다 마우스 현재 위치 기록(Playwright는 읽는 기능이 없음)
_mouse_pos: dict = {}


# ---------------------------------------------------------------- 대기

def pause(min_s: float = ACTION_MIN, max_s: float = ACTION_MAX) -> float:
    """밀리초 단위 무작위 대기. 실제 기다린 초를 돌려줌."""
    seconds = round(random.uniform(min_s, max_s), 3)
    time.sleep(seconds)
    return seconds


def _short(min_s: float, max_s: float) -> None:
    """동작 안의 짧은 멈춤(마우스 올리기, 키 누르는 시간 등) — 2~5초 조작 간격에 포함되지 않음."""
    time.sleep(random.uniform(min_s, max_s))


# ---------------------------------------------------------------- 마우스

def _viewport(page: Page) -> tuple:
    w, h = page.evaluate("() => [window.innerWidth, window.innerHeight]")
    return w, h


def _bezier(p0, p1, p2, p3, t):
    u = 1 - t
    return (
        u**3 * p0[0] + 3 * u**2 * t * p1[0] + 3 * u * t**2 * p2[0] + t**3 * p3[0],
        u**3 * p0[1] + 3 * u**2 * t * p1[1] + 3 * u * t**2 * p2[1] + t**3 * p3[1],
    )


def move_mouse(page: Page, x: float, y: float) -> None:
    """무작위 베지어 곡선을 따라 마우스를 (x, y)로 이동, 처음엔 빠르게 끝엔 느리게."""
    key = id(page)
    if key not in _mouse_pos:
        w, h = _viewport(page)
        _mouse_pos[key] = (random.uniform(w * 0.3, w * 0.7), random.uniform(h * 0.3, h * 0.7))
    sx, sy = _mouse_pos[key]
    dist = math.hypot(x - sx, y - sy)
    if dist < 2:
        return

    # 두 조절점을 시작·끝 연결선 양쪽으로 무작위로 벗어나게 해 자연스러운 곡선
    spread = min(dist * 0.3, 150)
    c1 = (sx + (x - sx) * random.uniform(0.2, 0.4) + random.uniform(-spread, spread),
          sy + (y - sy) * random.uniform(0.2, 0.4) + random.uniform(-spread, spread))
    c2 = (sx + (x - sx) * random.uniform(0.6, 0.8) + random.uniform(-spread, spread),
          sy + (y - sy) * random.uniform(0.6, 0.8) + random.uniform(-spread, spread))

    steps = int(min(max(dist / 12, 15), 60))
    for i in range(1, steps + 1):
        t = i / steps
        t = 1 - (1 - t) ** 2  # ease-out: 처음엔 빠르고 목표에 가까워지면 느려짐
        px, py = _bezier((sx, sy), c1, c2, (x, y), t)
        page.mouse.move(px, py)
        time.sleep(random.uniform(0.006, 0.018))

    # 가끔 살짝 지나쳤다가 다시 맞춤
    if random.random() < 0.25:
        ox, oy = x + random.uniform(-4, 4), y + random.uniform(-4, 4)
        page.mouse.move(ox, oy)
        _short(0.03, 0.08)
        page.mouse.move(x, y)

    _mouse_pos[key] = (x, y)


# ---------------------------------------------------------------- 스크롤

def scroll_by(page: Page, total_px: float) -> None:
    """작게 여러 번 나눠 스크롤, 전체 거리 약 total_px(양수는 아래, 음수는 위)."""
    direction = 1 if total_px > 0 else -1
    remaining = abs(total_px)
    while remaining > 0:
        step = min(remaining, random.uniform(80, 220))
        page.mouse.wheel(0, direction * step)
        remaining -= step
        _short(0.05, 0.25)
        if random.random() < 0.15:  # 가끔 멈춰서 ‘한 번 봄’
            _short(0.4, 1.2)


def _scroll_into_view(page: Page, locator: Locator) -> Optional[dict]:
    """휠로 요소를 화면 가운데 근처까지 스크롤하고 요소 위치를 돌려줌."""
    _, vh = _viewport(page)
    last_top = None
    for _ in range(30):
        box = locator.bounding_box()
        if box is None:
            return None
        top, bottom = box["y"], box["y"] + box["height"]
        if top >= 80 and bottom <= vh - 80:
            return box
        # 스크롤해도 안 움직이면(위에 고정된 도구 막대, 페이지 끝) 화면 안에만 있으면 그대로 사용
        if last_top is not None and abs(top - last_top) < 2:
            if top >= 0 and bottom <= vh:
                return box
            break
        last_top = top
        target = vh * random.uniform(0.35, 0.55)
        delta = (top + box["height"] / 2) - target
        scroll_by(page, delta)
        _short(0.2, 0.5)
    return locator.bounding_box()


# ---------------------------------------------------------------- 클릭과 이동

# 화면의 (x, y) 점에서 가장 위 요소가 목표(또는 목표 안)인지 확인, 웹 컴포넌트를 통과해서
_HIT_JS = """
([el, x, y]) => {
  let t = document.elementFromPoint(x, y);
  while (t && t.shadowRoot) {
    const inner = t.shadowRoot.elementFromPoint(x, y);
    if (!inner || inner === t) break;
    t = inner;
  }
  if (!t) return false;
  const up = (n) => n.parentElement || (n.getRootNode && n.getRootNode().host) || null;
  const within = (a, b) => { for (let n = b; n; n = up(n)) if (n === a) return true; return false; };
  // 목표 자체, 목표 안의 요소, 목표를 감싼 컴포넌트 껍데기를 누르면 모두 맞음
  return within(el, t) || (within(t, el) && t !== document.body && t !== document.documentElement);
}
"""


def move_away(page: Page) -> None:
    """마우스를 페이지 오른쪽 빈 곳으로 옮겨 마우스를 올리면 뜨는 설명(tooltip)이 사라지게."""
    w, h = _viewport(page)
    move_mouse(page, w * random.uniform(0.86, 0.95), h * random.uniform(0.55, 0.85))
    _short(0.6, 1.2)


def _hits(locator: Locator, x: float, y: float) -> bool:
    try:
        return bool(locator.evaluate(_HIT_JS, [x, y]))
    except Exception:
        return True   # 판단할 수 없으면 클릭을 막지 않음


def click(page: Page, locator: Locator, wait_after: bool = True) -> None:
    """
    사람처럼 클릭: 요소까지 스크롤 → 곡선으로 마우스 이동 → 잠깐 올려 둠 → 누름/뗌 → 2~5초 무작위 대기.
    클릭 전에 그 점이 다른 것(예: 마우스를 올리면 뜨는 설명)에 가려 있지 않은지 확인, 가려 있으면 마우스를 비켰다가 다른 위치로 다시.
    """
    locator = locator.first
    locator.wait_for(state="visible", timeout=15000)
    for attempt in range(4):
        box = _scroll_into_view(page, locator)
        if box is None:
            raise RuntimeError("클릭할 요소의 위치를 찾지 못했습니다")
        # 요소 안 가운데 쪽 무작위 위치를 누름(항상 정중앙은 아님)
        x = box["x"] + box["width"] * random.uniform(0.3, 0.7)
        y = box["y"] + box["height"] * random.uniform(0.35, 0.65)
        if _hits(locator, x, y):
            break
        move_away(page)   # 가려짐: 마우스를 치우고 뜬 창이 사라질 때까지 기다림
    move_mouse(page, x, y)
    _short(0.12, 0.45)          # 마우스 올려 두기
    if not _hits(locator, x, y):   # 이동 중에 다른 요소를 지나며 또 설명 창이 뜸
        move_away(page)
        box = _scroll_into_view(page, locator) or box
        x = box["x"] + box["width"] * random.uniform(0.3, 0.7)
        y = box["y"] + box["height"] * random.uniform(0.35, 0.65)
        move_mouse(page, x, y)
        _short(0.12, 0.3)
    page.mouse.down()
    _short(0.05, 0.14)          # 누르고 있기
    page.mouse.up()

    if wait_after:
        pause()


def click_box(page: Page, box: dict, wait_after: bool = True) -> None:
    """사각형 영역(페이지 좌표) 안 무작위 점을 사람처럼 클릭 — 요소는 못 찾고 위치만 알 때."""
    x = box["x"] + box["width"] * random.uniform(0.15, 0.6)
    y = box["y"] + box["height"] * random.uniform(0.35, 0.65)
    move_mouse(page, x, y)
    _short(0.12, 0.45)
    page.mouse.down()
    _short(0.05, 0.14)
    page.mouse.up()
    if wait_after:
        pause()


def type_text(page: Page, locator: Locator, text: str, clear: bool = True, wait_after: bool = True) -> None:
    """사람처럼 입력: 입력란 클릭 → 전체 선택해 지움 → 한 글자씩 입력(글자 사이 80~260밀리초 무작위)."""
    click(page, locator, wait_after=False)
    _short(0.2, 0.5)
    if clear:
        page.keyboard.press("Control+A")
        _short(0.08, 0.2)
        page.keyboard.press("Backspace")
        _short(0.15, 0.4)
    for ch in text:
        page.keyboard.type(ch)
        time.sleep(random.uniform(0.08, 0.26))
    _short(0.2, 0.5)
    if wait_after:
        pause()


def goto(page: Page, url: str, wait_after: bool = True) -> None:
    """주소로 바로 열기(주소창에 입력하는 것과 같음), 그다음 2~5초 무작위 대기."""
    page.goto(url, wait_until="domcontentloaded")
    if wait_after:
        pause()


def read_page(page: Page, max_screens: int = 8, back_to_top: bool = True) -> None:
    """
    사람처럼 페이지를 위에서 아래로 훑어봄: 한 번에 화면의 일부만큼 스크롤하고 ‘읽는’ 멈춤,
    끝까지 가면 맨 위로 돌아옴. 늦게 불러오는 내용도 불러옴.
    """
    _, vh = _viewport(page)
    for _ in range(max_screens):
        at_bottom = page.evaluate(
            "() => window.innerHeight + window.scrollY >= document.body.scrollHeight - 10"
        )
        if at_bottom:
            break
        scroll_by(page, vh * random.uniform(0.45, 0.8))
        _short(0.8, 2.0)  # 읽기
    if back_to_top:
        _short(0.5, 1.2)
        current = page.evaluate("() => window.scrollY")
        if current > 0:
            scroll_by(page, -current)
    pause()

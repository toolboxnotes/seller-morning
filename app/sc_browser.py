"""
Seller Central 로컬 브라우저 세션 관리

사용법:
    python sc_browser.py login            # 처음 로그인 / 세션 만료 뒤 다시 로그인(직접 진행)
    python sc_browser.py check            # 저장된 로그인이 아직 유효한지 확인
    python sc_browser.py check --headless # 창 없이 확인

다른 스크립트에서 재사용:
    from sc_browser import open_session, human_pause
    with open_session() as page:
        page.goto("https://sellercentral.amazon.com/...")
"""

import argparse
import os
import random
import sys
import time
from contextlib import contextmanager
from pathlib import Path

from playwright.sync_api import BrowserContext, Page, Playwright, sync_playwright
from playwright.sync_api import Error as PlaywrightError

import human

SC_HOME = "https://sellercentral.amazon.com/home"

# 이 컴퓨터에 설치된 브라우저 사용: "chrome"(Google Chrome) 또는 "msedge"(Microsoft Edge)
BROWSER_CHANNEL = "chrome"

# 로그인 상태(쿠키 등)는 이 폴더에 저장 — 로그인 정보와 같으니 공유하거나 올리지 말 것.
# 일부러 프로젝트 폴더 밖 이 컴퓨터 폴더에 둠(Windows는 %LOCALAPPDATA%\AmazonAIProject\browser_profile),
# 프로젝트가 OneDrive에 있어서: 동기화되면 쿠키가 클라우드에 올라가고, Chrome 실행 중 설정이 망가질 수 있음.
# 위치를 바꾸려면 환경 변수 SC_PROFILE_DIR 설정.
import platform_util as _pu
_LOCAL_BASE = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / ".local" / "share"))
# macOS: ~/Library/Application Support/AmazonAIProject/browser_profile
PROFILE_DIR = Path(os.environ.get("SC_PROFILE_DIR") or ((_pu.local_dir() if _pu.IS_MAC else _LOCAL_BASE / "AmazonAIProject") / "browser_profile"))

# 주소에 이 조각이 있으면 로그인 / 2단계 인증 / 보안 확인 페이지에 머물러 있다는 뜻
LOGIN_MARKERS = ("/ap/signin", "/ap/mfa", "/ap/cvf", "/ap/challenge", "/ap/forgotpassword")
# 매장이나 사이트가 여럿이면 Amazon이 먼저 계정을 고르게 함
SWITCHER_MARKERS = ("account-switcher", "/merchant-picker")


class SessionExpired(RuntimeError):
    """저장된 로그인이 만료됨, login을 다시 실행해야 함."""


class BrowserNotFound(RuntimeError):
    """이 컴퓨터에서 지정한 브라우저를 찾지 못함."""


def _launch(p: Playwright, headless: bool) -> BrowserContext:
    """이 컴퓨터에 설치된 브라우저 실행(Playwright 내장 Chromium은 쓰지 않음)."""
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        return p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            channel=BROWSER_CHANNEL,
            headless=headless,
            no_viewport=True,          # 웹 페이지 크기는 브라우저 창 크기를 따라감
            args=[] if headless else ["--start-maximized"],   # 매번 최대화해서 열고 지난번 창 크기를 쓰지 않음
        )
    except Exception as e:
        msg = str(e)
        if "Executable doesn't exist" in msg or "is not found" in msg or "not installed" in msg.lower():
            name = "Google Chrome" if BROWSER_CHANNEL == "chrome" else "Microsoft Edge"
            raise BrowserNotFound(
                f"이 PC에서 다음 브라우저를 찾지 못했습니다: {name}. 먼저 설치하거나, "
                f"sc_browser.py의 BROWSER_CHANNEL을 PC에 있는 브라우저로 바꿔 주세요. "
            ) from e
        raise


def page_state(page: Page) -> str:
    """지금 페이지 상태 판단: ok / need_login / need_account / unknown"""
    url = page.url
    if any(m in url for m in LOGIN_MARKERS):
        return "need_login"
    if any(m in url for m in SWITCHER_MARKERS):
        return "need_account"
    if "sellercentral.amazon.com" in url:
        return "ok"
    return "unknown"


def _first_page(ctx: BrowserContext) -> Page:
    return ctx.pages[0] if ctx.pages else ctx.new_page()


def human_pause(min_s: float = 2.0, max_s: float = 5.0) -> None:
    """2~5초 무작위 멈춤(밀리초 단위). 사람처럼 조작하려면 human.py 사용."""
    human.pause(min_s, max_s)


def login(timeout_min: int = 10) -> bool:
    """브라우저를 열고 직접 로그인할 때까지 기다렸다가, 로그인되면 상태를 자동 저장."""
    hints = {
        "need_login": "열린 브라우저에서 Seller Central에 로그인하세요(2단계 인증 포함). ",
        "need_account": "브라우저에서 사용할 스토어 / 마켓플레이스(United States)를 선택하세요. ",
        "unknown": "Seller Central로 이동하기를 기다리는 중...",
    }
    with sync_playwright() as p:
        ctx = _launch(p, headless=False)
        page = _first_page(ctx)
        page.goto(SC_HOME, wait_until="domcontentloaded")

        deadline = time.time() + timeout_min * 60
        last_state = None
        while time.time() < deadline:
            # 새 탭에서 작업할 수 있어서 모든 탭을 확인
            states = [page_state(pg) for pg in ctx.pages] or ["unknown"]
            if "ok" in states:
                # 잠시 기다려 로그인 페이지로 다시 넘어가지 않는지 확인
                page.wait_for_timeout(3000)
                if any(page_state(pg) == "ok" for pg in ctx.pages):
                    print(f"✅ 로그인 성공, 상태를 저장했습니다: {PROFILE_DIR}")
                    ctx.close()
                    return True
            state = states[0]
            if state != last_state:
                print(hints.get(state, hints["unknown"]))
                last_state = state
            page.wait_for_timeout(2000)

        print(f"⏱ {timeout_min}분 안에 로그인 완료를 확인하지 못해 종료합니다. login을 다시 실행할 수 있습니다. ")
        ctx.close()
        return False


@contextmanager
def open_session(headless: bool = False):
    """
    저장된 세션을 열어 Seller Central 홈으로 들어가고, 바로 쓸 수 있는 page를 돌려줌.
    로그인이 만료됐으면 SessionExpired를 냄 — 이때 `python sc_browser.py login`을 다시 실행.
    """
    with sync_playwright() as p:
        ctx = _launch(p, headless=headless)
        try:
            page = _first_page(ctx)
            page.goto(SC_HOME, wait_until="domcontentloaded")
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass  # 관리 화면은 긴 연결이 많아 networkidle을 못 기다려도 괜찮음
            state = page_state(page)
            if state != "ok":
                raise SessionExpired(
                    f"현재 상태: {state}({page.url}). `python sc_browser.py login`을 실행해 다시 로그인하세요. "
                )
            human.pause()  # 홈이 열리면 먼저 ‘한 번 보고’ 조작 시작
            yield page
        finally:
            ctx.close()


def check(headless: bool) -> bool:
    try:
        with open_session(headless=headless) as page:
            print("✅ 세션 유효")
            print("   페이지 제목:", page.title())
            print("   현재 주소:", page.url)
            return True
    except SessionExpired as e:
        print("❌ 세션 무효:", e)
        return False


def main() -> None:
    parser = argparse.ArgumentParser(description="Seller Central 로컬 세션 관리")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_login = sub.add_parser("login", help="직접 로그인하고 상태 저장")
    p_login.add_argument("--timeout", type=int, default=10, help="로그인 대기 시간(분, 기본 10)")

    p_check = sub.add_parser("check", help="저장된 로그인 상태 확인")
    p_check.add_argument("--headless", action="store_true", help="브라우저 창을 표시하지 않음")

    args = parser.parse_args()
    try:
        ok = login(args.timeout) if args.cmd == "login" else check(args.headless)
    except BrowserNotFound as e:
        print("❌", e)
        ok = False
    except PlaywrightError as e:
        print("❌ Seller Central을 열지 못했습니다. 네트워크 연결을 확인한 후 다시 시도하세요. ")
        print("   상세 정보:", str(e).splitlines()[0])
        ok = False
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

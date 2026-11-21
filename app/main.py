"""
SellerMorning - 기능 메뉴

실행:
    python app\\main.py
또는 start.bat 더블클릭

새 기능을 더할 때는 함수를 하나 만들어 아래 FEATURES 목록에 넣으면 됨.
"""

import xlsafe  # noqa: F401  Excel 수식 주입 방지(Amazon에서 받은 글자가 수식으로 실행되지 않게)
import sys
from dataclasses import dataclass
from typing import Callable, Optional

from playwright.sync_api import Error as PlaywrightError

import sc_browser
from sc_browser import BrowserNotFound, SessionExpired, open_session


@dataclass
class Feature:
    group: str
    name: str
    desc: str
    func: Optional[Callable[[], None]] = None  # None이면 개발 중
    status: Optional[Callable[[], str]] = None  # 메뉴에 추가로 보이는 한 줄(예: 마지막 실행 시각)

    @property
    def ready(self) -> bool:
        return self.func is not None


# ---------------------------------------------------------------- 만들어진 기능

def do_login() -> None:
    sc_browser.login()


def do_check() -> None:
    sc_browser.check(headless=True)


def do_open_backend() -> None:
    """저장된 로그인으로 관리 화면을 열어 직접 작업할 수 있게. 브라우저 창을 닫으면 메뉴로 돌아옴."""
    with open_session(headless=False) as page:
        print("✅ Seller Central을 열었습니다. 브라우저 창을 닫으면 메뉴로 돌아갑니다. ")
        page.wait_for_event("close", timeout=0)


def do_account_health() -> None:
    import account_health
    account_health.run()


def do_orders() -> None:
    import orders
    orders.run()


def _orders_status() -> str:
    import orders
    return orders.last_run_text()


def do_business() -> None:
    import business
    business.run()


def _business_status() -> str:
    import business
    return business.last_run_text()


def do_ads() -> None:
    import ads_agent
    ads_agent.run()


def do_ads_analyze() -> None:
    import ads_agent
    ads_agent.run_analyze_only()


def _ads_status() -> str:
    import ads_agent
    return ads_agent.status_text()


def do_explore() -> None:
    import explore
    explore.run()


def do_listing_check() -> None:
    import listing
    listing.run()


def _listing_status() -> str:
    import listing
    return listing.last_run_text()


def do_daily() -> None:
    import daily
    daily.menu()


def _daily_status() -> str:
    import daily
    return daily.status_text()


def do_listings() -> None:
    import inventory
    inventory.run_listings()


def do_fba() -> None:
    import inventory
    inventory.run_fba()


# ---------------------------------------------------------------- 기능 목록
# 순서가 메뉴 번호 순서. func=None이면 ‘개발 중’으로 표시.

FEATURES = [
    Feature("계정", "Seller Central 로그인", "처음 사용하거나 로그인이 만료되었을 때, 직접 로그인하고 상태를 저장합니다", do_login),
    Feature("계정", "로그인 상태 확인", "저장된 로그인이 아직 유효한지 확인합니다", do_check),
    Feature("계정", "셀러 센트럴 열기", "저장된 로그인으로 Seller Central을 열어 직접 조작합니다", do_open_backend),
    Feature("계정", "계정 상태", "계정 상태 점수, 이슈 상세, 영향 받는 상품을 확인하고 Excel 보고서로 내보냅니다", do_account_health),

    Feature("데이터", "전체 상품 재고", "All Listings Report(하루 한 번 가져옴), Excel로 변환하고 상품 상태를 평가합니다", do_listings),
    Feature("데이터", "FBA 재고", "공급망 보고서의 FBA 재고 보고서(하루 한 번 가져옴), Excel로 변환하고 재입고·장기 미판매·재고 기간을 평가합니다", do_fba),
    Feature("데이터", "주문 보고서", "전날 / 최근 7일 / 14일(이전 기간과 비교) / 30일(미국 중부 시간), Excel로 내보냅니다",
            do_orders, status=_orders_status),
    Feature("데이터", "비즈니스 보고서", "마케팅 워크스페이스 + 하위 ASIN별 트래픽·전환율·바이박스(전날 / 7 / 14 / 30일, 이전 기간과 비교), Excel로 내보냅니다",
            do_business, status=_business_status),

    Feature("광고", "광고 데이터 수집·조치 제안",
            "캠페인·키워드·검색어(전날 / 최근 7일) + Brand Analytics 키워드 → 입찰·부정 키워드·매치·상품명 조치 제안(수정 안 함)",
            do_ads, status=_ads_status),
    Feature("광고", "광고 분석만 다시 하기", "이미 받은 자료로 분석만 다시 합니다(ads_config.xlsx 수정 후)", do_ads_analyze),

    Feature("Listing", "Listing 점검",
            "제목 75자·금지 문자·반복 단어, 불릿·검색어·Item Highlights, 이미지·A+·평점, 판매 불가·미완성 상품 점검 → 수정 제안(수정 안 함)",
            do_listing_check, status=_listing_status),


    Feature("도구", "매일 자동 실행 · 오늘의 요약",
            "선택한 기능을 순서대로 실행하고 P0·P1 할 일을 한 장으로 정리합니다. Windows 작업 스케줄러로 매일 자동 실행 예약",
            do_daily, status=_daily_status),
]


# ---------------------------------------------------------------- 메뉴

def print_menu() -> None:
    print()
    print("=" * 52)
    print(f"      SellerMorning  {__import__('version').text()}")
    print("=" * 52)
    group = None
    for i, f in enumerate(FEATURES, start=1):
        if f.group != group:
            group = f.group
            print(f"\n  【{group}】")
        tag = "" if f.ready else "  (개발 중)"
        print(f"   {i:>2}. {f.name}{tag}")
        print(f"       {f.desc}")
        if f.status:
            try:
                print(f"       {f.status()}")
            except Exception:
                pass
    print("\n    0. 종료")
    print("-" * 52)


def run_feature(f: Feature) -> None:
    if not f.ready:
        print(f"\n“{f.name}” 기능은 아직 개발 중입니다. 조금만 기다려 주세요. ")
        return
    print(f"\n▶ {f.name}\n")
    try:
        import runlock
        with runlock.hold(f.name):
            f.func()
    except SessionExpired as e:
        print("❌ 로그인이 만료되었습니다:", e)
        print("   먼저 ‘Seller Central 로그인’을 선택해 주세요. ")
    except BrowserNotFound as e:
        print("❌", e)
    except PlaywrightError as e:
        print("❌ 브라우저 조작에 실패했습니다. 네트워크 연결을 확인하거나, 다른 스크립트가 브라우저를 사용 중인지 확인하세요. ")
        print("   상세 정보:", str(e).splitlines()[0])
    except ModuleNotFoundError as e:
        print(f"❌ 필요한 라이브러리가 없습니다: {e.name}. 프로젝트 폴더에서 다음을 실행하세요: ")
        print("   pip install -r requirements.txt")
    except PermissionError as e:
        print(f"❌ 파일을 저장할 수 없습니다: {e.filename}")
        print("   이 파일이 Excel에서 열려 있다면 닫은 후 다시 시도하세요. ")
    except KeyboardInterrupt:
        print("\n중단했습니다. 메뉴로 돌아갑니다. ")
    except Exception as e:
        print(f"❌ 예기치 않은 오류: {type(e).__name__}: {e}")
        print("   이 메시지를 개발자에게 보내 주세요. ")


def main() -> None:
    try:
        import daily
        daily.fix_bat()
    except Exception:
        pass
    if "--daily" in sys.argv:          # 작업 스케줄러(daily_run.bat)가 부름: 메뉴 없이 다 돌고 끝남
        import daily
        import runlock
        try:
            with runlock.hold("매일 자동 실행", wait_min=240):
                out = daily.run_all(open_summary=False)
            try:
                import notify
                notify.summary_toast(out)
            except Exception:
                pass
        except Exception as e:
            print(f"❌ 매일 자동 실행 오류: {type(e).__name__}: {e}")
        return
    while True:
        print_menu()
        try:
            choice = input("번호를 입력하고 Enter를 누르세요: ").strip()
        except (KeyboardInterrupt, EOFError):
            print()
            break

        if choice in ("0", "q", "Q"):
            break
        if not choice.isdigit() or not 1 <= int(choice) <= len(FEATURES):
            print(f"\n0부터 {len(FEATURES)} 사이의 숫자를 입력하세요. ")
            continue

        run_feature(FEATURES[int(choice) - 1])
        try:
            input("\nEnter를 누르면 메뉴로 돌아갑니다...")
        except (KeyboardInterrupt, EOFError):
            print()
            break

    print("안녕히 가세요!")


if __name__ == "__main__":
    main()

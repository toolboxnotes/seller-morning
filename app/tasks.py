"""
기능 목록(화면 프로그램 gui.py와 실행기 runner.py가 함께 사용)

각 기능: 키, 분류, 이름, 설명, 결과 폴더, 실행 함수(옵션, open_report) → 결과 파일 경로, 상태 함수
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional

BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS = BASE_DIR / "reports"

PERIODS = [("1", "전날"), ("2", "최근 7일"), ("3", "최근 14일"), ("4", "최근 30일")]


@dataclass
class Task:
    key: str
    group: str
    name: str
    desc: str
    folder: str = ""                       # reports/ 아래 폴더
    run: Optional[Callable] = None          # run(opts: dict, open_report: bool) -> 경로 또는 None
    status: Optional[Callable[[], str]] = None
    options: List[str] = field(default_factory=list)   # "period"
    long: bool = False                      # 오래 걸리는 기능(안내용)

    @property
    def ready(self) -> bool:
        return self.run is not None

    @property
    def folder_path(self) -> Path:
        return REPORTS / self.folder if self.folder else REPORTS


# ---------------------------------------------------------------- 실행 함수

def _login(o, open_report):
    import sc_browser
    sc_browser.login()


def _check(o, open_report):
    import sc_browser
    sc_browser.check(headless=True)


def _open_sc(o, open_report):
    from sc_browser import open_session
    with open_session(headless=False) as page:
        print("✅ Seller Central을 열었습니다. 브라우저 창을 닫으면 끝납니다.")
        page.wait_for_event("close", timeout=0)


def _account_health(o, open_report):
    import account_health
    return account_health.run(open_report=open_report)


def _listings(o, open_report):
    import inventory
    return inventory.run_listings(open_report=open_report)


def _fba(o, open_report):
    import inventory
    return inventory.run_fba(open_report=open_report)


def _orders(o, open_report):
    import orders
    return orders.run(period=o.get("period", "2"), open_report=open_report)


def _business(o, open_report):
    import business
    return business.run(period=o.get("period", "2"), open_report=open_report)


def _explore(o, open_report):
    import explore
    out = explore.run(open_report=False)
    return (out / "explore_summary.xlsx") if out else None


def _ads(o, open_report):
    import ads_agent
    return ads_agent.run(open_report=open_report)


def _ads_analyze(o, open_report):
    import ads_agent
    return ads_agent.run_analyze_only(open_report=open_report)


def _listing_check(o, open_report):
    import listing
    return listing.run(open_report=open_report)


def _daily_all(o, open_report):
    import daily
    return daily.run_all(open_summary=open_report)


def _daily_summary(o, open_report):
    import daily
    return daily.build_summary()


# ---------------------------------------------------------------- 상태 함수

def _st(mod: str, fn: str = "last_run_text"):
    def f():
        import importlib
        return getattr(importlib.import_module(mod), fn)()
    return f


TASKS: List[Task] = [
    Task("login", "계정", "Seller Central 로그인", "처음 사용하거나 로그인이 만료되었을 때, 열린 브라우저에서 직접 로그인하면 상태를 저장합니다(최대 10분 대기).",
         "", _login),
    Task("check", "계정", "로그인 상태 확인", "저장된 로그인이 아직 유효한지 확인합니다(창 없이).", "", _check),
    Task("open_sc", "계정", "셀러 센트럴 열기", "저장된 로그인으로 Seller Central을 열어 직접 조작합니다. 브라우저 창을 닫으면 끝납니다.", "", _open_sc),
    Task("account_health", "계정", "계정 상태", "계정 상태 점수·이슈 상세·영향 받는 상품 + 클레임(차지백·A-to-z·SAFE-T)·고객의 소리(VoC).",
         "account_health", _account_health),

    Task("listings", "데이터", "전체 상품 재고", "All Listings Report(하루 한 번)를 Excel로 변환하고 상품 상태를 평가합니다.", "inventory", _listings),
    Task("fba", "데이터", "FBA 재고", "FBA 재고 보고서 + FBA 대시보드(IPI·보관 기간 추가요금·장기 재고)·입고 배송(수량 부족·지연).",
         "inventory", _fba),
    Task("finance", "데이터", "재무·보상 점검(SP-API)",
         "창고 분실·파손인데 보상받지 못한 재고, 환불했는데 반품·보상이 없는 주문, FBA 수수료·크기 등급 변화를 찾아 신청 방법과 함께 정리합니다. "
         "조회만(신청은 직접), 보고서 4개라 5~15분. 매일 자동 실행에서는 주 1회(월요일).",
         "finance", lambda o, op: __import__("finance").run(open_report=op), _st("finance", "status_text")),
    Task("orders", "데이터", "주문 보고서", "주문·판매 수량·매출(이전 기간과 비교, 미국 중부 시간) + 반품 분석(최근 90일).",
         "orders", _orders, _st("orders"), ["period"]),
    Task("business", "데이터", "비즈니스 보고서", "마케팅 워크스페이스 + 하위 ASIN별 트래픽·전환율·바이박스(이전 기간과 비교).",
         "business", _business, _st("business"), ["period"]),

    Task("ads", "광고", "광고 데이터 수집·조치 제안", "캠페인·키워드·검색어(전날/7일) + Brand Analytics + 브랜드 노출 점유율 → 입찰·부정 키워드·매치·상품명 제안(수정 안 함). 광고 보고서는 Amazon Ads API 우선(설정 시), 안 되면 광고 콘솔 화면. 최대 약 1시간.",
         "ads", _ads, _st("ads_agent", "status_text"), long=True),
    Task("ads_analyze", "광고", "광고 분석만 다시 하기", "이미 받은 자료로 분석만 다시 합니다(ads_config.xlsx 수정 후). 브라우저를 열지 않습니다.", "ads", _ads_analyze),

    Task("buybox", "Listing", "Buy Box·합승 판매자 확인(SP-API)",
         "판매 중인 모든 ASIN의 Buy Box를 누가 가졌는지, 다른 판매자(합승)가 있는지, 내 가격이 Amazon 경쟁 가격 기준보다 높은지 확인합니다. "
         "어제 없던 판매자가 나타나면 P1. 조회만(SP-API Pricing 역할 필요, 몇 분).",
         "buybox", lambda o, op: __import__("buybox").run(open_report=op), _st("buybox", "status_text")),
    Task("funnel", "Listing", "주간 ASIN 퍼널 진단",
         "ASIN별 검색 노출 → 클릭 → 장바구니 → 구매를 지난 4주와 비교해 막힌 단계와 조치를 알려 줍니다(Brand Analytics SCP + 비즈니스 보고서, 최근 8주). "
         "‘리스팅·가격 수정’ 전후 2주 효과도 비교. SP-API 권한이 있으면 브라우저 없이 몇 분(처음은 8주치라 10~30분).",
         "funnel", lambda o, op: __import__("funnel").run(open_report=op), _st("funnel", "status_text")),
    Task("listing_check", "Listing", "Listing 점검", "제목 75자·금지 문자·반복 단어, 불릿·검색어·Item Highlights, 이미지·A+·평점, 리스팅 상태(검색 제한 등) → 수정 제안(수정 안 함).",
         "listing", _listing_check, _st("listing"), long=True),


    Task("changes", "수정(확인 후 적용)", "리스팅·가격 수정(SP-API)",
         "요청서(Excel)에 적은 제목·불릿·검색어·가격 변경을 Amazon 검사(실제 변경 없음) → 고른 줄만 적용. 원래 값 기록·되돌리기 가능. "
         "Listing 점검의 ‘제목 초안’을 요청서에 미리 채웁니다. 실행하면 수정 창이 열립니다.",
         "changes", lambda o, op: print("이 기능은 화면 프로그램(start_gui.bat · Mac: start_gui.command)의 수정 창에서 사용합니다.")),

    Task("action_items", "도구", "Amazon 할 일 확인(SP-API)", "Amazon이 알려 주는 문제·할 일·추천(Action Items)을 가져와 등급별로 정리합니다(API 전용, 브라우저 안 열림).",
         "daily", lambda o, op: __import__("daily").run_action_items(op)),
    Task("daily_all", "도구", "매일 실행(전체) · 오늘의 요약", "설정한 기능을 순서대로 모두 실행하고 P0·P1 할 일을 ‘오늘의 요약’ 한 장으로 정리합니다. 1~2시간.",
         "daily", _daily_all, _st("daily", "status_text"), long=True),
    Task("daily_summary", "도구", "오늘의 요약만 다시 만들기", "오늘 이미 실행한 기능들의 결과로 요약을 다시 만듭니다(브라우저 안 열림).", "daily", _daily_summary),
]

BY_KEY: Dict[str, Task] = {t.key: t for t in TASKS}

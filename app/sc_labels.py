"""
Seller Central 화면 이름 사전(영어 · 한국어 · 중국어)

Seller Central 화면 언어(오른쪽 위 언어 설정)가 영어·한국어·중국어 어느 것이어도 메뉴·버튼을 찾을 수 있게
모든 이름을 세 언어로 기록합니다. 출처: 2026-11-04 사이트 맵(Site map / 사이트 맵 / 站点地图) 화면.
새 이름이 생기면 SITEMAP 또는 EXTRA에 추가하세요.
페이지 안 버튼·탭·선택지 이름(약 600개)은 sc_labels_pages.json(2026-11-04 세 언어 화면에서 수집).

사용:  sc_labels.L("Supply chain reports")  →  ["Supply chain reports", "공급망 보고서", "供应链报告"]
"""

import json
from pathlib import Path
from typing import Dict, List, Optional

LANGS = ("en", "ko", "zh")
LANG_NAME = {"en": "영어", "ko": "한국어", "zh": "중국어"}

# (분류, 항목 영어, 한국어, 중국어, 상위 항목 영어 또는 "")
SITEMAP: List[tuple] = [
    # ---- Products · 상품 · 商品
    ("Products", "Products", "상품", "商品", ""),
    ("Products", "Manage products", "상품 관리", "管理商品", ""),
    ("Products", "Manage A+ content", "A+ 콘텐츠 관리", "管理 A+ 商品描述", "Manage products"),
    ("Products", "Manage images", "이미지 관리", "管理图片", "Manage products"),
    ("Products", "Manage videos", "동영상 관리", "管理视频", "Manage products"),
    ("Products", "Manage experiments", "실험 관리", "管理试验", "Manage products"),
    ("Products", "View selling applications", "판매 신청 보기", "查看销售申请", "Manage products"),
    ("Products", "Manage automotive fitment data", "자동차 부품 데이터 관리", "管理汽车配件数据", ""),
    ("Products", "Manage pricing", "가격 관리", "管理定价", ""),
    ("Products", "Automate pricing", "자동 가격 책정", "自动定价", "Manage pricing"),
    ("Products", "Flexible financing", "유연한 금융 서비스", "灵活融资", "Manage pricing"),
    ("Products", "Add products", "상품 등록", "添加商品", ""),
    ("Products", "Analyze reviews", "리뷰 분석", "分析评论", ""),
    ("Products", "Feedback manager", "피드백 관리자", "反馈管理器", "Analyze reviews"),
    ("Products", "Customer reviews", "고객 리뷰", "买家评论", "Analyze reviews"),
    ("Products", "Voice of customer", "고객의 소리", "买家之声", "Analyze reviews"),
    ("Products", "Research products", "연구 상품", "研究商品", ""),
    ("Products", "Product opportunity explorer", "상품 기회 탐색기", "商品商机探测器", "Research products"),
    ("Products", "Marketplace product guidance", "마켓플레이스 상품 지침", "商城商品指南", "Research products"),
    ("Products", "Sell globally", "글로벌 셀링", "全球开店", ""),
    ("Products", "Product reports", "상품 보고서", "商品报告", ""),
    ("Products", "Feedback report", "피드백 보고서", "反馈报告", "Product reports"),
    # ---- Supply chain · 공급망 · 供应链
    ("Supply chain", "Supply chain", "공급망", "供应链", ""),
    ("Supply chain", "Manage Amazon inventory", "아마존 재고 관리", "管理亚马逊库存", ""),
    ("Supply chain", "Manage AWD inventory", "AWD 재고 관리", "管理 AWD 库存", "Manage Amazon inventory"),
    ("Supply chain", "Manage stranded inventory", "문제 있는 재고 관리", "管理无在售信息的亚马逊库存", "Manage Amazon inventory"),
    ("Supply chain", "Manage unfulfillable inventory", "주문 처리 불가 재고 관리", "管理不可售库存", "Manage Amazon inventory"),
    ("Supply chain", "Manage global store", "글로벌 스토어 관리", "管理海外购", ""),
    ("Supply chain", "Manage seller fulfilled inventory", "셀러 주문 처리 재고 관리", "管理卖家自配送库存", ""),
    ("Supply chain", "Manage shipments", "배송 관리", "管理货件", ""),
    ("Supply chain", "Send to Amazon", "아마존으로 발송", "Send to Amazon", ""),
    ("Supply chain", "Send to AWD", "AWD로 발송", "发送至 AWD", "Send to Amazon"),
    ("Supply chain", "Move from AWD", "AWD에서 이동", "从 AWD 移出", "Send to Amazon"),
    ("Supply chain", "Manage channel destinations", "채널 대상 관리", "管理渠道目的地", ""),
    ("Supply chain", "Supply chain console", "공급망 콘솔", "供应链控制面板", ""),
    ("Supply chain", "Supply chain reports", "공급망 보고서", "供应链报告", ""),
    ("Supply chain", "Seller fulfilled reports", "셀러 주문 처리 보고서", "卖家自配送报告", "Supply chain reports"),
    ("Supply chain", "FBA inventory reports", "FBA 재고 보고서", "亚马逊物流库存报告", "Supply chain reports"),
    # ---- Orders · 주문 · 订单
    ("Orders", "Orders", "주문", "订单", ""),
    ("Orders", "Manage orders", "주문 관리", "管理订单", ""),
    ("Orders", "Create MCF order", "MCF 주문 생성", "创建多渠道配送订单", "Manage orders"),
    ("Orders", "Upload order files", "주문 파일 업로드", "上传订单文件", "Manage orders"),
    ("Orders", "Manage quotes", "견적 관리", "管理报价", ""),
    ("Orders", "Manage returns", "반품 관리", "管理退货", ""),
    ("Orders", "Returns and recovery insights", "반품 및 회수 인사이트", "退货与恢复洞察信息", "Manage returns"),
    ("Orders", "Manage FBA returns", "FBA 반품 관리", "管理亚马逊物流退货", "Manage returns"),
    ("Orders", "Manage claims", "클레임 관리", "管理索赔", ""),
    ("Orders", "Manage A-Z claims", "A-Z 클레임 관리", "管理亚马逊商城交易保障索赔", "Manage claims"),
    ("Orders", "Manage Safe-T claims", "Safe-T 클레임 관리", "管理Safe-T索赔", "Manage claims"),
    ("Orders", "Order reports", "주문 보고서", "订单报告", ""),
    ("Orders", "Seller fulfilled return reports", "셀러 주문 처리 반품 보고서", "卖家自配送退货报告", "Order reports"),
    ("Orders", "FBA customer return reports", "FBA 고객 반품 보고서", "亚马逊物流买家退货报告", "Order reports"),
    ("Orders", "Replacement reports", "교환 보고서", "换货报告", "Order reports"),
    # ---- Marketing · 마케팅 · 营销推广
    ("Marketing", "Marketing", "마케팅", "营销推广", ""),
    ("Marketing", "Manage promotions and deals", "프로모션 및 딜 관리", "管理促销和特惠", ""),
    ("Marketing", "Manage promotions", "프로모션 관리", "管理促销", "Manage promotions and deals"),
    ("Marketing", "Manage brand tailored promotions", "브랜드 맞춤형 프로모션 관리", "管理品牌定制促销", "Manage promotions and deals"),
    ("Marketing", "Manage stores", "스토어 관리", "管理商城", ""),
    ("Marketing", "Brand analytics", "브랜드 분석", "品牌分析", ""),
    ("Marketing", "Ads console", "광고 콘솔", "广告控制台", ""),
    ("Marketing", "Vine", "Vine", "Vine", ""),
    # ---- Finance · 재무 · 财务人员
    ("Finance", "Finance", "재무", "财务人员", ""),
    ("Finance", "Manage transactions", "거래 관리", "管理交易", ""),
    ("Finance", "Manage payouts", "지급액 관리", "管理付款", "Manage transactions"),
    ("Finance", "Seller wallet", "Seller Wallet", "卖家钱包", "Manage transactions"),
    ("Finance", "Currency converter", "통화 정산 서비스", "货币转换器", "Manage transactions"),
    ("Finance", "Manage taxes", "세금 관리", "管理税费", ""),
    ("Finance", "Manage VAT", "VAT 관리", "管理增值税", "Manage taxes"),
    ("Finance", "Profit analytics", "수익 분석", "利润分析", ""),
    ("Finance", "Revenue calculator", "FBA 매출 계산기", "亚马逊物流收入计算器", "Profit analytics"),
    ("Finance", "Selling economics", "판매 경제성", "销售成本", "Profit analytics"),
    ("Finance", "Finance reports", "재무 보고서", "财务报告", ""),
    ("Finance", "Payment reports", "판매 대금 인출 보고서", "付款报告", "Finance reports"),
    ("Finance", "Tax document library", "세금 자료 라이브러리", "税务文件库", "Finance reports"),
    # ---- Account health · 어카운트 헬스 · 账户状况
    ("Account health", "Account health", "어카운트 헬스", "账户状况", ""),
    ("Account health", "Manage account health", "어카운트 헬스 관리", "管理账户状况", ""),
    ("Account health", "Performance notifications", "퍼포먼스 알림", "绩效通知", "Manage account health"),
    ("Account health", "Manage compliance", "규정 준수 관리", "管理合规性", "Manage account health"),
    # ---- Settings · 설정 · 设置
    ("Settings", "Settings", "설정", "设置", ""),
    ("Settings", "Manage permissions", "권한 관리", "管理权限", ""),
    ("Settings", "Account information", "계정 정보", "账户信息", ""),
    ("Settings", "Account settings", "계정 설정", "账户设置", ""),
    ("Settings", "Notification preferences", "알림 기본 설정", "通知首选项", ""),
    ("Settings", "Login settings", "로그인 설정", "登录设置", ""),
    ("Settings", "Return settings", "반품 설정", "退货设置", ""),
    ("Settings", "Gift options", "선물 옵션", "礼品选项", ""),
    ("Settings", "Shipping settings", "배송 설정", "配送设置", ""),
    ("Settings", "Tax settings", "세금 설정", "税务设置", ""),
    ("Settings", "FBA settings", "FBA 설정", "亚马逊物流设置", ""),
    ("Settings", "Supply chain settings", "공급망 설정", "供应链设置", ""),
    ("Settings", "Premium services settings", "프리미엄 서비스 설정", "高级服务设置", ""),
    ("Settings", "Manage your apps", "앱 관리", "管理您的应用", ""),
    ("Settings", "Manage your certifications", "인증 관리", "管理您的认证", ""),
    ("Settings", "Manage your brands", "브랜드 관리", "管理您的品牌", ""),
    ("Settings", "Multichannel settings", "Multichannel settings", "Multichannel settings", ""),
    # ---- Programs and apps · 프로그램 및 앱 · 计划和应用程序
    ("Programs and apps", "Programs and apps", "프로그램 및 앱", "计划和应用程序", ""),
    ("Programs and apps", "Explore programs", "프로그램 탐색", "了解计划", ""),
    ("Programs and apps", "Selling partner appstore", "셀링 파트너 앱스토어", "销售伙伴应用商店", ""),
    ("Programs and apps", "Explore services", "서비스 탐색", "浏览服务", ""),
    # ---- Support resources · 지원 리소스 · 支持资源
    ("Support resources", "Support resources", "지원 리소스", "支持资源", ""),
    ("Support resources", "Get help and resources", "지원 및 리소스 요청", "获取帮助和资源", ""),
    ("Support resources", "Manage support cases", "지원 케이스 관리", "管理支持问题单", ""),
    ("Support resources", "Strategic account services", "전략 계정 서비스", "战略账户服务", ""),
    ("Support resources", "Seller university", "셀러 유니버시티", "卖家大学", ""),
    # ---- Other resources · 기타 리소스 · 其他资源
    ("Other resources", "Other resources", "기타 리소스", "其他资源", ""),
    ("Other resources", "Business reports", "비즈니스 보고서", "业务报告", ""),
    ("Other resources", "B2B central", "B2B 센트럴", "B2B 平台", ""),
    ("Other resources", "Custom analytics", "맞춤형 분석", "定制分析", ""),
    ("Other resources", "Virtual bundles", "가상 번들", "虚拟捆绑商品", ""),
    ("Other resources", "Brand transparency", "브랜드 Transparency", "品牌透明度", ""),
    ("Other resources", "Buyer seller messaging", "구매자-셀러 메시지 서비스", "买家与卖家消息服务", ""),
]

# 화면 위쪽 막대 · 기타(사이트 맵 밖)
HEADER: List[tuple] = [
    ("Header", "Menu", "메뉴", "菜单", ""),
    ("Header", "My business", "내 비즈니스", "我的业务", ""),          # 메뉴(☰) 첫 분류
    ("Header", "Workspace", "작업장", "工作区", ""),                   # 메뉴 각 분류 아래 항목
    ("Header", "Site map", "사이트 맵", "站点地图", ""),
    ("Header", "Add products", "상품 등록", "添加商品", ""),
    ("Header", "Seller Assistant", "Seller Assistant", "卖家助手", ""),
    ("Header", "Search for anything", "무엇이나 검색", "搜索任何内容", ""),
    ("Header", "Feedback", "피드", "反馈", ""),
    ("Header", "Program Policies", "프로그램 정책", "计划政策", ""),
]

# 예전 표기·다른 표기(대소문자 차이, 예전 화면 이름 등) — 찾을 때 함께 시도
EXTRA: Dict[str, List[str]] = {
    "Marketing": ["营销"],                                  # 예전 중국어 화면
    "Brand analytics": ["Brand Analytics"],
    "Ads console": ["Ads Console"],
    "Account health": ["Account Health", "账户健康"],
    "Business reports": ["Business Reports"],
    "FBA inventory reports": ["FBA Inventory Reports"],
}

_ALL = SITEMAP + HEADER
_BY_EN = {r[1]: r for r in _ALL}

# 페이지 안 버튼·탭·링크·선택지 이름(영어 → [한국어, 중국어])
# 2026-11-04 화면 이름 수집(영·한·중 전환해 같은 위치끼리 짝지음) 결과. 약 600개.
try:
    PAGES: Dict[str, List[str]] = json.loads((Path(__file__).with_name("sc_labels_pages.json")).read_text(encoding="utf-8"))
except Exception:
    PAGES = {}

_ANY: Dict[str, set] = {}      # 어느 언어 이름이든 → 같은 뜻의 세 언어 이름


def _add(names) -> None:
    names = [n for n in names if n]
    for n in names:
        _ANY.setdefault(n, set()).update(names)


for _r in _ALL:
    _add([_r[1], _r[2], _r[3]] + EXTRA.get(_r[1], []))
for _en, (_ko, _zh) in PAGES.items():
    _add([_en, _ko, _zh])


def L(en: str, *more: str) -> List[str]:
    """영어 이름 → [영어, 한국어, 중국어, 다른 표기…] (찾을 때 쓰는 목록, 중복 없음)."""
    r = _BY_EN.get(en)
    pg = PAGES.get(en) or []
    out = [en] + ([r[2], r[3]] if r else []) + list(pg) + EXTRA.get(en, []) + list(more)
    seen, res = set(), []
    for x in out:
        if x and x not in seen:
            seen.add(x)
            res.append(x)
    return res


def expand(labels: List[str]) -> List[str]:
    """찾을 이름 목록에 다른 언어 이름을 덧붙임(원래 순서 먼저). 화면 언어가 무엇이든 찾을 수 있게."""
    out, seen = [], set()
    for x in list(labels) + [y for l in labels for y in sorted(_ANY.get(l, ()))]:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


def name(en: str, lang: str) -> str:
    r = _BY_EN.get(en)
    if not r:
        return en
    return {"en": r[1], "ko": r[2], "zh": r[3]}.get(lang, en)


def all_rows() -> List[tuple]:
    return list(_ALL)


def detect(texts: List[str]) -> Optional[str]:
    """페이지·메뉴 글자 목록으로 화면 언어 추정(사전 이름이 가장 많이 맞는 언어)."""
    s = set(t.strip() for t in texts if t)
    score = {lang: sum(1 for r in _ALL if r[1 + LANGS.index(lang)] in s and r[1] != r[2] != r[3]) for lang in LANGS}
    best = max(score, key=score.get)
    return best if score[best] else None

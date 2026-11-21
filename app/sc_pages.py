"""
Seller Central / 광고 콘솔 페이지 읽기(보기만, 바꾸지 않음)

“전체 메뉴 스캔”에서 찾은, 기존 기능이 아직 쓰지 않는 페이지: 주소로 바로 열어 페이지 글자와 표를 읽고 버튼은 누르지 않음.
페이지마다 하루 한 번 읽음(결과는 reports/daily/날짜/pages/이름.json), 같은 날 다시 쓰면 저장된 결과를 읽음.

  home            홈: 오늘 판매, 판매 중 상품, IPI, 처리 대기 주문, 순수익, 광고(최근 7일)
  fba_dashboard   FBA 대시보드: IPI와 영향 요인, 보관 공간 사용 비율(보관 추가 요금), 재고 기간 추가 요금, 입고 문제
  shipments       입고 배송: 배송마다 받은 수 / 보낸 수(차이)
  returns         반품·회수 분석(최근 90일): 반품률, 환불, 반품 사유(ASIN별)
  listing_status  상품 워크스페이스: Listing 상태(검색 제한, 상세 페이지 삭제, 품절…), 고객 건강도, 피드백 점수
  claims          클레임: Chargeback, A-to-z, SAFE-T
  voc             고객의 소리(Voice of the Customer): CX 건강도, NCX 비율(ASIN별)
  ads_bis         광고: 브랜드 키워드 검색 상단 노출 점유율(Brand impression share)
  ads_recs        광고: 추천 / 처리할 일
"""

import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

import human
import inventory as inv
import sc_nav

BASE_DIR = Path(__file__).resolve().parent.parent
DAILY_DIR = BASE_DIR / "reports" / "daily"
DEBUG_DIR = DAILY_DIR / "debug"
ADS = "https://advertising.amazon.com"

URLS = {
    "home": "/home",
    "fba_dashboard": "/fba/dashboard/",
    "shipments": "/manage/shipments",
    "returns": "/manage/rr-central/return-insights?view=returns&~dateRange=T90&~marketplaceId=1&~fulfillment=FBA%2CMFN&~category=All",
    "listing_status": "/amazonsell/listings",
    "chargebacks": "/gp/chargebacks/home.html#/",
    "atoz": "/gp/guarantee-claims/home.html#/",
    "safet": "/safet-claims",
    "voc": "/voice-of-the-customer",
}
NAMES = {"home": "홈 화면", "fba_dashboard": "FBA 대시보드", "shipments": "입고 배송", "returns": "반품 분석",
         "listing_status": "상품 워크스페이스", "claims": "클레임", "voc": "고객의 소리(VoC)",
         "ads_bis": "브랜드 노출 점유율", "ads_recs": "광고 추천"}

_TABLES_JS = r"""
() => {
  const txt = (e) => (e.innerText || e.textContent || '').replace(/\s+/g, ' ').trim();
  const out = [];
  for (const t of document.querySelectorAll('table')) {
    const body = Array.from(t.querySelectorAll('tbody tr')).filter(r => r.querySelector('td'));
    let heads = [];
    const hr = t.querySelectorAll('thead tr');
    if (hr.length) heads = Array.from(hr[hr.length - 1].querySelectorAll('th')).map(txt);
    out.push({headers: heads, rows: body.map(r => Array.from(r.querySelectorAll('td')).map(txt))});
  }
  return out;
}
"""


# ================================================================ 작은 도구

def day_dir(d: Optional[date] = None) -> Path:
    p = DAILY_DIR / (d or date.today()).isoformat()
    p.mkdir(parents=True, exist_ok=True)
    return p


def _num(s) -> Optional[float]:
    if s is None:
        return None
    m = re.search(r"[-−]?\$?\s*([\d,]+(?:\.\d+)?)", str(s))
    if not m:
        return None
    v = float(m.group(1).replace(",", ""))
    return -v if str(s).strip().startswith(("-", "−")) else v


def _after(text: str, label: str, pat: str = r"([^\n]+)", span: int = 200) -> Optional[str]:
    """label 뒤에서 pat에 맞는 첫 내용."""
    i = text.find(label)
    if i < 0:
        return None
    m = re.search(pat, text[i + len(label): i + len(label) + span])
    return m.group(1).strip() if m else None


def _main_text(text: str) -> str:
    """왼쪽 메뉴 부분 제거(메뉴 글자는 Classic Seller Central 앞에 있음)."""
    i = text.find("Classic Seller Central")
    return text[i:] if i >= 0 else text


def _best_table(tables: List[Dict], must: str = "") -> Dict:
    cands = [t for t in tables if t.get("rows") and (not must or any(must.lower() in h.lower() for h in t.get("headers", [])))]
    if not cands:
        cands = [t for t in tables if t.get("rows")]
    return max(cands, key=lambda t: len(t["rows"])) if cands else {"headers": [], "rows": []}


def _rows(tbl: Dict) -> List[Dict]:
    h = [re.sub(r"\s*(unsorted|sorted ascending|sorted descending)\s*$", "", x).strip() for x in tbl.get("headers", [])]
    out = []
    for r in tbl.get("rows", []):
        if len(h) >= len(r) and h:
            out.append(dict(zip(h, r)))
        else:
            out.append({f"열{i + 1}": v for i, v in enumerate(r)})
    return out


def _dump(page, name: str) -> None:
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        stamp = f"{datetime.now():%Y%m%d_%H%M%S}"
        (DEBUG_DIR / f"{stamp}_{name}.txt").write_text(f"URL: {page.url}\n\n{page.inner_text('body')}", encoding="utf-8")
        page.screenshot(path=str(DEBUG_DIR / f"{stamp}_{name}.png"), full_page=True)
    except Exception:
        pass


# ================================================================ 분석(페이지 글자와 표만 사용, 오프라인 테스트가 쉬움)

def parse_home(text: str) -> Dict:
    t = _main_text(text)
    k = {}
    for key, label, pat in (("오늘 매출", "Today's Global Sales", r"(\$[\d,.]+)"),
                            ("판매 중 상품", "Active", r"(\d[\d,]*\s*/\s*[\d,]+)"),
                            ("재고 성과 지수(IPI)", "Performance Index", r"(\d[\d,]*\s*/\s*[\d,]+)"),
                            ("미처리 주문", "Open Orders", r"(\d[\d,]*)"),
                            ("순수익(Net proceeds)", "Net proceeds", r"(\$[\d,.]+)"),
                            ("오늘 주문 매출", "Ordered product sales", r"(\$[\d,.]+)"),
                            ("오늘 주문 수량", "Units ordered", r"(\d[\d,]*)"),
                            ("광고 매출(7일)", "Attributed sales", r"(\$[\d,.]+)"),
                            ("광고비(7일)", "\nSpend", r"(\$[\d,.]+)"),
                            ("광고 ROAS(7일)", "\nROAS", r"([\d.]+)"),
                            ("광고 클릭(7일)", "Total clicks", r"([\d,]+)"),
                            ("FBA 대기 주문(90일)", "Pending", r"(\d[\d,]*)")):
        v = _after(t, label, pat)
        if v:
            k[key] = v
    return {"kpis": k}


def parse_fba_dashboard(text: str) -> Dict:
    t = _main_text(text)
    k, alerts = {}, []
    ipi = re.search(r"Inventory Performance Index.*?(\d{2,4})\s*/\s*1,000", t, re.S)
    if ipi:
        k["IPI"] = int(ipi.group(1))
    for key, pat in (("과잉 재고 비율", r"Excess inventory:\s*([\d.]+%)"), ("과잉 재고", r"Excess inventory:[^\n]*\n\s*([\d,]+ SKUs \| [\d,]+ units)"),
                     ("판매 소진율(Sell-through)", r"Sell-through:\s*([\d.]+)"), ("재고 보유율(In-stock)", r"In-stock rate:\s*([\d.]+%)"),
                     ("재입고 필요", r"Restock (\d+) SKUs"), ("묶인 재고(Stranded)", r"Stranded inventory:\s*([\d.]+%)"),
                     ("미해결 입고 문제", r"Outstanding shipment problems:?\s*(\d+)"), ("추적번호 없는 입고", r"missing tracking IDs:?\s*(\d+)"),
                     ("보관 기간 비율(주)", r"Storage utilization ratio:\s*([\d.]+)\s*weeks"),
                     ("평균 재고 부피(ft³/일)", r"Average daily inventory volume\s*\n?\s*([\d.,]+)"),
                     ("평균 출고 부피(ft³/일)", r"Average daily shipped volume\s*\n?\s*([\d.,]+)"),
                     ("장기 재고 추가요금", r"will incur \$([\d,.]+) in aged inventory surcharge on ([A-Za-z]+ \d+)")):
        m = re.search(pat, t, re.I)
        if m:
            k[key] = " ".join(m.groups()) if len(m.groups()) > 1 else m.group(1)
    for key, pat in (("재입고 대상 SKU", r"(\d+) SKUs qualify for Restock"), ("아울렛 딜 대상 SKU", r"(\d+) SKUs qualify for Outlet Deal"),
                     ("광고 추천 SKU", r"(\d+) SKUs qualify for Sponsored Products Ad")):
        m = re.search(pat, t)
        if m:
            k[key] = int(m.group(1))
    age = re.findall(r"\n(0-180|181-270|271-365|366-455|456\+)\s+([\d,]+)\s+([\d.]+%)", t)
    if age:
        k["재고 기간 분포"] = ", ".join(f"{a}일 {b}개({c})" for a, b, c in age)
    # 판단
    sur = _num(k.get("보관 기간 비율(주)"))
    vol = _num(k.get("평균 재고 부피(ft³/일)"))
    if sur is not None and sur > 22:
        rate = 1.88 if sur > 52 else 1.58 if sur > 44 else 1.16 if sur > 36 else 0.76 if sur > 28 else 0.44
        est = f", 월 약 ${rate * vol:,.0f} 예상" if vol else ""
        alerts.append(("P1", "보관 기간 추가요금", f"보관 기간 비율 {sur}주(22주 초과 → ft³당 ${rate}{est})",
                       "과잉 재고(아울렛 딜·제거) 정리 또는 입고 수량 축소. 월말 기준으로 계산됩니다"))
    if k.get("장기 재고 추가요금"):
        alerts.append(("P1", "장기 재고 추가요금", f"${k['장기 재고 추가요금']} 부과 예정(180일 이상 재고)",
                       "부과일 전에 판매 촉진·아울렛 딜·제거 주문을 검토하세요"))
    if k.get("IPI") is not None and k["IPI"] < 400:
        alerts.append(("P1", "IPI", f"IPI {k['IPI']}(400 미만)", "과잉 재고를 줄이고 판매 소진율을 높이세요"))
    if _num(k.get("재고 보유율(In-stock)")) is not None and _num(k["재고 보유율(In-stock)"]) < 90:
        alerts.append(("P2", "재고 보유율", f"In-stock {k['재고 보유율(In-stock)']} · 재입고 필요 {k.get('재입고 필요', '?')}개 SKU",
                       "판매 중인 상품의 재입고를 먼저 진행하세요"))
    if (_num(k.get("추적번호 없는 입고")) or 0) > 0:
        alerts.append(("P2", "입고 추적번호", f"추적번호 없는 입고 {k['추적번호 없는 입고']}건", "입고 배송에 추적번호를 입력하세요(분실 시 청구에 필요)"))
    return {"kpis": k, "alerts": alerts}


_SHIP_RE = re.compile(
    r"\n([^\n]*?)\n(FBA[0-9A-Z]{6,})\s*\n\s*\n?\s*(Closed|Receiving|Shipped|In transit|Working|Delivered|Checked-in|Ready to ship|Cancelled)"
    r"(?:\n([^\n]*window[^\n]*))?.*?\n\s*([A-Z][a-z]{2} \d{1,2}, \d{4})\s*\n[^\n]*\n\s*\n?\s*([A-Z][a-z]{2} \d{1,2}, \d{4})"
    r".*?\n\s*([A-Z0-9]{4})\s*\n.*?(\d+) SKUs?\s*\n\s*(\d+) units\s*\n\s*\n?\s*(\d+)\s*/\s*(\d+)", re.S)


def parse_shipments(text: str, tables: Optional[List[Dict]] = None) -> Dict:
    rows = []
    t = _main_text(text)
    for m in _SHIP_RE.finditer(t):
        name, sid, status, window, created, updated, fc, skus, units, loc, exp = m.groups()
        loc, exp = int(loc), int(exp)
        rows.append({"배송 이름": name.strip(), "배송 ID": sid, "상태": status, "배송 기간": (window or "").strip(),
                     "생성일": created, "최근 업데이트": updated, "센터": fc, "SKU 수": int(skus), "보낸 수량": int(units),
                     "확인 수량": loc, "예정 수량": exp, "차이": loc - exp})
    alerts = []
    for r in rows:
        if r["상태"] == "Closed" and r["차이"] < 0:
            alerts.append(("P1", "입고 누락", f"{r['배송 ID']} 종료됐지만 {-r['차이']}개 부족({r['확인 수량']}/{r['예정 수량']})",
                           "입고 상세(View details)에서 누락 수량을 확인하고, 기한 내에 조사 요청(청구)을 검토하세요"))
        elif r["상태"] == "Receiving" and r["차이"] < 0:
            try:
                days = (date.today() - datetime.strptime(r["최근 업데이트"], "%b %d, %Y").date()).days
            except Exception:
                days = 0
            if days >= 7 or "Outside" in r["배송 기간"]:
                alerts.append(("P2", "입고 지연", f"{r['배송 ID']} 입고 중 {r['확인 수량']}/{r['예정 수량']}"
                               + (f", {days}일째 변화 없음" if days >= 7 else "") + (" · 배송 기간 초과" if "Outside" in r["배송 기간"] else ""),
                               "계속 늘지 않으면 입고 상세에서 문제를 확인하세요"))
    total = re.search(r"Manage Shipments \((\d+)\)", t)
    return {"kpis": {"입고 배송 수": int(total.group(1)) if total else len(rows),
                     "부족 수량 합계(종료)": sum(-r["차이"] for r in rows if r["상태"] == "Closed" and r["차이"] < 0)},
            "rows": rows, "alerts": alerts}


def parse_returns(text: str, tables: Optional[List[Dict]] = None) -> Dict:
    t = _main_text(text)
    k = {}
    per = re.search(r"Comparing current 90 days \(([^)]+)\)", t)
    if per:
        k["기간"] = per.group(1)
    for key, label in (("반품률", "Return rate"), ("반품 수량", "Returned units"), ("환불 금액", "Customer refund issued"),
                       ("심각 리뷰 비율", "Critical review rate"), ("FBA 반품 수수료", "FBA return fees")):
        m = re.search(re.escape(label) + r"\s*\n\s*\n?\s*([$\d.,%]+)\s*\n\s*([+−-]\s*[\d.,]+%)?", t)
        if m:
            k[key] = m.group(1) + (f" (이전 대비 {m.group(2).replace(' ', '')})" if m.group(2) else "")
    top = _after(t, "Top return reason", r"\n\s*([^\n]+)")
    if top:
        k["주요 반품 사유"] = top
    rows = []
    tbl = _best_table(tables or [], "Return rate")
    for r in _rows(tbl):
        prod = r.get("Product ASIN") or r.get("Product") or ""
        am = re.search(r"\b(B0[A-Z0-9]{8})\b", prod)
        if not am:
            continue
        g = lambda key: re.sub(r"\s+", " ", r.get(key, "")).strip()
        rows.append({"ASIN": am.group(1), "상품": prod.replace(am.group(1), "").strip()[:90],
                     "반품률": g("Return rate"), "반품 수량": g("Returned units"), "환불 금액": g("Customer refund issued"),
                     "심각 리뷰 비율": g("Critical review rate"), "반품 사유": g("Return reasons"), "반품 빈도 배지": g("Return frequency badge")})
    alerts = []
    for r in rows:
        rate = _num(r["반품률"].split(" ")[0]) or 0
        units = _num(r["반품 수량"].split(" ")[0]) or 0
        if rate >= 10 and units >= 3:
            alerts.append(("P1", "반품률 높음", f"{r['ASIN']} 반품률 {r['반품률'].split(' ')[0]}, {int(units)}개 · {r['반품 사유'][:60]}",
                           "반품 사유에 맞춰 상세 페이지(사이즈·사진·설명)를 고치거나 상품을 점검하세요"))
        if "false advertising" in r["반품 사유"].lower() or "misleading" in r["반품 사유"].lower():
            alerts.append(("P1", "설명과 다름 반품", f"{r['ASIN']}: {r['반품 사유'][:80]}",
                           "상세 페이지 내용이 실제 상품과 다른지 확인하세요(계정 상태에도 영향)"))
    return {"kpis": k, "rows": rows, "alerts": alerts}


LISTING_STATUS = ["Active", "Search suppressed", "Improve listing quality", "At-risk of search suppression", "Inactive",
                  "Detail page removed", "Out of stock", "Listing removed", "Missing offer", "Pricing issue",
                  "Approval required", "Closed"]


def parse_listing_status(text: str) -> Dict:
    t = _main_text(text)
    i = t.find("Listing Status")
    seg = t[i:i + 1500] if i >= 0 else ""
    st = {}
    for lab in LISTING_STATUS:
        m = re.search(r"\n" + re.escape(lab) + r"\s*\n\s*(\d[\d,]*)", seg)
        if m:
            st[lab] = int(m.group(1).replace(",", ""))
    ch = {}
    j = t.find("Customer health")
    for lab in ("Very poor", "Poor", "Fair", "Good", "Excellent"):
        m = re.search(r"\n" + lab + r"\s*\n\s*(\d+)", t[j:j + 400]) if j >= 0 else None
        if m:
            ch[lab] = int(m.group(1))
    fb = re.search(r"Feedback Rating.*?Average\s*\n\s*([\d.]+).*?\((\d+) ratings\).*?Negative\s*\n\s*(\d+).*?Positive\s*\n\s*(\d+)", t, re.S)
    k = {f"상태: {a}": b for a, b in st.items()}
    if ch:
        k["고객 건강도"] = ", ".join(f"{a} {b}" for a, b in ch.items())
    if fb:
        k["판매자 피드백"] = f"{fb.group(1)}점({fb.group(2)}개, 부정 {fb.group(3)}, 긍정 {fb.group(4)})"
    alerts = []
    for lab, lv, adv in (("Search suppressed", "P0", "상품 관리에서 ‘검색 제한’ 상품의 문제(이미지·필수 정보)를 고치세요"),
                         ("Detail page removed", "P0", "상세 페이지가 삭제된 상품의 원인(정책·카탈로그)을 확인하세요"),
                         ("Listing removed", "P0", "삭제된 리스팅의 사유를 계정 상태·알림에서 확인하세요"),
                         ("At-risk of search suppression", "P1", "검색 제한이 되기 전에 지적된 항목을 고치세요"),
                         ("Approval required", "P1", "카테고리·브랜드 승인 필요 여부를 확인하세요"),
                         ("Missing offer", "P1", "판매 제안(가격·재고)이 없는 상품을 확인하세요"),
                         ("Pricing issue", "P1", "가격 오류(너무 높거나 낮음)를 확인하세요"),
                         ("Improve listing quality", "P2", "리스팅 품질 개선 제안을 확인하세요")):
        if st.get(lab):
            alerts.append((lv, f"리스팅 상태: {lab}", f"{st[lab]}개", adv))
    if ch.get("Very poor") or ch.get("Poor"):
        alerts.append(("P1", "고객 건강도 낮음", f"Very poor {ch.get('Very poor', 0)}개, Poor {ch.get('Poor', 0)}개",
                       "고객의 소리(VoC)에서 해당 상품의 불만 사유를 확인하세요"))
    return {"kpis": k, "status": st, "customer_health": ch, "alerts": alerts}


def parse_claims(texts: Dict[str, str]) -> Dict:
    k, alerts = {}, []
    cb = _main_text(texts.get("chargebacks", ""))
    if "You have no chargebacks" in cb:
        k["차지백"] = 0
    elif cb:
        n = len(re.findall(r"\d{3}-\d{7}-\d{7}", cb))
        k["차지백"] = n
        if n:
            alerts.append(("P0", "차지백", f"{n}건", "Reply-By 날짜 전에 응답하세요(Action Required 탭)"))
    az = _main_text(texts.get("atoz", ""))
    if "There are no claims that currently require a response" in az:
        k["A-to-z(응답 필요)"] = 0
    elif az:
        n = len(re.findall(r"\d{3}-\d{7}-\d{7}", az))
        k["A-to-z(응답 필요)"] = n
        if n:
            alerts.append(("P0", "A-to-z 클레임", f"응답이 필요한 클레임 {n}건", "기한 안에 응답하세요(주문 결함률에 반영)"))
    sf = _main_text(texts.get("safet", ""))
    if sf:
        m = re.search(r"Claims granted:\s*(\d+) of (\d+)", sf)
        if m:
            k["SAFE-T(최근 30일, 승인/전체)"] = f"{m.group(1)}/{m.group(2)}"
        aw = len(re.findall(r"Awaiting Seller Response", sf)) - 1
        if "Your search did not find any results" not in sf and aw > 0:
            alerts.append(("P1", "SAFE-T", "판매자 응답 대기 클레임이 있을 수 있음", "SAFE-T 페이지의 Awaiting Seller Response 탭을 확인하세요"))
    return {"kpis": k, "alerts": alerts}


def parse_voc(text: str, tables: Optional[List[Dict]] = None) -> Dict:
    t = _main_text(text)
    k = {}
    m = re.search(r"CX Health breakdown of your listings.*?How is CX Health calculated\?\s*\n\s*(\d+)\s*\n\s*(\d+)\s*\n\s*(\d+)\s*\n\s*(\d+)\s*\n\s*(\d+)", t, re.S)
    if m:
        for lab, v in zip(("Very poor", "Poor", "Fair", "Good", "Excellent"), m.groups()):
            k[f"CX {lab}"] = int(v)
    n = re.search(r"(\d+) Offer Listings", t)
    if n:
        k["대상 리스팅"] = int(n.group(1))
    rows = []
    tbl = _best_table(tables or [], "NCX rate")
    for r in _rows(tbl):
        name = r.get("Product name ASIN") or r.get("Product name") or ""
        am = re.search(r"\b(B0[A-Z0-9]{8})\b", " ".join(r.values()))
        if not am:
            continue
        rows.append({"ASIN": am.group(1), "상품": re.sub(r"\bB0[A-Z0-9]{8}\b", "", name).strip()[:90],
                     "NCX 비율": r.get("NCX rate", ""), "NCX 주문": r.get("NCX orders", ""), "전체 주문": r.get("Total orders", ""),
                     "NCX 리뷰 비율": r.get("NCX Review Rate", ""), "별점": r.get("Star rating", ""),
                     "주요 불만 사유": r.get("Top NCX reason", ""), "CX 건강도": r.get("CX Health", ""),
                     "업데이트": r.get("Last updated", "")})
    alerts = []
    for r in rows:
        ch = r["CX 건강도"].lower()
        star = _num(r["별점"].split("(")[0]) if r["별점"] else None
        if "very poor" in ch or "poor" in ch:
            alerts.append(("P1", "CX 건강도 낮음", f"{r['ASIN']} {r['CX 건강도']} · NCX {r['NCX 비율']} · {r['주요 불만 사유'][:50]}",
                           "불만 사유를 확인하고 상품·상세 페이지를 고치세요(심하면 판매 중지될 수 있음)"))
        elif star is not None and star < 3.5:
            alerts.append(("P2", "별점 낮음", f"{r['ASIN']} 별점 {r['별점']}", "낮은 별점 리뷰 내용을 확인하세요"))
    if k.get("CX Very poor") or k.get("CX Poor"):
        alerts.insert(0, ("P1", "CX 건강도", f"Very poor {k.get('CX Very poor', 0)}개, Poor {k.get('CX Poor', 0)}개",
                          "VoC 페이지에서 CX Health 필터로 해당 상품을 확인하세요"))
    return {"kpis": k, "rows": rows, "alerts": alerts}


def parse_ads_bis(text: str, tables: Optional[List[Dict]] = None) -> Dict:
    t = text
    k = {}
    dr = re.search(r"Date range\s*:\s*([^\n]+)", t)
    if dr:
        k["기간"] = dr.group(1).strip()
    m = re.search(r"Your Sponsored Products and Sponsored Brands ads show up ([\d.]+)% of the time", t)
    if m:
        k["내 광고 상단 노출 점유율"] = f"{m.group(1)}%"
    m = re.search(r"Ads from other advertisers show up ([\d.]+)% of the time", t)
    if m:
        k["다른 광고주 점유율"] = f"{m.group(1)}%"
    for key, label in (("놓친 상단 클릭", "Lost top-of-search clicks"), ("놓친 상단 주문", "Lost top-of-search orders"),
                       ("놓친 상단 매출", "Lost top-of-search sales")):
        v = _after(t, label + "\n", r"\n?\s*([^\n]+)")
        if v and v != "—":
            k[key] = v
    rows = []
    for r in _rows(_best_table(tables or [], "Top-of-search IS")):
        if len(r) > 3:
            rows.append(r)
    alerts = []
    share = _num(k.get("내 광고 상단 노출 점유율"))
    if share is not None and share < 70:
        alerts.append(("P1", "브랜드 키워드 방어", f"내 브랜드 검색에서 내 광고 상단 노출 {share:.0f}%(다른 광고주 {k.get('다른 광고주 점유율', '?')})",
                       "브랜드 키워드(정확 일치) 캠페인의 상단 입찰 조정·예산을 검토하세요"))
    return {"kpis": k, "rows": rows, "alerts": alerts}


def parse_ads_recs(text: str) -> Dict:
    i = text.find("Needs attention")
    seg = text[i:] if i >= 0 else ""
    for stop in ("\nBenchmarks\n", "© 20"):
        j = seg.find(stop)
        seg = seg[:j] if j > 0 else seg
    recs = []
    for m in re.finditer(r"\n(Recommendation|Increase sales|Reach more shoppers[^\n]*|Improve [^\n]+|Fix [^\n]+|[A-Z][^\n]{8,60})\n\n([^\n]{30,400})", seg):
        title, body = m.group(1).strip(), m.group(2).strip()
        if title in ("Needs attention", "Recommended for you", "Opportunities"):
            continue
        recs.append({"제목": title, "내용": body})
    k = {"추천 수": len(recs)}
    na = re.search(r"Needs attention\s*\n\s*\n?Resolve these to keep your campaigns on track\.\s*\n\s*\n?([^\n]+)", seg)
    if na and not na.group(1).startswith("Recommended for you"):
        k["처리 필요"] = na.group(1)[:120]
    alerts = [("P3", "광고 추천", f"{r['제목']}: {r['내용'][:80]}", "광고 콘솔 → Recommendations에서 확인(자동 적용 안 함)") for r in recs[:5]]
    return {"kpis": k, "rows": recs, "alerts": alerts}


# ================================================================ 페이지 읽기(하루 한 번)

def _cache(key: str) -> Path:
    return day_dir() / "pages" / f"{key}.json"


def cached(key: str) -> Optional[Dict]:
    p = _cache(key)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _save(key: str, data: Dict) -> Dict:
    data["key"] = key
    data["name"] = NAMES.get(key, key)
    data["time"] = datetime.now().isoformat(timespec="seconds")
    data["alerts"] = [list(a) for a in data.get("alerts", [])]
    p = _cache(key)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return data


def _open(page, url: str, screens: int = 3) -> str:
    human.goto(page, url, wait_after=False)
    sc_nav._settle(page)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    human.read_page(page, max_screens=screens)
    return page.inner_text("body")


def _tables(page) -> List[Dict]:
    try:
        return page.evaluate(_TABLES_JS)
    except Exception:
        return []


def read(page, key: str, force: bool = False, entity_id: str = "") -> Dict:
    """페이지 하나 읽기(오늘 읽었으면 저장된 결과). 실패하면 {"error": ...}."""
    if not force:
        c = cached(key)
        if c and not c.get("error"):
            print(f"   {NAMES.get(key, key)}: 오늘 이미 읽음({c.get('time', '')[11:16]})")
            return c
    print(f"   {NAMES.get(key, key)} 읽는 중...")
    try:
        if key == "claims":
            texts = {k: _open(page, sc_nav.SC_BASE + URLS[k], 1) for k in ("chargebacks", "atoz", "safet")}
            data = parse_claims(texts)
        elif key in ("ads_bis", "ads_recs"):
            if not entity_id:
                raise RuntimeError("광고 계정(entityId)을 알 수 없습니다")
            url = f"{ADS}/bb/bis?entityId={entity_id}" if key == "ads_bis" else f"{ADS}/recommendations/prerelease?entityId={entity_id}"
            text = _open(page, url, 3)
            data = parse_ads_bis(text, _tables(page)) if key == "ads_bis" else parse_ads_recs(text)
        else:
            text = _open(page, sc_nav.SC_BASE + URLS[key], 4 if key in ("shipments", "returns", "voc") else 2)
            fn = {"home": lambda: parse_home(text), "fba_dashboard": lambda: parse_fba_dashboard(text),
                  "shipments": lambda: parse_shipments(text, _tables(page)), "returns": lambda: parse_returns(text, _tables(page)),
                  "listing_status": lambda: parse_listing_status(text), "voc": lambda: parse_voc(text, _tables(page))}[key]
            data = fn()
        if not data.get("kpis") and not data.get("rows"):
            _dump(page, f"page_{key}_empty")
            data["note"] = "페이지에서 값을 찾지 못했습니다(화면 구조 변경 가능, reports/daily/debug 확인)"
        print(f"   ✅ {NAMES.get(key, key)}: 항목 {len(data.get('kpis', {}))}개, 행 {len(data.get('rows', []))}개, "
              f"알림 {len(data.get('alerts', []))}개")
        return _save(key, data)
    except Exception as e:
        _dump(page, f"page_{key}_failed")
        msg = f"{type(e).__name__}: {str(e).splitlines()[0][:150]}"
        print(f"   ⚠ {NAMES.get(key, key)} 읽기 실패: {msg}")
        return {"key": key, "name": NAMES.get(key, key), "error": msg, "kpis": {}, "rows": [], "alerts": []}


def read_many(keys: List[str], page=None, entity_id: str = "") -> Dict[str, Dict]:
    """필요할 때만 브라우저를 엶(모두 오늘 결과가 있으면 열지 않음)."""
    out = {k: cached(k) for k in keys}
    todo = [k for k in keys if not out[k] or out[k].get("error")]
    if not todo:
        for k in keys:
            print(f"   {NAMES.get(k, k)}: 오늘 이미 읽음")
        return out
    if page is not None:
        for k in todo:
            out[k] = read(page, k, entity_id=entity_id)
        return out
    from sc_browser import open_session
    with open_session(headless=False) as pg:
        for k in todo:
            out[k] = read(pg, k, entity_id=entity_id)
    return out


# ================================================================ 기존 Excel에 시트 추가

def append_sheets(xlsx: Path, results: List[Dict]) -> None:
    """페이지 읽기 결과를 기존 보고서에 추가: 페이지마다 시트 하나(요점 + 알림 + 표)."""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    if not results or not xlsx or not Path(xlsx).exists():
        return
    wb = load_workbook(xlsx)
    FONT = inv.FONT
    HEAD = PatternFill("solid", fgColor="DCE6F1")
    fills = {"P0": PatternFill("solid", fgColor="F8D7DA"), "P1": PatternFill("solid", fgColor="FFE8CC")}
    for res in results:
        if not res:
            continue
        title = res.get("name", res.get("key", "페이지"))[:28]
        if title in wb.sheetnames:
            del wb[title]
        ws = wb.create_sheet(title)
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 70
        r = 1
        ws.cell(r, 1, f"{title} ({res.get('src', 'Seller Central 화면에서 읽음')} · {res.get('time', '')[:16].replace('T', ' ')})").font = Font(name=FONT, size=13, bold=True)
        r += 1
        if res.get("error"):
            ws.cell(r, 1, "읽기 실패").font = Font(name=FONT, bold=True, color="C01C1C")
            ws.cell(r, 2, res["error"]).font = Font(name=FONT)
            continue
        if res.get("note"):
            ws.cell(r, 1, res["note"]).font = Font(name=FONT, color="6B6B6B")
            r += 1
        r += 1
        for a, b in res.get("kpis", {}).items():
            ws.cell(r, 1, a).font = Font(name=FONT)
            ws.cell(r, 2, str(b)).font = Font(name=FONT, bold=True)
            r += 1
        if res.get("alerts"):
            r += 1
            for c, h in enumerate(["등급", "분류", "내용", "조치"], start=1):
                x = ws.cell(r, c, h)
                x.font, x.fill = Font(name=FONT, bold=True), HEAD
            ws.column_dimensions["C"].width = 60
            ws.column_dimensions["D"].width = 60
            r += 1
            for al in res["alerts"]:
                for c, v in enumerate(al, start=1):
                    x = ws.cell(r, c, v)
                    x.font = Font(name=FONT, size=10)
                    x.alignment = Alignment(wrap_text=True, vertical="top")
                if al[0] in fills:
                    ws.cell(r, 1).fill = fills[al[0]]
                r += 1
        rows = res.get("rows") or []
        if rows:
            r += 1
            heads = list(dict.fromkeys(k for row in rows for k in row))
            for c, h in enumerate(heads, start=1):
                x = ws.cell(r, c, h)
                x.font, x.fill = Font(name=FONT, bold=True), HEAD
                x.alignment = Alignment(wrap_text=True)
            r += 1
            for row in rows:
                for c, h in enumerate(heads, start=1):
                    ws.cell(r, c, row.get(h, "")).font = Font(name=FONT, size=10)
                r += 1
    wb.save(xlsx)


def print_alerts(results: List[Dict], limit: int = 6) -> None:
    al = [(a, r.get("name", "")) for r in results if r for a in r.get("alerts", [])]
    al.sort(key=lambda x: x[0][0])
    for a, name in al[:limit]:
        print(f"  [{a[0]}] {name} · {a[1]}: {a[2][:80]}")

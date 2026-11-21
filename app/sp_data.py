"""
SP-API로 가져오는 자료 → 기존 기능이 쓰던 형식(파일·열 이름)으로 변환

각 함수는 실패하면 예외를 냅니다. 기존 기능에서는 sp_api.try_api(...)로 감싸서
API가 안 되면 원래 방식(브라우저)으로 자동 전환합니다.
"""

import json
import re
import time
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import sp_api

PCT = lambda v: "" if v in (None, "") else f"{float(v) * 100:.2f}%"


def _c() -> sp_api.Client:
    return sp_api.client()


def _day_start(d: date) -> datetime:
    return datetime(d.year, d.month, d.day, tzinfo=timezone.utc)


def _now_minus(minutes: int = 3) -> datetime:
    return datetime.now(timezone.utc) - timedelta(minutes=minutes)


# ================================================================ 1. 전체 상품(All Listings) · FBA 재고

def all_listings(today: str) -> Optional[Path]:
    return _c().report("GET_MERCHANT_LISTINGS_ALL_DATA", save_as=f"all_listings_{today}")


def fba_inventory(today: str) -> Optional[Path]:
    """FBA 재고(재고 계획) 보고서: 판매 가능·입고·재고 기간·재고 일수·권장 조치 등(화면의 ‘FBA Inventory’ 보고서와 같은 내용)."""
    c = _c()
    try:
        return c.report("GET_FBA_INVENTORY_PLANNING_DATA", save_as=f"fba_inventory_{today}")
    except sp_api.SPAPIError as e:
        print(f"     재고 계획 보고서 실패({e}) → 관리 재고 보고서로 다시 시도")
        return c.report("GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA", save_as=f"fba_myi_{today}")


# ================================================================ 2. 주문 · 반품

def orders(start: date, end: date, tag: str) -> Optional[Path]:
    """주문일 기준 전체 주문(화면의 All Orders 보고서와 같은 열). 시간대 차이를 고려해 앞뒤 하루씩 넓게 받음."""
    s = _day_start(start) - timedelta(days=1)
    e = min(_day_start(end) + timedelta(days=2), _now_minus())
    return _c().report("GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL", s, e, save_as=f"orders_{tag}")


def returns(start: date, end: date, tag: str) -> Optional[Path]:
    s = _day_start(start) - timedelta(days=1)
    e = min(_day_start(end) + timedelta(days=2), _now_minus())
    return _c().report("GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA", s, e, save_as=f"returns_{tag}")


# ================================================================ 3. 비즈니스 보고서(하위 ASIN별 트래픽·전환)

ST_HEADERS = ["(Parent) ASIN", "(Child) ASIN", "Title", "SKU", "Sessions - Total", "Session Percentage - Total",
              "Page Views - Total", "Page Views Percentage - Total", "Featured Offer (Buy Box) Percentage", "Units Ordered",
              "Unit Session Percentage", "Ordered Product Sales", "Total Order Items"]


def sales_traffic(start: date, end: date, titles: Optional[Dict[str, str]] = None) -> Tuple[List[str], List[Dict]]:
    path = _c().report("GET_SALES_AND_TRAFFIC_REPORT", _day_start(start), _day_start(end) + timedelta(hours=23, minutes=59),
                       {"dateGranularity": "DAY", "asinGranularity": "CHILD"}, save_as=f"sales_traffic_{start}_{end}")
    if not path:
        return ST_HEADERS, []
    d = json.loads(path.read_text(encoding="utf-8"))
    rows = []
    for a in d.get("salesAndTrafficByAsin", []):
        s, t = a.get("salesByAsin", {}), a.get("trafficByAsin", {})
        child = a.get("childAsin") or a.get("parentAsin") or ""
        rows.append({
            "(Parent) ASIN": a.get("parentAsin", ""), "(Child) ASIN": child, "Title": (titles or {}).get(child, ""),
            "SKU": a.get("sku", ""), "Sessions - Total": t.get("sessions", 0),
            "Session Percentage - Total": f"{t.get('sessionPercentage', 0)}%", "Page Views - Total": t.get("pageViews", 0),
            "Page Views Percentage - Total": f"{t.get('pageViewsPercentage', 0)}%",
            "Featured Offer (Buy Box) Percentage": f"{t.get('buyBoxPercentage', 0)}%", "Units Ordered": s.get("unitsOrdered", 0),
            "Unit Session Percentage": f"{t.get('unitSessionPercentage', 0)}%",
            "Ordered Product Sales": (s.get("orderedProductSales") or {}).get("amount", 0),
            "Total Order Items": s.get("totalOrderItems", 0)})
    return ST_HEADERS, rows


def listing_titles() -> Dict[str, str]:
    """ASIN → 상품명(최근 All Listings 파일에서)."""
    import inventory as inv
    files = sorted((sp_api.BASE_DIR / "reports" / "inventory").glob("all_listings_*.xlsx"))
    out = {}
    if files:
        for r in inv.read_xlsx(files[-1])[1]:
            a = str(inv._get(r, "asin1", "asin", default="")).strip()
            if a:
                out[a] = str(inv._get(r, "item-name", default="") or "")
    return out


# ================================================================ 4. Brand Analytics(SCP · SQP)

def _week(week_end: str) -> Tuple[datetime, datetime]:
    e = date.fromisoformat(week_end)
    s = e - timedelta(days=6)
    return _day_start(s), _day_start(e)


def scp_week(week_end: str) -> Optional[List[Dict]]:
    """Search Catalog Performance(브랜드의 모든 ASIN) → 화면 표와 같은 열 이름."""
    s, e = _week(week_end)
    path = _c().report("GET_BRAND_ANALYTICS_SEARCH_CATALOG_PERFORMANCE_REPORT", s, e, {"reportPeriod": "WEEK"},
                       save_as=f"scp_{week_end}")
    if not path:
        return None
    titles = listing_titles()
    rows = []
    for a in json.loads(path.read_text(encoding="utf-8")).get("dataByAsin", []):
        im, cl, ca, pu = (a.get(k, {}) for k in ("impressionData", "clickData", "cartAddData", "purchaseData"))
        rows.append({"ASIN": a.get("asin", ""), "Title": titles.get(a.get("asin", ""), ""),
                     "Impressions": im.get("impressionCount", 0),
                     "Price (Median)": (im.get("impressionMedianPrice") or {}).get("amount", ""),
                     "Clicks": cl.get("clickCount", 0), "Click Rate (CTR)": PCT(cl.get("clickRate")),
                     "Cart Adds": ca.get("cartAddCount", 0), "Purchases": pu.get("purchaseCount", 0),
                     "Search Traffic Sales": (pu.get("searchTrafficSales") or {}).get("amount", 0),
                     "Conversion Rate": PCT(pu.get("conversionRate"))})
    return rows


def sqp_week(week_end: str, asins: List[str], max_asins: int = 90) -> Optional[List[Dict]]:
    """Search Query Performance: API는 ASIN 목록(200자 이내)으로만 요청 가능 → 여러 번 요청해 검색어별로 합쳐 ‘브랜드 보기’처럼 만듦."""
    asins = [a for a in dict.fromkeys(asins) if re.match(r"^B0[A-Z0-9]{8}$", a)][:max_asins]
    if not asins:
        return None
    s, e = _week(week_end)
    groups, cur = [], []
    for a in asins:
        if len(" ".join(cur + [a])) > 200:
            groups.append(cur)
            cur = []
        cur.append(a)
    if cur:
        groups.append(cur)
    c = _c()
    print(f"   SQP: ASIN {len(asins)}개를 {len(groups)}번에 나눠 요청")
    ids = []
    for g in groups:
        ids.append(c.create_report("GET_BRAND_ANALYTICS_SEARCH_QUERY_PERFORMANCE_REPORT", s, e,
                                   {"reportPeriod": "WEEK", "asin": " ".join(g)}))
        time.sleep(2)
    agg: Dict[str, Dict] = {}
    for i, rid in enumerate(ids, 1):
        doc = c.wait_report(rid, 45, f"SQP {i}/{len(ids)}")
        if not doc:
            continue
        data = json.loads(c.download(doc).decode("utf-8"))
        for a in data.get("dataByAsin", []):
            q = (a.get("searchQueryData") or {}).get("searchQuery")
            if not q:
                continue
            x = agg.setdefault(q, defaultdict(float))
            sq = a["searchQueryData"]
            x["score"] = sq.get("searchQueryScore") or x["score"]
            x["volume"] = max(x["volume"], sq.get("searchQueryVolume") or 0)
            for grp, tot_k, rate_k, asin_k in (("impressionData", "totalQueryImpressionCount", None, "asinImpressionCount"),
                                               ("clickData", "totalClickCount", "totalClickRate", "asinClickCount"),
                                               ("cartAddData", "totalCartAddCount", "totalCartAddRate", "asinCartAddCount"),
                                               ("purchaseData", "totalPurchaseCount", "totalPurchaseRate", "asinPurchaseCount")):
                gd = a.get(grp) or {}
                x[grp + "_total"] = max(x[grp + "_total"], gd.get(tot_k) or 0)
                if rate_k:
                    x[grp + "_rate"] = max(x[grp + "_rate"], gd.get(rate_k) or 0)
                x[grp + "_brand"] += gd.get(asin_k) or 0
    if not agg:
        return None
    rows = []
    for q, x in sorted(agg.items(), key=lambda kv: (kv[1]["score"] or 9999)):
        r = {"Search Query": q, "Search Query Score": int(x["score"] or 0), "Search Query Volume": int(x["volume"])}
        for grp, name, rate_name in (("impressionData", "Impressions", None), ("clickData", "Clicks", "Click Rate"),
                                     ("cartAddData", "Cart Adds", "Cart Add Rate"), ("purchaseData", "Purchases", "Purchase Rate")):
            tot, brand = x[grp + "_total"], x[grp + "_brand"]
            r[f"{name}: Total Count"] = int(tot)
            if rate_name:
                r[f"{name}: {rate_name}"] = PCT(x[grp + "_rate"])
            r[f"{name}: Brand Count"] = int(brand)
            r[f"{name}: Brand Share"] = PCT(brand / tot) if tot else "0.00%"
        rows.append(r)
    return rows


# ================================================================ 5. 계정 상태(성과 보고서)

STATUS_KO = {"GOOD": "좋음", "FAIR": "보통", "POOR": "나쁨", "BAD": "나쁨", "NONE": "없음", "NORMAL": "정상",
             "AT_RISK": "위험", "DEACTIVATED": "비활성", "WARNING": "경고", "WARNED": "경고", "HEALTHY": "건강", "CRITICAL": "위험"}


def performance() -> Dict:
    """GET_V2_SELLER_PERFORMANCE_REPORT → {"kpis", "rows", "alerts"} (sc_pages 결과와 같은 형식)."""
    path = _c().report("GET_V2_SELLER_PERFORMANCE_REPORT", save_as=f"seller_performance_{date.today()}")
    if not path:
        raise sp_api.SPAPIError(0, "성과 보고서가 비어 있음")
    d = json.loads(path.read_text(encoding="utf-8"))
    mp = sp_api.load_config().get("marketplace_id") or sp_api.US
    k, rows, alerts = {}, [], []
    for st in d.get("accountStatuses", []):
        if st.get("marketplaceId") == mp:
            k["계정 상태(API)"] = STATUS_KO.get(st.get("status"), st.get("status"))
            if st.get("status") != "NORMAL":
                alerts.append(("P0", "계정 상태", f"계정 상태: {st.get('status')}", "계정 상태 페이지에서 즉시 확인하세요"))
    pm = next((x for x in d.get("performanceMetrics", []) if x.get("marketplaceId") == mp), None) or \
        (d.get("performanceMetrics") or [{}])[0]
    ahr = pm.get("accountHealthRating") or {}
    if ahr:
        st_ = ahr.get("ahrStatus", "")
        k["계정 상태 등급(AHR)"] = (f"{ahr['ahrScore']} " if ahr.get("ahrScore") is not None else "") + STATUS_KO.get(st_, st_)
        if st_ and st_ not in ("GOOD", "NORMAL", "HEALTHY"):
            alerts.append(("P0" if st_ in ("CRITICAL", "AT_RISK", "BAD") else "P1", "계정 상태 등급", f"AHR 상태 {st_}",
                           "계정 상태 페이지에서 위반 항목을 확인하세요"))
    names = {"orderDefectRate": "주문 결함률(ODR)", "lateShipmentRate": "배송 지연율", "preFulfillmentCancellationRate": "출고 전 취소율",
             "validTrackingRate": "유효 추적률", "onTimeDeliveryRate": "정시 배송률", "invoiceDefectRate": "인보이스 결함률",
             "listingPolicyViolations": "리스팅 정책 위반", "productAuthenticityCustomerComplaints": "진품 불만",
             "productConditionCustomerComplaints": "상품 상태 불만", "productSafetyCustomerComplaints": "안전 불만",
             "receivedIntellectualPropertyComplaints": "지식재산권 신고", "restrictedProductPolicyViolations": "제한 상품 위반",
             "suspectedIntellectualPropertyViolations": "지식재산권 의심", "foodAndProductSafetyIssues": "식품·안전 이슈",
             "otherPolicyViolations": "기타 정책 위반", "documentRequests": "서류 요청"}
    for key, name in names.items():
        v = pm.get(key)
        if isinstance(v, list):          # orderDefectRate 등은 배송 방식별 목록
            for i, x in enumerate(v):
                _metric(rows, alerts, f"{name}({x.get('fulfillmentChannel', i)})", x)
        elif isinstance(v, dict) and any(isinstance(v.get(ch), dict) for ch in ("mfn", "afn")):
            for ch in ("afn", "mfn"):
                if isinstance(v.get(ch), dict):
                    _metric(rows, alerts, f"{name}({'FBA' if ch == 'afn' else '판매자 배송'})", v[ch])
        elif isinstance(v, dict):
            _metric(rows, alerts, name, v)
    for r in rows:
        if r["값"] not in ("", None):
            k[r["지표"]] = f"{r['값']} ({r['상태']})"
    return {"kpis": k, "rows": rows, "alerts": alerts}


def _metric(rows, alerts, name, x):
    st = x.get("status", "")
    val = x.get("rate")
    if val is not None:
        val = f"{val * 100:.2f}%"
    elif x.get("defectsCount") is not None:
        val = x.get("defectsCount")
    elif x.get("count") is not None:
        val = x.get("count")
    tgt = x.get("targetValue")
    rng = x.get("reportingDateRange") or {}
    extra = ""
    for sub, ko in (("claims", "A-to-z"), ("aToZGuaranteeClaims", "A-to-z"), ("chargebacks", "차지백"),
                    ("negativeFeedback", "부정 피드백"), ("negativeFeedbacks", "부정 피드백")):
        if isinstance(x.get(sub), dict):
            extra += f" {ko} {x[sub].get('count', '')}"
    rows.append({"지표": name, "값": val if val is not None else "", "상태": STATUS_KO.get(st, st),
                 "목표": ("" if tgt is None else (f"{tgt * 100:.2f}%" if isinstance(tgt, float) and tgt < 1 else tgt)),
                 "기간": f"{str(rng.get('reportingDateFrom', ''))[:10]} ~ {str(rng.get('reportingDateTo', ''))[:10]}",
                 "세부": extra.strip()})
    if st in ("BAD", "POOR"):
        alerts.append(("P0", f"계정 지표: {name}", f"{val} (목표 {tgt}) 상태 {st}", "계정 상태 페이지에서 원인을 확인하고 조치하세요"))
    elif st in ("FAIR", "WARNING"):
        alerts.append(("P1", f"계정 지표: {name}", f"{val} 상태 {st}", "기준을 넘기 전에 확인하세요"))


# ================================================================ 6. 리스팅(Listings Items API)

def listings_items(seller_id: str, limit: int = 3000) -> Dict[str, Dict]:
    """모든 SKU의 제목·불릿·검색어·이미지·Item Highlights + Amazon이 알려 주는 문제(issues)·상태."""
    c = _c()
    out = {}
    params = {"marketplaceIds": c.mp, "includedData": "summaries,attributes,issues", "pageSize": 20, "issueLocale": "en_US"}
    for it in c.pages(f"/listings/2021-08-01/items/{seller_id}", params, "items", limit=limit):
        sku = it.get("sku")
        if not sku:
            continue
        at = it.get("attributes") or {}
        vals = lambda k: [str(x.get("value", "")).strip() for x in (at.get(k) or []) if x.get("value")]
        summ = next((s for s in it.get("summaries", []) if s.get("marketplaceId") == c.mp), (it.get("summaries") or [{}])[0])
        imgs = 0
        main = ""
        for k, v in at.items():
            if "image_locator" in k and v:
                imgs += 1
                if k.startswith("main_"):
                    main = (v[0] or {}).get("media_location", "")
        hl_keys = [k for k in at if "highlight" in k]
        out[sku] = {"asin": summ.get("asin", ""), "title": (vals("item_name") or [summ.get("itemName", "")])[0],
                    "bullets": vals("bullet_point"), "search_terms": " ".join(vals("generic_keyword")),
                    "highlights": (" ".join(v for k in hl_keys for v in vals(k)) if hl_keys else None),
                    "main_image": main or (summ.get("mainImage") or {}).get("link", ""), "images": imgs,
                    "brand": (vals("brand") or [""])[0], "status": summ.get("status", []),
                    "issues": it.get("issues", [])}
    return out


def issue_findings(items: Dict[str, Dict]) -> List[Tuple[str, str, str, str, str, str]]:
    """Amazon이 알려 준 문제 → (sku, asin, 등급, 분류, 내용, 조치)."""
    out = []
    for sku, it in items.items():
        for iss in it.get("issues") or []:
            sev = iss.get("severity", "")
            acts = [a.get("action", "") for a in ((iss.get("enforcements") or {}).get("actions") or [])]
            lv = "P0" if acts or sev == "ERROR" else ("P2" if sev == "WARNING" else "P3")
            what = ", ".join(acts) if acts else ", ".join(iss.get("categories") or []) or sev
            attrs = ", ".join(iss.get("attributeNames") or [])
            out.append((sku, it.get("asin", ""), lv, f"Amazon 지적({what})",
                        f"{iss.get('message', '')[:200]}" + (f" [속성: {attrs}]" if attrs else "") + f" (코드 {iss.get('code', '')})",
                        "상품 편집에서 해당 속성을 고치세요" + (" — 노출·판매에 영향 중" if acts else "")))
    return out


# ================================================================ 7. 새 기능: Amazon 할 일(Action Items) · 리뷰 주제

def action_items(limit: int = 200) -> List[Dict]:
    c = _c()
    params = {"marketplaceIds": c.mp, "statuses": "PENDING_SELLING_PARTNER,OVERDUE", "includedData": "impacts",
              "pageSize": 20, "locale": "en-US"}
    return list(c.pages("/actionItems/2026-06-01/items", params, "actionItems", limit=limit))


def _txt(v) -> str:
    return (v or {}).get("text", "") if isinstance(v, dict) else str(v or "")


def action_alerts(items: List[Dict]) -> List[Tuple[str, str, str, str]]:
    out = []
    for it in items:
        ty = it.get("type") or {}
        cls = [c.upper() for c in ty.get("classifications") or []]
        name = _txt((ty.get("definition") or {}).get("displayName")) or (ty.get("definition") or {}).get("typeDefinitionId", "")
        st = [x.upper() for x in (it.get("status") or [])]
        imps = [c for im in it.get("impacts") or [] for c in (im.get("classifications") or [])]
        lv = ("P0" if "ENFORCEMENT" in imps or "OVERDUE" in st else
              "P1" if "ISSUE" in cls else "P2" if "TASK" in cls else "P3")
        benefit = ""
        for im in it.get("impacts") or []:
            for e in im.get("estimates") or []:
                if e.get("type") == "MONEY":
                    benefit = f" · {_txt((im.get('definition') or {}).get('displayName'))} {e.get('minimum', '')}~{e.get('maximum', '')} {e.get('currencyCode', '')}"
                    break
        due = it.get("dueDate") or ""
        out.append((lv, f"Amazon 할 일({'/'.join(cls) or '기타'})",
                    f"{name}: {_txt(it.get('description'))[:180]}{benefit}" + (f" (기한 {str(due)[:10]})" if due else ""),
                    "Seller Central에서 확인(자동 처리 안 함)"))
    return out


def review_topics(asin: str) -> Dict:
    c = _c()
    return c.get(f"/customerFeedback/2024-06-01/items/{asin}/reviews/topics", marketplaceId=c.mp, sortBy="MENTIONS")


def negative_topics(asins: List[str], max_asins: int = 15) -> List[Dict]:
    """광고비·매출 상위 ASIN의 부정 리뷰 주제(언급 수·별점 영향·예시 문장)."""
    rows = []
    for a in asins[:max_asins]:
        try:
            d = review_topics(a)
        except sp_api.SPAPIError as e:
            if e.status in (403, 404):
                continue
            raise
        for t in ((d.get("topics") or {}).get("negativeTopics") or [])[:5]:
            m = t.get("asinMetrics") or {}
            rows.append({"ASIN": a, "주제": t.get("topic", ""), "언급 수": m.get("numberOfMentions", ""),
                         "언급 비율(%)": m.get("occurrencePercentage", ""), "별점 영향": m.get("starRatingImpact", ""),
                         "예시": " / ".join((t.get("reviewSnippets") or [])[:3])[:300]})
        time.sleep(1.1)
    return rows

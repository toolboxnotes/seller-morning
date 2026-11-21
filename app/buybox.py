"""
Buy Box · 합승 판매자 확인(SP-API Product Pricing, 조회만)

판매 중인 ASIN마다(20개씩 묶어 10초 간격) 새 상품 오퍼를 확인:
  - Buy Box를 다른 판매자가 가져감 → P1 (가져간 판매자·가격, 내 가격)
  - 다른 판매자가 같은 ASIN을 판매(합승) → P2, 어제 없던 판매자가 새로 나타나면 P1
  - 추천 오퍼(Buy Box)가 아무에게도 없음 → P2 (가격이 다른 곳보다 높거나 재고 문제일 수 있음)
  - 내 가격이 Amazon ‘경쟁 가격 기준’보다 높음 → P2 (Buy Box가 숨겨질 수 있음)
  - 내 오퍼가 보이지 않음(품절·비활성) → P3
결과: reports/buybox/buybox_날짜.xlsx, 날짜별 기록 reports/buybox/data/날짜.json(어제와 비교)
필요 권한: 개발자 앱의 Pricing 역할
"""

import json
import os
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "reports" / "buybox"
DATA_DIR = OUT_DIR / "data"
BATCH = 20


def my_asins() -> Dict[str, Dict]:
    """판매 중 ASIN → {sku, title}: 오늘(없으면 최근) 전체 상품 보고서에서."""
    import inventory as inv
    files = sorted((BASE_DIR / "reports" / "inventory").glob("all_listings_*.xlsx"))
    out: Dict[str, Dict] = {}
    if not files:
        return out
    for r in inv.read_xlsx(files[-1])[1]:
        a = str(inv._get(r, "asin1", "asin", default="")).strip()
        st = str(inv._get(r, "status", default="Active") or "Active").lower()
        if a and st.startswith("active") and a not in out:
            out[a] = {"sku": str(inv._get(r, "seller-sku", "sku", default="") or ""),
                      "title": str(inv._get(r, "item-name", default="") or "")[:80],
                      "price": inv._get(r, "price", default="")}
    return out


def _money(x) -> Optional[float]:
    try:
        return float((x or {}).get("Amount"))
    except (TypeError, ValueError):
        return None


def fetch(c, asins: List[str], log=print) -> Dict[str, Dict]:
    """→ {asin: {"offers": [...], "summary": {...}} 또는 {"error": ...}}"""
    out: Dict[str, Dict] = {}
    groups = [asins[i:i + BATCH] for i in range(0, len(asins), BATCH)]
    for n, g in enumerate(groups, 1):
        body = {"requests": [{"uri": f"/products/pricing/v0/items/{a}/offers", "method": "GET", "MarketplaceId": c.mp,
                              "ItemCondition": "New", "CustomerType": "Consumer"} for a in g]}
        d = c.request("POST", "/batches/products/pricing/v0/itemOffers", None, body)
        for resp in d.get("responses", []):
            req = resp.get("request") or {}
            a = req.get("Asin") or ""
            pl = (resp.get("body") or {}).get("payload") or {}
            a = a or pl.get("ASIN", "")
            if not a:
                continue
            code = (resp.get("status") or {}).get("statusCode", 200)
            if code != 200 or not pl:
                errs = (resp.get("body") or {}).get("errors") or []
                out[a] = {"error": f"{code} {(errs[0].get('message') if errs else '')}"[:150]}
            else:
                out[a] = {"offers": pl.get("Offers") or [], "summary": pl.get("Summary") or {}}
        log(f"   [{n}/{len(groups)}] ASIN {len(g)}개 확인")
        if n < len(groups):
            time.sleep(10.5)                # 이 API는 10초에 1번
    return out


def analyze(data: Dict[str, Dict], info: Dict[str, Dict], seller: str, prev: Dict[str, Dict]) -> Tuple[List[Dict], List[Dict], List[Tuple]]:
    rows, others, alerts = [], [], []
    for a, d in data.items():
        it = info.get(a, {})
        name = it.get("sku") or a
        if d.get("error"):
            rows.append({"상태": "확인 실패", "ASIN": a, "SKU": it.get("sku", ""), "상품명": it.get("title", ""), "메모": d["error"]})
            continue
        offers, summ = d["offers"], d["summary"]
        mine = [o for o in offers if o.get("MyOffer") or (seller and o.get("SellerId") == seller)]
        rest = [o for o in offers if o not in mine]
        winner = next((o for o in offers if o.get("IsBuyBoxWinner")), None)
        landed = lambda o: (_money(o.get("ListingPrice")) or 0) + (_money(o.get("Shipping")) or 0)
        my_p = landed(mine[0]) if mine else None
        bb_price = next((_money(b.get("LandedPrice")) for b in summ.get("BuyBoxPrices") or [] if b.get("condition", "New").lower() == "new"), None)
        thr = _money(summ.get("CompetitivePriceThreshold"))      # 다른 쇼핑몰 기준 경쟁 가격(MoneyType)
        prev_sellers = set((prev.get(a) or {}).get("sellers", []))
        now_sellers = {o.get("SellerId") for o in rest if o.get("SellerId")}
        new_sellers = sorted(now_sellers - prev_sellers) if prev else []
        st, memo = "정상(내가 Buy Box)", []
        if not mine:
            st = "내 오퍼 안 보임"
            alerts.append(("P3", "내 오퍼 안 보임", f"{name}: 새 상품 오퍼 목록에 내 오퍼가 없음(품절·비활성·가격 순위 밖)", "재고·리스팅 상태 확인"))
        elif winner and winner not in mine:
            st = "Buy Box 잃음"
            wp = landed(winner)
            memo.append(f"Buy Box: 판매자 {winner.get('SellerId', '?')} ${wp:.2f}"
                        f"{'(FBA)' if winner.get('IsFulfilledByAmazon') else ''} · 내 가격 ${my_p:.2f}")
            alerts.append(("P1", "Buy Box 잃음", f"{name}: 다른 판매자 {winner.get('SellerId', '?')}가 ${wp:.2f}로 Buy Box(내 가격 ${my_p:.2f})",
                           "가격·재고·배송 기간 확인. 정품이 아닌 합승이면 Brand Registry → Report a violation"))
        elif not winner and not bb_price:
            st = "Buy Box 없음"
            alerts.append(("P2", "Buy Box 없음", f"{name}: 추천 오퍼(Buy Box)가 아무에게도 없음 · 내 가격 ${my_p:.2f}",
                           "가격이 다른 곳보다 높거나 배송·재고 문제일 수 있음 → 가격·재고 확인"))
        if mine and thr and my_p and my_p > thr + 0.005:
            memo.append(f"내 가격 ${my_p:.2f} > 경쟁 가격 기준 ${thr:.2f}")
            alerts.append(("P2", "경쟁 가격 기준 초과", f"{name}: 내 가격 ${my_p:.2f}가 Amazon 경쟁 가격 기준 ${thr:.2f}보다 높음",
                           "Buy Box가 숨겨질 수 있음 → 가격 검토(‘리스팅·가격 수정’)"))
        if rest:
            low = min(landed(o) for o in rest)
            memo.append(f"다른 판매자 {len(rest)}명(최저 ${low:.2f})")
            if new_sellers:
                alerts.append(("P1", "새 합승 판매자", f"{name}: 어제 없던 판매자 {', '.join(new_sellers)[:80]} 등장(최저 ${low:.2f})",
                               "판매자 페이지·상품 확인 → 위조·상표 침해면 Brand Registry → Report a violation"))
            elif st.startswith("정상"):
                alerts.append(("P2", "합승 판매자", f"{name}: 다른 판매자 {len(rest)}명 판매 중(최저 ${low:.2f})",
                               "판매자·가격 확인, 필요하면 Brand Registry → Report a violation"))
            for o in rest:
                others.append({"ASIN": a, "SKU": it.get("sku", ""), "판매자 ID": o.get("SellerId", ""),
                               "가격(배송 포함)": round(landed(o), 2), "FBA": "Y" if o.get("IsFulfilledByAmazon") else "",
                               "Buy Box": "Y" if o.get("IsBuyBoxWinner") else "", "새로 나타남": "Y" if o.get("SellerId") in new_sellers else "",
                               "판매자 평점": ((o.get("SellerFeedbackRating") or {}).get("SellerPositiveFeedbackRating", "")),
                               "판매자 페이지": f"https://www.amazon.com/sp?seller={o.get('SellerId', '')}"})
        rows.append({"상태": st, "ASIN": a, "SKU": it.get("sku", ""), "상품명": it.get("title", ""),
                     "내 가격": round(my_p, 2) if my_p else "", "Buy Box 가격": bb_price or "",
                     "경쟁 가격 기준": thr or "", "오퍼 수": summ.get("TotalOfferCount", len(offers)),
                     "다른 판매자": len(rest), "메모": " · ".join(memo),
                     "_sellers": sorted(now_sellers)})
    def rank(r):
        st = r["상태"]
        if st == "Buy Box 잃음":
            return 0
        if st == "Buy Box 없음":
            return 1
        if st.startswith("정상") and r.get("다른 판매자"):
            return 2
        return {"내 오퍼 안 보임": 3, "확인 실패": 5}.get(st, 4)
    rows.sort(key=lambda r: (rank(r), -(r.get("다른 판매자") or 0)))
    alerts.sort(key=lambda x: x[0])
    return rows, others, alerts


def run(open_report: bool = False) -> Optional[Path]:
    import sp_api
    import daily_bus as bus
    cfg = sp_api.load_config()
    if not sp_api.ready(cfg):
        print("❌ SP-API 설정이 필요합니다(⚙ 설정). Buy Box 확인은 API로만 합니다.")
        return None
    info = my_asins()
    if not info:
        print("❌ 판매 중인 ASIN 목록이 없습니다. ‘전체 상품 재고’를 먼저 실행하세요.")
        return None
    print(f"【1】 Buy Box·합승 확인: 판매 중 ASIN {len(info)}개(20개씩, 10초 간격 · 약 {len(info) // 20 * 10 + 10}초)")
    c = sp_api.Client(cfg)
    try:
        data = fetch(c, sorted(info))
    except Exception as e:
        msg = str(e)
        if getattr(e, "status", 0) == 403:
            msg = "권한 없음 — 개발자 앱에 ‘Pricing’ 역할을 추가하고 다시 Authorize → 새 Refresh Token을 ⚙ 설정에 넣으세요"
        print(f"❌ Buy Box 확인 실패: {msg[:300]}")
        sp_api._src("Buy Box 확인", "API 실패", msg[:200])
        return None
    sp_api._src("Buy Box 확인", "API", f"ASIN {len(data)}개")
    today = date.today()
    prev = {}
    for i in range(1, 8):
        pf = DATA_DIR / f"{(today - timedelta(days=i)).isoformat()}.json"
        if pf.exists():
            prev = json.loads(pf.read_text(encoding="utf-8"))
            break
    rows, others, alerts = analyze(data, info, sp_api.seller_id(), prev)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / f"{today.isoformat()}.json").write_text(json.dumps(
        {r["ASIN"]: {"status": r["상태"], "sellers": r.get("_sellers", []), "my_price": r.get("내 가격")} for r in rows},
        ensure_ascii=False), encoding="utf-8")
    from collections import Counter
    cnt = Counter(r["상태"] for r in rows)
    kpis = {"확인한 ASIN": len(rows), "Buy Box 잃음": cnt.get("Buy Box 잃음", 0), "Buy Box 없음": cnt.get("Buy Box 없음", 0),
            "합승 판매자 있는 ASIN": sum(1 for r in rows if r.get("다른 판매자")),
            "새 합승 판매자": sum(1 for a in alerts if a[1] == "새 합승 판매자")}
    print("【2】 결과: " + " · ".join(f"{k} {v}" for k, v in kpis.items()))
    out = OUT_DIR / f"buybox_{today.isoformat()}.xlsx"
    write(out, rows, others, alerts, kpis)
    bus.record("Buy Box", kpis, alerts, report=out)
    print(f"✅ Buy Box·합승 확인 → {out}")
    if open_report:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass
    return out


def write(out: Path, rows, others, alerts, kpis) -> None:
    import inventory as inv
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    F = inv.FONT
    ws.cell(1, 1, f"Buy Box · 합승 판매자 확인 · {date.today().isoformat()}").font = Font(name=F, size=15, bold=True, color="1F3864")
    r = 3
    for k, v in kpis.items():
        ws.cell(r, 1, k).font = Font(name=F, bold=True)
        ws.cell(r, 2, v).font = Font(name=F)
        r += 1
    r += 1
    ws.cell(r, 1, "조치 필요").font = Font(name=F, size=12, bold=True, color="1F3864")
    fills = {"P1": PatternFill("solid", fgColor="FFE8CC"), "P2": PatternFill("solid", fgColor="FFF6D6")}
    for a in alerts:
        r += 1
        for i, v in enumerate(a, 1):
            c = ws.cell(r, i, v)
            c.font = Font(name=F)
            if a[0] in fills:
                c.fill = fills[a[0]]
    if not alerts:
        ws.cell(r + 1, 1, "문제 없음 — 모든 ASIN에서 내가 Buy Box를 가지고 있고 다른 판매자가 없습니다.").font = Font(name=F)
    for col, w in zip("ABCD", (16, 18, 80, 60)):
        ws.column_dimensions[col].width = w
    s2 = wb.create_sheet("ASIN별 상태")
    head = ["상태", "ASIN", "SKU", "상품명", "내 가격", "Buy Box 가격", "경쟁 가격 기준", "오퍼 수", "다른 판매자", "메모"]
    inv.write_table_sheet(s2, head, rows, {"상품명": 40, "메모": 60})
    s3 = wb.create_sheet("다른 판매자")
    if others:
        inv.write_table_sheet(s3, list(others[0].keys()), others, {"판매자 페이지": 45})
    else:
        s3.append(["다른 판매자가 없습니다."])
    s4 = wb.create_sheet("안내")
    for line in ("Buy Box 잃음: 다른 판매자가 Buy Box(장바구니 버튼)를 가져감 → 가격·재고·배송 기간을 비교하세요.",
                 "합승 판매자: 같은 ASIN에 다른 판매자가 오퍼를 올림. 위조·상표 침해라면 Brand Registry → Report a violation 으로 신고.",
                 "경쟁 가격 기준(Competitive Price Threshold): 이보다 비싸면 Amazon이 Buy Box를 숨길 수 있음.",
                 "판매자 ID는 판매자 페이지(amazon.com/sp?seller=ID)에서 이름을 확인할 수 있습니다.",
                 "어제 기록과 비교해 새로 나타난 판매자는 P1로 표시합니다. 조회만 하며 가격을 바꾸지 않습니다."):
        s4.append([line])
    s4.column_dimensions["A"].width = 130
    wb.save(out)


def status_text() -> str:
    files = sorted(OUT_DIR.glob("buybox_*.xlsx")) if OUT_DIR.exists() else []
    return f"마지막 확인: {files[-1].stem.replace('buybox_', '')}" if files else "아직 실행하지 않음(SP-API Pricing 역할 필요)"

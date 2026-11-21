"""
재무·보상 점검(SP-API, 조회만) — 받아야 할 돈을 놓치지 않게

1. 창고 분실·파손 미보상: 재고 원장(Inventory Ledger, 18개월)의 분실·파손 수량 − 다시 찾은 수량(앞뒤 상관없이) − 보상받은 수량 > 0
   (같은 FNSKU의 ‘찾음’은 분실과 상계. 단 예전에 보상받은 분실을 찾은 것(보상 되돌림)은 상계하지 않음)
   (발생 후 30일 지난 것만 — Amazon이 자동 보상할 시간을 줌, 신청 기한 = 발생일 + 60일)
2. 환불했는데 반품도 보상도 없음: 환불(Finances API) 후 60일이 지났는데 반품 보고서·보상 보고서에 그 주문이 없음
3. FBA 수수료·크기 변화: FBA 수수료 미리 보기 보고서를 지난번과 비교(크기 등급·무게·배송비 10% 이상 변화 → 측정 오류일 수 있음)
4. 최근 90일 받은 보상 요약
필요 권한: 재고·보상·수수료 보고서 = Amazon Fulfillment(FBA 재고 보고서가 되면 있음), 환불 = Finance and Accounting
결과: reports/finance/finance_날짜.xlsx · 신청은 사람이 직접(자동 신청하지 않음)
"""

import json
import os
import sys
import time
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "reports" / "finance"
DATA_DIR = OUT_DIR / "data"
# Amazon 재고 조정 사유 코드(-: 줄어듦, +: 늘어남). D(폐기)·O(수정)·P/Q(상태 변경)·3/4(상품 재정의)는 분실이 아님
LOSS = {"M": "분실(창고, misplaced)", "5": "분실(창고)", "E": "파손(창고)", "H": "파손(창고)", "K": "파손(창고)",
        "U": "파손(창고)", "6": "파손(창고)", "7": "파손(창고)"}
FOUND = {"F", "N"}


def _reason(v) -> str:
    """보고서는 코드(M), 화면은 글(Inventory misplaced) — 둘 다 코드로."""
    t = str(v or "").strip()
    if len(t) <= 2:
        return t.upper()
    low = t.lower()
    if "found" in low:
        return "F"
    if "misplac" in low or "lost" in low:
        return "M"
    if "damag" in low:
        return "E"
    return t.upper()
WH_REIMB = ("lost_warehouse", "damaged_warehouse", "lost_inbound", "damaged_inbound", "lost", "damaged")
CLAIM_PATH = ("Seller Central 오른쪽 위 ? → Get help(Support resources → Get help and resources) → 검색 ‘lost and damaged inventory’ "
              "→ Fulfillment by Amazon → Lost and damaged inventory(FNSKU·Reference ID·날짜·수량 입력)")


def _g(r: Dict, *names, default=""):
    low = {str(k).strip().lower().replace("_", "-").replace(" ", "-"): v for k, v in r.items()}
    for n in names:
        v = low.get(n.lower().replace("_", "-").replace(" ", "-"))
        if v not in (None, ""):
            return v
    return default


def _num(v) -> float:
    try:
        return float(str(v).replace(",", "").replace("$", "") or 0)
    except ValueError:
        return 0.0


def _d(v) -> Optional[date]:
    s = str(v or "").strip()[:10]
    for f in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def _report(c, rtype: str, days: int, tag: str):
    import inventory as inv
    end = datetime.now(timezone.utc) - timedelta(minutes=5)
    start = end - timedelta(days=days)
    p = c.report(rtype, start, end, save_as=f"{tag}_{date.today()}")
    return inv.read_report(p)[1] if p else []


# ================================================================ 1. 분실·파손 미보상

def lost_damaged(ledger: List[Dict], reimb: List[Dict], today: date) -> Tuple[List[Dict], List[Dict]]:
    """FNSKU별로 분실·파손(−)과 찾음(+)·보상을 상계. 찾음·보상은 남아 있는 가장 오래된 분실부터 지움
    (찾음이 분실보다 먼저 있어도 같은 FNSKU면 상계 — Amazon 재고 조정 방식과 같게)."""
    ev = defaultdict(lambda: {"events": [], "found": 0.0, "found_rows": [], "asin": "", "sku": "", "title": ""})
    for r in ledger:
        if "adjust" not in str(_g(r, "event-type", "Event Type")).lower():
            continue
        fn = str(_g(r, "fnsku", "FNSKU")).strip()
        q = _num(_g(r, "quantity", "Quantity"))
        rs = _reason(_g(r, "reason", "Reason"))
        d = _d(_g(r, "date", "Date", "date-and-time", "Date and Time"))
        e = ev[fn]
        # Amazon이 분실·찾음을 서로 맞춰 둔 결과(보고서에 열이 있으면 그대로 사용 — 화면에서 펼치면 보이는 ‘Inventory found’ 연결)
        unrec_key = next((k for k in r if "unreconciled" in str(k).lower()), None)
        e["asin"], e["sku"] = e["asin"] or _g(r, "asin", "ASIN"), e["sku"] or _g(r, "msku", "MSKU", "sku")
        e["title"] = e["title"] or str(_g(r, "title", "Title"))[:60]
        if q < 0 and rs in LOSS and d:
            if unrec_key is not None:
                e["amazon_rec"] = True
                left_q = abs(_num(r.get(unrec_key)))
                if left_q <= 0:
                    e["matched"] = e.get("matched", 0) + (-q)
                    continue                 # Amazon이 이미 찾음(또는 보상)과 연결함
                q = -left_q
            e["events"].append({"date": d, "qty": -q, "why": LOSS[rs], "ref": str(_g(r, "reference-id", "Reference ID")),
                                "fc": str(_g(r, "fulfillment-center", "Fulfillment Center"))})
        elif q > 0 and rs in FOUND:
            if unrec_key is not None:
                e["amazon_rec"] = True
                left_q = abs(_num(r.get(unrec_key)))   # 아직 어떤 분실과도 연결되지 않은 찾음만 상계에 씀
                e["found_rows"].append(f"{d} +{q:.0f}" + ("" if left_q else "(분실과 연결됨)"))
                e["found"] += left_q
                continue
            e["found"] += q
            e["found_rows"].append(f"{d} +{q:.0f}")
    paid, rev = defaultdict(float), defaultdict(float)
    received = []
    for r in reimb:
        fn = str(_g(r, "fnsku")).strip()
        why = str(_g(r, "reason")).strip()
        qty = _num(_g(r, "quantity-reimbursed-total")) or (_num(_g(r, "quantity-reimbursed-cash")) + _num(_g(r, "quantity-reimbursed-inventory")))
        if "reversal" in why.lower():          # 보상받은 분실분을 나중에 찾아서 보상을 되돌림 → 그 ‘찾음’은 새 분실과 상계하면 안 됨
            rev[fn] += abs(qty) or 1
        elif why.lower() in WH_REIMB or "warehouse" in why.lower() or "inbound" in why.lower():
            paid[fn] += qty
        received.append({"승인일": str(_g(r, "approval-date"))[:10], "사유": why, "SKU": _g(r, "sku"), "FNSKU": fn,
                         "ASIN": _g(r, "asin"), "수량": qty, "금액($)": round(_num(_g(r, "amount-total")), 2),
                         "주문": _g(r, "amazon-order-id"), "케이스 ID": _g(r, "case-id"), "보상 ID": _g(r, "reimbursement-id")})
    out = []
    for fn, e in ev.items():
        lost = sum(x["qty"] for x in e["events"])
        credit = max(0.0, e["found"] - rev.get(fn, 0)) + paid.get(fn, 0)
        left = []
        for x in sorted(e["events"], key=lambda x: x["date"]):      # 오래된 분실부터 찾음·보상으로 지움
            use = min(credit, x["qty"])
            credit -= use
            if x["qty"] - use > 0:
                left.append(dict(x, qty=x["qty"] - use))
        net = sum(x["qty"] for x in left)
        if net <= 0:
            continue
        age = lambda x: (today - x["date"]).days
        can = [x for x in left if 30 <= age(x) <= 60]
        wait = [x for x in left if age(x) < 30]
        show = can or wait or left
        st = "신청 가능" if can else ("대기(Amazon 자동 보상 기다림)" if wait else "기한 지남(60일)")
        old_n = sum(x["qty"] for x in left if age(x) > 60)
        out.append({"상태": st, "FNSKU": fn, "SKU": e["sku"], "ASIN": e["asin"], "상품명": e["title"],
                    "분실·파손(18개월)": lost, "다시 찾음(18개월)": e["found"], "보상받음": paid.get(fn, 0), "미보상 수량": net,
                    "신청 가능 수량": sum(x["qty"] for x in can),
                    "사유": ", ".join(sorted({x["why"] for x in show})),
                    "발생일": ", ".join(x["date"].isoformat() for x in show[:5]),
                    "신청 기한(추정)": (min(x["date"] for x in show) + timedelta(days=60)).isoformat(),
                    "Reference ID": ", ".join(x["ref"] for x in show[:5] if x["ref"]),
                    "물류 센터": ", ".join(sorted({x["fc"] for x in show if x["fc"]})),
                    "찾음 기록": ", ".join(e["found_rows"][:5]) + (f" (그중 {rev[fn]:.0f}개는 예전 보상 되돌림)" if rev.get(fn) else ""),
                    "판단 방법": "Amazon 연결 정보(Unreconciled) 사용" if e.get("amazon_rec") else "분실 − 찾음 − 보상 상계(추정)",
                    "확인 방법": f"Inventory Ledger → Detailed view → FNSKU {fn} · Event type Adjustments · Last 365 days → "
                                 "분실 줄의 Reason 옆 ^ 를 펼쳐 ‘Inventory found’가 붙어 있으면 이미 찾은 것(신청 불필요)",
                    "신청 문구(영어)": (f"Hello, {sum(x['qty'] for x in can):.0f} unit(s) of FNSKU {fn} (SKU {e['sku']}, ASIN {e['asin']}) "
                                     f"were marked as {'misplaced/lost' if any('분실' in x['why'] for x in can) else 'damaged'} in the fulfillment center "
                                     + "; ".join(f"on {x['date']} (Reference ID {x['ref']}, FC {x['fc']}, qty {x['qty']:.0f})" for x in can[:5])
                                     + ". The Inventory Ledger shows no matching 'Inventory found' adjustment and there is no reimbursement "
                                     "in my Reimbursements report. Please investigate and reimburse. Thank you.") if can else "",
                    "메모": f"60일 지난 미보상 {old_n:.0f}개는 신청 기한이 지났을 수 있음" if old_n and st != "기한 지남(60일)" else ""})
    out.sort(key=lambda r: ({"신청 가능": 0, "대기(Amazon 자동 보상 기다림)": 1}.get(r["상태"], 2), -r["미보상 수량"]))
    return out, received


# ================================================================ 2. 환불했는데 반품·보상 없음

def refunds(c, days: int = 150) -> List[Dict]:
    """Finances API: 환불 이벤트(주문, 날짜, SKU, 금액)."""
    after = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    before = (datetime.now(timezone.utc) - timedelta(minutes=5)).strftime("%Y-%m-%dT%H:%M:%SZ")
    out, token, n = [], None, 0
    while True:
        params = {"MaxResultsPerPage": 100}
        if token:
            params["NextToken"] = token
        else:
            params.update({"PostedAfter": after, "PostedBefore": before})
        d = c.get("/finances/v0/financialEvents", **params)
        pl = d.get("payload") or {}
        for e in (pl.get("FinancialEvents") or {}).get("RefundEventList") or []:
            for it in e.get("ShipmentItemAdjustmentList") or []:
                amt = sum(_num((x.get("ChargeAmount") or {}).get("CurrencyAmount")) for x in it.get("ItemChargeAdjustmentList") or []
                          if x.get("ChargeType") == "Principal")
                out.append({"order": e.get("AmazonOrderId", ""), "date": _d(e.get("PostedDate")), "sku": it.get("SellerSKU", ""),
                            "qty": abs(_num(it.get("QuantityShipped"))) or 1, "amount": abs(amt)})
        token = pl.get("NextToken")
        n += 1
        if not token or n >= 60:
            break
        time.sleep(2.1)                  # 이 API는 초당 0.5번
    return out


def missing_returns(refs: List[Dict], returns: List[Dict], reimb: List[Dict], today: date) -> List[Dict]:
    returned = {str(_g(r, "order-id")).strip() for r in returns}
    paid = {str(_g(r, "amazon-order-id")).strip() for r in reimb}
    out = []
    for r in refs:
        if not r["date"] or not r["order"]:
            continue
        age = (today - r["date"]).days
        if age < 60 or r["order"] in returned or r["order"] in paid:
            continue
        out.append({"상태": "확인 필요", "주문": r["order"], "SKU": r["sku"], "환불일": r["date"].isoformat(), "경과일": age,
                    "수량": r["qty"], "환불 금액($)": round(r["amount"], 2),
                    "메모": "환불 후 60일 지났는데 반품 입고·보상 기록 없음(FBA 주문일 때만 해당)"})
    out.sort(key=lambda x: -x["환불 금액($)"])
    return out


# ================================================================ 3. 수수료·크기 변화

def fee_changes(fees: List[Dict], prev: Dict[str, Dict]) -> Tuple[List[Dict], Dict[str, Dict]]:
    snap, out = {}, []
    for r in fees:
        sku = str(_g(r, "sku")).strip()
        if not sku:
            continue
        cur = {"tier": str(_g(r, "product-size-tier")), "weight": _num(_g(r, "item-package-weight")),
               "long": _num(_g(r, "longest-side")), "fee": _num(_g(r, "expected-fulfillment-fee-per-unit")),
               "asin": str(_g(r, "asin")), "name": str(_g(r, "product-name"))[:60], "price": _num(_g(r, "your-price", "sales-price"))}
        snap[sku] = cur
        p = prev.get(sku)
        if not p:
            continue
        why = []
        if p.get("tier") and cur["tier"] and p["tier"] != cur["tier"]:
            why.append(f"크기 등급 {p['tier']} → {cur['tier']}")
        if p.get("fee") and cur["fee"] and cur["fee"] > p["fee"] * 1.1:
            why.append(f"FBA 배송비 ${p['fee']:.2f} → ${cur['fee']:.2f}")
        if p.get("weight") and cur["weight"] and abs(cur["weight"] - p["weight"]) / p["weight"] > 0.15:
            why.append(f"무게 {p['weight']} → {cur['weight']}")
        if p.get("long") and cur["long"] and abs(cur["long"] - p["long"]) / p["long"] > 0.15:
            why.append(f"가장 긴 변 {p['long']} → {cur['long']}")
        if why:
            out.append({"SKU": sku, "ASIN": cur["asin"], "상품명": cur["name"], "변화": " · ".join(why),
                        "지금 크기 등급": cur["tier"], "지금 FBA 배송비($)": cur["fee"],
                        "확인 방법": "실제 포장 크기·무게와 비교 → 다르면 ? → Get help → 검색 ‘FBA product dimensions’ 재측정 요청"})
    return out, snap


# ================================================================ 실행

def run(open_report: bool = False) -> Optional[Path]:
    import sp_api
    import daily_bus as bus
    cfg = sp_api.load_config()
    if not sp_api.ready(cfg):
        print("❌ SP-API 설정이 필요합니다(⚙ 설정). 재무·보상 점검은 API로만 합니다.")
        return None
    c = sp_api.Client(cfg)
    today = date.today()
    src, data = {}, {}

    def get(label, fn):
        try:
            print(f"   ▶ {label}...")
            data[label] = fn()
            src[label] = f"SP-API {len(data[label])}행"
        except Exception as e:
            data[label] = []
            st = getattr(e, "status", 0)
            src[label] = ("권한 없음(" + ("Finance and Accounting" if "환불" in label else "Amazon Fulfillment") + " 역할 필요)"
                          if st == 403 else f"실패: {str(e)[:150]}")
            print(f"   ⚠ {label}: {src[label]}")
        sp_api._src(f"재무·보상 {label}", "API" if not src[label].startswith(("권한", "실패")) else "API 실패", src[label])

    print("【1】 데이터 받기(보고서마다 1~5분)")
    get("재고 원장(18개월)", lambda: _report(c, "GET_LEDGER_DETAIL_VIEW_DATA", 540, "ledger"))
    get("보상 내역(18개월)", lambda: _report(c, "GET_FBA_REIMBURSEMENTS_DATA", 540, "reimbursements"))
    get("반품(180일)", lambda: _report(c, "GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA", 180, "returns_long"))
    get("FBA 수수료 미리 보기", lambda: _report(c, "GET_FBA_ESTIMATED_FBA_FEES_TXT_DATA", 30, "fba_fees"))
    get("환불(150일)", lambda: refunds(c))

    print("\n【2】 점검")
    lost, received = lost_damaged(data["재고 원장(18개월)"], data["보상 내역(18개월)"], today)
    miss = missing_returns(data["환불(150일)"], data["반품(180일)"], data["보상 내역(18개월)"], today) if data["환불(150일)"] else []
    prev, prev_day = {}, ""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    olds = sorted(DATA_DIR.glob("fees_*.json"))
    if olds:
        prev, prev_day = json.loads(olds[-1].read_text(encoding="utf-8")), olds[-1].stem.replace("fees_", "")
    fee_rows, snap = fee_changes(data["FBA 수수료 미리 보기"], prev)
    if snap:
        (DATA_DIR / f"fees_{today}.json").write_text(json.dumps(snap, ensure_ascii=False), encoding="utf-8")
    recent = [r for r in received if r["승인일"] >= (today - timedelta(days=90)).isoformat()]
    can = [r for r in lost if r["상태"] == "신청 가능"]
    alerts = []
    if can:
        alerts.append(("P1", "분실·파손 보상 신청 가능", f"{len(can)}개 FNSKU, 신청 가능 {sum(r['신청 가능 수량'] for r in can):.0f}개 "
                       f"(가장 빠른 기한 {min(r['신청 기한(추정)'] for r in can)})", CLAIM_PATH))
    for r in can[:10]:
        alerts.append(("P2", "분실·파손 미보상", f"{r['SKU'] or r['FNSKU']}: {r['신청 가능 수량']:.0f}개({r['사유']}, {r['발생일']}) · 기한 {r['신청 기한(추정)']}",
                       f"Reference ID {r['Reference ID'][:60]}"))
    if miss:
        alerts.append(("P2", "환불 후 반품·보상 없음", f"{len(miss)}건, ${sum(r['환불 금액($)'] for r in miss):,.2f}",
                       "? → Get help → 검색 ‘customer return refund reimbursement’ → 주문 번호로 문의"))
    for r in fee_rows[:10]:
        alerts.append(("P2", "FBA 수수료·크기 변화", f"{r['SKU']}: {r['변화']}", r["확인 방법"]))
    kpis = {"신청 가능 분실·파손(FNSKU)": len(can), "미보상 수량(신청 가능)": sum(r["신청 가능 수량"] for r in can),
            "대기 중(30일 미만)": sum(1 for r in lost if r["상태"].startswith("대기")),
            "환불 후 반품·보상 없음": len(miss), "환불 후 반품·보상 없음 금액($)": round(sum(r["환불 금액($)"] for r in miss), 2),
            "수수료·크기 변화 SKU": len(fee_rows), "최근 90일 받은 보상($)": round(sum(r["금액($)"] for r in recent), 2),
            "수수료 비교 기준": prev_day or "첫 실행(다음부터 비교)"}
    print("   " + " · ".join(f"{k} {v}" for k, v in kpis.items()))
    out = OUT_DIR / f"finance_{today}.xlsx"
    write(out, kpis, alerts, lost, miss, fee_rows, recent, src)
    bus.record("재무·보상", kpis, alerts, report=out)
    print(f"✅ 재무·보상 점검 → {out}")
    if open_report:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass
    return out


def write(out: Path, kpis, alerts, lost, miss, fee_rows, recent, src) -> None:
    import inventory as inv
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    F = inv.FONT
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.cell(1, 1, f"재무·보상 점검 · {date.today()}").font = Font(name=F, size=15, bold=True, color="1F3864")
    ws.cell(2, 1, "받아야 할 보상을 찾는 목록입니다(추정). 신청은 직접 하며, 이 프로그램은 신청하지 않습니다.").font = Font(name=F, color="6B6B6B")
    r = 4
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
        ws.cell(r + 1, 1, "찾은 문제가 없습니다.").font = Font(name=F)
    r += 2
    ws.cell(r, 1, "데이터 출처").font = Font(name=F, size=12, bold=True, color="1F3864")
    for k, v in src.items():
        r += 1
        ws.cell(r, 1, k).font = Font(name=F)
        ws.cell(r, 2, v).font = Font(name=F, color="6B6B6B")
    for col, w in zip("ABCD", (30, 26, 70, 70)):
        ws.column_dimensions[col].width = w

    def sheet(name, rows, empty, widths=None):
        s = wb.create_sheet(name)
        if rows:
            inv.write_table_sheet(s, list(rows[0].keys()), rows, widths or {})
        else:
            s.append([empty])
            s.column_dimensions["A"].width = 100
    sheet("분실·파손 미보상", lost, "미보상 분실·파손이 없습니다(또는 재고 원장 권한 없음).",
          {"상품명": 36, "Reference ID": 40, "확인 방법": 60, "신청 문구(영어)": 80, "찾음 기록": 40})
    sheet("환불 후 반품 없음", miss, "해당 없음(또는 Finance and Accounting 권한 없음).", {"메모": 50})
    sheet("수수료·크기 변화", fee_rows, "지난번과 달라진 SKU가 없습니다(첫 실행이면 다음 실행부터 비교).", {"변화": 50, "확인 방법": 70})
    sheet("받은 보상(90일)", recent, "최근 90일 보상 기록이 없습니다.")
    g = wb.create_sheet("신청 방법")
    for line in ("분실·파손: " + CLAIM_PATH,
                 "  확인: Supply chain → Supply chain reports → FBA inventory reports → Inventory Ledger → Detailed view "
                 "(FNSKU 입력, Event type ‘Adjustments’) — Reason M·E 등(-) 뒤에 F(+, 찾음)가 없는지",
                 "  받은 보상: Supply chain → Supply chain reports → FBA inventory reports → 왼쪽 Payments → Reimbursements (FNSKU로 검색)",
                 "  기준: 창고 분실·파손은 대부분 자동 보상되므로 발생 30일 뒤에도 보상이 없을 때 신청(발생 후 60일 안)",
                 "  신청 내용: FNSKU · SKU · ASIN · Reference ID · 발생일 · 수량 · Reason 코드 · 물류 센터(Fulfillment Center)",
                 "환불 후 반품 없음: ? → Get help → 검색 ‘customer return refund reimbursement’ → 주문 번호 입력",
                 "  확인: Orders → Order reports → FBA customer return reports(주문 번호) · 위의 Reimbursements",
                 "  (Amazon은 보통 환불 후 반품이 오지 않으면 자동 보상 — 그래도 기록이 없을 때만 문의)",
                 "수수료·크기: 실제 포장 크기·무게를 재서 다르면 ? → Get help → 검색 ‘FBA product dimensions’(재측정 요청)",
                 "  확인: Supply chain reports → FBA inventory reports → Fee Preview"):
        g.append([line])
    g.column_dimensions["A"].width = 140
    wb.save(out)


def run_weekly(open_report: bool = False) -> Optional[Path]:
    """매일 자동 실행용: 월요일이거나 마지막 점검이 7일 넘었을 때만."""
    files = sorted(OUT_DIR.glob("finance_*.xlsx")) if OUT_DIR.exists() else []
    last = _d(files[-1].stem.replace("finance_", "")) if files else None
    if last and date.today().weekday() != 0 and (date.today() - last).days < 7:
        print(f"   재무·보상 점검은 주 1회(월요일) — 마지막 점검 {last}. 이번에는 건너뜁니다.")
        return None
    return run(open_report)


def status_text() -> str:
    files = sorted(OUT_DIR.glob("finance_*.xlsx")) if OUT_DIR.exists() else []
    return f"마지막 점검: {files[-1].stem.replace('finance_', '')}" if files else "아직 실행하지 않음"

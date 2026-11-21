"""
FBA 재입고(보충) 제안 — 얼마나, 언제까지 보내야 하는지

하루 판매량 = 최근 30일 × 0.6 + 최근 90일 × 0.4 (최근 추세를 더 반영)
재주문 시점 = 하루 판매량 × (리드타임 + 안전 재고 일수)
  가지고 있는 재고(판매 가능 + 입고 중)가 재주문 시점 이하 → 지금 보내야 함
보낼 수량 = 하루 판매량 × (리드타임 + 안전 재고 일수 + 목표 보유 일수) − 가지고 있는 재고
보낼 날짜 = 재고가 재주문 시점에 닿는 날

설정: reports/inventory/restock_config.xlsx (처음 실행 때 자동 생성)
  ‘기본값’ 시트: 리드타임(생산+운송+입고), 안전 재고 일수, 목표 보유 일수
  ‘SKU별’ 시트: SKU마다 다르게 하고 싶을 때(빈칸은 기본값), 제외=Y 이면 제안하지 않음
제안만 하며, 배송을 만들거나 수량을 바꾸지 않습니다.
"""

from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List

BASE = Path(__file__).resolve().parent.parent
CONFIG = BASE / "reports" / "inventory" / "restock_config.xlsx"
DEFAULTS = {"리드타임(일)": 45, "안전 재고 일수": 14, "목표 보유 일수": 60}
MIN_DAILY = 0.1          # 하루 0.1개(한 달 3개) 미만은 참고만


def load_config() -> Dict:
    from openpyxl import Workbook, load_workbook
    if not CONFIG.exists():
        CONFIG.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "기본값"
        ws.append(["항목", "값", "설명"])
        ws.append(["리드타임(일)", DEFAULTS["리드타임(일)"], "주문(생산)부터 Amazon 창고 입고·판매 가능까지 걸리는 날 수"])
        ws.append(["안전 재고 일수", DEFAULTS["안전 재고 일수"], "판매가 갑자기 늘거나 입고가 늦어질 때를 대비한 여유분"])
        ws.append(["목표 보유 일수", DEFAULTS["목표 보유 일수"], "한 번 보낼 때 몇 일치를 보낼지(많을수록 보관료·장기 재고 위험 증가)"])
        for col, w in zip("ABC", (18, 10, 70)):
            ws.column_dimensions[col].width = w
        s2 = wb.create_sheet("SKU별")
        s2.append(["SKU", "리드타임(일)", "안전 재고 일수", "목표 보유 일수", "제외(Y)", "메모"])
        for col, w in zip("ABCDEF", (22, 12, 14, 14, 10, 40)):
            s2.column_dimensions[col].width = w
        wb.save(CONFIG)
    wb = load_workbook(CONFIG, data_only=True)
    base = dict(DEFAULTS)
    for r in wb["기본값"].iter_rows(min_row=2, values_only=True):
        if r and r[0] in base and isinstance(r[1], (int, float)):
            base[r[0]] = float(r[1])
    per = {}
    if "SKU별" in wb.sheetnames:
        for r in wb["SKU별"].iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                per[str(r[0]).strip()] = {"lead": r[1], "safety": r[2], "cover": r[3],
                                          "skip": str(r[4] or "").strip().upper() in ("Y", "YES", "O", "1")}
    return {"base": base, "per": per}


def plan(items: List[Dict], today: date = None) -> Dict:
    """build_rows 결과(FBA 재고) → {"kpis", "rows", "alerts"} (sc_pages 결과와 같은 형식)."""
    today = today or date.today()
    cfg = load_config()
    b = cfg["base"]
    rows, alerts = [], []
    for it in items:
        if not it.get("in_fba"):
            continue
        sku = it["sku"]
        o = cfg["per"].get(sku, {})
        if o.get("skip"):
            continue
        lead = float(o.get("lead") or b["리드타임(일)"])
        safety = float(o.get("safety") or b["안전 재고 일수"])
        cover = float(o.get("cover") or b["목표 보유 일수"])
        t30, t90 = it.get("t30"), it.get("t90")
        if t30 is None and t90 is None:
            continue
        daily = 0.6 * (t30 or 0) / 30 + 0.4 * (t90 or 0) / 90 if t90 is not None else (t30 or 0) / 30
        if daily <= 0:
            continue
        have = (it.get("available") or 0) + (it.get("inbound") or 0) + (it.get("reserved") or 0) * 0.5
        rop = daily * (lead + safety)
        days_left = have / daily
        qty = max(0, round(daily * (lead + safety + cover) - have))
        send_in = (have - rop) / daily
        if send_in > 30:          # 한 달 넘게 여유 → 목록에서 제외
            continue
        st = str(it.get("status") or "")
        if daily < MIN_DAILY:
            urg, lv = "판매 적음(필요하면 소량)", "P3"
        elif send_in <= 0:
            urg, lv = ("품절 위험(리드타임 안에 소진)" if days_left < lead else "지금 보내야 함"), ("P1" if days_left < lead else "P2")
        elif send_in <= 14:
            urg, lv = f"{send_in:.0f}일 안에 보내기", "P2"
        else:
            urg, lv = f"{send_in:.0f}일 후 보내기", "P3"
        rows.append({"긴급도": urg, "SKU": sku, "ASIN": it.get("asin", ""), "상품명": str(it.get("title", ""))[:70],
                     "하루 판매량": round(daily, 2), "판매 가능": it.get("available") or 0, "입고 중": it.get("inbound") or 0,
                     "남은 일수": round(days_left, 1), "보낼 날짜": (today + timedelta(days=max(0, send_in))).isoformat(),
                     "제안 수량": qty, "리드타임": lead, "안전 일수": safety, "목표 일수": cover, "상태": st,
                     "_sort": (0 if lv == "P1" else 1 if lv == "P2" else 2, -daily)})
        if lv in ("P1", "P2"):
            alerts.append((lv, "재입고", f"{sku}: {urg} · {qty}개 제안(남은 {days_left:.0f}일, 하루 {daily:.1f}개)",
                           "Send to Amazon으로 입고 계획을 만드세요(자동으로 만들지 않음)"))
    rows.sort(key=lambda r: r["_sort"])
    for r in rows:
        r.pop("_sort", None)
    urgent = sum(1 for r in rows if r["긴급도"].startswith(("품절", "지금")))
    rows = [r for r in rows]
    kpis = {"보충 제안 SKU": len(rows), "지금 보내야 하는 SKU": urgent, "제안 수량 합계": sum(r["제안 수량"] for r in rows),
            "기준": f"리드타임 {b['리드타임(일)']:.0f}일 · 안전 {b['안전 재고 일수']:.0f}일 · 목표 {b['목표 보유 일수']:.0f}일(설정: {CONFIG.name})"}
    alerts.sort(key=lambda a: a[0])
    return {"name": "보충 제안", "src": "FBA 재고 보고서로 계산", "kpis": kpis, "rows": rows, "alerts": alerts[:40]}

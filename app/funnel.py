"""
주간 ASIN 퍼널 진단 — 노출 → 클릭 → 장바구니 → 구매, 어느 단계에서 막히는지(조회만)

데이터(주 = 일~토, Amazon Brand Analytics 기준, 최근 8주)
  - Search Catalog Performance(SCP): ASIN별 검색 노출·클릭·장바구니·구매 — SP-API 우선, 없으면 광고 기능이 받아 둔 파일
  - 비즈니스 보고서(Sales & Traffic): ASIN별 세션·세션 전환율·Buy Box 비율 — SP-API(없으면 생략)
  받은 주는 reports/funnel/data/ 에 저장해 다시 받지 않음(매주 새 주 1개만 받음)

진단(이번 주 vs 지난 4주 평균, 데이터가 적은 ASIN은 판단하지 않음)
  노출 -30% ↓   → 키워드 순위·광고 노출 확인
  클릭률 -25% ↓ → 메인 이미지·가격·평점(가격이 올랐으면 표시)
  장바구니율 -25% ↓ → 상세 페이지(불릿·A+·이미지)
  구매 전환율 -25% ↓, 클릭 30회 이상인데 구매 0 → 리뷰·가격·재고·배송 약속
  Buy Box 90% 미만 → 가격·재고·다른 판매자 확인
수정 효과: ‘리스팅·가격 수정’ 기록과 연결 — 반영 전 2주 vs 반영 후 2주 비교
"""

import json
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "reports" / "funnel"
DATA_DIR = OUT_DIR / "data"
WEEKS = 8
DROP = {"impressions": 0.30, "ctr": 0.25, "cart_rate": 0.25, "cvr": 0.25}
MIN_IMP, MIN_CLICKS, MIN_CARTS = 500, 30, 10
BUYBOX_MIN = 90.0


# ================================================================ 데이터

def weeks(n: int = WEEKS) -> List[str]:
    import ba_collect as ba
    last = date.fromisoformat(ba.expected_week())
    return [(last - timedelta(days=7 * i)).isoformat() for i in range(n)][::-1]       # 오래된 주 → 최근 주


def _scp_file(week: str) -> Tuple[Optional[Path], str]:
    import ba_collect as ba
    full = ba.BA_DIR / f"scp_{week}.xlsx"
    if full.exists():
        return full, "Brand Analytics 전체 데이터"
    page = ba.BA_DIR / f"scp_{week}_page.xlsx"
    if page.exists():
        return page, "Brand Analytics 화면 표(일부 ASIN만일 수 있음)"
    return None, ""


def load_scp(week: str, use_api: bool) -> Tuple[Dict[str, Dict], str]:
    """{asin: {impressions, clicks, cart_adds, purchases, sales, price, title}}"""
    import ads_common as ac
    import ba_collect as ba
    p, src = _scp_file(week)
    if p is None and use_api:
        import sp_api
        import sp_data
        rows = sp_api.try_api(f"퍼널 SCP {week}", sp_data.scp_week, week)
        if rows:
            ba.BA_DIR.mkdir(parents=True, exist_ok=True)
            p = ba.BA_DIR / f"scp_{week}.xlsx"
            ac.save_rows(rows, p, "scp")
            src = "SP-API"
    if p is None:
        return {}, "없음"
    out = {}
    for x in ba.parse_scp(ac.load_rows(p)):
        out[x["asin"]] = x
    return out, src


def load_st(week: str, use_api: bool) -> Tuple[Dict[str, Dict], str]:
    """비즈니스 보고서(주간, ASIN별) → {asin: {sessions, units, sales, usp, buybox}}"""
    import ads_common as ac
    f = DATA_DIR / f"st_{week}.json"
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8")), "저장된 주간 데이터"
    if not use_api:
        return {}, "없음"
    import sp_api
    import sp_data
    e = date.fromisoformat(week)
    res = sp_api.try_api(f"퍼널 비즈니스 보고서 {week}", sp_data.sales_traffic, e - timedelta(days=6), e)
    if not res:
        return {}, "없음(SP-API 권한 필요)"
    _, rows = res
    out = {}
    for r in rows:
        a = r.get("(Child) ASIN")
        if not a:
            continue
        out[a] = {"sessions": ac.num(r.get("Sessions - Total")), "units": ac.num(r.get("Units Ordered")),
                  "sales": ac.num(r.get("Ordered Product Sales")), "usp": ac.num(r.get("Unit Session Percentage")),
                  "buybox": ac.num(r.get("Featured Offer (Buy Box) Percentage")), "sku": r.get("SKU", "")}
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if e < date.today() - timedelta(days=3):                 # 끝난 주만 저장(늦게 집계되는 주문 반영)
        f.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    return out, "SP-API"


def sku_map() -> Dict[str, str]:
    import inventory as inv
    files = sorted((BASE_DIR / "reports" / "inventory").glob("all_listings_*.xlsx"))
    out = {}
    if files:
        try:
            for r in inv.read_xlsx(files[-1])[1]:
                a = str(inv._get(r, "asin1", "asin", default="")).strip()
                if a and a not in out:
                    out[a] = str(inv._get(r, "seller-sku", "sku", default="") or "")
        except Exception:
            pass
    return out


# ================================================================ 계산

def rates(x: Dict) -> Dict:
    imp, cl, ca, pu = (x.get(k, 0) or 0 for k in ("impressions", "clicks", "cart_adds", "purchases"))
    return {"impressions": imp, "clicks": cl, "cart_adds": ca, "purchases": pu,
            "ctr": cl / imp if imp else None, "cart_rate": ca / cl if cl else None, "cvr": pu / cl if cl else None,
            "price": x.get("price") or None, "sales": x.get("sales", 0)}


def _avg(vals: List[Optional[float]]) -> Optional[float]:
    v = [x for x in vals if x is not None]
    return sum(v) / len(v) if v else None


def _chg(cur: Optional[float], base: Optional[float]) -> Optional[float]:
    if cur is None or not base:
        return None
    return cur / base - 1


def diagnose(asin: str, hist: List[Dict], st_now: Dict) -> Tuple[List[Tuple], Dict]:
    """hist: 주별 rates(오래된 → 최근). → ([(등급, 분류, 내용, 조치)], 비교 값)"""
    cur, prev4 = hist[-1], [h for h in hist[-5:-1] if h.get("impressions")]
    base = {k: _avg([h.get(k) for h in prev4]) for k in ("impressions", "ctr", "cart_rate", "cvr", "price")}
    cmp_ = {f"{k}_chg": _chg(cur.get(k), base[k]) for k in ("impressions", "ctr", "cart_rate", "cvr")}
    cmp_["base"] = base
    out = []
    if len(prev4) < 2:
        return out, cmp_
    ci, cc = cmp_["impressions_chg"], cmp_["ctr_chg"]
    if base["impressions"] and base["impressions"] >= MIN_IMP and ci is not None and ci <= -DROP["impressions"]:
        out.append(("P2", "노출 감소", f"검색 노출 {cur['impressions']:,.0f}회(4주 평균 대비 {ci:+.0%})",
                    "키워드 순위·광고 노출(예산 소진·입찰)·재고 확인, SQP에서 빠진 검색어 확인"))
    if cur["impressions"] >= MIN_IMP and cc is not None and cc <= -DROP["ctr"]:
        why = ""
        if cur.get("price") and base["price"] and cur["price"] > base["price"] * 1.03:
            why = f" · 가격 ${base['price']:.2f} → ${cur['price']:.2f} 상승"
        out.append(("P2", "클릭률 하락", f"CTR {cur['ctr']:.2%}(4주 평균 대비 {cc:+.0%}){why}",
                    "메인 이미지·가격·쿠폰 표시·평점, 경쟁 상품 확인"))
    ca = cmp_["cart_rate_chg"]
    if cur["clicks"] >= MIN_CLICKS and ca is not None and ca <= -DROP["cart_rate"]:
        out.append(("P2", "장바구니율 하락", f"장바구니 {cur['cart_adds']:,.0f}/클릭 {cur['clicks']:,.0f}(4주 평균 대비 {ca:+.0%})",
                    "상세 페이지(불릿·이미지·A+·변형 선택)·배송 약속일 확인"))
    cv = cmp_["cvr_chg"]
    if cur["clicks"] >= MIN_CLICKS and cur["purchases"] == 0:
        out.append(("P1", "구매 없음", f"클릭 {cur['clicks']:,.0f}회인데 검색 구매 0", "가격·재고·Buy Box·리뷰·리스팅 오류 바로 확인"))
    elif cur["clicks"] >= MIN_CLICKS and cv is not None and cv <= -DROP["cvr"]:
        out.append(("P2", "구매 전환 하락", f"구매 전환 {cur['cvr']:.1%}(4주 평균 대비 {cv:+.0%})",
                    "최근 리뷰·평점, 가격, 재고·배송 기간, Buy Box 확인"))
    bb = st_now.get("buybox")
    if st_now.get("sessions", 0) >= 50 and bb is not None and bb < BUYBOX_MIN:
        out.append(("P1" if bb < 70 else "P2", "Buy Box 비율 낮음", f"Buy Box {bb:.0f}%(세션 {st_now['sessions']:,.0f})",
                    "가격·재고 확인, 다른 판매자(합승) 여부 확인"))
    if base["impressions"] and base["impressions"] >= MIN_IMP and ci is not None and ci >= 0.5 and (cc or 0) > -0.1:
        out.append(("P3", "노출 증가", f"검색 노출 {ci:+.0%}", "재고가 충분한지 확인(품절 주의)"))
    return out, cmp_


def change_effects(hist_by_asin: Dict[str, Dict[str, Dict]], wk: List[str]) -> List[Dict]:
    """‘리스팅·가격 수정’ 기록 → 반영 전 2주 vs 반영 후 2주."""
    try:
        import changes
        entries = [e for e in changes.history() if e.get("kind") == "적용" and e.get("asin")
                   and e.get("status") in ("반영됨", "제출값 저장됨")]
    except Exception:
        return []
    out = []
    for e in entries:
        d = datetime.fromisoformat(e.get("confirmed_at") or e["time"]).date()
        before = [w for w in wk if date.fromisoformat(w) < d][-2:]
        after = [w for w in wk if date.fromisoformat(w) - timedelta(days=6) > d][:2]
        h = hist_by_asin.get(e["asin"], {})

        def agg(ws):
            if not ws or not all(w in h for w in ws):
                return None
            s = {k: sum(h[w].get(k, 0) or 0 for w in ws) for k in ("impressions", "clicks", "cart_adds", "purchases")}
            return rates(s)
        b, a = agg(before), agg(after)
        row = {"적용일": d.isoformat(), "SKU": e["sku"], "ASIN": e["asin"], "항목": e["field"],
               "새 값": str(e.get("new", ""))[:80], "상태": e.get("status", "")}
        if not b:
            row["결과"] = "반영 전 데이터 부족"
        elif not a:
            ready = d + timedelta(days=(5 - d.weekday()) % 7 + 14 + 1)
            row["결과"] = f"측정 대기(반영 후 2주 데이터가 쌓이면, {ready} 이후)"
        else:
            for k, name in (("impressions", "노출"), ("ctr", "CTR"), ("cvr", "구매 전환")):
                c = _chg(a[k], b[k])
                row[f"{name} 전"] = (f"{b[k]:,.0f}" if k == "impressions" else (f"{b[k]:.2%}" if b[k] is not None else ""))
                row[f"{name} 후"] = (f"{a[k]:,.0f}" if k == "impressions" else (f"{a[k]:.2%}" if a[k] is not None else ""))
                row[f"{name} 변화"] = "" if c is None else f"{c:+.0%}"
            ci, cc, cv = (_chg(a[k], b[k]) for k in ("impressions", "ctr", "cvr"))
            good = [n for n, c in (("노출", ci), ("CTR", cc), ("구매 전환", cv)) if c is not None and c >= 0.1]
            bad = [n for n, c in (("노출", ci), ("CTR", cc), ("구매 전환", cv)) if c is not None and c <= -0.1]
            row["결과"] = ("좋아짐: " + ", ".join(good) if good and not bad else "나빠짐: " + ", ".join(bad) if bad and not good
                         else ("섞임: +" + ",".join(good) + " / -" + ",".join(bad)) if good and bad else "큰 변화 없음(±10% 이내)")
        out.append(row)
    return out


# ================================================================ 실행

def run(open_report: bool = False) -> Optional[Path]:
    import sp_api
    import ba_collect as ba
    import daily_bus as bus
    cfg = sp_api.load_config()
    use_api = bool(cfg.get("use_api") and sp_api.ready(cfg))
    wk = weeks()
    print(f"【1】 주간 데이터 확인: {ba.week_text(wk[0]).split(' ~ ')[0]} ~ {wk[-1]} ({len(wk)}주)")
    scp, st, src = {}, {}, {}
    for i, w in enumerate(wk, 1):
        scp[w], s1 = load_scp(w, use_api)
        st[w], s2 = load_st(w, use_api) if w in wk[-5:] else ({}, "생략(최근 5주만)")
        src[w] = (s1, s2)
        print(f"   [{i}/{len(wk)}] {ba.week_text(w)}: SCP {len(scp[w])}개 ASIN({s1}) · 비즈니스 {len(st[w])}개({s2})")
    have = [w for w in wk if scp[w]]
    if not have:
        print("❌ Brand Analytics SCP 데이터가 한 주도 없습니다. SP-API Brand Analytics 권한을 받거나, ‘광고 데이터 수집’을 먼저 실행하세요.")
        return None
    cur_w = have[-1]
    if cur_w != wk[-1]:
        print(f"   ℹ 가장 최근 주({ba.week_text(wk[-1])}) 데이터가 아직 없어 {ba.week_text(cur_w)} 주로 진단합니다.")
    asins = sorted({a for w in have for a in scp[w]})
    skus = sku_map()
    hist_by_asin = {a: {w: rates(scp[w][a]) for w in have if a in scp[w]} for a in asins}

    print("\n【2】 진단")
    rows, alerts = [], []
    for a in asins:
        hist = [hist_by_asin[a].get(w, rates({})) for w in wk[:wk.index(cur_w) + 1]]
        cur = hist[-1]
        if not cur["impressions"] and not cur["clicks"]:
            continue
        sn = st.get(cur_w, {}).get(a, {})
        found, cmp_ = diagnose(a, hist, sn)
        title = scp[cur_w].get(a, {}).get("title") or ""
        sku = skus.get(a) or sn.get("sku", "")
        for lv, kind, det, adv in found:
            alerts.append((lv, kind, f"{sku or a}: {det}", adv))
        pct = lambda v: "" if v is None else f"{v:.2%}"
        chg = lambda v: "" if v is None else f"{v:+.0%}"
        rows.append({"진단": " / ".join(k for _, k, _, _ in found) or "정상", "ASIN": a, "SKU": sku, "상품명": str(title)[:60],
                     "노출": int(cur["impressions"]), "노출 변화(4주)": chg(cmp_.get("impressions_chg")),
                     "클릭": int(cur["clicks"]), "CTR": pct(cur["ctr"]), "CTR 변화": chg(cmp_.get("ctr_chg")),
                     "장바구니": int(cur["cart_adds"]), "장바구니율": pct(cur["cart_rate"]), "장바구니율 변화": chg(cmp_.get("cart_rate_chg")),
                     "구매": int(cur["purchases"]), "구매 전환": pct(cur["cvr"]), "구매 전환 변화": chg(cmp_.get("cvr_chg")),
                     "가격(중앙값)": cur.get("price") or "", "세션": int(sn.get("sessions", 0) or 0),
                     "세션 전환율": f"{sn['usp']:.1f}%" if sn.get("usp") is not None and sn else "",
                     "Buy Box": f"{sn['buybox']:.0f}%" if sn.get("buybox") is not None and sn else "",
                     "_lv": min([x[0] for x in found] or ["P9"]), "_imp": cur["impressions"]})
    rows.sort(key=lambda r: (r["_lv"], -r["_imp"]))
    alerts.sort(key=lambda x: x[0])
    tot = {w: rates({k: sum(scp[w][a].get(k, 0) or 0 for a in scp[w]) for k in ("impressions", "clicks", "cart_adds", "purchases")})
           for w in have}
    t_now = tot[cur_w]
    t_prev = tot[have[-2]] if len(have) >= 2 else None
    effects = change_effects(hist_by_asin, have)
    kpis = {"주": ba.week_text(cur_w), "ASIN": len(rows), "검색 노출": f"{t_now['impressions']:,.0f}",
            "클릭": f"{t_now['clicks']:,.0f}", "구매": f"{t_now['purchases']:,.0f}",
            "CTR": f"{t_now['ctr']:.2%}" if t_now["ctr"] else "", "구매 전환": f"{t_now['cvr']:.1%}" if t_now["cvr"] else "",
            "P1": sum(1 for a in alerts if a[0] == "P1"), "P2": sum(1 for a in alerts if a[0] == "P2")}
    if t_prev:
        for k, name in (("impressions", "노출"), ("purchases", "구매")):
            c = _chg(t_now[k], t_prev[k])
            kpis[f"{name} 전주 대비"] = "" if c is None else f"{c:+.0%}"
    print(f"   {kpis['주']}: ASIN {len(rows)}개 · 노출 {kpis['검색 노출']} · 구매 {kpis['구매']} · P1 {kpis['P1']} · P2 {kpis['P2']} · 수정 효과 {len(effects)}건")

    print("\n【3】 보고서 만들기")
    out = write_report(cur_w, wk, have, rows, alerts, kpis, tot, hist_by_asin, effects, src)
    bus.record("주간 퍼널", {k: v for k, v in kpis.items()}, alerts, report=out)
    print(f"✅ 주간 퍼널 진단 → {out}")
    if open_report:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass
    return out


def write_report(cur_w, wk, have, rows, alerts, kpis, tot, hist_by_asin, effects, src) -> Path:
    import inventory as inv
    import ba_collect as ba
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font, PatternFill
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / f"funnel_{cur_w}.xlsx"
    F = inv.FONT
    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 60
    ws.cell(1, 1, f"주간 ASIN 퍼널 진단 · {ba.week_text(cur_w)}").font = Font(name=F, size=15, bold=True, color="1F3864")
    ws.cell(2, 1, "검색 노출 → 클릭 → 장바구니 → 구매(Brand Analytics SCP) + 세션·Buy Box(비즈니스 보고서). 이번 주를 지난 4주 평균과 비교").font = Font(name=F, color="6B6B6B")
    r = 4
    for k, v in kpis.items():
        ws.cell(r, 1, k).font = Font(name=F, bold=True)
        ws.cell(r, 2, v).font = Font(name=F)
        r += 1
    r += 1
    ws.cell(r, 1, "조치 필요").font = Font(name=F, size=12, bold=True, color="1F3864")
    r += 1
    for i, h in enumerate(("등급", "분류", "내용", "조치"), 1):
        c = ws.cell(r, i, h)
        c.font, c.fill = Font(name=F, bold=True), PatternFill("solid", fgColor="DCE6F1")
    fills = {"P1": PatternFill("solid", fgColor="FFE8CC"), "P2": PatternFill("solid", fgColor="FFF6D6")}
    for a in alerts:
        r += 1
        for i, v in enumerate(a, 1):
            c = ws.cell(r, i, v)
            c.font = Font(name=F)
            if a[0] in fills:
                c.fill = fills[a[0]]
    if not alerts:
        ws.cell(r + 1, 1, "이번 주 크게 떨어진 단계가 없습니다.").font = Font(name=F)

    s2 = wb.create_sheet("퍼널(이번 주)")
    head = [k for k in rows[0].keys() if not k.startswith("_")] if rows else ["진단"]
    inv.write_table_sheet(s2, head, rows, {"상품명": 40, "진단": 26})
    for i, row in enumerate(rows, start=2):
        if row["_lv"] in fills:
            s2.cell(i, 1).fill = fills[row["_lv"]]

    # 주별 합계 + 그래프
    s3 = wb.create_sheet("주별 추세")
    s3.append(["주(끝나는 날)", "노출", "클릭", "장바구니", "구매", "CTR(%)", "장바구니율(%)", "구매 전환(%)"])
    for w in have:
        t = tot[w]
        s3.append([w, t["impressions"], t["clicks"], t["cart_adds"], t["purchases"],
                   round((t["ctr"] or 0) * 100, 2), round((t["cart_rate"] or 0) * 100, 2), round((t["cvr"] or 0) * 100, 2)])
    for col in "ABCDEFGH":
        s3.column_dimensions[col].width = 15
    n = len(have) + 1
    ch = LineChart()
    ch.title, ch.height, ch.width = "검색 노출·클릭·구매(주별 합계)", 8, 18
    ch.add_data(Reference(s3, min_col=2, max_col=3, min_row=1, max_row=n), titles_from_data=True)
    ch.set_categories(Reference(s3, min_col=1, min_row=2, max_row=n))
    s3.add_chart(ch, "J2")
    ch2 = LineChart()
    ch2.title, ch2.height, ch2.width = "CTR · 장바구니율 · 구매 전환(%)", 8, 18
    ch2.add_data(Reference(s3, min_col=6, max_col=8, min_row=1, max_row=n), titles_from_data=True)
    ch2.set_categories(Reference(s3, min_col=1, min_row=2, max_row=n))
    s3.add_chart(ch2, "J20")

    # ASIN별 8주(노출·CTR·구매 전환)
    s4 = wb.create_sheet("ASIN별 8주")
    top = [x["ASIN"] for x in rows][:60]
    s4.append(["ASIN", "SKU", "지표"] + have)
    sku_of = {x["ASIN"]: x["SKU"] for x in rows}
    for a in top:
        h = hist_by_asin.get(a, {})
        s4.append([a, sku_of.get(a, ""), "노출"] + [int(h[w]["impressions"]) if w in h else "" for w in have])
        s4.append([a, sku_of.get(a, ""), "CTR(%)"] + [round(h[w]["ctr"] * 100, 2) if w in h and h[w]["ctr"] is not None else "" for w in have])
        s4.append([a, sku_of.get(a, ""), "구매 전환(%)"] + [round(h[w]["cvr"] * 100, 1) if w in h and h[w]["cvr"] is not None else "" for w in have])
    for c in s4[1]:
        c.font, c.fill = Font(name=F, bold=True), PatternFill("solid", fgColor="DCE6F1")
    s4.column_dimensions["A"].width, s4.column_dimensions["B"].width, s4.column_dimensions["C"].width = 14, 18, 12
    s4.freeze_panes = "D2"

    s5 = wb.create_sheet("수정 효과")
    if effects:
        keys = []
        for e in effects:
            for k in e:
                if k not in keys:
                    keys.append(k)
        inv.write_table_sheet(s5, keys, effects, {"새 값": 40, "결과": 40})
    else:
        s5.append(["‘리스팅·가격 수정’으로 반영된 변경이 아직 없습니다. 변경하면 반영 전 2주와 반영 후 2주를 비교해 여기에 보여 줍니다."])

    s6 = wb.create_sheet("데이터 출처")
    s6.append(["주", "SCP", "비즈니스 보고서"])
    for w in wk:
        s6.append([ba.week_text(w), src[w][0], src[w][1]])
    s6.append([])
    s6.append(["판단 기준", f"노출 -{DROP['impressions']:.0%}, CTR·장바구니율·구매 전환 -{DROP['ctr']:.0%} 이하(지난 4주 평균 대비). "
                          f"노출 {MIN_IMP}회·클릭 {MIN_CLICKS}회 미만은 판단하지 않음. Buy Box {BUYBOX_MIN:.0f}% 미만 표시"])
    for col, w_ in zip("ABC", (26, 40, 40)):
        s6.column_dimensions[col].width = w_
    wb.save(out)
    return out


def status_text() -> str:
    files = sorted(OUT_DIR.glob("funnel_*.xlsx")) if OUT_DIR.exists() else []
    return f"마지막 진단: {files[-1].stem.replace('funnel_', '')} 주" if files else "아직 실행하지 않음"

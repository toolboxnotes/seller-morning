"""
매일 자동 실행 · 오늘의 요약

  1. 정해진 기능을 순서대로 실행합니다(홈 화면 → 계정 상태 → 전체 상품 → FBA → 주문 → 비즈니스 → 광고 → Listing).
     한 기능이 실패해도 다음 기능은 계속합니다(로그인 만료일 때만 멈춤).
  2. 각 기능이 남긴 요점(주요 수치 + 조치 필요 항목)을 모아 ‘오늘의 요약’ Excel을 만듭니다:
     reports/daily/daily_summary_날짜.xlsx
  3. Windows 작업 스케줄러에 등록하면 매일 정한 시간에 자동으로 실행합니다(컴퓨터가 켜져 있고 로그인된 상태여야 함).

모든 기능은 1단계(보기·제안만)이며, 상품·광고 설정은 바꾸지 않습니다.
"""

import json
import re
import os
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import daily_bus as bus

BASE_DIR = Path(__file__).resolve().parent.parent
DAILY_DIR = BASE_DIR / "reports" / "daily"
SETTINGS = DAILY_DIR / "daily_settings.json"
BAT = BASE_DIR / "daily_run.bat"
TASK_NAME = "AmazonAI_Daily"
LOG = DAILY_DIR / "daily_run.log"


# ================================================================ 실행 항목

def _home(open_report=False):
    import sc_pages
    r = sc_pages.read_many(["home"])["home"] or {}
    bus.record("홈 화면", r.get("kpis", {}), r.get("alerts", []))
    return None


def _account_health(open_report=False):
    import account_health
    return account_health.run(open_report=open_report)


def _listings(open_report=False):
    import inventory
    return inventory.run_listings(open_report=open_report)


def _fba(open_report=False):
    import inventory
    return inventory.run_fba(open_report=open_report)


def _orders(open_report=False):
    import orders
    return orders.run(period="2", open_report=open_report)


def _business(open_report=False):
    import business
    return business.run(period="2", open_report=open_report)


def _ads(open_report=False):
    import ads_agent
    return ads_agent.run(open_report=open_report)


def _listing_check(open_report=False):
    import listing
    return listing.run(open_report=open_report)


def _action_items(open_report=False):
    return run_action_items(open_report)


def run_action_items(open_report: bool = False) -> Optional[Path]:
    """SP-API Action Items: Amazon이 알려 주는 문제·할 일·추천 → Excel + 오늘의 요약."""
    import sp_api
    import sp_data
    items = sp_api.try_api("Amazon 할 일(Action Items)", sp_data.action_items)
    if items is None:
        print("   → Amazon 할 일은 API로만 가져올 수 있어 이번에는 건너뜁니다(실패가 아니라 설정 필요).")
        return None
    al = sp_data.action_alerts(items)
    from collections import Counter
    cnt = Counter(a[0] for a in al)
    bus.record("Amazon 할 일", {"전체": len(al), "P0": cnt.get("P0", 0), "P1": cnt.get("P1", 0),
                               "P2": cnt.get("P2", 0), "P3": cnt.get("P3", 0)}, al)
    import sc_pages
    from openpyxl import Workbook
    out = DAILY_DIR / f"action_items_{date.today().isoformat()}.xlsx"
    DAILY_DIR.mkdir(parents=True, exist_ok=True)
    Workbook().save(out)
    sc_pages.append_sheets(out, [{"name": "Amazon 할 일", "src": "SP-API", "time": datetime.now().isoformat(timespec="seconds"),
                                  "kpis": {"전체": len(al), "출처": "SP-API Action Items(대기·기한 지남)"},
                                  "alerts": sorted(al), "rows": []}])
    try:
        from openpyxl import load_workbook
        wb = load_workbook(out)
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]
            wb.save(out)
    except Exception:
        pass
    print(f"   ✅ Amazon 할 일 {len(al)}건(P0 {cnt.get('P0', 0)}, P1 {cnt.get('P1', 0)}) → {out.name}")
    if open_report:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass
    return out


STEPS: List[Tuple[str, str, Callable]] = [
    ("home", "홈 화면(매출·IPI·주문·순수익)", _home),
    ("action_items", "Amazon 할 일(SP-API)", _action_items),
    ("account_health", "계정 상태(+클레임·고객의 소리)", _account_health),
    ("listings", "전체 상품 재고", _listings),
    ("fba", "FBA 재고(+대시보드·입고 배송)", _fba),
    ("orders", "주문 보고서 최근 7일(+반품 분석)", _orders),
    ("business", "비즈니스 보고서 최근 7일", _business),
    ("ads", "광고 데이터 수집·조치 제안(+브랜드 노출·추천)", _ads),
    ("listing_check", "Listing 점검", _listing_check),
    ("buybox", "Buy Box·합승 판매자 확인(SP-API)", lambda open_report=False: __import__("buybox").run(open_report=open_report)),
    ("funnel", "주간 ASIN 퍼널 진단", lambda open_report=False: __import__("funnel").run(open_report=open_report)),
    ("finance", "재무·보상 점검(주 1회)", lambda open_report=False: __import__("finance").run_weekly(open_report=open_report)),
]
NEW_STEPS = ("action_items", "funnel", "buybox", "finance")          # 새로 생긴 항목은 한 번 자동으로 켬
DEFAULT = {"time": "08:10", "steps": [k for k, _, _ in STEPS]}


def load_settings() -> Dict:
    try:
        s = json.loads(SETTINGS.read_text(encoding="utf-8"))
        s = {**DEFAULT, **s}
        for k_new in NEW_STEPS:
            if not s.get(f"{k_new}_added"):
                if k_new not in s["steps"]:
                    s["steps"] = [k for k, _, _ in STEPS if k in s["steps"] or k == k_new]
                s[f"{k_new}_added"] = True
        return s
    except Exception:
        return dict(DEFAULT)


def save_settings(s: Dict) -> None:
    DAILY_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS.write_text(json.dumps(s, ensure_ascii=False, indent=1), encoding="utf-8")


# ================================================================ 전체 실행

def run_all(open_summary: bool = True) -> Optional[Path]:
    from sc_browser import BrowserNotFound, SessionExpired
    try:
        import housekeeping
        housekeeping.daily_once()
    except Exception:
        pass
    s = load_settings()
    todo = [(k, n, f) for k, n, f in STEPS if k in s["steps"]]
    t_all = time.time()
    import version
    print(f"\n=== 매일 자동 실행 시작 {datetime.now():%Y-%m-%d %H:%M} · {len(todo)}개 기능 · 프로그램 {version.text()} ===")
    print("   (1단계: 보기·제안만 합니다. 상품·광고 설정은 바꾸지 않습니다)")
    stop = ""
    pre_done = False
    for i, (key, name, fn) in enumerate(todo, 1):
        if stop:
            bus.step(key, name, "건너뜀", 0, error=stop)
            continue
        # 광고 보고서는 생성에 20~40분 걸림 → 첫 번째 브라우저 기능이 끝나면 미리 요청해 두고, 광고 단계에서는 받기만
        if not pre_done and key not in ("home", "action_items") and "ads" in [k for k, _, _ in todo]:
            pre_done = True
            t1 = time.time()
            try:
                import ads_agent
                print("\n────────── 광고 보고서 미리 요청(다른 기능이 도는 동안 생성) ──────────")
                msg = ads_agent.prerequest()
                print(f"   {msg}")
                bus.step("ads_pre", "광고 보고서 미리 요청", "완료", (time.time() - t1) / 60, error="" if msg else "")
            except (SessionExpired, BrowserNotFound) as e:
                stop = f"로그인 만료 또는 브라우저 문제로 이후 기능을 건너뜀: {e}"
                bus.step("ads_pre", "광고 보고서 미리 요청", "실패", (time.time() - t1) / 60, error=str(e)[:200])
                print(f"❌ {stop}")
                bus.step(key, name, "건너뜀", 0, error=stop)
                continue
            except Exception as e:
                bus.step("ads_pre", "광고 보고서 미리 요청", "실패", (time.time() - t1) / 60, error=f"{type(e).__name__}: {str(e)[:150]}")
                print(f"   ⚠ 미리 요청 실패({type(e).__name__}) → 광고 단계에서 평소대로 요청합니다.")
        print(f"\n────────── [{i}/{len(todo)}] {name} ──────────")
        t0 = time.time()
        try:
            out = fn(open_report=False)
            bus.step(key, name, "완료", (time.time() - t0) / 60, str(out) if out else "")
        except (SessionExpired, BrowserNotFound) as e:
            stop = f"로그인 만료 또는 브라우저 문제로 이후 기능을 건너뜀: {e}"
            bus.step(key, name, "실패", (time.time() - t0) / 60, error=str(e)[:200])
            print(f"❌ {stop}")
        except Exception as e:
            msg = f"{type(e).__name__}: {str(e).splitlines()[0][:200] if str(e) else ''}"
            bus.step(key, name, "실패", (time.time() - t0) / 60, error=msg)
            print(f"❌ {name} 실패: {msg} → 다음 기능으로 넘어갑니다.")
    try:                                   # 리스팅·가격 수정 반영 확인(API, 몇 초)
        import changes
        if changes.pending():
            print(f"\n────────── 리스팅·가격 수정 반영 확인({len(changes.pending())}건) ──────────")
            changes.verify()
    except Exception as e:
        print(f"   ⚠ 수정 반영 확인 실패: {type(e).__name__}: {str(e)[:120]}")
    out = build_summary()
    try:                                   # 기능별 걸린 시간(오래 걸린 순)
        st_ = bus.load().get("steps", {})
        tm = sorted(((v.get("name", k), v.get("minutes", 0), v.get("status", "")) for k, v in st_.items()), key=lambda x: -(x[1] or 0))
        print("\n   걸린 시간: " + " · ".join(f"{n} {m:.0f}분{'' if s_ == '완료' else '(' + s_ + ')'}" for n, m, s_ in tm[:10]))
    except Exception:
        pass
    print(f"\n=== 매일 자동 실행 끝 · {(time.time() - t_all) / 60:.0f}분 ===")
    if open_summary and out:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass
    return out


# ================================================================ 오늘의 요약

# ================================================================ 30일 추세

TREND_SKIP = ("기간", "출처", "시각", "상태", "등급(AHR)", "업데이트", "주요", "분포", "피드백", "건강도")
TREND_CHART = ["오늘 할 일 P0+P1", "광고 · 전날 광고비", "광고 · 전날 광고 매출", "광고 · 7일 ACoS",
               "FBA 재고 · FBA 대시보드: IPI", "주문 보고서 · 반품 분석: 반품률"]


def _num(v):
    import re
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"[-−]?\$?\s*([\d,]+(?:\.\d+)?)", str(v or ""))
    if not m:
        return None
    try:
        x = float(m.group(1).replace(",", ""))
    except ValueError:
        return None
    return -x if str(v).strip().startswith(("-", "−")) else x


def trend_data(d: Optional[date] = None, days: int = 30):
    """최근 days일의 status.json에서 숫자 지표 모으기 → (날짜 목록, {지표: {날짜: 값}})."""
    from datetime import timedelta as _td
    d = d or date.today()
    series, dates = {}, []
    for i in range(days - 1, -1, -1):
        day = d - _td(days=i)
        f = DAILY_DIR / day.isoformat() / "status.json"
        if not f.exists():
            continue
        try:
            st = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        mods = st.get("modules", {})
        if not mods:
            continue
        key = day.isoformat()
        dates.append(key)
        al = [a for m in mods.values() for a in m.get("alerts", [])]
        series.setdefault("오늘 할 일 P0+P1", {})[key] = sum(1 for a in al if a and a[0] in ("P0", "P1"))
        series.setdefault("오늘 할 일 P0", {})[key] = sum(1 for a in al if a and a[0] == "P0")
        for mname, m in mods.items():
            for k, v in (m.get("kpis") or {}).items():
                if any(x in k for x in TREND_SKIP):
                    continue
                n = _num(v)
                if n is not None:
                    series.setdefault(f"{mname} · {k}", {})[key] = n
    series = {k: v for k, v in series.items() if len(v) >= 2 or k.startswith("오늘 할 일")}
    return dates, series


def _trend_sheet(wb, d, FONT, HEAD):
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font
    dates, series = trend_data(d)
    ws = wb.create_sheet("30일 추세")
    if len(dates) < 2:
        ws.cell(1, 1, "추세는 이틀 이상 실행한 기록이 쌓이면 표시됩니다(매일 실행 결과를 모아 최근 30일을 보여 줌).").font = Font(name=FONT)
        return
    heads = ["지표"] + [x[5:] for x in dates] + ["변화(최근-이전)"]
    for c, h in enumerate(heads, start=1):
        x = ws.cell(1, c, h)
        x.font, x.fill = Font(name=FONT, bold=True), HEAD
    ws.column_dimensions["A"].width = 48
    order = [k for k in TREND_CHART if k in series] + sorted(k for k in series if k not in TREND_CHART)
    for r, k in enumerate(order, start=2):
        ws.cell(r, 1, k).font = Font(name=FONT, size=10)
        vals = [series[k].get(dt) for dt in dates]
        for c, v in enumerate(vals, start=2):
            ws.cell(r, c, v).font = Font(name=FONT, size=10)
        have = [v for v in vals if v is not None]
        if len(have) >= 2:
            ch = have[-1] - have[-2]
            x = ws.cell(r, len(dates) + 2, round(ch, 2))
            x.font = Font(name=FONT, size=10, color=("C01C1C" if ("P0" in k or "ACoS" in k or "반품" in k) and ch > 0 else None))
    ws.freeze_panes = "B2"
    # 주요 지표 그래프(있는 것만)
    pos = 0
    for k in TREND_CHART:
        if k not in series:
            continue
        r = order.index(k) + 2
        chart = LineChart()
        chart.title = k
        chart.height, chart.width = 6.5, 14
        chart.legend = None
        chart.add_data(Reference(ws, min_col=2, max_col=len(dates) + 1, min_row=r, max_row=r), from_rows=True, titles_from_data=False)
        chart.set_categories(Reference(ws, min_col=2, max_col=len(dates) + 1, min_row=1, max_row=1))
        ws.add_chart(chart, f"{'A' if pos % 2 == 0 else 'J'}{len(order) + 4 + (pos // 2) * 15}")
        pos += 1


def _group(items):
    """같은 기능·등급·분류가 3건 넘으면 한 줄로 묶음(전체는 ‘조치 필요 전체’ 시트)."""
    from collections import OrderedDict
    g = OrderedDict()
    for a, name in items:
        g.setdefault((a[0], name, a[1]), []).append(a)
    out = []
    for (lv, name, typ), arr in g.items():
        if len(arr) <= 3:
            out += [(a, name) for a in arr]
        else:
            ex = "; ".join(x[2][:40] for x in arr[:2])
            out.append(([lv, typ, f"{len(arr)}건 (예: {ex} …)", arr[0][3]], name))
    return out


def summary_alerts(data: Dict):
    """오늘의 요약에 나오는 조치 목록: ([(alert[등급,분류,내용,조치], 기능 이름)], 등급별 건수, 데이터 출처)."""
    mods = data.get("modules", {})
    alerts = [(a, name) for name, m in mods.items() for a in m.get("alerts", [])]
    srcs = data.get("sources", [])
    last_src = {}
    for s_ in srcs:
        last_src[s_["label"]] = s_
    for s_ in last_src.values():
        if s_["method"].startswith("API 실패"):
            alerts.append((["P2", "API 실패 → 브라우저로 대신함", f"{s_['label']}: {s_['detail']}",
                            "⚙ 설정(API·메일) → ‘전체 점검’으로 원인(권한·설정)을 확인하세요"], "데이터 출처"))
    try:
        import changes
        for a in changes.alerts():
            alerts.append((a, "리스팅·가격 수정"))
    except Exception:
        pass
    try:
        import sp_api
        for a in sp_api.token_reminders():
            alerts.append((a, "API 설정"))
    except Exception:
        pass
    alerts.sort(key=lambda x: (bus.LEVELS.index(x[0][0]) if x[0][0] in bus.LEVELS else 9))
    cnt = {lv: sum(1 for a, _ in alerts if a[0] == lv) for lv in bus.LEVELS}
    return alerts, cnt, last_src


def build_summary(d: Optional[date] = None) -> Optional[Path]:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    import inventory as inv
    d = d or date.today()
    data = bus.load(d)
    mods, steps = data.get("modules", {}), data.get("steps", {})
    FONT = inv.FONT
    HEAD = PatternFill("solid", fgColor="DCE6F1")
    fills = {"P0": PatternFill("solid", fgColor="F8D7DA"), "P1": PatternFill("solid", fgColor="FFE8CC")}
    alerts, cnt, last_src = summary_alerts(data)
    srcs = data.get("sources", [])

    wb = Workbook()
    ws = wb.active
    ws.title = "오늘의 요약"
    for col, w in zip("ABCDE", (30, 22, 12, 60, 50)):
        ws.column_dimensions[col].width = w
    r = 1
    ws.cell(r, 1, f"오늘의 요약 · {d.isoformat()}").font = Font(name=FONT, size=16, bold=True, color="1F3864")
    r += 1
    ws.cell(r, 1, f"만든 시각 {datetime.now():%H:%M} · 프로그램 v{__import__('version').VERSION} · P0 {cnt['P0']}건 · P1 {cnt['P1']}건 · P2 {cnt['P2']}건 · P3 {cnt['P3']}건").font = Font(name=FONT, color="6B6B6B")
    r += 1
    if last_src:
        from collections import Counter as _C
        mc = _C("API" if v["method"] == "API" else ("API 실패" if v["method"].startswith("API 실패") else "브라우저")
                for v in last_src.values())
        ws.cell(r, 1, f"데이터 출처: API {mc.get('API', 0)}건 · 브라우저 {mc.get('브라우저', 0)}건 · API 실패 후 브라우저 "
                      f"{mc.get('API 실패', 0)}건(‘데이터 출처’ 시트)").font = Font(name=FONT, color="6B6B6B")
        r += 1
    r += 1
    home = mods.get("홈 화면", {}).get("kpis", {})
    if home:
        ws.cell(r, 1, "홈 화면").font = Font(name=FONT, size=12, bold=True)
        r += 1
        for k, v in home.items():
            ws.cell(r, 1, k).font = Font(name=FONT)
            ws.cell(r, 2, v).font = Font(name=FONT, bold=True)
            r += 1
        r += 1
    ws.cell(r, 1, "오늘 처리할 일(P0·P1)").font = Font(name=FONT, size=12, bold=True, color="C01C1C")
    r += 1
    top = _group([x for x in alerts if x[0][0] in ("P0", "P1")])
    if not top:
        ws.cell(r, 1, "P0·P1 항목이 없습니다 ✓").font = Font(name=FONT)
        r += 1
    for (lv, typ, det, adv), name in top[:40]:
        for c, v in enumerate((f"[{lv}] {name}", typ, "", det, adv), start=1):
            x = ws.cell(r, c, v)
            x.font = Font(name=FONT, size=10, bold=(c == 1))
            x.alignment = Alignment(wrap_text=True, vertical="top")
        if lv in fills:
            ws.cell(r, 1).fill = fills[lv]
        r += 1
    if len(top) > 40:
        ws.cell(r, 1, f"… 외 {len(top) - 40}건(‘조치 필요 전체’ 시트)").font = Font(name=FONT, color="6B6B6B")
        r += 1
    r += 1
    ws.cell(r, 1, "실행 결과").font = Font(name=FONT, size=12, bold=True)
    r += 1
    for c, h in enumerate(("기능", "결과", "소요(분)", "보고서 / 오류", "실행 시각"), start=1):
        x = ws.cell(r, c, h)
        x.font, x.fill = Font(name=FONT, bold=True), HEAD
    r += 1
    for key, name, _ in STEPS:
        st = steps.get(key)
        if not st:
            continue
        vals = (name, st["status"], st["minutes"], st.get("report") or st.get("error", ""), st.get("time", "")[11:16])
        for c, v in enumerate(vals, start=1):
            x = ws.cell(r, c, v)
            x.font = Font(name=FONT, size=10, color=("C01C1C" if st["status"] == "실패" and c == 2 else None))
            x.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    # 전체 항목
    w2 = wb.create_sheet("조치 필요 전체")
    heads = ("등급", "기능", "분류", "내용", "조치")
    for c, (h, wd) in enumerate(zip(heads, (6, 22, 22, 70, 60)), start=1):
        x = w2.cell(1, c, h)
        x.font, x.fill = Font(name=FONT, bold=True), HEAD
        w2.column_dimensions[x.column_letter].width = wd
    for i, ((lv, typ, det, adv), name) in enumerate(alerts, start=2):
        for c, v in enumerate((lv, name, typ, det, adv), start=1):
            x = w2.cell(i, c, v)
            x.font = Font(name=FONT, size=10)
            x.alignment = Alignment(wrap_text=True, vertical="top")
        if lv in fills:
            w2.cell(i, 1).fill = fills[lv]
    w2.freeze_panes = "A2"
    if alerts:
        w2.auto_filter.ref = f"A1:E{len(alerts) + 1}"

    # 각 기능 수치
    w3 = wb.create_sheet("기능별 수치")
    w3.column_dimensions["A"].width = 26
    w3.column_dimensions["B"].width = 36
    w3.column_dimensions["C"].width = 30
    r = 1
    for name, m in mods.items():
        w3.cell(r, 1, name).font = Font(name=FONT, size=12, bold=True)
        w3.cell(r, 3, ("보고서: " + Path(m["report"]).name) if m.get("report") else "").font = Font(name=FONT, color="6B6B6B")
        r += 1
        for k, v in (m.get("kpis") or {}).items():
            w3.cell(r, 1, k).font = Font(name=FONT)
            w3.cell(r, 2, v).font = Font(name=FONT, bold=True)
            r += 1
        r += 1
    # 데이터 출처
    w4 = wb.create_sheet("데이터 출처")
    for c, (h, wd) in enumerate(zip(("시각", "자료", "가져온 방법", "내용"), (10, 40, 22, 90)), start=1):
        x = w4.cell(1, c, h)
        x.font, x.fill = Font(name=FONT, bold=True), HEAD
        w4.column_dimensions[x.column_letter].width = wd
    for i, s_ in enumerate(srcs, start=2):
        for c, v in enumerate((s_["time"][11:16], s_["label"], s_["method"], s_["detail"]), start=1):
            x = w4.cell(i, c, v)
            x.font = Font(name=FONT, size=10, color=("C01C1C" if s_["method"].startswith("API 실패") and c == 3 else None))
    # 30일 추세
    try:
        _trend_sheet(wb, d, FONT, HEAD)
    except Exception as e:
        print(f"   ⚠ 추세 시트를 만들지 못했습니다: {type(e).__name__}: {e}")
    DAILY_DIR.mkdir(parents=True, exist_ok=True)
    out = DAILY_DIR / f"daily_summary_{d.isoformat()}.xlsx"
    wb.save(out)
    print(f"\n📋 오늘의 요약: P0 {cnt['P0']}건, P1 {cnt['P1']}건, P2 {cnt['P2']}건")
    for (lv, typ, det, adv), name in top[:8]:
        print(f"  [{lv}] {name} · {typ}: {det[:80]}")
    print(f"   저장: {out}")
    return out


# ================================================================ 작업 스케줄러(Windows)

def _write_bat() -> Path:
    BAT.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "cd /d \"%~dp0\"\r\n"
        "set PYTHONUTF8=1\r\n"
        "if not exist \"reports\\daily\" mkdir \"reports\\daily\"\r\n"
        "if exist \".venv\\Scripts\\python.exe\" (\r\n"
        "    \".venv\\Scripts\\python.exe\" app\\main.py --daily >> \"reports\\daily\\daily_run.log\" 2>&1\r\n"
        ") else (\r\n"
        "    python app\\main.py --daily >> \"reports\\daily\\daily_run.log\" 2>&1\r\n"
        ")\r\n", encoding="utf-8")
    return BAT


# ---------------------------------------------------------------- macOS: launchd(매일 같은 시각, 잠자기였으면 깨어난 뒤 실행)
MAC_SH = BASE_DIR / "daily_run.command"
MAC_LABEL = "com.amazonaiproject.daily"
MAC_PLIST = Path.home() / "Library" / "LaunchAgents" / f"{MAC_LABEL}.plist"
def _remove_old_mac_plist() -> None:
    """예전 이름(com.<상점>.amazonai.daily)으로 만든 예약이 있으면 지움."""
    for old in (Path.home() / "Library" / "LaunchAgents").glob("com.*.amazonai.daily.plist"):
        subprocess.run(["launchctl", "unload", "-w", str(old)], capture_output=True)
        old.unlink(missing_ok=True)


def _write_mac_sh() -> Path:
    MAC_SH.write_text(
        "#!/bin/bash\n"
        'cd "$(dirname "$0")"\n'
        "export PYTHONUTF8=1 PYTHONIOENCODING=utf-8\n"
        "mkdir -p reports/daily\n"
        'VENV="$HOME/Library/Application Support/AmazonAIProject/venv"\n'
        'if [ -x "$VENV/bin/python" ]; then PY="$VENV/bin/python"; else PY="python3"; fi\n'
        '"$PY" app/main.py --daily >> "reports/daily/daily_run.log" 2>&1\n', encoding="utf-8")
    os.chmod(MAC_SH, 0o755)
    return MAC_SH


def _mac_schedule(hhmm: str) -> bool:
    import plistlib
    h, m = (int(x) for x in hhmm.split(":"))
    sh = _write_mac_sh()
    _remove_old_mac_plist()
    DAILY_DIR.mkdir(parents=True, exist_ok=True)
    MAC_PLIST.parent.mkdir(parents=True, exist_ok=True)
    with open(MAC_PLIST, "wb") as f:
        plistlib.dump({"Label": MAC_LABEL, "ProgramArguments": ["/bin/bash", str(sh)],
                       "StartCalendarInterval": {"Hour": h, "Minute": m}, "RunAtLoad": False,
                       "StandardOutPath": str(LOG), "StandardErrorPath": str(LOG)}, f)
    subprocess.run(["launchctl", "unload", str(MAC_PLIST)], capture_output=True)
    r = subprocess.run(["launchctl", "load", "-w", str(MAC_PLIST)], capture_output=True, text=True)
    if r.returncode == 0:
        print(f"✅ 매일 {hhmm}에 자동 실행하도록 예약했습니다(launchd {MAC_LABEL}). Mac이 잠자기 중이었으면 깨어난 뒤 실행합니다.")
        print("   ℹ 프로젝트가 OneDrive·문서·데스크탑 폴더에 있으면 처음 한 번 ‘폴더 접근 허용’ 창이 뜰 수 있습니다(허용).")
        return True
    print(f"❌ 예약하지 못했습니다: {(r.stderr or r.stdout).strip()[:300]}")
    return False


def _mac_unschedule() -> None:
    _remove_old_mac_plist()
    if not MAC_PLIST.exists():
        print("ℹ 예약된 자동 실행이 없습니다(해제할 것이 없음).")
        return
    subprocess.run(["launchctl", "unload", "-w", str(MAC_PLIST)], capture_output=True)
    MAC_PLIST.unlink(missing_ok=True)
    print("✅ 자동 실행 예약을 해제했습니다.")


def _mac_schedule_text() -> str:
    if not MAC_PLIST.exists():
        return "예약 안 됨"
    try:
        import plistlib
        t = plistlib.loads(MAC_PLIST.read_bytes()).get("StartCalendarInterval", {})
        return f"예약됨(매일 {t.get('Hour', 0):02d}:{t.get('Minute', 0):02d})"
    except Exception:
        return "예약됨"


def fix_bat() -> None:
    """폴더 구조를 바꾼 뒤(프로그램 파일 → app 폴더) 예전 daily_run.bat이 있으면 새 경로로 다시 씀."""
    if not sys.platform.startswith("win"):
        return
    try:
        if BAT.exists() and "app\\main.py" not in BAT.read_text(encoding="utf-8", errors="replace"):
            _write_bat()
    except Exception:
        pass


def _schtasks(args) -> Tuple[int, str]:
    """schtasks 실행. 한국어 Windows는 출력이 CP949라서 바이트로 받아 직접 변환(UTF-8로 읽으면 글자가 깨짐)."""
    r = subprocess.run(args, capture_output=True, creationflags=0x08000000 if sys.platform.startswith("win") else 0)
    raw = (r.stdout or b"") + (r.stderr or b"")
    encs = []
    try:
        import ctypes
        encs.append(f"cp{ctypes.windll.kernel32.GetOEMCP()}")
    except Exception:
        pass
    encs += ["cp949", "mbcs", "utf-8"]
    for enc in encs:
        try:
            return r.returncode, raw.decode(enc).strip()
        except (UnicodeDecodeError, LookupError):
            continue
    return r.returncode, raw.decode("utf-8", errors="replace").strip()


def _exists() -> bool:
    try:
        return _schtasks(["schtasks", "/Query", "/TN", TASK_NAME])[0] == 0
    except FileNotFoundError:
        return False


def schedule(hhmm: str) -> bool:
    if not re.fullmatch(r"([01]?\d|2[0-3]):[0-5]\d", (hhmm or "").strip()):   # 명령줄에 들어가므로 형식만 허용
        print(f"❌ 시간 형식이 올바르지 않습니다: {hhmm!r} (예: 08:10)")
        return False
    hhmm = hhmm.strip()
    if sys.platform == "darwin":
        return _mac_schedule(hhmm)
    bat = _write_bat()
    tr = f'\\"{bat}\\"' if " " in str(bat) else str(bat)
    cmd = f'schtasks /Create /F /SC DAILY /TN {TASK_NAME} /ST {hhmm} /TR "{tr}"'
    try:
        rc, out = _schtasks(cmd)
    except FileNotFoundError:
        print("❌ Windows에서만 예약할 수 있습니다(schtasks 없음).")
        return False
    if rc == 0:
        print(f"✅ 매일 {hhmm}에 자동 실행하도록 예약했습니다(작업 이름 {TASK_NAME}).")
    else:
        print(f"❌ 예약하지 못했습니다: {out}")
    return rc == 0


def unschedule() -> None:
    if sys.platform == "darwin":
        return _mac_unschedule()
    try:
        if not _exists():
            print("ℹ 예약된 자동 실행이 없습니다(해제할 것이 없음).")
            return
        rc, out = _schtasks(["schtasks", "/Delete", "/F", "/TN", TASK_NAME])
    except FileNotFoundError:
        print("❌ Windows에서만 사용할 수 있습니다.")
        return
    print("✅ 자동 실행 예약을 해제했습니다." if rc == 0 else f"⚠ 해제하지 못했습니다: {out}")


def schedule_text() -> str:
    if sys.platform == "darwin":
        return _mac_schedule_text()
    try:
        rc, out = _schtasks(["schtasks", "/Query", "/TN", TASK_NAME, "/FO", "LIST"])
        if rc != 0:
            return "예약 안 됨"
        nxt = [l.split(":", 1)[1].strip() for l in out.splitlines() if "Next Run Time" in l or "다음 실행 시간" in l]
        return f"예약됨(다음 실행: {nxt[0]})" if nxt else "예약됨"
    except Exception:
        return "예약 상태 확인 불가"


def status_text() -> str:
    s = load_settings()
    data = bus.load()
    done = [v for v in data.get("steps", {}).values() if v.get("status") == "완료"]
    today = f"오늘 {len(done)}개 기능 실행" if done else "오늘 아직 실행 안 함"
    return f"{today} · 자동 실행: {schedule_text() if sys.platform.startswith('win') or sys.platform == 'darwin' else '—'} · 시간 {s['time']}"


# ================================================================ 메뉴

def menu() -> None:
    while True:
        s = load_settings()
        print("\n  매일 자동 실행 · 오늘의 요약")
        print(f"   {status_text()}")
        print("   1. 지금 전체 실행(선택된 기능 모두 → 오늘의 요약)")
        print("   2. 오늘의 요약만 다시 만들기(이미 실행한 결과로)")
        print("   3. 실행할 기능 선택")
        print(f"   4. 매일 자동 실행 예약(현재 시간 {s['time']})")
        print("   5. 자동 실행 예약 해제")
        print("   0. 돌아가기")
        c = input("번호: ").strip()
        if c == "1":
            run_all()
        elif c == "2":
            out = build_summary()
            if out:
                try:
                    __import__("platform_util").open_path(str(out))
                except Exception:
                    pass
        elif c == "3":
            for i, (k, n, _) in enumerate(STEPS, 1):
                print(f"   {i}. [{'✓' if k in s['steps'] else ' '}] {n}")
            sel = input("켜고 끌 번호(쉼표로 여러 개, Enter=그대로): ").strip()
            for x in sel.replace(" ", "").split(","):
                if x.isdigit() and 1 <= int(x) <= len(STEPS):
                    k = STEPS[int(x) - 1][0]
                    s["steps"] = [y for y in s["steps"] if y != k] if k in s["steps"] else s["steps"] + [k]
            s["steps"] = [k for k, _, _ in STEPS if k in s["steps"]]
            save_settings(s)
        elif c == "4":
            t = input(f"실행 시간(HH:MM, Enter={s['time']}): ").strip() or s["time"]
            try:
                datetime.strptime(t, "%H:%M")
            except ValueError:
                print("HH:MM 형식으로 입력하세요(예: 08:10).")
                continue
            s["time"] = t
            save_settings(s)
            if schedule(t):
                print("   컴퓨터가 켜져 있고 Windows에 로그인된 상태에서 실행됩니다. 실행 기록: reports\\daily\\daily_run.log")
                print("   실행 중에는 이 프로그램의 브라우저(로그인 창)를 따로 열어 두지 마세요.")
        elif c == "5":
            unschedule()
        elif c in ("0", ""):
            return

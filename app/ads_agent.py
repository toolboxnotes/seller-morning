"""
광고 Agent · 1단계: 자료 수집 + 조치 제안(광고는 바꾸지 않음)

실행 순서(보고서 생성 시간이 일정하지 않아서):
  1. Ads console → Reporting: 필요한 보고서를 먼저 요청(최근 30일, 하루 단위, 매일 자동 실행으로 설정)
  2. 캠페인 관리 표 → Export(전날·최근 7일, 바로 받을 수 있음)
  3. Brand analytics: SQP / SCP / Top Search Terms(그동안 보고서가 뒤에서 생성됨)
  4. Reporting으로 돌아와 생성된 것은 다운로드, 아직이면 2~3분마다 확인(최대 40분)
  5. 분석 → Excel 보고서. 못 받은 보고서는 기록해 두고, 다음 실행 또는 ‘광고 데이터 수집·조치 제안’을 다시 실행하면 남은 것만 받음
"""

import json
import os
import sys
import random
import time
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

import ads_analyze as aa
import ads_collect as acl
import ads_common as ac
import ads_report as ar
import ba_collect as ba
import human
import sc_nav
import sc_reports
from sc_browser import open_session

HISTORY_DIR = ac.REPORT_DIR / "history"


def _step(name: str, fn, notes: List[str]):
    try:
        return fn()
    except Exception as e:
        msg = f"{name} 실패: {type(e).__name__}: {str(e).splitlines()[0][:160]}"
        print(f"   ⚠ {msg}")
        notes.append(msg + " (스크린샷과 페이지 구조는 reports/ads/debug)")
        return None


def _seed_terms(today: date) -> List[str]:
    """Top Search Terms에서 조회할 키워드: SQP 상위 검색어 + 광고비가 많은 키워드."""
    terms: List[str] = []
    p = ba.latest_file("sqp")
    if p:
        for x in sorted(ba.parse_sqp(ac.load_rows(p)), key=lambda x: -x["volume"])[:10]:
            terms.append(x["query"])
    rows, _ = aa._from_days("rp_sp_targeting.xlsx", today)
    spend = {}
    for r in ac.normalize(rows):
        t = (r.get("targeting") or "").strip()
        if t and not aa.is_auto_target(t) and not t.lower().startswith("asin"):
            spend[t] = spend.get(t, 0) + r.get("spend", 0)
    for t, _ in sorted(spend.items(), key=lambda x: -x[1])[:8]:
        if t.lower() not in (x.lower() for x in terms):
            terms.append(t)
    for t in str(_setting("인기 검색어 조회 키워드") or aa.DEFAULTS["인기 검색어 조회 키워드"]).split(","):
        t = t.strip()
        if t and t.lower() not in (x.lower() for x in terms):
            terms.append(t)
    return terms


def _setting(name: str):
    try:
        from openpyxl import load_workbook
        ws = load_workbook(ac.CONFIG_FILE, read_only=True)["설정"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] == name:
                return r[1]
    except Exception:
        return None
    return None


def _brand() -> str:
    import sp_api
    return sp_api.brand()


def prerequest() -> str:
    """매일 자동 실행 시작 때 광고 보고서를 미리 요청(생성에 20~40분 걸리므로 다른 기능이 도는 동안 만들어지게).
    광고 단계에서는 이미 요청된 보고서를 받기만 함. 오늘 이미 받았거나 요청했으면 아무것도 안 함."""
    today = ac.pacific_today()
    st_all = ac.load_state()
    same_day = st_all.get("date") == today.isoformat()
    keys = [k for k, *_ in acl.REPORTS]
    need = [k for k in keys if not (ac.day_dir(today) / f"rp_{k}.xlsx").exists()]
    if not need:
        return "오늘 광고 보고서는 이미 모두 받음"
    if same_day and (st_all.get("collected") or st_all.get("api_ids") or st_all.get("pre_requested")):
        return "오늘 이미 요청함"
    import ads_api
    if ads_api.ready():
        try:
            ids = ads_api.request_only(need, today)
            ac.save_state({"date": today.isoformat(), "reports": st_all.get("reports", {}) if same_day else {},
                           "api_ids": ids, "pre_requested": True})
            return f"Ads API로 {len(ids)}개 미리 요청"
        except Exception as e:
            print(f"   ⚠ Ads API 미리 요청 실패({str(e)[:120]}) → 광고 콘솔 화면으로 요청")
    state = st_all.get("reports", {}) if same_day else {}
    with open_session(headless=False) as page:
        eid = acl.open_ads(page)
        state = acl.request_reports(page, eid, state) or state
    ac.save_state({"date": today.isoformat(), "reports": state, "pre_requested": True})
    return f"광고 콘솔에서 {sum(1 for v in state.values() if v.get('status') in ('waiting', 'ready'))}개 미리 요청"


def collect(download_only: bool = False) -> Dict:
    today = ac.pacific_today()
    st_all = ac.load_state()
    same_day = st_all.get("date") == today.isoformat()
    state = st_all.get("reports", {}) if same_day else {}
    collected = bool(same_day and st_all.get("collected"))
    notes: List[str] = []
    t0 = time.time()
    # 오늘 이미 보고서 요청·키워드 자료 수집을 끝냈으면: 남은 보고서만 받기(같은 날 다시 실행했을 때)
    keys = [k for k, *_ in acl.REPORTS]
    state = {k: v for k, v in state.items() if k in keys}       # 지금은 쓰지 않는 보고서 기록은 지움
    retry: List[str] = []
    # Amazon Ads API로 먼저 받기(설정돼 있으면). 받은 보고서는 브라우저에서 다시 요청·다운로드하지 않음
    need = [k for k in keys if not (ac.day_dir(today) / f"rp_{k}.xlsx").exists()]
    if need:
        import ads_api
        print("\n【0】 광고 보고서(Amazon Ads API)")
        ads_api.try_ads("광고 보고서", ads_api.fetch_reports, need, today, ids_pre=st_all.get("api_ids") if same_day else None)
    for k in keys:
        f = ac.day_dir(today) / f"rp_{k}.xlsx"
        if f.exists():
            state[k] = {**state.get(k, {}), "status": "done", "file": str(f)}
    ac.save_state({"date": today.isoformat(), "reports": state, "collected": collected})
    if collected and not download_only:
        left = [k for k, v in state.items() if v.get("status") not in ("done", "failed")]
        retry = [k for k in keys if state.get(k, {}).get("status") == "failed" or k not in state]
        if not left and not retry and not ba.pending_text():
            print("   오늘 보고서 요청·다운로드와 키워드 자료 수집이 모두 끝났습니다. 브라우저를 열지 않고 분석만 다시 합니다.")
            return {"today": today, "notes": notes, "pending": [], "state": state}
        print(f"   오늘 이미 보고서를 요청하고 키워드 자료를 수집했습니다. 아직 받지 못한 것만 받습니다"
              f"(광고 보고서 {len(left)}개" + (f", Brand Analytics {ba.pending_text()}" if ba.pending_text() else "") + ")."
              + (f" 요청에 실패했던 보고서 {len(retry)}개는 다시 요청합니다." if retry else ""))
        download_only = True
    with open_session(headless=False) as page:
        print("\n【1】 Ads console")
        eid = _step("Ads console 열기", lambda: acl.open_ads(page), notes) or ""
        if retry:
            print(f"\n【2】 실패했던 광고 보고서 다시 요청: {', '.join(acl.report_name(k) for k in retry)}")
            state = _step("보고서 다시 요청", lambda: acl.request_reports(page, eid, state, only=retry), notes) or state
            ac.save_state({"date": today.isoformat(), "reports": state, "collected": True})
        if not download_only:
            print("\n【2】 광고 보고서 요청(Reporting)")
            if any(state.get(k, {}).get("status") != "done" for k in keys):
                state = _step("보고서 요청", lambda: acl.request_reports(page, eid, state), notes) or state
            else:
                print("   광고 보고서는 모두 Ads API로 받았습니다(화면에서 요청하지 않음)")
            ac.save_state({"date": today.isoformat(), "reports": state})
            print("\n【3】 Brand Analytics 키워드 자료(보고서가 생성되는 동안)")
            human.goto(page, sc_nav.SC_BASE + "/home", wait_after=False)
            ac.settle(page)
            for key in ("sqp", "scp"):
                _step(f"Brand Analytics {key.upper()}", lambda k=key: ba.fetch_dashboard(page, k), notes)
            _step("Top Search Terms", lambda: ba.fetch_top_search_terms(page, _seed_terms(today), _brand()), notes)
            print("\n【3-2】 브랜드 노출 점유율 · 광고 추천(보기만 함)")
            import sc_pages
            for pk in ("ads_bis", "ads_recs"):
                _step(sc_pages.NAMES[pk], lambda pk=pk: sc_pages.read(page, pk, entity_id=eid), notes)
            ac.save_state({"date": today.isoformat(), "reports": state, "collected": True})
        print("\n【4】 생성된 보고서 받기(광고 보고서 + Brand Analytics 전체 다운로드)")
        limit = acl.REPORT_WAIT_MIN
        while True:
            if any(state.get(k, {}).get("status") not in ("done", "failed") for k in keys):
                eid = _step("Ads console 열기", lambda: acl.open_ads(page), notes) or eid
                state = _step("광고 보고서 다운로드", lambda: acl.collect_reports(page, eid, state, wait_min=0), notes) or state
            human.goto(page, sc_nav.SC_BASE + "/home", wait_after=False)
            ac.settle(page)
            _step("Brand Analytics 다운로드", lambda: ba.collect_downloads(page, 0), notes)
            ads_left = [k for k, v in state.items() if v.get("status") not in ("done", "failed")]
            ba_left = ba.pending_text()
            used = (time.time() - t0) / 60
            if (not ads_left and not ba_left) or used >= limit:
                break
            wait = random.uniform(sc_reports.POLL_MIN_S, sc_reports.POLL_MAX_S)
            print(f"   아직 생성 중: 광고 보고서 {len(ads_left)}개" + (f", Brand Analytics {ba_left}" if ba_left else "") +
                  f" · {used:.0f}분 경과, {wait / 60:.1f}분 후 다시 확인(최대 {limit}분)")
            page.wait_for_timeout(int(wait * 1000))
    ac.save_state({"date": today.isoformat(), "reports": state, "collected": collected or not download_only})
    pending = [acl.report_name(k) for k, v in state.items() if v.get("status") not in ("done",)]
    if ba.pending_text():
        pending.append("Brand Analytics 전체 다운로드: " + ba.pending_text())
    return {"today": today, "notes": notes, "pending": pending, "state": state}


def analyze(today: date, notes: List[str], pending: List[str]) -> Path:
    print("\n【5】 분석 및 조치 제안")
    data = aa.load_all(today)
    data["notes"] = notes + data["notes"]
    adv = aa.Advisor(data, today)
    res = adv.run()
    k = aa.kpis(data, today)
    out = ac.REPORT_DIR / f"ads_report_{today}.xlsx"
    ar.write(out, today, data, res, k, pending)
    track = {}
    try:                                   # 지난 제안을 실행했는지 · 효과
        import ads_track
        t_rows, track = ads_track.check(today)
        if t_rows:
            ads_track.append_sheet(out, t_rows, track)
            print(f"   지난 제안 {track.get('확인한 제안', 0)}건: 실행됨 {track.get('실행됨', 0)} · 실행 안 됨 {track.get('실행 안 된 제안', 0)}"
                  f" · 실행률 {track.get('실행률', '—')} · 주당 절약(추정) {track.get('주당 절약(추정)', '')}")
    except Exception as e:
        print(f"   ⚠ 지난 제안 이행 확인 실패: {type(e).__name__}: {str(e)[:120]}")
    pages = []
    try:
        import sc_pages
        pages = [p for p in (sc_pages.cached("ads_bis"), sc_pages.cached("ads_recs")) if p]
        sc_pages.append_sheets(out, pages)
    except Exception as e:
        print(f"   ⚠ 브랜드 노출·광고 추천 시트를 추가하지 못했습니다: {type(e).__name__}")
    try:
        import daily_bus
        y_, w_ = k.get("yesterday") or {}, k.get("7d") or {}
        fa_ = lambda v: "—" if v in (None, float("inf")) else f"{v:.1%}"
        kp = {"전날 광고비": round(y_.get("spend", 0), 2), "전날 광고 매출": round(y_.get("sales", 0), 2), "전날 ACoS": fa_(y_.get("acos")),
              "7일 광고비": round(w_.get("spend", 0), 2), "7일 광고 매출": round(w_.get("sales", 0), 2), "7일 ACoS": fa_(w_.get("acos")),
              "조치 제안 수": len(res["actions"])}
        if track:
            kp["지난 제안 실행률"] = track.get("실행률", "—")
            kp["지난 제안 주당 절약(추정)"] = track.get("주당 절약(추정)", "")
        al = [(a["level"] if a["level"] in daily_bus.LEVELS else "P1", a["what"], f"{a['object']}: {a['detail']}", "광고 보고서 확인")
              for a in res["alerts"]]
        from collections import Counter as _C
        al += [("P2", "광고 조치 제안", f"{t} {n}건", "광고 보고서 ‘조치 제안’ 시트에서 확인 후 직접 적용")
               for t, n in _C(a["type"] for a in res["actions"]).most_common(6)]
        if track and track.get("실행 안 된 제안"):
            al.append(("P3", "지난 광고 제안 미실행", f"최근 30일 제안 중 실행 안 된 것 {track['실행 안 된 제안']}건(실행률 {track.get('실행률')})",
                       "광고 보고서 ‘지난 제안 이행·효과’ 시트 확인"))
        if pending:
            al.append(("P2", "광고 보고서 대기", ", ".join(pending)[:150], "나중에 ‘광고 데이터 수집·조치 제안’을 다시 실행(남은 보고서만 받음)"))
        daily_bus.record("광고", kp, al, out)
        daily_bus.merge_pages("광고", pages)
    except Exception:
        pass
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    (HISTORY_DIR / f"suggestions_{today}.json").write_text(
        json.dumps({"date": today.isoformat(), "created": datetime.now().isoformat(), "actions": res["actions"],
                    "alerts": res["alerts"]}, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    # 터미널 요약
    if k.get("7d"):
        y, w = k.get("yesterday") or {}, k["7d"]
        fa = lambda v: "—" if v in (None, float("inf")) else f"{v:.1%}"
        print(f"\n전날: 광고비 ${y.get('spend', 0):,.2f}, 매출 ${y.get('sales', 0):,.2f}, ACoS {fa(y.get('acos'))}")
        print(f"최근 7일: 광고비 ${w['spend']:,.2f}, 매출 ${w['sales']:,.2f}, ACoS {fa(w.get('acos'))}, 주문 {w['orders']:,.0f}")
    from collections import Counter
    cnt = Counter(a["type"] for a in res["actions"])
    print(f"조치 제안 {len(res['actions'])}건: " + ", ".join(f"{t} {n}" for t, n in cnt.most_common(8)))
    for al in res["alerts"][:5]:
        print(f"  [{al['level']}] {al['what']} · {al['object']}: {al['detail']}")
    if pending:
        print(f"⚠ 아직 생성 중인 보고서: {', '.join(pending)} → 나중에 ‘광고 데이터 수집·조치 제안’을 다시 실행하면 남은 보고서만 받습니다.")
    print(f"\n✅ 보고서 저장 완료: {out}")
    return out


def run(open_report: bool = True) -> Optional[Path]:
    ac.REPORT_DIR.mkdir(parents=True, exist_ok=True)
    print("   1단계: 자료 수집과 조치 제안만 합니다. 광고 설정은 바꾸지 않습니다.")
    r = collect()
    out = analyze(r["today"], r["notes"], r["pending"])
    _open(out, open_report)
    return out


def run_download_only(open_report: bool = True) -> Optional[Path]:
    """생성 중이던 광고 보고서만 받고, 분석을 다시 합니다."""
    r = collect(download_only=True)
    out = analyze(r["today"], r["notes"], r["pending"])
    _open(out, open_report)
    return out


def run_analyze_only(open_report: bool = True) -> Optional[Path]:
    """이미 받은 자료로 분석만 다시 합니다(ads_config.xlsx를 고친 뒤 등)."""
    out = analyze(ac.pacific_today(), [], [])
    _open(out, open_report)
    return out


def _open(out: Path, open_report: bool) -> None:
    if open_report:
        try:
            __import__("platform_util").open_path(str(out))
        except Exception:
            pass


def status_text() -> str:
    st = ac.load_state()
    files = sorted(ac.REPORT_DIR.glob("ads_report_*.xlsx"))
    last = f"마지막 보고서: {files[-1].stem.replace('ads_report_', '')}" if files else "아직 실행하지 않음"
    if st.get("date") == ac.pacific_today().isoformat():
        pend = [k for k, v in st.get("reports", {}).items() if v.get("status") != "done"]
        if pend:
            last += f" · 생성 중인 보고서 {len(pend)}개"
    return last


if __name__ == "__main__":
    run()

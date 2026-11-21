"""
지난 광고 조치 제안 — 실행했는지 · 실행 후 효과(조회만)

reports/ads/history/suggestions_날짜.json(지금까지의 제안) 과 그 뒤에 받은 광고 보고서를 비교:
  부정 키워드      → 제안 뒤 그 캠페인에서 그 검색어 클릭이 없으면 ‘실행됨’, 아낀 지출 추정
  키워드·상품 추가 → 제안 뒤 정확(또는 상품) 타겟으로 나타나면 ‘실행됨’, 새 타겟의 성과
  일시중지         → 제안 뒤 그 키워드 노출이 없으면 ‘실행됨’
  입찰가 인하·인상 → 제안 뒤 입찰가가 제안 쪽으로 바뀌면 ‘실행됨’, 전 7일 vs 후 7일 ACoS·지출·주문
  예산 증액        → 제안 뒤 예산이 늘었으면 ‘실행됨’
  게재 위치·리스팅 등 → 보고서로 확인할 수 없어 ‘확인 불가’
제안 뒤 최소 3일 데이터가 있어야 판단(그 전은 ‘판단 대기’). 최근 30일 제안만.
"""

import json
from collections import defaultdict
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

import ads_analyze as aa
import ads_common as ac

HISTORY = ac.REPORT_DIR / "history"
FILES = {"st": "rp_sp_search_term.xlsx", "tg": "rp_sp_targeting.xlsx", "cp": "rp_sp_campaign.xlsx"}
KEYS = ("date", "campaign", "ad_group", "targeting", "match_type", "search_term", "placement", "asin")


def merged(name: str, today: date, days: int = 45) -> List[Dict]:
    """날짜 폴더마다 받은 보고서를 합침(같은 날·같은 줄은 나중에 받은 것 사용)."""
    out: Dict[Tuple, Dict] = {}
    if not ac.DATA_DIR.exists():
        return []
    for d in sorted(p for p in ac.DATA_DIR.iterdir() if p.is_dir()):
        try:
            dd = date.fromisoformat(d.name)
        except ValueError:
            continue
        if dd < today - timedelta(days=days) or dd > today:
            continue
        f = d / name
        if not f.exists():
            continue
        try:
            rows = ac.normalize(ac.load_rows(f))
        except Exception:
            continue
        for r in rows:
            if r.get("date"):
                out[tuple(str(r.get(k) or "").strip().lower() for k in KEYS)] = r
    return list(out.values())


def _sum(rows: List[Dict]) -> Dict:
    a = {"impressions": 0.0, "clicks": 0.0, "spend": 0.0, "sales": 0.0, "orders": 0.0}
    for r in rows:
        for k in a:
            a[k] += r.get(k, 0) or 0
    a["acos"] = a["spend"] / a["sales"] if a["sales"] else None
    return a


def _between(rows, d0: date, d1: date):
    return [r for r in rows if r.get("date") and d0.isoformat() <= r["date"][:10] <= d1.isoformat()]


def _fmt(a: Dict) -> str:
    if not a or not (a["clicks"] or a["spend"]):
        return "활동 없음"
    acos = "—" if a["acos"] is None else f"{a['acos']:.0%}"
    return f"지출 ${a['spend']:.2f} · 클릭 {a['clicks']:.0f} · 주문 {a['orders']:.0f} · ACoS {acos}"


def load_suggestions(today: date, days: int = 30) -> List[Dict]:
    """같은 조치(유형·캠페인·광고 그룹·대상)는 처음 제안한 날 기준, 제안 횟수 기록."""
    first: Dict[Tuple, Dict] = {}
    if not HISTORY.exists():
        return []
    for f in sorted(HISTORY.glob("suggestions_*.json")):
        try:
            d = date.fromisoformat(f.stem.replace("suggestions_", ""))
        except ValueError:
            continue
        if d >= today or d < today - timedelta(days=days):
            continue
        try:
            acts = json.loads(f.read_text(encoding="utf-8")).get("actions", [])
        except Exception:
            continue
        for a in acts:
            k = (a.get("type"), a.get("campaign", ""), a.get("ad_group", ""), str(a.get("target", "")).lower())
            if k not in first:
                first[k] = dict(a, _date=d, _count=1)
            else:
                first[k]["_count"] += 1
                first[k]["_last"] = d
    return list(first.values())


def check(today: Optional[date] = None) -> Tuple[List[Dict], Dict]:
    today = today or ac.pacific_today()
    sugg = load_suggestions(today)
    if not sugg:
        return [], {}
    st, tg, cp = (merged(FILES[k], today) for k in ("st", "tg", "cp"))
    last_day = max([r["date"][:10] for r in st + tg + cp if r.get("date")] or [today.isoformat()])
    last = date.fromisoformat(last_day)
    out = []
    for a in sugg:
        t, c, g = a.get("type", ""), a.get("campaign", ""), a.get("ad_group", "")
        tgt = str(a.get("target", "")).strip()
        tl = tgt.lower()
        d = a["_date"]
        a0, a1 = d + timedelta(days=1), min(d + timedelta(days=7), last)
        b0, b1 = d - timedelta(days=7), d - timedelta(days=1)
        row = {"제안일": d.isoformat(), "제안 횟수": a.get("_count", 1), "조치 유형": t, "캠페인": c, "광고 그룹": g, "대상": tgt,
               "제안": a.get("suggested", ""), "실행": "", "판단 근거": "", "제안 전 7일": "", "제안 후": "", "평가": ""}
        days_after = (last - d).days
        if days_after < 3:
            row["실행"], row["판단 근거"] = "판단 대기", f"제안 후 데이터 {max(days_after, 0)}일(3일 이상 필요)"
            out.append(row)
            continue
        row["제안 후"] = f"({a0}~{a1})"
        if t.startswith("부정 키워드") and t != "부정 키워드 검토":
            srcs = [x.strip() for x in c.split(",")] if c else []
            rel = [r for r in st if (r.get("search_term") or "").strip().lower() == tl and (not srcs or r.get("campaign") in srcs)]
            if t == "부정 키워드(구문)":
                rel = [r for r in st if tl in (r.get("search_term") or "").lower()]
            before, after = _sum(_between(rel, b0, b1)), _sum(_between(rel, a0, a1))
            camp_after = _between([r for r in st if not srcs or r.get("campaign") in srcs], a0, a1)
            row["제안 전 7일"], row["제안 후"] = _fmt(before), _fmt(after) + " " + row["제안 후"]
            if after["clicks"] == 0 and camp_after:
                row["실행"], row["판단 근거"] = "실행됨(추정)", "제안 뒤 이 캠페인에서 이 검색어 클릭 없음"
                row["평가"] = f"낭비 지출 약 ${before['spend']:.2f}/주 절약" if before["spend"] else "효과 측정 어려움(이전 지출 적음)"
            elif after["clicks"] > 0:
                row["실행"], row["판단 근거"] = "실행 안 됨", f"제안 뒤에도 클릭 {after['clicks']:.0f}, 지출 ${after['spend']:.2f}"
                row["평가"] = "주문 없이 계속 지출 중 → 부정 키워드 추가 필요" if after["orders"] == 0 else "제안 뒤 주문 발생 → 다시 판단"
            else:
                row["실행"], row["판단 근거"] = "확인 불가", "제안 뒤 이 캠페인 검색어 데이터 없음"
        elif t in ("키워드 추가(정확)", "상품 타겟팅 추가"):
            want = "exact" if t == "키워드 추가(정확)" else "product"
            hit = [r for r in tg if (r.get("targeting") or "").strip().lower().strip('"').replace('asin="', "").rstrip('"') == tl
                   or (r.get("targeting") or "").strip().lower() == tl]
            hit = [r for r in hit if want == "product" or aa.norm_match(r.get("match_type", "")) == "exact"]
            after = _sum(_between(hit, a0, a1))
            if _between(hit, a0, a1):
                camps = sorted({r.get("campaign", "") for r in _between(hit, a0, a1)})
                row["실행"], row["판단 근거"] = "실행됨", f"정확 타겟으로 운영 중: {', '.join(camps)[:80]}"
                row["제안 후"] = _fmt(after) + " " + row["제안 후"]
                row["평가"] = ("좋음(주문 발생)" if after["orders"] else "아직 주문 없음 — 입찰가·관련성 확인") if after["clicks"] else "노출·클릭 적음 — 입찰가 확인"
            else:
                row["실행"], row["판단 근거"] = "실행 안 됨", "제안 뒤 정확 타겟 데이터에 이 키워드가 없음"
            src_rows = [r for r in st if (r.get("search_term") or "").strip().lower() == tl]
            row["제안 전 7일"] = "검색어로: " + _fmt(_sum(_between(src_rows, b0, b1)))
        elif t == "키워드 일시중지":
            rel = [r for r in tg if r.get("campaign") == c and r.get("ad_group", "") == g and (r.get("targeting") or "").lower() == tl]
            before, after = _sum(_between(rel, b0, b1)), _sum(_between(rel, a0, a1))
            row["제안 전 7일"], row["제안 후"] = _fmt(before), _fmt(after) + " " + row["제안 후"]
            if after["impressions"] == 0:
                row["실행"], row["판단 근거"] = "실행됨(추정)", "제안 뒤 노출 없음"
                row["평가"] = f"주문 없는 지출 약 ${before['spend']:.2f}/주 절약" if before["spend"] and not before["orders"] else ""
            else:
                row["실행"], row["판단 근거"] = "실행 안 됨", f"제안 뒤에도 노출 {after['impressions']:.0f}, 지출 ${after['spend']:.2f}"
        elif t in ("입찰가 인하", "입찰가 인상"):
            rel = [r for r in tg if r.get("campaign") == c and r.get("ad_group", "") == g and (r.get("targeting") or "").lower() == tl]
            ra = sorted(_between(rel, a0, a1), key=lambda r: r["date"])
            before, after = _sum(_between(rel, b0, b1)), _sum(ra)
            row["제안 전 7일"], row["제안 후"] = _fmt(before), _fmt(after) + " " + row["제안 후"]
            cur, sug = ac.num(a.get("current")), ac.num(a.get("suggested"))
            bids = [r.get("bid") for r in ra if r.get("bid")]
            nb = bids[-1] if bids else None
            if nb is None:
                row["실행"], row["판단 근거"] = "확인 불가", "보고서에 입찰가 정보 없음" if ra else "제안 뒤 데이터 없음"
            elif (t == "입찰가 인하" and nb < cur - 0.005) or (t == "입찰가 인상" and nb > cur + 0.005):
                row["실행"], row["판단 근거"] = "실행됨", f"입찰가 ${cur:.2f} → ${nb:.2f}(제안 ${sug:.2f})"
                if before["acos"] is not None and after["acos"] is not None:
                    ch = after["acos"] / before["acos"] - 1 if before["acos"] else 0
                    row["평가"] = ("좋아짐" if (t == "입찰가 인하" and ch <= -0.1) or (t == "입찰가 인상" and after["orders"] > before["orders"])
                                 else "나빠짐" if ch >= 0.15 else "큰 변화 없음") + f"(ACoS {before['acos']:.0%} → {after['acos']:.0%})"
            else:
                row["실행"], row["판단 근거"] = "실행 안 됨", f"입찰가 그대로 ${nb:.2f}"
        elif t == "예산 증액":
            rel = [r for r in cp if r.get("campaign") == c]
            ra = sorted(_between(rel, a0, a1), key=lambda r: r["date"])
            buds = [r.get("budget") for r in ra if r.get("budget")]
            cur = ac.num(a.get("current"))
            if not buds:
                row["실행"], row["판단 근거"] = "확인 불가", "보고서에 예산 정보 없음"
            elif buds[-1] > cur + 0.01:
                row["실행"], row["판단 근거"] = "실행됨", f"예산 ${cur:.2f} → ${buds[-1]:.2f}"
                b, af = _sum(_between(rel, b0, b1)), _sum(ra)
                row["제안 전 7일"], row["제안 후"] = _fmt(b), _fmt(af) + " " + row["제안 후"]
                row["평가"] = "주문 증가" if af["orders"] > b["orders"] * (len(ra) / 7 or 1) else "주문 증가 없음"
            else:
                row["실행"], row["판단 근거"] = "실행 안 됨", f"예산 그대로 ${buds[-1]:.2f}"
        else:
            row["실행"], row["판단 근거"] = "확인 불가", "광고 보고서로 확인할 수 없는 조치(게재 위치·리스팅 등) — 직접 확인"
        out.append(row)
    order = {"실행 안 됨": 0, "실행됨": 1, "실행됨(추정)": 1, "판단 대기": 2, "확인 불가": 3}
    out.sort(key=lambda r: (order.get(r["실행"], 4), r["제안일"]))
    judged = [r for r in out if r["실행"] not in ("판단 대기", "확인 불가")]
    done = [r for r in judged if r["실행"].startswith("실행됨")]
    by_type = defaultdict(lambda: [0, 0])
    for r in judged:
        by_type[r["조치 유형"]][0] += 1
        by_type[r["조치 유형"]][1] += r["실행"].startswith("실행됨")
    saved = sum(float(x.split("$")[1].split("/")[0]) for x in (r["평가"] for r in done) if "절약" in x and "$" in x)
    summary = {"확인한 제안": len(out), "판단 가능": len(judged), "실행됨": len(done),
               "실행률": f"{len(done) / len(judged):.0%}" if judged else "—",
               "실행 안 된 제안": sum(1 for r in judged if r["실행"] == "실행 안 됨"),
               "주당 절약(추정)": f"${saved:.2f}",
               "유형별 실행": " · ".join(f"{k} {v[1]}/{v[0]}" for k, v in sorted(by_type.items(), key=lambda x: -x[1][0]))}
    return out, summary


def append_sheet(path, rows: List[Dict], summary: Dict) -> None:
    """광고 보고서에 ‘지난 제안 이행·효과’ 시트 추가."""
    import inventory as inv
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    wb = load_workbook(path)
    if "지난 제안 이행·효과" in wb.sheetnames:
        del wb["지난 제안 이행·효과"]
    ws = wb.create_sheet("지난 제안 이행·효과", 1)
    F = inv.FONT
    ws.cell(1, 1, "지난 30일 조치 제안 — 실행했는지 · 효과").font = Font(name=F, size=13, bold=True, color="1F3864")
    ws.cell(2, 1, "광고 보고서로 추정한 결과입니다(실행 여부는 이후 데이터로 판단). ‘실행 안 됨’ 중 지금도 유효한 것은 오늘 ‘조치 제안’에도 다시 나옵니다.").font = Font(name=F, color="6B6B6B")
    r = 3
    for k, v in summary.items():
        ws.cell(r, 1, k).font = Font(name=F, bold=True)
        ws.cell(r, 2, v).font = Font(name=F)
        r += 1
    head = ["실행", "제안일", "제안 횟수", "조치 유형", "캠페인", "광고 그룹", "대상", "제안", "판단 근거", "제안 전 7일", "제안 후", "평가"]
    start = r + 1
    tmp = wb.create_sheet("_tmp")
    inv.write_table_sheet(tmp, head, rows, {"캠페인": 32, "제안": 30, "판단 근거": 40, "제안 전 7일": 40, "제안 후": 50, "평가": 36})
    fills = {"실행 안 됨": PatternFill("solid", fgColor="FFE8CC"), "실행됨": PatternFill("solid", fgColor="EAF6EC"),
             "실행됨(추정)": PatternFill("solid", fgColor="EAF6EC")}
    for i, row in enumerate(tmp.iter_rows(values_only=True)):
        for j, v in enumerate(row, 1):
            c = ws.cell(start + i, j, v)
            c.font = Font(name=F, bold=(i == 0), size=10)
            if i == 0:
                c.fill = PatternFill("solid", fgColor="DCE6F1")
            elif row[0] in fills:
                c.fill = fills[row[0]]
    for col, dim in tmp.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width
    del wb["_tmp"]
    ws.freeze_panes = ws.cell(start + 1, 1)
    wb.save(path)

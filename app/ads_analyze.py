"""
광고 자료 분석 → 필요한 조치(1단계: 제안만, 실행하지 않음)

기준: 《Amazon Ads Agent 관리 규격서》
  - 입찰 규칙(5절, 최근 14일, 최근 3일 제외)
  - 부정 키워드 / 검색어 수확 / 내부 경쟁(5·6절)
  - 매치 유형, 예산, 게재 위치(7절)
  - 재고 / Buy Box 연동, Listing(제목·메인 이미지·가격) 개선(4·8절)
  - Brand Analytics 키워드 기회(SQP / SCP / Top Search Terms)
제안마다 등급(L1 소폭·되돌릴 수 있음 / L2 승인 필요 / L3 사람만)과 근거 데이터·기간을 표시.
"""

import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import ads_common as ac
import ba_collect as ba
import inventory as inv

# ---------------------------------------------------------------- 기본값(ads_config.xlsx에서 바꿀 수 있음)
DEFAULTS = {
    "기본 목표 ACoS(%)": 30.0,          # 원가 정보가 없을 때 사용
    "목표 ACoS 계수 k": 0.7,            # 목표 ACoS = 손익분기 ACoS × k
    "부정 키워드 최소 클릭": 15,
    "어근 부정 최소 클릭": 40,
    "일시중지 최소 클릭": 30,
    "수확 최소 주문": 2,
    "단일 조정 상한(%)": 20,
    "최소 입찰가($)": 0.20,
    "최대 입찰가($)": 5.00,
    "재고 경고 일수": 14,
    "브랜드 키워드": "",          # 비어 있으면 설정의 브랜드 이름 사용
    "인기 검색어 조회 키워드": "fishing lures, bass fishing lures, frog lure, fishing gifts for men, glide bait",
}
STOP = set("for and the with of to in on a an by from men women kids set pack pcs piece pieces new best".split())
MATCH_KO = {"exact": "정확", "phrase": "구문", "broad": "확장"}


# ================================================================ 설정 파일

def load_config(campaigns: List[str], asins: Dict[str, Dict]) -> Dict:
    """ads_config.xlsx: 설정, ASIN 원가, 캠페인 용도, 보호 키워드. 없으면 새로 만들고, 새 캠페인·ASIN은 자동으로 추가."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
    path = ac.CONFIG_FILE
    ac.REPORT_DIR.mkdir(parents=True, exist_ok=True)
    if path.exists():
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.active.title = "설정"
    head = PatternFill("solid", fgColor="DCE6F1")

    def sheet(name, headers, widths):
        ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        if ws.cell(row=1, column=1).value is None:
            for c, h in enumerate(headers, start=1):
                x = ws.cell(row=1, column=c, value=h)
                x.font, x.fill = Font(name=inv.FONT, bold=True), head
                ws.column_dimensions[x.column_letter].width = widths[c - 1]
            ws.freeze_panes = "A2"
        return ws

    st = sheet("설정", ["항목", "값", "설명"], [24, 16, 60])
    have = {st.cell(row=r, column=1).value for r in range(2, st.max_row + 1)}
    notes = {"기본 목표 ACoS(%)": "원가 정보가 없는 상품에 쓰는 목표 ACoS",
             "목표 ACoS 계수 k": "목표 ACoS = 손익분기 ACoS × k (이익 단계 0.6~0.8, 신상품 1.0~1.5)",
             "브랜드 키워드": "쉼표로 구분. 이 단어가 들어간 검색어는 부정 키워드로 제안하지 않음",
             "인기 검색어 조회 키워드": "Top Search Terms에서 조회할 기본 키워드(SQP 상위 검색어와 함께 조회)"}
    for k, v in DEFAULTS.items():
        if k not in have:
            st.append([k, v, notes.get(k, "")])
    cfg = {st.cell(row=r, column=1).value: st.cell(row=r, column=2).value for r in range(2, st.max_row + 1)}

    cs = sheet("ASIN 원가", ["ASIN", "SKU", "상품명", "판매가($)", "원가 COGS($)", "물류·관세($)", "판매 수수료(%)",
                            "FBA 수수료($)", "기타($)", "손익분기 ACoS(%)", "목표 ACoS(%)"],
               [13, 18, 40, 10, 12, 12, 12, 12, 9, 14, 12])
    have = {cs.cell(row=r, column=1).value for r in range(2, cs.max_row + 1)}
    for a, info in sorted(asins.items()):
        if a not in have:
            cs.append([a, info.get("sku", ""), (info.get("title") or "")[:80], info.get("price") or None,
                       None, None, 15, None, None, None, None])
    costs = {}
    k = float(cfg.get("목표 ACoS 계수 k") or 0.7)
    for r in range(2, cs.max_row + 1):
        a = cs.cell(row=r, column=1).value
        vals = [cs.cell(row=r, column=c).value for c in range(4, 10)]
        price, cogs, freight, fee, fba_fee, other = [ac.num(v) if v not in (None, "") else None for v in vals]
        price = price or (asins.get(a, {}).get("price"))
        if price and cogs is not None:
            be = (price - cogs - (freight or 0) - price * (fee or 15) / 100 - (fba_fee or 0) - (other or 0)) / price
            cs.cell(row=r, column=10, value=round(be * 100, 1))
            tgt = cs.cell(row=r, column=11).value
            costs[a] = {"be": be, "target": ac.num(tgt) / 100 if tgt not in (None, "") else max(be * k, 0.05)}
            if tgt in (None, ""):
                cs.cell(row=r, column=11, value=round(costs[a]["target"] * 100, 1))
        elif cs.cell(row=r, column=11).value not in (None, ""):
            costs[a] = {"be": None, "target": ac.num(cs.cell(row=r, column=11).value) / 100}

    ps = sheet("캠페인 용도", ["캠페인", "추정 용도", "확정 용도(직접 입력)", "설명"], [45, 14, 18, 50])
    have = {ps.cell(row=r, column=1).value: r for r in range(2, ps.max_row + 1)}
    purposes = {}
    for c in campaigns:
        if c not in have:
            ps.append([c, "", "", "RESEARCH(자동) / DISCOVERY(확장·구문) / PERF(정확) / COMPETITOR(상품 타겟팅) / DEFENSE(브랜드) / 비워두면 추정값 사용"])
            have[c] = ps.max_row
    for c, r in have.items():
        fixed = ps.cell(row=r, column=3).value
        purposes[c] = {"row": r, "fixed": (str(fixed).strip().upper() if fixed else "")}

    ks = sheet("보호 키워드", ["키워드(부정 제안 안 함)", "메모"], [40, 40])
    protect = {str(ks.cell(row=r, column=1).value).strip().lower() for r in range(2, ks.max_row + 1)
               if ks.cell(row=r, column=1).value}
    brand_words = [w.strip().lower() for w in str(cfg.get("브랜드 키워드") or "").split(",") if w.strip()]
    if not brand_words:
        import sp_api
        b = (sp_api.load_config().get("brand") or "").strip().lower()
        brand_words = [b] if b else []
    return {"wb": wb, "path": path, "cfg": cfg, "costs": costs, "purposes": purposes, "purpose_sheet": ps,
            "protect": protect, "brand": brand_words}


def save_config(conf: Dict, guessed: Dict[str, str]) -> None:
    ps = conf["purpose_sheet"]
    for c, info in conf["purposes"].items():
        if c in guessed:
            ps.cell(row=info["row"], column=2, value=guessed[c])
    try:
        conf["wb"].save(conf["path"])
    except PermissionError:
        print(f"   ⚠ {conf['path'].name} 파일이 Excel에서 열려 있어 저장하지 못했습니다(다음에 다시 저장).")


# ================================================================ 자료 읽기

def _latest(folder: Path, pattern: str) -> Optional[Path]:
    files = sorted(folder.glob(pattern))
    return files[-1] if files else None


def _from_days(name: str, today: date, max_back: int = 3) -> Tuple[List[Dict], Optional[str]]:
    """오늘 데이터가 없으면 최근 며칠 것을 찾음(날짜 표시)."""
    for i in range(max_back + 1):
        d = today - timedelta(days=i)
        p = ac.DATA_DIR / d.isoformat() / name
        if p.exists():
            rows = ac.load_rows(p)
            if rows:
                return rows, (None if i == 0 else d.isoformat())
    return [], None


def load_all(today: date) -> Dict:
    data = {"notes": [], "sources": {}}
    for key, label in (("sp_search_term", "검색어 보고서(SP·SB)"), ("sp_targeting", "타겟팅 보고서"),
                       ("sp_campaign", "캠페인·게재 위치 보고서"), ("sp_advertised_product", "광고 상품 보고서")):
        rows, stale = _from_days(f"rp_{key}.xlsx", today)
        data[key] = ac.normalize(rows, {"ad_type": "SB" if key.startswith("sb") else "SP"})
        for r in data[key]:                       # 새 형식 보고서에는 Ad product 열이 있음: SP / SB / SD
            ap_ = (r.get("ad_product") or "").lower()
            if ap_:
                r["ad_type"] = "SB" if "brand" in ap_ else ("SD" if "display" in ap_ else ("SP" if "product" in ap_ else ap_.upper()))
        if rows:
            data["sources"][label] = f"{len(rows)}행" + (f"({stale}에 받은 데이터)" if stale else "")
            if stale:
                data["notes"].append(f"{label}: 오늘 보고서가 아직 생성되지 않아 {stale}에 받은 데이터를 사용했습니다.")
        else:
            data["sources"][label] = "없음(아직 생성되지 않았거나 받지 못함)"
    data["sb_search_term"] = []      # 새 형식 검색어 보고서에 SB도 들어 있음(Ad product 열로 구분)
    for view in ("campaigns", "ad_groups", "targeting", "search_terms"):
        for k in ("yesterday", "7d"):
            rows, stale = _from_days(f"cm_{view}_{k}.xlsx", today, 0)
            data[f"cm_{view}_{k}"] = ac.normalize(rows)
            if rows:
                data["sources"][f"광고 관리자 {view} ({'전날' if k == 'yesterday' else '최근 7일'})"] = f"{len(rows)}행"
    # Brand Analytics(최근 1주)
    for key in ("sqp", "scp", "tst"):
        p = ba.latest_file(key)
        rows = ac.load_rows(p) if p else []
        data[key] = ba.parse_sqp(rows) if key == "sqp" else (ba.parse_scp(rows) if key == "scp" else rows)
        if p:
            part = " · 페이지 표(전체 다운로드 전 임시)" if p.name.endswith("_page.xlsx") else ""
            data["sources"][f"Brand Analytics {key.upper()}"] = f"{p.name} ({len(rows)}행{part})"
    # 매장 자료(이미 있는 최신 파일 사용, 다시 받지 않음)
    inv_dir = ac.BASE_DIR / "reports" / "inventory"
    data["listings"], data["fba"] = {}, {}
    pl, pf = _latest(inv_dir, "all_listings_*.xlsx"), _latest(inv_dir, "fba_inventory_*.xlsx")
    al = inv.read_xlsx(pl)[1] if pl else []
    fb = inv.read_xlsx(pf)[1] if pf else []
    for it in inv.build_rows(al, fb):
        a = str(it.get("asin") or "").strip()
        if not a:
            continue
        x = data["listings"].setdefault(a, {"title": it.get("title") or "", "price": it.get("price") or 0,
                                            "sku": it.get("sku", ""), "status": it.get("status", "")})
        if not x["title"]:
            x["title"] = it.get("title") or ""
        if it.get("in_fba"):
            f = data["fba"].setdefault(a, {"available": 0, "dos": None})
            f["available"] += it.get("available") or 0
            if it.get("dos") is not None:
                f["dos"] = it["dos"] if f["dos"] is None else min(f["dos"], it["dos"])
    if pl:
        data["sources"]["전체 상품(제목·가격)"] = pl.name
    if pf:
        data["sources"]["FBA 재고"] = pf.name
    bdir = ac.BASE_DIR / "reports" / "business"
    p = _latest(bdir, "business_data_*.xlsx")
    data["buybox"] = {}
    if p:
        import business
        for a, x in business.parse(inv.read_xlsx(p)[1]).items():
            data["buybox"][a] = x
        data["sources"]["비즈니스 보고서(바이박스·전체 매출)"] = p.name
    return data


# ================================================================ 집계 도구

def agg(rows: List[Dict], keys: Tuple[str, ...]) -> Dict[Tuple, Dict]:
    out: Dict[Tuple, Dict] = {}
    for r in rows:
        k = tuple((r.get(x) or "").strip() if isinstance(r.get(x), str) else r.get(x) for x in keys)
        a = out.setdefault(k, {"impressions": 0, "clicks": 0, "spend": 0.0, "sales": 0.0, "orders": 0, "bid": 0,
                               "budget": 0, "rows": 0, "tos_is": 0})
        for f in ("impressions", "clicks", "spend", "sales", "orders"):
            a[f] += r.get(f, 0) or 0
        if r.get("bid"):
            a["bid"] = r["bid"]
        if r.get("budget"):
            a["budget"] = r["budget"]
        if r.get("tos_is"):
            a["tos_is"] = r["tos_is"]
        a["rows"] += 1
        for x in ("targeting_type", "status", "ad_type", "portfolio"):
            if r.get(x) and not a.get(x):
                a[x] = r[x]
    for a in out.values():
        a["acos"] = a["spend"] / a["sales"] if a["sales"] else None
        a["ctr"] = a["clicks"] / a["impressions"] if a["impressions"] else 0
        a["cpc"] = a["spend"] / a["clicks"] if a["clicks"] else 0
        a["cvr"] = a["orders"] / a["clicks"] if a["clicks"] else 0
        a["aov"] = a["sales"] / a["orders"] if a["orders"] else 0
    return out


def window(rows: List[Dict], rng) -> List[Dict]:
    return [r for r in rows if ac.in_range(r, rng)]


def tokens(text: str) -> List[str]:
    return [w for w in re.findall(r"[a-z0-9]+", (text or "").lower()) if len(w) >= 3 and w not in STOP]


def norm_match(m: str) -> str:
    m = (m or "").lower()
    for k in ("exact", "phrase", "broad"):
        if k in m:
            return k
    if any(x in m for x in ("close", "loose", "substitute", "complement", "auto")):
        return "auto"
    if "asin" in m or "product" in m or "category" in m:
        return "product"
    return m


def is_auto_target(t: str) -> bool:
    return bool(re.search(r"close-match|loose-match|substitutes|complements|^\*$", (t or "").lower()))


def is_asin(t: str) -> bool:
    return bool(re.fullmatch(r"b0[a-z0-9]{8}", (t or "").strip().lower()))


# ================================================================ 분석

class Advisor:
    def __init__(self, data: Dict, today: date):
        self.d = data
        self.today = today
        self.rng = ac.ranges(today)
        self.actions: List[Dict] = []
        self.alerts: List[Dict] = []

    # ---------------------------------------------------------------- 준비
    def prepare(self) -> None:
        d = self.d
        # 보고서 기간이 규칙 기간보다 짧으면(예: 최근 7일만 받음) 받은 기간 전체로 판단
        dates = sorted({r["date"] for r in d["sp_targeting"] + d["sp_search_term"] + d["sp_campaign"] if r.get("date")})
        if dates and dates[0] > self.rng["rule14"][0].isoformat():
            first = datetime.fromisoformat(dates[0]).date()
            self.rng["rule14"] = (first, self.rng["yesterday"][1])
            self.window_name = f"최근 {(self.rng['yesterday'][1] - first).days + 1}일"
            d["notes"].append(f"입찰·부정 키워드 규칙은 받은 기간({first} ~ {self.rng['yesterday'][1]}) 전체로 판단했습니다. "
                              "규격서 기준은 ‘최근 3일 제외 14일’이며, 최근 2~3일은 전환이 아직 다 집계되지 않았을 수 있습니다.")
        else:
            self.window_name = "최근 14일(최근 3일 제외)"
        st = d["sp_search_term"] + d["sb_search_term"]
        tg = d["sp_targeting"]
        cp = d["sp_campaign"]
        ap = d["sp_advertised_product"]
        self.has_daily = bool(st or tg or cp)
        # 캠페인 → 광고 상품
        self.camp_asins: Dict[str, Dict[str, float]] = defaultdict(lambda: defaultdict(float))
        self.adgroup_asins: Dict[Tuple[str, str], set] = defaultdict(set)
        for r in ap:
            a = (r.get("asin") or "").strip()
            if a:
                self.camp_asins[r.get("campaign", "")][a] += r.get("spend", 0)
                self.adgroup_asins[(r.get("campaign", ""), r.get("ad_group", ""))].add(a)
        campaigns = sorted({r.get("campaign", "") for r in cp + st + tg + ap + d["cm_campaigns_7d"] + d["cm_campaigns_yesterday"]}
                           - {""})
        asins = {}
        for a, x in d["listings"].items():
            asins[a] = x
        for c in self.camp_asins.values():
            for a in c:
                asins.setdefault(a, {})
        self.conf = load_config(campaigns, {a: x for a, x in asins.items() if a in self._advertised() or not self._advertised()})
        self.cfg = self.conf["cfg"]
        self.default_target = float(self.cfg.get("기본 목표 ACoS(%)") or 30) / 100
        # 투입한 키워드(검색어가 이미 정확 타겟인지 / 새 단어가 투입됐는지 판단용)
        self.targeted = defaultdict(set)   # 단어 -> {match}
        self.exact_where = defaultdict(set)
        for r in tg:
            t = (r.get("targeting") or "").strip().lower()
            m = norm_match(r.get("match_type") or "")
            if t:
                self.targeted[t].add(m)
                if m == "exact":
                    self.exact_where[t].add(r.get("campaign", ""))
        # 캠페인 용도
        self.purpose = {}
        guessed = {}
        by_c = defaultdict(list)
        for r in tg:
            by_c[r.get("campaign", "")].append(r)
        types = {}
        for r in d["cm_campaigns_7d"] + d["cm_campaigns_yesterday"] + cp:
            if r.get("targeting_type"):
                types[r.get("campaign", "")] = r["targeting_type"]
        for c in campaigns:
            rows = by_c.get(c, [])
            ms = [norm_match(r.get("match_type") or "") for r in rows]
            ts = [(r.get("targeting") or "").lower() for r in rows]
            if "auto" in (types.get(c, "").lower()) or any(is_auto_target(t) for t in ts) or "auto" in ms:
                g = "RESEARCH"
            elif ms and sum(1 for m in ms if m == "product") > len(ms) / 2 or any(t.startswith("asin") for t in ts):
                g = "COMPETITOR"
            elif ts and self.conf["brand"] and sum(1 for t in ts if any(b in t for b in self.conf["brand"])) > len(ts) / 2:
                g = "DEFENSE"
            elif ms and sum(1 for m in ms if m == "exact") >= len(ms) / 2:
                g = "PERF"
            elif ms:
                g = "DISCOVERY"
            else:
                g = ""
            guessed[c] = g
            fixed = self.conf["purposes"].get(c, {}).get("fixed")
            self.purpose[c] = fixed or g
        save_config(self.conf, guessed)

    def _advertised(self) -> set:
        s = set()
        for c in self.camp_asins.values():
            s |= set(c)
        return s

    def target_acos(self, campaign: str) -> float:
        asins = self.camp_asins.get(campaign)
        if asins:
            top = max(asins.items(), key=lambda x: x[1])[0]
            c = self.conf["costs"].get(top)
            if c:
                return c["target"]
        return self.default_target

    def protected(self, term: str) -> bool:
        t = (term or "").lower()
        return any(b in t for b in self.conf["brand"]) or t in self.conf["protect"]

    def add(self, **kw) -> None:
        kw.setdefault("priority", 3)
        kw.setdefault("level", "L2")
        try:
            kw["howto"] = self.howto(kw)
        except Exception:
            kw["howto"] = ""
        self.actions.append(kw)

    # ---------------------------------------------------------------- 실행 방법(광고 콘솔 영어 메뉴 그대로)
    def _same_product(self, campaigns, purpose: str) -> List[str]:
        """같은 상품(ASIN)을 광고하는 캠페인 중 목적이 purpose인 것."""
        asins = set()
        for c in campaigns:
            asins |= set(self.camp_asins.get(c, {}))
        out = [c for c, p in self.purpose.items() if p == purpose and (not asins or asins & set(self.camp_asins.get(c, {})))]
        return sorted(out)[:3]

    def howto(self, a: Dict) -> str:
        t, c, g = a.get("type", ""), a.get("campaign", ""), a.get("ad_group", "")
        tgt, m = a.get("target", ""), a.get("match", "")
        cm = f"Campaign Manager → Campaigns → ‘{c}’" if c and not c.startswith("(") else "Campaign Manager → Campaigns → (해당 캠페인)"
        ag = f" → Ad groups → ‘{g}’" if g else " → Ad groups → (해당 광고 그룹)"
        mt = {"정확": "Exact", "구문": "Phrase", "확장": "Broad", "exact": "Exact", "phrase": "Phrase", "broad": "Broad"}.get(
            str(m).split(" ")[0], m)
        auto = {"close-match": "Close match", "loose-match": "Loose match", "substitutes": "Substitutes",
                "complements": "Complements"}.get(str(tgt).lower())
        row = f"자동 타겟 ‘{auto}’ 줄" if auto else f"‘{tgt}’({mt}) 줄"
        if t == "키워드 일시중지":
            if auto:
                return (f"{cm}{ag} → Targeting 탭 → {row} 왼쪽 스위치를 끔(Paused). "
                        "다른 자동 타겟(" + "·".join(x for x in ("Close match", "Loose match", "Substitutes", "Complements") if x != auto)
                        + ")은 그대로")
            return (f"{cm}{ag} → Targeting 탭 → {row} 왼쪽 스위치를 끔(Paused). "
                    f"광고 그룹 Negative targeting 탭 → Add negative keywords → Negative exact로 ‘{tgt}’ 추가도 검토")
        if t in ("입찰가 인하", "입찰가 인상"):
            return f"{cm}{ag} → Targeting 탭 → {row}의 Bid 칸 클릭 → {a.get('current', '')} → {a.get('suggested', '')} 입력 → Save"
        if t == "부정 키워드(정확)":
            return (f"{cm}{ag} → Negative targeting 탭 → Add negative keywords → Match type ‘Negative exact’ → ‘{tgt}’ 입력 → "
                    "Add keywords → Save. (지름길: 같은 광고 그룹 Search terms 탭 → 이 검색어 줄 → Add as negative)")
        if t == "부정 키워드 검토":
            return "브랜드·보호 키워드라 자동으로 제외하지 않음. 필요하다고 판단되면 ‘부정 키워드(정확)’과 같은 방법(Negative exact)"
        if t == "부정 키워드(구문)":
            return (f"먼저 각 캠페인 Search terms 탭에서 ‘{tgt}’가 들어간 검색어 중 주문 있는 것이 없는지 확인 → "
                    f"해당 캠페인 → Ad groups → Negative targeting → Add negative keywords → ‘Negative phrase’ → ‘{tgt}’ → Save")
        if t in ("키워드 추가(정확)", "상품 타겟팅 추가"):
            srcs = [x.strip() for x in c.split(",") if x.strip()]
            bid = a.get("bid") or "Suggested bid의 낮은 값"
            if t == "키워드 추가(정확)":
                dst = self._same_product(srcs, "PERF")
                dest = (f"정확 캠페인 ‘{dst[0]}’" if dst else
                        "같은 상품의 정확(Exact) 수동 캠페인(없으면 Create campaign → Sponsored Products → Manual targeting → "
                        "Keyword targeting으로 새로 만들기, 예: ‘상품코드 Exact (Harvest)’)")
                return (f"① {dest} → Ad groups → Targeting 탭 → Add keywords → Enter list → Match type ‘Exact’만 체크 → "
                        f"Bid {bid} → ‘{tgt}’ → Add keywords → Save  "
                        f"② 출처 캠페인 ‘{srcs[0] if srcs else ''}’ → Ad groups → Negative targeting → Add negative keywords → "
                        f"Negative exact → ‘{tgt}’ → Save  (지름길: 출처 캠페인 Search terms 탭 → Add as keyword / Add as negative)")
            dst = self._same_product(srcs, "COMPETITOR")
            dest = (f"상품 타겟팅 캠페인 ‘{dst[0]}’" if dst else
                    "상품 타겟팅 수동 캠페인(없으면 Create campaign → Sponsored Products → Manual targeting → Product targeting)")
            return (f"① {dest} → Ad groups → Targeting 탭 → Add products → Individual products → ‘{tgt}’ 검색·Add(Bid {bid}) → Save  "
                    f"② 출처 캠페인 ‘{srcs[0] if srcs else ''}’ → Ad groups → Negative targeting → Negative products → ‘{tgt}’ → Save")
        if t == "매치 유형 변경":
            nxt = "Phrase" if mt == "Broad" else "Exact"
            return (f"{cm}{ag} → Targeting 탭 → Add keywords → Enter list → Match type ‘{nxt}’ → ‘{tgt}’ 추가 → "
                    f"기존 ‘{tgt}’({mt}) 줄은 Bid를 낮추거나 Paused. 또는 Search terms 탭에서 주문 없는 검색어를 Add as negative")
        if t == "예산 증액":
            low = str(a.get("suggested", "")).split("~")[0].strip()
            return f"{cm} → Campaign settings 탭 → Daily budget {a.get('current', '')} → {low} → Save changes(Campaigns 목록의 Budget 칸을 눌러 바로 바꿔도 됨)"
        if t == "입찰 점검(예산 아님)":
            return (f"예산은 그대로. {cm} → Ad groups → Targeting 탭 → Spend 순으로 정렬 → 지출이 크고 Orders 0이거나 ACoS가 높은 키워드부터 "
                    "Bid 10~20% 낮추기(‘입찰 조정’ 시트의 같은 캠페인 항목 참고)")
        if t == "게재 위치 조정":
            return f"{cm} → " + str(a.get("suggested", "")).replace("Campaign settings", "Campaign settings 탭", 1)
        if t == "재고 연동":
            cs = [x for x, asins in self.camp_asins.items() if tgt in asins][:4]
            where = ", ".join(f"‘{x}’" for x in cs) or "(이 ASIN을 광고하는 캠페인)"
            return (f"캠페인 {where} → Ad groups → Ads 탭 → ASIN {tgt} 줄 스위치 끔(Paused), 재고 입고 후 다시 켬. "
                    "또는 Targeting 탭에서 Bid를 낮춰 판매 속도 조절")
        if t in ("리스팅 개선(클릭률)", "리스팅 개선(장바구니)"):
            return (f"Seller Central → Catalog → Manage All Inventory → ‘{tgt}’ 검색 → Edit에서 제목·이미지 확인(가격은 Price 칸). "
                    f"amazon.com/dp/{tgt} 와 같은 검색어의 경쟁 상품을 비교. 제목·불릿 수정은 이 프로그램의 ‘리스팅·가격 수정’으로도 가능")
        if t == "상품명 키워드 추가":
            return ("‘리스팅·가격 수정’ → ① 빈 요청서 → 항목 ‘제목’ 또는 ‘불릿1~5’에 단어를 넣은 새 값 → ② 검사 → ③ 적용. "
                    "또는 Seller Central → Catalog → Manage All Inventory → Edit → Product details / Description")
        if t == "신규 키워드 후보":
            mm = "Exact" if "정확" in str(a.get("suggested", "")) else "Phrase"
            return (f"같은 상품의 수동 키워드 캠페인 → Ad groups → Targeting 탭 → Add keywords → Enter list → Match type ‘{mm}’ → "
                    f"‘{tgt}’ → Bid는 Suggested bid의 낮은 값부터 → Add keywords → Save")
        if t == "노출 확대":
            return (f"이 검색어를 타겟팅하는 캠페인의 Targeting 탭 → ‘{tgt}’ 키워드 Bid +10~15%, "
                    "또는 Campaign settings → Adjust bids by placement → ‘Top of search (first page)’ +10~20% → Save changes")
        if t == "상품 타겟팅 후보":
            return (f"먼저 amazon.com/dp/{tgt} 에서 가격·리뷰가 우리 상품보다 불리하지 않은지 확인 → 상품 타겟팅 캠페인 → Ad groups → "
                    f"Targeting 탭 → Add products → Individual products → ‘{tgt}’ → Bid는 Suggested bid 낮은 값 → Save")
        return ""

    # ---------------------------------------------------------------- 1. 입찰(5절)
    def bids(self) -> None:
        tg = window(self.d["sp_targeting"], self.rng["rule14"])
        if not tg:
            return
        cap_pct = float(self.cfg.get("단일 조정 상한(%)") or 20) / 100
        floor, ceil_ = float(self.cfg.get("최소 입찰가($)") or 0.2), float(self.cfg.get("최대 입찰가($)") or 5)
        camp = agg(window(self.d["sp_campaign"] or self.d["sp_targeting"], self.rng["rule14"]), ("campaign",))
        min_click = int(self.cfg.get("일시중지 최소 클릭") or 30)
        neg_click = int(self.cfg.get("부정 키워드 최소 클릭") or 15)
        # 현재 입찰가: 가장 최근 날의 값
        cur_bid = {}
        for r in sorted(self.d["sp_targeting"], key=lambda r: r.get("date") or ""):
            if r.get("bid"):
                cur_bid[(r.get("campaign"), r.get("ad_group"), r.get("targeting"), r.get("match_type"))] = r["bid"]
        for (c, g, t, m), a in agg(tg, ("campaign", "ad_group", "targeting", "match_type")).items():
            tgt = self.target_acos(c)
            ca = camp.get((c,), {})
            aov = a["aov"] or ca.get("aov") or 0
            cvr = a["cvr"] if a["orders"] else (ca.get("cvr") or 0)
            cpa_target = aov * tgt if aov else None
            cpc_cap = aov * cvr * tgt if aov and cvr else None
            bid = cur_bid.get((c, g, t, m)) or a["cpc"]
            base = dict(campaign=c, ad_group=g, target=t, match=m, window=self.window_name,
                        metrics=f"클릭 {a['clicks']:.0f}, 주문 {a['orders']:.0f}, 지출 ${a['spend']:.2f}, "
                                f"ACoS {_pct(a['acos'])}, 목표 {tgt:.0%}", spend=a["spend"])
            if a["clicks"] >= min_click and a["orders"] == 0:
                self.add(type="키워드 일시중지", level="L1", priority=1, current=_money(bid), suggested="일시중지 + 부정 후보",
                         rule="입찰-1: 클릭 ≥ 30, 주문 0", **base)
            elif neg_click <= a["clicks"] < min_click and a["orders"] == 0 and cpa_target and a["spend"] >= cpa_target:
                nb = max(floor, bid * (1 - cap_pct))
                self.add(type="입찰가 인하", level="L1", priority=2, current=_money(bid), suggested=_money(nb),
                         rule="입찰-2: 클릭 15~29, 주문 0, 지출 ≥ 목표 CPA", **base)
            elif a["orders"] >= 2 and a["acos"] is not None and a["acos"] > tgt * 1.3:
                nb = max(floor, bid * (1 - cap_pct), min(cpc_cap or bid, bid))
                self.add(type="입찰가 인하", level="L1", priority=2, current=_money(bid), suggested=_money(nb),
                         rule="입찰-3: 주문 ≥ 2, ACoS > 목표 × 1.3 → 목표 CPC 상한(최대 -20%)", **base)
            elif a["orders"] >= 2 and a["acos"] is not None and a["acos"] < tgt * 0.8:
                nb = min(ceil_, bid * (1 + min(cap_pct, 0.15)))
                if cpc_cap:
                    nb = min(nb, max(cpc_cap, bid))
                if nb > bid + 0.005:
                    self.add(type="입찰가 인상", level="L1", priority=3, current=_money(bid), suggested=_money(nb),
                             rule="입찰-5: 주문 ≥ 2, ACoS < 목표 × 0.8 (+10~20%, 목표 CPC 상한 이내)", **base)
            elif a["impressions"] < 100 and bid:
                self.add(type="입찰가 인상", level="L1", priority=4, current=_money(bid), suggested=_money(min(ceil_, bid * 1.1)),
                         rule="입찰-6: 노출 < 100(14일), 데이터 확보용 +10%", **base)

    # ---------------------------------------------------------------- 2. 검색어: 부정 / 수확 / 내부 경쟁
    def search_terms(self) -> None:
        st_all = self.d["sp_search_term"] + self.d["sb_search_term"]
        st = window(st_all, self.rng["rule14"])
        if not st:
            return
        neg_click = int(self.cfg.get("부정 키워드 최소 클릭") or 15)
        stem_click = int(self.cfg.get("어근 부정 최소 클릭") or 40)
        min_orders = int(self.cfg.get("수확 최소 주문") or 2)
        # 같은 광고 그룹의 같은 검색어가 여러 타겟에서 올 수 있음 → 합친 뒤 판단
        by_t = defaultdict(set)
        for r in st:
            if r.get("search_term"):
                by_t[(r.get("campaign", ""), r.get("ad_group", ""), r["search_term"].strip(), r.get("ad_type", ""))].add(
                    (r.get("targeting") or "", r.get("match_type") or ""))
        rows = agg(st, ("campaign", "ad_group", "search_term", "ad_type"))
        zero_stem = defaultdict(lambda: {"clicks": 0, "spend": 0.0, "terms": set(), "orders": 0})
        harvest = defaultdict(lambda: {"clicks": 0, "spend": 0.0, "sales": 0.0, "orders": 0, "src": set(), "cpc": []})
        for (c, g, q, typ), a in rows.items():
            if not q:
                continue
            ql = q.lower()
            tgts = by_t.get((c, g, q, typ), set())
            exact_self = any(norm_match(m) == "exact" and t.lower() == ql for t, m in tgts)
            src_txt = ", ".join(sorted(f"{t} {m}".strip() for t, m in tgts))[:60]
            tgt = self.target_acos(c)
            base = dict(campaign=c, ad_group=g, target=q, match=f"{typ} · {src_txt}", window=self.window_name,
                        metrics=f"클릭 {a['clicks']:.0f}, 주문 {a['orders']:.0f}, 지출 ${a['spend']:.2f}, ACoS {_pct(a['acos'])}",
                        spend=a["spend"])
            for w in set(tokens(q)):
                z = zero_stem[w]
                z["clicks"] += a["clicks"]
                z["spend"] += a["spend"]
                z["orders"] += a["orders"]
                z["terms"].add(q)
            if exact_self:
                continue                                          # 정확 키워드 자체에서 온 트래픽
            # 부정(정확)
            if a["clicks"] >= neg_click and a["orders"] == 0:
                if self.protected(q):
                    self.add(type="부정 키워드 검토", level="L3", priority=3, current="", suggested="수동 검토(브랜드/보호 키워드)",
                             rule="부정-보호: 브랜드·보호 키워드는 자동 부정 금지", **base)
                else:
                    self.add(type="부정 키워드(정확)", level="L1", priority=1, current="", suggested=f"‘{q}’ 정확 부정",
                             rule="부정-1: 클릭 ≥ 15, 주문 0", **base)
                continue
            # 내부 경쟁: 다른 캠페인에서 이미 정확 타겟
            if ql in self.exact_where and c not in self.exact_where[ql] and a["clicks"] >= 3:
                self.add(type="부정 키워드(정확)", level="L1", priority=2, current="",
                         suggested=f"‘{q}’ 정확 부정(정확 캠페인 {', '.join(sorted(self.exact_where[ql]))[:60]}과 내부 경쟁)",
                         rule="부정-3: 정확 캠페인에 이미 있는 검색어", **base)
                continue
            if a["orders"] > 0 or a["clicks"] > 0:
                h = harvest[(ql, typ)]
                for f_ in ("clicks", "spend", "sales", "orders"):
                    h[f_] += a[f_]
                h["src"].add(c)
                h["term"] = q
                h["tgt"] = min(h.get("tgt", 1), tgt)
        # 수확: 검색어별로 모든 출처 캠페인을 합침
        for (ql, typ), h in harvest.items():
            q = h["term"]
            acos = h["spend"] / h["sales"] if h["sales"] else None
            if h["orders"] < min_orders or acos is None or acos > h["tgt"]:
                continue
            if "exact" in self.targeted.get(ql, set()) or "product" in self.targeted.get(ql, set()):
                continue
            cpc = h["spend"] / h["clicks"] if h["clicks"] else 0
            aov = h["sales"] / h["orders"]
            cvr = h["orders"] / h["clicks"] if h["clicks"] else 0
            sb = min(x for x in (cpc * 1.1, aov * cvr * h["tgt"]) if x) if (cpc or cvr) else None
            srcs = ", ".join(sorted(h["src"]))
            base = dict(campaign=srcs[:120], ad_group="", target=q, match=f"{typ} · 출처 캠페인 {len(h['src'])}개",
                        window=self.window_name, spend=h["spend"],
                        metrics=f"클릭 {h['clicks']:.0f}, 주문 {h['orders']:.0f}, 지출 ${h['spend']:.2f}, ACoS {_pct(acos)}")
            if is_asin(q):
                self.add(type="상품 타겟팅 추가", level="L1", priority=2, current="", bid=_money(sb) if sb else "",
                         suggested=f"ASIN {q.upper()} 상품 타겟팅(COMPETITOR 캠페인)" + (f", 입찰가 {_money(sb)}" if sb else "") +
                                   " + 출처 캠페인에 ASIN 부정", rule="수확: 주문 ≥ 2, ACoS ≤ 목표", **base)
            else:
                self.add(type="키워드 추가(정확)", level="L1", priority=2, current="", bid=_money(sb) if sb else "",
                         suggested=f"‘{q}’ 정확 일치로 PERF 캠페인에 추가" + (f", 입찰가 {_money(sb)}" if sb else "") +
                                   " + 출처 캠페인에 정확 부정", rule="수확: 주문 ≥ 2, ACoS ≤ 목표 → 정확 캠페인으로 이동", **base)
        for w, z in zero_stem.items():
            if z["orders"] == 0 and z["clicks"] >= stem_click and len(z["terms"]) >= 2 and not self.protected(w):
                self.add(type="부정 키워드(구문)", level="L2", priority=2, campaign="(여러 캠페인)", ad_group="", target=w,
                         match="어근", window=self.window_name, current="",
                         suggested=f"‘{w}’ 구문 부정(영향 범위가 커서 승인 필요)",
                         metrics=f"이 단어가 들어간 검색어 {len(z['terms'])}개, 클릭 {z['clicks']:.0f}, 주문 0, 지출 ${z['spend']:.2f}",
                         rule="부정-2: 같은 어근 합계 클릭 ≥ 40, 주문 0", spend=z["spend"])
        self.st_rows = rows

    # ---------------------------------------------------------------- 3. 매치 유형
    def match_types(self) -> None:
        st = window(self.d["sp_search_term"], self.rng["rule14"])
        if not st:
            return
        by_kw = defaultdict(lambda: {"spend": 0.0, "waste": 0.0, "orders": 0, "terms": 0, "sales": 0.0})
        for r in st:
            m = norm_match(r.get("match_type") or "")
            if m not in ("broad", "phrase"):
                continue
            k = (r.get("campaign", ""), r.get("ad_group", ""), r.get("targeting", ""), m)
            b = by_kw[k]
            b["spend"] += r.get("spend", 0)
            b["sales"] += r.get("sales", 0)
            b["orders"] += r.get("orders", 0)
            if not r.get("orders"):
                b["waste"] += r.get("spend", 0)
        for (c, g, t, m), b in by_kw.items():
            if b["spend"] >= 10 and b["waste"] / b["spend"] >= 0.6:
                nxt = "구문" if m == "broad" else "정확"
                self.add(type="매치 유형 변경", level="L2", priority=3, campaign=c, ad_group=g, target=t, match=MATCH_KO.get(m, m),
                         current=MATCH_KO.get(m, m), suggested=f"{nxt} 일치로 좁히거나, 주문 없는 검색어를 부정 처리",
                         metrics=f"지출 ${b['spend']:.2f} 중 주문 없는 검색어 지출 {b['waste'] / b['spend']:.0%}, 주문 {b['orders']:.0f}",
                         window=self.window_name, rule="매치: 확장·구문 키워드의 낭비 지출 비율 ≥ 60%", spend=b["spend"])

    # ---------------------------------------------------------------- 4. 예산 / 지출 이상 / 게재 위치
    def budgets(self) -> None:
        cp = self.d["sp_campaign"]
        y = self.rng["yesterday"]
        if cp:
            yd = agg(window(cp, y), ("campaign",))
            d7 = agg(window(cp, self.rng["7d"]), ("campaign",))
            prev = window(cp, (y[0] - timedelta(days=7), y[0] - timedelta(days=1)))
            same_wd = agg([r for r in prev if r.get("date") and datetime.fromisoformat(r["date"]).weekday() == y[0].weekday()],
                          ("campaign",))
            prev7 = agg(prev, ("campaign",))
        else:
            yd = agg(self.d["cm_campaigns_yesterday"], ("campaign",))
            d7 = agg(self.d["cm_campaigns_7d"], ("campaign",))
            same_wd, prev7 = {}, {}
        budgets = {}
        for r in self.d["cm_campaigns_7d"] + self.d["cm_campaigns_yesterday"] + cp:
            if r.get("budget"):
                budgets[r.get("campaign", "")] = r["budget"]
        for (c,), a in yd.items():
            tgt = self.target_acos(c)
            w = d7.get((c,), a)
            bud = budgets.get(c) or a.get("budget")
            base = dict(campaign=c, ad_group="", target="", match="", spend=a["spend"])
            if bud and a["spend"] >= bud * 0.95:
                if w["acos"] is not None and w["acos"] <= tgt:
                    self.add(type="예산 증액", level="L2", priority=2, current=_money(bud), suggested=f"{_money(bud * 1.2)} ~ {_money(bud * 1.5)}",
                             metrics=f"전날 지출 ${a['spend']:.2f} / 예산 ${bud:.2f}, 7일 ACoS {_pct(w['acos'])} (목표 {tgt:.0%})",
                             window="전날 + 최근 7일", rule="예산-1: 예산 소진, ACoS ≤ 목표", **base)
                elif w["acos"] is None or w["acos"] > tgt * 1.3:
                    self.add(type="입찰 점검(예산 아님)", level="알림", priority=3, current=_money(bud), suggested="예산을 늘리지 말고 입찰가부터 낮추기",
                             metrics=f"전날 지출 ${a['spend']:.2f} / 예산 ${bud:.2f}, 7일 ACoS {_pct(w['acos'])}",
                             window="전날 + 최근 7일", rule="예산-2: 예산 소진이지만 ACoS > 목표 × 1.3", **base)
            base_spend = (same_wd.get((c,)) or {}).get("spend") or ((prev7.get((c,)) or {}).get("spend", 0) / 7 if prev7 else 0)
            if base_spend and a["spend"] > base_spend * 1.5 and a["spend"] >= 5:
                lvl = "P0" if a["spend"] > base_spend * 2 else "P1"
                self.alerts.append({"level": lvl, "object": c, "what": "지출 급증",
                                    "detail": f"전날 ${a['spend']:.2f}, 기준(지난주 같은 요일 또는 일평균) ${base_spend:.2f}",
                                    "action": "L1 입찰 인상 보류, 검색어 확인"})
        # 게재 위치(보고서 이름 → 캠페인 설정 Adjust bids by placement의 이름)
        def setting_name(p: str) -> str:
            low = p.lower()
            if "top" in low:
                return "Top of search (first page)"
            if "detail" in low or "product page" in low:
                return "Product pages"
            if "other" in low or "rest of search" in low:
                return "Rest of search"
            return p
        path = "Campaign settings → Campaign bidding strategy → Adjust bids by placement"
        pl = [r for r in window(cp, self.rng["rule14"]) if r.get("placement")]
        if pl:
            camp = agg(pl, ("campaign",))
            for (c, p), a in agg(pl, ("campaign", "placement")).items():
                tot = camp.get((c,))
                if not tot or tot["acos"] is None or a["orders"] < 2 or a["acos"] is None:
                    continue
                if "top" in p.lower() and a["acos"] < tot["acos"] * 0.8:
                    self.add(type="게재 위치 조정", level="L2", priority=3, campaign=c, ad_group="",
                             target=f"{setting_name(p)} (보고서: {p})", match="",
                             current="", suggested=f"{path} → ‘Top of search (first page)’ 비율 +10~20%p 올리기 → Save changes",
                             metrics=f"Top of search ACoS {_pct(a['acos'])} vs 캠페인 {_pct(tot['acos'])}, 주문 {a['orders']:.0f}",
                             window=self.window_name, rule="게재위치: 상단 ACoS가 전체보다 20% 이상 낮음", spend=a["spend"])
                elif a["acos"] > tot["acos"] * 1.3 and a["spend"] >= 10:
                    sn = setting_name(p)
                    self.add(type="게재 위치 조정", level="L2", priority=3, campaign=c, ad_group="",
                             target=f"{sn} (보고서: {p})", match="",
                             current="", suggested=(f"{path} → ‘{sn}’ 비율 낮추기(지금 값의 절반, 20% 이하면 0%) → Save changes. "
                                                    "이미 0%면 기본 입찰을 10~20% 낮추고 ACoS가 좋은 위치에 %를 더해 노출 유지"),
                             metrics=f"{sn} ACoS {_pct(a['acos'])} vs 캠페인 {_pct(tot['acos'])}, 지출 ${a['spend']:.2f}",
                             window=self.window_name, rule="게재위치: 위치 ACoS가 전체보다 30% 이상 높음", spend=a["spend"])

    # ---------------------------------------------------------------- 5. 재고 / Buy Box 연동
    def inventory(self) -> None:
        warn = float(self.cfg.get("재고 경고 일수") or 14)
        spend_by_asin = defaultdict(float)
        for r in window(self.d["sp_advertised_product"], self.rng["7d"]):
            if r.get("asin"):
                spend_by_asin[r["asin"]] += r.get("spend", 0)
        for a, sp in spend_by_asin.items():
            if sp <= 0:
                continue
            f = self.d["fba"].get(a)
            title = (self.d["listings"].get(a) or {}).get("title", "")
            if f and f["available"] <= 0:
                self.alerts.append({"level": "P0", "object": f"{a} {title[:30]}", "what": "품절인데 광고 지출 중",
                                    "detail": f"최근 7일 광고 지출 ${sp:.2f}, FBA 판매 가능 0", "action": "해당 ASIN 광고 일시중지(L1)"})
            elif f and f["dos"] is not None and f["dos"] < warn:
                self.add(type="재고 연동", level="L2", priority=2, campaign="(이 ASIN을 광고하는 캠페인)", ad_group="", target=a, match="",
                         current=f"판매 가능 {f['available']:.0f}개, 약 {f['dos']:.0f}일분", suggested="입찰가를 낮추거나 광고 일시중지 검토",
                         metrics=f"최근 7일 광고 지출 ${sp:.2f}", window="최근 7일", rule=f"재고: 판매 가능 일수 < {warn:.0f}일", spend=sp)
            bb = self.d["buybox"].get(a)
            if bb and bb.get("page_views") and bb.get("buybox", 100) < 90:
                self.alerts.append({"level": "P1", "object": f"{a} {title[:30]}", "what": "바이박스 낮음",
                                    "detail": f"바이박스 {bb['buybox']:.0f}%, 최근 7일 광고 지출 ${sp:.2f}",
                                    "action": "가격·재고·합승 판매자 확인 후 광고 유지 여부 결정"})

    # ---------------------------------------------------------------- 6. Listing(제목 / 메인 이미지 / 가격)
    def listings(self) -> None:
        scp = self.d["scp"]
        if scp:
            ctrs = sorted(x["clicks"] / x["impressions"] for x in scp if x["impressions"] >= 500)
            carts = sorted(x["cart_adds"] / x["clicks"] for x in scp if x["clicks"] >= 30)
            med_ctr = ctrs[len(ctrs) // 2] if ctrs else 0
            med_cart = carts[len(carts) // 2] if carts else 0
            for x in scp:
                title = x.get("title") or (self.d["listings"].get(x["asin"]) or {}).get("title", "")
                if x["impressions"] >= 1000 and med_ctr and x["clicks"] / x["impressions"] < med_ctr * 0.6:
                    self.add(type="리스팅 개선(클릭률)", level="L3", priority=3, campaign="", ad_group="", target=x["asin"], match="",
                             current=f"CTR {x['clicks'] / x['impressions']:.2%} (브랜드 중앙값 {med_ctr:.2%})",
                             suggested="메인 이미지·제목 앞부분·가격·리뷰 수 점검(검색 결과에서 클릭이 적음)",
                             metrics=f"노출 {x['impressions']:,.0f}, 클릭 {x['clicks']:,.0f} · {title[:40]}",
                             window="Brand Analytics 최근 1주", rule="리스팅-1: 검색 CTR < 브랜드 중앙값 × 0.6", spend=0)
                if x["clicks"] >= 50 and med_cart and x["cart_adds"] / x["clicks"] < med_cart * 0.6:
                    self.add(type="리스팅 개선(장바구니)", level="L3", priority=3, campaign="", ad_group="", target=x["asin"], match="",
                             current=f"장바구니율 {x['cart_adds'] / x['clicks']:.1%} (중앙값 {med_cart:.1%})",
                             suggested="가격·쿠폰·상세 이미지·A+·리뷰 점검(상세 페이지에서 이탈)",
                             metrics=f"클릭 {x['clicks']:,.0f}, 장바구니 {x['cart_adds']:,.0f} · {title[:40]}",
                             window="Brand Analytics 최근 1주", rule="리스팅-2: 장바구니율 < 중앙값 × 0.6", spend=0)
        # 전환이 잘 되는 검색어 중 제목에 없는 단어
        st = window(self.d["sp_search_term"], self.rng["30d"])
        if st and self.d["listings"]:
            words_by_asin = defaultdict(lambda: defaultdict(float))
            for r in st:
                if not r.get("orders") or is_asin(r.get("search_term", "")):
                    continue
                asins = self.adgroup_asins.get((r.get("campaign", ""), r.get("ad_group", ""))) or set(self.camp_asins.get(r.get("campaign", ""), {}))
                if len(asins) != 1:
                    continue
                a = next(iter(asins))
                for w in set(tokens(r["search_term"])):
                    words_by_asin[a][w] += r["orders"]
            for a, words in words_by_asin.items():
                title = (self.d["listings"].get(a) or {}).get("title", "")
                if not title:
                    continue
                have = set(tokens(title)) | {w.rstrip("s") for w in tokens(title)}
                missing = [w for w, o in sorted(words.items(), key=lambda x: -x[1]) if o >= 2 and w not in have and w.rstrip("s") not in have
                           and not self.protected(w)][:6]
                if missing:
                    self.add(type="상품명 키워드 추가", level="L3", priority=3, campaign="", ad_group="", target=a, match="",
                             current=title[:80], suggested="제목·불릿에 추가 검토: " + ", ".join(missing),
                             metrics="주문이 난 검색어에 자주 나오지만 제목에 없는 단어", window="최근 30일", rule="리스팅-3: 전환 검색어 단어 ⊄ 제목",
                             spend=0)

    # ---------------------------------------------------------------- 7. Brand Analytics: 키워드 기회
    def opportunities(self) -> None:
        sqp = sorted(self.d["sqp"], key=lambda x: -x["volume"])
        for x in sqp[:200]:
            q = x["query"].lower()
            tgt_matches = self.targeted.get(q, set())
            imp_share, clk_share, pur_share = x["imp_share"] / 100, x["clk_share"] / 100, x["pur_share"] / 100   # 보고서 값은 백분율
            x["imp_share_r"], x["clk_share_r"], x["pur_share_r"] = imp_share, clk_share, pur_share
            base = dict(campaign="", ad_group="", target=x["query"], window="Brand Analytics 최근 1주", spend=0)
            note = f"검색량 {x['volume']:,.0f}, 노출 점유 {imp_share:.1%}, 클릭 점유 {clk_share:.1%}, 구매 점유 {pur_share:.1%}"
            if not tgt_matches and x["volume"] >= 300 and not is_asin(q):
                n = len(q.split())
                self.add(type="신규 키워드 후보", level="L2", priority=3, match="", current="미투입",
                         suggested=f"{'정확' if n >= 3 else '구문'} 일치로 추가 검토", metrics=note,
                         rule="기회-1: 검색량 높은 검색어인데 키워드로 투입되지 않음", **base)
            elif tgt_matches and imp_share and clk_share > imp_share * 1.2 and imp_share < 0.05 and x["volume"] >= 300:
                self.add(type="노출 확대", level="L2", priority=3, match=",".join(sorted(tgt_matches)), current=f"노출 점유 {imp_share:.1%}",
                         suggested="노출되면 클릭이 잘 되는 검색어 → 입찰가·검색 상단 조정 상향 검토", metrics=note,
                         rule="기회-2: 클릭 점유 > 노출 점유 × 1.2, 노출 점유 < 5%", **base)
            elif imp_share >= 0.05 and clk_share < imp_share * 0.6:
                self.add(type="리스팅 개선(클릭률)", level="L3", priority=3, match="", current=f"노출 점유 {imp_share:.1%}, 클릭 점유 {clk_share:.1%}",
                         suggested="이 검색어에서 노출은 되지만 클릭이 적음 → 메인 이미지·제목·가격 경쟁력 점검", metrics=note,
                         rule="기회-3: 클릭 점유 < 노출 점유 × 0.6", **base)
        # Top Search Terms에서 클릭 상위 경쟁 ASIN
        own = set(self.d["listings"]) | self._advertised()
        seen = set()
        for r in self.d["tst"]:
            for k, v in r.items():
                m = re.search(r"\bB0[A-Z0-9]{8}\b", str(v))
                if m and m.group(0) not in own and m.group(0) not in seen and "asin" in k.lower():
                    seen.add(m.group(0))
                    self.add(type="상품 타겟팅 후보", level="L2", priority=4, campaign="", ad_group="", target=m.group(0), match="",
                             current="", suggested="경쟁 ASIN 상품 타겟팅(COMPETITOR 캠페인) 검토",
                             metrics=f"‘{r.get('조회한 키워드', '')}’ 검색 시 클릭 상위 상품", window="Brand Analytics 최근 1주",
                             rule="기회-4: 인기 검색어의 클릭 상위 경쟁 상품", spend=0)

    # ---------------------------------------------------------------- 정리
    def run(self) -> Dict:
        self.prepare()
        for f in (self.bids, self.search_terms, self.match_types, self.budgets, self.inventory, self.listings, self.opportunities):
            try:
                f()
            except Exception as e:
                self.d["notes"].append(f"{f.__name__} 분석 중 오류: {type(e).__name__}: {e}")
        self.actions.sort(key=lambda x: (x.get("priority", 9), -(x.get("spend") or 0)))
        return {"actions": self.actions, "alerts": self.alerts}


def _pct(v) -> str:
    if v is None:
        return "—"
    if v == float("inf"):
        return "∞(매출 없음)"
    return f"{v:.0%}"


def _money(v) -> str:
    return f"${v:.2f}" if v else "—"


# ================================================================ KPI

def kpis(data: Dict, today: date) -> Dict[str, Dict]:
    rng = ac.ranges(today)
    cp = data["sp_campaign"]
    out = {}
    if cp:
        for k in ("yesterday", "7d", "prev7d"):
            out[k] = ac.kpi(window(cp, rng[k]))
        out["source"] = "SP 캠페인 보고서(일별)"
    elif data["cm_campaigns_7d"] or data["cm_campaigns_yesterday"]:
        out["yesterday"] = ac.kpi(data["cm_campaigns_yesterday"])
        out["7d"] = ac.kpi(data["cm_campaigns_7d"])
        out["prev7d"] = None
        out["source"] = "광고 관리자 캠페인 표(내보내기, SP·SB 등 전체)"
    return out

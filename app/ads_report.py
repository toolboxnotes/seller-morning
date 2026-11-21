"""
광고 자료 + 필요한 조치 → Excel 보고서(한국어)
"""

from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List

import ads_analyze as aa
import ads_common as ac
import inventory as inv

TYPE_GROUPS = {
    "입찰 조정": ("입찰가 인하", "입찰가 인상", "키워드 일시중지"),
    "부정 키워드": ("부정 키워드(정확)", "부정 키워드(구문)", "부정 키워드 검토"),
    "키워드 추가": ("키워드 추가(정확)", "상품 타겟팅 추가", "신규 키워드 후보", "상품 타겟팅 후보", "노출 확대"),
    "매치·예산·위치": ("매치 유형 변경", "예산 증액", "입찰 점검(예산 아님)", "게재 위치 조정"),
    "재고·리스팅": ("재고 연동", "리스팅 개선(클릭률)", "리스팅 개선(장바구니)", "상품명 키워드 추가"),
}
LEVEL_NOTE = {"L1": "소폭·되돌릴 수 있음(나중에 자동 실행 후보)", "L2": "영향이 커서 승인 필요", "L3": "사람이 판단",
              "알림": "확인만"}


def write(out: Path, today: date, data: Dict, res: Dict, k: Dict, pending: List[str]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FONT = inv.FONT
    NAVY, GREY, RED, GREEN, ORANGE = "1F3864", "6B6B6B", "C01C1C", "1E7B34", "C26A00"
    HEAD = PatternFill("solid", fgColor="DCE6F1")
    thin = Side(style="thin", color="BFBFBF")
    BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
    rng = ac.ranges(today)
    f = lambda size=10.5, bold=False, color=None: Font(name=FONT, size=size, bold=bold, color=color)

    def cell(ws, r, c, v, font=None, num=None, fill=None, border=BOX, wrap=False, align=None):
        x = ws.cell(row=r, column=c, value=v)
        x.font = font or f(10)
        if num:
            x.number_format = num
        if fill:
            x.fill = fill
        if border:
            x.border = border
        x.alignment = Alignment(vertical="center", wrap_text=wrap, horizontal=align)
        return x

    def table(ws, headers, rows, widths, start=1, nums=None, wrap_cols=()):
        for c, h in enumerate(headers, start=1):
            cell(ws, start, c, h, f(10, True), fill=HEAD, wrap=True, align="center")
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
        ws.row_dimensions[start].height = 30
        for i, row in enumerate(rows, start=start + 1):
            for c, v in enumerate(row, start=1):
                cell(ws, i, c, v, num=(nums or {}).get(c), wrap=c in wrap_cols)
        if rows:
            ws.auto_filter.ref = f"A{start}:{get_column_letter(len(headers))}{start + len(rows)}"
        ws.freeze_panes = ws.cell(row=start + 1, column=1)

    wb = Workbook()
    acts, alerts = res["actions"], res["alerts"]

    # ================================================= 요약
    ws = wb.active
    ws.title = "요약"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", (24, 16, 16, 16, 14, 30)):
        ws.column_dimensions[col].width = w
    cell(ws, 1, 1, "광고 데이터 · 조치 제안", f(18, True, NAVY), border=None)
    cell(ws, 2, 1, f"전날 {rng['yesterday'][0]} · 최근 7일 {rng['7d'][0]} ~ {rng['7d'][1]} (미국 태평양 시간)   "
                   f"생성: {datetime.now():%Y-%m-%d %H:%M}   ※ 1단계: 수집·제안만, 광고는 수정하지 않음", f(9.5, color=GREY), border=None)
    r = 4
    cell(ws, r, 1, "핵심 지표" + (f" ({k.get('source')})" if k.get("source") else ""), f(13, True, NAVY), border=None)
    r += 1
    heads = ["지표", "전날", "최근 7일", "이전 7일", "7일 변화"]
    for c, h in enumerate(heads, start=1):
        cell(ws, r, c, h, f(10, True), fill=HEAD, align="center")
    rows = [("노출", "impressions", "#,##0"), ("클릭", "clicks", "#,##0"), ("광고비", "spend", '"$"#,##0.00'),
            ("광고 매출", "sales", '"$"#,##0.00'), ("주문", "orders", "#,##0"), ("ACoS", "acos", "0.0%"),
            ("ROAS", "roas", "0.00"), ("CTR", "ctr", "0.00%"), ("CPC", "cpc", '"$"#,##0.00'), ("전환율(CVR)", "cvr", "0.0%")]
    if k.get("7d"):
        for label, key, fmt in rows:
            r += 1
            cell(ws, r, 1, label, f(10, True))
            vals = []
            for p in ("yesterday", "7d", "prev7d"):
                v = (k.get(p) or {}).get(key) if k.get(p) else None
                v = None if v == float("inf") else v
                vals.append(v)
                cell(ws, r, len(vals) + 1, v if v is not None else "—", num=fmt, align=None if v is not None else "center")
            a, b = vals[1], vals[2]
            cell(ws, r, 5, (a / b - 1) if (a is not None and b) else "—", num="+0.0%;[Red]-0.0%;0.0%")
    else:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        cell(ws, r, 1, "캠페인 데이터를 아직 받지 못했습니다(보고서 생성 대기 중).", f(10, color=ORANGE), border=None)

    r += 2
    cell(ws, r, 1, "긴급 알림", f(13, True, NAVY), border=None)
    if alerts:
        for al in sorted(alerts, key=lambda x: x["level"]):
            r += 1
            color = RED if al["level"] == "P0" else ORANGE
            cell(ws, r, 1, al["level"], f(10, True, color))
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            cell(ws, r, 2, f"{al['what']} · {al['object']} — {al['detail']} → {al['action']}", wrap=True)
            ws.row_dimensions[r].height = 30
    else:
        r += 1
        cell(ws, r, 1, "없음 ✓", f(10, color=GREEN), border=None)

    r += 2
    cell(ws, r, 1, "조치 제안 요약", f(13, True, NAVY), border=None)
    cnt = Counter(a["type"] for a in acts)
    lv = Counter(a["level"] for a in acts)
    r += 1
    for c, h in enumerate(["분류", "제안 수", "시트"], start=1):
        cell(ws, r, c, h, f(10, True), fill=HEAD, align="center")
    for g, types in TYPE_GROUPS.items():
        r += 1
        cell(ws, r, 1, g, f(10, True))
        cell(ws, r, 2, sum(cnt.get(t, 0) for t in types), num="#,##0")
        cell(ws, r, 3, g)
    r += 1
    cell(ws, r, 1, "합계", f(10, True))
    cell(ws, r, 2, len(acts), f(10, True))
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    cell(ws, r, 1, "등급: " + ", ".join(f"{k_} {v}건({LEVEL_NOTE.get(k_, '')})" for k_, v in sorted(lv.items())),
         f(9, color=GREY), border=None, wrap=True)
    ws.row_dimensions[r].height = 30

    r += 2
    cell(ws, r, 1, "데이터 출처", f(13, True, NAVY), border=None)
    for kk, v in data["sources"].items():
        r += 1
        cell(ws, r, 1, kk, f(9.5, True), border=None)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        cell(ws, r, 2, v, f(9.5, color=GREY), border=None)
    notes = list(data["notes"])
    if pending:
        notes.insert(0, "아직 생성 중인 광고 보고서: " + ", ".join(pending) +
                     ". 나중에 ‘광고 데이터 수집·조치 제안’을 다시 실행하면 남은 보고서만 받아 키워드·검색어 분석이 추가됩니다.")
    notes.append("목표 ACoS는 ads_config.xlsx의 ‘ASIN 원가’(원가를 넣으면 손익분기 ACoS로 계산)와 ‘설정’의 기본값을 사용합니다.")
    notes.append("입찰·부정·수확 규칙은 최근 3일을 제외한 14일 데이터로 판단합니다(아마존 전환 데이터 지연 때문).")
    for n in notes:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        cell(ws, r, 1, "· " + n, f(9, color=GREY), border=None, wrap=True)
        ws.row_dimensions[r].height = 30

    # ================================================= 조치 제안(전체) + 분류
    headers = ["우선순위", "등급", "조치 유형", "캠페인", "광고 그룹", "대상(키워드·검색어·ASIN)", "매치/구분", "현재", "제안",
               "실행 방법(광고 콘솔 메뉴)", "근거 지표", "기간", "규칙", "광고비($)"]
    widths = [8, 7, 18, 34, 22, 32, 18, 18, 42, 70, 46, 18, 40, 10]

    def act_rows(items):
        return [[a.get("priority"), a.get("level"), a.get("type"), a.get("campaign", ""), a.get("ad_group", ""),
                 a.get("target", ""), a.get("match", ""), a.get("current", ""), a.get("suggested", ""), a.get("howto", ""),
                 a.get("metrics", ""), a.get("window", ""), a.get("rule", ""), round(a.get("spend") or 0, 2)] for a in items]

    s = wb.create_sheet("조치 제안")
    table(s, headers, act_rows(acts), widths, nums={14: '"$"#,##0.00'}, wrap_cols=(9, 10, 11))
    for g, types in TYPE_GROUPS.items():
        s = wb.create_sheet(g.replace("·", "·")[:31])
        table(s, headers, act_rows([a for a in acts if a["type"] in types]), widths, nums={14: '"$"#,##0.00'}, wrap_cols=(9, 10, 11))

    # ================================================= 성과표
    def perf_sheet(name, rows, keys, key_names, extra=None):
        s = wb.create_sheet(name)
        hd = key_names + ["노출", "클릭", "CTR", "광고비", "CPC", "주문", "매출", "ACoS", "전환율"] + (extra or [])
        out = []
        for kk, a in sorted(rows.items(), key=lambda x: -x[1]["spend"]):
            out.append(list(kk) + [a["impressions"], a["clicks"], a["ctr"], round(a["spend"], 2), round(a["cpc"], 2), a["orders"],
                                   round(a["sales"], 2), a["acos"] if a["acos"] is not None else "—", a["cvr"]])
        n = len(key_names)
        nums = {n + 1: "#,##0", n + 2: "#,##0", n + 3: "0.00%", n + 4: '"$"#,##0.00', n + 5: '"$"#,##0.00', n + 6: "#,##0",
                n + 7: '"$"#,##0.00', n + 8: "0.0%", n + 9: "0.0%"}
        table(s, hd, out, [34, 22, 30, 12][:n] + [10] * 9, nums=nums)
        return s

    cp = data["sp_campaign"] or []
    if cp:
        perf_sheet("캠페인(최근 7일)", aa.agg(aa.window(cp, rng["7d"]), ("campaign",)), ("campaign",), ["캠페인"])
        perf_sheet("캠페인(전날)", aa.agg(aa.window(cp, rng["yesterday"]), ("campaign",)), ("campaign",), ["캠페인"])
    else:
        if data["cm_campaigns_7d"]:
            perf_sheet("캠페인(최근 7일)", aa.agg(data["cm_campaigns_7d"], ("campaign",)), ("campaign",), ["캠페인"])
        if data["cm_campaigns_yesterday"]:
            perf_sheet("캠페인(전날)", aa.agg(data["cm_campaigns_yesterday"], ("campaign",)), ("campaign",), ["캠페인"])
    tg = data["sp_targeting"]
    if tg:
        perf_sheet("키워드(최근 7일)", aa.agg(aa.window(tg, rng["7d"]), ("campaign", "ad_group", "targeting", "match_type")),
                   (), ["캠페인", "광고 그룹", "키워드·타겟", "매치"])
    st = data["sp_search_term"] + data["sb_search_term"]
    if st:
        perf_sheet("검색어(최근 7일)", aa.agg(aa.window(st, rng["7d"]), ("campaign", "targeting", "search_term")),
                   (), ["캠페인", "키워드·타겟", "검색어"])
        perf_sheet("검색어(전날)", aa.agg(aa.window(st, rng["yesterday"]), ("campaign", "targeting", "search_term")),
                   (), ["캠페인", "키워드·타겟", "검색어"])

    # ================================================= Brand Analytics
    if data["sqp"]:
        s = wb.create_sheet("BA 검색어 성과(SQP)")
        hd = ["검색어", "순위 점수", "검색량", "노출(전체)", "노출(브랜드)", "노출 점유", "클릭(전체)", "클릭(브랜드)", "클릭 점유",
              "장바구니(브랜드)", "장바구니 점유", "구매(전체)", "구매(브랜드)", "구매 점유", "광고 투입 여부"]
        targeted = set()
        for r_ in data["sp_targeting"]:
            targeted.add((r_.get("targeting") or "").lower())
        share = lambda v: v / 100   # 보고서 값은 백분율
        rows = [[x["query"], x["score"], x["volume"], x["imp_total"], x["imp_brand"], share(x["imp_share"]), x["clk_total"],
                 x["clk_brand"], share(x["clk_share"]), x["cart_brand"], share(x["cart_share"]), x["pur_total"], x["pur_brand"],
                 share(x["pur_share"]), "투입" if x["query"].lower() in targeted else ("미확인" if not targeted else "미투입")]
                for x in sorted(data["sqp"], key=lambda x: -x["volume"])]
        table(s, hd, rows, [34, 9, 10, 11, 11, 9, 11, 11, 9, 11, 10, 10, 10, 9, 12],
              nums={3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "0.00%", 7: "#,##0", 8: "#,##0", 9: "0.00%", 10: "#,##0",
                    11: "0.00%", 12: "#,##0", 13: "#,##0", 14: "0.00%"})
    if data["scp"]:
        s = wb.create_sheet("BA 카탈로그 성과(SCP)")
        hd = ["ASIN", "상품명", "노출", "클릭", "CTR", "장바구니", "구매", "매출", "전환율", "가격(중앙값)"]
        rows = [[x["asin"], x["title"] or (data["listings"].get(x["asin"]) or {}).get("title", ""), x["impressions"], x["clicks"],
                 (x["clicks"] / x["impressions"]) if x["impressions"] else 0, x["cart_adds"], x["purchases"], x["sales"],
                 (x["purchases"] / x["clicks"]) if x["clicks"] else 0, x["price"]]
                for x in sorted(data["scp"], key=lambda x: -x["impressions"])]
        table(s, hd, rows, [13, 48, 11, 10, 8, 10, 8, 11, 8, 10],
              nums={3: "#,##0", 4: "#,##0", 5: "0.00%", 6: "#,##0", 7: "#,##0", 8: '"$"#,##0.00', 9: "0.0%", 10: '"$"#,##0.00'})
    if data["tst"]:
        s = wb.create_sheet("인기 검색어(TST)")
        hd = list(data["tst"][0].keys())
        table(s, hd, [[r_.get(h, "") for h in hd] for r_ in data["tst"]], [22] * len(hd))

    for sh in wb.worksheets:
        sh.page_setup.orientation = "portrait" if sh.title == "요약" else "landscape"
        sh.sheet_properties.pageSetUpPr.fitToPage = True
        sh.page_setup.fitToWidth, sh.page_setup.fitToHeight = 1, 0
    wb.save(out)

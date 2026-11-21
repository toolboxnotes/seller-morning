"""
2단계: 리스팅·가격 수정(SP-API) — 사람이 확인한 것만 적용

흐름
  1. 요청서(reports/changes/변경_요청서.xlsx)에 바꿀 내용을 적음: SKU · 항목 · 새 값 · 적용(Y)
     ‘요청서 만들기’를 누르면 Listing 점검의 ‘제목 초안’과 현재 가격표가 미리 들어감(적용 칸은 비워 둠)
  2. 검사: 현재 값을 API로 읽고, Amazon의 VALIDATION_PREVIEW(검사만, 실제 변경 없음)로 오류를 확인
  3. 화면에서 적용할 줄을 골라 ‘적용’ → 그때만 실제로 바뀜
  4. 모든 변경은 reports/changes/변경_기록.xlsx 와 change_log.json 에 원래 값과 함께 기록 → ‘되돌리기’ 가능

안전 장치
  - 가격: 0 이하·현재 가격의 ±30%를 넘는 변경은 ‘큰 변경’으로 표시, 따로 허용해야 적용
  - 제목 200자 초과는 오류, 75자 초과는 주의(카테고리 권장)
  - 한 번에 최대 50건, 요청 사이 간격을 둠(초당 5건 제한)
"""

import copy
import json
import time
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import sp_api

BASE_DIR = Path(__file__).resolve().parent.parent
DIR = BASE_DIR / "reports" / "changes"
REQ_FILE = DIR / "변경_요청서.xlsx"
LOG_JSON = DIR / "change_log.json"
LOG_XLSX = DIR / "변경_기록.xlsx"
MAX_BATCH = 50
PRICE_LIMIT = 0.30
TITLE_MAX, TITLE_WARN = 200, 75

# 요청서 ‘항목’ → (속성 이름, 설명)
FIELDS = {"제목": "item_name", "가격": "purchasable_offer", "불릿1": "bullet_point", "불릿2": "bullet_point",
          "불릿3": "bullet_point", "불릿4": "bullet_point", "불릿5": "bullet_point", "검색어": "generic_keyword",
          "설명": "product_description"}
FIELDS["하이라이트"] = "*highlight"        # Item Highlights — 속성 이름은 상품 유형 정의에서 찾음(제목 75자 미만일 때만 보임)
HIGHLIGHT_MAX = 125
_HL_CACHE: Dict[str, str] = {}


def _highlight_attr(c, item: Dict, product_type: str) -> str:
    """Item Highlights 속성 이름: ① 현재 리스팅 속성에 있으면 그 이름 ② 상품 유형 정의(Product Type Definitions)에서 찾기."""
    for k in (item.get("attributes") or {}):
        if "highlight" in k.lower():
            return k
    if product_type in _HL_CACHE:
        return _HL_CACHE[product_type]
    name = ""
    try:
        d = c.get(f"/definitions/2020-09-01/productTypes/{product_type}", marketplaceIds=c.mp, requirements="LISTING",
                  locale="en_US")
        url = (d.get("schema") or {}).get("link", {}).get("resource", "")
        if url:
            import urllib.request
            with urllib.request.urlopen(url, timeout=60) as r:
                schema = json.loads(r.read().decode("utf-8"))
            props = schema.get("properties") or {}
            cands = [k for k in props if "highlight" in k.lower()]
            if not cands:
                cands = [k for k, v in props.items() if "highlight" in str(v.get("title", "")).lower()]
            name = cands[0] if cands else ""
    except Exception:
        name = ""
    _HL_CACHE[product_type] = name
    return name

REQ_HEAD = ["적용(Y)", "SKU", "항목", "새 값", "메모(자유)", "참고: 현재 값(만들 때)"]


# ================================================================ 요청서

def _latest_listing_report() -> Optional[Path]:
    d = BASE_DIR / "reports" / "listing"
    files = sorted(d.glob("listing_check_*.xlsx")) if d.exists() else []
    return files[-1] if files else None


def title_drafts() -> List[Dict]:
    """최근 Listing 점검 결과의 ‘제목 초안’ 시트 → [{SKU, 현재 제목, 초안}]"""
    p = _latest_listing_report()
    if not p:
        return []
    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    if "제목 초안" not in wb.sheetnames:
        return []
    rows = list(wb["제목 초안"].iter_rows(values_only=True))
    out, head = [], None
    for r in rows:
        vals = [("" if v is None else str(v)) for v in r]
        if head is None:
            if "SKU" in vals and any(v.startswith("제목 초안") for v in vals):
                head = vals
            continue
        d = dict(zip(head, vals))
        if d.get("SKU"):
            out.append({"SKU": d["SKU"], "현재": d.get("현재 제목", ""),
                        "초안": next((d[k] for k in d if k.startswith("제목 초안")), ""),
                        "하이라이트": next((d[k] for k in d if k.startswith("Item Highlights 초안")), "")})
    return out


def make_request(prices: Optional[List[Dict]] = None, drafts: bool = False) -> Path:
    """요청서 만들기(기본: 빈 양식). 이미 있으면 백업하고 새로 만듦.
    drafts=True: Listing 점검의 제목·Item Highlights 초안을 미리 채움. prices: 현재 가격표 시트(참고용)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    DIR.mkdir(parents=True, exist_ok=True)
    if REQ_FILE.exists():
        REQ_FILE.rename(DIR / f"변경_요청서_백업_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "변경 요청"
    ws.append(REQ_HEAD)
    head_fill = PatternFill("solid", fgColor="232F3E")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
    for d in (title_drafts() if drafts else []):
        ws.append(["", d["SKU"], "제목", d["초안"], "Listing 점검 제목 초안(검토 후 Y)", d["현재"]])
        if d.get("하이라이트"):
            ws.append(["", d["SKU"], "하이라이트", d["하이라이트"], "Item Highlights 초안(제목을 75자 미만으로 바꿀 때 같이)", ""])
    for col, w in zip("ABCDEF", (9, 22, 10, 80, 30, 80)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        row[3].alignment = Alignment(wrap_text=True, vertical="top")
        row[5].alignment = Alignment(wrap_text=True, vertical="top")
    dv = DataValidation(type="list", formula1='"' + ",".join(FIELDS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C2:C2000")
    dv2 = DataValidation(type="list", formula1='"Y"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add("A2:A2000")
    ws.freeze_panes = "A2"
    g = wb.create_sheet("안내")
    for line in ("적용할 줄의 ‘적용(Y)’ 칸에 Y를 넣고 저장 → 화면 프로그램의 ‘리스팅·가격 수정’ 창에서 ‘검사’ → 확인 후 ‘적용’.",
                 "항목: " + ", ".join(FIELDS) + "   (불릿1~5는 해당 번호의 불릿만 바꿈, 하이라이트 = Item Highlights 125자·제목 75자 미만일 때만 보임)",
                 "가격: 숫자만(예: 19.99, 미국 달러). 현재 가격의 ±30%를 넘으면 ‘큰 변경 허용’을 체크해야 적용됩니다.",
                 "제목: 200자 이하(75자 권장). 금지 문자 ! $ ? _ { } ^ ¬ ¦ 는 쓰지 마세요.",
                 "검사는 Amazon이 오류만 확인하고 실제로 바꾸지 않습니다. ‘적용’을 눌러야 바뀝니다.",
                 "모든 변경은 ‘변경_기록.xlsx’에 원래 값과 함께 남고, 화면에서 되돌릴 수 있습니다.",
                 "Amazon 반영까지 보통 몇 분~최대 몇 시간 걸립니다. 다른 판매자/카탈로그 기여로 제목이 바뀌지 않을 수도 있습니다."):
        g.append([line])
    g.column_dimensions["A"].width = 130
    if prices:
        p = wb.create_sheet("현재 가격표")
        p.append(["SKU", "ASIN", "상품명", "현재 가격(USD)"])
        for r in prices:
            p.append([r.get("SKU"), r.get("ASIN"), r.get("상품명"), r.get("가격")])
        for col, w in zip("ABCD", (22, 14, 80, 14)):
            p.column_dimensions[col].width = w
    wb.save(REQ_FILE)
    return REQ_FILE


def current_prices() -> List[Dict]:
    """현재 가격표: 오늘 받은 전체 상품 보고서가 있으면 그것, 없으면 SP-API로 받음."""
    rows: List[Dict] = []
    try:
        import inventory as inv
        from datetime import date
        today = date.today().isoformat()
        x = inv.REPORT_DIR / f"all_listings_{today}.xlsx"
        if x.exists():
            import ads_common as ac
            rows = ac.load_rows(x)
        else:
            import sp_data
            raw = sp_data.all_listings(today)
            if raw:
                _, rows = inv.read_report(raw)
    except Exception as e:
        print(f"   ⚠ 현재 가격표를 만들지 못했습니다({type(e).__name__}: {str(e)[:120]}) — 요청서는 그대로 만듭니다")
        return []

    def g(r, *names):
        low = {str(k).lower().strip(): v for k, v in r.items()}
        for n in names:
            if low.get(n) not in (None, ""):
                return low[n]
        return ""
    out = []
    for r in rows:
        sku = g(r, "seller-sku", "sku")
        if sku:
            out.append({"SKU": sku, "ASIN": g(r, "asin1", "asin"), "상품명": str(g(r, "item-name", "상품명", "title"))[:120],
                        "가격": g(r, "price", "가격")})
    return out


def load_requests() -> List[Dict]:
    """요청서에서 적용(Y)인 줄."""
    if not REQ_FILE.exists():
        return []
    from openpyxl import load_workbook
    ws = load_workbook(REQ_FILE, data_only=True)["변경 요청"]
    out = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not r or str(r[0] or "").strip().upper() not in ("Y", "YES", "O"):
            continue
        sku, field, new = str(r[1] or "").strip(), str(r[2] or "").strip(), r[3]
        if sku and field:
            out.append({"row": i, "sku": sku, "field": field, "new": "" if new is None else str(new).strip(),
                        "memo": str(r[4] or "")})
    return out


# ================================================================ 현재 값 · 패치 만들기

def _get(c, seller: str, sku: str) -> Dict:
    return c.get(f"/listings/2021-08-01/items/{seller}/{urllib.parse.quote(sku, safe='')}", marketplaceIds=c.mp,
                 includedData="summaries,attributes,offers,issues")


def _text_values(attr: List[Dict]) -> List[str]:
    return [str(x.get("value", "")) for x in attr or []]


def _price_of(item: Dict) -> Optional[float]:
    po = (item.get("attributes") or {}).get("purchasable_offer") or []
    for o in po:
        for op in o.get("our_price") or []:
            for s in op.get("schedule") or []:
                if s.get("value_with_tax") is not None:
                    return float(s["value_with_tax"])
    for o in item.get("offers") or []:
        if o.get("offerType") == "B2C" and o.get("price", {}).get("amount") is not None:
            return float(o["price"]["amount"])
    return None


def build(c, req: Dict, item: Dict) -> Dict:
    """요청 1건 → {old, new, patch, warn, error}"""
    attrs = item.get("attributes") or {}
    field = req["field"]
    out = {"old": "", "new": req["new"], "patch": None, "warn": "", "error": "", "attr": FIELDS.get(field, "")}
    if field not in FIELDS:
        out["error"] = f"알 수 없는 항목 ‘{field}’(사용 가능: {', '.join(FIELDS)})"
        return out
    name = FIELDS[field]
    if name.startswith("*"):
        pt = ((item.get("summaries") or [{}])[0]).get("productType", "")
        name = _highlight_attr(c, item, pt)
        if not name:
            out["error"] = f"이 상품 유형({pt})에는 Item Highlights 항목이 없습니다"
            return out
        out["attr"] = name
    # 변형(variation): 부모 SKU는 판매 상품이 아님(가격 없음). 제목 등은 바꿀 수 있지만 고객에게는 보통 자식 제목이 보임
    level = ((_text_values(attrs.get("parentage_level", [])) or [""])[0] or "").lower()
    out["level"] = level
    if level == "parent":
        if field == "가격":
            out["error"] = "부모(Parent) SKU는 판매 상품이 아니라 가격이 없습니다 — 자식 SKU별로 바꾸세요"
            return out
        out["warn"] = "부모 SKU(검색 결과·상세 페이지에는 보통 자식 SKU 제목이 보임)"
    out["old_attr"] = copy.deepcopy(attrs.get(name))          # 되돌리기용 원래 값(속성 전체)
    tmpl = {"language_tag": "en_US", "marketplace_id": c.mp}
    if field == "가격":
        try:
            new = round(float(str(req["new"]).replace("$", "").replace(",", "")), 2)
        except ValueError:
            out["error"] = "가격은 숫자로 적어 주세요(예: 19.99)"
            return out
        old = _price_of(item)
        out["old"] = "" if old is None else f"{old:.2f}"
        out["new"] = f"{new:.2f}"
        if new <= 0:
            out["error"] = "가격은 0보다 커야 합니다"
            return out
        if old and abs(new - old) / old > PRICE_LIMIT:
            out["warn"] = f"큰 변경 {((new - old) / old):+.0%}"
            out["big"] = True
        po = copy.deepcopy(attrs.get("purchasable_offer")) or [{"marketplace_id": c.mp, "currency": "USD", "audience": "ALL"}]
        tgt = next((o for o in po if o.get("audience", "ALL") == "ALL"), po[0])
        tgt["our_price"] = [{"schedule": [{"value_with_tax": new}]}]
        out["patch"] = {"op": "replace", "path": "/attributes/purchasable_offer", "value": po}
        mn, mx = tgt.get("minimum_seller_allowed_price"), tgt.get("maximum_seller_allowed_price")
        try:
            lo = float(mn[0]["schedule"][0]["value_with_tax"]) if mn else None
            hi = float(mx[0]["schedule"][0]["value_with_tax"]) if mx else None
            if (lo and new < lo) or (hi and new > hi):
                out["warn"] = (out["warn"] + " · " if out["warn"] else "") + f"최소/최대 허용 가격({lo}~{hi}) 밖"
        except Exception:
            pass
        return out
    new = str(req["new"])
    cur = attrs.get(name) or []
    if field.startswith("불릿"):
        idx = int(field[-1]) - 1
        vals = copy.deepcopy(cur)
        out["old"] = _text_values(cur)[idx] if idx < len(cur) else ""
        while len(vals) <= idx:
            vals.append(dict(tmpl, value=""))
        vals[idx] = dict(vals[idx] if isinstance(vals[idx], dict) else tmpl, value=new)
        vals = [v for v in vals if v.get("value")]
        if len(new) > 500:
            out["warn"] = " · ".join(x for x in (out["warn"], f"불릿 {len(new)}자(500자 이하 권장)") if x)
    else:
        out["old"] = " | ".join(_text_values(cur))
        base = dict(cur[0]) if cur else dict(tmpl)
        base.pop("value", None)
        vals = [dict(base, value=new)]
        if field == "제목":
            bad = [ch for ch in "!$?_{}^¬¦" if ch in new]
            if len(new) > TITLE_MAX:
                out["error"] = f"제목 {len(new)}자(200자 초과)"
                return out
            if bad:
                out["warn"] = " · ".join(x for x in (out["warn"], "금지 문자 " + " ".join(bad)) if x)
            elif len(new) > TITLE_WARN:
                out["warn"] = " · ".join(x for x in (out["warn"], f"{len(new)}자(75자 권장 초과)") if x)
        if field == "하이라이트":
            if len(new) > HIGHLIGHT_MAX:
                out["error"] = f"Item Highlights {len(new)}자({HIGHLIGHT_MAX}자 초과)"
                return out
            title = " ".join(_text_values(attrs.get("item_name", [])))
            if len(title) >= 75:
                out["warn"] = " · ".join(x for x in (out["warn"], f"현재 제목 {len(title)}자 — 제목이 75자 미만이어야 고객에게 보임") if x)
        if field == "검색어" and len(new.encode("utf-8")) > 250:
            out["warn"] = " · ".join(x for x in (out["warn"], f"검색어 {len(new.encode('utf-8'))}바이트(250바이트 넘으면 색인 안 됨)") if x)
    if not new:
        out["error"] = "새 값이 비어 있습니다"
        return out
    out["patch"] = {"op": "replace", "path": f"/attributes/{name}", "value": vals}
    return out


# ================================================================ 가격 변경: 수수료·이익 미리 보기(Product Fees API)

_COST_CACHE: Optional[Dict[str, float]] = None


def unit_costs() -> Dict[str, float]:
    """광고 설정(ads_config.xlsx ‘ASIN 원가’)의 원가+물류·관세+기타 → {ASIN·SKU: 1개당 비용}. 없으면 빈 값."""
    global _COST_CACHE
    if _COST_CACHE is not None:
        return _COST_CACHE
    out: Dict[str, float] = {}
    try:
        import ads_common as ac
        from openpyxl import load_workbook
        wb = load_workbook(ac.CONFIG_FILE, read_only=True, data_only=True)
        if "ASIN 원가" in wb.sheetnames:
            for r in wb["ASIN 원가"].iter_rows(min_row=2, values_only=True):
                if not r or not r[0] or r[4] in (None, ""):
                    continue
                cost = sum(ac.num(v) for v in (r[4], r[5], r[8]) if v not in (None, ""))
                out[str(r[0]).strip()] = cost
                if r[1]:
                    out[str(r[1]).strip()] = cost
    except Exception:
        pass
    _COST_CACHE = out
    return out


def fee_estimate(c, sku: str, price: float) -> Optional[Dict]:
    """이 가격으로 팔 때 Amazon 수수료(판매 수수료 + FBA 배송비 등). 실패하면 None."""
    body = {"FeesEstimateRequest": {"MarketplaceId": c.mp, "IsAmazonFulfilled": True, "Identifier": f"{sku}-{price}",
                                    "PriceToEstimateFees": {"ListingPrice": {"CurrencyCode": "USD", "Amount": round(price, 2)},
                                                            "Shipping": {"CurrencyCode": "USD", "Amount": 0}}}}
    try:
        d = c.request("POST", f"/products/fees/v0/listings/{urllib.parse.quote(sku, safe='')}/feesEstimate", None, body)
    except Exception as e:
        return {"error": str(e)[:150], "status": getattr(e, "status", 0)}
    res = (d.get("payload") or {}).get("FeesEstimateResult") or {}
    est = res.get("FeesEstimate") or {}
    if res.get("Status") not in (None, "Success") or not est:
        return {"error": str((res.get("Error") or {}).get("Message") or res.get("Status"))[:150]}
    det = {x.get("FeeType", ""): float((x.get("FinalFee") or x.get("FeeAmount") or {}).get("Amount") or 0)
           for x in est.get("FeeDetailList") or []}
    total = float((est.get("TotalFeesEstimate") or {}).get("Amount") or sum(det.values()))
    return {"total": total, "referral": det.get("ReferralFee", 0.0), "fba": det.get("FBAFees", 0.0),
            "other": max(0.0, total - det.get("ReferralFee", 0.0) - det.get("FBAFees", 0.0))}


def profit_text(c, sku: str, asin: str, old: Optional[float], new: float) -> Tuple[str, bool]:
    """→ (‘수수료·이익’ 설명, 손해 여부)."""
    fe = fee_estimate(c, sku, new)
    if not fe or fe.get("error"):
        why = (fe or {}).get("error", "")
        if (fe or {}).get("status") == 403:
            why = "수수료 조회 권한 없음(Pricing 역할 필요)"
        return f"수수료 계산 못 함({why[:80]})", False
    keep = new - fe["total"]
    cost = unit_costs().get(sku) or unit_costs().get(asin)
    txt = f"수수료 ${fe['total']:.2f}(판매 ${fe['referral']:.2f}·FBA ${fe['fba']:.2f}) → 남는 금액 ${keep:.2f}/개"
    loss = False
    if cost is not None:
        prof = keep - cost
        loss = prof < 0
        txt += f" · 원가 ${cost:.2f} → 이익 ${prof:.2f}({prof / new:.0%})"
        if old:
            fo = fee_estimate(c, sku, old)
            if fo and not fo.get("error"):
                po = old - fo["total"] - cost
                txt += f" · 지금 가격 이익 ${po:.2f}에서 {prof - po:+.2f}"
    else:
        txt += " · 원가 미입력(reports/ads/ads_config.xlsx ‘ASIN 원가’에 넣으면 이익까지 계산)"
    return txt, loss


def _patch(c, seller: str, sku: str, product_type: str, patches: List[Dict], preview: bool) -> Dict:
    params = {"marketplaceIds": c.mp, "includedData": "issues", "issueLocale": "en_US"}
    if preview:
        params["mode"] = "VALIDATION_PREVIEW"
    return c.request("PATCH", f"/listings/2021-08-01/items/{seller}/{urllib.parse.quote(sku, safe='')}", params,
                     {"productType": product_type, "patches": patches})


def _issues(r: Dict) -> Tuple[List[str], List[str]]:
    err = [f"{i.get('code')}: {i.get('message', '')[:160]}" for i in r.get("issues", []) if i.get("severity") == "ERROR"]
    wrn = [f"{i.get('code')}: {i.get('message', '')[:160]}" for i in r.get("issues", []) if i.get("severity") == "WARNING"]
    return err, wrn


# ================================================================ 검사 · 적용 · 되돌리기

def check(reqs: List[Dict], log=print) -> List[Dict]:
    """요청 → 현재 값 읽기 + 검사(실제 변경 없음). 결과 줄에 ok(적용 가능) 표시."""
    c = sp_api.client()
    seller = sp_api.seller_id()
    items: Dict[str, Dict] = {}
    out = []
    for i, req in enumerate(reqs[:MAX_BATCH * 4], 1):
        res = dict(req, status="", ok=False, product_type="", old="", attr=FIELDS.get(req["field"], ""))
        try:
            if req["sku"] not in items:
                items[req["sku"]] = _get(c, seller, req["sku"])
                time.sleep(0.25)
            item = items[req["sku"]]
            prev = [x for x in out if x["sku"] == req["sku"] and x.get("ok") and (x.get("attr") == FIELDS.get(req["field"]) or (req["field"] == "하이라이트" and x["field"] == "하이라이트"))]
            res["product_type"] = ((item.get("summaries") or [{}])[0]).get("productType", "")
            res["asin"] = ((item.get("summaries") or [{}])[0]).get("asin", "")
            res.update(build(c, req, item))
            if res["error"]:
                res["status"] = "오류: " + res["error"]
            elif not res["product_type"]:
                res["status"] = "오류: 상품 유형(productType)을 알 수 없음"
            else:
                r = _patch(c, seller, req["sku"], res["product_type"], [res["patch"]], preview=True)
                err, wrn = _issues(r)
                if err or r.get("status") not in ("VALID", "ACCEPTED"):
                    res["status"] = "Amazon 검사 오류: " + ("; ".join(err) or str(r.get("status")))
                else:
                    res["ok"] = True
                    res["status"] = "적용 가능" + (" · 주의: " + res["warn"] if res.get("warn") else "") + \
                                    (" · Amazon 주의: " + "; ".join(wrn[:2]) if wrn else "") + \
                                    (f" · 같은 SKU의 앞 {len(prev)}줄 변경을 포함한 값" if prev else "")
                    if req["field"] == "가격":
                        try:
                            res["profit"], res["loss"] = profit_text(c, req["sku"], res.get("asin", ""),
                                                                     float(res["old"]) if res.get("old") else None, float(res["new"]))
                        except Exception as ex:
                            res["profit"] = f"수수료 계산 못 함({str(ex)[:60]})"
                        if res.get("loss"):
                            res["big"] = True
                            res["status"] += " · ⚠ 이 가격이면 손해(원가+수수료보다 낮음)"
                    # 같은 SKU·같은 속성(예: 불릿1과 불릿2)을 여러 줄 바꿀 때 앞 줄 변경이 덮어쓰이지 않게 이어서 반영
                    item.setdefault("attributes", {})[res["attr"]] = copy.deepcopy(res["patch"]["value"])
                time.sleep(0.25)
        except Exception as e:
            st = getattr(e, "status", 0)
            res["status"] = ("오류: SKU를 찾지 못함" if st == 404 else
                             "오류: 수정 권한 없음(Product Listing 역할)" if st == 403 else f"오류: {str(e)[:200]}")
        log(f"   [{i}/{len(reqs)}] {req['sku']} · {req['field']}: {res['status']}")
        out.append(res)
    return out


def _log_load() -> List[Dict]:
    try:
        return json.loads(LOG_JSON.read_text(encoding="utf-8"))
    except Exception:
        return []


STATUS_ORDER = ("반영 안 됨", "처리 중", "접수됨", "제출값 저장됨", "반영됨", "되돌림")
PENDING_HOURS = 24          # 이 시간이 지나도 반영되지 않으면 ‘반영 안 됨’


def _fmt(ts: str) -> str:
    return (ts or "").replace("T", " ")[:16]


def _log_save(entries: List[Dict]) -> None:
    DIR.mkdir(parents=True, exist_ok=True)
    LOG_JSON.write_text(json.dumps(entries, ensure_ascii=False, indent=1), encoding="utf-8")
    try:
        import inventory as inv
        head = ["상태", "확인 기준", "적용 시각", "제출값 저장 확인", "반영 확인 시각", "마지막 확인", "작업", "SKU", "ASIN", "항목", "원래 값", "새 값",
                "Amazon 메시지", "Amazon 마지막 업데이트", "API 응답", "메모", "ID"]
        rows = [{"상태": e.get("status", "접수됨"), "확인 기준": e.get("verified_by", ""), "적용 시각": _fmt(e["time"]),
                 "제출값 저장 확인": _fmt(e.get("submitted_at", "")), "반영 확인 시각": _fmt(e.get("confirmed_at", "")),
                 "마지막 확인": _fmt(e.get("checked_at", "")), "작업": e.get("kind", "적용"), "SKU": e["sku"],
                 "ASIN": e.get("asin", ""), "항목": e["field"], "원래 값": e.get("old", ""), "새 값": e.get("new", ""),
                 "Amazon 메시지": e.get("issues", ""), "Amazon 마지막 업데이트": _fmt(e.get("amazon_updated", "")),
                 "API 응답": e.get("result", ""), "메모": e.get("memo", ""), "ID": e["id"]} for e in reversed(entries)]
        inv.save_report_xlsx(head, rows, LOG_XLSX, "변경 기록")
    except Exception:
        pass


def _current_value(item: Dict, field: str, attr: str = "") -> str:
    attrs = item.get("attributes") or {}
    if field == "가격":
        p = _price_of(item)
        return "" if p is None else f"{p:.2f}"
    vals = _text_values(attrs.get(attr if attr and not attr.startswith("*") else FIELDS.get(field, ""), []))
    if field.startswith("불릿"):
        i = int(field[-1]) - 1
        return vals[i] if i < len(vals) else ""
    return " | ".join(vals)


def _same(field: str, cur: str, target: str) -> bool:
    if field == "가격":
        try:
            return abs(float(cur) - float(target)) < 0.005
        except ValueError:
            return False
    norm = lambda x: " ".join(str(x or "").split())
    return norm(cur) == norm(target)


LIVE_HOURS = 48            # 내 리스팅에 저장된 뒤 이 시간이 지나도 카탈로그가 다르면 ‘반영 안 됨’


def _offer_price(item: Dict) -> Optional[float]:
    for o in item.get("offers") or []:
        if o.get("offerType") == "B2C" and o.get("price", {}).get("amount") is not None:
            return float(o["price"]["amount"])
    return None


def _live_value(c, item: Dict, e: Dict, cat_cache: Dict) -> Tuple[Optional[str], str]:
    """고객에게 보이는 값. 가격 → 판매 중인 오퍼 가격, 제목·불릿·설명·하이라이트 → 카탈로그(Catalog Items API).
    확인할 수 없으면 (None, 이유)."""
    if e["field"] == "가격":
        p = _offer_price(item)
        return (None, "판매 가격(오퍼)을 확인할 수 없음 — 제출값 기준") if p is None else (f"{p:.2f}", "판매 가격(오퍼)")
    if e["field"] == "검색어":
        return None, "검색어는 고객에게 보이지 않아 카탈로그로 확인할 수 없음 — 제출값 기준"
    asin = e.get("asin") or ((item.get("summaries") or [{}])[0]).get("asin", "")
    if not asin:
        return None, "ASIN을 몰라 카탈로그로 확인할 수 없음 — 제출값 기준"
    if asin not in cat_cache:
        try:
            cat_cache[asin] = c.get(f"/catalog/2022-04-01/items/{asin}", marketplaceIds=c.mp, includedData="attributes,summaries")
        except Exception as ex:
            cat_cache[asin] = {"_error": str(ex)[:120]}
        time.sleep(0.5)
    cat = cat_cache[asin]
    if cat.get("_error"):
        return None, f"카탈로그 조회 실패({cat['_error']}) — 제출값 기준"
    attrs = cat.get("attributes") or {}
    name = e.get("attr") or FIELDS.get(e["field"], "")
    if name not in attrs:
        if e["field"] == "제목":
            t = ((cat.get("summaries") or [{}])[0]).get("itemName")
            if t:
                return t, "카탈로그(상세 페이지)"
        return None, f"카탈로그에 이 항목({name})이 없어 확인할 수 없음 — 제출값 기준"
    vals = _text_values(attrs[name])
    if e["field"].startswith("불릿"):
        i = int(e["field"][-1]) - 1
        return (vals[i] if i < len(vals) else ""), "카탈로그(상세 페이지)"
    return " | ".join(vals), "카탈로그(상세 페이지)"


def verify(ids: Optional[List[str]] = None, log=print, days: int = 30) -> Dict[str, int]:
    """적용·되돌리기 기록의 반영 여부 확인: 현재 값을 다시 읽어 새 값과 비교.
    ids 없으면 최근 days일 중 아직 ‘반영됨’이 아닌 것만. → 상태별 건수"""
    c = sp_api.client()
    seller = sp_api.seller_id()
    entries = _log_load()
    now = datetime.now()
    # 같은 SKU·항목은 가장 마지막 기록만 의미가 있음(앞의 것은 이후 변경으로 덮임)
    latest = {}
    for e in entries:
        latest[(e["sku"], e["field"])] = e["id"]
    todo = []
    for e in entries:
        if e.get("kind") == "적용" and e.get("reverted"):
            e["status"] = "되돌림"
            continue
        if ids is not None:
            if e["id"] in ids:
                todo.append(e)
            continue
        age = (now - datetime.fromisoformat(e["time"])).days
        # ‘반영됨’이라도 예전 방식(내 제출값만 비교)으로 확인한 것은 카탈로그로 다시 확인
        if (e.get("status") != "반영됨" or not e.get("verified_by")) and age <= days:
            todo.append(e)
    cache: Dict[str, Dict] = {}
    cat_cache: Dict[str, Dict] = {}
    for e in todo:
        if latest.get((e["sku"], e["field"])) != e["id"]:
            e["status"], e["issues"] = "이후 변경으로 대체", ""
            e["checked_at"] = now.isoformat(timespec="seconds")
            continue
        try:
            if e["sku"] not in cache:
                cache[e["sku"]] = _get(c, seller, e["sku"])
                time.sleep(0.25)
            item = cache[e["sku"]]
            cur = _current_value(item, e["field"], e.get("attr", ""))
            summ = (item.get("summaries") or [{}])[0]
            e["amazon_updated"] = str(summ.get("lastUpdatedDate", ""))[:19]
            errs = [f"{i.get('code')}: {i.get('message', '')[:160]}" for i in item.get("issues", [])
                    if i.get("severity") == "ERROR"]
            hours = (now - datetime.fromisoformat(e["time"])).total_seconds() / 3600
            if _same(e["field"], cur, e["new"]):
                # 1단계 통과: 내 리스팅(제출값)은 바뀜 → 2단계: 고객에게 보이는 값(카탈로그·판매 가격) 확인
                e.setdefault("submitted_at", now.isoformat(timespec="seconds"))
                live, how = _live_value(c, item, e, cat_cache)
                if live is None:                   # 카탈로그로 확인할 수 없는 항목(검색어 등)
                    e["status"], e["verified_by"] = "반영됨", how
                    e["issues"] = how
                    e.setdefault("confirmed_at", now.isoformat(timespec="seconds"))
                elif _same(e["field"], live, e["new"]):
                    e["status"], e["verified_by"], e["issues"] = "반영됨", how, ""
                    e.setdefault("confirmed_at", now.isoformat(timespec="seconds"))
                elif hours >= LIVE_HOURS:
                    e["status"] = "반영 안 됨"
                    e["issues"] = (f"내 리스팅에는 저장됐지만 {LIVE_HOURS}시간이 지나도 {how}가 다름(현재: {live[:80]}) — "
                                   "브랜드 소유자·다른 기여가 우선이거나 Amazon이 채택하지 않았을 수 있음")
                else:
                    e["status"] = "제출값 저장됨"
                    e["issues"] = f"내 리스팅에는 저장됨 · {how}는 아직 이전 값(현재: {live[:80]}) · 보통 15분~몇 시간"
            elif errs:
                e["status"], e["issues"] = "반영 안 됨", "; ".join(errs[:3])
            elif hours >= PENDING_HOURS:
                e["status"] = "반영 안 됨"
                e["issues"] = (f"{PENDING_HOURS}시간이 지나도 현재 값이 다름(현재: {cur[:80]}) — "
                               "브랜드/카탈로그 기여가 우선이거나 Amazon이 거부했을 수 있음")
            else:
                e["status"], e["issues"] = "처리 중", f"아직 이전 값(현재: {cur[:80]}) · 적용 후 {hours:.1f}시간"
        except Exception as ex:
            e["issues"] = f"확인 실패: {str(ex)[:150]}"
        e["checked_at"] = now.isoformat(timespec="seconds")
        log(f"   {e['sku']} · {e['field']}: {e.get('status', '')}" + (f" — {e['issues'][:120]}" if e.get("issues") else ""))
    _log_save(entries)
    cnt: Dict[str, int] = {}
    for e in todo:
        cnt[e.get("status", "")] = cnt.get(e.get("status", ""), 0) + 1
    return cnt


def pending() -> List[Dict]:
    """아직 확인이 필요한 기록(접수됨·처리 중, 최근 30일)."""
    now = datetime.now()
    return [e for e in _log_load() if (e.get("status", "접수됨") in ("접수됨", "처리 중", "제출값 저장됨")
                                       or (e.get("status") == "반영됨" and not e.get("verified_by"))) and not
            (e.get("kind") == "적용" and e.get("reverted")) and (now - datetime.fromisoformat(e["time"])).days <= 30]


def alerts() -> List[List[str]]:
    """오늘의 요약용: 반영 안 됨 → P2, 오래 처리 중 → P3."""
    out = []
    now = datetime.now()
    for e in _log_load():
        if (now - datetime.fromisoformat(e["time"])).days > 7:
            continue
        st = e.get("status", "접수됨")
        if st == "반영 안 됨":
            out.append(["P2", "수정 반영 안 됨", f"{e['sku']} {e['field']}: {e.get('new', '')[:60]} ({_fmt(e['time'])} 적용)",
                        e.get("issues", "")[:150] or "리스팅·가격 수정 창 → 변경 기록에서 확인"])
        elif st in ("접수됨", "처리 중", "제출값 저장됨") and (now - datetime.fromisoformat(e["time"])).total_seconds() > 6 * 3600:
            out.append(["P3", "수정 처리 중", f"{e['sku']} {e['field']}: 적용 후 6시간 넘게 확인 안 됨", "변경 기록 → ‘반영 확인’"])
    return out


def apply(rows: List[Dict], allow_big: bool = False, log=print) -> List[Dict]:
    """검사를 통과한(ok) 줄만 실제 적용. 원래 값은 기록."""
    c = sp_api.client()
    seller = sp_api.seller_id()
    entries = _log_load()
    done = []
    todo = [r for r in rows if r.get("ok")][:MAX_BATCH]
    for i, r in enumerate(todo, 1):
        if r.get("big") and not allow_big:
            r["result"] = "건너뜀: 큰 가격 변경(허용 안 함)"
            log(f"   [{i}/{len(todo)}] {r['sku']} · {r['field']}: {r['result']}")
            done.append(r)
            continue
        try:
            res = _patch(c, seller, r["sku"], r["product_type"], [r["patch"]], preview=False)
            err, _ = _issues(res)
            r["result"] = ("접수됨(ACCEPTED) — 반영까지 몇 분~몇 시간" if res.get("status") == "ACCEPTED" and not err
                           else f"실패: {res.get('status')} {'; '.join(err)}")
        except Exception as e:
            r["result"] = f"실패: {str(e)[:200]}"
        log(f"   [{i}/{len(todo)}] {r['sku']} · {r['field']}: {r['old'][:40]} → {r['new'][:40]} · {r['result']}")
        if not r["result"].startswith("실패"):
            entries.append({"id": f"{datetime.now():%Y%m%d%H%M%S}-{i}", "time": datetime.now().isoformat(timespec="seconds"),
                            "kind": "적용", "sku": r["sku"], "asin": r.get("asin", ""), "field": r["field"],
                            "attr": r["attr"], "product_type": r["product_type"], "old": r["old"], "new": r["new"],
                            "old_attr": r.get("old_attr"), "result": r["result"], "memo": r.get("memo", ""),
                            "reverted": False, "status": "접수됨"})
        done.append(r)
        time.sleep(0.3)
    _log_save(entries)
    _mark_applied([r for r in done if str(r.get("result", "")).startswith("접수")])
    return done


def _mark_applied(rows: List[Dict]) -> None:
    """요청서에서 적용한 줄의 Y를 ‘완료’로 바꿔 두 번 적용되지 않게."""
    if not rows or not REQ_FILE.exists():
        return
    try:
        from openpyxl import load_workbook
        wb = load_workbook(REQ_FILE)
        ws = wb["변경 요청"]
        for r in rows:
            ws.cell(r["row"], 1, f"완료 {datetime.now():%m-%d %H:%M}")
        wb.save(REQ_FILE)
    except Exception:
        pass   # 요청서가 Excel에서 열려 있으면 저장 실패 — 기록은 남아 있음


def history() -> List[Dict]:
    return list(reversed(_log_load()))


def revert(ids: List[str], log=print) -> List[Dict]:
    """기록의 원래 값으로 되돌림(속성 전체를 원래대로)."""
    c = sp_api.client()
    seller = sp_api.seller_id()
    entries = _log_load()
    out = []
    for e in entries:
        if e["id"] not in ids or e.get("reverted") or e.get("kind") != "적용":
            continue
        try:
            if e.get("old_attr"):
                patch = {"op": "replace", "path": f"/attributes/{e['attr']}", "value": e["old_attr"]}
            else:
                patch = {"op": "delete", "path": f"/attributes/{e['attr']}"}
            res = _patch(c, seller, e["sku"], e["product_type"], [patch], preview=False)
            err, _ = _issues(res)
            ok = res.get("status") == "ACCEPTED" and not err
            msg = "되돌림 접수됨" if ok else f"되돌리기 실패: {res.get('status')} {'; '.join(err)}"
        except Exception as ex:
            ok, msg = False, f"되돌리기 실패: {str(ex)[:200]}"
        log(f"   {e['sku']} · {e['field']}: {msg}")
        if ok:
            e["reverted"] = True
            entries.append({"id": f"{datetime.now():%Y%m%d%H%M%S}-r", "time": datetime.now().isoformat(timespec="seconds"),
                            "kind": "되돌리기", "sku": e["sku"], "asin": e.get("asin", ""), "field": e["field"], "attr": e["attr"],
                            "product_type": e["product_type"], "old": e["new"], "new": e["old"], "result": msg,
                            "memo": f"{e['id']} 되돌림", "reverted": False, "status": "접수됨"})
            e["status"] = "되돌림"
        out.append({"sku": e["sku"], "field": e["field"], "result": msg})
        time.sleep(0.3)
    _log_save(entries)
    return out

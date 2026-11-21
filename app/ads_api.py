"""
Amazon Ads API(광고) — 광고 보고서를 API로 받기(조회만, 입찰·예산·키워드는 바꾸지 않음)

설정(화면 ‘⚙ 설정’ → ‘Amazon Ads API’ 탭): Ads 앱의 LWA Client ID · Client Secret,
‘① 인증 페이지 열기’ → 승인 → 이동한 주소를 붙여 넣어 Refresh Token 받기, ‘프로필 찾기’로 미국 판매자 프로필 선택.

보고서(v3, 최근 30일, 하루 단위) → data/<미국 날짜>/rp_<종류>.xlsx (브라우저로 받던 파일과 같은 열 이름)
  sp_search_term        검색어      (spSearchTerm, groupBy searchTerm)
  sp_targeting          키워드·타겟 (spTargeting, groupBy targeting)
  sp_campaign           캠페인·게재 위치 (spCampaigns, groupBy campaign + campaignPlacement)
  sp_advertised_product 광고 상품   (spAdvertisedProduct, groupBy advertiser)
API가 안 되면 원래 방식(광고 콘솔 Reporting 화면)으로 진행합니다.
"""

import gzip
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import sp_api

ADS_ENDPOINT = "https://advertising-api.amazon.com"
AUTH_URL = "https://www.amazon.com/ap/oa"
TOKEN_URL = "https://api.amazon.com/auth/o2/token"
RAW_DIR = sp_api.BASE_DIR / "reports" / "api_raw"
DAYS = 30

COMMON = ["date", "campaignName", "campaignId", "adGroupName", "adGroupId", "impressions", "clicks", "cost",
          "purchases7d", "sales7d", "unitsSoldClicks7d"]
SPECS = {
    "sp_search_term": {"type": "spSearchTerm", "groupBy": ["searchTerm"],
                       "columns": COMMON + ["searchTerm", "keyword", "targeting", "matchType", "keywordType", "keywordBid",
                                            "campaignStatus", "campaignBudgetAmount"],
                       "minimal": COMMON + ["searchTerm", "targeting", "matchType"]},
    "sp_targeting": {"type": "spTargeting", "groupBy": ["targeting"],
                     "columns": COMMON + ["targeting", "keyword", "matchType", "keywordType", "keywordBid", "keywordId",
                                          "campaignStatus", "campaignBudgetAmount", "topOfSearchImpressionShare"],
                     "minimal": COMMON + ["targeting", "matchType", "keywordBid"]},
    "sp_campaign": {"type": "spCampaigns", "groupBy": ["campaign", "campaignPlacement"],
                    "columns": ["date", "campaignName", "campaignId", "campaignStatus", "campaignBudgetAmount",
                                "placementClassification", "impressions", "clicks", "cost", "purchases7d", "sales7d",
                                "unitsSoldClicks7d", "topOfSearchImpressionShare"],
                    "minimal": ["date", "campaignName", "campaignId", "campaignBudgetAmount", "placementClassification",
                                "impressions", "clicks", "cost", "purchases7d", "sales7d"]},
    "sp_advertised_product": {"type": "spAdvertisedProduct", "groupBy": ["advertiser"],
                              "columns": COMMON + ["advertisedAsin", "advertisedSku", "campaignStatus"],
                              "minimal": COMMON + ["advertisedAsin", "advertisedSku"]},
}
# 스폰서드 브랜드 검색어 — 브라우저 검색어 보고서처럼 sp_search_term 파일에 합침(SB 캠페인이 없거나 실패해도 SP는 그대로)
SB_SPEC = {"adProduct": "SPONSORED_BRANDS", "type": "sbSearchTerm", "groupBy": ["searchTerm"],
           "columns": ["date", "campaignName", "campaignId", "adGroupName", "adGroupId", "searchTerm", "keywordText",
                       "matchType", "impressions", "clicks", "cost", "purchases", "sales", "unitsSold"]}
# API 열 → 브라우저 보고서 열 이름(ads_common.FIELDS가 알아보는 이름)
COLNAME = {"date": "Date", "campaignName": "Campaign name", "campaignId": "Campaign ID", "adGroupName": "Ad group name",
           "adGroupId": "Ad group ID", "campaignStatus": "Campaign status", "campaignBudgetAmount": "Campaign budget amount",
           "searchTerm": "Customer search term", "targeting": "Targeting", "matchType": "Match type", "keywordBid": "Target bid",
           "placementClassification": "Placement classification", "advertisedAsin": "Advertised ASIN",
           "advertisedSku": "Advertised SKU", "impressions": "Impressions", "clicks": "Clicks", "cost": "Total cost",
           "purchases7d": "7 Day Total Orders", "sales7d": "7 Day Total Sales", "unitsSoldClicks7d": "7 Day Total Units",
           "topOfSearchImpressionShare": "Top-of-search impression share", "keywordId": "Keyword ID",
           "keywordType": "Keyword type", "keyword": "Keyword text", "keywordText": "Targeting",
           "purchases": "7 Day Total Orders", "sales": "7 Day Total Sales", "unitsSold": "7 Day Total Units"}


class AdsError(RuntimeError):
    def __init__(self, status: int, msg: str):
        super().__init__(f"Ads API {status}: {msg}")
        self.status = status


def cfg() -> Dict:
    return sp_api.load_config()


def ready(c: Optional[Dict] = None) -> bool:
    c = c or cfg()
    return bool(c.get("use_ads_api", True) and c.get("ads_client_id") and c.get("ads_client_secret")
                and c.get("ads_refresh_token") and c.get("ads_profile_id"))


def missing(c: Optional[Dict] = None) -> List[str]:
    c = c or cfg()
    names = {"ads_client_id": "Ads Client ID", "ads_client_secret": "Ads Client Secret",
             "ads_refresh_token": "Ads Refresh Token", "ads_profile_id": "Ads 프로필 ID"}
    return [n for k, n in names.items() if not c.get(k)]


# ================================================================ 인증(처음 한 번)

def consent_url(c: Dict) -> str:
    q = {"client_id": c.get("ads_client_id", ""), "scope": "advertising::campaign_management",
         "response_type": "code", "redirect_uri": c.get("ads_redirect_uri") or "https://amazon.com"}
    return AUTH_URL + "?" + urllib.parse.urlencode(q)


def exchange_code(c: Dict, code_or_url: str) -> str:
    """승인 후 이동한 주소(…?code=XXXX…) 또는 code 값 → Refresh Token."""
    code = code_or_url.strip()
    if "code=" in code:
        code = urllib.parse.parse_qs(urllib.parse.urlparse(code).query).get("code", [""])[0]
    if not code:
        raise AdsError(0, "주소에서 code 값을 찾지 못했습니다")
    body = urllib.parse.urlencode({"grant_type": "authorization_code", "code": code,
                                   "redirect_uri": c.get("ads_redirect_uri") or "https://amazon.com",
                                   "client_id": c["ads_client_id"], "client_secret": c["ads_client_secret"]}).encode()
    req = urllib.request.Request(TOKEN_URL, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode())["refresh_token"]
    except urllib.error.HTTPError as e:
        raise AdsError(e.code, "토큰 교환 실패(코드는 몇 분 안에 한 번만 쓸 수 있음, Redirect URI가 앱 설정과 같은지 확인): "
                       + e.read().decode(errors="replace")[:300])


# ================================================================ 클라이언트

class AdsClient:
    def __init__(self, c: Optional[Dict] = None, need_profile: bool = True):
        self.c = c or cfg()
        need = ["ads_client_id", "ads_client_secret", "ads_refresh_token"] + (["ads_profile_id"] if need_profile else [])
        miss = [k for k in need if not self.c.get(k)]
        if miss:
            raise AdsError(0, "Ads API 설정이 비어 있습니다: " + ", ".join(missing(self.c)))
        self._tok, self._exp = None, 0.0

    def token(self) -> str:
        if self._tok and time.time() < self._exp - 60:
            return self._tok
        body = urllib.parse.urlencode({"grant_type": "refresh_token", "refresh_token": self.c["ads_refresh_token"],
                                       "client_id": self.c["ads_client_id"],
                                       "client_secret": self.c["ads_client_secret"]}).encode()
        req = urllib.request.Request(TOKEN_URL, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                d = json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            raise AdsError(e.code, "인증 실패(Ads Client ID·Secret·Refresh Token 확인): " + e.read().decode(errors="replace")[:300])
        self._tok, self._exp = d["access_token"], time.time() + int(d.get("expires_in", 3600))
        return self._tok

    def request(self, method: str, path: str, body=None, ctype: str = "application/json", accept: str = "application/json",
                profile: bool = True, tries: int = 6):
        data = json.dumps(body).encode() if body is not None else None
        wait = 3.0
        for i in range(tries):
            h = {"Authorization": f"Bearer {self.token()}", "Amazon-Advertising-API-ClientId": self.c["ads_client_id"],
                 "Content-Type": ctype, "Accept": accept}
            if profile:
                h["Amazon-Advertising-API-Scope"] = str(self.c["ads_profile_id"])
            req = urllib.request.Request(ADS_ENDPOINT + path, data=data, method=method, headers=h)
            try:
                with urllib.request.urlopen(req, timeout=60) as r:
                    txt = r.read().decode("utf-8")
                    return json.loads(txt) if txt else {}
            except urllib.error.HTTPError as e:
                msg = e.read().decode(errors="replace")[:500]
                if e.code in (429, 500, 502, 503, 504) and i < tries - 1:
                    time.sleep(wait)
                    wait = min(wait * 2, 60)
                    continue
                if e.code == 425:          # 같은 보고서가 이미 요청됨 → 응답에 기존 ID
                    m = re.search(r"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})", msg)
                    if m:
                        return {"reportId": m.group(1), "duplicate": True}
                raise AdsError(e.code, msg)
            except urllib.error.URLError as e:
                if i < tries - 1:
                    time.sleep(wait)
                    continue
                raise AdsError(0, f"네트워크 오류: {e.reason}")

    def profiles(self) -> List[Dict]:
        return self.request("GET", "/v2/profiles", profile=False) or []

    def create(self, key: str, start: date, end: date, minimal: bool = False) -> str:
        sp = SB_SPEC if key == "sb_search_term" else SPECS[key]
        body = {"name": f"AIAGENT_{key}_{start}_{end}_{datetime.now():%H%M%S}", "startDate": start.isoformat(),
                "endDate": end.isoformat(),
                "configuration": {"adProduct": sp.get("adProduct", "SPONSORED_PRODUCTS"), "groupBy": sp["groupBy"],
                                  "columns": sp["minimal" if minimal and "minimal" in sp else "columns"], "reportTypeId": sp["type"],
                                  "timeUnit": "DAILY", "format": "GZIP_JSON"}}
        d = self.request("POST", "/reporting/reports", body, ctype="application/vnd.createasyncreportrequest.v3+json")
        return d["reportId"]

    def status(self, rid: str) -> Dict:
        return self.request("GET", f"/reporting/reports/{rid}")

    def download(self, url: str) -> List[Dict]:
        with urllib.request.urlopen(url, timeout=300) as r:
            data = r.read()
        if data[:2] == b"\x1f\x8b":
            data = gzip.decompress(data)
        return json.loads(data.decode("utf-8") or "[]")


# ================================================================ 보고서 → 기존 형식

def to_rows(items: List[Dict], product: str = "Sponsored Products") -> List[Dict]:
    out = []
    for it in items:
        r = {"Ad product": product}
        for k, v in it.items():
            name = COLNAME.get(k, k)
            if k == "matchType":
                if not v:
                    continue
                if str(v).upper().startswith("TARGETING_EXPRESSION"):   # 자동·상품 타겟은 화면 보고서처럼 '-'
                    v = "-"
            r[name] = v
        if not r.get("Match type") and it.get("keywordType"):
            r["Match type"] = it["keywordType"]
        if not r.get("Targeting") and it.get("keyword"):
            r["Targeting"] = it["keyword"]
        out.append(r)
    return out


def request_only(keys: List[str], today: Optional[date] = None, days: int = DAYS) -> Dict[str, str]:
    """보고서만 미리 요청(기다리지 않음) → {key: reportId}. 매일 자동 실행이 시작할 때 불러 두면 광고 단계에서 바로 받음."""
    import ads_common as ac
    today = today or ac.pacific_today()
    end = today - timedelta(days=1)
    start = end - timedelta(days=days - 1)
    cli = AdsClient()
    ids: Dict[str, str] = {}
    for k in keys:
        try:
            ids[k] = cli.create(k, start, end)
        except AdsError as e:
            if e.status != 400:
                raise
            ids[k] = cli.create(k, start, end, minimal=True)
        time.sleep(1.5)
    if "sp_search_term" in keys:
        try:
            ids["sb_search_term"] = cli.create("sb_search_term", start, end)
        except Exception:
            pass
    return ids


def fetch_reports(keys: List[str], today: Optional[date] = None, days: int = DAYS, timeout_min: float = 60,
                  ids_pre: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    """keys 보고서를 한꺼번에 요청 → 생성되는 대로 받기 → rp_<key>.xlsx. {key: 결과}. 실패한 key는 결과에 없음.
    ids_pre: 미리 요청해 둔 보고서 ID(있으면 다시 요청하지 않음)."""
    import ads_common as ac
    today = today or ac.pacific_today()
    end = today - timedelta(days=1)
    start = end - timedelta(days=days - 1)
    cli = AdsClient()
    ids: Dict[str, str] = {}
    pre = dict(ids_pre or {})
    for k in keys:
        if pre.get(k):
            ids[k] = pre[k]
            print(f"   Ads API 보고서: {k} — 미리 요청해 둔 보고서 사용")
            continue
        try:
            ids[k] = cli.create(k, start, end)
        except AdsError as e:
            if e.status == 400:      # 열 이름이 맞지 않으면 최소 열로 다시
                print(f"     {k}: 열 구성 오류 → 기본 열로 다시 요청({str(e)[:120]})")
                ids[k] = cli.create(k, start, end, minimal=True)
            else:
                raise
        print(f"   Ads API 보고서 요청: {k} ({start} ~ {end})")
        time.sleep(1.5)
    sb_id, sb_rows = pre.get("sb_search_term"), []
    if "sp_search_term" in keys and not sb_id:
        try:
            sb_id = cli.create("sb_search_term", start, end)
            print("   Ads API 보고서 요청: 스폰서드 브랜드 검색어(검색어 보고서에 합침)")
        except Exception as e:
            print(f"     스폰서드 브랜드 검색어는 건너뜀({str(e)[:120]})")
    done: Dict[str, str] = {}
    t0 = time.time()
    n = 0
    while len(done) < len(ids):
        if sb_id:                      # SB 검색어를 먼저 받아 둠(SP 검색어 파일에 합치려고)
            try:
                st = cli.status(sb_id)
                if st.get("status") == "COMPLETED" and st.get("url"):
                    sb_rows = to_rows(cli.download(st["url"]), "Sponsored Brands")
                    sb_id = None
                elif st.get("status") == "FAILED" or (time.time() - t0) / 60 > timeout_min / 2:
                    print("     스폰서드 브랜드 검색어 보고서를 받지 못해 SP만 사용합니다")
                    sb_id = None
            except Exception as e:
                print(f"     스폰서드 브랜드 검색어 확인 실패 → SP만 사용({str(e)[:100]})")
                sb_id = None
        for k, rid in ids.items():
            if k in done:
                continue
            if k == "sp_search_term" and sb_id:      # SB가 끝날 때까지 기다렸다가 합침
                continue
            st = cli.status(rid)
            s = st.get("status")
            if s == "COMPLETED" and st.get("url"):
                items = cli.download(st["url"])
                RAW_DIR.mkdir(parents=True, exist_ok=True)
                (RAW_DIR / f"ads_{k}_{today}.json").write_text(json.dumps(items)[:50_000_000], encoding="utf-8")
                rows = to_rows(items)
                if k == "sp_search_term" and sb_rows:
                    rows += sb_rows
                    print(f"     스폰서드 브랜드 검색어 {len(sb_rows)}행을 합쳤습니다")
                out = ac.day_dir(today) / f"rp_{k}.xlsx"
                ac.save_rows(rows, out, k)
                done[k] = f"{len(rows)}행"
                print(f"   ✅ Ads API {k}: {len(rows)}행 → {out.name}")
            elif s == "FAILED":
                raise AdsError(0, f"{k} 보고서 생성 실패: {st.get('failureReason', '')}")
        if len(done) == len(ids):
            break
        mins = (time.time() - t0) / 60
        if mins >= timeout_min:
            raise AdsError(0, f"{timeout_min:.0f}분 안에 생성되지 않은 보고서: {', '.join(k for k in ids if k not in done)}")
        n += 1
        if n % 3 == 1:
            print(f"     Ads API 보고서 생성 중: 남은 {len(ids) - len(done)}개 · {mins:.0f}분 경과")
        time.sleep(30 if mins < 5 else 60)
    return done


def try_ads(label: str, fn, *a, **kw):
    """Ads API 우선 → 설정 없거나 실패하면 None(원래 방식). 데이터 출처 기록."""
    c = cfg()
    if not c.get("use_ads_api", True):
        sp_api._src(label, "브라우저", "설정에서 Ads API 사용 안 함")
        return None
    if not ready(c):
        print(f"   ℹ {label}: Ads API 설정이 아직 완성되지 않았습니다(빠진 항목: {', '.join(missing(c))}) → 원래 방식(광고 콘솔 화면)으로 진행")
        sp_api._src(label, "브라우저", "Ads API 설정 없음: " + ", ".join(missing(c)))
        return None
    try:
        print(f"   ▶ {label}: Amazon Ads API로 가져오는 중...")
        t0 = time.time()
        res = fn(*a, **kw)
        sp_api._src(label, "API", f"Ads API {time.time() - t0:.0f}초")
        return res
    except Exception as e:
        print(f"   ⚠ {label}: Ads API 실패({type(e).__name__}: {str(e)[:200]}) → 원래 방식(광고 콘솔 화면)으로 진행")
        sp_api._src(label, "API 실패 → 브라우저", f"{type(e).__name__}: {str(e)[:250]}")
        return None


def find_profiles(c: Dict) -> List[Dict]:
    cli = AdsClient(c, need_profile=False)
    out = []
    for p in cli.profiles():
        ai = p.get("accountInfo") or {}
        out.append({"profileId": str(p.get("profileId")), "country": p.get("countryCode", ""), "type": ai.get("type", ""),
                    "name": ai.get("name", ""), "marketplace": ai.get("marketplaceStringId", "")})
    return out


def test(c: Dict) -> str:
    cli = AdsClient(c)
    cli.token()
    prof = [p for p in find_profiles(c) if p["profileId"] == str(c.get("ads_profile_id"))]
    return "✅ Ads API 연결 성공" + (f" · {prof[0]['name']}({prof[0]['country']}, {prof[0]['type']})" if prof else
                                   " · (주의: 프로필 ID가 목록에 없음)")

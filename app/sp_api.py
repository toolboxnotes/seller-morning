"""
Amazon Selling Partner API(SP-API) 연결

  - 설정: %LOCALAPPDATA%\\AmazonAIProject\\sp_api.json (OneDrive 밖, 이 컴퓨터에만 저장)
          client secret · refresh token은 Windows 사용자 계정으로 암호화(DPAPI)해서 저장합니다.
  - 화면 프로그램의 ‘설정(API·메일)’에서 입력·연결 테스트
  - 프로젝트 폴더에 sp_api_import.json 이 있으면 한 번 읽어서 설정에 넣고 그 파일은 지웁니다.

  API 우선 → 실패하면 원래 방식(브라우저)으로 자동 전환하는 기능들이 이 모듈을 사용합니다.
  조회(GET)·보고서 요청만 사용하며, 상품·가격·광고는 바꾸지 않습니다.
"""

import base64
import gzip
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Tuple, Dict, Iterable, List, Optional

BASE_DIR = Path(__file__).resolve().parent.parent
import platform_util as pu
CONFIG_DIR = pu.local_dir()          # Windows: %LOCALAPPDATA%\AmazonAIProject · macOS: ~/Library/Application Support/AmazonAIProject
CONFIG_FILE = CONFIG_DIR / "sp_api.json"
IMPORT_FILE = BASE_DIR / "sp_api_import.json"
RAW_DIR = BASE_DIR / "reports" / "api_raw"

ENDPOINTS = {"na": "https://sellingpartnerapi-na.amazon.com", "eu": "https://sellingpartnerapi-eu.amazon.com",
             "fe": "https://sellingpartnerapi-fe.amazon.com"}
TOKEN_URL = "https://api.amazon.com/auth/o2/token"
US = "ATVPDKIKX0DER"
SECRET_KEYS = ("client_secret", "refresh_token", "ads_client_secret", "ads_refresh_token")
DEFAULTS = {"client_id": "", "client_secret": "", "refresh_token": "", "seller_id": "",
            "marketplace_id": US, "region": "na", "use_api": True, "report_email": "",
            # 판매자 정보(코드에 넣지 않고 설정에만 저장): 브랜드 이름
            "brand": "",
            # Amazon Ads API(광고) — SP-API와 별도 앱·토큰
            "ads_client_id": "", "ads_client_secret": "", "ads_refresh_token": "", "ads_profile_id": "",
            "ads_redirect_uri": "https://amazon.com", "use_ads_api": True,
            "token_saved": "", "ads_token_saved": "", "notify": True}


class SPAPIError(RuntimeError):
    def __init__(self, status: int, msg: str):
        super().__init__(f"SP-API {status}: {msg}")
        self.status = status


# ================================================================ 비밀값 보관(Windows DPAPI · macOS 키체인)

def _protect(text: str, key: str = "") -> str:
    """Windows: 계정 암호화(DPAPI) · macOS: 키체인에 넣고 설정 파일에는 ‘kc:이름’만."""
    if not text:
        return ""
    if pu.IS_MAC and key:
        if pu.keychain_set(key, text):
            return "kc:" + key
    if sys.platform.startswith("win"):
        try:
            import ctypes
            from ctypes import wintypes

            class BLOB(ctypes.Structure):
                _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
            data = text.encode("utf-8")
            inp = BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char)))
            out = BLOB()
            if ctypes.windll.crypt32.CryptProtectData(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
                enc = ctypes.string_at(out.pbData, out.cbData)
                ctypes.windll.kernel32.LocalFree(out.pbData)
                return "dpapi:" + base64.b64encode(enc).decode()
        except Exception:
            pass
    return "b64:" + base64.b64encode(text.encode("utf-8")).decode()


def _unprotect(val: str) -> str:
    if not val:
        return ""
    if val.startswith("kc:"):
        v = pu.keychain_get(val[3:])
        if v is None:
            raise RuntimeError("키체인에서 비밀값을 찾지 못했습니다(키체인 접근을 거부했거나 다른 Mac). 설정에서 다시 입력하세요.")
        return v
    if val.startswith("b64:"):
        return base64.b64decode(val[4:]).decode("utf-8")
    if val.startswith("dpapi:"):
        import ctypes
        from ctypes import wintypes

        class BLOB(ctypes.Structure):
            _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
        raw = base64.b64decode(val[6:])
        inp = BLOB(len(raw), ctypes.cast(ctypes.create_string_buffer(raw, len(raw)), ctypes.POINTER(ctypes.c_char)))
        out = BLOB()
        if ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
            dec = ctypes.string_at(out.pbData, out.cbData)
            ctypes.windll.kernel32.LocalFree(out.pbData)
            return dec.decode("utf-8")
        raise RuntimeError("저장된 비밀값을 풀 수 없습니다(다른 Windows 사용자로 저장된 설정). 설정에서 다시 입력하세요.")
    return val


# ================================================================ 설정

def load_config() -> Dict:
    _import_file()
    cfg = dict(DEFAULTS)
    try:
        raw = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        cfg.update(raw)
        for k in SECRET_KEYS:
            cfg[k] = _unprotect(raw.get(k, ""))
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"⚠ SP-API 설정을 읽지 못했습니다: {e}")
    return cfg


def save_config(cfg: Dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    out = {k: cfg.get(k, v) for k, v in DEFAULTS.items()}
    for k in SECRET_KEYS:
        out[k] = _protect((out.get(k) or "").strip(), k)
    for k in ("client_id", "seller_id", "marketplace_id", "report_email", "brand", "ads_client_id", "ads_profile_id",
              "ads_redirect_uri"):
        out[k] = str(out.get(k) or "").strip()
    # 토큰을 바꾼 날짜 기록(설정 화면에 표시)
    try:
        old = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        old_rt = _unprotect(old.get("refresh_token", ""))
        old_ads = _unprotect(old.get("ads_refresh_token", ""))
    except Exception:
        old_rt = old_ads = None
    today = datetime.now().strftime("%Y-%m-%d")
    if (cfg.get("refresh_token") or "") and (cfg.get("refresh_token") or "").strip() != old_rt:
        out["token_saved"] = today
    if (cfg.get("ads_refresh_token") or "") and (cfg.get("ads_refresh_token") or "").strip() != old_ads:
        out["ads_token_saved"] = today
    CONFIG_FILE.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")


def _import_file() -> None:
    """프로젝트 폴더의 sp_api_import.json → 암호화된 설정으로 옮기고 원본은 삭제(OneDrive에 남지 않게)."""
    if not IMPORT_FILE.exists():
        return
    try:
        data = json.loads(IMPORT_FILE.read_text(encoding="utf-8"))
        cur = dict(DEFAULTS)
        try:
            raw = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            cur.update(raw)
            for k in SECRET_KEYS:
                cur[k] = _unprotect(raw.get(k, ""))
        except Exception:
            pass
        cur.update({k: v for k, v in data.items() if v not in (None, "")})
        save_config(cur)
        IMPORT_FILE.unlink()
    except Exception as e:
        print(f"⚠ sp_api_import.json을 가져오지 못했습니다: {e}")


def settings() -> Dict:
    return load_config()


def report_email() -> str:
    return (load_config().get("report_email") or "").strip()


def brand() -> str:
    """브랜드 이름: 설정 → 없으면 광고 설정(ads_config.xlsx)의 ‘브랜드 키워드’ 첫 단어 → 없으면 빈 값."""
    b = (load_config().get("brand") or "").strip()
    if b:
        return b
    try:
        from openpyxl import load_workbook
        cf = BASE_DIR / "reports" / "ads" / "ads_config.xlsx"
        ws = load_workbook(cf, read_only=True)["설정"]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] == "브랜드 키워드" and r[1]:
                return str(r[1]).split(",")[0].strip().upper()
    except Exception:
        pass
    return ""


def seller_id() -> str:
    """Seller ID(설정에만 저장). 비어 있으면 알기 쉬운 오류."""
    s = (load_config().get("seller_id") or "").strip()
    if not s:
        raise RuntimeError("Seller ID가 설정에 없습니다. 화면 프로그램 ‘설정 → SP-API’에서 Seller ID(판매자 토큰)를 입력하고 저장하세요.")
    return s


def profile_missing(cfg: Optional[Dict] = None) -> list:
    """설정에서 비어 있는 판매자 정보(화면에 안내용)."""
    c = cfg or load_config()
    out = []
    if not (c.get("seller_id") or "").strip():
        out.append("Seller ID")
    if not (c.get("report_email") or "").strip():
        out.append("확인 메일")
    if not brand():
        out.append("브랜드 이름")
    return out


def ready(cfg: Optional[Dict] = None) -> bool:
    c = cfg or load_config()
    return bool(c.get("use_api") and c.get("client_id") and c.get("client_secret") and c.get("refresh_token"))


def missing(cfg: Optional[Dict] = None) -> List[str]:
    c = cfg or load_config()
    names = {"client_id": "LWA Client ID", "client_secret": "LWA Client Secret", "refresh_token": "Refresh Token"}
    return [n for k, n in names.items() if not c.get(k)]


# ================================================================ 클라이언트

class Client:
    def __init__(self, cfg: Optional[Dict] = None):
        self.cfg = cfg or load_config()
        if not (self.cfg.get("client_id") and self.cfg.get("client_secret") and self.cfg.get("refresh_token")):
            raise SPAPIError(0, "SP-API 설정이 비어 있습니다: " + ", ".join(missing(self.cfg)))
        self.base = ENDPOINTS.get(self.cfg.get("region", "na"), ENDPOINTS["na"])
        self.mp = self.cfg.get("marketplace_id") or US
        self._token = None
        self._exp = 0.0

    # ---- 인증
    def token(self) -> str:
        if self._token and time.time() < self._exp - 60:
            return self._token
        body = urllib.parse.urlencode({"grant_type": "refresh_token", "refresh_token": self.cfg["refresh_token"],
                                       "client_id": self.cfg["client_id"], "client_secret": self.cfg["client_secret"]}).encode()
        req = urllib.request.Request(TOKEN_URL, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                d = json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            raise SPAPIError(e.code, "인증 실패(토큰·Client ID·Secret 확인): " + e.read().decode(errors="replace")[:300])
        self._token = d["access_token"]
        self._exp = time.time() + int(d.get("expires_in", 3600))
        return self._token

    # ---- 요청(429·5xx는 기다렸다가 다시)
    def request(self, method: str, path: str, params: Optional[Dict] = None, body: Optional[Dict] = None,
                tries: int = 6) -> Dict:
        q = ""
        if params:
            q = "?" + urllib.parse.urlencode({k: (",".join(v) if isinstance(v, (list, tuple)) else v)
                                              for k, v in params.items() if v not in (None, "")})
        data = json.dumps(body).encode() if body is not None else None
        wait = 2.0
        for i in range(max(tries, 10)):
            req = urllib.request.Request(self.base + path + q, data=data, method=method, headers={
                "x-amz-access-token": self.token(), "Content-Type": "application/json", "Accept": "application/json",
                "User-Agent": f"SellerMorning/{__import__('version').VERSION} (Language=Python)"})
            try:
                with urllib.request.urlopen(req, timeout=60) as r:
                    txt = r.read().decode("utf-8")
                    return json.loads(txt) if txt else {}
            except urllib.error.HTTPError as e:
                msg = e.read().decode(errors="replace")[:400]
                if e.code == 429 and i < max(tries, 10) - 1:
                    # 요청 한도 초과(QuotaExceeded): 한도가 채워질 때까지 기다렸다가 다시(보고서 요청은 1분에 1개 정도)
                    try:
                        rate = float(e.headers.get("x-amzn-RateLimit-Limit") or 0)
                    except (TypeError, ValueError):
                        rate = 0
                    need = (1.0 / rate * 1.1) if rate > 0 else 0
                    sleep = min(max(wait, need, 5.0), 120)
                    print(f"     (요청 한도 초과 — {sleep:.0f}초 기다렸다가 다시 시도 {i + 1}/{max(tries, 10) - 1})")
                    time.sleep(sleep)
                    wait = min(wait * 2, 90)
                    continue
                if e.code in (500, 502, 503, 504) and i < tries - 1:
                    time.sleep(wait)
                    wait = min(wait * 2, 60)
                    continue
                if e.code == 403:
                    msg = "권한 없음(앱에 이 API 역할이 없거나 토큰이 다른 계정용): " + msg
                raise SPAPIError(e.code, msg)
            except urllib.error.URLError as e:
                if i < tries - 1:
                    time.sleep(wait)
                    wait = min(wait * 2, 60)
                    continue
                raise SPAPIError(0, f"네트워크 오류: {e.reason}")
        raise SPAPIError(0, "요청 실패")

    def get(self, path: str, **params) -> Dict:
        return self.request("GET", path, params)

    # ---- 연결 확인
    def check(self) -> str:
        d = self.get("/sellers/v1/marketplaceParticipations")
        names = [f"{x['marketplace']['name']}({x['marketplace']['id']})" for x in d.get("payload", [])]
        return ", ".join(names) or "(마켓플레이스 없음)"

    # ---- 보고서
    def create_report(self, report_type: str, start: Optional[datetime] = None, end: Optional[datetime] = None,
                      options: Optional[Dict] = None) -> str:
        body = {"reportType": report_type, "marketplaceIds": [self.mp]}
        if start:
            body["dataStartTime"] = _iso(start)
        if end:
            body["dataEndTime"] = _iso(end)
        if options:
            body["reportOptions"] = options
        return self.request("POST", "/reports/2021-06-30/reports", body=body)["reportId"]

    def wait_report(self, report_id: str, timeout_min: float = 45, label: str = "") -> Optional[str]:
        """reportDocumentId(데이터 없음이면 None). 30초~1분 간격으로 상태 확인."""
        t0 = time.time()
        n = 0
        while True:
            d = self.get(f"/reports/2021-06-30/reports/{report_id}")
            st = d.get("processingStatus")
            if st == "DONE":
                return d.get("reportDocumentId")
            if st == "CANCELLED":
                return None            # 이 기간에 데이터 없음
            if st == "FATAL":
                doc = d.get("reportDocumentId")
                detail = ""
                if doc:
                    try:
                        detail = self.download(doc).decode("utf-8", errors="replace")[:300]
                    except Exception:
                        pass
                raise SPAPIError(0, f"보고서 생성 실패(FATAL) {detail}")
            mins = (time.time() - t0) / 60
            if mins >= timeout_min:
                raise SPAPIError(0, f"보고서가 {timeout_min:.0f}분 안에 생성되지 않음(상태 {st})")
            n += 1
            if n % 4 == 1:
                print(f"     API 보고서 생성 중{(' · ' + label) if label else ''}({st}, {mins:.0f}분 경과)...")
            time.sleep(30 if mins < 5 else 60)

    def download(self, doc_id: str) -> bytes:
        d = self.get(f"/reports/2021-06-30/documents/{doc_id}")
        with urllib.request.urlopen(d["url"], timeout=300) as r:
            data = r.read()
        if d.get("compressionAlgorithm") == "GZIP" or data[:2] == b"\x1f\x8b":
            data = gzip.decompress(data)
        return data

    def report(self, report_type: str, start: Optional[datetime] = None, end: Optional[datetime] = None,
               options: Optional[Dict] = None, timeout_min: float = 45, save_as: Optional[str] = None) -> Optional[Path]:
        """보고서 요청 → 생성 대기 → 다운로드 → reports/api_raw/ 에 저장. 데이터 없으면 None."""
        print(f"   API 보고서 요청: {report_type}" + (f" {start:%Y-%m-%d}~{end:%Y-%m-%d}" if start and end else ""))
        rid = self.create_report(report_type, start, end, options)
        doc = self.wait_report(rid, timeout_min, report_type)
        if not doc:
            print("     (이 기간·조건에는 데이터가 없습니다)")
            return None
        data = self.download(doc)
        RAW_DIR.mkdir(parents=True, exist_ok=True)
        name = save_as or f"{report_type}_{datetime.now():%Y%m%d_%H%M%S}"
        ext = ".json" if data[:1] in (b"{", b"[") else ".txt"
        out = RAW_DIR / (name + ext)
        out.write_bytes(data)
        print(f"   ✅ API 보고서 받음: {out.name} ({len(data) / 1024:,.0f} KB)")
        return out

    # ---- 목록 페이지 넘기기
    def pages(self, path: str, params: Dict, items_key: str, token_param: str = "pageToken",
              next_key: str = "nextToken", limit: int = 5000) -> Iterable[Dict]:
        p = dict(params)
        n = 0
        while True:
            d = self.get(path, **p)
            box = d.get("payload", d)
            for it in box.get(items_key, []) or []:
                yield it
                n += 1
                if n >= limit:
                    return
            nxt = box.get(next_key) or (box.get("pagination") or {}).get(next_key)
            if not nxt:
                return
            p[token_param] = nxt
            time.sleep(0.25)


def _iso(dt) -> str:
    if isinstance(dt, date) and not isinstance(dt, datetime):
        dt = datetime(dt.year, dt.month, dt.day)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ================================================================ "API 우선, 실패하면 원래 방식"

_client: Optional[Client] = None


def client() -> Client:
    global _client
    if _client is None:
        _client = Client()
    return _client


def _src(label: str, method: str, detail: str = "") -> None:
    try:
        import daily_bus
        daily_bus.source(label, method, detail)
    except Exception:
        pass


def try_api(label: str, fn, *args, **kwargs):
    """API로 해 보고, 설정이 없거나 실패하면 None(→ 호출한 쪽이 원래 방식으로 진행). 결과는 ‘데이터 출처’에 기록."""
    cfg = load_config()
    if not cfg.get("use_api"):
        _src(label, "브라우저", "설정에서 SP-API 사용 안 함")
        return None
    if not ready(cfg):
        print(f"   ℹ {label}: SP-API 설정이 아직 완성되지 않았습니다(빠진 항목: {', '.join(missing(cfg))}) "
              f"→ 화면 오른쪽 위 ‘⚙ 설정(API·메일)’에서 입력하세요. 이번에는 원래 방식으로 진행합니다.")
        _src(label, "브라우저", "SP-API 설정 없음: " + ", ".join(missing(cfg)))
        return None
    try:
        print(f"   ▶ {label}: SP-API로 가져오는 중...")
        t0 = time.time()
        res = fn(*args, **kwargs)
        if res is None:
            print(f"   ℹ {label}: API에 데이터가 없어 원래 방식으로 진행합니다.")
            _src(label, "브라우저", "API에 데이터 없음")
        else:
            _src(label, "API", f"{time.time() - t0:.0f}초")
        return res
    except Exception as e:
        print(f"   ⚠ {label}: API 실패({type(e).__name__}: {str(e)[:200]}) → 원래 방식(브라우저)으로 진행합니다.")
        _src(label, "API 실패 → 브라우저", f"{type(e).__name__}: {str(e)[:250]}")
        return None


def test_connection(cfg: Dict) -> str:
    c = Client(cfg)
    c.token()
    return "✅ 연결 성공 · " + c.check()


# ================================================================ 전체 점검(각 API를 가장 작은 요청으로 확인)

def self_check(cfg: Optional[Dict] = None) -> List[Dict]:
    """[{"항목", "결과"(✅/❌/⚠), "내용"}] — SP-API + Amazon Ads API."""
    cfg = cfg or load_config()
    out: List[Dict] = []
    if cfg.get("client_id") or cfg.get("refresh_token"):
        out += _sp_check(cfg)
    else:
        out.append({"항목": "SP-API", "결과": "⚠", "내용": "설정 없음(입력하지 않으면 원래 방식으로 진행)"})
    out += ads_check(cfg)
    return out


TOKEN_MAX_DAYS = 180


def token_age(saved: str) -> Optional[int]:
    try:
        return (date.today() - date.fromisoformat(saved)).days
    except Exception:
        return None


def token_reminders(cfg: Optional[Dict] = None) -> List[List[str]]:
    """Refresh Token 교체 알림(오래됐거나 저장 날짜를 모름) → [등급, 분류, 내용, 조치]."""
    cfg = cfg or load_config()
    out = []
    for key, saved, name in (("refresh_token", "token_saved", "SP-API"), ("ads_refresh_token", "ads_token_saved", "Ads API")):
        if not cfg.get(key):
            continue
        age = token_age(cfg.get(saved, ""))
        if age is None:
            out.append(["P3", "토큰 교체", f"{name} Refresh Token을 언제 받았는지 기록이 없습니다(대화·메일 등에 노출된 적이 있으면 교체 권장)",
                        "⚙ 설정 → 토큰 탭의 ‘교체 방법’대로 새 토큰을 받아 붙여 넣고 저장"])
        elif age >= TOKEN_MAX_DAYS:
            out.append(["P3", "토큰 교체", f"{name} Refresh Token을 받은 지 {age}일 지났습니다", "보안을 위해 새로 승인해 교체를 권장"])
    return out


def preview_patch(c: "Client", seller: str, sku: str, product_type: str, patches: List[Dict]) -> Dict:
    """VALIDATION_PREVIEW: Amazon이 검사만 하고 실제로는 바꾸지 않음."""
    return c.request("PATCH", f"/listings/2021-08-01/items/{seller}/{urllib.parse.quote(sku, safe='')}",
                     {"marketplaceIds": c.mp, "mode": "VALIDATION_PREVIEW", "includedData": "issues",
                      "issueLocale": "en_US"},
                     {"productType": product_type, "patches": patches})


def write_check(c: "Client", seller: str, sku: str) -> List[Tuple]:
    """수정 권한 점검(실제로 바꾸지 않음): 현재 제목·가격과 같은 값으로 VALIDATION_PREVIEW 요청.
    → [(항목, True/False/None, 내용)]"""
    out = []
    if not sku:
        return [("리스팅 수정 권한", None, "점검할 SKU가 없음")]
    d = c.get(f"/listings/2021-08-01/items/{seller}/{urllib.parse.quote(sku, safe='')}", marketplaceIds=c.mp,
              includedData="summaries,attributes,offers")
    pt = ((d.get("summaries") or [{}])[0]).get("productType", "")
    attrs = d.get("attributes") or {}

    def res(name, patches, what):
        try:
            r = preview_patch(c, seller, sku, pt, patches)
            errs = [i for i in r.get("issues", []) if i.get("severity") == "ERROR"]
            if r.get("status") in ("VALID", "ACCEPTED") and not errs:
                out.append((name, True, f"가능 — 현재 {what} 그대로 검사만 해 봄(실제 변경 없음, SKU {sku})"))
            else:
                msg = "; ".join(f"{i.get('code')}: {i.get('message', '')[:120]}" for i in errs[:2])
                out.append((name, None, f"권한은 있음, 검사 결과 {r.get('status')}: {msg}"))
        except Exception as e:
            st = getattr(e, "status", 0)
            out.append((name, False, "권한 없음 — 개발자 앱에 ‘Product Listing’ 역할을 추가하고 다시 승인(Authorize)하세요"
                         if st == 403 else str(e)[:200]))

    if attrs.get("item_name"):
        res("리스팅 수정(제목·불릿 등)", [{"op": "replace", "path": "/attributes/item_name", "value": attrs["item_name"]}], "제목")
    else:
        out.append(("리스팅 수정(제목·불릿 등)", None, f"SKU {sku}에 제목 정보가 없어 점검하지 못함"))
    time.sleep(0.5)
    po = attrs.get("purchasable_offer")
    if not po:
        offers = d.get("offers") or []
        amt = next((o.get("price", {}).get("amount") for o in offers if o.get("offerType") == "B2C"), None)
        if amt is not None:
            po = [{"marketplace_id": c.mp, "currency": "USD", "our_price": [{"schedule": [{"value_with_tax": float(amt)}]}]}]
    if po:
        res("가격 수정", [{"op": "replace", "path": "/attributes/purchasable_offer", "value": po}], "가격")
    else:
        out.append(("가격 수정", None, f"SKU {sku}의 현재 가격을 찾지 못해 점검하지 못함"))
    return out


def ads_check(cfg: Dict) -> List[Dict]:
    if not cfg.get("ads_refresh_token"):
        return [{"항목": "Ads API", "결과": "⚠", "내용": "아직 승인 전(Refresh Token 없음) — 광고 보고서는 광고 콘솔 화면으로 받습니다"}]
    if not any(cfg.get(k) for k in ("ads_client_id", "ads_refresh_token")):
        return [{"항목": "Ads API", "결과": "⚠", "내용": "설정 없음(입력하지 않으면 광고 콘솔 화면으로 보고서 받기)"}]
    try:
        import ads_api
        prof = ads_api.find_profiles(cfg)
        mine = [p for p in prof if p["profileId"] == str(cfg.get("ads_profile_id"))]
        if not cfg.get("ads_profile_id"):
            return [{"항목": "Ads API", "결과": "⚠", "내용": f"인증 정상 · 프로필 {len(prof)}개 — ‘프로필 찾기’로 미국 판매자 프로필을 고르세요"}]
        if not mine:
            return [{"항목": "Ads API", "결과": "❌", "내용": "인증 정상이지만 프로필 ID가 목록에 없음 — ‘프로필 찾기’로 다시 고르세요"}]
        p = mine[0]
        return [{"항목": "Ads API(광고 보고서)", "결과": "✅", "내용": f"{p['name']} · {p['country']} · {p['type']} 프로필 사용"}]
    except Exception as e:
        msg = str(e)[:220]
        if getattr(e, "status", 0) in (400, 401):
            msg = "인증 실패 — Ads Client ID·Secret·Refresh Token을 확인하거나 다시 승인하세요: " + msg[:150]
        return [{"항목": "Ads API", "결과": "❌", "내용": msg}]


def _sp_check(cfg: Dict) -> List[Dict]:
    out: List[Dict] = []

    def add(name, ok, msg):
        out.append({"항목": name, "결과": "✅" if ok is True else ("⚠" if ok is None else "❌"), "내용": msg})

    def why(e) -> str:
        s = str(e)
        if getattr(e, "status", 0) == 403:
            return "권한 없음 — 개발자 앱에 해당 역할(Role)을 추가하고 다시 승인(Authorize)하세요"
        if getattr(e, "status", 0) == 400 and "seller" in s.lower():
            return "Seller ID(판매자 토큰)를 확인하세요: " + s[:150]
        return s[:200]

    try:
        c = Client(cfg)
        c.token()
        add("인증(LWA)", True, "Client ID·Secret·Refresh Token 정상")
    except Exception as e:
        add("인증(LWA)", False, why(e))
        return out
    try:
        add("판매자 계정", True, c.check())
    except Exception as e:
        add("판매자 계정", False, why(e))
    for name, rtype in (("전체 상품(All Listings)", "GET_MERCHANT_LISTINGS_ALL_DATA"),
                        ("FBA 재고", "GET_FBA_INVENTORY_PLANNING_DATA"),
                        ("주문", "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL"),
                        ("반품", "GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA"),
                        ("비즈니스 보고서", "GET_SALES_AND_TRAFFIC_REPORT"),
                        ("Brand Analytics(SQP·SCP)", "GET_BRAND_ANALYTICS_SEARCH_CATALOG_PERFORMANCE_REPORT"),
                        ("계정 성과", "GET_V2_SELLER_PERFORMANCE_REPORT")):
        try:
            d = c.get("/reports/2021-06-30/reports", reportTypes=rtype, pageSize=1)
            n = len(d.get("reports", []))
            add(name, True, f"보고서 조회 가능({rtype})" + (" · 이전 요청 있음" if n else ""))
        except Exception as e:
            add(name, False, why(e))
        time.sleep(0.6)
    asin = ""
    try:
        d = c.get(f"/listings/2021-08-01/items/{(cfg or c.cfg).get('seller_id', '')}", marketplaceIds=c.mp,
                  includedData="summaries", pageSize=1)
        it = (d.get("items") or [{}])[0]
        asin = ((it.get("summaries") or [{}])[0]).get("asin", "")
        add("리스팅 상세(Listings Items)", True, f"SKU 조회 가능(예: {it.get('sku', '')})")
        for r in write_check(c, (cfg or c.cfg).get("seller_id", ""), it.get("sku", "")):
            add(*r)
    except Exception as e:
        add("리스팅 상세(Listings Items)", False, why(e))
    try:
        c.get("/actionItems/2026-06-01/items", marketplaceIds=c.mp, pageSize=1)
        add("Amazon 할 일(Action Items)", True, "조회 가능")
    except Exception as e:
        add("Amazon 할 일(Action Items)", False, why(e))
    if asin:
        try:
            c.get(f"/customerFeedback/2024-06-01/items/{asin}/reviews/topics", marketplaceId=c.mp, sortBy="MENTIONS")
            add("리뷰 주제(Customer Feedback)", True, "조회 가능")
        except Exception as e:
            add("리뷰 주제(Customer Feedback)", None if getattr(e, "status", 0) == 404 else False,
                "이 ASIN은 리뷰 주제 데이터 없음(정상일 수 있음)" if getattr(e, "status", 0) == 404 else why(e))
    return out

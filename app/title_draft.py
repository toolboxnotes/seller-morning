"""
75자 제목 · Item Highlights(125자) 초안 만들기 — 제안만, 상품에 적용하지 않음

방법
  1. 금지 문자(! $ ? _ { } ^ ¬ ¦)·홍보 문구를 빼고, |  -  ,  ; 기준으로 구절을 나눔
  2. 첫 구절(핵심 상품명)을 중심으로, Brand Analytics 검색어(SQP) 검색량이 큰 단어가 많은 구절을 75자 안에서 덧붙임
     (75자가 넘으면 검색량이 작은 단어부터 뺌, 같은 단어는 2번까지)
  3. 제목에 못 들어간 구절(재질·용도·대상 어종 등)은 Item Highlights 초안(125자)으로
  4. 원래 제목에는 있었는데 초안에서 빠진 인기 검색어를 표시 → 사람이 보고 고침
"""

import re
from collections import Counter, defaultdict
from typing import Dict, Iterable, List, Tuple

TITLE_MAX = 75
HL_MAX = 125
BAD = set("!$?_{}^¬¦")
STOP = {"a", "an", "and", "or", "for", "the", "with", "in", "on", "of", "to", "by", "at", "from", "as", "&", "is", "your", "you"}
PROMO = re.compile(r"\b(free shipping|best seller|bestseller|hot sale|on sale|cheap|100% quality|quality guaranteed|"
                   r"satisfaction guaranteed|money back|guaranteed|#1|top rated|new arrival|limited time|discount|lowest price|"
                   r"best price|buy now)\b", re.I)


def words(t: str) -> List[str]:
    return re.findall(r"[a-z0-9]+(?:['-][a-z0-9]+)*", (t or "").lower())


def weights(sqp: Iterable[Tuple[str, float]]) -> Dict[str, float]:
    w: Dict[str, float] = defaultdict(float)
    for q, vol in sqp:
        for tok in set(words(q)) - STOP:
            w[tok] += float(vol or 0)
    return w


def _clean(t: str) -> str:
    t = "".join(" " if c in BAD else c for c in t or "")
    t = PROMO.sub(" ", t)
    return re.sub(r"\s+", " ", t).strip(" ,;|-–—")


def _segments(t: str) -> List[str]:
    parts = re.split(r"\s*[|｜;/]\s*|\s+[-–—]\s+|,\s+|\(\s*|\s*\)", t)
    return [p.strip(" ,;|-–—") for p in parts if p and p.strip(" ,;|-–—")]


def _limit_repeats(text: str, brand: str) -> str:
    seen = Counter()
    out = []
    for tok in text.split(" "):
        k = re.sub(r"[^a-z0-9]", "", tok.lower())
        if k and k not in STOP:
            seen[k] += 1
            if seen[k] > 2:
                continue
        out.append(tok)
    return re.sub(r"\s+", " ", " ".join(out)).strip(" ,")


def _fit(text: str, w: Dict[str, float], keep: set, limit: int) -> str:
    """limit자를 넘으면 검색량이 작은 단어부터(브랜드·숫자 단위 제외) 뺌."""
    toks = text.split(" ")
    while len(" ".join(toks)) > limit and len(toks) > 1:
        cand = [(w.get(re.sub(r"[^a-z0-9]", "", x.lower()), 0), -i, i) for i, x in enumerate(toks)
                if re.sub(r"[^a-z0-9]", "", x.lower()) not in keep and i > 0]
        if not cand:
            toks = toks[:-1]
            continue
        cand.sort()
        del toks[cand[0][2]]
    return " ".join(toks).strip(" ,")


def draft(title: str, brand: str, w: Dict[str, float]) -> Tuple[str, str]:
    t = _clean(title)
    segs = _segments(t)
    if not segs:
        return t[:TITLE_MAX], ""
    core = segs[0]
    keep = {brand.lower()} if brand else set()
    if brand and t.lower().startswith(brand.lower()) and not core.lower().startswith(brand.lower()):
        core = f"{brand} {core}"
    used = [0]
    out = _fit(_limit_repeats(core, brand), w, keep, TITLE_MAX)
    # 검색량 높은 단어가 많은 구절부터 덧붙임
    rest = sorted(range(1, len(segs)), key=lambda i: -sum(w.get(x, 0) for x in set(words(segs[i])) - set(words(out))))
    for i in rest:
        add = " ".join(x for x in segs[i].split(" ") if re.sub(r"[^a-z0-9]", "", x.lower()) not in set(words(out)) - STOP)
        if not add.strip():
            used.append(i)
            continue
        cand = _limit_repeats(f"{out}, {add}", brand)
        if len(cand) <= TITLE_MAX:
            out = cand
            used.append(i)
    hl_parts = [segs[i] for i in range(1, len(segs)) if i not in used]
    hl = ""
    for p in hl_parts:
        c = f"{hl}, {p}" if hl else p
        if len(c) <= HL_MAX:
            hl = c
    return out.strip(" ,"), hl


def lost_keywords(title: str, new: str, sqp: List[Tuple[str, float]], top: int = 5) -> List[str]:
    ow, nw = set(words(title)), set(words(new))
    lost = []
    for q, vol in sorted(sqp, key=lambda x: -float(x[1] or 0)):
        qt = set(words(q)) - STOP
        if qt and qt <= ow and not qt <= nw:
            lost.append(q)
        if len(lost) >= top:
            break
    return lost


def build(rows: List[Dict], brand: str, sqp: List[Tuple[str, float]]) -> List[Dict]:
    """rows: [{"SKU","ASIN","상태","현재 제목"}] (75자 초과만) → 초안 표."""
    w = weights(sqp)
    cache: Dict[str, Tuple[str, str]] = {}
    out = []
    for r in rows:
        t = str(r.get("현재 제목") or "").strip()
        if len(t) <= TITLE_MAX:
            continue
        if t not in cache:
            cache[t] = draft(t, brand, w)
        d, hl = cache[t]
        out.append({"SKU": r.get("SKU", ""), "ASIN": r.get("ASIN", ""), "상태": r.get("상태", ""),
                    "현재 글자 수": len(t), "제목 초안(75자 이내)": d, "초안 글자 수": len(d),
                    "Item Highlights 초안(125자 이내)": hl, "초안에서 빠진 인기 검색어": ", ".join(lost_keywords(t, d, sqp)),
                    "현재 제목": t, "확인": "검토 후 Seller Central에서 직접 수정(자동 적용 안 함)"})
    return out

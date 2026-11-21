"""
계정 상태 · 이슈 상세

계정 상태 페이지에서 정책 항목 건수 > 0일 때:
  1. 그 항목을 눌러 이슈 목록 읽기(ASIN / SKU / 상품명 / 사유 / 날짜 / 상태)
  2. 계정 상태 페이지로 돌아와 다음 항목
  3. 영향받은 ASIN은 Amazon.com 상품 페이지를 열어 기본 정보 읽기

페이지 구조가 일정하지 않아 ‘표 먼저, 안 되면 글자’로 추출하고, 페이지마다
원문 글자와 스크린샷을 저장해 확인·규칙 조정에 씀.
"""

import re
from pathlib import Path
from typing import Callable, Dict, List, Optional

import human

PRODUCT_URL = "https://www.amazon.com/dp/{asin}"

ASIN_RE = re.compile(r"\b(B0[A-Z0-9]{8})\b")
DATE_RE = re.compile(
    r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.? \d{1,2}, \d{4}\b"
    r"|\b\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\b|\b\d{1,2}/\d{1,2}/\d{4}\b"
    r"|\d{4}년 ?\d{1,2}월 ?\d{1,2}일|\d{4}年\d{1,2}月\d{1,2}日"
)

# 머리글 키워드 → 필드(영 / 중 / 한)
HEADER_FIELDS = [
    ("asin", ["asin"]),
    ("sku", ["sku"]),
    ("title", ["product name", "title", "product", "商品名称", "商品", "상품명", "제품명", "상품", "제품"]),
    ("reason", ["reason", "policy", "issue", "violation", "原因", "政策", "问题", "사유", "정책", "문제", "위반"]),
    ("date", ["date", "日期", "时间", "날짜", "일자"]),
    ("status", ["status", "action", "impact", "状态", "影响", "措施", "상태", "조치", "영향"]),
]

# 페이지(Shadow DOM 포함)의 모든 표 줄 수집
_ROWS_JS = """
() => {
  const out = [];
  const seen = new Set();
  const rowSel = 'tr, [role="row"], kat-table-row';
  const cellSel = 'td, th, [role="cell"], [role="gridcell"], [role="columnheader"], kat-table-cell, kat-table-head-cell';
  const walk = (root) => {
    root.querySelectorAll(rowSel).forEach(r => {
      if (seen.has(r)) return;
      seen.add(r);
      const cells = [...r.querySelectorAll(cellSel)]
        .map(c => (c.innerText || '').trim().replace(/\\s+/g, ' '))
        .filter(Boolean);
      if (cells.length >= 2) out.push(cells);
    });
    root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) walk(el.shadowRoot); });
  };
  walk(document);
  return out.slice(0, 500);
}
"""

# Amazon.com 상품 페이지 기본 정보
_PRODUCT_JS = """
() => {
  const t = (sel) => { const e = document.querySelector(sel); return e ? (e.innerText || e.textContent || '').trim().replace(/\\s+/g, ' ') : ''; };
  const attr = (sel, a) => { const e = document.querySelector(sel); return e ? (e.getAttribute(a) || '') : ''; };
  let bsr = '';
  const detail = document.querySelector('#detailBulletsWrapper_feature_div, #productDetails_feature_div, #prodDetails');
  if (detail) {
    const m = detail.innerText.match(/Best Sellers Rank[:\\s]*([^\\n]+(?:\\n\\s*#[^\\n]+)*)/i);
    if (m) bsr = m[1].replace(/\\s+/g, ' ').trim();
  }
  return {
    title: t('#productTitle'),
    brand: t('#bylineInfo'),
    price: t('#corePrice_feature_div .a-offscreen') || t('#apex_desktop .a-price .a-offscreen') || t('.a-price .a-offscreen'),
    rating: attr('#acrPopover', 'title') || t('#acrPopover .a-icon-alt'),
    reviews: t('#acrCustomerReviewText'),
    availability: t('#availability'),
    bsr: bsr,
    image: attr('#landingImage', 'src'),
    page_title: document.title,
  };
}
"""


# ================================================================ 도구

def _field_for_header(h: str) -> Optional[str]:
    low = h.lower()
    for field, words in HEADER_FIELDS:
        if any(w in low for w in words):
            return field
    return None


def _open_link(page, locator):
    """사람처럼 링크 클릭. 새 탭이 열리면(늦게 열릴 수 있음) 새 페이지를, 아니면 원래 페이지를 돌려줌."""
    ctx = page.context
    try:
        with ctx.expect_page(timeout=8000) as pi:
            human.click(page, locator, wait_after=False)
        new = pi.value
        new.wait_for_load_state("domcontentloaded")
        return new, True
    except Exception:
        pass
    try:
        page.wait_for_load_state("domcontentloaded", timeout=20000)
    except Exception:
        pass
    return page, False


def _wait_settled(page) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    human.pause()


def _find_category_link(page, labels: List[str]):
    for lab in labels:
        loc = page.get_by_role("link", name=lab, exact=True)
        if loc.count() and loc.first.is_visible():
            return loc.first
    for lab in labels:  # <a>가 아닌 링크가 있어 글자로 찾음
        loc = page.get_by_text(lab, exact=True)
        if loc.count() and loc.first.is_visible():
            return loc.first
    return None


# ================================================================ 이슈 추출

# 정책 준수 이슈 페이지(Manage compliance)에서 이슈 하나는 이런 글자 묶음:
#   사유 / Review details / 날짜 / 상품명 / ASIN xxx / SKU: xxx /
#   판매 위험 / 취해진 조치(1~2줄) / 등급 영향 / View listing / Status: xxx
ASIN_LINE = re.compile(r"^ASIN[:：]?\s*(B0[A-Z0-9]{8})$")
SKU_LINE = re.compile(r"^SKU[:：]?\s*(.+)$")
STATUS_PREFIXES = ("status:", "status: ", "상태:", "상태 :", "状态：", "状态:")
BUTTON_TEXTS = {"review details", "세부 정보 검토", "세부 정보 보기", "查看详情", "view listing", "리스팅 보기",
                "查看商品", "submit", "제출", "提交"}
END_TEXTS = {"open", "archived", "열기", "보관됨", "未结", "已存档"}
REVIEW_RE = re.compile(r"^\W*(Review details|세부 정보 검토|세부 정보 보기|查看详情)\W*$", re.I)


def _lines(text: str) -> List[str]:
    return [l.strip() for l in text.splitlines() if l.strip()]


def _is_button_text(line: str) -> bool:
    """버튼 글자(앞뒤에 아이콘 기호가 붙을 수 있음, 예: “✦Review details”)."""
    return re.sub(r"^\W+|\W+$", "", line).lower() in BUTTON_TEXTS


def parse_issue_blocks(text: str) -> List[Dict]:
    lines = _lines(text)
    anchors = [i for i, l in enumerate(lines) if ASIN_LINE.match(l)]
    if not anchors:
        return []

    # 먼저 이슈마다 시작점(사유 줄)을 정함
    blocks = []
    for i in anchors:
        di = next((j for j in range(i - 1, max(-1, i - 6), -1) if DATE_RE.fullmatch(lines[j])), None)
        ri = None
        if di is not None:
            ri = next((j for j in range(di - 1, max(-1, di - 5), -1)
                       if not _is_button_text(lines[j])), None)
        blocks.append((ri if ri is not None else (di if di is not None else i - 1), di, i))

    issues = []
    for n, (ri, di, i) in enumerate(blocks):
        limit = blocks[n + 1][0] if n + 1 < len(blocks) else len(lines)
        asin = ASIN_LINE.match(lines[i]).group(1)
        reason = lines[ri] if di is not None and ri is not None and ri != di else ""
        date = lines[di] if di is not None else ""
        title = " ".join(lines[di + 1:i]) if di is not None else lines[i - 1]

        k = i + 1
        sku = ""
        if k < limit and SKU_LINE.match(lines[k]):
            sku = SKU_LINE.match(lines[k]).group(1).strip()
            k += 1

        seg, status = [], ""
        while k < limit:
            line = lines[k]
            low = line.lower()
            if low.startswith(STATUS_PREFIXES):
                status = re.split(r"[:：]", line, 1)[1].strip()
                break
            if low in END_TEXTS or line.startswith("©") or low.startswith("site map"):
                break
            if not _is_button_text(line):
                seg.append(line)
            k += 1

        at_risk = seg[0] if seg else ""
        middle = seg[1:]
        ahr_impact = middle[-1] if middle else ""
        action = "; ".join(middle[:-1]) if len(middle) > 1 else ""

        issues.append({
            "source": "block", "asin": asin, "sku": sku, "title": title, "reason": reason, "date": date,
            "at_risk": at_risk, "action": action, "ahr_impact": ahr_impact, "status": status,
            "context": lines[ri:k + 1],
        })
    return issues


def extract_issues(rows: List[List[str]], text: str) -> List[Dict]:
    """차례로 시도: 이슈 묶음(정책 준수 페이지) → 머리글 있는 표 → ASIN 앞뒤 글자."""
    issues = parse_issue_blocks(text)
    if issues:
        return _dedupe(issues)

    headers: List[str] = []
    for cells in rows:
        has_asin = any(ASIN_RE.search(c) for c in cells)
        if not has_asin:
            mapped = [_field_for_header(c) for c in cells]
            if sum(1 for m in mapped if m) >= 2 and all(len(c) < 60 for c in cells):
                headers = cells
            continue
        if not headers:
            continue  # 머리글 없는 줄(예: 달력)은 제외
        issue = {"source": "table", "cells": cells}
        if len(headers) == len(cells):
            for h, c in zip(headers, cells):
                f = _field_for_header(h)
                if f and f not in issue:
                    issue[f] = c
        joined = " | ".join(cells)
        m = ASIN_RE.search(joined)
        issue["asin"] = m.group(1) if m else issue.get("asin")
        if "date" not in issue:
            d = DATE_RE.search(joined)
            if d:
                issue["date"] = d.group(0)
        issues.append(issue)
    if issues:
        return _dedupe(issues)

    lines = _lines(text)
    for i, line in enumerate(lines):
        for m in ASIN_RE.finditer(line):
            ctx = lines[max(0, i - 4): i + 8]
            d = DATE_RE.search(" | ".join(ctx))
            issues.append({"source": "text", "asin": m.group(1), "date": d.group(0) if d else None,
                           "context": ctx})
    return _dedupe(issues)


def _dedupe(issues: List[Dict]) -> List[Dict]:
    seen, out = set(), []
    for it in issues:
        key = (it.get("asin"), it.get("sku"), it.get("reason"), it.get("date"),
               tuple(it.get("cells") or it.get("context") or []))
        if key not in seen:
            seen.add(key)
            out.append(it)
    return out


# ================================================================ 상세 사유(새 이슈만 한 번)

def _xp(s: str) -> str:
    """글자를 XPath 문자열로(따옴표 처리)."""
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in s.split("'")) + ")"


def _issue_row(page, it: Dict):
    """이 이슈가 있는 줄 찾기: ASIN과 사유를 모두 포함하는 가장 작은 영역."""
    asin_txt = f"ASIN {it['asin']}"
    inner = page.locator(
        f"xpath=//*[contains(normalize-space(.), {_xp(asin_txt)})]"
        f"[not(.//*[contains(normalize-space(.), {_xp(asin_txt)})])]")
    if not inner.count():
        return None
    if it.get("reason"):
        row = inner.first.locator(f"xpath=ancestor::*[contains(normalize-space(.), {_xp(it['reason'][:40])})][1]")
    else:
        row = inner.first.locator("xpath=ancestor::*[self::tr or @role='row' or self::kat-table-row][1]")
    return row.first if row.count() else None


def _new_lines(before: set, page) -> List[str]:
    return [l for l in _lines(page.inner_text("body")) if l not in before]


# 정책 도움말 페이지: 본문이 가장 많은 영역(문단·목록이 가장 많은 요소)을 찾고 왼쪽 목차는 피함
_ARTICLE_JS = """
() => {
  let best = null, bestScore = 0;
  for (const el of document.querySelectorAll('main, article, section, div')) {
    let score = 0;
    for (const c of el.children) {
      if (/^(P|UL|OL|H2|H3|H4|TABLE|DL)$/.test(c.tagName)) score += (c.innerText || '').length;
    }
    if (score > bestScore) { bestScore = score; best = el; }
  }
  const h1 = document.querySelector('h1');
  return { title: h1 ? h1.innerText.trim() : document.title,
           text: best ? best.innerText.trim() : '' };
}
"""


def _read_policy_page(target) -> Dict:
    _wait_settled(target)
    human.read_page(target, max_screens=4)
    art = target.evaluate(_ARTICLE_JS)
    text = re.sub(r"\n{3,}", "\n\n", art.get("text") or "").strip()
    return {"policy_title": art.get("title") or "", "policy_url": target.url, "policy_text": text[:4000]}


def collect_reason_detail(page, it: Dict, list_lines: List[str], out_file: Path,
                          policy_cache: Optional[Dict[str, Dict]] = None) -> Dict:
    """
    이슈 목록 페이지에서 이슈 하나에 대해:
      1. “Review details”를 눌러 펼쳐진 뒤 새로 나온 글자 기록
      2. 사유 링크를 눌러 정책 설명 페이지 읽기(같은 사유는 한 번만), 그다음 목록으로 돌아옴
    """
    policy_cache = policy_cache if policy_cache is not None else {}
    out = {"review_details": "", "detail_text": "", "detail_file": None,
           "policy_title": "", "policy_url": "", "policy_text": ""}
    row = _issue_row(page, it)
    if row is None:
        out["detail_text"] = "(페이지에서 이 이슈의 위치를 찾지 못했습니다)"
        return out
    log = [f"ASIN {it.get('asin')}  SKU {it.get('sku')}", f"사유: {it.get('reason')}", ""]

    # ---- 1. Review details(버튼 글자가 컴포넌트 안에 있을 수 있어 역할·글자 둘 다로 찾음)
    btn = row.get_by_role("button", name=REVIEW_RE).or_(row.get_by_text(REVIEW_RE))
    btn = next((btn.nth(i) for i in range(min(btn.count(), 5)) if btn.nth(i).is_visible()), None)
    if btn is not None:
        before = set(_lines(page.inner_text("body")))
        human.click(page, btn)
        page.wait_for_timeout(2000)
        added = _new_lines(before, page)
        out["review_details"] = "\n".join(added[:40])
        log += ["=== Review details 펼친 후 추가된 텍스트 ===", out["review_details"] or "(추가된 텍스트 없음)", ""]
    else:
        log += ["=== Review details 버튼을 찾지 못했습니다 ===", ""]

    # ---- 2. 사유 링크 → 정책 설명 페이지
    reason = it.get("reason") or ""
    cached = policy_cache.get(reason)
    if cached:
        out.update(cached)
        log += [f"=== 정책 설명(앞과 같은 사유, 이미 읽음){cached['policy_url']} ===", cached["policy_text"]]
    elif reason:
        cand = row.locator("a").filter(has_text=reason[:30])
        link = cand.first if cand.count() else None
        if link is not None:
            start_url = page.url
            before = set(_lines(page.inner_text("body")))
            target, is_new = _open_link(page, link)
            if is_new or target.url != start_url:
                info = _read_policy_page(target)
                out.update(info)
                policy_cache[reason] = info
                log += [f"=== 정책 설명 {info['policy_title']} {info['policy_url']} ===", info["policy_text"]]
                if is_new:
                    target.close()
                    human.pause()
                else:
                    target.go_back(wait_until="domcontentloaded")
                    _wait_settled(page)
            else:
                page.wait_for_timeout(1500)
                added = _new_lines(before, page)
                out["detail_text"] = "\n".join(added[:80])
                log += ["=== 사유 클릭 후 추가된 텍스트(팝업) ===", out["detail_text"] or "(추가된 텍스트 없음)"]
                if added:
                    page.keyboard.press("Escape")
                    human.pause()

    out_file.write_text("\n".join(log), encoding="utf-8")
    out["detail_file"] = out_file.name
    out["detail_ok"] = bool(out["review_details"] or len(out["policy_text"]) > 50 or len(out["detail_text"]) > 50)
    return out


# ================================================================ 메인 흐름

def collect_issue_details(page, policies: List[Dict], label_map: Dict[str, List[str]],
                          out_prefix: Path, is_known: Optional[Callable[[str, Dict], bool]] = None) -> List[Dict]:
    """
    건수 > 0인 정책 항목을 하나씩 열어 이슈 추출, 새 이슈는 상세 사유도 확인.
    돌려줌 [{key, name, count, url, issues:[...], text_file, screenshot, error}]
    """
    results = []
    for pol in policies:
        if not pol.get("count"):
            continue
        key, name = pol["key"], pol["name"]
        print(f"   확인: “{name}” 이슈 {pol['count']}건...")
        entry = {"key": key, "name": name, "count": pol["count"], "url": None, "issues": [],
                 "text_file": None, "screenshot": None, "error": None}
        try:
            link = _find_category_link(page, label_map[key])
            if link is None:
                entry["error"] = "계정 상태 페이지에서 이 분류의 링크를 찾지 못했습니다"
                results.append(entry)
                continue

            start_url = page.url
            target, is_new = _open_link(page, link)
            _wait_settled(target)
            human.read_page(target)

            text = target.inner_text("body")
            rows = target.evaluate(_ROWS_JS)
            entry["url"] = target.url
            entry["issues"] = extract_issues(rows, text)

            txt = Path(f"{out_prefix}_{key}_page.txt")
            png = Path(f"{out_prefix}_{key}.png")
            txt.write_text(f"URL: {target.url}\n\n{text}\n\n=== 표 행 ===\n"
                           + "\n".join(" | ".join(r) for r in rows), encoding="utf-8")
            target.screenshot(path=str(png), full_page=True)
            entry["text_file"], entry["screenshot"] = txt.name, png.name

            # 새 이슈(또는 전에 상세 사유를 못 읽은 이슈): 상세 사유 확인, 이미 읽은 것은 다시 보지 않음
            list_lines = _lines(text)
            new_items = [it for it in entry["issues"] if it.get("asin") and not (is_known and is_known(key, it))]
            policy_cache: Dict[str, Dict] = {}
            for n, it in enumerate(new_items, start=1):
                print(f"     이슈 {n}/{len(new_items)} 상세 사유 확인({it['asin']})...")
                try:
                    it.update(collect_reason_detail(
                        target, it, list_lines, Path(f"{out_prefix}_{key}_{it['asin']}_{n}_detail.txt"),
                        policy_cache))
                except Exception as e:
                    it["detail_text"] = f"(상세 사유를 읽지 못했습니다: {type(e).__name__}: {str(e).splitlines()[0][:150]})"

            # 계정 상태 페이지로 돌아감
            if is_new:
                target.close()
                human.pause()
            elif page.url != start_url:
                page.goto(start_url, wait_until="domcontentloaded")  # 중간에 상세 페이지를 봤을 수 있어 계정 상태 페이지로 바로 돌아감
                _wait_settled(page)
        except Exception as e:  # 한 항목이 실패해도 다른 항목은 계속
            entry["error"] = f"{type(e).__name__}: {str(e).splitlines()[0][:200]}"
        print(f"     기록 {len(entry['issues'])}건 발견" + (f"({entry['error']})" if entry["error"] else ""))
        results.append(entry)
    return results


def collect_product_info(page, asins: List[str], max_products: int = 15) -> List[Dict]:
    """새 탭에서 Amazon.com 상품 페이지를 열어 기본 정보 읽기."""
    products = []
    if not asins:
        return products
    tab = page.context.new_page()
    try:
        for asin in asins[:max_products]:
            url = PRODUCT_URL.format(asin=asin)
            print(f"   상품 확인: {asin} ...")
            info = {"asin": asin, "url": url, "error": None}
            try:
                human.goto(tab, url, wait_after=False)
                _wait_settled(tab)
                human.scroll_by(tab, 600)  # 조금 내려서 봄(상세 정보는 아래쪽)
                human.pause()
                info.update(tab.evaluate(_PRODUCT_JS))
                # “Visit the XXX Store” / “Brand: XXX” → XXX
                b = re.sub(r"^(Visit the |Brand:\s*)", "", info.get("brand") or "")
                info["brand"] = re.sub(r"\s+Store$", "", b).strip()
                if not info.get("title"):
                    info["error"] = "상품 페이지에 제목이 없습니다. 판매 중지·삭제되었거나 페이지가 열리지 않을 수 있습니다"
            except Exception as e:
                info["error"] = f"{type(e).__name__}: {str(e).splitlines()[0][:200]}"
            products.append(info)
    finally:
        tab.close()
    return products

"""
기능마다 실행이 끝나면 요점(주요 숫자 + 처리할 일)을 오늘 요약에 기록: reports/daily/날짜/status.json
“오늘의 요약”은 여기서 만들어짐. 기록에 실패해도 기능 자체에는 영향 없음.
"""

import json
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional

BASE_DIR = Path(__file__).resolve().parent.parent
DAILY_DIR = BASE_DIR / "reports" / "daily"
LEVELS = ("P0", "P1", "P2", "P3")
OLD_LEVEL = {"높음": "P1", "중간": "P2", "낮음": "P3"}


def _file(d: Optional[date] = None) -> Path:
    p = DAILY_DIR / (d or date.today()).isoformat()
    p.mkdir(parents=True, exist_ok=True)
    return p / "status.json"


def load(d: Optional[date] = None) -> Dict:
    f = _file(d)
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return {"modules": {}, "steps": {}}


def _save(data: Dict, d: Optional[date] = None) -> None:
    _file(d).write_text(json.dumps(data, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


def norm_alert(a) -> List[str]:
    """(등급, 분류, 내용, 조치) 형식으로 통일."""
    if isinstance(a, dict):
        lv = a.get("level", "P2")
        lv = OLD_LEVEL.get(lv, lv)
        return [lv if lv in LEVELS else "P2", str(a.get("type") or a.get("what") or ""),
                str(a.get("detail") or ""), str(a.get("advice") or a.get("object") or "")]
    if isinstance(a, str):
        return ["P2", "", a, ""]
    a = list(a) + [""] * 4
    lv = OLD_LEVEL.get(a[0], a[0])
    return [lv if lv in LEVELS else "P2", str(a[1]), str(a[2]), str(a[3])]


def record(module: str, kpis: Optional[Dict] = None, alerts: Optional[Iterable] = None,
           report: Optional[Path] = None, extra_alerts: Optional[Iterable] = None) -> None:
    try:
        data = load()
        m = data["modules"].setdefault(module, {})
        m["time"] = datetime.now().isoformat(timespec="seconds")
        if kpis is not None:
            m["kpis"] = {str(k): (v if isinstance(v, (int, float, str)) else str(v)) for k, v in kpis.items()}
        al = [norm_alert(a) for a in (alerts or [])] + [norm_alert(a) for a in (extra_alerts or [])]
        if alerts is not None or extra_alerts is not None:
            m["alerts"] = al
        if report:
            m["report"] = str(report)
        _save(data)
    except Exception:
        pass


def add_alerts(module: str, alerts: Iterable) -> None:
    """이미 기록된 기능에 항목을 더 붙임(예: 페이지 읽기 결과)."""
    try:
        data = load()
        m = data["modules"].setdefault(module, {})
        m.setdefault("alerts", [])
        m["alerts"] += [norm_alert(a) for a in alerts]
        _save(data)
    except Exception:
        pass


def step(key: str, name: str, status: str, minutes: float, report: Optional[str] = None, error: str = "") -> None:
    try:
        data = load()
        data["steps"][key] = {"name": name, "status": status, "minutes": round(minutes, 1), "report": report or "",
                              "error": error, "time": datetime.now().isoformat(timespec="seconds")}
        _save(data)
    except Exception:
        pass


def merge_pages(module: str, results) -> None:
    """페이지 읽기 결과(sc_pages)의 숫자와 항목을 어떤 기능 아래에 더함."""
    try:
        data = load()
        m = data["modules"].setdefault(module, {})
        m.setdefault("kpis", {})
        m.setdefault("alerts", [])
        for r in results or []:
            if not r:
                continue
            for k, v in (r.get("kpis") or {}).items():
                if not isinstance(v, (list, dict)):
                    m["kpis"][f"{r.get('name', '')}: {k}"] = v
            m["alerts"] += [norm_alert(a) for a in r.get("alerts", [])]
        _save(data)
    except Exception:
        pass


def points(pts, advice: str = "") -> List[List[str]]:
    """보고서 요점 (색, 글자): 빨강 → P1, 주황 → P2, 나머지 → P3."""
    out = []
    for p in pts or []:
        color, text = (p if isinstance(p, (tuple, list)) and len(p) == 2 else (None, p))
        lv = {"C01C1C": "P1", "C26A00": "P2"}.get(color or "", "P3")
        out.append([lv, "요점", str(text), advice if lv != "P3" else ""])
    return out


def source(label: str, method: str, detail: str = "") -> None:
    """데이터를 어디서 가져왔는지(API / 브라우저) 기록 — 오늘의 요약 ‘데이터 출처’."""
    try:
        data = load()
        data.setdefault("sources", []).append({"time": datetime.now().isoformat(timespec="seconds"), "label": label,
                                               "method": method, "detail": detail[:300]})
        data["sources"] = data["sources"][-300:]
        _save(data)
    except Exception:
        pass

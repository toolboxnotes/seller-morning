"""
오래된 디버그·임시 파일 정리(하루 한 번 자동)

reports 폴더는 OneDrive에 있어서, 스크린샷·페이지 구조 파일이 쌓이면 동기화가 무거워지고 화면이 느려집니다.
  - debug 폴더(스크린샷·페이지 구조): 14일 지난 것 삭제
  - 실행 로그(reports/logs), API 원본(reports/api_raw), 다운로드 원본(raw): 30일 지난 것 삭제
  - 메뉴 스캔·광고 탐색 결과: 최근 2번만 남김
보고서(Excel)·대장·설정·이력은 지우지 않습니다.
"""

import os
import shutil
import time
from datetime import date
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
REPORTS = BASE / "reports"
DEBUG_DAYS = 14
RAW_DAYS = 30
KEEP_SCANS = 2
MARK = REPORTS / ".housekeeping_date"


def _old(p: Path, days: int) -> bool:
    try:
        return time.time() - p.stat().st_mtime > days * 86400
    except OSError:
        return False


def _clean_dir(d: Path, days: int) -> tuple:
    n, size = 0, 0
    if not d.exists():
        return n, size
    for root, _, files in os.walk(d):
        for f in files:
            p = Path(root) / f
            if _old(p, days):
                try:
                    size += p.stat().st_size
                    p.unlink()
                    n += 1
                except OSError:
                    pass
    for root, dirs, files in os.walk(d, topdown=False):     # 빈 하위 폴더 정리
        for x in dirs:
            try:
                (Path(root) / x).rmdir()
            except OSError:
                pass
    return n, size


def run(verbose: bool = True) -> str:
    if not REPORTS.exists():
        return ""
    n, size = 0, 0
    for d in REPORTS.rglob("debug"):
        if d.is_dir():
            a, b = _clean_dir(d, DEBUG_DAYS)
            n, size = n + a, size + b
    for d in [REPORTS / "logs", REPORTS / "api_raw", REPORTS / "daily" / "debug"] + [p for p in REPORTS.rglob("raw") if p.is_dir()]:
        a, b = _clean_dir(d, RAW_DAYS)
        n, size = n + a, size + b
    for sub in ("scan", "explore"):
        root = REPORTS / sub
        if root.exists():
            runs = sorted([p for p in root.iterdir() if p.is_dir() and p.name[:8].isdigit()], key=lambda p: p.name)
            for p in runs[:-KEEP_SCANS]:
                try:
                    size += sum(f.stat().st_size for f in p.rglob("*") if f.is_file())
                    n += sum(1 for f in p.rglob("*") if f.is_file())
                    shutil.rmtree(p)
                except OSError:
                    pass
    msg = f"오래된 디버그·임시 파일 {n}개({size / 1024 / 1024:,.0f} MB) 정리" if n else ""
    if verbose and msg:
        print(f"   🧹 {msg}")
    return msg


def daily_once() -> None:
    """오늘 아직 정리하지 않았으면 한 번 정리."""
    try:
        today = date.today().isoformat()
        if MARK.exists() and MARK.read_text().strip() == today:
            return
        run()
        REPORTS.mkdir(parents=True, exist_ok=True)
        MARK.write_text(today)
    except Exception:
        pass

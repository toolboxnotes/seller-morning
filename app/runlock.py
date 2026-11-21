"""
동시에 두 작업이 실행되지 않게 하는 잠금(같은 로그인 브라우저를 함께 쓰면 둘 다 실패하기 때문)

  with runlock.hold("주문 보고서", wait_min=180):
      ...
다른 작업이 실행 중이면 끝날 때까지 1분마다 확인하며 기다립니다(최대 wait_min분).
잠금 파일: %LOCALAPPDATA%\\AmazonAIProject\\run.lock  (프로그램이 비정상 종료돼 남은 잠금은 자동으로 무시)
"""

import json
import os
import sys
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import platform_util as _pu
LOCK_DIR = _pu.local_dir()
LOCK = LOCK_DIR / "run.lock"


def _alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform.startswith("win"):
        try:
            import ctypes
            h = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)   # PROCESS_QUERY_LIMITED_INFORMATION
            if not h:
                return False
            code = ctypes.c_ulong()
            ok = ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
            ctypes.windll.kernel32.CloseHandle(h)
            return bool(ok) and code.value == 259                           # STILL_ACTIVE
        except Exception:
            return True
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def current() -> dict:
    try:
        d = json.loads(LOCK.read_text(encoding="utf-8"))
        return d if _alive(int(d.get("pid", 0))) else {}
    except Exception:
        return {}


def _try_take(name: str) -> bool:
    LOCK_DIR.mkdir(parents=True, exist_ok=True)
    cur = current()
    if cur and int(cur.get("pid", 0)) != os.getpid():
        return False
    tmp = LOCK.with_suffix(".tmp")
    tmp.write_text(json.dumps({"pid": os.getpid(), "task": name, "since": datetime.now().isoformat(timespec="seconds")},
                              ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, LOCK)
    time.sleep(0.2)
    return int(current().get("pid", 0)) == os.getpid()


def release() -> None:
    try:
        d = json.loads(LOCK.read_text(encoding="utf-8"))
        if int(d.get("pid", 0)) == os.getpid():
            LOCK.unlink()
    except Exception:
        pass


@contextmanager
def hold(name: str, wait_min: float = 180):
    t0 = time.time()
    told = False
    while not _try_take(name):
        cur = current()
        mins = (time.time() - t0) / 60
        if mins >= wait_min:
            raise RuntimeError(f"다른 작업({cur.get('task', '?')}, {cur.get('since', '')[11:16]} 시작)이 {wait_min:.0f}분 넘게 "
                               f"실행 중이라 이번 실행을 취소합니다")
        if not told or int(mins) % 5 == 0:
            print(f"   ⏳ 다른 작업이 실행 중입니다: {cur.get('task', '?')}({cur.get('since', '')[11:16]} 시작). "
                  f"끝나면 자동으로 시작합니다(기다린 시간 {mins:.0f}분, 최대 {wait_min:.0f}분).", flush=True)
            told = True
        time.sleep(60)
    try:
        yield
    finally:
        release()

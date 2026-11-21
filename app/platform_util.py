"""
Windows · macOS 공통 처리(파일 열기, 비밀값 보관, 알림, 이 컴퓨터 전용 폴더)

  Windows: %LOCALAPPDATA%\\AmazonAIProject · 비밀값은 Windows 계정 암호화(DPAPI) · 작업 스케줄러 · Windows 알림
  macOS  : ~/Library/Application Support/AmazonAIProject · 비밀값은 macOS 키체인 · launchd · 알림 센터
"""

import os
import subprocess
import sys
from pathlib import Path
from typing import Optional

IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
APP = "AmazonAIProject"
KEYCHAIN_SERVICE = "AmazonAIProject"


def local_dir() -> Path:
    """이 컴퓨터에만 두는 폴더(OneDrive 밖): 설정·로그인 프로필·실행 잠금."""
    if IS_WIN:
        return Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local")) / APP
    if IS_MAC:
        return Path.home() / "Library" / "Application Support" / APP
    return Path(os.environ.get("LOCALAPPDATA") or (Path.home() / ".config")) / APP


def open_path(p) -> bool:
    """파일·폴더·주소를 기본 프로그램으로 열기(Excel, Finder/탐색기, 브라우저)."""
    try:
        s = str(p)
        if IS_WIN:
            os.startfile(s)
        elif IS_MAC:
            subprocess.Popen(["open", s])
        else:
            subprocess.Popen(["xdg-open", s])
        return True
    except Exception:
        return False


def show_in_folder(p) -> None:
    p = Path(p)
    if IS_WIN and p.exists() and p.is_file():
        subprocess.Popen(f'explorer /select,"{p}"')
    elif IS_MAC and p.exists():
        subprocess.Popen(["open", "-R", str(p)])
    else:
        open_path(p if p.is_dir() else p.parent)


# ---------------------------------------------------------------- macOS 키체인
def keychain_set(account: str, secret: str) -> bool:
    try:
        import keyring
        keyring.set_password(KEYCHAIN_SERVICE, account, secret)
        return True
    except Exception:
        pass
    try:   # keyring이 없으면 macOS 기본 명령(security)
        r = subprocess.run(["security", "add-generic-password", "-U", "-s", KEYCHAIN_SERVICE, "-a", account, "-w", secret],
                           capture_output=True, timeout=20)
        return r.returncode == 0
    except Exception:
        return False


def keychain_get(account: str) -> Optional[str]:
    try:
        import keyring
        v = keyring.get_password(KEYCHAIN_SERVICE, account)
        if v is not None:
            return v
    except Exception:
        pass
    try:
        r = subprocess.run(["security", "find-generic-password", "-s", KEYCHAIN_SERVICE, "-a", account, "-w"],
                           capture_output=True, text=True, timeout=20)
        return r.stdout.rstrip("\n") if r.returncode == 0 else None
    except Exception:
        return None


# ---------------------------------------------------------------- 알림
def mac_notify(title: str, message: str, path=None) -> bool:
    """macOS 알림 센터. 누르면 열기는 지원되지 않아 결과 파일 위치는 알림 내용에 적음."""
    def esc(t: str) -> str:
        return str(t).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " · ")
    msg = message + (f" · {Path(path).name}" if path else "")
    try:
        r = subprocess.run(["osascript", "-e", f'display notification "{esc(msg)[:300]}" with title "{esc(title)[:100]}" '
                            f'sound name "Glass"'], capture_output=True, timeout=20)
        return r.returncode == 0
    except Exception:
        return False


def no_window_flags() -> int:
    return 0x08000000 if IS_WIN else 0        # CREATE_NO_WINDOW(Windows만)

"""
Windows 알림(오른쪽 아래 토스트) · macOS 알림 센터 — 매일 자동 실행이 끝났을 때, 화면 프로그램이 뒤에 있을 때 기능이 끝났을 때.
알림을 누르면 결과 파일이 열립니다. 설정에서 끌 수 있습니다(⚙ 설정 → 알림).
알림을 띄우지 못해도 기능에는 영향이 없습니다.
"""

import base64
import subprocess
import sys
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

APP_ID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe"   # 등록된 PowerShell 앱 ID


def enabled() -> bool:
    try:
        import sp_api
        return bool(sp_api.load_config().get("notify", True))
    except Exception:
        return True


def toast(title: str, message: str, path: Optional[Path] = None, force: bool = False) -> bool:
    if not (force or enabled()):
        return False
    if sys.platform == "darwin":                 # macOS 알림 센터
        import platform_util as pu
        return pu.mac_notify(title, message, path)
    if not sys.platform.startswith("win"):
        return False
    launch = ""
    if path:
        launch = f' activationType="protocol" launch="{escape(Path(path).resolve().as_uri(), {chr(34): "&quot;"})}"'
    xml = (f'<toast{launch}><visual><binding template="ToastGeneric"><text>{escape(title[:120])}</text>'
           f'<text>{escape(message[:400])}</text></binding></visual></toast>')
    b64 = base64.b64encode(xml.encode("utf-8")).decode()     # 한글이 깨지지 않게
    ps = ("[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;"
          "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null;"
          "$x = New-Object Windows.Data.Xml.Dom.XmlDocument;"
          f"$x.LoadXml([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{b64}')));"
          "$t = New-Object Windows.UI.Notifications.ToastNotification $x;"
          f"[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('{APP_ID}').Show($t)")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                           capture_output=True, timeout=30, creationflags=0x08000000)
        return r.returncode == 0
    except Exception:
        return False


def summary_toast(out: Optional[Path]) -> None:
    """오늘의 요약 결과를 알림으로: P0·P1 건수 + 가장 급한 일 1건."""
    try:
        import daily
        import daily_bus as bus
        alerts, cnt, _ = daily.summary_alerts(bus.load())
        steps = bus.load().get("steps", {})
        fail = [v["name"] for v in steps.values() if v.get("status") == "실패"]
        top = next((a for a, _ in alerts if a[0] in ("P0", "P1")), None)
        msg = f"P0 {cnt['P0']}건 · P1 {cnt['P1']}건 · P2 {cnt['P2']}건"
        if fail:
            msg += f"\n실패: {', '.join(fail)[:80]}"
        if top:
            msg += f"\n[{top[0]}] {top[2][:110]}"
        toast("매일 자동 실행 끝 — 오늘의 요약", msg, out)
    except Exception:
        toast("매일 자동 실행 끝", "오늘의 요약을 확인하세요", out)

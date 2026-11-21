"""
기능 하나를 실행하는 실행기(화면 프로그램이 별도 프로세스로 실행합니다)

    python -u runner.py <기능 키> [--period 2] [--open]

마지막에 다음 줄을 출력합니다(화면 프로그램이 읽음):
    @@OUTPUT@@ <결과 파일 경로>
    @@DONE@@ ok | fail | stopped
"""

import xlsafe  # noqa: F401  Excel 수식 주입 방지(Amazon에서 받은 글자가 수식으로 실행되지 않게)
import argparse
import sys
import time
from pathlib import Path


def main() -> int:
    try:   # 콘솔이 cp949여도 한글·기호가 깨지지 않게
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("key")
    ap.add_argument("--period", default="2")
    ap.add_argument("--open", action="store_true", help="끝나면 결과 파일 열기")
    a = ap.parse_args()

    import tasks
    t = tasks.BY_KEY.get(a.key)
    if t is None or not t.ready:
        print(f"❌ 실행할 수 없는 기능입니다: {a.key}")
        print("@@DONE@@ fail")
        return 2
    import version
    print(f"▶ {t.name}  ({time.strftime('%Y-%m-%d %H:%M:%S')} · 프로그램 {version.text()})\n")
    from playwright.sync_api import Error as PlaywrightError
    from sc_browser import BrowserNotFound, SessionExpired
    ok = False
    try:
        import housekeeping
        import runlock
        housekeeping.daily_once()
        with runlock.hold(t.name):
            out = t.run({"period": a.period}, a.open)
        if out:
            p = Path(str(out))
            if p.exists():
                print(f"@@OUTPUT@@ {p.resolve()}")
        ok = True
    except SessionExpired as e:
        print("❌ 로그인이 만료되었습니다:", e)
        print("   먼저 ‘Seller Central 로그인’을 실행하세요.")
    except BrowserNotFound as e:
        print("❌", e)
    except PlaywrightError as e:
        print("❌ 브라우저 조작에 실패했습니다. 네트워크 연결, 또는 이 프로그램의 브라우저가 이미 열려 있지 않은지 확인하세요.")
        print("   상세 정보:", str(e).splitlines()[0] if str(e) else "")
    except ModuleNotFoundError as e:
        print(f"❌ 필요한 라이브러리가 없습니다: {e.name}. 프로젝트 폴더에서 pip install -r requirements.txt 를 실행하세요.")
    except PermissionError as e:
        print(f"❌ 파일을 저장할 수 없습니다: {e.filename}\n   이 파일이 Excel에서 열려 있다면 닫은 후 다시 실행하세요.")
    except (KeyboardInterrupt, EOFError):
        print("\n중단했습니다.")
    except Exception as e:
        import traceback
        print(f"❌ 예기치 않은 오류: {type(e).__name__}: {e}")
        traceback.print_exc()
    print(f"\n@@DONE@@ {'ok' if ok else 'fail'}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

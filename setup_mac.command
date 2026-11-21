#!/bin/bash
# ================================================================
#  SellerMorning — Mac 처음 설치(한 번만)
#  터미널에서:  bash setup_mac.command   (또는 Finder에서 우클릭 → 열기)
#  1) Python·Chrome 확인  2) 프로그램 패키지 설치(이 Mac 전용 폴더)
#  3) 바탕화면 바로가기  4) Seller Central 로그인  5) 화면 프로그램 실행
# ================================================================
cd "$(dirname "$0")" || exit 1
PROJECT="$(pwd)"
APPDIR="$HOME/Library/Application Support/AmazonAIProject"
VENV="$APPDIR/venv"
pause() { read -r -p "Enter를 누르면 닫힙니다..." _; }

echo "== SellerMorning · Mac 설치 =="
echo "프로젝트 폴더: $PROJECT"

PY="$(command -v python3)"
if [ -z "$PY" ]; then
  echo "❌ Python 3가 없습니다. https://www.python.org/downloads/macos/ 에서 설치한 뒤 다시 실행하세요."
  pause; exit 1
fi
if ! "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)'; then
  echo "❌ Python 3.9 이상이 필요합니다(지금: $("$PY" --version 2>&1)). python.org에서 최신 버전을 설치하세요."
  pause; exit 1
fi
if ! "$PY" -c 'import tkinter' 2>/dev/null; then
  echo "❌ 이 Python에는 화면 기능(tkinter)이 없습니다."
  echo "   python.org 설치본을 쓰거나, Homebrew Python이면:  brew install python-tk"
  pause; exit 1
fi
echo "✅ $("$PY" --version 2>&1)"

if [ ! -d "/Applications/Google Chrome.app" ] && [ ! -d "$HOME/Applications/Google Chrome.app" ]; then
  echo "⚠ Google Chrome이 없습니다. https://www.google.com/chrome/ 에서 설치하세요(Seller Central 화면 기능에 필요)."
fi

mkdir -p "$APPDIR"
if [ ! -x "$VENV/bin/python" ]; then
  echo "▶ 프로그램 전용 Python 환경 만드는 중(OneDrive 밖: $VENV)..."
  "$PY" -m venv "$VENV" || { echo "❌ 환경을 만들지 못했습니다."; pause; exit 1; }
fi
echo "▶ 패키지 설치 중(1~3분)..."
"$VENV/bin/python" -m pip install --upgrade pip -q
if ! "$VENV/bin/python" -m pip install -r requirements.txt -q; then
  echo "❌ 패키지 설치 실패(인터넷 연결 확인 후 다시 실행)"; pause; exit 1
fi
echo "✅ 패키지 설치 완료"

# 바탕화면 바로가기(이 Mac에만 생김 — OneDrive 폴더에 실행 권한이 없어도 동작)
SC="$HOME/Desktop/SellerMorning.command"
cat > "$SC" <<EOS
#!/bin/bash
exec /bin/bash "$PROJECT/start_gui.command"
EOS
chmod +x "$SC"
chmod +x "$PROJECT"/*.command 2>/dev/null
echo "✅ 바탕화면에 ‘SellerMorning’ 바로가기를 만들었습니다."

echo
read -r -p "지금 Seller Central에 로그인할까요? 브라우저가 열리면 직접 로그인·2단계 인증 (Y/n) " A
if [ "$A" != "n" ] && [ "$A" != "N" ]; then
  PYTHONUTF8=1 "$VENV/bin/python" app/sc_browser.py login
fi

echo
echo "✅ 설치가 끝났습니다. 화면 프로그램을 엽니다."
echo "   다음부터는 바탕화면의 ‘SellerMorning’를 더블클릭하세요."
echo "   SP-API 설정은 화면 오른쪽 위 ‘⚙ 설정’에서 다시 입력하세요(비밀값은 Mac 키체인에 저장)."
/bin/bash "$PROJECT/start_gui.command"

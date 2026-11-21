#!/bin/bash
# Mac: 화면 프로그램 실행(더블클릭 또는  bash start_gui.command)
cd "$(dirname "$0")" || exit 1
VENV="$HOME/Library/Application Support/AmazonAIProject/venv"
if [ -x "$VENV/bin/python" ]; then PY="$VENV/bin/python"; else PY="python3"; fi
if ! "$PY" -c 'import playwright, openpyxl' 2>/dev/null; then
  echo "처음 설치가 필요합니다 → bash setup_mac.command 를 먼저 실행하세요."
  read -r -p "Enter를 누르면 닫힙니다..." _; exit 1
fi
export PYTHONUTF8=1 PYTHONIOENCODING=utf-8
nohup "$PY" app/gui.py > /dev/null 2>&1 &
echo "화면 프로그램을 열었습니다. 이 터미널 창은 닫아도 됩니다."

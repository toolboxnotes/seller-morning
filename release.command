#!/bin/bash
# Mac: 버전 패키지 관리(더블클릭 또는  bash release.command)
cd "$(dirname "$0")" || exit 1
VENV="$HOME/Library/Application Support/AmazonAIProject/venv"
if [ -x "$VENV/bin/python" ]; then PY="$VENV/bin/python"; else PY="python3"; fi
export PYTHONUTF8=1 PYTHONIOENCODING=utf-8
echo "============================================================"
echo " 버전 패키지 관리 (releases 폴더)"
echo "============================================================"
echo "  1  지금 버전으로 패키지 만들기"
echo "  2  패키지 목록 보기"
echo "  3  예전 버전으로 되돌리기"
read -r -p "번호: " CH
case "$CH" in
  1) "$PY" app/release.py ;;
  2) "$PY" app/release.py list ;;
  3) "$PY" app/release.py list; read -r -p "되돌릴 버전(예: 1.0.0): " VER; "$PY" app/release.py restore "$VER" ;;
esac
read -r -p "Enter를 누르면 닫힙니다." _

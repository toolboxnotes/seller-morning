#!/bin/bash
# Mac: 콘솔 메뉴 실행(더블클릭 또는  bash start.command)
cd "$(dirname "$0")" || exit 1
VENV="$HOME/Library/Application Support/AmazonAIProject/venv"
if [ -x "$VENV/bin/python" ]; then PY="$VENV/bin/python"; else PY="python3"; fi
export PYTHONUTF8=1 PYTHONIOENCODING=utf-8
"$PY" app/main.py

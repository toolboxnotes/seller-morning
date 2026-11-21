@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
rem 화면 프로그램(창) 실행: 콘솔 창 없이 pythonw로 엽니다.
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" app\gui.py
    exit /b
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw app\gui.py
    exit /b
)
python app\gui.py
pause

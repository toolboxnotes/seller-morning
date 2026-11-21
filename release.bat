@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
title 버전 패키지 관리
if exist ".venv\Scripts\python.exe" (set PY=.venv\Scripts\python.exe) else (set PY=python)

echo ============================================================
echo  버전 패키지 관리 (releases 폴더)
echo ============================================================
echo   1  지금 버전으로 패키지 만들기
echo   2  패키지 목록 보기
echo   3  예전 버전으로 되돌리기
echo.
set /p CH=번호: 
if "%CH%"=="1" %PY% app\release.py
if "%CH%"=="2" %PY% app\release.py list
if "%CH%"=="3" (
    %PY% app\release.py list
    echo.
    set /p VER=되돌릴 버전(예: 1.0.0): 
    call %PY% app\release.py restore %%VER%%
)
echo.
pause

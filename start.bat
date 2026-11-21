@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1

rem ---- Window size: WIN_COLS = width (characters), WIN_LINES = height (lines) ----
set WIN_COLS=140
set WIN_LINES=45

rem Resize the window but keep a long scroll-back buffer (3000 lines),
rem so earlier output can still be scrolled back to.
powershell -NoProfile -Command "try { $u=(Get-Host).UI.RawUI; $b=$u.BufferSize; if ($b.Width -lt %WIN_COLS%) { $b.Width=%WIN_COLS% }; $b.Height=3000; $u.BufferSize=$b; $s=$u.WindowSize; $s.Width=[Math]::Min(%WIN_COLS%, $u.MaxPhysicalWindowSize.Width); $s.Height=[Math]::Min(%WIN_LINES%, $u.MaxPhysicalWindowSize.Height); $u.WindowSize=$s; $b.Width=%WIN_COLS%; $u.BufferSize=$b } catch {}" 2>nul

title Amazon Seller Central Assistant

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app\main.py
) else (
    python app\main.py
)
pause

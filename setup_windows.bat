@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Windows 처음 설치(한 번만): 패키지 설치 → 로그인 → 화면 프로그램
echo == SellerMorning · Windows 설치 ==
where python >nul 2>nul
if errorlevel 1 (
    echo Python이 없습니다. https://www.python.org/downloads/windows/ 에서 설치할 때 "Add python.exe to PATH"를 체크하세요.
    pause
    exit /b
)
python -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)"
if errorlevel 1 (
    echo Python 3.9 이상이 필요합니다.
    pause
    exit /b
)
echo 패키지 설치 중(1~3분)...
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo 패키지 설치 실패. 인터넷 연결을 확인하고 다시 실행하세요.
    pause
    exit /b
)
echo 설치 완료.
set /p A=지금 Seller Central에 로그인할까요? (Y/n) 
if /i not "%A%"=="n" python app\sc_browser.py login
echo 화면 프로그램을 엽니다. 처음 한 번 오른쪽 위 "바탕화면 바로가기"를 누르면 다음부터 바탕화면에서 실행됩니다.
call start_gui.bat

@echo off
chcp 65001 > nul
echo ========================================
echo   RFID 재고조사 시스템 시작
echo ========================================
echo.

:: Python 설치 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 Python 3.10 이상을 설치해주세요.
    pause
    exit /b
)

:: 패키지 설치 확인 및 자동 설치
echo 필요한 패키지를 확인합니다...
pip install -r requirements.txt --quiet
echo.

:: Streamlit 앱 실행
echo 브라우저에서 http://localhost:8501 로 접속하세요.
echo 종료하려면 이 창을 닫거나 Ctrl+C 를 누르세요.
echo.
streamlit run app.py
pause

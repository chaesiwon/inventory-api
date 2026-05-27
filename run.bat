@echo off
chcp 65001 > nul
title 장기재고 소진계획 관리 시스템

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║    장기재고 소진계획 관리 시스템              ║
echo  ╚══════════════════════════════════════════════╝
echo.

:: Python(py) 확인
py --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo 설치 방법:
    echo   1. https://python.org/downloads 접속
    echo   2. Download Python 3.x.x 클릭 후 설치
    echo   3. 설치 시 반드시 [Add Python to PATH] 체크!
    echo   4. 설치 완료 후 이 창 닫고 다시 실행
    pause
    exit /b 1
)

echo [1/4] Python 버전 확인:
py --version
echo.

echo [2/4] 필수 패키지 설치 중... (최초 1회, 약 2-5분 소요)
py -m pip install --upgrade pip --quiet
py -m pip install fastapi uvicorn starlette python-multipart pandas openpyxl xlsxwriter itsdangerous jinja2 aiofiles python-pptx --quiet

if %errorlevel% neq 0 (
    echo.
    echo [오류] 패키지 설치 실패. 아래를 확인하세요:
    echo   - 인터넷 연결 상태
    echo   - 방화벽/프록시 설정
    pause
    exit /b 1
)
echo 패키지 설치 완료 ✓
echo.

echo [3/4] DB 마이그레이션 (기존 데이터 유지)...
py migrate.py
echo.

echo [4/4] 서버 시작...
echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║  브라우저 주소 :  http://localhost:8000      ║
echo  ║  아이디        :  admin                      ║
echo  ║  비밀번호      :  admin1234                  ║
echo  ╠══════════════════════════════════════════════╣
echo  ║  종료 방법 : 이 창에서 Ctrl+C 누르기        ║
echo  ╚══════════════════════════════════════════════╝
echo.
echo 잠시 후 브라우저를 열고 http://localhost:8000 을 입력하세요.
echo.

py -m uvicorn main:app --host 0.0.0.0 --port 8000

pause

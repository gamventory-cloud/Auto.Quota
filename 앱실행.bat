@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title 설문 데이터 툴킷

set "VPY=.venv\Scripts\python.exe"

if not exist "Home.py" goto NOHOME

REM ============================================================
REM  1. 최신 코드 받기 (Git 이 있고, 이 폴더가 저장소일 때만)
REM ============================================================
where git >nul 2>&1
if errorlevel 1 goto SKIPPULL

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 goto SKIPPULL

REM 이 폴더에서 파일을 고친 흔적이 있으면 덮어쓰지 않는다
git diff --quiet
if errorlevel 1 goto DIRTY
git diff --cached --quiet
if errorlevel 1 goto DIRTY

echo [*] 최신 코드를 확인합니다...
git pull --ff-only
if errorlevel 1 goto PULLFAIL
goto SKIPPULL

:DIRTY
echo [!] 이 폴더에서 수정된 파일이 있어 자동 업데이트를 건너뜁니다.
echo     코드를 직접 고치지 않으셨다면 관리자에게 알려 주세요.
echo.
goto SKIPPULL

:PULLFAIL
echo [!] 최신 코드를 받지 못했습니다. 기존 코드로 실행합니다.
echo     인터넷 연결을 확인하거나 관리자에게 문의해 주세요.
echo.

:SKIPPULL

REM ============================================================
REM  2. 가상환경 점검
REM     다른 PC 에서 복사해 온 .venv 는 원래 PC 의 파이썬 경로를
REM     가리켜 동작하지 않는다. 있는지만 보지 말고 실행되는지 본다.
REM ============================================================
if not exist "%VPY%" goto MAKEVENV
"%VPY%" -c "pass" >nul 2>&1
if not errorlevel 1 goto CHECKPKG
echo [*] 가상환경이 이 PC 와 맞지 않습니다. 새로 만듭니다...
echo     (다른 PC 에서 복사해 온 경우 정상입니다)
echo.
rmdir /s /q ".venv" >nul 2>&1

:MAKEVENV
echo [*] 가상환경을 만듭니다...
python -m venv .venv
if errorlevel 1 goto NOPYTHON
if not exist "%VPY%" goto NOPYTHON

REM ============================================================
REM  3. 패키지 점검
REM     streamlit 이 있는지, 그리고 requirements.txt 가 바뀌었는지
REM     둘 다 확인한다. 코드 업데이트로 패키지가 추가됐을 수 있다.
REM ============================================================
:CHECKPKG
if not exist "requirements.txt" goto NOREQ

set "NEWHASH="
for /f "skip=1 delims=" %%H in ('certutil -hashfile "requirements.txt" MD5 2^>nul') do if not defined NEWHASH set "NEWHASH=%%H"

set "OLDHASH="
if exist ".venv\.reqhash" set /p OLDHASH=<".venv\.reqhash"

"%VPY%" -c "import streamlit" >nul 2>&1
if errorlevel 1 goto INSTALL
if not defined NEWHASH goto RUN
if not "%NEWHASH%"=="%OLDHASH%" goto INSTALL
goto RUN

:INSTALL
echo [*] 필요한 패키지를 설치합니다. 몇 분 걸릴 수 있습니다...
echo.
"%VPY%" -m pip install --upgrade pip --quiet
"%VPY%" -m pip install -r requirements.txt
echo.
"%VPY%" -c "import streamlit" >nul 2>&1
if errorlevel 1 goto PIPFAIL
if defined NEWHASH > ".venv\.reqhash" echo %NEWHASH%
echo [*] 설치가 끝났습니다.

REM ============================================================
REM  4. 실행
REM ============================================================
:RUN
echo.
echo [*] 앱을 실행합니다. 브라우저가 자동으로 열립니다.
echo     끄실 때는 이 창에서 Ctrl+C 를 누르거나 창을 닫으세요.
echo.
"%VPY%" -m streamlit run Home.py
pause
exit /b 0

:NOHOME
echo.
echo [X] Home.py 를 찾을 수 없습니다.
echo     이 파일을 Home.py 와 같은 폴더에 두고 실행해 주세요.
echo     지금 위치: %CD%
echo.
pause
exit /b 1

:NOPYTHON
echo.
echo [X] 파이썬을 찾을 수 없습니다.
echo.
echo     python.org 에서 Python 3.12 를 설치하세요.
echo     설치 화면 맨 아래 "Add python.exe to PATH" 체크박스를
echo     반드시 켜야 합니다.
echo.
pause
exit /b 1

:NOREQ
echo.
echo [X] requirements.txt 가 없습니다.
echo     코드를 내려받은 폴더가 맞는지 확인해 주세요.
echo.
pause
exit /b 1

:PIPFAIL
echo.
echo [X] 패키지 설치가 끝나지 않았습니다. 위에 뜬 오류 메시지를
echo     확인해 주세요.
echo.
echo     자주 있는 원인:
echo       - 파이썬 버전이 너무 최신 (3.12 권장)
echo       - 폴더 경로에 한글이나 띄어쓰기가 있음
echo       - 인터넷 연결이 끊김
echo.
echo     아래 명령을 직접 실행하면 원인이 더 자세히 보입니다:
echo       .venv\Scripts\python.exe -m pip install -r requirements.txt
echo.
pause
exit /b 1
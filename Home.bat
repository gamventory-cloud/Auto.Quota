@echo off
cd /d "%~dp0"
echo 필요한 라이브러리를 확인하고 설치합니다...
pip install -r requirements.txt
cls
echo 실행 중입니다...
streamlit run Home.py
pause
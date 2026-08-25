#!/bin/bash
cd "$(dirname "$0")" || exit 1

VPY=".venv/bin/python"

hold() { echo; read -r -p "엔터를 누르면 창이 닫힙니다..."; exit 1; }

if [ ! -f "Home.py" ]; then
    echo
    echo "[X] Home.py 를 찾을 수 없습니다."
    echo "    이 파일을 Home.py 와 같은 폴더에 두고 실행해 주세요."
    echo "    지금 위치: $(pwd)"
    hold
fi

# ── 가상환경 점검
# 다른 PC 에서 복사해 온 .venv 는 원래 PC 의 파이썬 경로를 가리켜 동작하지 않는다.
# 파일이 있는지만 보지 말고 실제로 실행되는지 확인한다.
if [ -e "$VPY" ] && ! "$VPY" -c "pass" >/dev/null 2>&1; then
    echo "[*] 가상환경이 이 PC 와 맞지 않습니다. 새로 만듭니다..."
    echo "    (다른 PC 에서 복사해 온 경우 정상입니다)"
    echo
    rm -rf .venv
fi

if [ ! -x "$VPY" ]; then
    if ! command -v python3 >/dev/null 2>&1; then
        echo
        echo "[X] 파이썬을 찾을 수 없습니다."
        echo "    python.org 에서 Python 3.12 를 설치하거나"
        echo "    터미널에서 'brew install python@3.12' 를 실행하세요."
        hold
    fi
    echo "[*] 가상환경을 만듭니다..."
    python3 -m venv .venv || hold
fi

# ── 폴더만 있고 패키지가 없는 경우까지 잡아낸다
if ! "$VPY" -c "import streamlit" >/dev/null 2>&1; then
    echo "[*] 필요한 패키지를 설치합니다. 처음 한 번만 하며 몇 분 걸립니다..."
    echo
    if [ ! -f "requirements.txt" ]; then
        echo "[X] requirements.txt 가 없습니다."
        echo "    코드를 내려받은 폴더가 맞는지 확인해 주세요."
        hold
    fi
    "$VPY" -m pip install --upgrade pip --quiet
    "$VPY" -m pip install -r requirements.txt
    echo
    if ! "$VPY" -c "import streamlit" >/dev/null 2>&1; then
        echo "[X] 패키지 설치가 끝나지 않았습니다. 위 오류 메시지를 확인해 주세요."
        echo
        echo "    자주 있는 원인:"
        echo "      - 파이썬 버전이 너무 최신 (3.12 권장)"
        echo "      - 폴더 경로에 한글이나 띄어쓰기가 있음"
        echo "      - 인터넷 연결이 끊김"
        echo
        echo "    아래 명령을 직접 실행하면 원인이 더 자세히 보입니다:"
        echo "      .venv/bin/python -m pip install -r requirements.txt"
        hold
    fi
    echo "[*] 설치가 끝났습니다."
fi

echo
echo "[*] 앱을 실행합니다. 브라우저가 자동으로 열립니다."
echo "    끄실 때는 이 창에서 Ctrl+C 를 누르세요."
echo
"$VPY" -m streamlit run Home.py

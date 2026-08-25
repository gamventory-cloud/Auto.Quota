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

# ============================================================
#  1. 최신 코드 받기 (Git 이 있고, 이 폴더가 저장소일 때만)
# ============================================================
if command -v git >/dev/null 2>&1 && git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    if ! git diff --quiet || ! git diff --cached --quiet; then
        echo "[!] 이 폴더에서 수정된 파일이 있어 자동 업데이트를 건너뜁니다."
        echo "    코드를 직접 고치지 않으셨다면 관리자에게 알려 주세요."
        echo
    else
        echo "[*] 최신 코드를 확인합니다..."
        if ! git pull --ff-only; then
            echo "[!] 최신 코드를 받지 못했습니다. 기존 코드로 실행합니다."
            echo "    인터넷 연결을 확인하거나 관리자에게 문의해 주세요."
            echo
        fi
    fi
fi

# ============================================================
#  2. 가상환경 점검
#     다른 PC 에서 복사해 온 .venv 는 원래 PC 의 파이썬 경로를
#     가리켜 동작하지 않는다. 있는지만 보지 말고 실행되는지 본다.
# ============================================================
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

# ============================================================
#  3. 패키지 점검
#     streamlit 이 있는지, requirements.txt 가 바뀌었는지 둘 다 본다.
#     코드 업데이트로 패키지가 추가됐을 수 있다.
# ============================================================
if [ ! -f "requirements.txt" ]; then
    echo "[X] requirements.txt 가 없습니다."
    echo "    코드를 내려받은 폴더가 맞는지 확인해 주세요."
    hold
fi

NEWHASH=$(cksum requirements.txt 2>/dev/null | awk '{print $1"-"$2}')
OLDHASH=""
[ -f ".venv/.reqhash" ] && OLDHASH=$(cat ".venv/.reqhash")

NEED_INSTALL=0
"$VPY" -c "import streamlit" >/dev/null 2>&1 || NEED_INSTALL=1
[ -n "$NEWHASH" ] && [ "$NEWHASH" != "$OLDHASH" ] && NEED_INSTALL=1

if [ "$NEED_INSTALL" = "1" ]; then
    echo "[*] 필요한 패키지를 설치합니다. 몇 분 걸릴 수 있습니다..."
    echo
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
    [ -n "$NEWHASH" ] && echo "$NEWHASH" > ".venv/.reqhash"
    echo "[*] 설치가 끝났습니다."
fi

# ============================================================
#  4. 실행
# ============================================================
echo
echo "[*] 앱을 실행합니다. 브라우저가 자동으로 열립니다."
echo "    끄실 때는 이 창에서 Ctrl+C 를 누르세요."
echo
"$VPY" -m streamlit run Home.py

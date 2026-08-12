"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : check_setup.py                                                  ║
║  위치   : 리포지토리 최상단                                                ║
║  사용법 : python check_setup.py                                           ║
║                                                                          ║
║  깃에 푸시하기 전에 파일 배치가 맞는지 확인합니다.                           ║
║  스트림릿을 띄우지 않고 돌아가므로 배포 전 점검용으로 쓰세요.                 ║
╚══════════════════════════════════════════════════════════════════════════╝
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ok_all = True


def ok(msg):
    print(f"  ✅ {msg}")


def bad(msg, fix=""):
    global ok_all
    ok_all = False
    print(f"  ❌ {msg}")
    if fix:
        print(f"      → {fix}")


def warn(msg):
    print(f"  ⚠️  {msg}")


print("\n[1] 파일 위치")
for f in ("Home.py", "utils.py", "requirements.txt"):
    if os.path.isfile(os.path.join(HERE, f)):
        ok(f"{f} (최상단)")
    else:
        bad(f"{f} 없음", "리포 최상단에 두세요")

pages = os.path.join(HERE, "pages")
if not os.path.isdir(pages):
    bad("pages/ 폴더 없음", "mkdir pages 후 페이지 파일을 옮기세요")
else:
    py = [f for f in os.listdir(pages) if f.endswith(".py")]
    ok(f"pages/ ({len(py)}개 페이지: {', '.join(py) or '없음'})") if py else \
        bad("pages/ 가 비어 있음", "쿼터 솔루션 페이지 파일을 pages/ 로 옮기세요")

if os.path.isfile(os.path.join(HERE, "quota_ilp.py")):
    ok("quota_ilp.py (최상단) — ILP 사용 가능")
else:
    warn("quota_ilp.py 없음 — 그리디 방식만 사용 가능")

print("\n[2] 파일 내용이 섞이지 않았는지")
UI_MARKS = ("st.set_page_config", "st.file_uploader", "매칭 시작")
LIB_MARKS = ("def norm_val", "def simulation_worker", "def check_password")


def read(path):
    try:
        with open(path, encoding="utf-8") as fh:
            return fh.read()
    except Exception:
        return ""


def code_only(src):
    """
    맨 앞의 모듈 docstring을 제거한 소스를 반환한다.
    배너 설명문에 들어 있는 예시 문자열이 오탐을 일으키지 않도록 하기 위함.
    """
    s = src.lstrip()
    for q in ('"""', "'''"):
        if s.startswith(q):
            end = s.find(q, len(q))
            if end != -1:
                return s[end + len(q):]
    return src


u = os.path.join(HERE, "utils.py")
if os.path.isfile(u):
    src = code_only(read(u))
    ui_found = [m for m in UI_MARKS if m in src]
    lib_missing = [m for m in LIB_MARKS if m not in src]
    if ui_found:
        bad(f"utils.py 에 화면 코드가 있습니다: {ui_found}",
            "이 내용은 pages/2___쿼터_솔루션.py 로 옮기고, utils.py 는 공용 모듈로 교체")
    elif lib_missing:
        bad(f"utils.py 에 필수 함수가 없습니다: {lib_missing}",
            "utils.py 를 새 버전으로 교체하세요")
    else:
        ok("utils.py — 공용 모듈 맞음 (화면 코드 없음)")

if os.path.isdir(pages):
    for f in os.listdir(pages):
        if not f.endswith(".py"):
            continue
        src = code_only(read(os.path.join(pages, f)))
        if "def norm_val" in src and "st.file_uploader" not in src:
            bad(f"pages/{f} 는 화면 파일이 아니라 공용 모듈처럼 보입니다",
                "utils.py 내용이 pages/ 에 들어간 것 같습니다")
        elif "import utils" in src:
            ok(f"pages/{f} — 화면 파일 맞음")

print("\n[3] 모듈 임포트")
sys.path.insert(0, HERE)
try:
    import utils as _u
    if getattr(_u, "MODULE_ROLE", None) == "utils":
        ok(f"import utils 성공 (v{getattr(_u, '__version__', '?')})")
    else:
        bad("utils.py 가 임포트는 되지만 내용이 다릅니다",
            "MODULE_ROLE = 'utils' 가 있는 새 버전으로 교체")
except Exception as e:
    bad(f"import utils 실패 — {type(e).__name__}: {e}",
        "utils.py 에 화면 코드가 들어갔을 가능성이 높습니다")

for mod, note in (("quota_ilp", "ILP 솔버"), ("ortools", "ILP 엔진"),
                  ("streamlit", ""), ("pandas", ""), ("numpy", ""),
                  ("altair", ""), ("joblib", ""), ("openpyxl", "엑셀 읽기"),
                  ("xlsxwriter", "엑셀 쓰기"), ("chardet", "CSV 인코딩")):
    try:
        __import__(mod)
        ok(f"{mod}" + (f" ({note})" if note else ""))
    except Exception:
        if mod in ("quota_ilp", "ortools"):
            warn(f"{mod} 없음 — ILP 사용 불가, 그리디만 동작")
        else:
            bad(f"{mod} 없음", "pip install -r requirements.txt")

print("\n[4] 비밀값 관리")
gi = os.path.join(HERE, ".gitignore")
if os.path.isfile(gi):
    g = read(gi)
    if "secrets.toml" in g:
        ok(".gitignore 가 secrets.toml 을 제외함")
    else:
        bad(".gitignore 에 secrets.toml 이 없음",
            ".streamlit/secrets.toml 한 줄을 추가하세요")
    if any(x in g for x in ("*.xlsx", "*.csv")):
        ok(".gitignore 가 데이터 파일을 제외함")
    else:
        warn(".gitignore 에 *.xlsx / *.csv 추가를 권장 (개인정보 유출 방지)")
else:
    bad(".gitignore 없음", "제공된 .gitignore 를 최상단에 두세요")

sec = os.path.join(HERE, ".streamlit", "secrets.toml")
if os.path.isfile(sec):
    ok(".streamlit/secrets.toml 존재 (로컬 실행용 — 커밋되지 않는지 확인)")
else:
    warn(".streamlit/secrets.toml 없음 — 로컬에서는 비밀번호 확인이 실패합니다")

cfg = os.path.join(HERE, ".streamlit", "config.toml")
ok(".streamlit/config.toml 존재") if os.path.isfile(cfg) else \
    warn(".streamlit/config.toml 없음 — 업로드 상한 등 기본값 사용")

print("\n" + "=" * 60)
print("  ✅ 배포 준비 완료" if ok_all else "  ❌ 위 항목을 먼저 수정하세요")
print("=" * 60 + "\n")
sys.exit(0 if ok_all else 1)

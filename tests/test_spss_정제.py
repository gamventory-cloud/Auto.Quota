"""pages/3_🛠️_SPSS_정제.py 의 순수 로직 검증 (streamlit 을 스텁으로 대체).

    python3 tests/test_spss_정제.py
"""
import importlib.util
import sys
from unittest.mock import MagicMock

import pandas as pd

# st.columns(n) 이 n개를 돌려주도록 스텁을 맞춰준다
_stub = MagicMock()
_stub.columns.side_effect = lambda spec=2, **kw: [
    MagicMock() for _ in range(spec if isinstance(spec, int) else len(spec))]
sys.modules["streamlit"] = _stub
ROOT = "/".join(__file__.split("/")[:-2]) or "."
sys.path.insert(0, ROOT)

spec = importlib.util.spec_from_file_location("page", f"{ROOT}/pages/3_🛠️_SPSS_정제.py")
page = importlib.util.module_from_spec(spec)
spec.loader.exec_module(page)

fails = 0


def check(cond, msg):
    global fails
    print(("  OK   " if cond else "  실패 ") + msg)
    if not cond:
        fails += 1


print("[정적 검사]")
try:
    import subprocess

    r = subprocess.run([sys.executable, "-m", "pyflakes",
                        f"{ROOT}/pages/3_🛠️_SPSS_정제.py", f"{ROOT}/excel_style.py"],
                       capture_output=True, text=True)
    # 정의되지 않은 이름은 화면을 눌러봐야 터진다. 배포 전에 여기서 먼저 잡는다.
    undefined = [l for l in r.stdout.splitlines() if "undefined name" in l]
    check(not undefined, "정의되지 않은 이름 없음"
          + ("" if not undefined else f" -> {undefined[:3]}"))
except Exception as exc:
    print(f"  건너뜀 (pyflakes 미설치: pip install pyflakes) [{type(exc).__name__}]")

print("\n[변수명 규칙]")
for cand, fb, want in [("성별", "Q3", "Q3"), ("SQ1", "Q1", "SQ1"), ("1차 구매", "Q7", "Q7"),
                       ("TO", "Q9", "TO_"), ("만족도_1", "Q5_1", "Q5_1")]:
    got = page.make_valid_name(cand, fb)
    check(got == want, f"make_valid_name({cand!r}, {fb!r}) -> {got!r} (기대 {want!r})")

print("\n[매칭]")
df_raw = pd.DataFrame(columns=["ID", "Q1", "Q1_1", "Q1_2", "Q5_1", "Q5_2", "Q9", "Q10"])
df_code = pd.DataFrame([
    ["Q1", "SQ1. 성별"], ["문2", "성별 무관 질문"],
    ["Q5_RK1", "Q5. 선호 순위 (1순위)"], ["Q5_RK2", "Q5. 선호 순위 (2순위)"],
    ["Q9", "만족도"], ["Q10", "만족도"],
])
final, updates, warns = page.analyze(df_raw, df_code, label_col=1)
names = [r["변경할 변수명"] for r in final if r["변경할 변수명"]]
check("ID" not in [r["Raw 변수명"] for r in final], "관리용 열(ID) 제외")
check(not [n for n in set(names) if names.count(n) > 1], "변수명 중복 없음")
check(all(page.is_valid_name(n) for n in names), "모두 SPSS 규칙 통과")
check(any(r["상태"] == "매칭 성공 (순위 문항)" for r in final), "순위 문항(RK) 매칭")

_, _, w2 = page.analyze(pd.DataFrame(columns=["Q1", "q1"]), pd.DataFrame([["Q1", "SQ1. x"]]), 1)
check(bool(w2), "대소문자만 다른 동명 열 경고")

print("\n[신텍스]")
_, cnt0 = page.build_syntax(pd.DataFrame([{"Raw 변수명": "Q1", "변경할 변수명": ""}]), "t")
syn, cnt1 = page.build_syntax(pd.DataFrame([{"Raw 변수명": "Q1", "변경할 변수명": "SQ1"}]), "t")
check(cnt0 == 0, "변환 대상 없으면 RENAME 구문 생략")
check("(Q1 = SQ1)" in syn and cnt1 == 1, "정상 구문 생성")

print("\n[sav]")
df = pd.DataFrame({"q1": [1, 2], "q2": [3, 4]})
edited = pd.DataFrame([
    {"Raw 변수명": "q1", "변경할 변수명": "SQ1", "질문 내용": "성별"},
    {"Raw 변수명": "q2", "변경할 변수명": "SQ2", "질문 내용": "연령"},
])
blob, info = page.build_sav(df, edited)
check(info["vars"] == 2 and blob[:4] == b"$FL2", "sav 생성 (변수 2개)")

print("\n" + ("모두 통과" if fails == 0 else f"실패 {fails}건"))
sys.exit(1 if fails else 0)

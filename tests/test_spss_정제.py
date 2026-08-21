"""streamlit 을 스텁으로 갈아끼워 순수 로직(analyze)만 검증."""
import sys, types, importlib.util
from unittest.mock import MagicMock
import pandas as pd

stub = MagicMock(name="streamlit")
stub.columns.return_value = [MagicMock(), MagicMock(), MagicMock()]
sys.modules["streamlit"] = stub
sys.path.insert(0, "/home/claude/repo")

spec = importlib.util.spec_from_file_location("page", "/home/claude/repo/pages/3_🛠️_SPSS_정제.py")
page = importlib.util.module_from_spec(spec)
spec.loader.exec_module(page)

# --- make_valid_name ---
cases = [("성별", "Q3", "Q3"), ("SQ1", "Q1", "SQ1"), ("1차 구매", "Q7", "Q7"),
         ("", "", ""), ("TO", "Q9", "TO_"), ("만족도_1", "Q5_1", "Q5_1")]
for cand, fb, want in cases:
    got = page.make_valid_name(cand, fb)
    print(("OK  " if got == want else "실패"), f"make_valid_name({cand!r},{fb!r}) -> {got!r} (기대 {want!r})")

# --- analyze ---
df_raw = pd.DataFrame(columns=["ID", "Q1", "Q1_1", "Q1_2", "Q5_1", "Q5_2", "Q9", "Q10"])
df_code = pd.DataFrame([
    ["Q1",     "SQ1. 성별"],
    ["문2",    "성별 무관 질문"],          # 한글 라벨 -> 코드명 폴백
    ["Q5_RK1", "Q5. 선호 순위 (1순위)"],
    ["Q5_RK2", "Q5. 선호 순위 (2순위)"],
    ["Q9",     "만족도"],                  # 한글 라벨
    ["Q10",    "만족도"],                  # 같은 한글 라벨 -> 중복
])
final, updates, warns = page.analyze(df_raw, df_code, label_col=1)
print("\n경고:", warns)
for r in final:
    print(f"  {r['Raw 변수명']:6s} -> {r['변경할 변수명']:10s} [{r['상태']}]")
names = [r["변경할 변수명"] for r in final if r["변경할 변수명"]]
print("\n중복:", [n for n in set(names) if names.count(n) > 1])
print("규칙 위반:", [n for n in names if not page.is_valid_name(n)])
print("ID 제외됨:", "ID" not in [r["Raw 변수명"] for r in final])

# --- 대소문자 충돌 경고 ---
_, _, w2 = page.analyze(pd.DataFrame(columns=["Q1", "q1"]), pd.DataFrame([["Q1", "SQ1. x"]]), 1)
print("대소문자 경고:", bool(w2))

# --- 빈 RENAME 방지 ---
empty = pd.DataFrame([{"Raw 변수명": "Q1", "변경할 변수명": ""}])
syn, cnt = page.build_syntax(empty, "test")
print("\n빈 구문 count:", cnt, "| RENAME 포함:", "RENAME VARIABLES" in syn)
syn2, cnt2 = page.build_syntax(pd.DataFrame([{"Raw 변수명": "Q1", "변경할 변수명": "SQ1"}]), "test")
print("정상 구문 count:", cnt2, "|", [l for l in syn2.splitlines() if "=" in l])

"""변경 내용 추적(track changes)이 걸린 워드 설문지 파싱 검증.

    python3 tests/fixture_tracked_test.py

워드에서 추적을 켜고 수정하면 삽입된 글자는 <w:ins>, 삭제된 글자는
<w:del> + <w:delText> 로 저장된다. python-docx 의 paragraph.text 는 <w:ins> 안의
글자를 놓치므로 수정된 문항이 통째로 사라진다. spss_labels.element_text 가 이를 처리한다.

fixture_patterns.docx 와 .doc 양쪽에 같은 문항이 들어 있어 두 경로를 모두 검증한다.
견본에 문항을 더 넣으려면 tests/add_tracked_questions.py 를 참고하세요.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import spss_labels as sl

FIXTURES = [ROOT / "tests" / "fixture_patterns.docx",
            ROOT / "tests" / "fixture_patterns.doc"]

# (변수명, 검사, 설명)
CHECKS = [
    ("Q19", lambda v: sorted(v.values) == [1, 2] and "추적으로 삽입된" in v.label,
     "머리글과 보기가 모두 삽입된 문항"),
    ("Q20", lambda v: sorted(v.values) == [1, 2, 4],
     "보기 하나가 삭제되어 코드가 1,2,4 로 끊긴 문항"),
    ("Q21", lambda v: "해당 조직 구성원(들)" in v.label,
     "문단 중간만 삽입된 문항 (라벨이 잘리지 않음)"),
    ("Q22_1", lambda v: v.values.get(3) == "매우 그렇다(삽입)",
     "표 머리행 라벨이 삽입된 격자 문항"),
    ("Q22_2", lambda v: v.kind == "grid", "격자 두 번째 항목"),
]

fails = 0


def check(cond, msg):
    global fails
    print(("  OK   " if cond else "  실패 ") + msg)
    if not cond:
        fails += 1


for fixture in FIXTURES:
    if not fixture.exists():
        print(f"\n=== {fixture.name}: 파일 없음 (건너뜀) ===")
        continue
    variables = sl.parse_upload(fixture.read_bytes())
    table = {v.name: v for v in variables}
    print(f"\n=== {fixture.name} — 변수 {len(variables)}개 ===")

    for name, test, desc in CHECKS:
        var = table.get(name)
        if var is None:
            check(False, f"{name}: 변수 없음 ({desc})")
            continue
        check(test(var), f"{name}: {desc}")

    # 삭제된 글자는 결과에 남아 있으면 안 된다
    leaked = [v.name for v in variables
              if any("삭제된 보기" in str(x) for x in v.values.values())
              or "삭제된 보기" in v.label]
    check(not leaked, f"삭제된 글자가 결과에 남지 않음 (발견: {leaked[:3]})")

print("\n" + ("모두 통과" if fails == 0 else f"실패 {fails}건"))
sys.exit(1 if fails else 0)

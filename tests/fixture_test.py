"""견본 설문지 파싱 결과 검증 — 기대값을 직접 명시한 테스트.

    python3 tests/fixture_test.py

regression_check.py 는 "예전과 달라졌는지"만 봅니다. 이 테스트는 "맞는지"를 봅니다.
각 항목은 실제 설문지에서 한 번씩 문제를 일으켰던 서식 패턴입니다.
파서를 고친 뒤 이 파일을 돌려서 전부 통과하면, 그 패턴들은 안전합니다.

새 서식 패턴을 만나 파서를 고쳤다면
  1. tests/build_fixture.js 에 그 패턴을 추가하고 `node tests/build_fixture.js`
  2. 이 파일에 기대값을 한 줄 추가
하면 다음부터 자동으로 지켜집니다.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import spss_labels as sl

FIXTURE = ROOT / "tests" / "fixture_patterns.docx"
LEGACY = ROOT / "tests" / "fixture_patterns.doc"

# (변수명, 기대 조건 dict, 이 항목이 검증하는 서식 패턴)
CHECKS: list[tuple[str, dict, str]] = [
    ("SQ1", {"kind": "single", "codes": [1, 2], "label_has": "성별"},
     "한 단락에 탭으로 나열된 보기"),
    ("SQ2", {"kind": "numeric", "vtype": "numeric"},
     "밑줄 기입란(_____년) → 숫자 문항, 하위문항 있어도 사라지지 않음"),
    ("SQ2_1", {"kind": "single", "codes": [1, 2, 3]},
     "하위문항 번호"),
    ("SQ3", {"kind": "single", "n_values": 17, "value_at": (17, "제주")},
     "보기가 표 안에 여러 줄로 (17개 시·도 전부)"),
    ("SQ4", {"kind": "single", "n_values": 4, "todo": True},
     "워드 자동 불릿 보기 → 순차 코드 부여 + 확인필요"),
    ("SQ4_1", {"kind": "numeric"},
     "[RANGE : 1~23] → 직접 입력이라도 숫자 문항"),
    ("Q1", {"kind": "single", "codes": [1, 2, 3, 4], "value_at": (2, "보건소")},
     "동그라미 보기 ①②③④"),
    ("Q2", {"kind": "single", "missing": "99"},
     "모름/무응답 → 사용자 결측 자동 제안"),
    ("Q3_1_1", {"kind": "entry_check", "codes": [0, 1]},
     "체크박스 □ 열 → 0/1 선택 변수"),
    ("Q3_1_2", {"kind": "entry", "vtype": "numeric", "n_values": 0},
     "금액 입력 칸 → 숫자 변수 (열 제목을 코드로 오인하지 않음)"),
    ("Q4_1", {"kind": "grid", "n_values": 5},
     "격자 + 코드 칸이 `1)` 형태"),
    ("Q4_6", {"kind": "grid"},
     "격자 속성 번호에 결번이 있을 때 인쇄된 번호를 사용 (q4_5 아님)"),
    ("Q5", {"kind": "scale", "n_values": 5, "value_at": (1, "매우 불만족")},
     "항목 열이 없는 척도표 → 첫 척도점이 사라지지 않음"),
    ("Q6", {"kind": "scale", "n_values": 5, "todo": True},
     "코드가 전혀 없는 척도표 → 순차 부여 + 확인필요"),
    ("Q7", {"kind": "scale", "n_values": 5, "label_has": "제품 외관"},
     "소프트 리턴으로 두 줄이 된 머리글"),
    ("Q8_1", {"kind": "multi", "n_values": 5, "note_has": "복수응답"},
     "복수응답 [복수] → 보기 코드 방식"),
    ("Q9_R1", {"kind": "rank", "n_values": 4},
     "순위 표(1순위|2순위|3순위) + 보기 목록 → 숫자 순위형"),
    ("Q9_R3", {"kind": "rank"}, "순위 3개 모두 생성"),
    ("Q10_A_1_1", {"kind": "entry", "vtype": "numeric"},
     "문자 접미 머리글(Q10-a) + 전/후 숫자 입력 격자"),
    ("Q11_1", {"kind": "multi", "value_at": (1, "치즈")},
     "Code1..CodeN 보기 표 + 최대 N개 → 복수응답"),
    ("Q12_R1", {"kind": "rank", "n_values": 4},
     "'최대 3순위 선택' 문구 → 순위형"),
    ("Q13_1", {"kind": "scale", "label_has": "공정"},
     "번호 없는 진술문 + 척도표"),
    ("Q14", {"kind": "single", "n_values": 0, "todo": True},
     "드롭다운 → 보기 목록 없음으로 표시"),
    ("Q15", {"kind": "text", "vtype": "string"},
     "[직접 기입] 주관식 → 문자형"),
    ("Q16_1", {"kind": "grid", "label_has": "KBS"},
     "문항 서두가 길어도 항목명이 라벨에서 살아남음"),
    ("EQD1", {"kind": "grid", "label_has": "노력한 만큼"},
     "격자 항목이 자기 변수명을 갖는 경우(EQD1)"),
    ("COM1_1", {"kind": "entry", "vtype": "string"},
     "빈칸 기입 양식표 → 문자형"),
    ("COM1_2_DUP2", {"kind": "single", "n_values": 3, "todo": True},
     "변수명 충돌(Com1 격자 항목 vs Com1-2 문항) → 확인필요로 표시"),
    ("Q17_1", {"kind": "multi", "n_values": 4, "value_at": (4, "핸드볼")},
     "구분자가 탭+공백이거나 공백 하나여도 보기가 흡수되지 않음"),
    ("Q17_4", {"kind": "multi"},
     "줄 끝 지시문 `[PROG : 4) …]` 이 보기 분리를 방해하지 않음"),
    ("Q18", {"kind": "scale", "n_values": 9, "value_at": (1, "전혀")},
     "양 끝에만 라벨이 있는 척도 — 중간 코드(2~8)도 빈 라벨로 유지"),
    ("IN1_FAIL", {"kind": "dp_instruction", "label_has": "IN1_FAIL"},
     "[DP: … 변수 만들어주세요] → 변수 생성 + 라벨 유지"),
]

# 생성되어서는 안 되는 변수
MUST_NOT_EXIST = [
    ("Q13", "하위문항만 있는 블록 머리글은 변수가 되지 않아야 함"),
    ("Q4_5", "격자 속성 결번(5) 자리에 변수가 생기면 안 됨"),
]


def check_one(var, cond: dict) -> list[str]:
    fails = []
    if "kind" in cond and var.kind != cond["kind"]:
        fails.append(f"유형 {var.kind} (기대 {cond['kind']})")
    if "vtype" in cond and var.vtype != cond["vtype"]:
        fails.append(f"자료형 {var.vtype} (기대 {cond['vtype']})")
    if "codes" in cond and sorted(var.values) != cond["codes"]:
        fails.append(f"코드 {sorted(var.values)} (기대 {cond['codes']})")
    if "n_values" in cond and len(var.values) != cond["n_values"]:
        fails.append(f"값라벨 {len(var.values)}개 (기대 {cond['n_values']}개)")
    if "value_at" in cond:
        code, expected = cond["value_at"]
        got = var.values.get(code, "")
        if expected not in got:
            fails.append(f"코드 {code} 라벨 '{got}' (기대 '{expected}' 포함)")
    if "label_has" in cond and cond["label_has"] not in var.label:
        fails.append(f"라벨에 '{cond['label_has']}' 없음: {var.label[:40]}")
    if "note_has" in cond and cond["note_has"] not in var.note:
        fails.append(f"비고에 '{cond['note_has']}' 없음")
    if "missing" in cond and var.missing != cond["missing"]:
        fails.append(f"결측 '{var.missing}' (기대 '{cond['missing']}')")
    if cond.get("todo") and "확인필요" not in var.note:
        fails.append("확인필요 표시 없음")
    if cond.get("todo") is False and "확인필요" in var.note:
        fails.append("불필요한 확인필요 표시")
    return fails


def run(path: Path, tag: str) -> int:
    variables = {v.name: v for v in sl.parse_upload(path.read_bytes())}
    print(f"\n=== {tag} — 변수 {len(variables)}개 ===")
    failed = 0
    for name, cond, pattern in CHECKS:
        var = variables.get(name)
        if var is None:
            print(f"  실패  {name:12s} 변수 없음  ({pattern})")
            failed += 1
            continue
        fails = check_one(var, cond)
        if fails:
            print(f"  실패  {name:12s} {'; '.join(fails)}  ({pattern})")
            failed += 1
    for name, why in MUST_NOT_EXIST:
        if name in variables:
            print(f"  실패  {name:12s} 생성되면 안 되는 변수  ({why})")
            failed += 1

    passed = len(CHECKS) + len(MUST_NOT_EXIST) - failed
    print(f"  통과 {passed} / {len(CHECKS) + len(MUST_NOT_EXIST)}")
    return failed


def main() -> int:
    if not FIXTURE.exists():
        print(f"견본 파일이 없습니다: {FIXTURE}\n  node tests/build_fixture.js 로 생성하세요.")
        return 1
    failed = run(FIXTURE, "fixture_patterns.docx")

    if LEGACY.exists():
        import shutil

        if shutil.which("soffice") or shutil.which("libreoffice"):
            failed += run(LEGACY, "fixture_patterns.doc (구형 이진 포맷)")
        else:
            print("\n=== .doc 검증 건너뜀 (LibreOffice 없음) ===")

    print("\n" + ("모두 통과" if failed == 0 else f"실패 {failed}건"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

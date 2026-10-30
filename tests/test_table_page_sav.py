"""테이블 생성 화면의 .sav 읽기 — EUC-KR 선언 sav 에 CP949 확장 한글('똠')이 있으면 기본 읽기가 실패했다.

    python -m unittest tests.test_table_page_sav      (저장소 맨 위 폴더에서)

화면 파일은 불러오면 streamlit 화면을 그리므로 load_sav 함수만 꺼내 돌린다 (dp-skills#53, 2026-10-31).
합성 sav(2행)는 시험 때 dp-skills 의 sav_cp949 로 만든다 (.sav 는 .gitignore 대상이라 저장소에 두지 않음) — 스킬 폴더가 없으면 건너뜀.
"""
import ast
import os
import sys
import tempfile
import unittest
from pathlib import Path

import pyreadstat

ROOT = Path(__file__).resolve().parent.parent
PAGE = ROOT / "pages" / "11. 테이블_생성.py"
SKILL = next((p for p in (Path.home() / ".agents" / "skills" / "dp-editing" / "scripts",
                          Path.home() / ".claude" / "skills" / "dp-editing" / "scripts") if (p / "sav_cp949.py").is_file()), None)


def _load_sav():
    tree = ast.parse(PAGE.read_text(encoding="utf-8"))
    fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "load_sav")
    fn.decorator_list = []                     # st.cache_data 없이
    ns = {"tempfile": tempfile, "pyreadstat": pyreadstat, "Path": Path}
    exec(compile(ast.Module(body=[fn], type_ignores=[]), str(PAGE), "exec"), ns)
    return ns["load_sav"]


@unittest.skipUnless(SKILL, "dp-skills 의 sav_cp949.py 없음")
class TablePageSavTest(unittest.TestCase):
    def setUp(self):
        sys.path.insert(0, str(SKILL))
        import sav_cp949
        d = tempfile.mkdtemp()
        self.addCleanup(lambda: __import__("shutil").rmtree(d, True))
        self.sav = os.path.join(d, "ttom.sav")
        sav_cp949.write_sav(self.sav, ["NO", "Q1", "Q2"], {"NO": "F8.0", "Q1": "F8.0", "Q2": "A20"}, {"Q1": "Q1. 똠 문항", "Q2": "Q2. 의견"},
                            {"Q1": {1.0: "똠양꿍", 2.0: "보통"}}, [{"NO": 1.0, "Q1": 1.0, "Q2": "똠얌"}, {"NO": 2.0, "Q1": 2.0, "Q2": "보통"}])

    def test_extended_hangul_sav(self):
        with self.assertRaises(pyreadstat.ReadstatError):   # 픽스처가 기본 읽기로는 실패하는 꼴인지
            pyreadstat.read_sav(self.sav)
        df, meta = _load_sav()(open(self.sav, "rb").read())
        self.assertEqual(list(df["Q2"]), ["똠얌", "보통"])
        self.assertEqual(meta.column_names_to_labels["Q1"], "Q1. 똠 문항")
        self.assertEqual(meta.variable_value_labels["Q1"][1.0], "똠양꿍")


if __name__ == "__main__":
    unittest.main()

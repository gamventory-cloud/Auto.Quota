# -*- coding: utf-8 -*-
"""hwp_survey 회귀 테스트.

.hwpx는 ZIP+XML이라 테스트용 파일을 코드로 만들 수 있다.
바이너리 픽스처를 저장소에 넣지 않아도 CI에서 전 과정을 돌려볼 수 있다.
"""

import zipfile

import pytest

from hwp_survey import items_to_dsl, parse_dsl, read_survey, summarize
from hwp_survey.parser import scale_columns
from hwp_survey.reader import clean
from hwp_survey.writer import SurveyWriter, build_docx

NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def para(text):
    return f"<hp:p><hp:run><hp:t>{text}</hp:t></hp:run></hp:p>"


def cell(text):
    return f"<hp:tc><hp:subList>{para(text)}</hp:subList></hp:tc>"


def table(rows):
    body = "".join("<hp:tr>" + "".join(cell(c) for c in r) + "</hp:tr>" for r in rows)
    return f"<hp:p><hp:run><hp:tbl>{body}</hp:tbl></hp:run></hp:p>"


@pytest.fixture
def sample_hwpx(tmp_path):
    body = [
        para("직무교육 만족도 조사"),
        para("안녕하십니까? 본 조사는 교육과정 개선을 목적으로 합니다."),
        para("Ⅰ. 응답자 일반사항"),
        para("1. 귀하의 소속은 어디입니까? ① 영업 ② 개발 ③ 관리"),
        para("2. 귀하의 직급을 선택해 주십시오."),
        para("① 사원"), para("② 대리"), para("③ 과장 이상"),
        para("3. 유익했던 내용을 모두 골라 주십시오."),
        para("① 실습 ② 사례 발표"),
        para("Ⅱ. 교육 만족도"),
        table([["◀", "①", "②", "③", "④", "⑤", "▶"],
               ["전혀 그렇지 않다", "매우 그렇다"]]),
        table([["‘내용’에 관한 문항입니다."],
               ["1", "강의 내용이 유익했다.", "①", "②", "③", "④", "⑤"],
               ["2", "난이도가 적절했다.", "①", "②", "③", "④", "⑤"]]),
        para("4. 향후 교육에 대한 의견을 자유롭게 적어 주십시오."),
    ]
    xml = (f'<?xml version="1.0" encoding="UTF-8"?><hp:sec xmlns:hp="{NS}">'
           + "".join(body) + "</hp:sec>")
    path = tmp_path / "설문지.hwpx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/section0.xml", xml)
    return str(path)


# ------------------------------------------------------------------ 추출
def test_reader_extracts_paragraphs_and_tables(sample_hwpx):
    items = read_survey(sample_hwpx)
    assert sum(1 for k, _ in items if k == "table") == 2
    assert ("p", "Ⅰ. 응답자 일반사항") in items


def test_clean_joins_run_fragments():
    # 한글은 글자 모양이 바뀌는 지점마다 텍스트를 조각내므로 붙여 읽어야 한다
    assert clean("‘ 자율성 ’") == "‘자율성’"
    assert clean("200 만원 미만") == "200만원 미만"
    assert clean("[ 설문지 종료 ]") == "[설문지 종료]"


def test_clean_tightens_linebreak_padding():
    assert clean("기분이 좋   다.") == "기분이 좋다."


# ------------------------------------------------------------------ 파싱
def test_inline_options_split_into_choices(sample_hwpx):
    dsl = items_to_dsl(read_survey(sample_hwpx))
    assert "1. 귀하의 소속은 어디입니까? [단일]" in dsl
    assert "- 영업" in dsl


def test_options_on_following_lines_set_single_type(sample_hwpx):
    dsl = items_to_dsl(read_survey(sample_hwpx))
    assert "2. 귀하의 직급을 선택해 주십시오. [단일]" in dsl   # [단답]이 아니어야 한다


def test_multi_and_open_types(sample_hwpx):
    dsl = items_to_dsl(read_survey(sample_hwpx))
    assert "[복수]" in dsl        # '모두' 골라 주십시오
    assert "[장문]" in dsl        # '자유롭게' 적어 주십시오


def test_scale_legend_becomes_matrix_header():
    cols = scale_columns([["◀", "①", "②", "③", "④", "⑤", "▶"],
                          ["전혀 그렇지 않다", "매우 그렇다"]])
    assert len(cols) == 5
    assert cols[0].startswith("①") and "전혀" in cols[0]
    assert "매우" in cols[-1]


def test_matrix_table_with_group_row(sample_hwpx):
    blocks = parse_dsl(items_to_dsl(read_survey(sample_hwpx)))
    matrix = [b for b in blocks if b["kind"] == "question" and b["type"] == "표"]
    assert len(matrix) == 1
    kinds = [o["type"] for o in matrix[0]["options"]]
    assert kinds.count("group") == 1
    assert kinds.count("row") == 2


def test_summary_counts(sample_hwpx):
    found = summarize(parse_dsl(items_to_dsl(read_survey(sample_hwpx))))
    assert found["문항"] == 5          # 4문항 + 매트릭스 1
    assert found["매트릭스 표"] == 1
    assert found["섹션"] == 2


def test_dsl_roundtrip_is_stable(sample_hwpx):
    dsl = items_to_dsl(read_survey(sample_hwpx))
    assert summarize(parse_dsl(dsl)) == summarize(parse_dsl(dsl))


# ------------------------------------------------------------------ 출력
def test_build_docx_returns_openable_document(sample_hwpx):
    import io

    from docx import Document

    data = build_docx(parse_dsl(items_to_dsl(read_survey(sample_hwpx))))
    assert data[:2] == b"PK"                       # docx = ZIP
    doc = Document(io.BytesIO(data))
    assert len(doc.tables) == 1
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "1. 귀하의 소속은 어디입니까?" in text
    assert "감사" in text


def test_writer_options_apply(sample_hwpx):
    blocks = parse_dsl(items_to_dsl(read_survey(sample_hwpx)))
    writer = SurveyWriter(font="바탕", multi_mark="[ ]").write(blocks)
    xml = writer.doc.element.xml
    assert "바탕" in xml
    assert "[ ]" in xml


def test_unsupported_extension_raises(tmp_path):
    path = tmp_path / "설문지.doc"
    path.write_bytes(b"x")
    with pytest.raises(ValueError):
        read_survey(str(path))

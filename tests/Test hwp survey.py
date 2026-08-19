# -*- coding: utf-8 -*-
"""hwp_survey 회귀 테스트.

.hwpx는 ZIP+XML이라 테스트용 파일을 코드로 만들 수 있다.
바이너리 픽스처를 저장소에 넣지 않아도 CI에서 전 과정을 돌려볼 수 있다.
"""

import xml.sax.saxutils as sax
import zipfile

import pytest

from hwp_survey import items_to_dsl, parse_dsl, read_survey, summarize
from hwp_survey.parser import scale_columns
from hwp_survey.reader import clean
from hwp_survey.writer import SurveyWriter, build_docx

NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def para(text):
    # <제목> 처럼 꺾쇠가 들어간 본문도 있으므로 XML 이스케이프가 필요하다
    return f"<hp:p><hp:run><hp:t>{sax.escape(text)}</hp:t></hp:run></hp:p>"


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


# ================================================================== DP 스크립트
from hwp_survey.dp import (DPWriter, items_to_dp_dsl, parse_dp,  # noqa: E402
                           summarize_dp)


@pytest.fixture
def dp_hwpx(tmp_path):
    """리서치 초안 형태(문단만 있는 설문지) 샘플."""
    body = [
        para("<IT 리뷰 유튜버 설문지>"),
        para("■ 조사 대상자: 전국 만 20~39세 남녀"),
        para("■ 샘플 수: 250샘플"),
        para("■ 쿼터:"),
        para("20대 30대 합계"),
        para("남 62 63 125"),
        para("여 63 62 125"),
        para("---"),
        para("SQ1. 귀하의 성별은 무엇입니까?"),
        para("① 남성"), para("② 여성"),
        para("SQ2. 귀하의 연령은 어떻게 되십니까?"),
        para("(출생년도 입력)"),
        para("[PROG: 만20~39세만 진행]"),
        para("SQ3. 거주하고 계신 지역은 어디입니까?"),
        para("(운동 설문과 동일)"),
        para("SQ4. 평소 온라인 동영상 플랫폼을 이용하십니까?"),
        para("① 유튜브"), para("② 넷플릭스"), para("③ 이용하지 않음"),
        para("Q1. 다음은 유튜버의 특성에 관한 질문입니다."),
        para("(5점 척도)"),
        para("① 매력적이다."), para("② 멋있다."),
    ]
    xml = (f'<?xml version="1.0" encoding="UTF-8"?><hp:sec xmlns:hp="{NS}">'
           + "".join(body) + "</hp:sec>")
    path = tmp_path / "초안.hwpx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/section0.xml", xml)
    return str(path)


def test_dp_header_fields_and_quota(dp_hwpx):
    doc = parse_dp(items_to_dp_dsl(read_survey(dp_hwpx)))
    assert doc["제목"] == "IT 리뷰 유튜버 설문지"
    assert doc["샘플수"] == "250샘플"
    assert doc["쿼터표"][0] == ["", "20대", "30대", "합계"]   # 좌상단 빈 칸
    assert doc["쿼터표"][1] == ["남", "62", "63", "125"]


def test_dp_labels_are_preserved(dp_hwpx):
    doc = parse_dp(items_to_dp_dsl(read_survey(dp_hwpx)))
    labels = [b["label"] for b in doc["blocks"] if b["kind"] == "question"]
    assert labels == ["SQ1", "SQ2", "SQ3", "SQ4", "Q1"]        # 번호를 다시 매기지 않는다


def test_dp_response_tags(dp_hwpx):
    doc = parse_dp(items_to_dp_dsl(read_survey(dp_hwpx)))
    tags = {b["label"]: b["tag"] for b in doc["blocks"] if b["kind"] == "question"}
    assert tags["SQ1"] == "1개선택"
    assert tags["SQ2"] == "출생년도 입력"
    assert tags["SQ3"] == "지도에서 선택"
    assert tags["SQ4"] == "모두선택"          # '이용하지 않음' 배타 보기가 있으므로
    assert tags["Q1"] == "행별 1개선택"


def test_dp_matrix_scale_and_rows(dp_hwpx):
    doc = parse_dp(items_to_dp_dsl(read_survey(dp_hwpx)))
    q1 = [b for b in doc["blocks"] if b.get("label") == "Q1"][0]
    assert q1["scale"][0] == "전혀 그렇지 않다" and len(q1["scale"]) == 5
    assert [o["text"] for o in q1["options"]] == ["매력적이다.", "멋있다."]
    assert "체크해 주세요" in q1["text"]      # 안내 문장 자동 추가


def test_dp_alone_prog_added(dp_hwpx):
    dsl = items_to_dp_dsl(read_survey(dp_hwpx))
    assert "%PROG: 3번 보기는 단독선택만 가능" in dsl


def test_dp_alone_prog_can_be_disabled(dp_hwpx):
    dsl = items_to_dp_dsl(read_survey(dp_hwpx), add_alone_prog=False)
    assert "단독선택만 가능" not in dsl


def test_dp_prog_lines_kept(dp_hwpx):
    doc = parse_dp(items_to_dp_dsl(read_survey(dp_hwpx)))
    progs = [b["text"] for b in doc["blocks"] if b["kind"] == "prog"]
    assert "만20~39세만 진행" in progs
    assert "17개시도 지도제시" in progs        # 지도형 기본 지시문


def test_dp_summary(dp_hwpx):
    found = summarize_dp(parse_dp(items_to_dp_dsl(read_survey(dp_hwpx))))
    assert found["문항"] == 5
    assert found["선정문항(SQ)"] == 4
    assert found["행별 표"] == 1


def test_dp_docx_structure(dp_hwpx):
    import io

    from docx import Document

    doc = parse_dp(items_to_dp_dsl(read_survey(dp_hwpx)))
    doc["제외"] = "2026060452 참여자 제외"
    doc["blocks"].append({"kind": "verify", "text": "Q1 일자찍기 제외"})
    data = DPWriter().write(doc).to_bytes()

    d = Document(io.BytesIO(data))
    spec = d.tables[0]
    assert [c.text for c in spec.rows[0].cells][0] == "대상자"
    assert len(spec.rows[2].cells[1].tables) == 1          # 쿼터 격자는 중첩 표
    matrix = d.tables[-1]
    assert len(matrix.columns) == 6 and len(matrix.rows) == 3
    assert matrix.rows[0].cells[1].text.endswith("1")      # 라벨 + 척도 번호

    xml = d.element.xml
    assert 'w:val="yellow"' in xml                          # 제외 줄 형광
    assert 'w:color w:val="0000FF"' in xml                  # PROG 파란색
    assert 'w:color w:val="FF0000"' in xml                  # 검증 빨간색
    assert "나눔고딕" in xml


def test_dp_handles_table_based_survey(sample_hwpx):
    """표로 리커트를 짠 학술 설문지도 DP 스크립트로 변환된다."""
    doc = parse_dp(items_to_dp_dsl(read_survey(sample_hwpx)))
    qs = [b for b in doc["blocks"] if b["kind"] == "question"]
    labels = [q["label"] for q in qs]

    assert labels[:3] == ["SQ1", "SQ2", "SQ3"]        # 행별 문항 앞은 선정 문항
    assert any(l.startswith("Q") for l in labels)     # 행별 문항부터 본 문항
    matrix = [q for q in qs if q["tag"] == "행별 1개선택"]
    assert len(matrix) == 1
    assert [o["type"] for o in matrix[0]["options"]].count("group") == 1
    assert doc["쿼터표"] == []                        # 표를 쿼터로 오인하지 않는다


def test_dp_terminate_option_becomes_prog(tmp_path):
    body = [para("1. 이용 경험이 있으십니까?"), para("① 예   ② 아니오[설문지 종료]")]
    xml = (f'<?xml version="1.0" encoding="UTF-8"?><hp:sec xmlns:hp="{NS}">'
           + "".join(body) + "</hp:sec>")
    path = tmp_path / "종료.hwpx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("Contents/section0.xml", xml)

    dsl = items_to_dp_dsl(read_survey(str(path)))
    assert "- 아니오" in dsl and "설문지 종료]" not in dsl
    assert "%PROG: 2번 선택자 설문 종료" in dsl


def test_dp_scale_middle_labels_filled(tmp_path):
    """양 끝만 라벨이 있는 척도 안내표의 가운데 라벨을 채운다."""
    from hwp_survey.dp import fill_scale
    cols = ["① 전혀 그렇지 않다", "②", "③", "④", "⑤ 매우 그렇다"]
    assert fill_scale(cols) == ["전혀 그렇지 않다", "그렇지 않다", "보통이다",
                                "그렇다", "매우 그렇다"]

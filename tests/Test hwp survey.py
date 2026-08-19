# -*- coding: utf-8 -*-
"""한글 설문지(.hwp/.hwpx) -> 워드(.docx) 변환 페이지.

두 가지 출력 서식을 고를 수 있다.
  * DP 스크립트   : 리서치사 납품용. 조사 개요표 + SQ/Q 번호 + [PROG] 지시문
  * 인쇄용 설문지 : 응답자 배포용. 보기 기호와 체크 칸

문서 처리 로직은 hwp_survey 패키지에 있고, 이 파일은 화면만 담당한다.
"""

import csv
import io
import os
import tempfile

import streamlit as st

from hwp_survey import items_to_dsl, parse_dsl, read_survey, summarize
from hwp_survey.dp import DPWriter, items_to_dp_dsl, parse_dp, summarize_dp
from hwp_survey.writer import SurveyWriter

PAGE_TITLE = "설문지 변환 (한글 → 워드)"
SS = "survey_docx"          # 다른 페이지와 섞이지 않도록 세션 키 접두어

st.set_page_config(page_title=PAGE_TITLE, page_icon="📋", layout="wide")
st.title(PAGE_TITLE)

HELP_DP = """
| 표기 | 뜻 |
|---|---|
| `@제목:` `@대상자:` `@샘플수:` | 조사 개요표 |
| `@쿼터:` `@쿼터표: ,남자,여자` | 쿼터 설명 + 쿼터 격자(줄마다 한 행) |
| `@제외:` | 개요표 마지막 줄(노란 형광) |
| `SQ1. 문항 [1개선택]` | 선정 문항 |
| `Q1. 문항 [행별 1개선택]` | 본 문항 |
| `@행별: 전혀 그렇지 않다,...` | 행별 표의 열 라벨 |
| `- 보기` | 보기 / 행별 표의 행 |
| `@표: 셀,셀,셀` | 원본 표를 격자 그대로 (줄마다 한 행) |
| `## 영역 이름` | 영역 배너(진한 바탕 한 칸 표) |
| `%PROG: ...` | 파란색 프로그래밍 지시문 |
| `%검증: ...` | 빨간색 데이터 검증 지시문 |
| `~ 상자글` | 테두리 상자 |
| `! 안내` | 상자 안 보조 줄 |

응답 방식 태그는 그대로 문서에 찍히므로 자유롭게 바꿔도 됩니다:
`[1개선택]` `[모두선택]` `[행별 1개선택]` `[출생년도 입력]` `[지도에서 선택]`
"""

HELP_PRINT = """
| 표기 | 뜻 |
|---|---|
| `# 제목` / `> 안내문` / `~ 상자글` | 표지 요소 |
| `## Ⅰ. 섹션` | 섹션 제목 |
| `1. 문항 [단일]` | 문항 + 유형 |
| `- 보기` | 보기 (표 유형이면 표의 행) |
| `-- 소제목` | 표 안 소제목 행 |
| `! 지시문` | 문항 아래 작은 안내 |
| `@표: 셀,셀,셀` | 원본 표를 격자 그대로 (줄마다 한 행) |

유형 태그: `[단일]` `[복수]` `[단답]` `[장문]` `[척도:1-7]`
`[표:① 전혀 그렇지 않다,②,③,④,⑤ 매우 그렇다]`
"""


@st.cache_data(show_spinner=False)
def extract(file_bytes: bytes, suffix: str, tighten: bool, style: str,
            matrix_hint: bool, alone_prog: bool):
    """업로드 파일에서 문단·표를 뽑아 중간 텍스트로. 바이트가 같으면 재사용된다."""
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        path = tmp.name
    try:
        items = read_survey(path, tighten=tighten)
    finally:
        os.unlink(path)
    stats = {"문단": sum(1 for k, _ in items if k == "p"),
             "표": sum(1 for k, _ in items if k == "table")}
    if style == "DP 스크립트":
        return items_to_dp_dsl(items, add_matrix_hint=matrix_hint,
                               add_alone_prog=alone_prog), stats
    return items_to_dsl(items), stats


def to_csv(rows) -> bytes:
    buf = io.StringIO()
    csv.writer(buf).writerows(rows)
    return buf.getvalue().encode("utf-8-sig")       # 엑셀에서 한글 깨지지 않게


def items_csv_print(blocks) -> bytes:
    rows = [["문항번호", "유형", "섹션", "문항", "보기/세부항목"]]
    section, qno = "", 0
    for b in blocks:
        if b["kind"] == "section":
            section = b["text"]
        elif b["kind"] == "question":
            qno += 1
            items = [o["text"] for o in b["options"] if o["type"] == "row"]
            if b["type"] == "표":
                rows += [[f"{qno}-{i}", b["type"], section, b["text"], t]
                         for i, t in enumerate(items, 1)]
            else:
                rows.append([qno, b["type"], section, b["text"], " | ".join(items)])
    return to_csv(rows)


def items_csv_dp(doc) -> bytes:
    rows = [["문항번호", "응답방식", "문항", "보기/행"]]
    for b in doc["blocks"]:
        if b["kind"] != "question":
            continue
        items = [o["text"] for o in b["options"] if o["type"] == "row"]
        if b["tag"].startswith("행별"):
            rows += [[f"{b['label']}_{i}", b["tag"], b["text"], t]
                     for i, t in enumerate(items, 1)]
        else:
            rows.append([b["label"], b["tag"], b["text"], " | ".join(items)])
    return to_csv(rows)


with st.sidebar:
    st.subheader("출력 서식")
    style = st.radio("서식", ["DP 스크립트", "인쇄용 설문지"], index=0,
                     captions=["리서치사 납품용 스크립트", "응답자 배포용 설문지"],
                     label_visibility="collapsed")
    st.divider()

    if style == "DP 스크립트":
        st.subheader("조사 개요")
        quota_note = st.text_input("쿼터 설명", "성별*연령대별 균등할당")
        exclude = st.text_input("개요표 마지막 줄(형광)", "",
                                placeholder="예: 2026060452 참여자 제외")
        verify = st.text_input("데이터 검증 지시문", "",
                               placeholder="예: Q1~Q2 일자찍기는 최종 납품 데이터에서 제외")

        st.subheader("자동 처리")
        matrix_hint = st.toggle(
            "행별 문항에 안내 문장 붙이기", value=True,
            help="'귀하의 의견과 가장 일치하는 정도에 체크해 주세요.'를 문항 끝에 붙입니다.")
        alone_prog = st.toggle(
            "배타 보기에 단독선택 PROG 넣기", value=True,
            help="'이 중 없음', '이용하지 않음' 같은 보기가 있으면 "
                 "[PROG: N번 보기는 단독선택만 가능]을 넣습니다.")

        st.subheader("문서 서식")
        font = st.selectbox("글꼴", ["나눔고딕", "맑은 고딕", "함초롬돋움"], index=0)
        base_pt = st.slider("글자 크기(pt)", 9.0, 11.0, 10.0, 0.5)
        row_label_cm = st.slider("행별 표 문항열 너비(cm)", 7.0, 12.0, 9.85, 0.05)
    else:
        matrix_hint = alone_prog = False
        st.subheader("문서 서식")
        font = st.selectbox("한글 글꼴",
                            ["맑은 고딕", "함초롬돋움", "바탕", "굴림", "나눔고딕"], index=0)
        base_pt = st.slider("본문 글자 크기(pt)", 9.0, 12.0, 10.5, 0.5)
        single_mark = st.selectbox("단일 응답 기호", ["○", "□", "( )"], index=0)
        multi_mark = st.selectbox("복수 응답 기호", ["□", "☐", "[ ]"], index=0)
        row_label_cm = st.slider("표 첫 열 너비(cm)", 5.0, 12.0, 9.0, 0.5)
        accent = st.color_picker("섹션 제목 색", "#1F3B63")

    st.divider()
    st.subheader("추출 옵션")
    tighten = st.toggle(
        "줄바꿈 채움 공백 정리", value=True,
        help="한글에서 줄 끝을 공백으로 채운 흔적을 붙입니다. '만족 감을'→'만족감을'로 "
             "살아나지만, 단어 사이를 여러 칸 띄운 곳은 붙어버릴 수 있습니다.")

uploaded = st.file_uploader("한글 설문지 올리기", type=["hwp", "hwpx"])

if uploaded is None:
    st.info("변환할 .hwp 또는 .hwpx 파일을 올려주세요. 암호가 걸린 파일과 HWP 3.x 이하 "
            "옛 형식은 한글에서 다시 저장한 뒤 올려주세요.")
    with st.expander("중간 텍스트 문법 보기"):
        st.markdown(HELP_DP if style == "DP 스크립트" else HELP_PRINT)
    st.stop()

suffix = os.path.splitext(uploaded.name)[1].lower()
key = f"{uploaded.name}:{uploaded.size}:{tighten}:{style}:{matrix_hint}:{alone_prog}"

try:
    with st.spinner("한글 파일에서 문단과 표를 읽는 중"):
        auto_dsl, stats = extract(uploaded.getvalue(), suffix, tighten, style,
                                  matrix_hint, alone_prog)
except Exception as err:                                        # noqa: BLE001
    st.error(f"파일을 읽지 못했습니다: {err}")
    st.caption("한글에서 '다른 이름으로 저장 → HWPX'로 저장한 파일이 가장 잘 읽힙니다.")
    st.stop()

if st.session_state.get(f"{SS}_key") != key:                    # 조건이 바뀌면 초기화
    st.session_state[f"{SS}_key"] = key
    st.session_state[f"{SS}_dsl"] = auto_dsl

left, right = st.columns([3, 2], gap="large")

with left:
    st.subheader("중간 텍스트")
    st.caption("자동 인식이 어긋난 곳을 여기서 고치면 그대로 문서에 반영됩니다.")
    st.session_state[f"{SS}_dsl"] = st.text_area(
        "중간 텍스트", value=st.session_state[f"{SS}_dsl"], height=560,
        label_visibility="collapsed")
    if st.button("자동 인식 결과로 되돌리기"):
        st.session_state[f"{SS}_dsl"] = auto_dsl
        st.rerun()

with right:
    st.subheader("인식 결과")
    dsl = st.session_state[f"{SS}_dsl"]
    stem = os.path.splitext(uploaded.name)[0]
    mime = ("application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document")

    a, b = st.columns(2)
    a.metric("문단", stats["문단"])
    b.metric("표", stats["표"])

    if style == "DP 스크립트":
        doc = parse_dp(dsl)
        if quota_note and not doc["쿼터"]:
            doc["쿼터"] = quota_note
        if exclude:
            doc["제외"] = exclude
        if verify:
            doc["blocks"].append({"kind": "verify", "text": verify})

        found = summarize_dp(doc)
        c, d = st.columns(2)
        c.metric("문항", found["문항"])
        d.metric("선정문항(SQ)", found["선정문항(SQ)"])
        e, f = st.columns(2)
        e.metric("행별 표", found["행별 표"])
        f.metric("PROG 지시문", found["PROG 지시문"])
        st.metric("그대로 옮긴 표", found["일반 표"],
                  help="분류되지 않은 원본 표(빈도 표, 기입 표)를 격자 그대로 옮긴 개수")
        empty = found["문항"] == 0
        docx_bytes = DPWriter(font=font, base_pt=base_pt,
                              row_label_cm=row_label_cm).write(doc).to_bytes()
        csv_bytes = items_csv_dp(doc)
    else:
        blocks = parse_dsl(dsl)
        found = summarize(blocks)
        c, d = st.columns(2)
        c.metric("문항", found["문항"])
        d.metric("섹션", found["섹션"])
        e, f = st.columns(2)
        e.metric("매트릭스 표", found["매트릭스 표"])
        f.metric("매트릭스 세부항목", found["매트릭스 세부항목"])
        st.metric("그대로 옮긴 표", found["일반 표"],
                  help="분류되지 않은 원본 표(빈도 표, 기입 표)를 격자 그대로 옮긴 개수")
        empty = found["문항"] == 0
        docx_bytes = SurveyWriter(
            font=font, base_pt=base_pt, single_mark=single_mark,
            multi_mark=multi_mark, row_label_cm=row_label_cm,
            accent=accent.lstrip("#").upper()).write(blocks).to_bytes()
        csv_bytes = items_csv_print(blocks)

    if empty:
        st.warning("문항을 찾지 못했습니다. 왼쪽 텍스트의 문항 줄 형식을 확인해 주세요.")

    st.download_button("워드 파일 내려받기", docx_bytes, f"{stem}.docx", mime,
                       type="primary", use_container_width=True)
    st.download_button("중간 텍스트 내려받기", dsl, f"{stem}.txt", "text/plain",
                       use_container_width=True)
    st.download_button("문항 목록 CSV 내려받기", csv_bytes, f"{stem}_문항목록.csv",
                       "text/csv", use_container_width=True,
                       help="변수 라벨 작업에 넘겨 쓸 수 있는 문항·보기 목록입니다.")

    with st.expander("문법 도움말"):
        st.markdown(HELP_DP if style == "DP 스크립트" else HELP_PRINT)

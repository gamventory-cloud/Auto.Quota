# -*- coding: utf-8 -*-
"""한글 설문지(.hwp/.hwpx) -> 워드(.docx) 설문지 변환 페이지.

문서 처리 로직은 hwp_survey 패키지에 있고, 이 파일은 화면만 담당한다.
"""

import csv
import io
import os
import tempfile

import streamlit as st

from hwp_survey import items_to_dsl, parse_dsl, read_survey, summarize
from hwp_survey.writer import SurveyWriter

PAGE_TITLE = "설문지 변환 (한글 → 워드)"
SS = "survey_docx"          # 다른 페이지와 섞이지 않도록 세션 키 접두어

st.set_page_config(page_title=PAGE_TITLE, page_icon="📋", layout="wide")
st.title(PAGE_TITLE)
st.caption("한글 설문지를 워드 설문지로 옮깁니다. 표로 짠 리커트 척도 문항을 표 그대로 살립니다.")

DSL_HELP = """
| 표기 | 뜻 |
|---|---|
| `# 제목` | 설문 제목 |
| `> 안내문` | 인사말·연구 목적 |
| `~ 상자글` | 테두리 상자(용어 정의, 유의사항) |
| `## Ⅰ. 섹션` | 섹션 제목 |
| `! 지시문` | 문항 아래 작은 안내 |
| `1. 문항 [단일]` | 문항 + 유형 |
| `- 보기` | 보기 (표 유형이면 표의 행) |
| `-- 소제목` | 표 안 소제목 행 |

유형 태그: `[단일]` `[복수]` `[단답]` `[장문]` `[척도:1-7]`
`[표:① 전혀 그렇지 않다,②,③,④,⑤ 매우 그렇다]`
"""


@st.cache_data(show_spinner=False)
def extract(file_bytes: bytes, suffix: str, tighten: bool):
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
    return items_to_dsl(items), stats


def items_csv(blocks) -> bytes:
    """문항 목록을 CSV로. 변수 라벨 작업(spss_labels)에 넘겨 쓰기 위한 것."""
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["문항번호", "유형", "섹션", "문항", "보기/세부항목"])
    section = ""
    qno = 0
    for b in blocks:
        if b["kind"] == "section":
            section = b["text"]
        elif b["kind"] == "question":
            qno += 1
            rows = [o["text"] for o in b["options"] if o["type"] == "row"]
            if b["type"] == "표":
                for i, row in enumerate(rows, 1):
                    w.writerow([f"{qno}-{i}", b["type"], section, b["text"], row])
            else:
                w.writerow([qno, b["type"], section, b["text"], " | ".join(rows)])
    return buf.getvalue().encode("utf-8-sig")      # 엑셀에서 한글 깨지지 않게


with st.sidebar:
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
        st.markdown(DSL_HELP)
    st.stop()

suffix = os.path.splitext(uploaded.name)[1].lower()
key = f"{uploaded.name}:{uploaded.size}:{tighten}"

try:
    with st.spinner("한글 파일에서 문단과 표를 읽는 중"):
        auto_dsl, stats = extract(uploaded.getvalue(), suffix, tighten)
except Exception as err:                                        # noqa: BLE001
    st.error(f"파일을 읽지 못했습니다: {err}")
    st.caption("한글에서 '다른 이름으로 저장 → HWPX'로 저장한 파일이 가장 잘 읽힙니다.")
    st.stop()

if st.session_state.get(f"{SS}_key") != key:                    # 새 파일이면 편집본 초기화
    st.session_state[f"{SS}_key"] = key
    st.session_state[f"{SS}_dsl"] = auto_dsl

left, right = st.columns([3, 2], gap="large")

with left:
    st.subheader("중간 텍스트")
    st.caption("자동 인식이 어긋난 곳을 여기서 고치면 그대로 문서에 반영됩니다.")
    st.session_state[f"{SS}_dsl"] = st.text_area(
        "중간 텍스트", value=st.session_state[f"{SS}_dsl"], height=520,
        label_visibility="collapsed")
    if st.button("자동 인식 결과로 되돌리기"):
        st.session_state[f"{SS}_dsl"] = auto_dsl
        st.rerun()

with right:
    st.subheader("인식 결과")
    blocks = parse_dsl(st.session_state[f"{SS}_dsl"])
    found = summarize(blocks)

    a, b = st.columns(2)
    a.metric("문단", stats["문단"])
    b.metric("표", stats["표"])
    c, d = st.columns(2)
    c.metric("문항", found["문항"])
    d.metric("섹션", found["섹션"])
    e, f = st.columns(2)
    e.metric("매트릭스 표", found["매트릭스 표"])
    f.metric("매트릭스 세부항목", found["매트릭스 세부항목"])

    if found["문항"] == 0:
        st.warning("문항을 찾지 못했습니다. 왼쪽에서 문항 줄을 `1. 문항 [단일]` "
                   "형태로 맞춰주세요.")

    docx_bytes = SurveyWriter(
        font=font, base_pt=base_pt, single_mark=single_mark, multi_mark=multi_mark,
        row_label_cm=row_label_cm, accent=accent.lstrip("#").upper()
    ).write(blocks).to_bytes()

    stem = os.path.splitext(uploaded.name)[0]
    st.download_button("워드 파일 내려받기", docx_bytes, f"{stem}.docx",
                       "application/vnd.openxmlformats-officedocument."
                       "wordprocessingml.document",
                       type="primary", use_container_width=True)
    st.download_button("중간 텍스트 내려받기", st.session_state[f"{SS}_dsl"],
                       f"{stem}.txt", "text/plain", use_container_width=True)
    st.download_button("문항 목록 CSV 내려받기", items_csv(blocks),
                       f"{stem}_문항목록.csv", "text/csv",
                       use_container_width=True,
                       help="변수 라벨 작업에 넘겨 쓸 수 있는 문항·보기 목록입니다.")

    with st.expander("문법 도움말"):
        st.markdown(DSL_HELP)

# -*- coding: utf-8 -*-
"""한글 설문지(.hwp/.hwpx) -> 워드(.docx) 변환 페이지.

두 가지 출력 서식을 고를 수 있다.
  * ISAS 표준   : 부서 표준 서식. 납품본에서 실측한 값으로 재현
  * DP 스크립트 : 리서치사 납품용. 조사 개요표 + SQ/Q 번호 + [PROG] 지시문

문서 처리 로직은 hwp_survey 패키지에 있고, 이 파일은 화면만 담당한다.

Streamlit 멀티페이지에서는 각 페이지 스크립트가 독립적으로 실행된다.
Home.py 의 인증은 여기까지 미치지 않으므로, 이 파일에서도 직접
utils.check_password() 를 호출해야 한다.
"""

import csv
import io
import os
import tempfile

import streamlit as st

import utils
from hwp_survey import read_survey
from hwp_survey.dp import DPWriter, items_to_dp_dsl, parse_dp, summarize_dp
from hwp_survey.isas import (ISASWriter, items_to_isas_dsl, parse_isas,
                             summarize_isas)
try:                                     # 검증 기능은 선택 사항
    from hwp_survey.verify import compare, docx_text, hwp_text, pdf_text
    VERIFY_READY = True
except ImportError as err:
    VERIFY_READY = False
    VERIFY_ERROR = str(err)

PAGE_TITLE = "설문지 변환 (한글 → 워드)"
SS = "survey_docx"          # 다른 페이지와 섞이지 않도록 세션 키 접두어

st.set_page_config(page_title=PAGE_TITLE, page_icon="📋", layout="wide")

# 인증 통과 전에는 아래 코드가 한 줄도 실행되지 않아야 한다.
# (사이드바 위젯이나 업로더가 먼저 그려지면 화면이 잠깐 노출된다.)
if not utils.check_password():
    st.stop()

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

HELP_ISAS = """
DP 스크립트와 같은 문법을 쓰고, 번호와 태그만 ISAS 관행으로 바뀝니다.

| 표기 | 뜻 |
|---|---|
| `## PART 1. 제목` | 구역 제목 (본문 번호 `Q1-1`의 기준) |
| `SQ1. 문항 [1개 선택]` | 선별 문항 |
| `Q1-1. 문항 [행별 1개 선택]` | 본문 문항 |
| `@행별: 전혀 그렇지 않다,…` | 행별 표의 보기 라벨 (표 안은 코드로 채움) |
| `%PROG: …` | 파란색 프로그래밍 지시문 |
| `%검증: …` | 초록색 `[DATA: …]` 지시문 |
| `~ 상자글` / `! 안내` | 테두리 상자 / 상자 안 보조 줄 |

번호 체계 — 선별 구역 `SQ`, 참여 동의 `AQ`, 본문 구역 k번째 `Qk` 또는
`Qk-1 Qk-2`, 배경 구역 `DQ`. 분류되지 않은 줄은 **빨간색**으로 남습니다(규칙 2-8).
"""

@st.cache_data(show_spinner=False)
def extract(file_bytes: bytes, suffix: str, tighten: bool, style: str,
            matrix_hint: bool, alone_prog: bool, renumber: bool = False):
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
    if style == "ISAS 표준":
        return items_to_isas_dsl(items, renumber=renumber), stats
    return items_to_dp_dsl(items, add_matrix_hint=matrix_hint,
                           add_alone_prog=alone_prog), stats


def to_csv(rows) -> bytes:
    buf = io.StringIO()
    csv.writer(buf).writerows(rows)
    return buf.getvalue().encode("utf-8-sig")       # 엑셀에서 한글 깨지지 않게


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


def help_text(style):
    return HELP_ISAS if style == "ISAS 표준" else HELP_DP


with st.sidebar:
    st.subheader("출력 서식")
    style = st.radio("서식", ["ISAS 표준", "DP 스크립트"], index=0,
                     captions=["부서 표준 서식", "리서치사 납품용 스크립트"],
                     label_visibility="collapsed")
    st.divider()

    if style == "ISAS 표준":
        st.subheader("조사 개요표")
        st.caption("원본에 없는 정보라 직접 입력합니다.")
        target = st.text_input("조사대상", "")
        sample = st.text_input("샘플 수(명)", "")
        quota_note = st.text_input("할당 설명", "")
        quota_raw = st.text_area(
            "할당 격자 (줄마다 한 행, 쉼표로 칸 구분)", "",
            placeholder=",20-29세,30-39세,합계\n여성,40,40,80", height=90)
        verify = st.text_input("데이터 검증 지시문", "",
                               placeholder="예: Q2~Q7 모두 동일 값 응답시 데이터에서 제외")

        st.subheader("문서 서식")
        font = st.selectbox("글꼴", ["나눔고딕", "맑은 고딕", "함초롬돋움"], index=0)
        base_pt = st.slider("글자 크기(pt)", 8.0, 11.0, 9.0, 0.5)
        row_label_cm = st.slider("행별 표 문항열 너비(cm)", 6.0, 12.0, 8.82, 0.01)
        doc_header = st.text_input("머리말(문서 상단)", "",
                                  placeholder="비우면 설문 제목이 들어갑니다")

        st.subheader("자동 처리")
        split_matrix = st.slider(
            "행별 표 최대 행 수 (0 = 쪼개지 않음)", 0, 40, 0, 1,
            help="이 행 수를 넘는 행별 표를 여러 문항으로 등분합니다. "
                 "예: 20으로 두면 24행 표가 12행 + 12행 두 문항이 됩니다.")
        renumber = st.toggle(
            "문항 번호 다시 매기기 (SQ / Q1-1 / DQ)", value=False,
            help="끄면 원본 번호(문1, A3-2, B0 …)를 그대로 씁니다. "
                 "[PROG] 지시문이 원본 번호를 가리키므로 기본값은 끄기입니다.")
        matrix_hint = alone_prog = False
    elif style == "DP 스크립트":
        renumber = False
        split_matrix = 0
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
        st.markdown(help_text(style))
    st.stop()

suffix = os.path.splitext(uploaded.name)[1].lower()
key = (f"{uploaded.name}:{uploaded.size}:{tighten}:{style}"
       f":{matrix_hint}:{alone_prog}:{renumber}")

try:
    with st.spinner("한글 파일에서 문단과 표를 읽는 중"):
        auto_dsl, stats = extract(uploaded.getvalue(), suffix, tighten, style,
                                  matrix_hint, alone_prog, renumber)
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

    if style == "ISAS 표준":
        doc = parse_isas(dsl, split_matrix=split_matrix)
        doc["대상자"] = target or doc["대상자"]
        doc["샘플수"] = sample or doc["샘플수"]
        doc["쿼터"] = quota_note or doc["쿼터"]
        if quota_raw.strip():
            doc["쿼터표"] = [[c.strip() for c in row.split(",")]
                           for row in quota_raw.strip().splitlines() if row.strip()]
        if verify:
            doc["blocks"].append({"kind": "verify", "text": verify})

        found = summarize_isas(doc)
        c, d = st.columns(2)
        c.metric("문항", found["문항"])
        d.metric("선별문항(SQ)", found["선별문항(SQ)"])
        e, f = st.columns(2)
        e.metric("행별 표", found["행별 표"])
        f.metric("PROG 지시문", found["PROG 지시문"])
        st.metric("그대로 옮긴 표", found["일반 표"])
        empty = found["문항"] == 0
        docx_bytes = ISASWriter(font=font, base_pt=base_pt,
                                row_label_cm=row_label_cm,
                                doc_header=doc_header).write(doc).to_bytes()
        csv_bytes = items_csv_dp(doc)
    elif style == "DP 스크립트":
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
        st.markdown(help_text(style))

st.divider()
st.subheader("변환 누락 검증")
if not VERIFY_READY:
    st.info(f"검증 기능을 불러오지 못했습니다 ({VERIFY_ERROR}). "
            "hwp_survey/verify.py 와 requirements.txt 의 pypdf 를 확인해 주세요.")
    st.stop()
st.caption("한글에서 '다른 이름으로 저장 → PDF'로 저장한 원본을 올리면, 변환 결과와 "
           "문장 단위로 대조해 빠진 내용을 찾습니다.")

pdf_file = st.file_uploader("원본 PDF (선택)", type=["pdf"], key=f"{SS}_pdf")
use_parser = st.checkbox(
    "PDF 없이 검사 (파서가 읽은 텍스트를 기준으로)", value=False,
    help="파서가 애초에 놓친 내용은 이 방식으로 검출되지 않습니다. "
         "렌더링 단계의 누락만 잡힙니다.")

if pdf_file is not None or use_parser:
    if pdf_file is not None:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(pdf_file.getvalue())
            src_path = tmp.name
        try:
            source = pdf_text(src_path)
        finally:
            os.unlink(src_path)
    else:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(uploaded.getvalue())
            src_path = tmp.name
        try:
            source = hwp_text(src_path)
        finally:
            os.unlink(src_path)

    result = compare(source, docx_text(io.BytesIO(docx_bytes)))
    v1, v2, v3 = st.columns(3)
    v1.metric("대조 문장", result["대조 문장"])
    v2.metric("부분 일치", result["부분 일치"])
    v3.metric("누락", result["누락"], delta=f"{result['누락률']:.1%}",
              delta_color="inverse")

    if result["누락"]:
        st.warning("아래 내용이 변환 결과에서 발견되지 않았습니다. 확인해 주세요.")
        st.dataframe({"누락된 문장": result["누락 목록"]},
                     use_container_width=True, hide_index=True)
    else:
        st.success("원본의 모든 문장이 변환 결과에서 확인되었습니다.")

    if result["부분 일치"]:
        with st.expander(f"부분 일치 {result['부분 일치']}건 "
                         "(지시문 분리·문구 변형이 대부분입니다)"):
            st.dataframe({"문장": result["부분 일치 목록"]},
                         use_container_width=True, hide_index=True)

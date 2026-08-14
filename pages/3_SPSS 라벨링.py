"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : pages/3___SPSS_라벨링.py                                         ║
║  위치   : pages/ 폴더 안                                                  ║
║                                                                          ║
║  이 파일은 화면(UI) 코드만 담습니다. 파싱·내보내기 로직은                    ║
║  최상단 spss_labels.py 에 있습니다.                                        ║
╚══════════════════════════════════════════════════════════════════════════╝

워드 설문지를 올리면 문항번호·보기코드를 읽어 코드북을 만들고,
검수 후 SPSS 구문(.sps) 또는 데이터(.sav)로 내려준다.

세션 키는 모두 SL_ 접두사를 쓴다 (다른 페이지와 충돌 방지).
"""

import streamlit as st

import spss_labels as sl
import utils

# 파일 내용이 뒤섞이는 사고를 즉시 잡아낸다.
assert utils.MODULE_ROLE == "utils", "utils.py 가 아닌 파일이 import 되었습니다."
assert sl.MODULE_ROLE == "spss_labels", "spss_labels.py 가 아닌 파일이 import 되었습니다."

st.set_page_config(page_title="SPSS 라벨링", page_icon="📋", layout="wide")

if not utils.check_password():
    st.stop()

# 세션 키
K_VARS = "SL_vars"      # 파싱 결과 (list[Var])
K_EDIT = "SL_edit"      # 편집 중인 코드북 (DataFrame)
K_SRC = "SL_src"        # 원본 파일명
K_FILES = "SL_files"    # 생성된 산출물 {파일명: bytes}
K_REPORT = "SL_report"  # 라벨 적용 결과


def ss(key, default=None):
    """st.session_state.get() 대신 사용. 테스트 하네스에서도 동일하게 동작한다."""
    return st.session_state[key] if key in st.session_state else default


def wide():
    """전체 폭 인자. Streamlit 1.49+ 는 width='stretch', 이전은 use_container_width."""
    import inspect

    try:
        params = inspect.signature(st.download_button).parameters
    except (TypeError, ValueError):
        return {"use_container_width": True}
    return {"width": "stretch"} if "width" in params else {"use_container_width": True}


WIDE = wide()


# ==============================================================================
# 0. 캐시 : 같은 파일을 다시 파싱하지 않는다
# ==============================================================================
@st.cache_data(show_spinner=False)
def parse_cached(docx_bytes, base0_key, full_labels):
    """base0_key 는 캐시 키로 쓰기 위해 tuple 로 받는다."""
    return sl.parse_upload(docx_bytes, base0=list(base0_key), full_labels=full_labels)


def reset_state():
    for k in (K_VARS, K_EDIT, K_SRC, K_FILES, K_REPORT):
        st.session_state.pop(k, None)
    parse_cached.clear()


# ==============================================================================
# 1. 설문지 업로드
# ==============================================================================
st.title("📋 SPSS 라벨링")
st.caption("워드 설문지의 문항번호·보기코드를 읽어 코드북을 만들고, SPSS 구문(.sps) 또는 데이터(.sav)로 내보냅니다.")

with st.sidebar:
    st.header("파싱 설정")
    base0_raw = st.text_input(
        "코드를 0부터 부여할 변수 접두사",
        value="",
        placeholder="q14_1, q14_2",
        help="설문지 표에 코드가 표기되지 않은 척도는 1부터 부여됩니다. "
             "0부터 시작하는 척도라면 접두사를 쉼표로 구분해 적으세요.",
    )
    full_labels = st.checkbox(
        "긴 보기 라벨 원문 유지", value=False,
        help="'분노: 설명…' 형태를 '분노'로 줄이지 않습니다. "
             "SPSS 값라벨은 120바이트(한글 40자) 한도에서 잘립니다.",
    )
    st.divider()
    st.markdown(
        "**처리 순서**\n\n"
        "1. 설문지 업로드 → 자동 파싱\n"
        "2. 코드북 검수 (확인필요 항목)\n"
        "3. 형식 선택 후 다운로드"
    )

base0 = tuple(p.strip() for p in base0_raw.replace(",", " ").split() if p.strip())

st.subheader("1. 파일 업로드")
c1, c2 = st.columns(2)
with c1:
    up_docx = st.file_uploader("워드 설문지 (.docx)", type=["docx"], key="SL_up_docx")
    if up_docx is not None and st.button("설문지 파싱", type="primary"):
        with st.spinner("설문지를 읽고 있습니다…"):
            try:
                variables = parse_cached(up_docx.getvalue(), base0, full_labels)
            except Exception as e:
                st.error(f"파싱 실패: {type(e).__name__}: {e}")
                st.stop()
        st.session_state[K_VARS] = variables
        st.session_state[K_EDIT] = sl.vars_to_frame(variables)
        st.session_state[K_SRC] = up_docx.name
        st.session_state.pop(K_FILES, None)

with c2:
    up_cb = st.file_uploader(
        "이미 검수한 코드북 (.xlsx)", type=["xlsx"], key="SL_up_cb",
        help="설문지 대신 코드북을 올려 산출물만 다시 받을 수도 있습니다.",
    )
    if up_cb is not None and st.button("코드북 불러오기"):
        try:
            variables = sl.codebook_upload_to_vars(up_cb.getvalue())
        except Exception as e:
            st.error(f"코드북 읽기 실패: {type(e).__name__}: {e}")
            st.stop()
        st.session_state[K_VARS] = variables
        st.session_state[K_EDIT] = sl.vars_to_frame(variables)
        st.session_state[K_SRC] = up_cb.name
        st.session_state.pop(K_FILES, None)

if K_EDIT not in st.session_state:
    st.info("설문지(.docx) 또는 검수된 코드북(.xlsx)을 올리면 시작합니다.")
    st.stop()


# ==============================================================================
# 2. 코드북 검수
# ==============================================================================
st.subheader("2. 코드북 검수")

summary = sl.summarize(sl.frame_to_vars(st.session_state[K_EDIT]))
m1, m2, m3 = st.columns(3)
m1.metric("변수", summary["total"])
m2.metric("값라벨 보유", summary["with_labels"])
m3.metric("확인필요", len(summary["todo"]))
st.write("**문항유형** " + " · ".join(f"{k} {n}" for k, n in summary["kinds"].items()))

if summary["todo"]:
    todo = sorted(summary["todo"], key=utils.natural_key)
    st.warning(
        "자동 판정이 불확실한 변수입니다. 값라벨을 확인하세요 — "
        + ", ".join(todo[:20]) + (" …" if len(todo) > 20 else "")
    )
    if st.checkbox("확인필요 항목만 보기"):
        df_all = st.session_state[K_EDIT]
        st.dataframe(df_all[df_all["변수명"].isin(todo)], **WIDE)

st.caption("값라벨 형식 `1=남성 | 2=여성` · 결측값 `99` 또는 `98,99` 또는 범위 `90-99`")
edited = st.data_editor(
    st.session_state[K_EDIT],
    height=420,
    num_rows="dynamic",
    column_config={
        "변수라벨": st.column_config.TextColumn(width="large"),
        "값라벨": st.column_config.TextColumn(width="large"),
        "유형": st.column_config.SelectboxColumn(options=["numeric", "string"]),
        "측도": st.column_config.SelectboxColumn(options=["nominal", "ordinal", "scale"]),
        "문항번호": st.column_config.TextColumn(disabled=True),
        "문항유형": st.column_config.TextColumn(disabled=True),
    },
    key="SL_editor",
    **WIDE,
)
st.session_state[K_EDIT] = edited
variables = sl.frame_to_vars(edited)


# ==============================================================================
# 3. 내보내기
# ==============================================================================
st.subheader("3. 내보내기")

fmt = st.radio("형식", ["SPSS 구문 (.sps)", "SPSS 데이터 (.sav)", "둘 다"],
               horizontal=True, index=2)
want_sps = fmt != "SPSS 데이터 (.sav)"
want_sav = fmt != "SPSS 구문 (.sps)"

data = None
if want_sav:
    up_data = st.file_uploader(
        "원자료 (csv / xlsx / sav) — 없으면 0케이스 사전 파일로 생성",
        type=["csv", "xlsx", "xls", "sav"], key="SL_up_data",
        help="원자료를 올리면 라벨이 적용된 .sav 가 나옵니다. 올리지 않으면 "
             "SPSS 에서 APPLY DICTIONARY 로 씌울 수 있는 사전 파일이 나옵니다.",
    )
    if up_data is not None:
        if up_data.name.lower().endswith(".sav"):
            # utils.load_df 는 sav 를 다루지 않으므로 이 경로만 별도 처리
            try:
                data = sl.read_data_upload(up_data.getvalue(), up_data.name)
            except Exception as e:
                st.error(f"sav 로드 실패: {type(e).__name__}: {e}")
                data = None
        else:
            data = utils.load_df(up_data)   # 실패 시 None 반환 + st.error 표시
        if data is not None:
            st.success(f"원자료 {data.shape[0]:,}케이스 × {data.shape[1]}열")

stem = utils.sanitize_sheet_name(
    str(ss(K_SRC, "survey")).rsplit(".", 1)[0].replace("_codebook", "")
)

if st.button("산출물 생성", type="primary"):
    if not variables:
        st.error("변수가 없습니다.")
        st.stop()
    files = {}
    with st.spinner("생성 중…"):
        try:
            if want_sps:
                files[f"{stem}.sps"] = sl.syntax_bytes(variables, source=str(ss(K_SRC, "")))
            if want_sav:
                blob, report = sl.sav_bytes(variables, data=data)
                files[f"{stem}.sav"] = blob
                st.session_state[K_REPORT] = report
            files[f"{stem}_codebook.xlsx"] = sl.codebook_bytes(variables)
        except Exception as e:
            st.error(f"생성 실패: {type(e).__name__}: {e}")
            st.stop()
    st.session_state[K_FILES] = files

files = ss(K_FILES)
if files:
    report = ss(K_REPORT)
    if report and data is not None:
        st.info(f"라벨 적용 {len(report['labeled'])}개")
        if report["missing_in_data"]:
            with st.expander(f"데이터에 없는 변수 {len(report['missing_in_data'])}개"):
                st.write(", ".join(sorted(report["missing_in_data"], key=utils.natural_key)))
        if report["unlabeled_in_data"]:
            with st.expander(f"코드북에 없는 데이터 열 {len(report['unlabeled_in_data'])}개 (라벨 없이 보존)"):
                st.write(", ".join(report["unlabeled_in_data"]))

    MIMES = {
        "sps": "text/plain",
        "sav": "application/octet-stream",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    cols = st.columns(len(files) + 1)
    for col, (name, blob) in zip(cols, files.items()):
        ext = name.rsplit(".", 1)[-1]
        col.download_button(f"⬇ {ext.upper()}", blob, file_name=name,
                            mime=MIMES.get(ext, "application/octet-stream"), **WIDE)
    cols[-1].download_button("⬇ 전체 ZIP", sl.zip_bytes(files),
                             file_name=f"{stem}_spss.zip", mime="application/zip", **WIDE)

    sps_blob = next((b for n, b in files.items() if n.endswith(".sps")), None)
    if sps_blob:
        with st.expander("구문 미리보기"):
            st.code(sps_blob.decode("utf-8-sig")[:4000], language="text")

st.divider()
if st.button("초기화"):
    reset_state()
    st.rerun()

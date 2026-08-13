"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : pages/7___세로_가로_변환.py                                      ║
║  위치   : pages/ 폴더 안                                                  ║
║                                                                          ║
║  화면(UI) 전용. 변환 로직은 utils.long_to_wide 에 있다.                     ║
║  utils.py 는 2.1-longwide 이상이어야 한다.                                 ║
╚══════════════════════════════════════════════════════════════════════════╝

세로로 쌓인 데이터(long)를 선택한 변수만 가로(wide)로 펼친다.

    ID  항목    값                ID  키   몸무게
    1   키      170     ──▶       1   170  65
    1   몸무게  65                2   160  50
    2   키      160
    2   몸무게  50
"""

import io

import pandas as pd
import streamlit as st

import utils

# utils.py 가 아닌 다른 파일이 import 되는 사고를 즉시 잡는다
assert getattr(utils, "MODULE_ROLE", None) == "utils", \
    "utils.py 가 아닌 모듈이 import 되었습니다."

st.set_page_config(page_title="세로 → 가로 변환", page_icon="↔️", layout="wide")

if not utils.check_password():
    st.stop()

if not hasattr(utils, "long_to_wide"):
    st.error(
        f"utils.py 가 구버전입니다 (현재 {getattr(utils, '__version__', '?')}). "
        "long_to_wide 가 포함된 2.1-longwide 이상으로 교체하세요."
    )
    st.stop()

st.title("↔️ 세로 → 가로 변환")
st.caption(
    "한 응답자의 값이 여러 행에 세로로 쌓여 있는 데이터를, "
    "응답자 1명 = 1행 형태로 펼칩니다. 필요한 변수만 골라서 변환할 수 있습니다."
)


# ==============================================================================
# 파일 읽기 (시트·머리글 행 선택을 지원하므로 load_df 대신 직접 읽는다)
# ==============================================================================
@st.cache_data(show_spinner=False)
def list_sheets(data: bytes, name: str):
    if name.lower().endswith(".csv"):
        return ["(CSV)"]
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner="파일을 읽는 중…")
def read_table(data: bytes, name: str, sheet: str, header: int):
    """실패 시 예외를 raise 한다. 인코딩 폴백은 utils.load_df 와 동일한 순서."""
    if name.lower().endswith(".csv"):
        import chardet
        enc = chardet.detect(data)["encoding"] or "utf-8"
        for cand in (enc, "utf-8-sig", "cp949", "euc-kr"):
            try:
                return pd.read_csv(io.BytesIO(data), header=header, encoding=cand)
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 판별하지 못했습니다. utf-8 로 저장 후 다시 올려주세요.")
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=header)


up = st.file_uploader(
    "데이터 파일 (엑셀 또는 CSV)",
    type=["xlsx", "xlsm", "xls", "csv"],
    key="lw_file",
)
if up is None:
    with st.expander("어떤 데이터를 넣어야 하나요?"):
        st.markdown(
            "**세로(long) 형태** — 이런 데이터를 넣습니다.\n\n"
            "| 학번 | 항목 | 값 |\n|---|---|---|\n"
            "| 101 | 국어 | 88 |\n| 101 | 수학 | 92 |\n| 102 | 국어 | 75 |\n\n"
            "**가로(wide) 형태** — 이렇게 바뀝니다.\n\n"
            "| 학번 | 국어 | 수학 |\n|---|---|---|\n| 101 | 88 | 92 |\n| 102 | 75 |  |"
        )
    st.stop()

raw = up.getvalue()

try:
    sheets = list_sheets(raw, up.name)
except Exception as e:
    st.error(f"파일을 열 수 없습니다: {type(e).__name__}: {e}")
    st.stop()

c1, c2 = st.columns([3, 1])
with c1:
    sheet = st.selectbox("시트", sheets, key="lw_sheet",
                         disabled=(len(sheets) == 1))
with c2:
    header_row = st.number_input(
        "머리글 행", min_value=1, max_value=100, value=1, step=1,
        key="lw_header", help="열 이름이 들어 있는 행 번호. 보통 1행입니다.",
    )

try:
    df = read_table(raw, up.name, sheet, int(header_row) - 1)
except Exception as e:
    st.error(f"읽기 실패: {type(e).__name__}: {e}")
    st.stop()

if df is None or df.empty:
    st.error("데이터가 비어 있습니다. 시트와 머리글 행을 확인하세요.")
    st.stop()

df.columns = [str(c).strip() for c in df.columns]

# 중복 열 이름은 pivot 단계에서 조용히 문제를 일으키므로 먼저 막는다
dup_cols = [c for c, n in pd.Series(df.columns).value_counts().items() if n > 1]
if dup_cols:
    st.error(
        f"열 이름이 중복됩니다: {dup_cols}\n\n"
        "엑셀에서 열 이름을 서로 다르게 고친 뒤 다시 올려주세요."
    )
    st.stop()

st.success(f"{len(df):,}행 × {len(df.columns)}열 읽음")
with st.expander("원본 미리보기", expanded=False):
    st.dataframe(df.head(20))

cols = list(df.columns)

# ==============================================================================
# 열 지정
# ==============================================================================
st.divider()
st.subheader("1. 열 지정")

left, right = st.columns(2)
with left:
    id_cols = st.multiselect(
        "① ID 열 — 한 행을 식별하는 기준",
        cols, key="lw_id",
        help="응답자ID, 학번 등. 여러 개 선택하면 조합으로 식별합니다.",
    )
    key_col = st.selectbox(
        "③ 기준 열 — 가로로 펼칠 변수명이 담긴 열",
        ["(선택하세요)"] + cols, key="lw_key",
        help="'항목', '문항번호', '시점' 처럼 변수 이름이 세로로 쌓여 있는 열",
    )
with right:
    value_cols = st.multiselect(
        "② 값 열 — 실제 값이 담긴 열",
        cols, key="lw_val",
        help="2개 이상 고르면 열 이름이 '점수_1차', '시간_1차' 형태로 조합됩니다.",
    )
    agg_label = st.selectbox(
        "④ 중복값 처리 — 같은 ID에 같은 변수가 2번 이상일 때",
        list(utils.AGG_FUNCS.keys()), key="lw_agg",
    )

if key_col == "(선택하세요)":
    st.info("③ 기준 열을 선택하면 펼칠 변수 목록이 나타납니다.")
    st.stop()

norm_keys = st.checkbox(
    "변수명 정규화 (권장)", value=True, key="lw_norm",
    help=f"'1.0' → '1', 공백/결측 → '{utils.NA_TOKEN}'. "
         "utils.norm_val 을 쓰므로 앱의 다른 페이지와 열 이름이 일관됩니다.",
)

# ==============================================================================
# 펼칠 변수 선택
# ==============================================================================
st.divider()
st.subheader("2. 펼칠 변수 선택")

key_series = utils.norm_series(df[key_col]) if norm_keys else df[key_col].astype(str)
uniq = list(pd.unique(key_series.dropna()))

if len(uniq) > 300:
    st.warning(
        f"'{key_col}' 열에 서로 다른 값이 {len(uniq):,}개 있습니다. "
        "기준 열이 아니라 값 열을 고르신 게 아닌지 확인해 주세요."
    )

sc1, sc2 = st.columns([1, 3])
with sc1:
    sort_mode = st.radio(
        "목록 순서", ["원본 등장 순", "자연 정렬(q2 < q10)"],
        key="lw_sort", horizontal=False,
    )
if sort_mode.startswith("자연"):
    uniq = sorted(uniq, key=utils.natural_key)

with sc2:
    counts = key_series.value_counts()
    st.caption(f"총 {len(uniq):,}개 변수. 괄호 안은 해당 변수의 행 개수입니다.")
    labels = {f"{k}  ({counts.get(k, 0):,}행)": k for k in uniq}
    all_labels = list(labels.keys())

    # 전체선택/해제는 반드시 on_click 콜백으로 처리한다.
    # 스크립트 본문에서 st.session_state["lw_keys"] 를 직접 대입하면
    # 위젯이 이미 생성된 뒤라 StreamlitAPIException 이 발생한다.
    # 콜백은 다음 실행 시작 시점(위젯 생성 전)에 돌기 때문에 안전하다.
    def _pick_all(opts=all_labels):
        st.session_state["lw_keys"] = list(opts)

    def _pick_none():
        st.session_state["lw_keys"] = []

    b1, b2 = st.columns(2)
    b1.button("전체 선택", on_click=_pick_all)
    b2.button("전체 해제", on_click=_pick_none)

    # 기준 열을 바꾸면 옵션 목록이 완전히 달라진다. 이전 선택값이 새 옵션에
    # 없으면 multiselect 가 "default value must exist in options" 로 죽으므로
    # 위젯 생성 전에 미리 걸러낸다. (default= 인자는 쓰지 않는다. session_state
    # 와 같이 쓰면 경고가 나고 default 가 무시된다.)
    if "lw_keys" not in st.session_state:
        st.session_state["lw_keys"] = all_labels if len(uniq) <= 30 else []
    else:
        valid = [l for l in st.session_state["lw_keys"] if l in labels]
        if len(valid) != len(st.session_state["lw_keys"]):
            st.session_state["lw_keys"] = valid

    picked_labels = st.multiselect(
        "가로로 펼칠 변수 (선택한 순서대로 열이 배치됩니다)",
        all_labels,
        key="lw_keys",
    )
    keep_keys = [labels[l] for l in picked_labels if l in labels]

# ==============================================================================
# 중복 사전 점검
# ==============================================================================
if id_cols and keep_keys:
    try:
        dup, dup_total = utils.find_duplicate_cells(
            df[df[key_col].isin(keep_keys) if not norm_keys
               else key_series.isin(keep_keys)],
            id_cols, key_col, normalize_keys=norm_keys,
        )
    except Exception:
        dup, dup_total = pd.DataFrame(), 0

    if dup_total:
        st.warning(
            f"같은 ID에 같은 변수가 중복된 칸이 {len(dup):,}곳 있습니다 "
            f"(총 {dup_total:,}행). 현재 설정은 **{agg_label}** 로 처리합니다."
        )
        with st.expander("중복 내역 보기"):
            st.dataframe(dup)

# ==============================================================================
# 변환
# ==============================================================================
st.divider()
st.subheader("3. 변환")

if not id_cols or not value_cols or not keep_keys:
    st.info("ID 열, 값 열, 펼칠 변수를 모두 지정하면 변환 버튼이 활성화됩니다.")
    st.stop()

n_expected = len(id_cols) + len(value_cols) * len(keep_keys)
st.caption(f"예상 결과: ID {len(id_cols)}열 + 값 {len(value_cols)}개 × "
           f"변수 {len(keep_keys)}개 = 총 {n_expected}열")

if st.button("변환 실행", type="primary"):
    try:
        with st.spinner("변환 중…"):
            result = utils.long_to_wide(
                df,
                id_cols=id_cols,
                key_col=key_col,
                value_cols=value_cols,
                keep_keys=keep_keys,
                aggfunc=utils.AGG_FUNCS[agg_label],
                normalize_keys=norm_keys,
            )
        st.session_state["lw_result"] = result
    except ValueError as e:
        st.error(str(e))
        st.session_state.pop("lw_result", None)
    except Exception as e:
        st.error(f"예상치 못한 오류: {type(e).__name__}: {e}")
        st.session_state.pop("lw_result", None)

result = st.session_state.get("lw_result")
if result is not None:
    st.success(f"변환 완료 — {len(result):,}행 × {len(result.columns)}열")
    st.dataframe(result.head(100))
    if len(result) > 100:
        st.caption(f"위 표는 앞 100행만 표시합니다. 전체 {len(result):,}행은 파일로 받으세요.")

    empty_cols = [c for c in result.columns if result[c].isna().all()]
    if empty_cols:
        st.warning(f"값이 전부 비어 있는 열: {empty_cols}")

    base = up.name.rsplit(".", 1)[0]

    used = set()
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            result.to_excel(w, sheet_name=utils.unique_sheet_name("wide", used),
                            index=False)
    except Exception:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            result.to_excel(w, sheet_name="wide", index=False)

    d1, d2 = st.columns(2)
    d1.download_button(
        "엑셀 다운로드", data=buf.getvalue(),
        file_name=f"{base}_wide.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    d2.download_button(
        "CSV 다운로드", data=result.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{base}_wide.csv", mime="text/csv",
    )

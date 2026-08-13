"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : pages/7___세로_가로_변환.py                                      ║
║  위치   : pages/ 폴더 안                                                  ║
║                                                                          ║
║  ★ 단일 파일 버전 ★                                                       ║
║  utils.py 를 전혀 수정하지 않아도 동작합니다.                                ║
║  이 파일 하나만 pages/ 에 넣으면 끝입니다.                                   ║
║                                                                          ║
║  utils.py 가 있으면 check_password / norm_val 등을 그대로 재사용하고,        ║
║  없거나 구버전이면 이 파일 안의 동일한 구현으로 자동 대체합니다.               ║
╚══════════════════════════════════════════════════════════════════════════╝

세로로 쌓인 데이터(long)를 선택한 변수만 가로(wide)로 펼친다.

    ID  항목    값                ID  키   몸무게
    1   키      170     ──▶       1   170  65
    1   몸무게  65                2   160  50
    2   키      160
    2   몸무게  50

필요 패키지: streamlit, pandas, openpyxl, XlsxWriter, chardet
(모두 기존 requirements.txt 에 이미 포함되어 있음)
"""

import hmac
import io
import re

import pandas as pd
import streamlit as st

# ==============================================================================
# 0. utils.py 연동 (있으면 재사용, 없으면 자체 구현)
#    이 페이지는 utils.py 의 버전에 의존하지 않는다.
# ==============================================================================
try:
    import utils as _u
except Exception:
    _u = None


def _has(name):
    """utils 에 해당 함수가 실제로 있는지 확인."""
    return _u is not None and callable(getattr(_u, name, None))


# 결측/공백 토큰. utils 의 값을 우선 사용해 다른 페이지와 표기를 맞춘다.
NA_TOKEN = getattr(_u, "NA_TOKEN", "(무응답)") if _u else "(무응답)"


def norm_val(v):
    """
    값을 열 이름으로 쓸 수 있는 정규화된 문자열로 변환.
    utils.norm_val 이 있으면 그것을 쓴다 (앱 전체와 동일한 규칙 보장).

    규칙: 결측/공백 -> NA_TOKEN, 앞뒤 공백 제거,
          정수를 float 로 읽은 경우만 ".0" 제거 (1.0 -> "1", 1.5 는 유지)
    """
    if _has("norm_val"):
        return _u.norm_val(v)

    if v is None:
        return NA_TOKEN
    try:
        if pd.isna(v):
            return NA_TOKEN
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if s == "":
        return NA_TOKEN
    if s.endswith(".0") and s[:-2].lstrip("+-").isdigit():
        s = s[:-2]
    return s


def norm_series(s):
    if _has("norm_series"):
        return _u.norm_series(s)
    return pd.Series(s).map(norm_val)


def natural_key(string_):
    """'q10' 이 'q9' 뒤에 오도록 정렬."""
    if _has("natural_key"):
        return _u.natural_key(string_)
    return [int(x) if x.isdigit() else x.lower()
            for x in re.split(r"(\d+)", str(string_))]


def sanitize_sheet_name(name):
    if _has("sanitize_sheet_name"):
        return _u.sanitize_sheet_name(name)
    safe = re.sub(r"[\\/*?:\[\]]", "_", str(name)).strip("'") or "Sheet"
    return safe[:28] + ".." if len(safe) > 30 else safe


def check_password():
    """
    utils.check_password 가 있으면 그대로 사용 (기존 앱과 동일한 게이트).
    없을 때만 동일 동작의 자체 구현을 쓴다.
    """
    if _has("check_password"):
        return _u.check_password()

    def entered():
        try:
            secret = st.secrets["password"]
        except Exception:
            st.session_state["lw_pw_ok"] = False
            st.session_state["lw_pw_msg"] = (
                "서버에 비밀번호가 설정되어 있지 않습니다. "
                ".streamlit/secrets.toml 의 password 항목을 확인하세요.")
            return
        # compare_digest 는 비ASCII 문자열을 그대로 넘기면 TypeError 를 낸다.
        # (한글 비밀번호를 입력하면 앱이 죽는다) 반드시 bytes 로 비교한다.
        entered = str(st.session_state.get("lw_pw", "")).encode("utf-8")
        if hmac.compare_digest(entered, str(secret).encode("utf-8")):
            st.session_state["lw_pw_ok"] = True
            st.session_state["lw_pw_msg"] = None
            del st.session_state["lw_pw"]
        else:
            st.session_state["lw_pw_ok"] = False
            st.session_state["lw_pw_msg"] = "비밀번호가 올바르지 않습니다."

    # 다른 페이지에서 이미 인증했으면 통과
    if st.session_state.get("password_correct", False) or \
            st.session_state.get("lw_pw_ok", False):
        return True

    st.title("🔒 접속 제한")
    st.text_input("비밀번호를 입력하세요", type="password",
                  on_change=entered, key="lw_pw")
    msg = st.session_state.get("lw_pw_msg")
    if msg:
        st.error(msg)
    return False


# ==============================================================================
# 1. 변환 로직 (전부 이 파일 안에 있다 — utils.py 와 무관)
# ==============================================================================
AGG_FUNCS = {
    "첫 번째 값만": "first",
    "마지막 값만": "last",
    "평균": "mean",
    "합계": "sum",
    "최댓값": "max",
    "최솟값": "min",
    "개수": "count",
    "쉼표로 이어붙이기": lambda s: ", ".join(str(v) for v in s.dropna()),
}


def _restore_ints(df):
    """
    소수점이 필요 없는 실수 열을 정수(Int64)로 되돌린다.
    pivot 이 int 를 float 로 승격시켜 '170.0' 으로 보이는 문제를 없앤다.
    """
    for c in df.columns:
        s = df[c]
        if s.dtype.kind == "f":
            nn = s.dropna()
            if not nn.empty and (nn % 1 == 0).all():
                try:
                    df[c] = s.astype("Int64")
                except (TypeError, ValueError):
                    pass
    return df


def long_to_wide(df, id_cols, key_col, value_cols, keep_keys=None,
                 aggfunc="first", name_sep="_", normalize_keys=True):
    """
    세로(long) → 가로(wide) 변환. 실패 시 ValueError 를 raise 한다.

    id_cols    : 한 행을 식별하는 열 (예: ["학번"], ["학번","이름"])
    key_col    : 가로로 펼칠 변수명이 담긴 열 (예: "항목", "시점")
    value_cols : 값이 담긴 열. 2개 이상이면 "점수_1차" 처럼 조합된다
    keep_keys  : 펼칠 변수 목록. 준 순서가 결과 열 순서가 된다
    """
    if df is None or df.empty:
        raise ValueError("데이터가 비어 있습니다.")

    id_cols, value_cols = list(id_cols), list(value_cols)
    if not id_cols:
        raise ValueError("ID 열을 1개 이상 선택해야 합니다.")
    if not value_cols:
        raise ValueError("값 열을 1개 이상 선택해야 합니다.")
    if not key_col:
        raise ValueError("기준 열을 선택해야 합니다.")
    if key_col in id_cols or key_col in value_cols:
        raise ValueError(f"기준 열 '{key_col}' 이 ID 열 또는 값 열과 겹칩니다.")
    both = set(id_cols) & set(value_cols)
    if both:
        raise ValueError(f"ID 열과 값 열에 같은 열이 들어갔습니다: {sorted(both)}")
    missing = [c for c in id_cols + value_cols + [key_col] if c not in df.columns]
    if missing:
        raise ValueError(f"데이터에 없는 열입니다: {missing}")

    work = df.loc[:, id_cols + [key_col] + value_cols].copy()
    # 열 이름이 될 값이므로 항상 문자열로 통일. 이걸 빠뜨리면 int 1 과
    # 화면에서 고른 "1" 이 어긋나 "해당 행이 없습니다" 가 뜬다.
    work[key_col] = norm_series(work[key_col]) if normalize_keys \
        else work[key_col].astype("string")

    if keep_keys is None:
        keep_keys = list(pd.unique(work[key_col].dropna()))
    else:
        keep_keys = [norm_val(k) if normalize_keys else str(k) for k in keep_keys]
        keep_keys = list(dict.fromkeys(keep_keys))      # 순서 유지 중복 제거
        work = work[work[key_col].isin(keep_keys)]
        if work.empty:
            raise ValueError("선택한 변수에 해당하는 행이 없습니다.")

    # ID 조합의 원래 등장 순서 보존 (pivot 은 정렬해 버린다)
    order = df.loc[:, id_cols].drop_duplicates().reset_index(drop=True)

    try:
        wide = pd.pivot_table(
            work, index=id_cols, columns=key_col, values=value_cols,
            aggfunc=aggfunc,
            dropna=True,     # False 로 두면 정수가 불필요하게 실수로 승격된다
            observed=True,
        )
    except Exception as e:
        raise ValueError(
            f"피벗 실패 ({type(e).__name__}: {e})\n"
            "값 열에 숫자가 아닌 값이 섞여 있는데 평균/합계를 고른 경우일 수 있습니다. "
            "중복값 처리를 '첫 번째 값만' 으로 바꿔 보세요.")

    if isinstance(wide.columns, pd.MultiIndex):
        if len(value_cols) == 1:
            wide.columns = [str(c[-1]) for c in wide.columns]
            desired = [str(k) for k in keep_keys]
        else:
            wide.columns = [f"{c[0]}{name_sep}{c[1]}" for c in wide.columns]
            desired = [f"{v}{name_sep}{k}" for v in value_cols for k in keep_keys]
    else:
        wide.columns = [str(c) for c in wide.columns]
        desired = [str(k) for k in keep_keys]

    wide = wide.reset_index()

    # 값이 전부 비어 사라진 변수도 빈 열로 되살린다 (열 개수를 예측 가능하게)
    for c in desired:
        if c not in wide.columns:
            wide[c] = pd.NA

    cols = id_cols + list(desired)
    cols += [c for c in wide.columns if c not in cols]
    wide = wide.loc[:, cols]

    wide = order.merge(wide, on=id_cols, how="left")   # ID 원래 순서 복원
    return _restore_ints(wide)


def find_duplicate_cells(df, id_cols, key_col, normalize_keys=True, limit=200):
    """'같은 ID + 같은 변수' 가 2회 이상인 칸을 미리 찾는다."""
    key = norm_series(df[key_col]) if normalize_keys \
        else df[key_col].astype("string")
    tmp = df.loc[:, id_cols].copy()
    tmp[key_col] = key
    g = tmp.groupby(id_cols + [key_col], dropna=False, observed=True).size()
    dup = g[g > 1].reset_index(name="중복 횟수")
    total = int(dup["중복 횟수"].sum()) if not dup.empty else 0
    return dup.head(limit), total


# ==============================================================================
# 2. 화면
# ==============================================================================
st.set_page_config(page_title="세로 → 가로 변환", page_icon="↔️", layout="wide")

if not check_password():
    st.stop()

st.title("↔️ 세로 → 가로 변환")
st.caption("한 응답자의 값이 여러 행에 세로로 쌓여 있는 데이터를, "
           "응답자 1명 = 1행 형태로 펼칩니다. 필요한 변수만 골라서 변환할 수 있습니다.")


@st.cache_data(show_spinner=False)
def list_sheets(data: bytes, name: str):
    if name.lower().endswith(".csv"):
        return ["(CSV)"]
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner="파일을 읽는 중…")
def read_table(data: bytes, name: str, sheet: str, header: int):
    if name.lower().endswith(".csv"):
        try:
            import chardet
            enc = chardet.detect(data)["encoding"] or "utf-8"
        except Exception:
            enc = "utf-8"
        for cand in (enc, "utf-8-sig", "cp949", "euc-kr"):
            try:
                return pd.read_csv(io.BytesIO(data), header=header, encoding=cand)
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 판별하지 못했습니다. utf-8 로 저장 후 올려주세요.")
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=header)


up = st.file_uploader("데이터 파일 (엑셀 또는 CSV)",
                      type=["xlsx", "xlsm", "xls", "csv"], key="lw_file")
if up is None:
    with st.expander("어떤 데이터를 넣어야 하나요?"):
        st.markdown(
            "**세로(long) 형태** — 이런 데이터를 넣습니다.\n\n"
            "| 학번 | 항목 | 값 |\n|---|---|---|\n"
            "| 101 | 국어 | 88 |\n| 101 | 수학 | 92 |\n| 102 | 국어 | 75 |\n\n"
            "**가로(wide) 형태** — 이렇게 바뀝니다.\n\n"
            "| 학번 | 국어 | 수학 |\n|---|---|---|\n| 101 | 88 | 92 |\n| 102 | 75 |  |")
    st.stop()

raw = up.getvalue()
try:
    sheets = list_sheets(raw, up.name)
except Exception as e:
    st.error(f"파일을 열 수 없습니다: {type(e).__name__}: {e}")
    st.stop()

c1, c2 = st.columns([3, 1])
with c1:
    sheet = st.selectbox("시트", sheets, key="lw_sheet", disabled=(len(sheets) == 1))
with c2:
    header_row = st.number_input("머리글 행", min_value=1, max_value=100, value=1,
                                 step=1, key="lw_header",
                                 help="열 이름이 들어 있는 행 번호. 보통 1행입니다.")

try:
    df = read_table(raw, up.name, sheet, int(header_row) - 1)
except Exception as e:
    st.error(f"읽기 실패: {type(e).__name__}: {e}")
    st.stop()

if df is None or df.empty:
    st.error("데이터가 비어 있습니다. 시트와 머리글 행을 확인하세요.")
    st.stop()

df.columns = [str(c).strip() for c in df.columns]

dup_cols = [c for c, n in pd.Series(df.columns).value_counts().items() if n > 1]
if dup_cols:
    st.error(f"열 이름이 중복됩니다: {dup_cols}\n\n"
             "엑셀에서 열 이름을 서로 다르게 고친 뒤 다시 올려주세요.")
    st.stop()

st.success(f"{len(df):,}행 × {len(df.columns)}열 읽음")
with st.expander("원본 미리보기", expanded=False):
    st.dataframe(df.head(20))

cols = list(df.columns)

st.divider()
st.subheader("1. 열 지정")

left, right = st.columns(2)
with left:
    id_cols = st.multiselect("① ID 열 — 한 행을 식별하는 기준", cols, key="lw_id",
                             help="응답자ID, 학번 등. 여러 개면 조합으로 식별합니다.")
    key_col = st.selectbox("③ 기준 열 — 가로로 펼칠 변수명이 담긴 열",
                           ["(선택하세요)"] + cols, key="lw_key",
                           help="'항목', '문항번호', '시점' 처럼 변수 이름이 쌓인 열")
with right:
    value_cols = st.multiselect("② 값 열 — 실제 값이 담긴 열", cols, key="lw_val",
                                help="2개 이상이면 '점수_1차' 형태로 조합됩니다.")
    agg_label = st.selectbox("④ 중복값 처리 — 같은 ID에 같은 변수가 2번 이상일 때",
                             list(AGG_FUNCS.keys()), key="lw_agg")

if key_col == "(선택하세요)":
    st.info("③ 기준 열을 선택하면 펼칠 변수 목록이 나타납니다.")
    st.stop()

norm_keys = st.checkbox(
    "변수명 정규화 (권장)", value=True, key="lw_norm",
    help=f"'1.0' → '1', 공백/결측 → '{NA_TOKEN}'. 다른 페이지와 열 이름이 일관해집니다.")

st.divider()
st.subheader("2. 펼칠 변수 선택")

key_series = norm_series(df[key_col]) if norm_keys else df[key_col].astype("string")
uniq = list(pd.unique(key_series.dropna()))

if len(uniq) > 300:
    st.warning(f"'{key_col}' 열에 서로 다른 값이 {len(uniq):,}개 있습니다. "
               "기준 열이 아니라 값 열을 고르신 게 아닌지 확인해 주세요.")

sc1, sc2 = st.columns([1, 3])
with sc1:
    sort_mode = st.radio("목록 순서", ["원본 등장 순", "자연 정렬(q2 < q10)"], key="lw_sort")
if sort_mode.startswith("자연"):
    uniq = sorted(uniq, key=natural_key)

with sc2:
    counts = key_series.value_counts()
    st.caption(f"총 {len(uniq):,}개 변수. 괄호 안은 해당 변수의 행 개수입니다.")
    labels = {f"{k}  ({counts.get(k, 0):,}행)": k for k in uniq}
    all_labels = list(labels.keys())

    # 전체선택/해제는 반드시 on_click 콜백으로. 본문에서 session_state 를 직접
    # 대입하면 위젯 생성 후 수정이라 StreamlitAPIException 이 발생한다.
    def _pick_all(opts=all_labels):
        st.session_state["lw_keys"] = list(opts)

    def _pick_none():
        st.session_state["lw_keys"] = []

    b1, b2 = st.columns(2)
    b1.button("전체 선택", on_click=_pick_all)
    b2.button("전체 해제", on_click=_pick_none)

    # 기준 열을 바꾸면 옵션이 전부 달라진다. 이전 선택값이 새 옵션에 없으면
    # multiselect 가 죽으므로 위젯 생성 전에 걸러낸다.
    if "lw_keys" not in st.session_state:
        st.session_state["lw_keys"] = all_labels if len(uniq) <= 30 else []
    else:
        valid = [l for l in st.session_state["lw_keys"] if l in labels]
        if len(valid) != len(st.session_state["lw_keys"]):
            st.session_state["lw_keys"] = valid

    picked_labels = st.multiselect(
        "가로로 펼칠 변수 (선택한 순서대로 열이 배치됩니다)", all_labels, key="lw_keys")
    keep_keys = [labels[l] for l in picked_labels if l in labels]

if id_cols and keep_keys:
    try:
        dup, dup_total = find_duplicate_cells(
            df[key_series.isin(keep_keys)], id_cols, key_col,
            normalize_keys=norm_keys)
    except Exception:
        dup, dup_total = pd.DataFrame(), 0
    if dup_total:
        st.warning(f"같은 ID에 같은 변수가 중복된 칸이 {len(dup):,}곳 있습니다 "
                   f"(총 {dup_total:,}행). 현재 설정은 **{agg_label}** 로 처리합니다.")
        with st.expander("중복 내역 보기"):
            st.dataframe(dup)

st.divider()
st.subheader("3. 변환")

if not id_cols or not value_cols or not keep_keys:
    st.info("ID 열, 값 열, 펼칠 변수를 모두 지정하면 변환 버튼이 활성화됩니다.")
    st.stop()

st.caption(f"예상 결과: ID {len(id_cols)}열 + 값 {len(value_cols)}개 × "
           f"변수 {len(keep_keys)}개 = 총 {len(id_cols) + len(value_cols) * len(keep_keys)}열")

if st.button("변환 실행", type="primary"):
    try:
        with st.spinner("변환 중…"):
            st.session_state["lw_result"] = long_to_wide(
                df, id_cols=id_cols, key_col=key_col, value_cols=value_cols,
                keep_keys=keep_keys, aggfunc=AGG_FUNCS[agg_label],
                normalize_keys=norm_keys)
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
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            result.to_excel(w, sheet_name=sanitize_sheet_name("wide"), index=False)
    except Exception:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            result.to_excel(w, sheet_name="wide", index=False)

    d1, d2 = st.columns(2)
    d1.download_button(
        "엑셀 다운로드", data=buf.getvalue(), file_name=f"{base}_wide.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button(
        "CSV 다운로드", data=result.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{base}_wide.csv", mime="text/csv")

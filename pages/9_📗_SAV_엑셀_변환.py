# -*- coding: utf-8 -*-
"""
SAV → 엑셀 변환 (v1.3)

SPSS .sav 파일을 업로드하면 4개 시트로 구성된 엑셀 파일을 내려받습니다.
  · Raw        : 숫자 코드 그대로
  · Label      : 값 레이블로 치환 (레이블 없는 변수는 원래 값 유지)
  · Open       : 키 변수(NO, id) + 문자형(주관식) 변수
                 문자형 변수가 없어도 키 변수만으로 만든다
  · 변수 가이드 : 변수명 + 변수 설명

이 페이지는 utils.py 없이도 단독으로 동작합니다.
"""

import io
import os
import tempfile

import numpy as np
import pandas as pd
import pyreadstat
import streamlit as st
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────
# 비밀번호 (utils.py 있으면 사용, 없으면 통과)
# ──────────────────────────────────────────────────────────────
try:
    from utils import check_password  # type: ignore
except Exception:  # pragma: no cover
    def check_password() -> bool:
        return True


st.set_page_config(page_title="SAV → 엑셀 변환", page_icon="📗", layout="wide")

if not check_password():
    st.stop()

st.title("📗 SAV → 엑셀 변환")
st.caption("SPSS .sav 파일을 Raw / Label / 변수 가이드 3개 시트의 엑셀로 바꿔 드립니다.")


# ──────────────────────────────────────────────────────────────
# 변환 로직
# ──────────────────────────────────────────────────────────────
def _clean(v):
    """NaN은 빈칸으로, 소수점 없는 실수는 정수로."""
    if v is None:
        return None
    if isinstance(v, float):
        if np.isnan(v):
            return None
        if float(v).is_integer():
            return int(v)
    return v


@st.cache_data(show_spinner=False, max_entries=5)
def read_sav_bytes(data: bytes, filename: str):
    """업로드된 바이트를 임시파일로 떨어뜨려 pyreadstat으로 읽는다."""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".sav", delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        df, meta = pyreadstat.read_sav(tmp_path, apply_value_formats=False)
        return (
            df,
            dict(meta.column_names_to_labels or {}),
            dict(meta.variable_value_labels or {}),
            dict(getattr(meta, "readstat_variable_types", {}) or {}),
        )
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def build_raw(df: pd.DataFrame) -> pd.DataFrame:
    return df.map(_clean)


def build_label(df: pd.DataFrame, value_labels: dict) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if c in value_labels:
            m = {k: str(v).strip() for k, v in value_labels[c].items()}
            out[c] = out[c].map(lambda x, m=m: m.get(x, _clean(x)))
        else:
            out[c] = out[c].map(_clean)
    return out


def find_key_cols(df: pd.DataFrame) -> list:
    """응답자를 되짚을 키 열. NO / id 를 원래 순서대로 모으고, 없으면 첫 열."""
    keys = [c for c in df.columns if str(c).strip().lower() in ("no", "id")]
    return keys if keys else [df.columns[0]]


def find_text_cols(df: pd.DataFrame, var_types: dict) -> list:
    """SAV에서 문자형으로 선언된 변수 목록."""
    return [c for c in df.columns if str(var_types.get(c, "")).lower() == "string"]


def build_open(df: pd.DataFrame, text_cols: list, key_cols: list) -> pd.DataFrame:
    """키 변수 + 주관식 응답. 행 순서는 Raw/Label과 동일하게 유지.

    문자형 변수가 없으면 키 변수만 담긴 시트가 된다. 주관식 코딩을 할 때
    이 시트에 열을 직접 추가해 쓸 수 있도록 빈 채로라도 만들어 둔다.
    """
    cols = list(key_cols) + [c for c in text_cols if c not in key_cols]
    out = df[cols].copy()
    for c in out.columns:
        if c in text_cols:
            out[c] = out[c].map(
                lambda x: None
                if x is None or (isinstance(x, float) and np.isnan(x)) or str(x).strip() == ""
                else str(x).strip()
            )
        else:
            out[c] = out[c].map(_clean)
    return out


def build_guide(df: pd.DataFrame, col_labels: dict) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "변수명": list(df.columns),
            "변수 내용": [str(col_labels.get(c) or "").strip() for c in df.columns],
        }
    )


def to_excel(sheets: dict, n_keys: int = 1) -> bytes:
    """{시트명: DataFrame} → 엑셀 바이트. 헤더 굵게 + 첫 행 고정."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            for cell in ws[1]:
                cell.font = Font(bold=True)
            ws.freeze_panes = "A2"
            if name == "변수 가이드":
                ws.column_dimensions["A"].width = 18
                ws.column_dimensions["B"].width = 100
            elif name == "Open":
                # 주관식 본문은 길어서 좁으면 읽기 어렵다. 열 수가 적어 상한을 두지 않는다.
                for j in range(1, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(j)].width = 12 if j <= n_keys else 45
            # Raw / Label 은 너비를 지정하지 않는다.
            # 열이 수백 개까지 늘어날 수 있어, 일부만 지정되면 오히려 들쭉날쭉해진다.
            # 엑셀 기본 너비로 두면 사용자가 전체 선택 후 한 번에 조절할 수 있다.
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────
up = st.file_uploader("SAV 파일을 올려주세요", type=["sav"])

if not up:
    st.info("SPSS .sav 파일을 올리면 미리보기와 다운로드 버튼이 나타납니다.")
    st.stop()

try:
    df, col_labels, value_labels, var_types = read_sav_bytes(up.getvalue(), up.name)
except Exception as e:
    st.error(f"파일을 읽지 못했습니다: {e}")
    st.stop()

n_rows, n_cols = df.shape
n_labeled = sum(1 for c in df.columns if c in value_labels)
text_cols = find_text_cols(df, var_types)
key_cols = find_key_cols(df)

c1, c2, c3, c4 = st.columns(4)
c1.metric("응답자 수", f"{n_rows:,}")
c2.metric("변수 수", f"{n_cols:,}")
c3.metric("값 레이블이 있는 변수", f"{n_labeled:,}")
c4.metric("문자형 변수", f"{len(text_cols):,}")

if n_rows > 10_000:
    st.warning(
        f"행이 {n_rows:,}개입니다. 1만 행이 넘으면 변환이 느리거나 "
        "메모리 한도에 걸릴 수 있습니다."
    )

st.divider()

st.subheader("담을 시트 고르기")
s1, s2, s3, s4 = st.columns(4)
want_raw = s1.checkbox("Raw (숫자 코드)", value=True)
want_label = s2.checkbox("Label (값 레이블)", value=True)
want_open = s3.checkbox(
    "Open (주관식)",
    value=True,
    help="키 변수(" + ", ".join(key_cols) + ")와 문자형 변수를 담습니다.",
)
want_guide = s4.checkbox("변수 가이드", value=True)

if want_open and not text_cols:
    st.info(
        "이 파일에는 문자형 변수가 없어 Open 시트에 키 변수("
        + ", ".join(key_cols)
        + ")만 담깁니다. 주관식 응답을 옆에 붙여 코딩하실 때 쓰시면 됩니다."
    )

if not (want_raw or want_label or want_open or want_guide):
    st.warning("시트를 하나 이상 선택해주세요.")
    st.stop()

# ── 시트 구성 (Raw → Label → Open → 변수 가이드) ──
sheets = {}
if want_raw:
    sheets["Raw"] = build_raw(df)
if want_label:
    sheets["Label"] = build_label(df, value_labels)
if want_open:
    sheets["Open"] = build_open(df, text_cols, key_cols)
if want_guide:
    sheets["변수 가이드"] = build_guide(df, col_labels)

st.subheader("미리보기")
tabs = st.tabs(list(sheets.keys()))
for tab, (name, frame) in zip(tabs, sheets.items()):
    with tab:
        st.dataframe(
            frame.head(20).astype(str).replace("None", ""),
            use_container_width=True,
            hide_index=True,
        )
        if len(frame) > 20:
            st.caption(f"위 20행만 표시 · 전체 {len(frame):,}행")

st.divider()

if st.button("엑셀로 변환하기", type="primary", use_container_width=True):
    with st.spinner("엑셀 파일을 만드는 중입니다…"):
        try:
            xlsx_bytes = to_excel(sheets, n_keys=len(key_cols))
        except Exception as e:
            st.error(f"엑셀 생성에 실패했습니다: {e}")
            st.stop()
    st.session_state["sav2xlsx_bytes"] = xlsx_bytes
    st.session_state["sav2xlsx_name"] = os.path.splitext(up.name)[0] + ".xlsx"

if st.session_state.get("sav2xlsx_bytes"):
    st.success("변환이 끝났습니다.")
    st.download_button(
        "엑셀 파일 내려받기",
        data=st.session_state["sav2xlsx_bytes"],
        file_name=st.session_state["sav2xlsx_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

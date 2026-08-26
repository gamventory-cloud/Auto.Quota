# -*- coding: utf-8 -*-
"""
1___데이터_검증.py  (v1.0)

불성실 응답 / 데이터 품질 검증 페이지.

설계 원칙
  - 자동 삭제하지 않는다. "의심 케이스 + 사유"만 산출하고 판정은 사람이 한다.
  - 검사 함수는 모두 (df, cfg) -> BoolSeries 형태로 통일해 추가/제거가 쉽도록 한다.
  - 컬럼 역할 지정은 JSON으로 저장/불러오기 (같은 조사를 매일 돌리는 용도).

set_page_config 와 비밀번호 확인은 Home.py 에서 이미 처리합니다.
"""

from __future__ import annotations

import io
import json
import re
import tempfile
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# utils.py 가 있으면 쓰고, 없으면 내부 구현으로 폴백
# ---------------------------------------------------------------------------
try:  # pragma: no cover
    import utils as _utils
except Exception:  # pragma: no cover
    _utils = None


def _fallback_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, frame in sheets.items():
            frame.to_excel(xw, sheet_name=name[:31], index=False)
    return buf.getvalue()


to_excel_bytes = getattr(_utils, "to_excel_bytes", None) or _fallback_to_excel_bytes

# ---------------------------------------------------------------------------
# 상수
# ---------------------------------------------------------------------------
VERSION = "v1.0"
JUDGE_COL = "판정"
JUDGE_OPTIONS = ["미검토", "확인 완료", "삭제 대상", "보류"]

# 관리 컬럼 후보 (SAV 변환 페이지와 동일한 감각으로 기본 제외 대상 추정)
ADMIN_HINTS = [
    "quota_id", "outvar", "areaM", "status", "ip", "uid", "panel",
    "start", "end", "date", "time", "modify", "sample",
]

GRADE_CONFIRMED = "확정"
GRADE_HIGH = "높음"
GRADE_MID = "중간"
GRADE_LOW = "낮음"
GRADE_NONE = "-"


# ---------------------------------------------------------------------------
# 파일 읽기
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def read_sav(file_bytes: bytes, filename: str):
    """SAV 읽기. Windows 에서 NamedTemporaryFile 재개방 문제가 있어 TemporaryDirectory 사용."""
    import pyreadstat

    with tempfile.TemporaryDirectory() as tmpdir:
        path = Path(tmpdir) / (Path(filename).name or "input.sav")
        path.write_bytes(file_bytes)
        df, meta = pyreadstat.read_sav(str(path), apply_value_formats=False)

    value_labels = dict(getattr(meta, "variable_value_labels", {}) or {})
    col_labels = dict(zip(meta.column_names, meta.column_labels or []))
    return df, value_labels, col_labels


@st.cache_data(show_spinner=False)
def read_excel(file_bytes: bytes, filename: str, sheet: str | int = 0):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    return df, {}, {}


@st.cache_data(show_spinner=False)
def excel_sheet_names(file_bytes: bytes) -> list[str]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    return list(xl.sheet_names)


# ---------------------------------------------------------------------------
# 매트릭스 문항 그룹 자동 감지
# ---------------------------------------------------------------------------
# 접두사(base) 자체에 숫자가 들어갈 수 있으므로(Q5_1) base 는 느슨하게 두고
# 구분자를 별도 그룹으로 잡는다. 구분자까지 키에 포함시켜 Q5_1 과 Q51 이 섞이지 않게 한다.
GROUP_PAT = re.compile(r"^(?P<base>.+?)(?P<sep>[_\-.\s]?)(?P<num>\d{1,3})$")


def detect_matrix_groups(columns, exclude: set[str], min_items: int = 3) -> dict[str, list[str]]:
    """Q5_1 ~ Q5_10 처럼 접두사 + 숫자 패턴으로 묶인 컬럼 그룹을 추정."""
    buckets: dict[str, list[tuple[int, str]]] = defaultdict(list)
    for col in columns:
        if col in exclude:
            continue
        m = GROUP_PAT.match(str(col).strip())
        if not m:
            continue
        buckets[m.group("base") + m.group("sep")].append((int(m.group("num")), col))

    groups: dict[str, list[str]] = {}
    for base, items in buckets.items():
        if len(items) < min_items:
            continue
        items.sort()
        groups[base] = [c for _, c in items]
    return dict(sorted(groups.items()))


def numeric_frame(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    return df[cols].apply(pd.to_numeric, errors="coerce")


# ---------------------------------------------------------------------------
# 검사 함수  ((df, cfg) -> BoolSeries)
# ---------------------------------------------------------------------------
def chk_intval_ratio(df, cfg) -> pd.Series:
    """소요시간이 중위값 대비 일정 비율 미만. 하위 이상치 판정의 주력."""
    col = cfg.get("intval_col")
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)
    vals = pd.to_numeric(df[col], errors="coerce")
    med = vals.median()
    if not np.isfinite(med) or med <= 0:
        return pd.Series(False, index=df.index)
    limit = med * (cfg["ratio_pct"] / 100.0)
    return vals.notna() & (vals < limit)


def chk_intval_abs_low(df, cfg) -> pd.Series:
    """절대 초 하한. 물리적으로 불가능한 시간."""
    col = cfg.get("intval_col")
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)
    vals = pd.to_numeric(df[col], errors="coerce")
    return vals.notna() & (vals < cfg["abs_low"])


def chk_intval_high(df, cfg) -> pd.Series:
    """상한 초과. 단독으로는 삭제 사유가 아니지만 다른 검사와 겹치면 신호."""
    col = cfg.get("intval_col")
    if not col or col not in df.columns or not cfg.get("use_high"):
        return pd.Series(False, index=df.index)
    vals = pd.to_numeric(df[col], errors="coerce")
    return vals.notna() & (vals > cfg["abs_high"])


def chk_straightline(df, cfg):
    """매트릭스 그룹 내 모든 값이 동일. 역문항은 그룹에서 제외해 오탐을 줄인다."""
    mask = pd.Series(False, index=df.index)
    detail: dict[int, list[str]] = defaultdict(list)
    reverse = set(cfg.get("reverse_cols", []))
    min_items = cfg.get("sl_min_items", 3)

    for base, cols in cfg.get("groups", {}).items():
        use = [c for c in cols if c not in reverse and c in df.columns]
        if len(use) < min_items:
            continue
        sub = numeric_frame(df, use)
        hit = sub.notna().all(axis=1) & (sub.nunique(axis=1) == 1)
        for idx in df.index[hit]:
            detail[idx].append(base)
        mask |= hit
    return mask, {k: ", ".join(v) for k, v in detail.items()}


def chk_zigzag(df, cfg):
    """1-5-1-5 식 기계적 반복.

    부호 반전만으로 판정하면 정상 응답의 자연스러운 진동까지 대량으로 걸린다
    (6문항 그룹이면 우연히 반전할 확률이 6% 내외). 따라서
      ① 인접 차이의 부호가 계속 반전
      ② 등장한 값의 종류가 적음 (기본 3개 이하)
      ③ 진폭이 일정 이상 (기본 2 이상)
    을 모두 만족할 때만 적발한다.
    """
    mask = pd.Series(False, index=df.index)
    detail: dict[int, list[str]] = defaultdict(list)
    reverse = set(cfg.get("reverse_cols", []))
    max_uniq = int(cfg.get("zz_max_unique", 3))
    min_amp = float(cfg.get("zz_min_amp", 2))

    for base, cols in cfg.get("groups", {}).items():
        use = [c for c in cols if c not in reverse and c in df.columns]
        if len(use) < 4:
            continue
        sub = numeric_frame(df, use)
        arr = sub.to_numpy(dtype=float)
        valid = ~np.isnan(arr).any(axis=1)
        d = np.diff(arr, axis=1)
        with np.errstate(invalid="ignore"):
            nonzero = (d != 0).all(axis=1)
            signs = np.sign(d)
            alt = (signs[:, 1:] * signs[:, :-1] < 0).all(axis=1)
            amp = np.nanmax(np.abs(d), axis=1) >= min_amp
        few = (sub.nunique(axis=1) <= max_uniq).to_numpy()
        hit = pd.Series(valid & nonzero & alt & amp & few, index=df.index)
        for idx in df.index[hit]:
            detail[idx].append(base)
        mask |= hit
    return mask, {k: ", ".join(v) for k, v in detail.items()}


def chk_required_missing(df, cfg):
    cols = [c for c in cfg.get("required_cols", []) if c in df.columns]
    if not cols:
        return pd.Series(False, index=df.index), {}
    na = df[cols].isna()
    mask = na.any(axis=1)
    detail = {}
    for idx in df.index[mask]:
        miss = [c for c in cols if bool(na.at[idx, c])]
        detail[idx] = ", ".join(miss[:6]) + ("..." if len(miss) > 6 else "")
    return mask, detail


def chk_dup_id(df, cfg) -> pd.Series:
    col = cfg.get("id_col")
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)
    return df[col].duplicated(keep=False) & df[col].notna()


def chk_dup_vector(df, cfg) -> pd.Series:
    """관리 컬럼을 뺀 실질 응답 벡터가 완전히 동일한 케이스 (복붙 의심)."""
    cols = cfg.get("substantive_cols", [])
    cols = [c for c in cols if c in df.columns]
    if len(cols) < 5:
        return pd.Series(False, index=df.index)
    return df[cols].duplicated(keep=False)


def chk_range_violation(df, cfg):
    """value label 에 정의되지 않은 코드값. pyreadstat 메타를 그대로 활용."""
    labels = cfg.get("value_labels", {}) or {}
    targets = [c for c in labels if c in df.columns and c in cfg.get("substantive_cols", [])]
    if not targets:
        return pd.Series(False, index=df.index), {}

    mask = pd.Series(False, index=df.index)
    detail: dict[int, list[str]] = defaultdict(list)
    for col in targets:
        allowed = set(labels[col].keys())
        vals = df[col]
        hit = vals.notna() & ~vals.isin(allowed)
        if hit.any():
            for idx in df.index[hit]:
                detail[idx].append(f"{col}={vals.at[idx]}")
            mask |= hit
    return mask, {k: ", ".join(v[:5]) for k, v in detail.items()}


# 검사 레지스트리: (key, 표시명, 기본 가중치, 함수, 확정등급 여부)
CHECKS = [
    ("intval_ratio", "소요시간 중위값 대비 미달", 2.0, chk_intval_ratio, False),
    ("intval_abs", "소요시간 절대 하한 미달", 3.0, chk_intval_abs_low, False),
    ("intval_high", "소요시간 상한 초과", 0.5, chk_intval_high, False),
    ("straightline", "직진성", 2.0, chk_straightline, False),
    ("zigzag", "지그재그 패턴", 1.5, chk_zigzag, False),
    ("required", "필수 문항 결측", 1.5, chk_required_missing, False),
    ("dup_id", "ID 중복", 0.0, chk_dup_id, True),
    ("dup_vector", "응답벡터 완전 동일", 0.0, chk_dup_vector, True),
    ("range", "미정의 코드값", 1.0, chk_range_violation, False),
]


# ---------------------------------------------------------------------------
# 화면
# ---------------------------------------------------------------------------
st.title("🔍 데이터 검증")
st.caption(f"불성실 응답 / 품질 검증 · {VERSION} · 자동 삭제하지 않고 검토 대상만 산출합니다.")

up = st.file_uploader("원자료 업로드 (.sav / .xlsx / .xls)", type=["sav", "xlsx", "xls"])
if not up:
    st.info("파일을 올리면 컬럼 역할 지정 화면이 나타납니다.")
    st.stop()

raw_bytes = up.getvalue()
suffix = Path(up.name).suffix.lower()

if suffix == ".sav":
    df, value_labels, col_labels = read_sav(raw_bytes, up.name)
else:
    sheets = excel_sheet_names(raw_bytes)
    sheet = st.selectbox("시트 선택", sheets, index=0) if len(sheets) > 1 else sheets[0]
    df, value_labels, col_labels = read_excel(raw_bytes, up.name, sheet)

df = df.reset_index(drop=True)
all_cols = list(df.columns)
st.success(f"{len(df):,} 케이스 × {len(all_cols):,} 컬럼")

# --- 설정 불러오기 -----------------------------------------------------------
with st.expander("① 컬럼 역할 지정", expanded=True):
    cfg_file = st.file_uploader(
        "저장한 설정 JSON 불러오기 (선택)", type=["json"], key="cfgup"
    )
    saved: dict = {}
    if cfg_file is not None:
        try:
            saved = json.loads(cfg_file.getvalue().decode("utf-8"))
            st.caption("설정을 불러왔습니다. 아래 값에 반영되어 있습니다.")
        except Exception as e:
            st.warning(f"설정 파일을 읽지 못했습니다: {e}")

    def _pick(key, default):
        v = saved.get(key, default)
        if isinstance(default, list):
            return [x for x in (v or []) if x in all_cols]
        return v if v in all_cols else default

    lower = {c.lower(): c for c in all_cols}
    id_guess = _pick("id_col", lower.get("id") or lower.get("no") or all_cols[0])
    intval_guess = _pick("intval_col", lower.get("intval"))

    c1, c2 = st.columns(2)
    id_col = c1.selectbox(
        "ID 컬럼", all_cols, index=all_cols.index(id_guess) if id_guess in all_cols else 0
    )
    intval_opts = ["(없음)"] + all_cols
    intval_col = c2.selectbox(
        "소요시간 컬럼 (초)",
        intval_opts,
        index=intval_opts.index(intval_guess) if intval_guess in all_cols else 0,
        help="쿼터 솔루션 페이지에서 쓰는 intval 과 같은 컬럼입니다.",
    )
    intval_col = None if intval_col == "(없음)" else intval_col

    admin_guess = saved.get("admin_cols")
    if admin_guess is None:
        admin_guess = [
            c for c in all_cols
            if any(h.lower() in str(c).lower() for h in ADMIN_HINTS)
        ]
    admin_cols = st.multiselect(
        "제외할 관리 컬럼",
        all_cols,
        default=[c for c in admin_guess if c in all_cols],
        help="응답벡터 중복·코드값 검사에서 제외됩니다. ID·소요시간은 자동 제외됩니다.",
    )

    reserved = set(admin_cols) | {id_col} | ({intval_col} if intval_col else set())
    substantive_cols = [c for c in all_cols if c not in reserved]

    detected = detect_matrix_groups(all_cols, reserved)
    saved_groups = saved.get("groups") or {}
    group_names = st.multiselect(
        "매트릭스 문항 그룹",
        list(detected.keys()),
        default=[g for g in (saved_groups.keys() or detected.keys()) if g in detected],
        help="접두사 + 숫자 패턴으로 자동 감지한 후보입니다. 직진성·지그재그 판정 단위가 됩니다.",
    )
    groups = {g: detected[g] for g in group_names}
    if groups:
        st.caption(
            " · ".join(f"{g} ({len(cols)}문항)" for g, cols in list(groups.items())[:12])
        )

    group_cols_flat = [c for cols in groups.values() for c in cols]
    reverse_cols = st.multiselect(
        "역문항 (직진성 판정에서 제외)",
        group_cols_flat,
        default=_pick("reverse_cols", []),
        help="역문항이 섞인 그룹은 직진성이 오히려 정상 응답일 수 있어 오탐이 납니다.",
    )

    required_cols = st.multiselect(
        "필수 응답 문항",
        substantive_cols,
        default=_pick("required_cols", []),
    )

# --- 임계값 -----------------------------------------------------------------
with st.expander("② 검사 항목 · 임계값", expanded=True):
    enabled: dict[str, bool] = {}
    weights: dict[str, float] = {}
    saved_en = saved.get("enabled", {})
    saved_w = saved.get("weights", {})

    st.markdown("**소요시간**")
    if intval_col:
        vals = pd.to_numeric(df[intval_col], errors="coerce").dropna()
        med = float(vals.median()) if len(vals) else 0.0
        st.caption(
            f"중위값 {med:,.0f}초 ({med/60:,.1f}분) · "
            f"최소 {vals.min():,.0f} / 최대 {vals.max():,.0f}"
        )
        t1, t2, t3 = st.columns(3)
        ratio_pct = t1.slider("중위값 대비 하한 (%)", 5, 90, int(saved.get("ratio_pct", 40)), 5)
        abs_low = t2.number_input(
            "절대 초 하한", min_value=0, value=int(saved.get("abs_low", 120)), step=10
        )
        use_high = t3.checkbox("상한도 검사", value=bool(saved.get("use_high", False)))
        abs_high = t3.number_input(
            "절대 초 상한",
            min_value=0,
            value=int(saved.get("abs_high", max(3600, int(med * 5) if med else 3600))),
            step=300,
            disabled=not use_high,
        )

        # 임계선을 얹은 분포. 슬라이더를 움직이며 걸리는 건수를 보고 정한다.
        try:
            import altair as alt

            base = alt.Chart(pd.DataFrame({intval_col: vals.to_numpy()}))
            hist = base.mark_bar(opacity=0.75).encode(
                alt.X(f"{intval_col}:Q", bin=alt.Bin(maxbins=60), title="소요시간(초)"),
                alt.Y("count()", title="케이스"),
            )
            lines = [("중위값 대비", med * ratio_pct / 100.0), ("절대 하한", float(abs_low))]
            if use_high:
                lines.append(("상한", float(abs_high)))
            rule = (
                alt.Chart(pd.DataFrame({"x": [v for _, v in lines], "구분": [k for k, _ in lines]}))
                .mark_rule(strokeWidth=2, strokeDash=[4, 3])
                .encode(x="x:Q", color=alt.Color("구분:N", legend=alt.Legend(title=None)))
            )
            st.altair_chart(hist + rule, use_container_width=True)
        except Exception:
            st.bar_chart(np.histogram(vals, bins=40)[0])

        n_ratio = int((vals < med * ratio_pct / 100.0).sum())
        n_abs = int((vals < abs_low).sum())
        n_high = int((vals > abs_high).sum()) if use_high else 0
        st.caption(
            f"현재 임계값으로 → 중위값 미달 {n_ratio:,}건 · 절대 하한 미달 {n_abs:,}건"
            + (f" · 상한 초과 {n_high:,}건" if use_high else "")
        )
    else:
        ratio_pct, abs_low, abs_high, use_high = 40, 0, 0, False
        st.caption("소요시간 컬럼이 지정되지 않아 시간 검사는 비활성화됩니다.")

    st.divider()
    st.markdown("**패턴 검사 파라미터**")
    p1, p2, p3 = st.columns(3)
    sl_min_items = p1.number_input(
        "직진성 최소 문항 수",
        min_value=3,
        value=int(saved.get("sl_min_items", 3)),
        step=1,
        help="그룹 내 유효 문항이 이보다 적으면 직진성 판정을 건너뜁니다.",
    )
    zz_max_unique = p2.number_input(
        "지그재그 최대 값 종류",
        min_value=2,
        max_value=5,
        value=int(saved.get("zz_max_unique", 2)),
        step=1,
        help="2 면 1-5-1-5 처럼 두 값만 오가는 경우만 잡습니다. 늘리면 느슨해집니다.",
    )
    zz_min_amp = p3.number_input(
        "지그재그 최소 진폭",
        min_value=1.0,
        value=float(saved.get("zz_min_amp", 2.0)),
        step=1.0,
        help="인접 문항 간 최대 변화폭. 부호 반전만으로 판정하면 정상 응답도 대량 적발됩니다.",
    )

    st.divider()
    st.markdown("**검사 사용 여부 · 가중치**")
    st.caption("가중치 합계로 등급을 매깁니다. ID 중복·응답벡터 동일은 단독으로 확정 등급입니다.")

    for key, label, w_default, _fn, is_confirm in CHECKS:
        r1, r2 = st.columns([3, 1])
        time_check = key.startswith("intval")
        disabled = time_check and not intval_col
        enabled[key] = r1.checkbox(
            label + ("  ← 확정" if is_confirm else ""),
            value=bool(saved_en.get(key, not disabled)) and not disabled,
            disabled=disabled,
            key=f"en_{key}",
        )
        weights[key] = (
            0.0
            if is_confirm
            else r2.number_input(
                "가중치",
                min_value=0.0,
                max_value=10.0,
                value=float(saved_w.get(key, w_default)),
                step=0.5,
                key=f"w_{key}",
                label_visibility="collapsed",
                disabled=is_confirm or disabled,
            )
        )

    g1, g2 = st.columns(2)
    high_cut = g1.number_input(
        "‘높음’ 점수 기준", min_value=0.5, value=float(saved.get("high_cut", 3.5)), step=0.5
    )
    mid_cut = g2.number_input(
        "‘중간’ 점수 기준", min_value=0.5, value=float(saved.get("mid_cut", 2.0)), step=0.5
    )

cfg = {
    "id_col": id_col,
    "intval_col": intval_col,
    "admin_cols": admin_cols,
    "substantive_cols": substantive_cols,
    "groups": groups,
    "reverse_cols": reverse_cols,
    "required_cols": required_cols,
    "value_labels": value_labels,
    "ratio_pct": ratio_pct,
    "abs_low": abs_low,
    "abs_high": abs_high,
    "use_high": use_high,
    "sl_min_items": sl_min_items,
    "zz_max_unique": zz_max_unique,
    "zz_min_amp": zz_min_amp,
    "enabled": enabled,
    "weights": weights,
    "high_cut": high_cut,
    "mid_cut": mid_cut,
}

# 설정 저장 (value_labels 는 제외)
save_payload = {k: v for k, v in cfg.items() if k != "value_labels"}
st.download_button(
    "설정 JSON 저장",
    data=json.dumps(save_payload, ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="검증설정.json",
    mime="application/json",
)

st.divider()

# --- 검사 실행 --------------------------------------------------------------
masks: dict[str, pd.Series] = {}
details: dict[str, dict] = {}

for key, label, _w, fn, _c in CHECKS:
    if not enabled.get(key):
        masks[key] = pd.Series(False, index=df.index)
        details[key] = {}
        continue
    out = fn(df, cfg)
    if isinstance(out, tuple):
        m, d = out
    else:
        m, d = out, {}
    masks[key] = m.fillna(False).astype(bool)
    details[key] = d

label_of = {k: lab for k, lab, _w, _f, _c in CHECKS}
confirm_keys = [k for k, _lab, _w, _f, c in CHECKS if c]

score = pd.Series(0.0, index=df.index)
for key, _lab, _w, _f, is_confirm in CHECKS:
    if is_confirm:
        continue
    score += masks[key].astype(float) * float(weights.get(key, 0.0))

confirmed = pd.Series(False, index=df.index)
for k in confirm_keys:
    if enabled.get(k):
        confirmed |= masks[k]


def grade_of(idx) -> str:
    if confirmed.at[idx]:
        return GRADE_CONFIRMED
    s = score.at[idx]
    if s >= high_cut:
        return GRADE_HIGH
    if s >= mid_cut:
        return GRADE_MID
    if s > 0:
        return GRADE_LOW
    return GRADE_NONE


reason_rows = []
for idx in df.index:
    hits = [k for k in masks if masks[k].at[idx]]
    if not hits:
        continue
    parts = []
    for k in hits:
        d = details.get(k, {}).get(idx)
        parts.append(f"{label_of[k]}({d})" if d else label_of[k])
    reason_rows.append(
        {
            "행": int(idx) + 2,
            "ID": df.at[idx, id_col] if id_col in df.columns else idx,
            "등급": grade_of(idx),
            "점수": round(float(score.at[idx]), 2),
            "걸린 검사 수": len(hits),
            "사유": " / ".join(parts),
            JUDGE_COL: JUDGE_OPTIONS[0],
            "_idx": int(idx),
        }
    )

flagged = pd.DataFrame(reason_rows)

# --- 요약 -------------------------------------------------------------------
st.subheader("검사 결과")

summary = pd.DataFrame(
    [
        {
            "검사": label_of[k],
            "사용": "○" if enabled.get(k) else "-",
            "적발": int(masks[k].sum()),
            "비율(%)": round(masks[k].sum() / max(len(df), 1) * 100, 2),
            "가중치": weights.get(k, 0.0) if k not in confirm_keys else "확정",
        }
        for k in masks
    ]
)

m1, m2, m3, m4 = st.columns(4)
m1.metric("전체 케이스", f"{len(df):,}")
if flagged.empty:
    m2.metric("확정", "0")
    m3.metric("높음", "0")
    m4.metric("검토 대상", "0")
    st.success("현재 임계값에서 적발된 케이스가 없습니다.")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    st.stop()

grades = flagged["등급"]
m2.metric("확정", f"{int((grades == GRADE_CONFIRMED).sum()):,}")
m3.metric("높음", f"{int((grades == GRADE_HIGH).sum()):,}")
m4.metric("검토 대상 전체", f"{len(flagged):,}")

st.dataframe(summary, use_container_width=True, hide_index=True)

# --- 상세 + 판정 ------------------------------------------------------------
st.markdown("**상세 (등급 · 점수 내림차순)**")

order = pd.Categorical(
    flagged["등급"],
    categories=[GRADE_CONFIRMED, GRADE_HIGH, GRADE_MID, GRADE_LOW, GRADE_NONE],
    ordered=True,
)
flagged = flagged.assign(_ord=order).sort_values(
    ["_ord", "점수"], ascending=[True, False]
).drop(columns="_ord").reset_index(drop=True)

grade_filter = st.multiselect(
    "등급 필터",
    [GRADE_CONFIRMED, GRADE_HIGH, GRADE_MID, GRADE_LOW],
    default=[GRADE_CONFIRMED, GRADE_HIGH, GRADE_MID, GRADE_LOW],
)
view = flagged[flagged["등급"].isin(grade_filter)].copy()

editor_key = "verify_editor"
edited = st.data_editor(
    view.drop(columns=["_idx"]),
    key=editor_key,
    use_container_width=True,
    hide_index=True,
    disabled=[c for c in view.columns if c not in (JUDGE_COL, "_idx")],
    column_config={
        JUDGE_COL: st.column_config.SelectboxColumn(
            JUDGE_COL, options=JUDGE_OPTIONS, required=True, width="small"
        ),
        "사유": st.column_config.TextColumn("사유", width="large"),
    },
)
edited["_idx"] = view["_idx"].to_numpy()

# --- 내보내기 ---------------------------------------------------------------
st.divider()
st.markdown("**내보내기**")

del_idx = edited.loc[edited[JUDGE_COL] == "삭제 대상", "_idx"].tolist()

sheets = {
    "요약": summary,
    "사유": edited.drop(columns=["_idx"]),
    "의심케이스_Raw": df.loc[edited["_idx"].tolist()].reset_index(drop=True),
}

e1, e2 = st.columns(2)
e1.download_button(
    "검토용 엑셀 다운로드",
    data=to_excel_bytes(sheets),
    file_name=f"데이터검증_{Path(up.name).stem}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

if del_idx:
    clean = df.drop(index=del_idx).reset_index(drop=True)
    e2.download_button(
        f"정제 데이터 다운로드 ({len(del_idx):,}건 제외)",
        data=to_excel_bytes({"Raw": clean}),
        file_name=f"정제_{Path(up.name).stem}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    e2.caption("판정에서 ‘삭제 대상’을 지정하면 정제 데이터 다운로드가 활성화됩니다.")

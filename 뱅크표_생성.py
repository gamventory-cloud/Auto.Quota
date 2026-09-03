# 뱅크표_생성.py
#
# SPSS .sav 파일로 뱅크표(배너표)를 만드는 페이지입니다.
# 계산 로직은 banner_table_engine.py 에 있습니다.
#
# set_page_config 와 비밀번호 확인은 Home.py 가 처리하므로 여기에 두지 않습니다.
#
# ── session_state 키 ─────────────────────────────────────────────────
#   멀티페이지 앱은 session_state 를 모든 페이지가 함께 씁니다.
#   다른 페이지와 겹치지 않도록 이 페이지의 키는 모두 'bt_' 로 시작합니다.

import tempfile
from pathlib import Path

import pandas as pd
import pyreadstat
import streamlit as st

from banner_table_engine import (
    BANNER_COL,
    BANNER_ROW,
    BannerSpec,
    SigSpec,
    blocks_to_json,
    build_battery_block,
    build_block,
    compare_waves,
    compute_frequencies,
    compute_table,
    freq_to_frame,
    load_settings,
    missing_vars,
    parse_sps,
    parse_summary_spec,
    read_sps_text,
    result_to_frame,
    safe_stem,
    title_with_marker,
    write_freq_xlsx,
    write_tables_xlsx,
)
from banner_table_form import (
    blocks_to_form,
    read_form,
    write_filled_form,
    write_form_template,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── st.dataframe 폭 지정 ─────────────────────────────────────────────
# use_container_width 는 지원 종료 예고가 붙어 경고가 뜨고, width="stretch" 는
# 예전 버전에서 통하지 않습니다. 버전을 보고 한 번만 정해 둡니다.
try:
    _ver = tuple(int(x) for x in st.__version__.split(".")[:2])
    _WIDE = {"width": "stretch"} if _ver >= (1, 49) else {"use_container_width": True}
except Exception:                                        # noqa: BLE001
    _WIDE = {"use_container_width": True}


@st.cache_data(show_spinner=False)
def load_sav(file_bytes: bytes):
    """업로드한 .sav 를 읽는다. 같은 파일을 다시 읽지 않도록 캐시한다."""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".sav", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        return pyreadstat.read_sav(tmp_path)
    finally:
        # pyreadstat 이 메모리로 다 읽으므로 임시파일은 바로 지운다.
        if tmp_path:
            Path(tmp_path).unlink(missing_ok=True)


@st.cache_data(show_spinner=False)
def load_sps(file_bytes: bytes) -> str:
    return read_sps_text(file_bytes)


st.header("뱅크표 생성")
st.caption(
    "SPSS .sav 를 올리고 배너와 행 변수를 골라 뱅크표를 만듭니다. "
    "Embrain 'Table' 매크로 신텍스(.sps)가 있으면 그 표들을 한 번에 뽑을 수도 있습니다."
)

# 데이터 업로드는 화면 맨 위에 둔다. (탭마다 필요한 폼·신텍스는 해당 탭에서 받는다)
sav_file = st.file_uploader("SPSS 데이터 (.sav)", type=["sav"], key="bt_sav")

if sav_file is None:
    st.info("먼저 .sav 파일을 올려 주세요. 표를 만드는 방법은 세 가지입니다 — "
            "변수를 직접 고르거나, 엑셀 폼에 적어 올리거나, 기존 Table 신텍스를 올리면 됩니다.")
    st.stop()

try:
    df, meta = load_sav(sav_file.getvalue())
except Exception as e:                                   # noqa: BLE001
    st.error(f".sav 파일을 읽지 못했습니다 — {e}")
    st.stop()

value_labels = meta.variable_value_labels
col_labels = meta.column_names_to_labels

st.success(f"{sav_file.name} · {len(df):,}행 × {len(df.columns)}열")

# 내려받는 파일 이름은 올린 데이터 이름을 따른다.
# 데이터마다 변수 구성이 다르므로 이름이 같으면 서로 섞이기 쉽다.
SAV_STEM = safe_stem(sav_file.name)


def label_for(col: str) -> str:
    """변수명 뒤에 변수 라벨을 붙여 고르기 쉽게 한다."""
    lbl = col_labels.get(col)
    return f"{col} ({lbl})" if lbl and lbl != col else col


DISPLAY_MAP = {label_for(c): c for c in df.columns}
DISPLAY_NAMES = list(DISPLAY_MAP.keys())


def to_vars(display_names) -> list[str]:
    return [DISPLAY_MAP[d] for d in display_names]


st.session_state.setdefault("bt_merge_banners", [])
st.session_state.setdefault("bt_results", [])
st.session_state.setdefault("bt_blocks", [])      # 담아둔 표의 '정의' (설정 저장용)

tab_manual, tab_form, tab_quick, tab_syntax = st.tabs(
    ["변수 골라서 만들기", "엑셀 폼으로 만들기", "빈도 · 교차표", "신텍스로 한 번에"]
)


# =============================================================================
# 변수 골라서 만들기 (.sav 만 있으면 됨)
# =============================================================================
with tab_manual:
    st.subheader("1. 배너 (열)")
    banner_disp = st.multiselect(
        "배너로 쓸 변수 — 변수 하나가 배너 그룹 하나가 되고, 값 라벨이 하위 컬럼이 됩니다",
        DISPLAY_NAMES,
        key="bt_banner_single",
    )

    with st.expander("여러 변수를 하나의 다중응답 배너로 묶기 (예: 시설유형 4개 변수 → 배너 1개)"):
        c1, c2 = st.columns([1, 2])
        merge_label = c1.text_input("배너 그룹 이름", key="bt_merge_label")
        merge_vars_disp = c2.multiselect(
            "묶을 변수들 — 변수마다 자기 코드값만 갖고 나머지는 결측인 방식을 가정합니다",
            DISPLAY_NAMES,
            key="bt_merge_vars",
        )
        if st.button("배너 그룹 추가", key="bt_add_merge"):
            if merge_label and len(merge_vars_disp) >= 2:
                st.session_state["bt_merge_banners"].append(
                    {"label": merge_label, "varlist": to_vars(merge_vars_disp)}
                )
            else:
                st.warning("그룹 이름과 변수 2개 이상이 필요합니다.")

        for i, g in enumerate(st.session_state["bt_merge_banners"]):
            gc1, gc2 = st.columns([6, 1])
            gc1.write(f"**{g['label']}** — {', '.join(g['varlist'])}")
            if gc2.button("삭제", key=f"bt_del_merge_{i}"):
                st.session_state["bt_merge_banners"].pop(i)

    banners = [BannerSpec(kind="single", var=v) for v in to_vars(banner_disp)] + [
        BannerSpec(kind="merge", label=g["label"], varlist=g["varlist"])
        for g in st.session_state["bt_merge_banners"]
    ]
    if not banners:
        st.info("배너로 쓸 변수를 최소 1개 골라 주세요. '전체' 컬럼은 항상 자동으로 들어갑니다.")

    st.subheader("2. 행 변수")
    row_type_disp = st.radio(
        "행 변수 유형",
        ["단일 응답 (단수)", "다중 응답 (복수)", "수치형 (평균 · 중위값)",
         "척도 종합표 (문항 여러 개를 한 표에)"],
        horizontal=True,
        key="bt_row_type",
    )

    row_vars: list[str] = []
    row_ma_mode = "category"
    obser_stats: list[str] = []
    obser_show_values = False
    summaries: list = []
    battery_metric: str | None = None

    def summary_ui(var: str | None, key_suffix: str = "") -> list:
        """척도 요약 (Top2 · Middle · Bottom2 · 평균) 을 고르는 칸.

        단수와 수치형이 같은 칸을 쓴다. 다른 점은 '보기' 를 어디서 가져오는지
        뿐이다 — 단수는 값 라벨, 수치형은 실제 응답된 값.
        """
        picked_sum: list = []
        with st.expander(
            "척도 요약 — Top2 · Middle · Bottom2 · 평균을 '계' 뒤에 붙이기"
        ):
            if not var:
                st.caption("변수를 먼저 골라 주세요.")
                return picked_sum

            vl_here = value_labels.get(var, {})
            if vl_here:
                codes = sorted(vl_here.keys())
                labels_txt = ", ".join(
                    f"{int(c) if float(c).is_integer() else c}={vl_here[c]}"
                    for c in codes
                )
                st.caption(f"보기 {len(codes)}개 — {labels_txt}")
            else:
                codes = sorted(df[var].dropna().unique().tolist())
                if not codes:
                    st.caption("응답된 값이 없어 요약을 만들 수 없습니다.")
                    return picked_sum
                st.caption(
                    f"값 라벨이 없는 변수입니다. 응답된 값 {len(codes)}종을 "
                    "보기로 봅니다 — 평균·표준편차는 값 자체로 계산됩니다."
                )

            sc1, sc2, sc3, sc4, sc5 = st.columns(5)
            top_n = sc1.number_input("상위 몇 개", 0, len(codes), 0,
                                     key=f"bt_sum_top{key_suffix}")
            bot_n = sc2.number_input("하위 몇 개", 0, len(codes), 0,
                                     key=f"bt_sum_bot{key_suffix}")
            use_mid = sc3.checkbox("중간(나머지)", key=f"bt_sum_mid{key_suffix}")
            use_mean = sc4.checkbox("평균", key=f"bt_sum_mean{key_suffix}")
            use_std = sc5.checkbox("표준편차", key=f"bt_sum_std{key_suffix}")

            parts = []
            if top_n:
                parts.append(f"상{int(top_n)}")
            if use_mid:
                parts.append("중")
            if bot_n:
                parts.append(f"하{int(bot_n)}")
            if use_mean:
                parts.append("평균")
            if use_std:
                parts.append("표준편차")
            if not parts:
                return picked_sum

            picked_sum, sum_problems = parse_summary_spec(
                ",".join(parts), codes, decimals=1
            )
            for msg in sum_problems:
                st.warning(msg)
            if picked_sum:
                def _code_txt(c):
                    if vl_here:
                        return str(vl_here[c])
                    return f"{int(c) if float(c).is_integer() else c}"

                st.caption(
                    "붙는 칸: "
                    + " · ".join(
                        x.label if x.kind != "group" else
                        f"{x.label}(" + ",".join(_code_txt(c) for c in x.codes) + ")"
                        for x in picked_sum
                    )
                )
        return picked_sum

    if row_type_disp.startswith("단일"):
        row_type = "single"
        picked = st.selectbox("행 변수", DISPLAY_NAMES, key="bt_row_single")
        row_vars = [DISPLAY_MAP[picked]] if picked else []
        summaries = summary_ui(row_vars[0] if row_vars else None)

    elif row_type_disp.startswith("다중"):
        row_type = "multi"
        picked_multi = st.multiselect(
            "행에 쓸 다중응답 변수들 (예: 봉안시설_1 ~ 봉안시설_4)",
            DISPLAY_NAMES,
            key="bt_row_multi",
        )
        row_vars = to_vars(picked_multi)
        st.caption(
            "변수마다 자기 코드값을 갖고 해당 없으면 결측인 방식으로 읽습니다 "
            "(SPSS 다중응답 세트의 일반적인 형태)."
        )

    elif row_type_disp.startswith("척도 종합"):
        row_type = "battery"
        picked_bat = st.multiselect(
            "한 표에 넣을 문항들 — 척도가 같은 문항끼리 (예: Q5_1 ~ Q5_5)",
            DISPLAY_NAMES,
            key="bt_row_battery",
        )
        row_vars = to_vars(picked_bat)
        shape = st.radio(
            "표 모양",
            ["보기 분포형 — 열이 보기 + 계 + 요약",
             "평균 서머리(격자형) — 열이 배너, 값은 지표 하나"],
            key="bt_bat_shape",
        )
        summaries = summary_ui(row_vars[0] if row_vars else None, "_bat")
        if shape.startswith("평균 서머리"):
            choices = ["mean", "std"] + [s.label for s in summaries
                                         if s.kind == "group"]
            battery_metric = st.selectbox(
                "격자에 넣을 지표",
                choices,
                format_func=lambda s: {"mean": "평균", "std": "표준편차"}.get(s, s),
                key="bt_bat_metric",
                help="Top2 같은 묶음을 쓰려면 위 '척도 요약' 에서 먼저 정의하세요.",
            )
            st.caption(
                "격자형은 열(배너)끼리 비교하므로 아래에서 유의성 검정을 켤 수 있습니다."
            )
        else:
            st.caption(
                "보기 분포형은 행끼리(문항끼리) 비교하는 표입니다. 같은 응답자가 모든 "
                "문항에 답했으므로 유의성 검정은 하지 않습니다."
            )
        if row_vars:
            lab_sets = {tuple(sorted(value_labels.get(v, {}).keys())) for v in row_vars}
            if len(lab_sets) > 1:
                st.warning(
                    "고른 문항들의 보기가 서로 다릅니다. 합집합으로 계산하지만, "
                    "척도가 같은 문항끼리 묶는 것이 좋습니다."
                )

    else:
        row_type = "obser"
        picked = st.selectbox("수치형 변수 (이용료·나이처럼 값이 숫자인 문항)",
                              DISPLAY_NAMES, key="bt_row_obser")
        row_vars = [DISPLAY_MAP[picked]] if picked else []
        obser_stats = st.multiselect(
            "표시할 통계",
            ["MEAN", "MEDIAN", "MIN", "MAX"],
            default=["MEAN", "MEDIAN", "MIN", "MAX"],
            format_func=lambda s: {"MEAN": "평균", "MEDIAN": "중위값",
                                   "MIN": "최소값", "MAX": "최대값"}[s],
            key="bt_obser_stats",
        )
        obser_show_values = st.checkbox(
            "응답된 값의 분포도 함께 (단수 표처럼 값별 %/N → 계 → 통계)",
            key="bt_obser_values",
        )
        if obser_show_values and row_vars:
            # 값 종류가 많으면 표가 아주 넓어지므로 미리 알려준다
            n_values = int(df[row_vars[0]].dropna().nunique())
            if n_values > 30:
                st.warning(
                    f"'{row_vars[0]}' 는 응답된 값이 {n_values}종이라 보기가 "
                    f"{n_values}개 나옵니다. 표가 너무 넓으면 이 옵션을 끄고 통계만 내거나, "
                    "값을 묶은 변수를 쓰세요."
                )
            else:
                st.caption(f"응답된 값 {n_values}종이 보기로 들어갑니다.")
        summaries = summary_ui(row_vars[0] if row_vars else None, "_obs")
        # '표시할 통계' 에 평균이 이미 있으면 요약의 평균은 같은 숫자라 뺀다
        if "MEAN" in obser_stats:
            summaries = [s for s in summaries if s.kind != "mean"]

    st.subheader("3. 옵션")
    use_filter = st.checkbox("특정 조건만 계산 (예: 주체 = 공설)", key="bt_use_filter")
    extra_cond = None
    if use_filter:
        fc1, fc2 = st.columns(2)
        filt_disp = fc1.selectbox("필터 변수", DISPLAY_NAMES, key="bt_filt_var")
        filt_var = DISPLAY_MAP[filt_disp]
        vl = value_labels.get(filt_var, {})
        if vl:
            filt_val = fc2.selectbox("값", list(vl.values()), key="bt_filt_val")
            code = next(k for k, v in vl.items() if v == filt_val)
        else:
            code = fc2.number_input(f"{filt_var} 값", key="bt_filt_val_num")
        extra_cond = f"{filt_var}={code}"

    orient_disp = st.radio(
        "표 방향",
        ["배너를 행으로 (SPSS 산출물과 같음)", "배너를 열로"],
        horizontal=True,
        key="bt_orientation",
        help="배너를 행으로 두면 왼쪽에 권역·지자체 등이 오고 위에 문항 보기가 옵니다. "
             "열로 두면 그 반대입니다. 숫자는 같고 보는 방향만 바뀝니다.",
    )
    orientation = BANNER_ROW if orient_disp.startswith("배너를 행") else BANNER_COL

    o1, o2, o3 = st.columns(3)
    if row_type == "obser":
        obser_decimals = o1.number_input("통계 소수점 자리", 0, 4, 2, key="bt_obser_dec")
        if obser_show_values:
            show_pct = o2.checkbox("값 분포를 %로 (끄면 N)", value=True,
                                   key="bt_obser_pct")
            show_total_row = o3.checkbox("'계' 표시", value=True,
                                         key="bt_obser_total")
            decimals = 1
        else:
            show_pct, decimals, show_total_row = False, 1, False
    else:
        show_pct = o1.checkbox("퍼센트(%)로 표시", value=True, key="bt_show_pct")
        decimals = o2.number_input("소수점 자리", 0, 3, 1, key="bt_pct_dec")
        show_total_row = o3.checkbox("'계' 표시", value=True, key="bt_show_total")
        obser_decimals = 2

    # ── 유의성 검정 · 소표본 · 정렬 ──
    # 격자형이 아닌 종합표는 행끼리 비교하는 표라서 검정을 걸 수 없다.
    sig_allowed = not (row_type == "battery" and not battery_metric)
    x1, x2, x3 = st.columns(3)
    with x1:
        sig_disp = st.selectbox(
            "유의성 검정",
            ["안 함", "95%", "99%"],
            key="bt_sig_level",
            disabled=not sig_allowed,
            help="같은 배너 그룹 안에서 세그먼트끼리 비교합니다. 비율은 두 비율 "
                 "z검정, 평균은 Welch t검정. 유의하게 높은 칸에 상대 글자(a/b/c)를 "
                 "적습니다. 켜면 그 표의 값이 '42.6 b' 같은 문자로 나가서 엑셀 "
                 "계산에는 못 씁니다.",
        )
    sig = None
    if sig_allowed and sig_disp != "안 함":
        sig = SigSpec(enabled=True,
                      level=0.99 if sig_disp.startswith("99") else 0.95)
    with x2:
        min_base_show = st.number_input(
            "소표본 감추기 (사례수 미만)", 0, 500, 0, step=5,
            key="bt_min_base",
            help="0 이면 안 감춥니다. 30 으로 두면 사례수 30 미만인 배너는 값이 "
                 "'-' 로 나옵니다. N=3 에서 33.3% 가 그대로 나가는 것을 막습니다.",
        )
    with x3:
        sort_label = ("평균 높은 문항부터" if row_type == "battery"
                      else "응답 많은 보기부터")
        sort_values = st.checkbox(
            sort_label, key="bt_sort_values",
            help="'기타'·'모름'·'무응답' 계열은 응답이 많아도 맨 뒤로 보냅니다.",
        )

    if not sig_allowed and sig_disp != "안 함":
        st.caption("보기 분포형 종합표에는 검정을 걸 수 없습니다 (위 설명 참고).")

    # ── 표 제목 ──
    # 이름의 '본체' 만 직접 짓고, ' - %' / ' - N' 표시는 자동으로 붙는다.
    # 같은 문항으로 % 표와 N 표를 만들면 이름이 같아져 목록·엑셀에서 구분이
    # 안 되기 때문이다. 표시를 바꾸면 이름도 따라 바뀐다.
    #
    # 본체는 직접 고쳐 쓰면 그대로 지키되, 행 변수(또는 문항유형)가 바뀌면
    # 다른 표이므로 자동 이름으로 되돌린다.
    if not row_vars:
        auto_base = "표"
    elif row_type == "battery" and len(row_vars) > 1:
        # 종합표는 문항이 여러 개라 첫 변수명만 쓰면 무슨 표인지 알기 어렵다
        auto_base = f"{row_vars[0]} 외 {len(row_vars) - 1}문항"
    else:
        auto_base = row_vars[0]
    subject = f"{row_type}|{battery_metric}|{'|'.join(row_vars)}"
    untouched = (st.session_state.get("bt_title_base")
                 == st.session_state.get("bt_title_base_auto"))
    subject_changed = subject != st.session_state.get("bt_title_subject")

    if ("bt_title_base" not in st.session_state or untouched or subject_changed):
        st.session_state["bt_title_base"] = auto_base
    st.session_state["bt_title_base_auto"] = auto_base
    st.session_state["bt_title_subject"] = subject

    t1, t2 = st.columns([3, 1])
    base = t1.text_input(
        "표 제목", key="bt_title_base",
        help="행 변수를 바꾸면 자동 이름으로 돌아갑니다. %/N 표시는 아래 설정에 "
             "따라 자동으로 붙습니다.",
    )
    mark_title = t2.checkbox("이름에 %/N 붙이기", value=True, key="bt_title_mark")

    # 분포가 없는 수치형(통계만) 표와 평균·표준편차 격자는 %/N 이라는 게 없다
    has_pct_or_n = (
        ((row_type != "obser") or obser_show_values)
        and not (row_type == "battery" and battery_metric in ("mean", "std"))
    )
    kind = ("pct" if show_pct else "n") if (mark_title and has_pct_or_n) else None
    title = title_with_marker(base or auto_base, kind)
    st.caption(f"표 이름 → **{title}**")

    # 보기 분포형 종합표는 배너를 쓰지 않으므로 배너 없이도 만들 수 있다
    needs_banner = not (row_type == "battery" and not battery_metric)
    can_build = bool(row_vars) and (bool(banners) or not needs_banner)
    if st.button("표 만들기", type="primary", disabled=not can_build, key="bt_build"):
        try:
            if row_type == "battery":
                block = build_battery_block(
                    battery_vars=row_vars,
                    title=title,
                    banners=banners if battery_metric else None,
                    metric=battery_metric,
                    summaries=summaries,
                    show_pct=show_pct,
                    decimals=int(decimals),
                    show_total_row=show_total_row,
                    extra_cond=extra_cond,
                    orientation=orientation,
                    sig=sig,
                    min_base_show=int(min_base_show),
                    sort_rows=sort_values,
                )
            else:
                block = build_block(
                    row_type=row_type,
                    row_vars=row_vars,
                    banners=banners,
                    title=title,
                    row_ma_mode=row_ma_mode,
                    obser_stats=obser_stats or None,
                    extra_cond=extra_cond,
                    show_pct=show_pct,
                    decimals=int(decimals),
                    obser_decimals=int(obser_decimals),
                    show_total_row=show_total_row,
                    orientation=orientation,
                    obser_show_values=obser_show_values,
                    summaries=summaries,
                    sig=sig,
                    min_base_show=int(min_base_show),
                    sort_values=sort_values,
                )
            st.session_state["bt_last"] = compute_table(df, meta, block)
            st.session_state["bt_last_block"] = block
        except Exception as e:                           # noqa: BLE001
            st.error(f"계산 중 오류 — {e}")

    if "bt_last" in st.session_state:
        last = st.session_state["bt_last"]
        st.markdown(f"**{last.title}**")
        st.dataframe(result_to_frame(last), **_WIDE)
        for note in last.notes:
            st.caption(f"· {note}")
        if last.has_marks:
            st.caption(
                "글자는 같은 배너 그룹 안에서 그 칸이 유의하게 높은 상대를 뜻합니다 "
                "— '남성 (a)' 행의 `42.6 b` 는 여성(b)보다 높다는 뜻입니다."
            )
        b1, b2 = st.columns(2)
        if b1.button("아래 목록에 담기", key="bt_keep"):
            st.session_state["bt_results"].append(last)
            st.session_state["bt_blocks"].append(st.session_state["bt_last_block"])
        b2.download_button(
            "이 표만 엑셀로",
            data=write_tables_xlsx([last]),
            file_name=f"{SAV_STEM}_{safe_stem(last.title, '표')}.xlsx",
            mime=XLSX_MIME,
            key="bt_dl_one",
        )

    # ── 설정 저장 / 불러오기 ──
    #
    # 불러오기를 저장 버튼보다 먼저 처리한다. Streamlit 은 스크립트를 위에서
    # 아래로 실행하므로, 불러오기를 뒤에 두면 그 실행에서는 아래 '담아둔 표'
    # 목록이 이미 그려진 뒤라 불러온 표가 바로 보이지 않는다.
    st.divider()
    st.subheader("설정 저장 · 불러오기")
    st.caption(
        "표 정의만 저장합니다. 다음에 같은 구조의 새 .sav 를 올리고 설정을 "
        "불러오면 바뀐 데이터로 그대로 다시 계산됩니다."
    )

    s1, s2 = st.columns(2)

    with s2:
        cfg = st.file_uploader("설정 파일 (.json)", type=["json"], key="bt_cfg_up")
        if cfg is not None and st.button("불러와서 다시 계산", key="bt_load_cfg"):
            try:
                loaded, info = load_settings(cfg.getvalue())
            except ValueError as e:
                st.error(str(e))
                loaded, info = [], {}

            if loaded:
                src = info.get("source_file") or "(기록 없음)"
                when = (info.get("saved_at") or "")[:16].replace("T", " ")
                st.caption(f"설정 출처: {src}" + (f" · 저장 {when}" if when else ""))
                if info.get("source_file") and safe_stem(src) != SAV_STEM:
                    st.info(
                        f"이 설정은 '{src}' 로 만든 것이고 지금 올린 파일은 "
                        f"'{sav_file.name}' 입니다. 변수 이름이 같으면 그대로 계산됩니다."
                    )

                cols = list(df.columns)
                ok_blocks, ok_results, skipped = [], [], []
                for b in loaded:
                    gone = missing_vars(b, cols)
                    if gone:
                        skipped.append((b.title, gone))
                        continue
                    try:
                        ok_results.append(compute_table(df, meta, b))
                        ok_blocks.append(b)
                    except Exception as e:            # noqa: BLE001
                        skipped.append((b.title, [f"계산 오류: {e}"]))

                st.session_state["bt_results"] = ok_results
                st.session_state["bt_blocks"] = ok_blocks
                if ok_results:
                    st.success(f"{len(ok_results)}개 표를 지금 데이터로 다시 계산했습니다.")
                for title, why in skipped:
                    st.warning(f"'{title}' 건너뜀 — 이 .sav 에 없는 변수: {', '.join(why)}")

    with s1:
        if st.session_state["bt_blocks"]:
            st.download_button(
                f"설정 저장 ({len(st.session_state['bt_blocks'])}개 표)",
                data=blocks_to_json(st.session_state["bt_blocks"],
                                    source_file=sav_file.name),
                file_name=f"{SAV_STEM}_뱅크표설정.json",
                mime="application/json",
                key="bt_save_cfg",
            )
        else:
            st.caption("표를 목록에 담으면 저장할 수 있습니다.")

    if st.session_state["bt_results"]:
        st.divider()
        st.subheader(f"담아둔 표 {len(st.session_state['bt_results'])}개")
        st.caption("엑셀에 나가는 순서는 이 목록 순서입니다.")

        def move(i: int, step: int) -> None:
            """목록에서 표 하나를 위/아래로 옮긴다. 정의도 같이 움직인다."""
            res_list = st.session_state["bt_results"]
            blk_list = st.session_state["bt_blocks"]
            j = i + step
            if not (0 <= j < len(res_list)):
                return
            res_list[i], res_list[j] = res_list[j], res_list[i]
            if i < len(blk_list) and j < len(blk_list):
                blk_list[i], blk_list[j] = blk_list[j], blk_list[i]

        n_kept = len(st.session_state["bt_results"])
        for i, res in enumerate(st.session_state["bt_results"]):
            with st.expander(f"{i + 1}. {res.title}"):
                st.dataframe(result_to_frame(res), **_WIDE)
                for note in res.notes:
                    st.caption(f"· {note}")
                m1, m2, m3 = st.columns([1, 1, 4])
                m1.button("↑ 위로", key=f"bt_up_{i}", disabled=(i == 0),
                          on_click=move, args=(i, -1))
                m2.button("↓ 아래로", key=f"bt_down_{i}",
                          disabled=(i == n_kept - 1), on_click=move, args=(i, 1))
                if m3.button("빼기", key=f"bt_del_result_{i}"):
                    st.session_state["bt_results"].pop(i)
                    if i < len(st.session_state["bt_blocks"]):
                        st.session_state["bt_blocks"].pop(i)

        split_sheets = st.checkbox(
            "표마다 시트를 나누기", key="bt_split_sheets",
            help="끄면 'Table' 시트 하나에 표들을 위아래로 이어 붙입니다 "
                 "(SPSS 산출물과 같은 모양). 켜면 표마다 시트가 하나씩 생깁니다.",
        )
        e1, e2 = st.columns(2)
        e1.download_button(
            "담아둔 표 전체 엑셀로",
            data=write_tables_xlsx(st.session_state["bt_results"],
                                   split_sheets=split_sheets),
            file_name=f"{SAV_STEM}_뱅크표.xlsx",
            mime=XLSX_MIME,
            key="bt_dl_all",
        )
        if st.session_state["bt_blocks"]:
            # 화면에서 만든 표를 양식으로 빼두면, 엑셀에서 고쳐 다시 올릴 수 있다
            e2.download_button(
                "엑셀 양식으로 내보내기",
                data=blocks_to_form(st.session_state["bt_blocks"], df, meta),
                file_name=f"{SAV_STEM}_뱅크표양식.xlsx",
                mime=XLSX_MIME,
                key="bt_dl_form",
                help="이 표들이 채워진 양식이 나옵니다. 엑셀에서 고쳐 '엑셀 폼으로 만들기' 탭에 다시 올리세요.",
            )

        # ── 차수 비교 ──
        if st.session_state["bt_blocks"]:
            st.divider()
            with st.expander("차수 비교 — 지난 차수 파일과 나란히 보기"):
                st.caption(
                    "지난 차수의 .sav 를 올리면 담아둔 표마다 **이번 차수 · 지난 차수 · "
                    "증감(%p)** 세 표가 나옵니다. 표 정의는 그대로 쓰므로 변수 이름이 "
                    "같아야 합니다."
                )
                prev_sav = st.file_uploader("지난 차수 데이터 (.sav)", type=["sav"],
                                            key="bt_prev_sav")
                w1, w2 = st.columns(2)
                label_now = w1.text_input("이번 차수 이름", value="이번 차수",
                                          key="bt_wave_now")
                label_bef = w2.text_input("지난 차수 이름", value="지난 차수",
                                          key="bt_wave_bef")

                if prev_sav is not None:
                    try:
                        df_b, meta_b = load_sav(prev_sav.getvalue())
                    except Exception as e:               # noqa: BLE001
                        st.error(f"지난 차수 파일을 읽지 못했습니다 — {e}")
                        df_b = None

                    if df_b is not None:
                        st.caption(
                            f"{prev_sav.name} · {len(df_b):,}행 × {len(df_b.columns)}열"
                        )
                        wave_results, wave_problems = compare_waves(
                            df, meta, df_b, meta_b,
                            st.session_state["bt_blocks"],
                            label_now=label_now or "이번 차수",
                            label_before=label_bef or "지난 차수",
                        )
                        for msg in wave_problems:
                            st.warning(msg)
                        for res in wave_results:
                            st.markdown(f"**{res.title}**")
                            st.dataframe(result_to_frame(res), **_WIDE)
                        st.download_button(
                            f"차수 비교 엑셀로 ({len(wave_results)}개 표)",
                            data=write_tables_xlsx(wave_results,
                                                   split_sheets=split_sheets),
                            file_name=f"{SAV_STEM}_차수비교.xlsx",
                            mime=XLSX_MIME,
                            key="bt_dl_waves",
                        )

# =============================================================================
# 엑셀 폼으로 만들기 (.sav + 엑셀 양식)
# =============================================================================
with tab_form:
    st.write(
        "신텍스 없이, 엑셀 양식에 표를 한 줄씩 적어 올리면 그대로 계산합니다. "
        "양식을 내려받아 채운 뒤 다시 올리세요."
    )

    @st.cache_data(show_spinner=False)
    def _filled_form(file_bytes: bytes):
        """.sav 를 보고 자동으로 채운 양식. 같은 파일이면 다시 만들지 않는다."""
        return write_filled_form(df, meta)

    filled, fill_notes = _filled_form(sav_file.getvalue())

    f1, f2, f3 = st.columns(3)
    with f1:
        st.download_button(
            "① 자동 채운 양식 내려받기",
            data=filled,
            file_name=f"{SAV_STEM}_뱅크표양식.xlsx",
            mime=XLSX_MIME,
            key="bt_form_filled",
            type="primary",
            help="올린 .sav 의 변수 라벨·값 라벨을 보고 표 목록을 미리 채워 둡니다.",
        )
        st.caption("문항이 채워진 상태로 나옵니다")
    with f2:
        st.download_button(
            "빈 양식 내려받기",
            data=write_form_template(df, meta),
            file_name=f"{SAV_STEM}_뱅크표양식_빈것.xlsx",
            mime=XLSX_MIME,
            key="bt_form_tpl",
        )
        st.caption("직접 처음부터 적을 때")
    with f3:
        form_file = st.file_uploader("② 채운 양식 올리기 (.xlsx)", type=["xlsx"],
                                     key="bt_form_up")

    if fill_notes:
        with st.expander("자동 채우기가 무엇을 넣고 뺐는지"):
            for note in fill_notes:
                st.write(f"- {note}")

    if form_file is None:
        st.info(
            "①로 자동 채운 양식을 받아 엑셀에서 확인·수정한 뒤 ②로 올리면 됩니다. "
            "배너는 후보만 넣어 뒀으니 실제로 쓸 것만 남기세요. "
            "각 칸에 무엇을 적는지는 양식의 '사용법' 시트에 있습니다."
        )
    else:
        try:
            form_blocks, problems = read_form(form_file.getvalue(), df, meta)
        except ValueError as e:
            st.error(str(e))
            form_blocks, problems = [], []

        if problems:
            with st.expander(f"확인할 것 {len(problems)}건", expanded=not form_blocks):
                for msg in problems:
                    st.warning(msg)

        if form_blocks:
            st.success(f"표 {len(form_blocks)}개를 읽었습니다.")

            computed_form = []
            for b in form_blocks:
                try:
                    computed_form.append(compute_table(df, meta, b))
                except Exception as e:                   # noqa: BLE001
                    st.error(f"'{b.title}' 계산 중 오류 — {e}")

            if computed_form:
                form_split = st.checkbox("표마다 시트를 나누기",
                                         key="bt_form_split")
                g1, g2 = st.columns(2)
                g1.download_button(
                    "표 전체 엑셀로 (목차 + Table 시트)",
                    data=write_tables_xlsx(computed_form,
                                           split_sheets=form_split),
                    file_name=f"{SAV_STEM}_뱅크표.xlsx",
                    mime=XLSX_MIME,
                    key="bt_form_dl",
                )
                # 폼으로 만든 표도 설정으로 저장해 두면 다음엔 폼 없이 쓸 수 있다
                g2.download_button(
                    "이 표들을 설정으로 저장",
                    data=blocks_to_json(
                        form_blocks,
                        source_file=form_file.name,
                        note=f"{sav_file.name} 로 계산",
                    ),
                    file_name=f"{safe_stem(form_file.name)}_뱅크표설정.json",
                    mime="application/json",
                    key="bt_form_cfg",
                )

                for res in computed_form:
                    st.markdown(f"**{res.title}**")
                    st.dataframe(result_to_frame(res), **_WIDE)
                    for note in res.notes:
                        st.caption(f"· {note}")


# =============================================================================
# 빈도 · 교차표 (빠르게 훑어볼 때)
# =============================================================================
with tab_quick:
    freq_mode, cross_mode = st.tabs(["빈도표 (여러 변수 한 번에)", "교차표"])

    def apply_labels(series: pd.Series, varname: str) -> pd.Series:
        vl = value_labels.get(varname)
        return series if not vl else series.map(lambda v: vl.get(v, v))

    # ── 빈도표: 변수를 여러 개 골라 한 번에 ──
    with freq_mode:
        st.caption(
            "고른 변수마다 빈도표를 하나씩 만듭니다. 값 라벨에 정의된 보기는 "
            "응답이 0이어도 나오고, 라벨에 없는 코드가 데이터에 있으면 따로 알려 줍니다."
        )

        # 자주 쓰는 묶음은 버튼으로 골라 넣는다. 변수가 수십~수백 개라
        # 매번 하나씩 고르는 것이 이 탭에서 제일 번거로운 일이다.
        st.session_state.setdefault("bt_freq_vars", [])

        def set_freq_vars(names: list[str]) -> None:
            st.session_state["bt_freq_vars"] = [label_for(c) for c in names]

        labelled = [c for c in df.columns if value_labels.get(c)]
        numeric_only = [c for c in df.columns
                        if not value_labels.get(c)
                        and pd.api.types.is_numeric_dtype(df[c])]

        p1, p2, p3, p4 = st.columns(4)
        p1.button("전체", key="bt_freq_all", on_click=set_freq_vars,
                  args=(list(df.columns),))
        p2.button(f"값 라벨 있는 것만 ({len(labelled)})", key="bt_freq_lab",
                  on_click=set_freq_vars, args=(labelled,))
        p3.button(f"숫자 변수만 ({len(numeric_only)})", key="bt_freq_num",
                  on_click=set_freq_vars, args=(numeric_only,))
        p4.button("비우기", key="bt_freq_clear", on_click=set_freq_vars, args=([],))

        freq_disp = st.multiselect("빈도표를 뽑을 변수", DISPLAY_NAMES,
                                   key="bt_freq_vars")
        freq_vars = to_vars(freq_disp)

        q1, q2, q3 = st.columns(3)
        freq_missing = q1.checkbox("무응답(결측) 행 표시", value=True,
                                   key="bt_freq_missing")
        freq_sort = q2.checkbox("응답 많은 보기부터", key="bt_freq_sort",
                                help="'기타'·'모름'·'무응답' 계열은 맨 뒤로 보냅니다.")
        freq_split = q3.checkbox("변수마다 시트 나누기", key="bt_freq_split")

        if not freq_vars:
            st.info("위에서 변수를 고르거나 '전체' 같은 버튼을 눌러 주세요.")
        else:
            freq_tables = compute_frequencies(
                df, meta, freq_vars,
                show_missing=freq_missing, sort_by_count=freq_sort,
            )
            flagged = [t for t in freq_tables
                       if any("값 라벨에 없는 코드" in n for n in t.notes)]
            if flagged:
                st.warning(
                    "값 라벨에 없는 코드가 있는 변수 "
                    f"{len(flagged)}개 — {', '.join(t.var for t in flagged)}. "
                    "코딩 오류이거나 라벨을 안 붙인 것이니 확인해 보세요."
                )

            st.download_button(
                f"빈도표 {len(freq_tables)}개 엑셀로",
                data=write_freq_xlsx(freq_tables, split_sheets=freq_split),
                file_name=f"{SAV_STEM}_빈도표.xlsx",
                mime=XLSX_MIME,
                key="bt_freq_dl",
                type="primary",
            )

            for t in freq_tables:
                with st.expander(t.title, expanded=len(freq_tables) == 1):
                    frame = freq_to_frame(t)
                    if frame.empty:
                        st.caption("표로 만들 값이 없습니다.")
                    else:
                        st.dataframe(frame, **_WIDE)
                    if t.stats:
                        st.caption(" · ".join(
                            f"{k} {v:,}" for k, v in t.stats.items() if v is not None
                        ))
                    st.caption(
                        f"전체 {t.total_n:,} · 유효 {t.valid_n:,} · 무응답 {t.missing_n:,}"
                    )
                    for note in t.notes:
                        st.caption(f"· {note}")

    # ── 교차표: 두 변수 ──
    with cross_mode:
        st.caption("값 라벨을 붙인 교차표를 빠르게 봅니다.")
        row_disp = st.selectbox("행 변수", DISPLAY_NAMES, key="bt_q_row")
        col_disp = st.selectbox("열 변수", DISPLAY_NAMES, key="bt_q_col")
        row_var, col_var = DISPLAY_MAP[row_disp], DISPLAY_MAP[col_disp]

        mode = st.radio("표시", ["빈도(N)", "열 기준 %", "행 기준 %"],
                        horizontal=True, key="bt_q_mode")
        ct = pd.crosstab(apply_labels(df[row_var], row_var),
                         apply_labels(df[col_var], col_var))
        if mode == "열 기준 %":
            ct = (ct / ct.sum(axis=0) * 100).round(1)
        elif mode == "행 기준 %":
            ct = (ct.div(ct.sum(axis=1), axis=0) * 100).round(1)
        st.dataframe(ct, **_WIDE)
        st.download_button("이 표 CSV로", data=ct.to_csv().encode("utf-8-sig"),
                           file_name=f"{row_var}_x_{col_var}.csv",
                           mime="text/csv", key="bt_q_dl")


# =============================================================================
# 신텍스로 한 번에 (.sav + .sps)
# =============================================================================
with tab_syntax:
    st.caption(
        "Embrain 'Table' 매크로 신텍스를 읽어 정의된 표를 그대로 계산합니다. "
        "매크로 원본 정의를 보지 못한 상태에서 문법을 역추적한 것이라, "
        "실제 업무에 쓰기 전 SPSS 결과와 숫자를 한 번 대조해 주세요."
    )

    sps_file = st.file_uploader("Table 매크로 신텍스 (.sps)", type=["sps"], key="bt_sps")

    if sps_file is None:
        st.info("기존에 쓰던 .sps 신텍스를 올리면 그 안에 정의된 표를 그대로 계산합니다.")
    else:
        try:
            blocks = parse_sps(load_sps(sps_file.getvalue()))
        except Exception as e:                           # noqa: BLE001
            st.error(f"신텍스를 읽지 못했습니다 — {e}")
            blocks = []

        if not blocks:
            st.warning(
                "표 블록을 찾지 못했습니다. 이 도구는 'Table ... /mrg= /table= "
                "/statistics= /title=' 형태의 매크로 호출만 인식합니다."
            )
        else:
            st.write(f"표 **{len(blocks)}개**를 찾았습니다.")
            titles = [f"{i + 1}. {b.title}" for i, b in enumerate(blocks)]
            picked = st.multiselect("볼 표 (비우면 전체)", titles, key="bt_syn_pick")
            targets = blocks if not picked else [
                b for t, b in zip(titles, blocks) if t in picked
            ]

            syn_orient = st.radio(
                "표 방향",
                ["배너를 행으로 (SPSS 산출물과 같음)", "배너를 열로"],
                horizontal=True,
                key="bt_syn_orientation",
            )
            y1, y2, y3 = st.columns(3)
            syn_sig_disp = y1.selectbox("유의성 검정", ["안 함", "95%", "99%"],
                                        key="bt_syn_sig")
            syn_min_base = y2.number_input("소표본 감추기 (사례수 미만)", 0, 500, 0,
                                           step=5, key="bt_syn_min_base")
            syn_split = y3.checkbox("표마다 시트를 나누기", key="bt_syn_split")

            syn_sig = None
            if syn_sig_disp != "안 함":
                syn_sig = SigSpec(
                    enabled=True,
                    level=0.99 if syn_sig_disp.startswith("99") else 0.95,
                )
            for b in targets:
                b.orientation = (
                    BANNER_ROW if syn_orient.startswith("배너를 행") else BANNER_COL
                )
                b.sig = syn_sig
                b.min_base_show = int(syn_min_base)

            computed = []
            for b in targets:
                try:
                    computed.append(compute_table(df, meta, b))
                except Exception as e:                   # noqa: BLE001
                    st.error(f"'{b.title}' 계산 중 오류 — {e}")

            if computed:
                d1, d2 = st.columns(2)
                d1.download_button(
                    "전체 엑셀로 (목차 + Table 시트)",
                    data=write_tables_xlsx(computed, split_sheets=syn_split),
                    file_name=f"{safe_stem(sps_file.name)}_뱅크표.xlsx",
                    mime=XLSX_MIME,
                    key="bt_syn_dl",
                )
                # 신텍스를 설정으로 저장해 두면, 다음엔 .sps 없이 새 .sav 에
                # 바로 적용할 수 있다.
                d2.download_button(
                    "이 표들을 설정으로 저장",
                    data=blocks_to_json(
                        targets,
                        source_file=sps_file.name,
                        note=f"{sav_file.name} 로 계산한 신텍스 표",
                    ),
                    file_name=f"{safe_stem(sps_file.name)}_뱅크표설정.json",
                    mime="application/json",
                    key="bt_syn_cfg",
                )
                for res in computed:
                    st.markdown(f"**{res.title}**")
                    st.dataframe(result_to_frame(res), **_WIDE)
                    for note in res.notes:
                        st.caption(f"· {note}")

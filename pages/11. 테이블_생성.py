# 테이블_생성.py
#
# SPSS .sav 파일로 테이블을 만드는 페이지입니다.
# 계산 로직은 banner_table_engine.py 에 있습니다.
#
# set_page_config 와 비밀번호 확인은 Home.py 가 처리하므로 여기에 두지 않습니다.
#
# ── session_state 키 ─────────────────────────────────────────────────
#   멀티페이지 앱은 session_state 를 모든 페이지가 함께 씁니다.
#   다른 페이지와 겹치지 않도록 이 페이지의 키는 모두 'bt_' 로 시작합니다.

import collections
import hashlib
import re
import tempfile
from pathlib import Path

import pandas as pd
import pyreadstat
import streamlit as st

from banner_table_engine import (
    BANNER_COL,
    BANNER_ROW,
    TITLE_FILL,
    UNDEF_FILL,
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
import table_guide as tg
from table_picker import (
    KIND_LABEL,
    KIND_MULTI,
    KIND_OBSER,
    KIND_SINGLE,
    LABEL_KIND,
    build_plan,
    detect_items,
    find_batteries,
    merge_picks,
    parse_spec,
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


st.header("테이블 생성")
st.caption(
    "SPSS .sav 를 올리고 배너와 행 변수를 골라 테이블을 만듭니다. "
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


# ── 문항 목록 (여러 문항 한 번에 · 복수응답 세트 묶기) ──
# 복제본 찾기가 열마다 해시를 구하므로 파일당 한 번만 한다.
@st.cache_data(show_spinner=False)
def _pick_items(file_bytes: bytes):
    return detect_items(df, meta)


@st.cache_data(show_spinner=False)
def _battery_sets(file_bytes: bytes):
    return find_batteries(df, meta)


PICK_ITEMS, PICK_NOTES = _pick_items(sav_file.getvalue())
PICK_BY_DISPLAY = {it.display: it for it in PICK_ITEMS}
MA_SETS = {it.key: it.vars for it in PICK_ITEMS if it.kind == KIND_MULTI}
PICK_BATCH = "여러 문항 한 번에"
PICK_ONE = "한 문항씩 자세히"


st.session_state.setdefault("bt_merge_banners", [])
st.session_state.setdefault("bt_results", [])
st.session_state.setdefault("bt_blocks", [])      # 담아둔 표의 '정의' (설정 저장용)

tab_manual, tab_form, tab_quick, tab_syntax, tab_guide = st.tabs(
    ["변수 골라서 만들기", "엑셀 폼으로 만들기", "빈도 · 교차표", "신텍스로 한 번에",
     "가이드로 신텍스 만들기"]
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

        # 삭제는 on_click 콜백으로 한다. 본문에서 pop 하면 지금 돌고 있는
        # 루프의 목록이 줄어들어 바로 뒤 항목 하나가 이번 화면에서 빠진다.
        # 콜백은 다시 그리기 전에 실행되므로 그런 일이 없다.
        def _del_merge(idx: int) -> None:
            lst = st.session_state["bt_merge_banners"]
            if 0 <= idx < len(lst):
                lst.pop(idx)

        for i, g in enumerate(list(st.session_state["bt_merge_banners"])):
            gc1, gc2 = st.columns([6, 1])
            gc1.write(f"**{g['label']}** — {', '.join(g['varlist'])}")
            gc2.button("삭제", key=f"bt_del_merge_{i}",
                       on_click=_del_merge, args=(i,))

    banners = [BannerSpec(kind="single", var=v) for v in to_vars(banner_disp)] + [
        BannerSpec(kind="merge", label=g["label"], varlist=g["varlist"])
        for g in st.session_state["bt_merge_banners"]
    ]
    if not banners:
        st.info("배너로 쓸 변수를 최소 1개 골라 주세요. '전체' 컬럼은 항상 자동으로 들어갑니다.")

    st.subheader("2. 행 변수")
    pick_mode = st.radio(
        "만드는 방식", [PICK_BATCH, PICK_ONE], horizontal=True, key="bt_pick_mode",
        help="여러 문항 한 번에 : 고른 문항마다 표가 하나씩 만들어져 바로 목록에 "
             "담깁니다. 유형은 자동으로 판단하고, 틀린 것만 표에서 고칩니다.\n\n"
             "한 문항씩 자세히 : 척도 종합표, 수치형 값 분포처럼 세밀한 설정이 "
             "필요할 때 씁니다.",
    )
    batch = pick_mode == PICK_BATCH
    batch_plan: list[dict] = []

    def set_picker(sets: dict[str, list[str]], target_key: str, what: str) -> None:
        """자동으로 찾은 묶음을 골라 아래 변수 칸을 한 번에 채운다.

        콜백에서 채운다 — 위젯이 그려진 뒤에 session_state 를 바꾸면
        Streamlit 이 오류를 낸다. 고른 뒤에도 아래 칸에서 더하고 뺄 수 있다.
        """
        if not sets:
            return
        opts = ["(직접 고르기)"] + [f"{k} ({len(v)}개: {v[0]} ~ {v[-1]})"
                                  for k, v in sets.items()]
        keys = [None] + list(sets)

        def _fill() -> None:
            k = keys[opts.index(st.session_state[f"{target_key}_set"])]
            if k:
                st.session_state[target_key] = [label_for(v) for v in sets[k]]

        st.selectbox(f"자동으로 찾은 {what}에서 한 번에 고르기", opts,
                     key=f"{target_key}_set", on_change=_fill)

    if batch:
        row_type = "batch"
        row_vars = []
        row_ma_mode = "category"
        obser_stats = ["MEAN", "MEDIAN", "MIN", "MAX"]
        obser_show_values = False
        summaries = []
        battery_metric = None

        typed = st.text_area(
            "문항 입력 — 쉼표·줄바꿈으로 구분, 범위는 ~ (예: SQ1, Q1_1~Q1_5, Q3)",
            key="bt_batch_text", height=80,
            placeholder="SQ1, Q1_1~Q1_5\nQ3",
            help="설문지의 문항 번호를 붙여 넣어도 됩니다. 대소문자는 가리지 "
                 "않습니다. 'Q3' 처럼 묶음 이름만 쓰면 Q3_1, Q3_2 … 전체가 "
                 "들어가고, 복수응답 세트면 표 하나로 묶입니다.",
        )
        typed_items, typed_problems = parse_spec(typed, list(df.columns), PICK_ITEMS)
        for msg in typed_problems:
            st.warning(msg)
        picked_disp = st.multiselect(
            "목록에서 고르기 — 복수응답 세트는 한 줄로 묶여 있습니다",
            list(PICK_BY_DISPLAY), key="bt_batch_pick",
        )
        for note in PICK_NOTES:
            st.caption(note)
        chosen = merge_picks(typed_items,
                             [PICK_BY_DISPLAY[d] for d in picked_disp
                              if d in PICK_BY_DISPLAY])

        if chosen:
            def _vars_txt(vs: list[str]) -> str:
                return ", ".join(vs) if len(vs) <= 3 else f"{vs[0]} ~ {vs[-1]} ({len(vs)}개)"

            plan_df = pd.DataFrame({
                "문항": [it.key for it in chosen],
                "유형": [KIND_LABEL[it.kind] for it in chosen],
                "표 제목": [it.key for it in chosen],
                "척도 요약": [it.summary if it.kind == KIND_SINGLE else ""
                          for it in chosen],
                "문항 문구": [it.label for it in chosen],
                "변수": [_vars_txt(it.vars) for it in chosen],
            })
            # 고른 문항이 바뀌면 표를 새로 받는다. 같은 key 를 쓰면 예전에
            # 고친 칸이 줄 번호만 보고 엉뚱한 문항에 붙는다.
            _sig = hashlib.md5("|".join(",".join(it.vars) for it in chosen)
                               .encode()).hexdigest()[:12]
            ed = st.data_editor(
                plan_df, key=f"bt_batch_ed_{_sig}", hide_index=True,
                disabled=["문항", "문항 문구", "변수"],
                column_config={
                    "유형": st.column_config.SelectboxColumn(
                        options=list(LABEL_KIND), required=True,
                        help="자동으로 판단한 유형입니다. 틀렸으면 고르세요. "
                             "여러 변수를 단수·수치형으로 바꾸면 변수마다 표가 "
                             "따로 만들어집니다."),
                    "척도 요약": st.column_config.TextColumn(
                        help="예: 상2,중,하2,평균 — 단수 문항에만 붙습니다. "
                             "비우면 붙이지 않습니다."),
                },
                **_WIDE,
            )
            batch_plan = build_plan(
                chosen,
                [LABEL_KIND.get(k, "") for k in ed["유형"]],
                ed["표 제목"].fillna("").tolist(),
                ed["척도 요약"].fillna("").tolist(),
            )
            row_vars = [v for p in batch_plan for v in p["vars"]]
            st.caption(f"표 **{len(batch_plan)}개**가 만들어집니다. 제목·유형·척도 "
                       "요약은 위 표에서 바로 고칠 수 있습니다.")
        else:
            st.info("문항을 입력하거나 목록에서 고르세요.")

    else:
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
            set_picker(MA_SETS, "bt_row_multi", "복수응답 세트")
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
            set_picker(_battery_sets(sav_file.getvalue()), "bt_row_battery",
                       "보기가 같은 문항 묶음")
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

    # ── 옵션 배치 규칙 ──
    # 한 줄에는 **같은 종류의 위젯만** 놓는다. 체크박스는 라벨이 위에 붙지
    # 않고 숫자·선택 상자는 붙어서, 한 줄에 섞으면 기준선이 어긋나 지저분해
    # 보인다. 그래서 '표시'(체크박스) 줄과 '소수점 · 검정'(입력) 줄로 나눴다.

    # 격자형이 아닌 종합표는 행끼리 비교하는 표라서 검정을 걸 수 없다.
    sig_allowed = not (row_type == "battery" and not battery_metric)
    # 분포가 없는 수치형(통계만) 표와 평균·표준편차 격자는 %/N 이라는 게 없다
    has_pct_or_n = (
        ((row_type != "obser") or obser_show_values)
        and not (row_type == "battery" and battery_metric in ("mean", "std"))
    )

    st.markdown("**표시**")
    d1, d2, d3, d4 = st.columns(4)
    if row_type == "obser":
        if obser_show_values:
            show_pct = d1.checkbox("값 분포를 %로 (끄면 N)", value=True,
                                   key="bt_obser_pct")
            show_total_row = d2.checkbox("'계' 표시", value=True,
                                         key="bt_obser_total")
        else:
            d1.checkbox("값 분포를 %로 (끄면 N)", value=True, disabled=True,
                        key="bt_obser_pct_off",
                        help="위에서 '응답된 값의 분포도 함께' 를 켜면 씁니다.")
            d2.checkbox("'계' 표시", value=False, disabled=True,
                        key="bt_obser_total_off",
                        help="분포가 없으면 더할 것이 없습니다.")
            show_pct, show_total_row = False, False
        decimals = 1
    else:
        show_pct = d1.checkbox("퍼센트(%)로 표시", value=True, key="bt_show_pct")
        show_total_row = d2.checkbox("'계' 표시", value=True, key="bt_show_total")

    sort_label = ("평균 높은 문항부터" if row_type == "battery"
                  else "응답 많은 보기부터")
    sort_values = d3.checkbox(
        sort_label, key="bt_sort_values",
        help="'기타'·'모름'·'무응답' 계열은 응답이 많아도 맨 뒤로 보냅니다.",
    )
    mark_title = d4.checkbox(
        "표 이름에 %/N 붙이기", value=True, key="bt_title_mark",
        disabled=not has_pct_or_n,
        help="같은 문항으로 % 표와 N 표를 만들면 이름이 같아져 목록·엑셀에서 "
             "구분이 안 되므로, 이름 끝에 표시를 붙입니다.",
    )

    st.markdown("**소수점 · 검정**")
    n1, n2, n3 = st.columns(3)
    if row_type == "obser":
        obser_decimals = n1.number_input("통계 소수점 자리", 0, 4, 2,
                                         key="bt_obser_dec")
    else:
        decimals = n1.number_input("소수점 자리", 0, 3, 1, key="bt_pct_dec")
        obser_decimals = 2
    sig_disp = n2.selectbox(
        "유의성 검정",
        ["안 함", "95%", "99%"],
        key="bt_sig_level",
        disabled=not sig_allowed,
        help="같은 배너 그룹 안에서 세그먼트끼리 비교합니다. 비율은 두 비율 "
             "z검정, 평균은 Welch t검정. 유의하게 높은 칸에 상대 글자(a/b/c)를 "
             "적습니다. 켜면 그 표의 값이 '42.6 b' 같은 문자로 나가서 엑셀 "
             "계산에는 못 씁니다.",
    )
    min_base_show = n3.number_input(
        "소표본 감추기 (사례수 미만)", 0, 500, 0, step=5,
        key="bt_min_base",
        help="0 이면 안 감춥니다. 30 으로 두면 사례수 30 미만인 배너는 값이 "
             "'-' 로 나옵니다. N=3 에서 33.3% 가 그대로 나가는 것을 막습니다.",
    )

    sig = None
    if sig_allowed and sig_disp != "안 함":
        sig = SigSpec(enabled=True,
                      level=0.99 if sig_disp.startswith("99") else 0.95)
    if not sig_allowed and sig_disp != "안 함":
        st.caption("보기 분포형 종합표에는 검정을 걸 수 없습니다 (위 설명 참고).")

    # 표를 목록에 담는다. 만들면 바로 담기므로 '담기' 단계가 따로 없다.
    def keep_table(res, block) -> None:
        st.session_state["bt_results"].append(res)
        st.session_state["bt_blocks"].append(block)

    if batch:
        title = ""
        can_build = bool(batch_plan) and bool(banners)
        if batch_plan and not banners:
            st.caption("위에서 배너를 고르면 만들 수 있습니다.")
        if st.button(f"표 {len(batch_plan)}개 만들어 목록에 담기", type="primary",
                     disabled=not can_build, key="bt_build_batch"):
            made, failed, warned = 0, [], []
            bar = st.progress(0.0, text="표 계산 중…")
            for n, plan in enumerate(batch_plan, start=1):
                bar.progress(n / len(batch_plan),
                             text=f"표 계산 중… {n}/{len(batch_plan)}")
                is_obs = plan["kind"] == KIND_OBSER
                try:
                    summ = []
                    if plan["summary"]:
                        v0 = plan["vars"][0]
                        codes = (sorted(value_labels.get(v0, {}).keys())
                                 or sorted(df[v0].dropna().unique().tolist()))
                        summ, sum_problems = parse_summary_spec(
                            plan["summary"], codes, decimals=1)
                        warned += [f"{plan['title']}: {m}" for m in sum_problems]
                    mark = (None if (is_obs or not mark_title)
                            else ("pct" if show_pct else "n"))
                    block = build_block(
                        row_type=plan["kind"],
                        row_vars=plan["vars"],
                        banners=banners,
                        title=title_with_marker(plan["title"], mark),
                        row_ma_mode="category",
                        obser_stats=obser_stats if is_obs else None,
                        extra_cond=extra_cond,
                        show_pct=show_pct and not is_obs,
                        decimals=int(decimals),
                        obser_decimals=int(obser_decimals),
                        show_total_row=show_total_row and not is_obs,
                        orientation=orientation,
                        obser_show_values=False,
                        summaries=summ,
                        sig=sig,
                        min_base_show=int(min_base_show),
                        sort_values=sort_values,
                    )
                    keep_table(compute_table(df, meta, block), block)
                    made += 1
                except Exception as e:                   # noqa: BLE001
                    failed.append(f"{plan['title']} — {e}")
            bar.empty()
            st.session_state["bt_batch_msg"] = (made, failed, warned)

        msg = st.session_state.pop("bt_batch_msg", None)
        if msg:
            made, failed, warned = msg
            if made:
                st.success(f"표 {made}개를 만들어 아래 '담아둔 표' 에 넣었습니다.")
            for m in warned:
                st.warning(m)
            for m in failed:
                st.error(f"만들지 못한 표: {m}")
    else:
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

        # 제목 칸은 화면 폭을 다 쓴다. 여기에 도움말 아이콘(?)을 붙이면 라벨과
        # 멀리 떨어진 오른쪽 끝에 혼자 앉아 무엇에 붙은 설명인지 알 수 없다.
        # 짧은 설명이라 라벨에 넣었다.
        base = st.text_input("표 제목 — 행 변수를 바꾸면 자동 이름으로 돌아갑니다",
                             key="bt_title_base")
        kind = ("pct" if show_pct else "n") if (mark_title and has_pct_or_n) else None
        title = title_with_marker(base or auto_base, kind)
        st.caption(f"표 이름 → **{title}**")

        # 보기 분포형 종합표는 배너를 쓰지 않으므로 배너 없이도 만들 수 있다
        needs_banner = not (row_type == "battery" and not battery_metric)
        can_build = bool(row_vars) and (bool(banners) or not needs_banner)
        if st.button("표 만들고 목록에 담기", type="primary", disabled=not can_build,
                     key="bt_build"):
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
                res = compute_table(df, meta, block)
                keep_table(res, block)
                st.session_state["bt_last"] = res
            except Exception as e:                       # noqa: BLE001
                st.error(f"계산 중 오류 — {e}")

        if "bt_last" in st.session_state:
            last = st.session_state["bt_last"]
            kept = st.session_state["bt_results"]
            # 목록에서 몇 번째인지. 빼거나 옮겼으면 달라지므로 매번 찾는다.
            pos = next((i for i, r in enumerate(kept) if r is last), None)
            st.markdown(f"**{last.title}**")
            st.dataframe(result_to_frame(last), **_WIDE)
            for note in last.notes:
                st.caption(f"· {note}")
            if last.has_marks:
                st.caption(
                    "글자는 같은 배너 그룹 안에서 그 칸이 유의하게 높은 상대를 뜻합니다 "
                    "— '남성 (a)' 행의 `42.6 b` 는 여성(b)보다 높다는 뜻입니다."
                )

            def _undo_last() -> None:
                res_list = st.session_state["bt_results"]
                blk_list = st.session_state["bt_blocks"]
                i = next((i for i, r in enumerate(res_list)
                          if r is st.session_state.get("bt_last")), None)
                if i is not None:
                    res_list.pop(i)
                    if i < len(blk_list):
                        blk_list.pop(i)
                st.session_state.pop("bt_last", None)

            b1, b2, b3 = st.columns([2, 1, 1])
            b1.caption(f"✅ 담아둔 표 {pos + 1}번에 들어갔습니다."
                       if pos is not None else "목록에서 뺀 표입니다.")
            b2.button("방금 표 빼기", key="bt_undo_last", on_click=_undo_last,
                      disabled=pos is None)
            b3.download_button(
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
                file_name=f"{SAV_STEM}_테이블설정.json",
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

        def drop(i: int) -> None:
            """목록에서 표 하나를 뺀다. move 와 같은 이유로 on_click 에서 한다.
            본문에서 pop 하면 이번 화면에서 바로 뒤 표가 그려지지 않는다."""
            res_list = st.session_state["bt_results"]
            blk_list = st.session_state["bt_blocks"]
            if 0 <= i < len(res_list):
                res_list.pop(i)
                if i < len(blk_list):
                    blk_list.pop(i)

        n_kept = len(st.session_state["bt_results"])
        for i, res in enumerate(list(st.session_state["bt_results"])):
            with st.expander(f"{i + 1}. {res.title}"):
                st.dataframe(result_to_frame(res), **_WIDE)
                for note in res.notes:
                    st.caption(f"· {note}")
                m1, m2, m3 = st.columns([1, 1, 4])
                m1.button("↑ 위로", key=f"bt_up_{i}", disabled=(i == 0),
                          on_click=move, args=(i, -1))
                m2.button("↓ 아래로", key=f"bt_down_{i}",
                          disabled=(i == n_kept - 1), on_click=move, args=(i, 1))
                m3.button("빼기", key=f"bt_del_result_{i}",
                          on_click=drop, args=(i,))

        s1, s2 = st.columns([3, 1])
        split_sheets = s1.checkbox(
            "표마다 시트를 나누기", key="bt_split_sheets",
            help="끄면 'Table' 시트 하나에 표들을 위아래로 이어 붙입니다 "
                 "(SPSS 산출물과 같은 모양). 켜면 표마다 시트가 하나씩 생깁니다.",
        )
        # 테이블은 SPSS 산출물과 같은 모양이 기본이라 색 없이(흰색) 시작한다
        bank_fill = s2.color_picker(
            "표 제목 줄 색", value="#FFFFFF", key="bt_bank_fill",
            help="흰색이면 색을 넣지 않습니다 (SPSS 산출물과 같은 모양).",
        )
        # 아래 옵션 줄들과 같은 3칸을 써서 체크박스가 세로로 줄을 맞춘다
        e1, e2, _e3 = st.columns(3)
        e1.download_button(
            "담아둔 표 전체 엑셀로",
            data=write_tables_xlsx(st.session_state["bt_results"],
                                   split_sheets=split_sheets,
                                   title_fill=bank_fill),
            file_name=f"{SAV_STEM}_테이블.xlsx",
            mime=XLSX_MIME,
            key="bt_dl_all",
        )
        if st.session_state["bt_blocks"]:
            # 화면에서 만든 표를 양식으로 빼두면, 엑셀에서 고쳐 다시 올릴 수 있다
            e2.download_button(
                "엑셀 양식으로 내보내기",
                data=blocks_to_form(st.session_state["bt_blocks"], df, meta),
                file_name=f"{SAV_STEM}_테이블양식.xlsx",
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
                                                   split_sheets=split_sheets,
                                                   title_fill=bank_fill),
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
            file_name=f"{SAV_STEM}_테이블양식.xlsx",
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
            file_name=f"{SAV_STEM}_테이블양식_빈것.xlsx",
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
                fs1, fs2 = st.columns([3, 1])
                form_split = fs1.checkbox("표마다 시트를 나누기",
                                          key="bt_form_split")
                form_fill = fs2.color_picker(
                    "표 제목 줄 색", value="#FFFFFF", key="bt_form_fill",
                    help="흰색이면 색을 넣지 않습니다.",
                )
                g1, g2 = st.columns(2)
                g1.download_button(
                    "표 전체 엑셀로 (목차 + Table 시트)",
                    data=write_tables_xlsx(computed_form,
                                           split_sheets=form_split,
                                           title_fill=form_fill),
                    file_name=f"{SAV_STEM}_테이블.xlsx",
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
                    file_name=f"{safe_stem(form_file.name)}_테이블설정.json",
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

        # ── 빼는 변수 ──
        # 주관식 문자 변수는 응답자마다 값이 달라서 빈도표가 사실상 원자료
        # 나열이 된다 ('시설명' 400명 → 보기 400개). 데이터를 훑는 목적에는
        # 방해가 되므로 기본으로 뺀다. 주관식을 정말 세고 싶으면 끄면 된다.
        text_vars = [c for c in df.columns
                     if not pd.api.types.is_numeric_dtype(df[c])]
        # 응답이 아예 없는 변수 — '기타 open' 처럼 아무도 안 적은 칸.
        # 다만 응답 0 자체가 확인거리일 수 있어서(로직·쿼터 점검) 기본은 켜 둔다.
        empty_vars = [c for c in df.columns if int(df[c].notna().sum()) == 0]

        st.session_state.setdefault("bt_freq_drop_text", True)
        st.session_state.setdefault("bt_freq_drop_empty", False)

        def kept(names: list[str]) -> list[str]:
            """'빼기' 체크에 걸리는 변수를 걸러 낸다."""
            out = list(names)
            if st.session_state.get("bt_freq_drop_text"):
                out = [c for c in out if c not in set(text_vars)]
            if st.session_state.get("bt_freq_drop_empty"):
                out = [c for c in out if c not in set(empty_vars)]
            return out

        def pick(names: list[str]) -> None:
            set_freq_vars(kept(names))

        p1, p2, p3, p4 = st.columns(4)
        p1.button("전체", key="bt_freq_all", on_click=pick,
                  args=(list(df.columns),))
        p2.button(f"값 라벨 있는 것만 ({len(labelled)})", key="bt_freq_lab",
                  on_click=pick, args=(labelled,))
        p3.button(f"숫자 변수만 ({len(numeric_only)})", key="bt_freq_num",
                  on_click=pick, args=(numeric_only,))
        p4.button("비우기", key="bt_freq_clear", on_click=set_freq_vars, args=([],))

        # 체크박스를 버튼 아래·고르는 칸 위에 둔다. 버튼이 이 설정을 따르므로
        # 순서가 그렇게 읽혀야 한다.
        # 아래 옵션 줄들과 같은 3칸을 써서 체크박스가 세로로 줄을 맞춘다
        e1, e2, _e3 = st.columns(3)
        e1.checkbox(
            f"문자 변수 빼기 ({len(text_vars)}개)", key="bt_freq_drop_text",
            disabled=not text_vars,
            help="주관식처럼 값이 응답자마다 다른 문자 변수는 빈도표가 원자료 "
                 "나열이 됩니다. 끄면 문자 변수도 넣되 많이 나온 값만 냅니다.",
        )
        e2.checkbox(
            f"응답 없는 변수 빼기 ({len(empty_vars)}개)",
            key="bt_freq_drop_empty", disabled=not empty_vars,
            help="아무도 답하지 않은 변수('기타 open' 등)입니다. 응답이 0인 "
                 "것 자체가 확인거리일 수 있어 기본은 넣어 둡니다.",
        )

        freq_disp = st.multiselect("빈도표를 뽑을 변수", DISPLAY_NAMES,
                                   key="bt_freq_vars")
        # 손으로 고른 변수도 같은 규칙으로 걸러 내고, 뺀 것은 밝혀 준다.
        # 조용히 빼면 "왜 이 표가 없지" 를 데이터에서 찾게 된다.
        freq_vars = kept(to_vars(freq_disp))
        dropped = [c for c in to_vars(freq_disp) if c not in set(freq_vars)]
        if dropped:
            shown = ", ".join(dropped[:8])
            more = f" 외 {len(dropped) - 8}개" if len(dropped) > 8 else ""
            st.caption(f"빼기 설정으로 {len(dropped)}개 제외 — {shown}{more}")

        q1, q2, q3 = st.columns(3)
        freq_missing = q1.checkbox("무응답(결측) 행 표시", value=True,
                                   key="bt_freq_missing")
        freq_sort = q2.checkbox("응답 많은 보기부터", key="bt_freq_sort",
                                help="'기타'·'모름'·'무응답' 계열은 맨 뒤로 보냅니다.")
        freq_split = q3.checkbox("변수마다 시트 나누기", key="bt_freq_split")

        r1, r2 = st.columns(2)
        freq_group = r1.checkbox(
            "복수응답 문항은 묶어서 한 표로", value=True, key="bt_freq_group",
            help="'X_1','X_2' 처럼 짝지어진 다중응답 묶음을 표 하나로 합칩니다. "
                 "이름만 보고 묶지 않고, 변수마다 자기 코드값 하나만 갖는지까지 "
                 "확인합니다 (5점 척도 배터리는 안 묶입니다).",
        )
        freq_all_values = r2.checkbox(
            "값이 많아도 전부 나열", value=True, key="bt_freq_all_values",
            help="응답된 값을 하나도 빼지 않고 냅니다. 끄면 고유값이 30개를 "
                 "넘는 변수는 줄이고, 숫자 변수는 통계 요약으로 갈음합니다.",
        )

        if not freq_vars:
            st.info("위에서 변수를 고르거나 '전체' 같은 버튼을 눌러 주세요.")
        else:
            if freq_all_values:
                # 값이 몇 백 종인 변수를 여러 개 고르면 표가 아주 길어진다.
                # 막지는 않고 몇 줄이 될지 미리 알려만 준다.
                est = sum(int(df[v].nunique()) for v in freq_vars)
                if est > 2000:
                    st.warning(
                        f"고른 변수들의 고유값을 합치면 약 {est:,}줄이 됩니다. "
                        "화면이 아주 길어지니 엑셀로 받아 보시는 편이 낫습니다."
                    )

            freq_tables = compute_frequencies(
                df, meta, freq_vars,
                show_missing=freq_missing, sort_by_count=freq_sort,
                text_limit=0 if freq_all_values else 30,
                group_multi=freq_group,
            )
            grouped = [t for t in freq_tables if t.table_kind == "multi"]
            if grouped:
                st.caption(
                    "복수응답으로 묶은 문항 — "
                    + " · ".join(f"{t.label}({len(t.members)}개)" for t in grouped)
                )
            flagged = [t for t in freq_tables
                       if any("값 라벨에 없는 코드" in n for n in t.notes)]
            if flagged:
                st.warning(
                    "값 라벨에 없는 코드가 있는 변수 "
                    f"{len(flagged)}개 — {', '.join(t.var for t in flagged)}. "
                    "코딩 오류이거나 라벨을 안 붙인 것이니 확인해 보세요."
                )

            c1, c1b, c2 = st.columns([1, 1, 3])
            freq_fill = c1.color_picker(
                "표 제목 줄 색", value=f"#{TITLE_FILL}", key="bt_freq_fill",
                help="표 제목 줄의 바탕색입니다. 표가 여러 개 이어 붙을 때 "
                     "구분선 역할을 합니다. 흰색으로 두면 색을 넣지 않습니다.",
            )
            freq_undef_fill = c1b.color_picker(
                "라벨 누락 줄 색", value=f"#{UNDEF_FILL}", key="bt_freq_undef_fill",
                help="값 라벨이 없는 코드 줄의 바탕색입니다. 표가 많으면 "
                     "'라벨없음' 이라는 글자만으로는 지나치기 쉬워서 줄 전체를 "
                     "칠합니다. 흰색으로 두면 칠하지 않습니다.",
            )
            with c2:
                st.download_button(
                    f"빈도표 {len(freq_tables)}개 엑셀로",
                    data=write_freq_xlsx(freq_tables, split_sheets=freq_split,
                                         title_fill=freq_fill,
                                         undef_fill=freq_undef_fill),
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
            y1, y2, y3, y4 = st.columns([2, 2, 2, 1])
            syn_sig_disp = y1.selectbox("유의성 검정", ["안 함", "95%", "99%"],
                                        key="bt_syn_sig")
            syn_min_base = y2.number_input("소표본 감추기 (사례수 미만)", 0, 500, 0,
                                           step=5, key="bt_syn_min_base")
            syn_split = y3.checkbox("표마다 시트를 나누기", key="bt_syn_split")
            syn_fill = y4.color_picker("제목 줄 색", value="#FFFFFF",
                                       key="bt_syn_fill",
                                       help="흰색이면 색을 넣지 않습니다.")

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
                    data=write_tables_xlsx(computed, split_sheets=syn_split,
                                           title_fill=syn_fill),
                    file_name=f"{safe_stem(sps_file.name)}_테이블.xlsx",
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
                    file_name=f"{safe_stem(sps_file.name)}_테이블설정.json",
                    mime="application/json",
                    key="bt_syn_cfg",
                )
                for res in computed:
                    st.markdown(f"**{res.title}**")
                    st.dataframe(result_to_frame(res), **_WIDE)
                    for note in res.notes:
                        st.caption(f"· {note}")


# =============================================================================
# 테이블 가이드로 신텍스 만들기 (.sav + 사내 테이블 가이드 엑셀)
# =============================================================================
with tab_guide:
    st.write(
        "사내 **테이블 가이드**(Basic Table · Banner 시트)와 위에서 올린 .sav 로 "
        "`3] Command.sps`(배너·리코드) 와 `4] Table.sps` 를 만듭니다. "
        "규칙으로 확신하지 못한 줄은 표시해 두니 **그 줄만 확인**하면 됩니다."
    )
    guide_file = st.file_uploader("테이블 가이드 (.xlsx)", type=["xlsx"], key="bt_guide_up")

    if guide_file is None:
        st.info("가이드 엑셀을 올려 주세요. 'Basic Table' 시트의 표 목록과 'Banner' 시트의 "
                "배너를 읽습니다. 'TG' 시트가 있으면 격자형 문항의 베이스를 푸는 데 씁니다.")
    else:
        @st.cache_data(show_spinner="가이드를 읽고 표를 맞추는 중…")
        def _guide_infer(guide_bytes: bytes, sav_bytes: bytes):
            g = tg.read_guide(guide_bytes)
            specs = tg.infer_tables(g, df, meta)
            return g, specs, tg.infer_banners(g, specs, df, meta)

        try:
            guide, specs0, banners0 = _guide_infer(guide_file.getvalue(), sav_file.getvalue())
        except ValueError as e:
            st.error(str(e))
            guide = None

        if guide is not None:
            _sig = hashlib.md5(guide_file.getvalue() + sav_file.getvalue()).hexdigest()[:12]

            # ── 양식 ──
            st.subheader("1. 양식")
            orient_g = st.radio(
                "배너 위치", ["배너를 행으로 (왼쪽에 배너)", "배너를 열로 (위쪽에 배너)"],
                horizontal=True, key="bt_guide_orient",
                help="배너를 행으로: 왼쪽에 성별·연령 등 배너가 세로로 오고 위쪽에 보기가 "
                     "옵니다 (`/table=@t3+ bv1+ … by t2+ …`).\n\n"
                     "배너를 열로: 위쪽에 배너가 가로로 오고 왼쪽에 보기가 옵니다 "
                     "(`/table=t3+ … by @t2+ bv1+ …`).",
            )
            orientation_g = tg.ROW if orient_g.startswith("배너를 행") else tg.COL
            proj = re.sub(r"_DATA.*$", "", Path(sav_file.name).stem, flags=re.I).strip()
            kms = guide.info.get("kms", "")
            p1, p2, p3 = st.columns([2, 1.3, 1.3])
            work_dir = p1.text_input(
                "작업 폴더 (CD)",
                value=f"D:\\{kms[:4]}\\({kms}) {proj}" if kms[:4].isdigit() else f"D:\\{proj}",
                key=f"bt_guide_cd_{_sig}")
            data_file = p2.text_input("원본 데이터 파일", value=sav_file.name,
                                      key=f"bt_guide_data_{_sig}")
            final_file = p3.text_input("Command 결과 파일", value=f"{proj}_Final.sav",
                                       key=f"bt_guide_final_{_sig}")

            # ── 배너 ──
            st.subheader("2. 배너")
            main0 = [b for b in banners0 if not b.extra]
            bdf = pd.DataFrame({
                "배너": [b.name for b in main0],
                "상태": [tg.STATUS_LABEL[b.status] for b in main0],
                "이름": [b.label for b in main0],
                "원본 변수": [b.source for b in main0],
                "리코드": [b.recode for b in main0],
                "보기": [" / ".join(f"{c}) {l}" for c, l in b.values)[:80] for b in main0],
                "사유": [b.reason for b in main0],
            })
            bed = st.data_editor(
                bdf, key=f"bt_guide_ban_{_sig}", hide_index=True, **_WIDE,
                disabled=["배너", "상태", "보기", "사유"],
                column_config={
                    "리코드": st.column_config.TextColumn(
                        help="비우면 COMPUTE 로 그대로 씁니다. 예: (1 2=1)(3=2)(4=3)"),
                    "원본 변수": st.column_config.TextColumn(help="배너로 쓸 .sav 변수 이름"),
                },
            )
            cols_up = {c.upper(): c for c in df.columns}
            banners_g = []
            for b, (_, r) in zip(main0, bed.iterrows()):
                src = cols_up.get(str(r["원본 변수"]).strip().upper(), str(r["원본 변수"]).strip())
                edited = (src != b.source or str(r["리코드"]).strip() != b.recode
                          or str(r["이름"]).strip() != b.label)
                banners_g.append(tg.BannerSpec(
                    b.name, str(r["이름"]).strip() or b.label, src,
                    str(r["리코드"] or "").strip(), b.values,
                    tg.OK if edited and src in df.columns else b.status, b.reason))
            bad_src = [b.name for b in banners_g if b.source not in df.columns]
            if bad_src:
                st.warning(f"원본 변수가 .sav 에 없는 배너: {', '.join(bad_src)}")

            # ── 표 목록 ──
            st.subheader("3. 표 목록")
            cnt = collections.Counter(s.status for s in specs0)
            st.caption(
                f"가이드 {len(specs0)}줄 — {tg.STATUS_LABEL[tg.OK]} {cnt[tg.OK]} · "
                f"{tg.STATUS_LABEL[tg.CHECK]} {cnt[tg.CHECK]} · "
                f"{tg.STATUS_LABEL[tg.NEED]} {cnt[tg.NEED]}. "
                "확인이 필요한 줄이 위로 오게 정렬했습니다. 신텍스는 가이드 순서대로 나갑니다.")
            order = sorted(range(len(specs0)),
                           key=lambda i: ({tg.NEED: 0, tg.CHECK: 1, tg.OK: 2}[specs0[i].status], i))
            tdf = pd.DataFrame({
                "#": [i + 1 for i in order],
                "상태": [tg.STATUS_LABEL[specs0[i].status] for i in order],
                "제목": [specs0[i].title for i in order],
                "유형": [tg.KIND_LABEL[specs0[i].kind] for i in order],
                "변수": [specs0[i].vars for i in order],
                "조건": [specs0[i].cond for i in order],
                "정렬": [specs0[i].sort for i in order],
                "추가배너": [specs0[i].extra_banner for i in order],
                "통계값": [specs0[i].stats for i in order],
                "가이드 베이스": [specs0[i].base for i in order],
                "사유": [specs0[i].reason for i in order],
            })
            ted = st.data_editor(
                tdf, key=f"bt_guide_tab_{_sig}", hide_index=True, **_WIDE,
                disabled=["#", "상태", "가이드 베이스", "사유"],
                column_config={
                    "#": st.column_config.NumberColumn(width="small"),
                    "상태": st.column_config.TextColumn(width="small"),
                    "제목": st.column_config.TextColumn(width="medium"),
                    "유형": st.column_config.SelectboxColumn(
                        options=[tg.KIND_LABEL[k] for k in tg.KINDS], required=True,
                        width="small"),
                    "변수": st.column_config.TextColumn(
                        help="'q1' · 'q16_1 to q16_3' · 'b2_1 b2_2 …' (Summary 는 구성 변수)"),
                    "조건": st.column_config.TextColumn(
                        help="Select if 에 & 로 붙는 SPSS 조건. 비우면 응답자 전체. "
                             "예: A5=1 · Range(A6,2,4) · any(D1,2,3)"),
                    "정렬": st.column_config.CheckboxColumn(help="내림차순 (/sort= m_down)"),
                    "추가배너": st.column_config.TextColumn(help="이 표에만 더 붙일 배너 변수"),
                    "통계값": st.column_config.TextColumn(
                        help="척도 묶음과 100점 환산. 예: BOT2/SoSo/TOP2,Mean, 100점환산 평균"),
                },
            )
            edited_by_no = {int(r["#"]): r for _, r in ted.iterrows()}
            specs_g, var_problems = [], []
            for i, s0 in enumerate(specs0):
                r = edited_by_no[i + 1]
                vars_txt = str(r["변수"] or "").strip()
                _vs, bad = tg.parse_vars(vars_txt, tg.Data(df, meta))
                if bad:
                    var_problems.append(f"#{i + 1} {s0.title[:30]} — 없는 변수: {', '.join(bad)}")
                changed = (vars_txt != s0.vars or str(r["조건"] or "").strip() != s0.cond
                           or tg.LABEL_KIND.get(r["유형"], s0.kind) != s0.kind)
                specs_g.append(tg.TableSpec(
                    title=str(r["제목"]).strip() or s0.title,
                    kind=tg.LABEL_KIND.get(r["유형"], s0.kind),
                    vars=vars_txt, cond=str(r["조건"] or "").strip(),
                    sort=bool(r["정렬"]), extra_banner=str(r["추가배너"] or "").strip(),
                    stats=str(r["통계값"] or ""), base=s0.base,
                    status=(tg.OK if changed and vars_txt and not bad else s0.status),
                    reason=s0.reason, row=s0.row))
            for msg in var_problems[:8]:
                st.warning(msg)
            banners_g += tg.extra_banners(specs_g, df, meta, start=len(banners_g) + 1)

            # ── 내려받기 ──
            st.subheader("4. 내려받기")
            left = sum(1 for s in specs_g if s.status == tg.NEED)
            if left:
                st.warning(f"❌ 직접 입력이 남은 표 {left}개는 신텍스에 '* 확인 필요' 주석과 함께 "
                           "나갑니다. 변수가 비어 있으면 그 표는 건너뜁니다.")
            enc_disp = st.radio("인코딩", ["CP949 (SPSS 한글 윈도우)", "UTF-8"],
                                horizontal=True, key="bt_guide_enc")
            enc = "cp949" if enc_disp.startswith("CP949") else "utf-8"
            try:
                cmd_txt = tg.write_command(banners_g, specs_g, df, meta, work_dir=work_dir,
                                           data_file=data_file, final_file=final_file)
                tab_txt = tg.write_table(specs_g, banners_g, df, meta,
                                         orientation=orientation_g, work_dir=work_dir,
                                         final_file=final_file,
                                         section_title=guide.section_title)
            except Exception as e:                           # noqa: BLE001
                st.error(f"신텍스를 만들지 못했습니다 — {e}")
                cmd_txt = tab_txt = None
            if cmd_txt is not None:
                cmd_b, bad1 = tg.encode_sps(cmd_txt, enc)
                tab_b, bad2 = tg.encode_sps(tab_txt, enc)
                if bad1 + bad2:
                    st.warning(f"CP949 에 없는 글자 {bad1 + bad2}개가 '?' 로 바뀝니다. "
                               "제목에 특수문자가 있으면 UTF-8 로 받으세요.")
                d1, d2 = st.columns(2)
                d1.download_button("3] Command.sps 내려받기", data=cmd_b,
                                   file_name=f"3] {proj} - Command.sps",
                                   mime="text/plain", key="bt_guide_dl_cmd")
                d2.download_button("4] Table.sps 내려받기", data=tab_b,
                                   file_name=f"4] {proj} - Table.sps",
                                   mime="text/plain", key="bt_guide_dl_tab")
                with st.expander("미리 보기"):
                    v1, v2 = st.tabs(["Command", "Table"])
                    v1.code("\n".join(cmd_txt.splitlines()[:120]), language="sql")
                    v2.code("\n".join(tab_txt.splitlines()[:160]), language="sql")

"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : 2___쿼터_솔루션.py                                               ║
║  위치   : pages/2___쿼터_솔루션.py   ← 반드시 pages/ 폴더 안!               ║
║                                                                          ║
║  이 파일은 화면(UI) 코드입니다. utils.py 에 붙여넣으면                      ║
║  NameError: name 'utils' is not defined 가 발생합니다.                     ║
║  utils.py 는 별도 파일이며 리포 최상단에 둡니다.                            ║
╚══════════════════════════════════════════════════════════════════════════╝

2___쿼터_솔루션.py — 쿼터 자동 할당 화면

주요 변경점
-----------
1. 로컬 normalize_val / clean_series 삭제 -> utils.norm_val / norm_series 로 통일
2. 키 생성을 설정 화면과 실행 시점이 같은 함수(utils.build_*_keys)로 공유
   -> 화면에 보이는 집계와 실제 매칭 대상이 어긋나던 문제 해소
   -> df_proc 사본 자체가 불필요해져 제거
3. @st.cache_data 로 리런마다 반복되던 전처리/집계 제거 (메모리 상한 포함)
4. 목표값 입력 오류를 조용히 삼키지 않고 경고로 표시
5. 실행 전 정합성 프리플라이트 : 데이터에 아예 없는 쿼터 셀을 미리 경고
6. 인덱스 int 강제 캐스팅 제거 (문자열 ID 인덱스에서 죽던 문제)
7. 시트명 충돌 방지, Main_Status index=False, 0 나누기 가드
8. 계산 방식 선택 : 정확해(ILP) / 휴리스틱(그리디)
"""

import streamlit as st
import pandas as pd
import io
import collections
import numpy as np
import altair as alt
from joblib import Parallel, delayed, cpu_count
import sys
import os
import traceback

# 리포 최상단(부모 디렉터리)과 현재 디렉터리를 모두 경로에 넣는다.
# 이 파일이 pages/ 안에 있어도, 실수로 최상단에 있어도 utils 를 찾는다.
_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (os.path.dirname(_HERE), _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# --- utils 임포트 가드 -------------------------------------------------------
# 파일 내용이 뒤섞였을 때 원시 트레이스백 대신 무엇을 고쳐야 하는지 알려준다.
try:
    import utils
    if getattr(utils, "MODULE_ROLE", None) != "utils":
        raise ImportError("utils.py 의 내용이 공용 모듈이 아닙니다.")
except Exception:
    st.error(
        "### ❌ utils.py 를 불러오지 못했습니다\n\n"
        "**파일 내용이 서로 바뀐 경우가 대부분입니다.** 아래를 확인하세요.\n\n"
        "| 파일 | 있어야 하는 것 | 없어야 하는 것 |\n"
        "|---|---|---|\n"
        "| `utils.py` (최상단) | `def norm_val`, `def check_password` | `st.set_page_config` |\n"
        "| `pages/2___쿼터_솔루션.py` | `st.file_uploader`, `import utils` | `def norm_val` |\n\n"
        "`utils.py` 안에 `import utils` 나 `st.set_page_config` 가 보이면 "
        "그 파일에 화면 코드가 잘못 들어간 것입니다."
    )
    st.code(traceback.format_exc())
    st.stop()

# ILP 솔버는 선택적 의존성 (pip install ortools)
try:
    import quota_ilp
    if getattr(quota_ilp, "MODULE_ROLE", None) != "quota_ilp":
        raise ImportError("quota_ilp.py 의 내용이 올바르지 않습니다.")
    HAS_ILP, ILP_ERR = True, None
except Exception as _e:
    quota_ilp, HAS_ILP, ILP_ERR = None, False, str(_e)

st.set_page_config(page_title="쿼터 솔루션", layout="wide")

if not utils.check_password():
    st.stop()

st.title("📊 쿼터 자동 할당 솔루션")
n_cores = cpu_count()
st.sidebar.caption(f"🖥️ CPU 코어: {n_cores}개")

MAX_GRID_CELLS = 20000       # 교차표 폭발 방지 임계치
CACHE_MAX_ENTRIES = 6        # 캐시 항목 상한 (배포 환경 메모리 보호)
CACHE_TTL = 3600             # 캐시 유효 시간(초)


# ==============================================================================
# 캐시 래퍼 : 설정 화면과 실행 시점이 같은 결과를 공유하도록 보장한다
# ==============================================================================
@st.cache_data(show_spinner=False, max_entries=CACHE_MAX_ENTRIES, ttl=CACHE_TTL)
def cached_simple_keys(df, cols):
    return utils.build_simple_keys(df, list(cols))


@st.cache_data(show_spinner=False, max_entries=CACHE_MAX_ENTRIES, ttl=CACHE_TTL)
def cached_tuple_keys(df, cols):
    return utils.build_tuple_keys(df, list(cols))


@st.cache_data(show_spinner=False, max_entries=CACHE_MAX_ENTRIES, ttl=CACHE_TTL)
def cached_grid_keys(df, cols):
    return utils.build_grid_keys(df, list(cols))


@st.cache_data(show_spinner=False, max_entries=CACHE_MAX_ENTRIES, ttl=CACHE_TTL)
def cached_pivot(df, row_cols, col_name):
    """교차표(행 변수 × 열 변수) 집계. 값은 전부 norm_val 로 정규화된 상태."""
    cols = list(row_cols) + [col_name]
    base = pd.DataFrame({c: utils.norm_series(df[c]) for c in cols})
    for c in cols:
        uv = sorted(base[c].unique(), key=utils.natural_key)
        base[c] = pd.Categorical(base[c], categories=uv, ordered=True)
    return base.groupby(cols, observed=False).size().unstack(fill_value=0)


def parse_target(v):
    """목표값 파싱. 유효하지 않으면 None."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if pd.isna(f) or f < 0 or f != int(f):
        return None
    return int(f)


def warn_bad(bad_labels, where):
    if bad_labels:
        head = ", ".join(str(b) for b in bad_labels[:5])
        more = f" 외 {len(bad_labels) - 5}건" if len(bad_labels) > 5 else ""
        st.warning(
            f"⚠️ {where}: 목표값이 올바르지 않아 **{len(bad_labels)}개 항목을 건너뛰었습니다** "
            f"({head}{more}). 0 이상의 정수를 입력하세요."
        )


# ==============================================================================
# 1. 데이터 업로드
# ==============================================================================
st.subheader("1. 데이터 업로드")
data_file = st.file_uploader("설문 데이터", type=['csv', 'xlsx'], key="quota_up")

if data_file:
    df_survey = utils.load_df(data_file)

    # [수정] load_df 는 실패 시 None 을 반환한다. 곧바로 len() 하면 TypeError.
    if df_survey is None:
        st.stop()
    if df_survey.empty:
        st.error("데이터가 비어 있습니다.")
        st.stop()
    if not df_survey.index.is_unique:
        st.warning("인덱스에 중복이 있어 0부터 다시 매깁니다.")
        df_survey = df_survey.reset_index(drop=True)

    st.success(f"로드 완료: {len(df_survey)}명")
    st.divider()

    # ==========================================================================
    # 2. 쿼터 설정
    # ==========================================================================
    st.subheader("2. 쿼터 설정")
    use_main = st.checkbox("✅ 메인 쿼터 사용", value=True)
    main_map = {}
    algo_main_cols = []
    main_mode = 'grid'

    if use_main:
        q_mode = st.radio("메인 쿼터 방식", ["엑셀 업로드", "화면 설계"], horizontal=True)

        if q_mode == "엑셀 업로드":
            qf = st.file_uploader("쿼터 파일", type=['xlsx'])
            c1, c2, c3 = st.columns(3)
            with c1: q1 = st.selectbox("qt1", df_survey.columns)
            with c2: q2 = st.selectbox("qt2", df_survey.columns)
            with c3: q3 = st.selectbox("qt3", df_survey.columns)
            if qf:
                algo_main_cols = [q1, q2, q3]
                try:
                    raw = pd.read_excel(qf, 0, header=None)
                    flat = utils.transform_pivoted_quota(raw)
                    # 키는 utils.norm_val 로 이미 정규화되어 있다
                    main_map = {
                        (r.qt1, r.qt2, r.qt3): int(r.target)
                        for r in flat.itertuples()
                    }
                    st.caption(f"쿼터 셀 {len(main_map)}개 / 목표 합계 {sum(main_map.values()):,}명")
                except Exception as e:
                    # [수정] bare except 제거. 원인을 그대로 보여준다.
                    st.error(f"쿼터 엑셀 파싱 실패 — {type(e).__name__}: {e}")
                    with st.expander("상세 오류"):
                        st.code(traceback.format_exc())

        else:
            rv = st.multiselect("행(Row) 변수", df_survey.columns)
            cv = st.selectbox("열(Col) 변수", ["(선택)"] + list(df_survey.columns))
            if rv and cv != "(선택)":
                if cv in rv:
                    st.error("열 변수는 행 변수와 달라야 합니다.")
                else:
                    algo_main_cols = rv + [cv]
                    pi = cached_pivot(df_survey, tuple(rv), cv)
                    if pi.size > MAX_GRID_CELLS:
                        st.error(f"교차표 셀이 {pi.size:,}개로 너무 많습니다. 변수를 줄이세요.")
                    else:
                        ed = st.data_editor(pi.reset_index(), use_container_width=True, disabled=rv)
                        mlt = ed.melt(id_vars=rv, var_name=cv, value_name='target')
                        bad = []
                        for _, r in mlt.iterrows():
                            key = tuple(utils.norm_val(r[c]) for c in algo_main_cols)
                            t = parse_target(r['target'])
                            if t is None:
                                bad.append(" / ".join(key))
                                continue
                            if t > 0:
                                main_map[key] = t
                        warn_bad(bad, "메인 쿼터")
                        st.caption(f"쿼터 셀 {len(main_map)}개 / 목표 합계 {sum(main_map.values()):,}명")
    else:
        main_map = {('All',): st.number_input("전체 목표", 1, 1000000, 1000)}
        algo_main_cols = []

    # --------------------------------------------------------------------------
    # 추가 쿼터
    # --------------------------------------------------------------------------
    ex_configs = []
    tabs = st.tabs(["추가 1", "추가 2", "추가 3", "추가 4"])

    for i, tab in enumerate(tabs):
        with tab:
            ex_mode = st.radio(
                f"설정 방식 (그룹 {i+1})",
                ["단순형 (변수 값별 할당)", "조합형 (행/열 교차 할당)"],
                key=f"ex_mode_{i}", horizontal=True
            )
            config = {'cols': [], 'map': {}, 'name': f"Extra_{i+1}", 'mode': 'simple'}

            if ex_mode.startswith("단순형"):
                config['mode'] = 'simple'
                cols = st.multiselect(f"변수 선택 (그룹 {i+1})", df_survey.columns, key=f"ms{i}")
                if cols:
                    config['cols'] = cols
                    config['name'] = "_".join(str(c) for c in cols)

                    # [핵심 수정] 실행 시점과 완전히 동일한 함수로 키를 만든다.
                    # 예전 코드는 여기서 collect_values_from_cols(중복제거·결측제외)를 쓰고
                    # 실행 시점엔 [str(r[c]) for c in cols] 를 써서 결과가 어긋났다.
                    keys_setup = cached_simple_keys(df_survey, tuple(cols))
                    counter = collections.Counter(v for ks in keys_setup for v in ks)

                    if not counter:
                        st.info("유효한 값이 없습니다 (전부 결측).")
                    else:
                        cnt = pd.DataFrame(
                            sorted(counter.items(), key=lambda kv: utils.natural_key(kv[0])),
                            columns=['값', '현재']
                        )
                        cnt['목표'] = cnt['현재']
                        ed = st.data_editor(cnt, use_container_width=True,
                                            disabled=['값', '현재'], key=f"ed{i}", hide_index=True)
                        bad = []
                        for _, r in ed.iterrows():
                            t = parse_target(r['목표'])
                            if t is None:
                                bad.append(r['값'])
                                continue
                            if t > 0:
                                config['map'][str(r['값'])] = t
                        warn_bad(bad, f"추가 쿼터 {i+1}")

            else:
                config['mode'] = 'grid'
                st.caption("행과 열을 교차하여 상세 목표를 설정합니다.")
                ex_rv = st.multiselect(f"행(Row) 변수 (그룹 {i+1})", df_survey.columns, key=f"ex_rv_{i}")
                ex_cv = st.selectbox(f"열(Col) 변수 (그룹 {i+1})",
                                     ["(선택)"] + list(df_survey.columns), key=f"ex_cv_{i}")

                if ex_rv and ex_cv != "(선택)":
                    if ex_cv in ex_rv:
                        st.error("열 변수는 행 변수와 달라야 합니다.")
                    else:
                        target_cols = ex_rv + [ex_cv]
                        config['cols'] = target_cols
                        config['name'] = "_".join(str(c) for c in target_cols)

                        pi = cached_pivot(df_survey, tuple(ex_rv), ex_cv)
                        if pi.size > MAX_GRID_CELLS:
                            st.error(f"교차표 셀이 {pi.size:,}개로 너무 많습니다.")
                        else:
                            ed = st.data_editor(pi.reset_index(), use_container_width=True,
                                                disabled=ex_rv, key=f"ex_ed_grid_{i}")
                            mlt = ed.melt(id_vars=ex_rv, var_name=ex_cv, value_name='target')
                            bad = []
                            for _, r in mlt.iterrows():
                                key = tuple(utils.norm_val(r[c]) for c in target_cols)
                                t = parse_target(r['target'])
                                if t is None:
                                    bad.append(" / ".join(key))
                                    continue
                                if t > 0:
                                    config['map'][key] = t
                            warn_bad(bad, f"추가 쿼터 {i+1}")

            ex_configs.append(config)

    # ==========================================================================
    # 3. 실행 옵션
    # ==========================================================================
    st.divider()
    st.subheader("3. 실행 옵션")
    solver_opts = (["정확해 (ILP)", "휴리스틱 (그리디)"] if HAS_ILP
                   else ["휴리스틱 (그리디)"])
    solver_kind = st.radio(
        "계산 방식", solver_opts, horizontal=True,
        help=("ILP: 최적해임을 증명하고, 미달 시 어느 상한이 막는지 정확히 알려줍니다. "
              "그리디: 랜덤 재시작 휴리스틱으로 최적성 보장이 없습니다.")
    )
    if not HAS_ILP:
        st.caption(f"ℹ️ ILP 사용 불가 (`pip install ortools`) — {ILP_ERR}")
    use_ilp = solver_kind.startswith("정확해")

    c1, c2 = st.columns(2)
    with c1:
        c_no = st.selectbox("ID 컬럼", df_survey.columns)
        tol = st.number_input("허용 오차", 0, 100, 0)
        use_intval = st.checkbox(
            "intval 최적화", value=True,
            help="쿼터 조건이 완전히 같은 응답자들 사이에서, intval 값이 낮은 쪽을 "
                 "먼저 탈락시킵니다. 조건이 다른 응답자끼리는 영향이 없으므로 "
                 "최종 통과 인원수는 달라지지 않습니다.")
        c_int = st.selectbox("intval 컬럼", df_survey.columns) if use_intval else None
        if not use_intval:
            rand_pick = st.checkbox(
                "동일 조건 응답자 무작위 선택", value=True,
                help="끄면 데이터 순서대로 뽑아 결과가 완전히 재현됩니다.")
        else:
            rand_pick = False
    with c2:
        if use_ilp:
            time_limit = st.number_input("시간 제한(초)", 5, 600, 60, 5)
            balance = st.checkbox(
                "부족분 고르게 분산", value=True,
                help="총 부족 인원을 최소화한 뒤, 특정 셀에 부족이 몰리지 않도록 재조정합니다.")
            iters, backend, jitter = 0, None, 0.0
        else:
            iters = st.number_input("시도 횟수", 100, 1000000, 10000, 1000)
            jitter = st.slider("탐색 폭 (지터)", 0.0, 0.5, 0.15, 0.05,
                               help="0이면 항상 같은 해만 나옵니다.")
            backend = st.selectbox(
                "병렬 방식", ["프로세스 (loky)", "스레드 (threading)"],
                help=("워커가 파이썬 루프 위주라 스레드는 GIL 때문에 거의 빨라지지 "
                      "않습니다. 데이터가 매우 크면 직렬화 비용 때문에 스레드가 "
                      "나을 수도 있습니다.")
            )
            time_limit, balance = 0, False

    if st.button("🚀 매칭 시작", type="primary"):
        if not main_map:
            st.error("목표가 설정되지 않았습니다.")
            st.stop()
        if use_main and not algo_main_cols:
            st.error("메인 쿼터 변수를 선택하세요.")
            st.stop()

        try:
            with st.spinner("희소성 계산 및 병렬 연산 중..."):
                # ------------------------------------------------------------------
                # 키 생성 : 설정 화면과 동일한 캐시 함수를 호출한다 (df_proc 불필요)
                # ------------------------------------------------------------------
                if use_main:
                    m_keys = cached_tuple_keys(df_survey, tuple(algo_main_cols))
                else:
                    m_keys = [('All',)] * len(df_survey)

                ex_keys_list = []
                for cfg in ex_configs:
                    if not cfg['cols']:
                        ex_keys_list.append([[] for _ in range(len(df_survey))])
                    elif cfg['mode'] == 'simple':
                        ex_keys_list.append(cached_simple_keys(df_survey, tuple(cfg['cols'])))
                    else:
                        ex_keys_list.append(cached_grid_keys(df_survey, tuple(cfg['cols'])))

                target_total = sum(main_map.values())
                soft_target = max(0, target_total - tol)
                m_cnt = collections.Counter(m_keys)

                # ------------------------------------------------------------------
                # 프리플라이트 : 데이터에 아예 존재하지 않는 쿼터 키 경고
                # (정규화 불일치를 실행 전에 잡아내는 안전망)
                # ------------------------------------------------------------------
                ghosts = [k for k in main_map if m_cnt.get(k, 0) == 0]
                if ghosts:
                    st.warning(
                        f"⚠️ 메인 쿼터 {len(ghosts)}개 셀이 데이터에 한 명도 없습니다. "
                        f"목표 {sum(main_map[k] for k in ghosts):,}명은 달성 불가입니다.\n\n"
                        + ", ".join(" / ".join(k) for k in ghosts[:10])
                        + (" ..." if len(ghosts) > 10 else "")
                    )

                ex_maps = [c['map'] for c in ex_configs]
                indices = df_survey.index.to_numpy()
                pick_rng = np.random.default_rng(0) if rand_pick else None
                ilp_sol = None

                # --- intval 타이브레이크 ---
                tiebreak = None
                if use_intval and c_int:
                    tiebreak, n_ok, n_bad = utils.build_tiebreak(df_survey, c_int)
                    if n_ok == 0:
                        st.error(
                            f"`{c_int}` 컬럼에서 숫자를 하나도 읽지 못했습니다. "
                            "intval 최적화를 끄거나 숫자 컬럼을 선택하세요.")
                        st.stop()
                    if n_bad:
                        st.warning(
                            f"⚠️ `{c_int}` 컬럼에 숫자가 아닌 값/결측이 {n_bad:,}건 "
                            "있습니다. 해당 응답자는 **가장 먼저 탈락**합니다.")

            # ======================================================================
            # 실행 (A) 정확해 : 정수계획법
            # ======================================================================
            if use_ilp:
                with st.spinner("정수계획법으로 최적해 탐색 중..."):
                    ilp_sol = quota_ilp.solve_quota_ilp(
                        m_keys, ex_keys_list, main_map, ex_maps, indices,
                        balance=balance, time_limit=time_limit,
                        workers=n_cores, rng=pick_rng, tiebreak=tiebreak)
                g_best_cnt, g_best_idxs = ilp_sol.total, ilp_sol.selected

            # ======================================================================
            # 실행 (B) 휴리스틱 : 랜덤 재시작 그리디
            # ======================================================================
            else:
              with st.spinner("희소성 계산 및 병렬 연산 중..."):
                # 희소성 점수 : 보유/목표 비율이 낮을수록 먼저 뽑는다
                if use_main:
                    score_main = np.array([
                        m_cnt.get(k, 0) / main_map[k] if main_map.get(k, 0) > 0
                        else utils.MISS_PENALTY
                        for k in m_keys
                    ], dtype=float)
                else:
                    score_main = np.ones(len(df_survey), dtype=float)

                score_extras = np.zeros(len(df_survey), dtype=float)
                n_active_ex = sum(1 for c in ex_configs if c['cols'])
                for j, cfg in enumerate(ex_configs):
                    if not cfg['cols']:
                        continue
                    ex_cnt_total = collections.Counter(
                        v for keys in ex_keys_list[j] for v in keys
                    )
                    ex_map = cfg['map']
                    row_scores = np.empty(len(df_survey), dtype=float)
                    for ridx, keys in enumerate(ex_keys_list[j]):
                        if not keys:
                            row_scores[ridx] = 1.0
                            continue
                        best = utils.MISS_PENALTY
                        for k in keys:
                            cap = ex_map.get(k, 0)
                            s = ex_cnt_total[k] / cap if cap > 0 else utils.MISS_PENALTY
                            if s < best:
                                best = s
                        row_scores[ridx] = best
                    score_extras += row_scores

                # [수정] 추가 그룹 수만큼 점수가 커져 메인 쿼터 영향력이 희석되던 문제.
                # 그룹 평균을 써서 메인:추가 = 1:1 스케일로 맞춘다.
                if n_active_ex:
                    score_extras /= n_active_ex
                final_scarcity_scores = score_main + score_extras

                # ------------------------------------------------------------------
                # 병렬 실행
                # ------------------------------------------------------------------
                jl_backend = "loky" if backend.startswith("프로세스") else "threading"
                ipc = max(1, -(-int(iters) // n_cores))    # 올림 분배
                indices = df_survey.index.to_numpy()

                res = Parallel(n_jobs=n_cores, backend=jl_backend)(
                    delayed(utils.simulation_worker)(
                        seed, ipc, indices, final_scarcity_scores, m_keys, ex_keys_list,
                        main_map, [c['map'] for c in ex_configs],
                        soft_target, target_total, jitter, tiebreak
                    ) for seed in range(n_cores)
                )

                g_best_cnt, g_best_idxs = 0, []
                for c, ixs in res:
                    if c > g_best_cnt:
                        g_best_cnt, g_best_idxs = c, ixs

            is_fail = g_best_cnt < soft_target

            # ==================================================================
            # 결과 집계
            # ==================================================================
            # [수정] int() 강제 캐스팅 제거. indices 는 원본 인덱스 라벨 그대로다.
            fin_idxs = list(g_best_idxs)
            pos_of = {lbl: p for p, lbl in enumerate(df_survey.index)}

            final_m = collections.Counter()
            final_exs = [collections.Counter() for _ in ex_configs]
            for lbl in fin_idxs:
                p = pos_of[lbl]
                final_m[m_keys[p]] += 1
                for j, cfg in enumerate(ex_configs):
                    if cfg['cols']:
                        for k in ex_keys_list[j][p]:
                            final_exs[j][k] += 1

            # ------------------------------------------------------------------
            # 부족분 진단
            # ------------------------------------------------------------------
            recs = []
            if is_fail:
                if use_main:
                    for k, tgt in main_map.items():
                        act = final_m.get(k, 0)
                        diff = tgt - act
                        if diff > 0:
                            raw_avail = m_cnt.get(k, 0)
                            reason = "⚠️ 물리적 부족" if raw_avail < tgt else "⚔️ 경합 부족"
                            recs.append({'순서': 0, '구분': '메인 쿼터', '항목': " / ".join(k),
                                         '목표': tgt, '현재': act, '부족': diff,
                                         '진단': reason, '전체보유': raw_avail})

                for j, cfg in enumerate(ex_configs):
                    if not cfg['cols']:
                        continue
                    raw_cnt_map = collections.Counter(
                        v for keys in ex_keys_list[j] for v in keys
                    )
                    for k, tgt in cfg['map'].items():
                        act = final_exs[j].get(k, 0)
                        diff = tgt - act
                        if diff > 0:
                            raw_avail = raw_cnt_map.get(k, 0)
                            reason = "⚠️ 물리적 부족" if raw_avail < tgt else "⚔️ 경합 부족"
                            label = " / ".join(k) if isinstance(k, tuple) else str(k)
                            recs.append({'순서': j + 1, '구분': cfg['name'], '항목': label,
                                         '목표': tgt, '현재': act, '부족': diff,
                                         '진단': reason, '전체보유': raw_avail})

            # ------------------------------------------------------------------
            # 엑셀 저장
            # ------------------------------------------------------------------
            df_out = df_survey.copy()
            df_out['Chk'] = "제외"
            df_out.loc[fin_idxs, 'Chk'] = "통과"

            df_all = df_out.sort_values(by=c_no, ascending=True)
            df_pass = df_out[df_out['Chk'] == "통과"].sort_values(c_no, ascending=True)

            out = io.BytesIO()
            used_sheets = set()
            sheet_names = {}
            with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                df_all.to_excel(w, index=False, sheet_name='Result_All')
                df_pass.to_excel(w, index=False, sheet_name='Result_Pass')

                if recs:
                    df_excel = pd.DataFrame(recs)
                    df_excel['sort_val'] = df_excel['항목'].map(lambda x: tuple(utils.natural_key(x)))
                    df_excel = df_excel.sort_values(by=['순서', 'sort_val'])
                    df_excel.drop(columns=['순서', 'sort_val']).to_excel(
                        w, index=False, sheet_name='Shortage_Analysis')

                if use_main:
                    pd.DataFrame([
                        {'Group': " / ".join(k), 'Target': v, 'Actual': final_m[k],
                         'Diff': v - final_m[k]}
                        for k, v in main_map.items()
                    ]).to_excel(w, index=False, sheet_name='Main_Status')

                for j, cfg in enumerate(ex_configs):
                    if not cfg['cols']:
                        continue
                    # [수정] 같은 컬럼 조합이면 시트명이 충돌해 xlsxwriter 가 죽었다
                    sname = utils.unique_sheet_name(cfg['name'], used_sheets)
                    sheet_names[j] = sname
                    data_e = [
                        {'Value': " / ".join(k) if isinstance(k, tuple) else str(k),
                         'Target': t, 'Actual': final_exs[j][k], 'Diff': t - final_exs[j][k]}
                        for k, t in cfg['map'].items()
                    ]
                    if data_e:
                        pd.DataFrame(data_e).sort_values(
                            'Value', key=lambda c: c.map(utils.natural_key)
                        ).to_excel(w, sheet_name=sname, index=False)

            # ==================================================================
            # 결과 표시
            # ==================================================================
            st.divider()
            st.subheader("📊 할당 결과")

            total_rows, pass_rows = len(df_out), len(df_pass)
            st.info(f"💾 총 **{total_rows:,}명** "
                    f"(통과 {pass_rows:,}명 + 제외 {total_rows - pass_rows:,}명) 저장 완료")

            st.download_button(
                "📥 결과 파일 다운로드" if not is_fail else "⚠️ 실패한 결과라도 다운로드",
                out.getvalue(), "result.xlsx", type="primary", use_container_width=True
            )

            rate = (g_best_cnt / target_total * 100) if target_total else 0.0   # 0 나누기 가드
            c1, c2, c3 = st.columns(3)
            c1.metric("📌 전체 목표", f"{target_total:,}명")
            c2.metric("✅ 매칭 성공", f"{g_best_cnt:,}명")
            c3.metric("📈 달성률", f"{rate:.1f}%",
                      delta=f"{g_best_cnt - target_total}명" if is_fail else "목표 달성",
                      delta_color="inverse" if is_fail else "normal")

            # ------------------------------------------------------------------
            # intval 적용 검증 : 통과자의 intval 이 실제로 더 높은지 확인
            # ------------------------------------------------------------------
            if tiebreak is not None and fin_idxs:
                tb_all = pd.Series(tiebreak, index=df_survey.index).replace(
                    [-np.inf, np.inf], np.nan)
                tb_pass = tb_all.loc[fin_idxs].dropna()
                tb_drop = tb_all.drop(index=fin_idxs).dropna()
                if len(tb_pass) and len(tb_drop):
                    i1, i2, i3 = st.columns(3)
                    i1.metric(f"통과자 {c_int} 평균", f"{tb_pass.mean():,.1f}")
                    i2.metric(f"탈락자 {c_int} 평균", f"{tb_drop.mean():,.1f}",
                              delta=f"{tb_drop.mean() - tb_pass.mean():,.1f}")
                    i3.metric(f"통과자 {c_int} 최소", f"{tb_pass.min():,.1f}")
                    st.caption(
                        f"쿼터 조건이 같은 응답자 중 `{c_int}` 값이 낮은 쪽을 먼저 "
                        "탈락시킨 결과입니다. 조건이 다른 응답자끼리는 비교하지 않으므로 "
                        "통과자 평균이 항상 더 높다고 보장되지는 않습니다 "
                        "(희소한 셀에서는 값이 낮아도 뽑아야 합니다)."
                    )

            if is_fail:
                st.error("⚠️ 목표 인원을 달성하지 못했습니다. 아래 분석을 확인하세요.")
            else:
                st.success("🎉 목표 인원을 모두 달성했습니다!")

            # ------------------------------------------------------------------
            # ILP 전용: 최적성 보증 + 병목 진단
            # ------------------------------------------------------------------
            if ilp_sol is not None:
                if ilp_sol.proven_optimal:
                    st.success(
                        f"✅ **최적해임이 증명되었습니다.** 이 조건에서 {ilp_sol.total:,}명보다 "
                        f"많이 뽑는 방법은 존재하지 않습니다. "
                        f"(프로파일 {ilp_sol.n_profiles:,}개로 집약 / {ilp_sol.solve_sec:.2f}초)"
                    )
                else:
                    st.warning(
                        f"⏱️ 시간 제한({time_limit}초) 내에 최적성을 증명하지 못했습니다 "
                        f"(상태: {ilp_sol.status}). 현재 해는 유효하지만 더 나은 해가 "
                        f"있을 수 있습니다. 시간 제한을 늘려보세요."
                    )

                d = ilp_sol.diagnosis
                if is_fail:
                    st.markdown("#### 🧭 왜 목표를 못 채웠는가")

                    if d.group_relax_gain:
                        rows = [{'추가 쿼터 그룹': ex_configs[j]['name'],
                                 '이 그룹 상한을 없애면': f"+{gain:,}명 확보 가능"}
                                for j, gain in sorted(d.group_relax_gain.items(),
                                                      key=lambda x: -x[1])]
                        st.markdown("**어느 그룹이 막고 있는지**")
                        st.dataframe(pd.DataFrame(rows), use_container_width=True,
                                     hide_index=True)

                    if d.value_relax_gain:
                        rows = [{'그룹': ex_configs[j]['name'],
                                 '항목': " / ".join(k) if isinstance(k, tuple) else str(k),
                                 '상한 +1명당 확보': f"+{gain:,}명"}
                                for (j, k), gain in sorted(d.value_relax_gain.items(),
                                                           key=lambda x: -x[1])]
                        st.markdown("**한도를 조금만 풀면 효과가 큰 항목** (섀도 프라이스)")
                        st.dataframe(pd.DataFrame(rows), use_container_width=True,
                                     hide_index=True)
                        st.caption("추가 쿼터 상한을 1명 늘렸을 때 전체 확보 인원이 "
                                   "몇 명 늘어나는지를 실제로 재계산한 값입니다.")

                    if not d.group_relax_gain and not d.value_relax_gain:
                        st.info("추가 쿼터를 전부 해제해도 인원이 늘지 않습니다. "
                                "미달은 순수하게 **데이터에 해당 응답자가 없어서**입니다. "
                                "표본을 더 확보하거나 목표를 조정해야 합니다.")

                if d.binding:
                    with st.expander(f"한도까지 꽉 찬 추가 쿼터 {len(d.binding)}건 (병목 후보)"):
                        st.dataframe(pd.DataFrame([
                            {'그룹': ex_configs[b['group']]['name'],
                             '항목': " / ".join(b['key']) if isinstance(b['key'], tuple)
                                     else str(b['key']),
                             '상한': b['cap'], '사용': b['used']}
                            for b in d.binding
                        ]), use_container_width=True, hide_index=True)

            # ------------------------------------------------------------------
            # 차트
            # ------------------------------------------------------------------
            def draw_chart(pairs, height_hint):
                """pairs: [(라벨, 목표, 달성), ...]"""
                if not pairs:
                    st.info("표시할 항목이 없습니다.")
                    return
                rows = []
                for label, tgt, act in pairs:
                    rows.append({'Label': label, 'Type': '1.목표', 'Value': tgt})
                    rows.append({'Label': label, 'Type': '2.달성', 'Value': act})
                dfc = pd.DataFrame(rows)
                dfc['sort_val'] = dfc['Label'].map(lambda x: tuple(utils.natural_key(x)))
                dfc = dfc.sort_values('sort_val')
                order = dfc['Label'].unique().tolist()
                chart = alt.Chart(dfc.drop(columns=['sort_val'])).mark_bar().encode(
                    y=alt.Y('Label:N', axis=alt.Axis(title=None), sort=order),
                    x=alt.X('Value:Q', axis=alt.Axis(title='인원수')),
                    color=alt.Color('Type:N',
                                    scale=alt.Scale(domain=['1.목표', '2.달성'],
                                                    range=['#e0e0e0', '#4c78a8']),
                                    legend=alt.Legend(title="구분")),
                    yOffset='Type:N'
                ).properties(height=min(4000, max(300, height_hint * 25)))
                st.altair_chart(chart, use_container_width=True)

            st.markdown("### 🔍 쿼터별 상세 현황")
            active_ex = [(j, cfg) for j, cfg in enumerate(ex_configs) if cfg['cols']]
            v_tabs = st.tabs(["메인 쿼터"] + [sheet_names.get(j, cfg['name']) for j, cfg in active_ex])

            with v_tabs[0]:
                if use_main:
                    draw_chart([(" / ".join(k), t, final_m[k]) for k, t in main_map.items()],
                               len(main_map))
                else:
                    st.info("메인 쿼터 설정이 없습니다.")

            for idx, (j, cfg) in enumerate(active_ex):
                with v_tabs[idx + 1]:
                    draw_chart(
                        [(" / ".join(k) if isinstance(k, tuple) else str(k), t, final_exs[j][k])
                         for k, t in cfg['map'].items()],
                        len(cfg['map'])
                    )

            if recs:
                st.divider()
                st.subheader("📉 부족 쿼터 분석 및 진단")
                df_recs = pd.DataFrame(recs)
                df_recs['sort_val'] = df_recs['항목'].map(lambda x: tuple(utils.natural_key(x)))
                df_recs = df_recs.sort_values(by=['순서', 'sort_val'])
                st.dataframe(df_recs.drop(columns=['순서', 'sort_val']),
                             use_container_width=True, hide_index=True)

        except Exception:
            st.error("오류 발생")
            st.code(traceback.format_exc())

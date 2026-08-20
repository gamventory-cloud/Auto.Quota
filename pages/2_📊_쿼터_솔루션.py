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
8. 계산 방식 선택 : 최선 보장(ILP) / 빠른 근사(그리디)

v3 변경점 (추가 쿼터 100% 할당)
------------------------------
 9. "🎯 추가 쿼터도 목표로 100% 맞추기" 옵션 추가 (ILP 전용)
    -> quota_ilp v2.0 의 ex_as_target. 추가 쿼터를 상한이 아니라 등식+부족변수로
       모델링해 부족까지 최소화한다. 초과는 두 모드 모두 금지.
10. 추가 쿼터 목표 0 을 '금지'로 보존 (예전엔 t>0 만 저장해 '무제한'으로 뒤집혔다)
11. 성공/실패 판정을 총량(total_fail)과 개별 쿼터(is_fail)로 분리.
    부족 분석은 항상 계산한다. 예전엔 총량만 채우면 미달이 숨겨졌다.
12. 실행 전 산술 프리플라이트 : quota_ilp.preflight_targets 로 교체.
    메인/추가 양쪽 유령 키·물리적 부족, 추가 목표 합계 정합성, 무응답자 수를 검사.
    (단일응답 추가 쿼터는 "그룹 목표 합계 == 메인 목표 합계" 가 100% 달성의
     필요조건이다. 이걸 실행 전에 숫자로 알려준다.)
13. min_fill 슬라이더 표시 버그 수정 (0.0~1.0 + "%.0f%%" -> 70% 가 "1%" 로 보였다)
14. 부족 분산 기준 선택 : 인원수 vs 목표 대비 비율(balance_relative)
15. Run_Info 시트 추가 (재현성 기록)
16. 하드 쿼터 + 추가 쿼터 허용 편차
    - "🔒 메인 쿼터를 하드 쿼터로" : 셀별 목표를 정확히 충족
    - "추가 쿼터 허용 편차" : 정확히 맞춤 / ±N명 / ±N% / 제한 없음
      총 선정 인원은 메인 쿼터가 정하므로 바뀌지 않고, 추가 쿼터의 개별
      항목만 목표 위아래로 나뉘어 흔들린다.
          50/50/50/50 (합 200)  ->  55/45/47/53 (합 200)
    - 완화 순서 : ① 추가 쿼터 편차 한계 -> ② 메인 하드
      (메인이 총량을 정의하므로 메인을 마지막에 풀어준다. 어느 단계에서
       풀렸는지는 ilp_sol.notes 로 화면에 표시된다)
    - 솔버는 항상 편차를 최소화하므로, 허용 편차 설정은 '이 범위를 넘으면
      알려달라'는 경고선으로 작동한다.
17. 추가 수집 지시서 (quota_ilp.plan_recruitment)
    - 메인 쿼터가 미달하면 "어떤 조건의 응답자를 몇 명 더 수집해야 하는지" 역산
    - 기존 표본 활용을 최대화해 필요 인원을 최소화한 뒤, 남는 추가 쿼터 편차를
      제곱 편차 기준으로 여러 항목에 고르게 분산
    - 데이터에 실제로 관측된 조합만 후보로 쓴다 (모집단에 없는 조건을 지시하면
      의미가 없으므로). 목표 0(금지) 키를 포함한 조합은 제외
    - 결과 엑셀 Recruit_Plan 시트로도 저장
18. 편차 분산에 제곱 편차 단계 추가
    - 최소최대만 쓰면 물리적으로 불가피한 큰 편차 하나가 최댓값을 포화시켜
      나머지를 고르게 나눌 동기가 사라진다 (자영 -84 vs 전문/자영/기타 각 -28)
19. '메인 쿼터를 하드 쿼터로' 옵션 제거
    - 사전식 최적화라 메인은 이미 사실상 하드다. 켜고 끄고에 따라 결과가
      달라지는 경우는 메인이 달성 불가능할 때뿐인데, 그때는 자동 완화로
      되돌아가므로 결국 같은 결과가 된다. 혼란만 주어 화면에서 뺐다.
    - quota_ilp 의 main_hard 인자는 기본값 False 로 남겨 호환을 유지한다.
20. [버그 수정] preflight_targets 호출에 존재하지 않는 인자를 넘기던 문제
    - main_hard / overflow_weight / ex_tol_* 가 잘못 섞여 들어가 있었다.
      실행 시 TypeError 로 죽는 자리였다. 유효 인자만 넘기도록 수정.
21. 화면 문구를 일상어로 전면 교체
    - 해(解)/최적성/희소성/섀도 프라이스/프로파일 같은 최적화 용어를 걷어냈다.
      "최적해임이 증명되었습니다" -> "이보다 많이 뽑을 수는 없습니다"
      "물리적 부족 / 경합 부족"   -> "표본이 모자람 / 다른 쿼터에 밀림"
      "정확해(ILP) / 휴리스틱"    -> "최선 보장(정밀) / 빠른 근사(간이)"
    - 코드 주석과 함수 문서는 원래 용어를 유지한다(유지보수용).
22. ID 컬럼과 intval 컬럼의 기본 선택을 이름으로 자동 매칭
    - intval / int_val / intValue 컬럼이 있으면 그것을 기본값으로 잡는다.
      대소문자와 앞뒤 공백은 무시한다. 없으면 첫 컬럼.
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

# ==============================================================================
# 사이드바 : 사용 방법 안내
#   본문 흐름을 방해하지 않도록 expander 로 접어 둔다. 순서는 화면 순서와 동일.
# ==============================================================================
with st.sidebar:
    st.divider()
    st.markdown("### 📖 사용 방법")

    with st.expander("① 데이터 올리기"):
        st.markdown(
            "응답자 원자료를 엑셀(.xlsx)이나 CSV로 올립니다. "
            "**한 행이 한 명**이어야 하고, 쿼터에 쓸 항목(성별·연령·지역·직업 등)이 "
            "각각 컬럼으로 있어야 합니다.\n\n"
            "- 값은 자동으로 정리됩니다 (`1.0` → `1`, 앞뒤 공백 제거)\n"
            "- 빈칸·결측은 `(무응답)`으로 묶입니다\n"
            "- CSV 인코딩은 자동 판별합니다 (UTF-8, CP949 등)"
        )

    with st.expander("② 쿼터 정하기"):
        st.markdown(
            "**메인 쿼터** — 표본의 뼈대입니다. 3개 항목을 교차해 셀마다 목표 인원을 "
            "정합니다. 예: 남 × 30대 × 서울 = 42명.\n\n"
            "- *엑셀 업로드* : 쓰던 쿼터표가 있으면 그대로 올립니다\n"
            "- *화면 설계* : 현재 분포가 채워진 표가 나오니 숫자만 고칩니다\n\n"
            "**추가 쿼터** — 직업·학력처럼 따로 관리할 항목입니다. 최대 4개까지 "
            "겹쳐 걸 수 있습니다.\n\n"
            "- *단순형* : 값 하나에 목표 하나 (복수응답 컬럼도 됩니다)\n"
            "- *조합형* : 행/열을 교차해서 목표를 줍니다\n\n"
            "목표에 **0을 적으면 '뽑지 않음'**이 됩니다. 칸을 비워두는 것과 다릅니다."
        )

    with st.expander("③ 실행 옵션 고르기"):
        st.markdown(
            "**계산 방식** — 특별한 이유가 없으면 `최선 보장`을 쓰세요. 더 정확하고 "
            "대개 더 빠릅니다.\n\n"
            "**총 인원 허용 오차** — 전체 합계 기준입니다. 0으로 두는 것을 권합니다. "
            "값을 주면 어느 셀이 모자라든 넘어갑니다.\n\n"
            "**intval 최적화** — 쿼터 조건이 똑같은 응답자 중에서 값이 낮은 쪽을 먼저 "
            "탈락시킵니다. 뽑는 인원수는 바뀌지 않고 누구를 뽑을지만 달라집니다. "
            "여유 인원이 있는 셀에서만 효과가 있습니다.\n\n"
            "**추가 쿼터도 목표로 100% 맞추기** — 켜면 추가 쿼터의 부족까지 최소화합니다. "
            "끄면 '넘지만 마라'는 상한으로만 씁니다."
        )

    with st.expander("④ 결과 읽기"):
        st.markdown(
            "**✅ 이보다 많이 뽑을 수는 없습니다** — 최선이라는 확인까지 끝났다는 "
            "뜻입니다. 더 손댈 게 없습니다.\n\n"
            "**⏱️ 시간 안에 끝내지 못했습니다** — 결과는 쓸 수 있지만 더 나은 조합이 "
            "있을 수 있습니다. 계산 시간 제한을 늘려보세요.\n\n"
            "부족이 생기면 사유가 셋 중 하나로 표시됩니다.\n\n"
            "- ⚠️ **표본이 모자람** → 사람을 더 모아야 합니다\n"
            "- ⚔️ **다른 쿼터에 밀림** → 쿼터 목표를 조정하면 풀립니다\n"
            "- ⚖️ **목표 합계가 안 맞음** → 추가 쿼터 합계를 메인과 맞추세요"
        )

    with st.expander("⑤ 받은 엑셀 파일"):
        st.markdown(
            "- `Result_Pass` : 최종 선정된 응답자\n"
            "- `Result_All` : 전체 응답자 + 선정/제외 표시\n"
            "- `Main_Status` : 메인 쿼터 셀별 목표 대비 달성\n"
            "- `Shortage_Analysis` : 모자란 쿼터와 그 이유\n"
            "- `Recruit_Plan` : 무엇을 몇 명 더 모아야 하는지\n"
            "- `Run_Info` : 실행 시각과 설정값 (재현용)\n"
            "- 추가 쿼터별 시트 : 항목마다 목표 대비 실제"
        )

    with st.expander("💡 잘 안 맞을 때"):
        st.markdown(
            "**추가 쿼터를 100%로 맞추려면** 그 쿼터의 목표 합계가 메인 쿼터 목표 "
            "합계와 같아야 합니다. 응답자 한 명은 항목 하나에만 계상되기 때문입니다. "
            "실행 전에 자동으로 검사해서 알려줍니다.\n\n"
            "**특정 셀만 텅 비면** `구하기 어려운 쿼터 먼저 채우기`를 켜고 "
            "`셀별 최소 달성률`을 조정해 보세요.\n\n"
            "**추가 쿼터가 목표에서 벗어나도 괜찮다면** 허용 편차를 `±N명`이나 "
            "`±N%`로 두세요. 총 인원은 그대로고 항목별로만 흔들립니다.\n\n"
            "**목표 표에 빈칸이 있으면** 그 값은 쿼터 관리 대상에서 빠집니다. "
            "실행 전 경고에 '목표 목록에 없는 값' 인원이 뜨면 확인해 보세요."
        )

    st.divider()
    st.caption("문제가 생기면 화면에 뜬 경고 문구와 함께 문의해 주세요.")

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
                            # [수정] 0 도 저장한다. 0 = "이 값은 뽑지 않는다"(금지).
                            # 예전엔 0 을 버려서 '무제한'으로 뒤집혔다.
                            config['map'][str(r['값'])] = t
                        warn_bad(bad, f"추가 쿼터 {i+1}")
                        st.caption(f"항목 {len(config['map'])}개 / "
                                   f"목표 합계 {sum(config['map'].values()):,}명")

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
                                config['map'][key] = t      # 0 = 금지 (위와 동일)
                            warn_bad(bad, f"추가 쿼터 {i+1}")
                            st.caption(f"셀 {len(config['map'])}개 / "
                                       f"목표 합계 {sum(config['map'].values()):,}명")

            ex_configs.append(config)

    # ==========================================================================
    # 3. 실행 옵션
    # ==========================================================================
    st.divider()
    st.subheader("3. 실행 옵션")
    solver_opts = (["최선 보장 (정밀)", "빠른 근사 (간이)"] if HAS_ILP
                   else ["휴리스틱 (그리디)"])
    solver_kind = st.radio(
        "계산 방식", solver_opts, horizontal=True,
        help=("최선 보장: 이보다 많이 뽑는 방법이 없다는 것까지 확인하고 끝냅니다. "
              "미달하면 어느 쿼터가 막고 있는지도 알려줍니다. 보통 이쪽이 더 빠릅니다. "
              "빠른 근사: 여러 번 시도해서 제일 좋았던 결과를 씁니다. "
              "더 나은 조합이 있을 수도 있습니다.")
    )
    if not HAS_ILP:
        st.caption(f"ℹ️ '최선 보장' 방식을 쓸 수 없습니다 (`pip install ortools` 필요) "
                   f"— {ILP_ERR}")
    use_ilp = solver_kind.startswith("최선 보장")

    # ── 추가 쿼터를 상한이 아니라 '목표'로 다룰지 ──────────────────────────
    ex_as_target = st.checkbox(
        "🎯 추가 쿼터도 목표로 100% 맞추기", value=False, disabled=not use_ilp,
        help="끄면 추가 쿼터는 상한으로만 작동합니다(초과 금지, 부족 허용). "
             "켜면 부족도 최소화합니다. 초과는 두 경우 모두 금지됩니다. "
             "'최선 보장' 방식에서만 지원합니다.")
    if ex_as_target:
        st.info(
            "추가 쿼터 그룹이 단일응답이면 **그룹의 목표 합계가 메인 목표 합계와 "
            "같아야** 양쪽 100%가 가능합니다. 응답자 한 명은 그 그룹의 항목 하나에 "
            "1명으로만 계상되기 때문입니다. 실행 직전에 자동 검사합니다."
        )
        unlisted_pol = st.radio(
            "목표 목록에 없는 값의 처리", ["제약 없이 선택 가능", "선택 대상에서 제외"],
            horizontal=True,
            help="목표 표에서 지운 값을 가진 응답자를 어떻게 볼지 결정합니다.")
        unlisted = "free" if unlisted_pol.startswith("제약") else "forbid"
        # ── 추가 쿼터 허용 편차 ────────────────────────────────────────
        # 총 선정 인원은 메인 쿼터가 정하므로 바뀌지 않는다. 개별 항목만
        # 목표 위아래로 흔들린다.  50/50/50/50 → 55/45/47/53 (합 200 유지)
        st.markdown("**추가 쿼터 허용 편차**")
        tol_mode = st.radio(
            "항목별로 목표에서 얼마나 벗어나도 되는지",
            ["정확히 맞춤 (하드)", "±N명까지", "±N%까지", "제한 없음"],
            horizontal=True, label_visibility="collapsed",
            help="솔버는 항상 편차를 최소로 만듭니다. 이 설정은 '이 범위를 넘으면 "
                 "알려달라'는 경고선입니다. 범위 안에서 맞출 수 없으면 자동으로 "
                 "한계를 풀고, 그때 최소 편차가 얼마인지 알려줍니다.")
        ex_tol_abs, ex_tol_pct, ex_tol_unlimited = 0, 0.0, False
        if tol_mode.startswith("±N명"):
            ex_tol_abs = st.number_input("허용 편차 (명)", 1, 10000, 10)
        elif tol_mode.startswith("±N%"):
            ex_tol_pct = st.number_input("허용 편차 (%)", 1, 100, 5) / 100.0
        elif tol_mode.startswith("제한"):
            ex_tol_unlimited = True
        # '정확히 맞춤'도 달성 불가하면 자동 완화되므로 편차 허용 자체는 켜둔다
        ex_overflow = not tol_mode.startswith("정확히")
        overflow_weight = 1
    else:
        unlisted = "free"
        ex_overflow, overflow_weight = False, 1
        ex_tol_abs, ex_tol_pct, ex_tol_unlimited = 0, 0.0, False

    # [제거] '메인 쿼터를 하드 쿼터로' 옵션
    #   사전식 최적화라 1단계에서 메인 부족을 최소화하고 그 값을 고정한 뒤에야
    #   추가 쿼터를 다룬다. 따라서 메인이 달성 가능하면 이 옵션과 무관하게 항상
    #   100% 채워지고, 추가 쿼터에 양보하는 일은 구조적으로 없다.
    #   달성 불가능할 때만 동작이 갈리는데 그때는 INFEASIBLE 이 나서 자동 완화로
    #   되돌아가므로 결국 끈 것과 같은 결과가 된다. 혼란만 주어 화면에서 뺐다.
    #   quota_ilp.solve_quota_ilp 의 main_hard 인자는 기본값 False 로 남아 있다.

    c1, c2 = st.columns(2)
    with c1:
        def _col_idx(cols, *names):
            """컬럼 목록에서 이름이 일치하는 것을 찾아 기본 선택 위치를 돌려준다.
            대소문자와 앞뒤 공백은 무시한다. 없으면 0(첫 컬럼)."""
            low = [str(c).strip().lower() for c in cols]
            for nm in names:
                if nm.lower() in low:
                    return low.index(nm.lower())
            return 0

        cols_all = list(df_survey.columns)
        c_no = st.selectbox("ID 컬럼", cols_all,
                            index=_col_idx(cols_all, "id", "no", "번호", "일련번호"))
        tol = st.number_input(
            "총 인원 허용 오차(명)", 0, 100, 0,
            help="목표 인원에서 이 인원까지 모자라도 '달성'으로 봅니다. "
                 "쿼터별이 아니라 전체 합계 기준입니다. 0이면 한 명도 모자라면 안 됩니다.")
        use_intval = st.checkbox(
            "intval 최적화", value=True,
            help="쿼터 조건이 완전히 같은 응답자들 사이에서, intval 값이 낮은 쪽을 "
                 "먼저 탈락시킵니다. 조건이 다른 응답자끼리는 영향이 없으므로 "
                 "최종 통과 인원수는 달라지지 않습니다.")
        c_int = st.selectbox(
            "intval 컬럼", cols_all,
            index=_col_idx(cols_all, "intval", "int_val", "intValue")
        ) if use_intval else None
        if not use_intval:
            rand_pick = st.checkbox(
                "동일 조건 응답자 무작위 선택", value=True,
                help="끄면 데이터 순서대로 뽑아 결과가 완전히 재현됩니다.")
        else:
            rand_pick = False
    with c2:
        if use_ilp:
            time_limit = st.number_input(
                "계산 시간 제한(초)", 5, 600, 60, 5,
                help="이 시간 안에 끝내지 못하면 그때까지 찾은 가장 좋은 조합을 씁니다.")
            ilp_priority = st.checkbox(
                "구하기 어려운 쿼터 먼저 채우기", value=True,
                help="뽑을 인원을 최대로 확정한 뒤, 대신할 사람이 없는 귀한 조건의 셀을 "
                     "먼저 채웁니다. 총 인원은 줄지 않습니다. 끄면 어느 셀을 "
                     "채울지 임의로 결정됩니다.")
            # [수정] 0.0~1.0 값에 format="%.0f%%" 를 쓰면 0.7 이 "1%" 로 표시됐다.
            min_fill = (st.slider(
                "셀별 최소 달성률", 0, 100, 70, 5, format="%d%%",
                disabled=not ilp_priority,
                help="귀한 셀을 먼저 채우더라도 어떤 셀도 이 비율 밑으로 떨어지지 "
                     "않게 합니다. 0%로 두면 흔한 셀이 0명이 될 수 있습니다. "
                     "만족 불가능하면 자동으로 하한 없이 재계산하고 알려줍니다."
            ) / 100.0) if ilp_priority else 0.0
            balance = st.checkbox(
                "부족분 고르게 분산", value=True,
                help="귀한 셀을 먼저 채운 뒤, 남은 부족분이 특정 셀에 몰리지 않도록 "
                     "나눠 줍니다. 나중에 적용되므로 귀한 셀의 자리를 빼앗지 않습니다.")
            want_plan = st.checkbox(
                "미달 시 추가 수집 지시서 계산", value=True,
                help="메인 쿼터가 미달하면 '어떤 조건의 응답자를 몇 명 더 수집해야 "
                     "하는지'를 역산합니다. 계산이 한 번 더 돌아가므로 표본이 매우 "
                     "크면 시간이 조금 늘어납니다.")
            balance_rel = st.checkbox(
                "부족을 목표 대비 비율로 분산", value=True,
                disabled=not balance,
                help="목표 1000인 셀의 50명 부족(5%)과 목표 100인 셀의 50명 "
                     "부족(50%)을 같게 보지 않습니다. 끄면 인원수 기준입니다.")
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
            time_limit, balance, ilp_priority, min_fill = 0, False, False, 0.0
            balance_rel, want_plan = False, False

    if st.button("🚀 매칭 시작", type="primary"):
        if not main_map:
            st.error("목표가 설정되지 않았습니다.")
            st.stop()
        if use_main and not algo_main_cols:
            st.error("메인 쿼터 변수를 선택하세요.")
            st.stop()

        try:
            with st.spinner("쿼터 조건 정리 중..."):
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
                ex_maps = [c['map'] for c in ex_configs]

                # ------------------------------------------------------------------
                # 프리플라이트 : 데이터에 아예 존재하지 않는 쿼터 키 경고
                # (정규화 불일치를 실행 전에 잡아내는 안전망)
                # ------------------------------------------------------------------
                # [교체] 메인 유령셀만 보던 검사를 quota_ilp 의 종합 프리플라이트로.
                #  - 메인/추가 양쪽의 유령 키·물리적 부족
                #  - 추가 쿼터 목표 합계가 메인 합계와 맞는지 (초과 금지 시 필수 조건)
                #  - 추가 쿼터 변수의 무응답자 수
                pre = []
                if HAS_ILP:
                    pre = quota_ilp.preflight_targets(
                        m_keys, ex_keys_list, main_map, ex_maps,
                        ex_as_target=ex_as_target, unlisted=unlisted,
                        ex_overflow=ex_overflow)
                else:
                    ghosts = [k for k in main_map if m_cnt.get(k, 0) == 0]
                    if ghosts:
                        pre = [{'level': 'error', 'group': None, 'kind': 'main_ghost',
                                'msg': (f"메인 쿼터 {len(ghosts)}개 셀이 데이터에 한 명도 "
                                        f"없습니다. 목표 "
                                        f"{sum(main_map[k] for k in ghosts):,}명은 "
                                        f"달성 불가입니다.")}]

                def _gname(d):
                    j = d.get('group')
                    return "" if j is None else f"[{ex_configs[j]['name']}] "

                for d in pre:
                    if d['level'] == 'error':
                        st.error(f"❌ {_gname(d)}{d['msg']}")
                    elif d['level'] == 'warn':
                        st.warning(f"⚠️ {_gname(d)}{d['msg']}")
                    else:
                        st.caption(f"✅ {_gname(d)}{d['msg']}")

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
                with st.spinner("최적 조합 계산 중..."):
                    ilp_sol = quota_ilp.solve_quota_ilp(
                        m_keys, ex_keys_list, main_map, ex_maps, indices,
                        priority=ilp_priority, balance=balance,
                        balance_relative=balance_rel,
                        min_fill=min_fill, time_limit=time_limit,
                        workers=n_cores, rng=pick_rng, tiebreak=tiebreak,
                        ex_as_target=ex_as_target, unlisted=unlisted,
                        ex_overflow=ex_overflow,
                        overflow_weight=overflow_weight,
                        ex_tol_abs=ex_tol_abs, ex_tol_pct=ex_tol_pct,
                        ex_tol_unlimited=ex_tol_unlimited)
                g_best_cnt, g_best_idxs = ilp_sol.total, ilp_sol.selected

            # ======================================================================
            # 실행 (B) 휴리스틱 : 랜덤 재시작 그리디
            # ======================================================================
            else:
              with st.spinner("여러 조합을 시도하는 중..."):
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

            # [수정] 총량만 보던 판정을 분리한다. 총량을 채웠어도 개별 쿼터가
            # 미달일 수 있고, 예전에는 그 경우 부족 분석이 통째로 생략됐다.
            total_fail = g_best_cnt < soft_target

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
            if True:            # [수정] 항상 계산한다 (예전엔 is_fail 일 때만)
                if use_main:
                    for k, tgt in main_map.items():
                        act = final_m.get(k, 0)
                        diff = tgt - act
                        if diff > 0:
                            raw_avail = m_cnt.get(k, 0)
                            reason = ("⚠️ 표본이 모자람" if raw_avail < tgt
                                      else "⚔️ 다른 쿼터에 밀림")
                            recs.append({'순서': 0, '구분': '메인 쿼터', '항목': " / ".join(k),
                                         '목표': tgt, '현재': act, '부족': diff,
                                         '진단': reason, '전체보유': raw_avail})

                for j, cfg in enumerate(ex_configs):
                    if not cfg['cols']:
                        continue
                    raw_cnt_map = collections.Counter(
                        v for keys in ex_keys_list[j] for v in keys
                    )
                    struct_bad = {d.get('group') for d in pre
                                  if d['kind'] in ('group_sum_low', 'group_sum_high')}
                    for k, tgt in cfg['map'].items():
                        act = final_exs[j].get(k, 0)
                        diff = tgt - act
                        if diff > 0:
                            raw_avail = raw_cnt_map.get(k, 0)
                            if raw_avail < tgt:
                                reason = "⚠️ 표본이 모자람"
                            elif j in struct_bad:
                                reason = "⚖️ 목표 합계가 안 맞음"
                            else:
                                reason = "⚔️ 다른 쿼터에 밀림"
                            label = " / ".join(k) if isinstance(k, tuple) else str(k)
                            recs.append({'순서': j + 1, '구분': cfg['name'], '항목': label,
                                         '목표': tgt, '현재': act, '부족': diff,
                                         '진단': reason, '전체보유': raw_avail})

            # ------------------------------------------------------------------
            # 최종 판정 : 총량 미달 또는 개별 쿼터 미달
            # ------------------------------------------------------------------
            main_short_recs = [r for r in recs if r['구분'] == '메인 쿼터']
            ex_short_recs = [r for r in recs if r['구분'] != '메인 쿼터']
            ex_short_sum = sum(r['부족'] for r in ex_short_recs)
            if ex_as_target:
                # 편차를 허용한 경우 추가 쿼터의 벗어남은 실패가 아니라 '편차'로 본다.
                # 정확히 맞춤(하드)을 요구했을 때만 미달을 실패로 판정한다.
                is_fail = total_fail or bool(main_short_recs) or (
                    not ex_overflow and bool(ex_short_recs))
            else:
                is_fail = total_fail

            # ------------------------------------------------------------------
            # 추가 수집 지시서 : 부족분을 어떤 구성으로 보충해야 하는가
            # ------------------------------------------------------------------
            plan, plan_rows = None, []
            if want_plan and use_ilp and HAS_ILP and main_short_recs:
                with st.spinner("추가 수집 지시서 역산 중..."):
                    try:
                        plan = quota_ilp.plan_recruitment(
                            m_keys, ex_keys_list, main_map, ex_maps,
                            unlisted=unlisted, ex_tol_abs=ex_tol_abs,
                            ex_tol_pct=ex_tol_pct,
                            ex_tol_unlimited=(ex_tol_unlimited or not ex_overflow),
                            time_limit=max(20, time_limit), workers=n_cores)
                    except Exception as _pe:                      # noqa: BLE001
                        st.warning(f"⚠️ 추가 수집 지시서 계산 실패 — "
                                   f"{type(_pe).__name__}: {_pe}")
                if plan is not None and plan.feasible:
                    for r in plan.rows:
                        cond = " · ".join(
                            f"{ex_configs[j]['name']}={'/'.join(str(x) for x in ks)}"
                            for j, ks in r['pattern'].items())
                        plan_rows.append({
                            '메인 셀': " / ".join(r['cell']),
                            '추가 조건': cond or "(조건 없음)",
                            '추가 수집 인원': r['n']})

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

                if plan_rows:
                    pd.DataFrame(plan_rows).to_excel(
                        w, index=False, sheet_name='Recruit_Plan')

                pd.DataFrame([
                    {'항목': '실행 시각', '값': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')},
                    {'항목': '계산 방식', '값': solver_kind},
                    {'항목': '총 인원 허용 오차(명)', '값': tol},
                    {'항목': '추가 쿼터 처리', '값': '목표(100% 지향)' if ex_as_target else '상한'},
                    {'항목': '추가 쿼터 허용 편차',
                     '값': (f"±{ex_tol_abs}명" if ex_tol_abs else
                            f"±{ex_tol_pct:.0%}" if ex_tol_pct else
                            "제한 없음" if ex_tol_unlimited else "정확히 맞춤")},
                    {'항목': '목록 밖 값', '값': unlisted},
                    {'항목': '메인 목표 합계', '값': target_total},
                    {'항목': '선정 인원', '값': len(df_pass)},
                    {'항목': '추가 쿼터 부족', '값': ex_short_sum},
                    {'항목': '계산 시간 제한(초)', '값': time_limit},
                    {'항목': '귀한 쿼터 우선 / 최소달성률',
                     '값': f"{ilp_priority} / {min_fill:.0%}"},
                    {'항목': '부족 분산 / 비율기준', '값': f"{balance} / {balance_rel}"},
                    {'항목': '시도 횟수(휴리스틱)', '값': iters},
                    {'항목': '지터(휴리스틱)', '값': jitter},
                    {'항목': 'intval 컬럼', '값': str(c_int)},
                ]).to_excel(w, index=False, sheet_name='Run_Info')

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
                        "(구하기 어려운 셀에서는 값이 낮아도 뽑아야 합니다)."
                    )

            # 추가 쿼터 달성 현황 (목표 모드에서만 의미가 있다)
            if ex_as_target and ex_configs:
                ex_target_sum = sum(sum(c['map'].values())
                                    for c in ex_configs if c['cols'])
                if ex_target_sum:
                    e1, e2, e3 = st.columns(3)
                    e1.metric("🎯 추가 쿼터 목표 합계", f"{ex_target_sum:,}명")
                    e2.metric("✅ 추가 쿼터 달성",
                              f"{ex_target_sum - ex_short_sum:,}명")
                    e3.metric("📉 추가 쿼터 부족", f"{ex_short_sum:,}명",
                              delta="충족" if ex_short_sum == 0 else f"-{ex_short_sum:,}명",
                              delta_color="normal" if ex_short_sum == 0 else "inverse")

            # 추가 쿼터 편차 현황 (총 인원은 유지되고 항목만 흔들린다)
            ex_dev_recs = []
            if ex_as_target:
                for j, cfg in enumerate(ex_configs):
                    if not cfg['cols']:
                        continue
                    for k, tgt in cfg['map'].items():
                        act = final_exs[j].get(k, 0)
                        if act != tgt:
                            ex_dev_recs.append({
                                '구분': cfg['name'],
                                '항목': " / ".join(k) if isinstance(k, tuple) else str(k),
                                '목표': tgt, '실제': act, '편차': act - tgt,
                                '편차율': (act - tgt) / tgt if tgt else None})
                if ex_dev_recs:
                    mx = max(abs(r['편차']) for r in ex_dev_recs)
                    mxr = max((abs(r['편차율']) for r in ex_dev_recs
                               if r['편차율'] is not None), default=0)
                    d1, d2, d3 = st.columns(3)
                    d1.metric("📐 편차 발생 항목", f"{len(ex_dev_recs)}개")
                    d2.metric("최대 편차", f"{mx:,}명")
                    d3.metric("최대 편차율", f"{mxr:.1%}")
                    st.caption(
                        "총 선정 인원은 메인 쿼터가 정하므로 그대로이고, 추가 쿼터의 "
                        "개별 항목만 목표 위아래로 나뉘어 흔들립니다. 아래 편차는 "
                        "이 조건에서 가능한 최소값입니다.")
                    dfd = pd.DataFrame(ex_dev_recs)
                    st.dataframe(
                        dfd.style.format({'편차율': '{:+.1%}'}, na_rep="-")
                           .background_gradient(subset=['편차'], cmap='RdYlGn_r'),
                        use_container_width=True, hide_index=True)

            if not is_fail:
                if ex_as_target:
                    st.success("🎉 메인 쿼터와 추가 쿼터를 **모두 100% 달성**했습니다!")
                else:
                    st.success("🎉 목표 인원을 모두 달성했습니다!")
            elif not total_fail:
                st.warning(
                    f"⚠️ 메인 목표 인원({target_total:,}명)은 채웠지만 개별 쿼터가 "
                    f"미달입니다. 메인 {len(main_short_recs)}개 셀 / 추가 "
                    f"{len(ex_short_recs)}개 항목 — 아래 분석을 확인하세요.")
            else:
                st.error("⚠️ 목표 인원을 달성하지 못했습니다. 아래 분석을 확인하세요.")

            # ------------------------------------------------------------------
            # ILP 전용: 최적성 보증 + 병목 진단
            # ------------------------------------------------------------------
            if ilp_sol is not None:
                if ilp_sol.proven_optimal:
                    extra = ""
                    if ilp_sol.ex_as_target:
                        extra = (f" 추가 쿼터가 모자란 {sum(ilp_sol.ex_short_total):,}명도 "
                                 "더 줄일 수 없는 최소치입니다.")
                    st.success(
                        f"✅ **이보다 많이 뽑을 수는 없습니다.** 지금 쿼터 조건에서 "
                        f"{ilp_sol.total:,}명이 최대이고, 어떤 조합을 시도해도 이 숫자를 "
                        f"넘길 수 없다는 것까지 확인했습니다.{extra} "
                        f"(응답자를 {ilp_sol.n_profiles:,}개 유형으로 묶어 "
                        f"{ilp_sol.solve_sec:.2f}초 만에 계산)"
                    )
                else:
                    st.warning(
                        f"⏱️ {time_limit}초 안에 계산을 끝내지 못했습니다. 아래 결과는 "
                        f"그때까지 찾은 것 중 가장 좋은 조합이며, 더 나은 조합이 있을 "
                        f"수도 있습니다. 시간 제한을 늘리면 확인할 수 있습니다. "
                        f"(내부 상태: {ilp_sol.status})"
                    )

                for _n in getattr(ilp_sol, "notes", []):
                    st.warning(f"⚠️ {_n}")

                d = ilp_sol.diagnosis
                if is_fail:
                    st.markdown("#### 🧭 왜 목표를 못 채웠나")

                    if d.group_relax_gain:
                        rows = [{'추가 쿼터': ex_configs[j]['name'],
                                 '이 쿼터를 빼면': f"{gain:,}명 더 뽑을 수 있음"}
                                for j, gain in sorted(d.group_relax_gain.items(),
                                                      key=lambda x: -x[1])]
                        st.markdown("**어느 쿼터가 막고 있는지** "
                                    "(그 쿼터를 아예 빼고 계산했을 때)")
                        st.dataframe(pd.DataFrame(rows), use_container_width=True,
                                     hide_index=True)

                    if d.value_relax_gain:
                        rows = [{'그룹': ex_configs[j]['name'],
                                 '항목': " / ".join(k) if isinstance(k, tuple) else str(k),
                                 '목표 1명 늘릴 때': f"{gain:,}명 더 뽑힘"}
                                for (j, k), gain in sorted(d.value_relax_gain.items(),
                                                           key=lambda x: -x[1])]
                        st.markdown("**조금만 늘려주면 효과가 큰 항목**")
                        st.dataframe(pd.DataFrame(rows), use_container_width=True,
                                     hide_index=True)
                        st.caption("이 항목의 목표를 딱 1명 늘려서 다시 계산해 본 결과입니다. "
                                   "1명 늘렸는데 여러 명이 더 뽑힌다면, 그 항목이 전체를 "
                                   "막고 있다는 뜻입니다.")

                    if not d.group_relax_gain and not d.value_relax_gain:
                        st.info("추가 쿼터를 전부 없애고 계산해도 인원이 늘지 않습니다. "
                                "**데이터에 그 조건의 응답자가 아예 없어서** 모자란 "
                                "것입니다. 표본을 더 모으거나 목표를 낮춰야 합니다.")

                if d.binding:
                    with st.expander(f"목표를 다 채워 더 못 받는 추가 쿼터 {len(d.binding)}건"):
                        st.dataframe(pd.DataFrame([
                            {'그룹': ex_configs[b['group']]['name'],
                             '항목': " / ".join(b['key']) if isinstance(b['key'], tuple)
                                     else str(b['key']),
                             '목표': b['cap'], '채운 인원': b['used']}
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

            if plan is not None and plan.feasible and plan_rows:
                st.divider()
                st.subheader("🧾 추가 수집 지시서")
                p1, p2, p3 = st.columns(3)
                p1.metric("총 추가 수집 필요", f"{plan.total_needed:,}명")
                p2.metric("대상 메인 셀", f"{len(plan.by_cell)}개")
                p3.metric("보충 후 최대 편차", f"{plan.max_dev_after:,}명")
                st.caption(
                    "메인 쿼터를 100% 채우기 위해 어떤 조건의 응답자를 몇 명 더 "
                    "확보해야 하는지 역산한 결과입니다. 기존 표본을 최대한 활용하는 "
                    "전제에서 필요 인원이 최소가 되도록 계산했고, 추가 쿼터 조건은 "
                    "데이터에 실제로 존재하는 조합만 제시합니다.")
                st.dataframe(pd.DataFrame(plan_rows), use_container_width=True,
                             hide_index=True)
                for _n in plan.notes:
                    st.warning(f"⚠️ {_n}")
                if plan.max_dev_after:
                    st.caption(
                        f"이 인원을 모두 확보해도 추가 쿼터에 최대 "
                        f"{plan.max_dev_after:,}명의 편차가 남습니다. 이미 선정이 "
                        f"확정된 응답자 구성 때문에 피할 수 없는 부분입니다.")
            elif plan is not None and not plan.feasible:
                st.divider()
                st.subheader("🧾 추가 수집 지시서")
                for _n in (plan.notes or ["계산 결과가 없습니다."]):
                    st.error(f"❌ {_n}")

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

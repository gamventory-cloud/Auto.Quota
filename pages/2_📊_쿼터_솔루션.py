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
20-B. [버그 수정] Styler.background_gradient 가 matplotlib 없이 죽던 문제
    - 추가 쿼터 편차 표에서 ImportError: background_gradient requires matplotlib.
      matplotlib 은 pandas 필수 의존성이 아니라 로컬 환경에서 터졌다.
    - Styler 를 걷어내고 st.column_config 로 교체했다. 편차 크기는 색 대신
      ProgressColumn 막대로 표현하므로 추가 패키지가 필요 없다.
21. 화면 문구를 일상어로 전면 교체
    - 해(解)/최적성/희소성/섀도 프라이스/프로파일 같은 최적화 용어를 걷어냈다.
      "최적해임이 증명되었습니다" -> "이보다 많이 뽑을 수는 없습니다"
      "물리적 부족 / 경합 부족"   -> "표본이 모자람 / 다른 쿼터에 밀림"
      "정확해(ILP) / 휴리스틱"    -> "최선 보장(정밀) / 빠른 근사(간이)"
    - 코드 주석과 함수 문서는 원래 용어를 유지한다(유지보수용).
22-B. intval 범위 밖 응답자 제외 (최소값 / 최대값 두 칸)
    - "값이 범위를 벗어난 응답자는 후보에서 제외" + 최소·최대 직접 입력
    - 메인 쿼터 목표에 없는 키를 부여해 후보 목록에서 빼는 방식이라
      quota_ilp / utils 를 건드리지 않는다 (main_map.get(k,0) <= 0 이면 제외)
    - 제외된 응답자는 결과 엑셀 Chk 열에 "제외(intval 범위)" 로 사유가 남는다
    - 위젯에 분포(최소/중앙값/최대/하위 1%/상위 1%)와 제외 인원을 미리 보여준다
    - 최소값이 최대값보다 크면 전원 제외되므로 입력 단계에서 경고한다
23. 추가 쿼터 그룹을 4개 -> 8개로 확장 (MAX_EXTRA)
24. 추가 쿼터 달성 현황을 '합산' 에서 '그룹별' 로 교체
    - 그룹끼리 목표를 더하면 안 된다. 단일응답이면 응답자 한 명이 그룹마다
      1명씩 계상되므로, 3개 그룹이면 합계가 전체 목표의 3배가 된다.
      (전체 목표 1,300명인데 "추가 쿼터 목표 합계 3,900명" 으로 표시되던 문제)
    - 그룹별 목표/달성/부족/초과/어긋난 항목 수와 상태를 표로 보여준다.
25. 쿼터 설정 프리셋 (파일명 기준 저장/복원)
    - 같은 조사를 여러 번 처리할 때 목표를 매번 다시 입력하지 않게 한다.
    - 저장 : .quota_presets/<파일명>.json  +  JSON 다운로드
      (클라우드는 재시작 시 디스크가 비므로 JSON 병행이 필수다)
    - 불러오기 : 같은 파일명이 올라오면 자동 감지 -> 확인 후 적용.
      조용히 덮어쓰지 않는다.
    - 데이터가 바뀐 경우 대조해서 알려준다.
      "이번에 새로 생긴 값은 현재 분포로 채움 / 저장된 설정에만 있던 값은 없음"
    - 복원 대상 : 메인 사용 여부·방식·행열 변수·셀별 목표,
      추가 쿼터 그룹별 방식·변수·목표, 실행 옵션 일부
26. SPSS .sav 업로드 지원
    - 값 라벨을 "1) 서울" 형태로 합쳐서 읽는다 (utils.read_sav_combined).
      라벨만 쓰면 화면 정렬이 가나다순이 되어 코드 순서와 어긋난다.
      예: 1서울 2부산 3대구 -> 광주, 대구, 부산, 서울 ... 로 뒤섞임.
      코드를 앞에 붙이면 natural_key 가 숫자를 먼저 보므로 순서가 유지된다.
    - 값 라벨이 없는 변수는 원래 값 그대로 둔다.
    - pyreadstat 필요 (requirements.txt 에 추가).
27. 결과를 SPSS(.sav) 로도 받을 수 있게 함
    - "1) 서울" 로 읽어들인 값을 코드(1)로 되돌리고 값 라벨을 다시 입혀 저장한다.
      그대로 저장하면 SPSS 에서 문자열 변수가 되어 납품에 쓸 수 없다.
    - 값 라벨·변수 라벨이 원본과 동일하게 복원되는 것을 왕복 테스트로 확인했다.
    - .sav 에는 시트가 없으므로 선정자 데이터만 담는다. 부족 분석·지시서 등은
      엑셀 파일에만 들어간다.
    - 엑셀/CSV 를 올린 경우엔 값 라벨 정보가 없어 값 그대로 저장되며,
      한글 컬럼명은 SPSS 변수명 규칙에 맞게 바꾸고 무엇이 바뀌었는지 알려준다.
28. [프리셋 보완] 메인 쿼터 미사용 시의 '전체 목표' 도 저장/복원
    - 저장 대상이 아니어서 불러온 뒤 매번 1000 에서 다시 입력해야 했다.
    - 고정 key("QS_total_only") 를 붙여 리런 중에도 값이 유지되게 했다.
29. [프리셋 버그] 불러오기가 일부 위젯에 안 먹던 문제
    - key 가 붙은 위젯(ms{i}, ed{i}, ex_mode_{i} 등)은 session_state 값이
      default/index 보다 우선한다. 그래서 한 번이라도 만진 뒤 프리셋을
      불러오면 이전 값이 그대로 남아 "불러왔는데 안 바뀐다" 가 됐다.
    - 프리셋 적용/해제/업로드 시 _reset_widget_state() 로 해당 키만 지운다.
      ("이름 + 숫자" 형태만 정규식으로 정확히 집는다)
30. ID 컬럼과 intval 컬럼의 기본 선택을 이름으로 자동 매칭
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


# ==============================================================================
# 쿼터 설정 프리셋
#   같은 조사를 여러 번 처리할 때 목표를 매번 다시 입력하지 않도록 저장/복원한다.
#
#   저장 위치가 두 곳인 이유
#     - 로컬 실행 : .quota_presets/ 폴더에 그대로 남는다
#     - 클라우드   : 앱이 재시작되면 디스크가 초기화되므로 폴더 저장은 사라진다.
#                   그래서 JSON 다운로드/업로드를 함께 제공한다.
#   매칭은 업로드한 데이터 파일명 기준이다.
# ==============================================================================
import json
import re as _re
from datetime import datetime
from pathlib import Path

PRESET_DIR = Path(".quota_presets")
PRESET_VER = 1
KEY_SEP = "\u0001"          # 튜플 키를 JSON 문자열로 만들 때 쓰는 구분자


def _preset_slug(filename: str) -> str:
    """파일명을 저장용 이름으로 바꾼다 (확장자 제거, 특수문자 정리)."""
    stem = Path(str(filename)).stem
    return _re.sub(r'[^0-9A-Za-z가-힣._-]+', '_', stem)[:120] or "preset"


def _enc_key(k):
    return KEY_SEP.join(str(x) for x in k) if isinstance(k, tuple) else str(k)


def _dec_key(s, as_tuple):
    return tuple(s.split(KEY_SEP)) if as_tuple else s


def preset_payload(source_name, main_state, extras_state, options):
    return {
        "version": PRESET_VER,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_file": str(source_name),
        "main": main_state,
        "extras": extras_state,
        "options": options,
    }


def preset_save_local(slug, payload):
    """로컬 폴더에 저장. 쓰기 권한이 없으면 조용히 실패하고 사유를 돌려준다."""
    try:
        PRESET_DIR.mkdir(exist_ok=True)
        (PRESET_DIR / f"{slug}.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return True, str(PRESET_DIR / f"{slug}.json")
    except Exception as e:                                    # noqa: BLE001
        return False, f"{type(e).__name__}: {e}"


def preset_load_local(slug):
    f = PRESET_DIR / f"{slug}.json"
    if not f.exists():
        return None
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:                                         # noqa: BLE001
        return None


def preset_list_local():
    if not PRESET_DIR.exists():
        return []
    out = []
    for f in sorted(PRESET_DIR.glob("*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            out.append((f.stem, d.get("saved_at", ""), d.get("source_file", "")))
        except Exception:                                     # noqa: BLE001
            continue
    return out


def _reset_widget_state():
    """
    프리셋을 적용/해제할 때 입력 위젯의 저장된 상태를 지운다.

    key 가 붙은 위젯은 st.session_state 값이 default/index 보다 우선한다.
    그래서 지우지 않으면 프리셋을 불러와도 이전에 만졌던 값이 그대로 남아
    "불러왔는데 안 바뀐다" 가 된다.
    """
    # 접두사만 보면 msg_xxx, edit_xxx 처럼 무관한 키까지 걸리므로
    # "이름 + 숫자" 형태만 정확히 집는다.
    pat = _re.compile(r'^(ms|ed|ex_mode_|ex_rv_|ex_cv_|ex_ed_grid_)\d+$')
    for k in [k for k in list(st.session_state)
              if k == "QS_total_only" or pat.match(k)]:
        st.session_state.pop(k, None)


def preset_get(key, default=None):
    """현재 적용된 프리셋에서 값을 꺼낸다. 없으면 default."""
    p = st.session_state.get("QS_preset")
    if not p:
        return default
    cur, path = p, key.split(".")
    for step in path:
        if not isinstance(cur, dict) or step not in cur:
            return default
        cur = cur[step]
    return cur


def preset_targets(section, idx=None):
    """저장된 목표 dict 를 {키: 목표} 로 되돌린다. 없으면 빈 dict."""
    if section == "main":
        raw = preset_get("main.targets")
        as_tuple = True
    else:
        exs = preset_get("extras") or []
        if idx is None or idx >= len(exs):
            return {}
        raw = (exs[idx] or {}).get("targets")
        as_tuple = (exs[idx] or {}).get("mode") == "grid"
    if not isinstance(raw, dict):
        return {}
    return {_dec_key(k, as_tuple): v for k, v in raw.items()}


def apply_preset_col(df_col_values, preset_map, fallback):
    """
    프리셋 목표를 현재 데이터의 값 목록에 맞춰 채운다.
      - 프리셋에 있는 값 -> 저장된 목표
      - 이번에 새로 생긴 값 -> fallback (현재 분포)
    반환: (목표 리스트, 프리셋에만 있던 값, 이번에만 있는 값)
    """
    vals = list(df_col_values)
    out, only_new = [], []
    for v, fb in zip(vals, fallback):
        if v in preset_map:
            out.append(int(preset_map[v]))
        else:
            out.append(fb)
            only_new.append(v)
    only_old = [k for k in preset_map if k not in set(vals)]
    return out, only_old, only_new


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
            "**추가 쿼터** — 직업·학력처럼 따로 관리할 항목입니다. 최대 8개까지 "
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
data_file = st.file_uploader("설문 데이터", type=['csv', 'xlsx', 'sav'],
                             key="quota_up",
                             help="SPSS(.sav) 는 값 라벨을 \"1) 서울\" 형태로 읽습니다. "
                                  "코드 번호가 앞에 붙어야 화면과 쿼터표의 값 순서가 "
                                  "코드 순서대로 유지됩니다.")

if data_file:
    df_survey, _sav_meta = utils.load_df(data_file, with_meta=True)
    SAV_META = utils.sav_meta_dict(_sav_meta)     # 결과를 .sav 로 되돌릴 때 쓴다

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
    if str(getattr(data_file, "name", "")).lower().endswith(".sav"):
        st.caption(
            "SPSS 파일이라 값 라벨을 `1) 서울` 형태로 읽었습니다. "
            "코드 번호가 앞에 붙어 있어야 값이 코드 순서대로 정렬됩니다. "
            "결과 엑셀에도 이 형태로 저장됩니다.")

    # ── 저장된 쿼터 설정 찾기 ────────────────────────────────────────────
    _slug = _preset_slug(getattr(data_file, "name", "data"))
    if st.session_state.get("QS_slug") != _slug:
        # 다른 파일이 올라오면 이전 프리셋 적용 상태를 버린다
        st.session_state["QS_slug"] = _slug
        st.session_state.pop("QS_preset", None)

    _found = preset_load_local(_slug)
    _applied = st.session_state.get("QS_preset")

    with st.expander("💾 쿼터 설정 저장 / 불러오기",
                     expanded=bool(_found and not _applied)):
        if _applied:
            st.success(
                f"✅ 저장된 설정을 적용했습니다 "
                f"({_applied.get('saved_at','')} 저장, "
                f"원본 `{_applied.get('source_file','')}`). "
                "아래 표에 이전 목표가 채워져 있습니다.")
            if st.button("↩️ 적용 취소하고 새로 설정", key="QS_clear"):
                st.session_state.pop("QS_preset", None)
                _reset_widget_state()
                st.rerun()
        elif _found:
            st.info(
                f"📋 `{_slug}` 이름으로 저장된 설정이 있습니다 "
                f"({_found.get('saved_at','')} 저장 · 메인 "
                f"{len(_found.get('main',{}).get('targets',{})):,}셀 · 추가 "
                f"{len(_found.get('extras',[])):,}개).")
            if st.button("📥 이 설정 불러오기", key="QS_load", type="primary"):
                st.session_state["QS_preset"] = _found
                _reset_widget_state()
                st.rerun()
        else:
            st.caption(
                f"`{_slug}` 이름으로 저장된 설정이 없습니다. "
                "쿼터를 입력한 뒤 아래에서 저장하면, 다음에 같은 파일명을 "
                "올렸을 때 자동으로 찾아 줍니다.")

        # 클라우드는 앱이 재시작되면 폴더가 비므로 JSON 으로도 주고받는다
        up = st.file_uploader("설정 파일(JSON) 불러오기", type=["json"],
                              key="QS_upload")
        if up is not None and st.button("📤 올린 설정 적용", key="QS_apply_up"):
            try:
                st.session_state["QS_preset"] = json.loads(
                    up.getvalue().decode("utf-8"))
                _reset_widget_state()
                st.rerun()
            except Exception as _je:                          # noqa: BLE001
                st.error(f"설정 파일을 읽지 못했습니다 — {type(_je).__name__}: {_je}")

        _saved = preset_list_local()
        if _saved:
            with st.popover(f"저장된 설정 {len(_saved)}개 보기"):
                st.dataframe(
                    pd.DataFrame(_saved, columns=["이름", "저장 시각", "원본 파일"]),
                    use_container_width=True, hide_index=True)

    st.divider()

    # ==========================================================================
    # 2. 쿼터 설정
    # ==========================================================================
    st.subheader("2. 쿼터 설정")
    use_main = st.checkbox("✅ 메인 쿼터 사용",
                           value=bool(preset_get("main.use_main", True)))
    main_map = {}
    algo_main_cols = []
    main_mode = 'grid'

    if use_main:
        _qm = ["엑셀 업로드", "화면 설계"]
        q_mode = st.radio(
            "메인 쿼터 방식", _qm, horizontal=True,
            index=_qm.index(preset_get("main.q_mode", "화면 설계"))
            if preset_get("main.q_mode") in _qm else 1)

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
            _cols = list(df_survey.columns)
            _p_rv = [c for c in (preset_get("main.rv") or []) if c in _cols]
            _p_cv = preset_get("main.cv")
            rv = st.multiselect("행(Row) 변수", _cols, default=_p_rv)
            cv = st.selectbox(
                "열(Col) 변수", ["(선택)"] + _cols,
                index=(_cols.index(_p_cv) + 1) if _p_cv in _cols else 0)
            if rv and cv != "(선택)":
                if cv in rv:
                    st.error("열 변수는 행 변수와 달라야 합니다.")
                else:
                    algo_main_cols = rv + [cv]
                    pi = cached_pivot(df_survey, tuple(rv), cv)
                    if pi.size > MAX_GRID_CELLS:
                        st.error(f"교차표 셀이 {pi.size:,}개로 너무 많습니다. 변수를 줄이세요.")
                    else:
                        pi_init = pi.reset_index()
                        _pm = preset_targets("main")
                        if _pm:
                            # 저장된 목표를 교차표 칸에 되돌려 놓는다.
                            # 이번 데이터에 없던 조합은 현재 분포를 그대로 둔다.
                            _hit = 0
                            for _r in range(len(pi_init)):
                                for _c in pi.columns:
                                    _k = tuple(
                                        [utils.norm_val(pi_init.iloc[_r][x]) for x in rv]
                                        + [utils.norm_val(_c)])
                                    if _k in _pm:
                                        pi_init.at[_r, _c] = int(_pm[_k])
                                        _hit += 1
                            st.caption(f"저장된 설정에서 {_hit:,}개 셀의 목표를 "
                                       f"불러왔습니다 (전체 {pi.size:,}칸).")
                        ed = st.data_editor(pi_init, use_container_width=True, disabled=rv)
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
        # 메인 쿼터를 쓰지 않을 때의 총 목표 인원.
        # key 를 주지 않으면 값이 리런마다 흔들릴 수 있어 고정 key 를 붙인다.
        _p_total = preset_get("main.total")
        try:
            _p_total = int(_p_total) if _p_total else 1000
        except (TypeError, ValueError):
            _p_total = 1000
        main_map = {('All',): st.number_input(
            "전체 목표", 1, 1000000, max(1, _p_total), key="QS_total_only",
            help="메인 쿼터를 쓰지 않을 때 뽑을 총 인원입니다. "
                 "설정을 저장하면 이 값도 함께 저장됩니다.")}
        algo_main_cols = []

    # --------------------------------------------------------------------------
    # 추가 쿼터
    # --------------------------------------------------------------------------
    ex_configs = []
    MAX_EXTRA = 8          # 추가 쿼터 그룹 최대 개수
    tabs = st.tabs([f"추가 {i+1}" for i in range(MAX_EXTRA)])

    for i, tab in enumerate(tabs):
        with tab:
            _em = ["단순형 (변수 값별 할당)", "조합형 (행/열 교차 할당)"]
            _p_ex = (preset_get("extras") or [])
            _p_this = _p_ex[i] if i < len(_p_ex) else {}
            _p_mode = (_p_this or {}).get("mode")
            ex_mode = st.radio(
                f"설정 방식 (그룹 {i+1})", _em, key=f"ex_mode_{i}", horizontal=True,
                index=1 if _p_mode == "grid" else 0
            )
            config = {'cols': [], 'map': {}, 'name': f"Extra_{i+1}", 'mode': 'simple'}

            if ex_mode.startswith("단순형"):
                config['mode'] = 'simple'
                _pc = [c for c in ((_p_this or {}).get("cols") or [])
                       if c in df_survey.columns] if _p_mode == "simple" else []
                cols = st.multiselect(f"변수 선택 (그룹 {i+1})", df_survey.columns,
                                      default=_pc, key=f"ms{i}")
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
                        _pm = preset_targets("extra", i) if _p_mode == "simple" else {}
                        if _pm:
                            tg, only_old, only_new = apply_preset_col(
                                cnt['값'].tolist(), _pm, cnt['현재'].tolist())
                            cnt['목표'] = tg
                            msg = (f"저장된 설정에서 목표 "
                                   f"{len(cnt) - len(only_new):,}개를 불러왔습니다.")
                            if only_new:
                                msg += f" 이번에 새로 생긴 값 {len(only_new)}개는 현재 분포로 채웠습니다: " \
                                       + ", ".join(map(str, only_new[:6]))
                            if only_old:
                                msg += f" / 저장된 설정에만 있던 값 {len(only_old)}개는 이번 데이터에 없습니다: " \
                                       + ", ".join(map(str, only_old[:6]))
                            (st.warning if (only_new or only_old) else st.caption)(msg)
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
                _gc = list((_p_this or {}).get("cols") or []) if _p_mode == "grid" else []
                _g_rv = [c for c in _gc[:-1] if c in df_survey.columns] if len(_gc) > 1 else []
                _g_cv = _gc[-1] if _gc else None
                _all = list(df_survey.columns)
                ex_rv = st.multiselect(f"행(Row) 변수 (그룹 {i+1})", _all,
                                       default=_g_rv, key=f"ex_rv_{i}")
                ex_cv = st.selectbox(
                    f"열(Col) 변수 (그룹 {i+1})", ["(선택)"] + _all,
                    index=(_all.index(_g_cv) + 1) if _g_cv in _all else 0,
                    key=f"ex_cv_{i}")

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
                            pi_init = pi.reset_index()
                            _pmg = preset_targets("extra", i) if _p_mode == "grid" else {}
                            if _pmg:
                                _hit = 0
                                for _r in range(len(pi_init)):
                                    for _c in pi.columns:
                                        _k = tuple(
                                            [utils.norm_val(pi_init.iloc[_r][x]) for x in ex_rv]
                                            + [utils.norm_val(_c)])
                                        if _k in _pmg:
                                            pi_init.at[_r, _c] = int(_pmg[_k])
                                            _hit += 1
                                st.caption(f"저장된 설정에서 {_hit:,}개 셀의 목표를 "
                                           f"불러왔습니다.")
                            ed = st.data_editor(pi_init, use_container_width=True,
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

        # 값이 범위를 벗어난 응답자를 후보에서 아예 빼는 필터
        iv_cap_on, iv_min, iv_max = False, None, None
        if use_intval and c_int:
            _iv = pd.to_numeric(
                pd.Series(df_survey[c_int]).replace('', np.nan), errors='coerce')
            _valid = _iv.dropna()
            iv_cap_on = st.checkbox(
                "값이 범위를 벗어난 응답자는 후보에서 제외", value=False,
                help="너무 낮거나 너무 높은 응답자를 아예 빼고 싶을 때 씁니다. "
                     "제외된 응답자는 어떤 셀에도 배정되지 않으므로 표본이 "
                     "그만큼 줄어듭니다.")
            if iv_cap_on:
                _lo = int(_valid.min()) if len(_valid) else 0
                _hi = int(_valid.max()) if len(_valid) else 1000
                iv_c1, iv_c2 = st.columns(2)
                iv_min = iv_c1.number_input(
                    "최소값 (이 값 미만 제외)", min_value=0,
                    max_value=max(_hi * 10, 1_000_000), value=_lo, step=1,
                    help="입력한 값보다 작은 응답자를 제외합니다. 같은 값은 남습니다.")
                iv_max = iv_c2.number_input(
                    "최대값 (이 값 초과 제외)", min_value=0,
                    max_value=max(_hi * 10, 1_000_000), value=_hi, step=1,
                    help="입력한 값보다 큰 응답자를 제외합니다. 같은 값은 남습니다.")
                if len(_valid):
                    st.caption(
                        f"`{c_int}` 최소 {int(_valid.min()):,} / 중앙값 "
                        f"{int(_valid.median()):,} / 최대 {int(_valid.max()):,} "
                        f"· 하위 1% {int(_valid.quantile(0.01)):,} "
                        f"· 상위 1% {int(_valid.quantile(0.99)):,}")
                if iv_min > iv_max:
                    st.error("최소값이 최대값보다 큽니다. 이대로면 전원 제외됩니다.")
                _lo_n = int((_iv < float(iv_min)).sum())
                _hi_n = int((_iv > float(iv_max)).sum())
                _n_over = _lo_n + _hi_n
                st.caption(
                    ("제외 대상 없음" if _n_over == 0 else
                     f"⚠️ 최소 미만 {_lo_n:,}명 + 최대 초과 {_hi_n:,}명 = "
                     f"{_n_over:,}명 제외 "
                     f"(남는 후보 {len(df_survey) - _n_over:,}명)"))
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

    # ── 현재 설정 저장 ───────────────────────────────────────────────────
    with st.expander("💾 지금 설정을 저장해 두기"):
        st.caption(
            f"파일명 `{_slug}` 로 저장됩니다. 다음에 같은 이름의 데이터를 올리면 "
            "자동으로 찾아서 목표를 채워 줍니다. 클라우드에서는 앱이 재시작되면 "
            "저장한 것이 사라지므로, JSON 파일도 함께 내려받아 두세요.")

        _main_state = {
            "use_main": bool(use_main),
            "q_mode": q_mode if use_main else None,
            "rv": rv if (use_main and q_mode == "화면 설계") else [],
            "cv": cv if (use_main and q_mode == "화면 설계" and cv != "(선택)") else None,
            "cols": list(algo_main_cols),
            "total": (int(list(main_map.values())[0])
                      if (not use_main and main_map) else None),
            "targets": {_enc_key(k): int(v) for k, v in main_map.items()},
        }
        _extras_state = []
        for _c in ex_configs:
            _extras_state.append({
                "name": _c['name'],
                "mode": _c['mode'],
                "cols": list(_c['cols']),
                "targets": {_enc_key(k): int(v) for k, v in _c['map'].items()},
            })
        _options = {
            "tol": int(tol), "use_intval": bool(use_intval),
            "c_int": str(c_int) if c_int else None,
            "c_no": str(c_no),
            "iv_cap_on": bool(iv_cap_on),
            "iv_min": int(iv_min) if iv_min is not None else None,
            "iv_max": int(iv_max) if iv_max is not None else None,
            "ex_as_target": bool(ex_as_target),
        }
        _payload = preset_payload(getattr(data_file, "name", "data"),
                                  _main_state, _extras_state, _options)

        _n_ex = sum(1 for c in ex_configs if c['cols'])
        st.write(f"저장될 내용 — 메인 {len(main_map):,}셀 / 추가 쿼터 {_n_ex}개")

        sc1, sc2 = st.columns(2)
        if sc1.button("💾 이 컴퓨터에 저장", use_container_width=True,
                      disabled=not main_map):
            ok, info = preset_save_local(_slug, _payload)
            if ok:
                st.success(f"저장했습니다 → `{info}`")
            else:
                st.warning(
                    f"폴더에 저장하지 못했습니다 ({info}). "
                    "클라우드에서는 정상이며, 옆의 JSON 다운로드를 쓰세요.")
        sc2.download_button(
            "⬇️ 설정 JSON 내려받기",
            json.dumps(_payload, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"쿼터설정_{_slug}.json", mime="application/json",
            use_container_width=True, disabled=not main_map)

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

                # --- intval 상한 초과 제외 ---
                #  메인 쿼터 목표에 없는 키를 부여하면 솔버와 근사 계산 양쪽에서
                #  후보 목록에 아예 오르지 않는다. (main_map.get(k, 0) <= 0 이면 제외)
                #  덕분에 quota_ilp / utils 를 건드리지 않고 처리할 수 있다.
                iv_over_idx = []
                if (use_intval and c_int and iv_cap_on
                        and iv_min is not None and iv_max is not None):
                    _ivr = pd.to_numeric(
                        pd.Series(df_survey[c_int]).replace('', np.nan),
                        errors='coerce')
                    _lo_m = (_ivr < float(iv_min)).to_numpy()
                    _hi_m = (_ivr > float(iv_max)).to_numpy()
                    _over = _lo_m | _hi_m
                    iv_over_idx = list(df_survey.index[_over])
                    if iv_over_idx:
                        _klen = len(m_keys[0]) if m_keys else 1
                        EXCLUDED = ("__intval_범위밖__",) * _klen
                        m_keys = [EXCLUDED if _over[i] else k
                                  for i, k in enumerate(m_keys)]
                        st.info(
                            f"ℹ️ `{c_int}` 값이 {int(iv_min):,}~{int(iv_max):,} "
                            f"범위를 벗어난 {len(iv_over_idx):,}명을 후보에서 "
                            f"제외했습니다 (최소 미만 {int(_lo_m.sum()):,}명, "
                            f"최대 초과 {int(_hi_m.sum()):,}명). 남은 후보 "
                            f"{len(df_survey) - len(iv_over_idx):,}명으로 계산합니다.")

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
            # 상한 초과로 애초에 후보가 아니었던 응답자는 사유를 남긴다.
            # (이들은 통과 대상이 될 수 없으므로 "통과" 를 덮어쓸 위험이 없다)
            if iv_over_idx:
                df_out.loc[iv_over_idx, 'Chk'] = "제외(intval 범위)"

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
                    {'항목': 'intval 범위 제외',
                     '값': (f"{int(iv_min):,}~{int(iv_max):,} 범위 밖 "
                            f"{len(iv_over_idx):,}명 제외"
                            if iv_over_idx else "사용 안 함")},
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

            dl1, dl2 = st.columns(2)
            dl1.download_button(
                "📥 엑셀로 받기" if not is_fail else "⚠️ 실패한 결과라도 받기 (엑셀)",
                out.getvalue(), "result.xlsx", type="primary",
                use_container_width=True
            )

            # ── SPSS 저장 ────────────────────────────────────────────────
            #  원본이 .sav 였다면 "1) 서울" 로 읽어들인 값을 코드(1)로 되돌리고
            #  값 라벨을 다시 입혀서 저장한다. 그래야 SPSS 에서 열었을 때
            #  문자열이 아니라 원래대로 코드+라벨이 된다.
            try:
                _sav_src = df_pass.drop(columns=["Chk"], errors="ignore")
                _sav_df, _miss = utils.sav_restore_codes(
                    _sav_src, SAV_META.get("value_labels"))
                _sav_bytes, _renamed, _warns = utils.write_sav_bytes(
                    _sav_df, SAV_META.get("value_labels"),
                    SAV_META.get("column_labels"))
                dl2.download_button(
                    "📥 SPSS(.sav) 로 받기 — 선정자만",
                    _sav_bytes, "result.sav", use_container_width=True,
                    help="통과한 응답자만 담깁니다. 분석 시트는 .sav 에 넣을 수 "
                         "없으므로 엑셀 파일을 함께 받아 두세요.")
                _notes = []
                if not SAV_META.get("value_labels"):
                    _notes.append(
                        "원본이 SPSS 파일이 아니라 값 라벨 정보가 없습니다. "
                        "값이 있는 그대로 저장됩니다.")
                if _miss:
                    _notes.append(
                        "코드로 되돌리지 못한 값이 있습니다 — "
                        + ", ".join(f"{k} {v:,}건" for k, v in list(_miss.items())[:5]))
                if _renamed:
                    _notes.append(
                        f"SPSS 변수명 규칙에 맞춰 {len(_renamed)}개 컬럼 이름을 "
                        "바꿨습니다: "
                        + ", ".join(f"{a}→{b}" for a, b in list(_renamed.items())[:5]))
                _notes += _warns
                if _notes:
                    with dl2.popover("SPSS 저장 참고사항"):
                        for _n in _notes:
                            st.caption("• " + _n)
            except ImportError:
                dl2.caption("SPSS 저장에는 `pyreadstat` 이 필요합니다.")
            except Exception as _se:                          # noqa: BLE001
                dl2.warning(f"SPSS 저장 실패 — {type(_se).__name__}: {_se}")

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

            # 추가 쿼터 달성 현황 — 그룹별로 따로 본다.
            #   그룹끼리 목표를 합산하면 안 된다. 단일응답이면 응답자 한 명이
            #   그룹마다 1명씩 계상되므로, 3개 그룹이면 합계가 전체 목표의 3배가
            #   되어 "목표 3,900명"처럼 실제와 동떨어진 숫자가 나온다.
            if ex_as_target and ex_configs:
                g_rows = []
                for j, cfg in enumerate(ex_configs):
                    if not cfg['cols'] or not cfg['map']:
                        continue
                    tgt = sum(cfg['map'].values())
                    sh = sum(max(0, v - final_exs[j].get(k, 0))
                             for k, v in cfg['map'].items())
                    ov = sum(max(0, final_exs[j].get(k, 0) - v)
                             for k, v in cfg['map'].items())
                    n_bad_items = sum(1 for k, v in cfg['map'].items()
                                      if final_exs[j].get(k, 0) != v)
                    if sh == 0 and ov == 0:
                        state = "✅ 충족"
                    elif sh == 0:
                        state = f"↗ {ov:,}명 초과"
                    else:
                        state = f"⚠️ {sh:,}명 부족"
                    g_rows.append({
                        "추가 쿼터": cfg['name'],
                        "항목 수": len(cfg['map']),
                        "목표": tgt,
                        "달성": tgt - sh,
                        "부족": sh,
                        "초과": ov,
                        "어긋난 항목": n_bad_items,
                        "상태": state,
                    })
                if g_rows:
                    n_ok = sum(1 for r in g_rows if r["상태"] == "✅ 충족")
                    st.markdown(
                        f"**추가 쿼터 달성 현황 — {len(g_rows)}개 중 {n_ok}개 충족**")
                    st.dataframe(
                        pd.DataFrame(g_rows), use_container_width=True,
                        hide_index=True,
                        column_config={
                            "목표": st.column_config.NumberColumn(format="%d"),
                            "달성": st.column_config.NumberColumn(format="%d"),
                            "부족": st.column_config.NumberColumn(format="%d"),
                            "초과": st.column_config.NumberColumn(format="%d"),
                            "어긋난 항목": st.column_config.NumberColumn(
                                format="%d",
                                help="목표와 정확히 일치하지 않는 항목(코드) 수"),
                        })
                    st.caption(
                        "추가 쿼터는 그룹마다 따로 판단합니다. 단일응답 변수라면 "
                        "각 그룹의 목표 합계가 전체 목표 인원과 같아야 정상입니다 "
                        "(응답자 한 명이 그룹마다 한 번씩 계상되기 때문입니다).")

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
                                '편차 크기': abs(act - tgt),
                                # NumberColumn 의 %% 서식은 0~100 스케일을 쓴다
                                '편차율': (100.0 * (act - tgt) / tgt) if tgt else None})
                if ex_dev_recs:
                    mx = max(abs(r['편차']) for r in ex_dev_recs)
                    mxr = max((abs(r['편차율']) for r in ex_dev_recs
                               if r['편차율'] is not None), default=0.0) / 100.0
                    d1, d2, d3 = st.columns(3)
                    d1.metric("📐 편차 발생 항목", f"{len(ex_dev_recs)}개")
                    d2.metric("최대 편차", f"{mx:,}명")
                    d3.metric("최대 편차율", f"{mxr:.1%}")
                    st.caption(
                        "총 선정 인원은 메인 쿼터가 정하므로 그대로이고, 추가 쿼터의 "
                        "개별 항목만 목표 위아래로 나뉘어 흔들립니다. 아래 편차는 "
                        "이 조건에서 가능한 최소값입니다.")
                    # [수정] Styler.background_gradient 는 matplotlib 가 없으면
                    # ImportError 로 죽는다. matplotlib 은 pandas 필수 의존성이
                    # 아니므로, 패키지를 늘리지 않고 Streamlit 기본 기능으로 표현한다.
                    #   - 편차 크기는 ProgressColumn 막대로 (색 대신 길이)
                    #   - 편차율은 NumberColumn 서식으로
                    dfd = pd.DataFrame(ex_dev_recs)
                    st.dataframe(
                        dfd, use_container_width=True, hide_index=True,
                        column_config={
                            '목표': st.column_config.NumberColumn(format="%d"),
                            '실제': st.column_config.NumberColumn(format="%d"),
                            '편차': st.column_config.NumberColumn(
                                "편차(명)", format="%+d",
                                help="양수는 목표보다 많이, 음수는 적게 들어온 인원"),
                            '편차 크기': st.column_config.ProgressColumn(
                                "편차 크기", format="%d명", min_value=0,
                                max_value=int(mx) if mx else 1),
                            '편차율': st.column_config.NumberColumn(
                                format="%+.1f%%",
                                help="목표 대비 벗어난 비율"),
                        })

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

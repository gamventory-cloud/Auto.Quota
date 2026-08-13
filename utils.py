"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : utils.py                                                        ║
║  위치   : 리포지토리 최상단  (pages/ 폴더 안이 아님!)                       ║
║                                                                          ║
║  이 파일은 공용 함수 모듈입니다. 화면(UI) 코드가 전혀 없습니다.              ║
║  파일을 열었을 때 페이지 설정·파일 업로더·"매칭" 실행 버튼 같은              ║
║  화면 코드가 보이면, 그건 이 파일이 아니라                                  ║
║  pages/2___쿼터_솔루션.py 의 내용입니다.                                    ║
╚══════════════════════════════════════════════════════════════════════════╝

utils.py — 쿼터 솔루션 공용 모듈

주요 변경점
-----------
1. 정규화 함수 단일화 : norm_val / norm_series 하나만 쓴다.
   - 기존 clean_val / clean_series / normalize_val 3종 불일치 제거
   - 결측치는 NA_TOKEN 하나로 통일 ("NaN" vs "nan" 불일치 제거)
   - strip() 을 먼저, .0 제거를 나중에 (순서 버그 수정)
   - "1.5" -> "1" 로 뭉개던 split('.')[0] 방식 폐기
2. 키 생성 함수 신설 : build_simple_keys / build_tuple_keys / build_grid_keys
   - 설정 화면과 실행 시점이 반드시 같은 함수를 호출하도록 강제
   - 단순형 중복 카운트 버그 수정 (한 명이 같은 값으로 2칸 소모하던 문제)
3. simulation_worker 재작성
   - np.random.default_rng 로 스레드별 독립 RNG (전역 seed 경합 제거)
   - 상대(곱셈) 지터로 데이터 스케일과 무관한 탐색 강도
   - 목표 도달 시 즉시 break (내부 루프 낭비 제거)
   - 목표 0인 메인 키의 행은 정렬 전에 사전 제외
4. transform_pivoted_quota : bare except 제거, 실패 원인을 예외로 전달
5. check_password : hmac 비교, secrets 누락 대응, 첫 진입 시 오류 미표시
6. unique_sheet_name : 시트명 충돌 방지
7. [2.1] long_to_wide / find_duplicate_cells / AGG_FUNCS 추가
   - pages/7___세로_가로_변환.py 에서 사용
   - key_col 정규화는 norm_val 을 재사용하므로 앱 전체와 열 이름이 일관됨
   - 기존 함수는 하나도 수정하지 않았다 (섹션 6만 추가)
"""

import streamlit as st
import pandas as pd
import chardet
import io
import re
import hmac
import numpy as np
import collections

# 이 파일이 진짜 utils.py 인지 호출부에서 확인하는 표식.
# 파일 내용이 뒤섞이는 사고를 즉시 잡아낸다.
MODULE_ROLE = "utils"
__version__ = "2.1-longwide"

# 결측/공백을 나타내는 단일 토큰. 화면·엑셀·매칭 전부 이 값을 공유한다.
NA_TOKEN = "(무응답)"

# 쿼터 맵에 없는 키에 부여하는 희소성 페널티
MISS_PENALTY = 999.0


# ==============================================================================
# 0. 정규화 : 단일 진실 공급원 (Single Source of Truth)
# ==============================================================================
def norm_val(v):
    """
    임의의 값을 쿼터 키로 쓸 수 있는 정규화된 문자열로 변환한다.

    규칙 (순서 중요):
      1) 결측 / 빈 문자열  -> NA_TOKEN
      2) 앞뒤 공백 제거
      3) 정수를 float 로 읽어온 경우에만 뒤의 ".0" 제거  (1.0 -> "1")
         - "1.5" 는 그대로 "1.5"  (기존 clean_val 은 "1" 로 뭉갰음)
         - "서울.강남" 도 그대로  (기존 clean_val 은 "서울" 로 잘랐음)
    """
    if v is None:
        return NA_TOKEN
    try:
        if pd.isna(v):
            return NA_TOKEN
    except (TypeError, ValueError):
        # 배열/리스트 등 isna 가 스칼라를 반환하지 않는 타입은 그냥 문자열화
        pass

    s = str(v).strip()
    if s == "":
        return NA_TOKEN
    if s.endswith(".0") and s[:-2].lstrip("+-").isdigit():
        s = s[:-2]
    return s


def norm_series(s):
    """
    시리즈 정규화. norm_val 을 그대로 map 하므로 스칼라 경로와 100% 동일하다.
    (성능을 위해 벡터화하고 싶더라도, 결과가 norm_val 과 어긋나면 안 된다.)
    """
    return pd.Series(s).map(norm_val)


# --- 하위 호환용. 신규 코드에서는 쓰지 말 것 ------------------------------------
def clean_val(v):
    """[DEPRECATED] 다른 페이지 호환용. 쿼터 로직에서는 norm_val 을 쓸 것."""
    if pd.isna(v):
        return "NaN"
    return str(v).strip().split('.')[0]


def collect_values_from_cols(row, columns):
    """[DEPRECATED] build_simple_keys 로 대체됨."""
    values = {}
    for c in columns:
        s = norm_val(row[c])
        if s != NA_TOKEN:
            values[s] = None
    return list(values)


# ==============================================================================
# 1. 쿼터 키 생성 : 설정 화면과 실행 시점이 반드시 이 함수들을 공유해야 한다
# ==============================================================================
def build_simple_keys(df, cols):
    """
    단순형(값별 할당) 키.
    행마다 [정규화된 값, ...] 리스트를 반환한다.

    - 결측/공백은 제외한다 (해당 행은 그 그룹의 제약을 받지 않음)
    - 같은 행 안에서 값이 중복되면 1회만 센다
      (다중응답 1순위/2순위가 같은 값일 때 쿼터를 2칸 먹던 버그 수정)
    """
    out = []
    for row in df[cols].itertuples(index=False, name=None):
        seen = {}
        for v in row:
            s = norm_val(v)
            if s != NA_TOKEN:
                seen[s] = None
        out.append(list(seen))
    return out


def build_tuple_keys(df, cols):
    """조합형/메인 쿼터 키. 행마다 (값1, 값2, ...) 튜플을 반환한다."""
    cols_normed = [norm_series(df[c]).to_numpy() for c in cols]
    return list(zip(*cols_normed))


def build_grid_keys(df, cols):
    """조합형 추가 쿼터 키. 행마다 [튜플] (길이 1 리스트)."""
    return [[t] for t in build_tuple_keys(df, cols)]


def build_tiebreak(df, col):
    """
    intval 타이브레이크용 배열을 만든다.

    "쿼터 조건이 같으면 값이 낮은 응답자를 먼저 제외" 규칙이므로
    **값이 큰 쪽을 먼저 선택**하는 실수 배열을 반환한다.
    숫자로 변환할 수 없는 값과 결측은 -inf 로 두어 가장 먼저 탈락시킨다.

    반환: (배열, 유효 개수, 무효/결측 개수)
    """
    s = pd.to_numeric(pd.Series(df[col]).replace('', np.nan), errors='coerce')
    bad = int(s.isna().sum())
    arr = s.fillna(-np.inf).to_numpy(dtype=float)
    return arr, int(len(s) - bad), bad


# ==============================================================================
# 2. 텍스트 / 파일 유틸
# ==============================================================================
def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).strip()
    return text.replace("\n", "").replace("\r", "").replace("\t", "")


def extract_base_name(text):
    text = clean_text(text)
    if "." in text:
        return text.split(".")[0].strip()
    return text.strip()


def sanitize_var_name(text):
    text = str(text)
    text = text.replace("-", "_").replace(" ", "_")
    text = re.sub(r"[^a-zA-Z0-9_]", "", text)
    text = re.sub(r"__+", "_", text)
    return text


def natural_key(string_):
    """'q10' 이 'q9' 뒤에 오도록 정렬. 항상 [str, int, str, ...] 교대라 타입 안전."""
    target = str(string_)
    return [int(s) if s.isdigit() else s.lower() for s in re.split(r'(\d+)', target)]


def load_df(file):
    """
    실패 시 None 을 반환한다. 호출부는 반드시 None 을 검사할 것.
      df = utils.load_df(f)
      if df is None: st.stop()
    """
    if file is None:
        return None
    try:
        if file.name.lower().endswith('.csv'):
            raw = file.read()
            enc = chardet.detect(raw)['encoding'] or 'utf-8'
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc)
            except UnicodeDecodeError:
                # chardet 오탐 대비 국내 인코딩 폴백
                for fb in ('utf-8-sig', 'cp949', 'euc-kr'):
                    try:
                        return pd.read_csv(io.BytesIO(raw), encoding=fb)
                    except UnicodeDecodeError:
                        continue
                raise
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"파일 로드 실패: {type(e).__name__}: {e}")
        return None


RESERVED_SHEETS = {'Result_All', 'Result_Pass', 'Shortage_Analysis', 'Main_Status'}


def sanitize_sheet_name(name):
    safe_name = re.sub(r'[\\/*?:\[\]]', '_', str(name)).strip("'")
    if not safe_name:
        safe_name = "Sheet"
    if len(safe_name) > 30:
        return safe_name[:28] + ".."
    return safe_name


def unique_sheet_name(name, used):
    """
    엑셀 시트명 충돌 방지. used 는 호출자가 유지하는 set 이며 이 함수가 갱신한다.
    (추가 쿼터 두 그룹이 같은 컬럼 조합일 때 xlsxwriter 가 죽던 문제 수정)
    """
    base = sanitize_sheet_name(name)
    cand, i = base, 2
    while cand in used or cand in RESERVED_SHEETS:
        suffix = f"_{i}"
        cand = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(cand)
    return cand


# ==============================================================================
# 3. 쿼터 엑셀 파싱
# ==============================================================================
def transform_pivoted_quota(df_raw):
    """
    피벗 형태의 쿼터 시트를 평면 테이블(qt1, qt2, qt3, target)로 변환한다.

    기대 레이아웃 (header=None 으로 읽은 상태)
      - 2번째 행(iloc[1]) C열 이후 : qt3 라벨
      - 3번째 행(iloc[2]) 이후     : 데이터
      - A열 = qt1 (병합셀 대응 ffill), B열 = qt2

    실패 시 ValueError 를 raise 한다. (기존의 `except: return None` 은
    원인을 완전히 삼켜서 "엑셀 오류" 다섯 글자만 남겼다.)
    """
    if df_raw is None or df_raw.empty:
        raise ValueError("쿼터 시트가 비어 있습니다.")
    if df_raw.shape[0] < 3 or df_raw.shape[1] < 3:
        raise ValueError(
            f"쿼터 시트 크기가 부족합니다 (행 {df_raw.shape[0]}, 열 {df_raw.shape[1]}). "
            "2행 C열부터 qt3 라벨, 3행부터 데이터, A/B열이 qt1/qt2 여야 합니다."
        )

    qt3_labels = [norm_val(x) for x in df_raw.iloc[1, 2:].dropna().values]
    if not qt3_labels:
        raise ValueError("2행 C열 이후에서 qt3 라벨을 찾지 못했습니다.")

    dups = [k for k, v in collections.Counter(qt3_labels).items() if v > 1]
    if dups:
        raise ValueError(f"qt3 라벨이 중복됩니다: {dups}")

    n = len(qt3_labels)
    data_rows = df_raw.iloc[2:, :2 + n].copy()
    data_rows.iloc[:, 0] = data_rows.iloc[:, 0].ffill()
    data_rows.columns = ['qt1', 'qt2'] + qt3_labels

    flat = data_rows.melt(id_vars=['qt1', 'qt2'], var_name='qt3', value_name='target')
    for col in ('qt1', 'qt2', 'qt3'):
        flat[col] = flat[col].map(norm_val)
    flat['target'] = pd.to_numeric(flat['target'], errors='coerce').fillna(0).astype(int)

    flat = flat[flat['target'] > 0].reset_index(drop=True)
    if flat.empty:
        raise ValueError("목표 인원이 1 이상인 셀이 하나도 없습니다.")
    return flat


# ==============================================================================
# 4. 비밀번호
# ==============================================================================
def check_password():
    """올바른 비밀번호가 입력되었으면 True."""
    def password_entered():
        try:
            secret = st.secrets["password"]
        except Exception:
            st.session_state["password_correct"] = False
            st.session_state["password_msg"] = (
                "서버에 비밀번호가 설정되어 있지 않습니다. "
                ".streamlit/secrets.toml 의 password 항목을 확인하세요."
            )
            return

        entered = str(st.session_state.get("password", ""))
        if hmac.compare_digest(entered, str(secret)):
            st.session_state["password_correct"] = True
            st.session_state["password_msg"] = None
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
            st.session_state["password_msg"] = "비밀번호가 올바르지 않습니다."

    if st.session_state.get("password_correct", False):
        return True

    st.title("🔒 접속 제한")
    st.text_input("비밀번호를 입력하세요", type="password",
                  on_change=password_entered, key="password")
    msg = st.session_state.get("password_msg")
    if msg:
        st.error(msg)
    else:
        st.caption("지인들만 사용 가능한 비공개 프로그램입니다.")
    return False


# ==============================================================================
# 5. 시뮬레이션 워커
# ==============================================================================
def simulation_worker(seed, num_iters, indices, scarcity_scores, m_keys, ex_keys_list,
                      main_map, ex_maps, soft_target, target_total=None, jitter=0.15,
                      tiebreak=None):
    """
    희소성 점수 순으로 응답자를 그리디 선택하되, 매 반복마다 점수에 지터를 주어
    서로 다른 해를 탐색한다. 최선의 (선택 인원수, 인덱스 리스트) 를 반환.

    tiebreak : 행별 실수 배열. 쿼터 조건이 완전히 같은 응답자들 사이에서
               **값이 큰 쪽을 먼저 선택**한다 (= 값이 작은 쪽이 먼저 탈락).
               결측은 -inf 로 넣어 가장 먼저 탈락시킨다.
               쿼터 조건이 다른 응답자끼리는 영향을 주지 않으므로
               최종 인원수는 이 값과 무관하다.

    변경점
      - np.random.default_rng(seed) : 스레드마다 독립 RNG.
        기존 np.random.seed() 는 전역 상태라 threading 백엔드에서 서로 덮어썼다.
      - 곱셈 지터 : scarcity_scores 의 절대 스케일(추가 쿼터 수, 999 페널티)에
        관계없이 일정한 탐색 강도를 유지한다.
        기존 `+ uniform(0, 0.5)` 는 점수가 수백대면 순서를 전혀 못 바꿨다.
      - 지터를 행 단위가 아니라 **프로파일(동일 쿼터 조건) 단위**로 뽑는다.
        같은 조건 응답자는 서로 교환 가능하므로 그들 사이를 무작위로 섞는 것은
        탐색에 아무 도움이 안 되고, tiebreak 순서만 망가뜨린다.
      - 목표 인원 도달 시 내부 루프 즉시 종료.
      - 목표가 0인 메인 키에 속한 행은 정렬 대상에서 사전 제외.
    """
    rng = np.random.default_rng(seed)
    if target_total is None:
        target_total = sum(main_map.values())

    # 실제로 제약이 걸린 추가 쿼터 그룹만 남긴다
    active = [(j, m) for j, m in enumerate(ex_maps) if m]

    # --- 사전 필터 : 메인 쿼터 목표가 0인 행은 어차피 못 뽑는다 ---
    elig = np.fromiter((main_map.get(k, 0) > 0 for k in m_keys),
                       dtype=bool, count=len(m_keys))
    pos = np.flatnonzero(elig)
    if pos.size == 0:
        return 0, []

    base_scores = np.asarray(scarcity_scores, dtype=float)[pos]
    m_keys_e = [m_keys[i] for i in pos]
    indices_e = np.asarray(indices)[pos]
    ex_keys_e = [[ex_keys_list[j][i] for i in pos] for j, _ in active]
    n = pos.size

    tb = None
    if tiebreak is not None:
        tb = np.asarray(tiebreak, dtype=float)[pos]

    # --- 프로파일 id : 쿼터 조건이 완전히 같은 행끼리 같은 번호 ---
    # 지터를 프로파일 단위로 주고, 프로파일 내부는 tiebreak 순서를 유지한다.
    sig_to_id, prof = {}, np.empty(n, dtype=np.int64)
    for a in range(n):
        sig = (m_keys_e[a],
               tuple(tuple(ex_keys_e[g][a]) for g in range(len(active))))
        prof[a] = sig_to_id.setdefault(sig, len(sig_to_id))
    n_prof = len(sig_to_id)

    best_cnt, best_idxs = 0, []

    for _ in range(num_iters):
        scores = base_scores * rng.uniform(1.0 - jitter, 1.0 + jitter, size=n_prof)[prof]
        if tb is None:
            order = np.argsort(scores, kind='stable')
        else:
            # lexsort 는 마지막 키가 1차 기준. 점수 오름차순 → tiebreak 내림차순
            order = np.lexsort((-tb, scores))

        m_cnt = collections.defaultdict(int)
        ex_cnts = [collections.defaultdict(int) for _ in active]
        chosen = []
        cnt = 0

        for p in order:
            mk = m_keys_e[p]
            if m_cnt[mk] >= main_map[mk]:
                continue

            ok = True
            for a, (_, e_map) in enumerate(active):
                for k in ex_keys_e[a][p]:
                    cap = e_map.get(k)
                    if cap is not None and ex_cnts[a][k] >= cap:
                        ok = False
                        break
                if not ok:
                    break
            if not ok:
                continue

            m_cnt[mk] += 1
            for a, _ in enumerate(active):
                ec = ex_cnts[a]
                for k in ex_keys_e[a][p]:
                    ec[k] += 1

            chosen.append(indices_e[p])
            cnt += 1
            if cnt >= target_total:      # 더 뽑을 자리가 없다
                break

        if cnt > best_cnt:
            best_cnt = cnt
            best_idxs = list(chosen)
            if best_cnt >= soft_target:
                break

    return best_cnt, best_idxs


# ==============================================================================
# 6. Long → Wide 변환 (세로로 쌓인 데이터를 가로로 펼치기)
# ==============================================================================
# 중복값(같은 ID + 같은 변수가 2회 이상) 처리 방법.
# 화면의 selectbox 와 이 dict 하나만 공유하도록 한다.
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
    pivot 이 int 열을 float 로 승격시켜 '170.0' 처럼 보이는 문제를 없앤다.
    Int64(nullable) 를 쓰므로 결측은 <NA> 로 남고 엑셀에서 빈 칸이 된다.
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
    세로(long) 데이터를 가로(wide)로 펼친다.

        ID  항목    값                ID  키   몸무게
        1   키      170     ──▶       1   170  65
        1   몸무게  65                2   160  50
        2   키      160
        2   몸무게  50

    파라미터
    --------
    id_cols        : 한 행을 식별하는 열 목록 (예: ["ID"], ["학번", "이름"])
    key_col        : 가로로 펼칠 변수명이 담긴 열 (예: "항목", "시점")
    value_cols     : 실제 값이 담긴 열 목록. 2개 이상이면 열 이름이
                     "값열{name_sep}변수명" 으로 조합된다 (점수_1차, 시간_1차 …)
    keep_keys      : 펼칠 변수 목록. None 이면 전체.
                     여기 준 순서가 결과 열 순서가 된다.
    aggfunc        : 중복 발생 시 집계 방법. AGG_FUNCS 의 값을 넘긴다.
    normalize_keys : key_col 을 norm_val 로 정규화한다. 1.0 -> "1",
                     공백 -> NA_TOKEN 이 되어 열 이름이 앱 전체와 일관해진다.

    실패 시 ValueError 를 raise 한다. 호출부에서 st.error 로 보여줄 것.
    """
    if df is None or df.empty:
        raise ValueError("데이터가 비어 있습니다.")

    id_cols = list(id_cols)
    value_cols = list(value_cols)

    if not id_cols:
        raise ValueError("ID 열을 1개 이상 선택해야 합니다.")
    if not value_cols:
        raise ValueError("값 열을 1개 이상 선택해야 합니다.")
    if not key_col:
        raise ValueError("기준 열을 선택해야 합니다.")

    overlap = (set(id_cols) | set(value_cols)) & {key_col}
    if overlap:
        raise ValueError(f"기준 열 '{key_col}' 이 ID 열 또는 값 열과 겹칩니다.")
    both = set(id_cols) & set(value_cols)
    if both:
        raise ValueError(f"ID 열과 값 열에 같은 열이 들어갔습니다: {sorted(both)}")

    missing = [c for c in id_cols + value_cols + [key_col] if c not in df.columns]
    if missing:
        raise ValueError(f"데이터에 없는 열입니다: {missing}")

    work = df.loc[:, id_cols + [key_col] + value_cols].copy()
    # 열 이름이 될 값이므로 항상 문자열로 통일한다. 정규화를 끄더라도
    # astype("string") 을 거치지 않으면 int 1 과 화면에서 고른 "1" 이 어긋난다.
    if normalize_keys:
        work[key_col] = norm_series(work[key_col])
    else:
        work[key_col] = work[key_col].astype("string")

    # 펼칠 변수 선택
    if keep_keys is None:
        keep_keys = list(pd.unique(work[key_col].dropna()))
    else:
        keep_keys = [norm_val(k) if normalize_keys else str(k) for k in keep_keys]
        # 순서 유지 중복 제거
        keep_keys = list(dict.fromkeys(keep_keys))
        work = work[work[key_col].isin(keep_keys)]
        if work.empty:
            raise ValueError("선택한 변수에 해당하는 행이 없습니다.")

    # ID 조합의 원래 등장 순서 보존 (pivot 은 정렬해 버린다)
    order = df.loc[:, id_cols].drop_duplicates().reset_index(drop=True)

    try:
        wide = pd.pivot_table(
            work,
            index=id_cols,
            columns=key_col,
            values=value_cols,
            aggfunc=aggfunc,
            dropna=True,      # False 로 두면 정수가 불필요하게 실수로 승격된다
            observed=True,
        )
    except Exception as e:
        raise ValueError(
            f"피벗 실패 ({type(e).__name__}: {e})\n"
            "값 열에 숫자가 아닌 값이 섞여 있는데 평균/합계를 고른 경우일 수 있습니다. "
            "중복값 처리를 '첫 번째 값만' 으로 바꿔 보세요."
        )

    # 열 이름 정리
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

    # ID 원래 순서 복원
    wide = order.merge(wide, on=id_cols, how="left")
    return _restore_ints(wide)


def find_duplicate_cells(df, id_cols, key_col, value_cols=None,
                         normalize_keys=True, limit=200):
    """
    long_to_wide 실행 전에 '같은 ID + 같은 변수' 가 2회 이상 있는지 미리 찾는다.
    중복이 있으면 집계 방법에 따라 값이 조용히 바뀌므로, 화면에서 먼저 경고한다.

    반환: (중복 건수 DataFrame, 영향받은 행 총 개수)
    """
    key = norm_series(df[key_col]) if normalize_keys else df[key_col].astype("string")
    tmp = df.loc[:, id_cols].copy()
    tmp[key_col] = key
    g = tmp.groupby(id_cols + [key_col], dropna=False, observed=True).size()
    dup = g[g > 1].reset_index(name="중복 횟수")
    total = int(dup["중복 횟수"].sum()) if not dup.empty else 0
    return dup.head(limit), total

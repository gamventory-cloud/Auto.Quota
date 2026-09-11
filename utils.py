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
7. [2.0.1] check_password 의 비ASCII 크래시 수정
   - 비밀번호 칸에 한글을 입력하면 hmac.compare_digest 가 TypeError 를 내며
     앱 전체가 죽었다. 양쪽을 utf-8 bytes 로 인코딩해 비교하도록 변경.
   - 이 파일에서 바뀐 것은 위 한 곳뿐이며, 다른 함수는 손대지 않았다.
6. unique_sheet_name : 시트명 충돌 방지
10. [2.0.4] write_sav_bytes 의 '문자로 저장' 오판 수정
   - 값이 전부 결측인 열을 문자열로 잘못 판정해 헛경고가 떴다.
     실제로는 원본이 숫자형이면 그대로 숫자로 저장된다.
   - 이미 숫자형이거나 전부 결측이면 손대지 않고 값 라벨도 유지한다.
   - 일부만 숫자인 열은 나머지가 결측이 되므로 그 건수를 알려준다.
9. [2.0.3] SPSS .sav 읽기 추가 (read_sav_combined)
   - load_df 가 .sav 를 받으면 값 라벨을 "1) 서울" 형태로 합쳐서 돌려준다.
     라벨만 쓰면 화면 정렬이 가나다순이 되어 코드 순서와 어긋나기 때문이다.
   - mode="label" / "code" 로 다른 방식도 쓸 수 있다.
   - pyreadstat 은 파일 경로만 받으므로 임시 폴더를 거친다.
8. [2.0.2] RESERVED_SHEETS 에 Run_Info / Recruit_Plan 추가
   - 화면 쪽에서 재현성 기록 시트(Run_Info)와 추가 수집 지시서 시트
     (Recruit_Plan)를 새로 쓴다. 추가 쿼터 그룹 이름이 우연히 이 둘과 같아지면
     xlsxwriter 가 시트명 충돌로 죽으므로 예약어에 등록한다.
   - 이 파일에서 바뀐 것은 RESERVED_SHEETS 한 줄뿐이며, 함수는 손대지 않았다.
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
__version__ = "2.0.4-savfix"

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


def read_sav_combined(file, mode="combined"):
    """
    SPSS .sav 를 읽는다.

    mode
      "combined" (기본) : 값 라벨이 있는 변수를 "1) 서울" 형태로 만든다.
          라벨만 쓰면(=apply_value_formats) 화면 정렬이 가나다순이 되어
          코드 순서와 어긋난다. 예: 1서울 2부산 3대구 -> 광주,대구,부산,서울...
          앞에 코드를 붙이면 natural_key 가 숫자를 먼저 보므로 코드 순서가 유지되고
          라벨도 그대로 읽힌다.
      "label" : 라벨만 ("서울")
      "code"  : 코드만 (1, 2, 3)

    pyreadstat 은 파일 경로만 받으므로 업로드된 바이트를 임시 폴더에 풀어서 읽는다.
    (Windows 에서는 NamedTemporaryFile 을 열어둔 채 다시 열 수 없어 폴더를 쓴다)
    """
    import tempfile
    import os as _os
    try:
        import pyreadstat
    except ImportError as e:
        raise ImportError("SPSS(.sav) 를 읽으려면 pyreadstat 이 필요합니다. "
                          "pip install pyreadstat") from e

    raw = file.read() if hasattr(file, "read") else open(file, "rb").read()
    with tempfile.TemporaryDirectory() as td:
        path = _os.path.join(td, "upload.sav")
        with open(path, "wb") as f:
            f.write(raw)
        if mode == "label":
            df, meta = pyreadstat.read_sav(path, apply_value_formats=True)
            return df, meta
        df, meta = pyreadstat.read_sav(path, apply_value_formats=False)

    if mode == "code":
        return df, meta

    vlabels = getattr(meta, "variable_value_labels", None) or {}
    for col, vmap in vlabels.items():
        if col not in df.columns:
            continue
        conv = {}
        for code, lab in vmap.items():
            c = int(code) if isinstance(code, float) and float(code).is_integer() else code
            conv[code] = f"{c}) {lab}"
        df[col] = df[col].map(lambda v: conv.get(v, v))
    return df, meta


def load_df(file, with_meta=False):
    """
    실패 시 None 을 반환한다. 호출부는 반드시 None 을 검사할 것.
      df = utils.load_df(f)
      if df is None: st.stop()

    .sav 는 값 라벨을 "1) 서울" 형태로 합쳐서 읽는다 (read_sav_combined 참고).

    with_meta=True 면 (df, meta) 를 돌려준다. .sav 가 아니면 meta 는 None.
    결과를 다시 .sav 로 내보낼 때 값 라벨을 복원하려면 meta 가 필요하다.
    """
    if file is None:
        return (None, None) if with_meta else None
    try:
        if str(getattr(file, "name", file)).lower().endswith('.sav'):
            df, _meta = read_sav_combined(file, mode="combined")
            return (df, _meta) if with_meta else df
        if file.name.lower().endswith('.csv'):
            raw = file.read()
            enc = chardet.detect(raw)['encoding'] or 'utf-8'
            out = None
            try:
                out = pd.read_csv(io.BytesIO(raw), encoding=enc)
            except UnicodeDecodeError:
                # chardet 오탐 대비 국내 인코딩 폴백
                for fb in ('utf-8-sig', 'cp949', 'euc-kr'):
                    try:
                        out = pd.read_csv(io.BytesIO(raw), encoding=fb)
                        break
                    except UnicodeDecodeError:
                        continue
                if out is None:
                    raise
            return (out, None) if with_meta else out
        out = pd.read_excel(file)
        return (out, None) if with_meta else out
    except Exception as e:
        st.error(f"파일 로드 실패: {type(e).__name__}: {e}")
        return (None, None) if with_meta else None


# Run_Info      : 실행 설정·시각을 남기는 재현성 기록 시트
# Recruit_Plan  : 미달 시 "무엇을 몇 명 더 수집해야 하는지" 추가 수집 지시서 시트
RESERVED_SHEETS = {'Result_All', 'Result_Pass', 'Shortage_Analysis', 'Main_Status',
                   'Run_Info', 'Recruit_Plan'}


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

        # compare_digest 는 비ASCII 문자열을 그대로 넘기면
        #   TypeError: comparing strings with non-ASCII characters is not supported
        # 를 낸다. 사용자가 비밀번호 칸에 한글을 입력하면 "틀렸습니다" 가 아니라
        # 앱이 그대로 죽었다. 양쪽 모두 bytes 로 인코딩해서 비교한다.
        entered = str(st.session_state.get("password", "")).encode("utf-8")
        if hmac.compare_digest(entered, str(secret).encode("utf-8")):
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


def sav_meta_dict(meta):
    """write_sav 에 필요한 정보만 추려 담는다. meta 가 없으면 빈 dict."""
    if meta is None:
        return {}
    return {
        "value_labels": dict(getattr(meta, "variable_value_labels", None) or {}),
        "column_labels": dict(getattr(meta, "column_names_to_labels", None) or {}),
    }


def sav_restore_codes(df, value_labels):
    """
    "1) 서울" 형태로 읽어들인 값을 원래 코드(1)로 되돌린다.

    read_sav_combined 가 만든 문자열을 정확히 역으로 매핑한다.
    되돌릴 수 없는 값(직접 입력했거나 라벨이 없던 값)은 그대로 둔다.
    반환: (되돌린 DataFrame, 컬럼별 되돌리지 못한 값 수)
    """
    out = df.copy()
    misses = {}
    for col, vmap in (value_labels or {}).items():
        if col not in out.columns:
            continue
        rev = {}
        for code, lab in vmap.items():
            c = int(code) if isinstance(code, float) and float(code).is_integer() else code
            rev[f"{c}) {lab}"] = code
        col_s = out[col]
        n_miss = 0

        def _back(v, _rev=rev):
            nonlocal n_miss
            if pd.isna(v):
                return v
            if v in _rev:
                return _rev[v]
            n_miss += 1
            return v

        out[col] = col_s.map(_back)
        if n_miss:
            misses[col] = n_miss
    return out, misses


def sav_safe_columns(df):
    """
    SPSS 변수명 규칙에 맞게 컬럼명을 손본다.
      - 영문/숫자/밑줄만 남기고, 숫자로 시작하면 V 를 붙인다
      - 64바이트 제한, 선행 밑줄 금지
      - 한글 등으로 이름이 통째로 사라지면 VAR1, VAR2 로 대체
    반환: (이름 바꾼 DataFrame, {원래이름: 새이름} 중 바뀐 것만)
    """
    used, mapping = set(), {}
    for i, c in enumerate(df.columns, start=1):
        name = str(c)
        new = re.sub(r'[^0-9A-Za-z_]', '_', name).strip('_')
        new = re.sub(r'__+', '_', new)
        if not new or not re.match(r'^[A-Za-z]', new):
            new = ("V" + new) if new else f"VAR{i}"
        while len(new.encode('utf-8')) > 60:
            new = new[:-1]
        base, k = new, 2
        while new.upper() in used:
            new = f"{base[:57]}_{k}"
            k += 1
        used.add(new.upper())
        if new != name:
            mapping[name] = new
    out = df.rename(columns=mapping) if mapping else df
    return out, mapping


def write_sav_bytes(df, value_labels=None, column_labels=None):
    """
    DataFrame 을 .sav 바이트로 만든다.

    value_labels 를 주면 코드값과 값 라벨이 함께 저장되어, SPSS 에서 열었을 때
    원본과 같은 형태가 된다. (문자열 "1) 서울" 로 저장되는 것을 막는다)
    반환: (bytes, 이름이 바뀐 컬럼 dict, 경고 메시지 리스트)
    """
    import tempfile
    import os as _os
    try:
        import pyreadstat
    except ImportError as e:
        raise ImportError("SPSS(.sav) 로 저장하려면 pyreadstat 이 필요합니다. "
                          "pip install pyreadstat") from e

    warns = []
    out, renamed = sav_safe_columns(df)

    vl = {renamed.get(k, k): v for k, v in (value_labels or {}).items()
          if renamed.get(k, k) in out.columns}
    cl = {renamed.get(k, k): v for k, v in (column_labels or {}).items()
          if renamed.get(k, k) in out.columns}

    # 값 라벨이 붙은 열은 숫자여야 SPSS 가 제대로 읽는다.
    #
    # [수정] 예전 판정은 "숫자로 바뀐 값이 하나도 없으면 문자열"이었는데,
    # 값이 전부 결측인 열도 여기에 걸려서 "문자로 저장합니다" 라는 잘못된 경고가
    # 떴다. 실제로는 원본이 숫자형이면 그대로 숫자로 저장된다.
    # 이제 세 경우를 나눠서 본다.
    #   ① 이미 숫자형        -> 손대지 않고 라벨 유지 (경고 없음)
    #   ② 전부 결측          -> 숫자로 두고 라벨 유지 (경고 없음)
    #   ③ 값이 있는데 하나도 숫자로 안 바뀜 -> 진짜 문자열. 라벨 포기 + 경고
    # 일부만 숫자인 경우는 나머지가 결측이 되므로 그 사실을 알려준다.
    for c in list(vl):
        col = out[c]
        if pd.api.types.is_numeric_dtype(col):
            continue
        n_filled = int(col.notna().sum())
        conv = pd.to_numeric(col, errors='coerce')
        n_num = int(conv.notna().sum())
        if n_filled == 0:
            out[c] = conv
            continue
        if n_num == 0:
            warns.append(f"`{c}` 는 숫자가 아니라서 값 라벨 없이 문자로 저장합니다.")
            vl.pop(c)
            continue
        if n_num < n_filled:
            warns.append(
                f"`{c}` 에서 숫자로 바꿀 수 없는 값 {n_filled - n_num:,}건은 "
                "결측으로 저장됩니다.")
        out[c] = conv

    with tempfile.TemporaryDirectory() as td:
        path = _os.path.join(td, "out.sav")
        pyreadstat.write_sav(out, path,
                             variable_value_labels=vl or None,
                             column_labels=[cl.get(c, "") for c in out.columns]
                             if cl else None)
        with open(path, "rb") as f:
            data = f.read()
    return data, renamed, warns


# ==============================================================================
# 6. 쿼터표 엑셀 — 목록(평면) 형식 파서
#    피벗 형식(transform_pivoted_quota)은 변수 3개로 고정이다.
#    목록 형식은 변수 개수 제한이 없고, 머리글이 데이터 컬럼명이라
#    qt1/qt2/qt3 를 화면에서 지정할 필요도 없다.
# ==============================================================================
TARGET_ALIASES = {"목표", "목표수", "목표인원", "target", "n", "쿼터", "할당"}


def _find_target_col(cols):
    for c in cols:
        if str(c).strip().lower().replace(" ", "") in TARGET_ALIASES:
            return c
    return None


def parse_main_flat(df_raw, data_columns):
    """
    목록 형식 메인 쿼터표를 읽는다.

        SQ1   | SQ2  | SQ3  | SQ4  | 목표
        남성   | 서울  | 20대 | 사무  | 12

    - 머리글은 **데이터의 컬럼명**이어야 한다 (변수 지정 단계가 없어진다)
    - 마지막에 목표 열이 있어야 한다 (목표/목표수/target 등)
    - 변수 개수는 제한 없다
    - 목표가 1 이상인 행만 쓴다

    반환: (main_map, 사용된 컬럼 목록, 오류/경고 목록)
    """
    errs = []
    if df_raw is None or df_raw.empty:
        return {}, [], ["쿼터표 시트가 비어 있습니다."]

    tcol = _find_target_col(df_raw.columns)
    if tcol is None:
        return {}, [], ["목표 열을 찾지 못했습니다. "
                        "머리글에 '목표' 라는 열이 있어야 합니다."]

    qcols = [c for c in df_raw.columns
             if c != tcol and not str(c).startswith("Unnamed")]
    if not qcols:
        return {}, [], ["쿼터 변수 열이 없습니다."]

    # 표 오른쪽에 메모나 합계 같은 보조 칸이 있으면 그것도 열로 읽힌다.
    # 값이 거의 없는 열은 쿼터 변수가 아니라고 보고 무시한다.
    dset = set(data_columns)
    ignored = []
    for c in list(qcols):
        if c in dset:
            continue
        filled = df_raw[c].notna().sum()
        if filled <= max(1, int(len(df_raw) * 0.1)):
            qcols.remove(c)
            ignored.append(c)
    if ignored:
        errs.append("표 밖의 열로 보여 무시했습니다: "
                    + ", ".join(map(str, ignored[:8])))

    unknown = [c for c in qcols if c not in dset]
    if unknown:
        return {}, [], [
            "쿼터표 머리글이 데이터 컬럼명과 다릅니다: "
            + ", ".join(map(str, unknown[:8]))
            + ". 머리글을 데이터 컬럼명과 똑같이 적어 주세요."]

    main_map, dups, bad = {}, [], []
    for i, r in df_raw.iterrows():
        ln = i + 2
        key = tuple(norm_val(r[c]) for c in qcols)
        if all(k == NA_TOKEN for k in key):
            continue
        try:
            t = int(float(r[tcol])) if pd.notna(r[tcol]) else 0
        except (TypeError, ValueError):
            bad.append(f"{ln}행({' / '.join(key)})")
            continue
        if t <= 0:
            continue
        if key in main_map:
            dups.append(f"{ln}행({' / '.join(key)})")
            continue
        main_map[key] = t

    if bad:
        errs.append(f"목표를 숫자로 읽지 못한 행 {len(bad)}개를 건너뜁니다: "
                    + ", ".join(bad[:5]))
    if dups:
        errs.append(f"같은 조합이 두 번 나온 행 {len(dups)}개를 건너뜁니다: "
                    + ", ".join(dups[:5]))
    if not main_map:
        errs.append("목표가 1 이상인 행이 하나도 없습니다.")
    return main_map, qcols, errs


def parse_extra_flat(df_raw, data_columns):
    """
    '추가쿼터' 시트를 읽어 추가 쿼터 설정 목록을 만든다.

        그룹명 | 방식 | 변수            | 값1   | 값2     | 값3 | 목표
        직업   | 단순 | DQ1            | 사무직 |         |     | 350
        브랜드 | 단순 | Q5_1,Q5_2,Q5_3 | 삼성   |         |     | 400
        권역x연령| 조합 | 권역,연령대     | 수도권 | 20~30대 |     | 180

    - 같은 그룹명끼리 한 그룹으로 묶는다
    - 단순형은 값1 만, 조합형은 변수 개수만큼 값1..값3 을 채운다
    - 값을 한 칸에 쉼표로 잇지 않는 이유 : "1,000만원 이상" 처럼 라벨 안에
      쉼표가 들어가는 경우가 있다

    반환: ([{'name','mode','cols','map'}, ...], 오류 목록)
    """
    need = ["그룹명", "방식", "변수"]
    if df_raw is None:
        return [], []
    tcol = _find_target_col(df_raw.columns)
    miss = [c for c in need if c not in df_raw.columns]
    if miss or tcol is None:
        # 머리글이 아예 없으면 '시트가 비어 있다' 로 본다 (오류 아님)
        if df_raw.empty:
            return [], []
        return [], ["필수 열이 없습니다: "
                    + ", ".join(miss + ([] if tcol else ["목표"]))]
    if df_raw.empty:
        return [], []                      # 머리글만 있는 시트 = 추가 쿼터 없음

    dcols = set(data_columns)
    errs, order, buf = [], [], {}
    for i, r in df_raw.iterrows():
        ln = i + 2
        name = str(r["그룹명"]).strip() if pd.notna(r["그룹명"]) else ""
        if not name:
            continue
        mode_raw = str(r["방식"]).strip() if pd.notna(r["방식"]) else ""
        if not (mode_raw.startswith("조합") or mode_raw.startswith("단순")):
            errs.append(f"{ln}행: 방식은 '단순' 또는 '조합' 이어야 합니다 "
                        f"('{mode_raw}')")
            continue
        mode = "grid" if mode_raw.startswith("조합") else "simple"

        cols = [c.strip() for c in str(r["변수"]).split(",") if c.strip()] \
            if pd.notna(r["변수"]) else []
        if not cols:
            errs.append(f"{ln}행: 변수가 비어 있습니다")
            continue
        nf = [c for c in cols if c not in dcols]
        if nf:
            errs.append(f"{ln}행: 데이터에 없는 변수 {', '.join(nf)}")
            continue
        try:
            t = int(float(r[tcol])) if pd.notna(r[tcol]) else 0
        except (TypeError, ValueError):
            errs.append(f"{ln}행: 목표를 숫자로 읽을 수 없습니다 ('{r[tcol]}')")
            continue

        vals = []
        for k in (1, 2, 3):
            v = r.get(f"값{k}")
            if pd.notna(v) and str(v).strip() != "":
                vals.append(norm_val(v))
        if mode == "simple":
            if len(vals) != 1:
                errs.append(f"{ln}행: 단순형은 값1 만 채워야 합니다 "
                            f"(지금 {len(vals)}개)")
                continue
            key = vals[0]
        else:
            if len(vals) != len(cols):
                errs.append(f"{ln}행: 조합형은 변수 {len(cols)}개만큼 값을 "
                            f"채워야 합니다 (지금 {len(vals)}개)")
                continue
            key = tuple(vals)

        if name not in buf:
            buf[name] = {"name": name, "mode": mode, "cols": cols, "map": {}}
            order.append(name)
        cfg = buf[name]
        if cfg["mode"] != mode or cfg["cols"] != cols:
            errs.append(f"{ln}행: 그룹 '{name}' 안에서 방식/변수가 다릅니다")
            continue
        if key in cfg["map"]:
            errs.append(f"{ln}행: 그룹 '{name}' 에 값 '{key}' 가 중복됩니다")
            continue
        cfg["map"][key] = t

    return [buf[n] for n in order], errs


def build_quota_form_xlsx(main_map, main_cols, ex_configs, source_name=""):
    """
    현재 화면 설정으로 '쿼터표 업로드 양식' 엑셀을 만든다.

    한 번 화면에서 설정한 뒤 이 양식을 받아두면, 다음 조사에서는 목표 숫자만
    고쳐서 업로드하면 된다. 만들어지는 시트는 parse_main_flat /
    parse_extra_flat 이 그대로 읽을 수 있는 형식이다.

      메인쿼터 : <컬럼명들> | 목표          (목록 형식, 변수 개수 제한 없음)
      추가쿼터 : 그룹명 | 방식 | 변수 | 값1 | 값2 | 값3 | 목표
      사용법   : 이 파일을 어떻게 고쳐 쓰는지

    반환: bytes
    """
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    F = "Arial"
    HDR = PatternFill("solid", fgColor="1F3864")
    HF = Font(name=F, bold=True, color="FFFFFF", size=10)
    KEY = PatternFill("solid", fgColor="DDEBF7")
    YEL = PatternFill("solid", fgColor="FFF2CC")
    GRY = PatternFill("solid", fgColor="F2F2F2")
    BODY = Font(name=F, size=10)
    BOLD = Font(name=F, bold=True, size=10)
    THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

    wb = Workbook()

    # ── 메인쿼터 (목록 형식) ──────────────────────────────────────────
    ws = wb.active
    ws.title = "메인쿼터"
    cols = [str(c) for c in (main_cols or [])]
    no_main = not cols
    if no_main:
        # 메인 쿼터를 쓰지 않는 설정. 목록 형식으로 표현할 수 없으므로
        # 머리글만 남기고, 총 목표 인원은 사용법 시트에 적어 둔다.
        cols = []
    ws.append(cols + ["목표"])
    for j in range(1, len(cols) + 2):
        c = ws.cell(1, j)
        c.fill, c.font = HDR, HF
        c.alignment = Alignment(horizontal="center")

    def _sk(k):
        return tuple(k) if isinstance(k, tuple) else (k,)

    rows = sorted(main_map.items(), key=lambda kv: [natural_key(x)
                                                    for x in _sk(kv[0])])
    for k, v in rows:
        ws.append(list(_sk(k)) + [int(v)])
    for i in range(2, 2 + len(rows)):
        for j in range(1, len(cols) + 2):
            c = ws.cell(i, j)
            c.font, c.border = BODY, THIN
            if j == len(cols) + 1:
                c.fill = YEL
                c.alignment = Alignment(horizontal="center")
            else:
                c.fill = KEY
    ws.freeze_panes = "A2"
    for j in range(1, len(cols) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 13

    # ── 추가쿼터 ──────────────────────────────────────────────────────
    ws = wb.create_sheet("추가쿼터")
    EC = ["그룹명", "방식", "변수", "값1", "값2", "값3", "목표"]
    ws.append(EC)
    for j in range(1, len(EC) + 1):
        c = ws.cell(1, j)
        c.fill, c.font = HDR, HF
        c.alignment = Alignment(horizontal="center")

    n_ex = 0
    for cfg in (ex_configs or []):
        if not cfg.get("cols") or not cfg.get("map"):
            continue
        mode_txt = "조합" if cfg.get("mode") == "grid" else "단순"
        var_txt = ",".join(str(c) for c in cfg["cols"])
        for k, v in sorted(cfg["map"].items(),
                           key=lambda kv: [natural_key(x) for x in _sk(kv[0])]):
            vals = list(_sk(k))[:3] + [None] * (3 - min(3, len(_sk(k))))
            ws.append([cfg.get("name") or f"추가{n_ex+1}", mode_txt, var_txt]
                      + vals + [int(v)])
        n_ex += 1
    last = ws.max_row
    for i in range(2, last + 1):
        simple = ws.cell(i, 2).value == "단순"
        for j in range(1, len(EC) + 1):
            c = ws.cell(i, j)
            c.font, c.border = BODY, THIN
            if j <= 3:
                c.fill = KEY
            elif j == len(EC):
                c.fill = YEL
                c.alignment = Alignment(horizontal="center")
            else:
                c.fill = GRY if (simple and j > 4) else YEL
    ws.freeze_panes = "A2"
    for j, w in enumerate([14, 8, 22, 14, 14, 14, 9], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    # [주의] 여기에 안내문 행을 넣으면 파서가 데이터로 읽어 오류가 난다.
    # 추가 쿼터가 없으면 머리글만 남긴다. 설명은 사용법 시트에 있다.

    # ── 사용법 ────────────────────────────────────────────────────────
    ws = wb.create_sheet("사용법")
    guide = [
        ["쿼터표 업로드 양식 (현재 화면 설정으로 자동 생성)"],
        [""],
        ["만든 기준", f"데이터 파일: {source_name or '-'}"],
        ["", (f"메인 쿼터: {' × '.join(cols)} / {len(main_map):,}셀 / "
              f"목표 합계 {sum(main_map.values()):,}명") if not no_main else
             (f"메인 쿼터: 사용 안 함 (전체 목표 "
              f"{sum(main_map.values()):,}명 — 화면에서 직접 입력하세요)")],
        ["", f"추가 쿼터: {n_ex}개 그룹"],
        [""],
        ["쓰는 법", "1) 노란 칸(목표)의 숫자만 고칩니다."],
        ["", "2) 이 파일을 그대로 '쿼터 파일' 로 업로드합니다."],
        ["", "3) 메인 쿼터 방식에서 '엑셀 업로드' 를 고르면 됩니다."],
        [""],
        ["메인쿼터 시트", "머리글이 데이터 컬럼명입니다. 그래서 화면에서 변수를"],
        ["", "지정하지 않아도 됩니다. 컬럼명을 바꾸지 마세요."],
        ["", "목표가 0이거나 빈 행은 '이 셀은 쓰지 않음' 으로 처리됩니다."],
        ["", "행을 지워도 되고, 새 조합을 아래에 추가해도 됩니다."],
        [""],
        ["추가쿼터 시트", "같은 '그룹명' 끼리 한 그룹으로 묶입니다."],
        ["", "방식은 '단순' 또는 '조합' 입니다."],
        ["", "변수는 데이터 컬럼명이며 여러 개면 쉼표로 잇습니다."],
        ["", "단순형은 값1만, 조합형은 변수 개수만큼 값1~값3을 채웁니다."],
        ["", "회색 칸은 그 줄에서 쓰지 않는 칸입니다."],
        [""],
        ["주의", "· 값 표기를 바꾸지 마세요. 데이터의 값과 정확히 같아야 합니다."],
        ["", "· SPSS 파일을 올린 경우 값이 '1) 서울' 형태입니다. 그대로 두세요."],
        ["", "· 데이터 시트에 메모나 합계 칸을 넣지 마세요. 열로 읽힙니다."],
        ["", "· 같은 조합/값을 두 번 적으면 그 줄은 건너뜁니다."],
    ]
    for g in guide:
        ws.append(g)
    ws["A1"].font = Font(name=F, bold=True, size=12)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if not c.font.bold:
                c.font = BODY
    for rr in (3, 7, 11, 16, 22):
        ws.cell(rr, 1).font = BOLD
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 86

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_quota_report(main_map, main_cols, main_actual,
                       ex_configs, ex_actual, n_selected):
    """
    쿼터표와 같은 배치로 '목표 / 달성 / 차이' 표를 만든다.

    SPSS 에서 프리퀀시·크로스탭을 다시 돌리지 않아도 되도록, 화면에서 설정한
    쿼터표 모양 그대로 실적을 채워 준다. 소프트 쿼터로 목표와 어긋난 셀이
    어디인지 바로 보인다.

    교차 방향은 쿼터 설정과 같게 잡는다 : 마지막 변수를 열, 나머지를 행.
    변수가 1개면 교차 없이 목록으로 만든다.

    반환: [(제목, DataFrame), ...]  — 호출부가 한 시트에 세로로 이어 붙인다
    """
    blocks = []
    cols = [str(c) for c in (main_cols or [])]

    def _sk(k):
        return tuple(k) if isinstance(k, tuple) else (k,)

    # ── 메인 쿼터 ────────────────────────────────────────────────────
    if main_map:
        keys = list(main_map)
        if len(cols) >= 2:
            row_cols = cols[:-1]
            row_keys = sorted({_sk(k)[:-1] for k in keys},
                              key=lambda t: [natural_key(x) for x in t])
            col_keys = sorted({_sk(k)[-1] for k in keys}, key=natural_key)

            def _pivot(getval):
                data = []
                for rk in row_keys:
                    row = list(rk)
                    for ck in col_keys:
                        row.append(getval(tuple(rk) + (ck,)))
                    row.append(sum(getval(tuple(rk) + (c,)) for c in col_keys))
                    data.append(row)
                tot = [""] * (len(row_cols) - 1) + ["합계"]
                for ck in col_keys:
                    tot.append(sum(getval(tuple(rk) + (ck,)) for rk in row_keys))
                tot.append(sum(tot[len(row_cols):]))
                data.append(tot)
                return pd.DataFrame(
                    data, columns=row_cols + [str(c) for c in col_keys] + ["합계"])

            blocks.append(("메인 쿼터 — 목표", _pivot(
                lambda k: int(main_map.get(k, 0)))))
            blocks.append((f"메인 쿼터 — 달성 (통과 {n_selected:,}명)", _pivot(
                lambda k: int(main_actual.get(k, 0)))))
            blocks.append(("메인 쿼터 — 차이 (달성 − 목표)", _pivot(
                lambda k: int(main_actual.get(k, 0)) - int(main_map.get(k, 0)))))
        else:
            name = cols[0] if cols else "전체"
            rows = []
            for k in sorted(keys, key=lambda x: [natural_key(y) for y in _sk(x)]):
                t, a = int(main_map[k]), int(main_actual.get(k, 0))
                rows.append({name: " / ".join(_sk(k)), "목표": t, "달성": a,
                             "차이": a - t})
            df = pd.DataFrame(rows)
            df.loc[len(df)] = {name: "합계", "목표": df["목표"].sum(),
                               "달성": df["달성"].sum(), "차이": df["차이"].sum()}
            blocks.append((f"메인 쿼터 (통과 {n_selected:,}명)", df))

    # ── 추가 쿼터 ────────────────────────────────────────────────────
    for j, cfg in enumerate(ex_configs or []):
        if not cfg.get("cols") or not cfg.get("map"):
            continue
        act = ex_actual[j] if j < len(ex_actual) else {}
        is_grid = cfg.get("mode") == "grid" and len(cfg["cols"]) >= 2

        if is_grid:
            row_cols = cfg["cols"][:-1]
            rks = sorted({_sk(k)[:-1] for k in cfg["map"]},
                         key=lambda t: [natural_key(x) for x in t])
            cks = sorted({_sk(k)[-1] for k in cfg["map"]}, key=natural_key)
            data = []
            for rk in rks:
                row = list(rk)
                for ck in cks:
                    kk = tuple(rk) + (ck,)
                    row.append(f"{int(act.get(kk, 0))} / {int(cfg['map'].get(kk, 0))}")
                data.append(row)
            df = pd.DataFrame(data, columns=[str(c) for c in row_cols]
                              + [str(c) for c in cks])
            blocks.append((f"추가 쿼터 — {cfg.get('name')} "
                           f"(칸 안은 달성 / 목표)", df))
        else:
            rows = []
            for k in sorted(cfg["map"], key=lambda x: natural_key(
                    " / ".join(_sk(x)))):
                t = int(cfg["map"][k])
                a = int(act.get(k, 0))
                rows.append({
                    "값": " / ".join(_sk(k)) if isinstance(k, tuple) else str(k),
                    "목표": t, "달성": a, "차이": a - t,
                    "달성률": (a / t) if t else None,
                    "구성비": (a / n_selected) if n_selected else None})
            df = pd.DataFrame(rows)
            df.loc[len(df)] = {
                "값": "합계", "목표": df["목표"].sum(), "달성": df["달성"].sum(),
                "차이": df["차이"].sum(),
                "달성률": (df["달성"].sum() / df["목표"].sum())
                if df["목표"].sum() else None,
                "구성비": (df["달성"].sum() / n_selected) if n_selected else None}
            multi = len(cfg["cols"]) > 1
            blocks.append((f"추가 쿼터 — {cfg.get('name')}"
                           + ("  (복수응답: 구성비 합이 100%를 넘을 수 있음)"
                              if multi else ""), df))
    return blocks

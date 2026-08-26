"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : pages/7_🔄_행열 변환.py                                         ║
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
             pyreadstat  ← SPSS(.sav) 읽기/저장용. requirements.txt 에 추가.
                            없으면 sav 관련 기능만 막히고 나머지는 정상 동작.

수정 이력
---------
v1.1  1) 변환 후 설정을 바꾸면 옛 결과가 남아 화면과 다운로드 내용이 어긋나던 문제
         수정. 설정 지문(_SIG)을 결과와 함께 저장해 두고 달라지면 폐기한다.
      2) 다운로드용 xlsx/csv 를 변환 시점에 1회만 생성. 이전에는 위젯을 건드릴
         때마다 xlsx 를 처음부터 다시 써서 큰 데이터에서 매번 수 초씩 멈췄다.
      3) @st.cache_data 에 max_entries 지정. 큰 파일을 여러 개 올렸을 때
         캐시가 무한히 쌓여 메모리를 다 쓰는 것을 막는다.
v1.2  4) 결과 열 이름이 겹칠 때 _2, _3 을 붙여 자동 해소 (_dedupe_names).
         이전에는 판다스 내부 에러가 그대로 노출됐다.
           · ID='학번' 인데 기준 열에 '학번' 값 -> cannot insert 학번
           · 값 열 2개 조합이 같아짐      -> AttributeError: no attribute 'dtype'
      5) 정규화로 서로 다른 값이 한 열로 합쳐질 때 어떤 값들이 합쳐지는지
         목록으로 표시. 기존 중복 경고만으로는 원인을 알 수 없었다.
      6) 표시용 _arrow_safe 추가. 기준 열에 숫자와 문자가 섞이면 st.dataframe
         이 매번 Arrow 변환에 실패해 로그에 에러를 대량으로 남겼다.
v1.3  7) 중복값 처리에 "여러 열로 펼치기" 모드 추가 (expand_duplicates).
         같은 ID + 같은 변수가 N번 나오면 취미_1, 취미_2 … 로 펼친다.
         값을 버리지 않으므로 이 모드가 기본값이다.
         v1.2 의 _dedupe_names 는 '열 이름 충돌' 만 다루므로 별개다.
           · always_suffix   : 중복 없는 변수에도 _1 을 붙일지
           · max_occurrences : 변수당 최대 개수 (0 = 제한 없음)
           · max_result_cols : ID 열을 잘못 지정해 열이 폭발하는 것을 막는 상한
v1.4  8) DEFAULT_ID_COLS / DEFAULT_VALUE_COLS / DEFAULT_KEY_COL 로 기본 선택
         열 지정. 해당 이름의 열이 있으면 자동 선택하고 없으면 비워 둔다.
         겸사겸사, 다른 파일을 올렸을 때 이전 선택값이 새 옵션에 없어
         multiselect 가 죽던 문제도 함께 막힌다 (lw_colsig 로 감지).
v1.5  9) 변수 이름 바꾸기 추가 (rename_map). 변수 단위로 바꾸므로 펼침 모드에서는
         바뀐 이름 뒤에 순번이 붙는다 (Q3_추천 -> 추천 이면 추천_1, 추천_2).
         서로 같은 이름을 넣거나 ID 열 이름과 같게 바꾸면 _dedupe_names 가
         _2, _3 을 붙여 살리고 경고를 띄운다.
     10) DEFAULT_VALUE_COLS 를 intVal 로 정정.
v1.6 11) SPSS .sav 다운로드 추가 (build_sav / _spss_safe_names).
         SPSS 변수명 규칙은 실측으로 확인했다.
           · 한글은 되지만 길이 한도가 64'바이트' 라 한글은 21자까지
           · 공백·괄호·- 등 특수문자 불가 (기본 결측표기 '(무응답)' 도 거부됨)
           · 첫 글자는 반드시 문자. 밑줄로 시작해도 거부된다
             (readstat 에러 메시지는 밑줄이 허용되는 것처럼 안내하지만 아니다)
         바꾼 이름은 경고로 알리고 원래 이름은 SPSS 변수 라벨로 보존한다.
         엑셀·CSV 는 원래 이름을 그대로 쓴다.
v1.7 12) 변수 이름 바꾸기를 st.expander 밖으로 꺼내 항상 펼쳐진 표로 바꿨다.
         expander 제목에 "N개 변경됨" 을 넣었더니 이름을 고칠 때마다 제목이
         바뀌어 Streamlit 이 다른 위젯으로 인식하고 패널이 저절로 닫혔다.
         data_editor 에는 변수 목록으로 만든 해시를 key 로 주어, 목록이
         그대로인 동안 편집 상태가 유지되고 목록이 바뀌면 새 표로 시작한다.
v1.8 13) .sav 입력 지원 (read_sav_bytes). 변수 라벨·값 라벨을 함께 읽어
         결과 .sav 에 물려준다. 원본 .sav 로 시작하면 1='매우 불만족' 같은
         값 라벨이 SPSS 에서 그대로 보인다 (엑셀 원본은 라벨이 없으니 해당 없음).
     14) 가로 → 세로 (VARSTOCASES) 모드 추가 (wide_to_long / screen_wide_to_long).
         화면 맨 위 라디오로 방향을 고른다. 기존 세로→가로 코드는 그대로 두고
         가로→세로는 별도 함수로 분리해, 한쪽 수정이 다른 쪽에 닿지 않게 했다.
           · /NULL=DROP 에 해당하는 '빈 값 행 제외', /INDEX 유사한 '순번 열'
           · 내린 열들의 값 라벨이 모두 같을 때만 값 열에 라벨을 붙인다
             (서로 다르면 어느 것을 붙여도 틀리므로 넣지 않는다)
v1.9 15) 반복측정 지원 = SPSS /MAKE 여러 개 (detect_value_groups /
         wide_to_long_multi / screen_w2l_multi). 가로→세로의 핵심 기능.
           점수_1차 시간_1차 점수_2차 시간_2차
             -> 시점 | 점수 | 시간   (값 열이 여러 개 생긴다)
         열 이름에서 '접두사 + 꼬리표' 구조를 자동으로 찾는다. 구분자
         후보(_ . - 끝자리숫자)를 모두 시도해 그룹이 가장 많이 잡히는 것을
         고르고, 결과를 표로 보여주고 이름만 고치게 한다.
           · 수준이 하나뿐인 열은 그룹이 못 되어 결과에서 빠지므로,
             어떤 열이 빠지는지 경고로 반드시 알린다 (조용히 사라지면 안 된다)
           · 그룹별로 원본 열들의 값 라벨이 모두 같을 때만 라벨을 붙인다
           · 자동 인식 실패 시 '값 열 1개' 모드로 안내한다
     16) 다운로드/저장 블록을 emit_downloads, store_outputs 로 공통화.
"""

import hashlib
import hmac
import io
import os
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


def _dedupe_names(names, reserved=()):
    """
    결과 열 이름이 겹치면 _2, _3 … 을 붙여 유일하게 만든다.

    겹치는 경로가 두 가지 있다.
      ③ 기준 열의 값이 ID 열 이름과 같은 경우
         (ID='학번' 인데 항목 열에 '학번' 이라는 값이 있음)
      ④ 값 열이 2개 이상일 때 조합된 이름이 서로 같아지는 경우
         ('점수' + '_' + '1_차' 와 '점수_1' + '_' + '차' 가 모두 '점수_1_차')
    처리하지 않으면 판다스 내부 에러가 그대로 사용자에게 노출된다.

    reserved : ID 열 이름. 여기에도 겹치면 안 된다.
    반환     : (확정된 이름 목록, [(원래 이름, 바뀐 이름), …])
    """
    used = {str(r) for r in reserved}
    out, renamed = [], []
    for n in names:
        n = str(n)
        if n not in used:
            used.add(n)
            out.append(n)
            continue
        i = 2
        while f"{n}_{i}" in used:
            i += 1
        new = f"{n}_{i}"
        used.add(new)
        out.append(new)
        renamed.append((n, new))
    return out, renamed


def long_to_wide(df, id_cols, key_col, value_cols, keep_keys=None,
                 aggfunc="first", name_sep="_", normalize_keys=True,
                 expand_duplicates=False, always_suffix=False,
                 max_occurrences=0, max_result_cols=2000, rename_map=None):
    """
    세로(long) → 가로(wide) 변환. 실패 시 ValueError 를 raise 한다.

    id_cols    : 한 행을 식별하는 열 (예: ["학번"], ["학번","이름"])
    key_col    : 가로로 펼칠 변수명이 담긴 열 (예: "항목", "시점")
    value_cols : 값이 담긴 열. 2개 이상이면 "점수_1차" 처럼 조합된다
    keep_keys  : 펼칠 변수 목록. 준 순서가 결과 열 순서가 된다
    rename_map : {원본 변수값: 결과 열 이름}. 지정한 변수만 이름이 바뀐다.
                 펼침 모드에서는 바뀐 이름 뒤에 순번이 붙는다 (만족도_1, 만족도_2).
                 이름이 서로 겹치면 _dedupe_names 가 _2, _3 을 붙여 살린다.

    중복(같은 ID + 같은 변수가 2번 이상) 처리 방식 두 가지
    ------------------------------------------------------
    expand_duplicates=False : aggfunc 으로 하나만 남긴다 (값이 버려짐)
    expand_duplicates=True  : 등장 순서대로 취미_1, 취미_2, 취미_3 … 으로
                              여러 열에 펼친다 (값이 보존됨)
        always_suffix   : 중복이 없는 변수에도 _1 을 붙일지
        max_occurrences : 변수당 최대 몇 개까지 펼칠지 (0 = 제한 없음)
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

    # 이름 바꾸기 표. 빈 문자열은 '안 바꿈' 으로 취급한다.
    rmap = {}
    for k, v in (rename_map or {}).items():
        v = "" if v is None else str(v).strip()
        if v:
            rmap[str(k)] = v

    def _base_name(v_col, key):
        """결과 열의 기본 이름. 값 열이 2개 이상이면 값 열 이름과 조합한다."""
        nm = rmap.get(str(key), str(key))
        return nm if len(value_cols) == 1 else f"{v_col}{name_sep}{nm}"

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

    occ_max, truncated = {}, 0

    if expand_duplicates:
        # 같은 ID + 같은 변수 안에서 등장 순번(1,2,3…)을 매긴다.
        # 이 순번을 열 축에 함께 넣으면 중복이 열로 펼쳐진다.
        work["__occ"] = work.groupby(id_cols + [key_col], dropna=False,
                                     observed=True).cumcount() + 1

        if max_occurrences and max_occurrences > 0:
            over = int((work["__occ"] > max_occurrences).sum())
            if over:
                truncated = over
                work = work[work["__occ"] <= max_occurrences]

        occ_max = {k: int(v) for k, v in
                   work.groupby(key_col, dropna=False, observed=True)["__occ"]
                   .max().items()}

        n_cols = len(value_cols) * sum(max(occ_max.get(k, 1), 1) for k in keep_keys)
        if n_cols > max_result_cols:
            worst = sorted(occ_max.items(), key=lambda x: -x[1])[:3]
            raise ValueError(
                f"펼치면 열이 {n_cols:,}개가 됩니다 (상한 {max_result_cols:,}개).\n"
                f"가장 많이 중복된 변수: "
                + ", ".join(f"{k}({v}회)" for k, v in worst)
                + "\nID 열 지정이 잘못되었을 가능성이 높습니다. "
                  "'변수당 최대 개수' 를 지정하거나 ID 열을 다시 확인하세요.")

        try:
            wide = pd.pivot_table(
                work, index=id_cols, columns=[key_col, "__occ"],
                values=value_cols,
                aggfunc="first",   # (ID, 변수, 순번) 은 유일하므로 집계가 아니다
                dropna=True, observed=True,
            )
        except Exception as e:
            raise ValueError(f"피벗 실패 ({type(e).__name__}: {e})")

        pairs, raw_names = [], []
        for v in value_cols:
            for k in keep_keys:
                mx = max(int(occ_max.get(k, 1)), 1)
                base = _base_name(v, k)
                for o in range(1, mx + 1):
                    pairs.append((v, k, o))
                    # 중복이 없는 변수는 이름을 깨끗하게 둔다
                    raw_names.append(base if (mx == 1 and not always_suffix)
                                     else f"{base}{name_sep}{o}")
    else:
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

        pairs = [(v, k) for v in value_cols for k in keep_keys]
        raw_names = [_base_name(v, k) for v, k in pairs]

    # ── 결과 열 이름 확정 ────────────────────────────────────────────
    # (값 열, 변수[, 순번]) 쌍을 표준 식별자로 삼고 이름은 마지막에 붙인다.
    # 이름이 겹쳐도 어느 칸인지 잃지 않기 위한 것이다.
    if not isinstance(wide.columns, pd.MultiIndex):
        wide.columns = pd.MultiIndex.from_tuples(
            [(value_cols[0], c) for c in wide.columns])

    # 값이 전부 비어 사라진 칸도 빈 열로 되살리고, 선택 순서대로 정렬한다
    wide = wide.reindex(columns=pd.MultiIndex.from_tuples(pairs))

    final_names, renamed = _dedupe_names(raw_names, reserved=id_cols)
    wide.columns = final_names

    wide = wide.reset_index()
    wide = order.merge(wide, on=id_cols, how="left")   # ID 원래 순서 복원
    wide = _restore_ints(wide)
    # 결과 열 -> 원본 값 열 대응. .sav 저장 시 값 라벨을 물려주는 데 쓴다.
    wide.attrs["col_source"] = {
        fn: pr[0] for fn, pr in zip(final_names, pairs)}
    # 화면에서 경고로 알리기 위한 부가 정보
    wide.attrs["renamed_columns"] = renamed
    wide.attrs["occ_max"] = occ_max
    wide.attrs["truncated_rows"] = truncated
    return wide


def _spss_safe_names(cols, limit_bytes=64):
    """
    SPSS 변수명 규칙에 맞게 열 이름을 손질한다.

    실측한 제약 (pyreadstat/readstat)
      · 한글은 허용된다. 단 길이 한도는 문자 수가 아니라 **64바이트** 이므로
        UTF-8 한글은 21자까지다.
      · 공백, 괄호, -, % 같은 특수문자는 거부된다.
        기본 결측 표기인 '(무응답)' 도 괄호 때문에 그대로는 저장되지 않는다.
      · 숫자로 시작하면 거부된다. 밑줄로 시작해도 거부된다.
        (readstat 의 에러 메시지는 밑줄이 허용되는 것처럼 안내하지만
         실제로 저장을 시도하면 거부된다. 실측으로 확인했다.)
      · 허용: 문자·숫자·밑줄, 그리고 . @ # $ (단 첫 글자는 문자만)

    반환: (안전한 이름 목록, [(원래 이름, 바뀐 이름), …])
    바뀐 이름은 호출부에서 경고로 보여주고, 원래 이름은 변수 라벨로 보존한다.
    """
    out, changed, used = [], [], set()
    for c in cols:
        orig = str(c)
        # 허용되지 않는 문자를 밑줄로. \w 는 유니코드 문자·숫자·밑줄을 포함한다.
        name = re.sub(r"[^\w.@#$]", "_", orig, flags=re.UNICODE)
        # 첫 글자가 문자가 아니면 접두사를 붙인다 (밑줄·숫자·기호 모두 해당)
        if not name or not name[0].isalpha():
            name = "V" + name
        # 64바이트로 자른다 (한글이 잘려 깨지지 않도록 바이트 단위로 확인)
        while len(name.encode("utf-8")) > limit_bytes:
            name = name[:-1]
        # 자르거나 치환한 뒤 겹칠 수 있으므로 유일하게 만든다
        base, i = name, 2
        while name in used:
            suffix = f"_{i}"
            name = base
            while len((name + suffix).encode("utf-8")) > limit_bytes:
                name = name[:-1]
            name = name + suffix
            i += 1
        used.add(name)
        out.append(name)
        if name != orig:
            changed.append((orig, name))
    return out, changed


def build_sav(result, sanitize=True, col_labels=None, value_labels=None):
    """
    변환 결과를 SPSS .sav 바이트로 만든다.

    col_labels   : {결과 열 이름: 변수 라벨}. 없으면 결과 열 이름을 라벨로 쓴다.
    value_labels : {결과 열 이름: {코드: 라벨}}. 원본이 .sav 일 때 값 라벨을
                   그대로 물려주기 위한 것이다.

    반환: (bytes, [(원래 이름, 바뀐 이름), …])
    pyreadstat 이 없으면 ImportError 를 그대로 올린다 (호출부에서 안내).
    """
    import tempfile

    import pyreadstat

    d = result.copy()

    # pandas 확장 dtype(Int64, string)은 그대로도 대체로 쓰이지만,
    # 버전에 따라 실패하므로 numpy/object 로 내려서 안전하게 만든다.
    for c in d.columns:
        s = d[c]
        if isinstance(s.dtype, pd.CategoricalDtype):
            d[c] = s.astype(object)
        elif str(s.dtype) in ("Int64", "Int32", "Float64"):
            d[c] = s.astype("float64")
        elif str(s.dtype).startswith("string") or s.dtype == object:
            d[c] = s.apply(lambda v: None if pd.isna(v) else str(v))
        elif str(s.dtype) == "boolean":
            d[c] = s.astype("float64")

    # 변수 라벨: 지정된 것이 있으면 쓰고, 없으면 결과 열 이름을 그대로 넣는다.
    _cl = col_labels or {}
    labels = [str(_cl.get(c) or c) for c in d.columns]

    changed = []
    orig_cols = list(d.columns)
    if sanitize:
        safe, changed = _spss_safe_names(d.columns)
        d.columns = safe

    # 값 라벨은 '손질된 변수명' 기준으로 다시 매핑해야 한다
    vl = {}
    for before, after in zip(orig_cols, d.columns):
        got = (value_labels or {}).get(before)
        if got:
            vl[after] = {k: str(v) for k, v in got.items()}

    # Windows 에서 NamedTemporaryFile 은 열린 상태로 다시 열 수 없어
    # 실패하므로 TemporaryDirectory 를 쓴다.
    with tempfile.TemporaryDirectory() as dirpath:
        path = os.path.join(dirpath, "out.sav")
        pyreadstat.write_sav(d, path, column_labels=labels,
                             variable_value_labels=(vl or None))
        with open(path, "rb") as f:
            return f.read(), changed


def detect_value_groups(cols):
    """
    열 이름에서 '값 그룹' 과 '수준(시점)' 을 자동으로 찾아낸다.

        점수_1차, 시간_1차, 점수_2차, 시간_2차
          -> 그룹 {점수: {1차: 점수_1차, 2차: 점수_2차},
                   시간: {1차: 시간_1차, 2차: 시간_2차}}
             수준 [1차, 2차]

    구분자 후보를 하나씩 시도해 보고, 그룹이 가장 많이 잡히는 것을 고른다.
    반환: (그룹 dict, 수준 목록, 사용한 구분자)  실패하면 ({}, [], None)
    """
    best = None
    for kind in ("_", ".", "-", "digit"):
        stubs = {}
        for c in cols:
            name = str(c)
            if kind == "digit":
                # 끝의 숫자를 수준으로 본다 (Q1_1 처럼 구분자가 없는 경우)
                mo = re.match(r"^(.*?)[ _.\-]?(\d+)$", name)
                if not mo or not mo.group(1):
                    continue
                stub, suf = mo.group(1), mo.group(2)
            else:
                if kind not in name:
                    continue
                stub, suf = name.rsplit(kind, 1)
                if not stub or not suf:
                    continue
            stubs.setdefault(stub, {})[suf] = c

        groups = {k: v for k, v in stubs.items() if len(v) >= 2}
        if not groups:
            continue
        score = (len(groups), sum(len(v) for v in groups.values()))
        if best is None or score > best[0]:
            levels = []
            for v in groups.values():
                for suf in v:
                    if suf not in levels:
                        levels.append(suf)
            levels = sorted(levels, key=natural_key)
            best = (score, groups, levels, kind)

    if best is None:
        return {}, [], None
    return best[1], best[2], best[3]


def wide_to_long_multi(df, id_cols, groups, levels, index_name="시점",
                       level_labels=None, drop_missing=True):
    """
    값 열을 여러 개 만들면서 가로 → 세로 변환. SPSS 의

        VARSTOCASES /MAKE 점수 FROM 점수_1차 점수_2차
                    /MAKE 시간 FROM 시간_1차 시간_2차
                    /INDEX = 시점.

    에 해당한다.

    groups       : {결과 값 열 이름: {수준: 원본 열 이름}}
    levels       : 수준 순서 (결과 행 순서가 된다)
    index_name   : 수준이 들어갈 새 열 이름 (예: '시점')
    level_labels : {수준: 화면에 찍힐 이름}
    drop_missing : 값 열이 '전부' 비어 있는 행을 버린다.
                   (SPSS /NULL=DROP 과 같은 기준. 하나라도 값이 있으면 남긴다)
    """
    if df is None or df.empty:
        raise ValueError("데이터가 비어 있습니다.")
    id_cols = list(id_cols)
    if not groups:
        raise ValueError("값 그룹이 없습니다. 그룹을 1개 이상 지정하세요.")
    if not levels:
        raise ValueError("수준(시점)이 없습니다.")

    index_name = str(index_name).strip() or "시점"
    vnames = list(groups.keys())

    if len(set(vnames)) != len(vnames):
        raise ValueError("값 열 이름이 서로 겹칩니다. 다르게 지정하세요.")
    if index_name in vnames:
        raise ValueError(f"'{index_name}' 이 값 열 이름과 겹칩니다.")
    clash = set(id_cols) & (set(vnames) | {index_name})
    if clash:
        raise ValueError(
            f"새 열 이름이 유지할 열과 겹칩니다: {sorted(clash)}. 다른 이름을 쓰세요.")

    used = [c for m in groups.values() for c in m.values()]
    missing = [c for c in id_cols + used if c not in df.columns]
    if missing:
        raise ValueError(f"데이터에 없는 열입니다: {missing}")
    both = set(id_cols) & set(used)
    if both:
        raise ValueError(f"유지할 열이 값 그룹에도 들어 있습니다: {sorted(both)}")

    lab = level_labels or {}
    frames = []
    for i, lv in enumerate(levels):
        part = df.loc[:, id_cols].copy()
        part["__row"] = range(len(df))
        part["__lv"] = i
        part[index_name] = lab.get(lv, lv)
        for vname, mapping in groups.items():
            src = mapping.get(lv)
            # 그룹마다 수준이 다 있지 않을 수 있다. 없는 칸은 빈 값으로 둔다.
            part[vname] = df[src].values if src is not None else pd.NA
        frames.append(part)

    out = pd.concat(frames, ignore_index=True)
    # 응답자별로 묶이도록 원래 행 순서 → 수준 순서로 정렬
    out = out.sort_values(["__row", "__lv"], kind="mergesort")

    if drop_missing:
        out = out[out[vnames].notna().any(axis=1)]

    out = out.drop(columns=["__row", "__lv"]).reset_index(drop=True)
    return out.loc[:, id_cols + [index_name] + vnames]


def wide_to_long(df, id_cols, var_cols, var_name="변수", value_name="값",
                 drop_missing=True, add_index=False, rename_map=None):
    """
    가로(wide) → 세로(long) 변환. SPSS 의 VARSTOCASES 에 대응한다.

        ID  국어 수학              ID  변수  값
        1   88   92     ──▶       1   국어  88
        2   70   85                1   수학  92
                                   2   국어  70
                                   2   수학  85

    id_cols      : 그대로 유지할 열
    var_cols     : 세로로 내릴 열 (여기 준 순서가 결과 순서)
    var_name     : 변수명이 들어갈 새 열 이름
    value_name   : 값이 들어갈 새 열 이름
    drop_missing : 값이 빈 행을 버린다 (VARSTOCASES 의 /NULL=DROP 에 해당)
    add_index    : ID 안에서 1,2,3… 순번 열을 추가한다 (/INDEX 유사)
    rename_map   : {원본 열 이름: 변수 열에 찍힐 이름}
    """
    if df is None or df.empty:
        raise ValueError("데이터가 비어 있습니다.")

    id_cols, var_cols = list(id_cols), list(var_cols)
    if not var_cols:
        raise ValueError("세로로 내릴 열을 1개 이상 선택해야 합니다.")

    overlap = set(id_cols) & set(var_cols)
    if overlap:
        raise ValueError(
            f"유지할 열과 내릴 열에 같은 열이 있습니다: {sorted(overlap)}")
    missing = [c for c in id_cols + var_cols if c not in df.columns]
    if missing:
        raise ValueError(f"데이터에 없는 열입니다: {missing}")

    var_name = str(var_name).strip() or "변수"
    value_name = str(value_name).strip() or "값"
    if var_name == value_name:
        raise ValueError("변수 열 이름과 값 열 이름이 같습니다.")
    clash = set(id_cols) & {var_name, value_name}
    if clash:
        raise ValueError(
            f"새 열 이름이 유지할 열과 겹칩니다: {sorted(clash)}. 다른 이름을 쓰세요.")

    rmap = {}
    for k, v in (rename_map or {}).items():
        v = "" if v is None else str(v).strip()
        if v:
            rmap[str(k)] = v

    work = df.loc[:, id_cols + var_cols].copy()
    work["__row"] = range(len(work))       # 원래 행 순서 보존용

    out = work.melt(id_vars=id_cols + ["__row"], value_vars=var_cols,
                    var_name=var_name, value_name=value_name)

    # melt 는 변수별로 뭉쳐서 내놓는다. 응답자별로 묶이도록 되돌린다.
    _order = {c: i for i, c in enumerate(var_cols)}
    out["__vord"] = out[var_name].map(_order)
    out = out.sort_values(["__row", "__vord"], kind="mergesort")

    if drop_missing:
        out = out[out[value_name].notna()]

    if add_index:
        out["순번"] = out.groupby(id_cols, dropna=False, observed=True).cumcount() + 1

    # 변수 이름 바꾸기는 정렬이 끝난 뒤에 적용한다 (정렬 기준이 원본 이름이므로)
    if rmap:
        out[var_name] = out[var_name].map(lambda v: rmap.get(str(v), v))

    out = out.drop(columns=["__row", "__vord"]).reset_index(drop=True)

    # 열 순서: ID … 변수 [순번] 값
    tail = [var_name] + (["순번"] if add_index else []) + [value_name]
    out = out.loc[:, id_cols + tail]
    return out


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
def _arrow_safe(d):
    """
    st.dataframe 표시용 변환.
    한 열에 숫자와 문자가 섞여 있으면(예: 항목 열에 1 과 '국어') Arrow 변환이
    실패하고 Streamlit 이 매번 자동 복구를 시도하면서 로그에 대량의
    ArrowTypeError 를 남긴다. object 열만 문자열로 바꿔 미리 막는다.
    (숫자 열은 그대로 두므로 정렬·서식은 유지된다. 표시용이며 원본은 안 건드린다.)
    """
    out = d.copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype(str)
    return out


EXPAND_LABEL = "여러 열로 펼치기 (변수_1, 변수_2 …)"
DUP_MODES = [EXPAND_LABEL] + list(AGG_FUNCS.keys())

# ── 기본으로 선택될 열 이름 ──────────────────────────────────────────
# 올린 파일에 아래 이름의 열이 있으면 자동으로 선택해 둔다.
# 없으면 그냥 비워 두므로, 다른 형식의 파일도 문제없이 쓸 수 있다.
# 자주 쓰는 형식이 바뀌면 이 세 줄만 고치면 된다.
DEFAULT_ID_COLS = ["panel_id"]      # ① ID 열
DEFAULT_VALUE_COLS = ["intVal"]     # ② 값 열
DEFAULT_KEY_COL = "page_name"       # ③ 기준 열
NO_KEY = "(선택하세요)"


st.set_page_config(page_title="행열 변환", page_icon="🔄", layout="wide")

if not check_password():
    st.stop()

st.title("🔄 행열 변환")

MODE_L2W = "세로 → 가로 (SPSS CASESTOVARS)"
MODE_W2L = "가로 → 세로 (SPSS VARSTOCASES)"
mode = st.radio("변환 방향", [MODE_L2W, MODE_W2L], key="lw_mode", horizontal=True)

if mode == MODE_L2W:
    st.caption("한 응답자의 값이 여러 행에 세로로 쌓여 있는 데이터를, "
               "응답자 1명 = 1행 형태로 펼칩니다.")
else:
    st.caption("한 행에 여러 열로 퍼져 있는 값을, 열 하나당 한 행이 되도록 "
               "세로로 내립니다.")


def is_sav(name):
    return str(name).lower().endswith(".sav")


@st.cache_data(show_spinner=False, max_entries=4)
def list_sheets(data: bytes, name: str):
    if is_sav(name):
        return ["(SPSS)"]
    if name.lower().endswith(".csv"):
        return ["(CSV)"]
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


def read_sav_bytes(data: bytes):
    """
    .sav 바이트를 읽어 (DataFrame, 변수라벨, 값라벨) 을 돌려준다.
    pyreadstat 은 파일 경로만 받으므로 임시 파일을 거친다.
    Windows 에서는 NamedTemporaryFile 을 열어둔 채로 다시 열 수 없으므로
    TemporaryDirectory 를 쓴다.
    """
    import tempfile

    import pyreadstat

    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "in.sav")
        with open(path, "wb") as f:
            f.write(data)
        df, meta = pyreadstat.read_sav(path, apply_value_formats=False)
    labels = dict(getattr(meta, "column_names_to_labels", {}) or {})
    vlabels = dict(getattr(meta, "variable_value_labels", {}) or {})
    return df, labels, vlabels


# max_entries 를 걸지 않으면 큰 파일을 여러 개 올릴 때 원본 바이트와
# 데이터프레임이 캐시에 계속 쌓여 메모리를 다 쓴다 (Cloud 무료 플랜 1GB).
@st.cache_data(show_spinner="파일을 읽는 중…", max_entries=2)
def read_table(data: bytes, name: str, sheet: str, header: int):
    """반환: (DataFrame, 변수라벨 dict, 값라벨 dict). 엑셀/CSV 는 라벨이 빈 dict."""
    if is_sav(name):
        return read_sav_bytes(data)
    if name.lower().endswith(".csv"):
        try:
            import chardet
            enc = chardet.detect(data)["encoding"] or "utf-8"
        except Exception:
            enc = "utf-8"
        for cand in (enc, "utf-8-sig", "cp949", "euc-kr"):
            try:
                return (pd.read_csv(io.BytesIO(data), header=header,
                                    encoding=cand), {}, {})
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 판별하지 못했습니다. utf-8 로 저장 후 올려주세요.")
    return (pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=header), {}, {})


up = st.file_uploader("데이터 파일 (엑셀 / CSV / SPSS .sav)",
                      type=["xlsx", "xlsm", "xls", "csv", "sav"], key="lw_file")
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

_is_sav = is_sav(up.name)
c1, c2 = st.columns([3, 1])
with c1:
    sheet = st.selectbox("시트", sheets, key="lw_sheet",
                         disabled=(len(sheets) == 1))
with c2:
    header_row = st.number_input(
        "머리글 행", min_value=1, max_value=100, value=1, step=1,
        key="lw_header", disabled=_is_sav,
        help="열 이름이 들어 있는 행 번호. 보통 1행입니다. "
             "(.sav 는 해당 없음)")

try:
    df, src_labels, src_vlabels = read_table(
        raw, up.name, sheet, int(header_row) - 1)
except ImportError:
    st.error("SPSS 파일을 읽으려면 pyreadstat 이 필요합니다. "
             "requirements.txt 에 `pyreadstat` 을 추가하고 앱을 다시 배포하세요.")
    st.stop()
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

st.success(f"{len(df):,}행 × {len(df.columns)}열 읽음"
           + (f" · SPSS 메타데이터 (변수라벨 {len(src_labels)}개, "
              f"값라벨 {len(src_vlabels)}개) 를 함께 읽었습니다."
              if _is_sav else ""))
with st.expander("원본 미리보기", expanded=False):
    st.dataframe(_arrow_safe(df.head(20)))
    if _is_sav and src_vlabels:
        st.caption("값 라벨이 있는 변수: "
                   + ", ".join(list(src_vlabels.keys())[:20])
                   + (" …" if len(src_vlabels) > 20 else ""))

cols = list(df.columns)


# ══════════════════════════════════════════════════════════════════════
#  가로 → 세로 화면 (VARSTOCASES). 이 모드면 여기서 끝내고,
#  아래의 세로 → 가로 코드는 실행되지 않는다.
# ══════════════════════════════════════════════════════════════════════
def emit_downloads(res, base_name, prefix, sheet="long",
                   col_labels=None, value_labels=None):
    """
    변환 결과를 엑셀/CSV/SPSS 로 내보내는 공용 블록.
    파일은 변환 시점에 한 번만 만들어 session_state 에 담고, 여기서는
    이미 만들어진 바이트만 버튼에 연결한다.
    prefix : 세션 키 접두사 (모드별로 분리)
    """
    sav = st.session_state.get(f"{prefix}_sav")
    savmsg = st.session_state.get(f"{prefix}_savmsg")
    d1, d2, d3 = st.columns(3)
    d1.download_button(
        "엑셀 다운로드", data=st.session_state[f"{prefix}_xlsx"],
        file_name=f"{base_name}_{sheet}.xlsx", key=f"{prefix}_dx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button(
        "CSV 다운로드", data=st.session_state[f"{prefix}_csv"],
        file_name=f"{base_name}_{sheet}.csv", mime="text/csv",
        key=f"{prefix}_dc")
    if sav:
        d3.download_button(
            "SPSS(.sav) 다운로드", data=sav,
            file_name=f"{base_name}_{sheet}.sav", key=f"{prefix}_ds",
            mime="application/octet-stream")
    else:
        d3.button("SPSS(.sav) 다운로드", disabled=True, key=f"{prefix}_ds_off")

    if savmsg == "no_pyreadstat":
        st.info("SPSS 파일을 만들려면 pyreadstat 이 필요합니다. "
                "requirements.txt 에 `pyreadstat` 을 추가하고 앱을 "
                "다시 배포(Reboot)하세요. 엑셀·CSV 는 그대로 쓸 수 있습니다.")
    elif isinstance(savmsg, str) and savmsg.startswith("error:"):
        st.warning(f"SPSS 파일 생성에 실패했습니다 ({savmsg[6:]}). "
                   "엑셀·CSV 는 정상입니다.")
    elif savmsg:
        st.caption(f"SPSS 변수명 규칙에 맞춰 {len(savmsg)}개 이름을 바꿨습니다: "
                   + ", ".join(f"`{a}` → `{b}`" for a, b in savmsg[:10])
                   + " · 원래 이름은 변수 라벨에 남아 있습니다.")


def store_outputs(res, prefix, sig, sheet="long",
                  col_labels=None, value_labels=None):
    """변환 결과와 다운로드 파일 바이트를 session_state 에 담는다."""
    st.session_state[f"{prefix}_result"] = res
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            res.to_excel(w, sheet_name=sanitize_sheet_name(sheet), index=False)
    except Exception:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            res.to_excel(w, sheet_name=sheet, index=False)
    st.session_state[f"{prefix}_xlsx"] = buf.getvalue()
    st.session_state[f"{prefix}_csv"] = res.to_csv(index=False).encode("utf-8-sig")
    try:
        sav, ch = build_sav(res, sanitize=True, col_labels=col_labels,
                            value_labels=value_labels)
        st.session_state[f"{prefix}_sav"] = sav
        st.session_state[f"{prefix}_savmsg"] = ch
    except ImportError:
        st.session_state[f"{prefix}_sav"] = None
        st.session_state[f"{prefix}_savmsg"] = "no_pyreadstat"
    except Exception as e:
        st.session_state[f"{prefix}_sav"] = None
        st.session_state[f"{prefix}_savmsg"] = f"error:{type(e).__name__}: {e}"
    st.session_state[f"{prefix}_sig"] = sig


W2L_SINGLE = "값 열 1개 — 선택한 열을 한 열에 모음"
W2L_MULTI = "값 열 여러 개 — 반복측정 (SPSS /MAKE)"


def screen_w2l_multi(df, cols, base_name, src_labels, src_vlabels):
    """
    반복측정 가로 → 세로. SPSS 의
        VARSTOCASES /MAKE 점수 FROM 점수_1차 점수_2차
                    /MAKE 시간 FROM 시간_1차 시간_2차
                    /INDEX = 시점.
    에 해당한다. 열 이름에서 그룹과 수준을 자동으로 찾아 표로 보여준다.
    """
    st.divider()
    st.subheader("1. 열 지정")

    sig0 = str(cols)
    if st.session_state.get("wm_colsig") != sig0:
        st.session_state["wm_colsig"] = sig0
        st.session_state["wm_id"] = [c for c in DEFAULT_ID_COLS if c in cols]
        st.session_state.pop("wm_gname", None)
        st.session_state.pop("wm_lname", None)

    id_cols = st.multiselect(
        "① 그대로 유지할 열 (ID 등)", cols, key="wm_id",
        help="응답자ID처럼 모든 수준에서 반복될 열입니다.")

    target = [c for c in cols if c not in id_cols]
    groups, levels, kind = detect_value_groups(target)

    if not groups:
        st.error(
            "열 이름에서 반복 구조를 찾지 못했습니다.\n\n"
            "`점수_1차`, `점수_2차` 처럼 **같은 접두사 + 다른 꼬리표** 형태여야 "
            "자동 인식이 됩니다. 지금 열 이름이 그런 형태가 아니면 위에서 "
            f"**'{W2L_SINGLE}'** 를 선택하세요.")
        st.stop()

    st.caption(f"열 이름을 '{kind if kind != 'digit' else '끝자리 숫자'}' 기준으로 "
               f"나눠 **{len(groups)}개 그룹 × {len(levels)}개 수준**을 찾았습니다. "
               "아래 표에서 이름만 고치면 됩니다.")

    # 어떤 그룹에도 들어가지 못한 열을 반드시 알려준다.
    # (한 수준만 있는 열은 그룹이 되지 못해 조용히 빠진다)
    assigned = {c for mp in groups.values() for c in mp.values()}
    left_out = [c for c in target if c not in assigned]
    if left_out:
        st.warning(
            f"어느 그룹에도 속하지 않아 **결과에서 빠지는 열 {len(left_out)}개**: "
            + ", ".join(f"`{c}`" for c in left_out[:15])
            + (" …" if len(left_out) > 15 else "")
            + "\n\n수준이 하나뿐인 열은 그룹이 되지 못합니다. 유지해야 하는 값이면 "
              "①에서 '그대로 유지할 열' 로 옮기세요.")

    # ── 값 그룹 이름 ─────────────────────────────────────────────
    st.divider()
    st.subheader("2. 값 그룹")
    gsaved = st.session_state.get("wm_gname", {})
    grows = pd.DataFrame({
        "그룹": list(groups.keys()),
        "결과 값 열 이름": [gsaved.get(k, k) for k in groups],
        "포함된 열": [", ".join(str(groups[k][lv]) for lv in levels
                            if lv in groups[k]) for k in groups],
        "사용": [True] * len(groups),
    })
    gkey = "wm_g_" + hashlib.md5(
        "|".join(groups.keys()).encode()).hexdigest()[:10]
    if hasattr(st, "data_editor"):
        ged = st.data_editor(
            grows, key=gkey, hide_index=True, num_rows="fixed",
            column_config={
                "그룹": st.column_config.TextColumn("인식된 접두사", disabled=True),
                "결과 값 열 이름": st.column_config.TextColumn("결과 값 열 이름"),
                "포함된 열": st.column_config.TextColumn("포함된 열", disabled=True),
                "사용": st.column_config.CheckboxColumn("사용"),
            })
    else:
        ged = grows
    st.session_state["wm_gname"] = {
        k: str(v).strip() for k, v in zip(ged["그룹"], ged["결과 값 열 이름"])
        if str(v).strip()}

    # ── 수준(시점) 이름 ──────────────────────────────────────────
    st.subheader("3. 수준 (시점)")
    c1, c2 = st.columns([1, 2])
    with c1:
        index_name = st.text_input("수준이 들어갈 열 이름", value="시점",
                                   key="wm_idxname")
        drop_na = st.checkbox(
            "값이 전부 빈 행은 제외", value=True, key="wm_dropna",
            help="값 열이 모두 비어 있는 행만 버립니다. 하나라도 값이 있으면 "
                 "남깁니다. (SPSS /NULL=DROP 과 같은 기준)")
    with c2:
        lsaved = st.session_state.get("wm_lname", {})
        lrows = pd.DataFrame({
            "수준": [str(x) for x in levels],
            f"'{index_name}' 에 찍힐 값": [lsaved.get(str(x), str(x))
                                      for x in levels],
        })
        ltgt = f"'{index_name}' 에 찍힐 값"
        lkey = "wm_l_" + hashlib.md5(
            ("|".join(map(str, levels)) + index_name).encode()).hexdigest()[:10]
        if hasattr(st, "data_editor"):
            led = st.data_editor(
                lrows, key=lkey, hide_index=True, num_rows="fixed",
                column_config={
                    "수준": st.column_config.TextColumn("인식된 꼬리표",
                                                     disabled=True),
                    ltgt: st.column_config.TextColumn(ltgt)})
        else:
            led = lrows
        st.session_state["wm_lname"] = {
            k: str(v).strip() for k, v in zip(led["수준"], led[ltgt])
            if str(v).strip()}

    # 화면 입력을 실제 인자로 조립
    use_map = dict(zip(ged["그룹"], ged["사용"]))
    name_map = dict(zip(ged["그룹"], ged["결과 값 열 이름"]))
    final_groups = {}
    for g in groups:
        if not use_map.get(g, True):
            continue
        nm = str(name_map.get(g) or g).strip() or g
        final_groups[nm] = groups[g]
    level_labels = {str(k): (str(v).strip() or str(k))
                    for k, v in zip(led["수준"], led[ltgt])}

    # ── 변환 ─────────────────────────────────────────────────────
    st.divider()
    st.subheader("4. 변환")
    if not final_groups:
        st.info("값 그룹을 최소 1개는 '사용' 상태로 두세요.")
        st.stop()

    st.caption(f"예상 결과: {len(df):,}행 × {len(levels)}수준 = "
               f"최대 {len(df) * len(levels):,}행 · "
               f"열 {len(id_cols)} + 1 + {len(final_groups)}개")

    SIG = str((base_name, len(df), tuple(cols), tuple(id_cols),
               tuple(sorted((k, tuple(v.items())) for k, v in final_groups.items())),
               tuple(levels), index_name, tuple(sorted(level_labels.items())),
               bool(drop_na)))

    def clear():
        for k in ("wm_result", "wm_xlsx", "wm_csv", "wm_sav", "wm_savmsg", "wm_sig"):
            st.session_state.pop(k, None)

    if st.button("변환 실행", type="primary", key="wm_run"):
        try:
            with st.spinner("변환 중…"):
                res = wide_to_long_multi(
                    df, id_cols, final_groups, levels, index_name=index_name,
                    level_labels=level_labels, drop_missing=drop_na)
            with st.spinner("다운로드 파일 준비 중…"):
                # 각 값 열의 값 라벨: 그 그룹에 속한 원본 열들의 라벨이
                # 모두 같을 때만 붙인다 (다르면 어느 것을 붙여도 틀린다)
                vl = {}
                for nm, mp in final_groups.items():
                    got = [src_vlabels.get(c) for c in mp.values()
                           if src_vlabels.get(c)]
                    if got and len(got) == len(mp) and all(g == got[0] for g in got):
                        vl[nm] = got[0]
                lab = {c: src_labels.get(c, c) for c in id_cols} if src_labels else {}
                store_outputs(res, "wm", SIG, sheet="long",
                              col_labels=(lab or None), value_labels=(vl or None))
        except ValueError as e:
            clear()
            st.error(str(e))
        except Exception as e:
            clear()
            st.error(f"예상치 못한 오류: {type(e).__name__}: {e}")

    if "wm_result" in st.session_state and st.session_state.get("wm_sig") != SIG:
        clear()
        st.info("설정이 바뀌었습니다. [변환 실행] 을 다시 눌러주세요.")

    res = st.session_state.get("wm_result")
    if res is not None:
        st.success(f"변환 완료 — {len(res):,}행 × {len(res.columns)}열")
        st.dataframe(_arrow_safe(res.head(100)))
        if len(res) > 100:
            st.caption(f"위 표는 앞 100행만 표시합니다. "
                       f"전체 {len(res):,}행은 파일로 받으세요.")
        gone = len(df) * len(levels) - len(res)
        if drop_na and gone > 0:
            st.caption(f"값이 전부 비어 제외된 행: {gone:,}개")
        emit_downloads(res, base_name, "wm", sheet="long")


def screen_wide_to_long(df, cols, base_name, src_labels, src_vlabels):
    kind = st.radio("값 열 구성", [W2L_SINGLE, W2L_MULTI], key="wl_kind",
                    help="반복측정 데이터(점수_1차, 시간_1차 …)는 아래쪽을 "
                         "고르세요. 값 열이 여러 개 만들어집니다.")
    if kind == W2L_MULTI:
        screen_w2l_multi(df, cols, base_name, src_labels, src_vlabels)
        return

    st.divider()
    st.subheader("1. 열 지정")

    # 열 구성이 바뀌면 이전 선택을 정리한다 (위젯이 죽는 것을 막는다)
    sig = str(cols)
    if st.session_state.get("wl_colsig") != sig:
        st.session_state["wl_colsig"] = sig
        st.session_state["wl_id"] = [c for c in DEFAULT_ID_COLS if c in cols]
        st.session_state["wl_var"] = []

    left, right = st.columns(2)
    with left:
        id_cols = st.multiselect(
            "① 그대로 유지할 열 (ID 등)", cols, key="wl_id",
            help="응답자ID처럼 모든 행에 반복될 열입니다.")
    with right:
        rest = [c for c in cols if c not in id_cols]
        var_cols = st.multiselect(
            "② 세로로 내릴 열", rest, key="wl_var",
            help="선택한 순서가 결과 순서가 됩니다.")

    b1, b2, b3 = st.columns([1, 1, 3])
    b1.button("나머지 전체 선택",
              on_click=lambda r=rest: st.session_state.update(wl_var=list(r)))
    b2.button("전체 해제", on_click=lambda: st.session_state.update(wl_var=[]))
    if st.checkbox("자연 정렬로 내리기 (q2 < q10)", key="wl_sort") and var_cols:
        var_cols = sorted(var_cols, key=natural_key)
        b3.caption("정렬 적용: " + ", ".join(var_cols[:8])
                   + (" …" if len(var_cols) > 8 else ""))

    n1, n2 = st.columns(2)
    with n1:
        var_name = st.text_input("③ 변수명이 들어갈 열 이름", value="변수",
                                 key="wl_varname")
    with n2:
        value_name = st.text_input("④ 값이 들어갈 열 이름", value="값",
                                   key="wl_valname")

    o1, o2 = st.columns(2)
    with o1:
        drop_na = st.checkbox(
            "값이 빈 행은 제외", value=True, key="wl_dropna",
            help="SPSS VARSTOCASES 의 /NULL=DROP 과 같습니다. "
                 "끄면 빈 값도 행으로 남습니다.")
    with o2:
        add_idx = st.checkbox("ID별 순번 열 추가", value=False, key="wl_idx",
                              help="SPSS 의 /INDEX 와 비슷합니다.")

    if not var_cols:
        st.info("② 세로로 내릴 열을 1개 이상 선택하세요.")
        st.stop()

    # ── 변수 이름 바꾸기 ──────────────────────────────────────────
    st.divider()
    st.subheader("2. 변수 이름 바꾸기")
    st.caption(f"'{var_name}' 열에 찍힐 이름입니다. 비워 두면 원본 열 이름을 씁니다."
               + (" 원본이 .sav 이므로 변수 라벨을 참고용으로 함께 보여줍니다."
                  if src_labels else ""))

    saved = st.session_state.get("wl_rename", {})
    rows = {"열 이름": [str(c) for c in var_cols]}
    if src_labels:
        rows["원본 변수 라벨"] = [str(src_labels.get(c, "")) for c in var_cols]
    rows[f"'{var_name}' 에 찍힐 이름"] = [saved.get(str(c), "") for c in var_cols]
    rows_df = pd.DataFrame(rows)
    tgt = f"'{var_name}' 에 찍힐 이름"

    if hasattr(st, "data_editor"):
        ed_key = "wl_rn_" + hashlib.md5(
            ("|".join(str(c) for c in var_cols) + var_name).encode()).hexdigest()[:10]

        def reset_names(k=ed_key):
            st.session_state["wl_rename"] = {}
            st.session_state.pop(k, None)

        cfg = {"열 이름": st.column_config.TextColumn("열 이름 (원본)", disabled=True),
               tgt: st.column_config.TextColumn(tgt, help="비워 두면 원본 이름 사용")}
        if src_labels:
            cfg["원본 변수 라벨"] = st.column_config.TextColumn(
                "원본 변수 라벨", disabled=True)
        kw = {}
        if len(rows_df) > 8:
            kw["height"] = min(420, 45 + 35 * len(rows_df))
        edited = st.data_editor(rows_df, key=ed_key, hide_index=True,
                               num_rows="fixed", column_config=cfg, **kw)
        new = dict(zip(edited["열 이름"], edited[tgt]))
    else:
        def reset_names():
            st.session_state["wl_rename"] = {}
            for c in var_cols:
                st.session_state.pop(f"wl_rn_{c}", None)

        new = {}
        for c in var_cols:
            new[str(c)] = st.text_input(str(c), value=saved.get(str(c), ""),
                                        key=f"wl_rn_{c}")

    rename_map = {k: str(v).strip() for k, v in new.items()
                  if v is not None and str(v).strip()}
    st.session_state["wl_rename"] = rename_map
    if rename_map:
        rc1, rc2 = st.columns([4, 1])
        rc1.caption("적용될 이름: "
                    + ", ".join(f"`{k}` → `{v}`"
                                for k, v in list(rename_map.items())[:10])
                    + (" …" if len(rename_map) > 10 else ""))
        rc2.button("이름 전부 초기화", on_click=reset_names)

    # ── 변환 ─────────────────────────────────────────────────────
    st.divider()
    st.subheader("3. 변환")
    st.caption(f"예상 결과: 최대 {len(df):,}행 × {len(var_cols)}열 = "
               f"{len(df) * len(var_cols):,}행"
               + (" (빈 값 제외 시 줄어듭니다)" if drop_na else ""))

    SIG = str((base_name, len(df), tuple(cols), tuple(id_cols), tuple(var_cols),
               var_name, value_name, bool(drop_na), bool(add_idx),
               tuple(sorted(rename_map.items()))))
    KEYS = ("wl_result", "wl_xlsx", "wl_csv", "wl_sav", "wl_savmsg", "wl_sig")

    def clear():
        for k in KEYS:
            st.session_state.pop(k, None)

    if st.button("변환 실행", type="primary", key="wl_run"):
        try:
            with st.spinner("변환 중…"):
                res = wide_to_long(df, id_cols, var_cols, var_name=var_name,
                                   value_name=value_name, drop_missing=drop_na,
                                   add_index=add_idx, rename_map=rename_map)
            with st.spinner("다운로드 파일 준비 중…"):
                st.session_state["wl_result"] = res
                buf = io.BytesIO()
                try:
                    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                        res.to_excel(w, sheet_name=sanitize_sheet_name("long"),
                                     index=False)
                except Exception:
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as w:
                        res.to_excel(w, sheet_name="long", index=False)
                st.session_state["wl_xlsx"] = buf.getvalue()
                st.session_state["wl_csv"] = res.to_csv(index=False).encode("utf-8-sig")

                # 내린 열들의 값 라벨이 모두 같으면 값 열에 그대로 물려준다.
                # 서로 다르면 어느 것을 붙여도 틀리므로 넣지 않는다.
                vl = {}
                got = [src_vlabels.get(c) for c in var_cols if src_vlabels.get(c)]
                if got and len(got) == len(var_cols) and \
                        all(g == got[0] for g in got):
                    vl[value_name] = got[0]
                lab = {c: src_labels.get(c, c) for c in id_cols} if src_labels else {}
                try:
                    sav, ch = build_sav(res, sanitize=True,
                                        col_labels=lab or None,
                                        value_labels=vl or None)
                    st.session_state["wl_sav"] = sav
                    st.session_state["wl_savmsg"] = ch
                except ImportError:
                    st.session_state["wl_sav"] = None
                    st.session_state["wl_savmsg"] = "no_pyreadstat"
                except Exception as e:
                    st.session_state["wl_sav"] = None
                    st.session_state["wl_savmsg"] = f"error:{type(e).__name__}: {e}"
                st.session_state["wl_sig"] = SIG
        except ValueError as e:
            clear()
            st.error(str(e))
        except Exception as e:
            clear()
            st.error(f"예상치 못한 오류: {type(e).__name__}: {e}")

    if "wl_result" in st.session_state and st.session_state.get("wl_sig") != SIG:
        clear()
        st.info("설정이 바뀌었습니다. [변환 실행] 을 다시 눌러주세요.")

    res = st.session_state.get("wl_result")
    if res is not None:
        st.success(f"변환 완료 — {len(res):,}행 × {len(res.columns)}열")
        st.dataframe(_arrow_safe(res.head(100)))
        if len(res) > 100:
            st.caption(f"위 표는 앞 100행만 표시합니다. "
                       f"전체 {len(res):,}행은 파일로 받으세요.")

        dropped = len(df) * len(var_cols) - len(res)
        if drop_na and dropped > 0:
            st.caption(f"값이 비어 제외된 행: {dropped:,}개")

        sav = st.session_state.get("wl_sav")
        savmsg = st.session_state.get("wl_savmsg")
        d1, d2, d3 = st.columns(3)
        d1.download_button("엑셀 다운로드", data=st.session_state["wl_xlsx"],
                           file_name=f"{base_name}_long.xlsx", key="wl_dx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet")
        d2.download_button("CSV 다운로드", data=st.session_state["wl_csv"],
                           file_name=f"{base_name}_long.csv", mime="text/csv",
                           key="wl_dc")
        if sav:
            d3.download_button("SPSS(.sav) 다운로드", data=sav,
                               file_name=f"{base_name}_long.sav", key="wl_ds",
                               mime="application/octet-stream")
        else:
            d3.button("SPSS(.sav) 다운로드", disabled=True, key="wl_ds_off")

        if savmsg == "no_pyreadstat":
            st.info("SPSS 파일을 만들려면 pyreadstat 이 필요합니다. "
                    "requirements.txt 에 `pyreadstat` 을 추가하고 앱을 "
                    "다시 배포(Reboot)하세요.")
        elif isinstance(savmsg, str) and savmsg.startswith("error:"):
            st.warning(f"SPSS 파일 생성에 실패했습니다 ({savmsg[6:]}). "
                       "엑셀·CSV 는 정상입니다.")
        elif savmsg:
            st.caption(f"SPSS 변수명 규칙에 맞춰 {len(savmsg)}개 이름을 바꿨습니다: "
                       + ", ".join(f"`{a}` → `{b}`" for a, b in savmsg[:10])
                       + " · 원래 이름은 변수 라벨에 남아 있습니다.")


if mode == MODE_W2L:
    screen_wide_to_long(df, cols, up.name.rsplit(".", 1)[0],
                        src_labels, src_vlabels)
    st.stop()


# 열 구성이 바뀌면(= 다른 파일/시트) 기본값을 다시 적용한다.
# 위젯 생성 전이라 session_state 대입이 허용된다.
# 이 초기화는 이전 파일의 선택값이 새 파일 옵션에 없어 위젯이 죽는 것도 막아준다.
_colsig = str(cols)
if st.session_state.get("lw_colsig") != _colsig:
    st.session_state["lw_colsig"] = _colsig
    st.session_state["lw_id"] = [c for c in DEFAULT_ID_COLS if c in cols]
    st.session_state["lw_val"] = [c for c in DEFAULT_VALUE_COLS if c in cols]
    st.session_state["lw_key"] = DEFAULT_KEY_COL if DEFAULT_KEY_COL in cols else NO_KEY
    st.session_state.pop("lw_keys", None)      # 변수 선택은 기준 열에 딸린 것이라 초기화

st.divider()
st.subheader("1. 열 지정")

left, right = st.columns(2)
with left:
    id_cols = st.multiselect("① ID 열 — 한 행을 식별하는 기준", cols, key="lw_id",
                             help="응답자ID, 학번 등. 여러 개면 조합으로 식별합니다.")
    key_col = st.selectbox("③ 기준 열 — 가로로 펼칠 변수명이 담긴 열",
                           [NO_KEY] + cols, key="lw_key",
                           help="'항목', '문항번호', '시점' 처럼 변수 이름이 쌓인 열")
with right:
    value_cols = st.multiselect("② 값 열 — 실제 값이 담긴 열", cols, key="lw_val",
                                help="2개 이상이면 '점수_1차' 형태로 조합됩니다.")
    agg_label = st.selectbox(
        "④ 중복값 처리 — 같은 ID에 같은 변수가 2번 이상일 때",
        DUP_MODES, key="lw_agg",
        help="'여러 열로 펼치기' 는 값을 버리지 않습니다. 나머지는 하나만 남깁니다.")

expand_dup = (agg_label == EXPAND_LABEL)
always_sfx, max_occ = False, 0
if expand_dup:
    e1, e2 = st.columns([1, 1])
    with e1:
        always_sfx = st.checkbox(
            "중복이 없는 변수에도 _1 붙이기", value=False, key="lw_sfx",
            help="끄면 중복이 있는 변수만 취미_1, 취미_2 … 가 되고 "
                 "나머지는 성별 처럼 그대로 남습니다.")
    with e2:
        max_occ = int(st.number_input(
            "변수당 최대 개수 (0 = 제한 없음)", min_value=0, max_value=500,
            value=0, step=1, key="lw_maxocc",
            help="예: 3 으로 두면 취미_1~취미_3 까지만 만들고 나머지는 버립니다."))

if key_col == NO_KEY:
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

# ── 변수 이름 바꾸기 ────────────────────────────────────────────────
# 항상 펼쳐진 표로 보여준다. 예전에는 st.expander 안에 넣고 제목에
# "N개 변경됨" 을 표시했는데, 이름을 고칠 때마다 제목이 바뀌어
# Streamlit 이 다른 위젯으로 인식해 패널이 저절로 닫혔다.
rename_map = {}
if keep_keys:
    st.markdown("**변수 이름 바꾸기** (고칠 것만 수정하세요)")
    st.caption("'결과 열 이름' 칸을 고치면 됩니다. 비워 두면 원본 이름을 그대로 씁니다. "
               "펼침 모드에서는 바꾼 이름 뒤에 순번이 붙습니다 (추천 → 추천_1, 추천_2).")

    _saved = st.session_state.get("lw_rename", {})
    _rows = pd.DataFrame({
        "변수": [str(k) for k in keep_keys],
        "결과 열 이름": [_saved.get(str(k), "") for k in keep_keys],
    })

    if hasattr(st, "data_editor"):
        # 변수 목록이 그대로인 동안은 같은 key 를 써서 편집 상태를 유지하고,
        # 목록이 바뀌면 key 가 달라져 새 표로 시작한다. (행 위치가 어긋난 채
        # 이전 편집이 다른 변수에 잘못 적용되는 것을 막는다.)
        _ed_key = "lw_rn_" + hashlib.md5(
            "|".join(str(k) for k in keep_keys).encode("utf-8")).hexdigest()[:10]

        def _reset_names(k=_ed_key):
            st.session_state["lw_rename"] = {}
            st.session_state.pop(k, None)

        # height 는 None 을 허용하지 않으므로 필요할 때만 넘긴다
        _ed_kw = {}
        if len(_rows) > 8:
            _ed_kw["height"] = min(420, 45 + 35 * len(_rows))

        _edited = st.data_editor(
            _rows, key=_ed_key, hide_index=True, num_rows="fixed",
            column_config={
                "변수": st.column_config.TextColumn("변수 (원본)", disabled=True),
                "결과 열 이름": st.column_config.TextColumn(
                    "결과 열 이름", help="비워 두면 원본 이름 사용"),
            }, **_ed_kw)
        _new = dict(zip(_edited["변수"], _edited["결과 열 이름"]))
    else:
        # 아주 오래된 Streamlit 대비 (data_editor 이전 버전)
        def _reset_names():
            st.session_state["lw_rename"] = {}
            for _k in keep_keys:
                st.session_state.pop(f"lw_rn_{_k}", None)

        _new = {}
        for _k in keep_keys:
            _new[str(_k)] = st.text_input(
                str(_k), value=_saved.get(str(_k), ""), key=f"lw_rn_{_k}")

    rename_map = {k: str(v).strip() for k, v in _new.items()
                  if v is not None and str(v).strip()}
    st.session_state["lw_rename"] = rename_map

    if rename_map:
        rc1, rc2 = st.columns([4, 1])
        rc1.caption("적용될 이름: "
                    + ", ".join(f"`{k}` → `{v}`"
                                for k, v in list(rename_map.items())[:10])
                    + (" …" if len(rename_map) > 10 else ""))
        rc2.button("이름 전부 초기화", on_click=_reset_names)

if id_cols and keep_keys:
    try:
        dup, dup_total = find_duplicate_cells(
            df[key_series.isin(keep_keys)], id_cols, key_col,
            normalize_keys=norm_keys)
    except Exception:
        dup, dup_total = pd.DataFrame(), 0
    if dup_total:
        if expand_dup:
            st.info(f"같은 ID에 같은 변수가 중복된 칸이 {len(dup):,}곳 있습니다 "
                    f"(총 {dup_total:,}행). **여러 열로 펼쳐지며 값은 보존됩니다.**")
        else:
            st.warning(f"같은 ID에 같은 변수가 중복된 칸이 {len(dup):,}곳 있습니다 "
                       f"(총 {dup_total:,}행). 현재 설정은 **{agg_label}** 이므로 "
                       f"각 칸에서 하나만 남고 나머지는 버려집니다.")
        with st.expander("중복 내역 보기"):
            st.dataframe(_arrow_safe(dup))

# 정규화 때문에 서로 다른 원본 값이 한 열로 합쳐지는 경우를 명시적으로 알린다.
# 위의 중복 경고만으로는 원인이 정규화라는 것을 알 수 없다.
if norm_keys and keep_keys:
    try:
        _pair = pd.DataFrame({"o": df[key_col], "n": key_series})
        _pair = _pair[_pair["n"].isin(keep_keys)]
        _grp = _pair.groupby("n", dropna=False)["o"].unique()
        _merged = [(k, v) for k, v in _grp.items() if len(v) > 1]
    except Exception:
        _merged = []
    if _merged:
        st.warning(
            f"서로 다른 값 {sum(len(v) for _, v in _merged)}개가 정규화되어 "
            f"{len(_merged)}개 열로 합쳐집니다. 합쳐진 값들은 위의 중복값 처리 "
            + ("설정에 따라 **여러 열로 펼쳐집니다.**" if expand_dup
               else f"규칙(**{agg_label}**)에 따라 하나만 남습니다."))
        with st.expander("합쳐지는 값 보기"):
            st.dataframe(_arrow_safe(pd.DataFrame({
                "결과 열 이름": [str(k) for k, _ in _merged],
                "합쳐지는 원본 값": [", ".join(repr(x) for x in v) for _, v in _merged],
            })))
            st.caption("따옴표와 공백까지 보이도록 원본 그대로 표시했습니다. "
                       "의도한 것이 아니라면 '변수명 정규화'를 끄거나 "
                       "엑셀에서 값을 통일한 뒤 다시 올려주세요.")

st.divider()
st.subheader("3. 변환")

if not id_cols or not value_cols or not keep_keys:
    st.info("ID 열, 값 열, 펼칠 변수를 모두 지정하면 변환 버튼이 활성화됩니다.")
    st.stop()

if expand_dup:
    st.caption(f"예상 결과: ID {len(id_cols)}열 + 값 {len(value_cols)}개 × "
               f"변수 {len(keep_keys)}개 (중복이 있는 변수는 개수만큼 늘어납니다)")
else:
    st.caption(f"예상 결과: ID {len(id_cols)}열 + 값 {len(value_cols)}개 × "
               f"변수 {len(keep_keys)}개 = 총 "
               f"{len(id_cols) + len(value_cols) * len(keep_keys)}열")

# 결과에 영향을 주는 모든 설정의 지문. 변환 시점의 지문을 함께 저장해 두고
# 현재 지문과 다르면 결과를 폐기한다. 이게 없으면 변환 후 설정을 바꿨을 때
# 화면의 설정과 표/다운로드 내용이 어긋난 채로 남아 옛 파일을 받게 된다.
_SIG = str((up.name, len(raw), sheet, int(header_row), tuple(id_cols), key_col,
            tuple(value_cols), tuple(keep_keys), agg_label, bool(norm_keys),
            bool(always_sfx), int(max_occ),
            tuple(sorted(rename_map.items()))))
_KEYS = ("lw_result", "lw_xlsx", "lw_csv", "lw_sav", "lw_savmsg",
         "lw_sig", "lw_base")


def _clear_result():
    for k in _KEYS:
        st.session_state.pop(k, None)


def _build_xlsx(res):
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            res.to_excel(w, sheet_name=sanitize_sheet_name("wide"), index=False)
    except Exception:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            res.to_excel(w, sheet_name="wide", index=False)
    return buf.getvalue()


if st.button("변환 실행", type="primary"):
    try:
        with st.spinner("변환 중…"):
            res = long_to_wide(
                df, id_cols=id_cols, key_col=key_col, value_cols=value_cols,
                keep_keys=keep_keys,
                aggfunc=AGG_FUNCS.get(agg_label, "first"),
                normalize_keys=norm_keys,
                expand_duplicates=expand_dup,
                always_suffix=always_sfx,
                max_occurrences=max_occ,
                rename_map=rename_map)
        # 다운로드 파일은 여기서 딱 한 번만 만든다. 버튼 밖에서 만들면
        # 체크박스 하나 누를 때마다 xlsx 를 처음부터 다시 쓴다.
        with st.spinner("다운로드 파일 준비 중…"):
            st.session_state["lw_result"] = res
            st.session_state["lw_xlsx"] = _build_xlsx(res)
            st.session_state["lw_csv"] = res.to_csv(index=False).encode("utf-8-sig")
            # SPSS .sav 는 pyreadstat 이 있을 때만. 없으면 안내만 남기고
            # 엑셀/CSV 는 정상적으로 제공한다.
            try:
                # 원본이 .sav 였다면, 각 결과 열에 그 출처 값 열의
                # 값 라벨을 그대로 물려준다 (1=매우 불만족 …)
                _src = res.attrs.get("col_source") or {}
                _vl = {c: src_vlabels[v] for c, v in _src.items()
                       if v in src_vlabels}
                _sav, _savch = build_sav(res, sanitize=True, value_labels=_vl)
                st.session_state["lw_sav"] = _sav
                st.session_state["lw_savmsg"] = _savch
            except ImportError:
                st.session_state["lw_sav"] = None
                st.session_state["lw_savmsg"] = "no_pyreadstat"
            except Exception as _e:
                st.session_state["lw_sav"] = None
                st.session_state["lw_savmsg"] = f"error:{type(_e).__name__}: {_e}"
            st.session_state["lw_sig"] = _SIG
            st.session_state["lw_base"] = up.name.rsplit(".", 1)[0]
    except ValueError as e:
        _clear_result()
        st.error(str(e))
    except Exception as e:
        _clear_result()
        st.error(f"예상치 못한 오류: {type(e).__name__}: {e}")

# 변환 후 설정이 바뀌었으면 결과를 버린다
if "lw_result" in st.session_state and st.session_state.get("lw_sig") != _SIG:
    _clear_result()
    st.info("설정이 바뀌었습니다. [변환 실행] 을 다시 눌러주세요.")

result = st.session_state.get("lw_result")
if result is not None:
    st.success(f"변환 완료 — {len(result):,}행 × {len(result.columns)}열")
    st.dataframe(_arrow_safe(result.head(100)))
    if len(result) > 100:
        st.caption(f"위 표는 앞 100행만 표시합니다. 전체 {len(result):,}행은 파일로 받으세요.")

    empty_cols = [c for c in result.columns if result[c].isna().all()]
    if empty_cols:
        st.warning(f"값이 전부 비어 있는 열: {empty_cols}")

    _occ = result.attrs.get("occ_max") or {}
    _multi = {k: v for k, v in _occ.items() if v > 1}
    if _multi:
        _top = sorted(_multi.items(), key=lambda x: -x[1])
        st.info("여러 열로 펼쳐진 변수: "
                + ", ".join(f"`{rename_map.get(str(k), str(k))}` → {v}개"
                            for k, v in _top[:15])
                + (f" 외 {len(_top) - 15}개" if len(_top) > 15 else ""))

    _tr = result.attrs.get("truncated_rows") or 0
    if _tr:
        st.warning(f"'변수당 최대 개수' 제한으로 {_tr:,}개 값이 버려졌습니다. "
                   "모두 살리려면 제한을 0으로 두세요.")

    _rn = result.attrs.get("renamed_columns") or []
    if _rn:
        _shown = ", ".join(f"`{a}` → `{b}`" for a, b in _rn[:20])
        st.warning(
            f"열 이름이 겹쳐서 {len(_rn)}개를 자동으로 바꿨습니다: {_shown}"
            + (f" 외 {len(_rn) - 20}건" if len(_rn) > 20 else "")
            + "\n\n이름 바꾸기에서 서로 같은 이름을 넣었거나, ID 열 이름과 같은 "
              "변수가 있거나, 값 열 2개 이상을 고를 때 조합된 이름이 같아지면 "
              "발생합니다. 데이터는 손실되지 않았습니다.")

    base = st.session_state.get("lw_base", "result")
    _sav = st.session_state.get("lw_sav")
    _savmsg = st.session_state.get("lw_savmsg")

    d1, d2, d3 = st.columns(3)
    d1.download_button(
        "엑셀 다운로드", data=st.session_state["lw_xlsx"],
        file_name=f"{base}_wide.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button(
        "CSV 다운로드", data=st.session_state["lw_csv"],
        file_name=f"{base}_wide.csv", mime="text/csv")
    if _sav:
        d3.download_button(
            "SPSS(.sav) 다운로드", data=_sav,
            file_name=f"{base}_wide.sav", mime="application/octet-stream")
    else:
        d3.button("SPSS(.sav) 다운로드", disabled=True)

    if _savmsg == "no_pyreadstat":
        st.info("SPSS 파일을 만들려면 pyreadstat 이 필요합니다. "
                "requirements.txt 에 `pyreadstat` 한 줄을 추가하고 앱을 "
                "다시 배포(Reboot)하세요. 엑셀·CSV 는 그대로 사용할 수 있습니다.")
    elif isinstance(_savmsg, str) and _savmsg.startswith("error:"):
        st.warning(f"SPSS 파일 생성에 실패했습니다 ({_savmsg[6:]}). "
                   "엑셀·CSV 는 정상입니다.")
    elif _savmsg:
        _sh = ", ".join(f"`{a}` → `{b}`" for a, b in _savmsg[:10])
        st.caption(
            f"SPSS 변수명 규칙(공백·특수문자 불가, 문자로 시작, 64바이트 이내)에 "
            f"맞춰 {len(_savmsg)}개 이름을 바꿨습니다: {_sh}"
            + (" …" if len(_savmsg) > 10 else "")
            + " · 원래 이름은 SPSS 변수 라벨에 그대로 남아 있습니다. "
              "(엑셀·CSV 는 원래 이름을 씁니다)")

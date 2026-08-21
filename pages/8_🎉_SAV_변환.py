"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : pages/8___SAV_변환.py                                            ║
║  위치   : pages/ 폴더 안                                                   ║
║  버전   : 1.0                                                             ║
║                                                                          ║
║  ★ 단일 파일 버전 ★                                                        ║
║  utils.py 를 전혀 수정하지 않아도 동작합니다.                                 ║
║  utils.py 가 있으면 check_password 를 재사용하고,                            ║
║  없거나 구버전이면 이 파일 안의 동일한 구현으로 자동 대체합니다.                ║
║                                                                          ║
║  필요 패키지 : streamlit, pandas, openpyxl, pyreadstat                     ║
║               (chardet 은 있으면 쓰고 없으면 건너뜁니다)                      ║
║               pyreadstat 이 없으면 안내만 띄우고 앱은 죽지 않습니다.           ║
╚══════════════════════════════════════════════════════════════════════════╝

엑셀/CSV 표를 SPSS 파일(.sav)로 저장한다.

  ┌ 하는 일 ────────────────────────────────────────────────┐
  │ 1. 열 이름을 SPSS 규칙에 맞게 정리 (한글 이름 → 변수 라벨로) │
  │ 2. 열마다 숫자/문자/코드화 여부와 측정 수준을 지정          │
  │ 3. 문자 응답을 1,2,3… 숫자로 바꾸고 값 라벨을 붙임          │
  │ 4. 결측 코드(99, 999 등)를 SPSS 결측값으로 등록            │
  └────────────────────────────────────────────────────────┘
"""

import io
import os
import re
import hashlib
import tempfile
import unicodedata

import pandas as pd
import streamlit as st

PAGE_VERSION = "1.0"

# ==============================================================================
# 0. 선택적 의존성
# ==============================================================================
try:
    import pyreadstat
    HAS_PYREADSTAT = True
except Exception:                                    # noqa: BLE001
    pyreadstat = None
    HAS_PYREADSTAT = False

try:
    import chardet
    HAS_CHARDET = True
except Exception:                                    # noqa: BLE001
    chardet = None
    HAS_CHARDET = False

try:
    import utils as _u
except Exception:                                    # noqa: BLE001
    _u = None


# ==============================================================================
# 1. 로그인 (utils 에 있으면 그것을, 없으면 아래 구현을 쓴다)
# ==============================================================================
def _fallback_check_password() -> bool:
    """secrets 에 password 가 없으면 잠금을 걸지 않는다(로컬 실행 편의)."""
    try:
        real = st.secrets["password"]
    except Exception:                                # noqa: BLE001
        return True

    if st.session_state.get("password_correct"):
        return True

    def _same(a: str, b: str) -> bool:
        # hmac.compare_digest 는 비ASCII 입력에서 터진다 → sha256 으로 비교
        ha = hashlib.sha256(str(a).encode("utf-8")).hexdigest()
        hb = hashlib.sha256(str(b).encode("utf-8")).hexdigest()
        return ha == hb

    def _submit():
        st.session_state["password_correct"] = _same(
            st.session_state.get("password", ""), real
        )
        st.session_state["password"] = ""

    st.text_input("비밀번호", type="password", key="password", on_change=_submit)
    if st.session_state.get("password_correct") is False:
        st.error("비밀번호가 맞지 않습니다.")
    return False


check_password = getattr(_u, "check_password", None) or _fallback_check_password


# ==============================================================================
# 1-b. Streamlit 버전 호환 (width="stretch" 는 최근 버전에서 추가된 인자)
# ==============================================================================
import inspect as _inspect


def _wide(fn):
    """가로를 꽉 채우는 인자를 버전에 맞게 돌려준다."""
    try:
        params = _inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return {}
    if "width" in params:
        return {"width": "stretch"}
    if "use_container_width" in params:
        return {"use_container_width": True}
    return {}


# ==============================================================================
# 2. 화면 기본 설정
# ==============================================================================
st.set_page_config(page_title="SAV 변환", page_icon="💾", layout="wide")

if not check_password():
    st.stop()

st.title("💾 엑셀 → SPSS(.sav) 변환")
st.caption(
    "엑셀이나 CSV 표를 SPSS에서 바로 열 수 있는 .sav 파일로 만듭니다. "
    "한글 열 이름은 변수 라벨로 옮기고, 문자로 적힌 응답은 숫자 코드와 값 라벨로 "
    "바꿀 수 있습니다."
)

if not HAS_PYREADSTAT:
    st.error(
        "`pyreadstat` 패키지가 설치돼 있지 않아 .sav 파일을 만들 수 없습니다.\n\n"
        "`requirements.txt` 에 아래 한 줄을 추가하고 다시 배포하세요.\n\n"
        "```\npyreadstat\n```"
    )
    st.stop()


# ==============================================================================
# 3. 파일 읽기
# ==============================================================================
@st.cache_data(show_spinner=False, max_entries=5)
def list_sheets(data: bytes, name: str):
    if name.lower().endswith(".csv"):
        return ["(CSV)"]
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner="파일을 읽는 중…", max_entries=5)
def read_table(data: bytes, name: str, sheet: str, header_row: int) -> pd.DataFrame:
    """header_row 는 1부터 세는 화면 표기값."""
    hdr = header_row - 1
    if name.lower().endswith(".csv"):
        enc = None
        if HAS_CHARDET:
            guess = chardet.detect(data[:200_000]).get("encoding")
            enc = guess
        for cand in [enc, "utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]:
            if not cand:
                continue
            try:
                return pd.read_csv(io.BytesIO(data), header=hdr, encoding=cand,
                                   dtype=object, keep_default_na=True)
            except Exception:                        # noqa: BLE001, PERF203
                continue
        raise ValueError("CSV 인코딩을 판별하지 못했습니다.")
    return pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=hdr,
                         dtype=object)


# ==============================================================================
# 4. SPSS 변수명 정리
# ==============================================================================
_SPSS_RESERVED = {
    "ALL", "AND", "BY", "EQ", "GE", "GT", "LE", "LT", "NE", "NOT", "OR",
    "TO", "WITH",
}

TYPE_AUTO, TYPE_NUM, TYPE_STR, TYPE_CODE, TYPE_DATE = (
    "자동", "숫자", "문자", "문자→코드화", "날짜")
MEASURE_MAP = {"명목": "nominal", "서열": "ordinal", "연속": "scale"}


def to_spss_name(raw: str, order: int, used: set) -> str:
    """SPSS 변수명 규칙(영문/숫자/밑줄, 첫 글자는 영문, 64바이트)에 맞게 고친다."""
    s = unicodedata.normalize("NFKC", str(raw)).strip()
    s = re.sub(r"[^0-9A-Za-z_]", "_", s)
    s = re.sub(r"_{2,}", "_", s).strip("_")

    if not s or not re.match(r"^[A-Za-z]", s):
        s = f"V{order}" if not s else f"V{order}_{s}"

    s = s[:60]                      # 뒤에 중복 꼬리표가 붙을 여유를 남긴다
    if s.upper() in _SPSS_RESERVED:
        s = s + "_"

    base, n = s, 2
    while s.lower() in used:
        tail = f"_{n}"
        s = base[: 60 - len(tail)] + tail
        n += 1
    used.add(s.lower())
    return s


def looks_numeric(sr: pd.Series) -> bool:
    """빈칸을 뺀 값이 전부 숫자로 읽히면 True."""
    v = sr.dropna()
    v = v[v.astype(str).str.strip() != ""]
    if v.empty:
        return True
    return pd.to_numeric(v, errors="coerce").notna().all()


def _as_text(x):
    """빈칸·결측은 None, 나머지는 앞뒤 공백을 없앤 문자열로."""
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    s = str(x).strip()
    return s or None


def text_list(sr: pd.Series) -> list:
    """pandas 버전마다 map 의 결측 처리가 달라서 파이썬 리스트로 직접 돈다."""
    return [_as_text(x) for x in sr.tolist()]


_DATE_RE = re.compile(r"^\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}")


def looks_like_datetime_text(sr: pd.Series) -> bool:
    """‘2026-08-20 오후 5:39:19’ 처럼 글자로 적힌 날짜인지."""
    vals = [x for x in text_list(sr) if x is not None][:200]
    if not vals:
        return False
    hit = sum(1 for v in vals if _DATE_RE.match(v))
    return hit >= len(vals) * 0.8


def to_datetime_series(sr: pd.Series) -> pd.Series:
    """오전/오후 표기를 포함한 날짜 문자열을 최대한 살려서 변환한다."""
    import warnings

    s = pd.Series(text_list(sr), dtype=object)
    have = int(s.notna().sum())
    best = pd.Series([pd.NaT] * len(s))

    def _try(series, **kw):
        nonlocal best
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            try:
                cand = pd.to_datetime(series, errors="coerce", **kw)
            except Exception:                        # noqa: BLE001
                return False
        if int(cand.notna().sum()) > int(best.notna().sum()):
            best = cand
        return int(best.notna().sum()) >= have

    # 오전/오후가 섞여 있으면 그것부터 (dateutil 은 한글 표기를 모른다)
    if s.dropna().astype(str).str.contains("오전|오후").any():
        s2 = s.map(lambda x: None if x is None
                   else x.replace("오전", "AM").replace("오후", "PM"))
        for fmt in ("%Y-%m-%d %p %I:%M:%S", "%Y-%m-%d %p %I:%M",
                    "%Y/%m/%d %p %I:%M:%S", "%Y.%m.%d %p %I:%M:%S",
                    "%Y-%m-%d %I:%M:%S %p"):
            if _try(s2, format=fmt):
                return best

    _try(s)
    return best


def is_datetime_col(sr: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(sr):
        return True
    v = sr.dropna()
    if v.empty:
        return False
    import datetime as _dt
    if all(isinstance(x, (pd.Timestamp, _dt.datetime, _dt.date)) for x in v.head(50)):
        return True
    return looks_like_datetime_text(sr)


def has_leading_zero(sr: pd.Series) -> bool:
    """‘0001’ 처럼 앞자리 0 이 의미를 갖는 코드인지. 숫자로 바꾸면 0 이 날아간다."""
    for x in sr.dropna().tolist()[:1000]:
        if re.fullmatch(r"0\d+", str(x).strip()):
            return True
    return False


def auto_kind(sr: pd.Series, suggest_code: bool = False) -> str:
    """열 하나를 보고 유형을 추천한다."""
    if is_datetime_col(sr):
        return TYPE_DATE
    if has_leading_zero(sr):
        return TYPE_STR
    if looks_numeric(sr):
        return TYPE_NUM
    if suggest_code:
        vals = [v for v in text_list(sr) if v is not None]
        uniq = len(set(vals))
        # 값이 몇 종류 안 되고, 응답자 수보다 확실히 적을 때만 코드화를 권한다.
        # (ID 처럼 행마다 다른 값은 코드화하면 라벨만 수백 개 생긴다)
        if vals and 2 <= uniq <= 30 and uniq <= max(2, len(vals) * 0.6):
            if max(len(v) for v in vals) <= 40:
                return TYPE_CODE
    return TYPE_STR


def auto_measure(sr: pd.Series, kind: str) -> str:
    """측정 수준 추천. 값이 몇 개 안 되는 정수는 척도가 아니라 범주로 본다."""
    if kind in (TYPE_STR, TYPE_CODE):
        return "명목"
    if kind == TYPE_DATE:
        return "연속"
    v = pd.to_numeric(sr, errors="coerce").dropna()
    if v.empty:
        return "명목"
    if ((v % 1) == 0).all() and v.nunique() <= 10:
        return "명목"
    return "연속"


# ==============================================================================
# 5. .sav 만들기 (순수 함수 — 화면과 분리해서 테스트할 수 있게)
# ==============================================================================
def parse_codes(text) -> list:
    """‘99, 999’ 같은 문자열을 [99.0, 999.0] 으로. 숫자가 아니면 버린다."""
    if text is None:
        return []
    out = []
    for tok in re.split(r"[,\s]+", str(text).strip()):
        if not tok:
            continue
        try:
            out.append(float(tok))
        except ValueError:
            continue
    return out


def num_format(sr: pd.Series) -> str:
    """전부 정수면 소수점을 없앤다(SPSS에서 1.00 대신 1 로 보이게)."""
    v = sr.dropna()
    if v.empty:
        return "F8.0"
    width = max(8, len(str(int(abs(v).max()))) + 2)
    if ((v % 1) == 0).all():
        return f"F{width}.0"
    return f"F{width}.2"


SAVE_ALL, SAVE_NUM, SAVE_STR = "숫자＋문자", "숫자만", "문자만"


def build_sav(df: pd.DataFrame,
              spec: pd.DataFrame,
              file_label: str = "",
              save_mode: str = SAVE_ALL,
              always: tuple = ()) -> tuple:
    """
    df    : 원본 표
    spec  : 열마다 한 행. 필요한 칸 —
            포함 / 원본열 / 변수명 / 변수라벨 / 유형 / 측정 / 결측코드
    save_mode : 숫자＋문자 / 숫자만 / 문자만
    always    : 저장 범위와 상관없이 남길 열(ID 등)
    반환  : (sav bytes, 값라벨 dict, 경고 목록, 제외된 열 목록)
    """
    df = df.reset_index(drop=True)
    out, labels, value_labels, measures = {}, {}, {}, {}
    formats, missing_ranges = {}, {}
    warns, skipped = [], []

    for _, row in spec.iterrows():
        if not row["포함"]:
            continue
        src, name = row["원본열"], str(row["변수명"]).strip()
        if not name:
            warns.append(f"‘{src}’ 는 변수명이 비어 있어 건너뜁니다.")
            continue

        sr = df[src]
        kind = row["유형"]

        if kind == TYPE_AUTO:
            kind = auto_kind(sr)

        # 저장 범위 적용 (항상 남길 열은 건너뛰지 않는다)
        if str(src) not in set(map(str, always)):
            if save_mode == SAVE_NUM and kind == TYPE_STR:
                skipped.append(str(src))
                continue
            if save_mode == SAVE_STR:
                if kind in (TYPE_NUM, TYPE_DATE):
                    skipped.append(str(src))
                    continue
                if kind == TYPE_CODE:                # 코드 대신 원래 글자를 남긴다
                    kind = TYPE_STR

        if kind == TYPE_NUM:
            conv = pd.to_numeric(sr, errors="coerce")
            bad = int(sr.notna().sum() - conv.notna().sum())
            if bad:
                warns.append(f"‘{src}’ 에서 숫자로 못 읽은 값 {bad}개는 빈칸이 됩니다.")
            out[name] = conv.reset_index(drop=True).astype("float64")
            formats[name] = num_format(conv)

        elif kind == TYPE_DATE:
            conv = to_datetime_series(sr)
            miss = int(len(text_list(sr)) - sum(v is None for v in text_list(sr))
                       - conv.notna().sum())
            if miss > 0:
                warns.append(f"‘{src}’ 에서 날짜로 못 읽은 값 {miss}개는 빈칸이 됩니다.")
            out[name] = conv.reset_index(drop=True)

        elif kind == TYPE_CODE:
            txt = text_list(sr)
            cats = sorted({v for v in txt if v is not None})
            if len(cats) > 200:
                warns.append(
                    f"‘{src}’ 는 서로 다른 값이 {len(cats)}개라 코드화하면 "
                    "라벨이 지나치게 많아집니다. 문자 그대로 저장했습니다."
                )
                out[name] = pd.Series(["" if v is None else v for v in txt],
                                      dtype=object)
            else:
                code = {c: i + 1 for i, c in enumerate(cats)}
                out[name] = pd.Series(
                    [None if v is None else float(code[v]) for v in txt],
                    dtype="float64")
                value_labels[name] = {float(i): c for c, i in code.items()}
                formats[name] = "F8.0"

        else:                                        # TYPE_STR
            out[name] = pd.Series(
                [("" if v is None else v) for v in text_list(sr)], dtype=object)

        labels[name] = str(row["변수라벨"]).strip() or str(src)
        measures[name] = MEASURE_MAP.get(row["측정"], "nominal")

        codes = parse_codes(row.get("결측코드", ""))
        if codes:
            if pd.api.types.is_numeric_dtype(out[name]):
                missing_ranges[name] = [{"lo": c, "hi": c} for c in codes]
            else:
                warns.append(f"‘{src}’ 는 숫자 열이 아니라 결측 코드를 건너뜁니다.")

    if not out:
        raise ValueError(
            "내보낼 열이 하나도 없습니다."
            + (f" (저장 범위를 ‘{save_mode}’ 로 두어 {len(skipped)}개가 빠졌습니다)"
               if skipped else "")
        )

    res = pd.DataFrame(out)

    tmp = tempfile.NamedTemporaryFile(suffix=".sav", delete=False)
    tmp.close()
    try:
        pyreadstat.write_sav(
            res, tmp.name,
            file_label=file_label[:64] if file_label else None,
            column_labels=[labels[c] for c in res.columns],
            variable_value_labels=value_labels or None,
            variable_measure=measures or None,
            variable_format=formats or None,
            missing_ranges=missing_ranges or None,
        )
        with open(tmp.name, "rb") as f:
            data = f.read()
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass

    return data, value_labels, warns, skipped


# ==============================================================================
# 5-b. 여러 시트 합치기
# ==============================================================================
def key_series(sr: pd.Series) -> pd.Series:
    """연결 키를 문자로 통일한다. 시트마다 50 / '50' / 50.0 으로 읽히는 걸 막는다."""
    def one(x):
        v = _as_text(x)
        if v is None:
            return None
        if re.fullmatch(r"-?\d+\.0+", v):            # 50.0 → 50
            v = v.split(".")[0]
        return v
    return pd.Series([one(x) for x in sr.tolist()], dtype=object)


def merge_side(frames: dict, key: str, keep_all: bool) -> tuple:
    """시트를 옆으로 붙인다(같은 응답자, 다른 문항)."""
    notes, base, base_name = [], None, None
    seen_cols = set()

    for name, d in frames.items():
        if key not in d.columns:
            raise ValueError(f"‘{name}’ 시트에 연결 열 ‘{key}’ 가 없습니다.")
        d = d.copy()
        d["__key__"] = key_series(d[key])

        blank_key = int(d["__key__"].isna().sum())
        if blank_key:
            notes.append(f"‘{name}’ 시트에서 {key} 가 빈 행 {blank_key}개를 뺐습니다.")
            d = d[d["__key__"].notna()]

        dup = int(d["__key__"].duplicated().sum())
        if dup:
            raise ValueError(
                f"‘{name}’ 시트에 {key} 가 겹치는 행이 {dup}개 있습니다. "
                "옆으로 붙이려면 한 사람이 한 행이어야 합니다."
            )

        if base is None:
            base, base_name = d, name
            seen_cols = {c for c in d.columns if c != "__key__"}
            notes.append(f"‘{name}’ 시트 {len(d):,}행 × {len(d.columns)-1}열로 시작")
            continue

        drop = [c for c in d.columns
                if c != "__key__" and c in seen_cols]
        overlap = [c for c in drop if c != key]
        if overlap:
            # 같은 이름의 열이 양쪽에 있을 때, 값까지 같은지 확인한다
            chk = base[["__key__"] + overlap].merge(
                d[["__key__"] + overlap], on="__key__", how="inner",
                suffixes=("__L", "__R"))
            diff_cols = []
            for c in overlap:
                l = chk[f"{c}__L"].map(_as_text)
                r = chk[f"{c}__R"].map(_as_text)
                n = int((l != r).sum())
                if n:
                    diff_cols.append(f"{c}({n}건)")
            head = (f"‘{name}’ 시트에도 있는 열 "
                    f"{', '.join(map(str, overlap[:8]))}"
                    f"{' 외 %d개' % (len(overlap)-8) if len(overlap) > 8 else ''} 는 "
                    f"‘{base_name}’ 쪽 값을 씁니다.")
            if diff_cols:
                head += ("  ⚠ 값이 다른 열: " + ", ".join(diff_cols[:8])
                         + " — 어느 쪽이 맞는지 확인하세요.")
            notes.append(head)
        d = d.drop(columns=drop)

        only_l = len(set(base["__key__"]) - set(d["__key__"]))
        only_r = len(set(d["__key__"]) - set(base["__key__"]))
        if only_l or only_r:
            notes.append(
                f"‘{name}’ 시트와 대조: 앞쪽에만 있는 응답자 {only_l}명, "
                f"이 시트에만 있는 응답자 {only_r}명"
            )

        base = base.merge(d, on="__key__", how="outer" if keep_all else "inner")
        seen_cols |= {c for c in d.columns if c != "__key__"}
        notes.append(f"‘{name}’ 붙인 뒤 {len(base):,}행 × {len(base.columns)-1}열")

    return base.drop(columns="__key__"), notes


def merge_stack(frames: dict, add_sheet_col: bool) -> tuple:
    """시트를 위아래로 잇는다(같은 문항, 다른 응답자)."""
    notes, parts = [], []
    all_cols = []
    for name, d in frames.items():
        for c in d.columns:
            if c not in all_cols:
                all_cols.append(c)
    for name, d in frames.items():
        missing = [c for c in all_cols if c not in d.columns]
        if missing:
            notes.append(
                f"‘{name}’ 시트에 없는 열 {len(missing)}개는 빈칸으로 채웁니다 — "
                + ", ".join(map(str, missing[:8]))
            )
        d = d.reindex(columns=all_cols)
        if add_sheet_col:
            d.insert(0, "시트", name)
        parts.append(d)
        notes.append(f"‘{name}’ 시트 {len(d):,}행")
    out = pd.concat(parts, ignore_index=True)
    notes.append(f"합계 {len(out):,}행 × {len(out.columns)}열")
    return out, notes


# ==============================================================================
# 6. 화면
# ==============================================================================
up = st.file_uploader("엑셀 또는 CSV 파일", type=["xlsx", "xls", "csv"])
if up is None:
    st.info("파일을 올리면 열 설정 화면이 나타납니다.")
    st.stop()

raw = up.getvalue()

try:
    sheets = list_sheets(raw, up.name)
except Exception as e:                               # noqa: BLE001
    st.error(f"파일을 열지 못했습니다 — {e}")
    st.stop()

MODE_SIDE = "옆으로 붙이기 (같은 응답자, 다른 문항)"
MODE_STACK = "위아래로 잇기 (같은 문항, 다른 응답자)"

merge_mode, join_key, keep_all, add_sheet_col = None, None, True, False

if len(sheets) == 1:
    sel = sheets
    c2, = st.columns(1)
    header_row = st.number_input("머리글 행", 1, 50, 1, key="sav_hdr",
                                 help="열 이름이 들어 있는 행 번호")
else:
    st.info(f"시트가 {len(sheets)}개입니다. 합쳐서 하나의 .sav 로 만들 수 있습니다.")
    c1, c2 = st.columns([3, 1])
    with c1:
        sel = st.multiselect("사용할 시트", sheets, default=sheets, key="sav_sheets")
    with c2:
        header_row = st.number_input("머리글 행", 1, 50, 1, key="sav_hdr",
                                     help="열 이름이 들어 있는 행 번호")
    if not sel:
        st.warning("시트를 하나 이상 고르세요.")
        st.stop()
    if len(sel) > 1:
        merge_mode = st.radio("합치는 방식", [MODE_SIDE, MODE_STACK],
                              key="sav_mode", horizontal=True)

# ── 시트 읽기 ─────────────────────────────────────────────────────────────
frames = {}
try:
    for s in sel:
        d = read_table(raw, up.name, s, int(header_row))
        d = d.loc[:, [c for c in d.columns if not str(c).startswith("Unnamed:")]]
        frames[s] = d
except Exception as e:                               # noqa: BLE001
    st.error(f"파일을 읽지 못했습니다 — {e}")
    st.stop()

merge_notes = []
if len(sel) == 1:
    df = frames[sel[0]]
elif merge_mode == MODE_SIDE:
    common = [c for c in frames[sel[0]].columns
              if all(c in d.columns for d in frames.values())]
    if not common:
        st.error("시트끼리 공통으로 가진 열이 없어 옆으로 붙일 수 없습니다.")
        st.stop()
    k1, k2 = st.columns([2, 2])
    with k1:
        default_key = next((c for c in common
                            if str(c).lower() in ("id", "panel_id", "respondent_id")),
                           common[0])
        join_key = st.selectbox("연결 열", common,
                                index=common.index(default_key), key="sav_key",
                                help="시트끼리 같은 응답자를 알아보는 기준 열")
    with k2:
        keep_all = st.radio(
            "한쪽에만 있는 응답자",
            ["모두 남기기", "양쪽에 다 있는 사람만"],
            key="sav_keep", horizontal=True) == "모두 남기기"
    try:
        df, merge_notes = merge_side(frames, join_key, keep_all)
    except Exception as e:                           # noqa: BLE001
        st.error(f"시트를 붙이지 못했습니다 — {e}")
        st.stop()
else:
    add_sheet_col = st.checkbox("어느 시트에서 왔는지 열로 남기기", value=True,
                                key="sav_srccol")
    df, merge_notes = merge_stack(frames, add_sheet_col)

if df.empty or not len(df.columns):
    st.error("읽어들인 표가 비어 있습니다. 머리글 행 번호를 확인해 주세요.")
    st.stop()

st.success(f"{len(df):,}행 × {len(df.columns)}열을 읽었습니다.")
if merge_notes:
    with st.expander("시트를 합친 과정", expanded=False):
        for n in merge_notes:
            st.write("· " + n)
with st.expander("원본 미리보기 (앞 20행)", expanded=False):
    st.dataframe(df.head(20).astype(str), **_wide(st.dataframe))

# ── 열 설정 표 ────────────────────────────────────────────────────────────
sig = hashlib.md5(
    f"{up.name}|{len(raw)}|{'/'.join(sel)}|{header_row}|{merge_mode}|"
    f"{join_key}|{keep_all}|{add_sheet_col}".encode()
).hexdigest()

if st.session_state.get("sav_sig") != sig:
    used, rows, empties = set(), [], []
    for i, c in enumerate(df.columns, start=1):
        sr = df[c]
        kind = auto_kind(sr, suggest_code=True)
        blank = bool(sr.dropna().empty) or all(
            v is None for v in text_list(sr))
        if blank:
            empties.append(str(c))
        rows.append({
            "포함": True,
            "원본열": str(c),
            "변수명": to_spss_name(c, i, used),
            "변수라벨": str(c),
            "유형": kind,
            "측정": auto_measure(sr, kind),
            "결측코드": "",
        })
    st.session_state["sav_spec"] = pd.DataFrame(rows)
    st.session_state["sav_sig"] = sig
    st.session_state["sav_empty"] = empties
    st.session_state["sav_editor_n"] = st.session_state.get("sav_editor_n", 0) + 1
    st.session_state.pop("sav_bytes", None)

st.subheader("열 설정")
_empty = st.session_state.get("sav_empty") or []
if _empty:
    e1, e2 = st.columns([3, 1])
    with e1:
        st.caption(
            f"값이 하나도 없는 열이 {len(_empty)}개 있습니다 — "
            + ", ".join(_empty[:12]) + (" …" if len(_empty) > 12 else "")
            + "  (문항 구조를 맞추기 위해 기본은 포함입니다)"
        )
    with e2:
        if st.button("빈 열 빼기", **_wide(st.button)):
            cur = st.session_state["sav_spec"].copy()
            cur.loc[cur["원본열"].isin(_empty), "포함"] = False
            st.session_state["sav_spec"] = cur
            st.session_state["sav_editor_n"] = \
                st.session_state.get("sav_editor_n", 0) + 1
            st.session_state.pop("sav_bytes", None)
            st.rerun()
st.caption(
    "**변수명** 은 SPSS에서 쓸 이름(영문·숫자·밑줄만)이고, **변수라벨** 은 원래 "
    "한글 이름입니다. **문자→코드화** 를 고르면 응답 텍스트가 1, 2, 3… 숫자로 "
    "바뀌고 값 라벨이 함께 저장됩니다."
)

spec = st.data_editor(
    st.session_state["sav_spec"],
    key=f"sav_editor_{st.session_state.get('sav_editor_n', 0)}",
    **_wide(st.data_editor),
    hide_index=True,
    num_rows="fixed",
    column_config={
        "포함": st.column_config.CheckboxColumn("포함", width="small"),
        "원본열": st.column_config.TextColumn("원본 열", disabled=True),
        "변수명": st.column_config.TextColumn("SPSS 변수명", required=True),
        "변수라벨": st.column_config.TextColumn("변수 라벨"),
        "유형": st.column_config.SelectboxColumn(
            "유형", options=[TYPE_AUTO, TYPE_NUM, TYPE_STR, TYPE_CODE, TYPE_DATE],
            required=True),
        "측정": st.column_config.SelectboxColumn(
            "측정", options=list(MEASURE_MAP), required=True, width="small"),
        "결측코드": st.column_config.TextColumn(
            "결측 코드", width="small",
            help="이 열에서 결측으로 볼 값. 쉼표로 여러 개. 예: 99, 999"),
    },
)

b1, b2 = st.columns([2, 1])
with b1:
    bulk = st.text_input(
        "결측 코드 일괄 입력", value="", placeholder="예: 99, 999",
        help="아래 버튼을 누르면 숫자 유형인 열의 ‘결측 코드’ 칸을 이 값으로 채웁니다. "
             "열마다 다르게 두고 싶으면 표에서 직접 고치세요.",
    )
with b2:
    st.write("")
    st.write("")
    if st.button("숫자 열에 채우기", **_wide(st.button)):
        cur = spec.copy()
        cur.loc[cur["유형"].isin([TYPE_NUM, TYPE_AUTO]), "결측코드"] = bulk.strip()
        st.session_state["sav_spec"] = cur
        st.session_state["sav_editor_n"] = st.session_state.get("sav_editor_n", 0) + 1
        st.session_state.pop("sav_bytes", None)
        st.rerun()

o1, o2, o3 = st.columns([2, 2, 1])
with o1:
    save_mode = st.radio("저장할 변수", [SAVE_ALL, SAVE_NUM, SAVE_STR],
                         key="sav_mode_save", horizontal=True,
                         help="‘숫자만’ 은 글자로 된 변수를 빼고, ‘문자만’ 은 숫자·날짜를 "
                              "빼고 코드화 대상도 원래 글자 그대로 저장합니다.")
with o2:
    file_label = st.text_input("파일 설명(선택)", value="",
                               placeholder="예: 2026년 상반기 본조사")
with o3:
    st.write("")
    st.write("")
    run = st.button("변환 실행", type="primary", **_wide(st.button))

always_cols = []
if save_mode != SAVE_ALL:
    cand = [str(c) for c in df.columns]
    guess = [c for c in cand
             if str(c).lower() in ("id", "panel_id", "respondent_id", "no")
             or c == join_key]
    always_cols = st.multiselect(
        "저장 범위와 상관없이 남길 열", cand,
        default=[c for c in dict.fromkeys(guess)][:3], key="sav_always",
        help="ID 처럼 나중에 다시 붙일 때 필요한 열은 여기에 두면 항상 들어갑니다.",
    )

# ── 이름 중복 검사 (변환 전에 잡는다) ─────────────────────────────────────
active = spec[spec["포함"]]
dups = active["변수명"].str.lower().duplicated(keep=False)
if dups.any():
    st.error("변수명이 겹칩니다 — " + ", ".join(sorted(set(active.loc[dups, "변수명"]))))
    run = False

bad_name = active[~active["변수명"].str.match(r"^[A-Za-z][0-9A-Za-z_]{0,63}$", na=False)]
if len(bad_name):
    st.error(
        "SPSS 변수명 규칙에 어긋납니다(첫 글자는 영문, 영문·숫자·밑줄만, 64자 이내) — "
        + ", ".join(bad_name["변수명"].astype(str))
    )
    run = False

# ── 실행 ────────────────────────────────────────────────────────────────
if run:
    st.session_state["sav_spec"] = spec
    try:
        with st.spinner("SPSS 파일을 만드는 중…"):
            data, vlabels, warns, skipped = build_sav(
                df, spec, file_label, save_mode, tuple(always_cols))
        st.session_state["sav_bytes"] = data
        st.session_state["sav_vlabels"] = vlabels
        st.session_state["sav_warns"] = warns
        st.session_state["sav_skipped"] = skipped
        st.session_state["sav_mode_used"] = save_mode
    except Exception as e:                           # noqa: BLE001
        st.session_state.pop("sav_bytes", None)
        st.error(f"변환에 실패했습니다 — {e}")

if st.session_state.get("sav_bytes"):
    data = st.session_state["sav_bytes"]
    for w in st.session_state.get("sav_warns", []):
        st.warning(w)

    _skip = st.session_state.get("sav_skipped") or []
    _mode = st.session_state.get("sav_mode_used", SAVE_ALL)
    if _skip:
        st.caption(
            f"‘{_mode}’ 라서 {len(_skip)}개 열을 뺐습니다 — "
            + ", ".join(_skip[:15]) + (" …" if len(_skip) > 15 else "")
        )

    suffix = {SAVE_NUM: "_숫자", SAVE_STR: "_문자"}.get(_mode, "")
    st.download_button(
        "💾 .sav 내려받기",
        data=data,
        file_name=os.path.splitext(up.name)[0] + suffix + ".sav",
        mime="application/octet-stream",
        type="primary",
    )
    st.caption(f"파일 크기 {len(data)/1024:,.0f} KB")

    vlabels = st.session_state.get("sav_vlabels", {})
    if vlabels:
        with st.expander(f"값 라벨 확인 ({len(vlabels)}개 변수)", expanded=False):
            recs = [{"변수": v, "코드": int(k), "라벨": lab}
                    for v, m in vlabels.items() for k, lab in sorted(m.items())]
            st.dataframe(pd.DataFrame(recs), hide_index=True, **_wide(st.dataframe))

    with st.expander("저장된 내용 되읽어 확인", expanded=False):
        tmp = tempfile.NamedTemporaryFile(suffix=".sav", delete=False)
        tmp.write(data)
        tmp.close()
        try:
            back, meta = pyreadstat.read_sav(tmp.name, user_missing=True)
            st.write(f"{len(back):,}행 × {len(back.columns)}열")
            mr = meta.missing_ranges or {}
            st.dataframe(
                pd.DataFrame({
                    "변수": list(back.columns),
                    "라벨": [meta.column_names_to_labels.get(c, "") for c in back.columns],
                    "형식": [meta.original_variable_types.get(c, "") for c in back.columns],
                    "측정": [meta.variable_measure.get(c, "") for c in back.columns],
                    "결측": [", ".join(str(int(r["lo"])) for r in mr.get(c, []))
                             for c in back.columns],
                }),
                hide_index=True, **_wide(st.dataframe)
            )
            st.dataframe(back.head(10), **_wide(st.dataframe))
        finally:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass

st.divider()
st.caption(f"SAV 변환 v{PAGE_VERSION}")

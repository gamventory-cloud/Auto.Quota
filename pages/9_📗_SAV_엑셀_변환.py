# -*- coding: utf-8 -*-
"""
SAV → 엑셀 변환 (v2.7)

SPSS .sav 파일을 업로드하면 여러 시트로 구성된 엑셀 파일을 내려받습니다.
  · Raw        : 숫자 코드 그대로
  · Label      : 값 레이블로 치환 (레이블 없는 변수는 원래 값 유지)
  · Open       : 키 변수(NO, id) + 문자형(주관식) 변수
                 문자형 변수가 없어도 키 변수만으로 만든다
  · Code       : DP 코드북 형식 (변수마다 문항 + 코드값/보기 블록)
  · 변수 가이드 : 변수명 + 변수 설명

'엑셀 값 반영하기' 탭에서는 반대로, 코딩·수정을 마친 엑셀을 올려
ID 로 짝을 맞춰 SAV 값을 덮어씁니다. 결과는 SAV 로도, 위 시트 구성의
엑셀로도 바로 받을 수 있습니다. (SAV 를 다시 올릴 필요 없음)

이 페이지는 utils.py 없이도 단독으로 동작합니다.
"""

import io
import os
import random
import re
import tempfile
import zipfile

import numpy as np
import pandas as pd
import pyreadstat
import streamlit as st
from openpyxl.styles import Font, PatternFill

# ──────────────────────────────────────────────────────────────
# 비밀번호 (utils.py 있으면 사용, 없으면 통과)
# ──────────────────────────────────────────────────────────────
try:
    from utils import check_password  # type: ignore
except Exception:  # pragma: no cover
    def check_password() -> bool:
        return True


st.set_page_config(page_title="SAV → 엑셀 변환", page_icon="📗", layout="wide")

if not check_password():
    st.stop()

st.title("📗 SAV → 엑셀 변환")
st.caption("SPSS .sav 파일을 Raw / Label / Open / Code / 변수 가이드 시트의 엑셀로 바꿔 드립니다.")


# ──────────────────────────────────────────────────────────────
# 변환 로직
# ──────────────────────────────────────────────────────────────
def _clean(v):
    """NaN은 빈칸으로, 소수점 없는 실수는 정수로."""
    if v is None:
        return None
    if isinstance(v, float):
        if np.isnan(v):
            return None
        if float(v).is_integer():
            return int(v)
    return v


@st.cache_data(show_spinner=False, max_entries=5)
def read_sav_bytes(data: bytes, filename: str):
    """업로드된 바이트를 임시파일로 떨어뜨려 pyreadstat으로 읽는다.

    .sav 에 문자셋 표시가 없는 경우가 있다 (pyreadstat 으로 쓴 파일이 그렇다).
    그러면 읽는 쪽이 짐작하는데, 빗나가면 한글이 깨지거나
    ReadstatError 로 아예 못 읽는다. 기본 → UTF-8 → CP949 순으로 시도한다.

    반환 끝에 쓴 인코딩을 붙인다 ("" = 기본값으로 읽음).
    """
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".sav", delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name

        last = None
        for enc in (None, "UTF-8", "CP949"):
            kw = {} if enc is None else {"encoding": enc}
            try:
                df, meta = pyreadstat.read_sav(
                    tmp_path, apply_value_formats=False, **kw)
            except Exception as exc:              # noqa: BLE001, PERF203
                last = exc
                continue
            return (
                df,
                dict(meta.column_names_to_labels or {}),
                dict(meta.variable_value_labels or {}),
                dict(getattr(meta, "readstat_variable_types", {}) or {}),
                {"read": enc or "",
                 # 파일이 스스로 밝힌 인코딩. 라벨 한도(120바이트)는 이 기준이라
                 # 잘림 판정에 쓴다. (EUC-KR 파일을 UTF-8 로 재면 못 잡는다)
                 "file": str(getattr(meta, "file_encoding", "") or "")},
            )
        raise last
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


# SPSS 한도 — 값 라벨 120바이트, 변수 라벨 256바이트.
# 한도는 **파일에 저장된 인코딩** 기준이다 (EUC-KR 한글 2B, UTF-8 3B).
VALLABEL_LIMIT = 120
VARLABEL_LIMIT = 256

# 한도에서 자르면 마지막 온전한 글자 경계는 한도-2 ~ 한도 사이에 떨어진다
# (2바이트 글자면 119·120, 3바이트 글자면 118·119·120).
# 이보다 넓게 잡으면 우연히 그 길이인 멀쩡한 라벨까지 잘렸다고 잡는다.
CEIL_SLACK = 2


_NUM_PREFIX = re.compile(r"^\s*\d+\s*[)\.]\s*")


def _label_core(text) -> str:
    """앞에 붙은 번호('  2) ')를 뗀 본문. 길이를 견줄 때 쓴다."""
    return _NUM_PREFIX.sub("", str(text).strip()).rstrip()


def _enc_len(text, encoding: str) -> int:
    """그 인코딩으로 저장했을 때의 바이트 수."""
    try:
        return len(str(text).encode(encoding, errors="replace"))
    except LookupError:
        return len(str(text).encode("utf-8", errors="replace"))


def _find_ceiling(labels, limit: int, encodings) -> str | None:
    """라벨 길이가 한도에 '막혀' 있으면 그 기준이 된 인코딩을 돌려준다.

    한도는 파일에 저장된 인코딩 기준이다. 같은 라벨이라도 EUC-KR 로는
    120바이트인데 UTF-8 로 재면 168바이트라, 한 가지 기준으로만 보면
    놓친다. 그래서 후보 인코딩으로 각각 재어 보고
      · 한도를 넘는 라벨이 하나도 없고
      · 한도에 딱 붙은 라벨이 여럿 있는
    인코딩을 찾는다. 자연스럽게 생긴 길이 분포는 이렇게 되지 않는다.
    """
    best = None
    for enc in encodings:
        lens = [_enc_len(s, enc) for s in labels if str(s).strip()]
        if not lens:
            continue
        top = max(lens)
        if top > limit or top < limit - CEIL_SLACK:
            continue                      # 천장이 없거나 한도와 무관하다
        at_ceiling = sum(1 for n in lens if n >= limit - CEIL_SLACK)
        if at_ceiling >= 3 and (best is None or at_ceiling > best[1]):
            best = (enc, at_ceiling)
    return best[0] if best else None


def _looks_truncated(text, limit: int, encoding: str | None = None) -> bool:
    """잘려 나간 라벨로 보이는지.

    ① 천장이 확인된 인코딩에서 한도에 닿아 있으면 잘린 것이다.
       (1차 조사 파일처럼 '?' 없이 문장만 뚝 끊긴 경우가 여기 해당한다)
    ② 천장을 못 찾았어도, 한도 근처에서 '?' 나 '�' 로 끝나면 잘린 것이다.
       한글 한 글자가 경계에 걸려 반토막 난 흔적이다.
       '만족하십니까?' 같은 멀쩡한 라벨은 길이가 한참 짧아 걸리지 않는다.
    """
    s = str(text)
    # 문장이 끝맺어져 있으면 길이가 한도에 닿았어도 잘린 것이 아니다.
    # (우연히 118~120바이트인 멀쩡한 라벨이 실제로 있다 — '…지급합니다.')
    # 잘린 라벨은 '…국민연금 가' 처럼 단어 중간에서 끊긴다.
    if s.rstrip().endswith((".", "!", "。", "…")):
        return False
    if encoding and _enc_len(s, encoding) >= limit - CEIL_SLACK:
        return True
    return (s.rstrip().endswith(("?", "�"))
            and len(s.encode("utf-8", errors="replace")) >= limit - CEIL_SLACK)


def label_encodings(file_encoding: str = "") -> list:
    """천장을 찾아볼 인코딩 후보. 파일이 말하는 것을 먼저 본다."""
    out = []
    for e in (file_encoding, "utf-8", "euc-kr"):
        e = (e or "").strip()
        if e and e.lower() not in {x.lower() for x in out}:
            out.append(e)
    return out


def find_broken_labels(col_labels: dict, value_labels: dict,
                       file_encoding: str = "") -> list:
    """잘려 나간 것으로 보이는 라벨 목록. [(변수, 위치, 라벨), ...]"""
    encs = label_encodings(file_encoding)
    v_labs = [l for mp in (value_labels or {}).values()
              for l in (mp or {}).values()]
    c_labs = [l for l in (col_labels or {}).values() if l]
    v_enc = _find_ceiling(v_labs, VALLABEL_LIMIT, encs)
    c_enc = _find_ceiling(c_labs, VARLABEL_LIMIT, encs)

    bad = []
    for var, lab in (col_labels or {}).items():
        if lab and _looks_truncated(lab, VARLABEL_LIMIT, c_enc):
            bad.append((str(var), "변수 라벨", str(lab)))
    for var, mapping in (value_labels or {}).items():
        for code, lab in (mapping or {}).items():
            if _looks_truncated(lab, VALLABEL_LIMIT, v_enc):
                bad.append((str(var), f"코드 {_code_str(code)}", str(lab)))
    return bad


# ──────────────────────────────────────────────────────────────
# 코드북에서 잘리지 않은 원문 가져오기
#
#   .sav 의 값 라벨은 120바이트(한글 40자)에서 잘려 있다. 잘려 나간 글자는
#   파일에 남아 있지 않아 .sav 만으로는 되살릴 수 없다.
#   원문은 코드북에 있으므로, 그것을 받아 Code/Label 시트를 채운다.
# ──────────────────────────────────────────────────────────────
CB_HEADER_HINTS = ("변수명", "변수 명")
GUIDE_VNAME_HINTS = ("v변수명", "V변수명", "원변수명", "원본변수명")
GUIDE_DESC_HINTS = ("변수 설명", "변수설명", "변수라벨", "변수 내용", "문항")


def _code_blocks(ws) -> list:
    """'코드값 | 보기' 머리줄로 시작하는 블록을 **순서대로** 읽는다.

    변수명이 적혀 있지 않은 Code 시트가 있다. 그런 코드북은 블록 순서가
    변수가이드의 행 순서와 같다는 것이 유일한 연결 고리라, 코드가 하나도
    없는 블록(주관식 문항 등)도 자리를 지키도록 빈 dict 로 남겨 둔다.

    열 위치는 파일마다 다르므로(A/B 인 것도, B/C 인 것도 있다) 머리줄에서
    '코드값' 과 '보기' 가 나란히 있는 자리를 찾아 쓴다.
    """
    blocks, cur, c_code, c_lab = [], None, None, None
    for row in ws.iter_rows(values_only=True):
        cells = ["" if c is None else str(c).strip() for c in row]
        head = next((i for i in range(len(cells) - 1)
                     if cells[i] == "코드값" and cells[i + 1] == "보기"), None)
        if head is not None:
            c_code, c_lab = head, head + 1
            cur = {}
            blocks.append(cur)
            continue
        if cur is None or c_code is None:
            continue
        code = cells[c_code] if c_code < len(cells) else ""
        lab = cells[c_lab] if c_lab < len(cells) else ""
        if not code:
            cur = None                      # 빈 줄 = 블록 끝
            continue
        if lab:
            num = _as_number(code)
            cur[_code_str(num if num is not None else code)] = lab
    return blocks


def _read_guide_pair(wb):
    """'변수가이드' + '「코드값/보기」만 있는 Code' 두 시트로 된 코드북.

    Code 시트에 변수명이 없으므로 **블록 순서 ↔ 변수가이드 행 순서** 로만
    짝을 지을 수 있다. 개수가 다르면 어느 한쪽이 밀린 것이므로, 엉뚱한
    변수에 라벨을 붙이는 대신 읽기를 포기한다.
    """
    guide = None
    for ws in wb.worksheets:
        try:
            header = [str(h or "").strip()
                      for h in next(ws.iter_rows(max_row=1, values_only=True))]
        except StopIteration:
            continue
        if any(h in CB_HEADER_HINTS for h in header) and \
                any(h in GUIDE_DESC_HINTS for h in header):
            guide = (ws, header)
            break
    if guide is None:
        return None

    ws_g, header = guide
    i_name = next(i for i, h in enumerate(header) if h in CB_HEADER_HINTS)
    i_desc = next(i for i, h in enumerate(header) if h in GUIDE_DESC_HINTS)
    i_vname = next((i for i, h in enumerate(header) if h in GUIDE_VNAME_HINTS),
                   None)

    rows = []
    for row in ws_g.iter_rows(min_row=2, values_only=True):
        name = "" if i_name >= len(row) or row[i_name] is None \
            else str(row[i_name]).strip()
        if not name:
            continue
        desc = "" if i_desc >= len(row) or row[i_desc] is None \
            else str(row[i_desc]).strip()
        vname = ""
        if i_vname is not None and i_vname < len(row) and row[i_vname]:
            vname = str(row[i_vname]).strip()
        rows.append((name, vname, desc))
    if not rows:
        return None

    for ws in wb.worksheets:
        if ws is ws_g:
            continue
        blocks = _code_blocks(ws)
        if not blocks:
            continue
        if len(blocks) != len(rows):
            raise ValueError(
                f"‘{ws.title}’ 시트의 보기 블록이 {len(blocks)}개인데 "
                f"‘{ws_g.title}’ 의 변수는 {len(rows)}개입니다. "
                "이 코드북은 Code 시트에 변수명이 없어 **순서로만** 짝을 지을 수 "
                "있는데, 개수가 다르면 한 칸씩 밀려 엉뚱한 변수에 라벨이 붙습니다. "
                "두 시트의 행을 맞춘 뒤 다시 올려 주세요."
            )
        cl, vl = {}, {}
        for (name, vname, desc), block in zip(rows, blocks):
            for key in (name, vname):
                if not key:
                    continue
                if desc:
                    cl.setdefault(key, desc)
                if block:
                    vl.setdefault(key, dict(block))
        if vl:
            return cl, vl, f"코드북 ({ws_g.title} + {ws.title} 시트, 순서로 매칭)"
    return None


def _cb_split_values(text) -> dict:
    """'1=남성 | 2=여성' -> {'1': '남성', '2': '여성'}"""
    out = {}
    for chunk in str(text or "").split("|"):
        if "=" not in chunk:
            continue
        code, _, lab = chunk.partition("=")
        code, lab = code.strip(), lab.strip()
        if code and lab:
            out[_code_str(_as_number(code) if _as_number(code) is not None
                          else code)] = lab
    return out


def read_codebook_xlsx(data: bytes):
    """코드북 엑셀에서 라벨을 뽑는다. 세 가지 형식을 받는다.

    ① SPSS 라벨링 페이지의 코드북
         변수명 · 변수라벨 · 값라벨('1=남성 | 2=여성') 열을 가진 표
    ② 변수가이드 + Code 두 시트로 된 코드북 (DP 납품본에서 흔하다)
         Code 시트에 변수명이 없고 '코드값 | 보기' 블록만 이어진다.
         블록 순서가 변수가이드 행 순서와 같다는 것으로 짝을 짓는다.
    ③ 이 페이지가 내보낸 Code 시트
         변수 / 내용 두 열에 블록이 이어지는 형태. 받은 엑셀에서 잘린 칸만
         고쳐 다시 올리는 방법이라, 코드북이 없어도 쓸 수 있다.

    반환: (변수라벨 dict, 값라벨 dict{변수: {코드문자열: 라벨}}, 형식이름)
    """
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        # ① 코드북 형식 — '변수명' 머리글이 있는 시트를 찾는다
        for ws in wb.worksheets:
            rows = ws.iter_rows(values_only=True)
            try:
                header = [str(h or "").strip() for h in next(rows)]
            except StopIteration:
                continue
            if not any(h in CB_HEADER_HINTS for h in header):
                continue
            i_name = next(i for i, h in enumerate(header) if h in CB_HEADER_HINTS)
            i_lab = header.index("변수라벨") if "변수라벨" in header else None
            i_val = header.index("값라벨") if "값라벨" in header else None
            cl, vl = {}, {}
            for row in rows:
                if i_name >= len(row) or row[i_name] is None:
                    continue
                name = str(row[i_name]).strip()
                if not name:
                    continue
                if i_lab is not None and i_lab < len(row) and row[i_lab]:
                    cl[name] = str(row[i_lab]).strip()
                if i_val is not None and i_val < len(row) and row[i_val]:
                    got = _cb_split_values(row[i_val])
                    if got:
                        vl[name] = got
            if cl or vl:
                return cl, vl, f"코드북 ({ws.title} 시트)"

        # ② 변수가이드 + Code 두 시트 (Code 에 변수명이 없는 코드북)
        got = _read_guide_pair(wb)
        if got:
            return got

        # ③ Code 시트 형식 — '코드값' 줄 바로 위가 변수 머리줄이다
        for ws in wb.worksheets:
            cl, vl = {}, {}
            prev = None
            cur = None
            for row in ws.iter_rows(min_col=1, max_col=2, values_only=True):
                a = "" if row[0] is None else str(row[0]).strip()
                b = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
                if a == "코드값":
                    if prev and prev[0]:
                        cur = prev[0]
                        cl.setdefault(cur, prev[1])
                        vl.setdefault(cur, {})
                    continue
                if not a:
                    cur = None
                elif cur is not None:
                    if b:
                        vl[cur][_code_str(_as_number(a) if _as_number(a)
                                          is not None else a)] = b
                prev = (a, b)
            vl = {k: v for k, v in vl.items() if v}
            if vl:
                return cl, vl, f"Code 시트 ({ws.title})"
    finally:
        wb.close()

    raise ValueError(
        "라벨을 찾지 못했습니다. SPSS 라벨링 페이지의 코드북(변수명·값라벨 열이 "
        "있는 표)이나, 이 페이지가 내보낸 엑셀의 Code 시트를 올려 주세요."
    )


def apply_codebook(col_labels: dict, value_labels: dict,
                   cb_col: dict, cb_val: dict, only_broken: bool,
                   file_encoding: str = ""):
    """코드북 라벨로 .sav 라벨을 채운다. (새 변수라벨, 새 값라벨, 리포트)

    file_encoding 은 '이미 잘려 있는 라벨' 을 가려내는 데 쓴다.
    한도는 파일에 저장된 인코딩 기준이라 이것 없이는 판정이 어긋난다.
    """
    # 복구 전에 어느 라벨이 잘려 있었는지 먼저 확정한다. 복구 뒤에 다시
    # 찾으면 천장이 사라져(긴 원문이 들어가서) 못 채운 것까지 멀쩡해 보인다.
    orig_broken = {(v, w) for v, w, _ in
                   find_broken_labels(col_labels, value_labels, file_encoding)}
    encs = label_encodings(file_encoding)
    _v_enc = _find_ceiling(
        [l for mp in (value_labels or {}).values() for l in (mp or {}).values()],
        VALLABEL_LIMIT, encs)
    _c_enc = _find_ceiling(
        [l for l in (col_labels or {}).values() if l], VARLABEL_LIMIT, encs)
    new_col = dict(col_labels or {})
    new_val = {k: dict(v) for k, v in (value_labels or {}).items()}
    filled, skipped, unknown = [], [], []

    # 순서로 짝을 지은 코드북은 한 칸만 밀려도 엉뚱한 변수에 라벨이 붙는다.
    # 양쪽에 다 있는 변수의 '코드 집합' 이 같은지 세어 두고, 어긋나는 것이
    # 많으면 화면에서 경고한다.
    matched_vars, mismatched = 0, []
    for var, mapping in (cb_val or {}).items():
        if var not in new_val:
            continue
        sav_codes = {_code_str(c) for c in new_val[var]}
        cb_codes = set(mapping)
        if cb_codes and sav_codes and cb_codes <= sav_codes:
            matched_vars += 1
        else:
            mismatched.append(var)

    def _no_gain(new_text, old_text) -> bool:
        """이 라벨로 바꿔봐야 얻는 것이 없는 경우.

        내보낸 Code 시트를 고치지 않고 그대로 다시 올리는 일이 흔하다.
        그 안의 라벨은 잘린 채로 돌아오는 데다 번호 접두사('  2) ')까지
        떨어져 나가 있어서, 덮어쓰면 오히려 글자가 줄고 잘림 경고에서도
        빠져버린다. 잘림을 메우는 복구는 **반드시 길어진다**는 점을 쓴다.

        ① 여전히 '?'·'�' 로 끝나면 그쪽도 잘린 것이라 쓰지 않는다
        ② 지금 것보다 길어지지 않으면 복구가 아니다
        """
        new_s, old_s = str(new_text).rstrip(), str(old_text).rstrip()
        if new_s.endswith(("?", "�")):
            return True
        # 길이는 번호 접두사를 뗀 본문끼리 견준다. .sav 라벨에는 '  2) ' 가
        # 붙어 있고 코드북에는 없는 경우가 많아, 그대로 재면 제대로 된 원문도
        # '안 길어졌다' 고 걸러진다.
        return len(_label_core(new_s)) <= len(_label_core(old_s))

    for var, lab in (cb_col or {}).items():
        if var not in new_col or _no_gain(lab, new_col[var]):
            continue
        if only_broken and not _looks_truncated(new_col[var], VARLABEL_LIMIT,
                                                _c_enc):
            continue
        if str(new_col[var]).strip() != str(lab).strip():
            new_col[var] = lab
            filled.append((var, "변수 라벨"))

    for var, mapping in (cb_val or {}).items():
        if var not in new_val:
            unknown.append(var)
            continue
        by_code = {_code_str(c): c for c in new_val[var]}
        for code_s, lab in mapping.items():
            key = by_code.get(code_s)
            if key is None or _no_gain(lab, new_val[var][key]):
                continue
            if only_broken and not _looks_truncated(new_val[var][key],
                                                    VALLABEL_LIMIT, _v_enc):
                continue
            if str(new_val[var][key]).strip() == str(lab).strip():
                continue
            new_val[var][key] = lab
            filled.append((var, f"코드 {code_s}"))

    done = set(filled)
    skipped = sorted(orig_broken - done)

    return new_col, new_val, {"filled": filled, "still_broken": skipped,
                              "unknown_vars": sorted(set(unknown)),
                              "code_ok": matched_vars,
                              "code_mismatch": sorted(set(mismatched))}


def build_raw(df: pd.DataFrame) -> pd.DataFrame:
    return df.map(_clean)


def build_label(df: pd.DataFrame, value_labels: dict) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if c in value_labels:
            m = {k: str(v).strip() for k, v in value_labels[c].items()}
            out[c] = out[c].map(lambda x, m=m: m.get(x, _clean(x)))
        else:
            out[c] = out[c].map(_clean)
    return out


def find_key_cols(df: pd.DataFrame) -> list:
    """응답자를 되짚을 키 열. NO / id 를 원래 순서대로 모으고, 없으면 첫 열."""
    keys = [c for c in df.columns if str(c).strip().lower() in ("no", "id")]
    return keys if keys else [df.columns[0]]


def find_text_cols(df: pd.DataFrame, var_types: dict) -> list:
    """SAV에서 문자형으로 선언된 변수 목록."""
    return [c for c in df.columns if str(var_types.get(c, "")).lower() == "string"]


def build_open(df: pd.DataFrame, text_cols: list, key_cols: list) -> pd.DataFrame:
    """키 변수 + 주관식 응답. 행 순서는 Raw/Label과 동일하게 유지.

    문자형 변수가 없으면 키 변수만 담긴 시트가 된다. 주관식 코딩을 할 때
    이 시트에 열을 직접 추가해 쓸 수 있도록 빈 채로라도 만들어 둔다.
    """
    cols = list(key_cols) + [c for c in text_cols if c not in key_cols]
    out = df[cols].copy()
    for c in out.columns:
        if c in text_cols:
            out[c] = out[c].map(
                lambda x: None
                if x is None or (isinstance(x, float) and np.isnan(x)) or str(x).strip() == ""
                else str(x).strip()
            )
        else:
            out[c] = out[c].map(_clean)
    return out


# 머리글에 칠할 색. 엑셀 기본 팔레트의 연한 색들로, 검은 글씨가 잘 보인다.
# 순서가 화면에 나오는 순서다. 기본값은 맨 앞 항목.
HEAD_COLORS = {
    "주황": "#FCE4D6",
    "파랑": "#D9E1F2",
    "초록": "#E2EFDA",
    "노랑": "#FFF2CC",
    "회색": "#EDEDED",
    "자주": "#E4DFEC",
}
DEFAULT_HEAD_COLOR = next(iter(HEAD_COLORS.values()))

# 머리글 띠를 몇 열까지 칠할지. 모든 시트에 같이 적용된다.
# 데이터가 있는 데까지만 칠하면 오른쪽이 끊겨 보이므로,
# 최소 이 수까지 칠하고 데이터가 더 많으면 그 끝에서 여유분만큼 더 칠한다.
#
# openpyxl 은 RowDimension 에 fill 을 걸어도 스타일 번호만 붙이고 색은
# 넣지 않으므로(fillId 가 0 으로 남는다) 셀을 하나씩 칠해야 한다.
HEAD_FILL_COLS = 40
HEAD_FILL_MARGIN = 8

# 엑셀 열 상한. 이걸 넘기면 엑셀이 파일을 못 여는데 openpyxl 은 막지 않는다.
EXCEL_MAX_COLS = 16384


def _fill(color: str):
    """'#FCE4D6' 또는 'FCE4D6' → PatternFill. 빈 값이면 None."""
    if not color:
        return None
    return PatternFill("solid", fgColor=str(color).lstrip("#").upper())


def _code_str(v) -> str:
    """값 라벨의 코드를 표시용 문자열로. 1.0 -> '1'"""
    if isinstance(v, float) and float(v).is_integer():
        return str(int(v))
    return str(v)


def _option_text(code_str: str, label) -> str:
    """보기 문구에서 앞에 붙은 코드값을 뗀다.

    SPSS 값 라벨은 '  1) 남성' 처럼 코드가 앞에 붙어 저장되는 경우가 많다.
    코드값은 A열에 따로 들어가므로 중복이라 뗀다.
    떼고 나면 아무것도 안 남는 경우(척도 중간값처럼 '  5)' 만 있는 경우)는
    원문을 그대로 둔다. 빈 칸으로 보이면 누락처럼 읽히기 때문이다.
    """
    s = str(label).strip()
    m = re.match(r"^" + re.escape(code_str) + r"\s*[)\.]\s*", s)
    if m and s[m.end():].strip():
        return s[m.end():].strip()
    return s


def build_codebook(df: pd.DataFrame, col_labels: dict,
                   value_labels: dict, key_cols: list) -> pd.DataFrame:
    """DP 코드북 형식. 변수마다 블록 하나.

        q1        SQ1. 귀하의 성별은 무엇입니까?
        코드값     보기
        1         남성
        2         여성
        (빈 줄)
        (빈 줄)

    값 라벨이 없는 변수도 머리글까지는 넣고 코드 부분만 비운다.
    키 변수(no, id)는 문항이 아니므로 제외한다.
    """
    rows = []
    for c in df.columns:
        if c in key_cols:
            continue
        rows.append([str(c), str(col_labels.get(c) or "").strip()])
        rows.append(["코드값", "보기"])
        for code, lab in sorted((value_labels.get(c) or {}).items()):
            cs = _code_str(code)
            rows.append([cs, _option_text(cs, lab)])
        rows.append([None, None])
        rows.append([None, None])
    return pd.DataFrame(rows, columns=["변수", "내용"])


def build_guide(df: pd.DataFrame, col_labels: dict) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "변수명": list(df.columns),
            "변수 내용": [str(col_labels.get(c) or "").strip() for c in df.columns],
        }
    )


@st.cache_data(show_spinner=False, max_entries=5)
def list_sheets(data: bytes, filename: str):
    """엑셀 시트 이름 목록. CSV 면 None."""
    if os.path.splitext(filename)[1].lower() in (".csv", ".txt"):
        return None
    return pd.ExcelFile(io.BytesIO(data)).sheet_names


@st.cache_data(show_spinner=False, max_entries=5)
def read_table_bytes(data: bytes, filename: str, sheet=0) -> pd.DataFrame:
    """엑셀/CSV 업로드를 DataFrame 으로. 값은 손대지 않고 그대로 읽는다.

    sheet: 엑셀 시트 이름. CSV 면 무시한다.
    """
    ext = os.path.splitext(filename)[1].lower()
    bio = io.BytesIO(data)
    if ext in (".csv", ".txt"):
        for enc in ("utf-8-sig", "cp949", "utf-8"):
            try:
                bio.seek(0)
                return pd.read_csv(bio, dtype=object, encoding=enc)
            except UnicodeDecodeError:
                continue
        bio.seek(0)
        return pd.read_csv(bio, dtype=object, encoding="latin1")
    return pd.read_excel(bio, sheet_name=sheet, dtype=object)


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    return str(v).strip() == ""


def _as_number(v):
    """숫자로 읽히면 float, 아니면 None."""
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _to_text(v) -> str:
    """문자형으로 바꿀 때의 표시. 1.0 -> '1'"""
    if _is_blank(v):
        return ""
    if isinstance(v, float) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


SKIP_LABEL = "(넘기기)"


def guess_mapping(df: pd.DataFrame, patch: pd.DataFrame,
                  sav_key: str, patch_key: str) -> pd.DataFrame:
    """엑셀 열마다 어느 SAV 변수에 넣을지 짐작한 표.

    이름이 같으면(대소문자 무시) 그 변수를, 없으면 '(넘기기)'.
    화면에서 사람이 고쳐 쓸 수 있게 값 예시와 채워진 칸 수를 함께 담는다.
    """
    sav_by_lower = {str(c).lower(): c for c in df.columns}
    rows = []
    for pc in patch.columns:
        if str(pc) == str(patch_key):
            continue
        vals = [v for v in patch[pc] if not _is_blank(v)]
        hit = sav_by_lower.get(str(pc).lower())
        if hit is None or hit == sav_key:
            hit = SKIP_LABEL
        rows.append({
            "엑셀 열": str(pc),
            "값 예시": " / ".join(_to_text(v) for v in vals[:3]),
            "채워진 칸": len(vals),
            "SAV 변수": hit,
        })
    return pd.DataFrame(rows)


def apply_patch(df: pd.DataFrame, value_labels: dict, var_types: dict,
                patch: pd.DataFrame, sav_key: str, patch_key: str,
                mapping: dict) -> tuple:
    """엑셀 값을 SAV 데이터에 덮어쓴다.

    · ID 로 행을 짝지은 뒤, mapping 에 적힌 대로 열을 덮어쓴다.
      mapping: {엑셀 열 이름: SAV 변수 이름}. '(넘기기)' 는 건너뛴다.
    · 엑셀의 빈칸은 건드리지 않는다. 기존 값이 그대로 남는다.
    · 숫자 변수에 문자 값이 하나라도 들어오면 그 변수 전체를 문자형으로
      바꾼다. SPSS 는 한 변수에 숫자와 문자를 섞을 수 없기 때문이다.
      이때 값 라벨은 숫자 코드에 붙는 것이라 쓸 수 없게 되므로 버린다.

    반환: (새 df, 새 value_labels, 리포트 dict)
    """
    out = df.copy()
    new_labels = {k: dict(v) for k, v in value_labels.items()}

    pairs, only_in_patch = [], []
    for pc in patch.columns:
        if str(pc) == str(patch_key):
            continue
        sc = mapping.get(str(pc), SKIP_LABEL)
        if sc in (SKIP_LABEL, None, "") or sc not in out.columns:
            only_in_patch.append(str(pc))
        else:
            pairs.append((sc, pc))

    # ── ID 짝 맞추기 ──
    # 같은 ID 가 두 번 이상 나오면 한쪽만 반영되고 나머지는 조용히 사라진다.
    # 덮어쓰기 작업이라 조용히 넘기면 안 되므로 양쪽 모두 세어서 알린다.
    key_map, dup_sav = {}, []
    for i, v in enumerate(out[sav_key]):
        kid = _to_text(v)
        if kid == "":
            continue
        if kid in key_map:
            dup_sav.append(kid)          # 먼저 나온 행을 쓴다
            continue
        key_map[kid] = i

    changes, to_text_cols, unmatched_ids = {}, [], []
    matched_ids, dup_patch = set(), []

    for _, prow in patch.iterrows():
        kid = _to_text(prow[patch_key])
        if kid == "":
            continue
        idx = key_map.get(kid)
        if idx is None:
            unmatched_ids.append(kid)
            continue
        if kid in matched_ids:
            dup_patch.append(kid)        # 뒤에 오는 행이 앞을 덮어쓴다
        matched_ids.add(kid)
        for sc, pc in pairs:
            v = prow[pc]
            if _is_blank(v):
                continue                      # 빈칸은 건드리지 않는다
            changes.setdefault(sc, []).append((idx, v))

    # ── 열마다 반영 ──
    for sc, items in changes.items():
        numeric_col = str(var_types.get(sc, "")).lower() != "string"
        has_text = any(_as_number(v) is None for _, v in items)

        if numeric_col and has_text:
            # 열 전체를 문자형으로 바꾼다
            out[sc] = out[sc].map(_to_text)
            to_text_cols.append(sc)
            new_labels.pop(sc, None)
            for idx, v in items:
                out.iat[idx, out.columns.get_loc(sc)] = _to_text(v)
        elif numeric_col:
            col = out.columns.get_loc(sc)
            for idx, v in items:
                out.iat[idx, col] = _as_number(v)
        else:
            out[sc] = out[sc].map(_to_text)
            col = out.columns.get_loc(sc)
            for idx, v in items:
                out.iat[idx, col] = _to_text(v)

    report = {
        "pairs": [sc for sc, _ in pairs],
        "only_in_patch": only_in_patch,
        "changed": {k: len(v) for k, v in changes.items()},
        "to_text": to_text_cols,
        "matched": len(matched_ids),
        "unmatched_ids": unmatched_ids,
        "dup_sav": sorted(set(dup_sav)),
        "dup_patch": sorted(set(dup_patch)),
    }
    return out, new_labels, report


def write_sav(df: pd.DataFrame, col_labels: dict, value_labels: dict) -> bytes:
    """DataFrame 을 .sav 바이트로."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "out.sav")
        pyreadstat.write_sav(
            df, path,
            column_labels=[col_labels.get(c) or "" for c in df.columns],
            variable_value_labels={k: v for k, v in value_labels.items()
                                   if k in df.columns and v} or None,
        )
        with open(path, "rb") as f:
            return f.read()


def to_excel(sheets: dict, head_color: str = DEFAULT_HEAD_COLOR) -> bytes:
    """{시트명: DataFrame} → 엑셀 바이트.

    head_color: 머리글에 칠할 색. 빈 문자열이면 칠하지 않는다.
                Code 시트는 변수명·문항 줄, 나머지는 첫 행에 적용된다.
    """
    head_fill = _fill(head_color)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            # Code 는 변수 블록이 이어지는 형태라 표 머리글이 없다.
            is_code = name == "Code"
            frame.to_excel(writer, sheet_name=name, index=False,
                           header=not is_code)
            ws = writer.sheets[name]
            wide = max(ws.max_column + HEAD_FILL_MARGIN, HEAD_FILL_COLS)
            wide = min(wide, EXCEL_MAX_COLS)

            if is_code:
                ws.column_dimensions["A"].width = 14
                ws.column_dimensions["B"].width = 90
                # 변수명·문항 줄은 색을 채우고, '코드값' 줄은 굵게만.
                # 블록이 이어지는 시트라 눈으로 경계를 찾을 수 있어야 한다.
                #
                # 행 전체를 칠하려면 RowDimension 에 서식을 건다.
                # customFormat="1" 로 저장되는데, 엑셀에서 행 머리글을 눌러
                # 색을 칠했을 때와 같은 방식이라 오른쪽 끝까지 칠해진다.
                # 값이 든 칸(A·B)에도 따로 지정한다. 일부 뷰어가
                # 행 서식을 무시하고 셀 서식만 보기 때문이다.
                for r in range(1, ws.max_row + 1):
                    if ws.cell(row=r, column=1).value != "코드값":
                        continue
                    for c in (1, 2):
                        ws.cell(row=r, column=c).font = Font(bold=True)
                    if r > 1:
                        for c in range(1, wide + 1):
                            head = ws.cell(row=r - 1, column=c)
                            if c <= 2:
                                head.font = Font(bold=True)
                            if head_fill:
                                head.fill = head_fill
                continue

            for c in range(1, wide + 1):
                cell = ws.cell(row=1, column=c)
                if c <= ws.max_column:      # 굵게는 이름이 있는 칸만
                    cell.font = Font(bold=True)
                if head_fill:
                    cell.fill = head_fill
            ws.freeze_panes = "A2"
            if name == "변수 가이드":
                ws.column_dimensions["A"].width = 18
                ws.column_dimensions["B"].width = 100
            # Raw / Label / Open 은 너비를 지정하지 않는다.
            # 엑셀 기본 너비로 두면 사용자가 전체 선택 후 한 번에 조절할 수 있다.
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────
# UI
#   흐름은 위에서 아래로 한 줄이다.
#     SAV 올리기 → (선택) 엑셀 값 반영 → 시트 고르기 → 내려받기
#
#   파일을 여러 개 올리면 '엑셀 값 반영' 은 쓸 수 없다.
#   반영은 파일마다 ID 열과 변수 짝을 따로 정해야 하는 작업이라
#   여러 파일에 한꺼번에 적용할 수가 없다. 대신 같은 시트·색 설정으로
#   전부 변환해서 zip 하나로 내려준다.
# ──────────────────────────────────────────────────────────────
ups = st.file_uploader("SAV 파일을 올려주세요", type=["sav"],
                       accept_multiple_files=True)

if not ups:
    st.info("SPSS .sav 파일을 올리면 다음 단계가 나타납니다. 여러 개도 됩니다.")
    st.stop()

multi = len(ups) > 1

# ── 올린 파일 읽기 ──
loaded = []          # [(파일, df, col_labels, value_labels, var_types), ...]
fallback_enc = []    # 기본값으로 못 읽어 다른 인코딩으로 넘어간 파일
broken_labels = []   # 원본에서 이미 잘려 있는 라벨
file_encs = []       # 파일이 밝힌 인코딩 (라벨 한도 판정에 쓴다)
for f in ups:
    try:
        _d, _cl, _vl, _vt, _info = read_sav_bytes(f.getvalue(), f.name)
    except Exception as e:
        st.error(f"‘{f.name}’ 을 읽지 못했습니다: {e}")
        st.stop()
    loaded.append((f, _d, _cl, _vl, _vt))
    file_encs.append(_info.get("file", ""))
    if _info.get("read"):
        fallback_enc.append((f.name, _info["read"]))
    for _var, _where, _lab in find_broken_labels(_cl, _vl, _info.get("file", "")):
        broken_labels.append((f.name, _var, _where, _lab))

if fallback_enc:
    st.info(
        "문자셋 표시가 없어 인코딩을 바꿔 읽었습니다 — "
        + ", ".join(f"{n} → {e}" for n, e in fallback_enc)
        + ". 한글이 깨져 보이면 알려 주세요."
    )

if broken_labels:
    st.warning(
        f"**원본 .sav 에서 이미 잘려 있는 라벨이 {len(broken_labels)}개 있습니다.** "
        "SPSS 는 값 라벨을 120바이트(한글 40자), 변수 라벨을 256바이트에서 자르는데, "
        "경계가 한글 글자 가운데에 걸리면 그 글자가 `?` 로 남습니다. "
        "**잘려 나간 글자는 파일에 남아 있지 않아 이 도구가 되살릴 수 없습니다.** "
        "아래 목록의 라벨은 설문지나 코드북을 보고 직접 채워 주세요."
    )
    with st.expander(f"잘린 라벨 {len(broken_labels)}개 보기"):
        st.dataframe(
            pd.DataFrame(broken_labels,
                         columns=["파일", "변수", "위치", "잘린 라벨"]),
            hide_index=True, use_container_width=True,
        )

# ── 코드북으로 원문 채우기 ────────────────────────────────────────────────
with st.expander("📑 코드북으로 잘린 라벨 채우기 (선택)",
                 expanded=bool(broken_labels)):
    st.caption(
        "`.sav` 의 값 라벨은 120바이트(한글 40자)에서 잘려 있어 원문을 되살릴 수 "
        "없습니다. 원문이 남아 있는 코드북을 올리면 Code·Label 시트를 채웁니다.\n\n"
        "올릴 수 있는 것 — **SPSS 라벨링 페이지가 만든 코드북**(변수명·값라벨 열이 "
        "있는 표), 또는 **이 페이지가 내보낸 엑셀의 Code 시트**(받은 파일에서 잘린 "
        "칸만 고쳐 다시 올리면 됩니다)."
    )
    cb_file = st.file_uploader("코드북 또는 Code 시트 (.xlsx)", type=["xlsx"],
                               key="SX_codebook")
    only_broken = st.checkbox(
        "잘린 라벨만 채우기", value=True, key="SX_cb_only_broken",
        help="끄면 코드북에 있는 라벨로 전부 덮어씁니다. 코드북이 이 데이터의 "
             "것이 맞는지 확실할 때만 끄세요.",
    )

    if cb_file is not None:
        try:
            cb_col, cb_val, cb_kind = read_codebook_xlsx(cb_file.getvalue())
        except Exception as exc:                          # noqa: BLE001
            st.error(f"코드북을 읽지 못했습니다 — {exc}")
        else:
            st.success(
                f"{cb_kind} 에서 읽었습니다 — 변수 {len(cb_val)}개 · "
                f"값 라벨 {sum(len(v) for v in cb_val.values())}개"
            )
            # 먼저 계산만 해 보고, 짝이 어긋나 보이면 적용하지 않는다.
            # 순서로 짝지은 코드북은 한 칸만 밀려도 엉뚱한 변수에 라벨이 붙는데,
            # 같은 척도(1~4)를 쓰는 문항이 많으면 코드만으로는 티가 안 난다.
            merged, reports = [], []
            for f, d, cl, vl, vt in loaded:
                cl2, vl2, rep = apply_codebook(cl, vl, cb_col, cb_val,
                                               only_broken,
                                               file_encs[0] if file_encs else "")
                merged.append((f, d, cl2, vl2, vt))
                reports.append((f.name, rep))

            total_filled = sum(len(r["filled"]) for _, r in reports)
            ok = sum(r["code_ok"] for _, r in reports)
            bad = sorted({v for _, r in reports for v in r["code_mismatch"]})
            shifted = len(bad) > max(3, ok * 0.1)

            if bad:
                (st.error if shifted else st.warning)(
                    f"⚠️ **코드북에만 있는 코드를 가진 변수가 {len(bad)}개입니다** "
                    f"(코드가 맞는 변수 {ok}개). 짝이 제대로 맞은 코드북은 보통 "
                    "0개입니다. 코드북 행이 밀렸거나 다른 조사의 코드북일 수 "
                    "있습니다 — " + ", ".join(bad[:12])
                    + (" …" if len(bad) > 12 else "")
                )
            else:
                st.caption(
                    f"코드북과 .sav 의 코드가 모두 일치합니다 (변수 {ok}개). "
                    "짝이 제대로 맞았습니다."
                )

            force = False
            if shifted:
                force = st.checkbox(
                    "그래도 적용하기", value=False, key="SX_cb_force",
                    help="엉뚱한 변수에 라벨이 붙을 수 있습니다. 위 목록을 "
                         "확인하고 문제가 없을 때만 켜세요.",
                )
                if not force:
                    st.info(
                        "짝이 어긋나 보여 **적용하지 않았습니다.** 코드북을 "
                        "확인해 다시 올리시거나, 확인을 마쳤으면 위 상자를 켜세요."
                    )

            if not shifted or force:
                loaded[:] = merged
                if total_filled:
                    st.success(f"✅ 라벨 {total_filled}개를 원문으로 채웠습니다.")
                else:
                    st.warning(
                        "채운 라벨이 없습니다. 변수명과 코드가 .sav 와 맞는지 "
                        "확인해 주세요."
                    )

            rows = []
            for fname, rep in reports:
                for var, where in rep["filled"]:
                    rows.append({"파일": fname, "변수": var, "위치": where,
                                 "결과": "채움"})
                for var, where in rep["still_broken"]:
                    rows.append({"파일": fname, "변수": var, "위치": where,
                                 "결과": "코드북에도 없음 — 그대로 잘려 있음"})
            if rows:
                st.dataframe(pd.DataFrame(rows), hide_index=True,
                             use_container_width=True)

            unknown = sorted({v for _, r in reports
                              for v in r["unknown_vars"]})
            if unknown:
                st.caption(
                    f"코드북에는 있지만 이 .sav 에 없는 변수 {len(unknown)}개 — "
                    + ", ".join(unknown[:12])
                    + (" …" if len(unknown) > 12 else "")
                )
            st.caption(
                "채운 라벨은 **엑셀 시트에만** 들어갑니다. SPSS 형식이 120바이트를 "
                "넘는 값 라벨을 담지 못하므로, 내려받는 `.sav` 는 그대로 잘린 "
                "라벨을 씁니다."
            )

if broken_labels:
    st.caption(
        "SPSS 에서 라벨을 40자 이내로 줄여 다시 저장하면 애초에 `?` 가 생기지 "
        "않습니다. 다만 문장 자체가 조사 대상이면 줄이기 어려우니, 위의 "
        "코드북 채우기를 쓰시는 편이 낫습니다."
    )

if multi:
    st.dataframe(
        pd.DataFrame([{
            "파일": f.name,
            "응답자": len(d),
            "변수": len(d.columns),
            "값 레이블 변수": sum(1 for c in d.columns if c in vl),
            "문자형 변수": len(find_text_cols(d, vt)),
        } for f, d, cl, vl, vt in loaded]),
        hide_index=True, use_container_width=True,
    )
    big = [f.name for f, d, *_ in loaded if len(d) > 10_000]
    if big:
        st.warning(
            f"1만 행이 넘는 파일이 {len(big)}개 있습니다. 변환이 느리거나 "
            "메모리 한도에 걸릴 수 있습니다 — " + ", ".join(big[:5])
        )
else:
    up, df, col_labels, value_labels, var_types = loaded[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("응답자 수", f"{len(df):,}")
    c2.metric("변수 수", f"{len(df.columns):,}")
    c3.metric("값 레이블이 있는 변수",
              f"{sum(1 for c in df.columns if c in value_labels):,}")
    c4.metric("문자형 변수", f"{len(find_text_cols(df, var_types)):,}")
    if len(df) > 10_000:
        st.warning(
            f"행이 {len(df):,}개입니다. 1만 행이 넘으면 변환이 느리거나 "
            "메모리 한도에 걸릴 수 있습니다."
        )

st.divider()

# ══════════════════════════════════════════════════════════════
#  1. 엑셀 값 반영 (파일 하나일 때만)
# ══════════════════════════════════════════════════════════════
patched, rep = False, None

if multi:
    st.subheader("1. 엑셀 값 반영")
    st.info(
        "파일이 여러 개일 때는 쓸 수 없습니다. 반영은 파일마다 ID 열과 "
        "변수 짝을 따로 정해야 하기 때문입니다. "
        "값을 반영하려면 SAV 를 하나만 올려 주세요."
    )
else:
    st.subheader("1. 엑셀 값 반영 " + "(선택)")

    work_df, work_labels, work_types = df, value_labels, var_types

    do_patch = st.checkbox(
        "엑셀 파일의 값으로 덮어쓰기",
        value=False,
        help="코딩·수정을 마친 엑셀을 올리면 ID 로 짝을 맞춰 지정한 변수를 "
             "덮어씁니다. 엑셀의 빈칸은 건드리지 않습니다.",
    )

    if do_patch:
        pf = st.file_uploader("수정 값이 든 엑셀 또는 CSV",
                              type=["xlsx", "xls", "csv"], key="SX_patch")
        if not pf:
            st.info("엑셀 파일을 올리면 짝을 맞춰 보여드립니다.")
            st.stop()

        try:
            sheet_names = list_sheets(pf.getvalue(), pf.name)
        except Exception as e:
            st.error(
                f"시트 목록을 읽지 못했습니다: {e}\n\n"
                ".xls 파일이라면 requirements.txt 에 xlrd 가 있는지 확인해 주세요."
            )
            st.stop()

        sheet = 0
        if sheet_names:
            if len(sheet_names) == 1:
                sheet = sheet_names[0]
                st.caption(f"시트: {sheet}")
            else:
                HINTS = ("코딩", "수정", "반영", "결과", "data", "raw")
                guess = next((i for i, s in enumerate(sheet_names)
                              if any(h in str(s).lower() for h in HINTS)), 0)
                sheet = st.selectbox(
                    f"시트 고르기 (총 {len(sheet_names)}개)",
                    sheet_names, index=guess,
                    help="값이 든 시트를 고르세요. 첫 시트가 표지인 경우가 많습니다.",
                )

        try:
            patch = read_table_bytes(pf.getvalue(), pf.name, sheet)
        except Exception as e:
            st.error(f"파일을 읽지 못했습니다: {e}")
            st.stop()

        if patch.empty or not len(patch.columns):
            st.warning("고른 시트가 비어 있습니다. 다른 시트를 골라 주세요.")
            st.stop()

        st.write(f"올리신 파일: {len(patch):,}행 × {len(patch.columns)}열")
        with st.expander("고른 시트 미리보기"):
            st.dataframe(
                patch.head(10).astype(str).replace("None", "").replace("nan", ""),
                hide_index=True, use_container_width=True,
            )

        kc = find_key_cols(df)
        k1, k2 = st.columns(2)
        with k1:
            sav_key = st.selectbox(
                "SAV 의 ID 변수", list(df.columns),
                index=list(df.columns).index(kc[0]) if kc else 0,
            )
        with k2:
            pcols = [str(c) for c in patch.columns]
            g = next((i for i, c in enumerate(pcols)
                      if c.lower() == str(sav_key).lower()), 0)
            patch_key = st.selectbox("엑셀의 ID 열", pcols, index=g)

        # ── 변수 짝 맞추기 (표에서 직접 고칠 수 있다) ──
        st.markdown("**변수 짝 맞추기**")
        st.caption(
            "이름이 같은 변수는 미리 채워 뒀습니다. 'SAV 변수' 칸을 눌러 바꾸거나, "
            "넣지 않을 열은 " + SKIP_LABEL + " 로 두세요."
        )

        guess_df = guess_mapping(df, patch, sav_key, patch_key)
        if guess_df.empty:
            st.warning("ID 열 말고는 열이 없습니다. 다른 시트를 골라 주세요.")
            st.stop()

        edited = st.data_editor(
            guess_df,
            hide_index=True,
            use_container_width=True,
            key=f"SX_map_{pf.name}_{sheet}_{sav_key}_{patch_key}",
            column_config={
                "엑셀 열": st.column_config.TextColumn("엑셀 열", disabled=True),
                "값 예시": st.column_config.TextColumn("값 예시", disabled=True,
                                                    width="medium"),
                "채워진 칸": st.column_config.NumberColumn("채워진 칸",
                                                       disabled=True,
                                                       width="small"),
                "SAV 변수": st.column_config.SelectboxColumn(
                    "SAV 변수", options=[SKIP_LABEL] + list(df.columns),
                    required=True,
                ),
            },
        )

        mapping = dict(zip(edited["엑셀 열"], edited["SAV 변수"]))

        used = [v for v in mapping.values() if v != SKIP_LABEL]
        dups = sorted({v for v in used if used.count(v) > 1})
        if dups:
            st.warning(
                "같은 SAV 변수에 엑셀 열이 둘 이상 연결됐습니다. "
                "표 아래쪽 열이 위쪽을 덮어씁니다 — " + ", ".join(dups)
            )

        with st.spinner("맞춰 보는 중입니다…"):
            new_df, new_labels, rep = apply_patch(
                df, value_labels, var_types, patch, sav_key, patch_key, mapping)

        m1, m2, m3 = st.columns(3)
        m1.metric("짝이 맞은 응답자", f"{rep['matched']:,}")
        m2.metric("덮어쓸 변수", f"{len(rep['changed']):,}")
        m3.metric("바뀌는 셀", f"{sum(rep['changed'].values()):,}")

        if rep["unmatched_ids"]:
            st.warning(
                f"SAV 에 없는 ID {len(rep['unmatched_ids'])}개는 넘겼습니다 — "
                + ", ".join(rep["unmatched_ids"][:10])
                + (" …" if len(rep["unmatched_ids"]) > 10 else "")
            )

        if rep["dup_sav"]:
            st.error(
                f"**SAV 의 `{sav_key}` 에 같은 값이 여러 행에 있습니다 "
                f"({len(rep['dup_sav'])}개 ID).** 각 ID 의 **첫 번째 행에만** "
                "값이 들어가고 나머지 행은 그대로 남습니다. ID 열을 잘못 고른 "
                "것은 아닌지 확인해 주세요 — "
                + ", ".join(rep["dup_sav"][:10])
                + (" …" if len(rep["dup_sav"]) > 10 else "")
            )

        if rep["dup_patch"]:
            st.error(
                f"**엑셀의 `{patch_key}` 에 같은 값이 여러 행에 있습니다 "
                f"({len(rep['dup_patch'])}개 ID).** 같은 ID 끼리는 **아래쪽 행이 "
                "위쪽을 덮어씁니다.** 의도한 것이 아니라면 엑셀에서 정리한 뒤 "
                "다시 올려 주세요 — "
                + ", ".join(rep["dup_patch"][:10])
                + (" …" if len(rep["dup_patch"]) > 10 else "")
            )

        if rep["only_in_patch"]:
            st.info(
                f"연결하지 않아 넘긴 열 {len(rep['only_in_patch'])}개 — "
                + ", ".join(rep["only_in_patch"][:10])
                + (" …" if len(rep["only_in_patch"]) > 10 else "")
            )

        if rep["to_text"]:
            st.warning(
                "문자 값이 섞여 아래 변수는 **문자형으로 바뀝니다**. "
                "SPSS 는 한 변수에 숫자와 문자를 섞을 수 없어서, "
                "다른 응답자의 숫자도 글자가 되고 값 라벨은 버려집니다.\n\n"
                + ", ".join(rep["to_text"])
            )

        if not rep["changed"]:
            st.warning(
                "덮어쓸 값이 없습니다. ID 열이 맞는지, 위 표에서 SAV 변수를 "
                "연결했는지 확인해 주세요."
            )
            st.stop()

        with st.expander(f"변수별 변경 셀 수 ({len(rep['changed'])}개 변수)"):
            st.dataframe(
                pd.DataFrame({
                    "변수": list(rep["changed"].keys()),
                    "바뀌는 셀": list(rep["changed"].values()),
                    "문자형으로 바뀜": ["예" if k in rep["to_text"] else ""
                                    for k in rep["changed"]],
                }),
                hide_index=True, use_container_width=True,
            )

        work_df, work_labels = new_df, new_labels
        work_types = dict(var_types)
        for c in rep["to_text"]:
            work_types[c] = "string"
        patched = True

st.divider()

# ══════════════════════════════════════════════════════════════
#  2. 담을 시트 고르기
# ══════════════════════════════════════════════════════════════
st.subheader("2. 담을 시트 고르기")

if multi:
    key_hint, text_cols = "NO, id", None
else:
    key_cols = find_key_cols(work_df)
    text_cols = find_text_cols(work_df, work_types)
    key_hint = ", ".join(key_cols)

s1, s2, s3, s4, s5 = st.columns(5)
want_raw = s1.checkbox("Raw (숫자 코드)", value=True)
want_label = s2.checkbox("Label (값 레이블)", value=True)
want_open = s3.checkbox(
    "Open (주관식)", value=True,
    help=f"키 변수({key_hint})와 문자형 변수를 담습니다.",
)
want_code = s4.checkbox(
    "Code (코드북)", value=True,
    help=f"변수마다 문항과 코드값/보기를 블록으로 정리합니다. "
         f"키 변수({key_hint})는 제외합니다.",
)
want_guide = s5.checkbox("변수 가이드", value=True)

# ── 머리글 색 ──
PICK_OWN, NO_FILL, RANDOM = "직접 고르기", "색 없음", "랜덤"
choice = st.radio(
    "머리글 색",
    [RANDOM] + list(HEAD_COLORS) + [PICK_OWN, NO_FILL],
    horizontal=True,
    help="Code 시트는 변수명·문항 줄, 나머지 시트는 첫 행에 칠합니다."
         + ("  ‘랜덤’ 은 파일끼리 색이 겹치지 않게 섞어 줍니다." if multi
            else "  ‘랜덤’ 은 프리셋 중 하나를 골라 씁니다."),
)
random_color = choice == RANDOM
if choice == NO_FILL:
    head_color = ""
elif choice == RANDOM:
    head_color = DEFAULT_HEAD_COLOR      # 실제 색은 파일마다 따로 정한다
elif choice == PICK_OWN:
    head_color = st.color_picker("색 고르기", DEFAULT_HEAD_COLOR)
else:
    head_color = HEAD_COLORS[choice]

if random_color:
    st.caption(
        (f"프리셋 {len(HEAD_COLORS)}가지를 섞어 파일마다 하나씩 씁니다. "
         f"파일이 {len(HEAD_COLORS)}개를 넘으면 색이 돌아옵니다.")
        if multi else
        f"만들 때마다 프리셋 {len(HEAD_COLORS)}가지 중 하나를 골라 씁니다."
    )
elif head_color:
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;'
        f'font-size:13px;opacity:.75;margin:2px 0 6px;">'
        f'<span style="display:inline-block;width:74px;height:20px;'
        f'background:{head_color};border:1px solid rgba(128,128,128,.4);'
        f'border-radius:3px;"></span>{head_color.upper()}</div>',
        unsafe_allow_html=True,
    )
else:
    st.caption("색 없이 굵게만 표시됩니다.")

if want_open and text_cols is not None and not text_cols:
    st.info(
        f"문자형 변수가 없어 Open 시트에 키 변수({key_hint})만 담깁니다. "
        "주관식 응답을 옆에 붙여 코딩하실 때 쓰시면 됩니다."
    )

if not (want_raw or want_label or want_open or want_code or want_guide):
    st.warning("시트를 하나 이상 선택해주세요.")
    st.stop()


def build_sheets(d: pd.DataFrame, cl: dict, vlab: dict, vtyp: dict) -> dict:
    """고른 설정대로 시트를 만든다. 파일마다 이 함수를 쓴다."""
    kc = find_key_cols(d)
    tc = find_text_cols(d, vtyp)
    out = {}
    if want_raw:
        out["Raw"] = build_raw(d)
    if want_label:
        out["Label"] = build_label(d, vlab)
    if want_open:
        out["Open"] = build_open(d, tc, kc)
    if want_code:
        out["Code"] = build_codebook(d, cl, vlab, kc)
    if want_guide:
        out["변수 가이드"] = build_guide(d, cl)
    return out


if not multi:
    sheets = build_sheets(work_df, col_labels, work_labels, work_types)
    with st.expander("미리보기", expanded=not patched):
        tabs = st.tabs(list(sheets.keys()))
        for tab, (name, frame) in zip(tabs, sheets.items()):
            with tab:
                st.dataframe(
                    frame.head(20).astype(str).replace("None", ""),
                    use_container_width=True, hide_index=True,
                )
                if len(frame) > 20:
                    st.caption(f"위 20행만 표시 · 전체 {len(frame):,}행")

st.divider()

# ══════════════════════════════════════════════════════════════
#  3. 파일 만들기
# ══════════════════════════════════════════════════════════════
st.subheader("3. 파일 만들기")

if multi:
    st.caption(f"{len(loaded)}개 파일을 같은 설정으로 변환해 zip 하나로 묶습니다.")

if st.button("만들기", type="primary", use_container_width=True):
    if multi:
        buf, used_names = io.BytesIO(), set()
        # 파일마다 다른 색을 주려면 프리셋을 섞어 돌려 쓴다.
        # 섞어 쓰므로 파일 수가 프리셋 수 이하면 색이 겹치지 않는다.
        palette = list(HEAD_COLORS.values())
        random.shuffle(palette)
        picked_colors = []
        bar = st.progress(0.0, text="시작합니다…")
        try:
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, (f, d, cl, vlab, vtyp) in enumerate(loaded, start=1):
                    bar.progress((i - 1) / len(loaded),
                                 text=f"{i}/{len(loaded)} · {f.name}")
                    stem = os.path.splitext(f.name)[0]
                    name = stem + ".xlsx"
                    n = 2
                    while name in used_names:      # 이름이 겹치면 번호를 붙인다
                        name = f"{stem}({n}).xlsx"
                        n += 1
                    used_names.add(name)
                    color = (palette[(i - 1) % len(palette)]
                             if random_color else head_color)
                    picked_colors.append({"파일": name, "머리글 색": color or "없음"})
                    zf.writestr(name,
                                to_excel(build_sheets(d, cl, vlab, vtyp),
                                         color))
            bar.progress(1.0, text="다 됐습니다.")
        except Exception as e:
            st.error(f"파일 생성에 실패했습니다: {e}")
            st.stop()
        st.session_state["SX_zip"] = buf.getvalue()
        st.session_state["SX_zip_name"] = f"SAV_엑셀변환_{len(loaded)}개.zip"
        st.session_state["SX_colors"] = picked_colors if random_color else None
        st.session_state.pop("SX_xlsx", None)
        st.session_state.pop("SX_sav", None)
    else:
        color = (random.choice(list(HEAD_COLORS.values()))
                 if random_color else head_color)
        with st.spinner("파일을 만드는 중입니다…"):
            try:
                st.session_state["SX_xlsx"] = to_excel(sheets, color)
                st.session_state["SX_sav"] = (
                    write_sav(work_df, col_labels, work_labels)
                    if patched else None)
            except Exception as e:
                st.error(f"파일 생성에 실패했습니다: {e}")
                st.stop()
        st.session_state["SX_stem"] = (
            os.path.splitext(up.name)[0] + ("_반영" if patched else ""))
        st.session_state["SX_used_color"] = color if random_color else None
        st.session_state.pop("SX_zip", None)
        st.session_state.pop("SX_colors", None)

if st.session_state.get("SX_zip"):
    st.success("다 됐습니다.")
    if st.session_state.get("SX_colors"):
        with st.expander("파일마다 쓴 색"):
            st.dataframe(pd.DataFrame(st.session_state["SX_colors"]),
                         hide_index=True, use_container_width=True)
    st.download_button(
        "zip 내려받기",
        data=st.session_state["SX_zip"],
        file_name=st.session_state["SX_zip_name"],
        mime="application/zip",
        use_container_width=True,
    )

if st.session_state.get("SX_xlsx"):
    st.success("다 됐습니다.")
    if st.session_state.get("SX_used_color"):
        st.caption(f"머리글 색: {st.session_state['SX_used_color'].upper()}")
    stem = st.session_state.get("SX_stem", "output")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "엑셀 내려받기",
            data=st.session_state["SX_xlsx"],
            file_name=stem + ".xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    if st.session_state.get("SX_sav"):
        with d2:
            st.download_button(
                "SAV 내려받기",
                data=st.session_state["SX_sav"],
                file_name=stem + ".sav",
                mime="application/octet-stream",
                use_container_width=True,
            )

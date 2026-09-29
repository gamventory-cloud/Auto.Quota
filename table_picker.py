# table_picker.py
#
# 테이블 생성 페이지에서 '문항을 고르는' 쪽을 맡는다. 계산은 하지 않는다.
#
#   detect_items  : .sav 를 보고 고를 수 있는 문항 목록을 만든다.
#                   복수응답 세트(Q3_1 ~ Q3_10)는 한 문항으로 묶는다.
#   parse_spec    : 'Q1, Q3_1~Q3_5, Q7' 같은 입력을 문항 목록으로 푼다.
#   find_batteries: 보기가 같은 문항 묶음(척도 종합표 후보)을 찾는다.
#
# 판정 규칙(복제본·복수응답 묶음)은 '엑셀 폼' 자동 채우기와 같은 함수를 쓴다.
# 두 곳이 다르게 판정하면 같은 데이터인데 표가 다르게 나온다.

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

from banner_table_form import (
    _numeric_suffix_family,
    _scale_summary,
    find_duplicates,
    find_ma_sets,
)

KIND_SINGLE, KIND_MULTI, KIND_OBSER = "single", "multi", "obser"
KIND_LABEL = {KIND_SINGLE: "단수", KIND_MULTI: "복수", KIND_OBSER: "수치형"}
LABEL_KIND = {v: k for k, v in KIND_LABEL.items()}


@dataclass
class Item:
    """고를 수 있는 문항 하나. 복수응답 세트면 변수가 여러 개다."""
    key: str                     # 부르는 이름 — 변수명, 복수 세트는 접두사
    kind: str                    # single | multi | obser
    vars: list[str] = field(default_factory=list)
    label: str = ""              # 문항 문구 (변수 라벨)
    summary: str = ""            # 척도 요약 제안 ('상2,중,하2,평균' 등)

    @property
    def display(self) -> str:
        """목록에 보이는 이름."""
        if self.kind == KIND_MULTI:
            tag = f" [복수 {len(self.vars)}개]"
        elif self.kind == KIND_OBSER:
            tag = " [수치]"
        else:
            tag = ""
        lbl = f" · {self.label}" if self.label and self.label != self.key else ""
        return f"{self.key}{tag}{lbl}"


def detect_items(df: pd.DataFrame, meta) -> tuple[list[Item], list[str]]:
    """고를 수 있는 문항 목록 (파일 순서) 과 빠진 변수 안내.

    직접 고르는 목록이라 '엑셀 폼' 자동 채우기와 달리 복제본(bv1 등)이나
    보기가 많은 변수도 남긴다. 계산할 수 없는 문자 변수만 뺀다.
    """
    vl = meta.variable_value_labels
    cl = meta.column_names_to_labels
    ma_sets = find_ma_sets(df, meta, find_duplicates(df))
    in_set = {v: p for p, members in ma_sets.items() for v in members}

    items: list[Item] = []
    done_sets: set[str] = set()
    skipped_text: list[str] = []
    for col in df.columns:
        if col in in_set:
            p = in_set[col]
            if p in done_sets:
                continue
            done_sets.add(p)
            members = ma_sets[p]
            items.append(Item(p, KIND_MULTI, list(members),
                              cl.get(members[0]) or ""))
            continue
        labels = vl.get(col, {})
        if labels:
            items.append(Item(col, KIND_SINGLE, [col], cl.get(col) or "",
                              _scale_summary(labels)))
        elif pd.api.types.is_numeric_dtype(df[col]):
            items.append(Item(col, KIND_OBSER, [col], cl.get(col) or ""))
        else:
            skipped_text.append(col)

    notes = []
    if skipped_text:
        notes.append(f"문자 변수 {len(skipped_text)}개는 집계할 수 없어 목록에서 "
                     f"뺐습니다 ({', '.join(skipped_text[:4])}"
                     f"{'…' if len(skipped_text) > 4 else ''}).")
    return items, notes


def find_batteries(df: pd.DataFrame, meta) -> dict[str, list[str]]:
    """보기가 모두 같은 '앞부분 + 숫자' 묶음 중 복수응답이 아닌 것.

    Q5_1 ~ Q5_5 처럼 같은 척도로 묻는 문항들이다. 척도 종합표 후보로 쓴다.
    """
    vl = meta.variable_value_labels
    ma = find_ma_sets(df, meta, find_duplicates(df))
    out: dict[str, list[str]] = {}
    for prefix, members in _numeric_suffix_family(list(df.columns)).items():
        if prefix in ma:
            continue
        members = [m for m in members if vl.get(m)]
        if len(members) < 2:
            continue
        first = tuple(sorted(vl[members[0]].items()))
        if all(tuple(sorted(vl[m].items())) == first for m in members):
            out[prefix] = members
    return out


# =============================================================================
# 입력 해석 — 'Q1, Q3_1~Q3_5, Q7'
# =============================================================================
def _tokens(text: str) -> list[str]:
    """쉼표·공백·줄바꿈으로 나눈다. 'A ~ B' · 'A to B' 는 한 조각으로 붙인다."""
    text = re.sub(r"\s+to\s+", "~", text or "", flags=re.IGNORECASE)
    text = re.sub(r"\s*~\s*", "~", text)
    return [t for t in re.split(r"[,\s;]+", text) if t]


def parse_spec(text: str, columns: list[str],
               items: list[Item]) -> tuple[list[Item], list[str]]:
    """입력한 문항을 풀어 Item 목록으로 돌려준다. (문항들, 문제들)

      Q1            변수 하나. 변수가 없고 Q1_1, Q1_2 … 가 있으면 그 묶음 전체
                    (복수 세트 이름을 쓰면 세트 전체가 복수 문항 하나)
      Q3_1~Q3_5     파일 순서로 Q3_1 부터 Q3_5 까지  ('Q3_1 to Q3_5' 도 됨)

    대소문자는 가리지 않는다. 복수 세트에 속한 변수를 일부만 적으면
    그 변수들만 묶은 복수 문항이 된다.
    """
    lower = {str(c).lower(): c for c in columns}
    by_key = {it.key.lower(): it for it in items}
    by_var = {v: it for it in items for v in it.vars}
    pos = {c: i for i, c in enumerate(columns)}
    families = {p.lower(): m for p, m in _numeric_suffix_family(list(columns)).items()}

    def name(tok: str):
        return lower.get(tok.lower())

    picked: list[str] = []
    problems: list[str] = []
    for tok in _tokens(text):
        if "~" in tok:
            a, _, b = tok.partition("~")
            ca, cb = name(a), name(b)
            if ca is None or cb is None:
                problems.append(f"'{tok}' — 없는 변수: "
                                + ", ".join(x for x, c in ((a, ca), (b, cb)) if c is None))
                continue
            i0, i1 = sorted((pos[ca], pos[cb]))
            picked += list(columns[i0:i1 + 1])
        elif name(tok) is not None:
            picked.append(name(tok))
        elif tok.lower() in by_key:
            picked += by_key[tok.lower()].vars
        elif tok.lower() in families:     # 'Q1' → Q1_1 ~ Q1_5 (각각 단수)
            picked += families[tok.lower()]
        else:
            problems.append(f"'{tok}' — 이 파일에 없는 변수입니다")

    # 중복 제거 (처음 나온 순서 유지)
    seen: set[str] = set()
    picked = [v for v in picked if not (v in seen or seen.add(v))]

    out: list[Item] = []
    skipped: list[str] = []
    i = 0
    while i < len(picked):
        v = picked[i]
        it = by_var.get(v)
        if it is None:                     # 문자 변수 등 목록에 없는 것
            skipped.append(v)
            i += 1
            continue
        if it.kind != KIND_MULTI:
            out.append(it)
            i += 1
            continue
        # 같은 세트 변수가 이어지면 한 문항으로 묶는다
        j = i
        while j < len(picked) and by_var.get(picked[j]) is it:
            j += 1
        chunk = picked[i:j]
        if chunk == it.vars:
            out.append(it)
        else:
            key = chunk[0] if len(chunk) == 1 else f"{chunk[0]}~{chunk[-1]}"
            out.append(Item(key, KIND_MULTI, chunk, it.label))
        i = j
    if skipped:
        problems.append(f"집계할 수 없는 문자 변수라 뺐습니다: {', '.join(skipped[:6])}"
                        f"{'…' if len(skipped) > 6 else ''}")
    return out, problems


def build_plan(chosen: list[Item], kinds: list[str], titles: list[str],
               summaries: list[str]) -> list[dict]:
    """고른 문항과 편집표 값(유형·제목·척도 요약)으로 만들 표 목록을 짠다.

    kinds 는 'single' 같은 내부 이름이다. 여러 변수 문항을 단수·수치형으로
    바꾸면 변수마다 표가 따로 나온다 (제목은 변수명). 척도 요약은 단수에만 붙는다.
    """
    plan: list[dict] = []
    for it, kind, title, spec in zip(chosen, kinds, titles, summaries):
        kind = kind if kind in KIND_LABEL else it.kind
        title = str(title or "").strip() or it.key
        spec = str(spec or "").strip() if kind == KIND_SINGLE else ""
        if kind == KIND_MULTI or len(it.vars) == 1:
            plan.append({"kind": kind, "vars": list(it.vars),
                         "title": title, "summary": spec})
        else:
            for v in it.vars:
                plan.append({"kind": kind, "vars": [v],
                             "title": v, "summary": spec})
    return plan


def merge_picks(*groups: list[Item]) -> list[Item]:
    """여러 경로로 고른 문항을 합친다. 같은 변수 묶음은 한 번만."""
    seen: set[tuple] = set()
    out = []
    for g in groups:
        for it in g:
            k = tuple(it.vars)
            if k not in seen:
                seen.add(k)
                out.append(it)
    return out

# -*- coding: utf-8 -*-
"""
설문지(.docx) 파서 + 에디팅 신택스 생성기  (v2)

v1 에서 바뀐 점
  v1 은 지시문을 정규식으로 해석해 '필수값=2' 까지 단정했다. 표기가 조금
  다른 문서를 만나면 못 읽은 지시문이 조용히 사라져, 그럴듯하지만 틀린
  신택스가 나왔다. v2 는 파서의 역할을 낮춘다.

    자동    문항 머리글 / 보기 번호와 개수 / 매트릭스 표의 행 수와 척도 폭
            → 문서 구조에서 나오는 것이라 표기 흔들림에 강하다.
    화면    [PROG:] [Range:] (DATA:) 등 모든 지시문
            → 인식하든 못 하든 전부 목록에 올려 사람이 확인한다.
    오류    번호 건너뜀 / 보기 비연속 / 대괄호 짝 / 표 미연결
            → 조용히 보정하지 않고 띄운다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
T_SINGLE = "단일선택"
T_MULTI = "복수응답"
T_MATRIX = "매트릭스"
T_NUM = "숫자기입"
T_TEXT = "주관식"
T_UNKNOWN = "미정"

ALL_TYPES = [T_SINGLE, T_MULTI, T_MATRIX, T_NUM, T_TEXT, T_UNKNOWN]

LV_ERROR = "오류"
LV_CHECK = "확인"


@dataclass
class Option:
    code: int
    text: str
    lo: float | None = None
    hi: float | None = None


@dataclass
class Question:
    qid: str
    title: str = ""
    qtype: str = T_UNKNOWN
    section: str = ""
    options: list[Option] = field(default_factory=list)
    n_items: int = 0                      # 매트릭스 진술문 수
    scale_codes: list[float] = field(default_factory=list)
    item_texts: list[str] = field(default_factory=list)
    table_no: int | None = None           # 붙은 표 번호
    seq: int = 0

    @property
    def codes(self) -> list[int]:
        return [o.code for o in self.options]

    @property
    def contiguous(self) -> bool:
        c = self.codes
        return bool(c) and c == list(range(1, len(c) + 1))

    @property
    def n_expected(self) -> int:
        """이 문항에 배정될 변수 개수의 기본값."""
        if self.qtype == T_MULTI:
            return len(self.options)
        if self.qtype == T_MATRIX:
            return self.n_items
        return 1


@dataclass
class Directive:
    did: str
    scope: str            # 문항 ID 또는 'Q1~Q4'
    raw: str
    kind: str = ""        # 인식 종류 ('' = 미해석)
    params: dict = field(default_factory=dict)
    note: str = ""        # 사람이 읽을 해석 설명


@dataclass
class Issue:
    level: str
    where: str
    what: str


# ---------------------------------------------------------------------------
# 정규식
# ---------------------------------------------------------------------------
QCODE_RE = re.compile(
    r"^(?P<qid>[A-Za-z]{1,4}\d{1,3}(?:[-_][A-Za-z0-9]{1,6})?)\s*\.\s*(?P<rest>.*)$")
BRACKET_RE = re.compile(r"\[([^\]]*)\]")
# '(DATA: (역문항) 전체 동일 응답시 제외)' 처럼 안에 괄호가 한 번 더 들어간다
PAREN_DIR_RE = re.compile(
    r"\(\s*(DATA|PROG|DP)\s*[:：]((?:[^()]|\([^()]*\))*)\)", re.IGNORECASE)
PAREN_UNIT_RE = re.compile(r"\(([^)]*)\)\s*(년|개월|세|점|명|회|시간|일|원)")
OPT_RE = re.compile(r"(?<![\d\-.])(\d{1,2})\s*\)")
SPAN_RE = re.compile(r"(-?\d+)\s*[~∼－—–-]\s*(-?\d+)")
QREF_RANGE_RE = re.compile(
    r"([A-Za-z]{1,4}\d{1,3})\s*[~∼－—–-]\s*([A-Za-z]{0,4}\d{1,3})")

TYPE_HINTS = [
    ("행별", T_MATRIX), ("행 별", T_MATRIX),
    ("모두 선택", T_MULTI), ("복수", T_MULTI), ("중복", T_MULTI),
    ("1개 선택", T_SINGLE), ("한개 선택", T_SINGLE), ("하나 선택", T_SINGLE),
    ("숫자 입력", T_NUM), ("숫자입력", T_NUM),
    ("숫자 기입", T_NUM), ("숫자기입", T_NUM), ("출생연도", T_NUM),
    ("직접 기입", T_TEXT), ("직접기입", T_TEXT), ("직접 입력", T_TEXT),
]
SECTION_WORDS = ("screening", "본질문", "demographics", "스크리닝", "인구통계")


def is_type_hint(text: str) -> bool:
    """유형 표기인지 판정. 지시문을 유형 표기로 오인하면 그 지시문이 사라진다.

    '[PROG: 0년 0개월 기입 불가]' 가 '기입' 때문에 유형 표기로 처리돼
    목록에서 빠지는 일이 있었다. 지시문은 항상 'PROG:' 처럼 콜론을 쓰므로
    콜론이 있으면 유형 표기로 보지 않는다.
    """
    t = text.strip()
    if ":" in t or "：" in t:
        return False
    if len(t) > 20:
        return False
    if re.search(r"불가|금지|제외|진행|필수", t):
        return False
    return any(w in t for w, _ in TYPE_HINTS)


def detect_type(text: str) -> str:
    for word, typ in TYPE_HINTS:
        if word in text:
            return typ
    return T_UNKNOWN


# ---------------------------------------------------------------------------
# 보기 줄 분리
#   v1 은 번호가 연속하지 않으면 줄 전체를 버렸다. v2 는 그대로 담고 오류로 띄운다.
# ---------------------------------------------------------------------------
def split_options(line: str, expect_first: int) -> list[Option] | None:
    hits = list(OPT_RE.finditer(line))
    if not hits:
        return None
    if int(hits[0].group(1)) != expect_first:
        return None            # 본문의 괄호 숫자를 보기로 오인하는 것을 막는다
    out = []
    for i, m in enumerate(hits):
        end = hits[i + 1].start() if i + 1 < len(hits) else len(line)
        text = line[m.end():end].strip(" \t.·:_")
        lo = hi = None
        sm = SPAN_RE.search(text)
        if sm:
            lo, hi = float(sm.group(1)), float(sm.group(2))
        out.append(Option(code=int(m.group(1)), text=text, lo=lo, hi=hi))
    return out


# ---------------------------------------------------------------------------
# 지시문 인식 — 실패해도 버리지 않는다
# ---------------------------------------------------------------------------
def recognize(raw: str) -> tuple[str, dict, str]:
    """(kind, params, note). 인식 못 하면 ('', {}, '')."""
    b = raw.strip()

    # 여러 문항 동일값 → 불량
    if re.search(r"모두\s*동일\s*값", b) or re.search(r"동일\s*한?\s*값.*(?:불량|제외)", b):
        m = QREF_RANGE_RE.search(b)
        if m:
            return ("same_all",
                    {"from": m.group(1).upper(), "to": m.group(2).upper()},
                    f"{m.group(1)}~{m.group(2)} 전체 동일값 금지")
        return "same_all", {}, "동일값 금지 (대상 문항 확인 필요)"

    # 응답란 전부 0 불가
    if re.search(r"모두\s*0.*(?:불가|금지)", b) or re.search(r"0\s*년\s*0\s*개?월", b):
        return "not_all_zero", {}, "응답란 전부 0 금지"

    # 단독 선택 보기
    #   'SQ4=8 응답자는 다른 보기 선택 불가' 에서 8 을 집어야 한다.
    #   숫자를 앞에서부터 찾으면 문항코드 SQ4 의 4 를 잡으므로 '=' 뒤를 먼저 본다.
    if re.search(r"(?:다른\s*보기|다른보기)[^)]*(?:불가|금지)", b):
        m = re.search(r"=\s*(-?\d{1,2})", b)
        if not m:
            m = re.search(r"(\d{1,2})\s*\)", b)
        if m:
            return "exclusive", {"code": int(m.group(1))}, f"보기 {m.group(1)} 단독 선택"
        return "exclusive", {}, "단독 선택 보기 (번호 확인 필요)"

    # 범위 진행조건:  '20세 이상~59세 이하만 조사 진행'  /  '20~60세 진행'
    m = re.search(r"(-?\d+)\s*세?\s*이상\s*[~∼－—–-]?\s*(-?\d+)\s*세?\s*이하", b)
    if not m:
        m = re.search(r"(-?\d+)\s*[~∼－—–-]\s*(-?\d+)\s*세", b)
    if m and re.search(r"진행|조사|참여", b):
        return ("range", {"lo": m.group(1), "hi": m.group(2)},
                f"범위 {m.group(1)}~{m.group(2)}")

    # 특정 보기만 진행
    m = re.search(r"([A-Za-z]{1,4}\d{1,3}(?:[-_]\w+)?)\s*=\s*(-?\d+)[^)]*?만[^)]*진행", b)
    if m:
        return ("require", {"value": m.group(2), "of": m.group(1).upper()},
                f"{m.group(1)} = {m.group(2)} 만 진행")
    m = re.search(r"(\d{1,2})\s*\)[^)]*?만[^)]*진행", b)
    if m:
        return "require", {"value": m.group(1)}, f"보기 {m.group(1)} 만 진행"
    m = re.search(r"([A-Za-z]{1,4}\d{1,3}(?:[-_]\w+)?)\s*=\s*(-?\d+)\s*응답자만", b)
    if m:
        return ("require", {"value": m.group(2), "of": m.group(1).upper()},
                f"{m.group(1)} = {m.group(2)} 만 진행")

    # Range 표기
    m = re.match(r"^\s*Range\s*[:：]\s*(.+)$", b, re.IGNORECASE)
    body = m.group(1) if m else b
    sm = SPAN_RE.search(body)
    if sm and (m or re.search(r"범위|Range", b, re.IGNORECASE)):
        return ("range", {"lo": sm.group(1), "hi": sm.group(2)},
                f"범위 {sm.group(1)}~{sm.group(2)}")
    # '(0~11) 개월' 처럼 단위가 붙은 순수 범위
    if sm and re.search(r"(년|개월|세|점|명|회|시간|일|원)\s*$", b):
        return ("range", {"lo": sm.group(1), "hi": sm.group(2)},
                f"범위 {sm.group(1)}~{sm.group(2)}")

    return "", {}, ""


# ---------------------------------------------------------------------------
# 본문 파싱
# ---------------------------------------------------------------------------
def parse_survey(path: str):
    """(questions, directives, issues, leftovers)"""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(path)
    questions: list[Question] = []
    directives: list[Directive] = []
    issues: list[Issue] = []
    leftovers: list[str] = []
    section = ""
    cur: Question | None = None
    expect_opt = 1
    n_tbl = 0
    seq = 0
    seen: dict[str, int] = {}

    def add_dir(raw: str, scope: str) -> None:
        kind, params, note = recognize(raw)
        sc = scope
        if kind == "same_all":
            m = QREF_RANGE_RE.search(raw)
            if m:
                sc = f"{m.group(1).upper()}~{m.group(2).upper()}"
        directives.append(Directive(
            did=f"D{len(directives)+1:02d}", scope=sc, raw=raw.strip(),
            kind=kind, params=params, note=note))

    def harvest(line: str, scope: str) -> bool:
        """지시문을 걷어낸다. 유형 표기만 있으면 False."""
        got = False
        for b in BRACKET_RE.findall(line):
            if is_type_hint(b) or not b.strip():
                continue
            # 'Range: 0~(본인나이-15)년 / 0~11개월' 처럼 한 대괄호에 응답란별
            # 범위가 여러 개 들어온다. 쪼개지 않으면 첫 범위만 잡히고
            # 나머지가 조용히 사라진다.
            parts = [b]
            if re.match(r"^\s*Range\s*[:：]", b, re.IGNORECASE) and "/" in b:
                parts = [x.strip() for x in b.split("/") if x.strip()]
            for part in parts:
                add_dir(part, scope)
            got = True
        rest = BRACKET_RE.sub(" ", line)
        for m in PAREN_DIR_RE.finditer(rest):
            add_dir(m.group(0), scope)
            got = True
        rest2 = PAREN_DIR_RE.sub(" ", rest)
        for m in PAREN_UNIT_RE.finditer(rest2):
            inner = m.group(1).strip().strip("_ ")
            if inner:
                add_dir(f"{m.group(1)} {m.group(2)}", scope)
                got = True
        return got

    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]

        # ------------------------------------------------------------ 표
        if tag == "tbl":
            tb = Table(child, doc)
            n_tbl += 1
            seq += 1
            rows, cols = len(tb.rows), len(tb.columns)
            first = tb.rows[0].cells[0].text.strip()

            if rows == 1 or cols < 2:
                if any(w in first.lower() for w in SECTION_WORDS):
                    section = first
                continue

            if cur is None:
                issues.append(Issue(LV_CHECK, f"표 {n_tbl}",
                                    "앞에 문항이 없어 연결하지 못했습니다."))
                continue

            if cur.qtype not in (T_MATRIX, T_UNKNOWN):
                issues.append(Issue(
                    LV_CHECK, f"{cur.qid} / 표 {n_tbl}",
                    f"문항 유형이 '{cur.qtype}' 인데 표가 나왔습니다. 유형을 확인해 주세요."))
                continue

            if cur.table_no is not None:
                issues.append(Issue(
                    LV_ERROR, cur.qid,
                    f"표가 둘 이상 붙었습니다 (표 {cur.table_no}, {n_tbl}). "
                    "문항 머리글이 누락됐을 수 있습니다."))
                continue

            hdr = [c.text.strip() for c in tb.rows[0].cells][1:]
            codes: list[float] = []
            for h in hdr:
                # '전혀\n그렇지\n않다\n1' 처럼 라벨 끝에 코드가 붙어 있다
                mm = re.search(r"(-?\d+)\s*$", h.replace("\n", " ").strip())
                if not mm:
                    codes = []
                    break
                codes.append(float(mm.group(1)))
            if not codes:
                codes = [float(i) for i in range(1, len(hdr) + 1)]
                issues.append(Issue(
                    LV_CHECK, f"{cur.qid} / 표 {n_tbl}",
                    f"척도 머리글에서 코드를 읽지 못해 1~{len(hdr)} 로 두었습니다."))
            elif codes != [float(x) for x in
                           range(int(codes[0]), int(codes[0]) + len(codes))]:
                issues.append(Issue(
                    LV_ERROR, f"{cur.qid} / 표 {n_tbl}",
                    f"척도 코드가 연속하지 않습니다: {[int(c) for c in codes]}"))

            cur.scale_codes = codes
            cur.item_texts = [r.cells[0].text.strip() for r in tb.rows[1:]]
            cur.n_items = len(cur.item_texts)
            cur.table_no = n_tbl
            if cur.qtype == T_UNKNOWN:
                cur.qtype = T_MATRIX
            continue

        if tag != "p":
            continue

        line = Paragraph(child, doc).text.strip()
        if not line:
            continue
        seq += 1

        # 대괄호 짝 검사 — '행별 1개 선택]' 처럼 여는 괄호가 빠진 경우
        if line.count("[") != line.count("]"):
            issues.append(Issue(
                LV_ERROR, cur.qid if cur else f"{seq}번째 줄",
                f"대괄호 짝이 맞지 않습니다: {line[:60]}"))

        # ------------------------------------------------- 문항 머리글
        m = QCODE_RE.match(line)
        if m:
            qid = m.group("qid").upper()
            rest = m.group("rest")
            if qid in seen:
                issues.append(Issue(LV_ERROR, qid, "같은 문항 코드가 두 번 나옵니다."))
            seen[qid] = seq
            cur = Question(
                qid=qid,
                title=BRACKET_RE.sub("", rest).strip(),
                qtype=detect_type(rest),
                section=section,
                seq=seq,
            )
            questions.append(cur)
            harvest(line, qid)
            expect_opt = 1
            continue

        if cur is None:
            leftovers.append(line)
            continue

        # ------------------------------------------------- 보기
        opts = split_options(line, expect_opt)
        if opts is not None:
            cur.options.extend(opts)
            expect_opt = cur.options[-1].code + 1
            harvest(line, cur.qid)
            if cur.qtype == T_UNKNOWN:
                cur.qtype = T_SINGLE
            continue

        # ---------------------------- 유형 표기가 다음 줄에 있는 경우
        typ = detect_type(line)
        if typ != T_UNKNOWN and cur.qtype == T_UNKNOWN:
            cur.qtype = typ
            if not cur.title:
                cur.title = BRACKET_RE.sub("", line).strip()
            harvest(line, cur.qid)
            continue

        if harvest(line, cur.qid):
            continue

        if not cur.title:
            cur.title = line
            continue

        leftovers.append(line)

    # ------------------------------------------------------ 사후 점검
    for q in questions:
        if q.options and not q.contiguous:
            issues.append(Issue(
                LV_ERROR, q.qid,
                f"보기 번호가 1부터 연속하지 않습니다: {q.codes}. "
                "설문지 오타인지 실제 코딩인지 데이터로 확인해 주세요."))
        if q.qtype == T_MATRIX and q.table_no is None:
            issues.append(Issue(LV_ERROR, q.qid, "매트릭스인데 표를 찾지 못했습니다."))
        if q.qtype == T_UNKNOWN:
            issues.append(Issue(LV_CHECK, q.qid, "유형을 판단하지 못했습니다. 직접 골라 주세요."))

    # 같은 문항에 응답란이 여럿이면 (년/개월) 순서대로 변수를 배정해야 한다.
    # 자리 계산을 '인식 성공' 기준으로 하면, 앞의 년 범위가 미해석일 때
    # 개월 범위가 첫 변수에 붙어 조용히 틀린 조건이 나온다.
    UNIT_TAIL = re.compile(r"(년|개월|세|점|명|회|시간|일|원)\s*$")
    slot: dict[str, int] = {}
    for d in directives:
        looks_range = bool(UNIT_TAIL.search(d.raw)) or bool(
            re.match(r"^\s*Range\s*[:：]", d.raw, re.IGNORECASE))
        if not looks_range:
            continue
        d.params["idx"] = slot.get(d.scope, 0)
        slot[d.scope] = d.params["idx"] + 1

    # 번호 건너뜀
    groups: dict[str, list[int]] = {}
    for q in questions:
        mm = re.match(r"^([A-Za-z]+)(\d+)", q.qid)
        if mm:
            groups.setdefault(mm.group(1), []).append(int(mm.group(2)))
    for pre, nums in groups.items():
        uniq = sorted(set(nums))
        gaps = [n for n in range(uniq[0], uniq[-1] + 1) if n not in uniq]
        if gaps:
            issues.append(Issue(
                LV_ERROR, pre,
                "번호가 건너뜁니다 — 없는 문항: "
                + ", ".join(f"{pre}{g}" for g in gaps)))

    return questions, directives, issues, leftovers


# ---------------------------------------------------------------------------
# 변수 배정
# ---------------------------------------------------------------------------
def expand_vars(start: str, n: int) -> list[str]:
    if not start:
        return []
    if n <= 1:
        return [start]
    m = re.match(r"^(?P<base>.*?)(?P<num>\d+)$", start)
    if not m:
        return [start]
    base, num = m.group("base"), int(m.group("num"))
    return [f"{base}{num + i}" for i in range(n)]


def mapping_skeleton(questions: list[Question]) -> list[dict]:
    rows = []
    for q in questions:
        if q.qtype == T_MATRIX:
            info = f"진술문 {q.n_items} · 척도 {len(q.scale_codes)}"
        elif q.qtype in (T_MULTI, T_SINGLE):
            info = f"보기 {len(q.options)}"
        elif q.qtype == T_NUM:
            info = "숫자"
        else:
            info = ""
        rows.append({
            "문항": q.qid, "유형": q.qtype, "개수": q.n_expected,
            "시작변수": "", "구조": info, "문항명": (q.title or "")[:44],
        })
    return rows


# ---------------------------------------------------------------------------
# 지시문 → 조건식 제안
# ---------------------------------------------------------------------------
def suggest_cond(d: Directive, mapping: dict[str, dict],
                 questions: dict[str, Question] | None = None) -> tuple[str, str]:
    """(제안 조건식, 표시할 변수). 만들 수 없으면 ('', '')."""
    questions = questions or {}

    def V(qid: str) -> list[str]:
        return mapping.get(qid, {}).get("vars", []) or []

    v = V(d.scope)

    if d.kind == "range":
        lo, hi = d.params.get("lo"), d.params.get("hi")
        idx = int(d.params.get("idx", 0) or 0)
        if v and lo is not None and hi is not None and idx < len(v):
            # 응답란이 여럿이면 (년/개월) 순번대로 변수를 고른다
            return f"~Range({v[idx]},{lo},{hi})", v[idx]

    elif d.kind == "require":
        of = d.params.get("of") or d.scope
        tv = V(of)
        val = d.params.get("value")
        if tv and val is not None:
            q = questions.get(of)
            # 복수응답은 보기별로 변수가 따로 있으므로 그 보기의 변수를 봐야 한다
            if q is not None and q.qtype == T_MULTI:
                k = int(val)
                if 1 <= k <= len(tv):
                    return f"~Any({tv[k-1]},{k})", tv[k - 1]
                return "", ""
            return f"~Any({tv[0]},{val})", tv[0]

    elif d.kind == "exclusive":
        k = int(d.params.get("code", 0) or 0)
        if v and 1 <= k <= len(v):
            others = ",".join(v[:k - 1] + v[k:])
            if others:
                return f"{v[k-1]}>0 and max({others})>=0", f"{v[0]} to {v[-1]}"

    elif d.kind == "require_sel":
        k = int(d.params.get("code", 0) or 0)
        if v and 1 <= k <= len(v):
            return f"~Any({v[k-1]},{k})", v[k - 1]

    elif d.kind == "not_all_zero":
        if len(v) >= 2:
            return " and ".join(f"{x}=0" for x in v), " ".join(v)

    elif d.kind == "same_all":
        va = V(d.params.get("from", ""))
        vb = V(d.params.get("to", ""))
        if va and vb:
            span = f"{va[0]} to {vb[-1]}"
            return f"(max({span})-min({span})=0)", span

    return "", ""


# ---------------------------------------------------------------------------
# 신택스 생성
# ---------------------------------------------------------------------------
def _blk(cond: str, listvars: str, comment: str = "") -> str:
    head = f"* {comment}.\n" if comment else ""
    return f"{head}Temp.\nSelect If {cond} .\nList Var no id {listvars}  .\n"


def build_syntax(
    questions: list[Question],
    mapping: dict[str, dict],
    logic_blocks: list[tuple[str, str, str]],
    *,
    project: str = "",
    src_sav: str = "",
    first_var: str = "",
    last_var: str = "",
    straightline_only: set[str] | None = None,
    qsort: dict[str, list[int]] | None = None,
    extra_checks: str = "",
) -> tuple[str, list[str]]:
    """logic_blocks: [(라벨, 조건식, 표시변수)] — 지시문 화면에서 확정된 논리 검사"""
    qsort = qsort or {}
    straightline_only = straightline_only or set()
    L: list[str] = []
    warn: list[str] = []

    def V(qid: str) -> list[str]:
        return mapping.get(qid, {}).get("vars", []) or []

    if project:
        L.append(f'CD "{project}".\n')
    if src_sav:
        L.append(f"Get file='{src_sav}'.\n")
    L.append("*_ Check Syntax _______________________________________________________________.\n")
    L.append("SET PRINTBACK=ON.")
    L.append("SET Length=None Width=255.")
    if first_var and last_var:
        L.append(f"Recode {first_var} to {last_var} (SYSMIS=-1) .")
    else:
        warn.append("결측→-1 리코드 범위가 지정되지 않았습니다.")
    L.append("")

    # 음수 척도 매트릭스 사전 리코드
    for q in questions:
        v = V(q.qid)
        if q.qtype == T_MATRIX and len(v) > 1 and q.scale_codes and min(q.scale_codes) < 1:
            maps = "".join(f"({int(c)}={i+1})" for i, c in enumerate(q.scale_codes))
            L.append(f"RECODE {v[0]} to {v[-1]} {maps}.")
            L.append("")

    # ---- 구조에서 나오는 범위 검사 ---------------------------------------
    L.append("*_ 범위 Check _________________________________________________________________.\n")
    for q in questions:
        v = V(q.qid)
        if not v:
            if q.qtype not in (T_TEXT, T_UNKNOWN):
                warn.append(f"{q.qid}: 변수가 배정되지 않아 건너뜀")
            continue

        if q.qtype == T_SINGLE and q.options:
            if q.contiguous:
                L.append(_blk(f"~Range({v[0]},1,{len(q.options)})", v[0], q.qid))
            else:
                codes = ",".join(str(c) for c in q.codes)
                L.append(_blk(f"~Any({v[0]},{codes})", v[0],
                              f"{q.qid} 보기 번호 비연속 — 확인 필요"))
                warn.append(
                    f"{q.qid}: 보기 번호가 {q.codes} 라서 ANY 검사로 만들었습니다. "
                    "데이터 코딩과 맞는지 확인해 주세요.")

        elif q.qtype == T_MULTI and len(v) >= 2:
            span = f"{v[0]} to {v[-1]}"
            L.append(_blk(f"(max({span})<=0)", span, f"{q.qid} 전체 미선택"))

        elif q.qtype == T_MATRIX and q.scale_codes:
            span = f"{v[0]} to {v[-1]}" if len(v) > 1 else v[0]
            n = len(q.scale_codes)
            L.append(_blk(f"(max({span})>{n} or min({span})<1)", span, q.qid))
            if q.qid in straightline_only and len(v) > 1:
                L.append(_blk(f"(max({span})-min({span})=0)", span, f"{q.qid} 직진성"))
            dist = qsort.get(q.qid)
            if dist:
                if len(dist) != n:
                    warn.append(f"{q.qid}: 강제분포 길이 {len(dist)} ≠ 척도 {n}")
                elif sum(dist) != q.n_items:
                    warn.append(f"{q.qid}: 강제분포 합계 {sum(dist)} ≠ 진술문 {q.n_items}")
                else:
                    stem = re.sub(r"[\W_]*\d+$", "", v[0]) or "q"
                    cvar = f"v{stem}"
                    for i in range(1, n + 1):
                        L.append(f"Count {cvar}_{i}={span} ({i}).")
                    L.append("")
                    for i, cnt in enumerate(dist, start=1):
                        L.append(_blk(f"~Any({cvar}_{i},{cnt})", f"{cvar}_{i}",
                                      f"{q.qid} 강제분포 {i}점 {cnt}개"))

        elif q.qtype == T_TEXT:
            warn.append(f"{q.qid}: 주관식 — 눈으로 확인 후 Error 지정이 필요합니다")

    # ---- 지시문에서 나오는 논리 검사 -------------------------------------
    used = [b for b in logic_blocks if b[1].strip()]
    if used:
        L.append("*_ Logic Check ________________________________________________________________.\n")
        for label, cond, lv in used:
            L.append(_blk(cond.strip(), lv.strip(), label))

    if extra_checks.strip():
        L.append("*_ 추가 Check _________________________________________________________________.\n")
        L.append(extra_checks.strip())
        L.append("")

    if first_var and last_var:
        L.append(f"Recode {first_var} to {last_var} (-1=SYSMIS) .")
    L.append("SET PRINTBACK=OFF.")
    L.append("")
    L.append("*_ 객관식 에러 데이터 구분 ____________________________________________________.\n")
    L.append("compute error = -1.")
    L.append("*if (no =  ) Error = 1.")
    L.append("")
    L.append("*_ 주관식 에러 데이터 구분 ____________________________________________________.")
    L.append("** OUT.")
    L.append("*if (no =  ) Error = 2.")
    L.append("")
    L.append("** 불성실.")
    L.append("*if (no =  ) Error = 3.")
    L.append("")

    return re.sub(r"\n{3,}", "\n\n", "\n".join(L)), warn

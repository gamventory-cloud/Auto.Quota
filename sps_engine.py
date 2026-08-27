# -*- coding: utf-8 -*-
"""
SPSS 에디팅 신택스를 읽어 데이터에 그대로 실행하는 엔진 (프로토타입).

목적
  SPSS 는 TEMP / SELECT IF / LIST VAR 블록마다 별도 출력을 뿌리므로,
  체크가 40개면 40개의 목록을 눈으로 훑어 케이스 번호를 모아야 한다.
  이 엔진은 같은 조건을 pandas 로 평가해서
  '케이스 × 걸린 체크' 한 장으로 합친다.

지원 범위 (의도적으로 좁게)
  RECODE (값 매핑, SYSMIS, ELSE=COPY INTO), COUNT, COMPUTE(상수),
  IF (단순 대입), TEMP / SELECT IF / LIST VAR
  함수: ANY, RANGE, MAX, MIN, MISSING, SYSMIS
  연산: ~ NOT AND OR = <> > >= < <= + - * /
  변수 범위: 'a TO b' (파일 열 순서 기준)

SPSS 결측 의미론을 3값 논리로 재현한다. 이 부분을 numpy 기본 동작으로
두면 결측이 있는 케이스에서 조용히 다른 답이 나온다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# 3값 논리 (T / F / 결측)
#   float Series 로 표현한다: 1.0 = 참, 0.0 = 거짓, NaN = 결측
#   SPSS 는 SELECT IF 에서 결측을 '선택하지 않음' 으로 처리한다.
# ---------------------------------------------------------------------------
def t3_not(a: pd.Series) -> pd.Series:
    return 1.0 - a


def t3_and(a: pd.Series, b: pd.Series) -> pd.Series:
    # 하나라도 거짓이면 거짓. 그 외에 결측이 있으면 결측.
    out = pd.Series(np.nan, index=a.index, dtype=float)
    false_ = (a == 0) | (b == 0)
    true_ = (a == 1) & (b == 1)
    out[true_] = 1.0
    out[false_] = 0.0
    return out


def t3_or(a: pd.Series, b: pd.Series) -> pd.Series:
    # 하나라도 참이면 참. 그 외에 결측이 있으면 결측.
    out = pd.Series(np.nan, index=a.index, dtype=float)
    true_ = (a == 1) | (b == 1)
    false_ = (a == 0) & (b == 0)
    out[false_] = 0.0
    out[true_] = 1.0
    return out


def t3_cmp(op: str, x: pd.Series, y: pd.Series) -> pd.Series:
    valid = x.notna() & y.notna()
    with np.errstate(invalid="ignore"):
        if op == "=":
            r = x == y
        elif op == "<>":
            r = x != y
        elif op == ">":
            r = x > y
        elif op == ">=":
            r = x >= y
        elif op == "<":
            r = x < y
        elif op == "<=":
            r = x <= y
        else:
            raise ValueError(f"모르는 연산자: {op}")
    out = pd.Series(np.nan, index=x.index, dtype=float)
    out[valid] = r[valid].astype(float)
    return out


# ---------------------------------------------------------------------------
# 토크나이저
# ---------------------------------------------------------------------------
TOKEN_RE = re.compile(
    r"""
    (?P<num>-?\d+\.?\d*)
  | (?P<op><>|>=|<=|=|>|<|\+|-|\*|/|~|\(|\)|,)
  | (?P<name>[A-Za-z_$][A-Za-z0-9_.$]*)
  | (?P<ws>\s+)
    """,
    re.VERBOSE,
)


def tokenize(s: str) -> list[tuple[str, str]]:
    out, pos = [], 0
    while pos < len(s):
        m = TOKEN_RE.match(s, pos)
        if not m:
            raise SyntaxError(f"해석할 수 없는 문자: {s[pos:pos+20]!r}")
        pos = m.end()
        if m.lastgroup == "ws":
            continue
        out.append((m.lastgroup, m.group()))
    return out


# ---------------------------------------------------------------------------
# 파서 (재귀 하강) + 평가기
#   파싱과 평가를 한 번에 한다. 조건식이 짧고 한 번만 쓰이므로
#   AST 를 따로 만들 이득이 없다.
# ---------------------------------------------------------------------------
class Ctx:
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.lookup = {c.lower(): c for c in df.columns}
        self.order = [c.lower() for c in df.columns]

    def col(self, name: str) -> pd.Series:
        key = name.lower()
        if key not in self.lookup:
            raise KeyError(name)
        return pd.to_numeric(self.df[self.lookup[key]], errors="coerce")

    def has(self, name: str) -> bool:
        return name.lower() in self.lookup

    def expand(self, a: str, b: str) -> list[str]:
        """'v5 TO v12' → 파일 열 순서 기준 구간."""
        ia, ib = self.order.index(a.lower()), self.order.index(b.lower())
        if ia > ib:
            ia, ib = ib, ia
        return [self.lookup[c] for c in self.order[ia:ib + 1]]


class Parser:
    def __init__(self, tokens, ctx: Ctx):
        self.tk = tokens
        self.i = 0
        self.ctx = ctx

    # -- 도구 -------------------------------------------------------------
    def peek(self):
        return self.tk[self.i] if self.i < len(self.tk) else (None, None)

    def next(self):
        t = self.peek()
        self.i += 1
        return t

    def accept_op(self, *vals):
        k, v = self.peek()
        if k == "op" and v in vals:
            self.i += 1
            return v
        return None

    def accept_kw(self, *words):
        k, v = self.peek()
        if k == "name" and v.upper() in words:
            self.i += 1
            return v.upper()
        return None

    def expect_op(self, val):
        if self.accept_op(val) is None:
            raise SyntaxError(f"'{val}' 를 기대했습니다 (위치 {self.i})")

    # -- 문법 -------------------------------------------------------------
    def parse(self) -> pd.Series:
        r = self.or_expr()
        if self.i != len(self.tk):
            raise SyntaxError(f"해석되지 않은 꼬리: {self.tk[self.i:]}")
        return r

    def or_expr(self):
        left = self.and_expr()
        while self.accept_kw("OR"):
            left = t3_or(left, self.and_expr())
        return left

    def and_expr(self):
        left = self.not_expr()
        while self.accept_kw("AND"):
            left = t3_and(left, self.not_expr())
        return left

    def not_expr(self):
        if self.accept_op("~") or self.accept_kw("NOT"):
            return t3_not(self.not_expr())
        return self.comparison()

    def comparison(self):
        left = self.arith()
        op = self.accept_op("=", "<>", ">=", "<=", ">", "<")
        if op is None:
            # 논리값을 그대로 쓴 경우 (예: MISSING(x) 단독)
            return left
        return t3_cmp(op, left, self.arith())

    def arith(self):
        left = self.term()
        while True:
            op = self.accept_op("+", "-")
            if not op:
                return left
            r = self.term()
            left = left + r if op == "+" else left - r

    def term(self):
        left = self.factor()
        while True:
            op = self.accept_op("*", "/")
            if not op:
                return left
            r = self.factor()
            left = left * r if op == "*" else left / r

    def factor(self):
        if self.accept_op("("):
            r = self.or_expr()
            self.expect_op(")")
            return r
        if self.accept_op("-"):
            return -self.factor()

        k, v = self.next()
        if k == "num":
            return pd.Series(float(v), index=self.ctx.df.index, dtype=float)
        if k != "name":
            raise SyntaxError(f"예상 밖 토큰: {v!r}")

        up = v.upper()
        if self.peek() == ("op", "("):
            return self.call(up)
        if not self.ctx.has(v):
            raise KeyError(v)
        return self.ctx.col(v)

    # -- 함수 -------------------------------------------------------------
    def arg_list(self) -> list:
        """인수 목록. 'a TO b' 를 만나면 열 이름 리스트로 펼친다."""
        self.expect_op("(")
        args = []
        while True:
            k, v = self.peek()
            if k == "name" and self.ctx.has(v):
                save = self.i
                self.i += 1
                if self.accept_kw("TO"):
                    k2, v2 = self.next()
                    if k2 != "name":
                        raise SyntaxError("TO 뒤에 변수명이 필요합니다")
                    args.append(("cols", self.ctx.expand(v, v2)))
                else:
                    self.i = save
                    args.append(("val", self.or_expr()))
            else:
                args.append(("val", self.or_expr()))
            if self.accept_op(","):
                continue
            self.expect_op(")")
            return args

    def call(self, fname: str):
        args = self.arg_list()

        def flat_cols():
            cols = []
            for kind, a in args:
                if kind == "cols":
                    cols.extend(a)
                else:
                    raise SyntaxError(f"{fname} 인수에 식은 쓸 수 없습니다")
            return cols

        def series_list():
            out = []
            for kind, a in args:
                if kind == "cols":
                    out.extend(self.ctx.col(c) for c in a)
                else:
                    out.append(a)
            return out

        idx = self.ctx.df.index

        if fname == "ANY":
            x = args[0][1]
            vals = [a for k, a in args[1:]]
            hit = pd.Series(False, index=idx)
            for v in vals:
                hit |= (x == v).fillna(False)
            out = pd.Series(np.nan, index=idx, dtype=float)
            ok = x.notna()
            out[ok] = hit[ok].astype(float)
            # 결측이라도 일치가 있으면 참 (SPSS ANY 동작)
            out[hit] = 1.0
            return out

        if fname == "RANGE":
            x = args[0][1]
            rest = [a for k, a in args[1:]]
            if len(rest) % 2:
                raise SyntaxError("RANGE 는 (변수, 하한, 상한, ...) 형태여야 합니다")
            hit = pd.Series(False, index=idx)
            valid = x.notna()
            for lo, hi in zip(rest[0::2], rest[1::2]):
                valid &= lo.notna() & hi.notna()
                hit |= ((x >= lo) & (x <= hi)).fillna(False)
            out = pd.Series(np.nan, index=idx, dtype=float)
            out[valid] = hit[valid].astype(float)
            return out

        if fname in ("MAX", "MIN"):
            frame = pd.concat(series_list(), axis=1)
            return frame.max(axis=1) if fname == "MAX" else frame.min(axis=1)

        if fname in ("MISSING", "SYSMIS"):
            x = args[0][1]
            return x.isna().astype(float)

        if fname == "SUM":
            return pd.concat(series_list(), axis=1).sum(axis=1, min_count=1)

        if fname == "MEAN":
            return pd.concat(series_list(), axis=1).mean(axis=1)

        raise SyntaxError(f"지원하지 않는 함수: {fname}")


def eval_cond(expr: str, ctx: Ctx) -> pd.Series:
    return Parser(tokenize(expr), ctx).parse()


# ---------------------------------------------------------------------------
# 신택스 문장 쪼개기
# ---------------------------------------------------------------------------
def read_sps(path: str) -> str:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp949", errors="replace")


def split_statements(text: str) -> list[str]:
    """'.' 로 끝나는 줄을 문장 끝으로 본다. 주석(*로 시작)은 버린다."""
    stmts, buf = [], []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if not buf and (s.startswith("*") or s.startswith("!")):
            # 주석 / 매크로 호출은 종결 '.' 까지 통째로 건너뛴다
            if s.endswith("."):
                continue
            buf = ["<<SKIP>>"]
            continue
        if buf and buf[0] == "<<SKIP>>":
            if s.endswith("."):
                buf = []
            continue
        buf.append(s)
        if s.endswith("."):
            stmts.append(" ".join(buf)[:-1].strip())
            buf = []
    return [s for s in stmts if s]


# ---------------------------------------------------------------------------
# 실행기
# ---------------------------------------------------------------------------
@dataclass
class Check:
    seq: int
    cond: str
    list_vars: list[str] = field(default_factory=list)
    n_hit: int = 0
    cases: list = field(default_factory=list)
    error: str | None = None


RECODE_RE = re.compile(r"^RECODE\s+(?P<vars>.+?)\s*(?P<maps>\(.+\))\s*(?:INTO\s+(?P<into>[\w.$]+))?$",
                       re.IGNORECASE | re.DOTALL)
COUNT_RE = re.compile(r"^COUNT\s+(?P<new>[\w.$]+)\s*=\s*(?P<vars>.+?)\s*\((?P<vals>[^)]*)\)$",
                      re.IGNORECASE)
IF_RE = re.compile(r"^IF\s*\((?P<cond>.+?)\)\s*(?P<target>[\w.$]+)\s*=\s*(?P<val>-?[\d.]+)$",
                   re.IGNORECASE)
COMPUTE_RE = re.compile(r"^COMPUTE\s+(?P<target>[\w.$]+)\s*=\s*(?P<expr>.+)$", re.IGNORECASE)
SELECT_RE = re.compile(r"^SELE(?:CT)?\s+IF\s+(?P<cond>.+)$", re.IGNORECASE)
LIST_RE = re.compile(r"^LIST\s+(?:VAR(?:IABLES)?\s*=?\s*)?(?P<vars>.*)$", re.IGNORECASE)


def parse_var_spec(spec: str, ctx: Ctx) -> list[str]:
    """'v5 to v12' 또는 'v1 v2 v3' 또는 'v5, v6' → 열 이름 리스트."""
    toks = re.split(r"[\s,]+", spec.strip())
    cols, i = [], 0
    while i < len(toks):
        t = toks[i]
        if not t:
            i += 1
            continue
        if i + 2 < len(toks) and toks[i + 1].upper() == "TO":
            cols.extend(ctx.expand(t, toks[i + 2]))
            i += 3
        else:
            if ctx.has(t):
                cols.append(ctx.lookup[t.lower()])
            i += 1
    return cols


def parse_maps(maps: str):
    """'(-4=1)(-3=2)(SYSMIS=-1)(ELSE=COPY)' → [(from, to), ...]"""
    out = []
    for grp in re.findall(r"\(([^)]*)\)", maps):
        if "=" not in grp:
            continue
        lhs, rhs = grp.split("=", 1)
        out.append((lhs.strip().upper(), rhs.strip().upper()))
    return out


def run(sps_path: str, df: pd.DataFrame, id_cols=("no", "id")):
    df = df.copy()
    ctx = Ctx(df)
    checks: list[Check] = []
    notes: list[str] = []
    pending_cond: str | None = None
    seq = 0

    for stmt in split_statements(read_sps(sps_path)):
        head = stmt.split()[0].upper() if stmt.split() else ""

        # ---- 무시하는 명령 -------------------------------------------
        if head in {"CD", "GET", "SET", "TEMP", "TEMPORARY", "EXECUTE",
                    "SAVE", "CROSSTAB", "CROSSTABS", "FILTER", "WEIGHT",
                    "MISSING", "VARIABLE", "VALUE", "FORMATS", "DATASET"}:
            continue

        # ---- SELECT IF : 다음 프로시저의 조건 -------------------------
        m = SELECT_RE.match(stmt)
        if m:
            pending_cond = m.group("cond").strip()
            continue

        # ---- LIST : 앞의 SELECT IF 를 체크로 확정 ---------------------
        m = LIST_RE.match(stmt)
        if m:
            if pending_cond is None:
                continue
            seq += 1
            chk = Check(seq=seq, cond=pending_cond,
                        list_vars=parse_var_spec(m.group("vars"), ctx))
            try:
                res = eval_cond(pending_cond, ctx)
                hit = (res == 1.0).fillna(False)
                chk.n_hit = int(hit.sum())
                key = next((c for c in df.columns if c.lower() in id_cols), df.columns[0])
                chk.cases = df.loc[hit, key].tolist()
            except Exception as e:  # noqa: BLE001
                chk.error = f"{e.__class__.__name__}: {e}"
            checks.append(chk)
            pending_cond = None
            continue

        # ---- RECODE ---------------------------------------------------
        m = RECODE_RE.match(stmt)
        if m:
            cols = parse_var_spec(m.group("vars"), ctx)
            maps = parse_maps(m.group("maps"))
            into = m.group("into")
            if into:
                if len(cols) != 1:
                    notes.append(f"INTO 는 변수 1개만 지원: {stmt[:50]}")
                    continue
                df[into] = df[cols[0]]
                ctx = Ctx(df)
                continue
            # RECODE 는 원본 값을 기준으로 '한 번만' 매핑한다.
            # 순차 적용하면 (-4=1)...(1=6) 에서 -4→1 이 뒤의 규칙에 다시 걸려
            # 6 이 되어 버린다. 그래서 원본으로 마스크를 만들어 동시에 대입한다.
            for c in cols:
                src = pd.to_numeric(df[c], errors="coerce")
                dst = src.copy()
                for lhs, rhs in maps:
                    if lhs == "ELSE":
                        continue
                    tgt = np.nan if rhs == "SYSMIS" else float(rhs)
                    if lhs in ("SYSMIS", "MISSING"):
                        dst = dst.where(src.notna(), tgt)
                    else:
                        dst = dst.mask(src == float(lhs), tgt)
                df[c] = dst
            ctx = Ctx(df)
            continue

        # ---- COUNT ----------------------------------------------------
        m = COUNT_RE.match(stmt)
        if m:
            cols = parse_var_spec(m.group("vars"), ctx)
            vals = [float(v) for v in re.split(r"[\s,]+", m.group("vals").strip()) if v]
            sub = df[cols].apply(pd.to_numeric, errors="coerce")
            df[m.group("new")] = sub.isin(vals).sum(axis=1).astype(float)
            ctx = Ctx(df)
            continue

        # ---- COMPUTE / IF ---------------------------------------------
        m = COMPUTE_RE.match(stmt)
        if m:
            try:
                df[m.group("target")] = eval_cond(m.group("expr"), ctx)
                ctx = Ctx(df)
            except Exception as e:  # noqa: BLE001
                notes.append(f"COMPUTE 실패 [{stmt[:40]}] {e}")
            continue

        m = IF_RE.match(stmt)
        if m:
            tgt = m.group("target")
            try:
                cond = (eval_cond(m.group("cond"), ctx) == 1.0).fillna(False)
                key = ctx.lookup.get(tgt.lower(), tgt)
                if key not in df.columns:
                    df[key] = np.nan
                df.loc[cond, key] = float(m.group("val"))
                ctx = Ctx(df)
            except Exception as e:  # noqa: BLE001
                notes.append(f"IF 실패 [{stmt[:40]}] {e}")
            continue

        notes.append(f"처리하지 않은 명령: {stmt[:60]}")

    return checks, notes, df

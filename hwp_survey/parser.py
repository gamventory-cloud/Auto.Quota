# -*- coding: utf-8 -*-
"""추출 결과 <-> 중간 설문 문법(DSL) <-> 렌더링 블록.

중간 문법을 한 단계 끼워 넣은 이유: 한글 설문지는 서식이 제각각이라
자동 인식이 항상 맞지는 않는다. 사람이 텍스트 한 판을 눈으로 훑고
고치는 것이 docx를 직접 손보는 것보다 훨씬 빠르다.

    # 제목
    > 안내문
    ~ 박스 안내문(테두리 상자)
    ## 섹션 제목
    ! 지시문
    1. 문항 [단일|복수|단답|장문|척도:1-5|표:보기1,보기2,...]
    - 보기 (표 유형이면 표의 행)
    -- 표 안의 소제목 행
"""

from __future__ import annotations

import re

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳➀➁➂➃➄➅➆➇➈➉"
ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ"

# '1.' '문1.' 뿐 아니라 리서치 표기인 'SQ1.' 'Q1.' 'DQ1.'도 문항으로 본다
RE_Q = re.compile(r"^\s*(?:문\s*|SQ\s*|DQ\s*|Q\s*)?(\d{1,2})\s*[.)]\s*(.+)$",
                  re.IGNORECASE)
#: '문1.' 'SQ1.' 처럼 접두어가 붙은 문항. 이 표기를 쓰는 설문지에서는
#: 맨 앞이 숫자인 줄('1. 읽었다')은 문항이 아니라 보기다.
RE_Q_PREFIXED = re.compile(r"^\s*((?:문|SQ|DQ|Q)\s*\d{1,3})\s*[.)]\s*(.+)$",
                           re.IGNORECASE)
#: 한 줄에 코드 보기가 여러 개 있는 경우: '1. 서울 2. 부산 3. 대구'
RE_CODE_SPLIT = re.compile(r"(?:(?<=^)|(?<=\s))(\d{1,4})\.\s*([^\d]|\d(?!\s*\.))+?(?=\s+\d{1,4}\.\s|$)")
#: 문항 끝에 붙은 응답 방식 표기
RE_RESP_TAG = re.compile(r"\[\s*(복수\s*응답|중복\s*응답|단일\s*응답|1\s*개\s*선택|모두\s*선택)\s*\]\s*$")
#: 보기 번호는 9997(기타), 9998(모름) 같은 코드까지 쓰인다
RE_OPT_CODE = re.compile(r"^\s*(\d{1,4})\s*[.)]\s*(.+)$")
RE_OPT_SPLIT = re.compile(rf"[{CIRCLED}]\s*[^{CIRCLED}]*")
RE_LEAD_MARK = re.compile(
    rf"^\s*(?:[{CIRCLED}]|\(\s*\d+\s*\)|\d{{1,4}}\s*[.)]|[-•·▪])\s*")
RE_TYPE_TAG = re.compile(r"\[([^\[\]]+)\]\s*$")
RE_ROMAN_HEAD = re.compile(rf"^\s*(?:[{ROMAN}]|[IVX]{{1,4}})\s*[.)]?\s*$")
RE_SECTION_LINE = re.compile(
    rf"^\s*(?:[{ROMAN}]|[IVX]{{1,4}}|[가-힣])\s*[.)]\s*\S")

#: '□ 예  □ 아니오(설문 종료)' 처럼 네모 기호로 나열한 보기
RE_BOX_SPLIT = re.compile(r"[□☐▢]\s*[^□☐▢]+")
#: '※ 1=전혀 그렇지 않다, 3=보통이다, 5=매우 그렇다' 형태의 척도 안내
RE_SCALE_NOTE = re.compile(r"(\d)\s*=\s*([^,;/]+)")

RE_FOOTNOTE = re.compile(r"^\s*\*{1,3}\s*\S")          # * ** *** 각주
RE_LEADIN = re.compile(r"^\s*[♣◈▣▶]\s*\S")             # ♣ 다음으로 ...
RE_FIELDWORK = re.compile(r"^\s*▷?\s*조사원\s*[:：]\s*(.+)$")  # ▷ 조사원: ...

MULTI_HINTS = ("모두", "복수")
OPEN_HINTS = ("자유롭게", "서술", "의견을", "적어 주", "기술해")
SCALE_HINTS = ("전혀", "매우", "그렇지", "만족", "동의", "보통")


# =====================================================================
# 1) 추출 결과 -> DSL 텍스트
# =====================================================================
def detect_label_style(items) -> str:
    """'문1.'/'SQ1.' 표기를 쓰는 설문지인지 판별.

    이 표기를 쓰면 '1. 읽었다' 같은 줄은 보기이고, 쓰지 않으면 문항이다.
    한 문서 안에서 둘을 섞어 쓰는 경우는 드물어 문서 단위로 정한다.
    """
    prefixed = sum(1 for k, v in items if k == "p" and RE_Q_PREFIXED.match(v))
    for k, v in items:
        if k == "table":
            prefixed += sum(1 for r in v for c in r if RE_Q_PREFIXED.match(c))
    return "prefixed" if prefixed >= 3 else "bare"


def items_to_dsl(items) -> str:
    style = detect_label_style(items)
    lines: list[str] = []
    pending_scale: list[str] | None = None
    title_done = False

    def push_question(num, body, options=None, qtype=None):
        tag = qtype or guess_type(body, options)
        lines.append(f"{num}. {body} [{tag}]")
        for opt in options or []:
            lines.append(f"- {opt}")

    for kind, payload in items:
        if kind == "p":
            text = payload
            if not title_done and is_title_like(text):
                lines.append(f"# {text}")
                title_done = True
                continue
            lines.extend(classify_paragraph(text, style))
            continue

        rows = payload
        flat = " ".join(c for r in rows for c in r)

        if is_banner(rows):                          # '신문 이용' 같은 영역 배너
            lines.append(f"## {rows[0][0].strip()}")
            continue

        screening = screening_rows(rows)
        if screening:
            for label, body in screening:
                lines.extend(classify_paragraph(f"{label} {body}", style))
            continue

        # (a) 척도 안내 표: ◀ ① ② ③ ④ ⑤ ▶ / 전혀 그렇지 않다 ~ 매우 그렇다
        cols = scale_columns(rows)
        if cols:
            pending_scale = cols
            continue

        # (b) 매트릭스(리커트) 표
        matrix = matrix_rows(rows)
        head = pending_scale
        if matrix is None:
            fallback = header_matrix(rows)       # 칸이 빈 표
            if fallback:
                head, matrix = fallback
        if matrix:
            head = head or ["①", "②", "③", "④", "⑤"]
            label, stem = pop_matrix_stem(lines)
            stem = stem or "다음 각 항목에 대해 응답해 주십시오."
            prefix = f"{label}. " if label else ""
            lines.append(f"{prefix}{stem} [표:{','.join(head)}]")
            lines.extend(matrix)
            lines.append("")
            continue

        # (c) 섹션 머리표: Ⅰ | 다음은 ... 항목입니다.
        if len(rows) == 1 and len(rows[0]) >= 2 and RE_ROMAN_HEAD.match(rows[0][0]):
            label, desc = rows[0][0].strip(), " ".join(rows[0][1:]).strip()
            head, _, rest = desc.partition(".")
            lines.append(f"## {label}. {head.strip()}")
            if rest.strip():
                lines.append(f"! {rest.strip()}")
            continue

        # (d) 동의 여부 표
        if "동의" in flat and "□" in flat:
            for r in rows:
                cells = [c for c in r if c]
                marks = [c for c in cells if c.startswith("□")]
                if marks:
                    push_question(len(consent_nums(lines)) + 900, "동의 여부",
                                  [m.lstrip("□ ").strip() for m in marks], "단일")
                    lines[-len(marks) - 1] = "동의 여부를 표시해 주십시오. [단일]"
                else:
                    for c in cells:
                        lines.append(f"~ {c}")
            continue

        if len(rows) >= 2:
            lines.extend(grid_lines(rows))           # 빈도·기입 표는 격자 그대로
            continue

        # (e) 그 밖의 상자(제목/인사말/용어 정의)
        for r in rows:
            for c in r:
                if not c:
                    continue
                if not title_done and is_title_like(c):
                    lines.append(f"# {c}")
                    title_done = True
                elif any(k in c for k in ("안녕하십니", "감사", "협조")):
                    lines.append(f"> {c}")
                else:
                    lines.append(f"~ {c}")

    return "\n".join(collapse_blanks(retype_questions(lines)))


def retype_questions(lines):
    """문항 다음 줄에 보기가 붙어 있으면 [단답] -> [단일]/[복수]로 바로잡는다."""
    for i, line in enumerate(lines):
        m = RE_Q.match(line)
        if not m or not line.rstrip().endswith("[단답]"):
            continue
        nxt = next((l for l in lines[i + 1:]
                    if l.strip() and not l.startswith("!")), "")   # 각주는 건너뛴다
        if nxt.startswith("-"):
            body = m.group(2)
            tag = "복수" if any(k in body for k in MULTI_HINTS) else "단일"
            lines[i] = re.sub(r"\[단답\]$", f"[{tag}]", line.rstrip())
    return lines


def consent_nums(lines):
    return [l for l in lines if l.startswith("동의")]


def is_title_like(text: str) -> bool:
    if len(text) > 70 or RE_Q.match(text):
        return False
    return ("설 문 지" in text or "설문지" in text or "조사" in text
            or "영향" in text) and "안녕" not in text


def classify_paragraph(text: str, style: str = "bare") -> list[str]:
    """본문 문단 한 줄을 DSL 한 줄 이상으로."""
    out: list[str] = []

    field = RE_FIELDWORK.match(text)
    if field:
        return [f"! 조사원: {field.group(1).strip()}"]
    if RE_FOOTNOTE.match(text) or RE_LEADIN.match(text):
        return [f"! {text.strip()}"]
    m = (RE_Q_PREFIXED if style == "prefixed" else RE_Q).match(text)
    if m:
        num, body = m.group(1), m.group(2).strip()
        if style == "prefixed":
            num = re.sub(r"\s+", "", num)          # '문 5' -> '문5'
        resp = RE_RESP_TAG.search(body)            # '[복수 응답]' 같은 꼬리표
        forced = None
        if resp:
            body = body[: resp.start()].strip()
            forced = "복수" if "복수" in resp.group(1) or "중복" in resp.group(1) \
                or "모두" in resp.group(1) else "단일"

        inline = [o.strip() for o in RE_OPT_SPLIT.findall(body)]
        codes = code_options(body) if style == "prefixed" else []
        if len(inline) >= 2:
            head = body[: body.index(inline[0][0])].strip()
            out.append(f"{num}. {head} [{forced or guess_type(head, inline)}]")
            out += [f"- {strip_mark(o)}" for o in inline if strip_mark(o)]
        elif len(codes) >= 2:                      # '1. 서울 2. 부산 …'
            head = body[: body.index(codes[0])].strip()
            out.append(f"{num}. {head} [{forced or guess_type(head, codes)}]")
            out += [f"- {c}" for c in codes]
        else:
            out.append(f"{num}. {body} [{guess_type(body, None)}]")
        return out

    if RE_SECTION_LINE.match(text) and len(text) <= 40 and not RE_LEAD_MARK.match(text):
        return [f"## {text.strip()}"]

    inline = [o.strip() for o in RE_OPT_SPLIT.findall(text)]
    if inline:                                  # 보기만 있는 줄
        return [f"- {strip_mark(o)}" for o in inline if strip_mark(o)]
    code = RE_OPT_CODE.match(text)
    if code and style == "prefixed":            # 응답 코드는 그대로 살린다
        return [f"- {code.group(1)}. {code.group(2).strip()}"]
    if RE_LEAD_MARK.match(text):
        return [f"- {RE_LEAD_MARK.sub('', text).strip()}"]
    if any(k in text for k in ("안녕하십니", "감사합니", "협조")):
        return [f"> {text}"]
    return [f"! {text}"]


def code_options(text: str) -> list[str]:
    """'1. 서울 2. 부산 3. 대구' -> ['1. 서울', '2. 부산', '3. 대구']"""
    marks = [m.start() for m in re.finditer(r"(?:(?<=^)|(?<=\s))\d{1,4}\.\s", text)]
    if len(marks) < 2:
        return []
    bounds = marks + [len(text)]
    return [text[bounds[i]:bounds[i + 1]].strip() for i in range(len(marks))]


def strip_mark(text: str) -> str:
    return RE_LEAD_MARK.sub("", text).strip()


def guess_type(body: str, options) -> str:
    if any(k in body for k in MULTI_HINTS):
        return "복수"
    if not options and any(k in body for k in OPEN_HINTS):
        return "장문"
    if not options:
        return "단답"
    return "단일"


def scale_columns(rows) -> list[str] | None:
    """척도 안내 표를 열 라벨로. 예: ['① 전혀 그렇지 않다','②','③','④','⑤ 매우 그렇다']"""
    if len(rows) > 3:
        return None
    flat = " ".join(c for r in rows for c in r)
    marks = [c for c in flat if c in CIRCLED]
    if len(marks) < 3 or not any(k in flat for k in SCALE_HINTS):
        return None
    labels = [c.strip() for r in rows for c in r
              if c.strip() and not any(ch in c for ch in CIRCLED + "◀▶")]
    cols = list(dict.fromkeys(marks))
    if labels:
        cols[0] = f"{cols[0]} {labels[0]}"
        cols[-1] = f"{cols[-1]} {labels[-1]}"
    return cols


def matrix_rows(rows) -> list[str] | None:
    """리커트 표 -> ['-- 소제목', '- 1. 문항', ...]. 매트릭스가 아니면 None."""
    scored = [r for r in rows
              if len(r) >= 3 and sum(1 for c in r if c.strip() in
                                     [ch for ch in CIRCLED]) >= 3]
    if len(scored) < 2:
        return None

    out: list[str] = []
    for r in rows:
        cells = [c.strip() for c in r]
        marks = sum(1 for c in cells if c in [ch for ch in CIRCLED])
        texts = [c for c in cells if c and c not in [ch for ch in CIRCLED]]
        if marks >= 3 and texts:
            num = texts[0] if texts[0].isdigit() else None
            body = " ".join(texts[1:]) if num else " ".join(texts)
            body = re.sub(r"\s+([,.?!)])", r"\1", body)
            out.append(f"- {num}. {body}" if num else f"- {body}")
        elif texts:                              # 소제목 행('자율성'에 관한 문항)
            out.append(f"-- {' '.join(texts)}")
    return out


def header_matrix(rows) -> tuple[list[str], list[str]] | None:
    """칸이 비어 있는 표: 첫 행이 척도 라벨이면 매트릭스로 본다."""
    if len(rows) < 2:
        return None
    head = [c.strip() for c in rows[0]]
    if len(head) < 3:
        return None
    joined = " ".join(head)
    scale_like = (sum(k in joined for k in SCALE_HINTS) >= 2
                  or sum(bool(re.fullmatch(r"\d", c)) for c in head) >= 3)
    if not scale_like:
        return None
    labels = [r[0].strip() for r in rows[1:] if r and r[0].strip()]
    if len(labels) < 2:
        return None
    return [c for c in head[1:] if c], [f"- {l}" for l in labels]


def scale_from_note(text: str, default=None) -> list[str] | None:
    """'※ 1=전혀 그렇지 않다, 3=보통이다, 5=매우 그렇다' -> 5칸 라벨.

    비어 있는 자리(2, 4번)는 표준 5점 라벨로 채운다.
    """
    if "=" not in text:
        return None
    pairs = [(int(n), lab.strip()) for n, lab in RE_SCALE_NOTE.findall(text)]
    pairs = [(n, lab) for n, lab in pairs if lab]
    if len(pairs) < 2:
        return None
    top = max(n for n, _ in pairs)
    if not 3 <= top <= 10:
        return None
    base = list(default or [])
    if len(base) != top:
        base = [""] * top
    out = base[:]
    for n, lab in pairs:
        out[n - 1] = lab
    return out


def is_banner(rows) -> bool:
    """한 칸짜리 짧은 표 = 영역 구분 배너('신문 이용')."""
    return (len(rows) == 1 and len(rows[0]) == 1
            and 0 < len(rows[0][0].strip()) <= 30)


def screening_rows(rows) -> list[tuple[str, str]] | None:
    """'SQ1. 거주 | 1. 서울 2. 부산 …' 형태의 스크리닝 표."""
    out = []
    for r in rows:
        if len(r) < 2 or not r[0].strip() or not r[1].strip():
            return None
        if not re.match(r"^\s*(?:SQ|Q|문)\s*\d", r[0], re.I):
            return None
        out.append((r[0].strip(), " ".join(c for c in r[1:] if c).strip()))
    return out or None


def grid_lines(rows) -> list[str]:
    """분류되지 않은 표는 격자 그대로 옮긴다(빈도 표, 기입 표 등)."""
    return [f"@표: {','.join(c.replace(',', ' ') for c in r)}" for r in rows]


def pop_matrix_stem(lines) -> tuple[str | None, str | None]:
    """표 바로 앞의 지시문/문항을 매트릭스의 (번호, 문항 문장)으로 끌어올린다."""
    for i in range(len(lines) - 1, -1, -1):
        s = lines[i].strip()
        if not s:
            continue
        if s.startswith("!") and len(s) > 6:
            return None, lines.pop(i).lstrip("! ").strip()
        m = RE_Q_PREFIXED.match(s)
        if m:                                        # '문40. …' 번호를 지킨다
            lines.pop(i)
            return re.sub(r"\s+", "", m.group(1)), \
                RE_TYPE_TAG.sub("", m.group(2)).strip()
        if RE_Q.match(s):
            body = RE_TYPE_TAG.sub("", lines.pop(i)).strip()
            return None, RE_Q.match(body).group(2).strip()
        return None, None
    return None, None


def collapse_blanks(lines):
    out = []
    for line in lines:
        if not line.strip() and (not out or not out[-1].strip()):
            continue
        out.append(line)
    return out


# =====================================================================
# 2) DSL 텍스트 -> 렌더링 블록
# =====================================================================
def parse_dsl(text: str) -> list[dict]:
    blocks: list[dict] = []
    cur = None

    for raw in text.splitlines():
        s = raw.strip()
        if not s:
            continue

        if s.startswith("@표:"):
            row = [c.strip() for c in s.split(":", 1)[1].split(",")]
            if blocks and blocks[-1]["kind"] == "grid":
                blocks[-1]["rows"].append(row)
            else:
                cur = None
                blocks.append({"kind": "grid", "rows": [row]})
            continue

        if s.startswith("##"):
            cur = None
            blocks.append({"kind": "section", "text": s.lstrip("#").strip()})
        elif s.startswith("#"):
            cur = None
            blocks.append({"kind": "title", "text": s.lstrip("#").strip()})
        elif s.startswith(">"):
            cur = None
            blocks.append({"kind": "intro", "text": s.lstrip("> ").strip()})
        elif s.startswith("~"):
            cur = None
            blocks.append({"kind": "box", "text": s.lstrip("~ ").strip()})
        elif s.startswith("--"):
            if cur:
                cur["options"].append({"type": "group", "text": s.lstrip("- ").strip()})
        elif s.startswith("-"):
            if cur:
                cur["options"].append({"type": "row", "text": s.lstrip("- ").strip()})
            else:
                blocks.append({"kind": "note", "text": s.lstrip("- ").strip()})
        elif s.startswith("!"):
            note = s.lstrip("! ").strip()
            if cur:
                cur["notes"].append(note)
            else:
                blocks.append({"kind": "note", "text": note})
        else:
            label = None
            m_label = RE_Q_PREFIXED.match(s)
            if m_label:
                label = re.sub(r"\s+", "", m_label.group(1))
            m = RE_Q.match(s)
            body = (m_label or m).group(2) if (m_label or m) else s
            qtype, scale, matrix = "단일", None, None
            tag = RE_TYPE_TAG.search(body)
            if tag:
                name = tag.group(1).strip()
                body = body[: tag.start()].strip()
                if name.startswith("척도"):
                    qtype = "척도"
                    nums = re.findall(r"\d+", name)
                    scale = (int(nums[0]), int(nums[1])) if len(nums) >= 2 else (1, 5)
                elif name.startswith("표"):
                    qtype = "표"
                    matrix = [x.strip() for x in name.split(":", 1)[1].split(",")]
                else:
                    qtype = name
            cur = {"kind": "question", "text": body.strip(), "type": qtype,
                   "label": label, "scale": scale, "matrix": matrix,
                   "options": [], "notes": []}
            blocks.append(cur)

    return blocks


def _finalize(blocks):
    for b in blocks:
        if b["kind"] != "question":
            continue
        if b["type"] == "단답" and b["options"]:
            b["type"] = "복수" if any(k in b["text"] for k in MULTI_HINTS) else "단일"
    return _finalize(blocks)


def summarize(blocks) -> dict:
    q = [b for b in blocks if b["kind"] == "question"]
    return {
        "문항": len(q),
        "매트릭스 표": sum(1 for b in q if b["type"] == "표"),
        "매트릭스 세부항목": sum(len([o for o in b["options"] if o["type"] == "row"])
                          for b in q if b["type"] == "표"),
        "섹션": sum(1 for b in blocks if b["kind"] == "section"),
        "일반 표": sum(1 for b in blocks if b["kind"] == "grid"),
    }

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

RE_Q = re.compile(r"^\s*(?:문\s*)?(\d{1,2})\s*[.)]\s*(.+)$")
RE_OPT_SPLIT = re.compile(rf"[{CIRCLED}]\s*[^{CIRCLED}]*")
RE_LEAD_MARK = re.compile(rf"^\s*(?:[{CIRCLED}]|\(\s*\d+\s*\)|\d\s*\)|[-•·▪])\s*")
RE_TYPE_TAG = re.compile(r"\[([^\[\]]+)\]\s*$")
RE_ROMAN_HEAD = re.compile(rf"^\s*(?:[{ROMAN}]|[IVX]{{1,4}})\s*[.)]?\s*$")
RE_SECTION_LINE = re.compile(
    rf"^\s*(?:[{ROMAN}]|[IVX]{{1,4}}|[가-힣])\s*[.)]\s*\S")

MULTI_HINTS = ("모두", "복수")
OPEN_HINTS = ("자유롭게", "서술", "의견을", "적어 주", "기술해")
SCALE_HINTS = ("전혀", "매우", "그렇지", "만족", "동의", "보통")


# =====================================================================
# 1) 추출 결과 -> DSL 텍스트
# =====================================================================
def items_to_dsl(items) -> str:
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
            lines.extend(classify_paragraph(text))
            continue

        rows = payload
        flat = " ".join(c for r in rows for c in r)

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
            stem = pop_matrix_stem(lines) or "다음 각 항목에 대해 응답해 주십시오."
            lines.append(f"{stem} [표:{','.join(head)}]")
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
        nxt = next((l for l in lines[i + 1:] if l.strip()), "")
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


def classify_paragraph(text: str) -> list[str]:
    """본문 문단 한 줄을 DSL 한 줄 이상으로."""
    out: list[str] = []
    m = RE_Q.match(text)
    if m:
        num, body = m.group(1), m.group(2).strip()
        inline = [o.strip() for o in RE_OPT_SPLIT.findall(body)]
        if len(inline) >= 2:
            head = body[: body.index(inline[0][0])].strip()
            out.append(f"{num}. {head} [{guess_type(head, inline)}]")
            out += [f"- {strip_mark(o)}" for o in inline if strip_mark(o)]
        else:
            out.append(f"{num}. {body} [{guess_type(body, None)}]")
        return out

    if RE_SECTION_LINE.match(text) and len(text) <= 40 and not RE_LEAD_MARK.match(text):
        return [f"## {text.strip()}"]

    inline = [o.strip() for o in RE_OPT_SPLIT.findall(text)]
    if inline:                                  # 보기만 있는 줄
        return [f"- {strip_mark(o)}" for o in inline if strip_mark(o)]
    if RE_LEAD_MARK.match(text):
        return [f"- {RE_LEAD_MARK.sub('', text).strip()}"]
    if any(k in text for k in ("안녕하십니", "감사합니", "협조")):
        return [f"> {text}"]
    return [f"! {text}"]


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
    if len(head) < 3 or not head[0]:
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


def pop_matrix_stem(lines) -> str | None:
    """표 바로 앞의 지시문/문항을 매트릭스의 문항 문장으로 끌어올린다."""
    for i in range(len(lines) - 1, -1, -1):
        s = lines[i].strip()
        if not s:
            continue
        if s.startswith("!") and len(s) > 6:
            return lines.pop(i).lstrip("! ").strip()
        if RE_Q.match(s):
            body = RE_TYPE_TAG.sub("", lines.pop(i)).strip()
            return RE_Q.match(body).group(2).strip()
        return None
    return None


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
            m = RE_Q.match(s)
            body = m.group(2) if m else s
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
                   "scale": scale, "matrix": matrix, "options": [], "notes": []}
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
    }

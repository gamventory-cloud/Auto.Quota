#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
한글(.hwp / .hwpx) 설문지 -> 워드(.docx) 설문지 변환

사용법:
    python hwp_survey_to_docx.py 설문지.hwp  결과.docx
    python hwp_survey_to_docx.py 설문지.hwpx 결과.docx
    python hwp_survey_to_docx.py 설문지.hwp  결과.docx --dump 중간확인.txt

준비:
    pip install python-docx pyhwp lxml beautifulsoup4
    (.hwpx만 다룰 경우 pyhwp는 필요 없음)

동작 순서
    1) .hwpx  -> ZIP 안의 Contents/section*.xml 을 직접 파싱 (표 구조 보존)
       .hwp   -> pyhwp의 hwp5html 로 변환 후 XHTML 파싱 (표 구조 보존)
                 실패 시 hwp5txt(순수 텍스트)로 자동 폴백
    2) 추출된 문단/표를 설문 문법(#, ##, 문항, -, [유형])으로 자동 변환
    3) survey_to_docx.py 의 렌더러로 .docx 생성

주의: 자동 인식은 100%가 아니므로 --dump 로 중간 텍스트를 뽑아
      직접 손본 뒤 `python survey_to_docx.py 중간확인.txt 결과.docx` 로
      다시 굽는 방식을 권장합니다.
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from survey_to_docx import parse, render   # 같은 폴더에 두세요

# ---------------------------------------------------------------------------
# 1단계: 원문 추출  -> [("p", "문단텍스트"), ("table", [[셀,셀],[셀,셀]]), ...]
# ---------------------------------------------------------------------------
def local(tag):
    return tag.rsplit("}", 1)[-1]


def read_hwpx(path):
    """.hwpx = ZIP + OWPML(XML). 네임스페이스가 버전마다 달라 localname으로 처리."""
    items = []
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist()
                       if re.match(r"Contents/section\d+\.xml$", n))
        if not names:  # 드물게 경로가 다른 경우
            names = sorted(n for n in z.namelist() if n.endswith(".xml")
                           and "section" in n.lower())
        for name in names:
            root = ET.fromstring(z.read(name))
            items.extend(walk_owpml(root))
    return items


def text_of(el):
    """하위 <t> 요소를 모두 이어붙여 문단 텍스트를 만든다."""
    buf = []
    for node in el.iter():
        if local(node.tag) == "t" and node.text:
            buf.append(node.text)
        if local(node.tag) in ("lineBreak", "tab"):
            buf.append(" ")
    return re.sub(r"\s+", " ", "".join(buf)).strip()


def walk_owpml(el, items=None):
    if items is None:
        items = []
    for child in el:
        name = local(child.tag)
        if name == "tbl":
            rows = []
            for tr in child.iter():
                if local(tr.tag) != "tr":
                    continue
                cells = [text_of(tc) for tc in tr if local(tc.tag) == "tc"]
                if any(cells):
                    rows.append(cells)
            if rows:
                items.append(("table", rows))
            continue                      # 표 안쪽은 위에서 이미 처리
        if name == "p":
            has_table = any(local(n.tag) == "tbl" for n in child.iter())
            if has_table:                 # 표를 품은 문단이면 안쪽을 계속 탐색
                walk_owpml(child, items)
                continue
            t = text_of(child)
            if t:
                items.append(("p", t))
            continue
        walk_owpml(child, items)
    return items


def read_hwp(path):
    """.hwp(HWP 5.x): pyhwp hwp5html -> XHTML 파싱. 실패 시 hwp5txt."""
    from bs4 import BeautifulSoup

    tmp = tempfile.mkdtemp(prefix="hwp5_")
    try:
        subprocess.run(["hwp5html", "--output", tmp, path],
                       check=True, capture_output=True)
        html_path = next((os.path.join(tmp, f) for f in os.listdir(tmp)
                          if f.endswith((".xhtml", ".html"))), None)
        if not html_path:
            raise RuntimeError("hwp5html 출력물을 찾을 수 없음")
        with open(html_path, encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "lxml")
        body = soup.body or soup
        items = []
        for el in body.find_all(["p", "div", "table"], recursive=True):
            if el.name == "table":
                rows = []
                for tr in el.find_all("tr"):
                    cells = [re.sub(r"\s+", " ", td.get_text(" ")).strip()
                             for td in tr.find_all(["td", "th"])]
                    if any(cells):
                        rows.append(cells)
                if rows:
                    items.append(("table", rows))
            elif not el.find(["p", "div", "table"]):     # 최하위 블록만
                t = re.sub(r"\s+", " ", el.get_text(" ")).strip()
                if t:
                    items.append(("p", t))
        return items
    except (subprocess.CalledProcessError, FileNotFoundError, RuntimeError) as e:
        print(f"[알림] hwp5html 실패({e}) → hwp5txt 텍스트 모드로 진행합니다.")
        out = subprocess.run(["hwp5txt", path], check=True,
                             capture_output=True, text=True).stdout
        return [("p", line.strip()) for line in out.splitlines() if line.strip()]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------
# 2단계: 설문 문법으로 자동 변환
# ---------------------------------------------------------------------------
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

RE_SECTION = re.compile(r"^\s*(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|[IVX]{1,4}|[가-힣]\s*[.)])[\s.)]+\S")
RE_QUESTION = re.compile(r"^\s*(?:문\s*)?(\d{1,2})\s*[.)]\s*(.+)$")
RE_OPTION = re.compile(rf"^\s*(?:[{CIRCLED}]|\(\s*\d+\s*\)|\d\s*\)|[-•·▪])\s*(.+)$")
RE_INLINE_OPTS = re.compile(rf"[{CIRCLED}]\s*[^{CIRCLED}]+")
SCALE_WORDS = ("전혀", "매우", "그렇지", "만족", "불만족", "보통", "동의")


def looks_like_scale_header(cells):
    joined = " ".join(cells)
    return sum(w in joined for w in SCALE_WORDS) >= 2 or \
        sum(bool(re.fullmatch(r"\d", c.strip())) for c in cells) >= 3


def to_survey_text(items):
    """추출 결과 -> survey_to_docx가 읽는 텍스트 문법."""
    lines, title_done, qcount = [], False, 0

    for kind, payload in items:
        if kind == "table":
            rows = payload
            head = rows[0]
            if len(rows) >= 2 and len(head) >= 3 and looks_like_scale_header(head):
                # 행렬형(매트릭스) 문항으로 변환
                cols = [c for c in head[1:] if c]
                tag = f"[표:{','.join(cols)}]"
                # 바로 위 줄이 보기 없는 문항이면 그 문항을 표 유형으로 승격
                prev = next((i for i in range(len(lines) - 1, -1, -1)
                             if lines[i].strip()), None)
                if prev is not None and RE_QUESTION.match(lines[prev]) \
                        and lines[prev].rstrip().endswith(("[단일]", "[복수]")):
                    lines[prev] = re.sub(r"\[(단일|복수)\]$", tag,
                                         lines[prev].rstrip())
                else:
                    lines.append(f"다음 각 항목에 응답해 주십시오. {tag}")
                for r in rows[1:]:
                    if r and r[0]:
                        lines.append(f"- {r[0]}")
                lines.append("")
            else:
                # 표에 문항이 들어 있는 흔한 형태 → 셀 텍스트를 순서대로 흘려보냄
                for r in rows:
                    for c in r:
                        if c:
                            lines.extend(classify_line(c))
            continue

        text = payload
        if not title_done and len(text) <= 60 and not RE_QUESTION.match(text):
            lines.append(f"# {text}")
            title_done = True
            continue
        lines.extend(classify_line(text))

    # 문항 하나도 못 찾은 경우 알림용 카운트
    qcount = sum(1 for l in lines if RE_QUESTION.match(l))
    if qcount == 0:
        print("[경고] 문항 번호를 인식하지 못했습니다. --dump 결과를 직접 손보세요.")
    return "\n".join(lines)


def classify_line(text):
    """한 줄을 설문 문법 한 줄 이상으로."""
    out = []
    if RE_OPTION.match(text) and not RE_QUESTION.match(text):
        out.append(f"- {RE_OPTION.match(text).group(1).strip()}")
        return out

    m = RE_QUESTION.match(text)
    if m:
        body = m.group(2).strip()
        inline = RE_INLINE_OPTS.findall(body)      # "1. 성별? ① 남 ② 여" 처리
        if len(inline) >= 2:
            head = body[: body.index(inline[0])].strip()
            qtype = "복수" if ("모두" in head or "복수" in head) else "단일"
            out.append(f"{m.group(1)}. {head} [{qtype}]")
            out += [f"- {o.lstrip(CIRCLED).strip()}" for o in inline]
        else:
            qtype = "단일"
            if "모두" in body or "복수" in body:
                qtype = "복수"
            elif any(k in body for k in ("자유롭게", "의견", "서술", "적어")):
                qtype = "장문"
            out.append(f"{m.group(1)}. {body} [{qtype}]")
        return out

    if RE_SECTION.match(text) and len(text) <= 40:
        out.append(f"## {text.strip()}")
        return out

    if any(k in text for k in ("안녕", "감사", "협조", "목적", "소요")):
        out.append(f"> {text}")
        return out

    out.append(f"! {text}")   # 나머지는 지시문으로
    return out


# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    dump = sys.argv[sys.argv.index("--dump") + 1] if "--dump" in sys.argv else None

    ext = os.path.splitext(src)[1].lower()
    if ext == ".hwpx":
        items = read_hwpx(src)
    elif ext == ".hwp":
        items = read_hwp(src)
    else:
        print("지원 형식: .hwp, .hwpx")
        sys.exit(1)
    print(f"추출: 문단 {sum(1 for k,_ in items if k=='p')}개, "
          f"표 {sum(1 for k,_ in items if k=='table')}개")

    survey_text = to_survey_text(items)
    if dump:
        with open(dump, "w", encoding="utf-8") as f:
            f.write(survey_text)
        print("중간 텍스트 저장:", dump)

    print("저장 완료:", render(parse(survey_text), dst))


if __name__ == "__main__":
    main()

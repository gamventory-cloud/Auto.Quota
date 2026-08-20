# -*- coding: utf-8 -*-
"""한글 파일(.hwp / .hwpx)에서 문단과 표를 뽑아낸다.

반환 형식은 (kind, payload) 튜플의 리스트:
    ("p",     "문단 텍스트")
    ("table", [["셀", "셀"], ["셀", "셀"]])

표 구조를 보존하는 것이 핵심이다. 순수 텍스트 추출(hwp5txt)은 표를
'<표>' 한 줄로 날려버리기 때문에, 리커트 척도 문항이 통째로 사라진다.
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import warnings
import xml.etree.ElementTree as ET
import zipfile

WS = re.compile(r"\s+")

# 한글 파일은 글자 모양이 바뀌는 지점마다 텍스트를 조각내서 저장한다
# (한글/영문/기호가 각각 다른 run). 조각을 공백으로 이으면
# "심리 적행복감", "200 만원", "‘ 자율성 ’" 같은 군더더기가 생기므로
# 조각은 공백 없이 붙이고, 아래 규칙으로 원문 공백만 정리한다.
_FIXES = [
    (re.compile(r"([가-힣])\s+(다\.|다$|까\?|요\.|요$|니다\.)"), r"\1\2"),
    (re.compile(r"\s+([,.?!:;)\]}’”])"), r"\1"),
    (re.compile(r"([(\[{‘“])\s+"), r"\1"),
    (re.compile(r"(\d)\s+(만원|원|년|세|점|개|명|월|일|시간|%)"), r"\1\2"),
]


# 한글에서 줄 끝을 공백으로 채워 정렬한 흔적: 한글 사이의 2칸 이상 공백은
# 단어 경계가 아니라 줄바꿈 자리다("만족  " + " 감을" -> "만족감을").
_PAD = re.compile(r"(?<=[가-힣])[ \t]{2,}(?=[가-힣])")


#: 줄바꿈 채움 공백을 붙일지 여부. 붙이면 "만족 감을"->"만족감을"으로 살아나지만
#: 저자가 단어 사이를 여러 칸으로 띄운 곳은 "때 편안함"->"때편안함"이 될 수 있다.
TIGHTEN = True


def clean(text: str) -> str:
    raw = (text or "").replace("\xa0", " ")
    out = WS.sub(" ", _PAD.sub("", raw) if TIGHTEN else raw).strip()
    for pattern, repl in _FIXES:
        out = pattern.sub(repl, out)
    return out


def read_survey(path: str, tighten: bool = True) -> list[tuple[str, object]]:
    global TIGHTEN
    TIGHTEN = tighten
    ext = os.path.splitext(path)[1].lower()
    if ext == ".hwpx":
        return read_hwpx(path)
    if ext == ".hwp":
        return read_hwp(path)
    raise ValueError("지원 형식은 .hwp 와 .hwpx 입니다.")


# ---------------------------------------------------------------- .hwpx
def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _text_of(el, keep_lines: bool = False) -> str:
    if keep_lines:                                  # 문단마다 한 줄로
        parts = []
        for node in el.iter():
            if _local(node.tag) != "p":
                continue
            text = clean("".join(t.text or "" for t in node.iter()
                                 if _local(t.tag) == "t"))
            if text:
                parts.append(text)
        if len(parts) > 1 and _is_question_block(parts):
            return "\n".join(parts)      # 문항 덩어리만 줄을 남긴다
    buf = []
    for node in el.iter():
        name = _local(node.tag)
        if name == "t" and node.text:
            buf.append(node.text)          # 조각은 공백 없이 이어붙인다
        elif name in ("lineBreak", "tab"):
            buf.append(" ")
    return clean("".join(buf))


def _walk_owpml(el, items):
    for child in el:
        name = _local(child.tag)
        if name == "tbl":
            trs = [n for n in child.iter() if _local(n.tag) == "tr"]
            single = (len(trs) == 1
                      and len([c for c in trs[0] if _local(c.tag) == "tc"]) == 1)
            rows = []
            for tr in trs:
                cells = [_text_of(tc, keep_lines=single)
                         for tc in tr if _local(tc.tag) == "tc"]
                if any(cells):
                    rows.append(cells)
            if rows:
                items.append(("table", rows))
            continue
        if name == "p":
            if any(_local(n.tag) == "tbl" for n in child.iter()):
                _walk_owpml(child, items)      # 표를 품은 문단
                continue
            text = _text_of(child)
            if text:
                items.append(("p", text))
            continue
        _walk_owpml(child, items)
    return items


def read_hwpx(path: str) -> list[tuple[str, object]]:
    """.hwpx = ZIP + OWPML(XML). 네임스페이스가 버전마다 달라 localname으로 처리."""
    items: list[tuple[str, object]] = []
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist()
                       if re.match(r"Contents/section\d+\.xml$", n))
        if not names:
            names = sorted(n for n in z.namelist()
                           if n.endswith(".xml") and "section" in n.lower())
        for name in names:
            _walk_owpml(ET.fromstring(z.read(name)), items)
    return items


# ---------------------------------------------------------------- .hwp
def read_hwp(path: str) -> list[tuple[str, object]]:
    """HWP 5.x 바이너리. pyhwp의 XHTML 변환을 인프로세스로 호출한다.

    CLI(hwp5html)를 subprocess로 부르지 않는 이유: Streamlit Cloud 같은
    환경에서 PATH에 스크립트가 없을 수 있다.
    """
    from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
    from hwp5.hwp5html import HTMLTransform
    from hwp5.xmlmodel import Hwp5File

    warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

    tmp = tempfile.mkdtemp(prefix="hwp5_")
    try:
        hwp5 = Hwp5File(path)
        try:
            HTMLTransform().transform_hwp5_to_dir(hwp5, tmp)
        finally:
            hwp5.close()

        html_path = next((os.path.join(tmp, f) for f in sorted(os.listdir(tmp))
                          if f.endswith((".xhtml", ".html"))), None)
        if html_path is None:
            raise RuntimeError("변환 결과에서 XHTML을 찾지 못했습니다.")
        with open(html_path, encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "lxml")
        items = _items_from_html(soup)
        return _recover_dropped(path, items)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


#: 한 칸짜리 표 안에 문항·보기가 통째로 들어 있는지 판별할 때 쓰는 표시들
_MARKER = re.compile(
    r"^\s*(?:[①-⑳➀-➉□☐▢]|(?:문|SQ|DQ|Q)?\s*\d{1,2}(?:-\d{1,2})?\s*[.)】])")


def _is_question_block(parts) -> bool:
    """문항이나 보기로 보이는 줄이 둘 이상이면 줄 구분을 살려야 한다."""
    return sum(1 for x in parts if _MARKER.match(x)) >= 2


def _block_text(el, keep_lines: bool = False) -> str:
    """셀/문단의 텍스트. 내부 <p>(줄 단위)는 공백으로, run 조각은 붙여서.

    keep_lines=True 면 문단 경계를 줄바꿈으로 남긴다. 한 칸짜리 표 안에
    문항과 보기가 통째로 들어 있는 경우가 많아, 이때는 줄 구분이 필요하다.
    """
    paras = el.find_all("p")
    if paras:
        parts = [clean(p.get_text("")) for p in paras]
    else:
        parts = [clean(el.get_text(""))]
    parts = [x for x in parts if x]
    if keep_lines and len(parts) > 1 and _is_question_block(parts):
        return "\n".join(parts)          # 문항 덩어리만 줄을 남긴다
    return clean(" ".join(parts))


def _model_paragraphs(path: str) -> list[str]:
    """HWP 이진 구조에서 문단 텍스트를 순서대로 꺼낸다.

    XHTML 변환기는 글상자 같은 개체 안의 문단을 그리지 않고 버린다.
    문항 제목이 글상자에 들어 있으면 통째로 사라지므로, 이진 구조를
    한 번 더 훑어 빠진 문단을 되찾는다.
    """
    from hwp5.treeop import STARTEVENT
    from hwp5.xmlmodel import Hwp5File

    hwp5 = Hwp5File(path)
    try:
        indexes = list(hwp5.bodytext.section_indexes())
        paras, buf = [], []
        for index in indexes:
            for event, item in hwp5.bodytext.section(index).events():
                model = item[0] if isinstance(item, (tuple, list)) else item
                attrs = (item[1] if isinstance(item, (tuple, list)) and len(item) > 1
                         and isinstance(item[1], dict) else {})
                name = getattr(model, "__name__", "")
                if event is not STARTEVENT:
                    continue
                if name == "Paragraph" and buf:
                    paras.append(clean("".join(buf)))
                    buf = []
                elif name == "Text":
                    text = attrs.get("text", "")
                    if isinstance(text, str):
                        buf.append(text)
        if buf:
            paras.append(clean("".join(buf)))
        return [p for p in paras if p]
    finally:
        hwp5.close()


#: 'SQ4-1.', 'A3-2-1.', 'B0.' 같은 문항 번호. 되찾은 문단의 제자리를 찾는 데 쓴다.
_LABEL = re.compile(r"^\s*(SQ|DQ|AQ|[A-Z]|문)\s*(\d+(?:[-_]\d+)*)\s*[.)\]】]", re.I)


def _label_key(text: str):
    """문항 번호를 정렬 가능한 값으로. 'A3-2-1' -> (1, (3, 2, 1))"""
    m = _LABEL.match(text or "")
    if not m:
        return None
    head = m.group(1).upper()
    rank = 0 if head in ("SQ", "DQ", "AQ", "문") else ord(head) - ord("A") + 1
    nums = tuple(int(n) for n in re.split(r"[-_]", m.group(2)))
    return (rank, nums)


def _recover_dropped(path: str, items: list) -> list:
    """XHTML에서 누락된 문단을 이진 구조에서 찾아 제자리에 끼워 넣는다."""
    try:
        paragraphs = _model_paragraphs(path)
    except Exception:                                 # noqa: BLE001
        return items                                  # 되찾기는 있으면 좋은 기능이다

    pool = "\n".join(v if k == "p" else " ".join(c for r in v for c in r)
                      for k, v in items)
    pool = WS.sub("", pool)

    recovered, cursor = [], 0
    for text in paragraphs:
        key = WS.sub("", text)
        if len(key) < 6:
            continue
        if key in pool:
            for n in range(cursor, len(items)):       # 위치를 따라간다
                kind, payload = items[n]
                blob = payload if kind == "p" else " ".join(
                    c for r in payload for c in r)
                if key in WS.sub("", blob):
                    cursor = n + 1
                    break
            continue
        recovered.append((_place(items, text, cursor), text))

    recovered.sort(key=lambda pair: pair[0])
    for offset, (position, text) in enumerate(recovered):
        items.insert(position + offset, ("p", text))
    return items


def _place(items, text, cursor) -> int:
    """되찾은 문단을 넣을 자리.

    기본은 이진 구조에서 읽어온 순서(cursor)를 따른다. 다만 글상자는 문서
    흐름과 다른 자리에 저장되기도 해서, 같은 계열의 더 큰 번호(A3-3)를 이미
    지나쳐 버린 경우에만 번호 차례에 맞게 앞으로 되돌린다.
    """
    key = _label_key(text)
    if key is None:
        return cursor
    for n in range(min(cursor, len(items))):
        kind, payload = items[n]
        blob = payload if kind == "p" else (payload[0][0] if payload and payload[0]
                                            else "")
        other = _label_key(blob)
        if other and other[0] == key[0] and other[1] > key[1]:
            return _before_options(items, n, key)   # 번호를 앞질렀다
    return cursor


_OPTION_HEAD = re.compile(r"^\s*[①-⑳➀-➉□☐]")
#: 보기 뒤에 붙는 지시문. 문항 자리를 찾을 때 함께 건너뛴다.
_TRAILING = re.compile(r"^\s*[\[（(]?\s*(?:PROG|DP|DATA)\s*[:：]|^\s*※")


def _before_options(items, index: int, key) -> int:
    """보기 목록 앞으로 자리를 물린다. 문항이 자기 보기 뒤에 놓이지 않도록.

    번호가 앞선 문항을 넘어가지는 않는다(A3-4 가 A3-3 앞으로 가면 안 된다).
    """
    while index > 0:
        kind, payload = items[index - 1]
        if kind == "p":
            other = _label_key(payload)
            if other and other < key:
                break                     # 앞 번호 문항까지만
        if kind == "p" and _TRAILING.match(payload):
            index -= 1
            continue
        cells = ([payload] if kind == "p"
                 else [c for r in payload for c in r if c.strip()])
        # '기타()', '해당 없음'처럼 기호 없는 보기가 섞이므로 과반이면 보기로 본다
        marked = sum(1 for c in cells if _OPTION_HEAD.match(c) or len(c) < 3)
        if cells and marked >= len(cells) - 1:
            index -= 1
            continue
        break
    return index


def _items_from_html(soup) -> list[tuple[str, object]]:
    body = soup.body or soup
    items: list[tuple[str, object]] = []

    for el in body.find_all(["p", "div", "table"]):
        if el.find_parent("table") is not None:
            continue                              # 표 내부는 표에서 처리
        if el.name == "table":
            trs = el.find_all("tr")
            single = len(trs) == 1 and len(trs[0].find_all(["td", "th"])) == 1
            rows = []
            for tr in trs:
                cells = [_block_text(td, keep_lines=single)
                         for td in tr.find_all(["td", "th"])]
                if any(cells):
                    rows.append(cells)
            if rows:
                items.append(("table", rows))
            continue
        if el.find(["p", "div", "table"]):
            continue                              # 최하위 블록만
        text = _block_text(el)
        if text:
            items.append(("p", text))
    return items

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


def _text_of(el) -> str:
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
            rows = []
            for tr in child.iter():
                if _local(tr.tag) != "tr":
                    continue
                cells = [_text_of(tc) for tc in tr if _local(tc.tag) == "tc"]
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
        return _items_from_html(soup)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _block_text(el) -> str:
    """셀/문단의 텍스트. 내부 <p>(줄 단위)는 공백으로, run 조각은 붙여서."""
    paras = el.find_all("p")
    if paras:
        parts = [clean(p.get_text("")) for p in paras]
    else:
        parts = [clean(el.get_text(""))]
    return clean(" ".join(x for x in parts if x))


def _items_from_html(soup) -> list[tuple[str, object]]:
    body = soup.body or soup
    items: list[tuple[str, object]] = []

    for el in body.find_all(["p", "div", "table"]):
        if el.find_parent("table") is not None:
            continue                              # 표 내부는 표에서 처리
        if el.name == "table":
            rows = []
            for tr in el.find_all("tr"):
                cells = [_block_text(td) for td in tr.find_all(["td", "th"])]
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

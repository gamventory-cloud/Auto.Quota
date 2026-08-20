# -*- coding: utf-8 -*-
"""변환 누락 검증: 원본(PDF) ↔ 변환 결과(워드) 대조.

왜 PDF인가
    변환기의 파서로 읽은 텍스트를 기준으로 삼으면, 파서가 애초에 놓친 것은
    양쪽 모두에 없으므로 검출되지 않는다(같은 눈으로 두 번 보는 셈).
    한글에서 "PDF로 저장"한 파일은 파서를 거치지 않은 독립적인 기준이 된다.

왜 줄 단위 비교가 아닌가
    변환은 텍스트를 재배치한다. '(□ 예, □ 아니오) 나는 …' 한 줄이 표의 열
    머리와 행으로 쪼개지고, 번호가 바뀌고, 보기가 셀로 들어간다. 줄 단위로
    대조하면 정상 변환도 절반 이상이 "불일치"로 잡힌다(실측 44~62%).
    그래서 기호·번호를 지운 뒤 문장 단위로 존재 여부만 확인한다.

판정 세 갈래
    누락      앞부분조차 결과에서 찾을 수 없음 -> 사람이 확인해야 한다
    부분 일치 앞부분은 있으나 뒤가 다름 -> 지시문 분리·문구 변형이 대부분
    확인      그대로 있음
"""

from __future__ import annotations

import re

MARKS = "□☐▢①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳➀➁➂➃➄○◯●·▪◀▶※✓√"
PAREN_MARK = re.compile(rf"\([^)]*[{MARKS}][^)]*\)")      # '(□ 예, □ 아니오)'
LABEL = re.compile(r"^\s*(?:문|배문|SQ|DQ|AQ|Q|S|A)?\s*\d{0,3}(?:[-_]\d{1,3})?\s*[.)]\s*",
                   re.I)
SCALE_NOTE = re.compile(r"\d\s*=\s*\S")                   # '1=전혀 그렇지 않다'
DROP = re.compile(r"[\s()\[\]{}·:,.\-~=＝\"'’‘“”/|]")
PREFIX_LEN = 12                                           # 앞부분 일치 판정 길이
MIN_LEN = 5                                               # 이보다 짧은 문장은 무시


# ---------------------------------------------------------------- 텍스트 읽기
def pdf_text(path: str) -> str:
    """PDF 본문 텍스트. pypdf로 읽고, 빈약하면 pdfplumber로 다시 읽는다."""
    from pypdf import PdfReader

    pages = [(p.extract_text() or "") for p in PdfReader(path).pages]
    text = "\n".join(pages)
    if len(DROP.sub("", text)) >= 200:
        return text
    try:                                                  # 표가 많은 문서 대비
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:                                     # noqa: BLE001
        return text


def docx_text(path_or_stream) -> str:
    """워드의 문단과 표 셀을 모두 모은다."""
    from docx import Document

    d = Document(path_or_stream)
    texts = [p.text for p in d.paragraphs]
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
                for inner in cell.tables:             # 중첩 표(할당 격자 등)
                    texts += [c.text for r in inner.rows for c in r.cells]
    return "\n".join(t for t in texts if t and t.strip())


def hwp_text(path: str) -> str:
    """PDF가 없을 때의 차선책. 파서가 놓친 것은 검출되지 않음에 주의."""
    from .reader import read_survey

    out = []
    for kind, payload in read_survey(path):
        if kind == "p":
            out.append(payload)
        else:
            out += [c for row in payload for c in row if c.strip()]
    return "\n".join(out)


# ---------------------------------------------------------------- 정규화
def norm(text: str) -> str:
    text = PAREN_MARK.sub(" ", text or "")
    text = re.sub(rf"[{MARKS}]", " ", text)
    return DROP.sub("", text)


def sentences(text: str) -> list[tuple[str, str]]:
    """비교 단위(원문, 정규화형) 목록. 기호·번호를 지우고 문장으로 자른다."""
    text = PAREN_MARK.sub("\n", text or "")
    text = re.sub(rf"[{MARKS}]", "\n", text)
    # 괄호 안 지시문('(설문 종료)', '(적을 것: …)')은 변환 과정에서 PROG 줄로
    # 옮겨가므로 본문과 분리해서 따로 대조한다.
    text = re.sub(r"[()（）\[\]【】]", "\n", text)
    out, seen = [], set()
    for chunk in re.split(r"[\n\t]+|(?<=[.?!])\s+", text):
        chunk = LABEL.sub("", chunk.strip())
        if SCALE_NOTE.search(chunk):        # 척도 안내문은 표 머리로 흡수된다
            continue
        key = norm(chunk)
        if len(key) >= MIN_LEN and key not in seen:
            seen.add(key)
            out.append((chunk.strip(), key))
    return out


# ---------------------------------------------------------------- 대조
def compare(source_text: str, target_text: str) -> dict:
    units = sentences(source_text)
    pool = norm(target_text)

    missing, partial = [], []
    for raw, key in units:
        if key in pool:
            continue
        if key[:PREFIX_LEN] in pool:
            partial.append(raw)
        else:
            missing.append(raw)

    return {
        "대조 문장": len(units),
        "확인": len(units) - len(missing) - len(partial),
        "부분 일치": len(partial),
        "누락": len(missing),
        "누락 목록": missing,
        "부분 일치 목록": partial,
        "누락률": (len(missing) / len(units)) if units else 0.0,
    }


def compare_files(source: str, target) -> dict:
    """source 는 .pdf / .hwp / .hwpx, target 은 .docx 경로나 스트림."""
    reader = pdf_text if str(source).lower().endswith(".pdf") else hwp_text
    return compare(reader(source), docx_text(target))


def main():
    import argparse

    ap = argparse.ArgumentParser(description="변환 누락 검증 (원본 ↔ 워드)")
    ap.add_argument("source", help="원본 .pdf (없으면 .hwp/.hwpx)")
    ap.add_argument("target", help="변환 결과 .docx")
    ap.add_argument("--all", action="store_true", help="부분 일치까지 모두 출력")
    args = ap.parse_args()

    result = compare_files(args.source, args.target)
    print(f"대조 문장 {result['대조 문장']}개 · 확인 {result['확인']}"
          f" · 부분 일치 {result['부분 일치']} · 누락 {result['누락']}"
          f" ({result['누락률']:.1%})")
    for text in result["누락 목록"]:
        print("  [누락]", text[:90])
    if args.all:
        for text in result["부분 일치 목록"]:
            print("  [부분]", text[:90])


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""명령줄 사용: python -m hwp_survey.cli 설문지.hwp 결과.docx [--dsl 중간.txt]"""
import argparse

from . import read_survey, items_to_dsl, parse_dsl, summarize, SurveyWriter


def main():
    ap = argparse.ArgumentParser(description="한글 설문지 -> 워드 설문지")
    ap.add_argument("src", help="입력 .hwp 또는 .hwpx")
    ap.add_argument("dst", help="출력 .docx")
    ap.add_argument("--dsl", help="중간 텍스트를 이 경로에 저장(수정용)")
    ap.add_argument("--from-dsl", action="store_true",
                    help="src를 중간 텍스트(.txt)로 간주하고 바로 렌더링")
    ap.add_argument("--font", default=None, help="한글 글꼴 (기본: 맑은 고딕)")
    args = ap.parse_args()

    if args.from_dsl:
        with open(args.src, encoding="utf-8") as f:
            dsl = f.read()
    else:
        items = read_survey(args.src)
        print(f"추출: 문단 {sum(1 for k, _ in items if k == 'p')}개, "
              f"표 {sum(1 for k, _ in items if k == 'table')}개")
        dsl = items_to_dsl(items)

    if args.dsl:
        with open(args.dsl, "w", encoding="utf-8") as f:
            f.write(dsl)
        print("중간 텍스트:", args.dsl)

    blocks = parse_dsl(dsl)
    print("인식:", summarize(blocks))
    SurveyWriter(font=args.font).write(blocks).save(args.dst)
    print("저장 완료:", args.dst)


if __name__ == "__main__":
    main()

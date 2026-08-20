# -*- coding: utf-8 -*-
"""명령줄 사용: python -m hwp_survey.cli 설문지.hwp 결과.docx [--dsl 중간.txt]"""
import argparse

from . import (DPWriter, ISASWriter, items_to_dp_dsl, items_to_isas_dsl,
               parse_dp, parse_isas, read_survey, summarize_dp, summarize_isas)


def main():
    ap = argparse.ArgumentParser(description="한글 설문지 -> 워드 설문지")
    ap.add_argument("src", help="입력 .hwp 또는 .hwpx")
    ap.add_argument("dst", help="출력 .docx")
    ap.add_argument("--dsl", help="중간 텍스트를 이 경로에 저장(수정용)")
    ap.add_argument("--from-dsl", action="store_true",
                    help="src를 중간 텍스트(.txt)로 간주하고 바로 렌더링")
    ap.add_argument("--font", default=None, help="한글 글꼴")
    ap.add_argument("--renumber", action="store_true",
                    help="ISAS 번호 체계(SQ/Q1-1/DQ)로 다시 매긴다. "
                         "기본은 원본 번호 유지")
    ap.add_argument("--style", choices=["isas", "dp"], default="isas",
                    help="isas=ISAS 표준, dp=DP 스크립트")
    args = ap.parse_args()

    if args.from_dsl:
        with open(args.src, encoding="utf-8") as f:
            dsl = f.read()
    else:
        items = read_survey(args.src)
        print(f"추출: 문단 {sum(1 for k, _ in items if k == 'p')}개, "
              f"표 {sum(1 for k, _ in items if k == 'table')}개")
        dsl = (items_to_isas_dsl(items, renumber=args.renumber)
               if args.style == "isas" else items_to_dp_dsl(items))

    if args.dsl:
        with open(args.dsl, "w", encoding="utf-8") as f:
            f.write(dsl)
        print("중간 텍스트:", args.dsl)

    if args.style == "isas":
        doc = parse_isas(dsl)
        print("인식:", summarize_isas(doc))
        ISASWriter(**({"font": args.font} if args.font else {})).write(doc).save(args.dst)
    else:
        doc = parse_dp(dsl)
        print("인식:", summarize_dp(doc))
        DPWriter(**({"font": args.font} if args.font else {})).write(doc).save(args.dst)
    print("저장 완료:", args.dst)


if __name__ == "__main__":
    main()

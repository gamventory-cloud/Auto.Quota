# -*- coding: utf-8 -*-
"""한글 설문지(.hwp/.hwpx) -> 워드 설문지 변환."""
from .reader import read_survey, read_hwp, read_hwpx
from .parser import items_to_dsl, parse_dsl, summarize

#: items_to_dsl / parse_dsl / summarize 는 DP·ISAS 가 함께 쓰는 공용 중간 표현이다.
__all__ = ["read_survey", "read_hwp", "read_hwpx", "items_to_dsl",
           "parse_dsl", "summarize"]
__version__ = "0.1.0"

from .dp import (DPWriter, build_dp_docx, items_to_dp_dsl, parse_dp,  # noqa: E402
                 split_long_matrices, summarize_dp)

__all__ += ["DPWriter", "build_dp_docx", "items_to_dp_dsl", "parse_dp",
            "split_long_matrices", "summarize_dp"]

from .isas import (ISASWriter, build_isas_docx, items_to_isas_dsl, parse_isas,  # noqa: E402
                   summarize_isas)

__all__ += ["ISASWriter", "build_isas_docx", "items_to_isas_dsl", "parse_isas",
            "summarize_isas"]

# 검증 기능은 선택 사항이다. verify.py 가 아직 없거나 pypdf 가 설치되지 않은
# 환경에서도 변환 기능은 그대로 동작해야 하므로 실패를 삼킨다.
try:
    from .verify import compare, compare_files, docx_text, pdf_text  # noqa: E402

    __all__ += ["compare", "compare_files", "docx_text", "pdf_text"]
    VERIFY_AVAILABLE = True
except ImportError as _verify_error:                        # noqa: BLE001
    VERIFY_AVAILABLE = False
    VERIFY_IMPORT_ERROR = str(_verify_error)

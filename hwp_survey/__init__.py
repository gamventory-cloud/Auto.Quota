# -*- coding: utf-8 -*-
"""한글 설문지(.hwp/.hwpx) -> 워드 설문지 변환."""
from .reader import read_survey, read_hwp, read_hwpx
from .parser import items_to_dsl, parse_dsl, summarize
from .writer import SurveyWriter, build_docx

__all__ = ["read_survey", "read_hwp", "read_hwpx", "items_to_dsl",
           "parse_dsl", "summarize", "SurveyWriter", "build_docx"]
__version__ = "0.1.0"

from .dp import DPWriter, build_dp_docx, items_to_dp_dsl, parse_dp, summarize_dp  # noqa: E402

__all__ += ["DPWriter", "build_dp_docx", "items_to_dp_dsl", "parse_dp",
            "summarize_dp"]

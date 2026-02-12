import streamlit as st
import pandas as pd
import sys
import os
import re
import io
import textwrap
import collections
from collections import Counter

# 워드/엑셀 관련 라이브러리
try:
    from docx import Document
    from docx.document import Document as _Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn 
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    st.error("필수 라이브러리가 설치되지 않았습니다. requirements.txt에 'python-docx'와 'openpyxl'을 추가하세요.")
    st.stop()

# 1. 상위 폴더의 utils.py를 불러오기 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

# 2. 페이지 기본 설정
st.set_page_config(page_title="설문지 코드북 생성", layout="wide")

# 3. 비밀번호 잠금
if not utils.check_password():
    st.stop()

st.title("📝 설문지 읽기 & 코드북/신텍스 자동 생성 (Matrix 인식 강화)")

# ==============================================================================
# [Part 2] 유틸리티 및 텍스트 처리 함수
# ==============================================================================

CIRCLE_MAP = {'①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'}

def clean_empty_parentheses(text):
    if not text: return text
    return re.sub(r"\(\s*\)", "", text).strip()

def clean_header_text(text):
    text = text.strip()
    match = re.search(r"([①-⑩]|\d+)", text)
    if match:
        raw_code = match.group(1)
        code = CIRCLE_MAP.get(raw_code, raw_code)
        label = re.sub(r"[\(\[\{\<]?\s*" + re.escape(raw_code) + r"\s*[\)\]\}\>]?[\.]?", "", text).strip()
        if not label: label = f"{code}점"
        return f"{code}={label}"
    return f"{text}={text}"

def extract_options_from_line(text):
    pattern = re.compile(r"([①-⑩]|(?:\d+|[a-zA-Z])[\)\.])")
    matches = list(pattern.finditer(text))
    if not matches: return []
    results = []
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(text)
        item = text[start:end].strip()
        item = clean_empty_parentheses(item)
        if item: results.append(item)
    return results

def iter_block_items(parent):
    if isinstance(parent, _Document): parent_elm = parent.element.body
    elif isinstance(parent, _Cell): parent_elm = parent._tc
    else: raise ValueError("지원하지 않는 부모 객체입니다.")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P): yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl): yield Table(child, parent)

# ==============================================================================
# [Part 3] 테이블 추출기 (Matrix 대응 강화)
# ==============================================================================

def extract_table_scale(table):
    rows = table.rows
    if len(rows) < 2: return None, False
    headers = [cell.text.strip().replace('\n', ' ') for cell in rows[0].cells]
    first_data_row = [cell.text.strip() for cell in rows[1].cells]
    
    numeric_cells = []
    for cell_text in first_data_row:
        match = re.search(r"([①-⑩]|\d+)", cell_text)
        if match:
            raw_code = match.group(1)
            numeric_cells.append(CIRCLE_MAP.get(raw_code, raw_code))
        else: numeric_cells.append(None)
    
    body_numeric_count = sum(1 for x in numeric_cells if x is not None)
    if len(first_data_row) > 0 and (body_numeric_count / len(first_data_row)) >= 0.3:
        scale_pairs = []
        for i, h_text in enumerate(headers):
            if i >= len(numeric_cells) or numeric_cells[i] is None: continue
            if h_text: scale_pairs.append(f"{numeric_cells[i]}={h_text}")
        if scale_pairs: return "\n".join(scale_pairs), True
    return None, False

def extract_single_choice_options(table):
    options = []
    for row in table.rows:
        cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
        if not cells_text: continue
        match = re.match(r"^([①-⑩]|\d+[\)\.])", cells_text[0])
        if match:
            raw_code = match.group(1).replace(')','').replace('.','')
            code = CIRCLE_MAP.get(raw_code, raw_code)
            label = " - ".join(cells_text[1:]) if len(cells_text) > 1 else cells_text[0][len(match.group(0)):].strip()
            options.append(f"{code}={label}")
    return "\n".join(options)

# ==============================================================================
# [Part 5] 메인 파서 (B1~B4 매트릭스 특화)
# ==============================================================================

def parse_word_to_df(docx_file):
    doc = Document(docx_file)
    extracted_data = []
    var_pattern = re.compile(r"^([a-zA-Z가-힣0-9\-\_]+)(?:[\.\s]|\s+)(.*)")
    current_entry = None
    is_parent_added = False 
    variable_map = {} 

    def flush_entry(entry):
        entry["질문 내용"] = clean_empty_parentheses(entry["질문 내용"])
        raw_options = entry.get("보기_list", [])
        clean_opts = []
        for opt in raw_options:
            opt_match = re.match(r"^\s*([①-⑩]|\d+[\)\.])\s*(.*)", opt)
            if opt_match:
                raw_code = opt_match.group(1).replace(')','').replace('.','')
                code = CIRCLE_MAP.get(raw_code, raw_code)
                clean_opts.append(f"{code}={opt_match.group(2)}")
            else: clean_opts.append(opt)
        entry["보기 값"] = "\n".join(clean_opts)
        if "보기_list" in entry: del entry["보기_list"]
        return [entry]

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text: continue
            match_var = var_pattern.match(text)
            if match_var and any(match_var.group(1).upper().startswith(p) for p in ['Q','S','A','B','C','D']):
                if current_entry and not is_parent_added:
                    extracted_data.extend(flush_entry(current_entry))
                var_name = match_var.group(1).replace("-", "_")
                current_entry = { "변수명": var_name, "질문 내용": match_var.group(2), "보기_list": extract_options_from_line(match_var.group(2)), "유형": "Single" }
                is_parent_added = False
            elif current_entry:
                opts = extract_options_from_line(text)
                if opts: current_entry["보기_list"].extend(opts)
                elif not current_entry["보기_list"]: current_entry["질문 내용"] += " " + text

        elif isinstance(block, Table):
            if not current_entry: continue
            rows = block.rows
            # 매트릭스 여부 및 척도 추출
            table_vals, is_matrix_scale = extract_table_scale(block)
            
            if is_matrix_scale:
                sub_cnt = 0
                for row in rows[1:]:
                    row_label = row.cells[0].text.strip()
                    if not row_label or row_label in ["①", "②"]: continue
                    sub_cnt += 1
                    extracted_data.append({
                        "변수명": f"{current_entry['변수명']}_{sub_cnt}",
                        "질문 내용": f"[{current_entry['변수명']}] {row_label}",
                        "보기 값": table_vals, "유형": "Matrix"
                    })
                is_parent_added = True
            elif not is_parent_added:
                # 단순 보기 테이블인 경우
                opt_str = extract_single_choice_options(block)
                if opt_str:
                    current_entry["보기 값"] = opt_str
                    extracted_data.extend(flush_entry(current_entry))
                    is_parent_added = True

    if current_entry and not is_parent_added:
        extracted_data.extend(flush_entry(current_entry))
            
    return pd.DataFrame(extracted_data)

# (이하 엑셀 생성 및 SPSS UI 로직은 원본과 동일하게 유지)
def to_excel_with_usage_flag(df):
    rows = []
    for _, row in df.iterrows():
        rows.append({ "사용여부": "O", "V변수": "", "변수명": row['변수명'], "질문 내용": row['질문 내용'], "보기(Values)": row['보기 값'] })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(rows).to_excel(writer, index=False)
    return output.getvalue()

tab1, tab2 = st.tabs(["1단계: 워드 ➡️ 엑셀", "2단계: 엑셀 ➡️ SPSS"])
with tab1:
    f = st.file_uploader("설문지(.docx) 업로드", type=["docx"])
    if f and st.button("분석 시작"):
        df_raw = parse_word_to_df(f)
        st.session_state['df_raw'] = df_raw
        st.dataframe(df_raw, use_container_width=True)
        st.download_button("📥 코드북 다운로드", to_excel_with_usage_flag(df_raw), "Codebook.xlsx")

with tab2:
    excel = st.file_uploader("수정된 엑셀 업로드", type=["xlsx"])
    if excel:
        df_edit = pd.read_excel(excel)
        spss_utf8 = utils.generate_spss_final(df_edit, encoding_type='utf-8')
        st.code(spss_utf8, language="spss")
        st.download_button("💾 SPSS 신택스 다운로드", spss_utf8.encode('utf-8-sig'), "Syntax.sps")

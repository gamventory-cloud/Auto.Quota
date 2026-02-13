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

# 3. 비밀번호 잠금 (utils.py 참조)
if not utils.check_password():
    st.stop()

st.title("📝 설문지 읽기 & 코드북/신텍스 자동 생성 (통합 업데이트)")

# ==============================================================================
# [Part 1] 핵심 유틸리티 (동그라미 숫자 대응 추가)
# ==============================================================================

CIRCLE_MAP = {'①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'}

def clean_empty_parentheses(text):
    if not text: return text
    return re.sub(r"\(\s*\)", "", text).strip()

def extract_options_from_line(text):
    # 동그라미 숫자 및 일반 숫자/기호 패턴 대응
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
# [Part 2] 지능형 테이블 분석 및 매트릭스 추출 (B1-B4 강화)
# ==============================================================================

def extract_matrix_info(table):
    """B1~B4와 같은 매트릭스 7점 척도 표에서 척도와 질문을 분리함"""
    rows = table.rows
    if len(rows) < 2: return None, False
    
    # 헤더에서 척도 레이블 추출 (예: 전혀 그렇지 않다, 매우 그렇다 등)
    headers = [cell.text.strip().replace('\n', ' ') for cell in rows[0].cells]
    
    # 첫 데이터 행에서 동그라미 숫자가 있는지 확인하여 척도 값 확정
    first_data_cells = [cell.text.strip() for cell in rows[1].cells]
    scale_values = []
    for cell_text in first_data_cells:
        match = re.search(r"([①-⑩]|\d+)", cell_text)
        if match:
            raw = match.group(1)
            scale_values.append(CIRCLE_MAP.get(raw, raw))
        else: scale_values.append(None)
            
    valid_vals = [v for v in scale_values if v is not None]
    if len(first_data_cells) > 0 and (len(valid_vals) / len(first_data_cells)) >= 0.3:
        scale_pairs = []
        for i, val in enumerate(scale_values):
            if val is not None and i < len(headers) and headers[i]:
                # 중복된 텍스트 제거 및 깔끔한 매핑
                clean_h = re.sub(r"\s+", " ", headers[i]).strip()
                scale_pairs.append(f"{val}={clean_h}")
        return "\n".join(scale_pairs), True
    return None, False

def extract_child_demographics_table(table, current_var):
    headers = [c.text.strip() for c in table.rows[0].cells]
    gender_col_idx = -1; birth_col_idx = -1
    for idx, h in enumerate(headers):
        if "성별" in h: gender_col_idx = idx
        if "생년" in h or "생일" in h: birth_col_idx = idx
    if gender_col_idx == -1 or birth_col_idx == -1: return None 
    extracted_entries = []
    for i, row in enumerate(table.rows[1:]): 
        cells = row.cells
        if len(cells) <= max(gender_col_idx, birth_col_idx): continue
        row_label = cells[0].text.strip()
        if not row_label: continue 
        gender_vals_str = ""
        gender_opts = extract_options_from_line(cells[gender_col_idx].text.strip())
        if gender_opts:
            g_lines = []
            for opt in gender_opts:
                m = re.match(r"^([①-⑩]|\d+|[a-zA-Z])[\)\.]?\s*(.*)", opt)
                if m: 
                    code = CIRCLE_MAP.get(m.group(1), m.group(1).replace(')','').replace('.',''))
                    g_lines.append(f"{code}={m.group(2).strip()}")
            gender_vals_str = "\n".join(g_lines)
        extracted_entries.append({ "변수명": f"{current_var['변수명']}_{i+1}_1", "질문 내용": f"[{current_var['변수명']}] {row_label} - 성별", "보기 값": gender_vals_str, "유형": "Single" })
        extracted_entries.append({ "변수명": f"{current_var['변수명']}_{i+1}_2", "질문 내용": f"[{current_var['변수명']}] {row_label} - 생년", "보기 값": "(숫자입력)", "유형": "Open" })
    return extracted_entries

def analyze_table_structure(table):
    rows = table.rows
    if len(rows) < 1: return "UNKNOWN"
    all_text = " ".join([c.text.strip() for row in rows for c in row.cells])
    
    # 매트릭스 척도 우선 감지
    _, is_matrix = extract_matrix_info(table)
    if is_matrix: return "MATRIX_SCALE"
    
    if "성별" in all_text and ("생년" in all_text or "생일" in all_text): return "CHILD_DEMO"
    if "시간" in all_text and "분" in all_text and ("입력" in all_text or "(" in all_text): return "TIME_SPLIT"
    if "합계" in all_text and ("%" in all_text or "100" in all_text): return "CONSTANT_SUM"
    
    return "STANDARD"

# ==============================================================================
# [Part 3] 메인 파서
# ==============================================================================

def parse_word_to_df(docx_file):
    doc = Document(docx_file)
    extracted_data = []
    var_pattern = re.compile(r"^([a-zA-Z가-힣0-9\-\_]+)(?:[\.\s]|\s+)(.*)")
    multi_keywords = ["복수응답", "중복선택", "모두 골라", "모두 선택", "중복 응답", "중복 선택", "중복 응답 가능"]
    current_entry = None
    is_parent_added = False 

    def flush_entry(entry):
        entry["질문 내용"] = clean_empty_parentheses(entry["질문 내용"])
        raw_options = entry.get("보기_list", [])
        is_multi = any(k in entry["질문 내용"] for k in multi_keywords)
        
        clean_opts_list = []
        for opt in raw_options:
            m = re.match(r"^\s*([①-⑩]|\d+[\)\.])\s*(.*)", opt)
            if m:
                raw_code = m.group(1).replace(')','').replace('.','')
                code = CIRCLE_MAP.get(raw_code, raw_code)
                clean_opts_list.append(f"{code}={m.group(2).strip()}")
        
        if is_multi and clean_opts_list:
            full_val = "\n".join(clean_opts_list)
            results = []
            for opt_str in clean_opts_list:
                c, l = opt_str.split('=', 1)
                results.append({"변수명": f"{entry['변수명']}_{c}", "질문 내용": f"{entry['질문 내용']} ({l})", "보기 값": full_val, "유형": "Multi"})
            return results
        else:
            entry["보기 값"] = "\n".join(clean_opts_list)
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
                current_entry = {"변수명": match_var.group(1).replace("-", "_"), "질문 내용": match_var.group(2), "보기_list": extract_options_from_line(match_var.group(2)), "유형": "Single"}
                is_parent_added = False
            elif current_entry:
                opts = extract_options_from_line(text)
                if opts: current_entry["보기_list"].extend(opts)
                elif not current_entry["보기_list"]: current_entry["질문 내용"] += " " + text

        elif isinstance(block, Table):
            if not current_entry: continue
            t_type = analyze_table_structure(block)
            
            if t_type == "MATRIX_SCALE":
                scale_str, _ = extract_matrix_info(block)
                sub_cnt = 0
                for row in block.rows[1:]:
                    row_label = row.cells[0].text.strip()
                    if not row_label or row_label in CIRCLE_MAP: continue
                    sub_cnt += 1
                    extracted_data.append({"변수명": f"{current_entry['변수명']}_{sub_cnt}", "질문 내용": f"[{current_entry['변수명']}] {row_label}", "보기 값": scale_str, "유형": "Matrix"})
                is_parent_added = True
            elif t_type == "CHILD_DEMO":
                res = extract_child_demographics_table(block, current_entry)
                if res: extracted_data.extend(res); is_parent_added = True
            elif t_type == "STANDARD":
                opts = extract_options_from_line(" ".join([c.text for row in block.rows for c in row.cells]))
                if opts: current_entry["보기_list"].extend(opts)
            
    if current_entry and not is_parent_added:
        extracted_data.extend(flush_entry(current_entry))
    return pd.DataFrame(extracted_data)

# ==============================================================================
# [Part 4] SPSS 신텍스 안전 생성
# ==============================================================================

def generate_spss_safe(df, encoding='utf-8'):
    try:
        # utils.py에 함수가 있을 경우 사용
        return utils.generate_spss_final(df, encoding_type=encoding)
    except (AttributeError, TypeError):
        # 함수가 없거나 인자가 다를 경우 자체 폴백 로직
        syntax = ["* SPSS Syntax Generated (Integrated).", "SET UNICODE=ON." if encoding=='utf-8' else "SET UNICODE=OFF.", "", "VARIABLE LABELS"]
        for _, row in df.iterrows():
            syntax.append(f'  {row["변수명"]} "{row["질문 내용"]}"')
        syntax.append(".\nVALUE LABELS")
        for _, row in df.iterrows():
            val = str(row.get('보기(Values)', row.get('보기 값', '')))
            if val and '=' in val:
                syntax.append(f"  {row['변수명']}")
                for pair in val.split('\n'):
                    if '=' in pair:
                        c, l = pair.split('=', 1)
                        syntax.append(f'    {c} "{l.strip()}"')
        syntax.append(".\nEXECUTE.")
        return "\n".join(syntax)

# ==============================================================================
# [Part 5] UI 및 엑셀 출력
# ==============================================================================

tab1, tab2 = st.tabs(["1단계: 워드 분석", "2단계: SPSS 생성"])

with tab1:
    f = st.file_uploader("설문지(.docx) 업로드", type=["docx"])
    if f and st.button("분석 시작"):
        df_raw = parse_word_to_df(f)
        st.session_state['df_raw'] = df_raw
        st.dataframe(df_raw, use_container_width=True)
        
        output = io.BytesIO()
        df_raw.to_excel(output, index=False)
        st.download_button("📥 코드북 다운로드", output.getvalue(), "Codebook.xlsx")

with tab2:
    excel_file = st.file_uploader("수정된 엑셀 업로드", type=["xlsx"])
    if excel_file:
        df_edit = pd.read_excel(excel_file)
        spss_syntax = generate_spss_safe(df_edit)
        st.code(spss_syntax, language="spss")
        st.download_button("💾 신텍스 다운로드", spss_syntax.encode('utf-8-sig'), "Syntax.sps")
``` [cite: 7, 11, 14, 19, 24, 30, 31, 32, 33, 34, 35, 36, 37, 38, 41, 45, 51, 57, 59, 65, 71, 77, 83]

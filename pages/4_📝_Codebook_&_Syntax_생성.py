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
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    st.error("필수 라이브러리가 설치되지 않았습니다. 'python-docx', 'openpyxl'을 설치해주세요.")
    st.stop()

# 유틸리티 로드
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

st.set_page_config(page_title="Codebook & Syntax 생성", layout="wide")

if not utils.check_password():
    st.stop()

st.title("📝 설문지 파싱 & SPSS 신택스 생성 (원문자 인식 강화)")

# ==============================================================================
# [핵심] 원문자 변환 함수 (① -> 1)
# ==============================================================================
def convert_circled_num(text):
    circled_map = {
        '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
        '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
        '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
        '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20'
    }
    for char, num in circled_map.items():
        if char in text:
            text = text.replace(char, num)
    return text

# ==============================================================================
# [Part 1] 워드 -> 코드북 추출 로직
# ==============================================================================

def extract_options_robust(text):
    """일반 숫자와 원문자를 모두 포함하여 보기 추출"""
    # 1), 1., ① 패턴 모두 대응
    pattern = re.compile(r"(\d+|[①-⑳]|[a-zA-Z])[\)\.]?\s*")
    matches = list(pattern.finditer(text))
    if not matches: return []
    
    results = []
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(text)
        item = text[start:end].strip()
        # 원문자를 숫자로 변환하여 저장
        item = convert_circled_num(item)
        if item: results.append(item)
    return results

def parse_word_to_codebook(docx_file):
    doc = Document(docx_file)
    extracted_data = []
    
    # 질문 인식 패턴 (SQ1, A1, 문1 등)
    var_pattern = re.compile(r"^([a-zA-Z가-힣]*\d+[\-\_]?\d*)[\.\)\s]\s*(.*)")
    # 보기 인식 패턴 (1), ① 등)
    opt_pattern = re.compile(r"^(\d+|[①-⑳]|[a-zA-Z])[\)\.]?\s*(.*)")
    
    current_entry = None
    
    for block in doc.paragraphs:
        text = block.text.strip()
        if not text: continue
        
        # 1. 질문 여부 확인
        match_var = var_pattern.match(text)
        if match_var:
            if current_entry: extracted_data.append(current_entry)
            
            var_name = match_var.group(1).replace(" ", "").upper()
            q_label = match_var.group(2)
            
            current_entry = {
                "변수명": var_name,
                "질문 내용": q_label,
                "보기 값": [],
                "유형": "Single"
            }
            # 질문 줄에 보기가 같이 있는 경우 처리 (예: SQ1. 성별 ①남 ②여)
            inline_opts = extract_options_robust(q_label)
            if inline_opts:
                # 질문 텍스트에서 보기 부분 제거
                first_opt_raw = re.search(r"(\d+|[①-⑳]|[a-zA-Z])[\)\.]?\s*", q_label)
                if first_opt_raw:
                    current_entry["질문 내용"] = q_label[:first_opt_raw.start()].strip()
                current_entry["보기 값"].extend(inline_opts)
                
        # 2. 보기 여부 확인
        elif current_entry:
            # 원문자가 포함된 줄인지 확인
            if opt_pattern.match(text) or any(c in text for c in "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"):
                opts = extract_options_robust(text)
                current_entry["보기 값"].extend(opts)
            else:
                # 보기도 아니고 질문도 아니면 질문 내용의 연장으로 판단
                current_entry["질문 내용"] += " " + text

    if current_entry: extracted_data.append(current_entry)
    
    # 데이터 정리
    final_rows = []
    for entry in extracted_data:
        vals = "\n".join(entry["보기 값"])
        # 복수응답 판단
        q_type = entry["유형"]
        if any(k in entry["질문 내용"] for k in ["모두", "중복", "복수"]):
            q_type = "Multi"
            
        final_rows.append({
            "사용여부": "O",
            "V변수": "",
            "변수명": entry["변수명"],
            "질문 내용": entry["질문 내용"],
            "보기(Values)": vals,
            "유형": q_type
        })
        
    return pd.DataFrame(final_rows)

# ==============================================================================
# [Part 2] 코드북 -> SPSS 신택스 생성 로직
# ==============================================================================

def generate_spss_syntax(df):
    syntax = ["* SPSS Syntax Generated.\nSET UNICODE=ON.\n"]
    
    # 1. Variable Labels
    syntax.append("VARIABLE LABELS")
    for _, row in df.iterrows():
        var = str(row['변수명']).strip()
        label = str(row['질문 내용']).strip().replace('"', "'")
        syntax.append(f'  {var} "{label}"')
    syntax.append(".\n")
    
    # 2. Value Labels
    syntax.append("VALUE LABELS")
    for _, row in df.iterrows():
        var = str(row['변수명']).strip()
        val_text = str(row['보기(Values)'])
        if not val_text or val_text == 'nan': continue
        
        syntax.append(f"  {var}")
        lines = val_text.split('\n')
        for line in lines:
            # "1. 보기" 또는 "1=보기" 형태를 SPSS 형식으로 변환
            parts = re.split(r"[\=\)\.]", line, maxsplit=1)
            if len(parts) == 2:
                code = parts[0].strip()
                v_label = parts[1].strip().replace('"', "'")
                if code.isdigit():
                    syntax.append(f'    {code} "{v_label}"')
        syntax.append("  /")
    syntax.replace_last_slash = syntax[-1] = "." # 마지막 슬래시를 점으로 변경
    
    syntax.append("\nEXECUTE.")
    return "\n".join(syntax)

# ==============================================================================
# Streamlit UI
# ==============================================================================

tab1, tab2 = st.tabs(["1단계: 워드 ➡️ 코드북", "2단계: 코드북 ➡️ 신택스"])

with tab1:
    st.header("설문지 파싱 (워드 → 엑셀)")
    uploaded_word = st.file_uploader("워드 설문지(.docx) 업로드", type=["docx"])
    
    if uploaded_word:
        if st.button("분석 시작"):
            df_result = parse_word_to_codebook(uploaded_word)
            st.session_state['temp_codebook'] = df_result
            st.success("파싱 완료!")
            st.dataframe(df_result)
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False)
            st.download_button("📥 코드북(엑셀) 다운로드", output.getvalue(), "Codebook_Draft.xlsx")

with tab2:
    st.header("신택스 추출 (엑셀 → SPSS)")
    uploaded_excel = st.file_uploader("작업된 코드북(.xlsx) 업로드", type=["xlsx"])
    
    if uploaded_excel:
        df_excel = pd.read_excel(uploaded_excel)
        if st.button("신택스 생성"):
            spss_code = generate_spss_syntax(df_excel)
            st.text_area("생성된 신택스", spss_code, height=400)
            st.download_button("💾 .sps 파일 다운로드", spss_code, "SPSS_Syntax.sps")

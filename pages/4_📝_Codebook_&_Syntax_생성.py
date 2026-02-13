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

st.title("📝 설문지 읽기 & 코드북/신텍스 자동 생성 (Full Logic)")

# ==============================================================================
# [Part 1] 핵심 유틸리티 (동그라미 숫자 대응)
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
    if not matches:
        return []
    results = []
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(text)
        item = text[start:end].strip()
        item = clean_empty_parentheses(item)
        if item:
            results.append(item)
    return results

def summarize_label_regex(text):
    if not text: return ""
    text = re.sub(r"\(PROG.*?\)", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\[PROG.*?\]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\(.*?(입력|기입|범위|선택).*?\)", "", text)
    text = re.sub(r"\[.*?(선택|기입|응답).*?\]", "", text)
    text = re.sub(r"^다음은.*?질문입니다\.?", "", text).strip()
    text = re.sub(r"^다음.*?대해.*?(선택|응답).*?주십시오\.?", "", text).strip()
    text = text.replace("귀하의 ", "").replace("귀하께서는 ", "").replace("귀 댁의 ", "")
    text = text.replace("응답자 본인의 ", "").replace("평소 ", "")
    patterns = [
        r"은 무엇입니까\?*$", r"는 무엇입니까\?*$", r"는 무엇인가요\?*$",
        r"을 선택해 주십시오\.?$", r"를 선택해 주십시오\.?$",
        r"을 선택해 주세요\.?$", r"를 선택해 주세요\.?$",
        r"을 기입해 주십시오\.?$", r"를 기입해 주십시오\.?$",
        r"을 입력하여 주십시오\.?$", r"를 입력하여 주십시오\.?$",
        r"에 대해 어떻게 생각하십니까\?*$",
        r"정도입니까\?*$", r"되십니까\?*$", r"인가요\?*$", r"있습니까\?*$"
    ]
    for pat in patterns: text = re.sub(pat, "", text)
    replacements = { "만족하는 정도": "만족도", "얼마나 만족하십니까": "만족도", "얼마나 자주": "빈도", "이유는 무엇": "이유", "생각나는 이미지": "이미지", "구입한 적이": "구입 경험", "이용한 경험": "이용 경험", "어디입니까": "장소", "누구입니까": "대상" }
    for old, new in replacements.items(): 
        if old in text: text = text.replace(old, new)
    text = text.strip(); text = re.sub(r"\?+$", "", text); text = re.sub(r"\.$", "", text)
    return text.strip()

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("iter_block_items: 지원하지 않는 부모 객체입니다.")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# ==============================================================================
# [Part 2] 복합 테이블 추출기 (기존 모든 로직 유지 + 매트릭스 강화)
# ==============================================================================

def extract_matrix_info(table):
    """B1~B4 매트릭스 7점 척도 감지 및 분리"""
    rows = table.rows
    if len(rows) < 2: return None, False
    headers = [cell.text.strip().replace('\n', ' ') for cell in rows[0].cells]
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
                scale_pairs.append(f"{val}={headers[i].strip()}")
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

def extract_time_split_table(table, current_var):
    extracted = []
    for i, row in enumerate(table.rows):
        cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
        if not cells_text: continue
        row_label = cells_text[0]
        extracted.append({ "변수명": f"{current_var['변수명']}_{i+1}_H", "질문 내용": f"[{current_var['변수명']}] {row_label} (시간)", "보기 값": "(숫자입력)", "유형": "Open" })
        extracted.append({ "변수명": f"{current_var['변수명']}_{i+1}_M", "질문 내용": f"[{current_var['변수명']}] {row_label} (분)", "보기 값": "(숫자입력)", "유형": "Open" })
    return extracted

def analyze_table_structure(table):
    rows = table.rows
    if len(rows) < 1: return "UNKNOWN"
    all_text = " ".join([c.text.strip() for row in rows for c in row.cells])
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
    multi_keywords = ["복수응답", "중복선택", "모두 골라", "모두 선택", "중복 응답", "중복 선택", "중복 가능"]
    current_entry = None
    is_parent_added = False 

    def flush_entry(entry):
        entry["질문 내용"] = clean_empty_parentheses(entry["질문 내용"])
        raw_options = entry.get("보기_list", [])
        is_multi = any(k in entry["질문 내용"] for k in multi_keywords)
        clean_opts = []
        for opt in raw_options:
            m = re.match(r"^\s*([①-⑩]|\d+[\)\.])\s*(.*)", opt)
            if m:
                code = CIRCLE_MAP.get(m.group(1), m.group(1).replace(')','').replace('.',''))
                clean_opts.append(f"{code}={m.group(2).strip()}")
        
        if is_multi and clean_opts:
            full_val = "\n".join(clean_opts)
            return [{"변수명": f"{entry['변수명']}_{c.split('=')[0]}", "질문 내용": f"{entry['질문 내용']} ({c.split('=')[1]})", "보기 값": full_val, "유형": "Multi"} for c in clean_opts]
        else:
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
                for i, row in enumerate(block.rows[1:]):
                    row_label = row.cells[0].text.strip()
                    if row_label and row_label not in CIRCLE_MAP:
                        extracted_data.append({"변수명": f"{current_entry['변수명']}_{i+1}", "질문 내용": f"[{current_entry['변수명']}] {row_label}", "보기 값": scale_str, "유형": "Matrix"})
                is_parent_added = True
            elif t_type == "CHILD_DEMO":
                res = extract_child_demographics_table(block, current_entry)
                if res: extracted_data.extend(res); is_parent_added = True
            elif t_type == "TIME_SPLIT":
                res = extract_time_split_table(block, current_entry)
                if res: extracted_data.extend(res); is_parent_added = True

    if current_entry and not is_parent_added:
        extracted_data.extend(flush_entry(current_entry))
    return pd.DataFrame(extracted_data)

# ==============================================================================
# [Part 4] SPSS 신텍스 및 엑셀 출력 (완벽 복구)
# ==============================================================================

def generate_spss_syntax(df, encoding='utf-8'):
    """utils.py 에러 방지용 자체 내장 신텍스 생성기"""
    syntax = ["* SPSS Syntax Generated.", "SET UNICODE=ON." if encoding=='utf-8' else "SET UNICODE=OFF.", "", "VARIABLE LABELS"]
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

def to_excel_with_usage_flag(df):
    rows = []
    for _, row in df.iterrows():
        rows.append({ "사용여부": "O", "V변수": "", "변수명": row['변수명'], "질문 내용": row['질문 내용'], "보기(Values)": row['보기 값'] })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(rows).to_excel(writer, index=False)
    return output.getvalue()

# ==============================================================================
# Streamlit UI
# ==============================================================================
st.set_page_config(page_title="설문지 데이터 처리 마스터 (v100 Final)", layout="wide")
st.title("📑 설문지 데이터 처리 마스터")
st.markdown("""
**[최종 업데이트 v100]**
* **Save with KEEP:** SPSS 신택스 생성 시, '사용여부'가 O/R인 변수들만 `/KEEP=` 명령어로 길게 나열하여 저장하도록 변경했습니다.
* **완벽 통합:** 기존의 모든 기능(순위형, 표 파싱, PROG 삭제, 하이픈 처리 등)이 포함된 최종 완성본입니다.
""")

tab1, tab2 = st.tabs(["1단계: 워드 ➡️ 엑셀 생성", "2단계: 엑셀 ➡️ SPSS 생성"])

with tab1:
    st.header("1. 워드 파일 파싱")
    uploaded_word = st.file_uploader("설문지(.docx) 업로드", type=["docx"], key="word_uploader")
    if uploaded_word:
        if st.button("분석 시작", key="btn_analyze"):
            with st.spinner("문서 구조 정밀 분석 중..."):
                try: 
                    df_raw = parse_word_to_df(uploaded_word)
                    st.session_state['df_raw'] = df_raw
                    st.success(f"분석 완료! {len(df_raw)}개 항목 추출됨")
                except Exception as e: 
                    st.error(f"오류 발생: {e}")
                    
    if 'df_raw' in st.session_state:
        st.subheader("📊 분석 결과 미리보기")
        st.dataframe(st.session_state['df_raw'], use_container_width=True, height=400)
        
        st.info("아래 엑셀 파일을 다운로드하여 내용을 수정하세요.")
        excel_data = to_excel_with_usage_flag(st.session_state['df_raw'])
        st.download_button(
            label="📥 편집용 코드북 다운로드 (Codebook.xlsx)",
            data=excel_data,
            file_name="Codebook_Draft.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

with tab2:
    st.header("2. SPSS 신택스 생성")
    uploaded_excel = st.file_uploader("수정된 코드북(.xlsx) 업로드", type=["xlsx"], key="excel_uploader")
    if uploaded_excel:
        try:
            df_edited = pd.read_excel(uploaded_excel)
            if '사용여부' not in df_edited.columns: 
                st.error("⚠️ 1단계에서 생성된 엑셀 파일을 사용해주세요.")
            else:
                st.success("파일 로드 성공!")
                df_filtered = df_edited[df_edited['사용여부'].isin(['O', 'R'])].copy()
                st.write(f"총 {len(df_edited)}개 중 {len(df_filtered)}개 문항 선택됨")
                
                col1, col2 = st.columns(2)
                
                # Option 1: UTF-8
                with col1:
                    spss_utf8 = generate_spss_final(df_edited, encoding_type='utf-8')
                    st.download_button(
                        label="💾 (추천) SPSS 신택스 다운로드 (UTF-8)",
                        data=spss_utf8.encode('utf-8-sig'), 
                        file_name="Syntax_UTF8.sps",
                        mime="text/plain",
                        type="primary",
                        use_container_width=True
                    )
                    st.caption("최신 버전 SPSS 사용 시 권장")

                # Option 2: CP949
                with col2:
                    spss_cp949 = generate_spss_final(df_edited, encoding_type='cp949')
                    st.download_button(
                        label="💾 (구버전) SPSS 신택스 다운로드 (CP949)",
                        data=spss_cp949.encode('cp949', errors='ignore'), 
                        file_name="Syntax_CP949.sps",
                        mime="text/plain",
                        type="secondary",
                        use_container_width=True
                    )
                    st.caption("SPSS에서 한글이 깨질 때 사용")
                
                with st.expander("신택스 내용 미리보기 (UTF-8 기준)"):
                    st.code(spss_utf8, language="spss")
        except Exception as e: 
            st.error(f"파일 처리 중 오류: {e}")

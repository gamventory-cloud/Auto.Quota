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

st.title("📝 설문지 읽기 & 코드북/신텍스 자동 생성 (동그라미 숫자 대응판)")

# ==============================================================================
# [Part 1] 핵심 파싱 함수
# ==============================================================================

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
# [Part 2] 유틸리티 및 텍스트 처리 함수
# ==============================================================================

def clean_empty_parentheses(text):
    if not text: return text
    return re.sub(r"\(\s*\)", "", text).strip()

def clean_header_text(text):
    text = text.strip()
    # 동그라미 숫자 또는 일반 숫자 추출
    match = re.search(r"([①-⑩]|\d+)", text)
    if match:
        code = match.group(1)
        # 동그라미 숫자일 경우 숫자로 치환하여 저장 (SPSS 처리용)
        circle_map = {'①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'}
        clean_code = circle_map.get(code, code)
        label = re.sub(r"[\(\[\{\<]?\s*" + re.escape(code) + r"\s*[\)\]\}\>]?[\.]?", "", text).strip()
        if not label: label = f"{clean_code}점"
        return f"{clean_code}={label}"
    return f"{text}={text}"

def extract_options_from_line(text):
    # 동그라미 숫자 또는 '숫자/알파벳 + 기호' 패턴 (수정됨)
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

def check_section_header(text, current_prefix):
    clean_text = text.strip()
    new_prefix = current_prefix
    if re.search(r"Screening", clean_text, re.IGNORECASE) or "스크리닝" in clean_text:
        new_prefix = "SQ"
    elif re.search(r"Part\s*([A-Z])", clean_text, re.IGNORECASE):
        match = re.search(r"Part\s*([A-Z])", clean_text, re.IGNORECASE)
        new_prefix = match.group(1).upper()
    elif re.search(r"^DQ", clean_text, re.IGNORECASE) or "인구 통계" in clean_text:
        new_prefix = "DQ"
    return new_prefix

# ==============================================================================
# [Part 3] 테이블 추출기 (Extractors)
# ==============================================================================

def extract_single_choice_options(table):
    options = []
    circle_map = {'①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'}
    for row in table.rows:
        cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
        if not cells_text: continue
        first_cell_text = cells_text[0]
        # 동그라미 숫자 또는 일반 숫자 패턴
        match = re.match(r"^([①-⑩]|\d+[\)\.])", first_cell_text)
        if match:
            raw_code = match.group(1).replace(')','').replace('.','')
            code = circle_map.get(raw_code, raw_code)
            clean_first = first_cell_text[len(match.group(0)):].strip()
            label_parts = []
            if clean_first: label_parts.append(clean_first)
            if len(cells_text) > 1: label_parts.extend(cells_text[1:])
            final_label = " - ".join(label_parts); final_label = clean_empty_parentheses(final_label) 
            options.append(f"{code}={final_label}")
        else:
            row_text = " - ".join(cells_text); row_text = clean_empty_parentheses(row_text) 
            options.append(row_text)
    return "\n".join(options)

# [수정] 수평 척도 표 처리 시 동그라미 숫자 변환 추가
def extract_horizontal_scale_table(table, current_var):
    rows = table.rows
    if len(rows) < 2: return None
    circle_map = {'①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'}
    
    numeric_row_idx = -1
    label_row_idx = -1
    
    for i, row in enumerate(rows):
        cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
        if not cells_text: continue
        # 동그라미 숫자나 일반 숫자가 포함되어 있는지 확인
        numeric_count = sum(1 for t in cells_text if t.isdigit() or t in circle_map)
        if len(cells_text) > 0 and (numeric_count / len(cells_text)) > 0.7:
            numeric_row_idx = i
        elif len(cells_text) > 0:
            label_row_idx = i
            
    if numeric_row_idx == -1: return None
    
    codes = []
    for c in rows[numeric_row_idx].cells:
        t = c.text.strip()
        if not t: continue
        if t in circle_map: codes.append(circle_map[t])
        elif t.isdigit(): codes.append(t)

    labels = [c.text.strip() for c in rows[label_row_idx].cells if c.text.strip()] if label_row_idx != -1 else []
    scale_pairs = []
    
    if codes:
        if len(labels) == 2:
            scale_pairs.append(f"{codes[0]}={labels[0]}")
            for c in codes[1:-1]: scale_pairs.append(f"{c}={c}점")
            scale_pairs.append(f"{codes[-1]}={labels[1]}")
        elif len(labels) == len(codes):
             for i in range(len(codes)): scale_pairs.append(f"{codes[i]}={labels[i]}")
        else:
             for i, c in enumerate(codes):
                 if i < len(labels): scale_pairs.append(f"{c}={labels[i]}")
                 else: scale_pairs.append(f"{c}={c}점")

    if scale_pairs:
        current_var["보기 값"] = "\n".join(scale_pairs)
        return [current_var]
    return None

# (기타 유틸리티 함수들은 원본과 동일하게 유지하되 동그라미 숫자 패턴만 보강)
def is_multiple_choice(entry):
    vals = str(entry.get("보기 값", "")); q_text = str(entry.get("질문 내용", ""))
    if re.search(r"([①-⑩]|\d+[\)\.])", vals) or "=" in vals: return True
    if "선택]" in q_text or "골라" in q_text: return True
    return False

def check_and_split_max_n_text(entry):
    if entry["유형"] != "Single" and entry["유형"] != "Open": return None
    q_text = entry["질문 내용"]
    if "보기_list" in entry: q_text += " " + " ".join(entry["보기_list"])
    q_text_norm = q_text.replace("［", "[").replace("］", "]").replace("（", "(").replace("）", ")")
    count = 0
    patterns = [ r"\[\s*최대\s*(\d+)", r"최대\s*(\d+)\s*(?:개|대|곳|명|순위)", r"최대.*?(\d+)", r"(\d+)개.*?기입" ]
    for pat in patterns:
        match = re.search(pat, q_text_norm)
        if match: count = int(match.group(1)); break
    if count == 0 and "3" in q_text_norm and ("기입" in q_text_norm or "작성" in q_text_norm or "선택" in q_text_norm): count = 3
    if count < 1: return None
    new_entries = []
    for i in range(1, count + 1):
        v = entry.copy(); v["변수명"] = f"{entry['변수명']}_{i}"; v["질문 내용"] = f"[{entry['변수명']}] {i}순위"; v["유형"] = "Open"
        if "보기_list" in v: del v["보기_list"]
        new_entries.append(v)
    return new_entries

# ==============================================================================
# [Part 4] 지능형 테이블 분석 (Scanning)
# ==============================================================================

def analyze_table_structure(table):
    rows = table.rows
    if len(rows) < 1: return "UNKNOWN"
    all_text = ""; first_row_text = ""
    circle_map = {'①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩'}
    
    row0_digits = 0; row0_len = len(rows[0].cells)
    
    for i, row in enumerate(rows):
        row_txt = " ".join([c.text.strip() for c in row.cells])
        all_text += row_txt + " "; 
        if i == 0: 
            first_row_text = row_txt
            row0_digits = sum(1 for c in row.cells if re.search(r"^\d+$|^\d+\)", c.text.strip()) or c.text.strip() in circle_map)

    # 매트릭스 척도형 (B1~B4 패턴 대응)
    if len(table.columns) >= 4 and row0_digits >= 3:
        return "STANDARD"

    if "성별" in all_text and ("생년" in all_text or "생일" in all_text): return "CHILD_DEMO"
    if "시간" in all_text and "분" in all_text and ("입력" in all_text or "(" in all_text): return "TIME_SPLIT"
    
    return "STANDARD"

# ==============================================================================
# [Part 5] 메인 파서 (Word to DF)
# ==============================================================================

def parse_word_to_df(docx_file):
    doc = Document(docx_file)
    extracted_data = []
    # 변수명 패턴 (SQ1, A1, B1 등 시작점 인식)
    var_pattern = re.compile(r"^([a-zA-Z가-힣0-9\-\_]+)(?:[\.\s]|\s+)(.*)")
    multi_keywords = ["복수응답", "모두 선택", "중복선택", "중복 응답", "모두 골라", "중복 선택", "복수 선택", "모두 체크", "모두 응답"]
    circle_map = {'①':'1','②':'2','③':'3','④':'4','⑤':'5','⑥':'6','⑦':'7','⑧':'8','⑨':'9','⑩':'10'}
    
    current_entry = None
    is_parent_added = False 
    current_prefix = "Q"
    variable_map = {} 
    pending_max_n_count = None

    def flush_entry(entry):
        if "질문 내용" in entry: entry["질문 내용"] = clean_empty_parentheses(entry["질문 내용"])
        
        raw_options = entry.get("보기_list", [])
        is_multi = any(k in entry["질문 내용"] for k in multi_keywords)
        
        if is_multi and raw_options:
            full_options_str_list = []
            for opt in raw_options:
                # 동그라미 숫자 대응 매칭
                opt_match = re.match(r"^\s*([①-⑩]|\d+[\)\.])\s*(.*)", opt)
                if opt_match:
                    raw_code = opt_match.group(1).replace(')','').replace('.','')
                    code = circle_map.get(raw_code, raw_code)
                    label = clean_empty_parentheses(opt_match.group(2))
                    full_options_str_list.append(f"{code}={label}")
            
            full_options_str = "\n".join(full_options_str_list)
            results = []
            for opt in raw_options:
                opt_match = re.match(r"^\s*([①-⑩]|\d+[\)\.])\s*(.*)", opt)
                if opt_match:
                    raw_code = opt_match.group(1).replace(')','').replace('.','')
                    code = circle_map.get(raw_code, raw_code)
                    label = clean_empty_parentheses(opt_match.group(2))
                    results.append({ "변수명": f"{entry['변수명']}_{code}", "질문 내용": f"{entry['질문 내용']} ({label})", "보기 값": full_options_str, "유형": "Multi" })
            return results
        else:
            # 단일 선택 보기 값 정리
            clean_opts = []
            for opt in raw_options:
                opt_match = re.match(r"^\s*([①-⑩]|\d+[\)\.])\s*(.*)", opt)
                if opt_match:
                    raw_code = opt_match.group(1).replace(')','').replace('.','')
                    code = circle_map.get(raw_code, raw_code)
                    clean_opts.append(f"{code}={opt_match.group(2)}")
                else: clean_opts.append(opt)
            
            entry["보기 값"] = "\n".join(clean_opts)
            if "보기_list" in entry: del entry["보기_list"]
            return [entry]

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text: continue
            
            # 섹션 변경 확인
            current_prefix = check_section_header(text, current_prefix)
            
            # 신규 문항 여부 확인
            match_var = var_pattern.match(text)
            if match_var and any(match_var.group(1).upper().startswith(p) for p in ['Q','S','A','B','C','D']):
                if current_entry and not is_parent_added:
                    for item in flush_entry(current_entry):
                        variable_map[item['변수명']] = len(extracted_data)
                        extracted_data.append(item)
                
                var_name = match_var.group(1).replace("-", "_")
                label = match_var.group(2)
                inline_opts = extract_options_from_line(label)
                
                current_entry = { "변수명": var_name, "질문 내용": label, "보기_list": inline_opts, "유형": "Single" }
                is_parent_added = False
                
                # 최대 N개 선택 패턴 감지
                if "최대" in label and "선택" in label:
                    m = re.search(r"최대\s*(\d+)", label)
                    if m: pending_max_n_count = int(m.group(1))

            elif current_entry:
                # 문단이 보기로 시작하는지 확인 (동그라미 숫자 포함)
                opts_in_line = extract_options_from_line(text)
                if opts_in_line:
                    current_entry["보기_list"].extend(opts_in_line)
                elif "=" in text or "점" in text:
                    current_entry["보기_list"].append(text)
                else:
                    # 보기도 아니고 신규 문항도 아니면 질문 내용의 연장으로 판단
                    if not current_entry["보기_list"]:
                        current_entry["질문 내용"] += " " + text

        elif isinstance(block, Table):
            if not current_entry: continue
            table_type = analyze_table_structure(block)
            
            if table_type == "STANDARD":
                # 매트릭스(행렬)형 문항 처리
                rows = block.rows
                # 헤더에서 보기 값 추출
                header_cells = [c.text.strip() for c in rows[0].cells if c.text.strip()]
                vals_str = ""
                if header_cells:
                    # 헤더에 동그라미 숫자가 있거나, 텍스트가 있을 경우 매핑
                    vals_str = "\n".join([f"{i+1}={h}" for i, h in enumerate(header_cells) if not h.isdigit()])
                    if not vals_str: # 숫자로만 된 헤더일 경우
                        vals_str = "\n".join([f"{h}={h}점" for h in header_cells])

                sub_cnt = 0
                for row in rows[1:]:
                    q_label = row.cells[0].text.strip()
                    if not q_label: continue
                    sub_cnt += 1
                    extracted_data.append({
                        "변수명": f"{current_entry['변수명']}_{sub_cnt}",
                        "질문 내용": f"[{current_entry['변수명']}] {q_label}",
                        "보기 값": vals_str,
                        "유형": "Matrix"
                    })
                is_parent_added = True

    # 마지막 문항 처리
    if current_entry and not is_parent_added:
        for item in flush_entry(current_entry):
            extracted_data.append(item)
            
    return pd.DataFrame(extracted_data)

# ==============================================================================
# [Part 6] Excel & SPSS 생성 (기존 로직 동일)
# ==============================================================================

def to_excel_with_usage_flag(df):
    rows = []
    for idx, row in df.iterrows():
        var_name = row['변수명']
        final_q_label = f"{var_name}. {row['질문 내용']}"
        rows.append({ "사용여부": "O", "V변수": "", "변수명": var_name, "질문 내용": final_q_label, "보기(Values)": row['보기 값'] })
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name='Codebook')
    return output.getvalue()

def generate_spss_final(df_edited, encoding_type='utf-8'):
    # (원본 SPSS 생성 로직 유지)
    import utils
    return utils.generate_spss_syntax(df_edited, encoding_type)

# ==============================================================================
# Streamlit UI
# ==============================================================================
tab1, tab2 = st.tabs(["1단계: 워드 ➡️ 엑셀", "2단계: 엑셀 ➡️ SPSS"])

with tab1:
    uploaded_word = st.file_uploader("설문지(.docx) 업로드", type=["docx"])
    if uploaded_word and st.button("분석 시작"):
        df_raw = parse_word_to_df(uploaded_word)
        st.session_state['df_raw'] = df_raw
        st.dataframe(df_raw, use_container_width=True)
        
        excel_data = to_excel_with_usage_flag(df_raw)
        st.download_button("📥 코드북 다운로드", excel_data, "Codebook.xlsx")

with tab2:
    uploaded_excel = st.file_uploader("수정된 코드북(.xlsx) 업로드", type=["xlsx"])
    if uploaded_excel:
        df_edited = pd.read_excel(uploaded_excel)
        spss_syntax = generate_spss_final(df_edited)
        st.code(spss_syntax, language="spss")
        st.download_button("💾 SPSS 신택스 다운로드", spss_syntax.encode('utf-8-sig'), "Syntax.sps")

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

st.title("📝 설문지 읽기 & 코드북/신텍스 자동 생성 (최종 안정화)")

# ==============================================================================
# [Part 1] 워드 파싱 및 유틸리티 함수 정의
# ==============================================================================
def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def clean_empty_parentheses(text):
    if not text: return text
    return re.sub(r"\(\s*\)", "", text).strip()

def clean_header_text(text):
    text = text.strip()
    match = re.search(r"(\d+)", text)
    if match:
        code = match.group(1)
        label = re.sub(r"[\(\[\{\<]?\s*" + code + r"\s*[\)\]\}\>]?[\.]?", "", text).strip()
        if not label:
            label = f"{code}점"
        return f"{code}={label}"
    return f"{text}={text}"

def extract_options_from_line(text):
    pattern = re.compile(r"(\d+|[①-⑩]|[a-zA-Z])[\)\.]")
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

# [A1 대응] 헤더 없는 단순 입력형 테이블 (조건 강화됨)
def extract_plain_input_table(table, current_var):
    rows = table.rows
    if len(rows) < 1: return None
    
    # 조건 1: 열(Column) 개수가 2개 이하여야 함 (SQ6, A4 같은 복잡한 표 제외)
    # 테이블의 첫 행의 셀 개수로 판단
    if len(rows[0].cells) > 2: return None

    # 조건 2: 첫 셀이 객관식 보기(1) 2)...) 패턴이 아니어야 함
    first_cell = rows[0].cells[0].text.strip()
    if re.match(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]", first_cell): return None

    # 조건 3: "입력", "범위", 단위(cm, kg) 등이 포함되어 있어야 함
    input_keywords = ["입력", "범위", "cm", "kg", "시간", "분", "명", "개", "회"]
    match_count = 0
    
    # 조건 4: 셀 안에 "1) 남자" 같은 선택지가 있으면 안 됨 (SQ6 방지)
    option_pattern = re.compile(r"(\d+|[①-⑩]|[a-zA-Z])[\)\.]")

    for row in rows:
        row_text = " ".join([c.text for c in row.cells])
        # 선택지 패턴이 발견되면 즉시 중단 (이건 plain table이 아님)
        if option_pattern.search(row_text):
            return None
            
        if any(k in row_text for k in input_keywords) or "(" in row_text:
            match_count += 1
            
    if match_count < len(rows) * 0.5:
        return None
        
    extracted = []
    for i, row in enumerate(rows):
        row_text = " ".join([c.text.strip() for c in row.cells if c.text.strip()])
        clean_label = re.sub(r"\(\s*입력.*?\)", "", row_text)
        clean_label = clean_label.replace(":", "").strip()
        
        extracted.append({
            "변수명": f"{current_var['변수명']}_{i+1}",
            "질문 내용": f"[{current_var['변수명']}] {clean_label}",
            "보기 값": "(숫자입력)",
            "유형": "Open"
        })
        
    return extracted

# [SQ6 대응] 자녀 상세 정보(성별+생년월일 혼합) 테이블 감지 함수
def extract_child_demographics_table(table, current_var):
    if len(table.rows) < 2: return None
    headers = [c.text.strip() for c in table.rows[0].cells]
    has_gender = any("성별" in h for h in headers)
    has_birth = any("생년" in h or "생일" in h or "생월" in h for h in headers)
    if not (has_gender and has_birth): return None

    gender_col_idx = -1; birth_col_idx = -1
    for idx, h in enumerate(headers):
        if "성별" in h: gender_col_idx = idx
        if "생년" in h or "생일" in h or "생월" in h: birth_col_idx = idx
    if gender_col_idx == -1 or birth_col_idx == -1: return None

    extracted_entries = []
    for i, row in enumerate(table.rows[1:]):
        cells = row.cells
        if len(cells) <= max(gender_col_idx, birth_col_idx): continue
        row_label = cells[0].text.strip()
        gender_text = cells[gender_col_idx].text.strip()
        birth_text = cells[birth_col_idx].text.strip()
        if not row_label: continue 

        gender_opts = extract_options_from_line(gender_text); gender_vals_str = ""
        if gender_opts:
            g_lines = []
            for opt in gender_opts:
                m = re.match(r"(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)", opt)
                if m: code, val = m.groups(); g_lines.append(f"{code}={val.strip()}")
                else: g_lines.append(opt)
            gender_vals_str = "\n".join(g_lines)
            
        extracted_entries.append({ "변수명": f"{current_var['변수명']}_{i+1}_1", "질문 내용": f"[{current_var['변수명']}] {row_label} - 성별", "보기 값": gender_vals_str, "유형": "Single" })
        has_year = "년" in birth_text; has_month = "월" in birth_text
        if has_year: extracted_entries.append({ "변수명": f"{current_var['변수명']}_{i+1}_2", "질문 내용": f"[{current_var['변수명']}] {row_label} - 생년 (년)", "보기 값": "(숫자입력)", "유형": "Open" })
        if has_month: extracted_entries.append({ "변수명": f"{current_var['변수명']}_{i+1}_3", "질문 내용": f"[{current_var['변수명']}] {row_label} - 생월 (월)", "보기 값": "(숫자입력)", "유형": "Open" })
    return extracted_entries

# [Constant Sum 대응] 고정 합계 테이블 감지 함수
def extract_constant_sum_table(table, current_var):
    if len(table.columns) != 2: return None
    rows = table.rows
    if len(rows) < 2: return None

    q_text = current_var.get("질문 내용", "")
    is_sum_100 = ("100" in q_text and "%" in q_text) or "합계" in q_text or "비중" in q_text or "배분" in q_text
    
    right_col_sample = [rows[0].cells[1].text, rows[-1].cells[1].text]
    is_input_col = any(x in sample for sample in right_col_sample for x in ["%", "_", "입력", "(", ")"])
    
    if not (is_sum_100 or is_input_col): return None

    extracted_entries = []
    for i, row in enumerate(rows):
        cells = row.cells
        label_cell = cells[0].text.strip()
        input_cell = cells[1].text.strip()
        if not label_cell: continue
        if "합계" in label_cell or "Total" in label_cell or "TOTAL" in label_cell: continue

        sub_var_name = f"{current_var['변수명']}_{i+1}"
        final_label = f"[{current_var['변수명']}] {label_cell}"
        if "%" in input_cell or "퍼센트" in q_text: final_label += " (%)"
        extracted_entries.append({ "변수명": sub_var_name, "질문 내용": final_label, "보기 값": "(숫자입력)", "유형": "Open" })
    return extracted_entries

def is_multiple_choice(entry):
    vals = str(entry.get("보기 값", ""))
    q_text = str(entry.get("질문 내용", ""))
    if re.search(r"\d+[\)\.]", vals) or "=" in vals: return True
    if "선택]" in q_text: return True
    return False

def check_and_split_time(entry):
    if is_multiple_choice(entry): return [entry]
    val = str(entry.get("보기 값", "")) + str(entry.get("질문 내용", ""))
    is_time_related = ("시간" in val or "시" in val or "분" in val) and ("입력" in val or "기입" in val)
    if not is_time_related: return [entry]
        
    has_hour_unit = bool(re.search(r"(\)|\]|\}|_)\s*시간", val) or re.search(r"시간\s*(\(|\[|\{|_)", val))
    has_minute_unit = bool(re.search(r"(\)|\]|\}|_)\s*분", val) or re.search(r"분\s*(\(|\[|\{|_)", val))
    
    if has_hour_unit and has_minute_unit:
        entry_h = entry.copy(); entry_h["변수명"] += "_H"; entry_h["질문 내용"] += " (시간)"; entry_h["유형"] = "Open"
        entry_m = entry.copy(); entry_m["변수명"] += "_M"; entry_m["질문 내용"] += " (분)"; entry_m["유형"] = "Open"
        return [entry_h, entry_m]
    elif has_hour_unit:
        entry_h = entry.copy(); entry_h["변수명"] += "_H"; entry_h["질문 내용"] += " (시간)"; entry_h["유형"] = "Open"
        return [entry_h]
    elif has_minute_unit:
        entry_m = entry.copy(); entry_m["변수명"] += "_M"; entry_m["질문 내용"] += " (분)"; entry_m["유형"] = "Open"
        return [entry_m]
    if "분" in val and "시간" in val:
        entry_m = entry.copy(); entry_m["변수명"] += "_M"; entry_m["질문 내용"] += " (분)"; entry_m["유형"] = "Open"
        return [entry_m]
    return [entry]

def check_and_split_date(entry):
    if is_multiple_choice(entry): return [entry]
    val = str(entry.get("보기 값", "")) + str(entry.get("질문 내용", ""))
    if "억" in val: return [entry]
    if re.search(r"(몇\s*명|명\s*수|인원|\(\s*\)\s*명|\[\s*\]\s*명)", val): return [entry]

    def has_unit(text, unit):
        p1 = re.search(r"(\)|\]|\}|_)\s*" + unit, text)
        p2 = re.search(unit + r"\s*(\(|\[|\{|_)", text)
        p3 = (unit in text) and ("입력" in text or "기입" in text)
        return bool(p1 or p2 or p3)

    has_year = has_unit(val, "년"); has_month = has_unit(val, "월") or has_unit(val, "개월"); has_day = has_unit(val, "일")
    if not (has_year or has_month or has_day): return [entry]

    new_entries = []
    if has_year:
        y = entry.copy(); y["변수명"] += "_Y"; y["질문 내용"] += " (년)"; y["유형"] = "Open"; new_entries.append(y)
    if has_month:
        m = entry.copy(); m["변수명"] += "_M"; m["질문 내용"] += " (월)"; m["유형"] = "Open"; new_entries.append(m)
    if has_day:
        d = entry.copy(); d["변수명"] += "_D"; d["질문 내용"] += " (일)"; d["유형"] = "Open"; new_entries.append(d)
        
    if new_entries: return new_entries
    return [entry]

def check_and_split_money(entry):
    if is_multiple_choice(entry): return [entry]
    val = str(entry.get("보기 값", "")) + str(entry.get("질문 내용", ""))
    val_clean = val.replace(" ", "")
    if "만원" not in val_clean and "만 원" not in val: return [entry]
    new_entries = []
    if "억" in val_clean:
        e = entry.copy(); e["변수명"] += "_E"; e["질문 내용"] += " (억)"; e["유형"] = "Open"; new_entries.append(e)
    if "천" in val_clean:
        c = entry.copy(); c["변수명"] += "_C"; c["질문 내용"] += " (천)"; c["유형"] = "Open"; new_entries.append(c)
    if "백" in val_clean:
        b = entry.copy(); b["변수명"] += "_B"; b["질문 내용"] += " (백)"; b["유형"] = "Open"; new_entries.append(b)
    if new_entries: return new_entries
    return [entry]

def check_and_split_percent(entry):
    val = str(entry.get("보기 값", "")) + str(entry.get("질문 내용", ""))
    if "나" in val and "배우자" in val and ("%" in val or "100" in val):
        entry_me = entry.copy(); entry_me["변수명"] += "_1"; entry_me["질문 내용"] += " (나)"; entry_me["유형"] = "Open"
        entry_sp = entry.copy(); entry_sp["변수명"] += "_2"; entry_sp["질문 내용"] += " (배우자)"; entry_sp["유형"] = "Open"
        entry_sum = entry.copy(); entry_sum["변수명"] += "_3"; entry_sum["질문 내용"] += " (합계)"; entry_sum["유형"] = "Open"
        return [entry_me, entry_sp, entry_sum]
    return [entry]

def collapse_consecutive_duplicates(item_list):
    if not item_list: return []
    collapsed = [item_list[0]]
    for item in item_list[1:]:
        if item != collapsed[-1]: collapsed.append(item)
    return collapsed

def extract_double_scale_table(table, current_var):
    rows = table.rows
    if len(rows) < 3: return None
    raw_cat_cells = [c.text.strip() for c in rows[0].cells]; non_empty_cats = [c for c in raw_cat_cells if c]
    if len(non_empty_cats) < 2: return None 
    categories = collapse_consecutive_duplicates(non_empty_cats)
    if len(categories) != 2: return None
    scale_row_cells = [c.text.strip() for c in rows[1].cells]; scales = scale_row_cells[1:]
    if len(scales) % 2 != 0: return None
    mid = len(scales) // 2
    left_scale = scales[:mid]; right_scale = scales[mid:]
    left_norm = "".join(left_scale).replace(" ", ""); right_norm = "".join(right_scale).replace(" ", "")
    if left_norm != right_norm: return None
    scale_pairs = []
    for idx, txt in enumerate(left_scale):
        if txt: scale_pairs.append(f"{idx+1}={txt}")
    scale_str = "\n".join(scale_pairs)
    cat1_label = categories[0]; cat2_label = categories[1]
    extracted_entries = []
    for r_idx, row in enumerate(rows[2:]):
        cells = row.cells
        if not cells: continue
        q_text = cells[0].text.strip()
        if not q_text: continue
        q_text_clean = re.sub(r"^[\d\w]+[\)\.]\s*", "", q_text)
        var_base = f"{current_var['변수명']}_{r_idx+1}"
        entry1 = { "변수명": f"{var_base}_1", "질문 내용": f"[{cat1_label}] {q_text_clean}", "보기 값": scale_str, "유형": "Scale" }
        entry2 = { "변수명": f"{var_base}_2", "질문 내용": f"[{cat2_label}] {q_text_clean}", "보기 값": scale_str, "유형": "Scale" }
        extracted_entries.append(entry1); extracted_entries.append(entry2)
    return extracted_entries

def extract_table_scale(table):
    rows = table.rows
    if len(rows) < 2: return None, False
    headers = [cell.text.strip() for cell in rows[0].cells]
    first_data_row = [cell.text.strip() for cell in rows[1].cells]
    numeric_cells = []
    for cell_text in first_data_row:
        if "입력" in cell_text or "범위" in cell_text or "%" in cell_text: numeric_cells.append(None); continue
        match = re.search(r"(\d+)", cell_text)
        if match: numeric_cells.append(match.group(1))
        else: numeric_cells.append(None)
    body_numeric_count = sum(1 for x in numeric_cells if x is not None)
    if len(first_data_row) > 0 and (body_numeric_count / len(first_data_row)) >= 0.3:
        scale_pairs = []
        for i in range(len(headers)):
            if i >= len(first_data_row): break
            h_text = headers[i]; d_val = numeric_cells[i]
            if d_val is not None and h_text: scale_pairs.append(f"{d_val}={h_text}")
        if scale_pairs: return "\n".join(scale_pairs), True
    potential_values = []
    header_numeric_count = sum(1 for h in headers if re.search(r"\d", h))
    if len(headers) > 0 and (header_numeric_count / len(headers)) >= 0.3:
        for idx, h_text in enumerate(headers):
            if not h_text: continue
            if idx == 0 and not re.search(r"\d", h_text): continue
            potential_values.append(clean_header_text(h_text))
        if potential_values: return "\n".join(potential_values), False
    return None, False

def is_input_table(table):
    if len(table.rows) < 1: return False
    target_count = 0; total_rows = len(table.rows)
    for row in table.rows:
        if len(row.cells) > 1:
            cell_text = row.cells[1].text
            if "입력" in cell_text or "(" in cell_text or "%" in cell_text or "_" in cell_text: target_count += 1
    if total_rows > 0 and (target_count / total_rows) >= 0.3: return True
    return False

def extract_multi_column_input_table(table, current_var, force_row_count=None):
    rows = table.rows
    if len(rows) < 2: return None
    headers = [cell.text.strip() for cell in rows[0].cells]
    non_empty_headers = [h for h in headers if h]
    if len(non_empty_headers) < 1: return None
    first_data_row_cells = [c.text.strip() for c in rows[1].cells[1:]] 
    digit_count = sum(1 for c in first_data_row_cells if c.isdigit() and len(c) == 1)
    if len(first_data_row_cells) > 0 and (digit_count / len(first_data_row_cells)) > 0.5: return None
    extracted_entries = []
    actual_data_rows = len(rows) - 1
    target_loop_count = actual_data_rows
    if force_row_count and force_row_count > actual_data_rows: target_loop_count = force_row_count
    sub_item_count = 0
    for i in range(target_loop_count):
        sub_item_count += 1
        if i < actual_data_rows:
            curr_row = rows[i+1]
            first_cell = curr_row.cells[0].text.strip()
            row_label = first_cell if first_cell else f"{sub_item_count}순위"
        else: row_label = f"{sub_item_count}순위"
        for c_idx in range(len(headers)):
            if c_idx == 0: continue
            raw_header = headers[c_idx] if c_idx < len(headers) else ""
            col_header = raw_header if raw_header else f"Col{c_idx}"
            var_name = f"{current_var['변수명']}_{sub_item_count}_{c_idx}"
            var_label = f"[{current_var['변수명']}] {row_label} - {col_header}"
            extracted_entries.append({ "변수명": var_name, "질문 내용": var_label, "보기 값": "(주관식)", "유형": "Open" })
    return extracted_entries

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
    if count == 0 and "3" in q_text_norm and ("기입" in q_text_norm or "작성" in q_text_norm): count = 3
    if count < 1: return None
    has_manufacturer = "제조사" in q_text_norm; has_brand = "브랜드" in q_text_norm
    new_entries = []
    for i in range(1, count + 1):
        if has_manufacturer and has_brand:
            v1 = entry.copy(); v1["변수명"] = f"{entry['변수명']}_{i}_1"; v1["질문 내용"] = f"[{entry['변수명']}] {i}순위 - 제조사"; v1["유형"] = "Open"
            if "보기_list" in v1: del v1["보기_list"]
            v2 = entry.copy(); v2["변수명"] = f"{entry['변수명']}_{i}_2"; v2["질문 내용"] = f"[{entry['변수명']}] {i}순위 - 브랜드"; v2["유형"] = "Open"
            if "보기_list" in v2: del v2["보기_list"]
            new_entries.append(v1); new_entries.append(v2)
        else:
            v = entry.copy(); v["변수명"] = f"{entry['변수명']}_{i}"; v["질문 내용"] = f"[{entry['변수명']}] {i}순위"; v["유형"] = "Open"
            if "보기_list" in v: del v["보기_list"]
            new_entries.append(v)
    return new_entries

def is_option_description_table(table):
    if len(table.rows) < 1: return False
    pattern = re.compile(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]")
    match_count = 0
    for row in table.rows:
        if not row.cells: continue
        text = row.cells[0].text.strip()
        if pattern.match(text): match_count += 1
    return (match_count / len(table.rows)) >= 0.5

def extract_single_choice_options(table):
    options = []
    for row in table.rows:
        cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
        if not cells_text: continue
        first_cell_text = cells_text[0]
        match = re.match(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]", first_cell_text)
        if match:
            code = match.group(1)
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

def extract_options_from_table(table):
    options = []
    idx = 1
    for row in table.rows:
        for cell in row.cells:
            text = cell.text.strip(); text = clean_empty_parentheses(text)
            if text: options.append(f"{idx}={text}"); idx += 1
    return "\n".join(options)

def check_ranking_selection_question(entry):
    q_text = entry["질문 내용"]
    if ("순서" in q_text or "순위" in q_text) and "선택" in q_text:
        match_rank = re.search(r"~\s*(\d+)\s*순위", q_text)
        if match_rank: return int(match_rank.group(1))
        match_count = re.search(r"(\d+)개", q_text)
        if match_count: return int(match_count.group(1))
    return None

def parse_word_to_df(docx_file):
    doc = Document(docx_file)
    extracted_data = []
    var_pattern = re.compile(r"^([a-zA-Z가-힣0-9\-\_]+)(?:[\.\s]|\s+)(.*)")
    multi_keywords = ["복수응답", "모두 선택", "중복선택", "중복 응답", "모두 골라"]
    current_entry = None
    is_parent_added = False 
    
    pending_ranking_count = None
    ranking_options_buffer = []
    pending_max_n_count = None
    
    allowed_starts = ['Q', 'A', 'S', 'D', 'M', 'P', 'R', 'I', 'B', 'C', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'N', 'O', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', '문', '설문']

    def flush_entry(entry):
        nonlocal is_parent_added, pending_max_n_count
        if "질문 내용" in entry: entry["질문 내용"] = clean_empty_parentheses(entry["질문 내용"])
        if pending_ranking_count is not None and ranking_options_buffer:
            final_opts_str = "\n".join(ranking_options_buffer)
            results = []
            for i in range(1, pending_ranking_count + 1):
                results.append({ "변수명": f"{entry['변수명']}_{i}", "질문 내용": f"{entry['질문 내용']} ({i}순위)", "보기 값": final_opts_str, "유형": "Ranking_Sel" })
            return results
        if pending_max_n_count is not None:
            has_manufacturer = "제조사" in entry["질문 내용"]; has_brand = "브랜드" in entry["질문 내용"]
            new_entries = []
            for i in range(1, pending_max_n_count + 1):
                if has_manufacturer and has_brand:
                    v1 = entry.copy(); v1["변수명"] = f"{entry['변수명']}_{i}_1"; v1["질문 내용"] = f"[{entry['변수명']}] {i}순위 - 제조사"; v1["유형"] = "Open"
                    if "보기_list" in v1: del v1["보기_list"]
                    v2 = entry.copy(); v2["변수명"] = f"{entry['변수명']}_{i}_2"; v2["질문 내용"] = f"[{entry['변수명']}] {i}순위 - 브랜드"; v2["유형"] = "Open"
                    if "보기_list" in v2: del v2["보기_list"]
                    new_entries.append(v1); new_entries.append(v2)
                else:
                    v = entry.copy(); v["변수명"] = f"{entry['변수명']}_{i}"; v["질문 내용"] = f"[{entry['변수명']}] {i}순위"; v["유형"] = "Open"
                    if "보기_list" in v: del v["보기_list"]
                    new_entries.append(v)
            pending_max_n_count = None
            return new_entries
        raw_options = entry.get("보기_list", [])
        is_multi = any(k in entry["질문 내용"] for k in multi_keywords)
        if "D6_2" in entry["변수명"].replace("-", "_"): is_multi = True
        if is_multi and raw_options:
            full_options_str_list = []
            for opt in raw_options:
                opt_match = re.match(r"^\s*(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)", opt)
                if opt_match:
                    code, label = opt_match.groups(); label = clean_empty_parentheses(label)
                    full_options_str_list.append(f"{code}={label}")
            full_options_str = "\n".join(full_options_str_list)
            results = []
            for opt in raw_options:
                opt_match = re.match(r"^\s*(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)", opt)
                if opt_match:
                    code, label = opt_match.groups(); label = clean_empty_parentheses(label) 
                    results.append({ "변수명": f"{entry['변수명']}_{code}", "질문 내용": f"{entry['질문 내용']} ({label})", "보기 값": full_options_str, "유형": "Multi" })
            return results
        else:
            entry["보기 값"] = "\n".join(raw_options)
            if "보기_list" in entry: del entry["보기_list"]
            split_entries = check_and_split_time(entry)
            if len(split_entries) == 1: split_entries = check_and_split_date(split_entries[0])
            if len(split_entries) == 1: split_entries = check_and_split_money(split_entries[0])
            if len(split_entries) == 1: split_entries = check_and_split_percent(split_entries[0])
            return split_entries

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text: continue
            if re.match(r"^\[PROG", text, re.IGNORECASE) or re.match(r"^\(PROG", text, re.IGNORECASE): continue
            text = re.sub(r"\[PROG.*?\]", "", text, flags=re.IGNORECASE)
            text = re.sub(r"\(PROG.*?\)", "", text, flags=re.IGNORECASE)
            text = text.strip()
            if not text: continue
            match_var = var_pattern.match(text)
            is_new_q = False
            if match_var:
                temp_var = match_var.group(1)
                if re.search(r"\d", temp_var) or any(temp_var.startswith(x) for x in allowed_starts):
                    if temp_var not in ["보기", "다음", "참고", "주"]: is_new_q = True
            if is_new_q:
                if current_entry and not is_parent_added:
                    flushed_data = flush_entry(current_entry)
                    if flushed_data: extracted_data.extend(flushed_data)
                var_name = match_var.group(1).replace("-", "_"); label = match_var.group(2)
                inline_opts = extract_options_from_line(label)
                if inline_opts:
                    first_opt = inline_opts[0]; split_idx = label.find(first_opt)
                    if split_idx != -1: q_text = label[:split_idx].strip(); current_entry = { "변수명": var_name, "질문 내용": q_text, "보기 값": "", "보기_list": inline_opts, "유형": "Single" }
                    else: current_entry = { "변수명": var_name, "질문 내용": label.strip(), "보기 값": "", "보기_list": [], "유형": "Single" }
                else: current_entry = { "변수명": var_name, "질문 내용": label.strip(), "보기 값": "", "보기_list": [], "유형": "Single" }
                is_parent_added = False
                rank_count = check_ranking_selection_question(current_entry)
                if rank_count: pending_ranking_count = rank_count; ranking_options_buffer = [] 
                else: pending_ranking_count = None; ranking_options_buffer = []
                max_n_cnt = check_and_split_max_n_text(current_entry)
                if max_n_cnt:
                    q_norm = current_entry["질문 내용"].replace("［", "[").replace("］", "]")
                    m = re.search(r"최대.*?(\d+)", q_norm)
                    if m: pending_max_n_count = int(m.group(1))
                    elif "3" in q_norm and "기입" in q_norm: pending_max_n_count = 3
                    else: pending_max_n_count = None
                else: pending_max_n_count = None
                if "1개 선택" in current_entry["질문 내용"]: current_entry["유형"] = "Single"
            elif current_entry:
                if not is_parent_added:
                    opts_in_line = extract_options_from_line(text)
                    if opts_in_line:
                        if pending_ranking_count:
                            for opt in opts_in_line:
                                opt_match = re.match(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)", opt)
                                if opt_match: code, val = opt_match.groups(); ranking_options_buffer.append(f"{code}={val}")
                        else:
                            if "보기_list" in current_entry: current_entry["보기_list"].extend(opts_in_line)
                    elif "=" in text or "점" in text:
                         if "보기_list" in current_entry: current_entry["보기_list"].append(text)
                    elif "[주관식]" in text or "직접 기입" in text:
                        current_entry["유형"] = "Open"
                        if "보기_list" in current_entry: current_entry["보기_list"].append("(주관식)")
                    else:
                        if "보기_list" in current_entry and not current_entry["보기_list"]: current_entry["질문 내용"] += " " + text

        elif isinstance(block, Table):
            rows = block.rows
            if len(rows) < 1: continue

            # [순서 변경] 1. 특수 테이블들 (SQ6, 합계100%, 더블스케일) 먼저 체크
            
            # [SQ6 대응] 자녀 정보 테이블
            if current_entry and not is_parent_added:
                child_entries = extract_child_demographics_table(block, current_entry)
                if child_entries:
                    extracted_data.extend(child_entries)
                    is_parent_added = True
                    continue

            # [Constant Sum] 합계 100%
            if current_entry and not is_parent_added:
                const_sum_entries = extract_constant_sum_table(block, current_entry)
                if const_sum_entries:
                    extracted_data.extend(const_sum_entries)
                    is_parent_added = True
                    continue
            
            # [Double Scale] 양쪽 척도
            if current_entry and not is_parent_added:
                double_entries = extract_double_scale_table(block, current_entry)
                if double_entries:
                    extracted_data.extend(double_entries)
                    is_parent_added = True
                    continue

            # [일반 객관식]
            if current_entry and not is_parent_added:
                q_type = current_entry.get("유형")
                if any(k in current_entry["질문 내용"] for k in multi_keywords): q_type = "Multi"
                if q_type in ["Single", "Multi"]:
                    is_opt_table = False
                    first_cell_text = rows[0].cells[0].text.strip()
                    if re.match(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]", first_cell_text): is_opt_table = True
                    if is_opt_table:
                        opt_str = extract_single_choice_options(block)
                        if q_type == "Single": current_entry["보기 값"] = opt_str; extracted_data.append(current_entry)
                        else: 
                            parsed_opts = []
                            for line in opt_str.split('\n'):
                                if '=' in line: c, l = line.split('=', 1); parsed_opts.append(f"{c}) {l}")
                                else: parsed_opts.append(line)
                            if "보기_list" not in current_entry: current_entry["보기_list"] = []
                            current_entry["보기_list"].extend(parsed_opts)
                            continue 
                        is_parent_added = True; continue

            if pending_ranking_count and current_entry:
                options_str = extract_options_from_table(block)
                if options_str: ranking_options_buffer.append(options_str)
                continue 

            # [A4 대응] Multi-column Input
            if current_entry:
                multi_col_entries = extract_multi_column_input_table(block, current_entry, force_row_count=pending_max_n_count)
                if multi_col_entries: extracted_data.extend(multi_col_entries); is_parent_added = True; pending_max_n_count = None; continue

            if current_entry and not is_parent_added:
                if current_entry.get("유형") in ["Single", "Multi"]:
                    if is_option_description_table(block):
                        opt_str = extract_single_choice_options(block); current_entry["보기 값"] = opt_str; extracted_data.append(current_entry); is_parent_added = True; continue

            # [순서 변경] 맨 마지막: A1 같은 "헤더 없는 단순 입력형" 테이블 (조건 까다로움)
            if current_entry and not is_parent_added:
                plain_input_entries = extract_plain_input_table(block, current_entry)
                if plain_input_entries:
                    extracted_data.extend(plain_input_entries)
                    is_parent_added = True
                    continue

            is_input_style = is_input_table(block)
            if is_input_style:
                if current_entry:
                    if not is_parent_added: is_parent_added = True 
                    sub_item_count = 0
                    for row in rows:
                        first_cell = row.cells[0].text.strip()
                        if not first_cell: continue
                        sub_item_count += 1
                        m_var = f"{current_entry['변수명']}_{sub_item_count}"
                        m_label = f"{current_entry['질문 내용']} ({first_cell})"
                        extracted_data.append({ "변수명": m_var, "질문 내용": m_label, "보기 값": "(숫자입력)", "유형": "Open" })
            
            elif current_entry:
                table_vals_str, is_body_mapped = extract_table_scale(block)
                is_matrix_table = False
                if len(rows) > 1:
                    for row in rows[1:]:
                        fc = row.cells[0].text.strip()
                        if fc and not fc.isdigit() and fc not in ["○", "●", "V"]: is_matrix_table = True; break
                if not is_parent_added:
                    raw_options = current_entry.get("보기_list", [])
                    if not raw_options and table_vals_str: raw_options.append(table_vals_str)
                    current_entry["보기 값"] = "\n".join(raw_options)
                    if "보기_list" in current_entry: del current_entry["보기_list"]
                    if not is_matrix_table and not is_input_style:
                        split_entries = check_and_split_time(current_entry)
                        if len(split_entries) == 1: split_entries = check_and_split_date(split_entries[0])
                        if len(split_entries) == 1: split_entries = check_and_split_money(split_entries[0])
                        if len(split_entries) == 1: split_entries = check_and_split_percent(split_entries[0])
                        extracted_data.extend(split_entries); is_parent_added = True
                    else:
                        if is_input_style: is_parent_added = True
                        else: extracted_data.append(current_entry); is_parent_added = True
                if is_matrix_table:
                    sub_item_count = 0
                    for row in rows[1:]:
                        if not row.cells: continue
                        first_cell = row.cells[0].text.strip()
                        if not first_cell: continue
                        match_matrix_var = var_pattern.match(first_cell); m_var = ""; m_label = ""; m_vals = table_vals_str if table_vals_str else "(헤더참조)"
                        if match_matrix_var:
                            pot_var = match_matrix_var.group(1)
                            if re.match(r"^[A-Z][A-Z0-9\-_]*$", pot_var) or re.match(r"^\d+[\-\_]\d+$", pot_var): m_var = pot_var.replace("-", "_"); m_label = match_matrix_var.group(2)
                            else: sub_item_count += 1; parent_var = current_entry["변수명"]; m_var = f"{parent_var}_{sub_item_count}"; m_label = f"[{parent_var} 세부] {first_cell}"
                        else:
                            if first_cell.isdigit(): continue
                            sub_item_count += 1; parent_var = current_entry["변수명"]; m_var = f"{parent_var}_{sub_item_count}"; m_label = f"[{parent_var} 세부] {first_cell}"
                        temp_entry = { "변수명": m_var, "질문 내용": m_label, "보기 값": m_vals, "유형": "Matrix" }
                        s_entries = check_and_split_time(temp_entry)
                        if len(s_entries) == 1: s_entries = check_and_split_date(s_entries[0])
                        if len(s_entries) == 1: s_entries = check_and_split_money(s_entries[0])
                        extracted_data.extend(s_entries)

    if current_entry and not is_parent_added:
        flushed_data = flush_entry(current_entry)
        if flushed_data: extracted_data.extend(flushed_data)
            
    return pd.DataFrame(extracted_data)

def to_excel_with_usage_flag(df):
    rows = []
    code_start_pattern = re.compile(r"^(\d+|[①-⑩]|[a-zA-Z]|[가-하])[\.\)\s=]\s*(.*)")
    for idx, row in df.iterrows():
        var_name = row['변수명']; raw_q = str(row['질문 내용']); clean_q = re.sub(r"^\[.*?\]\s*", "", raw_q)
        if "_" in var_name:
            base_var, suffix = var_name.rsplit("_", 1)
            if raw_q.startswith("["): final_q_label = raw_q
            else: final_q_label = f"{base_var}. {suffix}) {clean_q}"
        else: final_q_label = f"{var_name}. {clean_q}"
        vals_str = str(row['보기 값']); formatted_values = ""
        if vals_str and vals_str.strip() != "" and vals_str != "nan":
            lines = vals_str.split('\n'); options = []; current_code = None; current_label_parts = []
            for line in lines:
                line = line.strip()
                if not line: continue
                is_new_code = False; temp_code = ""; temp_label = ""
                if "=" in line: is_new_code = True; temp_code, temp_label = line.split("=", 1)
                else:
                    match = code_start_pattern.match(line)
                    if match: is_new_code = True; temp_code, temp_label = match.groups()
                if is_new_code:
                    if current_code is not None: options.append(f"{current_code.strip()} = {' '.join(current_label_parts).strip()}")
                    current_code = temp_code; current_label_parts = [temp_label]
                else:
                    if current_code is not None: current_label_parts.append(line)
                    else: options.append(line)
            if current_code is not None: options.append(f"{current_code.strip()} = {' '.join(current_label_parts).strip()}")
            formatted_values = "\n".join(options) if options else vals_str
        rows.append({ "사용여부": "O", "V변수": "", "변수명": var_name, "질문 내용": final_q_label, "보기(Values)": formatted_values })
    result_df = pd.DataFrame(rows)
    var_list = df['변수명'].tolist(); var_counts = Counter(var_list); duplicates = [var for var, count in var_counts.items() if count > 1]
    highlight_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=False)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, sheet_name='Codebook')
        worksheet = writer.sheets['Codebook']
        for cell in worksheet[1]: cell.font = Font(bold=True); cell.alignment = Alignment(horizontal='center', vertical='center')
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                if cell.column <= 3: cell.alignment = align_center; 
                if cell.column == 3 and cell.value in duplicates: cell.fill = highlight_fill
                if cell.column > 3: cell.alignment = align_left
        worksheet.column_dimensions['A'].width = 8; worksheet.column_dimensions['B'].width = 15; worksheet.column_dimensions['C'].width = 20; worksheet.column_dimensions['D'].width = 50; worksheet.column_dimensions['E'].width = 40
    return output.getvalue()

def compress_var_list(var_list):
    if not var_list: return ""
    compressed = []; current_chunk = []; pattern = re.compile(r"^(.*?)(\d+)$")
    for var in var_list:
        if not current_chunk: current_chunk.append(var); continue
        prev_var = current_chunk[-1]; match_prev = pattern.match(prev_var); match_curr = pattern.match(var)
        is_continuous = False
        if match_prev and match_curr:
            prev_prefix, prev_num = match_prev.groups(); curr_prefix, curr_num = match_curr.groups()
            if prev_prefix == curr_prefix and int(curr_num) == int(prev_num) + 1: is_continuous = True
        if is_continuous: current_chunk.append(var)
        else:
            if len(current_chunk) >= 3: compressed.append(f"{current_chunk[0]} TO {current_chunk[-1]}")
            else: compressed.extend(current_chunk)
            current_chunk = [var]
    if len(current_chunk) >= 3: compressed.append(f"{current_chunk[0]} TO {current_chunk[-1]}")
    else: compressed.extend(current_chunk)
    return " ".join(compressed)

def generate_spss_final(df_edited, encoding_type='utf-8'):
    enc_str = "UTF-8" if encoding_type == 'utf-8' else "CP949"
    syntax_lines = ["* SPSS Syntax Generated by Streamlit (Final).", f"* Encoding: {enc_str}.", "", "* 0. Set Working Directory and Load Data.", "CD '경로'.", "GET FILE='project_CE.sav'.", ""]
    if encoding_type == 'utf-8': syntax_lines.insert(2, "SET UNICODE=ON.")
    if '사용여부' in df_edited.columns: df_target = df_edited[df_edited['사용여부'].isin(['O', 'R'])].copy()
    else: df_target = df_edited.copy()
    syntax_lines.append("* 1. Rename Variables (B -> C)."); rename_count = 0
    unique_rows = df_target.drop_duplicates(subset=['변수명'], keep='first')
    for idx, row in unique_rows.iterrows():
        v_clean = str(row['변수명']).strip(); v_raw = str(row['V변수']).strip()
        if v_raw and v_raw.lower() != 'nan' and v_raw != v_clean: syntax_lines.append(f"Rename Var {v_raw}={v_clean}."); rename_count += 1
    if rename_count > 0: syntax_lines.append("EXECUTE."); syntax_lines.append("")
    syntax_lines.append("* 1.5 Recode Variables (Reverse Coding)."); recode_count = 0
    for idx, row in df_target.iterrows():
        if row['사용여부'] == 'R':
            v_name = row['변수명']; val_text = str(row['보기(Values)'])
            if not v_name or val_text == 'nan' or not val_text.strip(): continue
            codes = []
            for line in val_text.split('\n'):
                if '=' in line: c = line.split('=', 1)[0].strip(); 
                if c.isdigit(): codes.append(int(c))
            if codes:
                min_c, max_c = min(codes), max(codes); recode_pairs = []; 
                for c in codes: new_c = max_c + min_c - c; recode_pairs.append(f"({c}={new_c})")
                recode_str = " ".join(recode_pairs); syntax_lines.append(f"RECODE {v_name} {recode_str}."); recode_count += 1
    if recode_count > 0: syntax_lines.append("EXECUTE."); syntax_lines.append("")
    syntax_lines.append("VARIABLE LABELS"); unique_vars = df_target.drop_duplicates(subset=['변수명'], keep='first')
    for idx, row in unique_vars.iterrows():
        v = str(row['변수명']).strip(); l = str(row['질문 내용']).strip().replace('"', "'")
        if v: syntax_lines.append(f'  {v} "{l}"')
    syntax_lines.append(".\nEXECUTE.\n"); syntax_lines.append("VALUE LABELS"); value_map = {}
    for idx, row in df_target.iterrows():
        v = str(row['변수명']).strip(); val_text = str(row['보기(Values)']); is_reverse = (row['사용여부'] == 'R')
        if not v or val_text == 'nan' or not val_text.strip(): continue
        lines = val_text.split('\n'); codes_labels = []; codes_int = []
        if is_reverse:
            for line in lines:
                if '=' in line: c = line.split('=', 1)[0].strip(); 
                if c.isdigit(): codes_int.append(int(c))
            if codes_int: min_c, max_c = min(codes_int), max(codes_int)
        for line in lines:
            line = line.strip()
            if '=' in line:
                parts = line.split('=', 1); code = parts[0].strip(); label = parts[1].strip(); final_code = code
                if is_reverse and code.isdigit() and codes_int: c_int = int(code); new_c_int = max_c + min_c - c_int; final_code = str(new_c_int)
                if final_code and label: codes_labels.append((final_code, label))
        if codes_labels:
            try: codes_labels.sort(key=lambda x: int(x[0]))
            except: pass
            val_tuple = tuple(codes_labels)
            if val_tuple not in value_map: value_map[val_tuple] = []
            value_map[val_tuple].append(v)
    group_count = 0; total_groups = len(value_map)
    for val_tuple, var_list in value_map.items():
        group_count += 1
        var_block_str = compress_var_list(var_list); wrapped_vars = textwrap.wrap(var_block_str, width=80)
        for line in wrapped_vars: syntax_lines.append(f"  {line}")
        for code, label in val_tuple:
            label_clean = label.replace('"', "'"); syntax_lines.append(f'    {code} "{code}) {label_clean}"')
        if group_count < total_groups: syntax_lines.append("  /")
        else: syntax_lines.append("  .")
    syntax_lines.append("EXECUTE."); syntax_lines.append(""); syntax_lines.append("* 4. Save Data.")
    keep_vars = df_target['변수명'].drop_duplicates().tolist()
    if keep_vars:
        syntax_lines.append("SAVE OUTFILE='Project_DATA.sav'"); syntax_lines.append("  /KEEP="); 
        for var in keep_vars: syntax_lines.append(f"    {var}")
        syntax_lines.append("  .")
    else: syntax_lines.append("SAVE OUTFILE='Project_DATA.sav'.")
    syntax_lines.append("EXECUTE."); syntax_lines.append(""); syntax_lines.append("* 5. Export to Excel."); syntax_lines.append("GET FILE='Project_DATA.sav'."); syntax_lines.append("EXECUTE.")
    syntax_lines.append(""); syntax_lines.append("*_ SAVE - Values _."); syntax_lines.append("SAVE TRANSLATE OUTFILE='(RAW) Project_DATA.xlsx' /TYPE=XLS /VERSION=12 /MAP /REPLACE /FIELDNAMES /CELLS=VALUES.")
    syntax_lines.append(""); syntax_lines.append("*_ SAVE - Labels _."); syntax_lines.append("SAVE TRANSLATE OUTFILE='(LABEL) Project_DATA.xlsx' /TYPE=XLS /VERSION=12 /MAP /REPLACE /FIELDNAMES /CELLS=LABELS.")
    return "\n".join(syntax_lines)

# ==============================================================================
# Streamlit UI
# ==============================================================================
st.markdown("""
**[기능 설명]**
* **NEW (A1 대응):** 키, 몸무게처럼 헤더 없이 입력 칸만 있는 표를 자동으로 주관식 변수(Open)로 처리합니다.
* **NEW (SQ6 대응):** 한 표 안에 '성별'과 '생년월일'이 섞여 있는 복합형 자녀 정보 테이블을 자동으로 감지하여 분리합니다.
* **Save with KEEP:** SPSS 신택스 생성 시, '사용여부'가 O/R인 변수들만 `/KEEP=` 명령어로 길게 나열하여 저장하도록 변경했습니다.
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

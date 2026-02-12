import streamlit as st
import pandas as pd
import sys
import os
import re
import io
import collections

# 워드/엑셀 라이브러리
try:
    from docx import Document
    from docx.document import Document as _Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    st.error("필수 라이브러리(python-docx, openpyxl)가 설치되지 않았습니다.")
    st.stop()

# 유틸리티 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

# 페이지 설정
st.set_page_config(page_title="설문지 구조화 파싱 (V2)", layout="wide")

# 비밀번호 체크
if not utils.check_password():
    st.stop()

st.title("🧩 설문지 구조화 파싱 엔진 (V2: ETL 방식)")
st.markdown("""
기존 방식과 달리 **[문서 평탄화] -> [구조 분석] -> [변수 생성]** 3단계 공정을 거쳐, 
복잡한 표나 숨겨진 자동 번호를 더욱 안정적으로 처리합니다.
""")

# ==============================================================================
# [Step 1] 문서 평탄화 (Flattening)
# : 워드(XML)의 복잡성을 제거하고, 사람이 읽기 쉬운 선형 리스트로 변환
# ==============================================================================

def iter_block_items(parent):
    """문서의 흐름대로 Paragraph와 Table을 순서대로 반환"""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("지원하지 않는 부모 객체입니다.")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_flattened_content(doc):
    """
    워드 파일을 읽어 [TYPE, CONTENT] 형태의 리스트로 변환
    - TYPE: HEADER(섹션), TEXT(일반글), OPTION(보기), TABLE(표)
    """
    flattened = []
    
    # 자동 번호 인식을 위한 카운터 {(numId, ilvl): count}
    auto_num_counters = collections.defaultdict(int)
    
    current_section = "Common" # 기본 섹션
    
    for block in iter_block_items(doc):
        # 1. 텍스트(Paragraph) 처리
        if isinstance(block, Paragraph):
            text = block.text.strip()
            
            # (1) 섹션 헤더 감지
            if re.match(r"^Part\s*[A-Z]", text, re.IGNORECASE):
                current_section = text
                flattened.append({"type": "SECTION", "content": text})
                continue
            if "Screening" in text or "스크리닝" in text:
                current_section = "SQ"
                flattened.append({"type": "SECTION", "content": "Screening"})
                continue
            if re.match(r"^DQ", text, re.IGNORECASE) or "인구 통계" in text:
                current_section = "DQ"
                flattened.append({"type": "SECTION", "content": "DQ"})
                continue

            # (2) 워드 자동 번호(Auto Numbering) 추출
            if block._p.pPr is not None and block._p.pPr.numPr is not None:
                try:
                    num_id = block._p.pPr.numPr.numId.val
                    ilvl = block._p.pPr.numPr.ilvl.val if block._p.pPr.numPr.ilvl is not None else 0
                    auto_num_counters[(num_id, ilvl)] += 1
                    num_val = auto_num_counters[(num_id, ilvl)]
                    
                    # 텍스트에 이미 번호가 없으면 강제 부착
                    if not re.match(r"^(\d+|[a-zA-Z])[\)\.]", text):
                        # 길이가 길거나 물음표가 있으면 질문(Q), 아니면 보기(Opt)로 추정
                        if len(text) > 40 or "?" in text or "시오" in text:
                            text = f"Q{num_val}. {text}" # 임시 마킹
                        else:
                            text = f"{num_val}) {text}"
                except:
                    pass
            
            if not text: continue
            
            # (3) 텍스트 유형 분류
            # 질문 (Q1. SQ1. A1. 등)
            if re.match(r"^([A-Z]*\d+[\-\_]?\d*)[\.\)]", text) or re.match(r"^Q\d+", text):
                flattened.append({"type": "QUESTION", "content": text, "section": current_section})
            # 보기 (1) 1. ① 등)
            elif re.match(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]", text):
                flattened.append({"type": "OPTION", "content": text})
            else:
                # 그 외 (안내문구 등) -> 앞 질문의 부가 설명일 수 있음
                flattened.append({"type": "TEXT", "content": text})

        # 2. 표(Table) 처리 -> 2차원 리스트로 변환
        elif isinstance(block, Table):
            table_data = []
            for row in block.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            
            if table_data:
                flattened.append({"type": "TABLE", "content": table_data, "section": current_section})

    return flattened

# ==============================================================================
# [Step 2 & 3] 구조 분석 및 코드북 생성 (Logic Application)
# ==============================================================================

def analyze_and_generate_codebook(flattened_data):
    codebook = []
    
    # 상태 변수들
    current_q = None # 현재 처리 중인 질문 {var, label, type, options...}
    
    # 질문 번호 카운터 (자동 부여용)
    q_counters = collections.defaultdict(int)
    
    # 변수 매핑 테이블 처리를 위한 인덱스 맵 {var_name: index_in_codebook}
    var_index_map = {} 
    
    # 정규식 패턴
    var_pattern = re.compile(r"^([A-Z]*\d+[\-\_]?\d*)[\.\)]\s*(.*)") # A1. 질문
    opt_pattern = re.compile(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)") # 1) 보기
    
    def flush_current_q():
        """현재 질문을 마무리하고 코드북에 등록"""
        nonlocal current_q
        if not current_q: return

        # 1. 보기 리스트를 텍스트로 변환
        opts = current_q.get('options', [])
        opt_text = ""
        if opts:
            lines = []
            for o in opts:
                # 이미 code=label 형태면 그대로, 아니면 변환
                if "=" in o: lines.append(o)
                else:
                    m = opt_pattern.match(o)
                    if m: lines.append(f"{m.group(1)}={m.group(2)}")
                    else: lines.append(o)
            opt_text = "\n".join(lines)
        
        current_q['values'] = opt_text
        
        # 2. Max N개 선택 로직 (변수 쪼개기)
        # 텍스트에 "최대 N개" 또는 "순서대로 N개" 등이 있으면 Ranking_Sel로 변경
        q_text = current_q['label']
        max_n = 0
        if "최대" in q_text and "개" in q_text:
            m = re.search(r"최대\s*(\d+)", q_text)
            if m: max_n = int(m.group(1))
        elif "순서대로" in q_text and "개" in q_text:
             m = re.search(r"(\d+)개", q_text)
             if m: max_n = int(m.group(1))
             
        if max_n > 1:
            # 1순위, 2순위... 변수 생성
            base_var = current_q['var']
            for i in range(1, max_n + 1):
                new_entry = {
                    "var": f"{base_var}_{i}",
                    "label": f"[{base_var}] {q_text} ({i}순위/선택)",
                    "type": "Ranking_Sel",
                    "values": opt_text
                }
                codebook.append(new_entry)
                var_index_map[new_entry['var']] = len(codebook) - 1
        
        # 3. 복수응답 (모두 선택) 로직
        elif "모두 선택" in q_text or "중복" in q_text or "복수" in q_text:
             # 보기가 있으면 보기별로 쪼개기 (Multi)
             if opts:
                 for o in opts:
                     m = opt_pattern.match(o)
                     if m:
                         code, label = m.group(1), m.group(2)
                         var_name = f"{current_q['var']}_{code}"
                         new_entry = {
                             "var": var_name,
                             "label": f"[{current_q['var']}] {q_text} ({label})",
                             "type": "Multi",
                             "values": opt_text # 전체 보기를 넣어줄지, 0/1로 할지는 선택. 보통 전체 보기 참조용으로 넣음
                         }
                         codebook.append(new_entry)
                         var_index_map[new_entry['var']] = len(codebook) - 1
             else:
                 # 보기가 아직 없으면(나중에 표에서 올 수도 있음) 일단 단일로 저장
                 codebook.append({
                     "var": current_q['var'], "label": q_text, "type": "Multi_Pending", "values": ""
                 })
                 var_index_map[current_q['var']] = len(codebook) - 1
        
        # 4. 일반 단일응답/주관식
        else:
            q_type = "Open" if ("직접 기입" in q_text or "입력" in q_text) else "Single"
            codebook.append({
                "var": current_q['var'],
                "label": q_text,
                "type": q_type,
                "values": opt_text
            })
            var_index_map[current_q['var']] = len(codebook) - 1
            
        current_q = None

    # --- Main Loop ---
    for item in flattened_data:
        itype = item['type']
        content = item['content']
        
        if itype == "QUESTION":
            flush_current_q()
            # 변수명과 질문 분리
            match = var_pattern.match(content)
            if match:
                var_name = match.group(1).replace("-", "_")
                label = match.group(2)
                
                # 임시 번호(Q1, Q2..)인 경우 섹션 접두어 붙이기
                if var_name.startswith("Q") and item.get('section') != "Common":
                    # 이미 섹션이 붙어있지 않다면 (예: SQ1이 아니라 Q1인 경우)
                    if item['section'] == "SQ" and not var_name.startswith("SQ"):
                        var_name = "SQ" + var_name[1:]
                    elif len(item['section']) == 1 and not var_name.startswith(item['section']):
                        # Part A -> A1
                        var_name = item['section'] + var_name[1:]

                current_q = {
                    "var": var_name,
                    "label": label,
                    "options": []
                }
            else:
                # 매칭 안되면 텍스트로 취급
                if current_q: current_q['label'] += " " + content

        elif itype == "OPTION":
            if current_q:
                current_q['options'].append(content)
        
        elif itype == "TEXT":
            if current_q:
                current_q['label'] += " " + content
                
        elif itype == "TABLE":
            # 표 처리 전략: 표의 특징을 보고 어떤 유형인지 판단
            table = content # list of lists
            if not table: continue
            
            # A. 보기 매핑 테이블 (SQ8, SQ8-1 등) - 헤더에 변수명이 있는 경우
            header = table[0]
            mapped_vars = []
            for idx, h in enumerate(header):
                clean_h = re.sub(r"[^A-Z0-9\_]", "", h.upper().replace("-", "_"))
                if clean_h and len(clean_h) >= 2: # 최소 SQ, A1 등 2글자
                    mapped_vars.append((idx, clean_h))
            
            if len(mapped_vars) >= 1 and "보기" in "".join(header):
                # 매핑 로직 실행
                opt_col = -1
                for i, h in enumerate(header): 
                    if "보기" in h: opt_col = i; break
                
                if opt_col != -1:
                    # 표 내용을 읽어서 각 변수에 할당
                    var_options = {v: [] for _, v in mapped_vars}
                    for row in table[1:]:
                        if len(row) <= opt_col: continue
                        opt_text = row[opt_col]
                        # 코드가 있으면 추출
                        code = ""; val = opt_text
                        m = opt_pattern.match(opt_text)
                        if m: code, val = m.group(1), m.group(2)
                        
                        for col_idx, v_name in mapped_vars:
                            if col_idx == opt_col: continue
                            if len(row) > col_idx and row[col_idx].strip():
                                # 체크된 값이 있으면 해당 변수의 보기로 추가
                                final_code = row[col_idx].strip() if row[col_idx].strip().isdigit() else code
                                if final_code:
                                    var_options[v_name].append(f"{final_code}={val}")
                    
                    # 변수 업데이트 (과거 변수 + 현재 변수)
                    for v_name, opts_list in var_options.items():
                        # 현재 작성 중인 변수라면
                        if current_q and current_q['var'] == v_name:
                            current_q['options'] = opts_list # 덮어쓰기
                        
                        # 이미 작성된 변수라면 (Retroactive Update)
                        elif v_name in var_index_map:
                            target_idx = var_index_map[v_name]
                            target_item = codebook[target_idx]
                            
                            # 기존 변수가 Multi_Pending 이었다면 Multi로 변환하며 폭파
                            if target_item['type'] == 'Multi_Pending' or target_item['type'] == 'Multi':
                                # 기존꺼 지우고 새로 폭파 (간략화: Values만 업데이트하고 타입은 Multi 유지)
                                # 원래는 여기서 Explode해야 하지만, 복잡도를 줄이기 위해 Values 업데이트로 처리
                                target_item['values'] = "\n".join(opts_list)
                                target_item['type'] = 'Multi' # 확정
                                
                                # 만약 Multi인데 단일 변수 하나만 있다면 -> 폭파 필요 (고급 로직)
                                # (이 부분은 사용자 요청 시 추가)
                            else:
                                target_item['values'] = "\n".join(opts_list)
                                
                continue # 표 처리 완료
            
            # B. 가로형 척도 (B1-1)
            # 조건: 숫자로만 된 행이 있다
            num_row_idx = -1
            lbl_row_idx = -1
            for r_i, row in enumerate(table):
                nums = [c for c in row if c.isdigit()]
                if len(nums) >= 3 and len(nums)/len([c for c in row if c]) > 0.7:
                    num_row_idx = r_i
                elif any(c for c in row):
                    lbl_row_idx = r_i
            
            if num_row_idx != -1 and current_q:
                # 척도 매핑
                codes = [c for c in table[num_row_idx] if c.isdigit()]
                labels = [c for c in table[lbl_row_idx] if c] if lbl_row_idx != -1 else []
                
                scale_opts = []
                if labels:
                    # 양극단 매핑 (1=전혀, 7=매우)
                    scale_opts.append(f"{codes[0]}={labels[0]}")
                    if len(labels) >= 2:
                        scale_opts.append(f"{codes[-1]}={labels[-1]}")
                    # 중간값들은 그냥 숫자로
                    for c in codes[1:-1]:
                        scale_opts.append(f"{c}={c}")
                else:
                    scale_opts = [f"{c}={c}점" for c in codes]
                
                current_q['options'] = scale_opts
                current_q['type'] = "Scale"
                continue

            # C. 단위 입력 (SQ6) - 가족수 등
            # 조건: '명', '세' 등의 단위가 포함된 열이 있다
            unit_col = -1
            for c_i, cell in enumerate(table[0]):
                if any(u in cell for u in ["명", "세", "개", "원"]): unit_col = c_i; break
            
            if unit_col != -1 or (len(table)>1 and any("입력" in c for c in table[0])):
                # 현재 질문 flush하고, 표의 각 행을 하위 질문으로 등록
                flush_current_q() # 상위 질문 저장
                base_var = codebook[-1]['var'] # 방금 저장된 변수명
                
                for r_i, row in enumerate(table):
                    label = row[0] # 첫 열을 라벨로 가정
                    if not label: continue
                    if "구분" in label or "입력" in label: continue # 헤더 스킵
                    
                    codebook.append({
                        "var": f"{base_var}_{r_i+1}",
                        "label": f"[{base_var}] {label}",
                        "type": "Open",
                        "values": "(숫자입력)"
                    })
                continue
            
            # D. 그 외 일반 표 -> 보기가 나열된 것으로 간주 (1열이 코드, 2열이 값 등)
            if current_q:
                # 단순 보기 추가
                for row in table:
                    clean_row = [c for c in row if c]
                    for cell in clean_row:
                         if opt_pattern.match(cell):
                             current_q['options'].append(cell)

    # 마지막 질문 처리
    flush_current_q()
    
    return pd.DataFrame(codebook)

# ==============================================================================
# [UI] Streamlit 인터페이스
# ==============================================================================

st.header("1. 설문지 업로드 및 분석")
uploaded_file = st.file_uploader("설문지(.docx) 파일 업로드", type=["docx"])

if uploaded_file:
    if st.button("분석 시작"):
        with st.spinner("1단계: 문서 평탄화 (Flattening) 진행 중..."):
            doc = Document(uploaded_file)
            flattened_data = extract_flattened_content(doc)
            st.success(f"평탄화 완료! 총 {len(flattened_data)}개의 블록 추출")
            
            # 디버깅용: 평탄화 결과 일부 보여주기
            with st.expander("평탄화된 데이터 확인 (Debug)"):
                st.write(flattened_data[:20])

        with st.spinner("2단계: 구조 분석 및 코드북 생성 중..."):
            df_codebook = analyze_and_generate_codebook(flattened_data)
            st.session_state['df_codebook_v2'] = df_codebook
            st.success(f"생성 완료! 총 {len(df_codebook)}개 변수 추출")

if 'df_codebook_v2' in st.session_state:
    st.subheader("📊 생성된 코드북")
    st.dataframe(st.session_state['df_codebook_v2'], use_container_width=True, height=500)
    
    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state['df_codebook_v2'].to_excel(writer, index=False)
    
    st.download_button(
        label="📥 엑셀 다운로드",
        data=output.getvalue(),
        file_name="Codebook_V2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
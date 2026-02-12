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
st.set_page_config(page_title="설문지 구조화 파싱 (V3)", layout="wide")

# 비밀번호 체크
if not utils.check_password():
    st.stop()

st.title("🧩 설문지 구조화 파싱 엔진 (V3: 인식률 강화)")
st.markdown("""
**[개선 사항]**
* **문항 인식 강화:** `SQ1`, `A1`, `Q1.` 등 다양한 문항 번호 패턴을 더 유연하게 잡아냅니다.
* **줄바꿈 처리:** 한 문단 안에 질문과 보기가 섞여 있어도 엔터(`\n`) 기준으로 분리하여 인식합니다.
* **표 인식 개선:** 표 안에 숨어있는 질문과 보기를 더 정확하게 추출합니다.
""")

# ==============================================================================
# [Step 1] 문서 평탄화 (Flattening)
# ==============================================================================

def iter_block_items(parent):
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
    flattened = []
    auto_num_counters = collections.defaultdict(int)
    current_section = "Common"
    
    # 정규식 패턴 정의 (미리 컴파일)
    # 질문 패턴: (문자열)(숫자)(특수문자)(공백)(내용)
    # 예: SQ1. 질문 / A-1) 질문 / [Q1] 질문 / 문1. 질문
    re_question = re.compile(r"^[\(\[]?([A-Za-z가-힣]*\s*\d+(?:[\-\_]\d+)?)[\)\]\.\:]?\s+(.*)")
    
    # 보기 패턴: (숫자/문자)(특수문자)(공백)(내용)
    # 예: 1) 보기 / ① 보기 / a. 보기
    re_option = re.compile(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)")

    for block in iter_block_items(doc):
        # 1. 텍스트(Paragraph) 처리
        if isinstance(block, Paragraph):
            # 1-1. 워드 자동 번호 처리
            if block._p.pPr is not None and block._p.pPr.numPr is not None:
                try:
                    num_id = block._p.pPr.numPr.numId.val
                    ilvl = block._p.pPr.numPr.ilvl.val if block._p.pPr.numPr.ilvl is not None else 0
                    auto_num_counters[(num_id, ilvl)] += 1
                    num_val = auto_num_counters[(num_id, ilvl)]
                    
                    # 텍스트에 번호가 없으면 강제 병합
                    # 단, 번호가 1, 2, 3... 인지 1), 2) 인지는 알 수 없으므로 텍스트 문맥으로 추측
                    raw_text = block.text.strip()
                    if raw_text and not re.match(r"^[\(\[]?(\d+|[a-zA-Z])[\)\.\:]", raw_text):
                        if "?" in raw_text or "시오" in raw_text or len(raw_text) > 40:
                            # 질문으로 추정되면 Q넘버링은 나중에 하고 일단 텍스트만 둠
                            pass 
                        else:
                            # 보기로 추정되면 번호 붙임
                            block.text = f"{num_val}) {raw_text}"
                except:
                    pass

            # 1-2. 줄바꿈(\n) 기준으로 텍스트 분리 (중요!)
            full_text = block.text.strip()
            if not full_text: continue
            
            lines = full_text.split('\n')
            
            for text in lines:
                text = text.strip()
                if not text: continue

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

                # (2) 유형 분류
                # 보기(Option) 우선 체크
                if re_option.match(text):
                    flattened.append({"type": "OPTION", "content": text})
                
                # 질문(Question) 체크
                elif re_question.match(text):
                    # 보기 패턴이랑 비슷하지만 질문인 경우 (예: 1. 다음 중...) 구분
                    # 보통 질문은 길이가 길거나 '?'가 있음
                    flattened.append({"type": "QUESTION", "content": text, "section": current_section})
                
                # 그 외 (Text)
                else:
                    flattened.append({"type": "TEXT", "content": text})

        # 2. 표(Table) 처리
        elif isinstance(block, Table):
            table_data = []
            for row in block.rows:
                # 빈 셀 제외하고 텍스트만 추출
                row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_data:
                    table_data.append(row_data)
            
            if table_data:
                flattened.append({"type": "TABLE", "content": table_data, "section": current_section})

    return flattened

# ==============================================================================
# [Step 2 & 3] 구조 분석 및 코드북 생성
# ==============================================================================

def analyze_and_generate_codebook(flattened_data):
    codebook = []
    current_q = None 
    var_index_map = {} 
    
    # 정규식 (분석용)
    re_q_split = re.compile(r"^[\(\[]?([A-Za-z가-힣]*\s*\d+(?:[\-\_]\d+)?)[\)\]\.\:]?\s+(.*)")
    re_opt_split = re.compile(r"^(\d+|[①-⑩]|[a-zA-Z])[\)\.]\s*(.*)")
    
    def flush_current_q():
        nonlocal current_q
        if not current_q: return

        # 보기 처리
        opts = current_q.get('options', [])
        opt_lines = []
        for o in opts:
            m = re_opt_split.match(o)
            if m: opt_lines.append(f"{m.group(1)}={m.group(2)}")
            elif "=" in o: opt_lines.append(o)
            else: opt_lines.append(o) # 그냥 텍스트인 경우
        
        opt_text = "\n".join(opt_lines)
        q_label = current_q['label']
        var_name = current_q['var']
        
        # 로직: Max N개 / 복수응답 / 주관식 / 단일응답 결정
        
        # 1. Max N (순위형)
        max_n = 0
        norm_label = q_label.replace("[", "").replace("]", "")
        if "최대" in norm_label and "개" in norm_label:
            m = re.search(r"최대\s*(\d+)", norm_label)
            if m: max_n = int(m.group(1))
        elif "순서대로" in norm_label and "개" in norm_label:
             m = re.search(r"(\d+)개", norm_label)
             if m: max_n = int(m.group(1))
             
        if max_n > 1:
            for i in range(1, max_n + 1):
                new_entry = { "var": f"{var_name}_{i}", "label": f"[{var_name}] {q_label} ({i}순위)", "type": "Ranking_Sel", "values": opt_text }
                codebook.append(new_entry)
                var_index_map[new_entry['var']] = len(codebook) - 1
                
        # 2. 복수응답 (Multi)
        elif any(k in q_label for k in ["모두 선택", "중복", "복수", "모두 골라"]):
             if opts:
                 for o in opts:
                     m = re_opt_split.match(o)
                     if m:
                         c, l = m.group(1), m.group(2)
                         v_name = f"{var_name}_{c}"
                         new_entry = { "var": v_name, "label": f"[{var_name}] {q_label} ({l})", "type": "Multi", "values": opt_text }
                         codebook.append(new_entry)
                         var_index_map[v_name] = len(codebook) - 1
             else:
                 # 보기가 없으면 일단 Multi 타입으로 저장 (나중에 표에서 채워질 수 있음)
                 codebook.append({ "var": var_name, "label": q_label, "type": "Multi", "values": "" })
                 var_index_map[var_name] = len(codebook) - 1
                 
        # 3. 주관식 / 단일응답
        else:
            q_type = "Open" if ("직접 기입" in q_label or "입력" in q_label) else "Single"
            codebook.append({ "var": var_name, "label": q_label, "type": q_type, "values": opt_text })
            var_index_map[var_name] = len(codebook) - 1
            
        current_q = None

    # --- Main Loop ---
    for item in flattened_data:
        itype = item['type']
        content = item['content']
        
        if itype == "QUESTION":
            flush_current_q()
            m = re_q_split.match(content)
            if m:
                # 변수명 정제 (SQ 1 -> SQ1, A-1 -> A1)
                raw_var = m.group(1).replace(" ", "").replace("-", "_").upper()
                
                # Q1, Q2 등 임시 번호인 경우 섹션 붙이기
                if raw_var.startswith("Q") and len(raw_var) < 4:
                    sec = item.get('section', 'Common')
                    if sec != 'Common' and not raw_var.startswith(sec):
                        raw_var = sec + raw_var[1:] # Q1 -> SQ1 or A1
                        
                current_q = { "var": raw_var, "label": m.group(2), "options": [] }
            else:
                # 매칭 실패 시 그냥 텍스트로 처리
                if current_q: current_q['label'] += " " + content

        elif itype == "OPTION":
            if current_q: current_q['options'].append(content)
        
        elif itype == "TEXT":
            if current_q: current_q['label'] += " " + content
                
        elif itype == "TABLE":
            # 표 처리 (보기 매핑 등)
            table = content
            if not table: continue
            
            # (A) 보기 매핑 테이블 (SQ8 등)
            header = table[0]
            header_str = "".join(header)
            
            # 변수명 매핑 로직
            mapped_vars = []
            opt_col = -1
            
            for idx, col_text in enumerate(header):
                if "보기" in col_text: opt_col = idx
                # 헤더가 변수명처럼 생겼는지 확인 (SQ8, SQ8-1 ...)
                clean_h = re.sub(r"[^A-Z0-9\_]", "", col_text.upper().replace("-", "_"))
                if clean_h and len(clean_h) >= 2:
                    mapped_vars.append((idx, clean_h))
            
            if opt_col != -1 and mapped_vars:
                # 보기를 추출하여 각 변수에 할당
                var_opts = {v: [] for _, v in mapped_vars}
                
                for row in table[1:]:
                    if len(row) <= opt_col: continue
                    opt_text = row[opt_col]
                    
                    # 보기 코드/값 분리
                    code = ""; val = opt_text
                    m = re_opt_split.match(opt_text)
                    if m: code, val = m.group(1), m.group(2)
                    
                    for c_idx, v_name in mapped_vars:
                        if c_idx == opt_col: continue
                        if len(row) > c_idx and row[c_idx]: # 값이 있으면 해당 보기 사용
                            final_code = row[c_idx] if row[c_idx].isdigit() else code
                            if final_code: var_opts[v_name].append(f"{final_code}={val}")
                
                # extracted_data (codebook) 업데이트
                for v_name, opts in var_opts.items():
                    # 현재 변수라면
                    if current_q and current_q['var'] == v_name:
                        current_q['options'] = [o.replace("=", ") ", 1) for o in opts] # 포맷 맞춤
                    # 이미 저장된 변수라면 (Retroactive)
                    elif v_name in var_index_map:
                        idx = var_index_map[v_name]
                        # 기존 값 덮어쓰기
                        codebook[idx]['values'] = "\n".join(opts)
                        # 만약 Multi였다면 여기서 폭파(Explode) 로직을 다시 수행해야 할 수도 있음 (여기선 생략)

            # (B) 척도형 테이블 (B1-1 등)
            # 숫자 행 찾기
            num_row_idx = -1
            lbl_row_idx = -1
            for i, row in enumerate(table):
                digits = [x for x in row if x.isdigit()]
                if len(digits) >= 3: num_row_idx = i
                elif any(x for x in row): lbl_row_idx = i
            
            if num_row_idx != -1 and current_q:
                codes = [x for x in table[num_row_idx] if x.isdigit()]
                labels = table[lbl_row_idx] if lbl_row_idx != -1 else []
                
                scale_opts = []
                # 매핑: 양극단
                if codes and labels:
                    if len(labels) >= 2:
                        scale_opts.append(f"{codes[0]}={labels[0]}")
                        scale_opts.append(f"{codes[-1]}={labels[-1]}")
                    else:
                        scale_opts = [f"{c}={c}" for c in codes]
                elif codes:
                    scale_opts = [f"{c}={c}" for c in codes]
                
                current_q['options'] = scale_opts
                current_q['type'] = "Scale"

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
            
            with st.expander("평탄화된 데이터 확인 (Debug)"):
                st.write(flattened_data[:30]) # 앞부분만 확인

        with st.spinner("2단계: 구조 분석 및 코드북 생성 중..."):
            df_codebook = analyze_and_generate_codebook(flattened_data)
            st.session_state['df_codebook_v3'] = df_codebook
            st.success(f"생성 완료! 총 {len(df_codebook)}개 변수 추출")

if 'df_codebook_v3' in st.session_state:
    st.subheader("📊 생성된 코드북")
    st.dataframe(st.session_state['df_codebook_v3'], use_container_width=True, height=500)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state['df_codebook_v3'].to_excel(writer, index=False)
    
    st.download_button(
        label="📥 엑셀 다운로드",
        data=output.getvalue(),
        file_name="Codebook_V3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

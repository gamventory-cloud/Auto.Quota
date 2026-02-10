import streamlit as st
import pandas as pd
import io
import collections
import traceback
import sys
import os

# (주의) utils 모듈이 같은 폴더나 상위 폴더에 있어야 합니다.
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

st.set_page_config(page_title="SPSS 변수명 정제", layout="wide")

if not utils.check_password():
    st.stop()

st.header("📊 SPSS 변수명 자동 정제 & 신텍스 생성")
st.markdown("""
**Raw 데이터**와 **Code북**을 비교하여 SPSS 변수명 변경 신텍스를 생성합니다.
* **Code북 규칙:** 1열=변수명(Q1), **2열=질문라벨(SQ1. 성별...)**
* **기능 1:** 라벨의 앞부분(SQ1)을 추출하여 변수명으로 자동 변환
* **기능 2:** 척도 문항 등으로 변수명이 중복될 경우, 자동으로 `_1`, `_2`, `_3`을 붙여서 구분
* **기능 3:** 엑셀 다운로드 시 **순수 데이터(디자인 없음)** + **1행: 새변수명, 2행: 기존변수명** 적용
""")

# 1. 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="spss_file_uploader")

if uploaded_file:
    try:
        # 엑셀 파일 로드 및 시트명 확인
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        # 시트 선택 UI
        col1, col2 = st.columns(2)
        with col1:
            raw_sheet = st.selectbox("Raw 데이터 시트", sheet_names, index=0, key="raw_sheet_select")
        with col2:
            # 보통 Code북은 뒤쪽에 있으므로 자동 선택 시도
            code_idx = 2 if len(sheet_names) > 2 else (1 if len(sheet_names) > 1 else 0)
            code_sheet = st.selectbox("Code북 시트", sheet_names, index=code_idx, key="code_sheet_select")
        
        # 분석 시작 버튼
        if st.button("분석 시작", key="analyze_btn"):
            with st.spinner('데이터 분석 및 매칭 중...'):
                # [NEW] 분석 시작 시 모든 시트를 미리 읽어둠 (다운로드용)
                st.session_state['spss_all_sheets'] = pd.read_excel(uploaded_file, sheet_name=None)
                st.session_state['spss_target_sheets'] = [raw_sheet] # 기본 타겟은 선택한 Raw 시트

                # 데이터프레임 로드 (분석용)
                df_raw = st.session_state['spss_all_sheets'][raw_sheet]
                
                # [수정] header=None 옵션 추가: 첫 번째 줄(Q1)도 데이터로 읽기 위해
                df_code = pd.read_excel(uploaded_file, sheet_name=code_sheet, header=None)
                
                # Raw 데이터 컬럼 매핑 (소문자 -> 원본)
                raw_cols_map = {str(col).strip().lower(): str(col).strip() for col in df_raw.columns}
                
                temp_vars = []
                
                # --- [Step 1] Code북 순회 (무조건 1, 2열 사용) ---
                for idx, row in df_code.iterrows():
                    if len(row) < 2: continue
                    if pd.isna(row.iloc[0]): continue
                    
                    col_a_val = utils.clean_text(row.iloc[0]) # 변수명 (Code) - 예: Q1
                    col_c_val = utils.clean_text(row.iloc[1]) # 질문 라벨 - 예: SQ1. 성별
                    
                    if not col_a_val: continue
                    
                    # [핵심] 라벨에서 기본 이름 추출 (예: "SQ1. 성별" -> "SQ1")
                    label_base = utils.extract_base_name(col_c_val)
                    if not label_base: 
                        label_base = col_a_val # 실패 시 Code명 사용

                    # [스마트 매칭 로직]
                    # 1. 정확히 일치하는 경우
                    if col_a_val.lower() in raw_cols_map:
                        raw_original = raw_cols_map[col_a_val.lower()]
                        new_var_name = utils.sanitize_var_name(label_base)
                        
                        temp_vars.append({
                            "Raw 변수명": raw_original,
                            "Code 변수명": col_a_val,
                            "질문 내용": col_c_val,
                            "변경할 변수명": new_var_name,
                            "상태": "매칭 성공"
                        })

                    # 2. 복수응답/세트 문항 탐색 (예: Q5 -> q5_1, q5_2...)
                    prefix = col_a_val.lower() + "_"
                    found_multiples = []
                    for rc_lower, rc_original in raw_cols_map.items():
                        if rc_lower.startswith(prefix):
                            found_multiples.append((rc_lower, rc_original))
                    
                    # 찾은 복수응답 컬럼들 추가
                    for _, rc_original in found_multiples:
                        # 접미사 추출
                        suffix = rc_original[len(col_a_val):] 
                        if not suffix.startswith('_') and not suffix.startswith('-'):
                            suffix = "_" + suffix

                        # 라벨 기반 이름 + 접미사
                        new_name = utils.sanitize_var_name(label_base + suffix)
                        
                        temp_vars.append({
                            "Raw 변수명": rc_original,
                            "Code 변수명": col_a_val,
                            "질문 내용": col_c_val,
                            "변경할 변수명": new_name,
                            "상태": "매칭 성공 (세트)"
                        })

                # --- [Step 2] 중복 변수명 처리 로직 (추가됨) ---
                # 1. 먼저 생성된 모든 변수명의 빈도수를 체크
                name_freq = collections.Counter([item['변경할 변수명'] for item in temp_vars])
                
                # 2. 중복 카운터 준비
                name_counter = collections.defaultdict(int)
                
                final_data = []
                seen_raw = set()
                
                # 3. 리스트를 다시 돌면서 중복인 경우 번호 부여
                for item in temp_vars:
                    # 이미 처리한 Raw 변수는 패스
                    if item['Raw 변수명'] in seen_raw: continue
                    
                    candidate_name = item['변경할 변수명']
                    
                    # 중복이 발생하는 이름인 경우에만 번호 붙임 (단독은 그대로)
                    if name_freq[candidate_name] > 1:
                        name_counter[candidate_name] += 1
                        # _1, _2 ... 순서대로 붙임
                        final_name = f"{candidate_name}_{name_counter[candidate_name]}"
                    else:
                        final_name = candidate_name
                        
                    item['변경할 변수명'] = final_name
                    final_data.append(item)
                    seen_raw.add(item['Raw 변수명'])

                # --- [Step 3] 매칭 실패 항목 찾기 ---
                for raw_col in df_raw.columns:
                    raw_col_str = str(raw_col).strip()
                    
                    # [수정] NO, ID 등 불필요한 컬럼은 실패 목록에서 제외
                    if raw_col_str.lower() in ['no', 'id', '번호', '순번']: continue
                    
                    if raw_col_str not in seen_raw:
                        final_data.append({
                            "Raw 변수명": raw_col_str,
                            "Code 변수명": "-",
                            "질문 내용": "-",
                            "변경할 변수명": "", 
                            "상태": "매칭 실패 (확인 필요)"
                        })
                
                st.session_state['spss_result_df'] = pd.DataFrame(final_data)
                st.session_state['spss_file_name'] = uploaded_file.name.split('.')[0]
                st.success("분석이 완료되었습니다! 아래 표에서 결과를 확인하세요.")
                
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())

# 2. 결과 확인 및 수정 에디터
if 'spss_result_df' in st.session_state:
    st.markdown("---")
    st.markdown("### 2. 결과 확인 및 수정")
    st.info("💡 **'변경할 변수명'** 컬럼을 더블클릭하여 직접 수정할 수 있습니다.")
    
    edited_df = st.data_editor(
        st.session_state['spss_result_df'],
        column_config={
            "상태": st.column_config.TextColumn("상태", disabled=True),
            "Raw 변수명": st.column_config.TextColumn(disabled=True),
            "Code 변수명": st.column_config.TextColumn(disabled=True),
            "질문 내용": st.column_config.TextColumn(disabled=True),
        },
        use_container_width=True,
        height=600,
        hide_index=True,
        key="data_editor"
    )
    
    # 3. 다운로드 버튼
    st.markdown("---")
    st.markdown("### 3. 파일 내보내기")
    
    c1, c2, c3 = st.columns(3) # 컬럼 3개로 변경
    
    with c1:
        if st.button("📥 SPSS Syntax 생성 (.sps)", key="gen_syntax_btn"):
            sps_lines = []
            sps_lines.append(f"* Auto Generated Syntax for {st.session_state['spss_file_name']}.")
            sps_lines.append(f"GET FILE='{st.session_state['spss_file_name']}.sav'.")
            sps_lines.append("RENAME VARIABLES")
            
            count = 0
            for _, row in edited_df.iterrows():
                old_v = str(row['Raw 변수명']).strip()
                new_v = str(row['변경할 변수명']).strip()
                
                if old_v and new_v and (old_v.lower() != new_v.lower()):
                    sps_lines.append(f"  ({old_v} = {new_v})")
                    count += 1
                    
            sps_lines.append(".")
            sps_lines.append("EXECUTE.")
            sps_lines.append(f"SAVE OUTFILE='{st.session_state['spss_file_name']}_Renamed.sav'.")
            sps_lines.append("EXECUTE.")
            
            final_sps = "\n".join(sps_lines)
            
            # [수정] 한글 깨짐 방지를 위해 cp949 인코딩 적용
            # cp949가 지원하지 않는 문자가 있을 경우를 대비해 errors='replace' 옵션 고려 가능하지만,
            # 변수명은 보통 영문/숫자/한글이므로 cp949로 충분합니다.
            try:
                final_sps_bytes = final_sps.encode('cp949')
            except UnicodeEncodeError:
                # cp949로 변환 안 되는 특수문자가 있는 경우 utf-8-sig로 폴백 (혹은 에러 처리)
                final_sps_bytes = final_sps.encode('utf-8-sig')
                st.warning("경고: 변수명에 한글 표준(CP949)으로 저장할 수 없는 특수문자가 포함되어 있어 UTF-8로 저장되었습니다. SPSS 버전에 따라 글자가 깨질 수 있습니다.")

            st.download_button(
                label="📄 Syntax 파일 다운로드",
                data=final_sps_bytes,
                file_name=f"{st.session_state['spss_file_name']}_Rename.sps",
                mime="text/plain"
            )
            st.success(f"총 {count}개의 변수 변환 구문이 생성되었습니다.")

    with c2:
        # [수정] 매핑 테이블을 엑셀로 변경
        out_map = io.BytesIO()
        with pd.ExcelWriter(out_map, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False)
            
        st.download_button(
            label="📄 매핑 테이블(XLSX) 다운로드",
            data=out_map.getvalue(),
            file_name=f"{st.session_state['spss_file_name']}_Mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with c3:
        # [NEW] 변환된 데이터 엑셀 다운로드 (스타일 제거: 헤더를 데이터로 처리)
        if 'spss_all_sheets' in st.session_state:
            out_data = io.BytesIO()
            
            with pd.ExcelWriter(out_data, engine='xlsxwriter') as writer:
                # 1. 변경할 이름 딕셔너리 생성
                rename_map = {}
                for _, row in edited_df.iterrows():
                    if row['변경할 변수명'] and str(row['변경할 변수명']).strip():
                        rename_map[row['Raw 변수명']] = str(row['변경할 변수명']).strip()
                
                # 2. 모든 시트 순회
                for sheet_name, df_sheet in st.session_state['spss_all_sheets'].items():
                    # 타겟 시트 확인 (DATA, LABEL, 또는 선택한 Raw 시트)
                    is_target = (sheet_name == st.session_state.get('spss_target_sheets', [''])[0]) or \
                                ('DATA' in sheet_name.upper()) or ('LABEL' in sheet_name.upper())
                    
                    if is_target:
                        # 1행: 새 변수명 (매칭된 것, 없으면 원래 이름)
                        row1 = [rename_map.get(str(col).strip(), str(col).strip()) for col in df_sheet.columns]
                        # 2행: 기존 변수명 (Original Header)
                        row2 = df_sheet.columns.tolist()
                        
                        # 데이터프레임 조립 (헤더 스타일 제거를 위해 데이터로 취급)
                        # Header DF (2줄)
                        df_header = pd.DataFrame([row1, row2]) 
                        # Data DF (Index 무시하고 값만)
                        df_body = pd.DataFrame(df_sheet.values)
                        
                        # 합치기
                        df_export = pd.concat([df_header, df_body], ignore_index=True)
                        
                        # 저장 (header=False, index=False -> 스타일 없는 순수 데이터)
                        df_export.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
                        
                    else:
                        # 타겟 아니면 원본 그대로 (단, 스타일 제거를 위해 헤더를 데이터로 내림)
                        row1 = df_sheet.columns.tolist()
                        df_header = pd.DataFrame([row1])
                        df_body = pd.DataFrame(df_sheet.values)
                        
                        df_export = pd.concat([df_header, df_body], ignore_index=True)
                        df_export.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
            
            st.download_button(
                label="📊 변환된 데이터(XLSX) 다운로드",
                data=out_data.getvalue(),
                file_name=f"{st.session_state['spss_file_name']}_Renamed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

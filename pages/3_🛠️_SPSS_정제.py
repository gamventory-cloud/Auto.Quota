import streamlit as st
import pandas as pd
import io
import collections
import traceback
import sys
import os
import re

# (주의) utils 모듈이 같은 폴더나 상위 폴더에 있어야 합니다.
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

st.set_page_config(page_title="SPSS 변수명 정제", layout="wide")

if not utils.check_password():
    st.stop()

st.header("📊 SPSS 변수명 자동 정제 & 신텍스 생성")
st.markdown("""
**Raw 데이터**와 **Code북**을 비교하여 SPSS 변수명 변경 신텍스를 생성합니다.
* **기능 1:** 라벨의 앞부분(SQ1)을 추출하여 변수명으로 자동 변환
* **기능 2:** Code북에 `문1`, `문2_1`로 표기된 변수를 `Q1`, `Q2_1`로 자동 치환하여 인식
* **기능 3:** 척도 문항 중복 시 `_1`, `_2` 자동 부여 및 순위 문항(RK) 완벽 매칭
* **기능 4:** 파생 변수(`_7`, `_etc` 등) 탐색 시 **라벨에 `[기타]` 꼬리표 자동 추가**
* **기능 5:** 엑셀 다운로드 시 **순수 데이터(디자인 없음)** + **1행: 새변수명, 2행: 기존변수명** 적용
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
            code_idx = 2 if len(sheet_names) > 2 else (1 if len(sheet_names) > 1 else 0)
            code_sheet = st.selectbox("Code북 시트", sheet_names, index=code_idx, key="code_sheet_select")
        
        # 분석 시작 버튼
        if st.button("분석 시작", key="analyze_btn"):
            with st.spinner('데이터 분석 및 매칭 중...'):
                st.session_state['spss_all_sheets'] = pd.read_excel(uploaded_file, sheet_name=None)
                st.session_state['spss_target_sheets'] = [raw_sheet]

                df_raw = st.session_state['spss_all_sheets'][raw_sheet]
                df_code = pd.read_excel(uploaded_file, sheet_name=code_sheet, header=None)
                
                raw_cols_map = {str(col).strip().lower(): str(col).strip() for col in df_raw.columns}
                temp_vars = []
                
                # --- [Step 1] Code북 순회하며 후보군 싹쓸이 ---
                for idx, row in df_code.iterrows():
                    if len(row) < 2: continue
                    if pd.isna(row.iloc[0]): continue
                    
                    col_a_val = utils.clean_text(row.iloc[0]) # 예: Q39RK1 또는 문1
                    
                    # [핵심 수정] Code명에 '문1', '문 2_1' 등으로 적혀있으면 'Q1', 'Q2_1'로 자동 치환
                    if col_a_val:
                        col_a_val = re.sub(r'^문\s*(\d)', r'Q\1', col_a_val)

                    col_c_val = utils.clean_text(row.iloc[1]) # 예: (1순위) Q14. 가장 좋아하는...
                    
                    if not col_a_val: continue
                    
                    # 1. 불필요한 라벨 텍스트 정리 (순위 등)
                    clean_label = re.sub(r'[\(\[<]?\s*\d+\s*순위\s*[\)\]>]?\s*', '', col_c_val).strip()
                    current_label_base = utils.extract_base_name(clean_label)
                    
                    # [핵심 수정] 라벨에서 뽑아낸 베이스명도 '문1' 형태라면 'Q1'로 통일
                    if current_label_base:
                        current_label_base = re.sub(r'^문\s*(\d)', r'Q\1', current_label_base)
                    else:
                        current_label_base = col_a_val 

                    is_matched = False
                    search_base_raw = col_a_val.lower() 
                    search_label_base = current_label_base

                    # [로직 1] 정확히 일치하는 경우
                    if col_a_val.lower() in raw_cols_map:
                        raw_original = raw_cols_map[col_a_val.lower()]
                        new_var_name = utils.sanitize_var_name(current_label_base)
                        
                        temp_vars.append({
                            "Raw 변수명": raw_original,
                            "Code 변수명": col_a_val,
                            "질문 내용": col_c_val,
                            "변경할 변수명": new_var_name,
                            "상태": "매칭 성공"
                        })
                        is_matched = True

                    # [로직 2] 순위 문항 탐색 (예: Code북 Q39RK1 -> Raw Q39_1)
                    if not is_matched:
                        rk_match = re.search(r'^(.*?)_?rk(\d+)$', col_a_val.lower())
                        if rk_match:
                            base_raw = rk_match.group(1)   # 예: q39
                            rank_num = rk_match.group(2)   # 예: 1
                            expected_raw_col = f"{base_raw}_{rank_num}" 
                            
                            if expected_raw_col in raw_cols_map:
                                raw_original = raw_cols_map[expected_raw_col]
                                new_var_name = utils.sanitize_var_name(f"{current_label_base}_{rank_num}")
                                
                                temp_vars.append({
                                    "Raw 변수명": raw_original,
                                    "Code 변수명": col_a_val,
                                    "질문 내용": col_c_val,
                                    "변경할 변수명": new_var_name,
                                    "상태": "매칭 성공 (순위 문항)"
                                })
                                is_matched = True
                                search_base_raw = base_raw 
                                search_label_base = current_label_base

                    # [로직 3] 기타/파생 변수 탐색 (Q39_7 등 주관식 싹쓸이)
                    prefix = search_base_raw + "_"
                    found_multiples = []
                    for rc_lower, rc_original in raw_cols_map.items():
                        if rc_lower.startswith(prefix):
                            found_multiples.append((rc_lower, rc_original))
                    
                    for rc_lower, rc_original in found_multiples:
                        suffix = rc_original[len(search_base_raw):] 
                        if not suffix.startswith('_') and not suffix.startswith('-'):
                            suffix = "_" + suffix

                        new_name = utils.sanitize_var_name(search_label_base + suffix)
                        
                        # 파생/기타 변수는 라벨 뒤에 [기타]를 명시적으로 붙여줌
                        if is_matched:
                            state_msg = "매칭 성공 (기타/파생 변수)"
                            display_label = f"{clean_label} [기타]"
                        else:
                            state_msg = "매칭 성공 (세트 문항)"
                            display_label = col_c_val
                        
                        temp_vars.append({
                            "Raw 변수명": rc_original,
                            "Code 변수명": col_a_val,
                            "질문 내용": display_label,
                            "변경할 변수명": new_name,
                            "상태": state_msg
                        })

                # --- [Step 2] 최적 매칭 선정 (중복/경합 방지) ---
                best_match_dict = {}
                for item in temp_vars:
                    raw_col = item['Raw 변수명']
                    
                    # 우선순위: 1순위(정확한 매칭/순위) > 2순위(세트) > 3순위(잡다한 파생)
                    def get_prio(s):
                        if s in ["매칭 성공", "매칭 성공 (순위 문항)"]: return 1
                        if s == "매칭 성공 (세트 문항)": return 2
                        return 3 
                    
                    if raw_col not in best_match_dict:
                        best_match_dict[raw_col] = item
                    else:
                        if get_prio(item['상태']) < get_prio(best_match_dict[raw_col]['상태']):
                            best_match_dict[raw_col] = item

                # 이름 빈도수 계산
                name_freq = collections.Counter([item['변경할 변수명'] for item in best_match_dict.values()])
                name_counter = collections.defaultdict(int)
                
                final_data = []
                
                # --- [Step 3] Raw 데이터 원본 순서대로 뷰어 구성 ---
                for raw_col in df_raw.columns:
                    raw_col_str = str(raw_col).strip()
                    
                    if raw_col_str.lower() in ['no', 'id', '번호', '순번']: continue
                    
                    if raw_col_str in best_match_dict:
                        item = best_match_dict[raw_col_str]
                        candidate_name = item['변경할 변수명']
                        
                        if name_freq[candidate_name] > 1:
                            name_counter[candidate_name] += 1
                            final_name = f"{candidate_name}_{name_counter[candidate_name]}"
                        else:
                            final_name = candidate_name
                            
                        item['변경할 변수명'] = final_name
                        final_data.append(item)
                    else:
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
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        sps_lines = []
        sps_lines.append(f'* Auto Generated Syntax for {st.session_state["spss_file_name"]}.')
        sps_lines.append(f'GET FILE="{st.session_state["spss_file_name"]}.sav".')
        sps_lines.append("RENAME VARIABLES")
        
        count = 0
        for _, row in edited_df.iterrows():
            old_v = str(row['Raw 변수명']).strip()
            new_v = str(row['변경할 변수명']).strip()
            
            if old_v and new_v and (old_v.lower() != new_v.lower()) and new_v != "nan":
                sps_lines.append(f"  ({old_v} = {new_v})")
                count += 1
                
        sps_lines.append(".")
        sps_lines.append("EXECUTE.")
        sps_lines.append(f'SAVE OUTFILE="{st.session_state["spss_file_name"]}_Renamed.sav".')
        sps_lines.append("EXECUTE.")
        
        final_sps = "\n".join(sps_lines)
        
        try:
            final_sps_bytes = final_sps.encode('cp949')
        except UnicodeEncodeError:
            final_sps_bytes = final_sps.encode('utf-8-sig')
            st.warning("⚠️ 특수문자 포함으로 인해 UTF-8로 저장되었습니다.")

        st.download_button(
            label="📄 Syntax 생성 및 다운로드 (.sps)",
            data=final_sps_bytes,
            file_name=f"{st.session_state['spss_file_name']}_Rename.sps",
            mime="text/plain",
            type="primary"
        )
        if count > 0:
            st.caption(f"✅ 총 {count}개의 변환 구문이 포함됩니다.")

    with c2:
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
        if 'spss_all_sheets' in st.session_state:
            out_data = io.BytesIO()
            
            with pd.ExcelWriter(out_data, engine='xlsxwriter') as writer:
                rename_map = {}
                for _, row in edited_df.iterrows():
                    if row['변경할 변수명'] and str(row['변경할 변수명']).strip() and str(row['변경할 변수명']).strip() != "nan":
                        rename_map[row['Raw 변수명']] = str(row['변경할 변수명']).strip()
                
                for sheet_name, df_sheet in st.session_state['spss_all_sheets'].items():
                    is_target = (sheet_name == st.session_state.get('spss_target_sheets', [''])[0]) or \
                                ('DATA' in sheet_name.upper()) or ('LABEL' in sheet_name.upper())
                    
                    if is_target:
                        row1 = [rename_map.get(str(col).strip(), str(col).strip()) for col in df_sheet.columns]
                        row2 = df_sheet.columns.tolist()
                        
                        df_header = pd.DataFrame([row1, row2]) 
                        df_body = pd.DataFrame(df_sheet.values)
                        
                        df_export = pd.concat([df_header, df_body], ignore_index=True)
                        df_export.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
                        
                    else:
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

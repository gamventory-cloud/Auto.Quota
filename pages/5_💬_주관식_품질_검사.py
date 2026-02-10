import streamlit as st
import pandas as pd
import re
import io
import sys
import os

# 1. 상위 폴더의 utils.py를 불러오기 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

st.set_page_config(page_title="주관식 품질 검사", layout="wide")

if not utils.check_password():
    st.stop()

st.title("💬 주관식 응답 품질 검사기 (Advanced)")
st.markdown("""
* **다중 시트 지원:** 엑셀 파일의 시트별 데이터 개수를 확인하고 **원하는 시트만** 합칠 수 있습니다.
* **헤더 조정:** 데이터가 몇 번째 줄부터 시작하는지 직접 지정할 수 있습니다.
""")

# ==============================================================================
# 1. 데이터 로드 및 시트 병합 설정
# ==============================================================================
st.subheader("1. 데이터 파일 업로드")
data_file = st.file_uploader("데이터 파일 업로드 (CSV, Excel, XLS)", type=['csv', 'xlsx', 'xls'])

if data_file:
    # 1. 파일 기본 정보 확인
    filename = data_file.name.lower()
    merged_df = None
    
    # 2. 엑셀 파일인 경우 시트 분석 및 옵션 제공
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        engine = 'xlrd' if filename.endswith('.xls') else 'openpyxl'
        try:
            # 모든 시트를 일단 읽음 (헤더 없이 읽어서 구조 파악)
            xls = pd.ExcelFile(data_file, engine=engine)
            sheet_names = xls.sheet_names
            
            st.info(f"📄 총 {len(sheet_names)}개의 시트가 감지되었습니다.")
            
            # --- [설정 옵션] ---
            c_opt1, c_opt2 = st.columns(2)
            with c_opt1:
                # 시트 선택 (기본: 모두 선택)
                selected_sheets = st.multiselect("합칠 시트 선택", sheet_names, default=sheet_names)
            with c_opt2:
                # 헤더 위치 지정 (기본: 0번째 줄)
                header_row_idx = st.number_input("변수명(Header)이 있는 행 번호 (0부터 시작)", min_value=0, value=0)
            
            if not selected_sheets:
                st.warning("최소 하나 이상의 시트를 선택해주세요.")
                st.stop()
                
            # --- [병합 로직] ---
            all_dfs = []
            valid_rows_log = []
            
            for sht in selected_sheets:
                # 사용자가 지정한 header 위치로 다시 읽기
                df_sht = pd.read_excel(data_file, sheet_name=sht, header=header_row_idx, engine=engine)
                
                # 데이터가 있는 경우만 처리
                if not df_sht.empty:
                    df_sht['_Origin_Sheet'] = sht # 출처 기록
                    all_dfs.append(df_sht)
                    valid_rows_log.append(f"- **{sht}**: {len(df_sht)}명")
                else:
                    valid_rows_log.append(f"- {sht}: (비어있음)")
            
            if all_dfs:
                # ignore_index=True로 인덱스 재설정 (표지 시트의 빈 공간 제거 효과)
                merged_df = pd.concat(all_dfs, ignore_index=True)
                
                # 병합 로그 출력
                with st.expander(f"📊 시트별 데이터 현황 확인 (총 {len(merged_df)}행)"):
                    st.markdown("\n".join(valid_rows_log))
                    
        except Exception as e:
            st.error(f"엑셀 읽기 오류: {e}")
            st.stop()
            
    else: # CSV 파일
        try:
            merged_df = utils.load_df(data_file)
        except Exception as e:
            st.error(f"CSV 읽기 오류: {e}")
            st.stop()

    # ==============================================================================
    # 2. 데이터 미리보기 및 컬럼 선택
    # ==============================================================================
    if merged_df is not None and not merged_df.empty:
        st.success(f"✅ 데이터 병합 완료: 총 {len(merged_df)}행 로드됨")
        
        # 미리보기 (전체 데이터프레임 모드)
        st.caption("▼ 병합된 데이터 미리보기 (상단 100개 행)")
        st.dataframe(merged_df.head(100), use_container_width=True)
        
        st.divider()
        st.subheader("2. 검사 대상 및 기준 설정")

        # 검사할 컬럼 다중 선택
        # (숫자가 아닌 컬럼만 필터링해서 보여주면 더 찾기 쉬움)
        cols = merged_df.columns.tolist()
        target_cols = st.multiselect("검사할 주관식 문항 선택 (다중 선택 가능)", cols)
        
        # 검사 옵션
        c1, c2, c3 = st.columns(3)
        with c1:
            min_len = st.number_input("최소 글자 수 (이것보다 짧으면 의심)", 1, 10, 2)
        with c2:
            check_korean_g = st.checkbox("자음/모음 남발 (예: ㅋㅋㅋ, ㅠㅠ)", value=True)
        with c3:
            check_repeat = st.checkbox("동일 문자 반복 (예: aaaa, ...)", value=True)
        
        default_bad_words = "없음, 모름, 몰라, 몰라요, 그냥, 굿, good, no, nothing, ., .., -, ?, !!"
        bad_words_input = st.text_area("🚫 거절/회피 단어 리스트 (쉼표로 구분)", value=default_bad_words)
        bad_words = [w.strip() for w in bad_words_input.split(",") if w.strip()]

        # ==============================================================================
        # 3. 분석 로직 (다중 컬럼 일괄 처리)
        # ==============================================================================
        if st.button("🔍 일괄 분석 시작", type="primary"):
            if not target_cols:
                st.warning("분석할 컬럼을 선택해주세요.")
                st.stop()

            all_bad_records = []
            progress_bar = st.progress(0)
            
            for idx, col in enumerate(target_cols):
                progress_bar.progress((idx + 1) / len(target_cols), text=f"검사 중: {col}")
                
                # 해당 컬럼 추출 (문자열 변환)
                target_series = merged_df[col].astype(str).fillna("")
                
                for row_idx, text in target_series.items():
                    detected = []
                    clean_text = text.strip()
                    
                    # 'nan', 'None' 등의 문자열 제외
                    if not clean_text or clean_text.lower() in ['nan', 'none', '']:
                        continue

                    # (1) 길이 체크
                    if len(clean_text) < min_len: detected.append("길이 미달")
                    # (2) 회피 단어
                    if clean_text in bad_words: detected.append("회피 단어")
                    # (3) 자음/모음
                    if check_korean_g and re.fullmatch(r"[ㄱ-ㅎㅏ-ㅣ\s]+", clean_text): detected.append("자음/모음 남발")
                    # (4) 반복
                    if check_repeat and re.search(r"(.)\1\1", clean_text): detected.append("문자 반복")
                    # (5) 특수문자
                    if re.fullmatch(r"[^가-힣a-zA-Z0-9]+", clean_text): detected.append("특수문자/숫자만 있음")

                    if detected:
                        record = {
                            'Index': row_idx,
                            '출처_시트': merged_df.loc[row_idx, '_Origin_Sheet'] if '_Origin_Sheet' in merged_df.columns else 'Single',
                            '대상_문항': col,
                            '응답_내용': text,
                            '의심_사유': ", ".join(detected)
                        }
                        all_bad_records.append(record)

            progress_bar.empty()

            # 결과 리포트
            st.divider()
            if all_bad_records:
                bad_df = pd.DataFrame(all_bad_records)
                
                c_res1, c_res2 = st.columns([1, 3])
                with c_res1:
                    st.error(f"🚨 총 {len(bad_df)}건 발견")
                    st.metric("발견된 불성실 응답", f"{len(bad_df)}건")
                with c_res2:
                    st.caption("문항별 발생 건수")
                    st.bar_chart(bad_df['대상_문항'].value_counts())
                
                st.dataframe(bad_df, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    bad_df.to_excel(writer, index=False)
                
                st.download_button(
                    "📥 불성실 리스트 다운로드 (xlsx)",
                    output.getvalue(),
                    "Bad_OpenEnds_Report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.success("✅ 선택한 문항들에서 불성실 패턴이 발견되지 않았습니다.")
    
    elif merged_df is not None and merged_df.empty:
        st.warning("⚠️ 데이터를 읽어왔지만 내용이 비어있습니다. '헤더 행 번호'를 조정해보세요.")

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

st.title("💬 주관식 응답 품질 검사기 (Multi-Sheet & Multi-Column)")
st.markdown("""
* **다중 시트 지원:** 엑셀 파일에 여러 시트가 있다면 자동으로 **하나로 합쳐서** 불러옵니다.
* **다중 컬럼 검사:** 여러 개의 주관식 문항을 **한 번에 선택**해서 일괄 검사합니다.
""")

# ==============================================================================
# 1. 데이터 로드 (모든 시트 통합 기능)
# ==============================================================================
data_file = st.file_uploader("데이터 파일 업로드 (CSV, Excel, XLS)", type=['csv', 'xlsx', 'xls'])

@st.cache_data(ttl=3600)
def load_data_all_sheets(file):
    """엑셀의 모든 시트를 읽어서 하나로 합치는 함수"""
    filename = file.name.lower()
    
    try:
        if filename.endswith('.csv'):
            return utils.load_df(file) # CSV는 기존 방식대로
            
        elif filename.endswith('.xlsx') or filename.endswith('.xls'):
            # 엔진 설정
            engine = 'xlrd' if filename.endswith('.xls') else 'openpyxl'
            
            # sheet_name=None이면 모든 시트를 dict 형태로 읽음 {'시트명': df, ...}
            sheets_dict = pd.read_excel(file, sheet_name=None, engine=engine)
            
            # 모든 시트 데이터프레임 리스트
            all_dfs = []
            for sheet_name, sheet_df in sheets_dict.items():
                # 데이터가 비어있지 않은 경우에만 추가
                if not sheet_df.empty:
                    # 시트 구분을 위해 'Sheet_Name' 컬럼 추가
                    sheet_df['_Origin_Sheet'] = sheet_name
                    all_dfs.append(sheet_df)
            
            if not all_dfs:
                return None
                
            # 하나로 병합 (컬럼이 달라도 합집합으로 합침)
            merged_df = pd.concat(all_dfs, ignore_index=True)
            return merged_df
            
    except Exception as e:
        st.error(f"파일 로드 중 오류 발생: {e}")
        return None
    return None

if data_file:
    df = load_data_all_sheets(data_file)
    
    if df is not None and not df.empty:
        st.success(f"데이터 로드 완료: 총 {len(df)}명 (모든 시트 통합됨)")
        
        with st.expander("데이터 미리보기"):
            st.dataframe(df.head())
        
        st.markdown("---")
        
        # 2. 검사할 컬럼 다중 선택
        target_cols = st.multiselect("검사할 주관식 문항들 (다중 선택 가능)", df.columns)
        
        # 3. 검사 옵션 설정
        st.subheader("⚙️ 검사 기준 설정")
        c1, c2, c3 = st.columns(3)
        with c1:
            min_len = st.number_input("최소 글자 수 (이것보다 짧으면 의심)", 1, 10, 2)
        with c2:
            check_korean_g = st.checkbox("자음/모음 남발 (예: ㅋㅋㅋ, ㅠㅠ)", value=True)
        with c3:
            check_repeat = st.checkbox("동일 문자 반복 (예: aaaa, ...)", value=True)
        
        # 불성실 키워드 사전
        default_bad_words = "없음, 모름, 몰라, 몰라요, 그냥, 굿, good, no, nothing, ., .., -, ?, !!"
        bad_words_input = st.text_area("🚫 거절/회피 단어 리스트 (쉼표로 구분)", value=default_bad_words)
        bad_words = [w.strip() for w in bad_words_input.split(",") if w.strip()]

        # 4. 분석 로직 (다중 컬럼 반복)
        if st.button("🔍 일괄 분석 시작", type="primary"):
            if not target_cols:
                st.warning("분석할 컬럼을 하나 이상 선택해주세요.")
                st.stop()

            # 결과 수집용 리스트
            all_bad_records = []
            
            # 진행률 표시
            progress_bar = st.progress(0)
            
            for idx, col in enumerate(target_cols):
                # 컬럼별 진행률 업데이트
                progress_bar.progress((idx + 1) / len(target_cols), text=f"검사 중: {col}")
                
                # 해당 컬럼 데이터 추출
                target_series = df[col].astype(str).fillna("")
                
                # 행 단위 검사
                for row_idx, text in target_series.items():
                    detected = []
                    clean_text = text.strip()
                    
                    # (1) 빈 값/nan 패스
                    if not clean_text or clean_text.lower() == 'nan':
                        continue

                    # (2) 길이 체크
                    if len(clean_text) < min_len:
                        detected.append("길이 미달")
                    
                    # (3) 거절/회피 단어 체크
                    if clean_text in bad_words:
                        detected.append("회피 단어")
                    
                    # (4) 자음/모음 남발
                    if check_korean_g:
                        if re.fullmatch(r"[ㄱ-ㅎㅏ-ㅣ\s]+", clean_text):
                            detected.append("자음/모음 남발")
                    
                    # (5) 동일 문자 반복
                    if check_repeat:
                        if re.search(r"(.)\1\1", clean_text):
                            detected.append("문자 반복")
                    
                    # (6) 특수문자만 있는 경우
                    if re.fullmatch(r"[^가-힣a-zA-Z0-9]+", clean_text):
                        detected.append("특수문자/숫자만 있음")

                    # 문제가 발견되면 기록
                    if detected:
                        record = {
                            'Index': row_idx,
                            '대상_문항': col,
                            '응답_내용': text,
                            '의심_사유': ", ".join(detected),
                            'Origin_Sheet': df.loc[row_idx, '_Origin_Sheet'] if '_Origin_Sheet' in df.columns else 'Single'
                        }
                        all_bad_records.append(record)

            progress_bar.empty()

            # 5. 결과 리포트
            st.divider()
            
            if all_bad_records:
                bad_df = pd.DataFrame(all_bad_records)
                
                st.error(f"🚨 총 {len(bad_df)}건의 불성실 의심 응답이 발견되었습니다!")
                
                # 문항별 발생 건수 차트
                st.caption("문항별 의심 응답 건수")
                st.bar_chart(bad_df['대상_문항'].value_counts())
                
                # 데이터프레임 표시
                st.dataframe(bad_df, use_container_width=True)
                
                # 다운로드
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    bad_df.to_excel(writer, index=False)
                
                st.download_button(
                    "📥 불성실 리스트 통합 다운로드 (xlsx)",
                    output.getvalue(),
                    "Bad_OpenEnds_All.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("✅ 선택한 모든 문항에서 불성실 응답 패턴이 발견되지 않았습니다.")
    
    elif df is None:
        pass # 에러 메시지는 load 함수에서 출력됨
    else:
        st.warning("⚠️ 데이터를 읽어왔지만 내용이 비어있습니다. 파일 내용을 확인해주세요.")

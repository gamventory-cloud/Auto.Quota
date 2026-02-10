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

st.title("💬 주관식 응답 품질 검사기")
st.markdown("""
주관식(Open-ended) 문항에서 **무성의한 응답(욕설, 자음남발, 무의미한 반복, 거절 등)**을 자동으로 탐지합니다.
""")

# 데이터 업로드
data_file = st.file_uploader("데이터 파일 업로드", type=['csv', 'xlsx'])

if data_file:
    df = utils.load_df(data_file)
    st.info(f"데이터 로드 완료: 총 {len(df)}명")
    
    st.markdown("---")
    
    # 1. 검사할 컬럼 선택
    text_col = st.selectbox("검사할 주관식 문항(Column) 선택", df.columns)
    
    # 2. 검사 옵션 설정
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

    # 3. 분석 로직
    if st.button("🔍 주관식 분석 시작", type="primary"):
        # 원본 보존
        df_res = df.copy()
        target_series = df_res[text_col].astype(str).fillna("")
        
        # 의심 사유를 담을 리스트
        reasons = []
        is_bad = []
        
        for text in target_series:
            detected = []
            clean_text = text.strip()
            
            # (1) 길이 체크
            if len(clean_text) < min_len:
                detected.append("길이 미달")
            
            # (2) 거절/회피 단어 체크
            if clean_text in bad_words:
                detected.append("회피 단어")
            
            # (3) 자음/모음 남발 (정규식)
            if check_korean_g:
                # 자음(ㄱ-ㅎ) 혹은 모음(ㅏ-ㅣ)만으로 구성된 경우
                if re.fullmatch(r"[ㄱ-ㅎㅏ-ㅣ\s]+", clean_text):
                    detected.append("자음/모음 남발")
            
            # (4) 동일 문자 반복 (3회 이상)
            if check_repeat:
                # 어떤 문자든 3번 이상 연속 (aaa, ..., 111)
                if re.search(r"(.)\1\1", clean_text):
                    detected.append("문자 반복")
            
            # (5) 특수문자만 있는 경우
            if re.fullmatch(r"[^가-힣a-zA-Z0-9]+", clean_text):
                detected.append("특수문자/숫자만 있음")

            if detected:
                is_bad.append(True)
                reasons.append(", ".join(detected))
            else:
                is_bad.append(False)
                reasons.append("통과")
        
        # 결과 컬럼 추가
        df_res['진단_결과'] = reasons
        
        # 4. 결과 보여주기
        bad_df = df_res[is_bad].copy()
        
        st.divider()
        if len(bad_df) > 0:
            st.error(f"🚨 총 {len(df)}명 중 {len(bad_df)}명의 불성실 의심 응답이 발견되었습니다!")
            
            # 비율 보여주기
            st.progress(len(bad_df) / len(df), text=f"불성실 비율: {(len(bad_df)/len(df))*100:.1f}%")
            
            # 미리보기 (중요 컬럼만)
            st.dataframe(bad_df[[text_col, '진단_결과']], use_container_width=True)
            
            # 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                bad_df.to_excel(writer, index=False)
            
            st.download_button(
                "📥 불성실 의심 리스트 다운로드",
                output.getvalue(),
                f"Bad_OpenEnds_{text_col}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("✅ 불성실한 응답 패턴이 발견되지 않았습니다.")
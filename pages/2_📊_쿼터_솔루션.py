import streamlit as st
import pandas as pd
import numpy as np
import io
import sys
import os
import random  # [추가] 랜덤 선발을 위해 필요

# 상위 폴더의 utils.py를 불러오기 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

# 페이지 설정
st.set_page_config(page_title="쿼터 솔루션", layout="wide")

# 비밀번호 잠금
if not utils.check_password():
    st.stop()

st.title("📊 쿼터(Quota) 관리 솔루션")

# 탭 구성
tab1, tab2 = st.tabs(["🎯 쿼터 맞추기 (Matching)", "📋 쿼터 현황 확인 (Checking)"])

# ==============================================================================
# [핵심 수정] 데이터 정규화 함수 (1, 1.0, "1"을 모두 "1"로 통일)
# ==============================================================================
def normalize_val(val):
    """모든 값을 문자열로 변환하고 소수점(.0) 제거 및 공백 제거"""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# ==============================================================================
# [공통 함수] Gap 계산 (정규화 적용)
# ==============================================================================
def calculate_gaps(current_df, quota_df):
    gaps = []
    for _, row in quota_df.iterrows():
        var_name = str(row['변수명']).strip()
        # [수정] 목표값 정규화
        target_val = normalize_val(row['값'])
        target_count = int(row['목표수'])
        
        if current_df.empty:
            current_count = 0
        else:
            # [수정] 현재 데이터도 정규화해서 비교
            if var_name in current_df.columns:
                # 해당 컬럼을 문자열로 변환 -> .0 제거 -> 공백 제거
                current_col_str = current_df[var_name].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                current_count = (current_col_str == target_val).sum()
            else:
                current_count = 0
            
        gap = target_count - current_count
        
        gaps.append({
            "var": var_name,
            "val": target_val, # 정규화된 값 사용
            "target": target_count,
            "current": current_count,
            "gap": gap,
            "priority": gap / target_count if target_count > 0 else 0
        })
    return pd.DataFrame(gaps)

# ==============================================================================
# [핵심 함수] 최적화 선발 로직 (정규화 적용)
# ==============================================================================
def best_fit_selection(raw_df, quota_df):
    # 1. 데이터 복사 및 ID 생성
    df_pool = raw_df.copy()
    if 'RESP_ID' not in df_pool.columns:
        df_pool['RESP_ID'] = range(len(df_pool))
        
    df_selected = pd.DataFrame(columns=raw_df.columns)
    
    # 2. 총 목표 N 계산
    if quota_df.empty: return df_selected, pd.DataFrame()
    first_var = quota_df.iloc[0]['변수명']
    total_target_n = quota_df[quota_df['변수명'] == first_var]['목표수'].sum()
    
    # UI 요소
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 3. 반복 선발 (최대 루프 제한)
    max_loops = int(total_target_n * 1.5)
    
    for i in range(max_loops):
        # (1) Gap 계산
        gap_df = calculate_gaps(df_selected, quota_df)
        
        if gap_df['gap'].max() <= 0:
            status_text.success("🎉 목표 달성 완료!")
            progress_bar.progress(1.0)
            break
            
        if df_pool.empty:
            status_text.warning("⚠️ 가용 데이터가 부족합니다.")
            break
            
        # (2) 우선순위 선정
        active_gaps = gap_df[gap_df['gap'] > 0]
        if active_gaps.empty: break
            
        # (3) 필요 집합(Needs) 파악
        needs = {}
        for _, r in active_gaps.iterrows():
            if r['var'] not in needs: needs[r['var']] = []
            needs[r['var']].append(r['val'])
            
        # (4) 최우선 타겟 선정 (Priority 1위)
        top_gap_row = active_gaps.sort_values('priority', ascending=False).iloc[0]
        target_var = top_gap_row['var']
        target_val = top_gap_row['val'] # 이미 normalize됨
        
        # (5) 후보자 필터링 (정규화 비교) [중요 수정]
        if target_var in df_pool.columns:
            # 풀 데이터를 정규화해서 비교
            pool_col_norm = df_pool[target_var].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            candidates = df_pool[pool_col_norm == target_val]
        else:
            candidates = pd.DataFrame()
        
        if candidates.empty:
            continue # 해당 조건 만족하는 사람 없으면 패스
            
        # (6) 점수 산정 (다른 쿼터 기여도)
        scores = []
        for idx, row in candidates.iterrows():
            score = 0
            for var, needed_vals in needs.items():
                if var == target_var: continue
                if var in row:
                    # 비교 시 정규화
                    val_norm = normalize_val(row[var])
                    if val_norm in needed_vals:
                        score += 1
            scores.append((idx, score))
            
        # (7) 선발 (점수 높은 순 -> 랜덤)
        scores.sort(key=lambda x: x[1], reverse=True)
        best_score = scores[0][1]
        top_candidates = [x[0] for x in scores if x[1] == best_score]
        chosen_idx = random.choice(top_candidates)
        
        # (8) 이동
        person = df_pool.loc[[chosen_idx]]
        df_selected = pd.concat([df_selected, person])
        df_pool = df_pool.drop(chosen_idx)
        
        # 진행률 업데이트
        if total_target_n > 0:
            prog = min(len(df_selected) / total_target_n, 1.0)
            progress_bar.progress(prog)
            status_text.text(f"매칭 중... ({len(df_selected)} / {total_target_n})")

    return df_selected, gap_df

# ==============================================================================
# [Tab 1] 쿼터 맞추기 UI (기존 유지)
# ==============================================================================
with tab1:
    st.header("🎯 최적 쿼터 매칭 (Best-Fit)")
    col1, col2 = st.columns(2)
    with col1:
        raw_file = st.file_uploader("1. 전체 데이터(.xlsx)", type=["xlsx", "csv"], key="match_raw")
    with col2:
        quota_file = st.file_uploader("2. 쿼터 설정표(.xlsx)", type=["xlsx", "csv"], key="match_quota")
        
    if raw_file and quota_file:
        try:
            df_raw = pd.read_excel(raw_file) if raw_file.name.endswith('xlsx') else pd.read_csv(raw_file)
            df_quota = pd.read_excel(quota_file) if quota_file.name.endswith('xlsx') else pd.read_csv(quota_file)
            
            st.info(f"데이터 로드 완료: 총 {len(df_raw)}명")
            
            if st.button("🚀 매칭 시작", type="primary"):
                with st.spinner("최적의 조합을 찾는 중..."):
                    final_df, final_gap = best_fit_selection(df_raw, df_quota)
                    
                    st.success(f"완료! {len(final_df)}명 선발됨")
                    
                    # 다운로드
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False)
                    st.download_button("📥 결과 다운로드", output.getvalue(), "Selected_Data.xlsx")
                    
                    # 결과 표
                    st.subheader("📈 달성 결과")
                    final_gap['달성률'] = (final_gap['current'] / final_gap['target'] * 100).fillna(0).round(1).astype(str) + "%"
                    st.dataframe(final_gap[['var', 'val', 'target', 'current', 'gap', '달성률']], use_container_width=True)
                        
        except Exception as e:
            st.error(f"오류: {e}")

# ==============================================================================
# [Tab 2] 쿼터 현황 확인 UI (기존 유지)
# ==============================================================================
with tab2:
    st.header("📋 쿼터 현황 점검")
    col3, col4 = st.columns(2)
    with col3:
        check_raw_file = st.file_uploader("1. 현재 데이터(.xlsx)", type=["xlsx", "csv"], key="check_raw")
    with col4:
        check_quota_file = st.file_uploader("2. 쿼터 설정표(.xlsx)", type=["xlsx", "csv"], key="check_quota")
        
    if check_raw_file and check_quota_file:
        try:
            df_check_raw = pd.read_excel(check_raw_file) if check_raw_file.name.endswith('xlsx') else pd.read_csv(check_raw_file)
            df_check_quota = pd.read_excel(check_quota_file) if check_quota_file.name.endswith('xlsx') else pd.read_csv(check_quota_file)
            
            if st.button("🔍 점검 하기"):
                # 정규화 로직이 포함된 calculate_gaps 사용
                gap_result = calculate_gaps(df_check_raw, df_check_quota)
                
                gap_result['달성률(%)'] = (gap_result['current'] / gap_result['target'] * 100).fillna(0).round(1)
                display_df = gap_result[['var', 'val', 'target', 'current', 'gap', '달성률(%)']].rename(columns={'var':'변수', 'val':'값', 'target':'목표', 'current':'현재', 'gap':'차이'})
                
                def highlight(row):
                    return ['background-color: #ffe6e6'] * len(row) if row['차이'] > 0 else ['background-color: #e6ffe6'] * len(row)

                st.dataframe(display_df.style.apply(highlight, axis=1), use_container_width=True)
                
        except Exception as e:
            st.error(f"오류: {e}")

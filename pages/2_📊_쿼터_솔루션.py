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
# [공통 함수] 데이터 정규화 및 갭 계산 (핵심 로직 개선)
# ==============================================================================

def normalize_val(val):
    """
    모든 값을 문자열로 변환하고, 엑셀에서 흔한 실수(.0) 및 공백을 제거하여 통일시킴
    예: 1 (int) -> "1", 1.0 (float) -> "1", "1.0" (str) -> "1", " 1 " -> "1"
    """
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def calculate_gaps(current_df, quota_df):
    """
    현재 데이터(current_df)와 목표(quota_df) 간의 차이(Gap)를 정밀하게 계산
    """
    gaps = []
    
    for _, row in quota_df.iterrows():
        var_name = str(row['변수명']).strip()
        # [핵심] 목표값 정규화
        target_val = normalize_val(row['값'])
        target_count = int(row['목표수'])
        
        if current_df.empty:
            current_count = 0
        else:
            # [핵심] 현재 데이터도 정규화하여 비교
            # 해당 컬럼을 문자열로 변환 -> .0 제거 -> 공백 제거
            if var_name in current_df.columns:
                current_col_str = current_df[var_name].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                current_count = (current_col_str == target_val).sum()
            else:
                current_count = 0 # 변수명이 없으면 0 처리
            
        gap = target_count - current_count
        
        gaps.append({
            "var": var_name,
            "val": target_val, # 정규화된 값 저장
            "target": target_count,
            "current": current_count,
            "gap": gap,
            # 우선순위: 남은 비율이 높을수록(달성률이 낮을수록) 높게 설정
            "priority": gap / target_count if target_count > 0 else 0 
        })
        
    return pd.DataFrame(gaps)

def best_fit_selection(raw_df, quota_df):
    """
    최적화 알고리즘: 목표 대비 가장 부족한(Gap이 큰) 그룹을 우선적으로 채우는 방식
    """
    df_pool = raw_df.copy()
    
    # 고유 ID 생성 (없으면)
    if 'RESP_ID' not in df_pool.columns:
        df_pool['RESP_ID'] = range(len(df_pool))
        
    df_selected = pd.DataFrame(columns=raw_df.columns)
    
    # 총 목표 N 계산 (첫 번째 변수의 목표 합계를 전체 N으로 가정)
    if quota_df.empty:
        return df_selected, pd.DataFrame()
        
    first_var = quota_df.iloc[0]['변수명']
    total_target_n = quota_df[quota_df['변수명'] == first_var]['목표수'].sum()
    
    # 진행 상황 표시
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 무한 루프 방지 (목표의 1.5배수까지만 반복)
    max_loops = int(total_target_n * 1.5)
    
    for i in range(max_loops):
        # 1. 현재 Gap 계산
        gap_df = calculate_gaps(df_selected, quota_df)
        
        # 종료 조건: 모든 쿼터가 충족되었으면(Gap <= 0) 종료
        if gap_df['gap'].max() <= 0:
            status_text.success("🎉 모든 쿼터 목표 달성 완료!")
            progress_bar.progress(1.0)
            break
            
        # 종료 조건: 더 이상 뽑을 사람이 없으면 종료
        if df_pool.empty:
            status_text.warning("⚠️ 가용 풀이 소진되었습니다.")
            break
            
        # 2. 우선순위 선정 (아직 덜 채운 조건들 중 Priority 높은 순)
        active_gaps = gap_df[gap_df['gap'] > 0]
        if active_gaps.empty:
            break # 이론상 위에서 걸러지지만 안전장치
            
        # 3. 필요 집합(Needs) 생성 (각 변수별로 필요한 값들 미리 파악)
        needs = {}
        for _, r in active_gaps.iterrows():
            if r['var'] not in needs: needs[r['var']] = []
            needs[r['var']].append(r['val'])
            
        # 4. 최우선 타겟 선정 (가장 급한 불 끄기)
        top_gap_row = active_gaps.sort_values('priority', ascending=False).iloc[0]
        target_var = top_gap_row['var']
        target_val = top_gap_row['val'] # 이미 normalize됨
        
        # 5. 후보자 필터링 (정규화 비교 적용)
        # 풀의 해당 컬럼을 문자열로 변환 -> .0 제거 -> 타겟값과 비교
        if target_var in df_pool.columns:
            pool_col_norm = df_pool[target_var].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            candidates_mask = (pool_col_norm == target_val)
            candidates = df_pool[candidates_mask]
        else:
            candidates = pd.DataFrame()
        
        if candidates.empty:
            # 이 조건을 만족하는 사람이 없으면 다음 루프로 (해당 조건은 포기 상태가 됨)
            # 무한 루프 방지를 위해 임시로 gap_df 조작 등이 필요할 수 있으나,
            # 여기선 우선순위가 계속 바뀌므로 자연스럽게 다른 조건을 탐색하게 둠
            continue
            
        # 6. 점수 산정 (이 사람을 뽑았을 때 다른 쿼터도 얼마나 채워주는지)
        scores = []
        for idx, row in candidates.iterrows():
            score = 0
            for var, needed_vals in needs.items():
                if var == target_var: continue # 이미 타겟 조건은 만족함
                
                # 다른 변수 값도 정규화해서 비교
                if var in row:
                    val_norm = normalize_val(row[var])
                    if val_norm in needed_vals:
                        score += 1
            scores.append((idx, score))
            
        # 7. 선발 (점수 높은 순, 동점이면 랜덤)
        scores.sort(key=lambda x: x[1], reverse=True)
        best_score = scores[0][1]
        top_candidates = [x[0] for x in scores if x[1] == best_score]
        chosen_idx = random.choice(top_candidates)
        
        # 8. 이동 (Pool -> Selected)
        person = df_pool.loc[[chosen_idx]]
        df_selected = pd.concat([df_selected, person])
        df_pool = df_pool.drop(chosen_idx)
        
        # 진행률 업데이트
        if total_target_n > 0:
            prog = min(len(df_selected) / total_target_n, 1.0)
            progress_bar.progress(prog)
            status_text.text(f"매칭 진행 중... ({len(df_selected)} / {total_target_n} 명)")

    return df_selected, gap_df

# ==============================================================================
# [Tab 1] 쿼터 맞추기 (Matching)
# ==============================================================================
with tab1:
    st.header("🎯 최적 쿼터 매칭 (Best-Fit)")
    st.markdown("전체 데이터에서 **목표 쿼터에 딱 맞는 인원**을 최적의 조합으로 선발합니다.")
    
    col1, col2 = st.columns(2)
    with col1:
        raw_file = st.file_uploader("1. 전체 응답자 데이터(.xlsx)", type=["xlsx", "csv"], key="match_raw")
    with col2:
        quota_file = st.file_uploader("2. 목표 쿼터 설정표(.xlsx)", type=["xlsx", "csv"], key="match_quota")
        
    if raw_file and quota_file:
        try:
            df_raw = pd.read_excel(raw_file) if raw_file.name.endswith('xlsx') else pd.read_csv(raw_file)
            df_quota = pd.read_excel(quota_file) if quota_file.name.endswith('xlsx') else pd.read_csv(quota_file)
            
            st.info(f"원본 데이터: {len(df_raw)}명 로드됨")
            
            if st.button("🚀 쿼터 매칭 시작", type="primary"):
                with st.spinner("알고리즘이 최적의 조합을 계산 중입니다... (1분 내외 소요)"):
                    final_df, final_gap = best_fit_selection(df_raw, df_quota)
                    
                    st.success(f"매칭 완료! 총 {len(final_df)}명 선발됨")
                    
                    # 1. 결과 다운로드
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 선발된 데이터 다운로드 (Selected_Data.xlsx)",
                        data=output.getvalue(),
                        file_name="Selected_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # 2. 결과 리포트
                    st.subheader("📈 쿼터 달성 결과")
                    
                    # 달성률 계산 및 스타일링
                    final_gap['달성률'] = (final_gap['current'] / final_gap['target'] * 100).fillna(0).round(1).astype(str) + "%"
                    
                    def style_gap(v):
                        return 'color: red; font-weight: bold;' if v > 0 else 'color: green;'
                    
                    st.dataframe(
                        final_gap[['var', 'val', 'target', 'current', 'gap', '달성률']].style.applymap(style_gap, subset=['gap']),
                        use_container_width=True,
                        height=400
                    )
                    
                    # 미달 항목 안내
                    failed = final_gap[final_gap['gap'] > 0]
                    if not failed.empty:
                        st.error(f"총 {len(failed)}개 항목에서 목표를 채우지 못했습니다.")
                    else:
                        st.balloons()
                        
        except Exception as e:
            st.error(f"오류 발생: {e}")

# ==============================================================================
# [Tab 2] 쿼터 현황 확인 (Checking)
# ==============================================================================
with tab2:
    st.header("📋 현재 쿼터 달성 현황 점검")
    st.markdown("현재 수집된 데이터가 **목표 쿼터를 얼마나 달성했는지** 확인합니다.")
    
    col3, col4 = st.columns(2)
    with col3:
        check_raw_file = st.file_uploader("1. 현재 수집 데이터(.xlsx)", type=["xlsx", "csv"], key="check_raw")
    with col4:
        check_quota_file = st.file_uploader("2. 목표 쿼터 설정표(.xlsx)", type=["xlsx", "csv"], key="check_quota")
        
    if check_raw_file and check_quota_file:
        try:
            df_check_raw = pd.read_excel(check_raw_file) if check_raw_file.name.endswith('xlsx') else pd.read_csv(check_raw_file)
            df_check_quota = pd.read_excel(check_quota_file) if check_quota_file.name.endswith('xlsx') else pd.read_csv(check_quota_file)
            
            if st.button("🔍 현황 점검"):
                # calculate_gaps 함수 재사용 (정규화 로직 포함됨)
                gap_result = calculate_gaps(df_check_raw, df_check_quota)
                
                # 달성률 계산
                gap_result['달성률(%)'] = (gap_result['current'] / gap_result['target'] * 100).fillna(0).round(1)
                
                # 보기 좋게 컬럼 정리
                display_df = gap_result[['var', 'val', 'target', 'current', 'gap', '달성률(%)']].rename(columns={
                    'var': '변수명', 'val': '값', 'target': '목표N', 'current': '현재N', 'gap': '부족분'
                })
                
                # 스타일링 (부족하면 빨강, 달성하면 초록)
                def highlight_status(row):
                    if row['부족분'] > 0:
                        return ['background-color: #ffe6e6'] * len(row) # 연한 빨강
                    else:
                        return ['background-color: #e6ffe6'] * len(row) # 연한 초록

                st.subheader("📊 점검 결과")
                st.dataframe(display_df.style.apply(highlight_status, axis=1), use_container_width=True, height=600)
                
        except Exception as e:
            st.error(f"오류 발생: {e}")

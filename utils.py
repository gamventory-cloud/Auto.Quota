import streamlit as st
import pandas as pd
import numpy as np
import re
import collections
import random
import chardet # CSV 인코딩 감지용
import io

# ==============================================================================
# 1. 비밀번호 및 보안 설정
# ==============================================================================
def check_password():
    """
    스트림릿 앱의 비밀번호를 확인합니다.
    .streamlit/secrets.toml 파일에 [secrets] password="..." 설정이 필요합니다.
    """
    if "password" not in st.secrets:
        st.error("❌ 비밀번호가 설정되지 않았습니다. (.streamlit/secrets.toml 확인 필요)")
        return False

    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    # 비밀번호 입력 폼
    with st.form("password_form"):
        password = st.text_input("비밀번호를 입력하세요", type="password")
        submitted = st.form_submit_button("로그인")
        
        if submitted:
            if password == st.secrets["password"]:
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("⛔ 비밀번호가 틀렸습니다.")
    return False


# ==============================================================================
# 2. 데이터 로딩 (CSV, XLSX, XLS 지원)
# ==============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def load_df(file):
    """
    업로드된 파일을 데이터프레임으로 변환합니다.
    지원 형식: CSV (자동 인코딩 감지), XLSX (openpyxl), XLS (xlrd)
    """
    if file is None:
        return None
        
    filename = file.name.lower()
    
    try:
        if filename.endswith('.csv'):
            # CSV는 인코딩 자동 감지 시도
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding'] if result['encoding'] else 'utf-8'
            
            # 한글 윈도우(cp949)와 호환성 높이기
            if encoding and 'EUC-KR' in encoding.upper():
                encoding = 'cp949'
                
            file.seek(0)
            return pd.read_csv(file, encoding=encoding)
            
        elif filename.endswith('.xlsx'):
            # 신형 엑셀 (openpyxl 엔진)
            return pd.read_excel(file, engine='openpyxl')
            
        elif filename.endswith('.xls'):
            # [NEW] 구형 엑셀 (xlrd 엔진) - requirements.txt에 xlrd 필수
            return pd.read_excel(file, engine='xlrd')
            
    except Exception as e:
        st.error(f"파일을 읽는 중 에러가 발생했습니다: {e}")
        return None
    return None


# ==============================================================================
# 3. 데이터 전처리 및 유틸리티 함수
# ==============================================================================
def clean_val(x):
    """데이터 값을 문자열로 변환하고 앞뒤 공백을 제거합니다."""
    if pd.isna(x): return ""
    return str(x).strip()

def natural_key(string_):
    """
    사람이 읽는 순서대로 정렬하기 위한 키 생성 (예: 1, 2, 10 순서 보장)
    """
    if not isinstance(string_, str):
        string_ = str(string_)
    return [int(s) if s.isdigit() else s for s in re.split(r'(\d+)', string_)]

def collect_values_from_cols(row, cols):
    """특정 행(row)에서 여러 컬럼(cols)의 값을 리스트로 추출"""
    vals = []
    for c in cols:
        v = row[c]
        vals.append(v if not pd.isna(v) else "")
    return vals

def sanitize_sheet_name(name):
    """엑셀 시트 이름으로 쓸 수 없는 문자 제거 및 길이 제한"""
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name[:31]  # 엑셀 시트 이름 최대 길이 31자


# ==============================================================================
# 4. 쿼터 솔루션 관련 로직 (Page 2)
# ==============================================================================
def transform_pivoted_quota(df_raw):
    """
    업로드된 엑셀 쿼터 파일(Pivot 형태)을 Flat한 형태(Target List)로 변환
    """
    # 1행, 2행이 헤더인지 데이터인지 판단하여 구조화
    # (사용자가 업로드하는 엑셀 양식에 따라 조정됨. 기본적으로 1열, 2열이 변수, 나머지가 값이라고 가정)
    
    # 간단한 구조 변환 로직 (예시)
    # 실제로는 업로드 양식에 맞춰 유연하게 짜야 함. 여기서는 3열 변수 구조 가정
    df = df_raw.fillna(method='ffill', axis=0) # 병합된 셀 처리
    
    flat_data = []
    # 가정: 0,1,2열은 쿼터 변수(성별,연령,지역), 3열은 목표값(Target)
    # 만약 헤더가 없다면:
    if len(df_raw.columns) >= 4:
        df_raw.columns = ['qt1', 'qt2', 'qt3', 'target'] + list(df_raw.columns[4:])
        return df_raw[['qt1', 'qt2', 'qt3', 'target']]
    
    return pd.DataFrame() # 실패 시 빈 DF 반환


def simulation_worker(worker_id, iters, indices, scores, m_keys, ex_keys_list, main_map, ex_maps_list, target_threshold):
    """
    [병렬 연산용] 쿼터 시뮬레이션 워커
    - indices: 응답자 인덱스 배열
    - scores: 응답자별 희소성 점수 (높을수록 귀한 사람)
    - m_keys: 메인 쿼터 키 리스트
    - ex_keys_list: 서브 쿼터 키 리스트들의 리스트
    - main_map, ex_maps_list: 쿼터 목표 맵
    """
    
    # 데이터 준비 (Numpy 변환으로 속도 향상)
    n = len(indices)
    
    best_cnt = -1
    best_idxs = []
    
    # 원본 맵 복사본 생성을 피하기 위해 루프 내에서 초기화
    # 맵의 값(Target)만 배열로 관리하면 더 빠르지만, 구조 복잡성을 위해 딕셔너리 복사 사용
    
    # 점수가 높은(희소한) 순서대로 정렬하되, 약간의 랜덤성 추가
    # 랜덤 노이즈를 추가하여 매번 다른 정렬 순서 생성
    base_order = np.argsort(scores)[::-1] # 내림차순 (점수 높은 사람 먼저)
    
    for _ in range(iters):
        # 1. 셔플: 점수 기반 + 랜덤성
        # 상위 50%는 유지하되 내부 순서 섞기 등 다양한 전략 가능
        # 여기서는 전체를 랜덤하게 섞되, 점수가 높은 구간을 우선하도록 가중치 셔플과 유사하게 처리
        
        # 방식: Score에 랜덤 노이즈를 더해서 정렬
        noise = np.random.normal(0, 0.5, n) # 표준편차 0.5 정도의 노이즈
        noisy_scores = scores + noise
        current_order = np.argsort(noisy_scores)[::-1]
        
        # 2. 그리디 매칭
        # 현재 쿼터 상황 초기화
        curr_main_map = main_map.copy()
        curr_ex_maps = [m.copy() for m in ex_maps_list]
        
        selected = []
        success_cnt = 0
        
        for idx_pos in current_order:
            real_idx = indices[idx_pos] # 실제 데이터프레임 인덱스
            # *주의: m_keys 등은 0부터 시작하는 순차 리스트이므로 idx_pos(위치)가 아니라 real_idx에 매핑된 값을 찾아야 함
            # 하지만 여기서 m_keys는 [key0, key1...] 형태이므로 real_idx가 아니라 "몇 번째 사람"인지가 중요함.
            # load_df로 불러온 df의 인덱스가 0,1,2... 순차적이라면 real_idx == i 번째 사람.
            # 안전을 위해 df.reset_index(drop=True)를 전제하거나, 리스트 인덱싱을 사용
            
            # 입력받은 m_keys 등은 리스트 형태이므로, real_idx가 아니라 "원본 데이터의 순서(0~N)"를 써야 함.
            # 따라서 indices는 [0, 1, 2 ...] 형태여야 함.
            
            # 메인 쿼터 체크
            m_key = m_keys[real_idx]
            if curr_main_map.get(m_key, 0) > 0:
                # 서브 쿼터 체크 (모두 통과해야 함)
                pass_all_extras = True
                
                # 임시 차감용 (만약 실패하면 롤백해야 하므로 체크만 먼저)
                for j, ex_key_list in enumerate(ex_keys_list):
                    # 해당 그룹에 설정된 쿼터가 없으면 패스
                    if not ex_key_list[real_idx]: 
                        continue
                        
                    # 여러 키 중 하나라도 쿼터가 남아있어야 함 (OR 조건? 보통은 AND 조건)
                    # 여기서는 1인당 1개의 속성만 갖는다고 가정하거나,
                    # 복수 속성일 경우 "하나라도 여유가 있으면 통과" or "모두 여유가 있어야 통과"
                    # 단순화를 위해: 리스트의 모든 키에 대해 쿼터 여유 확인
                    
                    # (수정) Page 2 로직상 Grid는 1:1 매핑이므로 리스트 길이는 1임.
                    # Simple 모드는 여러 속성일 수 있음.
                    # "해당 응답자가 가진 속성(키)들에 대해 쿼터가 남아있는가?"
                    
                    for k in ex_key_list[real_idx]:
                        if k in curr_ex_maps[j]: # 맵에 있는 키라면
                            if curr_ex_maps[j][k] <= 0:
                                pass_all_extras = False
                                break
                    if not pass_all_extras: break
                
                if pass_all_extras:
                    # 매칭 성공 -> 쿼터 차감
                    curr_main_map[m_key] -= 1
                    
                    for j, ex_key_list in enumerate(ex_keys_list):
                        for k in ex_key_list[real_idx]:
                            if k in curr_ex_maps[j]:
                                curr_ex_maps[j][k] -= 1
                    
                    selected.append(real_idx)
                    success_cnt += 1
        
        # 3. 결과 기록
        if success_cnt > best_cnt:
            best_cnt = success_cnt
            best_idxs = list(selected)
            
            # 목표치 달성하면 조기 종료 (옵션)
            if best_cnt >= target_threshold:
                break
                
    return best_cnt, best_idxs

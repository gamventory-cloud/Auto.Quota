import streamlit as st
import pandas as pd
import numpy as np
import re
import collections
import chardet
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
    if file is None:
        return None
        
    filename = file.name.lower()
    
    try:
        if filename.endswith('.csv'):
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding'] if result['encoding'] else 'utf-8'
            if encoding and 'EUC-KR' in encoding.upper(): encoding = 'cp949'
            file.seek(0)
            return pd.read_csv(file, encoding=encoding)
            
        elif filename.endswith('.xlsx'):
            return pd.read_excel(file, engine='openpyxl')
            
        elif filename.endswith('.xls'):
            return pd.read_excel(file, engine='xlrd')
            
    except Exception as e:
        st.error(f"파일을 읽는 중 에러가 발생했습니다: {e}")
        return None
    return None


# ==============================================================================
# 3. 데이터 전처리 및 유틸리티 함수
# ==============================================================================
def clean_val(x):
    if pd.isna(x): return ""
    return str(x).strip()

def natural_key(string_):
    if not isinstance(string_, str): string_ = str(string_)
    return [int(s) if s.isdigit() else s for s in re.split(r'(\d+)', string_)]

def collect_values_from_cols(row, cols):
    vals = []
    for c in cols:
        v = row[c]
        vals.append(v if not pd.isna(v) else "")
    return vals

def sanitize_sheet_name(name):
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for char in invalid_chars: name = name.replace(char, '_')
    return name[:31]


# ==============================================================================
# 4. 쿼터 솔루션 엔진 (Index Safe Version)
# ==============================================================================
def transform_pivoted_quota(df_raw):
    df = df_raw.fillna(method='ffill', axis=0)
    if len(df_raw.columns) >= 4:
        df_raw.columns = ['qt1', 'qt2', 'qt3', 'target'] + list(df_raw.columns[4:])
        return df_raw[['qt1', 'qt2', 'qt3', 'target']]
    return pd.DataFrame()

def simulation_worker(worker_id, iters, indices, scores, m_keys, ex_keys_list, main_map, ex_maps_list, target_threshold):
    """
    [안전장치 추가됨]
    indices: 원본 데이터프레임의 인덱스 라벨 (예: 1001, 1002...)
    m_keys, ex_keys_list: 0번부터 순서대로 정렬된 리스트
    """
    n = len(indices)
    best_cnt = -1
    best_idxs = []
    
    # 내림차순 정렬 (점수 높은 순)
    base_order = np.argsort(scores)[::-1]
    
    for _ in range(iters):
        # 1. 셔플 (노이즈 추가)
        noise = np.random.normal(0, 0.5, n)
        noisy_scores = scores + noise
        
        # current_order는 0부터 N-1까지의 '순서(Position)'를 담고 있음
        current_order = np.argsort(noisy_scores)[::-1]
        
        # 2. 그리디 매칭
        curr_main_map = main_map.copy()
        curr_ex_maps = [m.copy() for m in ex_maps_list]
        
        selected = []
        success_cnt = 0
        
        for idx_pos in current_order:
            # idx_pos: 리스트 내의 위치 (0 ~ N-1)
            
            # 메인 쿼터 키 조회 (위치 기반)
            m_key = m_keys[idx_pos]
            
            if curr_main_map.get(m_key, 0) > 0:
                pass_all_extras = True
                
                # 서브 쿼터 체크
                for j, ex_key_list in enumerate(ex_keys_list):
                    # 해당 사람의 서브 쿼터 키 조회 (위치 기반)
                    keys_for_person = ex_key_list[idx_pos]
                    
                    if not keys_for_person: continue
                    
                    for k in keys_for_person:
                        if k in curr_ex_maps[j]:
                            if curr_ex_maps[j][k] <= 0:
                                pass_all_extras = False
                                break
                    if not pass_all_extras: break
                
                if pass_all_extras:
                    # 매칭 성공 -> 차감
                    curr_main_map[m_key] -= 1
                    
                    for j, ex_key_list in enumerate(ex_keys_list):
                        keys_for_person = ex_key_list[idx_pos]
                        for k in keys_for_person:
                            if k in curr_ex_maps[j]:
                                curr_ex_maps[j][k] -= 1
                    
                    # [중요] 결과에는 원본 데이터프레임의 '라벨 인덱스'를 저장
                    selected.append(indices[idx_pos])
                    success_cnt += 1
        
        if success_cnt > best_cnt:
            best_cnt = success_cnt
            best_idxs = list(selected)
            if best_cnt >= target_threshold:
                break
                
    return best_cnt, best_idxs

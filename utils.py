# utils.py
import streamlit as st
import pandas as pd
import chardet
import io
import re
import numpy as np
import collections

# 1. 텍스트 정제 함수
def clean_text(text):
    """줄바꿈, 탭, 불필요한 공백을 제거합니다."""
    if pd.isna(text): return ""
    text = str(text).strip()
    return text.replace("\n", "").replace("\r", "").replace("\t", "")

def extract_base_name(text):
    """질문 라벨에서 마침표(.) 앞부분만 추출합니다."""
    text = clean_text(text)
    if "." in text:
        return text.split(".")[0].strip()
    return text.strip()

def sanitize_var_name(text):
    """SPSS 변수명 규칙에 맞게 특수문자를 제거합니다."""
    text = str(text)
    text = text.replace("-", "_").replace(" ", "_")
    text = re.sub(r"[^a-zA-Z0-9_]", "", text)
    text = re.sub(r"__+", "_", text)
    return text

# 2. 파일 로드 함수
def load_df(file):
    if file is None: return None
    try:
        if file.name.endswith('.csv'):
            raw = file.read(); enc = chardet.detect(raw)['encoding']
            return pd.read_csv(io.BytesIO(raw), encoding=enc if enc else 'utf-8')
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"파일 로드 실패: {e}"); return None

# 3. 쿼터/데이터 처리 관련 함수
def clean_val(v):
    if pd.isna(v): return "NaN"
    return str(v).strip().split('.')[0]

def collect_values_from_cols(row, columns):
    values = set()
    for c in columns:
        val = row[c]
        if pd.notna(val) and str(val).strip() != "":
            values.add(str(val).strip().split('.')[0])
    return sorted(list(values))

def natural_key(string_):
    target = str(string_)
    return [int(s) if s.isdigit() else s.lower() for s in re.split(r'(\d+)', target)]

def transform_pivoted_quota(df_raw):
    try:
        qt3_labels = [clean_val(x) for x in df_raw.iloc[1, 2:].dropna().values]
        data_rows = df_raw.iloc[2:].copy()
        data_rows.iloc[:, 0] = data_rows.iloc[:, 0].ffill()
        data_rows.columns = ['qt1', 'qt2'] + qt3_labels
        flat = data_rows.melt(id_vars=['qt1', 'qt2'], var_name='qt3', value_name='target')
        for col in ['qt1', 'qt2', 'qt3']: flat[col] = flat[col].apply(clean_val)
        flat['target'] = pd.to_numeric(flat['target'], errors='coerce').fillna(0).astype(int)
        return flat
    except: return None

def sanitize_sheet_name(name):
    safe_name = re.sub(r'[\\/*?:\[\]]', '_', str(name))
    if len(safe_name) > 30:
        return safe_name[:28] + ".."
    return safe_name

# 4. 비밀번호 체크 함수 (모든 페이지 상단에 붙일 것)
def check_password():
    """Returns `True` if the user had the correct password."""
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔒 접속 제한")
        st.text_input("비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("지인들만 사용 가능한 비공개 프로그램입니다.")
        return False
    else:
        return True

# 5. 시뮬레이션 워커 (쿼터용)
def simulation_worker(seed, num_iters, indices, scarcity_scores, m_keys, ex_keys_list, main_map, ex_maps, soft_target):
    np.random.seed(seed)
    local_best_cnt = 0
    local_best_idxs = []
    n_rows = len(indices)
    
    for _ in range(num_iters):
        noise = np.random.uniform(0, 0.5, size=n_rows)
        scores = scarcity_scores + noise
        sorted_arg = np.argsort(scores) 
        
        m_cnt = collections.defaultdict(int)
        ex_cnts = [collections.defaultdict(int) for _ in range(len(ex_maps))]
        curr_idx = []
        curr_c = 0
        
        for i in sorted_arg:
            mk = m_keys[i]
            limit = main_map.get(mk, 0)
            if limit > 0 and m_cnt[mk] < limit:
                all_extras_ok = True
                for j, e_map in enumerate(ex_maps):
                    if not e_map: continue 
                    keys = ex_keys_list[j][i]
                    for k in keys:
                        if k in e_map and ex_cnts[j][k] >= e_map[k]:
                            all_extras_ok = False; break
                    if not all_extras_ok: break
                
                if all_extras_ok:
                    m_cnt[mk] += 1
                    for j, e_map in enumerate(ex_maps):
                        if e_map:
                            for k in ex_keys_list[j][i]: ex_cnts[j][k] += 1
                    curr_idx.append(indices[i])
                    curr_c += 1
        
        if curr_c > local_best_cnt:
            local_best_cnt = curr_c
            local_best_idxs = list(curr_idx)
            if local_best_cnt >= soft_target: break
                
    return local_best_cnt, local_best_idxs
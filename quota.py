import streamlit as st
import pandas as pd
import chardet
import io
import time
import collections
import traceback
import numpy as np
import re
import os
import altair as alt
from joblib import Parallel, delayed, cpu_count

# ==============================================================================
# [공통 함수] 텍스트 정제 및 변수명 처리 함수 (전역 위치)
# ==============================================================================
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
    # [수정] 하이픈(-)과 공백을 먼저 언더바(_)로 치환하여 숫자 붙음 방지
    text = text.replace("-", "_").replace(" ", "_")
    # 괄호, 슬래시 등 제거 (알파벳, 숫자, 언더바만 남김)
    text = re.sub(r"[^a-zA-Z0-9_]", "", text)
    # 연속된 언더바는 하나로
    text = re.sub(r"__+", "_", text)
    return text

# [비밀번호 잠금 기능 시작] ---------------------------------------------
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 보안을 위해 비밀번호 삭제
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # 처음 접속 시 초기화
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        # 비밀번호 입력창 보여주기
        st.title("🔒 접속 제한")
        st.text_input(
            "비밀번호를 입력하세요", type="password", on_change=password_entered, key="password"
        )
        st.error("지인들만 사용 가능한 비공개 프로그램입니다.")
        return False
    else:
        # 비밀번호 맞음
        return True

if not check_password():
    st.stop()  # 비밀번호 틀리면 여기서 코드 실행 중단! (아래 내용 안 보여줌)
# [비밀번호 잠금 기능 끝] ---------------------------------------------


# 1. 페이지 설정
st.set_page_config(page_title="Quota Master Pro", layout="wide")

# 사이드바
st.sidebar.title("🧰 작업 메뉴")
app_mode = st.sidebar.radio(
    "메뉴 선택",
    ["🧹 1. 불성실 응답자 에디터", 
     "📊 2. 쿼터 자동 할당 솔루션 (Turbo)", 
     "🛠️ 3. SPSS 변수명 정제"] # 메뉴 추가됨
)
st.sidebar.markdown("---")
n_cores = cpu_count()
st.sidebar.caption(f"🖥️ CPU 코어: {n_cores}개 가동")

# --- 헬퍼 함수 ---
def load_df(file):
    if file is None: return None
    try:
        if file.name.endswith('.csv'):
            raw = file.read(); enc = chardet.detect(raw)['encoding']
            return pd.read_csv(io.BytesIO(raw), encoding=enc if enc else 'utf-8')
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"파일 로드 실패: {e}"); return None

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

# 병렬 처리 워커
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

# ================================================================================
# APP MODE 1: 불성실 에디터
# ================================================================================
if app_mode == "🧹 1. 불성실 응답자 에디터":
    st.title("🧹 불성실 응답자 제거 에디터")
    data_file = st.file_uploader("데이터 업로드", type=['csv', 'xlsx'])
    if data_file:
        df_raw = load_df(data_file)
        st.write(f"데이터: {len(df_raw)}명")
        
        if 'ed_grps' not in st.session_state: st.session_state.ed_grps = [{'cols':[]}]
        
        c1, c2 = st.columns([1,5])
        with c1: 
            if st.button("➕ 그룹추가"): st.session_state.ed_grps.append({'cols':[]}); st.rerun()
        with c2:
            if len(st.session_state.ed_grps)>1 and st.button("➖ 삭제"): st.session_state.ed_grps.pop(); st.rerun()

        c_tool1, c_tool2 = st.columns([1, 3])
        with c_tool1:
            target_idx = st.selectbox("담을 그룹", range(len(st.session_state.ed_grps)), format_func=lambda x: f"그룹 {x+1}")
            w_key_target = f"ed_ms_{target_idx}"
        with c_tool2:
            t1, t2 = st.tabs(["🔤 키워드", "↔️ 범위"])
            with t1:
                ck1, ck2 = st.columns([2,1])
                kwd = ck1.text_input("키워드", placeholder="Q1_", label_visibility="collapsed")
                if ck2.button("담기 (키워드)"):
                    if kwd:
                        found = [c for c in df_raw.columns if kwd in c]
                        cur = set(st.session_state.ed_grps[target_idx]['cols'])
                        upd = list(cur.union(set(found)))
                        upd.sort(key=lambda x: list(df_raw.columns).index(x))
                        st.session_state.ed_grps[target_idx]['cols'] = upd
                        st.session_state[w_key_target] = upd
                        st.rerun()
            with t2:
                cr1, cr2, cr3 = st.columns([1,1,1])
                cols = list(df_raw.columns)
                s_c = cr1.selectbox("Start", cols)
                e_c = cr2.selectbox("End", cols)
                if cr3.button("담기 (범위)"):
                    try:
                        si = cols.index(s_c); ei = cols.index(e_c)
                        if si<=ei:
                            rng = cols[si:ei+1]
                            cur = set(st.session_state.ed_grps[target_idx]['cols'])
                            upd = list(cur.union(set(rng)))
                            upd.sort(key=lambda x: cols.index(x))
                            st.session_state.ed_grps[target_idx]['cols'] = upd
                            st.session_state[w_key_target] = upd
                            st.rerun()
                    except: pass

        df_cln = df_raw.copy(); bad_ids = set()
        for i, g in enumerate(st.session_state.ed_grps):
            k=f"ed_ms_{i}"; 
            if k not in st.session_state: st.session_state[k]=g['cols']
            sel = st.multiselect(f"그룹 {i+1}", df_raw.columns, key=k)
            st.session_state.ed_grps[i]['cols']=sel
            if sel:
                try:
                    std = df_raw[sel].apply(pd.to_numeric, errors='coerce').std(axis=1)
                    bad = std[std==0].index.tolist()
                    if bad: st.error(f"{len(bad)}명 불성실"); bad_ids.update(bad)
                except: pass
        
        if bad_ids:
            if st.button("제거 후 다운로드"):
                final = df_cln.drop(index=list(bad_ids))
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w: final.to_excel(w, index=False)
                st.download_button("다운로드", out.getvalue(), "cleaned.xlsx")

# ================================================================================
# APP MODE 2: 쿼터 솔루션 (보존 로직 강화)
# ================================================================================
elif app_mode == "📊 2. 쿼터 자동 할당 솔루션 (Turbo)":
    st.title("📊 쿼터 자동 할당 솔루션 (Turbo + Visual)")
    
    st.subheader("1. 데이터 업로드")
    data_file = st.file_uploader("설문 데이터", type=['csv', 'xlsx'], key="quota_up")
    
    if data_file:
        df_survey = load_df(data_file)
        st.success(f"로드 완료: {len(df_survey)}명")
        st.divider()

        st.subheader("2. 쿼터 설정")
        use_main = st.checkbox("✅ 메인 쿼터 사용", value=True)
        main_map = {}; algo_main_cols = []
        
        if use_main:
            q_mode = st.radio("메인 쿼터 방식", ["엑셀 업로드", "화면 설계"], horizontal=True)
            if q_mode == "엑셀 업로드":
                qf = st.file_uploader("쿼터 파일", type=['xlsx'])
                c1,c2,c3 = st.columns(3)
                with c1: q1=st.selectbox("qt1", df_survey.columns)
                with c2: q2=st.selectbox("qt2", df_survey.columns)
                with c3: q3=st.selectbox("qt3", df_survey.columns)
                if qf:
                    algo_main_cols=[q1,q2,q3]
                    try:
                        raw = pd.read_excel(qf,0,header=None)
                        flat = transform_pivoted_quota(raw)
                        main_map = {(r.qt1, r.qt2, r.qt3): r.target for r in flat.itertuples()}
                    except: st.error("엑셀 오류")
            else:
                rv = st.multiselect("행(Row) 변수", df_survey.columns)
                cv = st.selectbox("열(Col) 변수", ["(선택)"]+list(df_survey.columns))
                if rv and cv!="(선택)":
                    algo_main_cols = rv+[cv]
                    base = df_survey.copy()
                    for c in algo_main_cols:
                        base[c]=base[c].apply(clean_val)
                        uv=sorted(base[c].unique(), key=natural_key)
                        base[c]=pd.Categorical(base[c], categories=uv, ordered=True)
                    pi = base.groupby(algo_main_cols, observed=False).size().unstack(fill_value=0)
                    ed = st.data_editor(pi.reset_index(), use_container_width=True, disabled=rv)
                    mlt = ed.melt(id_vars=rv, var_name=cv, value_name='target')
                    for _,r in mlt.iterrows():
                        try:
                            t=int(r['target'])
                            if t>0: main_map[tuple(str(r[c]) for c in algo_main_cols)]=t
                        except: pass
        else:
            main_map = {('All',): st.number_input("전체 목표", 1, 10000, 1000)}; algo_main_cols=[]

        ex_configs = []
        tabs = st.tabs(["추가 1", "추가 2", "추가 3", "추가 4"])
        
        for i, tab in enumerate(tabs):
            with tab:
                ex_mode = st.radio(f"설정 방식 (그룹 {i+1})", ["단순형 (변수 값별 할당)", "조합형 (행/열 교차 할당)"], key=f"ex_mode_{i}", horizontal=True)
                
                config = {'cols': [], 'map': {}, 'name': f"Extra_{i+1}", 'mode': 'simple'}
                
                if ex_mode.startswith("단순형"):
                    config['mode'] = 'simple'
                    cols = st.multiselect(f"변수 선택 (그룹 {i+1})", df_survey.columns, key=f"ms{i}")
                    if cols:
                        config['cols'] = cols
                        auto_name = "_".join([str(c) for c in cols])
                        config['name'] = sanitize_sheet_name(auto_name)
                        
                        vals = []
                        for _, r in df_survey[cols].fillna("").iterrows(): vals.extend(collect_values_from_cols(r, cols))
                        cnt = pd.DataFrame.from_dict(collections.Counter(vals), orient='index', columns=['현재']).reset_index()
                        cnt.columns=['값','현재']; cnt['목표']=cnt['현재']
                        cnt['srt']=cnt['값'].apply(natural_key)
                        ed = st.data_editor(cnt.sort_values('srt').drop(columns=['srt']), use_container_width=True, key=f"ed{i}")
                        for _,r in ed.iterrows(): 
                            if r['목표']>0: config['map'][str(r['값'])]=int(r['목표'])
                
                else:
                    config['mode'] = 'grid'
                    st.caption("메인 쿼터처럼 행과 열을 교차하여 상세 목표를 설정합니다.")
                    ex_rv = st.multiselect(f"행(Row) 변수 (그룹 {i+1})", df_survey.columns, key=f"ex_rv_{i}")
                    ex_cv = st.selectbox(f"열(Col) 변수 (그룹 {i+1})", ["(선택)"]+list(df_survey.columns), key=f"ex_cv_{i}")
                    
                    if ex_rv and ex_cv != "(선택)":
                        target_cols = ex_rv + [ex_cv]
                        config['cols'] = target_cols
                        auto_name = "_".join([str(c) for c in target_cols])
                        config['name'] = sanitize_sheet_name(auto_name)
                        
                        base = df_survey.copy()
                        for c in target_cols:
                            base[c] = base[c].apply(clean_val)
                            uv = sorted(base[c].unique(), key=natural_key)
                            base[c] = pd.Categorical(base[c], categories=uv, ordered=True)
                        
                        pi = base.groupby(target_cols, observed=False).size().unstack(fill_value=0)
                        ed = st.data_editor(pi.reset_index(), use_container_width=True, disabled=ex_rv, key=f"ex_ed_grid_{i}")
                        
                        mlt = ed.melt(id_vars=ex_rv, var_name=ex_cv, value_name='target')
                        for _, r in mlt.iterrows():
                            try:
                                t = int(r['target'])
                                if t > 0:
                                    key_tuple = tuple(str(r[c]) for c in target_cols)
                                    config['map'][key_tuple] = t
                            except: pass

                ex_configs.append(config)

        st.divider()
        st.subheader("3. 실행 옵션")
        c1, c2 = st.columns(2)
        with c1:
            c_no = st.selectbox("ID 컬럼", df_survey.columns)
            tol = st.number_input("허용 오차", 0, 100, 0)
        with c2:
            iters = st.number_input("시도 횟수", 100, 1000000, 10000, 1000)
            use_intval = st.checkbox("intval 최적화", value=True)
            c_int = st.selectbox("intval 컬럼", df_survey.columns) if use_intval else None

        if st.button("🚀 매칭 시작 (Turbo)", type="primary"):
            if not main_map: st.error("목표 없음"); st.stop()
            
            try:
                with st.spinner("종합 희소성 계산 및 병렬 연산 중..."):
                    df_proc = df_survey.copy()
                    if use_main:
                        for c in algo_main_cols: df_proc[c] = df_proc[c].apply(clean_val)
                        m_keys = list(zip(*[df_proc[c] for c in algo_main_cols]))
                    else: m_keys = [('All',) for _ in range(len(df_proc))]

                    ex_keys_list = []
                    for cfg in ex_configs:
                        if not cfg['cols']:
                            ex_keys_list.append([[] for _ in range(len(df_proc))])
                            continue
                            
                        if cfg['mode'] == 'simple':
                            keys = df_proc.apply(lambda r: collect_values_from_cols(r, cfg['cols']), axis=1).tolist()
                        else:
                            for c in cfg['cols']: df_proc[c] = df_proc[c].apply(clean_val)
                            tuples = list(zip(*[df_proc[c] for c in cfg['cols']]))
                            keys = [[t] for t in tuples]
                        ex_keys_list.append(keys)

                    target_total = sum(main_map.values())
                    soft_target = target_total - tol
                    
                    # Score Calculation
                    m_cnt = collections.Counter(m_keys)
                    if use_main:
                        score_main = np.array([m_cnt.get(k,0)/main_map.get(k,1) if main_map.get(k,0)>0 else 999 for k in m_keys])
                    else:
                        score_main = np.ones(len(df_proc))

                    score_extras = np.zeros(len(df_proc))
                    for j, cfg in enumerate(ex_configs):
                        if not cfg['cols']: continue
                        all_vals = []
                        for keys in ex_keys_list[j]: all_vals.extend(keys)
                        ex_cnt_total = collections.Counter(all_vals)
                        row_scores = []
                        ex_map = cfg['map']
                        for keys in ex_keys_list[j]:
                            if not keys: row_scores.append(1.0); continue
                            s_vals = []
                            for k in keys:
                                if k in ex_map and ex_map[k] > 0: s_vals.append(ex_cnt_total[k] / ex_map[k])
                                else: s_vals.append(999)
                            row_scores.append(min(s_vals))
                        score_extras += np.array(row_scores)
                    
                    final_scarcity_scores = score_main + score_extras
                    
                    # Parallel
                    ipc = max(1, iters // n_cores)
                    res = Parallel(n_jobs=-1, backend="threading")(delayed(simulation_worker)(
                        i, ipc, df_proc.index.to_numpy(), final_scarcity_scores, m_keys, ex_keys_list, main_map, [c['map'] for c in ex_configs], soft_target
                    ) for i in range(n_cores))
                    
                    g_best_cnt = 0; g_best_idxs = []
                    for c, ixs in res:
                        if c > g_best_cnt: g_best_cnt=c; g_best_idxs=ixs

                is_fail = g_best_cnt < soft_target
                
                # -------------------------------------------------------------
                # 엑셀 데이터 및 분석 준비
                # -------------------------------------------------------------
                fin_idxs = list(g_best_idxs)
                m_keys_map = {idx: k for idx, k in zip(df_proc.index, m_keys)}
                ex_keys_maps = [{idx: k for idx, k in zip(df_proc.index, k_list)} for k_list in ex_keys_list]
                
                final_m = collections.Counter()
                final_exs = [collections.Counter() for _ in range(len(ex_configs))]
                clean_fin_idxs = [int(idx) for idx in fin_idxs]
                
                for idx in clean_fin_idxs:
                    final_m[m_keys_map[idx]] += 1
                    for j, cfg in enumerate(ex_configs):
                        if cfg['cols']:
                            for k in ex_keys_maps[j][idx]: final_exs[j][k] += 1

                recs = []
                # 부족분 분석 (엑셀용)
                if is_fail:
                    if use_main:
                        for k, tgt in main_map.items():
                            act = final_m.get(k, 0); diff = tgt - act
                            if diff > 0: 
                                raw_avail = m_cnt.get(k, 0)
                                reason = "⚠️ 물리적 부족" if raw_avail < tgt else "⚔️ 경합 부족"
                                recs.append({'순서': 0, '구분': '메인 쿼터', '항목': " / ".join(k), '목표': tgt, '현재': act, '부족': diff, '진단': reason, '전체보유': raw_avail})
                    
                    for j, cfg in enumerate(ex_configs):
                        if cfg['cols']:
                            all_vals_raw = []
                            for keys in ex_keys_list[j]: all_vals_raw.extend(keys)
                            raw_cnt_map = collections.Counter(all_vals_raw)
                            for k, tgt in cfg['map'].items():
                                act = final_exs[j].get(k, 0); diff = tgt - act
                                if diff > 0: 
                                    raw_avail = raw_cnt_map.get(k, 0)
                                    reason = "⚠️ 물리적 부족" if raw_avail < tgt else "⚔️ 경합 부족"
                                    display_item = " / ".join(k) if isinstance(k, tuple) else k
                                    recs.append({'순서': j+1, '구분': cfg['name'], '항목': display_item, '목표': tgt, '현재': act, '부족': diff, '진단': reason, '전체보유': raw_avail})

                # [중요 변경] 엑셀 데이터 생성 시 정렬 기준 변경
                df_survey['Chk'] = "제외"
                df_survey.loc[clean_fin_idxs, 'Chk'] = "통과"
                
                # 시트1: Result_All (전체 데이터)
                # 오해 방지를 위해 'ID' 컬럼 기준으로만 정렬합니다. (통과/제외가 섞여서 나옴 -> 삭제 안 된 것 확인 가능)
                df_all = df_survey.sort_values(by=c_no, ascending=True)
                
                # 시트2: Result_Pass (통과 데이터만)
                # '통과'인 행만 뽑아서 별도로 저장
                df_pass = df_survey[df_survey['Chk'] == "통과"].sort_values(c_no, ascending=True)
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                    # 전체 데이터 (섞여있음)
                    df_all.to_excel(w, index=False, sheet_name='Result_All')
                    # 통과 데이터 (깔끔함)
                    df_pass.to_excel(w, index=False, sheet_name='Result_Pass')
                    
                    if recs: 
                        df_excel = pd.DataFrame(recs)
                        df_excel['sort_val'] = df_excel['항목'].apply(lambda x: tuple(natural_key(x)))
                        df_excel = df_excel.sort_values(by=['순서', 'sort_val'], ascending=[True, True])
                        df_excel.drop(columns=['순서', 'sort_val']).to_excel(w, index=False, sheet_name='Shortage_Analysis')
                    
                    if use_main:
                            pd.DataFrame([{'G':str(k), 'T':v, 'A':final_m[k]} for k,v in main_map.items()]).to_excel(w, sheet_name='Main_Status')

                    for j, cfg in enumerate(ex_configs):
                        if cfg['cols']:
                            data_e = []
                            for k, t in cfg['map'].items():
                                k_str = " / ".join(k) if isinstance(k, tuple) else k
                                data_e.append({'Value': k_str, 'Target': t, 'Actual': final_exs[j][k], 'Diff': t - final_exs[j][k]})
                            pd.DataFrame(data_e).sort_values('Value', key=lambda c: c.map(natural_key)).to_excel(w, sheet_name=cfg['name'], index=False)
                
                # -------------------------------------------------------------
                # 다운로드 버튼 및 검증 메시지
                # -------------------------------------------------------------
                st.divider()
                st.subheader("📊 할당 결과 시각화")
                
                # [NEW] 데이터 검증 메시지
                total_rows = len(df_survey)
                pass_rows = len(df_pass)
                exclude_rows = total_rows - pass_rows
                st.info(f"💾 **데이터 저장 완료**: 총 **{total_rows:,}명** (통과 {pass_rows:,}명 + 제외 {exclude_rows:,}명)이 엑셀에 모두 저장되었습니다.")

                # 다운로드 버튼
                btn_label = "📥 결과 파일 다운로드 (Result.xlsx)" if not is_fail else "⚠️ 실패한 결과라도 다운로드"
                st.download_button(btn_label, out.getvalue(), "result.xlsx", type="primary", use_container_width=True)
                
                # 상단 메트릭
                rate = (g_best_cnt / target_total) * 100
                c1, c2, c3 = st.columns(3)
                c1.metric("📌 전체 목표", f"{target_total:,}명")
                c2.metric("✅ 매칭 성공", f"{g_best_cnt:,}명")
                delta_color = "normal" if not is_fail else "inverse"
                c3.metric("📈 달성률", f"{rate:.1f}%", delta=f"{g_best_cnt - target_total}명" if is_fail else "목표 달성", delta_color=delta_color)

                if is_fail:
                    st.error("⚠️ 목표 인원을 달성하지 못했습니다. 아래 분석 결과를 확인하세요.")
                else:
                    st.success("🎉 목표 인원을 모두 달성했습니다!")
                
                st.markdown("### 🔍 쿼터별 상세 현황")
                
                active_ex_cfgs = [(j, cfg) for j, cfg in enumerate(ex_configs) if cfg['cols']]
                v_tabs = st.tabs(["메인 쿼터"] + [cfg['name'] for _, cfg in active_ex_cfgs])
                
                with v_tabs[0]:
                    if use_main:
                        data_m = []
                        for k, tgt in main_map.items():
                            k_str = " / ".join(k)
                            act = final_m[k]
                            data_m.append({'Label': k_str, 'Type': '1.목표', 'Value': tgt})
                            data_m.append({'Label': k_str, 'Type': '2.달성', 'Value': act})
                        
                        if data_m:
                            df_chart_m = pd.DataFrame(data_m)
                            df_chart_m['sort_val'] = df_chart_m['Label'].apply(lambda x: tuple(natural_key(x)))
                            df_chart_m = df_chart_m.sort_values('sort_val')
                            sorted_labels = df_chart_m['Label'].unique().tolist()
                            
                            chart_data = df_chart_m.drop(columns=['sort_val'])
                            chart = alt.Chart(chart_data).mark_bar().encode(
                                y=alt.Y('Label:N', axis=alt.Axis(title=None), sort=sorted_labels),
                                x=alt.X('Value:Q', axis=alt.Axis(title='인원수')),
                                color=alt.Color('Type:N', scale=alt.Scale(domain=['1.목표', '2.달성'], range=['#e0e0e0', '#4c78a8']), legend=alt.Legend(title="구분")),
                                yOffset='Type:N'
                            ).properties(height=max(300, len(main_map)*25))
                            st.altair_chart(chart, use_container_width=True)
                    else:
                        st.info("메인 쿼터 설정이 없습니다.")

                for idx, (j, cfg) in enumerate(active_ex_cfgs):
                    with v_tabs[idx + 1]:
                        data_e = []
                        for k, tgt in cfg['map'].items():
                            k_str = " / ".join(k) if isinstance(k, tuple) else k
                            act = final_exs[j][k]
                            data_e.append({'Label': k_str, 'Type': '1.목표', 'Value': tgt})
                            data_e.append({'Label': k_str, 'Type': '2.달성', 'Value': act})
                        
                        if data_e:
                            df_chart_e = pd.DataFrame(data_e)
                            df_chart_e['sort_val'] = df_chart_e['Label'].apply(lambda x: tuple(natural_key(x)))
                            df_chart_e = df_chart_e.sort_values('sort_val')
                            sorted_labels_e = df_chart_e['Label'].unique().tolist()
                            
                            chart_data_e = df_chart_e.drop(columns=['sort_val'])
                            chart = alt.Chart(chart_data_e).mark_bar().encode(
                                y=alt.Y('Label:N', axis=alt.Axis(title=None), sort=sorted_labels_e),
                                x=alt.X('Value:Q', axis=alt.Axis(title='인원수')),
                                color=alt.Color('Type:N', scale=alt.Scale(domain=['1.목표', '2.달성'], range=['#e0e0e0', '#4c78a8']), legend=alt.Legend(title="구분")),
                                yOffset='Type:N'
                            ).properties(height=max(300, len(cfg['map'])*25))
                            st.altair_chart(chart, use_container_width=True)
                
                if recs:
                    st.divider()
                    st.subheader("📉 부족 쿼터 분석 및 진단")
                    df_recs = pd.DataFrame(recs)
                    df_recs['sort_val'] = df_recs['항목'].apply(lambda x: tuple(natural_key(x)))
                    df_recs = df_recs.sort_values(by=['순서', 'sort_val'], ascending=[True, True])
                    st.dataframe(df_recs.drop(columns=['순서', 'sort_val']), use_container_width=True, hide_index=True)

            except Exception as e: st.error("오류 발생"); st.code(traceback.format_exc())

# ==============================================================================
# APP MODE 3: SPSS 변수명 정제 (수정됨: 중복 변수명에 자동 번호 부여)
# ==============================================================================
elif app_mode == "🛠️ 3. SPSS 변수명 정제":
    st.header("📊 SPSS 변수명 자동 정제 & 신텍스 생성")
    st.markdown("""
    **Raw 데이터**와 **Code북**을 비교하여 SPSS 변수명 변경 신텍스를 생성합니다.
    * **Code북 규칙:** 1열=변수명(Q1), **2열=질문라벨(SQ1. 성별...)**
    * **기능 1:** 라벨의 앞부분(SQ1)을 추출하여 변수명으로 자동 변환
    * **기능 2:** 척도 문항 등으로 변수명이 중복될 경우, 자동으로 `_1`, `_2`, `_3`을 붙여서 구분
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
                    # 데이터프레임 로드
                    df_raw = pd.read_excel(uploaded_file, sheet_name=raw_sheet)
                    # [수정] header=None 옵션 추가: 첫 번째 줄(Q1)도 데이터로 읽기 위해
                    df_code = pd.read_excel(uploaded_file, sheet_name=code_sheet, header=None)
                    
                    # Raw 데이터 컬럼 매핑 (소문자 -> 원본)
                    raw_cols_map = {str(col).strip().lower(): str(col).strip() for col in df_raw.columns}
                    
                    temp_vars = []
                    
                    # --- [Step 1] Code북 순회 (무조건 1, 2열 사용) ---
                    for idx, row in df_code.iterrows():
                        if len(row) < 2: continue
                        if pd.isna(row.iloc[0]): continue
                        
                        col_a_val = clean_text(row.iloc[0]) # 변수명 (Code) - 예: Q1
                        col_c_val = clean_text(row.iloc[1]) # 질문 라벨 - 예: SQ1. 성별
                        
                        if not col_a_val: continue
                        
                        # [핵심] 라벨에서 기본 이름 추출 (예: "SQ1. 성별" -> "SQ1")
                        label_base = extract_base_name(col_c_val)
                        if not label_base: 
                            label_base = col_a_val # 실패 시 Code명 사용

                        # [스마트 매칭 로직]
                        # 1. 정확히 일치하는 경우
                        if col_a_val.lower() in raw_cols_map:
                            raw_original = raw_cols_map[col_a_val.lower()]
                            new_var_name = sanitize_var_name(label_base)
                            
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
                            new_name = sanitize_var_name(label_base + suffix)
                            
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
        
        c1, c2 = st.columns(2)
        
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
                
                st.download_button(
                    label="📄 Syntax 파일 다운로드",
                    data=final_sps,
                    file_name=f"{st.session_state['spss_file_name']}_Rename.sps",
                    mime="text/plain"
                )
                st.success(f"총 {count}개의 변수 변환 구문이 생성되었습니다.")

        with c2:
            csv_buffer = io.BytesIO()
            edited_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
            st.download_button(
                label="📄 매핑 테이블(CSV) 다운로드",
                data=csv_buffer,
                file_name=f"{st.session_state['spss_file_name']}_Mapping.csv",
                mime="text/csv"
            )

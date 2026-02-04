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

# 1. 페이지 설정
st.set_page_config(page_title="Quota Master Pro", layout="wide")

# 사이드바
st.sidebar.title("🧰 작업 메뉴")
app_mode = st.sidebar.radio(
    "메뉴 선택",
    ["🧹 1. 불성실 응답자 에디터", "📊 2. 쿼터 자동 할당 솔루션 (Turbo)"]
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
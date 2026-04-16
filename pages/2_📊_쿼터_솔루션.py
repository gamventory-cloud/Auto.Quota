import streamlit as st
import pandas as pd
import io
import collections
import numpy as np
import altair as alt
from joblib import Parallel, delayed, cpu_count
import sys
import os
import traceback

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

st.set_page_config(page_title="쿼터 솔루션", layout="wide")

if not utils.check_password():
    st.stop()

st.title("📊 쿼터 자동 할당 솔루션 (Turbo + Visual + Robust)")
n_cores = cpu_count()
st.sidebar.caption(f"🖥️ CPU 코어: {n_cores}개 가동")

# ==============================================================================
# [데이터 정규화 함수] (1, 1.0, "1" 통일)
# ==============================================================================
def normalize_val(val):
    """모든 값을 문자열로 변환하고 소수점(.0) 제거 및 공백 제거"""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# ==============================================================================
# 1. 데이터 업로드
# ==============================================================================
st.subheader("1. 데이터 업로드")
data_file = st.file_uploader("설문 데이터", type=['csv', 'xlsx'], key="quota_up")

if data_file:
    df_survey = utils.load_df(data_file)
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
                    flat = utils.transform_pivoted_quota(raw)
                    # 쿼터 키 정규화
                    main_map = {
                        tuple(normalize_val(v) for v in (r.qt1, r.qt2, r.qt3)): r.target 
                        for r in flat.itertuples()
                    }
                except: st.error("엑셀 오류")
        else:
            rv = st.multiselect("행(Row) 변수", df_survey.columns)
            cv = st.selectbox("열(Col) 변수", ["(선택)"]+list(df_survey.columns))
            if rv and cv!="(선택)":
                algo_main_cols = rv+[cv]
                base = df_survey.copy()
                for c in algo_main_cols:
                    base[c]=base[c].apply(utils.clean_val)
                    uv=sorted(base[c].unique(), key=utils.natural_key)
                    base[c]=pd.Categorical(base[c], categories=uv, ordered=True)
                pi = base.groupby(algo_main_cols, observed=False).size().unstack(fill_value=0)
                ed = st.data_editor(pi.reset_index(), use_container_width=True, disabled=rv)
                mlt = ed.melt(id_vars=rv, var_name=cv, value_name='target')
                for _,r in mlt.iterrows():
                    try:
                        t=int(r['target'])
                        if t>0: 
                            # 키 정규화
                            key = tuple(normalize_val(r[c]) for c in algo_main_cols)
                            main_map[key]=t
                    except: pass
    else:
        main_map = {('All',): st.number_input("전체 목표", 1, 10000, 1000)}; algo_main_cols=[]

    ex_configs = []
    
    # 추가 쿼터 6개로 확장
    tabs = st.tabs(["추가 1", "추가 2", "추가 3", "추가 4", "추가 5", "추가 6"])
    
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
                    config['name'] = utils.sanitize_sheet_name(auto_name)
                    
                    vals = []
                    for _, r in df_survey[cols].fillna("").iterrows(): 
                        raw_vals = utils.collect_values_from_cols(r, cols)
                        norm_vals = [normalize_val(v) for v in raw_vals]
                        vals.extend(norm_vals)

                    cnt = pd.DataFrame.from_dict(collections.Counter(vals), orient='index', columns=['현재']).reset_index()
                    cnt.columns=['값','현재']; cnt['목표']=cnt['현재']
                    cnt['srt']=cnt['값'].apply(utils.natural_key)
                    ed = st.data_editor(cnt.sort_values('srt').drop(columns=['srt']), use_container_width=True, key=f"ed{i}")
                    for _,r in ed.iterrows(): 
                        if r['목표']>0: 
                            config['map'][normalize_val(r['값'])]=int(r['목표'])
            
            else:
                config['mode'] = 'grid'
                st.caption("메인 쿼터처럼 행과 열을 교차하여 상세 목표를 설정합니다.")
                ex_rv = st.multiselect(f"행(Row) 변수 (그룹 {i+1})", df_survey.columns, key=f"ex_rv_{i}")
                ex_cv = st.selectbox(f"열(Col) 변수 (그룹 {i+1})", ["(선택)"]+list(df_survey.columns), key=f"ex_cv_{i}")
                
                if ex_rv and ex_cv != "(선택)":
                    target_cols = ex_rv + [ex_cv]
                    config['cols'] = target_cols
                    auto_name = "_".join([str(c) for c in target_cols])
                    config['name'] = utils.sanitize_sheet_name(auto_name)
                    
                    base = df_survey.copy()
                    for c in target_cols:
                        base[c] = base[c].apply(utils.clean_val)
                        uv = sorted(base[c].unique(), key=utils.natural_key)
                        base[c] = pd.Categorical(base[c], categories=uv, ordered=True)
                    
                    pi = base.groupby(target_cols, observed=False).size().unstack(fill_value=0)
                    ed = st.data_editor(pi.reset_index(), use_container_width=True, disabled=ex_rv, key=f"ex_ed_grid_{i}")
                    
                    mlt = ed.melt(id_vars=ex_rv, var_name=ex_cv, value_name='target')
                    for _, r in mlt.iterrows():
                        try:
                            t = int(r['target'])
                            if t > 0:
                                key_tuple = tuple(normalize_val(r[c]) for c in target_cols)
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
                
                # [버그 수정] utils.clean_val을 유지하면서 안전하게 normalize_val 적용
                if use_main:
                    for c in algo_main_cols: 
                        df_proc[c] = df_proc[c].apply(utils.clean_val).apply(normalize_val)
                    m_keys = list(zip(*[df_proc[c] for c in algo_main_cols]))
                else: 
                    m_keys = [('All',) for _ in range(len(df_proc))]

                ex_keys_list = []
                for cfg in ex_configs:
                    if not cfg['cols']:
                        ex_keys_list.append([[] for _ in range(len(df_proc))])
                        continue
                        
                    if cfg['mode'] == 'simple':
                        # [버그 수정] utils.collect_values_from_cols 복구 (다중응답 쪼개기 기능 등 유지)
                        keys = df_proc.apply(
                            lambda r: [normalize_val(v) for v in utils.collect_values_from_cols(r, cfg['cols'])], 
                            axis=1
                        ).tolist()
                    else:
                        for c in cfg['cols']: 
                            df_proc[c] = df_proc[c].apply(utils.clean_val).apply(normalize_val)
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
                res = Parallel(n_jobs=-1, backend="threading")(delayed(utils.simulation_worker)(
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

            # 엑셀 저장
            df_survey['Chk'] = "제외"
            df_survey.loc[clean_fin_idxs, 'Chk'] = "통과"
            
            df_all = df_survey.sort_values(by=c_no, ascending=True)
            df_pass = df_survey[df_survey['Chk'] == "통과"].sort_values(c_no, ascending=True)
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                df_all.to_excel(w, index=False, sheet_name='Result_All')
                df_pass.to_excel(w, index=False, sheet_name='Result_Pass')
                
                if recs: 
                    df_excel = pd.DataFrame(recs)
                    df_excel['sort_val'] = df_excel['항목'].apply(lambda x: tuple(utils.natural_key(x)))
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
                        pd.DataFrame(data_e).sort_values('Value', key=lambda c: c.map(utils.natural_key)).to_excel(w, sheet_name=cfg['name'], index=False)
            
            # -------------------------------------------------------------
            # 다운로드 버튼 및 검증 메시지
            # -------------------------------------------------------------
            st.divider()
            st.subheader("📊 할당 결과 시각화")
            
            total_rows = len(df_survey)
            pass_rows = len(df_pass)
            exclude_rows = total_rows - pass_rows
            st.info(f"💾 **데이터 저장 완료**: 총 **{total_rows:,}명** (통과 {pass_rows:,}명 + 제외 {exclude_rows:,}명)이 엑셀에 모두 저장되었습니다.")

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
                        df_chart_m['sort_val'] = df_chart_m['Label'].apply(lambda x: tuple(utils.natural_key(x)))
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
                        df_chart_e['sort_val'] = df_chart_e['Label'].apply(lambda x: tuple(utils.natural_key(x)))
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
                df_recs['sort_val'] = df_recs['항목'].apply(lambda x: tuple(utils.natural_key(x)))
                df_recs = df_recs.sort_values(by=['순서', 'sort_val'], ascending=[True, True])
                st.dataframe(df_recs.drop(columns=['순서', 'sort_val']), use_container_width=True, hide_index=True)

        except Exception as e: st.error("오류 발생"); st.code(traceback.format_exc())

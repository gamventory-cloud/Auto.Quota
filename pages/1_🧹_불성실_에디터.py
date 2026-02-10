import streamlit as st
import pandas as pd
import io
import sys
import os

# 상위 폴더의 utils를 불러오기 위한 경로 설정
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils

st.set_page_config(page_title="불성실 에디터", layout="wide")

if not utils.check_password():
    st.stop()

st.title("🧹 불성실 응답자 제거 에디터")
data_file = st.file_uploader("데이터 업로드", type=['csv', 'xlsx'])

if data_file:
    df_raw = utils.load_df(data_file)
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
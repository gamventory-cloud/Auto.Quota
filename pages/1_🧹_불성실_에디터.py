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
data_file = st.file_uploader("데이터 업로드", type=['csv', 'xlsx', 'xls'])

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
    
    # 검사 옵션
    st.markdown("---")
    st.subheader("🔍 검사 옵션")
    check_method = st.radio(
        "어떤 불성실 패턴을 찾을까요?",
        ["1️⃣ 한 줄 찍기 (1,1,1,1...)", "2️⃣ 계단/지그재그 (1,2,3,2,1...)"],
        index=0,
        horizontal=True
    )
    
    # 그룹별 검사 진행
    for i, g in enumerate(st.session_state.ed_grps):
        k=f"ed_ms_{i}"; 
        if k not in st.session_state: st.session_state[k]=g['cols']
        sel = st.multiselect(f"그룹 {i+1} 변수 확인", df_raw.columns, key=k)
        st.session_state.ed_grps[i]['cols']=sel
        
        if sel:
            try:
                temp_df = df_raw[sel].apply(pd.to_numeric, errors='coerce')
                bad_indices = []
                
                if "한 줄 찍기" in check_method:
                    std = temp_df.std(axis=1)
                    bad_indices = std[std==0].index.tolist()
                    
                else: # 계단/지그재그
                    diffs = temp_df.diff(axis=1).iloc[:, 1:]
                    abs_diffs = diffs.abs()
                    is_zigzag = abs_diffs.eq(1).all(axis=1)
                    bad_indices = is_zigzag[is_zigzag].index.tolist()

                if bad_indices:
                    st.error(f"🚨 그룹 {i+1}: {len(bad_indices)}명 의심 패턴 발견")
                    bad_ids.update(bad_indices)
                else:
                    st.success(f"✅ 그룹 {i+1}: 해당 패턴 없음")
                    
            except Exception as e: 
                st.warning(f"계산 불가: {e}")
    
    # [NEW] 결과 확인 및 다운로드 섹션
    st.markdown("---")
    if bad_ids:
        st.subheader(f"🧐 불성실 의심 응답자 확인 (총 {len(bad_ids)}명)")
        st.caption("제거하기 전에 아래 표에서 응답 패턴을 눈으로 직접 확인하세요.")
        
        # 의심되는 사람들의 데이터만 추출
        bad_df_preview = df_raw.loc[list(bad_ids)]
        
        # 1. 엑셀처럼 보여주기 (여기서 눈으로 확인!)
        st.dataframe(bad_df_preview, use_container_width=True)
        
        st.markdown("---")
        col_down1, col_down2 = st.columns([1, 1])
        
        with col_down1:
            if st.button("🗑️ 확인했습니다. 제거하고 다운로드", type="primary"):
                final = df_cln.drop(index=list(bad_ids))
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as w: final.to_excel(w, index=False)
                st.download_button("📥 정제된 파일 받기", out.getvalue(), "cleaned_data.xlsx")
                
        with col_down2:
            # 의심자 목록만 따로 다운로드하고 싶을 수도 있으니 추가
            out_bad = io.BytesIO()
            with pd.ExcelWriter(out_bad, engine='xlsxwriter') as w: bad_df_preview.to_excel(w, index=False)
            st.download_button("📥 (참고용) 불성실 의심자 목록만 다운로드", out_bad.getvalue(), "bad_respondents.xlsx")
            
    else:
        st.info("검출된 불성실 응답자가 없습니다.")


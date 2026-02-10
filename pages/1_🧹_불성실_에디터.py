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
    
    # [NEW] 검사 옵션 선택 기능 추가
    st.markdown("---")
    st.subheader("🔍 검사 옵션")
    check_method = st.radio(
        "어떤 불성실 패턴을 찾을까요?",
        ["1️⃣ 한 줄 찍기 (1,1,1,1...)", "2️⃣ 계단/지그재그 (1,2,3,2,1...)"],
        index=0,
        horizontal=True
    )
    
    for i, g in enumerate(st.session_state.ed_grps):
        k=f"ed_ms_{i}"; 
        if k not in st.session_state: st.session_state[k]=g['cols']
        sel = st.multiselect(f"그룹 {i+1} 변수 확인", df_raw.columns, key=k)
        st.session_state.ed_grps[i]['cols']=sel
        
        if sel:
            try:
                # 데이터 숫자로 변환
                temp_df = df_raw[sel].apply(pd.to_numeric, errors='coerce')
                
                bad_indices = []
                
                if "한 줄 찍기" in check_method:
                    # 기존 로직: 표준편차 0
                    std = temp_df.std(axis=1)
                    bad_indices = std[std==0].index.tolist()
                    
                else: # 계단/지그재그 (1,2,3,2,1)
                    # 신규 로직: 앞뒤 차이의 절댓값이 모두 1인지 확인
                    # diff(axis=1)은 앞 열과의 차이를 구함
                    diffs = temp_df.diff(axis=1).iloc[:, 1:] # 첫 열은 NaN이므로 제외
                    abs_diffs = diffs.abs()
                    
                    # 모든 칸의 차이가 정확히 1인 행만 찾음 (all)
                    # (실수 오차 방지를 위해 isclose 대신 간단히 eq(1) 사용)
                    is_zigzag = abs_diffs.eq(1).all(axis=1)
                    bad_indices = is_zigzag[is_zigzag].index.tolist()

                if bad_indices:
                    st.error(f"🚨 그룹 {i+1}: {len(bad_indices)}명 불성실 의심")
                    bad_ids.update(bad_indices)
                else:
                    st.success(f"✅ 그룹 {i+1}: 해당 패턴 없음")
                    
            except Exception as e: 
                st.warning(f"계산 불가 (숫자형 데이터인지 확인 필요): {e}")
    
    st.markdown("---")
    if bad_ids:
        st.write(f"🛑 **총 제거 대상:** {len(bad_ids)}명")
        if st.button("🗑️ 불성실 응답자 제거 후 다운로드", type="primary"):
            final = df_cln.drop(index=list(bad_ids))
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w: final.to_excel(w, index=False)
            st.download_button("📥 정제된 파일 다운로드", out.getvalue(), "cleaned_data.xlsx")
    else:
        st.info("검출된 불성실 응답자가 없습니다.")

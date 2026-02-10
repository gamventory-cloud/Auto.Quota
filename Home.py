# Home.py
import streamlit as st
import utils # 위에서 만든 utils.py를 불러옵니다

st.set_page_config(page_title="Quota Master Pro", layout="wide")

# 비밀번호 체크
if not utils.check_password():
    st.stop()

st.title("🏠 Quota Master Pro 홈")
st.markdown("""
### 환영합니다! 👋
왼쪽 사이드바 메뉴에서 원하시는 작업을 선택해주세요.

---
#### 🛠️ 제공 기능
1.  **🧹 불성실 응답자 에디터:** * 특정 문항 범위나 키워드로 불성실 응답(한 줄 찍기 등)을 찾아내고 제거합니다.
2.  **📊 쿼터 자동 할당 솔루션:**
    * 메인 쿼터와 복잡한 서브 쿼터를 동시에 만족하는 최적의 응답자 조합을 찾아냅니다.
3.  **🛠️ SPSS 변수명 정제:**
    * Raw 데이터와 Code북을 비교하여 SPSS 변수명(Q1 -> SQ1 등)을 자동으로 매칭하고 변경합니다.
""")
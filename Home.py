# Home.py
import streamlit as st
import utils

st.set_page_config(page_title="Quota Master Pro", layout="wide",
                   page_icon="◧", initial_sidebar_state="expanded")

# 비밀번호 체크
if not utils.check_password():
    st.stop()


def _html(block: str) -> str:
    """
    마크다운이 HTML을 건드리지 않도록 한 줄로 눌러서 넘긴다.

    Streamlit 은 st.markdown 에 넘긴 문자열을 먼저 마크다운으로 해석한다.
    이때 빈 줄 다음에 4칸 이상 들여쓴 줄이 오면 그 부분을 **코드 블록**으로
    보고 태그를 글자 그대로 출력해 버린다. (카드 div 가 통째로 깨졌던 원인)

    각 줄의 앞뒤 공백을 없애고 빈 줄을 버린 뒤 공백 하나로 이어 붙이면
    들여쓰기와 빈 줄이 모두 사라져서 이 문제가 생기지 않는다.
    공백으로 잇는 이유는, 줄바꿈으로 나뉜 한글 문장이 붙어버리는 것을 막기 위해서다.
    """
    return " ".join(line.strip() for line in block.splitlines() if line.strip())


# ==============================================================================
# 스타일
#   - 배경/글자색은 Streamlit 테마를 따라가도록 반투명 + currentColor 로 처리한다.
#     (라이트/다크 어느 쪽에서도 카드가 겉돌지 않는다)
#   - 고정색은 강조색 하나뿐이다.
# ==============================================================================
st.markdown(_html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+KR:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
:root{
  --accent:#0E7C7B;
  --line:rgba(128,144,170,.28);
  --surface:rgba(128,144,170,.07);
  --mono:'IBM Plex Mono',ui-monospace,monospace;
}
.qm *{font-family:'IBM Plex Sans KR',-apple-system,'Malgun Gothic',sans-serif;}
.qm-hero{padding:8px 0 22px;border-bottom:1px solid var(--line);margin-bottom:26px;}
.qm-eyebrow{font-family:var(--mono);font-size:11.5px;letter-spacing:.18em;
  text-transform:uppercase;color:var(--accent);font-weight:600;margin-bottom:12px;}
.qm-title{font-size:44px;font-weight:700;letter-spacing:-.03em;line-height:1.1;
  margin:0 0 10px;color:inherit;}
.qm-title em{font-style:normal;color:var(--accent);}
.qm-sub{font-size:15.5px;line-height:1.65;opacity:.72;max-width:58ch;margin:0;}
.qm-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;align-items:stretch;}
.qm-card{border:1px solid var(--line);border-radius:12px;padding:22px 22px 18px;
  background:var(--surface);position:relative;overflow:hidden;
  display:flex;flex-direction:column;
  transition:transform .16s ease,border-color .16s ease;}
.qm-card:hover{transform:translateY(-2px);border-color:var(--accent);}
.qm-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;
  background:var(--accent);opacity:0;transition:opacity .16s ease;}
.qm-card:hover::before{opacity:1;}
.qm-card.feature::before{opacity:1;}
.qm-stage{font-family:var(--mono);font-size:11px;font-weight:600;letter-spacing:.1em;
  text-transform:uppercase;color:var(--accent);margin-bottom:9px;}
.qm-name{font-size:19px;font-weight:600;letter-spacing:-.02em;margin:0 0 8px;}
.qm-desc{font-size:14px;line-height:1.6;opacity:.74;margin:0 0 16px;flex:1 1 auto;}
.qm-desc code{font-family:var(--mono);font-size:12.5px;padding:1px 5px;
  border-radius:4px;background:var(--surface);border:1px solid var(--line);}
.qm-io{display:flex;align-items:center;gap:9px;flex-wrap:wrap;
  padding-top:14px;border-top:1px dashed var(--line);}
.qm-chip{font-family:var(--mono);font-size:11px;padding:4px 9px;border-radius:5px;
  border:1px solid var(--line);white-space:nowrap;opacity:.85;}
.qm-chip.out{border-color:var(--accent);color:var(--accent);opacity:1;}
.qm-io-arrow{font-family:var(--mono);font-size:12px;opacity:.4;}
.qm-foot{margin-top:28px;padding-top:18px;border-top:1px solid var(--line);
  font-size:13px;opacity:.6;line-height:1.7;}
@media (max-width:1200px){ .qm-grid{grid-template-columns:repeat(2,1fr);} }
@media (max-width:760px){ .qm-grid{grid-template-columns:1fr;} .qm-title{font-size:32px;} }
@media (prefers-reduced-motion:reduce){
  .qm-card{transition:none;} .qm-card:hover{transform:none;}
}
</style>
"""), unsafe_allow_html=True)


# ==============================================================================
# 본문
#   설명 문구와 입출력 칩은 이 블록만 고치면 된다.
# ==============================================================================
st.markdown(_html("""
<div class="qm">
  <div class="qm-hero">
    <div class="qm-eyebrow">Survey Data Toolkit</div>
    <h1 class="qm-title">설문지에서 <em>최종 표본</em>까지</h1>
    <p class="qm-sub">
      조사 데이터를 정리하고 쿼터를 맞추는 작업을 한곳에서 처리합니다.
      왼쪽 사이드바에서 도구를 선택하세요.
    </p>
  </div>
  <div class="qm-grid">
    <div class="qm-card feature">
      <div class="qm-stage">표본 확정</div>
      <h2 class="qm-name">쿼터 솔루션</h2>
      <p class="qm-desc">
        메인 쿼터와 추가 쿼터를 동시에 만족하는 응답자 조합을 찾습니다.
        목표를 못 채우면 원인을 가려내고, 어떤 조건의 응답자를
        몇 명 더 모아야 하는지까지 알려줍니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">정제된 데이터</span>
        <span class="qm-chip">쿼터표</span>
        <span class="qm-io-arrow">&rarr;</span>
        <span class="qm-chip out">최종 표본</span>
        <span class="qm-chip out">달성 현황 리포트</span>
        <span class="qm-chip out">추가 수집 지시서</span>
      </div>
    </div>
    <div class="qm-card">
      <div class="qm-stage">자료 준비</div>
      <h2 class="qm-name">설문지 변환</h2>
      <p class="qm-desc">
        설문지 워드 파일을 읽어 문항 구조를 분석하고 엑셀 코드북으로 만듭니다.
        단수·복수·표·순위형을 구분해 정리합니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">설문지 .docx</span>
        <span class="qm-io-arrow">&rarr;</span>
        <span class="qm-chip out">코드북 .xlsx</span>
      </div>
    </div>
    <div class="qm-card">
      <div class="qm-stage">자료 준비</div>
      <h2 class="qm-name">SPSS 라벨링</h2>
      <p class="qm-desc">
        코드북을 바탕으로 SPSS 초기 세팅 신택스를 만듭니다.
        변수 라벨과 값 라벨을 한 번에 입힙니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">코드북 .xlsx</span>
        <span class="qm-io-arrow">&rarr;</span>
        <span class="qm-chip out">신택스 .sps</span>
      </div>
    </div>
    <div class="qm-card">
      <div class="qm-stage">데이터 정리</div>
      <h2 class="qm-name">SPSS 정제</h2>
      <p class="qm-desc">
        원자료와 코드북을 대조해 변수명을 자동으로 맞춥니다.
        <code>Q1</code>을 <code>SQ1</code>로 바꾸는 식의 작업을 일괄 처리합니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">원자료</span>
        <span class="qm-chip">코드북</span>
        <span class="qm-io-arrow">&rarr;</span>
        <span class="qm-chip out">변수명 정리된 데이터</span>
      </div>
    </div>
    <div class="qm-card">
      <div class="qm-stage">데이터 정리</div>
      <h2 class="qm-name">응답시간 변환</h2>
      <p class="qm-desc">
        문항별 체류시간이 세로로 쌓여 있는 데이터를 한 명당 한 줄로 돌립니다.
        응답자 단위로 붙여야 원자료와 합치거나 불성실 응답을 걸러낼 수 있습니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">체류시간 원자료 (세로)</span>
        <span class="qm-io-arrow">&rarr;</span>
        <span class="qm-chip out">응답자별 가로 데이터</span>
      </div>
    </div>
    <div class="qm-card">
      <div class="qm-stage">내보내기</div>
      <h2 class="qm-name">SAV 변환</h2>
      <p class="qm-desc">
        엑셀·CSV 표를 SPSS에서 바로 열리는 <code>.sav</code> 파일로 만듭니다.
        한글 열 이름은 변수 라벨로, 글자 응답은 숫자 코드와 값 라벨로 바꿉니다.
        문항이 여러 시트에 나뉘어 있으면 <code>id</code> 기준으로 붙입니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">원자료 .xlsx / .csv</span>
        <span class="qm-chip">여러 시트</span>
        <span class="qm-io-arrow">&rarr;</span>
        <span class="qm-chip out">SPSS .sav</span>
      </div>
    </div>
  </div>
  <div class="qm-foot">
    지인들만 사용하는 비공개 도구입니다.
  </div>
</div>
"""), unsafe_allow_html=True)

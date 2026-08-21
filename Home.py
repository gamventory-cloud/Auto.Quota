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
  --ink:#1F3864;          /* 결과 엑셀 헤더와 같은 남색 — 제품 전체의 기준색 */
  --accent:#0E7C7B;       /* 강조 : 단 한 곳에만 */
  --line:rgba(128,144,170,.28);
  --surface:rgba(128,144,170,.07);
  --mono:'IBM Plex Mono',ui-monospace,monospace;
}
.qm *{font-family:'IBM Plex Sans KR',-apple-system,'Malgun Gothic',sans-serif;}

/* ── 히어로 ───────────────────────────────────────────── */
.qm-hero{padding:8px 0 22px;border-bottom:1px solid var(--line);margin-bottom:26px;}
.qm-eyebrow{font-family:var(--mono);font-size:11.5px;letter-spacing:.18em;
  text-transform:uppercase;color:var(--accent);font-weight:600;margin-bottom:12px;}
.qm-title{font-size:44px;font-weight:700;letter-spacing:-.03em;line-height:1.1;
  margin:0 0 10px;color:inherit;}
.qm-title em{font-style:normal;color:var(--accent);}
.qm-sub{font-size:15.5px;line-height:1.65;opacity:.72;max-width:56ch;margin:0;}

/* ── 작업 흐름 리본 ────────────────────────────────────── */
.qm-flow{display:flex;flex-wrap:wrap;align-items:center;gap:10px;
  padding:14px 16px;border:1px solid var(--line);border-radius:10px;
  background:var(--surface);margin-bottom:30px;}
.qm-flow-label{font-family:var(--mono);font-size:11px;letter-spacing:.14em;
  text-transform:uppercase;opacity:.55;margin-right:4px;}
.qm-step{font-size:13px;font-weight:500;white-space:nowrap;}
.qm-step b{font-family:var(--mono);font-size:11px;color:var(--accent);
  margin-right:5px;font-weight:600;}
.qm-arrow{opacity:.35;font-size:12px;}

/* ── 카드 그리드 ───────────────────────────────────────── */
.qm-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:16px;}
.qm-card{border:1px solid var(--line);border-radius:12px;padding:22px 22px 18px;
  background:var(--surface);position:relative;overflow:hidden;
  transition:transform .16s ease,border-color .16s ease;}
.qm-card:hover{transform:translateY(-2px);border-color:var(--accent);}
.qm-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;
  background:var(--accent);opacity:0;transition:opacity .16s ease;}
.qm-card:hover::before{opacity:1;}
.qm-num{font-family:var(--mono);font-size:11px;font-weight:600;letter-spacing:.1em;
  color:var(--accent);margin-bottom:9px;}
.qm-name{font-size:19px;font-weight:600;letter-spacing:-.02em;margin:0 0 8px;}
.qm-desc{font-size:14px;line-height:1.6;opacity:.74;margin:0 0 16px;}
.qm-desc + .qm-desc{margin-top:-8px;}

/* 입력 → 출력 : 이 도구들은 전부 파일을 받아 파일을 내놓는다 */
.qm-io{display:flex;align-items:center;gap:9px;flex-wrap:wrap;
  padding-top:14px;border-top:1px dashed var(--line);}
.qm-chip{font-family:var(--mono);font-size:11px;padding:4px 9px;border-radius:5px;
  border:1px solid var(--line);white-space:nowrap;opacity:.85;}
.qm-chip.out{border-color:var(--accent);color:var(--accent);opacity:1;}
.qm-io-arrow{font-family:var(--mono);font-size:12px;opacity:.4;}

.qm-foot{margin-top:30px;padding-top:18px;border-top:1px solid var(--line);
  font-size:13px;opacity:.6;line-height:1.7;}
.qm-foot code{font-family:var(--mono);font-size:12px;}
.qm-desc code{font-family:var(--mono);font-size:12.5px;padding:1px 5px;
  border-radius:4px;background:var(--surface);border:1px solid var(--line);}

@media (max-width:640px){ .qm-title{font-size:32px;} }
@media (prefers-reduced-motion:reduce){
  .qm-card{transition:none;} .qm-card:hover{transform:none;}
}
</style>
"""), unsafe_allow_html=True)


# ==============================================================================
# 본문
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

  <div class="qm-flow">
    <span class="qm-flow-label">작업 순서</span>
    <span class="qm-step"><b>04</b>코드북 만들기</span>
    <span class="qm-arrow">→</span>
    <span class="qm-step"><b>03</b>변수명 맞추기</span>
    <span class="qm-arrow">→</span>
    <span class="qm-step"><b>01</b>불성실 응답 걸러내기</span>
    <span class="qm-arrow">→</span>
    <span class="qm-step"><b>02</b>쿼터 확정</span>
  </div>

  <div class="qm-grid">

    <div class="qm-card">
      <div class="qm-num">01 / 데이터 정제</div>
      <h2 class="qm-name">불성실 응답자 에디터</h2>
      <p class="qm-desc">
        한 줄 찍기처럼 성의 없는 응답을 찾아 걸러냅니다.
        문항 범위나 키워드로 조건을 지정할 수 있습니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">응답 원자료</span>
        <span class="qm-io-arrow">→</span>
        <span class="qm-chip out">정제된 데이터</span>
      </div>
    </div>

    <div class="qm-card">
      <div class="qm-num">02 / 표본 확정</div>
      <h2 class="qm-name">쿼터 자동 할당</h2>
      <p class="qm-desc">
        메인 쿼터와 추가 쿼터를 동시에 만족하는 응답자 조합을 찾습니다.
        못 맞추면 어느 쿼터가 막고 있는지, 누구를 더 모아야 하는지 알려줍니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">정제된 데이터 + 쿼터표</span>
        <span class="qm-io-arrow">→</span>
        <span class="qm-chip out">최종 표본</span>
      </div>
    </div>

    <div class="qm-card">
      <div class="qm-num">03 / 변수 정리</div>
      <h2 class="qm-name">SPSS 변수명 정제</h2>
      <p class="qm-desc">
        원자료와 코드북을 대조해 변수명을 자동으로 맞춥니다.
        <code>Q1</code>을 <code>SQ1</code>로 바꾸는 식의 작업을 일괄 처리합니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">원자료 + 코드북</span>
        <span class="qm-io-arrow">→</span>
        <span class="qm-chip out">변수명 정리된 데이터</span>
      </div>
    </div>

    <div class="qm-card">
      <div class="qm-num">04 / 자료 준비</div>
      <h2 class="qm-name">코드북 · 신택스 생성</h2>
      <p class="qm-desc">
        설문지 워드 파일을 읽어 문항 구조(단수·복수·표·순위형)를 분석하고
        엑셀 코드북을 만듭니다.
      </p>
      <p class="qm-desc">
        그 코드북으로 SPSS 초기 세팅 신택스까지 이어서 만듭니다.
        Variable Label, Value Label, Rename, Recode가 들어갑니다.
      </p>
      <div class="qm-io">
        <span class="qm-chip">설문지 .docx</span>
        <span class="qm-io-arrow">→</span>
        <span class="qm-chip out">코드북 .xlsx</span>
        <span class="qm-io-arrow">→</span>
        <span class="qm-chip out">신택스 .sps</span>
      </div>
    </div>

  </div>

  <div class="qm-foot">
    번호는 사이드바 메뉴 순서입니다. 실제 작업은 위 <b>작업 순서</b>대로 진행하면 됩니다.<br>
    지인들만 사용하는 비공개 도구입니다.
  </div>

</div>
"""), unsafe_allow_html=True)

# Home.py
#
# 이 파일은 '라우터' 입니다. 페이지를 이동할 때마다 매번 실행되므로
# 여기에는 공통 설정만 두고, 화면 내용은 각 페이지 파일에 둡니다.
# 홈 화면의 카드 그리드는 홈화면.py 에 있습니다.
#
# ── 파일 경로를 적지 않는 이유 ────────────────────────────────────────
#   경로를 정확히 적으면 번호 접두사가 바뀌거나 이름을 조금 고쳐도 깨집니다.
#   그래서 '핵심 단어'로 찾도록 했습니다.
#   예: ("쿼터",) 는 2___쿼터_솔루션.py 든 쿼터_배분.py 든 찾아냅니다.
#
# ── 페이지 추가하는 법 ────────────────────────────────────────────────
#   PAGES 표에 (찾을 단어, 제외할 단어, 사이드바 이름, 아이콘) 을 넣습니다.
#   못 찾으면 사이드바에 안내가 뜨고 실제 파일 목록도 함께 보여줍니다.
# =====================================================================

from pathlib import Path

import streamlit as st

import utils

st.set_page_config(page_title="Quota Master Pro", layout="wide",
                   page_icon="◧", initial_sidebar_state="expanded")

PAGES: dict = {
    "시작": [
        (("홈화면",), (), "홈", "🏠"),
    ],
    "표본 확정": [
        (("쿼터",), (), "쿼터 솔루션", "🎯"),
    ],
    "내보내기": [
        (("SAV", "변환"), ("엑셀",), "Excel → Sav", "💾"),
        (("SAV", "엑셀"), (), "Sav → Excel", "📗"),
    ],
    "집계": [
        (("뱅크표",), (), "뱅크표 생성", "📊"),
    ],
    "자료 준비": [
        (("라벨",), (), "SPSS 라벨링", "🏷️"),
        (("설문지",), (), "HWP → Word", "📄"),
#        (("에디팅",), (), "에디팅 신택스", "🧾"),
    ],
    "데이터 정리": [
        (("정제",), (), "RD 변수명 수정", "🧹"),
        (("행열",), (), "행열 데이터 변환", "🔄"),
        (("지역코드",), (), "지역코드 검증", "📍"),
    ],
}

HOME_TITLE = "홈"
HOME_SECTION = "시작"      # 이 구역은 항상 펼쳐진 상태로 위에 둔다


# ── 후보 파일 모으기 ─────────────────────────────────────────────────
# 이 파일이 있는 폴더와 그 아래 한 단계까지 훑는다.
# pages 폴더 이름이 다르거나 파일이 최상단에 있어도 찾아낸다.
here = Path(__file__).resolve().parent
SKIP = {"home.py", "utils.py", "spss_labels.py",
        "dp_syntax.py", "sps_engine.py"}

candidates: list = []
for p in sorted(here.glob("*.py")):
    if p.name.lower() not in SKIP:
        candidates.append(p)
for sub in sorted(here.iterdir()):
    if sub.is_dir() and not sub.name.startswith((".", "__")):
        for p in sorted(sub.glob("*.py")):
            if p.name.lower() not in SKIP:
                candidates.append(p)


def find_page(want: tuple, avoid: tuple):
    """파일명에 want 단어가 모두 있고 avoid 단어는 없는 파일. 대소문자 무시."""
    hits = [
        p for p in candidates
        if all(w.lower() in p.stem.lower() for w in want)
        and not any(w.lower() in p.stem.lower() for w in avoid)
    ]
    if not hits:
        return None
    # 여러 개면 이름이 짧은 쪽을 고른다
    return min(hits, key=lambda p: len(p.stem))


def file_list() -> str:
    return "\n".join(p.relative_to(here).as_posix() for p in candidates) or "(없음)"


nav: dict = {}
missing: list = []
bad_icons: list = []
used: set = set()

for section, items in PAGES.items():
    built = []
    for want, avoid, title, icon in items:
        hit = find_page(want, avoid)
        if hit is None or hit in used:
            missing.append(title)
            continue
        used.add(hit)
        rel = hit.relative_to(here).as_posix()
        # st.Page 의 icon 은 진짜 이모지만 받는다. 실패하면 아이콘만 버린다.
        try:
            pg = st.Page(rel, title=title, icon=icon,
                         default=(title == HOME_TITLE))
        except Exception:                             # noqa: BLE001
            bad_icons.append(f"{title} ({icon})")
            pg = st.Page(rel, title=title, default=(title == HOME_TITLE))
        built.append(pg)
    if built:
        nav[section] = built

# ── 아무것도 못 찾았을 때: 실제 파일 목록을 보여준다 ─────────────────
if not nav:
    st.error("사이드바에 넣을 페이지를 하나도 찾지 못했습니다.")
    st.write("**이 파일이 있는 폴더**")
    st.code(str(here))
    st.write(f"**찾은 .py 파일 {len(candidates)}개**")
    st.code(file_list())
    st.caption(
        "목록에 페이지 파일이 보이는데도 안 잡혔다면 이름을 알려 주세요. "
        "목록 자체가 비어 있다면 이 Home.py 가 저장소 최상단에 있는지 확인해 주세요."
    )
    st.stop()

# ── st.navigation 을 먼저 호출한다 ───────────────────────────────────
# 이 명령이 실행되는 순간부터 Streamlit 이 pages/ 자동 탐색을 끈다.
# 비밀번호 확인 뒤에 두면, 로그인 화면에서는 아직 호출되지 않은 상태라
# 자동 탐색으로 만들어진 옛 사이드바(파일명 그대로인 목록)가 그대로 보인다.
#
# position="hidden" 이므로 이 호출 자체는 화면에 아무것도 그리지 않는다.
# 사이드바는 아래에서 손으로 만든다.
page = st.navigation(nav, position="hidden")

# ── 비밀번호 ─────────────────────────────────────────────────────────
# 사이드바를 그리기 전에 막는다. 로그인 화면에는 사이드바가 비어 있게 된다.
if not utils.check_password():
    st.stop()

if "nav_open" not in st.session_state:
    st.session_state["nav_open"] = True

with st.sidebar:
    # 홈은 접힘과 무관하게 항상 보인다
    for pg in nav.get(HOME_SECTION, []):
        st.page_link(pg)

    others = {k: v for k, v in nav.items() if k != HOME_SECTION}
    if others:
        opened = st.session_state["nav_open"]
        # tertiary = 테두리 없는 링크 형태. 구역 캡션과 톤이 맞는다.
        # use_container_width 를 주면 다시 박스처럼 보이므로 쓰지 않는다.
        if st.button("도구 접기" if opened else "도구 펼치기",
                     type="tertiary", icon=":material/unfold_less:" if opened
                     else ":material/unfold_more:"):
            st.session_state["nav_open"] = not opened
            st.rerun()

        if st.session_state["nav_open"]:
            for section, pgs in others.items():
                st.caption(section)
                for pg in pgs:
                    st.page_link(pg)

    if missing:
        st.warning("목록에서 빠진 페이지 — " + ", ".join(missing))
        with st.expander("실제 파일 목록 보기"):
            st.code(str(here))
            st.code(file_list())

    if bad_icons:
        st.info("이모지가 아니라 아이콘 없이 표시합니다 — " + ", ".join(bad_icons))

page.run()
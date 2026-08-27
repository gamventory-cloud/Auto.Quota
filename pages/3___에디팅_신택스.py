# -*- coding: utf-8 -*-
"""
3___에디팅_신택스.py  (v2)

설문지(.docx)를 읽어 에디팅 체크 신택스(.sps)를 만든다.

화면 순서
  ① 오류·확인      파서가 이상하다고 본 것. 먼저 보고 넘어간다.
  ② 변수 배정      문항 → 시작 변수. 사람이 맞춘다.
  ③ 지시문 처리    인식한 것도 못 한 것도 전부. 조건식을 확정한다.
  ④ 신택스         생성 · 다운로드
  ⑤ 검증           실제 데이터에 실행해 체크별 건수 확인

set_page_config 와 비밀번호 확인은 Home.py 에서 이미 처리합니다.
"""

from __future__ import annotations

import io
import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import dp_syntax as DP
import sps_engine as ENG

VERSION = "v2"
NONE = "(미지정)"

KIND_LABEL = {
    "": "미해석",
    "range": "범위",
    "require": "특정 값 필수",
    "exclusive": "단독 선택 보기",
    "require_sel": "필수 선택 보기",
    "not_all_zero": "전부 0 금지",
    "same_all": "여러 문항 동일값 금지",
}


# ---------------------------------------------------------------------------
# 파일 읽기
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def parse_bytes(file_bytes: bytes, name: str):
    # Windows 에서 NamedTemporaryFile 은 열린 채로 다시 열 수 없어 디렉터리를 쓴다
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / (Path(name).name or "survey.docx")
        p.write_bytes(file_bytes)
        return DP.parse_survey(str(p))


@st.cache_data(show_spinner=False)
def read_data(file_bytes: bytes, name: str) -> pd.DataFrame:
    suf = Path(name).suffix.lower()
    if suf == ".sav":
        import pyreadstat
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / Path(name).name
            p.write_bytes(file_bytes)
            df, _ = pyreadstat.read_sav(str(p), apply_value_formats=False)
        return df
    if suf == ".xls":
        return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
    if suf == ".csv":
        return pd.read_csv(io.BytesIO(file_bytes))
    return pd.read_excel(io.BytesIO(file_bytes))


def value_summary(df: pd.DataFrame | None, cols: list[str]) -> str:
    """배정한 변수의 실제 값 범위. 오배정을 눈으로 잡는 장치."""
    if df is None or not cols:
        return ""
    present = [c for c in cols if c in df.columns]
    if not present:
        return "⚠ 데이터에 없음"
    missing_names = [c for c in cols if c not in df.columns]
    s = df[present].apply(pd.to_numeric, errors="coerce")
    arr = s.to_numpy(dtype="float64", na_value=float("nan")).ravel()
    vals = pd.Series(arr).dropna()
    if len(vals) == 0:
        return "전부 결측"
    out = f"{vals.min():g}~{vals.max():g} · 종류 {vals.nunique()}"
    if missing_names:
        out += f" · ⚠ 없는 변수 {len(missing_names)}개"
    return out


# ---------------------------------------------------------------------------
st.title("🧾 에디팅 신택스")
st.caption(
    f"설문지에서 체크 신택스를 만듭니다 · {VERSION} · "
    "문항 구조는 자동, 지시문 해석과 변수 배정은 직접 확인합니다."
)

c1, c2 = st.columns(2)
doc_file = c1.file_uploader("설문지 / DP 문서 (.docx)", type=["docx"])
data_file = c2.file_uploader(
    "데이터 파일 (선택 · .sav / .xlsx / .xls / .csv)",
    type=["sav", "xlsx", "xls", "csv"],
    help="올리면 변수명을 목록에서 고를 수 있고, 만든 신택스를 바로 실행해 검증합니다.",
)

if not doc_file:
    st.info("설문지를 올리면 문항 목록과 지시문 목록이 만들어집니다.")
    st.stop()

questions, directives, issues, leftovers = parse_bytes(doc_file.getvalue(), doc_file.name)
if not questions:
    st.error("문항을 찾지 못했습니다. 문서 형식이 예상과 다를 수 있습니다.")
    st.stop()

df = None
if data_file:
    try:
        df = read_data(data_file.getvalue(), data_file.name)
    except Exception as e:  # noqa: BLE001
        st.warning(f"데이터를 읽지 못했습니다: {e}")

data_cols = list(df.columns) if df is not None else []
qmap = {q.qid: q for q in questions}

m1, m2, m3, m4 = st.columns(4)
m1.metric("문항", f"{len(questions)}")
m2.metric("지시문", f"{len(directives)}")
m3.metric("오류", f"{sum(1 for i in issues if i.level == DP.LV_ERROR)}")
m4.metric("데이터", f"{len(df):,}행" if df is not None else "없음")

# ---------------------------------------------------------------------------
# ① 오류 · 확인
# ---------------------------------------------------------------------------
st.divider()
st.subheader("① 오류 · 확인")

if issues:
    err = [i for i in issues if i.level == DP.LV_ERROR]
    chk = [i for i in issues if i.level == DP.LV_CHECK]
    if err:
        st.error(f"오류 {len(err)}건 — 신택스를 만들기 전에 확인해 주세요.")
        st.dataframe(
            pd.DataFrame([{"위치": i.where, "내용": i.what} for i in err]),
            use_container_width=True, hide_index=True)
    if chk:
        with st.expander(f"확인 {len(chk)}건", expanded=not err):
            st.dataframe(
                pd.DataFrame([{"위치": i.where, "내용": i.what} for i in chk]),
                use_container_width=True, hide_index=True)
else:
    st.success("파서가 이상을 발견하지 못했습니다.")

if leftovers:
    with st.expander(f"어디에도 넣지 못한 줄 {len(leftovers)}개"):
        st.code("\n".join(leftovers[:60]))

# ---------------------------------------------------------------------------
# ② 변수 배정
# ---------------------------------------------------------------------------
st.divider()
st.subheader("② 변수 배정")
st.caption(
    "**시작 변수**만 고르면 개수만큼 이어서 배정합니다 (v13_1 에 36개 → v13_1 ~ v13_36). "
    "유형과 개수도 고칠 수 있습니다."
)

saved: dict = {}
mfile = st.file_uploader("저장한 배정 JSON 불러오기 (선택)", type=["json"], key="mapjson")
if mfile is not None:
    try:
        saved = json.loads(mfile.getvalue().decode("utf-8"))
        st.caption("배정을 불러왔습니다.")
    except Exception as e:  # noqa: BLE001
        st.warning(f"불러오지 못했습니다: {e}")

skel = DP.mapping_skeleton(questions)
for row in skel:
    prev = saved.get(row["문항"], {})
    row["시작변수"] = prev.get("start", "")
    if prev.get("n"):
        row["개수"] = int(prev["n"])
    if prev.get("type") in DP.ALL_TYPES:
        row["유형"] = prev["type"]

base = pd.DataFrame(skel)[["문항", "유형", "개수", "시작변수", "구조", "문항명"]]
cfg = {
    "문항": st.column_config.TextColumn("문항", disabled=True, width="small"),
    "유형": st.column_config.SelectboxColumn("유형", options=DP.ALL_TYPES, width="small"),
    "개수": st.column_config.NumberColumn("개수", min_value=0, max_value=300, step=1,
                                        width="small"),
    "구조": st.column_config.TextColumn("문서 구조", disabled=True, width="small"),
    "문항명": st.column_config.TextColumn("문항명", disabled=True, width="large"),
    "시작변수": (st.column_config.SelectboxColumn("시작변수", options=[""] + data_cols,
                                             width="small")
             if data_cols else st.column_config.TextColumn("시작변수", width="small")),
}
edited = st.data_editor(base, key="mapedit", use_container_width=True,
                        hide_index=True, column_config=cfg)

mapping: dict[str, dict] = {}
for _, r in edited.iterrows():
    start = str(r["시작변수"] or "").strip()
    n = int(r["개수"] or 0)
    if not start or n <= 0:
        continue
    mapping[r["문항"]] = {"vars": DP.expand_vars(start, n), "n": n,
                          "start": start, "type": r["유형"]}
    if r["문항"] in qmap:
        qmap[r["문항"]].qtype = r["유형"]

if mapping:
    rows = []
    for qid, info in mapping.items():
        q = qmap[qid]
        expect = ""
        if q.qtype == DP.T_SINGLE and q.options:
            expect = (f"1~{len(q.options)}" if q.contiguous
                      else ",".join(str(c) for c in q.codes))
        elif q.qtype == DP.T_MATRIX and q.scale_codes:
            expect = f"{q.scale_codes[0]:g}~{q.scale_codes[-1]:g}"
        v = info["vars"]
        rows.append({
            "문항": qid,
            "배정": v[0] if len(v) == 1 else f"{v[0]} ~ {v[-1]} ({len(v)})",
            "기대 값": expect,
            "실제 값": value_summary(df, v),
        })
    with st.expander("배정 검토 — 기대 값과 실제 값이 어긋나면 잘못 고른 것입니다",
                     expanded=df is not None):
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

st.download_button(
    "배정 JSON 저장",
    data=json.dumps({k: {"start": v["start"], "n": v["n"], "type": v["type"]}
                     for k, v in mapping.items()},
                    ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="변수배정.json", mime="application/json")

# ---------------------------------------------------------------------------
# ③ 지시문 처리
# ---------------------------------------------------------------------------
st.divider()
st.subheader("③ 지시문 처리")
st.caption(
    "설문지에 있는 지시문 전부입니다. 인식한 것은 조건식을 채워뒀고, "
    "못 한 것은 비어 있습니다. **비워 두면 신택스에 들어가지 않습니다.**"
)

if not directives:
    st.info("지시문을 찾지 못했습니다.")

logic_blocks: list[tuple[str, str, str]] = []
n_filled = 0

show_all = st.checkbox("인식한 지시문까지 모두 펼쳐 보기", value=False)

for d in directives:
    sug, lv = DP.suggest_cond(d, mapping, qmap)
    default_lv = lv
    if not default_lv:
        v = mapping.get(d.scope, {}).get("vars", [])
        default_lv = f"{v[0]} to {v[-1]}" if len(v) > 1 else (v[0] if v else "")

    header = f"{d.did} · {d.scope} · {KIND_LABEL.get(d.kind, d.kind)}"
    with st.expander(header, expanded=show_all or not sug):
        st.caption(d.raw)
        if d.note:
            st.caption(f"해석: {d.note}")
        cc1, cc2 = st.columns([3, 2])
        cond = cc1.text_input("조건식 (위반인 케이스)", value=sug, key=f"cond_{d.did}",
                              placeholder="예: ~Range(v5,0,11)")
        lvars = cc2.text_input("List 변수", value=default_lv, key=f"lv_{d.did}")
    if cond.strip():
        n_filled += 1
        logic_blocks.append((f"{d.did} {d.scope} — {d.raw[:44]}", cond, lvars))

st.caption(f"조건식이 채워진 지시문 {n_filled} / {len(directives)}")

# ---------------------------------------------------------------------------
# ④ 신택스
# ---------------------------------------------------------------------------
st.divider()
st.subheader("④ 신택스")

if not mapping:
    st.info("변수를 하나 이상 배정하면 신택스가 만들어집니다.")
    st.stop()

o1, o2 = st.columns(2)
with o1:
    st.markdown("**결측 → -1 리코드 범위**")
    st.caption("이 범위의 결측을 -1 로 바꿔야 범위·ANY 검사에 무응답이 걸립니다.")
    if data_cols:
        first_var = st.selectbox("첫 변수", [NONE] + data_cols, index=1)
        last_var = st.selectbox("마지막 변수", [NONE] + data_cols, index=len(data_cols))
    else:
        first_var = st.text_input("첫 변수", value="")
        last_var = st.text_input("마지막 변수", value="")
    first_var = "" if first_var == NONE else first_var
    last_var = "" if last_var == NONE else last_var
with o2:
    st.markdown("**파일 경로 (선택)**")
    project = st.text_input("작업 폴더", value="", placeholder=r"D:\2026\(과제번호) 조사명")
    src_sav = st.text_input("원본 SAV 파일명", value="", placeholder="조사명_원본.sav")

matrices = [q for q in questions if q.qtype == DP.T_MATRIX and q.qid in mapping]
sl_targets: list[str] = []
qsort: dict[str, list[int]] = {}
if matrices:
    with st.expander(f"매트릭스 추가 체크 ({len(matrices)}문항)"):
        sl_targets = st.multiselect(
            "직진성 검사를 넣을 문항", [q.qid for q in matrices],
            help="문항 수가 적거나 척도가 좁으면 정상 응답도 걸립니다. "
                 "여러 문항에 걸친 동일값 규칙은 ③ 지시문 처리에서 다룹니다.")
        for q in matrices:
            txt = st.text_input(
                f"{q.qid} 강제분포 (척도 {len(q.scale_codes)} · 진술문 {q.n_items})",
                value="", placeholder="예: 2,3,4,5,8,5,4,3,2", key=f"qs_{q.qid}")
            if txt.strip():
                try:
                    qsort[q.qid] = [int(x) for x in txt.replace(" ", "").split(",") if x]
                except ValueError:
                    st.caption("⚠ 숫자를 쉼표로 구분해 주세요")

extra = st.text_area(
    "추가 체크 (설문지에 없는 항목 — 패널 변수, 체류시간 등)", value="", height=100,
    placeholder="Temp.\nSelect If ~Any(areaM,-1) .\nList Var no id areaM  .")

sps_text, warns = DP.build_syntax(
    questions, mapping, logic_blocks,
    project=project.strip(), src_sav=src_sav.strip(),
    first_var=first_var, last_var=last_var,
    straightline_only=set(sl_targets), qsort=qsort, extra_checks=extra,
)

st.caption(f"체크 {sps_text.lower().count('select if')}개 생성")
if warns:
    with st.expander(f"확인이 필요한 항목 {len(warns)}개", expanded=True):
        for w in warns:
            st.markdown(f"- {w}")

st.code(sps_text, language="text")

e1, e2 = st.columns(2)
enc = e1.radio("인코딩", ["cp949 (국내 SPSS)", "utf-8"], horizontal=True)
codec = "cp949" if enc.startswith("cp949") else "utf-8"
try:
    payload = sps_text.replace("\n", "\r\n").encode(codec)
except UnicodeEncodeError:
    payload = sps_text.replace("\n", "\r\n").encode(codec, errors="replace")
    st.caption("⚠ 일부 문자를 cp949 로 바꿀 수 없어 대체했습니다.")
e2.download_button("신택스 다운로드 (.sps)", data=payload,
                   file_name=f"{Path(doc_file.name).stem}_Editing.sps",
                   mime="text/plain", use_container_width=True)

# ---------------------------------------------------------------------------
# ⑤ 검증
# ---------------------------------------------------------------------------
if df is None:
    st.info("데이터 파일을 올리면 만든 신택스를 바로 실행해 확인할 수 있습니다.")
    st.stop()

st.divider()
st.subheader("⑤ 검증")
st.caption("특정 체크가 전원 위반으로 나오면 데이터가 아니라 변수 배정이 잘못된 신호입니다.")

if not st.button("데이터에 실행", type="primary"):
    st.stop()

with tempfile.TemporaryDirectory() as tmp:
    p = Path(tmp) / "generated.sps"
    p.write_text(sps_text, encoding="utf-8")
    try:
        checks, notes, out = ENG.run(str(p), df)
    except Exception as e:  # noqa: BLE001
        st.error(f"실행 중 오류: {e}")
        st.stop()

n = len(df)
rows = []
for c in checks:
    ratio = c.n_hit / max(n, 1) * 100
    if c.error:
        flag = "파싱 실패"
    elif c.n_hit == n:
        flag = "전원 위반 — 배정 확인"
    elif ratio >= 30:
        flag = "비율 높음"
    else:
        flag = ""
    rows.append({"#": c.seq, "조건": c.cond, "위반": c.n_hit,
                 "비율(%)": round(ratio, 1), "신호": flag,
                 "오류": c.error or "",
                 "케이스": ", ".join(str(x) for x in c.cases[:12])
                          + ("..." if len(c.cases) > 12 else "")})
res = pd.DataFrame(rows)

k1, k2, k3 = st.columns(3)
k1.metric("체크", f"{len(checks):,}")
k2.metric("위반 발생", f"{int((res['위반'] > 0).sum()):,}")
k3.metric("의심 신호", f"{int((res['신호'] != '').sum()):,}")

sus = res[res["신호"] != ""]
if not sus.empty:
    st.warning("아래 체크는 배정이나 조건식을 다시 볼 필요가 있습니다.")
    st.dataframe(sus, use_container_width=True, hide_index=True)

st.dataframe(res, use_container_width=True, hide_index=True)

if notes:
    with st.expander(f"엔진이 처리하지 않은 명령 {len(notes)}개"):
        for x in notes:
            st.markdown(f"- `{x}`")

hit_map: dict = {}
key = next((c for c in df.columns if c.lower() in ("no", "id")), df.columns[0])
for c in checks:
    for case in c.cases:
        hit_map.setdefault(case, []).append(str(c.seq))

if hit_map:
    st.markdown("**케이스별 통합** — SPSS 는 체크마다 따로 출력하므로 이 표가 따로 필요합니다.")
    merged = pd.DataFrame(
        [{"케이스": k, "걸린 체크 수": len(v), "체크 번호": ", ".join(v)}
         for k, v in sorted(hit_map.items(), key=lambda kv: -len(kv[1]))])
    st.dataframe(merged, use_container_width=True, hide_index=True)
    st.markdown("**Error 지정 신택스** — 검토 후 필요한 줄만 남기세요.")
    st.code("\n".join(
        f"if ({key.lower()} = {k} ) Error = 1.   /* 체크 {', '.join(v)} */"
        for k, v in sorted(hit_map.items())), language="text")
else:
    st.success("위반이 없습니다.")

# -*- coding: utf-8 -*-
"""
우편번호 DB(지역별 주소 DB) → 읍면동 대조표 변환기

인터넷우체국(epost.go.kr)에서 받은 시도별 txt 파일을 읽어,
「시군구 + 도로명 + 건물번호 → 읍면동」 대조표를 시도별 csv.gz 로 만듭니다.

사용법
    python 우편번호DB_변환.py <txt폴더> [출력폴더]

예
    python 우편번호DB_변환.py C:\\Users\\me\\Downloads\\우편번호DB  emd_data

출력
    emd_data/서울특별시.csv.gz  …  17개 파일 (총 20~25MB)
    emd_data/_요약.csv          변환 결과 점검표

주의
  · 원본 txt 는 UTF-8(BOM) 또는 CP949 로 배포됩니다. 둘 다 자동 처리합니다.
  · 컬럼 순서는 헤더 줄을 읽어 맞추므로, 우편번호 DB 개정으로 컬럼이 늘어도
    이름이 유지되는 한 그대로 동작합니다. 헤더가 없으면 26개 기본 순서를 씁니다.
"""
from __future__ import annotations

import gzip
import sys
from pathlib import Path

import pandas as pd

# 우편번호 DB(지역별 주소 DB) 기본 컬럼 순서 — 헤더가 없는 파일용 폴백
DEFAULT_COLS = [
    "우편번호", "시도", "시도영문", "시군구", "시군구영문", "읍면", "읍면영문",
    "도로명코드", "도로명", "도로명영문", "지하여부", "건물번호본번", "건물번호부번",
    "건물관리번호", "다량배달처명", "시군구용건물명", "법정동코드", "법정동명", "리명",
    "행정동명", "산여부", "지번본번", "읍면동일련번호", "지번부번", "구우편번호",
    "우편번호일련번호",
]

NEEDED = ["우편번호", "시도", "시군구", "읍면", "도로명", "지하여부",
          "건물번호본번", "건물번호부번", "법정동명", "리명", "행정동명"]

OUT_COLS = ["시도", "시군구", "도로명", "지하", "본번", "부번",
            "행정동", "법정동", "우편번호"]


def read_db(path: Path) -> pd.DataFrame:
    """시도별 txt 한 개를 읽는다. 인코딩과 헤더 유무를 자동 판별."""
    last_err = None
    for enc in ("utf-8-sig", "cp949", "utf-8"):
        try:
            with open(path, encoding=enc) as fh:
                first = fh.readline()
            has_header = "우편번호" in first and "도로명" in first
            kwargs = dict(sep="|", dtype=str, encoding=enc, keep_default_na=False,
                          engine="python", on_bad_lines="skip")
            if has_header:
                df = pd.read_csv(path, header=0, **kwargs)
                df.columns = [c.strip().lstrip("\ufeff") for c in df.columns]
            else:
                df = pd.read_csv(path, header=None, names=DEFAULT_COLS, **kwargs)
            missing = [c for c in NEEDED if c not in df.columns]
            if missing:
                raise ValueError(f"필요한 컬럼이 없습니다: {missing}")
            return df
        except (UnicodeDecodeError, UnicodeError) as exc:
            last_err = exc
            continue
    raise RuntimeError(f"{path.name} 인코딩을 알 수 없습니다 ({last_err})")


def build(df: pd.DataFrame) -> pd.DataFrame:
    """조회용 대조표로 정리."""
    for c in NEEDED:
        df[c] = df[c].astype(str).str.strip()

    out = pd.DataFrame({
        "시도": df["시도"],
        # 세종특별자치시는 시군구 단계가 없어 빈 값으로 남는다
        "시군구": df["시군구"],
        "도로명": df["도로명"],
        "지하": (df["지하여부"] == "1").astype("int8"),
        "본번": pd.to_numeric(df["건물번호본번"], errors="coerce").fillna(0).astype("int32"),
        "부번": pd.to_numeric(df["건물번호부번"], errors="coerce").fillna(0).astype("int32"),
        # 행정동 : 행정동명 → 법정동명 → 읍면 순으로 채운다.
        # 시도마다 채움 방식이 달라(세종의 읍면 지역은 행정동명이 비어 있다)
        # 한 단계만 폴백하면 읍면 지역이 통째로 버려진다.
        "행정동": df["행정동명"]
            .where(df["행정동명"] != "", df["법정동명"])
            .where(lambda x: x != "", df["읍면"]),
        # 법정동 : 읍면 지역은 읍면명, 동 지역은 법정동명
        "법정동": df["읍면"].where(df["읍면"] != "", df["법정동명"]),
        "우편번호": df["우편번호"],
    })
    # 버리는 기준: 도로명이 없거나, 행정동·법정동이 모두 없는 줄
    drop_road = int((out["도로명"] == "").sum())
    drop_dong = int(((out["도로명"] != "") &
                     (out["행정동"] == "") & (out["법정동"] == "")).sum())
    out = out[(out["도로명"] != "") & ((out["행정동"] != "") | (out["법정동"] != ""))]
    key = ["시도", "시군구", "도로명", "지하", "본번", "부번"]
    before = len(out)
    out = out.drop_duplicates(subset=key).sort_values(key).reset_index(drop=True)
    out.attrs["drop_road"] = drop_road
    out.attrs["drop_dong"] = drop_dong
    out.attrs["drop_dupe"] = before - len(out)
    return out


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 1
    src = Path(argv[1])
    dst = Path(argv[2]) if len(argv) > 2 else Path("emd_data")
    if not src.is_dir():
        print(f"폴더를 찾을 수 없습니다: {src}")
        return 1

    files = sorted(p for p in src.iterdir() if p.suffix.lower() == ".txt")
    if not files:
        print(f"{src} 안에 .txt 파일이 없습니다.")
        return 1
    dst.mkdir(parents=True, exist_ok=True)

    report, total_rows, total_bytes = [], 0, 0
    for path in files:
        try:
            raw = read_db(path)
            table = build(raw)
        except Exception as exc:
            print(f"  [실패] {path.name}: {exc}")
            report.append({"파일": path.name, "원본행": 0, "대조표행": 0,
                           "용량KB": 0, "비고": f"실패: {exc}"})
            continue

        out_path = dst / (path.stem + ".csv.gz")
        with gzip.open(out_path, "wt", encoding="utf-8", newline="") as fh:
            table[OUT_COLS].to_csv(fh, index=False)
        size = out_path.stat().st_size
        total_rows += len(table)
        total_bytes += size

        # 품질 점검 : 완전키가 행정동 하나로 확정되는 비율
        dup = raw.copy()
        for c in NEEDED:
            dup[c] = dup[c].astype(str).str.strip()
        dup["_ad"] = (dup["행정동명"].where(dup["행정동명"] != "", dup["법정동명"])
                      .where(lambda x: x != "", dup["읍면"]))
        dup["_bon"] = pd.to_numeric(dup["건물번호본번"], errors="coerce").fillna(0)
        dup["_bu"] = pd.to_numeric(dup["건물번호부번"], errors="coerce").fillna(0)
        g = dup.groupby(["시군구", "도로명", "지하여부", "_bon", "_bu"])["_ad"].nunique()
        clean = (g <= 1).mean() * 100 if len(g) else 100.0

        dr = table.attrs.get("drop_road", 0)
        dd = table.attrs.get("drop_dong", 0)
        du = table.attrs.get("drop_dupe", 0)
        print(f"  {path.stem:<12} 원본 {len(raw):>8,}행 → 대조표 {len(table):>8,}행 "
              f"({size/1024:>7,.0f}KB) 확정률 {clean:.2f}%"
              f"  [제외 도로명없음 {dr:,} · 동정보없음 {dd:,} · 중복키 {du:,}]")
        report.append({"파일": path.name, "원본행": len(raw), "대조표행": len(table),
                       "제외_도로명없음": dr, "제외_동정보없음": dd, "제외_중복키": du,
                       "용량KB": round(size / 1024), "비고": f"확정률 {clean:.2f}%"})

    pd.DataFrame(report).to_csv(dst / "_요약.csv", index=False, encoding="utf-8-sig")
    print(f"\n완료 — {len(files)}개 파일, 대조표 {total_rows:,}행, 합계 {total_bytes/1e6:.1f}MB")
    print(f"저장 위치: {dst.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))

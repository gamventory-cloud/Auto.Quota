"""코드북 회귀 비교 — 파서를 고친 뒤 기존 설문지 결과가 바뀌지 않았는지 확인.

    # 현재 버전 결과를 기준선으로 저장
    python3 regression_check.py snapshot baseline/ 설문지1.docx 설문지2.docx

    # 파서를 고친 뒤, 기준선과 비교
    python3 regression_check.py compare baseline/ 설문지1.docx 설문지2.docx

변수명·라벨·유형·측도·값라벨·결측을 하나씩 대조해 다음을 보고한다.
    사라진 변수 / 새로 생긴 변수 / 라벨 변경 / 값라벨 변경 / 유형 변경

개수만 같은 것으로는 회귀를 잡을 수 없다. 값라벨이 조용히 바뀌는 쪽이 더 위험하다.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import spss_labels as sl


# 설문지별 파싱 옵션. 기준선을 만들 때 쓴 옵션과 반드시 같아야 한다.
# (예: Q14 를 0부터 코딩했다면 비교할 때도 같은 옵션을 줘야 한다)
OPTIONS_FILE = "parse_options.json"


def options_for(path: str, basedir: Path) -> dict:
    src = basedir / OPTIONS_FILE
    if not src.exists():
        return {}
    conf = json.loads(src.read_text(encoding="utf-8"))
    return conf.get(Path(path).stem[:60], {})


def fingerprint(path: str, opts: dict | None = None) -> dict:
    """설문지 하나의 파싱 결과를 비교 가능한 dict 로."""
    opts = opts or {}
    variables = sl.parse_upload(
        Path(path).read_bytes(),
        base0=opts.get("base0", []),
        full_labels=opts.get("full_labels", False),
        multi_style=opts.get("multi_style", "category"),
    )
    return {
        v.name: {
            "label": v.label,
            "vtype": v.vtype,
            "measure": v.measure,
            "values": {str(k): val for k, val in sorted(v.values.items())},
            "missing": v.missing,
            "kind": v.kind,
        }
        for v in variables
    }


def snapshot(outdir: Path, docs: list[str]) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    for doc in docs:
        data = fingerprint(doc, options_for(doc, outdir))
        dst = outdir / (Path(doc).stem[:60] + ".json")
        dst.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"기준선 저장: {dst.name}  ({len(data)}개 변수)")


def compare_one(base: dict, now: dict, tag: str) -> int:
    """차이 개수를 반환 (0이면 완전 동일)."""
    gone = [n for n in base if n not in now]
    added = [n for n in now if n not in base]
    changed: dict[str, list[str]] = {}
    for name in base:
        if name not in now:
            continue
        diffs = []
        for field in ("label", "vtype", "measure", "missing", "values"):
            if base[name].get(field) != now[name].get(field):
                diffs.append(field)
        if diffs:
            changed[name] = diffs

    total = len(gone) + len(added) + len(changed)
    mark = "동일" if total == 0 else f"차이 {total}건"
    print(f"\n=== {tag}: 기준 {len(base)}개 -> 현재 {len(now)}개  [{mark}] ===")

    if gone:
        print(f"  [사라진 변수 {len(gone)}] " + ", ".join(gone[:15])
              + (" …" if len(gone) > 15 else ""))
    if added:
        print(f"  [새 변수 {len(added)}] " + ", ".join(added[:15])
              + (" …" if len(added) > 15 else ""))
    if changed:
        print(f"  [변경된 변수 {len(changed)}]")
        for name, fields in list(changed.items())[:12]:
            print(f"    - {name}: {', '.join(fields)}")
            if "values" in fields:
                b = base[name]["values"]
                n = now[name]["values"]
                print(f"        기준: {list(b.items())[:4]}")
                print(f"        현재: {list(n.items())[:4]}")
            if "label" in fields:
                print(f"        기준: {base[name]['label'][:70]}")
                print(f"        현재: {now[name]['label'][:70]}")
        if len(changed) > 12:
            print(f"    … 외 {len(changed) - 12}건")
    return total


def compare(basedir: Path, docs: list[str]) -> int:
    total = 0
    for doc in docs:
        src = basedir / (Path(doc).stem[:60] + ".json")
        if not src.exists():
            print(f"\n=== {Path(doc).name}: 기준선 없음 (건너뜀) ===")
            continue
        base = json.loads(src.read_text(encoding="utf-8"))
        total += compare_one(base, fingerprint(doc, options_for(doc, basedir)),
                             Path(doc).stem[:40])
    print(f"\n총 차이: {total}건")
    return total


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        raise SystemExit(1)
    mode, target, *documents = sys.argv[1:]
    if mode == "snapshot":
        snapshot(Path(target), documents)
    elif mode == "compare":
        raise SystemExit(1 if compare(Path(target), documents) else 0)
    else:
        print(f"알 수 없는 모드: {mode}")
        raise SystemExit(1)

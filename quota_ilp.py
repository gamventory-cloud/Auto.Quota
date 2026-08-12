"""
╔══════════════════════════════════════════════════════════════════════════╗
║  파일명 : quota_ilp.py                                                    ║
║  위치   : 리포지토리 최상단  (pages/ 폴더 안이 아님!)                       ║
║  필요   : pip install ortools   (없으면 앱은 그리디 방식만 노출)             ║
╚══════════════════════════════════════════════════════════════════════════╝

quota_ilp.py — 쿼터 할당을 정수계획법(ILP)으로 정확히 푸는 솔버

랜덤 재시작 그리디와의 차이
--------------------------
  그리디 : 10,000번 돌려도 "이게 최선인지" 알 수 없다. 미달이면 원인도 추정뿐이다.
  ILP    : 최적해임이 증명되고, 미달일 때 어느 제약이 막고 있는지 정확히 나온다.

핵심 기법 : 프로파일 집약 (aggregation)
-------------------------------------
응답자 5만 명에게 0/1 변수를 5만 개 두면 느리다. 그런데 쿼터 제약은 오직
"각 키를 몇 명 뽑았는가"에만 의존하므로, 동일한 (메인 키, 추가 키 조합) 을
가진 응답자들은 서로 완전히 교환 가능하다.

  → 같은 서명(signature)을 가진 응답자를 하나의 프로파일로 묶고,
    "이 프로파일에서 몇 명 뽑을까" 라는 정수 변수 하나만 둔다.
    변수 개수가 5만 → 수백~수천으로 줄고, 해를 다시 개인으로 펼치기만 하면 된다.

이 집약이 최적성을 훼손하지 않는 이유: 0 ≤ n_p ≤ avail_p 인 어떤 정수 배분도
실제 응답자 선택으로 항상 실현 가능하고, 모든 제약과 목적함수가 n_p 만의
함수이기 때문이다.

수학 모형
--------
  변수   n_p ∈ {0, ..., avail_p}         프로파일 p 에서 뽑을 인원
         short_k ∈ {0, ..., T_k}         메인 셀 k 의 부족분

  제약   Σ_{p ∈ cell(k)} n_p + short_k = T_k        (메인 쿼터, 등식 + 부족분)
         Σ_{p ∋ (j,v)}   n_p           ≤ C_{j,v}    (추가 쿼터, 상한)

  목표   min Σ_k w_k · short_k                       (1단계: 총 부족 최소화)
         min max_k short_k                           (2단계: 부족분 고르게 분산)

의존성 : pip install ortools
"""

from __future__ import annotations

import collections
import time
from dataclasses import dataclass, field

try:
    from ortools.sat.python import cp_model
except ImportError as e:  # pragma: no cover
    raise ImportError(
        "OR-Tools가 필요합니다.  pip install ortools"
    ) from e


# 파일 내용이 뒤섞이는 사고를 잡아내기 위한 표식
MODULE_ROLE = "quota_ilp"
__version__ = "1.0"


# ==============================================================================
# 결과 컨테이너
# ==============================================================================
@dataclass
class Diagnosis:
    """미달 원인 진단."""
    main_short: dict = field(default_factory=dict)      # 메인 셀 -> 부족 인원
    main_avail: dict = field(default_factory=dict)      # 메인 셀 -> 데이터 보유 인원
    main_reason: dict = field(default_factory=dict)     # 메인 셀 -> 사유 문자열
    binding: list = field(default_factory=list)         # 한도까지 꽉 찬 추가 쿼터
    group_relax_gain: dict = field(default_factory=dict)   # 그룹 해제 시 추가 확보 인원
    value_relax_gain: dict = field(default_factory=dict)   # 특정 한도 +1 시 확보 인원


@dataclass
class IlpSolution:
    status: str                  # OPTIMAL / FEASIBLE / UNKNOWN
    proven_optimal: bool         # 최적해임이 증명되었는가
    selected: list               # 선택된 응답자의 인덱스 라벨
    total: int                   # 선택 인원
    target_total: int            # 목표 합계
    main_actual: dict            # 메인 셀 -> 달성 인원
    ex_actual: list              # 그룹별 {키 -> 달성 인원}
    diagnosis: Diagnosis
    n_profiles: int              # 집약된 프로파일 수
    n_rows_considered: int       # 모형에 들어간 응답자 수
    solve_sec: float


# ==============================================================================
# 1. 프로파일 집약
# ==============================================================================
def build_profiles(m_keys, ex_keys_list, main_map, ex_maps):
    """
    응답자를 동일 서명 단위로 묶는다.

    서명 = (메인 키, 그룹별로 '제약이 걸린 키'만 정렬한 튜플)

    - 메인 목표가 없는(또는 0인) 셀의 응답자는 애초에 뽑을 수 없으므로 제외
    - ex_maps 에 없는 추가 키는 제약을 만들지 않으므로 서명에서 제거
      (이것만으로도 프로파일 수가 크게 줄어든다)

    반환: profiles = [(서명, [행 위치, ...]), ...]
    """
    n_rows = len(m_keys)
    n_groups = len(ex_maps)
    buckets = collections.defaultdict(list)

    for i in range(n_rows):
        mk = m_keys[i]
        if main_map.get(mk, 0) <= 0:
            continue                                # 목표 없는 셀 → 선택 불가
        sig_ex = []
        for j in range(n_groups):
            e_map = ex_maps[j]
            if not e_map:
                sig_ex.append(())
                continue
            ks = ex_keys_list[j][i]
            # 상한이 걸린 키만 남긴다 (중복 제거 + 정렬로 서명 정규화)
            constrained = sorted({k for k in ks if e_map.get(k, 0) > 0}, key=repr)
            sig_ex.append(tuple(constrained))
        buckets[(mk, tuple(sig_ex))].append(i)

    return list(buckets.items())


# ==============================================================================
# 2. 모형 구성 및 해 구하기
# ==============================================================================
def _solve_core(profiles, main_map, ex_maps, weights=None,
                skip_groups=frozenset(), cap_bonus=None,
                balance=False, time_limit=30.0, workers=8, log=False):
    """
    집약된 모형을 한 번 푼다. 내부용.

    skip_groups : 이 그룹들의 추가 쿼터 제약을 무시한다 (완화 실험용)
    cap_bonus   : {(그룹, 키): +Δ} 로 특정 한도를 늘린다 (민감도 분석용)
    """
    model = cp_model.CpModel()

    # --- 변수: 프로파일별 선택 인원 ---
    n_vars = []
    for idx, (_sig, rows) in enumerate(profiles):
        n_vars.append(model.NewIntVar(0, len(rows), f"n{idx}"))

    # --- 메인 쿼터: 등식 + 부족분 슬랙 ---
    by_main = collections.defaultdict(list)
    for idx, ((mk, _), _rows) in enumerate(profiles):
        by_main[mk].append(idx)

    short_vars = {}
    for k, tgt in main_map.items():
        s = model.NewIntVar(0, tgt, f"short_{abs(hash(k)) % 10**8}")
        short_vars[k] = s
        model.Add(sum(n_vars[i] for i in by_main.get(k, [])) + s == tgt)

    # --- 추가 쿼터: 상한 ---
    by_ex = collections.defaultdict(list)
    for idx, ((_mk, sig_ex), _rows) in enumerate(profiles):
        for j, keys in enumerate(sig_ex):
            if j in skip_groups:
                continue
            for k in keys:
                by_ex[(j, k)].append(idx)

    for j, e_map in enumerate(ex_maps):
        if not e_map or j in skip_groups:
            continue
        for k, cap in e_map.items():
            if cap <= 0:
                continue
            members = by_ex.get((j, k), [])
            if not members:
                continue
            eff = cap + (cap_bonus or {}).get((j, k), 0)
            model.Add(sum(n_vars[i] for i in members) <= eff)

    # --- 1단계 목표: 가중 총 부족 최소화 ---
    w = weights or {}
    model.Minimize(sum(w.get(k, 1) * s for k, s in short_vars.items()))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(time_limit)
    solver.parameters.num_search_workers = int(workers)
    solver.parameters.log_search_progress = bool(log)
    st = solver.Solve(model)

    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, st, solver, n_vars, short_vars

    # --- 2단계(선택): 총 부족을 고정하고 최대 부족을 최소화 → 고르게 분산 ---
    if balance and st == cp_model.OPTIMAL:
        best_obj = int(round(solver.ObjectiveValue()))
        model.Add(sum(w.get(k, 1) * s for k, s in short_vars.items()) == best_obj)
        worst = model.NewIntVar(0, max(main_map.values()), "worst")
        model.AddMaxEquality(worst, list(short_vars.values()))
        model.Minimize(worst)
        solver2 = cp_model.CpSolver()
        solver2.parameters.max_time_in_seconds = float(time_limit)
        solver2.parameters.num_search_workers = int(workers)
        st2 = solver2.Solve(model)
        if st2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            solver, st = solver2, st2

    return True, st, solver, n_vars, short_vars


def solve_quota_ilp(m_keys, ex_keys_list, main_map, ex_maps, indices,
                    weights=None, balance=False, time_limit=30.0, workers=8,
                    diagnose=True, max_value_probes=8, rng=None):
    """
    쿼터 할당 최적화.

    인자는 utils.build_*_keys / simulation_worker 와 동일한 형태를 그대로 받는다.
      m_keys       : 행별 메인 키 (튜플)
      ex_keys_list : 그룹별 [행별 키 리스트]
      main_map     : {메인 키: 목표}
      ex_maps      : [{추가 키: 상한}, ...]
      indices      : 행별 인덱스 라벨 (df.index.to_numpy())

    반환: IlpSolution
    """
    t0 = time.perf_counter()
    target_total = sum(main_map.values())

    profiles = build_profiles(m_keys, ex_keys_list, main_map, ex_maps)
    n_considered = sum(len(rows) for _sig, rows in profiles)

    if not profiles:
        return IlpSolution(
            status="EMPTY", proven_optimal=True, selected=[], total=0,
            target_total=target_total, main_actual={},
            ex_actual=[{} for _ in ex_maps],
            diagnosis=Diagnosis(
                main_short=dict(main_map),
                main_avail={k: 0 for k in main_map},
                main_reason={k: "⚠️ 물리적 부족 (데이터 없음)" for k in main_map},
            ),
            n_profiles=0, n_rows_considered=0,
            solve_sec=time.perf_counter() - t0,
        )

    ok, st, solver, n_vars, short_vars = _solve_core(
        profiles, main_map, ex_maps, weights=weights,
        balance=balance, time_limit=time_limit, workers=workers)

    if not ok:
        return IlpSolution(
            status=solver.StatusName(st), proven_optimal=False, selected=[], total=0,
            target_total=target_total, main_actual={},
            ex_actual=[{} for _ in ex_maps], diagnosis=Diagnosis(),
            n_profiles=len(profiles), n_rows_considered=n_considered,
            solve_sec=time.perf_counter() - t0,
        )

    # ------------------------------------------------------------------
    # 해를 개인 단위로 펼치기
    # ------------------------------------------------------------------
    selected = []
    for idx, (_sig, rows) in enumerate(profiles):
        take = solver.Value(n_vars[idx])
        if take <= 0:
            continue
        if rng is not None and take < len(rows):
            pick = list(rng.choice(rows, size=take, replace=False))
        else:
            pick = rows[:take]
        selected.extend(indices[i] for i in pick)

    # ------------------------------------------------------------------
    # 달성 현황 집계
    # ------------------------------------------------------------------
    main_actual = {k: 0 for k in main_map}
    ex_actual = [collections.Counter() for _ in ex_maps]
    for idx, ((mk, sig_ex), _rows) in enumerate(profiles):
        take = solver.Value(n_vars[idx])
        if take <= 0:
            continue
        main_actual[mk] = main_actual.get(mk, 0) + take
        for j, keys in enumerate(sig_ex):
            for k in keys:
                ex_actual[j][k] += take
    ex_actual = [dict(c) for c in ex_actual]

    total = sum(main_actual.values())

    # ------------------------------------------------------------------
    # 진단
    # ------------------------------------------------------------------
    diag = Diagnosis()
    if diagnose:
        avail = collections.Counter()
        for (mk, _), rows in profiles:
            avail[mk] += len(rows)

        for k, tgt in main_map.items():
            short = tgt - main_actual.get(k, 0)
            diag.main_avail[k] = avail.get(k, 0)
            if short <= 0:
                continue
            diag.main_short[k] = short
            if avail.get(k, 0) < tgt:
                phys = tgt - avail.get(k, 0)
                if short <= phys:
                    diag.main_reason[k] = f"⚠️ 물리적 부족 (보유 {avail.get(k, 0)}명)"
                else:
                    diag.main_reason[k] = (
                        f"⚠️+⚔️ 물리적 {phys}명 + 경합 {short - phys}명")
            else:
                diag.main_reason[k] = "⚔️ 경합 부족 (추가 쿼터 상한에 막힘)"

        # 한도까지 꽉 찬 추가 쿼터 = 병목
        for j, e_map in enumerate(ex_maps):
            for k, cap in e_map.items():
                used = ex_actual[j].get(k, 0)
                if cap > 0 and used >= cap:
                    diag.binding.append({'group': j, 'key': k, 'cap': cap, 'used': used})

        # 미달이면 원인을 정량화한다
        if total < target_total:
            active = [j for j, m in enumerate(ex_maps) if m]

            # (a) 그룹 단위 완화: 이 그룹 제약을 통째로 풀면 몇 명 더 뽑히는가
            for j in active:
                ok2, st2, s2, nv2, _ = _solve_core(
                    profiles, main_map, ex_maps, weights=weights,
                    skip_groups=frozenset([j]),
                    time_limit=min(10.0, time_limit), workers=workers)
                if ok2:
                    gain = sum(s2.Value(v) for v in nv2) - total
                    if gain > 0:
                        diag.group_relax_gain[j] = gain

            # (b) 값 단위 민감도: 병목 한도를 1 늘리면 몇 명 더 뽑히는가 (섀도 프라이스)
            probes = sorted(diag.binding, key=lambda d: -d['cap'])[:max_value_probes]
            for b in probes:
                ok3, st3, s3, nv3, _ = _solve_core(
                    profiles, main_map, ex_maps, weights=weights,
                    cap_bonus={(b['group'], b['key']): 1},
                    time_limit=min(5.0, time_limit), workers=workers)
                if ok3:
                    gain = sum(s3.Value(v) for v in nv3) - total
                    if gain > 0:
                        diag.value_relax_gain[(b['group'], b['key'])] = gain

    return IlpSolution(
        status=solver.StatusName(st),
        proven_optimal=(st == cp_model.OPTIMAL),
        selected=selected,
        total=total,
        target_total=target_total,
        main_actual=main_actual,
        ex_actual=ex_actual,
        diagnosis=diag,
        n_profiles=len(profiles),
        n_rows_considered=n_considered,
        solve_sec=time.perf_counter() - t0,
    )
